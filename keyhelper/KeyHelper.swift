// Key Helper - the key broker for this Mac (security review 09/29/2026).
//
// The key library (login-keychain item service "automation-qbo", label "credentials", a base64 JSON blob)
// used to be readable by ANY program running as the owner, silently, because "Always Allow" had been given
// to /usr/bin/security. After adoption the item trusts exactly one program - this helper, pinned to its
// build - so nothing else reads it without a macOS password dialog.
//
// What the helper does:
//   * one "Touch ID or password" per work session; keys held in memory only, wiped on screen lock, sleep,
//     logout, "Lock now", quit, or after 8 hours
//   * renews QuickBooks itself and hands out one-hour passes - never the refresh token or the app secret
//   * for a profile a workspace may only READ, it runs the GET itself and returns the result (no pass at
//     all), so read-only holds by construction
//   * answers only the owner's user, only the pinned Python, only scripts inside a registered workspace
//   * logs every request (never a secret) to ~/Library/Logs/Proficient/keyhelper/handouts.log
//
// Wire: a Unix socket <keys dir>/k.sock (dir 0700, socket 0600). One JSON request line in, one JSON
// response out, then close. Workspaces register themselves in <keys dir>/workspaces/*.json.
// Build + install: keyhelper/install.sh. Python client: shared/key_broker.py.

import AppKit
import Darwin
import Foundation
import LocalAuthentication
import Security

// MARK: - paths and constants

let home = FileManager.default.homeDirectoryForCurrentUser.path
let keysDir = home + "/Library/Application Support/Proficient/keys"
let sockPath = keysDir + "/k.sock"
let markerPath = keysDir + "/adopted"
let workspacesDir = keysDir + "/workspaces"
let logDir = home + "/Library/Logs/Proficient/keyhelper"
let logPath = logDir + "/handouts.log"
let venvPython = home + "/.venvs/proficient/bin/python"

let itemService = "automation-qbo"
let itemLabel = "credentials"
let itemAccount = ProcessInfo.processInfo.environment["USER"] ?? NSUserName()
let adoptTempService = "automation-qbo.adopting"

let sessionMax: TimeInterval = 8 * 3600
let passMargin: TimeInterval = 300            // hand out a fresh pass when the cached one has < 5 min left
let tokenURL = URL(string: "https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer")!
let apiBase = "https://quickbooks.api.intuit.com/v3/company/"
let masterKeys: Set<String> = ["QBO_CLIENT_ID", "QBO_CLIENT_SECRET"]

// MARK: - log (never a secret)

let logLock = NSLock()
func log(_ fields: [String: Any]) {
    logLock.lock(); defer { logLock.unlock() }
    try? FileManager.default.createDirectory(atPath: logDir, withIntermediateDirectories: true,
                                             attributes: [.posixPermissions: 0o700])
    var f = fields
    f["at"] = ISO8601DateFormatter().string(from: Date())
    guard let d = try? JSONSerialization.data(withJSONObject: f, options: [.sortedKeys]) else { return }
    if !FileManager.default.fileExists(atPath: logPath) {
        FileManager.default.createFile(atPath: logPath, contents: nil, attributes: [.posixPermissions: 0o600])
    }
    if let h = FileHandle(forWritingAtPath: logPath) {
        h.seekToEndOfFile(); h.write(d); h.write("\n".data(using: .utf8)!); h.closeFile()
    }
}

// MARK: - who is asking

struct Caller {
    let pid: pid_t
    let uid: uid_t
    let program: String
    let argv: [String]
    let cwd: String
    var script: String {
        // python [-flags] script.py args  |  python -m module  |  python -c ...  |  python - (stdin)
        var i = 1
        while i < argv.count {
            let a = argv[i]
            if a == "-m" { return "module:" + (i + 1 < argv.count ? argv[i + 1] : "?") }
            if a == "-c" || a == "-" { return "inline" }
            if a.hasPrefix("-") { i += 1; continue }
            return a.hasPrefix("/") ? a : (cwd as NSString).appendingPathComponent(a)
        }
        return "inline"
    }
    // where the code lives: the script's folder, or the working folder for inline / -m code
    var location: String { script.hasPrefix("/") ? (script as NSString).standardizingPath : cwd }
}

func peerCaller(_ fd: Int32) -> Caller? {
    var pid: pid_t = 0
    var len = socklen_t(MemoryLayout<pid_t>.size)
    guard getsockopt(fd, 0 /* SOL_LOCAL */, 2 /* LOCAL_PEERPID */, &pid, &len) == 0 else { return nil }
    var uid: uid_t = 0, gid: gid_t = 0
    guard getpeereid(fd, &uid, &gid) == 0 else { return nil }
    var buf = [CChar](repeating: 0, count: 4096)
    guard proc_pidpath(pid, &buf, UInt32(buf.count)) > 0 else { return nil }
    let program = String(cString: buf)
    return Caller(pid: pid, uid: uid, program: program, argv: processArgs(pid), cwd: processCwd(pid))
}

func processArgs(_ pid: pid_t) -> [String] {
    var mib: [Int32] = [CTL_KERN, KERN_PROCARGS2, pid]
    var size = 0
    guard sysctl(&mib, 3, nil, &size, nil, 0) == 0, size > 0 else { return [] }
    var buf = [UInt8](repeating: 0, count: size)
    guard sysctl(&mib, 3, &buf, &size, nil, 0) == 0, size > 4 else { return [] }
    let argc = Int(buf.withUnsafeBytes { $0.load(as: Int32.self) })
    var i = 4
    while i < size && buf[i] != 0 { i += 1 }          // exec path
    while i < size && buf[i] == 0 { i += 1 }          // padding
    var out: [String] = []
    while out.count < argc && i < size {
        let start = i
        while i < size && buf[i] != 0 { i += 1 }
        out.append(String(decoding: buf[start..<i], as: UTF8.self))
        i += 1
    }
    return out
}

func processCwd(_ pid: pid_t) -> String {
    var info = proc_vnodepathinfo()
    let n = proc_pidinfo(pid, PROC_PIDVNODEPATHINFO, 0, &info, Int32(MemoryLayout<proc_vnodepathinfo>.size))
    guard n > 0 else { return "" }
    return withUnsafePointer(to: &info.pvi_cdir.vip_path) {
        $0.withMemoryRebound(to: CChar.self, capacity: Int(MAXPATHLEN)) { String(cString: $0) }
    }
}

// The only interpreter the helper answers: the pinned Python (python-env) - Homebrew python@X.Y. Both
// ~/.venvs/proficient/bin/python and a plain `python3` from the same Homebrew formula run this binary.
func allowedPrograms() -> Set<String> {
    var out: Set<String> = []
    let real = (venvPython as NSString).resolvingSymlinksInPath          // .../Versions/3.14/bin/python3.14
    out.insert(real)
    let versionDir = ((real as NSString).deletingLastPathComponent as NSString).deletingLastPathComponent
    out.insert(versionDir + "/Resources/Python.app/Contents/MacOS/Python")
    return out
}

// MARK: - workspaces (registered by each repo's install step)

struct ProfileRule { let companyKey: String; let refreshKey: String; let mode: String }   // mode: full | read
struct Workspace {
    let name: String
    let roots: [String]
    let profiles: [String: ProfileRule]
    let keys: Set<String>
    let put: Set<String>
    let admin: Bool
    func owns(_ path: String) -> Bool {
        roots.contains { r in path == r || path.hasPrefix(r + "/") || path.hasPrefix(r + "-wt-") }
    }
}

func loadWorkspaces() -> [Workspace] {
    guard let names = try? FileManager.default.contentsOfDirectory(atPath: workspacesDir) else { return [] }
    var out: [Workspace] = []
    for n in names.sorted() where n.hasSuffix(".json") {
        let p = workspacesDir + "/" + n
        // a registration another account (or group/world) can write is ignored
        guard let at = try? FileManager.default.attributesOfItem(atPath: p),
              (at[.ownerAccountID] as? NSNumber)?.uint32Value == getuid(),
              ((at[.posixPermissions] as? NSNumber)?.intValue ?? 0o777) & 0o022 == 0,
              let d = FileManager.default.contents(atPath: p),
              let j = try? JSONSerialization.jsonObject(with: d) as? [String: Any] else { continue }
        var profiles: [String: ProfileRule] = [:]
        for (k, v) in (j["profiles"] as? [String: [String: String]]) ?? [:] {
            if let c = v["company_key"], let r = v["refresh_key"] {
                profiles[k] = ProfileRule(companyKey: c, refreshKey: r, mode: v["mode"] == "full" ? "full" : "read")
            }
        }
        out.append(Workspace(name: j["name"] as? String ?? n, roots: j["roots"] as? [String] ?? [],
                             profiles: profiles, keys: Set(j["keys"] as? [String] ?? []),
                             put: Set(j["put"] as? [String] ?? []), admin: j["admin"] as? Bool ?? false))
    }
    return out
}

// MARK: - keychain

enum KCError: Error { case notFound, needsUI(OSStatus), failed(OSStatus), corrupt }

func kcRead(service: String, allowUI: Bool) throws -> Data {
    var q: [String: Any] = [kSecClass as String: kSecClassGenericPassword,
                            kSecAttrService as String: service,
                            kSecAttrAccount as String: itemAccount,
                            kSecMatchLimit as String: kSecMatchLimitOne,
                            kSecReturnData as String: true]
    if !allowUI { q[kSecUseAuthenticationUI as String] = kSecUseAuthenticationUIFail }
    var out: CFTypeRef?
    let st = SecItemCopyMatching(q as CFDictionary, &out)
    if st == errSecItemNotFound { throw KCError.notFound }
    if st == errSecInteractionNotAllowed || st == errSecAuthFailed { throw KCError.needsUI(st) }
    guard st == errSecSuccess, let d = out as? Data else { throw KCError.failed(st) }
    return d
}

func kcUpdate(service: String, data: Data) throws {                  // data only - the ACL stays as it is
    let q: [String: Any] = [kSecClass as String: kSecClassGenericPassword,
                            kSecAttrService as String: service,
                            kSecAttrAccount as String: itemAccount]
    let st = SecItemUpdate(q as CFDictionary, [kSecValueData as String: data] as CFDictionary)
    guard st == errSecSuccess else { throw KCError.failed(st) }
}

// Adding an item that trusts ONLY this helper goes through `security -i` on stdin (never argv).
func secInteractive(_ line: String) -> Int32 {
    let p = Process()
    p.executableURL = URL(fileURLWithPath: "/usr/bin/security")
    p.arguments = ["-i"]
    let inPipe = Pipe(); p.standardInput = inPipe
    p.standardOutput = FileHandle.nullDevice; p.standardError = FileHandle.nullDevice
    do { try p.run() } catch { return -1 }
    inPipe.fileHandleForWriting.write((line + "\n").data(using: .utf8)!)
    inPipe.fileHandleForWriting.closeFile()
    p.waitUntilExit()
    return p.terminationStatus
}

func kcAddTrustingSelf(service: String, label: String, base64: String) -> Bool {
    let me = Bundle.main.executablePath ?? CommandLine.arguments[0]
    for v in [itemAccount, service, label, me] where v.contains("\"") || v.contains("\n") { return false }
    return secInteractive("add-generic-password -a \"\(itemAccount)\" -s \"\(service)\" -l \"\(label)\" "
                          + "-w \"\(base64)\" -U -T \"\(me)\"") == 0
}

func kcDelete(service: String) {
    let q: [String: Any] = [kSecClass as String: kSecClassGenericPassword,
                            kSecAttrService as String: service,
                            kSecAttrAccount as String: itemAccount]
    SecItemDelete(q as CFDictionary)
}

func decodeBlob(_ d: Data) throws -> [String: String] {
    guard let s = String(data: d, encoding: .utf8)?.trimmingCharacters(in: .whitespacesAndNewlines),
          let raw = Data(base64Encoded: s),
          let j = try? JSONSerialization.jsonObject(with: raw) as? [String: String] else { throw KCError.corrupt }
    return j
}

func encodeBlob(_ b: [String: String]) -> String {
    let d = (try? JSONSerialization.data(withJSONObject: b, options: [.sortedKeys])) ?? Data()
    return d.base64EncodedString()
}

// Re-home the item so it trusts only this helper: temp copy first, verified, then the real item, verified,
// then the temp copy goes. A failure part way leaves the temp copy holding the keys (never nothing).
func adoptItem(_ blob: [String: String]) -> String? {
    let b64 = encodeBlob(blob)
    kcDelete(service: adoptTempService)
    guard kcAddTrustingSelf(service: adoptTempService, label: itemLabel, base64: b64),
          let t = try? kcRead(service: adoptTempService, allowUI: false), (try? decodeBlob(t)) == blob else {
        kcDelete(service: adoptTempService)
        return "could not write a verified copy - nothing was changed"
    }
    kcDelete(service: itemService)
    guard kcAddTrustingSelf(service: itemService, label: itemLabel, base64: b64),
          let r = try? kcRead(service: itemService, allowUI: false), (try? decodeBlob(r)) == blob else {
        return "the keys are safe in '\(adoptTempService)' but the main item did not verify - run install again"
    }
    kcDelete(service: adoptTempService)
    FileManager.default.createFile(atPath: markerPath, contents: Data(ISO8601DateFormatter().string(from: Date()).utf8),
                                   attributes: [.posixPermissions: 0o600])
    return nil
}

// MARK: - session

final class Session {
    var blob: [String: String]? = nil
    var unlockedAt: Date? = nil
    var passes: [String: (token: String, expires: Date)] = [:]      // keyed by refresh-key name
    var onChange: (() -> Void)?

    var isUnlocked: Bool {
        guard blob != nil, let t = unlockedAt else { return false }
        return Date().timeIntervalSince(t) < sessionMax
    }

    func lock(_ why: String) {
        if blob != nil { log(["op": "lock", "why": why]) }
        if var b = blob { for k in Array(b.keys) { b[k] = String(repeating: "\0", count: 8) } }
        blob = nil; unlockedAt = nil; passes.removeAll()
        onChange?()
    }

    func askOwner(_ reason: String) -> Bool {
        let ctx = LAContext()
        ctx.localizedFallbackTitle = "Use Password"
        let done = DispatchSemaphore(value: 0)
        var ok = false
        ctx.evaluatePolicy(.deviceOwnerAuthentication, localizedReason: reason) { success, _ in
            ok = success; done.signal()
        }
        done.wait()
        return ok
    }

    // Unlock for the session: owner check, then the keychain. After a rebuild the pinned build no longer
    // matches the item, so the read needs the macOS dialog once and the item is re-homed to this build.
    func unlock(_ reason: String) -> String? {
        if isUnlocked { return nil }
        guard askOwner(reason) else { return "not unlocked - Touch ID / password was cancelled" }
        do {
            blob = try decodeBlob(try kcRead(service: itemService, allowUI: false))
        } catch KCError.needsUI {
            guard let d = try? kcRead(service: itemService, allowUI: true), let b = try? decodeBlob(d) else {
                return "the key library could not be opened (the macOS dialog was denied)"
            }
            if let err = adoptItem(b) { return err }
            blob = b
            log(["op": "re-home", "why": "helper build changed or first adoption"])
        } catch KCError.notFound {
            return "no key library in the Keychain yet - run shared/setup_qbo.py"
        } catch {
            return "the key library could not be read (\(error))"
        }
        unlockedAt = Date()
        log(["op": "unlock"])
        onChange?()
        return nil
    }

    func save() throws {
        guard let b = blob else { return }
        try kcUpdate(service: itemService, data: Data(encodeBlob(b).utf8))
    }
}

// MARK: - QuickBooks

func httpSync(_ req: URLRequest) -> (Int, Data)? {
    let done = DispatchSemaphore(value: 0)
    var out: (Int, Data)? = nil
    URLSession.shared.dataTask(with: req) { d, r, _ in
        if let h = r as? HTTPURLResponse { out = (h.statusCode, d ?? Data()) }
        done.signal()
    }.resume()
    done.wait()
    return out
}

func exchange(_ s: Session, rule: ProfileRule) -> (String, Date)? {
    guard var b = s.blob, let cid = b["QBO_CLIENT_ID"], let sec = b["QBO_CLIENT_SECRET"],
          let rt = b[rule.refreshKey], !rt.isEmpty else { return nil }
    var req = URLRequest(url: tokenURL, timeoutInterval: 30)
    req.httpMethod = "POST"
    req.setValue("Basic " + Data("\(cid):\(sec)".utf8).base64EncodedString(), forHTTPHeaderField: "Authorization")
    req.setValue("application/x-www-form-urlencoded", forHTTPHeaderField: "Content-Type")
    req.setValue("application/json", forHTTPHeaderField: "Accept")
    let enc = rt.addingPercentEncoding(withAllowedCharacters: .alphanumerics) ?? rt
    req.httpBody = Data("grant_type=refresh_token&refresh_token=\(enc)".utf8)
    var res: (Int, Data)? = nil
    for attempt in 0..<4 {                                           // 1 try + 3 retries on network / 5xx
        res = httpSync(req)
        if let r = res, r.0 < 500 { break }
        if attempt < 3 { Thread.sleep(forTimeInterval: Double(attempt + 1) * 3) }
    }
    guard let (code, data) = res, code == 200,
          let j = try? JSONSerialization.jsonObject(with: data) as? [String: Any],
          let access = j["access_token"] as? String else {
        log(["op": "renew", "result": "failed", "status": res?.0 ?? 0])
        return nil
    }
    if let newRT = j["refresh_token"] as? String, newRT != rt {
        b[rule.refreshKey] = newRT
        s.blob = b
        do { try s.save() } catch { log(["op": "renew", "result": "rotated token NOT saved", "error": "\(error)"]) }
    }
    let secs = (j["expires_in"] as? Double) ?? 3600
    return (access, Date().addingTimeInterval(secs))
}

func pass(_ s: Session, rule: ProfileRule, fresh: Bool) -> (String, Date)? {
    if !fresh, let p = s.passes[rule.refreshKey], p.expires.timeIntervalSinceNow > passMargin { return p }
    guard let p = exchange(s, rule: rule) else { return nil }
    s.passes[rule.refreshKey] = (p.0, p.1)
    return p
}

// MARK: - requests

let sessionQueue = DispatchQueue(label: "keyhelper.session")      // one request at a time: no double prompts / renewals
let session = Session()

func handle(_ req: [String: Any], caller: Caller) -> [String: Any] {
    let op = req["op"] as? String ?? ""
    var entry: [String: Any] = ["op": op, "pid": caller.pid, "script": caller.script]
    defer { log(entry) }

    guard caller.uid == getuid() else { entry["result"] = "refused: other user"; return ["error": "refused: other user"] }
    guard allowedPrograms().contains(caller.program) else {
        entry["result"] = "refused: program"; entry["program"] = caller.program
        return ["error": "refused: only the pinned Python (python-env) may ask for keys"]
    }
    guard let ws = loadWorkspaces().first(where: { $0.owns(caller.location) }) else {
        entry["result"] = "refused: not in a registered workspace"
        return ["error": "refused: the script is not inside a registered workspace"]
    }
    entry["workspace"] = ws.name
    let who = caller.script.hasPrefix("/") ? (caller.script as NSString).lastPathComponent : caller.script
    let reason = "unlock the QuickBooks keys for this work session (asked by \(who))"

    switch op {
    case "status":
        entry["result"] = "ok"
        return ["ok": true, "adopted": FileManager.default.fileExists(atPath: markerPath),
                "unlocked": session.isUnlocked, "workspace": ws.name,
                "profiles": ws.profiles.mapValues { $0.mode },
                "locks_in": session.unlockedAt.map { Int(sessionMax - Date().timeIntervalSince($0)) } ?? 0]

    case "lock":
        session.lock("asked by \(who)")
        entry["result"] = "ok"
        return ["ok": true]

    case "adopt":
        guard ws.admin else { entry["result"] = "refused: not admin"; return ["error": "refused: adopt runs from the main workspace"] }
        guard session.askOwner("hand the QuickBooks key library to Key Helper (one time)") else {
            entry["result"] = "cancelled"; return ["error": "cancelled"]
        }
        let b: [String: String]
        do { b = try decodeBlob(try kcRead(service: itemService, allowUI: true)) } catch {
            entry["result"] = "read failed"; return ["error": "the key library could not be read: \(error)"]
        }
        if let err = adoptItem(b) { entry["result"] = err; return ["error": err] }
        session.blob = b; session.unlockedAt = Date(); session.onChange?()
        entry["result"] = "adopted"
        return ["ok": true]

    case "pass":
        let name = req["profile"] as? String ?? ""
        entry["profile"] = name
        guard let rule = ws.profiles[name] else { entry["result"] = "refused: profile"; return ["error": "refused: this workspace has no profile '\(name)'"] }
        guard rule.mode == "full" else { entry["result"] = "refused: read-only"; return ["error": "refused: '\(name)' is read-only here - use the read request, not a pass"] }
        if let err = session.unlock(reason) { entry["result"] = err; return ["error": err] }
        guard let p = pass(session, rule: rule, fresh: req["fresh"] as? Bool ?? false),
              let company = session.blob?[rule.companyKey] else {
            entry["result"] = "renew failed"; return ["error": "QuickBooks did not renew the login (see the helper log) - run shared/setup_qbo.py --test"]
        }
        entry["result"] = "pass"
        return ["ok": true, "access_token": p.0, "company_id": company, "expires_at": p.1.timeIntervalSince1970]

    case "get":                                                   // read-only QuickBooks request, run here
        let name = req["profile"] as? String ?? ""
        let path = req["path"] as? String ?? ""
        entry["profile"] = name; entry["path"] = path
        guard let rule = ws.profiles[name] else { entry["result"] = "refused: profile"; return ["error": "refused: this workspace has no profile '\(name)'"] }
        guard !path.isEmpty, !path.contains(".."), path.range(of: "^[A-Za-z0-9/_-]+$", options: .regularExpression) != nil else {
            entry["result"] = "refused: path"; return ["error": "refused: bad path"]
        }
        if let err = session.unlock(reason) { entry["result"] = err; return ["error": err] }
        guard let company = session.blob?[rule.companyKey] else { entry["result"] = "no company"; return ["error": "no company id stored for '\(name)'"] }
        var comps = URLComponents(string: apiBase + company + "/" + path)!
        comps.queryItems = ((req["params"] as? [String: Any]) ?? [:]).map { URLQueryItem(name: $0.key, value: "\($0.value)") }
        var result: (Int, Data)? = nil
        for attempt in 0..<5 {
            guard let p = pass(session, rule: rule, fresh: attempt > 0 && result?.0 == 401) else { break }
            var r = URLRequest(url: comps.url!, timeoutInterval: 90)
            r.httpMethod = "GET"
            r.setValue("Bearer " + p.0, forHTTPHeaderField: "Authorization")
            r.setValue("application/json", forHTTPHeaderField: "Accept")
            result = httpSync(r)
            guard let code = result?.0, code == 401 || code == 429 || code >= 500 else { break }
            Thread.sleep(forTimeInterval: pow(2.0, Double(attempt)))
        }
        guard let (code, data) = result else { entry["result"] = "network"; return ["error": "QuickBooks did not answer"] }
        entry["result"] = "get \(code)"
        return ["ok": true, "status": code, "body": String(decoding: data, as: UTF8.self)]

    case "key", "keys":
        let names = op == "key" ? [req["name"] as? String ?? ""] : Array(ws.keys)
        entry["keys"] = names.sorted()
        guard names.allSatisfy({ ws.keys.contains($0) && !masterKeys.contains($0) }) else {
            entry["result"] = "refused: key"; return ["error": "refused: this workspace may not read \(names.joined(separator: ", "))"]
        }
        if let err = session.unlock(reason) { entry["result"] = err; return ["error": err] }
        var out: [String: String] = [:]
        for n in names { if let v = session.blob?[n] { out[n] = v } }
        entry["result"] = "ok"
        return ["ok": true, "values": out]

    case "names":
        if let err = session.unlock(reason) { entry["result"] = err; return ["error": err] }
        entry["result"] = "ok"
        return ["ok": true, "names": (session.blob ?? [:]).keys.sorted()]

    case "put", "delete", "purge":                                 // writes: a fresh owner check every time
        let values = req["values"] as? [String: String] ?? [:]
        let dels = req["names"] as? [String] ?? []
        entry["keys"] = (Array(values.keys) + dels).sorted()
        if op == "purge" { guard ws.admin else { entry["result"] = "refused"; return ["error": "refused: purge runs from the main workspace"] } }
        else {
            guard (Array(values.keys) + dels).allSatisfy({ ws.put.contains($0) }) else {
                entry["result"] = "refused: key"; return ["error": "refused: this workspace may not change those keys"]
            }
        }
        if let err = session.unlock(reason) { entry["result"] = err; return ["error": err] }
        guard session.askOwner("change the QuickBooks key library (\(op))") else { entry["result"] = "cancelled"; return ["error": "cancelled"] }
        if op == "purge" {
            kcDelete(service: itemService); session.lock("purged")
            entry["result"] = "purged"; return ["ok": true]
        }
        var b = session.blob ?? [:]
        for (k, v) in values { b[k] = v }
        for k in dels { b.removeValue(forKey: k) }
        session.blob = b
        session.passes.removeAll()
        do { try session.save() } catch { entry["result"] = "save failed"; return ["error": "the key library could not be saved: \(error)"] }
        entry["result"] = "saved"
        return ["ok": true]

    default:
        entry["result"] = "unknown op"
        return ["error": "unknown request '\(op)'"]
    }
}

// MARK: - socket server

func serve() {
    try? FileManager.default.createDirectory(atPath: keysDir, withIntermediateDirectories: true,
                                             attributes: [.posixPermissions: 0o700])
    chmod(keysDir, 0o700)
    unlink(sockPath)
    let fd = socket(AF_UNIX, SOCK_STREAM, 0)
    var addr = sockaddr_un()
    addr.sun_family = sa_family_t(AF_UNIX)
    _ = withUnsafeMutablePointer(to: &addr.sun_path) {
        $0.withMemoryRebound(to: CChar.self, capacity: 104) { strncpy($0, sockPath, 103) }
    }
    let bound = withUnsafePointer(to: &addr) {
        $0.withMemoryRebound(to: sockaddr.self, capacity: 1) { bind(fd, $0, socklen_t(MemoryLayout<sockaddr_un>.size)) }
    }
    guard bound == 0, listen(fd, 16) == 0 else { log(["op": "start", "result": "socket failed"]); exit(1) }
    chmod(sockPath, 0o600)
    log(["op": "start", "result": "listening"])
    Thread.detachNewThread {
        while true {
            let c = accept(fd, nil, nil)
            if c < 0 { continue }
            var tv = timeval(tv_sec: 10, tv_usec: 0)
            setsockopt(c, SOL_SOCKET, SO_RCVTIMEO, &tv, socklen_t(MemoryLayout<timeval>.size))
            var data = Data()
            var buf = [UInt8](repeating: 0, count: 65536)
            while data.count < 1_048_576 {
                let n = read(c, &buf, buf.count)
                if n <= 0 { break }
                data.append(contentsOf: buf[0..<n])
                if buf[0..<n].contains(10) { break }
            }
            guard let caller = peerCaller(c) else { close(c); continue }
            sessionQueue.async {
                let req = (try? JSONSerialization.jsonObject(with: data) as? [String: Any]) ?? [:]
                let resp = handle(req, caller: caller)
                var out = (try? JSONSerialization.data(withJSONObject: resp)) ?? Data("{\"error\":\"encode\"}".utf8)
                out.append(10)
                out.withUnsafeBytes { p in
                    var off = 0
                    while off < p.count {
                        let w = write(c, p.baseAddress! + off, p.count - off)
                        if w <= 0 { break }
                        off += w
                    }
                }
                close(c)
            }
        }
    }
}

// MARK: - menu bar + lock triggers

final class AppDelegate: NSObject, NSApplicationDelegate {
    var item: NSStatusItem!
    let stateLine = NSMenuItem(title: "", action: nil, keyEquivalent: "")

    func applicationDidFinishLaunching(_ n: Notification) {
        item = NSStatusBar.system.statusItem(withLength: NSStatusItem.variableLength)
        let menu = NSMenu()
        menu.addItem(stateLine)
        menu.addItem(.separator())
        menu.addItem(NSMenuItem(title: "Lock now", action: #selector(lockNow), keyEquivalent: "l"))
        menu.addItem(NSMenuItem(title: "Show the handout log", action: #selector(showLog), keyEquivalent: ""))
        menu.addItem(.separator())
        menu.addItem(NSMenuItem(title: "Quit Key Helper", action: #selector(quit), keyEquivalent: "q"))
        item.menu = menu
        session.onChange = { DispatchQueue.main.async { self.refresh() } }
        refresh()

        let lockOn: (String) -> (Notification) -> Void = { why in { _ in sessionQueue.async { session.lock(why) } } }
        DistributedNotificationCenter.default().addObserver(forName: .init("com.apple.screenIsLocked"), object: nil,
                                                            queue: nil, using: lockOn("screen locked"))
        let ws = NSWorkspace.shared.notificationCenter
        ws.addObserver(forName: NSWorkspace.willSleepNotification, object: nil, queue: nil, using: lockOn("sleep"))
        ws.addObserver(forName: NSWorkspace.sessionDidResignActiveNotification, object: nil, queue: nil, using: lockOn("user switched out"))
        ws.addObserver(forName: NSWorkspace.willPowerOffNotification, object: nil, queue: nil, using: lockOn("logout"))
        Timer.scheduledTimer(withTimeInterval: 60, repeats: true) { _ in
            sessionQueue.async { if session.blob != nil && !session.isUnlocked { session.lock("8 hours") } }
            self.refresh()
        }
        serve()
    }

    func refresh() {
        let open = session.isUnlocked
        item.button?.title = open ? "🔓" : "🔒"
        item.button?.toolTip = "Key Helper"
        if open, let t = session.unlockedAt {
            let f = DateFormatter(); f.dateFormat = "h:mm a"
            stateLine.title = "Keys unlocked since \(f.string(from: t)) (locks by \(f.string(from: t.addingTimeInterval(sessionMax))))"
        } else {
            stateLine.title = "Keys locked"
        }
    }

    @objc func lockNow() { sessionQueue.async { session.lock("Lock now") } }
    @objc func showLog() { NSWorkspace.shared.open(URL(fileURLWithPath: logDir)) }
    @objc func quit() { sessionQueue.sync { session.lock("quit") }; unlink(sockPath); NSApp.terminate(nil) }
    func applicationWillTerminate(_ n: Notification) { unlink(sockPath) }
}

let app = NSApplication.shared
app.setActivationPolicy(.accessory)
let delegate = AppDelegate()
app.delegate = delegate
app.run()
