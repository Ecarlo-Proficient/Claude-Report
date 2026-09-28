// presence.swift - ask macOS to confirm the Mac's owner is here: Touch ID, or the Mac login password.
// Built once by ledger/presence.py into ~/Library/Application Support/Proficient/bin/ (never committed).
// Usage: presence "<reason shown in the macOS dialog>"   exit 0 = confirmed, 1 = refused / cancelled, 2 = unavailable
import Foundation
import LocalAuthentication

let reason = CommandLine.arguments.count > 1 ? CommandLine.arguments[1] : "confirm this action"
let ctx = LAContext()
var err: NSError?
guard ctx.canEvaluatePolicy(.deviceOwnerAuthentication, error: &err) else {
    FileHandle.standardError.write("unavailable: \(err?.localizedDescription ?? "no owner authentication on this Mac")\n".data(using: .utf8)!)
    exit(2)
}
let done = DispatchSemaphore(value: 0)
var ok = false
ctx.evaluatePolicy(.deviceOwnerAuthentication, localizedReason: reason) { success, e in
    ok = success
    if !success { FileHandle.standardError.write("refused: \(e?.localizedDescription ?? "cancelled")\n".data(using: .utf8)!) }
    done.signal()
}
if done.wait(timeout: .now() + 120) == .timedOut {
    FileHandle.standardError.write("refused: no answer in 2 minutes\n".data(using: .utf8)!)
    exit(1)
}
exit(ok ? 0 : 1)
