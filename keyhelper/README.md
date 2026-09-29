# keyhelper/ - Key Helper, the Mac's key broker

Built from the 09/29/2026 security review: the key library (the login-keychain blob `automation-qbo`)
was readable by any program running as the owner, silently, because "Always Allow" had been given to
`/usr/bin/security`. Key Helper closes that.

## What it does

- **The Keychain item trusts only Key Helper**, pinned to its exact build. Anything else that asks gets a
  macOS password dialog (click **Deny**), never the keys.
- **One unlock per work session** - Touch ID, or the Mac password when the lid is closed. The keys live in
  the helper's memory only and are wiped on screen lock, sleep, logout, **Lock now** (menu bar), quit, or
  after 8 hours.
- **Tools never get the QuickBooks master keys.** The helper renews the login itself and hands out
  one-hour passes; the refresh token and the app secret never leave it.
- **Read-only profiles get no pass at all:** the helper runs the GET and returns the result, so a
  workspace allowed only to read can never write.
- **Only the pinned Python (python-env) running a script inside a registered workspace** is answered.
- **Every request is logged** (never a secret): `~/Library/Logs/Proficient/keyhelper/handouts.log`.

## Files

| File | What |
|---|---|
| `KeyHelper.swift` | the whole helper: socket server, caller check, session, Keychain, QuickBooks renewal + read proxy, menu bar |
| `Info.plist` | app bundle metadata (menu-bar only, no Dock icon) |
| `install.sh` | build + ad-hoc sign + install + register this workspace; `--adopt` = the cutover; `--status`; `--register` (no rebuild) |
| `migrate_keys.py` | one-time move of the Notion token + Teams webhook out of the older `proficient-automation-worker` store (dry run by default) |
| `../shared/key_broker.py` | the Python client every tool uses (through `qbo_api.get_pass()` / `qbo_vault`) |

Outside the repo (never committed): the app at `~/Library/Application Support/Proficient/bin/Key Helper.app`;
`~/Library/Application Support/Proficient/keys/` (0700) holds the socket `k.sock`, the adoption marker
`adopted`, and `workspaces/*.json` - each workspace's registration (which profiles, full or read, which keys).

## Commands (from the repo root)

```bash
bash keyhelper/install.sh
```

```bash
bash keyhelper/install.sh --adopt
```

```bash
bash keyhelper/install.sh --status
```

The first builds, installs and registers - nothing changes for the tools. The second is the one-time
cutover: macOS asks for the login password for `automation-qbo` (click **Allow**, not Always Allow), then
Touch ID / password; the item is rewritten to trust only Key Helper (a verified temp copy first, so the keys
are never missing), and the `adopted` marker switches every tool over. The third shows locked / unlocked.

**After a rebuild** (a new `install.sh` run) the Keychain no longer recognises the build, so the next unlock
shows the macOS dialog once more and the item is re-homed to the new build automatically.

**If a macOS dialog ever asks for `automation-qbo` for something that is not Key Helper** (an old worktree,
a stray script): click **Deny**. "Always Allow" there reopens the hole this tool closes.

## Limits (stated plainly)

- A program running as the owner while the session is unlocked can ask for a one-hour pass - the caller
  check sees the pinned Python, not which script is honest. It is logged and it never gets the master key.
- Someone with admin control of the Mac can read any program's memory. FileVault + the login password
  protect that, not this tool.
- Every key now lives in the library. The Notion token and the Teams paid-notice webhook moved in on
  09/29/2026 (`migrate_keys.py`); read any non-QuickBooks key with `shared/qbo_vault.get_secret(NAME)`.
