#!/usr/bin/env bash
# keyhelper/install.sh - build Key Helper, install it, register this workspace.
#
#   bash keyhelper/install.sh            build + install + register (safe: nothing changes for the tools)
#   bash keyhelper/install.sh --adopt    ...then hand the Keychain key library to the helper (the cutover:
#                                        one macOS dialog - type the login password, click Allow, NOT
#                                        "Always Allow" - then Touch ID / password once)
#   bash keyhelper/install.sh --status   what the helper reports (locked / unlocked, profiles)
#
# The app lives OUTSIDE the repo: ~/Library/Application Support/Proficient/bin/Key Helper.app (ad-hoc
# signed; the Keychain pins this exact build, so a rebuild asks for the macOS dialog once more).
set -euo pipefail
here="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
repo="$(cd "$here/.." && pwd)"
common="$(git -C "$repo" rev-parse --path-format=absolute --git-common-dir)"
main_root="$(dirname "$common")"
. "$repo/python-env/python.sh"

base="$HOME/Library/Application Support/Proficient"
app="$base/bin/Key Helper.app"
keys="$base/keys"

if [ "${1:-}" = "--status" ]; then
  cd "$repo" && "$ACB_PY" -c 'import json,sys; sys.path.insert(0,"."); from shared import key_broker as k; print(json.dumps(k.status(), indent=2))'
  exit 0
fi

echo "== build"
tmp="$(mktemp -d)"
trap 'rm -rf "$tmp"' EXIT
xcrun swiftc -O -swift-version 5 "$here/KeyHelper.swift" -o "$tmp/KeyHelper" 2> "$tmp/build.log" \
  || { grep error "$tmp/build.log"; exit 1; }
mkdir -p "$tmp/Key Helper.app/Contents/MacOS"
cp "$tmp/KeyHelper" "$tmp/Key Helper.app/Contents/MacOS/KeyHelper"
cp "$here/Info.plist" "$tmp/Key Helper.app/Contents/Info.plist"
codesign --force --sign - --identifier local.proficient.keyhelper "$tmp/Key Helper.app" >/dev/null

echo "== install -> $app"
pkill -f "Key Helper.app/Contents/MacOS/KeyHelper" 2>/dev/null || true
sleep 1
mkdir -p "$base/bin"
rm -rf "$app"
cp -R "$tmp/Key Helper.app" "$app"

echo "== register this workspace"
mkdir -p "$keys/workspaces"
chmod 700 "$keys" "$keys/workspaces"
"$ACB_PY" - "$main_root" "$keys/workspaces/concrete.json" <<'PY'
import json, os, sys
root, out = sys.argv[1], sys.argv[2]
reg = {
    "name": "concrete",
    "roots": [root],                       # worktrees (<root>-wt-*) are covered by the helper
    "admin": True,                         # adopt / purge run from here
    "profiles": {"proficient": {"company_key": "QBO_COMPANY_ID", "refresh_key": "QBO_REFRESH_TOKEN", "mode": "full"}},
    "keys": ["QBO_COMPANY_ID", "MIRROR_KEY", "JT_GRANT_KEY", "TEAMS_STMT_WEBHOOK",
             "GRAPH_TENANT_ID", "GRAPH_CLIENT_ID", "GRAPH_CLIENT_SECRET", "GRAPH_BILLING_MAILBOX"],
    "put": ["QBO_CLIENT_ID", "QBO_CLIENT_SECRET", "QBO_COMPANY_ID", "QBO_REFRESH_TOKEN", "MIRROR_KEY",
            "JT_GRANT_KEY", "TEAMS_STMT_WEBHOOK", "GRAPH_TENANT_ID", "GRAPH_CLIENT_ID", "GRAPH_CLIENT_SECRET",
            "GRAPH_BILLING_MAILBOX"],
}
fd = os.open(out, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
with os.fdopen(fd, "w") as f:
    json.dump(reg, f, indent=2)
os.chmod(out, 0o600)
print("   registered", out)
PY

open -g "$app"
echo "== Key Helper is running (lock icon in the menu bar)"

if [ "${1:-}" = "--adopt" ]; then
  echo "== adopt: macOS will ask for the login password for 'automation-qbo' - click Allow (not Always Allow)"
  cd "$repo" && "$ACB_PY" -c 'import sys; sys.path.insert(0,"."); from shared import key_broker as k; k.request("adopt"); print("   adopted - the key library now trusts only Key Helper")'
fi
