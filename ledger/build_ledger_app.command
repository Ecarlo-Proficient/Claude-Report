#!/bin/bash
# build_ledger_app.command — build "Project Ledger.app", a Dock on/off switch.
#
# Double-click me ONCE. It creates a small app in ~/Applications that gives you:
#   • one-click open  — launch it → starts the dashboard + opens it in your browser
#   • an on/off indicator — its Dock icon is present while the server is ON, gone when OFF
#   • a clean off switch — Quit it (Cmd-Q / right-click Dock icon → Quit), log out, or
#     shut down → the server stops. Closing the lid (sleep) also stops it.
#   • NOT always-on — it only runs while you keep it open; nothing starts at login.
#
# Self-locating: no hard-coded user paths in this file. If you move the repo, just
# double-click this again to rebuild the app against the new location.
set -euo pipefail

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"   # .../ledger
REPO="$(cd "$HERE/.." && pwd)"                          # repo root
APP_DIR="$HOME/Applications"
APP="$APP_DIR/Project Ledger.app"
SRC="$(mktemp -t ledger_app).applescript"

mkdir -p "$APP_DIR"

# AppleScript source — written verbatim (quoted heredoc), repo path injected after.
cat > "$SRC" <<'APPLESCRIPT'
property port : 8787
property launcher : "REPO_PLACEHOLDER/ledger/open_ledger.command"
property lastTick : missing value

on run
	startAndOpen()
end run

on reopen
	startAndOpen()
end reopen

on startAndOpen()
	try
		do shell script "/bin/bash " & quoted form of launcher
	end try
	set lastTick to current date
end startAndOpen

on serverUp()
	try
		do shell script "curl -s -o /dev/null http://127.0.0.1:" & (port as text) & "/api/health"
		return true
	on error
		return false
	end try
end serverUp

on stopServer()
	try
		do shell script "pkill -f 'ledger/dashboard.py' >/dev/null 2>&1; true"
	end try
end stopServer

on idle
	-- a big gap between idle ticks means the Mac slept → turn off on wake
	if lastTick is not missing value then
		if ((current date) - lastTick) > 90 then
			stopServer()
			quit
			return 0
		end if
	end if
	-- keep the Dock indicator honest: if the server is gone, disappear too
	if not serverUp() then
		quit
		return 0
	end if
	set lastTick to current date
	return 20
end idle

on quit
	stopServer()
	continue quit
end quit
APPLESCRIPT

# Bake the real repo path in (keeps THIS tracked file free of any /Users path).
/usr/bin/sed -i '' "s|REPO_PLACEHOLDER|$REPO|g" "$SRC"

rm -rf "$APP"
/usr/bin/osacompile -o "$APP" "$SRC"
rm -f "$SRC"

echo ""
echo "  Built:  $APP"
echo ""
echo "  Next:"
echo "    1. Open it once (double-click in ~/Applications)."
echo "    2. Right-click its Dock icon → Options → Keep in Dock."
echo ""
echo "  From then on: click the Dock icon to open the ledger. Icon present = ON."
echo "  Quit it (Cmd-Q), sleep, log out, or shut down = OFF. Never runs in the background."
echo ""
