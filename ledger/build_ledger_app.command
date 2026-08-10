#!/bin/bash
# build_ledger_app.command — build "Project Ledger.app", a real Dock on/off switch.
#
# Double-click me ONCE. It builds a proper macOS app (via PyObjC + py2app) into
# ~/Applications:
#   • one-click open  — launch it → the dashboard server starts + your browser opens.
#     Click the Dock icon again any time to reopen the tab.
#   • on/off indicator — the Dock icon is present exactly while the ledger is ON.
#   • clean off switch — Cmd-Q / right-click Dock icon → Quit / log out / shut down,
#     and real system sleep (closing the lid), all stop the server cleanly.
#   • NOT always-on — it only runs while you keep it open; nothing starts at login.
#
# Custom icon: drop a square PNG at ledger/app/app_icon.png and re-run me — it's
# converted to an .icns and baked into the app.
#
# Self-locating (no hard-coded user paths). Moved the repo? Double-click me again.
set -euo pipefail

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"   # .../ledger
APPDIR="$HERE/app"
DEST="$HOME/Applications"
LOG_DIR="$HOME/Library/Logs/Proficient/ledger-dashboard"
BUILDLOG="$LOG_DIR/build.log"
mkdir -p "$LOG_DIR" "$DEST"

# One-time: make sure the Mac-app toolkit is installed.
if ! /usr/bin/python3 -c "import AppKit, py2app" >/dev/null 2>&1; then
  echo "Installing the Mac app toolkit (one-time — this can take a minute)…"
  /usr/bin/python3 -m pip install --break-system-packages --quiet pyobjc-framework-Cocoa py2app
fi

# Optional custom icon: app_icon.png -> app_icon.icns (built-in tools only).
if [ -f "$APPDIR/app_icon.png" ] && { [ ! -f "$APPDIR/app_icon.icns" ] || [ "$APPDIR/app_icon.png" -nt "$APPDIR/app_icon.icns" ]; }; then
  echo "Making the app icon from app_icon.png…"
  TMPSET="$(mktemp -d)/icon.iconset"; mkdir -p "$TMPSET"
  for s in 16 32 64 128 256 512; do
    /usr/bin/sips -z "$s" "$s" "$APPDIR/app_icon.png" --out "$TMPSET/icon_${s}x${s}.png" >/dev/null 2>&1 || true
    d=$((s * 2))
    /usr/bin/sips -z "$d" "$d" "$APPDIR/app_icon.png" --out "$TMPSET/icon_${s}x${s}@2x.png" >/dev/null 2>&1 || true
  done
  /usr/bin/iconutil -c icns "$TMPSET" -o "$APPDIR/app_icon.icns" 2>/dev/null || true
  rm -rf "$(dirname "$TMPSET")"
fi

cd "$APPDIR"
rm -rf build dist
echo "Building Project Ledger.app…"
if ! /usr/bin/python3 setup.py py2app -A >"$BUILDLOG" 2>&1; then
  echo "Build failed — last lines of $BUILDLOG:"; tail -15 "$BUILDLOG"; exit 1
fi
rm -rf "$DEST/Project Ledger.app"
cp -R "dist/Project Ledger.app" "$DEST/"
rm -rf build dist

echo ""
echo "  Built:  $DEST/Project Ledger.app"
echo ""
echo "  Next:  open it (double-click), then right-click its Dock icon →"
echo "         Options → Keep in Dock."
echo ""
echo "  Dock icon present = ledger is ON. Cmd-Q / sleep / log out / shut down = OFF."
echo "  Never runs at login; nothing lingers in the background."
echo ""
