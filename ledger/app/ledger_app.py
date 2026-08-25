#!/usr/bin/env python3
"""
Project Ledger.app — a proper Dock app that runs the local ledger dashboard.

A real Cocoa app (via PyObjC + py2app), so:
  * its Dock icon is present exactly while the ledger is ON (the indicator),
  * clicking it opens the dashboard in your browser (starts the server if needed),
  * Cmd-Q / right-click Dock → Quit / log out / shut down stop the server cleanly,
  * real system sleep (closing the lid) also stops it — nothing lingers,
  * it never runs at login; it's on only while you keep it open.

It manages the dashboard server (ledger/dashboard.py) as a child process — the SAME
server the terminal launcher uses — and stops it on quit. Built by build_ledger_app.command.
"""
import os
import subprocess
import threading
import time
import urllib.request
import webbrowser

import AppKit
from PyObjCTools import AppHelper

PORT = str(os.environ.get("ACB_LEDGER_PORT", "8787"))
URL = f"http://127.0.0.1:{PORT}"
# The repo path is baked into the app's Info.plist (LSEnvironment) at build time,
# so it never depends on where the .app bundle lives; fall back to __file__ for a
# plain (non-bundled) run.
REPO = os.environ.get("ACB_LEDGER_REPO") or os.path.dirname(
    os.path.dirname(os.path.dirname(os.path.realpath(__file__))))
LEDGER_DIR = os.path.join(REPO, "ledger")
LOG_DIR = os.path.expanduser("~/Library/Logs/Proficient/ledger-dashboard")

# py2app injects PYTHONPATH/PYTHONHOME into our environment; the child server is a
# plain /usr/bin/python3 and must NOT inherit them (they hide its site-packages,
# e.g. `requests`). Hand it a cleaned environment.
_PY_ENV_STRIP = ("PYTHONPATH", "PYTHONHOME", "PYTHONEXECUTABLE", "PYTHONNOUSERSITE",
                 "PYTHONDONTWRITEBYTECODE", "PYTHONSTARTUP", "RESOURCEPATH", "ARGVZERO")

_proc = None


def _server_up() -> bool:
    try:
        urllib.request.urlopen(URL + "/api/health", timeout=1)
        return True
    except Exception:
        return False


def _start_server() -> None:
    global _proc
    if _server_up():
        return                                        # already running — adopt it
    os.makedirs(LOG_DIR, exist_ok=True)
    logf = open(os.path.join(LOG_DIR, "server.log"), "a")
    env = {k: v for k, v in os.environ.items() if k not in _PY_ENV_STRIP}
    _proc = subprocess.Popen(
        ["/usr/bin/python3", os.path.join(LEDGER_DIR, "dashboard.py"),
         "--no-open", "--port", PORT],
        cwd=REPO, stdout=logf, stderr=subprocess.STDOUT, stdin=subprocess.DEVNULL, env=env)


# Chromium-family browsers open a chromeless "app mode" window (no address bar / tabs) via
# --app=<url>. The owner didn't want the 127.0.0.1 address showing in a browser bar (2026-08-25).
_APP_BROWSERS = [
    "/Applications/Google Chrome.app",
    "/Applications/Microsoft Edge.app",
    "/Applications/Brave Browser.app",
    "/Applications/Chromium.app",
]


def _open_browser() -> None:
    for _ in range(80):                               # wait up to ~20s for the server
        if _server_up():
            break
        time.sleep(0.25)
    for app in _APP_BROWSERS:                          # prefer a chromeless app-mode window
        if os.path.isdir(app):
            try:
                subprocess.Popen(["/usr/bin/open", "-na", app, "--args",
                                  "--app=" + URL, "--window-size=1400,900"])
                return
            except Exception:
                pass
    webbrowser.open(URL)                               # fallback: default browser (shows its address bar)


def _stop_server() -> None:
    # Guarantee the ledger is OFF when the app quits, whoever started the server.
    try:
        subprocess.call(["/usr/bin/pkill", "-f", "ledger/dashboard.py"])
    except Exception:
        pass


class Delegate(AppKit.NSObject):
    def applicationDidFinishLaunching_(self, note):
        _start_server()
        threading.Thread(target=_open_browser, daemon=True).start()
        # Real system sleep (lid close) → stop cleanly. This is the reliable sleep
        # signal AppleScript couldn't get; PyObjC observes it directly.
        nc = AppKit.NSWorkspace.sharedWorkspace().notificationCenter()
        nc.addObserver_selector_name_object_(
            self, "onSleep:", "NSWorkspaceWillSleepNotification", None)

    def onSleep_(self, note):
        AppKit.NSApplication.sharedApplication().terminate_(None)

    def applicationShouldHandleReopen_hasVisibleWindows_(self, app, flag):
        # Clicking the Dock icon again reopens the dashboard tab.
        _start_server()
        threading.Thread(target=_open_browser, daemon=True).start()
        return True

    def applicationShouldTerminate_(self, sender):
        _stop_server()
        return AppKit.NSTerminateNow


def main():
    app = AppKit.NSApplication.sharedApplication()
    app.setActivationPolicy_(AppKit.NSApplicationActivationPolicyRegular)   # Dock icon
    delegate = Delegate.alloc().init()
    app.setDelegate_(delegate)
    AppHelper.runEventLoop()


if __name__ == "__main__":
    main()
