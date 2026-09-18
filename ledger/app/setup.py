"""py2app build config for Project Ledger.app.

Build via ../build_ledger_app.command (alias mode). A custom icon is picked up
automatically if ledger/app/app_icon.icns exists (the builder makes it from
app_icon.png when present).
"""
import os

from setuptools import setup

HERE = os.path.dirname(os.path.abspath(__file__))

_repo = os.path.dirname(os.path.dirname(HERE))   # HERE=<repo>/ledger/app → <repo>

_plist = {
    "CFBundleName": "Project Ledger",
    "CFBundleDisplayName": "Project Ledger",
    "CFBundleIdentifier": "local.proficient.ledger-dashboard",
    "CFBundleShortVersionString": "1.0",
    "CFBundleVersion": "1.0",
    "LSUIElement": False,          # show in the Dock (this IS the on/off indicator)
    "LSMinimumSystemVersion": "10.13",
    # Baked at build time so the app finds the repo no matter where the .app lives.
    "LSEnvironment": {"ACB_LEDGER_REPO": _repo},
}

_options = {"plist": _plist, "argv_emulation": False}
_icon = os.path.join(HERE, "app_icon.icns")
if os.path.exists(_icon):
    _options["iconfile"] = _icon

setup(
    app=["ledger_app.py"],
    options={"py2app": _options},
    setup_requires=["py2app"],
)
