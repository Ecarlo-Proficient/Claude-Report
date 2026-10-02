"""
presence.py - the owner is at THIS Mac right now: Touch ID (or the Mac login password) before a QBO money write.

`confirm(reason)` shows macOS' own dialog with `reason` in it (the vendor count, total and bank account), so a
web page can neither fake it nor click through it. The Swift helper (presence.swift) is compiled on first use into
~/Library/Application Support/Proficient/bin/, rebuilt when its source changes. Fails CLOSED: no helper, no
compiler, no answer in 2 minutes, cancelled -> not confirmed.

    python ledger/presence.py "test"      prompts once, prints the answer
"""
from __future__ import annotations

import hashlib
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
SRC = HERE / "presence.swift"
BIN_DIR = Path.home() / "Library" / "Application Support" / "Proficient" / "bin"


NAME = "Presence Requested Vendor Pay Approval"      # macOS names the asker by its file name: "<NAME> is trying to ..."


def _binary() -> Path:
    tag = hashlib.sha256(SRC.read_bytes()).hexdigest()[:12]
    exe = BIN_DIR / f"presence-{tag}" / NAME
    if not exe.exists():
        import shutil                               # noqa: PLC0415
        BIN_DIR.mkdir(parents=True, exist_ok=True)
        for old in BIN_DIR.glob("presence-*"):
            shutil.rmtree(old, ignore_errors=True) if old.is_dir() else old.unlink(missing_ok=True)
        exe.parent.mkdir(parents=True, exist_ok=True)
        r = subprocess.run(["/usr/bin/swiftc", "-O", "-o", str(exe), str(SRC)], capture_output=True, text=True, timeout=300)
        if r.returncode != 0:
            raise RuntimeError(f"could not build the Touch ID check: {r.stderr.strip()[:300]}")
    return exe


def confirm(reason: str) -> tuple:
    """(True, "") when the Mac's owner confirmed; (False, why) otherwise. Never raises."""
    try:
        r = subprocess.run([str(_binary()), reason[:900]], capture_output=True, text=True, timeout=150)
    except Exception as e:                          # noqa: BLE001
        return False, f"Touch ID check unavailable: {e}"
    if r.returncode == 0:
        return True, ""
    return False, (r.stderr.strip() or "not confirmed")


if __name__ == "__main__":
    ok, why = confirm(sys.argv[1] if len(sys.argv) > 1 else "Project Ledger - Touch ID test")
    print("confirmed" if ok else f"NOT confirmed: {why}")
    sys.exit(0 if ok else 1)
