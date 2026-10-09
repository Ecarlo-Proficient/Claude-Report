#!/usr/bin/env python3
"""
keyhelper/migrate_keys.py - move the last keys out of the older Keychain store into the key library.

Until 09/29/2026 the Notion token and the Teams paid-notice webhook lived in the `keyring` store
(service `proficient-automation-worker`), readable by programs outside Key Helper. This moves each one into
the key library (behind Key Helper) under its library name, reads it back THROUGH the helper to verify, and
only then removes the old item. A key that does not verify is left where it was.

    python3 keyhelper/migrate_keys.py            dry run: which old items exist (names only, no values)
    python3 keyhelper/migrate_keys.py --apply    move them (macOS may ask once for the old item - click
                                                 Allow, NOT Always Allow; saving asks Touch ID / password)

Needs the adoption done (`bash keyhelper/install.sh --adopt`) and the names registered
(`bash keyhelper/install.sh --register`).
"""
from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from shared import key_broker, qbo_vault as kc  # noqa: E402


def _old_item_exists(account: str) -> bool:
    """Attributes only - never reads the secret, never prompts."""
    r = subprocess.run(["/usr/bin/security", "find-generic-password", "-s", kc.LEGACY_SERVICE, "-a", account],
                       capture_output=True)
    return r.returncode == 0


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--apply", action="store_true", help="move the keys (default: dry run)")
    a = ap.parse_args()

    if not key_broker.active():
        print("✗ the key library is not adopted yet - run: bash keyhelper/install.sh --adopt")
        return 1

    present = {name: acct for name, acct in kc.LEGACY_ACCOUNTS.items() if _old_item_exists(acct)}
    for name, acct in kc.LEGACY_ACCOUNTS.items():
        print(f"  {kc.LEGACY_SERVICE}/{acct:<24} -> {name:<24} {'to move' if name in present else 'not stored'}")
    if not present:
        print("nothing to move")
        return 0
    if not a.apply:
        print("dry run - add --apply to move")
        return 0

    import keyring
    values = {}
    for name, acct in present.items():
        v = (keyring.get_password(kc.LEGACY_SERVICE, acct) or "").strip()
        if v:
            values[name] = v
    if not values:
        print("✗ the old items could not be read (dialog denied?) - nothing changed")
        return 1

    kc.put_all(values)                                   # one Touch ID / password for the write
    moved = []
    for name, v in values.items():
        if key_broker.key(name) != v:
            print(f"  ✗ {name}: the library copy did not verify - the old item is left in place")
            continue
        try:
            keyring.delete_password(kc.LEGACY_SERVICE, present[name])
        except Exception as e:  # noqa: BLE001
            print(f"  ! {name}: moved and verified, but the old item could not be removed ({type(e).__name__})")
            continue
        moved.append(name)
        print(f"  ✓ {name}: moved, verified through Key Helper, old item removed")
    values.clear()
    return 0 if len(moved) == len(present) else 1


if __name__ == "__main__":
    sys.exit(main())
