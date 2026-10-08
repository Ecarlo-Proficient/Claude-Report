#!/usr/bin/env python3
"""carry_history.py - switch-day step: carry the Mac's cost-code miscode history to the office server (10/08/2026).

Each side keeps its own `cost_code_history.json` (the Bill Tracker's `Audit - History`: when a miscoded bill was
first seen, how many runs, OPEN / FIXED). The Mac's is the long record (since 09/01); the server's started with the
test week. When the server becomes the writer its record must continue the Mac's, or every count restarts.

The merge: the Mac's entries and run log, plus any entry only the server has; for an entry on both, the earliest
first-seen and the newest state (last seen, status, fixed / reopened dates). The server's own run log is dropped -
those were test-week runs every 15 minutes and would swamp the "new per run" rate.

Dry run by default (prints what it would write). `--commit` backs up the server's copy beside it, then writes. It
refuses while the server's AP job is mid-run (it writes this file). Run it on the Mac, right before the writer file
is set to "server" (README step 8):

    ~/.venvs/proficient/bin/python docker/carry_history.py --commit
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import shutil
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))
from shared import paths  # noqa: E402

NAME = "cost_code_history.json"
STATE_FIELDS = ("last_seen", "status", "fixed_on", "reopened_on", "amount", "project")


def merge(mac: dict, server: dict) -> dict:
    out = {"schema": mac.get("schema", server.get("schema")), "runs": list(mac.get("runs") or []),
           "entries": dict(mac.get("entries") or {})}
    for key, s in (server.get("entries") or {}).items():
        m = out["entries"].get(key)
        if m is None:
            out["entries"][key] = s
            continue
        m = dict(m)
        m["first_seen"] = min(m.get("first_seen") or s.get("first_seen") or "", s.get("first_seen") or m["first_seen"])
        if (s.get("last_seen") or "") > (m.get("last_seen") or ""):
            for f in STATE_FIELDS:
                if f in s:
                    m[f] = s[f]
        out["entries"][key] = m
    return out


def _server_ap_running(share: Path) -> bool:
    try:
        ap = json.loads((share / "_automation" / "server-status.json").read_text())["jobs"]["ap"]
    except (OSError, ValueError, KeyError):
        return False
    return (ap.get("last_run") or "") > (ap.get("last_ok") or "") and ap.get("exit") in (0, None)


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("--commit", action="store_true", help="write the merged file to the server share (default: dry run)")
    args = ap.parse_args()

    mac_file = paths.register_file(NAME)
    share = paths.accounting_base()
    server_file = share / "_automation" / "registers" / NAME
    if not share.is_dir():
        print(f"carry_history: the Accounting share is not mounted ({share}) - connect it and run again")
        return 2
    mac = json.loads(mac_file.read_text())
    server = json.loads(server_file.read_text()) if server_file.exists() else {"entries": {}, "runs": []}
    out = merge(mac, server)

    open_ct = sum(1 for e in out["entries"].values() if e.get("status") == "open")
    only_server = len(set(server.get("entries") or {}) - set(mac.get("entries") or {}))
    print(f"Mac: {len(mac.get('entries') or {})} entries, {len(mac.get('runs') or [])} runs  |  "
          f"server: {len(server.get('entries') or {})} entries ({only_server} only on the server)")
    print(f"merged: {len(out['entries'])} entries, {len(out['runs'])} runs, {open_ct} open")
    if not args.commit:
        print("dry run - nothing written. Add --commit to write it to the server share.")
        return 0
    if _server_ap_running(share):
        print("carry_history: the server's AP job is running right now (it writes this file) - wait a minute, run again")
        return 3
    if server_file.exists():
        backup = server_file.with_name(f"{server_file.stem}.before-carry-{dt.datetime.now():%m-%d-%Y-%H%M}.json")
        shutil.copy2(server_file, backup)
        print(f"backup: {backup.name}")
    tmp = server_file.with_suffix(".tmp")
    tmp.write_text(json.dumps(out, indent=1, sort_keys=True))
    os.replace(tmp, server_file)
    print(f"written: {server_file}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
