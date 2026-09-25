#!/usr/bin/env python3
"""download_attachments.py - save a set of bills' QBO scans to a folder (JSON to stdout).

The Audit tab's "Download scans" hands the flagged bills to the responsible party. This
resolves each bill's FRESH `TempDownloadUri` (minutes-lived) and writes the file(s) into a
local folder, named "<Bill#> <Vendor> - <original>" so the scans line up with the copied
table. The ledger dashboard SUBPROCESSES this (never imports it - tools never import tools),
then reveals the folder in the OS file manager so the owner can drag the scans into the message.

Manifest is JSON on stdin (or --manifest FILE):
    [{"txnId": "123", "bill_no": "5567", "vendor": "COWTOWN", "type": "Bill"}, ...]
Prints exactly one JSON line:
    {"ok": true, "folder": "...", "count": N, "bills": M, "errors": [...]}

Read-only on QBO. Uses the shared attachable index (the disk cache reused from the P&L); a
fresh link is fetched per file at call time. Only the JSON goes to stdout (no realm, no secrets).
"""
import argparse
import json
import os
import re
import sys
import urllib.request
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from shared import qbo_attachments  # noqa: E402

_UNSAFE = re.compile(r'[<>:"/\\|?*\x00-\x1f]+')       # characters no OS allows in a filename


def _safe(s: str, limit: int = 90) -> str:
    """A filename-safe slice of a label (bill #, vendor, original name)."""
    s = _UNSAFE.sub("_", (s or "").strip()).strip(". ")
    return s[:limit].strip() or "file"


def _download(url: str, dest: Path) -> int:
    """GET a pre-signed QBO temp URL and write the bytes. https only (never file://)."""
    if not str(url).lower().startswith("https://"):
        raise ValueError("non-https url")
    req = urllib.request.Request(url, headers={"User-Agent": "ledger-audit/1"})
    with urllib.request.urlopen(req, timeout=90) as r:  # noqa: S310 - pre-signed QBO https temp URL
        data = r.read()
    dest.write_bytes(data)
    return len(data)


def _unique(dest_dir: Path, name: str, used: set) -> Path:
    """A path in dest_dir that collides with nothing already used or on disk."""
    p = dest_dir / name
    stem, ext = os.path.splitext(name)
    i = 2
    while str(p) in used or p.exists():
        p = dest_dir / f"{stem} ({i}){ext}"
        i += 1
    return p


def _run(bills: list, dest_dir: Path) -> dict:
    idx = qbo_attachments.index_from_cache()
    if idx is None:
        return {"ok": False, "error": "attachment index not built yet - run a P&L export to build it"}
    from shared.qbo_api import _api_get, load_credentials
    try:
        access, company_id = load_credentials()
    except Exception as e:                              # noqa: BLE001
        return {"ok": False, "error": f"auth failed: {e}"}

    dest_dir.mkdir(parents=True, exist_ok=True)
    count, bills_with, errors, used = 0, 0, [], set()
    for b in bills:
        txn = str((b or {}).get("txnId") or "").strip()
        if not txn.isdigit():
            continue
        label = _safe(f"{b.get('bill_no') or txn} {b.get('vendor') or ''}".strip())
        try:
            files = qbo_attachments.fresh_links(access, company_id, idx, txn, _api_get, b.get("type") or "Bill")
        except Exception as e:                          # noqa: BLE001
            errors.append(f"{label}: {e}")
            continue
        got = 0
        for f in files:
            p = _unique(dest_dir, f"{label} - {_safe(f.get('name') or 'scan')}", used)
            try:
                _download(f["url"], p)
                used.add(str(p))
                count += 1
                got += 1
            except Exception as e:                      # noqa: BLE001 - one bad file shouldn't sink the rest
                errors.append(f"{label}/{f.get('name')}: {e}")
        if got:
            bills_with += 1
    return {"ok": True, "folder": str(dest_dir), "count": count, "bills": bills_with, "errors": errors[:20]}


def main() -> int:
    ap = argparse.ArgumentParser(description="Download bills' QBO scans to a folder (JSON to stdout).")
    ap.add_argument("--dest", required=True, help="destination folder (created if missing).")
    ap.add_argument("--manifest", default="-", help="JSON manifest file, or - for stdin (default).")
    a = ap.parse_args()

    raw = sys.stdin.read() if a.manifest == "-" else Path(a.manifest).read_text(encoding="utf-8")
    try:
        bills = json.loads(raw or "[]")
    except json.JSONDecodeError:
        print(json.dumps({"ok": False, "error": "bad manifest"}))
        return 0
    if not isinstance(bills, list) or not bills:
        print(json.dumps({"ok": False, "error": "no bills"}))
        return 0
    print(json.dumps(_run(bills, Path(a.dest).expanduser())))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
