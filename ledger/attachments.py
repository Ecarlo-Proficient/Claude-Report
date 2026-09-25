#!/usr/bin/env python3
"""Resolve a bill's QBO attachment links and print them as JSON.

    {"ok": true, "files": [{"name": "...", "url": "<fresh TempDownloadUri>"}, ...]}

The ledger dashboard SUBPROCESSES this (never imports it - tools never import tools) to
turn a bill txnId into fresh, minutes-lived download links, so the owner opens the scan
without going into QBO. Uses the shared attachable index (the disk cache reused from the
P&L); a fresh link is fetched per file at call time. Read-only on QBO.

Only the JSON goes to stdout (no realm, no secrets). A missing index is reported, not
built here - the sweep is slow, so it stays an explicit P&L/console step.
"""
import argparse
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from shared import qbo_attachments  # noqa: E402


def main() -> int:
    ap = argparse.ArgumentParser(description="Resolve a bill's QBO attachment links (JSON to stdout).")
    ap.add_argument("txn_id")
    ap.add_argument("--type", default="Bill", help="Bill (default) / Expense / Purchase")
    a = ap.parse_args()

    if not str(a.txn_id).isdigit():
        print(json.dumps({"ok": False, "error": "bad txn id"}))
        return 0
    idx = qbo_attachments.index_from_cache()
    if idx is None:
        print(json.dumps({"ok": False, "error": "attachment index not built yet - run a P&L export to build it"}))
        return 0
    try:
        from shared.qbo_api import load_credentials, _api_get
        access, company_id = load_credentials()
        files = qbo_attachments.fresh_links(access, company_id, idx, a.txn_id, _api_get, a.type)
        print(json.dumps({"ok": True, "files": files}))
    except Exception as e:  # noqa: BLE001
        print(json.dumps({"ok": False, "error": str(e)}))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
