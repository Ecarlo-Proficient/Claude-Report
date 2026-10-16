#!/usr/bin/env python3
"""QuickBooks' printed copy of one invoice, for the ledger's invoice viewer (owner 2026-10-11: "i want that
to open just like attachment and see the invoice load from qbo ... so you don't have to leave the ledger").

    python3 ledger/invoice_pdf.py <invoice id> [--sync <SyncToken>]
    -> {"ok": true, "path": "<cache>/<id>-<sync>.pdf"}     (the JSON is the last stdout line)

The dashboard SUBPROCESSES this (never imports it), like attachments.py. The PDF is cached outside the repo
(~/Library/Application Support/Proficient/invoice-pdf/), one file per invoice VERSION: the mirror's SyncToken
is in the name, so an invoice edited in QuickBooks is fetched again and the old copy dropped. Read-only on
QuickBooks; no realm or secret ever printed.
"""
import argparse
import json
import os
import sys
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

CACHE = Path.home() / "Library" / "Application Support" / "Proficient" / "invoice-pdf"


def cached(txn_id: str, sync: str) -> Path:
    return CACHE / f"{txn_id}-{sync or '0'}.pdf"


def main() -> int:
    ap = argparse.ArgumentParser(description="QuickBooks' PDF of one invoice (JSON to stdout).")
    ap.add_argument("txn_id")
    ap.add_argument("--sync", default="0", help="the invoice's SyncToken (the cache key)")
    a = ap.parse_args()
    if not str(a.txn_id).isdigit() or not str(a.sync).isdigit():
        print(json.dumps({"ok": False, "error": "bad invoice id"}))
        return 0
    out = cached(a.txn_id, a.sync)
    if out.exists() and out.stat().st_size > 0:
        print(json.dumps({"ok": True, "path": str(out), "cached": True}))
        return 0
    try:
        from shared.qbo_api import get_form_pdf
        pdf = get_form_pdf("invoice", a.txn_id)
    except Exception as e:  # noqa: BLE001 - the reason goes back to the viewer
        print(json.dumps({"ok": False, "error": str(e)}))
        return 0
    CACHE.mkdir(parents=True, exist_ok=True)
    for old in CACHE.glob(f"{a.txn_id}-*.pdf"):   # an older version of the same invoice
        old.unlink(missing_ok=True)
    tmp = out.with_suffix(".part")
    tmp.write_bytes(pdf)
    tmp.replace(out)
    print(json.dumps({"ok": True, "path": str(out), "cached": False}))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
