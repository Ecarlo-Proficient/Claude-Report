#!/usr/bin/env python3
"""
reapply_payment.py - THE write behind the ledger's Payments QBO changed -> Write to QuickBooks: put a client payment
back on the invoices and credit memos QuickBooks took it off (owner 2026-10-16). The plan, guards and the check to the
cent are shared/payment_reapply.py; only the POST lives here, outside shared/, because shared/ is copied to the office
server and the server must not be able to write to QuickBooks (tests/test_office_server_image.py).

    ledger/reapply_payment.py 1316642            dry run: what would go back on (live read)
    ledger/reapply_payment.py 1316642 --commit   write it, then check every line and balance
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import requests  # noqa: E402

from shared import payment_reapply, qbo_api  # noqa: E402


def post(access: str, cid: str, ent: str, body: dict) -> dict:
    r = requests.post(f"{qbo_api.API_BASE}/v3/company/{cid}/{ent.lower()}",
                      headers={"Authorization": f"Bearer {qbo_api._CURRENT_ACCESS or access}",
                               "Content-Type": "application/json", "Accept": "application/json"},
                      params={"minorversion": qbo_api.MINOR_VERSION}, json=body, timeout=120)
    if r.status_code != 200:
        try:
            e = r.json()["Fault"]["Error"][0]
            msg = f"{e.get('Message')}: {e.get('Detail')}"
        except Exception:                                   # noqa: BLE001
            msg = r.text[:400]
        raise RuntimeError(f"QuickBooks refused the {ent} update ({r.status_code}): {msg} - nothing was changed")
    return r.json()[ent]


def dry_run(payment_id: str) -> dict:
    return payment_reapply.dry_run(payment_id)


def commit(payment_id: str, sync_token: str, n: int, amount: float) -> dict:
    """Write only the plan the owner confirmed (same payment version, count and amount)."""
    return payment_reapply.commit(post, payment_id, sync_token, n, amount)


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("payment_id")
    ap.add_argument("--commit", action="store_true")
    a = ap.parse_args()
    d = dry_run(a.payment_id)
    print(f"Payment #{d['ref']} {d['date']} {d['customer']} ${d['total']:,.2f}")
    for b in d["back"]:
        print(f"   put back {b['type']:<10} {b['doc']:<12} ${b['amount']:>10,.2f}   open {b['open_now']:,.2f} -> {b['open_after']:,.2f}")
    for w in d["why"]:
        print("  ", w)
    if not a.commit or not d["ok"]:
        return 0 if d["ok"] or d["why"] == [payment_reapply.NOTHING] else 1
    res = commit(a.payment_id, d["sync_token"], d["n"], d["amount"])
    print("Written and checked." if res["ok"] else "NOT BACK: " + res.get("error", ""))
    return 0 if res["ok"] else 1


if __name__ == "__main__":
    sys.exit(main())
