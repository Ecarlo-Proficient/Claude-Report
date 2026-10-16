#!/usr/bin/env python3
"""
payment_drift.py - client payments QBO quietly stripped. The ledger's "Payments QBO changed"
audit (Company, beside Checks QBO changed) - /api/paymentdrift. The customer-side twin of
check_drift.py (owner 2026-10-16: "if any client payment gets offed we need to fix JUST like
the bill payments getting unapplied").

Read-only on the raw QBO mirror (shared/qbo_mirror). No QBO call, no QBO write.

How it happens: change the CUSTOMER on a credit memo (moving a backcharge from the parent onto
the project) and QBO takes the credit off the payment that used it. The payment keeps its total,
so QBO takes the same amount off an invoice line instead - that invoice reopens. QBO does NOT
bump the payment's SyncToken when it does this, so the mirror keeps the payment as it was: the
mirror copy IS the before copy. Worked cases 10/16/2026: BKGP23052 on check #0087532 (invoice
34527 open $240), BKGP23026 on check #0086618 (invoice 34461 open $100).

The test: for every invoice / credit memo a payment applies to, its open balance in the mirror
should be at most its total less everything the mirror's payments apply to it. More open than
that = some payment lost it. A row per payment, with each document that came off.

    python3 ledger/payment_drift.py               every stripped payment on file
    python3 ledger/payment_drift.py --payment ID  one payment in full
"""
from __future__ import annotations

import argparse
import sys
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from shared import qbo_mirror as mirror    # noqa: E402

EPS = 0.005
MIN_GAP = 0.01
DOCS = ("Invoice", "CreditMemo")


def _fmt(d) -> str:
    s = str(d or "")[:10]
    return f"{s[5:7]}/{s[8:10]}/{s[:4]}" if len(s) == 10 else s


def links(pay: dict) -> list:
    """[(TxnType, TxnId, amount)] - what the client payment applies, line by line."""
    return [(lt["TxnType"], str(lt["TxnId"]), round(float(ln.get("Amount") or 0), 2))
            for ln in pay.get("Line") or [] for lt in ln.get("LinkedTxn") or []]


def open_of(rec: dict) -> float:
    v = rec.get("RemainingCredit") if "RemainingCredit" in rec else rec.get("Balance")
    return round(float(v or 0), 2)


def audit(con=None, payment: str | None = None) -> dict:
    own = con is None
    con = con or mirror.connect()
    try:
        pays = mirror.load("Payment", con=con)
        applied = defaultdict(float)            # (type, id) -> what every payment in the mirror puts on it
        for p in pays:
            for t, i, a in links(p):
                if t in DOCS:
                    applied[(t, i)] += a
        docs = {(t, r["Id"]): r for t in DOCS for r in mirror.load(t, con=con)}   # deleted ones left out
        rows = []
        for p in pays:
            if payment and p["Id"] != str(payment):
                continue
            off = []
            for t, i, a in links(p):
                rec = docs.get((t, i))
                if not rec or t not in DOCS:
                    continue
                total = round(float(rec.get("TotalAmt") or 0), 2)
                expect = round(max(total - applied[(t, i)], 0), 2)
                gap = round(open_of(rec) - expect, 2)
                if gap >= MIN_GAP:
                    off.append({"type": t, "id": i, "doc": rec.get("DocNumber") or i,
                                "customer": (rec.get("CustomerRef") or {}).get("name") or "",
                                "date": rec.get("TxnDate"), "total": total, "applied": a,
                                "open": open_of(rec), "gap": min(gap, a),
                                "updated": (rec.get("MetaData") or {}).get("LastUpdatedTime")})
            if off:
                rows.append({"payment_id": p["Id"], "ref": p.get("PaymentRefNum") or "", "date": p.get("TxnDate"),
                             "customer": (p.get("CustomerRef") or {}).get("name") or "",
                             "total": round(float(p.get("TotalAmt") or 0), 2), "sync_token": p.get("SyncToken"),
                             "off": off, "amount_off": round(max(sum(o["gap"] for o in off if o["type"] == t) for t in DOCS), 2),
                             "credits_off": [o["doc"] for o in off if o["type"] == "CreditMemo"],
                             "invoices_reopened": [o["doc"] for o in off if o["type"] == "Invoice"]})
        rows.sort(key=lambda r: max((o["updated"] or "") for o in r["off"]), reverse=True)
        last = None
        try:
            last = mirror._meta_get(con, "last_refresh")
        except Exception:                                   # noqa: BLE001 - a faked mirror in tests has no meta
            pass
        return {"ok": True, "rows": rows, "count": len(rows), "last_refresh": last,
                "amount": round(sum(r["amount_off"] for r in rows), 2)}
    finally:
        if own:
            con.close()


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--payment")
    a = ap.parse_args()
    res = audit(payment=a.payment)
    print(f"{res['count']} client payment(s) QBO stripped, ${res['amount']:,.2f} off")
    for r in res["rows"]:
        print(f"\n#{r['ref'] or '-'} {_fmt(r['date'])} {r['customer']} ${r['total']:,.2f} (payment {r['payment_id']})")
        for o in r["off"]:
            print(f"   {o['type']:<10} {o['doc']:<12} {o['customer']:<22} applied ${o['applied']:>10,.2f}"
                  f"  open now ${o['open']:>10,.2f}  off ${o['gap']:,.2f}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
