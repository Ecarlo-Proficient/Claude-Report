#!/usr/bin/env python3
"""
mirror_parity.py - PROOF that the raw QBO mirror answers every reader the way
QBO does. For each (entity, WHERE) shape the tools actually use, pull the rows
LIVE from QBO and from the mirror, and compare {Id: SyncToken}. Any row missing,
extra, or at a different version fails the run. Read-only; refreshes the mirror
first so both sides describe the same instant.

    python3 one-offs/mirror_parity.py            every shape
    python3 one-offs/mirror_parity.py --quick    the small entities only
"""
from __future__ import annotations

import argparse
import datetime as dt
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from shared import qbo_api                 # noqa: E402
from shared import qbo_mirror as mirror    # noqa: E402

Y = "2026-01-01"
TODAY = dt.date.today().isoformat()
SHAPES = [
    ("Account", ""), ("Account", "Active=true"),
    ("Account", "AccountType IN ('Bank','Credit Card') AND Active=true"),
    ("Account", "Classification = 'Liability'"),
    ("Vendor", ""), ("Vendor", "DisplayName LIKE '%Concrete%'"),
    ("Customer", ""), ("Customer", "Active = false"), ("Customer", "Active IN (true, false)"),
    ("Item", ""), ("Class", ""), ("Class", "Active = false"),
    ("Term", "Active IN (true, false)"), ("PaymentMethod", ""), ("PurchaseOrder", ""),
    ("Bill", "Balance > '0'"), ("Bill", f"Balance = '0' AND TxnDate >= '{Y}'"),
    ("Bill", f"TxnDate >= '{Y}'"), ("Purchase", f"TxnDate >= '{Y}'"),
    ("Invoice", ""), ("Invoice", "Balance > '0'"),
    ("Invoice", f"TxnDate >= '2025-01-01' AND TxnDate <= '{TODAY}'"),
    ("Invoice", "DocNumber IN ('34432','34520','34582')"),
    ("Payment", ""), ("Payment", f"TxnDate >= '{Y}'"),
    ("BillPayment", f"TxnDate >= '{Y}'"), ("JournalEntry", f"TxnDate >= '{Y}'"),
    ("CreditMemo", ""), ("VendorCredit", ""), ("Deposit", f"TxnDate >= '{Y}'"),
]
QUICK = {"Account", "Vendor", "Customer", "Item", "Class", "Term", "PaymentMethod", "PurchaseOrder"}


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--quick", action="store_true")
    ap.add_argument("--no-refresh", action="store_true")
    a = ap.parse_args()
    access = qbo_api.refresh_access()
    _, cid = qbo_api.load_credentials()
    if not a.no_refresh:
        con = mirror.connect()
        mirror.refresh(con, access, cid, progress=lambda s: None)
        con.close()
    shapes = [s for s in SHAPES if not a.quick or s[0] in QUICK]
    fails = 0
    for entity, where in shapes:
        t0 = time.time()
        live = {str(r["Id"]): str(r.get("SyncToken")) for r in qbo_api.query_all_live(access, cid, entity, where)}
        t1 = time.time()
        mine = {str(r["Id"]): str(r.get("SyncToken")) for r in mirror.query(entity, where)}
        t2 = time.time()
        missing = sorted(set(live) - set(mine))
        extra = sorted(set(mine) - set(live))
        stale = sorted(i for i in set(live) & set(mine) if live[i] != mine[i])
        bad = bool(missing or extra or stale)
        fails += bad
        flag = "FAIL" if bad else "ok  "
        print(f"{flag} {entity:13s} {where[:52]:52s} live {len(live):>6,} ({t1-t0:4.1f}s)  "
              f"mirror {len(mine):>6,} ({t2-t1:4.1f}s)"
              + (f"  missing {len(missing)} extra {len(extra)} stale {len(stale)}" if bad else ""))
        for label, ids in (("missing", missing), ("extra", extra), ("stale", stale)):
            if ids:
                print(f"       {label}: {ids[:12]}{' ...' if len(ids) > 12 else ''}")
    print(f"\n{len(shapes) - fails}/{len(shapes)} shapes identical")
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
