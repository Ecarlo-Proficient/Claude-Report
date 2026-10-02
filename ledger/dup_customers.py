#!/usr/bin/env python3
"""
dup_customers.py - one project # set up as two QuickBooks customers. The ledger's
"Duplicate customers" audit (Company, beside Uncleared checks) - /api/dupcustomers.

Read-only on the raw QBO mirror (shared/qbo_mirror). No QBO call, no QBO write.

Every tool keys a job by its project # (shared/qbo_api.build_project_customer_map)
and can only use ONE customer per number. When QBO holds two active customers
whose names carry the same project #, the tools keep one
(shared/qbo_api.pick_customer - exact name, most invoices, oldest) and anything
posted to the other is invisible to them - a loose top-level customer held
old bills no P&L or WIP ever counted (found 09/30/2026).
sync-all prints the same pairs as "duplicate customers for ..." lines.

Status per project (the worst of its non-kept records):
  Both have invoices    two customers each carry invoices - someone picks the right builder
  Money on duplicate    the record the tools skip carries invoices / bills / payments
  Name typo             'RP7340 -FTW' - a space hides the job's own suffix; rename, don't retire
  Separate job?         'RP7152-1' - a suffix the project # reader does not know
  Empty duplicate       nothing on it - make it inactive
A pair leaves the page when the extra customer is made inactive in QBO (the
mirror only serves active customers) and the mirror is refreshed.

The grouping is shared/qbo_trust.duplicate_groups - the SAME code the WIP
update's QuickBooks trust gate holds jobs on, so this page and the gate agree.

    python3 ledger/dup_customers.py        print the list
"""
from __future__ import annotations

import sys
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from shared import qbo_mirror as mirror    # noqa: E402
from shared import qbo_trust               # noqa: E402

ORDER = qbo_trust.STATUS_ORDER
_CACHE: dict = {}          # {stamp -> result}: the mirror only changes on a refresh


def audit(con=None) -> dict:
    own = con is None
    con = con or mirror.connect()
    try:
        st = mirror.stamp(con)
        key = st.isoformat() if st else ""
        if key and key in _CACHE:
            return _CACHE[key]
        groups = qbo_trust.duplicate_groups(con)
        counts = defaultdict(int)
        for g in groups:
            counts[g["status"]] += 1
        res = {"ok": True, "groups": groups, "counts": dict(counts), "statuses": ORDER[:-1], "loaded_at": key or None}
        if key:
            _CACHE.clear()
            _CACHE[key] = res
        return res
    finally:
        if own:
            con.close()


def main() -> int:
    res = audit()
    print(f"{len(res['groups'])} project #s on two or more QBO customers (mirror as of {res['loaded_at']})")
    for g in res["groups"]:
        print(f"\n{g['project']} - {g['status']}")
        for r in g["customers"]:
            inv = ", ".join(d["doc"] or d["id"] for d in r["invoices"]) or "-"
            bills = ", ".join(d["doc"] or d["id"] for d in r["bills"]) or "-"
            print(f"   {'*' if r['used'] else ' '} {r['name']} (Id {r['id']}) - {r['status']}; "
                  f"invoices {inv} {r['invoice_total']:,.2f}; bills {bills} {r['bill_total']:,.2f}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
