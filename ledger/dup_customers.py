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

    python3 ledger/dup_customers.py        print the list
"""
from __future__ import annotations

import re
import sys
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from shared import qbo_mirror as mirror                   # noqa: E402
from shared.qbo_api import extract_proj, pick_customer    # noqa: E402

COST_ENTITIES = ("Bill", "Purchase", "JournalEntry", "VendorCredit")
SALES_ENTITIES = ("Invoice", "Payment", "Estimate", "CreditMemo", "SalesReceipt")
# worst first - a project takes the status of its worst record
ORDER = ["Both have invoices", "Money on duplicate", "Name typo", "Separate job?", "Empty duplicate", "Real"]

_CACHE: dict = {}          # {stamp -> result}: the mirror only changes on a refresh


def _day(s) -> str:
    return str(s or "")[:10]


def _status(proj: str, c: dict, keep_id: str, busy: bool, tie: bool) -> str:
    dn = (c.get("DisplayName") or "").strip()
    squashed = re.sub(r"\s+", "", dn)
    if tie:
        return "Both have invoices"
    if dn.upper() != proj.upper():
        if squashed.upper() != proj.upper() and extract_proj(squashed) not in (None, proj):
            return "Name typo"
        if squashed.upper() != proj.upper():
            return "Separate job?"
    if c["Id"] == keep_id:
        return "Real"
    return "Money on duplicate" if busy else "Empty duplicate"


def audit(con=None) -> dict:
    own = con is None
    con = con or mirror.connect()
    try:
        st = mirror.stamp(con)
        key = st.isoformat() if st else ""
        if key and key in _CACHE:
            return _CACHE[key]
        res = _audit(con)
        res["loaded_at"] = key or None
        if key:
            _CACHE.clear()
            _CACHE[key] = res
        return res
    finally:
        if own:
            con.close()


def _audit(con) -> dict:
    groups: dict = defaultdict(list)
    for c in mirror.query("Customer", "", con=con):          # active only - an inactive duplicate is fixed
        p = extract_proj(c.get("DisplayName") or c.get("CompanyName") or "")
        if p:
            groups[p].append(c)
    dups = {p: cs for p, cs in groups.items() if len(cs) > 1}
    ids = {c["Id"] for cs in dups.values() for c in cs}
    sales: dict = defaultdict(lambda: defaultdict(list))     # cid -> entity -> [doc]
    costs: dict = defaultdict(dict)                          # cid -> {(entity, id): doc}
    if ids:
        for ent in SALES_ENTITIES:
            for r in mirror.load(ent, con=con):
                cid = (r.get("CustomerRef") or {}).get("value")
                if cid in ids:
                    sales[cid][ent].append({"id": r["Id"], "doc": r.get("DocNumber") or "", "date": _day(r.get("TxnDate")),
                                            "amount": round(float(r.get("TotalAmt") or 0), 2)})
        for ent in COST_ENTITIES:
            sign = -1 if ent == "VendorCredit" else 1
            for r in mirror.load(ent, con=con):
                for ln in r.get("Line") or []:
                    for det in ln.values():
                        if not (isinstance(det, dict) and "CustomerRef" in det):
                            continue
                        cid = (det.get("CustomerRef") or {}).get("value")
                        if cid not in ids:
                            continue
                        doc = costs[cid].setdefault((ent, r["Id"]), {
                            "type": ent, "id": r["Id"], "doc": r.get("DocNumber") or "", "date": _day(r.get("TxnDate")),
                            "vendor": (r.get("VendorRef") or r.get("EntityRef") or {}).get("name", ""), "amount": 0.0, "lines": 0})
                        doc["amount"] = round(doc["amount"] + sign * float(ln.get("Amount") or 0), 2)
                        doc["lines"] += 1
    out = []
    for p in sorted(dups):
        cs = dups[p]
        inv_n = {c["Id"]: len(sales[c["Id"]]["Invoice"]) for c in cs}
        keep = pick_customer(p, cs, inv_n)
        tie = sum(1 for c in cs if inv_n[c["Id"]]) > 1
        recs = []
        for c in sorted(cs, key=lambda c: (c["Id"] != keep["Id"], int(c["Id"]))):
            cid = c["Id"]
            s = sales[cid]
            bills = sorted(costs[cid].values(), key=lambda d: d["date"], reverse=True)
            busy = bool(bills or any(s[e] for e in SALES_ENTITIES))
            recs.append({
                "id": cid, "name": c.get("FullyQualifiedName") or c.get("DisplayName") or "",
                "used": cid == keep["Id"], "status": _status(p, c, keep["Id"], busy, tie),
                "created": _day((c.get("MetaData") or {}).get("CreateTime")),
                "invoices": sorted(s["Invoice"], key=lambda d: d["date"], reverse=True),
                "invoice_total": round(sum(d["amount"] for d in s["Invoice"]), 2),
                "payments": len(s["Payment"]), "estimates": len(s["Estimate"]),
                "bills": bills, "bill_total": round(sum(d["amount"] for d in bills), 2),
            })
        others = [r for r in recs if not r["used"]] or recs
        status = min((r["status"] for r in recs if r["status"] != "Real"), key=ORDER.index, default="Real")
        out.append({"project": p, "status": status, "customers": recs,
                    "created": max(r["created"] for r in recs),
                    "money_on_duplicates": round(sum(r["bill_total"] + r["invoice_total"] for r in others), 2)})
    out.sort(key=lambda g: (ORDER.index(g["status"]), g["project"]))
    counts = defaultdict(int)
    for g in out:
        counts[g["status"]] += 1
    return {"ok": True, "groups": out, "counts": dict(counts), "statuses": ORDER[:-1]}


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
