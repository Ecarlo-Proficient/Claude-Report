"""
qbo_trust.py - is QuickBooks good to trust for a job's costs / billed to date?
The gate every WIP update passes BEFORE it overwrites a number (the owner
2026-09-30: "we need to first have a ci check that qbo is good to trust ... that
way we can trust that when we overwrite the costs/billed we can be certain these
are real projects booked in qbo").

Every WIP tool reads a job's costs and billed from ONE QuickBooks customer - the
one `shared/qbo_api.build_project_customer_map` picks for its project #. That is
only the whole truth when the job is booked cleanly. Read-only on the raw QBO
mirror (shared/qbo_mirror); no QBO call, no QBO write. The checks:

  STALE      the mirror is older than MAX_AGE_HOURS - every job is held (the
             customer map comes from it; a job booked since would not be seen)
  UNREADABLE the mirror could not be read at all - every job is held (fail closed)
  DUPLICATE  the project # sits on two active customers and the one the tools
             skip carries invoices / bills / payments - that money never reaches
             the WIP (RP7401-FTW: two builders, one invoice each). costs + billed
  NAME TYPO  a customer whose name only reads as this job with the spaces taken
             out ('RP7340 -FTW') - the job's own customer is invisible. costs + billed
  NOT CODED  a cost line with NO project that names this job (and only this job)
             in its description, class or memo - cost the P&L never sees. costs
  FW ON SLAB a flatwork (FW) cost code on the slab RP#### while its RP####-FTW
             twin exists - the slab reads high, the flatwork low. costs, both jobs

`assess(project_nums)` -> Trust: `holds[PN] = {field: [reason, ...]}` for the
fields that must not be overwritten ("costs" / "billed"), plus `global_reason`
when every job is held. The WIP review marks those cells HELD (never written,
never approvable); a held job not yet on the tab is not added. A hold clears by
itself once QuickBooks is fixed and the mirror refreshed.

    python3 shared/qbo_trust.py RP7401-FTW CP610 ...     print the holds
"""
from __future__ import annotations

import re
import sys
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterable, List, Optional

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from shared import job_rulings                                       # noqa: E402
from shared import qbo_mirror as mirror                              # noqa: E402
from shared.qbo_api import PROJ_RE, extract_proj, pick_customer      # noqa: E402

MAX_AGE_HOURS = mirror.MAX_AGE_HOURS
COST_ENTITIES = ("Bill", "Purchase", "JournalEntry", "VendorCredit")
SALES_ENTITIES = ("Invoice", "Payment", "Estimate", "CreditMemo", "SalesReceipt")
MIN_AMOUNT = 1.00          # cents of rounding are not a reason to hold a job
# NOT CODED / FW ON SLAB hold only from this much (the owner 2026-09-30, "$1,000+ holds"): holding keeps the WIP's
# OLDER number, which misses that line AND everything booked since - over a $12 line that is less accurate, not safer.
# Below it the job still updates and the cell carries a CHECK note with the documents. Duplicates / typos always hold.
HOLD_AMOUNT = 1000.00
STATUS_ORDER = ["Both have invoices", "Money on duplicate", "Name typo", "Separate job?", "Empty duplicate", "Real"]
_FW_RE = re.compile(r"^FW(51|52|[1-9])$", re.IGNORECASE)
_RP_BASE_RE = re.compile(r"^RP\d{4}$")


@dataclass
class Trust:
    holds: Dict[str, Dict[str, List[str]]] = field(default_factory=dict)
    checks: Dict[str, Dict[str, List[str]]] = field(default_factory=dict)   # below HOLD_AMOUNT: say it, still update
    global_reason: Optional[str] = None
    mirror_as_of: Optional[str] = None

    def reasons(self, pn: str, key: str) -> List[str]:
        """Why this job's `key` ('costs' / 'billed') must not be overwritten; [] = trusted."""
        if self.global_reason and key in ("costs", "billed"):
            return [self.global_reason]
        return list((self.holds.get(str(pn).strip().upper()) or {}).get(key) or [])

    def held(self, pn: str) -> bool:
        return bool(self.reasons(pn, "costs") or self.reasons(pn, "billed"))

    def notes(self, pn: str, key: str) -> List[str]:
        """Small findings that do NOT hold the number - shown beside it so someone fixes them."""
        return list((self.checks.get(str(pn).strip().upper()) or {}).get(key) or [])

    def add(self, pn: str, keys: Iterable[str], reason: str, hold: bool = True) -> None:
        d = (self.holds if hold else self.checks).setdefault(pn.upper(), {})
        for k in keys:
            if reason not in d.setdefault(k, []):
                d[k].append(reason)


def _money(v: float) -> str:
    return f"${v:,.2f}"


def _day(s) -> str:
    return str(s or "")[:10]


def _squashed_proj(name: str) -> Optional[str]:
    return extract_proj(re.sub(r"\s+", "", name or ""))


# ── duplicate customers: the ONE copy (the ledger's Duplicate customers page reads it too) ──
def customer_status(proj: str, c: dict, keep_id: str, busy: bool, tie: bool) -> str:
    dn = (c.get("DisplayName") or "").strip()
    squashed = re.sub(r"\s+", "", dn)
    if tie:
        return "Both have invoices"
    if dn.upper() != proj.upper():
        if squashed.upper() != proj.upper() and _squashed_proj(dn) not in (None, proj):
            return "Name typo"
        if squashed.upper() != proj.upper():
            return "Separate job?"
    if c["Id"] == keep_id:
        return "Real"
    return "Money on duplicate" if busy else "Empty duplicate"


def duplicate_groups(con) -> List[dict]:
    """Every project # carried by two or more ACTIVE customers: the one the tools use
    (pick_customer, the same rule as build_project_customer_map), and every invoice /
    payment / estimate and bill / expense / JE / vendor-credit line on each."""
    groups: dict = defaultdict(list)
    for c in mirror.query("Customer", "", con=con):           # active only - an inactive duplicate is fixed
        p = extract_proj(c.get("DisplayName") or c.get("CompanyName") or "")
        if p:
            groups[p].append(c)
    # A job the owner ruled split billing (shared/job_rulings `customers: all`, RP7401-FTW)
    # is not a duplicate: every reader counts all its customers, so it neither holds the
    # WIP nor shows on the ledger's Duplicate customers page.
    dups = {p: cs for p, cs in groups.items() if len(cs) > 1 and not job_rulings.all_customers(p)}
    ids = {c["Id"] for cs in dups.values() for c in cs}
    sales: dict = defaultdict(lambda: defaultdict(list))
    costs: dict = defaultdict(dict)
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
            cid, s = c["Id"], sales[c["Id"]]
            bills = sorted(costs[cid].values(), key=lambda d: d["date"], reverse=True)
            busy = bool(bills or any(s[e] for e in SALES_ENTITIES))
            recs.append({
                "id": cid, "name": c.get("FullyQualifiedName") or c.get("DisplayName") or "",
                "display_name": (c.get("DisplayName") or "").strip(),
                "used": cid == keep["Id"], "status": customer_status(p, c, keep["Id"], busy, tie),
                "created": _day((c.get("MetaData") or {}).get("CreateTime")),
                "invoices": sorted(s["Invoice"], key=lambda d: d["date"], reverse=True),
                "invoice_total": round(sum(d["amount"] for d in s["Invoice"]), 2),
                "payments": len(s["Payment"]), "estimates": len(s["Estimate"]),
                "bills": bills, "bill_total": round(sum(d["amount"] for d in bills), 2),
            })
        others = [r for r in recs if not r["used"]] or recs
        status = min((r["status"] for r in recs if r["status"] != "Real"), key=STATUS_ORDER.index, default="Real")
        out.append({"project": p, "status": status, "customers": recs,
                    "created": max(r["created"] for r in recs),
                    "money_on_duplicates": round(sum(r["bill_total"] + r["invoice_total"] for r in others), 2)})
    out.sort(key=lambda g: (STATUS_ORDER.index(g["status"]), g["project"]))
    return out


# ── the gate ─────────────────────────────────────────────────────────────────
def assess(project_nums: Iterable[str], con=None) -> Trust:
    """The holds for these jobs. Fails CLOSED: a mirror that cannot be read holds every job."""
    pns = {str(p).strip().upper() for p in project_nums if p and str(p).strip()}
    t = Trust()
    own = con is None
    try:
        con = con or mirror.connect()
        st = mirror.stamp(con)
        if st is None:
            t.global_reason = "QuickBooks copy was never loaded on this Mac - run sync-all first"
            return t
        t.mirror_as_of = st.isoformat()
        hours = (mirror.now_utc() - st).total_seconds() / 3600
        if hours > MAX_AGE_HOURS:
            t.global_reason = (f"QuickBooks copy is {hours:.0f} h old (limit {MAX_AGE_HOURS} h) - "
                               f"refresh it (sync-all) before the WIP overwrites costs / billed")
            return t
        _check_duplicates(t, pns, con)
        _check_lines(t, pns, con)
    except Exception as e:                  # noqa: BLE001 - fail closed, say why
        t.global_reason = f"QuickBooks trust check could not run ({type(e).__name__}: {e}) - nothing overwritten"
    finally:
        if own and con is not None:
            try:
                con.close()
            except Exception:               # noqa: BLE001
                pass
    return t


def _check_duplicates(t: Trust, pns: set, con) -> None:
    for g in duplicate_groups(con):
        p = g["project"]
        for r in g["customers"]:
            if r["used"]:
                continue
            if r["status"] == "Name typo":
                # its own job, hidden behind a space: 'RP7340 -FTW' is RP7340-FTW, not RP7340
                real = _squashed_proj(r["display_name"])
                if real in pns:
                    t.add(real, ("costs", "billed"),
                          f"QuickBooks customer '{r['name']}' has a space in its name, so the tools cannot "
                          f"find it as {real} - rename it to {re.sub(chr(32), '', r['display_name'])}")
                continue
            if p not in pns:
                continue
            money = r["invoice_total"] + r["bill_total"]
            if r["status"] == "Both have invoices":
                inv = ", ".join(d["doc"] or d["id"] for d in r["invoices"])
                t.add(p, ("costs", "billed"),
                      f"{p} is on two QuickBooks customers and both have invoices - '{r['name']}' "
                      f"(invoice {inv}) is left out; pick the right builder and move it")
            elif money >= MIN_AMOUNT or r["payments"]:
                what = []
                if r["invoices"]:
                    what.append(f"invoices {', '.join(d['doc'] or d['id'] for d in r['invoices'][:5])} ({_money(r['invoice_total'])})")
                if r["bills"]:
                    what.append(f"bills {', '.join(d['doc'] or d['id'] for d in r['bills'][:5])} ({_money(r['bill_total'])})")
                if r["payments"] and not what:
                    what.append(f"{r['payments']} payment(s)")
                t.add(p, ("costs", "billed"),
                      f"{p} is on two QuickBooks customers - '{r['name']}' is left out and carries "
                      + "; ".join(what) + " - move them onto the used customer and make it inactive")


def _check_lines(t: Trust, pns: set, con) -> None:
    """NOT CODED + FW ON SLAB, in one pass over the cost lines."""
    try:
        from shared import job_rulings
    except Exception:                       # noqa: BLE001
        job_rulings = None
    cust_proj: Dict[str, Optional[str]] = {}
    all_projs = set()
    for c in mirror.load("Customer", con=con):          # every customer, inactive too, for the id -> job map
        p = extract_proj(c.get("DisplayName") or c.get("CompanyName") or "")
        cust_proj[c["Id"]] = p
        if p and c.get("Active", True) is not False:
            all_projs.add(p)
    uncoded: dict = defaultdict(lambda: [0.0, set()])   # pn -> [amount, {doc #}]
    fw: dict = defaultdict(lambda: [0.0, set()])        # slab pn -> [amount, {doc #}]
    for ent in COST_ENTITIES:
        sign = -1 if ent == "VendorCredit" else 1
        for r in mirror.load(ent, con=con):
            memo = f"{r.get('PrivateNote') or ''} {r.get('Memo') or ''}"
            doc = r.get("DocNumber") or r["Id"]
            for ln in r.get("Line") or []:
                amt = sign * float(ln.get("Amount") or 0)
                if not amt:
                    continue
                for det in ln.values():
                    if not isinstance(det, dict) or not ("AccountRef" in det or "ItemRef" in det):
                        continue
                    cref = det.get("CustomerRef") or det.get("Entity", {}).get("EntityRef") if ent == "JournalEntry" else det.get("CustomerRef")
                    cid = (cref or {}).get("value") if isinstance(cref, dict) else None
                    if ent == "JournalEntry" and cid and (det.get("Entity") or {}).get("Type") not in (None, "Customer"):
                        cid = None
                    if cid:
                        slab = cust_proj.get(cid)
                        code = ((det.get("ItemRef") or {}).get("name") or "").split(":")[-1].strip()
                        if (slab and _RP_BASE_RE.match(slab) and _FW_RE.match(code)
                                and f"{slab}-FTW" in all_projs and (slab in pns or f"{slab}-FTW" in pns)):
                            fw[slab][0] += amt
                            fw[slab][1].add(doc)
                        continue
                    if ent == "JournalEntry":
                        continue                               # a JE line with no customer is book-keeping, not a job cost
                    cls = (det.get("ClassRef") or {}).get("name") or ""
                    named = {m.upper() for m in PROJ_RE.findall(f"{ln.get('Description') or ''} {cls} {memo}")}
                    if len(named) != 1:
                        continue                               # no job named, or 2+ (never guess a split)
                    pn = next(iter(named))
                    if pn not in pns:
                        continue
                    if job_rulings is not None and pn in cls.upper():
                        rule = job_rulings.cost_rule(pn)
                        if rule and str(rule.get("rule", "")).lower() == "class":
                            continue                           # the class-only cost ruling already counts it (MFD295)
                    uncoded[pn][0] += amt
                    uncoded[pn][1].add(doc)
    for pn, (amt, docs) in uncoded.items():
        if abs(amt) >= MIN_AMOUNT:
            t.add(pn, ("costs",), hold=abs(amt) >= HOLD_AMOUNT, reason=f"{len(docs)} QuickBooks cost document(s) name {pn} but have no project on the line "
                                  f"({_money(amt)}: {', '.join(sorted(map(str, docs))[:6])}{' …' if len(docs) > 6 else ''}) "
                                  f"- add the project so the cost counts")
    for slab, (amt, docs) in fw.items():
        if abs(amt) < MIN_AMOUNT:
            continue
        refs = f"{', '.join(sorted(map(str, docs))[:6])}{' …' if len(docs) > 6 else ''}"
        for pn in (slab, f"{slab}-FTW"):
            if pn in pns:
                t.add(pn, ("costs",), hold=abs(amt) >= HOLD_AMOUNT, reason=f"flatwork (FW) cost coded to the slab {slab} while {slab}-FTW exists "
                                      f"({_money(amt)}: {refs}) - move those lines to {slab}-FTW")


def main(argv: List[str]) -> int:
    pns = [a for a in argv if not a.startswith("-")]
    if not pns:
        print("usage: qbo_trust.py <PROJECT#> [...]")
        return 2
    t = assess(pns)
    if t.global_reason:
        print(f"EVERY JOB HELD: {t.global_reason}")
        return 1
    for pn in pns:
        pn = pn.upper()
        if not t.held(pn) and not (t.notes(pn, "costs") or t.notes(pn, "billed")):
            print(f"{pn}: trusted")
            continue
        for key in ("costs", "billed"):
            for why in t.reasons(pn, key):
                print(f"{pn} {key} HELD: {why}")
            for why in t.notes(pn, key):
                print(f"{pn} {key} check: {why}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
