#!/usr/bin/env python3
"""
legacy_job_cost_pull.py — costs + billing for an OLDER job whose lines were
never consistently project-coded.

WHY THIS EXISTS: jobs that predate consistent project coding carry only part
of their cost on the project customer. The rest was coded to the parent or to
nothing at all, with the job named only in the line description or in the
bill's memo. A plain "costs for this customer" pull silently under-reports
those jobs — on the first job this was built for, by ~2.4% of total cost.

ATTRIBUTION — Bill + Purchase LINE ITEMS, never txn totals (most bills are
multi-line and only some lines belong to the job). A line is taken if any of
these is true, in this order:

  1. project     the line's own CustomerRef is the project customer
  2. line text   the line Description or line CustomerRef.name names the job
  3. bill note   the BILL's PrivateNote names the job AND names exactly ONE
                 job number AND the line's own text names no job at all

  GUARD on rule 3: a memo listing more than one job number (shared pump and
  material vendors do this constantly) is SKIPPED, never split. Missing a
  shared bill is cheaper than attributing another job's money to this one.

BILLING is separate and simpler: invoices ARE customer coded, but on an older
job they often sit on the PARENT customer rather than the project. Both
customers are pulled and the invoices whose PrivateNote names the job are
kept.

Read-only against QBO. One credential unlock per run. The company-wide Bill
and Purchase pulls are cached on disk under ~/Library/Logs/Proficient/ and
shared across jobs, so the second job you run costs nothing; --refresh
re-pulls.

USAGE
  python3 one-offs/legacy_job_cost_pull.py --project MFD172 --alias "BONDS RANCH"
  python3 one-offs/legacy_job_cost_pull.py --project MFD172 --alias "BONDS RANCH" \
      --as-of 2026-03-31 --interim 2025-12-31 --etc 4232000
  python3 one-offs/legacy_job_cost_pull.py --project MFD172 --csv
  python3 one-offs/legacy_job_cost_pull.py --project MFD172 --expect <fixture.json>

`--expect` takes a JSON file of known-good figures and prints ✓/✗ per line —
that is how a pull is proven to reproduce. Keep fixtures OUTSIDE this repo
(the log dir is the natural home); job dollars are not repo content.
"""
from __future__ import annotations

import argparse
import csv
import json
import re
import sys
import time
from pathlib import Path
from typing import Dict, List, Optional, Tuple

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from shared import qbo_api
from shared.job_lines import JobMatcher

DEFAULT_SINCE = "2023-01-01"
CACHE_DIR = Path.home() / "Library" / "Logs" / "Proficient" / "legacy-job-pull"
CACHE_TTL = 12 * 3600


# ── the job ───────────────────────────────────────────────────────────

def resolve_job(access: str, cid: str, project: str,
                project_id: Optional[str], parent_id: Optional[str]
                ) -> Tuple[str, Optional[str], str]:
    """(project customer id, parent customer id, display name). Explicit ids
    win; otherwise the project # is matched against the customer list."""
    if project_id:
        return project_id, parent_id, project
    for c in qbo_api.query_all(access, cid, "Customer"):
        name = c.get("DisplayName") or c.get("CompanyName") or ""
        if qbo_api.extract_proj(name) == project.upper():
            return (c["Id"], parent_id or (c.get("ParentRef") or {}).get("value"),
                    name)
    print(f"✗  No customer found for {project}. Pass --project-id explicitly.")
    raise SystemExit(1)


# ── raw pulls (disk-cached) ───────────────────────────────────────────

def _cached(name: str, refresh: bool, fetch) -> List[dict]:
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    f = CACHE_DIR / f"{name}.json"
    if not refresh and f.exists() and (time.time() - f.stat().st_mtime) < CACHE_TTL:
        try:
            return json.loads(f.read_text())
        except Exception:
            pass
    rows = fetch()
    f.write_text(json.dumps(rows))
    return rows


def pull(access: str, cid: str, since: str, project_id: str,
         parent_id: Optional[str], refresh: bool
         ) -> Tuple[List[dict], List[dict], List[dict]]:
    where = f"TxnDate >= '{since}'"
    # Company-wide and job-independent: cached once, reused by every job.
    bills = _cached(f"bills_{since}", refresh,
                    lambda: qbo_api.query_all(access, cid, "Bill", where))
    purchases = _cached(f"purchases_{since}", refresh,
                        lambda: qbo_api.query_all(access, cid, "Purchase", where))

    def _inv():
        out: List[dict] = []
        # The query parser is AND-only — one call per customer, merged here.
        for cust in [c for c in (project_id, parent_id) if c]:
            out += qbo_api.query_all(access, cid, "Invoice",
                                     f"CustomerRef = '{cust}'")
        return out

    invoices = _cached(f"invoices_{project_id}", refresh, _inv)
    return bills, purchases, invoices


# ── attribution ───────────────────────────────────────────────────────
# The three rules and their guard live in shared/job_lines.py — project-pnl
# runs on the SAME matcher (`--legacy`), so the P&L and this pull can never
# disagree about what belongs to a job.

def attribute(txns: List[dict], tx_type: str, vendor_field: str,
              matcher: JobMatcher) -> List[dict]:
    """One row per attributed LINE. `rule` records which of the three fired."""
    rows: List[dict] = []
    for t in txns or []:
        memo = (t.get("PrivateNote") or "").strip()
        vendor = ((t.get(vendor_field) or {}).get("name") or "").strip()
        for idx, ln in enumerate(t.get("Line") or []):
            det = (ln.get("AccountBasedExpenseLineDetail")
                   or ln.get("ItemBasedExpenseLineDetail"))
            if not det:
                continue
            rule = matcher.rule(det, ln, t)
            if not rule:
                continue
            cref = det.get("CustomerRef") or {}
            desc = (ln.get("Description") or "").strip()

            rows.append({
                "rule": rule,
                "txn_type": tx_type,
                "txn_id": t.get("Id", ""),
                "line_id": str(ln.get("Id") or idx),
                "date": t.get("TxnDate") or "",
                "vendor": vendor,
                "amount": float(ln.get("Amount", 0) or 0),
                "customer_id": cref.get("value") or "",
                "customer": cref.get("name") or "",
                "description": desc,
                "memo": memo,
            })
    return rows


def invoice_rows(invoices: List[dict], matcher: JobMatcher) -> List[dict]:
    """Job invoices: customer-coded to project OR parent, memo names the job.
    Voided invoices are dropped (QBO zeroes them and prefixes 'Voided - ')."""
    out, seen = [], set()
    for inv in invoices or []:
        memo = (inv.get("PrivateNote") or "").strip()
        if not matcher.invoice_belongs(inv):
            continue
        if inv["Id"] in seen:            # both customer pulls can return it
            continue
        seen.add(inv["Id"])
        out.append({
            "txn_id": inv["Id"],
            "doc": inv.get("DocNumber") or "",
            "date": inv.get("TxnDate") or "",
            "customer_id": (inv.get("CustomerRef") or {}).get("value") or "",
            "amount": float(inv.get("TotalAmt", 0) or 0),
            "memo": memo,
        })
    return sorted(out, key=lambda r: (r["date"], r["doc"]))


def classify_invoice(memo: str) -> str:
    """Structure, not keyword guessing: the GC's memo says what it is.
    'retain' anywhere wins (a memo can read 'October Draw 2025 - Retainage'
    and still be a release, not a draw); 'draw' is a monthly draw; anything
    else was billed outside the draw schedule — typically a CO whose memo
    never uses the words 'change order'."""
    m = memo.lower()
    if "retain" in m:
        return "retainage release"
    if "draw" in m:
        return "monthly draw"
    return "change order / extra"


# ── reporting ─────────────────────────────────────────────────────────

def _money(x: float) -> str:
    return f"{x:>16,.2f}"


def _mark(got: float, exp, n_got: Optional[int] = None, n_exp=None) -> str:
    if exp is None:
        return " "
    ok = abs(got - float(exp)) < 0.01 or (
        float(exp) == round(float(exp)) and abs(got - float(exp)) < 1.0)
    if n_exp is not None and n_got is not None:
        ok = ok and n_got == int(n_exp)
    return "✓" if ok else "✗"


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Costs + billing for an older, inconsistently coded job")
    ap.add_argument("--project", required=True, help="e.g. MFD172")
    ap.add_argument("--alias", action="append", default=[],
                    help="street name, repeatable (e.g. 'BONDS RANCH')")
    ap.add_argument("--project-id", help="override the project customer id")
    ap.add_argument("--parent-id", help="override the parent customer id")
    ap.add_argument("--since", default=DEFAULT_SINCE)
    ap.add_argument("--as-of", required=True, help="inclusive cutoff")
    ap.add_argument("--interim", help="second cutoff to report alongside")
    ap.add_argument("--etc", type=float,
                    help="budget/ETC, to print the variance line")
    ap.add_argument("--expect", help="JSON of known-good figures to verify")
    ap.add_argument("--refresh", action="store_true", help="re-pull from QBO")
    ap.add_argument("--csv", action="store_true", help="write line detail")
    a = ap.parse_args()

    exp = json.loads(Path(a.expect).read_text()) if a.expect else {}
    access, cid = qbo_api.load_credentials()          # realm never printed
    pid, parent, disp = resolve_job(access, cid, a.project,
                                    a.project_id, a.parent_id)
    matcher = JobMatcher(pid, a.project, a.alias, legacy=True)

    print(f"{disp}" + (f"  ({', '.join(a.alias)})" if a.alias else ""))
    print(f"cost window {a.since} .. {a.as_of}   line-item attribution\n")

    bills, purchases, invoices = pull(access, cid, a.since, pid, parent,
                                      a.refresh)
    print(f"pulled  {len(bills):,} bills   {len(purchases):,} purchases   "
          f"{len(invoices):,} invoices on the project/parent customers\n")

    lines = (attribute(bills, "Bill", "VendorRef", matcher)
             + attribute(purchases, "Purchase", "EntityRef", matcher))

    def tot(rows, cutoff=None):
        rows = [r for r in rows if cutoff is None or r["date"] <= cutoff]
        return round(sum(r["amount"] for r in rows), 2), len(rows)

    cut = [r for r in lines if r["date"] <= a.as_of]
    print("COST  (Bill + Purchase line items)")
    for rule, key in (("project", "cost_project"),
                      ("line text", "cost_line_text"),
                      ("bill note", "cost_bill_note")):
        amt, n = tot([r for r in cut if r["rule"] == rule])
        e = exp.get(key) or [None, None]
        print(f"  {'via ' + rule:<16}{_money(amt)}  {n:>6,} lines   "
              f"{_mark(amt, e[0], n, e[1])}")
    t_amt, t_n = tot(cut)
    e = exp.get("cost_total") or [None, None]
    print(f"  {'TOTAL':<16}{_money(t_amt)}  {t_n:>6,} lines   "
          f"{_mark(t_amt, e[0], t_n, e[1])}")
    if a.interim:
        i_amt, i_n = tot(lines, a.interim)
        e = exp.get("cost_total_interim") or [None, None]
        print(f"  {'thru ' + a.interim:<16}{_money(i_amt)}  {i_n:>6,} lines   "
              f"{_mark(i_amt, e[0])}")
    print()

    inv = invoice_rows(invoices, matcher)
    b_asof = round(sum(r["amount"] for r in inv if r["date"] <= a.as_of), 2)
    kinds: Dict[str, int] = {}
    for r in inv:
        if r["date"] <= a.as_of:
            k = classify_invoice(r["memo"])
            kinds[k] = kinds.get(k, 0) + 1

    print("BILLED  (invoices on the project OR parent, memo names the job)")
    if a.interim:
        b_int = round(sum(r["amount"] for r in inv
                          if r["date"] <= a.interim), 2)
        e = exp.get("billed_interim") or [None]
        print(f"  {'thru ' + a.interim:<16}{_money(b_int)}  {_mark(b_int, e[0])}")
    e = exp.get("billed_asof") or [None]
    print(f"  {'thru ' + a.as_of:<16}{_money(b_asof)}  {_mark(b_asof, e[0])}")
    print("  " + "   ".join(f"{n} {k}{'s' if n != 1 else ''}"
                            for k, n in sorted(kinds.items())) + "\n")

    gp = b_asof - t_amt
    print("RESULT")
    print(f"  {'contract':<16}{_money(b_asof)}")
    print(f"  {'actual cost':<16}{_money(t_amt)}")
    print(f"  {'gross profit':<16}{_money(gp)}   "
          f"{(gp / b_asof * 100 if b_asof else 0):.2f}%")
    if a.etc:
        print(f"  {'ETC / budget':<16}{_money(a.etc)}  (budget, not actual)")
        print(f"  {'vs budget':<16}{_money(t_amt - a.etc)}")

    if a.csv:
        CACHE_DIR.mkdir(parents=True, exist_ok=True)
        for rows, stem in ((cut, "cost_lines"), (inv, "invoices")):
            if not rows:
                continue
            out = CACHE_DIR / f"{a.project}_{stem}_{a.as_of}.csv"
            with out.open("w", newline="") as fh:
                w = csv.DictWriter(fh, fieldnames=list(rows[0].keys()))
                w.writeheader()
                w.writerows(rows)
            print(f"detail → {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
