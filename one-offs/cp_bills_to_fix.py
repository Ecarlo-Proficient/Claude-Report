#!/usr/bin/env python3
"""
cp_bills_to_fix.py - the QBO bills to fix so every CP job carries its own cost
(read-only on QBO; writes one Excel list).

The owner 10/02/2026: "i need the bill list like a table to fix". A cost line
belongs on a job's P&L when it is CODED to the job's QBO project. This lists,
one row per bill and job, grouped by vendor, every line that is NOT, but should be looked at:

  ADD PROJECT     no project on the line, and its memo / description names
                  this job and only this job
  ON JOB CLASS    no project, but the line sits on the job's own class
  CHECK CODING    coded to ANOTHER job while the memo names only this one -
                  or on this job's class while the memo names another job
Left out on purpose: bills whose memo names several jobs (split line by line
already), payroll, and estimating / takeoff fees (overhead, not job cost).

On an ACTIVE job these lines are missing from the P&L today. On a FINISHED job
the P&L is built --legacy and already counts the memo-named ones; coding the
project in QBO makes that permanent.

Only bills AFTER the QBO closing date - a closed-period bill is locked. The
Bill Tracker's Audit - Coding sheet is the clerk's standing list (it now carries
Wrong Job? too); this is the CP-only cut of it.

USAGE
  python3 one-offs/cp_bills_to_fix.py                 # -> CompanyHealth/Analysis/CP bills to fix (date)/
"""
from __future__ import annotations

import datetime as dt
import re
import sys
from collections import defaultdict
from pathlib import Path
from urllib.parse import quote

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "project-pnl"))
from openpyxl import Workbook  # noqa: E402
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side  # noqa: E402

import completed_pnl as cp  # noqa: E402
from shared import paths, pnl_paths, qbo_api, qbo_costs as qc  # noqa: E402
from shared.job_lines import discover_job_classes  # noqa: E402
from shared.xlsx_guard import put  # noqa: E402
from shared.xlsx_verify import assert_clean  # noqa: E402

# job numbers as the clerk writes them - 'CP745', 'CP790-9001 LANDMARK' (the
# same pattern the bill audit's Wrong Job? check uses; tools never import tools)
_JOB = re.compile(r"\b(MFD|CP|RP)\s?-?(\d{3,5})(-FTW)?(?=\b|-)", re.IGNORECASE)


def _named(text: str) -> set:
    return {f"{m.group(1).upper()}{m.group(2)}{(m.group(3) or '').upper()}" for m in _JOB.finditer(text or "")}


PAY = re.compile(r"payroll|wages|salar|employee benefit|workers.?comp", re.I)
OVERHEAD = re.compile(r"admin contract labor|takeoff|estimat", re.I)


def _url(kind: str, txn_id: str, realm: str) -> str:
    page = "bill" if kind == "Bill" else "expense"
    return (f"https://qbo.intuit.com/app/login?pagereq={quote(f'{page}?txnId={txn_id}')}"
            f"&deeplinkcompanyid={realm}")


def main() -> int:
    base = pnl_paths.CP_AWARDED_BASE
    loaded, _ = cp.load_division(cp._iter_jobs(base, "CP"), base, cp.resolve_year(cp.DIVISIONS["cp"]))
    status = {j: s.get("status") for j, s, _t, _p in loaded}
    jobs = set(status)

    access, cid = qbo_api.load_credentials()
    pmap = qbo_api.build_project_customer_map(access, cid)
    acct = qc.build_account_map(access, cid)
    c2p = {v["id"]: p for p, v in pmap.items()}
    classes = (qbo_api.query_all(access, cid, "Class")
               + qbo_api.query_all(access, cid, "Class", "Active = false"))
    cls_of = {}
    for j in jobs:
        for k in discover_job_classes(classes, j):
            cls_of[str(k)] = j
    # OPEN PERIOD ONLY (the owner 10/02: "those bills are in 2025 ... locked by admin
    # closing date"): nothing on or before the QBO closing date - nobody can fix it
    pref = qbo_api.query_all(access, cid, "Preferences")
    closed = ((pref[0].get("AccountingInfoPrefs") or {}).get("BookCloseDate") if pref else None)
    since = (dt.date.fromisoformat(closed) + dt.timedelta(days=1)).isoformat() if closed else "2026-01-01"
    bills, purchases = qc.pull_expense_txns(access, cid, since)

    rows = defaultdict(lambda: {"amt": 0.0, "n": 0})
    for kind, txns in (("Bill", bills), ("Expense", purchases)):
        for t in txns:
            memo = str(t.get("PrivateNote") or "")
            vendor = ((t.get("VendorRef") or t.get("EntityRef")) or {}).get("name", "")
            for ln in t.get("Line") or []:
                det = ln.get("AccountBasedExpenseLineDetail") or ln.get("ItemBasedExpenseLineDetail")
                if not det:
                    continue
                amt = float(ln.get("Amount") or 0)
                leaf = qc.cost_leaf(det, acct, fallback="")
                if not amt or PAY.search(leaf):
                    continue
                coded = c2p.get((det.get("CustomerRef") or {}).get("value"))
                named = _named(memo + " " + str(ln.get("Description") or ""))
                cls = str(((det.get("ClassRef") or ln.get("ClassRef") or t.get("ClassRef")) or {}).get("value") or "")
                on_class = cls_of.get(cls)
                job = action = None
                one = next(iter(named)) if len(named) == 1 else ""
                # a job and its -FTW twin are one family (same rule as the bill audit)
                if one and one in jobs and coded != one and (coded or "").split("-")[0] != one.split("-")[0]:
                    job = one
                    action = ("ADD PROJECT" if not coded else "CHECK CODING")
                elif on_class and not coded and not named:
                    job, action = on_class, "ON JOB CLASS"
                elif on_class and not coded and named and on_class not in named:
                    job, action = on_class, "CHECK CODING"
                if not job:
                    continue
                if OVERHEAD.search(leaf + " " + memo):
                    continue
                r = rows[(job, kind, t["Id"], action)]
                r.update(doc=t.get("DocNumber") or "", date=t.get("TxnDate"), vendor=vendor,
                         coded=coded or "", memo=memo.replace("\n", " ")[:120],
                         named=", ".join(sorted(named)) or "", leaf=leaf)
                r["amt"] += amt
                r["n"] += 1

    def fix(job, action, r):
        if action == "ADD PROJECT":
            return f"No project - set {job}"
        if action == "ON JOB CLASS":
            return f"On {job}'s class, no project - set {job}"
        if r["coded"]:
            return f"Coded to {r['coded']}, memo says {job} - confirm the job"
        return f"On {job}'s class, memo says {r['named']} - fix class or memo"

    # THE CLERK'S LIST (the owner 10/02: "too confusing for clerk ... just like the
    # table you have shown me, simple, no freeze panes, grouped by vendor"):
    # one plain table, a vendor row with its total, its bills under it.
    by_vendor = defaultdict(list)
    for (job, kind, tid, action), r in rows.items():
        by_vendor[r["vendor"] or "(no vendor)"].append((job, kind, tid, action, r))
    out = [x for v in by_vendor.values() for x in v]
    realm = cid
    wb = Workbook()
    ws = wb.active
    ws.title = "Bills to fix"
    thin = Side(style="thin", color="BFBFBF")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    ws["A1"] = "Bills to fix in QuickBooks"
    ws["A1"].font = Font(bold=True, size=14)
    heads = ["Vendor / Job", "Bill #", "Date", "Amount", "Fix"]
    for c, h in enumerate(heads, 1):
        cell = ws.cell(3, c, h)
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="D9D9D9")
        cell.border = box
    rr = 4
    for vendor in sorted(by_vendor, key=str.lower):
        items = sorted(by_vendor[vendor], key=lambda x: str(x[4]["date"]))
        put(ws, rr, 1, vendor)
        ws.cell(rr, 4, round(sum(x[4]["amt"] for x in items), 2))
        for c in range(1, 6):
            cell = ws.cell(rr, c)
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor="EEF1F5")
            cell.border = box
        ws.cell(rr, 4).number_format = '"$"#,##0.00;-"$"#,##0.00'
        rr += 1
        for job, kind, tid, action, r in items:
            put(ws, rr, 1, job)
            ws.cell(rr, 1).alignment = Alignment(indent=1)
            doc = put(ws, rr, 2, str(r["doc"]) or "(no #)")
            doc.hyperlink = _url(kind, tid, realm)          # click -> the bill in QBO
            doc.font = Font(color="0563C1", underline="single")
            d = r["date"]
            try:                                        # QBO gives 'YYYY-MM-DD' text; show mm/dd/yyyy
                d = dt.date.fromisoformat(str(d)[:10])
            except ValueError:
                pass
            ws.cell(rr, 3, d).number_format = "mm/dd/yyyy"
            amt = ws.cell(rr, 4, round(r["amt"], 2))
            amt.number_format = '"$"#,##0.00;-"$"#,##0.00'
            put(ws, rr, 5, fix(job, action, r))
            for c in range(1, 6):
                ws.cell(rr, c).border = box
            rr += 1
    for col, w in zip("ABCDE", (34, 16, 12, 14, 52)):
        ws.column_dimensions[col].width = w
    out_dir = paths.analysis_dir("CP bills to fix")
    path = out_dir / "CP bills to fix.xlsx"
    wb.save(path)
    assert_clean(path)
    print(f"{len(out)} bills · {len(by_vendor)} vendors · → {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
