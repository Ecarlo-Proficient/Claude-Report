#!/usr/bin/env python3
"""
cp_bills_to_fix.py - the QBO bills to fix so every CP job carries its own cost
(read-only on QBO; writes one Excel list).

The owner 10/02/2026: "i need the bill list like a table to fix". A cost line
belongs on a job's P&L when it is CODED to the job's QBO project. This lists,
one row per bill and job, every line that is NOT, but should be looked at:

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

USAGE
  python3 one-offs/cp_bills_to_fix.py                 # -> CompanyHealth/Analysis/CP bills to fix (date)/
"""
from __future__ import annotations

import re
import sys
from collections import defaultdict
from pathlib import Path
from urllib.parse import quote

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "project-pnl"))
from openpyxl import Workbook  # noqa: E402
from openpyxl.formatting.rule import FormulaRule  # noqa: E402
from openpyxl.styles import Alignment, Font, PatternFill  # noqa: E402
from openpyxl.worksheet.table import Table, TableStyleInfo  # noqa: E402

import completed_pnl as cp  # noqa: E402
from shared import paths, pnl_paths, qbo_api, qbo_costs as qc  # noqa: E402
from shared.job_lines import discover_job_classes, jobs_named_in  # noqa: E402
from shared.xlsx_guard import put  # noqa: E402
from shared.xlsx_verify import assert_clean  # noqa: E402

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
    bills, purchases = qc.pull_expense_txns(access, cid, "2020-01-01")

    rows = defaultdict(lambda: {"amt": 0.0, "n": 0})
    skipped_oh = 0.0
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
                named = jobs_named_in(memo + " " + str(ln.get("Description") or ""))
                cls = str(((det.get("ClassRef") or ln.get("ClassRef") or t.get("ClassRef")) or {}).get("value") or "")
                on_class = cls_of.get(cls)
                job = action = None
                if len(named) == 1 and next(iter(named)) in jobs and coded != next(iter(named)):
                    job = next(iter(named))
                    action = ("ADD PROJECT" if not coded else "CHECK CODING")
                elif on_class and not coded and not named:
                    job, action = on_class, "ON JOB CLASS"
                elif on_class and not coded and named and on_class not in named:
                    job, action = on_class, "CHECK CODING"
                if not job:
                    continue
                if OVERHEAD.search(leaf + " " + memo):
                    skipped_oh += amt
                    continue
                r = rows[(job, kind, t["Id"], action)]
                r.update(doc=t.get("DocNumber") or "", date=t.get("TxnDate"), vendor=vendor,
                         coded=coded or "", memo=memo.replace("\n", " ")[:120],
                         named=", ".join(sorted(named)) or "", leaf=leaf)
                r["amt"] += amt
                r["n"] += 1

    def note(job, action, r):
        if action == "ADD PROJECT":
            return f"Set the project to {job} on these lines"
        if action == "ON JOB CLASS":
            return f"On {job}'s class with no project - set the project to {job} if it is this job's"
        if r["coded"]:
            return f"Coded to {r['coded']} but the memo names only {job} - confirm which job it is"
        return f"On {job}'s class but the memo names {r['named']} - fix the class or the memo"

    out = sorted(rows.items(), key=lambda kv: (status[kv[0][0]] != "Active", kv[0][3] != "ADD PROJECT",
                                                -abs(kv[1]["amt"])))
    realm = cid
    wb = Workbook()
    ws = wb.active
    ws.title = "Bills to fix"
    ws.sheet_view.showGridLines = False
    hdr = ["Job", "Job status", "On the P&L today?", "Fix", "What to do", "Type", "Bill #", "Date",
           "Vendor", "Amount", "Lines", "Cost account / code", "Bill memo", "Open in QBO"]
    # compact summary on top
    ws["B1"] = "CP bills to fix in QuickBooks"
    ws["B1"].font = Font(bold=True, size=16, color="1F3A5F")
    summ = defaultdict(lambda: [0, 0.0])
    for (job, kind, tid, action), r in out:
        summ[(status[job], action)][0] += 1
        summ[(status[job], action)][1] += r["amt"]
    sr = 3
    for (st, action), (n, amt) in sorted(summ.items(), key=lambda x: (x[0][0] != "Active", x[0][1])):
        put(ws, sr, 2, f"{'Active' if st == 'Active' else 'Finished'} job · {action}")
        ws.cell(sr, 3, n).number_format = "0"
        c = ws.cell(sr, 4, round(amt, 2))
        c.number_format = '"$"#,##0.00;[Red]-"$"#,##0.00'
        for cc in (2, 3, 4):
            ws.cell(sr, cc).fill = PatternFill("solid", fgColor="FBEFD5" if st == "Active" else "EEF1F5")
        sr += 1
    put(ws, sr, 2, f"left out: estimating / takeoff fees on these jobs (overhead) ${skipped_oh:,.2f}")
    ws.cell(sr, 2).font = Font(italic=True, color="4B5563")
    top = sr + 2
    for i, h in enumerate(hdr):
        ws.cell(top, 2 + i, h).font = Font(bold=True)
    r0 = top + 1
    for k, ((job, kind, tid, action), r) in enumerate(out):
        rr = r0 + k
        active = status[job] == "Active"
        vals = [job, "Active" if active else "Finished",
                ("NO - missing" if active or action != "ADD PROJECT" else "Yes - by memo (legacy)"),
                action, note(job, action, r), kind, str(r["doc"]), r["date"], r["vendor"],
                round(r["amt"], 2), r["n"], r["leaf"], r["memo"]]
        for i, v in enumerate(vals):
            if isinstance(v, str):
                put(ws, rr, 2 + i, v)
            else:
                ws.cell(rr, 2 + i, v)
        ws.cell(rr, 9).number_format = "mm/dd/yyyy"
        ws.cell(rr, 11).number_format = '"$"#,##0.00;[Red]-"$"#,##0.00'
        link = ws.cell(rr, 15, "open")
        link.hyperlink = _url(kind, tid, realm)
        link.font = Font(color="0563C1", underline="single")
    last = r0 + len(out) - 1
    if out:
        from openpyxl.utils import get_column_letter as gl
        ref = f"B{top}:{gl(1 + len(hdr))}{last}"
        tab = Table(displayName="BillsToFix", ref=ref)
        tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
        ws.add_table(tab)
        ws.conditional_formatting.add(f"D{r0}:D{last}", FormulaRule(
            formula=[f'LEFT($D{r0},2)="NO"'], font=Font(bold=True, color="B00020")))
        ws.conditional_formatting.add(f"E{r0}:E{last}", FormulaRule(
            formula=[f'$E{r0}="CHECK CODING"'], fill=PatternFill("solid", fgColor="FBEFD5")))
    for col, w in zip("BCDEFGHIJKLMNO", (9, 11, 20, 15, 58, 9, 16, 11, 30, 13, 7, 26, 50, 9)):
        ws.column_dimensions[col].width = w
    for row in ws.iter_rows(min_row=r0, max_row=last, min_col=6, max_col=6):
        for c in row:
            c.alignment = Alignment(wrap_text=False)
    ws.freeze_panes = f"A{r0}"
    out_dir = paths.analysis_dir("CP bills to fix")
    path = out_dir / "CP bills to fix.xlsx"
    wb.save(path)
    assert_clean(path)
    print(f"{len(out)} bills · → {path}")
    for (st, action), (n, amt) in sorted(summ.items()):
        print(f"  {st:<10} {action:<14} {n:>4}  ${amt:,.2f}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
