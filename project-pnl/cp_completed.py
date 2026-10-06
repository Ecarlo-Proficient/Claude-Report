#!/usr/bin/env python3
"""
cp_completed.py - the CP COMPLETED workbook: every finished CP job of the year,
its result and whether it is really closed.

The owner 10/02/2026: "show me active excel and work on completed excel
separately so it's two excels" - and "for completed projects you also need to
be wary of projects not being done". So this book carries no projection (a
finished job's P&L falls back to an old WIP report for contract / ETC -
unconfirmed) and asks QuickBooks, job by job, what is still open:

  PAID IN FULL            nothing open in QBO
  RETAINAGE STILL OWED    the only open invoices are retainage releases
  CHECK - MAY NOT BE DONE an open regular invoice (a short-paid draw), or costs
                          landing more than 30 days after the last invoice

Same sheets as the CP Overview (`cp_overview.py`): one row per job, a sheet per
job with its draw coverage and every transaction. Billed and costs come from the
job's P&L workbook, built with `project_pnl_export.py --legacy` (these jobs
predate consistent project coding: cost lines named in the memo and invoices on
the GC count); the close-out facts (open
balances by invoice, last invoice date) are read from QBO directly. The job's
P&L marks a retainage release paid even while QBO still has it open, so QBO is
the only source for what is owed.

Read-only on QBO. Writes `Completed Projects/CP Completed <year>.xlsx`.

USAGE
  python3 project-pnl/cp_completed.py                 # this year
  python3 project-pnl/cp_completed.py --year 2026 --out /tmp/x.xlsx
"""
from __future__ import annotations

import argparse
import datetime as dt
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.path.insert(0, str(Path(__file__).resolve().parent))
import completed_pnl as cp  # noqa: E402
import cp_overview  # noqa: E402
from shared import pnl_paths, qbo_api  # noqa: E402
from shared.job_lines import JobMatcher  # noqa: E402

LATE_COST_DAYS = 30         # costs this long after the last invoice = maybe not done
_RET = re.compile(r"retainage", re.I)


def _d(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    try:
        return dt.date.fromisoformat(str(v)[:10])
    except ValueError:
        return None


def closeout(job: str, src: dict, invoices: list) -> dict:
    """What QBO still has open on a finished job, and which bucket it lands in."""
    open_ret, open_reg = [], []
    for inv in invoices:
        bal = float(inv.get("Balance") or 0)
        if bal <= 0.005:
            continue
        items = [((ln.get("SalesItemLineDetail") or {}).get("ItemRef") or {}).get("name", "")
                 for ln in inv.get("Line", []) if ln.get("DetailType") == "SalesItemLineDetail"]
        tag = (inv.get("DocNumber"), _d(inv.get("TxnDate")), bal)
        (open_ret if items and all(_RET.search(x) for x in items) else open_reg).append(tag)
    last_inv = max((_d(i.get("TxnDate")) for i in invoices if _d(i.get("TxnDate"))), default=None)
    costs = [_d(ln["date"]) for s in src["sections"] for a in s["accounts"] for v in a["vendors"]
             for ln in v["lines"] if _d(ln["date"])]
    last_cost = max(costs, default=None)
    late = bool(last_inv and last_cost and (last_cost - last_inv).days > LATE_COST_DAYS)
    owed = round(sum(x[2] for x in open_ret + open_reg), 2)

    notes = []
    for doc, date, bal in open_reg:
        notes.append(f"draw #{doc} {date:%m/%d/%Y} short-paid, ${bal:,.2f} open")
    for doc, date, bal in open_ret:
        notes.append(f"retainage #{doc} {date:%m/%d/%Y} unpaid, ${bal:,.2f}")
    if late:
        notes.append(f"costs to {last_cost:%m/%d/%Y}, {(last_cost - last_inv).days} days after "
                     f"the last invoice")
    if open_reg or late:
        bucket = "check"
    elif open_ret:
        bucket = "retainage"
    else:
        bucket = "closed"
        notes.append(f"paid in full · last invoice {last_inv:%m/%d/%Y}" if last_inv else "paid in full")
    return {"bucket": bucket, "owed": owed, "text": " · ".join(notes)}


def main() -> int:
    ap = argparse.ArgumentParser(description="CP Completed <year>.xlsx - finished CP jobs, QBO close-out")
    ap.add_argument("--year", type=int, default=dt.date.today().year)
    ap.add_argument("--out", default=None, help="write here instead of Completed Projects/")
    a = ap.parse_args()

    base = pnl_paths.CP_AWARDED_BASE
    div = cp.DIVISIONS["cp"]
    loaded, skipped = cp.load_division(cp._iter_jobs(base, "CP"), base, a.year)
    done = [x for x in loaded if x[1].get("status") != "Active"]
    if skipped:
        print(f"  · nothing in {a.year}: {', '.join(sorted(skipped))}")
    if not done:
        print("✗  no finished CP jobs for", a.year)
        return 1

    access, cid = qbo_api.load_credentials()
    pmap = qbo_api.build_project_customer_map(access, cid)
    jobs = []
    for job, src, t, path in done:
        info = pmap.get(job)
        invs = (qbo_api.query_all(access, cid, "Invoice", f"CustomerRef = '{info['id']}'")
                if info else [])
        if info and info.get("parent_id"):
            # finished jobs predate project coding: some were invoiced on the GC
            # (CP714 Draw #1) - the memo decides, as in the --legacy P&L
            m = JobMatcher(info["id"], job, [], legacy=True)
            have = {i["Id"] for i in invs}
            invs += [i for i in qbo_api.query_all(access, cid, "Invoice",
                                                   f"CustomerRef = '{info['parent_id']}'")
                     if i["Id"] not in have and m.invoice_belongs(i)]
        co = closeout(job, src, invs) if info else {"bucket": "check", "owed": None,
                                                      "text": "not a QBO project - check by hand"}
        ex = cp_overview.read_extra(path)
        ex.setdefault("static", {}).update({"ar": co["owed"] or None, "closeout": co["text"]})
        ex["bucket"] = co["bucket"]
        jobs.append((job, src, t, ex))
        print(f"  {job:<7} {co['bucket']:<10} {co['text']}")

    sections = [
        ("PAID IN FULL", "", lambda j: j[3]["bucket"] == "closed"),
        ("RETAINAGE STILL OWED", "work finished, retainage not collected",
         lambda j: j[3]["bucket"] == "retainage"),
        ("CHECK - MAY NOT BE DONE", "open billing or late costs - confirm before closing",
         lambda j: j[3]["bucket"] == "check"),
    ]
    out = (Path(a.out).expanduser() if a.out
           else base / pnl_paths.CP_COMPLETED_SUBDIR / f"CP Completed {a.year}.xlsx")
    cp_overview.build(jobs, out, dict(div, scope=f"{a.year}"), sections=sections,
                      title=f"COMMERCIAL COMPLETED {a.year}  ·  RESULTS AND CLOSE-OUT",
                      cols=cp_overview.COMPLETED_COLS)
    print(f"  → {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
