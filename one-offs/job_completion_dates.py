#!/usr/bin/env python3
"""
job_completion_dates.py — when did each job actually FINISH?

The question this exists to answer (the user 2026-09-01): "how are you going to
know which CP projects were completed in 2026?"

For CP you cannot, from anything local:
  * the WIP master carries a STATUS (Active / Closed) but **no completion
    date** - it says a job is done, never when it finished;
  * the ledger holds only the two WIP snapshots taken this month, so there is
    no status history to find an Active -> Closed flip in;
  * the ledger's `billing_event` table is thin - most Closed CP jobs have 0-2
    rows, so "last invoice" cannot be read from it either.

The one honest source is QBO's invoice history. A job's LAST INVOICE dates its
completion: on CP that is the final draw or the retainage release. This pulls
every invoice once, groups by project #, and joins the WIP master's status, so
"completed in <year>" becomes a set you can name instead of a guess.

**RP works completely differently, and the data says so plainly.** Not one RP
row in the WIP master is ever marked Closed - because a finished RP job is
REMOVED from the master, not flagged. Measured 2026-09-01: of 745 RP jobs ever
invoiced, the 57 on the master are ALL still moving (none quiet 60 days), and
every one of the 688 off it has gone quiet. So for RP:

    on the WIP master  = running        off it = finished
    last invoice year  = the year it finished

The `--quiet-days` rule below is kept as the CROSS-CHECK that proved this, not
as the primary test - it agreed with WIP-master presence on all 745 jobs.

Read-only. Writes nothing but stdout, and never touches the ledger.

USAGE
  python3 one-offs/job_completion_dates.py                 # all divisions
  python3 one-offs/job_completion_dates.py --division CP --closed-in 2026
"""
from __future__ import annotations

import argparse
import datetime as dt
import os
import sqlite3
import sys
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from shared.qbo_api import PROJ_RE, load_credentials, query_all


LEDGER_DB = Path(os.environ.get(
    "ACB_LEDGER_DB",
    Path.home() / "Library" / "Application Support" / "Proficient" / "ledger.sqlite3"))


def _last_costs() -> dict:
    """{project -> last cost date} from the ledger. Read-only; empty if the
    ledger has never been built on this machine."""
    try:
        con = sqlite3.connect(f"file:{LEDGER_DB}?mode=ro", uri=True)
    except sqlite3.Error:
        return {}
    try:
        return {p: d for p, d in con.execute(
            "select project_no, max(txn_date) from cost_line group by 1")}
    except sqlite3.Error:
        return {}
    finally:
        con.close()


def _proj_of(inv: dict) -> str:
    """The project # on an invoice. CustomerRef.name is
    `Parent:Project # Name`, so search it - never match from the start."""
    name = ((inv.get("CustomerRef") or {}).get("name") or "")
    m = PROJ_RE.search(name)
    return m.group(0).upper().replace(" ", "") if m else ""


def main() -> int:
    ap = argparse.ArgumentParser(description="Last-invoice date per job")
    ap.add_argument("--division", default="", help="CP / RP / MFD (default: all)")
    ap.add_argument("--closed-in", type=int, default=None,
                    help="list only jobs the WIP master calls Closed whose last "
                         "invoice falls in this year")
    ap.add_argument("--quiet-days", type=int, default=60,
                    help="RP completion rule: a job counts as finished when its "
                         "last invoice is at least this old and no cost has "
                         "landed since (default 60)")
    ap.add_argument("--finished-in", type=int, default=None,
                    help="list jobs INFERRED finished in this year by the "
                         "quiet-days rule - the only route that works for RP")
    ap.add_argument("--wip-master", default="")
    a = ap.parse_args()

    # WIP status first - it is a local file read, so a bad path fails before auth
    status = {}
    try:
        import importlib.util
        spec = importlib.util.spec_from_file_location(
            "ppe", str(Path(__file__).resolve().parent.parent
                       / "project-pnl" / "project_pnl_export.py"))
        ppe = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(ppe)
        wm = ppe.load_wip_master(Path(a.wip_master or str(ppe.DEFAULT_WIP_MASTER)))
        status = {k: str(v.get("status") or "").strip() for k, v in wm.items()}
    except Exception as e:                      # the report still works without it
        print(f"  ⚠ WIP status unavailable ({e}) — showing dates only")

    access, company_id = load_credentials()
    invoices = query_all(access, company_id, "Invoice")
    jobs: dict = defaultdict(lambda: {"first": "", "last": "", "n": 0, "amt": 0.0})
    for inv in invoices:
        p = _proj_of(inv)
        if not p:
            continue
        d = str(inv.get("TxnDate") or "")
        j = jobs[p]
        j["n"] += 1
        j["amt"] += float(inv.get("TotalAmt") or 0)
        if not j["first"] or d < j["first"]:
            j["first"] = d
        if d > j["last"]:
            j["last"] = d

    costs = _last_costs()
    today = dt.date.today()
    for p, v in jobs.items():
        v["last_cost"] = costs.get(p, "")
        try:
            quiet = (today - dt.date.fromisoformat(v["last"])).days
        except ValueError:
            quiet = -1
        v["quiet"] = quiet
        # finished = invoiced, gone quiet, and nothing spent since that invoice
        v["finished"] = (quiet >= a.quiet_days
                         and (not v["last_cost"] or v["last_cost"] <= v["last"]))

    div = a.division.upper()
    rows = [(p, v) for p, v in jobs.items() if not div or p.startswith(div)]
    if a.finished_in:
        rows = [(p, v) for p, v in rows
                if v["finished"] and v["last"][:4] == str(a.finished_in)]
        print(f"\n{div or 'ALL'} jobs INFERRED finished in {a.finished_in} "
              f"(last invoice ≥{a.quiet_days}d old, no cost since):")
    if a.closed_in:
        rows = [(p, v) for p, v in rows
                if status.get(p) == "Closed"
                and v["last"][:4] == str(a.closed_in)]
        print(f"\n{div or 'ALL'} jobs the WIP master calls Closed whose LAST "
              f"INVOICE is in {a.closed_in}:")
    rows.sort(key=lambda x: x[1]["last"], reverse=True)

    print(f"\n{'job':12} {'status':8} {'first inv':11} {'last inv':11} "
          f"{'last cost':11} {'quiet':>6} {'invs':>5} {'billed':>14}  done?")
    for p, v in rows:
        print(f"{p:12} {status.get(p, '?'):8} {v['first']:11} {v['last']:11} "
              f"{v['last_cost'] or '-':11} {v['quiet']:6} {v['n']:5} "
              f"{v['amt']:14,.2f}  {'yes' if v['finished'] else ''}")
    print(f"\n{len(rows)} job(s)  ·  {len(invoices):,} invoices scanned")
    if not (a.closed_in or a.finished_in):
        for yr in ("2024", "2025", "2026"):
            n = sum(1 for p, v in rows
                    if status.get(p) == "Closed" and v["last"][:4] == yr)
            f = sum(1 for p, v in rows
                    if v["finished"] and v["last"][:4] == yr)
            print(f"  {yr}:  WIP-master Closed {n:4}   inferred finished {f:4}")
        live = sum(1 for p, v in rows if not v["finished"])
        print(f"  still running (not quiet {a.quiet_days}d): {live}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
