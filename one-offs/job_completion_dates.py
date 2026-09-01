#!/usr/bin/env python3
"""
job_completion_dates.py — when did each job actually FINISH?

The question this exists to answer (the user 2026-09-01): "how are you going to
know which CP projects were completed in 2026?"

You cannot, from anything local:
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

Read-only. Writes nothing but stdout.

USAGE
  python3 one-offs/job_completion_dates.py                 # all divisions
  python3 one-offs/job_completion_dates.py --division CP --closed-in 2026
"""
from __future__ import annotations

import argparse
import sys
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from shared.qbo_api import PROJ_RE, load_credentials, query_all


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

    div = a.division.upper()
    rows = [(p, v) for p, v in jobs.items() if not div or p.startswith(div)]
    if a.closed_in:
        rows = [(p, v) for p, v in rows
                if status.get(p) == "Closed"
                and v["last"][:4] == str(a.closed_in)]
        print(f"\n{div or 'ALL'} jobs the WIP master calls Closed whose LAST "
              f"INVOICE is in {a.closed_in}:")
    rows.sort(key=lambda x: x[1]["last"], reverse=True)

    print(f"\n{'job':12} {'status':8} {'first inv':11} {'last inv':11} "
          f"{'invs':>5} {'billed':>15}")
    for p, v in rows:
        print(f"{p:12} {status.get(p, '?'):8} {v['first']:11} {v['last']:11} "
              f"{v['n']:5} {v['amt']:15,.2f}")
    print(f"\n{len(rows)} job(s)  ·  {len(invoices):,} invoices scanned")
    if not a.closed_in and status:
        for yr in ("2024", "2025", "2026"):
            n = sum(1 for p, v in rows
                    if status.get(p) == "Closed" and v["last"][:4] == yr)
            print(f"  Closed, last invoice in {yr}: {n}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
