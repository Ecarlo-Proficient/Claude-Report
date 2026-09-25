#!/usr/bin/env python3
"""
job_reality_audit.py — the full cost-to-invoice history of every job, and what
it says about which jobs are REAL.

Three questions the owner asked (2026-09-02), all answered from ONE read-only
QBO pull:

  1. WHICH JOBS ARE REAL - a job number in QBO is not proof a job happened.
     Laying every cost line beside every invoice, in date order, separates real
     work from data: a job with costs and no invoice is unbilled work (or a
     miscode), a job with an invoice and no costs is a pass-through (or someone
     else's costs), and costs landing long after the final invoice are leakage
     onto a job that was already closed out.

  3. WHICH SLAB JOBS CARRY FLATWORK - an `FW` cost code belongs on a `-FTW`
     project. On a base RP slab it inflates that slab's cost-to-date and starves
     the -FTW twin, so both look wrong and the slab looks over budget. Same rule
     as bill-tracker's `Audit - FW Misplaced` (the owner 2026-08-06), but rolled
     up PER JOB with dollars, which is what you need to fix it.

WHY IT DOES ITS OWN PULL
`ledger/load_costs.py` pulls exactly the same lines but WRITES them, scoped to
the ~173 projects on the WIP master. The universe here is every project QBO
knows about (745+ have invoices), and other sessions work in that ledger DB, so
this reads and writes nothing but a cache and stdout.

The pull is cached to the scratchpad; `--refresh` re-pulls. Read-only against
QBO and against the ledger.

USAGE
  python3 one-offs/job_reality_audit.py                 # all three sections
  python3 one-offs/job_reality_audit.py --section fw    # just the flatwork one
  python3 one-offs/job_reality_audit.py --division RP --refresh
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import re
import sys
from collections import defaultdict
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))
from shared import qbo_costs as qc                      # noqa: E402
from shared.qbo_api import (build_project_customer_map, load_credentials,      # noqa: E402
                            project_of_invoice, query_all)

CACHE = Path(os.environ.get(
    "ACB_AUDIT_CACHE",
    Path(os.environ.get("TMPDIR", "/tmp")) / "job_reality_audit.json"))

# An FW cost code is flatwork. Same shape bill-tracker matches on.
FW_RE = re.compile(r"^FW\d+$", re.IGNORECASE)


def pull(refresh: bool) -> dict:
    """{invoices: [...], costs: [...]} — cached, because three sections and any
    number of re-slices should not cost three QBO pulls."""
    if CACHE.exists() and not refresh:
        d = json.loads(CACHE.read_text())
        print(f"  cache: {len(d['invoices']):,} invoices · {len(d['costs']):,} "
              f"cost lines (pulled {d.get('pulled_at', '?')})  [--refresh to re-pull]")
        return d
    print("  Authenticating to QBO (Touch ID)…")
    access, company_id = load_credentials()
    print("  authenticated.")            # never echo the realm (owner 2026-08-06)
    raw = query_all(access, company_id, "Invoice")
    invoices = [{"proj": project_of_invoice(i), "date": str(i.get("TxnDate") or ""),
                 "amt": float(i.get("TotalAmt") or 0), "doc": i.get("DocNumber") or "",
                 "bal": float(i.get("Balance") or 0),
                 "cust": ((i.get("CustomerRef") or {}).get("name") or ""),
                 "memo": (i.get("PrivateNote") or "")[:120]}
                for i in raw]
    unattr = [i for i in invoices if not i["proj"]]
    invoices = [i for i in invoices if i["proj"]]
    print(f"  {len(invoices):,} of {len(raw):,} invoices attributed to a project "
          f"({len(unattr):,} name no job, or name two)")
    account_names = qc.build_account_map(access, company_id)
    proj_map = build_project_customer_map(access, company_id)
    cust_to_proj = {v["id"]: p for p, v in proj_map.items()}
    costs = [{"proj": r["project_no"], "date": r["txn_date"], "amt": r["amount"],
              "code": r["cost_code"] or "", "acct": r["account"] or "",
              "vendor": r["vendor"] or "", "sub": r["is_sub"],
              "txn": r["qbo_txn_id"], "type": r.get("txn_type") or "",
              "doc": r.get("doc_number") or "", "line": r.get("line_no") or "",
              "desc": (r["description"] or "")[:80],
              "memo": (r.get("memo") or "")[:80]}
             for r in qc.iter_cost_lines(access, company_id, account_names,
                                         cust_to_proj)
             if r["project_no"]]
    print(f"  {len(costs):,} cost lines attributed to a project")
    d = {"pulled_at": dt.datetime.now().isoformat(timespec="seconds"),
         "realm": company_id, "invoices": invoices, "costs": costs}
    CACHE.parent.mkdir(parents=True, exist_ok=True)
    CACHE.write_text(json.dumps(d))
    return d


def roll(data: dict, division: str) -> dict:
    """One row per job: the cost and invoice history collapsed to its edges."""
    jobs: dict = defaultdict(lambda: {
        "cost": 0.0, "billed": 0.0, "open": 0.0, "n_cost": 0, "n_inv": 0,
        "cost_lo": "", "cost_hi": "", "inv_lo": "", "inv_hi": "",
        "fw": 0.0, "n_fw": 0, "fw_codes": set(), "sub": 0.0})
    for c in data["costs"]:
        p = c["proj"].upper()
        if division and not p.startswith(division):
            continue
        j = jobs[p]
        j["cost"] += c["amt"]; j["n_cost"] += 1
        if c["sub"]:
            j["sub"] += c["amt"]
        d = c["date"] or ""
        if d and (not j["cost_lo"] or d < j["cost_lo"]):
            j["cost_lo"] = d
        if d > j["cost_hi"]:
            j["cost_hi"] = d
        leaf = (c["code"] or "").split(":")[-1].strip()
        if FW_RE.match(leaf):
            j["fw"] += c["amt"]; j["n_fw"] += 1; j["fw_codes"].add(leaf.upper())
    for i in data["invoices"]:
        p = i["proj"].upper()
        if division and not p.startswith(division):
            continue
        j = jobs[p]
        j["billed"] += i["amt"]; j["open"] += i["bal"]; j["n_inv"] += 1
        d = i["date"] or ""
        if d and (not j["inv_lo"] or d < j["inv_lo"]):
            j["inv_lo"] = d
        if d > j["inv_hi"]:
            j["inv_hi"] = d
    return jobs


def _money(x) -> str:
    return f"{x:,.0f}"


def section_real(jobs: dict, limit: int) -> None:
    print("\n" + "=" * 78)
    print("1. WHICH JOBS ARE REAL — cost history vs invoice history")
    print("=" * 78)
    costs_no_inv, inv_no_costs, after_final = [], [], []
    for p, j in jobs.items():
        if j["cost"] > 1 and j["n_inv"] == 0:
            costs_no_inv.append((p, j))
        elif j["billed"] > 1 and j["n_cost"] == 0:
            inv_no_costs.append((p, j))
        elif j["inv_hi"] and j["cost_hi"] > j["inv_hi"]:
            late = j["cost_hi"]
            gap = (dt.date.fromisoformat(late) - dt.date.fromisoformat(j["inv_hi"])).days
            if gap >= 45:
                after_final.append((p, j, gap))

    print(f"\n  COSTS BUT NEVER INVOICED — work went out, nothing billed  "
          f"({len(costs_no_inv)} jobs, {_money(sum(j['cost'] for _p, j in costs_no_inv))})")
    print(f"  {'job':14} {'cost':>12} {'lines':>6}  {'first cost':11} {'last cost':11}")
    for p, j in sorted(costs_no_inv, key=lambda x: -x[1]["cost"])[:limit]:
        print(f"  {p:14} {_money(j['cost']):>12} {j['n_cost']:6}  "
              f"{j['cost_lo']:11} {j['cost_hi']:11}")

    print(f"\n  INVOICED BUT NO COSTS — billed with nothing spent  "
          f"({len(inv_no_costs)} jobs, {_money(sum(j['billed'] for _p, j in inv_no_costs))})")
    print(f"  {'job':14} {'billed':>12} {'invs':>6}  {'first inv':11} {'last inv':11}")
    for p, j in sorted(inv_no_costs, key=lambda x: -x[1]["billed"])[:limit]:
        print(f"  {p:14} {_money(j['billed']):>12} {j['n_inv']:6}  "
              f"{j['inv_lo']:11} {j['inv_hi']:11}")

    print(f"\n  COSTS AFTER THE FINAL INVOICE (45d+) — leakage onto a closed job  "
          f"({len(after_final)} jobs)")
    print(f"  {'job':14} {'cost':>12} {'billed':>12}  {'last inv':11} {'last cost':11} {'gap':>5}")
    for p, j, gap in sorted(after_final, key=lambda x: -x[2])[:limit]:
        print(f"  {p:14} {_money(j['cost']):>12} {_money(j['billed']):>12}  "
              f"{j['inv_hi']:11} {j['cost_hi']:11} {gap:5}d")


def section_fw(jobs: dict, limit: int) -> None:
    print("\n" + "=" * 78)
    print("3. SLAB JOBS CARRYING FLATWORK — FW codes off their -FTW twin")
    print("=" * 78)
    rows = [(p, j) for p, j in jobs.items()
            if j["fw"] > 1 and not p.endswith("-FTW")]
    rows.sort(key=lambda x: -x[1]["fw"])
    tot = sum(j["fw"] for _p, j in rows)
    print(f"\n  {len(rows)} jobs carrying {_money(tot)} of flatwork cost that is not "
          f"on a -FTW project.")
    print(f"\n  {'job':14} {'FW cost':>11} {'of job cost':>12} {'%':>6} "
          f"{'twin':14} {'twin cost':>11}  codes")
    for p, j in rows[:limit]:
        twin = f"{p}-FTW"
        tw = jobs.get(twin)
        pct = (j["fw"] / j["cost"] * 100) if j["cost"] else 0
        print(f"  {p:14} {_money(j['fw']):>11} {_money(j['cost']):>12} {pct:5.1f}% "
              f"{(twin if tw else '— none —'):14} "
              f"{(_money(tw['cost']) if tw else ''):>11}  "
              f"{','.join(sorted(j['fw_codes']))}")
    print(f"\n  A slab carrying its own flatwork reads over budget while the -FTW "
          f"twin reads under.\n  Where there is no twin, the flatwork was never "
          f"split out as its own job at all.")


def main() -> int:
    ap = argparse.ArgumentParser(description="Cost-to-invoice reality audit")
    ap.add_argument("--division", default="", help="CP / RP / MFD (default: all)")
    ap.add_argument("--section", choices=["real", "fw", "all"], default="all")
    ap.add_argument("--limit", type=int, default=25)
    ap.add_argument("--refresh", action="store_true", help="re-pull from QBO")
    a = ap.parse_args()

    data = pull(a.refresh)
    jobs = roll(data, a.division.upper())
    print(f"\n  {len(jobs)} job(s) with any cost or invoice history"
          f"{' in ' + a.division.upper() if a.division else ''}")
    if a.section in ("real", "all"):
        section_real(jobs, a.limit)
    if a.section in ("fw", "all"):
        section_fw(jobs, a.limit)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
