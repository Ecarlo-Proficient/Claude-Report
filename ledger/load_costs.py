#!/usr/bin/env python3
"""
load_costs.py — land COMPLETE job costs (incl. subs) into the ledger, by cost code.

Pulls QBO expense transactions (Bills + Purchases) and writes one `cost_line` per
expense line, keyed to the project by its line `CustomerRef` and to the cost code
by the shared `cost_leaf()` resolver — the SAME engine project-pnl uses, so the
ledger and the P&L can never drift. This is the complete cost source Bill Tracker
couldn't be: subs are included, and reconciles to `wip_snapshot.costs_to_date`.

WHY IT NEEDS A QBO PULL
There is no on-disk artifact with current + complete costs (Bill Tracker excludes
subs; the P&L workbooks are stale and per-file). This tool reads QBO directly via
the shared vault — one Touch ID — and is READ-ONLY against QBO.

SAFETY
    * READ-ONLY against QBO (GET only, via shared/qbo_api).
    * Writes only the local ledger; scoped full-replace of source='qbo' cost_line
      for the target projects (idempotent; handles QBO deletions).
    * --dry-run pulls and reports the reconciliation WITHOUT writing.
    * --selftest runs the whole pipeline offline on a throwaway DB (no QBO, no
      touch to your real ledger) — proves the wiring before the first real pull.

USAGE
    python3 ledger/load_costs.py --selftest              # offline proof, no QBO
    python3 ledger/load_costs.py --active --show 20       # Active projects (Touch ID)
    python3 ledger/load_costs.py --division cp --dry-run  # pull + reconcile, write nothing
    python3 ledger/load_costs.py --project MFD177         # one project
    python3 ledger/load_costs.py --since 2025-01-01       # limit the pull window
"""
from __future__ import annotations

import argparse
import datetime as dt
import sqlite3
import sys
import tempfile
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))
from shared import paths  # noqa: E402
from shared import qbo_costs as qc  # noqa: E402

HERE = Path(__file__).resolve().parent
SCHEMA_SQL = HERE / "schema.sql"

DEFAULT_DB = paths.get_path(
    "ACB_LEDGER_DB",
    Path.home() / "Library" / "Application Support" / "Proficient" / "ledger.sqlite3",
)

DIVISION = {"cp": "Commercial", "rp": "Residential", "mfd": "Multi Family"}


def _connect(db_path: Path) -> sqlite3.Connection:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(db_path)
    con.execute("PRAGMA foreign_keys = ON;")
    con.executescript(SCHEMA_SQL.read_text(encoding="utf-8"))
    _migrate_cost_line(con)
    return con


def _migrate_cost_line(con) -> None:
    """cost_line predates its fleshed-out columns in DBs created by earlier loaders.
    It is a placeholder (empty until load_costs runs), so a shape upgrade is a safe
    drop + recreate. Refuse if it somehow already holds rows."""
    cols = {r[1] for r in con.execute("PRAGMA table_info(cost_line)")}
    if cols and "account" not in cols:
        if con.execute("SELECT COUNT(*) FROM cost_line").fetchone()[0]:
            sys.exit("cost_line has the legacy shape AND rows — refusing to auto-migrate. "
                     "Back up the ledger and migrate cost_line manually.")
        con.execute("DROP TABLE cost_line")
        con.executescript(SCHEMA_SQL.read_text(encoding="utf-8"))
        con.commit()


def target_projects(con, division: str | None, active: bool, projects: list[str] | None) -> set:
    """Which projects to load costs for — drawn from the ledger's own tables."""
    if projects:
        want = {p.upper() for p in projects}
        have = {r[0] for r in con.execute("SELECT project_no FROM project")}
        return want & have
    rows = con.execute("SELECT project_no, division, status FROM v_wip_latest").fetchall()
    out = set()
    for pn, div, status in rows:
        if division and div != DIVISION[division]:
            continue
        if active and not (status == "Active" or (div == "Multi Family" and status is None)):
            continue
        out.add(pn)
    return out


def write_cost_lines(con, records: list[dict], targets: set, now: str) -> dict:
    """Scoped full-replace of source='qbo' cost_line for the target projects.
    Upserts cost_code first (FK), then cost_line. Returns counts."""
    kept = [r for r in records if r["project_no"] in targets]
    # cost codes first (FK target)
    codes = {r["cost_code"] for r in kept if r["cost_code"]}
    for code in sorted(codes):
        m = qc.cost_code_meta(code)
        con.execute(
            "INSERT INTO cost_code (code, prefix, description) VALUES (:code,:prefix,:description) "
            "ON CONFLICT(code) DO UPDATE SET prefix=excluded.prefix, description=excluded.description",
            m,
        )
    # scoped replace so re-runs mirror QBO (and drop deleted txns)
    ph = ",".join("?" for _ in targets)
    if targets:
        con.execute(f"DELETE FROM cost_line WHERE source='qbo' AND project_no IN ({ph})", tuple(targets))
    cols = ["qbo_txn_id", "qbo_line_id", "txn_type", "project_no", "cost_code", "account",
            "amount", "txn_date", "is_sub", "vendor", "description", "source", "loaded_at"]
    ins = 0
    for r in kept:
        row = {**r, "source": "qbo", "loaded_at": now}
        con.execute(
            f"INSERT INTO cost_line ({','.join(cols)}) VALUES ({','.join(':'+c for c in cols)}) "
            f"ON CONFLICT(qbo_txn_id, qbo_line_id) DO UPDATE SET "
            + ", ".join(f"{c}=excluded.{c}" for c in cols if c not in ("qbo_txn_id", "qbo_line_id")),
            {c: row.get(c) for c in cols},
        )
        ins += 1
    con.commit()
    return {"lines": ins, "codes": len(codes),
            "skipped_off_target": len(records) - len(kept)}


def reconcile(con, targets: set, show: int) -> None:
    """Per-project: costs loaded here vs wip_snapshot.costs_to_date (the QBO truth)."""
    rows = con.execute("""
        SELECT p.project_no, s.costs_to_date AS wip, c.costs_loaded AS loaded, c.sub_costs, c.lines
        FROM project p
        LEFT JOIN v_cost_by_project c ON c.project_no = p.project_no
        LEFT JOIN wip_snapshot s ON s.project_no = p.project_no
        WHERE p.project_no IN (%s)
        ORDER BY COALESCE(c.costs_loaded,0) DESC
    """ % ",".join("?" for _ in targets), tuple(targets)).fetchall()
    reconciled = mism = 0
    print("\nReconcile — loaded cost vs WIP costs_to_date (the QBO truth):")
    shown = 0
    for pn, wip, loaded, subs, lines in rows:
        loaded = loaded or 0
        if wip:
            gap = abs(loaded - wip) / wip
            ok = gap <= 0.05
            reconciled += ok
            mism += (not ok)
            if shown < show:
                flag = "OK " if ok else f"GAP {gap*100:.0f}%"
                print(f"  {pn:<10} loaded ${loaded:>13,.0f}   WIP ${wip:>13,.0f}   "
                      f"sub ${subs or 0:>12,.0f}  {flag}")
                shown += 1
    print(f"\nReconciled within 5%: {reconciled}   mismatches: {mism}   "
          f"(projects with a WIP cost figure)")


def _selftest() -> None:
    """Prove the pipeline offline: fabricated txns → cost_lines → DB → views."""
    print("SELFTEST — offline, no QBO, throwaway DB.\n")
    account_names = {"a1": "Concrete"}
    customer_to_project = {"c1": "RP7358"}
    bills = [{
        "Id": "B1", "TxnDate": "2026-08-01", "PrivateNote": "RP7358 slab",
        "VendorRef": {"name": "Ready Mix Co"}, "Line": [
            {"Id": "1", "Amount": 5000, "AccountBasedExpenseLineDetail": {
                "CustomerRef": {"value": "c1"}, "AccountRef": {"value": "a1", "name": "Job Materials:Concrete"}}},
            {"Id": "2", "Amount": 8000, "ItemBasedExpenseLineDetail": {
                "CustomerRef": {"value": "c1"}, "ItemRef": {"name": "SL1"}}},
        ]}]
    purchases = [{
        "Id": "P1", "TxnDate": "2026-08-02", "PrivateNote": "RP7358 sub labor pour",
        "EntityRef": {"name": "Framing Crew"}, "Line": [
            {"Id": "1", "Amount": 12000, "ItemBasedExpenseLineDetail": {
                "CustomerRef": {"value": "c1"}, "ItemRef": {"name": "SL6"}}},
        ]}]
    records = list(qc.cost_lines_from_txns(bills, "Bill", "VendorRef", account_names, customer_to_project))
    records += list(qc.cost_lines_from_txns(purchases, "Expense", "EntityRef", account_names, customer_to_project))
    for r in records:
        print(f"  {r['txn_type']:<8} {r['project_no']} code={r['cost_code'] or '—':<5} "
              f"acct={r['account'] or '—':<10} ${r['amount']:>8,.0f} sub={r['is_sub']}  {r['vendor']}")

    with tempfile.TemporaryDirectory() as d:
        con = _connect(Path(d) / "selftest.sqlite3")
        con.execute("INSERT INTO project (project_no, division, is_ftw, updated_at) "
                    "VALUES ('RP7358','Residential',0,'t')")
        con.execute("INSERT INTO wip_snapshot (project_no, report_date, costs_to_date, loaded_at) "
                    "VALUES ('RP7358','2026-08-07',25000,'t')")
        con.commit()
        res = write_cost_lines(con, records, {"RP7358"}, "t")
        print(f"\nwrote: {res}")
        print("v_cost_by_code:")
        for row in con.execute("SELECT code, cost_code, actual, lines FROM v_cost_by_code WHERE project_no='RP7358' ORDER BY actual DESC"):
            print(f"  {row[0]:<14} cost_code={row[1] or '—':<5} ${row[2]:>8,.0f}  ({row[3]} lines)")
        reconcile(con, {"RP7358"}, show=5)
        con.close()
    print("\nSELFTEST OK — pipeline resolves codes, writes cost_line, reconciles.")


def run(db_path: Path, division, active, projects, since, dry_run, show):
    con = _connect(db_path)
    targets = target_projects(con, division, active, projects)
    if not targets:
        con.close()
        sys.exit("No target projects in the ledger. Load the WIP master first "
                 "(load_wip_master.py), or check --division/--project.")
    scope = (f"division {division}" if division else
             "active" if active else
             f"{len(projects)} named" if projects else "all")
    print(f"Target projects: {len(targets)} ({scope}){' since ' + since if since else ''}")

    print("Authenticating to QBO (Touch ID)…")
    from shared.qbo_api import load_credentials, build_project_customer_map
    access, company_id = load_credentials()
    print("  authenticated.")  # never echo the realm/company id (owner 2026-08-06)

    account_names = qc.build_account_map(access, company_id)
    proj_map = build_project_customer_map(access, company_id)
    customer_to_project = {v["id"]: p for p, v in proj_map.items()}
    print(f"  {len(account_names)} accounts · {len(customer_to_project)} project customers")

    records = list(qc.iter_cost_lines(access, company_id, account_names, customer_to_project, since))
    in_scope = [r for r in records if r["project_no"] in targets]
    print(f"Pulled {len(records)} cost lines · {len(in_scope)} in scope")

    if dry_run:
        print("\n--dry-run: nothing written. Reconciliation preview:")
        # reconcile against a temp copy so we don't touch the real ledger
        with tempfile.TemporaryDirectory() as d:
            tmp = _connect(Path(d) / "preview.sqlite3")
            for pn in targets:
                row = con.execute("SELECT division, is_ftw, name FROM project WHERE project_no=?", (pn,)).fetchone()
                if row:
                    tmp.execute("INSERT INTO project (project_no,division,is_ftw,name,updated_at) VALUES (?,?,?,?,'t')",
                                (pn, row[0], row[1], row[2]))
                s = con.execute("SELECT costs_to_date FROM wip_snapshot WHERE project_no=? ORDER BY report_date DESC LIMIT 1", (pn,)).fetchone()
                if s:
                    tmp.execute("INSERT INTO wip_snapshot (project_no,report_date,costs_to_date,loaded_at) VALUES (?,?,?,'t')",
                                (pn, "preview", s[0]))
            tmp.commit()
            write_cost_lines(tmp, records, targets, "preview")
            reconcile(tmp, targets, show)
            tmp.close()
        con.close()
        return

    now = dt.datetime.now().isoformat(timespec="seconds")
    res = write_cost_lines(con, records, targets, now)
    print(f"\nWrote {res['lines']} cost lines · {res['codes']} cost codes -> {db_path}")
    reconcile(con, targets, show)
    con.close()


def main():
    ap = argparse.ArgumentParser(description="Load complete QBO job costs (by cost code) into the ledger.")
    ap.add_argument("--db", type=Path, default=DEFAULT_DB)
    ap.add_argument("--division", choices=["cp", "rp", "mfd"], help="Only this division.")
    ap.add_argument("--active", action="store_true", help="Only Active projects (+ MFD).")
    ap.add_argument("--project", nargs="+", help="Only these project #s.")
    ap.add_argument("--since", help="Inclusive ISO date filter on TxnDate (e.g. 2025-01-01).")
    ap.add_argument("--dry-run", action="store_true", help="Pull + reconcile; write nothing.")
    ap.add_argument("--show", type=int, default=15, help="Reconciliation rows to print.")
    ap.add_argument("--selftest", action="store_true", help="Offline pipeline proof (no QBO).")
    args = ap.parse_args()
    if args.selftest:
        _selftest()
        return
    run(args.db, args.division, args.active, args.project, args.since, args.dry_run, args.show)


if __name__ == "__main__":
    main()
