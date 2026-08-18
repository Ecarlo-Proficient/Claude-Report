#!/usr/bin/env python3
"""
load_sub_loc.py - load the subcontractor LOC float model into the ledger.

Runs the shared engine (shared/sub_loc.py: read-only QBO pull -> chronological per-project
FIFO) and writes the result into the ledger so the dashboard's "Sub LOC" tab can show:
  * outstanding  = how much we have FRONTED to subs and NOT yet collected (today's float)
  * peak         = the high-water float = the LOC you truly need
  * the DRAW/REPAY timeline with running balance, and which client payment settled which subs

Full-replaces `sub_loc_event` and upserts the single `sub_loc_run` summary each run. The
ledger is the mirror; QBO stays the source. One Touch ID per run (the QBO pull).

USAGE
  python3 ledger/load_sub_loc.py                       # window = first Friday 3 months back -> today
  python3 ledger/load_sub_loc.py --months 6            # widen the window
  python3 ledger/load_sub_loc.py --start 2026-04-03    # explicit window start
  python3 ledger/load_sub_loc.py --dry-run             # pull + compute, write nothing
  python3 ledger/load_sub_loc.py --selftest            # offline proof (no QBO), throwaway DB
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import sqlite3
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
PROJECT_ROOT = HERE.parent
sys.path.insert(0, str(PROJECT_ROOT))
from shared import paths          # noqa: E402
from shared import sub_loc as sl  # noqa: E402

SCHEMA_SQL = HERE / "schema.sql"
DEFAULT_DB = paths.get_path(
    "ACB_LEDGER_DB",
    Path.home() / "Library" / "Application Support" / "Proficient" / "ledger.sqlite3")


def _connect(db_path: Path) -> sqlite3.Connection:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(str(db_path))
    con.row_factory = sqlite3.Row
    con.executescript(SCHEMA_SQL.read_text(encoding="utf-8"))
    # Migrate an existing sub_loc_run (CREATE TABLE IF NOT EXISTS won't add a new column):
    # add open_by_project so a re-run on a ledger from before the drill-down doesn't crash.
    have = {r[1] for r in con.execute("PRAGMA table_info(sub_loc_run)")}
    if "open_by_project" not in have:
        con.execute("ALTER TABLE sub_loc_run ADD COLUMN open_by_project TEXT")
        con.commit()
    return con


def _iso(d) -> str:
    return d.isoformat() if hasattr(d, "isoformat") else (str(d) if d else "")


def write(con: sqlite3.Connection, events, summary, projects, start, today) -> None:
    """Full-replace the event timeline; upsert the single run summary."""
    now = dt.datetime.now().isoformat(timespec="seconds")
    con.execute("DELETE FROM sub_loc_event")
    rows = []
    for i, e in enumerate(events):
        reimb = json.dumps([[inv, _iso(d)] for inv, d in e.get("reimb", [])]) if e.get("reimb") else None
        rows.append((
            i, _iso(e["date"]), e["type"], e.get("project"), sl.division_of(e.get("project") or ""),
            e.get("party"), e.get("out", 0.0), e.get("inn", 0.0), e.get("lag"),
            e.get("balance", 0.0), e.get("note") or None, e.get("invoice") or None, reimb, now))
    con.executemany(
        "INSERT INTO sub_loc_event (seq, event_date, type, project, division, party, out_amt, "
        "in_amt, lag_days, balance, note, invoice, reimb, loaded_at) "
        "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)", rows)
    con.execute(
        "INSERT INTO sub_loc_run (id, window_start, window_end, peak, peak_date, outstanding, "
        "total_drawn, total_repaid, prefunded, avg_lag, n_draws, divisions, projects, open_by_project, loaded_at) "
        "VALUES (1,?,?,?,?,?,?,?,?,?,?,?,?,?,?) "
        "ON CONFLICT(id) DO UPDATE SET window_start=excluded.window_start, window_end=excluded.window_end, "
        "peak=excluded.peak, peak_date=excluded.peak_date, outstanding=excluded.outstanding, "
        "total_drawn=excluded.total_drawn, total_repaid=excluded.total_repaid, prefunded=excluded.prefunded, "
        "avg_lag=excluded.avg_lag, n_draws=excluded.n_draws, divisions=excluded.divisions, "
        "projects=excluded.projects, open_by_project=excluded.open_by_project, loaded_at=excluded.loaded_at",
        (_iso(start), _iso(today), summary["peak"], _iso(summary["peak_date"]), summary["outstanding"],
         summary["total_drawn"], summary["total_repaid"], summary["prefunded"], summary["avg_lag"],
         summary["n_draws"], json.dumps(summary["divisions"], default=str),
         json.dumps(projects, default=str),
         json.dumps(summary.get("open_by_project", {}), default=str), now))
    con.commit()


def run(db_path: Path, months: int, start_override, dry_run: bool) -> None:
    today = dt.date.today()
    start = sl._parse(start_override) if start_override else sl.first_friday_months_back(today, months)
    print(f"\n  SUBCONTRACTOR LOC - float model")
    print(f"  window {start} -> {today}")
    from shared.qbo_api import load_credentials
    access, company_id = load_credentials()           # Touch ID; company_id never printed
    print("  pulling QBO (sub bills, invoices, sub payments, client payments) ...")
    events, summary, projects = sl.compute(access, company_id, start, today)
    print(f"    {summary['n_draws']} draw event(s), {summary['n_repay_chunks']} repay chunk(s)")
    print(f"  PEAK LOC needed:  ${summary['peak']:,.0f} on {summary['peak_date']}")
    print(f"  outstanding NOW:  ${summary['outstanding']:,.0f} fronted, still uncollected")
    print(f"  avg draw->repay:  {summary['avg_lag']:.1f} days")
    if dry_run:
        print("  (dry-run: nothing written)")
        return
    con = _connect(db_path)
    try:
        write(con, events, summary, projects, start, today)
    finally:
        con.close()
    print(f"  wrote {len(events)} event(s) + the run summary to {db_path}")


def selftest() -> int:
    """Offline proof of the engine + writer on a throwaway DB (no QBO)."""
    import tempfile
    D = dt.date
    draws = [
        {"date": D(2026, 6, 2), "project": "CP100", "party": "SubA", "amount": 100.0, "period": "2026-06"},
        {"date": D(2026, 6, 6), "project": "CP100", "party": "SubA", "amount": 50.0, "period": "2026-06"},
    ]
    repays = [
        {"date": D(2026, 6, 12), "project": "CP100", "party": "GC1", "amount": 80.0, "period": "2026-06", "invoice": "INV1"},
    ]
    # carry a bill id on the draws so the drill-down (open subs by draw) can be checked
    for i, d in enumerate(draws):
        d["bill_id"] = f"900{i}"; d["bill_ref"] = f"B-{i}"
    events, summary = sl.run_fifo(draws, repays)
    assert abs(summary["peak"] - 150.0) < 1e-6, summary["peak"]
    assert abs(summary["outstanding"] - 70.0) < 1e-6, summary["outstanding"]
    # the project drill-down: one CP100 group (2026-06 draw) with the still-open sub + bill id
    inv_meta = {"i1": {"project": "CP100", "draw_month": "2026-06", "doc": "INV1",
                       "total": 120.0, "balance": 50.0, "txn_date": "2026-06-10", "cust_id": "42"}}
    sl.attach_open_by_project(summary, inv_meta)
    obp = summary["open_by_project"]
    assert "CP100" in obp and obp["CP100"]["cust_id"] == "42", obp
    g = obp["CP100"]["groups"][0]
    assert g["draw"] and g["draw"]["status"] == "Partially Paid", g
    assert any(s["bill_id"] for s in g["subs"]), g["subs"]
    projects = sl.per_project(events)
    with tempfile.TemporaryDirectory() as td:
        con = _connect(Path(td) / "t.sqlite3")
        write(con, events, summary, projects, D(2026, 6, 1), D(2026, 6, 30))
        n = con.execute("SELECT COUNT(*) FROM sub_loc_event").fetchone()[0]
        run_row = con.execute("SELECT peak, outstanding FROM sub_loc_run WHERE id=1").fetchone()
        con.close()
    assert n == len(events), (n, len(events))
    assert abs(run_row[0] - 150.0) < 1e-6 and abs(run_row[1] - 70.0) < 1e-6, tuple(run_row)
    print(f"  selftest OK: {n} events, peak ${run_row[0]:,.0f}, outstanding ${run_row[1]:,.0f}")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(description="Load the subcontractor LOC float model into the ledger.")
    ap.add_argument("--db", type=Path, default=DEFAULT_DB)
    ap.add_argument("--months", type=int, default=3, help="months back to the first Friday (default 3)")
    ap.add_argument("--start", type=str, help="override window start YYYY-MM-DD")
    ap.add_argument("--dry-run", action="store_true", help="pull + compute; write nothing")
    ap.add_argument("--selftest", action="store_true", help="offline proof (no QBO), throwaway DB")
    args = ap.parse_args()
    if args.selftest:
        return selftest()
    run(args.db, args.months, args.start, args.dry_run)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
