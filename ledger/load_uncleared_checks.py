#!/usr/bin/env python3
"""load_uncleared_checks.py - every check QuickBooks shows as UNCLEARED -> ledger.uncleared_check.

The owner 2026-09-24: "a list of all unmatched checks and/or any checks that haven't been deposited".
QBO's API hides cleared status as a column (asking for it is silently dropped - 07/17) but the
TransactionList report FILTERS on it: `cleared=Uncleared` returns exactly the checks not yet matched /
cleared in QuickBooks. In QuickBooks "unmatched" and "not deposited" are the same flag.

Also writes `bank_match`: the newest CLEARED transaction per bank account = how far the bank-feed
matching has got. An uncleared check older than that date was not cashed, or was cashed and never
matched. An account with no cleared transaction in a year (Joint Checks Account) is marked
feed_matched=0 - its checks never clear, and the page keeps them apart.

Read-only on QBO (report calls). Never echoes the realm. Reloads both tables whole.

    python3 ledger/load_uncleared_checks.py
    python3 ledger/load_uncleared_checks.py --dry-run     pull + report, write nothing
"""
from __future__ import annotations

import argparse
import datetime as dt
import sqlite3
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))
from shared import paths                   # noqa: E402
from shared import qbo_mirror as mirror    # noqa: E402

HERE = Path(__file__).resolve().parent
SCHEMA_SQL = HERE / "schema.sql"
DEFAULT_DB = paths.get_path(
    "ACB_LEDGER_DB",
    Path.home() / "Library" / "Application Support" / "Proficient" / "ledger.sqlite3",
)
CHECK_TYPES = ("Check", "Bill Payment (Check)")
FIRST_YEAR = 2019
COLS = "tx_date,txn_type,doc_num,name,account_name,subt_nat_amount"


def _rows(rep: dict) -> list:
    cols = [c.get("ColTitle", "") for c in (rep.get("Columns") or {}).get("Column", [])]
    out = []

    def walk(block):
        for x in (block.get("Row") or []):
            if x.get("type") == "Data":
                cd = x.get("ColData") or []
                d = {cols[i]: c.get("value") for i, c in enumerate(cd) if i < len(cols)}
                d["_id"] = (cd[1] if len(cd) > 1 else {}).get("id")
                out.append(d)
            if x.get("Rows"):
                walk(x["Rows"])
    walk(rep.get("Rows") or {})
    return out


def _windows(today: dt.date):
    """Year by year - one report per year keeps each call well under QBO's row limits."""
    for y in range(FIRST_YEAR, today.year + 1):
        yield f"{y}-01-01", (f"{y}-12-31" if y < today.year else today.isoformat())


def pull(access: str, cid: str, bank_ids: list, today: dt.date) -> tuple:
    from shared.qbo_api import report
    acct = ",".join(bank_ids)
    unclr = []
    for a, b in _windows(today):
        unclr += _rows(report(access, cid, "TransactionList",
                              {"start_date": a, "end_date": b, "account": acct, "cleared": "Uncleared", "columns": COLS}))
    thru = {}
    start = (today - dt.timedelta(days=365)).isoformat()
    for r in _rows(report(access, cid, "TransactionList",
                          {"start_date": start, "end_date": today.isoformat(), "account": acct,
                           "cleared": "Cleared", "columns": COLS})):
        a = r.get("Account") or ""
        thru[a] = max(thru.get(a, ""), r.get("Date") or "")
    checks = [r for r in unclr if r.get("Transaction Type") in CHECK_TYPES]
    return checks, thru


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--db", type=Path, default=DEFAULT_DB)
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()
    from shared.qbo_api import load_credentials
    access, cid = load_credentials()
    mcon = mirror.connect()
    try:
        banks = {x["Id"]: x.get("Name") for x in mirror.load("Account", con=mcon) if x.get("AccountType") == "Bank"}
    finally:
        mcon.close()
    today = dt.date.today()
    checks, thru = pull(access, cid, list(banks), today)
    accounts = sorted({c.get("Account") or "" for c in checks} | set(thru))
    total = sum(-float(c.get("Amount") or 0) for c in checks)
    print(f"uncleared checks: {len(checks)} · ${total:,.2f}")
    for acc in accounts:
        n = sum(1 for c in checks if c.get("Account") == acc)
        print(f"  {acc[:34]:<34} {n:>5} uncleared · matched through {thru.get(acc) or 'never (not feed-matched)'}")
    if a.dry_run:
        print("dry run - nothing written")
        return 0
    now = dt.datetime.now().isoformat(timespec="seconds")
    con = sqlite3.connect(str(a.db))
    try:
        con.executescript(SCHEMA_SQL.read_text(encoding="utf-8"))
        with con:
            con.execute("DELETE FROM uncleared_check")
            con.execute("DELETE FROM bank_match")
            con.executemany(
                "INSERT INTO uncleared_check(qbo_txn_id, txn_type, txn_date, check_no, payee, account, amount, loaded_at) "
                "VALUES (?,?,?,?,?,?,?,?)",
                [(c.get("_id"), c.get("Transaction Type"), c.get("Date"), c.get("Num") or "", c.get("Name") or "",
                  c.get("Account") or "", round(-float(c.get("Amount") or 0), 2), now) for c in checks])
            con.executemany("INSERT INTO bank_match(account, matched_through, feed_matched, loaded_at) VALUES (?,?,?,?)",
                            [(acc, thru.get(acc), 1 if thru.get(acc) else 0, now) for acc in accounts])
    finally:
        con.close()
    print(f"wrote {len(checks)} checks, {len(accounts)} accounts -> {a.db.name}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
