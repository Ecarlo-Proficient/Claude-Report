"""load_attachments.py - every QBO attachment, indexed into the ledger (owner 2026-09-03: "every
transaction needs an attachment that we can view on the ledger and open ... invoices, bill scans,
etc need attachments everywhere from qbo").

QBO's Attachable entity links a file to one or more transactions (Bill, Purchase, Invoice, Payment,
BillPayment, ...). The shared resolver (`shared/qbo_attachments`) already sweeps them into a
week-cached disk index; this loader writes that index into the ledger as `attachment(etype, txn_id,
attachable_id, file_name)` so EVERY row the dashboard renders can carry a scan count with one join,
and the click resolves a fresh (minutes-lived) download link through /api/attachment.

Read-only on QBO. `--refresh` forces a new sweep instead of the disk cache. `--selftest` offline.
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

from shared import paths, qbo_attachments  # noqa: E402

LEDGER_DB = Path(paths.get_path("ACB_LEDGER_DB", Path.home() / "Library" / "Application Support" / "Proficient" / "ledger.sqlite3"))
SCHEMA_SQL = PROJECT_ROOT / "ledger" / "schema.sql"
_KEEP = ("Bill", "Purchase", "Invoice", "Payment", "BillPayment", "Estimate", "CreditMemo", "Deposit", "JournalEntry", "SalesReceipt", "VendorCredit")


def _connect(db_path: Path) -> sqlite3.Connection:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(str(db_path))
    con.executescript(SCHEMA_SQL.read_text(encoding="utf-8"))
    return con


def write_index(con: sqlite3.Connection, idx: dict, now: str) -> dict:
    """Full replace of `attachment` from the (etype, txn_id) -> [{Id, FileName}] index."""
    rows = []
    for (etype, txn_id), files in idx.items():
        if etype not in _KEEP:
            continue
        for f in files:
            rows.append((etype, str(txn_id), str(f.get("Id") or ""), f.get("FileName") or "", now))
    con.execute("DELETE FROM attachment")
    con.executemany("INSERT OR REPLACE INTO attachment (etype, txn_id, attachable_id, file_name, loaded_at) VALUES (?,?,?,?,?)", rows)
    con.commit()
    by = {}
    for r in rows:
        by[r[0]] = by.get(r[0], 0) + 1
    return {"files": len(rows), "by_entity": by}


def run(db_path: Path, refresh: bool, dry_run: bool) -> None:
    from shared.qbo_api import load_credentials, query_all
    access, company_id = load_credentials()
    print("  authenticated.")                          # never echo the realm / company id
    idx = qbo_attachments.build_index(access, company_id, query_all, force=refresh)
    n = sum(len(v) for v in idx.values())
    print(f"Index: {len(idx)} transactions with files · {n} file links" + (" (fresh sweep)" if refresh else " (disk cache if fresh, else swept)"))
    if dry_run:
        print("dry run - nothing written"); return
    con = _connect(db_path)
    try:
        res = write_index(con, idx, dt.datetime.now().isoformat(timespec="seconds"))
    finally:
        con.close()
    print(f"Wrote {res['files']} attachment rows -> {db_path}")
    for k, v in sorted(res["by_entity"].items(), key=lambda x: -x[1]):
        print(f"  {k:<12} {v:>7}")


def _selftest() -> None:
    idx = {("Bill", "1"): [{"Id": "a", "FileName": "scan.pdf"}, {"Id": "b", "FileName": "scan2.pdf"}],
           ("Invoice", "9"): [{"Id": "c", "FileName": "draw.pdf"}], ("Vendor", "5"): [{"Id": "d", "FileName": "w9.pdf"}]}
    with tempfile.TemporaryDirectory() as d:
        con = _connect(Path(d) / "t.sqlite3")
        res = write_index(con, idx, "t")
        assert res == {"files": 3, "by_entity": {"Bill": 2, "Invoice": 1}}, res
        n = con.execute("SELECT COUNT(*) FROM attachment WHERE etype='Bill' AND txn_id='1'").fetchone()[0]
        assert n == 2
        res2 = write_index(con, {("Bill", "1"): [{"Id": "a", "FileName": "scan.pdf"}]}, "t2")   # full replace
        assert res2["files"] == 1 and con.execute("SELECT COUNT(*) FROM attachment").fetchone()[0] == 1
        con.close()
    print("load_attachments selftest OK - index -> attachment rows, non-transaction entities skipped, full replace")


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="Index every QBO attachment into the ledger (attachment table).")
    ap.add_argument("--db", default=str(LEDGER_DB))
    ap.add_argument("--refresh", action="store_true", help="Force a fresh Attachable sweep (ignore the week-old disk cache).")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--selftest", action="store_true")
    a = ap.parse_args()
    if a.selftest:
        _selftest()
    else:
        run(Path(a.db), a.refresh, a.dry_run)
