#!/usr/bin/env python3
"""load_bill_payments.py — QBO BillPayment (money OUT to vendors) → ledger.bill_payment (+ _line).

The Vendor page shows a vendor's payments this year WITHOUT going into QuickBooks. A BillPayment is
the real payment event (one cheque/ACH can pay several bills); Line[].LinkedTxn holds the bills it
covered. This lands them in the LOCAL ledger DB (the same private file that already holds costs /
money-in payments — never pushed, never cloud), so the vendor page reads them on demand and the
refresh stays light — no bulk-load bloat, no new place data lives.

Read-only on QBO. Never echoes the realm. Idempotent: DELETE the window, reload.

    python3 ledger/load_bill_payments.py                 # this calendar year
    python3 ledger/load_bill_payments.py --since 2025-01-01
    python3 ledger/load_bill_payments.py --dry-run       # pull + report, write nothing
    python3 ledger/load_bill_payments.py --selftest      # offline proof, no QBO
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

HERE = Path(__file__).resolve().parent
SCHEMA_SQL = HERE / "schema.sql"

DEFAULT_DB = paths.get_path(
    "ACB_LEDGER_DB",
    Path.home() / "Library" / "Application Support" / "Proficient" / "ledger.sqlite3",
)

_BP_COLS = ("qbo_txn_id", "txn_date", "vendor", "vendor_id",
            "total_amt", "pay_type", "ref_no", "source", "loaded_at")
_BPL_COLS = ("payment_id", "bill_id", "amount")


def _num(x) -> float:
    try:
        return round(float(x), 2)
    except (TypeError, ValueError):
        return 0.0


def _connect(db_path: Path) -> sqlite3.Connection:
    con = sqlite3.connect(str(db_path))
    con.row_factory = sqlite3.Row
    con.executescript(SCHEMA_SQL.read_text(encoding="utf-8"))   # CREATE TABLE IF NOT EXISTS → new tables land
    return con


def _rows_from_billpayment(bp: dict) -> tuple[dict, list[dict]]:
    """One QBO BillPayment → (payment row, [line rows]). Each line's LinkedTxn of type Bill is a bill
    this payment covered; the line Amount is the slice applied to that bill."""
    pid = str(bp.get("Id") or "")
    ven = bp.get("VendorRef") or {}
    prow = {
        "qbo_txn_id": pid,
        "txn_date": (bp.get("TxnDate") or "")[:10] or None,
        "vendor": ven.get("name"), "vendor_id": ven.get("value"),
        "total_amt": _num(bp.get("TotalAmt")),
        "pay_type": bp.get("PayType"),
        "ref_no": bp.get("DocNumber"),
    }
    lines = []
    for ln in bp.get("Line") or []:
        amt = _num(ln.get("Amount"))
        for lt in ln.get("LinkedTxn") or []:
            if lt.get("TxnType") == "Bill":
                lines.append({"payment_id": pid, "bill_id": str(lt.get("TxnId") or ""), "amount": amt})
    return prow, lines


def _load(con: sqlite3.Connection, since: str, dry_run: bool) -> int:
    from shared.qbo_api import load_credentials, query_all
    access, company_id = load_credentials()
    print("  authenticated.")                       # never echo the realm/company id (owner 2026-08-06)
    raw = query_all(access, company_id, "BillPayment", f"TxnDate >= '{since}'")
    print(f"  pulled {len(raw)} bill payments since {since}.")

    payments, lines = [], []
    for bp in raw:
        prow, lrows = _rows_from_billpayment(bp)
        if not prow["qbo_txn_id"]:
            continue
        payments.append(prow)
        lines.extend(lrows)

    total = round(sum(p["total_amt"] for p in payments), 2)
    print(f"  {len(payments)} payments · {total:,.2f} paid out · {len(lines)} bill links.")
    if dry_run:
        print("  --dry-run: nothing written.")
        return len(payments)

    now = dt.datetime.now().isoformat(timespec="seconds")
    con.execute("DELETE FROM bill_payment_line")
    con.execute("DELETE FROM bill_payment")
    pph = ", ".join(f":{c}" for c in _BP_COLS)
    for p in payments:
        row = {**p, "source": "qbo_billpayment", "loaded_at": now}
        con.execute(f"INSERT OR REPLACE INTO bill_payment ({', '.join(_BP_COLS)}) VALUES ({pph})",
                    {c: row.get(c) for c in _BP_COLS})
    lph = ", ".join(f":{c}" for c in _BPL_COLS)
    seen = set()
    for ln in lines:
        k = (ln["payment_id"], ln["bill_id"])
        if k in seen:                              # PK is (payment, bill); collapse dup links
            continue
        seen.add(k)
        con.execute(f"INSERT OR REPLACE INTO bill_payment_line ({', '.join(_BPL_COLS)}) VALUES ({lph})",
                    {c: ln.get(c) for c in _BPL_COLS})
    con.commit()
    print(f"  wrote {len(payments)} payments + {len(seen)} bill links.")
    return len(payments)


def _selftest() -> int:
    """Offline proof: schema applies and a synthetic BillPayment + its two bill links round-trip."""
    with tempfile.TemporaryDirectory() as td:
        con = _connect(Path(td) / "t.sqlite3")
        bp = {"Id": "BP1", "TxnDate": "2026-08-14", "TotalAmt": "12500.00", "PayType": "Check",
              "DocNumber": "5567", "VendorRef": {"name": "COWTOWN REDI MIX CONCRETE", "value": "42"},
              "Line": [
                  {"Amount": "10000.00", "LinkedTxn": [{"TxnId": "BILL_A", "TxnType": "Bill"}]},
                  {"Amount": "2500.00", "LinkedTxn": [{"TxnId": "BILL_B", "TxnType": "Bill"}]},
              ]}
        prow, lrows = _rows_from_billpayment(bp)
        assert prow["total_amt"] == 12500.0 and prow["vendor_id"] == "42" and prow["ref_no"] == "5567", prow
        assert len(lrows) == 2 and {x["bill_id"] for x in lrows} == {"BILL_A", "BILL_B"}, lrows
        now = "x"
        con.execute("INSERT INTO bill_payment (" + ", ".join(_BP_COLS) + ") VALUES (" + ", ".join("?" * len(_BP_COLS)) + ")",
                    [prow["qbo_txn_id"], prow["txn_date"], prow["vendor"], prow["vendor_id"],
                     prow["total_amt"], prow["pay_type"], prow["ref_no"], "qbo_billpayment", now])
        for ln in lrows:
            con.execute("INSERT INTO bill_payment_line (payment_id, bill_id, amount) VALUES (?,?,?)",
                        [ln["payment_id"], ln["bill_id"], ln["amount"]])
        got = con.execute("SELECT COUNT(*) n, SUM(amount) s FROM bill_payment_line WHERE payment_id='BP1'").fetchone()
        assert got["n"] == 2 and round(got["s"], 2) == 12500.0, dict(got)
        print("selftest OK — BillPayment + 2 bill links round-trip.")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(description="Load QBO BillPayment (money out to vendors) into the ledger.")
    ap.add_argument("--db", type=Path, default=DEFAULT_DB)
    ap.add_argument("--since", default=f"{dt.date.today().year}-01-01",
                    help="pull payments on/after this date (default: Jan 1 of this year).")
    ap.add_argument("--dry-run", action="store_true", help="pull + report, write nothing.")
    ap.add_argument("--selftest", action="store_true", help="offline proof; no QBO.")
    args = ap.parse_args()

    if args.selftest:
        return _selftest()

    con = _connect(args.db)
    print(f"Bill payments → {args.db}")
    try:
        _load(con, args.since, args.dry_run)
    finally:
        con.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
