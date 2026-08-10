#!/usr/bin/env python3
"""
load_invoices.py — land AR invoices (the draws the GC pays YOU) into the ledger.

Pulls every QBO Invoice and writes one `billing_event` per invoice, keyed to the
project by its `CustomerRef` and to the draw by its **Invoice # (DocNumber)** — the
same number `ap_bill_line.invoice_no` carries, so the Draws view can show
billed-to-GC (money IN) next to paid-to-vendors (money OUT) on every draw.

WHY A QBO PULL
Open_Invoices.xlsx has only the still-OPEN invoices; most draws on the board are
already GC-paid, so their amount lives only in QBO. This reads QBO directly via the
shared vault — one Touch ID — and is READ-ONLY against QBO.

SAFETY
    * READ-ONLY against QBO (GET only, via shared/qbo_api).
    * Writes only the local ledger; scoped full-replace of source='qbo_invoice'
      billing_event (idempotent; a re-run mirrors QBO, handling deletions).
    * --dry-run pulls and reports coverage WITHOUT writing.
    * --selftest runs the whole pipeline offline on a throwaway DB (no QBO).

USAGE
    python3 ledger/load_invoices.py --selftest     # offline proof, no QBO
    python3 ledger/load_invoices.py                # pull all invoices (Touch ID) + write
    python3 ledger/load_invoices.py --dry-run      # pull + coverage, write nothing
"""
from __future__ import annotations

import argparse
import datetime as dt
import re
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

# Project # anywhere in the CustomerRef name ("Parent:Project# Name"); -FTW is a
# distinct project, matched strictly (never family-rolled).
_PROJ_RE = re.compile(r"(MFD|CP|RP)\s*(\d+)(-FTW)?", re.IGNORECASE)
_DIV_PREFIX = {"MFD": "Multi Family", "CP": "Commercial", "RP": "Residential"}


def _extract_proj(text) -> str | None:
    if not text:
        return None
    m = _PROJ_RE.search(str(text))
    if not m:
        return None
    return (m.group(1) + m.group(2) + (m.group(3) or "")).upper()


def _division_for(proj) -> str | None:
    if not proj:
        return None
    for pre, div in _DIV_PREFIX.items():
        if proj.upper().startswith(pre):
            return div
    return None


def _num(v):
    try:
        return round(float(v), 2)
    except (TypeError, ValueError):
        return None


def _connect(db_path: Path) -> sqlite3.Connection:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(db_path)
    con.executescript(SCHEMA_SQL.read_text(encoding="utf-8"))
    _migrate_billing_event(con)
    return con


def _migrate_billing_event(con) -> None:
    """billing_event predates the AR columns (doc_number/amount/balance/…) in DBs
    created by earlier schema. It's a placeholder (empty until this loader runs), so
    a shape upgrade is a safe drop + recreate. Refuse if it somehow already holds rows."""
    cols = {r[1] for r in con.execute("PRAGMA table_info(billing_event)")}
    if cols and "doc_number" not in cols:
        if con.execute("SELECT COUNT(*) FROM billing_event").fetchone()[0]:
            sys.exit("billing_event has the legacy shape AND rows — refusing to auto-migrate. "
                     "Back up the ledger and migrate billing_event manually.")
        con.execute("DROP TABLE billing_event")
        con.executescript(SCHEMA_SQL.read_text(encoding="utf-8"))
        con.commit()


def parse_invoice(inv: dict) -> dict | None:
    """One QBO Invoice → a billing_event record. None only if there's no Id."""
    if not inv.get("Id"):
        return None
    cust = inv.get("CustomerRef") or {}
    proj = _extract_proj(cust.get("name"))
    memo = ((inv.get("CustomerMemo") or {}).get("value")) or inv.get("PrivateNote") or ""
    bal = _num(inv.get("Balance"))
    return {
        "qbo_txn_id": str(inv.get("Id")),
        "doc_number": (str(inv.get("DocNumber") or "").strip() or None),
        "project_no": proj,
        "division": _division_for(proj),
        "customer": cust.get("name"),
        "memo": (memo.strip() or None),
        "amount": _num(inv.get("TotalAmt")),
        "balance": bal,
        "txn_date": inv.get("TxnDate"),
        "status": ("Paid" if (bal is not None and bal <= 0.005) else "Open"),
        "draw_period": None,
        "source": "qbo_invoice",
    }


_COLS = ["qbo_txn_id", "doc_number", "project_no", "division", "customer", "memo",
         "amount", "balance", "txn_date", "status", "draw_period", "source", "loaded_at"]


def write_events(con, records: list[dict], now: str) -> int:
    con.execute("DELETE FROM billing_event WHERE source='qbo_invoice'")
    ph = ", ".join(f":{c}" for c in _COLS)
    for r in records:
        row = {**r, "loaded_at": now}
        con.execute(f"INSERT OR REPLACE INTO billing_event ({', '.join(_COLS)}) VALUES ({ph})",
                    {c: row.get(c) for c in _COLS})
    con.commit()
    return len(records)


def _draw_coverage(con, show: int) -> None:
    """How many MFD/CP draws on the board now have a matched AR invoice — the payoff."""
    row = con.execute(
        "SELECT COUNT(DISTINCT a.invoice_no) draws, "
        "       COUNT(DISTINCT b.doc_number) matched "
        "FROM ap_bill_line a LEFT JOIN billing_event b ON b.doc_number = a.invoice_no "
        "WHERE a.invoice_no IS NOT NULL AND a.invoice_no <> '' "
        "  AND COALESCE(a.project_no,'') NOT LIKE 'RP%' AND a.matched_invoice NOT LIKE '%— RP%'"
    ).fetchone()
    billed_in = con.execute("SELECT SUM(amount), SUM(balance) FROM billing_event WHERE source='qbo_invoice'").fetchone()
    print(f"\nDraw coverage: {row[1] or 0} of {row[0] or 0} MFD/CP draws now matched to an AR invoice.")
    if billed_in and billed_in[0]:
        print(f"AR loaded: ${billed_in[0]:,.0f} billed · ${(billed_in[1] or 0):,.0f} still open (GC owes you).")
    if show:
        print(f"\nSample (top {show} by amount):")
        for r in con.execute(
                "SELECT doc_number, project_no, status, amount, balance FROM billing_event "
                "WHERE source='qbo_invoice' AND project_no IS NOT NULL ORDER BY amount DESC LIMIT ?", (show,)):
            print(f"  #{(r[0] or '—'):<8} {(r[1] or '—'):<10} {r[2]:<5} "
                  f"${(r[3] or 0):>12,.0f}  open ${(r[4] or 0):>10,.0f}")


def run(db_path: Path, dry_run: bool, show: int) -> None:
    con = _connect(db_path)
    print("Authenticating to QBO (Touch ID)…")
    from shared.qbo_api import load_credentials, query_all  # lazy: pulls in requests
    access, company_id = load_credentials()
    print("  authenticated.")  # never echo the realm / company id (owner 2026-08-06)

    invoices = query_all(access, company_id, "Invoice")
    records = [r for r in (parse_invoice(i) for i in invoices) if r]
    matched = sum(1 for r in records if r["project_no"])
    print(f"Pulled {len(invoices)} invoices · {matched} matched to a project #")

    if not dry_run:
        now = dt.datetime.now().isoformat(timespec="seconds")
        n = write_events(con, records, now)
        print(f"Wrote {n} billing_event rows → {db_path}")
    else:
        print("--dry-run: nothing written.")
        # coverage against a temp copy so we don't touch the real ledger
        with tempfile.TemporaryDirectory() as d:
            tmp = _connect(Path(d) / "preview.sqlite3")
            for r in con.execute("SELECT line_uid, project_no, invoice_no, matched_invoice FROM ap_bill_line"):
                tmp.execute("INSERT INTO ap_bill_line (line_uid, project_no, invoice_no, matched_invoice, source, loaded_at) "
                            "VALUES (?,?,?,?,'copy','t')", tuple(r))
            write_events(tmp, records, "preview")
            _draw_coverage(tmp, show)
            tmp.close()
            con.close()
            return
    _draw_coverage(con, show)
    con.close()


def _selftest() -> None:
    """Offline: fabricate invoices, parse, write, and prove the draw join by Invoice #."""
    fake = [
        {"Id": "1001", "DocNumber": "34319", "TxnDate": "2026-05-15", "TotalAmt": 78032.0,
         "Balance": 0.0, "CustomerRef": {"name": "Multi Family:MFD325 Mesquite Briarwood", "value": "5"},
         "CustomerMemo": {"value": "Mesquite - Briarwood - May Draw 2026"}},
        {"Id": "1002", "DocNumber": "34535", "TxnDate": "2026-08-06", "TotalAmt": 91578.90,
         "Balance": 91578.90, "CustomerRef": {"name": "Residential:RP7470 315 Woolley St"},
         "PrivateNote": "315 Woolley - Aug Draw"},
        {"Id": "1003", "DocNumber": "", "TxnDate": "2026-01-01", "TotalAmt": 100.0, "Balance": 0.0,
         "CustomerRef": {"name": "Some Customer With No Project"}},
    ]
    records = [r for r in (parse_invoice(i) for i in fake) if r]
    with tempfile.TemporaryDirectory() as d:
        con = _connect(Path(d) / "selftest.sqlite3")
        con.execute("INSERT INTO ap_bill_line (line_uid, project_no, invoice_no, matched_invoice, "
                    "bill_total, source, loaded_at) VALUES "
                    "('t1','MFD325','34319','34319 — MFD325 - Mesquite - Briarwood - May Draw 2026',12000,'test','t')")
        con.commit()
        n = write_events(con, records, "selftest")
        assert n == 3, f"expected 3 rows, got {n}"

        paid = con.execute("SELECT amount, status, project_no, division FROM billing_event "
                           "WHERE doc_number='34319'").fetchone()
        assert paid and abs(paid[0] - 78032.0) < 0.01 and paid[1] == "Paid", paid
        assert paid[2] == "MFD325" and paid[3] == "Multi Family", paid

        openinv = con.execute("SELECT status, project_no, division, balance FROM billing_event "
                              "WHERE doc_number='34535'").fetchone()
        assert openinv == ("Open", "RP7470", "Residential", 91578.90), openinv

        joined = con.execute(
            "SELECT b.amount, b.status FROM ap_bill_line a "
            "JOIN billing_event b ON b.doc_number = a.invoice_no WHERE a.project_no='MFD325'").fetchone()
        assert joined and abs(joined[0] - 78032.0) < 0.01 and joined[1] == "Paid", joined
        con.close()
    print("selftest OK: 3 invoices parsed · project+division extracted · Paid/Open derived · "
          "draw ↔ invoice join by Invoice # works.")


def main() -> int:
    ap = argparse.ArgumentParser(description="Load QBO AR invoices (the draws the GC pays you) into the ledger.")
    ap.add_argument("--db", type=Path, default=DEFAULT_DB)
    ap.add_argument("--dry-run", action="store_true", help="Pull + coverage; write nothing.")
    ap.add_argument("--show", type=int, default=12, help="Sample rows to print.")
    ap.add_argument("--selftest", action="store_true", help="Offline pipeline proof (no QBO).")
    args = ap.parse_args()

    if args.selftest:
        _selftest()
        return 0
    run(args.db, args.dry_run, args.show)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
