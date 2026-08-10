#!/usr/bin/env python3
"""
load_invoices.py — land AR invoices (the draws the GC pays YOU) into the ledger,
read FROM the Invoice Tracker (Notion) — NOT a second QBO pull.

Systems connect, they don't each re-pull QBO. `invoice-sync` already mirrors every
QBO invoice into the **Invoice Tracker** Notion DBs (Res/Com + MFD) every run, and
keeps paid invoices on file for 12 months. So the ledger reads THAT (via the shared
Notion token — the same way `load_customers.py` reads the Customer List), and writes
one `billing_event` per invoice keyed by **Invoice #** — the same number
`ap_bill_line.invoice_no` carries, so the Draws view can show billed-to-GC (money IN)
next to paid-to-vendors (money OUT) on every draw.

    QBO ──(invoice-sync)──▶ Invoice Tracker (Notion) ──(this)──▶ ledger.billing_event

SAFETY
    * Read-only on Notion; writes only the local ledger. Scoped full-replace of
      source='invoice_tracker' billing_event (idempotent; a re-run mirrors Notion).
    * No QBO, no Touch ID — just the Notion token.
    * --dry-run reads + reports draw coverage WITHOUT writing.
    * --selftest runs the whole pipeline offline on a throwaway DB (no Notion).

USAGE
    python3 ledger/load_invoices.py --selftest     # offline proof, no Notion
    python3 ledger/load_invoices.py                # read Invoice Tracker + write
    python3 ledger/load_invoices.py --dry-run      # read + coverage, write nothing
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

# Invoice Tracker data-source ids (public — documented in docs/Invoice Tracker —
# System Reference.md). Overridable via machine.env if they ever change.
RESCOM_DS = paths.get("ACB_INVOICE_RESCOM_DS_ID") or "265b24f7-5585-803c-bcae-000ba27328cd"
MFD_DS = paths.get("ACB_INVOICE_MFD_DS_ID") or "0f8e7cdf-16fe-4137-82e6-255e2ff400ce"

_PROJ_RE = re.compile(r"(MFD|CP|RP)\s*(\d+)(-FTW)?", re.IGNORECASE)
_DIV_PREFIX = {"MFD": "Multi Family", "CP": "Commercial", "RP": "Residential"}


def _extract_proj(text) -> str | None:
    if not text:
        return None
    m = _PROJ_RE.search(str(text))
    return (m.group(1) + m.group(2) + (m.group(3) or "")).upper() if m else None


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


def _prop(props: dict, name: str):
    """Value of a Notion property by name, whatever its type."""
    p = props.get(name) or {}
    t = p.get("type")
    if t in ("title", "rich_text"):
        return ("".join(x.get("plain_text", "") for x in (p.get(t) or [])).strip() or None)
    if t == "number":
        return p.get("number")
    if t == "select":
        return (p.get("select") or {}).get("name")
    if t == "multi_select":
        names = [o.get("name") for o in (p.get("multi_select") or [])]
        return names[0] if names else None
    if t == "date":
        return (p.get("date") or {}).get("start")
    if t == "formula":
        f = p.get("formula") or {}
        return f.get(f.get("type"))
    return None


def parse_invoice_page(page: dict, division_default: str | None = None) -> dict | None:
    """One Invoice Tracker Notion page → a billing_event record."""
    props = page.get("properties", {})
    inv_id = _prop(props, "Invoice ID")
    doc = _prop(props, "Invoice #")
    if not (inv_id or doc):
        return None
    proj = ((_prop(props, "Project #") or "").upper().strip() or None) \
        or _extract_proj(_prop(props, "Customer (raw)")) or _extract_proj(_prop(props, "Memo"))
    return {
        "qbo_txn_id": str(inv_id or doc),
        "doc_number": (str(doc).strip() if doc else None),
        "project_no": proj,
        "division": _division_for(proj) or division_default,
        "customer": _prop(props, "Customer (raw)"),
        "memo": _prop(props, "Memo"),
        "amount": _num(_prop(props, "Total Amount")),
        "balance": _num(_prop(props, "Open balance")),
        "txn_date": _prop(props, "Date"),
        "status": _prop(props, "Status"),          # Unpaid | Partially Paid | Paid
        "draw_period": None,
        "source": "invoice_tracker",
    }


_COLS = ["qbo_txn_id", "doc_number", "project_no", "division", "customer", "memo",
         "amount", "balance", "txn_date", "status", "draw_period", "source", "loaded_at"]


def _connect(db_path: Path) -> sqlite3.Connection:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(db_path)
    con.executescript(SCHEMA_SQL.read_text(encoding="utf-8"))
    _migrate_billing_event(con)
    return con


def _migrate_billing_event(con) -> None:
    """billing_event predates the AR columns in DBs made by earlier schema. It's a
    placeholder (empty until this loader runs), so a shape upgrade is a safe drop +
    recreate. Refuse if it somehow already holds rows."""
    cols = {r[1] for r in con.execute("PRAGMA table_info(billing_event)")}
    if cols and "doc_number" not in cols:
        if con.execute("SELECT COUNT(*) FROM billing_event").fetchone()[0]:
            sys.exit("billing_event has the legacy shape AND rows — refusing to auto-migrate.")
        con.execute("DROP TABLE billing_event")
        con.executescript(SCHEMA_SQL.read_text(encoding="utf-8"))
        con.commit()


def write_events(con, records: list[dict], now: str) -> int:
    con.execute("DELETE FROM billing_event WHERE source='invoice_tracker'")
    ph = ", ".join(f":{c}" for c in _COLS)
    for r in records:
        row = {**r, "loaded_at": now}
        con.execute(f"INSERT OR REPLACE INTO billing_event ({', '.join(_COLS)}) VALUES ({ph})",
                    {c: row.get(c) for c in _COLS})
    con.commit()
    return len(records)


def _draw_coverage(con, show: int) -> None:
    row = con.execute(
        "SELECT COUNT(DISTINCT a.invoice_no) draws, COUNT(DISTINCT b.doc_number) matched "
        "FROM ap_bill_line a LEFT JOIN billing_event b ON b.doc_number = a.invoice_no "
        "WHERE a.invoice_no IS NOT NULL AND a.invoice_no <> '' "
        "  AND COALESCE(a.project_no,'') NOT LIKE 'RP%' AND a.matched_invoice NOT LIKE '%— RP%'"
    ).fetchone()
    tot = con.execute("SELECT SUM(amount), SUM(balance) FROM billing_event WHERE source='invoice_tracker'").fetchone()
    print(f"\nDraw coverage: {row[1] or 0} of {row[0] or 0} MFD/CP draws matched to an AR invoice.")
    if tot and tot[0]:
        print(f"AR loaded: ${tot[0]:,.0f} billed · ${(tot[1] or 0):,.0f} still open (GC owes you).")
    if show:
        for r in con.execute(
                "SELECT doc_number, project_no, status, amount, balance FROM billing_event "
                "WHERE source='invoice_tracker' AND project_no IS NOT NULL ORDER BY amount DESC LIMIT ?", (show,)):
            print(f"  #{(r[0] or '—'):<8} {(r[1] or '—'):<10} {(r[2] or '—'):<14} "
                  f"${(r[3] or 0):>12,.0f}  open ${(r[4] or 0):>10,.0f}")


def run(db_path: Path, dry_run: bool, show: int) -> None:
    con = _connect(db_path)
    from shared.notion_client import NotionClient
    nc = NotionClient()
    records: list[dict] = []
    for label, ds_id, div_default in (("Res/Com", RESCOM_DS, None), ("MFD", MFD_DS, "Multi Family")):
        if not ds_id:
            print(f"  skip {label}: no data-source id"); continue
        n0 = len(records)
        for page in nc.query_data_source(ds_id):
            rec = parse_invoice_page(page, div_default)
            if rec:
                records.append(rec)
        print(f"  Invoice Tracker ({label}): {len(records) - n0} invoices")
    matched = sum(1 for r in records if r["project_no"])
    print(f"Read {len(records)} invoices · {matched} with a project #")

    if dry_run:
        print("--dry-run: nothing written.")
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

    now = dt.datetime.now().isoformat(timespec="seconds")
    n = write_events(con, records, now)
    print(f"Wrote {n} billing_event rows → {db_path}")
    _draw_coverage(con, show)
    con.close()


def _selftest() -> None:
    """Offline: fabricate Invoice Tracker Notion pages, parse, write, prove the draw join."""
    def page(inv_id, num, proj, total, bal, status, date, memo, cust=None):
        props = {
            "Invoice ID": {"type": "rich_text", "rich_text": [{"plain_text": inv_id}]},
            "Invoice #": {"type": "title", "title": [{"plain_text": num}]},
            "Project #": {"type": "rich_text", "rich_text": [{"plain_text": proj}]},
            "Customer (raw)": {"type": "rich_text", "rich_text": [{"plain_text": cust or ""}]},
            "Total Amount": {"type": "number", "number": total},
            "Open balance": {"type": "number", "number": bal},
            "Status": {"type": "select", "select": {"name": status}},
            "Date": {"type": "date", "date": {"start": date}},
            "Memo": {"type": "rich_text", "rich_text": [{"plain_text": memo}]},
        }
        return {"properties": props}

    pages = [
        page("5001", "34319", "MFD325", 78032.0, 0.0, "Paid", "2026-05-15", "Mesquite - Briarwood - May Draw 2026"),
        page("5002", "34535", "RP7470", 91578.90, 91578.90, "Unpaid", "2026-08-06", "315 Woolley - Aug Draw"),
        page("5003", "34600", "MFD192", 200000.0, 50000.0, "Partially Paid", "2026-07-01", "Mayhill - July Draw"),
    ]
    records = [r for r in (parse_invoice_page(p, None) for p in pages) if r]
    with tempfile.TemporaryDirectory() as d:
        con = _connect(Path(d) / "selftest.sqlite3")
        con.execute("INSERT INTO ap_bill_line (line_uid, project_no, invoice_no, matched_invoice, "
                    "bill_total, source, loaded_at) VALUES "
                    "('t1','MFD325','34319','34319 — MFD325 - Mesquite - Briarwood - May Draw 2026',12000,'test','t')")
        con.commit()
        n = write_events(con, records, "selftest")
        assert n == 3, n

        paid = con.execute("SELECT amount, status, project_no, division FROM billing_event WHERE doc_number='34319'").fetchone()
        assert paid == (78032.0, "Paid", "MFD325", "Multi Family"), paid

        part = con.execute("SELECT status, balance, division FROM billing_event WHERE doc_number='34600'").fetchone()
        assert part == ("Partially Paid", 50000.0, "Multi Family"), part

        joined = con.execute("SELECT b.amount, b.status FROM ap_bill_line a "
                             "JOIN billing_event b ON b.doc_number = a.invoice_no WHERE a.project_no='MFD325'").fetchone()
        assert joined == (78032.0, "Paid"), joined
        con.close()
    print("selftest OK: 3 Invoice Tracker pages parsed · project+division · Paid/Partially/Unpaid · "
          "draw ↔ invoice join by Invoice # works.")


def main() -> int:
    ap = argparse.ArgumentParser(description="Load AR invoices from the Invoice Tracker (Notion) into the ledger.")
    ap.add_argument("--db", type=Path, default=DEFAULT_DB)
    ap.add_argument("--dry-run", action="store_true", help="Read + coverage; write nothing.")
    ap.add_argument("--show", type=int, default=12, help="Sample rows to print.")
    ap.add_argument("--selftest", action="store_true", help="Offline pipeline proof (no Notion).")
    args = ap.parse_args()

    if args.selftest:
        _selftest()
        return 0
    run(args.db, args.dry_run, args.show)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
