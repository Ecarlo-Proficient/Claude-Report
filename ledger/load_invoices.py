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
                       └────────────── GAP FALLBACK (this, QBO) ──────────┘

The tracker is authoritative and covers the open + recent invoices. A handful of
OLDER draws were never entered there, so their draw shows no "billed (in)". The GAP
FALLBACK closes exactly those holes: after loading Notion, it finds CP/MFD draw
invoice #s with no billing_event match and pulls ONLY those from QBO by DocNumber
(source='qbo_fallback'). Tracker first; QBO fills only what it lacks (owner 2026-08-11).

SAFETY
    * Read-only on Notion AND QBO; writes only the local ledger. Scoped full-replace
      of each source ('invoice_tracker', 'qbo_fallback') — idempotent; a re-run mirrors.
    * The gap fallback authenticates to QBO → ONE Touch ID on this Mac. Skip with
      --no-qbo (Notion-only). --selftest stays fully offline (no Notion, no QBO).
    * --dry-run reads + reports draw coverage + the gap count WITHOUT writing or pulling QBO.

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
from shared import paths, lien_status as liens, notion_customers as customers  # noqa: E402

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
# Lien Tracker (Notion). Each invoice page carries a `Lien` relation back to this DB;
# we resolve it to the lien's Status for the Open Invoices tab. Read-only.
LIEN_DS = paths.get("ACB_LIEN_TRACKER_DS_ID") or "2c5b24f7-5585-80c5-b2d9-000bfcdaa084"
# Customer-list data sources the invoice `Customer` relation points at, so the client column
# shows the PARENT client (the GC), not the project-level `Customer (raw)` child. Res/Com and
# MFD keep separate lists; both are read (merged) so every invoice resolves. Override via env.
RESCOM_CUST_DS = paths.get("ACB_INVOICE_RESCOM_CUST_DS_ID") or "19db24f7-5585-81af-a4e1-000bbe22e6cc"
MFD_CUST_DS = paths.get("ACB_INVOICE_MFD_CUST_DS_ID") or "34bb24f7-5585-80d3-94fa-000b847f04e2"

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
    if t == "checkbox":
        return bool(p.get("checkbox"))
    if t == "date":
        return (p.get("date") or {}).get("start")
    if t == "formula":
        f = p.get("formula") or {}
        return f.get(f.get("type"))
    return None


def parse_invoice_page(page: dict, division_default: str | None = None,
                       lien_by_page: dict | None = None,
                       cust_cache: dict | None = None) -> dict | None:
    """One Invoice Tracker Notion page → a billing_event record."""
    props = page.get("properties", {})
    inv_id = _prop(props, "Invoice ID")
    doc = _prop(props, "Invoice #")
    if not (inv_id or doc):
        return None
    proj = ((_prop(props, "Project #") or "").upper().strip() or None) \
        or _extract_proj(_prop(props, "Customer (raw)")) or _extract_proj(_prop(props, "Memo"))
    lien_stat, lien_notice = liens.for_invoice(props, lien_by_page or {})   # shared resolver
    # The PARENT client (the GC), resolved from the `Customer` relation; the raw name is the
    # project-level child ("MFD177 - MERRITT PARK") and is only a fallback.
    client = customers.relation_title(props.get("Customer"), cust_cache or {}) \
        or _prop(props, "Customer (raw)")
    return {
        "qbo_txn_id": str(inv_id or doc),
        "doc_number": (str(doc).strip() if doc else None),
        "project_no": proj,
        "division": _division_for(proj) or division_default,
        "customer": client,
        "memo": _prop(props, "Memo"),
        "amount": _num(_prop(props, "Total Amount")),
        "balance": _num(_prop(props, "Open balance")),
        "txn_date": _prop(props, "Date"),
        "status": _prop(props, "Status"),          # Unpaid | Partially Paid | Paid
        "due_date": _prop(props, "Due Date"),
        "paid_date": _prop(props, "Paid Date"),
        "net_terms": _prop(props, "Net Terms"),
        "aging_bucket": _prop(props, "Aging Bucket"),
        "litigation": 1 if _prop(props, "Litigation") else 0,
        "lien_status": lien_stat,
        "lien_notice": lien_notice,
        "draw_period": None,
        "note": _prop(props, "Quick Status"),      # the collections one-liner ("GC paying Fri") for the AR view
        "source": "invoice_tracker",
    }


def load_lien_index(nc, ds_id: str) -> dict:
    """Resilient wrapper over `shared.lien_status.index_from_pages`: query the Lien Tracker
    once → {lien_page_id: {status, notice}}. The lien column is an add-on, so any read error
    (or the DB not shared with the integration) degrades to an empty index and the core
    invoice load still runs - the column just shows blank."""
    try:
        return liens.index_from_pages(nc.query_data_source(ds_id))
    except Exception as e:  # noqa: BLE001 - never let the lien add-on break the AR load
        print(f"  Lien Tracker: not readable ({type(e).__name__}) - lien column will be blank. "
              "Share the Lien Tracker DB with the automation integration to enable it.")
        return {}


def load_customer_titles(nc) -> dict:
    """Merged {customer_page_id -> parent name} from the Res/Com + MFD customer lists the
    invoice `Customer` relations point at (via `shared.notion_customers`). Resilient: a list
    the token can't read is skipped and those invoices fall back to `Customer (raw)`, so the
    load never breaks over the client-name nicety."""
    cache: dict = {}
    for label, ds in (("Res/Com", RESCOM_CUST_DS), ("MFD", MFD_CUST_DS)):
        if not ds:
            continue
        try:
            cache.update(customers.build_title_cache(nc.query_data_source(ds)))
        except Exception as e:  # noqa: BLE001 - client name is a nicety, never break the load
            print(f"  Customer list ({label}): not readable ({type(e).__name__}); "
                  "those invoices show the project name until the DB is shared with the integration.")
    return cache


_COLS = ["qbo_txn_id", "doc_number", "project_no", "division", "customer", "memo",
         "amount", "balance", "txn_date", "status", "due_date", "paid_date", "net_terms",
         "aging_bucket", "litigation", "lien_status", "lien_notice",
         "draw_period", "note", "source", "loaded_at"]


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
        cols = {r[1] for r in con.execute("PRAGMA table_info(billing_event)")}
    # Additive AR/lien columns - ALTER onto DBs that already have the doc_number shape
    # (CREATE TABLE IF NOT EXISTS won't add columns to an existing table).
    add = [("due_date", "TEXT"), ("net_terms", "TEXT"), ("aging_bucket", "TEXT"),
           ("litigation", "INTEGER NOT NULL DEFAULT 0"), ("lien_status", "TEXT"),
           ("lien_notice", "TEXT"), ("paid_date", "TEXT"), ("note", "TEXT")]
    for name, decl in add:
        if cols and name not in cols:
            con.execute(f"ALTER TABLE billing_event ADD COLUMN {name} {decl}")
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


def _invoice_from_qbo(inv: dict) -> dict:
    """One raw QBO Invoice → a billing_event record (source='qbo_fallback')."""
    total = _num(inv.get("TotalAmt"))
    bal = _num(inv.get("Balance"))
    if bal is None:
        status = None
    elif bal <= 0.005:
        status = "Paid"
    elif total and bal < total - 0.005:
        status = "Partially Paid"
    else:
        status = "Unpaid"
    cust = (inv.get("CustomerRef") or {}).get("name")
    memo = inv.get("PrivateNote") or (inv.get("CustomerMemo") or {}).get("value")
    proj = _extract_proj(cust) or _extract_proj(memo)
    doc = inv.get("DocNumber")
    return {
        "qbo_txn_id": str(inv.get("Id") or doc),
        "doc_number": (str(doc).strip() if doc else None),
        "project_no": proj,
        "division": _division_for(proj),
        "customer": cust,
        "memo": memo,
        "amount": total,
        "balance": bal,
        "txn_date": inv.get("TxnDate"),
        "status": status,
        "due_date": inv.get("DueDate"),
        "paid_date": None,
        "net_terms": None,
        "aging_bucket": None,
        "litigation": 0,
        "lien_status": None,
        "lien_notice": None,
        "draw_period": None,
        "source": "qbo_fallback",
    }


def fill_gaps_from_qbo(con, now: str, dry_run: bool, batch: int = 25) -> int:
    """Fill the ONE thing the Invoice Tracker can't: draws whose AR invoice was never
    entered there. Find CP/MFD draw invoice_nos with no billing_event match, pull ONLY
    those from QBO by DocNumber, and land them as source='qbo_fallback'. The tracker
    stays authoritative (loaded first); QBO fills only the holes. Returns rows filled."""
    gaps = [r[0] for r in con.execute(
        "SELECT DISTINCT a.invoice_no FROM ap_bill_line a "
        "LEFT JOIN billing_event b ON b.doc_number = a.invoice_no "
        "WHERE a.invoice_no GLOB '[0-9]*' AND b.doc_number IS NULL "
        "  AND COALESCE(a.project_no,'') NOT LIKE 'RP%' AND a.matched_invoice NOT LIKE '%— RP%'")]
    if not gaps:
        print("QBO fallback: no gaps — every CP/MFD draw is covered by the Invoice Tracker.")
        return 0
    print(f"QBO fallback: {len(gaps)} draw invoices missing from the tracker.")
    if dry_run:
        print("  --dry-run: not pulling QBO.")
        return 0
    from shared.qbo_api import load_credentials, query_all
    access, company_id = load_credentials()
    print("  authenticated.")                     # never echo the realm/company id (owner 2026-08-06)
    recs: list[dict] = []
    for i in range(0, len(gaps), batch):
        inlist = ", ".join("'" + d.replace("'", "") + "'" for d in gaps[i:i + batch])
        for inv in query_all(access, company_id, "Invoice", f"DocNumber IN ({inlist})"):
            recs.append(_invoice_from_qbo(inv))
    con.execute("DELETE FROM billing_event WHERE source='qbo_fallback'")
    ph = ", ".join(f":{c}" for c in _COLS)
    for r in recs:
        row = {**r, "loaded_at": now}
        con.execute(f"INSERT OR REPLACE INTO billing_event ({', '.join(_COLS)}) VALUES ({ph})",
                    {c: row.get(c) for c in _COLS})
    con.commit()
    got = {r["doc_number"] for r in recs if r["doc_number"]}
    still = [g for g in gaps if g not in got]
    print(f"  filled {len(got)} of {len(gaps)} gap invoices from QBO"
          + (f" · {len(still)} still unmatched (voided/deleted in QBO?): {', '.join(still[:8])}" if still else ""))
    return len(got)


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


def run(db_path: Path, dry_run: bool, show: int, no_qbo: bool = False) -> None:
    con = _connect(db_path)
    from shared.notion_client import NotionClient
    nc = NotionClient()
    lien_by_page = load_lien_index(nc, LIEN_DS) if LIEN_DS else {}
    if lien_by_page:
        print(f"  Lien Tracker: {len(lien_by_page)} lien rows indexed")
    cust_cache = load_customer_titles(nc)
    if cust_cache:
        print(f"  Customer lists: {len(cust_cache)} clients indexed (parent-name resolution)")
    records: list[dict] = []
    for label, ds_id, div_default in (("Res/Com", RESCOM_DS, None), ("MFD", MFD_DS, "Multi Family")):
        if not ds_id:
            print(f"  skip {label}: no data-source id"); continue
        n0 = len(records)
        for page in nc.query_data_source(ds_id):
            rec = parse_invoice_page(page, div_default, lien_by_page, cust_cache)
            if rec:
                records.append(rec)
        print(f"  Invoice Tracker ({label}): {len(records) - n0} invoices")
    liened = sum(1 for r in records if r["lien_status"])
    if liened:
        print(f"  {liened} open/closed invoices carry a lien status")
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
            if not no_qbo:
                fill_gaps_from_qbo(tmp, "preview", dry_run=True)   # reports the gap count only
            _draw_coverage(tmp, show)
            tmp.close()
        con.close()
        return

    now = dt.datetime.now().isoformat(timespec="seconds")
    n = write_events(con, records, now)
    print(f"Wrote {n} billing_event rows → {db_path}")
    if not no_qbo:
        fill_gaps_from_qbo(con, now, dry_run=False)               # QBO fills the tracker's holes
    _draw_coverage(con, show)
    con.close()


def _selftest() -> None:
    """Offline: fabricate Invoice Tracker Notion pages, parse, write, prove the draw join."""
    def page(inv_id, num, proj, total, bal, status, date, memo, cust=None,
             due=None, terms=None, bucket=None, litig=False, lien_ids=None, cust_rel=None):
        props = {
            "Invoice ID": {"type": "rich_text", "rich_text": [{"plain_text": inv_id}]},
            "Invoice #": {"type": "title", "title": [{"plain_text": num}]},
            "Project #": {"type": "rich_text", "rich_text": [{"plain_text": proj}]},
            "Customer (raw)": {"type": "rich_text", "rich_text": [{"plain_text": cust or ""}]},
            "Customer": {"type": "relation", "relation": ([{"id": cust_rel}] if cust_rel else [])},
            "Total Amount": {"type": "number", "number": total},
            "Open balance": {"type": "number", "number": bal},
            "Status": {"type": "select", "select": {"name": status}},
            "Date": {"type": "date", "date": {"start": date}},
            "Due Date": {"type": "date", "date": ({"start": due} if due else None)},
            "Net Terms": {"type": "select", "select": ({"name": terms} if terms else None)},
            "Aging Bucket": {"type": "select", "select": ({"name": bucket} if bucket else None)},
            "Litigation": {"type": "checkbox", "checkbox": litig},
            "Lien": {"type": "relation", "relation": [{"id": i} for i in (lien_ids or [])]},
            "Memo": {"type": "rich_text", "rich_text": [{"plain_text": memo}]},
        }
        return {"properties": props}

    # A fabricated Lien Tracker index: two liens, different statuses, on one invoice.
    lien_by_page = {
        "L1": {"status": "Mailed", "notice": "RP Notice"},
        "L2": {"status": "Lien", "notice": "Affidavit of Lien Claimed"},
    }
    # A fabricated customer title cache (the `Customer` relation resolves to the PARENT client).
    cust_cache = {"CUST_FIRE": "Firestone Building Co"}
    pages = [
        page("5001", "34319", "MFD325", 78032.0, 0.0, "Paid", "2026-05-15", "Mesquite - Briarwood - May Draw 2026"),
        page("5002", "34535", "RP7470", 91578.90, 91578.90, "Unpaid", "2026-08-06", "315 Woolley - Aug Draw",
             due="2026-09-05", terms="Net 30", bucket="Current", lien_ids=["L1", "L2"],
             cust="RP7470 - 315 WOOLLEY", cust_rel="CUST_FIRE"),
        page("5003", "34600", "MFD192", 200000.0, 50000.0, "Partially Paid", "2026-07-01", "Mayhill - July Draw",
             cust="MFD192 - MAYHILL"),
    ]
    records = [r for r in (parse_invoice_page(p, None, lien_by_page, cust_cache) for p in pages) if r]
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

        ar = con.execute("SELECT due_date, net_terms, lien_status, lien_notice, litigation "
                         "FROM billing_event WHERE doc_number='34535'").fetchone()
        # Two related liens (Mailed + Lien) → most-escalated wins.
        assert ar == ("2026-09-05", "Net 30", "Lien", "Affidavit of Lien Claimed", 0), ar
        assert con.execute("SELECT lien_status FROM billing_event WHERE doc_number='34319'").fetchone()[0] is None

        # Client = the PARENT (from the Customer relation), NOT the "Customer (raw)" project child.
        assert con.execute("SELECT customer FROM billing_event WHERE doc_number='34535'").fetchone()[0] == "Firestone Building Co"
        # No relation → falls back to the raw project-level name.
        assert con.execute("SELECT customer FROM billing_event WHERE doc_number='34600'").fetchone()[0] == "MFD192 - MAYHILL"
        con.close()
    print("selftest OK: 3 Invoice Tracker pages parsed · project+division · Paid/Partially/Unpaid · "
          "draw ↔ invoice join by Invoice # · due date + escalated lien · parent-client resolve.")


def main() -> int:
    ap = argparse.ArgumentParser(description="Load AR invoices from the Invoice Tracker (Notion) into the ledger.")
    ap.add_argument("--db", type=Path, default=DEFAULT_DB)
    ap.add_argument("--dry-run", action="store_true", help="Read + coverage; write nothing.")
    ap.add_argument("--show", type=int, default=12, help="Sample rows to print.")
    ap.add_argument("--no-qbo", action="store_true",
                    help="Skip the QBO gap-fallback (Notion-only; no Touch ID).")
    ap.add_argument("--selftest", action="store_true", help="Offline pipeline proof (no Notion).")
    args = ap.parse_args()

    if args.selftest:
        _selftest()
        return 0
    run(args.db, args.dry_run, args.show, args.no_qbo)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
