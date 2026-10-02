#!/usr/bin/env python3
"""
load_payments.py - land received customer PAYMENTS (money IN, as transactions)
into the ledger, straight from QBO.

A payment is the moment the GC hands over cash. ONE payment can settle several
invoices, so the ledger keeps two row-shapes:

    payment ................ the transaction: total, date, who paid
    payment_application .... one row per invoice that payment landed on (+ amount)

That is exactly what the dashboard's Payments tab reads: the payment on top, the
invoice(s) it paid grouped beneath it.

    QBO Payment ──(this)──▶ ledger.payment  +  payment_application
                            └── invoice #, project resolved via billing_event ──┘

Why a QBO pull (not the tracker): billing_event is invoice-level (amount/balance).
A single cheque that pays three draws is only a real thing in QBO's Payment object
(Line[].LinkedTxn). Reconstructing it from invoices would be guessing, so we read
the actual transaction.

SAFETY
    * Read-only on QBO (GET only); writes only the local ledger.
    * Full replace over a rolling window each run (source='qbo_payment') - a re-run
      mirrors, never double-counts.
    * Authenticates to QBO → ONE Touch ID on this Mac. --selftest stays fully offline.
    * --dry-run pulls + reports WITHOUT writing.

USAGE
    python3 ledger/load_payments.py --selftest      # offline proof, no QBO
    python3 ledger/load_payments.py                 # pull last 12 months + write
    python3 ledger/load_payments.py --months 24     # wider window
    python3 ledger/load_payments.py --dry-run       # pull + report, write nothing
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

_PAYMENT_COLS = ("qbo_txn_id", "txn_date", "customer", "customer_id",
                 "parent_customer", "parent_customer_id",
                 "total_amt", "unapplied_amt", "method", "ref_no", "source", "loaded_at")
_APP_COLS = ("payment_txn_id", "invoice_txn_id", "invoice_no",
             "project_no", "division", "amount", "invoice_open")


def _num(x) -> float:
    try:
        return round(float(x), 2)
    except (TypeError, ValueError):
        return 0.0


def _division_of(proj: str | None) -> str | None:
    """Project # prefix → QBO division (spelled out, as QBO stores classes)."""
    if not proj:
        return None
    u = proj.upper()
    if u.startswith("MFD"):
        return "Multi Family"
    if u.startswith("CP"):
        return "Commercial"
    if u.startswith("RP"):
        return "Residential"
    return None


def _connect(db_path: Path) -> sqlite3.Connection:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(db_path)
    con.row_factory = sqlite3.Row
    con.executescript(SCHEMA_SQL.read_text(encoding="utf-8"))   # CREATE TABLE IF NOT EXISTS → new tables land
    _migrate_payment(con)
    return con


def _migrate_payment(con) -> None:
    """Add columns to payment / payment_application tables made by an earlier schema.
    (IF NOT EXISTS won't alter an existing table, so add them explicitly.)"""
    pcols = {r[1] for r in con.execute("PRAGMA table_info(payment)")}
    for name in ("parent_customer", "parent_customer_id"):
        if pcols and name not in pcols:
            con.execute(f"ALTER TABLE payment ADD COLUMN {name} TEXT")
    acols = {r[1] for r in con.execute("PRAGMA table_info(payment_application)")}
    if acols and "invoice_open" not in acols:
        con.execute("ALTER TABLE payment_application ADD COLUMN invoice_open NUMERIC")


def _cutoff(months: int) -> str:
    """ISO date `months` whole months back from today (day clamped)."""
    today = dt.date.today()
    y, m = today.year, today.month - months
    while m <= 0:
        m += 12
        y -= 1
    day = min(today.day, 28)
    return dt.date(y, m, day).isoformat()


def _customer_gc_map(query_all, access, company_id, extract_proj=None) -> tuple:
    """One customer pull → two maps.
      gc_map   {customer_id → (gc_name, gc_id)}: every QBO customer resolved to the TOP of its
               hierarchy - the GC. A Payment's CustomerRef is the bare leaf (often a project
               sub-customer 'RP6676-FTW'); its parent chain ends at the client (LONESTAR GREEN HOMES).
      proj_map {project_no → (gc_name, gc_id)}: because every project is a QBO sub-customer whose
               NAME carries the project # (Customer:Project), reversing that gives the client for
               EVERY project - not just the ones that happen to have a payment. Empty if no
               extract_proj is passed."""
    cust = {}
    for c in query_all(access, company_id, "Customer"):
        cid = str(c.get("Id") or "")
        if cid:
            cust[cid] = (c.get("DisplayName") or "", str((c.get("ParentRef") or {}).get("value") or ""))

    def root(i: str) -> str:
        # climb to the top parent, but stop at the deepest KNOWN ancestor: an inactive
        # parent won't be in this (active-only) pull, so never step onto a missing id.
        seen = set()
        while cust.get(i, ("", ""))[1] in cust and cust[i][1] and i not in seen:
            seen.add(i)
            i = cust[i][1]
        return i

    gc_map = {i: (cust[root(i)][0], root(i)) for i in cust}
    proj_map: dict = {}
    if extract_proj:
        for cid, (disp, _) in cust.items():
            pn = extract_proj(disp or "")
            if not pn:
                continue
            r = root(cid)
            gc_name = cust[r][0]
            # only when there's a REAL client above the project (its "GC" isn't itself another project name)
            if gc_name and not extract_proj(gc_name):
                proj_map.setdefault(pn.upper(), (gc_name, r))
    return gc_map, proj_map


def _payment_method_map(query_all, access, company_id) -> dict:
    """{payment_method_id → name} (Check / ACH / Cash / ...). A Payment's
    PaymentMethodRef comes back as an id only (no name), so resolve it once."""
    out = {}
    for m in query_all(access, company_id, "PaymentMethod"):
        mid = str(m.get("Id") or "")
        if mid:
            out[mid] = m.get("Name")
    return out


def _invoice_index(con) -> dict:
    """{QBO invoice id → (invoice_no, project_no, division, open_balance)} from
    billing_event, so a payment's LinkedTxn resolves to the draw it paid."""
    idx = {}
    try:
        for r in con.execute("SELECT qbo_txn_id, doc_number, project_no, division, balance FROM billing_event"):
            if r["qbo_txn_id"]:
                idx[str(r["qbo_txn_id"])] = (r["doc_number"], r["project_no"], r["division"], r["balance"])
    except sqlite3.OperationalError:
        pass
    return idx


def _rows_from_payment(p: dict, inv_idx: dict) -> tuple[dict, list[dict]]:
    """One QBO Payment → (payment row, [application rows]). The line Amount is the
    money applied; when a line links to >1 invoice (rare) it is split evenly."""
    pid = str(p.get("Id") or "")
    cref = p.get("CustomerRef") or {}
    payment = {
        "qbo_txn_id": pid,
        "txn_date": (p.get("TxnDate") or "")[:10] or None,
        "customer": cref.get("name"),
        "customer_id": cref.get("value"),
        "parent_customer": None,        # filled from the QBO customer hierarchy in _load
        "parent_customer_id": None,
        "total_amt": _num(p.get("TotalAmt")),
        "unapplied_amt": _num(p.get("UnappliedAmt")),
        "method": (p.get("PaymentMethodRef") or {}).get("name"),
        "ref_no": p.get("PaymentRefNum"),
    }
    apps: list[dict] = []
    for line in p.get("Line") or []:
        lamt = _num(line.get("Amount"))
        inv_links = [lt for lt in (line.get("LinkedTxn") or []) if lt.get("TxnType") == "Invoice"]
        if not inv_links:
            continue
        share = round(lamt / len(inv_links), 2) if inv_links else lamt
        for lt in inv_links:
            inv_id = str(lt.get("TxnId") or "")
            if not inv_id:
                continue
            doc, proj, div, bal = inv_idx.get(inv_id, (None, None, None, None))
            apps.append({
                "payment_txn_id": pid,
                "invoice_txn_id": inv_id,
                "invoice_no": doc,
                "project_no": proj,
                "division": div,
                "amount": share,
                "invoice_open": bal,
            })
    return payment, apps


def _load(con, months: int, dry_run: bool) -> int:
    inv_idx = _invoice_index(con)
    cutoff = _cutoff(months)
    from shared.qbo_api import load_credentials, query_all, extract_proj
    access, company_id = load_credentials()
    print("  authenticated.")                       # never echo the realm/company id (owner 2026-08-06)
    gc_map, proj_cust = _customer_gc_map(query_all, access, company_id, extract_proj)
    print(f"  mapped {len(gc_map)} customers to their GC; {len(proj_cust)} projects to a client.")
    pm_map = _payment_method_map(query_all, access, company_id)     # PaymentMethodRef is an id only - resolve the name
    raw = query_all(access, company_id, "Payment", f"TxnDate >= '{cutoff}'")
    print(f"  pulled {len(raw)} payments since {cutoff}.")

    payments, apps = [], []
    for p in raw:
        prow, arows = _rows_from_payment(p, inv_idx)
        if not prow["qbo_txn_id"]:
            continue
        gc = gc_map.get(prow["customer_id"] or "")     # normalize the payer to the GC (parent of the project sub-customer)
        if gc and gc[0]:
            prow["parent_customer"], prow["parent_customer_id"] = gc
        mid = str((p.get("PaymentMethodRef") or {}).get("value") or "")
        if mid and pm_map.get(mid):
            prow["method"] = pm_map[mid]
        payments.append(prow)
        apps.extend(arows)

    # billing_event only holds open + recently-paid draws (~351), so most links to
    # older paid invoices don't resolve there. Pull just those invoices from QBO by
    # Id to fill invoice # + project, so the grouped view shows real numbers, not ids.
    unresolved = sorted({a["invoice_txn_id"] for a in apps if not a["invoice_no"] and a["invoice_txn_id"]})
    if unresolved:
        from shared.qbo_api import extract_proj
        extra: dict[str, tuple] = {}
        B = 80
        for i in range(0, len(unresolved), B):
            inlist = ", ".join("'" + x.replace("'", "") + "'" for x in unresolved[i:i + B])
            for inv in query_all(access, company_id, "Invoice", f"Id IN ({inlist})"):
                iid = str(inv.get("Id") or "")
                proj = extract_proj((inv.get("CustomerRef") or {}).get("name") or "")
                extra[iid] = (inv.get("DocNumber"), proj, _division_of(proj), _num(inv.get("Balance")))
        for a in apps:
            if not a["invoice_no"] and a["invoice_txn_id"] in extra:
                a["invoice_no"], a["project_no"], a["division"], a["invoice_open"] = extra[a["invoice_txn_id"]]
        print(f"  resolved {sum(1 for v in extra.values() if v[0])} more invoices from QBO by id.")

    total = round(sum(p["total_amt"] for p in payments), 2)
    resolved = sum(1 for a in apps if a["invoice_no"])
    print(f"  {len(payments)} payments · {total:,.2f} received · {len(apps)} invoice links "
          f"({resolved} resolved to a draw, {len(apps) - resolved} by id only).")
    if dry_run:
        print("  --dry-run: nothing written.")
        return len(payments)

    now = dt.datetime.now().isoformat(timespec="seconds")
    con.execute("DELETE FROM payment_application")
    con.execute("DELETE FROM payment")
    pph = ", ".join(f":{c}" for c in _PAYMENT_COLS)
    for p in payments:
        row = {**p, "source": "qbo_payment", "loaded_at": now}
        con.execute(f"INSERT OR REPLACE INTO payment ({', '.join(_PAYMENT_COLS)}) VALUES ({pph})",
                    {c: row.get(c) for c in _PAYMENT_COLS})
    aph = ", ".join(f":{c}" for c in _APP_COLS)
    seen = set()
    for a in apps:
        k = (a["payment_txn_id"], a["invoice_txn_id"])
        if k in seen:                                # PK is (payment, invoice); collapse dup links
            continue
        seen.add(k)
        con.execute(f"INSERT OR REPLACE INTO payment_application ({', '.join(_APP_COLS)}) VALUES ({aph})",
                    {c: a.get(c) for c in _APP_COLS})
    # project -> client (the GC), reversed from the QBO Customer:Project hierarchy, for EVERY project.
    con.execute("DELETE FROM project_customer")
    con.executemany(
        "INSERT OR REPLACE INTO project_customer (project_no, client, client_id, loaded_at) VALUES (?,?,?,?)",
        [(pn, gc, gid, now) for pn, (gc, gid) in proj_cust.items()])
    con.commit()
    print(f"  wrote {len(payments)} payments + {len(seen)} applications + {len(proj_cust)} project→client rows.")
    return len(payments)


def _selftest() -> int:
    """Offline proof: schema applies, a synthetic payment + its two invoice links
    round-trip, and the grouped read (payment → applications) reconstructs."""
    with tempfile.TemporaryDirectory() as td:
        con = _connect(Path(td) / "t.sqlite3")
        # a draw the payment will resolve against (balance 5000 = still partly open)
        con.execute("INSERT INTO billing_event (qbo_txn_id, doc_number, project_no, division, balance, source, loaded_at) "
                    "VALUES ('INV9','34999','CP861','Commercial',5000,'qbo_invoice','x')")
        pay = {"Id": "P1", "TxnDate": "2026-08-10", "TotalAmt": "150000.00", "UnappliedAmt": "0",
               "CustomerRef": {"name": "Firestone Building Co", "value": "77"},
               "PaymentMethodRef": {"name": "Check"}, "PaymentRefNum": "10231",
               "Line": [
                   {"Amount": "100000.00", "LinkedTxn": [{"TxnId": "INV9", "TxnType": "Invoice"}]},
                   {"Amount": "50000.00", "LinkedTxn": [{"TxnId": "INV_UNKNOWN", "TxnType": "Invoice"}]},
               ]}
        prow, arows = _rows_from_payment(pay, _invoice_index(con))
        assert prow["total_amt"] == 150000.0 and prow["customer_id"] == "77", prow
        assert len(arows) == 2, arows
        # parent-GC hierarchy walk: leaf 'RP1234 ...' (99) → parent 'Firestone' (77), 2 levels deep
        from shared.qbo_api import extract_proj as _xp
        fake_q = lambda a, c, e, where="": [
            {"Id": "77", "DisplayName": "Firestone Building Co", "ParentRef": None},
            {"Id": "99", "DisplayName": "RP1234 - SOME JOB", "ParentRef": {"value": "77"}},
        ]
        gmap, pmap = _customer_gc_map(fake_q, None, None, _xp)
        assert gmap["99"] == ("Firestone Building Co", "77"), gmap        # leaf resolves up to the GC
        assert gmap["77"] == ("Firestone Building Co", "77"), gmap        # a top-level GC resolves to itself
        assert pmap.get("RP1234") == ("Firestone Building Co", "77"), pmap   # project reversed from Customer:Project
        prow["parent_customer"], prow["parent_customer_id"] = gmap["99"]
        got = {a["invoice_txn_id"]: a for a in arows}
        assert got["INV9"]["invoice_no"] == "34999" and got["INV9"]["project_no"] == "CP861", got
        assert got["INV_UNKNOWN"]["invoice_no"] is None, got            # unresolved still recorded by id
        assert got["INV9"]["amount"] == 100000.0 and got["INV_UNKNOWN"]["amount"] == 50000.0, got
        assert got["INV9"]["invoice_open"] == 5000 and got["INV_UNKNOWN"]["invoice_open"] is None, got   # open balance flows for resolved
        # write + grouped read-back
        now = "2026-08-19T00:00:00"
        con.execute("INSERT INTO payment (qbo_txn_id, txn_date, customer, customer_id, parent_customer, "
                    "parent_customer_id, total_amt, unapplied_amt, method, ref_no, source, loaded_at) VALUES "
                    "(:qbo_txn_id,:txn_date,:customer,:customer_id,:parent_customer,:parent_customer_id,"
                    ":total_amt,:unapplied_amt,:method,:ref_no,'qbo_payment',:loaded_at)",
                    {**prow, "loaded_at": now})
        for a in arows:
            con.execute("INSERT INTO payment_application (payment_txn_id, invoice_txn_id, invoice_no, "
                        "project_no, division, amount, invoice_open) VALUES "
                        "(:payment_txn_id,:invoice_txn_id,:invoice_no,:project_no,:division,:amount,:invoice_open)", a)
        con.commit()
        n = con.execute("SELECT COUNT(*) FROM payment_application WHERE payment_txn_id='P1'").fetchone()[0]
        assert n == 2, n
        print("selftest OK - payment + 2 applications round-trip; unresolved link kept by id.")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(description="Load received QBO customer payments into the ledger.")
    ap.add_argument("--db", type=Path, default=DEFAULT_DB)
    ap.add_argument("--months", type=int, default=12, help="rolling window to pull (default 12).")
    ap.add_argument("--dry-run", action="store_true", help="pull + report, write nothing.")
    ap.add_argument("--selftest", action="store_true", help="offline proof; no QBO.")
    args = ap.parse_args()

    if args.selftest:
        return _selftest()

    con = _connect(args.db)
    print(f"Payments → {args.db}")
    _load(con, args.months, args.dry_run)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
