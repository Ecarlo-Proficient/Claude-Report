#!/usr/bin/env python3
"""bill_payment_stub.py - a check / bill-payment stub in QuickBooks' report look, the PAYMENT on
top and the bills it paid below (owner 2026-09-22: "it needs to look like qbo format, but we need
the payment on top and the bills it's paying below").

One stub per payment, one payment per page. Columns are a registry the caller adds to or drops
from; the default set is QBO's own "Bills and Applied Payments" columns:

    date · type · number · memo · amount · open_balance

`amount` is what THIS payment applied to the bill (the column ties to the payment total); the
bill's own total is the optional `bill_total` column. Both stubs read each bill as Open balance - Paid = New
balance (owner 2026-10-07: "shown on amount paid - open balance = new total"; the vendor stub too the same day:
"fix the partials ... use columns", no partial sentence): the balance just before THIS payment (the bill total
less every earlier payment on it, from the bill's own LinkedTxn), what it applied and what is left. On the internal
stub a check drawn on the Joint Checks account shows its type as "Joint check".

Reads the raw QBO mirror only (shared/qbo_mirror: BillPayment -> Line[].LinkedTxn -> Bill; Vendor
for the print name and address). Read-only everywhere; writes just the HTML/PDF it was asked for.
PDF = the HTML printed by Chrome headless (Letter portrait), the same recipe as the process guides.

    python3 ledger/bill_payment_stub.py --vendor COWTOWN --date 2026-09-21
    python3 ledger/bill_payment_stub.py --payment 1318748 --payment 1318747
    python3 ledger/bill_payment_stub.py --vendor COWTOWN --from 2026-09-01 --to 2026-09-30
    --columns date,number,memo,amount          the exact set, in that order
    --add bill_total,project --drop type       edit the default set
    --list-columns                             every column the registry knows
    --company "<name on the stub>"             default: ACB_COMPANY_NAME in machine.env
    --out <file.pdf>                           default: Accounting/Accounts Payable/Bill Payment Stubs/<vendor>/
                                               <Vendor> - Bill Payment Stub <check # | ref> <mm-dd-yyyy>.pdf (ONE
                                               file per payment;
                                               the share must be mounted, a missing mount stops the run)
    --html                                     keep the .html beside the .pdf
    --selftest                                 offline proof on synthetic records, no mirror, no Chrome

Two kinds (owner 2026-10-07: "a vendor stub and a internal stub that shows the inv it is tied to"):
`vendor` (the default, the one mailed with the check) and `internal`: the client in the header, then each
CLIENT INVOICE the money came from on top with the bills it paid below it (plus the Project column). A joint
check is tied to its invoices through QuickBooks itself: the client's Payment with the same date and amount
(the check # as its ref when typed) lists the invoices, and the bills fill them top to bottom. Any other
payment is grouped by the invoice each bill is billed through (the caller's bill -> invoice map, the ledger's
vendor page record from the Bill Tracker); a bill with none sits under "No client invoice yet". An internal
stub files under `<vendor>/Internal/`.

History (owner 2026-09-22: "in case a bill payment gets changed, deleted, voided and us still see
those historical"): every print is a row in the ledger's `bill_payment_stub` table with the stub
record as printed (snapshot), the file path and the QBO SyncToken at print time. `history()` judges
each print against the mirror NOW - current / changed / voided / deleted - and the PDF stays on disk,
so a payment QuickBooks no longer shows still has its stubs. The dashboard's vendor page prints
through `print_stub()` and lists `history()`; `--history` shows it on the command line.
"""
from __future__ import annotations

import argparse
import datetime as dt
import html
import json
import shutil
import sqlite3
import subprocess
import sys
import tempfile
import time
from pathlib import Path
from typing import Callable, Dict, List, Optional, Sequence, Tuple

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))
from shared import paths            # noqa: E402
from shared.qbo_api import PROJ_RE  # noqa: E402

CHROME_CANDIDATES = (
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    "/Applications/Chromium.app/Contents/MacOS/Chromium",
    "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
)


# ───────────────────────── helpers ─────────────────────────

def _num(x) -> float:
    try:
        return round(float(x), 2)
    except (TypeError, ValueError):
        return 0.0


def money(x: float) -> str:
    """QBO report style: 27,251.94 / (1,000.00) for negatives, no $ sign inside the table."""
    x = _num(x)
    s = f"{abs(x):,.2f}"
    return f"({s})" if x < 0 else s


def mdy(iso: Optional[str]) -> str:
    """ISO date -> mm/dd/yyyy (the owner's date format, never year-first)."""
    if not iso:
        return ""
    try:
        return dt.date.fromisoformat(str(iso)[:10]).strftime("%m/%d/%Y")
    except ValueError:
        return str(iso)


def _esc(s) -> str:
    return html.escape("" if s is None else str(s))


def _projects_of_bill(bill: dict) -> str:
    """Distinct job numbers the bill's lines are coded to, in line order (RP7530 · RP6676-FTW)."""
    seen: List[str] = []
    for ln in bill.get("Line", []) or []:
        det = ln.get("ItemBasedExpenseLineDetail") or ln.get("AccountBasedExpenseLineDetail") or {}
        name = (det.get("CustomerRef") or {}).get("name") or ""
        m = PROJ_RE.search(name)
        if m:
            p = m.group(1).upper()
            if p not in seen:
                seen.append(p)
    return " · ".join(seen)


# ───────────────────────── the column registry ─────────────────────────
# key -> (header label, alignment, value getter over a stub row). Add a column here and every
# caller (CLI today, the ledger's stub button later) can switch it on.

Row = Dict[str, object]

COLUMNS: Dict[str, Tuple[str, str, Callable[[Row], str]]] = {
    "date":         ("Date",               "left",  lambda r: mdy(r["bill_date"])),
    "type":         ("Transaction type",   "left",  lambda r: str(r.get("type") or "Bill")),
    "number":       ("Transaction number", "left",  lambda r: str(r["number"] or "")),
    "memo":         ("Memo/Description",   "left",  lambda r: str(r["memo"] or "")),
    "project":      ("Project",            "left",  lambda r: str(r["project"] or "")),
    "due_date":     ("Due date",           "left",  lambda r: mdy(r["due_date"])),
    "bill_total":   ("Bill amount",        "right", lambda r: money(r["bill_total"])),
    # the money as of THIS payment: Open balance - Paid = New balance (owner 2026-10-07; the picker shows these
    # same labels, so what is ticked is what prints - the keys stay for saved picks and the print history)
    "balance_before": ("Open balance",     "right", lambda r: "" if r.get("balance_before") is None else money(r["balance_before"])),
    "amount":       ("Paid",               "right", lambda r: money(r["paid"])),
    "open_balance": ("New balance",        "right", lambda r: "" if r.get("balance_after") is None else money(r["balance_after"])),
}
KINDS = ("vendor", "internal")
MONEY_ORDER = ("bill_total", "balance_before", "amount", "open_balance")   # always on the right, in this order
DEFAULT_COLUMNS: Tuple[str, ...] = ("date", "type", "number", "memo", "balance_before", "amount", "open_balance")
WIDTHS = {"date": 64, "type": 88, "number": 100, "project": 90, "due_date": 64, "bill_total": 74, "amount": 74,
          "open_balance": 80, "balance_before": 80}   # px on a 7.3in Letter page; memo takes what is left
MONEY_COLUMNS = {"bill_total", "balance_before", "amount", "open_balance"}   # get a total on the last row


def resolve_columns(columns: Optional[Sequence[str]] = None,
                    add: Sequence[str] = (), drop: Sequence[str] = ()) -> List[str]:
    """The ordered column keys for a render: an explicit set, or the default edited by add/drop.
    Unknown keys raise so a typo never silently drops a column."""
    base = list(columns) if columns else list(DEFAULT_COLUMNS)
    for k in add:
        if k not in base:
            base.append(k)
    base = [k for k in base if k not in set(drop)]
    bad = [k for k in base if k not in COLUMNS]
    if bad:
        raise ValueError(f"unknown column(s): {', '.join(bad)} - known: {', '.join(COLUMNS)}")
    if not base:
        raise ValueError("no columns left to print")
    return base


def columns_for(kind: str, columns: Optional[Sequence[str]] = None) -> List[str]:
    """The columns a print of this kind gets. Vendor: exactly the pick (owner 2026-10-07: "columns not matching
    stub"), the money ones moved to the right in MONEY_ORDER. Internal: the pick without Transaction type (always
    "Bill"; its width goes to the memo), Project after the memo, and always Open balance - Paid = New balance.
    Idempotent."""
    if kind not in KINDS:
        raise ValueError(f"unknown stub kind {kind!r} - {', '.join(KINDS)}")
    cols = resolve_columns(columns)
    if kind == "vendor":
        return [k for k in cols if k not in MONEY_ORDER] + [k for k in MONEY_ORDER if k in cols]
    tail = ["balance_before", "amount", "open_balance"]
    cols = [k for k in cols if k != "type"] or cols
    cols = [k for k in cols if k not in tail] + tail
    if "project" not in cols:
        at = cols.index("memo") + 1 if "memo" in cols else (cols.index("number") + 1 if "number" in cols else 0)
        cols.insert(at, "project")
    return cols


INVOICE_STATE_TEXT = {"awaiting": "Not invoiced yet", "noproject": "No project #", "untracked": "Not tracked", "open": "Not paid"}


def _draw_label(inv: dict) -> str:
    """The invoice's own name for itself: the last " - " part of its memo's first line (… - May Draw 2026)."""
    first = str(inv.get("PrivateNote") or "").splitlines()[:1]
    line = first[0].strip() if first else ""
    return line.rsplit(" - ", 1)[-1].strip() if line else ""


def _group(inv: Optional[dict], label: str = "") -> dict:
    inv = inv or {}
    m = PROJ_RE.search(str((inv.get("CustomerRef") or {}).get("name") or ""))
    return {"invoice_no": inv.get("DocNumber") or "", "date": inv.get("TxnDate"), "label": label or _draw_label(inv),
            "project": m.group(1).upper() if m else "",
            "invoice_total": _num(inv.get("TotalAmt")) if inv else None, "rows": [], "paid": 0.0}


def _segment(r: Row, portion: float, before: Optional[float]) -> Row:
    seg = dict(r)
    seg["paid"] = round(portion, 2)
    seg["balance_before"] = before
    seg["balance_after"] = None if before is None else round(before - portion, 2)
    return seg


def apply_internal(stub: dict, joint_payment: Optional[dict], txn_of: Callable[[str, str], Optional[dict]],
                   invoice_by_doc: Callable[[str], Optional[dict]], client_of: Callable[[str], str],
                   invoice_of: Optional[Dict[str, dict]] = None) -> dict:
    """The internal stub's shape (owner 2026-10-07: "the invoice on top and the bills below it"): stub["groups"] =
    one block per client invoice, each with the bills (or the part of a bill) this payment paid from it, and
    stub["client"] for the header.

    Joint check: `joint_payment` is the client's Payment that funded it; its invoices, oldest first, are filled
    by the bills top to bottom, so a bill can straddle two invoices (each part shows its own Open balance - Paid =
    New balance, chained). Otherwise: each bill under the invoice it is billed through (`invoice_of`, the vendor
    page record), a bill with none under "No client invoice yet"."""
    invoice_of = invoice_of or {}
    stub["joint"] = bool(joint_payment) or (stub.get("method") == "Check" and is_joint_account(stub.get("account")))
    groups: List[dict] = []
    bills = [r for r in stub["rows"] if r.get("txn_type") == "Bill"]
    others = [r for r in stub["rows"] if r.get("txn_type") != "Bill"]
    clients: List[str] = []
    if joint_payment:
        c = client_of(str((joint_payment.get("CustomerRef") or {}).get("value") or ""))
        if c:
            clients.append(c)
        invs = []
        for ln in joint_payment.get("Line") or []:
            for t in ln.get("LinkedTxn") or []:
                if t.get("TxnType") == "Invoice":
                    invs.append((txn_of("Invoice", str(t["TxnId"])) or {"DocNumber": str(t["TxnId"])}, _num(ln.get("Amount"))))
        invs.sort(key=lambda x: (str(x[0].get("TxnDate") or ""), _natural(x[0].get("DocNumber"))))
        queue = [[r, _num(r["paid"]), r.get("balance_before")] for r in bills]    # bill, left to place, its running balance
        for inv, applied in invs:
            g = _group(inv)
            room = applied
            while queue and room > 0.005:
                r, left, before = queue[0]
                take = round(min(left, room), 2)
                g["rows"].append(_segment(r, take, before))
                room = round(room - take, 2)
                queue[0][1] = round(left - take, 2)
                queue[0][2] = None if before is None else round(before - take, 2)
                if queue[0][1] <= 0.005:
                    queue.pop(0)
            groups.append(g)
        if queue:                                         # the check paid more than the client's payment covered
            g = _group(None, "Not covered by the client's payment")
            g["rows"] = [_segment(r, left, before) for r, left, before in queue]
            groups.append(g)
    else:
        by: Dict[str, dict] = {}
        for r in bills:
            iv = invoice_of.get(str(r.get("bill_id") or "")) or {}
            for c in iv.get("clients") or []:
                if c not in clients:
                    clients.append(c)
            no = str(iv.get("invoice_no") or "")
            key = no or "none"
            if key not in by:
                by[key] = _group(invoice_by_doc(no) if no else None, "" if no else "No client invoice yet")
                if no and not by[key]["invoice_no"]:
                    by[key]["invoice_no"] = no
                groups.append(by[key])
            by[key]["rows"].append(_segment(r, _num(r["paid"]), r.get("balance_before")))
        groups.sort(key=lambda g: (not g["invoice_no"], str(g["date"] or ""), _natural(g["invoice_no"])))
    if others:
        g = _group(None, "Credits and amounts not on a bill")
        g["rows"] = [_segment(r, _num(r["paid"]), None) for r in others]
        groups.append(g)
    for g in groups:
        g["paid"] = round(sum(_num(r["paid"]) for r in g["rows"]), 2)
        if not g.get("project"):                          # the invoice's own job; else the jobs of the bills under it
            g["project"] = " · ".join(dict.fromkeys(p for r in g["rows"] for p in str(r.get("project") or "").split(" · ") if p))
    stub["groups"] = groups
    stub["client"] = ", ".join(clients)
    return stub


def find_joint_payment(bp: dict, payments_on: Callable[[str], List[dict]]) -> Optional[dict]:
    """The client's Payment behind a joint check: same date, same amount; the check # as its ref breaks a tie.
    None when the check is not on the Joint Checks account or no single payment matches."""
    method, ref, acct = _pay_method(bp)
    if method != "Check" or not is_joint_account(acct):
        return None
    total = _num(bp.get("TotalAmt"))
    hits = [p for p in payments_on(str(bp.get("TxnDate") or "")) if abs(_num(p.get("TotalAmt")) - total) < 0.005]
    if len(hits) > 1 and ref:
        hits = [p for p in hits if str(p.get("PaymentRefNum") or "") == str(ref)] or hits
    return hits[0] if len(hits) == 1 else None


# ───────────────────────── the data (mirror -> stub records) ─────────────────────────

def _pay_method(bp: dict) -> Tuple[str, str, str]:
    """(method label, reference number, account name) for the payment header."""
    ptype = bp.get("PayType") or ""
    ref = bp.get("DocNumber") or ""
    if ptype == "Check":
        acct = ((bp.get("CheckPayment") or {}).get("BankAccountRef") or {}).get("name") or ""
        return "Check", ref, acct
    if ptype == "CreditCard":
        acct = ((bp.get("CreditCardPayment") or {}).get("CCAccountRef") or {}).get("name") or ""
        return "Credit card", ref, acct
    return ptype or "Payment", ref, ""


def is_joint_account(acct_name: str) -> bool:
    """A check written from the Joint Checks account is a joint check (the client's check to us + the vendor)."""
    return "joint" in str(acct_name or "").lower()


CHECK_METHODS = ("Check", "Joint check")

TYPE_LABEL = {"Bill": "Bill", "VendorCredit": "Vendor credit", "Deposit": "Deposit", "JournalEntry": "Journal entry", "Purchase": "Expense"}


def _natural(s) -> tuple:
    import re
    return tuple((0, int(t)) if t.isdigit() else (1, t.lower()) for t in re.findall(r"\d+|\D+", str(s or "")))


def _applied_to(pay: dict, bill_id: str) -> float:
    """What one bill payment applied to one bill (a bill can sit on several of its lines)."""
    return round(sum(_num(ln.get("Amount")) for ln in pay.get("Line") or []
                     if any(str(t.get("TxnId")) == str(bill_id) and (t.get("TxnType") or "Bill") == "Bill"
                            for t in ln.get("LinkedTxn") or [])), 2)


def _balance_before(bill: dict, bp: dict, bill_id: str, txn_of: Callable[[str, str], Optional[dict]]) -> float:
    """The bill's open balance just before THIS payment: its total less every payment QuickBooks links to it that
    comes earlier (by date, then by QuickBooks id for two the same day)."""
    me = (str(bp.get("TxnDate") or ""), int(bp.get("Id") or 0) if str(bp.get("Id") or "").isdigit() else 0)
    prior = 0.0
    for t in bill.get("LinkedTxn") or []:
        if not str(t.get("TxnType") or "").startswith("BillPayment") or str(t.get("TxnId")) == str(bp.get("Id")):
            continue
        pay = txn_of("BillPayment", str(t.get("TxnId"))) or {}
        pid = str(pay.get("Id") or "")
        key = (str(pay.get("TxnDate") or ""), int(pid) if pid.isdigit() else 0)
        if pay and key < me:
            prior += _applied_to(pay, bill_id)
    return round(_num(bill.get("TotalAmt")) - prior, 2)


def build_stub(bp: dict, txn_of: Callable[[str, str], Optional[dict]],
               vendor: Optional[dict] = None) -> dict:
    """One BillPayment (as QBO returned it) -> the stub record: header + rows + total. `txn_of(type, id)`
    resolves a linked transaction (Bill, or a VendorCredit / Deposit / JournalEntry QBO applied). Lines are
    merged per linked transaction (a bill linked on two lines is ONE row) and an amount applied to nothing
    is its own "Unapplied" row, so the rows always add up to the payment."""
    method, ref, acct = _pay_method(bp)
    vref = bp.get("VendorRef") or {}
    vendor = vendor or {}
    addr = vendor.get("BillAddr") or {}
    addr_lines = [addr.get(k) for k in ("Line1", "Line2", "Line3") if addr.get(k)]
    city = " ".join(x for x in (addr.get("City"), addr.get("CountrySubDivisionCode"), addr.get("PostalCode")) if x)
    if city:
        addr_lines.append(city)
    rows: List[Row] = []
    agg: Dict[Tuple[str, str], float] = {}
    unapplied = 0.0
    for ln in bp.get("Line", []) or []:
        paid = _num(ln.get("Amount"))
        linked = [t for t in (ln.get("LinkedTxn") or []) if t.get("TxnId")]
        if not linked:
            unapplied = round(unapplied + paid, 2)
            continue
        key = (str(linked[0].get("TxnType") or "Bill"), str(linked[0]["TxnId"]))
        agg[key] = round(agg.get(key, 0.0) + paid, 2)
    for (ttype, tid), paid in agg.items():
        rec = txn_of(ttype, tid) or {}
        total = _num(rec.get("TotalAmt"))
        is_bill = ttype == "Bill"
        before = _balance_before(rec, bp, tid, txn_of) if is_bill and rec else None
        rows.append({
            "bill_id": tid,
            "txn_type": ttype,
            "type": TYPE_LABEL.get(ttype, ttype),
            "bill_date": rec.get("TxnDate"),
            "due_date": rec.get("DueDate"),
            "number": rec.get("DocNumber") or "",
            "memo": rec.get("PrivateNote") or "",
            "project": _projects_of_bill(rec) if is_bill else "",
            "bill_total": total,
            "paid": paid,
            "open_balance": _num(rec.get("Balance")),
            "balance_before": before,
            "balance_after": None if before is None else round(before - paid, 2),
            "partial": is_bill and bool(rec) and abs(total - paid) > 0.005,
        })
    rows.sort(key=lambda r: (str(r["bill_date"] or ""), _natural(r["number"])))
    if abs(unapplied) > 0.005:
        rows.append({"bill_id": "", "txn_type": "", "type": "Unapplied", "bill_date": bp.get("TxnDate"), "due_date": None,
                     "number": "", "memo": "amount not applied to a bill", "project": "", "bill_total": unapplied,
                     "paid": unapplied, "open_balance": 0.0, "balance_before": None, "balance_after": None, "partial": False})
    total_paid = round(sum(r["paid"] for r in rows), 2)   # type: ignore[misc]
    return {
        "payment_id": str(bp.get("Id")),
        "txn_date": bp.get("TxnDate"),
        "vendor": vref.get("name") or vendor.get("DisplayName") or "",
        "pay_to": vendor.get("PrintOnCheckName") or vendor.get("DisplayName") or vref.get("name") or "",
        "address": addr_lines,
        "method": method,
        "ref": ref,
        "account": acct,
        "total": _num(bp.get("TotalAmt")),
        "rows": rows,
        "rows_total": total_paid,
        "ties": abs(total_paid - _num(bp.get("TotalAmt"))) < 0.005,
    }


def load_payments(vendor: str = "", date_from: str = "", date_to: str = "",
                  payment_ids: Sequence[str] = ()) -> List[dict]:
    """Stub records from the mirror: by payment id, or by vendor (substring, case-insensitive)
    within a date window. Newest payment first, then by reference."""
    from shared import qbo_mirror as mirror
    con = mirror.connect()
    try:
        if payment_ids:
            bps = [b for b in (mirror.get("BillPayment", i, con=con) for i in payment_ids) if b]
        else:
            where, params = [], []
            if date_from:
                where.append("txn_date >= ?")
                params.append(date_from)
            if date_to:
                where.append("txn_date <= ?")
                params.append(date_to)
            bps = mirror.load("BillPayment", " AND ".join(where), tuple(params), con=con)
            if vendor:
                v = vendor.lower()
                bps = [b for b in bps if v in ((b.get("VendorRef") or {}).get("name") or "").lower()]
        vendors: Dict[str, Optional[dict]] = {}
        txns: Dict[Tuple[str, str], Optional[dict]] = {}

        def txn_of(ttype: str, tid: str) -> Optional[dict]:
            if (ttype, tid) not in txns:
                txns[(ttype, tid)] = mirror.get(ttype, tid, con=con) if ttype in mirror.ENTITIES else None
            return txns[(ttype, tid)]

        out = []
        for bp in bps:
            vid = str((bp.get("VendorRef") or {}).get("value") or "")
            if vid and vid not in vendors:
                vendors[vid] = mirror.get("Vendor", vid, con=con)
            st = build_stub(bp, txn_of, vendors.get(vid))
            st["vendor_id"] = vid
            st["_raw"] = {"SyncToken": bp.get("SyncToken"), "MetaData": bp.get("MetaData")}
            out.append(st)
        out.sort(key=lambda s: (str(s["txn_date"] or ""), str(s["ref"])), reverse=True)
        return out
    finally:
        con.close()


# ───────────────────────── history (the ledger's bill_payment_stub table) ─────────────────────────

DEFAULT_LEDGER_DB = paths.get_path(
    "ACB_LEDGER_DB", Path.home() / "Library" / "Application Support" / "Proficient" / "ledger.sqlite3")

_STUB_TABLE = """
CREATE TABLE IF NOT EXISTS bill_payment_stub (
    id INTEGER PRIMARY KEY AUTOINCREMENT, payment_id TEXT NOT NULL, vendor TEXT, vendor_id TEXT,
    txn_date TEXT, method TEXT, ref TEXT, total NUMERIC, n_bills INTEGER, sync_token TEXT,
    last_updated TEXT, columns TEXT NOT NULL, snapshot TEXT NOT NULL, file_path TEXT, archive_path TEXT, printed_at TEXT NOT NULL,
    kind TEXT NOT NULL DEFAULT 'vendor');
CREATE INDEX IF NOT EXISTS ix_bpstub_payment ON bill_payment_stub (payment_id);
CREATE INDEX IF NOT EXISTS ix_bpstub_vendor ON bill_payment_stub (vendor);
"""


def ledger_connect(db_path: Optional[Path] = None) -> sqlite3.Connection:
    """Read-write connection to the ledger for the history table only (the dashboard itself opens
    the ledger read-only; this is the one write this module makes)."""
    con = sqlite3.connect(str(db_path or DEFAULT_LEDGER_DB))
    con.row_factory = sqlite3.Row
    con.executescript(_STUB_TABLE)
    cols = {r[1] for r in con.execute("PRAGMA table_info(bill_payment_stub)")}
    if "archive_path" not in cols:                       # table made before the archive copy existed
        con.execute("ALTER TABLE bill_payment_stub ADD COLUMN archive_path TEXT")
        con.commit()
    if "kind" not in cols:                               # table made before the internal stub existed
        con.execute("ALTER TABLE bill_payment_stub ADD COLUMN kind TEXT NOT NULL DEFAULT 'vendor'")
        con.commit()
    return con


STUB_ARCHIVE_DIR = paths.get_path(
    "ACB_STUB_ARCHIVE_DIR", Path.home() / "Library" / "Application Support" / "Proficient" / "bill-payment-stubs")


def record_print(con: sqlite3.Connection, stub: dict, columns: Sequence[str], file_path: Path,
                 raw: Optional[dict] = None, printed_at: Optional[str] = None,
                 archive_dir: Optional[Path] = None) -> int:
    """One row per print, the stub exactly as it went on paper. The vendor-folder file is the LATEST
    print (a re-print overwrites it, same name); the history keeps its own immutable copy of every
    print beside the ledger DB (`<archive>/<history id>.pdf`), so an older print can always be opened
    as it went out. Returns the history id."""
    raw = raw or {}
    cur = con.execute(
        "INSERT INTO bill_payment_stub (payment_id, vendor, vendor_id, txn_date, method, ref, total, n_bills, "
        "sync_token, last_updated, columns, snapshot, file_path, printed_at, kind) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
        (stub["payment_id"], stub["vendor"], stub.get("vendor_id"), stub["txn_date"], stub["method"], stub["ref"],
         stub["total"], len(stub["rows"]), str(raw.get("SyncToken", "")),
         (raw.get("MetaData") or {}).get("LastUpdatedTime", ""), ",".join(columns), json.dumps(stub, default=str),
         str(file_path), printed_at or dt.datetime.now().isoformat(timespec="seconds"), stub.get("kind") or "vendor"))
    hid = int(cur.lastrowid)
    arch = None
    try:
        adir = archive_dir or STUB_ARCHIVE_DIR
        adir.mkdir(parents=True, exist_ok=True)
        arch = adir / f"{hid}.pdf"
        shutil.copyfile(file_path, arch)
        con.execute("UPDATE bill_payment_stub SET archive_path=? WHERE id=?", (str(arch), hid))
    except OSError:
        arch = None                                       # the vendor-folder file still serves
    con.commit()
    return hid


def mirror_state(payment_id: str, con=None) -> dict:
    """What the mirror says about the payment NOW: present / deleted / voided, its SyncToken, total."""
    from shared import qbo_mirror as mirror
    own = con is None
    con = con or mirror.connect()
    try:
        r = con.execute("SELECT sync_token, last_updated, total, deleted, deleted_at, json FROM qbo_billpayment WHERE id=?",
                        (str(payment_id),)).fetchone()
        if not r:
            return {"present": False, "deleted": False, "voided": False, "sync_token": None, "total": None}
        rec = mirror._unpack(r[5])
        voided = is_voided(rec)
        return {"present": True, "deleted": bool(r[3]), "deleted_at": r[4], "voided": voided,
                "sync_token": _tok(r[0]), "last_updated": r[1], "total": _num(r[2])}
    finally:
        if own:
            con.close()


def _tok(x) -> str:
    """A SyncToken as text. QBO's first token is 0 - falsy - so never `x or ""`."""
    return "" if x is None else str(x)


def is_voided(rec: dict) -> bool:
    """QBO voids a bill payment by zeroing it, writing "Voided…" in the memo and dropping its lines - the ONE
    rule (the vendor page's `_mark_voided_payments` applies the same tell)."""
    note = str(rec.get("PrivateNote") or "")
    return _num(rec.get("TotalAmt")) == 0 and (note.strip().upper().startswith("VOID") or not rec.get("Line"))


def print_status(row: dict, now: dict) -> str:
    """current · changed · voided · deleted - the print judged against QBO today."""
    if now.get("deleted") or not now.get("present"):
        return "deleted"
    if now.get("voided") and _num(row.get("total")) != 0:
        return "voided"
    if _tok(now.get("sync_token")) != _tok(row.get("sync_token")) or abs(_num(now.get("total")) - _num(row.get("total"))) > 0.005:
        return "changed"
    return "current"


def history(con: sqlite3.Connection, vendor: str = "", payment_id: str = "") -> List[dict]:
    """Every print (newest first) with its status against the mirror now and whether the PDF still exists."""
    from shared import qbo_mirror as mirror
    where, params = [], []
    if vendor:
        where.append("vendor = ?")
        params.append(vendor)
    if payment_id:
        where.append("payment_id = ?")
        params.append(str(payment_id))
    sql = "SELECT id, payment_id, vendor, vendor_id, txn_date, method, ref, total, n_bills, sync_token, last_updated, " \
          "columns, file_path, archive_path, printed_at, kind FROM bill_payment_stub"
    if where:
        sql += " WHERE " + " AND ".join(where)
    sql += " ORDER BY printed_at DESC, id DESC"
    rows = [dict(r) for r in con.execute(sql, params)]
    if not rows:
        return []
    states: Dict[str, dict] = {}
    mcon = mirror.connect() if mirror.db_path().exists() else None
    try:
        for r in rows:
            pid = r["payment_id"]
            if mcon and pid not in states:
                states[pid] = mirror_state(pid, mcon)
            now = states.get(pid)
            r["checked"] = bool(now)
            r["status"] = print_status(r, now) if now else "unchecked"     # no mirror on this machine: nothing to judge against
            r["qbo_total_now"] = now.get("total") if now else None
            r["deleted_at"] = now.get("deleted_at") if now else None
            # a print with an archive copy exists only if THAT copy is there; the vendor-folder file is the latest
            # print and stands in only for rows made before the archive existed
            r["file_exists"] = Path(r["archive_path"]).is_file() if r["archive_path"] else bool(r["file_path"] and Path(r["file_path"]).is_file())
            r["columns"] = r["columns"].split(",") if r["columns"] else list(DEFAULT_COLUMNS)
    finally:
        if mcon:
            mcon.close()
    return rows


def history_snapshot(con: sqlite3.Connection, hist_id: int) -> Optional[dict]:
    r = con.execute("SELECT snapshot, columns, file_path, archive_path FROM bill_payment_stub WHERE id=?", (int(hist_id),)).fetchone()
    if not r:
        return None
    # the print AS IT WENT OUT: the archived copy; the vendor-folder file (the LATEST print) stands in only for a
    # row made before the archive existed - never for a row whose own copy is missing (that would show a newer print)
    if r[3]:
        served = r[3] if Path(r[3]).is_file() else None
    else:
        served = r[2] if r[2] and Path(r[2]).is_file() else None
    return {"stub": json.loads(r[0]), "columns": r[1].split(","), "file_path": r[2], "archive_path": r[3], "served_path": served}


def print_stub(payment_id: str, columns: Optional[Sequence[str]] = None, company: str = "",
               db_path: Optional[Path] = None, kind: str = "vendor",
               invoice_of: Optional[Dict[str, dict]] = None) -> dict:
    """The dashboard's print: build from the mirror (a deleted payment still builds - the mirror keeps
    it), write the PDF into the vendor folder, record the history row. Returns {ok, id, file, stub}.
    kind="internal" groups the bills under the client invoices they were paid from (see apply_internal). A payment the
    mirror has not seen yet answers {ok: False, missing: True} so the caller can refresh and retry."""
    from shared import qbo_mirror as mirror
    cols = columns_for(kind, columns)
    mcon = mirror.connect()
    try:
        raw = mirror.get("BillPayment", str(payment_id), con=mcon)
        if not raw:
            return {"ok": False, "missing": True,
                    "error": f"payment {payment_id} is not in the ledger's copy of QuickBooks yet - Refresh from QuickBooks and print again"}
        if is_voided(raw):
            return {"ok": False, "error": f"payment {raw.get('DocNumber') or payment_id} is voided in QuickBooks - nothing to print"}
        vid = str((raw.get("VendorRef") or {}).get("value") or "")
        vendor = mirror.get("Vendor", vid, con=mcon) if vid else None
        get = lambda t, i: mirror.get(t, i, con=mcon) if t in mirror.ENTITIES else None   # noqa: E731
        stub = build_stub(raw, get, vendor)
        stub["vendor_id"] = vid
        stub["kind"] = kind
        if kind == "internal":
            def invoice_by_doc(doc: str) -> Optional[dict]:
                hit = mirror.load("Invoice", "doc_number = ?", (str(doc),), con=mcon)
                return hit[0] if len(hit) == 1 else None

            def client_of(cust_id: str) -> str:        # the top parent we know (the GC), never the project
                name, seen = "", set()
                while cust_id and cust_id not in seen:
                    seen.add(cust_id)
                    c = mirror.get("Customer", cust_id, con=mcon)
                    if not c:
                        break
                    name = c.get("DisplayName") or name
                    cust_id = str((c.get("ParentRef") or {}).get("value") or "")
                return name
            joint = find_joint_payment(raw, lambda d: mirror.load("Payment", "txn_date = ?", (d,), con=mcon))
            apply_internal(stub, joint, get, invoice_by_doc, client_of, invoice_of)
    finally:
        mcon.close()
    out = default_out(stub)
    out.parent.mkdir(parents=True, exist_ok=True)
    html_path = out.with_suffix(".html")
    html_path.write_text(stub_html([stub], cols, company), encoding="utf-8")
    ok = render_pdf(html_path, out)
    html_path.unlink(missing_ok=True)
    if not ok:
        return {"ok": False, "error": "Chrome not found or the print failed"}
    lcon = ledger_connect(db_path)
    try:
        hid = record_print(lcon, stub, cols, out, raw)
    finally:
        lcon.close()
    return {"ok": True, "id": hid, "file": str(out), "columns": cols, "kind": kind, "stub": stub}


def columns_registry() -> List[dict]:
    """The column picker (the internal print adapts the same pick: see columns_for)."""
    return [{"key": k, "label": v[0], "default": k in DEFAULT_COLUMNS} for k, v in COLUMNS.items()]


# ───────────────────────── the page ─────────────────────────

CSS = """
@page { size: Letter portrait; margin: 0.55in 0.6in; }
* { box-sizing: border-box; }
body { margin: 0; font-family: Helvetica, Arial, -apple-system, sans-serif; font-size: 10.5px; color: #111; }
.stub { page-break-after: always; }
.stub:last-child { page-break-after: auto; }
.hdr { text-align: center; margin: 0 0 18px; }
.hdr .co { font-size: 20px; font-weight: 700; margin: 0; }
.hdr .rpt { font-size: 14px; margin: 3px 0 0; }
.hdr .asof { font-size: 12px; margin: 2px 0 0; }
.pay { display: flex; justify-content: space-between; align-items: flex-start; border-top: 1px solid #333;
       border-bottom: 1px solid #333; padding: 9px 4px; margin-bottom: 14px; }
.pay .to { max-width: 55%; }
.pay .to .lbl, .pay .kv .lbl { font-size: 9.5px; color: #555; text-transform: uppercase; letter-spacing: .04em; }
.pay .to .name { font-size: 13px; font-weight: 700; margin: 2px 0 1px; }
.pay .to .addr { color: #333; line-height: 1.35; }
.pay .kv { display: grid; grid-template-columns: max-content max-content; column-gap: 14px; row-gap: 2px; text-align: left; }
.pay .kv .v { font-weight: 600; }
.pay .kv .amt { font-size: 14px; font-weight: 700; }
table { width: 100%; border-collapse: collapse; }
th { font-weight: 400; text-align: left; padding: 5px 4px; border-top: 1px solid #333; border-bottom: 1px solid #333;
     vertical-align: bottom; white-space: nowrap; }
td { padding: 3px 4px; vertical-align: top; }
td.memo, th.memo { width: auto; }
th.r, td.r { text-align: right; white-space: nowrap; }
tr.band td { background: #e9e9ec; font-weight: 400; padding: 5px 4px; }
tr.row td { border-bottom: 1px solid #f0f0f0; }
td.memo { word-break: break-word; }
tr.total td { border-top: 1px solid #333; border-bottom: 2px solid #333; font-weight: 700; padding: 6px 4px; }
.warn { margin-top: 8px; color: #a33; font-weight: 600; }
.internal { display: inline-block; margin-top: 4px; padding: 2px 8px; border: 1px solid #111; font-size: 10px;
            font-weight: 700; letter-spacing: .06em; }
.internal-stub th { white-space: normal; }
.rule { margin: 0 0 22px; }
.pay .kv .kvrule { grid-column: 1 / -1; border-top: 1px solid #333; margin: 4px 0 3px; }
.pay .kv .lbl.amtl { font-size: 14px; align-self: baseline; }
"""


def _internal_table(s: dict, cols: Sequence[str]) -> str:
    """Internal stub body (owner 2026-10-07: "invoice one row, total of that, ___ then the bills it was applied to"):
    the client invoice(s) this payment came from, one row each with what it took from it, and their total; a rule;
    then the bills it paid, each ONCE, as Open balance - Paid = New balance, and their total."""
    spec = COLUMNS
    word = "check" if s["method"] == "Check" else "payment"
    out = ["<table class='invs'><thead><tr><th>Client invoice</th><th>Invoice date</th><th>Project</th><th>Draw</th>"
           "<th class='r'>Invoice amount</th><th class='r'>Applied</th></tr></thead><tbody>"]
    for g in s.get("groups") or []:
        if g["invoice_no"]:
            out.append(f"<tr class='row'><td>{_esc(g['invoice_no'])}</td><td>{_esc(mdy(g['date']))}</td><td>{_esc(g['project'])}</td><td>{_esc(g['label'])}</td>"
                       f"<td class='r'>{'' if g.get('invoice_total') is None else money(g['invoice_total'])}</td>"
                       f"<td class='r'>{money(g['paid'])}</td></tr>")
        else:
            out.append(f"<tr class='row'><td colspan='2'>{_esc(g['label'])}</td><td>{_esc(g['project'])}</td><td></td><td class='r'>{money(g['paid'])}</td></tr>")
    out.append(f"<tr class='total'><td colspan='5'>Total</td><td class='r'>${money(sum(g['paid'] for g in s.get('groups') or []))}</td></tr>"
               "</tbody></table><div class='rule'></div>")
    out.append("<table><thead><tr>")
    for k in cols:
        label, align, _ = spec[k]
        w = f" style='width:{WIDTHS[k]}px'" if k in WIDTHS else ""
        out.append(f"<th class='{align[0]}{' memo' if k == 'memo' else ''}'{w}>{_esc(label)}</th>")
    out.append("</tr></thead><tbody>")
    for r in s["rows"]:
        out.append("<tr class='row'>")
        for k in cols:
            _, align, get = spec[k]
            out.append(f"<td class='{align[0]}{' memo' if k == 'memo' else ''}'>{_esc(get(r))}</td>")
        out.append("</tr>")
    paid_at = cols.index("amount") if "amount" in cols else len(cols) - 1
    out.append(f"<tr class='total'><td colspan='{paid_at}'>Total paid on this {word}</td><td class='r'>${money(s['rows_total'])}</td>"
               + "<td></td>" * (len(cols) - paid_at - 1) + "</tr></tbody></table>")
    if not s["ties"]:
        out.append(f"<div class='warn'>Bill lines total ${money(s['rows_total'])} but the payment is "
                   f"${money(s['total'])} - check this payment in QuickBooks.</div>")
    return "".join(out)


def stub_html(stubs: Sequence[dict], columns: Optional[Sequence[str]] = None,
              company: str = "", report_date: Optional[dt.date] = None) -> str:
    """The printable page(s): one stub per payment, the payment on top, its bills below."""
    picked = resolve_columns(columns)
    report_date = report_date or dt.date.today()
    parts = ["<!doctype html><html><head><meta charset='utf-8'><title>Bill payment stubs</title>",
             f"<style>{CSS}</style></head><body>"]
    for s in stubs:
        cols = columns_for(s.get("kind") or "vendor", picked)
        parts.append("<section class='stub internal-stub'>" if s.get("kind") == "internal" else "<section class='stub'>")
        parts.append("<div class='hdr'>")
        if company:
            parts.append(f"<p class='co'>{_esc(company)}</p>")
        internal = s.get("kind") == "internal"
        parts.append(f"<p class='rpt'>Bill Payment Stub{' - Internal' if internal else ''}</p>")
        parts.append(f"<p class='asof'>{_esc(mdy(s['txn_date']))}</p>")
        if internal:
            parts.append("<div class='internal'>INTERNAL - NOT FOR THE VENDOR</div>")
        parts.append("</div>")
        # the payment block
        parts.append("<div class='pay'><div class='to'><div class='lbl'>Paid to</div>")
        parts.append(f"<div class='name'>{_esc(s['pay_to'])}</div>")
        if s.get("address"):
            parts.append("<div class='addr'>" + "<br>".join(_esc(a) for a in s["address"]) + "</div>")
        parts.append("</div><div class='kv'>")
        # no Account line - the owner's own stub does not show it (2026-09-22); the value stays in the record
        kv = [("Payment date", mdy(s["txn_date"])),
              ("Payment type", "Joint check" if internal and s.get("joint") else s["method"]),
              ("Check number" if s["method"] == "Check" else "Reference", s["ref"])]
        if internal:
            kv.append(("Client", s.get("client") or ""))
        kv.append(("Amount", "$" + money(s["total"])))
        for lbl, val in kv:
            if not val:
                continue
            cls = "v amt" if lbl == "Amount" else "v"
            if internal and lbl == "Amount":              # owner 2026-10-07: a divider above Amount, its label as big as the $
                parts.append("<div class='kvrule'></div>")
                parts.append(f"<div class='lbl amtl'>{_esc(lbl)}</div><div class='{cls}'>{_esc(val)}</div>")
                continue
            parts.append(f"<div class='lbl'>{_esc(lbl)}</div><div class='{cls}'>{_esc(val)}</div>")
        parts.append("</div></div>")
        if internal:
            parts.append(_internal_table(s, cols))
            parts.append("</section>")
            continue
        # the bills table
        spec = COLUMNS
        parts.append("<table><thead><tr>")
        for k in cols:
            label, align, _ = spec[k]
            w = f" style='width:{WIDTHS[k]}px'" if k in WIDTHS else ""
            parts.append(f"<th class='{align[0]}{' memo' if k == 'memo' else ''}'{w}>{_esc(label)}</th>")
        parts.append("</tr></thead><tbody>")
        parts.append(f"<tr class='band'><td colspan='{len(cols)}'>{_esc(s['vendor'])}</td></tr>")
        for r in s["rows"]:
            parts.append("<tr class='row'>")
            for k in cols:
                _, align, get = spec[k]
                parts.append(f"<td class='{align[0]}{' memo' if k == 'memo' else ''}'>{_esc(get(r))}</td>")
            parts.append("</tr>")
        parts.append("<tr class='total'>")
        first_money = next((i for i, k in enumerate(cols) if k in MONEY_COLUMNS), len(cols))
        for i, k in enumerate(cols):
            if i == 0 and first_money > 0:
                parts.append(f"<td colspan='{first_money}'>Total for {_esc(s['vendor'])}</td>")
            elif i < first_money:
                continue
            elif k == "amount":
                parts.append(f"<td class='r'>${money(s['rows_total'])}</td>")
            elif k == "bill_total":
                parts.append(f"<td class='r'>${money(sum(r['bill_total'] for r in s['rows']))}</td>")
            elif k == "balance_before":
                parts.append(f"<td class='r'>${money(sum(r.get('balance_before') or 0 for r in s['rows']))}</td>")
            elif k == "open_balance":
                parts.append(f"<td class='r'>${money(sum(r.get('balance_after') or 0 for r in s['rows']))}</td>")
            else:
                parts.append("<td></td>")
        parts.append("</tr></tbody></table>")
        if not s["ties"]:
            parts.append(f"<div class='warn'>Bill lines total ${money(s['rows_total'])} but the payment is "
                         f"${money(s['total'])} - check this payment in QuickBooks.</div>")
        parts.append("</section>")
    parts.append("</body></html>")
    return "".join(parts)


def chrome_path() -> Optional[str]:
    for c in CHROME_CANDIDATES:
        if Path(c).exists():
            return c
    return shutil.which("google-chrome") or shutil.which("chromium")


def render_pdf(html_path: Path, pdf_path: Path) -> bool:
    """Print the HTML to PDF with Chrome headless (Letter portrait, no header/footer)."""
    exe = chrome_path()
    if not exe:
        return False
    # a bare, throwaway printer: its own temp profile and no updater / first run / background network. Started by the
    # ledger (Python), Chrome's updater reaching for Google Chrome.app trips macOS App Management ("python3.14 was
    # prevented from modifying apps on your Mac", owner 2026-09-25). With its own profile Chrome writes the PDF and
    # then never exits, so wait for its "bytes written to file" line and close it ourselves.
    try:
        pdf_path.unlink()
    except FileNotFoundError:
        pass
    with tempfile.TemporaryDirectory(prefix="ledger-chrome-") as prof:
        cmd = [exe, "--headless=new", "--disable-gpu", "--no-pdf-header-footer", f"--user-data-dir={prof}",
               "--use-mock-keychain", "--password-store=basic",    # a fresh profile otherwise asks the keychain and hangs
               "--no-first-run", "--no-default-browser-check", "--disable-component-update",
               "--disable-background-networking", "--disable-sync", "--disable-extensions",
               f"--print-to-pdf={pdf_path}", html_path.resolve().as_uri()]
        log = Path(prof) / "chrome.log"
        with open(log, "w") as out:
            try:
                proc = subprocess.Popen(cmd, stdout=out, stderr=subprocess.STDOUT)
            except OSError:
                return False
            try:
                deadline = time.monotonic() + 120
                while time.monotonic() < deadline:
                    if proc.poll() is not None or "written to file" in log.read_text(errors="ignore"):
                        break
                    time.sleep(0.25)
            finally:
                if proc.poll() is None:
                    proc.terminate()
                    try:
                        proc.wait(timeout=10)
                    except subprocess.TimeoutExpired:
                        proc.kill()
                        proc.wait()
    return pdf_path.exists() and pdf_path.stat().st_size > 0


def pdf_page_count(pdf_path: Path) -> int:
    import re
    data = pdf_path.read_bytes()
    return len(re.findall(rb"/Type\s*/Page[^s]", data))


# ───────────────────────── CLI ─────────────────────────

def _safe_name(s: str) -> str:
    return "".join("-" if c in '/\\:*?"<>|' else c for c in s).strip().lstrip(".") or "vendor"


def _file_name(stub: dict) -> str:
    """One payment = one file, named the owner's way (2026-09-22):
    "[Vendor] - Bill Payment Stub [Ref #] [date]" -> COWTOWN REDI MIX CONCRETE - Bill Payment Stub 25760 09-21-2026.pdf
    The vendor is the QBO name (same as the folder); the ref is the check # or the card reference."""
    vend = stub.get("vendor") or "vendor"
    tag = mdy(stub["txn_date"]).replace("/", "-") if stub.get("txn_date") else dt.date.today().strftime("%m-%d-%Y")
    ref = stub["ref"] or f"{stub['method']} {stub['payment_id']}"   # no check # (ACH, card): the QBO id keeps two same-day payments apart
    suffix = " INTERNAL" if stub.get("kind") == "internal" else ""
    return f"{_safe_name(vend)} - Bill Payment Stub {_safe_name(ref)} {tag}{suffix}.pdf"


def default_out(stub: dict) -> Path:
    """Accounting/Accounts Payable/Bill Payment Stubs/<QBO vendor name>/<file>.pdf, one file per
    payment. The share must be mounted: a missing mount raises so a stub never lands elsewhere."""
    base = paths.bill_payment_stubs_dir()
    if paths.get("ACB_BILL_PAYMENT_STUBS_DIR", ""):
        if not base.parent.exists():
            raise FileNotFoundError(f"the stub folder's parent does not exist ({base.parent})")
    else:
        try:
            paths.require_accounting_share("the bill payment stub")
        except SystemExit as e:
            raise FileNotFoundError(str(e)) from None
    folder = base / _safe_name(stub["vendor"])
    if stub.get("kind") == "internal":
        folder = folder / "Internal"                      # kept apart so it is never mailed with the check
    return folder / _file_name(stub)


def _selftest() -> int:
    bills = {
        "b1": {"Id": "b1", "DocNumber": "1001", "TxnDate": "2026-06-17", "DueDate": "2026-07-17",
               "TotalAmt": 1000.0, "Balance": 0, "PrivateNote": "RP1234 - job one",
               "Line": [{"Amount": 1000.0, "ItemBasedExpenseLineDetail": {"CustomerRef": {"name": "GC:RP1234 SL1"}}}]},
        "b2": {"Id": "b2", "DocNumber": "1002", "TxnDate": "2026-06-18", "DueDate": "2026-07-18",
               "TotalAmt": 500.0, "Balance": 200.0, "PrivateNote": "RP1234 - job one (partial)",
               "Line": [{"Amount": 500.0, "ItemBasedExpenseLineDetail": {"CustomerRef": {"name": "GC:RP1234-FTW FW1"}}}]},
    }
    bp = {"Id": "p1", "TxnDate": "2026-09-21", "TotalAmt": 1300.0, "PayType": "Check", "DocNumber": "25760",
          "VendorRef": {"value": "v1", "name": "TEST VENDOR"},
          "CheckPayment": {"BankAccountRef": {"name": "Test Bank ****0000"}},
          "Line": [{"Amount": 1000.0, "LinkedTxn": [{"TxnType": "Bill", "TxnId": "b1"}]},
                   {"Amount": 300.0, "LinkedTxn": [{"TxnType": "Bill", "TxnId": "b2"}]}]}
    vendor = {"DisplayName": "TEST VENDOR", "PrintOnCheckName": "TEST VENDOR INC",
              "BillAddr": {"Line1": "PO BOX 1", "City": "FORT WORTH", "CountrySubDivisionCode": "TX", "PostalCode": "76100"}}
    s = build_stub(bp, lambda t, i: bills.get(i) if t == "Bill" else None, vendor)
    assert s["ties"] and s["rows_total"] == 1300.0, s
    assert [r["number"] for r in s["rows"]] == ["1001", "1002"]
    assert s["rows"][1]["partial"] and not s["rows"][0]["partial"]
    assert s["rows"][1]["project"] == "RP1234-FTW"
    assert s["pay_to"] == "TEST VENDOR INC" and s["address"][-1] == "FORT WORTH TX 76100"
    assert s["method"] == "Check" and s["ref"] == "25760" and s["account"] == "Test Bank ****0000"
    assert resolve_columns() == list(DEFAULT_COLUMNS)
    assert resolve_columns(add=["bill_total"], drop=["type"]) == ["date", "number", "memo", "balance_before", "amount", "open_balance", "bill_total"]
    assert resolve_columns(["memo", "amount"]) == ["memo", "amount"]
    for bad in ({"add": ["nope"]}, {"columns": ["memo", "typo"]}):
        try:
            resolve_columns(**bad)   # type: ignore[arg-type]
            raise AssertionError("unknown column accepted")
        except ValueError:
            pass
    page = stub_html([s], company="Test Co, LLC")
    for needle in ("Test Co, LLC", "Bill Payment Stub", "09/21/2026", "TEST VENDOR INC", "Check number", "25760",
                   "Memo/Description", "Open balance", "New balance", money(1000), "$" + money(1300), "06/17/2026"):
        assert needle in page, needle
    assert "2026-09-21" not in page.replace("Bill Payment Stub", "")   # never year-first on the page
    nomemo = stub_html([s], columns=["date", "number", "amount"])       # memo unticked at print time
    assert "Memo/Description" not in nomemo and "Paid" in nomemo and "New balance" not in nomemo   # exactly the pick
    assert "partial payment" not in page and "Transaction type" in page       # columns, never a sentence
    assert ">500.00</td><td class='r'>300.00</td><td class='r'>200.00<" in page, "partial bill = 500 - 300 = 200"
    tot = "".join(f"<td class='r'>${money(x)}</td>" for x in (1500, 1300, 200))
    assert tot in page                                                        # totals tie: before - paid = new
    assert "Account" not in page and "Test Bank" not in page          # the owner's stub shows no account
    assert _file_name(s) == "TEST VENDOR - Bill Payment Stub 25760 09-21-2026.pdf"
    assert _file_name(dict(s, method="Credit card", ref="AMEX-0000")) == "TEST VENDOR - Bill Payment Stub AMEX-0000 09-21-2026.pdf"
    assert _file_name(dict(s, ref="")) == "TEST VENDOR - Bill Payment Stub Check p1 09-21-2026.pdf"   # no check #: the id keeps same-day payments apart
    # a bill on two lines is ONE row; a vendor credit and an unapplied amount are rows of their own; everything still ties
    bp2 = dict(bp, TotalAmt=1350.0, Line=[{"Amount": 600.0, "LinkedTxn": [{"TxnType": "Bill", "TxnId": "b1"}]},
                                        {"Amount": 400.0, "LinkedTxn": [{"TxnType": "Bill", "TxnId": "b1"}]},
                                        {"Amount": 300.0, "LinkedTxn": [{"TxnType": "VendorCredit", "TxnId": "vc1"}]},
                                        {"Amount": 50.0}])
    s2 = build_stub(bp2, lambda t, i: bills.get(i) if t == "Bill" else ({"DocNumber": "VC-9", "TxnDate": "2026-06-01", "TotalAmt": 300.0} if i == "vc1" else None), vendor)
    assert s2["ties"] and [r["type"] for r in s2["rows"]] == ["Vendor credit", "Bill", "Unapplied"], s2["rows"]
    assert s2["rows"][1]["paid"] == 1000.0 and not s2["rows"][1]["partial"]
    assert "Vendor credit" in stub_html([s2]) and "Unapplied" in stub_html([s2])
    assert not (_natural("1001") < _natural("999")) and _natural("999") < _natural("1001")
    assert _safe_name("A/B: C") == "A-B- C"
    assert "Bill amount" not in page and "Bill amount" in stub_html([s], columns=["number", "bill_total", "amount"])
    assert "check this payment" not in stub_html([s])
    bp_bad = dict(bp, TotalAmt=1400.0)
    assert "check this payment" in stub_html([build_stub(bp_bad, lambda t, i: bills.get(i) if t == "Bill" else None, vendor)])
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "stub.html"
        p.write_text(page, encoding="utf-8")
        assert p.stat().st_size > 1000
        # history: record, judge against a fake mirror state, snapshot round trip
        con = ledger_connect(Path(td) / "ledger.sqlite3")
        hid = record_print(con, s, list(DEFAULT_COLUMNS), p, {"SyncToken": "3", "MetaData": {"LastUpdatedTime": "x"}},
                           archive_dir=Path(td) / "archive")
        snap = history_snapshot(con, hid)
        assert snap and snap["stub"]["ref"] == "25760" and snap["columns"] == list(DEFAULT_COLUMNS)
        assert snap["served_path"] == str(Path(td) / "archive" / f"{hid}.pdf") and Path(snap["served_path"]).read_bytes() == p.read_bytes()
        p.write_text("overwritten by a re-print", encoding="utf-8")            # the vendor file moves on, the archive does not
        assert Path(snap["served_path"]).read_text(encoding="utf-8") == page
        Path(snap["served_path"]).unlink()                                     # its own copy gone -> nothing served, never the newer print
        assert history_snapshot(con, hid)["served_path"] is None
        assert history(con)[0]["status"] in ("unchecked", "current", "changed", "deleted", "voided") and "checked" in history(con)[0]
        row = {"total": 1300.0, "sync_token": "3"}
        assert print_status(row, {"present": True, "deleted": False, "voided": False, "sync_token": "3", "total": 1300.0}) == "current"
        assert print_status(row, {"present": True, "deleted": False, "voided": False, "sync_token": "4", "total": 1300.0}) == "changed"
        assert print_status(row, {"present": True, "deleted": False, "voided": True, "sync_token": "4", "total": 0.0}) == "voided"
        assert print_status(row, {"present": True, "deleted": True, "voided": False, "sync_token": "3", "total": 1300.0}) == "deleted"
        assert print_status(row, {"present": False}) == "deleted"
        assert is_voided({"TotalAmt": 0, "PrivateNote": "Voided", "Line": []}) and is_voided({"TotalAmt": 0, "Line": []}) and not is_voided(bp)
        assert print_status({"total": 1.0, "sync_token": "0"}, {"present": True, "deleted": False, "voided": False, "sync_token": 0, "total": 1.0}) == "current"   # token 0 is not "missing"
        con.close()
        assert [c["key"] for c in columns_registry() if c["default"]] == list(DEFAULT_COLUMNS)
    # the internal stub: invoice on top, bills below, Open balance - Paid = New balance; joint check from the client's Payment
    assert columns_for("vendor") == list(DEFAULT_COLUMNS)
    assert columns_for("vendor", ["open_balance", "date", "amount", "number"]) == ["date", "number", "amount", "open_balance"]
    assert [c["label"] for c in columns_registry()][-3:] == ["Open balance", "Paid", "New balance"]   # picker = stub
    assert columns_for("vendor", columns_for("vendor")) == columns_for("vendor")
    assert columns_for("internal") == ["date", "number", "memo", "project", "balance_before", "amount", "open_balance"]
    bills["b2"]["LinkedTxn"] = [{"TxnId": "p0", "TxnType": "BillPaymentCheck"}, {"TxnId": "pj", "TxnType": "BillPaymentCheck"}]
    pays = {"p0": {"Id": "p0", "TxnDate": "2026-06-30", "Line": [{"Amount": 100.0, "LinkedTxn": [{"TxnType": "Bill", "TxnId": "b2"}]}]}}
    invs = {"i1": {"Id": "i1", "DocNumber": "34457", "TxnDate": "2026-05-20", "TotalAmt": 600.0, "PrivateNote": "MFD1 - Job - May Draw 2026\n(Period)"},
            "i2": {"Id": "i2", "DocNumber": "34458", "TxnDate": "2026-06-20", "TotalAmt": 900.0, "PrivateNote": "MFD1 - Job - June Draw 2026"}}
    getj = lambda t, i: bills.get(i) if t == "Bill" else pays.get(i) if t == "BillPayment" else invs.get(i) if t == "Invoice" else None   # noqa: E731
    bpj = dict(bp, Id="pj", CheckPayment={"BankAccountRef": {"name": "Joint Checks Account"}})
    sj = dict(build_stub(bpj, getj, vendor), kind="internal")
    assert sj["method"] == "Check" and [r["balance_before"] for r in sj["rows"]] == [1000.0, 400.0]   # b2: 500 less p0's 100 earlier
    cpay = {"Id": "c1", "TxnDate": bp["TxnDate"], "TotalAmt": 1300.0, "PaymentRefNum": "25760", "CustomerRef": {"value": "proj"},
            "Line": [{"Amount": 600.0, "LinkedTxn": [{"TxnType": "Invoice", "TxnId": "i1"}]}, {"Amount": 700.0, "LinkedTxn": [{"TxnType": "Invoice", "TxnId": "i2"}]}]}
    assert find_joint_payment(bpj, lambda d: [cpay, dict(cpay, Id="c2", TotalAmt=5.0)]) is cpay
    assert find_joint_payment(bp, lambda d: [cpay]) is None                       # not on the joint account
    apply_internal(sj, cpay, getj, lambda d: None, lambda c: "GC ONE")
    assert sj["joint"] and sj["client"] == "GC ONE"
    assert [(g["invoice_no"], g["label"], g["paid"]) for g in sj["groups"]] == [("34457", "May Draw 2026", 600.0), ("34458", "June Draw 2026", 700.0)]
    seg = [(r["number"], r["balance_before"], r["paid"], r["balance_after"]) for g in sj["groups"] for r in g["rows"]]
    assert seg == [("1001", 1000.0, 600.0, 400.0), ("1001", 400.0, 400.0, 0.0), ("1002", 400.0, 300.0, 100.0)], seg
    ip = stub_html([sj], columns_for("internal"))
    for needle in ("INTERNAL - NOT FOR THE VENDOR", "Joint check", "GC ONE", "34457", "05/20/2026", "May Draw 2026",
                   "Client invoice", "New balance", "Total paid on this check"):
        assert needle in ip, needle
    assert "partial payment" not in ip and "Transaction type" not in ip
    vp = stub_html([dict(sj, kind="vendor")])                                      # the vendor stub: no client, no invoices
    assert "Joint check" not in vp and "partial payment" not in vp and "New balance" in vp and "INTERNAL" not in vp
    assert "Client invoice" not in vp
    sn = apply_internal(dict(build_stub(bp, getj, vendor), kind="internal"), None, getj, lambda d: invs["i2"] if d == "34458" else None,
                        lambda c: "", {"b1": {"invoice_no": "34458", "clients": ["GC ONE"]}})
    assert [(g["invoice_no"] or g["label"], g["paid"]) for g in sn["groups"]] == [("34458", 1000.0), ("No client invoice yet", 300.0)]
    assert _file_name(sj) == "TEST VENDOR - Bill Payment Stub 25760 09-21-2026 INTERNAL.pdf"
    print("selftest OK: build_stub, column registry, page, history")
    return 0


def main(argv: Optional[Sequence[str]] = None) -> int:
    ap = argparse.ArgumentParser(description="Bill payment stubs (payment on top, bills below) from the QBO mirror.")
    ap.add_argument("--vendor", default="", help="vendor name substring (case-insensitive)")
    ap.add_argument("--date", default="", help="one payment date, YYYY-MM-DD")
    ap.add_argument("--from", dest="date_from", default="")
    ap.add_argument("--to", dest="date_to", default="")
    ap.add_argument("--payment", action="append", default=[], help="a BillPayment id (repeatable)")
    ap.add_argument("--columns", default="", help="comma list: the exact columns, in order")
    ap.add_argument("--add", default="", help="comma list of columns to add to the default set")
    ap.add_argument("--drop", default="", help="comma list of columns to drop from the default set")
    ap.add_argument("--list-columns", action="store_true")
    ap.add_argument("--company", default=paths.get("ACB_COMPANY_NAME", ""))
    ap.add_argument("--out", default="", help="force everything into ONE PDF at this path (default: one file per payment in the vendor folder)")
    ap.add_argument("--html", action="store_true", help="keep the .html beside the .pdf")
    ap.add_argument("--history", action="store_true", help="list every printed stub (for --vendor / --payment) with its status vs QBO now")
    ap.add_argument("--selftest", action="store_true")
    a = ap.parse_args(argv)

    if a.selftest:
        return _selftest()
    if a.list_columns:
        for k, (label, align, _) in COLUMNS.items():
            print(f"  {k:13s} {label:20s} {'(default)' if k in DEFAULT_COLUMNS else ''}")
        return 0
    if a.history:
        con = ledger_connect()
        try:
            rows = history(con, "", a.payment[0] if a.payment else "")
        finally:
            con.close()
        if a.vendor and not a.payment:
            rows = [r for r in rows if a.vendor.lower() in (r["vendor"] or "").lower()]
        for r in rows:
            print(f"  {r['printed_at'][:16]}  {r['status']:8s} {mdy(r['txn_date'])}  {r['method']:11s} {r['ref']:12s} "
                  f"{(r['vendor'] or '')[:30]:30s} ${money(r['total'])}  {'' if r['file_exists'] else '(file missing) '}{r['file_path']}")
        if not rows:
            print("no printed stubs yet")
        return 0
    split = lambda s: [x.strip() for x in s.split(",") if x.strip()]   # noqa: E731
    try:
        cols = resolve_columns(split(a.columns) or None, split(a.add), split(a.drop))
    except ValueError as e:
        print(f"error: {e}", file=sys.stderr)
        return 2
    if not (a.payment or a.vendor or a.date or a.date_from):
        print("error: give --payment, or --vendor with --date / --from / --to", file=sys.stderr)
        return 2
    d_from, d_to = (a.date, a.date) if a.date else (a.date_from, a.date_to)
    stubs = load_payments(a.vendor, d_from, d_to, a.payment)
    if not stubs:
        print("no bill payments matched", file=sys.stderr)
        return 1
    if not a.company:
        print("note: no company name on the stub (set ACB_COMPANY_NAME in machine.env or pass --company)")
    for s in stubs:
        flag = "" if s["ties"] else "   <- lines do not tie to the payment"
        print(f"  {mdy(s['txn_date'])}  {s['method']:11s} {s['ref']:12s} {s['vendor'][:34]:34s} "
              f"{len(s['rows']):3d} bills  ${money(s['total'])}{flag}")
    # one PDF per PAYMENT, in the vendor's folder; --out forces everything into a single file
    groups: List[List[dict]] = [stubs] if a.out else [[s] for s in stubs]
    rc = 0
    for group in groups:
        try:
            out = Path(a.out).expanduser() if a.out else default_out(group[0])
            if a.out and out.suffix.lower() != ".pdf":
                out = out.with_suffix(".pdf")           # the .html beside it is the source; --out names the PDF
        except FileNotFoundError as e:
            print(f"error: {e}", file=sys.stderr)
            return 1
        out.parent.mkdir(parents=True, exist_ok=True)
        html_path = out.with_suffix(".html")
        html_path.write_text(stub_html(group, cols, a.company), encoding="utf-8")
        if render_pdf(html_path, out):
            print(f"wrote {out}  ({pdf_page_count(out)} page(s), columns: {', '.join(cols)})")
            if not a.html:
                html_path.unlink(missing_ok=True)
            if not a.out:                       # a real print -> the history (one row per payment)
                lcon = ledger_connect()
                try:
                    for st in group:
                        record_print(lcon, st, cols, out, st.get("_raw"))
                finally:
                    lcon.close()
        else:
            print(f"Chrome not found or the print failed - the HTML is at {html_path}", file=sys.stderr)
            rc = 1
    return rc


if __name__ == "__main__":
    sys.exit(main())
