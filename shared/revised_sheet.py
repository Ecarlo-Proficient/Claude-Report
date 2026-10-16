"""shared/revised_sheet.py - the "Revised Bills" sheet in every vendor's statement workbook (10/12/2026).

The owner: "I need a way to have the revised bills in a process so we know where they stand - did the
clerk send the bill to get revised? when? have they responded? with the new revised bill? is the new
revised bill uploaded to QBO?" Then: "keep the process in the Excel ... the revised bill sheet needs to be
in the same statement Excel, not a mix ... the bill info with our status on the right ... when the stage
goes to Done, move those bills in the Summary up and keep the Revised Bills sheet as the history."

A revised bill = a bill the vendor must re-issue: the PM changed its total in QuickBooks (the
`REVISED - <reason>` memo line, AP-13), or the vendor's statement shows an amount QuickBooks does not
(the Summary's amount-mismatch and sales-tax rows). The re-issued bill must land on the SAME QuickBooks
bill - a second entry is the pickle.

The sheet (`Reconciliation - <Vendor> - as of <date>.xlsx` › `Revised Bills`, one vendor per workbook,
an Excel Table), two parts and nothing else:

  LEFT, the bill info the clerk copies into the email (select A:E, copy, paste):
      Invoice # · Invoice date · Amount billed · Corrected amount · Reason
  RIGHT, our status, typed by the clerk and carried forward on every run:
      Stage (a dropdown) · Sent on · Replied on · Received on · Entered on · Notes
  and, hidden: _Key = the QuickBooks bill id (the carry-forward key), _Since = the day the row first appeared.

  A SECOND section under it (owner 10/12/2026): the bills the vendor's statement still shows open that
  QuickBooks paid - the bill info + the check (payment) #, date and amount, for the vendor to apply:
      Invoice # · Invoice date · Amount · Check # · Check date · Paid amount  |  Notes · _Key

The reconciler (statement-reconciler/statement_reconciler.py) builds the rows each run, reads the old
workbook's sheet first (`harvest`) so nothing typed is lost, keeps every row ever on the sheet (the
history), moves Stage forward only - to "Done - statement agrees" when the vendor's statement shows the
corrected amount - and writes the sheet (`write_sheet`). The ledger reads the rows from the vendor's
statement record (shared/statement_record), never the workbook. Kept in shared/ so both tools read the
same contract (repo rule).
"""
from __future__ import annotations

import datetime as dt
import re
from pathlib import Path
from typing import Dict, List, Optional

SHEET = "Revised Bills"
TABLE = "RevisedBills"
BILL_COLS = [("Invoice #", 14), ("Invoice date", 12), ("Amount billed", 14), ("Corrected amount", 16), ("Reason", 48)]
STATUS_COLS = [("Stage", 24), ("Sent on", 11), ("Replied on", 11), ("Received on", 11), ("Entered on", 11), ("Notes", 40)]
KEY = "_Key"
SINCE = "_Since"                       # the day the row first appeared on the sheet (hidden) - "where it stands" counts from it
HEADERS = [h for h, _ in BILL_COLS + STATUS_COLS] + [KEY, SINCE]
WIDTHS = [w for _, w in BILL_COLS + STATUS_COLS] + [9, 11]
TYPED = ["Stage", "Sent on", "Replied on", "Received on", "Entered on", "Notes"]   # the clerk's - carried, never overwritten
DATE_COLS = {"Invoice date", "Sent on", "Replied on", "Received on", "Entered on", SINCE}
MONEY_COLS = {"Amount billed", "Corrected amount"}

PAID_TITLE = "Paid in QuickBooks - the statement still shows these open"
PAID_COLS = [("Invoice #", 14), ("Invoice date", 12), ("Amount", 14), ("Check #", 14), ("Check date", 12), ("Paid amount", 14), ("Notes", 40)]
PAID_HEADERS = [h for h, _ in PAID_COLS] + [KEY]
PAID_TABLE = "RevisedBillsPaid"
PAID_DATE_COLS = {"Invoice date", "Check date"}
PAID_MONEY_COLS = {"Amount", "Paid amount"}

STAGES = ["To send", "Sent to vendor", "Vendor replied", "Revised bill received", "Entered in QBO",
          "Done - statement agrees", "Dropped"]
STAGE_RANK = {s: i for i, s in enumerate(STAGES)}
FINAL = {"Done - statement agrees", "Dropped"}
STAGE_LIST = ",".join(STAGES)          # the dropdown's inline list (no commas inside a stage, under 255 characters)
STAGE_FILL = {"To send": "F4B6B6", "Sent to vendor": "F9D29B", "Vendor replied": "FBE8A0", "Revised bill received": "BFD7F2",
              "Entered in QBO": "C9E8C9", "Done - statement agrees": "8FD19E", "Dropped": "D9D9D9"}
# the question each stage answers, in the owner's words - "where it stands" names the NEXT one
NEXT_QUESTION = {
    "To send": "the clerk has not told the vendor yet",
    "Sent to vendor": "waiting on the vendor's reply",
    "Vendor replied": "waiting on the re-issued bill",
    "Revised bill received": "the re-issued bill is not on the QuickBooks bill yet",
    "Entered in QBO": "waiting for a vendor statement that shows the corrected amount",
    "Done - statement agrees": "done",
    "Dropped": "dropped",
}
QBO_BILL_URL = "https://qbo.intuit.com/app/bill?txnId={bill_id}"


def col_key(name: str) -> str:
    """'Sent on' -> 'sent_on', '_Key' -> 'key', 'Invoice #' -> 'invoice'."""
    return re.sub(r"[^a-z0-9]+", "_", name.lower().replace("#", "")).strip("_")


def stage_of(text) -> str:
    """The clerk's typed stage, matched loosely (case, spacing, a first word); '' when blank or unknown."""
    t = re.sub(r"\s+", " ", str(text or "")).strip().lower()
    if not t:
        return ""
    for s in STAGES:
        if t == s.lower() or t == s.lower().split(" - ")[0]:
            return s
    for s in STAGES:                                   # 'sent', 'replied', 'received', 'entered', 'done'
        if t.split()[0] in s.lower().split():
            return s
    return ""


def iso(v) -> str:
    """A cell value (date, datetime, 'mm/dd/yyyy', ISO) -> ISO date or ''."""
    if v is None or v == "":
        return ""
    if isinstance(v, dt.datetime):
        return v.date().isoformat()
    if isinstance(v, dt.date):
        return v.isoformat()
    s = str(v).strip()
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{2,4})$", s)
    if m:
        y = int(m.group(3)); y = y + 2000 if y < 100 else y
        try:
            return dt.date(y, int(m.group(1)), int(m.group(2))).isoformat()
        except ValueError:
            return ""
    return s[:10] if re.match(r"^\d{4}-\d{2}-\d{2}", s) else ""


def _date(v):
    try:
        return dt.date.fromisoformat(str(v)[:10]) if v else None
    except ValueError:
        return None


# ─────────────────────────── where it stands ───────────────────────────

def days_since(iso_date: str, today: Optional[dt.date] = None) -> Optional[int]:
    d = _date(iso_date)
    return None if d is None else ((today or dt.date.today()) - d).days


def standing(row: dict, today: Optional[dt.date] = None) -> dict:
    """One plain sentence + how long it has sat there, from the stage and the clerk's dates (revised_on = the day
    the PM changed the bill, else the day the row first appeared on the sheet)."""
    stage = stage_of(row.get("stage")) or "To send"
    since = {"To send": row.get("revised_on"), "Sent to vendor": row.get("sent_on"), "Vendor replied": row.get("replied_on"),
             "Revised bill received": row.get("received_on"), "Entered in QBO": row.get("entered_on")}.get(stage) or row.get("revised_on")
    days = days_since(since or "", today)
    what = NEXT_QUESTION.get(stage, stage)
    text = what if stage in FINAL else (f"{what} - {days} day{'s' if days != 1 else ''}" if days is not None else what)
    return {"stage": stage, "rank": STAGE_RANK.get(stage, 0), "since": since or "", "days": days, "text": text, "final": stage in FINAL}


# ─────────────────────────── the old sheet, read back ───────────────────────────

def harvest(xlsx: Path) -> Dict[str, dict]:
    """Every row of the workbook's Revised Bills sheet, keyed by the QuickBooks bill id: the bill
    columns as written and the clerk's columns as typed (ISO dates, normalised stage). {} when the
    workbook has no such sheet yet. The history: a row read here is never dropped."""
    try:
        from openpyxl import load_workbook
        wb = load_workbook(xlsx, read_only=True, data_only=True)
    except Exception:                                  # noqa: BLE001 - no file / locked = nothing yet
        return {}
    try:
        if SHEET not in wb.sheetnames:
            return {}
        rows = wb[SHEET].iter_rows(values_only=True)
        hdr = [str(h or "").strip() for h in next(rows, [])]
        if KEY not in hdr:
            return {}
        out: Dict[str, dict] = {}
        for vals in rows:
            if vals and vals[0] == "Invoice #":          # the second section's header: the paid rows start - not this reader's
                break
            d: dict = {}
            for h, v in zip(hdr, vals):
                k = col_key(h)
                if not k:
                    continue
                if h in DATE_COLS:
                    d[k] = iso(v)
                elif h in MONEY_COLS:
                    try:
                        d[k] = round(float(v), 2) if v not in (None, "") else None
                    except (TypeError, ValueError):
                        d[k] = None
                else:
                    d[k] = str(v).strip() if v is not None else ""
            key = str(d.get("key") or "").strip()
            if not key:
                continue
            d["key"] = key
            d["stage"] = stage_of(d.get("stage"))
            out[key] = d
        return out
    finally:
        wb.close()


def harvest_paid(xlsx: Path) -> Dict[str, str]:
    """The second section's Notes, by bill id (the only column the clerk types there)."""
    try:
        from openpyxl import load_workbook
        wb = load_workbook(xlsx, read_only=True, data_only=True)
    except Exception:                                  # noqa: BLE001
        return {}
    try:
        if SHEET not in wb.sheetnames:
            return {}
        out: Dict[str, str] = {}
        hdr = None
        for vals in wb[SHEET].iter_rows(values_only=True):
            if vals and vals[0] == "Invoice #" and "Check #" in [str(v or "") for v in vals]:
                hdr = [str(h or "").strip() for h in vals]
                continue
            if hdr is None:
                continue
            d = dict(zip(hdr, vals))
            key = str(d.get(KEY) or "").strip()
            if key:
                out[key] = str(d.get("Notes") or "").strip()
        return out
    finally:
        wb.close()


def payments_for(bill_ids: List[str], con=None, since: str = "") -> Dict[str, List[dict]]:
    """{bill id: [{check_no, check_date, amount, pay_type}]} - every QuickBooks bill payment that paid
    (part of) the bill, from the QuickBooks copy. `since` (ISO) bounds the scan by payment date.
    Best-effort - no mirror / no key = {}."""
    out: Dict[str, List[dict]] = {}
    want = {str(b) for b in bill_ids if b}
    if not want:
        return out
    own = con is None
    if own:
        try:
            from shared import qbo_mirror
            if not qbo_mirror.db_path().exists():
                return out
            con = qbo_mirror.connect()
        except Exception:                              # noqa: BLE001
            return out
    try:
        from shared import qbo_mirror
        q = "SELECT id, txn_date, doc_number, total, json FROM qbo_billpayment WHERE deleted=0"
        args: list = []
        if since:
            q += " AND txn_date >= ?"
            args.append(since)
        for r in con.execute(q, args):
            try:
                j = qbo_mirror._unpack(r[4]) or {}
            except Exception:                          # noqa: BLE001
                continue
            for ln in j.get("Line") or []:
                for lt in ln.get("LinkedTxn") or []:
                    if lt.get("TxnType") == "Bill" and str(lt.get("TxnId")) in want:
                        out.setdefault(str(lt.get("TxnId")), []).append({
                            "check_no": r[2] or "", "check_date": r[1] or "", "amount": round(float(ln.get("Amount") or 0), 2),
                            "pay_type": j.get("PayType") or "", "payment_id": str(r[0])})
    except Exception:                                  # noqa: BLE001
        return {}
    finally:
        if own:
            con.close()
    for v in out.values():
        v.sort(key=lambda x: x["check_date"])
    return out


# ─────────────────────────── the vendor's original total (the change log) ───────────────────────────

def original_totals(bill_ids: List[str]) -> Dict[str, dict]:
    """{bill id: {original, revised_on}} from the QuickBooks copy's change log: the total before the
    LAST 'amount changed' edit and that edit's date. Best-effort - no mirror / no key = {}."""
    out: Dict[str, dict] = {}
    if not bill_ids:
        return out
    try:
        from shared import qbo_mirror
        if not qbo_mirror.db_path().exists():
            return out
        con = qbo_mirror.connect()
    except Exception:                                  # noqa: BLE001
        return out
    try:
        ids = [str(b) for b in bill_ids]
        for i in range(0, len(ids), 500):
            chunk = ids[i:i + 500]
            for c in con.execute(f"SELECT rec_id, kind, changed_at, total_before, total_after, flags FROM mirror_change "
                                 f"WHERE entity='Bill' AND rec_id IN ({','.join('?' * len(chunk))}) ORDER BY changed_at", chunk):
                e = out.setdefault(str(c[0]), {"original": None, "revised_on": "", "first": None})
                if c[1] == "created" and e["first"] is None:
                    e["first"] = c[4]
                if c[1] == "edited" and "amount changed" in (c[5] or ""):
                    if e["original"] is None:
                        e["original"] = c[3]
                    e["revised_on"] = (c[2] or "")[:10]
        for e in out.values():
            if e["original"] is None:
                e["original"] = e.pop("first", None)
            else:
                e.pop("first", None)
    except Exception:                                  # noqa: BLE001
        return {}
    finally:
        con.close()
    return out


# ─────────────────────────── the sheet ───────────────────────────

def write_sheet(wb, rows: List[dict], paid: Optional[List[dict]] = None, index: int = 1) -> int:
    """Write the sheet (replacing any older one) at `index`: two parts, one Table, the Stage dropdown,
    the colour that says the stage. `rows` = [{bill_id, invoice, invoice_date, amount_billed,
    corrected_amount, reason, stage, sent_on, replied_on, received_on, entered_on, notes}] (ISO dates).
    `paid` = the second section: [{bill_id, invoice, invoice_date, amount, check_no, check_date, paid_amount, notes}]."""
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.worksheet.table import Table, TableStyleInfo
    from shared.xlsx_guard import put as xl_text
    if SHEET in wb.sheetnames:
        del wb[SHEET]
    ws = wb.create_sheet(SHEET, index)
    ws.sheet_view.showGridLines = False
    n_bill = len(BILL_COLS)
    for i, (h, w) in enumerate(zip(HEADERS, WIDTHS), 1):
        c = ws.cell(row=1, column=i, value=h)
        c.fill = PatternFill("solid", fgColor="1F5C99" if i <= n_bill else "595959")
        c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "A2"
    for h in (KEY, SINCE):                                                    # the carry-forward key and the first-seen day: not for the eye
        ws.column_dimensions[get_column_letter(HEADERS.index(h) + 1)].hidden = True
    thin = Side(style="thin", color="9DB6D1")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    r_i = 2
    for row in rows:
        vals = [row.get("invoice", ""), _date(row.get("invoice_date")), row.get("amount_billed"), row.get("corrected_amount"), row.get("reason", ""),
                stage_of(row.get("stage")) or "To send", _date(row.get("sent_on")), _date(row.get("replied_on")), _date(row.get("received_on")),
                _date(row.get("entered_on")), row.get("notes", ""), str(row.get("bill_id") or ""), _date(row.get("revised_on"))]
        for i, (h, v) in enumerate(zip(HEADERS, vals), 1):
            if h in DATE_COLS:
                c = ws.cell(row=r_i, column=i, value=v)
                c.number_format = "mm/dd/yyyy"
            elif h in MONEY_COLS:
                c = ws.cell(row=r_i, column=i, value=round(float(v), 2) if v not in (None, "") else None)
                c.number_format = '"$"#,##0.00'
            else:
                c = xl_text(ws, r_i, i, v if v is not None else "")      # vendor / QuickBooks text: text, never a formula
                if h in ("Reason", "Notes"):
                    c.alignment = Alignment(wrap_text=True, vertical="top")
            if i <= n_bill:
                c.border = box
        r_i += 1
    last = max(r_i - 1, 2)
    if not rows:
        ws.cell(row=2, column=1, value="(nothing to re-issue)")
    tbl = Table(displayName=TABLE, ref=f"A1:{get_column_letter(len(HEADERS))}{last}")
    tbl.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showFirstColumn=False, showLastColumn=False,
                                        showRowStripes=False, showColumnStripes=False)
    ws.add_table(tbl)
    S = get_column_letter(HEADERS.index("Stage") + 1)
    L = max(last, 200)
    dv = DataValidation(type="list", formula1=f'"{STAGE_LIST}"', allow_blank=True, showErrorMessage=False)
    dv.add(f"{S}2:{S}{L}")
    ws.add_data_validation(dv)
    for st, fc in STAGE_FILL.items():
        ws.conditional_formatting.add(f"{S}2:{S}{L}", FormulaRule(formula=[f'${S}2="{st}"'], fill=PatternFill("solid", fgColor=fc, bgColor=fc), font=Font(bold=True)))
    D, B = get_column_letter(HEADERS.index("Corrected amount") + 1), get_column_letter(HEADERS.index("Amount billed") + 1)
    ws.conditional_formatting.add(f"{D}2:{D}{L}", FormulaRule(formula=[f'AND(ISNUMBER(${D}2),${D}2<>${B}2)'], font=Font(bold=True, color="1F5C99")))
    # ── the second section: paid in QuickBooks, the statement still shows it open ──
    paid = paid or []
    t = last + 3
    ws.cell(row=t, column=1, value=PAID_TITLE).font = Font(name="Calibri", size=12, bold=True, color="1F5C99")
    h = t + 1
    n_v = len(PAID_COLS) - 1                                 # Invoice # .. Paid amount go to the vendor; Notes is ours
    for i, name in enumerate(PAID_HEADERS, 1):
        c = ws.cell(row=h, column=i, value=name)
        c.fill = PatternFill("solid", fgColor="1F5C99" if i <= n_v else "595959")
        c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[h].height = 30
    r_i = h + 1
    for row in paid:
        vals = [row.get("invoice", ""), _date(row.get("invoice_date")), row.get("amount"), row.get("check_no", ""), _date(row.get("check_date")),
                row.get("paid_amount"), row.get("notes", ""), str(row.get("bill_id") or "")]
        for i, (name, v) in enumerate(zip(PAID_HEADERS, vals), 1):
            if name in PAID_DATE_COLS:
                c = ws.cell(row=r_i, column=i, value=v)
                c.number_format = "mm/dd/yyyy"
            elif name in PAID_MONEY_COLS:
                c = ws.cell(row=r_i, column=i, value=round(float(v), 2) if v not in (None, "") else None)
                c.number_format = '"$"#,##0.00'
            else:
                c = xl_text(ws, r_i, i, v if v is not None else "")
                if name == "Notes":
                    c.alignment = Alignment(wrap_text=True, vertical="top")
            if i <= n_v:
                c.border = box
        r_i += 1
    plast = max(r_i - 1, h + 1)
    if not paid:
        ws.cell(row=h + 1, column=1, value="(none - every paid bill has dropped off the statement)")
    tbl2 = Table(displayName=PAID_TABLE, ref=f"A{h}:{get_column_letter(len(PAID_HEADERS))}{plast}")
    tbl2.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showFirstColumn=False, showLastColumn=False,
                                         showRowStripes=False, showColumnStripes=False)
    ws.add_table(tbl2)
    return len(rows)
