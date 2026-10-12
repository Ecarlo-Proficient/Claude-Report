"""shared/revised_sheet.py - the Bill Tracker's "Revised Bills" sheet: ONE contract, two tools (10/12/2026).

The owner: "I need a way to have the revised bills in a process so we know where they stand - did
the clerk send the bill to get revised? when? have they responded? with the new revised bill? is the
new revised bill uploaded to QBO? that's the part where we can get in a pickle." And, the same day:
"keep the process in the Excel, forget about Notion. It's only the Excel."

A revised bill = a bill a PM fixed in QuickBooks where the TOTAL changed, marked by the
`REVISED - <reason>` memo line (AP-13 in the process registry). The vendor must re-issue it, and the
re-issued bill must land on the SAME QuickBooks bill - a second entry is the pickle.

The sheet (`Bill Tracker.xlsx` › `Revised Bills`, an Excel Table) has one row per such bill:

  the AP run writes     Vendor · Bill # · Bill date · Project · PM reason · Vendor billed · QBO now ·
  (bill-tracker/        Open · Approval · Revised in QBO · Statement says · Flags · Where it stands ·
   revised_bills.py)    QBO (link) · _Key (the QuickBooks bill id - the carry-forward key)
  the clerk types       Stage (a dropdown) · Sent on · Replied on · Received on · Entered on · Notes
                        - read back from the sheet before every rebuild, like the Lien tags and Notes
  the run moves Stage   forward only, and only to "Done - statement agrees" (the vendor's own statement
                        shows the revised amount) or "Dropped" (the bill left QuickBooks)

The ledger (ledger/dashboard.py, /api/revisions) READS the sheet through `read_sheet` and shows where
every revised bill stands; it never writes the workbook. Kept in shared/ because both tools need the
same column names and the same reading of a row (repo rule).
"""
from __future__ import annotations

import datetime as dt
import re
from pathlib import Path
from typing import Dict, List, Optional

SHEET = "Revised Bills"
TABLE = "RevisedBills"
HEADERS = ["Vendor", "Bill #", "Bill date", "Project", "PM reason", "Vendor billed", "QBO now", "Open", "Approval",
           "Revised in QBO", "Statement says", "Flags", "Where it stands",
           "Stage", "Sent on", "Replied on", "Received on", "Entered on", "Notes", "QBO", "_Key"]
KINDS = ["text", "text", "date", "text", "text", "money", "money", "money", "text",
         "date", "text", "text", "text",
         "text", "date", "date", "date", "date", "text", "url", "text"]
WIDTHS = [30, 14, 11, 11, 40, 13, 13, 12, 12, 12, 40, 40, 36, 22, 11, 11, 11, 11, 36, 6, 10]
TYPED = ["Stage", "Sent on", "Replied on", "Received on", "Entered on", "Notes"]   # the clerk's - carried forward, never overwritten
KEY = "_Key"

STAGES = ["To send", "Sent to vendor", "Vendor replied", "Revised bill received", "Entered in QBO",
          "Done - statement agrees", "Dropped"]
STAGE_RANK = {s: i for i, s in enumerate(STAGES)}
FINAL = {"Done - statement agrees", "Dropped"}
STAGE_LIST = ",".join(STAGES)          # the dropdown's inline list (no commas inside a stage, under 255 characters)
# the question each stage answers, in the owner's words - "where it stands" names the NEXT one
NEXT_QUESTION = {
    "To send": "the clerk has not told the vendor yet",
    "Sent to vendor": "waiting on the vendor's reply",
    "Vendor replied": "waiting on the re-issued bill",
    "Revised bill received": "the re-issued bill is not on the QuickBooks bill yet",
    "Entered in QBO": "waiting for a vendor statement that shows the revised amount",
    "Done - statement agrees": "done",
    "Dropped": "dropped",
}
QBO_BILL_URL = "https://qbo.intuit.com/app/bill?txnId={bill_id}"


def col_key(name: str) -> str:
    """'Sent on' -> 'sent_on', '_Key' -> 'key', 'Bill #' -> 'bill'."""
    return re.sub(r"[^a-z0-9]+", "_", name.lower().replace("#", "")).strip("_")


def stage_of(text) -> str:
    """The clerk's typed stage, matched loosely (case, spacing); '' when blank or unknown."""
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


def _iso(v) -> str:
    if v is None or v == "":
        return ""
    if isinstance(v, dt.datetime):
        return v.date().isoformat()
    if isinstance(v, dt.date):
        return v.isoformat()
    s = str(v).strip()
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{2,4})$", s)          # a typed mm/dd/yyyy
    if m:
        y = int(m.group(3)); y = y + 2000 if y < 100 else y
        try:
            return dt.date(y, int(m.group(1)), int(m.group(2))).isoformat()
        except ValueError:
            return ""
    return s[:10] if re.match(r"^\d{4}-\d{2}-\d{2}", s) else ""


def read_sheet(path: Path) -> List[dict]:
    """Every row of the sheet as a flat dict keyed by col_key(header) (dates ISO, money float,
    stage normalised). [] when the workbook or the sheet does not exist yet."""
    try:
        from openpyxl import load_workbook
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception:                                  # noqa: BLE001 - no file / locked = nothing yet
        return []
    try:
        if SHEET not in wb.sheetnames:
            return []
        ws = wb[SHEET]
        rows = ws.iter_rows(values_only=True)
        hdr = [str(h or "").strip() for h in next(rows, [])]
        if KEY not in hdr:
            return []
        out = []
        for vals in rows:
            d: dict = {}
            for h, v in zip(hdr, vals):
                k = col_key(h)
                if not k:
                    continue
                kind = KINDS[HEADERS.index(h)] if h in HEADERS else "text"
                if kind == "date":
                    d[k] = _iso(v)
                elif kind == "money":
                    try:
                        d[k] = round(float(v), 2) if v not in (None, "") else None
                    except (TypeError, ValueError):
                        d[k] = None
                elif kind == "url":
                    d[k] = str(v or "")
                else:
                    d[k] = str(v).strip() if v is not None else ""
            if not d.get("key"):
                continue
            d["stage"] = stage_of(d.get("stage"))
            d["key"] = str(d["key"]).strip()
            out.append(d)
        return out
    finally:
        wb.close()


def typed_by_key(rows: List[dict]) -> Dict[str, dict]:
    """The clerk's columns per bill id - what the rebuild must carry forward."""
    return {r["key"]: {col_key(c): r.get(col_key(c), "") for c in TYPED} for r in rows if r.get("key")}


def days_since(iso: str, today: Optional[dt.date] = None) -> Optional[int]:
    if not iso:
        return None
    try:
        d = dt.date.fromisoformat(iso[:10])
    except ValueError:
        return None
    return ((today or dt.date.today()) - d).days


def standing(row: dict, today: Optional[dt.date] = None) -> dict:
    """One plain sentence + how long it has sat there, from the stage and the clerk's dates."""
    stage = row.get("stage") or "To send"
    since = {"To send": row.get("revised_in_qbo"), "Sent to vendor": row.get("sent_on"), "Vendor replied": row.get("replied_on"),
             "Revised bill received": row.get("received_on"), "Entered in QBO": row.get("entered_on")}.get(stage) or row.get("revised_in_qbo")
    days = days_since(since, today)
    what = NEXT_QUESTION.get(stage, stage)
    text = what if stage in FINAL else (f"{what} - {days} day{'s' if days != 1 else ''}" if days is not None else what)
    return {"stage": stage, "rank": STAGE_RANK.get(stage, 0), "since": since or "", "days": days, "text": text, "final": stage in FINAL}


def flags_text(flags: List[str]) -> str:
    return " · ".join(f for f in flags if f)
