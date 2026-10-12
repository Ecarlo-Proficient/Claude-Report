"""shared/revised_board.py - the Notion "Revised Bills" board: ONE schema, read by two tools (10/12/2026).

The owner: "I need a way to have the revised bills in a process so we know where they stand - did
the clerk send the bill to get revised? when? have they responded? with the new revised bill? is the
new revised bill uploaded to QBO? that's the part where we can get in a pickle."

A revised bill = a bill a PM fixed in QuickBooks where the TOTAL changed, marked by the
`REVISED - <reason>` memo line (AP-13 in the process registry). The vendor must re-issue it, and
the re-issued bill must land on the SAME QuickBooks bill - a second entry is the pickle.

The board (Notion, data source ACB_REVISED_DS_ID) has one row per such bill:

  the automation owns   Name · Vendor · Bill # · Bill date · Project · PM reason · QBO total ·
  (bill-tracker/        Vendor original · Open balance · Approval · Revised in QBO · Statement says ·
   revised_bills.py)    Flags · QBO bill · Last checked · Key (the QuickBooks bill id, the upsert key)
  the clerk owns        Stage · Sent on · Replied on · Received on · Entered on · Notes
  the automation moves  Stage forward only, and only to "Done - statement agrees" (the vendor's own
                        statement shows the revised amount) or "Dropped" (the bill left QuickBooks)

The ledger (ledger/dashboard.py, /api/revisions) READS the board through `read_rows` and shows
where every revised bill stands; it never writes here. Kept in shared/ because both tools need
the same column names and the same reading of a row (repo rule).
"""
from __future__ import annotations

import datetime as dt
import re
from typing import Dict, Iterator, List, Optional

from shared import paths

DS_ENV = "ACB_REVISED_DS_ID"

STAGES = ["To send", "Sent to vendor", "Vendor replied", "Revised bill received", "Entered in QBO",
          "Done - statement agrees", "Dropped"]
STAGE_RANK = {s: i for i, s in enumerate(STAGES)}
FINAL = {"Done - statement agrees", "Dropped"}
# the question each stage answers, in the owner's words - the ledger shows the NEXT one as "waiting on"
NEXT_QUESTION = {
    "To send": "the clerk has not told the vendor yet",
    "Sent to vendor": "waiting on the vendor's reply",
    "Vendor replied": "waiting on the re-issued bill",
    "Revised bill received": "the re-issued bill is not on the QuickBooks bill yet",
    "Entered in QBO": "waiting for a vendor statement that shows the revised amount",
    "Done - statement agrees": "done",
    "Dropped": "dropped",
}
CLERK_DATES = ["Sent on", "Replied on", "Received on", "Entered on"]
AUTO_TEXT = ["Vendor", "Bill #", "Project", "PM reason", "Approval", "Statement says", "Flags", "Key"]
AUTO_NUMBER = ["QBO total", "Vendor original", "Open balance"]
AUTO_DATE = ["Bill date", "Revised in QBO", "Last checked"]
QBO_BILL_URL = "https://qbo.intuit.com/app/bill?txnId={bill_id}"


def ds_id() -> str:
    """The board's data-source id (machine.env / server.env), '' when the board is not set up."""
    return (paths.get(DS_ENV) or "").strip()


# ─────────────────────────── property builders (write side) ───────────────────────────

def _rt(text: str) -> dict:
    return {"rich_text": [{"type": "text", "text": {"content": (text or "")[:1900]}}]}


def props_auto(row: dict) -> dict:
    """The automation's columns for one revised bill (see revised_bills.find). Never a clerk column."""
    out: dict = {"Name": {"title": [{"type": "text", "text": {"content": f"{row['vendor']} · {row['ref'] or '(no #)'}"[:200]}}]}}
    for k, src in (("Vendor", "vendor"), ("Bill #", "ref"), ("Project", "project"), ("PM reason", "reason"),
                   ("Approval", "approval"), ("Statement says", "statement_says"), ("Flags", "flags_text"), ("Key", "bill_id")):
        out[k] = _rt(str(row.get(src) or ""))
    for k, src in (("QBO total", "total"), ("Vendor original", "original"), ("Open balance", "open")):
        v = row.get(src)
        out[k] = {"number": round(float(v), 2) if v is not None else None}
    for k, src in (("Bill date", "bill_date"), ("Revised in QBO", "revised_at"), ("Last checked", "checked_at")):
        v = row.get(src)
        out[k] = {"date": {"start": str(v)[:19] if v else None}} if v else {"date": None}
    out["QBO bill"] = {"url": QBO_BILL_URL.format(bill_id=row["bill_id"]) if row.get("bill_id") else None}
    return out


def props_stage(stage: str) -> dict:
    return {"Stage": {"select": {"name": stage}}}


# ─────────────────────────── reading a row (both sides) ───────────────────────────

def _plain(prop: Optional[dict]) -> str:
    if not prop:
        return ""
    t = prop.get("type")
    if t in ("rich_text", "title"):
        return "".join(x.get("plain_text", "") for x in prop.get(t) or [])
    if t == "select":
        return ((prop.get("select") or {}).get("name")) or ""
    if t == "url":
        return prop.get("url") or ""
    return ""


def _date(prop: Optional[dict]) -> str:
    if not prop or prop.get("type") != "date":
        return ""
    return ((prop.get("date") or {}).get("start")) or ""


def _number(prop: Optional[dict]):
    if not prop or prop.get("type") != "number":
        return None
    return prop.get("number")


def parse_page(page: dict) -> dict:
    """A Notion page of the board -> one flat row (ISO dates as strings, '' when blank)."""
    p = page.get("properties") or {}
    row = {"page_id": page.get("id"), "url": page.get("url") or "",
           "name": _plain(p.get("Name")), "stage": _plain(p.get("Stage")) or "", "notes": _plain(p.get("Notes"))}
    for k in AUTO_TEXT:
        row[_key(k)] = _plain(p.get(k))
    for k in AUTO_NUMBER:
        row[_key(k)] = _number(p.get(k))
    for k in AUTO_DATE + CLERK_DATES:
        row[_key(k)] = _date(p.get(k))
    row["last_edited"] = page.get("last_edited_time") or ""
    return row


def _key(name: str) -> str:
    return re.sub(r"[^a-z0-9]+", "_", name.lower()).strip("_")


def read_rows(client, ds: Optional[str] = None) -> List[dict]:
    """Every row on the board (the ledger's read, the automation's prior state)."""
    ds = ds or ds_id()
    if not ds:
        return []
    return [parse_page(pg) for pg in client.query_data_source(ds, page_size=100)]


def iter_pages(client, ds: Optional[str] = None) -> Iterator[dict]:
    ds = ds or ds_id()
    if ds:
        yield from client.query_data_source(ds, page_size=100)


# ─────────────────────────── where it stands (both sides) ───────────────────────────

def days_since(iso: str, today: Optional[dt.date] = None) -> Optional[int]:
    if not iso:
        return None
    try:
        d = dt.date.fromisoformat(iso[:10])
    except ValueError:
        return None
    return ((today or dt.date.today()) - d).days


def standing(row: dict, today: Optional[dt.date] = None) -> dict:
    """One plain sentence + how long it has sat there, from the stage and the clerk's dates.
    The ledger prints this; the automation writes nothing of it (the dates speak)."""
    stage = row.get("stage") or "To send"
    since = {"To send": row.get("revised_in_qbo"), "Sent to vendor": row.get("sent_on"), "Vendor replied": row.get("replied_on"),
             "Revised bill received": row.get("received_on"), "Entered in QBO": row.get("entered_on")}.get(stage) or row.get("revised_in_qbo")
    days = days_since(since, today)
    what = NEXT_QUESTION.get(stage, stage)
    text = what if stage in FINAL else (f"{what} - {days} day{'s' if days != 1 else ''}" if days is not None else what)
    return {"stage": stage, "rank": STAGE_RANK.get(stage, 0), "since": since or "", "days": days, "text": text, "final": stage in FINAL}


def flags_text(flags: List[str]) -> str:
    return " · ".join(f for f in flags if f)
