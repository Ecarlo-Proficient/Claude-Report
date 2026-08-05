"""
aging_sheet.py — the "AR Aging" tab of Open_Invoices.xlsx.

Why this exists (the user 2026-08-05):
    Notion is good for reading ONE invoice page. It is bad at the thing the
    owner actually does every week — scanning a hundred rows at once to see
    "who owes me money, how old is it, and what is holding it up." That is a
    QBO-style aging view: one row per invoice, the open balance dropped into
    a Current / 1-30 / 31-60 / 61-90 / 90+ column, all three divisions in one
    place, rolled up under the parent client.

What this tab shows that QBO's own aging does NOT:
    1. **Notes** — the collections clerk's running note on each invoice
       (Notion `Quick Status`) plus the date they last touched it.
    2. **Vendor status** — for MFD and CP, whether we still owe subs/suppliers
       on the draw period that invoice covers. An invoice we are chasing while
       our own vendors are unpaid is a different collection conversation than
       one where we already fronted the money. Sourced from the bill-tracker's
       Excel output (see `load_vendor_bill_map`).

Rules baked in:
    - **Litigation invoices are excluded** (the `Litigation` checkbox in both
      Notion trackers). They are not collections work anymore; they are legal
      work, and leaving them in the aging inflates every bucket.
    - Aging is by DUE DATE, matching QBO's default AR aging and the
      `Aging Bucket` select that invoice_sync already writes to Notion.
    - Parent-client groups are **collapsed by default** — the owner opens the
      client they care about instead of scrolling 60 invoice rows.

This module only builds the worksheet; `export_invoices_xlsx.py` owns the
workbook and the Notion pull.
"""
from __future__ import annotations

import datetime as dt
import logging
import sys
from collections import defaultdict
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from shared import paths


log = logging.getLogger("automation_worker.aging_sheet")

SHEET_TITLE = "AR Aging"

# The bill-tracker's Excel output. Same env key the bill-tracker itself uses,
# so a machine.env override moves both together. We read the FILE, not the
# bill-tracker's code — tools never import tools (repo rule 3).
BILL_TRACKER_PATH = paths.get_path(
    "ACB_BILL_TRACKER_XLSX",
    paths.onedrive_base() / "Automations-" / "Bill Tracker.xlsx",
)

# Vendor status literals (the user's wording: the status is "Vendor Unpaid Bills").
VENDOR_UNPAID = "Vendor Unpaid Bills"
VENDOR_CLEAR = "Vendors Paid"
VENDOR_NA = ""          # RP — the draw-period match doesn't apply, see below
VENDOR_UNKNOWN = "?"    # bill tracker file missing / unreadable

# Divisions where a bill line is matched to the invoice that authorizes paying
# it via the DRAW PERIOD (bill-tracker README, "How matching works"). RP matches
# on "earliest invoice on/after bill date", which is not a draw-period statement,
# so the vendor column is left blank for RP rather than implying a match.
DRAW_DIVISIONS = ("MFD", "CP")

# (header, width, number_format)
COLUMNS: List[Tuple[str, int, Optional[str]]] = [
    ("Client / Invoice", 34, None),
    ("Division",          9, None),
    ("Project #",        13, None),
    ("Invoice #",        11, None),
    ("Date",             11, "mm/dd/yyyy"),
    ("Due Date",         11, "mm/dd/yyyy"),
    ("Days Past Due",     9, '"+"0;-0;0'),
    ("Current",          14, '"$"#,##0.00'),
    ("1-30",             14, '"$"#,##0.00'),
    ("31-60",            14, '"$"#,##0.00'),
    ("61-90",            14, '"$"#,##0.00'),
    ("90+",              14, '"$"#,##0.00'),
    ("Total Open",       15, '"$"#,##0.00'),
    ("Vendor Status",    20, None),
    ("Open Bills",        9, "0"),
    ("Vendor $ Open",    15, '"$"#,##0.00'),
    ("Notes",            46, None),
    ("Last Action",      12, "mm/dd/yyyy"),
]

# 0-based positions used when writing rows (kept in sync with COLUMNS above).
C_LABEL, C_DIV, C_PROJ, C_INV, C_DATE, C_DUE, C_DPD = range(7)
C_CURRENT, C_1_30, C_31_60, C_61_90, C_90 = range(7, 12)
C_TOTAL, C_VSTATUS, C_VBILLS, C_VAMT, C_NOTES, C_ACTION = range(12, 18)

BUCKET_COLS = (C_CURRENT, C_1_30, C_31_60, C_61_90, C_90)

_THIN = Side(style="thin", color="000000")
_MEDIUM = Side(style="medium", color="000000")


# ─────────────────────── aging ───────────────────────

def bucket_index(days_past_due: Optional[int]) -> int:
    """Position in BUCKET_COLS for a signed days-past-due value.

    Positive = overdue. <= 0 (not yet due, or due today) = Current, which is how
    QBO ages AR and how invoice_sync assigns the Notion `Aging Bucket` select.
    An invoice with no due date can't be aged — it lands in Current rather than
    silently disappearing from the row total.
    """
    if days_past_due is None or days_past_due <= 0:
        return 0
    if days_past_due <= 30:
        return 1
    if days_past_due <= 60:
        return 2
    if days_past_due <= 90:
        return 3
    return 4


# ─────────────────── vendor bills (bill tracker) ───────────────────

def _humanize_age(stamp: dt.datetime) -> str:
    """'3 hours' / '2 days' — for telling the reader how old the vendor data is."""
    delta = dt.datetime.now() - stamp
    hours = delta.days * 24 + delta.seconds // 3600
    if hours < 1:
        return "under an hour"
    if hours < 48:
        return f"{hours} hour{'s' if hours != 1 else ''}"
    return f"{delta.days} days"

def load_vendor_bill_map(
    path: Path = BILL_TRACKER_PATH,
) -> Tuple[Optional[Dict[str, Tuple[float, int, int]]], Optional[dt.datetime]]:
    """{invoice # → (open $, bill count, vendor count)} from Bill Tracker.xlsx.

    The bill-tracker's `Bills` sheet is LINE-level: one row per bill line, with
    the bill's own `Bill Open Bal` repeated on every line of that bill. Summing
    the column directly would multiply a bill by its line count, so we dedupe on
    (vendor, bill #, bill date) per invoice before adding.

    Only MFD/CP rows carry a draw-period invoice match, and only lines with an
    open balance are owed — everything else is already paid and irrelevant here.

    Returns (None, None) if the file is missing or unreadable; the caller then
    shows "?" instead of claiming vendors are paid on stale/absent data.
    """
    if not path.exists():
        log.warning(
            "Bill Tracker not found at %s — Vendor Status will show '%s'. "
            "Run `sync-ap` to generate it.", path, VENDOR_UNKNOWN,
        )
        return None, None

    try:
        from openpyxl import load_workbook

        as_of = dt.datetime.fromtimestamp(path.stat().st_mtime)
        wb = load_workbook(path, read_only=True, data_only=True)
        try:
            ws = wb["Bills"]
            rows = ws.iter_rows(min_row=2, values_only=True)  # row 1 = banner
            header = next(rows)
            idx = {name: i for i, name in enumerate(header) if name}
            need = ("Division", "Invoice #", "Bill Open Bal", "Vendor", "Bill #", "Bill Date")
            missing = [c for c in need if c not in idx]
            if missing:
                log.warning("Bill Tracker 'Bills' sheet missing columns %s", missing)
                return None, None

            seen: set = set()
            agg: Dict[str, List[Any]] = defaultdict(lambda: [0.0, 0, set()])
            for row in rows:
                if not any(row):
                    continue
                if row[idx["Division"]] not in DRAW_DIVISIONS:
                    continue
                invoice_num = str(row[idx["Invoice #"]] or "").strip()
                if not invoice_num:
                    continue
                try:
                    balance = float(row[idx["Bill Open Bal"]] or 0)
                except (TypeError, ValueError):
                    continue
                if balance <= 0:
                    continue
                vendor = row[idx["Vendor"]]
                bill_key = (invoice_num, vendor, row[idx["Bill #"]], row[idx["Bill Date"]])
                if bill_key in seen:
                    continue
                seen.add(bill_key)
                entry = agg[invoice_num]
                entry[0] += balance
                entry[1] += 1
                entry[2].add(vendor)
        finally:
            wb.close()

        result = {inv: (amt, bills, len(vendors)) for inv, (amt, bills, vendors) in agg.items()}
        log.info(
            "Vendor bills: %d MFD/CP invoices still carry unpaid bills (tracker as of %s)",
            len(result), as_of.strftime("%Y-%m-%d %H:%M"),
        )
        # AR runs after AP by design (`sync-all`). A tracker older than today
        # means that order was broken — say so at run time instead of letting a
        # day-old vendor column pass for current.
        if as_of.date() < dt.date.today():
            log.warning(
                "Bill Tracker is from %s, not today — the Vendor columns are "
                "%s old. Run `sync-ap` (or `sync-all`, which runs AP first) "
                "to refresh them.",
                as_of.strftime("%b %d"), _humanize_age(as_of),
            )
        return result, as_of
    except Exception as e:
        log.warning("Could not read Bill Tracker (%s): %s", path, e)
        return None, None


def vendor_cells(
    division: str,
    invoice_num: str,
    vendor_map: Optional[Dict[str, Tuple[float, int, int]]],
) -> Tuple[str, Optional[int], Optional[float]]:
    """(status text, open bill count, open $) for one invoice's Vendor columns.

    The status string is one of a small fixed set on purpose — the bill count
    lives in its own column so the Vendor Status autofilter stays a two-item
    dropdown instead of one entry per distinct count.
    """
    if division not in DRAW_DIVISIONS:
        return VENDOR_NA, None, None
    if vendor_map is None:
        return VENDOR_UNKNOWN, None, None
    hit = vendor_map.get(str(invoice_num).strip())
    if not hit:
        return VENDOR_CLEAR, None, None
    amount, bills, _vendors = hit
    return VENDOR_UNPAID, bills, amount


# ─────────────────────── sheet build ───────────────────────

def _write_row(ws: Worksheet, row_num: int, values: List[Any], *, bold: bool) -> None:
    for offset, value in enumerate(values):
        cell = ws.cell(row=row_num, column=offset + 1, value=value)
        number_format = COLUMNS[offset][2]
        if number_format:
            cell.number_format = number_format
        if bold:
            cell.font = Font(bold=True)
    ws.cell(row=row_num, column=C_NOTES + 1).alignment = Alignment(
        horizontal="left", vertical="top", wrap_text=True
    )


def build_aging_sheet(
    ws: Worksheet,
    invoices: List[dict],
    *,
    today: dt.date,
    litigation_excluded: int,
    vendor_as_of: Optional[dt.datetime] = None,
) -> None:
    """Write the aging tab.

    `invoices` are plain dicts (built by export_invoices_xlsx._aging_record) so
    this module stays free of Notion property plumbing:
        parent, division, project_num, invoice_num, invoice_date, due_date,
        open_balance, notes, last_action

    Layout: a header, an always-visible ALL CLIENTS total, then one bold summary
    row per parent client with its invoices grouped underneath and COLLAPSED
    (the owner's ask — Notion-style one-page scanning, drill in on demand).
    Summary-above-detail requires outlinePr.summaryBelow = False; without it
    Excel puts the collapse toggle on the row after the block and the grouping
    reads backwards.
    """
    # Group by parent client. Unresolved relations fall back to a stable label
    # rather than being dropped — an invoice with no parent is still money owed.
    by_parent: Dict[str, List[dict]] = defaultdict(list)
    for rec in invoices:
        by_parent[rec["parent"] or "(no client on file)"].append(rec)

    # ── title block ──
    ws.cell(row=1, column=1, value=f"AR AGING — as of {today.strftime('%b %d, %Y').upper()}").font = Font(bold=True)
    # The vendor columns have a different clock from everything else on this tab:
    # they come from the AP tool's last run, not from this one. Say which, in
    # plain words, so nobody reads a day-old vendor figure as current.
    if vendor_as_of is None:
        vendor_note = "Vendor bill status UNAVAILABLE — run `sync-ap`, then re-run this"
    elif vendor_as_of.date() < today:
        vendor_note = (
            f"⚠ VENDOR COLUMNS ARE {_humanize_age(vendor_as_of).upper()} OLD "
            f"(Bill Tracker last run {vendor_as_of.strftime('%b %d, %I:%M %p')}) — "
            f"run `sync-all`, which runs AP before AR"
        )
    else:
        vendor_note = f"Vendor bill status current as of {vendor_as_of.strftime('%b %d, %I:%M %p')}"
    subtitle = (
        f"{len(invoices)} open invoices · litigation excluded ({litigation_excluded}) · {vendor_note}"
    )
    cell = ws.cell(row=2, column=1, value=subtitle)
    if vendor_as_of is None or vendor_as_of.date() < today:
        cell.font = Font(bold=True, color="C00000")

    header_row = 4
    for offset, (name, width, _) in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=header_row, column=offset, value=name)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(bottom=_MEDIUM)
        ws.column_dimensions[get_column_letter(offset)].width = width
    ws.row_dimensions[header_row].height = 26
    ws.freeze_panes = ws.cell(row=header_row + 1, column=1)

    # ── grand total (always visible, never grouped) ──
    # Roll the grand total up from the same per-invoice values the parent rows
    # use, so the top row always ties to the sum of the groups below it.
    grand = [0.0] * len(BUCKET_COLS)
    grand_vendor = 0.0
    grand_bills = 0
    for rec in invoices:
        grand[bucket_index(rec["days_past_due"])] += rec["open_balance"] or 0.0
        grand_vendor += rec["vendor_amount"] or 0.0
        grand_bills += rec["vendor_bills"] or 0

    row_num = header_row + 1
    total_row: List[Any] = [""] * len(COLUMNS)
    total_row[C_LABEL] = f"ALL CLIENTS ({len(by_parent)})"
    for pos, amount in zip(BUCKET_COLS, grand):
        total_row[pos] = amount or None
    total_row[C_TOTAL] = sum(grand)
    total_row[C_VBILLS] = grand_bills or None
    total_row[C_VAMT] = grand_vendor or None
    _write_row(ws, row_num, total_row, bold=True)
    for offset in range(len(COLUMNS)):
        ws.cell(row=row_num, column=offset + 1).border = Border(bottom=_THIN)
    row_num += 1

    # ── parent groups, biggest balance first ──
    def parent_total(records: List[dict]) -> float:
        return sum(r["open_balance"] or 0.0 for r in records)

    for parent in sorted(by_parent, key=lambda p: parent_total(by_parent[p]), reverse=True):
        records = sorted(
            by_parent[parent],
            key=lambda r: (r["due_date"] or dt.date.max, r["invoice_num"]),
        )
        buckets = [0.0] * len(BUCKET_COLS)
        vendor_sum = 0.0
        vendor_bills = 0
        for rec in records:
            buckets[bucket_index(rec["days_past_due"])] += rec["open_balance"] or 0.0
            vendor_sum += rec["vendor_amount"] or 0.0
            vendor_bills += rec["vendor_bills"] or 0

        # A parent row carries a Division so the Division filter keeps the whole
        # group visible; clients that span divisions say so instead of picking one.
        divisions = sorted({r["division"] for r in records if r["division"]})
        division_label = divisions[0] if len(divisions) == 1 else "(mixed)"

        summary: List[Any] = [""] * len(COLUMNS)
        summary[C_LABEL] = parent
        summary[C_DIV] = division_label
        summary[C_INV] = f"{len(records)} inv"
        for pos, amount in zip(BUCKET_COLS, buckets):
            summary[pos] = amount or None
        summary[C_TOTAL] = sum(buckets)
        summary[C_VBILLS] = vendor_bills or None
        summary[C_VAMT] = vendor_sum or None
        _write_row(ws, row_num, summary, bold=True)
        for offset in range(len(COLUMNS)):
            ws.cell(row=row_num, column=offset + 1).border = Border(top=_THIN)
        row_num += 1

        for rec in records:
            detail: List[Any] = [""] * len(COLUMNS)
            detail[C_LABEL] = f"    {rec['memo'] or rec['project_num'] or ''}"[:120]
            detail[C_DIV] = rec["division"]
            detail[C_PROJ] = rec["project_num"]
            detail[C_INV] = rec["invoice_num"]
            detail[C_DATE] = rec["invoice_date"]
            detail[C_DUE] = rec["due_date"]
            detail[C_DPD] = rec["days_past_due"] if rec["days_past_due"] is not None else ""
            detail[BUCKET_COLS[bucket_index(rec["days_past_due"])]] = rec["open_balance"]
            detail[C_TOTAL] = rec["open_balance"]
            detail[C_VSTATUS] = rec["vendor_status"]
            detail[C_VBILLS] = rec["vendor_bills"]
            detail[C_VAMT] = rec["vendor_amount"]
            detail[C_NOTES] = rec["notes"]
            detail[C_ACTION] = rec["last_action"]
            _write_row(ws, row_num, detail, bold=False)

            # >30 days overdue reads as red, same cue as the Open Invoices tab.
            days = rec["days_past_due"]
            if isinstance(days, int) and days > 30:
                ws.cell(row=row_num, column=C_DPD + 1).font = Font(bold=True, color="C00000")

            ws.row_dimensions[row_num].outlineLevel = 1
            ws.row_dimensions[row_num].hidden = True  # collapsed by default
            row_num += 1

    ws.sheet_properties.outlinePr.summaryBelow = False
    ws.sheet_properties.outlinePr.applyStyles = False
    ws.auto_filter.ref = f"A{header_row}:{get_column_letter(len(COLUMNS))}{row_num - 1}"
    ws.sheet_view.showGridLines = True
