#!/usr/bin/env python3
"""
excel_bill_sync.py — daily QBO → xlsx AP bill tracker.

Output: /Users/sebas/Documents/CompanyHealth/Bill Tracker.xlsx  (chmod 600)

Five sheets, every one an Excel Table with AutoFilter on every column.
Same column set across all five, sorted differently per sheet:

  1. Pay List        — Pay = x AND not yet paid; vendor → date
  2. Open Bills      — open bills only (Balance > 0); vendor → date → line
  3. Project Ledger  — open + paid since paid_cutoff; project → vendor → date
  4. Client Ledger   — open + paid since paid_cutoff; client → project → date
  5. By Division     — open + paid since paid_cutoff; div → proj → inv → bill

Editable cells: Lien (Notice Sent / Lien Filed), Notes (free text). Both are
preserved across syncs by _Key (bill_id). Pay marking was removed 2026-06-18 —
the owner now copies the bills to pay into a separate workbook, so Bill
Tracker.xlsx is script-owned and the clerk treats it as read-only except for
the Lien tag and Notes.

Lien tag (escalation, two steps): "Notice Sent" = a supplier/sub sent a
preliminary lien notice on an unpaid bill; "Lien Filed" = a real lien was
recorded. Both persist across syncs (a lien doesn't auto-clear on payment).

Color: RED row tint = NOT APPROVED. Lien cell tint: amber = Notice Sent,
red = Lien Filed.

USAGE
  python3 bill-tracker/excel_bill_sync.py
  python3 bill-tracker/excel_bill_sync.py --dry-run   # build rows, don't write
  python3 bill-tracker/excel_bill_sync.py --limit 50  # smoke test
"""
from __future__ import annotations

import argparse
import datetime as dt
import os
import re
import stat
import sys
from collections import defaultdict
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple


# ─────────────────────── path bootstraps ───────────────────────

PROJECT_ROOT = Path(__file__).resolve().parent.parent
BILL_TRACKER_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(PROJECT_ROOT))           # qbo_vault.py
sys.path.insert(0, str(BILL_TRACKER_DIR))       # qbo_bill_tracker.py

from shared import paths


# ─────────────────────── deps ───────────────────────

try:
    import requests  # noqa: F401  (used transitively by QBO extraction)
except ImportError:
    print("✗ pip3 install --break-system-packages requests")
    sys.exit(1)

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Color
    from openpyxl.styles.differential import DifferentialStyle
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table, TableStyleInfo
    from openpyxl.worksheet.filters import AutoFilter, FilterColumn, Filters
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.formatting.rule import Rule
    from openpyxl.worksheet.formula import ArrayFormula
except ImportError:
    print("✗ pip3 install --break-system-packages openpyxl")
    sys.exit(1)

from qbo_bill_tracker import (
    load_credentials, query_all, is_sub_bill, is_excluded_invoice,
    parse_date,
    STATUS_OK_TO_PAY, STATUS_AWAITING_PAYMENT, STATUS_AWAITING_INVOICE,
    STATUS_PAID, STATUS_NO_PROJECT, STATUS_PARTIAL_PAID,
)
from bill_rows import (
    build_account_maps, build_po_map, build_payment_map, build_rows,
    display_status, approved_text,
    collapse_rows, multi_project_bill_ids, MULTI_MARKER,
)


# ─────────────────────── constants ───────────────────────

OUTPUT_PATH = paths.get_path(
    "ACB_BILL_TRACKER_XLSX",
    paths.onedrive_base() / "Automations-/Bill Tracker.xlsx",
)
BACKUP_RETENTION_DAYS = 14

# Paid bills lookback. Per Ted 2026-05-27: trailing 12mo was too much; use
# fixed YTD start date instead. Adjust here when crossing a fiscal year.
PAID_CUTOFF_DATE = "2026-01-01"

# Invoice query cutoff. Used to filter the Invoice pull — invoices older
# than this are skipped on the assumption they can't match any open bill.
# Pushed back to 2024-01-01 to comfortably cover the oldest open bills
# (the December 2024 MCP residual etc.) without dragging in years of
# historic invoices that quintuple the QBO query load.
INVOICE_CUTOFF_DATE = "2024-01-01"

QBO_BILL_URL_TEMPLATE = "https://qbo.intuit.com/app/bill?txnId={bill_id}"
QBO_INVOICE_URL_TEMPLATE = "https://qbo.intuit.com/app/invoice?txnId={inv_id}"

# Muted palette — chrome is calm so the two CF signals can read at a glance.
HEADER_FILL = PatternFill("solid", start_color="6E7E94")
DIVIDER_FILL = PatternFill("solid", start_color="2F3B4D")
HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
BODY_FONT = Font(name="Calibri", size=11)
ALIGN_CENTER = Alignment(horizontal="center", vertical="center")
ALIGN_LEFT = Alignment(horizontal="left", vertical="center")
ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")

# Supra-header band colors — one per section. Semantic mapping:
#   BILL    = navy   (cool, identification)
#   STATUS  = amber  (warm, action needed)
#   INVOICE = green  (stable, verified data)
# Ted 2026-06-04: swapped INVOICE and STATUS so green sits on Invoice.
SUPRA_BILL_FILL    = PatternFill("solid", start_color="1F3864")  # deep navy
SUPRA_STATUS_FILL  = PatternFill("solid", start_color="BF8F00")  # warm amber
SUPRA_INVOICE_FILL = PatternFill("solid", start_color="375623")  # deep green
SUPRA_PAY_FILL     = PatternFill("solid", start_color="31849B")  # teal (AP cash-out)
SUPRA_FONT         = Font(name="Calibri", size=12, bold=True,
                          color="FFFFFF", italic=False)
SUPRA_HEIGHT       = 22

MONEY_FMT = '"$"#,##0.00;[Red]("$"#,##0.00);"-"'
DATE_FMT = "m/d/yyyy"

# Status × Approved row colors. Pay-cell green stacks on top.
# Ted 2026-06-04: softened "hold" (red was too aggressive) and bumped
# "partial" to a warmer amber-yellow so it reads as actionable.
CF_GREEN_READY = "C6EFCE"        # Invoice paid + approved → pay the vendor
CF_ORANGE_URGENT = "FFC299"      # Invoice paid + not approved → approve now
CF_RED_HOLD = "FCE4E4"           # Awaiting + not approved → held (soft pink)
CF_GRAY_DONE = "D9D9D9"          # Bill paid + approved → done
CF_YELLOW_AUDIT = "FFEB9C"       # Bill paid + not approved → audit
CF_PURPLE_NEEDS_PROJECT = "E4D7F5"  # No project # → needs project
CF_PEACH_PARTIAL = "FFCC66"      # Partial paid → warm amber, action needed
CF_LIEN_NOTICE = "FFC000"        # amber — lien NOTICE sent (escalation step 1)
CF_LIEN_FILED  = "FF0000"        # red — real LIEN filed (escalation step 2)
CF_LIEN_RELEASED = "92D050"      # green — lien RELEASED / satisfied (resolved)

# Lien-NOTICE timer (Ted 2026-06-18): an UNPAID, UNTAGGED bill whose supplier
# lien-notice date (15th of the 2nd month after the bill date — TX practice)
# is approaching. A yellow→orange countdown. It NEVER goes red: red is reserved
# for an actual FILED lien (a real legal event), not a "this is coming" warning.
# 2026-07-10 (Ted): old ramp was too soft — ≤30d was near-invisible and ≤15d
# blended into the Notice Sent amber. Each step now one notch hotter.
# 2026-07-10 (Ted): these cells are SCRIPT-written (the countdown), so they read
# as soft/tentative — PALE tints + grey italic text — vs the human lien tags,
# which stay saturated + bold black. The tier still escalates gold→coral→rose;
# the bucket text carries the exact window. PAST is a distinct rose (NOT red —
# red stays reserved for an actual FILED lien).
CF_TIMER_YELLOW = "FFF2CC"       # ≤ 30 days out — pale gold
CF_TIMER_ORANGE = "FCE4D6"       # ≤ 15 days out — pale peach
CF_TIMER_HOT    = "F8CBAD"       # ≤ 7 days out — pale coral
CF_TIMER_PAST   = "E6B8B8"       # notice deadline passed — pale rose
CF_TIMER_TEXT   = "808080"       # grey italic — signals "script-written, not a tag"

# The Lien cell itself shows the countdown bucket as text (Ted 2026-07-10) so a
# clerk reads the state without cross-referencing the key. These label strings
# are the single source of truth: written into the cell AND matched by the CF
# rules that color it, so text and color can never disagree.
TIMER_LABEL_30   = "Notice due in ≤30d"
TIMER_LABEL_15   = "Notice due in ≤15d"
TIMER_LABEL_7    = "Notice due in ≤7d"
TIMER_LABEL_PAST = "Notice PAST due"

# Lien tag values (escalation order). "Outstanding" = NOTICE or FILED; a bill
# being paid does NOT resolve a lien — only RELEASED does.
LIEN_NOTICE = "Notice Sent"
LIEN_FILED = "Lien Filed"
LIEN_RELEASED = "✓ Released"
LIEN_OUTSTANDING = (LIEN_NOTICE, LIEN_FILED)
LIEN_ALL = (LIEN_NOTICE, LIEN_FILED, LIEN_RELEASED)

# Aging bucket boundaries — strictly informational column on every sheet
AGING_BUCKETS: List[Tuple[str, int, int]] = [
    ("0–30",  0,  30),
    ("31–60", 31, 60),
    ("61–90", 61, 90),
    ("90+",   91, 10**9),
]


# ─────────────────────── unified bill-row column spec ───────────────────────
#
# Every visible sheet emits rows in this exact column order. Hidden columns
# (_Key) are present for the merge logic but hidden from view. Divider columns
# (│) are 1.5-wide dark stripes splitting bill / invoice / status sections.

BILL_ROW_COLS: List[Tuple[str, str]] = [
    # (header, kind) — kind controls number_format / alignment / font
    # 2026-06-03: dropped Days Open / Bucket / Bill Type. Moved Open to col A.
    # 2026-06-04: full reorder — BILL → STATUS → INVOICE.
    # 2026-06-04: Line Amount restored at col 9 — hidden on Bills (where at
    #   bill grain it equals Bill Total) but VISIBLE on Inventory (where
    #   it's the per-line dollar amount, the whole point of that sheet).
    ("Open",               "link"),      # 1  — tiny QBO bill link
    # ── BILL section ──
    ("Vendor",             "text"),      # 2
    ("Bill #",             "text"),      # 3
    ("Bill Date",          "date"),      # 4
    ("Project #",          "text"),      # 5
    ("Division",           "text"),      # 6
    ("Account",            "text"),      # 7
    ("Line Description",   "text"),      # 8
    ("Line Amount",        "money"),     # 9  — hidden on Bills, shown on Inventory
    ("Bill Total",         "money"),     # 10 — original bill amount
    ("Bill Open Bal",      "money"),     # 11 — what's still owed
    ("│",                  "div"),       # 12 — divider bill | status
    # ── STATUS & ACTION section ──
    ("Status",             "text"),      # 13 — pipeline only
    ("Approved",           "text"),      # 14 — "approved" / "not approved"
    ("Lien",               "text"),      # 15 — Notice Sent / Lien Filed; hidden on Inventory
    ("Notes",              "text"),      # 16
    ("┃",                  "div"),       # 17 — divider status | invoice
    # ── INVOICE section (verification only) ──
    ("Invoice #",          "text"),      # 18 — just the # → QBO link
    ("Matched Invoice",    "text"),      # 19 — # + memo (scope), for eyeballing the match
    ("Invoice Date",       "date"),      # 20
    ("Invoice Total",      "money"),     # 21
    ("Payment Date",       "date"),      # 22 — INVOICE paid date (GC → us, money IN)
    ("Client",             "text"),      # 23 — GC/parent customer; filter by client
    # ── PAYMENT section (AP cash-OUT: the bill payment we cut) ──
    ("Pay Ref #",          "text"),      # 24 — check # we paid with (blank for CC)
    ("Pay Date",           "date"),      # 25 — when we paid the vendor
    ("Pay Method",         "text"),      # 26 — Check / CC / (multiple)
    ("_Key",               "text"),      # 27 — hidden merge join key
]
LINE_AMT_COL_INDEX = next(i + 1 for i, (h, _) in enumerate(BILL_ROW_COLS) if h == "Line Amount")
HEADERS = [h for h, _ in BILL_ROW_COLS]
KINDS   = [k for _, k in BILL_ROW_COLS]
# Divider columns are marked by kind=="div"; their distinct headers (│ vs ┃)
# render visually similar but satisfy Excel Table's unique-column-name rule.
DIVIDER_COL_INDEXES = {i + 1 for i, (_, k) in enumerate(BILL_ROW_COLS) if k == "div"}
KEY_COL_INDEX = HEADERS.index("_Key") + 1
LIEN_COL_INDEX = HEADERS.index("Lien") + 1
NOTES_COL_INDEX = HEADERS.index("Notes") + 1
BILL_OPEN_BAL_COL_INDEX = HEADERS.index("Bill Open Bal") + 1
STATUS_COL_INDEX = HEADERS.index("Status") + 1
APPROVED_COL_INDEX = HEADERS.index("Approved") + 1
OPEN_COL_INDEX = HEADERS.index("Open") + 1

# Per-sheet column widths — matches BILL → STATUS → INVOICE.
COL_WIDTHS: Dict[int, float] = {
    1: 5,                                            # Open
    # BILL section
    2: 28,                                           # Vendor
    3: 12, 4: 11,                                    # Bill # / Bill Date
    5: 11, 6: 8,                                     # Project # / Division
    7: 22, 8: 35,                                    # Account / Line Description
    9: 13,                                           # Line Amount (Inventory only)
    10: 13, 11: 13,                                  # Bill Total / Bill Open Bal
    12: 1.5,                                         # divider │
    # STATUS section
    13: 18, 14: 14,                                  # Status / Approved
    15: 13, 16: 28,                                  # Lien / Notes
    17: 1.5,                                         # divider ┃
    # INVOICE section
    18: 10, 19: 60, 20: 11, 21: 12, 22: 12,          # Invoice # / Matched Inv / Inv Date / Inv Total / Pay Date
    23: 26,                                          # Client (GC / parent customer)
    # PAYMENT section (AP cash-out)
    24: 14, 25: 11, 26: 12,                          # Pay Ref # / Pay Date / Pay Method
    27: 14,                                          # _Key (hidden)
}


# ─────────────────────── small helpers ───────────────────────

def _col_letter(idx_1based: int) -> str:
    return get_column_letter(idx_1based)


def _bucket_for_age(days: Optional[int]) -> str:
    if days is None:
        return ""
    for label, lo, hi in AGING_BUCKETS:
        if lo <= days <= hi:
            return label
    return ""


def _qbo_link(bill_id: str) -> str:
    """Render the QBO bill link as a single-glyph 'button'. Column is ~5 wide
    so we keep the label tiny; the cell is centered and underlined by
    _format_data_cell's link branch."""
    if not bill_id:
        return ""
    url = QBO_BILL_URL_TEMPLATE.format(bill_id=bill_id)
    return f'=HYPERLINK("{url}","↗")'


def _invoice_cell(inv_doc: str, inv_memo: str) -> str:
    """Format Invoice # together with the invoice's PrivateNote (draw memo)
    so the AP team can eyeball whether the QBO match is correct. No truncation
    — column is wide enough, and clipping the period notation defeats the
    purpose of showing the memo at all."""
    if not inv_doc:
        return ""
    memo = (inv_memo or "").strip()
    if not memo:
        return inv_doc
    return f"{inv_doc} — {memo}"


def _esc_col(name: str) -> str:
    """Escape a column name for use inside Excel `tbl[Col]` structured refs.
    `#`, `'`, `[`, `]` are reserved and must be prefixed with `'`.
    LibreOffice accepts unescaped; Excel silently strips formulas without it."""
    return (name.replace("'", "''")
                .replace("#", "'#")
                .replace("[", "'[")
                .replace("]", "']"))


def _format_data_cell(c, kind: str) -> None:
    c.font = BODY_FONT
    if kind == "money":
        c.number_format = MONEY_FMT
        c.alignment = ALIGN_RIGHT
    elif kind == "date":
        c.number_format = DATE_FMT
        c.alignment = ALIGN_CENTER
    elif kind == "flag":
        c.number_format = "0"
        c.alignment = ALIGN_CENTER
    elif kind == "link":
        c.font = Font(name="Calibri", size=11, color="0563C1", underline="single")
        c.alignment = ALIGN_CENTER
    elif kind == "div":
        # Divider cells get the dark fill; no value, no alignment
        c.fill = DIVIDER_FILL
    else:
        c.alignment = ALIGN_LEFT


def _cf_fill_rule(formula: str, hex_color: str, bold: bool = False,
                  stop_if_true: bool = False, font_color: str = "",
                  italic: bool = False) -> Rule:
    """Build a properly-formed conditional-formatting Rule.

    openpyxl's PatternFill("solid", start_color=…) shortcut emits a dxf that
    Excel rejects (missing patternType). We construct DifferentialStyle with
    explicit fg/bg colors; the post_process_xlsx step further patches the XML
    if openpyxl strips patternType anyway.

    `stop_if_true=True` halts CF evaluation when this rule matches — used for
    the NOT APPROVED red rule so it wins over OK-TO-PAY green when a row could
    match both (NOT APPROVED bill that's also OK TO PAY = still don't pay).

    `font_color` sets the dxf font color; `italic`/`bold` set the weight/style
    (grey italic marks script-written text; bold black marks a human tag).
    """
    fill = PatternFill(
        patternType="solid",
        fgColor=Color(rgb=f"FF{hex_color}"),
        bgColor=Color(rgb=f"FF{hex_color}"),
    )
    if bold or font_color or italic:
        dxf = DifferentialStyle(
            fill=fill,
            font=Font(bold=bold, italic=italic, color=(font_color or None)))
    else:
        dxf = DifferentialStyle(fill=fill)
    return Rule(type="expression", formula=[formula], dxf=dxf, stopIfTrue=stop_if_true)


# ─────────────────────── post-process Excel quirk fixes ───────────────────────

def post_process_xlsx(path: Path) -> None:
    """Patch openpyxl's two known Excel-rejection quirks via XML regex.

    1. dxf <patternFill> elements emitted by openpyxl drop `patternType="solid"`.
       Excel needs it; LibreOffice tolerates absence. Inject it.
    2. type="list" data validations get serialized with `operator="between"`
       and `<formula2>0</formula2>`. Excel rejects on list validations.
       Strip both.

    Both issues produce "We found a problem with some content" warnings on open.
    """
    import zipfile
    import shutil

    tmp = path.with_suffix(".xlsx.tmp")
    with zipfile.ZipFile(path, "r") as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/styles.xml":
                text = data.decode("utf-8")
                text = re.sub(
                    r'<patternFill(?![^>]*patternType)([^>/]*)>',
                    r'<patternFill patternType="solid"\1>',
                    text,
                )
                data = text.encode("utf-8")
            elif item.filename.startswith("xl/worksheets/sheet") and item.filename.endswith(".xml"):
                text = data.decode("utf-8")
                def _fix_dv(m: "re.Match") -> str:
                    block = m.group(0)
                    if 'type="list"' not in block:
                        return block
                    block = re.sub(r'\s+operator="[^"]*"', "", block)
                    block = re.sub(r'<formula2>[^<]*</formula2>', "", block)
                    return block
                text = re.sub(
                    r"<dataValidation\b[^>]*>.*?</dataValidation>",
                    _fix_dv, text, flags=re.S,
                )
                data = text.encode("utf-8")
            zout.writestr(item, data)
    shutil.move(tmp, path)


# ─────────────────────── xlsx validator (Excel-strict preflight) ───────────────────────

def validate_xlsx(path: Path) -> List[str]:
    """Inspect the built xlsx for patterns Excel rejects but LibreOffice tolerates.

    Returns a list of human-readable failure messages. Empty list = clean.
    Run AFTER post_process_xlsx so we catch only un-patched issues.

    Patterns checked (each one has burned us before):
      A. dxf <patternFill> missing patternType="solid"
      B. dataValidation type="list" with operator="between" or <formula2>
      C. Excel Table with duplicate column names
      D. Unescaped # / [ / ] inside tbl[Col] structured references
      E. Excel 365 dynamic-array functions (FILTER/SORT/UNIQUE/etc) in cells
         without t="array" on the <f> tag
    """
    import zipfile
    DYNAMIC_ARRAY_FNS = ("FILTER", "SORT", "SORTBY", "UNIQUE", "SEQUENCE",
                         "RANDARRAY", "TOROW", "TOCOL", "WRAPROWS", "WRAPCOLS")
    failures: List[str] = []

    with zipfile.ZipFile(path, "r") as z:
        names = z.namelist()

        # A. dxf patternType
        if "xl/styles.xml" in names:
            text = z.read("xl/styles.xml").decode("utf-8")
            m = re.search(r"<dxfs[^>]*>(.*?)</dxfs>", text, re.S)
            if m:
                content = m.group(1)
                total = len(re.findall(r"<patternFill\b", content))
                ok = len(re.findall(r'<patternFill[^>]*patternType=', content))
                if total != ok:
                    failures.append(
                        f"A. styles.xml dxf: {total - ok} <patternFill> missing patternType "
                        f"(Excel rejects — openpyxl drops it in DifferentialStyle output)"
                    )

        # Per-sheet checks
        for fn in names:
            if not (fn.startswith("xl/worksheets/sheet") and fn.endswith(".xml")):
                continue
            sheet_text = z.read(fn).decode("utf-8")
            sheet_label = fn.rsplit("/", 1)[-1]

            # B. dataValidation list with stray operator / formula2
            for dv in re.findall(r"<dataValidation\b[^>]*>.*?</dataValidation>", sheet_text, re.S):
                if 'type="list"' in dv:
                    if 'operator="between"' in dv:
                        failures.append(f"B. {sheet_label}: dataValidation type=list has operator=between")
                    if "<formula2>" in dv:
                        failures.append(f"B. {sheet_label}: dataValidation type=list has stray <formula2>")

            # D. Unescaped # / [ / ] inside tbl[Col]. Match `tblName[Foo#Bar]` style.
            for ref in re.findall(r"\btbl[A-Za-z][A-Za-z0-9_]*\[([^]]*)\]", sheet_text):
                if "#" in ref and "'#" not in ref:
                    fixed = ref.replace("#", "'#")
                    failures.append(
                        f"D. {sheet_label}: structured ref contains unescaped # → "
                        f"tbl[{ref}] (should be tbl[{fixed}])"
                    )

            # E. dynamic-array functions written without t="array"
            array_formula_refs: List[str] = []  # collect for F check below
            for fn_match in re.finditer(r"<c\b[^>]*>(.*?)</c>", sheet_text, re.S):
                cell_xml = fn_match.group(0)
                f_match = re.search(r"<f\b([^>]*)>([^<]*)</f>", cell_xml)
                if not f_match:
                    continue
                f_attrs, f_text = f_match.group(1), f_match.group(2)
                if 't="array"' in f_attrs:
                    ref_m = re.search(r'ref="([^"]+)"', f_attrs)
                    if ref_m:
                        array_formula_refs.append(ref_m.group(1))
                    continue
                # Check if the formula uses any dynamic-array function
                f_text_upper = f_text.upper()
                for fn_name in DYNAMIC_ARRAY_FNS:
                    if re.search(rf"\b{fn_name}\s*\(", f_text_upper):
                        failures.append(
                            f"E. {sheet_label}: dynamic-array fn {fn_name}() in cell formula "
                            f"without t=\"array\" attribute → Excel rejects on open. "
                            f"Use openpyxl's ArrayFormula(ref=…, text=…) instead."
                        )
                        break

            # F. ArrayFormula spill range contains other formatted cells.
            # `<c r="A3" s="2"/>` (style-only, no value) inside the declared
            # spill envelope blocks Excel's array spill → "Removed Records".
            # Detect by checking for any `<c s="…"/>` cells within the ref range.
            for spill_ref in array_formula_refs:
                # parse range like "A2:Y1001"
                m = re.match(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", spill_ref)
                if not m:
                    continue
                start_row, end_row = int(m.group(2)), int(m.group(4))
                # find any <c> within the row range that has s= but no value
                # (the FILTER's anchor cell at start_row is exempt — it's the array)
                blockers = 0
                for cell_m in re.finditer(r'<c r="([A-Z]+)(\d+)"([^>]*)/>', sheet_text):
                    cell_row = int(cell_m.group(2))
                    if cell_row <= start_row or cell_row > end_row:
                        continue
                    attrs = cell_m.group(3)
                    if "s=" in attrs:  # style-attached, blocks spill
                        blockers += 1
                if blockers > 0:
                    failures.append(
                        f"F. {sheet_label}: ArrayFormula spill range {spill_ref} contains "
                        f"{blockers} style-attached empty cells — Excel will refuse spill. "
                        f"Remove per-cell formatting inside the spill range; use column "
                        f"defaults (ws.column_dimensions[X].number_format) instead."
                    )

        # C. Tables with duplicate column names
        for fn in names:
            if not (fn.startswith("xl/tables/table") and fn.endswith(".xml")):
                continue
            tbl_text = z.read(fn).decode("utf-8")
            tbl_name_m = re.search(r'displayName="([^"]+)"', tbl_text)
            tbl_name = tbl_name_m.group(1) if tbl_name_m else fn
            cols = re.findall(r'<tableColumn[^/]*name="([^"]+)"', tbl_text)
            dupes = sorted({c for c in cols if cols.count(c) > 1})
            if dupes:
                failures.append(
                    f"C. Table {tbl_name}: duplicate column names {dupes} "
                    f"(Excel Tables require unique headers)"
                )

    return failures


# ─────────────────────── backup rotation ───────────────────────

def rotate_backup(workbook_path: Path, retention_days: int = BACKUP_RETENTION_DAYS) -> None:
    """Snapshot the existing workbook before overwriting. One file per day;
    prune anything older than `retention_days`.
    """
    import shutil
    if not workbook_path.exists():
        return
    backups_dir = workbook_path.parent / "_backups"
    backups_dir.mkdir(parents=True, exist_ok=True)
    stem = workbook_path.stem
    today_iso = dt.date.today().isoformat()
    target = backups_dir / f"{stem} — {today_iso}{workbook_path.suffix}"
    try:
        shutil.copy2(workbook_path, target)
    except Exception as e:
        print(f"  ⚠ backup copy failed: {e}")
        return

    cutoff = dt.date.today() - dt.timedelta(days=retention_days)
    pruned = 0
    for f in backups_dir.glob(f"{stem} — *.xlsx"):
        try:
            date_str = f.stem.rsplit(" — ", 1)[1]
            file_date = dt.date.fromisoformat(date_str)
        except (IndexError, ValueError):
            continue
        if file_date < cutoff:
            try:
                f.unlink()
                pruned += 1
            except Exception as e:
                print(f"  ⚠ prune failed for {f.name}: {e}")
    print(f"  ✓ backup → _backups/{target.name}"
          f"{f' (pruned {pruned} older than {retention_days}d)' if pruned else ''}")


# ─────────────────────── Lien/Notes preservation (single-sheet) ───────────────────────
#
# Bills is the only editable sheet (division + Audit sheets are read-only,
# rebuilt each sync). Preservation reads the existing Bills sheet and restores
# the two human-edited columns — Lien and Notes — by _Key (bill_id) on the
# next build. Pay was removed 2026-06-18 (owner pastes the pay list into a
# separate workbook), so it is no longer carried forward.

SHEET_NAMES = ["Bills"]
BILLS_TABLE = "tblBills"


def _read_sheet_edits(ws) -> Dict[str, Dict[str, str]]:
    """Return {_Key → {Lien, Notes}} from a sheet that follows BILL_ROW_COLS.

    Auto-detects whether the header row is row 1 (legacy single-header) or
    row 2 (current layout with supra-header at row 1). Requires "_Key" and
    "Notes"; "Lien" is optional (absent in pre-2026-06-18 workbooks → blank).
    """
    out: Dict[str, Dict[str, str]] = {}
    if ws.max_row < 2:
        return out

    # Detect header row — try row 1 first (legacy), fall back to row 2 (current)
    header_row_idx = None
    for candidate in (1, 2):
        try:
            cells = [c.value for c in ws[candidate]]
        except Exception:
            continue
        if "_Key" in cells and "Notes" in cells:
            header_row_idx = candidate
            break
    if header_row_idx is None:
        return out

    header_row = [c.value for c in ws[header_row_idx]]
    key_i = header_row.index("_Key")
    notes_i = header_row.index("Notes")
    lien_i = header_row.index("Lien") if "Lien" in header_row else None

    data_start = header_row_idx + 1
    for row in ws.iter_rows(min_row=data_start, values_only=True):
        if key_i >= len(row):
            continue
        key = row[key_i]
        if not key:
            continue
        lien_val = ""
        if lien_i is not None and lien_i < len(row) and row[lien_i] is not None:
            # Only real tags are preserved — the notice-countdown text the sheet
            # writes into blank Lien cells must NOT be carried forward as an edit
            # (it's recomputed each build).
            if row[lien_i] in LIEN_ALL:
                lien_val = row[lien_i]
        out[str(key)] = {
            "Lien": lien_val,
            "Notes": (row[notes_i] if notes_i < len(row) and row[notes_i] is not None else ""),
        }
    return out


def preserve_edits(path: Path) -> Dict[str, Dict[str, str]]:
    """Read existing workbook's Bills sheet and return {_Key → {Lien, Notes}}.

    Bills is the only editable sheet — division sheets and the Audit sheet are
    read-only views, rebuilt each sync. Both Lien and Notes are ALWAYS
    preserved (never auto-cleared — a lien persists until manually released,
    and Notes are commentary).

    Legacy migration: _Key was per-line (`bill_id-line_id`) before the Bills
    sheet collapsed to bill-grain rows. Any line-level keys found are collapsed
    to bill-level so edits carry over. First non-empty value wins per bill.

    Returns empty dict on first run (no file) or unreadable workbook —
    safe to feed into `_render_rows` either way.
    """
    if not path.exists():
        return {}
    try:
        wb = load_workbook(path, data_only=True)
    except Exception as e:
        print(f"  ⚠ couldn't open existing workbook for preservation: {e}")
        return {}
    # Fall back through legacy sheet names so an upgrade from older layouts
    # (Open Bills, Pay List, AP Workspace) doesn't lose pending edits.
    raw: Dict[str, Dict[str, str]] = {}
    for name in ("Bills", "Open Bills", "Pay List", "AP Workspace"):
        if name in wb.sheetnames:
            raw = _read_sheet_edits(wb[name])
            break
    if not raw:
        return {}

    # Migrate line-level keys (e.g. "12345-0") to bill-level ("12345").
    # Detection: a key with a trailing "-<digits>" segment is line-level.
    migrated: Dict[str, Dict[str, str]] = {}
    line_keys_seen = 0
    for k, v in raw.items():
        m = re.match(r"^(.+?)-(\d+)$", k)
        bill_key = m.group(1) if m else k
        if m:
            line_keys_seen += 1
        slot = migrated.setdefault(bill_key, {"Lien": "", "Notes": ""})
        # First non-empty value wins per bill (Lien and Notes alike).
        if (v.get("Lien") or "").strip() and not slot["Lien"]:
            slot["Lien"] = v["Lien"]
        if (v.get("Notes") or "").strip() and not slot["Notes"]:
            slot["Notes"] = v["Notes"]
    if line_keys_seen:
        print(f"  ↻ migrated {line_keys_seen} line-level _Keys → "
              f"{len(migrated)} bill-level _Keys for collapsed view")
    return migrated


# ─────────────────────── sheet rendering ───────────────────────

def _apply_header(ws, hide_cols: Optional[List[int]] = None) -> None:
    """Write a two-row header: supra-header bands at row 1 (BILL / INVOICE /
    STATUS & ACTION) + column-name headers at row 2 (these are the Table
    headers with AutoFilter dropdowns).

    The supra-header sits OUTSIDE the Table — Excel only allows one header
    row per Table, so the Table's ref begins at row 2.
    """
    # Row 1 — supra-header section bands.
    # The three sections are defined by the divider column positions:
    #   BILL    = cols 1 .. divider1-1
    #   INVOICE = cols divider1+1 .. divider2-1
    #   STATUS  = cols divider2+1 .. end (excluding hidden _Key)
    dividers = sorted(DIVIDER_COL_INDEXES)  # [11, 16] for current layout
    d1, d2 = dividers[0], dividers[1]
    # The PAYMENT band (AP cash-out) sits at the end. It has no divider column
    # of its own (that would break the 2-divider assumptions elsewhere) — the
    # band-color change alone separates it from INVOICE INFO.
    pay_from = HEADERS.index("Pay Ref #") + 1
    sections = [
        ("BILL INFO",          1,        d1 - 1,                     SUPRA_BILL_FILL),
        ("STATUS & ACTION",    d1 + 1,   d2 - 1,                     SUPRA_STATUS_FILL),
        ("INVOICE INFO",       d2 + 1,   pay_from - 1,               SUPRA_INVOICE_FILL),
        ("PAYMENT",            pay_from, KEY_COL_INDEX - 1,          SUPRA_PAY_FILL),
    ]
    for label, c_from, c_to, fill in sections:
        ws.cell(row=1, column=c_from, value=label)
        ws.merge_cells(start_row=1, start_column=c_from, end_row=1, end_column=c_to)
        # Apply font + fill to every cell in the merged range
        # (Excel needs all cells styled or the merge can de-style on edit)
        for ci in range(c_from, c_to + 1):
            cell = ws.cell(row=1, column=ci)
            cell.fill = fill
            cell.font = SUPRA_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center")
    # Divider columns in row 1 also get the dark fill so the section bands
    # are visually separated by the same divider as the data rows.
    for c_i in DIVIDER_COL_INDEXES:
        ws.cell(row=1, column=c_i).fill = DIVIDER_FILL
    # _Key column in row 1 stays plain (it will be hidden anyway)

    ws.row_dimensions[1].height = SUPRA_HEIGHT

    # Row 2 — column headers (this row is the Table's header).
    for col_i, name in enumerate(HEADERS, 1):
        c = ws.cell(row=2, column=col_i, value=name)
        c.font = HEADER_FONT
        c.alignment = ALIGN_CENTER
        c.fill = DIVIDER_FILL if col_i in DIVIDER_COL_INDEXES else HEADER_FILL
    ws.row_dimensions[2].height = 22

    # Freeze both header rows so they stay visible as user scrolls
    ws.freeze_panes = "A3"

    for col_i, width in COL_WIDTHS.items():
        ws.column_dimensions[_col_letter(col_i)].width = width
    # hide _Key always, plus any sheet-specific hidden columns (Line Amount
    # on Bills, Pay on Inventory).
    ws.column_dimensions[_col_letter(KEY_COL_INDEX)].hidden = True
    for col_i in (hide_cols or []):
        ws.column_dimensions[_col_letter(col_i)].hidden = True


def _notice_date(bill_date) -> Optional[dt.date]:
    """Supplier lien-notice deadline: 15th of the 2nd month after the bill date
    (TX practice). None if the bill date is missing/invalid."""
    if not isinstance(bill_date, dt.date):
        return None
    total = bill_date.year * 12 + (bill_date.month - 1) + 2
    y, m0 = divmod(total, 12)
    return dt.date(y, m0 + 1, 15)


def _notice_timer_label(bill_date, balance, lien_val: str,
                        today: dt.date) -> Optional[str]:
    """Bucket label for an UNPAID, UNTAGGED bill whose lien-notice deadline is
    near or past. None when tagged, paid, no bill date, or >30 days out."""
    if lien_val:
        return None                       # a real tag wins over the countdown
    if not (isinstance(balance, (int, float)) and balance > 0):
        return None                       # paid → no notice risk
    nd = _notice_date(bill_date)
    if nd is None:
        return None
    days = (nd - today).days
    if days < 0:
        return TIMER_LABEL_PAST
    if days <= 7:
        return TIMER_LABEL_7
    if days <= 15:
        return TIMER_LABEL_15
    if days <= 30:
        return TIMER_LABEL_30
    return None


def _write_bill_row(ws, r_i: int, r: dict, edits: Dict[str, Dict[str, str]],
                    today: dt.date) -> None:
    """Render one bill-line row using the unified BILL_ROW_COLS layout.

    Lien and Notes both reflect the preserved value (neither auto-clears — a
    lien persists until manually released, Notes are commentary).

    2026-06-03: dropped Days Open / Bucket / Bill Type. Open link moved to
    col A. Status split into pipeline + Approved columns.
    2026-06-18: Pay column replaced by Lien (Notice Sent / Lien Filed).
    """
    key = r.get("key", "") or ""
    pres = edits.get(key, {"Lien": "", "Notes": ""})
    approved = bool(r.get("approved"))
    pipeline = display_status(r.get("auto_status", ""))
    lien_val = pres.get("Lien", "") or ""
    # AP cash-out: the payment(s) that paid this bill (set in main() for the run).
    _pm = _BILL_PAY_MAP.get(r.get("bill_id", "") or "", {})
    # Untagged + unpaid + near the notice deadline → the cell shows the countdown
    # bucket as text (colored by the matching CF rule). A real tag always wins;
    # this computed text is never preserved (see _read_sheet_edits).
    lien_display = lien_val or (
        _notice_timer_label(r.get("bill_date"), r.get("bill_balance"),
                            lien_val, today) or "")

    values: List[Any] = [
        _qbo_link(r.get("bill_id", "")),                         # 1  Open
        # ── BILL section ──
        r.get("vendor", ""),                                     # 2
        r.get("bill_doc", ""),                                   # 3  Bill #
        r.get("bill_date"),                                      # 4
        r.get("project_num", ""),                                # 5
        r.get("division", ""),                                   # 6
        r.get("account", ""),                                    # 7
        r.get("line_desc", ""),                                  # 8
        r.get("line_amount"),                                    # 9  Line Amount
        r.get("bill_total"),                                     # 10 Bill Total
        r.get("bill_balance"),                                   # 11 Bill Open Bal
        None,                                                    # 12 divider │
        # ── STATUS & ACTION section ──
        pipeline,                                                # 13 Status
        approved_text(approved),                                 # 14 Approved
        lien_display,                                            # 15 Lien (tag or countdown)
        pres.get("Notes", "") or "",                             # 16 Notes
        None,                                                    # 17 divider ┃
        # ── INVOICE section (verification only) ──
        r.get("inv_doc", ""),                                    # 18 Invoice # (→ link)
        _invoice_cell(r.get("inv_doc", ""), r.get("inv_memo", "")),  # 19 Matched Invoice (# + memo)
        r.get("inv_date"),                                       # 20
        r.get("inv_total"),                                      # 21
        r.get("payment_date"),                                   # 22
        r.get("gc_name", ""),                                    # 23 Client (GC/parent)
        # ── PAYMENT section (AP cash-out) ──
        _pm.get("ref", ""),                                      # 24 Pay Ref #
        _pm.get("date"),                                         # 25 Pay Date
        _pm.get("method", ""),                                   # 26 Pay Method
        key,                                                     # 27 _Key
    ]
    for c_i, (val, kind) in enumerate(zip(values, KINDS), start=1):
        c = ws.cell(row=r_i, column=c_i, value=val)
        _format_data_cell(c, kind)

    # Hyperlink the dedicated "Invoice #" cell → the QBO invoice (just the # is
    # the link; the "Matched Invoice" column keeps the # + memo for reading).
    inv_id = r.get("inv_id")
    if inv_id:
        mi = ws.cell(row=r_i, column=HEADERS.index("Invoice #") + 1)
        if mi.value:
            mi.hyperlink = QBO_INVOICE_URL_TEMPLATE.format(inv_id=inv_id)
            mi.font = Font(name="Calibri", size=11, color="0563C1", underline="single")


def _lien_legend(ws, start_col: int) -> None:
    """HORIZONTAL key for the Lien column, on the frozen header row (row 1) to
    the right of the table. A single row so AutoFilter (which hides data rows)
    can never fragment it (Ted 2026-06-18). Each cell is a colored swatch with
    its label inside. Human TAGS (bold black on a saturated fill) come first,
    then the SCRIPT countdown (grey italic on a pale tint) — the same styling
    the data cells get, so the key reads exactly like the column."""
    # (hex fill or None, label, font hex, italic) — italic marks script-written.
    items = [
        (None,             "LIEN KEY →",        "1F3864", False),
        (CF_LIEN_NOTICE,   LIEN_NOTICE,         "000000", False),
        (CF_LIEN_FILED,    LIEN_FILED,          "000000", False),
        (CF_LIEN_RELEASED, LIEN_RELEASED,       "000000", False),
        (CF_TIMER_YELLOW,  TIMER_LABEL_30,      CF_TIMER_TEXT, True),
        (CF_TIMER_ORANGE,  TIMER_LABEL_15,      CF_TIMER_TEXT, True),
        (CF_TIMER_HOT,     TIMER_LABEL_7,       CF_TIMER_TEXT, True),
        (CF_TIMER_PAST,    TIMER_LABEL_PAST,    CF_TIMER_TEXT, True),
    ]
    for i, (hex_color, text, font_hex, italic) in enumerate(items):
        col = start_col + i
        ws.column_dimensions[_col_letter(col)].width = 20
        c = ws.cell(row=1, column=col, value=text)
        c.alignment = Alignment(horizontal="center", vertical="center")
        if hex_color:
            c.fill = PatternFill(patternType="solid",
                                 fgColor=Color(rgb=f"FF{hex_color}"),
                                 bgColor=Color(rgb=f"FF{hex_color}"))
            # Tags read bold; script countdown reads grey italic (not bold).
            c.font = Font(bold=not italic, italic=italic, size=10, color=font_hex)
        else:
            c.font = Font(bold=True, size=11, color=font_hex)


def _finalize_sheet(ws, table_name: str, last_row: int,
                    hide_paid_default: bool = False,
                    lien_editable: bool = False) -> None:
    """Wrap rendered range in an Excel Table (gives AutoFilter on every col)
    and apply universal conditional formatting.

    Tables require ≥1 data row, so we ensure at least row 3 exists even when
    empty (data starts at row 3 since row 1 = supra-header, row 2 = headers).

    `hide_paid_default=True` pre-filters the Status column to hide "Bill paid"
    rows on open — used for the Bills sheet so paid bills don't clutter the
    active view. User clears the filter (Status dropdown → check all) to see
    paid bills when needed.
    """
    HEADER_ROW = 2   # column-name headers (Table header row)
    DATA_START = 3   # data rows begin here

    if last_row < DATA_START:
        # Empty sheet — drop a single blank data row so the Table has body
        for c_i in range(1, len(HEADERS) + 1):
            ws.cell(row=DATA_START, column=c_i, value=None)
        last_row = DATA_START

    table_ref = f"A{HEADER_ROW}:{_col_letter(len(HEADERS))}{last_row}"
    tbl = Table(displayName=table_name, ref=table_ref)
    tbl.tableStyleInfo = TableStyleInfo(
        name="TableStyleLight15",
        showFirstColumn=False, showLastColumn=False,
        showRowStripes=True, showColumnStripes=False,
    )

    # Default filter: hide "Bill paid + …" rows on Bills sheet.
    # FilterColumn.colId is 0-indexed from the first column of the Table.
    # Status sits at table-colId = STATUS_COL_INDEX - 1.
    # Table.autoFilter is None by default — instantiate an AutoFilter first.
    if hide_paid_default:
        # Status is pipeline-only. "Partial paid" added 2026-06-04 — emitted
        # by collapse_rows for multi-project bills whose lines have a mix of
        # funded/unfunded GC invoices. Visible by default (it's an action
        # signal — decide to float or wait).
        visible_statuses = [
            "Invoice paid",
            "Partial paid",
            "Awaiting Payment",
            "Awaiting Invoice",
            "No project #",
        ]
        tbl.autoFilter = AutoFilter(ref=table_ref)
        tbl.autoFilter.filterColumn = [
            FilterColumn(
                colId=STATUS_COL_INDEX - 1,
                filters=Filters(filter=visible_statuses),
            )
        ]
        # openpyxl writes the filter criteria to XML, but Excel doesn't
        # re-evaluate the filter on file open — it just trusts each row's
        # hidden flag. So we also physically hide every row whose Status
        # == "Bill paid" to match the filter on first open. Without this,
        # paid rows show until the user clicks the filter.
        for r in range(DATA_START, last_row + 1):
            if ws.cell(row=r, column=STATUS_COL_INDEX).value == "Bill paid":
                # Exception: a paid bill with an OUTSTANDING lien (Notice Sent
                # or Lien Filed) stays visible — a lien outlives payment until
                # it's Released. Without this, liened-but-paid bills would
                # vanish from the default view and from a Lien-column filter.
                lien_here = ws.cell(row=r, column=LIEN_COL_INDEX).value
                if lien_here in LIEN_OUTSTANDING:
                    continue
                ws.row_dimensions[r].hidden = True
    ws.add_table(tbl)

    # Universal CF — color the row band based on Status × Approved.
    # 2026-06-04: CF ranges EXCLUDE divider columns so the dark bands don't
    # get wiped when a row is tinted. Multi-range syntax = space-separated
    # ranges. Section split = [1..d1-1] [d1+1..d2-1] [d2+1..end]
    dividers = sorted(DIVIDER_COL_INDEXES)
    d1, d2 = dividers[0], dividers[1]
    last_col = _col_letter(len(HEADERS))
    section_ranges = " ".join([
        f"A{DATA_START}:{_col_letter(d1 - 1)}{last_row}",
        f"{_col_letter(d1 + 1)}{DATA_START}:{_col_letter(d2 - 1)}{last_row}",
        f"{_col_letter(d2 + 1)}{DATA_START}:{last_col}{last_row}",
    ])
    status_letter = _col_letter(STATUS_COL_INDEX)
    approved_letter = _col_letter(APPROVED_COL_INDEX)
    lien_letter = _col_letter(LIEN_COL_INDEX)

    # CF formula row anchor — use DATA_START (3) so absolute-row references
    # in $-prefixed formulas point at the first DATA row, not the header.
    fr = DATA_START

    # ── Lien column CF FIRST so it out-prioritizes the row-status band ──
    # The Lien cell sits inside the status section, so the row band would
    # otherwise win and the lien color wouldn't show (Ted 2026-06-18). Adding
    # the lien rules before the row rules gives them higher CF priority: a lien
    # tag or an active timer paints the Lien cell; when neither applies, the row
    # band fills it so the colored band stays continuous.
    if lien_editable:
        lien_range = f"{lien_letter}{DATA_START}:{lien_letter}{last_row}"
        ob_letter = _col_letter(BILL_OPEN_BAL_COL_INDEX)
        unpaid = f"${ob_letter}{fr}>0"
        # Outstanding liens stay colored even when paid (lien outlives payment
        # until formally released).
        ws.conditional_formatting.add(lien_range,
            _cf_fill_rule(f'${lien_letter}{fr}="{LIEN_NOTICE}"', CF_LIEN_NOTICE, bold=True))
        ws.conditional_formatting.add(lien_range,
            _cf_fill_rule(f'${lien_letter}{fr}="{LIEN_FILED}"', CF_LIEN_FILED, bold=True))
        # Released tints only while UNPAID → paid+released archives clean.
        ws.conditional_formatting.add(lien_range,
            _cf_fill_rule(f'AND(${lien_letter}{fr}="{LIEN_RELEASED}",{unpaid})',
                          CF_LIEN_RELEASED, bold=True))
        # Supplier-notice timer: the Lien cell CARRIES the countdown bucket as
        # text (written in _write_bill_row), so color it by matching that text.
        # Script-written → pale tint + GREY ITALIC (vs the human tags above,
        # which are saturated + bold black). stop-if-true so the band below
        # doesn't repaint them.
        for label, color in (
            (TIMER_LABEL_PAST, CF_TIMER_PAST),
            (TIMER_LABEL_7,    CF_TIMER_HOT),
            (TIMER_LABEL_15,   CF_TIMER_ORANGE),
            (TIMER_LABEL_30,   CF_TIMER_YELLOW),
        ):
            ws.conditional_formatting.add(lien_range,
                _cf_fill_rule(f'${lien_letter}{fr}="{label}"', color,
                              italic=True, font_color=CF_TIMER_TEXT,
                              stop_if_true=True))

    # ── Row-status band (added AFTER lien → lower CF priority) ──
    # (status_value, approved_value_or_None_for_any, color)
    cf_rules: List[Tuple[str, Optional[str], str]] = [
        ("Invoice paid",     "approved",      CF_GREEN_READY),     # safe to pay
        ("Invoice paid",     "not approved",  CF_ORANGE_URGENT),   # approve now
        ("Partial paid",     None,            CF_PEACH_PARTIAL),   # decide: float or wait
        ("Awaiting Payment", "not approved",  CF_RED_HOLD),        # blocked
        ("Awaiting Invoice", "not approved",  CF_RED_HOLD),        # blocked
        # Awaiting + approved → no tint (normal pipeline state)
        ("Bill paid",        "approved",      CF_GRAY_DONE),       # done
        ("Bill paid",        "not approved",  CF_YELLOW_AUDIT),    # audit
        ("No project #",     None,            CF_PURPLE_NEEDS_PROJECT),
    ]
    for status_val, approved_val, color in cf_rules:
        if approved_val is None:
            formula = f'${status_letter}{fr}="{status_val}"'
        else:
            formula = (f'AND(${status_letter}{fr}="{status_val}",'
                       f'${approved_letter}{fr}="{approved_val}")')
        ws.conditional_formatting.add(
            section_ranges,
            _cf_fill_rule(formula, color),
        )

    # Lien legend (horizontal, frozen row 1) + dropdown. The lien CF rules
    # themselves were added ABOVE the row-status loop so they out-prioritize
    # the row band on the Lien cell.
    if lien_editable:
        _lien_legend(ws, len(HEADERS) + 2)
        # Dropdown so the tag is foolproof (no typos that would break the CF
        # match). Inline list — no formula2, so the Excel-strict validator
        # passes. allow_blank lets a bill carry no lien tag.
        lien_dv = DataValidation(
            type="list",
            formula1=f'"{LIEN_NOTICE},{LIEN_FILED},{LIEN_RELEASED}"',
            allow_blank=True,
        )
        ws.add_data_validation(lien_dv)
        lien_dv.add(f"{lien_letter}{DATA_START}:{lien_letter}{last_row}")

    # Repaint divider columns dark on every data row (Tables zebra-stripe
    # overrides per-cell fills otherwise; setting on each row is reliable).
    # The CF section_ranges above EXCLUDE divider columns so this fill
    # survives even when the row is tinted by status CF.
    for r_i in range(DATA_START, last_row + 1):
        for c_i in DIVIDER_COL_INDEXES:
            ws.cell(row=r_i, column=c_i).fill = DIVIDER_FILL


def _render_rows(
    ws,
    rows: List[dict],
    table_name: str,
    edits: Dict[str, Dict[str, str]],
    row_filter: Optional[Callable[[dict], bool]] = None,
    sort_key: Optional[Callable[[dict], Tuple]] = None,
    hide_paid_default: bool = False,
    hide_cols: Optional[List[int]] = None,
    inv_anchor_map: Optional[Dict[str, int]] = None,
    lien_editable: bool = False,
) -> Tuple[int, Dict[str, int]]:
    """Apply filter + sort, write rows, finalize as Excel Table.

    Returns (row_count, bill_id_to_row_map).

    Layout: row 1 = supra-header bands, row 2 = column headers (Table
    header row), row 3+ = data. Table.ref starts at row 2.

    `hide_cols` — extra columns to hide (in addition to _Key). Use for
    Bills (hide Line Amount, redundant with Bill Total at bill grain) and
    Inventory (hide Pay — Bills-only decision).

    `inv_anchor_map` — {bill_id → row_on_Inventory}. If provided and a
    row's `is_multi_project` is True, an internal hyperlink is attached to
    the Project # cell so clicking "(multiple)" jumps to that bill's first
    line on the Inventory sheet.
    """
    _apply_header(ws, hide_cols=hide_cols)

    items = [r for r in rows if (row_filter is None or row_filter(r))]
    if sort_key is not None:
        items = sorted(items, key=sort_key)

    today = dt.date.today()
    bill_id_to_row: Dict[str, int] = {}
    proj_col_letter = _col_letter(HEADERS.index("Project #") + 1)
    for r_i, r in enumerate(items, start=3):
        _write_bill_row(ws, r_i, r, edits, today)
        bid = r.get("bill_id")
        if bid and bid not in bill_id_to_row:
            bill_id_to_row[bid] = r_i
        # Hyperlink the "(multiple)" cell on multi-project bills →
        # Inventory anchor row. Cell value stays as the visible text.
        if inv_anchor_map and r.get("is_multi_project") and bid in inv_anchor_map:
            target_row = inv_anchor_map[bid]
            cell = ws[f"{proj_col_letter}{r_i}"]
            cell.hyperlink = f"#Inventory!A{target_row}"
            cell.font = Font(
                name="Calibri", size=11, color="0563C1", underline="single",
            )

    last_row = max(len(items) + 2, 3)
    _finalize_sheet(ws, table_name, last_row, hide_paid_default=hide_paid_default,
                    lien_editable=lien_editable)
    return len(items), bill_id_to_row


# ─────────────────────── sort key ───────────────────────

def _vendor_key(r: dict) -> Tuple:
    """Default sort on the Bills sheet: Vendor → Bill Date → Bill # → Line."""
    return ((r.get("vendor") or "").upper(),
            r.get("bill_date") or dt.date(1900, 1, 1),
            r.get("bill_doc", ""),
            int(r.get("line_id") or 0) if str(r.get("line_id") or "").isdigit() else 0)


# ─────────────────────── Pay List (static, rebuilt each sync) ───────────────────────

# build_pay_list_sheet — removed 2026-05-29.
# Workflow: manager filters Bills sheet by Pay=x (Status filter already hides
# paid bills by default), selects visible rows, copy-pastes to a new workbook
# for AP. Simpler than maintaining a static-snapshot sheet that confused
# people about freshness.


# ─────────────────────── Liens sheet (live register of liened bills) ───────────────────────

# A LIVE view, not a sync-time snapshot. Each row pulls from the Bills sheet
# with SMALL/IF/INDEX formulas, so the instant the clerk tags a lien on Bills
# the row appears here — no waiting for the next sync (like a Notion filtered
# view). Read-only: liens are edited on Bills.
#
# Why not FILTER(): openpyxl can't write a true dynamic-array spill — Excel
# imports it as a single-cell legacy array and nothing spills (verified — the
# sheet came up blank). SMALL/IF/INDEX is the version-proof equivalent: a
# hidden helper column finds the k-th Bills row whose Lien is set (CSE array),
# and each visible cell INDEXes that row out of the Bills sheet. Because it
# reads the Bills rows directly, it also shows liened bills that are paid
# (those rows are kept visible on Bills).

LIENS_MAX_ROWS = 200          # pre-filled live-formula slots (liens are rare)
LIENS_SCAN_END = 5000         # Bills row range each helper scans

# (Liens header, Bills source header). INDEX pulls each column live from Bills.
LIENS_COLS = [
    ("Vendor",        "Vendor"),
    ("Bill #",        "Bill #"),
    ("Bill Date",     "Bill Date"),
    ("Project #",     "Project #"),
    ("Division",      "Division"),
    ("Client",        "Client"),
    ("Bill Total",    "Bill Total"),
    ("Bill Open Bal", "Bill Open Bal"),
    ("Status",        "Status"),
    ("Lien",          "Lien"),
    ("Notes",         "Notes"),
]
_LIENS_KIND = {"Bill Date": "date", "Bill Total": "money", "Bill Open Bal": "money"}
_LIENS_WIDTH = {"Vendor": 26, "Bill #": 12, "Bill Date": 11, "Project #": 12,
                "Division": 9, "Client": 26, "Bill Total": 13, "Bill Open Bal": 13,
                "Status": 16, "Lien": 14, "Notes": 34}


def build_liens_sheet(ws) -> None:
    """Live register of every bill whose Lien tag is set (SMALL/IF/INDEX pulls
    from Bills). Updates on edit, no sync required."""
    n = len(LIENS_COLS)
    help_letter = _col_letter(n + 1)                 # hidden helper column
    bills_lien = _col_letter(LIEN_COL_INDEX)         # Lien column letter on Bills
    lien_out_letter = _col_letter(
        next(i + 1 for i, (h, _) in enumerate(LIENS_COLS) if h == "Lien"))

    # Header + widths.
    for c_i, (h, _) in enumerate(LIENS_COLS, 1):
        c = ws.cell(row=1, column=c_i, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = ALIGN_CENTER
        ws.column_dimensions[_col_letter(c_i)].width = _LIENS_WIDTH.get(h, 14)
    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 22
    ws.column_dimensions[help_letter].hidden = True

    last = LIENS_MAX_ROWS + 1
    scan = (f"Bills!${bills_lien}$3:${bills_lien}${LIENS_SCAN_END}")
    for r in range(2, last + 1):
        # Helper (CSE array): worksheet-row of the k-th Bills bill with a lien,
        # where k = ROW()-1. Blank once we run past the last lien.
        ws.cell(row=r, column=n + 1).value = ArrayFormula(
            ref=f"{help_letter}{r}",
            text=(f'=IFERROR(SMALL(IF({scan}<>"",ROW({scan})),ROW()-1),"")'),
        )
        for c_i, (h, src) in enumerate(LIENS_COLS, 1):
            src_letter = _col_letter(HEADERS.index(src) + 1)
            kind = _LIENS_KIND.get(h)
            idx = f'INDEX(Bills!{src_letter}:{src_letter},${help_letter}{r})'
            # Text columns: append &"" so an empty source cell shows blank, not
            # the 0 that a bare INDEX returns. Money/date stay numeric so a real
            # 0 / date isn't turned into text.
            expr = f'{idx}&""' if kind is None else idx
            cell = ws.cell(
                row=r, column=c_i,
                value=f'=IF(${help_letter}{r}="","",{expr})',
            )
            cell.font = BODY_FONT
            if kind == "money":
                cell.number_format = MONEY_FMT
                cell.alignment = ALIGN_RIGHT
            elif kind == "date":
                cell.number_format = DATE_FMT
                cell.alignment = ALIGN_CENTER

    # AutoFilter so AP can slice by Division / Lien within the register.
    ws.auto_filter.ref = f"A1:{_col_letter(n)}{last}"

    # Lien cell tint (amber / red / green) over the slot range.
    lien_range = f"{lien_out_letter}2:{lien_out_letter}{last}"
    for val, color in ((LIEN_NOTICE, CF_LIEN_NOTICE),
                       (LIEN_FILED, CF_LIEN_FILED),
                       (LIEN_RELEASED, CF_LIEN_RELEASED)):
        ws.conditional_formatting.add(
            lien_range,
            _cf_fill_rule(f'${lien_out_letter}2="{val}"', color, bold=True),
        )


# ─────────────────────── Inventory sheet (multi-project drill-down) ───────────────────────

# Multi-project bills (Martin Marietta-style inventory tickets that distribute
# across jobs) appear on the master Bills sheet as ONE summary row with
# Project# = "(multiple)". The per-line / per-project detail lives here so AP
# can see which jobs ate which dollars + each line's individual GC-invoice
# status (so the bill-level Status on the Bills sheet can roll up to
# "Partial paid" when some lines are funded and others aren't).
#
# 2026-06-04: Inventory now mirrors the Bills sheet layout exactly (same 22
# columns, same supra-headers, same CF). Pay column is hidden — that decision
# lives on the Bills sheet. Line Amount is visible here (the whole point —
# per-line dollar attribution).

INVENTORY_TABLE = "tblInventory"


def build_inventory_sheet(ws, line_rows: List[dict]
                          ) -> Tuple[int, Dict[str, int]]:
    """Render multi-project bills' line-level rows using the same layout as
    the master Bills sheet, with the Pay column hidden.

    Returns (line_count, {bill_id → first_row_on_inventory}) so the Bills
    renderer can attach hyperlinks from each "(multiple)" cell back to the
    matching anchor row here.

    Sort: Vendor → Bill Date → Bill # → Project # — keeps every bill's lines
    grouped together so the hyperlink lands you at the FIRST line of the
    bill and the rest sit immediately below.
    """
    multi_ids = multi_project_bill_ids(line_rows)
    items = [r for r in line_rows if r.get("bill_id") in multi_ids]

    sort_key = lambda r: (
        (r.get("vendor") or "").upper(),
        r.get("bill_date") or dt.date(1900, 1, 1),
        r.get("bill_doc", ""),
        (r.get("project_num") or ""),
    )

    # Inventory hides Pay (col 15) — Bills sheet is the source of truth for
    # the pay decision. _Key is always hidden by _apply_header.
    n, bill_id_to_row = _render_rows(
        ws,
        items,
        INVENTORY_TABLE,
        edits={},                       # no Lien/Notes preservation here
        sort_key=sort_key,
        hide_paid_default=True,         # same default as Bills
        hide_cols=[LIEN_COL_INDEX],
        inv_anchor_map=None,            # this IS the inventory sheet
    )
    return n, bill_id_to_row


# ─────────────────────── Bill payments (AP cash-out) ───────────────────────
# Rather than a separate sheet, the payment that paid each bill is shown as
# Pay Ref # / Pay Date / Pay Method columns on the Bills sheet, so the clerk can
# just filter by check # or pay date (Ted 2026-07-13). No out-of-window bill
# fetch and no second sheet — we only annotate bills already loaded. The map is
# built once per run and read by _write_bill_row via this module global.
_BILL_PAY_MAP: Dict[str, dict] = {}


def _bp_method(bp: dict) -> str:
    """Normalize a BillPayment's PayType to a short label the clerk reads as
    'which statement do I open'."""
    pt = (bp.get("PayType") or "").strip().lower()
    if pt == "check":
        return "Check"
    if pt in ("creditcard", "cc"):
        return "CC"
    if pt == "cash":
        return "Cash"
    return (bp.get("PayType") or "?")


def _bp_linked_bills(bp: dict) -> List[Tuple[str, Optional[float]]]:
    """Return [(bill_id, amount_applied)] for a BillPayment. QBO puts the link
    either on each Line (with the applied Amount) or at the top level (no
    amount). Line-level wins; top-level fills any gap with amount None."""
    seen: Dict[str, Optional[float]] = {}
    for ln in bp.get("Line") or []:
        amt = float(ln.get("Amount") or 0)
        for lk in ln.get("LinkedTxn") or []:
            if lk.get("TxnType") == "Bill" and lk.get("TxnId"):
                seen[lk["TxnId"]] = amt
    for lk in bp.get("LinkedTxn") or []:
        if lk.get("TxnType") == "Bill" and lk.get("TxnId") and lk["TxnId"] not in seen:
            seen[lk["TxnId"]] = None
    return list(seen.items())


def build_bill_payment_map(bill_payments: List[dict]) -> Dict[str, dict]:
    """Map each paid bill_id -> {ref, date, method} from its BillPayment(s).
    Checks show the check # (DocNumber); CC refs are left blank (a generic
    label is no lookup key -- pay date + vendor + amount locate it). A bill paid
    by more than one payment shows the latest date and '(multiple)'."""
    acc: Dict[str, List[Tuple[Optional[dt.date], str, str]]] = defaultdict(list)
    for bp in bill_payments:
        pay_date = parse_date(bp.get("TxnDate"))
        method = _bp_method(bp)
        ref = (bp.get("DocNumber") or "").strip() if method == "Check" else ""
        for bid, _amt in _bp_linked_bills(bp):
            acc[bid].append((pay_date, method, ref))
    out: Dict[str, dict] = {}
    for bid, lst in acc.items():
        lst.sort(key=lambda t: t[0] or dt.date.min)
        if len(lst) == 1:
            d, m, rf = lst[0]
            out[bid] = {"date": d, "method": m, "ref": rf}
        else:
            methods = {t[1] for t in lst}
            out[bid] = {
                "date": lst[-1][0],
                "method": methods.pop() if len(methods) == 1 else "(multiple)",
                "ref": "(multiple)",
            }
    return out

# ─────────────────────── QBO Audit sheet ───────────────────────

# QBO Class field uses spelled-out division names, not the RP/CP/MFD codes.
# Normalize before comparing in the audit. Anything not in this map gets the
# raw class string echoed back (so the audit message says exactly what's in
# QBO when there's a real mismatch — easier to spot bad data entry).
CLASS_TO_DIVISION = {
    "RESIDENTIAL": "RP",
    "RP":          "RP",
    "RES":         "RP",
    "COMMERCIAL":  "CP",
    "CP":          "CP",
    "COM":         "CP",
    "MULTI FAMILY":  "MFD",
    "MULTIFAMILY":   "MFD",
    "MULTI-FAMILY":  "MFD",
    "MFD":           "MFD",
    "MF":            "MFD",
}


def _normalize_class(class_name: str) -> Optional[str]:
    """Map a QBO Class value to a division code (RP/CP/MFD), or None if
    unrecognized. Case- and whitespace-insensitive."""
    if not class_name:
        return None
    key = " ".join(class_name.strip().upper().split())  # collapse internal spaces
    return CLASS_TO_DIVISION.get(key)


AUDIT_HEADERS = [
    "Bill #", "Vendor", "Bill Date", "Customer/Project (QBO)",
    "Project # (parsed)", "Division (expected)", "Class field (QBO)",
    "Memo (PrivateNote)", "Line Description",
    "Mismatch", "Open",
]

# 2026-06-04: rolling buffer for the stale-not-approved audit. NOT APPROVED
# bills older than this many days surface in their own audit section so the
# clerk can chase approvals before they age further. 30 days mirrors the
# old end-of-month batch cadence and self-corrects past month boundaries
# (a bill from May 30 doesn't become urgent until ~June 29, not June 1).
NOT_APPROVED_BUFFER_DAYS = 30

# Audit section banner styling — distinct color per section so the two
# concern types (approval pipeline vs data entry) read as separate work.
AUDIT_SECTION_STALE_FILL = PatternFill(patternType="solid",
                                       fgColor=Color(rgb="FFFFCC66"),
                                       bgColor=Color(rgb="FFFFCC66"))
AUDIT_SECTION_DATA_FILL  = PatternFill(patternType="solid",
                                       fgColor=Color(rgb="FFD9E1F2"),
                                       bgColor=Color(rgb="FFD9E1F2"))
AUDIT_SECTION_FONT       = Font(name="Calibri", size=12, bold=True, color="1F3864")
AUDIT_SECTION_HEIGHT     = 24

# Section 3 — uncoded JOB-COST lines (a real cost about to be left off a
# project). Strong orange so it reads as the most actionable section. Overhead-
# only uncoded lines are deliberately NOT shown (Ted 2026-06-18: they'd flood
# the audit and are legitimately projectless). Sub bills are already excluded
# upstream, so they don't reach here either.
AUDIT_SECTION_UNCODED_FILL = PatternFill(patternType="solid",
                                         fgColor=Color(rgb="FFF4B183"),
                                         bgColor=Color(rgb="FFF4B183"))
AUDIT_UNCODED_CELL_FILL    = PatternFill(patternType="solid",
                                         fgColor=Color(rgb="FFFCE4D6"),
                                         bgColor=Color(rgb="FFFCE4D6"))

# Age escalation on the Bill Date cell in the MISSING-project section. A blank
# project # is tolerable while the job is still being identified, but a real
# job cost shouldn't sit uncoded for long — the project is knowable within ~2
# weeks. >15 days old → yellow, >30 days → red. (Ted 2026-07-10.)
AUDIT_MISSING_AGE_YELLOW_DAYS = 15
AUDIT_MISSING_AGE_RED_DAYS    = 30
AUDIT_AGE_YELLOW_FILL = PatternFill(patternType="solid",
                                    fgColor=Color(rgb="FFFFE699"),
                                    bgColor=Color(rgb="FFFFE699"))
AUDIT_AGE_RED_FILL    = PatternFill(patternType="solid",
                                    fgColor=Color(rgb="FFFF9999"),
                                    bgColor=Color(rgb="FFFF9999"))

# Project-code pattern for the uncoded job-cost check (own constant so the
# existing _audit_row_checks stays untouched).
_UNCODED_PROJ_RE = re.compile(r"\b(MFD|CP|RP)\d+(?:-FTW)?\b", re.IGNORECASE)

# Aging bucket sub-banner inside the stale section. Lighter amber than the
# main section banner — reads as a sub-heading, not a peer break.
AUDIT_BUCKET_FILL        = PatternFill(patternType="solid",
                                       fgColor=Color(rgb="FFFFE5B0"),
                                       bgColor=Color(rgb="FFFFE5B0"))
AUDIT_BUCKET_FONT        = Font(name="Calibri", size=11, bold=True, color="1F3864")
AUDIT_BUCKET_HEIGHT      = 20

# Section 4 — duplicate bill # within a vendor tree (double-entry / double-pay
# risk). Purple family so it reads as its own concern, distinct from the amber
# approval section and blue data-entry section. Always rendered — even at zero,
# the clerk sees the check ran and found nothing.
AUDIT_SECTION_DUP_FILL   = PatternFill(patternType="solid",
                                       fgColor=Color(rgb="FFB1A0C7"),
                                       bgColor=Color(rgb="FFB1A0C7"))
AUDIT_DUP_GROUP_FILL     = PatternFill(patternType="solid",
                                       fgColor=Color(rgb="FFCCC0DA"),
                                       bgColor=Color(rgb="FFCCC0DA"))
AUDIT_DUP_CELL_FILL      = PatternFill(patternType="solid",
                                       fgColor=Color(rgb="FFE4DFEC"),
                                       bgColor=Color(rgb="FFE4DFEC"))

# Shared cell-kind list for every audit data row (all sections use the same
# 11-column AUDIT_HEADERS layout).
AUDIT_ROW_KINDS = ["text", "text", "date", "text", "text", "text", "text",
                   "text", "text", "text", "link"]

# Aging buckets — most overdue first, so the clerk's eye lands on the worst
# offenders. (label, lower_inclusive, upper_exclusive)
AGING_BUCKETS: List[Tuple[str, int, int]] = [
    ("90+ days",   90, 10**9),
    ("60–90 days", 60, 90),
    ("30–60 days", 30, 60),
]


def _aging_bucket_for(days_old: int) -> str:
    """Map a days-old value to its aging-bucket label."""
    for label, lo, hi in AGING_BUCKETS:
        if lo <= days_old < hi:
            return label
    return AGING_BUCKETS[0][0]  # 90+ catches anything past the top


def _stale_not_approved_bills(
    line_rows: List[dict],
    today: dt.date,
    threshold_days: int = NOT_APPROVED_BUFFER_DAYS,
) -> List[dict]:
    """Return bills that are NOT APPROVED and aged past `threshold_days`.

    Deduped by bill_id — approval is bill-level, so one entry per bill even
    when the bill has many lines. Sorted oldest-first so the most overdue
    bills sit at the top of the audit section.

    `today` is passed in by the caller so the date stays consistent across
    the entire audit pass.
    """
    by_bill: Dict[str, dict] = {}
    for r in line_rows:
        bid = r.get("bill_id") or ""
        if not bid or bid in by_bill:
            continue
        if r.get("approved", True):
            continue  # bill IS approved — not in our scope
        bill_date = r.get("bill_date")
        if not isinstance(bill_date, dt.date):
            continue
        days_old = (today - bill_date).days
        if days_old < threshold_days:
            continue
        by_bill[bid] = {
            "bill_id": bid,
            "bill_doc": r.get("bill_doc") or "",
            "vendor": r.get("vendor") or "",
            "bill_date": bill_date,
            "customer_name": r.get("customer_name") or "",
            "project_num": r.get("project_num") or "",
            "division": r.get("division") or "",
            "class_name": r.get("class_name") or "",
            "days_old": days_old,
        }
    return sorted(by_bill.values(), key=lambda x: -x["days_old"])


def _audit_section_banner(ws, row_idx: int, text: str,
                          fill: PatternFill, n_cols: int) -> None:
    """Render a merged section banner row on the audit sheet."""
    ws.cell(row=row_idx, column=1, value=text)
    ws.merge_cells(start_row=row_idx, start_column=1,
                   end_row=row_idx, end_column=n_cols)
    for c_i in range(1, n_cols + 1):
        cell = ws.cell(row=row_idx, column=c_i)
        cell.fill = fill
        cell.font = AUDIT_SECTION_FONT
    ws.cell(row=row_idx, column=1).alignment = Alignment(
        horizontal="left", vertical="center", indent=1
    )
    ws.row_dimensions[row_idx].height = AUDIT_SECTION_HEIGHT


def _audit_row_checks(r: dict) -> List[str]:
    """Run all audit checks for a single bill-line row. Return list of
    mismatch labels (empty list = no problems)."""
    issues: List[str] = []
    project_num = (r.get("project_num") or "").upper()
    expected_div = (r.get("division") or "").upper()
    actual_class = (r.get("class_name") or "").strip()

    # 1) Class field empty — only meaningful if project IS assigned
    if expected_div and not actual_class:
        issues.append("Class not set")

    # 2) Class set but normalizes to wrong division.
    #    QBO Class is spelled-out ("Residential" / "Commercial" / "Multi Family"),
    #    not the RP/CP/MFD codes — normalize via CLASS_TO_DIVISION first.
    if expected_div and actual_class:
        normalized = _normalize_class(actual_class)
        if normalized is None:
            # Class set but unrecognized — clerk typed something off-script
            issues.append(f"Class={actual_class!r} (not a recognized division)")
        elif normalized != expected_div:
            issues.append(f"Class={actual_class!r} ({normalized}) but project is {expected_div}")

    # 3) Line description has a different project # than Customer/Project
    line_desc = r.get("line_desc", "") or ""
    desc_match = re.search(r"\b(MFD|CP|RP)\d+(?:-FTW)?\b", line_desc, re.IGNORECASE)
    if desc_match:
        desc_proj = desc_match.group(0).upper()
        if project_num and desc_proj != project_num:
            issues.append(f"Line desc says {desc_proj} but project is {project_num}")
    return issues


def _uncoded_job_cost(r: dict) -> Optional[Tuple[str, str]]:
    """Guardrail: a line with NO project that shows job-cost signals is a real
    miss (a cost about to be left off a project). Return (reason, division) or
    None.

    Signals (high-confidence job work):
      1. item / cost-code line (bill_type == "COGS"), or
      2. Class set to a TOP-PARENT division class (_normalize_class → RP/CP/MFD;
         sub-classes / off-script classes don't count), or
      3. line description carries a project code (MFD/CP/RP####).

    Lines with no signal are treated as legitimate overhead and intentionally
    NOT flagged — that's the "don't flood the audit" guardrail (Ted 2026-06-18).
    """
    if (r.get("project_num") or "").strip():
        return None  # has a project — covered by _audit_row_checks, not here

    is_item   = r.get("bill_type") == "COGS"
    class_div = _normalize_class(r.get("class_name") or "")   # top-parent division only
    desc      = r.get("line_desc") or ""
    code_m    = _UNCODED_PROJ_RE.search(desc)
    if not (is_item or class_div or code_m):
        return None  # no job-cost signal → overhead, skip

    cust = (r.get("customer_name") or "").strip()
    reason = (f"Missing project # (parent only: {cust})" if cust
              else "Missing Customer/Project")
    hints = []
    if is_item:
        hints.append("item/cost-code line")
    if class_div:
        hints.append(f"class={class_div}")
    if code_m:
        hints.append(f"desc says {code_m.group(0).upper()}")
    if hints:
        reason += " — " + ", ".join(hints)
    return reason, (class_div or "(none)")


def _norm_ref(doc: str) -> str:
    """Normalize a bill ref # for duplicate matching: trim + uppercase."""
    return (doc or "").strip().upper()


# Credit-card-fee bills reuse a generic label ("CC", "CC FEE", "MONTHLY CC FEE")
# as their ref #, so the same label recurs across many unrelated bills and dates
# — not real duplicates. A CC-marker ref only counts as a duplicate when the
# copies also land on the SAME DAY (and same vendor), i.e. a genuine same-day
# double entry. Ted 2026-07-10.
_CC_REF_RE = re.compile(r"\bCC\b")


def _is_cc_ref(ref_key: str) -> bool:
    """True if a normalized ref # is a credit-card-fee marker (has 'CC' as a
    standalone token)."""
    return bool(_CC_REF_RE.search(ref_key))


def _build_vendor_root(vendors: List[dict]) -> Dict[str, str]:
    """Map each vendor Id to its top-most ancestor Id by walking ParentRef
    (cycle-guarded). Root + every sub-vendor collapse to one tree key, so a
    duplicate ref # booked once to a parent and once to a sub-vendor is still
    caught. A vendor with no parent maps to itself."""
    parent: Dict[str, str] = {}
    ids = set()
    for v in vendors:
        vid = v.get("Id")
        if not vid:
            continue
        ids.add(vid)
        pid = (v.get("ParentRef") or {}).get("value")
        if pid and pid != vid:
            parent[vid] = pid

    def root_of(vid: str) -> str:
        seen = set()
        cur = vid
        while cur in parent and cur not in seen:
            seen.add(cur)
            cur = parent[cur]
        return cur

    return {vid: root_of(vid) for vid in ids}


def _duplicate_bill_groups(
    rows: List[dict],
    vendor_root: Optional[Dict[str, str]] = None,
    vendor_map: Optional[Dict[str, str]] = None,
) -> List[List[dict]]:
    """Group bills (deduped by bill_id) by (vendor tree root, normalized ref #)
    and return only the groups with 2+ distinct bills — the same ref # entered
    more than once under one vendor tree. Blank ref #s are skipped (nothing to
    compare). Groups sorted by tree name then ref #; bills within a group by
    date."""
    vendor_root = vendor_root or {}
    vendor_map = vendor_map or {}
    bills: Dict[str, dict] = {}
    for r in rows:
        bid = r.get("bill_id") or ""
        if not bid or bid in bills:
            continue
        doc_key = _norm_ref(r.get("bill_doc"))
        if not doc_key:
            continue
        vid = r.get("vendor_id") or ""
        root_id = vendor_root.get(vid, vid)
        bills[bid] = {
            "bill_id": bid,
            "bill_doc": r.get("bill_doc") or "",
            "vendor": r.get("vendor") or "",
            "root_id": root_id,
            "root_name": vendor_map.get(root_id, "") or (r.get("vendor") or ""),
            "bill_date": r.get("bill_date"),
            "bill_total": float(r.get("bill_total") or 0),
            "customer_name": r.get("customer_name") or "",
            "doc_key": doc_key,
        }
    groups: Dict[Tuple[str, str, str], List[dict]] = defaultdict(list)
    for b in bills.values():
        # Fall back to vendor name when we have no id (defensive — a row should
        # always carry vendor_id, but never merge two vendors on an empty key).
        tree_key = b["root_id"] or ("NAME:" + b["vendor"].upper())
        # CC-fee markers only group within the same day (same-day double entry);
        # a real ref # groups across all dates (date_part stays empty).
        date_part = ""
        if _is_cc_ref(b["doc_key"]):
            date_part = b["bill_date"].isoformat() if b["bill_date"] else "NODATE"
        groups[(tree_key, b["doc_key"], date_part)].append(b)
    out = [g for g in groups.values() if len(g) >= 2]
    for g in out:
        g.sort(key=lambda x: (x["bill_date"] or dt.date.min, x["bill_doc"]))
    out.sort(key=lambda g: ((g[0]["root_name"] or g[0]["vendor"] or "").upper(),
                            g[0]["doc_key"]))
    return out


def _audit_write_row(ws, row_idx: int, values: List, tint_fill: Optional[PatternFill],
                     level: int) -> None:
    """Write one 11-column audit data row, tint the Mismatch cell (col 10), and
    place it at the given outline level so the section is collapsible."""
    for c_i, (val, kind) in enumerate(zip(values, AUDIT_ROW_KINDS), start=1):
        c = ws.cell(row=row_idx, column=c_i, value=val)
        _format_data_cell(c, kind)
    if tint_fill is not None:
        ws.cell(row=row_idx, column=10).fill = tint_fill
    ws.row_dimensions[row_idx].outline_level = level


def _audit_sub_banner(ws, row_idx: int, text: str, fill: PatternFill,
                      n_cols: int, level: int) -> None:
    """Render a nested sub-group banner (aging bucket, duplicate group) at the
    given outline level."""
    ws.cell(row=row_idx, column=1, value=text)
    ws.merge_cells(start_row=row_idx, start_column=1,
                   end_row=row_idx, end_column=n_cols)
    for c_i in range(1, n_cols + 1):
        cell = ws.cell(row=row_idx, column=c_i)
        cell.fill = fill
        cell.font = AUDIT_BUCKET_FONT
    ws.cell(row=row_idx, column=1).alignment = Alignment(
        horizontal="left", vertical="center", indent=2
    )
    ws.row_dimensions[row_idx].height = AUDIT_BUCKET_HEIGHT
    ws.row_dimensions[row_idx].outline_level = level


def _audit_none_row(ws, row_idx: int) -> None:
    """Placeholder row shown when a section has zero findings — the header still
    renders so the clerk sees at a glance that the check ran and was clean."""
    c = ws.cell(row=row_idx, column=1, value="✓ none")
    c.font = Font(italic=True, color="808080")
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[row_idx].outline_level = 1


def _missing_project_age_fill(bill_date, today: dt.date) -> Optional[PatternFill]:
    """Escalate the Bill Date cell of a MISSING-project line by age: yellow past
    AUDIT_MISSING_AGE_YELLOW_DAYS, red past AUDIT_MISSING_AGE_RED_DAYS. A recent
    bill (project still being identified) stays uncolored. None = no fill."""
    if not isinstance(bill_date, dt.date):
        return None
    days_old = (today - bill_date).days
    if days_old > AUDIT_MISSING_AGE_RED_DAYS:
        return AUDIT_AGE_RED_FILL
    if days_old > AUDIT_MISSING_AGE_YELLOW_DAYS:
        return AUDIT_AGE_YELLOW_FILL
    return None


def build_audit_sheet(ws, rows: List[dict],
                      vendor_root: Optional[Dict[str, str]] = None,
                      vendor_map: Optional[Dict[str, str]] = None) -> int:
    """Audit sheet — four sections, each under a collapsible outline group so
    the clerk can collapse to the section banners and read the whole audit at a
    glance:
      1. Stale NOT APPROVED bills (bill-grain, deduped), sub-grouped by aging
         bucket. Chase approvals for any bill older than NOT_APPROVED_BUFFER_DAYS.
      2. Data entry issues (line-grain). Empty Class, Class/division mismatch,
         line-desc project mismatch, etc.
      3. Likely job cost missing a project (line-grain, high-confidence only).
      4. Duplicate bill # within a vendor tree (bill-grain) — double-entry /
         double-pay risk.

    Every section renders its banner + count even at zero (a "✓ none" row),
    so no check silently disappears. All sections share the AUDIT_HEADERS
    column layout. Returns total flagged-entry count.
    """
    today = dt.date.today()
    n_cols = len(AUDIT_HEADERS)

    # Outline grouping: banners are summaries that sit ABOVE their detail rows,
    # so the collapse (+/−) control lands on the banner, not below the group.
    ws.sheet_properties.outlinePr.summaryBelow = False
    ws.sheet_properties.outlinePr.summaryRight = False

    # Header row
    for col_i, name in enumerate(AUDIT_HEADERS, 1):
        c = ws.cell(row=1, column=col_i, value=name)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = ALIGN_CENTER
    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 22
    widths = {1: 12, 2: 28, 3: 11, 4: 40, 5: 12, 6: 9, 7: 14,
              8: 40, 9: 35, 10: 50, 11: 9}
    for col, w in widths.items():
        ws.column_dimensions[_col_letter(col)].width = w

    # Collect all sections' work
    stale_bills = _stale_not_approved_bills(rows, today)
    flagged: List[Tuple[dict, List[str]]] = []
    uncoded: List[Tuple[dict, str, str]] = []   # (row, reason, division)
    for r in rows:
        issues = _audit_row_checks(r)
        if issues:
            flagged.append((r, issues))
        uc = _uncoded_job_cost(r)               # uncoded lines never have a project,
        if uc:                                  # so they can't also appear in `flagged`
            uncoded.append((r, uc[0], uc[1]))
    flagged.sort(key=lambda x: ((x[0].get("vendor") or "").upper(),
                                x[0].get("bill_doc", "")))
    uncoded.sort(key=lambda x: ((x[0].get("vendor") or "").upper(),
                                x[0].get("bill_doc", "")))
    dup_groups = _duplicate_bill_groups(rows, vendor_root, vendor_map)

    print(f"  audit: {len(stale_bills)} stale, {len(flagged)} data-entry, "
          f"{len(uncoded)} uncoded job-cost, {len(dup_groups)} duplicate-ref group(s)")

    # Amber tint for the Mismatch cell — matches the stale section banner.
    stale_tint = PatternFill(patternType="solid",
                             fgColor=Color(rgb="FFFFCC66"), bgColor=Color(rgb="FFFFCC66"))
    data_tint = PatternFill(patternType="solid",
                            fgColor=Color(rgb="FFFCE4E4"), bgColor=Color(rgb="FFFCE4E4"))

    cur_row = 2

    # ─── Section 1: Stale NOT APPROVED bills ───
    # Sub-grouped by aging bucket (90+ → 60–90 → 30–60), vendor-first within
    # each bucket so the clerk works one vendor at a time without flipping
    # back and forth across the list.
    _audit_section_banner(
        ws, cur_row,
        f"⚠  NOT APPROVED bills > {NOT_APPROVED_BUFFER_DAYS} days old  "
        f"({len(stale_bills)} bill{'s' if len(stale_bills) != 1 else ''})",
        AUDIT_SECTION_STALE_FILL,
        n_cols,
    )
    cur_row += 1
    if not stale_bills:
        _audit_none_row(ws, cur_row)
        cur_row += 1
    else:
        by_bucket: Dict[str, List[dict]] = defaultdict(list)
        for s in stale_bills:
            by_bucket[_aging_bucket_for(s["days_old"])].append(s)

        # Render buckets in canonical order (most overdue first)
        for bucket_label, _, _ in AGING_BUCKETS:
            bucket_bills = by_bucket.get(bucket_label, [])
            if not bucket_bills:
                continue
            # Vendor-first within bucket; oldest first for a vendor's own group
            bucket_bills.sort(key=lambda x: (
                (x["vendor"] or "").upper(),
                -x["days_old"],
                x["bill_doc"],
            ))
            _audit_sub_banner(
                ws, cur_row,
                f"   {bucket_label}  ({len(bucket_bills)} bill"
                f"{'s' if len(bucket_bills) != 1 else ''})",
                AUDIT_BUCKET_FILL, n_cols, level=1,
            )
            cur_row += 1
            for s in bucket_bills:
                values = [
                    s["bill_doc"], s["vendor"], s["bill_date"], s["customer_name"],
                    s["project_num"] or "(none)", s["division"] or "(none)",
                    "",                       # class — bill-level, n/a
                    "NOT APPROVED",
                    "",                       # line desc — bill-level, n/a
                    f"NOT APPROVED — {s['days_old']} days old",
                    _qbo_link(s["bill_id"]),
                ]
                _audit_write_row(ws, cur_row, values, stale_tint, level=2)
                cur_row += 1
    cur_row += 1   # spacer row between sections

    # ─── Section 2: Data entry issues ───
    _audit_section_banner(
        ws, cur_row,
        f"Data entry issues  ({len(flagged)} flagged line"
        f"{'s' if len(flagged) != 1 else ''})",
        AUDIT_SECTION_DATA_FILL,
        n_cols,
    )
    cur_row += 1
    if not flagged:
        _audit_none_row(ws, cur_row)
        cur_row += 1
    else:
        for r, issues in flagged:
            values = [
                r.get("bill_doc", ""), r.get("vendor", ""), r.get("bill_date"),
                r.get("customer_name", ""),
                r.get("project_num", "") or "(none)",
                r.get("division", "") or "(none)",
                r.get("class_name", "") or "(empty)",
                "",                       # memo placeholder
                r.get("line_desc", ""),
                " · ".join(issues),
                _qbo_link(r.get("bill_id", "")),
            ]
            _audit_write_row(ws, cur_row, values, data_tint, level=1)
            cur_row += 1
    cur_row += 1   # spacer row between sections

    # ─── Section 3: Uncoded JOB-COST lines (missing project) ───
    # Only high-confidence job-cost misses — overhead-only uncoded lines are
    # skipped so the audit isn't flooded (Ted 2026-06-18).
    _audit_section_banner(
        ws, cur_row,
        f"⛔  Likely job cost — MISSING project  ({len(uncoded)} line"
        f"{'s' if len(uncoded) != 1 else ''})",
        AUDIT_SECTION_UNCODED_FILL,
        n_cols,
    )
    cur_row += 1
    if not uncoded:
        _audit_none_row(ws, cur_row)
        cur_row += 1
    else:
        for r, reason, division in uncoded:
            values = [
                r.get("bill_doc", ""), r.get("vendor", ""), r.get("bill_date"),
                r.get("customer_name", "") or "(none)",
                "(none)",                                  # Project # (parsed)
                division,                                  # Division (implied by class)
                r.get("class_name", "") or "(empty)",
                "",                                        # memo placeholder
                r.get("line_desc", ""),
                reason,
                _qbo_link(r.get("bill_id", "")),
            ]
            _audit_write_row(ws, cur_row, values, AUDIT_UNCODED_CELL_FILL, level=1)
            # Age-escalate the Bill Date cell (col 3): a project # left blank
            # because it's unknown should be resolved within ~2 weeks.
            age_fill = _missing_project_age_fill(r.get("bill_date"), today)
            if age_fill is not None:
                ws.cell(row=cur_row, column=3).fill = age_fill
            cur_row += 1
    cur_row += 1   # spacer row between sections

    # ─── Section 4: Duplicate bill # within a vendor tree ───
    # Same ref # on 2+ bills under one vendor tree = double-entry / double-pay
    # risk. Grouped per (tree, ref #); matching totals are the strongest signal.
    n_dup_bills = sum(len(g) for g in dup_groups)
    _audit_section_banner(
        ws, cur_row,
        f"🔁  Duplicate bill # in a vendor tree  ({len(dup_groups)} group"
        f"{'s' if len(dup_groups) != 1 else ''})",
        AUDIT_SECTION_DUP_FILL,
        n_cols,
    )
    cur_row += 1
    if not dup_groups:
        _audit_none_row(ws, cur_row)
        cur_row += 1
    else:
        for g in dup_groups:
            totals = {round(b["bill_total"], 2) for b in g}
            same_amt = len(totals) == 1
            amt_note = (f"same amount ${next(iter(totals)):,.2f}" if same_amt
                        else "AMOUNTS DIFFER")
            root_name = g[0]["root_name"] or g[0]["vendor"]
            _audit_sub_banner(
                ws, cur_row,
                f"   #{g[0]['bill_doc']}  ·  {root_name}  ·  {len(g)} bills  ·  {amt_note}",
                AUDIT_DUP_GROUP_FILL, n_cols, level=1,
            )
            cur_row += 1
            for b in g:
                mismatch = ("Duplicate ref — same $ as its copies"
                            if same_amt else "Duplicate ref — amount differs")
                values = [
                    b["bill_doc"], b["vendor"], b["bill_date"],
                    b["customer_name"],
                    "", "", "",               # project / division / class — n/a here
                    "",                       # memo placeholder
                    f"Bill total ${b['bill_total']:,.2f}",
                    mismatch,
                    _qbo_link(b["bill_id"]),
                ]
                _audit_write_row(ws, cur_row, values, AUDIT_DUP_CELL_FILL, level=2)
                cur_row += 1

    return len(stale_bills) + len(flagged) + len(uncoded) + n_dup_bills


# ─────────────────────── main ───────────────────────

def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true", help="build rows but do not write xlsx")
    ap.add_argument("--limit", type=int, default=0, help="cap row count (smoke test)")
    args = ap.parse_args()

    started = dt.datetime.now()
    print(f"→ {started:%Y-%m-%d %H:%M:%S}  excel bill sync starting")
    print(f"   output: {OUTPUT_PATH}")

    print("→ authenticating to QBO (Touch ID) …")
    qbo_access, qbo_cid = load_credentials()
    print(f"  ok. company_id={qbo_cid}")

    print("→ fetching vendors …")
    vendors = query_all(qbo_access, qbo_cid, "Vendor")
    vendor_map = {
        v["Id"]: v.get("DisplayName") or v.get("CompanyName") or f"Vendor {v['Id']}"
        for v in vendors
    }
    vendor_root = _build_vendor_root(vendors)   # id → top-most parent id, for the dup audit
    print(f"  {len(vendor_map)} vendors")

    print("→ building account + item maps …")
    account_map, item_map = build_account_maps(qbo_access, qbo_cid)
    print(f"  {len(account_map)} accounts, {len(item_map)} items")

    print("→ building PO map …")
    po_map = build_po_map(qbo_access, qbo_cid)
    print(f"  {len(po_map)} purchase orders")

    print("→ building invoice → payment date map …")
    payment_map = build_payment_map(qbo_access, qbo_cid)
    print(f"  {len(payment_map)} invoices with payment dates")

    # Open bills — any date, Balance > 0
    print("→ fetching OPEN bills (Balance > 0) …")
    open_bills = query_all(qbo_access, qbo_cid, "Bill", where="Balance > '0'")
    n0 = len(open_bills)
    open_bills = [b for b in open_bills if not is_sub_bill(b)]
    print(f"  {n0} → {len(open_bills)} after sub exclusion")

    # Paid bills since cutoff — for ledger sheets
    print(f"→ fetching PAID bills since {PAID_CUTOFF_DATE} (Balance = 0) …")
    paid_bills = query_all(
        qbo_access, qbo_cid, "Bill",
        where=f"Balance = '0' AND TxnDate >= '{PAID_CUTOFF_DATE}'",
    )
    n0 = len(paid_bills)
    paid_bills = [b for b in paid_bills if not is_sub_bill(b)]
    print(f"  {n0} → {len(paid_bills)} after sub exclusion")

    print(f"→ fetching invoices since {INVOICE_CUTOFF_DATE} …")
    invoices_raw = query_all(
        qbo_access, qbo_cid, "Invoice",
        where=f"TxnDate >= '{INVOICE_CUTOFF_DATE}'",
    )
    invoices = [inv for inv in invoices_raw if not is_excluded_invoice(inv)]
    print(f"  {len(invoices_raw)} → {len(invoices)} after retainage exclusion")

    invoices_by_customer: Dict[str, List[dict]] = defaultdict(list)
    for inv in invoices:
        cv = (inv.get("CustomerRef") or {}).get("value", "")
        if cv:
            invoices_by_customer[cv].append(inv)

    print("→ building rows …")
    open_rows = build_rows(
        open_bills, invoices_by_customer, vendor_map, account_map, item_map, po_map,
        payment_map=payment_map,
    )
    paid_rows = build_rows(
        paid_bills, invoices_by_customer, vendor_map, account_map, item_map, po_map,
        payment_map=payment_map,
    )
    all_rows = open_rows + paid_rows
    print(f"  {len(open_rows)} open + {len(paid_rows)} paid = {len(all_rows)} total")

    # status / division breakdowns
    breakdown: Dict[str, int] = defaultdict(int)
    by_division: Dict[str, int] = defaultdict(int)
    for r in all_rows:
        breakdown[r["auto_status"]] += 1
        by_division[r.get("division") or "(none)"] += 1
    print("→ status breakdown (all rows):")
    for s, n in sorted(breakdown.items(), key=lambda x: -x[1]):
        print(f"    {s}: {n}")
    print("→ by division:")
    for d, n in sorted(by_division.items()):
        print(f"    {d}: {n}")

    if args.limit > 0:
        print(f"→ --limit {args.limit}: capping row sets")
        open_rows = open_rows[: args.limit]
        all_rows = all_rows[: args.limit]

    if args.dry_run:
        elapsed = (dt.datetime.now() - started).total_seconds()
        print(f"\n✓ dry run complete in {elapsed:.1f}s — workbook NOT written")
        return 0

    # Read existing workbook's Bills sheet for Lien/Notes preservation
    print("→ reading existing workbook for Lien/Notes preservation …")
    edits = preserve_edits(OUTPUT_PATH)
    if edits:
        nonblank = sum(1 for v in edits.values()
                       if (v.get("Lien") or "") or (v.get("Notes") or ""))
        print(f"  {nonblank} _Keys with Lien/Notes edits carried forward")
    else:
        print("  no existing workbook — first run, nothing to preserve")

    print("→ rotating backup …")
    rotate_backup(OUTPUT_PATH)

    # Two derived row-sets from the line-level data:
    #   • bills_view_rows  — bill-grain (one row per bill). Multi-project
    #                         bills show Project# = "(multiple)".
    #   • all_rows         — line-level. Inventory + Audit sheets use this.
    # (The per-division MFD/RP/CP grain was dropped 2026-07-13 — the Project #
    # already lives on the Bills sheet and the division sheets went unused.)
    print("→ collapsing rows for display …")
    bills_view_rows = collapse_rows(all_rows, grain="bill")
    n_multi = len(multi_project_bill_ids(all_rows))
    print(f"  {len(bills_view_rows)} bills · "
          f"{n_multi} multi-project bills routed to Inventory")

    # Bill payments (AP cash-out) → Pay Ref #/Date/Method columns on the Bills
    # sheet (no separate sheet, no out-of-window bill fetch — we only annotate
    # bills already loaded). One BillPayment fetch since the cutoff, then a
    # bill_id → payment map that _write_bill_row reads via the module global.
    global _BILL_PAY_MAP
    print(f"→ fetching bill payments since {PAID_CUTOFF_DATE} …")
    bill_payments = query_all(qbo_access, qbo_cid, "BillPayment",
                              where=f"TxnDate >= '{PAID_CUTOFF_DATE}'")
    _BILL_PAY_MAP = build_bill_payment_map(bill_payments)
    print(f"  {len(bill_payments)} bill payments → {len(_BILL_PAY_MAP)} bills annotated")

    print("→ building workbook …")
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    wb.remove(wb.active)
    ws_bills = wb.create_sheet("Bills")
    ws_liens = wb.create_sheet("Liens")
    ws_inv   = wb.create_sheet("Inventory")
    ws_audit = wb.create_sheet("QBO Audit")
    # MFD/RP/CP division sheets removed 2026-07-13 (unused; Project # is on Bills).
    # Bill payments show as Pay columns on Bills, not a separate sheet (2026-07-13).

    # Render order matters: build Inventory FIRST so we have the
    # {bill_id → row} anchor map to wire the Bills sheet hyperlinks.
    n_inv, inv_anchor_map = build_inventory_sheet(ws_inv, all_rows)

    n_bills, _ = _render_rows(
        ws_bills, bills_view_rows, BILLS_TABLE, edits,
        sort_key=_vendor_key,
        hide_paid_default=True,
        hide_cols=[LINE_AMT_COL_INDEX],   # Bills hides Line Amount (= Bill Total at bill grain)
        inv_anchor_map=inv_anchor_map,    # hyperlink "(multiple)" → Inventory
        lien_editable=True,               # Bills is the only sheet with the Lien tag + dropdown
    )
    build_liens_sheet(ws_liens)       # live FILTER view of tblBills (Lien set)
    n_audit = build_audit_sheet(ws_audit, all_rows,
                                vendor_root=vendor_root, vendor_map=vendor_map)
    print(f"  Bills: {n_bills} bills (open + paid since {PAID_CUTOFF_DATE})")
    print(f"  Liens: live view  ·  Inventory: {n_inv} lines  ·  Audit: {n_audit} flagged")

    wb.save(OUTPUT_PATH)
    post_process_xlsx(OUTPUT_PATH)

    # Excel-strict preflight — bail loud if any known-bad pattern survived.
    # Better to fail here than ship a workbook that Excel will mangle on open.
    print("→ validating xlsx (Excel-strict preflight) …")
    failures = validate_xlsx(OUTPUT_PATH)
    if failures:
        print("✗ validation FAILED — Excel would warn on open:")
        for f in failures:
            print(f"    {f}")
        print(f"\n  workbook saved anyway at: {OUTPUT_PATH}")
        print("  but you should NOT trust it until these are fixed.")
        return 2
    print(f"  ✓ saved {OUTPUT_PATH}")

    try:
        os.chmod(OUTPUT_PATH, stat.S_IRUSR | stat.S_IWUSR)
    except Exception as e:
        print(f"  ⚠ chmod 600 failed: {e}")

    elapsed = (dt.datetime.now() - started).total_seconds()
    print(f"\n✓ done in {elapsed:.1f}s")
    print(f"   open: {OUTPUT_PATH}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
