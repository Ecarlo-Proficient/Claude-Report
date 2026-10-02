"""
aging_sheet.py — the per-division aging tabs of Open_Invoices.xlsx.

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
    2. **The lien clock** — the `Lien` column (which replaced Days Past Due,
       the user 2026-08-10) gives the date a Ch. 53 notice must be MAILED by.
       Days-past-due was already legible from which bucket the money sits in;
       the lien deadline is not, and it is the one that EXPIRES. Dates come from
       `shared/lien_clock.py`, shared with money_bleeds so both tools can never
       drift apart on a statutory date.
    3. **Why the draw isn't funded yet** — for MFD and CP, the state of the
       PREVIOUS draw. The funding is a chain (the user 2026-08-05): the GC funds
       draw N, we pay draw N's vendor bills, those vendors issue unconditional
       waivers, and the GC needs the waivers before releasing draw N+1. So an
       unpaid draw is rarely about its own bills — it's about whether the one
       before it is cleared. `draw_chain.py` builds the sequence;
       `load_vendor_bill_map` reads what's still owed, from the bill-tracker's
       output file.

       The verdict splits the two very different holds: `PAY BILLS → unlock`
       (previous draw funded, our vendors still owed — **ours** to fix) versus
       `Waiting GC on prev` (previous draw not funded either — upstream of us).
       `This Draw $ Open` keeps the older same-draw figure alongside it.

Rules baked in:
    - Litigation invoices (the `Litigation` checkbox in both Notion trackers)
      were excluded until 2026-10-02; now they are shown and flagged - see below.
    - Aging is by DUE DATE, matching QBO's default AR aging and the
      `Aging Bucket` select that invoice_sync already writes to Notion.
    - Parent-client groups open **expanded** (the user 2026-10-02; collapsed
      from 2026-08-05) - the outline toggles still collapse a client.
    - **Litigation invoices are ON the tabs** since 2026-10-02, flagged in a
      `Litigation` column at the end to filter them out; client and total rows
      are SUBTOTALs, so they follow that filter.
    - **Every row the same height, no wrapping** (the user 2026-10-02): a long
      memo or note is cut off at the cell edge.
    - **Client, then invoices - no project layer** (the user 2026-10-01: "we
      already see the project in the column which we can filter"). The one
      exception is JPI on the MFD tab, which keeps a project sub-group per job
      (`PROJECT_SPLIT_CLIENTS`).
    - **Age is ONE column** (`Aging` = Current / 1-30 / 31-60 / 61-90 / 90+),
      filterable like a property, instead of five money columns (the user
      2026-10-01). Open Balance and Total Amount sit right after Due Date; the
      lien columns sit last.
    - **A tight top** (the user 2026-10-02): row 1 = title in A + the one-line
      summary beside it in B, row 2 = the column headers. No KEY / BY AGE /
      footnote block under the TOTAL row any more.

This module only builds the worksheet; `export_invoices_xlsx.py` owns the
workbook and the Notion pull.
"""
from __future__ import annotations

import datetime as dt
import logging
import re
import sys
from collections import defaultdict
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

import draw_chain

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from shared import paths
from shared import lien_clock


log = logging.getLogger("automation_worker.aging_sheet")

# One tab per division (the user 2026-08-10 — "keep cp and mfd separated").
# (sheet name, division, title, columns to omit) — RP drops the whole
# previous-draw block because RP doesn't bill in draws.
DIVISION_TABS = (
    ("CP Aging", "CP", "CP AGING"),
    ("MFD Aging", "MFD", "MFD AGING"),
    ("RP Aging", "RP", "RP AGING"),
)

# Clients whose invoices keep a PROJECT sub-group under the client row (the user
# 2026-10-01: every tab is client -> invoices, "except JPI in mfd, that one needs
# to be separated"). Keyed by division; matched on the parent client's name.
PROJECT_SPLIT_CLIENTS: Dict[str, "re.Pattern[str]"] = {
    "MFD": re.compile(r"\bJPI\b", re.IGNORECASE),
}

# Type sizes (the user 2026-08-05 — "make the font 12 to see it a bit bigger",
# client rows 13). Client rows carry a point more because the sheet is read
# by client: the client name is the line that has to land first.
BODY_PT = 13          # bumped from 12 with the taller rows (the user 2026-10-02)
CLIENT_PT = 14
TITLE_PT = 15
# Fixed row heights (points): every row the same height, text never wraps - a long
# memo or note is cut off at the cell edge instead of stretching its row (the user
# 2026-10-02). Slightly taller than Excel's default for readability.
ROW_PT = 21
CLIENT_ROW_PT = 23
# Widths are in characters of the default 11pt font; _autofit scales for 12pt.
# The cap stops a long memo or note from creating a column you have to scroll.
MAX_COL_WIDTH = 62
# The Client / Invoice column stops here: memos are cut off at the edge anyway,
# and a wider column pushes the money off the screen (2026-10-02).
MAX_LABEL_WIDTH = 46
# Default zoom on every aging tab: the 13pt font and taller rows read too large
# at 100% (the user 2026-10-02: "zoom it out a little by default").
ZOOM_PCT = 85

# The bill-tracker's Excel output, via the ONE shared resolver so the path can't
# drift between writer and reader. We read the FILE, not the bill-tracker's code
# (tools never import tools, repo rule 3). NOTE: the Dockerized invoice-sync must
# set ACB_BILL_TRACKER_XLSX to wherever the AP file is mounted in-container.
BILL_TRACKER_PATH = paths.bill_tracker_xlsx()

# Previous-draw verdicts. The funding chain (the user 2026-08-05): the GC funds
# draw N, we pay draw N's vendor bills, the vendors issue unconditional waivers,
# and the GC needs those waivers before funding draw N+1. So for an unpaid draw
# the question is never "are THIS draw's bills paid" — it's whether the PREVIOUS
# draw is cleared. These are a small fixed set so the column filters to two
# clicks; counts and dollars live in their own columns.
PREV_BLOCKED = "PAY BILLS → unlock"   # prev draw funded, our vendors still owed — WE are the blocker
PREV_WAITING_GC = "Waiting GC on prev"  # prev draw itself unpaid — blocker is upstream of us
PREV_CLEAR = "Clear"                  # prev draw funded and its vendors paid
PREV_FIRST = "First draw"             # provably first (Draw #1), nothing gates it
PREV_NOT_SYNCED = "Prev not synced"   # a later draw whose predecessor never entered Notion
PREV_NOT_DRAW = "Not a draw"          # retainage / turnkey one-offs — no chain
PREV_MULTI = "Multi-contract"         # parallel contracts, bills unattributable (see draw_chain)
VENDOR_NA = "n/a"                     # RP — no draws at all
VENDOR_UNKNOWN = "?"                  # bill tracker file missing / unreadable

# Verdicts with no previous draw to report on. The count/$ cells under them are
# not "zero", they are unanswerable — so the whole block greys out. The verdict
# text itself survives (see _grey_out_vendor_block: it only fills EMPTY cells),
# which keeps "Multi-contract" and "First draw" readable while making clear
# there is no number beside them.
NO_CHAIN_VERDICTS = (
    VENDOR_NA, PREV_NOT_DRAW, PREV_MULTI, PREV_FIRST, PREV_NOT_SYNCED, VENDOR_UNKNOWN,
)

# The divisions that bill in draws at all. Bill lines are matched to the invoice
# authorising their payment via the DRAW PERIOD here (bill-tracker README, "How
# matching works"). RP matches on "earliest invoice on/after bill date" — not a
# draw period, and RP has no draw chain to walk — so the whole block is n/a.
DRAW_DIVISIONS = ("MFD", "CP")

# (header, width, number_format)
# No Division column: each tab IS one division (the user 2026-08-10), so it
# would repeat the tab name on every row.
# Order (the user 2026-10-01 / 10-02): who / which invoice / when, the money
# (Open Balance, Total Amount), THEN Due Date and its age (Aging), the Project #,
# the draw block and notes, the lien columns, and Division / Litigation LAST so
# they filter. Division shows only on the All Open tab (the division tabs ARE
# their division). The "Prev Draw" invoice # column was removed 2026-10-02.
COLUMNS: List[Tuple[str, int, Optional[str]]] = [
    ("Client / Invoice", 34, None),
    ("Invoice #",        11, None),     # notes_preserve finds this column by its header
    ("Date",             11, "mm/dd/yyyy"),
    # Open Balance first, then the invoice's original Total Amount, so the pair
    # reads open->total left-to-right (the user 2026-08-11). Both bold and framed
    # (the user 2026-10-02: "more pronounced"). Open Balance is amber-flagged when
    # it differs from Total Amount, i.e. the invoice is partly paid (2026-08-14).
    ("Open Balance",     16, '"$"#,##0.00'),
    ("Total Amount",     16, '"$"#,##0.00'),
    ("Due Date",         11, "mm/dd/yyyy"),
    # The five bucket columns collapsed into one filterable value (the user
    # 2026-10-01). The cell keeps its green->red tint, so colour still reads age.
    ("Aging",            10, None),
    ("Project #",        13, None),
    ("Prev Draw Status", 19, None),
    ("Prev Bills Open",   9, "0"),
    ("Prev $ Open",      15, '"$"#,##0.00'),
    ("This Draw $ Open", 15, '"$"#,##0.00'),
    # "Notion Notes": the collections clerk's Quick Status from Notion - named so it
    # is never confused with an Excel cell Note (the user 2026-10-02).
    ("Notion Notes",     16, None),
    ("Last Action",      12, "mm/dd/yyyy"),
    # Replaced Days Past Due (the user 2026-08-10): the date a notice must be
    # MAILED by - the one deadline that expires.
    ("Lien",             22, None),
    # The Notion Lien Tracker status (Mailed / Lien filed / ...) - the SAME value the
    # dashboard's Open Invoices Lien column shows, kept beside the deadline clock so the
    # workbook and the site never disconnect (the user 2026-08-18). Blank = no lien linked.
    ("Lien status",      18, None),
    ("Division",         10, None),
    # Litigation invoices are ON the tabs since 2026-10-02 (the user: "i want
    # litigation not excluded ... add column at end to filter out"). Yes / No.
    ("Litigation",       11, None),
]

# 0-based positions used when writing rows (kept in sync with COLUMNS above).
# C_TOTAL = Open Balance, C_INVTOTAL = the invoice's original Total Amount.
C_LABEL, C_INV, C_DATE, C_TOTAL, C_INVTOTAL, C_DUE, C_AGING, C_PROJ = range(8)
C_VSTATUS, C_VBILLS, C_VAMT, C_THIS, C_NOTES, C_ACTION = range(8, 14)
C_LIEN, C_LIENSTATUS, C_DIV, C_LITIG = range(14, 18)

# Width caps for free-text columns: past this the text is cut off at the cell
# edge rather than spreading the sheet out (the user 2026-10-02: "always check
# width to not spread info out").
COL_CAPS = {C_LABEL: MAX_LABEL_WIDTH, C_PROJ: 24, C_NOTES: 40, C_LIEN: 28, C_LIENSTATUS: 18}

# Row 1 = title + summary, row 2 = column headers, data from row 3.
HEADER_ROW = 2

# Short codes and dates read best centred; everything is vertically centred
# (the user's screenshot 2026-10-02: "2 inv" and totals sat low in taller rows).
CENTERED_COLS = (C_INV, C_DATE, C_DUE, C_AGING, C_ACTION, C_DIV, C_LITIG)

# Roll-up columns: on client / total rows these are SUBTOTAL(9, ...) formulas,
# so they follow a filter (untick Litigation "Yes" and every total drops it) yet
# still count rows hidden by collapsing a group (function 9, not 109).
SUM_COLS = (C_TOTAL, C_INVTOTAL, C_VBILLS, C_VAMT, C_THIS)

# The values the Aging column takes, youngest first (index = bucket_index()).
AGING_LABELS = ("Current", "1-30", "31-60", "61-90", "90+")

# The previous-draw block - greyed out wherever there is no chain to read.
VENDOR_COLS = (C_VSTATUS, C_VBILLS, C_VAMT)

# Columns the RP tab omits: the whole previous-draw block plus the same-draw
# figure. RP doesn't bill in draws, so on the combined tab these are 34 rows of
# grey "n/a". (the user 2026-08-05)
RP_DROP_COLUMNS = (C_VSTATUS, C_VBILLS, C_VAMT, C_THIS, C_DIV)
# The CP / MFD tabs drop only the Division column (the tab IS the division).
# The draw block was taken off every tab 2026-10-02 (the user: "remove prev draw
# from cp and mfd columns ... i meant the whole block"), so the CP / MFD tabs now
# drop the same columns RP always did. The block's code stays for a revival.
DIVISION_TAB_DROP = RP_DROP_COLUMNS
# Display order of divisions on the All Open tab.
DIVISION_ORDER = {"CP": 0, "MFD": 1, "RP": 2, "Lease": 3}

_THIN = Side(style="thin", color="000000")
_MEDIUM = Side(style="medium", color="000000")


# ─────────────────────── palette ───────────────────────
# Colour is here to answer "how bad is this" without reading a number, so it
# only ever encodes AGE (the five buckets) and STATE (vendors, dead cells).
# Nothing decorative gets a fill.

_HEADER_FILL = PatternFill("solid", fgColor="1F4E78")   # matches the Open Invoices tab
_HEADER_FONT = Font(bold=True, color="FFFFFF")

# Aging cell tint, green -> red as the money ages (Current, 1-30, ... 90+), so
# the one Aging column still reads severity at a glance.
BUCKET_CELL_FILLS = (
    PatternFill("solid", fgColor="E8F5E9"),
    PatternFill("solid", fgColor="F1F8E9"),
    PatternFill("solid", fgColor="FDF2E0"),
    PatternFill("solid", fgColor="FAE7E0"),
    PatternFill("solid", fgColor="F8DDDA"),
)

# Hierarchy bands (the user 2026-08-14 — "professional colors"): structure gets a
# quiet neutral so the saturated palette stays reserved for DATA. High-contrast at
# every step: darkest→slate→grey→white, white text on the two dark bands.
_GRAND_FILL = PatternFill("solid", fgColor="333F50")    # all-clients roll-up (darkest)
_PARENT_FILL = PatternFill("solid", fgColor="44546A")   # client header — slate
_PROJECT_FILL = PatternFill("solid", fgColor="F2F2F2")  # project sub-group — neutral grey
_ON_DARK = "FFFFFF"       # text on the slate / dark-slate bands
_ON_LIGHT = "3F3F3F"      # text on the grey project band
_ROW_TEXT = "333333"      # detail (invoice) rows
_MEMO_TEXT = "595959"     # the memo label on a detail row (secondary)
_HAIR = Side(style="thin", color="D9D9D9")              # quiet row separator
# Open Balance ≠ Total Amount = the invoice is partly paid. Amber flag (Excel's
# "Neutral" pair) replaces the per-row data bar (the user 2026-08-14).
_PARTIAL_FILL = PatternFill("solid", fgColor="FFEB9C")
_PARTIAL_TEXT = "9C6500"

# Dead cells: the previous-draw block where there's nothing to read. Grey fill,
# darker grey text — "nothing to see here" without looking like missing data.
_NA_FILL = PatternFill("solid", fgColor="D9D9D9")
_NA_COLOR = "808080"

_BLOCKED_FILL = PatternFill("solid", fgColor="F8DDDA")  # the one row-level call to action
_LINK_COLOR = "0563C1"                                  # Excel's own hyperlink blue

# Lien deadlines are the only HARD expiry on this sheet — miss one and the right
# is gone, not merely late. Past due gets the only reversed-out cell here.
_LIEN_PAST_FILL = PatternFill("solid", fgColor="922B21")
_LIEN_URGENT_FILL = PatternFill("solid", fgColor="F8DDDA")
_LIEN_WATCH_FILL = PatternFill("solid", fgColor="FDF2E0")


# ─────────────────────── aging ───────────────────────

def bucket_index(days_past_due: Optional[int]) -> int:
    """Position in AGING_LABELS for a signed days-past-due value.

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


def _open_bills_for(
    invoice_num: str, vendor_map: Optional[Dict[str, Tuple[float, int, int]]]
) -> Tuple[Optional[int], Optional[float]]:
    """(open bill count, open $) still owed against one invoice's draw period."""
    if vendor_map is None:
        return None, None
    hit = vendor_map.get(str(invoice_num).strip())
    if not hit:
        return None, None
    amount, bills, _vendors = hit
    return bills, amount


def vendor_cells(
    division: str,
    invoice_num: str,
    vendor_map: Optional[Dict[str, Tuple[float, int, int]]],
    chains: Optional[Any] = None,
) -> Tuple[str, str, Optional[int], Optional[float], Optional[float]]:
    """The previous-draw block for one invoice.

    Returns (prev draw label, verdict, prev open bill count, prev open $,
    this draw's own open $).

    The verdict separates the two very different reasons a draw sits unpaid:

      * `PAY BILLS → unlock` — the previous draw WAS funded, but our vendors on
        it are still owed, so no unconditional waivers exist and the GC won't
        release this draw. **We** are the blocker, and the fix is ours.
      * `Waiting GC on prev` — the previous draw hasn't been funded either, so
        we couldn't have paid those vendors yet. The blocker is upstream.

    Collapsing those into one "unpaid bills" flag (as the first cut of this tab
    did) hides which of the two you're looking at, and they lead to opposite
    actions.
    """
    this_bills, this_amount = _open_bills_for(invoice_num, vendor_map)
    if division not in DRAW_DIVISIONS:
        return "", VENDOR_NA, None, None, None
    if chains is None:
        return "", VENDOR_UNKNOWN, None, None, this_amount

    outcome, prev = chains.previous_draw(invoice_num)
    if outcome == draw_chain.CHAIN_NOT_A_DRAW:
        return "", PREV_NOT_DRAW, None, None, this_amount
    if outcome == draw_chain.CHAIN_MULTI_CONTRACT:
        return "", PREV_MULTI, None, None, this_amount
    if outcome == draw_chain.CHAIN_PREV_UNKNOWN:
        return "", PREV_NOT_SYNCED, None, None, this_amount
    if outcome == draw_chain.CHAIN_FIRST_DRAW or prev is None:
        return "", PREV_FIRST, None, None, this_amount

    prev_num = prev["invoice_num"]
    if vendor_map is None:
        return prev_num, VENDOR_UNKNOWN, None, None, this_amount

    prev_bills, prev_amount = _open_bills_for(prev_num, vendor_map)
    if not prev["is_paid"]:
        # Show the bills anyway — they're what we'll owe once it does fund.
        return prev_num, PREV_WAITING_GC, prev_bills, prev_amount, this_amount
    if prev_bills:
        return prev_num, PREV_BLOCKED, prev_bills, prev_amount, this_amount
    return prev_num, PREV_CLEAR, None, None, this_amount


# ─────────────────────── sheet build ───────────────────────

class _Grid:
    """Maps logical column indices (the C_* constants) to physical ones.

    A tab can drop columns — the RP view hides the whole previous-draw block,
    which is meaningless there — but every row is still built at full width
    against the C_* constants. This does the projection in one place so no
    caller has to think about which physical column a field landed in.
    """

    def __init__(self, drop: Tuple[int, ...] = ()) -> None:
        self.visible = [i for i in range(len(COLUMNS)) if i not in drop]
        self._pos = {logical: n for n, logical in enumerate(self.visible)}

    def __contains__(self, logical: int) -> bool:
        return logical in self._pos

    def col(self, logical: int) -> int:
        """1-based physical column for a logical index."""
        return self._pos[logical] + 1

    def cell(self, ws: Worksheet, row: int, logical: int):
        """The cell for a logical column, or None when that column is hidden."""
        if logical not in self._pos:
            return None
        return ws.cell(row=row, column=self.col(logical))

    @property
    def width(self) -> int:
        return len(self.visible)


def _autofit(ws: Worksheet, grid: _Grid, last_row: int,
             headers: Optional[Dict[int, str]] = None, uncapped: Tuple[int, ...] = ()) -> None:
    """Size every column to its widest cell, in the sheet's own font.

    openpyxl has no real autofit — Excel computes widths at render time and
    openpyxl never renders. Measuring the strings we wrote is the honest
    approximation: character count scaled for the 12pt body (wider than the
    11pt default the width unit assumes), plus padding for the filter arrow.
    """
    for logical in grid.visible:
        name, floor_width, number_format = COLUMNS[logical]
        name = (headers or {}).get(logical, name)
        letter = get_column_letter(grid.col(logical))
        # The header is bold, LEFT-aligned and carries a filter arrow: ~1.3 width
        # units per character plus ~3 for the arrow. (Centred, the name ran into
        # the arrow - the user's screenshot 2026-10-02: "Due Date", "Aging".)
        header_need = len(name) * 1.3 + 3
        widest = 0
        for row in range(HEADER_ROW + 1, last_row + 1):
            value = ws.cell(row=row, column=grid.col(logical)).value
            if value is None:
                continue
            if isinstance(value, str) and value.startswith("="):
                continue                           # a SUBTOTAL formula, not its shown width
            if isinstance(value, dt.date):
                rendered = 10                      # mm/dd/yyyy
            elif isinstance(value, float) and number_format and "$" in number_format:
                rendered = len(f"{value:,.2f}") + 2   # "$" plus separators
            else:
                rendered = len(str(value))
            widest = max(widest, rendered)
        width = max(floor_width, header_need, widest * BODY_PT / 11.0 + 1.5)
        cap = MAX_COL_WIDTH if logical in uncapped else COL_CAPS.get(logical, MAX_COL_WIDTH)
        ws.column_dimensions[letter].width = min(width, cap)


def _write_row(
    ws: Worksheet,
    grid: _Grid,
    row_num: int,
    values: List[Any],
    *,
    bold: bool,
    size: float = BODY_PT,
    fill: Optional[PatternFill] = None,
    color: str = "000000",
) -> None:
    for logical in grid.visible:
        cell = ws.cell(row=row_num, column=grid.col(logical), value=values[logical])
        number_format = COLUMNS[logical][2]
        if number_format:
            cell.number_format = number_format
        cell.font = Font(bold=bold, size=size, color=color)
        cell.alignment = Alignment(
            horizontal="center" if logical in CENTERED_COLS else None,
            vertical="center", wrap_text=False,
        )
        if fill:
            cell.fill = fill


def _flag_partial(cell, open_v, total_v, *, bold: bool, size: float) -> None:
    """Amber-flag an Open Balance cell when it doesn't equal the invoice total,
    i.e. the invoice is partly paid (the user 2026-08-14, replacing the data bar).
    Full-open invoices (open == total) get nothing, so only the exceptions stand out."""
    if cell is None:
        return
    if round(open_v or 0.0, 2) != round(total_v or 0.0, 2):
        cell.fill = _PARTIAL_FILL
        cell.font = Font(bold=bold, size=size, color=_PARTIAL_TEXT)


def _grey_out_vendor_block(ws: Worksheet, grid: _Grid, row_num: int) -> None:
    """Mark the previous-draw cells as not-applicable on this row.

    Used wherever there is no draw chain to read — RP (no draws at all), non-draw
    invoices, first draws, and the multi-contract projects whose bills can't be
    attributed. A blank would read as "not looked up yet"; grey fill with darker
    grey text says the cell is intentionally dead. (the user)

    Only EMPTY cells get the "n/a" text, so a verdict already written there
    ("Multi-contract", "First draw") stays readable inside the grey block.
    """
    for logical in VENDOR_COLS:
        cell = grid.cell(ws, row_num, logical)
        if cell is None:
            continue
        if cell.value in (None, ""):
            cell.value = VENDOR_NA
        cell.number_format = "General"  # else "n/a" fights the $ / 0 formats
        cell.fill = _NA_FILL
        cell.font = Font(color=_NA_COLOR, italic=True, size=BODY_PT)
        cell.alignment = Alignment(horizontal="center", vertical="center")


def _subtotal(ws: Worksheet, grid: "_Grid", row: int, first: int, last: int) -> None:
    """Turn a roll-up row's money cells into SUBTOTAL(9, ...) over the rows under it.
    SUBTOTAL skips nested SUBTOTALs (project rows inside a client) and filtered-out
    rows, but keeps rows hidden by collapsing a group."""
    if last < first:
        return
    for logical in SUM_COLS:
        cell = grid.cell(ws, row, logical)
        if cell is None:
            continue
        letter = get_column_letter(grid.col(logical))
        cell.value = f"=SUBTOTAL(9,{letter}{first}:{letter}{last})"


def _frame_money(ws: Worksheet, grid: "_Grid", first: int, last: int) -> None:
    """A thin rule either side of the Open Balance / Total Amount pair, so the two
    money columns read as one block (the user 2026-10-02: "more pronounced")."""
    left, right = grid.col(C_TOTAL), grid.col(C_INVTOTAL)
    for r in range(first, last + 1):
        for col, side in ((left, "left"), (right, "right")):
            cell = ws.cell(row=r, column=col)
            b = cell.border
            kw = {"left": b.left, "right": b.right, "top": b.top, "bottom": b.bottom}
            kw[side] = _THIN
            cell.border = Border(**kw)


def build_aging_sheet(
    ws: Worksheet,
    invoices: List[dict],
    *,
    today: dt.date,
    litigation_excluded: Optional[int],
    vendor_as_of: Optional[dt.datetime] = None,
    drop_columns: Tuple[int, ...] = (),
    title: str = "AR AGING",
    scope_note: str = "",
    split_clients: Optional["re.Pattern[str]"] = None,
    header_overrides: Optional[Dict[int, str]] = None,
    group_by_division: bool = False,
    uncapped: Tuple[int, ...] = (),
) -> None:
    """Write an aging tab.

    `invoices` are plain dicts (built by export_invoices_xlsx._aging_record /
    _lease_record) so this module stays free of Notion and QBO plumbing.

    Layout: a header, an always-visible ALL CLIENTS total, then one bold summary
    row per parent client with its invoices underneath, expanded (the user
    2026-10-02), with outline toggles to collapse a client. Client -> invoice,
    no project layer (the user 2026-10-01); a client matching `split_clients`
    keeps a project sub-group (JPI on MFD). Summary-above-detail requires
    outlinePr.summaryBelow = False; without it Excel puts the collapse toggle on
    the row after the block and the grouping reads backwards.

    `drop_columns` hides logical columns entirely - the RP tab passes the whole
    previous-draw block, which has no meaning for a division that doesn't bill
    in draws. `header_overrides` renames a column on this tab only (the lease tab
    shows its QBO item where the project # would be). `litigation_excluded=None` omits that count.
    """
    grid = _Grid(drop_columns)
    shows_vendor_block = C_VSTATUS in grid
    shows_lien = C_LIEN in grid
    headers = {i: c[0] for i, c in enumerate(COLUMNS)}
    headers.update(header_overrides or {})

    # Group by parent client. Unresolved relations fall back to a stable label
    # rather than being dropped - an invoice with no parent is still money owed.
    # On All Open (`group_by_division`) a client is grouped within its division,
    # so every client row carries ONE Division value and a Division filter keeps
    # whole groups (header, invoices, total) together.
    by_parent: Dict[Tuple[str, str], List[dict]] = defaultdict(list)
    for rec in invoices:
        div = rec["division"] if group_by_division else ""
        by_parent[(div, rec["parent"] or "(no client on file)")].append(rec)

    # ── title block ──
    ws.cell(
        row=1, column=1,
        value=f"{title} - as of {today.strftime('%m/%d/%Y')}",
    ).font = Font(bold=True, size=TITLE_PT)

    parts = [f"{len(invoices)} open invoices"]
    if litigation_excluded is not None:
        parts.append(f"litigation excluded ({litigation_excluded})")
    if scope_note:
        parts.insert(0, scope_note)
    stale = vendor_as_of is None or vendor_as_of.date() < today
    if shows_vendor_block:
        # The vendor columns have a different clock from everything else on this
        # tab: they come from the AP tool's last run, not from this one. Say
        # which, in plain words, so nobody reads a day-old figure as current.
        if vendor_as_of is None:
            parts.append("Vendor bill status UNAVAILABLE - run `sync-ap`, then re-run this")
        elif vendor_as_of.date() < today:
            parts.append(
                f"VENDOR COLUMNS ARE {_humanize_age(vendor_as_of).upper()} OLD "
                f"(Bill Tracker last run {vendor_as_of.strftime('%m/%d, %I:%M %p')}) - "
                f"run `sync-all`, which runs AP before AR"
            )
        else:
            parts.append(
                f"Vendor bill status current as of {vendor_as_of.strftime('%m/%d, %I:%M %p')}"
            )
    # Title in A1, its one-line summary right beside it in B1 (the user's mockup
    # 2026-10-02: "new top header for all pages like this. keep it tight").
    subtitle_cell = ws.cell(row=1, column=2, value=" · ".join(parts))
    subtitle_cell.alignment = Alignment(vertical="center")
    ws.cell(row=1, column=1).alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 26
    subtitle_cell.font = (
        Font(bold=True, color="C00000", size=BODY_PT)
        if (shows_vendor_block and stale)
        else Font(size=BODY_PT)
    )

    header_row = HEADER_ROW
    for logical in grid.visible:
        cell = ws.cell(row=header_row, column=grid.col(logical), value=headers[logical])
        cell.font = Font(bold=True, color="FFFFFF", size=BODY_PT)
        cell.fill = _HEADER_FILL
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=False)
        cell.border = Border(bottom=_MEDIUM)
    ws.row_dimensions[header_row].height = 30
    ws.freeze_panes = ws.cell(row=header_row + 1, column=1)

    def _sums(records: List[dict]):
        """(per-age open $, open $, total $, prev bills, prev $, this-draw $)."""
        ages = [0.0] * len(AGING_LABELS)
        vendor_sum = 0.0
        vendor_bills = 0
        this_sum = 0.0
        invtotal_sum = 0.0
        for rec in records:
            ages[bucket_index(rec["days_past_due"])] += rec["open_balance"] or 0.0
            vendor_sum += rec.get("vendor_amount") or 0.0
            vendor_bills += rec.get("vendor_bills") or 0
            this_sum += rec.get("this_draw_amount") or 0.0
            invtotal_sum += rec["total_amount"] or 0.0
        return ages, sum(ages), invtotal_sum, vendor_bills, vendor_sum, this_sum

    def _rollup(label, inv_label, records) -> List[Any]:
        _ages, open_sum, invtotal_sum, vendor_bills, vendor_sum, this_sum = _sums(records)
        row: List[Any] = [""] * len(COLUMNS)
        row[C_LABEL] = label
        row[C_INV] = inv_label
        row[C_TOTAL] = open_sum
        row[C_INVTOTAL] = invtotal_sum or None
        row[C_VBILLS] = vendor_bills or None
        row[C_VAMT] = vendor_sum or None
        row[C_THIS] = this_sum or None
        return row

    # ── grand total (always visible, never grouped) ──
    row_num = header_row + 1
    _write_row(ws, grid, row_num, _rollup(f"ALL CLIENTS ({len({k[1] for k in by_parent})})", "", invoices),
               bold=True, size=CLIENT_PT, fill=_GRAND_FILL, color=_ON_DARK)
    for logical in grid.visible:
        ws.cell(row=row_num, column=grid.col(logical)).border = Border(bottom=_HAIR)
    ws.row_dimensions[row_num].height = CLIENT_ROW_PT
    grand_row = row_num
    row_num += 1

    def write_header(label, records, *, fill, size, level, color, division) -> None:
        """A group's NAME row on top (client at level 0, project at level 1): the
        name and its invoice count. Its totals sit on the group's total row at the
        bottom (the user 2026-10-02), QuickBooks-report style."""
        nonlocal row_num
        head: List[Any] = [""] * len(COLUMNS)
        head[C_LABEL] = label
        head[C_INV] = f"{len(records)} inv"
        head[C_DIV] = division
        _write_row(ws, grid, row_num, head, bold=True, size=size, fill=fill, color=color)
        grid.cell(ws, row_num, C_LABEL).alignment = Alignment(indent=level, vertical="center")
        if level > 0:
            ws.row_dimensions[row_num].outlineLevel = level
        ws.row_dimensions[row_num].height = CLIENT_ROW_PT if level == 0 else ROW_PT
        row_num += 1

    def write_total(label, records, first: int, *, level, division) -> None:
        """A group's TOTAL row under its invoices: SUBTOTALs over rows first..here-1,
        so it follows filters and skips nested project totals."""
        nonlocal row_num
        total = _rollup(f"Total {label}", "", records)
        total[C_DIV] = division
        _write_row(ws, grid, row_num, total, bold=True, size=BODY_PT if level else CLIENT_PT - 1,
                   fill=_PROJECT_FILL, color=_ON_LIGHT)
        grid.cell(ws, row_num, C_LABEL).alignment = Alignment(indent=level, vertical="center")
        for logical in grid.visible:
            ws.cell(row=row_num, column=grid.col(logical)).border = Border(top=_HAIR, bottom=_THIN)
        # A group with no MFD/CP work has no vendor answer either - grey the
        # block, or an all-RP group reads as "clear".
        if shows_vendor_block and not any(r["division"] in DRAW_DIVISIONS for r in records):
            _grey_out_vendor_block(ws, grid, row_num)
        _subtotal(ws, grid, row_num, first, row_num - 1)
        if level > 0:
            ws.row_dimensions[row_num].outlineLevel = level
        ws.row_dimensions[row_num].height = ROW_PT if level else CLIENT_ROW_PT
        row_num += 1

    def write_detail(rec, level) -> None:
        nonlocal row_num
        slot = bucket_index(rec["days_past_due"])
        detail: List[Any] = [""] * len(COLUMNS)
        # QBO memos carry hard line breaks ("... Draw 2026\n(Period: ...)").
        # Left raw they render as one run-on line or a stray box, since this
        # column doesn't wrap - collapse to single-spaced text.
        label = " ".join((rec["memo"] or rec["project_num"] or "").split())
        detail[C_LABEL] = label[:120]
        detail[C_PROJ] = rec["project_num"]
        detail[C_INV] = rec["invoice_num"]
        detail[C_DATE] = rec["invoice_date"]
        detail[C_DUE] = rec["due_date"]
        detail[C_TOTAL] = rec["open_balance"]
        detail[C_INVTOTAL] = rec["total_amount"]
        detail[C_AGING] = AGING_LABELS[slot]
        detail[C_VSTATUS] = rec.get("vendor_status")
        detail[C_VBILLS] = rec.get("vendor_bills")
        detail[C_VAMT] = rec.get("vendor_amount")
        detail[C_THIS] = rec.get("this_draw_amount")
        detail[C_NOTES] = rec.get("notes")
        detail[C_ACTION] = rec.get("last_action")
        if shows_lien and rec.get("lien") is not None:       # lease rows carry no lien clock
            detail[C_LIEN] = rec["lien"].label
            detail[C_LIENSTATUS] = rec.get("lien_status") or ""
        detail[C_DIV] = rec["division"]
        detail[C_LITIG] = "Yes" if rec.get("litigation") else "No"
        _write_row(ws, grid, row_num, detail, bold=False, color=_ROW_TEXT)
        # Indent matches the outline depth: 2 under a project, 1 under a client.
        # Memo reads as secondary (lighter grey).
        label_cell = grid.cell(ws, row_num, C_LABEL)
        label_cell.alignment = Alignment(indent=level, vertical="center")
        label_cell.font = Font(size=BODY_PT, color=_MEMO_TEXT)

        # The invoice number opens the invoice in QBO. Putting the link on the
        # number rather than in its own column keeps the sheet narrow and puts
        # the click where the eye already is.
        if rec.get("qbo_link"):
            inv_cell = grid.cell(ws, row_num, C_INV)
            if inv_cell is not None:
                inv_cell.hyperlink = rec["qbo_link"]
                inv_cell.font = Font(color=_LINK_COLOR, underline="single", size=BODY_PT)

        # The Aging cell keeps the green->red tint the five bucket columns used
        # to carry, so colour still answers "how old" without reading the word.
        aging_cell = grid.cell(ws, row_num, C_AGING)
        if aging_cell is not None:
            aging_cell.fill = BUCKET_CELL_FILLS[slot]
            aging_cell.alignment = Alignment(horizontal="center", vertical="center")

        # The lien cell carries the only hard expiry on this sheet, so it gets
        # the strongest cue: a missed or imminent notice deadline is a right
        # that disappears, not just money that is late.
        lien_cell = grid.cell(ws, row_num, C_LIEN)
        if lien_cell is not None and rec.get("lien") is not None:
            state = rec["lien"].state
            if state == lien_clock.STATE_PAST:
                lien_cell.font = Font(bold=True, color="FFFFFF", size=BODY_PT)
                lien_cell.fill = _LIEN_PAST_FILL
            elif state == lien_clock.STATE_URGENT:
                lien_cell.font = Font(bold=True, color="922B21", size=BODY_PT)
                lien_cell.fill = _LIEN_URGENT_FILL
            elif state == lien_clock.STATE_WATCH:
                lien_cell.font = Font(color="8A5A00", size=BODY_PT)
                lien_cell.fill = _LIEN_WATCH_FILL
            elif state == lien_clock.STATE_SENT:
                lien_cell.font = Font(color="2E7D32", size=BODY_PT)

        if shows_vendor_block:
            verdict = rec.get("vendor_status")
            if verdict in NO_CHAIN_VERDICTS:
                _grey_out_vendor_block(ws, grid, row_num)
            else:
                status_cell = grid.cell(ws, row_num, C_VSTATUS)
                if verdict == PREV_BLOCKED:
                    # The one verdict that is ours to act on - pay these, get
                    # the waivers, unlock this draw.
                    status_cell.font = Font(bold=True, color="C00000", size=BODY_PT)
                    status_cell.fill = _BLOCKED_FILL
                elif verdict == PREV_WAITING_GC:
                    status_cell.font = Font(color="B06000", size=BODY_PT)
                elif verdict == PREV_CLEAR:
                    status_cell.font = Font(color="2E7D32", size=BODY_PT)

        # Open Balance + Total Amount carry the row: bold (the user 2026-10-02).
        for logical in SUM_COLS:
            money = grid.cell(ws, row_num, logical)
            if money is not None:
                money.font = Font(bold=True, size=BODY_PT, color="000000")
        # Amber-flag the Open Balance when the invoice is only partly paid.
        _flag_partial(grid.cell(ws, row_num, C_TOTAL),
                      rec["open_balance"], rec["total_amount"], bold=True, size=BODY_PT)
        if rec.get("litigation"):
            lit = grid.cell(ws, row_num, C_LITIG)
            if lit is not None:
                lit.font = Font(bold=True, color="922B21", size=BODY_PT)
                lit.fill = _LIEN_URGENT_FILL

        for logical in grid.visible:                   # a quiet rule between invoices
            ws.cell(row=row_num, column=grid.col(logical)).border = Border(bottom=_HAIR)
        ws.row_dimensions[row_num].outlineLevel = level
        ws.row_dimensions[row_num].height = ROW_PT   # expanded by default (the user 2026-10-02)
        row_num += 1

    def _oldest_first(r: dict):
        return (r["invoice_date"] or dt.date.max, r["due_date"] or dt.date.max, str(r["invoice_num"]))

    # Alphabetical by client (the user 2026-08-14) - a predictable order to scan,
    # not biggest-balance-first (All Open: by division first). Inside a client:
    # invoice date, OLDEST first (the user 2026-10-02) - the Project # column
    # filters a single job.
    for div, parent in sorted(by_parent, key=lambda k: (DIVISION_ORDER.get(k[0], 9), (k[1] or "").upper())):
        records = sorted(by_parent[(div, parent)], key=_oldest_first)
        write_header(parent, records, fill=_PARENT_FILL, size=CLIENT_PT, level=0, color=_ON_DARK,
                     division=div)
        first = row_num

        by_project: Dict[str, List[dict]] = defaultdict(list)
        for rec in records:
            by_project[rec["project_num"] or "(no project #)"].append(rec)

        # Client -> invoice (the user 2026-10-01). Only a client named in
        # PROJECT_SPLIT_CLIENTS (JPI on MFD) keeps the project layer between them:
        # client 0 -> project 1 -> invoice 2, each project with its own total.
        if split_clients is not None and split_clients.search(parent) and len(by_project) > 1:
            for proj in sorted(by_project, key=lambda p: (p or "").upper()):
                write_header(proj, by_project[proj], fill=_PROJECT_FILL, size=BODY_PT, level=1,
                             color=_ON_LIGHT, division=div)
                proj_first = row_num
                for rec in by_project[proj]:
                    write_detail(rec, level=2)
                write_total(proj, by_project[proj], proj_first, level=1, division=div)
        else:
            for rec in records:
                write_detail(rec, level=1)
        write_total(parent, records, first, level=0, division=div)

    last_data_row = row_num - 1
    _subtotal(ws, grid, grand_row, grand_row + 1, last_data_row)

    # Bottom total (the user 2026-08-14 - "need a sum"): a footer that mirrors the
    # ALL CLIENTS roll-up, so the number is there at the end of a long scroll too.
    # It sits BELOW last_data_row, so it's outside the autofilter range and a
    # filter can never hide it.
    _write_row(ws, grid, row_num, _rollup("TOTAL", "", invoices), bold=True, size=CLIENT_PT,
               fill=_GRAND_FILL, color=_ON_DARK)
    for logical in grid.visible:
        ws.cell(row=row_num, column=grid.col(logical)).border = Border(top=_MEDIUM)
    ws.row_dimensions[row_num].height = CLIENT_ROW_PT
    _subtotal(ws, grid, row_num, grand_row + 1, last_data_row)
    row_num += 1

    ws.sheet_properties.outlinePr.summaryBelow = False
    ws.sheet_properties.outlinePr.applyStyles = False
    # Filter spans the header + data only - never the TOTAL row below it.
    ws.auto_filter.ref = (
        f"A{header_row}:{get_column_letter(grid.width)}{last_data_row}"
    )
    _frame_money(ws, grid, header_row, last_data_row + 1)
    _autofit(ws, grid, last_data_row, headers, uncapped)
    ws.sheet_view.showGridLines = False  # designed look; structure carried by fills/rules
    ws.sheet_view.zoomScale = ZOOM_PCT
    # Printing: landscape, every column on one page width, the header row repeated
    # on each page, narrow margins - so a printed aging reads like the screen.
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = f"{header_row}:{header_row}"
    ws.page_margins.left = ws.page_margins.right = 0.3
    ws.page_margins.top = ws.page_margins.bottom = 0.5
