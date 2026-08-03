#!/usr/bin/env python3
"""
cp_wip_reader.py — CP division WIP row builder.

Reads Commercial takeoffs from Synology, joins each project with its QBO
Billed/Costs, and writes results to the "Test - CP" tab of the SharePoint
WIP file. Live tabs (WIP - CP, WIP Master) are never touched — the guard
in wip_excel_guard.py enforces this at the code level.

Per-project extraction rules (locked 2026-06-30):
  Contract Price = cell immediately right of "GRAND TOTAL" on sheet
                   "Commercial Proposal"
  ETC            = cell AP1961 on sheet "Bid"
  Change Orders  = from the draw only (G702 Line 2). A project with no draw
                   yet has no approved COs, so takeoff CO sheets and the
                   Change Orders/ sub-folder are NOT read.

Failure modes never crash the run — they surface as a Status column value
so the user can triage from the Test - CP tab.

Data sources:
  - Synology (READ-ONLY):
      Active:    /Volumes/Common/CURRENT PROJECTS/Awarded Projects Commercial projects
      Completed: same path + /Completed Projects
  - QBO (READ-ONLY GET): via project-pnl/project_pnl_export.py reused
    helpers (load_credentials, build_project_customer_map, fetch_project_pl,
    extract_pl_totals).
  - OneDrive local mirror of SharePoint (WRITE, Test - CP only):
      ~/Library/CloudStorage/OneDrive-ProficientConcrete,LLC/
        Company Files - WIP Report/WIP - MASTER new.xlsx

Usage:
  python3 cp_wip_reader.py --dry-run                     # preview, no write
  python3 cp_wip_reader.py --project CP672               # one project only
  python3 cp_wip_reader.py                                # live run, all CP
  python3 cp_wip_reader.py --no-qbo                       # skip QBO join
                                                          #   (fast local test)

Overridable via env or flags:
  CP_ACTIVE_DIR, CP_COMPLETED_DIR, WIP_EXCEL_PATH
"""
from __future__ import annotations

import argparse
import datetime as dt
import logging
import os
import re
import sys
import warnings
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple

# Silence known-benign noise BEFORE importing openpyxl / requests.
# 1. openpyxl warns on cross-sheet INDIRECT() print areas (harmless — we're
#    only reading data, print settings don't matter).
# 2. macOS system Python 3.9 ships with LibreSSL; urllib3 v2 prefers OpenSSL.
#    HTTPS still works fine over LibreSSL; the warning is cosmetic.
warnings.filterwarnings("ignore", message="Print area cannot be set to Defined name.*")
warnings.filterwarnings("ignore", message=".*LibreSSL.*")
try:
    import urllib3
    urllib3.disable_warnings(urllib3.exceptions.NotOpenSSLWarning)  # type: ignore[attr-defined]
except Exception:
    pass

from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from wip_excel_guard import (
    ALLOWED_WRITE_SHEETS,
    WipWriteDenied,
    assert_write_allowed,
    open_wip_workbook_for_write,
)


# ─────────────────────── pretty terminal output ───────────────────
class _Term:
    """Small ANSI helper — degrades gracefully when stdout isn't a TTY."""
    ENABLED = sys.stdout.isatty()

    RESET = "\033[0m"     if ENABLED else ""
    BOLD  = "\033[1m"     if ENABLED else ""
    DIM   = "\033[2m"     if ENABLED else ""
    CYAN  = "\033[36m"    if ENABLED else ""
    GREEN = "\033[32m"    if ENABLED else ""
    AMBER = "\033[33m"    if ENABLED else ""
    RED   = "\033[31m"    if ENABLED else ""
    GRAY  = "\033[90m"    if ENABLED else ""

    @staticmethod
    def color(code: str, text: str) -> str:
        return f"{code}{text}{_Term.RESET}"


_ANSI_RE = re.compile(r"\033\[[0-9;]*m")

def _vlen(s: str) -> int:
    """Visible length (strips ANSI escape codes so padding math works)."""
    return len(_ANSI_RE.sub("", s))

def _rpad(s: str, width: int) -> str:
    """Right-pad to `width` visible chars (for left-aligned columns)."""
    return s + " " * max(0, width - _vlen(s))

def _lpad(s: str, width: int) -> str:
    """Left-pad to `width` visible chars (for right-aligned columns)."""
    return " " * max(0, width - _vlen(s)) + s


def _section(title: str) -> None:
    """Print a section header."""
    print()
    print(_Term.color(_Term.BOLD + _Term.CYAN, f"▸ {title}"))
    print(_Term.color(_Term.DIM, "─" * (len(title) + 2)))


def _kv(label: str, value, indent: int = 2) -> None:
    """Print a key/value pair aligned."""
    print(f"{' ' * indent}{label:<18} {value}")


def _fmt_money(v: Optional[float]) -> str:
    """Money as a plain (uncolored) string. Coloring happens at pad time."""
    if v is None:
        return "—"
    return f"${v:,.2f}"


def _fmt_pct(v: Optional[float]) -> str:
    if v is None:
        return "—"
    return f"{v*100:.1f}%"


def _dim_if_dash(s: str) -> str:
    """Dim the em-dash placeholder so real numbers pop; leave money plain."""
    return _Term.color(_Term.DIM, s) if s == "—" else s


# ─────────────────────── config / paths ────────────────────────────
PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))
from shared import paths
from shared import qbo_api  # QBO helpers (formerly loaded from project-pnl by raw file path)

CP_ACTIVE_DIR = Path(os.getenv(
    "CP_ACTIVE_DIR",
    "/Volumes/Common/CURRENT PROJECTS/Awarded Projects Commercial projects",
))
CP_COMPLETED_DIR = Path(os.getenv(
    "CP_COMPLETED_DIR",
    str(CP_ACTIVE_DIR / "Completed Projects"),
))
WIP_EXCEL_PATH = paths.get_path(
    "WIP_EXCEL_PATH",
    paths.onedrive_base() / "Company Files - WIP Report/WIP - MASTER new.xlsx",
)

TEST_TAB = "Test - CP"  # write target — must be in ALLOWED_WRITE_SHEETS
MASTER_TITLE_SHEET = "WIP Master"   # read-only: the formatting reference sheet

# Cell anchors on the takeoff (locked)
PROPOSAL_SHEET = "Commercial Proposal"
GRAND_TOTAL_LABEL = "GRAND TOTAL"
BID_SHEET = "Bid"
ETC_CELL = "AP1961"

# Draw discovery + G702 parsing moved to shared/draws.py (2026-07-16) — the
# money-bleeds health report needed it too (repo rule: tools never import
# tools; common code lives in shared/). Aliased to keep local call sites.
from shared.draws import (                                    # noqa: E402
    G702_SHEET,
    coerce_float as _coerce_float,
    find_latest_draw,
    read_draw_g702,
)

# Project # from folder name — e.g. "CP672 - FIRESTONE RED OAK" → "CP672"
_CP_FOLDER_RE = re.compile(r"^(CP\d{3,4})\b", re.IGNORECASE)


log = logging.getLogger("cp_wip_reader")


# QBO helpers come from shared/qbo_api.py (2026-07-13 restructure) — the old
# importlib load of project-pnl/project_pnl_export.py is gone.

# Realm (company id) captured by enrich_with_qbo so the Excel writer can build
# QBO deep links (customer page / project P&L report) on the number cells.
QBO_REALM = ""


# ─────────────────────── data models ───────────────────────────────
@dataclass
class CpRow:
    """One row per CP project, ready to write to Test - CP."""
    project_num: str
    project_name: str                # from folder name after " - "
    is_completed: bool               # from Completed Projects/ subfolder
    base_contract: Optional[float]   # from Commercial Proposal Grand Total (audit)
    co_revenue: Optional[float]      # approved COs from the draw (G702 Line 2);
                                     # None until Draw #1 (no draw ⇒ no COs yet)
    base_etc: Optional[float]        # from Bid!AP1961 (audit — pre-CO)
    billed_to_date: Optional[float]  # QBO P&L income = GROSS billed (incl retainage)
    costs_to_date: Optional[float]   # from QBO P&L COGS + Expenses
    retainage_held: Optional[float] = None  # gross billed − net collectible (retainage receivable)
    status_flags: List[str] = field(default_factory=list)  # TRUE flags only: the script
                                          # hit something it could not confirm as fact and a
                                          # human must fix it (unreadable takeoff, QBO failure,
                                          # ambiguous proposal). NOT business observations —
                                          # the WIP report shows over-budget / CO $ itself.
    notes: List[str] = field(default_factory=list)  # informational; the script IS certain
                                          # (e.g. 'Draw #6…', 'No draw yet') — never a flag.
    takeoff_path: Optional[Path] = None   # audit trail: first included takeoff (hyperlink anchor)
    included_takeoffs: List[Path] = field(default_factory=list)  # takeoff(s) summed into this row
    folder_path: Optional[Path] = None   # explicit folder for the project-name link
                                          # (RP sets this; CP falls back to takeoff_path.parent)
    draw_num: Optional[int] = None       # latest draw # that sourced contract/billed/retainage
                                          # (None = pre-Draw#1 → takeoff proposal + QBO instead)
    draw_path: Optional[Path] = None     # the latest draw's G702/G703 workbook (audit trail)
    qbo_customer_id: Optional[str] = None  # QBO customer id → deep links on Billed/Costs cells
    needs_review: bool = False           # number doesn't look right / flagged → red font in Excel
    client: Optional[str] = None         # builder/client display name (RP tab)
    home_type: Optional[str] = None      # 'Tract' / 'Custom' (RP tab)
    why_link: Optional[str] = None       # path to the justification workbook (temp WHY column)
    why_fragment: Optional[str] = None   # "#'SHEET'!A<row>" — jump straight to this line's row
    src_link: Optional[str] = None       # source workbook the numbers came from (PROJECT # cell link)
    src_fragment: Optional[str] = None   # "#'SHEET'!C<row>" — exact source row (the user 2026-07-15)
    section: Optional[str] = None        # master-sheet grouping (SECTION column)

    @property
    def contract_price(self) -> Optional[float]:
        """Total Contract Price = base contract + summed COs. Returns None when
        the base is undetermined (e.g. Missing Grand Total / contract not yet
        decided) — a CO-only total would be misleading, so the cell stays blank
        and the flag explains why."""
        if self.base_contract is None:
            return None
        return self.base_contract + (self.co_revenue or 0.0)

    @property
    def co_cost_estimate(self) -> Optional[float]:
        """CO Cost is None until estimators add a real cost cell to the CO
        template. No proxy — the user's rule (2026-07-01): "no false numbers,
        don't populate if there is no source. If no CO costs, don't put
        it, just flag it."

        Property name kept as `co_cost_estimate` for column-mapping
        stability; semantics are now "CO cost sourced from the draw/template"
        and the value is None until such sourcing exists (blocked on a
        template change that adds a CO cost line)."""
        return None

    @property
    def etc(self) -> Optional[float]:
        """Revised ETC.
        - No COs → equals Base ETC (revised == base is truthful here).
        - COs present but no CO Cost data → defaults to Base ETC and the
          row is flagged provisional (the user 2026-07-02). This WIP is a
          MONITORING view, fully script-driven from takeoffs — no manual
          entry. Defaulting to Original ETC keeps the row alive and
          surfaces the over-budget signal (Costs > Original ETC) instead
          of going dark. The bounded error (excludes the CO's cost) is
          disclosed by the `% based on Original ETC — excludes CO cost`
          flag. We still never FABRICATE a CO cost — the base ETC is a
          real number, just scoped pre-CO. See [[project-cp-wip-takeoff-extraction]]."""
        if self.base_etc is None:
            return None
        if not self.co_revenue:
            return self.base_etc                # no COs → Revised = Base
        if self.co_cost_estimate is None:
            return self.base_etc                # CO Rev but no CO Cost → provisional (Original ETC)
        return self.base_etc + self.co_cost_estimate

    @property
    def has_missing_co_cost(self) -> bool:
        """True when the project has CO Revenue but no CO Cost data — a
        signal that Revised ETC (and the % / Earned / Over-Under derived
        from it) is PROVISIONAL: computed off Original ETC, which excludes
        the CO's cost. Drives the provisional flag."""
        return bool(self.co_revenue) and self.co_cost_estimate is None

    @property
    def status(self) -> str:
        if not self.status_flags:
            return "OK"
        return "; ".join(self.status_flags)

    @property
    def notes_text(self) -> str:
        """Informational notes joined for the NOTES column (blank when none)."""
        return "; ".join(self.notes)


# ─────────────────────── takeoff parsing ───────────────────────────
def _find_label_cells(ws, label: str) -> List[tuple]:
    """Return ALL (row, col) matches for the given label (case-insensitive,
    trailing colon/whitespace stripped).

    Returns a list because CO templates have TWO 'TOTAL:' cells: a header
    label at the top of the description table AND the summary total at
    the bottom. Caller must try each until one yields a numeric value to
    the right (see _read_number_to_right)."""
    label_norm = label.strip().upper().rstrip(":").strip()
    results: List[tuple] = []
    for row in ws.iter_rows(values_only=False):
        for cell in row:
            v = cell.value
            if v is None:
                continue
            s = str(v).strip().upper().rstrip(":").strip()
            if s == label_norm:
                results.append((cell.row, cell.column))
    return results


def _resolve_sheet_name(wb, target: str) -> Optional[str]:
    """Case-insensitive, whitespace-tolerant sheet name lookup. Returns the
    ACTUAL sheet name as it appears in the workbook, or None if not found.
    Handles: 'BID' vs 'Bid' vs 'bid ' vs '  Commercial Proposal '."""
    target_norm = target.strip().lower()
    for name in wb.sheetnames:
        if name.strip().lower() == target_norm:
            return name
    return None


def _read_number_to_right(ws_data, ws_formula, row: int, start_col: int,
                          max_scan: int = 12) -> Optional[float]:
    """Scan rightward from `start_col` looking for the first numeric value.

    Skips empty cells (None), cells that are part of a merged region but
    aren't the top-left anchor, and cells that hold just a currency
    symbol like '$'. This handles two real-world CP takeoff patterns:

      1. TOTAL: label lives in a merged cell — the value sits past the
         merge's right edge, not one column right of the label anchor.
      2. Currency symbol is a separate cell from the number:
             [ ...merged label... ] [ $ ] [ 3,840.00 ]

    Returns the first cell that resolves to a number via
    `_read_number_smart` (which itself handles formula fallback). None
    if nothing found within `max_scan` columns."""
    for offset in range(1, max_scan + 1):
        col = start_col + offset
        try:
            ref = f"{get_column_letter(col)}{row}"
        except ValueError:
            return None
        # Try the smart reader first — handles cached values + formula refs.
        v = _read_number_smart(ws_data, ws_formula, ref)
        if v is not None:
            return v
        # If not numeric, check if the raw string is a currency marker; if
        # so, keep scanning. Any other non-empty non-numeric value → treat
        # as a stop signal (we've passed the value region).
        raw = ws_data[ref].value
        if raw is None or raw == "":
            continue
        s = str(raw).strip()
        if s in ("$", "USD", "€", "£", "-", "—"):
            continue
        # Non-numeric text (like a next-row label) — bail.
        return None
    return None


# ─────────────────────── formula-safe cell reader ────────────────
# openpyxl(data_only=True) returns cached formula results. If the takeoff was
# saved without recalc (common when files are edited via Google Sheets export,
# LibreOffice, or a Python tool), the cache is empty and formula cells come
# back as None. In that case we open a SECOND view of the workbook (data_only
# =False) to see the formula string, then evaluate it dynamically against the
# cached view — so we NEVER hardcode the specific cell-relationship (the
# estimators can change AP1961's formula without breaking us).
_CELL_REF_RE = re.compile(r"^[A-Z]+\d+$")
_TOKEN_RE    = re.compile(r"[A-Z]+\d+|[+\-]")


def _read_number_smart(ws_data, ws_formula, cell_ref: str, depth: int = 0) -> Optional[float]:
    """Read a numeric value from `cell_ref`. Order of attempts:
       1. Cached value from data_only workbook.
       2. If the cell holds a formula, evaluate simple +/- expressions of
          cell references (with recursive fallback for each ref).

    Recursion depth is bounded to prevent runaway on circular refs."""
    if depth > 4:
        return None
    try:
        cached = ws_data[cell_ref].value
    except (KeyError, ValueError):
        return None
    if cached is not None:
        coerced = _coerce_float(cached)
        if coerced is not None:
            return coerced

    # Cached is missing / non-numeric — try the formula view.
    try:
        raw = ws_formula[cell_ref].value
    except (KeyError, ValueError):
        return None
    if not isinstance(raw, str) or not raw.strip().startswith("="):
        return None

    expr = raw.strip().lstrip("=").replace(" ", "").upper()
    # Support only simple +/- chains of cell refs (e.g. "AM1948+AN1948+AO1948").
    # Anything more complex (SUM, ranges, other sheets) falls back to None +
    # the caller will flag the row for manual review.
    if not re.fullmatch(r"[A-Z]+\d+([+\-][A-Z]+\d+)*", expr):
        return None
    tokens = _TOKEN_RE.findall(expr)
    if not tokens:
        return None

    total = 0.0
    sign = 1.0
    expect_ref = True
    for tok in tokens:
        if tok in ("+", "-"):
            if expect_ref:
                return None  # malformed
            sign = 1.0 if tok == "+" else -1.0
            expect_ref = True
        else:
            if not expect_ref or not _CELL_REF_RE.match(tok):
                return None
            v = _read_number_smart(ws_data, ws_formula, tok, depth + 1)
            if v is None:
                return None
            total += sign * v
            expect_ref = False
            sign = 1.0
    return total


_WIP_TAG_RE = re.compile(r"\bwip\b", re.IGNORECASE)
# Auxiliary xlsx that live alongside takeoffs in a project folder and are NOT
# takeoffs (used only as a fallback when no file is named 'takeoff').
_AUX_XLSX_RE = re.compile(r"cost\s*code|explanation", re.IGNORECASE)


def _find_contract_total(ws_data, ws_formula):
    """Read the contract total off a proposal sheet. Templates are inconsistent
    (the user 2026-07-02) — the overall total is labeled 'GRAND TOTAL', 'SUB TOTAL',
    or just 'TOTAL'. Match is EXACT (normalized, colon-stripped) so 'TOTAL SQFT',
    'TOTAL YARDS', etc. are NOT caught — only cells that are exactly one of the
    three dollar-total labels. Collect every value from all three and take the
    **largest** = the overall contract total (an overall is ≥ any section
    subtotal). Returns (value_or_None, label_used_or_None)."""
    best = None
    best_label = None
    for label in ("GRAND TOTAL", "SUB TOTAL", "TOTAL"):
        for (r, c) in _find_label_cells(ws_data, label):
            v = _read_number_to_right(ws_data, ws_formula, r, c)
            if v is not None and (best is None or v > best):
                best = v
                best_label = label
    return best, best_label


def _select_proposal_sheet(wb):
    """Pick which proposal tab to read the contract GRAND TOTAL from, when a
    takeoff has multiple proposal sheets (the user 2026-07-02):
      - a tab with 'final' in its name → the final proposal, use it;
      - else if exactly ONE proposal tab → use it;
      - else (multiple proposals, none marked FINAL) → don't guess, flag it.
    'Proposal' tabs = any sheet whose name contains 'proposal' (Commercial /
    Residential / Alternative / '11.14.25 Proposal' …). Non-proposal tabs
    (Bid, JMP Subcontract, Change Order#N, Cost Codes) are ignored here.
    Returns (sheet_name_or_None, flag_or_None)."""
    proposals = [s for s in wb.sheetnames if "proposal" in s.lower()]
    if not proposals:
        return None, "Missing Proposal Sheet"
    finals = [s for s in proposals if "final" in s.lower()]
    if len(finals) == 1:
        return finals[0], None
    if len(finals) > 1:
        return finals[0], f"Multiple FINAL proposals — used '{finals[0]}'"
    if len(proposals) == 1:
        return proposals[0], None
    return None, (f"Multiple proposals ({len(proposals)}), none marked FINAL — "
                  f"mark the final one: {', '.join(proposals[:4])}"
                  f"{'...' if len(proposals) > 4 else ''}")


def _parse_one_takeoff(tk: Path):
    """Read Contract Price (final proposal Grand Total) + ETC (Bid!AP1961)
    from ONE takeoff file. Change Orders are NOT read here — approved COs only
    ever come from a draw, and a no-draw project (the only caller of this path)
    has no COs yet. Returns (contract, etc, flags). Never raises."""
    flags: List[str] = []
    # Two views: cached values (data_only) fast path + formulas (fallback for
    # cells saved without recalc). Random-access needed → not read_only.
    try:
        wb_data    = load_workbook(tk, data_only=True)
        wb_formula = load_workbook(tk, data_only=False)
    except Exception as e:
        return None, None, [f"Takeoff Read Failed: {type(e).__name__}"]

    contract = etc = None
    try:
        prop_sheet, prop_flag = _select_proposal_sheet(wb_data)
        if prop_flag:
            flags.append(prop_flag)
        if prop_sheet is not None:
            cp, label_used = _find_contract_total(wb_data[prop_sheet],
                                                  wb_formula[prop_sheet])
            if cp is None:
                has_label = (_find_label_cells(wb_data[prop_sheet], "GRAND TOTAL")
                             or _find_label_cells(wb_data[prop_sheet], "SUB TOTAL"))
                flags.append("Bad Contract Price" if has_label
                             else "Missing Grand/Sub Total")
            else:
                contract = cp

        bid_sheet = _resolve_sheet_name(wb_data, BID_SHEET)
        if not bid_sheet:
            flags.append("Missing Bid Sheet")
        else:
            e = _read_number_smart(wb_data[bid_sheet], wb_formula[bid_sheet], ETC_CELL)
            if e is None:
                flags.append(f"Bad ETC ({ETC_CELL})")
            else:
                etc = e

    finally:
        wb_data.close()
        wb_formula.close()
    return contract, etc, flags


def _select_takeoffs(folder: Path):
    """Identify which takeoff file(s) to read in a project folder. Shared by
    parse_takeoff (full contract+ETC read) and parse_takeoff_etc (ETC-only,
    used when a draw supplies the contract). Rules (the user 2026-07-02):
      - Only files with 'takeoff' in the name are takeoffs; auxiliary xlsx
        (Cost Codes, Explanation OH, …) are ignored. If none is named
        'takeoff', fall back to the single non-auxiliary xlsx.
      - ONE takeoff → use it; MULTIPLE → only the 'WIP'-tagged one(s), summed.
      - Multiple, none tagged WIP → don't guess.
    Returns (included_list, flag_or_None)."""
    xlsx_files = sorted([p for p in folder.iterdir()
                         if p.suffix.lower() in (".xlsx", ".xlsm")
                         and not p.name.startswith("~$")])
    if len(xlsx_files) == 0:
        return [], "No Takeoff"
    takeoffs = [p for p in xlsx_files if "takeoff" in p.name.lower()]
    if not takeoffs:
        non_aux = [p for p in xlsx_files if not _AUX_XLSX_RE.search(p.name)]
        if len(non_aux) == 1:
            takeoffs = non_aux
        else:
            return [], (f"No takeoff file identified ({len(xlsx_files)} xlsx, none "
                        f"named 'takeoff') — rename the takeoff to include 'takeoff'")
    if len(takeoffs) == 1:
        return takeoffs, None
    included = [p for p in takeoffs if _WIP_TAG_RE.search(p.name)]
    if not included:
        return [], (f"Multiple takeoffs ({len(takeoffs)}) — none tagged 'WIP'; "
                    f"estimator must tag the one(s) to include")
    return included, None


def parse_takeoff(folder: Path, row: CpRow) -> None:
    """Extract Contract Price + ETC into `row` from the takeoff. Used when the
    project has NO draw yet (pre-Draw#1) — contract comes from the proposal
    Grand/Sub Total, ETC from Bid!AP1961. No COs are read: a project that
    hasn't started billing has no approved change orders yet (COs come from a
    draw). Appends to row.status_flags on any failure; never raises."""
    included, flag = _select_takeoffs(folder)
    if flag:
        row.status_flags.append(flag)
    if not included:
        return

    row.included_takeoffs = included
    row.takeoff_path = included[0]                  # hyperlink anchor

    contract_total = 0.0
    etc_total = 0.0
    got_contract = False
    got_etc = False
    multi = len(included) > 1
    for tk in included:
        c, e, fflags = _parse_one_takeoff(tk)
        prefix = f"{tk.name}: " if multi else ""
        for f in fflags:
            row.status_flags.append(prefix + f)
        if c is not None:
            contract_total += c
            got_contract = True
        if e is not None:
            etc_total += e
            got_etc = True

    row.base_contract = contract_total if got_contract else None
    row.base_etc = etc_total if got_etc else None

    if multi:
        row.notes.append(
            f"WIP takeoffs summed ({len(included)}): "
            f"{', '.join(p.name for p in included)}")


# ─────────────────────── ETC-only takeoff read (draw path) ────────
def parse_takeoff_etc(folder: Path, row: CpRow) -> None:
    """Read ONLY the ETC (Bid!AP1961) from the project's takeoff. Used when a
    DRAW supplies contract/CO/billed/retainage but the cost estimate still
    lives in the takeoff (the user 2026-07-09: "ETC — still keep the takeoff
    costs"). The proposal/contract/CO parsing is skipped, so a draw-backed row
    isn't cluttered with contract-side takeoff flags. Never raises."""
    included, flag = _select_takeoffs(folder)
    if not included:
        # Missing takeoff only matters for ETC now; keep the message specific.
        row.status_flags.append(f"ETC: {flag}" if flag else "ETC: no takeoff")
        return
    if row.takeoff_path is None:
        row.takeoff_path = included[0]
    row.included_takeoffs = included

    etc_total = 0.0
    got_etc = False
    multi = len(included) > 1
    for tk in included:
        try:
            wb_data = load_workbook(tk, data_only=True)
            wb_formula = load_workbook(tk, data_only=False)
        except Exception as e:
            row.status_flags.append(f"ETC read failed ({tk.name}): {type(e).__name__}")
            continue
        try:
            bid_sheet = _resolve_sheet_name(wb_data, BID_SHEET)
            prefix = f"{tk.name}: " if multi else ""
            if not bid_sheet:
                row.status_flags.append(prefix + "Missing Bid Sheet (ETC)")
                continue
            e = _read_number_smart(wb_data[bid_sheet], wb_formula[bid_sheet], ETC_CELL)
            if e is None:
                row.status_flags.append(prefix + f"Bad ETC ({ETC_CELL})")
            else:
                etc_total += e
                got_etc = True
        finally:
            wb_data.close()
            wb_formula.close()
    if got_etc:
        row.base_etc = etc_total


# ─────────────────────── draw (G702/G703) read ────────────────────
def parse_draw(project_folder: Path, row: CpRow) -> bool:
    """If the project has a draw, read the LATEST one's G702 into `row`
    (contract, CO, billed, retainage) and return True. ETC still comes from the
    takeoff and costs from QBO — those are handled by the caller. Returns False
    when there's no draw yet (caller falls back to the takeoff proposal)."""
    found = find_latest_draw(project_folder)
    if not found:
        return False
    draw_num, draw_file = found
    row.draw_num = draw_num
    row.draw_path = draw_file

    data, flags = read_draw_g702(draw_file)
    for f in flags:
        row.status_flags.append(f"Draw #{draw_num}: {f}")
    if data is None or data["contract_to_date"] is None or data["billed"] is None:
        # A draw exists but is unreadable — do NOT fall through to takeoff/QBO
        # for billing (that would silently mix sources). Flag for triage.
        row.status_flags.append(
            f"Draw #{draw_num} unreadable — contract/billed left blank for review")
        return True

    row.base_contract = data["orig_contract"]
    row.co_revenue = data["net_co"]
    row.billed_to_date = data["billed"]
    row.retainage_held = data["retainage"]
    row.notes.append(
        f"Draw #{draw_num}: billed ${data['billed']:,.0f} (gross), "
        f"retainage ${(data['retainage'] or 0):,.0f}, "
        f"contract ${(data['contract_to_date'] or 0):,.0f}")
    return True


# ─────────────────────── folder scan ───────────────────────────────
def _project_num_from_folder(folder: Path) -> Optional[str]:
    m = _CP_FOLDER_RE.match(folder.name)
    return m.group(1).upper() if m else None


def _project_name_from_folder(folder: Path) -> str:
    """'CP672 - FIRESTONE RED OAK' → 'FIRESTONE RED OAK'"""
    if " - " in folder.name:
        return folder.name.split(" - ", 1)[1].strip()
    return folder.name


def scan_cp_folders(root: Path, is_completed: bool) -> List[CpRow]:
    """Iterate CP#### folders under root. Returns one CpRow per project.
    Skips non-CP folders silently (folder naming discipline is on the
    estimators). Handles Change Orders/ sub-folder = flag-and-skip."""
    rows: List[CpRow] = []
    if not root.exists():
        log.warning("CP root does not exist: %s (Synology unmounted?)", root)
        return rows

    for entry in sorted(root.iterdir()):
        if not entry.is_dir():
            continue
        if entry.name == "Completed Projects":  # handled in a separate pass
            continue

        proj_num = _project_num_from_folder(entry)
        if not proj_num:
            continue  # not a CP folder — skip silently

        row = CpRow(
            project_num=proj_num,
            project_name=_project_name_from_folder(entry),
            is_completed=is_completed,
            base_contract=None,
            co_revenue=None,
            base_etc=None,
            billed_to_date=None,
            costs_to_date=None,
        )

        # Draw-first (the user 2026-07-09): if the project has a draw (AIA G702/G703
        # payment application), the LATEST draw IS the WIP update — it supplies
        # Contract Price, Approved COs, Billed-to-Date (gross), and Retainage
        # Held. ETC still comes from the takeoff; Costs from QBO. Only before
        # Draw #1 lands do we fall back to the takeoff proposal for contract/CO
        # and QBO for billed/retainage.
        if parse_draw(entry, row):
            row.folder_path = entry                 # project-name link target
            parse_takeoff_etc(entry, row)           # ETC (Bid!AP1961) only
        else:
            # No draw yet — takeoff proposal drives contract; QBO drives
            # billed/retainage (in enrich_with_qbo). No COs: a project that
            # hasn't started its first draw has no approved change orders yet,
            # so the Change Orders/ sub-folder is intentionally not read.
            row.notes.append("No draw yet — contract from takeoff proposal")
            parse_takeoff(entry, row)

        rows.append(row)

    return rows


# ─────────────────────── QBO join ──────────────────────────────────
# Retainage detection — case-insensitive match on "Retainage Not Billed"
# anywhere in the invoice's PrivateNote (memo field). Pattern harvested
# from legacy wip/wip_sync.py; see [[reference-wip-qbo-aggregation-patterns]].
_RETAINAGE_NOT_BILLED_RE = re.compile(r"retainage\s+not\s+billed", re.IGNORECASE)


def _linked_txn_types(inv: dict) -> List[str]:
    """Normalized TxnType strings of every transaction linked to `inv`
    (payments, journal entries, credits) from QBO's LinkedTxn array.
    Spaces stripped + lowercased so 'JournalEntry' / 'Journal Entry' both
    normalize to 'journalentry'."""
    out = []
    for lt in (inv.get("LinkedTxn") or []):
        t = str(lt.get("TxnType", "")).replace(" ", "").lower()
        if t:
            out.append(t)
    return out


def _cleared_by_journal_entry(inv: dict) -> bool:
    """True if a Journal Entry is linked to this invoice — the reclass that
    moves the balance into the Retainage Receivable account (i.e. it was NOT
    settled by a real customer cash payment)."""
    return "journalentry" in _linked_txn_types(inv)


def _is_retainage_receivable_invoice(inv: dict) -> bool:
    """True if this is a standalone retainage-receivable invoice → its amount
    is retainage HELD (not collectible cash). It is excluded from Billed's
    net-collectible component and surfaces in the RETAINAGE HELD column.

    Keyed on the 'Retainage Not Billed' memo — the user's explicit tag for
    retainage moved to (or awaiting) the Retainage Receivable account.

    CORRECTED 2026-07-02 (live run): an earlier version ALSO required a
    Journal Entry on the invoice's LinkedTxn. That left retainage held
    understated by an order of magnitude on CP672 because the JE
    that reclasses AR → Retainage Receivable is applied to the invoice via a
    PAYMENT (the "Payment on 12/31/25" that carries the JE credit) — so the
    JE never appears directly in Invoice.LinkedTxn and the check missed it.
    The memo is the reliable signal and correctly handles BOTH methods:
      * standalone 'Retainage Not Billed' invoice → matched here → held.
      * 2026 draw-with-retainage-LINE → memo is 'Draw #N' (not matched); its
        withheld portion is captured via gross income − net TotalAmt.
    (`_cleared_by_journal_entry` is retained for audit logging only.)"""
    memo = inv.get("PrivateNote") or ""
    return bool(_RETAINAGE_NOT_BILLED_RE.search(memo))


def _fetch_billed_ex_retainage(pnl, access: str, company_id: str,
                               customer_id: str) -> float:
    """Sum TotalAmt of all invoices for a customer, EXCLUDING standalone
    retainage-receivable invoices (memo 'Retainage Not Billed'). This is the
    NET-COLLECTIBLE billed; the caller derives Retainage Held = gross − net.
    Returns 0.0 if the query fails or there are no invoices.

    Every retainage-memo invoice is logged with its LinkedTxn types + whether
    a JE is directly attached, so the retainage reclass path stays auditable."""
    try:
        invoices = pnl.fetch_customer_invoices(access, company_id, customer_id)
    except Exception as e:
        log.warning("Invoice fetch failed for customer %s: %s — falling back to 0",
                    customer_id, e)
        return 0.0

    total = 0.0
    skipped = 0
    for inv in invoices:
        if _is_retainage_receivable_invoice(inv):
            skipped += 1
            doc = inv.get("DocNumber", "?")
            amt = inv.get("TotalAmt", "?")
            types = _linked_txn_types(inv) or ["<none>"]
            je = "direct" if _cleared_by_journal_entry(inv) else "via-payment/none"
            log.info("  EXCL retainage-receivable inv #%s ($%s) — memo-tagged "
                     "(LinkedTxn=%s, JE=%s) → counts as Retainage Held",
                     doc, amt, ",".join(types), je)
            continue
        total += float(inv.get("TotalAmt", 0) or 0)

    if skipped:
        log.info("Customer %s: %d retainage-receivable invoice(s) → Retainage Held",
                 customer_id, skipped)
    return total


def enrich_with_qbo(rows: List[CpRow]) -> None:
    """Fetch QBO Billed/Costs per project and populate rows in-place.
    All-time window (start_date = 2019-01-01, end_date = today) — CP is
    slow-turn commercial work, worth the extra API cost. Never raises.
    Side effect: records the realm + per-row customer id so the Excel writer
    can attach QBO deep links to the Billed/Costs cells."""
    global QBO_REALM
    pnl = qbo_api
    try:
        access, company_id = pnl.load_credentials()
    except SystemExit:
        log.error("QBO auth failed — leaving Billed/Costs blank on all rows")
        for r in rows:
            r.status_flags.append("QBO Auth Failed")
        return
    except Exception as e:
        log.error("QBO auth error: %s", e)
        for r in rows:
            r.status_flags.append(f"QBO Auth Error: {type(e).__name__}")
        return

    try:
        proj_map = pnl.build_project_customer_map(access, company_id)
    except Exception as e:
        log.error("QBO customer map fetch failed: %s", e)
        for r in rows:
            r.status_flags.append(f"QBO Customer Map Failed: {type(e).__name__}")
        return

    QBO_REALM = company_id
    start_date = "2019-01-01"
    end_date = dt.date.today().isoformat()

    for row in rows:
        cust = proj_map.get(row.project_num)
        if not cust:
            row.status_flags.append("QBO Not Found")
            continue
        row.qbo_customer_id = cust["id"]

        try:
            report_data = pnl.fetch_project_pl(
                access, company_id, cust["id"], start_date, end_date
            )
            totals = pnl.extract_pl_totals(report_data)
            # Billed to Date + Retainage Held come from the DRAW when the
            # project has one (the user 2026-07-09: the latest G702 is the billing
            # source of record). Only PRE-Draw#1 projects fall back to QBO for
            # billing. This also saves the per-customer invoice fetch on
            # draw-backed jobs.
            if row.draw_num is None:
                # QBO fallback (no draw yet). Billed = GROSS, incl retainage
                # (Marcum/CFMA basis, verified 2026-07-02): QBO P&L Total Income
                # is gross billed. Net-collectible excludes 'Retainage Not
                # Billed' memo invoices; Retainage Held = gross − net.
                gross_billed = float(totals.get("income", 0.0) or 0.0)
                net_collectible = _fetch_billed_ex_retainage(
                    pnl, access, company_id, cust["id"]
                )
                row.billed_to_date = gross_billed
                row.retainage_held = max(gross_billed - net_collectible, 0.0)
                log.info("  %s billed(gross)=%.2f net-collectible=%.2f retainage-held=%.2f (QBO)",
                         row.project_num, gross_billed, net_collectible, row.retainage_held)
            else:
                log.info("  %s billed/retainage from Draw #%s — QBO billing skipped",
                         row.project_num, row.draw_num)
            # Costs = COGS + Expenses. QBO Projects UI sums both; per GAAP
            # job costing, direct project spending is a project cost
            # regardless of which account category it lands in. Coding
            # errors in QBO (bills posted to Expenses instead of COGS)
            # should NOT hide costs from the WIP — the WIP should surface
            # true total spend so the coding error becomes visible.
            row.costs_to_date = (
                (totals.get("cogs", 0.0) or 0.0)
                + (totals.get("expenses", 0.0) or 0.0)
            )
            # Over-budget is NOT flagged: it's a business observation the WIP
            # report surfaces itself (Costs > ETC, and the uncapped % column).
            # Flags are reserved for things the script could not confirm as
            # fact and a human must fix — not report readings. % stays UNCAPPED
            # so the overage remains visible in the numbers.
        except Exception as e:
            row.status_flags.append(f"QBO P&L Failed: {type(e).__name__}")


# ─────────────────────── Excel write (Test - CP only) ──────────────
# Clean layout matching the team's WIP: light-gray bold header, thin grid
# borders on every cell, light-yellow ONLY on sourced inputs, white calcs,
# no green rows (the user 2026-07-02 "make it clean and easy to read").
# Font + number formats MATCH the real 'WIP Master' sheet (the user
# 2026-07-29: same font, size, formatting as the master): Tahoma 8
# throughout, no-cents currency, 0.00% percents.
MASTER_FONT_NAME = "Tahoma"
MASTER_FONT_SIZE = 8
HDR_FILL = PatternFill("solid", fgColor="D9D9D9")        # light gray header
HDR_FONT = Font(name=MASTER_FONT_NAME, size=MASTER_FONT_SIZE,
                bold=True, color="000000")
DATA_FONT = Font(name=MASTER_FONT_NAME, size=MASTER_FONT_SIZE)
FLAG_FILL = PatternFill("solid", fgColor="FFF2CC")       # flag cell
FLAG_FONT = Font(name=MASTER_FONT_NAME, size=MASTER_FONT_SIZE,
                 italic=True, color="7F6000")
INPUT_FILL = PatternFill("solid", fgColor="FFFF99")      # light yellow — SOURCED inputs
LINK_FONT = Font(name=MASTER_FONT_NAME, size=MASTER_FONT_SIZE,
                 color="0563C1", underline="single")     # Excel link
_SIDE = Side(style="thin", color="BFBFBF")
CELL_BORDER = Border(left=_SIDE, right=_SIDE, top=_SIDE, bottom=_SIDE)
# Title block, copied cell-for-cell from the real 'WIP Master' sheet.
TITLE_FONT = Font(name=MASTER_FONT_NAME, size=MASTER_FONT_SIZE, bold=True)
_MEDIUM = Side(style="medium", color="000000")
CURRENCY_FMT = '"$"#,##0_);[Red]\\("$"#,##0\\)'          # verbatim from 'WIP Master'
PCT_FMT = "0.00%"                                        # verbatim from 'WIP Master'

# Which fields render as currency vs percent (drives number_format in the write
# loop). Everything money-valued is currency; the two ratios are percent.
_MONEY_FIELDS = frozenset({
    "base_contract", "co_revenue", "contract_price", "base_etc",
    "co_cost_estimate", "etc", "original_profit", "costs_to_date",
    "cost_to_complete", "_earned_revenue", "profit_earned", "future_profit",
    "billed_to_date", "retainage_held", "left_to_bill", "overbillings",
    "underbillings", "job_borrow",
})
_PCT_FIELDS = frozenset({"_pct_complete", "gross_profit_pct"})

# SOURCED cells (raw numbers straight from the takeoff or QBO) get a yellow
# fill so they're visually distinct from the white CALCULATED cells (Excel
# formulas). The user 2026-07-02: "yellow = metrics that have sources, calculations
# leave white." Identifier/metadata columns (project #, name, status, flags,
# last synced) stay white — they're not metrics.
_SOURCE_FIELDS = frozenset({
    "contract_price",    # TOTAL CONTRACT PRICE (Original + COs)
    "etc",               # ESTIMATED TOTAL COSTS (Revised ETC)
    "billed_to_date",    # QBO (gross income)
    "costs_to_date",     # QBO
    # trailing cross-check columns (also sourced → yellow)
    "co_revenue",        # APPROVED COs
    "retainage_held",    # QBO (gross − net collectible)
})


def _sheet_fragment(sheet_name: str, cell_ref: str = "A1") -> str:
    """Build the '#SheetName!CellRef' fragment appended to a file:// URI so
    Excel opens the workbook AND jumps to the specified cell.

    Sheet names with spaces or special chars need single-quote wrapping in
    the URL fragment (Excel convention). Cell ref defaults to A1."""
    return f"#'{sheet_name}'!{cell_ref}"


def _apply_hyperlink(cell, target: Optional[Path], fragment: str = "") -> None:
    """Attach a file:// hyperlink to a cell and apply link styling.

    `target` may be a Path to a file OR a directory. If None, this is a
    no-op so callers can safely pass None when the source is missing.
    `fragment` is optional (e.g. \"#'Bid'!AP1961\" to jump to a cell).

    **Never attaches a hyperlink to a cell whose value is None.** openpyxl's
    behavior in that case is to display the hyperlink URL as the cell text,
    which produces confusing "why is there a URL in this blank cell?"
    output. If the cell is empty, there's nothing to click anyway.

    The sheet/cell jump goes in the hyperlink's `location` ATTRIBUTE, never
    appended to the URI (2026-07-21): a fragment like #'Small Jobs'!C7 puts
    raw spaces into the .rels target URI — invalid XML that makes Excel
    demand a repair on open."""
    from openpyxl.worksheet.hyperlink import Hyperlink
    if target is None:
        return
    if cell.value is None or cell.value == "":
        return
    try:
        uri = target.as_uri()
    except (ValueError, OSError):
        # Path can't be converted to URI (unusual on macOS/Linux) — skip
        # rather than crash the whole run.
        return
    loc = fragment.lstrip("#") if fragment else None
    cell.hyperlink = Hyperlink(ref=cell.coordinate, target=uri,
                               location=loc or None)
    cell.font = LINK_FONT

# Full standard construction-WIP column set. Formulas verified 2026-07-02
# against Marcum LLP / Construction Executive, Wouch Maloney CPAs, CFMA/
# AICPA-referenced guidance. Billed to Date is GROSS (incl retainage); a
# separate RETAINAGE HELD column keeps the cash-collection timeline visible.
# Column order mirrors the team's WIP schedule — reads left→right as a story:
# the deal (contract, cost) → progress (billed, spent) → projection (cost to
# complete, profit) → recognition (% , earned) → billing position (over/under,
# left to bill) → profitability (GP%, future profit, job borrow). The four
# yellow inputs (contract, cost, billed, costs) + retainage lead; everything
# after is derived. Draw #/no-draw context lives in the NOTES column; FLAGS is
# reserved for genuine script must-fix issues (kept clean per the user 2026-07-02).
COLS = [
    # ── Identifiers ──
    ("PROJECT #",                                       12, "project_num"),
    ("PROJECT NAME",                                    30, "project_name"),
    # PROJECT FOLDER / DATA SOURCE link columns REMOVED (the user 2026-07-29:
    # file hyperlinks weigh the workbook down — the only links anywhere are
    # the QBO deep links on Billed/Costs).
    ("STATUS",                                          10, "_active_status"),
    # ── Inputs (yellow) — the 4 core WIP inputs, matching the team sheet ──
    ("TOTAL CONTRACT PRICE",                            16, "contract_price"),   # Original + COs
    ("ESTIMATED TOTAL COSTS",                           16, "etc"),              # Revised ETC
    ("BILLED TO DATE",                                  15, "billed_to_date"),   # gross (incl retainage)
    ("COSTS TO DATE",                                   15, "costs_to_date"),
    # ── Derived story (white) ──
    ("COST TO COMPLETE",                                14, "cost_to_complete"),
    ("ORIGINAL PROFIT",                                 14, "original_profit"),
    ("PERCENT COMPLETE",                                11, "_pct_complete"),
    ("REVENUES EARNED TO DATE",                         16, "_earned_revenue"),
    ("PROFIT EARNED TO DATE",                           15, "profit_earned"),
    # Short names (the user 2026-07-16): the long "BILLINGS IN EXCESS OF…"
    # labels kept getting clipped every sync — width now fits the numbers.
    ("OVERBILLINGS",                                    14, "overbillings"),
    ("UNDERBILLINGS",                                   14, "underbillings"),
    ("LEFT TO BILL",                                    14, "left_to_bill"),
    ("GROSS PROFIT %",                                  12, "gross_profit_pct"),
    ("FUTURE PROFIT TO EARN",                           15, "future_profit"),
    ("PURE JOB BORROW",                                 14, "job_borrow"),
    # ── Trailing cross-check / reference (yellow, sourced) ──
    ("APPROVED COs",                                    14, "co_revenue"),
    ("RETAINAGE HELD",                                  14, "retainage_held"),
    ("LAST SYNCED",                                     18, "_last_synced"),
    # ONE commentary column (the user 2026-07-31: "stick to one" — NOTES and
    # FLAGS were two columns saying overlapping things). Informational notes
    # (Draw #, no-draw, the owner's ACTION text) and genuine script must-fix
    # flags now share this cell; the cell turns yellow/italic when it carries
    # a flag.
    ("NOTES",                                           60, "_notes_all"),
]


# Fields written as Excel FORMULAS (not static values) so any input change
# auto-recalculates and the formula is visible in the formula bar for audit.
# Mirrors the project-pnl workbook's formula-driven design.
FORMULA_FIELDS: frozenset = frozenset({
    # Total Contract Price (contract_price) + Estimated Total Costs (etc) are
    # now written as VALUES (the yellow inputs, = Original+COs / Revised ETC via
    # the CpRow properties); the derived columns below reference those cells.
    "original_profit",   # ORIGINAL PROFIT   = Total Contract − Estimated Costs
    "gross_profit_pct",  # GROSS PROFIT %    = Original Profit / Revised Contract
    "cost_to_complete",  # COST TO COMPLETE  = Revised ETC − Costs to Date
    "_pct_complete",     # % COMPLETE        = Costs / Revised ETC
    "_earned_revenue",   # REVENUES EARNED   = Revised Contract × (Costs / Revised ETC)
    "profit_earned",     # PROFIT EARNED     = Earned Revenue − Costs to Date
    "future_profit",     # FUTURE PROFIT     = Original Profit − Profit Earned
    "left_to_bill",      # LEFT TO BILL      = Revised Contract − Billed
    "overbillings",      # OVERBILLINGS      = MAX(Billed − Earned, 0)
    "underbillings",     # UNDERBILLINGS     = MAX(Earned − Billed, 0)
    "job_borrow",        # PURE JOB BORROW   = MAX(Cost to Complete − Left to Bill, 0)
})


def _build_formula(field_name: str, row_num: int,
                   col_letter_by_field: Dict[str, str]) -> str:
    """Build the Excel formula string for a derived-field cell in row `row_num`.

    Every reference is resolved by FIELD NAME → column letter (dynamic), so
    the formulas survive any column reordering. Blank inputs propagate to a
    blank output (never a false zero or #DIV/0), matching the Python
    None-cascades."""
    def ref(f: str) -> str:
        return col_letter_by_field[f] + str(row_num)

    F  = ref("contract_price")     # Total Contract Price (value)
    I  = ref("etc")                # Estimated Total Costs (value)
    J  = ref("billed_to_date")     # Billed (gross)
    K  = ref("costs_to_date")      # Costs to Date
    OP = ref("original_profit")    # Original Profit
    CTC = ref("cost_to_complete")  # Cost to Complete
    ER = ref("_earned_revenue")    # Revenues Earned to Date
    PE = ref("profit_earned")      # Profit Earned to Date
    LTB = ref("left_to_bill")      # Left to Bill

    if field_name == "original_profit":
        return f'=IF(OR({F}="",{I}=""),"",{F}-{I})'
    if field_name == "gross_profit_pct":
        return f'=IF(OR({OP}="",{F}="",{F}=0),"",{OP}/{F})'
    if field_name == "cost_to_complete":
        return f'=IF(OR({I}="",{K}=""),"",{I}-{K})'
    if field_name == "_pct_complete":
        return f'=IF(OR({K}="",{I}="",{I}=0),"",{K}/{I})'
    if field_name == "_earned_revenue":
        return f'=IF(OR({F}="",{K}="",{I}="",{I}=0),"",{F}*{K}/{I})'
    if field_name == "profit_earned":
        return f'=IF(OR({ER}="",{K}=""),"",{ER}-{K})'
    if field_name == "future_profit":
        return f'=IF(OR({OP}="",{PE}=""),"",{OP}-{PE})'
    if field_name == "left_to_bill":
        return f'=IF(OR({F}="",{J}=""),"",{F}-{J})'
    if field_name == "overbillings":
        # BIE — billings in excess of earned revenue (liability). 0 if underbilled.
        return f'=IF(OR({J}="",{ER}=""),"",MAX({J}-{ER},0))'
    if field_name == "underbillings":
        # CIE — earned revenue in excess of billings (asset). 0 if overbilled.
        return f'=IF(OR({ER}="",{J}=""),"",MAX({ER}-{J},0))'
    if field_name == "job_borrow":
        # Pure Job Borrow — remaining costs in excess of remaining billing
        # capacity (Wouch Maloney CPAs): costs-to-complete beyond what's left
        # to bill = cash this job must borrow from others. 0 if none.
        return f'=IF(OR({CTC}="",{LTB}=""),"",MAX({CTC}-{LTB},0))'
    raise ValueError(f"No formula defined for field {field_name!r}")


def _row_display_value(row: CpRow, field_name: str, sync_ts: str):
    """Return the value to write into the cell for a NON-formula `field_name`
    (formula fields are written via _build_formula, never here)."""
    if field_name == "_active_status":
        return "Closed" if row.is_completed else "Active"
    if field_name == "why_link":
        return "why ⇗" if row.why_link else None
    if field_name == "_folder_link":
        return "folder ⇗" if (row.folder_path or row.takeoff_path) else None
    if field_name == "_source_link":
        return "source ⇗" if row.src_link else None
    if field_name == "_last_synced":
        return sync_ts
    if field_name == "_notes_all":
        # ONE commentary cell: the owner's ACTION text, the script's
        # informational notes, then any genuine must-fix flag.
        #
        # De-duplication is per SEGMENT, not per string (2026-08-03): the
        # owner's ACTION text arrives as one string that already contains
        # ' · ' separators, so whole-string matching let the same sentence
        # come back twice. Split everything to segments first, compare on a
        # normalised key, and keep the FIRST wording seen — the message is
        # never reworded, only the redundant 'note:' label is dropped (the
        # column is already called NOTES).
        parts, seen = [], set()
        for chunk in ([getattr(row, "action_note", None)] + list(row.notes)
                      + list(row.status_flags)):
            for seg in str(chunk or "").split(" · "):
                seg = seg.strip()
                for lead in ("RED:", "note:", "NOTE:"):
                    if seg.startswith(lead):
                        seg = seg[len(lead):].strip()
                key = re.sub(r"\s+", " ", seg).lower()
                if seg and key not in seen:
                    seen.add(key)
                    parts.append(seg)
        return " · ".join(parts) or None
    if field_name == "status":
        # RED numbers must explain themselves (the user 2026-07-16): a red
        # row with no script flag showed "OK" — surface the classify reason
        # here instead (tabs without a NOTES column had nowhere else).
        if row.status_flags:
            return "; ".join(row.status_flags)
        if row.needs_review:
            red = next((n[len("RED: "):] for n in row.notes
                        if n.startswith("RED: ")), None)
            return red or (row.notes[-1] if row.notes else "needs review")
        return "OK"
    return getattr(row, field_name, None)


# ───────────────── change audit (the user 2026-07-31) ─────────────
# Every sync must state, per division, what moved: jobs added, jobs removed,
# original contract/ETC changes, revised contract/ETC changes. The owner
# audits this on every run — it is never optional and never summarised away.
AUDIT_XLSX = Path.home() / "Downloads" / "WIP Changes.xlsx"

# Note segments the SCRIPT writes. Anything else in a NOTES cell was typed by
# a human and is carried forward across the full-replace (the user
# 2026-07-31: "be sure to preserve any notes").
_SCRIPT_NOTE_RE = re.compile(
    r"^(Draw #|No draw yet|QBO |No QBO project|Duplicate line in the RP file|"
    r"No budget \(ETC\)|Contract/ETC from |Data integrity:|On the .* schedule|"
    r"% based on Original ETC|RED: )", re.IGNORECASE)


def _division(pnum: str) -> str:
    """MFD / CP / RP from the project number (RP7234-FTW → RP)."""
    p = (pnum or "").strip().upper()
    for d in ("MFD", "CP", "RP"):
        if p.startswith(d):
            return d
    return "OTHER"


def _num(v):
    """Cell → float, or None (formulas/blanks/text read back as None)."""
    return float(v) if isinstance(v, (int, float)) else None


def _snapshot_tab(ws, hdr_row: Optional[int]) -> Dict[str, dict]:
    """The tab as it stands BEFORE the rewrite: per project #, the numbers the
    audit compares against plus its NOTES text. Read straight off the sheet so
    the baseline is what the owner last saw, not what any script remembers."""
    if not hdr_row:
        return {}
    idx = {ws.cell(hdr_row, c).value: c
           for c in range(1, (ws.max_column or 0) + 1)}
    pcol = idx.get("PROJECT #")
    if not pcol:
        return {}
    prior = {}
    for r in range(hdr_row + 1, (ws.max_row or 0) + 1):
        first = ws.cell(r, 1).value
        if isinstance(first, str) and first.strip() == "TOTALS":
            break                       # summary block — not data
        pnum = ws.cell(r, pcol).value
        if not pnum or not str(pnum).strip():
            continue
        def val(label):
            c = idx.get(label)
            return _num(ws.cell(r, c).value) if c else None
        cos = val("APPROVED COs") or 0.0
        rev_k = val("TOTAL CONTRACT PRICE")
        etc = val("ESTIMATED TOTAL COSTS")
        ncol = idx.get("NOTES")
        prior[str(pnum).strip().upper()] = {
            "name": ws.cell(r, idx["PROJECT NAME"]).value if idx.get("PROJECT NAME") else None,
            "rev_contract": rev_k,
            "orig_contract": (rev_k - cos) if rev_k is not None else None,
            "rev_etc": etc,
            "orig_etc": etc,      # identical until CO costs have a source
            "notes": (str(ws.cell(r, ncol).value).strip()
                      if ncol and ws.cell(r, ncol).value else ""),
        }
    return prior


def _changed(old, new) -> bool:
    if old is None and new is None:
        return False
    if old is None or new is None:
        return True
    return abs(float(old) - float(new)) >= 0.01


def audit_changes(prior: Dict[str, dict], rows: List["CpRow"]) -> Dict[str, dict]:
    """Diff the incoming rows against the tab's previous state, split by
    division. Returns {division: {added/removed/orig/revised}}."""
    out = {}
    def bucket(d):
        return out.setdefault(d, {"added": [], "removed": [],
                                  "orig": [], "revised": []})
    seen = set()
    for row in rows:
        p = row.project_num.strip().upper()
        seen.add(p)
        b = bucket(_division(p))
        was = prior.get(p)
        if was is None:
            b["added"].append((p, row.project_name, row.contract_price, row.etc))
            continue
        for key, label, old, new in (
                ("orig", "ORIGINAL CONTRACT", was["orig_contract"], row.base_contract),
                ("orig", "ORIGINAL ETC",      was["orig_etc"],      row.base_etc),
                ("revised", "REVISED CONTRACT", was["rev_contract"], row.contract_price),
                ("revised", "REVISED ETC",      was["rev_etc"],      row.etc)):
            if _changed(old, new):
                b[key].append((p, row.project_name, label, old, new))
    for p, was in prior.items():
        if p not in seen:
            bucket(_division(p))["removed"].append(
                (p, was.get("name"), was.get("rev_contract"), was.get("rev_etc")))
    return out


def _fmt(v) -> str:
    return "—" if v is None else f"${v:,.0f}"


def print_audit(audit: Dict[str, dict], tab_name: str) -> None:
    """The always-on change report. Loud, per division, never summarised away."""
    print()
    print(_Term.color(_Term.BOLD, f"  WIP CHANGE AUDIT — {tab_name}"))
    print("  " + "=" * 66)
    if not audit:
        print("  (no previous data on this tab — this run sets the baseline)")
        return
    total = 0
    for div in ("MFD", "CP", "RP", "OTHER"):
        b = audit.get(div)
        if not b or not any(b.values()):
            continue
        n = sum(len(v) for v in b.values())
        total += n
        print(_Term.color(_Term.BOLD + _Term.CYAN, f"\n  ── {div} ──"))
        for job, name, k, e in b["added"]:
            print(f"    + NEW      {job:<14} {str(name or '')[:30]:<30} "
                  f"contract {_fmt(k)} · ETC {_fmt(e)}")
        for job, name, k, e in b["removed"]:
            print(f"    − REMOVED  {job:<14} {str(name or '')[:30]:<30} "
                  f"was contract {_fmt(k)} · ETC {_fmt(e)}")
        for tag, key in (("ORIGINAL", "orig"), ("REVISED", "revised")):
            for job, name, label, old, new in b[key]:
                delta = (new or 0) - (old or 0)
                print(f"    ~ {label:<17} {job:<14} {_fmt(old)} → {_fmt(new)} "
                      f"({'+' if delta >= 0 else '−'}{_fmt(abs(delta))[1:]})")
        if not any(b.values()):
            print("    no changes")
    if total == 0:
        print("  No changes in any division since the last sync.")
    print()


def write_audit_xlsx(audit: Dict[str, dict], tab_name: str,
                     out_path: Path = AUDIT_XLSX) -> Optional[Path]:
    """One flat, plain sheet of every change so the owner can audit later.
    Overwrites the SAME file each run (never a v2). Skipped if it's open."""
    from openpyxl import Workbook
    if out_path.with_name("~$" + out_path.name).exists():
        print(_Term.color(_Term.AMBER,
              f"  ⚠ {out_path.name} is open in Excel — change file not written"))
        return None
    wb = Workbook()
    ws = wb.active
    ws.title = "CHANGES"
    ws.cell(1, 2, f"WIP CHANGES - {tab_name}").font = TITLE_FONT
    ws.cell(2, 2, f"REPORT DATE: {dt.date.today():%b %d, %Y}".upper()).font = TITLE_FONT
    hdr = ["DIVISION", "CHANGE", "JOB #", "JOB NAME", "FIELD", "OLD", "NEW", "CHANGE $"]
    for c, h in enumerate(hdr, start=1):
        cell = ws.cell(4, c, h)
        cell.font = HDR_FONT
        cell.fill = HDR_FILL
        cell.border = CELL_BORDER
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
    widths = (10, 12, 16, 34, 20, 16, 16, 16)
    for c, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(c)].width = w
    r = 5
    for div in ("MFD", "CP", "RP", "OTHER"):
        b = audit.get(div)
        if not b:
            continue
        for job, name, k, e in b["added"]:
            for c, v in enumerate([div, "NEW JOB", job, name, "CONTRACT",
                                   None, k, k], start=1):
                ws.cell(r, c, v)
            r += 1
            for c, v in enumerate([div, "NEW JOB", job, name, "ETC",
                                   None, e, e], start=1):
                ws.cell(r, c, v)
            r += 1
        for job, name, k, e in b["removed"]:
            for c, v in enumerate([div, "REMOVED", job, name, "CONTRACT / ETC", k, None, None], start=1):
                ws.cell(r, c, v)
            r += 1
        for key, tag in (("orig", "ORIGINAL"), ("revised", "REVISED")):
            for job, name, label, old, new in b[key]:
                delta = (new or 0) - (old or 0)
                for c, v in enumerate([div, tag + " CHANGE", job, name, label,
                                       old, new, delta], start=1):
                    ws.cell(r, c, v)
                r += 1
    for rr in range(5, r):
        for c in range(1, len(hdr) + 1):
            cell = ws.cell(rr, c)
            cell.font = DATA_FONT
            cell.border = CELL_BORDER
            if c >= 6:
                cell.number_format = CURRENCY_FMT
    if r == 5:
        ws.cell(5, 1, "No changes since the last sync.").font = DATA_FONT
    ws.freeze_panes = "A5"
    wb.save(out_path)
    return out_path


def _master_title_prefix(wb) -> str:
    """The company/report prefix from the real 'WIP Master' title cell (B1),
    e.g. 'PROFICIENT CONCRETE, LLC' out of '<company> - WIP REPORT'. Read at
    RUNTIME so the test tabs inherit the owner's own wording (and so no
    company name is hard-coded in this repo). '' when unreadable."""
    try:
        v = wb[MASTER_TITLE_SHEET].cell(row=1, column=2).value
    except (KeyError, AttributeError):
        return ""
    return str(v).split(" - ")[0].strip() if v else ""


def _find_header_row(ws) -> Optional[int]:
    """Locate the table header row (the one holding 'PROJECT #') in the first
    few rows — row 1 on plain tabs, below the title banner (and any legend
    block) on branded ones."""
    for r in range(1, min(ws.max_row or 0, 15) + 1):
        for c in range(1, (ws.max_column or 0) + 1):
            if ws.cell(r, c).value == "PROJECT #":
                return r
    return None


def _write_summary(ws, cols_, col_letter_by_field, data_start: int,
                   last_row: int, start_row: int) -> None:
    """WIP-master-style TOTALS row + FUTURE WIP CASH FLOW block (the user
    2026-07-16), written BELOW the table/appendix.

    TOTALS uses SUBTOTAL(109, …) over the table's data rows — it counts only
    VISIBLE rows, so filtering the table (e.g. hiding FTW BACKLOG or Closed)
    re-totals live. The cash-flow block derives everything from the TOTALS
    cells:  rev left = contract − earned · GP left = rev left − CTC ·
    net under/(over) = under − over · cash flow = GP left + net under."""
    L = col_letter_by_field.get
    tot = start_row

    # ── TOTALS row ──
    _sum_fields = [f for f in (
        "contract_price", "etc", "billed_to_date", "costs_to_date",
        "cost_to_complete", "original_profit", "_earned_revenue",
        "profit_earned", "overbillings", "underbillings", "left_to_bill",
        "future_profit", "job_borrow", "co_revenue", "retainage_held")
        if L(f)]
    for c in range(1, len(cols_) + 1):
        cell = ws.cell(row=tot, column=c)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.border = CELL_BORDER
    ws.cell(row=tot, column=1, value="TOTALS")
    for f in _sum_fields:
        cell = ws.cell(row=tot, column=column_index_from_string(L(f)))
        cell.value = f"=SUBTOTAL(109,{L(f)}{data_start}:{L(f)}{last_row})"
        cell.number_format = CURRENCY_FMT
    if L("_pct_complete") and L("costs_to_date") and L("etc"):
        c = ws.cell(row=tot, column=column_index_from_string(L("_pct_complete")))
        c.value = (f'=IF({L("etc")}{tot}=0,"",'
                   f'{L("costs_to_date")}{tot}/{L("etc")}{tot})')
        c.number_format = PCT_FMT
    if L("gross_profit_pct") and L("original_profit") and L("contract_price"):
        c = ws.cell(row=tot, column=column_index_from_string(L("gross_profit_pct")))
        c.value = (f'=IF({L("contract_price")}{tot}=0,"",'
                   f'{L("original_profit")}{tot}/{L("contract_price")}{tot})')
        c.number_format = PCT_FMT

    # ── FUTURE WIP CASH FLOW block ──
    vcol = column_index_from_string(L("contract_price") or "E")     # amounts stack in one column
    vL = get_column_letter(vcol)
    r0 = tot + 2
    lines = [
        ("FUTURE WIP CASH FLOW", None, True),
        ("TOTAL CONTRACT PRICE", f"={L('contract_price')}{tot}", False),
        ("REVENUE EARNED TO DATE", f"={L('_earned_revenue')}{tot}", False),
        ("REVENUE LEFT TO EARN", f"={vL}{r0 + 1}-{vL}{r0 + 2}", False),
        ("COST TO COMPLETE", f"={L('cost_to_complete')}{tot}", False),
        ("G.P. LEFT TO EARN", f"={vL}{r0 + 3}-{vL}{r0 + 4}", False),
        ("+/- NET UNDER/(OVERBILLINGS)",
         f"={L('underbillings')}{tot}-{L('overbillings')}{tot}", False),
        ("FUTURE WIP CASH FLOW", f"={vL}{r0 + 5}+{vL}{r0 + 6}", True),
    ]
    _bold = Font(name=MASTER_FONT_NAME, size=MASTER_FONT_SIZE, bold=True)
    for k, (label, formula, bold) in enumerate(lines):
        r = r0 + k
        lc = ws.cell(row=r, column=1, value=label)
        lc.font = _bold if bold else DATA_FONT
        lc.border = CELL_BORDER
        if formula:
            vc = ws.cell(row=r, column=vcol, value=formula)
            vc.number_format = CURRENCY_FMT
            vc.border = CELL_BORDER
            vc.font = _bold if bold else DATA_FONT
    # G.P. LEFT TO EARN margin % (of revenue left to earn) beside the amount
    pc = ws.cell(row=r0 + 5, column=vcol + 1)
    pc.value = f'=IF({vL}{r0 + 3}=0,"",{vL}{r0 + 5}/{vL}{r0 + 3})'
    pc.number_format = PCT_FMT
    pc.font = DATA_FONT
    pc.border = CELL_BORDER


def write_test_cp(rows: List[CpRow], wip_path: Path, dry_run: bool = False,
                  tab_name: str = TEST_TAB,
                  appendix: Optional[Tuple[str, List[CpRow]]] = None,
                  cols: Optional[List] = None,
                  default_filter_active: bool = False,
                  title: Optional[str] = None,
                  summary: bool = False,
                  qbo_links_only: bool = True,
                  legend: Optional[List] = None,
                  audit: bool = False,
                  audit_xlsx: Optional[Path] = None) -> bool:
    """Write rows to the given WIP tab (default 'Test - CP'; RP passes
    'Test - RP'). Same structure/formatting for every division. Guarded by
    wip_excel_guard. Returns True if written, False if skipped (dry-run, or the
    file is open in Excel).

    `appendix` = (section title, rows): written BELOW the main table under a
    gray band — RP uses it for the FTW backlog (bid with the slab, not poured
    yet; the user 2026-07-14: separated at the bottom, they read as expected
    wins rather than in-progress jobs).

    `title` (the user 2026-07-16): report banner across the top ("WIP REPORT
    as of …") with the two banner rows reserved as logo space — embedded
    images (the logo) survive every sync (openpyxl round-trips them; the
    rewrite only touches cells). `summary` adds the WIP-master-style TOTALS
    row under the table (live SUBTOTALs — they follow the table filter) plus
    the FUTURE WIP CASH FLOW block derived from it.

    `qbo_links_only` (the user 2026-07-29, now the DEFAULT — file hyperlinks
    weigh the workbook down): suppress every file:// hyperlink (Synology
    folders, takeoffs, draws, source workbooks, WHY) — the only links on the
    report are the QBO deep links on the Billed/Costs cells. Pass False to
    restore the full click-to-verify link set.

    `legend` (the user 2026-07-31): rows of (text, font_rgb_or_None, bold)
    rendered under the banner, one plain single-font cell per row (rich text
    is banned — it corrupts the workbook). Rows carrying `cell_marks`
    ({field: rgb}) get the owner's colour re-applied to those cells — his
    verified/changed/verify-me marks survive the sync."""
    assert_write_allowed(tab_name)  # tripwire before we even open the workbook
    cols_ = cols or COLS

    if dry_run:
        _print_rows_table(rows, wip_path, tab_name)
        return False

    if not wip_path.exists():
        raise FileNotFoundError(
            f"WIP Excel not found at {wip_path}. "
            f"Check WIP_EXCEL_PATH env or verify OneDrive is synced."
        )

    # Never clobber a workbook that's open in Excel. Excel drops a
    # ~$<name> owner-lock file next to an open workbook (project-pnl
    # safe_save pattern, the user 2026-06-23). If present, skip with a clear
    # message rather than writing underneath the open file — a
    # last-writer-wins save would silently lose the sync or the user's edits.
    lock = wip_path.with_name("~$" + wip_path.name)
    if lock.exists():
        print(_Term.color(_Term.AMBER,
              f"  ⚠ {wip_path.name} looks OPEN in Excel — skipped the write to "
              f"avoid overwriting it. Close the file and re-run."))
        return False

    wb = open_wip_workbook_for_write(wip_path)
    try:
        if tab_name not in wb.sheetnames:
            wb.create_sheet(tab_name)
        ws = wb[tab_name]
        assert_write_allowed(ws.title)  # belt + suspenders

        # A SHEET-level AutoFilter cannot coexist with the Table's own filter:
        # Excel calls the workbook damaged ("We found a problem with some
        # content…") and repairs it on open. A previous tool's filter survives
        # the cell wipe — 'Test - RP' still carried A2:L69 from the old
        # 12-column layout (2026-07-31). Clearing the ref also drops the
        # hidden _xlnm._FilterDatabase defined name openpyxl derives from it.
        ws.auto_filter.ref = None

        # PRESERVE USER CELL COMMENTS across the full-replace (the user
        # 2026-07-16: review notes typed on cells — e.g. "add the missing
        # $5k" on a contract price — must survive every sync). Harvest them
        # keyed by (PROJECT #, header label) before the wipe; re-attach to
        # the same project/column after the rewrite. A comment whose project
        # left the tab is PRINTED, never silently dropped.
        saved_comments = {}
        prior_hdr = _find_header_row(ws)
        # Baseline for the change audit + the owner's typed NOTES, both read
        # off the sheet BEFORE it is wiped (the user 2026-07-31).
        prior_state = _snapshot_tab(ws, prior_hdr)
        for _p, _s in prior_state.items():
            keep = [seg.strip() for seg in _s["notes"].split(" · ")
                    if seg.strip() and not _SCRIPT_NOTE_RE.match(seg.strip())]
            _s["manual_notes"] = keep
        hdr_labels = ({c: ws.cell(prior_hdr, c).value
                       for c in range(1, (ws.max_column or 0) + 1)}
                      if prior_hdr else {})
        pnum_col = next((c for c, v in hdr_labels.items() if v == "PROJECT #"), None)
        if pnum_col:
            for r in range(prior_hdr + 1, (ws.max_row or 0) + 1):
                pnum = ws.cell(r, pnum_col).value
                if not pnum:
                    continue
                for c, label in hdr_labels.items():
                    cm = ws.cell(r, c).comment
                    if cm and label:
                        saved_comments[(str(pnum).strip(), label)] = (cm.text, cm.author)

        # Full replace: clear existing rows AND explicitly wipe every cell
        # attribute (value, hyperlink, number_format, fill, font) up to the
        # prior extent so leftovers can't leak into new rows.
        # openpyxl's delete_rows sometimes leaves formatting or hyperlinks
        # behind on cells that were populated before — belt-and-suspenders
        # clear here to guarantee a clean slate.
        # Merged cells (a prior run's title banner) must be unmerged before
        # the clear — writing to a MergedCell raises in openpyxl.
        for rng in list(ws.merged_cells.ranges):
            ws.unmerge_cells(str(rng))
        prior_max_row = ws.max_row or 0
        prior_max_col = ws.max_column or 0
        off = (2 if title else 0) + (len(legend) if legend else 0)
        # banner + legend rows above the header
        _sects = ([appendix] if isinstance(appendix, tuple) else (appendix or []))
        n_total = (off + len(rows) + sum(len(a[1]) + 3 for a in _sects if a[1])
                   + (14 if summary else 0))
        for r in range(1, max(prior_max_row, n_total + 1) + 1):
            for c in range(1, max(prior_max_col, len(cols_)) + 1):
                cell = ws.cell(row=r, column=c)
                cell.value = None
                cell.hyperlink = None
                cell.number_format = "General"
                cell.comment = None   # harvested above; stale ones must not linger
        if prior_max_row > 0:
            ws.delete_rows(1, prior_max_row)
        # Reset stale hidden flags — rows shift between runs, and a hidden
        # flag left on the wrong row would silently hide live data.
        # Row HEIGHTS reset too (2026-08-03): the header/banner moves between
        # layouts, and a leftover 30pt height on what is now a blank or legend
        # row shows up as a random tall gap.
        for _rd in ws.row_dimensions.values():
            _rd.hidden = False
            _rd.height = None

        # Title banner (the user 2026-07-16): "WIP REPORT as of …" across the
        # table width, two tall rows that double as the logo's parking space
        # (the logo image floats over the cells — the rewrite never touches
        # it, so it stays put run after run).
        hdr_row = 1 + off
        if title:
            # TITLE BLOCK — copied from the real 'WIP Master' sheet and
            # BINDING (the user 2026-07-31: "keep the same format as the
            # original WIP Master sheet"). B1 = company + report name,
            # B2 = REPORT DATE, both Tahoma 8 bold and LEFT-aligned, a medium
            # rule above row 1 and below row 2. No merge-and-center, no
            # oversized font, no custom row heights. Do not redesign this.
            prefix = _master_title_prefix(wb)
            ws.cell(row=1, column=2,
                    value=f"{prefix} - {title}" if prefix else title
                    ).font = TITLE_FONT
            ws.cell(row=2, column=2,
                    value=f"REPORT DATE: {dt.date.today():%b %d, %Y}".upper()
                    ).font = TITLE_FONT
            for c in range(1, len(cols_) + 1):
                ws.cell(row=1, column=c).border = Border(top=_MEDIUM)
                ws.cell(row=2, column=c).border = Border(bottom=_MEDIUM)
        if legend:
            lr0 = 3 if title else 1
            for k, (txt, rgb, bold) in enumerate(legend):
                lc = ws.cell(row=lr0 + k, column=1, value=txt)
                lc.font = Font(name=MASTER_FONT_NAME, size=MASTER_FONT_SIZE,
                               color=(rgb or "000000"), bold=bold)

        # Header row — gray, bold, centered + wrapped, bordered.
        for c, (label, width, _key) in enumerate(cols_, start=1):
            cell = ws.cell(row=hdr_row, column=c, value=label)
            cell.fill = HDR_FILL
            cell.font = HDR_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center",
                                       wrap_text=True)
            cell.border = CELL_BORDER
            ws.column_dimensions[get_column_letter(c)].width = width
        ws.row_dimensions[hdr_row].height = 30
        ws.freeze_panes = f"A{hdr_row + 1}"

        sync_ts = dt.datetime.now().strftime("%Y-%m-%d %H:%M")

        # Column indices for hyperlink attachment (1-based like openpyxl)
        # + column LETTERS for building cross-cell formulas.
        col_idx = {field: i + 1 for i, (_, _, field) in enumerate(cols_)}
        col_letter_by_field = {
            field: get_column_letter(i + 1)
            for i, (_, _, field) in enumerate(cols_)
        }

        def _emit(i: int, row: CpRow) -> None:
            # Re-attach any NOTES text a human typed on this job's row last
            # time (script segments regenerate themselves; only human text is
            # carried) — the user 2026-07-31: "be sure to preserve any notes".
            # SKIPPED when the row's notes come from a source file the owner
            # edits (RP): that file is authoritative, and carrying its text
            # forward would resurrect a note he had just deleted from it.
            if not getattr(row, "notes_from_source", False):
                for _seg in (prior_state.get(row.project_num.strip().upper(), {})
                             .get("manual_notes") or []):
                    if _seg not in row.notes:
                        row.notes.append(_seg)
            # Invariant guard: if CO Cost is populated but CO Revenue is
            # not, refuse to write the CO Cost. That's either a bug or a
            # corrupted state, and quietly writing a made-up cost would
            # be worse than surfacing the issue.
            if row.co_revenue is None and row.co_cost_estimate is not None:
                log.warning(
                    "%s: CO Cost populated ($%s) with no CO Revenue — "
                    "refusing to write CO Cost. Investigate parse logic.",
                    row.project_num, row.co_cost_estimate
                )
                row.status_flags.append("Data integrity: CO Cost without CO Rev — dropped")

            for c, (_label, _width, field_name) in enumerate(cols_, start=1):
                if field_name in FORMULA_FIELDS:
                    # Derived cell — write an Excel formula referencing
                    # the input cells in the same row. Excel evaluates on
                    # open/edit so any input change auto-recalculates.
                    val = _build_formula(field_name, i, col_letter_by_field)
                else:
                    val = _row_display_value(row, field_name, sync_ts)
                cell = ws.cell(row=i, column=c, value=val)
                cell.border = CELL_BORDER
                cell.font = DATA_FONT          # master-sheet Tahoma 8 baseline
                if field_name in _MONEY_FIELDS:
                    cell.number_format = CURRENCY_FMT
                elif field_name in _PCT_FIELDS:
                    cell.number_format = PCT_FMT
                # Yellow = sourced input (raw from takeoff / QBO); white = calc.
                if field_name in _SOURCE_FIELDS:
                    cell.fill = INPUT_FILL
                elif field_name in ("status", "_notes_all") and (
                        row.status_flags or row.needs_review):
                    cell.fill = FLAG_FILL
                    cell.font = FLAG_FONT

            # PROJECT FOLDER → project (Awarded Project) folder; DATA SOURCE →
            # where the numbers came from (GL row / master tab / takeoff). The
            # user 2026-07-16: links moved OFF the # and name cells into their
            # own columns so the identifiers can be selected/copied without
            # Excel navigating away. (file:// links still get rewritten by
            # Windows/OneDrive Excel — kept for the Mac, the trace point.)
            # Folder target: explicit folder_path (RP) else takeoff's parent (CP).
            if not qbo_links_only:
                link_folder = (row.folder_path if row.folder_path is not None
                               else (row.takeoff_path.parent if row.takeoff_path else None))
                if "_folder_link" in col_idx:
                    _apply_hyperlink(ws.cell(row=i, column=col_idx["_folder_link"]),
                                     link_folder)
                if "_source_link" in col_idx and row.src_link:
                    _apply_hyperlink(ws.cell(row=i, column=col_idx["_source_link"]),
                                     Path(row.src_link), row.src_fragment or "")

                # Number-cell links (the user 2026-07-13: every number click-to-verify):
                # Contract/COs → the draw workbook (takeoff pre-draw); ETC → takeoff;
                # Billed/Retainage → QBO customer page (all invoices on one screen);
                # Costs → QBO project-filtered P&L report.
                for _f, _tgt in (("contract_price", row.draw_path or row.takeoff_path),
                                 ("co_revenue", row.draw_path),
                                 ("etc", row.takeoff_path)):
                    if _f in col_idx:
                        _apply_hyperlink(ws.cell(row=i, column=col_idx[_f]), _tgt)
                if "why_link" in col_idx and row.why_link:
                    _apply_hyperlink(ws.cell(row=i, column=col_idx["why_link"]),
                                     Path(row.why_link),
                                     row.why_fragment or "")
            if row.qbo_customer_id and QBO_REALM:
                _cu = qbo_api.customer_url(row.qbo_customer_id, QBO_REALM)
                _pu = qbo_api.project_pl_url(row.qbo_customer_id, QBO_REALM)
                for _f, _u in (("billed_to_date", _cu), ("retainage_held", _cu),
                               ("costs_to_date", _pu)):
                    if _f not in col_idx:
                        continue
                    _c = ws.cell(row=i, column=col_idx[_f])
                    if _u and _c.value not in (None, ""):
                        _c.hyperlink = _u
                        _c.font = LINK_FONT

            # Review pass LAST so it wins over link styling: numbers that don't
            # look right (row.needs_review) render RED (the user 2026-07-13);
            # underline is kept where the cell carries a link.
            if row.needs_review:
                for _f in (_MONEY_FIELDS | _PCT_FIELDS):
                    if _f in col_idx:
                        _c = ws.cell(row=i, column=col_idx[_f])
                        _c.font = Font(
                            name=MASTER_FONT_NAME, size=MASTER_FONT_SIZE,
                            color="C00000",
                            underline=("single" if _c.hyperlink else None))

            # The owner's colour marks are applied LAST — they outrank link
            # and review styling (the user 2026-07-31: "keep all the notes
            # and colors since they mean something").
            for _f, _rgb in (getattr(row, "cell_marks", None) or {}).items():
                if _f in col_idx:
                    _c = ws.cell(row=i, column=col_idx[_f])
                    _c.font = Font(
                        name=MASTER_FONT_NAME, size=MASTER_FONT_SIZE,
                        color=_rgb, bold=True,
                        underline=("single" if _c.hyperlink else None))

        data_start = hdr_row + 1
        written_rows = {}                   # PROJECT # → sheet row (for comments)
        for i, row in enumerate(rows, start=data_start):
            _emit(i, row)
            written_rows[row.project_num] = i

        # Wrap the range in an Excel Table — gives filter/sort dropdowns and a
        # structured, clean look. Explicit gray header + borders above override
        # the table style, so it stays clean (no row stripes).
        last_row = len(rows) + hdr_row
        last_col = get_column_letter(len(cols_))
        for tname in list(ws.tables):
            del ws.tables[tname]          # drop any prior run's table first
        tbl_name = re.sub(r"[^A-Za-z0-9]", "", tab_name) or "WIP"  # unique per tab
        table = Table(displayName=tbl_name,
                      ref=f"A{hdr_row}:{last_col}{last_row}")
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleLight1", showFirstColumn=False, showLastColumn=False,
            showRowStripes=False, showColumnStripes=False)
        if default_filter_active and "_active_status" in col_idx:
            # Default view = Active only (the user 2026-07-14): apply the
            # STATUS filter on the table AND hide the Closed rows — Excel
            # shows them again with one filter click.
            from openpyxl.worksheet.filters import (AutoFilter, FilterColumn,
                                                     Filters)
            fc = FilterColumn(colId=col_idx["_active_status"] - 1,
                              filters=Filters(filter=["Active"]))
            table.autoFilter = AutoFilter(ref=f"A{hdr_row}:{last_col}{last_row}",
                                          filterColumn=[fc])
            for i, row in enumerate(rows, start=data_start):
                if row.is_completed:
                    ws.row_dimensions[i].hidden = True
        ws.add_table(table)

        # Appendix section BELOW the table (outside it, so its rows don't
        # pollute the table's filters): gray band title, then the same row
        # rendering as the main block.
        sections = ([appendix] if isinstance(appendix, tuple) else (appendix or []))
        next_row = last_row + 2             # one blank spacer row under the table
        for sect_title, ap_rows in sections:
            if not ap_rows:
                continue
            band = next_row
            for c in range(1, len(cols_) + 1):
                bc = ws.cell(row=band, column=c)
                bc.fill = HDR_FILL
                bc.border = CELL_BORDER
            t = ws.cell(row=band, column=1, value=sect_title)
            t.font = HDR_FONT
            ws.row_dimensions[band].height = 22
            for k, row in enumerate(ap_rows, start=band + 1):
                _emit(k, row)
                written_rows[row.project_num] = k
            next_row = band + len(ap_rows) + 2

        if summary:
            _write_summary(ws, cols_, col_letter_by_field, data_start,
                           last_row, next_row)

        # Re-attach the harvested user comments to their project/column.
        col_by_label = {label: c for c, (label, _w, _f) in enumerate(cols_, start=1)}
        for (pnum, label), (text, author) in saved_comments.items():
            r, c = written_rows.get(pnum), col_by_label.get(label)
            if r and c:
                ws.cell(row=r, column=c).comment = Comment(text, author or "")
            else:
                print(_Term.color(_Term.AMBER,
                      f"  ⚠ comment on {pnum} / {label} has no row this run "
                      f"(line left the tab) — text was: {text!r}"))

        # Atomic write — save to a temp file then os.replace() so a crash
        # or interruption can't leave a half-written WIP (safe_save pattern).
        tmp = wip_path.with_name(wip_path.name + ".tmp")
        wb.save(str(tmp))
        try:
            os.replace(str(tmp), str(wip_path))
        except OSError as e:
            # Keep the good data in the .tmp rather than losing it.
            log.error("Could not replace %s (%s); new data saved to %s",
                      wip_path.name, e, tmp.name)
            raise
        log.info("Wrote %d rows to %s in %s", len(rows), tab_name, wip_path)
        if audit:
            _sections = ([appendix] if isinstance(appendix, tuple)
                         else (appendix or []))
            _all = list(rows) + [r for _t, ap in _sections for r in (ap or [])]
            _audit = audit_changes(prior_state, _all)
            print_audit(_audit, tab_name)
            if audit_xlsx is not None:
                _p = write_audit_xlsx(_audit, tab_name, audit_xlsx)
                if _p:
                    print(_Term.color(_Term.DIM, f"  change file → {_p}"))
    finally:
        wb.close()
    _qc_check(wip_path, tab_name, expected_rows=len(rows) + sum(
        len(a[1]) for a in ([appendix] if isinstance(appendix, tuple)
                            else (appendix or [])) if a[1]),
              active_only=default_filter_active)
    return True


def _qc_check(wip_path: Path, tab_name: str, expected_rows: int,
              active_only: bool) -> None:
    """Visual QC after EVERY write (the user 2026-07-15): re-open the saved
    file and verify what the reader believes matches what Excel will show.
    Never raises — prints ✓/⚠ lines so a bad write is loud, not silent."""
    try:
        wb = load_workbook(wip_path)
        ws = wb[tab_name]
        hdr = _find_header_row(ws) or 1     # header sits below any title banner
        hix = {ws.cell(hdr, c).value: c for c in range(1, ws.max_column + 1)}
        pcol = hix.get("PROJECT #")
        scol = hix.get("STATUS")
        lcols = [c for lbl, c in hix.items()
                 if lbl in ("PROJECT FOLDER", "DATA SOURCE")]
        n = vis_closed = links = 0
        last_data = hdr
        for r in range(hdr + 1, ws.max_row + 1):
            # The summary block (TOTALS + FUTURE WIP CASH FLOW) is written
            # below the data in column 1 — on a tab whose first column IS
            # 'PROJECT #' its labels would otherwise be miscounted as data
            # rows. Everything from TOTALS down is the summary; stop there.
            first = ws.cell(r, 1).value
            if isinstance(first, str) and first.strip() == "TOTALS":
                break
            j = ws.cell(r, pcol).value if pcol else None
            if not j:
                continue
            n += 1
            last_data = r
            if any(ws.cell(r, c).hyperlink for c in lcols):
                links += 1
            if (scol and ws.cell(r, scol).value == "Closed"
                    and not ws.row_dimensions[r].hidden):
                vis_closed += 1
        tbl_ok = True
        if ws.tables:
            ref = list(ws.tables.values())[0].ref            # e.g. A1:U130
            tbl_end = int(re.findall(r"(\d+)$", ref)[0])
            tbl_ok = tbl_end >= last_data
        # Excel "repair on open" tripwire (2026-07-31): a sheet-level
        # AutoFilter alongside the Table's own filter makes Excel declare the
        # workbook damaged. Loud here beats a repair dialog on the user's desk.
        sheet_filter = ws.auto_filter.ref if ws.tables else None
        wb.close()
        probs = []
        if sheet_filter:
            probs.append(f"sheet AutoFilter {sheet_filter} coexists with the "
                         f"table filter — Excel will demand a repair")
        if n != expected_rows:
            probs.append(f"rows {n} ≠ expected {expected_rows}")
        if active_only and vis_closed:
            probs.append(f"{vis_closed} Closed row(s) VISIBLE in an active view")
        if not tbl_ok:
            probs.append("table does not span all data rows")
        if lcols and links == 0 and n:
            # Only meaningful when the layout HAS link columns — the master
            # tab dropped them (qbo_links_only, the user 2026-07-29).
            probs.append("no source links found")
        if probs:
            print(_Term.color(_Term.AMBER, "  ⚠ QC: " + " · ".join(probs)))
        else:
            print(_Term.color(_Term.GREEN,
                  f"  ✓ QC: {n} rows · closed hidden · table spans all · links ok"))
    except Exception as e:                                    # QC must never kill a run
        print(_Term.color(_Term.AMBER, f"  ⚠ QC check failed to run: {e}"))


# ─────────────────────── pretty run report ────────────────────────
def _shorten(path: Path, max_len: int = 70) -> str:
    """Shorten long paths by inserting an ellipsis in the middle."""
    s = str(path)
    if len(s) <= max_len:
        return s
    head = s[: max_len // 2 - 2]
    tail = s[-(max_len // 2 - 2):]
    return f"{head}…{tail}"


def _wip_metrics(r: CpRow) -> Dict[str, Optional[float]]:
    """Compute the derived WIP numbers in Python for TERMINAL DISPLAY only —
    mirrors the Excel formulas exactly (same blank-propagation), so the
    dry-run preview matches what the workbook will show after recalc."""
    F = r.contract_price          # Revised Contract
    I = r.etc                     # Revised ETC
    K = r.costs_to_date
    J = r.billed_to_date          # Billed (gross)

    def sub(a, b):
        return (a - b) if (a is not None and b is not None) else None

    orig_profit = sub(F, I)
    gp_pct = (orig_profit / F) if (orig_profit is not None and F not in (None, 0)) else None
    ctc = sub(I, K)
    pct = (K / I) if (K is not None and I not in (None, 0)) else None
    earned = (F * K / I) if (F is not None and K is not None and I not in (None, 0)) else None
    profit_earned = sub(earned, K)
    future_profit = sub(orig_profit, profit_earned)
    left_to_bill = sub(F, J)
    overbill = max(J - earned, 0.0) if (J is not None and earned is not None) else None
    underbill = max(earned - J, 0.0) if (earned is not None and J is not None) else None
    job_borrow = max(ctc - left_to_bill, 0.0) if (ctc is not None and left_to_bill is not None) else None
    return {
        "orig_profit": orig_profit, "gp_pct": gp_pct, "ctc": ctc, "pct": pct,
        "earned": earned, "profit_earned": profit_earned,
        "future_profit": future_profit, "left_to_bill": left_to_bill,
        "overbill": overbill, "underbill": underbill, "job_borrow": job_borrow,
    }


def _print_rows_table(rows: List[CpRow], wip_path: Path, tab_name: str = TEST_TAB) -> None:
    """Dry-run preview. Vertical per-project detail for small runs (the
    single-project verify case), compact table for bulk runs. Values match
    the Excel formulas via _wip_metrics."""
    _section(f"DRY RUN — would write {len(rows)} row(s) to {tab_name!r}")
    print(f"  target: {_Term.color(_Term.DIM, _shorten(wip_path))}")

    clean = sum(1 for r in rows if not r.status_flags)
    flagged = len(rows) - clean

    if len(rows) <= 3:
        # Full vertical WIP card per project — every column, easy to verify.
        for r in rows:
            m = _wip_metrics(r)
            status = "Closed" if r.is_completed else "Active"
            print()
            print(_Term.color(_Term.BOLD + _Term.CYAN,
                  f"  ▌ {r.project_num}  {r.project_name}") +
                  _Term.color(_Term.DIM, f"   [{status}]"))
            rows_out = [
                ("Original Contract",   _fmt_money(r.base_contract)),
                ("Approved COs",        _fmt_money(r.co_revenue)),
                ("Revised Contract",    _fmt_money(r.contract_price)),
                ("Original ETC",        _fmt_money(r.base_etc)),
                ("CO Cost",             _fmt_money(r.co_cost_estimate)),
                ("Revised ETC",         _fmt_money(r.etc)),
                ("Original Profit",     _fmt_money(m["orig_profit"])),
                ("Gross Profit %",      _fmt_pct(m["gp_pct"])),
                ("Costs to Date",       _fmt_money(r.costs_to_date)),
                ("Cost to Complete",    _fmt_money(m["ctc"])),
                ("% Complete",          _fmt_pct(m["pct"])),
                ("Revenues Earned",     _fmt_money(m["earned"])),
                ("Profit Earned",       _fmt_money(m["profit_earned"])),
                ("Future Profit",       _fmt_money(m["future_profit"])),
                ("Billed (gross)",      _fmt_money(r.billed_to_date)),
                ("Retainage Held",      _fmt_money(r.retainage_held)),
                ("Left to Bill",        _fmt_money(m["left_to_bill"])),
                ("Overbillings",        _fmt_money(m["overbill"])),
                ("Underbillings",       _fmt_money(m["underbill"])),
                ("Pure Job Borrow",     _fmt_money(m["job_borrow"])),
            ]
            for label, val in rows_out:
                print(f"    {label:<20} {_lpad(_dim_if_dash(val), 16)}")
            if r.notes:
                print(_Term.color(_Term.DIM, "    · " + "; ".join(r.notes)))
            if r.status_flags:
                print(_Term.color(_Term.AMBER, "    ⚑ " + "; ".join(r.status_flags)))
            else:
                print(_Term.color(_Term.GREEN, "    ✓ no flags"))
    else:
        # Compact table — the headline columns for scanning many jobs.
        # NOTES (informational: Draw #, no-draw) and FLAGS (true script must-fix
        # issues) are kept in separate trailing columns.
        W_P, W_N, W_M, W_PC, W_NOTE = 8, 24, 13, 6, 44
        SEP = _Term.color(_Term.DIM, " │ ")
        hdr = [
            _rpad("PROJECT", W_P), _rpad("NAME", W_N),
            _lpad("CONTRACT", W_M), _lpad("CO $", W_M), _lpad("ETC", W_M),
            _lpad("COSTS", W_M), _lpad("%", W_PC), _lpad("BILLED", W_M),
            _lpad("RETAIN", W_M), _lpad("OVER", W_M), _lpad("UNDER", W_M),
            _lpad("BORROW", W_M), _rpad("NOTES", W_NOTE), "FLAGS",
        ]
        print()
        print(_Term.color(_Term.BOLD, "  " + SEP.join(hdr)))
        print(_Term.color(_Term.DIM, "  " + "─" * 196))
        for r in rows:
            m = _wip_metrics(r)
            flag = (_Term.color(_Term.AMBER, "⚑ " + "; ".join(r.status_flags))
                    if r.status_flags else _Term.color(_Term.GREEN, "✓"))
            note_txt = "; ".join(r.notes)
            note = (_Term.color(_Term.DIM, _rpad(note_txt[:W_NOTE], W_NOTE))
                    if note_txt else _rpad("", W_NOTE))
            name = _rpad(r.project_name[:W_N], W_N)
            if r.is_completed:
                name = _Term.color(_Term.GREEN, name)
            cells = [
                _rpad(r.project_num, W_P), name,
                _lpad(_dim_if_dash(_fmt_money(r.contract_price)), W_M),
                _lpad(_dim_if_dash(_fmt_money(r.co_revenue)), W_M),
                _lpad(_dim_if_dash(_fmt_money(r.etc)), W_M),
                _lpad(_dim_if_dash(_fmt_money(r.costs_to_date)), W_M),
                _lpad(_dim_if_dash(_fmt_pct(m["pct"])), W_PC),
                _lpad(_dim_if_dash(_fmt_money(r.billed_to_date)), W_M),
                _lpad(_dim_if_dash(_fmt_money(r.retainage_held)), W_M),
                _lpad(_dim_if_dash(_fmt_money(m["overbill"])), W_M),
                _lpad(_dim_if_dash(_fmt_money(m["underbill"])), W_M),
                _lpad(_dim_if_dash(_fmt_money(m["job_borrow"])), W_M),
                note, flag,
            ]
            print("  " + SEP.join(cells))

    print()
    print(f"  {_Term.color(_Term.BOLD, 'Summary')}:  "
          f"{len(rows)} total  ·  "
          f"{_Term.color(_Term.GREEN, str(clean) + ' clean')}  ·  "
          f"{_Term.color(_Term.AMBER, str(flagged) + ' flagged')}")

    # Audit trail — file paths the reader touched. Shown when small enough
    # to be useful (≤ 5 projects), so single-project runs always get it.
    if len(rows) <= 5:
        print()
        print(_Term.color(_Term.DIM, "  Audit trail:"))
        for r in rows:
            if r.draw_path:
                print(_Term.color(_Term.DIM,
                      f"    {r.project_num}  draw #{r.draw_num}:    {r.draw_path}"))
            if len(r.included_takeoffs) > 1:
                print(_Term.color(_Term.DIM, f"    {r.project_num}  takeoffs (summed, {len(r.included_takeoffs)}):"))
                for tk in r.included_takeoffs:
                    print(_Term.color(_Term.DIM, f"    {r.project_num}    · {tk.name}"))
            elif r.takeoff_path:
                print(_Term.color(_Term.DIM, f"    {r.project_num}  takeoff:      {r.takeoff_path}"))


# ─────────────────────── orchestration ─────────────────────────────
def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--dry-run", action="store_true",
                    help="Parse takeoffs + fetch QBO but don't write to Excel.")
    ap.add_argument("--project",
                    help="Filter to one project # (e.g. CP672). Case-insensitive.")
    ap.add_argument("--no-qbo", action="store_true",
                    help="Skip QBO join (fast local test of takeoff parsing).")
    ap.add_argument("--verbose", "-v", action="store_true",
                    help="Show debug logs (timestamps + module names).")
    args = ap.parse_args()

    # Interactive mode uses pretty terminal output (no log prefixes).
    # --verbose reverts to the standard timestamped log format for debugging.
    # Route logging to stdout (not stderr) so pretty print order is preserved
    # when stdout/stderr are merged.
    if args.verbose:
        logging.basicConfig(
            level=logging.DEBUG,
            format="%(asctime)s %(levelname)-7s %(name)s: %(message)s",
            stream=sys.stdout,
        )
    else:
        logging.basicConfig(
            level=logging.WARNING,
            format="  %(levelname)s: %(message)s",
            stream=sys.stdout,
        )

    # ── Header ──
    print()
    print(_Term.color(_Term.BOLD + _Term.CYAN, "  CP WIP Reader"))
    print(_Term.color(_Term.DIM, "  " + dt.datetime.now().strftime("%Y-%m-%d %H:%M")))

    _section("Configuration")
    _kv("Active dir",    _shorten(CP_ACTIVE_DIR))
    _kv("Completed dir", _shorten(CP_COMPLETED_DIR))
    _kv("WIP target",    _shorten(WIP_EXCEL_PATH))
    _kv("Write target",  f"{TEST_TAB!r} tab (guard allow-list: "
                         f"{', '.join(sorted(ALLOWED_WRITE_SHEETS))})")

    # ── Scan ──
    _section("Scanning CP folders on Synology")
    active_rows = scan_cp_folders(CP_ACTIVE_DIR, is_completed=False)
    completed_rows = scan_cp_folders(CP_COMPLETED_DIR, is_completed=True)
    rows = active_rows + completed_rows
    _kv("Active",    f"{len(active_rows)} folder(s)")
    _kv("Completed", f"{len(completed_rows)} folder(s)")
    _kv("Total",     f"{len(rows)} folder(s) scanned")

    if args.project:
        pf = args.project.upper()
        rows = [r for r in rows if r.project_num == pf]
        _kv(f"--project {pf}", f"{len(rows)} row(s) after filter")

    if not rows:
        print()
        print(_Term.color(_Term.AMBER, "  No CP projects to process — exiting"))
        return 0

    # ── QBO enrichment ──
    if not args.no_qbo:
        _section("Enriching with QBO Billed/Costs")
        print(f"  {_Term.color(_Term.DIM, 'Fetching...')}")
        t0 = dt.datetime.now()
        enrich_with_qbo(rows)
        elapsed = (dt.datetime.now() - t0).total_seconds()
        print(_Term.color(_Term.GREEN, f"  ✓ {len(rows)} project(s) enriched") +
              _Term.color(_Term.DIM, f" ({elapsed:.1f}s)"))
    else:
        _section("Skipping QBO join (--no-qbo)")

    # Flagged rows = the script couldn't verify a number → red in Excel
    # (the user 2026-07-13: bad-looking numbers render red).
    for row in rows:
        row.needs_review = bool(row.status_flags)

    # ── Write / dry-run report ──
    try:
        # Active-only default view (the user 2026-07-31: "don't want to see
        # Closed on default open") — same as the master tab.
        wrote = write_test_cp(rows, WIP_EXCEL_PATH, dry_run=args.dry_run,
                              default_filter_active=True,
                              title="CP WIP REPORT", summary=True)
    except WipWriteDenied as e:
        print(_Term.color(_Term.RED, f"  ✗ Guard blocked write: {e}"))
        return 2
    except FileNotFoundError as e:
        print(_Term.color(_Term.RED, f"  ✗ {e}"))
        return 3

    if not args.dry_run and wrote:
        _section("Done")
        print(_Term.color(_Term.GREEN, f"  ✓ Wrote {len(rows)} row(s) to {TEST_TAB!r}"))
    print()
    return 0


if __name__ == "__main__":
    sys.exit(main())
