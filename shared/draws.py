"""
draws.py — CP/commercial draw (AIA G702/G703) discovery + parsing.

Extracted from wip/cp_wip_reader.py (2026-07-16) the moment a second tool
(health-dashboard) needed it — repo rule: tools never import tools; common
code lives in shared/.

Draw folder detection (the user 2026-07-09): the WIP update for a CP project
comes from the LATEST draw (AIA G702/G703 payment application), not the
takeoff.
  • Container folder is named 'Draw' or 'Draws' (inclusive), never 'Drawings'.
  • Draw # is the SEQUENCE — the highest draw # wins, read from the filename
    or the numbered 'Draw #N' subfolder.
  • If no draw folder / no draw yet → callers fall back to the takeoff
    proposal (Original Contract Price); once Draw #1 lands, use the draw.
"""
from __future__ import annotations

import re
from pathlib import Path
from typing import List, Optional, Tuple

from openpyxl import load_workbook

DRAWS_FOLDER_RE = re.compile(r"^draws?\b", re.IGNORECASE)   # 'Draw' / 'Draws', not 'Drawings'
DRAW_NUM_RE = re.compile(r"draw\s*#?\s*(\d+)", re.IGNORECASE)  # 'Draw #4', 'DRAW#4', 'Draw 4'
G702_SHEET = "G702"


def coerce_float(v) -> Optional[float]:
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "").replace("$", "")
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def find_draws_folder(project_folder: Path) -> Optional[Path]:
    """Return the project's draw container ('Draw'/'Draws', case-insensitive),
    or None. 'Drawings' is excluded by the word-boundary in DRAWS_FOLDER_RE."""
    try:
        for entry in project_folder.iterdir():
            if entry.is_dir() and DRAWS_FOLDER_RE.match(entry.name.strip()):
                return entry
    except OSError:
        pass
    return None


def draw_num_from_name(name: str) -> Optional[int]:
    m = DRAW_NUM_RE.search(name)
    return int(m.group(1)) if m else None


def has_g702(xlsx: Path) -> bool:
    """True if the workbook has a G702 sheet (payment-application front page).
    read_only so it's a cheap header read, not a full parse."""
    try:
        wb = load_workbook(xlsx, read_only=True)
        try:
            return any(s.strip().lower() == G702_SHEET.lower() for s in wb.sheetnames)
        finally:
            wb.close()
    except Exception:
        return False


def find_latest_draw(project_folder: Path) -> Optional[Tuple[int, Path]]:
    """Locate the LATEST draw workbook for a project (the user 2026-07-09: draw # is
    the sequence — highest wins). Draw workbooks live in the 'Draws' container,
    either directly or inside a numbered 'Draw #N' subfolder; if there's no
    container, numbered draws directly under the project folder are also
    accepted. Only numbered draw subfolders are descended into — the Supplier
    Release folders (full of PDFs) are never opened. Returns (draw_num,
    draw_file) or None (→ caller falls back to the takeoff proposal)."""
    scan_root = find_draws_folder(project_folder) or project_folder
    candidates = []   # (draw_num, xlsx_path)
    try:
        entries = list(scan_root.iterdir())
    except OSError:
        return None
    for entry in entries:
        try:
            if entry.is_file():
                if entry.suffix.lower() in (".xlsx", ".xlsm") \
                        and not entry.name.startswith("~$"):
                    n = draw_num_from_name(entry.name)
                    if n is not None:
                        candidates.append((n, entry))
            elif entry.is_dir():
                n = draw_num_from_name(entry.name)      # numbered 'Draw #N' subfolder
                if n is not None:
                    for f in entry.iterdir():
                        if f.is_file() and f.suffix.lower() in (".xlsx", ".xlsm") \
                                and not f.name.startswith("~$"):
                            fn = draw_num_from_name(f.name)
                            candidates.append((fn if fn is not None else n, f))
        except OSError:
            continue
    if not candidates:
        return None
    # Highest draw # wins; require a real G702 so a stray xlsx can't win.
    for n, f in sorted(candidates, key=lambda x: x[0], reverse=True):
        if has_g702(f):
            return n, f
    return None


def _g702_value(ws, label_sub: str, max_scan: int = 14) -> Optional[float]:
    """Find `label_sub` (case-insensitive substring) in the G702 label column
    and read the first numeric value to its right (skipping '$'/blank cells).
    G702 line labels sit in column A with the amount a few columns right."""
    needle = label_sub.upper()
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if v and needle in str(v).upper():
                for off in range(1, max_scan + 1):
                    raw = ws.cell(r, c + off).value
                    num = coerce_float(raw)
                    if num is not None:
                        return num
                    if raw is None or str(raw).strip() in ("", "$", "USD", "-", "—"):
                        continue
                    break   # hit other text — stop scanning this label row
    return None


def read_draw_g702(draw_file: Path):
    """Read the WIP inputs off a draw's G702 (AIA payment application).
    Mapping (verified 2026-07-09 against CP585 Draws #1–#4):
      Contract Price  = Line 3  Contract Sum to Date (= Line 1 + Line 2)
      Approved COs    = Line 2  Net change by Change Orders
      Billed (gross)  = Line 4  Total Completed & Stored to Date
      Retainage       = Line 4 − Line 6  (Total Earned Less Retainage)
    RETAINAGE IS NOT read from the labeled 'Total Retainage' cell — that cell
    is unreliable across draws (0 / mismatched); Line 4 − Line 6 ties to the
    10% on Line 5a every time. Returns (dict, flags); never raises."""
    flags: List[str] = []
    try:
        wb = load_workbook(draw_file, data_only=True)
    except Exception as e:
        return None, [f"Draw read failed: {type(e).__name__}"]
    try:
        sheet = next((s for s in wb.sheetnames
                      if s.strip().lower() == G702_SHEET.lower()), None)
        if sheet is None:
            return None, ["Draw has no G702 sheet"]
        ws = wb[sheet]
        orig   = _g702_value(ws, "ORIGINAL CONTRACT SUM")          # Line 1
        net_co = _g702_value(ws, "NET CHANGE BY CHANGE ORDERS")    # Line 2
        c2d    = _g702_value(ws, "CONTRACT SUM TO DATE")           # Line 3
        billed = _g702_value(ws, "TOTAL COMPLETED")                # Line 4
        earned = _g702_value(ws, "TOTAL EARNED LESS RETAINAGE")    # Line 6
    finally:
        wb.close()

    # Contract: prefer Line 3; else reconstruct Line 1 + Line 2.
    if c2d is None and orig is not None:
        c2d = orig + (net_co or 0.0)
    # Original contract (base) so contract_price property = base + CO = Line 3.
    if orig is None and c2d is not None:
        orig = c2d - (net_co or 0.0)
    retainage = (billed - earned) if (billed is not None and earned is not None) else None

    data = {
        "orig_contract": orig, "net_co": net_co, "contract_to_date": c2d,
        "billed": billed, "earned_less_retainage": earned, "retainage": retainage,
    }
    missing = [k for k in ("contract_to_date", "billed") if data[k] is None]
    if missing:
        flags.append(f"Draw G702 missing {', '.join(missing)} — open & save the "
                     f"draw in Excel to refresh cached values")
    return data, flags
