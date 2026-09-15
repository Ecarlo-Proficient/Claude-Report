#!/usr/bin/env python3
"""
rp_review.py - the weekly RP review with the ops manager: every RP line on the WIP, WHERE
each number was grabbed from, and a SNAPSHOT of the source spreadsheet cut around the row
the number sits in (max 5 rows, that row highlighted) - the way the owner wants to see it
(2026-09-15: "instead of only showing the source file, create a snapshot of the excel of
only where it's at, max 5 rows with its row highlighted").

WHAT IT READS (all read-only)
  * the WIP master's RP tab ('Test - RP', or 'WIP-RP' if the office renamed it) - the four
    inputs per line, the «base» columns (what the last sync wrote), the QBO deep links
  * the owner's RP WIP file ('RP WIP', its CONTRACT FILE / ETC FILE paths, STATUS, ACTION;
    the 'Removed log' sheet = the finished lines and why they left)
  * the job folder on Common: the proposal PDF / takeoff bid sheet / cost sheet the number
    came from - located by VALUE (the exact amount) so the snapshot is honest: when the
    number is not in any document, the page says so
  * JobTread approved proposals (shared/jobtread, Touch ID once)
  * the daily crew schedules (shared/schedule_index for the last date; that day's sheet
    for the row snapshot)

WHAT IT WRITES
  ~/Library/Application Support/Proficient/rp-review/rp_review.json  (ACB_RP_REVIEW_DIR)
  plus a snapshot cache beside it (keyed by file path + mtime + value) so a rebuild costs
  seconds when nothing changed. The dashboard serves the JSON; the ops-manager marks live in
  the ledger DB (rp_review_mark + rp_review_mark_log) - the mark helpers are at the bottom.

  python3 ledger/rp_review.py [--no-jobtread] [--no-snapshots] [--limit N]
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import re
import sqlite3
import subprocess
import sys
import time
from pathlib import Path
from typing import Dict, List, Optional

_REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_REPO))

from openpyxl import load_workbook                     # noqa: E402
from PIL import Image, ImageDraw, ImageFont            # noqa: E402
from shared import paths, schedule as SCH, schedule_index as SI, jobtread, takeoff_etc as TK  # noqa: E402

OUT_DIR = Path(os.environ.get("ACB_RP_REVIEW_DIR",
                              Path.home() / "Library" / "Application Support" / "Proficient" / "rp-review"))
LEDGER_DB = Path(os.environ.get("ACB_LEDGER_DB",
                                Path.home() / "Library" / "Application Support" / "Proficient" / "ledger.sqlite3"))
RP_TABS = ("Test - RP", "WIP-RP")
MAX_ROWS = 5
MAX_COLS = 8
STALE_DAYS = 14
SNAP_VER = 3                 # bump when the snapshot shape changes - old cache entries are then ignored
IMG_DIR = OUT_DIR / "img"
_FONT = "/System/Library/Fonts/Supplemental/Arial.ttf"
_FONT_B = "/System/Library/Fonts/Supplemental/Arial Bold.ttf"
_THEME = {0: "FFFFFF", 1: "000000", 2: "E7E6E6", 3: "44546A", 4: "4472C4", 5: "ED7D31", 6: "A5A5A5", 7: "FFC000", 8: "5B9BD5", 9: "70AD47"}

_log_lines: List[str] = []
_WB: Dict[str, object] = {}          # in-run memo: data_only workbooks by path (the master is opened once, not 94 times)


def _wb(path: Path):
    k = str(path)
    if k not in _WB:
        _WB[k] = load_workbook(path, data_only=True)
    return _WB[k]


def log(msg: str) -> None:
    print(msg, flush=True)
    _log_lines.append(msg)


# ───────────────────────────── helpers ─────────────────────────────
def _num(v) -> Optional[float]:
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _fmt(v) -> str:
    """A cell for the snapshot: numbers with commas, dates mm/dd/yyyy, text as is."""
    if v is None:
        return ""
    if isinstance(v, (dt.datetime, dt.date)):
        return v.strftime("%m/%d/%Y")
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, (int, float)):
        return f"{v:,.2f}".rstrip("0").rstrip(".") if abs(v - round(v)) > 0.004 else f"{v:,.0f}"
    s = str(v)
    return s if len(s) <= 60 else s[:57] + "..."


def _master_path() -> Path:
    p = paths.get_path("WIP_EXCEL_PATH", paths.onedrive_base() / "Company Files - WIP Report" / "WIP - MASTER.xlsx")
    if not p.exists():
        alt = p.with_name("WIP - MASTER new.xlsx")
        if alt.exists():
            return alt
    return p


def _rp_file() -> Path:
    return paths.get_path("RP_WIP_FILE", paths.onedrive_base() / "RP WIP TO FIX_Final.xlsx")


# ───────────────────────────── snapshots ─────────────────────────────
class SnapCache:
    def __init__(self, path: Path):
        self.path = path
        try:
            self.d = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            self.d = {}
        self.hits = 0

    def key(self, f: Path, extra: str) -> Optional[str]:
        try:
            return f"v{SNAP_VER}|{f}|{int(f.stat().st_mtime)}|{extra}"
        except OSError:
            return None

    def get(self, k):
        v = self.d.get(k) if k else None
        if v is not None:
            self.hits += 1
        return v

    def put(self, k, v):
        if k:
            self.d[k] = v

    def save(self):
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.path.write_text(json.dumps(self.d), encoding="utf-8")


def _col_letter(c: int) -> str:
    s = ""
    while c:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s


def _snap_rows(ws, r0: int, cols: List[int], hi_row: int, title: str, path: Path, anchor: str) -> dict:
    """MAX_ROWS rows centred on hi_row (clipped to the sheet), the given columns."""
    lo = max(1, hi_row - 2)
    hi = min(ws.max_row, lo + MAX_ROWS - 1)
    lo = max(1, hi - MAX_ROWS + 1)
    rows = []
    for r in range(lo, hi + 1):
        rows.append({"n": r, "hi": r == hi_row,
                     "cells": [_fmt(ws.cell(r, c).value) for c in cols]})
    return {"kind": "xlsx", "file": path.name, "path": str(path), "sheet": ws.title, "anchor": anchor,
            "cols": [_col_letter(c) for c in cols], "rows": rows}


def _snap_rows_multi(ws, hi_rows: List[int], cols: List[int], path: Path, anchor: str) -> dict:
    """Up to MAX_ROWS rows: every highlighted row plus the row above each, oldest first."""
    want = sorted({r for h in hi_rows for r in (h - 1, h) if r >= 1})
    while len(want) > MAX_ROWS:
        extra = [r for r in want if r not in hi_rows]
        if not extra:
            break
        want.remove(extra[0])
    rows = [{"n": r, "hi": r in hi_rows, "cells": [_fmt(ws.cell(r, c).value) for c in cols]} for r in want]
    return {"kind": "xlsx", "file": path.name, "path": str(path), "sheet": ws.title, "anchor": anchor,
            "cols": [_col_letter(c) for c in cols], "rows": rows}


def _snap_cost_bands(path: Path, value: float, cache: SnapCache) -> Optional[dict]:
    """The ETC is the SUM of the cost sheet's band subtotals (SL + PR for a slab, FW for
    flatwork - shared/takeoff_etc). Find the band set that adds up to `value` and cut those
    subtotal rows (each highlighted)."""
    k = cache.key(path, f"bands={value:.2f}")
    hit = cache.get(k)
    if hit is not None:
        return hit or None
    out = None
    try:
        wb = _wb(path)
        for n in wb.sheetnames:
            if "COST" not in n.upper():
                continue
            bands = TK._cost_sheet_totals(wb[n])
            subs = {b: v for b, v in bands.items() if v.get("sub") is not None and v.get("cell")}
            combos = [("SL", "PR"), ("SL",), ("FW",), ("SL", "PR", "FW"), ("PR",)]
            for combo in combos:
                if all(b in subs for b in combo) and abs(sum(subs[b]["sub"] for b in combo) - value) < 1.0:
                    ws = wb[n]
                    rows = [ws[subs[b]["cell"]].row for b in combo]
                    col = ws[subs[combo[0]]["cell"]].column
                    out = _snap_rows_multi(ws, rows, _window_cols(col, ws.max_column), path, " + ".join(subs[b]["cell"] for b in combo))
                    out["note"] = " + ".join(f"{b} {subs[b]['sub']:,.2f}" for b in combo) + f" = {value:,.2f}"
                    break
            if out:
                break
    except Exception as e:                                        # noqa: BLE001
        log(f"    ! cost bands {path.name}: {type(e).__name__}")
    cache.put(k, out or {})
    return out


def _window_cols(anchor_col: int, max_col: int) -> List[int]:
    lo = max(1, anchor_col - (MAX_COLS - 2))
    hi = min(max_col, lo + MAX_COLS - 1)
    lo = max(1, hi - MAX_COLS + 1)
    return list(range(lo, hi + 1))


def _snap_xlsx_value(path: Path, value: float, cache: SnapCache, sheets_like: Optional[str] = None) -> Optional[dict]:
    """Find `value` in the workbook (bid/cost sheets first) and cut the rows around it."""
    k = cache.key(path, f"v={value:.2f}|{sheets_like or ''}")
    hit = cache.get(k)
    if hit is not None:
        return hit or None
    try:
        wb = load_workbook(path, data_only=True, read_only=True)
    except Exception as e:                                        # noqa: BLE001
        cache.put(k, {}); log(f"    ! cannot open {path.name}: {type(e).__name__}"); return None
    names = wb.sheetnames
    order = sorted(names, key=lambda n: (0 if (sheets_like and sheets_like in n.upper()) else
                                         1 if ("BID" in n.upper() or "COST" in n.upper()) else 2))
    out = None
    for n in order:
        ws = wb[n]
        for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 400)):
            for c in row:
                v = c.value
                if isinstance(v, (int, float)) and not isinstance(v, bool) and abs(v - value) < 0.6:
                    ws2 = _wb(path)[n]                            # random access for the window
                    cols = _window_cols(c.column, ws2.max_column)
                    out = _snap_rows(ws2, c.row, cols, c.row, n, path, c.coordinate)
                    break
            if out:
                break
        if out:
            break
    wb.close()
    cache.put(k, out or {})
    return out


def _snap_xlsx_row(path: Path, sheet: str, row: int, cols: List[int], cache: SnapCache, anchor_col: int) -> Optional[dict]:
    k = cache.key(path, f"r={sheet}|{row}|{','.join(map(str, cols))}")
    hit = cache.get(k)
    if hit is not None:
        return hit or None
    try:
        ws = _wb(path)[sheet]
    except Exception as e:                                        # noqa: BLE001
        cache.put(k, {}); log(f"    ! cannot open {path.name}: {type(e).__name__}"); return None
    out = _snap_rows(ws, row, cols, row, sheet, path, f"{_col_letter(anchor_col)}{row}")
    cache.put(k, out)
    return out


def _snap_pdf(path: Path, value: Optional[float], cache: SnapCache) -> Optional[dict]:
    """pdftotext -layout; the 5 lines around the line carrying the amount (or SUB TOTAL)."""
    k = cache.key(path, f"pdf={value if value is None else round(value, 2)}")
    hit = cache.get(k)
    if hit is not None:
        return hit or None
    try:
        txt = subprocess.run(["pdftotext", "-layout", str(path), "-"], capture_output=True,   # noqa: S603,S607
                             text=True, timeout=60).stdout
    except (OSError, subprocess.SubprocessError):
        cache.put(k, {}); return None
    lines = [ln.rstrip() for ln in txt.splitlines()]
    needle = f"{value:,.2f}" if value is not None else None
    idx = None
    for i, ln in enumerate(lines):
        if needle and needle in ln:
            idx = i
    if idx is None:
        for i, ln in enumerate(lines):
            if "SUB TOTAL" in ln.upper():
                idx = i
    if idx is None:
        cache.put(k, {}); return None
    nonempty = [i for i, ln in enumerate(lines) if ln.strip()]
    pos = nonempty.index(idx) if idx in nonempty else 0
    pick = nonempty[max(0, pos - 2): pos + 3][:MAX_ROWS]
    found = bool(needle and needle in lines[idx])
    m = re.search(r"\$\s*([\d,]+\.\d{2})", lines[idx])
    out = {"kind": "pdf", "file": path.name, "path": str(path), "anchor": f"line {idx + 1}",
           "rows": [{"n": i + 1, "hi": i == idx, "cells": [re.sub(r"\s{3,}", "   ", lines[i]).strip()]} for i in pick],
           "cols": ["text"], "found": found,
           "note": "" if found else (f"this proposal says {m.group(1)}, not {needle}" if (m and needle) else "the amount is not in this proposal")}
    cache.put(k, out)
    return out


def _find_in_folder(folder: Path, value: float, cache: SnapCache, want: str) -> Optional[dict]:
    """The document in the job folder that carries `value`: proposal PDFs and takeoff
    workbooks at the top level only (the ARCHIVE subfolder is skipped on purpose)."""
    if not folder.exists() or not folder.is_dir():
        return None
    try:
        items = sorted(folder.iterdir(), key=lambda p: p.stat().st_mtime, reverse=True)
    except OSError:
        return None
    pdfs = [p for p in items if p.suffix.lower() == ".pdf" and re.search(r"proposal|bid|flatwork|ftw", p.name, re.I)]
    xls = [p for p in items if p.suffix.lower() in (".xlsm", ".xlsx") and not p.name.startswith("~$")]
    seq = (pdfs + xls) if want == "contract" else (xls + pdfs)
    for p in seq[:12]:
        if p.suffix.lower() == ".pdf":
            s = _snap_pdf(p, value, cache)
            if s and s.get("found"):
                return s
        else:
            s = _snap_xlsx_value(p, value, cache, "COST" if want == "etc" else "BID")
            if not s and want == "etc":
                s = _snap_cost_bands(p, value, cache)
            if s:
                return s
    return None



# ───────────────────────────── pictures (the owner 2026-09-15: "a literal screenshot of the file with the file name") ─────
def _font(size: int, bold: bool = False):
    try:
        return ImageFont.truetype(_FONT_B if bold and Path(_FONT_B).exists() else _FONT, size)
    except OSError:
        return ImageFont.load_default()


def _rgb(color):
    """An openpyxl Color -> (r, g, b) or None (indexed / auto colours are skipped)."""
    if color is None:
        return None
    try:
        if color.type == "rgb" and isinstance(color.rgb, str) and len(color.rgb) >= 6:
            h = color.rgb[-6:]
        elif color.type == "theme" and color.theme in _THEME:
            h = _THEME[color.theme]
        else:
            return None
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
        t = float(color.tint or 0)
        if t > 0:
            r, g, b = (int(v + (255 - v) * t) for v in (r, g, b))
        elif t < 0:
            r, g, b = (int(v * (1 + t)) for v in (r, g, b))
        return (r, g, b)
    except (AttributeError, ValueError, TypeError):
        return None


def _render_sheet(ws, rows: List[int], cols: List[int], hi_rows: List[int], title: str, out: Path) -> Optional[str]:
    """Draw the cells as Excel shows them - values, fills, bold, font colour, column widths,
    row numbers and column letters - with the file name on top and the highlighted rows
    outlined. Returns the PNG path (relative to IMG_DIR) or None."""
    if not rows:
        return None
    widths = []
    for c in cols:
        w = ws.column_dimensions[_col_letter(c)].width if _col_letter(c) in ws.column_dimensions else None
        widths.append(max(66, min(300, int((w or 8.43) * 7.4))))
    row_h, hdr_h, rn_w, title_h = 22, 20, 36, 28
    W = rn_w + sum(widths) + 2
    H = title_h + hdr_h + row_h * len(rows) + 2
    img = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(img)
    f, fb, fs = _font(12), _font(12, True), _font(11)
    d.rectangle([0, 0, W, title_h], fill=(31, 36, 48))
    d.text((9, 7), title[:140], fill="white", font=fb)
    y0 = title_h
    d.rectangle([0, y0, W, y0 + hdr_h], fill=(236, 238, 241))
    x = rn_w
    for c, w in zip(cols, widths):
        d.text((x + w / 2 - 4, y0 + 4), _col_letter(c), fill=(95, 95, 95), font=fs)
        d.line([x, y0, x, H], fill=(214, 214, 214))
        x += w
    merged = {}
    _NONE = object()
    for rng in ws.merged_cells.ranges:
        merged[(rng.min_row, rng.min_col)] = rng
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                if (r, c) != (rng.min_row, rng.min_col):
                    merged[(r, c)] = None
    for i, r in enumerate(rows):
        y = y0 + hdr_h + i * row_h
        d.rectangle([0, y, rn_w, y + row_h], fill=(236, 238, 241))
        d.text((6, y + 5), str(r), fill=(95, 95, 95), font=fs)
        x = rn_w
        for c, w in zip(cols, widths):
            cell = ws.cell(r, c)
            fill = _rgb(cell.fill.fgColor) if (cell.fill is not None and cell.fill.fill_type == "solid") else None
            if fill and fill != (255, 255, 255):
                d.rectangle([x, y, x + w, y + row_h], fill=fill)
            m = merged.get((r, c), _NONE)
            if m is None:                                   # inside a merge, not its first cell
                x += w
                continue
            span_w = w
            if m is not _NONE:
                span_w = sum(ww for cc, ww in zip(cols, widths) if m.min_col <= cc <= m.max_col)
            txt = _fmt(cell.value)
            if txt:
                fc = _rgb(cell.font.color) if cell.font is not None else None
                fnt = fb if (cell.font is not None and cell.font.b) else f
                while txt and d.textlength(txt, font=fnt) > span_w - 8:
                    txt = txt[:-2] + "…" if len(txt) > 3 else ""
                tw = d.textlength(txt, font=fnt)
                right = isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool)
                tx = x + span_w - 5 - tw if right else x + 5
                d.text((tx, y + 5), txt, fill=fc or (0, 0, 0), font=fnt)
            x += w
        d.line([0, y + row_h, W, y + row_h], fill=(214, 214, 214))
    for i, r in enumerate(rows):
        if r in hi_rows:
            y = y0 + hdr_h + i * row_h
            d.rectangle([rn_w + 1, y + 1, W - 2, y + row_h - 1], outline=(232, 122, 0), width=3)
    out.parent.mkdir(parents=True, exist_ok=True)
    img.save(out, optimize=True)
    return str(out.relative_to(IMG_DIR))


def _pdf_page_png(path: Path, value: Optional[float], out: Path) -> Optional[dict]:
    """The proposal page that carries the amount (else the SUB TOTAL page) as a PNG."""
    try:
        info = subprocess.run(["pdfinfo", str(path)], capture_output=True, text=True, timeout=30).stdout   # noqa: S603,S607
        pages = int(re.search(r"Pages:\s+(\d+)", info).group(1))
    except (OSError, subprocess.SubprocessError, AttributeError, ValueError):
        pages = 1
    needle = f"{value:,.2f}" if value is not None else None
    pick, found, says = None, False, ""
    for pg in range(1, min(pages, 8) + 1):
        try:
            txt = subprocess.run(["pdftotext", "-f", str(pg), "-l", str(pg), "-layout", str(path), "-"],   # noqa: S603,S607
                                 capture_output=True, text=True, timeout=30).stdout
        except (OSError, subprocess.SubprocessError):
            continue
        if needle and needle in txt:
            pick, found = pg, True
            break
        if pick is None and "SUB TOTAL" in txt.upper():
            pick = pg
            m = re.search(r"SUB TOTAL:?\s*\$?\s*([\d,]+\.\d{2})", txt.upper())
            says = m.group(1) if m else ""
    if pick is None:
        return None
    out.parent.mkdir(parents=True, exist_ok=True)
    stem = out.with_suffix("")
    try:
        subprocess.run(["pdftoppm", "-png", "-r", "72", "-f", str(pick), "-l", str(pick), "-singlefile", str(path), str(stem)],   # noqa: S603,S607
                       capture_output=True, timeout=60)
    except (OSError, subprocess.SubprocessError):
        return None
    if not out.exists():
        return None
    return {"kind": "pdf", "file": path.name, "path": str(path), "page": pick, "pages": pages, "found": found,
            "img": str(out.relative_to(IMG_DIR)),
            "note": "" if found else (f"this proposal says {says}, not {needle}" if (says and needle) else "the amount is not on this proposal")}


def _find_value_cell(path: Path, value: float, sheets_like: Optional[str] = None):
    """(sheet name, row, col) of the first cell equal to `value` (bid / cost sheets first)."""
    try:
        wb = _wb(path)
    except Exception:                                             # noqa: BLE001
        return None
    order = sorted(wb.sheetnames, key=lambda n: (0 if (sheets_like and sheets_like in n.upper()) else
                                                 1 if ("BID" in n.upper() or "COST" in n.upper()) else 2))
    for n in order:
        ws = wb[n]
        for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 400)):
            for c in row:
                v = c.value
                if isinstance(v, (int, float)) and not isinstance(v, bool) and abs(v - value) < 0.6:
                    return n, c.row, c.column
    return None


def _sheet_page_png(path: Path, sheet: str, hi_rows: List[int], out: Path, whole: bool = True) -> Optional[dict]:
    """The sheet as a picture: the whole used page (up to 70 rows x 10 columns) with the
    rows highlighted, or only those rows when whole=False."""
    try:
        ws = _wb(path)[sheet]
    except Exception:                                             # noqa: BLE001
        return None
    last = max(hi_rows) if hi_rows else 1
    if whole:
        top = 1
        bottom = min(ws.max_row, max(last + 4, 24), 70)
        rows = list(range(top, bottom + 1))
    else:
        rows = sorted(set(hi_rows))
    cols = list(range(1, min(ws.max_column, 10) + 1))
    rel = _render_sheet(ws, rows, cols, hi_rows, f"{path.name}  ·  sheet {sheet}", out)
    if not rel:
        return None
    return {"kind": "xlsx", "file": path.name, "path": str(path), "sheet": sheet, "img": rel,
            "anchor": ", ".join(f"row {r}" for r in hi_rows)}


def _row_strip_png(path: Path, sheet: str, row: int, cols: List[int], out: Path, header_row: Optional[int] = None) -> Optional[dict]:
    """One row of a sheet as a picture (plus its header row when given)."""
    try:
        ws = _wb(path)[sheet]
    except Exception:                                             # noqa: BLE001
        return None
    rows = ([header_row] if header_row else []) + [row]
    rel = _render_sheet(ws, rows, cols, [row], f"{path.name}  ·  sheet {sheet}", out)
    return {"kind": "xlsx", "file": path.name, "path": str(path), "sheet": sheet, "img": rel, "anchor": f"row {row}"} if rel else None


_SCHED_RECS: Dict[str, list] = {}


def _sched_recs(path: Path) -> list:
    k = str(path)
    if k not in _SCHED_RECS:
        try:
            _SCHED_RECS[k] = SCH.parse_main_schedule(path)
        except Exception as e:                                    # noqa: BLE001
            log(f"    ! schedule {path.name}: {type(e).__name__}")
            _SCHED_RECS[k] = []
    return _SCHED_RECS[k]


def _sched_row_of(path: Path, base: str, want_ftw: bool):
    """(sheet name, header row, job row) on that day's Main Schedule for the slab or the flatwork line."""
    try:
        wb = _wb(path)
        name = next((n for n in wb.sheetnames if "MAIN" in n.upper()), wb.sheetnames[0])
        ws = wb[name]
    except Exception:                                             # noqa: BLE001
        return None
    hdr = None
    hits = []
    section = ""
    for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 600)):
        texts = [str(c.value or "").strip().upper() for c in row[:12]]
        if hdr is None and "ADDRESS" in texts:
            hdr = row[0].row
        first = texts[0] if texts else ""
        if first and all(not t for t in texts[1:6]):
            section = first
        if any(t == base or t.startswith(base + " ") for t in texts):
            stage = " ".join(texts)
            hits.append((row[0].row, bool(SI._FTW_RE.search(section + " " + stage))))
    if not hits:
        return None
    pick = next((r for r, ftw in hits if ftw == want_ftw), hits[0][0])
    return name, hdr, pick


def _timeline(line: str, files_by_date: Dict[dt.date, Path], index_files: dict) -> List[dict]:
    """Every day the line was on a crew schedule: [{date, section, stage}] oldest first."""
    base = line.replace("-FTW", "")
    want_ftw = line.endswith("-FTW")
    days = []
    for key, rec in index_files.items():
        if line not in rec.get("jobs", []):
            continue
        try:
            d = dt.date.fromisoformat(rec["date"])
        except (KeyError, ValueError):
            continue
        f = files_by_date.get(d) or Path(key)
        if not f.exists():
            continue
        recs = [r for r in _sched_recs(f) if (r.get("proj") or "").upper() == base]
        pick = None
        for r in recs:
            is_ftw = bool(SI._FTW_RE.search(f"{r.get('section', '')} {r.get('stage', '')}"))
            if is_ftw == want_ftw:
                pick = r
                break
        if pick is None and recs and not want_ftw:
            pick = recs[0]
        if pick is None:
            continue
        days.append({"date": d.isoformat(), "section": (pick.get("section") or "").strip(), "stage": (pick.get("stage") or "").strip()})
    days.sort(key=lambda x: x["date"])
    return days


def _runs(days: List[dict]) -> List[dict]:
    """Consecutive schedule days with the same section + task folded into one line."""
    out = []
    for d in days:
        if out and out[-1]["section"] == d["section"] and out[-1]["stage"] == d["stage"]:
            out[-1]["to"] = d["date"]
            out[-1]["days"] += 1
        else:
            out.append({"from": d["date"], "to": d["date"], "days": 1, "section": d["section"], "stage": d["stage"]})
    return out


# ───────────────────────────── readers ─────────────────────────────
def read_master(path: Path) -> dict:
    wb = load_workbook(path)                       # formulas kept: we read INPUT cells + base columns
    tab = next((t for t in RP_TABS if t in wb.sheetnames), None)
    if not tab:
        raise SystemExit(f"no RP tab ({' / '.join(RP_TABS)}) in {path.name}")
    ws = wb[tab]
    hr = next(r for r in range(1, 16) if ws.cell(r, 2).value == "PROJECT #")
    hdr = {str(ws.cell(hr, c).value or "").strip(): c for c in range(1, ws.max_column + 1)}
    rows = {}
    for r in range(hr + 1, ws.max_row + 1):
        p = str(ws.cell(r, 2).value or "").strip().upper()
        if not p.startswith("RP"):
            continue
        d = {k: ws.cell(r, c).value for k, c in hdr.items() if k}
        d["_row"] = r
        for k in ("COSTS TO DATE", "BILLED TO DATE"):
            c = ws.cell(r, hdr[k]) if k in hdr else None
            d[k + " LINK"] = c.hyperlink.target if (c is not None and c.hyperlink) else None
        rows[p] = d
    m = re.search(r"REPORT DATE:\s*(.+)", str(ws.cell(2, 2).value or ""))
    return {"tab": tab, "hdr": hdr, "rows": rows, "report_date": m.group(1).strip() if m else "",
            "path": path, "renamed": tab != RP_TABS[0]}


def read_rp_file(path: Path) -> dict:
    wb = load_workbook(path)
    ws = wb["RP WIP"]
    rows, band = {}, ""
    for r in range(3, ws.max_row + 1):
        v = str(ws.cell(r, 1).value or "").strip()
        if v.startswith("⚠") or v.upper().startswith("FTW BACKLOG"):
            band = v; continue
        j = v.upper()
        if not j.startswith("RP"):
            continue
        rows.setdefault(j, {"row": r, "address": ws.cell(r, 2).value, "builder": ws.cell(r, 3).value,
                            "contract": _num(ws.cell(r, 4).value), "etc": _num(ws.cell(r, 5).value),
                            "status": ws.cell(r, 12).value, "action": ws.cell(r, 13).value,
                            "contract_file": ws.cell(r, 15).value, "etc_file": ws.cell(r, 16).value,
                            "band": band})
    removed = []
    if "Removed log" in wb.sheetnames:
        lg = wb["Removed log"]
        hdr = [str(lg.cell(1, c).value or "").strip() for c in range(1, lg.max_column + 1)]
        for r in range(2, lg.max_row + 1):
            if not lg.cell(r, 2).value:
                continue
            d = {hdr[c - 1]: lg.cell(r, c).value for c in range(1, len(hdr) + 1) if hdr[c - 1]}
            d["_row"] = r
            removed.append(d)
    return {"rows": rows, "removed": removed, "path": path}


def _schedule_snap(job: str, day: Optional[dt.date], files: Dict[dt.date, Path], cache: SnapCache) -> Optional[dict]:
    f = files.get(day) if day else None
    if not f:
        return None
    k = cache.key(f, f"sched={job}")
    hit = cache.get(k)
    if hit is not None:
        return hit or None
    out = None
    try:
        wb = load_workbook(f, data_only=True, read_only=True)
        name = next((n for n in wb.sheetnames if "MAIN" in n.upper()), wb.sheetnames[0])
        ws = wb[name]
        base = job.replace("-FTW", "")
        want_ftw = job.endswith("-FTW")
        hits = []
        section = ""
        for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 600)):
            first = str(row[0].value or "").strip() if row else ""
            texts = [str(c.value or "").strip().upper() for c in row[:12]]
            if first and all(not str(c.value or "").strip() for c in row[1:6]):
                section = first.upper()
            if any(t == base or t.startswith(base + " ") for t in texts):
                hits.append((row[0].row, "FLATWORK" in section))
        pick = next((r for r, ftw in hits if ftw == want_ftw), hits[0][0] if hits else None)
        wb.close()
        if pick:
            ws2 = _wb(f)[name]
            # columns B..G = project #, address, city, builder, description, date (the super and
            # crew columns A / H are people - left out on purpose)
            out = _snap_rows(ws2, pick, list(range(2, min(ws2.max_column, 7) + 1)), pick, name, f, f"B{pick}")
    except Exception as e:                                        # noqa: BLE001
        log(f"    ! schedule {f.name}: {type(e).__name__}")
    cache.put(k, out or {})
    return out


# ───────────────────────────── the build ─────────────────────────────
def classify_contract(line: str, d: dict, r: dict, typed: bool):
    v = _num(d.get("ORIGINAL CONTRACT"))
    cf = str(r.get("contract_file") or "")
    if typed and v:
        hint = "equals the invoice amount" if abs(v - (_num(d.get("BILLED TO DATE")) or -1)) < 1 else "no paper named"
        return "typed", "Typed on the master", f"blank at the last sync, typed in by hand since ({hint}); OneDrive version history names who"
    if not v:
        return "blank", "BLANK", "no contract anywhere - the estimator enters it in JobTread"
    if cf.lower().endswith(".pdf"):
        return "pdf", "Proposal PDF", Path(cf).name
    if cf.lower().endswith((".xlsm", ".xlsx")):
        return "takeoff", "Takeoff bid sheet", Path(cf).name
    if cf:
        return "folder", "Job folder", "read from the proposal / takeoff in the folder"
    if "FTW BACKLOG" in str(r.get("band") or "").upper():
        return "gl", "General List", "flatwork price typed in the General List"
    return "master", "WIP master", "carried from an earlier master; no document on file"


def classify_etc(line: str, d: dict, r: dict, typed: bool):
    v = _num(d.get("ORIGINAL ESTIMATED COST"))
    ef = str(r.get("etc_file") or "")
    if typed and v:
        return "typed", "Typed on the master", "blank at the last sync, typed in by hand since, no cost sheet named"
    if not v:
        return "blank", "BLANK", "no budget anywhere - the estimator enters it in JobTread"
    if ef.lower().endswith((".xlsm", ".xlsx")):
        return "takeoff", "Takeoff cost sheet", Path(ef).name
    if ef:
        return "folder", "Job folder", "read from the takeoff cost sheet in the folder"
    if "manual" in str(r.get("action") or "").lower():
        return "manual", "Typed by hand", str(r.get("action"))[:90]
    return "master", "WIP master", "carried from an earlier master; no cost sheet on file"


def build(no_jobtread: bool = False, no_snapshots: bool = False, limit: int = 0) -> dict:
    t0 = time.time()
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    cache = SnapCache(OUT_DIR / "snapshot_cache.json")
    mp, fp = _master_path(), _rp_file()
    log(f"master: {mp.name}   RP file: {fp.name}")
    M = read_master(mp)
    F = read_rp_file(fp)
    log(f"RP tab '{M['tab']}' ({'RENAMED - the sync writes Test - RP' if M['renamed'] else 'as expected'}): "
        f"{len(M['rows'])} lines · RP file: {len(F['rows'])} active, {len(F['removed'])} removed")

    # last schedule date per job (cached index; refresh when the share is mounted)
    try:
        seen = SI.last_seen(refresh=Path(str(SCH.SCHEDULE_DIR)).exists())
    except Exception as e:                                        # noqa: BLE001
        log(f"  schedule index unavailable: {type(e).__name__}"); seen = {}
    sched_files = {d: f for d, f in SCH.all_schedule_files()} if Path(str(SCH.SCHEDULE_DIR)).exists() else {}

    jt: Dict[str, dict] = {}
    if not no_jobtread:
        nums = {ln.replace("-FTW", "") for ln in M["rows"]} | {str(x.get("JOB #") or "").upper().replace("-FTW", "") for x in F["removed"]}
        jt = jobtread.jobs(nums, log=log)
        log(f"JobTread: {sum(1 for v in jt.values() if v['docs'])} of {len(nums)} jobs have an approved proposal; {len(jt)} exist there")
    index_files = SI._load(SI.CACHE).get("files", {})

    def folder_of(r: dict) -> Optional[str]:
        for k in ("contract_file", "etc_file"):
            v = str(r.get(k) or "")
            if v:
                pth = Path(v)
                return str(pth if pth.is_dir() else pth.parent)
        return None

    def jt_block(base: str, line: str) -> dict:
        j = jt.get(base) or {}
        docs = j.get("docs") or []
        price, cost, date = (docs[-1] if docs else (None, None, ""))
        return {"price": price, "cost": cost, "date": date, "count": len(docs), "url": j.get("url") or "",
                "exists": bool(j),
                "scope_note": "the base job's proposal - may be the slab, not the flatwork" if (docs and line.endswith("-FTW")) else ""}

    as_of = dt.date.today()
    newest = max(sched_files) if sched_files else None          # the latest crew schedule on file
    current = []
    lines = list(M["rows"].items())[:limit] if limit else list(M["rows"].items())
    for i, (line, d) in enumerate(lines, 1):
        r = F["rows"].get(line, {})
        K, E = _num(d.get("ORIGINAL CONTRACT")), _num(d.get("ORIGINAL ESTIMATED COST"))
        bK, bE = _num(d.get("«base» base_contract")), _num(d.get("«base» base_etc"))
        typed_k = (K or 0) != (bK or 0) and "«base» base_contract" in d
        typed_e = (E or 0) != (bE or 0) and "«base» base_etc" in d
        kc, kl, kd = classify_contract(line, d, r, typed_k)
        ec, el, ed = classify_etc(line, d, r, typed_e)
        base = line.replace("-FTW", "")
        jtb = jt_block(base, line)
        jt_price = jtb["price"]
        days = _timeline(line, sched_files, index_files) if sched_files else []
        last = dt.date.fromisoformat(days[-1]["date"]) if days else (seen.get(line) or seen.get(base))
        first = dt.date.fromisoformat(days[0]["date"]) if days else None
        stale = bool(last and (as_of - last).days > STALE_DAYS)
        costs, billed = _num(d.get("COSTS TO DATE")), _num(d.get("BILLED TO DATE"))
        flags = []
        if typed_k or typed_e:
            flags.append("contract / ETC typed on the master after the last sync, no source named - which paper?")
        fk, fe = _num(r.get("contract")), _num(r.get("etc"))
        if K and fk and abs(K - fk) > 1:
            flags.append(f"the RP file says contract {fk:,.0f}")
        if E and fe and abs(E - fe) > 1:
            flags.append(f"the RP file says ETC {fe:,.0f}")
        if jt_price and K and not line.endswith("-FTW") and abs(jt_price - K) > 1:
            flags.append(f"JobTread price {jt_price:,.0f} differs from the WIP")
        if not jtb["count"] and not no_jobtread:
            flags.append("no approved proposal in JobTread" if jtb["exists"] else "not in JobTread")
        if stale:
            flags.append(f"not on the schedule since {last.strftime('%m/%d/%Y')}")
        if K and billed and billed >= K * 0.99:
            if newest and last == newest:
                flags.append("billed out, still on today's schedule")
            elif last:
                flags.append(f"billed out; last on the schedule {last.strftime('%m/%d/%Y')} - finished?")
        if K and costs and E and costs > E:
            flags.append(f"costs {costs:,.0f} already over the ETC {E:,.0f}")

        pics = {}
        src_k, src_e = r.get("contract_file"), r.get("etc_file")
        if not no_snapshots:
            jd = IMG_DIR / line
            # 1. the schedule: the job's own row on its first and last day
            for tag, day in (("first", first), ("last", last)):
                f = sched_files.get(day) if day else None
                if not f:
                    continue
                hit = _sched_row_of(f, base, line.endswith("-FTW"))
                if hit:
                    name, hdr, row = hit
                    pics["sched_" + tag] = _row_strip_png(f, name, row, list(range(2, 8)), jd / f"sched_{tag}.png", hdr)
                    if pics["sched_" + tag]:
                        pics["sched_" + tag]["date"] = day.isoformat()
            # 2. the contract: the page it sits on
            if K and kc in ("pdf", "takeoff", "folder"):
                cand = None
                if kc == "pdf" and src_k and Path(src_k).exists():
                    cand = _pdf_page_png(Path(src_k), K, jd / "contract.png")
                elif kc == "takeoff" and src_k and Path(src_k).exists():
                    hit = _find_value_cell(Path(src_k), K, "BID")
                    if hit:
                        cand = _sheet_page_png(Path(src_k), hit[0], [hit[1]], jd / "contract.png")
                elif kc == "folder" and src_k:
                    folder = Path(src_k)
                    try:
                        items = sorted(folder.iterdir(), key=lambda q: q.stat().st_mtime, reverse=True) if folder.is_dir() else []
                    except OSError:
                        items = []
                    for q in [x for x in items if x.suffix.lower() == ".pdf" and re.search(r"proposal|bid|flatwork|ftw", x.name, re.I)][:6]:
                        c2 = _pdf_page_png(q, K, jd / "contract.png")
                        if c2 and c2.get("found"):
                            cand = c2
                            break
                    if not cand:
                        for q in [x for x in items if x.suffix.lower() in (".xlsm", ".xlsx") and not x.name.startswith("~$")][:6]:
                            hit = _find_value_cell(q, K, "BID")
                            if hit:
                                cand = _sheet_page_png(q, hit[0], [hit[1]], jd / "contract.png")
                                break
                if cand:
                    pics["contract"] = cand
                    src_k = cand["path"]                        # the exact file the number sits in
                    if cand.get("kind") == "pdf" and not cand.get("found"):
                        flags.append(cand.get("note") or "the proposal PDF says a different number")
                else:
                    flags.append("contract amount not found in the folder's proposal or takeoff")
            # 3. the ETC: the cost sheet page (the band subtotals that add up)
            if E and ec in ("takeoff", "folder"):
                cand = None
                files = []
                if ec == "takeoff" and src_e and Path(src_e).exists():
                    files = [Path(src_e)]
                elif ec == "folder" and src_e:
                    folder = Path(src_e)
                    try:
                        files = sorted([x for x in folder.iterdir() if x.suffix.lower() in (".xlsm", ".xlsx") and not x.name.startswith("~$")],
                                       key=lambda q: q.stat().st_mtime, reverse=True)[:6] if folder.is_dir() else []
                    except OSError:
                        files = []
                for q in files:
                    hit = _find_value_cell(q, E, "COST")
                    if hit:
                        cand = _sheet_page_png(q, hit[0], [hit[1]], jd / "etc.png")
                    else:
                        bands = _snap_cost_bands(q, E, cache)
                        if bands:
                            cand = _sheet_page_png(q, bands["sheet"], [row["n"] for row in bands["rows"] if row["hi"]], jd / "etc.png")
                            if cand:
                                cand["note"] = bands.get("note", "")
                    if cand:
                        break
                if cand:
                    pics["etc"] = cand
                    src_e = cand["path"]
                else:
                    flags.append("ETC amount not found in the folder's cost sheet")
            # 4. the WIP master row and the RP file row
            hdr = M["hdr"]
            cols = [hdr[k] for k in ("PROJECT #", "PROJECT NAME", "ORIGINAL CONTRACT", "APPROVED COs",
                                     "ORIGINAL ESTIMATED COST", "CO COSTS", "COSTS TO DATE", "BILLED TO DATE") if k in hdr]
            hr_row = next(rr for rr in range(1, 16) if _wb(mp)[M["tab"]].cell(rr, 2).value == "PROJECT #")
            pics["master"] = _row_strip_png(mp, M["tab"], d["_row"], cols, jd / "master.png", hr_row)
            if r:
                pics["rpfile"] = _row_strip_png(fp, "RP WIP", r["row"], list(range(1, 8)), jd / "rpfile.png", 2)

        current.append({
            "line": line, "ftw": line.endswith("-FTW"), "name": d.get("PROJECT NAME"), "builder": d.get("BUILDER"),
            "type": d.get("TYPE"), "status": d.get("STATUS"), "category": d.get("CATEGORY"),
            "contract": K, "etc": E, "costs": costs, "billed": billed,
            "costs_link": d.get("COSTS TO DATE LINK"), "billed_link": d.get("BILLED TO DATE LINK"),
            "folder": folder_of(r),
            "src": {"contract": {"kind": kc, "label": kl, "detail": kd, "file": src_k},
                    "etc": {"kind": ec, "label": el, "detail": ed, "file": src_e},
                    "costs": {"kind": "qbo", "label": "QuickBooks", "detail": f"project P&L, synced {M['report_date']}"},
                    "billed": {"kind": "qbo", "label": "QuickBooks", "detail": f"invoices on the project, synced {M['report_date']}"}},
            "jt": jtb,
            "rp_status": r.get("status"), "rp_action": r.get("action"), "band": r.get("band"),
            "schedule": {"days": len(days), "first": first.isoformat() if first else None,
                         "last": last.isoformat() if last else None, "runs": _runs(days)},
            "last_seen": last.isoformat() if last else None, "stale": stale,
            "flags": flags, "problem": bool(flags) or not K or not E,
            "pics": {k: v for k, v in pics.items() if v},
        })
        if i % 10 == 0:
            log(f"  {i}/{len(lines)} lines · {time.time() - t0:.0f}s")

    finished = []
    for x in F["removed"]:
        job = str(x.get("JOB #") or "").upper()
        base = job.replace("-FTW", "")
        last_day = x.get("LAST DAY ON SCHEDULE")
        ld = None
        if isinstance(last_day, (dt.date, dt.datetime)):
            ld = last_day if isinstance(last_day, dt.date) else last_day.date()
        else:
            m = re.match(r"(\d{2})/(\d{2})/(\d{4})", str(last_day or ""))
            if m:
                ld = dt.date(int(m.group(3)), int(m.group(1)), int(m.group(2)))
        days = _timeline(job, sched_files, index_files) if sched_files else []
        pics = {}
        if not no_snapshots:
            jd = IMG_DIR / job
            pics["removed"] = _row_strip_png(fp, "Removed log", x["_row"], [2, 3, 5, 6, 7, 8, 9, 10, 11], jd / "removed.png", 1)
            for tag, day in (("first", dt.date.fromisoformat(days[0]["date"]) if days else None), ("last", ld or (dt.date.fromisoformat(days[-1]["date"]) if days else None))):
                f = sched_files.get(day) if day else None
                if not f:
                    continue
                hit = _sched_row_of(f, base, job.endswith("-FTW"))
                if hit:
                    name, hdr, row = hit
                    pics["sched_" + tag] = _row_strip_png(f, name, row, list(range(2, 8)), jd / f"sched_{tag}.png", hdr)
                    if pics["sched_" + tag]:
                        pics["sched_" + tag]["date"] = day.isoformat()
        finished.append({
            "line": job, "ftw": job.endswith("-FTW"), "name": x.get("ADDRESS"), "builder": x.get("BUILDER"),
            "contract": _num(x.get("CONTRACT $")), "etc": _num(x.get("ETC $")),
            "billed": _num(x.get("BILLED TO DATE $ (QBO)")), "costs": _num(x.get("COSTS TO DATE $ (QBO)")),
            "last_invoice": _fmt(x.get("LAST INVOICE")), "last_day": _fmt(last_day), "last_stage": x.get("LAST STAGE"),
            "why": x.get("WHY REMOVED"), "source": x.get("SOURCE"), "removed_on": _fmt(x.get("REMOVED ON")),
            "approved_by": x.get("APPROVED BY"),
            "jt": jt_block(base, job),
            "schedule": {"days": len(days), "first": days[0]["date"] if days else None, "last": days[-1]["date"] if days else None, "runs": _runs(days)},
            "pics": {k: v for k, v in pics.items() if v},
        })

    cache.save()
    kinds = {}
    for c in current:
        kinds[c["src"]["contract"]["kind"]] = kinds.get(c["src"]["contract"]["kind"], 0) + 1
    out = {
        "built_at": dt.datetime.now().isoformat(timespec="seconds"),
        "master": {"file": mp.name, "tab": M["tab"], "renamed": M["renamed"], "report_date": M["report_date"],
                   "mtime": dt.datetime.fromtimestamp(mp.stat().st_mtime).isoformat(timespec="seconds")},
        "rp_file": {"file": fp.name, "mtime": dt.datetime.fromtimestamp(fp.stat().st_mtime).isoformat(timespec="seconds")},
        "jobtread": {"queried": not no_jobtread, "approved": sum(1 for v in jt.values() if v["docs"]), "exists": len(jt)},
        "schedule_mounted": bool(sched_files),
        "newest_schedule": newest.isoformat() if newest else None,
        "counts": {"current": len(current), "finished": len(finished), "contract_kinds": kinds,
                   "problems": sum(1 for c in current if c["problem"]),
                   "in_jobtread": sum(1 for c in current if c["jt"]["count"]),
                   "jt_exists": sum(1 for c in current if c["jt"]["exists"]),
                   "typed": sum(1 for c in current if c["src"]["contract"]["kind"] == "typed" or c["src"]["etc"]["kind"] == "typed"),
                   "stale": sum(1 for c in current if c["stale"])},
        "current": current, "finished": finished,
        "log": _log_lines[-40:],
    }
    (OUT_DIR / "rp_review.json").write_text(json.dumps(out, default=str), encoding="utf-8")
    log(f"wrote {OUT_DIR / 'rp_review.json'} in {time.time() - t0:.0f}s (cache hits {cache.hits})")
    return out


def load() -> Optional[dict]:
    p = OUT_DIR / "rp_review.json"
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None


# ───────────────────────────── marks (the ledger's write surface) ─────────────────────────────
# The owner and the ops manager never edit the prepared numbers (2026-09-15: "we need our own
# space for our own numbers with the numbers you have prepopulated in case it's right, we check
# or mark x with notes/changes and you read and make sense of it"). Each line keeps ONE
# standing answer (rp_review_mark) plus every answer ever given (rp_review_mark_log):
#   decision   current: confirmed | fix        finished: agree | keep       '' clears
#   mode       ops (the ops manager is here, deciding together) | me (the owner alone)
#   our_contract / our_etc   THEIR numbers (prepopulated with the prepared value)
#   contract_ok / etc_ok     1 = check, 0 = X, NULL = not answered
#   note       free text - what could not be settled, what changed
MODES = ("ops", "me")
DECISIONS = {"current": ("confirmed", "fix", "noted"), "finished": ("agree", "keep", "noted")}
_COLS = "project_no, kind, decision, mode, note, at, our_contract, our_etc, contract_ok, etc_ok"


def _ensure(con: sqlite3.Connection) -> None:
    con.execute("CREATE TABLE IF NOT EXISTS rp_review_mark ("
                "project_no TEXT NOT NULL, kind TEXT NOT NULL, decision TEXT NOT NULL, mode TEXT NOT NULL, "
                "note TEXT, at TEXT NOT NULL, our_contract REAL, our_etc REAL, contract_ok INTEGER, etc_ok INTEGER, "
                "PRIMARY KEY (project_no, kind))")
    con.execute("CREATE TABLE IF NOT EXISTS rp_review_mark_log ("
                "id INTEGER PRIMARY KEY AUTOINCREMENT, project_no TEXT NOT NULL, kind TEXT NOT NULL, "
                "decision TEXT NOT NULL, mode TEXT NOT NULL, note TEXT, at TEXT NOT NULL, "
                "our_contract REAL, our_etc REAL, contract_ok INTEGER, etc_ok INTEGER)")
    for t in ("rp_review_mark", "rp_review_mark_log"):        # older tables: add the answer columns
        have = {r[1] for r in con.execute(f"PRAGMA table_info({t})")}
        for c, typ in (("our_contract", "REAL"), ("our_etc", "REAL"), ("contract_ok", "INTEGER"), ("etc_ok", "INTEGER")):
            if c not in have:
                con.execute(f"ALTER TABLE {t} ADD COLUMN {c} {typ}")


def _rowdict(r) -> dict:
    return {"project_no": r[0], "kind": r[1], "decision": r[2], "mode": r[3], "note": r[4], "at": r[5],
            "our_contract": r[6], "our_etc": r[7], "contract_ok": r[8], "etc_ok": r[9]}


def read_marks() -> dict:
    """{'<project_no>|<kind>': {...}} - absent-safe (no DB / no table -> {})."""
    if not LEDGER_DB.exists():
        return {}
    try:
        con = sqlite3.connect(f"file:{LEDGER_DB}?mode=ro", uri=True)
        try:
            rows = con.execute(f"SELECT {_COLS} FROM rp_review_mark").fetchall()
        finally:
            con.close()
    except sqlite3.OperationalError:
        return {}
    return {f"{r[0]}|{r[1]}": _rowdict(r) for r in rows}


def read_mark_log(limit: int = 400) -> list:
    if not LEDGER_DB.exists():
        return []
    try:
        con = sqlite3.connect(f"file:{LEDGER_DB}?mode=ro", uri=True)
        try:
            rows = con.execute(f"SELECT {_COLS} FROM rp_review_mark_log ORDER BY id DESC LIMIT ?", (limit,)).fetchall()
        finally:
            con.close()
    except sqlite3.OperationalError:
        return []
    return [_rowdict(r) for r in rows]


def set_mark(project_no: str, kind: str, decision: str, mode: str, note: str, now: str,
             our_contract=None, our_etc=None, contract_ok=None, etc_ok=None) -> None:
    """Write one answer (and append it to the log). decision '' clears the standing mark."""
    LEDGER_DB.parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(str(LEDGER_DB))
    try:
        _ensure(con)
        vals = (project_no, kind, decision, mode, note, now, our_contract, our_etc, contract_ok, etc_ok)
        if decision:
            con.execute(f"INSERT INTO rp_review_mark ({_COLS}) VALUES (?,?,?,?,?,?,?,?,?,?) "
                        "ON CONFLICT(project_no, kind) DO UPDATE SET decision=excluded.decision, mode=excluded.mode, "
                        "note=excluded.note, at=excluded.at, our_contract=excluded.our_contract, our_etc=excluded.our_etc, "
                        "contract_ok=excluded.contract_ok, etc_ok=excluded.etc_ok", vals)
        else:
            con.execute("DELETE FROM rp_review_mark WHERE project_no=? AND kind=?", (project_no, kind))
        con.execute(f"INSERT INTO rp_review_mark_log ({_COLS}) VALUES (?,?,?,?,?,?,?,?,?,?)",
                    (project_no, kind, decision or "cleared", mode, note, now, our_contract, our_etc, contract_ok, etc_ok))
        con.commit()
    finally:
        con.close()


def answers_report() -> str:
    """The answers as text - what a later session reads to 'make sense of it': every line the
    owner / ops manager marked, their numbers where they differ from the prepared ones, the X's,
    the notes. Prepared values come from the last build."""
    data = load() or {}
    prepared = {c["line"]: c for c in data.get("current", [])}
    prepared.update({f["line"] + "|finished": f for f in data.get("finished", [])})
    marks = read_marks()
    if not marks:
        return "no answers yet"
    out = [f"RP review answers - {len(marks)} line(s) marked (build {data.get('built_at', '?')})", ""]
    for key in sorted(marks):
        m = marks[key]; pn, kind = key.split("|")
        p = prepared.get(pn if kind == "current" else pn + "|finished", {})
        who = "OPS MANAGER + owner" if m["mode"] == "ops" else "owner alone"
        bits = [f"{pn} [{kind}] {m['decision'].upper()} by {who} on {m['at'][:16].replace('T', ' ')}"]
        for fld, ok, ours in (("contract", m["contract_ok"], m["our_contract"]), ("etc", m["etc_ok"], m["our_etc"])):
            mine = p.get(fld)
            mark = {1: "check", 0: "X"}.get(ok, "-")
            diff = f" THEIR {fld.upper()} {ours:,.2f} vs prepared {mine:,.2f}" if (ours is not None and mine is not None and abs(ours - mine) > 0.5) else ""
            if ok is not None or diff:
                bits.append(f"    {fld}: {mark}{diff}")
        if m["note"]:
            bits.append(f"    note: {m['note']}")
        out.extend(bits)
    return "\n".join(out)


def finalize_package() -> dict:
    """What the saved answers mean for the WIP report, as a package the guarded WIP writer can
    apply (rp-review/finalize.json + decisions_rp.json in the --apply-review shape). Nothing is
    written to the master here - the dashboard never touches the workbook (repo rule)."""
    data = load() or {}
    marks = read_marks()
    cur = {c["line"]: c for c in data.get("current", [])}
    fin = {f["line"]: f for f in data.get("finished", [])}
    changes, confirmed, fixes, done, keep, notes = [], [], [], [], [], []
    fields = {}
    for key, m in marks.items():
        pn, kind = key.split("|")
        if kind == "current" and pn in cur:
            c = cur[pn]
            if m["decision"] == "confirmed":
                confirmed.append(pn)
            elif m["decision"] == "fix":
                fixes.append({"line": pn, "note": m["note"] or ""})
            for fld, ours, prepared, dkey in (("contract", m["our_contract"], c.get("contract"), "orig_contract"),
                                              ("etc", m["our_etc"], c.get("etc"), "orig_etc")):
                if ours is not None and (prepared is None or abs(ours - prepared) > 0.5):
                    changes.append({"line": pn, "field": fld, "from": prepared, "to": ours, "who": m["mode"], "at": m["at"], "note": m["note"] or ""})
                    fields.setdefault(pn, {})[dkey] = {"approved": False, "revert": ours}
        elif kind == "finished" and pn in fin:
            (done if m["decision"] == "agree" else keep if m["decision"] == "keep" else []).append(pn)
        if m["note"]:
            notes.append({"line": pn, "kind": kind, "note": m["note"], "who": m["mode"], "at": m["at"]})
    pkg = {"generated": dt.datetime.now().isoformat(timespec="seconds"), "build": data.get("built_at"),
           "answered": len(marks), "confirmed": sorted(confirmed), "fixes": fixes, "changes": changes,
           "finished_agreed": sorted(done), "finished_keep": sorted(keep), "notes": notes,
           "master_tab": (data.get("master") or {}).get("tab"), "master_renamed": (data.get("master") or {}).get("renamed")}
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    (OUT_DIR / "finalize.json").write_text(json.dumps(pkg, indent=1, default=str), encoding="utf-8")
    (OUT_DIR / "decisions_rp.json").write_text(json.dumps({"fields": fields, "drop_added": []}, indent=1), encoding="utf-8")
    (OUT_DIR / "answers.txt").write_text(answers_report(), encoding="utf-8")
    return pkg


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--no-jobtread", action="store_true", help="skip the JobTread pull")
    ap.add_argument("--no-snapshots", action="store_true", help="skip the source snapshots (fast)")
    ap.add_argument("--limit", type=int, default=0, help="first N lines only (testing)")
    ap.add_argument("--answers", action="store_true", help="print the owner's / ops manager's answers and stop")
    ap.add_argument("--finalize", action="store_true", help="write the finalize package (finalize.json + decisions_rp.json + answers.txt) and stop")
    a = ap.parse_args()
    if a.answers:
        print(answers_report())
        return 0
    if a.finalize:
        pkg = finalize_package()
        print(json.dumps({k: (v if not isinstance(v, list) else len(v)) for k, v in pkg.items()}, indent=1))
        return 0
    build(a.no_jobtread, a.no_snapshots, a.limit)
    return 0


if __name__ == "__main__":
    sys.exit(main())
