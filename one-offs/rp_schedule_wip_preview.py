#!/usr/bin/env python3
"""
rp_schedule_wip_preview.py — PREVIEW of the schedule-driven RP WIP method
(the user 2026-07-16). READ-ONLY everywhere; never touches the WIP master.

THE METHOD BEING PREVIEWED
  The daily schedule's 'Main Schedule' tab is the truth for which RP jobs are
  ACTIVE — the General List lags it (e.g. RP6440's fence scope is on the
  schedule but has no GL row). Under the new method each schedule job gets:
    • CONTRACT  = the bid proposal PDF in the project folder (last
      "SUB TOTAL: $…" in the text — the signed client number)
    • ETC       = the takeoff workbook's cost sheet ('JobTread Cost Gral',
      the last sheet): slab = SL+PR sections, flatwork = FW section
  Scope: FLATWORK schedule section → RP#-FTW line; wreck rows follow their
  description; everything else (trench/piers/forms/grade) → slab line.

OUTPUT — one Excel (never the WIP): ~/Downloads/RP WIP - Schedule Method
Preview.xlsx, grouped NEW (on schedule, missing from the current WIP) /
CHANGED (numbers differ from the General List) / MATCHES, with the source
file of every number.

Usage:
  python3 rp_schedule_wip_preview.py
  python3 rp_schedule_wip_preview.py --project RP6440
"""
from __future__ import annotations

import argparse
import os
import re
import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_REPO))
sys.path.insert(0, str(_REPO / "wip"))

from openpyxl import Workbook, load_workbook

import rp_wip_reader as RP

MAIN_SHEET = "Main Schedule"
# Section bands on the Main Schedule tab → scope of the WIP line they feed.
SECTION_SCOPE = [
    ("WRECK", "desc"),          # wreck slab → slab · wreck flatwork → ftw
    ("FLATWORK", "ftw"),
    ("GRADE AND BACKOUT", "slab"),
    ("TRENCH", "slab"),
    ("PIERS", "slab"),
    ("FORM SET", "slab"),
    ("TRACTOR", "slab"),
    ("CONCRETE CUTTING", "desc"),
]
_JOB_RE = re.compile(r"^(RP|CP)\d{3,4}$", re.IGNORECASE)
_SUBTOTAL_RE = re.compile(r"(?:SUB\s*)?TOTAL\s*:?\s*\$\s*([\d,]+(?:\.\d{1,2})?)",
                          re.IGNORECASE)
# Scope words that mark a proposal/takeoff as a SIDE scope (not the base slab)
_SIDE_TOKENS = {"POOL", "CASITA", "CABANA", "FENCE", "CAPS", "FLATWORK",
                "RETAINING", "COURTYARD", "WALL", "PATIO", "DRIVEWAY",
                "FOOTINGS", "FOOTER", "PAVING"}


def _norm(s) -> str:
    return re.sub(r"\s+", " ", str(s or "").upper().strip())


def latest_schedule(sched_dir: Path):
    best = None
    for year_dir in sched_dir.iterdir():
        if not (year_dir.is_dir() and year_dir.name.isdigit()):
            continue
        for month_dir in year_dir.iterdir():
            if not month_dir.is_dir():
                continue
            for f in month_dir.iterdir():
                m = RP._SCHED_FILE_RE.search(f.name)
                if m:
                    mo, dy, yy = (int(g) for g in m.groups())
                    key = (2000 + yy, mo, dy)
                    if best is None or key > best[0]:
                        best = (key, f)
    return best


def read_main_schedule(path: Path):
    """'Main Schedule' → [{job, scope, section, address, builder, desc}].
    Section bands tracked from the band-title rows; a data row is any row
    whose PROJECT column matches RP####/CP####."""
    wb = load_workbook(path, data_only=True, read_only=True)
    sheet = next((s for s in wb.sheetnames
                  if s.strip().lower() == MAIN_SHEET.lower()), None)
    if sheet is None:
        wb.close()
        raise SystemExit(f"No {MAIN_SHEET!r} tab in {path.name}")
    ws = wb[sheet]
    out, seen, section = [], {}, None
    for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 400)):
        vals = [str(c.value).strip() if c.value is not None else ""
                for c in row[:12]]
        joined = _norm(" ".join(vals))
        job = vals[1].upper() if len(vals) > 1 else ""
        if not _JOB_RE.match(job):
            for token, scope in SECTION_SCOPE:
                if token in joined and "NAME" not in joined \
                        and "SUPERINTENDENT" not in joined:
                    section = (token, scope)
                    break
            continue
        if section is None:
            continue
        desc = _norm(vals[5])
        token, scope = section
        if scope == "desc":
            scope = "ftw" if ("FLATWORK" in desc or "FTW" in desc
                              or "PAVING" in desc) else "slab"
        key = (job, scope)
        rec = seen.get(key)
        if rec:
            if desc and desc not in rec["desc"]:
                rec["desc"] += f" · {desc}"
            continue
        rec = {"job": job, "scope": scope, "section": token,
               "address": _norm(vals[2]), "city": _norm(vals[3]),
               "builder": _norm(vals[4]), "desc": desc}
        seen[key] = rec
        out.append(rec)
    wb.close()
    return out


def read_gl_jobs(path: Path):
    """ALL job numbers present in the General List (priced or not) — presence
    check for the lag report. RP.read_general_list drops unpriced rows, so
    presence needs its own pass."""
    wb = load_workbook(path, data_only=True, read_only=True)
    jobs = set()
    for sheet in (RP.ALPHA_SHEET, RP.SMALL_SHEET):
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        for row in ws.iter_rows(min_row=6, min_col=3, max_col=3):
            v = row[0].value
            if v:
                m = RP._JOB_RE.match(str(v).strip())
                if m:
                    jobs.add(m.group(1).upper())
    wb.close()
    return jobs


# ─────────────────── proposal PDF → contract ───────────────────────
def _desc_tokens(desc: str):
    return {w for w in re.split(r"[^A-Z]+", desc or "")
            if len(w) > 3 and w not in ("POUR", "READY", "MAKE")}


def _score_name(name_u: str, scope: str, desc: str) -> float:
    """Rank a proposal/takeoff filename for the (scope, description) asked."""
    toks = _desc_tokens(desc)
    side_in_name = {t for t in _SIDE_TOKENS if t in name_u}
    score = 0.0
    score += 2.0 * len({t for t in toks if t in name_u})       # desc words hit
    if scope == "ftw":
        score += 3.0 if ("FLATWORK" in name_u or "FTW" in name_u) else 0.0
    else:
        # Slab wants the BASE proposal — side-scope names only help when the
        # schedule description itself asked for them.
        stray = side_in_name - toks
        score -= 2.0 * len(stray)
    if "REVISED" in name_u or "UPDATED" in name_u:
        score += 1.0
    return score


def find_proposal(folder: Path, scope: str, desc: str):
    """Best bid-proposal PDF + its extracted SUB TOTAL. Returns
    (path, amount, note) — amount None when no priced proposal reads.

    FTW scope only accepts PDFs whose NAME says flatwork (or matches the
    schedule description) — the base proposal is the SLAB contract, and
    silently returning it would overstate the -FTW line (RP5542: $397K
    house bid vs the real flatwork scope). No such PDF → caller falls back
    to the takeoff's bid sheets."""
    import pdfplumber
    cands = []
    try:
        for f in folder.iterdir():
            if f.suffix.lower() != ".pdf":
                continue
            n = _norm(f.name)
            if "INVOICE" in n or "DIAGRAM" in n or "PLAN" in n:
                continue
            if "BID" in n or "PROPOSAL" in n or _desc_tokens(desc) & set(n.split()):
                if scope == "ftw" and not (
                        "FLATWORK" in n or "FTW" in n
                        or _desc_tokens(desc) & set(n.split())):
                    continue
                cands.append(f)
    except OSError:
        return None, None, "folder unreadable"
    if not cands:
        return None, None, ("no flatwork proposal PDF" if scope == "ftw"
                            else "no proposal PDF in folder")
    cands.sort(key=lambda f: (_score_name(_norm(f.name), scope, desc),
                              f.stat().st_mtime), reverse=True)
    for f in cands[:4]:                    # best-ranked few; first priced wins
        try:
            with pdfplumber.open(f) as pdf:
                txt = "\n".join(pg.extract_text() or "" for pg in pdf.pages)
            hits = _SUBTOTAL_RE.findall(txt)
            if hits:
                return f, float(hits[-1].replace(",", "")), None
        except Exception:
            continue
    return cands[0], None, "proposal found but no SUB TOTAL text (unpriced?)"


# ─────────────────── takeoff last sheet → ETC ──────────────────────
def _cost_sheet_totals(ws):
    """Parse 'JobTread Cost Gral' in BOTH layouts:
      A) code rows (SL#/PR#/FW# in col A, qty col C, unit cost col D)
      B) banded rows (SLAB/PIERS/FLATWORK in col A, items col B)

    THE NUMBER THAT COUNTS is the template's OWN subtotal cell at the foot
    of each band — NOT Σ(qty × cost) of the visible items. RP5542 FLATWORK:
    items sum to $47,758 but D33 (the blue subtotal the clerk keys into the
    General List) is $88,858 — its formula pulls from the Flatwork takeoff
    sheet beyond these rows (the user 2026-07-17). Items are kept only as
    the fallback when the subtotal cell errors (#N/A) + a mismatch check.

    Returns {'SL'|'PR'|'FW': {'sub': float|None, 'cell': 'D33'|None,
                              'items': float}}."""
    bands = {k: {"sub": None, "cell": None, "items": 0.0}
             for k in ("SL", "PR", "FW")}
    band = None
    for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 200)):
        a = _norm(row[0].value if len(row) > 0 else "")
        b = _norm(row[1].value if len(row) > 1 else "")
        qty = row[2].value if len(row) > 2 else None
        cost_cell = row[3] if len(row) > 3 else None
        cost = cost_cell.value if cost_cell is not None else None
        m = re.match(r"^(SL|PR|FW)\d", a)
        if m:                                           # layout A code row
            band = m.group(1)
            try:
                bands[band]["items"] += float(qty or 0) * float(cost or 0)
            except (TypeError, ValueError):
                pass
            continue
        if a.startswith("SLAB"):
            band = "SL"
        elif a.startswith("PIER"):
            band = "PR"
        elif a.startswith("FLATWORK"):
            band = "FW"
        if band and b and b not in ("DESCRIPTION",):    # layout B item row
            try:
                bands[band]["items"] += float(qty or 0) * float(cost or 0)
            except (TypeError, ValueError):
                pass
            continue
        # Band-foot subtotal: no code, no item label, a value in the COST
        # column — the template's own total (first one after the band wins).
        if band and not a and not b and cost is not None \
                and bands[band]["sub"] is None and bands[band]["cell"] is None:
            bands[band]["cell"] = cost_cell.coordinate
            if isinstance(cost, (int, float)):
                bands[band]["sub"] = float(cost)
            # strings like '#N/A' leave sub=None → items fallback
    return bands


def find_takeoff_etc(folder: Path, job: str, scope: str, desc: str):
    """Best takeoff .xlsm for the scope → ETC from its cost sheet's own
    subtotal cells. Returns (path, etc, note, fragment) — note names the
    sheet + cells the number came from; fragment jump-links there."""
    cands = []
    try:
        for f in folder.iterdir():
            if f.suffix.lower() not in (".xlsm", ".xlsx"):
                continue
            n = _norm(f.name)
            if n.startswith(job) or job in n:
                cands.append(f)
    except OSError:
        return None, None, "folder unreadable", None
    if not cands:
        return None, None, "no takeoff in folder", None
    cands.sort(key=lambda f: (_score_name(_norm(f.name), scope, desc),
                              f.stat().st_mtime), reverse=True)
    for f in cands[:4]:
        try:
            wb = load_workbook(f, data_only=True, read_only=True)
        except Exception:
            continue
        sheet = next((s for s in wb.sheetnames if "COST GRAL" in _norm(s)), None)
        if sheet is None:
            wb.close()
            continue
        bands = _cost_sheet_totals(wb[sheet])
        wb.close()
        # A SIDE-SCOPE takeoff (its name matches the schedule description —
        # fence, pool house, caps…) is a whole little job of its own: its
        # entire cost sheet IS the scope's ETC. Only the BASE takeoff splits
        # slab (SL+PR) vs flatwork (FW) bands.
        if _desc_tokens(desc) & set(_norm(f.name).split()):
            keys = ("SL", "PR", "FW")
        else:
            keys = ("FW",) if scope == "ftw" else ("SL", "PR")
        etc, cells, notes = 0.0, [], []
        for k in keys:
            b = bands[k]
            if b["sub"] is not None:
                etc += b["sub"]
                cells.append(b["cell"])
            elif b["items"]:
                etc += b["items"]
                cells.append(f"{k} items (subtotal cell "
                             f"{b['cell'] or 'missing'} unreadable)")
            if (b["sub"] is not None and b["items"]
                    and abs(b["sub"] - b["items"]) > 1):
                notes.append(f"{k} subtotal ${b['sub']:,.0f} ≠ item rows "
                             f"${b['items']:,.0f} — template pulls extra scope")
        if etc:
            src = f"'{sheet}' {' + '.join(cells)}"
            note = "; ".join([src] + notes)
            frag = f"#'{sheet}'!{next((c for c in cells if re.match(r'^[A-Z]+[0-9]+$', c)), 'A1')}"
            return f, round(etc, 2), note, frag
    return cands[0], None, "takeoff has no cost sheet (older template?)", None


# ─────────────────────────── report ────────────────────────────────
# Implied gross-margin sanity band for the NEW numbers (the user 2026-07-17,
# RP5542-FTW: contract $91K vs takeoff ETC $48K = 47.6% GP — "ETC way too
# low"). RP jobs run ~10–25% GP; outside the band the pair is mismatched
# (wrong scope file, stale takeoff, or missing cost sections).
_GP_HI = 0.35     # above → ETC too low vs the contract
_GP_LO = 0.05     # below → margin too thin (ETC too high / contract too low)


def margin_flag(contract, etc):
    """(gp_pct, flag_text) for a contract/ETC pair — flag None when sane."""
    if not contract or etc is None:
        return None, None
    gp = (contract - etc) / contract
    if gp > _GP_HI:
        return gp, (f"ETC WAY TOO LOW — implied GP {gp * 100:.0f}% "
                    f"(RP runs ~10–25%): scope mismatch or stale takeoff")
    if gp < 0:
        return gp, f"ETC EXCEEDS contract (GP {gp * 100:.0f}%) — check the pair"
    if gp < _GP_LO:
        return gp, (f"margin too thin — implied GP {gp * 100:.0f}%: "
                    f"ETC too high or contract too low")
    return gp, None


def _needs_rich(needs: str):
    """NEEDS as color-coded rich text (the user 2026-07-21): BLUE = cost/ETC
    issue, ORANGE = contract/bid-proposal issue, red = everything else
    (missing from the list, no folder…). One cell, per-segment colors."""
    if not needs:
        return None
    from openpyxl.cell.rich_text import CellRichText, TextBlock
    from openpyxl.cell.text import InlineFont
    BLUE = InlineFont(color="0070C0", b=True)
    ORANGE = InlineFont(color="ED7D31", b=True)
    RED = InlineFont(color="9C0006", b=True)
    parts = [p.strip() for p in needs.split(";") if p.strip()]
    blocks = []
    for i, p in enumerate(parts):
        u = p.upper()
        if "ETC" in u or "COST" in u or "TAKEOFF" in u or "MARGIN" in u:
            f = BLUE
        elif "PROPOSAL" in u or "CONTRACT" in u or "BID" in u:
            f = ORANGE
        else:
            f = RED
        blocks.append(TextBlock(f, p + ("; " if i < len(parts) - 1 else "")))
    return CellRichText(*blocks)


def _link(cell, target, fragment: str = ""):
    """file:// hyperlink + blue underline; no-op when target is None."""
    from openpyxl.styles import Font
    if target is None or cell.value in (None, ""):
        return
    try:
        cell.hyperlink = Path(target).as_uri() + fragment
    except (ValueError, OSError):
        return
    cell.font = Font(color="0563C1", underline="single")


def write_report(items, sched_label, out_path: Path) -> None:
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    GRAY = PatternFill("solid", fgColor="D9D9D9")
    YELLOW = PatternFill("solid", fgColor="FFF2CC")   # General List numbers
    GREEN = PatternFill("solid", fgColor="E2EFDA")    # proposal/takeoff numbers
    AMBER = PatternFill("solid", fgColor="FCE4D6")    # significant deltas
    NEW_F = PatternFill("solid", fgColor="F8CBAD")    # NEW group band
    BAD = Font(color="9C0006", bold=True)
    GOOD = Font(color="006100", bold=True)
    thin = Side(style="thin", color="000000")
    BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
    CUR = '"$"#,##0.00_);[Red]("$"#,##0.00)'

    lock = out_path.with_name("~$" + out_path.name)
    if lock.exists():
        raise SystemExit(f"{out_path.name} is open in Excel — close it first")
    wb = Workbook()
    ws = wb.active
    ws.title = "PREVIEW"
    ws["A1"] = (f"SCHEDULE-DRIVEN RP WIP — PREVIEW ONLY (schedule "
                f"{sched_label}; the WIP master was NOT touched). "
                f"YELLOW = General List numbers · GREEN = bid proposal / "
                f"takeoff numbers · NEEDS: BLUE = cost/ETC issue, ORANGE = "
                f"contract/proposal issue · every number links to its source.")
    ws["A1"].font = Font(bold=True)
    ws.append([])
    HDR = ["GROUP", "WIP LINE", "WORK DESC",
           "ADDRESS", "BUILDER", "IN GEN. LIST?", "GL CONTRACT $", "GL ETC $",
           "NEW CONTRACT $ (proposal)", "NEW ETC $ (takeoff)", "NEW GP %",
           "Δ CONTRACT", "Δ ETC", "PROPOSAL PDF", "TAKEOFF FILE", "NEEDS"]
    ws.append(HDR)
    for c in range(1, len(HDR) + 1):
        cell = ws.cell(3, c)
        cell.font = Font(bold=True)
        cell.fill = GRAY
        cell.alignment = Alignment(horizontal="center", vertical="center",
                                   wrap_text=True)
        cell.border = BORDER
    # Header tint mirrors the data: GL columns yellow, source columns green.
    for c in (7, 8):
        ws.cell(3, c).fill = YELLOW
    for c in (9, 10, 11):
        ws.cell(3, c).fill = GREEN

    order = {"NEW — not in WIP": 0, "CHANGED": 1, "MATCHES": 2}
    for it in sorted(items, key=lambda x: (order.get(x["group"], 3),
                                           x["line"])):
        d_k = (it["new_contract"] - it["gl_contract"]
               if it["new_contract"] is not None and it["gl_contract"] else None)
        d_e = (it["new_etc"] - it["gl_etc"]
               if it["new_etc"] is not None and it["gl_etc"] else None)
        gp, gp_flag = margin_flag(it["new_contract"], it["new_etc"])
        needs = "; ".join(x for x in (it["needs"], gp_flag) if x)
        ws.append([it["group"], it["line"], it["desc"],
                   it["address"], it["builder"],
                   ("yes" if it["in_gl"] else "NO"),
                   it["gl_contract"], it["gl_etc"],
                   it["new_contract"], it["new_etc"], gp,
                   d_k, d_e,
                   (it["proposal"].name if it["proposal"] else it["p_note"]),
                   ((it["takeoff"].name + (f"  [{it['t_note']}]" if it["t_note"] else ""))
                    if it["takeoff"] else it["t_note"]),
                   _needs_rich(needs)])
        r = ws.max_row
        for cc in range(1, len(HDR) + 1):
            ws.cell(r, cc).border = BORDER
            ws.cell(r, cc).alignment = Alignment(
                vertical="top", wrap_text=(cc in (3, 14, 15, 16)))
        for cc in (7, 8, 9, 10, 12, 13):
            ws.cell(r, cc).number_format = CUR
        ws.cell(r, 11).number_format = "0.0%"
        # Source-colored numbers: GL yellow · proposal/takeoff green.
        for cc in (7, 8):
            ws.cell(r, cc).fill = YELLOW
        for cc in (9, 10, 11):
            ws.cell(r, cc).fill = GREEN
        # Links: line → folder · GL $ → its General List row · new numbers +
        # file columns → the exact proposal PDF / takeoff workbook.
        _link(ws.cell(r, 2), it.get("folder"))
        gl_frag = (f"#'{it['gl_sheet']}'!C{it['gl_row']}"
                   if it.get("gl_row") else "")
        for cc in (7, 8):
            _link(ws.cell(r, cc), RP.ALPHA_PATH if it.get("gl_row") else None,
                  gl_frag)
        # NEW CONTRACT $ opens the PDF itself; the PROPOSAL PDF column
        # opens the FOLDER so the PDF can be grabbed/attached (the user
        # 2026-07-17 — a plain file link can't make Finder pre-select the
        # file, so the cell text carries the exact filename to look for).
        _link(ws.cell(r, 9), it.get("proposal"))
        _link(ws.cell(r, 14),
              it["proposal"].parent if it.get("proposal") else None)
        for cc in (10, 15):
            # Jump to the exact subtotal cell the ETC came from (the user
            # 2026-07-17: "list the sheet and cell and take me there").
            _link(ws.cell(r, cc), it.get("takeoff"), it.get("t_frag") or "")
        # Distinct groups: NEW rows banded orange in col A; big deltas amber
        # + bold red so a changed number can't be missed.
        if it["group"].startswith("NEW"):
            ws.cell(r, 1).fill = NEW_F
            ws.cell(r, 1).font = BAD
        elif it["group"] == "CHANGED":
            ws.cell(r, 1).font = Font(color="BF6000", bold=True)
        else:
            ws.cell(r, 1).font = GOOD
        for cc, dv, base in ((12, d_k, it["gl_contract"]),
                             (13, d_e, it["gl_etc"])):
            if dv is not None and base and abs(dv) > max(100, 0.01 * base):
                ws.cell(r, cc).fill = AMBER
                ws.cell(r, cc).font = BAD
        if gp is not None and gp_flag:
            ws.cell(r, 11).font = BAD
        if not it["in_gl"]:
            ws.cell(r, 6).font = BAD
    widths = (16, 12, 28, 22, 20, 9, 13, 13, 14, 13, 9, 12, 12, 30, 30, 42)
    for col, w in zip("ABCDEFGHIJKLMNOP", widths):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A4"
    wb.save(out_path)
    print(f"  ✓ Preview → {out_path}")


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--project", help="filter to one job #")
    ap.add_argument("--schedule", help="use this schedule file instead of the latest")
    args = ap.parse_args()

    print("\n  RP WIP — schedule-driven method PREVIEW (read-only)")
    print("  " + "─" * 62)
    if args.schedule:
        sched_path = Path(args.schedule)
        if not sched_path.exists():
            print(f"  ✗ schedule not found: {sched_path}")
            return 1
        m = RP._SCHED_FILE_RE.search(sched_path.name)
        label = "-".join(m.groups()) if m else sched_path.stem
    else:
        best = latest_schedule(RP.SCHEDULE_DIR)
        if best is None:
            print("  ✗ no schedule file found")
            return 1
        (_k, sched_path) = best
        label = f"{_k[1]}-{_k[2]}-{_k[0] % 100:02d}"
    print(f"  schedule: {sched_path.name}")
    sched = read_main_schedule(sched_path)
    if args.project:
        sched = [s for s in sched if s["job"] == args.project.upper()]
    print(f"  active schedule lines: {len(sched)}")

    gl_jobs = read_gl_jobs(RP.ALPHA_PATH)
    records, _m = RP.read_general_list(RP.ALPHA_PATH)
    gl_by_job = {r["job"]: r for r in records}

    rp_to_folders, addr_folders = RP.index_residential(RP.RP_ROOT)

    items = []
    for s in sched:
        job, scope = s["job"], s["scope"]
        line = job if (scope == "slab" or job.startswith("CP")) else f"{job}-FTW"
        rec = gl_by_job.get(job)
        if job.startswith("CP"):
            gl_k = ((rec["slab_bid"] or 0) + (rec["flat_bid"] or 0)) or None \
                if rec else None
            gl_e = ((rec["slab_cost"] or 0) + (rec["flat_cost"] or 0)) or None \
                if rec else None
        elif scope == "ftw":
            gl_k = rec and rec["flat_bid"]
            gl_e = rec and rec["flat_cost"]
            if rec and rec.get("flat_other"):
                gl_k = gl_e = None            # OTHER won it — no line today
        else:
            gl_k = rec and rec["slab_bid"]
            gl_e = rec and rec["slab_cost"]
        in_wip = bool(gl_k or gl_e)

        folders = sorted(rp_to_folders.get(job, ()),
                         key=lambda f: (f.parent.name, f.name))
        folder = folders[0] if folders else None
        if folder is None and s["address"]:
            parts = s["address"].split(None, 1)
            fake = {"house": parts[0] if parts else "",
                    "street": parts[1] if len(parts) > 1 else s["address"]}
            folder = RP.match_by_address(fake, addr_folders)

        prop = tkoff = None
        new_k = new_e = None
        p_note = t_note = ""
        t_frag = None
        if folder is not None:
            # PDF ONLY for the contract (the user 2026-07-21): the signed
            # proposal is the one legitimate contract source — no takeoff
            # bid-sheet substitution. Missing/unpriced PDF = a NEEDS flag,
            # not a fallback number.
            prop, new_k, p_note = find_proposal(folder, scope, s["desc"])
            tkoff, new_e, t_note, t_frag = find_takeoff_etc(
                folder, job, scope, s["desc"])
        else:
            p_note = t_note = "no project folder found"

        needs = []
        if not s.get("job"):
            needs.append("no job #")
        if not in_wip:
            if not rec:
                needs.append("not in General List" if job not in gl_jobs
                             else "in GL but unpriced")
        if new_k is None:
            needs.append("no priced bid proposal PDF")
        if new_e is None:
            needs.append("no takeoff cost")
        if not in_wip:
            group = "NEW — not in WIP"
        elif ((new_k is not None and gl_k and abs(new_k - gl_k) > max(100, 0.01 * gl_k))
              or (new_e is not None and gl_e and abs(new_e - gl_e) > max(100, 0.01 * gl_e))):
            group = "CHANGED"
        else:
            group = "MATCHES"
        items.append({
            "group": group, "line": line, "section": s["section"],
            "desc": s["desc"], "address": s["address"], "builder": s["builder"],
            "in_gl": job in gl_jobs, "gl_contract": gl_k or None,
            "gl_etc": gl_e or None, "new_contract": new_k, "new_etc": new_e,
            "proposal": prop, "p_note": p_note, "takeoff": tkoff,
            "t_note": t_note, "needs": "; ".join(needs),
            "folder": folder, "t_frag": t_frag,
            "gl_sheet": rec["source"] if rec else None,
            "gl_row": rec["gl_row"] if rec else None,
        })

    n_new = sum(1 for i in items if i["group"].startswith("NEW"))
    n_chg = sum(1 for i in items if i["group"] == "CHANGED")
    print(f"  NEW (schedule→WIP lag): {n_new} · CHANGED: {n_chg} · "
          f"MATCHES: {len(items) - n_new - n_chg}")
    out = Path(os.getenv("RP_SCHED_PREVIEW_XLSX",
               str(Path.home() / "Downloads"
                   / "RP WIP - Schedule Method Preview.xlsx")))
    write_report(items, label, out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
