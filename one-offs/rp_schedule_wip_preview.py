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
import json
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

# Team-corrected source paths (RP WIP Fixes.xlsx → rp_source_overrides.json,
# 2026-07-22): WIP-LINE → {proposal, takeoff} Mac paths that OVERRIDE the
# folder guess. The team fixed wrong folders, wrong builders, and typos, so
# these win over any automatic file match.
_OVR_PATH = Path(__file__).resolve().parent / "rp_source_overrides.json"
try:
    OVERRIDES = json.loads(_OVR_PATH.read_text()) if _OVR_PATH.exists() else {}
except Exception:
    OVERRIDES = {}

# TRACT volume builders (the team 2026-07-22): NO bid proposal — contract +
# cost come from P.O.'s / a builder price list. Don't false-flag 'no proposal';
# take the General Lista price as the contract.
TRACT_TOKENS = ("CAMDEN", "GRAND HOMES", "HABITAT", "HABITART", "WILLIAM RYAN")


def _is_tract(builder: str) -> bool:
    b = _norm(builder)
    return any(t in b for t in TRACT_TOKENS)


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
    """Best takeoff for the scope → BUDGET from its cost sheet's own
    subtotal cells. Returns (path, budget, note, fragment) — note names the
    sheet + cells the number came from; fragment jump-links there.

    COMMERCIAL TAKEOFF WINS for slab scope (the user 2026-07-21, RP6586):
    when the CP PM helps on an RP job he uses the CP template — a workbook
    with a 'BID' sheet, budget in AP1948 (AP1961 in some revisions). Those
    files may not carry the RP# in the name ('Peninsula Takeoff …'), so any
    *takeoff*-named workbook is a candidate, and a found BID sheet beats the
    residential 'JobTread Cost Gral' (which holds partial garbage on those
    jobs)."""
    cands = []
    try:
        for f in folder.iterdir():
            if f.suffix.lower() not in (".xlsm", ".xlsx"):
                continue
            if f.name.startswith("~$"):
                continue
            n = _norm(f.name)
            if n.startswith(job) or job in n or "TAKEOFF" in n:
                cands.append(f)
    except OSError:
        return None, None, "folder unreadable", None
    if not cands:
        return None, None, "No budget takeoff in this folder — add it", None
    cands.sort(key=lambda f: (_score_name(_norm(f.name), scope, desc),
                              f.stat().st_mtime), reverse=True)

    if scope != "ftw":
        for f in cands[:6]:
            try:
                wb = load_workbook(f, data_only=True, read_only=True)
            except Exception:
                continue
            bid = next((s for s in wb.sheetnames if _norm(s) == "BID"), None)
            if bid is None:
                wb.close()
                continue
            ws = wb[bid]
            for ref in ("AP1948", "AP1961"):
                v = ws[ref].value
                if isinstance(v, (int, float)) and v:
                    wb.close()
                    return (f, round(float(v), 2),
                            f"Commercial Takeoff '{bid}' {ref}",
                            f"#'{bid}'!{ref}")
            wb.close()

    found_takeoff = False
    for f in cands[:4]:
        try:
            wb = load_workbook(f, data_only=True, read_only=True)
        except Exception:
            continue
        found_takeoff = True
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
    return (cands[0], None,
            ("Missing 'JobTread Cost Gral' sheet — add the budget sheet to "
             "the takeoff" if found_takeoff
             else "No budget takeoff in this folder — add it"), None)


# ─────────────────────────── report ────────────────────────────────
# Implied gross-margin sanity band for the NEW numbers (the user 2026-07-17,
# RP5542-FTW: contract $91K vs takeoff budget $48K = 47.6% GP — "way too
# low"). RP jobs run ~10–25% GP; outside the band the pair is mismatched
# (wrong scope file, stale takeoff, or missing cost sections).
_GP_HI = 0.35     # above → budget too low vs the contract
_GP_LO = 0.05     # below → margin too thin (budget too high / contract too low)


def margin_flag(contract, budget):
    """(gp_pct, instructions) for a contract/budget pair — None when sane.
    Instructions are estimator to-dos (the user 2026-07-21), one per line;
    `_split_needs` routes them to the AR (proposal) or JR (budget) column."""
    if not contract or budget is None:
        return None, None
    gp = (contract - budget) / contract
    if gp > _GP_HI:
        return gp, (f"Budget way too low vs contract (implied GP "
                    f"{gp * 100:.0f}%; RP runs ~10–25%) — find the real "
                    f"budget takeoff sheet for this scope")
    if gp < 0:
        return gp, (f"Contract below budget (GP {gp * 100:.0f}%) — "
                    f"export/confirm the latest bid proposal PDF sent to "
                    f"the client\n"
                    f"Also confirm the budget takeoff isn't carrying "
                    f"extra scope")
    if gp < _GP_LO:
        return gp, (f"Margin too thin (GP {gp * 100:.0f}%) — verify the "
                    f"budget takeoff against the contract")
    return gp, None


def _split_needs(needs: str):
    """Route each NEEDS instruction to (AR text, JR text) — the user
    2026-07-21 wants owners in SEPARATE COLUMNS, not colored runs in one
    cell. AR (Accounts/estimator) = anything about the bid proposal PDF;
    JR = budget/takeoff/sheet/folder work; pure GL-status notes are dropped
    (the GROUP + 'IN GEN. LIST?' columns already say that).

    Separate single-font columns replace the old rich-text cell entirely —
    Mac Excel rejected openpyxl's multi-run inline strings ('String
    properties' repair), so the report now carries NO rich text at all."""
    ar, jr = [], []
    for p in (p.strip() for p in re.split(r"[;\n]", needs or "") if p.strip()):
        u = p.upper()
        if "GENERAL LIST" in u or "IN GL" in u or "UNPRICED" in u:
            continue
        (ar if "PROPOSAL" in u else jr).append(p)
    return "\n".join(ar), "\n".join(jr)


def _breadcrumb(path: Path) -> str:
    """Human-navigable path text starting at CURRENT PROJECTS (the user
    2026-07-21): the file:// links die on Windows/OneDrive, so estimators
    need a breadcrumb they can walk in Explorer — 'CURRENT PROJECTS >
    Residential > <client> > <address> > <file>'. Mac links stay attached."""
    parts = list(path.parts)
    for i, p in enumerate(parts):
        if p.upper() == "CURRENT PROJECTS":
            return " > ".join(parts[i:])
    return " > ".join(parts[-4:])


# ── extraction from a SPECIFIC file (used by team-corrected overrides) ──
def pdf_subtotal(path: Path):
    """Last 'SUB TOTAL: $…' in a specific proposal PDF, or None."""
    import pdfplumber
    try:
        with pdfplumber.open(path) as pdf:
            txt = "\n".join(pg.extract_text() or "" for pg in pdf.pages)
    except Exception:
        return None
    hits = _SUBTOTAL_RE.findall(txt)
    return float(hits[-1].replace(",", "")) if hits else None


def takeoff_budget_from(f: Path, scope: str, desc: str):
    """Budget from a SPECIFIC takeoff file → (budget, note, fragment). Tries
    the commercial 'BID' sheet (AP1948/AP1961) for slab scope, then the
    residential 'JobTread Cost Gral' bands."""
    try:
        wb = load_workbook(f, data_only=True, read_only=True)
    except Exception:
        return None, "takeoff unreadable", None
    if scope != "ftw":
        bid = next((s for s in wb.sheetnames if _norm(s) == "BID"), None)
        if bid is not None:
            ws = wb[bid]
            for ref in ("AP1948", "AP1961"):
                v = ws[ref].value
                if isinstance(v, (int, float)) and v:
                    wb.close()
                    return round(float(v), 2), f"'{bid}' {ref}", f"#'{bid}'!{ref}"
    sheet = next((s for s in wb.sheetnames if "COST GRAL" in _norm(s)), None)
    if sheet is None:
        wb.close()
        return (None, "Missing 'JobTread Cost Gral' sheet — add the budget "
                "sheet to the takeoff", None)
    bands = _cost_sheet_totals(wb[sheet])
    wb.close()
    if _desc_tokens(desc) & set(_norm(f.name).split()):
        keys = ("SL", "PR", "FW")
    else:
        keys = ("FW",) if scope == "ftw" else ("SL", "PR")
    budget, cells = 0.0, []
    for k in keys:
        b = bands[k]
        if b["sub"] is not None:
            budget += b["sub"]
            cells.append(b["cell"])
        elif b["items"]:
            budget += b["items"]
            cells.append(f"{k} items")
    if budget:
        frag = (f"#'{sheet}'!"
                + next((c for c in cells if re.match(r'^[A-Z]+[0-9]+$', c)), "A1"))
        return round(budget, 2), f"'{sheet}' {' + '.join(cells)}", frag
    return (None, "Missing 'JobTread Cost Gral' sheet — add the budget sheet "
            "to the takeoff", None)


def _blank_item(**kw):
    """Full item dict with defaults so cross-check rows (from the General
    Lista) carry every key the writer reads."""
    base = dict(group="", line="", section="", desc="", address="", builder="",
                in_gl=False, gl_contract=None, gl_etc=None, new_contract=None,
                new_etc=None, proposal=None, p_note="", takeoff=None, t_note="",
                needs="", folder=None, t_frag=None, gl_sheet=None, gl_row=None)
    base.update(kw)
    return base


def _link(cell, target, fragment: str = ""):
    """file:// hyperlink + blue underline; no-op when target is None.

    The sheet/cell jump goes in the hyperlink's `location` ATTRIBUTE, not
    appended to the URI (the user 2026-07-21): a fragment like
    #'JobTread Cost Gral'!D11 puts raw spaces in the .rels target URI —
    invalid XML that makes Excel demand a repair on open."""
    from openpyxl.styles import Font
    from openpyxl.worksheet.hyperlink import Hyperlink
    if target is None or cell.value in (None, ""):
        return
    try:
        uri = Path(target).as_uri()
    except (ValueError, OSError):
        return
    loc = fragment.lstrip("#") if fragment else None
    cell.hyperlink = Hyperlink(ref=cell.coordinate, target=uri,
                               location=loc or None)
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
                f"YELLOW = General Lista numbers · GREEN = bid proposal / "
                f"takeoff numbers · every number links to its source. "
                f"If we grabbed the WRONG file (or it says missing), delete "
                f"what's in the PROPOSAL PDF / TAKEOFF FILE box and paste the "
                f"correct full file path there.")
    ws["A1"].font = Font(bold=True)
    ws.append([])
    # Legend — two plain single-font cells (NO rich text anywhere in this
    # workbook: Mac Excel rejects openpyxl's multi-run inline strings with a
    # 'String properties' repair). AR and JR each read their own column.
    ORANGE_F = Font(color="ED7D31", bold=True, size=12)
    BLUE_F = Font(color="0070C0", bold=True, size=12)
    ws["A2"] = "ORANGE = AR — bid proposal / contract actions"
    ws["A2"].font = ORANGE_F
    ws["H2"] = "BLUE = JR — budget takeoff actions"
    ws["H2"].font = BLUE_F
    HDR = ["GROUP", "WIP LINE", "WORK DESC",
           "ADDRESS", "BUILDER", "IN GENERAL LISTA?",
           "General Lista CONTRACT $", "General Lista BUDGET $",
           "NEW CONTRACT $ (proposal)", "NEW BUDGET $ (takeoff)", "NEW GP %",
           "PROPOSAL PDF", "TAKEOFF FILE",
           "AR — PROPOSAL / CONTRACT", "JR — BUDGET / TAKEOFF"]
    ws.append(HDR)
    for c in range(1, len(HDR) + 1):
        cell = ws.cell(3, c)
        cell.font = Font(bold=True)
        cell.fill = GRAY
        cell.alignment = Alignment(horizontal="center", vertical="center",
                                   wrap_text=True)
        cell.border = BORDER
    # Header tint: General Lista columns yellow, source columns green, the
    # two owner columns carry their owner's color.
    for c in (7, 8):
        ws.cell(3, c).fill = YELLOW
    for c in (9, 10, 11):
        ws.cell(3, c).fill = GREEN
    ws.cell(3, 14).font = ORANGE_F
    ws.cell(3, 15).font = BLUE_F

    order = {"⚠ IN GENERAL LISTA, NOT ON SCHEDULE": 0, "NEW — not in WIP": 1,
             "CHANGED": 2, "MATCHES": 3, "FTW BACKLOG (GL, not scheduled)": 4,
             "FLATWORK TAKEN BY OTHER": 5}
    for it in sorted(items, key=lambda x: (order.get(x["group"], 6),
                                           x["line"])):
        gp, gp_flag = margin_flag(it["new_contract"], it["new_etc"])
        needs = "; ".join(x for x in (it["needs"], gp_flag) if x)
        ar_needs, jr_needs = _split_needs(needs)
        ws.append([it["group"], it["line"], it["desc"],
                   it["address"], it["builder"],
                   ("yes" if it["in_gl"] else "NO"),
                   it["gl_contract"], it["gl_etc"],
                   it["new_contract"], it["new_etc"], gp,
                   (_breadcrumb(it["proposal"]) if it["proposal"]
                    else it["p_note"]),
                   ((_breadcrumb(it["takeoff"])
                     + (f"\n[{it['t_note']}]" if it["t_note"] else ""))
                    if it["takeoff"] else it["t_note"]),
                   ar_needs, jr_needs])
        r = ws.max_row
        for cc in range(1, len(HDR) + 1):
            ws.cell(r, cc).border = BORDER
            ws.cell(r, cc).alignment = Alignment(
                vertical="top", wrap_text=(cc in (3, 12, 13, 14, 15)))
        for cc in (7, 8, 9, 10):
            ws.cell(r, cc).number_format = CUR
        ws.cell(r, 11).number_format = "0.0%"
        # Source-colored numbers: General Lista yellow · proposal/takeoff green.
        for cc in (7, 8):
            ws.cell(r, cc).fill = YELLOW
        for cc in (9, 10, 11):
            ws.cell(r, cc).fill = GREEN
        # Owner NEEDS columns carry the owner's font color.
        if ws.cell(r, 14).value:
            ws.cell(r, 14).font = Font(color="ED7D31", bold=True)
        if ws.cell(r, 15).value:
            ws.cell(r, 15).font = Font(color="0070C0", bold=True)
        # Links: line → folder · General Lista $ → its list row · new numbers
        # → the exact proposal PDF / takeoff cell; file columns → the folder
        # (they delete the path and paste the correct one in the same box).
        _link(ws.cell(r, 2), it.get("folder"))
        gl_frag = (f"#'{it['gl_sheet']}'!C{it['gl_row']}"
                   if it.get("gl_row") else "")
        for cc in (7, 8):
            _link(ws.cell(r, cc), RP.ALPHA_PATH if it.get("gl_row") else None,
                  gl_frag)
        _link(ws.cell(r, 9), it.get("proposal"))
        _link(ws.cell(r, 10), it.get("takeoff"), it.get("t_frag") or "")
        _link(ws.cell(r, 12),
              it["proposal"].parent if it.get("proposal") else None)
        _link(ws.cell(r, 13),
              it["takeoff"].parent if it.get("takeoff") else None)
        # GROUP color: RED-flag not-on-schedule = red fill; NEW banded orange;
        # CHANGED amber-text; MATCHES green; FTW backlog / OTHER neutral-band.
        RED_FILL = PatternFill("solid", fgColor="F4CCCC")
        if it["group"].startswith("⚠"):
            ws.cell(r, 1).fill = RED_FILL
            ws.cell(r, 1).font = BAD
        elif it["group"].startswith("NEW"):
            ws.cell(r, 1).fill = NEW_F
            ws.cell(r, 1).font = BAD
        elif it["group"] == "CHANGED":
            ws.cell(r, 1).font = Font(color="BF6000", bold=True)
        elif it["group"] == "MATCHES":
            ws.cell(r, 1).font = GOOD
        else:                      # FTW backlog / taken by OTHER
            ws.cell(r, 1).fill = GRAY
            ws.cell(r, 1).font = Font(bold=True)
        if gp is not None and gp_flag:
            ws.cell(r, 11).font = BAD
        if not it["in_gl"]:
            ws.cell(r, 6).font = BAD
    widths = (16, 12, 26, 22, 20, 11, 15, 15, 14, 13, 9, 34, 34, 32, 32)
    for col, w in zip("ABCDEFGHIJKLMNO", widths):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A4"

    # ── second sheet: budget-sheet CLEANUP CHECKLIST (the user 2026-07-22) ──
    # Every takeoff missing its 'JobTread Cost Gral' sheet (or with no takeoff
    # at all) — the #1 data gap the team surfaced. Add the sheet → the budget
    # extracts automatically next run.
    cw = wb.create_sheet("CLEANUP CHECKLIST")
    cw["A1"] = ("BUDGET-SHEET CLEANUP — takeoffs missing the 'JobTread Cost "
                "Gral' sheet (or no takeoff file). Add the sheet to the takeoff "
                "and the budget reads automatically.")
    cw["A1"].font = Font(bold=True)
    cw.append([])
    ch = ["JOB #", "BUILDER", "ISSUE", "TAKEOFF FILE / FOLDER"]
    cw.append(ch)
    for c in range(1, len(ch) + 1):
        cell = cw.cell(3, c)
        cell.font = Font(bold=True)
        cell.fill = GRAY
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
    seen = set()
    for it in sorted(items, key=lambda x: (x["builder"], x["line"])):
        tn = it["t_note"] or ""
        # Genuine misses ONLY — a SUCCESSFUL extraction's note also names the
        # 'Cost Gral' sheet, so match the miss phrasing, not any mention.
        miss = ("Missing 'JobTread Cost Gral'" in tn
                or "No budget takeoff" in tn or "takeoff unreadable" in tn)
        if miss and it["line"] not in seen:
            seen.add(it["line"])
            where = (_breadcrumb(it["takeoff"]) if it["takeoff"]
                     else (_breadcrumb(it["folder"]) if it["folder"] else "—"))
            cw.append([it["line"], it["builder"], tn, where])
            rr = cw.max_row
            for cc in range(1, len(ch) + 1):
                cw.cell(rr, cc).border = BORDER
                cw.cell(rr, cc).alignment = Alignment(vertical="top",
                                                      wrap_text=True)
    for col, w in zip("ABCD", (14, 24, 50, 58)):
        cw.column_dimensions[col].width = w
    cw.freeze_panes = "A4"

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
        ov = OVERRIDES.get(line, {})
        # Team-pinned files WIN (RP WIP Fixes.xlsx corrected wrong folders,
        # wrong builders and typos — trust the human over the folder guess).
        if ov.get("proposal"):
            prop = Path(ov["proposal"])
            new_k = pdf_subtotal(prop)
            p_note = ("team-pinned" if new_k is not None
                      else "team-pinned (price list — no SUB TOTAL)")
        if ov.get("takeoff"):
            tkoff = Path(ov["takeoff"])
            new_e, t_note, t_frag = takeoff_budget_from(tkoff, scope, s["desc"])
        # Folder search fills whatever the override didn't supply.
        if folder is not None:
            if prop is None:
                prop, new_k, p_note = find_proposal(folder, scope, s["desc"])
            if tkoff is None:
                tkoff, new_e, t_note, t_frag = find_takeoff_etc(
                    folder, job, scope, s["desc"])
        elif prop is None and tkoff is None:
            p_note = t_note = "no project folder found"

        # TRACT builders (Camden, Grand Homes, Habitat…): NO bid proposal —
        # contract comes from P.O.'s / a builder price list. Take the General
        # Lista price as the contract; never false-flag 'no proposal'.
        tract = _is_tract(s["builder"])
        if tract and new_k is None:
            p_note = "Tract — contract from P.O.'s / builder price list"
            if gl_k:
                new_k = gl_k

        needs = []
        if not in_wip and not tract and not rec:
            needs.append("not in General List" if job not in gl_jobs
                         else "in GL but unpriced")
        if new_k is None:
            needs.append("Tract — enter contract from P.O.'s / price list"
                         if tract else
                         "Export/find the bid proposal PDF that was sent "
                         "to the client → drop it in this project's folder")
        if new_e is None:
            needs.append(t_note or "no budget takeoff")
        if not in_wip:
            group = "NEW — not in WIP"
        elif ((new_k is not None and gl_k and abs(new_k - gl_k) > max(100, 0.01 * gl_k))
              or (new_e is not None and gl_e and abs(new_e - gl_e) > max(100, 0.01 * gl_e))):
            group = "CHANGED"
        else:
            group = "MATCHES"
        items.append(_blank_item(
            group=group, line=line, section=s["section"], desc=s["desc"],
            address=s["address"], builder=s["builder"],
            in_gl=(job in gl_jobs), gl_contract=gl_k or None, gl_etc=gl_e or None,
            new_contract=new_k, new_etc=new_e, proposal=prop, p_note=p_note,
            takeoff=tkoff, t_note=t_note, needs="; ".join(needs), folder=folder,
            t_frag=t_frag, gl_sheet=(rec["source"] if rec else None),
            gl_row=(rec["gl_row"] if rec else None)))

    # ── General Lista cross-checks (the user 2026-07-22) ──
    #   (1) A GL job in progress but NOT on the schedule → RED: every active
    #       job must be on the schedule.
    #   (2) Flatwork from the GL: AF=OTHER = won by another contractor
    #       (excluded); otherwise a priced -FTW not on today's schedule = the
    #       backlog the schedule can't show (expected, not scheduled yet).
    sched_bases = {s["job"] for s in sched}
    sched_ftw = {s["job"] for s in sched if s["scope"] == "ftw"}
    n_notsched = n_backlog = n_other = 0
    for rec in records:
        job = rec["job"]
        if job.startswith("CP"):
            continue
        comp = rec.get("completion")
        in_progress = comp is None or (isinstance(comp, (int, float)) and comp < 0.999)
        if (rec["slab_bid"] or rec["slab_cost"]) and in_progress \
                and job not in sched_bases:
            n_notsched += 1
            items.append(_blank_item(
                group="⚠ IN GENERAL LISTA, NOT ON SCHEDULE", line=job,
                section="—", desc="in progress on the list, missing from the schedule",
                builder=str(rec.get("builder") or ""), in_gl=True,
                gl_contract=rec["slab_bid"], gl_etc=rec["slab_cost"],
                needs="All active jobs must be on the schedule — add it, or "
                      "mark it complete if it's done",
                gl_sheet=rec["source"], gl_row=rec["gl_row"]))
        if rec["flat_bid"] or rec["flat_cost"]:
            if rec.get("flat_other"):
                n_other += 1
                items.append(_blank_item(
                    group="FLATWORK TAKEN BY OTHER", line=f"{job}-FTW",
                    section="OTHER",
                    desc="AF=OTHER — flatwork won by another contractor",
                    builder=str(rec.get("builder") or ""), in_gl=True,
                    gl_contract=rec["flat_bid"], gl_etc=rec["flat_cost"],
                    needs="Excluded — another contractor won the flatwork",
                    gl_sheet=rec["source"], gl_row=rec["gl_row"]))
            elif job not in sched_ftw:
                n_backlog += 1
                items.append(_blank_item(
                    group="FTW BACKLOG (GL, not scheduled)", line=f"{job}-FTW",
                    section="BACKLOG",
                    desc="flatwork bid, not yet on the schedule",
                    builder=str(rec.get("builder") or ""), in_gl=True,
                    gl_contract=rec["flat_bid"], gl_etc=rec["flat_cost"],
                    needs="Backlog — expected flatwork, not scheduled yet",
                    gl_sheet=rec["source"], gl_row=rec["gl_row"]))

    n_new = sum(1 for i in items if i["group"].startswith("NEW"))
    n_chg = sum(1 for i in items if i["group"] == "CHANGED")
    n_match = sum(1 for i in items if i["group"] == "MATCHES")
    print(f"  Schedule-active: NEW {n_new} · CHANGED {n_chg} · MATCHES {n_match}")
    print(f"  ⚠ GL active NOT on schedule: {n_notsched} · "
          f"FTW backlog (GL): {n_backlog} · FTW taken by OTHER: {n_other}")
    out = Path(os.getenv("RP_SCHED_PREVIEW_XLSX",
               str(Path.home() / "Downloads"
                   / "RP WIP - Schedule Method Preview.xlsx")))
    write_report(items, label, out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
