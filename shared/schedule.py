"""
schedule.py — weekly crew-schedule stage model (shared).

The daily crew schedules live at
`…/OPERATIONS/SCHEDULE/<yr>/<month>/Schedule M-D-YY.xlsx` ('Daily Schedule'
tab): crew/foreman names in col A above their jobs, each job row carrying an
address (col B), city (C), builder (D) and the STAGE that day (E, 'Description'
— Pour / Wreck / Forms / …).

This module reads a Mon–Fri week into a job × day × stage model, matches pricing
best-effort from the WIP master by address, and classifies each stage into a
coloured category. Extracted from one-offs/schedule_report.py (2026-07-24) the
moment company-tracker also needed it — repo rule: tools never import tools,
shared code lives here.
"""
from __future__ import annotations

import datetime as dt
import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from openpyxl import load_workbook

from shared import paths

SCHEDULE_DIR = paths.get_path("RP_SCHEDULE_DIR", "/Volumes/Common/OPERATIONS/SCHEDULE")
WIP_PATH = paths.get_path(
    "WIP_EXCEL_PATH",
    paths.onedrive_base() / "Company Files - WIP Report/WIP - MASTER new.xlsx")
GL_PATH = paths.get_path(
    "RP_ALPHA_PATH",
    "/Volumes/Common/OPERATIONS/GENERAL LIST/LISTA GENERAL AÑO 2026.xlsx")

# General List column map (1-based → 0-based for values_only): job C, house D,
# street E, slab-bid AI(35), flat-bid AK(37). Header row 4, data row 6+.
_GL_JOB, _GL_HOUSE, _GL_STREET, _GL_SLAB, _GL_FLAT = 2, 3, 4, 34, 36
_GL_PRICED_SHEETS = ("General list - Alpha order", "Small Jobs")
_JOB_RE = re.compile(r"^(RP\d{4}(?:-[A-Za-z]{2,6})?|CP\d{3,4})\b", re.IGNORECASE)

_FILE_RE = re.compile(r"Schedule (\d{1,2})-(\d{1,2})-(\d{2})\.xlsx$", re.IGNORECASE)

# stage → (category, hex color). First keyword hit wins.
_STAGE_RULES = [
    ("pour", ("Pour", "C00000")),
    ("wreck", ("Wreck", "305496")),
    ("form", ("Forms", "BF8F00")),
    ("cable", ("Cables", "7030A0")), ("tension", ("Cables", "7030A0")),
    ("stress", ("Stress", "1F6B4C")),
    ("plumb", ("Plumbing", "2E75B6")), ("trench", ("Plumbing", "2E75B6")),
    ("grade", ("Grade", "548235")), ("back", ("Grade", "548235")),
    ("punch", ("Punch", "7F7F7F")),
    ("set up", ("Set up", "9E480E")), ("setup", ("Set up", "9E480E")),
    ("pad", ("Pads", "9E480E")), ("patio", ("Pads", "9E480E")),
]
_DEFAULT_STAGE = ("Other", "404040")


def stage_cat(desc: str) -> Tuple[str, str]:
    lo = (desc or "").lower()
    for kw, cat in _STAGE_RULES:
        if kw in lo:
            return cat
    return _DEFAULT_STAGE


def legend() -> List[Tuple[str, str]]:
    seen: Dict[str, str] = {}
    for _, cat in _STAGE_RULES:
        seen.setdefault(cat[0], cat[1])
    seen[_DEFAULT_STAGE[0]] = _DEFAULT_STAGE[1]
    return list(seen.items())


# ────────────────────────── discovery + parse ──────────────

def _monday(d: dt.date) -> dt.date:
    return d - dt.timedelta(days=d.weekday())


def find_week_files(target: Optional[dt.date] = None) -> List[Tuple[dt.date, Path]]:
    """Mon–Fri Schedule files for the target week (or the latest week with
    files). Empty list if the schedule volume isn't mounted / none found."""
    found: Dict[dt.date, Path] = {}
    if not SCHEDULE_DIR.is_dir():
        return []
    for p in SCHEDULE_DIR.rglob("Schedule *.xlsx"):
        m = _FILE_RE.search(p.name)
        if not m or p.name.startswith("~$"):
            continue
        try:
            d = dt.date(2000 + int(m.group(3)), int(m.group(1)), int(m.group(2)))
        except ValueError:
            continue
        found[d] = p
    if not found:
        return []
    if target is None:
        # the week containing TODAY (not the newest file — next week's schedule
        # is often pre-loaded); fall back to the latest week that has files
        mon = _monday(dt.date.today())
        if not any((mon + dt.timedelta(days=i)) in found for i in range(5)):
            mon = _monday(max(found))
    else:
        mon = _monday(target)
    return [(mon + dt.timedelta(days=i), found[mon + dt.timedelta(days=i)])
            for i in range(5) if (mon + dt.timedelta(days=i)) in found]


def parse_daily(path: Path) -> List[dict]:
    """'Daily Schedule' tab → [{crew, address, city, builder, stage}]."""
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception:
        return []
    sheet = next((s for s in wb.sheetnames if s.strip().lower() == "daily schedule"),
                 wb.sheetnames[0])
    ws = wb[sheet]
    crew = None
    out = []
    for r in range(1, ws.max_row + 1):
        a = str(ws.cell(r, 1).value or "").strip()
        b = str(ws.cell(r, 2).value or "").strip()
        c = str(ws.cell(r, 3).value or "").strip()
        d = str(ws.cell(r, 4).value or "").strip()
        e = str(ws.cell(r, 5).value or "").strip()
        if a and not b and a.lower() != "name":
            crew = a
        elif b and b.lower() != "address":
            out.append({"crew": crew, "address": b, "city": c,
                        "builder": d, "stage": e})
    wb.close()
    return out


# ────────────────────────── pricing (best-effort) ──────────

_SUFFIX = {"DRIVE": "DR", "STREET": "ST", "AVENUE": "AVE", "LANE": "LN",
           "ROAD": "RD", "COURT": "CT", "BOULEVARD": "BLVD", "CIRCLE": "CIR",
           "PLACE": "PL", "TRAIL": "TRL", "PARKWAY": "PKWY", "HIGHWAY": "HWY",
           "TERRACE": "TER", "COVE": "CV", "RANCH": "RCH"}
_DIR = {"NORTH": "N", "SOUTH": "S", "EAST": "E", "WEST": "W"}


def _norm_addr(s: str) -> str:
    s = re.sub(r"[^A-Z0-9 ]", " ", (s or "").upper())
    toks = []
    for t in s.split():
        t = _SUFFIX.get(t, _DIR.get(t, t))
        toks.append(t)
    return " ".join(toks).strip()


def _money(v):
    try:
        f = float(v)
        return f if f else 0.0
    except (TypeError, ValueError):
        return 0.0


def _read_general_list(path: Path) -> Tuple[Dict[str, dict], Dict[str, float]]:
    """(by_addr, by_proj) from the General List: every RP/CP row across all
    sheets gives address→project# (broad coverage); the priced Alpha/Small
    sheets add the slab+flat bid as the contract."""
    by_addr: Dict[str, dict] = {}
    by_proj: Dict[str, float] = {}
    if not path.exists():
        return by_addr, by_proj
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception:
        return by_addr, by_proj
    for sn in wb.sheetnames:
        priced = sn in _GL_PRICED_SHEETS
        ws = wb[sn]
        for row in ws.iter_rows(min_row=6, values_only=True):
            if len(row) <= _GL_STREET:
                continue
            job = row[_GL_JOB]
            m = _JOB_RE.match(str(job).strip()) if job else None
            if not m:
                continue
            proj = m.group(1).upper()
            house = row[_GL_HOUSE] if len(row) > _GL_HOUSE else None
            street = row[_GL_STREET] if len(row) > _GL_STREET else None
            if not street:
                continue
            key = _norm_addr(f"{house or ''} {street}")
            contract = None
            if priced and len(row) > _GL_FLAT:
                c = _money(row[_GL_SLAB]) + _money(row[_GL_FLAT])
                contract = c or None
            if contract and proj not in by_proj:
                by_proj[proj] = contract
            if key:
                # don't let a later, price-less row clobber a priced one
                if key not in by_addr or (contract and by_addr[key].get("contract") is None):
                    by_addr[key] = {"proj": proj, "contract": contract, "name": street}
    wb.close()
    return by_addr, by_proj


def _read_wip(path: Path) -> Tuple[Dict[str, dict], Dict[str, float]]:
    """(by_addr, by_proj) from the WIP master Test-Master — clean active
    contracts (PROJECT NAME is the address for RP)."""
    by_addr: Dict[str, dict] = {}
    by_proj: Dict[str, float] = {}
    if not path.exists():
        return by_addr, by_proj
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception:
        return by_addr, by_proj
    if "Test-Master" not in wb.sheetnames:
        wb.close()
        return by_addr, by_proj
    ws = wb["Test-Master"]
    hdr_row, H = None, {}
    for r in range(1, 6):
        vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        if any("CONTRACT" in str(v).upper() for v in vals if v):
            hdr_row = r
            H = {str(v).strip().upper(): i + 1 for i, v in enumerate(vals) if v}
            break
    if not hdr_row:
        wb.close()
        return by_addr, by_proj
    pc, nc, cc = H.get("PROJECT #"), H.get("PROJECT NAME"), H.get("TOTAL CONTRACT PRICE")
    for r in range(hdr_row + 1, ws.max_row + 1):
        name = ws.cell(r, nc).value if nc else None
        proj = str(ws.cell(r, pc).value or "").upper() if pc else ""
        contract = ws.cell(r, cc).value if cc else None
        contract = contract if isinstance(contract, (int, float)) else None
        if proj and contract and proj not in by_proj:
            by_proj[proj] = contract
        if name:
            key = _norm_addr(str(name))
            if key and key not in by_addr:
                by_addr[key] = {"proj": proj, "contract": contract, "name": str(name)}
    wb.close()
    return by_addr, by_proj


def build_price_map(path: Path = WIP_PATH, gl_path: Path = GL_PATH) -> dict:
    """Address→pricing lookup, broadened: the General List gives address→
    project# across ALL its jobs plus bid prices on the priced sheets; the WIP
    master overlays clean active contracts. Returns {'by_addr':…, 'by_proj':…}.
    match_price cross-looks the project# in by_proj when an address has no price
    of its own."""
    gl_addr, gl_proj = _read_general_list(gl_path)
    wip_addr, wip_proj = _read_wip(path)
    by_addr = dict(gl_addr)
    by_addr.update(wip_addr)                       # WIP contract wins on collision
    by_proj = dict(gl_proj)
    by_proj.update(wip_proj)                       # WIP contract wins
    return {"by_addr": by_addr, "by_proj": by_proj}


def _with_price(hit: dict, by_proj: Dict[str, float]) -> dict:
    if hit and hit.get("contract") is None and hit.get("proj") in by_proj:
        return {**hit, "contract": by_proj[hit["proj"]]}
    return hit


def match_price(address: str, pmap: dict) -> dict:
    """Exact → house-number range → house# + street token-overlap fuzzy, then a
    project#→contract cross-lookup so a matched job shows a price even if its own
    address row had none."""
    by_addr = pmap.get("by_addr", {})
    by_proj = pmap.get("by_proj", {})
    miss = {"proj": "", "contract": None, "name": ""}
    key = _norm_addr(address)
    if not key:
        return miss
    if key in by_addr:
        return _with_price(by_addr[key], by_proj)
    toks = key.split()
    # house-number range (e.g. "420-422 ESA PARK" → "420 ESA PARK")
    if toks and "-" in toks[0]:
        alt = " ".join([toks[0].split("-")[0]] + toks[1:])
        if alt in by_addr:
            return _with_price(by_addr[alt], by_proj)
    # fuzzy: same house number + majority street-token overlap
    if len(toks) >= 2:
        house, street = toks[0], set(toks[1:])
        best, best_score = None, 0.0
        for k, v in by_addr.items():
            kt = k.split()
            if not kt or kt[0] != house:
                continue
            ks = set(kt[1:])
            if not ks:
                continue
            j = len(street & ks) / len(street | ks)
            if j > best_score:
                best, best_score = v, j
        if best and best_score >= 0.6:
            return _with_price(best, by_proj)
    return miss


# ────────────────────────── model ──────────────────────────

def build_model(week: Optional[List[Tuple[dt.date, Path]]] = None,
                pmap: Optional[Dict[str, dict]] = None) -> dict:
    """job (address) → {proj, contract, builder, current, days:{date:stage}}.
    Returns {'dates': [...], 'jobs': [...]} (empty if no schedule files)."""
    if week is None:
        week = find_week_files()
    if not week:
        return {"dates": [], "jobs": []}
    if pmap is None:
        pmap = build_price_map()
    jobs: Dict[str, dict] = {}
    for d, path in week:
        for row in parse_daily(path):
            addr = row["address"]
            j = jobs.setdefault(addr, {"address": addr, "builder": row["builder"],
                                       "city": row["city"], "days": {},
                                       "crew": row["crew"]})
            if row["stage"]:
                j["days"][d] = row["stage"]
            if row["builder"] and not j["builder"]:
                j["builder"] = row["builder"]
    for addr, j in jobs.items():
        pr = match_price(addr, pmap)
        j["proj"], j["contract"] = pr["proj"], pr["contract"]
        last = max((d for d in j["days"]), default=None)
        j["current"] = j["days"].get(last, "") if last else ""
    dates = [d for d, _ in week]
    ordered = sorted(jobs.values(), key=lambda j: (-len(j["days"]), j["address"]))
    return {"dates": dates, "jobs": ordered}
