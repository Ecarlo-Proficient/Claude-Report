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
    mon = _monday(target or max(found))
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
           "PLACE": "PL", "TRAIL": "TRL", "PARKWAY": "PKWY"}


def _norm_addr(s: str) -> str:
    s = re.sub(r"[^A-Z0-9 ]", " ", (s or "").upper())
    return " ".join(_SUFFIX.get(t, t) for t in s.split()).strip()


def build_price_map(path: Path = WIP_PATH) -> Dict[str, dict]:
    """{norm_address: {proj, contract, name}} from the WIP master Test-Master
    (PROJECT NAME is the address for RP jobs)."""
    out: Dict[str, dict] = {}
    if not path.exists():
        return out
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception:
        return out
    if "Test-Master" not in wb.sheetnames:
        wb.close()
        return out
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
        return out
    pc, nc, cc = H.get("PROJECT #"), H.get("PROJECT NAME"), H.get("TOTAL CONTRACT PRICE")
    for r in range(hdr_row + 1, ws.max_row + 1):
        name = ws.cell(r, nc).value if nc else None
        if not name:
            continue
        key = _norm_addr(str(name))
        if key and key not in out:
            out[key] = {"proj": str(ws.cell(r, pc).value or "") if pc else "",
                        "contract": (ws.cell(r, cc).value if cc else None),
                        "name": str(name)}
    wb.close()
    return out


def match_price(address: str, pmap: Dict[str, dict]) -> dict:
    key = _norm_addr(address)
    if key in pmap:
        return pmap[key]
    toks = key.split()
    if len(toks) >= 2:
        pre = f"{toks[0]} {toks[1]}"
        for k, v in pmap.items():
            if k.startswith(pre):
                return v
    return {"proj": "", "contract": None, "name": ""}


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
