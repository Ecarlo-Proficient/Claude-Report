#!/usr/bin/env python3
"""
cp_overview.py - the CP OVERVIEW: projections against actuals, with each job's
draw coverage on its own sheet.

The owner 2026-10-02: "a cp overview to review projects projections and their
actuals with the draw coverage in separate sheets, links from overview to
coverage that also shows all transactions - simple, good colour technique,
easy to read, professional and practical."

  Overview      one row per job, three coloured groups read left to right:
                PROJECTION (contract, est. cost, projected profit and %) ·
                ACTUAL TO DATE (billed, costs, gross profit, % and the gap to
                plan) · PROGRESS (% billed, % complete, coverage, costs waiting
                on the next draw). The job name opens its sheet.
  <job> sheet   the same three cards, the DRAW COVERAGE table (one row per
                draw, each draw name jumps to its transactions), then every
                transaction grouped by draw, newest first - the invoice(s)
                first, then the bills by vendor. Collapsible, open by default.

READS THE GENERATED P&L WORKBOOKS, NOT QBO - the same rule as the other
division Overviews (`completed_pnl.py`): billed and cost come from
`completed_pnl.read_source` / `_totals`, the numbers proven line-level against
QBO. The projection is the P&L's own ① block (Original Contract + Change
Orders, Original ETC + CO Costs), the draws are its DRAW COVERAGE table, the
lines are its 'Draw Data' / 'Draw N' / 'Next Draw' sheets. Every draw total on
the job sheet is a SUM of its lines, and a draw whose lines do not add up to the
P&L's own figure is reported on stdout.

Built by `completed_pnl.rebuild_overview("cp")` (and its CLI with
`--division cp`), so every CP P&L run refreshes it.
"""
from __future__ import annotations

import datetime as dt
import re
import sys
from pathlib import Path
from typing import Dict, List, Optional

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from shared.xlsx_verify import assert_clean  # noqa: E402

# ── palette: one hue per GROUP, so the eye sorts the columns before reading ──
# Projection = blue (the plan), Actual = green (what QBO says), Progress = amber
# (how far along / how well billing keeps up). Colour marks a group or a state
# (coverage, overrun) - never decoration.
FONT = "Calibri"
NAVY = "1F3A5F"
INK = "1F2937"
GREY = "4B5563"            # 7:1 on white - the 6B7280 grey failed Excel's contrast check
LINK = "0563C1"
GREEN_T = "1E6B3A"
RED_T = "B00020"
AMBER_T = "9A5B00"
GROUPS = {                      # band (white text) / header tint / card tint
    "proj": ("2F5D8A", "DCE7F3", "F2F6FB"),
    "act": ("2E7D5B", "DCEFE4", "F1F8F4"),
    "prog": ("A86A12", "F6E7CC", "FCF6EC"),
    "fin": ("2E7D5B", "DCEFE4", "F1F8F4"),
    "close": ("A86A12", "F6E7CC", "FCF6EC"),
}
F_SECTION = PatternFill("solid", fgColor="EEF1F5")
F_BAND = PatternFill("solid", fgColor="F8FAFC")
F_DRAW = PatternFill("solid", fgColor="E6ECF3")
F_GOOD = PatternFill("solid", fgColor="DDF0E3")
F_WARN = PatternFill("solid", fgColor="FBEFD5")
F_BAD = PatternFill("solid", fgColor="F8DADA")
HAIR = Side(style="thin", color="D9DEE5")
RULE = Side(style="medium", color=NAVY)
BOX = Side(style="thick", color="000000")

# Negatives carry a minus sign AND red - never colour alone (Excel's accessibility check)
MONEY = '"$"#,##0;[Red]-"$"#,##0;"-"'
MONEY_C = '"$"#,##0.00;[Red]-"$"#,##0.00;"-"'
PCT = '0.0%;[Red]-0.0%;"-"'
PTS = '+0.0%;[Red]-0.0%;0.0%'
DATE = "mm/dd/yyyy"
SZ = 12
SZ_SMALL = 11
SZ_TITLE = 20

OVERHEAD_PCT = 0.10
# Coverage = billed / costs. Under 100% the draws do not pay for the work done;
# up to 1/(1-10%) = 111% they pay the costs but not the overhead; past it they
# carry the job (the P&L's Net Cov % break-even, the same 10%).
COVER_COST = 1.0
COVER_OH = 1 / (1 - OVERHEAD_PCT)


def _cell(ws, r, c, v=None, *, size=SZ, bold=False, color=INK, fmt=None,
          fill=None, align=None, indent=0, wrap=False, italic=False):
    if isinstance(v, str) and not v.startswith("="):
        v = v.replace("—", "-")                     # QBO text carries em dashes; never on our sheets
    cell = ws.cell(row=r, column=c, value=v)
    cell.font = Font(name=FONT, size=size, bold=bold, color=color, italic=italic)
    if fmt:
        cell.number_format = fmt
    if fill is not None:
        cell.fill = fill
    if align or indent or wrap:
        cell.alignment = Alignment(horizontal=align, vertical="center",
                                   indent=indent, wrap_text=wrap)
    else:
        cell.alignment = Alignment(vertical="center")
    return cell


def _link(cell, target: str, size=SZ, bold=False):
    cell.hyperlink = target
    cell.font = Font(name=FONT, size=size, bold=bold, color=LINK, underline="single")
    return cell


def _num(v) -> float:
    return float(v) if isinstance(v, (int, float)) else 0.0


# ───────────────────────────── read the P&L ──────────────────────────────

def read_projection(ws) -> dict:
    """The ① block: contract (+ COs) and ETC (+ CO costs). 0 = not on file.
    Labels matched by prefix; the value is the cell to the right."""
    got = {"contract": 0.0, "co": 0.0, "etc": 0.0, "co_cost": 0.0}
    keys = (("original contract price", "contract"), ("bid proposal (contract)", "contract"),
            ("change orders", "co"), ("original etc", "etc"), ("co costs", "co_cost"))
    for r in range(1, 40):
        for c in (1, 2):
            lbl = str(ws.cell(r, c).value or "").strip().lower()
            if not lbl:
                continue
            val = ws.cell(r, c + 1).value
            if not isinstance(val, (int, float)):
                continue
            for prefix, k in keys:
                if lbl.startswith(prefix) and not got[k]:
                    got[k] = float(val)
                    break
    return {"contract": round(got["contract"] + got["co"], 2),
            "etc": round(got["etc"] + got["co_cost"], 2),
            "has_co": bool(got["co"] or got["co_cost"])}


def read_coverage(ws) -> tuple:
    """([draw rows], accumulating cost) off the P&L's DRAW COVERAGE table.
    The table sat at the top right until 2026-10-02; it now sits under the
    projection / actual blocks, so look down the sheet, not just its top."""
    hdr = None
    # the owner's new P&L layout moves the table down to ~row 25 from column A
    for r in range(1, 200):
        for c in range(1, 12):
            if str(ws.cell(r, c).value or "").strip() == "Draw":
                hdr = (r, c)
                break
        if hdr:
            break
    if not hdr:
        return [], 0.0
    hr, hc = hdr
    col = {str(ws.cell(hr, c).value or "").strip().lower(): c for c in range(hc, hc + 14)}
    draws = []
    acc = 0.0
    for r in range(hr + 1, hr + 80):
        name = str(ws.cell(r, hc).value or "").strip()
        # since 2026-10-02 the table sits on the Draws sheet, the next draw
        # (costs so far) is its first row and the total reads "TOTAL  (draws)"
        if name.lower().startswith("next draw"):
            acc = _num(ws.cell(r, col.get("costs", hc + 5)).value)
            continue
        if name.startswith("TOTAL"):
            # since 10/06/2026 the next draw sits UNDER the total
            for rr in range(r + 1, r + 6):
                if str(ws.cell(rr, hc).value or "").lower().startswith(("accumulating", "next draw")):
                    acc = _num(ws.cell(rr, col.get("costs", hc + 5)).value)
            break
        if not name:
            continue
        draws.append({"name": name,
                      "period": str(ws.cell(r, col.get("period", hc + 1)).value or ""),
                      "gross": _num(ws.cell(r, col.get("gross billed", hc + 2)).value),
                      "retained": _num(ws.cell(r, col.get("retained", hc + 3)).value),
                      "cost": _num(ws.cell(r, col.get("costs", hc + 5)).value)})
    return draws, acc


def read_draw_lines(wb) -> Dict[str, List[dict]]:
    """{draw name: [cost lines]} from the flat 'Draw Data' sheet."""
    out: Dict[str, List[dict]] = {}
    if "Draw Data" not in wb.sheetnames:
        return out
    ws = wb["Draw Data"]
    hdr = None
    for r in range(1, 6):
        for c in range(1, 4):
            if str(ws.cell(r, c).value or "").strip() == "Draw":
                hdr = (r, c)
                break
        if hdr:
            break
    if not hdr:
        return out
    hr, hc = hdr
    col = {str(ws.cell(hr, c).value or "").strip().lower(): c for c in range(hc, hc + 14)}
    for row in ws.iter_rows(min_row=hr + 1, values_only=True):
        def g(k, default=None):
            c = col.get(k)
            return row[c - 1] if c and c - 1 < len(row) else default
        name = g("draw")
        if not name:
            continue
        out.setdefault(str(name).strip(), []).append({
            "date": g("date"), "vendor": str(g("vendor") or ""),
            "code": str(g("cost code") or ""), "doc": g("bill #"),
            "amt": _num(g("amount")), "desc": str(g("description") or "")})
    return out


def read_draw_invoices(wb, draw: str) -> List[dict]:
    """The invoice block of one draw: its own 'Draw N' sheet, or - since
    2026-10-02 - its SECTION on the one 'Draws' sheet (the section band reads
    "<draw>   ·   <period>   ·   <PAID|UNPAID>")."""
    hdr = None
    if draw in wb.sheetnames:
        ws = wb[draw]
        start, stop = 1, 30
    elif "Draws" in wb.sheetnames:
        ws = wb["Draws"]
        band = re.compile(r"(^|\s)" + re.escape(draw) + r"\s+·")
        start = next((r for r in range(1, ws.max_row + 1)
                      if band.search(str(ws.cell(r, 2).value or ""))), None)
        if start is None:
            return []
        stop = start + 30
    else:
        return []
    for r in range(start, stop):
        for c in range(1, 4):
            if str(ws.cell(r, c).value or "").strip() == "Invoice #":
                hdr = (r, c)
                break
        if hdr:
            break
    if not hdr:
        return []
    hr, hc = hdr
    col = {str(ws.cell(hr, c).value or "").strip().lower(): c for c in range(hc, hc + 12)}
    out = []
    for r in range(hr + 1, hr + 60):
        doc = ws.cell(r, hc).value
        if doc is None or str(doc).startswith("TOTAL"):
            break
        out.append({"doc": doc, "date": ws.cell(r, col.get("date", hc + 1)).value,
                    # a retainage RELEASE collects retainage already billed in
                    # gross - not income (owner 09/23: gross = invoice lines
                    # except the retainage item); listed, never added
                    "billed": _num(ws.cell(r, col.get("gross billed", hc + 2)).value),
                    "release": _num(ws.cell(r, col.get("retainage billed", hc + 4)).value),
                    "paid": str(ws.cell(r, col.get("paid?", hc + 6)).value or ""),
                    "memo": str(ws.cell(r, col.get("memo", hc + 8)).value or "")})
    return out


def read_next_draw(wb) -> List[dict]:
    """Every bill on the 'Next Draw' sheet: Job Type › cost code › vendor › bill,
    told apart by their indent."""
    # since 2026-10-02 the next draw's bills sit on Draw Data under "Next
    # draw (forming)" (its sheet, when there is one, is in the draw format)
    nl = [dict(l) for k, v in read_draw_lines(wb).items()
          if k.lower().startswith("next draw") for l in v]
    if nl or "Next Draw" not in wb.sheetnames:
        return nl
    ws = wb["Next Draw"]
    out: List[dict] = []
    code = vendor = ""
    started = False
    for r in range(1, ws.max_row + 1):
        raw = ws.cell(r, 2).value
        txt = str(raw or "")
        if txt.startswith("Job Type"):
            started = True
            continue
        if not started:
            continue
        if txt.startswith("Total accumulating"):
            break
        depth = len(txt) - len(txt.lstrip(" "))
        s = txt.strip()
        if depth >= 12:
            out.append({"date": ws.cell(r, 3).value, "vendor": vendor, "code": code,
                        "doc": s, "amt": _num(ws.cell(r, 7).value),
                        "desc": str(ws.cell(r, 5).value or "")})
        elif depth >= 8:
            vendor = re.sub(r"\s*\(\d+\)\s*$", "", s)
        elif depth >= 4:
            code = s.split(" - ")[0]
    return out


def read_extra(path: Path) -> dict:
    """Projection, draws and their lines for one job's P&L workbook."""
    wb = load_workbook(str(path), data_only=False)
    try:
        ws = wb["P&L"] if "P&L" in wb.sheetnames else wb.worksheets[0]
        proj = read_projection(ws)
        rnb = None
        # when the P&L pulled QBO: its "Generated <date time>" stamp
        pulled = None
        for row in ws.iter_rows(min_row=1, max_row=3, values_only=True):
            for v in row:
                m = re.search(r"Generated (\d{4}-\d{2}-\d{2} \d{1,2}:\d{2} [AP]M)", str(v or ""))
                if m and pulled is None:
                    pulled = dt.datetime.strptime(m.group(1), "%Y-%m-%d %I:%M %p")
        for row in ws.iter_rows(min_row=1, max_row=120, max_col=3, values_only=True):
            m = next((x for x in (re.match(r"\s*#(\d+)\s.*retainage\s+not\s+billed",
                                           str(v), re.I) for v in row if v) if x), None)
            if m:
                rnb = m.group(1)
                break
        # the coverage sat on a Draws sheet for a day (2026-10-02); it is on
        # the P&L again
        draws, acc = read_coverage(wb["Draws"] if "Draws" in wb.sheetnames else ws)
        lines = read_draw_lines(wb)
        for d in draws:
            d["invoices"] = read_draw_invoices(wb, d["name"])
            d["bills"] = lines.get(d["name"], [])
        return dict(proj, draws=draws, awaiting=acc, next_lines=read_next_draw(wb),
                    rnb_doc=rnb, pulled=pulled)
    finally:
        wb.close()


# ───────────────────────────── write the book ─────────────────────────────

# Overview columns: (header, key, format, width, group)
OV_COLS = [
    ("Contract", "contract", MONEY, 15, "proj"),
    ("Est. cost (ETC)", "etc", MONEY, 13, "proj"),
    ("Proj. gross profit", "pp", MONEY, 13, "proj"),
    ("Proj. gross %", "ppm", PCT, 10, "proj"),
    (f"{OVERHEAD_PCT:.0%} OH (of contract)", "poh", MONEY, 14, "proj"),
    (f"Proj. net ({OVERHEAD_PCT:.0%} OH)", "pn", MONEY, 15, "proj"),
    ("Proj. net %", "pnm", PCT, 10, "proj"),
    ("Billed", "billed", MONEY, 15, "act"),
    ("Costs", "cost", MONEY, 13, "act"),
    ("Gross profit", "gp", MONEY, 13, "act"),
    ("Gross %", "gpm", PCT, 10, "act"),
    (f"{OVERHEAD_PCT:.0%} OH (of billed)", "oh", MONEY, 14, "act"),
    (f"Net profit ({OVERHEAD_PCT:.0%} OH)", "net", MONEY, 15, "act"),
    ("Net %", "netm", PCT, 10, "act"),
    ("Net vs plan", "drift", PTS, 13, "act"),
    ("% billed", "pbill", PCT, 11, "prog"),
    ("% complete", "pcomp", PCT, 13, "prog"),
    ("Coverage", "cov", PCT, 13, "prog"),
    ("Awaiting draw", "await", MONEY, 14, "prog"),
]
GROUP_TITLE = {"proj": "PROJECTION", "act": "ACTUAL TO DATE", "prog": "PROGRESS & COVERAGE",
               "fin": "RESULT", "close": "COVERAGE & CLOSE-OUT"}

# The COMPLETED book: a finished job's actuals ARE its result - no projection
# (its P&L falls back to an old WIP report for contract / ETC, unconfirmed;
# owner 10/02), no progress; coverage, what is still owed and why instead.
COMPLETED_COLS = [
    ("Billed", "billed", MONEY, 15, "fin"),
    ("Costs", "cost", MONEY, 15, "fin"),
    ("Gross profit", "gp", MONEY, 15, "fin"),
    ("Gross %", "gpm", PCT, 10, "fin"),
    (f"{OVERHEAD_PCT:.0%} OH (of billed)", "oh", MONEY, 14, "fin"),
    (f"Net profit ({OVERHEAD_PCT:.0%} OH)", "net", MONEY, 15, "fin"),
    ("Net %", "netm", PCT, 10, "fin"),
    ("Coverage", "cov", PCT, 11, "close"),
    ("Still owed", "ar", MONEY, 14, "close"),
    ("Close-out (QuickBooks)", "closeout", None, 52, "close"),
]


def _derived(k: str, L: dict, r: int) -> Optional[str]:
    """The formula for a derived Overview / card column on row r."""
    c = lambda key: f"{L[key]}{r}"  # noqa: E731
    # LAZY: a column set without a contract (the Completed book) must not
    # build the projection formulas it does not have
    f = {
        "pp": lambda: f'=IF(OR({c("contract")}=0,{c("etc")}=0),"",{c("contract")}-{c("etc")})',
        # PROJECTED overhead = the rate x the CONTRACT (owner 2026-09-03)
        "poh": lambda: f'=IF({c("pp")}="","",{c("contract")}*{OVERHEAD_PCT})',
        "pn": lambda: f'=IF({c("pp")}="","",{c("pp")}-{c("poh")})',
        "ppm": lambda: f'=IF({c("pp")}="","",{c("pp")}/{c("contract")})',
        "pnm": lambda: f'=IF({c("pn")}="","",{c("pn")}/{c("contract")})',
        "gp": lambda: f"={c('billed')}-{c('cost')}",
        "gpm": lambda: f'=IF({c("billed")}=0,"",{c("gp")}/{c("billed")})',
        # overhead TO DATE = the rate x GROSS BILLED - the sum of the per-draw
        # overhead on the P&L's draw table (owner 2026-09-23 / 10-01); at
        # completion billed = contract, so it lands on the projected figure
        "oh": lambda: f"={c('billed')}*{OVERHEAD_PCT}",
        "net": lambda: f"={c('gp')}-{c('oh')}",
        "netm": lambda: f'=IF({c("billed")}=0,"",{c("net")}/{c("billed")})',
        "drift": lambda: f'=IF(OR({c("pnm")}="",{c("netm")}=""),"",{c("netm")}-{c("pnm")})',
        "pbill": lambda: f'=IF({c("contract")}=0,"",{c("billed")}/{c("contract")})',
        "pcomp": lambda: f'=IF({c("etc")}=0,"",{c("cost")}/{c("etc")})',
        "cov": lambda: f'=IF({c("cost")}=0,"",{c("billed")}/{c("cost")})',
    }.get(k)
    return f() if f else None


def _coverage_rules(ws, rng: str, first: str) -> None:
    """Coverage cells: red under 100%, amber to the overhead break-even, green past it."""
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'AND(ISNUMBER({first}),{first}<{COVER_COST})'], fill=F_BAD,
        font=Font(name=FONT, color=RED_T, bold=True)))
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'AND(ISNUMBER({first}),{first}>={COVER_COST},{first}<{COVER_OH:.4f})'],
        fill=F_WARN, font=Font(name=FONT, color=AMBER_T, bold=True)))
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'AND(ISNUMBER({first}),{first}>={COVER_OH:.4f})'], fill=F_GOOD,
        font=Font(name=FONT, color=GREEN_T, bold=True)))


def _sign_rules(ws, rng: str, first: str) -> None:
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'AND(ISNUMBER({first}),{first}<0)'], font=Font(name=FONT, color=RED_T)))
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'AND(ISNUMBER({first}),{first}>0)'], font=Font(name=FONT, color=GREEN_T)))


def _legend(ws, r: int, c: int) -> None:
    """Label on row r, the three bands on row r + 1, from column c (the owner
    placed it under PROGRESS & COVERAGE, 10/02)."""
    _cell(ws, r, c, "Coverage = billed ÷ costs:", size=SZ_SMALL, color=GREY)
    r += 1
    for i, (txt, fill, col) in enumerate((
            ("under 100%  draws don't cover costs", F_BAD, RED_T),
            (f"100-{COVER_OH:.0%}  covers costs, not overhead", F_WARN, AMBER_T),
            (f"{COVER_OH:.0%}+  covers costs + 10% overhead", F_GOOD, GREEN_T))):
        cc = c + i * 3
        _cell(ws, r, cc, txt, size=SZ_SMALL, color=col, fill=fill, bold=True, indent=1)
        ws.merge_cells(start_row=r, start_column=cc, end_row=r, end_column=cc + 2)


def _excel_selections(ws) -> None:
    """A both-axes freeze saved the way Excel saves it: one selection per pane,
    each with an active cell, the main one inside the scrolling pane.
    openpyxl's default leaves two bare and parks the cursor at A1, outside
    the pane - the 'Repaired Records: View' shape xlsx_verify rejects."""
    from openpyxl.worksheet.views import Selection
    tl = ws.freeze_panes
    col = re.match(r"[A-Z]+", tl).group(0)
    row = re.search(r"\d+", tl).group(0)
    ws.sheet_view.selection = [Selection(pane="topRight", activeCell=f"{col}1", sqref=f"{col}1"),
                               Selection(pane="bottomLeft", activeCell=f"A{row}", sqref=f"A{row}"),
                               Selection(pane="bottomRight", activeCell=tl, sqref=tl)]


def _sheet_name(job: str) -> str:
    return job[:31]


def build(jobs: List[tuple], out: Path, div: dict, sections=None, title=None,
          cols=None) -> None:
    """jobs = [(job, src, totals, extra)]: src/totals from completed_pnl,
    extra from read_extra."""
    wb = Workbook()
    ov = wb.active
    ov.title = "Overview"
    ov.sheet_view.showGridLines = False
    ov.sheet_view.zoomScale = 110
    C0 = 2
    OV = cols or OV_COLS
    LCOL = C0 + len(OV) + 1                     # the "open" link column
    L = {k: get_column_letter(C0 + 1 + i) for i, (_h, k, _f, _w, _g) in enumerate(OV)}

    # job sheets FIRST (the Overview's cells point at their totals)
    refs: Dict[str, dict] = {}
    jobs_ex = {j: ex for j, _s, _t, ex in jobs}
    order = sorted(jobs, key=lambda x: (x[1].get("status") != "Active", x[0]))
    for job, src, t, ex in order:
        refs[job] = _job_sheet(wb, job, src, t, ex, div)

    now = dt.datetime.now()
    _cell(ov, 1, C0, title or "COMMERCIAL OVERVIEW  ·  PROJECTIONS VS ACTUALS", size=SZ_TITLE,
          bold=True, color=NAVY)
    _cell(ov, 1, LCOL, f"Generated {now:%m/%d/%Y %I:%M %p}", size=SZ_SMALL, color=GREY,
          align="right")
    n_act = sum(1 for j in jobs if j[1].get("status") == "Active")
    scope = f" · {div['scope']}" if div.get("scope") else ""
    stamps = sorted(ex["pulled"] for *_x, ex in jobs if ex.get("pulled"))
    qbo = ""
    if stamps:
        lo, hi = stamps[0], stamps[-1]
        qbo = (f" · QBO data {lo:%m/%d/%Y %I:%M %p}" if lo.date() == hi.date() and
               (hi - lo).total_seconds() < 3600 else
               f" · QBO data {lo:%m/%d/%Y %I:%M %p} to {hi:%m/%d/%Y %I:%M %p}")
    _cell(ov, 2, C0, f"{n_act} active · {len(jobs) - n_act} completed{scope}{qbo}",
          size=SZ_SMALL, color=GREY)
    for c in range(C0, LCOL + 1):
        ov.cell(2, c).border = Border(bottom=HAIR)
    ov.row_dimensions[1].height = 30

    hdr_r = 4
    # group band + column headers
    _cell(ov, hdr_r, C0, "", fill=PatternFill("solid", fgColor=NAVY))
    _cell(ov, hdr_r + 1, C0, "Job", size=SZ_SMALL, bold=True, color="FFFFFF",
          fill=PatternFill("solid", fgColor=NAVY), indent=1)
    gstart: Dict[str, int] = {}
    for i, (h, k, _f, _w, g) in enumerate(OV):
        c = C0 + 1 + i
        gstart.setdefault(g, c)
        band, tint, _card = GROUPS[g]
        _cell(ov, hdr_r, c, None, fill=PatternFill("solid", fgColor=band))
        _cell(ov, hdr_r + 1, c, h, size=SZ_SMALL, bold=True, color=INK,
              fill=PatternFill("solid", fgColor=tint), align="right", wrap=True)
    for g, c in gstart.items():
        last = max(C0 + 1 + i for i, x in enumerate(OV) if x[4] == g)
        _cell(ov, hdr_r, c, GROUP_TITLE[g], size=SZ_SMALL, bold=True, color="FFFFFF",
              fill=PatternFill("solid", fgColor=GROUPS[g][0]), align="center")
        ov.merge_cells(start_row=hdr_r, start_column=c, end_row=hdr_r, end_column=last)
        for rr in (hdr_r, hdr_r + 1):                     # white seam between groups
            ov.cell(rr, c).border = Border(left=Side(style="medium", color="FFFFFF"))
    _cell(ov, hdr_r, LCOL, None, fill=PatternFill("solid", fgColor=NAVY))
    _cell(ov, hdr_r + 1, LCOL, "Detail", size=SZ_SMALL, bold=True, color="FFFFFF",
          fill=PatternFill("solid", fgColor=NAVY), align="center")
    ov.row_dimensions[hdr_r].height = 20
    ov.row_dimensions[hdr_r + 1].height = 32

    r = hdr_r + 2
    sec_rows: List[tuple] = []

    def _row(rr, job, src, bold=False):
        ref = refs[job]
        sn = _sheet_name(job)
        nm = _cell(ov, rr, C0, job_label(job, src.get("title", "")), indent=1)
        _link(nm, f"#'{sn}'!A1", bold=True)
        for i, (_h, k, fmt, _w, _g) in enumerate(OV):
            c = C0 + 1 + i
            done = src.get("status") != "Active"
            if done and k in ("contract", "etc", "await"):
                # a FINISHED job is off the WIP master, so its P&L falls back
                # to an old WIP report for contract / ETC - unconfirmed (owner
                # 10/02). Its actuals ARE the result: projection left blank.
                v = None
            elif k in ("contract", "etc"):                 # the P&L's projection
                v = jobs_ex[job][k] or None
            elif k in jobs_ex[job].get("static", {}):      # a figure the caller supplies
                v = jobs_ex[job]["static"][k]
            elif k in ref:                                 # the job sheet's own cell
                v = f"='{sn}'!{ref[k]}"
            else:
                v = _derived(k, L, rr)
            _cell(ov, rr, c, v, fmt=fmt, align="right" if fmt else "left",
                  size=SZ if fmt else SZ_SMALL, bold=k in ("pn", "net"), indent=0 if fmt else 1)
        lk = _cell(ov, rr, LCOL, "Draws  ›", align="center", size=SZ_SMALL)
        _link(lk, f"#'{sn}'!{ref['cov_anchor']}", size=SZ_SMALL)
        ov.row_dimensions[rr].height = 21

    def _section(title, sel, note):
        nonlocal r
        if not sel:
            return
        if title:                       # one group = no band (the owner's layout 10/02)
            _cell(ov, r, C0, title, bold=True, color=NAVY, fill=F_SECTION, indent=1)
            _cell(ov, r, C0 + 1, note, size=SZ_SMALL, color=GREY, fill=F_SECTION, italic=True)
            for c in range(C0 + 2, LCOL + 1):
                ov.cell(r, c).fill = F_SECTION
            ov.row_dimensions[r].height = 22
            r += 1
        first = r
        for job, src, _t, _ex in sorted(sel, key=lambda x: -x[2]["billed"]):
            _row(r, job, src)
            if (r - first) % 2 == 1:
                for c in range(C0, LCOL + 1):
                    ov.cell(r, c).fill = F_BAND
            r += 1
        sec_rows.append((first, r - 1))
        _total(r, f"Subtotal · {len(sel)} job(s)", [(first, r - 1)], rule=HAIR)
        r += 2

    def _total(rr, label, ranges, rule):
        _cell(ov, rr, C0, label, bold=True, color=NAVY, indent=1)
        def tot(k):
            return "+".join(f"SUM({L[k]}{a}:{L[k]}{b})" for a, b in ranges)

        def where(k, cond_k):
            # the sum of k over the jobs that carry a projection figure in cond_k,
            # so a job with no contract / ETC never skews a projected total
            return "+".join(f"SUMPRODUCT(({L[cond_k]}{a}:{L[cond_k]}{b}<>0)*"
                            f"({L['contract']}{a}:{L['contract']}{b}<>0)*"
                            f"{L[k]}{a}:{L[k]}{b})" for a, b in ranges)
        for i, (_h, k, fmt, _w, _g) in enumerate(OV):
            c = C0 + 1 + i
            if not fmt:                                    # a text column has no total
                v = None
            elif k in ("contract", "etc", "billed", "cost", "await", "pp", "poh", "pn", "ar"):
                v = "=" + tot(k)
            elif k in ("pnm", "ppm"):
                num = L["pn" if k == "pnm" else "pp"]
                v = f'=IF(({where("contract", "etc")})=0,"",{num}{rr}/({where("contract", "etc")}))'
            elif k == "pbill":
                v = f'=IF(({tot("contract")})=0,"",({where("billed", "contract")})/({tot("contract")}))'
            elif k == "pcomp":
                v = f'=IF(({where("etc", "etc")})=0,"",({where("cost", "etc")})/({where("etc", "etc")}))'
            else:
                v = _derived(k, L, rr)
            _cell(ov, rr, c, v, fmt=fmt, align="right", bold=True)
        for c in range(C0, LCOL + 1):
            ov.cell(rr, c).border = Border(top=rule)
        ov.row_dimensions[rr].height = 22

    groups = sections or [(None, "", lambda _j: True)]
    for title, note, keep in groups:
        _section(title, [j for j in jobs if keep(j)], note)
    if len(sec_rows) > 1:               # several groups: a grand total under them
        all_r = r
        _total(all_r, f"ALL CP · {len(jobs)} jobs", sec_rows, rule=RULE)
    else:                               # one group: its subtotal IS the total
        all_r = r - 2
    first_job = hdr_r + 2
    cov_group = next(x[4] for x in OV if x[1] == "cov")
    _legend(ov, all_r + 3, gstart[cov_group])

    # conditional colour on the job and total rows
    rows = f"{first_job}:{all_r}"
    a, b = rows.split(":")
    _coverage_rules(ov, f"{L['cov']}{a}:{L['cov']}{b}", f"{L['cov']}{a}")
    for k in ("drift", "pp", "pn", "gp", "net"):
        if k in L:
            _sign_rules(ov, f"{L[k]}{a}:{L[k]}{b}", f"{L[k]}{a}")
    if "pcomp" in L:
        ov.conditional_formatting.add(f"{L['pcomp']}{a}:{L['pcomp']}{b}", FormulaRule(
            formula=[f'AND(ISNUMBER({L["pcomp"]}{a}),{L["pcomp"]}{a}>1)'],
            font=Font(name=FONT, color=RED_T, bold=True), fill=F_BAD))
    if "ar" in L:                                  # money still owed: red when there is any
        ov.conditional_formatting.add(f"{L['ar']}{a}:{L['ar']}{b}", FormulaRule(
            formula=[f'AND(ISNUMBER({L["ar"]}{a}),{L["ar"]}{a}>0)'],
            font=Font(name=FONT, color=RED_T, bold=True)))
    for k in [x for x in ("pbill", "pcomp") if x in L]:
        ov.conditional_formatting.add(f"{L[k]}{a}:{L[k]}{b}", DataBarRule(
            start_type="num", start_value=0, end_type="num", end_value=1,
            color="9DB8D9" if k == "pbill" else "E5B66B", showValue=True))

    # dividers: a heavy rule where each group starts, and between GROSS and NET
    # inside the projection and the actuals (the owner 10/02)
    div_cols = sorted(set(gstart.values()) | {C0 + 1 + i for i, x in enumerate(OV)
                                              if x[1] in ("poh", "oh")} | {LCOL})
    for c in div_cols:
        for rr in range(hdr_r + 1, all_r + 1):
            cur = ov.cell(rr, c).border
            ov.cell(rr, c).border = Border(left=Side(style="medium", color=NAVY),
                                           top=cur.top, bottom=cur.bottom)
    ov.column_dimensions["A"].width = 2
    ov.column_dimensions[get_column_letter(C0)].width = 36
    for i, (_h, _k, _f, w, _g) in enumerate(OV):
        ov.column_dimensions[get_column_letter(C0 + 1 + i)].width = w
    ov.column_dimensions[get_column_letter(LCOL)].width = 11
    # headers AND the job column stay put (the owner's freeze, 10/02); one
    # top-left pane, which xlsx_verify accepts
    ov.freeze_panes = f"{get_column_letter(C0 + 1)}{hdr_r + 2}"
    _excel_selections(ov)

    for ws in wb.worksheets:
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
    wb.move_sheet("Overview", offset=-(len(wb.sheetnames) - 1))
    wb.active = 0
    tmp = out.with_suffix(".tmp.xlsx")
    wb.save(str(tmp))
    assert_clean(tmp)
    tmp.replace(out)


def _job_sheet(wb, job, src, t, ex, div) -> dict:
    """One job: cards, the draw coverage table, every transaction by draw.
    Returns {overview key: cell address} for the Overview's references."""
    ws = wb.create_sheet(_sheet_name(job))
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 110
    ws.sheet_properties.outlinePr.summaryBelow = False
    # A gutter | B draw/date | C period/doc | D invoices/vendor | E code/retained
    # | F billed | G cost | H profit/paid | I OH/description | J net | K cov | L cum | M status
    widths = {"A": 2, "B": 13, "C": 21, "D": 32, "E": 13, "F": 15, "G": 15, "H": 15,
              "I": 13, "J": 15, "K": 12, "L": 12, "M": 14}
    for k, w in widths.items():
        ws.column_dimensions[k].width = w
    LAST = 13

    _cell(ws, 1, 2, job_label(job, src.get("title", "")), size=SZ + 4, bold=True,
          color=NAVY)
    if src.get("rel"):
        _link(_cell(ws, 1, LAST - 1, "P&L  ↗", align="right", size=SZ_SMALL),
              src["rel"], size=SZ_SMALL)
    if ex.get("pulled"):
        _cell(ws, 1, LAST - 3, f"QBO data {ex['pulled']:%m/%d/%Y %I:%M %p}", align="right",
              size=SZ_SMALL, color=GREY)
    _link(_cell(ws, 1, LAST, "‹ Overview", align="right", size=SZ_SMALL),
          "#'Overview'!A1", size=SZ_SMALL)
    ws.row_dimensions[1].height = 24
    ref: Dict[str, str] = {}

    # ── DRAW COVERAGE table ──
    r = 3
    ref["cov_anchor"] = f"B{r}"
    _cell(ws, r, 2, "DRAW COVERAGE", bold=True, color=NAVY)
    r += 1
    heads = [(2, "Draw", "left"), (3, "Period", "left"), (4, "Invoice(s)", "left"),
             (5, "Retained", "right"), (6, "Billed", "right"), (7, "Costs", "right"),
             (8, "Gross profit", "right"), (9, f"{OVERHEAD_PCT:.0%} OH", "right"),
             (10, "Net profit", "right"), (11, "Coverage", "right"),
             (12, "To-date cov.", "right"), (13, "Invoice status", "left")]
    for c, h, al in heads:
        _cell(ws, r, c, h, size=SZ_SMALL, bold=True, color="FFFFFF",
              fill=PatternFill("solid", fgColor=NAVY), align=al,
              indent=1 if al == "left" else 0)
    ws.row_dimensions[r].height = 22
    r += 1
    draws = ex["draws"]
    tbl_first = r
    draw_rows: Dict[str, int] = {}
    for d in draws:
        draw_rows[d["name"]] = r
        r += 1
    await_row = r
    r += 1
    resid_row = r
    r += 1
    tot_row = r
    r += 2

    # ── TRANSACTIONS by draw ──
    _cell(ws, r, 2, "TRANSACTIONS", bold=True, color=NAVY)
    r += 1
    for c, h, al in [(2, "Date", "left"), (3, "Doc #", "left"),
                     (4, "Vendor / customer", "left"), (5, "Cost code", "left"),
                     (6, "Billed", "right"), (7, "Cost", "right"), (8, "Paid?", "center"),
                     (9, "Description", "left")]:
        _cell(ws, r, c, h, size=SZ_SMALL, bold=True, color="FFFFFF",
              fill=PatternFill("solid", fgColor=NAVY), align=al,
              indent=1 if al == "left" else 0)
    for c in range(10, LAST + 1):
        ws.cell(r, c).fill = PatternFill("solid", fgColor=NAVY)
    ws.row_dimensions[r].height = 22
    r += 1

    # paid / QBO link lookups from the proven Transactions read
    bill_meta: Dict[str, tuple] = {}
    for s in src["sections"]:
        for a in s["accounts"]:
            for v in a["vendors"]:
                for ln in v["lines"]:
                    bill_meta.setdefault(str(ln["doc"]), (ln.get("paid"), ln.get("url")))
    inv_meta = {str(i["doc"]): (i.get("paid"), i.get("url")) for i in src["invoices"]}

    band_of: Dict[str, int] = {}

    def _block(key, title, period, invoices, bills, cov_row):
        nonlocal r
        band = r
        band_of[key] = band
        n_b = len({str(b["doc"]) for b in bills})
        what = " · ".join(x for x in (
            f"{len(invoices)} invoice{'s' if len(invoices) != 1 else ''}" if invoices else "",
            f"{n_b} bill{'s' if n_b != 1 else ''}" if bills else "") if x) or "nothing posted"
        _cell(ws, band, 2, title, bold=True, color=NAVY, fill=F_DRAW, indent=1)
        _cell(ws, band, 3, period, color=INK, fill=F_DRAW)
        _cell(ws, band, 4, what, size=SZ_SMALL, color=GREY, fill=F_DRAW)
        _cell(ws, band, 5, None, fill=F_DRAW)
        _cell(ws, band, 8, None, fill=F_DRAW)
        _cell(ws, band, 9, None, fill=F_DRAW)
        for c in range(10, LAST):
            _cell(ws, band, c, None, fill=F_DRAW)
        _link(_cell(ws, band, LAST, "↑ coverage", align="right", size=SZ_SMALL, fill=F_DRAW),
              f"#'{ws.title}'!B{cov_row}", size=SZ_SMALL)
        ws.row_dimensions[band].height = 22
        r += 1
        first = r
        for inv in sorted(invoices, key=lambda i: str(i["date"])):
            paid, url = inv_meta.get(str(inv["doc"]), (inv.get("paid"), None))
            paid = inv.get("paid") or paid or ""
            _cell(ws, r, 2, inv["date"], fmt=DATE, align="left", indent=1)
            dc = _cell(ws, r, 3, inv["doc"], align="left", bold=True)
            dc.number_format = "0"
            if url:
                _link(dc, url, bold=True)
            kind = inv.get("kind") or ("Ret. release" if inv.get("release") and not inv["billed"]
                                       else "Invoice")
            _cell(ws, r, 4, kind, bold=True, color=GREEN_T if kind != "Ret. release" else GREY)
            _cell(ws, r, 5, "")
            _cell(ws, r, 6, inv["billed"], fmt=MONEY_C, align="right", color=GREEN_T,
                  bold=True)
            _cell(ws, r, 8, _paid(paid), align="center", size=SZ_SMALL,
                  color=GREY if _paid(paid) == "PAID" else RED_T, bold=True)
            memo = inv["memo"]
            if inv.get("release"):
                memo = (f"retainage release ${inv['release']:,.2f} - collects retainage already "
                        f"billed, not counted again · " + memo)
            _cell(ws, r, 9, memo[:160], size=SZ_SMALL, color=GREY)
            ws.row_dimensions[r].outline_level = 1
            r += 1
        by_vendor: Dict[str, list] = {}
        for b in bills:
            by_vendor.setdefault(b["vendor"] or "(no vendor)", []).append(b)
        v_cells: List[str] = []
        for vname in sorted(by_vendor, key=str.lower):
            items = sorted(by_vendor[vname], key=lambda x: str(x["date"]))
            vr = r
            _cell(ws, vr, 2, None)
            _cell(ws, vr, 4, f"{vname}  ({len(items)})", bold=True)
            for c in range(2, LAST + 1):
                ws.cell(vr, c).border = Border(top=HAIR)
            ws.row_dimensions[vr].outline_level = 1
            r += 1
            for b in items:
                _bill_row(b)
            _cell(ws, vr, 7, f"=SUM(G{vr + 1}:G{r - 1})", fmt=MONEY_C, align="right",
                  bold=True)
            v_cells.append(f"G{vr}")
        last = r - 1
        fb = f"=SUM(F{first}:F{last})" if last >= first else 0
        gb = ("=" + "+".join(v_cells)) if v_cells else 0
        _cell(ws, band, 6, fb, fmt=MONEY, align="right", bold=True, fill=F_DRAW,
              color=GREEN_T)
        _cell(ws, band, 7, gb, fmt=MONEY, align="right", bold=True, fill=F_DRAW)
        # a THICK BLACK BOX per draw (owner 10/02). The bottom edge is the TOP of
        # the spacer row under the block - a visible row - so the box still closes
        # when the bill lines are folded away.
        for rr in range(band, r + 1):
            for c in range(2, LAST + 1):
                cur = ws.cell(rr, c).border
                if rr == r:
                    ws.cell(rr, c).border = Border(top=BOX)
                    continue
                ws.cell(rr, c).border = Border(
                    left=BOX if c == 2 else cur.left, right=BOX if c == LAST else cur.right,
                    top=BOX if rr == band else cur.top, bottom=cur.bottom)
        r += 1
        return band

    def _bill_row(b):
        nonlocal r
        paid, url = bill_meta.get(str(b["doc"]), (None, None))
        _cell(ws, r, 2, b["date"], fmt=DATE, align="left", indent=1)
        dc = _cell(ws, r, 3, b["doc"], align="left")
        if isinstance(b["doc"], (int, float)):
            dc.number_format = "0"
        if url:
            _link(dc, url)
        _cell(ws, r, 5, b["code"], color=GREY, size=SZ_SMALL)
        _cell(ws, r, 7, b["amt"], fmt=MONEY_C, align="right")
        pd = _paid(paid)
        _cell(ws, r, 8, pd, align="center", size=SZ_SMALL,
              color=GREY if pd == "PAID" else (RED_T if pd else GREY))
        _cell(ws, r, 9, b["desc"][:120], size=SZ_SMALL, color=GREY)
        ws.row_dimensions[r].outline_level = 2
        ws.row_dimensions[r].hidden = True          # opens at outline level 2: vendor totals
        r += 1

    dup = _spread_retainage(job, dict(src, rnb_doc=ex.get("rnb_doc")), draws)
    # A retainage RELEASE collects retainage that was already BOOKED as income
    # (held on an invoice, the not-billed record, or a PC00 retainage invoice -
    # completed_pnl.retainage_booked), so up to that much it is not billing.
    # Past it, the retainage was never booked (CP697: early draws entered net)
    # and the release is the only place that money shows as billed - it counts.
    from completed_pnl import retainage_booked
    booked = retainage_booked(src)
    budget = booked
    extra: Dict[str, float] = {}                 # doc -> part of the release that is billing
    for i in sorted(src["invoices"], key=lambda x: str(x["date"])):
        if i["ret_billed"]:
            take = min(i["ret_billed"], max(budget, 0))
            budget -= take
            if i["ret_billed"] - take > 0.5:
                extra[str(i["doc"])] = round(i["ret_billed"] - take, 2)
    released = round(sum(i["ret_billed"] for i in src["invoices"]), 2)
    counted = round(sum(extra.values()), 2)
    if released:
        tie = ("= the retainage booked" if abs(released - booked) < 1 else
               f"vs {booked:,.2f} booked - {counted:,.2f} never booked, counted as billing"
               if counted else f"vs {booked:,.2f} booked - {booked - released:,.2f} still held")
        print(f"    · {job}: releases {released:,.2f} {tie}; {released - counted:,.2f} left out "
              f"of billed (already in gross)")
    # the P&L's draw table counts a release inside the draw's gross billed; take
    # it back out so a draw's billed is new billing only
    for d in draws:
        rel = sum(x.get("release", 0) for x in d["invoices"])
        if rel and abs(sum(x["billed"] for x in d["invoices"]) + rel - d["gross"]) < 1:
            d["gross"] = round(d["gross"] - rel, 2)
        for x in d["invoices"]:
            add = extra.get(str(x["doc"]), 0)
            if add:
                x["billed"] = round(x["billed"] + add, 2)
                x["release"] = round(x["release"] - add, 2)
                d["gross"] = round(d["gross"] + add, 2)

    nxt = ex["next_lines"]
    if nxt or ex["awaiting"]:
        _block("__await", "Next draw", "open", [], nxt, await_row)
    for d in reversed(draws):
        _block(d["name"], d["name"], d["period"], d["invoices"], d["bills"],
               draw_rows[d["name"]])
    # Everything the draw windows do not hold (history before the P&L's first
    # draw period, retainage moved by journal entry): listed line by line so
    # the sheet carries EVERY transaction, then tied to the job's own totals.
    seen_b = {str(b["doc"]) for d in draws for b in d["bills"]} | {str(b["doc"]) for b in nxt}
    seen_i = {str(i["doc"]) for d in draws for i in d["invoices"]}
    # bills are matched by DATE, not doc #: the draw sheets join split bill
    # numbers ("1063-1064-1065") that the Transactions sheet keeps apart, so a
    # doc match double-lists them. The P&L's draws start at its first period;
    # every bill before it is the history the draw table leaves out.
    start = _period_start(draws[0]["period"]) if draws else None
    o_inv = [{"doc": i["doc"], "date": i["date"],
              "billed": round(i["gross"] + extra.get(str(i["doc"]), 0), 2),
              "release": round(i["ret_billed"] - extra.get(str(i["doc"]), 0), 2),
              "paid": i.get("paid") or "", "memo": str(i.get("memo") or "")}
             for i in src["invoices"] if str(i["doc"]) not in seen_i]
    o_inv = [x for x in o_inv if not any(not dd["je"] and str(dd["doc"]) == str(x["doc"])
                                         for dd in dup)]
    for dd in [x for x in dup if x["amt"] >= 1]:
        o_inv.append({"doc": dd["doc"], "date": dd["date"], "billed": 0, "paid": "",
                      "kind": "Not counted",
                      "memo": f"also holds ${dd['amt']:,.2f} already billed as retainage on "
                              f"{dd['on']} - a double count in QBO"})
    if src.get("not_billed") and not any(dd["je"] for dd in dup):
        o_inv.append({"doc": "JE", "date": None, "billed": src["not_billed"], "paid": "",
                      "memo": "retainage by journal entry"})
    o_bills = [{"date": ln["date"], "vendor": v["name"],
                "code": str(a["name"]).split(":")[-1].strip(), "doc": ln["doc"],
                "amt": ln["amt"], "desc": str(ln.get("desc") or "")}
               for sec in src["sections"] for a in sec["accounts"] for v in a["vendors"]
               for ln in v["lines"] if _before(ln["date"], start)
               or (start is None and str(ln["doc"]) not in seen_b)]
    # t["billed"] already leaves the releases out (completed_pnl._totals since
    # 3c1fe63) - taking them out again here read CP585 / CP672 / CP861 short by
    # exactly their releases (12,663 / 30,870 / 65,760.60, caught 10/02 by
    # one-offs/cp_overview_verify.py)
    want_b = (t["billed"] - sum(dd["amt"] for dd in dup if dd["amt"] >= 1)
              - sum(d["gross"] for d in draws))
    want_c = t["cost"] - sum(d["cost"] for d in draws) - ex["awaiting"]
    gap_b = want_b - sum(x["billed"] for x in o_inv)
    gap_c = want_c - sum(x["amt"] for x in o_bills)
    if abs(gap_b) >= 1:
        o_inv.append({"doc": "", "date": None, "billed": round(gap_b, 2), "paid": "",
                      "memo": "difference to the P&L total"})
    if abs(gap_c) >= 1:
        o_bills.append({"date": None, "vendor": "difference to the P&L total", "code": "",
                        "doc": "", "amt": round(gap_c, 2), "desc": ""})
    if o_inv or o_bills:
        _block("__out", "Outside draws" if draws else "All transactions",
               "earlier / other" if draws else "", o_inv, o_bills, resid_row)

    # ── now fill the coverage table, every figure bound to its block ──
    issues: List[str] = []

    def _profit_cells(rr, bold=False):
        # overhead per draw = the rate x that draw's GROSS BILLED - the P&L's own
        # draw table since 10/01 (owner 2026-09-23: net = gross - costs - OH)
        _cell(ws, rr, 8, f"=F{rr}-G{rr}", fmt=MONEY, align="right", bold=bold)
        _cell(ws, rr, 9, f"=F{rr}*{OVERHEAD_PCT}", fmt=MONEY, align="right", color=GREY,
              bold=bold)
        _cell(ws, rr, 10, f"=H{rr}-I{rr}", fmt=MONEY, align="right", bold=True)
    for i, d in enumerate(draws):
        rr = draw_rows[d["name"]]
        band = band_of[d["name"]]
        nm = _cell(ws, rr, 2, d["name"], indent=1)
        _link(nm, f"#'{ws.title}'!B{band}", bold=True)
        _cell(ws, rr, 3, d["period"], size=SZ_SMALL)
        invs = d["invoices"]
        _cell(ws, rr, 4, " · ".join(f"#{x['doc']}" for x in invs) or "-",
              size=SZ_SMALL, color=GREY)
        _cell(ws, rr, 5, d["retained"] or None, fmt=MONEY, align="right", color=GREY)
        # billed: the invoice lines when they add up to the P&L's figure, else the P&L's
        inv_sum = sum(x["billed"] for x in invs)
        f_bill = f"=F{band}" if invs and abs(inv_sum - d["gross"]) < 1 else d["gross"]
        if invs and abs(inv_sum - d["gross"]) >= 1:
            issues.append(f"{d['name']} invoices {inv_sum:,.2f} vs P&L {d['gross']:,.2f}")
        bill_sum = sum(x["amt"] for x in d["bills"])
        g_cost = f"=G{band}" if abs(bill_sum - d["cost"]) < 1 else d["cost"]
        if abs(bill_sum - d["cost"]) >= 1:
            issues.append(f"{d['name']} bills {bill_sum:,.2f} vs P&L {d['cost']:,.2f}")
        _cell(ws, rr, 6, f_bill, fmt=MONEY, align="right")
        _cell(ws, rr, 7, g_cost, fmt=MONEY, align="right")
        _profit_cells(rr)
        _cell(ws, rr, 11, f'=IF(G{rr}=0,"",F{rr}/G{rr})', fmt=PCT, align="right", bold=True)
        _cell(ws, rr, 12, f'=IF(SUM(G${tbl_first}:G{rr})=0,"",'
                          f'SUM(F${tbl_first}:F{rr})/SUM(G${tbl_first}:G{rr}))',
              fmt=PCT, align="right")
        paid_all = invs and all(_paid(x["paid"]) == "PAID" for x in invs)
        st = ("paid" if paid_all else ("unpaid" if invs else "no invoice"))
        _cell(ws, rr, 13, st, size=SZ_SMALL, bold=True, indent=1,
              color=GREY if paid_all else RED_T)
        if i % 2 == 1:
            for c in range(2, LAST + 1):
                ws.cell(rr, c).fill = F_BAND
        ws.row_dimensions[rr].height = 20

    # awaiting next draw: costs only
    nm = _cell(ws, await_row, 2, "Next draw", indent=1)
    if "__await" in band_of:
        _link(nm, f"#'{ws.title}'!B{band_of['__await']}")
        nxt_sum = sum(x["amt"] for x in nxt)
        g_aw = f"=G{band_of['__await']}" if abs(nxt_sum - ex["awaiting"]) < 1 else ex["awaiting"]
        if abs(nxt_sum - ex["awaiting"]) >= 1:
            issues.append(f"next-draw bills {nxt_sum:,.2f} vs P&L {ex['awaiting']:,.2f}")
    else:
        g_aw = 0
    _cell(ws, await_row, 6, 0, fmt=MONEY, align="right", color=GREY)
    _cell(ws, await_row, 7, g_aw, fmt=MONEY, align="right", color=AMBER_T, bold=True)
    _profit_cells(await_row)

    # outside the draw windows: what ties the draws to the job's totals
    # (retainage moved by journal entry, history before the first draw period)
    _cell(ws, resid_row, 2, "Outside draws", indent=1, color=GREY)
    if "__out" in band_of:
        _link(ws.cell(resid_row, 2), f"#'{ws.title}'!B{band_of['__out']}")
        b_res, c_res = f"=F{band_of['__out']}", f"=G{band_of['__out']}"
    else:
        b_res = c_res = 0
    _cell(ws, resid_row, 6, b_res, fmt=MONEY, align="right", color=GREY)
    _cell(ws, resid_row, 7, c_res, fmt=MONEY, align="right", color=GREY)
    _profit_cells(resid_row)

    _cell(ws, tot_row, 2, "TOTAL", bold=True, color=NAVY, indent=1)
    _cell(ws, tot_row, 5, f"=SUM(E{tbl_first}:E{resid_row})", fmt=MONEY, align="right",
          bold=True, color=GREY)
    _cell(ws, tot_row, 6, f"=SUM(F{tbl_first}:F{resid_row})", fmt=MONEY, align="right",
          bold=True)
    _cell(ws, tot_row, 7, f"=SUM(G{tbl_first}:G{resid_row})", fmt=MONEY, align="right",
          bold=True)
    _profit_cells(tot_row, bold=True)
    _cell(ws, tot_row, 11, f'=IF(G{tot_row}=0,"",F{tot_row}/G{tot_row})', fmt=PCT,
          align="right", bold=True)
    for c in range(2, LAST + 1):
        ws.cell(tot_row, c).border = Border(top=RULE)
    ws.row_dimensions[tot_row].height = 22
    _coverage_rules(ws, f"K{tbl_first}:L{tot_row}", f"K{tbl_first}")
    _sign_rules(ws, f"H{tbl_first}:H{tot_row}", f"H{tbl_first}")
    _sign_rules(ws, f"J{tbl_first}:J{tot_row}", f"J{tbl_first}")

    # the Overview reads the table's own totals, so the two can never disagree
    ref.update({"billed": f"F{tot_row}", "cost": f"G{tot_row}", "await": f"G{await_row}"})

    ws.freeze_panes = None
    for msg in issues:
        print(f"    ⚑ {job}: {msg} - the P&L's figure is used")
    return ref


def _no_overlap(src: dict, amt: float) -> str:
    """Why a retainage-not-billed amount cannot be a double count, or '' when
    nothing proves it: no invoice carries its own retainage line (nothing to
    overlap), or the GC released exactly what was booked."""
    withheld = sum(i["withheld"] for i in src["invoices"])
    released = sum(i["ret_billed"] for i in src["invoices"])
    if withheld < 1:
        return "no invoice carries its own retainage line, nothing to overlap"
    # a double count books MORE retainage than the GC owes, so it pays back
    # less; a release that covers everything booked rules one out
    if released and released >= withheld + amt - 1:
        return (f"the GC released {released:,.2f}, all of the {withheld + amt:,.2f} booked "
                f"- no double count")
    return ""


_RNB = re.compile(r"retainage\s+not\s+billed", re.I)


def _spread_retainage(job: str, src: dict, draws: List[dict]) -> List[dict]:
    """A CP "Retainage not Billed" invoice books the pay app's retainage TO DATE
    for early draws that were entered NET (no retainage line). Spread it back
    onto those draws at net / 9 (10% retainage) each; whatever is left that
    equals retainage a later draw's invoice ALREADY carries is a double count in
    QBO and is not counted (CP790: 22,720.10 = Draws 1-2 + Draw 3's 5,401.82;
    true gross 791,905.60). Mutates the draws; returns the double counts.
    A shape it does not recognise is left alone (listed outside the draws)."""
    in_draws = {str(i["doc"]) for d in draws for i in d["invoices"]}
    out: List[dict] = []
    # the invoice itself, or - once the 12/31 JE closed it into Retainage
    # Receivable - the "retainage by journal entry" the Transactions sheet shows
    cands = [dict(inv, amt=inv["gross"] + inv["ret_billed"], je=False)
             for inv in src["invoices"]
             if str(inv["doc"]) not in in_draws and _RNB.search(str(inv.get("memo") or ""))]
    if src.get("not_billed"):
        cands.append({"doc": src.get("rnb_doc") or "JE", "date": None, "paid": "",
                      "amt": src["not_billed"], "je": True})
    for inv in cands:
        amt = inv["amt"]
        early = []
        for d in draws:
            if d["retained"]:
                break
            if d["gross"]:
                early.append(d)
        alloc = {d["name"]: round(d["gross"] / 9, 2) for d in early}
        rest = round(amt - sum(alloc.values()), 2)
        if not early or rest < -1:
            # the net-entered draws are older than the P&L's draw table: the
            # amount stays as its own line, tested against those invoices
            pre = sum(i["gross"] for i in src["invoices"]
                      if not i["withheld"] and i["gross"] and str(i["doc"]) not in in_draws)
            ties = bool(pre and abs(pre / 9 - amt) < 1)
            why = ("= net invoices before the draw table " f"{pre:,.2f} / 9" if ties
                   else _no_overlap(src, amt))
            print(f"    {'·' if why else '⚑'} {job}: retainage not billed {amt:,.2f} kept as its own "
                  f"line - {why or f'vs net invoices before the draw table {pre:,.2f} / 9 = {pre / 9 if pre else 0:,.2f} - DOES NOT TIE, check the pay app'}")
            continue
        dup_on, run = None, 0.0
        if rest >= 1:
            for d in draws[len(early):]:
                run += d["retained"]
                if abs(run - rest) < 1:
                    dup_on = d["name"] if dup_on is None else f"{dup_on.split(' - ')[0]} - {d['name']}"
                    break
                dup_on = dup_on or d["name"]
            else:
                dup_on = None
            if dup_on is None or abs(run - rest) >= 1:
                ok = _no_overlap(src, amt)
                print(f"    {'·' if ok else '⚑'} {job}: retainage not billed {amt:,.2f} kept as its own "
                      f"line - {rest:,.2f} after the spread matches no draw's retainage"
                      f"{'; ' + ok if ok else '; check it by hand'}")
                continue                          # unexplained remainder: leave it whole
        for d in early:
            a = alloc[d["name"]]
            d["invoices"].append({"doc": inv["doc"], "date": inv["date"], "billed": a,
                                  "paid": inv.get("paid") or "", "kind": "Retainage",
                                  "memo": f"retainage not billed - {d['name']} share (net / 9)"})
            d["gross"] = round(d["gross"] + a, 2)
            d["retained"] = round(d["retained"] + a, 2)
        out.append({"doc": inv["doc"], "date": inv["date"], "amt": rest, "on": dup_on,
                    "je": inv["je"], "consumed": True})
        if rest >= 1:
            print(f"    ⚑ {job}: #{inv['doc']} retainage not billed {amt:,.2f} spread over "
                  f"{', '.join(alloc)}; {rest:,.2f} duplicates {dup_on} - not counted")
    return out


def _period_start(period: str) -> Optional[dt.date]:
    m = re.match(r"\s*(\d{1,2})/(\d{1,2})/(\d{2,4})", period or "")
    if not m:
        return None
    mo, d, y = (int(x) for x in m.groups())
    return dt.date(y + 2000 if y < 100 else y, mo, d)


def _before(v, start: Optional[dt.date]) -> bool:
    if start is None or v is None:
        return False
    v = v.date() if isinstance(v, dt.datetime) else v
    return isinstance(v, dt.date) and v < start


def _paid(v) -> str:
    s = str(v or "").strip().upper()
    if s.startswith("PAID"):
        return "PAID"
    if s.startswith("PARTIAL"):
        return "PARTIAL"
    if s.startswith("UNPAID") or s.startswith("OPEN"):
        return "UNPAID"
    return s[:10]


def job_label(job: str, title: str) -> str:
    from completed_pnl import job_label as _jl          # one label rule, the Overview's
    return _jl(job, title).replace("—", "-")
