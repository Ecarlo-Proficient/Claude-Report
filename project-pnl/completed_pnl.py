#!/usr/bin/env python3
"""
completed_pnl.py — the SIMPLE report for a finished job.

`project_pnl_export.py` is built for a job in flight: draws, coverage, what to
bill next, WIP projection. On a job that is DONE those questions are settled and
the sheet is a wall. This is the other shape (the user 2026-08-27): "made
simply, not small font and easy to follow so we can get a birds eye view and
swoop into the details when needed."

THREE SHEETS, and that is the whole design:
  1. Summary   — one screen. A metrics strip ACROSS the top, then cost by
                 account beside the invoices that paid for it. Every figure is
                 a link into the detail.
  2. Costs     — account → vendor → line, collapsed to accounts by default.
  3. Invoices  — every invoice, its memo, and whether it was paid.

READS THE GENERATED WORKBOOK, NOT QBO. It re-shapes `Project_PnL_<job>.xlsx`,
whose numbers have already been proven line-level against QBO by
`one-offs/pnl_line_level_audit.py`. So this cannot introduce an attribution
bug, needs no credentials, and runs in a second.

USAGE
  python3 project-pnl/completed_pnl.py MFD133
  python3 project-pnl/completed_pnl.py --all
"""
from __future__ import annotations

import argparse
import datetime as dt
import sys
from pathlib import Path
from typing import Dict, List, Optional

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from shared import pnl_paths
from shared.xlsx_verify import assert_clean

ARCHIVE = "completed mfd project p&l"
OVERHEAD_PCT = 0.10          # company view: 10% of REVENUE
MFD_OVERHEAD_PCT = 0.09      # MFD's own view: 9% of COSTS (CLAUDE.md)

# Big and plain. The whole complaint about the old sheet was that it is dense.
SZ = 14                      # body
SZ_SMALL = 12                # detail rows
SZ_TITLE = 22
NAVY = "1F3A5F"
GREEN = "1E6B3A"
RED = "B00020"
LINK = "0563C1"
GREY = "595959"
MONEY = '#,##0;[Red](#,##0)'          # birds-eye: no cents
MONEY_C = '#,##0.00;[Red](#,##0.00)'  # detail: cents
PCT = "0.0%"

F_HDR = PatternFill("solid", fgColor=NAVY)
F_BAND = PatternFill("solid", fgColor="EDF3FA")
F_KPI = PatternFill("solid", fgColor="F2F2F2")
HAIR = Side(style="thin", color="D0D7E2")
BOX = Border(left=HAIR, right=HAIR, top=HAIR, bottom=HAIR)
UNDER = Border(bottom=Side(style="medium", color=NAVY))
THICK = Side(style="medium", color=NAVY)


def _thick_box(ws, r0: int, r1: int, c0: int, c1: int) -> None:
    """Heavy rule around a block so it reads as its own thing."""
    for rr in range(r0, r1 + 1):
        for cc in range(c0, c1 + 1):
            cur = ws.cell(row=rr, column=cc).border
            ws.cell(row=rr, column=cc).border = Border(
                left=THICK if cc == c0 else cur.left,
                right=THICK if cc == c1 else cur.right,
                top=THICK if rr == r0 else cur.top,
                bottom=THICK if rr == r1 else cur.bottom)


# ───────────────────────────── read the source ─────────────────────────────

def read_source(path: Path) -> dict:
    """Invoices + account/vendor/line detail out of a generated project P&L."""
    wb = load_workbook(str(path), data_only=False)
    try:
        inv: List[dict] = []
        ws = wb["Transactions"]
        hdr = next((r for r in range(1, 60)
                    if str(ws.cell(r, 1).value or "").strip() == "Inv #"), None)
        end = next((r for r in range(hdr + 1, ws.max_row + 1)
                    if str(ws.cell(r, 1).value or "").startswith("TOTAL")), None)

        def f(rr, cc):
            v = ws.cell(rr, cc).value
            return float(v) if isinstance(v, (int, float)) else 0.0

        for r in range(hdr + 1, end):
            inv.append({"doc": ws.cell(r, 1).value, "date": ws.cell(r, 2).value,
                        "memo": ws.cell(r, 3).value or "",
                        "gross": f(r, 4), "withheld": f(r, 5),
                        "ret_billed": f(r, 7),
                        "paid": ws.cell(r, 8).value or "",
                        "url": (ws.cell(r, 1).hyperlink.target
                                if ws.cell(r, 1).hyperlink else None)})
        # rows under the income total with NO account are retainage-not-billed
        not_billed = sum(f(r, 5) for r in range(end + 1, ws.max_row + 1)
                         if not ws.cell(r, 4).value)

        # By Account: outline 0 = account, 1 = vendor, 2 = line
        sections: List[dict] = []
        cur_sec = cur_acct = cur_vendor = None
        if "By Account" in wb.sheetnames:
            a = wb["By Account"]
            for r in range(1, a.max_row + 1):
                lvl = a.row_dimensions[r].outline_level or 0
                v1 = a.cell(r, 1).value
                amt = a.cell(r, 5).value
                txt = str(v1 or "").strip()
                if txt in ("COST OF GOODS SOLD", "OPERATING EXPENSES (non-COGS)"):
                    cur_sec = {"name": txt, "total": float(amt or 0), "accounts": []}
                    sections.append(cur_sec)
                    cur_acct = None
                    continue
                if not cur_sec or txt == "Account / Vendor / Line":
                    continue
                if lvl == 0 and txt and isinstance(amt, (int, float)):
                    cur_acct = {"name": txt, "total": float(amt), "vendors": []}
                    cur_sec["accounts"].append(cur_acct)
                    cur_vendor = None
                elif lvl == 1 and cur_acct is not None and txt:
                    cur_vendor = {"name": txt, "total": float(amt or 0), "lines": []}
                    cur_acct["vendors"].append(cur_vendor)
                elif lvl == 2 and cur_vendor is not None:
                    cur_vendor["lines"].append({
                        "date": a.cell(r, 2).value, "doc": a.cell(r, 3).value,
                        "desc": a.cell(r, 4).value or "",
                        "amt": float(a.cell(r, 5).value or 0),
                        "paid": a.cell(r, 6).value or "",
                        "url": (a.cell(r, 3).hyperlink.target
                                if a.cell(r, 3).hyperlink else None)})
        title = str(wb["P&L"].cell(1, 1).value or "")
        return {"invoices": inv, "not_billed": not_billed,
                "sections": sections, "title": title}
    finally:
        wb.close()


# ───────────────────────────── write the report ────────────────────────────

def _t(ws, r, c, v, *, size=SZ, bold=False, color="000000", fmt=None,
       fill=None, align=None, wrap=False, border=None, indent=0):
    cell = ws.cell(row=r, column=c, value=v)
    cell.font = Font(size=size, bold=bold, color=color)
    if fmt:
        cell.number_format = fmt
    if fill is not None:
        cell.fill = fill
    if align or wrap or indent:
        cell.alignment = Alignment(horizontal=align, wrap_text=wrap,
                                   vertical="center", indent=indent)
    if border is not None:
        cell.border = border
    return cell


def build(job: str, src: dict, out: Path) -> None:
    billed = sum(i["gross"] + i["ret_billed"] for i in src["invoices"]) + src["not_billed"]
    cogs = next((s["total"] for s in src["sections"] if s["name"].startswith("COST")), 0.0)
    opex = next((s["total"] for s in src["sections"] if s["name"].startswith("OPERATING")), 0.0)
    cost = cogs + opex
    gp = billed - cost
    oh = billed * OVERHEAD_PCT
    net = gp - oh

    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    ws.sheet_view.showGridLines = False

    _t(ws, 1, 1, f"{job} — JOB RESULT", size=SZ_TITLE, bold=True, color=NAVY)
    _t(ws, 2, 1, f"{src['title'].replace('PROJECT P&L — ', '')}   ·   completed job   "
                 f"·   {dt.datetime.now():%m/%d/%Y}", size=SZ_SMALL, color=GREY)

    # ── metrics ACROSS the top (the user 2026-08-27) ──
    kpis = [("BILLED", billed, MONEY, NAVY), ("COST", cost, MONEY, NAVY),
            ("GROSS PROFIT", gp, MONEY, GREEN if gp >= 0 else RED),
            ("GROSS MARGIN", gp / billed if billed else 0, PCT, GREEN if gp >= 0 else RED),
            ("OVERHEAD 10%", -oh, MONEY, GREY),
            ("NET PROFIT", net, MONEY, GREEN if net >= 0 else RED),
            ("NET MARGIN", net / billed if billed else 0, PCT, GREEN if net >= 0 else RED)]
    for i, (label, val, fmt, color) in enumerate(kpis):
        c = 1 + i * 2
        _t(ws, 4, c, label, size=SZ_SMALL, bold=True, color="FFFFFF",
           fill=F_HDR, align="center", border=BOX)
        _t(ws, 5, c, val, size=SZ + 4, bold=True, color=color, fmt=fmt,
           fill=F_KPI, align="center", border=BOX)
        ws.merge_cells(start_row=4, start_column=c, end_row=4, end_column=c + 1)
        ws.merge_cells(start_row=5, start_column=c, end_row=5, end_column=c + 1)
    ws.row_dimensions[4].height = 22
    ws.row_dimensions[5].height = 34

    # ── COST BY ACCOUNT (left) beside INVOICES (right) ──
    top = 8
    _t(ws, top, 1, "WHERE THE MONEY WENT", size=SZ + 2, bold=True, color=NAVY)
    _t(ws, top, 2, cost, size=SZ + 2, bold=True, color=NAVY, fmt=MONEY, align="right")
    _t(ws, top, 3, "% of cost", size=SZ_SMALL, bold=True, color=GREY, align="right")
    for c in (1, 2, 3):
        ws.cell(row=top, column=c).border = UNDER
    r = top + 1
    acct_rows: Dict[str, int] = {}
    for sec in src["sections"]:
        _t(ws, r, 1, sec["name"], size=SZ, bold=True, color="FFFFFF", fill=F_HDR)
        _t(ws, r, 2, sec["total"], size=SZ, bold=True, color="FFFFFF",
           fill=F_HDR, fmt=MONEY, align="right")
        ws.cell(row=r, column=3).fill = F_HDR
        r += 1
        for acct in sec["accounts"]:
            acct_rows[acct["name"]] = r
            cell = _t(ws, r, 1, acct["name"], size=SZ, color=LINK, indent=1)
            cell.font = Font(size=SZ, color=LINK, underline="single")
            _t(ws, r, 2, acct["total"], size=SZ, fmt=MONEY, align="right")
            _t(ws, r, 3, (acct["total"] / cost) if cost else 0, size=SZ_SMALL,
               fmt=PCT, color=GREY, align="right")
            if r % 2 == 0:
                for c in (1, 2, 3):
                    ws.cell(row=r, column=c).fill = F_BAND
            r += 1
        r += 1
    cost_bottom = r

    ir = top
    _t(ws, ir, 5, "WHAT WE BILLED", size=SZ + 2, bold=True, color=NAVY)
    _t(ws, ir, 8, billed, size=SZ + 2, bold=True, color=NAVY, fmt=MONEY, align="right")
    for c in range(5, 9):
        ws.cell(row=ir, column=c).border = UNDER
    ir += 1
    for c, h in ((5, "Invoice"), (6, "Date"), (7, "What for"), (8, "Amount")):
        _t(ws, ir, c, h, size=SZ_SMALL, bold=True, color="FFFFFF", fill=F_HDR)
    ir += 1
    for inv in sorted(src["invoices"], key=lambda i: str(i["date"]), reverse=True):
        amt = inv["gross"] + inv["ret_billed"]
        cell = _t(ws, ir, 5, inv["doc"], size=SZ_SMALL)
        if inv["url"]:
            cell.hyperlink = inv["url"]
            cell.font = Font(size=SZ_SMALL, color=LINK, underline="single")
        d = inv["date"]
        dc = _t(ws, ir, 6, d, size=SZ_SMALL)
        dc.number_format = "mm/dd/yyyy"
        _t(ws, ir, 7, str(inv["memo"])[:70], size=SZ_SMALL)
        _t(ws, ir, 8, amt, size=SZ_SMALL, fmt=MONEY, align="right")
        if str(inv["paid"]).startswith("PAID"):
            pass
        else:
            _t(ws, ir, 9, str(inv["paid"]), size=SZ_SMALL, bold=True, color=RED)
        if ir % 2 == 0:
            for c in range(5, 9):
                ws.cell(row=ir, column=c).fill = F_BAND
        ir += 1
    if src["not_billed"]:
        _t(ws, ir, 7, "retainage moved by journal entry", size=SZ_SMALL, color=GREY)
        _t(ws, ir, 8, src["not_billed"], size=SZ_SMALL, fmt=MONEY, align="right", color=GREY)
        ir += 1

    for col, w in zip("ABCDEFGHI", (46, 18, 12, 4, 14, 13, 52, 16, 22)):
        ws.column_dimensions[col].width = w

    # ── sheet 2: the detail, collapsed to accounts ──
    det = wb.create_sheet("Costs")
    det.sheet_view.showGridLines = False
    _t(det, 1, 1, f"{job} — COST DETAIL", size=SZ_TITLE - 4, bold=True, color=NAVY)
    _t(det, 2, 1, "Account → vendor → every line. Collapsed to accounts; "
                  "click + in the margin to open one.", size=SZ_SMALL, color=GREY)
    dr = 4
    anchors: Dict[str, int] = {}
    for sec in src["sections"]:
        _t(det, dr, 1, sec["name"], size=SZ + 2, bold=True, color="FFFFFF", fill=F_HDR)
        _t(det, dr, 5, sec["total"], size=SZ + 2, bold=True, color="FFFFFF",
           fill=F_HDR, fmt=MONEY_C, align="right")
        for c in range(2, 5):
            det.cell(row=dr, column=c).fill = F_HDR
        dr += 1
        for acct in sec["accounts"]:
            anchors[acct["name"]] = dr
            _t(det, dr, 1, acct["name"], size=SZ, bold=True, color=NAVY, fill=F_BAND)
            _t(det, dr, 5, acct["total"], size=SZ, bold=True, color=NAVY,
               fill=F_BAND, fmt=MONEY_C, align="right")
            for c in range(2, 5):
                det.cell(row=dr, column=c).fill = F_BAND
            dr += 1
            for v in acct["vendors"]:
                _t(det, dr, 1, v["name"], size=SZ_SMALL, bold=True, indent=1)
                _t(det, dr, 5, v["total"], size=SZ_SMALL, bold=True,
                   fmt=MONEY_C, align="right")
                det.row_dimensions[dr].outline_level = 1
                det.row_dimensions[dr].hidden = True
                dr += 1
                for ln in v["lines"]:
                    dc = _t(det, dr, 2, ln["date"], size=SZ_SMALL)
                    dc.number_format = "mm/dd/yyyy"
                    c3 = _t(det, dr, 3, ln["doc"], size=SZ_SMALL)
                    if ln["url"]:
                        c3.hyperlink = ln["url"]
                        c3.font = Font(size=SZ_SMALL, color=LINK, underline="single")
                    _t(det, dr, 4, str(ln["desc"])[:80], size=SZ_SMALL)
                    _t(det, dr, 5, ln["amt"], size=SZ_SMALL, fmt=MONEY_C, align="right")
                    _t(det, dr, 6, ln["paid"], size=SZ_SMALL, color=GREY)
                    det.row_dimensions[dr].outline_level = 2
                    det.row_dimensions[dr].hidden = True
                    dr += 1
        dr += 1
    for col, w in zip("ABCDEF", (52, 14, 16, 78, 18, 24)):
        det.column_dimensions[col].width = w
    det.sheet_properties.outlinePr.summaryBelow = False

    # summary account names jump to the detail block
    for name, srow in acct_rows.items():
        a = anchors.get(name)
        if a:
            ws.cell(row=srow, column=1).hyperlink = f"#'Costs'!A{a}"

    tmp = out.with_suffix(".tmp.xlsx")
    wb.save(str(tmp))
    assert_clean(tmp)
    tmp.replace(out)


def _totals(src: dict) -> dict:
    billed = sum(i["gross"] + i["ret_billed"] for i in src["invoices"]) + src["not_billed"]
    cogs = next((x["total"] for x in src["sections"] if x["name"].startswith("COST")), 0.0)
    opex = next((x["total"] for x in src["sections"] if x["name"].startswith("OPERATING")), 0.0)
    cost = cogs + opex
    gp = billed - cost
    oh = billed * OVERHEAD_PCT
    moh = cost * MFD_OVERHEAD_PCT
    return {"billed": billed, "cogs": cogs, "opex": opex, "cost": cost,
            "gp": gp, "oh": oh, "net": gp - oh,
            "gpm": gp / billed if billed else 0.0,
            "netm": (gp - oh) / billed if billed else 0.0,
            "moh": moh, "mnet": gp - moh,
            "mnetm": (gp - moh) / billed if billed else 0.0}


def _tiles(ws, r: int, items, spans) -> None:
    """Metric tiles on MERGED spans of their own.

    They used to borrow the data table's column widths, so a label like
    "Company — 10% of revenue" was clipped by an 18-wide column while its
    value landed at the far edge of a 74-wide description column, nowhere
    near it. A tile owns its span, and the label and value are centred in it
    (the user 2026-08-27: "it just looks amateurish")."""
    for (label, val, fmt, color), (c0, c1) in zip(items, spans):
        ws.merge_cells(start_row=r, start_column=c0, end_row=r, end_column=c1)
        ws.merge_cells(start_row=r + 1, start_column=c0, end_row=r + 1, end_column=c1)
        _t(ws, r, c0, label, size=SZ_SMALL, bold=True, color="FFFFFF",
           fill=F_HDR, align="center")
        _t(ws, r + 1, c0, val, size=SZ + 6, bold=True, color=color, fmt=fmt,
           fill=F_KPI, align="center")
        for cc in range(c0, c1 + 1):
            ws.cell(row=r, column=cc).border = BOX
            ws.cell(row=r + 1, column=cc).border = BOX
            ws.cell(row=r, column=cc).fill = F_HDR
            ws.cell(row=r + 1, column=cc).fill = F_KPI
    ws.row_dimensions[r].height = 20
    ws.row_dimensions[r + 1].height = 38


def _kpi_strip(ws, r: int, t: dict, spans) -> int:
    """Metrics ACROSS, exactly 4 cells wide so the strip lines up with the
    table beneath it instead of running off to column N (the user 2026-08-27).
    The two OVERHEAD views get their own boxed block below, because MFD's 9%
    of cost and the company's 10% of revenue are different answers and the
    MFD one must not get lost among the others."""
    _tiles(ws, r, [
        ("BILLED", t["billed"], MONEY, NAVY),
        ("COST", t["cost"], MONEY, NAVY),
        ("GROSS PROFIT", t["gp"], MONEY, GREEN if t["gp"] >= 0 else RED),
        ("GROSS MARGIN", t["gpm"], PCT, GREEN if t["gp"] >= 0 else RED)], spans)

    r2 = r + 3
    lab, oh_c, net_c, pct_c = spans[0][0], spans[1], spans[2], spans[3]
    ws.merge_cells(start_row=r2, start_column=lab, end_row=r2, end_column=spans[0][1])
    _t(ws, r2, lab, "AFTER OVERHEAD", size=SZ_SMALL, bold=True, color="FFFFFF",
       fill=F_HDR)
    for span, txt in ((oh_c, "OVERHEAD"), (net_c, "NET PROFIT"), (pct_c, "NET MARGIN")):
        ws.merge_cells(start_row=r2, start_column=span[0], end_row=r2, end_column=span[1])
        _t(ws, r2, span[0], txt, size=SZ_SMALL, bold=True, color="FFFFFF",
           fill=F_HDR, align="center")
    for cc in range(spans[0][0], spans[3][1] + 1):
        ws.cell(row=r2, column=cc).fill = F_HDR
    for i, (lbl, oh, net, netm) in enumerate((
            ("Company — 10% of revenue", t["oh"], t["net"], t["netm"]),
            ("MFD — 9% of cost", t["moh"], t["mnet"], t["mnetm"]))):
        rr = r2 + 1 + i
        ws.merge_cells(start_row=rr, start_column=lab, end_row=rr, end_column=spans[0][1])
        _t(ws, rr, lab, lbl, size=SZ, bold=(i == 1), indent=1)
        for span, val, fmt, col in (
                (oh_c, -oh, MONEY, GREY),
                (net_c, net, MONEY, GREEN if net >= 0 else RED),
                (pct_c, netm, PCT, GREEN if net >= 0 else RED)):
            ws.merge_cells(start_row=rr, start_column=span[0], end_row=rr,
                           end_column=span[1])
            _t(ws, rr, span[0], val, size=SZ + 1, bold=True, fmt=fmt,
               color=col, align="center")
        ws.row_dimensions[rr].height = 24
    _thick_box(ws, r2, r2 + 2, spans[0][0], spans[3][1])
    return r2 + 4


def build_bundle(jobs: List[tuple], out: Path) -> None:
    """MFD Overview Total: the overview sheet, then a sheet per job carrying
    its P&L and its transactions grouped account → vendor → line.

    EVERY link is INTERNAL. The per-job reports link to files beside them,
    which is right on the share and broken the moment the file is emailed —
    this one has to survive being sent (the user 2026-08-27).
    """
    wb = Workbook()
    sm = wb.active
    sm.title = "Summary"
    sm.sheet_view.showGridLines = False

    _t(sm, 1, 1, "MFD OVERVIEW — TOTAL", size=SZ_TITLE, bold=True, color=NAVY)
    _t(sm, 2, 1, f"{len(jobs)} finished jobs · click a job to open its detail · "
                 f"{dt.datetime.now():%m/%d/%Y}", size=SZ_SMALL, color=GREY)

    tot = {k: sum(t[k] for _, _, t in jobs)
           for k in ("billed", "cost", "gp", "oh", "net", "moh", "mnet")}
    tot["mnetm"] = tot["mnet"] / tot["billed"] if tot["billed"] else 0
    tot["gpm"] = tot["gp"] / tot["billed"] if tot["billed"] else 0
    tot["netm"] = tot["net"] / tot["billed"] if tot["billed"] else 0
    SUM_SPANS = [(1, 2), (3, 4), (5, 6), (7, 7)]
    r = _kpi_strip(sm, 4, tot, SUM_SPANS)

    heads = ["JOB", "BILLED", "COST", "GROSS PROFIT", "GP %",
             "NET  (company 10%)", "NET  (MFD 9%)"]
    for c, h in enumerate(heads, 1):
        _t(sm, r, c, h, size=SZ_SMALL, bold=True, color="FFFFFF", fill=F_HDR,
           align="center" if c > 1 else "left", wrap=True, border=BOX)
    sm.row_dimensions[r].height = 32
    r += 1
    for job, _src, t in sorted(jobs, key=lambda x: -x[2]["billed"]):
        cell = _t(sm, r, 1, job, size=SZ, bold=True)
        cell.hyperlink = f"#'{job}'!A1"
        cell.font = Font(size=SZ, bold=True, color=LINK, underline="single")
        _t(sm, r, 2, t["billed"], size=SZ, fmt=MONEY, align="right")
        _t(sm, r, 3, t["cost"], size=SZ, fmt=MONEY, align="right")
        _t(sm, r, 4, t["gp"], size=SZ, bold=True, fmt=MONEY, align="right",
           color=GREEN if t["gp"] >= 0 else RED)
        _t(sm, r, 5, t["gpm"], size=SZ, fmt=PCT, align="right",
           color=GREEN if t["gp"] >= 0 else RED)
        _t(sm, r, 6, t["net"], size=SZ, fmt=MONEY, align="right",
           color=GREEN if t["net"] >= 0 else RED)
        _t(sm, r, 7, t["mnet"], size=SZ, bold=True, fmt=MONEY, align="right",
           color=GREEN if t["mnet"] >= 0 else RED)
        for c in range(1, 8):
            sm.cell(row=r, column=c).border = BOX
            if r % 2 == 0:
                sm.cell(row=r, column=c).fill = F_BAND
        sm.row_dimensions[r].height = 22
        r += 1
    _t(sm, r, 1, f"TOTAL — {len(jobs)} JOBS", size=SZ, bold=True, fill=F_KPI)
    for c, k, fmt in ((2, "billed", MONEY), (3, "cost", MONEY), (4, "gp", MONEY),
                      (5, "gpm", PCT), (6, "net", MONEY), (7, "mnet", MONEY)):
        _t(sm, r, c, tot[k], size=SZ, bold=True, fmt=fmt, align="right",
           fill=F_KPI, color=(GREEN if tot["gp"] >= 0 else RED) if c in (4, 5, 6, 7) else "000000")
    for c in range(1, 8):
        sm.cell(row=r, column=c).border = BOX
    # MFD's own overhead column is boxed on the job table too, so it reads as
    # a distinct answer rather than another number in the row.
    _thick_box(sm, r - len(jobs) - 1, r, 7, 7)
    sm.row_dimensions[r].height = 24
    for col, w in zip("ABCDEFG", (19, 19, 19, 19, 12, 22, 22)):
        sm.column_dimensions[col].width = w
    sm.freeze_panes = "A4"

    for job, src, t in sorted(jobs, key=lambda x: -x[2]["billed"]):
        ws = wb.create_sheet(job[:31])
        ws.sheet_view.showGridLines = False
        _t(ws, 1, 1, f"{job} — JOB RESULT", size=SZ_TITLE - 2, bold=True, color=NAVY)
        back = _t(ws, 1, 7, "← back to Summary", size=SZ_SMALL, align="right")
        back.hyperlink = "#'Summary'!A1"
        back.font = Font(size=SZ_SMALL, color=LINK, underline="single")
        _t(ws, 2, 1, src["title"].replace("PROJECT P&L — ", ""), size=SZ_SMALL, color=GREY)
        JOB_SPANS = [(1, 2), (3, 4), (5, 6), (7, 7)]
        r = _kpi_strip(ws, 4, t, JOB_SPANS)

        # ONE grid for both blocks: A label/doc · B date · C what-for ·
        # D amount · E paid. No column is ever left blank in a data row —
        # the empty column A down the cost detail is what made the first cut
        # read as sloppy (the user 2026-08-27).
        _t(ws, r, 1, "INVOICED", size=SZ + 2, bold=True, color=NAVY)
        _t(ws, r, 6, t["billed"], size=SZ + 2, bold=True, color=NAVY,
           fmt=MONEY, align="right")
        r += 1
        for c, h in ((1, "Invoice"), (2, "Date"), (3, "What for"),
                     (6, "Amount"), (7, "Paid?")):
            _t(ws, r, c, h, size=SZ_SMALL, bold=True, color="FFFFFF",
               fill=F_HDR, align="center" if c in (6, 7) else "left")
        for c in range(1, 8):
            ws.cell(row=r, column=c).fill = F_HDR
            ws.cell(row=r, column=c).border = BOX
        ws.row_dimensions[r].height = 22
        r += 1
        for inv in sorted(src["invoices"], key=lambda i: str(i["date"]), reverse=True):
            # An invoice number is an IDENTIFIER, so it reads left. Numeric
            # right-alignment parked it at the far edge of a wide column,
            # disconnected from its own header.
            c1 = _t(ws, r, 1, inv["doc"], size=SZ_SMALL, align="left")
            c1.number_format = "0"
            if inv["url"]:                      # → the invoice in QBO
                c1.hyperlink = inv["url"]
                c1.font = Font(size=SZ_SMALL, color=LINK, underline="single")
            dc = _t(ws, r, 2, inv["date"], size=SZ_SMALL, align="left")
            dc.number_format = "mm/dd/yyyy"
            _t(ws, r, 3, str(inv["memo"])[:110], size=SZ_SMALL)   # spills D:E
            _t(ws, r, 6, inv["gross"] + inv["ret_billed"], size=SZ_SMALL,
               fmt=MONEY, align="right")
            _t(ws, r, 7, str(inv["paid"]), size=SZ_SMALL, align="center",
               color=GREY if str(inv["paid"]).startswith("PAID") else RED)
            if r % 2 == 0:
                for c in range(1, 8):
                    ws.cell(row=r, column=c).fill = F_BAND
            ws.row_dimensions[r].height = 20
            r += 1
        if src["not_billed"]:
            _t(ws, r, 1, "(journal entry)", size=SZ_SMALL, color=GREY, align="left")
            _t(ws, r, 3, "retainage moved by journal entry", size=SZ_SMALL, color=GREY)
            _t(ws, r, 6, src["not_billed"], size=SZ_SMALL, fmt=MONEY,
               align="right", color=GREY)
            r += 1
        r += 1

        _t(ws, r, 1, "COSTS — account, then vendor, then every line",
           size=SZ + 2, bold=True, color=NAVY)
        _t(ws, r, 6, t["cost"], size=SZ + 2, bold=True, color=NAVY,
           fmt=MONEY, align="right")
        r += 1
        for c, h in ((1, "Account / Vendor / Doc #"), (2, "Date"),
                     (3, "Description"), (6, "Amount"), (7, "Paid?")):
            _t(ws, r, c, h, size=SZ_SMALL, bold=True, color="FFFFFF",
               fill=F_HDR, align="center" if c in (6, 7) else "left")
        for c in range(1, 8):
            ws.cell(row=r, column=c).fill = F_HDR
            ws.cell(row=r, column=c).border = BOX
        ws.row_dimensions[r].height = 22
        r += 1
        for sec in src["sections"]:
            _t(ws, r, 1, sec["name"], size=SZ, bold=True, color="FFFFFF", fill=F_HDR)
            _t(ws, r, 6, sec["total"], size=SZ, bold=True, color="FFFFFF",
               fill=F_HDR, fmt=MONEY, align="right")
            for c in (2, 3, 4, 5, 7):
                ws.cell(row=r, column=c).fill = F_HDR
            ws.row_dimensions[r].height = 24
            r += 1
            for acct in sec["accounts"]:
                _t(ws, r, 1, acct["name"], size=SZ, bold=True, color=NAVY, fill=F_BAND)
                _t(ws, r, 6, acct["total"], size=SZ, bold=True, color=NAVY,
                   fill=F_BAND, fmt=MONEY, align="right")
                for c in (2, 3, 4, 5, 7):
                    ws.cell(row=r, column=c).fill = F_BAND
                ws.row_dimensions[r].height = 22
                r += 1
                for v in acct["vendors"]:
                    _t(ws, r, 1, v["name"], size=SZ_SMALL, bold=True, indent=1)
                    _t(ws, r, 6, v["total"], size=SZ_SMALL, bold=True,
                       fmt=MONEY, align="right")
                    ws.row_dimensions[r].outline_level = 1
                    ws.row_dimensions[r].hidden = True
                    r += 1
                    for ln in v["lines"]:
                        dc1 = _t(ws, r, 1, ln["doc"], size=SZ_SMALL, indent=2,
                                 align="left")
                        if ln["url"]:            # → the bill in QBO
                            dc1.hyperlink = ln["url"]
                            dc1.font = Font(size=SZ_SMALL, color=LINK,
                                            underline="single")
                        dc = _t(ws, r, 2, ln["date"], size=SZ_SMALL, align="left")
                        dc.number_format = "mm/dd/yyyy"
                        _t(ws, r, 3, str(ln["desc"])[:110], size=SZ_SMALL)
                        _t(ws, r, 6, ln["amt"], size=SZ_SMALL, fmt=MONEY_C,
                           align="right")
                        _t(ws, r, 7, ln["paid"], size=SZ_SMALL, color=GREY,
                           align="center")
                        ws.row_dimensions[r].outline_level = 2
                        ws.row_dimensions[r].hidden = True
                        r += 1
        # C carries the description and SPILLS across D:E (empty on data rows;
        # the metric tiles above are what keep those columns from reading as
        # gutters). Amount and Paid sit at the right edge, always aligned.
        for col, w in zip("ABCDEFG", (26, 13, 24, 24, 24, 20, 14)):
            ws.column_dimensions[col].width = w
        ws.freeze_panes = "A11"
        ws.sheet_properties.outlinePr.summaryBelow = False

    tmp = out.with_suffix(".tmp.xlsx")
    wb.save(str(tmp))
    assert_clean(tmp)
    for msg in lint_layout(wb):
        print(f"    ⚑ layout: {msg}")
    tmp.replace(out)


def lint_layout(wb) -> List[str]:
    """Catch the things that make a sheet READ badly but verify fine — the
    class of defect that shipped twice before this existed (the user
    2026-08-27: "why don't you inspect it after? it looks sloppy").

    Checks: a column left empty inside the used range (a gutter), and data
    rows whose first column is blank (a ragged left edge). Cell-level checks
    catch corruption; these catch ugly."""
    from openpyxl.utils import get_column_letter
    out: List[str] = []
    for ws in wb.worksheets:
        last = max((c.column for row in ws.iter_rows() for c in row
                    if c.value is not None), default=0)
        if not last:
            continue
        used = set()
        for row in ws.iter_rows(min_row=1, max_col=last):
            for c in row:
                if c.value is not None:
                    used.add(c.column)
        # A column covered by a MERGED range is not a gutter — only the
        # top-left cell of a merge carries the value.
        for mr in ws.merged_cells.ranges:
            used.update(range(mr.min_col, mr.max_col + 1))
        gutters = [get_column_letter(i) for i in range(1, last + 1) if i not in used]
        if gutters:
            out.append(f"{ws.title}: empty column(s) {gutters} inside A..{get_column_letter(last)}")
        blank_a = sum(1 for row in ws.iter_rows(min_row=6, max_col=last)
                      if row[0].value is None
                      and any(c.value is not None for c in row[1:]))
        if blank_a > 5:
            out.append(f"{ws.title}: {blank_a} rows have data but a blank column A")
    return out


def main() -> int:
    ap = argparse.ArgumentParser(description="Simple report for a finished job")
    ap.add_argument("jobs", nargs="*", help="e.g. MFD133")
    ap.add_argument("--all", action="store_true", help="every job in the archive folder")
    ap.add_argument("--folder", default=None)
    ap.add_argument("--bundle", action="store_true",
                    help="MFD Overview Total: ONE workbook with the overview "
                         "sheet plus a sheet per job (P&L + grouped "
                         "transactions). All links internal, so it survives "
                         "being emailed.")
    a = ap.parse_args()
    folder = (Path(a.folder).expanduser() if a.folder
              else pnl_paths.pnl_out_dir() / ARCHIVE)
    jobs = ([d.name for d in sorted(folder.iterdir()) if d.is_dir()]
            if a.all else [j.upper() for j in a.jobs])
    if not jobs:
        print("✗  name a job, or pass --all")
        return 1
    loaded = []
    for job in jobs:
        src_path = folder / job / f"Project_PnL_{job}.xlsx"
        if not src_path.exists():
            print(f"  ⚠ {job}: no {src_path.name}")
            continue
        src = read_source(src_path)
        if not src["sections"]:
            print(f"  ⚠ {job}: no 'By Account' sheet — regenerate it first")
            continue
        loaded.append((job, src, _totals(src)))
        if not a.bundle:
            out = folder / job / f"{job} Job Result.xlsx"
            build(job, src, out)
            print(f"  ✓ {job}  →  {out.name}")
    if a.bundle:
        if not loaded:
            print("✗  nothing to bundle")
            return 1
        out = folder / "MFD Overview Total.xlsx"
        build_bundle(loaded, out)
        tb = sum(t["billed"] for _, _, t in loaded)
        tc = sum(t["cost"] for _, _, t in loaded)
        print(f"  {len(loaded)} jobs   billed ${tb:,.0f}   cost ${tc:,.0f}   "
              f"GP ${tb - tc:,.0f} ({(tb - tc) / tb * 100 if tb else 0:.2f}%)")
        print(f"  → {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
