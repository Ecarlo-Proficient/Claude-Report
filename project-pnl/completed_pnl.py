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
OVERHEAD_PCT = 0.10

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


def main() -> int:
    ap = argparse.ArgumentParser(description="Simple report for a finished job")
    ap.add_argument("jobs", nargs="*", help="e.g. MFD133")
    ap.add_argument("--all", action="store_true", help="every job in the archive folder")
    ap.add_argument("--folder", default=None)
    a = ap.parse_args()
    folder = (Path(a.folder).expanduser() if a.folder
              else pnl_paths.pnl_out_dir() / ARCHIVE)
    jobs = ([d.name for d in sorted(folder.iterdir()) if d.is_dir()]
            if a.all else [j.upper() for j in a.jobs])
    if not jobs:
        print("✗  name a job, or pass --all")
        return 1
    for job in jobs:
        src_path = folder / job / f"Project_PnL_{job}.xlsx"
        if not src_path.exists():
            print(f"  ⚠ {job}: no {src_path.name}")
            continue
        src = read_source(src_path)
        if not src["sections"]:
            print(f"  ⚠ {job}: no 'By Account' sheet — regenerate it first")
            continue
        out = folder / job / f"{job} Job Result.xlsx"
        build(job, src, out)
        print(f"  ✓ {job}  →  {out.name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
