#!/usr/bin/env python3
"""
bizdev_cut_view.py - the owner's INTERNAL division P&L with the director's cut
charged at the end, in the exact shape of the division Overview, with every
cut line linked to its transaction in QuickBooks.

WHAT IT IS
The division Overview (`completed_pnl.py`) is the shared report: one row per
job, contract / billed / cost / gross profit, then the 10% and 9% overhead
views. It never shows what the outside parties in the cut register
(`shared/bizdev_cut.py`) were paid, because PMs read it. This workbook is
the same page for the OWNER ONLY, with the director's cut added at the end of
every row and the REAL NET after it at both overhead rates (the owner
2026-09-11: "exactly how it is but add the director's cut at the end to see
the final net profit"). Click a job and its sheet lists every line of the cut -
date, reference, what the line says, the account, how it was tied to the job -
each reference a deep link into QuickBooks, so the question "how did his money
get coded to that job?" is answered by the transaction itself, not by us.

STRAIGHT DIRECTOR ONLY (the owner 2026-09-11: "put jordan's cut in the costs of
the project instead, we want straight director only"). The register names the
director; every OTHER registered vendor's cut is charged into that job's COST
on this page, and the job sheet shows those lines under "in job cost" so the
owner can still see them. No overhead model lives here - the rate ladder and
the overhead review are gone from this file (the review is its own workbook,
`<DIV> OH Calculations.xlsx`).

WHERE THE NUMBERS COME FROM
  * contract / billed / cost: READ OUT OF THE GENERATED P&L WORKBOOKS through
    `completed_pnl.read_source` + `_totals` - the same reader the Overview
    uses, so this page can never disagree with it. Billed is gross of
    retainage, exactly as the P&L and the Overview carry it.
  * the cut: one QBO pull of every Bill, Check and Expense on each registered
    vendor (the TransactionList report by vendor names them; a Purchase cannot
    be queried by payee, so the checks and expenses are fetched by id). A line
    is a CUT or "they fronted it" by `bizdev_cut.is_cut` - the one test.
  * the job a cut line belongs to, in order: the QBO project on the line; the
    job named in the line text (`MFD177`, or the builder draw code `177-0-20-1`
    whose LEADING number is the job - see the vault's rosetta); the bill memo.
    A line whose text names several jobs is split equally between them, and
    the sheet says so. Dates are stripped before matching so `03/10/2023`
    never reads as a job.
  * a cut line the P&L workbook still carries in job cost is stripped out of
    COST here (the register test again), so every cut is counted ONCE: the
    director's in his column, the others' inside COST.

LOCAL ONLY. The file lives in <CompanyHealth>, never on OneDrive, Teams or in
the repo (the owner 2026-09-04: "do not put anywhere just local"). It is
rebuilt from scratch every run.

USAGE
  python3 project-pnl/bizdev_cut_view.py                 # MFD, default file
  python3 project-pnl/bizdev_cut_view.py --division mfd --out <path.xlsx>
"""
from __future__ import annotations

import argparse
import datetime as dt
import re
import sys
from pathlib import Path
from typing import Dict, List, Tuple

from openpyxl import Workbook
from openpyxl.styles import Border, Font
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.path.insert(0, str(Path(__file__).resolve().parent))
from shared import bizdev_cut, paths, pnl_paths, qbo_api          # noqa: E402
from shared.xlsx_verify import assert_clean                       # noqa: E402
import completed_pnl as cp                                        # noqa: E402
from completed_pnl import (C0, F_BAND, F_HDR, GREEN, GREY, GUTTER_W, HAIR,   # noqa: E402
                           INK, LINK, MFD_OVERHEAD_PCT, MONEY, MONEY_C, NAVY,
                           OVERHEAD_PCT, PCT, RED, RULE, SZ, SZ_SMALL,
                           SZ_TITLE, THICK, _spans, _t, _thick_box, _tiles,
                           job_label, lint_layout)

OUT_NAME = "{label} PnL - Internal - Director Cut.xlsx"
PNL_SHEET = "{label} P&L"
OLDER_SHEET = "Older {label} jobs"
NOJOB_SHEET = "No job named"
DETAIL_W = (13, 12, 22, 60, 30, 15, 18, 40)      # the line grid, B..I

# ─────────────────────────── the cut, from QBO ───────────────────────────
_DATE = re.compile(r"\b\d{1,2}/\d{1,2}(?:/\d{2,4})?\b")
_CODE = re.compile(r"(?<![\d-])(\d{3})-\d+-\d+-\d+\b")       # builder draw code
_JOB = re.compile(r"\b(MFD|CP|RP)\s?(\d{3,4})\b", re.I)


def _jobs_in(text: str, prefix: str) -> Tuple[set, str]:
    """(jobs named in `text`, how) - the literal job number or the draw code
    whose leading number is the job. Dates go first so they cannot match."""
    t = _DATE.sub(" ", text or "")
    lit = {f"{p.upper()}{n}" for p, n in _JOB.findall(t) if p.upper() == prefix}
    codes = {f"{prefix}{n}" for n in _CODE.findall(t)}
    if lit and not codes:
        return lit, "job number on the line"
    if codes:
        return lit | codes, "draw code on the line (leading number = the job)"
    return set(), ""


def _txn_url(kind: str, txn_id: str, realm: str) -> str:
    """The transaction in QBO - the same login+deeplink form the P&Ls use."""
    return (f"https://qbo.intuit.com/app/login?pagereq={kind}?txnId={txn_id}"
            f"&deeplinkcompanyid={realm}")


def _vendor_txn_ids(access: str, realm: str, vendor_id: str) -> Dict[str, List[str]]:
    """{'Bill': [...ids], 'Purchase': [...ids]} for one vendor, off the
    TransactionList report (a Purchase cannot be queried by payee)."""
    rep = qbo_api.report(access, realm, "TransactionList", {
        "start_date": "2015-01-01", "end_date": "2099-12-31",
        "vendor": vendor_id, "columns": "tx_date,txn_type,doc_num,name"})
    out: Dict[str, List[str]] = {"Bill": [], "Purchase": []}

    def walk(node):
        for row in node.get("Rows", {}).get("Row", []):
            if "Rows" in row:
                walk(row)
                continue
            cd = row.get("ColData", [])
            if len(cd) < 2:
                continue
            ttype, tid = cd[1].get("value") or "", cd[1].get("id") or ""
            if not tid:
                continue
            if ttype == "Bill":
                out["Bill"].append(tid)
            elif ttype in ("Check", "Expense", "Credit Card Expense", "Cash Expense"):
                out["Purchase"].append(tid)
    walk(rep)
    return out


def pull_cut_lines(prefix: str) -> Tuple[List[dict], str]:
    """Every line of every Bill / Check / Expense on the registered vendors,
    already tested (cut vs fronted), tied to a job, and marked director or
    not. Returns (lines, realm)."""
    reg = bizdev_cut.load()
    if not reg["vendors"]:
        raise SystemExit("no vendors in the cut register - nothing to build")
    access, realm = qbo_api.load_credentials()
    vendors: Dict[str, str] = {}
    for key in reg["vendors"]:
        for v in qbo_api.query_all(access, realm, "Vendor", f"DisplayName LIKE '%{key.title()}%'"):
            vendors[v["Id"]] = v["DisplayName"]
        for v in qbo_api.query_all(access, realm, "Vendor", f"DisplayName LIKE '%{key}%'"):
            vendors[v["Id"]] = v["DisplayName"]
    if not vendors:
        raise SystemExit("none of the registered vendors exists in QBO")
    lines: List[dict] = []
    for vid, vname in vendors.items():
        ids = _vendor_txn_ids(access, realm, vid)
        bills = qbo_api.query_all(access, realm, "Bill", f"VendorRef = '{vid}'")
        purchases = [qbo_api._api_get(f"/v3/company/{realm}/purchase/{pid}", access)
                     .get("Purchase") for pid in ids["Purchase"]]
        n_b = len(bills)
        for txn in bills + [p for p in purchases if p]:
            kind = "bill" if n_b > 0 else "expense"
            n_b -= 1
            if kind == "expense" and str(txn.get("PaymentType") or "") == "Check":
                kind = "check"
            memo = str(txn.get("PrivateNote") or "")
            memo_jobs, memo_how = _jobs_in(memo, prefix)
            for ln in txn.get("Line") or []:
                det = (ln.get("AccountBasedExpenseLineDetail")
                       or ln.get("ItemBasedExpenseLineDetail"))
                if not det:
                    continue
                amt = float(ln.get("Amount") or 0)
                desc = str(ln.get("Description") or "")
                acct = ((det.get("AccountRef") or {}).get("name")
                        or (det.get("ItemRef") or {}).get("name") or "")
                cust = str((det.get("CustomerRef") or {}).get("name") or "")
                m = _JOB.search(cust.replace(" ", ""))
                if m and m.group(1).upper() == prefix:
                    jobs, how = {f"{prefix}{m.group(2)}"}, "coded to the job in QuickBooks"
                else:
                    jobs, how = _jobs_in(desc, prefix)
                    if not jobs:
                        jobs, how = memo_jobs, (f"bill memo: {memo_how}" if memo_jobs else "")
                n = max(1, len(jobs))
                if n > 1:
                    how += f" - names {n} jobs, split {n} ways"
                cut = bizdev_cut.is_cut(vname, acct, desc or memo)
                for job in (sorted(jobs) or [""]):
                    lines.append({
                        "job": job, "vendor": vname, "label": bizdev_cut.label(vname),
                        "director": bizdev_cut.is_director(vname),
                        "date": str(txn.get("TxnDate") or ""),
                        "ref": str(txn.get("DocNumber") or "").strip() or
                               {"bill": "bill", "check": "check", "expense": "ACH / card"}[kind],
                        "url": _txn_url(kind, str(txn["Id"]), realm),
                        "text": desc or memo, "acct": acct,
                        "amt": amt / n, "cut": cut,
                        "how": how or "no job named on the line or the bill",
                    })
    return lines, realm


# ─────────────────────────── the workbook ───────────────────────────

def _strip_cut_from_cost(src: dict) -> float:
    """A cut line the P&L workbook still carries in job cost comes OUT of the
    sections here, so every cut is counted once. Returns what was removed."""
    removed = 0.0
    for sec in src["sections"]:
        for acct in sec["accounts"]:
            keep_v = []
            for v in acct["vendors"]:
                kept = [ln for ln in v["lines"]
                        if not bizdev_cut.is_cut(v["name"], acct["name"], ln["desc"])]
                removed += sum(ln["amt"] for ln in v["lines"]) - sum(ln["amt"] for ln in kept)
                if kept:
                    v["lines"] = kept
                    v["total"] = round(sum(ln["amt"] for ln in kept), 2)
                    keep_v.append(v)
            acct["vendors"] = keep_v
            acct["total"] = round(sum(v["total"] for v in keep_v), 2)
        sec["accounts"] = [a for a in sec["accounts"] if a["vendors"]]
        sec["total"] = round(sum(a["total"] for a in sec["accounts"]), 2)
    return round(removed, 2)


def _label(job: str, title: str) -> str:
    """The Overview's job label, with its em dash swapped for a hyphen - nothing
    this tool writes carries one (the owner's standing rule)."""
    return job_label(job, title).replace(" — ", " - ").replace("—", "-")


def _hdr(ws, r, c, text, align="right", indent=0):
    return _t(ws, r, c, text, size=SZ_SMALL - 1, bold=True, color="FFFFFF",
              fill=F_HDR, align=align, wrap=True, indent=indent)


def _who(ln: dict) -> str:
    """The Who column: the director's label as is; anyone else's cut is
    marked as job cost, which is where this page puts it."""
    lab = ln["label"] or ln["vendor"]
    return lab if ln["director"] or not ln["cut"] else f"{lab} (in job cost)"


def _lines_table(ws, r: int, rows: List[dict], dlabel: str, title: str,
                 first_col: int = C0) -> Tuple[int, str, str, str]:
    """The cut lines, one grid: ref (linked) · date · who · what it says ·
    account · the director's cut · in job cost · how it was tied to the job.
    Returns (next row, director total cell, other-cut-in-cost total cell,
    fronted total cell)."""
    B = first_col
    _t(ws, r, B, title, size=SZ + 2, bold=True, color=NAVY)
    r += 1
    heads = (("Ref #", "left"), ("Date", "left"), ("Who", "left"),
             ("What the line says", "left"), ("Account", "left"),
             (dlabel, "right"), ("In job cost", "right"),
             ("How it was tied to this job", "left"))
    for i, (h, al) in enumerate(heads):
        _hdr(ws, r, B + i, h, align=al, indent=1 if i == 0 else 0)
    ws.row_dimensions[r].height = 30
    r += 1
    first = r
    for ln in sorted(rows, key=lambda x: (x["date"], x["ref"]), reverse=True):
        c1 = _t(ws, r, B, ln["ref"], size=SZ_SMALL, align="left", indent=1)
        if ln["url"]:
            c1.hyperlink = ln["url"]
            c1.font = Font(size=SZ_SMALL, color=LINK, underline="single")
        d = _t(ws, r, B + 1, dt.date.fromisoformat(ln["date"]) if ln["date"] else "",
               size=SZ_SMALL, align="left")
        d.number_format = "mm/dd/yyyy"
        _t(ws, r, B + 2, _who(ln), size=SZ_SMALL)
        _t(ws, r, B + 3, ln["text"][:120], size=SZ_SMALL)
        _t(ws, r, B + 4, ln["acct"][:40], size=SZ_SMALL)
        his = ln["cut"] and ln["director"]
        _t(ws, r, B + 5, ln["amt"] if his else None, size=SZ_SMALL, fmt=MONEY_C, align="right")
        _t(ws, r, B + 6, None if his else ln["amt"], size=SZ_SMALL, fmt=MONEY_C,
           align="right", color=GREY)
        _t(ws, r, B + 7, ln["how"], size=SZ_SMALL - 1, color=GREY)
        if r % 2 == 0:
            for c in range(B, B + 8):
                ws.cell(row=r, column=c).fill = F_BAND
        ws.row_dimensions[r].height = 20
        r += 1
    last = r - 1
    if not rows:
        _t(ws, r, B, "nothing from them on this job", size=SZ_SMALL, color=GREY, indent=1)
        r += 1
        last = first          # an empty SUM range still has to be a real range
    who, cut, inc = (get_column_letter(B + 2), get_column_letter(B + 5),
                     get_column_letter(B + 6))
    for c in range(B, B + 8):
        ws.cell(row=r, column=c).border = Border(top=RULE)
    _t(ws, r, B, f"{dlabel} on this job", size=SZ, bold=True, color=NAVY, indent=1)
    d_cell = _t(ws, r, B + 5, f"=SUM({cut}{first}:{cut}{last})", size=SZ, bold=True,
                color=NAVY, fmt=MONEY, align="right")
    ws.row_dimensions[r].height = 24
    r += 1
    _t(ws, r, B, "other cuts charged into this job's COST", size=SZ_SMALL, color=GREY, indent=1)
    o_cell = _t(ws, r, B + 6, f'=SUMIF({who}{first}:{who}{last},"*(in job cost)",'
                              f'{inc}{first}:{inc}{last})',
                size=SZ_SMALL, bold=True, color=GREY, fmt=MONEY, align="right")
    r += 1
    _t(ws, r, B, "fronted for the job - already in COST, not a cut", size=SZ_SMALL,
       color=GREY, indent=1)
    f_cell = _t(ws, r, B + 6, f"=SUM({inc}{first}:{inc}{last})-{o_cell.coordinate}",
                size=SZ_SMALL, bold=True, color=GREY, fmt=MONEY, align="right")
    r += 2
    return r, d_cell.coordinate, o_cell.coordinate, f_cell.coordinate


def _page_setup(ws):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True


def _detail_sheet(wb, name: str, title: str, pnl_name: str):
    ws = wb.create_sheet(name[:31])
    ws.sheet_view.showGridLines = False
    _t(ws, 1, 1, title, size=SZ_TITLE - 2, bold=True, color=NAVY)
    back = _t(ws, 1, C0 + 7, f"← back to {pnl_name}", size=SZ_SMALL, align="right")
    back.hyperlink = f"#'{pnl_name}'!A1"
    back.font = Font(size=SZ_SMALL, color=LINK, underline="single")
    for c in range(1, C0 + 8):
        ws.cell(row=2, column=c).border = Border(bottom=HAIR)
    ws.row_dimensions[1].height = 30
    ws.row_dimensions[3].height = 8
    ws.column_dimensions["A"].width = GUTTER_W
    for col, w in zip("BCDEFGHI", DETAIL_W):
        ws.column_dimensions[col].width = w
    _page_setup(ws)
    return ws


def build(jobs: List[tuple], cut_lines: List[dict], out: Path, div: dict) -> dict:
    """jobs = [(job, src, totals, src_path)] from completed_pnl.load_division,
    with every registered cut already stripped from each src. Writes `out`
    from scratch."""
    label = div["label"]
    pnl_name = PNL_SHEET.format(label=label)
    older_name = OLDER_SHEET.format(label=label)
    dlabel = next((ln["label"] for ln in cut_lines if ln["director"] and ln["label"]),
                  "DIRECTOR CUT")
    dword = dlabel.replace(" CUT", "").strip().title() or "Director"
    by_job: Dict[str, List[dict]] = {}
    for ln in cut_lines:
        by_job.setdefault(ln["job"], []).append(ln)
    known = {j for j, _s, _t2, _p in jobs}

    wb = Workbook()
    sm = wb.active
    sm.title = pnl_name
    sm.sheet_view.showGridLines = False

    # ── columns: the Overview's, then the director's cut at the end ──
    cols = [("CONTRACT", "contract", MONEY, 18),
            ("BILLED", "billed", MONEY, 18), ("COST", "cost", MONEY, 18),
            ("GROSS PROFIT", "gp", MONEY, 18), ("GP %", "gpm", PCT, 11),
            (f"{OVERHEAD_PCT:.0%} OH", "oh", MONEY, 16),
            (f"FINAL NET  ({OVERHEAD_PCT:.0%} OH)", "net", MONEY, 20),
            (f"{MFD_OVERHEAD_PCT:.0%} OH", "moh", MONEY, 16),
            (f"FINAL NET  ({MFD_OVERHEAD_PCT:.0%} OH)", "mnet", MONEY, 20),
            (dlabel, "cut", MONEY, 17),
            (f"REAL NET  ({OVERHEAD_PCT:.0%} OH)", "rnet", MONEY, 20),
            (f"REAL NET  ({MFD_OVERHEAD_PCT:.0%} OH)", "rmnet", MONEY, 20),
            (f"REAL NET %  ({MFD_OVERHEAD_PCT:.0%} OH)", "rmnetm", PCT, 13)]
    keys = [k for _h, k, _f, _w in cols]
    L = {k: get_column_letter(C0 + 1 + i) for i, k in enumerate(keys)}
    LAST = C0 + len(cols) + 1
    first_cut_col = C0 + 1 + keys.index("cut")

    def _formula(k, rr, ranges=None):
        if k in ("contract", "billed", "cost", "cut"):
            if not ranges:
                return None
            return "=" + "+".join(f"SUM({L[k]}{a}:{L[k]}{b})" for a, b in ranges)
        if k == "gp":
            return f"={L['billed']}{rr}-{L['cost']}{rr}"
        if k == "gpm":
            return f'=IF({L["billed"]}{rr}=0,"",{L["gp"]}{rr}/{L["billed"]}{rr})'
        if k == "oh":
            return f"={L['contract']}{rr}*{OVERHEAD_PCT}"
        if k == "net":
            return f"={L['gp']}{rr}-{L['oh']}{rr}"
        if k == "moh":
            return f"={L['contract']}{rr}*{MFD_OVERHEAD_PCT}"
        if k == "mnet":
            return f"={L['gp']}{rr}-{L['moh']}{rr}"
        if k == "rnet":
            return f"={L['net']}{rr}-{L['cut']}{rr}"
        if k == "rmnet":
            return f"={L['mnet']}{rr}-{L['cut']}{rr}"
        if k == "rmnetm":
            return f'=IF({L["billed"]}{rr}=0,"",{L["rmnet"]}{rr}/{L["billed"]}{rr})'
        raise KeyError(k)

    # python-side figures for the colours (the cells themselves are formulas).
    # COST here = the workbook's cost + the OTHER vendors' cut on the job.
    fig: Dict[str, dict] = {}
    for job, src, t, _p in jobs:
        mine = by_job.get(job, [])
        other = sum(ln["amt"] for ln in mine if ln["cut"] and not ln["director"])
        d = dict(t)
        d["base_cost"] = t["cost"]
        d["other"] = other
        d["cost"] = t["cost"] + other
        d["gp"] = d["billed"] - d["cost"]
        d["gpm"] = d["gp"] / d["billed"] if d["billed"] else 0.0
        d["net"], d["mnet"] = d["gp"] - d["oh"], d["gp"] - d["moh"]
        d["cut"] = sum(ln["amt"] for ln in mine if ln["cut"] and ln["director"])
        d["rnet"], d["rmnet"] = d["net"] - d["cut"], d["mnet"] - d["cut"]
        d["rmnetm"] = d["rmnet"] / d["billed"] if d["billed"] else 0.0
        fig[job] = d

    def _sum(sel):
        d = {k: sum(fig[j][k] for j, _s, _t2, _p in sel)
             for k in ("contract", "billed", "cost", "gp", "oh", "net", "moh", "mnet",
                       "cut", "rnet", "rmnet", "other")}
        for a, b in (("gpm", "gp"), ("netm", "net"), ("mnetm", "mnet"), ("rmnetm", "rmnet")):
            d[a] = d[b] / d["billed"] if d["billed"] else 0
        return d

    tot = _sum(jobs)
    n_act = sum(1 for _j, s, _t2, _p in jobs if s.get("status") == "Active")
    _t(sm, 1, 1, f"{label} P&L - INTERNAL - THE {dword.upper()}'S CUT AT THE END",
       size=SZ_TITLE, bold=True, color=NAVY)
    _t(sm, 2, C0, f"{n_act} active · {len(jobs) - n_act} completed · the Overview's page "
                  f"with the {dword.lower()}'s cut charged after overhead · click a job for "
                  f"every line with its QuickBooks link · internal only",
       size=SZ_SMALL, color=GREY)
    _t(sm, 1, LAST, f"Generated {dt.datetime.now():%m/%d/%Y %I:%M %p}", size=SZ_SMALL,
       color=GREY, align="right")
    for c in range(1, LAST + 1):
        sm.cell(row=2, column=c).border = Border(bottom=HAIR)
    sm.row_dimensions[1].height = 30
    sm.row_dimensions[3].height = 8

    STRIP_ROWS = 8                       # tiles (2) + gap + box header + 2 views + gap
    r = 4 + STRIP_ROWS
    job_rows: Dict[str, int] = {}
    sec_ranges: list = []
    _hdr(sm, r, C0, "JOB", align="left", indent=1)
    for i, (h, _k, _f, _w) in enumerate(cols):
        _hdr(sm, r, C0 + 1 + i, h)
    _hdr(sm, r, LAST, "P&L FILE", align="center")
    sm.row_dimensions[r].height = 30
    hdr_row = r
    r += 1

    def _row_figures(rr, t, bold_keys=("gp", "mnet", "rmnet"), ranges=None):
        for i, (_h, k, fmt, _w) in enumerate(cols):
            pos = k in ("gp", "gpm", "net", "mnet", "rnet", "rmnet", "rmnetm")
            f = _formula(k, rr, ranges)
            _t(sm, rr, C0 + 1 + i, f if f is not None else t[k], size=SZ, fmt=fmt,
               align="right", bold=k in bold_keys,
               color=(GREEN if t[k] >= 0 else RED) if pos else INK)

    def _job_row(job, src, t):
        nonlocal r
        cell = _t(sm, r, C0, _label(job, src.get("title", "")), size=SZ, bold=True)
        cell.hyperlink = f"#'{job}'!A1"
        cell.font = Font(size=SZ, bold=True, color=LINK, underline="single")
        job_rows[job] = r
        _row_figures(r, t)
        if src.get("rel"):
            lk = _t(sm, r, LAST, "open P&L  ↗", size=SZ_SMALL, color=LINK)
            lk.hyperlink = src["rel"]
            lk.font = Font(size=SZ_SMALL, color=LINK, underline="single")
        if r % 2 == 0:
            for c in range(C0, LAST + 1):
                sm.cell(row=r, column=c).fill = F_BAND
        sm.row_dimensions[r].height = 22
        r += 1

    def _section(title, sel, note=""):
        nonlocal r
        if not sel:
            return
        _t(sm, r, C0, title, size=SZ, bold=True, color="FFFFFF", fill=F_HDR)
        if note:
            _t(sm, r, C0 + 1, note, size=SZ_SMALL - 1, color="FFFFFF", fill=F_HDR)
        for c in range(C0, LAST + 1):
            sm.cell(row=r, column=c).fill = F_HDR
        sm.row_dimensions[r].height = 22
        r += 1
        first = r
        for j, s, _t2, _p in sorted(sel, key=lambda x: -fig[x[0]]["billed"]):
            _job_row(j, s, fig[j])
        sec_ranges.append((first, r - 1))
        _t(sm, r, C0, f"subtotal - {len(sel)} job(s)", size=SZ, bold=True, color=NAVY)
        _row_figures(r, _sum(sel), bold_keys=tuple(keys), ranges=[(first, r - 1)])
        for c in range(C0, LAST + 1):
            sm.cell(row=r, column=c).border = Border(top=HAIR)
        r += 2

    active = [j for j in jobs if j[1].get("status") == "Active"]
    done = [j for j in jobs if j[1].get("status") != "Active"]
    _section("ACTIVE - in progress", active, "costs to date only - not finished")
    _section("COMPLETED", done)
    _t(sm, r, C0, f"ALL {label} - {len(jobs)} JOBS", size=SZ, bold=True, color=NAVY)
    _row_figures(r, tot, bold_keys=tuple(keys), ranges=sec_ranges)
    for c in range(C0, LAST + 1):
        sm.cell(row=r, column=c).border = Border(top=RULE)
    sm.row_dimensions[r].height = 24
    all_row = r
    # a heavy rule down the left of the cut block, so the added columns read
    # as their own thing against the Overview's page
    for rr in range(hdr_row, all_row + 1):
        cur = sm.cell(row=rr, column=first_cut_col).border
        sm.cell(row=rr, column=first_cut_col).border = Border(
            left=THICK, right=cur.right, top=cur.top, bottom=cur.bottom)
    r += 2

    # ── the strip at the top: tiles, then one box with both views ──
    spans4 = _spans(C0, LAST, 4)
    A = lambda span, row: f"{get_column_letter(span[0])}{row}"          # noqa: E731
    _tiles(sm, 4, [
        ("BILLED", f"={L['billed']}{all_row}", MONEY, NAVY),
        ("COST", f"={L['cost']}{all_row}", MONEY, NAVY),
        ("GROSS PROFIT", f"={A(spans4[0], 5)}-{A(spans4[1], 5)}", MONEY,
         GREEN if tot["gp"] >= 0 else RED),
        ("GROSS MARGIN", f'=IF({A(spans4[0], 5)}=0,"",{A(spans4[2], 5)}/{A(spans4[0], 5)})',
         PCT, GREEN if tot["gp"] >= 0 else RED)], spans4)
    spans7 = _spans(C0, LAST, 7)
    r2 = 7
    heads = ("AFTER OVERHEAD, THEN THE CUT", "OVERHEAD", "FINAL NET PROFIT", "NET MARGIN",
             dlabel, "REAL NET PROFIT", "REAL MARGIN")
    for span, txt in zip(spans7, heads):
        sm.merge_cells(start_row=r2, start_column=span[0], end_row=r2, end_column=span[1])
        _t(sm, r2, span[0], txt, size=SZ_SMALL - 1, bold=True, color=GREY,
           align="left" if span is spans7[0] else "right", indent=1 if span is spans7[0] else 0)
    for cc in range(spans7[0][0], spans7[-1][1] + 1):
        sm.cell(row=r2, column=cc).border = Border(bottom=HAIR)
    views = ((f"{OVERHEAD_PCT:.0%} OH", "oh", "net", "rnet", tot["net"], tot["rnet"]),
             (f"{MFD_OVERHEAD_PCT:.0%} OH", "moh", "mnet", "rmnet", tot["mnet"], tot["rmnet"]))
    for i, (lbl, ohk, netk, rk, netv, rv) in enumerate(views):
        rr = r2 + 1 + i
        sm.merge_cells(start_row=rr, start_column=spans7[0][0], end_row=rr,
                       end_column=spans7[0][1])
        _t(sm, rr, spans7[0][0], lbl, size=SZ, bold=(i == 1), color=INK, indent=1)
        vals = ((f"=-{L[ohk]}{all_row}", MONEY, GREY),
                (f"={L[netk]}{all_row}", MONEY, GREEN if netv >= 0 else RED),
                (f'=IF({L["billed"]}{all_row}=0,"",{L[netk]}{all_row}/{L["billed"]}{all_row})',
                 PCT, GREEN if netv >= 0 else RED),
                (f"=-{L['cut']}{all_row}", MONEY, GREY),
                (f"={L[rk]}{all_row}", MONEY, GREEN if rv >= 0 else RED),
                (f'=IF({L["billed"]}{all_row}=0,"",{L[rk]}{all_row}/{L["billed"]}{all_row})',
                 PCT, GREEN if rv >= 0 else RED))
        for span, (val, fmt, col) in zip(spans7[1:], vals):
            sm.merge_cells(start_row=rr, start_column=span[0], end_row=rr, end_column=span[1])
            _t(sm, rr, span[0], val, size=SZ + 2, bold=True, fmt=fmt, color=col, align="right")
        sm.row_dimensions[rr].height = 26
        if i:
            for cc in range(spans7[0][0], spans7[-1][1] + 1):
                cur = sm.cell(row=rr, column=cc).border
                sm.cell(row=rr, column=cc).border = Border(
                    left=cur.left, right=cur.right, bottom=cur.bottom, top=THICK)
    _thick_box(sm, r2, r2 + 2, spans7[0][0], spans7[-1][1])

    # ── per-job sheets, then the two catch-alls ──
    job_cells: Dict[str, Tuple[str, str]] = {}
    for job, src, _t2, _p in sorted(jobs, key=lambda x: -fig[x[0]]["billed"]):
        ws = _detail_sheet(wb, job, _label(job, src.get("title", "")), pnl_name)
        who = src["title"].replace("PROJECT P&L — ", "").replace("—", "-")
        _t(ws, 2, C0, f"{src.get('status', 'Completed').lower()} job · {who} · click a ref # "
                      f"to open that transaction in QuickBooks", size=SZ_SMALL, color=GREY)
        JLAST = C0 + 7
        jr = job_rows[job]
        P = f"'{pnl_name}'!"
        f = fig[job]
        _tiles(ws, 4, [
            ("BILLED", f"={P}{L['billed']}{jr}", MONEY, NAVY),
            ("COST", f"={P}{L['cost']}{jr}", MONEY, NAVY),
            ("GROSS PROFIT", f"={P}{L['gp']}{jr}", MONEY, GREEN if f["gp"] >= 0 else RED),
            ("GROSS MARGIN", f"={P}{L['gpm']}{jr}", PCT, GREEN if f["gp"] >= 0 else RED)],
            _spans(C0, JLAST, 4))
        sp7 = _spans(C0, JLAST, 7)
        for span, txt in zip(sp7, heads):
            ws.merge_cells(start_row=7, start_column=span[0], end_row=7, end_column=span[1])
            _t(ws, 7, span[0], txt, size=SZ_SMALL - 1, bold=True, color=GREY,
               align="left" if span is sp7[0] else "right", indent=1 if span is sp7[0] else 0)
        for cc in range(sp7[0][0], sp7[-1][1] + 1):
            ws.cell(row=7, column=cc).border = Border(bottom=HAIR)
        for i, (lbl, ohk, netk, rk, _nv, _rv) in enumerate(views):
            rr = 8 + i
            ws.merge_cells(start_row=rr, start_column=sp7[0][0], end_row=rr, end_column=sp7[0][1])
            _t(ws, rr, sp7[0][0], lbl, size=SZ, bold=(i == 1), color=INK, indent=1)
            vals = ((f"=-{P}{L[ohk]}{jr}", MONEY, GREY),
                    (f"={P}{L[netk]}{jr}", MONEY, GREEN if f[netk] >= 0 else RED),
                    (f'=IF({P}{L["billed"]}{jr}=0,"",{P}{L[netk]}{jr}/{P}{L["billed"]}{jr})',
                     PCT, GREEN if f[netk] >= 0 else RED),
                    (f"=-{P}{L['cut']}{jr}", MONEY, GREY),
                    (f"={P}{L[rk]}{jr}", MONEY, GREEN if f[rk] >= 0 else RED),
                    (f'=IF({P}{L["billed"]}{jr}=0,"",{P}{L[rk]}{jr}/{P}{L["billed"]}{jr})',
                     PCT, GREEN if f[rk] >= 0 else RED))
            for span, (val, fmt, col) in zip(sp7[1:], vals):
                ws.merge_cells(start_row=rr, start_column=span[0], end_row=rr,
                               end_column=span[1])
                _t(ws, rr, span[0], val, size=SZ + 2, bold=True, fmt=fmt, color=col,
                   align="right")
            ws.row_dimensions[rr].height = 26
        _thick_box(ws, 7, 9, sp7[0][0], sp7[-1][1])
        _r, d_cell, o_cell, _fc = _lines_table(
            ws, 12, by_job.get(job, []), dlabel, "EVERY LINE - newest first")
        job_cells[job] = (d_cell, o_cell)

    # the P&L page's cut and COST cells POINT AT the job sheets' totals, so a
    # figure and the lines behind it can never disagree
    for job, (d_cell, o_cell) in job_cells.items():
        rr = job_rows[job]
        sm.cell(row=rr, column=C0 + 1 + keys.index("cut")).value = f"='{job}'!{d_cell}"
        base = round(fig[job]["base_cost"], 2)
        sm.cell(row=rr, column=C0 + 1 + keys.index("cost")).value = \
            f"={base}+'{job}'!{o_cell}" if fig[job]["other"] else base

    older = {j: v for j, v in by_job.items() if j and j not in known}
    nojob = by_job.get("", [])
    ws = _detail_sheet(wb, older_name, f"OLDER {label} JOBS - no P&L workbook on file", pnl_name)
    rr = 4
    older_cells: List[str] = []
    for job in sorted(older, key=lambda j: -sum(x["amt"] for x in older[j] if x["cut"])):
        rr, d_cell, _o, _f = _lines_table(ws, rr, older[job], dlabel, job)
        older_cells.append(d_cell)
    ws = _detail_sheet(wb, NOJOB_SHEET, "NO JOB NAMED ON THE LINE OR THE BILL", pnl_name)
    _rr, nojob_cell, _o, _f = _lines_table(
        ws, 4, nojob, dlabel, "retainers, estimating, pay periods - nothing ties them to a job")

    # ── every dollar paid to the director, reconciled on the page ──
    his = [ln for ln in cut_lines if ln["director"]]
    paid = sum(ln["amt"] for ln in his)
    dates = sorted(ln["date"] for ln in his if ln["date"])
    _mdy = lambda iso: f"{iso[5:7]}/{iso[8:10]}/{iso[:4]}" if len(iso) >= 10 else iso  # noqa: E731
    _t(sm, r, C0, f"EVERY DOLLAR PAID TO THE {dword.upper()} - {_mdy(dates[0]) if dates else ''} "
                  f"to {_mdy(dates[-1]) if dates else ''} - where it went", size=SZ + 1,
       bold=True, color=NAVY)
    r += 1
    # the label spills across C:F, the amount sits in G - nothing gets clipped
    AMT = C0 + 5
    recon_first = r
    rows_ = [
        (f"on these {len(jobs)} jobs (the {dlabel} column above)", f"={L['cut']}{all_row}", None),
        (f"on older {label} jobs with no P&L on file",
         "=" + ("+".join(f"'{older_name}'!{c}" for c in older_cells) or "0"),
         f"#'{older_name}'!A1"),
        ("no job named on the line or the bill", f"='{NOJOB_SHEET}'!{nojob_cell}",
         f"#'{NOJOB_SHEET}'!A1"),
        ("fronted for the jobs - inside COST, not a cut",
         round(sum(ln["amt"] for ln in his if not ln["cut"]), 2), None),
    ]
    for lbl, val, link in rows_:
        c1 = _t(sm, r, C0, lbl, size=SZ, indent=1)
        if link:
            c1.hyperlink = link
            c1.font = Font(size=SZ, color=LINK, underline="single")
        _t(sm, r, AMT, val, size=SZ, fmt=MONEY, align="right", color=INK)
        r += 1
    rec_last = r - 1
    _t(sm, r, C0, f"TOTAL PAID TO THE {dword.upper()} - every bill, check and expense in QuickBooks",
       size=SZ, bold=True, color=NAVY, indent=1)
    _t(sm, r, AMT, round(paid, 2), size=SZ, bold=True, fmt=MONEY, align="right", color=NAVY)
    for c in range(C0, AMT + 1):
        sm.cell(row=r, column=c).border = Border(top=RULE)
    tot_cell = f"{get_column_letter(AMT)}{r}"
    r += 1
    _t(sm, r, C0, "check - the four lines above less the total (must be zero)", size=SZ_SMALL,
       color=GREY, indent=1)
    _t(sm, r, AMT, f"=SUM({get_column_letter(AMT)}{recon_first}:{get_column_letter(AMT)}{rec_last})"
                   f"-{tot_cell}", size=SZ_SMALL, fmt=MONEY_C, align="right", color=GREY)
    r += 2
    if tot["other"]:
        _t(sm, r, C0, f"COST above includes {tot['other']:,.0f} of other business-development "
                      f"cuts charged into the jobs - see 'in job cost' on each job sheet",
           size=SZ_SMALL, color=GREY, indent=1)
        r += 1

    sm.column_dimensions["A"].width = GUTTER_W
    sm.column_dimensions[get_column_letter(C0)].width = 34
    for i, (_h, _k, _f, w) in enumerate(cols):
        sm.column_dimensions[get_column_letter(C0 + 1 + i)].width = w
    sm.column_dimensions[get_column_letter(LAST)].width = 15
    sm.freeze_panes = f"A{hdr_row + 1}"      # rows only - a column freeze too trips the repair prompt
    _page_setup(sm)

    tmp = out.with_name(out.stem + ".tmp.xlsx")
    wb.save(str(tmp))
    assert_clean(tmp)
    for msg in lint_layout(wb, first_col=C0):
        print(f"    ⚑ layout: {msg}")
    tmp.replace(out)
    return {"path": out, "jobs": len(jobs), "paid": round(paid, 2),
            "cut_on_jobs": round(tot["cut"], 2), "other_in_cost": round(tot["other"], 2),
            "rnet": round(tot["rnet"], 2), "rmnet": round(tot["rmnet"], 2)}


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0],
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--division", choices=sorted(cp.DIVISIONS), default="mfd")
    ap.add_argument("--out", type=Path, default=None,
                    help="the workbook to build; default "
                         "<CompanyHealth>/<DIV> PnL - Internal - Director Cut.xlsx")
    a = ap.parse_args(argv)
    div = cp.DIVISIONS[a.division]
    div_dir = pnl_paths.division_dir(div["prefix"])
    out = a.out or (paths.companyhealth_dir() / OUT_NAME.format(label=div["label"]))
    print(f"reading the {div['label']} P&L workbooks in {div_dir}")
    loaded, _skipped = cp.load_division(cp._iter_jobs(div_dir, div["prefix"]), div_dir, None)
    if not loaded:
        print("no P&L workbooks found - nothing to build")
        return 1
    jobs = []
    for job, src, _t2, src_path in loaded:
        removed = _strip_cut_from_cost(src)
        if removed:
            print(f"  {job}: {removed:,.2f} of a registered cut was inside job cost - taken out")
        src["rel"] = cp._link_target(src_path, out.parent)
        jobs.append((job, src, cp._totals(src), src_path))
    print("pulling the cut from QuickBooks")
    lines, _realm = pull_cut_lines(div["prefix"])
    res = build(jobs, lines, out, div)
    print(f"✓ {res['path']}\n  {res['jobs']} jobs · paid to the director {res['paid']:,.2f} · on "
          f"these jobs {res['cut_on_jobs']:,.2f} · other cuts charged into COST "
          f"{res['other_in_cost']:,.2f} · real net {OVERHEAD_PCT:.0%} {res['rnet']:,.0f} / "
          f"{MFD_OVERHEAD_PCT:.0%} {res['rmnet']:,.0f}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
