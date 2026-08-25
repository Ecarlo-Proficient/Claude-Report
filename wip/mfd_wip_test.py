#!/usr/bin/env python3
"""
mfd_wip_test.py - builds and refreshes the 'Test - MFD' tab of
'WIP - MASTER new.xlsx': the MFD division WIP with an entry block MFD fills in
and a QBO block the script owns.

WHY THIS TAB EXISTS
The live 'WIP - MFD' tab is hand-maintained and code-locked
(wip_excel_guard.ALLOWED_WRITE_SHEETS). 'Test - MFD' is the allow-listed
staging copy that is MEANT TO TAKE OVER from it once the owner signs off
(the user 2026-08-25). It is therefore a faithful copy of 'WIP - MFD' - same
Calibri look, same input convention - plus six new columns.

FORMATTING NOTE (read before "fixing" this file)
Repo CLAUDE.md rail 5a freezes the WIP Test tabs to the 'WIP Master' look
(Tahoma 8). That rule governs the tabs the CP/RP readers write. THIS tab is
the exception the owner asked for: it must mirror 'WIP - MFD', because it
replaces 'WIP - MFD'. Do not restyle it to Tahoma 8.

LAYOUT - columns B..M are copied verbatim and NEVER touched again.
The script only ever owns N..S:

    N  ETC                MFD types it        grey/orange input style
    O  REVISED ETC        MFD types it        seeded '=N<row>'; type over on a CO
    P  GP %               formula             (REV. CONTRACT - REVISED ETC) / REV. CONTRACT
       Q5:S5              merged banner       'QBO - LAST SYNC mm/dd/yyyy h:mm AM'
    Q  COSTS TO DATE      from QBO            green header, tinted cell, comment
    R  BILLED TO DATE     from QBO            green header, tinted cell, comment
    S  RETAINAGE (QBO)    from QBO            green header; comment carries the
                                              variance against the tab's own col M
    T  COST TO COMPLETE   formula             REVISED ETC - COSTS TO DATE

RETAINAGE - WHERE THE NUMBER COMES FROM (probed 2026-08-25, do not re-derive)
QBO tracks retainage properly: the invoice item '99 - Retainage' posts to a real
Other Current Asset account, 'Retainage Receivable'. A negative retainage line on
a draw DEBITS that account (retainage moves out of AR); billing the retainage
later CREDITS it back out. So the per-job balance of that account IS "what QBO
has", and it is pulled from the GeneralLedger report filtered to that account.

Two traps that cost an hour, both already handled here:
  * The GL report's account filter is 'account' (SINGULAR). Passing 'accounts'
    is silently IGNORED and you get the whole 66k-row general ledger back,
    truncated to the first 11 accounts - with no error.
  * Do NOT derive retainage from invoice lines the way cp_wip_reader does
    (gross P&L income minus non-retainage invoice totals). That heuristic is
    built for CP and gives the WRONG answer on MFD - on the largest job it
    missed by more than twice the retainage actually at stake - because
    retainage that has since been BILLED still sits in the
    invoice history. The GL balance nets it out; the invoice scan does not.

The account balance is expected to DISAGREE with the tab's own 'Total Retainage'
(col M), and that disagreement is the point of the column. QBO stops counting
retainage once it has been billed to the GC; the WIP tab keeps carrying it. The
cell comment spells the variance out per job.

GP % mirrors 'WIP Master'!Q ((contract - ETC) / contract) so the two sheets
agree; COST TO COMPLETE mirrors 'WIP Master'!I.

THE MFD192 PROBLEM (the reason for the 'see MFD192' markers)
'WIP - MFD' carries THREE contract rows for job 192 (Hudsonwood 009, Offsite
010, base 008), but QBO has ONE project MFD192 - all 460 cost lines sit on one
customer with no contract marker (5 lines mention OFFSITE, together well under
1% of the job). Costs cannot be split. So per the owner's 2026-08-25 ruling the QBO
figures ANCHOR on the largest-contract row of each job group; the sibling rows
get a muted 'see MFD192' marker instead of a number. SUM() ignores text, so
the totals row still adds up exactly once.

USAGE
    python3 wip/mfd_wip_test.py --seed          first build (copies WIP - MFD)
    python3 wip/mfd_wip_test.py                 refresh QBO columns only
    python3 wip/mfd_wip_test.py --dry-run       show what would change
    python3 wip/mfd_wip_test.py --no-qbo        skip the QBO pull

A re-seed over an existing 'Test - MFD' DISCARDS whatever MFD typed there, so
it is gated behind CONFIRM=Y. The default refresh never touches B..P or S.
"""
from __future__ import annotations

import argparse
import datetime as dt
import os
import re
import sys
from copy import copy
from pathlib import Path
from typing import Dict, List, Optional

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from shared import paths, xlsx_verify
from wip_excel_guard import assert_write_allowed, open_wip_workbook_for_write

WIP_EXCEL_PATH = paths.get_path(
    "WIP_EXCEL_PATH",
    paths.onedrive_base() / "Company Files - WIP Report/WIP - MASTER new.xlsx",
)

SOURCE_TAB = "WIP - MFD"      # read-only template
TARGET_TAB = "Test - MFD"     # allow-listed write target

HDR_ROW = 6                   # 'PROJECT | MOBE DATE | ...'
FIRST_DATA_ROW = 7
BANNER_ROW = 5                # the row the sync stamp merges across

COL_CONTRACT = 6              # F
COL_CO = 7                    # G
COL_ETC = 14                  # N
COL_REV_ETC = 15              # O
COL_GP_PCT = 16               # P
COL_COSTS = 17                # Q
COL_BILLED = 18               # R
COL_RETAINAGE = 19            # S
COL_CTC = 20                  # T

NEW_COLS = (COL_ETC, COL_REV_ETC, COL_GP_PCT, COL_COSTS, COL_BILLED,
            COL_RETAINAGE, COL_CTC)
QBO_COLS = (COL_COSTS, COL_BILLED, COL_RETAINAGE)   # the green block
COL_TAB_RETAINAGE = 13        # M - MFD's own 'Total Retainage', for the variance

# ── styles: lifted verbatim from 'WIP - MFD' so the tab is indistinguishable ──
FONT_NAME = "Calibri"
CURRENCY = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'
PCT = "0.00%"

_THIN = Side(style="thin", color="000000")
BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)

# MFD's existing manual-entry look: bold orange on light grey (F/G/I/J/M)
INPUT_FILL = PatternFill("solid", fgColor="FFF2F2F2")
INPUT_FONT = Font(name=FONT_NAME, size=12, bold=True, color="FFFA7D00")

# plain calculated cell, same as H/K on the live tab
CALC_FONT = Font(name=FONT_NAME, size=11)

HDR_FONT = Font(name=FONT_NAME, size=11, bold=True)
HDR_INPUT_FILL = PatternFill("solid", fgColor="FFF2F2F2")
HDR_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)

# QBO block. Green is the signal "this came from QuickBooks, hands off".
QBO_HDR_FILL = PatternFill("solid", fgColor="FF2CA01C")
QBO_HDR_FONT = Font(name=FONT_NAME, size=11, bold=True, color="FFFFFFFF")
QBO_CELL_FILL = PatternFill("solid", fgColor="FFEBF4E8")
QBO_CELL_FONT = Font(name=FONT_NAME, size=11, color="FF375623")
BANNER_FILL = PatternFill("solid", fgColor="FFD9EAD3")
BANNER_FONT = Font(name=FONT_NAME, size=10, bold=True, color="FF1F4E20")

# sibling rows of a multi-contract job: muted, obviously not a number
MUTED_FILL = PatternFill("solid", fgColor="FFF1EFE8")
MUTED_FONT = Font(name=FONT_NAME, size=10, italic=True, color="FF7F7F7F")

QBO_NOTE = ("From QuickBooks Online (project P&L, all time).\n"
            "The sync overwrites this cell - do not type here.")

RETAINAGE_NOTE = (
    "From QuickBooks Online: the balance of the 'Retainage Receivable' account "
    "for this job.\nThe sync overwrites this cell - do not type here.\n\n"
    "This is EXPECTED to differ from the Total Retainage column. QBO stops "
    "counting retainage once it has been billed to the GC; this report keeps "
    "carrying it. A gap means retainage was invoiced and has not been taken off "
    "the WIP, or the reverse.")

_JOB_RE = re.compile(r"^\s*(\d{2,4})\b")


def job_of(label) -> Optional[str]:
    """'192 - JPI MAYHILL - 008' -> 'MFD192'. None when the row has no job #."""
    m = _JOB_RE.match(str(label or ""))
    return f"MFD{m.group(1)}" if m else None


def _num(v) -> float:
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


# ─────────────────────────── sheet geometry ───────────────────────────

def data_rows(ws) -> List[int]:
    """Contract rows: from FIRST_DATA_ROW down while column B carries a job #."""
    out = []
    r = FIRST_DATA_ROW
    while r <= ws.max_row:
        label = ws.cell(row=r, column=2).value
        if not str(label or "").strip():
            break
        if job_of(label):
            out.append(r)
        r += 1
    return out


def totals_rows(ws):
    """(label_row, sum_row) of the TOTALS block, or (None, None)."""
    for r in range(FIRST_DATA_ROW, min(ws.max_row, 60) + 1):
        for c in range(2, 14):
            if str(ws.cell(row=r, column=c).value or "").strip().upper() == "TOTALS:":
                return r, r + 1
    return None, None


def group_rows(ws, rows: List[int]) -> Dict[str, List[int]]:
    groups: Dict[str, List[int]] = {}
    for r in rows:
        job = job_of(ws.cell(row=r, column=2).value)
        if job:
            groups.setdefault(job, []).append(r)
    return groups


def anchor_row(ws, rows: List[int]) -> int:
    """The row a job's QBO figures land on: the biggest revised contract.
    For MFD192 that is the base 008 row, which is what the owner asked for."""
    return max(rows, key=lambda r: _num(ws.cell(row=r, column=COL_CONTRACT).value)
               + _num(ws.cell(row=r, column=COL_CO).value))


# ─────────────────────────── seed from the live tab ───────────────────────────

def seed(wb) -> None:
    """Rebuild 'Test - MFD' as a cell-for-cell copy of 'WIP - MFD'."""
    src = wb[SOURCE_TAB]
    if TARGET_TAB in wb.sheetnames:
        del wb[TARGET_TAB]
    idx = wb.sheetnames.index(SOURCE_TAB) + 1
    ws = wb.create_sheet(TARGET_TAB, idx)
    assert_write_allowed(ws.title)

    for row in src.iter_rows():
        for cell in row:
            if cell.value is None and not cell.has_style:
                continue
            tgt = ws.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                tgt._style = copy(cell._style)
            if cell.comment is not None:
                tgt.comment = Comment(cell.comment.text, cell.comment.author or "WIP")

    for rng in src.merged_cells.ranges:
        ws.merge_cells(str(rng))
    for key, dim in src.column_dimensions.items():
        ws.column_dimensions[key].width = dim.width
        ws.column_dimensions[key].hidden = dim.hidden
    for key, dim in src.row_dimensions.items():
        ws.row_dimensions[key].height = dim.height
        ws.row_dimensions[key].hidden = dim.hidden
    ws.sheet_format.defaultRowHeight = src.sheet_format.defaultRowHeight


# ─────────────────────────── the new columns ───────────────────────────

def _style(cell, font, fill=None, fmt=None, align=None, border=True):
    cell.font = font
    if fill is not None:
        cell.fill = fill
    if fmt:
        cell.number_format = fmt
    if align is not None:
        cell.alignment = align
    if border:
        cell.border = BORDER


def build_columns(ws) -> None:
    """Write headers, banner, input styling and formulas for N..S.
    Never reads or writes columns B..M."""
    rows = data_rows(ws)
    if not rows:
        raise RuntimeError(f"No contract rows found on {ws.title!r}")
    groups = group_rows(ws, rows)
    anchors = {anchor_row(ws, rs) for rs in groups.values()}

    for col, width in ((COL_ETC, 16), (COL_REV_ETC, 16), (COL_GP_PCT, 11),
                       (COL_COSTS, 17), (COL_BILLED, 17), (COL_RETAINAGE, 17),
                       (COL_CTC, 18)):
        ws.column_dimensions[get_column_letter(col)].width = width

    # merged sync banner sitting directly on top of the QBO headers. The block
    # grew from two columns to three when retainage was added (2026-08-25), so a
    # tab built by the earlier version carries a stale Q5:R5 merge - drop any
    # merge that starts on the banner row before re-merging, or Excel reports
    # overlapping merged cells and repairs the file.
    q = get_column_letter(COL_COSTS)
    last = get_column_letter(max(QBO_COLS))
    span = f"{q}{BANNER_ROW}:{last}{BANNER_ROW}"
    for m in [str(m) for m in ws.merged_cells.ranges]:
        if m != span and m.startswith(f"{q}{BANNER_ROW}:"):
            ws.unmerge_cells(m)
    if span not in {str(m) for m in ws.merged_cells.ranges}:
        ws.merge_cells(span)
    for c in QBO_COLS:
        _style(ws.cell(row=BANNER_ROW, column=c), BANNER_FONT, BANNER_FILL,
               align=Alignment(horizontal="center", vertical="center"))

    headers = (
        (COL_ETC, "ETC", HDR_FONT, HDR_INPUT_FILL),
        (COL_REV_ETC, "REVISED ETC", HDR_FONT, HDR_INPUT_FILL),
        (COL_GP_PCT, "GP %", HDR_FONT, None),
        (COL_COSTS, "COSTS TO DATE", QBO_HDR_FONT, QBO_HDR_FILL),
        (COL_BILLED, "BILLED TO DATE", QBO_HDR_FONT, QBO_HDR_FILL),
        (COL_RETAINAGE, "RETAINAGE (QBO)", QBO_HDR_FONT, QBO_HDR_FILL),
        (COL_CTC, "COST TO COMPLETE", HDR_FONT, None),
    )
    for col, label, font, fill in headers:
        cell = ws.cell(row=HDR_ROW, column=col, value=label)
        _style(cell, font, fill, align=HDR_ALIGN)

    for row in rows:
        n = get_column_letter(COL_ETC)
        o = get_column_letter(COL_REV_ETC)

        etc = ws.cell(row=row, column=COL_ETC)
        _style(etc, INPUT_FONT, INPUT_FILL, CURRENCY)

        rev = ws.cell(row=row, column=COL_REV_ETC)
        if rev.value in (None, ""):
            rev.value = f"={n}{row}"
        _style(rev, INPUT_FONT, INPUT_FILL, CURRENCY)
        rev.comment = Comment(
            "Starts equal to ETC. Type the new budget over it when an approved "
            "change order moves the cost.", "WIP")

        gp = ws.cell(row=row, column=COL_GP_PCT,
                     value=f"=IF(H{row}=0,0,(H{row}-{o}{row})/H{row})")
        _style(gp, CALC_FONT, fmt=PCT,
               align=Alignment(horizontal="center", vertical="center"))

        is_anchor = row in anchors
        for col in QBO_COLS:
            cell = ws.cell(row=row, column=col)
            # Migration guard: before retainage was added, COST TO COMPLETE sat
            # in S, which is now a QBO column. A tab built by that version still
            # holds its formula here, and a QBO column must never carry one.
            if isinstance(cell.value, str) and cell.value.startswith("="):
                cell.value = None
            if is_anchor:
                _style(cell, QBO_CELL_FONT, QBO_CELL_FILL, CURRENCY)
                cell.comment = Comment(QBO_NOTE, "WIP")
            else:
                _style(cell, MUTED_FONT, MUTED_FILL, fmt="General",
                       align=Alignment(horizontal="center", vertical="center"))

        ctc = ws.cell(row=row, column=COL_CTC)
        qc = f"{get_column_letter(COL_COSTS)}{row}"
        ctc.value = (f'=IF(OR({o}{row}=0,NOT(ISNUMBER({qc}))),"",{o}{row}-{qc})')
        _style(ctc, CALC_FONT, fmt=CURRENCY)

    _build_totals(ws, rows)
    _build_key(ws, rows)


def _build_totals(ws, rows: List[int]) -> None:
    """Extend the existing TOTALS block across N..S. Text markers are ignored
    by SUM(), so each job's QBO figure is counted exactly once."""
    label_row, sum_row = totals_rows(ws)
    if not label_row:
        return
    first, last = rows[0], rows[-1]
    labels = {COL_ETC: "ETC", COL_REV_ETC: "REVISED ETC", COL_GP_PCT: "GP %",
              COL_COSTS: "COSTS TO DATE", COL_BILLED: "BILLED TO DATE",
              COL_RETAINAGE: "RETAINAGE (QBO)", COL_CTC: "COST TO COMPLETE"}
    for col, label in labels.items():
        _style(ws.cell(row=label_row, column=col, value=label), HDR_FONT,
               align=HDR_ALIGN)

    o = get_column_letter(COL_REV_ETC)
    for col in (COL_ETC, COL_REV_ETC, COL_COSTS, COL_BILLED, COL_RETAINAGE,
                COL_CTC):
        L = get_column_letter(col)
        cell = ws.cell(row=sum_row, column=col, value=f"=SUM({L}{first}:{L}{last})")
        _style(cell, Font(name=FONT_NAME, size=11, bold=True), fmt=CURRENCY,
               align=Alignment(horizontal="center", vertical="center"))
    gp = ws.cell(row=sum_row, column=COL_GP_PCT,
                 value=f"=IF(H{sum_row}=0,0,(H{sum_row}-{o}{sum_row})/H{sum_row})")
    _style(gp, Font(name=FONT_NAME, size=11, bold=True), fmt=PCT,
           align=Alignment(horizontal="center", vertical="center"))


def _build_key(ws, rows: List[int]) -> None:
    """Two swatch cells under the table so MFD can see which cells are theirs.
    Sits below the PRODUCTION BALANCE block that lives in column K."""
    base = max(rows) + 8
    a = Alignment(horizontal="center", vertical="center")
    _style(ws.cell(row=base, column=COL_ETC, value="MFD ENTERS"),
           Font(name=FONT_NAME, size=10, bold=True, color="FFFA7D00"),
           INPUT_FILL, align=a)
    _style(ws.cell(row=base + 1, column=COL_ETC, value="FROM QBO"),
           Font(name=FONT_NAME, size=10, bold=True, color="FFFFFFFF"),
           QBO_HDR_FILL, align=a)
    for row, text in ((base, "type your numbers in these cells"),
                      (base + 1, "written by the sync - do not type here")):
        cell = ws.cell(row=row, column=COL_REV_ETC, value=text)
        cell.font = Font(name=FONT_NAME, size=10, color="FF7F7F7F")
        cell.alignment = Alignment(horizontal="left", vertical="center")


# ─────────────────────────── QBO ───────────────────────────

def _retainage_account_id(access: str, company_id: str, qbo_api) -> Optional[str]:
    """The 'Retainage Receivable' account. Matched on AccountSubType first
    (QBO has a dedicated 'Retainage' subtype) and only then on the name, so a
    rename in the chart of accounts does not silently zero the column."""
    try:
        accounts = qbo_api.query_all(access, company_id, "Account")
    except Exception:
        return None
    for a in accounts:
        if (a.get("AccountSubType") or "").lower() == "retainage":
            return a["Id"]
    for a in accounts:
        if (a.get("Name") or "").strip().lower() == "retainage receivable":
            return a["Id"]
    return None


def fetch_retainage(access: str, company_id: str, qbo_api) -> Dict[str, float]:
    """{job -> retainage still held} = the per-project balance of the
    'Retainage Receivable' account, walked out of the GeneralLedger report.

    The filter key is 'account', SINGULAR. 'accounts' is accepted and then
    silently ignored, which hands back the entire general ledger truncated to
    its first few accounts - looking like a clean empty result. Never raises."""
    acct = _retainage_account_id(access, company_id, qbo_api)
    if not acct:
        print("  no 'Retainage Receivable' account in QBO - retainage skipped")
        return {}
    try:
        rep = qbo_api.report(access, company_id, "GeneralLedger", {
            "start_date": "2019-01-01",
            "end_date": dt.date.today().isoformat(),
            "accounting_method": "Accrual",
            "account": acct,
            "columns": "tx_date,txn_type,doc_num,name,subt_nat_amount,rbal_nat_amount",
        })
    except Exception as e:
        print(f"  retainage GL failed ({type(e).__name__}) - retainage skipped")
        return {}

    lines: List[list] = []

    def walk(node):
        if node.get("Rows"):
            for child in node["Rows"].get("Row", []):
                walk(child)
        elif node.get("ColData"):
            lines.append([c.get("value") for c in node["ColData"]])

    for node in (rep.get("Rows") or {}).get("Row", []):
        walk(node)

    out: Dict[str, float] = {}
    for cols in lines:
        cols = (cols + [""] * 6)[:6]
        job = qbo_api.extract_proj(cols[3])
        if not job:
            continue
        try:
            out[job] = out.get(job, 0.0) + float(cols[4])
        except (TypeError, ValueError):
            continue
    return out


def fetch_qbo(jobs: List[str]) -> Dict[str, dict]:
    """{job -> {'costs': float, 'billed': float, 'retainage': float}}.
    Costs = COGS + Expenses (same basis as cp_wip_reader); billed = Total
    Income, i.e. GROSS billed with retainage included; retainage = the
    Retainage Receivable balance. Never raises."""
    from shared import qbo_api

    try:
        access, company_id = qbo_api.load_credentials()
    except Exception as e:
        print(f"  QBO auth failed ({type(e).__name__}) - leaving QBO columns as they are")
        return {}
    try:
        proj_map = qbo_api.build_project_customer_map(access, company_id)
    except Exception as e:
        print(f"  QBO customer map failed ({type(e).__name__})")
        return {}

    start, end = "2019-01-01", dt.date.today().isoformat()
    retainage = fetch_retainage(access, company_id, qbo_api)
    out: Dict[str, dict] = {}
    for job in jobs:
        cust = proj_map.get(job)
        if not cust:
            print(f"  {job}: not found in QBO - skipped")
            continue
        try:
            totals = qbo_api.extract_pl_totals(
                qbo_api.fetch_project_pl(access, company_id, cust["id"], start, end))
        except Exception as e:
            print(f"  {job}: P&L failed ({type(e).__name__}) - skipped")
            continue
        out[job] = {
            "costs": _num(totals.get("cogs")) + _num(totals.get("expenses")),
            "billed": _num(totals.get("income")),
            "retainage": retainage.get(job),
        }
        ret = out[job]["retainage"]
        ret_txt = f"{ret:>13,.2f}" if ret is not None else "          n/a"
        print(f"  {job}: costs {out[job]['costs']:>14,.2f}   "
              f"billed {out[job]['billed']:>14,.2f}   retainage {ret_txt}")
    return out


def _retainage_variance(ws, job_rows: List[int], qbo_amount: float) -> str:
    """The one line of the comment that does the comparing: QBO's balance
    against the sum of this job's own 'Total Retainage' cells (col M). Summed
    across the group because a job like MFD192 spreads its retainage over three
    contract rows while QBO holds one balance."""
    theirs = sum(_num(ws.cell(row=r, column=COL_TAB_RETAINAGE).value)
                 for r in job_rows)
    gap = qbo_amount - theirs
    if abs(gap) < 0.01:
        return f"This report says {theirs:,.2f} - they agree."
    return (f"This report says {theirs:,.2f}.\n"
            f"Difference: {gap:+,.2f} (QBO minus this report).")


def write_qbo(ws, data: Dict[str, dict], stamp: str) -> int:
    """Write the QBO figures onto each job's anchor row + the sync banner.
    Touches ONLY columns Q and R and the banner cell."""
    rows = data_rows(ws)
    groups = group_rows(ws, rows)
    written = 0
    for job, job_rows in groups.items():
        anchor = anchor_row(ws, job_rows)
        for row in job_rows:
            if row == anchor:
                continue
            for col in QBO_COLS:
                cell = ws.cell(row=row, column=col)
                cell.value = f"see {job}"
                _style(cell, MUTED_FONT, MUTED_FILL, fmt="General",
                       align=Alignment(horizontal="center", vertical="center"))
        vals = data.get(job)
        if not vals:
            continue
        for col, key in ((COL_COSTS, "costs"), (COL_BILLED, "billed"),
                         (COL_RETAINAGE, "retainage")):
            cell = ws.cell(row=anchor, column=col)
            amount = vals.get(key)
            if amount is None:                  # QBO had nothing for this job
                cell.value = None
                _style(cell, QBO_CELL_FONT, QBO_CELL_FILL, CURRENCY)
                cell.comment = Comment(QBO_NOTE, "WIP")
                continue
            cell.value = round(amount, 2)
            _style(cell, QBO_CELL_FONT, QBO_CELL_FILL, CURRENCY)
            note = QBO_NOTE
            if col == COL_RETAINAGE:
                note = f"{RETAINAGE_NOTE}\n\n{_retainage_variance(ws, job_rows, amount)}"
            cell.comment = Comment(note, "WIP")
        written += 1

    banner = ws.cell(row=BANNER_ROW, column=COL_COSTS)
    banner.value = f"QBO - LAST SYNC {stamp}"
    _style(banner, BANNER_FONT, BANNER_FILL,
           align=Alignment(horizontal="center", vertical="center"))
    return written


# ─────────────────────────── main ───────────────────────────

def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--seed", action="store_true",
                    help="Rebuild 'Test - MFD' from 'WIP - MFD' (discards MFD's "
                         "typed values on the test tab; needs CONFIRM=Y if it exists).")
    ap.add_argument("--no-qbo", action="store_true", help="Skip the QBO pull.")
    ap.add_argument("--dry-run", action="store_true",
                    help="Report what would change; write nothing.")
    args = ap.parse_args()

    print(f"\n  MFD WIP -> '{TARGET_TAB}'")
    print(f"  workbook: {WIP_EXCEL_PATH}")
    if not WIP_EXCEL_PATH.exists():
        print("  workbook not found (OneDrive synced?)")
        return 1

    wb = open_wip_workbook_for_write(WIP_EXCEL_PATH)
    if SOURCE_TAB not in wb.sheetnames:
        print(f"  source tab {SOURCE_TAB!r} missing")
        return 1

    exists = TARGET_TAB in wb.sheetnames
    if args.seed:
        if exists and os.environ.get("CONFIRM") != "Y":
            print(f"  '{TARGET_TAB}' already exists. Re-seeding DISCARDS whatever "
                  f"MFD typed there.\n  Re-run with CONFIRM=Y to proceed.")
            return 2
        print(f"  seeding from '{SOURCE_TAB}'")
        if not args.dry_run:
            seed(wb)
    elif not exists:
        print(f"  '{TARGET_TAB}' does not exist yet - run with --seed first.")
        return 2

    ws = wb[TARGET_TAB] if TARGET_TAB in wb.sheetnames else wb[SOURCE_TAB]
    assert_write_allowed(ws.title if TARGET_TAB in wb.sheetnames else TARGET_TAB)

    rows = data_rows(ws)
    groups = group_rows(ws, rows)
    print(f"  {len(rows)} contract row(s), {len(groups)} QBO job(s)")
    for job, job_rows in sorted(groups.items()):
        anchor = anchor_row(ws, job_rows)
        label = str(ws.cell(row=anchor, column=2).value or "").strip()
        extra = f"  (+{len(job_rows) - 1} sibling row(s) -> 'see {job}')" if len(job_rows) > 1 else ""
        print(f"    {job}: anchor row {anchor} - {label}{extra}")

    # Runs on every pass, not just --seed: it is non-destructive (it never
    # overwrites an ETC or a REVISED ETC that already has a value), so a row
    # MFD adds later picks up the styling and formulas on the next sync.
    if not args.dry_run:
        build_columns(ws)

    stamp = dt.datetime.now().strftime("%m/%d/%Y %-I:%M %p")
    data = {} if args.no_qbo else fetch_qbo(sorted(groups))
    if not args.dry_run:
        n = write_qbo(ws, data, stamp)
        print(f"  QBO figures written to {n} anchor row(s); banner stamped {stamp}")

    if args.dry_run:
        print("  dry run - nothing written")
        return 0

    wb.save(WIP_EXCEL_PATH)
    xlsx_verify.assert_clean(WIP_EXCEL_PATH)
    print(f"  saved and verified clean: {WIP_EXCEL_PATH.name}\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
