#!/usr/bin/env python3
"""
ftw_recode_list.py — the per-bill worklist for moving flatwork off the slabs.

`job_reality_audit.py` found FW (flatwork) cost sitting on jobs that are not
`-FTW` projects. This turns that into something a clerk can work: every
offending LINE, grouped by vendor, biggest vendor first, so one vendor gets
knocked out at a time instead of hopping between them.

THREE SHEETS, because they are three different actions:
  1. Recode by vendor  - the slab has a `-FTW` twin already, so the line just
                         moves job. Unambiguous; this is the work.
  2. Create -FTW first  - FW on an RP slab with NO twin. Somebody has to make
                         the job before anything can move.
  3. Wrong code         - FW on a CP or MFD job. Those divisions have no -FTW
                         twin at all, so this is a cost-code correction, not a
                         job move (the owner's rule, 2026-08-06).

Plain formatting on purpose (repo rule 5): this is a worklist, not a report.
Rows are grouped but NEVER hidden - the clerk can collapse a vendor with the
outline +/- if they want to.

Reads the cache `job_reality_audit.py` writes. No QBO pull of its own.

USAGE
  python3 one-offs/ftw_recode_list.py --out ~/Downloads/"FTW Recodes.xlsx"
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import re
import sys
from collections import defaultdict
from pathlib import Path
from urllib.parse import quote

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from shared.xlsx_verify import assert_clean                    # noqa: E402

CACHE = Path(os.environ.get(
    "ACB_AUDIT_CACHE",
    Path(os.environ.get("TMPDIR", "/tmp")) / "job_reality_audit.json"))
FW_RE = re.compile(r"^FW\d+$", re.IGNORECASE)
MONEY = '#,##0.00;[Red](#,##0.00)'
THIN = Side(style="thin", color="000000")


def qbo_url(tx_type: str, txn_id: str, realm: str) -> str:
    """Deep link straight to the bill, so the clerk clicks instead of searching."""
    if not txn_id or not realm:
        return ""
    page = "bill" if (tx_type or "").lower() == "bill" else "expense"
    return (f"https://qbo.intuit.com/app/login?pagereq="
            f"{quote(f'{page}?txnId={txn_id}')}&deeplinkcompanyid={realm}")


def classify(costs: list) -> tuple:
    """Split the FW lines into the three actions."""
    jobs_with_cost = {c["proj"].upper() for c in costs}
    move, need_job, wrong_code = [], [], []
    for c in costs:
        leaf = (c.get("code") or "").split(":")[-1].strip()
        if not FW_RE.match(leaf):
            continue
        p = c["proj"].upper()
        if p.endswith("-FTW"):
            continue
        if p.startswith("RP"):
            twin = f"{p}-FTW"
            (move if twin in jobs_with_cost else need_job).append(dict(c, target=twin))
        else:
            wrong_code.append(dict(c, target=""))
    return move, need_job, wrong_code


def _t(ws, r, c, v, *, bold=False, fmt=None, align=None, wrap=False, size=11):
    cell = ws.cell(row=r, column=c, value=v)
    cell.font = Font(size=size, bold=bold)
    if fmt:
        cell.number_format = fmt
    if align or wrap:
        cell.alignment = Alignment(horizontal=align, wrap_text=wrap, vertical="center")
    return cell


def sheet_by_vendor(wb, title: str, rows: list, realm: str, blurb: str,
                    show_target: bool) -> None:
    ws = wb.create_sheet(title)
    ws.sheet_view.showGridLines = False
    by_vendor = defaultdict(list)
    for r in rows:
        by_vendor[r.get("vendor") or "(no vendor)"].append(r)
    order = sorted(by_vendor, key=lambda v: -sum(x["amt"] for x in by_vendor[v]))
    total = sum(r["amt"] for r in rows)
    bills = len({(r.get("txn"), r.get("doc")) for r in rows})

    _t(ws, 1, 1, title.upper(), bold=True, size=14)
    _t(ws, 2, 1, blurb, size=10)
    _t(ws, 3, 1, f"{len(order)} vendors · {bills} bills · {len(rows)} lines · "
                 f"{total:,.2f} total · biggest vendor first", size=10)
    heads = ["Done", "Date", "Bill #", "Move FROM"] + \
            (["Move TO"] if show_target else []) + \
            ["Code", "Amount", "Description", "Open in QBO"]
    r = 5
    for i, h in enumerate(heads):
        c = _t(ws, r, 2 + i, h, bold=True,
               align="right" if h == "Amount" else "left")
        c.border = Border(bottom=THIN)
    r += 1
    for vendor in order:
        lines = sorted(by_vendor[vendor], key=lambda x: (x.get("date") or ""))
        sub = sum(x["amt"] for x in lines)
        _t(ws, r, 2, f"{vendor}   ({len(lines)} line(s))", bold=True)
        _t(ws, r, 2 + heads.index("Amount"), sub, bold=True, fmt=MONEY, align="right")
        for i in range(len(heads)):
            ws.cell(row=r, column=2 + i).border = Border(top=THIN)
        r += 1
        for x in lines:
            vals = [None, x.get("date") or "", str(x.get("doc") or ""),
                    x["proj"]] + \
                   ([x.get("target") or ""] if show_target else []) + \
                   [(x.get("code") or "").split(":")[-1], x["amt"],
                    (x.get("desc") or x.get("memo") or "")[:70], None]
            for i, v in enumerate(vals):
                cell = _t(ws, r, 2 + i, v,
                          fmt=MONEY if heads[i] == "Amount" else None,
                          align="right" if heads[i] == "Amount" else None)
                if heads[i] == "Done":
                    cell.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
            link = qbo_url(x.get("type", ""), x.get("txn", ""), realm)
            if link:
                lc = ws.cell(row=r, column=2 + len(heads) - 1, value="open bill")
                lc.hyperlink = link
                lc.font = Font(size=11, color="0563C1", underline="single")
            # grouped, NOT hidden (repo rule 5: no hidden rows) - the clerk can
            # collapse a vendor with the outline +/- if they want to
            ws.row_dimensions[r].outline_level = 1
            r += 1
    _t(ws, r, 2, "TOTAL", bold=True)
    _t(ws, r, 2 + heads.index("Amount"), total, bold=True, fmt=MONEY, align="right")
    for i in range(len(heads)):
        ws.cell(row=r, column=2 + i).border = Border(top=THIN)

    ws.column_dimensions["A"].width = 3
    widths = [7, 11, 14, 13] + ([13] if show_target else []) + [8, 14, 52, 13]
    for i, w in enumerate(widths):
        ws.column_dimensions[get_column_letter(2 + i)].width = w
    ws.sheet_properties.outlinePr.summaryBelow = False
    ws.freeze_panes = f"A{6}"


def main() -> int:
    ap = argparse.ArgumentParser(description="Per-bill FTW recode worklist")
    ap.add_argument("--out", default=str(Path.home() / "Downloads" / "FTW Recodes.xlsx"))
    ap.add_argument("--cache", default=str(CACHE))
    a = ap.parse_args()
    cache = Path(a.cache).expanduser()
    if not cache.exists():
        print(f"✗  no audit cache at {cache} — run one-offs/job_reality_audit.py first")
        return 1
    data = json.loads(cache.read_text())
    realm = data.get("realm", "")
    move, need_job, wrong_code = classify(data["costs"])

    wb = Workbook()
    wb.remove(wb.active)
    sheet_by_vendor(
        wb, "Recode by vendor", move, realm, show_target=True,
        blurb="Move each line from the slab to its -FTW twin. The twin already "
              "exists, so nothing needs creating first.")
    if need_job:
        sheet_by_vendor(
            wb, "Create -FTW first", need_job, realm, show_target=False,
            blurb="FW cost on an RP slab with NO -FTW twin. The flatwork job has "
                  "to be created before these can move.")
    if wrong_code:
        sheet_by_vendor(
            wb, "Wrong code", wrong_code, realm, show_target=False,
            blurb="FW cost on a CP or MFD job. Those divisions have no -FTW "
                  "twin, so this is a cost-code correction, not a job move.")
    out = Path(a.out).expanduser()
    out.parent.mkdir(parents=True, exist_ok=True)
    tmp = out.with_suffix(".tmp.xlsx")
    wb.save(str(tmp))
    assert_clean(tmp)
    tmp.replace(out)
    for label, rows in (("recode", move), ("create -FTW first", need_job),
                        ("wrong code", wrong_code)):
        print(f"  {label:20} {len(rows):5} lines  {sum(r['amt'] for r in rows):12,.2f}")
    print(f"\n  → {out}   (generated {dt.datetime.now():%m/%d/%Y %I:%M %p})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
