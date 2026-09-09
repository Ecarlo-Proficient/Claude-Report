#!/usr/bin/env python3
"""wip_month_end.py - the month-end WIP report as of a cutoff date, in the
bank layout the 12-31-25 and 3-31-26 reports use, plus the RP jobs that
closed during that month (the user 2026-09-09: "make an end of month August
WIP report, i also want to see all the closed RP jobs of that month").

HOW THE REPORT IS BUILT (the standard split - see the wip-standard-schedule
reference: estimator data from the master, accounting data from QBO):
  * POPULATION + CONTRACT + ETC  = the live master's Test-Master tab. That tab
    is the bank-facing active WIP for all three divisions (MFD from WIP Master,
    CP from the G702 draws, RP from the owner's RP WIP file). Type / name /
    bonded come from it too. Nothing is invented here.
  * COSTS TO DATE + BILLED TO DATE = QBO project P&L dated 2019-01-01 .. the
    cutoff - the SAME pull the live readers do (fetch_project_pl +
    extract_pl_totals), only date-bounded. Income = billed (gross, retainage
    included); COGS + expenses = costs.
  * The eleven derived columns are Excel formulas, exactly as wip_qc expects.

CLOSED RP JOBS (second sheet). The RP tab carries no Closed status (its source
is the owner's RP WIP file), and no RP project was made inactive in QBO during
the month, so "closed" is decided from evidence, and every row shows it:
  * BILLED OUT       - billed to the cutoff reaches the contract, and either the
                       final invoice is dated in the month or the job left the
                       crew schedule during the month.
  * LEFT THE SCHEDULE, BILLING OPEN - on the last schedule of the prior month,
                       gone from the cutoff's schedule, and NOT fully billed.
  * INVOICED, NOT ON THE WIP - an RP project invoiced in the month that has no
                       contract on the WIP (usually one-invoice flatwork).
Schedules read: the last 'Schedule m-d-yy.xlsx' on/before the prior month-end
and on/before the cutoff, 'Main Schedule' tab (PROJECT column; FLATWORK band
rows are the -FTW line).

OUTPUT: '<WIP History>/WIP <m-d-yy>.xlsx' (sheet 1 = the report, sheet 2 =
Closed RP). The file is scrubbed of the openpyxl fingerprint and verified
with shared.xlsx_verify as the LAST step. It never touches the master.
Operator detail (diffs against the live tab, per-division totals) goes to
stdout only - nothing internal is printed on the report.

Usage:
  python3 one-offs/wip_month_end.py                       (prior month-end)
  python3 one-offs/wip_month_end.py --cutoff 2026-08-31
  python3 one-offs/wip_month_end.py --cutoff 2026-08-31 --out-dir ~/Downloads
"""
from __future__ import annotations

import argparse
import calendar
import datetime as dt
import re
import shutil
import sys
import tempfile
import zipfile
from collections import defaultdict
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.table import Table, TableStyleInfo

_REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_REPO))
sys.path.insert(0, str(_REPO / "one-offs"))
from shared import paths, qbo_api                              # noqa: E402
from shared.xlsx_verify import assert_clean                    # noqa: E402
from rp_schedule_wip_preview import read_main_schedule         # noqa: E402

MASTER = paths.get_path(
    "WIP_EXCEL_PATH",
    paths.onedrive_base() / "Company Files - WIP Report" / "WIP - MASTER new.xlsx")
HISTORY = paths.get_path(
    "WIP_HISTORY_DIR",
    paths.onedrive_base() / "Company Files - WIP Report" / "WIP History")
RP_FILE = paths.get_path("RP_WIP_FILE", paths.onedrive_base() / "RP WIP TO FIX_Final.xlsx")
SCHED_ROOT = paths.get_path("RP_SCHEDULE_ROOT", "/Volumes/Common/OPERATIONS/SCHEDULE")
_SCHED_RE = re.compile(r"Schedule (\d{1,2})-(\d{1,2})-(\d{2})\.xlsx$", re.IGNORECASE)
QBO_START = "2019-01-01"
FULL_TOL = 0.01          # billed within 1% of the contract = fully billed

FONT = "Tahoma"
MONEY = '"$"#,##0_);[Red]("$"#,##0)'
PCT = "0.00%"
DATE = "mm/dd/yyyy"
HDR_FILL = PatternFill("solid", fgColor="D9D9D9")
_thin = Side(style="thin", color="000000")
BORDER = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)

HEADERS = [
    "TYPE", "PROJECT #", "PROJECT NAME", "BONDED", "TOTAL CONTRACT PRICE",
    "ESTIMATED TOTAL COSTS", "ORIGINAL PROFIT", "GROSS PROFIT %",
    "COSTS TO DATE", "COST TO COMPLETE", "PERCENT COMPLETE",
    "REVENUES EARNED TO DATE", "PROFIT EARNED TO DATE", "BILLED TO DATE",
    "OVERBILLINGS", "UNDERBILLINGS", "LEFT TO BILL", "FUTURE PROFIT TO EARN",
    "PURE JOB BORROW",
]
FORMULAS = {
    7:  '=IF(OR(E{r}="",F{r}=""),"",E{r}-F{r})',
    8:  '=IF(OR(G{r}="",E{r}="",E{r}=0),"",G{r}/E{r})',
    10: '=IF(OR(F{r}="",I{r}=""),"",F{r}-I{r})',
    11: '=IF(OR(I{r}="",F{r}="",F{r}=0),"",I{r}/F{r})',
    12: '=IF(OR(E{r}="",I{r}="",F{r}="",F{r}=0),"",E{r}*I{r}/F{r})',
    13: '=IF(OR(L{r}="",I{r}=""),"",L{r}-I{r})',
    15: '=IF(OR(N{r}="",L{r}=""),"",MAX(N{r}-L{r},0))',
    16: '=IF(OR(L{r}="",N{r}=""),"",MAX(L{r}-N{r},0))',
    17: '=IF(OR(E{r}="",N{r}=""),"",E{r}-N{r})',
    18: '=IF(OR(G{r}="",M{r}=""),"",G{r}-M{r})',
    19: '=IF(OR(J{r}="",Q{r}=""),"",MAX(J{r}-Q{r},0))',
}
MONEY_COLS = {5, 6, 7, 9, 10, 12, 13, 14, 15, 16, 17, 18, 19}
PCT_COLS = {8, 11}
CENTER_COLS = {1, 2, 4}
WIDTHS = {"A": 27.66, "B": 11.5, "C": 42.83, "D": 7.66, "E": 21.0, "F": 22.0,
          "G": 16.16, "H": 15.33, "I": 14.33, "J": 17.16, "L": 23.83, "M": 22.0,
          "N": 15.33, "O": 13.33, "P": 14.33, "Q": 13.33, "R": 22.0, "S": 16.16}


# ────────────────────────── helpers ──────────────────────────
def _num(v):
    return float(v) if isinstance(v, (int, float)) else None


def _label(s):
    """The TYPE label with the master's dash separators dropped, ASCII only
    ('Residential - Custom - Slab' -> 'Residential Custom Slab', the way the
    3-31-26 report writes it)."""
    s = str(s or "")
    s = re.sub(r"\s+[-–—]\s+", " ", s)
    return s.encode("ascii", "ignore").decode()


def _fmt_date(d: dt.date) -> str:
    return f"{d.month}-{d.day}-{d.year % 100:02d}"


def _prior_month_end(cutoff: dt.date) -> dt.date:
    return cutoff.replace(day=1) - dt.timedelta(days=1)


def _schedule_on_or_before(day: dt.date):
    """Newest 'Schedule m-d-yy.xlsx' under SCHED_ROOT dated on/before `day`."""
    best = None
    for f in SCHED_ROOT.glob("*/*/Schedule *.xlsx"):
        m = _SCHED_RE.search(f.name)
        if not m:
            continue
        mo, dy, yy = (int(g) for g in m.groups())
        try:
            d = dt.date(2000 + yy, mo, dy)
        except ValueError:
            continue
        if d <= day and (best is None or d > best[0]):
            best = (d, f)
    return best


def _schedule_jobs(day: dt.date):
    """{project# -> {address, builder, section}} on the newest schedule on/before
    `day`; FLATWORK-band rows are the -FTW line. (None, {}) if none is mounted."""
    found = _schedule_on_or_before(day)
    if not found:
        return None, {}
    sdate, path = found
    out = {}
    for rec in read_main_schedule(path):
        job = rec["job"].upper()
        if rec["scope"] == "ftw" and not job.endswith("-FTW"):
            job += "-FTW"
        out.setdefault(job, {"address": rec["address"], "builder": rec["builder"],
                             "section": rec["section"]})
    return sdate, out


# ────────────────────────── sources ──────────────────────────
def read_master():
    """Test-Master rows (values) + the company prefix from WIP Master!B1."""
    wb = load_workbook(MASTER, data_only=True)
    title = str(wb["WIP Master"]["B1"].value or "").strip()
    company = title.split(" - ")[0].strip() if " - " in title else title
    ws = wb["Test-Master"]
    hdr = next((r for r in range(1, 16) if ws.cell(r, 2).value == "PROJECT #"), 3)
    rows = []
    for r in range(hdr + 1, ws.max_row + 1):
        proj = str(ws.cell(r, 2).value or "").strip().upper()
        if not proj or proj.startswith("TOTAL") or not qbo_api.PROJ_RE.search(proj):
            continue
        rows.append({
            "type": _label(ws.cell(r, 1).value), "proj": proj,
            "name": str(ws.cell(r, 3).value or "").strip(),
            "bonded": "Yes" if str(ws.cell(r, 4).value or "").strip().upper() in ("Y", "YES") else "No",
            "contract": _num(ws.cell(r, 5).value), "etc": _num(ws.cell(r, 6).value),
            "tab_costs": _num(ws.cell(r, 9).value), "tab_billed": _num(ws.cell(r, 14).value),
        })
    wb.close()
    return company, rows


def read_rp_file():
    """Every RP line in the owner's RP WIP file (all bands): contract / ETC /
    address / builder by project #. First copy wins on duplicates."""
    out = {}
    if not RP_FILE.exists():
        return out
    wb = load_workbook(RP_FILE, data_only=True)
    ws = wb["RP WIP"] if "RP WIP" in wb.sheetnames else wb.worksheets[0]
    for r in range(1, ws.max_row + 1):
        job = str(ws.cell(r, 1).value or "").strip().upper()
        if not re.match(r"^RP\d{4}(-FTW)?$", job) or job in out:
            continue
        out[job] = {"address": str(ws.cell(r, 2).value or "").strip(),
                    "builder": str(ws.cell(r, 3).value or "").strip(),
                    "contract": _num(ws.cell(r, 4).value), "etc": _num(ws.cell(r, 5).value)}
    wb.close()
    return out


def _qbo_name(proj_map, proj):
    """The address part of the QBO customer name ('RP####-FTW - <address>' ->
    '<address>'); '' when there is no project."""
    name = (proj_map.get(proj) or {}).get("name") or ""
    return re.sub(r"^\s*" + re.escape(proj) + r"\s*[-:]?\s*", "", name, flags=re.IGNORECASE).strip()


def qbo_totals_at(access, cid, cust_id, cutoff_iso):
    t = qbo_api.extract_pl_totals(qbo_api.fetch_project_pl(access, cid, cust_id, QBO_START, cutoff_iso))
    billed = round(float(t.get("income", 0) or 0), 2)
    costs = round(float(t.get("cogs", 0) or 0) + float(t.get("expenses", 0) or 0), 2)
    return billed, costs


def invoices_by_project(access, cid, cutoff_iso):
    """{project# -> [(date, doc#, total)]} for every invoice dated on/before the
    cutoff since 2025-01-01, attributed with the ONE invoice->job rule."""
    inv = qbo_api.query_all(access, cid, "Invoice",
                            f"TxnDate >= '2025-01-01' AND TxnDate <= '{cutoff_iso}'")
    by = defaultdict(list)
    for i in inv:
        p = qbo_api.project_of_invoice(i)
        if p:
            by[p.upper()].append((dt.date.fromisoformat(i["TxnDate"]), i.get("DocNumber"),
                                  float(i.get("TotalAmt") or 0)))
    for v in by.values():
        v.sort()
    return by


# ────────────────────────── the report sheet ──────────────────────────
def _style(c, bold=False, fill=None, center=False, fmt=None, border=True):
    c.font = Font(name=FONT, size=8, bold=bold)
    if fill:
        c.fill = fill
    if border:
        c.border = BORDER
    if center:
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    if fmt:
        c.number_format = fmt


def write_report(ws, company, rows, cutoff: dt.date):
    ws["A1"] = f"{company} - WIP REPORT"
    ws["A2"] = f"REPORT DATE: {cutoff.strftime('%b %d, %Y').upper()}"
    for a in ("A1", "A2"):
        ws[a].font = Font(name=FONT, size=8, bold=True)
    for c, h in enumerate(HEADERS, 1):
        _style(ws.cell(3, c, h), bold=True, fill=HDR_FILL, center=True)
    ws.row_dimensions[3].height = 28
    r = 4
    for row in rows:
        vals = {1: row["type"], 2: row["proj"], 3: row["name"], 4: row["bonded"],
                5: row["contract"], 6: row["etc"], 9: row["costs"], 14: row["billed"]}
        for c in range(1, 20):
            v = vals.get(c) if c not in FORMULAS else FORMULAS[c].format(r=r)
            cell = ws.cell(r, c, v)
            _style(cell, center=(c in CENTER_COLS),
                   fmt=MONEY if c in MONEY_COLS else PCT if c in PCT_COLS else None)
        ws.row_dimensions[r].height = 14
        r += 1
    last = r - 1
    tot = r
    ws.cell(tot, 1, "TOTALS")
    for c in range(1, 20):
        cell = ws.cell(tot, c)
        col = cell.column_letter
        if c in MONEY_COLS:
            cell.value = f"=SUBTOTAL(109,{col}4:{col}{last})"
        elif c == 8:
            cell.value = f'=IF(E{tot}=0,"",G{tot}/E{tot})'
        elif c == 11:
            cell.value = f'=IF(F{tot}=0,"",I{tot}/F{tot})'
        _style(cell, bold=True, fill=HDR_FILL,
               fmt=MONEY if c in MONEY_COLS else PCT if c in PCT_COLS else None)
    tab = Table(displayName="Wip" + cutoff.strftime("%b"), ref=f"A3:S{last}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=False)
    ws.add_table(tab)
    for col, w in WIDTHS.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A4"
    ws.sheet_view.selection[0].activeCell = "A4"
    ws.sheet_view.selection[0].sqref = "A4"
    ws.print_area = f"A1:S{tot}"
    ws.print_title_rows = "1:3"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    return last, tot


# ────────────────────────── the closed-RP sheet ──────────────────────────
CLOSED_HDR = ["PROJECT #", "PROJECT NAME", "BUILDER", "TOTAL CONTRACT PRICE",
              "ESTIMATED TOTAL COSTS", "COSTS TO DATE", "PERCENT COMPLETE", "BILLED TO DATE",
              "% BILLED", "LAST INVOICE", "BILLED THIS MONTH", "ON PRIOR SCHEDULE",
              "ON CUTOFF SCHEDULE", "EVIDENCE"]
CLOSED_WIDTHS = {"A": 12, "B": 42, "C": 30, "D": 18, "E": 18, "F": 14, "G": 10, "H": 14,
                 "I": 9, "J": 12, "K": 16, "L": 12, "M": 12, "N": 70}
CLOSED_MONEY = {4, 5, 6, 8, 11}
CLOSED_PCT = {7, 9}
CLOSED_CENTER = {1, 7, 9, 10, 12, 13}


def classify_closed(cands, month_start: dt.date, cutoff: dt.date, prior_sd, cutoff_sd):
    """cands: {proj -> dict(contract, etc, costs, billed, last_inv, month_billed,
    on_prior, on_cutoff, name, builder)} -> three blocks."""
    p_lbl = _fmt_date(prior_sd) if prior_sd else "n/a"
    c_lbl = _fmt_date(cutoff_sd) if cutoff_sd else "n/a"
    out = {"BILLED OUT": [], "LEFT THE SCHEDULE, BILLING OPEN": [], "INVOICED, NOT ON THE WIP": []}
    for proj, c in sorted(cands.items()):
        K, B = c.get("contract"), c.get("billed") or 0.0
        li = c.get("last_inv")
        inv_in_month = bool(li and month_start <= li <= cutoff)
        left = bool(c["on_prior"] and not c["on_cutoff"])
        fully = bool(K and K > 0 and B >= K * (1 - FULL_TOL))
        why = []
        if K and K > 0:
            if fully and (inv_in_month or left):
                if li:
                    why.append(f"billed out - last invoice {li.strftime('%m/%d/%Y')}")
                if left:
                    why.append(f"left the crew schedule (on {p_lbl}, gone {c_lbl})")
                if c["on_cutoff"]:
                    why.append(f"still on the {c_lbl} schedule ({c['on_cutoff']})")
                out["BILLED OUT"].append((proj, c, "; ".join(why)))
            elif left and not fully:
                why.append(f"on the {p_lbl} schedule, gone {c_lbl}; billed {B:,.0f} of {K:,.0f}")
                if li:
                    why.append(f"last invoice {li.strftime('%m/%d/%Y')}")
                out["LEFT THE SCHEDULE, BILLING OPEN"].append((proj, c, "; ".join(why)))
        elif inv_in_month:
            why.append(f"invoiced {li.strftime('%m/%d/%Y')} - no contract on the WIP or the RP file")
            if c["on_cutoff"]:
                why.append(f"on the {c_lbl} schedule ({c['on_cutoff']})")
            out["INVOICED, NOT ON THE WIP"].append((proj, c, "; ".join(why)))
    return out


def write_closed(ws, company, blocks, cutoff: dt.date, month_name: str):
    ws["A1"] = f"{company} - RP JOBS CLOSED IN {month_name.upper()}"
    ws["A2"] = f"REPORT DATE: {cutoff.strftime('%b %d, %Y').upper()}"
    for a in ("A1", "A2"):
        ws[a].font = Font(name=FONT, size=8, bold=True)
    for c, h in enumerate(CLOSED_HDR, 1):
        _style(ws.cell(3, c, h), bold=True, fill=HDR_FILL, center=True)
    ws.row_dimensions[3].height = 28
    r = 4
    n = 0
    for title, items in blocks.items():
        cell = ws.cell(r, 1, f"{title} ({len(items)})")
        cell.font = Font(name=FONT, size=8, bold=True)
        r += 1
        for proj, c, why in items:
            K, B, E, C = c.get("contract"), c.get("billed"), c.get("etc"), c.get("costs")
            pct_billed = (B / K) if (K and B is not None) else None
            pct_done = (C / E) if (E and C is not None) else None
            vals = [proj, c.get("name") or "", c.get("builder") or "", K, E, C, pct_done, B,
                    pct_billed, c.get("last_inv"), c.get("month_billed") or 0.0,
                    "Yes" if c["on_prior"] else "No", "Yes" if c["on_cutoff"] else "No", why]
            for col, v in enumerate(vals, 1):
                cell = ws.cell(r, col, v)
                fmt = (MONEY if col in CLOSED_MONEY else PCT if col in CLOSED_PCT
                       else DATE if col == 10 else None)
                _style(cell, center=(col in CLOSED_CENTER), fmt=fmt)
            ws.row_dimensions[r].height = 14
            r += 1
            n += 1
        r += 1
    last = r - 2
    for col, w in CLOSED_WIDTHS.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A4"
    ws.sheet_view.selection[0].activeCell = "A4"
    ws.sheet_view.selection[0].sqref = "A4"
    ws.print_area = f"A1:N{max(last, 4)}"
    ws.print_title_rows = "1:3"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    return n


# ────────────────────────── fingerprint scrub ──────────────────────────
def scrub(path: Path, company: str) -> None:
    """Rewrite the zip: Excel's own app properties, the company as creator, no
    [trash] parts. The file is otherwise byte-identical."""
    tmp = path.with_suffix(".tmp.xlsx")
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename.startswith("[trash]/"):
                continue
            data = zin.read(item.filename)
            if item.filename == "docProps/app.xml":
                data = re.sub(rb"<Application>.*?</Application>",
                              b"<Application>Microsoft Excel</Application>", data)
                if b"<AppVersion>" in data:
                    data = re.sub(rb"<AppVersion>.*?</AppVersion>", b"<AppVersion>16.0300</AppVersion>", data)
                else:
                    data = data.replace(b"</Properties>", b"<AppVersion>16.0300</AppVersion></Properties>")
            elif item.filename == "docProps/core.xml":
                who = company.encode()
                data = re.sub(rb"(<dc:creator[^>]*>).*?(</dc:creator>)", lambda m: m.group(1) + who + m.group(2), data)
                data = re.sub(rb"(<cp:lastModifiedBy[^>]*>).*?(</cp:lastModifiedBy>)",
                              lambda m: m.group(1) + who + m.group(2), data)
            zout.writestr(item, data)
    shutil.move(str(tmp), str(path))


# ────────────────────────── main ──────────────────────────
def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--cutoff", help="report date YYYY-MM-DD (default: the prior month-end)")
    ap.add_argument("--out-dir", help="where to write (default: the WIP History folder)")
    args = ap.parse_args()

    if args.cutoff:
        cutoff = dt.date.fromisoformat(args.cutoff)
    else:
        cutoff = _prior_month_end(dt.date.today())
    month_start = cutoff.replace(day=1)
    month_name = calendar.month_name[cutoff.month] + f" {cutoff.year}"
    cutoff_iso = cutoff.isoformat()
    out_dir = Path(args.out_dir).expanduser() if args.out_dir else HISTORY
    out = out_dir / f"WIP {_fmt_date(cutoff)}.xlsx"
    if not MASTER.exists():
        print(f"master not mounted: {MASTER}", file=sys.stderr)
        return 2

    company, rows = read_master()
    rp_file = read_rp_file()
    print(f"Test-Master: {len(rows)} jobs · RP file: {len(rp_file)} lines · cutoff {cutoff_iso}")

    access, cid = qbo_api.load_credentials()
    proj_map = qbo_api.build_project_customer_map(access, cid)
    cache = {}

    def totals(proj):
        if proj not in cache:
            cust = proj_map.get(proj)
            cache[proj] = qbo_totals_at(access, cid, cust["id"], cutoff_iso) if cust else None
        return cache[proj]

    # 1. the report population, costs + billed at the cutoff
    missing, ahead = [], []
    for i, row in enumerate(rows, 1):
        t = totals(row["proj"])
        if t is None:
            row["billed"], row["costs"] = 0.0, 0.0
            missing.append(row["proj"])
        else:
            row["billed"], row["costs"] = t
            for fld in ("billed", "costs"):
                tab = row.get(f"tab_{fld}")
                if tab is not None and row[fld] > tab + 1:
                    ahead.append((row["proj"], fld, row[fld], tab))
        if i % 20 == 0:
            print(f"  ...{i}/{len(rows)} jobs pulled")
    if missing:
        print(f"  no QBO project for {len(missing)}: {', '.join(missing)} (costs/billed = 0)")
    if ahead:
        print(f"  {len(ahead)} figure(s) HIGHER at the cutoff than on the live tab - read these:")
        for p, f, v, tab in ahead:
            print(f"    {p} {f}: cutoff {v:,.2f} vs tab {tab:,.2f}")

    # 2. closed RP candidates
    prior_sd, prior = _schedule_jobs(_prior_month_end(cutoff))
    cutoff_sd, cur = _schedule_jobs(cutoff)
    print(f"schedules: prior {(_fmt_date(prior_sd) if prior_sd else 'none')} ({len(prior)} jobs) · "
          f"cutoff {(_fmt_date(cutoff_sd) if cutoff_sd else 'none')} ({len(cur)} jobs)")
    inv_by = invoices_by_project(access, cid, cutoff_iso)
    master_rp = {r["proj"]: r for r in rows if r["proj"].startswith("RP")}
    month_inv = {p for p, v in inv_by.items() if p.startswith("RP") and any(month_start <= d <= cutoff for d, _, _ in v)}
    universe = (set(master_rp) | {p for p in rp_file} | {p for p in prior if p.startswith("RP")}
                | {p for p in cur if p.startswith("RP")} | month_inv)
    cands = {}
    for proj in sorted(universe):
        m, f = master_rp.get(proj), rp_file.get(proj)
        on_prior = prior.get(proj, {}).get("section") if proj in prior else None
        on_cutoff = cur.get(proj, {}).get("section") if proj in cur else None
        invs = inv_by.get(proj, [])
        last_inv = invs[-1][0] if invs else None
        in_month = bool(last_inv and month_start <= last_inv <= cutoff)
        left = bool(on_prior and not on_cutoff)
        if not (in_month or left):
            continue                     # nothing happened to it this month
        t = totals(proj)
        billed, costs = t if t else (None, None)
        if t is None and invs:
            billed = sum(a for _, _, a in invs)
        contract = (m or {}).get("contract") or (f or {}).get("contract")
        cands[proj] = {
            "contract": contract, "etc": (m or {}).get("etc") or (f or {}).get("etc"),
            "costs": costs, "billed": billed, "last_inv": last_inv,
            "month_billed": sum(a for d, _, a in invs if month_start <= d <= cutoff),
            "on_prior": on_prior, "on_cutoff": on_cutoff,
            "name": (m or {}).get("name") or (f or {}).get("address")
                    or prior.get(proj, cur.get(proj, {})).get("address", "")
                    or _qbo_name(proj_map, proj),
            "builder": (f or {}).get("builder") or prior.get(proj, cur.get(proj, {})).get("builder", ""),
        }
    blocks = classify_closed(cands, month_start, cutoff, prior_sd, cutoff_sd)

    # 3. write, scrub, verify - in that order
    out_dir.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = f"WIP {_fmt_date(cutoff)}"
    last, tot = write_report(ws, company, rows, cutoff)
    ws2 = wb.create_sheet(f"Closed RP {cutoff.strftime('%b %Y')}")
    n_closed = write_closed(ws2, company, blocks, cutoff, month_name)
    wb.active = 0
    ws.sheet_view.tabSelected = True
    ws2.sheet_view.tabSelected = False
    tmp = Path(tempfile.mkdtemp()) / out.name
    wb.save(tmp)
    scrub(tmp, company)
    assert_clean(tmp)
    shutil.move(str(tmp), str(out))
    assert_clean(out)

    # 4. operator summary (stdout only)
    print(f"\nWIP {_fmt_date(cutoff)}: {len(rows)} jobs, rows 4-{last}, TOTALS row {tot}")
    div = defaultdict(lambda: [0, 0.0, 0.0, 0.0, 0.0])
    for r in rows:
        d = "MFD" if r["proj"].startswith("MFD") else "CP" if r["proj"].startswith("CP") else "RP"
        for key in (d, "ALL"):
            a = div[key]
            a[0] += 1
            a[1] += r["contract"] or 0
            a[2] += r["etc"] or 0
            a[3] += r["costs"] or 0
            a[4] += r["billed"] or 0
    print(f"  {'DIV':4} {'jobs':>4} {'contract':>14} {'ETC':>14} {'costs':>14} {'billed':>14}  GP%   %compl")
    for key in ("MFD", "CP", "RP", "ALL"):
        n, k, e, c, b = div[key]
        gp = (k - e) / k if k else 0
        pc = c / e if e else 0
        print(f"  {key:4} {n:4} {k:14,.0f} {e:14,.0f} {c:14,.0f} {b:14,.0f} {gp:6.2%} {pc:6.2%}")
    print(f"\nClosed RP ({month_name}): {n_closed} rows")
    for title, items in blocks.items():
        print(f"  {title}: {len(items)}")
        for proj, c, why in items:
            print(f"    {proj:12} {str(c.get('name') or '')[:34]:34} contract {(c.get('contract') or 0):>10,.0f} "
                  f"billed {(c.get('billed') or 0):>10,.0f}  {why}")
    print(f"\n  verified clean -> {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
