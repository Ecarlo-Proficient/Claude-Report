"""Historical WIP reconstruction — a point-in-time WIP as of a past date, built
SEPARATELY from the live WIP master (the user 2026-08-08). Two dates asked for:
12-31-2025 and 3-31-2026.

SOURCES (all confirmed present):
  · MFD + CP — the monthly .xlsb snapshot in 'Company Files - WIP Report/WIP
    History' already froze each division's WIP at month-end. Read the
    'WIP - CP' / 'WIP - MFD' tabs (pyxlsb) — that IS the truth as of the date,
    no QBO needed. Snapshot carries contract · COs · revised · billed
    (COMPLETED TO DATE) · % complete · retainage — NO cost column (it's a
    billing-based WIP).
  · RP — NOT in the old snapshots. Rebuilt from the schedule of that exact day
    → the RP jobs working then → each job's bid proposal vs. its invoice
    (match 100% ⇒ contract acquired, else FLAG) → QBO billed/costs dated
    on/before the report date. [stage B/C — added next]

Output: '~/Downloads/WIP as of <date>.xlsx', one tab per division. Touches
nothing else. Stage A here = MFD + CP from the snapshots.
"""
import argparse
import os
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
from pyxlsb import open_workbook as open_xlsb

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from shared.xlsx_verify import assert_clean, safe_table_name  # noqa: E402

WIP_HISTORY = Path(os.getenv(
    "WIP_HISTORY_DIR",
    str(Path.home() / "Library/CloudStorage/OneDrive-ProficientConcrete,LLC"
        / "Company Files - WIP Report" / "WIP History")))

# date key → (snapshot .xlsb, human label)
DATES = {
    "12-31-2025": ("WIP_12-31.25.xlsb", "December 31, 2025"),
    "3-31-2026":  ("WIP - 03-31-26.xlsb", "March 31, 2026"),
}

# The snapshot columns we lift, by division tab. (header label → out label)
CP_COLS = [
    ("PROJECT", "PROJECT"), ("CUSTOMER", "CUSTOMER"), ("CONTRACT", "CONTRACT"),
    ("CHANGE ORDERS", "CHANGE ORDERS"), ("REV. CONTRACT", "REVISED CONTRACT"),
    ("COMPLETED TO DATE", "BILLED TO DATE"), ("% COMPLETE", "% COMPLETE"),
    ("BALANCE TO FINISH INCL'N RET", "BALANCE TO FINISH"),
    ("Total Retainage", "RETAINAGE"),
]
MFD_COLS = CP_COLS  # same header names on the MFD tab

HDR_FILL = PatternFill("solid", fgColor="D9D9D9")
_thin = Side(style="thin", color="000000")
BORDER = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
MONEY = '"$"#,##0_);[Red]("$"#,##0)'
PCT = "0%"
FONT = "Tahoma"


def _num(v):
    return float(v) if isinstance(v, (int, float)) else None


def read_division(xlsb_path: Path, tab: str, colspec):
    """Rows from a snapshot division tab. Excludes rows explicitly marked
    COMPLETED in col A; keeps every other job (its status travels along).
    Returns (rows, header_out_labels)."""
    with open_xlsb(str(xlsb_path)) as wb:
        with wb.get_sheet(tab) as sh:
            grid = [[c.v for c in row] for row in sh.rows()]
    hdr_i = next((i for i, r in enumerate(grid)
                  if any(str(v).strip() == "PROJECT" for v in r if v)), None)
    if hdr_i is None:
        return [], [o for _, o in colspec]
    hdr = grid[hdr_i]
    idx = {str(v).strip(): j for j, v in enumerate(hdr) if v not in (None, "")}
    pcol = idx.get("% COMPLETE")
    bcol = idx.get("BALANCE TO FINISH INCL'N RET")
    out = []
    for r in grid[hdr_i + 1:]:
        if not r:
            continue
        status = str(r[0]).strip() if r[0] not in (None, "") else ""
        proj = r[idx["PROJECT"]] if idx.get("PROJECT") is not None and len(r) > idx["PROJECT"] else None
        if not proj or not str(proj).strip():
            continue
        # ACTIVE = still billing as of the snapshot. The tab is a cumulative
        # ledger (mostly completed jobs kept for reference), and the col-A
        # COMPLETED tag is only on some — so key off % complete: < 100% billed
        # is active; fully-billed/retainage-only jobs live in the RETAINAGE tab.
        pct = _num(r[pcol]) if (pcol is not None and len(r) > pcol) else None
        bal = _num(r[bcol]) if (bcol is not None and len(r) > bcol) else None
        if "COMPLETED" in status.upper():
            continue
        active = (pct is not None and pct < 0.999) or (pct is None and bool(bal and bal > 1))
        if not active:
            continue
        rec = {"STATUS": status or "active"}
        for src, dst in colspec:
            j = idx.get(src)
            v = r[j] if (j is not None and len(r) > j) else None
            rec[dst] = _num(v) if dst not in ("PROJECT", "CUSTOMER") else (
                str(v).strip() if v not in (None, "") else "")
        out.append(rec)
    return out, [o for _, o in colspec]


def _write_tab(ws, title, subtitle, rows, out_labels, table_name):
    ws.cell(1, 1, title).font = Font(name=FONT, size=11, bold=True)
    ws.cell(2, 1, subtitle).font = Font(name=FONT, size=8)
    headers = ["STATUS"] + out_labels
    hdr = 4
    for c, label in enumerate(headers, 1):
        cell = ws.cell(hdr, c, label)
        cell.fill = HDR_FILL
        cell.font = Font(name=FONT, size=8, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    widths = {"PROJECT": 26, "CUSTOMER": 22, "STATUS": 11}
    for c, label in enumerate(headers, 1):
        ws.column_dimensions[ws.cell(hdr, c).column_letter].width = widths.get(label, 15)
    for i, rec in enumerate(rows, hdr + 1):
        for c, label in enumerate(headers, 1):
            cell = ws.cell(i, c, rec.get(label))
            cell.font = Font(name=FONT, size=8)
            cell.border = BORDER
            if label in ("CONTRACT", "CHANGE ORDERS", "REVISED CONTRACT",
                         "BILLED TO DATE", "RETAINAGE"):
                cell.number_format = MONEY
            elif label == "% COMPLETE":
                cell.number_format = PCT
    last = hdr + len(rows)
    if rows:
        from openpyxl.utils import get_column_letter
        ref = f"A{hdr}:{get_column_letter(len(headers))}{last}"
        t = Table(displayName=table_name, ref=ref)   # already valid + unique
        t.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=False)
        ws.tables.add(t)
    ws.freeze_panes = f"A{hdr+1}"


def build(date_key: str, out_dir: Path) -> Path:
    snap, label = DATES[date_key]
    xlsb = WIP_HISTORY / snap
    if not xlsb.exists():
        raise FileNotFoundError(f"snapshot not found: {xlsb}")
    cp, cp_h = read_division(xlsb, "WIP - CP", CP_COLS)
    mfd, mfd_h = read_division(xlsb, "WIP - MFD", MFD_COLS)
    seen = set()
    wb = Workbook()
    _write_tab(wb.active, f"WIP as of {label} — COMMERCIAL",
               f"from snapshot '{snap}' · WIP - CP tab · billing-based (no cost column in source)",
               cp, cp_h, safe_table_name("histCP", seen))
    wb.active.title = "CP"
    _write_tab(wb.create_sheet("MFD"), f"WIP as of {label} — MULTI-FAMILY",
               f"from snapshot '{snap}' · WIP - MFD tab", mfd, mfd_h,
               safe_table_name("histMFD", seen))
    rp = wb.create_sheet("RP")
    rp.cell(1, 1, f"WIP as of {label} — RESIDENTIAL").font = Font(name=FONT, size=11, bold=True)
    rp.cell(3, 1, "[stage B/C] schedule-of-day + proposal↔invoice + QBO as-of-date — not built yet").font = Font(name=FONT, size=9, italic=True)
    out = out_dir / f"WIP as of {date_key}.xlsx"
    wb.save(out)
    assert_clean(out)          # NEVER hand over a file that would trip Excel repair
    print(f"  ✓ {date_key}: CP {len(cp)} active · MFD {len(mfd)} active · "
          f"xlsx verified clean → {out}")
    return out


def main() -> int:
    ap = argparse.ArgumentParser(description="Point-in-time WIP from the history snapshots (MFD/CP).")
    ap.add_argument("--date", choices=list(DATES) + ["all"], default="all")
    ap.add_argument("--out", default=str(Path.home() / "Downloads"))
    args = ap.parse_args()
    out_dir = Path(args.out).expanduser()
    keys = list(DATES) if args.date == "all" else [args.date]
    print("\n  HISTORICAL WIP — MFD/CP from snapshots (stage A)")
    for k in keys:
        build(k, out_dir)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
