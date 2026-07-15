#!/usr/bin/env python3
"""
master_wip_test.py — the UNIFIED WIP test sheet (the user 2026-07-15).

One run, one 'Test' tab in the WIP master (guard-allowed), all divisions
stacked in the user's order:
  main table:  MFD (from the 'WIP Master' tab: contract E / ETC F) then CP
               (live folder scan, draws model — ACTIVE projects only)
  sections:    RP SLABS — CUSTOM · RP SLABS — TRACT · FTW — ACTIVE ·
               FTW BACKLOG   (all from the General List via rp_wip_reader)
Billed + Costs always fresh from QBO per line. Read-only everywhere except
the 'Test' tab write.

Usage:  python3 master_wip_test.py [--dry-run]
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from openpyxl import load_workbook

import cp_wip_reader as CP
import rp_wip_reader as RP

MASTER_SHEET = "WIP Master"
# 'WIP Master' tab layout: header row 3, data row 4+.
MCOL_PROJ, MCOL_NAME, MCOL_CONTRACT, MCOL_ETC = 1, 2, 5, 6


def read_mfd_from_master(wip_path: Path):
    """MFD#### rows from the 'WIP Master' tab — the master sheet IS the MFD
    source for active jobs and pricing (the user 2026-07-15). Billed/Costs
    come from QBO, not from the tab."""
    wb = load_workbook(wip_path, data_only=True, read_only=True)
    ws = wb[MASTER_SHEET]
    rows = []
    for r in range(4, ws.max_row + 1):
        proj = ws.cell(r, MCOL_PROJ).value
        if not proj or not str(proj).strip().upper().startswith("MFD"):
            continue
        proj = str(proj).strip().upper()
        row = CP.CpRow(proj, str(ws.cell(r, MCOL_NAME).value or proj), False,
                       RP._money(ws.cell(r, MCOL_CONTRACT).value), None,
                       RP._money(ws.cell(r, MCOL_ETC).value), None, None)
        row.client = "Multi Family"
        row.takeoff_path = wip_path          # contract/ETC link → the master tab
        row.src_link = str(wip_path)
        row.src_fragment = CP._sheet_fragment(MASTER_SHEET, f"A{r}")
        row.notes.append(f"Contract/ETC from '{MASTER_SHEET}' row {r}")
        rows.append(row)
    wb.close()
    return rows


def master_cols():
    """CP layout minus the division-only columns (the user 2026-07-15:
    Approved COs / Retainage / NOTES live in the division sheets, and the
    master needs neither CLIENT nor the WHY column) + TYPE after the name."""
    drop = {"co_revenue", "retainage_held", "notes_text"}
    cols = []
    for label, width, field in CP.COLS:
        if field in drop:
            continue
        cols.append((label, width, field))
        if field == "project_name":
            cols.append(("TYPE", 9, "home_type"))
    return cols


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    print("\n  UNIFIED WIP — MFD → CP → RP → 'Test' tab")
    print("  " + "─" * 60)

    # ── MFD (master sheet) ──
    mfd_rows = read_mfd_from_master(CP.WIP_EXCEL_PATH)
    print(f"  MFD from '{MASTER_SHEET}': {len(mfd_rows)} job(s)")

    # ── CP (live folder scan, ACTIVE only — the unified sheet is a WIP) ──
    cp_rows = CP.scan_cp_folders(CP.CP_ACTIVE_DIR, is_completed=False)
    print(f"  CP active folders: {len(cp_rows)} project(s)")

    # ── RP (General List → lines → classify → partition) ──
    records, _missing = RP.read_general_list(RP.ALPHA_PATH)
    rp_to_folders, addr_folders = ({}, [])
    if RP.RP_ROOT.exists():
        rp_to_folders, addr_folders = RP.index_residential(RP.RP_ROOT)
    pairs = RP.build_lines(records, rp_to_folders, addr_folders)
    print(f"  RP lines from the General List: {len(pairs)}")

    # ── ONE QBO pass across everything ──
    print("  Enriching with QBO Billed/Costs (all divisions) …")
    CP.enrich_with_qbo(mfd_rows + cp_rows)
    RP.enrich_with_qbo(pairs)
    for row, comp, _rec in pairs:
        RP._classify(row, comp)
    for row in mfd_rows + cp_rows:
        row.needs_review = bool(row.status_flags)

    # ── RP partition: custom/tract slabs, FTW active, FTW backlog ──
    sched_addrs, sched_label = RP.read_schedule_flatwork(RP.SCHEDULE_DIR)
    slabs_custom, slabs_tract, ftw_active, ftw_backlog = [], [], [], []
    for row, comp, rec in pairs:
        if row.project_num.endswith("-FTW"):
            addr_n = RP._norm(f"{rec['house'] or ''} {rec['street'] or ''}")
            on_sched = bool(addr_n) and any(
                addr_n == a or a.startswith(addr_n) or addr_n.startswith(a)
                for a in sched_addrs)
            if row.billed_to_date or row.costs_to_date or on_sched:
                if on_sched:
                    row.notes.append(f"On the {sched_label} schedule — flatwork crew")
                ftw_active.append(row)
            else:
                row.needs_review = False
                ftw_backlog.append(row)
        elif row.home_type == "Tract":
            slabs_tract.append(row)
        elif row.project_num.startswith("CP"):
            cp_rows.append(row)              # CP-standalone GL jobs (e.g. CP865)
                                             # belong with CP (the user 2026-07-15)
        else:
            slabs_custom.append(row)

    # RP justification + WHY row-jump links (same workbook as the RP tab uses)
    justify_path = Path.home() / "Downloads" / "RP WIP - Justification.xlsx"
    just_rows = RP.write_justification(pairs, ftw_backlog, justify_path) or {}
    for row, _comp, _rec in pairs:
        row.why_link = str(justify_path)
        jr = just_rows.get(row.project_num)
        row.why_fragment = CP._sheet_fragment("JUSTIFICATION", f"A{jr}") if jr else None
        if row.needs_review:
            reason = "; ".join(row.status_flags) or (row.notes[-1] if row.notes else "")
            if reason and f"RED: {reason}" not in row.notes:
                row.notes.append(f"RED: {reason}")

    print(f"  Sections: MFD {len(mfd_rows)} · CP {len(cp_rows)} · "
          f"RP custom {len(slabs_custom)} · tract {len(slabs_tract)} · "
          f"FTW {len(ftw_active)} · backlog {len(ftw_backlog)}")

    try:
        wrote = CP.write_test_cp(
            mfd_rows + cp_rows, CP.WIP_EXCEL_PATH,
            dry_run=args.dry_run, tab_name="Test-Master",
            appendix=[("RP SLABS — CUSTOM", slabs_custom),
                      ("RP SLABS — TRACT", slabs_tract),
                      ("FTW — ACTIVE (won / working)", ftw_active),
                      ("FTW BACKLOG — bid with the slab, NOT poured yet "
                       "(expected wins)", ftw_backlog)],
            cols=master_cols())
    except CP.WipWriteDenied as e:
        print(f"  ✗ Guard blocked write: {e}")
        return 2
    except FileNotFoundError as e:
        print(f"  ✗ {e}")
        return 3
    total = (len(mfd_rows) + len(cp_rows) + len(slabs_custom)
             + len(slabs_tract) + len(ftw_active) + len(ftw_backlog))
    if not args.dry_run and wrote:
        print(f"  ✓ Wrote {total} line(s) to 'Test-Master'")
        _drop_stale_test_tab()
    return 0


def _drop_stale_test_tab() -> None:
    """Remove the superseded 'Test' tab (renamed to 'Test-Master' — the user
    2026-07-15). Only ever deletes 'Test'; guard-checked; live tabs untouched."""
    from openpyxl import load_workbook as _lw
    CP.assert_write_allowed("Test")
    wb = _lw(CP.WIP_EXCEL_PATH)
    if "Test" in wb.sheetnames:
        del wb["Test"]
        wb.save(CP.WIP_EXCEL_PATH)
        print("  ✓ removed superseded 'Test' tab")
    wb.close()


if __name__ == "__main__":
    sys.exit(main())
