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
    """CP layout minus the division-only columns (Retainage / NOTES live in
    the division sheets; the master needs neither CLIENT nor the WHY column).
    SECTION column first so ONE table spans every division (band rows broke
    the table's filters — the user 2026-07-15), then TYPE after the name.
    APPROVED COs kept (the user 2026-07-16): project-pnl reads Test-Master
    for Contract/ETC/COs/STATUS — original contract = total − COs.
    PROJECT FOLDER / DATA SOURCE are gone from CP.COLS itself (the user
    2026-07-29): no file links anywhere — the only links are QBO on
    Billed/Costs."""
    drop = {"retainage_held", "notes_text"}
    cols = [("SECTION", 15, "section")]
    for label, width, field in CP.COLS:
        if field in drop:
            continue
        cols.append((label, width, field))
        if field == "project_name":
            cols.append(("TYPE", 9, "home_type"))
    return cols


# The owner's fixed RP WIP ('RP WIP' sheet) — band rows partition the lines.
# Band substring → (SECTION on the master, TYPE on Test - RP). None = excluded:
# CP jobs belong to the CP folder scan, never the RP section (the user
# 2026-07-29). TYPE wording is the user's (2026-07-31): the top section is
# just GOOD; the other three bands are their own types.
_RP_FILE_BANDS = (
    ("CP JOBS",                  None),
    ("DROPPED OFF THE SCHEDULE", ("RP — DROPPED, UNBILLED", "DROPPED OFF SCHEDULE")),
    ("FTW WITH COSTS",           ("FTW — OFF-SCHEDULE (COSTS)", "FTW WITH COSTS")),
    ("FTW BACKLOG",              ("FTW BACKLOG", "FTW BACKLOG")),
)
_RP_FILE_COL = {"job": 1, "addr": 2, "builder": 3, "contract": 4, "etc": 5,
                "billed": 6, "costs": 7, "sched": 9, "action": 11, "co": 12}

# The owner's colour contract (rp_wip_update.py, the user 2026-07-30) — his
# font-colour marks are authoritative and must SURVIVE into the test tabs
# (the user 2026-07-31: "keep all the notes and colors since they mean
# something"). theme 9 = the workbook's orange → written as literal orange.
_OWNER_RGB = {"00B050": "00B050",     # green — the owner verified this number
              "FF0000": "FF0000"}     # red   — the owner changed this number
_OWNER_ORANGE = "ED7D31"              # theme 9 — ops manager must verify


def _rp_category(row) -> str:
    """The row's TYPE for a line from the owner's top section, decided AFTER
    the QBO pass (the user 2026-07-31: "rp7234-ftw is not good … there are no
    costs" — the top band alone doesn't make a job good).

    GOOD means work is actually underway: QBO shows costs or billing.
    Anything with no money moved yet is NOT STARTED when it's still on the
    schedule, and — for flatwork — plain backlog when it isn't (the locked
    FTW backlog rule: no activity AND not scheduled)."""
    if row.billed_to_date or row.costs_to_date:
        return "GOOD"
    if getattr(row, "on_schedule", False):
        return "NOT STARTED"
    return "FTW BACKLOG" if row.project_num.endswith("-FTW") else "NOT STARTED"


def _owner_mark(cell):
    """The owner's colour on this cell ('00B050' / 'FF0000' / orange), or None."""
    col = cell.font.color if cell.font else None
    if col is None:
        return None
    rgb = getattr(col, "rgb", None)
    if isinstance(rgb, str) and rgb[-6:].upper() in _OWNER_RGB:
        return _OWNER_RGB[rgb[-6:].upper()]
    th = getattr(col, "theme", None)
    return _OWNER_ORANGE if th == 9 else None


def read_rp_from_file(xlsx_path: Path):
    """RP rows from the owner's verified RP WIP workbook (the user 2026-07-29:
    'for RP, use this excel'). His contract/ETC/CO values are taken as-is —
    the file IS the RP source of record. Billed/Costs are pre-seeded from the
    file, then refreshed from QBO (QBO wins; the file value only survives a
    failed lookup). CP-numbered lines are EXCLUDED (they belong to the CP
    folder scan). Sections come from the file's band rows; lines above the
    first band split RP SLAB vs FTW — ACTIVE by the -FTW suffix."""
    import re as _re
    wb = load_workbook(xlsx_path, data_only=True, read_only=True)
    ws = wb["RP WIP"]
    C = _RP_FILE_COL
    rows, section, skipped_cp = [], "", []
    seen = {}                                          # project # → first row
    for r in range(3, ws.max_row + 1):
        raw = ws.cell(r, C["job"]).value
        if raw is None or not str(raw).strip():
            continue
        job = str(raw).strip().upper()
        if not _re.match(r"^(RP|CP)\d", job):          # band row → new section
            band = str(raw)
            for key, sect in _RP_FILE_BANDS:
                if key in band:
                    section = sect
                    break
            continue
        if job.startswith("CP"):
            skipped_cp.append(job)
            continue
        if section is None:                            # inside an excluded band
            skipped_cp.append(job)
            continue
        contract = RP._money(ws.cell(r, C["contract"]).value)
        etc = RP._money(ws.cell(r, C["etc"]).value)
        if job in seen:
            # The file lists the line twice (e.g. under two warning bands).
            # One WIP line per job — keep the FIRST (fuller) copy; if the
            # duplicate carries DIFFERENT numbers, flag it red so a human
            # settles which contract/ETC is real.
            first = seen[job]
            if (contract, etc) != (first.base_contract, first.base_etc):
                first.status_flags.append(
                    f"Duplicate line in the RP file with different numbers "
                    f"(row {r}: contract {contract}, ETC {etc}) — verify")
            print(f"    ⚠ {job}: duplicate line in the RP file (row {r}) — "
                  f"kept the first copy"
                  + ("" if (contract, etc) == (first.base_contract, first.base_etc)
                     else " (NUMBERS DIFFER — flagged)"))
            continue
        row = CP.CpRow(
            job, str(ws.cell(r, C["addr"]).value or job).strip(), False,
            contract,
            RP._money(ws.cell(r, C["co"]).value),
            etc,
            RP._money(ws.cell(r, C["billed"]).value),
            RP._money(ws.cell(r, C["costs"]).value))
        row.client = str(ws.cell(r, C["builder"]).value or "").strip() or None
        # TYPE = Tract / Custom (the user 2026-07-31 — it must not disappear
        # from the RP view). Same rule as the GL pipeline: production builders
        # (by name OR by the General-Lista code) are Tract, everyone else is
        # Custom.
        _b = RP._norm(row.client)
        row.home_type = ("Tract" if (_b in RP.TRACT_CLIENTS
                                     or _b in RP.TRACT_CODES) else "Custom")
        row.on_schedule = str(ws.cell(r, C["sched"]).value or "").strip() == "✓"
        if section:
            row.section, row.rp_type = section
        else:
            row.section = ("FTW — ACTIVE" if job.endswith("-FTW") else "RP SLAB")
            row.rp_type = None          # decided after QBO — see _rp_category()
        # The owner's notes + colour marks travel with the row (the user
        # 2026-07-31). A red/green/orange mark on Billed/Costs also means
        # HIS number stands — the QBO refresh must not overwrite it.
        row.action_note = str(ws.cell(r, C["action"]).value or "").strip() or None
        row.cell_marks, row.qbo_protect = {}, {}
        for fld, ccol in (("contract_price", "contract"), ("etc", "etc"),
                          ("billed_to_date", "billed"), ("costs_to_date", "costs"),
                          ("action_note", "action")):
            mark = _owner_mark(ws.cell(r, C[ccol]))
            if mark:
                row.cell_marks[fld] = mark
                if fld in ("billed_to_date", "costs_to_date"):
                    row.qbo_protect[fld] = getattr(row, fld)
        seen[job] = row
        rows.append(row)
    wb.close()
    if skipped_cp:
        print(f"  RP file: excluded {len(skipped_cp)} CP/band line(s): "
              f"{', '.join(skipped_cp)}")
    return rows


def existing_project_nums(wip_path: Path, tab_name: str = "Test-Master"):
    """PROJECT #s already on the Test-Master tab (read-only). Used to lock
    the RP refresh to jobs that are ALREADY in the report (the user
    2026-07-22: 'RP just the projects already in the WIP report … don't
    add')."""
    from openpyxl import load_workbook
    if not wip_path.exists():
        return set()
    wb = load_workbook(wip_path, data_only=True, read_only=True)
    if tab_name not in wb.sheetnames:
        wb.close()
        return set()
    ws = wb[tab_name]
    hdr = next((r for r in range(1, 8)
                for c in range(1, (ws.max_column or 0) + 1)
                if ws.cell(r, c).value == "PROJECT #"), None)
    nums = set()
    if hdr:
        pcol = next(c for c in range(1, ws.max_column + 1)
                    if ws.cell(hdr, c).value == "PROJECT #")
        for r in range(hdr + 1, ws.max_row + 1):
            v = ws.cell(r, pcol).value
            if v and str(v).strip().upper() != "PROJECT #":
                nums.add(str(v).strip().upper())
    wb.close()
    return nums


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--rp-existing-only", action="store_true",
                    help="RP: refresh only the project #s already in "
                         "Test-Master — never add new RP jobs (the user "
                         "2026-07-22)")
    ap.add_argument("--rp-from-file", metavar="XLSX",
                    help="RP: take the RP lines from the owner's verified "
                         "'RP WIP' workbook instead of the General List "
                         "pipeline (the user 2026-07-29); CP lines in it are "
                         "excluded, billed/costs still refresh from QBO")
    args = ap.parse_args()

    print("\n  UNIFIED WIP — MFD → CP → RP → 'Test' tab")
    print("  " + "─" * 60)

    # ── MFD (master sheet) ──
    mfd_rows = read_mfd_from_master(CP.WIP_EXCEL_PATH)
    print(f"  MFD from '{MASTER_SHEET}': {len(mfd_rows)} job(s)")

    # ── CP (live folder scan, ACTIVE only — the unified sheet is a WIP) ──
    if not CP.CP_ACTIVE_DIR.exists():
        # Tripwire (2026-07-29): an unmounted Synology once produced a WIP
        # with ZERO CP lines — a gutted report is worse than no report.
        print(f"  ✗ CP root unreachable: {CP.CP_ACTIVE_DIR}\n"
              f"    (Synology unmounted?) Refusing to write a WIP report "
              f"without its CP section — reconnect the volume and re-run.")
        return 4
    cp_rows = CP.scan_cp_folders(CP.CP_ACTIVE_DIR, is_completed=False)
    print(f"  CP active folders: {len(cp_rows)} project(s)")

    # ── RP: the owner's verified workbook (--rp-from-file) OR the
    #        General List pipeline (legacy) ──
    if args.rp_from_file:
        rp_file = Path(args.rp_from_file).expanduser()
        if not rp_file.exists():
            print(f"  ✗ RP file not found: {rp_file}")
            return 3
        rp_rows = read_rp_from_file(rp_file)
        pairs = [(row,) for row in rp_rows]     # RP.enrich_with_qbo unpacks row-first
        print(f"  RP lines from {rp_file.name!r}: {len(rp_rows)}")
    else:
        records, _missing = RP.read_general_list(RP.ALPHA_PATH)
        rp_to_folders, addr_folders = ({}, [])
        if RP.RP_ROOT.exists():
            rp_to_folders, addr_folders = RP.index_residential(RP.RP_ROOT)
        pairs = RP.build_lines(records, rp_to_folders, addr_folders)
        if args.rp_existing_only:
            keep = existing_project_nums(CP.WIP_EXCEL_PATH)
            before = len(pairs)
            pairs = [t for t in pairs if t[0].project_num in keep]
            print(f"  RP locked to existing Test-Master lines: {len(pairs)} kept "
                  f"of {before} (no new RP jobs added)")
        else:
            print(f"  RP lines from the General List: {len(pairs)}")

    # ── ONE QBO pass across everything ──
    print("  Enriching with QBO Billed/Costs (all divisions) …")
    CP.enrich_with_qbo(mfd_rows + cp_rows)
    RP.enrich_with_qbo(pairs)
    for row in mfd_rows + cp_rows:
        row.needs_review = bool(row.status_flags)
    for row in mfd_rows:
        row.section = "MFD"
    for row in cp_rows:
        row.section = "CP"

    if args.rp_from_file:
        # Sections came from the file's bands. Two post-QBO touches:
        # (1) FTW backlog rule (LOCKED, the user 2026-07-25): a backlog line
        #     with ANY costs/billing after the QBO pass is NOT backlog.
        # (2) backlog lines never render red — no QBO project is their
        #     normal state, not a must-fix.
        rp_rows = [t[0] for t in pairs]
        order = ["RP SLAB", "FTW — ACTIVE",
                 "FTW — OFF-SCHEDULE (COSTS)", "RP — DROPPED, UNBILLED",
                 "FTW BACKLOG"]
        for row in rp_rows:
            # Owner-marked Billed/Costs stand — undo the QBO overwrite (his
            # red/green/orange means he settled that number).
            for fld, val in (getattr(row, "qbo_protect", None) or {}).items():
                if getattr(row, fld) != val:
                    print(f"    ⚠ {row.project_num}: {fld} kept at the owner's "
                          f"marked value (QBO refresh discarded)")
                    setattr(row, fld, val)
            if (row.section == "FTW BACKLOG"
                    and (row.billed_to_date or row.costs_to_date)):
                row.section = "FTW — OFF-SCHEDULE (COSTS)"
                row.rp_type = "FTW WITH COSTS"
                print(f"    ⚠ {row.project_num}: QBO shows activity — "
                      f"reclassed out of FTW BACKLOG")
            if row.rp_type is None:            # owner's top section
                row.rp_type = _rp_category(row)
                if row.rp_type != "GOOD":
                    print(f"    • {row.project_num}: no costs/billing → "
                          f"{row.rp_type} (was in the top section)")
            # A line with no budget can't be measured — say so in NOTES.
            if row.base_etc is None:
                row.status_flags.append("No budget (ETC) on the RP file")
            row.needs_review = (bool(row.status_flags)
                                and row.section != "FTW BACKLOG")
            if row.project_num == "RP6901":     # the user 2026-07-31: invoiced?
                print(f"    ► RP6901 check: QBO billed = "
                      f"{row.billed_to_date or 0:,.2f} · costs = "
                      f"{row.costs_to_date or 0:,.2f}")
        rp_sorted = [r for s in order for r in rp_rows if r.section == s]
        all_rows = mfd_rows + cp_rows + rp_sorted
        counts = " · ".join(
            f"{s} {sum(1 for r in rp_sorted if r.section == s)}"
            for s in order)
        print(f"  Sections: MFD {len(mfd_rows)} · CP {len(cp_rows)} · {counts}")
    else:
        for row, comp, _rec in pairs:
            RP._classify(row, comp)

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
                cp_rows.append(row)          # CP-standalone GL jobs (e.g. CP865)
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

        for rows_, sect in ((cp_rows, "CP"),      # re-stamp: partition may have
                                                  # appended CP-standalone GL rows
                            (slabs_custom, "RP SLAB — CUSTOM"),
                            (slabs_tract, "RP SLAB — TRACT"),
                            (ftw_active, "FTW — ACTIVE"),
                            (ftw_backlog, "FTW BACKLOG")):
            for row in rows_:
                row.section = sect
        all_rows = (mfd_rows + cp_rows + slabs_custom + slabs_tract
                    + ftw_active + ftw_backlog)
    import datetime as dt
    try:
        wrote = CP.write_test_cp(
            all_rows, CP.WIP_EXCEL_PATH,
            dry_run=args.dry_run, tab_name="Test-Master",
            cols=master_cols(), default_filter_active=True,
            # Banner + logo space + TOTALS/cash-flow (the user 2026-07-16).
            # The logo image floats over the banner rows and survives syncs.
            title="WIP REPORT",
            summary=True,
            # The user 2026-07-29: the ONLY links on the report are the QBO
            # deep links on Billed/Costs — no Synology/file links.
            qbo_links_only=True,
            # The change audit runs HERE — Test-Master carries all three
            # divisions, so one report covers MFD + CP + RP (the user
            # 2026-07-31: "I need to always audit these things").
            audit=True, audit_xlsx=CP.AUDIT_XLSX)
    except CP.WipWriteDenied as e:
        print(f"  ✗ Guard blocked write: {e}")
        return 2
    except FileNotFoundError as e:
        print(f"  ✗ {e}")
        return 3
    total = len(all_rows)
    if not args.dry_run and wrote:
        print(f"  ✓ Wrote {total} line(s) to 'Test-Master'")
        _drop_stale_test_tab()

    # With --rp-from-file, 'Test - RP' becomes the SAME rows in the master
    # formatting, PLUS the user's 2026-07-31 asks: a TYPE column that
    # describes every row (GOOD / FTW WITH COSTS / DROPPED OFF SCHEDULE /
    # FTW BACKLOG), a legend on top explaining the types and the owner's
    # colours, the BUILDER, and the owner's notes carried in NOTES.
    if args.rp_from_file:
        type_order = ["GOOD", "NOT STARTED", "FTW WITH COSTS",
                      "DROPPED OFF SCHEDULE", "FTW BACKLOG"]
        rp_view = [r for t in type_order for r in all_rows
                   if r.section not in ("MFD", "CP")
                   and getattr(r, "rp_type", None) == t]
        # CATEGORY describes the row; TYPE stays Tract/Custom exactly as on
        # the master layout (the user 2026-07-31).
        rp_cols = [("CATEGORY", 20, "rp_type")]
        for label, width, field in master_cols():
            if field == "section":
                continue
            rp_cols.append((label, width, field))
            if field == "home_type":
                rp_cols.append(("BUILDER", 24, "client"))
        legend = [
            ("LEGEND — CATEGORY:", None, True),
            ("GOOD — work is underway: QBO shows costs and/or billing", None, False),
            ("NOT STARTED — on the schedule / priced, but no costs and no billing yet", None, False),
            ("FTW WITH COSTS — flatwork started off-schedule (has costs) — belongs on the schedule", None, False),
            ("DROPPED OFF SCHEDULE — left the schedule with money still on the table", None, False),
            ("FTW BACKLOG — flatwork priced with the slab, no activity and not scheduled (expected wins)", None, False),
            ("COLORS:  GREEN = the owner verified this number", "00B050", True),
            ("RED = the owner changed this number", "FF0000", True),
            ("ORANGE = ops manager must verify", "ED7D31", True),
        ]
        try:
            wrote_rp = CP.write_test_cp(
                rp_view, CP.WIP_EXCEL_PATH,
                dry_run=args.dry_run, tab_name="Test - RP",
                cols=rp_cols, default_filter_active=True,
                title="RP WIP REPORT",
                summary=True, qbo_links_only=True, legend=legend)
            if not args.dry_run and wrote_rp:
                print(f"  ✓ Wrote {len(rp_view)} RP line(s) to 'Test - RP'")
        except CP.WipWriteDenied as e:
            print(f"  ✗ Guard blocked the 'Test - RP' write: {e}")
            return 2
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
