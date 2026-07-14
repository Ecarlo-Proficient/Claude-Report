#!/usr/bin/env python3
"""
rp_wip_reader.py — RP (Residential) WIP reader — v2 (the user 2026-07-13).

MODEL (locked 2026-07-10, built 2026-07-13):
  • SOURCE = the General List workbook (READ-ONLY, never written): sheets
    'General list - Alpha order' + 'Small Jobs' (identical layout — Small Jobs
    holds the jobs kept out of Alpha to avoid clutter). Prices are entered
    there by hand: AI=SLAB BID $, AJ=SLAB COST $, AK=FLATWORK BID $,
    AL=FLATWORK COST $. Completion % in Z, job # in C.
  • The WIP AUTO-SPLITS each RP job into TWO lines: RP#### (slab — contract=AI,
    ETC=AJ) and RP####-FTW (flatwork — contract=AK, ETC=AL). Flatwork is
    pre-bid and assumed to follow the slab; RP####-FTW is its own standalone
    QBO project. CP#### jobs in the list are STANDALONE (never -FTW): one
    line, contract=AI+AK, ETC=AJ+AL, billed under the plain CP#.
  • QBO enrich per line (read-only GET): Billed = P&L income of the line's
    own customer; Costs = COGS + expenses. A job is DONE when 100% complete
    AND billed = contract (the RP manager's rule — billing is the truth).
  • needs_review → RED numbers in Excel (the user 2026-07-13): billed over
    contract · 100% but not fully billed (punch work) · billed with no
    contract in the sheet · fully billed but <100% on the list · contract
    with no QBO project.
  • Writes ONLY the 'Test - RP' tab of the WIP master (wip_excel_guard).
    Links: project name → Residential folder; Contract/ETC → the General
    List; Billed → QBO customer page; Costs → QBO project P&L report.

Usage:
  python3 rp_wip_reader.py --dry-run               # preview, no write
  python3 rp_wip_reader.py --project RP7538        # one job (both lines)
  python3 rp_wip_reader.py --no-qbo --dry-run      # fast local test
  python3 rp_wip_reader.py                         # live run → Test - RP
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import re
import sys
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

from openpyxl import load_workbook

# Reuse the CP writer so RP gets the exact same WIP structure/formatting/links,
# just written to the 'Test - RP' tab. Intra-folder import (allowed).
import cp_wip_reader as CP
from shared import qbo_api

ALPHA_PATH = Path(os.getenv(
    "RP_ALPHA_PATH",
    "/Volumes/Common/OPERATIONS/GENERAL LIST/LISTA GENERAL AÑO 2026.xlsx",
))
ALPHA_SHEET = "General list - Alpha order"
SMALL_SHEET = "Small Jobs"
RP_ROOT = Path(os.getenv(
    "RP_ROOT",
    "/Volumes/Common/CURRENT PROJECTS/Residential",
))

# General List column map (1-based), header row 4, data from row 6.
COL_JOB, COL_HOUSE, COL_STREET, COL_CITY, COL_COMPLETION = 3, 4, 5, 6, 26
COL_SLAB_BID, COL_SLAB_COST, COL_FLAT_BID, COL_FLAT_COST = 35, 36, 37, 38

# RP#### (with optional suffix like -FTW already in the list) or CP#### —
# CP jobs live here because the RP team runs them (standalone, never split).
_JOB_RE = re.compile(r"^(RP\d{4}|CP\d{3,4})\b", re.IGNORECASE)
_RP_RE = re.compile(r"RP\d{4}(?:-[A-Za-z]{2,6})?(?!\d)", re.IGNORECASE)

# Fully billed tolerance — within half a percent of contract counts as billed
# out (retainage/rounding noise).
_FULL_TOL = 0.005
_VARIANCE_CLOSE = 1000.0   # 100% + billed within $1K of contract → Closed w/ variance
                           # (the user 2026-07-14 materiality rule; covers WRH's 1.5% fee)


def _norm(s) -> str:
    return re.sub(r"\s+", " ", str(s or "").upper().strip())


def _money(v):
    try:
        f = float(v)
        return f if f else None
    except (TypeError, ValueError):
        return None


# ─────────────────── General List: jobs + prices ───────────────────
def read_general_list(path: Path):
    """Read BOTH source sheets (Alpha + Small Jobs) READ-ONLY.
    Returns [{job, source, completion, house, street, city, slab_bid,
    slab_cost, flat_bid, flat_cost}] for every RP/CP job with any price
    entered OR completion < 100%. Alpha wins on duplicates."""
    wb = load_workbook(path, data_only=True)
    out, seen = [], set()
    missing_sheets = []
    for sheet in (ALPHA_SHEET, SMALL_SHEET):
        if sheet not in wb.sheetnames:
            missing_sheets.append(sheet)
            continue
        ws = wb[sheet]
        hdr = _norm(ws.cell(4, COL_SLAB_BID).value)
        if "SLAB BID" not in hdr:
            print(f"  ⚠ {sheet!r}: no SLAB BID header at col AI — "
                  f"price columns missing, sheet skipped")
            continue
        for r in range(6, ws.max_row + 1):
            job = ws.cell(r, COL_JOB).value
            if not job:
                continue
            m = _JOB_RE.match(str(job).strip())
            if not m:
                continue
            job = m.group(1).upper()
            if job in seen:
                continue
            comp = ws.cell(r, COL_COMPLETION).value
            rec = {
                "job": job, "source": sheet, "gl_row": r,
                "completion": comp if isinstance(comp, (int, float)) else None,
                "house": ws.cell(r, COL_HOUSE).value,
                "street": ws.cell(r, COL_STREET).value,
                "city": ws.cell(r, COL_CITY).value,
                "slab_bid": _money(ws.cell(r, COL_SLAB_BID).value),
                "slab_cost": _money(ws.cell(r, COL_SLAB_COST).value),
                "flat_bid": _money(ws.cell(r, COL_FLAT_BID).value),
                "flat_cost": _money(ws.cell(r, COL_FLAT_COST).value),
            }
            has_price = any(rec[k] for k in
                            ("slab_bid", "slab_cost", "flat_bid", "flat_cost"))
            # PRICED ONLY (the user 2026-07-14): a job with no price for any
            # scope was SKIPPED on purpose ("unless we put the contract price
            # there just to be done") — it does not enter the WIP. Kills the
            # bloat of empty slab lines from Small Jobs / legacy Alpha rows.
            if has_price:
                seen.add(job)
                out.append(rec)
    wb.close()
    return out, missing_sheets


# ─────────────────── Today's schedule: flatwork crews ──────────────
SCHEDULE_DIR = Path(os.getenv(
    "RP_SCHEDULE_DIR", "/Volumes/Common/OPERATIONS/SCHEDULE"))
_SCHED_FILE_RE = re.compile(r"Schedule (\d{1,2})-(\d{1,2})-(\d{2})\.xlsx$",
                            re.IGNORECASE)


def read_schedule_flatwork(sched_dir: Path):
    """Latest daily schedule → the set of normalized ADDRESSES whose crew
    description mentions flatwork (the user 2026-07-14: anything on the
    schedule under Flatwork is won and being worked — never backlog).
    Returns (set_of_norm_addresses, schedule_label) — empty set if the
    schedule can't be found/read (backlog rule then falls back to $-only)."""
    best = None
    try:
        for year_dir in sched_dir.iterdir():
            if not (year_dir.is_dir() and year_dir.name.isdigit()):
                continue
            for month_dir in year_dir.iterdir():
                if not month_dir.is_dir():
                    continue
                for f in month_dir.iterdir():
                    m = _SCHED_FILE_RE.search(f.name)
                    if m:
                        mo, dy, yy = (int(g) for g in m.groups())
                        key = (2000 + yy, mo, dy)
                        if best is None or key > best[0]:
                            best = (key, f)
    except OSError:
        return set(), None
    if best is None:
        return set(), None
    _key, path = best
    label = f"{_key[1]}-{_key[2]}-{_key[0] % 100:02d}"
    addrs = set()
    try:
        wb = load_workbook(path, data_only=True, read_only=True)
        sheet = next((s for s in wb.sheetnames
                      if s.strip().lower() == "daily schedule"), wb.sheetnames[0])
        ws = wb[sheet]
        for row in ws.iter_rows(min_row=1, max_row=200):
            vals = [str(getattr(c, "value", "") or "") for c in row[:8]]
            desc = " ".join(vals).upper()
            if "FLATWORK" in desc or re.search(r"\bFTW\b", desc):
                # address is the first street-looking cell in the row
                for v in vals:
                    vn = _norm(v)
                    if re.match(r"^\d+[A-Z0-9 ,.'-]+$", vn) and len(vn) > 8:
                        addrs.add(vn)
                        break
        wb.close()
    except Exception:
        return set(), None
    return addrs, label


# ─────────────────── Residential: folder lookup ────────────────────
def _list_dir(path: Path):
    """One directory listing via scandir → (subdirs, filenames). Errors → empty."""
    subdirs, files = [], []
    try:
        with os.scandir(path) as it:
            for e in it:
                try:
                    if e.is_dir():
                        subdirs.append(Path(e.path))
                    elif e.is_file():
                        files.append(e.name)
                except OSError:
                    pass
    except OSError:
        pass
    return subdirs, files


def index_residential(root: Path, workers: int = 24):
    """RP#→folder index from filenames at client (depth 1) + address (depth 2)
    levels, listed in parallel (Synology round-trips overlap)."""
    root = Path(root)
    rp_to_folders = defaultdict(set)
    addr_folders = []

    clients, _root_files = _list_dir(root)
    with ThreadPoolExecutor(max_workers=workers) as ex:
        client_listings = list(ex.map(_list_dir, clients))

    addr_dirs = []
    for client, (subs, files) in zip(clients, client_listings):
        for fn in files:
            for m in _RP_RE.findall(fn):
                rp_to_folders[m.upper()].add(client)
        addr_dirs.extend(subs)

    with ThreadPoolExecutor(max_workers=workers) as ex:
        addr_listings = list(ex.map(_list_dir, addr_dirs))

    for addr, (_subs, files) in zip(addr_dirs, addr_listings):
        addr_folders.append((_norm(addr.name), addr))
        for fn in files:
            for m in _RP_RE.findall(fn):
                rp_to_folders[m.upper()].add(addr)
    return rp_to_folders, addr_folders


def match_by_address(rec, addr_folders):
    """Fallback: match house+street to an address folder name."""
    house = _norm(rec["house"])
    street_words = [w for w in _norm(rec["street"]).split()
                    if w not in ("ROAD", "STREET", "ST", "DRIVE", "DR", "LANE",
                                 "LN", "AVENUE", "AVE", "COURT", "CT", "TRAIL",
                                 "CIRCLE", "WAY", "BLVD", "N", "S", "E", "W")]
    for norm_name, folder in addr_folders:
        if house and house in norm_name and any(w in norm_name for w in street_words):
            return folder
    return None


# ─────────────────── line building + classification ────────────────
def _classify(row: CP.CpRow, completion) -> None:
    """The RP done-rule (billing is the truth): appends notes/flags and sets
    needs_review (→ RED numbers in Excel) in-place."""
    K = row.base_contract
    B = row.billed_to_date
    c100 = isinstance(completion, (int, float)) and completion >= 0.999
    if completion is not None:
        row.notes.append(f"{completion * 100:.0f}% complete (list)")

    if K and B is None and row.qbo_customer_id is None and row.status_flags:
        # QBO enrich already flagged (no project / fetch failed) — red it.
        row.needs_review = True
        return
    if not K:
        # No contract on this line — ANY QBO activity is a red flag
        # (the user 2026-07-14, RP5542: slab line had no bid yet QBO carried
        # $9.6K of costs — either the bid was never entered or the costs are
        # coded to the wrong project/scope).
        activity = []
        if B:
            activity.append(f"billed ${B:,.0f}")
        if row.costs_to_date:
            activity.append(f"costs ${row.costs_to_date:,.0f}")
        if activity:
            row.status_flags.append(
                "No contract in the list but QBO has "
                + " · ".join(activity) + " — enter the bid or fix the coding")
            row.needs_review = True
        return
    if B is None:
        return                      # --no-qbo run: contract-only view
    if B > K * (1 + _FULL_TOL):
        row.notes.append(f"Billed OVER contract by ${B - K:,.0f}")
        row.needs_review = True
    elif B >= K * (1 - _FULL_TOL):
        if c100:
            # The done-rule (RP manager): 100% complete AND fully billed →
            # STATUS reads 'Closed' (the user 2026-07-14: STATUS was hardcoded
            # Active for every line; it must derive from this rule).
            row.notes.append("CLOSED — fully billed + 100%")
            row.is_completed = True
        else:
            row.notes.append("Fully billed but list <100% — update completion")
            row.needs_review = True
    elif B == 0:
        if c100:
            row.notes.append("100% but $0 billed — punch/flatwork backlog")
            row.needs_review = True
        else:
            row.notes.append("Not billed yet")
    else:
        if c100:
            gap = K - B
            if gap <= _VARIANCE_CLOSE:
                # Materiality rule (the user 2026-07-14): 100% + billed within
                # $1K of contract = close with a documented small variance
                # (builder fee like WRH's 1.5%, or an approved write-down) —
                # amber note, NOT red. Over the threshold: chase the billing.
                row.notes.append(
                    f"CLOSED — small variance ${gap:,.0f} (fee/write-down)")
                row.is_completed = True
            else:
                row.notes.append(f"100% but only ${B:,.0f} billed — bill the rest")
                row.needs_review = True
        else:
            row.notes.append("Partially billed — treat as draw")


def build_lines(records, rp_to_folders, addr_folders):
    """General List records → CpRow lines. RP splits slab/-FTW; CP stays one
    standalone line (never -FTW — the user 2026-07-13)."""
    rows = []
    for rec in sorted(records, key=lambda x: x["job"]):
        name = _norm(f"{rec['house'] or ''} {rec['street'] or ''}") or rec["job"]
        folders = sorted(rp_to_folders.get(rec["job"], ()),
                         key=lambda f: (f.parent.name, f.name))
        folder = folders[0] if folders else match_by_address(rec, addr_folders)

        def _mk(line_num, contract, etc):
            row = CP.CpRow(line_num, name, False, contract, None, etc, None, None)
            row.folder_path = folder
            row.takeoff_path = ALPHA_PATH      # Contract/ETC cells link → the List
            if rec["source"] == SMALL_SHEET:
                row.notes.append("Small Jobs list")
            return row

        if rec["job"].startswith("CP"):
            # CP standalone: whole contract on one line, bills under CP#.
            contract = ((rec["slab_bid"] or 0) + (rec["flat_bid"] or 0)) or None
            etc = ((rec["slab_cost"] or 0) + (rec["flat_cost"] or 0)) or None
            row = _mk(rec["job"], contract, etc)
            row.notes.append("CP standalone (never -FTW)")
            rows.append((row, rec["completion"], rec))
            continue

        # A line exists ONLY for a scope that has pricing (the user
        # 2026-07-14): Small Jobs entries are usually flatwork-only — the
        # clerk confirmed what is real by pricing it. No slab price → no
        # slab line (e.g. RP5542: take only the -FTW line).
        if rec["slab_bid"] or rec["slab_cost"]:
            rows.append((_mk(rec["job"], rec["slab_bid"], rec["slab_cost"]),
                         rec["completion"], rec))
        if rec["flat_bid"] or rec["flat_cost"]:
            rows.append((_mk(rec["job"] + "-FTW", rec["flat_bid"],
                             rec["flat_cost"]), rec["completion"], rec))
    return rows


def enrich_with_qbo(pairs) -> None:
    """Billed (P&L income) + Costs (COGS+expenses) per LINE customer —
    read-only GETs. Sets CP.QBO_REALM so the writer builds QBO deep links."""
    try:
        access, company_id = qbo_api.load_credentials()
    except Exception as e:
        for row, *_ in pairs:
            row.status_flags.append(f"QBO Auth Failed: {type(e).__name__}")
        return
    try:
        proj_map = qbo_api.build_project_customer_map(access, company_id)
    except Exception as e:
        for row, *_ in pairs:
            row.status_flags.append(f"QBO Customer Map Failed: {type(e).__name__}")
        return
    CP.QBO_REALM = company_id
    start, end = "2019-01-01", dt.date.today().isoformat()
    for n, (row, *_rest) in enumerate(pairs, 1):
        cust = proj_map.get(row.project_num)
        if not cust:
            if row.base_contract:
                row.status_flags.append("No QBO project")
            continue
        row.qbo_customer_id = cust["id"]
        try:
            totals = qbo_api.extract_pl_totals(
                qbo_api.fetch_project_pl(access, company_id, cust["id"], start, end))
            row.billed_to_date = float(totals.get("income", 0.0) or 0.0)
            row.costs_to_date = ((totals.get("cogs", 0.0) or 0.0)
                                 + (totals.get("expenses", 0.0) or 0.0))
        except Exception as e:
            row.status_flags.append(f"QBO P&L Failed: {type(e).__name__}")
        if n % 20 == 0:
            print(f"    …{n}/{len(pairs)} lines enriched")


# ─────────────────────────── main ──────────────────────────────────
def main() -> int:
    ap = argparse.ArgumentParser(
        description="RP WIP — General List (Alpha + Small Jobs) → slab/-FTW "
                    "lines → QBO billed/costs → 'Test - RP' tab.")
    ap.add_argument("--alpha", help="override General List path")
    ap.add_argument("--root", help="override Residential root")
    ap.add_argument("--project", help="filter to one job # (both lines)")
    ap.add_argument("--no-qbo", action="store_true", help="skip QBO join")
    ap.add_argument("--dry-run", action="store_true", help="preview, no write")
    args = ap.parse_args()
    alpha = Path(args.alpha) if args.alpha else ALPHA_PATH
    root = Path(args.root) if args.root else RP_ROOT

    print()
    print("  RP WIP Reader — General List → slab/-FTW lines → Test - RP")
    print(f"  list:  {alpha}")
    print(f"  root:  {root}")
    print("  " + "─" * 74)

    if not alpha.exists():
        print(f"  ✗ General List not found: {alpha}  (Synology mounted?)")
        return 1
    records, missing_sheets = read_general_list(alpha)
    for s in missing_sheets:
        print(f"  ⚠ sheet {s!r} not in the workbook — skipped")
    print(f"  Jobs in scope (priced or active): {len(records)}")

    rp_to_folders, addr_folders = ({}, [])
    if root.exists():
        rp_to_folders, addr_folders = index_residential(root)
    else:
        print(f"  ⚠ Residential root not found: {root} — no folder links")

    pairs = build_lines(records, rp_to_folders, addr_folders)
    if args.project:
        pf = args.project.upper()
        pairs = [t for t in pairs
                 if t[0].project_num in (pf, f"{pf}-FTW")]
    print(f"  Lines (slab + -FTW + CP standalone): {len(pairs)}")
    if not pairs:
        print("  No RP lines to process — exiting")
        return 0

    if not args.no_qbo:
        print("  Enriching with QBO Billed/Costs …")
        enrich_with_qbo(pairs)
    for row, comp, _rec in pairs:
        _classify(row, comp)

    # FTW backlog → its own section at the BOTTOM of the tab (the user
    # 2026-07-14): flatwork bid together with the slab but not poured yet —
    # effectively unwon-but-expected work (~95% follows the slab). Not an
    # error state, so no red; invoiced under RP#-FTW when the pour lands.
    sched_addrs, sched_label = read_schedule_flatwork(SCHEDULE_DIR)
    if sched_label:
        print(f"  Schedule check: {sched_label} — "
              f"{len(sched_addrs)} flatwork address(es)")
    main_rows, backlog = [], []
    for row, comp, rec in pairs:
        if not row.project_num.endswith("-FTW"):
            main_rows.append(row)
            continue
        addr_n = _norm(f"{rec['house'] or ''} {rec['street'] or ''}")
        on_sched = bool(addr_n) and any(
            addr_n == a or a.startswith(addr_n) or addr_n.startswith(a)
            for a in sched_addrs)
        has_activity = bool(row.billed_to_date) or bool(row.costs_to_date)
        # WON + WORKING when there is ANY QBO activity (billed OR costs —
        # the user 2026-07-14: "how can this be a backlog if there are
        # costs? we won it!") or the job is on today's flatwork schedule.
        if has_activity or on_sched:
            if on_sched:
                row.notes.append(
                    f"On the {sched_label} schedule — flatwork crew")
            main_rows.append(row)
        else:
            row.needs_review = False   # expected pre-pour state, not an error
            backlog.append(row)
    rows = main_rows + backlog              # for counts/logs
    n_red = sum(1 for r in rows if r.needs_review)
    print(f"  Review (red): {n_red} line(s) · FTW backlog: {len(backlog)}")

    # Justification dump — one record per line so every WIP verdict can be
    # backed up job-by-job with the ops manager (the user 2026-07-14): where
    # it was found in the General List, completion, contract, QBO billed,
    # and the rule that produced the status.
    dump = []
    for row, comp, rec in pairs:
        dump.append({
            "line": row.project_num,
            "section": ("FTW BACKLOG" if row in backlog else "MAIN"),
            "gl_sheet": rec["source"], "gl_row": rec.get("gl_row"),
            "completion": comp,
            "contract": row.base_contract, "etc": row.base_etc,
            "billed": row.billed_to_date, "costs": row.costs_to_date,
            "status": "Closed" if row.is_completed else "Active",
            "red": row.needs_review,
            "notes": list(row.notes), "flags": list(row.status_flags),
        })
    log_dir = Path.home() / "Library" / "Logs" / "Proficient" / "wip"
    log_dir.mkdir(parents=True, exist_ok=True)
    (log_dir / "rp_wip_lines.json").write_text(json.dumps(dump, indent=1))

    try:
        wrote = CP.write_test_cp(
            main_rows, CP.WIP_EXCEL_PATH,
            dry_run=args.dry_run, tab_name="Test - RP",
            appendix=("FTW BACKLOG — flatwork bid with the slab, NOT poured "
                      "yet (expected wins; invoice under RP#-FTW when poured)",
                      backlog))
    except CP.WipWriteDenied as e:
        print(f"  ✗ Guard blocked write: {e}")
        return 2
    except FileNotFoundError as e:
        print(f"  ✗ {e}")
        return 3
    if not args.dry_run and wrote:
        print(f"  ✓ Wrote {len(rows)} line(s) to 'Test - RP'")
    print()
    return 0


if __name__ == "__main__":
    sys.exit(main())
