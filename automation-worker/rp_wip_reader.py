#!/usr/bin/env python3
"""
rp_wip_reader.py — RP (Residential) WIP reader.

ARCHITECTURE (Ted 2026-07-02), parallels CP:
  • ACTIVE PROJECTS come from the ALPHA LIST (the RP equivalent of CP's
    "Awarded Projects"): a job is active when COMPLETION < 100%.
  • The RESIDENTIAL folder tree is the TAKEOFF LOOKUP — for each active RP#,
    find its takeoff file by RP# in the filename; if the RP# isn't in any
    filename, fall back to matching the project's ADDRESS to an address folder.

STEP 1 (this build): read Alpha → active RP#s, then locate each one's takeoff
folder in Residential (RP# match, else address). Prints a table so Ted can
audit the active list + whether every project's takeoff was found. Next steps:
pick the right takeoff/proposal per RP#, then contract price + ETC.

Sources (Synology, read-only):
  Alpha:       /Volumes/Common/OPERATIONS/GENERAL LIST/LISTA GENERAL AÑO 2026.xlsx
               sheet 'General list - Alpha order' (hdr row 4; JOB NUMBER col 3,
               HOUSE col 4, STREET col 5, CITY col 6, COMPLETION col 26)
  Takeoffs:    /Volumes/Common/CURRENT PROJECTS/Residential/<CLIENT>/<ADDRESS>/

Usage:
  python3 rp_wip_reader.py
  python3 rp_wip_reader.py --alpha <path> --root <path>

Read-only. Never writes.
"""
from __future__ import annotations

import argparse
import os
import re
import sys
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

from openpyxl import load_workbook

# Reuse the CP writer so RP gets the exact same WIP structure/formatting,
# just written to the 'Test - RP' tab.
import cp_wip_reader as CP

ALPHA_PATH = Path(os.getenv(
    "RP_ALPHA_PATH",
    "/Volumes/Common/OPERATIONS/GENERAL LIST/LISTA GENERAL AÑO 2026.xlsx",
))
ALPHA_SHEET = "General list - Alpha order"
RP_ROOT = Path(os.getenv(
    "RP_ROOT",
    "/Volumes/Common/CURRENT PROJECTS/Residential",
))

# Alpha column map (1-based), header row 4, data from row 6.
COL_JOB, COL_HOUSE, COL_STREET, COL_CITY, COL_COMPLETION = 3, 4, 5, 6, 26

_RP_RE = re.compile(r"RP\d{4}(?:-[A-Za-z]{2,6})?(?!\d)", re.IGNORECASE)


def _norm(s) -> str:
    return re.sub(r"\s+", " ", str(s or "").upper().strip())


# ─────────────────────── Alpha: active projects ────────────────────
def read_alpha_active(path: Path):
    """Return [{rp, completion, house, street, city}] for RP jobs with
    COMPLETION < 100% (active). 100% or fully-greyed = done, excluded."""
    wb = load_workbook(path, data_only=True)
    ws = wb[ALPHA_SHEET]
    out = []
    for r in range(6, ws.max_row + 1):
        job = ws.cell(r, COL_JOB).value
        comp = ws.cell(r, COL_COMPLETION).value
        if not job:
            continue
        m = _RP_RE.match(str(job).strip())
        if not m:
            continue
        if isinstance(comp, (int, float)) and comp < 1.0:
            out.append({
                "rp": m.group(0).upper(),
                "completion": comp,
                "house": ws.cell(r, COL_HOUSE).value,
                "street": ws.cell(r, COL_STREET).value,
                "city": ws.cell(r, COL_CITY).value,
            })
    wb.close()
    return out


# ─────────────────────── Residential: takeoff lookup ───────────────
def _list_dir(path: Path):
    """One directory listing via scandir → (subdirs, filenames). Errors → empty.
    is_dir()/is_file() come cached from scandir, so no extra stat per entry."""
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
    """List ONLY the client (depth 1) and address (depth 2) folders — never the
    INVOICES/PLANS/… subfolders. The address-folder listings are done IN PARALLEL
    (thread pool) so the Synology network round-trips overlap instead of running
    one-at-a-time (the real slowdown). Takeoffs live directly in the
    client/address folder, so scanning those filenames is enough.
    Returns rp_to_folders {RP#: set(folder Path)} and addr_folders
    [(normalized name, Path)]."""
    root = Path(root)
    rp_to_folders = defaultdict(set)
    addr_folders = []

    clients, _root_files = _list_dir(root)          # depth 1 = client folders
    with ThreadPoolExecutor(max_workers=workers) as ex:
        client_listings = list(ex.map(_list_dir, clients))

    addr_dirs = []
    for client, (subs, files) in zip(clients, client_listings):
        for fn in files:                            # takeoffs occasionally sit at client level
            for m in _RP_RE.findall(fn):
                rp_to_folders[m.upper()].add(client)
        addr_dirs.extend(subs)

    with ThreadPoolExecutor(max_workers=workers) as ex:
        addr_listings = list(ex.map(_list_dir, addr_dirs))  # depth 2, parallel — the fast path

    for addr, (_subs, files) in zip(addr_dirs, addr_listings):
        addr_folders.append((_norm(addr.name), addr))
        for fn in files:
            for m in _RP_RE.findall(fn):
                rp_to_folders[m.upper()].add(addr)
    return rp_to_folders, addr_folders


def match_by_address(proj, addr_folders):
    """Fallback: match Alpha house+street to an address folder name.
    Requires the house number to appear in the folder name AND a street-word
    overlap. Returns the folder Path or None (returns first strong match)."""
    house = _norm(proj["house"])
    street_words = [w for w in _norm(proj["street"]).split()
                    if w not in ("ROAD", "STREET", "ST", "DRIVE", "DR", "LANE",
                                 "LN", "AVENUE", "AVE", "COURT", "CT", "TRAIL",
                                 "CIRCLE", "WAY", "BLVD", "N", "S", "E", "W")]
    for norm_name, folder in addr_folders:
        if house and house in norm_name and any(w in norm_name for w in street_words):
            return folder
    return None


def resolve_takeoff(proj, rp_to_folders, addr_folders):
    """Return (method, [folders]) for one active project.
    method ∈ {'by RP#', 'by address', 'MISSING'}."""
    rp = proj["rp"]
    if rp in rp_to_folders:
        return "by RP#", sorted(rp_to_folders[rp], key=lambda f: (f.parent.name, f.name))
    f = match_by_address(proj, addr_folders) if addr_folders else None
    if f:
        return "by address", [f]
    return "MISSING", []


def build_rows(active, rp_to_folders, addr_folders):
    """Convert active Alpha projects → CP.CpRow objects for the WIP write.
    Only the identifier + folder link are populated now; contract/ETC/QBO fill
    in later. Flags carry completion %, the match method, and any warnings."""
    rows = []
    for p in sorted(active, key=lambda x: x["rp"]):
        method, folders = resolve_takeoff(p, rp_to_folders, addr_folders)
        name = _norm(f"{p['house']} {p['street']}") or (folders[0].name if folders else p["rp"])
        row = CP.CpRow(p["rp"], name, False, None, None, None, None, None)
        if folders:
            row.folder_path = folders[0]          # project-name → folder link
        row.status_flags.append(f"{p['completion']*100:.0f}% complete (Alpha)")
        if method == "by address":
            row.status_flags.append(f"takeoff by ADDRESS — verify ({folders[0].name})")
        elif method == "MISSING":
            row.status_flags.append(f"NO TAKEOFF FOUND ({p['house']} {p['street']}, {p['city']})")
        elif len(folders) > 1:
            row.status_flags.append("⚠ takeoff in " + str(len(folders)) + " folders: "
                                    + ", ".join(f"{f.parent.name}/{f.name}" for f in folders))
        rows.append(row)
    return rows


# ─────────────────────── report ────────────────────────────────────
def main() -> int:
    ap = argparse.ArgumentParser(description="RP active projects (Alpha) + takeoff lookup (Residential).")
    ap.add_argument("--alpha", help="override Alpha List path")
    ap.add_argument("--root", help="override Residential root")
    ap.add_argument("--dry-run", action="store_true", help="preview, don't write the WIP")
    args = ap.parse_args()
    alpha = Path(args.alpha) if args.alpha else ALPHA_PATH
    root = Path(args.root) if args.root else RP_ROOT

    print()
    print("  RP SLAB — active projects (Alpha) + takeoff lookup (Residential)")
    print(f"  alpha: {alpha}")
    print(f"  root:  {root}")
    print("  " + "─" * 74)

    if not alpha.exists():
        print(f"  ✗ Alpha not found: {alpha}  (Synology mounted?)")
        return 1
    active = read_alpha_active(alpha)
    print(f"  Active RP (completion < 100%): {len(active)}")

    rp_to_folders, addr_folders = ({}, [])
    if root.exists():
        rp_to_folders, addr_folders = index_residential(root)
    else:
        print(f"  ⚠ Residential root not found: {root} — listing active only, no takeoff lookup")

    print()
    print(f"  {'RP #':11} {'%':>5}  {'MATCH':11} FOLDER")
    print("  " + "─" * 74)
    n_rp, n_addr, n_missing = 0, 0, 0
    for p in sorted(active, key=lambda x: x["rp"]):
        rp, comp = p["rp"], p["completion"]
        if rp in rp_to_folders:
            folders = sorted(rp_to_folders[rp], key=lambda f: (f.parent.name, f.name))
            if len(folders) == 1:
                loc = folders[0].name
            else:
                loc = f"⚠ {len(folders)} folders: " + ", ".join(
                    f"{f.parent.name}/{f.name}" for f in folders)
            method = "by RP#"; n_rp += 1
        elif addr_folders and (f := match_by_address(p, addr_folders)):
            loc = f.name + "   (addr-matched — verify)"; method = "by address"; n_addr += 1
        elif not root.exists():
            loc = f"{p['house']} {p['street']}, {p['city']}"; method = "—"
        else:
            loc = f"NO TAKEOFF FOUND  ({p['house']} {p['street']}, {p['city']})"; method = "MISSING"; n_missing += 1
        print(f"  {rp:11} {comp*100:>4.0f}%  {method:11} {loc}")

    print("  " + "─" * 74)
    print(f"  {len(active)} active  ·  {n_rp} matched by RP#  ·  {n_addr} by address  ·  {n_missing} no takeoff")

    # ── Write into the WIP (Test - RP tab), same structure as CP ──
    rows = build_rows(active, rp_to_folders, addr_folders)
    try:
        wrote = CP.write_test_cp(rows, CP.WIP_EXCEL_PATH,
                                 dry_run=args.dry_run, tab_name="Test - RP")
    except CP.WipWriteDenied as e:
        print(f"  ✗ Guard blocked write: {e}")
        return 2
    except FileNotFoundError as e:
        print(f"  ✗ {e}")
        return 3
    if not args.dry_run and wrote:
        print(f"  ✓ Wrote {len(rows)} row(s) to 'Test - RP'")
    print()
    return 0


if __name__ == "__main__":
    sys.exit(main())
