#!/usr/bin/env python3
"""
pnl_paths.py — resolve where a project's P&L workbook lives (and when it was last pulled).

The ONE resolver for the project-pnl output location, kept in shared/ so the ledger
dashboard can FIND an existing P&L and its "last pulled" time using EXACTLY the rules
project-pnl writes with — without importing that tool (repo rule: tools never import
tools; common code lives in shared/).

project-pnl writes  <folder>/Project_PnL_<proj>.xlsx  where <folder> is:
  * non-CP : <out>/<proj>
  * CP     : the Common-drive awarded folder's 'Profit and Loss' subfolder, falling
             back to <out>/<proj> when the drive isn't mounted or nothing matches.
<out> defaults to OneDrive 'Automations-/PROJECT P&Ls' (ACB_PNL_OUT_DIR override).

NOTE (dedupe follow-up): project-pnl still keeps its own copy of this logic
(_resolve_project_out_dir / _find_awarded_cp_folder). It should import these instead so
the two can't drift — deferred here only to avoid touching the 326 KB export script while
a concurrent session is editing shared/. Same pattern as shared/qbo_costs.py (cost_leaf).
"""
from __future__ import annotations

import datetime as _dt
import os
import re
from pathlib import Path

from shared import paths

# CP-only: Commercial P&Ls drop into the awarded-project folder on the Common drive
# (mirrors where the WIP report lives). Overridable for other machines / mounts.
CP_AWARDED_BASE = Path(os.environ.get(
    "ACB_CP_AWARDED_BASE",
    "/Volumes/Common/CURRENT PROJECTS/Awarded Projects Commercial projects"))
CP_PNL_SUBDIR = "Profit and Loss"
PNL_FILE = "Project_PnL_{proj}.xlsx"

# RP source jobs live under the Residential share, filed BY BUILDER then address
# (same root rp_wip_reader.py uses). We can land on the builder folder reliably;
# the exact job inside needs the General-List address index (rp_wip_reader).
RP_SOURCE_BASE = Path(os.environ.get(
    "RP_ROOT", "/Volumes/Common/CURRENT PROJECTS/Residential"))


def job_folder(proj: str, builder: "str | None" = None):
    """(folder, note) — the SOURCE job folder on the file server, or (None, note).

    CP → the awarded-project folder (Synology Common drive, matched by #). RP →
    best-effort match of the BUILDER's folder under Residential (the exact address
    folder is inside). MFD moves a lot → no stable Synology source (None; the caller
    falls back to the OneDrive P&L folder). `note` explains a miss for the UI.
    """
    pu = (proj or "").upper()
    if pu.startswith("CP"):
        try:
            if not CP_AWARDED_BASE.is_dir():
                return None, "Common drive not mounted"
        except OSError:
            return None, "Common drive not mounted"
        f = _find_awarded_cp_folder(CP_AWARDED_BASE, proj)
        return (f, None) if f else (None, f"no awarded folder for {proj}")
    if pu.startswith("RP"):
        try:
            if not RP_SOURCE_BASE.is_dir():
                return None, "Residential drive not mounted"
        except OSError:
            return None, "Residential drive not mounted"
        if builder:
            b = re.sub(r"[\s\-_,.]+", "", builder.upper())
            if b:
                for child in sorted(RP_SOURCE_BASE.iterdir()):
                    if child.is_dir() and b in re.sub(r"[\s\-_,.]+", "", child.name.upper()):
                        return child, "builder folder — the job is inside, by address"
        return RP_SOURCE_BASE, "Residential root (builder folder not matched)"
    return None, "MFD moves a lot → filed in OneDrive, not Synology"


def pnl_out_dir() -> Path:
    """The default P&L output root (ACB_PNL_OUT_DIR or OneDrive 'PROJECT P&Ls')."""
    return paths.get_path("ACB_PNL_OUT_DIR",
                          paths.onedrive_base() / "Automations-/PROJECT P&Ls")


def _find_awarded_cp_folder(base: Path, proj: str):
    """Awarded-project folder for a CP job, matched by project # (full match wins,
    bare-number match on a digit boundary as fallback). None if base unreachable."""
    try:
        if not base.is_dir():
            return None
    except OSError:
        return None
    pu = proj.upper()
    num = re.sub(r"\D", "", proj)          # 'CP672' -> '672'
    numbered = None
    for child in sorted(base.iterdir()):
        if not child.is_dir():
            continue
        compact = re.sub(r"[\s\-_]+", "", child.name.upper())
        if pu in compact:                  # strongest: name carries 'CP672'
            return child
        if num and numbered is None and re.search(rf"(?<!\d){num}(?!\d)", child.name):
            numbered = child               # weaker: bare-number match, keep first
    return numbered


def resolve_project_out_dir(proj: str, out_dir: "Path | None" = None):
    """(folder, note) — where project-pnl would put this project's workbook.
    `note` explains any CP → OneDrive fallback (surfaced in the UI)."""
    out_dir = out_dir or pnl_out_dir()
    if not proj.upper().startswith("CP"):
        return out_dir / proj, None
    try:
        mounted = CP_AWARDED_BASE.exists()
    except OSError:
        mounted = False
    if not mounted:
        return out_dir / proj, "Common drive not mounted → OneDrive"
    folder = _find_awarded_cp_folder(CP_AWARDED_BASE, proj)
    if folder is None:
        return out_dir / proj, f"no awarded folder for {proj} → OneDrive"
    return folder / CP_PNL_SUBDIR, None


def pnl_path(proj: str, out_dir: "Path | None" = None) -> Path:
    """The exact workbook path project-pnl would write for this project."""
    folder, _ = resolve_project_out_dir(proj, out_dir)
    return folder / PNL_FILE.format(proj=proj)


# Subfolders of PROJECT P&Ls that hold FINISHED jobs. Matched case-insensitively
# by prefix so "completed mfd project p&l", "Completed CP …" and friends all
# count without needing to be listed one by one.
ARCHIVE_PREFIXES = ("completed", "closed", "archive")


def _archive_dirs():
    """Existing archive subfolders under the P&L output root."""
    try:
        root = pnl_out_dir()
        return [d for d in sorted(root.iterdir())
                if d.is_dir()
                and d.name.lower().startswith(ARCHIVE_PREFIXES)]
    except OSError:
        return []


def _candidates(proj: str):
    """Every place this project's workbook might exist (newest one wins in find_pnl).
    Covers the resolved path (incl. Common drive if mounted), the per-project subfolder,
    and the older flat 'PROJECT P&Ls/' locations."""
    seen: set = set()
    out: list = []

    def add(p: Path):
        if p and str(p) not in seen:
            seen.add(str(p))
            out.append(p)

    fname = PNL_FILE.format(proj=proj)
    add(pnl_path(proj))                                   # exact resolved path
    add(pnl_out_dir() / proj / fname)                    # default per-project subfolder
    # Finished jobs are filed away under an ARCHIVE subfolder (the user
    # 2026-08-27) so the top level stays the live work. Look there too, or the
    # dashboard reports "never generated" the moment a job is filed.
    for _arch in _archive_dirs():
        add(_arch / proj / fname)
    ob = paths.onedrive_base()
    add(ob / "PROJECT P&Ls" / proj / fname)              # older flat tree, subfolder
    add(ob / "PROJECT P&Ls" / fname)                     # older flat tree, root
    return out


def find_pnl(proj: str) -> dict:
    """{exists, path, mtime, note} for the NEWEST existing P&L workbook of `proj`.
    `mtime` is ISO-minutes local time — the "last pulled" the owner asked to see;
    None when nothing has been generated yet."""
    _, note = resolve_project_out_dir(proj)
    best = None
    for p in _candidates(proj):
        try:
            st = p.stat()
        except OSError:
            continue
        if best is None or st.st_mtime > best[1]:
            best = (p, st.st_mtime)
    if best is None:
        return {"exists": False, "path": str(pnl_path(proj)), "mtime": None, "note": note}
    p, m = best
    return {"exists": True, "path": str(p),
            "mtime": _dt.datetime.fromtimestamp(m).isoformat(timespec="minutes"),
            "note": note}
