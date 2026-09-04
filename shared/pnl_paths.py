#!/usr/bin/env python3
"""
pnl_paths.py — resolve where a project's P&L workbook lives (and when it was last pulled).

The ONE resolver for the project-pnl output location, kept in shared/ so the ledger
dashboard can FIND an existing P&L and its "last pulled" time using EXACTLY the rules
project-pnl writes with — without importing that tool (repo rule: tools never import
tools; common code lives in shared/).

project-pnl writes  <folder>/Project_PnL_<proj>.xlsx  where <folder> is:
  * non-CP : <division folder>/<proj>
  * CP     : the Common-drive awarded folder's 'Profit and Loss' subfolder, falling
             back to <division folder>/<proj> when the drive isn't mounted or
             nothing matches.
The DIVISION folder is `division_dir` — the OneDrive 'Commercial'/'Multi-Family'/
'Residential' folder, or, when a division is mapped to a TEAMS CHANNEL and that
channel is synced on this Mac, the channel's folder (MFD → 'Project Financials',
the user 2026-09-03). <out> defaults to OneDrive 'Automations-/PROJECT P&Ls'
(ACB_PNL_OUT_DIR); a single division can be pinned with ACB_PNL_DIR_<DIV>.

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
# A FINISHED job's workbook is named FINAL (the owner 2026-09-04: "for
# completed just rename the P&L to final, for actives just use the original
# name of Project Pnl"). One file per job either way - the separate
# `<job> Job Result.xlsx` and `<job> FINAL Closeout.xlsx` reports were retired
# 2026-09-03 and their 25 leftover copies deleted 2026-09-04.
PNL_FILE_FINAL = "{proj} FINAL.xlsx"


def pnl_filename(proj: str, archived: bool) -> str:
    """The workbook name for this job - FINAL once it is filed under an
    archive folder, `Project_PnL_<proj>` while it is live."""
    return (PNL_FILE_FINAL if archived else PNL_FILE).format(proj=proj)


def is_archived_dir(folder: "Path") -> bool:
    """True when `folder` sits inside a 'completed …' archive."""
    return any(str(x.name).lower().startswith(ARCHIVE_PREFIXES)
               for x in Path(folder).parents)

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


# ── one folder per division, and P&Ls are SORTED into it ─────────────────
# Binding (the user 2026-08-31): "if a p&l routes to this folder you must add a
# new rule to sort in the division folders. i want to send the folder links but
# can't show other pms other p&l that arent' theres." The OneDrive folder link
# is the unit of sharing, so a division folder has to hold ONLY that division's
# jobs - a P&L landing at the root would expose every PM's numbers to whoever
# has the link. Folder names match what is already on OneDrive.
DIVISION_DIRS = {"CP": "Commercial", "MFD": "Multi-Family", "RP": "Residential"}


def division_of(proj: str) -> str:
    """'MFD' | 'CP' | 'RP' from a project #, or '' when it is none of them."""
    pu = (proj or "").strip().upper()
    for pre in ("MFD", "CP", "RP"):        # MFD first - longest prefix wins
        if pu.startswith(pre):
            return pre
    return ""


# ── a division folder may live in a TEAMS CHANNEL ────────────────────────
# (the user 2026-09-03): "make new python route Multifamily p&l to
# <the Project Financials channel>". A Teams channel's Files tab IS a
# SharePoint folder, so once the channel is synced it is an ordinary path on
# this Mac - exactly the shape of 'Company Files - WIP Report', which the wip/
# readers have written to for months. No Graph API and no new key: the route
# is a folder name, and OneDrive does the moving.
#
# MOVE, not mirror (the user 2026-09-03): the channel is the ONLY home for the
# division once it is synced. Two copies of 'MFD Overview.xlsx' would drift,
# and the folder link the owner shares has to be the one with the live numbers.
DIVISION_CHANNELS = {"MFD": "Project Financials"}

# A synced channel shows up under one of TWO naming schemes, because OneDrive
# syncs a team channel and a shared library differently:
#   * personal OneDrive root : '<Team> - <Channel>'
#       'Company Files - WIP Report'          (what the wip/ readers use)
#   * shared-library root    : '<Site> - <Library>', the site itself named
#       '<Team>-<Channel>'
#       'OneDrive-SharedLibraries-<Org>/Multi-Family-Project Financials - Documents'
# So: drop a trailing library segment, then require the channel at the END of
# what is left, on a separator boundary. Nothing looser - a bare substring
# match would happily route the book into any folder that merely mentions the
# words, and a P&L in the wrong folder is the exact leak division_dir exists
# to prevent.
_LIBRARY_SUFFIX_RE = re.compile(r"\s+-\s+(Documents|Shared Documents)$", re.I)
_channel_cache: dict = {}


def _channel_roots():
    """Every place a sync can land: the personal OneDrive root, plus each
    'OneDrive-SharedLibraries-*' root beside it."""
    roots = [paths.onedrive_base()]
    for r in roots[0].parent.glob("OneDrive-SharedLibraries-*"):
        if r.is_dir():
            roots.append(r)
    return roots


def channel_dir(channel: str) -> "Path | None":
    """The local folder of a synced Teams channel, or None when it isn't synced.

    Deterministic on a tie: an exact-name folder wins over a '<Team> - <Channel>'
    one, and roots are searched in a fixed order, so two syncs of the same
    channel can never send successive runs to different folders."""
    if channel in _channel_cache:
        return _channel_cache[channel]
    tail = re.compile(r"(?:^|[\s\-–])" + re.escape(channel) + r"$", re.I)
    exact, partial = None, None
    for root in _channel_roots():
        try:
            children = sorted(root.iterdir())
        except OSError:
            continue
        for child in children:
            if not child.is_dir():
                continue
            name = _LIBRARY_SUFFIX_RE.sub("", child.name).strip()
            if name.lower() == channel.lower():
                exact = exact or child
            elif partial is None and tail.search(name):
                partial = child
    found = exact or partial
    # A CloudStorage folder can be WRITABLE but not LISTABLE: macOS grants the
    # write and refuses `iterdir` until the running app has Full Disk Access
    # (2026-09-03, the shared-library sync of this very channel). That half
    # state is worse than no sync at all - workbooks would land there while
    # `_iter_jobs` saw an empty folder, so the Overview would quietly rebuild
    # from nothing and a finished job would regenerate outside its archive.
    # Require the listing, or treat the channel as not synced.
    if found is not None and not _listable(found):
        found = None
    _channel_cache[channel] = found
    return found


def _listable(d: Path) -> bool:
    """True when this process can actually enumerate `d` (see channel_dir)."""
    try:
        next(iter(d.iterdir()), None)
        return True
    except OSError:
        return False


def division_dir(proj: str, out_dir: "Path | None" = None,
                 forced: bool = False) -> Path:
    """The division folder a project's P&L belongs in."""
    return division_dir_note(proj, out_dir, forced)[0]


def division_dir_note(proj: str, out_dir: "Path | None" = None,
                      forced: bool = False):
    """(folder, note) — the division folder, and why, when it isn't the default.

    Resolution order: an explicit --out (`forced`) beats every route · an
    `ACB_PNL_DIR_<DIV>` override · the division's Teams channel when synced ·
    the OneDrive division folder.

    An unrecognised project # stays at the root rather than being filed into
    the wrong division - a misfiled P&L is exactly the leak this rule exists
    to prevent."""
    out_dir = out_dir or pnl_out_dir()
    div = division_of(proj)
    name = DIVISION_DIRS.get(div)
    default = out_dir / name if name else out_dir
    if forced:                       # someone named an output folder; obey it
        return default, None
    override = paths.get("ACB_PNL_DIR_" + div) if div else ""
    if override:
        return Path(override).expanduser(), f"ACB_PNL_DIR_{div} override"
    channel = DIVISION_CHANNELS.get(div)
    if channel:
        found = channel_dir(channel)
        if found is not None:
            return found, None
        # Fall back rather than fail a run - but SAY SO. A silent fallback is
        # how the owner ends up sharing a channel link to a folder the numbers
        # never reached.
        return default, (f"Teams channel '{channel}' is not readable on this Mac "
                         f"(not synced, or synced as a shared library this app "
                         f"cannot list) → wrote to {default.name} on OneDrive "
                         f"instead")
    return default, None


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
    base, dnote = division_dir_note(proj, out_dir)
    if not proj.upper().startswith("CP"):
        return base / proj, dnote
    try:
        mounted = CP_AWARDED_BASE.exists()
    except OSError:
        mounted = False
    if not mounted:
        return base / proj, "Common drive not mounted → OneDrive"
    folder = _find_awarded_cp_folder(CP_AWARDED_BASE, proj)
    if folder is None:
        return base / proj, f"no awarded folder for {proj} → OneDrive"
    return folder / CP_PNL_SUBDIR, None


def pnl_path(proj: str, out_dir: "Path | None" = None) -> Path:
    """The exact workbook path project-pnl would write for this project."""
    folder, _ = resolve_project_out_dir(proj, out_dir)
    return folder / pnl_filename(proj, is_archived_dir(folder))


# Subfolders of PROJECT P&Ls that hold FINISHED jobs. Matched case-insensitively
# by prefix so "completed mfd project p&l", "Completed CP …" and friends all
# count without needing to be listed one by one.
ARCHIVE_PREFIXES = ("completed", "closed", "archive")


def _archive_dirs():
    """Existing archive subfolders, under the P&L root AND inside each division
    folder ('Multi-Family/completed mfd project p&l'). The root is still swept
    for archives filed before the division rule."""
    out: list = []
    root = pnl_out_dir()
    # Resolved, not hard-coded: a division routed to a Teams channel keeps its
    # 'completed …' archive inside that channel, and a re-run has to regenerate
    # a finished job THERE rather than spawn a second copy at the top level.
    # THE ROUTED DIVISION FOLDER COMES FIRST. A finished job is regenerated
    # into the first archive that already holds it, so when the same archive
    # exists in the old OneDrive tree AND in the synced Teams channel, the
    # order here decides where the live copy lands. 2026-09-03: the old tree
    # was listed first and 11 finished MFD jobs regenerated into the folder
    # nobody reads any more while the channel kept the morning's copies.
    bases = []
    for _d in DIVISION_DIRS:
        _r = division_dir(_d)
        if _r not in bases:
            bases.append(_r)
    for _b in [root] + [root / n for n in DIVISION_DIRS.values()]:
        if _b not in bases:
            bases.append(_b)
    for base in bases:
        try:
            out += [d for d in sorted(base.iterdir())
                    if d.is_dir()
                    and d.name.lower().startswith(ARCHIVE_PREFIXES)]
        except OSError:
            continue
    return out


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
    fnames = (fname, PNL_FILE_FINAL.format(proj=proj))
    add(pnl_path(proj))                                   # exact resolved path
    add(division_dir(proj) / proj / fname)                # division-sorted
    _dn = DIVISION_DIRS.get(division_of(proj))
    if _dn:                                              # pre-Teams-move home
        add(pnl_out_dir() / _dn / proj / fname)
    add(pnl_out_dir() / proj / fname)                    # pre-division root subfolder
    # Finished jobs are filed away under an ARCHIVE subfolder (the user
    # 2026-08-27) so the top level stays the live work. Look there too, or the
    # dashboard reports "never generated" the moment a job is filed.
    for _arch in _archive_dirs():
        for _fn in fnames:                      # FINAL is the archived name
            add(_arch / proj / _fn)
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
