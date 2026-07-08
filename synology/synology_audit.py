#!/usr/bin/env python3
"""
synology_audit.py — age-distribution audit of one or more Synology shares.

Built for Proficient Concrete to figure out what on the Synology is *active*
vs *archive* before committing to any tree extract or migration. Same --root
interface as synology_tree.py but emits a stats report instead of a full tree.

USAGE
-----
    # Single share
    python3 synology_audit.py --root "/Volumes/Common"

    # Multiple shares, combined report
    python3 synology_audit.py \\
        --root "/Volumes/Common" \\
        --root "/Volumes/Accounting" \\
        --root "/Volumes/Field" \\
        --root "/Volumes/Multi Family" \\
        --root "/Volumes/Proinfo" \\
        --title "Proficient Concrete"

OUTPUT
------
A markdown file with:
  - Top-level summary (totals, newest/oldest file dates across all shares)
  - Aggregate age-distribution table (last 12mo, 1-2yr, 2-5yr, >5yr)
  - Per-share section: counts + age distribution + top-level folders sorted
    newest -> oldest (so the active/archive cliff is visible)
  - Aggregated error block at the bottom

Same noise filters as synology_tree.py (@eaDir, .DS_Store, ~$* lockfiles, etc.).
"""

from __future__ import annotations

import argparse
import os
import sys
import time
from datetime import datetime
from pathlib import Path

# ---------------------------------------------------------------------------
# Filters (mirror synology_tree.py)
# ---------------------------------------------------------------------------

SKIP_DIRS = {
    "@eaDir",
    ".Spotlight-V100",
    ".Trashes",
    ".fseventsd",
    ".TemporaryItems",
    "#recycle",
    ".DocumentRevisions-V100",
    ".com.apple.timemachine.donotpresent",
    "$RECYCLE.BIN",
}

SKIP_FILES_EXACT = {".DS_Store", "Thumbs.db", "desktop.ini", ".localized"}


def is_skipped_file(name: str) -> bool:
    if name in SKIP_FILES_EXACT:
        return True
    if name.startswith("~$"):
        return True
    if name.startswith(".~lock."):
        return True
    return False


# ---------------------------------------------------------------------------
# Buckets and formatters
# ---------------------------------------------------------------------------

DAY = 86400.0
BUCKETS = [
    ("Last 12 months", 0.0, 365 * DAY),
    ("1–2 years", 365 * DAY, 730 * DAY),
    ("2–5 years", 730 * DAY, 5 * 365 * DAY),
    ("Older than 5 years", 5 * 365 * DAY, float("inf")),
]


def bucket_for(age_seconds: float) -> str:
    for name, lo, hi in BUCKETS:
        if lo <= age_seconds < hi:
            return name
    return BUCKETS[-1][0]


def fmt_size(n: float) -> str:
    size = float(n)
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if size < 1024 or unit == "TB":
            if unit == "B":
                return f"{int(size)} B"
            return f"{size:.1f} {unit}"
        size /= 1024
    return f"{size:.1f} PB"


def fmt_date(ts: float) -> str:
    if ts <= 0 or ts == float("inf"):
        return "—"
    return datetime.fromtimestamp(ts).strftime("%Y-%m-%d")


# ---------------------------------------------------------------------------
# Stats containers
# ---------------------------------------------------------------------------

class FolderStats:
    __slots__ = ("name", "path", "files", "bytes", "newest", "oldest")

    def __init__(self, name: str, path: str) -> None:
        self.name = name
        self.path = path
        self.files = 0
        self.bytes = 0
        self.newest = 0.0
        self.oldest = float("inf")


class RootStats:
    def __init__(self, root: Path) -> None:
        self.root = root
        self.files = 0
        self.folders = 0
        self.bytes = 0
        self.bucket_files: dict[str, int] = {b[0]: 0 for b in BUCKETS}
        self.bucket_bytes: dict[str, int] = {b[0]: 0 for b in BUCKETS}
        self.newest: float = 0.0
        self.oldest: float = float("inf")
        self.newest_path: str = ""
        self.oldest_path: str = ""
        self.folder_stats: dict[str, FolderStats] = {}
        self.errors: list[str] = []


# ---------------------------------------------------------------------------
# Walker
# ---------------------------------------------------------------------------

def _record_file(
    rs: RootStats,
    top: FolderStats | None,
    path: str,
    st: os.stat_result,
    now: float,
) -> None:
    sz = st.st_size
    mt = st.st_mtime
    rs.files += 1
    rs.bytes += sz
    b = bucket_for(now - mt)
    rs.bucket_files[b] += 1
    rs.bucket_bytes[b] += sz
    if mt > rs.newest:
        rs.newest = mt
        rs.newest_path = path
    if mt < rs.oldest:
        rs.oldest = mt
        rs.oldest_path = path
    if top is not None:
        top.files += 1
        top.bytes += sz
        if mt > top.newest:
            top.newest = mt
        if mt < top.oldest:
            top.oldest = mt


def is_excluded(path: Path, excludes: set[Path]) -> bool:
    """True if path is at or under any --exclude path."""
    if not excludes:
        return False
    for ex in excludes:
        if path == ex:
            return True
        try:
            path.relative_to(ex)
            return True
        except ValueError:
            continue
    return False


def _recurse(
    d: str,
    top: FolderStats | None,
    rs: RootStats,
    now: float,
    excludes: set[Path],
) -> None:
    try:
        entries = list(os.scandir(d))
    except PermissionError:
        rs.errors.append(f"{d}: permission denied")
        return
    except OSError as e:
        rs.errors.append(f"{d}: {e}")
        return

    for entry in entries:
        try:
            if entry.is_dir(follow_symlinks=False):
                if entry.name in SKIP_DIRS:
                    continue
                if is_excluded(Path(entry.path).resolve(), excludes):
                    continue
                rs.folders += 1
                _recurse(entry.path, top, rs, now, excludes)
            elif entry.is_file(follow_symlinks=False):
                if is_skipped_file(entry.name):
                    continue
                try:
                    st = entry.stat(follow_symlinks=False)
                except OSError:
                    continue
                _record_file(rs, top, entry.path, st, now)
        except OSError:
            continue


def audit(root: Path, now: float, excludes: set[Path]) -> RootStats:
    rs = RootStats(root)
    try:
        top_entries = list(os.scandir(root))
    except OSError as e:
        rs.errors.append(f"{root}: {e}")
        return rs

    for entry in top_entries:
        try:
            if entry.is_dir(follow_symlinks=False):
                if entry.name in SKIP_DIRS:
                    continue
                if is_excluded(Path(entry.path).resolve(), excludes):
                    continue
                rs.folders += 1
                fs = FolderStats(entry.name, entry.path)
                rs.folder_stats[entry.name] = fs
                _recurse(entry.path, fs, rs, now, excludes)
            elif entry.is_file(follow_symlinks=False):
                if is_skipped_file(entry.name):
                    continue
                try:
                    st = entry.stat(follow_symlinks=False)
                except OSError:
                    continue
                _record_file(rs, None, entry.path, st, now)
        except OSError:
            continue

    return rs


# ---------------------------------------------------------------------------
# Markdown emission
# ---------------------------------------------------------------------------

def emit_markdown(
    roots_stats: list[RootStats],
    title: str | None,
    now: float,
) -> str:
    if title is None:
        title = roots_stats[0].root.name if len(roots_stats) == 1 else "Combined"

    agg_files = sum(rs.files for rs in roots_stats)
    agg_folders = sum(rs.folders for rs in roots_stats)
    agg_bytes = sum(rs.bytes for rs in roots_stats)
    agg_bucket_files = {b[0]: 0 for b in BUCKETS}
    agg_bucket_bytes = {b[0]: 0 for b in BUCKETS}
    for rs in roots_stats:
        for b_name in agg_bucket_files:
            agg_bucket_files[b_name] += rs.bucket_files[b_name]
            agg_bucket_bytes[b_name] += rs.bucket_bytes[b_name]

    nonempty = [rs for rs in roots_stats if rs.files > 0]
    agg_newest = max((rs.newest for rs in nonempty), default=0.0)
    agg_oldest = min((rs.oldest for rs in nonempty), default=float("inf"))

    out: list[str] = []
    out.append(f"# Synology Age Audit — {title}")
    out.append("")
    out.append(f"- **Roots:** {', '.join(rs.root.name for rs in roots_stats)}")
    out.append(
        f"- **Generated:** {datetime.fromtimestamp(now).strftime('%Y-%m-%d %H:%M:%S')}"
    )
    out.append(f"- **Total files:** {agg_files:,}")
    out.append(f"- **Total folders:** {agg_folders:,}")
    out.append(f"- **Total size:** {fmt_size(agg_bytes)}")
    out.append(f"- **Newest file:** {fmt_date(agg_newest)}")
    out.append(f"- **Oldest file:** {fmt_date(agg_oldest)}")
    out.append("")

    out.append("## Age distribution — all shares")
    out.append("")
    out.append("| Bucket | Files | Size | % files | % size |")
    out.append("|---|---:|---:|---:|---:|")
    for b_name, _, _ in BUCKETS:
        f = agg_bucket_files[b_name]
        s = agg_bucket_bytes[b_name]
        pf = (f / agg_files * 100) if agg_files else 0.0
        ps = (s / agg_bytes * 100) if agg_bytes else 0.0
        out.append(f"| {b_name} | {f:,} | {fmt_size(s)} | {pf:.1f}% | {ps:.1f}% |")
    out.append("")

    for rs in roots_stats:
        out.append(f"## {rs.root.name}")
        out.append("")
        out.append(
            f"`{rs.root}` — {rs.files:,} files, {rs.folders:,} folders, "
            f"{fmt_size(rs.bytes)}"
        )
        if rs.files:
            out.append(
                f"Newest: **{fmt_date(rs.newest)}** · "
                f"Oldest: **{fmt_date(rs.oldest)}**"
            )
        out.append("")

        out.append("**Age distribution:**")
        out.append("")
        out.append("| Bucket | Files | Size | % files |")
        out.append("|---|---:|---:|---:|")
        for b_name, _, _ in BUCKETS:
            f = rs.bucket_files[b_name]
            s = rs.bucket_bytes[b_name]
            pf = (f / rs.files * 100) if rs.files else 0.0
            out.append(f"| {b_name} | {f:,} | {fmt_size(s)} | {pf:.1f}% |")
        out.append("")

        if rs.folder_stats:
            sorted_folders = sorted(
                rs.folder_stats.values(),
                key=lambda f: f.newest,
                reverse=True,
            )
            out.append(
                "**Top-level folders** "
                "(sorted by most-recent file modified, newest → oldest):"
            )
            out.append("")
            out.append("| Folder | Files | Size | Newest | Oldest |")
            out.append("|---|---:|---:|---|---|")
            for fs in sorted_folders:
                out.append(
                    f"| {fs.name}/ | {fs.files:,} | {fmt_size(fs.bytes)} | "
                    f"{fmt_date(fs.newest)} | {fmt_date(fs.oldest)} |"
                )
            out.append("")

        if rs.errors:
            out.append("**Errors:**")
            out.append("")
            for err in rs.errors[:20]:
                out.append(f"- {err}")
            if len(rs.errors) > 20:
                out.append(f"- _… and {len(rs.errors) - 20} more_")
            out.append("")

    return "\n".join(out)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description=(
            "Audit one or more Synology shares: age distribution, top-level "
            "folder ages, total size."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p.add_argument(
        "--root",
        action="append",
        required=True,
        metavar="PATH",
        help="Path to a share or folder. Pass --root multiple times to combine shares.",
    )
    p.add_argument(
        "--out",
        default=None,
        help=(
            "Output markdown file. Default: ~/Library/Logs/Proficient/synology/synology_audit.md "
            "(project folder is AI-visible — company file/folder names stay outside)."
        ),
    )
    p.add_argument(
        "--title",
        default=None,
        help=(
            "Title for the report header. Defaults to share name when one "
            "--root, or 'Combined' when multiple."
        ),
    )
    p.add_argument(
        "--exclude",
        action="append",
        default=[],
        metavar="PATH",
        help=(
            "Folder to skip entirely (no recursion, no listing). Pass multiple "
            "times. Use full paths. Skips also apply to all subfolders."
        ),
    )
    return p.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)

    roots: list[Path] = []
    for raw in args.root:
        path = Path(raw).expanduser().resolve()
        if not path.exists():
            print(f"ERROR: root does not exist: {path}", file=sys.stderr)
            return 2
        if not path.is_dir():
            print(f"ERROR: root is not a directory: {path}", file=sys.stderr)
            return 2
        roots.append(path)

    excludes: set[Path] = set()
    for raw in args.exclude:
        ex = Path(raw).expanduser().resolve()
        excludes.add(ex)

    # Default output lives OUTSIDE the project folder (privacy: project folder
    # is AI-session-visible; Synology audit includes company file/folder names).
    out_path = (
        Path(args.out).expanduser().resolve()
        if args.out
        else Path.home() / "Library/Logs/Proficient/synology/synology_audit.md"
    )
    out_path.parent.mkdir(parents=True, exist_ok=True)

    now = time.time()
    print(f"Auditing {len(roots)} share(s)…")
    if excludes:
        print(f"  excluding {len(excludes)} path(s)")
    roots_stats: list[RootStats] = []
    for r in roots:
        print(f"  scanning {r} …", flush=True)
        rs = audit(r, now, excludes)
        print(
            f"    {rs.files:,} files, {rs.folders:,} folders, "
            f"{fmt_size(rs.bytes)}, errors: {len(rs.errors)}"
        )
        roots_stats.append(rs)

    md = emit_markdown(roots_stats, args.title, now)
    out_path.write_text(md, encoding="utf-8")

    print(f"\nWrote {out_path}")
    total_files = sum(rs.files for rs in roots_stats)
    total_bytes = sum(rs.bytes for rs in roots_stats)
    total_errors = sum(len(rs.errors) for rs in roots_stats)
    print(
        f"  total: {total_files:,} files, "
        f"{fmt_size(total_bytes)}, errors: {total_errors}"
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
