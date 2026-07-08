#!/usr/bin/env python3
"""
synology_tree.py — generate a markdown tree of one or more Synology shares.

Built for Proficient Concrete to inventory the Synology so files and folders
can be referenced as "assets" in Notion / Obsidian / chat with Claude. macOS
mounts SMB shares one-at-a-time under /Volumes/<sharename>, so this tool
accepts multiple --root arguments and combines them into a single tree file.

USAGE
-----
    # Single share
    python3 synology_tree.py --root "/Volumes/Common"

    # Multiple shares, one combined output (each share = its own section)
    python3 synology_tree.py \\
        --root "/Volumes/Common" \\
        --root "/Volumes/Accounting" \\
        --root "/Volumes/Field" \\
        --root "/Volumes/Multi Family" \\
        --root "/Volumes/Proinfo" \\
        --metadata --max-depth 4 \\
        --title "Proficient Concrete"

    # Folders only (skeleton)
    python3 synology_tree.py --root "/Volumes/Common" --no-files

    # Custom output path
    python3 synology_tree.py --root "/Volumes/Common" \\
        --out "./assets/synology_map.md"

OUTPUT
------
A single markdown file with:
  - Top-level header: title, list of roots, generated-at, total folder/file counts
  - One "## <share>" section per --root, each with its own stat line and tree
  - Tree using ├── └── │  characters
  - Files annotated with size (and optionally modified date with --metadata)
  - Synology / macOS noise filtered out by default
  - Aggregated error block at the bottom if any unreadable paths were hit

NOISE FILTERED
--------------
Folders: @eaDir, .Spotlight-V100, .Trashes, .fseventsd, .TemporaryItems, #recycle,
         .DocumentRevisions-V100, .com.apple.timemachine.donotpresent, $RECYCLE.BIN
Files:   .DS_Store, Thumbs.db, desktop.ini, .localized,
         ~$* (Office lockfiles), .~lock.* (LibreOffice lockfiles)
"""

from __future__ import annotations

import argparse
import os
import sys
from datetime import datetime
from pathlib import Path

# ---------------------------------------------------------------------------
# Filters
# ---------------------------------------------------------------------------

SKIP_DIRS = {
    "@eaDir",                                # Synology thumbnail/index dir
    ".Spotlight-V100",
    ".Trashes",
    ".fseventsd",
    ".TemporaryItems",
    "#recycle",                              # Synology recycle bin
    ".DocumentRevisions-V100",
    ".com.apple.timemachine.donotpresent",
    "$RECYCLE.BIN",
}

SKIP_FILES_EXACT = {
    ".DS_Store",
    "Thumbs.db",
    "desktop.ini",
    ".localized",
}


def is_skipped_file(name: str) -> bool:
    if name in SKIP_FILES_EXACT:
        return True
    # Office lockfiles: ~$Foo.xlsx
    if name.startswith("~$"):
        return True
    # LibreOffice lockfiles: .~lock.Foo.xlsx#
    if name.startswith(".~lock."):
        return True
    return False


# ---------------------------------------------------------------------------
# Formatting helpers
# ---------------------------------------------------------------------------

def fmt_size(n: int) -> str:
    """Human-readable size that handles ints and floats correctly."""
    size = float(n)
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if size < 1024 or unit == "TB":
            if unit == "B":
                return f"{int(size)} B"
            return f"{size:.1f} {unit}"
        size /= 1024
    return f"{size:.1f} PB"


def fmt_mtime(ts: float) -> str:
    return datetime.fromtimestamp(ts).strftime("%Y-%m-%d")


# ---------------------------------------------------------------------------
# Walker
# ---------------------------------------------------------------------------

class TreeStats:
    def __init__(self) -> None:
        self.folders = 0
        self.files = 0
        self.bytes = 0
        self.errors: list[str] = []


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


def walk(
    root: Path,
    *,
    include_files: bool,
    max_depth: int | None,
    metadata: bool,
    excludes: set[Path],
    stats: TreeStats,
) -> list[str]:
    """Return a list of pre-formatted tree lines (no leading newline)."""
    lines: list[str] = []

    def recurse(d: Path, depth: int, prefix: str) -> None:
        if max_depth is not None and depth > max_depth:
            return

        try:
            entries = list(os.scandir(d))
        except PermissionError:
            stats.errors.append(f"{d}: permission denied")
            lines.append(f"{prefix}└── _<permission denied>_")
            return
        except OSError as e:
            stats.errors.append(f"{d}: {e}")
            lines.append(f"{prefix}└── _<error: {e}>_")
            return

        # Split into dirs vs files, apply filters
        dirs: list[os.DirEntry] = []
        files: list[os.DirEntry] = []
        for entry in entries:
            try:
                if entry.is_dir(follow_symlinks=False):
                    if entry.name in SKIP_DIRS:
                        continue
                    if is_excluded(Path(entry.path).resolve(), excludes):
                        continue
                    dirs.append(entry)
                elif entry.is_file(follow_symlinks=False):
                    if not include_files:
                        continue
                    if is_skipped_file(entry.name):
                        continue
                    files.append(entry)
            except OSError:
                continue

        # Stable, case-insensitive ordering
        dirs.sort(key=lambda e: e.name.lower())
        files.sort(key=lambda e: e.name.lower())

        children: list[tuple[os.DirEntry, bool]] = (
            [(d, True) for d in dirs] + [(f, False) for f in files]
        )

        for i, (entry, is_dir) in enumerate(children):
            last = (i == len(children) - 1)
            connector = "└── " if last else "├── "
            child_prefix = prefix + ("    " if last else "│   ")

            if is_dir:
                stats.folders += 1
                # Trailing slash distinguishes folders. No markdown bold —
                # would render as literal '**' inside a fenced code block.
                lines.append(f"{prefix}{connector}{entry.name}/")
                recurse(Path(entry.path), depth + 1, child_prefix)
            else:
                stats.files += 1
                try:
                    st = entry.stat(follow_symlinks=False)
                    stats.bytes += st.st_size
                    if metadata:
                        suffix = f"  ({fmt_size(st.st_size)}, {fmt_mtime(st.st_mtime)})"
                    else:
                        suffix = f"  ({fmt_size(st.st_size)})"
                except OSError:
                    suffix = ""
                lines.append(f"{prefix}{connector}{entry.name}{suffix}")

    # Root line
    lines.append(f"{root}/")
    recurse(root, depth=1, prefix="")
    return lines


# ---------------------------------------------------------------------------
# Markdown assembly
# ---------------------------------------------------------------------------

def build_markdown(
    roots: list[Path],
    *,
    include_files: bool,
    max_depth: int | None,
    metadata: bool,
    excludes: set[Path],
    title: str | None,
) -> tuple[str, TreeStats]:
    aggregate = TreeStats()
    section_blocks: list[list[str]] = []

    for root in roots:
        section_stats = TreeStats()
        body_lines = walk(
            root,
            include_files=include_files,
            max_depth=max_depth,
            metadata=metadata,
            excludes=excludes,
            stats=section_stats,
        )
        aggregate.folders += section_stats.folders
        aggregate.files += section_stats.files
        aggregate.bytes += section_stats.bytes
        aggregate.errors.extend(section_stats.errors)

        if include_files:
            stat_line = (
                f"`{root}` — {section_stats.folders:,} folders, "
                f"{section_stats.files:,} files, "
                f"{fmt_size(section_stats.bytes)}"
            )
        else:
            stat_line = f"`{root}` — {section_stats.folders:,} folders"

        section_blocks.append([
            f"## {root.name}",
            "",
            stat_line,
            "",
            "```text",
            *body_lines,
            "```",
            "",
        ])

    # Top-level title
    if title is None:
        title = roots[0].name if len(roots) == 1 else "Combined"

    header_lines = [
        f"# Synology Tree — {title}",
        "",
        f"- **Roots:** {', '.join(r.name for r in roots)}",
        f"- **Generated:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"- **Folders:** {aggregate.folders:,}",
        (
            f"- **Files:** {aggregate.files:,}"
            if include_files
            else "- **Files:** _excluded_"
        ),
    ]
    if include_files:
        header_lines.append(f"- **Total size:** {fmt_size(aggregate.bytes)}")
    header_lines.append(
        f"- **Max depth:** {max_depth if max_depth is not None else 'unlimited'}"
    )
    header_lines.append("")

    error_lines: list[str] = []
    if aggregate.errors:
        error_lines.append("## Errors")
        error_lines.append("")
        for err in aggregate.errors[:50]:
            error_lines.append(f"- {err}")
        if len(aggregate.errors) > 50:
            error_lines.append(f"- _… and {len(aggregate.errors) - 50} more_")
        error_lines.append("")

    all_lines: list[str] = list(header_lines)
    for block in section_blocks:
        all_lines.extend(block)
    all_lines.extend(error_lines)

    return "\n".join(all_lines), aggregate


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description=(
            "Generate a markdown tree of one or more Synology shares "
            "(or any folder)."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p.add_argument(
        "--root",
        action="append",
        required=True,
        metavar="PATH",
        help=(
            "Path to a share or folder. Pass --root multiple times to combine "
            "shares into one tree (e.g. --root /Volumes/Common "
            "--root /Volumes/Accounting)."
        ),
    )
    p.add_argument(
        "--out",
        default=None,
        help=(
            "Output markdown file. Default: ~/Library/Logs/Proficient/synology/synology_tree.md "
            "(the project folder is AI-visible — generated trees with company file names "
            "must NOT land inside it)."
        ),
    )
    p.add_argument(
        "--title",
        default=None,
        help=(
            "Title for the combined tree header. Defaults to the share name "
            "when one --root is given, or 'Combined' when multiple."
        ),
    )
    p.add_argument(
        "--max-depth",
        type=int,
        default=None,
        help="Maximum recursion depth (default: unlimited). Root counts as depth 0.",
    )
    p.add_argument(
        "--no-files",
        action="store_true",
        help="Folders only, skip files",
    )
    p.add_argument(
        "--metadata",
        action="store_true",
        help="Include modified date next to each file (size is always shown)",
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
    # is AI-session-visible; Synology tree includes company file/folder names).
    out_path = (
        Path(args.out).expanduser().resolve()
        if args.out
        else Path.home() / "Library/Logs/Proficient/synology/synology_tree.md"
    )
    out_path.parent.mkdir(parents=True, exist_ok=True)

    md, stats = build_markdown(
        roots,
        include_files=not args.no_files,
        max_depth=args.max_depth,
        metadata=args.metadata,
        excludes=excludes,
        title=args.title,
    )

    out_path.write_text(md, encoding="utf-8")

    print(f"Wrote {out_path}")
    print(
        f"  shares: {len(roots)}  "
        f"folders: {stats.folders:,}  files: {stats.files:,}  "
        f"size: {fmt_size(stats.bytes)}  errors: {len(stats.errors)}"
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
