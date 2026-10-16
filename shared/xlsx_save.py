"""shared/xlsx_save.py - THE way a tool replaces a workbook that people open and type in.

`save_verified(wb, path)` never leaves a broken or half-written workbook at `path`:
  1. the workbook is built in a hidden temp file beside it (same folder, same disk),
  2. the tool's own fix-ups run on the temp file (`post_process`),
  3. the tool's own checks (`checks`) and the repo-wide Excel check (shared/xlsx_verify, rule 5b) run on it,
  4. only a temp file that passed everything is written over the real one.
A failed check raises NotSaved and the real file is left exactly as it was.

Step 4 writes the verified bytes INTO the existing file instead of renaming the temp over it. A rename makes
a NEW file, and on the office server a new file carries the automation user's permissions, not the share's:
nobody could open the Bill Tracker (docker/STATUS 10/05). Writing into the existing file keeps its owner and
permissions. A first run (no file yet) renames, so a file never exists half-made under its real name.

The copy is the only moment the real file is being written, and it is a straight local copy of a file that is
already complete. The callers pair it with a read guard: a workbook that cannot be opened stops the next run
instead of being treated as empty (bill-tracker `read_existing`), and the backup holds the last good copy.
"""
from __future__ import annotations

import os
import shutil
from pathlib import Path
from typing import Callable, Iterable, Optional

try:
    from shared import xlsx_verify
except ImportError:  # imported with shared/ itself on sys.path
    import xlsx_verify  # type: ignore


class NotSaved(Exception):
    """The new workbook failed a check; the file on disk was not touched."""


def _temp_for(path: Path) -> Path:
    # Hidden and never `~$...` (that prefix is Excel's own lock file, which the open-guards look for).
    return path.with_name(f".{path.stem}.building{path.suffix}")


def _write_into(src: Path, dst: Path) -> None:
    """Overwrite dst's contents with src's, keeping dst itself (owner, permissions, ACL)."""
    with open(src, "rb") as fin, open(dst, "r+b") as fout:
        shutil.copyfileobj(fin, fout, 1024 * 1024)
        fout.truncate()
        fout.flush()
        os.fsync(fout.fileno())


def save_verified(wb, path: Path,
                  post_process: Optional[Callable[[Path], None]] = None,
                  checks: Optional[Callable[[Path], Iterable[str]]] = None) -> None:
    """Build `wb` in a temp file, verify it, then put it at `path`. Raises NotSaved on a failed check."""
    path = Path(path)
    tmp = _temp_for(path)
    try:
        wb.save(str(tmp))
        if post_process:
            post_process(tmp)
        failures = list(checks(tmp)) if checks else []
        try:
            xlsx_verify.assert_clean(tmp)
        except ValueError as e:
            failures.append(str(e))
        if failures:
            raise NotSaved("; ".join(failures))
        if path.exists():
            _write_into(tmp, path)
        else:
            os.replace(tmp, path)
    finally:
        tmp.unlink(missing_ok=True)
