"""bill-tracker: the tracker is never written while someone has it open in Excel.

Excel leaves a hidden `~$<name>` lock beside an open workbook. The run skips (exit 0, the next run catches up);
a lock older than STALE_LOCK_HOURS still skips but exits EXIT_STALE_LOCK so the office server alerts.
"""
import os
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "bill-tracker"))

import excel_bill_sync as ebs  # noqa: E402


def test_no_lock_means_write(tmp_path):
    book = tmp_path / "Bill Tracker.xlsx"
    book.write_bytes(b"x")
    assert ebs.open_in_excel(book) is None
    assert ebs._skip_if_open(book, "this run") is None


def test_excel_lock_skips_cleanly(tmp_path):
    book = tmp_path / "Bill Tracker.xlsx"
    (tmp_path / "~$Bill Tracker.xlsx").write_bytes(b"")
    assert ebs.open_in_excel(book).name == "~$Bill Tracker.xlsx"
    assert ebs._skip_if_open(book, "this run") == 0


def test_short_lock_name_is_caught(tmp_path):
    book = tmp_path / "Bill Tracker.xlsx"
    (tmp_path / "~$ll Tracker.xlsx").write_bytes(b"")
    assert ebs._skip_if_open(book, "this run") == 0


def test_stale_lock_still_skips_but_alerts(tmp_path):
    book = tmp_path / "Bill Tracker.xlsx"
    lock = tmp_path / "~$Bill Tracker.xlsx"
    lock.write_bytes(b"")
    old = time.time() - (ebs.STALE_LOCK_HOURS + 1) * 3600
    os.utime(lock, (old, old))
    assert ebs._skip_if_open(book, "this run") == ebs.EXIT_STALE_LOCK != 0


def test_other_files_lock_is_ignored(tmp_path):
    book = tmp_path / "Bill Tracker.xlsx"
    (tmp_path / "~$Invoice Tracker.xlsx").write_bytes(b"")
    assert ebs._skip_if_open(book, "this run") is None
