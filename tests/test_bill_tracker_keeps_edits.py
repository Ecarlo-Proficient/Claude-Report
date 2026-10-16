"""bill-tracker: the clerks' typed Lien / Notes / audit marks can never be wiped by a run.

A tracker that cannot be opened stops the run (it is never read as "no edits"), and the backups only ever
copy a tracker that just opened cleanly: the day's first copy is kept, plus the copy from just before the run.
"""
import datetime as dt
import sys
from pathlib import Path

import pytest
from openpyxl import Workbook

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "bill-tracker"))

import excel_bill_sync as ebs  # noqa: E402


def _tracker(path, note="call GC"):
    wb = Workbook()
    ws = wb.active
    ws.title = "Bills"
    ws.append(["_Key", "Lien", "Notes"])
    ws.append(["101", "", note])
    audit = wb.create_sheet("Audit - Coding")
    audit.append(["_Key", "Status"])
    audit.append(["101", "fixed in QBO"])
    wb.save(path)


def test_no_tracker_yet_is_a_first_run(tmp_path):
    wb = ebs.read_existing(tmp_path / "Bill Tracker.xlsx")
    assert wb is None
    assert ebs.preserve_edits(wb) == {} and ebs.preserve_audit_marks(wb) == {}


def test_an_unreadable_tracker_stops_instead_of_reading_as_empty(tmp_path):
    path = tmp_path / "Bill Tracker.xlsx"
    path.write_bytes(b"PK\x03\x04 half a file")              # what a cut-off copy looks like
    with pytest.raises(ebs.SourceUnreadable):
        ebs.read_existing(path)


def test_edits_and_marks_come_from_one_open(tmp_path):
    path = tmp_path / "Bill Tracker.xlsx"
    _tracker(path)
    wb = ebs.read_existing(path)
    assert ebs.preserve_edits(wb) == {"101": {"Lien": "", "Notes": "call GC"}}
    assert ebs.preserve_audit_marks(wb) == {"101": "fixed in QBO"}


def test_backup_keeps_the_days_first_copy_and_the_previous_one(tmp_path):
    path = tmp_path / "Bill Tracker.xlsx"
    backups = tmp_path / "_backups"
    _tracker(path, note="morning")
    ebs.rotate_backup(path)
    _tracker(path, note="afternoon")
    ebs.rotate_backup(path)
    day = backups / f"Bill Tracker \u2014 {dt.date.today().isoformat()}.xlsx"
    prev = backups / "Bill Tracker - previous.xlsx"
    assert ebs.preserve_edits(ebs.read_existing(day))["101"]["Notes"] == "morning"
    assert ebs.preserve_edits(ebs.read_existing(prev))["101"]["Notes"] == "afternoon"
