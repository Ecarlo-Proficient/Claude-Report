"""A source workbook that EXISTS is either read or the run stops on it (shared/xlsx_read).

Never again "the WIP master was locked, so every job had no contract", or "the old P&L would not open, so the
PM's typed Contract / ETC and green confirmations were saved over". Missing files still read as empty.
"""
import sys
from pathlib import Path

import pytest
from openpyxl import Workbook

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "project-pnl"))
sys.path.insert(0, str(ROOT / "wip"))

from shared.xlsx_read import SourceUnreadable, open_source  # noqa: E402
import project_pnl_export as P  # noqa: E402
import master_wip_test as M  # noqa: E402


def _broken(path):
    path.write_bytes(b"PK\x03\x04 cut off mid-copy")
    return path


def test_open_source_tells_missing_from_unreadable(tmp_path):
    assert open_source(tmp_path / "nope.xlsx") is None
    with pytest.raises(SourceUnreadable, match="bad.xlsx exists but could not be opened"):
        open_source(_broken(tmp_path / "bad.xlsx"))
    wb = Workbook()
    wb.save(tmp_path / "ok.xlsx")
    assert open_source(tmp_path / "ok.xlsx") is not None


@pytest.mark.parametrize("reader", [
    P.load_wip_master,
    P.read_back_inputs,
    P.read_back_ledger_marks,
    lambda p: P._read_cost_code_sheet(p, "cost code"),
    P._read_yards_sheet,
    M.read_mfd_etc_from_division_tab,
])
def test_every_reader_stops_on_an_unreadable_file(tmp_path, reader):
    with pytest.raises(SourceUnreadable):
        reader(_broken(tmp_path / "Source.xlsx"))


def test_missing_files_still_read_as_empty(tmp_path):
    gone = tmp_path / "missing.xlsx"
    assert P.load_wip_master(gone) == {}
    assert P.read_back_inputs(gone) == {}
    assert P.read_back_ledger_marks(gone) == {}
    assert M.read_mfd_etc_from_division_tab(gone) == {}


def test_an_unreadable_wip_report_never_falls_through_to_an_older_one(tmp_path):
    from shared import wip_contracts as W
    _broken(tmp_path / "newest.xlsx")
    with pytest.raises(SourceUnreadable):
        W._scan(tmp_path / "newest.xlsx")
    assert W._scan(tmp_path / "not-there.xlsx") == {}
