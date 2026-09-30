"""The CP Overview's home = the Active Awarded Projects folder on Common (owner 2026-09-30)."""
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from shared import pnl_paths  # noqa: E402


def test_cp_overview_lives_in_the_awarded_folder(monkeypatch, tmp_path):
    monkeypatch.setattr(pnl_paths, "CP_AWARDED_BASE", tmp_path)
    assert pnl_paths.overview_dir("CP") == tmp_path


def test_cp_overview_stops_when_common_is_not_mounted(monkeypatch, tmp_path):
    monkeypatch.setattr(pnl_paths, "CP_AWARDED_BASE", tmp_path / "not-mounted")
    with pytest.raises(pnl_paths.HomeNotMounted):
        pnl_paths.overview_dir("CP")
