"""shared/pnl_paths._find_awarded_cp_folder - active awarded jobs first, then the
Completed Projects archive (by year), full match only in the archive."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from shared.pnl_paths import _find_awarded_cp_folder  # noqa: E402


def _tree(tmp_path, *dirs):
    for d in dirs:
        (tmp_path / d).mkdir(parents=True)
    return tmp_path


def test_active_folder_wins(tmp_path):
    base = _tree(tmp_path, "CP672 - ACTIVE JOB", "Completed Projects/2025/CP672 - OLD COPY")
    assert _find_awarded_cp_folder(base, "CP672").name == "CP672 - ACTIVE JOB"


def test_completed_year_folder_found(tmp_path):
    base = _tree(tmp_path, "CP997 - OTHER", "Completed Projects/2025/CP610 - SAUCE + VINE  Complete")
    assert _find_awarded_cp_folder(base, "CP610").name == "CP610 - SAUCE + VINE  Complete"


def test_completed_top_level_found(tmp_path):
    base = _tree(tmp_path, "Completed Projects/CP568 - Kemp ISD Renovations")
    assert _find_awarded_cp_folder(base, "CP568").name == "CP568 - Kemp ISD Renovations"


def test_archive_needs_full_match(tmp_path):
    base = _tree(tmp_path, "Completed Projects/2025/CP5921 - LOOKALIKE", "Completed Projects/2025/610 BARE")
    assert _find_awarded_cp_folder(base, "CP592") is None
    assert _find_awarded_cp_folder(base, "CP610") is None


def test_missing_base(tmp_path):
    assert _find_awarded_cp_folder(tmp_path / "nope", "CP610") is None
