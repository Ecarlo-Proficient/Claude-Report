"""shared/qbo_api.pick_customer - one customer per project # when QBO holds duplicates."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from shared.qbo_api import pick_customer  # noqa: E402


def _c(i, name, created):
    return {"Id": i, "DisplayName": name, "MetaData": {"CreateTime": created}}


def test_exact_name_beats_a_typo_twin():
    a, b = _c("1", "RP7340", "2026-01-28T09:02:55"), _c("2", "RP7340 -FTW", "2026-01-28T09:03:20")
    assert pick_customer("RP7340", [b, a], {})["Id"] == "1"


def test_invoices_beat_an_empty_duplicate():
    a, b = _c("29851", "RP7074", "2025-10-14"), _c("29251", "RP7074", "2025-08-21")
    assert pick_customer("RP7074", [a, b], {"29251": 1})["Id"] == "29251"


def test_oldest_record_when_nothing_else_separates_them():
    a, b = _c("30312", "RP7340", "2026-01-28T09:02:55"), _c("30313", "RP7340", "2026-01-28T09:03:20")
    assert pick_customer("RP7340", [b, a], {})["Id"] == "30312"
