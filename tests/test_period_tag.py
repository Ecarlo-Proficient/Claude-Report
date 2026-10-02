"""shared/draws.parse_period_tag - THE draw-period tag parser (parentheses optional)."""
import datetime as dt
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from shared.draws import PERIOD_TAG_RE, parse_period_tag  # noqa: E402

MAY21, JUN20 = dt.date(2026, 5, 21), dt.date(2026, 6, 20)


def test_parenthesised_tag():
    assert parse_period_tag("MFD281 - Draw 4 (Period: 05/21/2026 - 06/20/2026)") == (MAY21, JUN20)


def test_bare_tag_without_parentheses():
    # CP785 draws as typed in QBO (the user 2026-09-15): no parens, no space after the colon
    m = "CP785 - 7 BREW COFFEE - 1001 E US HWY 377 - Draw #1 -\nPeriod:05/21/2026 - 06/20/2026"
    assert parse_period_tag(m) == (MAY21, JUN20)


def test_two_digit_years_and_en_dash():
    assert parse_period_tag("(Period: 5/21/26 – 6/20/26)") == (MAY21, JUN20)


def test_no_tag_or_bad_tag():
    assert parse_period_tag("Draw #2 - retainage not billed") is None
    assert parse_period_tag("Period: 06/20/2026 - 05/21/2026") is None   # start > end
    assert parse_period_tag("") is None


def test_regex_strips_either_form():
    assert PERIOD_TAG_RE.sub("", "X (Period: 05/21/2026 - 06/20/2026) Y").split() == ["X", "Y"]
    assert PERIOD_TAG_RE.sub("", "X - Period:05/21/2026 - 06/20/2026").strip(" -") == "X"
