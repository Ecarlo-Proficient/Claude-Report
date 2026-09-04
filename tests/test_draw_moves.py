"""Pure-logic tests for shared/draw_moves - the 'push a bill into a later draw' rule
(no QBO, no ledger, no rules file: the rules are passed in)."""
import datetime as dt
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))                    # repo root -> shared

from shared import draw_moves as dm

RAW = {"project": "CP800", "vendor": "Preferred Materials", "after": "2026-07-20",
       "through": "2026-07-31", "move_to": "2026-08-01", "from_draw": 3, "to_draw": 4,
       "why": "agreed with the supplier"}
RULES = [dm._norm_rule(RAW)]


def test_rule_normalises():
    r = RULES[0]
    assert r["project"] == "CP800" and r["vendor"] == "preferred materials"
    assert r["after"] == dt.date(2026, 7, 20) and r["move_to"] == dt.date(2026, 8, 1)


def test_broken_rules_are_dropped():
    assert dm._norm_rule({**RAW, "vendor": ""}) is None
    assert dm._norm_rule({**RAW, "through": "2026-07-10"}) is None      # through <= after
    assert dm._norm_rule({**RAW, "move_to": "2026-07-25"}) is None      # move_to inside the old window
    assert dm._norm_rule({**RAW, "after": "not a date"}) is None


def test_find_move_window_is_after_cutoff_through_end():
    f = lambda d: dm.find_move("CP800", "Preferred Materials LLC", d, RULES)   # noqa: E731
    assert f("2026-07-20") is None          # the cutoff itself stays
    assert f("2026-07-21") is not None      # first day after the cutoff moves
    assert f("2026-07-31") is not None      # ... through the period end
    assert f("2026-08-01") is None          # already in the next draw - nothing to move
    assert f("2026-07-16") is None


def test_find_move_is_exact_project_and_vendor_substring():
    assert dm.find_move("CP800-FTW", "Preferred Materials LLC", "2026-07-22", RULES) is None   # -FTW is its own job
    assert dm.find_move("cp800", "PREFERRED MATERIALS LLC", "2026-07-22", RULES) is not None   # case-insensitive
    assert dm.find_move("CP800", "Cowtown Redi-Mix", "2026-07-22", RULES) is None
    assert dm.find_move("CP800", None, "2026-07-22", RULES) is None
    assert dm.find_move(None, "Preferred Materials", "2026-07-22", RULES) is None


def test_effective_date_keeps_the_callers_type():
    d, r = dm.effective_date("CP800", "Preferred Materials LLC", dt.date(2026, 7, 22), RULES)
    assert d == dt.date(2026, 8, 1) and r is RULES[0]
    s, r = dm.effective_date("CP800", "Preferred Materials LLC", "2026-07-22", RULES)
    assert s == "2026-08-01" and r is RULES[0]
    same, r = dm.effective_date("CP800", "Preferred Materials LLC", "2026-07-16", RULES)
    assert same == "2026-07-16" and r is None


def test_labels():
    assert dm.push_label(RULES[0]) == "pushed from Draw #3"
    note_in = dm.push_note(RULES[0])
    note_out = dm.push_note(RULES[0], "out")
    assert "after 07/20/26" in note_in and "from Draw #3" in note_in and "agreed with the supplier" in note_in
    assert "to Draw #4" in note_out


def test_no_rules_file_is_silent(tmp_path):
    assert dm.load_rules(tmp_path / "missing.json") == []
    (tmp_path / "bad.json").write_text("{not json")
    assert dm.load_rules(tmp_path / "bad.json") == []
    (tmp_path / "ok.json").write_text('{"rules": [%s]}' % __import__("json").dumps(RAW))
    assert len(dm.load_rules(tmp_path / "ok.json")) == 1
