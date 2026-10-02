"""bill-tracker audit: a line coded to one job while its memo / line names exactly
one OTHER job is a 'Wrong Job?' finding - subs included (owner 10/02/2026)."""
import importlib.util
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "bill-tracker"))
_spec = importlib.util.spec_from_file_location("ebs", ROOT / "bill-tracker" / "excel_bill_sync.py")
ebs = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(ebs)


def test_memo_names_other_job():
    r = {"project_num": "CP961", "bill_memo": "Sub Service: 07/17/2026 - 07/23/2026 CP745 - 8811 OLD DECATUR",
         "line_desc": "Pour slab", "is_sub": True}
    assert ebs._wrong_job(r) == "memo says CP745, coded to CP961 - confirm which job"


def test_same_job_with_address_suffix_is_fine():
    r = {"project_num": "CP790", "bill_memo": "CP790-A NORTH SITE - GC NAME", "line_desc": ""}
    assert ebs._wrong_job(r) is None


def test_multi_job_memo_is_not_a_finding():
    r = {"project_num": "RP7111", "bill_memo": "CP672, RP7111, RP7093 - pumping", "line_desc": ""}
    assert ebs._wrong_job(r) is None


def test_ftw_twin_is_not_a_wrong_job():
    # RP bills name the house in the memo and code flatwork to the -FTW twin
    r = {"project_num": "RP7242-FTW", "bill_memo": "RP7242 - 123 MAIN ST", "line_desc": ""}
    assert ebs._wrong_job(r) is None


def test_a_different_job_still_flags():
    r = {"project_num": "CP800", "bill_memo": "RP7566 - pumping", "line_desc": ""}
    assert ebs._wrong_job(r) == "memo says RP7566, coded to CP800 - confirm which job"


def test_no_project_is_left_to_missing_project():
    assert ebs._wrong_job({"project_num": "", "bill_memo": "CP790 - x"}) is None
