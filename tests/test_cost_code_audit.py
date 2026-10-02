"""shared/cost_code_audit - the vendor cost-code family audit (pure, offline)."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from shared.cost_code_audit import (  # noqa: E402
    aggregate_bills, aggregate_memo, code_families, flag_lines, is_nonjob)


def _row(bill_id, desc, code, **kw):
    number, name = code_families(code)
    return {"vendor": "BODIN CONCRETE L.P", "bill_id": bill_id, "bill_doc": bill_id,
            "cost_code": code, "number": number, "cost_name": name, "desc": desc,
            "account": "", "amount": 10.0, "date": "2026-08-24", **kw}


CONCRETE_VENDOR = {"BODIN CONCRETE L.P": "concrete"}


def test_memo_helpers():
    assert aggregate_memo("20yds of exposed pea gravel")
    assert not aggregate_memo("ENVIRONMENTAL")
    assert is_nonjob("FUEL SURCHARGE")


def test_pea_gravel_bill_is_read_as_a_whole():
    # Bodin 235198 (the user 2026-09-14): the product line vouches for the
    # add-on lines on the SAME bill - none of the four FW4 lines is a miscode.
    rows = [_row("B1", "20yds of exposed pea gravel", "FW4"),
            _row("B1", "FUEL SURCHARGE", "FW4"),
            _row("B1", "ENVIRONMENTAL", "FW4"),
            _row("B1", "TAXES", "FW4")]
    assert aggregate_bills(rows) == {"B1"}
    assert flag_lines(rows, CONCRETE_VENDOR) == []


def test_bill_memo_alone_vouches_for_aggregate():
    rows = [_row("B2", "", "FW4", bill_memo="RP7433-FTW pea gravel delivery"),
            _row("B2", "TAXES", "FW4", bill_memo="RP7433-FTW pea gravel delivery")]
    assert flag_lines(rows, CONCRETE_VENDOR) == []


def test_aggregate_code_without_an_aggregate_memo_still_flags():
    # A *4 on a concrete vendor with nothing on the bill reading as aggregate
    # is still the miscode it always was - the gate is memo-driven, not blanket.
    rows = [_row("B3", "10 CY 3000 psi", "FW4"),
            _row("B3", "TAXES", "FW4")]
    flagged = flag_lines(rows, CONCRETE_VENDOR)
    assert [r["desc"] for r in flagged] == ["10 CY 3000 psi", "TAXES"]
    assert all("expected *1 Concrete" in r["reason"] for r in flagged)


def test_sibling_bills_do_not_vouch_for_each_other():
    rows = [_row("B4", "pea gravel", "FW4"),
            _row("B5", "TAXES", "FW4")]
    flagged = flag_lines(rows, CONCRETE_VENDOR)
    assert [r["bill_id"] for r in flagged] == ["B5"]


def test_concrete_yardage_line_is_never_vouched_by_a_pea_gravel_sibling():
    # Bodin 235212: pea gravel + "10+yds of 3000psi regular rock" on one bill.
    # The add-on lines pass on the bill's memo; the concrete line still flags.
    rows = [_row("B6", "10+yds of exposed pea gravel", "FW4"),
            _row("B6", "10+yds of 3000psi regular rock", "FW4"),
            _row("B6", "ENVIRONMENTAL FEE", "FW4"),
            _row("B6", "TAXES", "FW4")]
    flagged = flag_lines(rows, CONCRETE_VENDOR)
    assert [r["desc"] for r in flagged] == ["10+yds of 3000psi regular rock"]
