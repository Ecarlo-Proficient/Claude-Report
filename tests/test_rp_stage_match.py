"""RP bills on a job billed in stages go to the FIRST stage invoice after the bill
that covers them, not the biggest invoice in the 60-day window (the owner 10/14/2026)."""
import datetime as dt
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "bill-tracker"))

import qbo_bill_tracker as q  # noqa: E402

CUST = "7597"


def _inv(doc, date, amt, memo):
    return {"DocNumber": doc, "TxnDate": date, "TotalAmt": amt, "PrivateNote": memo}


PIERS = _inv("34620", "2026-09-01", 32045.00, "RP7597 - 7204 Meadow Lake Avenue - Piers")
FOUND = _inv("34683", "2026-09-17", 59658.85, "RP7597 - 7204 Meadow Lake Avenue - Foundation")


def test_staged_job_bill_goes_to_first_stage_invoice_after_it():
    # K783197 08/19 4,947.46: Piers (09/01) is the first stage invoice after the bill
    # and covers it - it wins over the bigger Foundation invoice (09/17).
    m, basis = q.find_matching_invoice_ex(
        dt.date(2026, 8, 19), "RP", CUST, {CUST: [FOUND, PIERS]}, bill_amount=4947.46)
    assert m is PIERS and basis == q.MATCH_BASIS_STAGE


def test_staged_job_bill_after_piers_goes_to_foundation():
    m, basis = q.find_matching_invoice_ex(
        dt.date(2026, 9, 3), "RP", CUST, {CUST: [FOUND, PIERS]}, bill_amount=800)
    assert m is FOUND and basis == q.MATCH_BASIS_STAGE


def test_staged_job_first_stage_too_small_skips_to_next_stage():
    # A 40,000 bill: Piers (32,045) does not cover it, Foundation does.
    m, basis = q.find_matching_invoice_ex(
        dt.date(2026, 8, 19), "RP", CUST, {CUST: [FOUND, PIERS]}, bill_amount=40000)
    assert m is FOUND and basis == q.MATCH_BASIS_STAGE


def test_staged_job_nothing_covers_falls_through():
    # 90,000 bill: no stage invoice covers it -> the old rules (cover, then GL draw).
    m, basis = q.find_matching_invoice_ex(
        dt.date(2026, 8, 19), "RP", CUST, {CUST: [FOUND, PIERS]}, bill_amount=90000)
    assert m is None and basis == ""


def test_one_invoice_job_keeps_largest_cover_rule():
    # Foundation + a pump extra: not billed in stages, so the old rule holds and
    # the biggest covering invoice wins even though the extra comes first.
    extra = _inv("34500", "2026-08-25", 1200.00, "RP7000 - 1 Elm St - Pump charges")
    slab = _inv("34510", "2026-09-10", 25000.00, "RP7000 - 1 Elm St - Foundation")
    m, basis = q.find_matching_invoice_ex(
        dt.date(2026, 8, 19), "RP", CUST, {CUST: [extra, slab]}, bill_amount=500)
    assert m is slab and basis == q.MATCH_BASIS_COVER


def test_rp_billed_in_stages():
    assert q.rp_billed_in_stages([PIERS, FOUND])
    assert q.rp_billed_in_stages([PIERS])                      # one partial stage so far
    assert not q.rp_billed_in_stages([FOUND])
    assert not q.rp_billed_in_stages([])
