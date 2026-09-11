"""Pure-logic tests for the RP draw-matching semantics (no QBO, no GL file)."""
import datetime as dt
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))                    # repo root → shared
sys.path.insert(0, str(ROOT / "bill-tracker"))   # the module under test

import qbo_bill_tracker as q

CUST = "42"


def _inv(date, amt, memo=""):
    return {"TxnDate": date, "TotalAmt": amt, "PrivateNote": memo}


def _pool(*invs):
    return {CUST: list(invs)}


# ── rp_gl_state ──────────────────────────────────────────────────────────
def test_rp_gl_state():
    p = _pool(_inv("2026-07-01", 3000))
    assert q.rp_gl_state("RP1", CUST, p, None) is None            # no GL → degrade
    assert q.rp_gl_state(None, CUST, p, {"RP1": 100}) is None     # no project #
    assert q.rp_gl_state("RP1", CUST, p, {"RP1": 0}) is None      # no/zero contract
    assert q.rp_gl_state("RP1", CUST, {CUST: []}, {"RP1": 1e5}) == "unbilled"
    assert q.rp_gl_state("RP1", CUST, p, {"RP1": 1e5}) == "partial"   # 3000 < 100000
    assert q.rp_gl_state("RP1", CUST, p, {"RP1": 3000}) == "full"     # 3000 ≈ contract
    assert q.rp_gl_state("rp1", CUST, p, {"RP1": 1e5}) == "partial"   # key case-insensitive


# ── find_matching_invoice_ex, RP branch ──────────────────────────────────
def test_cover_still_wins():
    # A covering invoice inside the 60-day window matches first, basis=cover,
    # even when the GL says the job is only partially billed.
    cov = _inv("2026-07-10", 1000)
    m, basis = q.find_matching_invoice_ex(
        dt.date(2026, 7, 2), "RP", CUST, _pool(cov),
        bill_amount=500, project_num="RP1234", gl_contracts={"RP1234": 1e5})
    assert m is cov and basis == q.MATCH_BASIS_COVER


def test_partial_matches_next_draw():
    # No invoice covers the 5000 cost; the 7/1 draw predates it → the NEXT draw
    # (earliest dated on/after 7/2) authorizes it, no forward cap.
    before, nxt, later = _inv("2026-07-01", 2000), _inv("2026-07-20", 3000), _inv("2026-08-15", 4000)
    m, basis = q.find_matching_invoice_ex(
        dt.date(2026, 7, 2), "RP", CUST, _pool(before, nxt, later),
        bill_amount=5000, project_num="RP1234", gl_contracts={"RP1234": 1e5})
    assert m is nxt and basis == q.MATCH_BASIS_DRAW


def test_partial_no_next_draw_stays_unmatched():
    past = _inv("2026-06-01", 3000)
    m, basis = q.find_matching_invoice_ex(
        dt.date(2026, 7, 2), "RP", CUST, _pool(past),
        bill_amount=5000, project_num="RP1234", gl_contracts={"RP1234": 1e5})
    assert m is None and basis == ""


def test_full_attaches_to_last_draw():
    # 100%-billed job, late bill dated after every invoice → matches the last draw.
    d1, d2 = _inv("2026-05-01", 40000), _inv("2026-06-01", 60000)
    m, basis = q.find_matching_invoice_ex(
        dt.date(2026, 7, 2), "RP", CUST, _pool(d1, d2),
        bill_amount=500, project_num="RP1234", gl_contracts={"RP1234": 1e5})
    assert m is d2 and basis == q.MATCH_BASIS_FINAL


def test_gl_unavailable_degrades_to_cover_only():
    # gl_contracts=None (share unmounted): a non-covering invoice → no match,
    # exactly today's behavior.
    nxt = _inv("2026-07-20", 3000)
    m, basis = q.find_matching_invoice_ex(
        dt.date(2026, 7, 2), "RP", CUST, _pool(nxt),
        bill_amount=5000, project_num="RP1234", gl_contracts=None)
    assert m is None and basis == ""
