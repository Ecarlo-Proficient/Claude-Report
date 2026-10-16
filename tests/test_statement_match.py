"""statement-reconciler: pairing a statement's invoice # with the QBO bill's Ref #.

Found 10/09/2026 re-checking every vendor: two bills that were entered AND paid showed
as "not entered" - QBO '16018k' vs the statement's '16018K' (CMC), and the clerk's
'401417-CC FEE' for the statement's 401417 (Cowtown).
"""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "statement-reconciler"))

import statement_reconciler as sr  # noqa: E402


def _bill(doc, bal, bid="1", date="2026-09-01"):
    return sr.QboBill(bill_id=bid, doc_number=doc, txn_date=date,
                      open_balance=bal, total_amount=bal)


def _run(lines, bills, paid=()):
    return [row for _, _, row in sr.reconcile_iter(lines, bills, list(paid), "2026-10-01")]


def test_case_and_spacing_never_decide_a_match():
    rows = _run([sr.StmtLine("2026-08-31", "16018K", 73.50)], [],
                paid=[_bill("16018k ", 0.0)])
    assert [r.category for r in rows] == ["LIKELY_VENDOR_LAG"]


def test_a_clerk_suffix_pairs_when_only_one_bill_carries_it():
    rows = _run([sr.StmtLine("2026-09-21", "401417", 3210.00)], [],
                paid=[_bill("401417-CC FEE", 0.0)])
    assert [r.category for r in rows] == ["LIKELY_VENDOR_LAG"]


def test_an_open_suffixed_bill_is_matched_and_not_also_missing_from_the_statement():
    rows = _run([sr.StmtLine("2026-09-21", "401417", 3210.00)],
                [_bill("401417-CC FEE", 3210.00)])
    assert [r.category for r in rows] == ["MATCHED"]


def test_a_suffix_never_guesses_between_two_bills_or_on_a_short_number():
    two = _run([sr.StmtLine("2026-09-21", "401417", 50.0)], [],
               paid=[_bill("401417-A", 0.0, "1"), _bill("401417-B", 0.0, "2")])
    short = _run([sr.StmtLine("2026-09-21", "142", 50.0)], [], paid=[_bill("142-X", 0.0)])
    assert two[0].category == "MISSING_IN_QBO" and short[0].category == "MISSING_IN_QBO"


def test_a_suffixed_ref_that_is_its_own_statement_line_is_not_borrowed():
    rows = _run([sr.StmtLine("2026-09-21", "401417", 100.0),
                 sr.StmtLine("2026-09-21", "401417-CC FEE", 3210.0)], [],
                paid=[_bill("401417-CC FEE", 0.0)])
    assert [r.category for r in rows] == ["MISSING_IN_QBO", "LIKELY_VENDOR_LAG"]
