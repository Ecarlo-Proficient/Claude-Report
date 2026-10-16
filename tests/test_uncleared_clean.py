"""ledger/load_uncleared_checks.clean - voided and 'To print' checks are left off the list."""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "ledger"))

import load_uncleared_checks as U  # noqa: E402


def test_clean():
    recs = {("BillPayment", "1"): {"TotalAmt": 100, "CheckPayment": {"PrintStatus": "NeedToPrint"}},
            ("BillPayment", "2"): {"TotalAmt": 0, "CheckPayment": {"PrintStatus": "NotSet"}},
            ("Purchase", "3"): {"TotalAmt": 50, "PrintStatus": "PrintComplete"},
            ("Purchase", "4"): {"TotalAmt": 75, "PrintStatus": "NeedToPrint"}}
    checks = [{"_id": "1", "Transaction Type": "Bill Payment (Check)", "Amount": "-100"},
              {"_id": "2", "Transaction Type": "Bill Payment (Check)", "Amount": "-40"},
              {"_id": "3", "Transaction Type": "Check", "Amount": "-50"},
              {"_id": "4", "Transaction Type": "Check", "Amount": "-75"},
              {"_id": "5", "Transaction Type": "Check", "Amount": "0"},
              {"_id": "6", "Transaction Type": "Check", "Amount": "-9"}]
    kept, voided, to_print = U.clean(checks, lambda e, i: recs.get((e, i)))
    assert [c["_id"] for c in kept] == ["3", "6"]
    assert [c["_id"] for c in voided] == ["2", "5"]
    assert [c["_id"] for c in to_print] == ["1", "4"]


def test_cleared_delta():
    """A check the last pull listed that this pull no longer does -> a check_cleared row with its window."""
    previous = [{"qbo_txn_id": "10", "txn_type": "Bill Payment (Check)", "txn_date": "2026-09-01", "check_no": "48401",
                 "payee": "V", "account": "Frost", "amount": 500.0, "loaded_at": "2026-10-10T08:00:00"},
                {"qbo_txn_id": "11", "txn_type": "Check", "txn_date": "2026-09-02", "check_no": "48402",
                 "payee": "W", "account": "Frost", "amount": 60.0, "loaded_at": "2026-10-10T08:00:00"}]
    kept = [{"_id": "11", "Transaction Type": "Check", "Amount": "-60"}, {"_id": "12", "Transaction Type": "Check", "Amount": "-9"}]
    rows = U.cleared_delta(previous, kept, "2026-10-15T09:00:00")
    assert rows == [("10", "Bill Payment (Check)", "2026-09-01", "48401", "V", "Frost", 500.0, "2026-10-10T08:00:00", "2026-10-15T09:00:00")]
    assert U.cleared_delta([], kept, "x") == []
