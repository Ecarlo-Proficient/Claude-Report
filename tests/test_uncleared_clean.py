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
