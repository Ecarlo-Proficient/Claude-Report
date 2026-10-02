"""Pure-logic tests for ledger/bill_payment_stub - the payment-on-top / bills-below stub
(no mirror, no Chrome: synthetic BillPayment + Bill records, the column registry, the page)."""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "ledger"))

import bill_payment_stub as bps   # noqa: E402


def test_selftest_passes():
    assert bps._selftest() == 0


def test_money_and_dates():
    assert bps.money(1234.5) == "1,234.50"
    assert bps.money(-3) == "(3.00)"
    assert bps.mdy("2026-09-21") == "09/21/2026"
    assert bps.mdy(None) == ""


def test_columns_registry_toggles():
    assert bps.resolve_columns(drop=["type", "open_balance"]) == ["date", "number", "memo", "amount"]
    assert bps.resolve_columns(add=["project"])[-1] == "project"
    assert set(bps.DEFAULT_COLUMNS) <= set(bps.COLUMNS)
