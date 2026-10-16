"""Pay bills reads QuickBooks' open balance, not the workbook's (owner 2026-10-13): three bills he had just paid
in QuickBooks stayed open - and on the pay run - because the ledger's Bill Tracker load was four days old while
the mirror (refreshed every 3 minutes) already showed them at 0."""
import datetime as dt
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "ledger"))
import dashboard as dash   # noqa: E402


def _bills():
    return [
        {"bill_id": "1", "open_balance": 545.21, "qbo_balance": 0.0, "pay_selected": True, "pay_amount": None},
        {"bill_id": "2", "open_balance": 100.0, "qbo_balance": 40.0, "pay_selected": True, "pay_amount": 100.0},
        {"bill_id": "3", "open_balance": 75.0, "qbo_balance": None, "pay_selected": False, "pay_amount": None},
        {"bill_id": None, "open_balance": 9.0, "qbo_balance": None, "pay_selected": False, "pay_amount": None},
    ]


def test_quickbooks_balance_wins_and_a_paid_bill_leaves_the_run():
    bills = _bills()
    drop = dash._apply_qbo_balance(bills, use_mirror=True)
    assert bills[0]["open_balance"] == 0 and bills[0]["pay_selected"] is False and bills[0]["pay_amount"] is None
    assert bills[1]["open_balance"] == 40.0 and bills[1]["pay_selected"] is True   # still owed: stays on the run
    assert bills[2]["open_balance"] == 75.0                                         # not in the mirror: the workbook stands
    assert drop == ["1"]


def test_an_older_mirror_changes_nothing():
    bills = _bills()
    assert dash._apply_qbo_balance(bills, use_mirror=False) == []
    assert [b["open_balance"] for b in bills] == [545.21, 100.0, 75.0, 9.0]
    assert bills[0]["pay_selected"] is True


def test_mirror_newer_than_the_workbook_load(monkeypatch):
    from shared import qbo_mirror
    monkeypatch.setattr(qbo_mirror, "stamp", lambda con=None: dt.datetime(2026, 10, 13, 21, 54, tzinfo=dt.timezone.utc))
    local = dt.datetime(2026, 10, 13, 21, 54, tzinfo=dt.timezone.utc).astimezone().replace(tzinfo=None)
    assert dash._mirror_newer_than("2026-10-09T09:04:49") is True
    assert dash._mirror_newer_than((local + dt.timedelta(minutes=1)).isoformat(timespec="seconds")) is False
    assert dash._mirror_newer_than(None) is True
    monkeypatch.setattr(qbo_mirror, "stamp", lambda con=None: None)
    assert dash._mirror_newer_than("2026-10-09T09:04:49") is False
