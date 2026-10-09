"""Clear copy - the bulk re-apply gate (ledger/reapply_check.py, owner 2026-10-06).

A check is offered for the bulk fix only when the copy on file is provably the check as it was AND every bill /
credit lands back exactly where it was. Credits and short pays are the cases that bite (owner: "be careful with
credits and shortpays"), so each has a passing case and the failure that must block it. No QuickBooks: the live
reads are faked, the change log is a stub index."""
import copy
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "ledger"))

import reapply_check as rc  # noqa: E402

PID = "900"
STRIP = "2026-10-01T09:00:00-07:00"


def _pay(lines, total, **kw):
    rec = {"Id": PID, "DocNumber": "50001", "TotalAmt": total, "TxnDate": "2026-09-26", "PayType": "Check",
           "VendorRef": {"value": "7", "name": "Sub A"}, "CheckPayment": {"BankAccountRef": {"value": "35"}},
           "SyncToken": "4",
           "Line": [{"Amount": a, "LinkedTxn": [{"TxnId": i, "TxnType": t}]} for t, i, a in lines]}
    rec.update(kw)
    return rec


def _doc(typ, i, total, balance, paid_by=PID, vendor="7"):
    return {"Id": i, "DocNumber": f"{typ[0]}{i}", "TotalAmt": total, "Balance": balance, "TxnDate": "2026-09-20",
            "VendorRef": {"value": vendor}, "PrivateNote": "",
            "LinkedTxn": [{"TxnId": paid_by, "TxnType": "BillPaymentCheck"}] if paid_by else []}


class FakeIx(rc._Index):
    def __init__(self):
        self.by, self._before, self.n = {"BillPayment": {}, "Bill": {}, "VendorCredit": {}}, {}, 0

    def add(self, ent, rec_id, before, changed_at=STRIP, bal_after=None):
        self.n += 1
        self._before[self.n] = before
        self.by[ent].setdefault(str(rec_id), []).append({
            "id": self.n, "changed_at": changed_at, "has_before": True,
            "balance_before": before.get("Balance"), "balance_after": bal_after})


def _world(before_lines, total, pre_docs, live_lines=(), live_docs=None, live_pay_kw=None):
    """pre_docs: {(type, id): (total, balance while the check was on it)}; live_docs: {(type, id): (total, balance now)}."""
    ix = FakeIx()
    ix.add("BillPayment", PID, _pay(before_lines, total))
    for (t, i), (tot, bal) in pre_docs.items():
        ix.add(t, i, _doc(t, i, tot, bal))
    live = _pay(list(live_lines), total, **(live_pay_kw or {}))
    now = {}
    for (t, i), (tot, bal) in (live_docs or {}).items():
        now.setdefault(t, {})[i] = _doc(t, i, tot, bal, paid_by=None)
    return ix, live, now


@pytest.fixture
def run(monkeypatch):
    def go(ix, live, now):
        monkeypatch.setattr(rc, "_get_payment", lambda a, c, p: copy.deepcopy(live))
        monkeypatch.setattr(rc, "_live", lambda a, c, typ, ids: {i: now.get(typ, {})[i] for i in ids if i in now.get(typ, {})})
        ev = rc.evidence(ix, PID, live)
        p = rc.plan("x", "y", PID, ev["before"] or {"Line": []}, copy.deepcopy(live))
        why = rc._proved(ev, p)
        return why, rc._check_json(PID, ev, p, why)
    return go


def test_plain_full_pay_proves(run):
    why, j = run(*_world([("Bill", "1", 1000.0)], 1000.0, {("Bill", "1"): (1000.0, 0.0)},
                         live_docs={("Bill", "1"): (1000.0, 1000.0)}))
    assert why == [] and j["ok"] and j["amount"] == 1000.0 and j["applied_after"] == 1000.0


def test_credit_goes_back_to_where_it_was(run):
    why, j = run(*_world([("Bill", "1", 1000.0), ("VendorCredit", "5", 200.0)], 800.0,
                         {("Bill", "1"): (1000.0, 0.0), ("VendorCredit", "5"): (200.0, 0.0)},
                         live_docs={("Bill", "1"): (1000.0, 1000.0), ("VendorCredit", "5"): (200.0, 200.0)}))
    assert why == [] and j["applied_after"] == 800.0 and [c["doc"] for c in j["credits"]] == ["V5"]


def test_credit_used_on_another_check_since_blocks(run):
    why, _ = run(*_world([("Bill", "1", 1000.0), ("VendorCredit", "5", 200.0)], 800.0,
                         {("Bill", "1"): (1000.0, 0.0), ("VendorCredit", "5"): (200.0, 0.0)},
                         live_docs={("Bill", "1"): (1000.0, 1000.0), ("VendorCredit", "5"): (200.0, 0.0)}))
    assert why                                     # the bill alone would over-apply; the credit cannot go back


def test_credit_partly_used_since_blocks(run):
    why, _ = run(*_world([("Bill", "1", 1000.0), ("VendorCredit", "5", 200.0)], 800.0,
                         {("Bill", "1"): (1000.0, 0.0), ("VendorCredit", "5"): (300.0, 100.0)},
                         live_docs={("Bill", "1"): (1000.0, 1000.0), ("VendorCredit", "5"): (300.0, 250.0)}))
    assert any("credit" in w for w in why)        # would land at 50, was 100 - 50 of it went elsewhere


def test_credit_proved_by_its_release_when_it_does_not_name_the_check(run):
    ix, live, now = _world([("Bill", "1", 1000.0)], 1000.0, {("Bill", "1"): (1000.0, 0.0)},
                           live_docs={("Bill", "1"): (1000.0, 1000.0), ("VendorCredit", "5"): (200.0, 200.0)})
    ix.by["BillPayment"][PID][0]["changed_at"] = STRIP
    ix._before[1] = _pay([("Bill", "1", 1000.0), ("VendorCredit", "5", 200.0)], 800.0)
    live = _pay([], 800.0)
    rel = _doc("VendorCredit", "5", 200.0, 0.0, paid_by=None)
    ix.add("VendorCredit", "5", rel, changed_at="2026-10-01T09:00:40-07:00", bal_after=200.0)
    why, j = run(ix, live, now)
    assert why == [] and j["credits"]


def test_short_pay_keeps_its_open_amount(run):
    why, j = run(*_world([("Bill", "1", 600.0)], 600.0, {("Bill", "1"): (1000.0, 400.0)},
                         live_docs={("Bill", "1"): (1000.0, 1000.0)}))
    assert why == [] and j["short_pays"] == [{"doc": "B1", "id": "1", "amount": 600.0, "bill_total": 1000.0, "stays_open": 400.0}]


def test_short_pay_remainder_paid_since_blocks(run):
    why, _ = run(*_world([("Bill", "1", 600.0)], 600.0, {("Bill", "1"): (1000.0, 400.0)},
                         live_docs={("Bill", "1"): (1000.0, 600.0)}))
    assert any("would end at" in w for w in why)  # 600 back on -> 0 open, but it was 400 short before


def test_bill_amount_edited_blocks(run):
    why, _ = run(*_world([("Bill", "1", 1000.0)], 1000.0, {("Bill", "1"): (1000.0, 0.0)},
                         live_docs={("Bill", "1"): (1100.0, 1100.0)}))
    assert any("total changed" in w for w in why)


def test_bill_paid_again_blocks(run):
    why, _ = run(*_world([("Bill", "1", 1000.0)], 1000.0, {("Bill", "1"): (1000.0, 0.0)},
                         live_docs={("Bill", "1"): (1000.0, 0.0)}))
    assert why


def test_deleted_bill_blocks(run):
    why, _ = run(*_world([("Bill", "1", 1000.0)], 1000.0, {("Bill", "1"): (1000.0, 0.0)}, live_docs={}))
    assert why


def test_money_moved_to_a_later_bill_blocks(run):
    why, _ = run(*_world([("Bill", "1", 1000.0)], 1000.0, {("Bill", "1"): (1000.0, 0.0)},
                         live_lines=[("Bill", "2", 135.4)],
                         live_docs={("Bill", "1"): (1000.0, 1000.0), ("Bill", "2"): (500.0, 364.6)}))
    assert any("another bill" in w for w in why)


def test_copy_not_whole_blocks(run):
    why, _ = run(*_world([("Bill", "1", 900.0)], 1000.0, {("Bill", "1"): (900.0, 0.0)},
                         live_docs={("Bill", "1"): (900.0, 900.0)}))
    assert any("not the whole check" in w for w in why)


def test_two_different_copies_block(run):
    ix, live, now = _world([("Bill", "1", 1000.0)], 1000.0, {("Bill", "1"): (1000.0, 0.0)},
                           live_docs={("Bill", "1"): (1000.0, 1000.0)})
    ix.add("BillPayment", PID, _pay([("Bill", "2", 1000.0)], 1000.0), changed_at="2026-09-24T09:00:00-07:00")
    why, _ = run(ix, live, now)
    assert any("two different copies" in w for w in why)


def test_check_itself_changed_blocks(run):
    why, _ = run(*_world([("Bill", "1", 1000.0)], 1000.0, {("Bill", "1"): (1000.0, 0.0)},
                         live_docs={("Bill", "1"): (1000.0, 1000.0)}, live_pay_kw={"TxnDate": "2026-09-30"}))
    assert any("check itself was changed" in w for w in why)


def test_no_copy_blocks(run):
    ix = FakeIx()
    live = _pay([], 1000.0)
    why, j = run(ix, live, {})
    assert why and not j["ok"]


def test_no_bill_copy_blocks(run):
    ix, live, now = _world([("Bill", "1", 1000.0)], 1000.0, {}, live_docs={("Bill", "1"): (1000.0, 1000.0)})
    why, _ = run(ix, live, now)
    assert any("no copy of bill" in w for w in why)


def test_bulk_write_needs_touch_id(monkeypatch):
    import presence
    monkeypatch.setattr(presence, "confirm", lambda reason: (False, "cancelled"))
    monkeypatch.setattr(rc, "_get_payment", lambda *a: pytest.fail("read QuickBooks without Touch ID"))
    res = rc.bulk_commit([{"payment_id": PID, "sync_token": "4", "n": 1, "amount": 1000.0}])
    assert not res["ok"] and "nothing was written" in res["error"]
