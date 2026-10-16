"""Client payments QBO stripped - detection (ledger/payment_drift.py) and the re-apply plan (shared/payment_reapply.py).

Built from the worked case of 10/16/2026: BKGP23052 ($240) moved to a project by hand, QBO took it off check #0087532
and took $240 off invoice 34527 instead (payment total unchanged, SyncToken unchanged). No QuickBooks: plan() is pure,
the mirror is faked."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "ledger"))

import payment_drift as pdrift  # noqa: E402
from shared import payment_reapply as pr  # noqa: E402

PID = "1316642"


def _pay(lines, total=11177.45, **kw):
    rec = {"Id": PID, "PaymentRefNum": "0087532", "TxnDate": "2026-08-17", "TotalAmt": total, "UnappliedAmt": 0,
           "SyncToken": "1", "CustomerRef": {"value": "23553", "name": "Grand Homes, Inc."},
           "Line": [{"Amount": a, "LinkedTxn": [{"TxnId": i, "TxnType": t}]} for t, i, a in lines]}
    rec.update(kw)
    return rec


BEFORE = [("Invoice", "1314933", 265.0), ("Invoice", "1315763", 5463.9), ("Invoice", "1315737", 5983.55),
          ("CreditMemo", "1316640", 295.0), ("CreditMemo", "1316639", 240.0)]
STRIPPED = [("Invoice", "1314933", 265.0), ("Invoice", "1315763", 5463.9), ("Invoice", "1315737", 5743.55),
            ("CreditMemo", "1316640", 295.0)]


def _inv(i, total, bal):
    return {"Id": i, "DocNumber": f"I{i}", "TotalAmt": total, "Balance": bal, "CustomerRef": {"name": "RP7450-FTW"}}


def _cm(i, total, rem):
    return {"Id": i, "DocNumber": f"BK{i}", "TotalAmt": total, "RemainingCredit": rem, "Balance": rem,
            "CustomerRef": {"name": "RP7307"}}


def _docs(inv_open=240.0, cm_open=240.0):
    return {("Invoice", "1314933"): _inv("1314933", 265.0, 0), ("Invoice", "1315763"): _inv("1315763", 5463.9, 0),
            ("Invoice", "1315737"): _inv("1315737", 5983.55, inv_open),
            ("CreditMemo", "1316640"): _cm("1316640", 295.0, 0), ("CreditMemo", "1316639"): _cm("1316639", 240.0, cm_open)}


def test_stripped_payment_puts_back_credit_and_invoice():
    p = pr.plan(_pay(BEFORE), _pay(STRIPPED), _docs())
    assert p["ok"], p["why"]
    back = {(b["type"], b["id"]): (b["amount"], b["open_after"]) for b in p["back"]}
    assert back == {("Invoice", "1315737"): (240.0, 0.0), ("CreditMemo", "1316639"): (240.0, 0.0)}


def test_body_carries_every_line_of_the_copy():
    body = pr.body_for(_pay(STRIPPED, SyncToken="7"), _pay(BEFORE))
    assert body["SyncToken"] == "7" and sorted(pr.links(body)) == sorted(pr.links(_pay(BEFORE)))
    assert "MetaData" not in body


def test_nothing_to_do_when_already_back():
    p = pr.plan(_pay(BEFORE), _pay(BEFORE), _docs(0, 0))
    assert not p["ok"] and p["why"] == [pr.NOTHING]


def test_credit_used_elsewhere_blocks():
    p = pr.plan(_pay(BEFORE), _pay(STRIPPED), _docs(cm_open=0))
    assert not p["ok"] and any("used somewhere else" in w for w in p["why"])


def test_payment_amount_changed_blocks():
    p = pr.plan(_pay(BEFORE), _pay(STRIPPED, total=10000.0), _docs())
    assert not p["ok"] and any("amount changed" in w for w in p["why"])


def test_new_link_on_payment_blocks():
    p = pr.plan(_pay(BEFORE), _pay(STRIPPED + [("Invoice", "999", 240.0)]), _docs())
    assert not p["ok"] and any("more than the copy" in w for w in p["why"])


def test_deleted_document_blocks():
    d = _docs()
    del d[("CreditMemo", "1316639")]
    p = pr.plan(_pay(BEFORE), _pay(STRIPPED), d)
    assert not p["ok"] and any("no longer exists" in w for w in p["why"])


def _fake_mirror(monkeypatch, pays, docs):
    class Con:
        def close(self):
            pass
    by = {"Payment": pays, "Invoice": [], "CreditMemo": []}
    for (t, _i), rec in docs.items():
        by[t].append(rec)
    monkeypatch.setattr(pdrift.mirror, "connect", lambda *a, **k: Con())
    monkeypatch.setattr(pdrift.mirror, "load", lambda e, *a, **k: by[e])


def test_detects_the_stripped_payment(monkeypatch):
    # The mirror keeps the payment as it was (QBO did not bump it); the credit and invoice show open.
    _fake_mirror(monkeypatch, [_pay(BEFORE)], _docs())
    res = pdrift.audit()
    assert res["count"] == 1 and res["amount"] == 240.0
    assert res["rows"][0]["credits_off"] == ["BK1316639"] and res["rows"][0]["invoices_reopened"] == ["I1315737"]


def test_clean_payment_is_not_flagged(monkeypatch):
    _fake_mirror(monkeypatch, [_pay(BEFORE)], _docs(0, 0))
    assert pdrift.audit()["count"] == 0


def test_invoice_paid_by_two_payments_is_not_flagged(monkeypatch):
    other = _pay([("Invoice", "1315737", 100.0)], total=100.0, Id="2")
    d = _docs(0, 0)
    d[("Invoice", "1315737")] = _inv("1315737", 6083.55, 0)
    _fake_mirror(monkeypatch, [_pay(BEFORE), other], d)
    assert pdrift.audit()["count"] == 0
