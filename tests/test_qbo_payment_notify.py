"""QuickBooks Payments Teams cards (invoice-sync/qbo_payment_notify.py, the user 2026-10-07).

Pure logic over a fake mirror - no QBO, no Teams:
  * only e-invoice / QuickBooks Payments payments count; a typed-in check never does
  * the first run marks everything seen and posts nothing
  * a new payment posts once; a failed post is retried; an old unseen one never posts
"""
import datetime as dt
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "invoice-sync"))

import qbo_payment_notify as q  # noqa: E402

NOW = dt.datetime(2026, 10, 7, 16, 0, tzinfo=dt.timezone.utc)
METHODS = [{"Id": "1", "Name": "Check"}, {"Id": "9", "Name": "QuickBooks Payments-Bank"}]
CUSTOMERS = [{"Id": "10", "DisplayName": "GC LLC"},
             {"Id": "11", "DisplayName": "RP7260-FTW - MABEL", "ParentRef": {"value": "10"}}]
INVOICES = {"500": {"Id": "500", "DocNumber": "34558", "Balance": 0}}


def _pay(pid, hours_ago, online=True):
    return {"Id": pid, "TotalAmt": 5000, "TxnDate": "2026-10-07",
            "TxnSource": "EInvoice" if online else None,
            "PaymentMethodRef": {"value": "9" if online else "1"},
            "CustomerRef": {"value": "11"},
            "MetaData": {"CreateTime": (NOW - dt.timedelta(hours=hours_ago)).isoformat()},
            "Line": [{"LinkedTxn": [{"TxnType": "Invoice", "TxnId": "500"}]}]}


def _mirror(monkeypatch, pays):
    data = {"Payment": pays, "PaymentMethod": METHODS, "Customer": CUSTOMERS}
    monkeypatch.setattr(q.qbo_mirror, "serves", lambda e: True)
    monkeypatch.setattr(q.qbo_mirror, "load", lambda e, *a, **k: data[e])
    monkeypatch.setattr(q.qbo_mirror, "get", lambda e, i, *a: INVOICES.get(i))


def test_only_online_payments_count():
    names = {m["Id"]: m["Name"] for m in METHODS}
    assert q.is_online(_pay("1", 1), names)
    assert not q.is_online(_pay("2", 1, online=False), names)
    assert q.is_online({"TxnSource": "INTUITMASPAYMENT"}, {})


def test_first_run_seeds_then_posts_new_once(monkeypatch, tmp_path):
    sent = []
    monkeypatch.setattr(q.teams, "post", lambda hook, payload: sent.append(payload) or True)
    _mirror(monkeypatch, [_pay("1", 1)])
    assert q.run("hook", tmp_path, now=NOW)["seeded"] == 1 and not sent

    _mirror(monkeypatch, [_pay("1", 1), _pay("2", 1), _pay("3", 1, online=False), _pay("4", 200)])
    assert q.run("hook", tmp_path, now=NOW)["posted"] == 1          # 2 only: 3 is a check, 4 is too old
    facts = {f["title"]: f["value"] for f in sent[0]["attachments"][0]["content"]["body"][1]["facts"]}
    assert facts["Client"] == "GC LLC" and facts["Invoice"] == "#34558 - paid in full"
    assert q.run("hook", tmp_path, now=NOW)["posted"] == 0          # never twice


def test_failed_post_retries_and_dry_run_leaves_state(monkeypatch, tmp_path):
    (tmp_path / q.STATE_FILE).write_text(json.dumps({"seeded": NOW.isoformat(), "posted": {}}))
    _mirror(monkeypatch, [_pay("7", 2)])
    assert q.run("", tmp_path, dry_run=True, now=NOW)["would_post"] == 1
    monkeypatch.setattr(q.teams, "post", lambda hook, payload: False)
    assert q.run("hook", tmp_path, now=NOW)["failed"] == 1
    monkeypatch.setattr(q.teams, "post", lambda hook, payload: True)
    assert q.run("hook", tmp_path, now=NOW)["posted"] == 1


def test_no_webhook_is_off(monkeypatch, tmp_path):
    _mirror(monkeypatch, [_pay("1", 1)])
    assert q.run("", tmp_path, now=NOW) == {"posted": 0, "failed": 0, "would_post": 0, "seeded": 0}
    assert not (tmp_path / q.STATE_FILE).exists()


def test_dry_run_with_no_record_lists_the_window(monkeypatch, tmp_path):
    _mirror(monkeypatch, [_pay("1", 1), _pay("2", 200)])
    assert q.run("", tmp_path, dry_run=True, now=NOW)["would_post"] == 1      # only the one inside 72h
    assert not (tmp_path / q.STATE_FILE).exists()


def test_server_test_mode_keeps_its_own_record_and_says_test(monkeypatch, tmp_path):
    sent = []
    monkeypatch.setattr(q.teams, "post", lambda hook, payload: sent.append(payload) or True)
    _mirror(monkeypatch, [_pay("1", 1)])
    q.run("hook", tmp_path, now=NOW, test=True)                              # first run seeds the TEST record
    _mirror(monkeypatch, [_pay("1", 1), _pay("2", 1)])
    assert q.run("hook", tmp_path, now=NOW, test=True)["posted"] == 1
    assert (tmp_path / q.TEST_STATE_FILE).exists() and not (tmp_path / q.STATE_FILE).exists()
    title = sent[0]["attachments"][0]["content"]["body"][0]["text"]
    assert title == "TEST - QuickBooks payment received"
