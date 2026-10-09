"""The guardrails on the ledger's QBO money writes (ledger/pay_bills.py) - offline, no QBO, no Touch ID."""
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "ledger"))
sys.path.insert(0, str(ROOT))

import pay_bills as pb  # noqa: E402


def test_reference_cap_is_qbo_21():
    assert pb.clean_ref("ach", " 25783-123456831813153 ") == "25783-123456831813153"
    with pytest.raises(ValueError):
        pb.clean_ref("ach", "25783-1234568318131534")
    with pytest.raises(ValueError):
        pb.clean_ref("check", "")
    assert pb.clean_ref("print", "anything") == pb.TO_PRINT


def test_unauthorized_mac_is_refused(monkeypatch):
    monkeypatch.setattr(pb, "authorized", lambda: (False, "locked"))
    assert pb.commit({"run_token": "x" * 32})["error"] == "locked"
    assert pb.assign_number("123", "4567")["error"] == "locked"


def test_write_needs_a_live_review_token(monkeypatch):
    monkeypatch.setattr(pb, "authorized", lambda: (True, ""))
    monkeypatch.setattr(pb, "_audit", lambda *a, **k: None)
    assert "expired or was already used" in pb.commit({"run_token": "never-issued"})["error"]
    pb._TOKENS["t1"] = {"at": pb.time.time(), "vendors": {"9": {"100": 50.0}}, "credits": {"9": {}}, "names": {}}
    body = {"run_token": "t1", "vendors": [{"vendor_id": "9", "bills": [{"bill_id": "100", "amount": 51}]}]}
    assert "not what the review showed" in pb.commit(body)["error"]
    assert "expired or was already used" in pb.commit(body)["error"]      # spent even though it was refused


def test_approval_text_layout():
    t = pb.approval_text([("JMP", 10.5, "2026-09-28", "to print"), ("RCI", 2, "2026-09-28", "ACH 123")], "Operating ****0000", "to print")
    assert t.splitlines() == ["approve these vendor payments:", "JMP · $10.50 · 09/28/2026", "Operating ****0000 · to print",
                              "RCI · $2.00 · 09/28/2026", "Operating ****0000 · ACH 123", "Total $12.50"]


def _db(monkeypatch, tmp_path):
    monkeypatch.setattr(pb.bill_marks, "LEDGER_DB", tmp_path / "ledger.sqlite3")
    monkeypatch.setattr(pb, "_audit", lambda *a, **k: None)
    pb.bill_marks.set_pay_marks([{"bill_id": "100", "amount": 4000, "selected": True}], "2026-09-29T10:00:00")


def test_a_vendor_listed_twice_is_refused(monkeypatch):
    monkeypatch.setattr(pb, "authorized", lambda: (True, ""))
    monkeypatch.setattr(pb, "_audit", lambda *a, **k: None)
    pb._TOKENS["t2"] = {"at": pb.time.time(), "vendors": {"9": {"100": 50.0}}, "credits": {"9": {}}, "names": {}}
    v = {"vendor_id": "9", "bills": [{"bill_id": "100", "amount": 50}]}
    assert "listed twice" in pb.commit({"run_token": "t2", "vendors": [v, v]})["error"]


def test_no_answer_holds_the_partial_until_quickbooks_is_checked(monkeypatch, tmp_path):
    # review 2026-09-29 finding 1: 4,000 of a 10,000 bill, QBO never answered - the next review must NOT offer it again
    _db(monkeypatch, tmp_path)
    pb._attempt("r1", "sent", vendor_id="9", bills={"100": 4000.0}, body={"VendorRef": {"value": "9"}})
    pb._attempt("r1", "unknown", error="did not answer")
    monkeypatch.setattr(pb, "_find_posted", lambda *a: None)
    assert "got no answer" in pb._held("a", "c")["100"]
    monkeypatch.setattr(pb, "_find_posted", lambda *a: {"Id": "555"})            # found in QBO -> written, still held
    assert "already paid from the ledger (payment 555" in pb._held("a", "c")["100"]
    pb.bill_marks.set_pay_marks([{"bill_id": "100", "amount": 1000, "selected": True}], "2099-01-01T00:00:00")
    assert pb._held("a", "c") == {}                                                 # ticked again after: the owner asking


def test_an_unanswered_attempt_not_in_quickbooks_clears_after_the_settle_time(monkeypatch, tmp_path):
    _db(monkeypatch, tmp_path)
    pb._attempt("r2", "sent", vendor_id="9", bills={"100": 4000.0}, body={})
    con = pb._acon()
    con.execute("UPDATE pay_attempt SET at='2026-01-01T00:00:00'")
    con.commit()
    con.close()
    monkeypatch.setattr(pb, "_find_posted", lambda *a: None)
    assert pb._held("a", "c") == {}


def test_a_failed_login_is_an_error_not_an_exit(monkeypatch):
    def boom(*a, **k):
        raise SystemExit(1)
    monkeypatch.setattr(pb.qbo_api, "_api_get", boom)
    with pytest.raises(pb.qbo_api.AuthError):
        pb._get("/x", "a")


def _vendor_call(monkeypatch, tmp_path, live_bal, credits=None, approved=None):
    _db(monkeypatch, tmp_path)
    monkeypatch.setattr(pb, "LOG_DIR", tmp_path)
    monkeypatch.setattr(pb, "_live_ids", lambda a, c, typ, ids: {"100": {"Id": "100", "Balance": live_bal, "DocNumber": "B1"}})
    monkeypatch.setattr(pb, "_credits", lambda a, c, vid: credits or [])
    posts = []
    monkeypatch.setattr(pb, "_post", lambda cid, body, rid: posts.append(body) or {"Id": "777", "TotalAmt": body["TotalAmt"]})
    monkeypatch.setattr(pb, "_after_post", lambda *a, **k: None)
    lv = {"bills": [{"bill_id": "100", "amount": 4000.0, "balance": 10000.0, "doc": "B1"}]}
    v = {"vendor_id": "9", "bills": [{"bill_id": "100", "amount": 4000}], "credits": [c["id"] for c in credits or []]}
    res = {"vendor_id": "9", "vendor": "V"}
    t = {"credits": {"9": approved or {}}}
    pb._pay_vendor(v, "9", res, lv, t, "tok", "print", {"name": "Op"}, "35", "2026-09-29", {"9": pb.TO_PRINT}, "a", "c")
    return res, posts


def test_a_bill_paid_during_the_approval_is_not_written(monkeypatch, tmp_path):
    res, posts = _vendor_call(monkeypatch, tmp_path, live_bal=6000.0)          # someone paid 4,000 in QBO meanwhile
    assert not posts and "changed in QuickBooks during the approval" in res["error"]


def test_a_credit_that_moved_since_the_review_is_not_written(monkeypatch, tmp_path):
    res, posts = _vendor_call(monkeypatch, tmp_path, 10000.0, credits=[{"id": "c1", "doc": "", "balance": 300.0}], approved={"c1": 500.0})
    assert not posts and "credit's balance changed" in res["error"]


def test_same_payment_same_request_id(monkeypatch, tmp_path):
    res, posts = _vendor_call(monkeypatch, tmp_path, 10000.0)
    assert res["ok"] and len(posts) == 1
    con = pb._acon()
    rows = [dict(r) for r in con.execute("SELECT request_id, state, payment_id FROM pay_attempt")]
    con.close()
    assert rows[0]["state"] == "written" and rows[0]["payment_id"] == "777"
    import hashlib, json
    assert rows[0]["request_id"] == hashlib.sha1(json.dumps(posts[0], sort_keys=True).encode()).hexdigest()[:36]
