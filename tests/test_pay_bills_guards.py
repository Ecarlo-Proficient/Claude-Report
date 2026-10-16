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


def test_plan_on_given_bills_is_scoped_and_remembered(monkeypatch):
    # the vendor page (owner 2026-10-13): the review runs on the ticked bills, not the saved run, and the token
    # remembers which, so the write re-reads the same bills
    monkeypatch.setattr(pb, "_login", lambda: ("a", "c"))
    seen = {}

    def fake_plan_live(access, cid, marks=None):
        seen["marks"] = marks
        return {"vendors": {"9": {"vendor_id": "9", "vendor": "RCI", "bills": [{"bill_id": "100", "amount": 50.0}], "total": 50.0}}, "skipped": []}
    monkeypatch.setattr(pb, "_plan_live", fake_plan_live)
    monkeypatch.setattr(pb, "_credits", lambda *a: [])
    monkeypatch.setattr(pb, "accounts", lambda: [])
    monkeypatch.setattr(pb, "authorized", lambda: (True, ""))
    monkeypatch.setattr(pb, "_close_date", lambda *a: "")
    d = pb.plan(["100", " ", "101"])
    assert seen["marks"] == {"100": {"amount": None}, "101": {"amount": None}}
    assert d["scope"] == "bills" and pb._TOKENS[d["run_token"]]["marks"] == seen["marks"]
    d2 = pb.plan()
    assert seen["marks"] is None and d2["scope"] == "run" and pb._TOKENS[d2["run_token"]]["marks"] is None


def test_plan_live_on_given_bills_never_reads_the_saved_run(monkeypatch):
    def boom():
        raise AssertionError("read the saved run")
    monkeypatch.setattr(pb.bill_marks, "read_pay_marks", boom)
    bill = {"Id": "100", "DocNumber": "7", "TxnDate": "2026-10-01", "Balance": 80.0, "SyncToken": "1",
            "VendorRef": {"value": "9", "name": "RCI"}, "PrivateNote": "", "MetaData": {}}
    monkeypatch.setattr(pb, "_live_ids", lambda a, c, typ, ids: {"100": bill})
    monkeypatch.setattr(pb, "_held", lambda a, c: {})
    p = pb._plan_live("a", "c", pb.marks_for(["100"]))
    assert p["vendors"]["9"]["bills"][0]["amount"] == 80.0 and p["skipped"] == []
    assert p["vendors"]["9"]["bills"][0]["partial"] is False


# ── a payment MORE than its bills = a credit with the vendor (owner 2026-10-13, a joint check on a bill already paid) ──
def test_a_typed_amount_above_the_bills_becomes_the_credit(monkeypatch, tmp_path):
    _db(monkeypatch, tmp_path)
    monkeypatch.setattr(pb, "LOG_DIR", tmp_path)
    monkeypatch.setattr(pb, "_live_ids", lambda a, c, typ, ids: {"100": {"Id": "100", "Balance": 10000.0, "DocNumber": "B1"}})
    monkeypatch.setattr(pb, "_credits", lambda a, c, vid: [])
    posts = []
    monkeypatch.setattr(pb, "_post", lambda cid, body, rid: posts.append(body) or {"Id": "777", "TotalAmt": body["TotalAmt"]})
    monkeypatch.setattr(pb, "_after_post", lambda *a, **k: None)
    monkeypatch.setattr(pb, "_number_taken", lambda *a, **k: None)     # a typed check #: the mirror is not read here
    lv = {"bills": [{"bill_id": "100", "amount": 4000.0, "balance": 10000.0, "doc": "B1"}], "paid_bills": []}
    v = {"vendor_id": "9", "bills": [{"bill_id": "100", "amount": 4000}], "credits": [], "amount": 4301.0}
    res = {"vendor_id": "9", "vendor": "V"}
    pb._pay_vendor(v, "9", res, lv, {"credits": {}}, "tok", "check", {"name": "Op"}, "35", "2026-10-13", {"9": "20119"}, "a", "c")
    assert res["ok"] and res["credit"] == 301.0 and "warn" not in res
    assert posts[0]["TotalAmt"] == 4301.0 and [ln["Amount"] for ln in posts[0]["Line"]] == [4000.0]   # the excess is NO line
    # the same bills with no amount typed: the payment equals its bills, no credit
    posts.clear()
    res = {"vendor_id": "9", "vendor": "V"}
    pb._pay_vendor({**v, "amount": None}, "9", res, lv, {"credits": {}}, "tok", "check", {"name": "Op"}, "35", "2026-10-13", {"9": "20120"}, "a", "c")
    assert res["ok"] and res["credit"] == 0 and posts[0]["TotalAmt"] == 4000.0


def test_an_amount_below_the_bills_is_refused(monkeypatch, tmp_path):
    _db(monkeypatch, tmp_path)
    monkeypatch.setattr(pb, "LOG_DIR", tmp_path)
    monkeypatch.setattr(pb, "_live_ids", lambda a, c, typ, ids: {"100": {"Id": "100", "Balance": 10000.0, "DocNumber": "B1"}})
    monkeypatch.setattr(pb, "_credits", lambda a, c, vid: [])
    posts = []
    monkeypatch.setattr(pb, "_post", lambda cid, body, rid: posts.append(body))
    lv = {"bills": [{"bill_id": "100", "amount": 4000.0, "balance": 10000.0, "doc": "B1"}]}
    v = {"vendor_id": "9", "bills": [{"bill_id": "100", "amount": 4000}], "credits": [], "amount": 3999.0}
    res = {"vendor_id": "9", "vendor": "V"}
    pb._pay_vendor(v, "9", res, lv, {"credits": {}}, "tok", "print", {"name": "Op"}, "35", "2026-10-13", {"9": pb.TO_PRINT}, "a", "c")
    assert not posts and "less than the bills" in res["error"]


def test_paid_bills_only_pay_as_a_credit_with_no_lines(monkeypatch, tmp_path):
    """From the vendor page: the ticked bills are already paid (the joint check case) - the vendor stays on the
    review with paid_bills, the typed amount is the whole payment and QuickBooks gets no Line at all."""
    _db(monkeypatch, tmp_path)
    monkeypatch.setattr(pb, "LOG_DIR", tmp_path)
    monkeypatch.setattr(pb, "_live_ids", lambda a, c, typ, ids: {})
    monkeypatch.setattr(pb, "_credits", lambda a, c, vid: [])
    posts = []
    monkeypatch.setattr(pb, "_post", lambda cid, body, rid: posts.append(body) or {"Id": "778", "TotalAmt": body["TotalAmt"]})
    monkeypatch.setattr(pb, "_after_post", lambda *a, **k: None)
    monkeypatch.setattr(pb, "_number_taken", lambda *a, **k: None)
    lv = {"bills": [], "paid_bills": [{"bill_id": "100", "doc": "B1", "date": "2026-09-01", "memo": "", "total": 500.0}]}
    v = {"vendor_id": "9", "bills": [], "credits": [], "amount": 500.0}
    res = {"vendor_id": "9", "vendor": "V"}
    pb._pay_vendor(v, "9", res, lv, {"credits": {}}, "tok", "check", {"name": "Op"}, "35", "2026-10-13", {"9": "20119"}, "a", "c")
    assert res["ok"] and res["credit"] == 500.0 and res["bills"] == 0
    assert posts[0]["TotalAmt"] == 500.0 and "Line" not in posts[0]
    # no amount typed on a credit-only payment: nothing to write
    posts.clear()
    res = {"vendor_id": "9", "vendor": "V"}
    pb._pay_vendor({**v, "amount": None}, "9", res, lv, {"credits": {}}, "tok", "check", {"name": "Op"}, "35", "2026-10-13", {"9": "20119"}, "a", "c")
    assert not posts and "amount is 0" in res["error"]
    # the saved run never carries a paid bill's vendor; the vendor page's ticks do
    paid = {"Id": "100", "Balance": 0, "TotalAmt": 500.0, "DocNumber": "B1", "TxnDate": "2026-09-01", "VendorRef": {"value": "9", "name": "V"},
            "PrivateNote": "RP1 - paid by us 09/15", "MetaData": {"CreateTime": "2026-09-01"}}
    monkeypatch.setattr(pb, "_live_ids", lambda a, c, typ, ids: {"100": paid})
    monkeypatch.setattr(pb, "_held", lambda a, c: {})
    p = pb._plan_live("a", "c", pb.marks_for(["100"]))
    assert p["vendors"]["9"]["bills"] == [] and p["vendors"]["9"]["paid_bills"][0]["doc"] == "B1" and not p["skipped"]
    monkeypatch.setattr(pb.bill_marks, "read_pay_marks", lambda: {"100": {"amount": None}})
    p = pb._plan_live("a", "c")
    assert not p["vendors"] and p["skipped"][0]["why"] == "already paid in QuickBooks"
