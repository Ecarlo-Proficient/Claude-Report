"""After a payment is written, BOTH bill payment stubs are filed (owner 2026-10-16: "add show both internal and
external to print ... it didn't create a external/internal stub after paying"): the queue row carries each stub's
file, error and print-history id, an older queue table gains the columns on first open, and the result the page
reads says which stub is filed and how to open both. No QuickBooks, no Chrome - print_stub is replaced."""
import json
import sqlite3
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "ledger"))

import bill_payment_stub as bps   # noqa: E402
import pay_bills as pb            # noqa: E402


def _fresh_db(monkeypatch, tmp_path):
    db = tmp_path / "ledger.sqlite3"
    monkeypatch.setattr(pb.bill_marks, "LEDGER_DB", db)
    return db


def test_an_older_queue_table_gains_the_internal_stub_columns(monkeypatch, tmp_path):
    db = _fresh_db(monkeypatch, tmp_path)
    con = sqlite3.connect(db)
    con.execute(pb._PUSH_TABLE)                      # the table as it was before 10/16
    con.execute("INSERT INTO pay_push (payment_id, pushed_at) VALUES ('1', '2026-10-01T00:00:00')")
    con.commit(); con.close()
    con = pb._pcon()
    have = {r[1] for r in con.execute("PRAGMA table_info(pay_push)")}
    assert {"stub_internal", "stub_internal_error", "stub_ids"} <= have
    row = con.execute("SELECT * FROM pay_push WHERE payment_id='1'").fetchone()
    assert row["stub_internal"] is None and row["stub_ids"] is None
    con.close()


def test_stub_files_both_kinds_and_hands_the_invoice_map_to_the_internal_one(monkeypatch):
    calls = []

    def fake_print(pid, columns=None, company="", db_path=None, kind="vendor", invoice_of=None, invoices=False):
        calls.append((pid, kind, invoice_of, invoices))
        return {"ok": True, "id": 10 if kind == "vendor" else 11, "file": f"/x/{kind}.pdf"}
    monkeypatch.setattr(bps, "print_stub", fake_print)
    monkeypatch.setattr(pb, "invoice_map_for", lambda pid: {"b1": {"invoice_no": "34457"}})
    out = pb._stub("777")
    assert out["filed"] and out["id"] == 10 and out["file"].endswith("vendor.pdf")
    assert out["internal"]["filed"] and out["internal"]["id"] == 11
    assert [c[1] for c in calls] == ["vendor", "internal"]
    assert calls[0][2] is None and calls[1][2] == {"b1": {"invoice_no": "34457"}}   # the map feeds the internal stub only
    assert all(c[3] is False for c in calls)                                        # client invoices on top: never by default


def test_a_failed_internal_stub_never_undoes_the_external_one(monkeypatch):
    def fake_print(pid, columns=None, company="", db_path=None, kind="vendor", invoice_of=None, invoices=False):
        if kind == "internal":
            raise RuntimeError("Chrome died")
        return {"ok": True, "id": 5, "file": "/x/v.pdf"}
    monkeypatch.setattr(bps, "print_stub", fake_print)
    monkeypatch.setattr(pb, "invoice_map_for", None)
    out = pb._stub("1")
    assert out["filed"] and not out["internal"]["filed"] and "Chrome died" in out["internal"]["error"]


def test_store_and_result_round_trip(monkeypatch, tmp_path):
    _fresh_db(monkeypatch, tmp_path)
    con = pb._pcon()
    con.execute("INSERT INTO pay_push (payment_id, vendor, pushed_at, qbo_state, qbo_ref) VALUES ('9', 'V', '2026-10-16T11:00:00', 'matched', '37824')")
    row = con.execute("SELECT * FROM pay_push WHERE payment_id='9'").fetchone()
    # both filed
    pb._store_stubs(con, "9", row, {"filed": True, "file": "/s/v.pdf", "id": 65}, {"filed": True, "file": "/s/Internal/i.pdf", "id": 66}, "37824")
    r = con.execute("SELECT * FROM pay_push WHERE payment_id='9'").fetchone()
    assert r["stub_file"] == "/s/v.pdf" and r["stub_ref"] == "37824" and r["stub_at"] and r["stub_error"] is None
    assert r["stub_internal"] == "/s/Internal/i.pdf" and r["stub_internal_error"] is None
    assert json.loads(r["stub_ids"]) == {"vendor": 65, "internal": 66}
    res = pb._stub_result(r)
    assert res["filed"] and res["id"] == 65 and res["internal"]["filed"] and res["internal"]["id"] == 66
    # the internal re-filed alone and failing: the external untouched, the last good internal file kept, the error recorded
    pb._store_stubs(con, "9", r, None, {"filed": False, "error": "share unmounted"}, "37824")
    r2 = con.execute("SELECT * FROM pay_push WHERE payment_id='9'").fetchone()
    assert r2["stub_file"] == "/s/v.pdf" and r2["stub_at"] == r["stub_at"]
    assert r2["stub_internal"] == "/s/Internal/i.pdf" and r2["stub_internal_error"] == "share unmounted"
    assert json.loads(r2["stub_ids"]) == {"vendor": 65, "internal": 66}
    res2 = pb._stub_result(r2)
    assert res2["filed"] and not res2["internal"]["filed"] and res2["internal"]["error"] == "share unmounted"
    # a row with nothing filed yet
    assert pb._stub_result(None) == {"filed": False, "file": None, "error": None, "id": None,
                                     "internal": {"filed": False, "file": None, "error": None, "id": None}}
    # a row pushed before 10/16 never recorded its external stub's id: the print history fills it
    con.execute("UPDATE pay_push SET stub_ids=NULL WHERE payment_id='9'")
    assert pb._stub_result(con.execute("SELECT * FROM pay_push WHERE payment_id='9'").fetchone(), con)["id"] is None   # no history table: nothing to fill
    con.execute("CREATE TABLE bill_payment_stub (id INTEGER PRIMARY KEY, payment_id TEXT, kind TEXT)")
    con.execute("INSERT INTO bill_payment_stub (id, payment_id, kind) VALUES (40, '9', 'vendor'), (41, '9', 'vendor'), (42, '9', 'internal')")
    r3 = con.execute("SELECT * FROM pay_push WHERE payment_id='9'").fetchone()
    assert pb._history_ids(con, "9") == {"vendor": 41, "internal": 42}                     # the latest print of each kind
    assert pb._history_ids(con, "9", {"internal": 66}) == {"vendor": 41, "internal": 66}   # the row's own id wins
    assert pb._stub_result(r3, con)["id"] == 41 and pb._stub_result(r3, con)["internal"]["id"] == 42
    con.close()
