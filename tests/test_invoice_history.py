"""Audit > AR: one invoice's story (dashboard._invoice_history, 2026-10-11) - created, every logged change oldest
first, the payments applied, deleted or on file, and the Invoice Tracker page: the ledger's row when the tracker
still lists it, else the page kept in invoice_notion_ref (the loader fills it; the sync archives a deleted
invoice's page). In-memory mirror + a throwaway ledger file, no QuickBooks, no Notion."""
import sqlite3
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "ledger"))

import dashboard as d   # noqa: E402
import load_invoices   # noqa: E402
from shared import qbo_mirror as m   # noqa: E402


def _inv(total, balance, token="0", line_amt=None):
    return {"Id": "1", "SyncToken": token, "DocNumber": "34999", "TxnDate": "2026-09-30", "DueDate": "2026-10-30",
            "TotalAmt": total, "Balance": balance, "PrivateNote": "CP800 - Draw 3",
            "CustomerRef": {"value": "7", "name": "GC:CP800 - TOPAZ"},
            "MetaData": {"CreateTime": "2026-09-01T08:00:00-07:00", "LastUpdatedTime": "2026-10-01T09:00:00-07:00"},
            "LinkedTxn": [{"TxnId": "9", "TxnType": "Payment"}],
            "Line": [{"Id": "1", "Amount": line_amt if line_amt is not None else total, "DetailType": "SalesItemLineDetail",
                      "SalesItemLineDetail": {"ItemRef": {"name": "Draw"}}}]}


def _pay():
    return {"Id": "9", "SyncToken": "0", "TxnDate": "2026-10-02", "TotalAmt": 400.0, "PaymentRefNum": "CK 123",
            "CustomerRef": {"value": "7"}, "MetaData": {},
            "Line": [{"Amount": 400.0, "LinkedTxn": [{"TxnId": "1", "TxnType": "Invoice"}]}]}


def _ledger(tmp_path, with_row: bool):
    db = tmp_path / "ledger.sqlite3"
    con = sqlite3.connect(db)
    con.executescript((ROOT / "ledger" / "schema.sql").read_text(encoding="utf-8"))
    rec = {"qbo_txn_id": "1", "doc_number": "34999", "status": "Unpaid", "note": "GC paying Fri",
           "notion_url": "https://app.notion.com/p/34999-" + "a" * 32, "source": "invoice_tracker", "loaded_at": "t0"}
    cols = [c for c in load_invoices._COLS if c in rec]
    con.execute(f"INSERT INTO billing_event ({', '.join(cols)}) VALUES ({', '.join('?' * len(cols))})", [rec[c] for c in cols])
    con.commit()
    load_invoices.write_events(con, [dict(rec, loaded_at=None)] if with_row else [], "t1")   # fills the register, then (re)places the rows
    con.close()
    return db


def test_story_of_an_edited_then_deleted_invoice(tmp_path):
    con = m.connect(":memory:")
    m.upsert_many(con, "Invoice", [_inv(1000.0, 1000.0)], "s0")
    m.upsert_many(con, "Payment", [_pay()], "s0")
    m.upsert_many(con, "Invoice", [_inv(1200.0, 800.0, token="1", line_amt=1200.0)], "s1", track=True)
    db = _ledger(tmp_path, with_row=True)
    out = d._invoice_history("1", db_path=db, mcon=con)
    assert out["ok"]
    inv = out["invoice"]
    assert inv["created"] == "2026-09-01T08:00:00-07:00" and inv["doc_number"] == "34999" and inv["project"] == "CP800"
    assert inv["deleted"] is False and inv["total"] == 1200.0 and inv["balance"] == 800.0
    assert [h["kind"] for h in out["history"]] == ["edited"] and "amount changed" in out["history"][0]["flags"]
    assert inv["logged_since_birth"] is False
    assert out["payments"] == [{"id": "9", "date": "2026-10-02", "total": 400.0, "applied": 400.0, "ref": "CK 123", "deleted": False, "deleted_at": None}]
    assert out["notion"]["source"] == "ledger" and out["notion"]["status"] == "Unpaid" and out["notion"]["note"] == "GC paying Fri"

    # QuickBooks deletes it, and the next tracker load no longer lists it: the register still knows its page
    m.mark_deleted(con, "Invoice", ["1"], "s2", track=True, when_by_id={"1": "2026-10-10T10:00:00-07:00"})
    db2 = _ledger(tmp_path / "later", with_row=False) if (tmp_path / "later").mkdir() is None else None
    out = d._invoice_history("1", db_path=db2, mcon=con)
    assert out["ok"] and out["invoice"]["deleted"] is True and out["invoice"]["deleted_at"] == "2026-10-10T10:00:00-07:00"
    assert [h["kind"] for h in out["history"]] == ["edited", "deleted"] and out["history"][-1]["flags"] == ["deleted open invoice"]
    assert out["notion"]["source"] == "register" and out["notion"]["url"].endswith("a" * 32)


def test_unknown_invoice_and_no_ledger(tmp_path):
    con = m.connect(":memory:")
    assert d._invoice_history("404", db_path=None, mcon=con)["ok"] is False
    assert d._invoice_history("", db_path=None, mcon=con)["error"] == "no invoice id"


def test_register_only_grows(tmp_path):
    db = tmp_path / "l.sqlite3"
    con = sqlite3.connect(db)
    con.executescript((ROOT / "ledger" / "schema.sql").read_text(encoding="utf-8"))
    a = {"qbo_txn_id": "1", "doc_number": "1", "notion_url": "https://app.notion.com/p/1-" + "b" * 32, "source": "invoice_tracker"}
    b = {"qbo_txn_id": "2", "doc_number": "2", "notion_url": "https://app.notion.com/p/2-" + "c" * 32, "source": "invoice_tracker"}
    load_invoices.write_events(con, [a, b], "t1")
    load_invoices.write_events(con, [dict(b, notion_url=b["notion_url"] + "?v=2")], "t2")   # a dropped off the tracker
    rows = con.execute("SELECT qbo_txn_id, notion_url, first_seen, last_seen FROM invoice_notion_ref ORDER BY 1").fetchall()
    assert rows == [("1", a["notion_url"], "t1", "t1"), ("2", b["notion_url"] + "?v=2", "t1", "t2")]
    assert con.execute("SELECT COUNT(*) FROM billing_event").fetchone()[0] == 1
