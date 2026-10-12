"""The REVISED-bill follow-up (bill-tracker/revised_bills.py + shared/revised_board.py, 10/12/2026):
finding revised bills in a stand-in QuickBooks copy, the vendor's original total from the change
log, what the statement record says, the pickles flagged, and the board plan (forward-only stages,
the clerk's columns untouched). No network, no share."""
import datetime as dt
import json
import sqlite3
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "bill-tracker"))

import revised_bills as rv  # noqa: E402
from shared import qbo_mirror, revised_board as rb, statement_record as sr  # noqa: E402

TODAY = dt.date(2026, 10, 12)


def _con(monkeypatch):
    """An in-memory stand-in for the mirror: qbo_bill + mirror_change with the columns find() reads;
    json stored plain (the mirror's _unpack is monkeypatched to json.loads)."""
    con = sqlite3.connect(":memory:")
    con.execute("CREATE TABLE qbo_bill (id TEXT, txn_date TEXT, doc_number TEXT, total REAL, created TEXT, last_updated TEXT, deleted INT, json TEXT)")
    con.execute("CREATE TABLE mirror_change (entity TEXT, rec_id TEXT, kind TEXT, changed_at TEXT, total_before REAL, total_after REAL, flags TEXT)")
    monkeypatch.setattr(qbo_mirror, "_unpack", lambda b: json.loads(b))
    return con


def _bill(con, bid, vendor, ref, date, total, memo, created="2026-09-30T10:00:00-07:00", balance=None, deleted=0):
    j = {"VendorRef": {"name": vendor}, "PrivateNote": memo, "Balance": total if balance is None else balance,
         "Line": [{"ItemBasedExpenseLineDetail": {"CustomerRef": {"name": f"BUILDER:{memo.split(' - ')[0] if memo.startswith('RP') else 'RP0000'}"}}}]}
    con.execute("INSERT INTO qbo_bill VALUES (?,?,?,?,?,?,?,?)", (bid, date, ref, total, created, created, deleted, json.dumps(j)))


def _change(con, bid, kind, at, before, after, flags=""):
    con.execute("INSERT INTO mirror_change VALUES ('Bill',?,?,?,?,?,?)", (bid, kind, at, before, after, flags))


def _records():
    rec = {"vendor": "RCI READY CABLE", "as_of": "2026-09-30", "lines": [
        {"ref": "RW788288", "amount": 966.08, "bucket": "matched"},
        {"ref": "RW789171", "amount": 726.31, "bucket": "mismatch"}]}
    return {sr.key("RCI READY CABLE"): rec}


def test_find_reads_the_revised_memo_the_original_total_and_the_statement(monkeypatch):
    con = _con(monkeypatch)
    _bill(con, "1", "RCI READY CABLE", "RW788288", "2026-09-01", 966.08, "NOT APPROVED - RP7502 - 8606 GREAT STONE TRL - GRAND HOMES\n\nREVISED - NO DOWELS ORDERED")
    _change(con, "1", "created", "2026-09-29T11:00:00-07:00", None, 1156.20)
    _change(con, "1", "edited", "2026-10-01T07:03:03-07:00", 1156.20, 1045.78, "amount changed,lines changed")
    _change(con, "1", "edited", "2026-10-09T07:55:08-07:00", 1045.78, 966.08, "amount changed,lines changed")
    _bill(con, "2", "RCI READY CABLE", "RW789171", "2026-09-03", 670.96, "RP7622 - 1622 EAST ELMORE AVE\n\nREVISED - TAX EXEMPT JOB")
    _bill(con, "3", "RCI READY CABLE", "RW700000", "2026-09-03", 100.0, "RP7622 - plain bill, no revision")
    _bill(con, "4", "CARDER CONCRETE", "R1", "2026-09-03", 50.0, "REVISED - CP861 - 4231 STATE HWY 161", created="2026-03-01T10:00:00-07:00")   # paper-era line-1: not a follow-up
    found = rv.find(con, since=dt.date(2026, 8, 1), today=TODAY, records=_records())
    assert [f["ref"] for f in found] == ["RW788288", "RW789171"]
    a = found[0]
    assert a["original"] == 1156.20 and a["revised_at"] == "2026-10-09" and a["reason"] == "NO DOWELS ORDERED" and a["project"] == "RP7502"
    assert a["statement_matched"] is True and a["statement_says"].startswith("matched at $966.08 on the 09/30/2026 statement")
    assert a["approval"] == "not approved" and a["flags"] == []
    b = found[1]
    assert b["original"] is None and b["statement_matched"] is False and "still $726.31" in b["statement_says"]   # no change-log rows = the original is unknown, never guessed


def test_pickles_are_flagged(monkeypatch):
    con = _con(monkeypatch)
    _bill(con, "1", "COWTOWN REDI MIX CONCRETE", "401816", "2026-09-24", 600.0, "RP7428 - 4418 BALDWIN ST\n\nREVISED - TAX EXEMPT JOB")
    _change(con, "1", "edited", "2026-10-02T12:44:49-07:00", 650.0, 600.0, "amount changed,lines changed")
    _bill(con, "2", "COWTOWN REDI MIX CONCRETE", "401816-R", "2026-10-05", 600.0, "RP7428 - re-issued", created="2026-10-06T10:00:00-07:00")   # the pickle
    _bill(con, "3", "Preferred Materials LLC", "INV41671", "2026-08-18", 1871.64, "ADJUSTED - RP7447\n\nREVISED - PRICE PER", balance=0.0)
    _bill(con, "4", "RCI READY CABLE", "RW1", "2026-09-03", 10.0, "RP7622 - x\n\nREVISED - gone", deleted=1)
    found = {f["ref"]: f for f in rv.find(con, since=dt.date(2026, 8, 1), today=TODAY, records={})}
    assert found["401816"]["flags"] == ["entered twice? bill 401816-R (2026-10-05, $600.00) entered 2026-10-06"]
    assert found["INV41671"]["flags"] == ["paid before the vendor re-issued"]
    assert found["RW1"]["flags"] == ["bill gone from QuickBooks"] and found["RW1"]["deleted"]
    assert found["401816"]["statement_says"] == "no statement record for this vendor"


def test_plan_moves_stages_forward_only_and_leaves_the_clerks_columns():
    base = {"bill_id": "1", "vendor": "V", "ref": "A", "bill_date": "2026-09-01", "project": "", "reason": "x", "total": 10.0, "original": 12.0,
            "open": 10.0, "approval": "approved", "revised_at": "2026-10-01", "amount_edits": [{"at": "2026-10-01T07:00:00-07:00", "before": 12.0, "after": 10.0}],
            "deleted": False, "statement_says": "not on the 09/30/2026 statement", "statement_matched": None, "flags": [], "flags_text": "", "checked_at": "x"}
    # new -> created at To send
    p = rv.plan([base], [])
    assert p[0]["action"] == "create" and p[0]["stage_to"] == "To send"
    # the clerk moved it on; nothing changed -> keep
    prior = [{"page_id": "pg", "key": "1", "stage": "Sent to vendor", "sent_on": "2026-10-03", "statement_says": base["statement_says"], "flags": "",
              "approval": "approved", "pm_reason": "x", "qbo_total": 10.0, "open_balance": 10.0, "vendor_original": 12.0}]
    p = rv.plan([base], prior)
    assert p[0]["action"] == "keep" and p[0]["stage_to"] is None
    # the statement now agrees -> Done, forward from Sent to vendor
    done = dict(base, statement_matched=True, statement_says="matched at $10.00 on the 10/31/2026 statement")
    p = rv.plan([done], prior)
    assert p[0]["action"] == "update" and p[0]["stage_to"] == "Done - statement agrees"
    # the PM changed the amount AGAIN after she sent it -> flagged, stage untouched
    again = dict(base, amount_edits=base["amount_edits"] + [{"at": "2026-10-08T07:00:00-07:00", "before": 10.0, "after": 9.0}], total=9.0)
    p = rv.plan([again], prior)
    assert p[0]["action"] == "update" and p[0]["stage_to"] is None and "amount changed again" in p[0]["row"]["flags_text"]
    # a bill gone from QuickBooks -> Dropped; a Done row never goes back
    p = rv.plan([dict(base, deleted=True, flags=["bill gone from QuickBooks"])], prior)
    assert p[0]["stage_to"] == "Dropped"
    p = rv.plan([dict(base, deleted=True)], [dict(prior[0], stage="Done - statement agrees")])
    assert p[0]["stage_to"] is None


def test_props_never_touch_the_clerks_columns_and_parse_round_trips():
    row = {"bill_id": "77", "vendor": "V", "ref": "A1", "project": "RP1", "reason": "why", "approval": "approved", "statement_says": "s",
           "flags_text": "f", "total": 10.0, "original": 12.0, "open": 0.0, "bill_date": "2026-09-01", "revised_at": "2026-10-01", "checked_at": "2026-10-12T10:00"}
    props = rb.props_auto(row)
    assert not (set(props) & set(rb.CLERK_DATES)) and "Stage" not in props and "Notes" not in props
    assert props["QBO bill"]["url"].endswith("txnId=77") and props["Open balance"]["number"] == 0.0 and props["Bill date"]["date"]["start"] == "2026-09-01"
    page = {"id": "pg", "url": "https://n/pg", "last_edited_time": "2026-10-12T10:00:00.000Z", "properties": {
        **{k: {"type": "rich_text", "rich_text": [{"plain_text": v["rich_text"][0]["text"]["content"]}]} for k, v in props.items() if "rich_text" in v},
        "Name": {"type": "title", "title": [{"plain_text": "V · A1"}]}, "Stage": {"type": "select", "select": {"name": "Sent to vendor"}},
        "Sent on": {"type": "date", "date": {"start": "2026-10-03"}}, "QBO total": {"type": "number", "number": 10.0}, "QBO bill": {"type": "url", "url": "https://q"}}}
    r = rb.parse_page(page)
    assert r["key"] == "77" and r["stage"] == "Sent to vendor" and r["sent_on"] == "2026-10-03" and r["qbo_total"] == 10.0 and r["bill"] == "A1" and r["pm_reason"] == "why"
    st = rb.standing(r, TODAY)
    assert st["stage"] == "Sent to vendor" and st["days"] == 9 and st["text"] == "waiting on the vendor's reply - 9 days" and not st["final"]
    assert rb.standing({"stage": "Done - statement agrees"}, TODAY)["final"]
    assert rb.standing({"stage": "", "revised_in_qbo": "2026-10-10"}, TODAY)["text"] == "the clerk has not told the vendor yet - 2 days"
