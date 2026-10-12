"""The REVISED-bill follow-up in the Bill Tracker (bill-tracker/revised_bills.py + shared/revised_sheet.py,
10/12/2026): finding revised bills in a stand-in QuickBooks copy, the vendor's original total from the
change log, what the statement record says, the pickles flagged, the sheet written and read back with
the clerk's typed columns carried forward, and the stage rules (forward only). No network, no share."""
import datetime as dt
import json
import sqlite3
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "bill-tracker"))

import revised_bills as rv  # noqa: E402
from shared import qbo_mirror, revised_sheet as rs, statement_record as sr, xlsx_verify  # noqa: E402

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


def _found(**over):
    base = {"bill_id": "1", "vendor": "V", "ref": "A", "bill_date": "2026-09-01", "project": "RP1", "reason": "x", "total": 10.0, "original": 12.0,
            "open": 10.0, "approval": "approved", "revised_at": "2026-10-01", "amount_edits": [{"at": "2026-10-01T07:00:00-07:00", "before": 12.0, "after": 10.0}],
            "deleted": False, "statement_says": "not on the 09/30/2026 statement", "statement_matched": None, "flags": []}
    return dict(base, **over)


def test_plan_moves_stages_forward_only_and_carries_the_clerks_columns():
    p = rv.plan([_found()], [], TODAY)
    assert p[0]["stage"] == "To send" and p[0]["moved"] and p[0]["standing"]["text"] == "the clerk has not told the vendor yet - 11 days"
    # the clerk typed Sent to vendor + a date and a note: carried as typed, no move
    prior = [{"key": "1", "stage": "sent to vendor", "sent_on": "2026-10-03", "replied_on": "", "received_on": "", "entered_on": "", "notes": "called twice"}]
    p = rv.plan([_found()], prior, TODAY)
    assert p[0]["stage"] == "Sent to vendor" and not p[0]["moved"] and p[0]["sent_on"] == "2026-10-03" and p[0]["notes"] == "called twice"
    assert p[0]["standing"]["text"] == "waiting on the vendor's reply - 9 days"
    # the statement now agrees -> Done, forward from Sent to vendor
    p = rv.plan([_found(statement_matched=True, statement_says="matched at $10.00 on the 10/31/2026 statement")], prior, TODAY)
    assert p[0]["stage"] == "Done - statement agrees" and p[0]["moved"]
    # the PM changed the amount AGAIN after she sent it -> flagged, stage untouched
    again = _found(amount_edits=[{"at": "2026-10-01T07:00:00-07:00", "before": 12.0, "after": 10.0}, {"at": "2026-10-08T07:00:00-07:00", "before": 10.0, "after": 9.0}], total=9.0)
    p = rv.plan([again], prior, TODAY)
    assert p[0]["stage"] == "Sent to vendor" and "amount changed again" in p[0]["flags_text"]
    # a bill gone from QuickBooks -> Dropped; a Done row never goes back
    assert rv.plan([_found(deleted=True, flags=["bill gone from QuickBooks"])], prior, TODAY)[0]["stage"] == "Dropped"
    assert rv.plan([_found(deleted=True)], [dict(prior[0], stage="Done - statement agrees")], TODAY)[0]["stage"] == "Done - statement agrees"
    # a row no longer found but still mid-process is kept as typed; a finished one is let go
    p = rv.plan([], [dict(prior[0], vendor="V", bill="A", bill_date="2026-09-01"), {"key": "2", "stage": "Dropped", "vendor": "W", "bill": "B"}], TODAY)
    assert [r["bill_id"] for r in p] == ["1"] and p[0]["flags_text"].startswith("not in the last 120 days")


def test_sheet_round_trip_keeps_what_the_clerk_typed(tmp_path):
    from openpyxl import Workbook, load_workbook
    path = tmp_path / "Bill Tracker.xlsx"
    wb = Workbook(); wb.active.title = "Bills"; wb.active["A1"] = "x"
    rows = rv.plan([_found(), _found(bill_id="2", ref="B", vendor="W", flags=["paid before the vendor re-issued"], open=0.0)], [], TODAY)
    assert rv.write_sheet(wb, rows) == 2
    wb.save(path)
    xlsx_verify.assert_clean(path)                    # the table ref, the dropdown, the formulas all pass Excel's rules
    back = rs.read_sheet(path)
    assert [r["key"] for r in back] == ["1", "2"] and back[0]["stage"] == "To send" and back[0]["vendor_billed"] == 12.0 and back[0]["qbo_now"] == 10.0
    assert back[0]["bill_date"] == "2026-09-01" and back[0]["revised_in_qbo"] == "2026-10-01" and back[1]["flags"] == "paid before the vendor re-issued"
    # the clerk types in Excel: a stage from the dropdown, a date, a note
    wb2 = load_workbook(path); ws = wb2[rs.SHEET]
    hdr = [c.value for c in ws[1]]
    ws.cell(row=2, column=hdr.index("Stage") + 1, value="Vendor replied")
    ws.cell(row=2, column=hdr.index("Sent on") + 1, value=dt.date(2026, 10, 3))
    ws.cell(row=2, column=hdr.index("Replied on") + 1, value="10/09/2026")
    ws.cell(row=2, column=hdr.index("Notes") + 1, value="they will re-issue")
    wb2.save(path)
    prior = rv.prior_rows(path)
    assert prior[0]["stage"] == "Vendor replied" and prior[0]["sent_on"] == "2026-10-03" and prior[0]["replied_on"] == "2026-10-09"
    # the next run rebuilds the sheet: her columns come back, the automation's refresh
    rows = rv.plan([_found(statement_says="still $12.00 on the 10/31/2026 statement (QBO $10.00)", statement_matched=False)], prior, TODAY)
    wb3 = load_workbook(path); rv.write_sheet(wb3, rows); wb3.save(path)
    xlsx_verify.assert_clean(path)
    back = {r["key"]: r for r in rs.read_sheet(path)}
    assert set(back) == {"1", "2"}                    # bill 2 fell out of the scan but is mid-process: kept as typed, flagged
    assert back["2"]["flags"].startswith("not in the last 120 days") and back["2"]["stage"] == "To send"
    b1 = back["1"]
    assert b1["stage"] == "Vendor replied" and b1["sent_on"] == "2026-10-03" and b1["replied_on"] == "2026-10-09"
    assert b1["notes"] == "they will re-issue" and b1["statement_says"].startswith("still $12.00") and b1["where_it_stands"] == "waiting on the re-issued bill - 3 days"
    # the sheet keeps its place and its dropdown
    wb4 = load_workbook(path)
    assert wb4.sheetnames.index(rs.SHEET) == 1 and wb4[rs.SHEET].data_validations.dataValidation[0].formula1 == f'"{rs.STAGE_LIST}"'


def test_stage_words_match_loosely():
    assert rs.stage_of(" sent to vendor ") == "Sent to vendor" and rs.stage_of("DONE") == "Done - statement agrees"
    assert rs.stage_of("received") == "Revised bill received" and rs.stage_of("") == "" and rs.stage_of("whatever") == ""
    assert rs.standing({"stage": "Done - statement agrees"}, TODAY)["final"] and len(rs.STAGE_LIST) < 255 and "," not in "".join(rs.STAGES)
