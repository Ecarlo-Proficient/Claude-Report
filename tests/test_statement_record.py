"""The vendor statement record the ledger reads (shared/statement_record, 10/09/2026): built from a
reconcile result + the marked statements, written beside the state file with the pages as PNGs,
read back by vendor name however it is spelled, and every served path kept inside the vendor's
folder. No network, no share - a temp folder stands in for Vendor Statements."""
import json
import sys
from dataclasses import dataclass
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "statement-reconciler"))

import pytest  # noqa: E402

from shared import statement_record as sr  # noqa: E402


@dataclass
class L:                       # a StmtLine stand-in
    date: str
    ref: str
    amount: float


class FakeResult:              # statement_markup.Result as mark() fills it
    def __init__(self, lines, bands, pages, unread=(), not_found=()):
        self.lines, self.bands, self.pages = lines, bands, pages
        self.unread, self.not_found = list(unread), list(not_found)
        self.located, self.total = len(bands), len(lines)


class FakeImage:
    size = (1275, 1650)

    def save(self, path, format="PNG", optimize=True):
        Path(path).write_bytes(b"\x89PNG fake " + str(self.size).encode())


def _rows():
    return [
        {"category": "MATCHED", "bucket": "matched", "stmt_ref": "INV-100", "qbo_ref": "INV-100", "stmt_date": "2026-09-01",
         "qbo_date": "2026-09-01", "stmt_amount": 500.0, "qbo_amount": 500.0, "qbo_bill_id": "901", "job": "RP6612", "note": ""},
        {"category": "MISSING_IN_QBO", "bucket": "enter", "stmt_ref": "inv 101", "qbo_ref": "", "stmt_date": "2026-09-10",
         "qbo_date": "", "stmt_amount": 250.0, "qbo_amount": 0.0, "qbo_bill_id": "", "job": "", "note": "called the vendor"},
        {"category": "MISSING_ON_STATEMENT", "bucket": "unlisted", "stmt_ref": "", "qbo_ref": "INV-099", "stmt_date": "",
         "qbo_date": "2026-08-20", "stmt_amount": 0.0, "qbo_amount": 75.0, "qbo_bill_id": "899", "job": "CP861", "note": ""},
    ]


def _result():
    return {"vendor": "Burnco Texas LLC", "vendor_folder": "Burnco", "vendor_id": "42", "as_of": "2026-09-30",
            "checked": "10/09/2026 02:14 PM", "gap": 0.0, "amount_due": 750.0, "merged_total": 750.0, "matched": 1,
            "tieout": True, "clean": False, "open_items": {"To enter (missing in QBO)": 1},
            "statements": [{"as_of": "2026-09-30", "label": "Current", "kind": "open list", "file": "Statement 09-30.pdf"}],
            "_record": {"rows": _rows(), "files": ["Current/Statement 09-30.pdf"], "excel": "Current/Reconciliation - Burnco - as of 09-30-2026.xlsx"}}


def _markups():
    lines = [L("2026-09-01", "INV-100", 500.0), L("2026-09-10", "INV 101", 250.0), L("2026-09-15", "", -100.0)]
    bands = [{"line": 0, "page": 0, "top": 400.0, "bottom": 418.0, "bucket": "matched"},
             {"line": 1, "page": 1, "top": 120.0, "bottom": 138.0, "bucket": "enter"},
             {"line": 2, "page": 1, "top": 300.0, "bottom": 318.0, "bucket": "other"}]
    unread = [{"page": 1, "top": 500.0, "bottom": 520.0, "text": "09/20/2026 INV-102 99.00"}]
    return [("Statement 09-30.pdf", FakeResult(lines, bands, [FakeImage(), FakeImage()], unread))]


# ── build ──

def test_build_lines_carry_bucket_band_and_qbo_bill_and_the_qbo_only_bill_has_no_page():
    rec = sr.build(_result(), status="Not entered", standing="1 not entered as of 09/30/2026", notion_url="https://n/1", markups=_markups())
    assert rec["status"] == "Not entered" and rec["as_of"] == "2026-09-30" and rec["notion_url"] == "https://n/1"
    assert len(rec["pages"]) == 2 and rec["pages"][1]["source"] == "Statement 09-30.pdf"
    by_ref = {ln["ref"]: ln for ln in rec["lines"]}
    m = by_ref["INV-100"]
    assert m["bucket"] == "matched" and m["page"] == 0 and m["top"] == 400.0 and m["bill_id"] == "901" and m["job"] == "RP6612"
    e = by_ref["INV 101"]                       # 'inv 101' on the row meets 'INV 101' on the statement: the key folds case + spacing
    assert e["bucket"] == "enter" and e["page"] == 1 and e["note"] == "called the vendor" and e["bill_id"] == ""
    u = by_ref["INV-099"]
    assert u["bucket"] == "unlisted" and u["page"] is None and u["bill_id"] == "899" and u["amount"] == 75.0
    assert by_ref[""]["bucket"] == "other" and by_ref[""]["amount"] == -100.0
    assert rec["unread"] == [{"page": 1, "top": 500.0, "bottom": 520.0, "text": "09/20/2026 INV-102 99.00"}]
    assert rec["counts"] == {"matched": 1, "enter": 1, "other": 1, "unlisted": 1, "unread": 1}
    assert rec["located"] == 3 and rec["total"] == 3
    assert rec["files"] == ["Current/Statement 09-30.pdf"] and rec["excel"].startswith("Current/Reconciliation")


def test_build_without_markups_still_lists_every_statement_row_pageless():
    rec = sr.build(_result(), status="Not entered", standing="x", markups=[])
    refs = sorted(ln["ref"] for ln in rec["lines"])
    assert refs == ["INV-099", "INV-100", "inv 101"] and all(ln["page"] is None for ln in rec["lines"])
    assert rec["pages"] == [] and rec["total"] == 0


def test_pages_offset_across_two_statements():
    m1 = _markups()
    m2 = [("Letter 10-01.pdf", FakeResult([L("2026-10-01", "INV-200", 10.0)], [{"line": 0, "page": 0, "top": 50.0, "bottom": 60.0, "bucket": "matched"}], [FakeImage()]))]
    rec = sr.build(_result(), status="All entered", standing="x", markups=m1 + m2)
    assert len(rec["pages"]) == 3 and rec["pages"][2]["source"] == "Letter 10-01.pdf"
    assert next(ln for ln in rec["lines"] if ln["ref"] == "INV-200")["page"] == 2


# ── write + read ──

def test_write_saves_pages_and_json_and_read_all_finds_the_vendor_by_any_spelling(tmp_path):
    vdir = tmp_path / "Burnco"
    (vdir / "Current").mkdir(parents=True)
    (tmp_path / "Inbox").mkdir()                      # the dump folder is not a vendor
    (tmp_path / "Cintas").mkdir()                     # a folder the reconciler has not written a record for
    rec = sr.build(_result(), status="Not entered", standing="1 not entered as of 09/30/2026", markups=_markups())
    out = sr.write(vdir, rec, [FakeImage(), FakeImage()])
    assert out == vdir / sr.RECORD_FILE
    saved = json.loads(out.read_text())
    assert saved["version"] == 1 and saved["pages"][0]["file"] == ".marked/page-01.png" and saved["pages"][0]["width"] == 1275
    assert sorted(p.name for p in (vdir / ".marked").iterdir()) == ["page-01.png", "page-02.png"]
    allv = sr.read_all(tmp_path)
    assert allv["mounted"] is True
    for spelling in ("BURNCO TEXAS LLC", "Burnco Texas, LLC", "burnco"):
        s = allv["vendors"][sr.key(spelling)]
        assert s["status"] == "Not entered" and s["pages"] == 2 and s["lines"] == 4 and s["as_of"] == "2026-09-30"
    assert allv["vendors"][sr.key("Cintas")]["status"] == "No record"
    assert sr.key("Inbox") not in allv["vendors"]
    full = sr.find("BURNCO TEXAS LLC", tmp_path)
    assert full and full["folder"] == "Burnco" and full["_dir"] == str(vdir)
    assert sr.find("Nobody", tmp_path) is None


def test_rewrite_drops_stale_pages(tmp_path):
    vdir = tmp_path / "V"
    rec = sr.build(_result(), status="All entered", standing="x", markups=_markups())
    sr.write(vdir, rec, [FakeImage(), FakeImage()])
    rec2 = sr.build(_result(), status="All entered", standing="x", markups=[("one.pdf", FakeResult([], [], [FakeImage()]))])
    sr.write(vdir, rec2, [FakeImage()])
    assert [p.name for p in (vdir / ".marked").iterdir()] == ["page-01.png"]
    assert len(json.loads((vdir / sr.RECORD_FILE).read_text())["pages"]) == 1


def test_served_paths_stay_inside_the_vendor_folder(tmp_path):
    vdir = tmp_path / "V"; (vdir / "Current").mkdir(parents=True)
    (vdir / "Current" / "s.pdf").write_bytes(b"%PDF")
    (tmp_path / "secret.txt").write_text("no")
    rec = sr.build(_result(), status="All entered", standing="x", markups=_markups())
    rec["files"] = ["Current/s.pdf"]
    sr.write(vdir, rec, [FakeImage(), FakeImage()])
    full = sr.find("Burnco Texas LLC", tmp_path)
    assert sr.resolve_file(full, "Current/s.pdf") == (vdir / "Current" / "s.pdf").resolve()
    assert sr.resolve_file(full, ".marked/page-01.png").name == "page-01.png"
    assert sr.resolve_file(full, "../secret.txt") is None
    assert sr.resolve_file(full, "/etc/hosts") is None
    assert sr.resolve_file(full, "Current/missing.pdf") is None


def test_unmounted_root_reads_as_not_mounted(tmp_path):
    out = sr.read_all(tmp_path / "nope")
    assert out == {"mounted": False, "root": str(tmp_path / "nope"), "vendors": {}}
    assert sr.find("x", tmp_path / "nope") is None


def test_root_honours_the_override(monkeypatch, tmp_path):
    monkeypatch.setenv("ACB_STATEMENTS_ROOT", str(tmp_path))
    assert sr.root() == tmp_path
    monkeypatch.delenv("ACB_STATEMENTS_ROOT")
    assert sr.root().name == "Vendor Statements" and sr.root().parent.name == "Accounts Payable"


# ── the markup now reports where every band sits ──

def test_markup_result_carries_bands_and_unread_positions():
    import statement_markup as sm
    assert hasattr(sm.Result(), "bands") and hasattr(sm.Result(), "unread") and hasattr(sm.Result(), "lines")
    for k in sr.BUCKETS:
        if k not in ("unlisted", "unread"):
            assert k in sm.BUCKETS, k            # the ledger's legend speaks the markup's words
            assert sr.BUCKETS[k][1] == sm.BUCKETS[k][1], k


def test_reconciler_publish_returns_the_boards_writes_and_records_are_written_live(tmp_path, monkeypatch):
    """_write_records: a filed vendor gets its record + pages, an unreadable one only turns its
    status, a dry run writes nothing. notion_board is driven with no network (status computed)."""
    import statement_reconciler as r
    vdir = tmp_path / "Burnco"; (vdir / "Current").mkdir(parents=True)
    res = _result(); res.update({"action": "filed", "_vdir": str(vdir), "_markups": _markups(),
                                 "items": [{"kind": "enter", "ref": "INV 101", "date": "2026-09-10", "amount": 250.0, "bill_id": "", "job": ""}]})
    args = type("A", (), {"dry_run": True})()
    r._write_records([res], [], args)
    assert not (vdir / sr.RECORD_FILE).exists()
    args.dry_run = False
    r._write_records([res], [{"name": "Burnco Texas LLC", "status": "Not entered", "standing": "1 not entered as of 09/30/2026", "url": "https://n/p"}], args)
    rec = json.loads((vdir / sr.RECORD_FILE).read_text())
    assert rec["status"] == "Not entered" and rec["notion_url"] == "https://n/p" and rec["board_counts"]["enter"] == 1
    assert len(rec["pages"]) == 2 and (vdir / ".marked" / "page-02.png").exists()
    # the board had nothing to say (no Notion): the same status is computed from the items
    r._write_records([res], [], args)
    assert json.loads((vdir / sr.RECORD_FILE).read_text())["status"] == "Not entered"
    # unreadable: the lines and pages stay, the status turns
    r._write_records([{"action": "unreadable", "vendor_folder": "Burnco", "vendor": "Burnco Texas LLC", "vendor_id": "42",
                       "as_of": "2026-10-05", "_vdir": str(vdir)}], [], args)
    rec = json.loads((vdir / sr.RECORD_FILE).read_text())
    assert rec["status"] == "Unreadable" and rec["standing"] == "Statement as of 10/05/2026 unreadable"
    assert len(rec["pages"]) == 2 and len(rec["lines"]) == 4 and (vdir / ".marked" / "page-01.png").exists()


@pytest.mark.parametrize("a,b", [("BURNCO TEXAS LLC", "Burnco Texas, LLC"), ("R.C.I.", "RCI"), ("Cowtown  Redi Mix", "cowtown redi-mix")])
def test_key_folds_case_spacing_and_punctuation(a, b):
    assert sr.key(a) == sr.key(b)
