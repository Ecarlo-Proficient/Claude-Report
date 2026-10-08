"""The Notion Vendor Statements board (one running page per vendor, 10/07/2026):
merge rules, status, every bucket counted, grouping, Teams digest, a fake-Notion
round trip incl. a pre-10/07 page with month headings. No network."""
import datetime as dt
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "statement-reconciler"))

import notion_board as nb  # noqa: E402
from shared import teams_notify  # noqa: E402

TODAY = "10/07/2026"
EMPTY = {"items": {}, "cleared": {}}


def _it(kind, ref, amount=100.0, bill_id="", note=""):
    return {"kind": kind, "ref": ref, "date": "2026-08-14", "amount": amount,
            "bill_id": bill_id, "job": "RP6612", "clerk_note": note}


def _run(items, tieout=True, parsed=True):
    return {"items": items, "tieout": tieout, "parsed": parsed}


def _prior(items):
    return {"items": {nb._item_key(k, r): {"kind": k, "ref": r, "line": r, "url": "", "checked": c}
                      for k, r, c in items}, "cleared": {}}


# ── merge ──

def test_every_bucket_is_counted_and_unknown_kinds_ignored():
    m = nb.merge(_run([_it("enter", "A1"), _it("approve", "B2"), _it("lag", "C3"),
                       _it("unlisted", "D4"), _it("bogus", "E5")]), EMPTY, today=TODAY)
    assert m["not_entered"] == 1 and m["followups"] == 3
    assert m["counts"]["lag"] == 1 and m["counts"]["unlisted"] == 1 and m["counts"]["tax"] == 0
    assert nb.status(m, "2026-09-30", 0, dt.date(2026, 10, 7)) == (
        "Not entered", "1 not entered as of 09/30/2026")


def test_only_followups_is_all_entered_as_of_the_statement_date():
    m = nb.merge(_run([_it("approve", "B2")]), EMPTY, today=TODAY)
    assert nb.status(m, "2026-09-30", 0, dt.date(2026, 10, 7)) == (
        "All entered", "All entered as of 09/30/2026")


def test_statements_that_disagree_say_so_in_the_standing():
    m = nb.merge(_run([]), EMPTY, today=TODAY)
    assert nb.status(m, "2026-10-02", 12.34, dt.date(2026, 10, 7))[1] == (
        "All entered as of 10/02/2026 · statements disagree by $12.34")


def test_unreadable_carries_the_last_known_items():
    prior = _prior([("enter", "A1", False)])
    m = nb.merge(_run([], parsed=False), prior, today=TODAY)
    assert [r[0]["ref"] for r in m["rows"]] == ["A1"] and m["tieout"] is False
    assert nb.status(m, "2026-09-30", 0, dt.date(2026, 10, 7))[0] == "Unreadable"


def test_gone_from_the_run_moves_to_cleared():
    m = nb.merge(_run([]), _prior([("enter", "A1", False)]), today=TODAY)
    assert m["rows"] == [] and m["cleared"]["enter|A1"]["on"] == TODAY


def test_ticked_but_still_missing_is_unticked():
    m = nb.merge(_run([_it("enter", "A1")]), _prior([("enter", "A1", True)]), today=TODAY)
    assert m["rows"][0][1] is False and m["rows"][0][2] == f"still open in QBO {TODAY}"


def test_unverifiable_tick_is_kept_and_unchecked_kind_carried():
    prior = _prior([("checkqbo", "B2", True), ("print", "P1", False)])
    m = nb.merge(_run([_it("checkqbo", "B2")]), prior, frozenset({"print"}), today=TODAY)
    assert {(r[0]["ref"], r[1]) for r in m["rows"]} == {("B2", True), ("P1", False)}


def test_reappeared_item_is_flagged_back_open():
    prior = {"items": {}, "cleared": {"enter|A1": {"kind": "enter", "ref": "A1", "on": "09/01/2026"}}}
    m = nb.merge(_run([_it("enter", "A1")]), prior, today=TODAY)
    assert m["rows"][0][2] == "back open" and "enter|A1" not in m["cleared"]


def test_no_statement_in_60_days_goes_gray():
    m = nb.merge(_run([]), EMPTY, today=TODAY)
    assert nb.status(m, "2026-07-01", 0, dt.date(2026, 10, 7)) == (
        "No statement", "No statement since 07/01/2026")


def test_dates_never_year_first():
    assert nb._us("2026-08-14") == "08/14/2026"
    assert "08/14/2026" in nb._item_line(_it("enter", "A1"))


# ── grouping ──

def test_group_results_one_record_per_vendor():
    base = {"action": "filed", "vendor": "RCI READY CABLE", "vendor_folder": "RCI Ready Cable",
            "tieout": True, "items": [_it("enter", "A1")], "unchecked_kinds": ["print"],
            "as_of": "2026-09-30", "vendor_id": "55", "matched": 240, "gap": 0.0}
    recs = nb.group_results([base, {"action": "skipped", "vendor": "X"}])
    assert len(recs) == 1
    r = recs[0]
    assert r["key"] == "RCI READY CABLE" and r["vendor_id"] == "55" and r["matched"] == 240
    assert r["parsed"] is True and r["unchecked_kinds"] == {"print"}


def test_unreadable_vendor_record():
    recs = nb.group_results([{"action": "unreadable", "vendor_folder": "Sunbelt Rentals",
                              "as_of": "2026-10-01"}])
    assert recs[0]["parsed"] is False and recs[0]["tieout"] is False


# ── Teams ──

def test_teams_digest_lists_only_vendors_not_ready(monkeypatch):
    sent = {}
    monkeypatch.setattr(teams_notify, "post", lambda url, payload: sent.update(p=payload) or True)
    ok = teams_notify.post_statement_digest("https://hook", [
        {"name": "RCI", "status": "Not entered", "standing": "30 not entered (oldest 06-2026)",
         "url": "https://n/1"},
        {"name": "CINTAS", "status": "All entered", "standing": "All entered as of 10/02/2026",
         "url": ""}], "https://board")
    card = sent["p"]["attachments"][0]["content"]
    assert ok and card["body"][0]["text"] == "Vendor statements · 1 not ready to pay"
    assert "[RCI](https://n/1) · 30 not entered (oldest 06-2026)" in card["body"][1]["text"]
    assert "CINTAS" not in card["body"][1]["text"] and "✅ 1 vendor(s)" in card["body"][1]["text"]
    assert card["actions"][0]["url"] == "https://board"


# ── fake Notion round trip ──

class FakeNotion:
    """Just enough of NotionClient to run Board.sync repeatedly: pages + a block
    tree whose rich_text gets plain_text the way the real API returns it."""

    def __init__(self):
        self.pages, self.blocks, self.kids, self.n = {}, {}, {}, 0

    def _id(self):
        self.n += 1
        return f"b{self.n}"

    def _store(self, parent, children):
        out = []
        for c in children:
            c = dict(c)
            grand = (c.get(c["type"]) or {}).pop("children", None)
            for r in (c.get(c["type"]) or {}).get("rich_text", []):
                r["plain_text"] = r["text"]["content"]
            c["id"] = self._id()
            c["has_children"] = bool(grand)
            self.blocks[c["id"]] = c
            out.append(c)
            if grand:
                self._store(c["id"], grand)
        self.kids.setdefault(parent, []).extend(c["id"] for c in out)
        if parent in self.blocks and out:
            self.blocks[parent]["has_children"] = True
        return {"results": out}

    def query_by_property(self, ds, prop, kind, value):
        return next((p for p in self.pages.values()
                     if p["properties"]["Key"]["rich_text"][0]["text"]["content"] == value), None)

    def create_page(self, ds, props, children=None):
        pid = self._id()
        self.pages[pid] = {"id": pid, "url": f"https://n/{pid}", "properties": props}
        self._store(pid, children or [])
        return self.pages[pid]

    def update_page(self, pid, props):
        self.pages[pid]["properties"].update(props)

    def block_children(self, bid):
        return [self.blocks[i] for i in self.kids.get(bid, [])]

    def delete_block(self, bid):
        for lst in self.kids.values():
            if bid in lst:
                lst.remove(bid)

    def append_children(self, bid, children):
        return self._store(bid, children)

    def tick(self, pid, startswith):
        for hid in self.kids[pid]:
            for cid in self.kids.get(hid, []):
                c = self.blocks[cid]
                if c["type"] == "to_do" and nb._plain(c).startswith(startswith):
                    c["to_do"]["checked"] = True


class ReadOnlyNotion(FakeNotion):
    """Fails the test on any write - the preview must only read."""

    def create_page(self, *a, **k):
        raise AssertionError("preview wrote a page")

    def update_page(self, *a, **k):
        raise AssertionError("preview updated a page")

    def delete_block(self, *a, **k):
        raise AssertionError("preview deleted a block")

    def append_children(self, *a, **k):
        raise AssertionError("preview appended blocks")


def _rec(items, as_of="2026-09-30", **kw):
    return dict({"key": "RCI READY CABLE", "vendor": "RCI READY CABLE", "vendor_id": "55",
                 "folder": "RCI Ready Cable", "items": items, "tieout": True, "parsed": True,
                 "as_of": as_of, "unchecked_kinds": {"print"}, "matched": 10,
                 "checked_against": "1 statement as of 09/30/2026 vs QuickBooks ...",
                 "statements": [{"as_of": as_of, "label": "Current", "kind": "open list",
                                 "amount": 500.0, "name": "Statement RCI.pdf"}],
                 "changes": {"frm": "2026-08-31", "to": as_of, "new": [], "gone": [],
                             "changed": [{"ref": "C1", "date": "2026-08-04", "amount": 170.8,
                                          "was": 184.89, "where": "2100 MAYHILL"}]}}, **kw)


class SchemaNotion(FakeNotion):
    def retrieve_data_source(self, ds):
        return {"properties": {"Last checked": {"type": "date"}}}

    def _request(self, method, path, body):
        self.schema = body


def test_vendor_page_round_trip():
    fake = SchemaNotion()
    board = nb.Board("ds", client=fake)
    r1 = board.sync(_rec([_it("enter", "A1", bill_id="77", note="called RCI"), _it("tax", "T2")]),
                    log=lambda m: None)
    assert r1["status"] == "Not entered" and r1["standing"] == "1 not entered as of 09/30/2026"
    assert "Paid, vendor shows open" in fake.schema["properties"]
    assert fake.schema["properties"]["Last checked"]["date"] == {}
    pid = next(iter(fake.pages))
    props = fake.pages[pid]["properties"]
    assert props["Tax charged"]["number"] == 1 and props["Matched"]["number"] == 10
    assert props["Last statement"]["date"]["start"] == "2026-09-30"
    heads = [nb._plain(b) for b in fake.block_children(pid)]
    assert heads[0].startswith("Pay-run check: 1 bill not entered")
    assert heads[1:] == ["To do · as of 09/30/2026 · 2 open", "Changed since 08/31/2026 (1)",
                         "Statements (1)"]
    todo = fake.block_children(fake.kids[pid][1])
    assert nb._plain(todo[0]) == "To enter in QBO (1)"
    assert nb._first_link(todo[1]).endswith("txnId=77")
    assert [nb._plain(b) for b in fake.block_children(todo[1]["id"])] == ["called RCI"]   # the note, under the bill
    assert board.notes("RCI READY CABLE")["enter|A1"] == {"ref": "A1", "bill_id": "77", "note": "called RCI"}
    assert fake.pages[pid]["properties"]["Re-check"] == {"checkbox": False}

    # the clerk enters A1 and writes on the page; the preview must not write
    fake.append_children(pid, [{"type": "paragraph", "paragraph": {"rich_text": [
        {"type": "text", "text": {"content": "called the PM 09/29"}}]}}])
    ro = ReadOnlyNotion()
    ro.__dict__.update(fake.__dict__)
    p = nb.Board("ds", client=ro).preview(_rec([_it("tax", "T2")]))
    assert p["action"] == "UPDATE" and p["status"] == "All entered" and p["followups"] == 1

    r2 = board.sync(_rec([_it("tax", "T2")]), log=lambda m: None)
    assert r2["status"] == "All entered"
    texts = [nb._plain(b) for b in fake.block_children(pid)]
    assert "called the PM 09/29" in texts and any(t.startswith("Cleared (1)") for t in texts)


def test_a_page_written_with_months_carries_its_ticks():
    fake = FakeNotion()
    old = [nb._toggle("08-2026 · 1 open", [
               {"type": "to_do", "to_do": {"checked": False, "rich_text": [nb._rt("Month done")]}},
               nb._para("Not approved - chase PM (1)", bold=True),
               nb._todo(_it("approve", "B2"), True)]),
           nb._toggle("History (1)", [nb._para("07-2026 · done 09/30/2026")])]
    fake.create_page("ds", {"Key": {"rich_text": [nb._rt("RCI READY CABLE")]}}, children=old)
    fake.retrieve_data_source = lambda ds: {"properties": {}}
    fake._request = lambda *a: None
    board = nb.Board("ds", client=fake)
    board.sync(_rec([_it("approve", "B2")]), log=lambda m: None)
    pid = next(iter(fake.pages))
    heads = [nb._plain(b) for b in fake.block_children(pid)]
    assert not any(h.startswith(("08-2026", "History")) for h in heads)
    todo = fake.block_children([b for b in fake.kids[pid]
                                if nb._plain(fake.blocks[b]).startswith("To do")][0])
    assert todo[1]["to_do"]["checked"] is False       # verifiable + still open -> unticked


def test_a_note_typed_under_a_bill_on_notion_is_read_back():
    fake = FakeNotion()
    fake.retrieve_data_source = lambda ds: {"properties": {}}
    fake._request = lambda *a: None
    board = nb.Board("ds", client=fake)
    board.sync(_rec([_it("tax", "T2", bill_id="9")]), log=lambda m: None)
    pid = next(iter(fake.pages))
    todo = fake.block_children(fake.kids[pid][1])[1]
    fake.append_children(todo["id"], [nb._para("asked for a credit 10/08")])      # she types it
    assert board.notes("RCI READY CABLE")["tax|T2"]["note"] == "asked for a credit 10/08"
