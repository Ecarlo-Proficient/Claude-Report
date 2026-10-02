"""The Notion Vendor Statements board (one page per vendor): merge rules, month
closing, status, grouping, Teams digest. No network - the Board runs against a
fake client; the folder renames run against a temp directory."""
import datetime as dt
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "statement-reconciler"))

import notion_board as nb  # noqa: E402
from shared import teams_notify  # noqa: E402

TODAY = "10/02/2026"
EMPTY = {"months": {}, "cleared": {}, "history": {}}


def _it(kind, ref, amount=100.0, bill_id=""):
    return {"kind": kind, "ref": ref, "date": "2026-08-14", "amount": amount,
            "bill_id": bill_id, "job": "RP6612"}


def _run(items, tieout=True, parsed=True, stmt_date="2026-08-26"):
    return {"items": items, "tieout": tieout, "parsed": parsed, "stmt_date": stmt_date}


def _prior(month, items, done_tick=False, tieout=True):
    return {"months": {month: {"items": {nb._item_key(k, r): {"kind": k, "ref": r, "line": r,
                                                               "url": "", "checked": c}
                                         for k, r, c in items},
                               "done_tick": done_tick, "tieout": tieout}},
            "cleared": {}, "history": {}}


# ── merge ──

def test_first_run_counts_not_entered_apart_from_followups():
    m = nb.merge({"08-2026": _run([_it("enter", "A1"), _it("approve", "B2"), _it("lag", "C3")])},
                 EMPTY, today=TODAY)
    assert m["not_entered"] == 1 and m["followups"] == 1        # unknown kind ignored
    assert nb.status(m, "2026-08-26", dt.date(2026, 10, 2)) == ("Not entered",
                                                               "1 not entered (oldest 08-2026)")


def test_only_followups_is_still_all_entered():
    m = nb.merge({"08-2026": _run([_it("approve", "B2"), _it("mismatch", "B3")])}, EMPTY, today=TODAY)
    assert m["not_entered"] == 0 and m["followups"] == 2 and "08-2026" in m["months"]
    assert nb.status(m, "2026-08-26", dt.date(2026, 10, 2)) == ("All entered",
                                                               "All entered as of 10/02/2026")


def test_clean_month_closes_itself():
    m = nb.merge({"07-2026": _run([])}, EMPTY, today=TODAY)
    assert m["closing"] == {"07-2026": "clean"} and m["months"] == {}
    assert m["history"]["07-2026"] == {"on": TODAY, "how": "clean"}


def test_tie_out_failed_month_never_closes_and_marks_unreadable():
    m = nb.merge({"08-2026": _run([], tieout=False)}, EMPTY, today=TODAY)
    assert m["closing"] == {} and m["unreadable"] == ["08-2026"]
    assert nb.status(m, "2026-08-26")[0] == "Unreadable"


def test_ticked_month_done_closes_with_what_was_left():
    prior = _prior("08-2026", [("approve", "B2", False)], done_tick=True)
    m = nb.merge({"08-2026": _run([_it("approve", "B2")])}, prior, today=TODAY)
    assert m["closing"] == {"08-2026": "ticked done · 1 left open"}


def test_folder_done_by_hand_moves_month_to_history():
    prior = _prior("06-2026", [("enter", "A1", False)])
    m = nb.merge({}, prior, folder_done=frozenset({"06-2026"}), today=TODAY)
    assert m["closing"] == {"06-2026": "folder marked DONE"} and m["not_entered"] == 0


def test_failed_rename_keeps_the_month_open():
    m = nb.merge({"07-2026": _run([])}, EMPTY, keep_open=frozenset({"07-2026"}), today=TODAY)
    assert m["closing"] == {} and "07-2026" in m["months"]


def test_gone_from_qbo_moves_to_cleared_with_its_month():
    prior = _prior("08-2026", [("enter", "A1", False), ("approve", "B2", False)])
    m = nb.merge({"08-2026": _run([_it("approve", "B2")])}, prior, today=TODAY)
    c = m["cleared"][nb._item_key("enter", "A1")]
    assert c["month"] == "08-2026" and c["on"] == TODAY and m["not_entered"] == 0


def test_ticked_but_still_missing_is_unticked():
    prior = _prior("08-2026", [("enter", "A1", True)])
    m = nb.merge({"08-2026": _run([_it("enter", "A1")])}, prior, today=TODAY)
    (it, checked, note), = m["months"]["08-2026"]["rows"]
    assert checked is False and "still open in QBO" in note and m["not_entered"] == 1


def test_unverifiable_tick_is_kept():
    prior = _prior("08-2026", [("checkqbo", "Z9", True)])
    m = nb.merge({"08-2026": _run([_it("checkqbo", "Z9")])}, prior, today=TODAY)
    (it, checked, note), = m["months"]["08-2026"]["rows"]
    assert checked is True and m["followups"] == 0


def test_unchecked_kind_is_carried_not_cleared():
    prior = _prior("08-2026", [("print", "P1", False)])
    m = nb.merge({"08-2026": _run([])}, prior, frozenset({"print"}), today=TODAY)
    assert m["cleared"] == {} and m["followups"] == 1 and m["closing"] == {}


def test_a_bill_repeated_on_newer_statements_is_listed_once_under_the_newest():
    runs = {"06-2026": _run([_it("enter", "A1")]), "08-2026": _run([_it("enter", "A1")])}
    m = nb.merge(runs, EMPTY, today=TODAY)
    assert m["not_entered"] == 1 and m["closing"] == {"06-2026": "clean"}
    assert [r[0]["ref"] for r in m["months"]["08-2026"]["rows"]] == ["A1"]


def test_months_not_run_are_carried_untouched():
    prior = _prior("07-2026", [("enter", "A1", True), ("approve", "B2", False)])
    m = nb.merge({"08-2026": _run([_it("enter", "C3")])}, prior, today=TODAY)
    rows = m["months"]["07-2026"]["rows"]
    assert [(r[0]["ref"], r[1]) for r in rows] == [("A1", True), ("B2", False)]
    assert m["not_entered"] == 1 and m["cleared"] == {}


def test_unreadable_run_month_keeps_last_known_items():
    prior = _prior("08-2026", [("enter", "A1", False)])
    m = nb.merge({"08-2026": _run([], tieout=False, parsed=False)}, prior, today=TODAY)
    assert m["not_entered"] == 1 and m["unreadable"] == ["08-2026"] and m["cleared"] == {}


def test_reappeared_item_is_flagged_back_open():
    prior = dict(EMPTY, cleared={nb._item_key("enter", "A1"): {"kind": "enter", "month": "07-2026",
                                                                "ref": "A1", "on": "09/01/2026"}})
    m = nb.merge({"08-2026": _run([_it("enter", "A1")])}, prior, today=TODAY)
    assert m["months"]["08-2026"]["rows"][0][2] == "back open" and m["cleared"] == {}


# ── status ──

def test_no_statement_in_60_days_goes_gray():
    m = nb.merge({}, EMPTY, today=TODAY)
    assert nb.status(m, "2026-07-01", dt.date(2026, 10, 2)) == ("No statement",
                                                               "No statement since 07/01/2026")
    assert nb.status(m, "2026-08-05", dt.date(2026, 10, 2))[0] == "All entered"


def test_dates_never_year_first():
    assert nb._us("2026-08-14") == "08/14/2026"
    assert "08/14/2026" in nb._item_line(_it("enter", "A1"))


# ── grouping ──

def test_group_results_one_record_per_vendor():
    base = {"action": "filed", "vendor": "RCI READY CABLE", "vendor_folder": "RCI Ready Cable",
            "month": "08-2026", "tieout": True, "items": [_it("enter", "A1")],
            "unchecked_kinds": ["print"], "stmt_date": "2026-08-26", "vendor_id": "55"}
    recs = nb.group_results([
        base, dict(base, tieout=False, unchecked_kinds=None),
        dict(base, action="preview", month="07-2026"),
        {"action": "unreadable", "vendor_folder": "RCI Ready Cable", "month": "06-2026"},
        {"action": "skipped_done", "vendor": "X"}])
    assert len(recs) == 1
    r = recs[0]
    assert r["key"] == "RCI READY CABLE" and r["vendor"] == "RCI READY CABLE" and r["vendor_id"] == "55"
    assert set(r["months"]) == {"06-2026", "07-2026", "08-2026"}
    assert r["months"]["08-2026"]["tieout"] is False and len(r["months"]["08-2026"]["items"]) == 2
    assert r["months"]["06-2026"] == {"items": [], "tieout": False, "parsed": False, "stmt_date": ""}
    assert r["unchecked_kinds"] == {"print"}


def test_unreadable_file_beside_a_clean_one_keeps_the_month_parsed():
    ok = {"action": "filed", "vendor": "SUNBELT", "vendor_folder": "Sunbelt Rentals",
          "month": "09-2026", "items": []}
    recs = nb.group_results([ok, {"action": "unreadable", "vendor_folder": "Sunbelt Rentals",
                                  "month": "09-2026"}])
    mm = recs[0]["months"]["09-2026"]
    assert mm["parsed"] is True and mm["tieout"] is False


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


def _rec(months):
    return {"key": "RCI READY CABLE", "vendor": "RCI READY CABLE", "vendor_id": "55",
            "folder": "RCI Ready Cable", "months": months, "unchecked_kinds": {"print"}}


def test_vendor_page_round_trip(tmp_path):
    vdir = tmp_path / "RCI Ready Cable"
    for mo in ("07-2026", "08-2026"):
        (vdir / mo).mkdir(parents=True)
    fake = FakeNotion()
    board = nb.Board("ds", client=fake)
    logs = []
    r1 = board.sync(_rec({"07-2026": _run([], stmt_date="2026-07-28"),
                          "08-2026": _run([_it("enter", "A1", bill_id="77"), _it("approve", "B2")])}),
                    tmp_path, log=logs.append)
    assert r1["status"] == "Not entered" and r1["closed"] == {"07-2026": "clean"}
    assert (vdir / "07-2026 DONE").is_dir() and not (vdir / "07-2026").exists()
    pid = next(iter(fake.pages))
    props = fake.pages[pid]["properties"]
    assert props["Standing"]["rich_text"][0]["annotations"]["color"] == "red"
    assert props["Last statement"]["date"]["start"] == "2026-08-26"
    heads = [nb._plain(b) for b in fake.block_children(pid)]
    assert heads[0].startswith("Pay-run check: 1 bill not entered") and "08-2026: 1" in heads[0]
    assert heads[1:] == ["08-2026 · 2 open", "History (1)"]
    month = fake.block_children(fake.kids[pid][1])
    assert nb._plain(month[0]).startswith("Month done")
    assert [nb._plain(b) for b in month[1:3]] == ["To enter in QBO (1)", "A1 · 08/14/2026 · $100.00 · RP6612"]
    assert nb._first_link(month[2]).endswith("txnId=77")

    # the clerk enters A1, writes a note; the preview must not write anything
    fake.append_children(pid, [{"type": "paragraph", "paragraph": {"rich_text": [
        {"type": "text", "text": {"content": "called the PM 09/29"}}]}}])
    ro = ReadOnlyNotion()
    ro.__dict__.update(fake.__dict__)
    p = nb.Board("ds", client=ro).preview(_rec({"08-2026": _run([_it("approve", "B2")])}), tmp_path)
    assert p["action"] == "UPDATE" and p["status"] == "All entered" and p["followups"] == 1

    r2 = board.sync(_rec({"08-2026": _run([_it("approve", "B2")])}), tmp_path, log=logs.append)
    assert r2["status"] == "All entered" and r2["not_entered"] == 0
    texts = [nb._plain(b) for b in fake.block_children(pid)]
    assert "called the PM 09/29" in texts and any(t.startswith("Cleared (1)") for t in texts)
    assert fake.pages[pid]["properties"]["Standing"]["rich_text"][0]["annotations"]["color"] == "green"

    # the clerk ticks Month done with the approval still open -> folder filed, month to History
    fake.tick(pid, "Month done")
    r3 = board.sync(_rec({"08-2026": _run([_it("approve", "B2")])}), tmp_path, log=logs.append)
    assert r3["closed"] == {"08-2026": "ticked done · 1 left open"}
    assert (vdir / "08-2026 DONE").is_dir()
    hist = [b for b in fake.block_children(pid) if nb._plain(b).startswith("History")][0]
    assert {nb._plain(c).split(" · ")[0] for c in fake.block_children(hist["id"])} == {"07-2026", "08-2026"}
