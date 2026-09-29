"""The Notion Vendor Statements board: merge rules, status, grouping, Teams digest.
No network - the Board is exercised through its pure pieces and a fake client."""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "statement-reconciler"))

import notion_board as nb  # noqa: E402
from shared import teams_notify  # noqa: E402


def _it(kind, ref, amount=100.0, bill_id=""):
    return {"kind": kind, "ref": ref, "date": "2026-08-14", "amount": amount,
            "bill_id": bill_id, "job": "RP6612"}


def test_first_run_everything_new_and_open():
    m = nb.Board.merge([_it("enter", "A1"), _it("approve", "B2"), _it("lag", "C3")], {}, {})
    assert m["open"] == 2 and m["new"] == 2 and m["cleared_now"] == 0
    assert len(m["fyi"]["lag"]) == 1
    assert nb.Board.status("", m, True) == "Open"


def test_gone_from_qbo_moves_to_cleared():
    prior = {nb._item_key("enter", "A1"): {"checked": False, "line": "A1 · 08/14/2026 · $100.00"}}
    m = nb.Board.merge([], prior, {})
    assert m["cleared_now"] == 1 and m["open"] == 0
    assert m["cleared"][nb._item_key("enter", "A1")]["line"].startswith("A1")
    assert nb.Board.status("Open", m, True) == "Clean"


def test_ticked_verifiable_but_still_in_qbo_is_unticked():
    prior = {nb._item_key("enter", "A1"): {"checked": True, "line": "A1"}}
    m = nb.Board.merge([_it("enter", "A1")], prior, {})
    (it, checked, note), = m["sections"]["enter"]
    assert checked is False and "still open in QBO" in note and m["open"] == 1


def test_ticked_unverifiable_keeps_the_tick():
    prior = {nb._item_key("checkqbo", "Z9"): {"checked": True, "line": "Z9"}}
    m = nb.Board.merge([_it("checkqbo", "Z9")], prior, {})
    (it, checked, note), = m["sections"]["checkqbo"]
    assert checked is True and m["open"] == 0 and m["ticked"] == 1
    assert nb.Board.status("Open", m, True) == "Clean"


def test_unchecked_kind_is_carried_not_cleared():
    prior = {nb._item_key("print", "P1"): {"checked": False, "line": "P1 · 08/01/2026 · $5.00"}}
    m = nb.Board.merge([], prior, {}, frozenset({"print"}))
    assert m["cleared_now"] == 0 and m["open"] == 1
    assert m["sections"]["print"][0][0]["line"].startswith("P1")


def test_reappeared_item_is_flagged_back_open():
    key = nb._item_key("enter", "A1")
    m = nb.Board.merge([_it("enter", "A1")], {}, {key: {"kind": "enter", "ref": "A1", "on": "09/01/2026"}})
    assert m["sections"]["enter"][0][2] == "back open"
    assert key not in m["cleared"]


def test_duplicate_bill_on_two_statements_counts_once():
    m = nb.Board.merge([_it("enter", "A1"), _it("enter", "a1")], {}, {})
    assert m["open"] == 1


def test_done_is_never_overwritten_and_tieout_wins():
    m = nb.Board.merge([_it("enter", "A1")], {}, {})
    assert nb.Board.status("Done", m, True) == "Done"
    assert nb.Board.status("Open", m, False) == "Tie-out failed"


def test_heading_round_trip():
    for kind, (title, *_rest) in nb.KINDS.items():
        assert nb._heading_kind(f"{title} (12)") == kind
    assert nb._heading_kind("Cleared (3)") == "_cleared"
    assert nb._heading_kind("History") == "_history"
    assert nb._heading_kind("My own notes") is None


def test_group_results_one_record_per_vendor_month():
    base = {"action": "filed", "vendor": "RCI READY CABLE", "month": "08-2026",
            "stmt_total": 100.0, "qbo_open": 500.0, "tieout": True, "items": [_it("enter", "A1")]}
    recs = nb.group_results([base, dict(base, stmt_total=50.0, qbo_open=400.0, tieout=False),
                             {"action": "skipped_done", "vendor": "X"}])
    assert len(recs) == 1
    r = recs[0]
    assert r["stmt_total"] == 150.0 and r["qbo_open"] == 500.0 and r["tieout"] is False
    assert len(r["items"]) == 2


def test_dates_never_year_first():
    assert nb._us("2026-08-14") == "08/14/2026"
    assert "08/14/2026" in nb._item_line(_it("enter", "A1"))


def test_teams_digest_is_one_compact_card(monkeypatch):
    sent = {}
    monkeypatch.setattr(teams_notify, "post", lambda url, payload: sent.update(p=payload) or True)
    ok = teams_notify.post_statement_digest("https://hook", [
        {"name": "RCI · 08-2026", "status": "Open", "open": 75, "url": "https://n/1"},
        {"name": "CINTAS · 07-2026", "status": "Clean", "open": 0, "url": ""}], "https://board")
    card = sent["p"]["attachments"][0]["content"]
    assert ok and len(card["body"]) == 2
    assert "1 open · 75 items" in card["body"][0]["text"]
    assert "[RCI · 08-2026](https://n/1) · 75 open" in card["body"][1]["text"]
    assert "CINTAS · 07-2026 is clean" in card["body"][1]["text"]
    assert card["actions"][0]["url"] == "https://board"


class FakeNotion:
    """Just enough of NotionClient to run Board.sync twice: pages + a block tree
    whose rich_text gets plain_text the way the real API returns it."""

    def __init__(self):
        self.pages, self.blocks, self.kids, self.n = {}, {}, {}, 0

    def _id(self):
        self.n += 1
        return f"b{self.n}"

    def _store(self, parent, children, after=None):
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
        lst = self.kids.setdefault(parent, [])
        i = lst.index(after) + 1 if after in lst else len(lst)
        lst[i:i] = [c["id"] for c in out]
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

    def update_block(self, bid, body):
        for part in body.values():
            for r in part.get("rich_text", []):
                r["plain_text"] = r["text"]["content"]
        self.blocks[bid].update(body)

    def append_children(self, bid, children, after=None):
        return self._store(bid, children, after)

    def tick(self, pid, ref):
        for hid in self.kids[pid]:
            for cid in self.kids.get(hid, []):
                c = self.blocks[cid]
                if c["type"] == "to_do" and nb._plain(c).startswith(ref):
                    c["to_do"]["checked"] = True


def test_sync_round_trip_through_the_page():
    fake = FakeNotion()
    board = nb.Board("ds", client=fake)
    rec = {"key": "RCI|08-2026", "vendor": "RCI", "month": "08-2026", "tieout": True,
           "stmt_total": 300.0, "qbo_open": 200.0,
           "items": [_it("enter", "A1"), _it("approve", "B2", bill_id="77"), _it("checkqbo", "C3")]}
    r1 = board.sync(rec)
    assert r1["status"] == "Open" and r1["open"] == 3
    pid = next(iter(fake.pages))
    fake.tick(pid, "A1")                       # ticked but still missing in QBO
    fake.tick(pid, "C3")                       # unverifiable - tick stands
    rec2 = dict(rec, items=[_it("enter", "A1"), _it("checkqbo", "C3")])   # B2 approved
    r2 = board.sync(rec2)
    assert r2["cleared_now"] == 1 and r2["open"] == 1 and r2["status"] == "In progress"
    props = fake.pages[pid]["properties"]
    assert props["To enter"]["number"] == 1 and props["Not approved"]["number"] == 0
    assert props["Cleared"]["number"] == 1
    heads = [nb._plain(b) for b in fake.block_children(pid)]
    assert heads[0].startswith("Statement $300.00")          # callout stays first
    assert "Cleared (1)" in heads and heads[-1] == "History"
    hist = fake.block_children(fake.kids[pid][-1])
    assert len(hist) == 2                                    # one line per run
    r3 = board.sync(dict(rec, items=[_it("checkqbo", "C3")]))
    assert r3["status"] == "Clean" and r3["open"] == 0
    assert fake.pages[pid]["properties"]["Status"]["select"]["name"] == "Clean"
