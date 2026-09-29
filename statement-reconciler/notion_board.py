"""notion_board.py - the Notion "Vendor Statements" board (tool-local).

ONE Notion page per vendor-month that the reconciler keeps rewriting until the
month is done - the live record of that statement, replacing the stack of Teams
cards (owner 2026-09-29: "one card that keeps a live record of that statement
and goes through work until it's shown as done").

Page properties: Status · Open items · Statement date · First seen · Last checked
· QBO vendor · Folder · Key (Days open is a Notion formula). Page body - one
collapsible heading per kind of work, one to-do per bill (ref linked to QBO,
date, $, job):

  To enter in QBO · Not approved - chase PM · Amount mismatch · Tax charged ·
  Not printed · Approval pending? check QBO · Cleared (what got fixed, dated)

Every run re-derives the list from QBO and MERGES it with what the clerk ticked:
  * an item QBO no longer shows          -> moves to Cleared (ticked, dated)
  * a VERIFIABLE item she ticked that QBO still shows
                                          -> unticked, noted "still open in QBO"
  * an item the script cannot verify (approval pending in QBO's workflow)
                                          -> her tick is kept
Status is computed (Open / In progress / Clean / Tie-out failed); Done is the
human's call and is never overwritten. `sync_done` mirrors Done both ways with
the month folder on the Accounting share (`<MM-YYYY> DONE`). Anything else she
writes on the page is left alone - only the headings above are rewritten.

Setup: ACB_STATEMENTS_DS_ID (the data-source id) in machine.env, the database
shared with the Notion integration. No id -> the board is skipped with one line.
"""
from __future__ import annotations

import datetime as dt
import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from shared import paths
from shared.notion_client import NotionClient, NotionError

QBO_BILL_URL = "https://qbo.intuit.com/app/bill?txnId={bill_id}"
QBO_VENDOR_URL = "https://qbo.intuit.com/app/vendordetail?nameId={vendor_id}"
ITEMS_PER_SECTION = 99                    # Notion: 100 children per append

# kind -> (section heading, short label on Cleared lines,
#          verifiable: QBO can tell when it is fixed)
KINDS: Dict[str, Tuple[str, str, bool]] = {
    "enter":    ("To enter in QBO",                "Enter",   True),
    "approve":  ("Not approved - chase PM",        "Approve", True),
    "mismatch": ("Amount mismatch - fix the bill", "Amount",  True),
    "tax":      ("Tax charged - ask for a credit", "Tax",     True),
    "print":    ("Not printed",                    "Print",   True),
    "checkqbo": ("Approval pending? check QBO",    "Check",   False),
}
H_CLEARED = "Cleared"
_LABEL_TO_KIND = {v[1]: k for k, v in KINDS.items()}


# ─────────────────────────── helpers ───────────────────────────

def _us(iso: str) -> str:
    """mm/dd/yyyy from an ISO date - never year-first (owner rule)."""
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})", iso or "")
    return f"{m.group(2)}/{m.group(3)}/{m.group(1)}" if m else (iso or "")


def _today() -> str:
    return dt.date.today().strftime("%m/%d/%Y")


def _rt(text: str, url: str = "", **ann) -> dict:
    t: dict = {"type": "text", "text": {"content": text[:2000]}}
    if url:
        t["text"]["link"] = {"url": url}
    if ann:
        t["annotations"] = ann
    return t


def _plain(block: dict) -> str:
    body = block.get(block.get("type", ""), {}) or {}
    return "".join(r.get("plain_text", "") for r in body.get("rich_text", []))


def _item_key(kind: str, ref: str) -> str:
    return f"{kind}|{(ref or '').strip().upper()}"


def _item_line(it: dict) -> str:
    if it.get("line"):
        return it["line"]                     # carried forward as last shown
    bits = [it.get("ref") or "(no #)", _us(it.get("date", "")),
            f"${float(it.get('amount') or 0):,.2f}"]
    if it.get("job"):
        bits.append(it["job"])
    return " · ".join(bits)


def _todo(it: dict, checked: bool, note: str = "") -> dict:
    ref = it.get("ref") or "(no #)"
    url = QBO_BILL_URL.format(bill_id=it["bill_id"]) if it.get("bill_id") else ""
    rich = [_rt(ref, url), _rt(_item_line(it)[len(ref):])]
    if note:
        rich.append(_rt(f"  {note}", italic=True, color="gray"))
    return {"type": "to_do", "to_do": {"rich_text": rich, "checked": checked}}


def _toggle(title: str, children: list) -> dict:
    return {"type": "heading_3", "heading_3": {
        "rich_text": [_rt(title)], "is_toggleable": True, "children": children}}


def _heading_kind(text: str) -> Optional[str]:
    """Which managed section a heading's text belongs to (None = not ours)."""
    base = re.sub(r"\s*\(\d+\)\s*$", "", text).strip()
    for k, v in KINDS.items():
        if base == v[0]:
            return k
    return "_cleared" if base == H_CLEARED else None


def _parse_ref(text: str) -> str:
    return text.split(" · ", 1)[0].strip()


# ─────────────────────────── board ───────────────────────────

class Board:
    def __init__(self, ds_id: str, client: Optional[NotionClient] = None):
        self.ds = ds_id
        self.nc = client or NotionClient()
        self._url: Optional[str] = None

    @classmethod
    def from_env(cls) -> Optional["Board"]:
        ds = paths.get("ACB_STATEMENTS_DS_ID")
        return cls(ds) if ds else None

    def url(self) -> str:
        """The board's own Notion link (for the Teams digest button)."""
        if self._url is None:
            try:
                db = (self.nc.retrieve_data_source(self.ds).get("parent") or {}).get("database_id", "")
                self._url = f"https://www.notion.so/{db.replace('-', '')}" if db else ""
            except NotionError:
                self._url = ""
        return self._url

    # ── page body read-back ──
    def _read_body(self, page_id: str) -> dict:
        """Our headings on the page and the state inside them."""
        out: dict = {"managed": [], "checked": {}, "cleared": {}}
        for b in self.nc.block_children(page_id):
            if b.get("type") != "heading_3":
                continue
            kind = _heading_kind(_plain(b))
            if kind is None:
                continue
            out["managed"].append(b["id"])
            if not b.get("has_children"):
                continue
            for c in self.nc.block_children(b["id"]):
                if c.get("type") != "to_do":
                    continue
                text = _plain(c)
                if kind == "_cleared":
                    m = re.match(r"(\w+):\s*(.*)$", text)
                    if not m or m.group(1) not in _LABEL_TO_KIND:
                        continue
                    k2 = _LABEL_TO_KIND[m.group(1)]
                    d = re.search(r"cleared (\d\d/\d\d/\d{4})", text)
                    ref = _parse_ref(m.group(2))
                    out["cleared"][_item_key(k2, ref)] = {
                        "kind": k2, "ref": ref, "on": d.group(1) if d else "",
                        "line": m.group(2).split("  cleared")[0].strip()}
                else:
                    out["checked"][_item_key(kind, _parse_ref(text))] = {
                        "checked": bool(c.get("to_do", {}).get("checked")),
                        "line": text.split("  ", 1)[0].strip()}
        return out

    # ── merge ──
    @staticmethod
    def merge(items: List[dict], prior_checked: Dict[str, dict],
              prior_cleared: Dict[str, dict], unchecked_kinds: frozenset = frozenset()) -> dict:
        """Merge the fresh QBO list with last run's page. Pure (unit-tested).
        prior_checked: {item key: {checked, line}} read off the page's sections.
        unchecked_kinds: kinds this run did NOT evaluate (print status off) - their
        items are carried forward as they were, never "cleared" by absence."""
        today = _today()
        now_keys: Dict[str, dict] = {}
        sections: Dict[str, list] = {k: [] for k in KINDS}
        new = 0
        for it in items:
            key = _item_key(it["kind"], it.get("ref", ""))
            if it["kind"] not in KINDS or key in now_keys:
                continue                       # unknown kind / same bill on two statements
            now_keys[key] = it
            was = prior_checked[key]["checked"] if key in prior_checked else None
            if was is None:
                new += 1
                sections[it["kind"]].append((it, False, "back open" if key in prior_cleared else ""))
            elif was and KINDS[it["kind"]][2]:
                sections[it["kind"]].append((it, False, f"still open in QBO {today}"))
            else:
                sections[it["kind"]].append((it, bool(was), ""))
        for key, prev in prior_checked.items():
            kind, ref = key.split("|", 1)
            if kind in unchecked_kinds and key not in now_keys:
                now_keys[key] = {"kind": kind, "ref": ref, "line": prev.get("line", ref)}
                sections[kind].append((now_keys[key], bool(prev.get("checked")), ""))
        cleared = {k: v for k, v in prior_cleared.items() if k not in now_keys}
        cleared_now = 0
        for key, prev in prior_checked.items():
            if key not in now_keys and key not in cleared:
                kind, ref = key.split("|", 1)
                cleared[key] = {"kind": kind, "ref": ref, "line": prev.get("line") or ref, "on": today}
                cleared_now += 1
        open_n = sum(1 for rows in sections.values() for _it, chk, _n in rows if not chk)
        ticked = sum(1 for rows in sections.values() for _it, chk, _n in rows if chk)
        return {"sections": sections, "cleared": cleared, "new": new,
                "cleared_now": cleared_now, "open": open_n, "ticked": ticked}

    @staticmethod
    def status(prev: str, m: dict, tieout: bool) -> str:
        if prev == "Done":
            return "Done"                      # the human's call - never overwritten
        if not tieout:
            return "Tie-out failed"
        if m["open"] == 0:
            return "Clean"
        return "In progress" if (m["cleared"] or m["ticked"]) else "Open"

    # ── page body write ──
    @staticmethod
    def _body_blocks(m: dict) -> list:
        blocks = []
        for kind, rows in m["sections"].items():
            if not rows:
                continue
            kids = [_todo(it, chk, note) for it, chk, note in rows[:ITEMS_PER_SECTION]]
            if len(rows) > ITEMS_PER_SECTION:
                kids.append({"type": "paragraph", "paragraph": {"rich_text": [
                    _rt(f"… {len(rows) - ITEMS_PER_SECTION} more - see the reconciliation Excel",
                        italic=True, color="gray")]}})
            n_open = sum(1 for _it, chk, _n in rows if not chk)
            blocks.append(_toggle(f"{KINDS[kind][0]} ({n_open})", kids))
        if m["cleared"]:
            kids = []
            for c in sorted(m["cleared"].values(), key=lambda c: c.get("on", ""), reverse=True):
                label = KINDS.get(c["kind"], ("", c["kind"]))[1]
                kids.append({"type": "to_do", "to_do": {"checked": True, "rich_text": [
                    _rt(f"{label}: {c.get('line') or c['ref']}", strikethrough=True, color="gray"),
                    _rt(f"  cleared {c.get('on') or _today()}", italic=True, color="gray")]}})
            blocks.append(_toggle(f"{H_CLEARED} ({len(m['cleared'])})", kids[:ITEMS_PER_SECTION]))
        return blocks

    @staticmethod
    def _props(rec: dict, m: dict, status: str, prev_props: dict) -> dict:
        props = {
            "Name": {"title": [_rt(f"{rec['vendor']} · {rec['month']}")]},
            "Key": {"rich_text": [_rt(rec["key"])]},
            "Status": {"select": {"name": status}},
            "Open items": {"number": m["open"]},
            "Last checked": {"date": {"start": dt.datetime.now().astimezone().isoformat(timespec="minutes")}},
            "Folder": {"rich_text": [_rt(rec.get("folder", ""))]},
        }
        if rec.get("stmt_date"):
            props["Statement date"] = {"date": {"start": rec["stmt_date"]}}
        if rec.get("vendor_id"):
            props["QBO vendor"] = {"url": QBO_VENDOR_URL.format(vendor_id=rec["vendor_id"])}
        if not (prev_props.get("First seen") or {}).get("date"):
            props["First seen"] = {"date": {"start": dt.date.today().isoformat()}}
        return props

    # ── one vendor-month ──
    def sync(self, rec: dict) -> dict:
        """Create or rewrite the page for one vendor-month. `rec` = {key, vendor,
        month, items, tieout, stmt_date, vendor_id, folder, unchecked_kinds}.
        Returns {status, open, new, cleared_now, url, name}."""
        page = self.nc.query_by_property(self.ds, "Key", "rich_text", rec["key"])
        prev_props = (page or {}).get("properties", {})
        prev_status = ((prev_props.get("Status") or {}).get("select") or {}).get("name", "")
        body = self._read_body(page["id"]) if page else {"managed": [], "checked": {}, "cleared": {}}
        m = self.merge(rec.get("items", []), body["checked"], body["cleared"],
                       frozenset(rec.get("unchecked_kinds", ())))
        status = self.status(prev_status, m, rec.get("tieout", True))
        props = self._props(rec, m, status, prev_props)
        blocks = self._body_blocks(m)
        if page is None:
            page = self.nc.create_page(self.ds, props, children=blocks)
        else:
            self.nc.update_page(page["id"], props)
            for bid in body["managed"]:
                self.nc.delete_block(bid)
            if blocks:
                self.nc.append_children(page["id"], blocks)
        return {"status": status, "open": m["open"], "new": m["new"],
                "cleared_now": m["cleared_now"], "url": page.get("url", ""),
                "name": f"{rec['vendor']} · {rec['month']}"}

    # ── Done <-> month folder ──
    def sync_done(self, root: Path, apply: bool, log=print) -> List[str]:
        """Mirror Done both ways with the month folder on the share:
          Notion Done, folder open   -> rename the folder '<MM-YYYY> DONE'
          folder DONE, Notion not    -> set the page Done
        Nothing changes without apply. Returns the actions (done or proposed)."""
        acts: List[str] = []
        for page in self.nc.query_data_source(self.ds):
            props = page.get("properties", {})
            status = ((props.get("Status") or {}).get("select") or {}).get("name", "")
            folder = "".join(r.get("plain_text", "") for r in
                             (props.get("Folder") or {}).get("rich_text", []))
            if "/" not in folder:
                continue
            vend, month = folder.rsplit("/", 1)
            vdir = root / vend
            if not vdir.is_dir():
                continue
            done_dir = next((d for d in vdir.iterdir() if d.is_dir()
                             and d.name.split(maxsplit=1)[0] == month
                             and re.search(r"\bDONE\b", d.name, re.I)), None)
            open_dir = vdir / month
            if status == "Done" and done_dir is None and open_dir.is_dir():
                acts.append(f"rename folder {vend}/{month} -> {month} DONE")
                if apply:
                    open_dir.rename(vdir / f"{month} DONE")
            elif status != "Done" and done_dir is not None:
                acts.append(f"mark {vend} · {month} Done in Notion (folder is DONE)")
                if apply:
                    self.nc.update_page(page["id"], {"Status": {"select": {"name": "Done"}}})
        for a in acts:
            log(("  " if apply else "  would ") + a)
        return acts


def group_results(results: List[dict]) -> List[dict]:
    """Per-file reconcile results -> one record per vendor-month (two statements
    in a month become one page; items concatenate, tie-out must hold for both)."""
    groups: Dict[str, dict] = {}
    for res in results:
        if res.get("action") not in ("filed", "held") or not res.get("vendor"):
            continue
        key = f"{res['vendor']}|{res.get('month', '')}".upper()
        g = groups.setdefault(key, {
            "key": key, "vendor": res["vendor"], "month": res.get("month", ""),
            "items": [], "tieout": True, "stmt_date": res.get("stmt_date", ""),
            "vendor_id": res.get("vendor_id", ""), "folder": res.get("folder", ""),
            "unchecked_kinds": set()})
        g["items"] += res.get("items") or []
        g["tieout"] = g["tieout"] and bool(res.get("tieout", True))
        g["stmt_date"] = max(g["stmt_date"], res.get("stmt_date", ""))
        g["unchecked_kinds"].update(res.get("unchecked_kinds") or ())
    return list(groups.values())
