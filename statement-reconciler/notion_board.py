"""notion_board.py - the Notion "Vendor Statements" board (tool-local).

ONE Notion page per VENDOR (owner 09/30/2026). The row answers the one question
the clerk has before a pay run: is every bill on the vendor's statements entered
in QBO? Approvals, amount / tax fixes and printing are follow-ups that can wait
until after the payment, so they never change the row's colour.

Row (properties): Status (colour) · Standing ("All entered as of 10/02/2026",
"30 not entered (oldest 06-2026)", ...) · Not entered · Follow-ups · Last
statement · Last checked · QBO vendor · Key (the vendor's folder name).

  All entered   every statement bill is in QBO            -> safe to pay
  Not entered   the statement lists bills QBO lacks       -> enter them first
  Unreadable    a statement did not tie out / parse       -> check the PDF
  No statement  none in 60+ days                          -> ask the vendor

Page body, rewritten each run (anything else the clerk writes is left alone):
  Pay-run check callout · one heading per OPEN month, newest first (a "Month
  done" tick, then To enter in QBO, then the follow-ups, one to-do per bill) ·
  Cleared (what got fixed, dated) · History (months closed, dated).

Merge rules, per bill (key kind|ref, one line per vendor - a running-balance
statement repeats unpaid bills, so a bill sits under the NEWEST statement that
shows it):
  * gone from QBO's list                    -> Cleared (dated)
  * a verifiable item ticked but still open -> unticked, "still open in QBO"
  * unverifiable (approval pending) ticks   -> kept
A month closes (folder renamed '<MM-YYYY> DONE', month -> History) when it ties
out with nothing left, when the clerk ticks "Month done", or when someone
already renamed its folder DONE.

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
MAX_CHILDREN = 100                        # Notion: 100 children per block
STALE_DAYS = 60                           # no statement this long -> "No statement"

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
_LABEL_TO_KIND = {v[1]: k for k, v in KINDS.items()}
H_CLEARED, H_HISTORY = "Cleared", "History"
CHECK_PREFIX = "Pay-run check"
DONE_TICK = "Month done"
UNREADABLE = "statement unreadable - check the PDF"

# Status -> colour of the Standing text and the pay-run callout
STATUS_COLOR = {"All entered": "green", "Not entered": "red",
                "Unreadable": "purple", "No statement": "gray"}

_MONTH_RE = re.compile(r"^(\d\d-\d{4})\b")


# ─────────────────────────── helpers ───────────────────────────

def _us(iso: str) -> str:
    """mm/dd/yyyy from an ISO date - never year-first (owner rule)."""
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})", iso or "")
    return f"{m.group(2)}/{m.group(3)}/{m.group(1)}" if m else (iso or "")


def _today() -> str:
    return dt.date.today().strftime("%m/%d/%Y")


def _month_sort(month: str) -> Tuple[str, str]:
    """'08-2026' -> ('2026', '08') so months sort in time order."""
    mm, _, yyyy = month.partition("-")
    return (yyyy, mm)


def _us_sort(us: str) -> str:
    """'10/02/2026' -> '2026/10/02' for sorting only."""
    return us[-4:] + us[:5] if len(us) == 10 else us


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


def _first_link(block: dict) -> str:
    body = block.get(block.get("type", ""), {}) or {}
    for r in body.get("rich_text", []):
        url = r.get("href") or ((r.get("text") or {}).get("link") or {}).get("url")
        if url:
            return url
    return ""


def _item_key(kind: str, ref: str) -> str:
    return f"{kind}|{(ref or '').strip().upper()}"


def _row_key(row: tuple) -> str:
    return _item_key(row[0]["kind"], row[0].get("ref", ""))


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
    url = it.get("url") or (QBO_BILL_URL.format(bill_id=it["bill_id"]) if it.get("bill_id") else "")
    rich = [_rt(ref, url), _rt(_item_line(it)[len(ref):])]
    if note:
        rich.append(_rt(f"  {note}", italic=True, color="gray"))
    return {"type": "to_do", "to_do": {"rich_text": rich, "checked": checked}}


def _para(text: str, **ann) -> dict:
    return {"type": "paragraph", "paragraph": {"rich_text": [_rt(text, **ann)]}}


def _toggle(title: str, children: list) -> dict:
    return {"type": "heading_3", "heading_3": {
        "rich_text": [_rt(title)], "is_toggleable": True, "children": children}}


def _cap(kids: list) -> list:
    if len(kids) <= MAX_CHILDREN:
        return kids
    more = len(kids) - MAX_CHILDREN + 1
    return kids[:MAX_CHILDREN - 1] + [_para(f"… {more} more - see the reconciliation Excel",
                                            italic=True, color="gray")]


def _heading_kind(text: str) -> Optional[str]:
    """Which follow-up section a bold line inside a month belongs to."""
    base = re.sub(r"\s*\(\d+\)\s*$", "", text).strip()
    for k, v in KINDS.items():
        if base == v[0]:
            return k
    return None


def _top_heading(text: str) -> Optional[str]:
    """Which managed top-level heading this is: 'MM-YYYY', '_cleared', '_history'."""
    m = _MONTH_RE.match(text)
    if m:
        return m.group(1)
    base = re.sub(r"\s*\(\d+\)\s*$", "", text).strip()
    return {H_CLEARED: "_cleared", H_HISTORY: "_history"}.get(base)


def _parse_ref(text: str) -> str:
    return text.split(" · ", 1)[0].strip()


def _stale(last_iso: str, today: dt.date) -> bool:
    try:
        return (today - dt.date.fromisoformat(last_iso[:10])).days > STALE_DAYS
    except ValueError:
        return False


def _open_enter(rows: list) -> int:
    return sum(1 for it, chk, _n in rows if it["kind"] == "enter" and not chk)


# ─────────────────────────── the pure merge ───────────────────────────

def merge(run_months: Dict[str, dict], prior: dict, unchecked_kinds: frozenset = frozenset(),
          folder_done: frozenset = frozenset(), keep_open: frozenset = frozenset(),
          today: Optional[str] = None) -> dict:
    """Merge this run's reconcile with the vendor page as last written. Pure.

    run_months: {month: {items, tieout, parsed}} - the months reconciled this run
                (parsed False = no statement in that month could be read).
    prior:      the page read back - {months: {month: {items: {key: {kind, ref,
                line, url, checked}}, done_tick, tieout}}, cleared, history}.
    unchecked_kinds: kinds this run did not evaluate (print status off) - carried.
    folder_done: months whose folder someone already renamed DONE.
    keep_open:  months to keep open even if they would close (a failed rename).
    Returns {months, cleared, history, closing, not_entered, followups,
             unreadable, oldest_open}; months = {month: {rows, done_tick,
             tieout, ran}} with rows = [(item, checked, note)]."""
    today = today or _today()
    pm = prior.get("months", {})
    cleared = dict(prior.get("cleared", {}))
    history = dict(prior.get("history", {}))
    fresh = {mo for mo, r in run_months.items() if r.get("parsed", True)}
    months: Dict[str, dict] = {}

    # 1. carry the months this run did not re-derive (not run, or unreadable)
    for mo, st in pm.items():
        if mo in fresh:
            continue
        months[mo] = {"rows": [(dict(it), bool(it.get("checked")), "")
                               for it in st.get("items", {}).values()],
                      "done_tick": st.get("done_tick", False),
                      "tieout": (False if mo in run_months else st.get("tieout", True)),
                      "ran": mo in run_months}
    for mo in run_months:
        if mo not in fresh and mo not in months:
            months[mo] = {"rows": [], "done_tick": False, "tieout": False, "ran": True}

    # 2. fresh items: each bill once, under the newest month that shows it
    prior_checked: Dict[str, dict] = {}
    for mo in fresh & set(pm):
        for key, it in pm[mo].get("items", {}).items():
            prior_checked[key] = dict(it, month=mo)
    seen: Dict[str, str] = {}
    for mo in sorted(fresh, key=_month_sort, reverse=True):
        rows = []
        for it in run_months[mo].get("items", []):
            key = _item_key(it["kind"], it.get("ref", ""))
            if it["kind"] not in KINDS or key in seen:
                continue
            seen[key] = mo
            was = prior_checked.get(key)
            if was is None:
                rows.append((it, False, "back open" if key in cleared else ""))
            elif was.get("checked") and KINDS[it["kind"]][2]:
                rows.append((it, False, f"still open in QBO {today}"))
            else:
                rows.append((it, bool(was.get("checked")), ""))
            cleared.pop(key, None)
        months[mo] = {"rows": rows, "done_tick": pm.get(mo, {}).get("done_tick", False),
                      "tieout": bool(run_months[mo].get("tieout", True)), "ran": True}
    for mo, st in months.items():             # a fresh bill supersedes a carried copy
        if mo not in fresh:
            st["rows"] = [r for r in st["rows"] if _row_key(r) not in seen]

    # 3. prior bills of re-derived months that QBO no longer shows
    for key, prev in prior_checked.items():
        if key in seen:
            continue
        if prev["kind"] in unchecked_kinds:
            months[prev["month"]]["rows"].append((prev, bool(prev.get("checked")), ""))
            seen[key] = prev["month"]
        elif key not in cleared:
            cleared[key] = {"kind": prev["kind"], "month": prev["month"], "ref": prev.get("ref", ""),
                            "line": prev.get("line") or prev.get("ref", ""), "on": today}

    # 4. which months close
    closing: Dict[str, str] = {}
    for mo, st in months.items():
        if mo in keep_open:
            continue
        left = sum(1 for _it, chk, _n in st["rows"] if not chk)
        if mo in folder_done:
            closing[mo] = "folder marked DONE"
        elif st["done_tick"]:
            closing[mo] = "ticked done" + (f" · {left} left open" if left else "")
        elif st["ran"] and st["tieout"] and not st["rows"]:
            closing[mo] = "clean"
    for mo, how in closing.items():
        months.pop(mo)
        history[mo] = {"on": today, "how": how}

    enter_months = [mo for mo, st in months.items() if _open_enter(st["rows"])]
    return {"months": months, "cleared": cleared, "history": history, "closing": closing,
            "not_entered": sum(_open_enter(st["rows"]) for st in months.values()),
            "followups": sum(1 for st in months.values() for it, chk, _n in st["rows"]
                             if it["kind"] != "enter" and not chk),
            "unreadable": sorted((mo for mo, st in months.items() if not st["tieout"]),
                                 key=_month_sort),
            "oldest_open": min(enter_months, key=_month_sort) if enter_months else ""}


def status(m: dict, last_statement: str, today: Optional[dt.date] = None) -> Tuple[str, str]:
    """(Status, Standing text) for the row. Only 'is everything entered' colours it."""
    today = today or dt.date.today()
    if m["unreadable"]:
        extra = f" · {m['not_entered']} not entered" if m["not_entered"] else ""
        return "Unreadable", f"{m['unreadable'][-1]} statement unreadable{extra}"
    if m["not_entered"]:
        return "Not entered", f"{m['not_entered']} not entered (oldest {m['oldest_open']})"
    if last_statement and _stale(last_statement, today):
        return "No statement", f"No statement since {_us(last_statement)}"
    return "All entered", f"All entered as of {today.strftime('%m/%d/%Y')}"


def _month_folder_done(root: Optional[Path], vend: str, month: str) -> bool:
    if root is None or not vend:
        return False
    vdir = root / vend
    return vdir.is_dir() and any(
        d.is_dir() and d.name.split(maxsplit=1)[0] == month and re.search(r"\bDONE\b", d.name, re.I)
        for d in vdir.iterdir())


def _rename_month_done(root: Optional[Path], vend: str, month: str, log=print) -> bool:
    """'<Vendor>/<MM-YYYY>' -> '<MM-YYYY> DONE'. False (month stays open) on any trouble."""
    if root is None or not vend:
        return True                            # no share on this run - nothing to rename
    src = root / vend / month
    if not src.is_dir():
        return True
    dst = root / vend / f"{month} DONE"
    try:
        src.rename(dst)
        log(f"  filed {vend}/{month} -> {dst.name}")
        return True
    except OSError as e:
        log(f"  could not rename {vend}/{month} ({e}) - month stays open")
        return False


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
        """Our blocks on the page and the state inside them."""
        out: dict = {"managed": [], "months": {}, "cleared": {}, "history": {}}
        for b in self.nc.block_children(page_id):
            text = _plain(b)
            if b.get("type") == "callout" and text.startswith(CHECK_PREFIX):
                out["managed"].append(b["id"])
                continue
            if b.get("type") != "heading_3":
                continue
            which = _top_heading(text)
            if which is None:
                continue
            out["managed"].append(b["id"])
            kids = list(self.nc.block_children(b["id"])) if b.get("has_children") else []
            if which == "_cleared":
                for c in kids:
                    m = re.match(r"(\w+) (\d\d-\d{4}):\s*(.*)$", _plain(c))
                    if not m or m.group(1) not in _LABEL_TO_KIND:
                        continue
                    kind = _LABEL_TO_KIND[m.group(1)]
                    d = re.search(r"cleared (\d\d/\d\d/\d{4})", m.group(3))
                    line = m.group(3).split("  cleared")[0].strip()
                    out["cleared"][_item_key(kind, _parse_ref(line))] = {
                        "kind": kind, "month": m.group(2), "ref": _parse_ref(line),
                        "line": line, "on": d.group(1) if d else ""}
            elif which == "_history":
                for c in kids:
                    m = re.match(r"(\d\d-\d{4}) · done (\S+)(?: · (.*))?$", _plain(c))
                    if m:
                        out["history"][m.group(1)] = {"on": m.group(2), "how": m.group(3) or ""}
            else:
                st = {"items": {}, "done_tick": False, "tieout": UNREADABLE not in text}
                kind = None
                for c in kids:
                    ct = _plain(c)
                    if c.get("type") == "paragraph":
                        kind = _heading_kind(ct) or kind
                    elif c.get("type") == "to_do":
                        chk = bool(c.get("to_do", {}).get("checked"))
                        if ct.startswith(DONE_TICK):
                            st["done_tick"] = chk
                        elif kind:
                            line = ct.split("  ", 1)[0].strip()
                            ref = _parse_ref(line)
                            st["items"][_item_key(kind, ref)] = {
                                "kind": kind, "ref": ref, "line": line,
                                "url": _first_link(c), "checked": chk}
                out["months"][which] = st
        return out

    # ── page body write ──
    @staticmethod
    def _check_text(m: dict, st_name: str, standing: str) -> str:
        if st_name == "All entered":
            return f"{CHECK_PREFIX}: every statement bill is entered in QBO - {standing.lower()}."
        if st_name == "Not entered":
            by_mo = ", ".join(f"{mo}: {_open_enter(st['rows'])}" for mo, st in
                              sorted(m["months"].items(), key=lambda kv: _month_sort(kv[0]),
                                     reverse=True) if _open_enter(st["rows"]))
            n = m["not_entered"]
            return (f"{CHECK_PREFIX}: {n} bill{'s' if n != 1 else ''} not entered in QBO - "
                    f"enter them before the pay run ({by_mo}).")
        if st_name == "Unreadable":
            return f"{CHECK_PREFIX}: can't confirm - {standing}. Check the statement PDF by hand."
        return f"{CHECK_PREFIX}: can't confirm - {standing[0].lower()}{standing[1:]}. Ask the vendor for one."

    @classmethod
    def _body_blocks(cls, m: dict, st_name: str, standing: str) -> list:
        blocks = [{"type": "callout", "callout": {
            "rich_text": [_rt(cls._check_text(m, st_name, standing))],
            "color": f"{STATUS_COLOR.get(st_name, 'gray')}_background",
            "icon": {"type": "emoji", "emoji": "✅" if st_name == "All entered" else "⚠️"}}}]
        for mo, st in sorted(m["months"].items(), key=lambda kv: _month_sort(kv[0]), reverse=True):
            left = sum(1 for _it, chk, _n in st["rows"] if not chk)
            title = f"{mo} · {left} open" + ("" if st["tieout"] else f" · {UNREADABLE}")
            kids = [{"type": "to_do", "to_do": {"checked": False, "rich_text": [
                _rt(DONE_TICK, bold=True),
                _rt(" - tick when finished; the next run files the folder DONE",
                    italic=True, color="gray")]}}]
            for kind in KINDS:
                rows = [r for r in st["rows"] if r[0]["kind"] == kind]
                if not rows:
                    continue
                n_open = sum(1 for _it, chk, _n in rows if not chk)
                kids.append(_para(f"{KINDS[kind][0]} ({n_open})", bold=True))
                kids += [_todo(it, chk, note) for it, chk, note in rows]
            blocks.append(_toggle(title, _cap(kids)))
        if m["cleared"]:
            kids = []
            for c in sorted(m["cleared"].values(), key=lambda c: _us_sort(c.get("on", "")),
                            reverse=True):
                label = KINDS.get(c["kind"], ("", c["kind"]))[1]
                kids.append({"type": "to_do", "to_do": {"checked": True, "rich_text": [
                    _rt(f"{label} {c.get('month', '')}: {c.get('line') or c['ref']}",
                        strikethrough=True, color="gray"),
                    _rt(f"  cleared {c.get('on') or _today()}", italic=True, color="gray")]}})
            blocks.append(_toggle(f"{H_CLEARED} ({len(m['cleared'])})", kids[:MAX_CHILDREN]))
        if m["history"]:
            kids = [_para(f"{mo} · done {h['on']}" + (f" · {h['how']}" if h.get("how") else ""),
                          color="gray")
                    for mo, h in sorted(m["history"].items(), key=lambda kv: _month_sort(kv[0]),
                                        reverse=True)]
            blocks.append(_toggle(f"{H_HISTORY} ({len(m['history'])})", kids[:MAX_CHILDREN]))
        return blocks

    @staticmethod
    def _props(rec: dict, m: dict, st_name: str, standing: str, last_statement: str) -> dict:
        props = {
            "Name": {"title": [_rt(rec["vendor"])]},
            "Key": {"rich_text": [_rt(rec["key"])]},
            "Status": {"select": {"name": st_name}},
            "Standing": {"rich_text": [_rt(standing, color=STATUS_COLOR.get(st_name, "default"),
                                           bold=st_name != "All entered")]},
            "Not entered": {"number": m["not_entered"]},
            "Follow-ups": {"number": m["followups"]},
            "Last checked": {"date": {"start": dt.datetime.now().astimezone().isoformat(timespec="minutes")}},
        }
        if last_statement:
            props["Last statement"] = {"date": {"start": last_statement}}
        if rec.get("vendor_id"):
            props["QBO vendor"] = {"url": QBO_VENDOR_URL.format(vendor_id=rec["vendor_id"])}
        return props

    # ── one vendor ──
    def _plan(self, rec: dict, root: Optional[Path]) -> tuple:
        """Read the page (if any) and merge - reads only."""
        page = self.nc.query_by_property(self.ds, "Key", "rich_text", rec["key"])
        prior = self._read_body(page["id"]) if page else {"managed": [], "months": {},
                                                          "cleared": {}, "history": {}}
        folder_done = frozenset(mo for mo in set(prior["months"]) | set(rec["months"])
                                if _month_folder_done(root, rec.get("folder", ""), mo))
        m = merge(rec["months"], prior, frozenset(rec.get("unchecked_kinds", ())), folder_done)
        prev_last = ((((page or {}).get("properties", {}).get("Last statement") or {})
                      .get("date") or {}).get("start") or "")
        last = max([prev_last] + [r.get("stmt_date", "") for r in rec["months"].values()])
        st_name, standing = status(m, last)
        return page, prior, m, st_name, standing, last, folder_done

    def preview(self, rec: dict, root: Optional[Path] = None) -> dict:
        """What sync() would do for one vendor, writing nothing (--dry-run)."""
        page, _prior, m, st_name, standing, _last, folder_done = self._plan(rec, root)
        return {"name": rec["vendor"], "action": "UPDATE" if page else "NEW PAGE",
                "status": st_name, "standing": standing, "not_entered": m["not_entered"],
                "followups": m["followups"], "open_months": sorted(m["months"], key=_month_sort),
                "closing": {mo: how for mo, how in m["closing"].items() if mo not in folder_done}}

    def sync(self, rec: dict, root: Optional[Path] = None, log=print) -> dict:
        """Create or rewrite the page for one vendor, and rename the folders of the
        months that close. `rec` = {key, vendor, vendor_id, folder, months:
        {month: {items, tieout, parsed, stmt_date}}, unchecked_kinds}."""
        page, prior, m, st_name, standing, last, folder_done = self._plan(rec, root)
        failed = {mo for mo in m["closing"] if mo not in folder_done
                  and not _rename_month_done(root, rec.get("folder", ""), mo, log)}
        if failed:                             # a folder we could not rename stays open
            m = merge(rec["months"], prior, frozenset(rec.get("unchecked_kinds", ())),
                      folder_done, frozenset(failed))
            st_name, standing = status(m, last)
        props = self._props(rec, m, st_name, standing, last)
        blocks = self._body_blocks(m, st_name, standing)
        if page is None:
            page = self.nc.create_page(self.ds, props, children=blocks)
        else:
            self.nc.update_page(page["id"], props)
            for bid in prior["managed"]:
                self.nc.delete_block(bid)
            self.nc.append_children(page["id"], blocks)
        return {"name": rec["vendor"], "status": st_name, "standing": standing,
                "not_entered": m["not_entered"], "followups": m["followups"],
                "url": page.get("url", ""), "closed": dict(m["closing"])}


def group_results(results: List[dict]) -> List[dict]:
    """Per-file reconcile results -> one record per VENDOR (keyed by its folder on
    the share). Two statements in one month merge; tie-out must hold for both. A
    file that could not be read at all ('unreadable', vendor known from its
    folder) makes its month unreadable."""
    groups: Dict[str, dict] = {}
    for res in results:
        if res.get("action") not in ("filed", "held", "preview", "unreadable"):
            continue
        folder = res.get("vendor_folder") or res.get("vendor") or ""
        mo = res.get("month", "")
        if not folder or not mo:
            continue
        key = folder.upper()
        g = groups.setdefault(key, {"key": key, "vendor": folder, "vendor_id": "",
                                    "folder": folder, "months": {}, "unchecked_kinds": set()})
        mm = g["months"].setdefault(mo, {"items": [], "tieout": True, "parsed": False,
                                         "stmt_date": ""})
        if res["action"] == "unreadable":
            mm["tieout"] = False
            continue
        if res.get("vendor"):
            g["vendor"] = res["vendor"]
        g["vendor_id"] = g["vendor_id"] or res.get("vendor_id", "")
        mm["parsed"] = True
        mm["items"] += res.get("items") or []
        mm["tieout"] = mm["tieout"] and bool(res.get("tieout", True))
        mm["stmt_date"] = max(mm["stmt_date"], res.get("stmt_date", ""))
        g["unchecked_kinds"].update(res.get("unchecked_kinds") or ())
    return list(groups.values())
