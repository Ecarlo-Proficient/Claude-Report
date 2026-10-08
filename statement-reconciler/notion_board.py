"""notion_board.py - the Notion "Vendor Statements" board (tool-local).

ONE Notion page per VENDOR, one running record (owner 10/07/2026: "all entered AS
OF the last date the statement shows"). No months: the page is built from the
vendor's CURRENT statement(s) (statement_set.py decides which), and every
bucket the reconcile finds is a column on the row.

Row (properties): Status (colour) · Standing ("All entered as of 09/30/2026",
"39 not entered as of 09/30/2026", "... · statements disagree by $12.34") ·
one count per bucket (Not entered, Not approved, Amount mismatch, Tax charged,
Paid, vendor shows open, Not on statement, Not printed, Check QBO, Matched) ·
Follow-ups · Last statement (the as-of date) · Last checked · Checked against
(which statements vs which QuickBooks pull) · QBO vendor · Key (vendor folder).

  All entered   every statement bill is in QBO            -> safe to pay
  Not entered   the statement lists bills QBO lacks       -> enter them first
  Unreadable    a statement did not tie out / parse       -> check the PDF
  No statement  none in 60+ days                          -> ask the vendor

Page body, rewritten each run (anything else the clerk writes is left alone):
  Pay-run check callout · To do (as of <date>; one section per bucket, one
  to-do per bill, the clerk's Excel note beside it) · Changed since the last
  statement · Statements (every one received, newest first: Current / DONE /
  Replaced / Duplicate) · Cleared (what got fixed, dated).

Merge rules, per bill (key kind|ref):
  * gone from this run's buckets              -> Cleared (dated)
  * a verifiable item ticked but still open   -> unticked, "still open in QBO"
  * unverifiable (approval pending) ticks     -> kept
A page written before 10/07 (month headings) is read as one bucket, so the
clerk's ticks carry over.

Setup: ACB_STATEMENTS_DS_ID (the data-source id) in machine.env, the database
shared with the Notion integration. No id -> the board is skipped with one line.
Missing bucket columns are added to the database on the first live run.
"""
from __future__ import annotations

import datetime as dt
import re
from typing import Dict, List, Optional, Tuple

from shared import paths
from shared.notion_client import NotionClient, NotionError


QBO_BILL_URL = "https://qbo.intuit.com/app/bill?txnId={bill_id}"
QBO_VENDOR_URL = "https://qbo.intuit.com/app/vendordetail?nameId={vendor_id}"
MAX_CHILDREN = 100                        # Notion: 100 children per block
STALE_DAYS = 60                           # no statement this long -> "No statement"
BUCKET = "current"                        # the one bucket merge() works on

# kind -> (section heading, short label on Cleared lines,
#          verifiable: the next run can tell when it is fixed)
KINDS: Dict[str, Tuple[str, str, bool]] = {
    "enter":    ("To enter in QBO",                                  "Enter",    True),
    "approve":  ("Not approved - chase PM",                          "Approve",  True),
    "mismatch": ("Amount mismatch - fix the bill",                   "Amount",   True),
    "tax":      ("Tax charged - ask for a credit",                   "Tax",      True),
    "lag":      ("Paid in QBO, vendor still shows it open",          "Paid",     True),
    "unlisted": ("Open in QBO, not on the statement",                "Unlisted", True),
    "print":    ("Not printed",                                      "Print",    True),
    "unread":   ("On the statement, not read by the tool - check by hand", "Unread", True),
    "checkqbo": ("Approval pending? check QBO",                      "Check",    False),
}
_LABEL_TO_KIND = {v[1]: k for k, v in KINDS.items()}
# Row columns: one count per bucket (owner 10/07: every bucket, not just two).
BUCKET_PROPS: List[Tuple[str, str]] = [
    ("Not entered", "enter"), ("Not approved", "approve"), ("Amount mismatch", "mismatch"),
    ("Tax charged", "tax"), ("Paid, vendor shows open", "lag"), ("Not on statement", "unlisted"),
    ("Not printed", "print"), ("Check QBO", "checkqbo"), ("Not read", "unread")]
PROP_DESCRIPTIONS = {
    "Last checked": "When the last run pulled QuickBooks. Each run re-reads the vendor's current "
                    "statement(s) and compares every invoice # with QuickBooks open and recently "
                    "paid bills at that moment.",
    "Checked against": "Which statement(s) and which QuickBooks pull the counts on this row come from.",
    "Last statement": "The as-of date printed on the newest current statement.",
    "Matched": "Statement invoices found in QuickBooks at the same amount.",
    "Not read": "Rows on the statement that look like a bill but the tool did not read - "
                "outlined in red on the Excel's 'Statement (marked)' sheet.",
}
H_TODO, H_CHANGES, H_STATEMENTS, H_CLEARED = "To do", "Changed since", "Statements", "Cleared"
CHECK_PREFIX = "Pay-run check"
UNREADABLE = "statement unreadable - check the PDF"

# Status -> colour of the Standing text and the pay-run callout
STATUS_COLOR = {"All entered": "green", "Not entered": "red",
                "Unreadable": "purple", "No statement": "gray"}

_MONTH_RE = re.compile(r"^(\d\d-\d{4})\b")     # a pre-10/07 page's month heading


# ─────────────────────────── helpers ───────────────────────────

def _us(iso: str) -> str:
    """mm/dd/yyyy from an ISO date - never year-first (owner rule)."""
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})", iso or "")
    return f"{m.group(2)}/{m.group(3)}/{m.group(1)}" if m else (iso or "")


def _today() -> str:
    return dt.date.today().strftime("%m/%d/%Y")


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


def _money(x: float) -> str:
    return f"-${abs(x):,.2f}" if x < 0 else f"${x:,.2f}"


def _item_line(it: dict) -> str:
    if it.get("line"):
        return it["line"]                     # carried forward as last shown
    bits = [it.get("ref") or "(no #)", _us(it.get("date", "")), _money(float(it.get("amount") or 0))]
    if it.get("job"):
        bits.append(it["job"])
    return " · ".join(bits)


def _todo(it: dict, checked: bool, note: str = "") -> dict:
    ref = it.get("ref") or "(no #)"
    url = it.get("url") or (QBO_BILL_URL.format(bill_id=it["bill_id"]) if it.get("bill_id") else "")
    rich = [_rt(ref, url), _rt(_item_line(it)[len(ref):])]
    if note:
        rich.append(_rt(f"  {note}", italic=True, color="gray"))
    if it.get("clerk_note"):                  # her note from the Excel, follows the bill
        rich.append(_rt(f"  Note: {it['clerk_note']}", italic=True, color="brown"))
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
    """Which bucket a bold line inside the To do section belongs to."""
    base = re.sub(r"\s*\(\d+\)\s*$", "", text).strip()
    for k, v in KINDS.items():
        if base == v[0]:
            return k
    return None


def _top_heading(text: str) -> Optional[str]:
    """Which managed top-level heading this is: BUCKET (To do, or a pre-10/07
    month heading), '_cleared', or '_other' (rewritten from scratch each run)."""
    if text.startswith(H_TODO) or _MONTH_RE.match(text):
        return BUCKET
    if text.startswith(H_CLEARED):
        return "_cleared"
    if text.startswith((H_CHANGES, H_STATEMENTS, "History")):
        return "_other"
    return None


def _parse_ref(text: str) -> str:
    return text.split(" · ", 1)[0].strip()


def _stale(last_iso: str, today: dt.date) -> bool:
    try:
        return (today - dt.date.fromisoformat(last_iso[:10])).days > STALE_DAYS
    except ValueError:
        return False


def _open_of(rows: list, kind: str) -> int:
    return sum(1 for it, chk, _n in rows if it["kind"] == kind and not chk)


# ─────────────────────────── the pure merge ───────────────────────────

def merge(run: dict, prior: dict, unchecked_kinds: frozenset = frozenset(),
          today: Optional[str] = None) -> dict:
    """Merge this run's reconcile with the vendor page as last written. Pure.

    run:   {items, tieout, parsed} - parsed False = no current statement could be
           read this run (the last known items are carried untouched).
    prior: the page read back - {items: {key: {kind, ref, line, url, checked}},
           cleared: {key: {...}}}.
    unchecked_kinds: kinds this run did not evaluate (print status off) - carried.
    Returns {rows, cleared, counts, not_entered, followups, tieout} with rows =
    [(item, checked, note)]."""
    today = today or _today()
    cleared = dict(prior.get("cleared", {}))
    prev = prior.get("items", {})
    if not run.get("parsed", True):
        rows = [(dict(it), bool(it.get("checked")), "") for it in prev.values()]
        return _tally(rows, cleared, tieout=False)
    rows, seen = [], set()
    for it in run.get("items", []):
        key = _item_key(it["kind"], it.get("ref", ""))
        if it["kind"] not in KINDS or key in seen:
            continue
        seen.add(key)
        was = prev.get(key)
        if was is None:
            rows.append((it, False, "back open" if key in cleared else ""))
        elif was.get("checked") and KINDS[it["kind"]][2]:
            rows.append((it, False, f"still open in QBO {today}"))
        else:
            rows.append((it, bool(was.get("checked")), ""))
        cleared.pop(key, None)
    for key, it in prev.items():
        if key in seen:
            continue
        if it["kind"] in unchecked_kinds:
            rows.append((dict(it), bool(it.get("checked")), ""))
        elif key not in cleared:
            cleared[key] = {"kind": it["kind"], "ref": it.get("ref", ""), "on": today,
                            "line": it.get("line") or it.get("ref", ""),
                            "asof": it.get("asof", "")}
    return _tally(rows, cleared, tieout=bool(run.get("tieout", True)))


def _tally(rows: list, cleared: dict, tieout: bool) -> dict:
    counts = {k: _open_of(rows, k) for k in KINDS}
    return {"rows": rows, "cleared": cleared, "counts": counts, "tieout": tieout,
            "not_entered": counts["enter"],
            "followups": sum(v for k, v in counts.items() if k != "enter")}


def status(m: dict, as_of: str, gap: float = 0.0,
           today: Optional[dt.date] = None) -> Tuple[str, str]:
    """(Status, Standing text). Only 'is everything entered' colours the row;
    a disagreement between the vendor's own statements is added to the text."""
    today = today or dt.date.today()
    when = _us(as_of)
    if not m["tieout"]:
        extra = f" · {m['not_entered']} not entered" if m["not_entered"] else ""
        st = ("Unreadable", f"Statement as of {when} unreadable{extra}")
    elif m["not_entered"]:
        st = ("Not entered", f"{m['not_entered']} not entered as of {when}")
    elif as_of and _stale(as_of, today):
        st = ("No statement", f"No statement since {when}")
    else:
        st = ("All entered", f"All entered as of {when}" if when else "All entered")
    if abs(gap) >= 0.01 and m["tieout"]:
        st = (st[0], f"{st[1]} · statements disagree by ${abs(gap):,.2f}")
    return st


# ─────────────────────────── board ───────────────────────────

class Board:
    def __init__(self, ds_id: str, client: Optional[NotionClient] = None):
        self.ds = ds_id
        self.nc = client or NotionClient()
        self._url: Optional[str] = None
        self._schema_ok = False

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

    def ensure_schema(self, log=print) -> None:
        """Add the bucket columns + 'Checked against' if the database lacks them, and
        describe the date columns. Once per run; a failure only logs."""
        if self._schema_ok:
            return
        self._schema_ok = True
        try:
            have = self.nc.retrieve_data_source(self.ds).get("properties", {})
        except Exception as e:
            log(f"  Notion columns not checked ({e})")
            return
        want: Dict[str, dict] = {name: {"number": {}} for name, _k in BUCKET_PROPS + [("Matched", "")]}
        want["Checked against"] = {"rich_text": {}}
        add = {n: v for n, v in want.items() if n not in have}
        desc = {n: d for n, d in PROP_DESCRIPTIONS.items()
                if (have.get(n) or {}).get("description") != d}
        if not add and not desc:
            return
        # Notion wants the column's type beside a description ({"date": {}, ...}).
        body = {n: dict(add.get(n) or {(have.get(n) or {}).get("type", "rich_text"): {}},
                        **({"description": desc[n]} if n in desc else {}))
                for n in set(add) | set(desc)}
        for attempt in (body, add):           # descriptions are optional - retry without them
            if not attempt:
                continue
            try:
                self.nc._request("PATCH", f"/data_sources/{self.ds}", {"properties": attempt})
                log(f"  Notion columns updated: {', '.join(sorted(attempt))}")
                return
            except Exception as e:
                log(f"  Notion columns: {e}")

    # ── page body read-back ──
    def _read_body(self, page_id: str) -> dict:
        """Our blocks on the page and the state inside them."""
        out: dict = {"managed": [], "items": {}, "cleared": {}}
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
            if which == "_other":
                continue
            kids = list(self.nc.block_children(b["id"])) if b.get("has_children") else []
            if which == "_cleared":
                for c in kids:
                    m = re.match(r"(\w+) ([\d/-]+):\s*(.*)$", _plain(c))
                    if not m or m.group(1) not in _LABEL_TO_KIND:
                        continue
                    kind = _LABEL_TO_KIND[m.group(1)]
                    d = re.search(r"cleared (\d\d/\d\d/\d{4})", m.group(3))
                    line = m.group(3).split("  cleared")[0].strip()
                    out["cleared"][_item_key(kind, _parse_ref(line))] = {
                        "kind": kind, "asof": m.group(2), "ref": _parse_ref(line),
                        "line": line, "on": d.group(1) if d else ""}
                continue
            kind = None                        # BUCKET: To do, or an old month heading
            for c in kids:
                ct = _plain(c)
                if c.get("type") == "paragraph":
                    kind = _heading_kind(ct) or kind
                elif c.get("type") == "to_do" and kind and not ct.startswith("Month done"):
                    line = ct.split("  ", 1)[0].strip()
                    ref = _parse_ref(line)
                    key = _item_key(kind, ref)
                    chk = bool(c.get("to_do", {}).get("checked"))
                    if key in out["items"]:     # same bill under two old months: a tick wins
                        out["items"][key]["checked"] = out["items"][key]["checked"] or chk
                        continue
                    out["items"][key] = {"kind": kind, "ref": ref, "line": line,
                                         "url": _first_link(c), "checked": chk}
        return out

    # ── page body write ──
    @staticmethod
    def _check_text(m: dict, st_name: str, standing: str) -> str:
        if st_name == "All entered":
            return f"{CHECK_PREFIX}: every statement bill is entered in QBO - {standing[0].lower()}{standing[1:]}."
        if st_name == "Not entered":
            n = m["not_entered"]
            return (f"{CHECK_PREFIX}: {n} bill{'s' if n != 1 else ''} not entered in QBO - "
                    f"enter them before the pay run ({standing}).")
        if st_name == "Unreadable":
            return f"{CHECK_PREFIX}: can't confirm - {standing}. Check the statement PDF by hand."
        return f"{CHECK_PREFIX}: can't confirm - {standing[0].lower()}{standing[1:]}. Ask the vendor for one."

    @staticmethod
    def _changes_block(ch: Optional[dict]) -> Optional[dict]:
        if not ch:
            return None
        tot = lambda rows: sum(r["amount"] for r in rows)
        kids = [_para(f"New on the statement: {len(ch['new'])} invoice(s) · {_money(tot(ch['new']))}"),
                _para(f"Gone (paid or credited in between): {len(ch['gone'])} invoice(s) · "
                      f"{_money(tot(ch['gone']))}")]
        if ch["changed"]:
            kids.append(_para(f"Amount changed by the vendor ({len(ch['changed'])})", bold=True))
            for r in ch["changed"]:
                bits = [r["ref"], _us(r["date"]), f"{_money(r['was'])} -> {_money(r['amount'])}"]
                if r.get("where"):
                    bits.append(r["where"])
                kids.append(_para(" · ".join(bits)))
        n = len(ch["new"]) + len(ch["gone"]) + len(ch["changed"])
        return _toggle(f"{H_CHANGES} {_us(ch.get('frm', ''))} ({n})", _cap(kids))

    @classmethod
    def _body_blocks(cls, m: dict, st_name: str, standing: str, rec: dict) -> list:
        blocks = [{"type": "callout", "callout": {
            "rich_text": [_rt(cls._check_text(m, st_name, standing))],
            "color": f"{STATUS_COLOR.get(st_name, 'gray')}_background",
            "icon": {"type": "emoji", "emoji": "✅" if st_name == "All entered" else "⚠️"}}}]
        left = sum(1 for _it, chk, _n in m["rows"] if not chk)
        title = f"{H_TODO} · as of {_us(rec.get('as_of', ''))} · {left} open"
        if not m["tieout"]:
            title += f" · {UNREADABLE}"
        kids = []
        for kind in KINDS:
            rows = [r for r in m["rows"] if r[0]["kind"] == kind]
            if not rows:
                continue
            kids.append(_para(f"{KINDS[kind][0]} ({_open_of(rows, kind)})", bold=True))
            kids += [_todo(it, chk, note) for it, chk, note in rows]
        blocks.append(_toggle(title, _cap(kids or [_para("Nothing open.", color="gray")])))
        ch = cls._changes_block(rec.get("changes"))
        if ch:
            blocks.append(ch)
        stmts = rec.get("statements") or []
        if stmts:
            kids = []
            if abs(rec.get("gap") or 0) >= 0.01 and m["tieout"]:
                kids.append(_para(
                    f"The current statements disagree: their open invoices add to "
                    f"{_money(rec['merged_total'])}, the newest Amount Due is "
                    f"{_money(rec['amount_due'])} ({_money(abs(rec['gap']))} apart). "
                    f"Each statement ties out on its own.", color="orange"))
            for st in stmts:
                bits = [_us(st.get("as_of", "")) or st.get("as_of_raw", "undated"), st["label"]]
                if st.get("kind"):
                    bits.append(st["kind"])
                if st.get("amount") is not None:
                    bits.append(_money(st["amount"]))
                bits.append(st["name"])
                kids.append(_para(" · ".join(bits), color="default" if st["label"] == "Current" else "gray"))
            blocks.append(_toggle(f"{H_STATEMENTS} ({len(stmts)})", _cap(kids)))
        if m["cleared"]:
            kids = []
            for c in sorted(m["cleared"].values(), key=lambda c: _us_sort(c.get("on", "")),
                            reverse=True):
                label = KINDS.get(c["kind"], ("", c["kind"]))[1]
                kids.append({"type": "to_do", "to_do": {"checked": True, "rich_text": [
                    _rt(f"{label} {c.get('asof') or '-'}: {c.get('line') or c['ref']}",
                        strikethrough=True, color="gray"),
                    _rt(f"  cleared {c.get('on') or _today()}", italic=True, color="gray")]}})
            blocks.append(_toggle(f"{H_CLEARED} ({len(m['cleared'])})", kids[:MAX_CHILDREN]))
        return blocks

    @staticmethod
    def _props(rec: dict, m: dict, st_name: str, standing: str) -> dict:
        props = {
            "Name": {"title": [_rt(rec["vendor"])]},
            "Key": {"rich_text": [_rt(rec["key"])]},
            "Status": {"select": {"name": st_name}},
            "Standing": {"rich_text": [_rt(standing, color=STATUS_COLOR.get(st_name, "default"),
                                           bold=st_name != "All entered")]},
            "Follow-ups": {"number": m["followups"]},
            "Last checked": {"date": {"start": dt.datetime.now().astimezone().isoformat(timespec="minutes")}},
            "Checked against": {"rich_text": [_rt(rec.get("checked_against", ""))]},
        }
        for name, kind in BUCKET_PROPS:
            props[name] = {"number": m["counts"].get(kind, 0)}
        if rec.get("matched") is not None:
            props["Matched"] = {"number": rec["matched"]}
        if rec.get("as_of"):
            props["Last statement"] = {"date": {"start": rec["as_of"]}}
        if rec.get("vendor_id"):
            props["QBO vendor"] = {"url": QBO_VENDOR_URL.format(vendor_id=rec["vendor_id"])}
        return props

    # ── one vendor ──
    def _plan(self, rec: dict) -> tuple:
        """Read the page (if any) and merge - reads only."""
        page = self.nc.query_by_property(self.ds, "Key", "rich_text", rec["key"])
        prior = self._read_body(page["id"]) if page else {"managed": [], "items": {}, "cleared": {}}
        run = {"items": [dict(it, asof=_us(rec.get("as_of", ""))) for it in rec.get("items", [])],
               "tieout": rec.get("tieout", True), "parsed": rec.get("parsed", True)}
        m = merge(run, prior, frozenset(rec.get("unchecked_kinds", ())))
        as_of = rec.get("as_of") or ((((page or {}).get("properties", {}).get("Last statement")
                                       or {}).get("date") or {}).get("start") or "")
        st_name, standing = status(m, as_of, rec.get("gap", 0.0))
        return page, prior, m, st_name, standing

    def preview(self, rec: dict) -> dict:
        """What sync() would do for one vendor, writing nothing (--dry-run)."""
        page, _prior, m, st_name, standing = self._plan(rec)
        return {"name": rec["vendor"], "action": "UPDATE" if page else "NEW PAGE",
                "status": st_name, "standing": standing, "not_entered": m["not_entered"],
                "followups": m["followups"], "counts": m["counts"]}

    def sync(self, rec: dict, log=print) -> dict:
        """Create or rewrite the page for one vendor. `rec` = {key, vendor,
        vendor_id, as_of, items, tieout, parsed, gap, merged_total, amount_due,
        matched, statements, changes, checked_against, unchecked_kinds}."""
        self.ensure_schema(log)
        page, prior, m, st_name, standing = self._plan(rec)
        props = self._props(rec, m, st_name, standing)
        blocks = self._body_blocks(m, st_name, standing, rec)
        if page is None:
            page = self.nc.create_page(self.ds, props, children=blocks)
        else:
            self.nc.update_page(page["id"], props)
            for bid in prior["managed"]:
                self.nc.delete_block(bid)
            self.nc.append_children(page["id"], blocks)
        return {"name": rec["vendor"], "status": st_name, "standing": standing,
                "not_entered": m["not_entered"], "followups": m["followups"],
                "url": page.get("url", "")}


def group_results(results: List[dict]) -> List[dict]:
    """Per-vendor reconcile results -> board records, keyed by the vendor's folder
    on the share. A vendor with no readable current statement ('unreadable')
    keeps its last known items and shows Unreadable."""
    recs: Dict[str, dict] = {}
    for res in results:
        if res.get("action") not in ("filed", "held", "preview", "unreadable"):
            continue
        folder = res.get("vendor_folder") or res.get("vendor") or ""
        if not folder:
            continue
        key = folder.upper()
        rec = recs.setdefault(key, {"key": key, "vendor": folder, "vendor_id": "", "folder": folder,
                                    "items": [], "tieout": True, "parsed": False, "as_of": "",
                                    "unchecked_kinds": set()})
        if res["action"] == "unreadable":
            rec["tieout"] = False
            rec["as_of"] = rec["as_of"] or res.get("as_of", "")
            continue
        for k in ("vendor", "vendor_id", "as_of", "gap", "merged_total", "amount_due", "matched",
                  "statements", "changes", "checked_against"):
            if res.get(k) not in (None, ""):
                rec[k] = res[k]
        rec["parsed"] = True
        rec["items"] += res.get("items") or []
        rec["tieout"] = rec["tieout"] and bool(res.get("tieout", True))
        rec["unchecked_kinds"].update(res.get("unchecked_kinds") or ())
    return list(recs.values())
