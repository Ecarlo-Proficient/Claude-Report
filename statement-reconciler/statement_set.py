"""statement_set.py - one vendor's statements as ONE running record (tool-local).

The owner (10/07/2026): a vendor is not a stack of months - it is one record,
"all entered AS OF the last date the statement shows". The clerk dumps every
statement in the Inbox in any order; this module decides what each one is and
which ones the vendor's standing is built from.

Every statement covers a RANGE of invoice dates:
  full      a complete open list (QBO open balance, most vendors)
            -> covers every invoice dated up to its as-of date
  pastdue   a past-due letter (Cowtown) - lists only invoices already past due
            -> covers invoices up to the end of the month of its newest invoice
  activity  a statement with a "Balance forward" lump (Cowtown monthly)
            -> covers only the invoices dated after the balance forward

merge_docs() walks the statements oldest first: a statement wipes every invoice
inside its range, then adds its own. What is left is the vendor's open list. A
statement that no longer contributes a line is history (Replaced / DONE).
An activity statement's balance-forward lump and payments count only while no
other current statement itemizes that earlier period - otherwise the past-due
letter already shows those invoices net of the payments.

The vendor folder on the share:
  <Vendor>/Current/   the statement(s) the standing is built from + ONE Excel
  <Vendor>/History/   '<mm-dd-yyyy> DONE|Replaced|Duplicate - <file>'
  <Vendor>/.reconciler.json   last as-of, clean flag, changes, the clerk's notes

Pure logic here (no QBO, no Notion); the reconciler does the I/O around it.
"""
from __future__ import annotations

import calendar
import datetime as dt
import hashlib
import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple

CURRENT, HISTORY = "Current", "History"
STATE_FILE = ".reconciler.json"
REPORT_PREFIX = "Reconciliation - "
TIE_TOLERANCE = 0.50
LEGACY_MONTH_RE = re.compile(r"^(\d\d)-(\d{4})(?:\s+DONE)?$", re.I)
HISTORY_RE = re.compile(r"^(\S+) (DONE|Replaced|Duplicate|Not read) - (.+)$")

# Tool-written notes in the old 'Note' column - anything else there was typed by
# the clerk (pre-10/07 workbooks had no column of her own).
TOOL_NOTE_PREFIXES = ("Stmt = QBO", "Diff = $", "Paid in QBO", "ENTER BILL IN QBO",
                      "Open in QBO but not on statement", "QBO bill has no DocNumber")


@dataclass
class Doc:
    path: Path
    as_of: str                      # ISO date the statement is "as of"
    amount_due: float
    lines: list                     # StmtLine-like: .date .ref .amount (.po .address)
    template: str = ""
    kind: str = "full"              # full | pastdue | activity
    lo: str = ""                    # covers invoice dates > lo ("" = no floor)
    hi: str = ""                    # ... and <= hi ("" = no ceiling)
    sha: str = ""
    origin: str = "inbox"           # inbox | current | legacy | legacy_done | history
    status: str = ""                # current | replaced | duplicate | unreadable (set by plan)

    @property
    def line_sum(self) -> float:
        return round(sum(l.amount for l in self.lines), 2)

    @property
    def ties_out(self) -> bool:
        return bool(self.lines) and abs(self.line_sum - self.amount_due) <= TIE_TOLERANCE

    @property
    def problem(self) -> str:
        """Why this statement can't be trusted, '' when it can. The safety nets: a
        misread must never become the vendor's open list."""
        if not self.lines:
            return "no lines read (new layout?)"
        if not self.as_of:
            return "no statement date"
        if not self.ties_out:
            return (f"lines add to {self.line_sum:,.2f}, its Amount Due is "
                    f"{self.amount_due:,.2f}")
        newest = max((l.date for l in self.lines if l.ref and re.match(r"\d{4}-\d\d-\d\d$", l.date or "")),
                     default="")
        if newest and self.as_of < newest:
            return f"dated {us_date(self.as_of)}, before its own invoice of {us_date(newest)}"
        if self.as_of > (dt.date.today() + dt.timedelta(days=1)).isoformat():
            return f"dated {us_date(self.as_of)}, in the future"
        return ""

    @property
    def bf_date(self) -> str:
        return next((l.date for l in self.lines if not l.ref), "")

    def covers(self, date: str) -> bool:
        if not date:
            return not self.lo              # undated lines: only a full/past-due list owns them
        return (not self.lo or date > self.lo) and (not self.hi or date <= self.hi)


def _month_end(iso: str) -> str:
    y, m = int(iso[:4]), int(iso[5:7])
    return f"{y:04d}-{m:02d}-{calendar.monthrange(y, m)[1]:02d}"


def classify(doc: Doc) -> Doc:
    """Set doc.kind and its invoice-date range from its template and lines."""
    dated = sorted(l.date for l in doc.lines if l.ref and re.match(r"\d{4}-\d\d-\d\d$", l.date or ""))
    if doc.template == "vendor_cowtown":
        doc.kind, doc.lo = "pastdue", ""
        doc.hi = _month_end(dated[-1]) if dated else doc.as_of
    elif doc.bf_date:
        doc.kind, doc.lo, doc.hi = "activity", doc.bf_date, doc.as_of
    else:
        doc.kind, doc.lo, doc.hi = "full", "", doc.as_of
    return doc


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def _key(line, i: int, n: int) -> str:
    """Invoices key by their #; payments/credits are not unique by # (one check
    pays many jobs), so they key by their position."""
    return line.ref.strip().upper() if line.amount >= 0 else f"~{i}.{n}"


@dataclass
class Merged:
    lines: list = field(default_factory=list)     # the vendor's open list
    contributors: List[int] = field(default_factory=list)   # indexes into docs
    as_of: str = ""
    amount_due: float = 0.0
    line_sum: float = 0.0

    @property
    def gap(self) -> float:
        return round(self.line_sum - self.amount_due, 2)


def _partnered(d: Doc, docs: List[Doc]) -> bool:
    """Does another statement of the SAME batch (within BATCH_DAYS) itemize the
    period before this activity statement's balance forward? Cowtown's 10/01
    letter does for its 10/02 statement. A month-old open list does not - it is
    stale (Preferred 09/01 vs 10/05), so the balance forward stands for it."""
    lo = _days_before(d.as_of, BATCH_DAYS)
    return any(o is not d and o.kind != "activity" and lo <= o.as_of and o.covers(d.bf_date)
               for o in docs)


def merge_docs(docs: List[Doc]) -> Merged:
    """Merge readable statements (any order in) into one open list. Pure."""
    for d in docs:
        if d.kind == "activity":                 # unpartnered: its lump covers all before
            d.lo = d.bf_date if _partnered(d, docs) else ""
    order = sorted(range(len(docs)), key=lambda i: (docs[i].as_of, str(docs[i].path.name)))
    cur: Dict[str, Tuple[object, int]] = {}
    for i in order:
        d = docs[i]
        for k in [k for k, (l, _src) in cur.items() if d.covers(l.date)]:
            del cur[k]
        for n, l in enumerate(d.lines):
            if l.ref:
                cur[_key(l, i, n)] = (l, i)
    contrib = sorted({src for _l, src in cur.values()}, key=lambda i: order.index(i))
    out = [(l, src) for l, src in cur.values()]
    # Balance forward + payments of an activity statement: only while nobody else
    # itemizes the period before its balance forward.
    for i in contrib:
        d = docs[i]
        if d.kind != "activity":
            continue
        itemized = any(j != i and docs[j].kind != "activity" and docs[j].covers(d.bf_date)
                       for j in contrib)
        if itemized:
            out = [(l, src) for l, src in out if not (src == i and l.amount < 0)]
        else:
            out.insert(0, (next(l for l in d.lines if not l.ref), i))
    newest = docs[order[-1]] if order else None
    lines = [l for l, _src in out]
    return Merged(lines=lines, contributors=contrib,
                  as_of=newest.as_of if newest else "",
                  amount_due=newest.amount_due if newest else 0.0,
                  line_sum=round(sum(l.amount for l in lines), 2))


def plan(docs: List[Doc]) -> Tuple[Merged, List[Doc]]:
    """Duplicates out, unreadable out, merge the rest; label every doc. Docs
    already filed win a duplicate tie (the Inbox copy is the extra one)."""
    # Who keeps the file when two are the same: the one in Current, then the old
    # month folders, then History (a History twin only catches a NEW copy - it
    # must never knock out the live one: RCI refresh 10/08), then the Inbox.
    rank = {"current": 0, "legacy": 1, "legacy_done": 2, "history": 3, "inbox": 4, "manual": 5}
    seen: Dict[str, Doc] = {}
    for d in sorted(docs, key=lambda d: (rank.get(d.origin, 9), d.path.name)):
        if d.origin == "history":
            seen.setdefault(d.sha, d)
            continue
        if d.sha and d.sha in seen:
            d.status = "duplicate"
            continue
        if d.sha:
            seen[d.sha] = d
        if d.problem:
            d.status = "unreadable"
    live = [d for d in docs if d.origin != "history" and not d.status]
    m = merge_docs(live)
    for i, d in enumerate(live):
        d.status = "current" if i in m.contributors else "replaced"
    return m, docs


def diff_sets(old_lines: list, new_lines: list) -> dict:
    """What changed between two open lists: new invoices, invoices gone (paid or
    credited in between) and invoices whose amount the vendor changed."""
    o = {l.ref.strip().upper(): l for l in old_lines if l.ref and l.amount >= 0}
    n = {l.ref.strip().upper(): l for l in new_lines if l.ref and l.amount >= 0}
    row = lambda l: {"ref": l.ref, "date": l.date, "amount": round(l.amount, 2),
                     "where": (getattr(l, "address", "") or "")[:40]}
    return {
        "new": [row(n[r]) for r in n if r not in o],
        "gone": [row(o[r]) for r in o if r not in n],
        "changed": [dict(row(n[r]), was=round(o[r].amount, 2)) for r in n
                    if r in o and abs(o[r].amount - n[r].amount) > 0.005],
    }


BATCH_DAYS = 7          # statements this close together are one update (Cowtown 10-01 + 10-02)


def _days_before(iso: str, days: int) -> str:
    try:
        return (dt.date.fromisoformat(iso) - dt.timedelta(days=days)).isoformat()
    except ValueError:
        return iso


def changes_for(docs: List[Doc], merged: Merged, state: dict) -> Optional[dict]:
    """Changes since the previous update. The newest batch is every statement
    within BATCH_DAYS of the newest as-of; the previous update is everything older.
    Uses the statements here when there are older ones, else what the last run
    stored for this as-of."""
    cut = _days_before(merged.as_of, BATCH_DAYS)
    older = [d for d in docs if d.status in ("current", "replaced") and d.as_of < cut]
    if older:
        prev = merge_docs(older)
        return dict(diff_sets(prev.lines, merged.lines), frm=prev.as_of, to=merged.as_of)
    ch = state.get("changes")
    return ch if ch and ch.get("to") == merged.as_of else None


# ─────────────────────────── the vendor folder ───────────────────────────

def date_from_name(name: str, mtime: float = 0.0) -> str:
    """ISO date from a file name ('Statement BURNCO 06-09.xlsx', '... 09-25-2026',
    '... 9-25-26'); a name without a year takes the file's own year. '' if none -
    an undated Excel statement is never silently dated today."""
    m = re.search(r"(?<!\d)(\d{1,2})-(\d{1,2})(?:-(\d{2,4}))?(?!\d)", name)
    if not m:
        return ""
    mm, dd, yy = int(m.group(1)), int(m.group(2)), m.group(3)
    if yy:
        year = int(yy) + (2000 if len(yy) == 2 else 0)
    elif mtime:
        year = dt.date.fromtimestamp(mtime).year
    else:
        return ""
    try:
        return dt.date(year, mm, dd).isoformat()
    except ValueError:
        return ""


def us_date(iso: str, sep: str = "/") -> str:
    """mm/dd/yyyy (or mm-dd-yyyy) from ISO - never year-first (owner rule)."""
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})", iso or "")
    return sep.join((m.group(2), m.group(3), m.group(1))) if m else (iso or "")


def history_name(as_of: str, label: str, name: str) -> str:
    return f"{us_date(as_of, '-') or 'undated'} {label} - {name}"


def report_name(vendor_folder: str, as_of: str) -> str:
    return f"{REPORT_PREFIX}{vendor_folder} - as of {us_date(as_of, '-')}.xlsx"


def is_report(name: str) -> bool:
    """A reconciliation Excel (ours) - also once filed in History under a prefix."""
    m = HISTORY_RE.match(name)
    base = m.group(3) if m else name
    return base.startswith(REPORT_PREFIX) or base.startswith("Statement_Reconciliation_")


def is_junk(name: str) -> bool:
    return name.startswith(".") or name.startswith("~$") or name.endswith("#")


def load_state(vendor_dir: Path) -> dict:
    try:
        return json.loads((vendor_dir / STATE_FILE).read_text())
    except (OSError, ValueError):
        return {}


def save_state(vendor_dir: Path, state: dict) -> None:
    vendor_dir.mkdir(parents=True, exist_ok=True)
    tmp = vendor_dir / (STATE_FILE + ".tmp")
    tmp.write_text(json.dumps(state, indent=1, sort_keys=True))
    tmp.replace(vendor_dir / STATE_FILE)


# ─────────────────────────── the clerk's notes ───────────────────────────

def note_key(bill_id: str = "", ref: str = "") -> str:
    """A note follows the QBO bill (its id); a bill not in QBO yet follows its #."""
    if bill_id:
        return f"bill:{bill_id}"
    return f"ref:{ref.strip().upper()}" if ref and ref.strip() else ""


def note_for(notes: dict, bill_id: str = "", ref: str = "") -> str:
    for k in (note_key(bill_id=bill_id), note_key(ref=ref)):
        if k and k in notes:
            return notes[k].get("note", "")
    return ""


def harvest_notes(xlsx: Path) -> Dict[str, Optional[str]]:
    """The clerk's notes in a reconciliation workbook, keyed like note_key().
    A row whose notes cell is empty maps to None (she cleared it). Reads the
    'Clerk notes' column; on a pre-10/07 workbook, the 'Note' column text that
    the tool did not write."""
    from openpyxl import load_workbook
    out: Dict[str, Optional[str]] = {}
    try:
        wb = load_workbook(xlsx, data_only=True)
    except Exception:
        return out
    if "Summary" not in wb.sheetnames:
        return out
    ws = wb["Summary"]
    col = None                                   # 13 = Clerk notes, 12 = legacy Note
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        b = row[1].value if len(row) > 1 else None
        if b == "Stmt Ref #":
            # Her column: ours is 'Clerk notes'; before 10/07 she added her own
            # ('<name>'S NOTES') in column M - any header saying NOTES counts.
            heads = [str(c.value or "") for c in row]
            col = 13 if len(heads) >= 13 and "NOTES" in heads[12].upper() else 12
            continue
        if col is None or len(row) < col:
            continue
        a = row[0]
        stmt_ref = str(row[1].value or "").strip()
        qbo_ref = str(row[2].value or "").strip()
        if not (stmt_ref or qbo_ref):
            continue                              # section header / spacer row
        bill_id = ""
        if a.hyperlink is not None and a.hyperlink.target:
            m = re.search(r"txnId=(\w+)", a.hyperlink.target)
            bill_id = m.group(1) if m else ""
        text = str(row[col - 1].value or "").strip()
        k = note_key(bill_id=bill_id, ref=stmt_ref or qbo_ref)
        if not k:
            continue
        if col == 13:
            out[k] = text or None
        elif text and not text.startswith(TOOL_NOTE_PREFIXES):
            out[k] = text                         # legacy: only what she typed
    return out


def fold_notes(notes: dict, harvested: Dict[str, Optional[str]], today: str) -> dict:
    """Harvested notes win; an emptied cell drops the note."""
    out = dict(notes)
    for k, v in harvested.items():
        if v is None:
            out.pop(k, None)
        elif v and (out.get(k) or {}).get("note") != v:
            out[k] = {"note": v, "on": today}
    return out
