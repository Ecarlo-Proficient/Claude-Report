"""
draw_moves.py - the ONE "this bill rides a later draw" override (the push).

WHY THIS EXISTS
A bill lands in a draw by DATE: its TxnDate falls inside the draw invoice's "(Period: ...)" tag.
Three readers apply that rule independently - project-pnl's draw buckets (the P&L workbook's
draw sheets, coverage table and Labor/Concrete grids), the Bill Tracker's CP/MFD matcher
(`matched_invoice`, which the ledger mirrors), and the ledger dashboard's draw bands. When we
agree with a supplier to carry end-of-period bills into the NEXT draw (CP800, 2026-09-02:
Preferred Materials bills dated after 07/20/26 move from Draw #3 to Draw #4 - most of them were
end-of-billing-term), every reader must move the SAME bills the SAME way and say so on its face.
The bill keeps its real date everywhere; only the draw it rides changes, and the move is labelled
("pushed from Draw #3") wherever the bill is listed.

THE RULES FILE  <companyhealth_dir>/draw_moves.json   (local business data - never in the repo)
{
  "rules": [
    {"project": "CP800", "vendor": "Preferred Materials",
     "after": "2026-07-20", "through": "2026-07-31", "move_to": "2026-08-01",
     "from_draw": 3, "to_draw": 4,
     "why": "agreed with the supplier 2026-09-02 - end-of-term bills carry into Draw #4"}
  ]
}
  project    exact project # (case-insensitive) - never a family match (-FTW is its own job)
  vendor     case-insensitive substring of the QBO vendor name ("Preferred Materials")
  after      bills dated AFTER this date move (the cutoff itself stays in the old draw)
  through    ... and up to this date inclusive (the old draw's period end)
  move_to    the date the bill is bucketed AS - any day inside the target draw's period
  from_draw / to_draw   draw numbers, for the label only
  why        free text, shown as the note wherever the push is explained

Every reader is ABSENT-SAFE: no file / no rules -> nothing moves, behaviour exactly as before.
`shared/` is the only importable common code, so all three readers use this one module.
"""
from __future__ import annotations

import datetime as dt
import json
import re
from pathlib import Path
from typing import Iterable, List, Optional, Tuple, Union

try:                                   # sibling shared module (package import)
    from . import paths
except ImportError:                    # run with the repo root on sys.path
    import paths                       # type: ignore

RULES_FILE: Path = paths.companyhealth_dir() / "draw_moves.json"

DateLike = Union[str, dt.date, None]

_cache: Optional[List[dict]] = None
_cache_path: Optional[Path] = None


def _to_date(v: DateLike) -> Optional[dt.date]:
    if v is None or v == "":
        return None
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    m = re.match(r"\s*(\d{4})-(\d{1,2})-(\d{1,2})", str(v))
    if m:
        return dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    m = re.match(r"\s*(\d{1,2})/(\d{1,2})/(\d{2,4})", str(v))
    if m:
        y = int(m.group(3))
        y = y + 2000 if y < 100 else y
        return dt.date(y, int(m.group(1)), int(m.group(2)))
    return None


def _norm_rule(raw: dict) -> Optional[dict]:
    """Validate one rule; None when it cannot be applied (a broken rule must not
    silently move nothing AND look like it did - the caller logs the skip)."""
    project = str(raw.get("project") or "").strip().upper()
    vendor = str(raw.get("vendor") or "").strip().lower()
    after, through, move_to = (_to_date(raw.get("after")), _to_date(raw.get("through")),
                               _to_date(raw.get("move_to")))
    if not (project and vendor and after and through and move_to):
        return None
    if through <= after or move_to <= through:
        return None
    return {
        "project": project, "vendor": vendor, "after": after, "through": through,
        "move_to": move_to,
        "from_draw": raw.get("from_draw"), "to_draw": raw.get("to_draw"),
        "why": str(raw.get("why") or "").strip(),
    }


def load_rules(path: Optional[Path] = None) -> List[dict]:
    """The validated rules, cached per process. [] when the file is absent or empty."""
    global _cache, _cache_path
    p = Path(path) if path else RULES_FILE
    if _cache is not None and _cache_path == p:
        return _cache
    rules: List[dict] = []
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
        for raw in (data.get("rules") if isinstance(data, dict) else data) or []:
            r = _norm_rule(raw) if isinstance(raw, dict) else None
            if r:
                rules.append(r)
    except (OSError, ValueError):
        rules = []
    _cache, _cache_path = rules, p
    return rules


def reload() -> List[dict]:
    """Drop the cache (a long-running dashboard picks up an edited file)."""
    global _cache
    _cache = None
    return load_rules()


def find_move(project_no: Optional[str], vendor: Optional[str], txn_date: DateLike,
              rules: Optional[Iterable[dict]] = None) -> Optional[dict]:
    """The rule that moves this bill, or None. Exact project #, vendor substring,
    date strictly after the cutoff and up to `through` inclusive."""
    d = _to_date(txn_date)
    if not d or not project_no or not vendor:
        return None
    pn = str(project_no).strip().upper()
    vn = str(vendor).strip().lower()
    for r in (rules if rules is not None else load_rules()):
        if r["project"] == pn and r["vendor"] in vn and r["after"] < d <= r["through"]:
            return r
    return None


def effective_date(project_no: Optional[str], vendor: Optional[str], txn_date: DateLike,
                   rules: Optional[Iterable[dict]] = None) -> Tuple[DateLike, Optional[dict]]:
    """(the date to BUCKET this bill by, the rule that moved it or None). The returned
    date keeps the caller's type: a `date` in -> `date` out, an ISO string in -> ISO out.
    An unmoved bill comes back untouched."""
    r = find_move(project_no, vendor, txn_date, rules)
    if not r:
        return txn_date, None
    mv = r["move_to"]
    if isinstance(txn_date, dt.date):
        return mv, r
    return mv.isoformat(), r


def push_label(rule: dict) -> str:
    """Short mark for a bill row: 'pushed from Draw #3'."""
    fr = rule.get("from_draw")
    return f"pushed from Draw #{fr}" if fr else "pushed from the prior draw"


def push_note(rule: dict, direction: str = "in") -> str:
    """One sentence for a draw's summary. direction 'in' = the receiving draw,
    'out' = the draw the bills left."""
    fr, to = rule.get("from_draw"), rule.get("to_draw")
    cut = rule["after"].strftime("%m/%d/%y")
    who = rule["vendor"].title()
    if direction == "out":
        head = (f"{who} bills dated after {cut} were pushed to Draw #{to}"
                if to else f"{who} bills dated after {cut} were pushed to the next draw")
    else:
        head = (f"{who} bills dated after {cut} were pushed in from Draw #{fr}"
                if fr else f"{who} bills dated after {cut} were pushed in from the prior draw")
    return f"{head} - {rule['why']}" if rule.get("why") else head


if __name__ == "__main__":             # metadata-only check: which rules load
    rs = load_rules()
    print(f"{RULES_FILE}: {len(rs)} rule(s)")
    for r in rs:
        print(f"  {r['project']}  {r['vendor']!r}  after {r['after']} through {r['through']}"
              f"  -> as of {r['move_to']}  (Draw #{r['from_draw']} -> #{r['to_draw']})")
