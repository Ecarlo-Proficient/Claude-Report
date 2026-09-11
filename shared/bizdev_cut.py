#!/usr/bin/env python3
"""
bizdev_cut.py - the ONE test for "is this line a business-development cut?"
(the owner 2026-09-09).

THE REGISTER  <companyhealth_dir>/bizdev_cut.json   (local business data - never
in the repo, because it names vendors). Override: ACB_BIZDEV_CUT_FILE.

WHAT THIS IS FOR
Some outside parties are paid a cut of the work they bring in, and they invoice
it against whichever job is open at the time - the job number is written on the
line by them, in their own coding. That money is NOT what it took to build the
job, so it must not sit in job cost: it makes the jobs that happened to be open
look like losers and the jobs that were not look like winners, and no two jobs
are comparable. The owner, 2026-09-09, on their estimator's hours:

    "is it a job cost or not. it's not it's overhead"

THE TEST is one question: **did this line build the job?**
  yes -> job cost. Material, fuel, safety, rentals, petty cash, forms and
         batters they laid out for the job and were paid back for.
  no  -> overhead. Their draw, commission, project-management fee, allowance,
         vehicle, insurance, bonus, settlement, estimating fee, software, and
         the hours of the people THEY employ - even when the line carries a job
         number, a job cost code, an hourly rate and a burden %.

The ACCOUNT is NOT the test and never was. The same cut has been booked to
Subcontractors Expense (a COGS account) and to Business Development Services (an
operating account) on the same jobs; and real material they fronted has been
booked to Subcontractors Expense too. Only the line text separates them.

Register format:
{
  "_comment": "...",
  "vendors":  ["<vendor DisplayName substring>", ...],   case-insensitive
  "fronted":  ["<regex>", ...],   line text meaning "they laid this out for the
                                   job" - checked FIRST, wins over everything
  "pay_accounts": ["<account name prefix>", ...]         accounts that can only
                                   ever hold their own compensation
  "labels": {"<vendor key>": "<column header>"}          optional - what the
                                   owner's own cut workbook calls each one
  "categories": [{"name", "who", "pattern", "memo"?}]     optional - categories of
                                   the cut, first match wins (see categorize)
  "who_labels": {"<who>": "<column header>"}              optional
  "director": "<vendor key>"                              optional - the one whose
                                   cut that workbook charges AFTER overhead; every
                                   other registered vendor's cut is charged into
                                   the job's COST there (the owner 2026-09-11)
}
A line from a registered vendor is a CUT when its account is one of
`pay_accounts` and its text does not match `fronted`. A line in any other
account is a real expense they fronted, so it stays in job cost.

Who reads it (ONE test, so the reports cannot disagree):
  project-pnl   -> keeps cut lines out of job cost on every Job P&L, and so out
                   of the division Overview, which is built from those workbooks
  the owner's own CompanyHealth analysis -> totals the same lines as the cut

No register, or no vendors in it -> `is_cut()` is always False and nothing
changes. That is deliberate: a missing register must never silently move money.
"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path
from typing import List, Optional, Tuple

_HERE = Path(__file__).resolve().parent
if str(_HERE.parent) not in sys.path:
    sys.path.insert(0, str(_HERE.parent))
from shared import paths  # noqa: E402

CUT_FILE: Path = paths.get_path("ACB_BIZDEV_CUT_FILE",
                                paths.companyhealth_dir() / "bizdev_cut.json")

_cache: Optional[dict] = None
_cache_key = None


def load(path: Optional[Path] = None) -> dict:
    """{'vendors': [...], 'fronted': <compiled re|None>, 'pay_accounts': (...)}.

    Empty vendors when the register is missing or unreadable - callers then get
    False from every test and the reports are exactly as they were."""
    global _cache, _cache_key
    p = Path(path) if path else CUT_FILE
    try:
        key = (str(p), p.stat().st_mtime_ns)
    except OSError:
        return {"vendors": (), "fronted": None, "pay_accounts": (), "labels": {}, "director": "",
                "categories": (), "who_labels": {}}
    if _cache is not None and _cache_key == key:
        return _cache
    try:
        raw = json.loads(p.read_text(encoding="utf-8"))
    except (ValueError, OSError) as e:
        print(f"  (business-development register unreadable: {p}: {e})", file=sys.stderr)
        return {"vendors": (), "fronted": None, "pay_accounts": (), "labels": {}, "director": "",
                "categories": (), "who_labels": {}}
    pats: List[str] = [str(x) for x in (raw.get("fronted") or []) if str(x).strip()]
    try:
        fronted = re.compile("|".join(f"(?:{x})" for x in pats), re.I) if pats else None
    except re.error as e:
        print(f"  (business-development 'fronted' pattern bad: {e})", file=sys.stderr)
        fronted = None
    out = {
        "vendors": tuple(str(v).strip().upper() for v in (raw.get("vendors") or [])
                         if str(v).strip()),
        "fronted": fronted,
        "pay_accounts": tuple(str(a).strip().upper() for a in (raw.get("pay_accounts") or [])
                              if str(a).strip()),
        "labels": {str(k).strip().upper(): str(v).strip()
                   for k, v in (raw.get("labels") or {}).items() if str(v).strip()},
        "director": str(raw.get("director") or "").strip().upper(),
        "categories": _compile_categories(raw.get("categories") or []),
        "who_labels": {str(k).strip().lower(): str(v).strip()
                       for k, v in (raw.get("who_labels") or {}).items() if str(v).strip()},
    }
    _cache, _cache_key = out, key
    return out


def _compile_categories(raw_list) -> tuple:
    """[(name, who, desc regex, memo regex|None), ...] in register order."""
    out = []
    for item in raw_list:
        try:
            name = str(item.get("name") or "").strip()
            who = str(item.get("who") or "").strip().lower()
            pat = re.compile(str(item.get("pattern") or ""), re.I)
            mpat = re.compile(str(item["memo"]), re.I) if item.get("memo") else None
        except (AttributeError, re.error, KeyError) as e:
            print(f"  (business-development category skipped: {item!r}: {e})", file=sys.stderr)
            continue
        if name and who:
            out.append((name, who, pat, mpat))
    return tuple(out)


def categorize(vendor: str, account: str, text: str, memo: str = "",
               path: Optional[Path] = None) -> Tuple[str, str]:
    """(category, who) for ONE line of a registered vendor - the owner's
    "categories of his cut" (2026-09-11): draw, weekly pay, allowances, bonus,
    estimating, estimating software, and the three he will ask the director
    about (burden, sub service, hourly help). First match wins on the line
    text; memo-only rules run when the text matched nothing. A line the
    `fronted` test catches is job cost before any category; no rule at all is
    "Other" and a blank line is "Blank line", both "ask the director" so
    nothing is silently filed."""
    if not is_vendor(vendor, path):
        return "", ""
    if not is_cut(vendor, account, text or memo, path):
        return "Fronted (job cost)", "the job"
    reg = load(path)
    t, m = text or "", memo or ""
    for name, who, pat, _m in reg["categories"]:
        if pat.search(t):
            return name, who
    for name, who, _p, mpat in reg["categories"]:
        if mpat is not None and mpat.search(m):
            return name, who
    if not t.strip() and not m.strip():
        return "Blank line", "ask the director"
    return "Other", "ask the director"


def who_label(who: str, path: Optional[Path] = None) -> str:
    """The column header for a `who` - the register's `who_labels`, else the
    word itself in capitals."""
    return load(path)["who_labels"].get((who or "").lower()) or (who or "").upper()


def is_vendor(vendor: str, path: Optional[Path] = None) -> bool:
    """Is this vendor one the owner registered as taking a cut?"""
    v = (vendor or "").upper()
    reg = load(path)
    return bool(v) and any(w in v for w in reg["vendors"])


def is_cut(vendor: str, account: str, text: str, path: Optional[Path] = None) -> bool:
    """True when this ONE line is their own compensation, not a cost of building.

    vendor   the bill's vendor DisplayName
    account  the line's account name (fully qualified, e.g. "Subcontractors Expense")
    text     the line description; fall back to the bill memo when it is blank
    """
    if not is_vendor(vendor, path):
        return False
    reg = load(path)
    acct = (account or "").upper()
    if not any(acct.startswith(a) for a in reg["pay_accounts"]):
        return False                      # a real expense account -> they fronted it
    if reg["fronted"] is not None and reg["fronted"].search(text or ""):
        return False                      # the text says they laid it out
    return True


def vendor_key(vendor: str, path: Optional[Path] = None) -> str:
    """The register key a vendor DisplayName matched ("" when none)."""
    v = (vendor or "").upper()
    for w in load(path)["vendors"]:
        if w in v:
            return w
    return ""


def is_director(vendor: str, path: Optional[Path] = None) -> bool:
    """Is this the registered director - the one cut charged after overhead on
    the owner's page? With no `director` in the register, every registered
    vendor counts (the page then carries one column each)."""
    key = vendor_key(vendor, path)
    if not key:
        return False
    d = load(path)["director"]
    return key == d if d else True


def label(vendor: str, path: Optional[Path] = None) -> str:
    """What the owner's own cut workbook calls this vendor's column - the
    register's `labels` entry, else "<vendor> CUT". Only that workbook ever
    uses it; every shared report keeps `note()` and names nobody."""
    key = vendor_key(vendor, path)
    if not key:
        return ""
    return load(path)["labels"].get(key) or f"{key.title()} CUT"


def note(total: float) -> str:
    """The one line a report carries in place of the detail. Never names anyone:
    these workbooks are shared with the people the cut is paid to."""
    return (f"excludes business development fees of {total:,.0f} "
            f"(not a job cost - carried in overhead)")


def main(argv=None) -> int:
    import argparse
    ap = argparse.ArgumentParser(description="Show the business-development cut register")
    ap.add_argument("--file", type=Path, default=None)
    a = ap.parse_args(argv)
    p = a.file or CUT_FILE
    reg = load(p)
    print(f"register: {p}   {'FOUND' if p.exists() else 'MISSING'}")
    print(f"  vendors registered : {len(reg['vendors'])}")
    print(f"  pay accounts       : {len(reg['pay_accounts'])}")
    n = 0 if reg["fronted"] is None else len(reg["fronted"].pattern.split("|"))
    print(f"  'fronted' patterns : {n}")
    if not reg["vendors"]:
        print("  -> is_cut() is always False; no report changes.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
