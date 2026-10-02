#!/usr/bin/env python3
"""
sales_report_baseline_import.py - seed the sales report's week-over-week baseline
from a PREVIOUSLY RENDERED report PDF.

WHY
`sales_rep_leads_report.py` answers "what is new since the list you already sent"
by diffing against a snapshot it wrote the week before. The very first run has no
snapshot, so it can only fall back to a date window - and a date window cannot
tell a brand-new account from one that was on the list already and got contacted
again (`customer.last_contacted` holds one latest date; the loader full-replaces,
so there is no history in the ledger).

An already-rendered report IS that history. This reads the FULL LIST table out of
one, maps the company names back to `customer`, and writes it as the snapshot for
that report's date - after which the next run diffs against real ground truth.

SAFETY
  * READ-ONLY on the PDFs and on the ledger. Writes one snapshot JSON.
  * Refuses to overwrite an existing snapshot unless --force: a recorded snapshot
    is real history and a reconstructed one is not, so the recorded one wins.
  * Needs `pdftotext` (poppler) on PATH.

USAGE
  python3 one-offs/sales_report_baseline_import.py --pdf "<report>.pdf" [--pdf ...]
  python3 one-offs/sales_report_baseline_import.py --pdf a.pdf --dry-run
"""
from __future__ import annotations

import argparse
import json
import re
import shutil
import sqlite3
import subprocess
import sys
import tempfile
import datetime as dt
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))
from shared import paths  # noqa: E402

DEFAULT_DB = paths.get_path(
    "ACB_LEDGER_DB",
    Path.home() / "Library" / "Application Support" / "Proficient" / "ledger.sqlite3",
)
SNAP_DIR = Path.home() / "Library" / "Application Support" / "Proficient" / "sales-report"
TOUCHED = ("Contacted", "Follow up", "Interested", "No response")
STAGES = TOUCHED + ("Lead", "Closed - Won", "Closed - Lost")
ROW = re.compile(
    r"^\s*(?P<co>\S.*?)\s{2,}(?P<div>RP|CP|MFD)\s+"
    r"(?P<stage>" + "|".join(map(re.escape, STAGES)) + r")\s+"
    r"(?P<lc>\d{2}/\d{2}/\d{4}|-)"
)
HDR_DATE = re.compile(r"as of (\d{2})/(\d{2})/(\d{4})")
# the full-list column is width-clipped and a live-client row carries a badge;
# both mangle the company name, so matching is normalized + prefix-tolerant.
_BADGE_FULL = "LIVE CLIENT"


def strip_badge(co: str) -> str:
    """Drop a trailing live-client badge, including a column-clipped one.

    The table renders the badge as "LIVE CLIENT" but the column width can cut it
    anywhere - "LIVE CLIEN", "LIVE C", "LIVE". So peel trailing ALL-CAPS tokens
    while what they spell is still a prefix of the full badge. Company names are
    mixed case in this table, so a genuine name is not at risk."""
    parts = co.split()
    while parts:
        tail = " ".join(parts[-1:] if len(parts) == 1 else parts[-2:])
        for n in (2, 1):
            if len(parts) >= n:
                cand = " ".join(parts[-n:])
                if cand.isupper() and _BADGE_FULL.startswith(cand) and cand != co:
                    parts = parts[:-n]
                    break
        else:
            break
        if not tail:
            break
    return " ".join(parts).strip()


def norm(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", (s or "").lower())


def pdf_text(pdf: Path) -> str:
    if not shutil.which("pdftotext"):
        raise RuntimeError("pdftotext not on PATH - install poppler (brew install poppler)")
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / "t.txt"
        subprocess.run(["pdftotext", "-layout", str(pdf), str(out)],
                       check=True, capture_output=True, timeout=120)
        return out.read_text(encoding="utf-8", errors="replace")


def parse_report(pdf: Path) -> tuple[dt.date, dict[str, str]]:
    """-> (report date, {company as printed: last_contacted ISO or ''})"""
    lines = pdf_text(pdf).splitlines()
    m = next((HDR_DATE.search(ln) for ln in lines[:6] if HDR_DATE.search(ln)), None)
    if not m:
        raise ValueError(f"{pdf.name}: no 'as of MM/DD/YYYY' header - not a report of this shape")
    when = dt.date(int(m.group(3)), int(m.group(1)), int(m.group(2)))
    try:
        i = next(k for k, ln in enumerate(lines) if ln.strip().startswith("FULL LIST"))
    except StopIteration:
        raise ValueError(f"{pdf.name}: no FULL LIST table - this is a different kind of report")
    rows: dict[str, str] = {}
    for ln in lines[i + 1:]:
        g = ROW.match(ln)
        if not g:
            continue
        co = strip_badge(g.group("co").strip())
        if co.upper() == "COMPANY":
            continue
        lc = g.group("lc")
        rows[co] = "" if lc == "-" else f"{lc[6:10]}-{lc[0:2]}-{lc[3:5]}"
    return when, rows


def match_to_ledger(printed: dict[str, str], cust) -> tuple[dict[str, str], list[str]]:
    """Company-name-as-printed -> customer_key. Exact on normalized name, then a
    prefix match for names the table clipped (>=12 chars, must be unambiguous)."""
    by_norm = {norm(r["name"]): r["customer_key"] for r in cust}
    keys, misses = {}, []
    for co, lc in printed.items():
        n = norm(co)
        k = by_norm.get(n)
        if not k and len(n) >= 12:
            hits = [v for kk, v in by_norm.items() if kk.startswith(n)]
            if len(hits) == 1:
                k = hits[0]
        if k:
            keys[k] = lc
        else:
            misses.append(co)
    return keys, misses


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--pdf", type=Path, action="append", required=True,
                    help="a previously rendered report (repeatable; they merge, newest date wins)")
    ap.add_argument("--db", type=Path, default=DEFAULT_DB)
    ap.add_argument("--scope", default="all", help="snapshot scope key (default 'all')")
    ap.add_argument("--dry-run", action="store_true", help="report what would be written, write nothing")
    ap.add_argument("--force", action="store_true", help="overwrite an existing snapshot for that date")
    a = ap.parse_args()

    if not a.db.exists():
        print(f"ledger not found: {a.db}", file=sys.stderr)
        return 2

    con = sqlite3.connect(f"file:{a.db}?mode=ro", uri=True)
    con.row_factory = sqlite3.Row
    cust = con.execute("SELECT customer_key, name, sales_status, last_contacted FROM customer").fetchall()
    con.close()

    merged: dict[str, str] = {}
    as_of: dt.date | None = None
    for pdf in a.pdf:
        if not pdf.exists():
            print(f"  !! missing: {pdf}", file=sys.stderr)
            return 2
        try:
            when, printed = parse_report(pdf)
        except ValueError as e:
            print(f"  -- skipped: {e}")
            continue
        keys, misses = match_to_ledger(printed, cust)
        print(f"  {pdf.name}: as of {when:%m/%d/%Y}, {len(printed)} listed, "
              f"{len(keys)} matched, {len(misses)} unmatched")
        for m in misses:
            print(f"      unmatched: {m}")
        as_of = when if as_of is None or when > as_of else as_of
        merged.update(keys)

    if as_of is None:
        print("no usable report supplied - nothing written", file=sys.stderr)
        return 1

    # Accounts never on those reports (another rep's book) must not read as NEW.
    # Baseline them at their current date so they show as neither new nor fresh;
    # only an account whose contact date moved past the report date can surface.
    stamp = as_of.isoformat()
    filled = 0
    for r in cust:
        if r["sales_status"] in TOUCHED and r["customer_key"] not in merged:
            merged[r["customer_key"]] = r["last_contacted"] or ""
            filled += 1

    moved = sum(1 for r in cust if r["sales_status"] in TOUCHED
                and (r["last_contacted"] or "") > (merged.get(r["customer_key"]) or ""))
    print(f"\n  baseline {stamp}: {len(merged)} accounts "
          f"({len(merged)-filled} from the reports, {filled} carried at current date)")
    print(f"  -> next render will show {moved} with fresh outreach since {as_of:%m/%d/%Y}")

    dest = SNAP_DIR / f"{a.scope}-{stamp}.json"
    if a.dry_run:
        print(f"\n  dry run - would write {dest}")
        return 0
    if dest.exists() and not a.force:
        print(f"\n  refusing to overwrite {dest} (a recorded snapshot beats a reconstructed "
              f"one) - pass --force to override", file=sys.stderr)
        return 1
    SNAP_DIR.mkdir(parents=True, exist_ok=True)
    dest.write_text(json.dumps({
        "date": stamp, "rep": None, "accounts": merged,
        "provenance": "reconstructed from rendered reports: "
                      + ", ".join(p.name for p in a.pdf),
    }), encoding="utf-8")
    print(f"\n  wrote {dest}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
