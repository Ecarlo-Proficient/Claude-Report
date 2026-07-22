#!/usr/bin/env python3
"""
company_dashboard.py — one consolidated company-health view.

Reads the tracker workbooks that the other tools already generate (they are the
data layer) and renders ONE self-contained HTML page — no QBO calls, no Touch
ID, no server, works offline (the user 2026-07-17: "excel with html dashboard
reading all"). Regenerate the trackers first, then run this.

Sources (all read-only; each shown with its freshness):
  • Money Bleeds.xlsx   — draws-not-invoiced, lien clock, unused POs, open AP
  • Sub LOC Report.xlsx — subcontractor LOC peak + by-division float
  • Bill Tracker.xlsx / health_dashboard.xlsx — freshness only (folded in later)

OUTPUT
  ~/Documents/CompanyHealth/Company Dashboard.html  (chmod 600)

USAGE
  python3 health-dashboard/company_dashboard.py
  python3 health-dashboard/company_dashboard.py --open   # open in the browser
"""
from __future__ import annotations

import argparse
import datetime as dt
import html
import os
import re
import stat
import sys
import webbrowser
from pathlib import Path
from typing import Any, Dict, List, Optional

from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from shared import paths

CH = paths.companyhealth_dir()
MB_PATH = CH / "Money Bleeds.xlsx"
LOC_PATH = CH / "Sub LOC Report.xlsx"
BILL_PATH = CH / "Bill Tracker.xlsx"
HEALTH_PATH = CH / "health_dashboard.xlsx"
OUT_PATH = CH / "Company Dashboard.html"
STALE_DAYS = 7


# ────────────────────────── source reading ──────────────────────────

def _mtime(p: Path) -> Optional[dt.datetime]:
    try:
        return dt.datetime.fromtimestamp(p.stat().st_mtime)
    except OSError:
        return None


def _num(v) -> Optional[float]:
    if isinstance(v, (int, float)):
        return float(v)
    return None


def read_money_bleeds(path: Path) -> List[dict]:
    """Parse the Money Bleeds 'Dashboard' sheet into sections of cards.
    Each card: {section, label, count, amount, detail, severity}."""
    if not path.exists():
        return []
    wb = load_workbook(path, read_only=True, data_only=True)
    if "Dashboard" not in wb.sheetnames:
        wb.close()
        return []
    ws = wb["Dashboard"]
    cards: List[dict] = []
    section = ""
    for row in ws.iter_rows(min_row=2, values_only=True):
        b = row[1] if len(row) > 1 else None
        c = row[2] if len(row) > 2 else None
        d = row[3] if len(row) > 3 else None
        if not b or not str(b).strip():
            continue
        label = str(b).strip()
        if label.startswith(("Generated", "RP data")):    # subtitles — skip
            continue
        # a section header: the part before "—"/double-space is ALL CAPS and
        # doesn't itself end in a count (guards against "… synced 14:08")
        head = re.split(r"—|  ", label)[0].strip()
        if head.isupper() and not re.search(r":\s*\d+\s*$", head):
            section = head
            continue
        m = re.search(r":\s*(\d+)\s*$", label)
        if m:                                   # a KPI card ("… : N")
            cards.append({
                "section": section, "label": label, "count": int(m.group(1)),
                "amount": _num(c), "detail": str(d or "").strip(),
                "severity": _severity(section, label, int(m.group(1))),
            })
    wb.close()
    return cards


def _severity(section: str, label: str, count: int) -> str:
    lo = label.lower()
    if count == 0:
        return "ok"
    if any(k in lo for k in ("past", "urgent", "no invoice", "under-invoiced",
                             "coming up", "hasn't paid")):
        return "red"
    if any(k in lo for k in ("watch", "review", "pending", "not yet lien",
                             "unused po", "waiting on punch", "no invoice issued")):
        return "amber"
    return "info"


def read_sub_loc(path: Path) -> dict:
    """Return {summary:{label:value}, divisions:[{name,peak,drawn,outstanding,
    float}]}."""
    out: Dict[str, Any] = {"summary": {}, "divisions": []}
    if not path.exists():
        return out
    wb = load_workbook(path, read_only=True, data_only=True)
    if "Summary" in wb.sheetnames:
        for row in wb["Summary"].iter_rows(min_row=3, values_only=True):
            lbl = row[1] if len(row) > 1 else None
            val = row[2] if len(row) > 2 else None
            if lbl and val not in (None, ""):
                out["summary"][str(lbl).strip()] = val
    if "By Division" in wb.sheetnames:
        for row in wb["By Division"].iter_rows(min_row=2, values_only=True):
            name = row[0] if row else None
            if not name or str(name).strip() in ("TOTAL", ""):
                continue
            if str(name).strip().startswith("Note"):
                continue
            out["divisions"].append({
                "name": str(name).strip(),
                "peak": _num(row[1]) or 0, "drawn": _num(row[3]) or 0,
                "outstanding": _num(row[6]) or 0, "float": _num(row[7]) or 0,
            })
    wb.close()
    return out


def _find(summary: dict, needle: str) -> Optional[Any]:
    for k, v in summary.items():
        if needle.lower() in k.lower():
            return v
    return None


# ────────────────────────── HTML render ──────────────────────────

def _money(v) -> str:
    n = _num(v)
    return f"${n:,.0f}" if n is not None else (html.escape(str(v)) if v else "—")


def _fresh_badge(p: Path) -> str:
    mt = _mtime(p)
    if not mt:
        return f'<span class="badge missing">{html.escape(p.name)} · not generated</span>'
    age = (dt.datetime.now() - mt).days
    cls = "stale" if age > STALE_DAYS else "fresh"
    return (f'<span class="badge {cls}">{html.escape(p.name)} · '
            f'{mt:%Y-%m-%d %H:%M} ({age}d)</span>')


def _tile(card: dict) -> str:
    amt = f'<div class="amt">{_money(card["amount"])}</div>' if card["amount"] else ""
    label = html.escape(re.sub(r":\s*\d+\s*$", "", card["label"]))
    return f'''<div class="tile {card['severity']}">
      <div class="count">{card['count']}</div>
      <div class="label">{label}</div>
      {amt}
      <div class="detail">{html.escape(card['detail'])}</div>
    </div>'''


def _division_bars(divs: List[dict]) -> str:
    if not divs:
        return "<p class='muted'>No LOC data — run sub_loc_report.py.</p>"
    mx = max((d["peak"] for d in divs), default=1) or 1
    colors = {"MFD": "#305496", "CP": "#1F6B4C", "RP": "#7030A0", "Other": "#808080"}
    rows = []
    for d in sorted(divs, key=lambda x: -x["peak"]):
        w = max(2, round(100 * d["peak"] / mx))
        col = colors.get(d["name"], "#808080")
        rows.append(f'''<div class="bar-row">
          <div class="bar-name" style="color:{col}">{html.escape(d['name'])}</div>
          <div class="bar-track"><div class="bar-fill" style="width:{w}%;background:{col}">
            <span>${d['peak']:,.0f}</span></div></div>
          <div class="bar-meta">{d['float']:.0f}d float · ${d['outstanding']:,.0f} out</div>
        </div>''')
    return "".join(rows)


def render(cards, loc, sources) -> str:
    now = dt.datetime.now()
    # group Money Bleeds cards by section
    sections: Dict[str, List[dict]] = {}
    for c in cards:
        sections.setdefault(c["section"] or "OTHER", []).append(c)
    mb_html = ""
    for sec, cs in sections.items():
        mb_html += (f'<h3 class="sec">{html.escape(sec)}</h3>'
                    f'<div class="grid">{"".join(_tile(c) for c in cs)}</div>')
    if not mb_html:
        mb_html = "<p class='muted'>No Money Bleeds data — run money_bleeds.py.</p>"

    peak = _find(loc["summary"], "LOC you truly need")
    avg = _find(loc["summary"], "Avg days our cash")
    out_ = _find(loc["summary"], "Still outstanding")
    loc_tiles = f'''<div class="grid">
      <div class="tile red"><div class="count">{_money(peak)}</div>
        <div class="label">LOC you truly need (peak)</div>
        <div class="detail">company-wide high-water balance</div></div>
      <div class="tile amber"><div class="count">{avg if avg is not None else '—'}</div>
        <div class="label">Avg days cash is out</div>
        <div class="detail">draw → client repayment</div></div>
      <div class="tile info"><div class="count">{_money(out_)}</div>
        <div class="label">Still outstanding</div>
        <div class="detail">drawn, not yet repaid</div></div>
    </div>'''

    badges = " ".join(_fresh_badge(p) for p in sources)

    return f'''<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Company Dashboard</title>
<style>
  :root {{ --bg:#f4f6fa; --card:#fff; --ink:#1f2937; --muted:#6b7280;
           --navy:#1F3864; --red:#C00000; --amber:#BF8F00; --green:#1F6B4C; }}
  @media (prefers-color-scheme: dark) {{
    :root {{ --bg:#0f1420; --card:#1a2130; --ink:#e5e9f0; --muted:#9aa4b2; }}
  }}
  * {{ box-sizing:border-box; }}
  body {{ margin:0; background:var(--bg); color:var(--ink);
    font:15px/1.45 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif; }}
  header {{ background:var(--navy); color:#fff; padding:20px 28px; }}
  header h1 {{ margin:0; font-size:22px; }}
  header .sub {{ opacity:.8; font-size:13px; margin-top:4px; }}
  main {{ max-width:1180px; margin:0 auto; padding:24px 28px 60px; }}
  h2 {{ font-size:16px; letter-spacing:.04em; text-transform:uppercase;
    color:var(--muted); border-bottom:2px solid rgba(128,128,128,.2);
    padding-bottom:6px; margin:34px 0 14px; }}
  h3.sec {{ font-size:12px; letter-spacing:.06em; color:var(--muted);
    margin:18px 0 8px; }}
  .grid {{ display:grid; grid-template-columns:repeat(auto-fill,minmax(215px,1fr));
    gap:12px; }}
  .tile {{ background:var(--card); border-radius:10px; padding:14px 16px;
    border-left:5px solid var(--muted); box-shadow:0 1px 3px rgba(0,0,0,.08); }}
  .tile.red {{ border-left-color:var(--red); }}
  .tile.amber {{ border-left-color:var(--amber); }}
  .tile.ok {{ border-left-color:var(--green); }}
  .tile.info {{ border-left-color:var(--navy); }}
  .tile .count {{ font-size:26px; font-weight:700; }}
  .tile.red .count {{ color:var(--red); }} .tile.amber .count {{ color:var(--amber); }}
  .tile.ok .count {{ color:var(--green); }}
  .tile .amt {{ font-size:17px; font-weight:600; margin-top:2px; }}
  .tile .label {{ font-size:13px; margin-top:4px; }}
  .tile .detail {{ font-size:11px; color:var(--muted); margin-top:6px; }}
  .bar-row {{ display:grid; grid-template-columns:60px 1fr 190px;
    align-items:center; gap:12px; margin:8px 0; }}
  .bar-name {{ font-weight:700; }}
  .bar-track {{ background:rgba(128,128,128,.15); border-radius:6px; height:26px; }}
  .bar-fill {{ height:26px; border-radius:6px; display:flex; align-items:center;
    justify-content:flex-end; padding-right:8px; color:#fff; font-size:12px;
    font-weight:600; min-width:40px; }}
  .bar-meta {{ font-size:12px; color:var(--muted); }}
  .muted {{ color:var(--muted); }}
  footer {{ max-width:1180px; margin:0 auto; padding:0 28px 40px; }}
  .badge {{ display:inline-block; font-size:11px; padding:4px 9px; border-radius:20px;
    margin:3px 4px 0 0; background:rgba(128,128,128,.15); color:var(--muted); }}
  .badge.stale {{ background:rgba(192,0,0,.15); color:var(--red); font-weight:600; }}
  .badge.fresh {{ background:rgba(31,107,76,.15); color:var(--green); }}
  .badge.missing {{ background:rgba(191,143,0,.18); color:var(--amber); }}
</style></head><body>
<header>
  <h1>Company Dashboard</h1>
  <div class="sub">Generated {now:%Y-%m-%d %H:%M} · reads the tracker workbooks in
    ~/Documents/CompanyHealth · regenerate those first for fresh numbers</div>
</header>
<main>
  <h2>Money Bleeds</h2>
  {mb_html}
  <h2>Subcontractor LOC</h2>
  {loc_tiles}
  <h3 class="sec">Peak LOC by division (bar = peak $, judge the need per division)</h3>
  {_division_bars(loc["divisions"])}
</main>
<footer>
  <h2>Sources &amp; freshness</h2>
  {badges}
</footer>
</body></html>'''


def main() -> int:
    ap = argparse.ArgumentParser(description="Consolidated company dashboard (HTML)")
    ap.add_argument("--out", type=Path, default=OUT_PATH)
    ap.add_argument("--open", action="store_true", help="open in the browser after writing")
    args = ap.parse_args()

    print("\n  COMPANY DASHBOARD — reading the tracker workbooks")
    print("  " + "─" * 52)
    cards = read_money_bleeds(MB_PATH)
    print(f"  Money Bleeds: {len(cards)} KPI card(s)")
    loc = read_sub_loc(LOC_PATH)
    print(f"  Sub LOC: peak={_find(loc['summary'], 'LOC you truly need')} · "
          f"{len(loc['divisions'])} division(s)")

    htmlout = render(cards, loc, [MB_PATH, LOC_PATH, BILL_PATH, HEALTH_PATH])
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(htmlout, encoding="utf-8")
    os.chmod(args.out, stat.S_IRUSR | stat.S_IWUSR)
    print(f"\n  ✓ {args.out}  (chmod 600)")
    if args.open:
        webbrowser.open(f"file://{args.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
