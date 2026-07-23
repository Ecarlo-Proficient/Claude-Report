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
MOR_PATH = CH / "Money Out Register.xlsx"
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


def read_money_out(path: Path) -> dict:
    """Uncashed-check KPIs from the Money Out Register's Summary sheet."""
    out: Dict[str, Any] = {}
    if not path.exists():
        return out
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception:
        return out
    if "Summary" in wb.sheetnames:
        for row in wb["Summary"].iter_rows(min_row=3, values_only=True):
            lbl = str(row[1]).strip() if len(row) > 1 and row[1] else ""
            val = row[2] if len(row) > 2 else None
            lo = lbl.lower()
            if "aged > 30" in lo and "$" in lbl:
                out["aged_total"] = _num(val)
            elif "aged > 30" in lo:
                out["aged_count"] = _num(val)
            elif lo.startswith("unmarked checks"):
                out["unmarked_count"] = _num(val)
            elif lo.startswith("unmarked $"):
                out["unmarked_total"] = _num(val)
    wb.close()
    return out


BUCKET_ORDER = ["Current", "1-30", "31-60", "61-90", "90+"]


def read_health(path: Path) -> dict:
    """Cash total + per-account and AR/AP grand totals + aging buckets from the
    legacy health_dashboard.xlsx (qbo_health.py output). Carries its own
    'generated' date — this workbook is often stale, so the dashboard dates it."""
    out: Dict[str, Any] = {"generated": None, "cash_total": None, "accounts": [],
                           "ar_total": None, "ar_buckets": {},
                           "ap_total": None, "ap_buckets": {}}
    if not path.exists():
        return out
    wb = load_workbook(path, read_only=True, data_only=True)
    if "_Meta" in wb.sheetnames:
        for row in wb["_Meta"].iter_rows(values_only=True):
            if row and str(row[0]).strip() == "Generated" and len(row) > 1:
                try:
                    out["generated"] = dt.datetime.fromisoformat(str(row[1]))
                except ValueError:
                    pass
    if "Cash" in wb.sheetnames:
        tot = 0.0
        for row in wb["Cash"].iter_rows(min_row=2, values_only=True):
            name, bal = (row[0] if row else None), (_num(row[3]) if len(row) > 3 else None)
            if name and bal is not None:
                out["accounts"].append((str(name), bal))
                tot += bal
        out["cash_total"] = tot
    for sheet, tkey, bkey in (("AR Aging", "ar_total", "ar_buckets"),
                              ("AP Aging", "ap_total", "ap_buckets")):
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        it = ws.iter_rows(values_only=True)
        hdr = list(next(it))
        hi = {n: i for i, n in enumerate(hdr)}
        bi, bal_i = hi.get("Bucket"), hi.get("Balance")
        name_i = hi.get("Customer", hi.get("Vendor"))
        buckets: Dict[str, float] = {}
        for row in it:
            nm = str(row[name_i] or "") if name_i is not None else ""
            if nm.startswith("━"):
                continue
            if "GRAND TOTAL" in nm:
                out[tkey] = _num(row[bal_i]) or 0
                continue
            b = row[bi] if bi is not None else None
            v = _num(row[bal_i]) if bal_i is not None else None
            if b and v:
                buckets[str(b)] = buckets.get(str(b), 0) + v
        out[bkey] = buckets
    wb.close()
    return out


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


def _bucket_bars(buckets: dict, worst_first=True) -> str:
    if not buckets:
        return "<p class='muted'>—</p>"
    mx = max(buckets.values(), default=1) or 1
    # older buckets redder
    shade = {"Current": "#1F6B4C", "1-30": "#7F9F3F", "31-60": "#BF8F00",
             "61-90": "#D2691E", "90+": "#C00000"}
    rows = []
    for b in BUCKET_ORDER:
        v = buckets.get(b, 0)
        if not v:
            continue
        w = max(2, round(100 * v / mx))
        col = shade.get(b, "#808080")
        rows.append(f'''<div class="bar-row bkt">
          <div class="bar-name" style="color:{col}">{b}</div>
          <div class="bar-track"><div class="bar-fill" style="width:{w}%;background:{col}">
            <span>${v:,.0f}</span></div></div></div>''')
    return "".join(rows)


def _health_section(h: dict) -> str:
    if not h.get("generated") and h.get("cash_total") is None:
        return "<p class='muted'>No health_dashboard.xlsx — run qbo_health.py.</p>"
    gen = h.get("generated")
    age = (dt.datetime.now() - gen).days if gen else None
    stale = age is not None and age > STALE_DAYS
    asof = (f'<span class="asof {"stale" if stale else ""}">as of '
            f'{gen:%Y-%m-%d} ({age}d old)</span>' if gen else "")
    warn = ('<p class="warnbar">⚠ This cash/aging snapshot is stale — run '
            'qbo_health.py to refresh.</p>' if stale else "")
    cash = h.get("cash_total")
    cash_cls = "red" if (cash is not None and cash < 0) else "info"
    top = sorted([a for a in h.get("accounts", [])], key=lambda x: x[1])[:3]
    top_html = "".join(f'<div class="acct"><span>{html.escape(n)}</span>'
                       f'<span>{_money(b)}</span></div>' for n, b in top)
    return f'''<h2>Cash &amp; Aging {asof}</h2>{warn}
      <div class="grid">
        <div class="tile {cash_cls}"><div class="count">{_money(cash)}</div>
          <div class="label">Total cash (all bank accounts)</div>
          <div class="detail">{top_html}</div></div>
        <div class="tile info"><div class="count">{_money(h.get("ar_total"))}</div>
          <div class="label">Accounts Receivable</div>
          <div class="detail">money owed to us</div></div>
        <div class="tile amber"><div class="count">{_money(h.get("ap_total"))}</div>
          <div class="label">Accounts Payable</div>
          <div class="detail">money we owe</div></div>
      </div>
      <div class="two-col">
        <div><h3 class="sec">AR aging</h3>{_bucket_bars(h.get("ar_buckets", {}))}</div>
        <div><h3 class="sec">AP aging</h3>{_bucket_bars(h.get("ap_buckets", {}))}</div>
      </div>'''


def _money_out_tile(mor: dict) -> str:
    if not mor:
        return ""
    ac = int(mor.get("aged_count") or 0)
    return f'''<div class="grid"><div class="tile {'red' if ac else 'ok'}">
      <div class="count">{ac}</div>
      <div class="label">Uncashed checks aged &gt;30d — chase list</div>
      <div class="amt">{_money(mor.get("aged_total"))}</div>
      <div class="detail">{int(mor.get("unmarked_count") or 0)} unmarked total
        ({_money(mor.get("unmarked_total"))}) · mark cleared in the Money Out
        Register as checks clear</div></div></div>'''


def render(cards, loc, health, mor, sources) -> str:
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
  .asof {{ font-size:12px; font-weight:400; text-transform:none; letter-spacing:0;
    color:var(--muted); margin-left:8px; }}
  .asof.stale {{ color:var(--red); font-weight:600; }}
  .warnbar {{ background:rgba(192,0,0,.12); color:var(--red); padding:8px 12px;
    border-radius:8px; font-size:13px; margin:0 0 12px; }}
  .acct {{ display:flex; justify-content:space-between; gap:10px; font-size:11px;
    margin-top:3px; }}
  .two-col {{ display:grid; grid-template-columns:1fr 1fr; gap:24px; margin-top:6px; }}
  @media (max-width:720px) {{ .two-col {{ grid-template-columns:1fr; }} }}
  .bar-row.bkt {{ grid-template-columns:60px 1fr; }}
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
  {_health_section(health)}
  {_money_out_tile(mor)}
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
    health = read_health(HEALTH_PATH)
    print(f"  Cash/Aging: cash={health.get('cash_total')} AR={health.get('ar_total')} "
          f"AP={health.get('ap_total')} (as of {health.get('generated')})")
    mor = read_money_out(MOR_PATH)
    print(f"  Money Out: aged>30d checks={mor.get('aged_count')} ${mor.get('aged_total')}")

    htmlout = render(cards, loc, health, mor,
                     [MB_PATH, LOC_PATH, MOR_PATH, BILL_PATH, HEALTH_PATH])
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(htmlout, encoding="utf-8")
    os.chmod(args.out, stat.S_IRUSR | stat.S_IWUSR)
    print(f"\n  ✓ {args.out}  (chmod 600)")
    if args.open:
        webbrowser.open(f"file://{args.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
