#!/usr/bin/env python3
"""
sub_loc_report.py — Subcontractor line-of-credit (LOC) float model.

The question (the user 2026-07-17): if every subcontractor payment were funded
by a line of credit and repaid when the client pays us for that work, how big
a LOC do we truly need, and how long is our money out before it comes back?

MODEL (all four choices are the user's, 2026-07-17):
  • Sub bill = a QBO Bill whose memo/PrivateNote contains "sub" (same rule as
    bill-tracker's is_sub_bill). Worked at the LINE level — each line carries
    its project on CustomerRef; one bill can span several projects.
  • DRAW (money out) = when the sub is actually PAID — the QBO BillPayment
    date — allocated across the bill's lines pro-rata by line amount.
  • REPAY (money in) = when the client pays us — the QBO customer Payment
    date, mapped to a project through the invoice it was applied to.
  • Matching = per-project FIFO: a client payment repays the oldest still-out
    sub draws on that project first, only up to what was drawn (the margin
    stays as profit and does NOT pay down the LOC).
  • Window = first Friday of the month 3 months back → today; balance starts 0.

Running LOC balance = cumulative draws − cumulative applied repayments over
time; its PEAK is the LOC you truly need. Averages: amount-weighted draw→repay
lag, average draw, average repayment.

READ-ONLY against QBO. Output: ~/Documents/CompanyHealth/Sub LOC Report.xlsx
(chmod 600). One Touch ID per run.

USAGE
  python3 one-offs/sub_loc_report.py
  python3 one-offs/sub_loc_report.py --months 3
  python3 one-offs/sub_loc_report.py --start 2026-04-03 --out /path/x.xlsx
"""
from __future__ import annotations

import argparse
import datetime as dt
import os
import re
import stat
import sys
from collections import defaultdict
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.formatting.rule import DataBarRule
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from shared import qbo_api
from shared import paths

SUB_RE = re.compile(r"\bsub\b", re.IGNORECASE)     # same as bill-tracker
DEFAULT_OUTPUT = paths.companyhealth_dir() / "Sub LOC Report.xlsx"
# pull bills/invoices this far back so txns paid inside the window are covered
LOOKBACK_MONTHS = 9


# ────────────────────────── dates ──────────────────────────

def _today() -> dt.date:
    return dt.date.today()


def first_friday_months_back(today: dt.date, months: int) -> dt.date:
    """First Friday of the month `months` before `today`."""
    y, m = today.year, today.month - months
    while m <= 0:
        m += 12
        y -= 1
    d = dt.date(y, m, 1)
    return d + dt.timedelta(days=(4 - d.weekday()) % 7)   # 4 = Friday


def _parse(s: str) -> Optional[dt.date]:
    try:
        return dt.date.fromisoformat(s)
    except (TypeError, ValueError):
        return None


_PERIOD_RE = re.compile(
    r"(\d{1,2})/(\d{1,2})/(\d{2,4})\s*[-–]\s*(\d{1,2})/(\d{1,2})/(\d{2,4})")


def _mdy(m: str, d: str, y: str) -> Optional[dt.date]:
    try:
        yr = int(y)
        yr += 2000 if yr < 100 else 0
        return dt.date(yr, int(m), int(d))
    except ValueError:
        return None


def period_month(text: str) -> Optional[str]:
    """A draw/work period like '7/01/2026 - 7/31/2026' → the 'YYYY-MM' month
    containing its MIDPOINT (draws straddle month boundaries, e.g. 1/26–2/25 =
    the February draw). Returns None when no period range is present."""
    m = _PERIOD_RE.search(text or "")
    if not m:
        return None
    a = _mdy(m.group(1), m.group(2), m.group(3))
    b = _mdy(m.group(4), m.group(5), m.group(6))
    if not a or not b:
        return None
    mid = a + (b - a) / 2
    return f"{mid.year:04d}-{mid.month:02d}"


def _month_of(d: Optional[dt.date]) -> Optional[str]:
    return f"{d.year:04d}-{d.month:02d}" if d else None


# ────────────────────────── QBO pulls ──────────────────────────

def _line_project(line: dt) -> Optional[str]:
    for k in ("AccountBasedExpenseLineDetail", "ItemBasedExpenseLineDetail"):
        det = line.get(k) or {}
        name = (det.get("CustomerRef") or {}).get("name")
        if name:
            return qbo_api.extract_proj(name)
    return None


def build_sub_bill_lines(access, cid, since: str) -> Dict[str, dict]:
    """{billId: {vendor, lines:[(project, amount)], total, work_month}} for sub
    bills. work_month = the month of the memo's 'Period …' (the draw period the
    sub cost belongs to), else the bill-date month."""
    out: Dict[str, dict] = {}
    for b in qbo_api.query_all(access, cid, "Bill", f"TxnDate >= '{since}'"):
        memo = b.get("PrivateNote") or ""
        if not SUB_RE.search(memo):
            continue
        lines = []
        for ln in b.get("Line") or []:
            amt = float(ln.get("Amount") or 0)
            if amt <= 0:
                continue
            lines.append((_line_project(ln), amt))
        if not lines:
            continue
        wm = period_month(memo) or _month_of(_parse(b.get("TxnDate")))
        out[b["Id"]] = {
            "vendor": (b.get("VendorRef") or {}).get("name") or "",
            "lines": lines,
            "total": sum(a for _, a in lines),
            "work_month": wm,
        }
    return out


def build_invoice_meta(access, cid, since: str) -> Dict[str, dict]:
    """{invoiceId: {project, draw_month}}. draw_month = the month of the
    invoice memo's 'Draw #N (Period …)', else the invoice-date month (RP lump
    invoices have no period → their draw_month is unused, RP matches by
    project only)."""
    out = {}
    for i in qbo_api.query_all(access, cid, "Invoice", f"TxnDate >= '{since}'"):
        memo = ((i.get("PrivateNote") or "") + " " +
                ((i.get("CustomerMemo") or {}).get("value") or ""))
        out[i["Id"]] = {
            "project": qbo_api.extract_proj((i.get("CustomerRef") or {}).get("name") or ""),
            "draw_month": period_month(memo) or _month_of(_parse(i.get("TxnDate"))),
            "doc": str(i.get("DocNumber") or ""),
        }
    return out


def collect_draws(access, cid, sub_lines: Dict[str, dict],
                  start: dt.date) -> List[dict]:
    """Money OUT: each BillPayment that pays a sub bill → per-project draw
    events, allocated pro-rata across the bill's lines."""
    draws = []
    since = start.isoformat()
    for bp in qbo_api.query_all(access, cid, "BillPayment", f"TxnDate >= '{since}'"):
        d = _parse(bp.get("TxnDate"))
        if d is None or d < start:
            continue
        vendor = (bp.get("VendorRef") or {}).get("name") or ""
        for ln in bp.get("Line") or []:
            paid = float(ln.get("Amount") or 0)
            if paid <= 0:
                continue
            for lk in ln.get("LinkedTxn") or []:
                if lk.get("TxnType") != "Bill":
                    continue
                bill = sub_lines.get(lk.get("TxnId"))
                if not bill or bill["total"] <= 0:
                    continue
                for proj, line_amt in bill["lines"]:
                    draws.append({
                        "date": d, "project": proj or "(no project)",
                        "party": bill["vendor"] or vendor,
                        "amount": paid * line_amt / bill["total"],
                        "period": bill["work_month"],
                    })
    draws.sort(key=lambda x: x["date"])
    return draws


def collect_repays(access, cid, inv_meta: Dict[str, dict],
                   start: dt.date) -> List[dict]:
    """Money IN: each customer Payment → per-project repayment events, mapped
    through the invoice it was applied to (which carries the draw month)."""
    repays = []
    since = start.isoformat()
    for p in qbo_api.query_all(access, cid, "Payment", f"TxnDate >= '{since}'"):
        d = _parse(p.get("TxnDate"))
        if d is None or d < start:
            continue
        client = (p.get("CustomerRef") or {}).get("name") or ""
        for ln in p.get("Line") or []:
            amt = float(ln.get("Amount") or 0)
            for lk in ln.get("LinkedTxn") or []:
                if lk.get("TxnType") != "Invoice":
                    continue
                meta = inv_meta.get(lk.get("TxnId")) or {}
                repays.append({
                    "date": d, "project": meta.get("project") or "(no project)",
                    "party": client, "amount": amt if amt > 0 else 0,
                    "period": meta.get("draw_month"),
                    "invoice": meta.get("doc") or "",
                })
    repays.sort(key=lambda x: x["date"])
    return repays


# ────────────────────────── FIFO model ──────────────────────────

def run_fifo(draws: List[dict], repays: List[dict]) -> Tuple[List[dict], dict]:
    """Chronological per-project float model. Events processed in date order:
      • DRAW: first consume any client cash already received on that project
        (prefunded — the GC draw came in before we paid the sub, common on MFD);
        only the UN-prefunded remainder hits the LOC and joins the FIFO queue.
      • REPAY: pay down the oldest outstanding LOC draws first (that's the
        draw→repay float); any leftover becomes prefunding for future draws
        (capped there — the margin above sub cost never funds more than a
        later draw needs). Running LOC balance never goes negative.
    Peak running balance = the LOC we truly need. Lag average is over
    LOC-funded draws only (prefunded ones had zero money-out time)."""
    eps = 1e-6
    merged = ([{"kind": "D", **d} for d in draws]
              + [{"kind": "R", **r} for r in repays])
    # same day: DRAW before REPAY → conservative (higher) peak
    merged.sort(key=lambda e: (e["date"], 0 if e["kind"] == "D" else 1))

    # Match bucket: MFD/CP reimburse BY DRAW PERIOD — a draw only offsets sub
    # costs of the SAME (project, month). RP invoices are lump per scope (no
    # draw period) → match by project only. This is what stops a May draw from
    # phantom-'prefunding' June sub costs (the user 2026-07-17).
    def bucket(e) -> Tuple[str, str]:
        proj = e["project"]
        if proj == "(no project)" or proj.startswith("RP"):
            return (proj, "*")
        return (proj, e.get("period") or "*")

    queues: Dict[Tuple[str, str], List[dict]] = defaultdict(list)
    prepaid: Dict[Tuple[str, str], float] = defaultdict(float)
    events: List[dict] = []
    lags: List[Tuple[float, int]] = []
    applied_total = prefunded_total = 0.0
    bal = peak = 0.0
    peak_date = None

    for e in merged:
        p = bucket(e)
        if e["kind"] == "D":
            amt = e["amount"]
            use_pre = min(amt, prepaid[p])
            prepaid[p] -= use_pre
            prefunded_total += use_pre
            loc_part = amt - use_pre
            note = f"prefunded ${use_pre:,.0f}" if use_pre > eps else ""
            ev = {"date": e["date"], "type": "DRAW", "project": e["project"],
                  "party": e["party"], "out": amt, "inn": 0.0, "lag": None,
                  "note": note, "balance": bal, "reimb": []}
            if loc_part > eps:
                # node carries a ref to its ledger row so repayments can stamp
                # the reimbursing invoice + client-paid date back onto the draw
                queues[p].append({"date": e["date"], "remaining": loc_part, "ev": ev})
                bal += loc_part
            ev["balance"] = bal
            events.append(ev)
        else:
            R = e["amount"]
            applied_here = 0.0
            first_lag = None
            q = queues[p]
            while R > eps and q and q[0]["remaining"] > eps:
                node = q[0]
                take = min(R, node["remaining"])
                node["remaining"] -= take
                R -= take
                applied_here += take
                applied_total += take
                days = (e["date"] - node["date"]).days
                lags.append((take, days))
                if first_lag is None:
                    first_lag = days
                # stamp the reimbursing invoice + client-paid date on the draw
                node["ev"]["reimb"].append((e.get("invoice", ""), e["date"]))
                if node["remaining"] <= eps:
                    q.pop(0)
            bal -= applied_here
            surplus = R if R > eps else 0.0
            prepaid[p] += surplus
            note = f"surplus ${surplus:,.0f}" if surplus > eps else ""
            events.append({"date": e["date"], "type": "REPAY", "project": e["project"],
                           "party": e["party"], "out": 0.0, "inn": applied_here,
                           "lag": first_lag, "note": note, "balance": bal,
                           "invoice": e.get("invoice", ""), "reimb": []})
        if bal > peak:
            peak, peak_date = bal, e["date"]

    total_drawn = sum(d["amount"] for d in draws)
    outstanding = sum(n["remaining"] for q in queues.values() for n in q)
    wl = sum(a * days for a, days in lags)
    wa = sum(a for a, _ in lags)
    summary = {
        "n_draws": len(draws), "n_repay_chunks": len(lags),
        "total_drawn": total_drawn, "total_repaid": applied_total,
        "prefunded": prefunded_total, "outstanding": outstanding,
        "peak": peak, "peak_date": peak_date,
        "avg_lag": (wl / wa) if wa else 0.0,
        "avg_draw": (total_drawn / len(draws)) if draws else 0.0,
        "avg_repay": (applied_total / len(lags)) if lags else 0.0,
    }
    return events, summary


def per_project(events: List[dict]) -> List[dict]:
    agg: Dict[str, dict] = defaultdict(lambda: {"out": 0.0, "inn": 0.0,
                                                "wl": 0.0, "wa": 0.0})
    for e in events:
        a = agg[e["project"]]
        a["out"] += e["out"]
        a["inn"] += e["inn"]
        if e["type"] == "REPAY" and e["lag"] is not None:
            a["wl"] += e["inn"] * e["lag"]
            a["wa"] += e["inn"]
    rows = []
    for proj, a in agg.items():
        rows.append({"project": proj, "drawn": a["out"], "repaid": a["inn"],
                     "outstanding": a["out"] - a["inn"],
                     "avg_lag": (a["wl"] / a["wa"]) if a["wa"] else 0.0})
    rows.sort(key=lambda r: -r["outstanding"])
    return rows


# ────────────────────────── Excel ──────────────────────────

CUR = '#,##0'
_NAVY = PatternFill("solid", fgColor="1F3864")
_ZEBRA = PatternFill("solid", fgColor="EEF3FA")
_DRAW_FILL = PatternFill("solid", fgColor="FCE4D6")     # money out (peach)
_REPAY_FILL = PatternFill("solid", fgColor="E2EFDA")    # money in (green)
_PEAK_FILL = PatternFill("solid", fgColor="C00000")
_WHITE = Font(bold=True, color="FFFFFF")
_THIN = Side(style="thin", color="D0D7E5")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)


def _hdr(ws, headers, widths):
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(1, c)
        cell.font = _WHITE
        cell.fill = _NAVY
        cell.border = _BORDER
        cell.alignment = Alignment(vertical="center", horizontal="center",
                                   wrap_text=True)
    for i, w in enumerate(widths):
        ws.column_dimensions[get_column_letter(i + 1)].width = w
    ws.freeze_panes = "A2"


def write_workbook(out: Path, events, summary, projects, start, today):
    wb = Workbook()

    # ── Summary ──
    s = wb.active
    s.title = "Summary"
    s.sheet_view.showGridLines = False
    s.column_dimensions["A"].width = 3
    s.column_dimensions["B"].width = 42
    s.column_dimensions["C"].width = 20
    s.append(["", "SUBCONTRACTOR LOC — float model"])
    s.merge_cells("B1:C1")
    s.cell(1, 2).font = Font(bold=True, size=15, color="FFFFFF")
    for c in (1, 2, 3):
        s.cell(1, c).fill = _NAVY
    s.row_dimensions[1].height = 32
    s.append(["", f"Window {start:%Y-%m-%d} → {today:%Y-%m-%d}  ·  sub bills paid "
              f"from LOC, repaid when the client pays  ·  per-project FIFO"])
    s.cell(2, 2).font = Font(italic=True, color="595959")
    s.merge_cells("B2:C2")

    def line(label, value, money=True, big=False, bad=False):
        s.append(["", label, value])
        r = s.max_row
        s.cell(r, 2).font = Font(bold=big, size=12 if big else 11)
        vc = s.cell(r, 3)
        vc.font = Font(bold=True, size=13 if big else 11,
                       color="C00000" if bad else ("1F6B4C" if big else "000000"))
        if money:
            vc.number_format = CUR
        for c in (2, 3):
            s.cell(r, c).border = _BORDER
            if big:
                s.cell(r, c).fill = PatternFill("solid", fgColor="FCE4E4" if bad
                                                else "E7F2E7")

    s.append([])
    line("LOC you truly need (peak balance)", round(summary["peak"]), big=True, bad=True)
    pk = summary["peak_date"]
    line("…reached on", pk.isoformat() if pk else "—", money=False)
    s.append([])
    line("Total drawn (paid to subs)", round(summary["total_drawn"]))
    line("…of which prefunded by the client first", round(summary["prefunded"]))
    line("Total repaid (client → LOC)", round(summary["total_repaid"]))
    line("Still outstanding (not yet repaid)", round(summary["outstanding"]), bad=True)
    s.append([])
    line("Avg days our cash is out (draw→repay)", round(summary["avg_lag"], 1),
         money=False, big=True)
    line("Avg draw (paid to a sub)", round(summary["avg_draw"]))
    line("Avg repayment chunk", round(summary["avg_repay"]))
    line("# draws", summary["n_draws"], money=False)
    line("# repayment chunks", summary["n_repay_chunks"], money=False)
    s.sheet_properties.tabColor = "1F3864"

    # ── Ledger (running balance) ──
    lg = wb.create_sheet("Ledger")
    _hdr(lg, ["DATE", "TYPE", "PROJECT #", "PARTY", "DRAW OUT $", "REPAY IN $",
              "RUNNING LOC $", "LAG (days)", "NOTE", "REIMBURSING INVOICE #",
              "INVOICE PAID DATE"],
         [12, 8, 12, 32, 14, 14, 16, 11, 16, 20, 15])
    first = 2
    peak_marked = False
    for i, e in enumerate(events):
        reimb = e.get("reimb") or []
        if e["type"] == "DRAW":
            invs = sorted({str(x[0]) for x in reimb if x[0]})
            inv_str = ", ".join(invs)
            paid = max((x[1] for x in reimb), default=None)
            paid_str = paid.isoformat() if paid else ("" if not e["out"] else "still out")
        else:
            inv_str = e.get("invoice", "")
            paid_str = e["date"].isoformat()
        lg.append([e["date"].isoformat(), e["type"], e["project"], e["party"],
                   round(e["out"]) if e["out"] else "",
                   round(e["inn"]) if e["inn"] else "",
                   round(e["balance"]), e["lag"] if e["lag"] is not None else "",
                   e["note"], inv_str, paid_str])
        r = lg.max_row
        for c in range(1, 12):
            lg.cell(r, c).border = _BORDER
            if i % 2:
                lg.cell(r, c).fill = _ZEBRA
        lg.cell(r, 2).fill = _DRAW_FILL if e["type"] == "DRAW" else _REPAY_FILL
        lg.cell(r, 2).font = Font(bold=True,
                                  color="9C6500" if e["type"] == "DRAW" else "1F6B4C")
        for c in (5, 6, 7):
            lg.cell(r, c).number_format = CUR
        if (not peak_marked and e["date"] == summary["peak_date"]
                and round(e["balance"]) == round(summary["peak"])):
            lg.cell(r, 7).fill = _PEAK_FILL
            lg.cell(r, 7).font = Font(bold=True, color="FFFFFF")
            peak_marked = True
    last = lg.max_row
    lg.conditional_formatting.add(
        f"G{first}:G{last}",
        DataBarRule(start_type="num", start_value=0, end_type="max",
                    color="8EAADB", showValue=True))
    lg.auto_filter.ref = f"A1:K{last}"
    # legend
    lg.append([])
    lg.append(["Legend:  prefunded = this draw was covered by client cash already "
               "received for the same project+draw-month (no LOC needed).  "
               "surplus = the client paid more than we had drawn that period "
               "(our margin / advance toward the next draw).  REIMBURSING INVOICE "
               "= the client invoice(s) whose payment repaid this draw."])
    lg.cell(lg.max_row, 1).font = Font(italic=True, color="808080")
    lg.sheet_properties.tabColor = "C00000"

    # ── Per-Project ──
    pp = wb.create_sheet("Per-Project")
    _hdr(pp, ["PROJECT #", "DRAWN (paid subs) $", "REPAID $",
              "OUTSTANDING $", "AVG DAYS draw→repay"], [14, 20, 16, 16, 20])
    for i, p in enumerate(projects):
        pp.append([p["project"], round(p["drawn"]), round(p["repaid"]),
                   round(p["outstanding"]), round(p["avg_lag"], 1)])
        r = pp.max_row
        for c in range(1, 6):
            pp.cell(r, c).border = _BORDER
            if i % 2:
                pp.cell(r, c).fill = _ZEBRA
        for c in (2, 3, 4):
            pp.cell(r, c).number_format = CUR
        if p["outstanding"] > 0:
            pp.cell(r, 4).font = Font(bold=True, color="9C0006")
    last = pp.max_row
    for col in (2, 4):
        pp.conditional_formatting.add(
            f"{get_column_letter(col)}2:{get_column_letter(col)}{last}",
            DataBarRule(start_type="num", start_value=0, end_type="max",
                        color="F8696B" if col == 4 else "8EAADB", showValue=True))
    pp.auto_filter.ref = f"A1:E{last}"
    pp.sheet_properties.tabColor = "548235"

    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    os.chmod(out, stat.S_IRUSR | stat.S_IWUSR)


# ────────────────────────── main ──────────────────────────

def main() -> int:
    ap = argparse.ArgumentParser(description="Subcontractor LOC float model")
    ap.add_argument("--out", type=Path, default=DEFAULT_OUTPUT)
    ap.add_argument("--months", type=int, default=3,
                    help="months back to the first Friday (default 3)")
    ap.add_argument("--start", type=str, help="override start date YYYY-MM-DD")
    args = ap.parse_args()

    today = _today()
    start = _parse(args.start) if args.start else first_friday_months_back(today, args.months)
    lookback = first_friday_months_back(today, LOOKBACK_MONTHS).isoformat()

    print("\n  SUBCONTRACTOR LOC — float model")
    print("  " + "─" * 55)
    print(f"  window {start} → {today}")

    access, cid = qbo_api.load_credentials()
    print("  sub bill lines …")
    sub_lines = build_sub_bill_lines(access, cid, lookback)
    print(f"    {len(sub_lines)} sub bill(s)")
    print("  invoice → project map …")
    inv_proj = build_invoice_meta(access, cid, lookback)
    print("  draws (sub payments) …")
    draws = collect_draws(access, cid, sub_lines, start)
    print(f"    {len(draws)} draw line-event(s), ${sum(d['amount'] for d in draws):,.0f}")
    print("  repayments (client payments) …")
    repays = collect_repays(access, cid, inv_proj, start)
    print(f"    {len(repays)} repay event(s), ${sum(r['amount'] for r in repays):,.0f}")

    events, summary = run_fifo(draws, repays)
    projects = per_project(events)
    print(f"\n  PEAK LOC needed: ${summary['peak']:,.0f} on {summary['peak_date']}")
    print(f"  avg draw→repay: {summary['avg_lag']:.1f} days · "
          f"still out ${summary['outstanding']:,.0f}")

    write_workbook(args.out, events, summary, projects, start, today)
    print(f"\n  ✓ {args.out}  (chmod 600)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
