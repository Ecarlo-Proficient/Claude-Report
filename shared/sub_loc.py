"""
sub_loc.py - the subcontractor line-of-credit (LOC) float ENGINE.

The question (the owner 2026-07-17): if every subcontractor payment were funded by a
line of credit and repaid when the client pays us for that work, how big a LOC do we
truly need, and how long is our money out before it comes back?

This module is the pure/QBO-pulling model, shared by BOTH the standalone Excel report
(one-offs/sub_loc_report.py) and the ledger loader (ledger/load_sub_loc.py) - so the
numbers are identical everywhere. It holds NO Excel / no output paths.

MODEL (all choices are the owner's, 2026-07-17):
  - Sub bill = a QBO Bill whose memo/PrivateNote contains "sub" (same rule as bill-tracker's
    is_sub_bill). Worked at the LINE level - each line carries its project on CustomerRef;
    one bill can span several projects.
  - DRAW (money out) = when the sub is actually PAID - the QBO BillPayment date - allocated
    across the bill's lines pro-rata by line amount.
  - REPAY (money in) = when the client pays us - the QBO customer Payment date, mapped to a
    project through the invoice it was applied to.
  - Matching = per-project FIFO, matched BY DRAW PERIOD (MFD/CP: a draw only offsets sub
    costs of the same project+month; RP invoices are lump per scope -> project only). Events
    are processed CHRONOLOGICALLY so genuine prefunding is not miscounted as repayment.

Running LOC balance = cumulative draws - cumulative applied repayments over time; its PEAK
is the LOC you truly need, and `outstanding` is what is still fronted-but-uncollected NOW.
READ-ONLY against QBO.
"""
from __future__ import annotations

import datetime as dt
import re
from collections import defaultdict
from typing import Dict, List, Optional, Tuple

try:                                   # sibling shared module (package import)
    from . import qbo_api
except ImportError:                    # run with the repo root on sys.path
    import qbo_api                     # type: ignore

SUB_RE = re.compile(r"\bsub\b", re.IGNORECASE)     # same as bill-tracker's is_sub_bill
# pull bills/invoices this far back so txns paid inside the window are covered
LOOKBACK_MONTHS = 9


# ────────────────────────── dates ──────────────────────────

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
    """A draw/work period like '7/01/2026 - 7/31/2026' -> the 'YYYY-MM' month containing its
    MIDPOINT (draws straddle month boundaries, e.g. 1/26-2/25 = the February draw). None when
    no period range is present."""
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


def division_of(project: str) -> str:
    p = (project or "").upper()
    if p.startswith("MFD"):
        return "MFD"
    if p.startswith("CP"):
        return "CP"
    if p.startswith("RP"):
        return "RP"
    return "Other"


# ────────────────────────── QBO pulls (read-only) ──────────────────────────

def _line_project(line: dict) -> Optional[str]:
    for k in ("AccountBasedExpenseLineDetail", "ItemBasedExpenseLineDetail"):
        det = line.get(k) or {}
        name = (det.get("CustomerRef") or {}).get("name")
        if name:
            return qbo_api.extract_proj(name)
    return None


def build_sub_bill_lines(access, cid, since: str) -> Dict[str, dict]:
    """{billId: {vendor, lines:[(project, amount)], total, work_month}} for sub bills.
    work_month = the month of the memo's 'Period ...' (the draw period the sub cost belongs
    to), else the bill-date month."""
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
    """{invoiceId: {project, draw_month, doc}}. draw_month = the month of the invoice memo's
    'Draw #N (Period ...)', else the invoice-date month (RP lump invoices have no period ->
    their draw_month is unused, RP matches by project only)."""
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


def collect_draws(access, cid, sub_lines: Dict[str, dict], start: dt.date) -> List[dict]:
    """Money OUT: each BillPayment that pays a sub bill -> per-project draw events, allocated
    pro-rata across the bill's lines."""
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


def collect_repays(access, cid, inv_meta: Dict[str, dict], start: dt.date) -> List[dict]:
    """Money IN: each customer Payment -> per-project repayment events, mapped through the
    invoice it was applied to (which carries the draw month)."""
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
      - DRAW: first consume any client cash already received on that project (prefunded - the
        GC draw came in before we paid the sub, common on MFD); only the UN-prefunded
        remainder hits the LOC and joins the FIFO queue.
      - REPAY: pay down the oldest outstanding LOC draws first (that's the draw->repay float);
        any leftover becomes prefunding for future draws (capped there - the margin above sub
        cost never funds more than a later draw needs). Running LOC balance never goes negative.
    Peak running balance = the LOC we truly need. Lag average is over LOC-funded draws only
    (prefunded ones had zero money-out time)."""
    eps = 1e-6
    merged = ([{"kind": "D", **d} for d in draws]
              + [{"kind": "R", **r} for r in repays])
    # same day: DRAW before REPAY -> conservative (higher) peak
    merged.sort(key=lambda e: (e["date"], 0 if e["kind"] == "D" else 1))

    # Match bucket: MFD/CP reimburse BY DRAW PERIOD - a draw only offsets sub costs of the
    # SAME (project, month). RP invoices are lump per scope (no draw period) -> match by
    # project only. This stops a May draw from phantom-'prefunding' June sub costs.
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
    div: Dict[str, dict] = defaultdict(
        lambda: {"drawn": 0.0, "prefunded": 0.0, "repaid": 0.0, "wl": 0.0,
                 "wa": 0.0, "n_draws": 0, "outstanding": 0.0,
                 "peak": 0.0, "peak_date": None})

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
                  "note": note, "balance": bal, "reimb": [],
                  "loc_delta": loc_part if loc_part > eps else 0.0}
            if loc_part > eps:
                # node carries a ref to its event so repayments can stamp the reimbursing
                # invoice + client-paid date back onto the draw
                queues[p].append({"date": e["date"], "remaining": loc_part, "ev": ev})
                bal += loc_part
            ev["balance"] = bal
            events.append(ev)
            dv = div[division_of(e["project"])]
            dv["drawn"] += amt
            dv["prefunded"] += use_pre
            dv["n_draws"] += 1
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
                dv = div[division_of(e["project"])]
                dv["repaid"] += take
                dv["wl"] += take * days
                dv["wa"] += take
                if node["remaining"] <= eps:
                    q.pop(0)
            bal -= applied_here
            surplus = R if R > eps else 0.0
            prepaid[p] += surplus
            note = f"surplus ${surplus:,.0f}" if surplus > eps else ""
            events.append({"date": e["date"], "type": "REPAY", "project": e["project"],
                           "party": e["party"], "out": 0.0, "inn": applied_here,
                           "lag": first_lag, "note": note, "balance": bal,
                           "invoice": e.get("invoice", ""), "reimb": [],
                           "loc_delta": -applied_here})
        if bal > peak:
            peak, peak_date = bal, e["date"]

    total_drawn = sum(d["amount"] for d in draws)
    outstanding = sum(n["remaining"] for q in queues.values() for n in q)
    wl = sum(a * days for a, days in lags)
    wa = sum(a for a, _ in lags)

    # per-division outstanding (from the remaining FIFO queues)
    for (proj, _per), q in queues.items():
        div[division_of(proj)]["outstanding"] += sum(n["remaining"] for n in q)
    # per-division peak = high-water of that division's own running balance
    dbal: Dict[str, float] = defaultdict(float)
    for e in events:                          # events are in date order
        dv = div[division_of(e["project"])]
        dbal_key = division_of(e["project"])
        dbal[dbal_key] += e["loc_delta"]
        if dbal[dbal_key] > dv["peak"]:
            dv["peak"] = dbal[dbal_key]
            dv["peak_date"] = e["date"]
    divisions = {}
    for name, a in div.items():
        divisions[name] = {
            "peak": a["peak"], "peak_date": a["peak_date"],
            "drawn": a["drawn"], "prefunded": a["prefunded"],
            "repaid": a["repaid"], "outstanding": a["outstanding"],
            "n_draws": a["n_draws"],
            "avg_lag": (a["wl"] / a["wa"]) if a["wa"] else 0.0,
        }

    summary = {
        "n_draws": len(draws), "n_repay_chunks": len(lags),
        "total_drawn": total_drawn, "total_repaid": applied_total,
        "prefunded": prefunded_total, "outstanding": outstanding,
        "peak": peak, "peak_date": peak_date,
        "avg_lag": (wl / wa) if wa else 0.0,
        "avg_draw": (total_drawn / len(draws)) if draws else 0.0,
        "avg_repay": (applied_total / len(lags)) if lags else 0.0,
        "divisions": divisions,
    }
    return events, summary


def per_project(events: List[dict]) -> List[dict]:
    agg: Dict[str, dict] = defaultdict(lambda: {"out": 0.0, "inn": 0.0, "wl": 0.0, "wa": 0.0})
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


# ────────────────────────── orchestration ──────────────────────────

def compute(access, cid, start: dt.date, today: dt.date,
            lookback_months: int = LOOKBACK_MONTHS) -> Tuple[List[dict], dict, List[dict]]:
    """Pull QBO (read-only) and run the model. Returns (events, summary, per_project_rows).
    `start` = window start (balance begins at 0); `lookback_months` = how far back to pull the
    bill/invoice metadata so payments landing inside the window resolve."""
    lookback = first_friday_months_back(today, lookback_months).isoformat()
    sub_lines = build_sub_bill_lines(access, cid, lookback)
    inv_meta = build_invoice_meta(access, cid, lookback)
    draws = collect_draws(access, cid, sub_lines, start)
    repays = collect_repays(access, cid, inv_meta, start)
    events, summary = run_fifo(draws, repays)
    return events, summary, per_project(events)
