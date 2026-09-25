"""
recurring.py — the recurring-obligations register: what we pay every month,
what CHANGED, and what STOPPED.

Why this exists (the user 2026-08-03): break-even is only as good as the list of
fixed obligations behind it, and that list drifts — a lender re-amortises, an MCA
gets refinanced, a subscription lapses, a vendor raises a rate. So rather than
trusting a static number, we put every recurring obligation in front of the owner
each run with three questions answered:

    a. did the payment AMOUNT change?
    b. did the payments STOP?
    c. anything NEW / otherwise changed?

The owner reads it, tells us what really happened, and their note is preserved.

TWO SOURCES, one method — month-by-month history from QBO:
  • Overhead  → ProfitAndLoss, summarize_column_by=Month (per expense account)
  • Debt      → BalanceSheet, summarize_column_by=Month (per liability account);
    the month-over-month DROP in the balance IS the cash payment. That works
    here precisely because interest is not booked as transactions are entered —
    the CPA journals total interest at year-end — so the whole payment lands on
    the liability. (The user 2026-08-03: "it doesn't matter what the split is,
    just the payment" — correct for a CASH break-even.)
    CAVEAT: where interest/fees ARE booked (MCA fees sit in overhead), the
    balance delta is principal only and the booked fee is a separate line — do
    not add both for the same debt or the cost is double-counted.
"""
from __future__ import annotations

import datetime as dt
import re
from typing import Any, Dict, List, Optional

from shared import qbo_api

# Fixed / recurring overhead accounts (the user 2026-08-03, incl. the ones they
# confirmed after review). Anything not listed is treated as variable/one-off
# and stays OUT of the fixed nut.
FIXED_OVERHEAD = {
    "payroll expenses", "rent expense", "insurance expense", "accounting",
    "telephone expense", "utilities", "computer expenses", "security",
    "superintendent", "admin contract labor",
    # added after review — recurring but missing from the first list
    "overhead vehicle expenses", "bank expenses", "office expenses", "taxes",
    "overhead job supplies", "safety", "employee benefits",
    "sales & marketing", "dues and subscriptions", "postage & shipping",
    "printing and reproduction", "quickbooks payments fees", "transaction fee",
    "uniforms",
}
# Explicitly variable / contingent — excluded from fixed overhead.
VARIABLE_OVERHEAD = {
    "business development services",   # MFD director pay, contingent on draws
    "charitable contributions", "client/vendor gifts", "meals & entertainment",
    "legal and professional fees", "travel expense", "miscellaneous",
}
# MCAs REFINANCE each other (the user 2026-08-03): a new advance pays off the
# prior balance and issues a new one, and the CPA records that with a journal
# entry. So an MCA balance moves for reasons that are NOT cash leaving the
# business — the month-over-month delta is meaningless as a payment, a payoff
# looks like a giant "payment", and the successor looks "NEW" while the
# predecessor looks "STOPPED". MCA rows are therefore reported for review but
# never counted in debt service; their fees are already booked in overhead.
REFINANCING_DEBT = re.compile(r"\bMCA\b", re.IGNORECASE)
FEES_BOOKED_IN_OVERHEAD = REFINANCING_DEBT

# A change is only worth the owner's attention if it is BOTH proportionally
# meaningful and materially large — otherwise the register becomes noise.
CHANGE_TOLERANCE = 0.20      # >20% move …
CHANGE_MIN_DOLLARS = 500.0   # … AND at least this many dollars
STOPPED_MONTHS = 2           # no payment for this many closed months = stopped


def _f(v) -> Optional[float]:
    if v in (None, ""):
        return None
    try:
        return float(str(v).replace(",", "").replace("$", ""))
    except ValueError:
        return None


def _walk(node, out: List) -> List:
    """Flatten a QBO report tree into [(row_label, [col values])]."""
    for rw in (node.get("Rows", {}).get("Row", []) if isinstance(node, dict) else []):
        cd = rw.get("ColData")
        if cd and cd[0].get("value"):
            out.append((cd[0]["value"], [c.get("value", "") for c in cd[1:]]))
        _walk(rw, out)
        summ = rw.get("Summary")
        if summ and summ.get("ColData"):
            cd2 = summ["ColData"]
            if cd2[0].get("value"):
                out.append((cd2[0]["value"], [c.get("value", "") for c in cd2[1:]]))
    return out


def _months(report: dict) -> List[str]:
    """Closed months only. QBO labels a partial month 'Aug 1-3, 2026'; including
    it makes every obligation look like it collapsed, so it is dropped."""
    cols = [c.get("ColTitle", "") for c in
            (report.get("Columns", {}).get("Column") or [])]
    out = [c for c in cols[1:] if c and c.lower() != "total"]
    return [c for c in out if not re.search(r"\d+\s*-\s*\d+", c)]


def _classify(series: List[Optional[float]], months: List[str]) -> dict:
    """Given a per-month amount series, answer: last paid, last amount, prior
    amount, did it change, did it stop, is it new."""
    paid = [(m, v) for m, v in zip(months, series) if v]
    if not paid:
        return {"status": "none", "last_month": "", "last_amount": 0.0,
                "prior_amount": 0.0, "change": 0.0, "note": "no activity"}
    last_month, last_amount = paid[-1]
    prior_amount = paid[-2][1] if len(paid) > 1 else 0.0
    change = ((last_amount - prior_amount) / prior_amount) if prior_amount else 0.0
    # closed months since the last payment (ignore the partial current month)
    idx = months.index(last_month)
    gap = len(months) - 1 - idx
    first_month = paid[0][0]

    if gap >= STOPPED_MONTHS:
        status, note = "STOPPED", f"no payment since {last_month}"
    elif len(paid) == 1 and months.index(first_month) >= len(months) - 2:
        status, note = "NEW", f"first seen {first_month}"
    elif (prior_amount and abs(change) > CHANGE_TOLERANCE
          and abs(last_amount - prior_amount) >= CHANGE_MIN_DOLLARS):
        status = "CHANGED"
        note = (f"{prior_amount:,.0f} → {last_amount:,.0f} "
                f"({change*100:+.0f}%) in {last_month}")
    else:
        status, note = "steady", ""
    return {"status": status, "last_month": last_month, "last_amount": last_amount,
            "prior_amount": prior_amount, "change": change, "note": note,
            "months_paid": len(paid), "first_month": first_month}


def _liability_names(qc) -> set:
    """Names of real liability accounts — a falling ASSET balance (equipment,
    receivables, bank) is not a debt payment and must never be counted here."""
    try:
        accts = qc.entity("Account", "Classification = 'Liability'")
    except Exception:
        return set()
    return {str(a.get("Name", "")).strip().lower() for a in accts
            if a.get("Name") and "accounts payable" not in str(a["Name"]).lower()}


def build(qc, start: str = "", as_of: Optional[dt.date] = None) -> dict:
    """Recurring overhead + debt-service register with change/stop detection."""
    as_of = as_of or dt.date.today()
    start = start or f"{as_of.year}-01-01"
    access, cid = qc.credentials()

    pl = qbo_api.report(access, cid, "ProfitAndLoss",
                        {"start_date": start, "end_date": as_of.isoformat(),
                         "summarize_column_by": "Month",
                         "accounting_method": "Accrual"})
    bs = qbo_api.report(access, cid, "BalanceSheet",
                        {"start_date": start, "end_date": as_of.isoformat(),
                         "summarize_column_by": "Month",
                         "accounting_method": "Accrual"})
    months = _months(pl)

    # ── overhead (P&L expense accounts) ──
    overhead: List[dict] = []
    for name, vals in _walk(pl, []):
        key = name.strip().lower().removeprefix("total ").strip()
        if key not in FIXED_OVERHEAD and key not in VARIABLE_OVERHEAD:
            continue
        series = [_f(v) for v in vals[:len(months)]]
        if not any(series):
            continue
        info = _classify(series, months)
        overhead.append({
            "name": name.strip().removeprefix("Total ").strip(),
            "kind": "fixed" if key in FIXED_OVERHEAD else "variable",
            "series": series, **info,
            "ytd": sum(v for v in series if v),
        })
    # de-dupe (the report emits both a row and its summary)
    seen: Dict[str, dict] = {}
    for r in overhead:
        cur = seen.get(r["name"].lower())
        if not cur or r["ytd"] > cur["ytd"]:
            seen[r["name"].lower()] = r
    overhead = sorted(seen.values(), key=lambda r: -r["ytd"])

    # ── debt (liability balances; month-over-month drop = the payment) ──
    liabs = _liability_names(qc)
    debt: List[dict] = []
    for name, vals in _walk(bs, []):
        low = name.strip().lower()
        if low.startswith(("total", "net ")) or low not in liabs:
            continue
        bal = [_f(v) for v in vals[:len(months)]]
        if sum(1 for v in bal if v) < 2:
            continue
        # payment = previous balance − this balance (positive drops only)
        pays: List[Optional[float]] = [None]
        for i in range(1, len(bal)):
            a, b = bal[i - 1], bal[i]
            pays.append(round(a - b, 2) if (a and b and a - b > 0) else None)
        if not any(pays):
            continue
        info = _classify(pays, months)
        debt.append({
            "name": name.strip(), "series": pays, "balances": bal, **info,
            "balance_now": next((v for v in reversed(bal) if v), 0.0),
            "fees_booked": bool(FEES_BOOKED_IN_OVERHEAD.search(name)),
            "refinancing": bool(REFINANCING_DEBT.search(name)),
            "ytd": sum(v for v in pays if v),
        })
    seen2: Dict[str, dict] = {}
    for r in debt:
        cur = seen2.get(r["name"].lower())
        if not cur or r["ytd"] > cur["ytd"]:
            seen2[r["name"].lower()] = r
    debt = sorted(seen2.values(), key=lambda r: -r["last_amount"])

    fixed_month = sum(r["last_amount"] for r in overhead if r["kind"] == "fixed")
    # exclude refinancing debt: its balance delta is a JE, not cash out
    debt_month = sum(r["last_amount"] for r in debt
                     if r["status"] != "STOPPED" and not r["refinancing"])
    return {
        "months": months, "as_of": as_of,
        "overhead": overhead, "debt": debt,
        "fixed_overhead_month": fixed_month,
        "debt_service_month": debt_month,
        "total_monthly_obligation": fixed_month + debt_month,
        "alerts": [r for r in overhead + debt
                   if r["status"] in ("CHANGED", "STOPPED", "NEW")
                   and not r.get("refinancing")],
        "refinancing": [r for r in debt if r.get("refinancing")],
    }
