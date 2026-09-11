"""
qbo_pl.py - company-level P&L totals from the QBO ProfitAndLoss report.

The ONE place that walks a QBO P&L report tree and extracts the five totals the
break-even / margin math runs on (income, cogs, gross_profit, overhead,
net_operating) - shaped exactly like shared/breakeven.py's block dicts so the
two snap together. Used by ledger/load_health.py; health-dashboard/qbo_health.py
keeps its historical local copy until that tool retires with the fold-in.

Robust to construction COA naming (Job Costs vs COGS, Operating Expenses vs
Expenses) via the same keyword hints qbo_health hardened over time.
"""
from __future__ import annotations

from typing import Dict, List, Optional, Tuple

from shared import qbo_api

_INCOME_HINTS = ("total income", "total revenue", "total sales", "total revenues")
_COGS_HINTS = (
    "total cost of goods sold", "total cogs",
    "total job costs", "total job cost", "total job-related costs",
    "total direct costs", "total direct cost",
    "total cost of sales", "total cost of revenue",
    "total construction costs",
)
_EXPENSE_HINTS = (
    "total expenses", "total expense",
    "total operating expenses", "total operating expense",
    "total overhead", "total g&a",
    "total general & administrative",
)


def _walk(node: dict, out: List[Tuple[str, Optional[float]]]) -> List[Tuple[str, Optional[float]]]:
    """Flatten a QBO report tree to [(label, first_value)] - rows AND summaries."""
    for rw in (node.get("Rows", {}).get("Row", []) if isinstance(node, dict) else []):
        for cd in (rw.get("ColData"), (rw.get("Summary") or {}).get("ColData")):
            if cd and cd[0].get("value"):
                try:
                    v = float(str(cd[1].get("value", "")).replace(",", "")) if len(cd) > 1 else None
                except (ValueError, TypeError):
                    v = None
                out.append((cd[0]["value"], v))
        _walk(rw, out)
    return out


def _matches(key: str, hints: tuple) -> bool:
    return any(key.startswith(h) for h in hints)


def totals_from_report(data: dict) -> Dict[str, float]:
    """{'income','cogs','gross_profit','overhead','net_operating','net_income'}
    from one ProfitAndLoss report payload (breakeven-block key names).

    An EXACT hint match locks the field; a prefix match only fills an empty one.
    Without the lock, an account group like 'Total Overhead Vehicle Expenses'
    (prefix-matches 'total overhead') would overwrite the real section total
    'Total Expenses' - the walk visits it later, and last-wins burned us."""
    out: Dict[str, float] = {}
    exact: set = set()

    def put(field: str, key: str, v: float, hints: tuple) -> None:
        if key in hints:
            out[field] = v
            exact.add(field)
        elif field not in out and field not in exact:
            out[field] = v

    for label, v in _walk(data, []):
        if v is None:
            continue
        key = label.strip().lower()
        if _matches(key, _INCOME_HINTS):
            put("income", key, v, _INCOME_HINTS)
        elif _matches(key, _COGS_HINTS):
            put("cogs", key, v, _COGS_HINTS)
        elif key == "gross profit":
            out["gross_profit"] = v
        elif _matches(key, _EXPENSE_HINTS):
            put("overhead", key, v, _EXPENSE_HINTS)
        elif key.startswith(("net operating income", "operating income")):
            out["net_operating"] = v
        elif key.startswith("net income") or key == "net loss":
            out["net_income"] = v
    if "cogs" not in out and "income" in out and "gross_profit" in out:
        out["cogs"] = out["income"] - out["gross_profit"]
    if "gross_profit" not in out and "income" in out and "cogs" in out:
        out["gross_profit"] = out["income"] - out["cogs"]
    # Identity check: Net Operating Income = Gross Profit - Total Expenses, always.
    # If the matched overhead breaks it, derive overhead from the identity instead.
    if "gross_profit" in out and "net_operating" in out:
        derived = out["gross_profit"] - out["net_operating"]
        if abs(out.get("overhead", 0.0) - derived) > 1.0:
            out["overhead"] = derived
    return out


def pl_totals(access: str, company_id: str, start: str, end: str) -> Dict[str, float]:
    """One accrual P&L pull for [start, end] reduced to the block totals."""
    data = qbo_api.report(access, company_id, "ProfitAndLoss",
                          {"start_date": start, "end_date": end,
                           "accounting_method": "Accrual"})
    return totals_from_report(data)
