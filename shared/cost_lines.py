"""
cost_lines.py — cost-line classification + bill-line combining (shared).

WHY THIS EXISTS
project-pnl's Transactions sheet groups job costs CONCRETE → LABOR → MATERIALS
(the biggest money movers first — the user 2026-07-15) and combines a bill's
line items when they hit the same account — but NEVER combines labor lines
(subs/crews are paid per line and every line must stay visible). The concept
comes from bill-tracker's collapse_rows(); the repo rule is that tools never
import tools, so the reusable piece lives here in shared/.

CLASSIFICATION
Category is decided from the line's account name (QBO chart of accounts),
with a cost-code fallback (job-type prefix + number: …1 = Concrete,
…6 = Labor — see the CP/RP takeoff cost codes):
    "Job Materials: Concrete"        → Concrete
    "Subcontractors Expense: Labor"  → Labor        (any account naming labor)
    everything else                  → Materials    (rebar, lumber, aggregates,
                                                     pump, equipment, misc …)

COMBINING (grain = one bill × one account)
combine_bill_lines() merges records that share (txn_id, account) into one row —
amounts summed, description tagged "(N lines)" — ONLY when the account is not
Labor. Lines of the same bill that hit different accounts stay separate, so a
combined row always carries ONE real account and any per-account SUMIF over the
sheet still ties to the penny.
"""
from __future__ import annotations

import re
from typing import Dict, List

CONCRETE = "Concrete"
LABOR = "Labor"
MATERIALS = "Materials"

# category display order — concrete first, then labor, then materials
CATEGORY_ORDER = {CONCRETE: 0, LABOR: 1, MATERIALS: 2}

# Cost-code fallback: optional job-type prefix + cost number ("SL1", "PV6", "CS1").
_COST_CODE_RE = re.compile(r"^(?:SL|PV|FW|PR|WL|CS|MS)?\s*-?\s*(51|52|[1-9])\b",
                           re.IGNORECASE)
_CONCRETE_NUM = {"1"}
_LABOR_NUM = {"6"}


def line_category(account: str) -> str:
    """Concrete / Labor / Materials for a cost line's account name."""
    name = (account or "").strip()
    low = name.lower()
    if "labor" in low:
        return LABOR
    if "concrete" in low:
        return CONCRETE
    m = _COST_CODE_RE.match(name)
    if m:
        if m.group(1) in _CONCRETE_NUM:
            return CONCRETE
        if m.group(1) in _LABOR_NUM:
            return LABOR
    return MATERIALS


def category_sort_key(account: str) -> int:
    return CATEGORY_ORDER.get(line_category(account), len(CATEGORY_ORDER))


def combine_bill_lines(recs: List[dict]) -> List[dict]:
    """Merge same-(bill × account) lines for NON-labor accounts.

    Each rec is a dict with at least: txn_id (or ref), account, amount, desc.
    Labor lines pass through untouched — one row per line, always. Merged rows
    keep the first rec's fields, sum `amount`, tag desc with "(N lines)", and
    carry `line_count`. Order of first appearance is preserved.
    """
    out: List[dict] = []
    groups: Dict[tuple, dict] = {}
    for rec in recs:
        acct = rec.get("account") or ""
        key = (rec.get("txn_id") or rec.get("ref") or id(rec), acct)
        if line_category(acct) == LABOR:
            out.append(rec)                     # labor: never combined
            continue
        g = groups.get(key)
        if g is None:
            g = dict(rec)
            g["line_count"] = 1
            groups[key] = g
            out.append(g)
        else:
            g["amount"] = round(
                float(g.get("amount", 0) or 0) + float(rec.get("amount", 0) or 0), 2)
            g["line_count"] += 1
            if not g.get("desc") and rec.get("desc"):
                g["desc"] = rec["desc"]
    for g in groups.values():
        n = g.get("line_count", 1)
        if n > 1:
            g["desc"] = f"{(g.get('desc') or '').strip()}  ({n} lines)".strip()
    return out
