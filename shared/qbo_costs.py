"""
qbo_costs.py — the shared cost-extraction engine.

The ONE place that turns QBO expense transactions into cost lines keyed by our
cost code. Lifted out of project-pnl the moment a second tool (the ledger's
load_costs.py) needed it — so `cost_leaf()` is now the single authoritative
resolver for BOTH the P&L export and the ledger, and they can never drift.

WHAT LIVES HERE
    cost_leaf(det, account_names, fallback)  the canonical cost-code / account resolver
    is_cost_code(name)                       True for 'SL1', 'PV51', '9', ...
    cost_code_meta(code)                     ('SL1' → prefix SL, number 1, name 'Concrete')
    sub_evidence(memo)                       the memo token that flags a SUB bill (kept beside is_sub)
    build_account_map(access, company_id)    QBO account id → name
    pull_expense_txns(access, company_id)    fetch Bills + Purchases (Expenses)
    cost_lines_from_txns(...)                a generator over txns → structured cost lines
    iter_cost_lines(...)                     the whole pipeline: pull + resolve

QBO API pin (learned the hard way — see CLAUDE.md): OUR cost codes live in the
ITEM name (`ItemRef.name`), NOT the account. An item-based expense line carries
`ItemBasedExpenseLineDetail` with an `ItemRef` and NO line-level `AccountRef`, so
it resolves to the item = the cost code (SL1, PV6, CS1…). An account-based line
resolves to its account name ('Job Materials: Concrete'). NEVER resolve the item
to its posting account — that collapses every SL#/PV# into one account.
"""
from __future__ import annotations

import re
from typing import Dict, Iterable, Iterator, List, Optional, Tuple

from shared.qbo_api import query_all, extract_proj

_ILLEGAL_XML_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")

# cost-code number → the cost NAME it lands in (the user 2026-06-09).
_COST_CODE_NAMES = {
    "1": "Concrete", "2": "Rebar & Reinforcement", "3": "Formwork & Lumber",
    "4": "Aggregates", "5": "Equipment & Rentals", "51": "Pump", "52": "Saw Cutting",
    "6": "Labor", "7": "Specialty/Misc.", "8": "Fuel", "9": "Supplies",
}
# job-type prefix → full name
_JOB_TYPE_NAMES = {
    "SL": "Slab", "PV": "Paving", "FW": "Flatwork for Residential", "PR": "Piers",
    "WL": "Walls", "CS": "Commercial Sidewalks", "MS": "Miscellaneous",
}
# optional 2-letter JOB-TYPE prefix + cost-code number (51/52 before single digits)
_COST_CODE_RE = re.compile(r"^(?:SL|PV|FW|PR|WL|CS|MS)?(51|52|[1-9])$", re.IGNORECASE)
_COST_CODE_SPLIT_RE = re.compile(r"^(SL|PV|FW|PR|WL|CS|MS)?(51|52|[1-9])$", re.IGNORECASE)


def _xml_clean(s):
    """Strip XML-illegal control chars from a string (passes non-str through)."""
    if not isinstance(s, str):
        return s
    return _ILLEGAL_XML_RE.sub("", s)


def cost_leaf(det: dict, account_names: Dict[str, str],
              fallback: str = "(unclassified)") -> str:
    """THE cost-code / account a QBO expense LINE lands in — the single
    authoritative resolver (the user 2026-07-17). Used by the P&L export's
    cost buckets, Budget vs Actual, and the ledger, so every cost figure keyed
    by cost code ties. Resolution order:
        AccountRef account name → AccountRef.name last segment →
        ItemRef.name (the cost code) → fallback."""
    aref = det.get("AccountRef") or {}
    aid = aref.get("value")
    return _xml_clean(
        account_names.get(aid)
        or (aref.get("name") or "").split(":")[-1].strip()
        or (det.get("ItemRef") or {}).get("name")
        or fallback)


def is_cost_code(name) -> bool:
    """True if the string is a cost code (e.g. 'CS1', 'SL6', 'PV51', '9')."""
    return isinstance(name, str) and bool(_COST_CODE_RE.match(name.strip()))


def cost_code_meta(code: str) -> dict:
    """'SL1' → {code, prefix 'SL', number '1', description 'Concrete'}.
    A non-code string comes back with prefix/number/description = None."""
    m = _COST_CODE_SPLIT_RE.match((code or "").strip())
    if not m:
        return {"code": code, "prefix": None, "number": None, "description": None}
    prefix = (m.group(1) or "").upper() or None
    num = m.group(2)
    return {"code": code.strip(), "prefix": prefix, "number": num,
            "description": _COST_CODE_NAMES.get(num)}


def job_type_name(prefix: str) -> Optional[str]:
    """'SL' → 'Slab'. The job-type PREFIX name — the sub under a cost-type parent
    (the JobTread model: material → one cost-type parent, job-type as the sub)."""
    return _JOB_TYPE_NAMES.get((prefix or "").upper())


_SUB_TOKEN_RE = re.compile(r"\S*sub\S*", re.IGNORECASE)


def sub_evidence(memo) -> Optional[str]:
    """The memo token that makes a bill a SUB bill - 'sub' anywhere in the bill memo
    (CLAUDE.md), e.g. 'sub', 'Subcontractor', 'SUB-LABOR' - or None. Stored beside
    `is_sub` (2026-09-01) so a reader can see WHY a line was called a sub."""
    m = _SUB_TOKEN_RE.search(memo or "")
    return m.group(0) if m else None


# ── QBO pull ────────────────────────────────────────────────────────────────

def build_account_map(access: str, company_id: str) -> Dict[str, str]:
    """QBO Account Id → name (leaf name preferred, e.g. 'Concrete')."""
    out: Dict[str, str] = {}
    for a in query_all(access, company_id, "Account"):
        name = a.get("Name") or a.get("FullyQualifiedName") or f"Account {a['Id']}"
        out[a["Id"]] = name
    return out


def pull_expense_txns(access: str, company_id: str,
                      since: Optional[str] = None) -> Tuple[List[dict], List[dict]]:
    """Fetch (bills, purchases). `since` is an inclusive ISO date on TxnDate."""
    where = f"TxnDate >= '{since}'" if since else ""
    return (query_all(access, company_id, "Bill", where),
            query_all(access, company_id, "Purchase", where))


def cost_lines_from_txns(
    txns: Iterable[dict],
    tx_type: str,
    vendor_field: str,
    account_names: Dict[str, str],
    customer_to_project: Dict[str, str],
) -> Iterator[dict]:
    """Yield one structured cost-line dict per attributable expense line.

    Network-free (operates on already-pulled txns) so it is unit-testable. A
    line is attributed to a project by its own `CustomerRef` (falling back to a
    project # parsed from the memo). `cost_code` is set only when the leaf is a
    real code (SL/PV/…); account-based lines carry `account` with `cost_code`
    NULL. `is_sub` follows the 'sub' flag in the bill memo (CLAUDE.md).

    Since 2026-09-01 every line also carries the audit trail a reader asks for
    (ADDITIVE keys - every earlier key is unchanged, project-pnl imports from here):
    `doc_number` (the document's DocNumber), `memo` (the bill-level PrivateNote,
    kept SEPARATE - `description` is now ONLY the line's own Description, never
    the memo folded in), `line_no` (LineNum, else position), `bill_total` (the
    document TotalAmt - display only, never summed: cost is LINE amounts),
    `vendor_id`, `class_name`, and `is_sub_evidence` (the memo token behind is_sub).
    Which lines are yielded, and with what amount, is exactly as before."""
    for t in txns or []:
        memo = _xml_clean((t.get("PrivateNote") or "").strip())
        is_sub = "sub" in (memo or "").lower()
        evidence = (sub_evidence(memo) or "sub") if is_sub else None
        vref = t.get(vendor_field) or {}
        vendor = _xml_clean((vref.get("name") or "").strip()) or None
        vendor_id = str(vref.get("value")) if vref.get("value") not in (None, "") else None
        txn_id = t.get("Id", "")
        txn_date = t.get("TxnDate")
        doc_number = _xml_clean(str(t.get("DocNumber") or "").strip()) or None
        try:
            bill_total = float(t["TotalAmt"]) if t.get("TotalAmt") not in (None, "") else None
        except (TypeError, ValueError):
            bill_total = None
        txn_class = (t.get("ClassRef") or {}).get("name") or None
        for idx, ln in enumerate(t.get("Line") or []):
            det = (ln.get("AccountBasedExpenseLineDetail")
                   or ln.get("ItemBasedExpenseLineDetail"))
            if not det:
                continue
            amt = float(ln.get("Amount", 0) or 0)
            if amt == 0:
                continue
            cust_id = (det.get("CustomerRef") or {}).get("value")
            proj = customer_to_project.get(cust_id) if cust_id else None
            if not proj:
                proj = extract_proj(memo)
            leaf = cost_leaf(det, account_names, fallback="(unclassified)")
            if is_cost_code(leaf):
                cost_code, account = leaf, account_names.get((det.get("AccountRef") or {}).get("value"))
            else:
                cost_code, account = None, leaf
            try:
                line_no = int(ln["LineNum"]) if ln.get("LineNum") is not None else idx + 1
            except (TypeError, ValueError):
                line_no = idx + 1
            yield {
                "qbo_txn_id": txn_id,
                "qbo_line_id": str(ln.get("Id") or idx),
                "txn_type": tx_type,
                "project_no": proj,
                "customer_id": cust_id,
                "cost_code": cost_code,
                "account": account,
                "amount": amt,
                "txn_date": txn_date,
                "is_sub": 1 if is_sub else 0,
                "vendor": vendor,
                "description": _xml_clean((ln.get("Description") or "").strip()) or None,
                # additive (2026-09-01): the trail behind the line
                "doc_number": doc_number,
                "memo": memo or None,
                "line_no": line_no,
                "bill_total": bill_total,
                "vendor_id": vendor_id,
                "class_name": _xml_clean((det.get("ClassRef") or {}).get("name") or txn_class) or None,
                "is_sub_evidence": evidence,
            }


def iter_cost_lines(
    access: str,
    company_id: str,
    account_names: Dict[str, str],
    customer_to_project: Dict[str, str],
    since: Optional[str] = None,
) -> Iterator[dict]:
    """Pull Bills + Purchases and yield every attributable cost line."""
    bills, purchases = pull_expense_txns(access, company_id, since)
    yield from cost_lines_from_txns(bills, "Bill", "VendorRef", account_names, customer_to_project)
    yield from cost_lines_from_txns(purchases, "Expense", "EntityRef", account_names, customer_to_project)
