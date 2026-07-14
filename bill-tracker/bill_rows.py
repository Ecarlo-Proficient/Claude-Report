"""
bill_rows.py — pure QBO → row-dict helpers, shared by all bill-tracker sinks.

Extracted from notion_bill_sync.py so excel_bill_sync.py (and any future
destination) can build the same row set without a Notion dependency. No I/O
to any destination here — just QBO fetch + line-level row assembly.

Row dict shape (keys consumed by downstream sinks):
    key, bill_id, line_id, bill_date, vendor, bill_doc, po_num,
    bill_total, bill_balance, division, project_num, bill_type, account,
    line_amount, line_desc, inv_doc, inv_id, inv_date, inv_total,
    inv_balance, payment_date, auto_status, approved

Collapse helpers (2026-06-04):
    collapse_rows(line_rows, grain) — roll up line-level rows into bill-grain
        or (bill, project)-grain rows for display. See function docstring.
    multi_project_bill_ids(line_rows) — set of bill_ids whose lines span 2+
        distinct project_nums; used to route bills to the Inventory sheet.
"""
from __future__ import annotations

import datetime as dt
import re
from collections import defaultdict
from typing import Dict, List, Optional, Set, Tuple

from qbo_bill_tracker import (
    query_all,
    get_line_customer_ref,
    get_project_num,
    get_division,
    find_matching_invoice_ex,
    compute_status,
    compute_pay_status,
    compute_invoice_status,
    parse_date,
    STATUS_PAID,
    STATUS_NO_PROJECT,
    STATUS_OK_TO_PAY,
    STATUS_AWAITING_PAYMENT,
    STATUS_AWAITING_INVOICE,
    STATUS_PARTIAL_PAID,
    STATUS_UNPAID,
)


# ─────────────────────── approval flag ───────────────────────

def display_status(auto_status: str, approved: bool = False) -> str:
    """Return the pipeline-status label as-is for the Status column.

    The user 2026-06-03: Approval moved to its own column, so Status is just the
    pipeline state ("Awaiting Payment" / "Awaiting Invoice" / "Invoice paid" /
    "Bill paid" / "No project #"). The `approved` arg is kept (unused) so old
    callers don't break — pass it for forward-compat in case we re-merge
    later.
    """
    return auto_status


def approved_text(approved: bool) -> str:
    """One-word string for the Approved column. Lowercase to match the
    plain-English style used elsewhere on the sheet."""
    return "approved" if approved else "not approved"


def is_approved(bill: dict) -> bool:
    """Bill is approved unless PrivateNote (the Bill Memo) starts with
    'NOT APPROVED' (case-insensitive, leading whitespace OK).

    Per Proficient workflow: AP review tags un-approved bills by prefixing
    the memo. Untagged bills are assumed approved.
    """
    memo = (bill.get("PrivateNote") or "").lstrip()
    return not memo.upper().startswith("NOT APPROVED")


# ─────────────────────── invoice → payment date map ───────────────────────

def build_payment_map(qbo_access: str, qbo_cid: str) -> Dict[str, dt.date]:
    """Build {Invoice.Id → Payment.TxnDate} by walking every Payment object's
    LinkedTxn list. If an invoice is paid by multiple payments (split), we keep
    the LATEST payment date — the date the invoice actually cleared.

    Used to power the Paid This Week dashboard with the date the GC's money
    actually hit our bank, not the date we billed them.
    """
    out: Dict[str, dt.date] = {}
    for pay in query_all(qbo_access, qbo_cid, "Payment"):
        pay_date = parse_date(pay.get("TxnDate"))
        if not pay_date:
            continue
        for lt in (pay.get("Line") or []):
            for linked in (lt.get("LinkedTxn") or []):
                if linked.get("TxnType") != "Invoice":
                    continue
                inv_id = linked.get("TxnId", "")
                if not inv_id:
                    continue
                # latest payment wins (handles split-payment invoices)
                if inv_id not in out or pay_date > out[inv_id]:
                    out[inv_id] = pay_date
    return out


def build_account_maps(qbo_access: str, qbo_cid: str) -> Tuple[Dict[str, str], Dict[str, str]]:
    """Build (account_map, item_map) for resolving line account names + Bill Type.

    item_map collapses Item → expense-account name so item-based lines roll up
    to the same Account bucket as account-based lines.
    """
    account_map: Dict[str, str] = {}
    for a in query_all(qbo_access, qbo_cid, "Account"):
        account_map[a["Id"]] = a.get("Name") or f"Account {a['Id']}"
    item_map: Dict[str, str] = {}
    for it in query_all(qbo_access, qbo_cid, "Item"):
        ea_ref = it.get("ExpenseAccountRef") or {}
        ea_id = ea_ref.get("value")
        item_map[it["Id"]] = account_map.get(ea_id, ea_ref.get("name") or it.get("Name", ""))
    return account_map, item_map


def build_po_map(qbo_access: str, qbo_cid: str) -> Dict[str, str]:
    """Build {PurchaseOrder.Id → PurchaseOrder.DocNumber} for Bill.LinkedTxn lookup."""
    po_map: Dict[str, str] = {}
    for po in query_all(qbo_access, qbo_cid, "PurchaseOrder"):
        po_map[po["Id"]] = (po.get("DocNumber") or "").strip()
    return po_map


def get_po_number(bill: dict, po_map: Dict[str, str]) -> str:
    """Extract PO #(s) for a Bill via LinkedTxn (PO clerk-linked in QBO).

    Multiple POs come back comma-separated. Falls back to CustomField named
    'PO' for QBO files where bills aren't formally linked.
    """
    nums: List[str] = []
    for lt in (bill.get("LinkedTxn") or []):
        if lt.get("TxnType") == "PurchaseOrder":
            doc = po_map.get(lt.get("TxnId", ""), "")
            if doc:
                nums.append(doc)
    if nums:
        return ", ".join(nums)
    for cf in (bill.get("CustomField") or []):
        if "po" in (cf.get("Name") or "").lower():
            val = (cf.get("StringValue") or "").strip()
            if val:
                return val
    return ""


def line_account_and_type(
    line: dict,
    account_map: Dict[str, str],
    item_map: Dict[str, str],
) -> Tuple[str, str]:
    """Return (account_name, bill_type) for a bill line.

    bill_type = 'COGS' for ItemBasedExpenseLineDetail (items roll to inventory/COGS),
                'Other' for AccountBasedExpenseLineDetail (overhead/fixed).
    """
    dt_type = line.get("DetailType", "")
    if dt_type == "ItemBasedExpenseLineDetail":
        item_ref = (line.get("ItemBasedExpenseLineDetail") or {}).get("ItemRef") or {}
        item_id = item_ref.get("value", "")
        return (item_map.get(item_id) or item_ref.get("name", ""), "COGS")
    if dt_type == "AccountBasedExpenseLineDetail":
        acct_ref = (line.get("AccountBasedExpenseLineDetail") or {}).get("AccountRef") or {}
        acct_id = acct_ref.get("value", "")
        return (account_map.get(acct_id) or acct_ref.get("name", ""), "Other")
    return ("", "")


def build_rows(
    bills: List[dict],
    invoices_by_customer: Dict[str, List[dict]],
    vendor_map: Dict[str, str],
    account_map: Dict[str, str],
    item_map: Dict[str, str],
    po_map: Dict[str, str],
    payment_map: Optional[Dict[str, dt.date]] = None,
    gl_contracts: Optional[Dict[str, float]] = None,
) -> List[dict]:
    """One row per bill line. Includes BOTH Item lines (Bill Type=COGS) AND
    Account lines (Bill Type=Other). Lines without project # → Status=NO PROJECT #
    (kept for QC review, not skipped). Bills with Balance==0 → Status=PAID.
    """
    payment_map = payment_map or {}
    rows: List[dict] = []
    for bill in bills:
        bill_id = bill.get("Id", "")
        bill_date = parse_date(bill.get("TxnDate")) or dt.date.today()
        bill_total = float(bill.get("TotalAmt") or 0)
        bill_balance = float(bill.get("Balance") or 0)
        bill_doc = bill.get("DocNumber", "") or ""
        po_num = get_po_number(bill, po_map)
        v_ref = bill.get("VendorRef") or {}
        vendor_name = vendor_map.get(v_ref.get("value", ""), v_ref.get("name", "?"))
        approved = is_approved(bill)

        for line in (bill.get("Line") or []):
            dt_type = line.get("DetailType", "")
            if dt_type not in ("ItemBasedExpenseLineDetail", "AccountBasedExpenseLineDetail"):
                continue
            line_id = line.get("Id", "")
            line_amt = float(line.get("Amount") or 0)
            line_desc = line.get("Description", "") or ""

            cust = get_line_customer_ref(line)
            cust_id = cust.get("value", "")
            cust_name = cust.get("name", "")
            project_num = get_project_num(cust_name)
            division = get_division(project_num)

            account_name, bill_type = line_account_and_type(line, account_map, item_map)

            # ClassRef can live on the line (preferred — bills can split across
            # divisions) or on the bill itself (older QBO setups). Try line first.
            line_detail = line.get(dt_type) or {}
            class_name = ""
            line_class = (line_detail.get("ClassRef") or {}).get("name", "")
            if line_class:
                class_name = line_class
            else:
                bill_class = (bill.get("ClassRef") or {}).get("name", "")
                class_name = bill_class

            matched: Optional[dict] = None
            match_basis = ""
            if cust_id and division:
                # Concatenate description + account so the RP pump filter can
                # detect pump bills regardless of which field carries the cue.
                bill_text = f"{line_desc}  {account_name}"
                matched, match_basis = find_matching_invoice_ex(
                    bill_date, division, cust_id, invoices_by_customer,
                    bill_text=bill_text, bill_amount=line_amt,
                    project_num=project_num, gl_contracts=gl_contracts,
                )

            inv_doc = ""
            inv_id = ""
            inv_memo = ""
            inv_date: Optional[dt.date] = None
            inv_total: Optional[float] = None
            inv_balance: Optional[float] = None
            payment_date: Optional[dt.date] = None
            if matched:
                inv_doc = matched.get("DocNumber", "") or ""
                inv_id = matched.get("Id", "") or ""
                inv_memo = (matched.get("PrivateNote") or "").strip()
                inv_date = parse_date(matched.get("TxnDate"))
                inv_total = float(matched.get("TotalAmt") or 0)
                inv_balance = float(matched.get("Balance") or 0)
                payment_date = payment_map.get(inv_id) if inv_id else None

            # GC parent customer name = everything before the first ':' in
            # the line's CustomerRef (e.g. 'Forestar:RP7234 Project Name').
            # If there's no colon, the customer IS the GC (rare).
            gc_name = (cust_name.split(":", 1)[0] or "").strip() if cust_name else ""

            if bill_balance == 0:
                auto_status = STATUS_PAID
            elif not division:
                auto_status = STATUS_NO_PROJECT
            else:
                auto_status = compute_status(bill, matched, division)

            # Two-axis split (the user 2026-07-13): pay = did WE pay the vendor;
            # invoice = did the GC fund us (computed independent of payment).
            pay_status = compute_pay_status(bill)
            invoice_status = compute_invoice_status(matched, division)

            rows.append({
                "key": f"{bill_id}-{line_id}",
                "bill_id": bill_id,
                "line_id": line_id,
                "bill_date": bill_date,
                "vendor": vendor_name,
                "vendor_id": v_ref.get("value", ""),
                "bill_doc": bill_doc,
                "po_num": po_num,
                "bill_total": bill_total,
                "bill_balance": bill_balance,
                "division": division or "",
                "project_num": project_num or "",
                "bill_type": bill_type,
                "account": account_name,
                "line_amount": line_amt,
                "line_desc": line_desc,
                "inv_doc": inv_doc,
                "inv_id": inv_id,
                "inv_memo": inv_memo,
                "inv_date": inv_date,
                "inv_total": inv_total,
                "inv_balance": inv_balance,
                "payment_date": payment_date,
                "approved": approved,
                "gc_name": gc_name,
                "customer_name": cust_name,
                "class_name": class_name,
                "auto_status": auto_status,
                "pay_status": pay_status,
                "invoice_status": invoice_status,
                "match_basis": match_basis,
            })
    return rows


# ─────────────────────── collapse helpers ───────────────────────

# Sentinel string shown on the Bills sheet for multi-project bills' Project#,
# Division, Account columns. The Inventory sheet holds the per-project detail.
MULTI_MARKER = "(multiple)"


def multi_project_bill_ids(line_rows: List[dict]) -> Set[str]:
    """Return the set of bill_ids whose lines code to 2+ distinct project_nums.

    These are the "inventory" bills the user described — one supplier ticket
    distributed across multiple jobs. They get displayed as a single summary
    row on the master Bills sheet AND line-level on the Inventory drill-down.
    Single-project bills aren't in this set.

    Bills with no project on any line don't count as multi-project.
    """
    by_bill: Dict[str, Set[str]] = defaultdict(set)
    for r in line_rows:
        bid = r.get("bill_id") or ""
        proj = (r.get("project_num") or "").strip()
        if proj:
            by_bill[bid].add(proj)
    return {bid for bid, projs in by_bill.items() if len(projs) >= 2}


def _agg_distinct_or_multi(values: List[str]) -> str:
    """Return the single value if all non-empty values are identical, else
    return MULTI_MARKER. Empty strings are ignored when comparing."""
    distinct = {v for v in values if v}
    if not distinct:
        return ""
    if len(distinct) == 1:
        return next(iter(distinct))
    return MULTI_MARKER


# Word-boundary tax detector. Matches "tax", "taxes", "taxable", and the
# common compounds "sales tax", "use tax" as standalone words in either the
# line account or the line description. Word-boundary anchors on both ends
# so "Texarkana", "Texas", "taxi" don't false-positive.
# 2026-06-04: extended to "tax(es|able)?" after the user hit clerks typing "TAXES"
# (plural) on Martin Marietta yardage bills.
TAX_RE = re.compile(
    r"\b(?:sales\s+tax(?:es|able)?|use\s+tax(?:es|able)?|tax(?:es|able)?)\b",
    re.IGNORECASE,
)


def _is_tax_line(line: dict) -> bool:
    """True if this line looks like a sales/use tax line.

    Checks both the QBO account and the line description. Account-based
    detection catches the common case (account "Sales Tax Payable" etc.);
    description fallback catches lines where the clerk left the account
    generic but typed "Tax" in the memo.
    """
    acct = (line.get("account") or "").strip()
    desc = (line.get("line_desc") or "").strip()
    return bool(TAX_RE.search(acct) or TAX_RE.search(desc))


def _collapse_line_desc(lines: List[dict]) -> str:
    """Build a display description from N collapsed lines.

    Tax lines get pulled out of the "(+N more)" count and surfaced as a
    " + tax" suffix instead. Reason: most bills are one main item + one
    tax line, and showing "(+1 more)" for tax forced the AP person to
    mentally translate "+1 more = tax" every time. With this rule:

        1 line:                   "50 yds"
        main + tax:               "50 yds + tax"
        main + delivery + tax:    "50 yds  (+1 more) + tax"
        3 main + 1 tax:           "first desc  (+2 more) + tax"
        all tax (edge case):      "(tax only)"

    The summed Line Amount still includes tax in every case — this only
    changes the display string, not the math.
    """
    tax_lines = [l for l in lines if _is_tax_line(l)]
    main_lines = [l for l in lines if not _is_tax_line(l)]
    has_tax = bool(tax_lines)

    # Defensive: if every line on the bill detected as tax (would be weird),
    # don't drop them entirely — surface the situation so it's investigable.
    if not main_lines:
        return "(tax only)"

    descs = [(l.get("line_desc") or "").strip() for l in main_lines]
    nonempty = [d for d in descs if d]

    if len(main_lines) == 1 and nonempty:
        base = nonempty[0]
    elif nonempty:
        base = f"{nonempty[0]}  (+{len(main_lines) - 1} more)"
    elif len(main_lines) > 1:
        base = f"({len(main_lines)} lines)"
    else:
        base = ""

    if has_tax:
        return f"{base} + tax" if base else "+ tax"
    return base


def _aggregate_bill_status(line_statuses: List[str]) -> str:
    """Roll per-line statuses into a single bill-level status string.

    For a multi-project bill, each line independently matched its own
    project's invoice — so the lines can carry DIFFERENT statuses. The Bills
    sheet shows one row per bill, so we need to collapse the set into one
    label. The user 2026-06-04: the "some paid, some not" mixed state becomes
    `Partial paid`, surfacing the cash-position decision (float remaining
    out of operating, or wait for the last GC invoice to pay).

    Priority (most pessimistic / most informative wins):
      - any line has no project   → No project #
      - all lines are Bill paid   → Bill paid
      - any Invoice paid + any other → Partial paid  (NEW)
      - all lines Invoice paid    → Invoice paid
      - any Awaiting Invoice      → Awaiting Invoice
      - otherwise                 → Awaiting Payment
    """
    statuses = set(line_statuses)
    if not statuses:
        return STATUS_AWAITING_PAYMENT  # defensive default

    if STATUS_NO_PROJECT in statuses:
        return STATUS_NO_PROJECT
    if statuses == {STATUS_PAID}:
        return STATUS_PAID
    # Mixed-state with at least one paid line → Partial paid
    if STATUS_OK_TO_PAY in statuses and len(statuses) > 1:
        return STATUS_PARTIAL_PAID
    if statuses == {STATUS_OK_TO_PAY}:
        return STATUS_OK_TO_PAY
    if STATUS_AWAITING_INVOICE in statuses:
        return STATUS_AWAITING_INVOICE
    return STATUS_AWAITING_PAYMENT


def _aggregate_invoice_status(line_statuses: List[str]) -> str:
    """Roll per-line AR (invoice) statuses into one bill-level label. Same shape
    as _aggregate_bill_status but AR-only (never 'Bill paid' — that's the pay
    axis): mixed funded/unfunded → Partial paid."""
    s = {v for v in line_statuses if v}
    if not s:
        return STATUS_AWAITING_PAYMENT
    if STATUS_NO_PROJECT in s:
        return STATUS_NO_PROJECT
    if s == {STATUS_OK_TO_PAY}:
        return STATUS_OK_TO_PAY
    if STATUS_OK_TO_PAY in s and len(s) > 1:
        return STATUS_PARTIAL_PAID
    if STATUS_AWAITING_INVOICE in s:
        return STATUS_AWAITING_INVOICE
    return STATUS_AWAITING_PAYMENT


def collapse_rows(line_rows: List[dict], grain: str = "bill") -> List[dict]:
    """Roll up line-level rows into bill-grain or (bill, project)-grain rows.

    grain="bill" — One row per bill_id.
        Single-project bills: Project# / Division / Account keep their true
        values. Multi-project bills: those three columns become "(multiple)".
        Used by the master Bills sheet so each bill is one row regardless of
        how many lines QBO splits it across.

    grain="bill_project" — One row per (bill_id, project_num).
        Single-project bills produce one row (same as bill grain).
        Multi-project bills produce one row PER project chunk, each summing
        only that project's lines. Used by Division sheets and any
        cost-allocation view.

    Aggregation rules:
        * line_amount → SUM across the collapsed group
        * bill-level fields (vendor, bill_doc, bill_date, bill_balance,
          bill_total, auto_status, approved, gc_name, customer_name) →
          take from the first line; these are identical across lines of a bill
        * invoice fields → first line in the group (a (bill, project) chunk
          maps to a single matched invoice)
        * account, class_name → single value if all match, else "(multiple)"
        * line_desc → first non-empty + "(+N more)" badge
        * key → bill_id (grain="bill") or f"{bill_id}|{project_num}"
          (grain="bill_project")
        * line_id → "" — the collapsed row no longer represents a single line
        * is_multi_project (new) → bool flag; True iff lines span 2+ projects
        * line_count (new) → number of source lines folded in
    """
    if grain not in ("bill", "bill_project"):
        raise ValueError(f"unknown collapse grain: {grain!r}")

    if grain == "bill":
        def _grouper(r: dict) -> str:
            return r.get("bill_id") or ""
    else:
        def _grouper(r: dict) -> Tuple[str, str]:
            return (r.get("bill_id") or "", r.get("project_num") or "")

    # Pre-compute per-bill total line amount — used to pro-rate the bill's
    # open balance across project chunks on grain="bill_project". Without
    # pro-rata, summing bill_balance across multi-project chunks would
    # double-count.
    total_lines_per_bill: Dict[str, float] = defaultdict(float)
    # Also pre-compute the set of distinct projects per bill_id, because
    # "is multi-project" can ONLY be detected at the whole-bill level (within
    # a single bill_project group, all lines have the same project).
    projects_per_bill: Dict[str, Set[str]] = defaultdict(set)
    for r in line_rows:
        bid = r.get("bill_id") or ""
        total_lines_per_bill[bid] += float(r.get("line_amount") or 0)
        proj = (r.get("project_num") or "").strip()
        if proj:
            projects_per_bill[bid].add(proj)

    groups: Dict = defaultdict(list)
    seen_order: List = []
    for r in line_rows:
        k = _grouper(r)
        if k not in groups:
            seen_order.append(k)
        groups[k].append(r)

    collapsed: List[dict] = []
    for gk in seen_order:
        lines = groups[gk]
        first = lines[0]
        bill_id = first.get("bill_id") or ""

        # is_multi must be evaluated at the WHOLE-BILL level, not within
        # the current group: a bill_project group only contains lines for
        # ONE project, so an in-group check would always return False.
        is_multi = len(projects_per_bill.get(bill_id, set())) >= 2

        chunk_line_sum = sum(float(r.get("line_amount") or 0) for r in lines)
        bill_total_lines = total_lines_per_bill.get(bill_id, 0.0)
        # Pro-rata: chunk's share of the bill's lines. 1.0 for single-project.
        share = (chunk_line_sum / bill_total_lines
                 if (bill_total_lines and grain == "bill_project" and is_multi)
                 else 1.0)

        full_balance = first.get("bill_balance")
        full_total = first.get("bill_total")
        if share < 1.0:
            chunk_balance = (float(full_balance) * share) if full_balance is not None else None
            chunk_total = (float(full_total) * share) if full_total is not None else None
        else:
            chunk_balance = full_balance
            chunk_total = full_total

        if grain == "bill":
            if is_multi:
                project_num = MULTI_MARKER
                division = MULTI_MARKER
            else:
                project_num = first.get("project_num") or ""
                division = first.get("division") or ""
            account = _agg_distinct_or_multi([r.get("account") or "" for r in lines])
            class_name = _agg_distinct_or_multi([r.get("class_name") or "" for r in lines])
            # Bill grain: aggregate per-line statuses (mixed paid/unpaid →
            # Partial paid). For single-project bills the lines share the
            # same status, so aggregation returns it unchanged.
            agg_status = _aggregate_bill_status(
                [r.get("auto_status") or "" for r in lines]
            )
            key = bill_id
        else:  # bill_project
            project_num = first.get("project_num") or ""
            division = first.get("division") or ""
            account = _agg_distinct_or_multi([r.get("account") or "" for r in lines])
            class_name = _agg_distinct_or_multi([r.get("class_name") or "" for r in lines])
            # bill_project chunks always carry one project's status — no
            # aggregation needed; trust the first line in the group.
            agg_status = first.get("auto_status") or ""
            key = f"{bill_id}|{project_num}"

        collapsed.append({
            "key": key,
            "bill_id": bill_id,
            "line_id": "",
            "bill_date": first.get("bill_date"),
            "vendor": first.get("vendor", ""),
            "vendor_id": first.get("vendor_id", ""),
            "bill_doc": first.get("bill_doc", ""),
            "po_num": first.get("po_num", ""),
            "bill_total": chunk_total,
            "bill_balance": chunk_balance,
            "division": division,
            "project_num": project_num,
            "bill_type": first.get("bill_type", ""),
            "account": account,
            "line_amount": chunk_line_sum,
            "line_desc": _collapse_line_desc(lines),
            "inv_doc": first.get("inv_doc", ""),
            "inv_id": first.get("inv_id", ""),
            "inv_memo": first.get("inv_memo", ""),
            "inv_date": first.get("inv_date"),
            "inv_total": first.get("inv_total"),
            "inv_balance": first.get("inv_balance"),
            "payment_date": first.get("payment_date"),
            "approved": first.get("approved", False),
            "gc_name": first.get("gc_name", ""),
            "customer_name": first.get("customer_name", ""),
            "class_name": class_name,
            "auto_status": agg_status,
            # Pay status is bill-level (one balance) → same on every line.
            # Invoice status is per-line (per matched invoice) → aggregate.
            "pay_status": first.get("pay_status", ""),
            "invoice_status": (
                _aggregate_invoice_status([r.get("invoice_status") or "" for r in lines])
                if grain == "bill" else first.get("invoice_status", "")
            ),
            # match_basis is per-line (per matched invoice) → aggregate like account.
            "match_basis": (
                _agg_distinct_or_multi([r.get("match_basis") or "" for r in lines])
                if grain == "bill" else first.get("match_basis", "")
            ),
            "is_multi_project": is_multi,
            "line_count": len(lines),
            "share": share,  # 1.0 for single-project; <1.0 for multi-project chunks
        })
    return collapsed
