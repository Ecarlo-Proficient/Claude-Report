#!/usr/bin/env python3
"""
rp_invoicing.py - ONE-INVOICE job or SCOPE-BASED job? Read it off the invoices.

The owner (2026-09-08): "we have two different types, a one invoice one job vs
a scope based invoices for one job like this one" (RP6586). The P&L treats
them differently, so the kind has to be decided one way everywhere - here,
shared by project-pnl (the sheet-per-invoice template) and the RP invoicing
scan (one-offs/rp_stage_scan.py).

WHAT THE INVOICES SAY (measured 2026-09-08 over 2,464 invoices / 876 RP jobs
since 2025-01-01):
  * EVERY invoice memo reads "RP#### - <address> - <what was billed>". The
    third segment is the invoice's description, not a stage marker - a
    one-invoice tract job says "Foundation" or "Foundation Turnkey" there.
    So the memo suffix alone tells you nothing about the kind.
  * A tract job is billed ONCE for the slab, then often again for EXTRAS -
    pump charges, rock saw, drop brickledge, HTT5 bolts, "Foundation Turnkey
    Extra", sealer, sleeves - small add-ons, usually days apart. That is still
    a one-invoice job: the main scope went out on one invoice.
  * A scope-based (multi-stage) job bills the SLAB ITSELF in pieces as each
    finishes - lot preparation, piers, mud slab and pier caps, grade beams,
    walls, foundation - weeks or months apart. RP6586: five stage invoices,
    3/11 -> 7/27/2026, from a proposal priced in four sections.

THE RULE: an invoice is a STAGE invoice when its description names a piece of
the slab build and is not an extra. Two or more stage invoices = scope-based.
One stage invoice that is a PARTIAL stage (piers, lot prep, grade beams, mud
slab, walls) = scope-based with one stage billed so far. One stage invoice
that is the whole slab ("Foundation ...") = a one-invoice job, extras or not.
"""
from __future__ import annotations

import datetime as dt
import re
from typing import Dict, List, Optional

_SEG_RE = re.compile(r"\s+[-–—]\s+")
# a piece of the slab build (any of these makes the invoice a STAGE invoice)
STAGE_RE = re.compile(
    r"lot\s*prep|site\s*prep|pier|mud\s*slab|grade\s*beam|foundation|slab|"
    r"stem\s*wall|retaining\s*wall|wall\s*pour|concrete\s*wall|footing|fondation",
    re.IGNORECASE)
# a partial stage - the slab is NOT done when this is all that was billed
PARTIAL_RE = re.compile(
    r"lot\s*prep|site\s*prep|pier|mud\s*slab|grade\s*beam|stem\s*wall|"
    r"retaining\s*wall|wall\s*pour|concrete\s*wall|footing",
    re.IGNORECASE)
# an add-on billed beside the slab, never a stage of it
EXTRA_RE = re.compile(
    r"extra|pump|rock\s*saw|brick\s*ledge|brickledge|bolt|htt|sealer|sleeve|"
    r"drop\s*garage|repair|patch|punch|retainage|change\s*order|\bco\b|"
    r"haul|dirt|fill|sand|demo|scratch|crack|cable|rebar",
    re.IGNORECASE)

KIND_ONE = "One invoice, one job"
KIND_ONE_EXTRAS = "One invoice + extras"
KIND_SCOPE = "Scope-based (multi-stage)"
KIND_SCOPE_1 = "Scope-based, 1 stage billed so far"
KIND_NONE = "Not invoiced"


def scope_of(memo: str) -> str:
    """The invoice's own description - the memo's third segment
    ("RP#### - <address> - <scope>"), '' when the memo stops short."""
    segs = [s.strip() for s in _SEG_RE.split(memo or "") if s.strip()]
    return segs[-1] if len(segs) >= 3 else ""


def is_stage(scope: str) -> bool:
    return bool(scope) and bool(STAGE_RE.search(scope)) and not EXTRA_RE.search(scope)


def is_partial_stage(scope: str) -> bool:
    return is_stage(scope) and bool(PARTIAL_RE.search(scope))


def _date(v) -> Optional[dt.date]:
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    try:
        return dt.date.fromisoformat(str(v)[:10])
    except (TypeError, ValueError):
        return None


def classify(invoices: List[dict]) -> Dict:
    """invoices: [{doc_num, date, memo, amount, ...}] for ONE job.
    -> {kind, multi_stage, invoices (dated, oldest first), scopes, stage_flags,
        n, n_stage, span_days, label}"""
    invs = sorted((i for i in invoices if _date(i.get("date"))),
                  key=lambda i: (_date(i["date"]), str(i.get("doc_num", ""))))
    scopes = [scope_of(i.get("memo") or "") for i in invs]
    flags = [is_stage(s) for s in scopes]
    n, n_stage = len(invs), sum(flags)
    span = (_date(invs[-1]["date"]) - _date(invs[0]["date"])).days if n >= 2 else 0
    if n == 0:
        kind = KIND_NONE
    elif n_stage >= 2:
        kind = KIND_SCOPE
    elif n_stage == 1 and any(is_partial_stage(s) for s in scopes):
        kind = KIND_SCOPE_1
    elif n >= 2:
        kind = KIND_ONE_EXTRAS
    else:
        kind = KIND_ONE
    multi = kind == KIND_SCOPE
    if multi:
        label = f"Scope-based invoicing: {n_stage} stage invoices of {n} over {span} days"
    elif kind == KIND_SCOPE_1:
        label = f"Scope-based invoicing: 1 stage billed so far ({scopes[flags.index(True)]})"
    elif kind == KIND_ONE_EXTRAS:
        label = f"One-invoice job + {n - 1} extra invoice(s)"
    elif kind == KIND_ONE:
        label = "One-invoice job"
    else:
        label = "Not invoiced yet"
    return {"kind": kind, "multi_stage": multi, "invoices": invs, "scopes": scopes,
            "stage_flags": flags, "n": n, "n_stage": n_stage, "span_days": span,
            "label": label}
