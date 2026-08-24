"""
job_lines.py — the ONE test for "does this expense line belong to this job".

Modern jobs are project-coded: the line's own `CustomerRef` is the project, and
that is the whole answer. Jobs that predate consistent project coding are not.
On those, part of the cost sits on the project, part is named only in the LINE's
description, and part only in the BILL's memo — a plain customer test can run
several percent short on a multi-million job and make an over-budget job read as
on-budget.

Two modes, one object:

  STRICT (default)  the line's `CustomerRef` is the project customer. Byte-for-
                    byte what every caller did before this module existed.

  LEGACY            first rule that fires wins:
                      1. project    line `CustomerRef` is the project customer
                      2. line text  line Description or line `CustomerRef.name`
                                    names the job
                      3. bill note  the TXN's `PrivateNote` names the job AND
                                    names exactly ONE job number AND the line's
                                    own text names no job at all

GUARD on rule 3 (binding): a memo listing more than one job number — which
shared pump and material vendors do constantly — is SKIPPED, never split.
Missing a shared bill is cheaper than booking another job's money to this one.

Legacy mode is always opt-in. Turning it on for a modern job is harmless but
pointless; leaving it off for an old job silently under-reports it.

Billing has its own asymmetry, handled by `invoice_belongs`: an older job's
invoices usually sit on the PARENT customer, so both customers are pulled and
the memo decides.
"""
from __future__ import annotations

import re
from typing import Iterable, Optional

from shared.qbo_api import PROJ_RE


def job_pattern(project: str, aliases: Iterable[str] = ()) -> re.Pattern:
    """The job # plus any name it goes by on the street ('Bonds Ranch').
    Whitespace inside an alias matches any run of whitespace."""
    parts = [re.escape(project)]
    for a in aliases or ():
        a = (a or "").strip()
        if a:
            parts.append(r"\s*".join(re.escape(w) for w in a.split()))
    return re.compile("|".join(parts), re.IGNORECASE)


def jobs_named_in(text: str) -> set:
    """Every distinct job NUMBER appearing in a piece of text."""
    return {m.upper() for m in PROJ_RE.findall(text or "")}


def _line_text(det: dict, ln: dict) -> str:
    cref = det.get("CustomerRef") or {}
    return f"{(ln.get('Description') or '').strip()} {cref.get('name') or ''}"


class JobMatcher:
    """Callable line test for one job. `rule()` says which rule fired."""

    def __init__(self, customer_id: str, project: str = "",
                 aliases: Iterable[str] = (), legacy: bool = False):
        self.customer_id = str(customer_id)
        self.project = (project or "").upper()
        self.legacy = bool(legacy and self.project)
        self.pattern = job_pattern(self.project, aliases) if self.project else None

    def rule(self, det: dict, ln: dict, txn: dict) -> Optional[str]:
        if (det.get("CustomerRef") or {}).get("value") == self.customer_id:
            return "project"
        if not self.legacy or self.pattern is None:
            return None
        text = _line_text(det, ln)
        if self.pattern.search(text):
            return "line text"
        memo = (txn.get("PrivateNote") or "").strip()
        if (self.pattern.search(memo)
                and len(jobs_named_in(memo)) == 1
                and not jobs_named_in(text)):
            return "bill note"
        return None

    def __call__(self, det: dict, ln: dict, txn: dict) -> bool:
        return self.rule(det, ln, txn) is not None

    def txn_has_line(self, txn: dict) -> bool:
        """Does any expense line of this txn belong to the job?"""
        for ln in txn.get("Line") or []:
            det = (ln.get("AccountBasedExpenseLineDetail")
                   or ln.get("ItemBasedExpenseLineDetail"))
            if det and self(det, ln, txn):
                return True
        return False

    def line_total(self, txn: dict) -> float:
        tot = 0.0
        for ln in txn.get("Line") or []:
            det = (ln.get("AccountBasedExpenseLineDetail")
                   or ln.get("ItemBasedExpenseLineDetail"))
            if det and self(det, ln, txn):
                tot += float(ln.get("Amount", 0) or 0)
        return tot

    def invoice_belongs(self, inv: dict) -> bool:
        """An invoice is the job's when its memo names the job. Used for the
        PARENT customer's invoices on older jobs; a voided invoice never
        counts (QBO zeroes it and prefixes the memo 'Voided - ')."""
        memo = (inv.get("PrivateNote") or "").strip()
        if memo.lower().startswith("voided"):
            return False
        if (inv.get("CustomerRef") or {}).get("value") == self.customer_id:
            return True
        if not self.legacy or self.pattern is None:
            return False
        return bool(self.pattern.search(memo))
