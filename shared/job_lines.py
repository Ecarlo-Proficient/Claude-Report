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
                      2. class      the line's `ClassRef` sits under the job's
                                    OWN class branch (opt-in, see below)
                      3. line text  line Description or line `CustomerRef.name`
                                    names the job
                      4. bill note  the TXN's `PrivateNote` names the job AND
                                    names exactly ONE job number AND the line's
                                    own text names no job at all

GUARD on rule 3 (binding): a memo listing more than one job number — which
shared pump and material vendors do constantly — is SKIPPED, never split.
Missing a shared bill is cheaper than booking another job's money to this one.

Legacy mode is always opt-in. Turning it on for a modern job is harmless but
pointless; leaving it off for an old job silently under-reports it.

CLASS is opt-in per job and never inferred, because it is the easiest rule to
get wrong. Two traps, both seen on MFD228 (2026-08-25):

  * **The job's class is often INACTIVE.** A plain `SELECT * FROM Class` returns
    active classes only, so the branch you find may be the parent while every
    cost line carries a deleted child (`…:MARKER LAPIZ:MFD228 (deleted)`).
    Query classes with `Active IN (true, false)` or match on the branch PREFIX,
    which catches the live parent and the deleted leaf together.
  * **A DIVISION or BUILDER class is not a job class.** `MULTI FAMILY` or
    `Residential:CUSTOM HOMES OF TEXAS` sweeps in every other job under it.
    `class_prefix` therefore REFUSES a bare division name outright — pass the
    branch that identifies this job and nothing else.

Billing has its own asymmetry, handled by `invoice_belongs`: an older job's
invoices usually sit on the PARENT customer, so both customers are pulled and
the memo decides.
"""
from __future__ import annotations

import re
from typing import Iterable, Optional

from shared.qbo_api import PROJ_RE


_JOB_SPLIT_RE = re.compile(r"^([A-Z]+)\s*(\d+)\s*(?:[-\s]*(FTW|[A-Z]{2,6}))?$",
                           re.IGNORECASE)


def job_number_pattern(project: str) -> str:
    """Regex source matching THIS job number and no other.

    Two things people get wrong when they hand-roll this:

    * **The separator drifts.** Clerks write `MFD228`, `MFD 228`, `MFD-228`.
      A literal match on the canonical spelling silently drops the rest -
      it cost MFD228 $6,680 across 9 lines before this existed (2026-08-25).
    * **A suffix is a DIFFERENT JOB.** `RP7186` and `RP7186-FTW` are two
      projects and must never absorb each other (CLAUDE.md, binding). The base
      number therefore carries a negative lookahead for a suffix, and a
      suffixed job matches only with its suffix. `(?!\d)` stops `MFD228` from
      matching `MFD2281`. A letter glued straight on with no separator
      (`MFD228A`) still matches - there is no such job-numbering convention,
      so in real data that is an annotation of this job, not another one.
    """
    m = _JOB_SPLIT_RE.match((project or "").strip())
    if not m:
        return re.escape(project or "")
    pfx, num, suffix = m.group(1), m.group(2), m.group(3)
    base = rf"{re.escape(pfx)}\s*[-]?\s*{re.escape(num)}(?!\d)"
    if suffix:
        return base + rf"[\s-]*{re.escape(suffix)}"
    # Reject only a real SUFFIX. It is spelled hyphen-attached with NO spaces
    # (`RP7186-FTW`, per PROJ_RE), so the guard must NOT fire on the ordinary
    # memo form `MFD172 - 1392 E Bonds Ranch Rd` — a spaced hyphen is a
    # separator between fields, not a suffix. Getting this wrong silently
    # dropped 48 real lines before it was caught (2026-08-25). FTW is also
    # rejected in any spacing, since the schedule writes `RP7186 FTW`.
    return base + r"(?!-[A-Za-z])(?!\s*-?\s*FTW\b)"


def job_pattern(project: str, aliases: Iterable[str] = ()) -> re.Pattern:
    """The job # (separator-tolerant, suffix-exact) plus any name it goes by on
    the street ('Bonds Ranch'). Whitespace inside an alias matches any run of
    whitespace."""
    parts = [job_number_pattern(project)]
    for a in aliases or ():
        a = (a or "").strip()
        if a:
            parts.append(r"\s*".join(re.escape(w) for w in a.split()))
    return re.compile("|".join(parts), re.IGNORECASE)


def jobs_named_in(text: str) -> set:
    """Every distinct job NUMBER appearing in a piece of text."""
    return {m.upper() for m in PROJ_RE.findall(text or "")}


# A class at this level names a DIVISION, never a job — passing one as a job
# class would attribute every multifamily/residential/commercial line to one job.
_DIVISION_CLASSES = {"multi family", "multifamily", "residential", "commercial"}


def _line_class(det: dict, ln: dict, txn: dict) -> str:
    """The class that applies to this line. Line level wins — a bill can split
    across divisions, so the txn's class is only a fallback."""
    ref = (det.get("ClassRef") or ln.get("ClassRef") or txn.get("ClassRef") or {})
    return ref.get("name") or ""


def _line_text(det: dict, ln: dict) -> str:
    cref = det.get("CustomerRef") or {}
    return f"{(ln.get('Description') or '').strip()} {cref.get('name') or ''}"


class JobMatcher:
    """Callable line test for one job. `rule()` says which rule fired."""

    def __init__(self, customer_id: str, project: str = "",
                 aliases: Iterable[str] = (), legacy: bool = False,
                 class_prefix: str = ""):
        self.customer_id = str(customer_id)
        self.project = (project or "").upper()
        self.legacy = bool(legacy and self.project)
        self.pattern = job_pattern(self.project, aliases) if self.project else None
        cp = (class_prefix or "").strip()
        if cp and cp.strip(": ").lower() in _DIVISION_CLASSES:
            raise ValueError(
                f"{cp!r} is a DIVISION class, not a job class — it would claim "
                f"every job in that division. Pass the job's own class branch.")
        self.class_prefix = cp.upper()

    def rule(self, det: dict, ln: dict, txn: dict) -> Optional[str]:
        if (det.get("CustomerRef") or {}).get("value") == self.customer_id:
            return "project"
        if not self.legacy or self.pattern is None:
            return None
        if self.class_prefix:
            # Prefix, not equality: the live parent branch and the deleted
            # per-job leaf beneath it are the same job.
            if _line_class(det, ln, txn).upper().startswith(self.class_prefix):
                return "class"
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
