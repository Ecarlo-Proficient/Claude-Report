"""bill_approval.py - the ONE rule for "is this QBO bill approved?" (shared).

Used by the Bill Tracker (bill-tracker/bill_rows.py) and the statement reconciler,
so both say the same thing about the same bill (moved here 10/08/2026 when the
reconciler needed it - tools never import tools).

Bill approval moved into QBO's approval Workflow on 09/16/2026 (the user
2026-09-17). From then on AP no longer types NOT APPROVED at entry, and QBO's API
does NOT return a bill's approval status (checked 2026-09-18 at minor versions 70
and 75 on a bill sitting in the approval queue: no such field). Bills entered
before that still carry the paper-era memo tag until they are cleared.

Three states:
  not approved  the memo starts NOT APPROVED (the paper-era tag)
  check QBO     entered on/after APPROVAL_WORKFLOW_START and still unpaid - its
                approval lives in QBO Tasks and the API cannot say
  approved      everything else: an old-process bill with no tag, or a
                new-process bill that has been paid (QBO will not release a
                payment on a bill that is pending approval)
"""
from __future__ import annotations

import datetime as dt
from typing import Optional, Union

APPROVAL_WORKFLOW_START = dt.date(2026, 9, 16)
APPROVAL_NO = "not approved"
APPROVAL_CHECK = "check QBO"
APPROVAL_YES = "approved"
UNAPPROVED_TAG = "NOT APPROVED"


def memo_approved(memo: Optional[str]) -> bool:
    """False only when the memo STARTS with 'NOT APPROVED' (any case, leading
    whitespace allowed) - 'Approved 5/1 (was Not Approved)' is approved."""
    return not (memo or "").lstrip().upper().startswith(UNAPPROVED_TAG)


def _as_date(d: Union[dt.date, str, None]) -> Optional[dt.date]:
    if isinstance(d, dt.date):
        return d
    try:
        return dt.date.fromisoformat((d or "")[:10])
    except ValueError:
        return None


def approval_state(memo: Optional[str], created: Union[dt.date, str, None],
                   balance: Optional[float]) -> str:
    """The bill's state - APPROVAL_NO / APPROVAL_CHECK / APPROVAL_YES. `created` is
    the day the bill was ENTERED in QBO (MetaData.CreateTime), not its TxnDate."""
    if not memo_approved(memo):
        return APPROVAL_NO
    c = _as_date(created)
    try:
        open_amt = float(balance or 0)
    except (TypeError, ValueError):
        open_amt = 0.0
    if c and c >= APPROVAL_WORKFLOW_START and open_amt > 0:
        return APPROVAL_CHECK
    return APPROVAL_YES


def revised_reason(memo: str, allow_first_line: bool = False) -> str:
    """The reason off a `REVISED - <reason>` memo line, or "" when there is none.
    The PM writes it on the SECOND memo line (line 1 stays AP's project line), so
    line 1 is skipped by default: in the paper era AP itself opened line 1 with
    REVISED followed by the project line (`REVISED - CP861 - 4231 STATE HWY 161`),
    which is not a reason and not a vendor follow-up - 16 such bills sit in the
    2026 population. Pass allow_first_line=True for a bill entered under the QBO
    workflow, where a first-line REVISED can only be a PM who typed it in the wrong
    place. A bare REVISED returns "(no reason given)" so the bill still gets chased."""
    lines = (memo or "").splitlines()
    for line in (lines if allow_first_line else lines[1:]):
        t = line.strip()
        if t.upper().startswith("REVISED"):
            rest = t[len("REVISED"):].lstrip(" -:\u2013").strip()
            return rest or "(no reason given)"
    return ""
