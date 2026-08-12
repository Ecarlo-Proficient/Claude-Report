"""
lien_clock.py — the Texas lien notice clock, in ONE place.

Texas Prop. Code Ch. 53, first-tier subcontractor (we contract with GCs, not
owners). This was `health-dashboard/money_bleeds.py`'s private logic until the
AR Aging tab needed the same dates; per repo rule 3 it moved here rather than
being copied or cross-imported. **money_bleeds is the origin — its behaviour is
the reference. Do not "improve" the rules here without changing it there too.**

The deadlines
-------------
    Commercial / nonresidential (CP, MFD) — notice to owner AND original
        contractor by the **15th of the 3rd month** after each work month.
    Residential (RP) — by the **15th of the 2nd month** after each work month.

    MFD sits on the COMMERCIAL clock on purpose. Ch. 53's "residence" is a
    single-family house, duplex, triplex, quadruplex, or a unit whose title
    transfers individually — an apartment complex is none of those.

    Deadlines roll **BACKWARD** to the prior business day, never forward
    (weekends only here; holidays are not modelled).

Work month = INVOICE month (the user 2026-07-16, settled — do not re-add a
conservative offset). RP invoices go out the day the job finishes; draws bill
their own work month. The first build used an earlier month "to be safe" and
produced month-early false alarms, which is why this is written down.

Retainage
---------
    Retainage has its OWN deadline: the notice is due **30 days after the invoice
    date** (the user 2026-08-12, correcting the earlier "own track, not dated"
    treatment). § 53.057's retainage notice runs from completion, and the owner
    invoices retainage at completion, so invoice date + 30 days is the date to
    watch. Retainage is dated and banded like everything else here, just off a
    different clock; its label carries a `RET` marker.

What this clock does NOT cover
------------------------------
    * **Equipment-lease / note-payment invoices to subs are not construction
      income** (the user 2026-07-16) and carry no lien rights. Detecting those
      needs QBO line items; a caller with only a memo cannot see them — see
      `is_lease_text`.

This is a deadline *watchlist*, not legal advice. Project type, parcel, and
owning entity all have to be verified before anything is actually sent.
"""
from __future__ import annotations

import datetime as dt
import re
from typing import NamedTuple, Optional, Tuple

# Months after the work month in which notice is due, by division.
NOTICE_MONTHS = {"MFD": 3, "CP": 3, "RP": 2}
NOTICE_DAY = 15

# Retainage runs off a different clock: notice due 30 days after the invoice date
# (the user 2026-08-12). Invoiced at completion, so invoice date + 30 = § 53.057.
RETAINAGE_NOTICE_DAYS = 30

# Urgency bands (money_bleeds' values — keep them in step).
URGENT_DAYS = 15
WATCH_DAYS = 45

STATE_PAST = "PAST"
STATE_URGENT = "URGENT"
STATE_WATCH = "WATCH"
STATE_OK = "OK"
STATE_RETAINAGE = "RETAINAGE"
STATE_SENT = "SENT"

RETAINAGE_RE = re.compile(r"retainage|retention|retenci[oó]n", re.IGNORECASE)

# Equipment lease / note payment — no lien rights ride on these. Text-only
# match; the authoritative check needs QBO line items (money_bleeds does that).
NON_CONSTRUCTION_RE = re.compile(
    r"equipment lease|monthly equipment|note principal|principal payment"
    r"|interest (charge|fee)|lease payment",
    re.IGNORECASE,
)

# A human note saying the notice already went out. Deliberately narrow: it must
# name the notice or the affidavit, so "waiting on vendor unconditional" (a
# note about someone else's waiver) doesn't read as our notice being sent.
NOTICE_SENT_RE = re.compile(
    r"\b(lien\s+(notice|affidavit|filed)|notice\s+(sent|mailed|served)"
    r"|sent\s+lien|filed\s+lien)\b",
    re.IGNORECASE,
)


def add_months(year: int, month: int, n: int) -> Tuple[int, int]:
    m = month - 1 + n
    return year + m // 12, m % 12 + 1


def roll_back_weekend(d: dt.date) -> dt.date:
    """Statutory deadlines roll BACKWARD to the prior business day, never
    forward. (Holidays are not modelled — treat a Monday deadline with care.)"""
    while d.weekday() >= 5:  # 5=Sat, 6=Sun
        d -= dt.timedelta(days=1)
    return d


def notice_deadline(work_year: int, work_month: int, division: str) -> dt.date:
    """The date the notice must be MAILED by (service completes on mailing,
    § 53.003(c)). An unknown division gets the shorter residential clock —
    erring toward the earlier date is the safe direction for a deadline."""
    n = NOTICE_MONTHS.get(division, NOTICE_MONTHS["RP"])
    y, m = add_months(work_year, work_month, n)
    return roll_back_weekend(dt.date(y, m, NOTICE_DAY))


def is_retainage_text(*texts: str) -> bool:
    return any(RETAINAGE_RE.search(t) for t in texts if t)


def is_lease_text(*texts: str) -> bool:
    return any(NON_CONSTRUCTION_RE.search(t) for t in texts if t)


def notice_already_sent(*texts: str) -> bool:
    return any(NOTICE_SENT_RE.search(t) for t in texts if t)


class LienState(NamedTuple):
    state: str                       # one of the STATE_* constants
    deadline: Optional[dt.date]      # None when no monthly clock applies
    days_left: Optional[int]         # negative once past
    label: str                       # ready to drop in a cell

    @property
    def needs_action(self) -> bool:
        return self.state in (STATE_PAST, STATE_URGENT)


def _fmt(d: dt.date) -> str:
    """Month-day-year, abbreviated month. Never year-first."""
    return d.strftime("%b %d, %Y").replace(" 0", " ")


def retainage_deadline(invoice_date: dt.date) -> dt.date:
    """Notice deadline for a retainage invoice: 30 days after the invoice date."""
    return invoice_date + dt.timedelta(days=RETAINAGE_NOTICE_DAYS)


def _band(deadline: dt.date, today: dt.date, prefix: str = "") -> LienState:
    """Urgency band + label for a deadline. `prefix` marks a non-standard clock
    (e.g. 'RET' for retainage)."""
    days = (deadline - today).days
    p = f"{prefix} " if prefix else ""
    if days < 0:
        return LienState(STATE_PAST, deadline, days, f"{p}PAST DUE · {_fmt(deadline)}")
    if days <= URGENT_DAYS:
        return LienState(STATE_URGENT, deadline, days, f"{p}DUE {_fmt(deadline)} · {days}d")
    if days <= WATCH_DAYS:
        return LienState(STATE_WATCH, deadline, days, f"{p}{_fmt(deadline)} · {days}d")
    return LienState(STATE_OK, deadline, days, f"{p}{_fmt(deadline)}")


def lien_state(
    division: str,
    invoice_date: Optional[dt.date],
    today: dt.date,
    *,
    memo: str = "",
    note: str = "",
) -> LienState:
    """The lien-notice position of one open invoice.

    `memo` and `note` are free text (the QBO memo and the collections clerk's
    Quick Status) used to spot retainage and an already-sent notice. Both are
    text heuristics — a retainage invoice whose memo doesn't say "retainage"
    will be given a monthly deadline it may not be governed by.
    """
    if invoice_date is None:
        return LienState(STATE_OK, None, None, "")

    if notice_already_sent(note, memo):
        return LienState(STATE_SENT, None, None, "Notice sent")

    if is_retainage_text(memo, note):
        # Retainage notice is due 30 days after the invoice date (the user
        # 2026-08-12). Dated and banded like everything else, off its own clock;
        # the 'RET' prefix marks which clock it is.
        return _band(retainage_deadline(invoice_date), today, prefix="RET")

    deadline = notice_deadline(invoice_date.year, invoice_date.month, division)
    return _band(deadline, today)
