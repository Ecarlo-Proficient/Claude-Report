"""
draw_chain.py — order a project's draw invoices so "the previous draw" is answerable.

Why this exists (the user 2026-08-05):
    MFD and CP bill in draws, and the funding is a chain, not a set of
    independent invoices: **draw N is invoiced, the GC funds it, we pay draw N's
    vendor bills with that money, the vendors issue unconditional waivers, and
    the GC needs those waivers before it will fund draw N+1.**

    So when an open draw invoice isn't getting paid, the thing to look at is not
    its own bills — it's whether the PREVIOUS draw's bills are cleared. An
    unpaid vendor two draws back is what stops today's money.

Getting "previous" right is the whole job, and two things in the real data
break a naive sort by date:

  1. **Not every invoice on a project is a draw.** MFD177 has a
     `City Retainage` invoice dated between its April and May draws. Sorted by
     date it would become "the previous draw" — it isn't one. Only invoices
     whose memo names a draw enter the chain.
  2. **A project can run several contracts in parallel.** MFD192 bills a base
     contract, a `HUDSONWOOD CONTRACT`, and an `OFFSITE CONTRACT` on the same
     day each month. Their draws interleave perfectly by date.

For (2) the honest answer today is *we don't know*: the user 2026-08-05 —
"MFD192 - oddball, i have no definite way to see what bills belong to which
contract yet." Bills carry a project #, not a contract. So a multi-contract
project is REPORTED AS UNATTRIBUTABLE rather than split on a guess — see
`CHAIN_MULTI_CONTRACT`. Splitting it would produce a confident answer built on
an attribution nobody can currently make.

Lookback is the immediately-previous draw only (the user's choice), matching
how the waiver actually gates the next funding.
"""
from __future__ import annotations

import datetime as dt
import re
from collections import defaultdict
from typing import Dict, List, Optional, Tuple

# Chain outcomes for a single invoice.
CHAIN_NOT_A_DRAW = "not-a-draw"          # retainage, one-offs — nothing gates it
CHAIN_FIRST_DRAW = "first-draw"          # provably the first (Draw #1)
CHAIN_MULTI_CONTRACT = "multi-contract"  # parallel contracts, bills unattributable
CHAIN_HAS_PREV = "has-prev"
# Earliest draw we can SEE, but NOT provably the first: a numbered draw past #1
# whose predecessor never entered Notion (the open-invoice sync only ever created
# pages for draws that were open while it was running; a Draw #1 paid before then
# is invisible), or an unnumbered MFD draw where we can't confirm it's first. Was
# mislabelled "First draw" until 2026-08-12.
CHAIN_PREV_UNKNOWN = "prev-unknown"

_MONTHS = (
    "january|february|march|april|may|june|july|august|september|october|"
    "november|december|jan|feb|mar|apr|jun|jul|aug|sep|sept|oct|nov|dec"
)

# The draw designator, in every shape the memos actually use. Whatever survives
# in FRONT of it names the contract, so this has to strip all of them or two
# spellings of one contract look like two contracts:
#   "May Draw 2026" · "Draw #2" · "Draw # 6" · "Draw #3 December 2024"
#   "March 2025 Draw" · "Draw #1 -"  (trailing dash left by the period split)
_DRAW_TAIL = re.compile(
    rf"[-–—]?\s*(?:(?:{_MONTHS})\s+(?:20\d\d\s+)?)?draw\s*#?\s*\d*"
    rf"\s*(?:(?:{_MONTHS})\s*)?(?:20\d\d)?\s*[-–—]?\s*$",
    re.IGNORECASE,
)
# A retainage draw is the SAME contract's final draw, not a separate contract —
# "…- Retainage - Draw #5" must land on the same chain as "…- Draw #4".
_RETAINAGE_TAIL = re.compile(r"[-–—]?\s*retainage\s*[-–—]?\s*$", re.IGNORECASE)
# Periods appear as "(Period: …", "(Period:…", and bare "- Period:…".
_PERIOD = re.compile(r"\(?\s*period\s*:", re.IGNORECASE)
_HAS_DRAW = re.compile(r"\bdraws?\b", re.IGNORECASE)
_DRAW_NUM = re.compile(r"\bdraw\s*#\s*(\d+)", re.IGNORECASE)


def memo_head(memo: str, project_num: str = "") -> str:
    """Memo with the period block, line breaks and leading project # removed."""
    text = " ".join((memo or "").split())
    head = _PERIOD.split(text)[0].strip()
    if project_num:
        head = re.sub(
            rf"^\s*{re.escape(project_num)}\s*[-–—:]?\s*", "", head, flags=re.IGNORECASE
        )
    return head.strip(" -–—")


def is_draw(memo: str) -> bool:
    """True when this invoice is a draw at all (vs retainage, a one-off, etc.)."""
    return bool(_HAS_DRAW.search(memo_head(memo)))


def contract_label(memo: str, project_num: str = "") -> str:
    """The contract a draw belongs to — the memo with its draw designator cut off.

    "2100 S. Mayhill Rd - OFFSITE CONTRACT - July Draw 2026" → "2100 S. Mayhill
    Rd - OFFSITE CONTRACT", while the base contract's July draw yields
    "2100 S. Mayhill Rd". Two different labels on one project # = parallel
    contracts.

    Strips the draw and retainage tails REPEATEDLY, in either order: a
    "…- Draw #8 - Retainage" memo has the retainage AFTER the draw, so a single
    draw-then-retainage pass would leave "Draw #8" stuck in the label and split
    the retainage draw off its own contract's chain (fixed 2026-08-12).
    """
    head = memo_head(memo, project_num)
    prev = None
    while head != prev:
        prev = head
        head = _DRAW_TAIL.sub("", head)
        head = _RETAINAGE_TAIL.sub("", head)
        head = head.strip(" -–—")
    return head.upper()


def draw_number(memo: str) -> Optional[int]:
    """The "#2" in "Draw #2" — CP numbers its draws, MFD names them by month."""
    m = _DRAW_NUM.search(memo or "")
    return int(m.group(1)) if m else None


class DrawChains:
    """Per-(project, contract) ordered draw sequences built from every invoice.

    Feed it EVERY invoice on a project — paid ones included. The previous draw
    is almost always already paid, so a chain built from open invoices alone
    would be blind exactly where it needs to see.
    """

    def __init__(self) -> None:
        self._chains: Dict[Tuple[str, str], List[dict]] = defaultdict(list)
        self._contracts_per_project: Dict[str, set] = defaultdict(set)
        self._by_invoice: Dict[str, dict] = {}

    def add(
        self,
        *,
        invoice_num: str,
        project_num: str,
        memo: str,
        invoice_date: Optional[dt.date],
        is_paid: bool,
    ) -> None:
        if not project_num or not invoice_num:
            return
        rec = {
            "invoice_num": str(invoice_num).strip(),
            "project_num": project_num,
            "memo": memo or "",
            "date": invoice_date,
            "is_paid": is_paid,
            "draw_no": draw_number(memo),
        }
        self._by_invoice[rec["invoice_num"]] = rec
        if not is_draw(memo):
            return
        contract = contract_label(memo, project_num)
        rec["contract"] = contract
        self._chains[(project_num, contract)].append(rec)
        self._contracts_per_project[project_num].add(contract)

    def finalize(self) -> None:
        """Sort each chain oldest→newest. Draw # wins where present (CP numbers
        its draws); otherwise invoice date, with invoice # as a stable tiebreak
        for same-day draws."""
        for key, rows in self._chains.items():
            rows.sort(
                key=lambda r: (
                    r["draw_no"] if r["draw_no"] is not None else 10**6,
                    r["date"] or dt.date.min,
                    r["invoice_num"],
                )
            )

    def is_multi_contract(self, project_num: str) -> bool:
        return len(self._contracts_per_project.get(project_num, ())) > 1

    def previous_draw(self, invoice_num: str) -> Tuple[str, Optional[dict]]:
        """(outcome, previous draw record or None) for one invoice number."""
        rec = self._by_invoice.get(str(invoice_num).strip())
        if rec is None or not is_draw(rec["memo"]):
            return CHAIN_NOT_A_DRAW, None
        # Parallel contracts on one project # — bills can't be attributed to a
        # contract today, so refuse to name a previous draw rather than guess.
        if self.is_multi_contract(rec["project_num"]):
            return CHAIN_MULTI_CONTRACT, None
        chain = self._chains.get((rec["project_num"], rec["contract"]), [])
        for idx, row in enumerate(chain):
            if row["invoice_num"] == rec["invoice_num"]:
                if idx == 0:
                    return self._earliest_seen_outcome(rec), None
                return CHAIN_HAS_PREV, chain[idx - 1]
        return self._earliest_seen_outcome(rec), None

    @staticmethod
    def _earliest_seen_outcome(rec: dict) -> str:
        """Outcome for a draw with nothing before it in the chain. Only Draw #1 is
        provably first; a numbered draw past #1 (predecessor never synced) or an
        unnumbered draw is 'previous unknown', not 'first'. (the user 2026-08-12)"""
        return CHAIN_FIRST_DRAW if rec.get("draw_no") == 1 else CHAIN_PREV_UNKNOWN
