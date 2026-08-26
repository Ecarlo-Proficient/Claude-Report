"""
mfd_wip_cols.py - the column CONTRACT for the MFD division WIP tab.

WHY THIS FILE EXISTS (read before "simplifying" it back into constants)
The first cut of mfd_wip_test.py hardcoded positions - COL_ETC = 14, COL_CTC = 20.
Every column the owner added afterwards became surgery: shift the constants, re-merge
the banner, write a migration guard for tabs built by the previous version. Three of
the four rebuild cycles on 2026-08-25 were that, not changed requirements.

So column ORDER is data now, exactly like wip_writer.COLS. Reordering the list below
reorders the sheet; adding a column is one line. Positions are DERIVED - never write
a column letter or index anywhere else in the MFD tooling, ask this module.

THE GROUPING (the owner, 2026-08-25): everything MFD types sits in ONE run so entry is
a single left-to-right pass with no calculated cell interrupting it; the QBO block is
next; the metrics that drive decisions are furthest right.

ONE ETC, FOR THE WHOLE CONTRACT (the owner, 2026-08-26 - settled, do not re-litigate)
There is deliberately NO 'original ETC + CO costs = revised ETC' trio here, even though
wip_writer and project-pnl both build one. The owner's ruling: "just use one contract and
one ETC ... im not going to make them put the etc of the CO." So the single ETC column
means the estimated total cost of the ENTIRE contract, change orders included.

That also avoids a live double-count. The ETC on the sheet came from WIP Master's
`=(E/1.17)` divisor fallback, computed off a contract figure that ALREADY carried ~91% of
that job's change orders. Adding a separate CO-cost column on top of it would have counted
most of the CO cost twice and driven GP% badly negative. One column, one meaning.

KIND tells the writer what a column IS, and that drives both styling and safety:
  carry  - MFD's, script never writes it, only moves it during a reorder
  input  - MFD types it; bold-orange-on-grey, their existing convention
  qbo    - the script owns it; green header, tinted, "do not type here"
  calc   - an Excel formula the script owns
"""
from __future__ import annotations

from typing import Dict, List, Optional, Tuple

# (header, width, key, group label or "", kind)
COLS: List[Tuple[str, int, str, str, str]] = [
    # ── who and when ──────────────────────────────────────────────────────
    ("PROJECT",                      44, "project",      "",           "carry"),
    ("MOBE DATE",                    15, "mobe",         "",           "carry"),
    ("COMPLETION DATE SOG/PAVING",   17, "completion",   "",           "carry"),
    ("CUSTOMER",                     17, "customer",     "",           "carry"),

    # ── ONE typing run. Do not put a calculated column inside this block. ──
    ("CONTRACT",                     15, "contract",     "MFD ENTERS", "input"),
    ("CHANGE ORDERS",                15, "co",           "",           "input"),
    ("ETC",                          15, "etc",          "",           "input"),
    ("COMPLETED TO DATE",            16, "completed",    "",           "input"),
    ("EARNED LESS RET.",             16, "earned_less",  "",           "input"),
    ("Total Retainage",              15, "retainage",    "",           "input"),

    # ── the script's, from QBO. Sync stamp merges across this block. ──────
    ("COSTS TO DATE",                17, "qbo_costs",    "FROM QBO",   "qbo"),
    ("BILLED TO DATE",               17, "qbo_billed",   "",           "qbo"),
    ("RETAINAGE (QBO)",              17, "qbo_retain",   "",           "qbo"),

    # ── what the numbers mean, most decision-useful furthest right ────────
    ("REV. CONTRACT",                15, "rev_contract", "METRICS",    "calc"),
    ("EARNED REVENUE",               15, "earned_rev",   "",           "calc"),
    ("% COMPLETE",                   12, "pct",          "",           "calc"),
    ("BALANCE TO FINISH INCL'N RET", 16, "balance",      "",           "calc"),
    ("COST TO COMPLETE",             17, "ctc",          "",           "calc"),
    ("BILLED AHEAD",                 15, "over",         "",           "calc"),
    ("BILLED BEHIND",                15, "under",        "",           "calc"),
    ("GP %",                         12, "gp",           "",           "calc"),
]

FIRST_COL = 2                      # column B - column A is the narrow spacer

KEYS = [c[2] for c in COLS]
BY_KEY: Dict[str, tuple] = {c[2]: c for c in COLS}
INPUT_KEYS = [c[2] for c in COLS if c[4] == "input"]
QBO_KEYS = [c[2] for c in COLS if c[4] == "qbo"]
CALC_KEYS = [c[2] for c in COLS if c[4] == "calc"]
CARRY_KEYS = [c[2] for c in COLS if c[4] == "carry"]
# Every column whose value MFD owns - carried through a reorder untouched.
OWNED_BY_MFD = CARRY_KEYS + INPUT_KEYS


def index(key: str) -> int:
    """1-based worksheet column index for a key."""
    return FIRST_COL + KEYS.index(key)


def letter(key: str) -> str:
    from openpyxl.utils import get_column_letter
    return get_column_letter(index(key))


def kind(key: str) -> str:
    return BY_KEY[key][4]


def header(key: str) -> str:
    return BY_KEY[key][0]


def last_index() -> int:
    return FIRST_COL + len(COLS) - 1


def group_starts() -> Dict[str, str]:
    """{group label -> the key it starts on}, for the banner and section rules."""
    return {c[3]: c[2] for c in COLS if c[3]}


def group_keys(label: str) -> List[str]:
    """Every key belonging to a labelled group, in sheet order."""
    out: List[str] = []
    seen = False
    for hdr, _w, key, grp, _k in COLS:
        if grp == label:
            seen = True
        elif grp and seen:
            break
        if seen:
            out.append(key)
    return out


# ── formulas ──────────────────────────────────────────────────────────────
# One place, keyed by column. `L` resolves a key to its column letter, so a
# reorder rewrites every formula correctly with no edit here.
#
# EARNED REVENUE is cost-to-cost - contract x (costs / revised ETC) - which is
# the CPA/bank method and the ONLY basis on which billed-ahead/behind means
# anything. Deriving it from the tab's own % COMPLETE would make it equal
# COMPLETED TO DATE and both billing columns identically zero.
#
# % COMPLETE keeps the tab's OWN billing-based definition (completed / revised
# contract), which is not the standard's cost-based one. It predates this
# script and MFD reads it that way; changing it silently would be a defect.
#
# GP% is (revised contract - revised ETC) / revised contract. Costs to date must
# NEVER enter it - see the cardinal rule in the WIP standard: that column shows
# the BID, not the outcome.

def _f(key: str, r: int, L) -> Optional[str]:
    if key == "rev_contract":
        return f"={L('contract')}{r}+{L('co')}{r}"
    if key == "earned_rev":
        return (f'=IF(OR({L("etc")}{r}="",{L("etc")}{r}=0,'
                f'NOT(ISNUMBER({L("qbo_costs")}{r}))),"",'
                f'{L("rev_contract")}{r}*{L("qbo_costs")}{r}/{L("etc")}{r})')
    if key == "pct":
        return f'=IF({L("rev_contract")}{r}=0,0,{L("completed")}{r}/{L("rev_contract")}{r})'
    if key == "balance":
        return f'={L("rev_contract")}{r}-{L("earned_less")}{r}'
    if key == "ctc":
        return (f'=IF(OR({L("etc")}{r}=0,NOT(ISNUMBER({L("qbo_costs")}{r}))),"",'
                f'{L("etc")}{r}-{L("qbo_costs")}{r})')
    if key == "over":
        return (f'=IF({L("earned_rev")}{r}="","",'
                f'MAX({L("completed")}{r}-{L("earned_rev")}{r},0))')
    if key == "under":
        return (f'=IF({L("earned_rev")}{r}="","",'
                f'MAX({L("earned_rev")}{r}-{L("completed")}{r},0))')
    if key == "gp":
        return (f'=IF({L("rev_contract")}{r}=0,0,'
                f'({L("rev_contract")}{r}-{L("etc")}{r})/{L("rev_contract")}{r})')
    return None


def formula(key: str, row: int) -> Optional[str]:
    return _f(key, row, letter)


# Totals row: money columns subtotal, ratios recompute off the totals.
TOTAL_SUM_KEYS = ["contract", "co", "etc", "completed", "earned_less",
                  "retainage", "qbo_costs", "qbo_billed", "qbo_retain",
                  "rev_contract", "earned_rev", "balance", "ctc", "over", "under"]


def total_formula(key: str, row: int, first: int, last: int) -> Optional[str]:
    L = letter
    if key in TOTAL_SUM_KEYS:
        return f"=SUM({L(key)}{first}:{L(key)}{last})"
    if key == "pct":
        return f'=IF({L("rev_contract")}{row}=0,0,{L("completed")}{row}/{L("rev_contract")}{row})'
    if key == "gp":
        return (f'=IF({L("rev_contract")}{row}=0,0,'
                f'({L("rev_contract")}{row}-{L("etc")}{row})/{L("rev_contract")}{row})')
    return None


PCT_KEYS = frozenset({"pct", "gp"})
DATE_KEYS = frozenset({"mobe", "completion"})
TEXT_KEYS = frozenset({"project", "customer"})
