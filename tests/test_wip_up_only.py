"""The WIP update's up-only gate (the user 2026-09-29): costs / billed to date from QuickBooks only
go UP on the WIP. A drop is flagged `blocked` in the review and never written, even when approved -
in every division (MFD / CP since 2026-09-09, RP too now)."""
import sys
from pathlib import Path
from types import SimpleNamespace

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "wip"))

import wip_review_common as wrc   # noqa: E402


def _row(pn, costs, billed, retainage=None):
    return SimpleNamespace(project_num=pn, project_name=pn, costs_to_date=costs, billed_to_date=billed,
                           retainage_held=retainage, base_contract=None, co_revenue=None, base_etc=None,
                           co_cost_override=None)


def test_is_blocked():
    assert wrc.is_blocked("costs", 100.0, 90.0)
    assert wrc.is_blocked("billed", 100.0, None)           # going blank is a drop too
    assert not wrc.is_blocked("costs", 100.0, 100.004)     # within a cent
    assert not wrc.is_blocked("costs", 100.0, 150.0)
    assert not wrc.is_blocked("costs", None, 5.0)          # new job
    assert not wrc.is_blocked("retainage", 100.0, 0.0)     # retainage is released on purpose


def test_diff_flags_the_drop_in_every_division():
    for pn in ("RP6000", "CP700", "MFD200"):
        prior = {pn: {"costs": 1000.0, "billed": 2000.0}}
        rec = wrc.diff_rows([_row(pn, 900.0, 2500.0)], prior, division="x", tab_name="t", tab_kind="master")[0]
        f = {c["key"]: c for c in rec["fields"]}
        assert f["costs"]["blocked"] and f["costs"]["note"] == wrc.BLOCKED_NOTE
        assert not f["billed"]["blocked"]


def test_apply_never_writes_a_drop_even_when_approved():
    prior = {"RP6000": {"costs": 1000.0, "billed": 2000.0}}
    rows = [_row("RP6000", 900.0, 2500.0)]
    dec = {"fields": {"RP6000": {"costs": {"approved": True}, "billed": {"approved": True}}}}
    out = wrc.apply_decisions(rows, dec, prior, tab_kind="master")[0]
    assert out.costs_to_date == 1000.0       # held at the tab
    assert out.billed_to_date == 2500.0      # went up - written
