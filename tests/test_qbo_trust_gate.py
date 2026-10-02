"""The WIP update's QuickBooks trust gate (the owner 2026-09-30: "we can be certain these are real
projects booked in qbo"). A job QuickBooks is not trusted for is HELD: its costs / billed are never
written (the tab keeps its number), a held job not yet on the tab is not added, and findings under
$1,000 only ride as a CHECK note. Offline - the Trust object is built by hand, no mirror needed."""
import sys
from pathlib import Path
from types import SimpleNamespace

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "wip"))

import wip_review_common as wrc          # noqa: E402
from shared import qbo_trust             # noqa: E402


def _row(pn, costs, billed):
    return SimpleNamespace(project_num=pn, project_name=pn, costs_to_date=costs, billed_to_date=billed,
                           retainage_held=None, base_contract=None, co_revenue=None, base_etc=None,
                           co_cost_override=None)


def _trust():
    t = qbo_trust.Trust()
    t.add("RP7401-FTW", ("costs", "billed"), "two customers")
    t.add("CP790", ("costs",), "cost with no project")
    t.add("CP672", ("costs",), "$200 with no project", hold=False)
    return t


def test_trust_reasons_and_notes():
    t = _trust()
    assert t.held("RP7401-FTW") and t.reasons("rp7401-ftw", "billed") == ["two customers"]
    assert t.reasons("CP790", "costs") and not t.reasons("CP790", "billed")
    assert not t.held("CP672") and t.notes("CP672", "costs") == ["$200 with no project"]
    assert not t.held("RP6000")


def test_global_reason_holds_every_job():
    t = qbo_trust.Trust(global_reason="stale")
    assert t.reasons("ANY1", "costs") == ["stale"] and t.reasons("ANY1", "billed") == ["stale"]
    assert not t.reasons("ANY1", "retainage")


def test_diff_marks_held_cells_and_check_notes():
    prior = {p: {"costs": 100.0, "billed": 100.0} for p in ("RP7401-FTW", "CP790", "CP672")}
    rows = [_row("RP7401-FTW", 200.0, 200.0), _row("CP790", 200.0, 200.0), _row("CP672", 200.0, 200.0)]
    recs = {r["project_num"]: r for r in wrc.diff_rows(rows, prior, division="x", tab_name="t",
                                                        tab_kind="working", trust=_trust())}
    f = {c["key"]: c for c in recs["CP790"]["fields"]}
    assert f["costs"]["held"] and f["costs"]["note"].startswith(wrc.HELD_PREFIX)
    assert not f["billed"]["held"]
    assert recs["RP7401-FTW"]["held"] and recs["CP790"]["held"] and not recs["CP672"]["held"]
    g = {c["key"]: c for c in recs["CP672"]["fields"]}
    assert not g["costs"]["held"] and g["costs"]["check"] == "$200 with no project"


def test_hold_keeps_the_tab_even_when_approved():
    prior = {"RP7401-FTW": {"costs": 100.0, "billed": 100.0}, "CP790": {"costs": 100.0, "billed": 100.0}}
    rows = [_row("RP7401-FTW", 200.0, 200.0), _row("CP790", 200.0, 300.0)]
    dec = {"fields": {p: {"costs": {"approved": True}, "billed": {"approved": True}} for p in prior}}
    out = {r.project_num: r for r in wrc.apply_decisions(rows, dec, prior, tab_kind="working", trust=_trust())}
    assert (out["RP7401-FTW"].costs_to_date, out["RP7401-FTW"].billed_to_date) == (100.0, 100.0)
    assert out["CP790"].costs_to_date == 100.0 and out["CP790"].billed_to_date == 300.0


def test_held_job_not_on_the_tab_is_not_added():
    kept = wrc.hold_untrusted([_row("RP7401-FTW", 5.0, 5.0), _row("RP6000", 5.0, 5.0)], _trust(), prior={})
    assert [r.project_num for r in kept] == ["RP6000"]


def test_no_trust_means_no_gate():
    rows = [_row("RP7401-FTW", 5.0, 5.0)]
    assert wrc.hold_untrusted(rows, None, prior={}) == rows


def test_customer_status_names_the_problem():
    cs = qbo_trust.customer_status
    assert cs("RP7340", {"Id": "2", "DisplayName": "RP7340 -FTW"}, "1", False, False) == "Name typo"
    assert cs("RP7152", {"Id": "2", "DisplayName": "RP7152-1"}, "1", True, False) == "Separate job?"
    assert cs("RP2583", {"Id": "2", "DisplayName": "RP2583"}, "1", True, False) == "Money on duplicate"
    assert cs("RP7074", {"Id": "2", "DisplayName": "RP7074"}, "1", False, False) == "Empty duplicate"
    assert cs("RP7074", {"Id": "1", "DisplayName": "RP7074"}, "1", False, False) == "Real"
    assert cs("RP7401-FTW", {"Id": "2", "DisplayName": "RP7401-FTW"}, "1", True, True) == "Both have invoices"
