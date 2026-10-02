"""project-pnl _spread_not_billed_retainage - a "Retainage not billed" invoice goes back
onto the draws that were entered net, and a slice that repeats a retainage line already
on another invoice is a duplicate, never counted (CP790 #33659, 2026-10-02)."""
import importlib.util
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
_spec = importlib.util.spec_from_file_location(
    "project_pnl_export", ROOT / "project-pnl" / "project_pnl_export.py")
pnl = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(pnl)


def _inv(doc, date, gross, ret):
    return {"doc_num": doc, "date": date, "gross": gross, "retainage": ret,
            "retainage_billed": 0.0, "amount": round(gross - ret, 2)}


def _draw(*invs):
    return {"invoices": list(invs), "gross_income": sum(i["gross"] for i in invs),
            "retainage_held": sum(i["retainage"] for i in invs)}


def _cp790():
    return {
        "d1": _draw(_inv("33658", "2025-11-07", 29190.38, 0.0)),
        "d2": _draw(_inv("33898", "2026-01-13", 126674.10, 0.0)),
        "d3": _draw(_inv("33995", "2026-02-10", 54018.24, 5401.82)),
        "__retainage": {"invoices": [{"doc_num": "33659", "amount": 22720.10}],
                        "total": 22720.10},
    }


def test_cp790_spread_matches_the_pay_apps():
    g = _cp790()
    pnl._spread_not_billed_retainage(g)
    d1, d2 = g["d1"]["invoices"][0], g["d2"]["invoices"][0]
    assert (d1["gross"], d1["retainage"]) == (32433.76, 3243.38)     # G702 Draw 1
    assert (d2["gross"], d2["retainage"]) == (140749.0, 14074.9)
    assert d1["amount"] == 29190.38                                   # what the GC paid
    blk = g["__retainage"]
    assert blk["placed"] == 17318.28
    assert (round(blk["duplicate"], 2), blk["dup_doc"]) == (5401.82, "33995")
    assert blk["total"] == 0.0                                        # nothing left to count


def test_exact_pool_leaves_nothing_and_no_duplicate():
    g = {"d1": _draw(_inv("1", "2025-01-01", 90.0, 0.0)),
         "__retainage": {"invoices": [{"doc_num": "9", "amount": 10.0}], "total": 10.0}}
    pnl._spread_not_billed_retainage(g)
    assert g["d1"]["invoices"][0]["gross"] == 100.0
    assert g["__retainage"]["total"] == 0.0 and g["__retainage"]["duplicate"] == 0.0


def test_unmatched_leftover_still_counts():
    g = {"d1": _draw(_inv("1", "2025-01-01", 90.0, 0.0)),
         "d2": _draw(_inv("2", "2025-02-01", 500.0, 50.0)),
         "__retainage": {"invoices": [{"doc_num": "9", "amount": 30.0}], "total": 30.0}}
    pnl._spread_not_billed_retainage(g)
    assert g["__retainage"]["placed"] == 10.0
    assert g["__retainage"]["duplicate"] == 0.0                       # 20 matches no line
    assert g["__retainage"]["total"] == 20.0


def test_oldest_draws_fill_first_when_the_pool_is_short():
    g = {"d2": _draw(_inv("2", "2025-02-01", 90.0, 0.0)),
         "d1": _draw(_inv("1", "2025-01-01", 90.0, 0.0)),
         "__retainage": {"invoices": [{"doc_num": "9", "amount": 12.0}], "total": 12.0}}
    pnl._spread_not_billed_retainage(g)
    assert g["d1"]["invoices"][0]["retainage"] == 10.0
    assert g["d2"]["invoices"][0]["retainage"] == 2.0
    assert g["__retainage"]["total"] == 0.0
