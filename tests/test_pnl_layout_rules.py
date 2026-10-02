"""project-pnl layout review (2026-10-02, rounds 1-23): the rules the owner
corrected, pinned so a later layout change cannot quietly undo them.

- Billed is GROSS (retainage included): a draw's gross = net billed + retained.
- Every TO-DATE overhead is rate x BILLED, never the contract: the draw rows,
  the coverage TOTAL, and the P&L-by-account check against them. (The
  projection alone uses the contract.)
- The draw coverage sits on the P&L, newest draw first, each name linking to
  its own sheet; the P&L checks live on Reconciliations under the QuickBooks
  tie-out, every formula sheet-qualified.
- Dollar figures carry a $.
"""
import importlib.util
import sys
from pathlib import Path

from openpyxl import Workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "project-pnl"))
_spec = importlib.util.spec_from_file_location(
    "project_pnl_export", ROOT / "project-pnl" / "project_pnl_export.py")
pnl = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(pnl)

# (name, period, net billed, costs, retainage held, retainage billed)
DRAWS = [("Draw 1", "11/01/25–11/30/25", 29190.38, 1828.81, 3243.38, 0.0),
         ("Draw 2", "01/01/26–01/30/26", 126674.10, 69547.35, 14074.90, 0.0)]


def _coverage(ws, top=20, contract=None):
    return pnl.write_draw_coverage(
        ws, top, 1, DRAWS, {}, {"__outside": {"total": 500.0}}, None,
        overhead_pct=10.0, contract_ref=contract,
        sheet_links={"__next": "Next Draw", "11/01/25–11/30/25": "Draw 1",
                     "01/01/26–01/30/26": "Draw 2"})


def _row_of(ws, text):
    return next(r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value == text)


def test_coverage_gross_includes_retainage_and_overhead_is_on_billed():
    wb = Workbook()
    ws = wb.active
    ws.title = "P&L"
    cov = _coverage(ws)
    r2 = _row_of(ws, "Draw 2")
    assert ws.cell(r2, 5).value == round(126674.10 + 14074.90, 2)    # gross billed
    assert ws.cell(r2, 9).value == f"=E{r2}*0.1"                     # OH = 10% x gross billed
    assert ws.cell(r2, 10).value == f"=E{r2}-F{r2}-I{r2}"            # net = billed - costs - OH
    tot = _row_of(ws, "TOTAL  (draws)")
    assert cov["cov_oh_tot"] == f"'P&L'!I{tot}"
    assert cov["cov_gross_tot"] == f"'P&L'!E{tot}"


def test_coverage_is_newest_first_and_links_to_each_draw_sheet():
    wb = Workbook()
    ws = wb.active
    ws.title = "P&L"
    _coverage(ws, contract="'P&L'!B4")
    nxt, d2, d1 = (_row_of(ws, t) for t in ("Next draw (forming)", "Draw 2", "Draw 1"))
    assert nxt < d2 < d1
    assert ws.cell(d2, 1).hyperlink.target == "#'Draw 2'!B2"
    assert ws.cell(nxt, 1).hyperlink.target == "#'Next Draw'!B2"
    # % Billed: this draw and every older one, over the contract
    assert ws.cell(d2, 12).value == f"=IF('P&L'!B4=0,\"\",SUM(E{d2}:E{d1})/'P&L'!B4)"


def test_checks_are_sheet_qualified_and_compare_overhead_on_billed():
    wb = Workbook()
    ws = wb.active
    ws.title = "P&L"
    cov = _coverage(ws)
    refs = {"billed": "C5", "costs": "C6", "opex": "N4", "ret": "I8",
            "oh_billed": "R6", **{k: v for k, v in cov.items() if k.startswith("cov_")}}
    checks = pnl._wire_pl_support(wb, None, DRAWS, {}, refs)
    by = {c[0]: c for c in checks}
    assert by["Draws gross billed = Billed to Date"][1] == "'P&L'!C5"
    assert by["② overhead (on billed) = draws overhead"][1:3] == ("-'P&L'!R6", cov["cov_oh_tot"])
    assert by["Draws retained = retainage still owed"][1] == "'P&L'!I8"
    for _label, plf, supf, _link in checks:
        assert "'P&L'!" in plf and "!" in supf


def test_checks_join_the_reconciliations_table():
    wb = Workbook()
    checks = [("Draws gross billed = Billed to Date", "'P&L'!C5", "'P&L'!E30", None)]
    pnl.build_sheet_reconciliations(
        wb, "CP000", {"name": "Test"}, {}, 100.0, 50.0, 0.0,
        {"billed": "Transactions!D5", "cogs": "Transactions!D6", "exp": "Transactions!D7"},
        "10/02/2026", pl_checks=checks)
    ws = wb["Reconciliations"]
    hdr = _row_of(ws, "P&L ties to its support")
    assert ws.cell(hdr + 1, 1).value == checks[0][0]
    assert ws.cell(hdr + 1, 2).value == "='P&L'!C5"
    status = next(ws.cell(r, 2).value for r in range(1, ws.max_row + 1)
                  if ws.cell(r, 1).value == "RECONCILIATION STATUS")
    assert f"ABS(D{hdr + 1})<0.005" in status
    assert ws.cell(hdr + 1, 2).number_format.startswith('"$"')


def test_dollar_columns_widen_for_the_sign():
    wb = Workbook()
    ws = wb.active
    ws.column_dimensions["B"].width = 12
    c = ws.cell(1, 2, value=1234.5)
    c.number_format = pnl.CURR_FMT
    pnl._widen_for_dollar(wb)
    assert ws.column_dimensions["B"].width == 13.5
    assert pnl.CURR_FMT.startswith('"$"') and pnl.RET_FMT.startswith('"$"')


def test_billing_outside_the_draws_joins_the_checks():
    # CP585's shape: three untagged pre-period invoices, a release outside every
    # draw window, and not-billed retainage that was never placed on a draw
    groups = {"__untagged": {"net_billed": 96777.0, "retainage_held": 0.0,
                             "retainage_billed": 0.0},
              "__retainage_billed": {"net_billed": 12663.0, "retainage_held": 0.0,
                                     "retainage_billed": 12663.0},
              "__retainage": {"total": 10753.0}}
    out = pnl._outside_draws(groups)
    assert out == {"gross": 107530.0, "ret": -1910.0}
    wb = Workbook()
    ws = wb.active
    ws.title = "P&L"
    cov = _coverage(ws)
    refs = {"billed": "C5", "costs": "C6", "ret": "I8", "oh_billed": "R6",
            "outside": out, "oh_rate": 0.10,
            **{k: v for k, v in cov.items() if k.startswith("cov_")}}
    by = {c[0].split(" = ")[0]: c for c in pnl._wire_pl_support(wb, None, DRAWS, {}, refs)}
    assert by["Draws gross billed + 107,530.00 outside the draws"][2] == cov["cov_gross_tot"] + "+107530.0"
    assert by["Draws retained - 1,910.00 outside the draws"][2] == cov["cov_ret_tot"] + "-1910.0"
    assert by["② overhead (on billed)"][2] == cov["cov_oh_tot"] + "+10753.0"


def test_only_untagged_invoices_after_the_first_draw_are_forming():
    groups = {"__untagged": {"invoices": [{"date": "2025-01-20", "doc_num": "old"},
                                          {"date": "2026-09-25", "doc_num": "new"}]}}
    costs = {"__disregarded": {"anchor": "2026-04-20"}}
    assert [i["doc_num"] for i in pnl._forming_invoices(groups, costs)] == ["new"]
