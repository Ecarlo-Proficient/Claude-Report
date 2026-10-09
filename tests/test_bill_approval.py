"""shared/bill_approval - the ONE approval rule (Bill Tracker + statement reconciler),
and the reconciler's workbook showing all three states (owner 10/08/2026: bills are
approved in QBO now, but older ones still carry NOT APPROVED)."""
import datetime as dt
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "statement-reconciler"))

from shared import bill_approval as ba  # noqa: E402


def test_three_states():
    assert ba.approval_state("NOT APPROVED - CP861", "2026-08-01", 10) == ba.APPROVAL_NO
    assert ba.approval_state("  not approved", dt.date(2026, 10, 1), 10) == ba.APPROVAL_NO   # tag wins
    assert ba.approval_state("CP861", "2026-10-01T09:00:00", 10) == ba.APPROVAL_CHECK         # new, unpaid
    assert ba.approval_state("CP861", "2026-10-01", 0) == ba.APPROVAL_YES                     # new, paid
    assert ba.approval_state("CP861", "2026-09-15", 10) == ba.APPROVAL_YES                    # old, no tag
    assert ba.approval_state("Approved 5/1 (was Not Approved)", "2026-08-01", 10) == ba.APPROVAL_YES
    assert ba.approval_state(None, "", None) == ba.APPROVAL_YES


def test_reconciler_workbook_shows_all_three(tmp_path):
    import statement_reconciler as sr
    from openpyxl import load_workbook
    bills = [sr.QboBill("1", "A1", "2026-08-01", 10.0, 10.0, memo="NOT APPROVED", created="2026-08-01"),
             sr.QboBill("2", "B2", "2026-10-01", 20.0, 20.0, memo="CP861", created="2026-10-01"),
             sr.QboBill("3", "C3", "2026-08-01", 30.0, 30.0, memo="CP861", created="2026-08-01")]
    rows = [sr.ReconRow("MATCHED", b.txn_date, b.doc_number, b.open_balance, b.open_balance, "", "",
                        stmt_ref=b.doc_number, qbo_ref=b.doc_number, stmt_date=b.txn_date,
                        qbo_date=b.txn_date, qbo_memo=b.memo, qbo_bill_id=b.bill_id) for b in bills]
    out = tmp_path / "r.xlsx"
    sr.write_excel(out, "V", "2026-10-07", 60.0, bills, rows, line_sum=60.0)
    ws = load_workbook(out)["Summary"]
    approved = {r[1]: r[10] for r in ws.iter_rows(values_only=True) if r[1] in ("A1", "B2", "C3")}
    assert approved == {"A1": "Not Approved", "B2": "Check QBO", "C3": "Yes"}
    text = " ".join(str(c) for r in ws.iter_rows(values_only=True) for c in r if c)
    assert "approval is in QBO" in text and "NOT APPROVED by memo" in text
    import statement_markup as sm
    assert [sm.bucket_for("MATCHED", s) for s in (ba.APPROVAL_NO, ba.APPROVAL_CHECK, ba.APPROVAL_YES)] == \
        ["approve", "checkqbo", "matched"]
