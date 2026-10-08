"""Statement layouts learned from real statements (10/07/2026 Inbox sweep). Each
test holds a short sample of the text the parser actually sees, so a later change
can't quietly break a vendor. Amounts and numbers are illustrative; no names.
Rule for new layouts: add a sample here when a parser is added or fixed."""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "statement-reconciler"))

import statement_reconciler as sr  # noqa: E402
import statement_set as ss  # noqa: E402


def _parse(text):
    tpl = sr.detect_template(text)
    fn = {
        "vendor_abatix_online": sr.parse_statement_abatix_online,
        "vendor_croell_ar": sr.parse_statement_croell_ar,
        "vendor_quikrete": sr.parse_statement_quikrete,
        "qbo_invoice_list": sr.parse_statement_qbo_invoice_list,
        "qbo_open_invoices": sr.parse_statement_qbo_open_invoices,
        "qbo_statement": sr.parse_statement_qbo_statement,
        "vendor_cowtown": sr.parse_statement_cowtown,
        "qbo_customer_open_balance": sr.parse_statement_qbo_customer_open_balance,
    }[tpl]
    v, d, due, lines = fn(text)
    return tpl, d, due, lines


def _ties(due, lines):
    return abs(round(sum(x.amount for x in lines), 2) - due) <= 0.5


def test_abatix_online_with_a_credit_in_parentheses():
    t = """STATEMENT
As of Date: 10/1/2026
Proficient Concrete LLC Remit to: Total Due $415.13
Number Date Amount
8924376 8/19/2026 9/18/2026 DG7676 $378.33 $378.33
8931407 9/2/2026 10/2/2026 DG7676 $38.97 ($1.17)
8939832 9/21/2026 10/21/2026 $44.33 $37.97
Summary of Invoice Age for 222803
Total Amount Due For Customer: $415.13"""
    tpl, d, due, lines = _parse(t)
    assert tpl == "vendor_abatix_online" and d == "2026-10-01" and due == 415.13
    assert [x.ref for x in lines] == ["8924376", "8931407", "8939832"] and lines[1].amount == -1.17
    assert _ties(due, lines)


def test_croell_ar_balance_invoice_number_runs_into_the_po():
    t = """DEV AR Customer Balance
Customer 112768 Proficient Concrete LLC Invoices & On Acct 7,435.08
Date 10/07/2026 Finance Charges 749.12
Balance After Disc 7,435.08
Invoice # PO# Invoice Date Disc Date Type Invoice Amount Pending Prev Applied Current Due Disc Offered Prev Disc TakenDisc Avail Current Due
1080594MC6749 05/13/2026 06/02/2026 I 14,846.55 0.00 14,830.11 16.44 811.88 0.00 0.00 16.44
1102761Finance Charge 07/06/2026 F 749.12 0.00 0.00 749.12 0.00 0.00 0.00 749.12
11057227 BREW GRANBURY 07/14/2026 08/03/2026 I 6,669.52 0.00 0.00 6,669.52 0.00 0.00 0.00 6,669.52"""
    tpl, d, due, lines = _parse(t)
    assert tpl == "vendor_croell_ar" and d == "2026-10-07"
    assert [x.ref for x in lines] == ["1080594", "1102761", "1105722"]
    assert lines[0].amount == 16.44 and _ties(due, lines)


def test_quikrete_account_statement():
    t = """ACCOUNT STATEMENT Page No. 1 of 1 Date Statement No.
10/06/26 1373049
AS OF 09/30/26
INVOICE NO. INVOICE DATE DUE DATE REMARK CUST P.O. # INVOICE AMOUNT BALANCE DUE
Ship-to: 1043868 JOB A
RI 34216970 07/17/26 08/16/26 29215167 JA7387 47,825.00 47,825.00
RI 34826201 09/18/26 10/18/26 29935556 P.O.#: Customer Invoice 3,331.94 3,331.94
Total for: JOB A 51,156.94
TOTAL USD 51,156.94"""
    tpl, d, due, lines = _parse(t)
    assert tpl == "vendor_quikrete" and d == "2026-09-30" and due == 51156.94
    assert [x.ref for x in lines] == ["34216970", "34826201"] and _ties(due, lines)


def test_qbo_invoices_for_list_counts_only_open_ones():
    t = """VENDOR - DFW
3:58 PM
Invoices for Proficient Concrete, LLC
10/05/26
Accrual Basis All Transactions
Num Date Name Amount Open Balance
DFW26-02974 08/20/2026 Proficient Concrete, LLC:7725 Black Elk Court 5,253.30 5,253.30
DFW26-02999 08/22/2026 Proficient Concrete, LLC:1 Paid St 100.00 0.00
DFW26-03099 08/31/2026 Proficient Concrete, LLC:802 Ranch Road 400.00 218.00
Total 5,753.30 5,471.30"""
    tpl, d, due, lines = _parse(t)
    assert tpl == "qbo_invoice_list" and d == "2026-10-05" and due == 5471.30
    assert [x.ref for x in lines] == ["DFW26-02974", "DFW26-03099"] and _ties(due, lines)


def test_open_invoices_report_for_every_customer_takes_only_ours_across_a_page_break():
    t = """Vendor Pumping LLC
Open Invoices Report
As of Oct 7, 2026
Date Transaction type Num Term Due date Open balance
OTHER CUSTOMER LLC
09/16/2026 Invoice 1414 Net 30 10/16/2026 120.00
Total for OTHER CUSTOMER LLC $120.00
Proficient Concrete, LLC
05/30/2026 Invoice 1209 Net 30 06/29/2026 126.00
Wednesday, October 07, 2026 01:49 PM GMT-06:00 2/3
Vendor Pumping LLC
Open Invoices Report
As of Oct 7, 2026
Date Transaction type Num Term Due date Open balance
07/16/2026 Invoice 1256 Net 30 08/15/2026 237.70
09/22/2026 Payment 25763 09/22/2026 -200.00
Total for Proficient Concrete, LLC $163.70
ANOTHER ONE
08/31/2026 Invoice 1387 Net 30 09/30/2026 88.00
Total for ANOTHER ONE $88.00
TOTAL $371.70"""
    tpl, d, due, lines = _parse(t)
    assert tpl == "qbo_open_invoices" and d == "2026-10-07"
    assert [x.ref for x in lines] == ["1209", "1256", "25763"] and due == 163.70 and _ties(due, lines)


def test_qbo_statement_revised_invoice_and_a_header_date_far_from_its_label():
    t = """Vendor Concrete Company
Statement
13501 N Highway 171
Date
Office Phone# 817-928-5401
Email: x@example.com 10/7/2026
To:
Proficient Concrete, LLC
Amount Due Amount Enc.
$500.00
Date Transaction Amount Balance
08/05/2026 INV #16541. Due 09/04/2026. PO #MC7626. Orig. Amount 100.00 100.00
09/03/2026 INV #R16738. Due 10/03/2026. PO #MC7923. Orig. Amount 400.00 500.00
"""
    tpl, d, due, lines = _parse(t)
    assert tpl == "qbo_statement" and d == "2026-10-07"
    assert [x.ref for x in lines] == ["16541", "R16738"] and _ties(due, lines)


def test_activity_statement_payment_without_a_check_number():
    t = """Vendor Redi-Mix
Date
PO Box 162327 10/2/2026
Terms Amount Due
Net 10th $150.00
Date Transaction Due Date Amount Balance
08/31/2026 Balance forward 100.00
09/21/2026 INV #401417. Due 10/10/2026. 70.00 170.00
09/21/2026 PMT -20.00 150.00
"""
    tpl, d, due, lines = _parse(t)
    assert tpl == "qbo_statement" and d == "2026-10-02"
    assert [(x.ref, x.amount) for x in lines] == [("", 100.0), ("401417", 70.0), ("PMT", -20.0)]
    assert _ties(due, lines)


def test_past_due_letter_reissued_invoice_with_a_letter_suffix():
    t = """October 1, 2026 3400 BETHLEHEM AVE.
Past Due Amount: $300.00
Job Inv. No. Inv. Date Due Date Inv. Balance
A St 399238R 08/25/2026 09/10/2026 $200.00 $200.00
B St 399326 08/24/2026 09/10/2026 $100.00 $100.00"""
    tpl, d, due, lines = _parse(t)
    assert tpl == "vendor_cowtown" and [x.ref for x in lines] == ["399238R", "399326"] and _ties(due, lines)


def test_open_balance_credit_type_truncated_to_credit_m():
    t = """Vendor Cable, Inc
Customer Open Balance
Accrual Basis As of August 31, 2026
Type Date Num Memo Due Date Open Balance
PROFICIENT CONCRETE, LLC.
JOB
Invoice 7/7/2026 RWC768503 9401 MERRITT RD 8/6/2026 1,000.00
Credit M... 5/13/2026 CMC747836 9401 MERRITT RD 6/12/2026 -491.73
Total JOB 508.27
TOTAL 508.27"""
    tpl, d, due, lines = _parse(t)
    assert tpl == "qbo_customer_open_balance"
    assert [x.amount for x in lines] == [1000.0, -491.73] and _ties(due, lines)


# ── safety nets ──

def test_a_statement_dated_before_its_own_invoices_is_not_used():
    @__import__("dataclasses").dataclass
    class L:
        date: str
        ref: str
        amount: float
    d = ss.classify(ss.Doc(path=Path("/x/a.pdf"), as_of="2026-08-05", amount_due=10.0,
                           lines=[L("2026-09-08", "1", 10.0)]))
    assert "before its own invoice" in d.problem
    m, _ = ss.plan([d])
    assert d.status == "unreadable" and not m.lines


def test_undated_excel_takes_the_date_in_its_name_never_today():
    assert ss.date_from_name("Statement BURNCO 06-09.xlsx", 1780000000) == "2026-06-09"
    assert ss.date_from_name("core qbo proficient ex 9-25-2026.xlsx") == "2026-09-25"
    assert ss.date_from_name("Statement SUNBELT.xlsx", 1780000000) == ""


def test_vendor_folder_matches_without_a_legal_suffix(tmp_path):
    (tmp_path / "Core Concrete Pumping").mkdir()
    (tmp_path / "Croell, Inc").mkdir()
    assert sr._vendor_dir(tmp_path, "Core Concrete Pumping LLC").name == "Core Concrete Pumping"
    assert sr._vendor_dir(tmp_path, "CROELL, INC").name == "Croell, Inc"
    assert sr._vendor_dir(tmp_path, "New Vendor LLC").name == "New Vendor LLC"


def test_voidform_customer_statement():
    t = """Customer Statement
Date: 8/31/2026
Date Due Date Doc. Type Ref. Nbr. Ext. Ref. Nbr. Orig. Amount Amount Due Balance
7/29/2026 8/28/2026 Invoice AR-0325087 OM7391 916.03 916.03 916.03
3724 Job St
8/20/2026 9/19/2026 Invoice AR-0325801 OM7660 332.17 332.17 1,248.20
8231 Other St
Current 1 - 30 Days Past Due 31 - 60 Days Past Due 61 - 90 Days Past Due Over 90 Days Past Due Amount Due
332.17 916.03 0.00 0.00 0.00 1,248.20"""
    assert sr.detect_template(t) == "vendor_voidform"
    v, d, due, lines = sr.parse_statement_voidform(t)
    assert d == "2026-08-31" and due == 1248.20
    assert [(x.ref, x.po, x.address) for x in lines] == [("AR-0325087", "OM7391", "3724 Job St"),
                                                        ("AR-0325801", "OM7660", "8231 Other St")]


def test_scanned_statement_amount_due_falls_back_to_the_running_balance():
    t = """VENDOR FOUNDATION
7/28/2026
[Amount Due [__AmountEne._}
04/24/2026 INV #6171. Due 04/24/2026. PO #7014. Orig. 404.00 404.00
05/20/2026 INV #6187. Due 05/20/2026. PO #4509. Orig. Amount 250.00 654.00
"""
    v, d, due, lines = sr.parse_statement_qbo_statement(t)
    assert due == 654.00 and _ties(due, lines)


# ── the marked-up statement (statement_markup.locate) ──

def _page(*texts, ocr=False):
    import statement_markup as sm
    rows = []
    for i, t in enumerate(texts):
        words = [sm.Word(w, x * 50, x * 50 + 40, i * 20, i * 20 + 12) for x, w in enumerate(t.split())]
        rows.append(sm.Row(i * 20, i * 20 + 12, words))
    return sm.Page(None, rows, "ocr" if ocr else "text")


class _L:
    def __init__(self, ref, amount, date=""):
        self.ref, self.amount, self.date = ref, amount, date


def test_markup_bands_every_printed_copy_of_one_invoice():
    import statement_markup as sm
    pg = _page("09/16/2026 4282533059 $ 48.71", "header", "09/16/2026 10/10/2026 4282533059 $ 48.71")
    where = sm.locate([pg], [_L("4282533059", 48.71)])
    assert where == [(0, 0)] and sm._EXTRA == {0: [(0, 2)]}


def test_markup_lookalike_refs_never_take_a_neighbours_row():
    import statement_markup as sm
    pg = _page("09/18/2024 CM_SUNO0232 976.80", "09/18/2024 CM_SUNO0233 250.80", ocr=True)
    where = sm.locate([pg], [_L("CM_SUN00233", -250.80), _L("CM_SUN00232", -976.80)])
    assert where == [(0, 1), (0, 0)]


def test_markup_zero_rows_and_totals_are_not_bills():
    import statement_markup as sm
    pg = _page("01/01/2020 Balance Forward 0.00 0.00", "07/16/2026 Discount 0.00 685,810.67",
               "Total 06/30/2026 1,000.00", "07/17/2026 147780 08/16/2026 792.07 792.07")
    assert [sm._bill_like(r) for r in pg.rows] == [False, False, False, True]
