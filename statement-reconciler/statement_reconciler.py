#!/usr/bin/env python3
"""
statement_reconciler.py - Vendor Statement vs QBO Reconciler

Takes a vendor statement PDF, identifies the vendor, pulls every open
Bill in QBO for that vendor, and reconciles the two side-by-side.

Categorizes each statement line into one of four buckets:
  ✓ MATCHED              - same Ref#, same amount (clerk OK)
  ⚠ VENDOR_TAX_VIOLATION - same Ref#, amount diff is exactly 8.25% (TX
                            sales tax) of the QBO amount → vendor billed
                            tax in violation of Proficient/vendor no-tax
                            agreement. Vendor action, not clerk action.
  ✗ CLERK_AMOUNT_MISMATCH- same Ref#, amount diff is some other value
                            → clerk entered wrong amount in QBO.
  ✗ MISSING_IN_QBO       - on statement, no Bill in QBO with that Ref#
                            → clerk has not entered the bill yet.
  ✗ MISSING_ON_STATEMENT - Bill exists open in QBO, but vendor doesn't
                            show it on the statement → stale unpaid bill
                            vendor may have credited / already received
                            payment for; needs the user's eyes.

Writes an Excel report (Summary + one sheet per category) named
  Statement_Reconciliation_<date>_<vendor>.xlsx

Manually-passed files get the SAME treatment as an --inbox sweep: each Excel
lands in the Synology Reconciliations folder, and a source file that already
lives in the Statement Inbox is archived to its DONE subfolder on success. If
the Synology share isn't mounted, output falls back to OUTDIR_DEFAULT (below)
and no file is moved. (--out is accepted but ignored in this mode.)

SUPPORTED PDF TEMPLATES (auto-detected by report-type signature, never by vendor name)
  • QuickBooks Statement                 - vendor-issued statement with "INV #<num>. Due <date>" lines
  • QuickBooks Customer Open Balance     - QBO Customer Open Balance report, columnar
  • QuickBooks Open Invoices             - QBO Open Invoices report, columnar (4-col)
  • Plus per-vendor layouts (White Cap, Bobcat, Bodin, BURNCO, Cintas, Cow Town,
    Sunbelt, Croell) and generic tabular/columnar - see TEMPLATE_LABELS.

USAGE
  python3 statement_reconciler.py /path/to/statement.pdf            # inbox-style: Excel → Reconciliations, source → DONE if in inbox
  python3 statement_reconciler.py /path/to/statement.pdf --vendor "Exact QBO Display Name"
  python3 statement_reconciler.py /path/to/statement.pdf --dry-run  # reconcile + print, write/move nothing
  python3 statement_reconciler.py /path/to/statement.pdf --yes      # skip prompts
  python3 statement_reconciler.py --inbox                           # sweep the whole Statement Inbox

INTERACTIVE FLOW
  Two Y/N prompts before any QBO call so a misread PDF never wastes API
  roundtrips:
    1. After parse → shows vendor / date / total / first+last line. Confirm.
    2. After QBO vendor lookup → shows matched vendor name. Confirm.
  Ends with an INBOX SUMMARY (reconciled / moved-to-DONE / left-for-a-human) and
  a clickable link to the Reconciliations folder. Does not auto-open the Excel.

DEPENDENCIES
  bash python-env/setup.sh
"""
from __future__ import annotations

import argparse
import base64
import contextlib
import io
import datetime as dt
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import urllib.parse
from dataclasses import dataclass, field
from pathlib import Path
from collections import Counter
from typing import Dict, Iterator, List, Optional, Tuple

# Allow import of qbo_vault from project root (same pattern as qbo_bill_tracker)
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

try:
    import requests
except ImportError:
    print("✗ bash python-env/setup.sh")
    sys.exit(1)

try:
    import pdfplumber
except ImportError:
    print("✗ bash python-env/setup.sh")
    sys.exit(1)

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    print("✗ bash python-env/setup.sh")
    sys.exit(1)

from shared import qbo_vault as kc
from shared import qbo_api  # noqa: E402  (the one QuickBooks login)
from shared import paths
from shared.xlsx_guard import csv_cell, put as xl_text  # noqa: E402  (outside text never runs in Excel)
from shared import xlsx_verify
import statement_set as ss  # noqa: E402  (tool-local: one vendor = one running record)
import statement_markup as sm  # noqa: E402  (tool-local: the marked-up copy of the statement)
from shared import statement_record as sr  # noqa: E402  (the record the ledger's Vendor Center reads)

# ───────────────────────── constants ─────────────────────────

API_BASE = "https://quickbooks.api.intuit.com"
MINOR_VERSION = "70"
TX_SALES_TAX = 0.0825
TAX_TOLERANCE = 1.00  # dollars - accept within $1 of exact 8.25% match
EXACT_TOLERANCE = 0.01  # dollars - bill amount equality
LAG_LOOKBACK_DAYS = 60  # FLOOR for paid-bill lookback (days before stmt date)
LAG_LOOKBACK_STMT_PADDING_DAYS = 7  # extra padding before the statement's oldest line
# Effective cutoff = min(oldest_stmt_line_date - PADDING, stmt_date - LAG_LOOKBACK_DAYS)
# So statements with old bills extend the lookback as needed, but every run
# still pulls AT LEAST 60 days of paid history.
# Bill approval: the ONE rule lives in shared/bill_approval (shared with the Bill
# Tracker, 10/08/2026) - paper-era NOT APPROVED memo tag, or, for a bill entered
# since 09/16/2026 and still unpaid, "check QBO" (its approval queue is invisible
# to the API).
from shared import bill_approval  # noqa: E402
APPROVAL_WORKFLOW_START = bill_approval.APPROVAL_WORKFLOW_START.isoformat()


def _is_approved(memo: str) -> bool:
    """False only when the memo starts with NOT APPROVED (the paper-era tag)."""
    return bill_approval.memo_approved(memo)


def _approval(r: "ReconRow", by_id: Dict[str, "QboBill"]) -> str:
    """bill_approval state of a statement row's QBO bill: not approved / check QBO /
    approved."""
    b = by_id.get(r.qbo_bill_id)
    return bill_approval.approval_state(r.qbo_memo, b.created if b else "", b.open_balance if b else 0)

OUTDIR_DEFAULT = paths.get_path(
    "ACB_RECON_OUT_DIR",
    paths.onedrive_base() / "Automations-/statement reconciles",
)
ALIAS_FILE = Path(__file__).resolve().parent / "vendor_aliases.json"
CLERK_PERF_CSV = OUTDIR_DEFAULT / "clerk_performance.csv"

# QBO deep-link to open a Bill in the browser. Same format used by the sibling
# bill tracker (excel_bill_sync.py). Rendered as a ↗ hyperlink in the report.
QBO_BILL_URL_TEMPLATE = "https://qbo.intuit.com/app/bill?txnId={bill_id}"

# ── Inbox automation (Synology) ────────────────────────────────
# Folder-driven workflow root on the Accounting share (must be mounted). The
# Inbox is a pure dump: the sweep reads each statement, then FILES both the Excel
# and the source statement together under <root>/<Vendor>/<MM-YYYY>/. A statement
# whose reconciliation the clerk has marked DONE is left untouched on re-runs.
# (Moved 2026-09-16 from Automations/ into the new Accounts Payable/ folder.)
INBOX_ROOT = sr.root()      # ONE place names the root (shared/statement_record): the ledger reads the same folders
INBOX_SUPPORTED_EXTS = {".pdf", ".png", ".jpg", ".jpeg", ".heic", ".heif", ".xlsx", ".xls"}
STATEMENT_EMBED_MAX_PAGES = 20   # cap embedded statement pages to keep xlsx size sane
STATEMENT_EMBED_MAX_WIDTH = 900  # px - target on-sheet width per embedded page

# ── Template: QuickBooks Statement (vendor-issued) ─────────────
# Layout: a leading date + a transaction-type marker anchor every row.
#   Invoice: 02/25/2026 INV #589898. Due 03/15/2026 ...  80,620.00  84,770.31
#   Payment: 10/16/2025 PMT #23225.                        -38.49    -38.49
# PAYMENTS and CREDIT MEMOS (PMT/CM/…) have NO due-date column and a NEGATIVE
# amount (leading "-" or in parentheses).
QBO_STATEMENT_SIG = re.compile(r"INV\s*\#\d+\.\s*Due\s+\d", re.I)
# Match EVERY transaction row, not just invoices. Summing only INV rows drops
# credits/payments, so the line total comes out too HIGH by the credit and the
# tie-out to "Amount Due" falsely fails. This mirrors the 2026-08-12 fix to the
# Customer Open Balance parser (QBO_CUSTOMER_OPEN_BAL_LINE_RE) - the twin
# QBO-Statement parser never received it, so a Bee Line statement carrying a
# -38.49 payment tripped it (2026-09-10). The type token is generic so any QBO
# abbreviation (INV/PMT/CM/FC/…) is caught; amount and balance accept a sign or
# parentheses, and the due-date is folded into the optional "..." middle.
STMT_LINE_RE = re.compile(
    r"""
    ^\s*
    (?P<date>\d{1,2}/\d{1,2}/\d{2,4})           # 02/25/2026
    \s+
    (?P<type>[A-Z]{2,8})                        # INV / PMT / CM / FC ...
    \s*\#?(?P<ref>[A-Z]?\d+[A-Z]?)\.?             # #589898. / #23225. / #R16738. (revised)
    .*?                                          # ... Due <date> / description ...
    (?P<amount>\(?-?[\d,]+\.\d{2}\)?)           # 80,620.00 / -38.49 / (38.49)
    \s+
    (?P<balance>\(?-?[\d,]+\.\d{2}\)?)\s*$      # running balance (may be negative)
    """,
    re.VERBOSE | re.MULTILINE,
)

PO_RE = re.compile(r"PO\s*\#?\s*([^.]+?)\.", re.IGNORECASE)
ORIG_AMT_RE = re.compile(r"Orig\.\s*Amount\s*\$?([\d,]+\.\d{2})", re.IGNORECASE)
# Tried in order - first one with a hit wins. Most-specific label first.
STMT_DATE_PATTERNS = [
    # "Statement Date" / "Stmt Date" / "Statement Dt" label, then up to 40 chars to the date
    re.compile(r"(?:Statement|Stmt|Stat)\s*Da?te?[:\s][\s\S]{0,40}?(\d{1,2}/\d{1,2}/\d{2,4})", re.I),
    # Bare "Date" label, then up to 40 chars (vendor statement header noise between label and value)
    re.compile(r"\bDate\b[:\s][\s\S]{0,40}?(\d{1,2}/\d{1,2}/\d{2,4})", re.I),
    # Fallback: first standalone date on its own line - only used if labels missing
    re.compile(r"(?:^|\n)\s*(\d{1,2}/\d{1,2}/\d{4})\s*(?:\n|$)"),
]
# Amount Due: allow up to 80 chars (incl. newlines + "Amount Enc." label) between
# the "Amount Due" text and the dollar amount.
PMT_NO_REF_RE = re.compile(
    r"^\s*(\d{1,2}/\d{1,2}/\d{2,4})\s+(PMT|Discount)\s+(\(?-[\d,]+\.\d{2}\)?)\s+\(?-?[\d,]+\.\d{2}\)?\s*$",
    re.MULTILINE)
AMT_DUE_RE = re.compile(r"Amount\s+Due[\s\S]{0,80}?\$\s*([\d,]+\.\d{2})", re.IGNORECASE)

# ── Template: QuickBooks Customer Open Balance ─────────────────
# Vendor runs a "Customer Open Balance" report for Proficient Concrete from
# their own QBO. Layout (extracted text order):
#   <Vendor Name>                       ← vendor on its own line at top
#   9:35 AM                              ← timestamp
#   Customer Open Balance                ← report-title signature
#   05/15/26                             ← print date
#   Accrual Basis As of March 31, 2026   ← statement date
#   Type Date Num Memo Due Date Open Balance
#   PROFICIENT CONCRETE, LLC.            ← bill-to header
#   <SUB-CUSTOMER NAME>                  ← job
#   Invoice 3/18/2026 K733367 8811 OLD DECATUR RD 4/17/2026 269.95
#   Total <SUB-CUSTOMER>  269.95
#   ...
#   TOTAL 265,372.99
# Detection matches two title variants:
#   • "Customer Open Balance" (older QBO label)
#   • Header row "Type Date Num Memo Due Date Open Balance" (used by newer
#     QBO Statement reports titled "May 2026 Statement" or similar)
QBO_CUSTOMER_OPEN_BAL_SIG = re.compile(
    r"Customer\s+Open\s+Balance"
    r"|Type\s+Date\s+Num\s+Memo\s+Due\s+Date\s+Open\s+Balance",
    re.I,
)
QBO_AS_OF_RE = re.compile(
    r"As\s+of\s+([A-Za-z]+\s+\d{1,2},?\s+\d{4})", re.I)
# Matches EVERY transaction row in a Customer Open Balance report, not just
# invoices. Credit Memos and Payments carry NEGATIVE amounts (leading "-" or
# parentheses) and Payments have no due-date column - both must be parsed so
# the line-sum ties to the report's grand total. (Before 2026-08-12 only
# "Invoice" rows matched, so credits/payments were dropped, the line-sum came
# out too HIGH by the credit total, and every statement carrying a credit
# falsely failed the tie-out.) The due-date group is optional; the amount group
# accepts a sign or parentheses.
QBO_CUSTOMER_OPEN_BAL_LINE_RE = re.compile(
    r"""^\s*
    (?P<type>(?:Invoice|Credit(?:\s+M\w*)?|Payment|Discount|Journal|
                 Deposit|Sales\s+Receipt|Check|Bill\s+Pmt|Transfer)
             (?:\s*\.\.\.)?)\s+        # QBO truncates a narrow Type cell to "...":
                                       #   "Credit ...", "Credit Me..." and "Credit M..." all seen
    (?P<date>\d{1,2}/\d{1,2}/\d{2,4})\s+
    (?P<num>\S+)\s+
    (?P<memo>.+?)
    (?:\s+(?P<due>\d{1,2}/\d{1,2}/\d{2,4}))?\s+
    (?P<amount>\(?-?[\d,]+\.\d{2}\)?)\s*$""",
    re.VERBOSE | re.MULTILINE,
)
# Some CoB reports print the grand total across TWO columns ("Open Balance" +
# "Amount"), e.g. "TOTAL 383,940.69 383,940.69" - allow the extra trailing
# amount(s) or the total goes undetected and Amount Due reads $0 (Estrada CoB).
QBO_GRAND_TOTAL_RE = re.compile(
    r"^\s*TOTAL\s+([\d,]+\.\d{2})(?:\s+\(?-?[\d,]+\.\d{2}\)?)*\s*$", re.MULTILINE)
# Allow asterisks/lowercase in body so QBO's "**EXEMPT" / "*EXEMPT" markers
# and " - Other" suffix on the parent customer header still register as
# sub-customers. (Example: "Dallas Area Habitat for Humanity **EXEMPT".)
QBO_SUBCUST_RE = re.compile(
    r"^\s*(?!Total\b|TOTAL\b|Invoice\b|Credit\b|Type\b|Accrual\b|Cash\b|Page\b)"
    r"([A-Z][A-Za-z0-9\s\(\)\.,&'\-/\*]{3,})\s*$",
    re.MULTILINE,
)

# ── Template: QuickBooks Open Invoices ─────────────────────────
# Vendor runs an "Open Invoices" report for Proficient Concrete from their
# own QBO. Layout (extracted text order - vendor appears at bottom because
# pdfplumber reads the header columns first):
#   Date Num Due Date Open Balance            ← column header
#   Proficient Concrete                        ← bill-to header
#   <SUB-CUSTOMER NAME> [**EXEMPT marker]      ← job
#   09/26/2025 D80025 10/26/2025 372.96       ← invoice line (4 cols)
#   Total <SUB-CUSTOMER> 372.96
#   ...
#   Total Proficient Concrete 90,507.30
#   TOTAL 90,507.30
#   1:27 PM <Vendor Name>                      ← timestamp + vendor SAME line
#   05/21/26 Open Invoices                     ← print-date + title SAME line
#   As of May 21, 2026                          ← statement date
#   Page 1
QBO_OPEN_INVOICES_SIG = re.compile(r"\bOpen\s+Invoices\b", re.I)
QBO_OPEN_INV_LINE_RE = re.compile(
    r"""^\s*
    (?P<date>\d{1,2}/\d{1,2}/\d{2,4})\s+
    (?P<num>\S+)\s+
    (?P<due>\d{1,2}/\d{1,2}/\d{2,4})\s+
    (?P<amount>[\d,]+\.\d{2})\s*$""",
    re.VERBOSE | re.MULTILINE,
)
# Timestamp pattern shared by both QBO report templates - used by the
# generic vendor finder.
QBO_TIMESTAMP_RE = re.compile(
    r"^\s*(?P<time>\d{1,2}:\d{2}\s*(?:AM|PM))\s*(?P<rest>.*)$", re.I)

# ── Template: Vendor Statement (tabular, CMC-style) ────────────
# Vendors using line-based statement layouts with the header:
#   "Date Invoice Due Date Amount Pymt Date Payment Amount Tp Balance"
# Each data row: <date> <inv#> <due-date> <amount> [optional payment fields] <balance>
# Sub-customer groupings ("SHIP TO XYZ") and aging rows are skipped.
VENDOR_STMT_TABULAR_SIG = re.compile(
    r"Date\s+Invoice\s+Due\s+Date\s+Amount\b.*?\bBalance\b",
    re.I | re.DOTALL,
)
VENDOR_STMT_TABULAR_LINE_RE = re.compile(
    r"""^\s*
    (?P<date>\d{1,2}/\d{1,2}/\d{2,4})\s+
    (?P<ref>[A-Za-z0-9]+)\s+
    (?P<due>\d{1,2}/\d{1,2}/\d{2,4})\s+
    (?P<amount>[\d,]+\.\d{2})
    .*?
    (?P<balance>[\d,]+\.\d{2})\s*$
    """,
    re.VERBOSE | re.MULTILINE,
)
# CMC sub-customer header lines (skip these)
VENDOR_STMT_SHIPTO_RE = re.compile(r"^\s*SHIP[- ]TO\b", re.I)

# ── Template: Vendor Statement (columnar - Preferred Materials / Sunrise) ──
# These statements have a visible table grid in the PDF, but pdfplumber's
# default text extraction returns each COLUMN as a vertical stack of values
# (Date column, Description column, Charge column, Balance column each as
# their own list of lines, NOT one line per row). Parsing this requires
# pdfplumber.extract_words() with x/y coordinates to reconstruct rows.
# Detection: title "Statement" + the column-header tokens "Description" and
# "Charge" appearing in the extracted text without the CMC-style header.
VENDOR_STMT_COLUMNAR_SIG = re.compile(
    r"\bStatement\b.*?\bDescription\b.*?\bCharge\b.*?\bBalance\b",
    re.I | re.DOTALL,
)

# ── Template: Vendor Statement (White Cap / Billtrust-style) ───
# White Cap-issued statements via Billtrust. Layout: Transaction Date,
# Transaction No., T (type code), Original Transaction, Balance Due.
# Type codes: I=Invoice, C=Credit Memo, R=Rental, D=Debit Memo,
# U=Unapplied Payment, *=In Review. pdfplumber.extract_tables() returns
# all rows of the data table as ONE cell of newline-separated lines -
# easy to split and parse line-by-line.
VENDOR_STMT_WHITECAP_SIG = re.compile(
    r"(?=[\s\S]*?White\s+Cap)"
    r"(?=[\s\S]*?CLOSING\s+DATE)"
    r"(?=[\s\S]*?BALANCE\s+DUE)",
    re.I,
)
# Line format in pdfplumber's text extraction:
#   <date> <ref> <type> <original> <balance>  [<ref-dup> <po> <balance-dup>]
# The right-side remittance copy gets merged onto the same line - we capture
# only the first 5 fields and ignore the rest.
WHITECAP_ROW_RE = re.compile(
    r"""^\s*
    (?P<date>\d{1,2}/\d{1,2}/\d{2,4})\s+
    (?P<ref>\d+)\s+
    (?P<tp>[A-Z]\*?)\s+
    (?P<orig>-?[\d,]+\.\d{2})\s+
    (?P<balance>-?[\d,]+\.\d{2})
    (?:\s+.*)?$    # optional trailing fields (right-side remittance copy)
    """,
    re.VERBOSE | re.MULTILINE,
)

# Strings that should NEVER be treated as a vendor name when extracted.
_VENDOR_NOISE_RE = re.compile(
    r"^(?:Page\s+\d+|Open\s+Invoices|Customer\s+Open\s+Balance|Accrual\s+Basis|Cash\s+Basis|"
    r"As\s+of\b|Date\b|Type\b|Total\b|TOTAL\b|Invoice\b|Credit\b|"
    r"Statement\s*$|Statement\s+Date\b|Bill\s+To\b|Amount\s+Due\b|"
    r"Currency\b|Subsidiary\b|Company\s*$|Description\b|Charge\b|Balance\s+Forward\b)",
    re.I,
)

# Styles
HEADER_FILL   = PatternFill("solid", start_color="1F4E78")
HEADER_FONT   = Font(bold=True, color="FFFFFF", name="Arial", size=11)
TITLE_FONT    = Font(bold=True, name="Arial", size=14, color="1F4E78")
LABEL_FONT    = Font(bold=True, name="Arial", size=11)
BODY_FONT     = Font(name="Arial", size=10)
SUBTOTAL_FILL = PatternFill("solid", start_color="E7E6E6")
OK_FILL       = PatternFill("solid", start_color="C6EFCE")
WARN_FILL     = PatternFill("solid", start_color="FFEB9C")
BAD_FILL      = PatternFill("solid", start_color="FFC7CE")
THIN          = Side(border_style="thin", color="BFBFBF")
BORDER        = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER        = Alignment(horizontal="center", vertical="center")
LEFT          = Alignment(horizontal="left",   vertical="center", wrap_text=True)
RIGHT         = Alignment(horizontal="right",  vertical="center")
MONEY         = '"$"#,##0.00;[Red]("$"#,##0.00);"-"'

# ───────────────────────── terminal UI (stdlib only) ─────────────────────────

class _Term:
    """ANSI helpers. Auto-disables on non-tty or when NO_COLOR/--no-color is set."""
    enabled: bool = sys.stdout.isatty() and os.environ.get("NO_COLOR") is None
    G    = "\033[32m"  # green
    Y    = "\033[33m"  # yellow
    R    = "\033[31m"  # red
    B    = "\033[34m"  # blue
    C    = "\033[36m"  # cyan
    DIM  = "\033[2m"
    BOLD = "\033[1m"
    RESET    = "\033[0m"
    CLEARLINE = "\033[2K"  # clear entire line

    @classmethod
    def disable(cls) -> None:
        cls.enabled = False

    @classmethod
    def color(cls, code: str, text: str) -> str:
        if not cls.enabled:
            return text
        return f"{code}{text}{cls.RESET}"


def _width() -> int:
    return shutil.get_terminal_size((80, 24)).columns


def _phase(label: str) -> float:
    """Print a phase header → returns start time so the caller can stamp duration."""
    print(_Term.color(_Term.C, f"→ {label}"))
    return time.time()


def _done(t0: float, msg: str, marker: str = "✓", color: str = _Term.G) -> None:
    elapsed = time.time() - t0
    print(f"  {_Term.color(color, marker)} {msg}  {_Term.color(_Term.DIM, f'({elapsed*1000:.0f} ms)')}")


def _warn(msg: str) -> None:
    print(f"  {_Term.color(_Term.Y, '⚠')} {msg}")


def _fail(msg: str) -> None:
    print(f"  {_Term.color(_Term.R, '✗')} {msg}")


def _bar(current: int, total: int, suffix: str = "", width_target: int = 24) -> None:
    """Inline progress bar: redraws in place. No newline until caller adds one."""
    if total <= 0:
        return
    pct = max(0.0, min(1.0, current / total))
    filled = int(width_target * pct)
    bar = "█" * filled + "░" * (width_target - filled)
    line = f"\r  [{bar}] {current}/{total}  {suffix}"
    # Trim to terminal width so no wrap on tiny windows
    line = line[: max(20, _width() - 1)]
    sys.stdout.write(_Term.CLEARLINE + line)
    sys.stdout.flush()


def _bar_end() -> None:
    """Move to a fresh line after a progress bar finishes."""
    sys.stdout.write("\n")
    sys.stdout.flush()


# Category → display label + color for the live stream
_CAT_STYLE: Dict[str, Tuple[str, str]] = {
    "MATCHED":               ("✓ Match    ", _Term.G),
    "VENDOR_TAX_VIOLATION":  ("⚠ Tax viol ", _Term.Y),
    "CLERK_AMOUNT_MISMATCH": ("⚠ Mismatch ", _Term.Y),
    "LIKELY_VENDOR_LAG":     ("⊙ Paid lag ", _Term.B),
    "MISSING_IN_QBO":        ("✗ Not in QB", _Term.R),
    "MISSING_ON_STATEMENT":  ("✗ Not on St", _Term.R),
}


# ───────────────────────── data classes ─────────────────────────

@dataclass
class StmtLine:
    date: str           # YYYY-MM-DD
    ref: str            # invoice number
    amount: float       # line amount
    po: str = ""        # PO number from description
    address: str = ""   # job address text

@dataclass
class QboBill:
    bill_id: str
    doc_number: str
    txn_date: str
    open_balance: float
    total_amount: float
    memo: str = ""        # Bill memo from QBO. The QBO UI "Memo" field maps to
                          # the API `PrivateNote` field (verified via project
                          # script read_private_note.py). We also accept the
                          # rarer top-level `Memo` field as fallback. Contains
                          # "Not Approved" when AP/PM hasn't signed off.
    created: str = ""     # MetaData.CreateTime date (YYYY-MM-DD): a bill entered on/after
                          # APPROVAL_WORKFLOW_START and still open may be pending approval
                          # in QBO's workflow, which the API cannot see.

    @property
    def is_approved(self) -> bool:
        return _is_approved(self.memo)

@dataclass
class ReconRow:
    category: str       # MATCHED | VENDOR_TAX_VIOLATION | CLERK_AMOUNT_MISMATCH | LIKELY_VENDOR_LAG | MISSING_IN_QBO | MISSING_ON_STATEMENT
    # Note: `date` and `ref` are kept for backward-compatibility with action-list,
    # CSV, and detail-sheet code. For MISSING_ON_STATEMENT they hold the QBO data.
    # The summary sheet uses the more explicit stmt_*/qbo_* fields below.
    date: str           # canonical date (stmt date if available, else QBO date)
    ref: str            # canonical ref (stmt ref if available, else QBO ref)
    stmt_amount: float
    qbo_amount: float
    po: str
    address: str
    notes: str = ""
    # Side-by-side fields for the new summary layout
    stmt_ref: str = ""
    qbo_ref: str = ""
    stmt_date: str = ""
    qbo_date: str = ""
    qbo_memo: str = ""    # bill-level Memo from QBO; "" when there is no QBO bill (MISSING_IN_QBO)
    qbo_bill_id: str = "" # QBO internal Bill Id, used to build the ↗ deep-link; "" when no QBO bill

# ───────────────────────── PDF parsing ─────────────────────────

def _pdf_text(pdf_path: Path) -> str:
    parts = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            parts.append(page.extract_text() or "")
    text = "\n".join(parts)
    if len(text.strip()) < 40:
        # A scanned statement (a picture in a PDF - Power Jack, Bobcat 06/2026) has no
        # text: read it with OCR. It still has to tie out to be used (10/08/2026).
        try:
            text = _scanned_pdf_text(pdf_path) or text
        except Exception:
            pass
    return text


def _scanned_pdf_text(pdf_path: Path, max_pages: int = 10) -> str:
    import pypdfium2 as pdfium
    import pytesseract
    pdf = pdfium.PdfDocument(str(pdf_path))
    try:
        pages = []
        for i in range(min(len(pdf), max_pages)):
            img = pdf[i].render(scale=300 / 72).to_pil().convert("L")
            pages.append(pytesseract.image_to_string(img, config="--psm 6"))
        # OCR reads table rules as '|' between columns ('8,404.00 | 8,404.00').
        return re.sub(r"\s*\|\s*", " ", "\n".join(pages))
    finally:
        pdf.close()


def _image_to_text(img_path: Path) -> str:
    """OCR an image (PNG/JPG/etc.) to plain text using Tesseract.
    Lazy-imports pytesseract + Pillow so installs aren't required for users
    who never pass image statements. Prints actionable install instructions
    on failure rather than a stack trace."""
    try:
        import pytesseract
        from PIL import Image
    except ImportError as e:
        sys.exit(_Term.color(_Term.R,
            f"✗ Image OCR needs Tesseract + Python wrappers. Install once:\n"
            f"    brew install tesseract\n"
            f"    bash python-env/setup.sh\n"
            f"  (missing: {e.name})"))
    # Verify the tesseract binary itself is reachable
    try:
        pytesseract.get_tesseract_version()
    except Exception:
        sys.exit(_Term.color(_Term.R,
            "✗ Tesseract binary not found on PATH. Install once:\n"
            "    brew install tesseract"))
    # HEIC support (iPhone screenshots) needs the optional pillow-heif backend
    if img_path.suffix.lower() in (".heic", ".heif"):
        try:
            from pillow_heif import register_heif_opener
            register_heif_opener()
        except ImportError:
            sys.exit(_Term.color(_Term.R,
                "✗ HEIC images need the pillow-heif backend. Install once:\n"
                "    bash python-env/setup.sh"))
    try:
        img = Image.open(img_path)
        # --psm 6: assume a single uniform block of text. Critical for
        # statement tables - keeps each invoice row as a single line of OCR
        # output (default mode breaks columns into separate vertical stacks
        # which destroys the date-ref-amount alignment our parsers need).
        # preserve_interword_spaces=1 keeps wide-spaced columns readable.
        return pytesseract.image_to_string(
            img,
            config="--psm 6 -c preserve_interword_spaces=1"
        )
    except Exception as e:
        sys.exit(_Term.color(_Term.R, f"✗ OCR failed on {img_path.name}: {e}"))


TEMPLATE_LABELS = {
    "qbo_statement":             "QuickBooks Statement (PDF)",
    "qbo_customer_open_balance": "QuickBooks Customer Open Balance / Statement (PDF - Type Date Num Memo Due Date Open Balance header)",
    "qbo_open_invoices":         "QuickBooks Open Invoices (PDF)",
    "vendor_stmt_tabular":       "Vendor Statement, tabular (PDF - Date Invoice Due Date Amount ... Balance header, e.g. CMC)",
    "vendor_stmt_columnar":      "Vendor Statement, columnar (PDF - Date Description Charge Payment Balance grid, e.g. Preferred Materials / Sunrise)",
    "vendor_stmt_whitecap":      "Vendor Statement, White Cap / Billtrust (PDF - Transaction Date | Transaction No. | T | Original | Balance Due with I/C/R/D/U type codes)",
    "excel_columnar":            "Excel statement (.xlsx or .xls - columns: Inv Date | Invoice # | Original Invoice Amount | Balance | Due Date)",
    "vendor_bobcat":             "Vendor Statement, Bobcat/equipment (PDF - Invoice Number | Invoice Date | Due Date | Purchase Order | Balance)",
    "vendor_bodin":              "Vendor Statement, Bodin/concrete (PDF - Invoice | Date | Type | Reference | Yardage | Credit/Debit | Balance)",
    "vendor_burnco":             "Vendor Statement, BURNCO (PDF - Date | Number | Delivery Address | PO | Type | Original | Balance Due)",
    "vendor_cintas":             "Vendor Statement, Cintas (PDF - Date | Sold-To | Reference | Amount Due | Due Date)",
    "vendor_cowtown":            "Vendor Statement, past-due letter (PDF - Job | Inv. No. | Inv. Date | Due Date | Inv. Amount | Balance)",
    "vendor_sunbelt":            "Vendor Statement, Sunbelt Rentals (PDF - Date | Invoice | Job Description | Amount Due)",
    "vendor_croell":             "Vendor Statement, Croell Inc (PDF - Date | Cd | Invoice | Description | Amount | Balance doubled register/remittance layout)",
    "vendor_abatix":             "Vendor Statement, Abatix Corp (PDF - Invoice Date | Due Date | Invoice No | PO | Amount Due | Enclosed No)",
    "vendor_abatix_online":      "Vendor Statement, Abatix online (PDF - Invoice Number | Date | Due Date | PO | Original | Amount Due + 'Summary of Invoice Age')",
    "vendor_croell_ar":          "Vendor Statement, Croell AR Customer Balance (PDF - Invoice # | PO# | Invoice Date | Disc Date | Type | ... | Current Due Less Discounts)",
    "vendor_quikrete":           "Vendor Statement, Quikrete Ready Mix account statement (PDF - Invoice No. | Invoice Date | Due Date | Remark | Cust P.O. | Invoice Amount | Balance Due)",
    "qbo_invoice_list":          "QuickBooks 'Invoices for <customer>' list (PDF - Num | Date | Name | Amount | Open Balance)",
    "vendor_voidform":           "Vendor Statement, VoidForm (PDF - Date | Due Date | Doc. Type | Ref. Nbr. | Ext. Ref. Nbr. | Orig. Amount | Amount Due | Balance)",
}


def detect_template(text: str) -> str:
    """Map a PDF's extracted text to one of the supported template keys.
    Templates are named by REPORT TYPE, never by vendor. Returns "" if no
    supported template matches."""
    # Order matters: most-specific signature first.
    if ABATIX_ONLINE_SIG.search(text):
        return "vendor_abatix_online"
    if CROELL_AR_SIG.search(text):
        return "vendor_croell_ar"
    if QUIKRETE_SIG.search(text):
        return "vendor_quikrete"
    if QBO_INVOICE_LIST_SIG.search(text):
        return "qbo_invoice_list"
    if VOIDFORM_SIG.search(text):
        return "vendor_voidform"
    # QBO Statement: "INV #N. Due N" line anchor
    if QBO_STATEMENT_SIG.search(text):
        return "qbo_statement"
    # QBO Customer Open Balance / new "Statement" variant
    if QBO_CUSTOMER_OPEN_BAL_SIG.search(text):
        return "qbo_customer_open_balance"
    # QBO Open Invoices
    if QBO_OPEN_INVOICES_SIG.search(text):
        return "qbo_open_invoices"
    # White Cap / Billtrust - check BEFORE the more-generic columnar/tabular
    # signatures since they could partially match a White Cap layout
    if VENDOR_STMT_WHITECAP_SIG.search(text):
        return "vendor_stmt_whitecap"
    # Specific vendor-statement layouts (2026-07-01 batch) - checked before the
    # generic tabular/columnar signatures so their distinct headers win.
    if BOBCAT_SIG.search(text):
        return "vendor_bobcat"
    if BODIN_SIG.search(text):
        return "vendor_bodin"
    if BURNCO_SIG.search(text):
        return "vendor_burnco"
    if CINTAS_SIG.search(text):
        return "vendor_cintas"
    if COWTOWN_SIG.search(text):
        return "vendor_cowtown"
    if SUNBELT_SIG.search(text):
        return "vendor_sunbelt"
    # Croell Inc - doubled register/remittance header; must beat the generic
    # tabular/columnar sigs (its "Finance Charge" wording trips columnar).
    if CROELL_SIG.search(text):
        return "vendor_croell"
    if ABATIX_SIG.search(text):
        return "vendor_abatix"
    # Vendor tabular statement (Date Invoice Due Date Amount ... Balance header)
    if VENDOR_STMT_TABULAR_SIG.search(text):
        return "vendor_stmt_tabular"
    # Vendor columnar statement (Date / Description / Charge / Balance grid)
    if VENDOR_STMT_COLUMNAR_SIG.search(text):
        return "vendor_stmt_columnar"
    return ""


def _is_vendor_noise(s: str) -> bool:
    """True when a line is clearly NOT a vendor name (date, time, page, title,
    header row, our own company)."""
    if not s:
        return True
    if re.match(r"^\d{1,2}[/-]\d{1,2}([/-]\d{2,4})?\s*$", s):
        return True   # pure date
    if re.match(r"^\d{1,2}:\d{2}\s*(?:AM|PM)\s*$", s, re.I):
        return True   # pure time
    if _VENDOR_NOISE_RE.match(s):
        return True   # known title / header keyword
    if "PROFICIENT" in s.upper():
        return True   # our own company - never the vendor we're reconciling against
    return False


def _find_qbo_report_vendor(text: str) -> str:
    """Extract vendor name from a QBO report (Customer Open Balance, Open
    Invoices, or any future QBO report). The vendor is paired with a HH:MM
    AM/PM timestamp the report-runner's QBO prints - either on the SAME line
    after the timestamp, or on the line IMMEDIATELY BEFORE it. Handles both
    layouts observed in pdfplumber-extracted text order."""
    lines = text.splitlines()
    for i, line in enumerate(lines):
        m = QBO_TIMESTAMP_RE.match(line)
        if not m:
            continue
        # Case A: vendor inline AFTER timestamp on same line
        #   "1:27 PM Post-Tension Services of Texas"
        rest = m.group("rest").strip()
        if rest and not _is_vendor_noise(rest):
            return rest
        # Case B: vendor on the line immediately BEFORE the timestamp
        #   "Ready Cable, Inc\n9:35 AM\n..."
        if i > 0:
            prev = lines[i - 1].strip()
            if prev and not _is_vendor_noise(prev):
                return prev
    return ""


def _find_qbo_statement_vendor(text: str) -> str:
    """Extract vendor from a QBO Statement (vendor-issued). Vendor appears
    prominently as a header line containing a corporate suffix. Skip our
    own header."""
    for line in text.splitlines():
        if re.search(r"\b(LLC|L\.L\.C\.|INC|CO\.|CORP|COMPANY)\b", line, re.I):
            cand = line.strip()
            if not _is_vendor_noise(cand):
                return cand
    return ""


def _paren_amount(raw: str) -> float:
    """Parse a money token that may be negative via a leading '-' OR parentheses.
    Credits/payments on statements print either way, and both must net out.
        '(1,483.03)' -> -1483.03 · '-491.73' -> -491.73 · '75.00' -> 75.00"""
    s = raw.strip()
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()").replace(",", "").replace("$", "").strip()
    try:
        v = float(s)
    except ValueError:
        return 0.0
    return -v if neg else v


def parse_statement_qbo_customer_open_balance(full_text: str) -> Tuple[str, str, float, List[StmtLine]]:
    """Parse a QuickBooks Customer Open Balance report into the common shape."""
    vendor = _find_qbo_report_vendor(full_text)

    # Statement date - "As of <Month DD, YYYY>"
    stmt_date = ""
    m = QBO_AS_OF_RE.search(full_text)
    if m:
        raw = re.sub(r",", "", m.group(1))  # "March 31 2026"
        for fmt in ("%B %d %Y", "%b %d %Y"):
            try:
                stmt_date = dt.datetime.strptime(raw.strip(), fmt).strftime("%Y-%m-%d")
                break
            except ValueError:
                continue
    # Fallback: the newer Customer Open Balance layout prints the as-of date as a
    # bare MM/DD/YY on the line right under the "Customer Open Balance" title
    # (Estrada / Post-Tension), with no "As of" wording.
    if not stmt_date:
        m = re.search(r"Customer Open Balance\s*[\r\n]+\s*(\d{1,2}/\d{1,2}/\d{2,4})",
                      full_text, re.I)
        if m:
            stmt_date = _norm_date(m.group(1))

    # Grand total - "TOTAL <amount>" at end (not "Total <name> <amount>" - must be standalone TOTAL)
    amt_due = 0.0
    m = QBO_GRAND_TOTAL_RE.search(full_text)
    if m:
        try:
            amt_due = float(m.group(1).replace(",", ""))
        except ValueError:
            pass

    # Line items - walk lines, tracking the current sub-customer for PO/job tagging
    lines: List[StmtLine] = []
    current_subcust = ""
    text_lines = full_text.splitlines()
    for line in text_lines:
        stripped = line.strip()
        if not stripped:
            continue
        # Transaction row FIRST (Invoice / Credit Memo / Payment / ...). Checked
        # before the sub-customer test because a Payment row has no due-date
        # column and would otherwise satisfy the caps-leading header pattern,
        # getting swallowed as a bogus sub-customer instead of counted.
        inv_m = QBO_CUSTOMER_OPEN_BAL_LINE_RE.match(line)
        if inv_m:
            memo = inv_m.group("memo").strip()
            # Truncated memos: QBO shows "..." when memo is cut off
            memo = memo.rstrip(".").strip() if memo.endswith("...") else memo
            lines.append(StmtLine(
                date=_norm_date(inv_m.group("date")),
                ref=inv_m.group("num").strip(),
                amount=_paren_amount(inv_m.group("amount")),
                po=current_subcust, address=memo))
            continue
        # Sub-customer header? (caps-leading line, not a txn/total/header)
        sub_m = QBO_SUBCUST_RE.match(line)
        if sub_m and not stripped.startswith(
                ("Invoice", "Credit", "Payment", "Total", "TOTAL", "Type", "Accrual", "Cash")):
            cand = sub_m.group(1).strip().rstrip(".").strip()
            # Filter out the parent customer header (our own company).
            # Keep sub-customers with the parent name + suffix ("... - Other", etc.).
            if cand.upper().rstrip(".").rstrip(",").strip() in (
                "PROFICIENT CONCRETE, LLC", "PROFICIENT CONCRETE LLC", "PROFICIENT CONCRETE"
            ):
                continue
            current_subcust = cand
            continue
    return vendor, stmt_date, amt_due, lines


def parse_statement_qbo_open_invoices(full_text: str) -> Tuple[str, str, float, List[StmtLine]]:
    """Parse a QuickBooks Open Invoices report into the common shape.
    4-column layout: Date | Num | Due Date | Open Balance. Grouped by
    sub-customer under the bill-to header ("Proficient Concrete")."""
    vendor = _find_qbo_report_vendor(full_text)

    # Statement date - "As of <Month DD, YYYY>" (same convention as Customer Open Balance)
    stmt_date = ""
    m = QBO_AS_OF_RE.search(full_text)
    if m:
        raw = re.sub(r",", "", m.group(1))
        for fmt in ("%B %d %Y", "%b %d %Y"):
            try:
                stmt_date = dt.datetime.strptime(raw.strip(), fmt).strftime("%Y-%m-%d")
                break
            except ValueError:
                continue

    # Grand total
    amt_due = 0.0
    m = QBO_GRAND_TOTAL_RE.search(full_text)
    if m:
        try:
            amt_due = float(m.group(1).replace(",", ""))
        except ValueError:
            pass

    # Line items
    lines: List[StmtLine] = []
    current_subcust = ""
    for line in full_text.splitlines():
        stripped = line.strip()
        if not stripped:
            continue
        # Sub-customer header? (must NOT match an invoice-row pattern first)
        if QBO_OPEN_INV_LINE_RE.match(line):
            inv_m = QBO_OPEN_INV_LINE_RE.match(line)
            date_str = _norm_date(inv_m.group("date"))
            num = inv_m.group("num").strip()
            amount = float(inv_m.group("amount").replace(",", ""))
            lines.append(StmtLine(date=date_str, ref=num, amount=amount,
                                  po=current_subcust, address=""))
            continue
        sub_m = QBO_SUBCUST_RE.match(line)
        if sub_m and not stripped.startswith(("Invoice", "Total", "TOTAL", "Type", "Accrual", "Cash", "Page", "As ")):
            cand = sub_m.group(1).strip().rstrip(".").strip()
            # Skip the bill-to header (our own company, with or without LLC suffix)
            if cand.upper().rstrip(".").rstrip(",").strip() in (
                "PROFICIENT CONCRETE, LLC", "PROFICIENT CONCRETE LLC", "PROFICIENT CONCRETE"
            ):
                continue
            # Skip the column-header row
            if cand.upper().startswith("DATE NUM"):
                continue
            current_subcust = cand
            continue
    if not lines:
        sect = _qbo_open_invoices_sections(full_text)
        if sect is not None:
            lines, amt_due = sect
    return vendor, stmt_date, amt_due, lines


_OURS_RE = re.compile(r"\bPROFICIENT\b", re.I)


def _qbo_open_invoices_sections(full_text: str) -> Optional[Tuple[List[StmtLine], float]]:
    """The newer QBO Open Invoices layout (Date | Transaction type | Num | Term | Due
    date | Open balance), one section per CUSTOMER - Core Concrete Pumping sent its
    whole receivables report (10/07/2026). Only OUR section(s) count; the total is
    our 'Total for ...' line, never the report TOTAL. None = not this layout."""
    row = re.compile(r"^(?P<d>\d{1,2}/\d{1,2}/\d{4})\s+(?P<typ>Invoice|Credit Memo|Payment|"
                     r"Journal Entry|Sales Receipt|Refund|Bill)\s+(?P<num>\S+)\s+.*?"
                     r"(?P<amt>\(?-?\$?[\d,]+\.\d{2}\)?)$")
    sections: Dict[str, List[StmtLine]] = {}
    totals: Dict[str, float] = {}
    cur, expect_name = "", False
    for ln in (x.strip() for x in full_text.splitlines()):
        if not ln:
            continue
        if re.match(r"^Date\s+Transaction type", ln, re.I):
            expect_name = not cur          # the report's first column header; repeats per page
            continue
        m = row.match(ln)
        if m:
            if cur:
                sections.setdefault(cur, []).append(StmtLine(
                    date=_norm_date(m["d"]), ref=m["num"], amount=_paren_amount(m["amt"])))
            continue
        t = re.match(r"^Total for (?P<name>.+?)\s+(?P<amt>\(?-?\$?[\d,]+\.\d{2}\)?)$", ln)
        if t:
            totals[t["name"].strip()] = _paren_amount(t["amt"])
            expect_name = True             # the next name line opens the next customer
            continue
        # A customer name comes only right after a 'Total for' (or the first column
        # header) - page furniture (title, 'As of', the printed-on line) between
        # rows of one customer never splits it (Core: page 2 header mid-section).
        if expect_name and not ln.upper().startswith(("TOTAL", "AS OF")):
            cur, expect_name = ln, False
    if not sections:
        return None
    ours = [k for k in sections if _OURS_RE.search(k)]
    if len(sections) > 1:
        _warn(f"this report covers {len(sections)} customers - only "
              f"{', '.join(ours) or 'none of them is ours'} counted.")
    lines = [ln for k in ours for ln in sections[k]]
    _LAST_SKIPPED.extend(ln for k in sections if k not in ours for ln in sections[k])
    return lines, round(sum(totals.get(k, 0.0) for k in ours), 2)


# ── New vendor-statement templates (2026-07-01 batch) ─────────────────
# Each is keyed on a STRUCTURAL header signature, never the vendor name.
# All six were validated to tie their line-sum to the vendor's own stated
# total (Bobcat 19,395.99 · Bodin 169,039.44 · BURNCO 291,152.51 ·
# Cintas 389.08 · Cow Town 451,715.48 · Sunbelt 33,058.62).
BOBCAT_SIG  = re.compile(r"INVOICE NUMBER\s+INVOICE DATE\s+DUE DATE\s+PURCHASE ORDER\s+BALANCE", re.I)
BODIN_SIG   = re.compile(r"TYPE\s+REFERENCE\s+YARDAGE\s+CREDIT\s*/\s*DEBIT\s+BALANCE", re.I)
BURNCO_SIG  = re.compile(r"Delivery Address\s+PO Number\s+Type", re.I)
CINTAS_SIG  = re.compile(r"DATE\s+SOLD-TO\s+DESCRIPTION\s+REFERENCE\s+AMOUNT DUE\s+DUE DATE", re.I)
COWTOWN_SIG = re.compile(r"Inv\.\s*No\.\s+Inv\.\s*Date\s+Due Date", re.I)
SUNBELT_SIG = re.compile(r"DATE\s+INVOICE\s+JOB\s+DESCRIPTION\s+AMOUNT\s+DUE", re.I)
# Abatix Corp - columnar statement, distinctive doubled "Invoice Amount" header
# (Invoice Date | Due Date | Invoice No | PO | Amount Due | Enclosed #).
ABATIX_SIG = re.compile(r"Invoice\s+Due\s+Invoice\s+Amount\s+Invoice\s+Amount", re.I)
# Layouts first seen in the 10/07/2026 Inbox sweep.
ABATIX_ONLINE_SIG = re.compile(r"Summary of Invoice Age for", re.I)
CROELL_AR_SIG = re.compile(r"AR Customer Balance[\s\S]{0,600}?Balance After Disc", re.I)
QUIKRETE_SIG = re.compile(r"INVOICE NO\.\s+INVOICE DATE\s+DUE DATE\s+REMARK", re.I)
QBO_INVOICE_LIST_SIG = re.compile(r"Num\s+Date\s+Name\s+Amount\s+Open Balance", re.I)
VOIDFORM_SIG = re.compile(r"Doc\.\s*Type\s+Ref\.\s*Nbr\.\s+Ext\.\s*Ref\.\s*Nbr\.", re.I)

# ── Template: Croell Inc statement (added 2026-08-12) ─────────────────
# pdfplumber merges the left register and the right remittance stub onto one
# physical line, so each data row reads:
#   <date> <cd> <invoice> <description> <amount> <balance> <due date> \
#       <invoice-dup> <cd-dup> <amount-dup>
# We take the FIRST amount (the left "Amount" column) and ignore the duplicated
# remittance fields. Credits print in parentheses -> negative, so the line-sum
# nets to Balance Due. Cd codes seen: I=Invoice, F=Finance Charge.
# The doubled header is an unmistakable signature (checked before the generic
# tabular/columnar sigs, which the "Finance Charge" wording would otherwise trip).
CROELL_SIG = re.compile(
    r"Date\s+Cd\s+Invoice\s+Description\s+Amount\s+Balance\s+"
    r"Date\s+Due\s+Invoice\s+Cd\s+Amount", re.I)
CROELL_ROW_RE = re.compile(
    r"""^\s*
    (?P<date>\d{1,2}/\d{1,2}/\d{4})\s+
    (?P<cd>[A-Z])\s+
    (?P<num>\d+)\s+
    (?P<desc>.+?)\s+
    (?P<amount>\(?-?[\d,]+\.\d{2}\)?)\s+
    (?P<balance>\(?-?[\d,]+\.\d{2}\)?)\s+
    (?P<duedate>\d{1,2}/\d{1,2}/\d{4})\s+
    (?P<num2>\d+)\s+
    (?P<cd2>[A-Z])\s+
    (?P<amount2>\(?-?[\d,]+\.\d{2}\)?)\s*$""",
    re.VERBOSE | re.MULTILINE,
)


def _grab(pattern: str, text: str) -> str:
    m = re.search(pattern, text, re.I)
    return m.group(1) if m else ""


def _last_amt_after(label: str, text: str) -> float:
    """Last decimal amount on the block right after `label` (for aging-row totals)."""
    m = re.search(label + r"\s*[\r\n]+([$\d,. ]+)", text, re.I)
    if m:
        nums = re.findall(r"[\d,]+\.\d{2}", m.group(1))
        if nums:
            return float(nums[-1].replace(",", ""))
    return 0.0


def parse_statement_bobcat(full_text: str) -> Tuple[str, str, float, List[StmtLine]]:
    stmt_date = _norm_date(_grab(r"\bDATE\s+(\d{1,2}/\d{1,2}/\d{2})\b", full_text))
    amt_due = _last_amt_after("TOTAL DUE", full_text)
    rx = re.compile(r"^(?P<ref>[0-9A-Z]{5,})\s+(?P<d>\d{1,2}/\d{1,2}/\d{2})\s+\d{1,2}/\d{1,2}/\d{2}\s+(?P<mid>.*?)\s*(?P<amt>[\d,]+\.\d{2})\s*$")
    lines: List[StmtLine] = []
    for ln in full_text.splitlines():
        m = rx.match(ln.strip())
        if not m:
            continue
        po = re.sub(r"\bAPPROVED\b", "", m["mid"], flags=re.I).strip()
        lines.append(StmtLine(date=_norm_date(m["d"]), ref=m["ref"],
                              amount=float(m["amt"].replace(",", "")), po=po))
    if not amt_due:
        amt_due = round(sum(l.amount for l in lines), 2)
    return "", stmt_date, amt_due, lines


def parse_statement_bodin(full_text: str) -> Tuple[str, str, float, List[StmtLine]]:
    stmt_date = _norm_date(_grab(r"STATEMENT DATE[\s\S]{0,40}?(\d{1,2}/\d{1,2}/\d{4})", full_text))
    amt_due = 0.0
    m = re.search(r"THIS AMOUNT\s*[\r\n]+\s*([\d,]+\.\d{2})", full_text, re.I)
    if m:
        amt_due = float(m.group(1).replace(",", ""))
    rx = re.compile(r"^(?P<ref>\d+)\s+(?P<d>\d{1,2}/\d{1,2}/\d{4})\s+IN\s+(?P<mid>.*?)\s+[\d.]+\s*YDS\s+-?[\d,]+\.\d{2}\s+(?P<amt>-?[\d,]+\.\d{2})\s*$")
    lines: List[StmtLine] = []
    for ln in full_text.splitlines():
        m = rx.match(ln.strip())
        if not m:
            continue
        lines.append(StmtLine(date=_norm_date(m["d"]), ref=m["ref"],
                              amount=float(m["amt"].replace(",", "")), address=m["mid"].strip()))
    if not amt_due:
        amt_due = round(sum(l.amount for l in lines), 2)
    return "", stmt_date, amt_due, lines


def parse_statement_burnco(full_text: str) -> Tuple[str, str, float, List[StmtLine]]:
    stmt_date = _norm_date(_grab(r"STATEMENT DATE:\s*(\d{2}-\d{2}-\d{4})", full_text).replace("-", "/"))
    amt_due = 0.0
    m = re.search(r"BALANCE DUE:\s*([\d,]+\.\d{2})", full_text, re.I)
    if m:
        amt_due = float(m.group(1).replace(",", ""))
    rx = re.compile(r"^(?P<d>\d{2}-\d{2}-\d{4})\s+(?P<ref>SA\d+)\s+(?P<mid>.*?)\s+(?:Invoice|Credit)\s+-?[\d,]+\.\d{2}\s+(?P<amt>-?[\d,]+\.\d{2})\s*$", re.I)
    lines: List[StmtLine] = []
    for ln in full_text.splitlines():
        m = rx.match(ln.strip())
        if not m:
            continue
        lines.append(StmtLine(date=_norm_date(m["d"].replace("-", "/")), ref=m["ref"],
                              amount=float(m["amt"].replace(",", "")), address=m["mid"].strip()))
    if not amt_due:
        amt_due = round(sum(l.amount for l in lines), 2)
    return "", stmt_date, amt_due, lines


def parse_statement_cintas(full_text: str) -> Tuple[str, str, float, List[StmtLine]]:
    stmt_date = _norm_date(_grab(r"STATEMENT DATE\s+(\d{1,2}/\d{1,2}/\d{4})", full_text))
    amt_due = _last_amt_after("TOTAL DUE", full_text)
    rx = re.compile(r"^(?P<d>\d{1,2}/\d{1,2}/\d{4})\s+\d+\s+(?P<ref>\d+)\s+\$\s*(?P<amt>[\d,]+\.\d{2})\s+\d{1,2}/\d{1,2}/\d{4}\s*$")
    lines: List[StmtLine] = []
    for ln in full_text.splitlines():
        m = rx.match(ln.strip())
        if not m:
            continue
        lines.append(StmtLine(date=_norm_date(m["d"]), ref=m["ref"],
                              amount=float(m["amt"].replace(",", ""))))
    if not amt_due:
        amt_due = round(sum(l.amount for l in lines), 2)
    return "", stmt_date, amt_due, lines


def parse_statement_abatix(full_text: str) -> Tuple[str, str, float, List[StmtLine]]:
    """Abatix Corp columnar statement:
       Invoice Date | Due Date | Invoice No | PO Number | Amount Due | Enclosed No
    The Amount Due (e.g. '2,069.75') is the net per invoice; the trailing Enclosed
    number is the invoice # repeated. Tie-out = sum of Amount Due = Total."""
    stmt_date = _norm_date(_grab(r"As of Date:\s*(\d{1,2}/\d{1,2}/\d{4})", full_text))
    amt_due = 0.0
    m = re.search(r"Total Amount Due:\s*([\d,]+\.\d{2})", full_text, re.I)
    if m:
        amt_due = float(m.group(1).replace(",", ""))
    rx = re.compile(r"^(?P<d>\d{1,2}/\d{1,2}/\d{4})\s+\d{1,2}/\d{1,2}/\d{4}\s+"
                    r"(?P<ref>\d+)\s+.*?\s+(?P<amt>-?[\d,]+\.\d{2})\s+\d+\s*$")
    lines: List[StmtLine] = []
    for ln in full_text.splitlines():
        m2 = rx.match(ln.strip())
        if not m2:
            continue
        lines.append(StmtLine(date=_norm_date(m2["d"]), ref=m2["ref"],
                              amount=float(m2["amt"].replace(",", ""))))
    if not amt_due:
        amt_due = round(sum(l.amount for l in lines), 2)
    return "", stmt_date, amt_due, lines


def _money_tok(raw: str) -> float:
    """'$234.50' · '($1.17)' · '-5.00' · '1,234.00-' (trailing minus) -> float."""
    s = raw.strip()
    if s.endswith("-"):
        return -_paren_amount(s[:-1])
    return _paren_amount(s)


def parse_statement_abatix_online(full_text: str) -> Tuple[str, str, float, List[StmtLine]]:
    """Abatix online statement (10/2026): Invoice Number | Invoice Date | Due Date |
    PO | Original Amount | Amount Due, then 'Summary of Invoice Age'. Amount Due
    may be a credit in parentheses ('($1.17)')."""
    stmt_date = _norm_date(_grab(r"As of Date:\s*(\d{1,2}/\d{1,2}/\d{4})", full_text))
    m = (re.search(r"Total Amount Due For Customer:\s*\$?\s*([\d,]+\.\d{2})", full_text, re.I)
         or re.search(r"Total Due\s*\$\s*([\d,]+\.\d{2})", full_text, re.I))
    amt_due = float(m.group(1).replace(",", "")) if m else 0.0
    rx = re.compile(r"^(?P<ref>\d{5,})\s+(?P<d>\d{1,2}/\d{1,2}/\d{4})\s+\d{1,2}/\d{1,2}/\d{4}\s+"
                    r"(?:(?P<po>.*?)\s+)?\(?\$[\d,]+\.\d{2}\)?\s+(?P<amt>\(?\$[\d,]+\.\d{2}\)?)\s*$")
    lines = [StmtLine(date=_norm_date(m2["d"]), ref=m2["ref"], amount=_money_tok(m2["amt"]),
                      po=(m2["po"] or "").strip())
             for m2 in (rx.match(ln.strip()) for ln in full_text.splitlines()) if m2]
    return "", stmt_date, amt_due, lines


def parse_statement_croell_ar(full_text: str) -> Tuple[str, str, float, List[StmtLine]]:
    """Croell 'DEV AR Customer Balance' (10/2026). The 7-digit invoice # runs into
    the PO in the extracted text ('11057227 BREW GRANBURY' = 1105722 + '7 BREW');
    the LAST column (Current Due Less Discounts) is the open amount. Finance
    charges (type F) are part of the balance, like on the older Croell layout."""
    stmt_date = _norm_date(_grab(r"\bDate\s+(\d{2}/\d{2}/\d{4})", full_text))
    m = re.search(r"Balance After Disc\s+(\(?-?[\d,]+\.\d{2}\)?)", full_text, re.I)
    amt_due = _paren_amount(m.group(1)) if m else 0.0
    num = r"\(?-?[\d,]+\.\d{2}\)?"
    rx = re.compile(rf"^(?P<ref>\d{{7}})(?P<po>.*?)\s+(?P<d>\d{{2}}/\d{{2}}/\d{{4}})\s+"
                    rf"(?:\d{{2}}/\d{{2}}/\d{{4}}\s+)?(?P<typ>[A-Z])\s+(?P<nums>(?:{num}\s+){{7}}{num})\s*$")
    lines: List[StmtLine] = []
    for ln in full_text.splitlines():
        m2 = rx.match(ln.strip())
        if m2:
            lines.append(StmtLine(date=_norm_date(m2["d"]), ref=m2["ref"],
                                  amount=_paren_amount(m2["nums"].split()[-1]),
                                  po=m2["po"].strip()))
    return "Croell Inc", stmt_date, amt_due, lines


def parse_statement_quikrete(full_text: str) -> Tuple[str, str, float, List[StmtLine]]:
    """Quikrete Ready Mix account statement: 'RI 34216970 07/17/26 08/16/26 <remark>
    <PO> 47,825.00 47,825.00', grouped by ship-to with 'Total for:' lines. The
    statement is AS OF a date printed under the title. Ref = the number after the
    two-letter type."""
    stmt_date = _norm_date(_grab(r"AS OF\s+(\d{2}/\d{2}/\d{2,4})", full_text))
    m = re.search(r"TOTAL\s+USD\s+(-?[\d,]*\.\d{2}-?)", full_text)
    amt_due = _money_tok(m.group(1)) if m else 0.0
    amt = r"-?[\d,]*\.\d{2}-?"
    rx = re.compile(rf"^(?P<typ>[A-Z]{{2}})\s+(?P<ref>\d{{5,}})\s+(?P<d>\d{{2}}/\d{{2}}/\d{{2}})\s+"
                    rf"(?:\d{{2}}/\d{{2}}/\d{{2}}\s+)?(?P<rest>.*?)\s*(?P<inv>{amt})\s+(?P<bal>{amt})\s*$")
    lines: List[StmtLine] = []
    for ln in full_text.splitlines():
        m2 = rx.match(ln.strip())
        if m2:
            lines.append(StmtLine(date=_norm_date(m2["d"]), ref=m2["ref"],
                                  amount=_money_tok(m2["bal"]), po=m2["rest"].strip()))
    return "Quikrete Ready Mix", stmt_date, amt_due, lines


def parse_statement_qbo_invoice_list(full_text: str) -> Tuple[str, str, float, List[StmtLine]]:
    """A vendor's QuickBooks 'Invoices for <us>' list (CRI-DFW, 10/2026): Num | Date |
    Name | Amount | Open Balance. 'All Transactions' lists paid ones too (open 0.00)
    - only open lines count. As-of = the report's print date."""
    lines_txt = [ln.strip() for ln in full_text.splitlines() if ln.strip()]
    vendor = lines_txt[0] if lines_txt and not re.match(r"\d", lines_txt[0]) else ""
    m = re.search(r"Invoices for .*\n\s*(\d{1,2}/\d{1,2}/\d{2,4})", full_text)
    stmt_date = _norm_date(m.group(1)) if m else ""
    m = re.search(r"^\s*Total\s+-?[\d,]+\.\d{2}\s+(-?[\d,]+\.\d{2})\s*$", full_text, re.M)
    amt_due = _paren_amount(m.group(1)) if m else 0.0
    rx = re.compile(r"^(?P<ref>\S+)\s+(?P<d>\d{1,2}/\d{1,2}/\d{4})\s+(?P<name>.+?)\s+"
                    r"(?P<amt>\(?-?[\d,]+\.\d{2}\)?)\s+(?P<open>\(?-?[\d,]+\.\d{2}\)?)$")
    lines: List[StmtLine] = []
    for ln in lines_txt:
        m2 = rx.match(ln)
        if m2 and _paren_amount(m2["open"]) != 0:
            lines.append(StmtLine(date=_norm_date(m2["d"]), ref=m2["ref"],
                                  amount=_paren_amount(m2["open"]),
                                  address=m2["name"].split(":", 1)[-1].strip()))
    return vendor, stmt_date, amt_due, lines


def parse_statement_voidform(full_text: str) -> Tuple[str, str, float, List[StmtLine]]:
    """VoidForm Products 'Customer Statement' (Acumatica-style): Date | Due Date |
    Doc. Type | Ref. Nbr. | Ext. Ref. Nbr. (our PO) | Orig. Amount | Amount Due |
    Balance; the job address prints on the next line. Amount Due = the open amount."""
    stmt_date = _norm_date(_grab(r"\bDate:\s*(\d{1,2}/\d{1,2}/\d{4})", full_text))
    amt = r"\(?-?[\d,]+\.\d{2}\)?"
    m = re.search(rf"Over 90 Days Past Due\s+Amount Due\s*\n(?:\s*{amt}){{5}}\s+({amt})", full_text)
    amt_due = _paren_amount(m.group(1)) if m else 0.0
    rx = re.compile(rf"^(?P<d>\d{{1,2}}/\d{{1,2}}/\d{{4}})\s+\d{{1,2}}/\d{{1,2}}/\d{{4}}\s+"
                    rf"(?P<typ>Invoice|Credit Memo|Payment|Debit Memo|Prepayment)\s+(?P<ref>\S+)\s+"
                    rf"(?:(?P<po>\S+)\s+)?(?P<orig>{amt})\s+(?P<due>{amt})\s+(?P<bal>{amt})\s*$")
    lines: List[StmtLine] = []
    rows = full_text.splitlines()
    for i, ln in enumerate(rows):
        m2 = rx.match(ln.strip())
        if not m2:
            continue
        a = _paren_amount(m2["due"])
        if m2["typ"] in ("Credit Memo", "Payment", "Prepayment") and a > 0:
            a = -a
        nxt = rows[i + 1].strip() if i + 1 < len(rows) else ""
        lines.append(StmtLine(date=_norm_date(m2["d"]), ref=m2["ref"], amount=a, po=m2["po"] or "",
                              address="" if re.match(r"\d{1,2}/\d{1,2}/\d{4}|Current", nxt) else nxt[:50]))
    if not amt_due:
        amt_due = round(sum(x.amount for x in lines), 2)
    return "VoidForm Products, LLC", stmt_date, amt_due, lines


def parse_statement_cowtown(full_text: str) -> Tuple[str, str, float, List[StmtLine]]:
    stmt_date = ""
    m = re.search(r"([A-Z][a-z]+ \d{1,2}, \d{4})", full_text)
    if m:
        for fmt in ("%B %d %Y", "%b %d %Y"):
            try:
                stmt_date = dt.datetime.strptime(re.sub(r",", "", m.group(1)), fmt).strftime("%Y-%m-%d")
                break
            except ValueError:
                continue
    amt_due = 0.0
    m = re.search(r"Past Due Amount:\s*\$?([\d,]+\.\d{2})", full_text, re.I)
    if m:
        amt_due = float(m.group(1).replace(",", ""))
    rx = re.compile(r"(?P<ref>\d{6}[A-Z]?)\s+(?P<d>\d{1,2}/\d{1,2}/\d{4})\s+\d{1,2}/\d{1,2}/\d{4}\s+\$(?P<inv>[\d,]+\.\d{2})\s+\$(?P<amt>[\d,]+\.\d{2})")
    lines: List[StmtLine] = []
    for m in rx.finditer(full_text):
        lines.append(StmtLine(date=_norm_date(m["d"]), ref=m["ref"],
                              amount=float(m["amt"].replace(",", ""))))
    if not amt_due:
        amt_due = round(sum(l.amount for l in lines), 2)
    return "", stmt_date, amt_due, lines


def parse_statement_sunbelt(full_text: str) -> Tuple[str, str, float, List[StmtLine]]:
    stmt_date = _norm_date(_grab(r"[\r\n]\d{6,8}\s+(\d{1,2}/\d{1,2}/\d{2})\s+\d+", full_text))
    amt_due = 0.0
    m = re.search(r"TOTAL DUE\s*[\r\n]+\$?\s*([\d,]+\.\d{2})", full_text, re.I)
    if m:
        amt_due = float(m.group(1).replace(",", ""))
    rx = re.compile(r"^(?P<d>\d{1,2}/\d{1,2}/\d{2})\s+(?P<ref>\d{9}-\d{4})\s+(?P<desc>.*?)\s*(?P<amt>-?[\d,]+\.\d{2})\s*$")
    lines: List[StmtLine] = []
    for ln in full_text.splitlines():
        m = rx.match(ln.strip())
        if not m:
            continue
        lines.append(StmtLine(date=_norm_date(m["d"]), ref=m["ref"],
                              amount=float(m["amt"].replace(",", ""))))
    if not amt_due:
        amt_due = round(sum(l.amount for l in lines), 2)
    return "", stmt_date, amt_due, lines


def parse_statement_croell(full_text: str) -> Tuple[str, str, float, List[StmtLine]]:
    """Parse a Croell Inc statement (doubled register/remittance layout).
    Takes the left 'Amount' column; credits in parentheses become negatives so
    the line-sum nets to Balance Due."""
    vendor = "Croell Inc"
    # The "Statement Date" header row is doubled and the value sits on the NEXT
    # line ("... Page\n08/08/2026 08/08/2026 ..."), so grab the first date there.
    stmt_date = _norm_date(
        _grab(r"Statement\s+Date[^\n]*\n\s*(\d{1,2}/\d{1,2}/\d{4})", full_text))

    # Balance Due = rightmost value on the aging footer - the last line that is
    # nothing but money tokens (Current / 1-30 / 31-60 / Over 60 / Bal Due).
    amt_due = 0.0
    money_line = re.compile(r"^\s*(?:\(?-?[\d,]+\.\d{2}\)?\s+){2,}\(?-?[\d,]+\.\d{2}\)?\s*$")
    for ln in reversed(full_text.splitlines()):
        if money_line.match(ln):
            amt_due = _paren_amount(re.findall(r"\(?-?[\d,]+\.\d{2}\)?", ln)[-1])
            break

    lines: List[StmtLine] = []
    for m in CROELL_ROW_RE.finditer(full_text):
        # Use the BALANCE column (the still-open amount), not the first "Amount"
        # column (the original charge). A partly-paid invoice reads e.g. charge
        # 14,846.55 / balance 16.44 - summing the charge overshoots the tie-out
        # (INV 1080594, 2026-09-10). Balance == charge for a fully-open invoice.
        lines.append(StmtLine(date=_norm_date(m.group("date")), ref=m.group("num"),
                              amount=_paren_amount(m.group("balance")),
                              address=m.group("desc").strip()))
    if not amt_due:
        amt_due = round(sum(l.amount for l in lines), 2)
    return vendor, stmt_date, amt_due, lines


# ── Staple-vendor identity overrides ──────────────────────────────
# A few big, recurring vendors are impossible to identify from the generic
# body extraction alone: their name is either in a raster logo (no extractable
# text), absent from the letter body, or sits below OUR OWN bill-to name so the
# generic finder grabs "Proficient Concrete" instead. Left to the fuzzy
# filename/first-word fallback they collide with unrelated QBO names
# (BOB→'Bobby Tenison', COW→'COWBOY CHICKEN'). We hard-map any of a vendor's
# stable in-text markers to its EXACT QBO display name, applied across every
# template so it works regardless of which statement layout the vendor sends.
#
# Each entry: (exact QBO DisplayName, [lowercase marker substrings]).
# A marker must be specific enough that it can only mean this vendor. To add a
# staple: pick a token that always appears in its statements (an email domain,
# a letterhead street address, a distinctive brand spelling) and its QBO name.
_STAPLE_VENDORS: List[Tuple[str, List[str]]] = [
    # Bobcat is a multi-dealership brand (QBO has several "Bobcat of ..."); the
    # dealership logo is an image, but the code line ("Statement BOBNTXQ") and
    # credit-manager email ("...@bobcatntx.com") carry the region as text.
    ("Bobcat of North Texas", ["bobcatntx", "bobntx"]),
    # Cowtown Redi Mix: one statement variant is a past-due letter that only
    # shows its letterhead address; another is a QBO statement listing our own
    # name as bill-to. Match ONLY tokens unique to Cowtown - its brand word or
    # its letterhead address. NEVER on "redi mix" / "ready mix": that phrase is
    # on nearly every concrete vendor's statement, so it hard-mapped a Sunrise
    # Redi Mix statement to Cowtown (2026-09-10) - and a wrong staple map
    # reconciles against the wrong QBO vendor, which can clear the wrong bills.
    ("Cowtown Redi Mix Concrete", ["3400 bethlehem", "cowtown"]),
    # Its name heads the statement but the QBO-statement vendor finder misses it,
    # and the filename fallback 'GONZALEZ' matched SEVEN QBO vendors (10/07/2026).
    ("GONZALEZ BROTHERS BATCH PLANT, LP", ["gonzalezreadymix", "gonzalez brothers batch"]),
    # Sends its whole receivables report (all its customers) - the title line is
    # the only place its own name appears.
    ("Core Concrete Pumping LLC", ["core concrete pumping llc"]),
    # A scanned statement (read by OCR) - its name only shows in the web address.
    ("POWER JACK FOUNDATION REPAIR", ["powerjacktexas", "power jack foundation"]),
]


def _staple_vendor_override(full_text: str) -> str:
    """Return the exact QBO name for a known staple vendor if any of its
    markers appear in the statement text, else '' (no override)."""
    t = full_text.lower()
    for qbo_name, markers in _STAPLE_VENDORS:
        if any(marker in t for marker in markers):
            return qbo_name
    return ""


_LAST_TEMPLATE = ""
_LAST_SKIPPED: List["StmtLine"] = []


def last_skipped() -> List["StmtLine"]:
    """Rows the last parse left out ON PURPOSE (another customer's section of a
    vendor's all-customer report) - the marked statement greys them instead of
    flagging them as missed."""
    return list(_LAST_SKIPPED)


def parse_statement_ex(path: Path) -> Tuple[str, str, float, List[StmtLine], str]:
    """parse_statement + the template key that read it ('' for Excel / unknown) -
    statement_set uses it to tell a past-due letter from a full open list."""
    global _LAST_TEMPLATE
    _LAST_TEMPLATE = ""
    _LAST_SKIPPED.clear()
    vendor, stmt_date, amt_due, lines = parse_statement(path)
    return vendor, stmt_date, amt_due, lines, _LAST_TEMPLATE


def parse_statement(path: Path) -> Tuple[str, str, float, List[StmtLine]]:
    """Returns (vendor_guess, stmt_date_YYYY-MM-DD, amount_due_total, lines).
    Dispatches by file extension first:
      • .xlsx/.xls/.xlsm → Excel parser
      • .png/.jpg/.jpeg/.tiff/.bmp/.heic → OCR to text, then run through PDF template detection
      • everything else → PDF text extraction + template detection
    For Excel/Image: vendor and stmt date may come back empty - caller fills
    them from --vendor / --stmt-date or the fallback chain in process_pdf.

    On the text path, a known staple-vendor marker (see _STAPLE_VENDORS)
    overrides whatever vendor the template parser extracted - the staple's exact
    QBO name is more reliable than generic body/filename extraction."""
    ext = path.suffix.lower()
    if ext in (".xlsx", ".xls", ".xlsm"):
        return parse_statement_excel(path)
    if ext in (".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".heic", ".heif"):
        full_text = _image_to_text(path)
    else:
        full_text = _pdf_text(path)
    template = detect_template(full_text)
    global _LAST_TEMPLATE
    _LAST_TEMPLATE = template
    if template == "qbo_customer_open_balance":
        result = parse_statement_qbo_customer_open_balance(full_text)
    elif template == "qbo_open_invoices":
        result = parse_statement_qbo_open_invoices(full_text)
    elif template == "qbo_statement":
        result = parse_statement_qbo_statement(full_text)
    elif template == "vendor_stmt_tabular":
        result = parse_statement_vendor_tabular(full_text)
    elif template == "vendor_stmt_columnar":
        # Columnar parser re-opens the PDF for positional extraction
        # (raw text loses row/column alignment in this layout)
        result = parse_statement_vendor_columnar(path)
    elif template == "vendor_stmt_whitecap":
        result = parse_statement_vendor_whitecap(full_text)
    elif template == "vendor_bobcat":
        result = parse_statement_bobcat(full_text)
    elif template == "vendor_bodin":
        result = parse_statement_bodin(full_text)
    elif template == "vendor_burnco":
        result = parse_statement_burnco(full_text)
    elif template == "vendor_cintas":
        result = parse_statement_cintas(full_text)
    elif template == "vendor_cowtown":
        result = parse_statement_cowtown(full_text)
    elif template == "vendor_sunbelt":
        result = parse_statement_sunbelt(full_text)
    elif template == "vendor_croell":
        result = parse_statement_croell(full_text)
    elif template == "vendor_abatix":
        result = parse_statement_abatix(full_text)
    elif template == "vendor_abatix_online":
        result = parse_statement_abatix_online(full_text)
    elif template == "vendor_croell_ar":
        result = parse_statement_croell_ar(full_text)
    elif template == "vendor_quikrete":
        result = parse_statement_quikrete(full_text)
    elif template == "qbo_invoice_list":
        result = parse_statement_qbo_invoice_list(full_text)
    elif template == "vendor_voidform":
        result = parse_statement_voidform(full_text)
    else:
        # No supported template detected - return empty so the caller surfaces
        # the unsupported-template error with the full list of supported formats.
        return "", "", 0.0, []
    # Staple-vendor identity wins over the parser's extracted vendor.
    override = _staple_vendor_override(full_text)
    if override:
        _vendor, stmt_date, amt_due, lines = result
        result = (override, stmt_date, amt_due, lines)
    return result


def _scan_excel_for_vendor(xlsx_path: Path) -> str:
    """Scan the first 5 rows × 5 columns of an Excel for text that looks
    like a vendor name. Skips header keywords, dates, numeric cells, our own
    company. Prefers cells with corporate suffix tokens (LLC/INC/CO/etc.) or
    industry-typical words. Returns first matching candidate or ''."""
    try:
        from openpyxl import load_workbook
        wb = load_workbook(xlsx_path, data_only=True, read_only=True)
        ws = wb.active
    except Exception:
        return ""

    HEADER_KW = ("invoice", "inv date", "inv #", "due date", "balance",
                 "amount", "ref #", "ref#", "statement date", "page",
                 "total", "type", "date", "memo", "open", "as of")
    candidates: List[str] = []
    for row in ws.iter_rows(min_row=1, max_row=5, max_col=5, values_only=True):
        for cell in row:
            if not isinstance(cell, str):
                continue
            text = cell.strip()
            if not text or len(text) < 4:
                continue
            text_low = text.lower()
            if any(kw in text_low for kw in HEADER_KW):
                continue
            if re.match(r"^\d{1,4}[/.\-]\d{1,2}", text):   # MM/DD or YYYY-MM (date-ish)
                continue
            if re.match(r"^\$?[\d,]+\.?\d*$", text):       # money/numeric
                continue
            if "proficient" in text_low:                    # our own company
                continue
            candidates.append(text)
    # Try wb.close() if read_only mode allocated handles
    try: wb.close()
    except Exception: pass

    # Prefer cells with strong vendor-name signals
    STRONG = re.compile(
        r"\b(LLC|L\.L\.C\.|INC|CO\.|CORP|COMPANY|LTD|MATERIALS|SERVICES|"
        r"MIX|CABLE|TENSION|CONCRETE|SUPPLY|READY|EQUIPMENT|HOLDINGS)\b",
        re.I,
    )
    for c in candidates:
        if STRONG.search(c):
            return c
    return candidates[0] if candidates else ""


def _vendor_from_filename(path: Path) -> str:
    """Derive a vendor-name hint from a filename. Strips dates (slash/dash/
    underscore-separated, month names, 4-digit years), generic words like
    'statement'/'invoice', and separators. Returns '' if nothing left."""
    stem = path.stem
    # Normalize separators to spaces first
    stem = re.sub(r"[_\-.]+", " ", stem)
    # Date patterns (now space-separated after normalization)
    stem = re.sub(r"\b\d{4}\s+\d{1,2}\s+\d{1,2}\b", " ", stem)   # YYYY MM DD
    stem = re.sub(r"\b\d{1,2}\s+\d{1,2}\s+\d{2,4}\b", " ", stem) # MM DD YYYY
    stem = re.sub(r"\b\d{4}\s+\d{1,2}\b", " ", stem)             # YYYY MM
    stem = re.sub(r"\b\d{1,2}/\d{1,2}/\d{2,4}\b", " ", stem)     # MM/DD/YY (original slashes)
    stem = re.sub(r"\b(?:19|20)\d{2}\b", " ", stem)              # standalone YYYY
    # Month names (full + abbreviated)
    stem = re.sub(r"\b(january|february|march|april|may|june|july|august|"
                  r"september|october|november|december|jan|feb|mar|apr|jun|"
                  r"jul|aug|sept?|oct|nov|dec)\b", " ", stem, flags=re.I)
    # Generic statement-related words
    stem = re.sub(r"\b(statement|stmt|invoice|invoices|bill|bills|report|open|"
                  r"aging|monthly|thru|through|as\s+of|past\s+due)\b",
                  " ", stem, flags=re.I)
    # Collapse whitespace
    stem = re.sub(r"\s+", " ", stem).strip()
    # Strip any remaining standalone leading/trailing digit tokens
    stem = re.sub(r"^(\d+\s+)+", "", stem)
    stem = re.sub(r"(\s+\d+)+$", "", stem)
    stem = stem.strip()
    # If what's left has no real letters (e.g., just orphan digits), return ''
    if not re.search(r"[A-Za-z]{2,}", stem):
        return ""
    return stem


def _excel_date_to_str(val) -> str:
    """Convert any Excel-readable date cell to YYYY-MM-DD, or '' if unparseable."""
    if val is None:
        return ""
    if isinstance(val, dt.datetime):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, dt.date):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, (int, float)):
        # Excel serial date (days since 1899-12-30)
        try:
            base = dt.datetime(1899, 12, 30)
            return (base + dt.timedelta(days=int(val))).strftime("%Y-%m-%d")
        except (ValueError, OverflowError):
            return ""
    if isinstance(val, str):
        return _norm_date(val.strip())
    return ""


def _xls_to_xlsx_temp(xls_path: Path) -> Path:
    """Convert a legacy .xls (Excel 97-2003) workbook to a temp .xlsx file
    and return the new path. openpyxl can't read .xls, so we use xlrd<2.0
    to read and rewrite as .xlsx. The temp file is created in the OS temp
    dir; macOS cleans these up automatically. Date cells (xlrd stores them
    as floats with a datemode flag) are converted to real datetime objects
    so the downstream parser sees them as dates not numbers.
    Lazy-imports xlrd so users who never see .xls files don't need it."""
    import tempfile
    try:
        import xlrd
    except ImportError:
        sys.exit(_Term.color(_Term.R,
            "✗ Legacy .xls support needs xlrd (pinned in python-env). Repair the environment:\n"
            "    bash python-env/setup.sh"))
    from openpyxl import Workbook

    try:
        xls = xlrd.open_workbook(str(xls_path))
    except Exception as e:
        sys.exit(_Term.color(_Term.R, f"✗ couldn't open .xls file: {e}"))

    sheet = xls.sheet_by_index(0)
    new_wb = Workbook()
    new_ws = new_wb.active
    for r in range(sheet.nrows):
        for c in range(sheet.ncols):
            cell = sheet.cell(r, c)
            v = cell.value
            # xlrd encodes dates as floats - convert back to datetime
            if cell.ctype == xlrd.XL_CELL_DATE:
                try:
                    v = xlrd.xldate.xldate_as_datetime(v, xls.datemode)
                except (xlrd.XLDateError, ValueError):
                    pass
            new_ws.cell(row=r + 1, column=c + 1, value=v)
    tmp = Path(tempfile.gettempdir()) / f"_reconcile_{xls_path.stem}.xlsx"
    new_wb.save(tmp)
    return tmp


def _find_excel_as_of_date(ws) -> str:
    """Scan the top 10 rows for an 'As of <Month DD, YYYY>' label and return
    the parsed YYYY-MM-DD. Same convention as the QBO PDF templates."""
    AS_OF = re.compile(r"As\s+of\s+([A-Za-z]+\s+\d{1,2},?\s+\d{4})", re.I)
    for row in ws.iter_rows(min_row=1, max_row=10, values_only=True):
        for cell in row:
            if not isinstance(cell, str):
                continue
            m = AS_OF.search(cell)
            if not m:
                continue
            raw = re.sub(r",", "", m.group(1)).strip()
            for fmt in ("%B %d %Y", "%b %d %Y"):
                try:
                    return dt.datetime.strptime(raw, fmt).strftime("%Y-%m-%d")
                except ValueError:
                    continue
    return ""


def _find_excel_grand_total(ws) -> float:
    """Look for a 'Total ...' row (e.g., 'Total - Cust00000058 ...') and grab
    its rightmost positive numeric - that's the grand total for tie-out.
    Uses the LAST matching 'total' row: on a multi-customer A/R aging export
    the per-customer subtotals are also labeled 'Total ...' and appear first,
    so taking the first one would return a subtotal. The grand total is last.
    Single-vendor statements have exactly one 'total' row, so this is a no-op
    for them."""
    grand = 0.0
    for row in ws.iter_rows(values_only=False):
        first_text = ""
        for cell in row:
            if isinstance(cell.value, str) and cell.value.strip():
                first_text = cell.value.strip().lower()
                break
        if first_text.startswith("total"):
            nums = [c.value for c in row if isinstance(c.value, (int, float)) and c.value > 0]
            if nums:
                grand = float(max(nums))
    return grand


def _parse_excel_with_headers(ws) -> Tuple[List["StmtLine"], float]:
    """Header-based parser: scan first 10 rows for a header row with named
    columns (Invoice # + Original Invoice Amount, etc.). Map columns by name."""
    header_row = None
    col_map: Dict[str, int] = {}
    for row_idx in range(1, min(11, ws.max_row + 1)):
        mapping: Dict[str, int] = {}
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=row_idx, column=c).value
            if not v:
                continue
            vs = str(v).lower().strip()
            if "inv date" == vs or "invoice date" in vs:
                mapping.setdefault("date", c)
            elif "invoice #" in vs or vs in ("inv #", "invoice number", "invoice no", "invoice #"):
                mapping.setdefault("ref", c)
            elif "original invoice amount" in vs or vs == "original amount":
                mapping["amount"] = c
            elif vs == "amount" and "amount" not in mapping:
                mapping["amount"] = c
            elif vs in ("balance", "open balance", "open bal"):
                mapping.setdefault("balance", c)
            elif "due date" in vs or vs == "due":
                mapping.setdefault("due_date", c)
        if "ref" in mapping and "amount" in mapping:
            header_row = row_idx
            col_map = mapping
            break

    if header_row is None:
        return [], 0.0

    lines: List[StmtLine] = []
    total = 0.0
    for row_idx in range(header_row + 1, ws.max_row + 1):
        ref_v = ws.cell(row=row_idx, column=col_map["ref"]).value
        amount_v = ws.cell(row=row_idx, column=col_map["amount"]).value
        if ref_v is None or amount_v is None:
            continue
        ref = str(ref_v).strip()
        if not ref:
            continue
        try:
            amount = float(amount_v)
        except (ValueError, TypeError):
            continue
        date_v = ws.cell(row=row_idx, column=col_map["date"]).value if "date" in col_map else None
        lines.append(StmtLine(date=_excel_date_to_str(date_v), ref=ref, amount=amount, po="", address=""))
        total += amount
    return lines, round(total, 2)


def _parse_excel_content_pattern(ws) -> Tuple[List["StmtLine"], float]:
    """Header-less parser for layouts like A/R Aging Detail reports.

    Identifies an invoice row by content alone:
      • at least one date cell (datetime or MM/DD/YY-style string)
      • at least one ref-like token (letters + digits like INV38014 / CV6415,
        or pure digits ≥3 chars) - taken from the FIRST whitespace-delimited
        token of any text cell
      • a rightmost positive numeric cell → treated as the amount

    Rows missing any of these (headers, totals, blank rows, customer-header
    rows like 'Cust00000058 Proficient Concrete LLC') are skipped."""
    REF_TOKEN = re.compile(r"^[A-Za-z]*\d{3,}$")
    DATE_STR  = re.compile(r"^\d{1,2}[/.\-]\d{1,2}[/.\-]\d{2,4}$")
    lines: List[StmtLine] = []
    total = 0.0
    for row in ws.iter_rows(values_only=False):
        date_val = None
        ref_val = None
        last_numeric_val = None
        last_numeric_col = -1
        for cell in row:
            v = cell.value
            if v is None:
                continue
            if isinstance(v, (dt.datetime, dt.date)):
                if date_val is None:
                    date_val = v
                continue
            if isinstance(v, str):
                vs = v.strip()
                if not vs:
                    continue
                if DATE_STR.match(vs) and date_val is None:
                    date_val = vs
                    continue
                if ref_val is None:
                    first_token = vs.split()[0] if vs.split() else ""
                    if REF_TOKEN.match(first_token):
                        ref_val = first_token
                continue
            if isinstance(v, (int, float)) and v > 0:
                if cell.column > last_numeric_col:
                    last_numeric_val = float(v)
                    last_numeric_col = cell.column
        if date_val and ref_val and last_numeric_val is not None:
            lines.append(StmtLine(
                date=_excel_date_to_str(date_val),
                ref=ref_val,
                amount=last_numeric_val,
                po="", address="",
            ))
            total += last_numeric_val
    return lines, round(total, 2)


def parse_statement_excel(xlsx_path: Path) -> Tuple[str, str, float, List[StmtLine]]:
    """Parse an Excel statement. Two-pass detection:
       1. Header-based - looks for 'Invoice #' + 'Original Invoice Amount' (or
          similar) on a header row, then reads data rows by column position.
       2. Content-pattern (fallback) - header-less layouts like A/R Aging
          Detail reports. Detects invoice rows by content (date + ref# +
          rightmost numeric).

    Accepts both modern .xlsx and legacy .xls (Excel 97-2003). .xls files
    are transparently converted to a temp .xlsx first via xlrd.

    Statement date is auto-extracted from 'As of <Month DD, YYYY>' if present
    in the top rows. Grand total is taken from a 'Total ...' row when found,
    otherwise it's the sum of parsed line amounts.

    Vendor name is NOT extracted here - that's handled by the chain in
    process_pdf (--vendor flag → cached alias → Excel cell scan → cleaned
    filename → interactive prompt)."""
    # Legacy .xls → convert to temp .xlsx first so the rest of the code can
    # use openpyxl uniformly. _xls_to_xlsx_temp handles the conversion + date
    # cell normalization; openpyxl-only path stays unchanged.
    if xlsx_path.suffix.lower() == ".xls":
        xlsx_path = _xls_to_xlsx_temp(xlsx_path)
    from openpyxl import load_workbook
    wb = load_workbook(xlsx_path, data_only=True)
    ws = wb.active

    stmt_date = _find_excel_as_of_date(ws)

    # 1. Try header-based parser
    lines, total = _parse_excel_with_headers(ws)
    # 2. Fall through to content-pattern parser if no lines extracted
    if not lines:
        lines, total = _parse_excel_content_pattern(ws)

    # Prefer the "Total ..." row's amount if found - that's the report's own
    # grand total (more reliable than summing if any line was missed).
    grand_total = _find_excel_grand_total(ws)
    if grand_total > 0:
        total = grand_total

    return "", stmt_date, round(total, 2), lines


def parse_statement_vendor_tabular(full_text: str) -> Tuple[str, str, float, List[StmtLine]]:
    """Parse a tabular vendor statement (CMC-style). One line per invoice:
       <date> <ref> <due-date> <amount> [optional payment fields] <balance>
    Sub-customer headers (SHIP TO XYZ) and aging rows are skipped."""
    # Vendor - first non-empty line at the top usually has the company name.
    # CMC-style PDFs concatenate vendor name + "STATEMENT" on one line
    # ("CMC Construction Services STATEMENT"); strip the trailing word.
    vendor = ""
    for line in full_text.splitlines()[:5]:
        cand = line.strip()
        if not cand:
            continue
        # Strip trailing "STATEMENT" (case-insensitive) - common artifact
        cand = re.sub(r"\s+STATEMENT\s*$", "", cand, flags=re.I).strip()
        if not cand:
            continue
        if _is_vendor_noise(cand):
            continue
        # Skip address/state lines and phone numbers
        if re.match(r"^[\d\-\(\) ]+$", cand):
            continue
        vendor = cand
        break

    # Statement date - try labeled patterns first
    stmt_date = ""
    for pat in STMT_DATE_PATTERNS:
        m = pat.search(full_text)
        if m:
            stmt_date = _norm_date(m.group(1))
            if stmt_date:
                break

    # Amount due - look for "Total Due" label, then take the LAST dollar value
    # within the next ~200 chars (aging row format: labels on one line, totals
    # on next line; Total Due is the rightmost column).
    amt_due = 0.0
    m = re.search(r"Total\s+Due([\s\S]{0,200})", full_text, re.I)
    if m:
        nums = re.findall(r"[\d,]+\.\d{2}", m.group(1))
        if nums:
            try:
                amt_due = float(nums[-1].replace(",", ""))
            except ValueError:
                pass
    # Fallback to AMT_DUE_RE
    if amt_due == 0.0:
        m = AMT_DUE_RE.search(full_text)
        if m:
            try:
                amt_due = float(m.group(1).replace(",", ""))
            except ValueError:
                pass

    # Line items
    lines: List[StmtLine] = []
    current_subcust = ""
    for line in full_text.splitlines():
        stripped = line.strip()
        if not stripped:
            continue
        # Track sub-customer headers like "SHIP TO BRIW BRIARWOOD"
        if VENDOR_STMT_SHIPTO_RE.match(line):
            # Pull the description after "SHIP TO" / "SHIP-TO"
            sub = re.sub(r"^\s*SHIP[- ]TO\s*", "", stripped, flags=re.I).strip()
            current_subcust = sub
            continue
        # Skip "SHIP-TO TOTAL" lines
        if re.match(r"^\s*SHIP[- ]TO\s+TOTAL\b", line, re.I):
            continue
        # Skip aging row (Current 31-60 ...)
        if re.match(r"^\s*Current\s+\d|^\s*\d+\.\d+\s+\d+\.\d+\s+\d+\.\d+", line):
            continue
        # Try to match an invoice row
        m = VENDOR_STMT_TABULAR_LINE_RE.match(line)
        if not m:
            continue
        try:
            amount = float(m.group("amount").replace(",", ""))
            balance = float(m.group("balance").replace(",", ""))
        except (ValueError, TypeError):
            continue
        # Use BALANCE as the open amount (more accurate than original amount -
        # accounts for partial payments). For unpaid bills balance == amount.
        # Skip zero-balance rows (fully paid).
        if balance <= 0:
            continue
        lines.append(StmtLine(
            date=_norm_date(m.group("date")),
            ref=m.group("ref"),
            amount=balance,
            po=current_subcust,
            address="",
        ))
    return vendor, stmt_date, amt_due, lines


def parse_statement_vendor_columnar(pdf_path: Path) -> Tuple[str, str, float, List[StmtLine]]:
    """Parse a columnar vendor statement (Preferred Materials / Sunrise style).
    These have a visible table grid but pdfplumber's default extraction
    flattens each column into a vertical stack. We re-extract with positional
    coordinates and group words by Y (row) - then map by X (column header)."""
    import pdfplumber
    lines: List[StmtLine] = []
    vendor = ""
    stmt_date = ""
    amt_due_text = ""

    with pdfplumber.open(pdf_path) as pdf:
        all_words: List[dict] = []
        for page_idx, page in enumerate(pdf.pages):
            page_words = page.extract_words(use_text_flow=False)
            # Tag each word with which page it came from (for Y-deduplication across pages)
            for w in page_words:
                w["page"] = page_idx
                all_words.append(w)

        # Vendor + stmt date from page-1 text (still useful since headers
        # come through as roughly-sequential lines). Strip trailing
        # "Date MM/DD/YYYY" / "STATEMENT" artifacts that get concatenated
        # onto the vendor name by pdfplumber's flat text extraction.
        page1_text = pdf.pages[0].extract_text() or ""
        for line in page1_text.splitlines()[:8]:
            cand = line.strip()
            # Strip trailing "Date MM/DD/YYYY" suffix
            cand = re.sub(r"\s+Date\s+\d{1,2}/\d{1,2}/\d{2,4}\s*$", "", cand, flags=re.I).strip()
            # Strip trailing STATEMENT word
            cand = re.sub(r"\s+STATEMENT\s*$", "", cand, flags=re.I).strip()
            if cand and not _is_vendor_noise(cand) and len(cand) > 3:
                vendor = cand
                break
        # Statement date - usually "Date MM/DD/YYYY"
        m = re.search(r"\bDate\s+(\d{1,2}/\d{1,2}/\d{2,4})", page1_text)
        if m:
            stmt_date = _norm_date(m.group(1))
        # Amount Due
        m = re.search(r"Amount\s+Due\s*\$?\s*([\d,]+\.\d{2})", page1_text, re.I)
        if m:
            try:
                amt_due_text = m.group(1).replace(",", "")
            except Exception:
                pass

    # Group all words across all pages into rows by (page, y-coordinate)
    # Y proximity threshold: 3 pixels (rows are usually 12-15px apart)
    rows_by_pos: Dict[Tuple[int, int], List[dict]] = {}
    for w in all_words:
        key = (w["page"], round(w["top"] / 3))  # bucket by 3-pixel y-band
        rows_by_pos.setdefault(key, []).append(w)

    # Each data row's RIGHTMOST number is the running Balance. A line's amount is
    # the balance DELTA from the previous row. That nets credit memos, payments
    # and partial payments - and gets their SIGN right - without needing to know
    # which of the Charge / Payment / Open-Amount columns is populated, so the
    # line sum always ties to the closing balance / Amount Due.
    #   Was: "leftmost positive numeric = charge", which summed GROSS charges and
    #   dropped every non-"Invoice" row, overshooting the tie-out by the paid-down
    #   and credit amounts (Sunrise Redi Mix, Preferred Materials, 2026-09-10).
    MONEY_TOKEN = re.compile(r"^\(?-?\$?[\d,]+\.\d{2}\)?$")
    DATE_TOKEN = re.compile(r"^\d{1,2}/\d{1,2}/\d{2,4}$")
    REF_RE = re.compile(r"#\s*([A-Za-z0-9_\-]+)")
    prev_balance = 0.0
    for key, words in sorted(rows_by_pos.items()):
        words.sort(key=lambda w: w["x0"])
        tokens = [w["text"] for w in words]
        # Data rows lead with a date; header / bill-to / aging-footer rows don't.
        date_token = next((t for t in tokens if DATE_TOKEN.match(t)), None)
        if not date_token:
            continue
        money = [t for t in tokens if MONEY_TOKEN.match(t)]
        if not money:
            continue
        balance = _paren_amount(money[-1])            # rightmost column = running balance
        amount = round(balance - prev_balance, 2)     # delta nets credits / payments / partials
        prev_balance = balance
        if amount == 0.0:
            continue                                  # $0 / informational row
        row_text = " ".join(tokens)
        # A "Balance Forward" row carries a prior open balance (often a credit,
        # e.g. -645.91). It has no invoice ref, but it MUST count toward the sum
        # or the tie-out is off by the whole forward - so it flows through as a
        # normal delta line (balance - 0), just labelled and ref-less. A $0.00
        # forward nets to 0 above and is dropped.
        is_fwd = re.search(r"balance\s+forward", row_text, re.I)
        ref_m = None if is_fwd else REF_RE.search(row_text)
        lines.append(StmtLine(
            date=_norm_date(date_token),
            ref=ref_m.group(1) if ref_m else "",
            amount=amount,
            po="",
            address="Balance forward (prior open balance, not itemized)" if is_fwd else "",
        ))

    amt_due = 0.0
    if amt_due_text:
        try:
            amt_due = float(amt_due_text)
        except ValueError:
            pass
    if amt_due == 0.0:
        amt_due = round(prev_balance, 2)              # closing running balance

    return vendor, stmt_date, amt_due, lines


def parse_statement_vendor_whitecap(full_text: str) -> Tuple[str, str, float, List[StmtLine]]:
    """Parse a White Cap / Billtrust statement. pdfplumber merges the left
    (statement) and right (remittance advice) copies onto a single line per
    row, so each data row in the extracted text looks like:
        <date> <transaction_no> <type> <original> <balance> <ref-dup> <po> <balance-dup>
    We capture the first 5 fields and ignore the duplicate trailing ones.

    Type codes: I=Invoice, C=Credit Memo, R=Rental, D=Debit Memo,
    U=Unapplied Payment, * suffix = In Review. We include I/R/D rows (real
    bills), skip C and U (vendor accounting adjustments, not real bills).
    Uses BALANCE DUE as the amount (already accounts for partial payments)."""
    vendor = "White Cap, L.P."

    # Statement date - "CLOSING DATE" label, then date on a nearby line
    stmt_date = ""
    m = re.search(r"CLOSING\s+DATE[\s\S]{0,40}?(\d{1,2}/\d{1,2}/\d{2,4})",
                  full_text, re.I)
    if m:
        stmt_date = _norm_date(m.group(1))

    # Total Due
    amt_due = 0.0
    m = re.search(r"TOTAL\s+DUE[\s\S]{0,40}?\$?([\d,]+\.\d{2})", full_text, re.I)
    if m:
        try:
            amt_due = float(m.group(1).replace(",", ""))
        except ValueError:
            pass

    # Keep EVERY type as a line, credits/payments included (negative amounts).
    # Unapplied Payments (U) and Credit Memos (C) are netted into the vendor's
    # Total Due, so dropping them makes the line sum overshoot and the tie-out
    # fail (White Cap 07-01: a -3,173.96 type-U payment). They carry no invoice
    # ref, so they surface as MISSING_IN_QBO - which is correct: AP sees the
    # open credit/payment. (Was: skip U, which broke the tie-out.)
    SKIP_TYPES: set = set()
    lines: List[StmtLine] = []
    for m in WHITECAP_ROW_RE.finditer(full_text):
        tp = m.group("tp").rstrip("*")  # strip "*" in-review marker
        if tp in SKIP_TYPES:
            continue
        try:
            balance = float(m.group("balance").replace(",", ""))
        except (ValueError, TypeError):
            continue
        if balance == 0:
            continue
        lines.append(StmtLine(
            date=_norm_date(m.group("date")),
            ref=m.group("ref"),
            amount=balance,
            po="",
            address="",
        ))

    return vendor, stmt_date, amt_due, lines


def parse_statement_qbo_statement(full_text: str) -> Tuple[str, str, float, List[StmtLine]]:
    """Parse a QuickBooks Statement (vendor-issued statement with INV #N. Due N lines)."""
    vendor = _find_qbo_statement_vendor(full_text)

    # Statement date - try labeled patterns first, then standalone fallback
    stmt_date = ""
    for pat in STMT_DATE_PATTERNS:
        m = pat.search(full_text)
        if m:
            stmt_date = _norm_date(m.group(1))
            if stmt_date:
                break

    # Amount due - first match after "Amount Due" label
    amt_due = 0.0
    m = AMT_DUE_RE.search(full_text)
    if m:
        try:
            amt_due = float(m.group(1).replace(",", ""))
        except ValueError:
            pass

    # Lines - match against full text (multiline)
    lines: List[StmtLine] = []
    last_balance = None
    for m in STMT_LINE_RE.finditer(full_text):
        last_balance = _paren_amount(m.group("balance"))
        date = _norm_date(m.group("date"))
        ref = m.group("ref")
        amount = _paren_amount(m.group("amount"))   # signed: credits/payments net out
        # Pull PO and address from the matched line+surrounding chars - invoice
        # rows only. The 200-char look-ahead (needed for addresses that wrap
        # across rows) would otherwise bleed the NEXT invoice's PO/address onto
        # a credit/payment row, which carries neither.
        line_blob = full_text[max(0, m.start() - 20): m.end() + 200] if amount > 0 else ""
        po_m = PO_RE.search(line_blob)
        po = po_m.group(1).strip() if po_m else ""
        # Address: text between "Orig. Amount $X.XX." and the start of the
        # NEXT line (or end of table). PDF tables interject the line $amount
        # and running balance, and wrap long addresses across rows - both
        # need stripping.
        addr = ""
        orig_m = ORIG_AMT_RE.search(line_blob)
        if orig_m:
            tail = line_blob[orig_m.end():]
            # Cut at the next line boundary - whichever comes first:
            #   • next date (MM/DD/YYYY)
            #   • next "INV #..." marker
            #   • aging-bucket footer strings ("DAYS PAST", "CURRENT", "Amount Due")
            cutoff_patterns = [
                r"\d{1,2}/\d{1,2}/\d{2,4}",   # next line's date
                r"INV\s*\#",                  # next line's invoice marker
                r"\d{1,2}-\d{1,2}\b",         # aging range like "1-30", "31-60"
                r"DAYS\s+PAST",               # aging footer
                r"\bCURRENT\b",
                r"Amount\s+Due",
                r"OVER\s+\d+\s+DAYS",         # "OVER 90 DAYS PAST DUE"
            ]
            cutoff_re = re.compile("|".join(cutoff_patterns), re.I)
            m2 = cutoff_re.search(tail)
            if m2:
                tail = tail[:m2.start()]
            # Strip the "$N,NNN.NN $N,NNN.NN" pair (line amt + running balance)
            tail = re.sub(r"\s*[\d,]+\.\d{2}\s+[\d,]+\.\d{2}\s*", " ", tail)
            # Strip leading period+space left over from "Orig. Amount $X.XX. <addr>"
            tail = re.sub(r"^[\s.]+", "", tail)
            # Collapse newlines + multi-space to single space
            tail = re.sub(r"\s+", " ", tail).strip()
            # Truncate cleanly - first 50 chars, cut at last word boundary
            if len(tail) > 50:
                tail = tail[:50].rsplit(" ", 1)[0]
            addr = tail
        lines.append(StmtLine(date=date, ref=ref, amount=amount, po=po, address=addr))

    # A payment row with NO check number ("09/21/2026  PMT  -3,210.00  806,994.22",
    # CowTown 10-02) has no ref for STMT_LINE_RE to anchor on - read it here or the
    # statement never ties out. Ref = 'PMT' (payments are never matched to bills).
    # Same for a 'Discount -0.01' row (CowTown 07-02: six of them, why it was 0.13 off -
    # caught by the marked-up statement, 10/08/2026).
    for m in PMT_NO_REF_RE.finditer(full_text):
        lines.append(StmtLine(date=_norm_date(m.group(1)),
                              ref="PMT" if m.group(2) == "PMT" else "DISC",
                              amount=_paren_amount(m.group(3))))

    # The statement date sits in the header (right column, under "Date"), sometimes
    # too far from its label for STMT_DATE_PATTERNS - which then grab the FIRST
    # transaction date (Carder 10/07/2026 read as 08/05). A statement can't predate
    # its own newest line, so in that case take the header date that doesn't.
    newest = max((l.date for l in lines if re.match(r"\d{4}-\d\d-\d\d$", l.date or "")), default="")
    if newest and (not stmt_date or stmt_date < newest):
        first = STMT_LINE_RE.search(full_text)
        head = full_text[:first.start()] if first else full_text[:1500]
        cands = sorted(d for d in (_norm_date(x) for x in re.findall(r"\d{1,2}/\d{1,2}/\d{4}", head))
                       if d and d >= newest)
        if cands:
            stmt_date = cands[-1]

    # Amount Due box unreadable (a scanned statement, Power Jack 07-28): the running
    # balance of the last row is the same figure. The tie-out still has to hold, so
    # the two columns check each other.
    if not amt_due and last_balance:
        amt_due = last_balance

    # A "Balance forward" row rolls all prior-period open items into one opening
    # amount. It has no invoice ref and no running-balance pair, so STMT_LINE_RE
    # skips it - but it MUST be counted or the line sum falls short of Amount Due
    # by the entire carried balance (CowTown 09-01: a 638,262.39 forward on an
    # 810,837.10 statement). It is a lump - prior invoices aren't itemized here -
    # so it carries no ref to match against a QBO bill.
    bf = re.search(
        r"^\s*(\d{1,2}/\d{1,2}/\d{2,4})\s+Balance\s+forward\s+"
        r"(\(?-?[\d,]+\.\d{2}\)?)\s*$",
        full_text, re.I | re.MULTILINE)
    if bf and _paren_amount(bf.group(2)) != 0.0:
        lines.insert(0, StmtLine(
            date=_norm_date(bf.group(1)), ref="",
            amount=_paren_amount(bf.group(2)),
            address="Balance forward (prior open balance, not itemized)"))

    return vendor, stmt_date, amt_due, lines


def _norm_date(s: str) -> str:
    """MM/DD/YYYY or M/D/YY → YYYY-MM-DD."""
    for fmt in ("%m/%d/%Y", "%m/%d/%y"):
        try:
            return dt.datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return s

# ───────────────────────── QBO auth + queries ─────────────────────────

def load_credentials() -> Tuple[str, str]:
    """(access token, company id) from THE shared QuickBooks login (shared/qbo_api), which asks Key Helper
    once the key library is adopted. This tool's own copy of the refresh-token exchange was retired
    09/29/2026 (security review): one login path, and no tool ever holds the refresh token."""
    return qbo_api.load_credentials()


def _api_get(access: str, path: str, params: Optional[dict] = None) -> dict:
    p = dict(params or {})
    p["minorversion"] = MINOR_VERSION
    r = requests.get(f"{API_BASE}{path}",
                     headers={"Authorization": f"Bearer {access}",
                              "Accept": "application/json"},
                     params=p, timeout=60)
    if r.status_code != 200:
        raise RuntimeError(f"GET {path} → {r.status_code}: {r.text[:300]}")
    return r.json()


def query(access: str, cid: str, q: str) -> dict:
    return _api_get(access, f"/v3/company/{cid}/query", {"query": q})


def query_all(access: str, cid: str, entity: str, where: str = "") -> List[dict]:
    out: List[dict] = []
    start = 1
    page = 500
    while True:
        q = f"SELECT * FROM {entity}"
        if where:
            q += f" WHERE {where}"
        q += f" STARTPOSITION {start} MAXRESULTS {page}"
        data = query(access, cid, q)
        batch = data.get("QueryResponse", {}).get(entity, [])
        if not batch:
            break
        out.extend(batch)
        if len(batch) < page:
            break
        start += page
    return out


# ───────────────────────── vendor alias cache ─────────────────────────

def _alias_key(s: str) -> str:
    """Normalize a vendor string for cache lookup. Preserves visual identity
    but ignores trailing whitespace/punctuation noise that varies by PDF run."""
    return re.sub(r"[\s\.,;]+$", "", (s or "").strip()).lower()


def load_aliases() -> Dict[str, dict]:
    """Return {alias_key: {'pdf_name', 'qbo_id', 'qbo_name', 'saved'}} or {} if missing/corrupt."""
    if not ALIAS_FILE.exists():
        return {}
    try:
        data = json.loads(ALIAS_FILE.read_text())
        if not isinstance(data, dict):
            return {}
        return data
    except (json.JSONDecodeError, OSError) as e:
        _warn(f"vendor_aliases.json unreadable ({e}); ignoring cache.")
        return {}


def save_aliases(aliases: Dict[str, dict]) -> None:
    try:
        ALIAS_FILE.write_text(json.dumps(aliases, indent=2, sort_keys=True))
    except OSError as e:
        _warn(f"could not save vendor_aliases.json ({e}); alias not persisted.")


def remember_vendor(pdf_name: str, qbo_id: str, qbo_name: str) -> None:
    aliases = load_aliases()
    key = _alias_key(pdf_name)
    aliases[key] = {
        "pdf_name": pdf_name,
        "qbo_id": qbo_id,
        "qbo_name": qbo_name,
        "saved": dt.date.today().isoformat(),
    }
    save_aliases(aliases)


def forget_vendor(needle: str) -> bool:
    """Remove one alias. Matches by alias_key OR by qbo_name (case-insensitive)."""
    aliases = load_aliases()
    target = _alias_key(needle)
    to_remove = [k for k, v in aliases.items()
                 if k == target or v.get("qbo_name", "").lower() == needle.lower()]
    for k in to_remove:
        del aliases[k]
    if to_remove:
        save_aliases(aliases)
    return bool(to_remove)


def find_alias_by_filename(filename: str) -> Optional[Tuple[str, dict, int, str]]:
    """Fuzzy-match a filename against cached aliases by token overlap.

    Used for Excel statements (where vendor isn't in the body): if the user
    names files like 'Ferguson_2026-03.xlsx' or 'Estrada March.xlsx', we can
    pick up the vendor from a previously-confirmed alias.

    Returns (alias_key, alias_data, score, matched_tokens_str) for the best
    confident match, or None if there's no clear winner (no matches, or a tie
    where we can't pick safely).

    Scoring: each alias-name token (>=3 chars) found in the normalized
    filename = +1. The first token (usually most distinctive - vendor's
    primary surname) earns a +1 bonus. A "clear winner" is the highest score
    AND strictly above the runner-up; otherwise we abstain."""
    aliases = load_aliases()
    if not aliases:
        return None

    # Normalize filename: strip extension, strip common date patterns,
    # collapse separators to spaces, lowercase.
    stem = Path(filename).stem
    stem = re.sub(r"\b\d{1,2}[/.\-_]\d{1,2}([/.\-_]\d{2,4})?\b", " ", stem)  # MM/DD or MM-DD-YYYY
    stem = re.sub(r"\b\d{4}[/.\-_]?\d{1,2}([/.\-_]?\d{1,2})?\b", " ", stem)  # YYYY-MM-DD
    stem = re.sub(r"[_\-.]+", " ", stem)
    fn_norm = stem.lower()

    scored: List[Tuple[int, str, dict, List[str]]] = []
    for key, data in aliases.items():
        # Prefer the QBO display name (canonical) over the PDF-extracted name.
        name = (data.get("qbo_name") or data.get("pdf_name") or "").lower()
        if not name:
            continue
        # Tokens >=3 chars from the alias name. Skip corporate-suffix noise
        # so 'LLC'/'INC'/'CO' don't accidentally match on filenames like
        # 'Inc_2026.xlsx' (which would never be a real vendor file).
        tokens = [t for t in re.split(r"[^a-z0-9]+", name)
                  if len(t) >= 3 and t not in {"llc", "inc", "corp", "company", "ltd", "the"}]
        if not tokens:
            continue
        matched = [t for t in tokens if t in fn_norm]
        if not matched:
            continue
        score = len(matched)
        if tokens[0] in fn_norm:
            score += 1  # first-token bonus (usually most distinctive)
        scored.append((score, key, data, matched))

    if not scored:
        return None
    scored.sort(reverse=True, key=lambda x: (x[0], -len(x[1])))  # higher score wins; tie → prefer longer key
    best_score, best_key, best_data, best_matched = scored[0]
    # Clear winner only if strictly higher than runner-up (avoid ambiguous matches).
    if len(scored) == 1 or scored[0][0] > scored[1][0]:
        return (best_key, best_data, best_score, ", ".join(best_matched))
    return None


def validate_cached_vendor(access: str, cid: str, qbo_id: str) -> Optional[str]:
    """Cheap SELECT-by-Id check. Returns current DisplayName if vendor still exists, else None."""
    try:
        rows = query(access, cid,
                     f"SELECT Id, DisplayName FROM Vendor WHERE Id = '{qbo_id}'"
                     ).get("QueryResponse", {}).get("Vendor", [])
        if rows:
            return rows[0].get("DisplayName", "")
    except RuntimeError:
        return None
    return None


def find_vendor_id_cached(access: str, cid: str, pdf_vendor: str,
                          override: str = "", strict: bool = True) -> Tuple[str, str, bool]:
    """Resolve vendor → (id, current_display_name, from_cache).
    - If --vendor override is set, use it directly (no cache lookup).
    - Else check alias cache; validate via QBO; on hit return (id, name, True).
    - On miss/invalid, fall back to LIKE search and return (id, name, False).
    Caller is responsible for saving the alias after user confirmation.
    """
    if override:
        return (*find_vendor_id(access, cid, override, strict=strict), False)
    aliases = load_aliases()
    cached = aliases.get(_alias_key(pdf_vendor))
    cached_id = cached.get("qbo_id") if cached else None
    if cached_id:
        current_name = validate_cached_vendor(access, cid, cached_id)
        if current_name:
            return cached_id, current_name, True
        _warn(f"cached vendor id {cached_id} ({cached.get('qbo_name')}) no longer in QBO - re-resolving.")
    return (*find_vendor_id(access, cid, pdf_vendor, strict=strict), False)


def find_vendor_id(access: str, cid: str, vendor_name_hint: str,
                   strict: bool = True) -> Tuple[str, str]:
    """Returns (vendor_id, vendor_display_name), or ("", "") when strict=False and
    nothing resolves (so an unattended batch can SKIP one vendor instead of
    aborting the whole run - a bare no-match sys.exit is a SystemExit that slips
    past run_inbox's `except Exception`). strict=True keeps the fatal, helpful
    message for interactive single-file use.
    Tries progressively-relaxed LIKE searches against QBO Vendor.DisplayName:
      1. first 2 whitespace tokens with internal punctuation preserved
         ('Post-Tension Services' - handles hyphens/ampersands cleanly)
      2. first 1 token (handles vendors whose PDF shows only one word)
      3. hyphen-stripped variants of #1 and #2 (handles QBO names spelled
         without the hyphen - e.g., PDF 'Post-Tension' vs QBO 'Post Tension')
    Trailing punctuation like commas/periods is stripped from each token
    so 'Ready Cable, Inc' → tokens ['Ready', 'Cable', 'Inc']."""
    if not vendor_name_hint:
        if not strict:
            return "", ""
        sys.exit("✗ could not identify vendor from PDF. Pass --vendor explicitly.")
    raw_tokens = vendor_name_hint.split()
    tokens = [t.rstrip(",.;:") for t in raw_tokens if t.rstrip(",.;:")]
    if not tokens:
        if not strict:
            return "", ""
        sys.exit(f"✗ vendor hint '{vendor_name_hint}' contains no usable tokens.")

    # Build the ordered list of needles to try. Dedup while preserving order.
    attempts: List[Tuple[str, str]] = []
    seen: set = set()
    def _add(label: str, needle: str) -> None:
        n = needle.strip()
        if n and n not in seen:
            attempts.append((label, n))
            seen.add(n)

    # Full precise name first - a parser that identifies the exact entity
    # (e.g. 'Bobcat of North Texas') must beat the first-2-words reduction,
    # since 'Bobcat of' alone collides with 'Bobcat of Midland'. Only worth a
    # distinct attempt at 3+ tokens; at 1-2 tokens it equals the reductions
    # below and dedup skips it.
    if len(tokens) >= 3:
        _add("full name", " ".join(tokens))
    if len(tokens) >= 2:
        _add("first 2 words", f"{tokens[0]} {tokens[1]}")
    _add("first word", tokens[0])
    # Hyphen-stripped fallbacks - only if there's actually a hyphen
    if any("-" in t for t in tokens[:2]):
        if len(tokens) >= 2:
            _add("first 2 words (no hyphens)",
                 f"{tokens[0]} {tokens[1]}".replace("-", " "))
        _add("first word (no hyphens)", tokens[0].replace("-", " "))

    found_rows: List[dict] = []
    used_label = used_needle = ""
    for label, needle in attempts:
        safe = needle.replace("'", "''")
        rows = query(access, cid,
                     f"SELECT Id, DisplayName FROM Vendor WHERE DisplayName LIKE '%{safe}%'"
                     ).get("QueryResponse", {}).get("Vendor", [])
        if rows:
            found_rows, used_label, used_needle = rows, label, needle
            break

    if not found_rows:
        tried = ", ".join(repr(n) for _, n in attempts)
        if not strict:
            _warn(f"no QBO vendor matches. Tried: {tried}.")
            return "", ""
        sys.exit(f"✗ no QBO vendor matches. Tried: {tried}. "
                 "Pass --vendor with exact display name.")
    if len(found_rows) > 1:
        names = [r["DisplayName"] for r in found_rows]
        exact = [r for r in found_rows if r["DisplayName"].strip().upper() == used_needle.strip().upper()]
        if len(exact) != 1:
            # Never guess between several vendors (safety net, 10/07/2026: 'GONZALEZ'
            # matched seven and the first was a person, not the batch plant).
            msg = (f"{len(names)} QBO vendors match {used_needle!r}: {names}. Not guessing - "
                   "pass --vendor with the exact name (it is remembered after that).")
            if not strict:
                _warn(msg)
                return "", ""
            sys.exit(f"✗ {msg}")
        found_rows = exact
    if used_label != "first 2 words":
        print(f"  (matched via fallback: {used_label} → {used_needle!r})")
    return found_rows[0]["Id"], found_rows[0]["DisplayName"]


def _bills_to_objs(raw: List[dict]) -> List[QboBill]:
    out: List[QboBill] = []
    for b in raw:
        # QBO UI "Memo" on a Bill maps to API `PrivateNote`. A few bills
        # have a top-level `Memo` instead - accept either, prefer PrivateNote.
        memo_value = (str(b.get("PrivateNote") or "").strip()
                      or str(b.get("Memo") or "").strip())
        out.append(QboBill(
            bill_id=b["Id"],
            doc_number=str(b.get("DocNumber", "")).strip(),
            txn_date=b.get("TxnDate", ""),
            open_balance=float(b.get("Balance", 0)),
            total_amount=float(b.get("TotalAmt", 0)),
            memo=memo_value,
            created=str((b.get("MetaData") or {}).get("CreateTime") or "")[:10],
        ))
    return out


def get_open_bills_for_vendor(access: str, cid: str, vendor_id: str) -> List[QboBill]:
    """Pulls ALL open Bills (Balance > 0) regardless of past-due status."""
    raw = query_all(access, cid, "Bill",
                    where=f"VendorRef = '{vendor_id}' AND Balance > '0'")
    return _bills_to_objs(raw)


def get_recently_paid_bills_for_vendor(access: str, cid: str, vendor_id: str,
                                       since_date: str) -> List[QboBill]:
    """Pulls fully-paid Bills (Balance = 0) with TxnDate >= since_date.
    Used to catch vendor-lag: bills you paid before the statement was printed."""
    where = (f"VendorRef = '{vendor_id}' AND Balance = '0' "
             f"AND TxnDate >= '{since_date}'")
    raw = query_all(access, cid, "Bill", where=where)
    return _bills_to_objs(raw)

# ───────────────────────── matching ─────────────────────────

def _build_row(sl: StmtLine, bill: Optional[QboBill],
               paid_bill: Optional[QboBill] = None) -> ReconRow:
    """Pure classification for one statement line.
    Precedence: open Bill (MATCHED/TAX/MISMATCH) → paid Bill (LIKELY_VENDOR_LAG) → MISSING_IN_QBO."""
    if bill is None:
        if paid_bill is not None:
            return ReconRow(
                "LIKELY_VENDOR_LAG", sl.date, sl.ref,
                sl.amount, paid_bill.open_balance, sl.po, sl.address,
                f"Paid in QBO (bill dated {paid_bill.txn_date}, BillId={paid_bill.bill_id}). "
                "Vendor still shows it open - verify the check cleared, or send proof of payment.",
                stmt_ref=sl.ref, qbo_ref=paid_bill.doc_number,
                stmt_date=sl.date, qbo_date=paid_bill.txn_date,
                qbo_memo=paid_bill.memo, qbo_bill_id=paid_bill.bill_id,
            )
        return ReconRow(
            "MISSING_IN_QBO", sl.date, sl.ref,
            sl.amount, 0.0, sl.po, sl.address, "ENTER BILL IN QBO",
            stmt_ref=sl.ref, qbo_ref="",
            stmt_date=sl.date, qbo_date="",
        )
    diff = round(sl.amount - bill.open_balance, 2)
    common = dict(po=sl.po, address=sl.address,
                  stmt_ref=sl.ref, qbo_ref=bill.doc_number,
                  stmt_date=sl.date, qbo_date=bill.txn_date,
                  qbo_memo=bill.memo, qbo_bill_id=bill.bill_id)
    if abs(diff) < EXACT_TOLERANCE:
        return ReconRow("MATCHED", sl.date, sl.ref, sl.amount,
                        bill.open_balance, **common)
    expected_tax = round(bill.open_balance * TX_SALES_TAX, 2)
    if abs(diff - expected_tax) <= TAX_TOLERANCE:
        return ReconRow("VENDOR_TAX_VIOLATION", sl.date, sl.ref,
                        sl.amount, bill.open_balance,
                        notes="Stmt = QBO × 1.0825 (TX sales tax). Per agreement: NO TAX. Vendor action.",
                        **common)
    return ReconRow("CLERK_AMOUNT_MISMATCH", sl.date, sl.ref,
                    sl.amount, bill.open_balance,
                    notes=f"Diff = ${diff:,.2f}. Investigate.", **common)


def _ref_key(ref: str) -> str:
    """A Ref # as compared: case and spacing never decide a match (QBO '16018k' is the
    statement's '16018K', CMC 10/09)."""
    return re.sub(r"\s+", "", ref or "").upper()


def _suffixed_bill(ref: str, pool: List[QboBill], taken: set) -> Optional[QboBill]:
    """The ONE bill whose Ref # is the statement's number plus a clerk's suffix after a
    separator ('401417-CC FEE' for statement 401417, Cowtown 10/09). None when no bill or
    several carry it, when the number is too short to trust, or when the suffixed Ref # is
    itself a line on the statement (`taken`)."""
    key = _ref_key(ref)
    if len(key) < 4:
        return None
    pat = re.compile(re.escape(key) + r"[\s\-_/#(.]")
    hits = [b for b in pool if b.doc_number and _ref_key(b.doc_number) not in taken
            and _ref_key(b.doc_number) != key and pat.match(b.doc_number.strip().upper())]
    return hits[0] if len(hits) == 1 else None


def reconcile_iter(lines: List[StmtLine],
                   bills: List[QboBill],
                   paid_bills: Optional[List[QboBill]] = None,
                   stmt_date: str = "",
                   ) -> Iterator[Tuple[int, int, ReconRow]]:
    """Match by Ref# = DocNumber. Yields (current_index, total, row) per result.

    Total is computed up front: len(lines) + #unmatched-in-QBO bills (i.e., the
    MISSING_ON_STATEMENT rows that will be emitted after the main loop).

    `stmt_date` is the statement's "as of" date (YYYY-MM-DD). When provided,
    QBO bills with a txn_date AFTER stmt_date are excluded from MISSING_ON_STATEMENT
    because they couldn't possibly have been on a statement that hadn't been
    printed yet. Without this filter, every fresh bill in QBO would false-flag.

    Reliability warnings are printed before iteration starts:
      • Duplicate DocNumber across QBO bills (silent last-wins).
      • Duplicate Ref# on the statement (vendor reused an invoice number).
      • Bills with empty DocNumber in QBO (cannot be matched).
    """
    # Detect duplicate DocNumbers in QBO
    seen_docs: Dict[str, int] = {}
    no_doc_bills: List[QboBill] = []
    for b in bills:
        if not b.doc_number:
            no_doc_bills.append(b)
            continue
        k = _ref_key(b.doc_number)
        seen_docs[k] = seen_docs.get(k, 0) + 1
    dup_qbo = [d for d, n in seen_docs.items() if n > 1]
    if dup_qbo:
        _warn(f"QBO has duplicate DocNumber(s): {dup_qbo} - last-wins; review manually.")
    if no_doc_bills:
        _warn(f"{len(no_doc_bills)} open QBO bill(s) have empty DocNumber - cannot match by Ref#.")

    # Non-invoice statement rows are NOT bills to match against QBO: a Balance
    # forward (ref == "") lumps prior-period open items into one opening amount,
    # and payment/credit rows (amount < 0) are money the vendor applied, not bills
    # to enter. Matching either produces bogus MISSING_IN_QBO rows (CowTown 09-01:
    # a six-figure balance-forward + 18 payment rows all false-flagged "enter in QBO").
    # They stay in the tie-out (process_pdf's line_sum covers them) but are not
    # reconciled line-by-line here. Only real invoice rows are matched.
    invoice_lines = [l for l in lines if l.ref and l.amount >= 0]

    # Detect duplicate Ref# on the statement (invoice rows only - the shared
    # payment check-# would otherwise false-trigger this warning).
    seen_stmt: Dict[str, int] = {}
    for l in invoice_lines:
        seen_stmt[l.ref] = seen_stmt.get(l.ref, 0) + 1
    dup_stmt = [r for r, n in seen_stmt.items() if n > 1]
    if dup_stmt:
        _warn(f"Statement has duplicate Ref#(s): {dup_stmt} - both lines will match the same QBO bill.")

    by_doc = {_ref_key(b.doc_number): b for b in bills if b.doc_number}
    paid_by_doc = {_ref_key(b.doc_number): b for b in (paid_bills or []) if b.doc_number}
    stmt_refs = {_ref_key(l.ref) for l in invoice_lines}

    # Pair every statement line first (exact Ref #, then a unique clerk-suffixed one) so
    # a suffix-paired open bill is not also reported missing from the statement.
    pairs: List[Tuple[StmtLine, Optional[QboBill], Optional[QboBill]]] = []
    for sl in invoice_lines:
        k = _ref_key(sl.ref)
        open_match = by_doc.get(k)
        paid_match = paid_by_doc.get(k) if open_match is None else None
        if open_match is None and paid_match is None:
            open_match = _suffixed_bill(sl.ref, bills, stmt_refs)
            if open_match is None:
                paid_match = _suffixed_bill(sl.ref, paid_bills or [], stmt_refs)
        pairs.append((sl, open_match, paid_match))
    paired_ids = {id(o) for _, o, _p in pairs if o is not None}

    # Balance-forward date: QBO bills dated on/before it are folded into the
    # forward lump, so they are covered by the statement, not "missing" from it.
    bf_line = next((l for l in lines if l.ref == ""), None)
    bf_date = bf_line.date if bf_line else ""

    # MISSING_ON_STATEMENT: open in QBO, no matching ref on stmt, dated on/before
    # the statement's as-of date, AND after the balance-forward date. Post-stmt
    # bills couldn't have been on a statement that wasn't printed yet; pre-forward
    # bills are already carried in the balance-forward lump.
    def _on_or_before_stmt(b: QboBill) -> bool:
        if not stmt_date or not b.txn_date:
            return True   # no cutoff info → don't filter (preserve old behavior)
        return b.txn_date <= stmt_date

    def _covered_by_forward(b: QboBill) -> bool:
        return bool(bf_date) and bool(b.txn_date) and b.txn_date <= bf_date

    missing_on_stmt: List[QboBill] = []
    post_stmt_excluded = 0
    bf_covered: List[QboBill] = []
    for b in bills:
        if not b.doc_number or _ref_key(b.doc_number) in stmt_refs or id(b) in paired_ids:
            continue
        if not _on_or_before_stmt(b):
            post_stmt_excluded += 1
            continue
        if _covered_by_forward(b):
            bf_covered.append(b)
            continue
        missing_on_stmt.append(b)
    if post_stmt_excluded:
        _warn(f"excluded {post_stmt_excluded} QBO bill(s) dated AFTER stmt {stmt_date} "
              "from MISSING_ON_STATEMENT (can't be missing - statement is older).")
    if bf_covered:
        _warn(f"{len(bf_covered)} QBO bill(s) dated on/before the balance-forward date "
              f"{bf_date} (${sum(b.open_balance for b in bf_covered):,.2f}) are covered by "
              "the balance forward - excluded from MISSING_ON_STATEMENT.")

    no_doc_bills_filtered = [b for b in no_doc_bills
                             if _on_or_before_stmt(b) and not _covered_by_forward(b)]
    total = len(invoice_lines) + len(missing_on_stmt) + len(no_doc_bills_filtered)

    idx = 0
    for sl, open_match, paid_match in pairs:
        idx += 1
        yield idx, total, _build_row(sl, open_match, paid_match)

    # MISSING_ON_STATEMENT
    for b in missing_on_stmt:
        idx += 1
        yield idx, total, ReconRow(
            "MISSING_ON_STATEMENT", b.txn_date, b.doc_number,
            0.0, b.open_balance, "", "",
            "Open in QBO but not on statement. Vendor may have credited/applied payment.",
            stmt_ref="", qbo_ref=b.doc_number,
            stmt_date="", qbo_date=b.txn_date,
            qbo_memo=b.memo, qbo_bill_id=b.bill_id,
        )
    for b in no_doc_bills_filtered:
        idx += 1
        yield idx, total, ReconRow(
            "MISSING_ON_STATEMENT", b.txn_date, f"(no Ref#, BillId={b.bill_id})",
            0.0, b.open_balance, "", "",
            "QBO bill has no DocNumber - cannot match by Ref#. Add Ref # in QBO.",
            stmt_ref="", qbo_ref=f"(no Ref#, BillId={b.bill_id})",
            stmt_date="", qbo_date=b.txn_date,
            qbo_memo=b.memo, qbo_bill_id=b.bill_id,
        )


def reconcile(lines: List[StmtLine], bills: List[QboBill],
              paid_bills: Optional[List[QboBill]] = None,
              stmt_date: str = "") -> List[ReconRow]:
    """Backwards-compatible non-streaming wrapper."""
    return [row for _, _, row in reconcile_iter(lines, bills, paid_bills, stmt_date)]

# ───────────────────────── Excel writer ─────────────────────────

def _render_statement_pages(src: Path) -> List[Path]:
    """Render a statement file to PNG page images in a fresh temp dir.
    PDF → one PNG per page (via pypdfium2, ~150 DPI); image → one normalized
    PNG; Excel/other → [] (nothing to rasterize). Returns the list of PNG paths;
    caller owns the temp dir (paths[0].parent) and must clean it up after save."""
    ext = src.suffix.lower()
    tmpdir = Path(tempfile.mkdtemp(prefix="stmt_pages_"))
    imgs: List[Path] = []
    if ext == ".pdf":
        import pypdfium2 as pdfium
        pdf = pdfium.PdfDocument(str(src))
        try:
            n = min(len(pdf), STATEMENT_EMBED_MAX_PAGES)
            if len(pdf) > STATEMENT_EMBED_MAX_PAGES:
                _warn(f"statement has {len(pdf)} pages; embedding first "
                      f"{STATEMENT_EMBED_MAX_PAGES} only.")
            for i in range(n):
                page = pdf[i]
                pil = page.render(scale=150 / 72).to_pil()
                p = tmpdir / f"page_{i + 1:02d}.png"
                pil.save(p)
                imgs.append(p)
        finally:
            pdf.close()
    elif ext in (".png", ".jpg", ".jpeg", ".heic", ".heif"):
        from PIL import Image as PILImage
        if ext in (".heic", ".heif"):
            try:
                import pillow_heif
                pillow_heif.register_heif_opener()
            except Exception:
                pass
        with PILImage.open(src) as im:
            p = tmpdir / "page_01.png"
            im.convert("RGB").save(p)
            imgs.append(p)
    return imgs


def _embed_statement_tab(wb, src: Path) -> Optional[Path]:
    """Add a 'Statement' sheet with the source statement rendered as stacked
    page images. Returns the temp dir to clean up after wb.save (or None)."""
    from openpyxl.drawing.image import Image as XLImage
    from PIL import Image as PILImage
    pages = _render_statement_pages(src)
    if not pages:
        return None
    ws = wb.create_sheet("Statement")
    ws.sheet_view.showGridLines = False
    ws["A1"] = f"Source statement: {src.name}"
    ws["A1"].font = LABEL_FONT
    ws["A2"] = "Embedded image - travels inside this workbook."
    ws["A2"].font = Font(italic=True, name="Arial", size=10, color="808080")
    anchor_row = 4
    for p in pages:
        with PILImage.open(p) as im:
            w, h = im.size
        img = XLImage(str(p))
        if w > STATEMENT_EMBED_MAX_WIDTH:
            img.width = STATEMENT_EMBED_MAX_WIDTH
            img.height = int(h * (STATEMENT_EMBED_MAX_WIDTH / w))
        else:
            img.width, img.height = w, h
        ws.add_image(img, f"A{anchor_row}")
        anchor_row += max(2, int(img.height / 18)) + 2   # ~18px per sheet row + gap
    return pages[0].parent


def write_excel(out_path: Path, vendor: str, stmt_date: str, stmt_total: float,
                qbo_bills: List[QboBill], rows: List[ReconRow],
                statement_src: Optional[Path] = None,
                line_sum: Optional[float] = None,
                tieout_ok: bool = True,
                bf_amount: float = 0.0,
                payments_total: float = 0.0,
                stmt_lines: Optional[list] = None,
                notes: Optional[dict] = None,
                disagree: bool = False,
                markups: Optional[list] = None) -> None:
    """The vendor's ONE reconciliation workbook. `notes` = the clerk's saved notes
    (statement_set.note_key -> {note}); written into her own 'Notes' column so a
    re-run never loses them. `disagree` = the parse is fine, the vendor's statements
    differ. `markups` add the 'Statement (marked)' sheet. Which statements were
    used and what changed since the last one live on the Notion page (owner 10/08:
    no Statements / Changes sheets)."""
    notes = notes or {}
    # QBO open as-of stmt_date: exclude post-statement bills from the displayed
    # total so the reconciliation math lines up with the statement snapshot.
    def _as_of(b: QboBill) -> bool:
        return (not stmt_date) or (not b.txn_date) or (b.txn_date <= stmt_date)
    qbo_bills_asof = [b for b in qbo_bills if _as_of(b)]
    post_stmt_bills = [b for b in qbo_bills if not _as_of(b)]
    qbo_total = round(sum(b.open_balance for b in qbo_bills_asof), 2)
    post_stmt_total = round(sum(b.open_balance for b in post_stmt_bills), 2)
    # Split MATCHED by approval status so unapproved-but-matched bills don't
    # hide inside the "clean" pile - they still need PM signoff before payment.
    by_id = {b.bill_id: b for b in qbo_bills}
    state = {id(r): _approval(r, by_id) for r in rows if r.category == "MATCHED"}
    cats = {
        "MATCHED_APPROVED":      [r for r in rows if state.get(id(r)) == bill_approval.APPROVAL_YES],
        "MATCHED_CHECK_QBO":     [r for r in rows if state.get(id(r)) == bill_approval.APPROVAL_CHECK],
        "MATCHED_NOT_APPROVED":  [r for r in rows if state.get(id(r)) == bill_approval.APPROVAL_NO],
        "VENDOR_TAX_VIOLATION":  [r for r in rows if r.category == "VENDOR_TAX_VIOLATION"],
        "CLERK_AMOUNT_MISMATCH": [r for r in rows if r.category == "CLERK_AMOUNT_MISMATCH"],
        "LIKELY_VENDOR_LAG":     [r for r in rows if r.category == "LIKELY_VENDOR_LAG"],
        "MISSING_IN_QBO":        [r for r in rows if r.category == "MISSING_IN_QBO"],
        "MISSING_ON_STATEMENT":  [r for r in rows if r.category == "MISSING_ON_STATEMENT"],
    }

    wb = Workbook()

    # ── Summary ──
    s = wb.active
    s.title = "Summary"
    s.sheet_view.showGridLines = False
    s["A1"] = f"Statement Reconciliation - {vendor}"
    s["A1"].font = TITLE_FONT
    s["A1"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=False)
    s.merge_cells("A1:M1")
    s.row_dimensions[1].height = 26
    s["A2"] = (f"Statement as of: {ss.us_date(stmt_date)}    |    "
               f"Generated: {dt.date.today():%m/%d/%Y}")
    s["A2"].font = BODY_FONT
    s["A2"].alignment = Alignment(horizontal="left", vertical="center")
    s.merge_cells("A2:M2")
    s.row_dimensions[2].height = 18

    # Tie-out: label in merged A:C, dollar value in D. Wider label area so
    # "QBO open as of YYYY-MM-DD (NN bills)" fits without truncation.
    r = 3                       # no blank row 3 - it showed as a white band under the frozen title
    s.cell(row=r, column=1, value="TIE-OUT").font = HEADER_FONT
    s.cell(row=r, column=1).fill = HEADER_FILL
    s.cell(row=r, column=1).alignment = Alignment(horizontal="left", vertical="center")
    s.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    r += 1

    def _tieout_row(row_idx: int, label: str, value, font, fill=None, formula=False):
        s.cell(row=row_idx, column=1, value=label).font = font
        s.cell(row=row_idx, column=1).alignment = Alignment(horizontal="left", vertical="center")
        s.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=4)
        c = s.cell(row=row_idx, column=5, value=value)
        c.number_format = MONEY; c.font = font; c.alignment = RIGHT
        if fill:
            for col in (1, 2, 3, 4, 5):
                s.cell(row=row_idx, column=col).fill = fill

    stmt_row = r
    _tieout_row(r, "Statement total", stmt_total, BODY_FONT); r += 1

    # Parse completeness: the sum of the statement lines we actually parsed must
    # equal the statement's own Amount Due. When it doesn't, the parse dropped or
    # duplicated lines (e.g. an unsupported multi-page layout) and the whole
    # reconciliation is unreliable - surface the gap here and banner it below.
    if line_sum is not None:
        ls_row = r
        _tieout_row(r, "Sum of parsed lines", round(line_sum, 2), BODY_FONT); r += 1
        # Decompose the parsed total when the statement carries a balance-forward
        # lump and/or payment rows, so "Sum of parsed lines" reconciles visibly to
        # the itemized invoices shown in the sections below.
        if bf_amount or payments_total:
            itemized = round(line_sum - bf_amount - payments_total, 2)
            _tieout_row(r, "   · itemized invoices (this statement)", itemized, BODY_FONT); r += 1
            if bf_amount:
                _tieout_row(r, "   · balance forward (prior, not itemized)", bf_amount, BODY_FONT); r += 1
            if payments_total:
                _tieout_row(r, "   · payments / credits on statement", payments_total, BODY_FONT); r += 1
        _tieout_row(r, "Parse gap (lines − statement total)",
                    f"=E{ls_row}-E{stmt_row}", LABEL_FONT,
                    fill=(SUBTOTAL_FILL if tieout_ok else BAD_FILL)); r += 1

    qbo_label = (f"QBO open as of {stmt_date} ({len(qbo_bills_asof)} bills)"
                 if stmt_date else f"QBO open ({len(qbo_bills_asof)} open Bills)")
    qbo_row = r
    _tieout_row(r, qbo_label, qbo_total, BODY_FONT); r += 1
    _tieout_row(r, "Reconciling difference", f"=E{stmt_row}-E{qbo_row}", LABEL_FONT, fill=SUBTOTAL_FILL); r += 1

    # Loud, unmissable banner when the parse did not tie out. The report is still
    # written (a human may want to see the partial match) but it is marked NOT
    # reliable, and the caller keeps the source out of DONE.
    if not tieout_ok and disagree:
        gap = round((line_sum or 0.0) - stmt_total, 2)
        msg = (f"⚠ THE VENDOR'S STATEMENTS DISAGREE - their open invoices add to "
               f"${(line_sum or 0.0):,.2f} but the newest Amount Due is ${stmt_total:,.2f} "
               f"(${gap:,.2f} apart). Each statement ties out on its own.")
        c = s.cell(row=r, column=1, value=msg)
        c.font = Font(bold=True, name="Arial", size=12, color="7A4A00")
        c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        c.fill = PatternFill("solid", start_color="FFE0B2")
        s.merge_cells(start_row=r, start_column=1, end_row=r, end_column=12)
        s.row_dimensions[r].height = 42
        r += 1
    elif not tieout_ok:
        gap = round((line_sum or 0.0) - stmt_total, 2)
        msg = (f"⚠ TIE-OUT FAILED - parse incomplete: parsed lines total "
               f"${(line_sum or 0.0):,.2f} but the statement's Amount Due is "
               f"${stmt_total:,.2f} (gap ${gap:,.2f}). This reconciliation is NOT "
               f"reliable - do not action it until the statement is re-parsed.")
        c = s.cell(row=r, column=1, value=msg)
        c.font = Font(bold=True, name="Arial", size=12, color="FFFFFF")
        c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        c.fill = PatternFill("solid", start_color="C62828")
        s.merge_cells(start_row=r, start_column=1, end_row=r, end_column=12)
        s.row_dimensions[r].height = 42
        r += 1

    # FYI: post-statement bills excluded from the tie-out (wrap text, tall row)
    if post_stmt_bills:
        n = len(post_stmt_bills)
        info = (f"({n} QBO bill{'s' if n!=1 else ''} dated after {stmt_date}, "
                f"totaling ${post_stmt_total:,.2f}, "
                f"{'are' if n!=1 else 'is'} excluded from this tie-out - "
                f"{'they' if n!=1 else 'it'} couldn't have been on this statement.)")
        s.cell(row=r, column=1, value=info).font = Font(italic=True, name="Arial", size=10, color="808080")
        s.cell(row=r, column=1).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        s.merge_cells(start_row=r, start_column=1, end_row=r, end_column=12)
        s.row_dimensions[r].height = 32
        r += 1

    # Bills pending approval in QBO - counted from in-scope (as-of) bills only
    unapproved = [b for b in qbo_bills_asof if not b.is_approved]
    pending = [b for b in qbo_bills_asof if b.is_approved and bill_approval.approval_state(
        b.memo, b.created, b.open_balance) == bill_approval.APPROVAL_CHECK]
    if unapproved or pending:
        parts = []
        if unapproved:
            parts.append(f"{len(unapproved)} NOT APPROVED by memo "
                         f"(${round(sum(b.open_balance for b in unapproved), 2):,.2f}) - chase the PM")
        if pending:
            parts.append(f"{len(pending)} entered since 09/16 and unpaid "
                         f"(${round(sum(b.open_balance for b in pending), 2):,.2f}) - their approval is in QBO, check there")
        msg = "⚠ Approval: " + " · ".join(parts) + "  (see the 'Approved?' column)"
        c = s.cell(row=r, column=1, value=msg)
        c.font = Font(bold=True, name="Arial", size=11, color="C62828")
        c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        c.fill = PatternFill("solid", start_color="FFE5E5")
        s.merge_cells(start_row=r, start_column=1, end_row=r, end_column=12)
        s.row_dimensions[r].height = 30
        r += 1
    r += 1

    # ── Per-category bill lists with collapsible groups (summary above) ──
    # Layout: leading ↗ bill-link column, then 9 columns side-by-side stmt vs
    # QBO with check marks, then Approved? + Note.
    #   A: ↗ (open bill in QBO)
    #   B: Stmt Ref#   C: QBO Ref#   D: ✓?
    #   E: Stmt $      F: QBO $      G: ✓?
    #   H: Stmt Date   I: QBO Date   J: ✓?
    # Group rows collapse under the category-header row. MATCHED collapsed by
    # default (boring); others expanded (actionable).
    s.sheet_properties.outlinePr.summaryBelow = False  # summary row is ABOVE detail

    LAG_FILL = PatternFill("solid", start_color="DDEBF7")
    CHECK_OK = "✓"
    CHECK_NO = "✗"
    CHECK_NA = "-"

    GREEN_FONT = Font(bold=True, name="Arial", size=10, color="2E7D32")
    RED_FONT   = Font(bold=True, name="Arial", size=10, color="C62828")
    DIM_FONT   = Font(name="Arial", size=10, color="808080")
    LINK_FONT  = Font(bold=True, name="Arial", size=11, color="0563C1", underline="single")

    CAT_DEFS = [
        ("MATCHED_APPROVED",      "✓ MATCHED - APPROVED",                       OK_FILL,   True),   # collapse by default (boring/clean)
        ("MATCHED_NOT_APPROVED",  "⚠ MATCHED but NOT APPROVED (chase PM)",      WARN_FILL, False),  # surface - needs PM signoff
        ("MATCHED_CHECK_QBO",     "? MATCHED - approval is in QBO (entered since 09/16, unpaid): check its approval there",
                                  PatternFill("solid", start_color="D9F2F7"), False),
        ("VENDOR_TAX_VIOLATION",  "⚠ VENDOR TAX VIOLATION (8.25%)",             WARN_FILL, False),
        ("CLERK_AMOUNT_MISMATCH", "⚠ CLERK AMOUNT MISMATCH",                    WARN_FILL, False),
        ("LIKELY_VENDOR_LAG",     "⊙ LIKELY VENDOR LAG (paid in QBO)",          LAG_FILL,  False),
        ("MISSING_IN_QBO",        "✗ MISSING IN QBO",                           BAD_FILL,  False),
        ("MISSING_ON_STATEMENT",  "✗ MISSING ON STATEMENT",                     BAD_FILL,  False),
    ]
    SUMMARY_NCOLS = 13

    def _section_header_text(label: str, count: int, stmt_sum: float, qbo_sum: float) -> str:
        diff = round(stmt_sum - qbo_sum, 2)
        if count == 0:
            return f"{label}  -  0 bills"
        if "TAX" in label or "MISMATCH" in label:
            return (f"{label}  -  {count} bill{'s' if count!=1 else ''}  ·  "
                    f"stmt ${stmt_sum:,.2f}  vs  QBO ${qbo_sum:,.2f}  (diff ${diff:,.2f})")
        if "MISSING ON STATEMENT" in label:
            return f"{label}  -  {count} bill{'s' if count!=1 else ''}  ·  ${qbo_sum:,.2f} open in QBO"
        if "VENDOR LAG" in label:
            return f"{label}  -  {count} bill{'s' if count!=1 else ''}  ·  ${stmt_sum:,.2f} on stmt"
        return f"{label}  -  {count} bill{'s' if count!=1 else ''}  ·  ${stmt_sum:,.2f}"

    def _write_section_header_row(row_idx: int, text: str, fill) -> None:
        s.cell(row=row_idx, column=1, value=text).font = LABEL_FONT
        s.cell(row=row_idx, column=1).alignment = LEFT
        for col in range(1, SUMMARY_NCOLS + 1):
            s.cell(row=row_idx, column=col).fill = fill
            s.cell(row=row_idx, column=col).border = BORDER
        s.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=SUMMARY_NCOLS)

    def _write_subhead_row(row_idx: int) -> None:
        headers = ["Bill", "Stmt Ref #", "QBO Ref #", "✓", "Stmt $", "QBO $", "✓",
                   "Stmt Date", "QBO Date", "✓", "Approved?", "Finding", "Notes"]
        for i, h in enumerate(headers, 1):
            c = s.cell(row=row_idx, column=i, value=h)
            c.font = Font(bold=True, name="Arial", size=10, color="404040")
            c.alignment = CENTER
            c.fill = SUBTOTAL_FILL
            c.border = BORDER

    # Approval-state cell styling
    UNAPPROVED_FILL = PatternFill("solid", start_color="FFE5E5")
    UNAPPROVED_FONT = Font(bold=True, name="Arial", size=10, color="C62828")
    APPROVED_FONT   = Font(bold=True, name="Arial", size=10, color="2E7D32")

    def _write_compare_row(row_idx: int, rr: ReconRow) -> None:
        # Ref check
        stmt_ref, qbo_ref = rr.stmt_ref, rr.qbo_ref
        ref_check = (CHECK_OK if (stmt_ref and qbo_ref and _ref_key(stmt_ref) == _ref_key(qbo_ref))
                     else CHECK_NA if (not stmt_ref or not qbo_ref)
                     else CHECK_NO)
        # Amount check (within tolerance)
        amt_check = (CHECK_OK if abs(rr.stmt_amount - rr.qbo_amount) < EXACT_TOLERANCE
                     and (rr.stmt_amount > 0 or rr.qbo_amount > 0)
                     else CHECK_NA if (rr.stmt_amount == 0 or rr.qbo_amount == 0)
                          and rr.category in ("MISSING_IN_QBO", "MISSING_ON_STATEMENT")
                     else CHECK_NO)
        # Date check
        date_check = (CHECK_OK if (rr.stmt_date and rr.qbo_date and rr.stmt_date == rr.qbo_date)
                      else CHECK_NA if (not rr.stmt_date or not rr.qbo_date)
                      else CHECK_NO)
        # Approval state: only meaningful when there's a QBO bill (MISSING_IN_QBO has none)
        if rr.category == "MISSING_IN_QBO":
            approved_text = "-"
            approved_font = DIM_FONT
            approved_fill = None
        else:
            st = _approval(rr, by_id)
            is_app = st == bill_approval.APPROVAL_YES
            approved_text = {bill_approval.APPROVAL_YES: "Yes", bill_approval.APPROVAL_NO: "Not Approved",
                             bill_approval.APPROVAL_CHECK: "Check QBO"}[st]
            approved_font = APPROVED_FONT if is_app else (
                UNAPPROVED_FONT if st == bill_approval.APPROVAL_NO else Font(bold=True, name="Arial", size=10, color="00838F"))
            approved_fill = None if is_app else (
                UNAPPROVED_FILL if st == bill_approval.APPROVAL_NO else PatternFill("solid", start_color="D9F2F7"))

        # Column 1: ↗ deep-link that opens this Bill in QBO. Blank when there is
        # no QBO bill for the row (MISSING_IN_QBO). Ref# text stays non-clickable.
        link_cell = s.cell(row=row_idx, column=1, value=("↗" if rr.qbo_bill_id else ""))
        link_cell.alignment = CENTER
        link_cell.border = BORDER
        if rr.qbo_bill_id:
            link_cell.hyperlink = QBO_BILL_URL_TEMPLATE.format(bill_id=rr.qbo_bill_id)
            link_cell.font = LINK_FONT
        else:
            link_cell.font = BODY_FONT

        cells = [
            (rr.stmt_ref, LEFT, None),
            (rr.qbo_ref, LEFT, None),
            (ref_check, CENTER, GREEN_FONT if ref_check == CHECK_OK else (RED_FONT if ref_check == CHECK_NO else DIM_FONT)),
            (rr.stmt_amount if rr.stmt_amount else None, RIGHT, None),
            (rr.qbo_amount if rr.qbo_amount else None, RIGHT, None),
            (amt_check, CENTER, GREEN_FONT if amt_check == CHECK_OK else (RED_FONT if amt_check == CHECK_NO else DIM_FONT)),
            (rr.stmt_date, CENTER, None),
            (rr.qbo_date, CENTER, None),
            (date_check, CENTER, GREEN_FONT if date_check == CHECK_OK else (RED_FONT if date_check == CHECK_NO else DIM_FONT)),
            (approved_text, CENTER, approved_font),
            (rr.notes or "", LEFT, None),
        ]
        # Compare cells occupy columns 2..12 (column 1 is the ↗ link above).
        for i, (val, align, font) in enumerate(cells, 2):
            c = xl_text(s, row_idx, i, val)          # vendor statement text: stored as text, never a formula
            c.font = font or BODY_FONT
            c.alignment = align
            c.border = BORDER
            if i in (5, 6):
                c.number_format = MONEY
            # Highlight unapproved cell (Approved? column, now col 11) with red fill
            if i == 11 and approved_fill is not None:
                c.fill = approved_fill
        # Column 13 'Notes': the clerk's own, typed on Notion (since 10/08) - saved by bill id (else invoice #) and
        # written back on every run, so a re-run never wipes what she typed.
        nc = xl_text(s, row_idx, 13, ss.note_for(notes, rr.qbo_bill_id, rr.stmt_ref or rr.qbo_ref))
        nc.font = Font(italic=True, name="Arial", size=10, color="595959")   # shown here, typed on Notion
        nc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        nc.border = BORDER

    # ── PRINT STATUS (first section) ──────────────────────────────────────
    # Was each statement invoice ever received-and-printed? (AP-03 -> AP-01).
    # OPT-IN while we test-and-see (PRINT_STATUS=1) so a normal reconcile never
    # triggers a mailbox pull until it's proven. The bill clerk works the
    # "Unprinted bills" group (open by default, with a box to mark once printed -
    # so this workbook doubles as her worklist); the invoice clerk works the QBO
    # buckets below. Best-effort: any Graph failure skips the section, never breaks
    # the reconcile.
    _ps_on = os.environ.get("PRINT_STATUS", "").strip().lower() in ("1", "true", "yes", "on")
    if _ps_on and stmt_lines:
        _ps = None
        _idx = None
        try:
            import print_status as _ps
            _idx = _ps.printed_index()
        except Exception as e:
            _warn(f"Print Status skipped ({e}).")
        if _ps is not None and _idx is None:
            _warn(f"Print Status skipped: {_ps._INDEX_CACHE.get('err', 'index unavailable')}")
        elif _idx is not None:
            _qbo_present = ("MATCHED_APPROVED", "MATCHED_CHECK_QBO", "MATCHED_NOT_APPROVED", "VENDOR_TAX_VIOLATION",
                           "CLERK_AMOUNT_MISMATCH", "LIKELY_VENDOR_LAG")
            _qbo_refs = {rr.stmt_ref for k in _qbo_present for rr in cats[k] if rr.stmt_ref}
            _prows = _ps.build_print_rows(stmt_lines, _idx,
                                          search_fn=_ps.live_search_printed, qbo_refs=_qbo_refs)
            _pc = _ps.classify_print_rows(_prows)
            _printed, _to_print = _pc["printed"], (_pc["genuine"] + _pc["suspect"] + _pc["unverified"])

            _write_section_header_row(
                r, f"PRINT STATUS  -  {len(_prows)} invoices  ·  {len(_printed)} printed  "
                   f"·  {len(_to_print)} to check", HEADER_FILL)
            r += 1

            # Agreement-gate + index-health banners (the future-proof alarms)
            _alerts = [f"⚠ index health: {h}" for h in _idx.health]
            if _pc["suspect"]:
                _alerts.append(f"⚠ {len(_pc['suspect'])} bill(s) are IN QBO but no printed email "
                               f"was found - likely a reader blind spot for this vendor's format; "
                               f"verify before treating as un-printed.")
            if _pc["unverified"]:
                _alerts.append(f"⚠ {len(_pc['unverified'])} bill(s) UNVERIFIED (mailbox search "
                               f"error) - re-run; not counted as un-printed.")
            for _msg in _alerts:
                c = s.cell(row=r, column=1, value=_msg)
                c.font = Font(bold=True, name="Arial", size=10, color="C62828")
                c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                c.fill = PatternFill("solid", start_color="FFE5E5")
                s.merge_cells(start_row=r, start_column=1, end_row=r, end_column=SUMMARY_NCOLS)
                s.row_dimensions[r].height = 28
                r += 1

            # Group 1: Printed bills check (collapsed by default - the boring pile)
            _write_section_header_row(r, f"✓ Printed bills check  -  {len(_printed)}", OK_FILL)
            r += 1
            if _printed:
                _ph = ["", "Date", "Invoice #", "", "Amount", "", "", "Email date",
                       "Email subject", "", "", ""]
                for i, h in enumerate(_ph, 1):
                    cc = s.cell(row=r, column=i, value=h)
                    cc.font = Font(bold=True, name="Arial", size=10, color="404040")
                    cc.alignment = CENTER; cc.fill = SUBTOTAL_FILL; cc.border = BORDER
                s.row_dimensions[r].outline_level = 1; s.row_dimensions[r].hidden = True
                r += 1
                for pr in _printed:
                    s.cell(row=r, column=2, value=pr.date).font = BODY_FONT
                    xl_text(s, r, 3, pr.ref).font = BODY_FONT   # from an email subject: text only
                    cc = s.cell(row=r, column=5, value=pr.amount or None)
                    cc.number_format = MONEY; cc.alignment = RIGHT; cc.font = BODY_FONT
                    s.cell(row=r, column=8, value=pr.email_date).font = BODY_FONT
                    sc = s.cell(row=r, column=9, value=pr.email_subject)
                    sc.alignment = LEFT; sc.font = BODY_FONT
                    s.merge_cells(start_row=r, start_column=9, end_row=r, end_column=SUMMARY_NCOLS)
                    for col in range(1, SUMMARY_NCOLS + 1):
                        s.cell(row=r, column=col).border = BORDER
                    s.row_dimensions[r].outline_level = 1; s.row_dimensions[r].hidden = True
                    r += 1
            r += 1

            # Group 2: Unprinted bills (OPEN by default; a box she marks once printed)
            _write_section_header_row(
                r, f"✗ Unprinted bills  -  {len(_to_print)}  ·  print, then mark the box", BAD_FILL)
            r += 1
            if _to_print:
                _uh = ["", "Date", "Invoice #", "", "Amount", "", "In QBO?", "Printed ✓",
                       "Note", "", "", ""]
                for i, h in enumerate(_uh, 1):
                    cc = s.cell(row=r, column=i, value=h)
                    cc.font = Font(bold=True, name="Arial", size=10, color="404040")
                    cc.alignment = CENTER; cc.fill = SUBTOTAL_FILL; cc.border = BORDER
                r += 1
                _INPUT_FILL = PatternFill("solid", start_color="FFF9C4")
                for pr in _to_print:
                    note = ("in QBO - verify; may already be printed" if pr.reader_suspect
                            else "unverified - mailbox search error, re-run" if pr.unverified
                            else "")
                    s.cell(row=r, column=2, value=pr.date).font = BODY_FONT
                    xl_text(s, r, 3, pr.ref).font = BODY_FONT   # from an email subject: text only
                    cc = s.cell(row=r, column=5, value=pr.amount or None)
                    cc.number_format = MONEY; cc.alignment = RIGHT; cc.font = BODY_FONT
                    inq = "Yes" if pr.in_qbo else ("No" if pr.in_qbo is False else "-")
                    s.cell(row=r, column=7, value=inq).alignment = CENTER
                    s.cell(row=r, column=7).font = BODY_FONT
                    box = s.cell(row=r, column=8, value=""); box.fill = _INPUT_FILL; box.alignment = CENTER
                    nt = s.cell(row=r, column=9, value=note); nt.alignment = LEFT
                    nt.font = DIM_FONT if note else BODY_FONT
                    s.merge_cells(start_row=r, start_column=9, end_row=r, end_column=SUMMARY_NCOLS)
                    for col in range(1, SUMMARY_NCOLS + 1):
                        s.cell(row=r, column=col).border = BORDER
                    s.row_dimensions[r].outline_level = 1; s.row_dimensions[r].hidden = False
                    r += 1
            r += 1

    for key, label, fill, collapse_by_default in CAT_DEFS:
        bucket = cats[key]
        n = len(bucket)
        if n == 0:
            continue                     # drop empty buckets entirely - no clutter
        stmt_sum = round(sum(rr.stmt_amount for rr in bucket), 2)
        qbo_sum  = round(sum(rr.qbo_amount  for rr in bucket), 2)

        _write_section_header_row(r, _section_header_text(label, n, stmt_sum, qbo_sum), fill)
        r += 1

        # Sub-header and data rows - all at outline_level=1, grouped under the header
        sub_head_row = r
        _write_subhead_row(sub_head_row)
        s.row_dimensions[sub_head_row].outline_level = 1
        s.row_dimensions[sub_head_row].hidden = collapse_by_default
        r += 1

        for rr_row in bucket:
            _write_compare_row(r, rr_row)
            s.row_dimensions[r].outline_level = 1
            s.row_dimensions[r].hidden = collapse_by_default
            r += 1

        # Spacer row (not grouped) between categories
        r += 1

    # Column widths sized for the 12-col layout (↗ link + ref/amount/date
    # triplets + Approved? + Note)
    for col, w in {"A": 5, "B": 13, "C": 13, "D": 4, "E": 13, "F": 13, "G": 4,
                   "H": 12, "I": 12, "J": 4, "K": 14, "L": 36, "M": 40}.items():
        s.column_dimensions[col].width = w

    # Freeze the top rows (title + tie-out) so per-bill rows stay scrollable
    # under a fixed header. Freeze at row 4 + col A leaves the table header
    # area visible at all times.
    s.freeze_panes = "A3"

    marked_dir = None
    if markups:
        try:
            marked_dir = _write_marked_sheet(wb, markups)
        except Exception as e:
            _warn(f"could not add the marked statement ({e}); Excel saved without it.")

    # Embed the source statement as a self-contained "Statement" tab so the
    # workbook carries the original inside it - no external link to break when
    # the file is moved (Inbox → Reconciliations → Old-Done).
    cleanup_dir = None
    if statement_src is not None:
        try:
            cleanup_dir = _embed_statement_tab(wb, statement_src)
        except Exception as e:
            _warn(f"could not embed statement image ({e}); Excel saved without it.")

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    if cleanup_dir:
        shutil.rmtree(cleanup_dir, ignore_errors=True)
    if marked_dir:
        shutil.rmtree(marked_dir, ignore_errors=True)
    # Binding guard (repo rule 5b): never hand over an xlsx that would trip
    # Excel's "we found a problem with some content" repair prompt.
    xlsx_verify.assert_clean(out_path)

MARKED_WIDTH = 1000          # px on the sheet per statement page


LEGEND_COL = 16              # P - the page pictures fill A:N (4 + 13 x 10 wide = 1000px), O is a gutter


def _write_marked_sheet(wb, markups: list) -> Optional[Path]:
    """'Statement (marked)', the 2nd sheet: one line per statement (lines found on
    the page · rows on the page the tool did NOT read), the misses listed, then every
    page with each bill row banded in its bucket's colour; the colour legend sits to
    the RIGHT of the first page (column P). The statement file is only read. Returns
    the temp dir to clean after save."""
    from openpyxl.drawing.image import Image as XLImage
    ws = wb.create_sheet("Statement (marked)", 1)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 4
    for col in "BCDEFGHIJKLMN":               # 13 x ~75px = the 1000px page picture
        ws.column_dimensions[col].width = 10
    ws.column_dimensions["O"].width = 3
    ws.column_dimensions["P"].width = 5
    # Prints one page wide (the pages are pictures; a column-split print is useless).
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.orientation = "landscape"
    ws["A1"] = "Statement - each bill row coloured by what the check found"
    ws["A1"].font = Font(bold=True, name="Arial", size=12)
    r = 3
    tmp = Path(tempfile.mkdtemp(prefix="stmt_marked_"))
    legend_row = None
    for k, (name, res) in enumerate(markups):
        ok = res.located == res.total and not res.unread_rows
        msg = (f"{name}: {res.located} of {res.total} lines found on the page"
               + (f" · {len(res.unread_rows)} row(s) on the page NOT read" if res.unread_rows else "")
               + (" · all accounted for" if ok else ""))
        xl_text(ws, r, 1, msg).font = Font(bold=True, name="Arial", size=11,
                                           color="2E7D32" if ok else "C62828")
        r += 1
        for t in res.not_found:
            xl_text(ws, r, 2, f"read but not found on the page: {t}").font = Font(name="Arial", size=10, color="C62828")
            r += 1
        for t in res.unread_rows:
            xl_text(ws, r, 2, f"on the page, not read: {t}").font = Font(name="Arial", size=10, color="C62828")
            r += 1
        r += 1
        for i, im in enumerate(res.pages):
            w, h = im.size
            if w > MARKED_WIDTH:
                im = im.resize((MARKED_WIDTH, int(h * MARKED_WIDTH / w)))
            f = tmp / f"doc{k}_p{i + 1}.png"
            # 64-colour palette: a statement page is text + a few band colours; keeps
            # the workbook small on the share (RCI 6 pages: 3.5 MB full colour).
            im.convert("P", palette=1, colors=64).save(f, optimize=True)
            img = XLImage(str(f))
            legend_row = legend_row or r
            ws.add_image(img, f"A{r}")
            r += int(img.height / 20) + 2
        r += 1
    lr = legend_row or 3
    ws.cell(row=lr, column=LEGEND_COL, value="Colour key").font = Font(bold=True, name="Arial", size=11)
    for key, (label, rgb) in sm.BUCKETS.items():
        lr += 1
        ws.cell(row=lr, column=LEGEND_COL).fill = PatternFill("solid", start_color="%02X%02X%02X" % rgb)
        ws.cell(row=lr, column=LEGEND_COL + 1, value=label).font = BODY_FONT
    lr += 1
    ws.cell(row=lr, column=LEGEND_COL).border = Border(*(Side(style="medium", color="C62828"),) * 4)
    ws.cell(row=lr, column=LEGEND_COL + 1, value="Outlined in red: looks like a bill on the statement, "
                                                  "NOT read by the tool - check it by hand"
            ).font = Font(bold=True, name="Arial", size=10, color="C62828")
    # Excel prints the range of cells that hold values and clips pictures outside it -
    # name the print area so every page picture prints (10/08/2026, seen in Excel's PDF).
    ws.print_area = f"A1:Z{r}"
    return tmp


def _plain_sheet_header(ws, row: int, headers: List[str]) -> None:
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=i, value=h)
        c.font = Font(bold=True, name="Arial", size=10)
        c.border = BORDER
        c.alignment = CENTER


# ───────────────────────── interactive helpers ─────────────────────────

def _slug(s: str) -> str:
    return re.sub(r"[^A-Za-z0-9]+", "_", s).strip("_")[:40] or "Vendor"


def _confirm(prompt: str, default_yes: bool = True, skip: bool = False) -> bool:
    """Y/N confirmation. Returns True if user proceeds."""
    if skip:
        return True
    default_str = "[Y/n]" if default_yes else "[y/N]"
    try:
        ans = input(f"{prompt} {default_str} ").strip().lower()
    except (EOFError, KeyboardInterrupt):
        print()  # newline after ^C
        return False
    if not ans:
        return default_yes
    return ans in ("y", "yes")


def _hr():
    print("─" * 60)


def _open_file(path: Path) -> None:
    """Open a file with the OS default app (macOS only - `open`)."""
    try:
        subprocess.run(["/usr/bin/open", str(path)], check=False)
    except Exception as e:
        print(f"  (could not auto-open: {e})")


def _term_link(label: str, path: Path) -> str:
    """OSC-8 terminal hyperlink. Click in macOS Terminal/iTerm2 to open in Finder.
    Falls back to plain `label  <url>` when ANSI is disabled / non-tty."""
    abs_path = path.resolve()
    url = "file://" + urllib.parse.quote(str(abs_path))
    if not _Term.enabled:
        return f"{label}  {url}"
    # OSC-8: ESC ] 8 ;; URL ST  text  ESC ] 8 ;; ST
    return f"\033]8;;{url}\033\\{label}\033]8;;\033\\"


def append_clerk_perf(rows: List["ReconRow"], vendor_name: str, stmt_date: str,
                      stmt_total: float) -> Optional[Path]:
    """Append one row to clerk_performance.csv. Focused on clerk-action metrics -
    vendor-side issues (tax violations) are EXCLUDED since the user reads those from QBO."""
    import csv as _csv

    by_cat: Dict[str, List["ReconRow"]] = {}
    for r in rows:
        by_cat.setdefault(r.category, []).append(r)

    missing  = by_cat.get("MISSING_IN_QBO", [])
    mismatch = by_cat.get("CLERK_AMOUNT_MISMATCH", [])
    matched  = by_cat.get("MATCHED", [])
    lag      = by_cat.get("LIKELY_VENDOR_LAG", [])

    # Oldest missing in days = (stmt_date) - (earliest stmt_date among missing)
    oldest_missing_days = ""
    if missing:
        try:
            stmt_dt = dt.date.fromisoformat(stmt_date)
            ages = []
            for r in missing:
                try:
                    ages.append((stmt_dt - dt.date.fromisoformat(r.date)).days)
                except ValueError:
                    pass
            if ages:
                oldest_missing_days = max(ages)
        except ValueError:
            pass

    row = {
        "reconcile_date":         dt.date.today().isoformat(),
        "vendor":                 vendor_name,
        "statement_date":         stmt_date,
        "statement_total":        f"{stmt_total:.2f}",
        "matched_count":          len(matched),
        "clerk_mismatch_count":   len(mismatch),
        "missing_in_qbo_count":   len(missing),
        "missing_in_qbo_amt":     f"{sum(r.stmt_amount for r in missing):.2f}",
        "oldest_missing_days":    oldest_missing_days,
        "likely_vendor_lag_count": len(lag),
    }

    CLERK_PERF_CSV.parent.mkdir(parents=True, exist_ok=True)
    is_new = not CLERK_PERF_CSV.exists()
    try:
        with CLERK_PERF_CSV.open("a", newline="") as f:
            w = _csv.DictWriter(f, fieldnames=list(row.keys()))
            if is_new:
                w.writeheader()
            w.writerow({k: csv_cell(v) for k, v in row.items()})
        return CLERK_PERF_CSV
    except OSError as e:
        _warn(f"could not append clerk_performance.csv ({e})")
        return None


# ───────────────────────── main ─────────────────────────

def identify_file(pdf_path: Path, args: argparse.Namespace, access: str, cid: str
                  ) -> Optional[dict]:
    """Read one statement file and resolve its QBO vendor: {vendor_id, vendor_name,
    stmt_date, amt_due, lines, template, ties_out}. None = could not identify it
    (left where it is, listed for a human). A statement whose lines do not add up
    to its Amount Due is still identified - statement_set keeps it out of the
    merge and the vendor shows Unreadable."""
    print(_Term.color(_Term.BOLD, "━" * min(60, _width())))
    print(_Term.color(_Term.BOLD, f"  READING  ·  {pdf_path.name}"))
    print(_Term.color(_Term.BOLD, "━" * min(60, _width())))

    # ── parse statement (PDF / Excel / Image) ───────────────
    ext = pdf_path.suffix.lower()
    is_excel = ext in (".xlsx", ".xls", ".xlsm")
    is_image = ext in (".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".heic", ".heif")
    file_kind = "Excel" if is_excel else ("Image (OCR)" if is_image else "PDF")
    t0 = _phase(f"Reading {file_kind} {pdf_path.name}")
    try:
        vendor_guess, stmt_date, amt_due, lines, template = parse_statement_ex(pdf_path)
    except Exception as e:
        _fail(f"could not read {pdf_path.name}: {e}")
        return None

    # Vendor fallback chain - runs for ALL file types (PDF, Excel, Image).
    # Step 2 (file body extraction) wins for PDFs where the vendor is in the
    # header/body. For Excel and OCR'd images where body extraction often
    # comes up empty, the chain falls through to filename / cell scan / prompt.
    #
    # Order rationale: explicit user signals (flag, alias, deliberate filename)
    # beat incidental file body data. Excel cells in particular often contain
    # OUR OWN company name as bill-to ("Proficient Concrete LLC") and the
    # vendor's full legal name with corp suffixes that fuzz the QBO LIKE
    # search - so filename (user-curated) beats cell scan (incidental).
    vendor_source = ""
    if args.vendor:
        # 1. Explicit --vendor flag - overrides everything
        vendor_guess = args.vendor
        vendor_source = "from --vendor flag"
    elif vendor_guess:
        # 2. File body extraction (PDF parser, OCR'd text via PDF parser).
        #    For Excel this is always empty; for PDFs this usually works.
        vendor_source = f"from {file_kind} body"

    # 3. Cached alias by filename (works for any extension)
    if not vendor_guess and not args.no_cache:
        match = find_alias_by_filename(pdf_path.name)
        if match:
            _key, data, _score, matched_tokens = match
            vendor_guess = data.get("qbo_name", "")
            vendor_source = f"from cached alias (filename tokens: {matched_tokens})"

    # 4. Cleaned filename (user-curated signal - trust it before incidental file body)
    if not vendor_guess:
        cleaned = _vendor_from_filename(pdf_path)
        if cleaned and len(cleaned) >= 3:
            vendor_guess = cleaned
            vendor_source = "from filename"

    # 5. Scan Excel cells (Excel only - last automated step before prompt)
    if not vendor_guess and is_excel:
        scanned = _scan_excel_for_vendor(pdf_path)
        if scanned:
            vendor_guess = scanned
            vendor_source = "from Excel cell"

    # 6. Interactive prompt - last resort, works for any file type
    if not vendor_guess:
        print()
        print(_Term.color(_Term.Y, "  ⚠ Couldn't auto-detect vendor (no flag, no cached alias, "
                                    "no usable filename, no body extraction)."))
        try:
            vendor_guess = input("  Vendor name to search QBO for: ").strip()
            vendor_source = "user-entered"
        except (EOFError, KeyboardInterrupt):
            print()
            _fail("vendor input required - aborted.")
            return None

    if not vendor_guess:
        _fail("no vendor identified - aborted.")
        return None

    # Always log which source the vendor came from (helps debug surprises)
    print(f"  {_Term.color(_Term.G, '✓')} Vendor candidate: "
          f"{_Term.color(_Term.C, vendor_guess)}  "
          f"{_Term.color(_Term.DIM, f'({vendor_source})')}")

    # Excel/Image - default stmt_date to today if the file didn't provide one
    if not stmt_date:
        if args.stmt_date:
            stmt_date = args.stmt_date
        elif is_excel or is_image:
            stmt_date = ss.date_from_name(pdf_path.name, pdf_path.stat().st_mtime)
            if stmt_date:
                _warn(f"No statement date in {file_kind} - using the date in its file name "
                      f"({ss.us_date(stmt_date)}).")
            else:
                _warn(f"No statement date in {file_kind} or its file name - left for a human "
                      f"(rename it with the date, e.g. '... 09-30-2026', or pass --stmt-date).")

    vendor_hint = vendor_guess

    if not lines:
        _fail(f"No statement lines parsed from this {file_kind}.")
        print("    Supported formats:")
        for label in TEMPLATE_LABELS.values():
            print(f"      • {label}")
        print("    If this is a new format, share it so a parser can be added.")
        return None

    line_sum = round(sum(l.amount for l in lines), 2)
    sum_matches = abs(line_sum - amt_due) <= 0.50
    _done(t0, f"{len(lines)} lines · sum ${line_sum:,.2f} vs Amount Due ${amt_due:,.2f}  "
              + (_Term.color(_Term.G, "✓ ties out") if sum_matches
                 else _Term.color(_Term.R, f"⚠ DOES NOT TIE (gap ${line_sum - amt_due:,.2f})")))
    if not stmt_date:
        _fail("could not identify the statement date - left for a human.")
        return None
    newest = max((l.date for l in lines if l.ref and re.match(r"\d{4}-\d\d-\d\d$", l.date or "")), default="")
    if newest and stmt_date < newest:
        _warn(f"statement date {ss.us_date(stmt_date)} is before its own invoice of "
              f"{ss.us_date(newest)} - the date was misread; it will not be used.")

    print()
    # ── vendor resolve (uses alias cache when available) ────
    print()
    t0 = _phase("Resolving vendor in QBO")
    # strict=False: an unresolvable vendor returns "" and we SKIP this one file
    # rather than sys.exit and abort the whole batch (SystemExit slips past
    # run_inbox's except). The run's summary lists it for a human.
    if args.no_cache:
        vendor_id, vendor_name = find_vendor_id(access, cid, args.vendor or vendor_hint, strict=False)
        from_cache = False
    else:
        vendor_id, vendor_name, from_cache = find_vendor_id_cached(
            access, cid, vendor_hint, override=args.vendor, strict=False)
    if not vendor_id:
        _fail(f"no QBO vendor match for '{vendor_hint}' - skipped (add an alias or pass "
              "--vendor with the exact QBO display name). Left in place.")
        return None
    cache_marker = _Term.color(_Term.Y, " ★ from saved alias") if from_cache else ""
    _done(t0, f"Matched {_Term.color(_Term.BOLD, vendor_name)}  (id={vendor_id}){cache_marker}")

    # Unattended inbox mode: only auto-process vendors already confirmed in the
    # alias cache. A first-time vendor is left in the inbox so a human runs it
    # once (which caches it) - avoids reconciling against a wrong fuzzy match
    # with nobody watching.
    if getattr(args, "inbox_cached_only", False) and not from_cache and not args.vendor:
        _fail(f"vendor '{vendor_name}' is not in the alias cache yet - skipping in unattended "
              "inbox mode. Run it once manually (statement-reconcile <file>) to confirm + cache it.")
        return None

    # ── confirm #2: vendor match - skipped on cache hit ─────
    if from_cache:
        print(f"  {_Term.color(_Term.DIM, '(skipping vendor confirmation - cached.)')}")
    else:
        print()
        _hr()
        print(_Term.color(_Term.BOLD, "  VENDOR MATCH CONFIRMATION"))
        _hr()
        print(f"  Statement says:  {vendor_hint}")
        print(f"  Matched QBO to:  {_Term.color(_Term.C, vendor_name)}  (id={vendor_id})")
        _hr()
        if not _confirm("Is this the correct vendor?", default_yes=True, skip=args.yes):
            print(_Term.color(_Term.R, "✗ aborted."))
            return None
        if not args.no_cache:
            remember_vendor(vendor_hint, vendor_id, vendor_name)
            print(f"  {_Term.color(_Term.DIM, '(remembered this vendor - future statements will skip this step)')}")

    return {"vendor_id": vendor_id, "vendor_name": vendor_name, "stmt_date": stmt_date,
            "amt_due": amt_due, "lines": lines, "template": template, "ties_out": sum_matches}


def _mark_statements(sources: list, rows: List["ReconRow"], result: dict) -> list:
    """The marked-up copy of each current statement (statement_markup): every line
    found on its page and banded in its bucket's colour. Rows on the page that look
    like a bill but were not read become 'unread' items on the Notion board - the
    'did it grab everything' check. Read-only on the statement files."""
    by_id = {b.bill_id: b for b in result.get("_bills", [])}
    cat = {}
    for r in rows:
        if r.stmt_ref:
            cat[r.stmt_ref.strip().upper()] = sm.bucket_for(r.category, _approval(r, by_id))
    out = []
    for path, lines in sources:
        t0 = _phase(f"Marking up {path.name}")
        try:
            with contextlib.redirect_stdout(io.StringIO()):
                parse_statement_ex(path)                 # for rows left out on purpose
            skipped = last_skipped()
            res = sm.mark(sm.load_pages(path), lines,
                          lambda ln: cat.get((ln.ref or "").strip().upper(), "other"), skipped)
        except Exception as e:
            _warn(f"could not mark up {path.name}: {e}")
            continue
        out.append((path.name, res))
        missed = len(res.not_found) + len(res.unread_rows)
        _done(t0, f"{res.located}/{res.total} lines found on the page"
                  + (_Term.color(_Term.R, f" · {missed} to check by hand") if missed else " · all accounted for"))
        for t in res.not_found + res.unread_rows:
            print(_Term.color(_Term.R, f"    · {t}"))
            result.setdefault("items", []).append(
                {"kind": "unread", "ref": t[:80], "line": f"{path.name}: {t}"[:180],
                 "date": "", "amount": 0.0, "bill_id": "", "job": ""})
    return out


def reconcile_vendor(args: argparse.Namespace, access: str, cid: str, result: dict, *,
                     vendor_id: str, vendor_name: str, stmt_date: str, amt_due: float,
                     lines: List[StmtLine], out: Optional[Path], extras: dict
                     ) -> Tuple[bool, Dict[str, int], bool]:
    """Reconcile ONE vendor's open list (statement_set's merge of its current
    statements) against QBO, fill `result` for the Notion board and, unless
    --dry-run, write the vendor's one Excel to `out`. Returns (ok, counts,
    tieout_ok). extras: notes, statements, changes, disagree."""
    line_sum = round(sum(l.amount for l in lines), 2)
    sum_matches = abs(line_sum - amt_due) <= 0.50
    bf_amount = round(sum(l.amount for l in lines if not l.ref), 2)
    payments_total = round(sum(l.amount for l in lines if l.ref and l.amount < 0), 2)
    print()
    # ── pull bills ──────────────────────────────────────────
    print()
    t0 = _phase("Pulling open bills from QBO")
    bills = get_open_bills_for_vendor(access, cid, vendor_id)
    # As-of-stmt totals for the displayed tie-out (excludes post-stmt bills)
    bills_asof = [b for b in bills if not stmt_date or not b.txn_date or b.txn_date <= stmt_date]
    bills_post = [b for b in bills if stmt_date and b.txn_date and b.txn_date > stmt_date]
    qbo_total_all = sum(b.open_balance for b in bills)
    qbo_total = sum(b.open_balance for b in bills_asof)
    if bills_post:
        _done(t0, f"{_Term.color(_Term.BOLD, f'{len(bills)} open bills')} total "
                  f"${qbo_total_all:,.2f}  ·  "
                  f"as of {stmt_date}: {_Term.color(_Term.BOLD, f'{len(bills_asof)} bills, ${qbo_total:,.2f}')}  "
                  f"({len(bills_post)} dated after, ${sum(b.open_balance for b in bills_post):,.2f} excluded)")
    else:
        _done(t0, f"{_Term.color(_Term.BOLD, f'{len(bills_asof)} open bills')} totaling "
                  f"{_Term.color(_Term.BOLD, f'${qbo_total:,.2f}')}")

    if abs(len(bills_asof) - len(lines)) > max(5, len(lines) // 2):
        _warn(f"As-of bill counts differ a lot: stmt has {len(lines)}, QBO has {len(bills_asof)}. "
              "Could be normal (missing entries) or wrong vendor - verify before continuing.")

    # ── recently-paid bills (vendor-lag check) ──────────────
    # Data-driven lookback: extend far enough back to cover the OLDEST stmt
    # line (minus a week of padding), but never look back LESS than the
    # standard 60-day floor.
    floor_cutoff = dt.date.fromisoformat(stmt_date) - dt.timedelta(days=LAG_LOOKBACK_DAYS)
    # Only ISO-normalized dates are safe for fromisoformat(); _norm_date()
    # passes through unparseable strings unchanged, so filter those out here
    # rather than crash the run after QBO has already been queried.
    stmt_dates = [l.date for l in lines if l.date and re.match(r"^\d{4}-\d{2}-\d{2}$", l.date)]
    if stmt_dates:
        oldest_stmt = dt.date.fromisoformat(min(stmt_dates))
        padded_oldest = oldest_stmt - dt.timedelta(days=LAG_LOOKBACK_STMT_PADDING_DAYS)
        effective_cutoff = min(floor_cutoff, padded_oldest)
        cutoff_source = ("stmt-driven" if padded_oldest < floor_cutoff
                         else f"{LAG_LOOKBACK_DAYS}-day floor")
    else:
        effective_cutoff = floor_cutoff
        cutoff_source = f"{LAG_LOOKBACK_DAYS}-day floor"
    lag_cutoff = effective_cutoff.isoformat()
    t0 = _phase(f"Pulling paid bills since {lag_cutoff} ({cutoff_source}, vendor-lag check)")
    paid_bills = get_recently_paid_bills_for_vendor(access, cid, vendor_id, lag_cutoff)
    _done(t0, f"{len(paid_bills)} paid bills in lookback window", color=_Term.B)

    # ── live reconcile stream ───────────────────────────────
    print()
    t0 = _phase(f"Reconciling {len(lines)} statement lines vs {len(bills)} open + {len(paid_bills)} paid QBO bills")
    rows: List[ReconRow] = []
    counts: Dict[str, int] = {}
    for i, total, row in reconcile_iter(lines, bills, paid_bills, stmt_date):
        rows.append(row)
        counts[row.category] = counts.get(row.category, 0) + 1
        label, color = _CAT_STYLE.get(row.category, (row.category, _Term.RESET))
        running = (
            f"{_Term.color(_Term.G, '✓'+str(counts.get('MATCHED',0)))} "
            f"{_Term.color(_Term.Y, '⚠'+str(counts.get('VENDOR_TAX_VIOLATION',0)+counts.get('CLERK_AMOUNT_MISMATCH',0)))} "
            f"{_Term.color(_Term.B, '⊙'+str(counts.get('LIKELY_VENDOR_LAG',0)))} "
            f"{_Term.color(_Term.R, '✗'+str(counts.get('MISSING_IN_QBO',0)+counts.get('MISSING_ON_STATEMENT',0)))}"
        )
        suffix = f"{_Term.color(color, label)}  Ref #{row.ref:<10} {running}"
        _bar(i, total, suffix)
    _bar_end()
    _done(t0, f"Classified {len(rows)} bill comparisons")

    # ── summary ─────────────────────────────────────────────
    print()
    _hr()
    print(_Term.color(_Term.BOLD, "  RECONCILIATION SUMMARY"))
    _hr()
    print(f"  Statement total:           ${amt_due:,.2f}")
    if bf_amount or payments_total:
        itemized = round(line_sum - bf_amount - payments_total, 2)
        print(_Term.color(_Term.DIM, f"    · itemized invoices:      ${itemized:,.2f}"))
        if bf_amount:
            print(_Term.color(_Term.DIM, f"    · balance forward:        ${bf_amount:,.2f}"))
        if payments_total:
            print(_Term.color(_Term.DIM, f"    · payments/credits:       ${payments_total:,.2f}"))
    qbo_label = f"QBO open as of {stmt_date}:" if stmt_date else "QBO open total:"
    print(f"  {qbo_label:25s}  ${qbo_total:,.2f}  ({len(bills_asof)} bills)")
    diff = amt_due - qbo_total
    diff_color = _Term.G if abs(diff) < 0.01 else (_Term.Y if abs(diff) < 100 else _Term.R)
    print(f"  Reconciling diff:          {_Term.color(diff_color, f'${diff:,.2f}')}")
    if bills_post:
        post_total = sum(b.open_balance for b in bills_post)
        print(f"  {_Term.color(_Term.DIM, f'(plus {len(bills_post)} QBO bill(s) dated after {stmt_date} totaling ${post_total:,.2f} - excluded from tie-out)')}")
    print()
    pretty: List[Tuple[str, str, str]] = [
        ("MATCHED",               "✓ Matched",                       _Term.G),
        ("VENDOR_TAX_VIOLATION",  "⚠ Vendor tax violation (8.25%)",  _Term.Y),
        ("CLERK_AMOUNT_MISMATCH", "⚠ Clerk amount mismatch",         _Term.Y),
        ("LIKELY_VENDOR_LAG",     "⊙ Likely vendor lag (paid)",      _Term.B),
        ("MISSING_IN_QBO",        "✗ Missing in QBO",                _Term.R),
        ("MISSING_ON_STATEMENT",  "✗ Missing on Statement",          _Term.R),
    ]
    for key, label, color in pretty:
        n = counts.get(key, 0)
        bucket_total = sum(r.stmt_amount or r.qbo_amount for r in rows if r.category == key)
        line = f"  {label:38s} {n:3d}   ${bucket_total:>12,.2f}"
        print(_Term.color(color, line) if n > 0 else _Term.color(_Term.DIM, line))
    _hr()

    # Per-vendor actionable summary for the Teams task card (run_inbox posts one
    # card per vendor-month). The clerk's fix-list: bills to enter, approvals to
    # chase, amount/tax mismatches, and - if print-status ran - unprinted bills.
    _by_id = {b.bill_id: b for b in bills}
    _states = [_approval(r, _by_id) for r in rows if r.category == "MATCHED"]
    open_items = {
        "To enter (missing in QBO)": counts.get("MISSING_IN_QBO", 0),
        "Not approved (chase PM)": _states.count(bill_approval.APPROVAL_NO),
        "Approval in QBO (check)": _states.count(bill_approval.APPROVAL_CHECK),
        "Amount mismatch": counts.get("CLERK_AMOUNT_MISMATCH", 0),
        "Tax violation (8.25%)": counts.get("VENDOR_TAX_VIOLATION", 0),
    }
    unprinted: list = []
    if os.environ.get("PRINT_STATUS", "").strip().lower() in ("1", "true", "yes", "on") and lines:
        try:
            import print_status as _ps
            _idx = _ps.printed_index()
            if _idx is not None:
                _qref = {r.stmt_ref for r in rows
                         if r.category in ("MATCHED", "VENDOR_TAX_VIOLATION",
                                           "CLERK_AMOUNT_MISMATCH", "LIKELY_VENDOR_LAG")
                         and r.stmt_ref}
                _pc = _ps.classify_print_rows(
                    _ps.build_print_rows(lines, _idx,
                                         search_fn=_ps.live_search_printed, qbo_refs=_qref))
                open_items["Unprinted"] = len(_pc["genuine"]) + len(_pc["suspect"])
                unprinted = _pc["genuine"] + _pc["suspect"]
        except Exception:
            pass
    result["vendor"] = vendor_name
    result["vendor_folder"] = _vendor_folder(vendor_name)
    # The Notion board's checklist: one item per bill, with the ref numbers the
    # clerk needs (bill #, date, $, job, QBO link).
    result["items"] = _board_items(rows, bills, unprinted, extras.get("notes") or {})
    if "Unprinted" not in open_items:
        result["unchecked_kinds"] = ["print"]     # print status didn't run - keep last run's
    result.update({"vendor_id": vendor_id, "as_of": stmt_date,
                   "tieout": bool(sum_matches) or bool(extras.get("disagree")),
                   "matched": counts.get("MATCHED", 0)})
    result["open_items"] = {k: v for k, v in open_items.items() if v}
    result["clean"] = (not result["open_items"]) and sum_matches

    result["_bills"] = bills
    markups = _mark_statements(extras.get("sources") or [], rows, result)
    # what the ledger's record needs (shared/statement_record, 10/09/2026): every row with its
    # bucket, the clerk's note and the job, plus the marked pages - written after the publish
    _by = {b.bill_id: b for b in bills}
    _notes = extras.get("notes") or {}
    result["_record"] = {
        "rows": [{"category": r.category, "bucket": sm.bucket_for(r.category, _approval(r, _by)),
                  "stmt_ref": r.stmt_ref, "qbo_ref": r.qbo_ref, "stmt_date": r.stmt_date, "qbo_date": r.qbo_date,
                  "stmt_amount": r.stmt_amount, "qbo_amount": r.qbo_amount, "qbo_bill_id": r.qbo_bill_id,
                  "job": (lambda m: m.group(1).upper() if m else "")(qbo_api.PROJ_RE.search(f"{r.qbo_memo} {r.po} {r.address}")),
                  "note": ss.note_for(_notes, r.qbo_bill_id, r.stmt_ref or r.qbo_ref)} for r in rows],
        "files": [f"{ss.CURRENT}/{p.name}" for p, _l in (extras.get("sources") or [])],
        "excel": f"{ss.CURRENT}/{out.name}" if out else ""}
    result["_markups"] = markups
    result.pop("_bills", None)

    if args.dry_run:
        result["action"] = "preview"          # the Notion board preview (no writes)
        print(_Term.color(_Term.DIM, "--dry-run set; no Excel written."))
        return True, counts, sum_matches

    t0 = _phase("Writing Excel report")
    write_excel(out, vendor_name, stmt_date, amt_due, bills, rows,
                statement_src=None, line_sum=line_sum, tieout_ok=sum_matches,
                bf_amount=bf_amount, payments_total=payments_total, stmt_lines=lines,
                notes=extras.get("notes"), disagree=bool(extras.get("disagree")),
                markups=markups)
    _done(t0, f"Saved to {_Term.color(_Term.C, str(out))}")
    result["action"] = "filed" if result["tieout"] else "held"
    t0 = _phase("Appending clerk-performance history")
    csv_path = append_clerk_perf(rows, vendor_name, stmt_date, amt_due)
    if csv_path:
        _done(t0, f"+1 row to {_Term.color(_Term.C, str(csv_path))}")
    return True, counts, sum_matches


def _resolve_workflow_dirs(base: Path, strict: bool = True
                           ) -> Tuple[Optional[Path], Optional[Path]]:
    """Resolve (inbox, root) under `base`. `root` is where the per-vendor folders
    live (the Vendor Statements root); `inbox` is the dump subfolder inside it.
    Tolerant of the exact inbox spelling/spacing. With strict=True (the --inbox
    sweep) exits with a clear message if the share isn't mounted or the inbox
    can't be found; with strict=False returns (None, None) so a manually-passed
    file falls back to the default output dir instead of aborting."""
    def _bail(msg: str) -> Tuple[None, None]:
        if strict:
            sys.exit(_Term.color(_Term.R, msg))
        return None, None

    if not base.exists():
        return _bail(f"✗ not found: {base}\n"
                     "  Is the Accounting share mounted? (Finder → Go → Connect to Server)")
    entries = [d for d in base.iterdir() if d.is_dir()]

    inbox = None
    for d in entries:
        if "inbox" in d.name.lower():
            inbox = d
            break
    if not inbox:
        found = ", ".join(sorted(d.name for d in entries)) or "(none)"
        return _bail(f"✗ couldn't find the Statement Inbox folder under {base}\n"
                     f"  Folders present: {found}")
    return inbox, base


def _gather_inbox_files(inbox: Path) -> List[Path]:
    """Supported statement files sitting directly in the inbox (not DONE).
    Skips hidden files, Office lock files, and temp artifacts."""
    out: List[Path] = []
    for f in sorted(inbox.iterdir()):
        if not f.is_file():
            continue
        n = f.name
        if n.startswith(".") or n.startswith("~$") or n.endswith("#"):
            continue
        if f.suffix.lower() in INBOX_SUPPORTED_EXTS:
            out.append(f)
    return out


def _unique_dest(dest_dir: Path, name: str) -> Path:
    """A non-colliding path in dest_dir for `name` (adds ' (2)', ' (3)', …)."""
    if not (dest_dir / name).exists():
        return dest_dir / name
    stem, suf = Path(name).stem, Path(name).suffix
    i = 2
    while (dest_dir / f"{stem} ({i}){suf}").exists():
        i += 1
    return dest_dir / f"{stem} ({i}){suf}"


def _vendor_folder(name: str) -> str:
    """Filesystem-safe vendor folder from the resolved QBO vendor name. macOS is
    case-insensitive, so casing differences never make a duplicate folder."""
    s = re.sub(r'[\\/:*?"<>|]+', " ", name or "Vendor")
    return re.sub(r"\s+", " ", s).strip() or "Vendor"


_LEGAL_SUFFIX = {"LLC", "LLP", "LP", "INC", "CO", "COMPANY", "CORP", "CORPORATION", "LTD"}


def _folder_key(name: str) -> str:
    """'Core Concrete Pumping LLC' == 'Core Concrete Pumping' == 'CORE CONCRETE PUMPING, L.L.C.'"""
    toks = re.sub(r"[^A-Z0-9 ]+", " ", name.upper().replace(".", "")).split()
    while toks and toks[-1] in _LEGAL_SUFFIX:
        toks.pop()
    return " ".join(toks)


def _vendor_dir(root: Path, folder: str) -> Path:
    """The vendor's existing folder - same name in any casing, else the same name
    without a legal suffix (a hand-made 'Core Concrete Pumping' for QBO's 'Core
    Concrete Pumping LLC') - or where a new one goes. Never a second folder."""
    if root.exists():
        dirs = [d for d in root.iterdir() if d.is_dir()]
        for d in dirs:
            if d.name.upper() == folder.upper():
                return d
        key = _folder_key(folder)
        same = [d for d in dirs if _folder_key(d.name) == key]
        if len(same) == 1:
            return same[0]
    return root / folder


def _only_match(name: str, only: List[str]) -> bool:
    """--only: a whole-word match on a file or vendor folder name (any case)."""
    return not only or any(re.search(rf"(?<![A-Za-z0-9]){re.escape(t)}(?![A-Za-z0-9])", name, re.I)
                           for t in only)


def _board_items(rows: List["ReconRow"], bills: List["QboBill"], unprinted: list,
                 notes: Optional[dict] = None) -> List[dict]:
    """The Notion checklist for one vendor: one item per bill, kind = the bucket
    (see notion_board.KINDS), with the clerk's Excel note when she wrote one. A
    matched bill entered since QBO's approval workflow started and still open may
    be pending approval there - the API cannot say, so it is listed as 'check
    QBO', never as approved."""
    by_id = {b.bill_id: b for b in bills}
    kind_of = {"MISSING_IN_QBO": "enter", "CLERK_AMOUNT_MISMATCH": "mismatch",
               "VENDOR_TAX_VIOLATION": "tax", "LIKELY_VENDOR_LAG": "lag",
               "MISSING_ON_STATEMENT": "unlisted"}
    out: List[dict] = []
    for r in rows:
        kind = kind_of.get(r.category)
        if r.category == "MATCHED":
            kind = {bill_approval.APPROVAL_NO: "approve",
                    bill_approval.APPROVAL_CHECK: "checkqbo"}.get(_approval(r, by_id))
        if not kind:
            continue
        m = qbo_api.PROJ_RE.search(f"{r.qbo_memo} {r.po} {r.address}")
        out.append({"kind": kind, "ref": r.ref, "date": r.date,
                    "amount": r.stmt_amount or r.qbo_amount, "bill_id": r.qbo_bill_id,
                    "job": m.group(1).upper() if m else "",
                    "clerk_note": ss.note_for(notes or {}, r.qbo_bill_id, r.stmt_ref or r.qbo_ref)})
    for pr in unprinted or []:
        out.append({"kind": "print", "ref": pr.ref, "date": pr.date,
                    "amount": pr.amount, "bill_id": "", "job": ""})
    return out


# ─────────────────── one vendor, many statements (statement_set) ───────────────────

_KIND_WORDS = {"full": "open list", "pastdue": "past-due letter", "activity": "activity statement"}
_MONTH_DIR_RE = re.compile(r"^(\d\d)-(\d{4})(\s+DONE)?$", re.I)


def _statement_file(f: Path) -> bool:
    return (f.is_file() and not ss.is_junk(f.name) and not ss.is_report(f.name)
            and f.suffix.lower() in INBOX_SUPPORTED_EXTS)


def _quiet_doc(f: Path, origin: str) -> Tuple[Optional["ss.Doc"], str]:
    """A statement already filed under the vendor: parse it (no prompts). Returns
    (doc, vendor guess). A file that does not parse is an empty, unreadable doc."""
    try:
        vg, d, a, lines, t = parse_statement_ex(f)
    except Exception as e:
        _warn(f"{f.name}: could not read ({e})")
        vg, d, a, lines, t = "", "", 0.0, [], ""
    if not d and f.suffix.lower() not in (".pdf",):
        d = ss.date_from_name(f.name, f.stat().st_mtime)
    return ss.classify(ss.Doc(path=f, as_of=d, amount_due=a, lines=lines, template=t,
                              sha=ss.sha256(f), origin=origin)), vg


def _filed(vdir: Path) -> dict:
    """Everything already filed under one vendor: statements (Current, the old
    MM-YYYY folders, History - History only by checksum), the reconciliation
    Excels, other files in old month folders, and the old month folders."""
    out = {"docs": [], "guesses": [], "reports": [], "others": [], "month_dirs": [], "locks": []}
    if not vdir.is_dir():
        return out
    places = [(vdir / ss.CURRENT, "current")]
    for d in sorted(vdir.iterdir()):
        # Any other folder is an old month folder - '09-2026', '07-2026 DONE', and
        # hand-made ones like 'SEPT 2026' (Core Concrete Pumping).
        if d.is_dir() and d.name not in (ss.CURRENT, ss.HISTORY) and not d.name.startswith("."):
            places.append((d, "legacy_done" if re.search(r"\bDONE\b", d.name, re.I) else "legacy"))
            out["month_dirs"].append(d)
    for place, origin in places:
        if not place.is_dir():
            continue
        for f in sorted(place.iterdir()):
            if f.name.startswith("~$"):
                out["locks"].append(f)
            elif f.is_file() and ss.is_report(f.name):
                out["reports"].append((f, origin))
            elif _statement_file(f):
                doc, vg = _quiet_doc(f, origin)
                out["docs"].append(doc)
                if vg:
                    out["guesses"].append(vg)
            elif f.is_file() and not ss.is_junk(f.name):
                out["others"].append((f, origin))
    hist = vdir / ss.HISTORY
    if hist.is_dir():
        for f in sorted(hist.iterdir()):
            if _statement_file(f):
                m = ss.HISTORY_RE.match(f.name)
                iso = ""
                if m and re.match(r"\d\d-\d\d-\d{4}$", m.group(1)):
                    mm, dd, yy = m.group(1).split("-")
                    iso = f"{yy}-{mm}-{dd}"
                out["docs"].append(ss.Doc(path=f, as_of=iso, amount_due=0.0, lines=[],
                                          sha=ss.sha256(f), origin="history",
                                          status=(m.group(2) if m else "History")))
    return out


def _report_as_of(name: str) -> str:
    m = re.search(r"(\d{4})-(\d{2})-(\d{2})", name)
    if m:
        return m.group(0)
    m = re.search(r"as of (\d\d)-(\d\d)-(\d{4})", name)
    return f"{m.group(3)}-{m.group(1)}-{m.group(2)}" if m else ""


def _doc_label(d: "ss.Doc", prev_clean: bool) -> str:
    """What a statement is called once it leaves Current."""
    if d.status == "duplicate":
        return "Duplicate"
    if d.origin == "legacy_done" or (d.origin == "current" and prev_clean):
        return "DONE"
    return "Replaced"


def _file_plan(vdir: Path, docs: List["ss.Doc"], filed: dict, out: Path,
               prev_clean: bool) -> List[Tuple[Path, Path]]:
    """(source, destination) moves that leave the vendor folder as Current +
    History. A statement passed by hand from outside the Inbox is never moved;
    an unreadable one stays where it is (Inbox or Current) for a human."""
    cur, hist = vdir / ss.CURRENT, vdir / ss.HISTORY
    moves: List[Tuple[Path, Path]] = []
    for d in docs:
        if d.origin in ("history", "manual"):
            continue
        dst = None
        if d.status == "current":
            dst = None if d.path.parent == cur else cur / d.path.name
        elif d.status in ("replaced", "duplicate"):
            dst = hist / ss.history_name(d.as_of, _doc_label(d, prev_clean), d.path.name)
        elif d.status == "unreadable":
            # An old file from a month folder that can't be read (a printed email,
            # a scanned page, an unsupported layout) is filed 'Not read' - the
            # newest statements decide the vendor. A NEW one (Inbox) or one already
            # in Current stays where a person sees it.
            if d.origin in ("legacy", "legacy_done"):
                when = d.as_of or ss.date_from_name(d.path.name, d.path.stat().st_mtime)
                dst = hist / ss.history_name(when, "DONE" if d.origin == "legacy_done" else "Not read",
                                             d.path.name)
        if dst is not None:
            moves.append((d.path, dst))
    for f, origin in filed["reports"]:
        if f == out:
            continue
        label = "DONE" if origin == "legacy_done" or (origin == "current" and prev_clean) else "Replaced"
        moves.append((f, hist / ss.history_name(_report_as_of(f.name), label, f.name)))
    for f, origin in filed["others"]:
        if origin == "current":
            continue
        m = _MONTH_DIR_RE.match(f.parent.name)
        label = "DONE" if origin == "legacy_done" else "Replaced"
        when = f"{m.group(1)}-{m.group(2)}" if m else re.sub(r"\s*DONE\s*", "", f.parent.name, flags=re.I)
        moves.append((f, hist / f"{when} {label} - {f.name}"))
    return moves


def _statement_rows(docs: List["ss.Doc"], prev_clean: bool) -> List[dict]:
    """Every statement the vendor has sent, newest first - for the Notion page and
    the Excel's Statements sheet."""
    rows = []
    for d in docs:
        if d.origin == "history":
            m = ss.HISTORY_RE.match(d.path.name)
            rows.append({"as_of": d.as_of, "label": d.status or "History", "kind": "",
                         "amount": None, "name": m.group(3) if m else d.path.name})
            continue
        label = {"current": "Current",
                 "unreadable": ("Not read" if d.origin in ("legacy", "legacy_done")
                                else f"Can't read - not used ({d.problem})")}.get(
            d.status, _doc_label(d, prev_clean))
        rows.append({"as_of": d.as_of, "label": label, "kind": _KIND_WORDS.get(d.kind, d.kind),
                     "amount": d.amount_due, "name": d.path.name,
                     "lines_sum": d.line_sum})
    return sorted(rows, key=lambda r: (r["as_of"], r["label"] == "Current"), reverse=True)


def _resolve_filed_vendor(access: str, cid: str, folder: str, guesses: List[str]) -> Tuple[str, str]:
    """QBO vendor for a folder with no new statement this run: its statements'
    cached alias first, then the folder name (which is the QBO name)."""
    for g in guesses[:3]:
        vid, vname, from_cache = find_vendor_id_cached(access, cid, g, strict=False)
        if vid and from_cache:
            return vid, vname
    vid, vname = find_vendor_id(access, cid, folder, strict=False)
    return vid, vname


def run_vendors(args: argparse.Namespace, access: str, cid: str,
                inbox: Optional[Path], root: Path, new_files: List[Path],
                refresh_all: bool = False, only: Optional[List[str]] = None,
                folders: Optional[set] = None) -> int:
    """The run. Identify each new file's vendor, then per vendor: gather what is
    already filed, merge every statement into one open list (statement_set),
    reconcile it against QBO once, write ONE Excel in <Vendor>/Current, and file
    the rest to History. --dry-run prints the same plan and writes nothing."""
    only = only or []
    args.no_open = True
    args.inbox_cached_only = args.yes and getattr(args, "inbox", False)
    groups: Dict[str, dict] = {}
    failed: List[Tuple[str, str]] = []
    for f in new_files:
        print()
        info = identify_file(f, args, access, cid)
        if info is None:
            failed.append((f.name, "could not identify it (see above) - left in place"))
            continue
        folder = _vendor_dir(root, _vendor_folder(info["vendor_name"])).name
        g = groups.setdefault(folder.upper(), {"folder": folder, "vendor_id": info["vendor_id"],
                                               "vendor_name": info["vendor_name"], "docs": []})
        in_inbox = bool(inbox) and f.resolve().parent == inbox.resolve()
        g["docs"].append(ss.classify(ss.Doc(
            path=f, as_of=info["stmt_date"], amount_due=info["amt_due"], lines=info["lines"],
            template=info["template"], sha=ss.sha256(f), origin="inbox" if in_inbox else "manual")))
    if refresh_all:
        for vdir in sorted(root.iterdir()):
            if (not vdir.is_dir() or vdir.name.startswith(("-", ".", "~"))
                    or (inbox and vdir.resolve() == inbox.resolve())
                    or not _only_match(vdir.name, only)
                    or (folders is not None and vdir.name.upper() not in folders)):
                continue
            groups.setdefault(vdir.name.upper(), {"folder": vdir.name, "vendor_id": "",
                                                  "vendor_name": "", "docs": []})

    results: List[dict] = []
    summary: List[str] = []
    try:
        import notion_board as _nb
        board = _nb.Board.from_env()
    except Exception:
        board = None
    for key in sorted(groups):
        g = groups[key]
        vdir = _vendor_dir(root, g["folder"])
        folder = vdir.name
        filed = _filed(vdir)
        docs = g["docs"] + filed["docs"]
        live = [d for d in docs if d.origin != "history"]
        if not g["docs"] and not any(d.lines for d in live):
            continue                       # nothing current to check (e.g. only DONE months)
        print()
        _hr()
        print(_Term.color(_Term.BOLD, f"  VENDOR  ·  {folder}  ·  {len(live)} statement(s) to sort"))
        _hr()
        vendor_id, vendor_name = g["vendor_id"], g["vendor_name"]
        if not vendor_id:
            vendor_id, vendor_name = _resolve_filed_vendor(access, cid, folder, filed["guesses"])
        if not vendor_id:
            failed.append((folder, "no QBO vendor match for the folder - left as is"))
            continue
        if filed["locks"] and not args.dry_run:
            failed.append((folder, "a file is open in Excel (" + ", ".join(
                f.name for f in filed["locks"]) + ") - close it and re-run; nothing changed"))
            continue
        state = ss.load_state(vdir)
        prev_clean = bool(state.get("clean"))
        merged, _ = ss.plan(docs)
        stmts = _statement_rows(docs, prev_clean)
        for d in sorted(live, key=lambda d: d.as_of):
            print(f"  {ss.us_date(d.as_of) or 'undated':10s}  {d.status:10s}  "
                  f"{_KIND_WORDS.get(d.kind, d.kind):18s}  ${d.amount_due:>13,.2f}  "
                  f"{_Term.color(_Term.R, f'({d.problem}) ') if d.problem and d.status != 'duplicate' else ''}"
                  f"{d.path.name}")
        result: dict = {"vendor": vendor_name, "vendor_folder": folder, "vendor_id": vendor_id,
                        "statements": stmts}
        if not merged.lines:
            newest = max(live, key=lambda d: d.as_of)
            _fail(f"{folder}: no statement ties out - nothing reconciled; left as is.")
            results.append({"action": "unreadable", "vendor_folder": folder, "as_of": newest.as_of,
                            "vendor": vendor_name, "vendor_id": vendor_id, "_vdir": str(vdir)})
            failed.append((folder, "no readable statement"))
            continue
        cur_docs = [d for d in live if d.status == "current"]
        disagree = abs(merged.gap) > 0.50 and len(cur_docs) > 1
        if disagree:
            _warn(f"the current statements disagree by ${merged.gap:,.2f}: open invoices "
                  f"${merged.line_sum:,.2f} vs the newest Amount Due ${merged.amount_due:,.2f}.")
        # The clerk's notes live on Notion (owner 10/08): a line indented under a bill.
        # Read them first; the Excel only shows them. Once, before the move, pick up
        # anything typed in the Excel's Notes column so nothing is lost.
        today = dt.date.today().strftime("%m/%d/%Y")
        notes = dict(state.get("notes") or {})
        if not state.get("notes_on_notion"):
            for f, origin in filed["reports"]:
                if origin != "legacy_done":
                    notes = ss.fold_notes(notes, ss.harvest_notes(f), today)
        if board is not None:
            try:
                for it in board.notes(folder.upper()).values():
                    k = ss.note_key(it["bill_id"], it["ref"])
                    if not k:
                        continue
                    if it["note"]:
                        if (notes.get(k) or {}).get("note") != it["note"]:
                            notes[k] = {"note": it["note"], "on": today}
                    elif state.get("notes_on_notion"):
                        notes.pop(k, None)            # she cleared it on Notion
            except Exception as e:
                _warn(f"could not read the notes on Notion ({e}) - using the last saved ones")
        changes = ss.changes_for(docs, merged, state)
        when = dt.datetime.now().strftime("%m/%d/%Y %I:%M %p")
        dates = ", ".join(sorted({ss.us_date(d.as_of) for d in cur_docs}))
        result.update({
            "gap": merged.gap if disagree else 0.0, "merged_total": merged.line_sum,
            "amount_due": merged.amount_due, "changes": changes,
            "checked_against": (f"{len(cur_docs)} statement{'s' if len(cur_docs) != 1 else ''} "
                                f"as of {dates} vs QuickBooks open + paid bills pulled {when}"),
            "excel_url": ss.folder_link(vdir / ss.CURRENT)})
        out = vdir / ss.CURRENT / ss.report_name(folder, merged.as_of)
        ok, counts, _tie = reconcile_vendor(
            args, access, cid, result, vendor_id=vendor_id, vendor_name=vendor_name,
            stmt_date=merged.as_of, amt_due=merged.amount_due, lines=merged.lines, out=out,
            extras={"notes": notes, "statements": stmts, "changes": changes, "disagree": disagree,
                    "sources": [(d.path, d.lines) for d in cur_docs]})
        result["vendor_folder"] = folder
        result["_vdir"], result["checked"] = str(vdir), when     # for the ledger's record
        if any(d.status == "unreadable" and d.origin in ("inbox", "manual", "current")
               and (not d.as_of or d.as_of >= merged.as_of) for d in live):
            result["tieout"] = False       # a new statement can't be read - can't confirm
        results.append(result)
        moves = _file_plan(vdir, docs, filed, out, prev_clean)
        verb = "would file" if args.dry_run else "filed"
        for src, dst in moves:
            print(_Term.color(_Term.DIM, f"  {verb}: {src.parent.name}/{src.name}  ->  "
                                         f"{dst.parent.name}/{dst.name}"))
        if not args.dry_run:
            for src, dst in moves:
                try:
                    dst.parent.mkdir(parents=True, exist_ok=True)
                    shutil.move(str(src), str(_unique_dest(dst.parent, dst.name)))
                except OSError as e:
                    _warn(f"could not move {src.name}: {e}")
            for md in filed["month_dirs"]:
                left = [x for x in md.iterdir() if x.name != ".DS_Store"] if md.is_dir() else []
                if md.is_dir() and not left:
                    shutil.rmtree(md, ignore_errors=True)
            ss.save_state(vdir, {
                "as_of": merged.as_of, "checked": when, "changes": changes, "notes": notes,
                "notes_on_notion": board is not None,
                "clean": bool(result.get("tieout")) and not counts.get("MISSING_IN_QBO", 0)})
        kinds = Counter("Current" if d.status == "current" else
                        ("Not read" if d.origin in ("legacy", "legacy_done") else "Can't read")
                        if d.status == "unreadable" else _doc_label(d, prev_clean)
                        for d in live)
        summary.append(f"{folder}: as of {ss.us_date(merged.as_of)} · "
                       + ", ".join(f"{n} {k}" for k, n in sorted(kinds.items()))
                       + (f" · statements disagree by ${merged.gap:,.2f}" if disagree else ""))

    synced = _publish(results, args)
    _write_records(results, synced, args)

    print()
    _hr()
    print(_Term.color(_Term.BOLD, "  SUMMARY" + ("  (dry run - nothing written, moved or posted)"
                                                  if args.dry_run else "")))
    _hr()
    for line in summary:
        print(f"  {line}")
    if not summary:
        print("  No vendor had a statement to check.")
    if failed:
        print(_Term.color(_Term.R, f"  Left for a human: {len(failed)}"))
        for name, why in failed:
            print(_Term.color(_Term.DIM, f"    · {name} - {why}"))
    print(f"\n  {_Term.color(_Term.BOLD, 'Vendor folders:')} {_term_link(str(root), root)}")
    _hr()
    return 0 if not failed else 1


def run_from_notion(args: argparse.Namespace, base: Path) -> int:
    """--from-notion: the Notion 'Re-check' box asks for a fresh run of a vendor. Polled
    every few minutes; silent and free when there is nothing to do. Never prompts:
    a locked Key Helper or an unmounted share just waits for the next poll."""
    from shared import key_broker
    stamp = dt.datetime.now().strftime("%m/%d/%Y %I:%M %p")
    if key_broker.active():
        try:
            if not key_broker.status().get("unlocked"):
                print(f"{stamp} Key Helper is locked - skipped (next poll retries)")
                return 0
        except Exception as e:
            print(f"{stamp} Key Helper not answering ({e}) - skipped")
            return 0
    if not base.exists():
        print(f"{stamp} the Accounting share is not mounted - skipped")
        return 0
    import notion_board as nb
    board = nb.Board.from_env()
    if board is None:
        print(f"{stamp} no Notion board configured")
        return 0
    keys = {k.upper() for k in board.recheck_keys()}
    if not keys:
        return 0
    print(f"{stamp} Re-check asked for: {', '.join(sorted(keys))}")
    inbox, root = _resolve_workflow_dirs(base)
    args.yes = True
    args.no_teams = True
    args.dry_run = False
    access, cid = load_credentials()
    rc = run_vendors(args, access, cid, inbox, root, [], refresh_all=True, folders=keys)
    for k in keys:                          # a vendor with nothing to check still gets unticked
        try:
            board.untick(k)
        except Exception:
            pass
    return rc


def audit_parsing(root: Path, only: Optional[List[str]] = None) -> int:
    """--audit-parsing: every statement on the share through the parser and the
    statement_set safety nets, read-only. One line per statement that can't be
    trusted (why), then a count per template. Exit 1 if any NEW statement (Inbox
    or Current) fails - old ones in History / month folders are listed for info."""
    only = only or []
    files = []
    for f in sorted(root.rglob("*")):
        rel = f.relative_to(root)
        if (rel.parts[0].startswith(("-archive", ".")) or not _statement_file(f)
                or not _only_match(" ".join(rel.parts[:2]), only)):
            continue
        files.append(f)
    _hr()
    print(_Term.color(_Term.BOLD, f"  STATEMENT RECONCILER  ·  PARSING AUDIT  ·  {len(files)} file(s)"))
    _hr()
    bad_new, bad_old, per_tpl = [], [], Counter()
    for f in files:
        rel = f.relative_to(root)
        try:
            with contextlib.redirect_stdout(io.StringIO()):
                _v, d, a, lines, tpl = parse_statement_ex(f)
        except Exception as e:
            d, a, lines, tpl = "", 0.0, [], f"error: {e}"
        if not d and f.suffix.lower() != ".pdf":
            d = ss.date_from_name(f.name, f.stat().st_mtime)
        doc = ss.classify(ss.Doc(path=f, as_of=d, amount_due=a, lines=lines, template=tpl))
        per_tpl[tpl or "(no layout)"] += 1
        if doc.problem:
            new = rel.parts[0].startswith("-") or (len(rel.parts) > 1 and rel.parts[1] == ss.CURRENT)
            (bad_new if new else bad_old).append((str(rel), tpl or "(no layout)", doc.problem))
    for title, rows, color in (("NEW - fix before the next run", bad_new, _Term.R),
                               ("old (History / month folders) - for info", bad_old, _Term.DIM)):
        print(_Term.color(_Term.BOLD, f"\n  {title}: {len(rows)}"))
        for rel, tpl, why in rows:
            print(_Term.color(color, f"    {rel}  [{tpl}]  {why}"))
    print(_Term.color(_Term.BOLD, "\n  Read by layout:"))
    for tpl, n in per_tpl.most_common():
        print(f"    {n:4d}  {tpl}")
    _hr()
    return 1 if bad_new else 0


def _publish(results: List[dict], args: argparse.Namespace) -> List[dict]:
    """Push this run to the Notion board (one live page per vendor) and post ONE
    short Teams digest linking to it. --dry-run reads the board and prints what
    would change. Best-effort: no board id / no token / no webhook each skip with
    one line; nothing here can fail the run. Returns what the board wrote (one
    {name, status, standing, url} per vendor) so the ledger's record can carry it."""
    try:
        import notion_board as nb
        recs = nb.group_results(results)
    except Exception as e:
        _warn(f"Notion board skipped: {e}")
        return []
    if not recs or (getattr(args, "no_notion", False) and not args.dry_run):
        return []
    try:
        board = nb.Board.from_env()
    except Exception as e:
        _warn(f"Notion board skipped: {e}")
        return []
    if board is None:
        print(_Term.color(_Term.DIM, "  (Notion board off: set ACB_STATEMENTS_DS_ID in machine.env)"))
        return []
    if args.dry_run:
        print()
        print(_Term.color(_Term.BOLD, f"  NOTION BOARD PREVIEW  ·  {len(recs)} vendor(s)"))
        for rec in recs:
            try:
                r = board.preview(rec)
            except Exception as e:
                _warn(f"Notion preview: {rec['vendor']}: {e}")
                continue
            buckets = " · ".join(f"{name} {r['counts'].get(kind, 0)}" for name, kind in nb.BUCKET_PROPS
                                 if r["counts"].get(kind))
            print(f"    {r['name']}  {r['action']}  {r['standing']}")
            if buckets:
                print(_Term.color(_Term.DIM, f"      {buckets}"))
        print(_Term.color(_Term.DIM, "  (preview: nothing written. Run without --dry-run to publish.)"))
        return []
    synced: List[dict] = []
    t0 = _phase(f"Updating the Notion board ({len(recs)} vendor(s))")
    for rec in recs:
        try:
            synced.append(board.sync(rec, log=lambda m: print(_Term.color(_Term.DIM, m))))
        except Exception as e:
            _warn(f"Notion: {rec['vendor']}: {e}")
    _done(t0, f"{len(synced)} vendor page(s) written")
    for s in synced:
        print(f"    {s['name']}: {s['standing']}")
    if getattr(args, "no_teams", False):
        return []
    try:
        webhook = kc.get_secret("TEAMS_STMT_WEBHOOK") or ""
    except Exception:
        webhook = ""
    if not webhook or not synced:
        return []
    from shared import teams_notify
    if teams_notify.post_statement_digest(webhook, synced, board.url()):
        print(_Term.color(_Term.DIM, "  (posted the Teams digest)"))
    return synced



def _write_records(results: List[dict], synced: Optional[List[dict]], args: argparse.Namespace) -> None:
    """The ledger's record copy of every vendor reconciled this run (shared/statement_record,
    owner 10/09/2026): `<Vendor>/.statement-record.json` + the marked pages as PNGs under
    `<Vendor>/.marked/`. Status and standing are the Notion board's words (from what it just
    wrote, else computed the same way); an unreadable run keeps the last record's lines and
    pages and only turns the status. Never on --dry-run; nothing here can fail the run."""
    if getattr(args, "dry_run", False):
        return
    try:
        import notion_board as nb
    except Exception as e:                                   # noqa: BLE001
        _warn(f"ledger record skipped: {e}")
        return
    by_name = {s.get("name"): s for s in (synced or []) if s.get("name")}
    n = 0
    for res in results:
        vdir = Path(res.get("_vdir") or "")
        if not res.get("_vdir") or not vdir.is_dir():
            continue
        try:
            if res.get("action") == "unreadable":
                old = sr.read(vdir) or {"vendor": res.get("vendor") or vdir.name, "vendor_id": str(res.get("vendor_id") or ""),
                                        "folder": res.get("vendor_folder") or vdir.name, "lines": [], "pages": [], "counts": {}}
                old.pop("_dir", None)
                when = ss.us_date(res.get("as_of") or old.get("as_of") or "")
                old.update({"status": "Unreadable", "standing": f"Statement as of {when} unreadable" if when else "Statement unreadable",
                            "tieout": False, "clean": False, "checked": dt.datetime.now().strftime("%m/%d/%Y %I:%M %p"),
                            "checked_at": dt.datetime.now().isoformat(timespec="seconds")})
                s = by_name.get(old.get("vendor"))
                if s:
                    old["status"], old["standing"], old["notion_url"] = s["status"], s["standing"], s.get("url") or old.get("notion_url", "")
                sr.write(vdir, old)
                n += 1
                continue
            if res.get("action") not in ("filed", "held"):
                continue
            recs = nb.group_results([res])
            if not recs:
                continue
            rec = recs[0]
            m = nb.merge({"items": rec["items"], "tieout": rec["tieout"], "parsed": True}, {"items": {}, "cleared": {}},
                         frozenset(rec.get("unchecked_kinds") or ()))
            status, standing = nb.status(m, rec.get("as_of", ""), rec.get("gap", 0.0) or 0.0)
            url = ""
            s = by_name.get(rec["vendor"])
            if s:
                status, standing, url = s["status"], s["standing"], s.get("url") or ""
            markups = res.get("_markups") or []
            record = sr.build(res, status=status, standing=standing, notion_url=url, markups=markups)
            record["board_counts"] = m["counts"]
            pages = [im for _name, r in markups for im in (getattr(r, "pages", None) or [])]
            sr.write(vdir, record, pages)
            n += 1
        except Exception as e:                               # noqa: BLE001
            _warn(f"ledger record for {vdir.name}: {e}")
    if n:
        print(_Term.color(_Term.DIM, f"  (ledger record written for {n} vendor(s))"))

def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument("pdf", type=Path, nargs="*",
                   help="Path(s) to statement file(s) - accepts .pdf, .xlsx, .xls, or .png/.jpg "
                        "(images go through Tesseract OCR; .xls files auto-converted via xlrd)")
    p.add_argument("--vendor", default="",
                   help="Override vendor name. REQUIRED for .xlsx files (Excel has no vendor in body). "
                        "Otherwise guessed from PDF. In batch mode applies to ALL files - use with care.")
    p.add_argument("--stmt-date", default="", metavar="YYYY-MM-DD",
                   help="Statement as-of date. REQUIRED for .xlsx files if you want anything other "
                        "than today's date. PDFs read this from the body automatically. Format: YYYY-MM-DD.")
    p.add_argument("--out", type=Path, default=None, help="Override output xlsx path "
                                                          "(single-file mode only)")
    p.add_argument("--dry-run", action="store_true", help="Print findings without writing Excel")
    p.add_argument("--yes", action="store_true", help="Skip confirmation prompts")
    p.add_argument("--no-open", action="store_true", help="Do not auto-open the Excel report")
    p.add_argument("--no-color", action="store_true", help="Disable ANSI colors / progress bar")
    p.add_argument("--list-aliases", action="store_true", help="Print saved vendor aliases and exit")
    p.add_argument("--forget-vendor", default="", metavar="NAME",
                   help="Remove a saved alias by PDF name or QBO display name, then exit")
    p.add_argument("--no-cache", action="store_true", help="Do not use or update the vendor alias cache")
    p.add_argument("--inbox", action="store_true",
                   help="Reconcile what is in the Statement Inbox: each vendor's statements are "
                        "merged with what is already filed, checked once, filed to "
                        "<Vendor>/Current and <Vendor>/History.")
    p.add_argument("--only", default="", metavar="NAME[,NAME]",
                   help="Limit the run to these vendors - a whole word of the file name or the "
                        "vendor folder, e.g. --only cowtown,rci")
    p.add_argument("--inbox-root", type=Path, default=None,
                   help="Override the inbox workflow root (default: the Synology Vendor Statements folder).")
    p.add_argument("--embed", action="store_true",
                   help="Embed the source statement's pages as a 'Statement' tab (self-contained xlsx).")
    p.add_argument("--no-embed", action="store_true",
                   help="In --inbox mode, do NOT embed the statement image (smaller files).")
    p.add_argument("--audit-print-status", action="store_true",
                   help="Self-audit: sweep every reconciled statement (inbox + DONE) and report "
                        "print-status reader coverage per statement. No QBO, no files written. "
                        "Run this after any print_status matcher change.")
    p.add_argument("--audit-parsing", action="store_true",
                   help="Self-audit: read EVERY statement on the share (Inbox, Current, History, old "
                        "month folders) and list each one whose lines don't add up to its Amount Due, "
                        "has no date, or is dated before its own invoices. No QBO, nothing written. "
                        "Run this after any parser change.")
    p.add_argument("--refresh", action="store_true",
                   help="Re-check every vendor's current statement(s) against QBO (after the clerk "
                        "fixes bills), rewriting each vendor's Excel and Notion page.")
    p.add_argument("--from-notion", action="store_true",
                   help="Re-check the vendors whose 'Re-check' box is ticked on the Notion board, then untick "
                        "them. Run every few minutes by the Mac's scheduler; does nothing (no QBO login, no "
                        "Touch ID prompt) when nothing is ticked or Key Helper is locked.")
    p.add_argument("--no-notion", action="store_true",
                   help="Do not update the Notion Vendor Statements board this run.")
    p.add_argument("--no-teams", action="store_true",
                   help="Do not post the Teams digest this run (the Notion board still updates).")
    args = p.parse_args()

    if args.no_color:
        _Term.disable()

    # ── alias-cache utility modes ───────────────────────────
    if args.list_aliases:
        aliases = load_aliases()
        if not aliases:
            print("(no saved aliases yet)")
            return 0
        print(_Term.color(_Term.BOLD, "SAVED VENDOR ALIASES"))
        for key, v in sorted(aliases.items()):
            print(f"  {v.get('pdf_name','?'):42s} → {_Term.color(_Term.C, v.get('qbo_name','?'))}  "
                  f"(id={v.get('qbo_id','?')}, saved {v.get('saved','?')})")
        return 0

    if args.forget_vendor:
        if forget_vendor(args.forget_vendor):
            print(_Term.color(_Term.G, f"✓ forgot alias for: {args.forget_vendor}"))
            return 0
        print(_Term.color(_Term.Y, f"⚠ no alias matched: {args.forget_vendor}"))
        return 1

    # ── print-status self-audit mode (no QBO, no files) ─────
    if args.audit_print_status:
        base = args.inbox_root or INBOX_ROOT
        inbox, root = _resolve_workflow_dirs(base)
        # every filed statement PDF lives under <root>/<Vendor>/<MM-YYYY>/ (not the inbox)
        pdfs = sorted(p for p in root.rglob("*.pdf") if inbox not in p.parents)
        _hr()
        print(_Term.color(_Term.BOLD, "  STATEMENT RECONCILER  ·  PRINT-STATUS SELF-AUDIT"))
        _hr()
        print(f"  Corpus: {len(pdfs)} statement PDF(s) filed under {root}\n")
        import print_status as _ps
        return _ps.audit_print_status(pdfs)

    only = [t.strip() for t in args.only.split(",") if t.strip()]
    base = args.inbox_root or INBOX_ROOT

    if args.audit_parsing:
        _inbox, root = _resolve_workflow_dirs(base)
        return audit_parsing(root, only)

    if args.from_notion:
        return run_from_notion(args, base)

    # ── refresh: re-check every vendor's current statement(s) ──
    if args.refresh:
        inbox, root = _resolve_workflow_dirs(base)
        _hr()
        print(_Term.color(_Term.BOLD, "  STATEMENT RECONCILER  ·  REFRESH"
                                      + (f"  ·  only {', '.join(only)}" if only else "")))
        _hr()
        if args.dry_run:
            print(_Term.color(_Term.DIM, "  (dry run: reads QBO + Notion; writes, moves and posts nothing.)"))
        t0 = _phase("Authenticating to QBO (Touch ID may prompt)")
        access, cid = load_credentials()
        _done(t0, "Authenticated")
        return run_vendors(args, access, cid, inbox, root, [], refresh_all=True, only=only)

    # ── inbox: what the clerk dropped ──
    if args.inbox:
        inbox, root = _resolve_workflow_dirs(base)
        files = [f for f in _gather_inbox_files(inbox) if _only_match(f.name, only)]
        _hr()
        print(_Term.color(_Term.BOLD, "  STATEMENT RECONCILER  ·  INBOX"
                                      + (f"  ·  only {', '.join(only)}" if only else "")))
        _hr()
        print(f"  Inbox:  {inbox}")
        print(f"  Files → {root}/<Vendor>/Current + History\n")
        if not files:
            print("  Nothing in the Inbox" + (" for that vendor." if only else "."))
            return 0
        print(f"  Found {len(files)} file(s): {', '.join(f.name for f in files)}")
        if args.dry_run:
            print(_Term.color(_Term.DIM, "  (dry run: reads QBO + Notion; writes, moves and posts nothing.)"))
        t0 = _phase("Authenticating to QBO (Touch ID may prompt)")
        access, cid = load_credentials()
        _done(t0, "Authenticated")
        return run_vendors(args, access, cid, inbox, root, files, only=only)

    if not args.pdf:
        p.error("Statement file path required (.pdf / .xlsx / .png) - or use --inbox / --refresh.")

    pdf_paths: List[Path] = [Path(x) for x in args.pdf]
    for pp in pdf_paths:
        if not pp.exists():
            sys.exit(_Term.color(_Term.R, f"✗ not found: {pp}"))
    if args.out:
        _warn("--out is ignored; the Excel is written to <Vendor>/Current/.")
    inbox, root = _resolve_workflow_dirs(base)
    print(_Term.color(_Term.BOLD, "━" * min(60, _width())))
    print(_Term.color(_Term.BOLD, f"  STATEMENT RECONCILER  ·  {len(pdf_paths)} FILE(S)"))
    print(_Term.color(_Term.BOLD, "━" * min(60, _width())))
    if args.dry_run:
        print(_Term.color(_Term.DIM, "  (dry run: reads QBO + Notion; writes, moves and posts nothing.)"))
    t0 = _phase("Authenticating to QBO (Touch ID may prompt)")
    access, cid = load_credentials()
    _done(t0, "Authenticated")
    return run_vendors(args, access, cid, inbox, root, pdf_paths, only=only)


if __name__ == "__main__":
    sys.exit(main())
