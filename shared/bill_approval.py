"""bill_approval.py - the ONE rule for "is this QBO bill approved?" (shared).

Used by the Bill Tracker (bill-tracker/bill_rows.py) and the statement reconciler,
so both say the same thing about the same bill (moved here 10/08/2026 when the
reconciler needed it - tools never import tools).

Bill approval moved into QBO's approval Workflow on 09/16/2026 (the user
2026-09-17). From then on AP no longer types NOT APPROVED at entry, and QBO's API
does NOT return a bill's approval status (checked 2026-09-18 at minor versions 70
and 75 on a bill sitting in the approval queue: no such field). Bills entered
before that still carry the paper-era memo tag until they are cleared.

THE APPROVAL REPORT (the user 2026-10-13: "add new process to check bills, check
this excel while also keeping in mind the last time it synced from qbo")
    QuickBooks' own custom report "Bill Approval Status ALL" is synced out of QBO
    into the Accounts Payable channel's files (SharePoint, mirrored on OneDrive;
    `paths.bill_approval_xlsx()`). It lists every bill still in the approval
    workflow - Approval status "Pending Approval" (sent to the approver, waiting)
    or "Requires Approval" (flagged, not routed yet) - grouped by Vendor then
    Class, one row PER BILL LINE (a 12-line bill appears 12 times, the bill total
    on each), and its last row is the stamp QuickBooks printed when the report
    ran: "Tuesday, October 13, 2026 | 04:19 PM GMT -05:00". That stamp is the
    "last synced from QBO" moment every answer below is good as of.

    Bills the workflow has approved are NOT on the report (the report is the
    queue), so for a bill entered under the workflow:
      on the report            -> pending approval / requires approval (the text
                                  QBO uses, so the clerk can tell "waiting on the
                                  PM" from "nobody sent it yet")
      not on it, entered on or -> approved: it went through the queue and left it
        before the stamp
      entered AFTER the stamp  -> check QBO: the report cannot know about it yet
      report older than        -> check QBO for everything the memo does not
        MAX_AGE_HOURS, or none    settle; the report is not trusted when stale

Five states:
  not approved       the memo starts NOT APPROVED (the paper-era tag)
  pending approval   on the approval report, sent to the approver
  requires approval  on the approval report, QBO flagged it and nobody routed it
  check QBO          entered on/after APPROVAL_WORKFLOW_START and still unpaid,
                     and the report cannot vouch for it (none, stale, or the bill
                     is newer than the report)
  approved           everything else: an old-process bill with no tag, a
                     new-process bill that has been paid (QBO will not release a
                     payment on a bill that is pending approval), or a
                     new-process bill the fresh report does not list
"""
from __future__ import annotations

import datetime as dt
import os
import re
from pathlib import Path
from typing import Dict, Optional, Tuple, Union

APPROVAL_WORKFLOW_START = dt.date(2026, 9, 16)
APPROVAL_NO = "not approved"
APPROVAL_PENDING = "pending approval"
APPROVAL_REQUIRES = "requires approval"
APPROVAL_CHECK = "check QBO"
APPROVAL_YES = "approved"
UNAPPROVED_STATES = frozenset({APPROVAL_NO, APPROVAL_PENDING, APPROVAL_REQUIRES})
UNAPPROVED_TAG = "NOT APPROVED"

# The Bills / Inventory sheets show the state as ONE icon so the column stays narrow (owner
# 10/14/2026): a green check = approved, a red cross = not approved / pending / requires approval,
# an amber ? = check QBO. The ledger's loader turns the icon back into the state word (icon_state),
# so every reader downstream keeps the words; the Bill List sheet keeps the words too.
ICON_YES = "\u2713"
ICON_NO = "\u2717"
ICON_CHECK = "?"
STATE_ICON = {APPROVAL_YES: ICON_YES, APPROVAL_NO: ICON_NO, APPROVAL_PENDING: ICON_NO,
              APPROVAL_REQUIRES: ICON_NO, APPROVAL_CHECK: ICON_CHECK}
ICON_STATE = {ICON_YES: APPROVAL_YES, ICON_NO: APPROVAL_NO, ICON_CHECK: APPROVAL_CHECK}


def state_icon(state: Optional[str]) -> str:
    """The one-character cell for an approval state; an unknown state shows as check QBO."""
    return STATE_ICON.get((state or "").strip(), ICON_CHECK)


def icon_state(cell: Optional[str]) -> Optional[str]:
    """The state word behind a sheet cell: an icon maps back, a word (older files, the Bill
    List) passes through unchanged."""
    text = (cell or "").strip()
    return ICON_STATE.get(text, cell)

# How old the report may be before it stops vouching for bills (the sync out of
# QuickBooks is not continuous). Override: ACB_BILL_APPROVAL_MAX_AGE_HOURS.
MAX_AGE_HOURS = float(os.environ.get("ACB_BILL_APPROVAL_MAX_AGE_HOURS", "36") or 36)

# The report's own status words -> our states.
_REPORT_STATES = {"pending approval": APPROVAL_PENDING, "requires approval": APPROVAL_REQUIRES}


def memo_approved(memo: Optional[str]) -> bool:
    """False only when the memo STARTS with 'NOT APPROVED' (any case, leading
    whitespace allowed) - 'Approved 5/1 (was Not Approved)' is approved."""
    return not (memo or "").lstrip().upper().startswith(UNAPPROVED_TAG)


def _as_date(d: Union[dt.date, dt.datetime, str, None]) -> Optional[dt.date]:
    if isinstance(d, dt.datetime):
        return d.date()
    if isinstance(d, dt.date):
        return d
    try:
        return dt.date.fromisoformat((d or "")[:10])
    except ValueError:
        return None


def vendor_key(name: Optional[str]) -> str:
    """A vendor name folded to letters + digits, upper - the report prints QuickBooks'
    display name, the tools carry VendorRef.name; punctuation and spacing differ."""
    return re.sub(r"[^A-Z0-9]", "", (name or "").upper())


def ref_key(doc: Optional[str]) -> str:
    """A bill number folded for matching: upper, no spaces (the reconciler's rule)."""
    return re.sub(r"\s+", "", str(doc or "").strip().upper())


def _money(v) -> Optional[float]:
    try:
        return round(float(v), 2)
    except (TypeError, ValueError):
        return None


class ApprovalReport:
    """The QuickBooks approval-queue report, read once per run.

    bills:  (vendor_key, ref_key) -> {state, paid, date, amount, vendor, ref}
    by_amt: (vendor_key, date, amount) -> state   (a bill entered with no bill #)
    as_of:  when QuickBooks ran the report (its trailer stamp), else the file's time
    """

    def __init__(self, path: Path, *, now: Optional[dt.datetime] = None):
        self.path = Path(path)
        self.bills: Dict[Tuple[str, str], dict] = {}
        self.by_amt: Dict[Tuple[str, dt.date, float], str] = {}
        self.as_of: Optional[dt.datetime] = None
        self.stamp_from_report = False
        self.rows = 0
        self._now = now
        self._parse()

    # ── reading ──────────────────────────────────────────────────────────
    def _parse(self) -> None:
        import openpyxl  # the tools already depend on it; kept local so the rule itself stays light

        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        try:
            ws = wb.worksheets[0]
            rows = [r for r in ws.iter_rows(values_only=True) if any(v is not None for v in r)]
        finally:
            wb.close()
        head = None
        col: Dict[str, int] = {}
        for i, r in enumerate(rows):
            cells = [str(v).strip().lower() if v is not None else "" for v in r]
            if "bill number" in cells and "approval status" in cells:
                head = i
                col = {name: cells.index(name) for name in cells if name}
                break
        if head is None:
            raise ValueError(f"{self.path.name}: no 'Bill number' / 'Approval status' header row")
        c_ref, c_date, c_amt = col["bill number"], col.get("date"), col.get("amount")
        c_paid, c_state = col.get("paid status"), col["approval status"]
        vendor: Optional[str] = None
        klass: Optional[str] = None
        for r in rows[head + 1:]:
            label = r[c_ref]
            date = _as_date(r[c_date]) if c_date is not None else None
            state_txt = str(r[c_state] or "").strip().lower()
            if state_txt in _REPORT_STATES and date:
                self.rows += 1
                state = _REPORT_STATES[state_txt]
                amt = _money(r[c_amt]) if c_amt is not None else None
                paid = str(r[c_paid] or "").strip().lower() == "paid" if c_paid is not None else False
                vk = vendor_key(vendor)
                rk = ref_key(label)
                if rk:
                    self.bills.setdefault((vk, rk), {
                        "state": state, "paid": paid, "date": date, "amount": amt,
                        "vendor": vendor or "", "ref": str(label).strip()})
                if amt is not None:
                    self.by_amt.setdefault((vk, date, amt), state)
                continue
            if label is None:
                stamp = self._stamp(r[0])
                if stamp:
                    self.as_of, self.stamp_from_report = stamp, True
                continue
            text = str(label).strip()
            if text.startswith("Total for "):
                name = text[len("Total for "):]
                if name == "--" or name == klass:
                    klass = None
                elif name == vendor:
                    vendor, klass = None, None
                continue
            if vendor is None:
                vendor = text
            else:
                klass = text
        if self.as_of is None:
            self.as_of = dt.datetime.fromtimestamp(self.path.stat().st_mtime)

    @staticmethod
    def _stamp(v) -> Optional[dt.datetime]:
        """'Tuesday, October 13, 2026 | 04:19 PM GMT -05:00' -> the local datetime QuickBooks
        printed (the zone is the company's own; the Mac and the server run on it)."""
        if not isinstance(v, str) or "|" not in v:
            return None
        left, _, right = v.partition("|")
        day = left.split(",", 1)[-1].strip()
        time = " ".join(right.strip().split()[:2])
        for fmt in ("%B %d, %Y %I:%M %p", "%B %d, %Y %H:%M"):
            try:
                return dt.datetime.strptime(f"{day} {time}", fmt)
            except ValueError:
                continue
        try:
            return dt.datetime.strptime(day, "%B %d, %Y")
        except ValueError:
            return None

    # ── answering ────────────────────────────────────────────────────────
    @property
    def age_hours(self) -> float:
        now = self._now or dt.datetime.now()
        return max(0.0, (now - self.as_of).total_seconds() / 3600) if self.as_of else float("inf")

    @property
    def fresh(self) -> bool:
        return self.age_hours <= MAX_AGE_HOURS

    def lookup(self, vendor: Optional[str], docnum: Optional[str],
               txn_date: Union[dt.date, str, None] = None, amount=None) -> Optional[dict]:
        """The report's row for a bill - by vendor + bill #, else (a bill entered with no
        bill #) by vendor + date + amount. None when the report does not list it."""
        vk = vendor_key(vendor)
        rk = ref_key(docnum)
        if rk and (vk, rk) in self.bills:
            return self.bills[(vk, rk)]
        d, a = _as_date(txn_date), _money(amount)
        if not rk and d and a is not None and (vk, d, a) in self.by_amt:
            return {"state": self.by_amt[(vk, d, a)], "paid": False, "date": d, "amount": a,
                    "vendor": vendor or "", "ref": ""}
        return None

    def covers(self, created: Union[dt.date, str, None]) -> bool:
        """True when a bill entered on `created` had time to reach the report."""
        c = _as_date(created)
        return bool(self.fresh and self.as_of and c and c <= self.as_of.date())

    def describe(self) -> str:
        """One plain line for a banner: when it synced, and whether it is trusted."""
        when = self.as_of.strftime("%m/%d/%Y %I:%M %p") if self.as_of else "unknown time"
        how = "QuickBooks' stamp" if self.stamp_from_report else "file time"
        if self.fresh:
            return (f"Approval report synced from QuickBooks {when} ({how}) - {len(self.bills)} bills in the queue, "
                    f"{sum(1 for b in self.bills.values() if not b['paid'])} unpaid")
        return (f"Approval report is {self.age_hours / 24:.1f} days old (synced {when}) - too old to trust; "
                f"bills entered since 09/16 and unpaid read 'check QBO'")


_CURRENT: Optional[ApprovalReport] = None
_CURRENT_LOADED = False
_CURRENT_ERROR = ""


def load(path: Optional[Path] = None) -> Optional[ApprovalReport]:
    """The report at `path` (default: paths.bill_approval_xlsx()), or None when the file
    is missing or cannot be read - the rule then falls back to 'check QBO'."""
    global _CURRENT_ERROR
    try:
        if path is None:
            from shared import paths
            path = paths.bill_approval_xlsx()
        path = Path(path)
        if not path.is_file():
            _CURRENT_ERROR = f"no approval report at {path}"
            return None
        return ApprovalReport(path)
    except Exception as e:  # noqa: BLE001 - a bad report must never fail an AP run
        _CURRENT_ERROR = f"approval report not read ({e})"
        return None


def current(reload: bool = False) -> Optional[ApprovalReport]:
    """The report, read once per process."""
    global _CURRENT, _CURRENT_LOADED
    if reload or not _CURRENT_LOADED:
        _CURRENT = load()
        _CURRENT_LOADED = True
    return _CURRENT


def report_line(report: Optional[ApprovalReport]) -> str:
    """The banner line for a run: the report's own description, or why there is none."""
    if report is None:
        return (_CURRENT_ERROR or "no approval report") + " - bills entered since 09/16 and unpaid read 'check QBO'"
    return report.describe()


def approval_state(memo: Optional[str], created: Union[dt.date, str, None],
                   balance: Optional[float], *, report: Optional[ApprovalReport] = None,
                   vendor: Optional[str] = None, docnum: Optional[str] = None,
                   txn_date: Union[dt.date, str, None] = None, amount=None) -> str:
    """The bill's state - APPROVAL_NO / APPROVAL_PENDING / APPROVAL_REQUIRES /
    APPROVAL_CHECK / APPROVAL_YES. `created` is the day the bill was ENTERED in QBO
    (MetaData.CreateTime), not its TxnDate. Pass `report` (bill_approval.current())
    with the bill's vendor / bill # / date / total to read the approval queue; without
    it the answer is the memo + the entry date alone."""
    if not memo_approved(memo):
        return APPROVAL_NO
    try:
        open_amt = float(balance or 0)
    except (TypeError, ValueError):
        open_amt = 0.0
    if report is not None and report.fresh and open_amt > 0:
        row = report.lookup(vendor, docnum, txn_date, amount)
        if row:
            return row["state"]
    c = _as_date(created)
    if c and c >= APPROVAL_WORKFLOW_START and open_amt > 0:
        if report is not None and report.covers(c):
            return APPROVAL_YES          # in the workflow's reach, not in its queue: it left approved
        return APPROVAL_CHECK
    return APPROVAL_YES


def revised_reason(memo: str, allow_first_line: bool = False) -> str:
    """The reason off a `REVISED - <reason>` memo line, or "" when there is none.
    The PM writes it on the SECOND memo line (line 1 stays AP's project line), so
    line 1 is skipped by default: in the paper era AP itself opened line 1 with
    REVISED followed by the project line (`REVISED - CP861 - 4231 STATE HWY 161`),
    which is not a reason and not a vendor follow-up - 16 such bills sit in the
    2026 population. Pass allow_first_line=True for a bill entered under the QBO
    workflow, where a first-line REVISED can only be a PM who typed it in the wrong
    place. A bare REVISED returns "(no reason given)" so the bill still gets chased."""
    lines = (memo or "").splitlines()
    for line in (lines if allow_first_line else lines[1:]):
        t = line.strip()
        if t.upper().startswith("REVISED"):
            rest = t[len("REVISED"):].lstrip(" -:–").strip()
            return rest or "(no reason given)"
    return ""
