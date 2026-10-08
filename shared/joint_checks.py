"""joint_checks.py - the ONE joint-check rule (moved out of ledger/bill_payment_stub.py 10/08/2026,
when the project P&L needed it too: "I was trying to see all the payments made for that draw").

A joint check is the client's check made out to us AND a supplier. QuickBooks holds it as two
records for the same amount:

    Payment      (client -> us, applied to the draw's invoice, deposited to the Joint Checks account)
    BillPayment  (us -> the supplier, a Check drawn on the Joint Checks account)

The check # is the Payment's PaymentRefNum and the BillPayment's DocNumber. The two are usually entered on
different days, so the pair is found by check # + amount first, then date + amount. Pure functions on raw QBO records - no I/O.
"""
from __future__ import annotations

from typing import Callable, Dict, Iterable, List, Optional


def _num(v) -> float:
    try:
        return float(v or 0)
    except (TypeError, ValueError):
        return 0.0


def is_joint_account(acct_name: str) -> bool:
    """A check written from the Joint Checks account is a joint check (the client's check to us + the vendor)."""
    return "joint" in str(acct_name or "").lower()


def is_joint_bill_payment(bp: dict) -> bool:
    """A BillPayment paid by Check out of the Joint Checks account."""
    if (bp.get("PayType") or "") != "Check":
        return False
    acct = ((bp.get("CheckPayment") or {}).get("BankAccountRef") or {}).get("name") or ""
    return is_joint_account(acct)


def find_joint_payment(bp: dict, payments_on: Callable[[str], List[dict]]) -> Optional[dict]:
    """The client's Payment behind a joint check: same date, same amount; the check # as its ref breaks a tie.
    None when the check is not on the Joint Checks account or no single payment matches."""
    if not is_joint_bill_payment(bp):
        return None
    total = _num(bp.get("TotalAmt"))
    ref = str(bp.get("DocNumber") or "")
    hits = [p for p in payments_on(str(bp.get("TxnDate") or "")) if abs(_num(p.get("TotalAmt")) - total) < 0.005]
    if len(hits) > 1 and ref:
        hits = [p for p in hits if str(p.get("PaymentRefNum") or "") == ref] or hits
    return hits[0] if len(hits) == 1 else None


def _ref(v) -> str:
    return str(v or "").strip().lstrip("0")


def pair_joint_checks(payments: Iterable[dict], bill_payments: Iterable[dict]) -> Dict[str, dict]:
    """{client Payment Id -> the joint BillPayment that paid it out}. Two passes, each BillPayment used once:
    1. the same check # and amount, ANY date - the usual case: the client's payment is entered the day the
       check came in, the supplier's bill payment days later (up to 23 days apart in 2026);
    2. the same date and amount, for a payment typed without its check #.
    A payment with no joint check out (or an amount off by cents) is absent - that is the gap to look at."""
    joint = [b for b in bill_payments or [] if is_joint_bill_payment(b)]
    amt = lambda r: round(_num(r.get("TotalAmt")), 2)   # noqa: E731
    used: set = set()
    out: Dict[str, dict] = {}
    by_ref: Dict[tuple, List[dict]] = {}
    for b in joint:
        by_ref.setdefault((_ref(b.get("DocNumber")), amt(b)), []).append(b)
    pays = list(payments or [])
    for p in pays:
        r = _ref(p.get("PaymentRefNum"))
        for b in (by_ref.get((r, amt(p)), []) if r else []):
            if id(b) not in used:
                used.add(id(b))
                out[str(p.get("Id"))] = b
                break
    by_day: Dict[tuple, List[dict]] = {}
    for b in joint:
        if id(b) not in used:
            by_day.setdefault((str(b.get("TxnDate") or ""), amt(b)), []).append(b)
    for p in pays:
        if str(p.get("Id")) in out:
            continue
        for b in by_day.get((str(p.get("TxnDate") or ""), amt(p)), []):
            if id(b) not in used:
                used.add(id(b))
                out[str(p.get("Id"))] = b
                break
    return out


def joint_vendor_by_payment(payments: Iterable[dict], bill_payments: Iterable[dict]) -> Dict[str, str]:
    """{client Payment Id -> the supplier its joint check paid}. Payments with no joint check are absent."""
    return {pid: str((b.get("VendorRef") or {}).get("name") or "")
            for pid, b in pair_joint_checks(payments, bill_payments).items()}


def _selftest() -> None:
    pay = {"Id": "p1", "TxnDate": "2026-09-25", "TotalAmt": 133400.82, "PaymentRefNum": "37754"}
    twin = dict(pay, Id="p2", PaymentRefNum="37760")
    other = {"Id": "p3", "TxnDate": "2026-09-25", "TotalAmt": 500.0}
    bp = {"TxnDate": "2026-09-25", "TotalAmt": 133400.82, "DocNumber": "37754", "PayType": "Check",
          "VendorRef": {"name": "BURNCO TEXAS LLC"},
          "CheckPayment": {"BankAccountRef": {"name": "Joint Checks Account"}}}
    ours = dict(bp, CheckPayment={"BankAccountRef": {"name": "Operating"}})
    assert joint_vendor_by_payment([pay, other], [bp]) == {"p1": "BURNCO TEXAS LLC"}
    assert joint_vendor_by_payment([pay, twin], [bp]) == {"p1": "BURNCO TEXAS LLC"}   # the check # breaks the tie
    assert joint_vendor_by_payment([pay], [ours]) == {}                                # not the joint account
    assert find_joint_payment(dict(bp, PayType="CreditCard"), lambda d: [pay]) is None
    late = dict(bp, TxnDate="2026-10-10", DocNumber="0037754")                      # entered days later
    assert joint_vendor_by_payment([pay], [late]) == {"p1": "BURNCO TEXAS LLC"}
    assert joint_vendor_by_payment([pay], [dict(late, TotalAmt=133400.85)]) == {}    # cents off = a gap
    print("joint_checks selftest: OK")


if __name__ == "__main__":
    _selftest()
