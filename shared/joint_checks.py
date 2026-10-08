"""joint_checks.py - the ONE joint-check rule (moved out of ledger/bill_payment_stub.py 10/08/2026,
when the project P&L needed it too: "I was trying to see all the payments made for that draw").

A joint check is the client's check made out to us AND a supplier. QuickBooks holds it as two
records for the same amount:

    Payment      (client -> us, applied to the draw's invoice, deposited to the Joint Checks account)
    BillPayment  (us -> the supplier, a Check drawn on the Joint Checks account)

The two are usually entered on different days (up to 23 days apart in 2026), so they pair in three
steps, each BillPayment used once:

    1. check #     the Payment's PaymentRefNum = the BillPayment's DocNumber, same amount, any date
    2. date        same date, same amount (a Payment typed without its check #)
    3. amount      same amount within 45 days, and the ONLY such pair on both sides - the case where
                   the client sent no check # and AP enters the supplier's payment once the copy of
                   the joint check arrives (Van Brunt, 10/07/2026). Shown as "amount only".

A Payment deposited to the Joint Checks account with no BillPayment yet is still a joint check: the
supplier side is not entered. AR writes who it went to in the Payment's memo ("Joint check to RCI
Ready Cable & Proficient 09/24/2026 - paid to RCI"); `memo_of` hands that to the reader until AP
enters the payment out. Pure functions on raw QBO records - no I/O.
"""
from __future__ import annotations

import datetime as dt
from typing import Dict, Iterable, List, Optional

AMOUNT_ONLY_DAYS = 45


def _num(v) -> float:
    try:
        return float(v or 0)
    except (TypeError, ValueError):
        return 0.0


def _ref(v) -> str:
    return str(v or "").strip().lstrip("0")


def _day(s) -> Optional[dt.date]:
    try:
        return dt.date.fromisoformat(str(s or "")[:10])
    except ValueError:
        return None


def _amt(r: dict) -> float:
    return round(_num(r.get("TotalAmt")), 2)


def is_joint_account(acct_name: str) -> bool:
    """A check written from the Joint Checks account is a joint check (the client's check to us + the vendor)."""
    return "joint" in str(acct_name or "").lower()


def _bank(bp: dict) -> dict:
    return (bp.get("CheckPayment") or {}).get("BankAccountRef") or {}


def is_joint_bill_payment(bp: dict) -> bool:
    """A BillPayment paid by Check out of the Joint Checks account."""
    return (bp.get("PayType") or "") == "Check" and is_joint_account(_bank(bp).get("name"))


def joint_account_ids(bill_payments: Iterable[dict]) -> set:
    """The Joint Checks account id(s), learned from the joint BillPayments (a Payment's
    DepositToAccountRef carries the id only, never the name)."""
    return {str(_bank(b).get("value")) for b in bill_payments or []
            if is_joint_bill_payment(b) and _bank(b).get("value")}


def is_joint_deposit(payment: dict, joint_ids: set) -> bool:
    """A client Payment deposited to the Joint Checks account."""
    return str((payment.get("DepositToAccountRef") or {}).get("value") or "") in joint_ids


def memo_of(payment: dict) -> str:
    """The AR clerk's note on the Payment (who the joint check went to), one line."""
    return " ".join(str(payment.get("PrivateNote") or "").split())


def pair_joint_checks(payments: Iterable[dict], bill_payments: Iterable[dict]) -> Dict[str, dict]:
    """{client Payment Id -> {"bp": the joint BillPayment that paid it out, "how": "check #" | "date" |
    "amount only"}}. A Payment with no joint check out (or an amount off by cents) is absent - that is
    the gap to look at. Payments deposited anywhere but the Joint Checks account never pair."""
    joint = [b for b in bill_payments or [] if is_joint_bill_payment(b)]
    ids = joint_account_ids(joint)
    pays = [p for p in payments or []
            if not p.get("DepositToAccountRef") or not ids or is_joint_deposit(p, ids)]
    used: set = set()
    out: Dict[str, dict] = {}

    def take(p, b, how):
        used.add(id(b))
        out[str(p.get("Id"))] = {"bp": b, "how": how}

    by_ref: Dict[tuple, List[dict]] = {}
    for b in joint:
        by_ref.setdefault((_ref(b.get("DocNumber")), _amt(b)), []).append(b)
    for p in pays:
        r = _ref(p.get("PaymentRefNum"))
        for b in (by_ref.get((r, _amt(p)), []) if r else []):
            if id(b) not in used:
                take(p, b, "check #")
                break

    for p in pays:
        if str(p.get("Id")) in out:
            continue
        for b in joint:
            if (id(b) not in used and _amt(b) == _amt(p)
                    and str(b.get("TxnDate") or "") == str(p.get("TxnDate") or "")):
                take(p, b, "date")
                break

    # amount only: both sides must be alone in their amount within the window
    def near(a, b):
        da, db = _day(a.get("TxnDate")), _day(b.get("TxnDate"))
        return bool(da and db) and abs((da - db).days) <= AMOUNT_ONLY_DAYS
    left_p = [p for p in pays if str(p.get("Id")) not in out]
    left_b = [b for b in joint if id(b) not in used]
    for p in left_p:
        cands = [b for b in left_b if id(b) not in used and _amt(b) == _amt(p) and near(p, b)]
        if len(cands) != 1:
            continue
        rivals = [q for q in left_p if q is not p and str(q.get("Id")) not in out
                  and _amt(q) == _amt(p) and near(q, cands[0])]
        if not rivals:
            take(p, cands[0], "amount only")
    return out


def joint_vendor_by_payment(payments: Iterable[dict], bill_payments: Iterable[dict]) -> Dict[str, str]:
    """{client Payment Id -> the supplier its joint check paid}. Payments with no joint check are absent."""
    return {pid: str((m["bp"].get("VendorRef") or {}).get("name") or "")
            for pid, m in pair_joint_checks(payments, bill_payments).items()}


def find_joint_payment(bp: dict, payments: Iterable[dict],
                       bill_payments: Iterable[dict] = ()) -> Optional[dict]:
    """The client's Payment behind ONE joint check. `payments` = the candidate Payments (the stub
    passes every Payment of the same amount); `bill_payments` = the BillPayments of that amount, so two
    checks for the same amount cannot both claim one Payment. None when `bp` is not a joint check or
    nothing pairs."""
    if not is_joint_bill_payment(bp):
        return None
    pays = list(payments or [])
    others = [b for b in bill_payments or [] if str(b.get("Id")) != str(bp.get("Id"))] + [bp]
    for pid, m in pair_joint_checks(pays, others).items():
        if m["bp"] is bp:
            return next(p for p in pays if str(p.get("Id")) == pid)
    return None


def _selftest() -> None:
    acct = {"value": "2472", "name": "Joint Checks Account"}
    pay = {"Id": "p1", "TxnDate": "2026-09-25", "TotalAmt": 133400.82, "PaymentRefNum": "37754",
           "DepositToAccountRef": {"value": "2472"}}
    twin = dict(pay, Id="p2", PaymentRefNum="37760")
    other = {"Id": "p3", "TxnDate": "2026-09-25", "TotalAmt": 500.0}
    bp = {"Id": "b1", "TxnDate": "2026-09-25", "TotalAmt": 133400.82, "DocNumber": "37754", "PayType": "Check",
          "VendorRef": {"name": "BURNCO TEXAS LLC"}, "CheckPayment": {"BankAccountRef": acct}}
    ours = dict(bp, CheckPayment={"BankAccountRef": {"value": "1", "name": "Operating"}})
    assert joint_vendor_by_payment([pay, other], [bp]) == {"p1": "BURNCO TEXAS LLC"}
    assert joint_vendor_by_payment([pay, twin], [bp]) == {"p1": "BURNCO TEXAS LLC"}   # the check # picks
    assert joint_vendor_by_payment([pay], [ours]) == {}                                # not the joint account
    late = dict(bp, TxnDate="2026-10-10", DocNumber="0037754")                      # entered days later
    assert pair_joint_checks([pay], [late])["p1"]["how"] == "check #"
    assert joint_vendor_by_payment([pay], [dict(late, TotalAmt=133400.85)]) == {}    # cents off = a gap
    # no check # on the client's Payment (Van Brunt): the date first, then the amount alone
    vb = {"Id": "v1", "TxnDate": "2026-10-07", "TotalAmt": 16469.62, "DepositToAccountRef": {"value": "2472"},
          "PrivateNote": "Joint check to RCI Ready Cable & Proficient 09/24/2026 - paid to RCI"}
    rci = dict(bp, Id="b9", TxnDate="2026-10-15", TotalAmt=16469.62, DocNumber="4471",
               VendorRef={"name": "RCI READY CABLE"})
    assert pair_joint_checks([vb], [dict(rci, TxnDate="2026-10-07")])["v1"]["how"] == "date"
    assert pair_joint_checks([vb], [rci])["v1"]["how"] == "amount only"
    assert pair_joint_checks([vb], [dict(rci, TxnDate="2026-12-31")]) == {}           # outside 45 days
    assert pair_joint_checks([vb, dict(vb, Id="v2")], [rci]) == {}                    # two could claim it: none
    assert pair_joint_checks([dict(vb, DepositToAccountRef={"value": "1"})], [rci]) == {}   # not a joint deposit
    assert joint_account_ids([bp, ours]) == {"2472"} and is_joint_deposit(vb, {"2472"})
    assert memo_of(vb).startswith("Joint check to RCI")
    # the stub's single lookup
    assert find_joint_payment(bp, [pay, twin]) is pay
    assert find_joint_payment(rci, [vb]) is vb
    assert find_joint_payment(ours, [pay]) is None
    assert find_joint_payment(dict(bp, PayType="CreditCard"), [pay]) is None
    print("joint_checks selftest: OK")


if __name__ == "__main__":
    _selftest()
