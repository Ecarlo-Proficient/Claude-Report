"""invoice_view.py - the ledger's invoice viewer data (owner 2026-10-11: "invoice # not clickable i want that to
open just like attachment and see the invoice load from qbo ... and also see the line items as is so you don't
have to leave the ledger"; and for a draw billed on several invoices - MFD192, three contracts - "show clearly
... that there are more invoices adding to the total").

    view(["34417", "34418"], project="MFD192") -> {"ok": True, "invoices": [...], "total": ..., "open": ...}

Each invoice from the raw QBO mirror (shared/qbo_mirror), as QuickBooks holds it: the header, every line as
typed (item · description · qty · rate · amount; group lines opened, subtotals skipped), what it is open for,
and every payment applied to it - a joint check names the supplier it paid (shared/joint_checks), a joint
deposit with no payment out yet carries AR's memo. The PDF itself comes from invoice_pdf.py. Read-only.
"""
from __future__ import annotations

import sys
from pathlib import Path
from typing import Dict, List, Optional

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from shared import qbo_mirror     # noqa: E402
from shared import joint_checks   # noqa: E402


def _n(v) -> float:
    try:
        return round(float(v or 0), 2)
    except (TypeError, ValueError):
        return 0.0


def _lines(inv: dict) -> List[dict]:
    """The invoice's lines as typed in QuickBooks (group lines opened; QuickBooks' own subtotal line skipped)."""
    out: List[dict] = []

    def walk(lines, group=""):
        for ln in lines or []:
            dt = ln.get("DetailType") or ""
            if dt == "SubTotalLineDetail":
                continue
            if dt == "GroupLineDetail":
                g = ln.get("GroupLineDetail") or {}
                walk(g.get("Line"), (g.get("GroupItemRef") or {}).get("name") or "")
                continue
            d = ln.get("SalesItemLineDetail") or {}
            item = (d.get("ItemRef") or {}).get("name") or ("Discount" if dt == "DiscountLineDetail" else "")
            out.append({"item": item, "description": ln.get("Description") or "", "group": group,
                        "qty": d.get("Qty"), "rate": d.get("UnitPrice"), "amount": _n(ln.get("Amount")),
                        "service_date": d.get("ServiceDate") or ""})
    walk(inv.get("Line"))
    return out


_JOINT_IDS: Optional[set] = None


def _joint_ids(con) -> set:
    """The Joint Checks account id(s), learned once from this year's joint bill payments."""
    global _JOINT_IDS
    if _JOINT_IDS is None:
        _JOINT_IDS = joint_checks.joint_account_ids(qbo_mirror.load("BillPayment", "txn_date >= ?", ("2026-01-01",), con=con))
    return _JOINT_IDS


def _payments(inv: dict, con) -> List[dict]:
    """Every payment applied to this invoice (the invoice's own LinkedTxn), the slice that landed on it, and who a
    joint check paid."""
    rows = []
    for lt in inv.get("LinkedTxn") or []:
        if lt.get("TxnType") != "Payment":
            continue
        p = qbo_mirror.get("Payment", str(lt.get("TxnId")), con=con)
        if not p:
            continue
        applied = sum(_n(ln.get("Amount")) for ln in p.get("Line") or []
                      for x in ln.get("LinkedTxn") or [] if x.get("TxnType") == "Invoice" and str(x.get("TxnId")) == str(inv.get("Id")))
        amt = _n(p.get("TotalAmt"))
        same = qbo_mirror.load("Payment", "round(total, 2) = ?", (amt,), con=con)
        bps = qbo_mirror.load("BillPayment", "round(total, 2) = ?", (amt,), con=con)
        hit = joint_checks.pair_joint_checks(same, bps).get(str(p.get("Id")))
        joint_ids = joint_checks.joint_account_ids(bps) or _joint_ids(con)
        if hit:
            kind, to = "joint", str((hit["bp"].get("VendorRef") or {}).get("name") or "")
            how = hit["how"]
        elif joint_checks.is_joint_deposit(p, joint_ids):
            kind, to, how = "open", "", ""
        else:
            kind, to, how = "us", "", ""
        rows.append({"id": str(p.get("Id")), "date": p.get("TxnDate") or "", "ref": p.get("PaymentRefNum") or "",
                     "applied": round(applied, 2), "total": amt, "kind": kind, "to": to, "how": how,
                     "memo": joint_checks.memo_of(p) if kind == "open" else ""})
    rows.sort(key=lambda r: (r["date"], r["ref"]))
    return rows


def _pick(hits: List[dict], project: str) -> Optional[dict]:
    if len(hits) > 1 and project:
        pj = project.upper()
        mine = [h for h in hits if pj in str((h.get("CustomerRef") or {}).get("name") or "").upper()
                or pj in str(h.get("PrivateNote") or "").upper()]
        hits = mine or hits
    return hits[0] if hits else None


def view(nos: List[str], project: str = "") -> Dict:
    nos = [str(n).strip() for n in nos if str(n).strip()][:24]
    if not nos:
        return {"ok": False, "error": "no invoice #"}
    if not qbo_mirror.db_path().exists():
        return {"ok": False, "error": "the QuickBooks copy (mirror) is not on this Mac"}
    con = qbo_mirror.connect()
    try:
        out = []
        for no in nos:
            inv = _pick(qbo_mirror.load("Invoice", "doc_number = ?", (no,), con=con), project)
            if not inv:
                out.append({"doc_number": no, "missing": True})
                continue
            lines = _lines(inv)
            out.append({
                "doc_number": no, "id": str(inv.get("Id")), "sync": str(inv.get("SyncToken") or "0"),
                "date": inv.get("TxnDate") or "", "due": inv.get("DueDate") or "",
                "customer": (inv.get("CustomerRef") or {}).get("name") or "",
                "memo": " ".join(str(inv.get("PrivateNote") or "").split()),
                "message": " ".join(str((inv.get("CustomerMemo") or {}).get("value") or "").split()),
                "terms": (inv.get("SalesTermRef") or {}).get("name") or "",
                "total": _n(inv.get("TotalAmt")), "balance": _n(inv.get("Balance")),
                "gross": round(sum(ln["amount"] for ln in lines if ln["amount"] > 0), 2),
                "retainage": round(-sum(ln["amount"] for ln in lines if ln["amount"] < 0), 2),
                "lines": lines, "payments": _payments(inv, con),
            })
        found = [i for i in out if not i.get("missing")]
        return {"ok": True, "invoices": out, "total": round(sum(i["total"] for i in found), 2),
                "open": round(sum(i["balance"] for i in found), 2), "stamp": str(qbo_mirror.stamp() or "")}
    finally:
        con.close()


def _selftest() -> None:
    inv = {"Id": "9", "Line": [
        {"DetailType": "SalesItemLineDetail", "Amount": 1000, "Description": "Draw #3",
         "SalesItemLineDetail": {"ItemRef": {"name": "Progress"}, "Qty": 1, "UnitPrice": 1000}},
        {"DetailType": "GroupLineDetail", "GroupLineDetail": {"GroupItemRef": {"name": "Phase 2"}, "Line": [
            {"DetailType": "SalesItemLineDetail", "Amount": 50, "SalesItemLineDetail": {"ItemRef": {"name": "CO"}}}]}},
        {"DetailType": "SalesItemLineDetail", "Amount": -100, "Description": "Retainage",
         "SalesItemLineDetail": {"ItemRef": {"name": "Retainage"}}},
        {"DetailType": "SubTotalLineDetail", "Amount": 950}]}
    ls = _lines(inv)
    assert [x["amount"] for x in ls] == [1000, 50, -100] and ls[1]["group"] == "Phase 2"
    assert _pick([{"CustomerRef": {"name": "A:CP1"}}, {"CustomerRef": {"name": "B:MFD192 - X"}}], "MFD192")["CustomerRef"]["name"].startswith("B")
    assert view([])["ok"] is False
    print("invoice_view selftest: OK")


if __name__ == "__main__":
    _selftest()
