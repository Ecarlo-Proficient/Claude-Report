#!/usr/bin/env python3
"""
load_health.py - the company-health metric layer -> ledger `health_snapshot`.

The Health tab derives most of its numbers from tables other loaders already
fill (AR from billing_event, AP from ap_bill_line, backlog from wip_snapshot,
Sub LOC from sub_loc_run). This loader pulls ONLY what the ledger can't derive:

  * bank_accounts - Bank + Credit Card balances. Cash = Bank only. QBO's bank
    feed is broken, so balances move only when the owner uploads - every cash
    figure carries this pull's as_of stamp (ruling 2026-07-28).
  * retainage     - GL accounts named 'retainage' (receivable vs payable).
  * pl_blocks     - accrual P&L totals for MTD / YTD / prior-YTD (margins +
    the break-even inputs), via shared/qbo_pl.
  * weekly_flow   - 13 weeks of cash in (Payment) vs cash out (BillPayment +
    Purchase), plus the burn/runway summary computed off it.
  * recurring     - the recurring-obligations register (FIN-12), via
    shared/recurring: fixed overhead + debt service by month with
    CHANGED / STOPPED / NEW detection.

Read-only against QBO; one Touch ID per run. Full-replaces the health_snapshot
rows (idempotent). NOT a source of truth - QBO is.

USAGE
  python3 ledger/load_health.py                # pull + write
  python3 ledger/load_health.py --dry-run      # pull + print, write nothing
  python3 ledger/load_health.py --selftest     # offline proof (no QBO), throwaway DB
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import sqlite3
import sys
from collections import defaultdict
from pathlib import Path
from typing import Any, Dict, List, Optional

HERE = Path(__file__).resolve().parent
PROJECT_ROOT = HERE.parent
sys.path.insert(0, str(PROJECT_ROOT))
from shared import paths, qbo_api, qbo_pl, recurring  # noqa: E402
from shared.qbo_cache import QboCache                  # noqa: E402

SCHEMA_SQL = HERE / "schema.sql"
DEFAULT_DB = paths.get_path(
    "ACB_LEDGER_DB",
    Path.home() / "Library" / "Application Support" / "Proficient" / "ledger.sqlite3")

FLOW_WEEKS = 13   # trailing window for burn / runway (same as the legacy dashboard)


def _connect(db_path: Path) -> sqlite3.Connection:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(str(db_path))
    con.row_factory = sqlite3.Row
    con.executescript(SCHEMA_SQL.read_text(encoding="utf-8"))
    return con


# ── QBO pulls ────────────────────────────────────────────────────────────────

def fetch_bank_accounts(access: str, cid: str) -> List[Dict[str, Any]]:
    """Active Bank + Credit Card accounts with CurrentBalance (cash = Bank only)."""
    rows = qbo_api.query_all(access, cid, "Account",
                             "AccountType IN ('Bank','Credit Card') AND Active=true")
    out = [{"name": a.get("Name", ""), "type": a.get("AccountType", ""),
            "balance": float(a.get("CurrentBalance", 0) or 0)} for a in rows]
    return sorted(out, key=lambda r: (r["type"] != "Bank", -r["balance"]))


def fetch_retainage(access: str, cid: str) -> Dict[str, float]:
    """Totals of active GL accounts named 'retainage', receivable vs payable."""
    rows = qbo_api.query_all(access, cid, "Account", "Active=true")
    recv = pay = 0.0
    for a in rows:
        name = (a.get("Name") or "")
        if "retainage" not in name.lower():
            continue
        bal = float(a.get("CurrentBalance", 0) or 0)
        atype = (a.get("AccountType") or "").lower()
        if "liabilit" in atype or "payable" in atype or "payable" in name.lower():
            pay += bal
        else:
            recv += bal
    return {"receivable": recv, "payable": pay}


def fetch_pl_blocks(access: str, cid: str, today: dt.date) -> Dict[str, dict]:
    """{'mtd','ytd','prior'} accrual P&L totals (breakeven-shaped keys)."""
    j1 = dt.date(today.year, 1, 1)
    m1 = today.replace(day=1)
    pj1 = dt.date(today.year - 1, 1, 1)
    p_end = today.replace(year=today.year - 1) if not (today.month == 2 and today.day == 29) \
        else dt.date(today.year - 1, 2, 28)
    return {
        "mtd": qbo_pl.pl_totals(access, cid, m1.isoformat(), today.isoformat()),
        "ytd": qbo_pl.pl_totals(access, cid, j1.isoformat(), today.isoformat()),
        "prior": qbo_pl.pl_totals(access, cid, pj1.isoformat(), p_end.isoformat()),
    }


def _txn_totals(access: str, cid: str, entity: str, since: str) -> List[Dict[str, Any]]:
    rows = qbo_api.query_all(access, cid, entity, f"TxnDate >= '{since}'")
    return [{"TxnDate": (r.get("TxnDate") or "")[:10],
             "TotalAmt": float(r.get("TotalAmt", 0) or 0)} for r in rows]


def weekly_cash_flow(payments, bill_payments, purchases, today: dt.date,
                     weeks: int = FLOW_WEEKS) -> List[Dict[str, Any]]:
    """Weekly net cash flow, Monday-anchored. Inflow = customer Payments;
    outflow = BillPayments + Purchases (cash actually moved, not accruals)."""
    earliest = today - dt.timedelta(weeks=weeks)

    def anchor(s: str) -> Optional[dt.date]:
        try:
            d = dt.date.fromisoformat(s)
        except (ValueError, TypeError):
            return None
        if d < earliest:
            return None
        return d - dt.timedelta(days=d.weekday())

    inflow: Dict[dt.date, float] = defaultdict(float)
    outflow: Dict[dt.date, float] = defaultdict(float)
    for p in payments:
        a = anchor(p.get("TxnDate", ""))
        if a:
            inflow[a] += p.get("TotalAmt", 0)
    for src in (bill_payments, purchases):
        for p in src:
            a = anchor(p.get("TxnDate", ""))
            if a:
                outflow[a] += p.get("TotalAmt", 0)
    return [{"week_of": wk.isoformat(),
             "received": round(inflow.get(wk, 0.0), 2),
             "paid": round(outflow.get(wk, 0.0), 2),
             "net": round(inflow.get(wk, 0.0) - outflow.get(wk, 0.0), 2)}
            for wk in sorted(set(inflow) | set(outflow))]


def flow_summary(cash: float, flow: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Burn + runway off the trailing weekly flow. Runway = cash / avg negative
    net; None (unconstrained) while the trend is net cash-positive."""
    if not flow:
        return {"avg_weekly_paid": 0.0, "avg_weekly_net": 0.0, "runway_weeks": None,
                "note": "no cash-flow data"}
    avg_paid = sum(w["paid"] for w in flow) / len(flow)
    avg_net = sum(w["net"] for w in flow) / len(flow)
    if avg_net >= 0:
        return {"avg_weekly_paid": round(avg_paid, 2), "avg_weekly_net": round(avg_net, 2),
                "runway_weeks": None,
                "note": "net cash-positive over the window - runway not constrained"}
    burn = -avg_net
    note = f"{FLOW_WEEKS}-wk avg net burn ${burn:,.0f}/week"
    if cash <= 0:
        note += " · bank balance is negative in QBO"
    return {"avg_weekly_paid": round(avg_paid, 2), "avg_weekly_net": round(avg_net, 2),
            "runway_weeks": round(max(cash, 0.0) / burn, 1) if burn else None,
            "note": note}


# ── write ────────────────────────────────────────────────────────────────────

def write(con: sqlite3.Connection, payloads: Dict[str, Any], as_of: str) -> None:
    now = dt.datetime.now().isoformat(timespec="seconds")
    for key, payload in payloads.items():
        con.execute(
            "INSERT INTO health_snapshot (key, payload, as_of, loaded_at) VALUES (?,?,?,?) "
            "ON CONFLICT(key) DO UPDATE SET payload=excluded.payload, as_of=excluded.as_of, "
            "loaded_at=excluded.loaded_at",
            (key, json.dumps(payload, default=str), as_of, now))
    con.commit()


def run(db_path: Path, dry_run: bool) -> None:
    today = dt.date.today()
    as_of = dt.datetime.now().isoformat(timespec="minutes")
    print("\n  COMPANY HEALTH - metric layer pull")
    access, cid = qbo_api.load_credentials()
    print("  pulling bank + retainage accounts ...")
    banks = fetch_bank_accounts(access, cid)
    cash = sum(a["balance"] for a in banks if a["type"] == "Bank")
    retain = fetch_retainage(access, cid)
    print(f"    {len(banks)} account(s) · cash (bank only) ${cash:,.0f}")
    print("  pulling P&L blocks (MTD / YTD / prior-YTD) ...")
    blocks = fetch_pl_blocks(access, cid, today)
    ytd = blocks.get("ytd", {})
    if ytd.get("income"):
        gm = (ytd.get("gross_profit") or 0) / ytd["income"]
        print(f"    YTD income ${ytd['income']:,.0f} · GM {gm * 100:.1f}%")
    print(f"  pulling {FLOW_WEEKS} weeks of cash flow (payments / bill payments / purchases) ...")
    since = (today - dt.timedelta(weeks=FLOW_WEEKS)).isoformat()
    flow = weekly_cash_flow(_txn_totals(access, cid, "Payment", since),
                            _txn_totals(access, cid, "BillPayment", since),
                            _txn_totals(access, cid, "Purchase", since), today)
    summ = flow_summary(cash, flow)
    rw = summ.get("runway_weeks")
    print(f"    avg weekly out ${summ['avg_weekly_paid']:,.0f} · "
          f"runway {'unconstrained' if rw is None else f'{rw:.1f} wk'}")
    print("  building the recurring-obligations register ...")
    reg = recurring.build(QboCache(access, cid))
    print(f"    fixed overhead ${reg['fixed_overhead_month']:,.0f}/mo · "
          f"debt service ${reg['debt_service_month']:,.0f}/mo · "
          f"{len(reg['alerts'])} alert(s)")
    payloads = {
        "bank_accounts": {"accounts": banks, "cash": cash},
        "retainage": retain,
        "pl_blocks": blocks,
        "weekly_flow": {"weeks": flow, "summary": summ},
        "recurring": {
            "months": reg["months"],
            "fixed_overhead_month": reg["fixed_overhead_month"],
            "debt_service_month": reg["debt_service_month"],
            "total_monthly_obligation": reg["total_monthly_obligation"],
            "overhead": reg["overhead"], "debt": reg["debt"],
            "alerts": reg["alerts"], "refinancing": reg["refinancing"],
        },
    }
    if dry_run:
        print("  (dry-run: nothing written)")
        return
    con = _connect(db_path)
    try:
        write(con, payloads, as_of)
    finally:
        con.close()
    print(f"  wrote {len(payloads)} snapshot payload(s) to {db_path}")


def selftest() -> int:
    """Offline proof of the flow math + writer on a throwaway DB (no QBO)."""
    import tempfile
    today = dt.date(2026, 8, 31)
    pays = [{"TxnDate": "2026-08-24", "TotalAmt": 100.0},
            {"TxnDate": "2026-08-25", "TotalAmt": 50.0},
            {"TxnDate": "2020-01-01", "TotalAmt": 999.0}]      # outside the window - dropped
    bps = [{"TxnDate": "2026-08-24", "TotalAmt": 120.0}]
    purs = [{"TxnDate": "2026-08-18", "TotalAmt": 80.0}]
    flow = weekly_cash_flow(pays, bps, purs, today)
    assert len(flow) == 2, flow
    assert flow[-1]["received"] == 150.0 and flow[-1]["paid"] == 120.0, flow[-1]
    assert flow[0]["paid"] == 80.0 and flow[0]["net"] == -80.0, flow[0]
    s = flow_summary(1000.0, flow)
    # avg net = (-80 + 30) / 2 = -25 -> runway = 1000 / 25 = 40 weeks
    assert s["runway_weeks"] == 40.0, s
    assert flow_summary(1000.0, [{"paid": 10.0, "net": 5.0}])["runway_weeks"] is None
    with tempfile.TemporaryDirectory() as td:
        con = _connect(Path(td) / "t.sqlite3")
        write(con, {"weekly_flow": {"weeks": flow, "summary": s}}, "2026-08-31T12:00")
        write(con, {"weekly_flow": {"weeks": flow, "summary": s}}, "2026-08-31T12:05")  # idempotent upsert
        n = con.execute("SELECT COUNT(*) FROM health_snapshot").fetchone()[0]
        row = con.execute("SELECT payload, as_of FROM health_snapshot WHERE key='weekly_flow'").fetchone()
        con.close()
    assert n == 1, n
    assert row["as_of"] == "2026-08-31T12:05", row["as_of"]
    got = json.loads(row["payload"])
    assert got["summary"]["runway_weeks"] == 40.0, got["summary"]
    print("  selftest OK: 2-week flow, runway 40.0 wk, upsert idempotent")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(description="Load the company-health metric layer into the ledger.")
    ap.add_argument("--db", type=Path, default=DEFAULT_DB)
    ap.add_argument("--dry-run", action="store_true", help="pull + print; write nothing")
    ap.add_argument("--selftest", action="store_true", help="offline proof (no QBO), throwaway DB")
    args = ap.parse_args()
    if args.selftest:
        return selftest()
    run(args.db, args.dry_run)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
