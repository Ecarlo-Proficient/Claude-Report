"""trail.py - the money trail: every QBO line behind a project's Costs / Billed.

`GET /api/trail?project=<PN>&kind=costs|billed|both[&csv=1]` - the line-level audit
behind the two totals the ledger shows everywhere. Nothing is hidden: each cost line
is one Bill / Expense line (vendor, bill #, memo, description, code, amount) and each
billed line is one Invoice; a running total per kind is the "red line" against the
budget (ETC) and the contract. Read-only, parameterised SQL, no realm printed.

Why a separate module: `dashboard.py` only routes here (two lines), so the trail can be
unit-tested offline (`python3 ledger/trail.py --selftest`) and reused by a CSV export.

Rules that hold here (repo CLAUDE.md):
  * cost = LINE amounts; `bill_total` rides along for display only, never summed;
  * no created / edited stamps or "entered by" (parked by the owner, 2026-09-01);
  * the totals are the SAME numbers the drawer shows (SUM(cost_line), SUM(billing_event),
    v_wip_latest), so the page and the drawer can never disagree.
"""
from __future__ import annotations

import csv
import io
import json
import sqlite3
import sys
from typing import Optional

KINDS = ("costs", "billed", "both")
CSV_COLS = ["kind", "date", "party", "doc_number", "ref", "memo", "description", "cost_code",
            "account", "is_sub", "amount", "running_total", "txn_type", "txn_id", "line_id",
            "bill_total", "has_attachment", "qbo_url"]

# QBO deep links - the SAME company-scoped form the front end builds (app.js qboUrl):
# a bare /app/<kind>?txnId= link opens in whichever company the browser is on.
_QBO_KIND = {"Bill": "bill", "Expense": "expense", "Invoice": "invoice"}


def qbo_url(txn_type: str, txn_id: str, realm: Optional[str]) -> Optional[str]:
    kind = _QBO_KIND.get(txn_type or "")
    if not kind or not txn_id:
        return None
    if realm:
        return f"https://qbo.intuit.com/app/login?pagereq={kind}?txnId={txn_id}&deeplinkcompanyid={realm}"
    return f"https://qbo.intuit.com/app/{kind}?txnId={txn_id}"


def _cols(con, table: str) -> set:
    return {r[1] for r in con.execute(f"PRAGMA table_info({table})")}


def _realm(con) -> Optional[str]:
    try:
        r = con.execute("SELECT value FROM meta WHERE key='qbo_realm'").fetchone()
        return r[0] if r and r[0] else None
    except sqlite3.OperationalError:
        return None


def _f(v) -> Optional[float]:
    try:
        return None if v is None else float(v)
    except (TypeError, ValueError):
        return None


def build(con, project: str, kind: str = "both") -> Optional[dict]:
    """The trail payload for one project, or None when the project is unknown."""
    project = (project or "").strip().upper()
    kind = kind if kind in KINDS else "both"
    proj = con.execute("SELECT project_no, name FROM project WHERE project_no = ?", (project,)).fetchone()
    if not proj:
        return None
    realm = _realm(con)
    wip = con.execute(
        "SELECT report_date, total_contract_price, estimated_total_costs, costs_to_date, billed_to_date, "
        "retainage_held FROM v_wip_latest WHERE project_no = ?", (project,)).fetchone()
    loaded = con.execute("SELECT MAX(loaded_at) FROM cost_line WHERE project_no = ?", (project,)).fetchone()[0]
    have = _cols(con, "cost_line")
    opt = {c: (c if c in have else "NULL") for c in
           ("doc_number", "memo", "line_no", "bill_total", "has_attachment")}
    lines: list = []
    if kind in ("costs", "both"):
        sql = (f"SELECT txn_date, vendor, {opt['doc_number']} doc_number, {opt['memo']} memo, description, "
               f"cost_code, account, is_sub, amount, txn_type, qbo_txn_id, qbo_line_id, "
               f"{opt['line_no']} line_no, {opt['bill_total']} bill_total, {opt['has_attachment']} has_attachment "
               f"FROM cost_line WHERE project_no = ? "
               f"ORDER BY txn_date, doc_number, {opt['line_no']}, qbo_txn_id, qbo_line_id")
        for r in con.execute(sql, (project,)):
            ha = r["has_attachment"]
            lines.append({
                "kind": "cost", "date": r["txn_date"], "party": r["vendor"],
                "doc_number": r["doc_number"], "ref": r["doc_number"], "memo": r["memo"],
                "description": r["description"], "cost_code": r["cost_code"], "account": r["account"],
                "is_sub": int(r["is_sub"] or 0), "amount": _f(r["amount"]) or 0.0, "running_total": None,
                "txn_type": r["txn_type"] or "Bill", "txn_id": r["qbo_txn_id"], "line_id": r["qbo_line_id"],
                "qbo_url": qbo_url(r["txn_type"] or "Bill", r["qbo_txn_id"], realm),
                "bill_total": _f(r["bill_total"]),
                "has_attachment": None if ha is None else bool(ha),
            })
    if kind in ("billed", "both"):
        for r in con.execute(
                "SELECT txn_date, customer, doc_number, memo, draw_period, note, amount, balance, "
                "qbo_txn_id, status FROM billing_event WHERE project_no = ? "
                "ORDER BY txn_date, doc_number", (project,)):
            desc = r["draw_period"] or r["note"]
            lines.append({
                "kind": "billed", "date": r["txn_date"], "party": r["customer"],
                "doc_number": r["doc_number"], "ref": r["doc_number"], "memo": r["memo"],
                "description": desc, "cost_code": None, "account": r["status"],
                "is_sub": 0, "amount": _f(r["amount"]) or 0.0, "running_total": None,
                "txn_type": "Invoice", "txn_id": r["qbo_txn_id"], "line_id": None,
                "qbo_url": qbo_url("Invoice", r["qbo_txn_id"], realm),
                "bill_total": None, "has_attachment": None,
            })
    # one chronological list; the running total cumulates WITHIN its kind
    lines.sort(key=lambda d: (d["date"] or "", d["doc_number"] or "", d["txn_id"] or "", d["line_id"] or ""))
    run = {"cost": 0.0, "billed": 0.0}
    for d in lines:
        run[d["kind"]] = round(run[d["kind"]] + d["amount"], 2)
        d["running_total"] = run[d["kind"]]
    costs = run["cost"] if kind != "billed" else _f(con.execute(
        "SELECT COALESCE(SUM(amount),0) FROM cost_line WHERE project_no = ?", (project,)).fetchone()[0])
    billed = run["billed"] if kind != "costs" else _f(con.execute(
        "SELECT COALESCE(SUM(amount),0) FROM billing_event WHERE project_no = ?", (project,)).fetchone()[0])
    wip_c = _f(wip["costs_to_date"]) if wip else None
    wip_b = _f(wip["billed_to_date"]) if wip else None
    wip_ret = (_f(wip["retainage_held"]) or 0.0) if wip else 0.0
    report_date = wip["report_date"] if wip else None
    # Two honest comparisons: (1) WIP billed is GROSS incl. retainage while an invoice is the NET the
    # GC pays - so invoices + retainage held is the like-for-like figure; (2) WIP costs are the
    # report-date cut while QBO is live - lines dated after the report explain most of any gap.
    billed_gross = round((billed or 0.0) + wip_ret, 2)
    after = _f(con.execute(
        "SELECT COALESCE(SUM(amount),0) FROM cost_line WHERE project_no = ? AND txn_date > ?",
        (project, report_date or "9999")).fetchone()[0]) or 0.0
    return {
        "project": project, "name": proj["name"], "kind": kind,
        "as_of": {"qbo_loaded_at": loaded, "wip_report_date": wip["report_date"] if wip else None},
        "budget": {"etc": _f(wip["estimated_total_costs"]) if wip else None,
                   "contract": _f(wip["total_contract_price"]) if wip else None},
        "totals": {"costs": round(costs or 0.0, 2), "billed": round(billed or 0.0, 2),
                   "wip_costs_to_date": wip_c, "wip_billed_to_date": wip_b,
                   "wip_retainage_held": wip_ret, "billed_gross": billed_gross,
                   "costs_after_report": round(after, 2),
                   "delta_costs": None if wip_c is None else round((costs or 0.0) - wip_c, 2),
                   "delta_costs_unexplained": None if wip_c is None else round((costs or 0.0) - after - wip_c, 2),
                   "delta_billed": None if wip_b is None else round(billed_gross - wip_b, 2)},
        "lines": lines,
    }


def to_csv(payload: dict) -> bytes:
    buf = io.StringIO()
    w = csv.DictWriter(buf, fieldnames=CSV_COLS, extrasaction="ignore")
    w.writeheader()
    for d in payload["lines"]:
        w.writerow({k: ("" if d.get(k) is None else d.get(k)) for k in CSV_COLS})
    return ("﻿" + buf.getvalue()).encode("utf-8")


def respond(con, query: dict):
    """(code, body, content-type) for the dashboard handler."""
    project = (query.get("project") or query.get("p") or "").strip()
    kind = (query.get("kind") or "both").strip().lower()
    if not project:
        return 400, json.dumps({"error": "project required"}).encode(), "application/json; charset=utf-8"
    payload = build(con, project, kind)
    if payload is None:
        return 404, json.dumps({"error": f"unknown project {project.upper()}"}).encode(), "application/json; charset=utf-8"
    if query.get("csv") in ("1", "true", "yes"):
        return 200, to_csv(payload), f"text/csv; charset=utf-8"
    return 200, json.dumps(payload, default=str).encode("utf-8"), "application/json; charset=utf-8"


# ── selftest (offline, in-memory) ────────────────────────────────────────────
def _selftest() -> None:
    con = sqlite3.connect(":memory:")
    con.row_factory = sqlite3.Row
    con.executescript("""
    CREATE TABLE project (project_no TEXT PRIMARY KEY, division TEXT, is_ftw INT, name TEXT, type TEXT,
                          builder_or_gc TEXT, bonded INT, rp_category TEXT);
    CREATE TABLE wip_snapshot (project_no TEXT, report_date TEXT, total_contract_price NUM,
                          estimated_total_costs NUM, costs_to_date NUM, billed_to_date NUM, retainage_held NUM);
    CREATE VIEW v_wip_latest AS SELECT s.*, p.division, p.is_ftw, p.name AS project_name, p.type AS project_type,
       p.builder_or_gc, p.bonded, p.rp_category FROM wip_snapshot s JOIN project p ON p.project_no = s.project_no;
    CREATE TABLE cost_line (qbo_txn_id TEXT, qbo_line_id TEXT, txn_type TEXT, project_no TEXT, cost_code TEXT,
                          account TEXT, amount NUM, txn_date TEXT, is_sub INT, vendor TEXT, description TEXT,
                          customer_id TEXT, doc_number TEXT, memo TEXT, line_no INT, bill_total NUM,
                          has_attachment INT, loaded_at TEXT);
    CREATE TABLE billing_event (qbo_txn_id TEXT, doc_number TEXT, project_no TEXT, customer TEXT, memo TEXT,
                          amount NUM, balance NUM, txn_date TEXT, status TEXT, draw_period TEXT, note TEXT);
    CREATE TABLE meta (key TEXT, value TEXT);
    INSERT INTO project VALUES ('CP1','Commercial',0,'TEST JOB',NULL,NULL,0,NULL);
    INSERT INTO wip_snapshot VALUES ('CP1','2026-08-01',100000,80000,300,1000,100);
    INSERT INTO cost_line VALUES ('B1','1','Bill','CP1','SL1',NULL,200,'2026-08-02',0,'VENDOR A','5 sack','c1','4471','CP1 slab',1,300,1,'t');
    INSERT INTO cost_line VALUES ('B1','2','Bill','CP1','SL1',NULL,100,'2026-08-02',0,'VENDOR A',NULL,'c1','4471','CP1 slab',2,300,1,'t');
    INSERT INTO cost_line VALUES ('P1','1','Expense','CP1','SL6',NULL,50,'2026-08-01',1,'VENDOR B','pour','c1','1002','CP1 sub',1,50,NULL,'t');
    INSERT INTO billing_event VALUES ('I1','34001','CP1','CLIENT A','Draw #1',900,0,'2026-08-05','Paid','07/2026',NULL);
    """)
    p = build(con, "cp1", "both")
    assert p and p["name"] == "TEST JOB" and len(p["lines"]) == 4, p
    assert [d["running_total"] for d in p["lines"] if d["kind"] == "cost"] == [50.0, 250.0, 350.0]
    assert p["totals"]["costs"] == 350.0 and p["totals"]["billed"] == 900.0
    assert p["totals"]["delta_costs"] == 50.0 and p["totals"]["delta_billed"] == 0.0     # 900 net + 100 retainage = 1000 gross
    assert p["totals"]["costs_after_report"] == 300.0 and p["totals"]["delta_costs_unexplained"] == -250.0
    assert p["lines"][0]["qbo_url"] == "https://qbo.intuit.com/app/expense?txnId=P1"   # no realm -> bare
    assert p["lines"][1]["has_attachment"] is True and p["lines"][0]["has_attachment"] is None
    con.execute("INSERT INTO meta VALUES ('qbo_realm','9')")
    p2 = build(con, "CP1", "costs")
    assert p2["lines"][1]["qbo_url"].startswith("https://qbo.intuit.com/app/login?pagereq=bill?txnId=B1&deeplinkcompanyid=")
    assert len(p2["lines"]) == 3 and p2["totals"]["billed"] == 900.0
    code, body, ct = respond(con, {"project": "CP1", "csv": "1"})
    assert code == 200 and ct.startswith("text/csv") and body.decode("utf-8-sig").splitlines()[0].startswith("kind,date,party,doc_number")
    assert respond(con, {"project": "ZZ9"})[0] == 404 and respond(con, {})[0] == 400
    print("trail selftest OK - 4 lines, running totals per kind, deltas vs WIP, URL forms, CSV, 404/400")


if __name__ == "__main__":
    if "--selftest" in sys.argv:
        _selftest()
    else:
        print("usage: trail.py --selftest   (served by dashboard.py as /api/trail)")
