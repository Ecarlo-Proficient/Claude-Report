#!/usr/bin/env python3
"""
dashboard.py — a local web dashboard over the project ledger.

Reads the SQLite ledger (READ-ONLY) and serves a single-page dashboard at
http://127.0.0.1:<port>. No terminal, no SQL: portfolio KPIs, per-division
rollups, a searchable / sortable projects table, click-into-a-job detail, and
one-click copy + CSV export. Appearance — theme, font, size, density, content
width, which widgets and which columns show — is customizable in the UI and
saved in the browser (localStorage), per person.

SAFETY
  * The database is opened READ-ONLY (SQLite mode=ro) — the dashboard never writes.
  * The server binds to 127.0.0.1 only — it is not exposed on the network.

USAGE
  python3 ledger/dashboard.py                       # start + open your browser
  python3 ledger/dashboard.py --port 8787           # pick the port
  python3 ledger/dashboard.py --no-open             # don't auto-open a browser
  python3 ledger/dashboard.py --db /path/ledger.sqlite3
"""
from __future__ import annotations

import argparse
import json
import sqlite3
import sys
import threading
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))
from shared import paths  # noqa: E402

HERE = Path(__file__).resolve().parent
STATIC = HERE / "static"

DEFAULT_DB = paths.get_path(
    "ACB_LEDGER_DB",
    Path.home() / "Library" / "Application Support" / "Proficient" / "ledger.sqlite3",
)

CONTENT_TYPES = {
    ".html": "text/html; charset=utf-8",
    ".js": "application/javascript; charset=utf-8",
    ".css": "text/css; charset=utf-8",
    ".svg": "image/svg+xml",
}

# lien states that put a bill on the action watchlist, most-urgent first
LIEN_RANK = {
    "Notice PAST due": 0, "Notice due in ≤7d": 1, "Notice due in ≤15d": 2,
    "Notice due in ≤30d": 3, "Notice Sent": 4, "Lien Filed": 5,
}


def _fetch_ap(con) -> dict:
    """AP + lien view from ap_bill_line; empty (not an error) if the table is absent."""
    ap = {"summary": {"open_balance": 0, "open_lines": 0, "watch_count": 0},
          "lien_watch": [], "by_project": {}}
    try:
        rows = con.execute(
            "SELECT project_no, division, vendor, bill_ref, open_balance, lien_status "
            "FROM ap_bill_line").fetchall()
    except sqlite3.OperationalError:
        return ap
    open_bal, open_lines, watch = 0.0, 0, []
    for r in rows:
        ob = r["open_balance"] or 0
        if ob > 0:
            open_bal += ob
            open_lines += 1
        if r["lien_status"] in LIEN_RANK:
            watch.append(dict(r))
    watch.sort(key=lambda r: (LIEN_RANK[r["lien_status"]], -(r["open_balance"] or 0)))
    by_project = {}
    for r in con.execute("SELECT project_no, open_lines, open_balance FROM v_ap_by_project "
                         "WHERE COALESCE(open_balance,0) > 0"):
        if r["project_no"]:
            by_project[r["project_no"]] = {"open_lines": r["open_lines"], "open_balance": r["open_balance"]}
    ap["summary"] = {"open_balance": open_bal, "open_lines": open_lines, "watch_count": len(watch)}
    ap["lien_watch"] = watch[:500]   # full worklist for the Liens page
    ap["by_project"] = by_project
    return ap


def _fetch_costs(con) -> dict:
    """QBO cost rollups from cost_line; empty (not an error) if unloaded."""
    out = {"by_code": [], "by_project_code": {}, "by_project": {}, "loaded_total": 0}
    try:
        bp = {r["project_no"]: {"costs_loaded": r["costs_loaded"], "sub_costs": r["sub_costs"],
                                "lines": r["lines"]}
              for r in con.execute("SELECT project_no, costs_loaded, sub_costs, lines "
                                   "FROM v_cost_by_project")}
    except sqlite3.OperationalError:
        return out
    out["by_project"] = bp
    out["loaded_total"] = sum((v["costs_loaded"] or 0) for v in bp.values())
    # portfolio, by cost code (join the friendly description)
    for r in con.execute(
            "SELECT vc.code, vc.cost_code, cc.description, SUM(vc.actual) actual, SUM(vc.lines) lines "
            "FROM v_cost_by_code vc LEFT JOIN cost_code cc ON cc.code = vc.cost_code "
            "GROUP BY vc.code, vc.cost_code, cc.description ORDER BY actual DESC"):
        out["by_code"].append({"code": r["code"], "cost_code": r["cost_code"],
                               "description": r["description"], "actual": r["actual"], "lines": r["lines"]})
    # per-project, by cost code (for the job detail drill)
    pc: dict = {}
    for r in con.execute("SELECT project_no, code, cost_code, actual, lines "
                         "FROM v_cost_by_code ORDER BY actual DESC"):
        pc.setdefault(r["project_no"], []).append(
            {"code": r["code"], "cost_code": r["cost_code"], "actual": r["actual"], "lines": r["lines"]})
    out["by_project_code"] = pc

    # grouped: cost TYPE = parent, job TYPE = sub — the JobTread model (material
    # links to ONE cost-type parent; the job-type sub shows cost-to-budget).
    from shared.qbo_costs import cost_code_meta, job_type_name
    groups: dict = {}
    for c in out["by_code"]:
        if c["cost_code"]:
            m = cost_code_meta(c["cost_code"])
            parent = m["description"] or c["cost_code"]
            sub = job_type_name(m["prefix"]) or (m["prefix"] or "—")
        else:                                   # account-based line: no job-type split
            parent = c["code"]
            sub = "(account)"
        g = groups.setdefault(parent, {"parent": parent, "actual": 0, "lines": 0, "subs": {}})
        g["actual"] += c["actual"] or 0
        g["lines"] += c["lines"] or 0
        s = g["subs"].setdefault(sub, {"sub": sub, "code": c["cost_code"], "actual": 0, "lines": 0})
        s["actual"] += c["actual"] or 0
        s["lines"] += c["lines"] or 0
    by_type = []
    for g in groups.values():
        g["subs"] = sorted(g["subs"].values(), key=lambda s: -(s["actual"] or 0))
        by_type.append(g)
    by_type.sort(key=lambda g: -(g["actual"] or 0))
    out["by_cost_type"] = by_type

    # by vendor — the Vendors page (who we pay the most, subs vs suppliers)
    vend = []
    for r in con.execute(
            "SELECT vendor, SUM(amount) spend, COUNT(*) lines, "
            "COUNT(DISTINCT project_no) jobs, "
            "SUM(CASE WHEN is_sub=1 THEN amount ELSE 0 END) sub_spend "
            "FROM cost_line WHERE vendor IS NOT NULL AND vendor <> '' "
            "GROUP BY vendor ORDER BY spend DESC"):
        vend.append({"vendor": r["vendor"], "spend": r["spend"], "lines": r["lines"],
                     "jobs": r["jobs"], "sub_spend": r["sub_spend"]})
    out["by_vendor"] = vend
    return out


def _connect(db_path: Path) -> sqlite3.Connection:
    """Open the ledger READ-ONLY. New connection per request (SQLite + threads)."""
    con = sqlite3.connect(f"file:{db_path}?mode=ro", uri=True)
    con.row_factory = sqlite3.Row
    return con


def fetch_data(db_path: Path) -> dict:
    """Return {meta, projects} from v_wip_latest, or {error} if the db isn't ready."""
    if not db_path.exists():
        return {"error": f"No ledger database at {db_path}. "
                         f"Run:  python3 ledger/load_wip_master.py"}
    try:
        con = _connect(db_path)
    except sqlite3.OperationalError as e:
        return {"error": f"Could not open {db_path}: {e}"}
    try:
        rows = [dict(r) for r in con.execute("SELECT * FROM v_wip_latest")]
        report_date = None
        pcount = len(rows)
        if rows:
            report_date = max((r.get("report_date") or "") for r in rows) or None
        loaded_at = None
        cur = con.execute("SELECT MAX(loaded_at) FROM wip_snapshot")
        got = cur.fetchone()
        if got:
            loaded_at = got[0]
        ap = _fetch_ap(con)
        costs = _fetch_costs(con)
    except sqlite3.OperationalError as e:
        con.close()
        return {"error": f"Ledger schema not found ({e}). Run the loader first."}
    con.close()
    for r in rows:  # attach QBO cost rollup onto each project row
        cp = costs["by_project"].get(r["project_no"])
        r["costs_loaded"] = cp["costs_loaded"] if cp else None
        r["sub_costs"] = cp["sub_costs"] if cp else None
    return {
        "meta": {
            "db_path": str(db_path),
            "report_date": report_date,
            "loaded_at": loaded_at,
            "project_count": pcount,
        },
        "projects": rows,
        "ap": ap,
        "cost": costs,
    }


class Handler(BaseHTTPRequestHandler):
    db_path: Path = DEFAULT_DB

    def log_message(self, *args):  # quiet; no request spam in the terminal
        pass

    def _send(self, code: int, body: bytes, ctype: str):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def _json(self, obj, code: int = 200):
        self._send(code, json.dumps(obj, default=str).encode("utf-8"),
                   "application/json; charset=utf-8")

    def _static(self, name: str):
        # Serve only files that actually live in static/ (no traversal).
        target = (STATIC / name).resolve()
        if STATIC.resolve() not in target.parents or not target.is_file():
            self._send(404, b"not found", "text/plain; charset=utf-8")
            return
        ctype = CONTENT_TYPES.get(target.suffix, "application/octet-stream")
        self._send(200, target.read_bytes(), ctype)

    def do_GET(self):
        path = self.path.split("?", 1)[0]
        if path == "/":
            self._static("index.html")
        elif path == "/api/data":
            self._json(fetch_data(self.db_path))
        elif path == "/api/health":
            self._json({"ok": True})
        elif path.startswith("/static/"):
            self._static(path[len("/static/"):])
        else:
            self._send(404, b"not found", "text/plain; charset=utf-8")


def main():
    ap = argparse.ArgumentParser(description="Local web dashboard over the project ledger.")
    ap.add_argument("--db", type=Path, default=DEFAULT_DB, help="SQLite ledger to read.")
    ap.add_argument("--port", type=int, default=8787, help="Port (default 8787).")
    ap.add_argument("--no-open", action="store_true", help="Don't auto-open a browser.")
    args = ap.parse_args()

    Handler.db_path = args.db
    url = f"http://127.0.0.1:{args.port}"
    httpd = ThreadingHTTPServer(("127.0.0.1", args.port), Handler)

    ready = fetch_data(args.db)
    if "error" in ready:
        print(f"⚠  {ready['error']}")
    else:
        m = ready["meta"]
        print(f"Ledger: {m['project_count']} projects · report {m['report_date']} · {m['db_path']}")
    print(f"Dashboard: {url}   (Ctrl-C to stop)")

    if not args.no_open:
        threading.Timer(0.6, lambda: webbrowser.open(url)).start()
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nStopped.")
        httpd.shutdown()


if __name__ == "__main__":
    main()
