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
import datetime as _dt
import hashlib
import json
import re
import sqlite3
import subprocess
import sys
import threading
import time
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))
from shared import paths, pnl_paths  # noqa: E402

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

# ── P&L link (project-pnl) ──────────────────────────────────────────────────
# The dashboard can OPEN an existing per-project P&L and, on an explicit owner
# confirm, RUN project-pnl to (re)generate it. Generation shells out to the tool's
# own CLI (run_pnl.sh) — a subprocess, never an import (tools never import tools).
# QBO stays read-only inside project-pnl; the ONE data write (the .xlsx) is gated
# behind the UI confirm + a confirm flag on the request. Logs land under
# ~/Library/Logs/Proficient/ (never inside the repo).
_PNL_DIR = PROJECT_ROOT / "project-pnl"
_PNL_RUN = _PNL_DIR / "run_pnl.sh"
_PNL_LOG_DIR = Path.home() / "Library" / "Logs" / "Proficient" / "ledger-pnl"
_PROJ_RE = re.compile(r"^(MFD|CP|RP)\d+(-FTW)?$", re.IGNORECASE)
_PNL_JOBS: dict = {}                 # proj -> {state, started, log, detail, proc, file}
_PNL_LOCK = threading.Lock()


def _pnl_wait(proj: str) -> None:
    """Reap a generation subprocess and record its outcome (running → done/error)."""
    j = _PNL_JOBS.get(proj)
    if not j:
        return
    rc = j["proc"].wait()
    try:
        j["file"].close()
    except OSError:
        pass
    with _PNL_LOCK:
        j["state"] = "done" if rc == 0 else "error"
        j["rc"] = rc
        if rc != 0:
            j["detail"] = f"exit {rc} — see {j['log']}"


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
            "SELECT project_no, division, vendor, bill_ref, open_balance, lien_status, "
            "matched_invoice, invoice_no FROM ap_bill_line").fetchall()
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
    # vendor TYPE — Sub (labor) vs Supplier: <material> — from each vendor's cost mix.
    mix: dict = {}
    for r in con.execute("SELECT vendor, cost_code, account, SUM(amount) amt FROM cost_line "
                         "WHERE vendor IS NOT NULL AND vendor <> '' GROUP BY vendor, cost_code, account"):
        parent = cost_code_meta(r["cost_code"])["description"] if r["cost_code"] else None
        parent = parent or r["account"] or "Materials"
        mix.setdefault(r["vendor"], {})
        mix[r["vendor"]][parent] = mix[r["vendor"]].get(parent, 0) + (r["amt"] or 0)
    for v in vend:
        spend = v["spend"] or 0
        sub_share = (v["sub_spend"] or 0) / spend if spend else 0
        vm = mix.get(v["vendor"], {})
        top = max(vm, key=vm.get) if vm else None
        v["vtype"] = "Sub" if sub_share >= 0.5 else (f"Supplier: {top}" if top else "Supplier")
    out["by_vendor"] = vend
    return out


def _freshness(con) -> dict:
    """When each feed last landed (ledger loaded_at) + when each SOURCE sync wrote
    its file (mtime) — the owner's "is my data current?" strip."""
    import os
    out = {"ledger": {}, "sources": {}}
    for tbl, key in (("wip_snapshot", "WIP"), ("ap_bill_line", "AP (Bill Tracker)"),
                     ("cost_line", "Costs (QBO)")):
        try:
            r = con.execute(f"SELECT MAX(loaded_at) FROM {tbl}").fetchone()
            out["ledger"][key] = r[0] if r and r[0] else None
        except sqlite3.OperationalError:
            out["ledger"][key] = None

    def mtime(p):
        try:
            return _dt.datetime.fromtimestamp(os.path.getmtime(str(p))).isoformat(timespec="minutes")
        except OSError:
            return None

    ob = paths.onedrive_base()
    bt = paths.get_path("ACB_BILL_TRACKER_XLSX", ob / "Automations-/Bill Tracker.xlsx")
    wm = paths.get_path("WIP_EXCEL_PATH", ob / "Company Files - WIP Report/WIP - MASTER new.xlsx")
    out["sources"]["sync-ap"] = mtime(bt)
    out["sources"]["WIP master"] = mtime(wm)
    for cand in (ob / "Automations-/Open_Invoices.xlsx", ob / "Open_Invoices.xlsx",
                 ob / "Automations-/Collections/Open_Invoices.xlsx"):
        m = mtime(cand)
        if m:
            out["sources"]["sync-ar"] = m
            break
    return out


def _fetch_actions(con) -> dict:
    """Map action_key → {url, status} from the `action` table (Notion mirror).
    Empty (not an error) if sync_actions hasn't run."""
    out: dict = {}
    try:
        for r in con.execute("SELECT action_key, status, notion_url FROM action"):
            out[r["action_key"]] = {"status": r["status"], "url": r["notion_url"]}
    except sqlite3.OperationalError:
        pass
    return out


def _waiver_key(mi, vendor, bill_ref) -> str:
    """Deterministic key for a bill's waiver — survives ap_bill_line reloads."""
    raw = f"{mi or ''}|{vendor or ''}|{bill_ref or ''}"
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()[:16]


# race-through stages, in worklist priority — Ready-to-turn-in first (turn it in
# NOW to unlock the next draw), then pay vendors, then collect waivers.
_STAGE_ORDER = {"Ready to turn in": 0, "Fund in — pay vendors": 1,
                "Paid — collect waivers": 2, "Awaiting GC funding": 3}


def _fetch_draws(con, limit: int = 100) -> dict:
    """Roll AP bills up BY DRAW (matched invoice) → the race-through pipeline,
    joined to the owner's waiver marks. Empty (not an error) if not loaded."""
    try:
        rows = con.execute(
            "SELECT matched_invoice mi, project_no, division, vendor, bill_ref, "
            "MAX(bill_total) amount, MAX(open_balance) open_bal, pay_status, invoice_status, "
            "MAX(gc_paid_date) gc, MAX(pay_date) pd, MAX(bill_date) bd "
            "FROM ap_bill_line WHERE matched_invoice IS NOT NULL AND matched_invoice <> '' "
            # RP isn't draws — RP bills at completion / milestones, not formal draws (owner).
            "AND COALESCE(project_no,'') NOT LIKE 'RP%' AND matched_invoice NOT LIKE '%— RP%' "
            "GROUP BY matched_invoice, vendor, bill_ref").fetchall()
    except sqlite3.OperationalError:
        return {"draws": [], "total": 0}
    wmap = {w["waiver_key"]: w["received"] for w in con.execute("SELECT waiver_key, received FROM waiver")}
    draws: dict = {}
    for r in rows:
        d = draws.setdefault(r["mi"], {"matched_invoice": r["mi"], "project_no": r["project_no"],
                                       "division": r["division"], "bills": []})
        wk = _waiver_key(r["mi"], r["vendor"], r["bill_ref"])
        d["bills"].append({
            "vendor": r["vendor"], "bill_ref": r["bill_ref"], "amount": r["amount"] or 0,
            "open": r["open_bal"] or 0, "pay_status": r["pay_status"], "invoice_status": r["invoice_status"],
            "gc_paid": r["gc"], "pay_date": r["pd"], "bill_date": r["bd"],
            "waiver_key": wk, "waiver": bool(wmap.get(wk, 0)),
        })
    out = []
    for mi, d in draws.items():
        bills = d["bills"]
        n = len(bills)
        paid = sum(1 for b in bills if b["pay_date"])
        funded = any(b["gc_paid"] for b in bills)
        waivers = sum(1 for b in bills if b["waiver"])
        if not funded:
            stage = "Awaiting GC funding"
        elif paid < n:
            stage = "Fund in — pay vendors"
        elif waivers < n:
            stage = "Paid — collect waivers"
        else:
            stage = "Ready to turn in"
        d.update({
            "label": (mi or "").split("\n")[0].strip(), "n": n, "paid": paid, "funded": funded,
            "waivers": waivers, "total": sum(b["amount"] for b in bills), "stage": stage,
            "recency": max([(b["gc_paid"] or b["pay_date"] or b["bill_date"] or "") for b in bills] or [""]),
        })
        out.append(d)
    out.sort(key=lambda d: d["recency"], reverse=True)
    out.sort(key=lambda d: _STAGE_ORDER.get(d["stage"], 9))
    return {"draws": out[:limit], "total": len(out)}


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
        draws = _fetch_draws(con)
        actions = _fetch_actions(con)
        freshness = _freshness(con)
    except sqlite3.OperationalError as e:
        con.close()
        return {"error": f"Ledger schema not found ({e}). Run the loader first."}
    con.close()
    for r in rows:  # attach QBO cost rollup onto each project row
        cp = costs["by_project"].get(r["project_no"])
        r["costs_loaded"] = cp["costs_loaded"] if cp else None
        r["sub_costs"] = cp["sub_costs"] if cp else None
    for d in draws.get("draws", []):  # attach the Notion action link (if tracked)
        inv = (d.get("matched_invoice") or "").split("—")[0].strip()
        d["action"] = actions.get(f"draw:{inv}")
    return {
        "meta": {
            "db_path": str(db_path),
            "report_date": report_date,
            "loaded_at": loaded_at,
            "project_count": pcount,
            "freshness": freshness,
        },
        "projects": rows,
        "ap": ap,
        "cost": costs,
        "draws": draws,
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

    def _query(self) -> dict:
        return {k: v[0] for k, v in parse_qs(urlparse(self.path).query).items()}

    def do_GET(self):
        path = self.path.split("?", 1)[0]
        if path == "/":
            self._static("index.html")
        elif path == "/api/data":
            self._json(fetch_data(self.db_path))
        elif path == "/api/health":
            self._json({"ok": True})
        elif path == "/api/pnl":
            self._pnl_find(self._query().get("proj", ""))
        elif path == "/api/pnl/status":
            self._pnl_status(self._query().get("proj", ""))
        elif path.startswith("/static/"):
            self._static(path[len("/static/"):])
        else:
            self._send(404, b"not found", "text/plain; charset=utf-8")

    def do_POST(self):
        p = urlparse(self.path).path
        if p == "/api/waiver":            # the ONE ledger write — the owner's waiver marks
            self._set_waiver()
        elif p == "/api/pnl/open":        # open an existing P&L workbook (local `open`)
            self._pnl_open(self._query().get("proj", ""))
        elif p == "/api/pnl/generate":    # run project-pnl (gated by an explicit confirm)
            self._pnl_generate()
        else:
            self._send(404, b"not found", "text/plain; charset=utf-8")

    # ── P&L link handlers ───────────────────────────────────────────────────
    def _pnl_find(self, proj: str):
        proj = (proj or "").strip().upper()
        if not _PROJ_RE.match(proj):
            return self._json({"error": "bad or missing project"}, 400)
        info = pnl_paths.find_pnl(proj)
        j = _PNL_JOBS.get(proj)
        info["job"] = j["state"] if j else "idle"
        self._json(info)

    def _pnl_status(self, proj: str):
        proj = (proj or "").strip().upper()
        j = _PNL_JOBS.get(proj)
        if not j:
            return self._json({"state": "idle"})
        out = {"state": j["state"]}
        if j["state"] == "running":
            out["elapsed"] = int(time.time() - j["started"])
        if j.get("detail"):
            out["detail"] = j["detail"]
        self._json(out)

    def _pnl_open(self, proj: str):
        proj = (proj or "").strip().upper()
        if not _PROJ_RE.match(proj):
            return self._json({"error": "bad project"}, 400)
        info = pnl_paths.find_pnl(proj)
        if not info.get("exists"):
            return self._json({"error": "no P&L generated yet"}, 404)
        path = Path(info["path"])
        if path.name != f"Project_PnL_{proj}.xlsx":   # only ever open the resolved workbook
            return self._json({"error": "unexpected file"}, 400)
        try:
            subprocess.Popen(["open", str(path)])     # macOS: open in Excel
        except OSError as e:
            return self._json({"error": f"open failed: {e}"}, 500)
        self._json({"ok": True, "path": str(path)})

    def _pnl_generate(self):
        try:
            length = int(self.headers.get("Content-Length", 0))
            body = json.loads(self.rfile.read(length) or b"{}")
        except (ValueError, json.JSONDecodeError):
            return self._json({"error": "bad request"}, 400)
        proj = str(body.get("proj", "")).strip().upper()
        if not _PROJ_RE.match(proj):
            return self._json({"error": "bad or missing project"}, 400)
        if not body.get("confirm"):                   # gate: no confirm, no run
            return self._json({"error": "confirm required"}, 400)
        if not _PNL_RUN.exists():
            return self._json({"error": "project-pnl runner not found"}, 500)
        with _PNL_LOCK:
            j = _PNL_JOBS.get(proj)
            if j and j["state"] == "running":
                return self._json({"state": "running", "proj": proj, "already": True})
            try:
                _PNL_LOG_DIR.mkdir(parents=True, exist_ok=True)
                f = open(_PNL_LOG_DIR / f"{proj}.log", "w")
                proc = subprocess.Popen(
                    ["/bin/bash", str(_PNL_RUN), proj],
                    cwd=str(_PNL_DIR), stdout=f, stderr=subprocess.STDOUT)
            except OSError as e:
                return self._json({"error": f"spawn failed: {e}"}, 500)
            _PNL_JOBS[proj] = {"state": "running", "started": time.time(),
                               "log": str(_PNL_LOG_DIR / f"{proj}.log"),
                               "proc": proc, "file": f}
        threading.Thread(target=_pnl_wait, args=(proj,), daemon=True).start()
        self._json({"state": "running", "proj": proj})

    def _set_waiver(self):
        try:
            length = int(self.headers.get("Content-Length", 0))
            body = json.loads(self.rfile.read(length) or b"{}")
        except (ValueError, json.JSONDecodeError):
            return self._json({"error": "bad request"}, 400)
        mi, vendor, bill = body.get("matched_invoice"), body.get("vendor"), body.get("bill_ref")
        received = 1 if body.get("received") else 0
        if not mi:
            return self._json({"error": "matched_invoice required"}, 400)
        key = _waiver_key(mi, vendor, bill)
        now = _dt.datetime.now().isoformat(timespec="seconds")
        try:
            con = sqlite3.connect(self.db_path)          # WRITABLE (the one write surface)
            con.execute("CREATE TABLE IF NOT EXISTS waiver (waiver_key TEXT PRIMARY KEY, "
                        "matched_invoice TEXT, vendor TEXT, bill_ref TEXT, received INTEGER NOT NULL "
                        "DEFAULT 0, received_date TEXT, note TEXT, updated_at TEXT NOT NULL)")
            con.execute(
                "INSERT INTO waiver (waiver_key, matched_invoice, vendor, bill_ref, received, "
                "received_date, updated_at) VALUES (?,?,?,?,?,?,?) "
                "ON CONFLICT(waiver_key) DO UPDATE SET received=excluded.received, "
                "received_date=excluded.received_date, updated_at=excluded.updated_at",
                (key, mi, vendor, bill, received, now if received else None, now))
            con.commit(); con.close()
        except sqlite3.OperationalError as e:
            return self._json({"error": f"write failed: {e}"}, 500)
        self._json({"ok": True, "received": bool(received), "waiver_key": key})


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
