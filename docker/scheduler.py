#!/usr/bin/env python3
"""
scheduler.py - the office server's clock: the ONE process in the Synology container (2026-09-30).

What it runs (the owner's plan, Office server move 09/30/2026):
  mirror     every MIRROR_EVERY_SECONDS (180)   ledger/refresh_mirror.py   - QBO's change feed into the server's
                                                                              own mirror (its own QBO login)
  reconcile  once a night at RECONCILE_AT (02:30) ledger/refresh_mirror.py --reconcile
  AP/AR      every APAR_EVERY_SECONDS (900), AP first then AR, as ONE job:
               AP  bill-tracker/excel_bill_sync.py   -> Bill Tracker.xlsx (Accounting share)
               AR  invoice-sync/run_invoice_sync.py  -> Notion + Teams + Invoice Tracker.xlsx
               (test mode adds payments-test: invoice-sync/qbo_payment_notify.py --test - TEST QuickBooks payment cards)
             AR still runs when AP fails, so Notion stays current.

Two modes (ACB_SERVER_MODE):
  test (default) - the side-by-side week. The trackers are written to TEST_DIR (never the live files), AR runs
                   --dry-run (Notion and Teams untouched) and the Invoice Tracker is rebuilt read-only from Notion
                   by invoice-sync/preview_export.py. Compare the TEST_DIR files with the Mac's each day.
  live           - the server is THE writer, but only while the writer file on the Accounting share names it
                   ({"writer": "server"}). Flip it to "mac" and the server stands down at its next run - that is the
                   rollback. A missing or unreadable writer file stands the server down too (fails closed).

What it reports:
  STATUS_DIR/status.json  - every job's last run, last good run, exit code and the tail of its output; also copied
                            to SHARE_STATUS_FILE (optional) so the Mac ledger can show the server's last good run.
  Teams (TEAMS_WEBHOOK_ALERTS, optional) - a failed run, and no good mirror refresh for ALERT_QUIET_MINUTES.
                            One alert per job per hour at most.

Never a QuickBooks write: the image carries only read code (tests/test_office_server_image.py fails the build
otherwise). Logs go to LOG_DIR (/data/logs), never into the code.
"""
from __future__ import annotations

import datetime as dt
import json
import os
import shutil
import signal
import subprocess
import sys
import time
from pathlib import Path

APP = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(APP))
PY = sys.executable

MODE = (os.getenv("ACB_SERVER_MODE") or "test").strip().lower()
MIRROR_EVERY = int(os.getenv("MIRROR_EVERY_SECONDS", "180"))
APAR_EVERY = int(os.getenv("APAR_EVERY_SECONDS", "900"))
RECONCILE_AT = os.getenv("RECONCILE_AT", "02:30")
MIRROR_STALE_MIN = int(os.getenv("MIRROR_STALE_MINUTES", "30"))
QUIET_ALERT_MIN = int(os.getenv("ALERT_QUIET_MINUTES", "60"))
STATUS_DIR = Path(os.getenv("STATUS_DIR", "/data/status"))
LOG_DIR = Path(os.getenv("LOG_DIR", "/data/logs"))
SHARE_STATUS = os.getenv("SHARE_STATUS_FILE", "").strip()
WRITER_FILE = os.getenv("WRITER_FILE", "").strip()
TEST_DIR = Path(os.getenv("TEST_DIR", "/mnt/accounting/_server-test"))
HOST = os.getenv("ACB_HOST_NAME", "office-server")
TIMEOUTS = {"mirror": 2700, "reconcile": 3600, "ap": 1800, "ar": 1200, "ar-export": 900, "payments-test": 300}

_stop = False
_state: dict = {"mode": MODE, "host": HOST, "jobs": {}}
_alerted: dict = {}


def _now() -> dt.datetime:
    return dt.datetime.now().astimezone()


def _stamp(t: dt.datetime | None = None) -> str:
    return (t or _now()).isoformat(timespec="seconds")


def _say(msg: str) -> None:
    print(f"[{_now():%m/%d/%Y %H:%M:%S}] {msg}", flush=True)


def _write_status() -> None:
    STATUS_DIR.mkdir(parents=True, exist_ok=True)
    _state["written"] = _stamp()
    body = json.dumps(_state, indent=2)
    tmp = STATUS_DIR / "status.json.tmp"
    tmp.write_text(body)
    tmp.replace(STATUS_DIR / "status.json")          # one step: a reader never sees half a file
    if SHARE_STATUS:
        try:
            p = Path(SHARE_STATUS)
            p.parent.mkdir(parents=True, exist_ok=True)
            t2 = p.with_suffix(p.suffix + ".tmp")
            t2.write_text(body)
            t2.replace(p)
        except OSError as e:
            _say(f"status copy to the share failed: {e}")


def _alert(key: str, text: str) -> None:
    """One Teams alert per key per hour - a stuck job must not flood the channel."""
    hook = os.getenv("TEAMS_WEBHOOK_ALERTS", "").strip()
    last = _alerted.get(key)
    if last and (_now() - last).total_seconds() < 3600:
        return
    _alerted[key] = _now()
    _say(f"ALERT {key}: {text}")
    if not hook:
        return
    try:
        from shared import teams_notify
        teams_notify.post(hook, teams_notify._envelope([
            {"type": "TextBlock", "text": f"Office server ({HOST}, {MODE})", "weight": "Bolder"},
            {"type": "TextBlock", "text": text, "wrap": True}]))
    except Exception as e:                           # noqa: BLE001 - an alert must never stop the clock
        _say(f"alert post failed: {e}")


def _run(name: str, argv: list, env_extra: dict | None = None) -> bool:
    """Run one step, keep its output in LOG_DIR/<name>.log, record it in the status file."""
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    log = LOG_DIR / f"{name}.log"
    if log.exists() and log.stat().st_size > 5_000_000:  # keep one old copy, never an endless file
        log.replace(LOG_DIR / f"{name}.log.1")
    env = dict(os.environ)
    env.update(env_extra or {})
    started = _now()
    job = _state["jobs"].setdefault(name, {})
    job["last_run"] = _stamp(started)
    _say(f"{name}: start")
    try:
        p = subprocess.run(argv, cwd=str(APP), env=env, capture_output=True, text=True,
                           timeout=TIMEOUTS.get(name, 1800))
        out, code = (p.stdout or "") + (p.stderr or ""), p.returncode
    except subprocess.TimeoutExpired as e:
        out, code = f"timed out after {e.timeout}s", 124
    with log.open("a") as f:
        f.write(f"\n===== {name} {_stamp(started)} exit {code} =====\n{out}\n")
    job.update({"exit": code, "seconds": round((_now() - started).total_seconds()),
                "tail": out.strip().splitlines()[-6:]})
    if code == 0:
        job["last_ok"] = _stamp()
        _say(f"{name}: ok ({job['seconds']}s)")
    else:
        _say(f"{name}: FAILED exit {code}")
        _alert(name, f"{name} failed (exit {code}): " + " | ".join(job["tail"][-3:]))
    _write_status()
    return code == 0


def _writer() -> str:
    """Who may write the trackers and Notion: the writer file on the share says. Missing = nobody (fail closed)."""
    if not WRITER_FILE:
        return ""
    try:
        return str(json.loads(Path(WRITER_FILE).read_text()).get("writer") or "").strip().lower()
    except (OSError, ValueError):
        return ""


def _mirror_fresh() -> bool:
    ok = _state["jobs"].get("mirror", {}).get("last_ok")
    if not ok:
        return False
    return (_now() - dt.datetime.fromisoformat(ok)).total_seconds() < MIRROR_STALE_MIN * 60


def _seed_test_copies() -> None:
    """The first test run starts from copies of the live trackers, so the Lien / Notes carried forward match."""
    TEST_DIR.mkdir(parents=True, exist_ok=True)
    for live_env, name in (("LIVE_BILL_TRACKER_XLSX", "Bill Tracker.xlsx"), ("LIVE_INVOICE_TRACKER_XLSX", "Invoice Tracker.xlsx")):
        live, test = os.getenv(live_env, "").strip(), TEST_DIR / name
        if live and Path(live).exists() and not test.exists():
            shutil.copy2(live, test)                  # read the live file, never write it
            _say(f"test copy seeded: {name}")


def _apar() -> None:
    if not _mirror_fresh():
        _state["jobs"].setdefault("apar", {})["skipped"] = f"{_stamp()} mirror older than {MIRROR_STALE_MIN} min"
        _alert("apar-stale", f"AP/AR skipped - no good mirror refresh in the last {MIRROR_STALE_MIN} minutes")
        return _write_status()
    if MODE == "live":
        w = _writer()
        if w != "server":
            _state["jobs"].setdefault("apar", {})["skipped"] = f"{_stamp()} writer file says {w or 'nothing'} - standing down"
            if not w:
                _alert("apar-writer", "AP/AR skipped - the writer file is missing or unreadable (the server fails closed)")
            return _write_status()
        _run("ap", [PY, "bill-tracker/excel_bill_sync.py"])
        _run("ar", [PY, "invoice-sync/run_invoice_sync.py"])
    else:
        _seed_test_copies()
        _run("ap", [PY, "bill-tracker/excel_bill_sync.py"], {"ACB_BILL_TRACKER_XLSX": str(TEST_DIR / "Bill Tracker.xlsx")})
        quiet = {"TEAMS_WEBHOOK_MFD_PAID": "", "TEAMS_WEBHOOK_PAYMENTS": "", "ABSORB_NOTES": "0"}   # test: no cards, no notes to Notion
        _run("ar", [PY, "invoice-sync/run_invoice_sync.py", "--dry-run"], quiet)
        _run("ar-export", [PY, "invoice-sync/preview_export.py"], dict(quiet, PREVIEW_EXPORT_PATH=str(TEST_DIR / "Invoice Tracker.xlsx")))
        # the QuickBooks payment cards' test (owner 2026-10-07): "TEST" cards to TEAMS_WEBHOOK_PAYMENTS_TEST, or a dry run
        # listing what would post while that is blank. OFF until the base AP/AR run is proven on the server (owner, same
        # day: "have that part working BEFORE we start adding more") - turn on with PAYMENTS_TEST=1 in server.env.
        if os.getenv("PAYMENTS_TEST", "").strip() == "1":
            _run("payments-test", [PY, "invoice-sync/qbo_payment_notify.py", "--test"])


def _stop_handler(*_a) -> None:
    global _stop
    _stop = True
    _say("stop requested - finishing the current step")


def main() -> int:
    signal.signal(signal.SIGTERM, _stop_handler)
    signal.signal(signal.SIGINT, _stop_handler)
    if MODE not in ("test", "live"):
        _say(f"ACB_SERVER_MODE must be test or live, not {MODE!r}")
        return 2
    _say(f"office server scheduler - mode {MODE}, mirror every {MIRROR_EVERY}s, AP/AR every {APAR_EVERY}s, "
         f"reconcile at {RECONCILE_AT}")
    next_mirror = next_apar = 0.0
    last_reconcile_day = None
    started = _now()
    while not _stop:
        t = time.time()
        if t >= next_mirror:
            _run("mirror", [PY, "ledger/refresh_mirror.py"])
            next_mirror = time.time() + MIRROR_EVERY
        now = _now()
        if now.strftime("%H:%M") >= RECONCILE_AT and last_reconcile_day != now.date():
            _run("reconcile", [PY, "ledger/refresh_mirror.py", "--reconcile"])
            last_reconcile_day = now.date()
        if not _stop and time.time() >= next_apar:
            _apar()
            next_apar = time.time() + APAR_EVERY
        ok = _state["jobs"].get("mirror", {}).get("last_ok")
        quiet_since = dt.datetime.fromisoformat(ok) if ok else started
        if (_now() - quiet_since).total_seconds() > QUIET_ALERT_MIN * 60:
            _alert("quiet", f"no good mirror refresh since {quiet_since:%m/%d/%Y %H:%M}")
        for _ in range(30):                           # wake every 2 s so a stop is quick
            if _stop:
                break
            time.sleep(2)
    _say("stopped")
    return 0


if __name__ == "__main__":
    sys.exit(main())
