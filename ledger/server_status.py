"""ledger/server_status.py - the office server as the ledger watches it (10/09/2026).

The owner, the day the Synology went live: "now that we have the live, make sure we can
monitor action". The server's scheduler (docker/scheduler.py) writes
`Accounting/_automation/server-status.json` after every step - mode, host, every job's last
run / last good run / exit code / seconds / the tail of its output, and why AP/AR was skipped
- and the Mac honours `Accounting/_automation/writer.json` (who may write the trackers). This
module reads both and turns them into ONE judgement the ledger can show and colour:

  ok             every job's latest run exited 0, the heartbeat is fresh, the server writes
  late           the status file has not been rewritten for a while (the mirror runs every 3 min,
                 so 20 quiet minutes means a job is hanging or the clock stopped)
  down           no rewrite for an hour: the container is stopped or the share is cut off
  failing        a job's latest run exited non-zero (its tail says why)
  standing down  live mode, but the writer file does not name the server (rollback to the Mac)
                 or the mirror went stale so AP/AR skipped
  off            no status file (share not mounted here, or the server never ran)

Pure logic in `assess()` (tests/test_server_status.py); `read()` does the two file reads.
Never writes. Nothing here is a QuickBooks call.
"""
from __future__ import annotations

import datetime as dt
import json
from pathlib import Path
from typing import Optional

STATUS_FILE = "server-status.json"
WRITER_FILE = "writer.json"
LATE_MIN = 20          # heartbeat older than this = late (amber)
DOWN_MIN = 60          # ... than this = down (red)
JOB_WORDS = {"mirror": "QuickBooks copy", "reconcile": "Nightly QuickBooks count check", "ap": "Bill Tracker run",
             "ar": "Invoice sync run", "ar-export": "Invoice Tracker (test copy)", "payments-test": "Payment cards (test)",
             "apar": "AP / AR round"}
JOB_ORDER = ["mirror", "ap", "ar", "reconcile", "ar-export", "payments-test", "apar"]


def _parse(ts: Optional[str]) -> Optional[dt.datetime]:
    if not ts:
        return None
    try:
        t = dt.datetime.fromisoformat(ts)
    except ValueError:
        return None
    return t if t.tzinfo else t.replace(tzinfo=dt.datetime.now().astimezone().tzinfo)


def _minutes(a: Optional[dt.datetime], b: dt.datetime) -> Optional[float]:
    return None if a is None else round((b - a).total_seconds() / 60, 1)


def assess(status: Optional[dict], writer: Optional[str], now: Optional[dt.datetime] = None) -> dict:
    """status = the parsed status.json (None = no file), writer = writer.json's "writer" (None =
    unreadable / missing). Returns the judgement + a flat job list the page renders."""
    now = now or dt.datetime.now().astimezone()
    if not status:
        return {"present": False, "health": "off", "reason": "no status file - the Accounting share is not mounted here, or the server has not run",
                "mode": None, "writer": writer, "jobs": []}
    mode = (status.get("mode") or "").lower()
    written = _parse(status.get("written"))
    age = _minutes(written, now)
    jobs_in = status.get("jobs") or {}
    jobs = []
    failing = []
    for name in sorted(jobs_in, key=lambda n: (JOB_ORDER.index(n) if n in JOB_ORDER else 99, n)):
        j = jobs_in.get(name) or {}
        exit_code = j.get("exit")
        last_run, last_ok = _parse(j.get("last_run")), _parse(j.get("last_ok"))
        ok_now = exit_code == 0 if exit_code is not None else (last_ok is not None and last_ok == last_run)
        row = {"job": name, "label": JOB_WORDS.get(name, name), "last_run": j.get("last_run"), "last_ok": j.get("last_ok"),
               "exit": exit_code, "seconds": j.get("seconds"), "ok": bool(ok_now) if (last_run or exit_code is not None) else None,
               "skipped": j.get("skipped"), "tail": list(j.get("tail") or [])[-6:],
               "ran_min_ago": _minutes(last_run, now), "ok_min_ago": _minutes(last_ok, now)}
        jobs.append(row)
        if exit_code not in (None, 0):
            failing.append(name)
    health, reason = "ok", "every job's latest run finished clean"
    skipped = (jobs_in.get("apar") or {}).get("skipped") or ""
    if age is None:
        health, reason = "down", "the status file carries no time"
    elif age > DOWN_MIN:
        health, reason = "down", f"no word from the server for {int(age)} minutes - container stopped, or the share is cut off"
    elif failing:
        health, reason = "failing", "failed: " + ", ".join(JOB_WORDS.get(n, n) for n in failing)
    elif age > LATE_MIN:
        health, reason = "late", f"last word {int(age)} minutes ago (the copy refreshes every 3 minutes)"
    elif mode == "live" and (writer or "") != "server":
        health = "standing down"
        reason = (f"writer file says \"{writer}\" - the server is not writing the trackers" if writer
                  else "writer file missing or unreadable - the server fails closed (nothing written)")
    elif skipped and _skip_is_current(skipped, jobs_in, now):
        health, reason = "standing down", "AP / AR skipped: " + skipped.split(" ", 1)[-1]
    return {"present": True, "health": health, "reason": reason, "mode": mode, "host": status.get("host"),
            "writer": writer, "written": status.get("written"), "age_min": age, "jobs": jobs}


def _skip_is_current(skipped: str, jobs: dict, now: dt.datetime) -> bool:
    """The apar skip note carries its own stamp; it counts only when it is newer than the last AP
    run (the scheduler never clears the note) and within the last round (15 min + slack)."""
    stamp = _parse(skipped.split(" ", 1)[0])
    if stamp is None:
        return False
    ap_last = _parse((jobs.get("ap") or {}).get("last_run"))
    if ap_last and ap_last > stamp:
        return False
    return (now - stamp).total_seconds() < 40 * 60


def read(automation_dir: Path, now: Optional[dt.datetime] = None) -> dict:
    """Read `_automation/server-status.json` + `writer.json` under the Accounting share and assess."""
    status = writer = None
    try:
        status = json.loads((automation_dir / STATUS_FILE).read_text())
    except (OSError, ValueError):
        status = None
    try:
        writer = str(json.loads((automation_dir / WRITER_FILE).read_text()).get("writer") or "").strip().lower() or None
    except (OSError, ValueError, AttributeError):
        writer = None
    out = assess(status if isinstance(status, dict) else None, writer, now)
    out["dir"] = str(automation_dir)
    return out
