"""schedule_index.py - job # -> the LAST day it appeared on a daily crew schedule.

The owner (2026-09-09): the RP tab's CATEGORY must say where a line comes from and
the last date it was on that source - "Schedule 8-31-26" - not "GOOD". The daily
schedules are one file per day under OPERATIONS/SCHEDULE/<year>/<month>/, so the
answer needs every 'Main Schedule' tab ever written. Parsing them all takes
minutes; this module parses each file ONCE and keeps a JSON cache keyed by the
file's path and mtime, so a normal run only reads the new days.

  last_seen(sched_dir) -> {"RP7358": date(2026, 8, 20), "RP7358-FTW": ..., ...}

A job on a FLATWORK band row (or a row whose stage says flatwork) is indexed as
both the base job and its -FTW line, so a flatwork line finds its own date. The
cache lives outside the repo (Application Support); override with
ACB_SCHEDULE_INDEX. Read-only on the schedules. A missing share gives {} plus
whatever the cache already knows.
"""
from __future__ import annotations

import datetime as dt
import json
import os
import re
from pathlib import Path
from typing import Dict, Optional

from shared import schedule as SCH

CACHE = Path(os.environ.get(
    "ACB_SCHEDULE_INDEX",
    Path.home() / "Library" / "Application Support" / "Proficient" / "schedule-index.json"))
_FTW_RE = re.compile(r"FLATWORK|\bFTW\b|PAVING|DRIVEWAY|SIDEWALK|PATIO|POOL DECK", re.I)


def _load(cache: Path) -> dict:
    try:
        d = json.loads(cache.read_text(encoding="utf-8"))
        return d if isinstance(d, dict) and isinstance(d.get("files"), dict) else {"files": {}}
    except (OSError, ValueError):
        return {"files": {}}


def _jobs_in(path: Path) -> list:
    """Every job # on one schedule, flatwork rows doubled as -FTW lines."""
    out = set()
    for rec in SCH.parse_main_schedule(path):
        job = (rec.get("proj") or "").upper()
        if not job:
            continue
        out.add(job)
        text = f"{rec.get('section', '')} {rec.get('stage', '')}"
        if _FTW_RE.search(text) and not job.endswith("-FTW"):
            out.add(job + "-FTW")
    return sorted(out)


def last_seen(sched_dir: Optional[Path] = None, cache: Path = CACHE,
              refresh: bool = True) -> Dict[str, dt.date]:
    """{job# -> last schedule date it appeared on}, across every schedule file.
    `refresh=False` answers from the cache alone (no share access)."""
    idx = _load(cache)
    files = idx["files"]
    changed = False
    if refresh:
        for date, f in SCH.all_schedule_files(sched_dir):
            key = str(f)
            try:
                mtime = f.stat().st_mtime
            except OSError:
                continue
            rec = files.get(key)
            if rec and rec.get("mtime") == mtime:
                continue
            files[key] = {"mtime": mtime, "date": date.isoformat(), "jobs": _jobs_in(f)}
            changed = True
    if changed:
        try:
            cache.parent.mkdir(parents=True, exist_ok=True)
            cache.write_text(json.dumps(idx, indent=0), encoding="utf-8")
        except OSError:
            pass
    out: Dict[str, dt.date] = {}
    for rec in files.values():
        try:
            d = dt.date.fromisoformat(rec["date"])
        except (KeyError, ValueError):
            continue
        for job in rec.get("jobs", []):
            if job not in out or d > out[job]:
                out[job] = d
    return out


def label(d: Optional[dt.date]) -> str:
    """'8-31-26' - the schedule file's own naming."""
    return f"{d.month}-{d.day}-{d.year % 100:02d}" if d else ""
