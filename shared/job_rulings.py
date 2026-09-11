#!/usr/bin/env python3
"""
job_rulings.py - the standing per-job rulings register (the owner 2026-09-08).

THE REGISTER  <companyhealth_dir>/job_rulings.json   (local business data - never in the repo)
Override: ACB_JOB_RULINGS_FILE in machine.env.

A ruling is a fact about ONE job the owner has settled and does not want
re-discovered on every report: a known loss, an accepted overrun, a scope cut.
Without a register the same finding fires on every WIP run, every P&L and
every QC pass, someone explains it again, and the explanation dies with the
session. The owner: "this needs to be a note for that job permanently so when
we do the WIP report we don't flag, since we know why it went over".

Who reads it (ONE register, so every reader tells the same story):
  wip/rp_wip_reader, wip/cp_wip_reader (and master_wip_test through them)
      -> a `KNOWN: ...` segment in the NOTES column of Test - RP / Test - CP /
         Test-Master. Script-owned: it regenerates from the register each run,
         so retiring a ruling here retires the note on the next sync.
  wip/wip_qc        -> a finding on a job whose ruling `accept`s that check is
                       signed off with the ruling's reason - no sign-off file entry.
  project-pnl       -> the KNOWN LOSSES / RULINGS block on the Job P&L sheet
                       (RP card and the CP/MFD P&L sheet alike).
  ledger            -> `over_budget_accepted` on the project row (the Over budget
                       rule skips it) + the rulings shown on the project page.

Register format:
{
  "_comment": "...",
  "RP6586": [
    {"on": "2026-09-08", "kind": "loss", "amount": 23180,
     "line": "Prepare and pour concrete wall 8\\" wide and 6' tall (lnft) - 61 lnft x $380",
     "source": "Bid Proposal dated 2/4/2026 (GRADE BEAMS)",
     "note": "Wall fell and had to be redone at our cost - the overrun is known",
     "accept": ["OVER_BUDGET"], "by": "the owner"}
  ]
}
  on      the date the ruling was made (YYYY-MM-DD)
  kind    loss | overrun | scope | note      (loss = money we will not get back)
  amount  optional $ the ruling concerns (the bid line, the write-down ...)
  line    the bid / draw line it concerns, as written on the document
  source  the document (proposal date, draw #, CO #)
  note    the plain-English why - this is what every reader prints
  accept  checks this ruling retires: OVER_BUDGET (= OVER_100_PCT +
          COSTS_OVER_CONTRACT), or any wip_qc check name verbatim
  by      a ROLE, never a name (repo names policy)

Project numbers match EXACTLY - RP6586 and RP6586-FTW are two jobs (repo rule);
a ruling on one never leaks to the other.
"""
from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Dict, List, Optional

_HERE = Path(__file__).resolve().parent
if str(_HERE.parent) not in sys.path:
    sys.path.insert(0, str(_HERE.parent))
from shared import paths  # noqa: E402

RULES_FILE: Path = paths.get_path("ACB_JOB_RULINGS_FILE",
                                  paths.companyhealth_dir() / "job_rulings.json")

# A ruling says what it accepts in the owner's words; the QC checks have their
# own names. OVER_BUDGET is the one people mean - it retires BOTH the % past
# 100 and the costs-over-contract finding, because one overrun trips both.
ACCEPT_ALIASES: Dict[str, frozenset] = {
    "OVER_BUDGET": frozenset({"OVER_BUDGET", "OVER_100_PCT", "COSTS_OVER_CONTRACT"}),
}

NOTE_PREFIX = "KNOWN: "        # the NOTES segment every WIP tab carries

_cache: Optional[Dict[str, List[dict]]] = None
_cache_key = None


def _norm_job(proj) -> str:
    return str(proj or "").strip().upper()


def load(path: Optional[Path] = None) -> Dict[str, List[dict]]:
    """{PROJECT# -> [ruling, ...]} - {} when the register does not exist.
    Cached per process by (path, mtime) so a batch does not re-read it per job."""
    global _cache, _cache_key
    p = Path(path) if path else RULES_FILE
    try:
        key = (str(p), p.stat().st_mtime_ns)
    except OSError:
        return {}
    if _cache is not None and _cache_key == key:
        return _cache
    try:
        raw = json.loads(p.read_text(encoding="utf-8"))
    except (ValueError, OSError) as e:
        print(f"  (job rulings unreadable: {p}: {e})", file=sys.stderr)
        return {}
    out: Dict[str, List[dict]] = {}
    for job, rulings in (raw or {}).items():
        if str(job).startswith("_") or not isinstance(rulings, list):
            continue
        out[_norm_job(job)] = [dict(x) for x in rulings if isinstance(x, dict)]
    _cache, _cache_key = out, key
    return out


def for_job(proj, path: Optional[Path] = None) -> List[dict]:
    """The rulings on ONE job - exact project # match, never a family match."""
    return list(load(path).get(_norm_job(proj), ()))


def known_losses(proj, path: Optional[Path] = None) -> List[dict]:
    return [x for x in for_job(proj, path) if str(x.get("kind", "")).lower() == "loss"]


def _accepted_checks(ruling: dict) -> frozenset:
    out = set()
    for a in ruling.get("accept") or ():
        a = str(a).strip().upper()
        out |= ACCEPT_ALIASES.get(a, frozenset({a}))
    return frozenset(out)


def accepting(proj, check: str, path: Optional[Path] = None) -> Optional[dict]:
    """The ruling that retires `check` on this job, or None."""
    c = str(check or "").strip().upper()
    for x in for_job(proj, path):
        if c in _accepted_checks(x):
            return x
    return None


def accepts(proj, check: str, path: Optional[Path] = None) -> bool:
    return accepting(proj, check, path) is not None


def _money(v) -> str:
    try:
        return f"${float(v):,.0f}"
    except (TypeError, ValueError):
        return ""


def summary(ruling: dict) -> str:
    """One line, the way it should read anywhere: the why, then the line and
    the money it concerns, then the document."""
    parts = [str(ruling.get("note") or "").strip()]
    tail = " ".join(s for s in (str(ruling.get("line") or "").strip(),
                                _money(ruling.get("amount"))) if s)
    if tail:
        parts.append(tail)
    if ruling.get("source"):
        parts.append(str(ruling["source"]).strip())
    return " - ".join(p for p in parts if p)


def signoff_record(proj, check: str, path: Optional[Path] = None) -> Optional[dict]:
    """The wip_qc sign-off shape ({reason, by, on}) for a ruling that accepts
    `check`, so QC prints it exactly like a hand-signed finding."""
    x = accepting(proj, check, path)
    if not x:
        return None
    return {"reason": summary(x), "by": x.get("by") or "the owner",
            "on": x.get("on") or "?"}


def note_segments(proj, path: Optional[Path] = None) -> List[str]:
    """The NOTES-column segments for a job: 'KNOWN: <summary>' per ruling."""
    return [NOTE_PREFIX + summary(x) for x in for_job(proj, path)]


def annotate_rows(rows, attr: str = "project_num", path: Optional[Path] = None) -> int:
    """Append each job's KNOWN: segments to row.notes (idempotent) and hang the
    rulings on row.rulings. Returns how many rows got one. Works for any object
    with a project-number attribute and a `notes` list (the WIP CpRow)."""
    n = 0
    for row in rows or ():
        job = getattr(row, attr, None)
        rl = for_job(job, path)
        setattr(row, "rulings", rl)
        if not rl:
            continue
        notes = getattr(row, "notes", None)
        if notes is None:
            continue
        for seg in note_segments(job, path):
            if seg not in notes:
                notes.append(seg)
        n += 1
    return n


def main(argv=None) -> int:
    import argparse
    ap = argparse.ArgumentParser(description="print the job rulings register")
    ap.add_argument("projects", nargs="*", help="project #s (default: every job)")
    ap.add_argument("--file", type=Path, default=None, help="another register file")
    a = ap.parse_args(argv)
    p = a.file or RULES_FILE
    reg = load(p)
    print(f"register: {p}  ({'found' if p.exists() else 'MISSING'}, {len(reg)} job(s))")
    jobs = [_norm_job(x) for x in a.projects] or sorted(reg)
    for job in jobs:
        rl = reg.get(job, [])
        print(f"\n{job}: {len(rl)} ruling(s)")
        for x in rl:
            print(f"  [{x.get('on', '?')}] {str(x.get('kind', 'note')).upper()}: {summary(x)}")
            acc = sorted(_accepted_checks(x))
            if acc:
                print(f"      accepts: {', '.join(acc)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
