"""
rp_price.py - the ONE RP contract / ETC resolver: NEWEST DOCUMENT WINS (the owner 2026-09-17,
after "trust JobTread" moved two contracts down on 2025 proposals while the folder held newer
paper).

For one RP line (a slab job 'RP####' or its flatwork 'RP####-FTW') gather every dated candidate:
  contract  · JobTread's approved proposal (price, its created date)
            · the newest proposal PDF for the scope in the job folder (SUB TOTAL, the DATE printed
              on it - file mtime when the PDF carries none)
  ETC       · JobTread's approved proposal (cost, same date)
            · the takeoff cost sheet for the scope (shared/takeoff_etc, file mtime)
and pick the newest for each. A pick below what is already billed is HELD (the current value
stays) and flagged. Every candidate, its date and the reason it won or lost comes back with the
pick so the change log can say which paper won.

Used by wip/rp_price_sync.py (writes the RP WIP file) and readable by anything that shows the
candidates. Read-only on every source.
"""
from __future__ import annotations

import datetime as dt
import re
import subprocess
from pathlib import Path
from typing import Dict, List, Optional

try:
    from . import takeoff_etc as TK
except ImportError:                                            # pragma: no cover - script path
    import takeoff_etc as TK                                   # type: ignore

_FTW_NAME = re.compile(r"FLATWORK|\bFTW\b|PAVING|DRIVEWAY|SIDEWALK|PATIO|POOL DECK", re.I)
_PROPOSAL_NAME = re.compile(r"proposal|\bbid\b|flatwork|ftw", re.I)
_SKIP_NAME = re.compile(r"invoice|waiver|lien|plan|survey|w-?9|coi\b|insurance|receipt|statement", re.I)
_DATE_RE = re.compile(r"DATE:\s*(\d{1,2})/(\d{1,2})/(\d{4})")
_SUB_RE = re.compile(r"SUB TOTAL:?\s*\$?\s*([\d,]+\.\d{2})", re.I)


def _pdf_text(path: Path) -> str:
    try:
        return subprocess.run(["pdftotext", "-layout", str(path), "-"],       # noqa: S603,S607
                              capture_output=True, text=True, timeout=60).stdout
    except (OSError, subprocess.SubprocessError):
        return ""


def proposal_pdfs(folder: Path, scope: str) -> List[dict]:
    """Every proposal PDF for the scope in the folder (top level only):
    [{path, name, amount, date, dated_by}] newest first. scope 'slab' | 'ftw'."""
    out = []
    try:
        items = [p for p in folder.iterdir() if p.suffix.lower() == ".pdf"]
    except OSError:
        return out
    for p in items:
        n = p.name
        if _SKIP_NAME.search(n) or not _PROPOSAL_NAME.search(n):
            continue
        is_ftw = bool(_FTW_NAME.search(n))
        if (scope == "ftw") != is_ftw:
            continue
        txt = _pdf_text(p)
        m = _SUB_RE.search(txt.upper())
        if not m:
            continue
        amount = float(m.group(1).replace(",", ""))
        d = _DATE_RE.search(txt)
        try:
            date = dt.date(int(d.group(3)), int(d.group(1)), int(d.group(2))) if d else dt.date.fromtimestamp(p.stat().st_mtime)
            dated_by = "the date printed on it" if d else "file date"
        except (ValueError, OSError):
            date, dated_by = dt.date.fromtimestamp(p.stat().st_mtime), "file date"
        out.append({"path": str(p), "name": n, "amount": amount, "date": date, "dated_by": dated_by})
    out.sort(key=lambda x: (x["date"], x["name"]), reverse=True)
    return out


def takeoff_cost(folder: Path, job: str, scope: str) -> Optional[dict]:
    """The takeoff cost sheet's budget for the scope: {path, name, amount, date, note}."""
    path, budget, note, _frag = TK.find_takeoff_etc(folder, job.replace("-FTW", ""), scope, "")
    if not path or budget is None:
        return None
    try:
        date = dt.date.fromtimestamp(Path(path).stat().st_mtime)
    except OSError:
        date = None
    return {"path": str(path), "name": Path(path).name, "amount": float(budget), "date": date, "note": note or ""}


_TEXT_DATE = re.compile(r"(\d{1,2})/(\d{1,2})/(\d{4})")


def source_date(text: Optional[str]) -> Optional[dt.date]:
    """The date of the document a current value came from: a proposal PDF path -> the date printed
    on it (file date when none); a takeoff path -> its file date; text carrying mm/dd/yyyy (the
    'JobTread approved proposal 09/09/2026' notes) -> that date; anything else -> None."""
    t = str(text or "").strip()
    if not t:
        return None
    if t.startswith("/"):
        p = Path(t)
        if not p.exists() or p.is_dir():
            return None
        if p.suffix.lower() == ".pdf":
            d = _DATE_RE.search(_pdf_text(p))
            if d:
                try:
                    return dt.date(int(d.group(3)), int(d.group(1)), int(d.group(2)))
                except ValueError:
                    pass
        try:
            return dt.date.fromtimestamp(p.stat().st_mtime)
        except OSError:
            return None
    m = _TEXT_DATE.search(t)
    if m:
        try:
            return dt.date(int(m.group(3)), int(m.group(1)), int(m.group(2)))
        except ValueError:
            return None
    return None


def sane_etc(amount: Optional[float], contract: Optional[float], current: Optional[float]) -> Optional[str]:
    """Why a machine-read ETC is NOT believable, or None when it is. RP bids carry a margin, so the
    budget sits between a quarter of the contract and the contract; a read that jumps past three
    times the current ETC is a wrong cell, not a re-estimate."""
    if amount is None:
        return "no amount"
    if contract:
        if amount > contract * 1.0001:
            return f"ETC {amount:,.2f} above the contract {contract:,.2f}"
        if amount < contract * 0.25:
            return f"ETC {amount:,.2f} under a quarter of the contract {contract:,.2f}"
    elif amount < 500:
        return f"ETC {amount:,.2f} is not a budget"
    if current and current > 0 and amount > current * 3:
        return f"ETC {amount:,.2f} is more than three times the current {current:,.2f}"
    return None


def resolve(line: str, folder: Optional[Path], jt: Optional[dict], current_contract: Optional[float],
            current_etc: Optional[float], billed: Optional[float],
            current_contract_src: Optional[str] = None, current_etc_src: Optional[str] = None) -> dict:
    """The pick for one line. `jt` = shared.jobtread.jobs()[number] ({id, url, docs}) for the line's
    OWN number (RP####-FTW for a flatwork line), or None. `current_*_src` = the RP file's source
    text / path for the value on file - a documented value only moves for STRICTLY newer paper."""
    scope = "ftw" if line.endswith("-FTW") else "slab"
    cands_k, cands_e, flags = [], [], []
    docs = (jt or {}).get("docs") or []
    jd = None
    if docs:
        price, cost, date = docs[-1]
        try:
            jd = dt.date.fromisoformat(date)
        except (TypeError, ValueError):
            jd = None
        if price > 0:
            cands_k.append({"src": "JobTread approved proposal", "amount": float(price), "date": jd, "detail": f"job {line}"})
        if cost > 0:
            cands_e.append({"src": "JobTread approved proposal", "amount": float(cost), "date": jd, "detail": f"job {line}"})
    elif jt:
        flags.append("JobTread: job exists, no approved proposal")
    else:
        flags.append("not in JobTread")
    pdfs = proposal_pdfs(folder, scope) if folder else []
    multi = len({round(p["amount"]) for p in pdfs}) >= 3
    if pdfs:
        p = pdfs[0]
        cands_k.append({"src": f"proposal PDF {p['name']}", "amount": p["amount"], "date": p["date"], "detail": p["dated_by"]})
    else:
        flags.append("no proposal PDF for this scope in the folder" if folder else "no job folder known")
    tk = takeoff_cost(folder, line, scope) if folder else None
    if tk:
        cands_e.append({"src": f"takeoff cost sheet {tk['name']}", "amount": tk["amount"], "date": tk["date"], "detail": tk["note"]})
    elif folder:
        flags.append("no takeoff cost sheet for this scope in the folder")

    def newest(cands):
        dated = [c for c in cands if c.get("date")]
        return max(dated, key=lambda c: c["date"]) if dated else None
    cur_k_date, cur_e_date = source_date(current_contract_src), source_date(current_etc_src)
    pick_k = newest(cands_k)
    if docs and pdfs and jd and pdfs[0]["date"] and pdfs[0]["date"] > jd:
        flags.append(f"JobTread older than the folder ({jd.strftime('%m/%d/%Y')} vs {pdfs[0]['date'].strftime('%m/%d/%Y')})")
    held = []
    if multi:
        held.append(f"{len(pdfs)} proposals with different amounts in the folder (a multi-scope job) - contract held, sum it by hand")
        pick_k = None
    if pick_k and billed and pick_k["amount"] < billed * 0.99:
        held.append(f"{pick_k['src']} {pick_k['amount']:,.2f} is below what is already billed {billed:,.2f} - contract held")
        pick_k = None
    if pick_k and current_contract and cur_k_date and pick_k["date"] and pick_k["date"] <= cur_k_date and abs(pick_k["amount"] - current_contract) > 0.5:
        held.append(f"the value on file comes from paper dated {cur_k_date.strftime('%m/%d/%Y')}; nothing newer - contract held")
        pick_k = None
    contract_for_sanity = (pick_k["amount"] if pick_k else current_contract)
    # ETC: only believable reads, and JobTread's cost only when JobTread is at least as new as the contract's paper
    ok_e = []
    for c in cands_e:
        why = sane_etc(c["amount"], contract_for_sanity, current_etc)
        if why:
            flags.append(f"{c['src'][:40]}: {why} - not used")
            continue
        if c["src"].startswith("JobTread") and pick_k and not pick_k["src"].startswith("JobTread") and c["date"] and pick_k["date"] and c["date"] < pick_k["date"]:
            continue                                            # older JobTread cost against a newer proposal
        ok_e.append(c)
    pick_e = newest(ok_e)
    if pick_e and current_etc and cur_e_date and pick_e["date"] and pick_e["date"] <= cur_e_date and abs(pick_e["amount"] - current_etc) > 0.5:
        held.append(f"the ETC on file comes from paper dated {cur_e_date.strftime('%m/%d/%Y')}; nothing newer - ETC held")
        pick_e = None
    flags.extend(held)
    return {
        "line": line, "scope": scope, "folder": str(folder) if folder else None,
        "contract": pick_k, "etc": pick_e, "candidates": {"contract": cands_k, "etc": cands_e},
        "current": {"contract": current_contract, "etc": current_etc, "billed": billed,
                    "contract_src_date": cur_k_date, "etc_src_date": cur_e_date},
        "held": "; ".join(held) or None, "flags": flags,
        "changes": {
            "contract": bool(pick_k) and (current_contract is None or abs(pick_k["amount"] - float(current_contract)) > 0.5),
            "etc": bool(pick_e) and (current_etc in (None, 0) or abs(pick_e["amount"] - float(current_etc)) > 0.5),
        },
    }


def fmt_date(d) -> str:
    return d.strftime("%m/%d/%Y") if isinstance(d, dt.date) else ""
