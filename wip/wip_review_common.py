"""
wip_review_common.py - the shared diff/merge core for the ledger's WIP Review.

THE FLOW (the user 2026-08-25: "make the wip update give me the report to
accept/merge ... i can accept costs/billed to date and have pm answer on the
rest ... easy to see the before and afters"):

  1. EMIT  - each WIP tool (cp / rp / master) computes its rows exactly as it
             would before a real write, then instead of writing calls emit_review():
             snapshot the current tab, diff it against the fresh rows FIELD BY
             FIELD, and dump a review JSON. No tab is touched.
  2. REVIEW - the ledger reads those JSONs and shows every change as WAS -> NOW,
             split into the two blocks below. The owner approves/disapproves each.
  3. APPLY  - each tool recomputes, apply_decisions() reverts every DISAPPROVED
             field back to the current tab value (approved fields keep the fresh
             value), and the tool writes through its normal guarded writer.

THE TWO BLOCKS (this is the whole point of the split):
  * QBO   - costs / billed / retainage. Facts from QuickBooks; the owner ACCEPTS.
  * PM    - original contract / approved COs / original ETC / CO costs. Judgment
            the PM answers on.

This module is the ONE place the field set, the tab-header lookup, the diff, and
the revert live, so the CP / RP / Master tabs can never diff or merge differently.
It is pure (no QBO, no Excel writes of its own beyond reading a workbook snapshot),
so it is safe to import from any wip tool.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Dict, List, Optional

# ── the reviewable fields ────────────────────────────────────────────────────
# key         : stable id used in the JSON and the decisions file
# attr        : the CpRow attribute the value is read from / written back to
# block       : "qbo" (owner accepts) or "pm" (PM answers)
# working_hdr : column header on the working tabs (Test - CP / Test - RP)
# master_hdr  : column header on Test-Master (the lean bank view); None = absent
# For MFD, co_revenue is always None, so base_contract == TOTAL CONTRACT PRICE and
# base_etc == ESTIMATED TOTAL COSTS - which is why those two map cleanly to the
# master headers while the CO breakout columns do not exist there.
FIELDS = [
    dict(key="costs",         attr="costs_to_date",    block="qbo", label="Costs to date",
         working_hdr="COSTS TO DATE",         master_hdr="COSTS TO DATE"),
    dict(key="billed",        attr="billed_to_date",   block="qbo", label="Billed to date",
         working_hdr="BILLED TO DATE",        master_hdr="BILLED TO DATE"),
    dict(key="retainage",     attr="retainage_held",   block="qbo", label="Retainage held",
         working_hdr="RETAINAGE HELD",        master_hdr=None),
    dict(key="orig_contract", attr="base_contract",    block="pm",  label="Original contract",
         working_hdr="ORIGINAL CONTRACT",     master_hdr="TOTAL CONTRACT PRICE"),
    dict(key="approved_cos",  attr="co_revenue",       block="pm",  label="Approved COs",
         working_hdr="APPROVED COs",          master_hdr=None),
    dict(key="orig_etc",      attr="base_etc",         block="pm",  label="Original ETC",
         working_hdr="ORIGINAL ESTIMATED COST", master_hdr="ESTIMATED TOTAL COSTS"),
    dict(key="co_costs",      attr="co_cost_override", block="pm",  label="CO costs",
         working_hdr="CO COSTS",              master_hdr=None),
]
FIELD_BY_KEY = {f["key"]: f for f in FIELDS}
_EPS = 0.01                                   # money equal within a cent


def _num(v) -> Optional[float]:
    """Any cell/attr -> float, or None. Tolerates '$1,234', '(123)', '', formulas."""
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if not s or s.startswith("="):
        return None
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()").replace("$", "").replace(",", "").replace("%", "").strip()
    try:
        f = float(s)
    except ValueError:
        return None
    return -f if neg else f


def _changed(old, new) -> bool:
    a, b = _num(old), _num(new)
    if a is None and b is None:
        return False
    if a is None or b is None:
        return True
    return abs(a - b) >= _EPS


# ── read the current tab (the "before") ──────────────────────────────────────
def snapshot_tab(wip_path: Path, tab_name: str, tab_kind: str) -> Dict[str, dict]:
    """{PROJECT# -> {field_key: value, '_name':, '_notes':}} straight off the tab
    as it stands now. tab_kind is 'working' or 'master' (picks the header set).
    Missing file/tab -> {} so a first-ever run reads as all-ADDED."""
    from openpyxl import load_workbook
    wip_path = Path(wip_path)
    if not wip_path.exists():
        return {}
    wb = load_workbook(wip_path, data_only=True, read_only=True)
    if tab_name not in wb.sheetnames:
        wb.close()
        return {}
    ws = wb[tab_name]
    hdr = next((r for r in range(1, 16)
                for c in range(1, (ws.max_column or 0) + 1)
                if ws.cell(r, c).value == "PROJECT #"), None)
    if not hdr:
        wb.close()
        return {}
    idx = {ws.cell(hdr, c).value: c for c in range(1, (ws.max_column or 0) + 1)}
    hdr_key = "master_hdr" if tab_kind == "master" else "working_hdr"
    pcol, ncol = idx.get("PROJECT #"), idx.get("PROJECT NAME")
    out: Dict[str, dict] = {}
    for r in range(hdr + 1, (ws.max_row or 0) + 1):
        first = ws.cell(r, 1).value
        if isinstance(first, str) and first.strip() in ("TOTALS", "TOTAL"):
            break                                  # summary block, not data
        pn = ws.cell(r, pcol).value if pcol else None
        if not pn or not str(pn).strip():
            continue
        rec = {"_name": (ws.cell(r, ncol).value if ncol else None)}
        for f in FIELDS:
            col = idx.get(f[hdr_key]) if f[hdr_key] else None
            rec[f["key"]] = _num(ws.cell(r, col).value) if col else None
        out[str(pn).strip().upper()] = rec
    wb.close()
    return out


def row_value(row, key: str):
    """The fresh value a CpRow carries for a review field (the "after")."""
    return _num(getattr(row, FIELD_BY_KEY[key]["attr"], None))


# ── build the diff (the review payload for one tab) ──────────────────────────
def diff_rows(new_rows: List, prior: Dict[str, dict], *, division: str,
              tab_name: str, tab_kind: str) -> List[dict]:
    """One record per job: status ADDED/REMOVED/CHANGED/SAME and, for each field
    that exists on this tab, its was/now/changed. QBO block first, PM block next."""
    fields_here = [f for f in FIELDS
                   if (f["master_hdr"] if tab_kind == "master" else f["working_hdr"])]
    records, seen = [], set()
    for row in new_rows:
        pn = row.project_num.strip().upper()
        seen.add(pn)
        was = prior.get(pn)
        cells, any_change = [], False
        for f in fields_here:
            now = row_value(row, f["key"])
            old = was.get(f["key"]) if was else None
            ch = _changed(old, now) if was else (now is not None)
            any_change = any_change or ch
            cells.append(dict(key=f["key"], label=f["label"], block=f["block"],
                              was=old, now=now, changed=ch))
        status = "ADDED" if not was else ("CHANGED" if any_change else "SAME")
        records.append(dict(
            project_num=row.project_num, name=(row.project_name or ""),
            division=division, tab=tab_name, tab_kind=tab_kind,
            section=(getattr(row, "section", None) or getattr(row, "rp_type", None) or ""),
            status=status, fields=cells,
            flags="; ".join((getattr(row, "status_flags", None) or [])
                            + (getattr(row, "notes", None) or [])),
        ))
    for pn, was in prior.items():                  # on the tab last run, gone now
        if pn in seen:
            continue
        cells = [dict(key=f["key"], label=f["label"], block=f["block"],
                      was=was.get(f["key"]), now=None, changed=(was.get(f["key"]) is not None))
                 for f in fields_here]
        records.append(dict(
            project_num=pn, name=(was.get("_name") or ""), division=division,
            tab=tab_name, tab_kind=tab_kind, section="", status="REMOVED",
            fields=cells, flags=""))
    order = {"CHANGED": 0, "ADDED": 1, "REMOVED": 2, "SAME": 3}
    records.sort(key=lambda d: (order[d["status"]], d["project_num"]))
    return records


# ── apply the owner's decisions before a real write ──────────────────────────
def apply_decisions(rows: List, decisions: dict) -> List:
    """Return the row list to actually write, honouring the owner's marks.

    decisions = {
      "fields": {"<PN>": {"<field_key>": {"approved": bool, "revert": <num|null>}}},
      "drop_added": ["<PN>", ...],          # ADDED jobs the owner rejected
    }

    A DISAPPROVED field is reverted to `revert` - the exact "was" value the owner
    saw in the review - so every tab reverts the SAME source number and can't
    drift (the lean Master tab folds Contract+COs into one column, so re-reading
    it here would be wrong; the carried value avoids that). An APPROVED or unmarked
    field keeps its freshly computed value. An ADDED job in drop_added is dropped.
    (Reverting the source fields - base_contract, co_revenue - also fixes the
    Master tab's derived TOTAL CONTRACT PRICE / ESTIMATED TOTAL COSTS for free.)"""
    field_dec = (decisions or {}).get("fields", {}) or {}
    drop = {p.upper() for p in ((decisions or {}).get("drop_added", []) or [])}
    kept = []
    for row in rows:
        pn = row.project_num.strip().upper()
        if pn in drop:                              # rejected brand-new job
            continue
        for key, d in field_dec.get(pn, {}).items():
            if key not in FIELD_BY_KEY:
                continue
            approved = d.get("approved", True) if isinstance(d, dict) else bool(d)
            if approved:
                continue                            # approved -> keep fresh value
            revert = d.get("revert") if isinstance(d, dict) else None
            setattr(row, FIELD_BY_KEY[key]["attr"], revert)   # disapproved -> what they saw
        kept.append(row)
    return kept


# ── JSON on disk (what the tools write / the dashboard reads) ────────────────
def write_review_json(out_path: Path, division: str, tab_name: str,
                      records: List[dict]) -> Path:
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    payload = {"division": division, "tab": tab_name, "count": len(records),
               "changed": sum(r["status"] in ("CHANGED", "ADDED", "REMOVED") for r in records),
               "records": records}
    out_path.write_text(json.dumps(payload, indent=2, default=str), encoding="utf-8")
    return out_path


def load_decisions(path: Path) -> dict:
    try:
        return json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}
