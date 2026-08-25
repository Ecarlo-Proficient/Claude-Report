"""
shared/cost_code_audit.py — THE vendor cost-code coding-family audit.

Vendors must code their bill lines to the RIGHT cost-code FAMILY for what they
sell (the owner 2026-08-25). The cost-code NUMBER is the family (see the Cost Code
Sheet + shared/qbo_costs): 1 Concrete (ready-mix) · 2 Rebar · 3 Formwork/Lumber ·
4 Aggregates · 5 Equip/51 Pump/52 Saw · 6 Labor · 7 Specialty · 8 Fuel · 9 Supplies.

Vendor coding TYPES + their rule:
  • concrete supplier (ready-mix, e.g. Cowtown) → every line must be *1.
  • material supplier (e.g. RCI = lumber/rebar) → *2/*3/*4 only; NEVER *1
    (concrete), *5/*51/*52 (equipment), or *6 (labor).
  • both (e.g. Preferred Materials) → sells concrete AND material, so a line whose
    MEMO reads as concrete yardage / ready-mix MUST be *1, never another code.

The type is CAPTURED from each vendor's *1-vs-*2/3/4 split; an override JSON forces
it. This module is PURE (no QBO/IO) so it unit-tests offline and is shared by BOTH
`one-offs/concrete_cost_code_audit.py` (standalone QBO pull → Excel) and the
bill-tracker's `Audit - Cost Code` sheet (folded into sync-ap).

Row shape the functions expect (build it from whatever source):
  {vendor, number, cost_code, cost_name, desc, account, …passthrough}
`number`/`cost_name` come from `code_families(raw_item_name)`.
"""
from __future__ import annotations

import json
import re
from collections import defaultdict
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from shared.qbo_costs import cost_code_meta

# Cost-code number families.
CONCRETE = {"1"}
MATERIAL = {"2", "3", "4"}                  # rebar · lumber/formwork · aggregates
EQUIP = {"5", "51", "52"}
LABOR = {"6"}
# What a material-only vendor (e.g. RCI) must NEVER carry.
MATERIAL_FORBIDDEN = CONCRETE | EQUIP | LABOR

# A line memo that reads as CONCRETE YARDAGE — concrete is ordered by cubic yard
# or sack mix. For a both-supplier such a line must be coded *1.
_CONCRETE_MEMO_RE = re.compile(
    r"\b(?:\d+(?:\.\d+)?\s*(?:C\.?Y\.?|CU\.?\s*YD|YDS?|YARDS?|CUBIC)"
    r"|SACKS?|SCK|SACK\s*MIX|CONCRETE|READY[\s-]*MIX|REDI[\s-]*MIX|REDIMIX)\b",
    re.IGNORECASE)

TYPE_LABEL = {"concrete": "Concrete", "material": "Material", "both": "Both (conc+mat)",
              "hauler": "Hauler (haul-off OK)", "review": "Review - possible"}
TYPE_ORDER = {"concrete": 0, "material": 1, "both": 2, "hauler": 3, "review": 4}
# Override / forced-type keys (in priority order); "exclude" drops a vendor.
OVERRIDE_TYPES = ("concrete", "material", "both", "hauler")


def concrete_memo(desc: str) -> bool:
    """True if a line description reads as concrete yardage / ready-mix."""
    return bool(_CONCRETE_MEMO_RE.search(desc or ""))


def code_families(cost_code: str) -> Tuple[Optional[str], Optional[str]]:
    """Leaf cost code → (number, cost_name). 'Parent:SL1' → ('1', 'Concrete');
    a non-code → (None, None)."""
    leaf = (cost_code or "").split(":")[-1].strip()
    if not leaf:
        return None, None
    m = cost_code_meta(leaf)
    return m["number"], m["description"]


def po_origin(bill_number: Optional[str], po_numbers, po_found: bool) -> str:
    """Where a miscode came from: compare the bill line's cost-code family to the
    family numbers on its linked PO. Answers "is the clerk just trusting a wrong
    PO from the super/PM?" (the user 2026-08-25).
      • not po_found            → the bill has no PO (clerk coded it standalone)
      • po_numbers empty        → the PO carries no cost code to compare
      • bill_number in po_nums  → the PO ALSO carries this code → upstream (super/PM)
      • else                    → the PO's code differs → the bill deviated
    """
    if not po_found:
        return "No PO on the bill (clerk-coded)"
    nums = [n for n in (po_numbers or []) if n]
    if not nums:
        return "PO has no cost code"
    if bill_number is not None and bill_number in nums:
        return f"PO also codes *{bill_number} - upstream (super/PM)"
    return f"PO codes {'/'.join('*' + n for n in nums)} - bill deviated from PO"


def load_override(path: Path) -> Dict[str, set]:
    """JSON {"concrete":[…], "material":[…], "both":[…], "hauler":[…],
    "exclude":[…]} → dict of upper-cased name sets. Missing/broken file → empties.
    A forced type wins over auto-detection; "exclude" drops the vendor."""
    keys = OVERRIDE_TYPES + ("exclude",)
    out = {k: set() for k in keys}
    if not path or not Path(path).exists():
        return out
    try:
        data = json.loads(Path(path).read_text())
    except (OSError, json.JSONDecodeError):
        return out
    for k in keys:
        out[k] = {str(x).strip().upper() for x in data.get(k, [])}
    return out


def classify_vendors(rows: List[dict], threshold: float = 0.60, min_lines: int = 3,
                     review_floor: float = 0.25,
                     override: Optional[Dict[str, set]] = None
                     ) -> Tuple[Dict[str, dict], Dict[str, str]]:
    """Aggregate lines per vendor and classify each as a coding TYPE (concrete /
    material / both / hauler / review). Returns (agg_by_vendor_upper,
    type_by_vendor_upper). Auto-detected from the *1-vs-*2/3/4 split; an override
    forces the type (hauler is override-only - it can't be auto-detected)."""
    override = override or {}
    agg: Dict[str, dict] = defaultdict(
        lambda: {"vendor": "", "coded": 0, "concrete": 0, "material": 0,
                 "other": 0, "nocode": 0, "amount": 0.0})
    for r in rows:
        a = agg[r["vendor"].upper()]
        a["vendor"] = r["vendor"]
        a["amount"] += r.get("amount") or 0.0
        n = r.get("number")
        if n is None:
            a["nocode"] += 1
        else:
            a["coded"] += 1
            if n in CONCRETE:
                a["concrete"] += 1
            elif n in MATERIAL:
                a["material"] += 1
            else:
                a["other"] += 1

    exc = override.get("exclude", set())
    vtype: Dict[str, str] = {}
    for up, a in agg.items():
        c = a["coded"]
        a["pct_c"] = (a["concrete"] / c) if c else 0.0
        a["pct_m"] = (a["material"] / c) if c else 0.0
        if up in exc:
            continue
        forced = next((t for t in OVERRIDE_TYPES if up in override.get(t, set())), None)
        if forced:
            vtype[up] = forced
            continue
        if c < min_lines:
            if a["concrete"] or a["material"]:
                vtype[up] = "review"
            continue
        pc, pm = a["pct_c"], a["pct_m"]
        if pc >= 0.20 and pm >= 0.20:
            vtype[up] = "both"
        elif pc >= threshold and pm < 0.15:
            vtype[up] = "concrete"
        elif pm >= threshold and pc < 0.15:
            vtype[up] = "material"
        elif pc >= review_floor or pm >= review_floor:
            vtype[up] = "review"
    return agg, vtype


def flag_lines(rows: List[dict], vtype: Dict[str, str]) -> List[dict]:
    """Per-type miscodes. 'review' vendors are never auto-flagged (surfaced only on
    the vendor summary for the owner to type via the override). Each flagged row is
    the original dict + {reason, vtype}."""
    out: List[dict] = []
    for r in rows:
        t = vtype.get(r["vendor"].upper())
        if t in (None, "review"):
            continue
        n = r.get("number")
        code = r.get("cost_code") or ""
        name = r.get("cost_name") or ""
        reason = None

        if t == "concrete":
            if n == "1":
                continue
            if n is not None:
                reason = f"{code} = {name} (expected *1 Concrete)"
            elif code:
                reason = f'Item "{code}" is not a cost code (expected *1 Concrete)'
            else:
                reason = (f"No cost code - account line: "
                          f"{r.get('account') or '(none)'} (expected *1 Concrete)")

        elif t in ("material", "hauler"):
            # material: rebar/lumber/aggregates only. hauler: same PLUS haul-off /
            # equipment (*5/*51/*52) is legitimate (trucking, JA Rock, etc.), so
            # only concrete/labor flag.
            allowed = MATERIAL | (EQUIP if t == "hauler" else set())
            if n in allowed:
                continue
            forbidden = (CONCRETE | LABOR) if t == "hauler" else MATERIAL_FORBIDDEN
            if n in forbidden:
                fam = ("Concrete" if n in CONCRETE
                       else "Equipment" if n in EQUIP else "Labor")
                kind = "HAULER" if t == "hauler" else "MATERIAL"
                expect = ("rebar/lumber/aggregates or haul-off" if t == "hauler"
                          else "rebar/lumber *2/*3")
                reason = f"{code} = {name} - {fam} on a {kind} vendor (expected {expect})"
            else:
                continue                      # fuel/supplies/no-code: not the named set

        elif t == "both":
            if concrete_memo(r.get("desc") or "") and n != "1":
                reason = (f"Memo reads as concrete/yardage but coded "
                          f"{code or '(none)'}{' = ' + name if name else ''} - expected *1")
            else:
                continue

        if reason:
            out.append({**r, "reason": reason, "vtype": t})
    out.sort(key=lambda r: (r["vendor"].upper(), r.get("date") or "",
                            r.get("bill_doc") or ""))
    return out
