#!/usr/bin/env python3
"""ETC COMPOSITION AUDIT v2 — read-only.

v1 lessons (why v2 exists):
  • v1's phantom-pier test used `PR subtotal > 0`. 19 of 45 jobs have a PR
    subtotal of #N/A, so the test could never fire on them — the largest
    suspect population was invisible. v2 falls back to the PR ITEM sum.
  • v1 treated "the word PIER appears in the proposal" as "the proposal sells
    piers". v2 also captures the pier LINES and whether any carries a price,
    and prints samples so the signal can be judged rather than trusted.
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

_R = str(Path(__file__).resolve().parent.parent)   # repo root — self-locating
for _p in (_R, _R + "/one-offs", _R + "/wip"):
    sys.path.insert(0, _p)

import rp_wip_reader as RP            # noqa: E402
import rp_schedule_wip_preview as P   # noqa: E402

# NOTE: prefix match, NO trailing \b — `\bpier\b` silently misses the PLURAL
# "PIERS", which is how these proposals actually word it. That bug made v2's
# first run report 30 phantom jobs that were nothing but a regex miss.
PIER_RE = re.compile(r"(?i)\b(piers?|drilled|bell\s*bottom|caisson)")
MONEY_RE = re.compile(r"\$\s?[\d,]+(?:\.\d{2})?")


def pdf_lines(path):
    try:
        import pdfplumber
        with pdfplumber.open(path) as pdf:
            txt = "\n".join((pg.extract_text() or "") for pg in pdf.pages)
        return [l.strip() for l in txt.splitlines() if l.strip()]
    except Exception:
        return None


def pier_evidence(lines):
    """Returns (mentions, priced_lines, all_pier_lines)."""
    if lines is None:
        return False, [], []
    hits = [l for l in lines if PIER_RE.search(l)]
    priced = [l for l in hits if MONEY_RE.search(l)]
    return bool(hits), priced, hits


def bands_of(tk_path):
    from openpyxl import load_workbook as _lw
    try:
        wb = _lw(tk_path, data_only=True)
    except Exception:
        return None, "takeoff unreadable"
    try:
        sheet = next((n for n in wb.sheetnames if "cost gral" in n.lower()), None)
        if sheet is None:
            return None, "no 'Cost Gral' sheet"
        return P._cost_sheet_totals(wb[sheet]), ""
    finally:
        wb.close()


def main():
    print("\n  ETC COMPOSITION AUDIT v2 — read-only")
    print("  " + "─" * 74)
    best = P.latest_schedule(RP.SCHEDULE_DIR)
    (_k, sched_path) = best
    print(f"  schedule: {sched_path.name}")
    sched = [s for s in P.read_main_schedule(sched_path)
             if s["scope"] != "ftw" and not s["job"].startswith("CP")]
    print(f"  active SLAB lines: {len(sched)}")
    rp_to_folders, addr_folders = RP.index_residential(RP.RP_ROOT)
    print("  folder index built\n")

    rows, errors = [], []
    for s in sched:
        job, desc = s["job"], s["desc"]
        folders = sorted(rp_to_folders.get(job, ()), key=lambda f: (f.parent.name, f.name))
        folder = folders[0] if folders else None
        if folder is None and s["address"]:
            parts = s["address"].split(None, 1)
            folder = RP.match_by_address(
                {"house": parts[0] if parts else "",
                 "street": parts[1] if len(parts) > 1 else s["address"]}, addr_folders)

        ov = P.OVERRIDES.get(job, {})
        prop = Path(ov["proposal"]) if ov.get("proposal") else None
        tk = Path(ov["takeoff"]) if ov.get("takeoff") else None
        if folder is not None:
            if prop is None:
                prop, _a, _n = P.find_proposal(folder, "slab", desc)
            if tk is None:
                tk, _e, _n, _f = P.find_takeoff_etc(folder, job, "slab", desc)
        if tk is None:
            errors.append((job, s["address"], "no takeoff found"))
            continue
        bands, err = bands_of(tk)
        if bands is None:
            errors.append((job, s["address"], err))
            continue

        SL, PR, FW = (bands.get(k) or {} for k in ("SL", "PR", "FW"))
        pr_sub, pr_items = PR.get("sub"), PR.get("items") or 0.0
        # v2: a band "carries cost" via its subtotal OR (when #N/A) its items
        pr_cost = pr_sub if pr_sub is not None else pr_items
        sl_cost = SL.get("sub") if SL.get("sub") is not None else (SL.get("items") or 0.0)

        lines = pdf_lines(prop) if prop else None
        mentions, priced, hits = pier_evidence(lines)

        rows.append(dict(job=job, addr=s["address"], builder=s["builder"],
                         tract=P._is_tract(s["builder"]),
                         sl_sub=SL.get("sub"), sl_items=SL.get("items") or 0.0,
                         pr_sub=pr_sub, pr_items=pr_items, pr_cost=pr_cost,
                         fw_sub=FW.get("sub"), fw_items=FW.get("items") or 0.0,
                         sl_cost=sl_cost,
                         prop=Path(prop).name if prop else None,
                         mentions=mentions, priced=priced, hits=hits,
                         readable=lines is not None))

    # ── how good is the "proposal sells piers" signal?
    withprop = [r for r in rows if r["prop"] and r["readable"]]
    print("  ── SIGNAL QUALITY: does 'PIER' in the proposal mean anything? ──")
    print(f"    proposals read            : {len(withprop)}")
    print(f"    mention a pier word       : {sum(1 for r in withprop if r['mentions'])}")
    print(f"    have a PRICED pier line   : {sum(1 for r in withprop if r['priced'])}")
    print("\n    sample pier lines (judge boilerplate vs real scope):")
    seen = set()
    for r in withprop:
        for l in r["hits"][:2]:
            k = re.sub(r"[\d,.$]+", "#", l)[:70]
            if k in seen:
                continue
            seen.add(k)
            print(f"      [{r['job']}] {l[:110]}")
        if len(seen) >= 12:
            break

    # ── flags
    print("\n  ── FLAGS " + "─" * 65)
    phantom, missing, nafw, nasl, napr, fwleak = [], [], [], [], [], []
    for r in rows:
        if r["pr_cost"] and r["prop"] and r["readable"] and not r["mentions"]:
            phantom.append(r)
        if r["pr_cost"] and r["prop"] and r["readable"] and r["mentions"] and not r["priced"]:
            missing.append(r)          # mentioned but never priced — weak sell
        if r["sl_sub"] is None and r["sl_items"]:
            nasl.append(r)
        if r["pr_sub"] is None and r["pr_items"]:
            napr.append(r)
        if r["fw_sub"] is None and r["fw_items"]:
            nafw.append(r)
        if r["fw_sub"]:
            fwleak.append(r)

    def show(title, lst, amt):
        tot = sum(amt(r) for r in lst)
        print(f"\n  {title}: {len(lst)} jobs · ${tot:,.0f}")
        for r in sorted(lst, key=lambda r: -amt(r))[:12]:
            print(f"    {r['job']:<9} {r['addr'][:30]:<30} ${amt(r):>10,.0f}  {r['prop'] or '—'}")

    show("PHANTOM PIERS (pier cost, proposal never says pier)", phantom, lambda r: r["pr_cost"])
    show("PIERS MENTIONED BUT NEVER PRICED (pier cost carried)", missing, lambda r: r["pr_cost"])
    show("#N/A SLAB — slab band DROPPED from ETC", nasl, lambda r: r["sl_items"])
    show("#N/A PIERS — pier band DROPPED from ETC", napr, lambda r: r["pr_items"])
    show("FW LEAK — flatwork cost on a slab takeoff", fwleak, lambda r: r["fw_sub"] or 0)

    if errors:
        print("\n  ── COULD NOT AUDIT ──")
        for j, a, w in errors:
            print(f"    {j:<10} {a:<32} {w}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
