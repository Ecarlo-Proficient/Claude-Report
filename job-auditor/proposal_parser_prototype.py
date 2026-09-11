#!/usr/bin/env python3
"""Parse a bid proposal into PRICED LINE ITEMS, then answer the real question:
does the proposal SELL piers, and how does the sold pier line compare to the
takeoff's PR cost band?

Why a parser and not a regex: the item description WRAPS across lines. The line
holding the word "Piers" carries no money; the money lands on the last wrapped
line, which carries no "Piers". Three separate line-regex detectors gave wrong
answers before this was understood.

Item grammar (verified on the live template):
    <description, 1..n wrapped lines> <qty> $<unit> $<total>   → SOLD
    <description ...>                 <qty> $<unit>            → OPTIONAL (not sold)
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

M = r"\$\s?([\d,]+(?:\.\d{2})?)"
SOLD = re.compile(rf"^(?P<desc>.*?)\s*(?P<qty>[\d,.]+)\s+{M}\s+{M}\s*$")
OPT = re.compile(rf"^(?P<desc>.*?)\s*(?P<qty>[\w,.]+)\s+{M}\s*$")
PIER = re.compile(r"(?i)pier")


def money(s):
    try:
        return float(str(s).replace(",", ""))
    except (TypeError, ValueError):
        return None


def parse_items(pdf_path):
    """→ (items, subtotal). items = [{desc, qty, unit, total, sold}]"""
    try:
        import pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            lines = []
            for pg in pdf.pages:
                lines += [l.rstrip() for l in (pg.extract_text() or "").splitlines()]
    except Exception:
        return None, None

    items, buf, subtotal, started = [], [], None, False
    for raw in lines:
        l = raw.strip()
        if not l:
            continue
        if re.match(r"(?i)^description\b", l):
            started = True
            buf = []
            continue
        if re.match(r"(?i)^sub\s*total", l):
            m = re.search(M, l)
            subtotal = money(m.group(1)) if m else None
            break
        if not started:
            continue
        if re.match(r"(?i)^not valid until|^contractor\b|^\d+\.-", l):
            break
        # NB: named groups occupy 1 and 2 (desc, qty) — the money groups are
        # 3 and 4. Using 2/3 here shifted every unit and total one position
        # left and produced a whole run of plausible, wrong numbers.
        ms = SOLD.match(l)
        if ms:
            desc = " ".join(buf + [ms.group("desc")]).strip()
            items.append(dict(desc=desc, qty=ms.group("qty"),
                              unit=money(ms.group(3)), total=money(ms.group(4)),
                              sold=True))
            buf = []
            continue
        mo = OPT.match(l)
        if mo and money(mo.group(3)) is not None:
            desc = " ".join(buf + [mo.group("desc")]).strip()
            items.append(dict(desc=desc, qty=mo.group("qty"),
                              unit=money(mo.group(3)), total=None, sold=False))
            buf = []
            continue
        buf.append(l)
    return items, subtotal


def bands_of(tk):
    from openpyxl import load_workbook as _lw
    try:
        wb = _lw(tk, data_only=True)
    except Exception:
        return None
    try:
        s = next((n for n in wb.sheetnames if "cost gral" in n.lower()), None)
        return P._cost_sheet_totals(wb[s]) if s else None
    finally:
        wb.close()


def main():
    print("\n  PROPOSAL LINE-ITEM AUDIT — does the proposal SELL piers?")
    print("  " + "─" * 72)
    (_k, sp) = P.latest_schedule(RP.SCHEDULE_DIR)
    sched = [s for s in P.read_main_schedule(sp)
             if s["scope"] != "ftw" and not s["job"].startswith("CP")]
    rp_to_folders, addr_folders = RP.index_residential(RP.RP_ROOT)

    sells, nosell, noprop, unparsed = [], [], [], []
    for s in sched:
        job = s["job"]
        fs = sorted(rp_to_folders.get(job, ()), key=lambda f: (f.parent.name, f.name))
        folder = fs[0] if fs else None
        if folder is None and s["address"]:
            parts = s["address"].split(None, 1)
            folder = RP.match_by_address({"house": parts[0] if parts else "",
                                          "street": parts[1] if len(parts) > 1
                                          else s["address"]}, addr_folders)
        ov = P.OVERRIDES.get(job, {})
        prop = Path(ov["proposal"]) if ov.get("proposal") else None
        tk = Path(ov["takeoff"]) if ov.get("takeoff") else None
        if folder is not None:
            if prop is None:
                prop, _a, _n = P.find_proposal(folder, "slab", s["desc"])
            if tk is None:
                tk, _e, _n, _f = P.find_takeoff_etc(folder, job, "slab", s["desc"])
        if tk is None:
            continue
        bands = bands_of(tk)
        if bands is None:
            continue
        PR = bands.get("PR") or {}
        pr_cost = PR.get("sub") if PR.get("sub") is not None else (PR.get("items") or 0.0)

        if prop is None:
            if pr_cost:
                noprop.append((job, s["address"], pr_cost, s["builder"]))
            continue
        items, sub = parse_items(prop)
        if not items:
            unparsed.append((job, s["address"], Path(prop).name))
            continue
        # INTEGRITY GATE: a trustworthy parse reconciles to the proposal's own
        # SUB TOTAL. Anything else is not evidence and must not become a flag.
        tot = sum(i["total"] for i in items if i["sold"] and i["total"])
        if sub is None or abs(tot - sub) > 0.5:
            unparsed.append((job, s["address"],
                             f"{Path(prop).name[:34]} [Σ{tot:,.0f} vs sub"
                             f"{'—' if sub is None else format(sub, ',.0f')}]"))
            continue
        pier_items = [i for i in items if PIER.search(i["desc"])]
        pier_sold = [i for i in pier_items if i["sold"] and i["total"]]
        rev = sum(i["total"] for i in pier_sold)
        if pier_sold:
            sells.append((job, s["address"], rev, pr_cost))
        elif pr_cost:
            nosell.append((job, s["address"], pr_cost, Path(prop).name,
                           bool(pier_items)))

    print(f"\n  proposals SELLING a priced pier line : {len(sells)}")
    print(f"  pier COST but NO sold pier line      : {len(nosell)}   ← the flag")
    print(f"  pier COST but NO proposal at all     : {len(noprop)}")
    print(f"  proposal unparseable                 : {len(unparsed)}")

    print("\n  ── PIER LINE: sold vs cost ──")
    print(f"    {'JOB':<9} {'ADDRESS':<30} {'SOLD $':>11} {'PR COST $':>11} {'MARGIN':>8}")
    for job, addr, rev, cost in sorted(sells, key=lambda r: -r[2])[:16]:
        mg = f"{(rev - cost) / rev * 100:.1f}%" if rev else "—"
        print(f"    {job:<9} {addr[:30]:<30} {rev:>11,.0f} {cost:>11,.0f} {mg:>8}")

    if nosell:
        print("\n  ── ⚠ PIER COST, PROPOSAL SELLS NO PIERS ──")
        for job, addr, cost, pname, mentioned in sorted(nosell, key=lambda r: -r[2]):
            tag = "(described, not priced)" if mentioned else "(never mentioned)"
            print(f"    {job:<9} {addr[:28]:<28} ${cost:>10,.0f}  {tag}  {pname[:38]}")
    if noprop:
        print("\n  ── pier cost, no proposal on file ──")
        for job, addr, cost, b in sorted(noprop, key=lambda r: -r[2]):
            print(f"    {job:<9} {addr[:28]:<28} ${cost:>10,.0f}  {b[:28]}")
    if unparsed:
        print("\n  ── proposal could not be parsed ──")
        for job, addr, n in unparsed:
            print(f"    {job:<9} {addr[:28]:<28} {n[:50]}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
