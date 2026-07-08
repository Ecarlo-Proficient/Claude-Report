# WIP Reader — Criteria & Status (by division)

How each division's WIP is built, the exact rules the reader follows, and the
open items. **CP is fully loading — contract price, change orders, ETC, and the
QBO actuals all pull automatically. RP's active list + takeoff lookup work; RP
contract/ETC is the next step. MFD hasn't started.**

---

## CP — Commercial  ·  ✅ LOADING (contract + COs + ETC + QBO all working)

### Where projects come from
Synology folders, one per project, named `CP#### - PROJECT NAME`:
- **Active:** `/Volumes/Common/CURRENT PROJECTS/Awarded Projects Commercial projects/`
- **Completed:** same path + `/Completed Projects/` — a folder moved here (fully
  billed out) = **closed** (trusted, no QBO cross-check).

### The read chain (per project)

**1. Pick the takeoff file.** A takeoff is the `.xlsx` with **"takeoff"** in the
filename. Auxiliary files (`Cost Codes.xlsx`, `Explanation OH.xlsx`, PDFs) are
ignored.
- One takeoff → use it.
- Multiple takeoffs → only the ones tagged **`WIP`** in the filename, and sum
  them (a project can have more than one scope, e.g. FDT + PAVING).
- Multiple takeoffs, none tagged WIP → flag, leave blank.

**2. Pick the proposal tab** (takeoffs carry several — dated versions, Commercial
/ Residential / Alternative).
- A tab marked **`FINAL`** → use it.
- Only one proposal tab → use it.
- Multiple, none marked FINAL → flag, leave blank.

**3. Contract Price.** Read off the chosen proposal tab, from **`GRAND TOTAL`,
`SUB TOTAL`, or plain `TOTAL`** (templates use all three) — value in the cell to
the right; if a label repeats, take the largest (the overall). `TOTAL SQFT` /
`TOTAL YARDS` are ignored (exact-label match only).

**4. Change Orders → Approved COs.** Summed from both places, additive:
- In-takeoff `Change Order#N` sheets — from each sheet's `TOTAL:` cell.
- A `Change Orders/` subfolder of standalone CO files.
- **CO cost has no source yet** (the CO template has no cost cell), so Estimated
  Total Costs stays at the base ETC (provisional) and is flagged.

**5. ETC (Estimated Total Costs).** From `Bid!AP1961` (evaluates the formula if
the saved value is stale).

**6. QBO actuals** by project # (`CP####`):
- **Billed to Date** = QBO income — **GROSS, includes retainage** (standard WIP).
- **Retainage Held** = gross billed − net collectible (net excludes the
  "Retainage Not Billed" memo invoices — the reclass to Retainage Receivable).
- **Costs to Date** = COGS + Expenses.

**7. Derived WIP columns** written as live Excel formulas: Cost to Complete,
Original Profit, Gross Profit %, % Complete, Revenues Earned, Profit Earned,
Future Profit, Overbillings, Underbillings, Left to Bill, Pure Job Borrow.

**8. Write** to the `Test - CP` tab only (code-level guard; live tabs untouched).
Yellow = the sourced inputs, white = calculations, gray header, thin borders,
wrapped as an Excel table. Project name links to the Awarded Project folder
(works on Mac).

### Flags you'll see (and what they mean)
- **Multiple proposals, none marked FINAL** → estimator must mark the real one.
- **Missing Grand/Sub Total** → the proposal has no total the reader recognizes.
- **CO Rev without CO Cost** → Revised ETC is provisional (base ETC), pending a
  CO cost cell in the template.
- **⚠ OVER BUDGET — Costs exceed ETC** → costs have passed the estimate.
- **Multiple takeoffs — none tagged WIP** / **No takeoff file** → naming needs a fix.

### How to run
Full sync (writes the `Test - CP` tab):
```
python3 "/Users/sebas/Documents/Claude/Projects/Automate Concrete Business/automation-worker/cp_wip_reader.py"
```
Fast takeoff-only audit (skips the slow QBO join — contract/CO/ETC only):
```
python3 "/Users/sebas/Documents/Claude/Projects/Automate Concrete Business/automation-worker/cp_wip_reader.py" --dry-run --no-qbo
```
(Close the WIP file in Excel first, or the write safely skips.)

### Open items
- **CO cost cell** — estimators add a cost line to the CO template → Revised ETC
  can include CO cost instead of running provisional.
- **Verified / locked WIP + PM sign-off** — designed, not built (levels + the
  anti-lockout rule still to finalize).

---

## RP — Residential  ·  🟡 Active list + takeoff lookup working; contract/ETC next

### Architecture (parallels CP)
- **Active projects come from the ALPHA LIST** (the RP equivalent of "Awarded
  Projects"): `/Volumes/Common/OPERATIONS/GENERAL LIST/LISTA GENERAL AÑO 2026.xlsx`,
  sheet `General list - Alpha order`. A job is **active when COMPLETION < 100%**
  (100% or fully greyed = done). Currently **15 active RP slab projects**.
- **The Residential folder is the takeoff lookup:**
  `/Volumes/Common/CURRENT PROJECTS/Residential/<CLIENT>/<ADDRESS>/`, with the
  **RP# in the takeoff filenames**. For each active RP#, find its takeoff by
  RP# in the filename; fall back to matching the address if the RP# isn't in a
  filename. On the live run all 15 matched by RP#.

### Run
```
python3 "/Users/sebas/Documents/Claude/Projects/Automate Concrete Business/automation-worker/rp_wip_reader.py"
```

### Next step
Per RP#, pick the right takeoff/proposal inside the folder and pull **contract
price + ETC** — same FINAL-proposal + GRAND/SUB/TOTAL logic as CP, once we
confirm the RP takeoff layout. Then the QBO join (billed/costs/retainage by RP#)
is the same as CP. Flatwork (Schedule) is a separate, later pass.

---

## MFD — Multi-Family  ·  🔴 Not started — sources unknown

- MFD = parent customer **"Multi Family"** (or `MFD####` prefix). Class field
  not trusted; closures always manual.
- **We don't yet know where MFD contract price / ETC come from** — the source
  (takeoff? WIP Master snapshot? QBO budget?) hasn't been identified.
- MFD WIP currently lives in the WIP Master file as monthly snapshots.
- **First step:** find where the contract + ETC live before any automation.

---

## One-line status

| Division | Find projects | Contract / ETC | Actuals (QBO) | Status |
|---|---|---|---|---|
| **CP** | ✅ Synology folders | ✅ FINAL proposal → GRAND/SUB/TOTAL + Bid!AP1961, COs summed | ✅ billed/costs/retainage | **Loading** |
| **RP** | ✅ Alpha (active) + Residential (takeoff by RP#) | 🟡 next step | ✅ available by RP# | **In progress** |
| **MFD** | ⚠️ parent "Multi Family" | 🔴 source unknown | ✅ available by project # | **Not started** |
