# WIP Reader — Criteria & Status (by division)

How each division's WIP is built, the exact rules the reader follows, and the
open items. **CP is fully loading — the latest draw (G702) drives contract price,
change orders, billed, and retainage; the takeoff gives ETC; QBO gives costs.
RP's active list + takeoff lookup work; RP contract/ETC is the next step. MFD
hasn't started.**

---

## CP — Commercial  ·  ✅ LOADING (draw-based billing + takeoff ETC + QBO costs)

### Where projects come from
Synology folders, one per project, named `CP#### - PROJECT NAME`:
- **Active:** `/Volumes/Common/CURRENT PROJECTS/Awarded Projects Commercial projects/`
- **Completed:** same path + `/Completed Projects/` — a folder moved here (fully
  billed out) = **closed** (trusted, no QBO cross-check).

### The read chain (per project)

**The draw is the billing source of record** (Ted 2026-07-09). If a project has
a draw (AIA **G702/G703** payment application), the **latest draw** — highest
draw # — supplies Contract Price, Approved COs, Billed-to-Date, and Retainage.
The takeoff is used only for the cost estimate (ETC); QBO only for costs. Before
Draw #1 lands, the reader falls back to the takeoff proposal for contract/CO and
QBO for billed/retainage.

**1. Find the latest draw.** Look for a **`Draws`** (or `Draw`) folder in the
project folder (`Drawings` is not a match). Inside it, the draw workbook lives
either directly or in a numbered **`Draw #N`** subfolder. **Draw # is the
sequence — highest wins.** Supplier Release subfolders are never opened. The
winning file must actually contain a **G702** sheet.
- No draw folder / no draw yet → skip to the takeoff proposal (step 5).

**2. Read the draw's G702** (mapping verified 2026-07-09 against CP585 Draws
#1–#4):

| WIP field | G702 source |
|---|---|
| **Contract Price** | Line 3 Contract Sum to Date (= Line 1 + Line 2) |
| **Approved COs** | Line 2 Net change by Change Orders |
| **Billed to Date** (gross) | Line 4 Total Completed & Stored to Date |
| **Retainage Held** | **Line 4 − Line 6** (Total Earned Less Retainage) |

**Retainage is computed as Line 4 − Line 6, never read from the labeled "Total
Retainage" cell** — that cell is unreliable across draws (reads 0 or a
mismatched figure); Line 4 − Line 6 ties to the 10% on Line 5a every draw.

**3. ETC (Estimated Total Costs)** — from the takeoff **`Bid!AP1961`** (evaluates
the formula if the saved value is stale). ETC always comes from the takeoff,
draw or not.

**4. Costs to Date** — from **QBO** by project # (COGS + Expenses). Draw-backed
projects skip the QBO *billing* fetch entirely (billed/retainage come from the
draw), so the QBO call only pulls costs.

**5. Fallback — no draw yet (pre-Draw #1).** Contract/CO from the takeoff
proposal, billed/retainage from QBO:
- **Pick the takeoff file** — `.xlsx` with **"takeoff"** in the name (auxiliary
  `Cost Codes.xlsx` / `Explanation OH.xlsx` / PDFs ignored). One → use it;
  multiple → only the `WIP`-tagged one(s), summed; none tagged → flag.
- **Pick the proposal tab** — a tab marked `FINAL`, else the only proposal tab,
  else flag.
- **Contract Price** — from `GRAND TOTAL`, `SUB TOTAL`, or plain `TOTAL` (largest
  wins; `TOTAL SQFT`/`TOTAL YARDS` ignored).
- **Approved COs** — in-takeoff `Change Order#N` sheets + a `Change Orders/`
  subfolder, summed. CO cost has no source yet → ETC stays at base (provisional,
  flagged).
- **Billed to Date / Retainage Held** — QBO income (GROSS, incl retainage);
  Retainage = gross − net collectible (net excludes "Retainage Not Billed" memo
  invoices).

**7. Derived WIP columns** written as live Excel formulas: Cost to Complete,
Original Profit, Gross Profit %, % Complete, Revenues Earned, Profit Earned,
Future Profit, Overbillings, Underbillings, Left to Bill, Pure Job Borrow.

**8. Write** to the `Test - CP` tab only (code-level guard; live tabs untouched).
Yellow = the sourced inputs, white = calculations, gray header, thin borders,
wrapped as an Excel table. Project name links to the Awarded Project folder
(works on Mac).

### Flags you'll see (and what they mean)
- **Draw #N: billed … retainage … contract …** → normal; the row's billing came
  from that draw. Confirms which draw was used.
- **No draw yet — contract from takeoff proposal** → pre-Draw#1; billing is from
  QBO and contract from the proposal until the first draw lands.
- **Draw #N unreadable / Draw G702 missing …** → open and save the draw in Excel
  to refresh its cached values, then re-run.
- **Multiple proposals, none marked FINAL** → estimator must mark the real one
  (fallback path only).
- **Missing Grand/Sub Total** → the proposal has no total the reader recognizes
  (fallback path only).
- **CO Rev without CO Cost** → Revised ETC is provisional (base ETC), pending a
  CO cost cell in the template (fallback path only).
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
| **CP** | ✅ Synology folders | ✅ latest draw G702 (contract/CO/billed/retainage) + Bid!AP1961 ETC; takeoff proposal is the pre-Draw#1 fallback | ✅ costs (billed/retainage from draw) | **Loading** |
| **RP** | ✅ Alpha (active) + Residential (takeoff by RP#) | 🟡 next step | ✅ available by RP# | **In progress** |
| **MFD** | ⚠️ parent "Multi Family" | 🔴 source unknown | ✅ available by project # | **Not started** |
