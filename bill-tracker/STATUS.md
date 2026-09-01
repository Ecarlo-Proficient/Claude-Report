# STATUS — bill-tracker

Shared progression record (the user's sessions ↔ the developer's). Tool scope only —
no business findings, dollar exposures, or owner analyses (those live in the owner's vault).

## DONE / FINALIZED

- **Pipeline:** `run_tracker.sh` → `sync_view.py` → `excel_bill_sync.py` → `Bill Tracker.xlsx`
  (OneDrive `Automations-/`). Sheets: Bills · Inventory · QBO Audit · Liens.
- **Full pull incl. subs (2026-08-06).** Every bill is fetched (open any-date + paid since
  `PAID_CUTOFF_DATE`), subs included. Subs are marked per row (`is_sub`) and filtered to
  `display_rows` for Bills/Inventory/Liens; the audit sheet gets the full population.
- **Cost code capture (2026-08-06).** `bill_rows.line_cost_code()` keeps the raw QBO Item
  name (`SL1`/`FW2`…) on each row as `cost_code`, audit-only — no display column added.
- **QBO Audit = 6 sections (2026-08-06).** Added §5 FW-misplacement (FW on CP/MFD/base
  RP#### slab; legit only on `-FTW`) and §6 SUB bill missing project. §4 duplicates widened
  to the full population incl. subs.
- **Audit → one Excel Table per section (2026-08-06).** The single collapsible "QBO Audit"
  banner sheet is replaced by six **`Audit - …`** sheets (Not Approved · Data Entry · Missing
  Project · Duplicates · FW Misplaced · Sub No Project), each a proper Table (filter/sort on
  every column) — banners can't live inside a table. `build_audit_sheets()` +
  `_audit_table_sheet()`. Passes `validate_xlsx` (opens clean in Excel). Empty sections render
  a valid one-row table. NOTE: the old banner helpers/constants
  (`_audit_section_banner`, `AUDIT_HEADERS`, `AUDIT_SECTION_*_FILL`, …) are now dead code —
  prune in a follow-up.
- **Audit consolidation (2026-08-06).** `duplicate_bill_audit.py`, `item_no_project_audit.py`,
  `sub_bill_audit.py` retired — folded into the QBO Audit sheet. `job_coding_audit.py` kept as
  the interactive per-job drill (`audit-job` alias) and shared helper lib.
- **Sub No Project = cost-code items only (2026-08-12, owner).** §6 now keeps only item/COGS
  lines (`bill_type == "COGS"`) with no project #; account-based category lines (reimbursements,
  fees, overhead) are dropped — they carry no cost code and aren't job cost to chase onto a
  project. One-line filter in `build_audit_sheets()`.
- **Bills AR column reorder (2026-08-12, owner).** `Client` moved out of the identity block into
  the CLIENT PAYMENT (AR) section, right BEFORE `Matched Invoice`; `Invoice #` moved to right
  AFTER `Matched Invoice`. AR reads Invoice Status · Client · Matched Invoice · Invoice #. Only
  `BILL_ROW_COLS`, `COL_WIDTHS`, and the positional `values` list changed — every other column
  ref is name-keyed (`HEADERS.index`) and auto-followed (Liens sheet included).
- **Cell-only conditional formatting (2026-08-12, owner).** Killed the whole-row reconciliation
  band — the Pay×Invoice color now paints ONLY the `Invoice Status` cell. `Approved` cell gains a
  green tint for `"approved"` (kept red for `"not approved"`); `Lien` cell was already cell-scoped.
  No cell is colored by another cell's status. Legend relabeled `ROW KEY →` → `STATUS KEY →`.
  Verified: `validate_xlsx` clean; CF sqref = `N`(Approved) / `O`(Lien) / `R`(Invoice Status) only.
- **Invoice-status colors + new AR state (2026-08-12, owner).** `Invoice paid` is now ALWAYS green
  (dropped the gray "done" tint). New invoice status **`Partially Paid/Awaiting Remainder`** —
  emitted by `compute_invoice_status` when a matched invoice has `0 < Balance < Total` (GC paid part)
  — colored NEUTRAL tan (`FFEB9C`, Excel "Neutral"). `_aggregate_invoice_status` preserves it
  (single-project) and handles mixed multi-project sets. Legend: green=`Invoice paid`,
  tan=`Awaiting / partial pay`, gray "Done" removed. New constant in `qbo_bill_tracker.py`
  (`STATUS_PARTIALLY_PAID_REMAINDER` + `COLOR_PARTIAL_REMAINDER` + `STATUS_FILL_MAP` entry).
- **Invoice Open Bal column (2026-08-12, owner).** Added `Invoice Open Bal` (from `inv_balance`,
  the matched invoice's QBO Balance) to the Bills/Inventory sheets, positioned right BEFORE
  `Invoice Total`. 28 → 29 columns; positional `values` list + `COL_WIDTHS` updated, dividers now
  at cols 13/17. Verified: 29 cols, AR order `… Invoice Date · Invoice Open Bal · Invoice Total …`,
  `validate_xlsx` clean, 6/6 existing tests pass.
- **`Audit - Unused PO` sheet — PO tracker × QBO, one story (2026-08-25, owner).** New
  `po_tracker.py` reads the office PO tracker workbook (`Orders` tab, READ-ONLY) and reconciles
  it against QBO POs. New `bill_rows.build_po_index()` pulls `PurchaseOrder` once with
  `POStatus`/`TxnDate`/`VendorRef`/`TotalAmt`/`LinkedTxn`/job (replaces `build_po_map` in main;
  id→doc derived, so still ONE PO pull). Flags: **Open, no bill** · **Stale >60d** (open+unbilled
  aged past 60d) · **On tracker, not in QBO** (recent, UNBILLED tracker PO never issued — billed
  tracker rows are excluded as not-"unused"). One row per PO, QBO + tracker columns side by side,
  `polink` → QBO PO deep link; tracker freshness stamped to the right of the header. Tracker path
  = `ACB_PO_TRACKER_XLSX` (default OneDrive `Purchase Orders/Copy 05 dic.xlsx` — the CURRENT file;
  `1.0purchase-order-tracker.xlsx` is ~15 mo stale, do not use). Degrades to a QBO-only view if
  the tracker is unreadable. Verified offline (real tracker + mock QBO): all 3 flags fire, QBO rows
  link / tracker-only rows don't, `validate_xlsx` clean. **First live `sync-ap` validates the QBO
  `build_po_index` parse.**
- **`Audit - Cost Code` sheet — vendors code to their family (2026-08-25, owner).** 8th audit sheet.
  Cost-code NUMBER = family (1 concrete · 2/3/4 material · 5/51/52 equip · 6 labor). Auto-captures
  each vendor's TYPE from its `*1`-vs-`*2/3/4` split — **concrete** (→ all `*1`), **material** (RCI
  → `*2/*3/*4`, never `*1`/`*5`/`*6`), **both** (Preferred Materials → yardage/ready-mix MEMO line
  must be `*1`) — then flags lines breaking the rule. Runs over the full `all_rows` sync-ap already
  pulls. Logic lives in **`shared/cost_code_audit.py`** (shared with the standalone
  `one-offs/concrete_cost_code_audit.py` — repo rule: shared, not tool→tool). Types overridable via
  `<companyhealth>/concrete_suppliers.json` `{concrete/material/both/exclude}`. Verified offline
  (Cowtown/RCI/Preferred mock): types + all 4 flags correct incl. yardage-memo catch, `validate_xlsx`
  clean, 6/6 tests pass.
- **Cost Code audit → PO-origin cross-reference (2026-08-25, owner).** Each flagged line now shows
  `PO #` · `PO Cost Code` · `Origin`: did the linked PO ALSO carry the wrong family (**upstream** -
  super/PM built the PO wrong, clerk trusted it), did the bill **deviate** from a correct PO, or was
  there **no PO** (clerk-coded)? PO cost codes come from QBO (`bill_rows.build_po_index` now captures
  each PO's line `codes`/`numbers`); the PO tracker recovers the PO# by Bill# when QBO left the bill
  unlinked (the tracker's own Cost Code column is unusable - 2% filled, foreign format). Verdict in
  `shared/cost_code_audit.po_origin`. Verified offline (4 scenarios: PO-also-wrong, bill-deviated,
  tracker-recovered, no-PO), `validate_xlsx` clean. **Standalone one-off doesn't yet carry the PO
  column** (would need a PO pull); bill-tracker sheet is the PO-aware surface.
- **Cost Code: hauler type + dict overrides (2026-08-25, owner).** `shared/cost_code_audit`
  overrides moved tuple→dict `{concrete/material/both/hauler/exclude}`. New **hauler** type = a
  material vendor where haul-off/equipment (`*5/*51/*52`) is legit (trucking, JA Rock), so only
  `*1`/`*6` flag - kills the haul-off false positives. `both` fixes vendors like Preferred Materials.
  Overrides live OUTSIDE the repo at `<companyhealth>/concrete_suppliers.json` (vendor names). Live
  run (3,697 bills / 13,030 lines): overrides cut miscodes 163→94; origin split 29 upstream (PO/
  super-PM) · 63 no-PO (direct entry) · 0 deviated. `hauler` is override-only (not auto-detected).
- **Audit Status marks - acknowledge & keep (2026-08-25, owner).** `Audit - Coding` gained an
  editable **`Status`** column (+ hidden `_Key`=bill_id). Type a mark (e.g. `KEEP - reason`) to
  acknowledge a one-off you're deliberately keeping (vs. `audit_exclusions.json` which HIDES it):
  the mark **persists across runs** (`preserve_audit_marks`, keyed by bill_id) and **mirrors into
  the Bills `Notes`** (the audit is the entry point → pushed in `main`, idempotent append; Bills
  row key = bill_id, so it lands). Filter `Status`: blank = still needs review, marked = on the
  record. Nothing lost - stored in BOTH the audit and Bills Notes. Verified offline round-trip
  (build → mark → re-read → rebuild persists + pushes, `validate_xlsx` clean, 6/6 tests). Edge:
  editing/removing a mark may leave a stale copy in Bills Notes - clear both to fully remove.
- **Missing Project exclusions (EXCEL ONLY, 2026-08-25, owner).** Known-legit no-project vendors
  (equipment rental, overhead, insurance) and classes are suppressed from the Missing Project check
  via `<companyhealth>/audit_exclusions.json` (`{missing_project: {vendors, classes}}`,
  case-insensitive substring). Seeded: `Bobcat of North Texas` (owner said "North Dallas" - no such
  vendor), `CBUSA`, `Cindy Montgomery` (Agency LLC), class `10702-10704` (C.F Hawn Fwy). Loader
  `_load_audit_exclusions` + `_excluded` in `build_audits`. **Ledger untouched** - Excel audit only.
- **Audit consolidation → 3 themed sheets + Missing PO (2026-08-25, owner).** De-bloat: the 8
  `Audit - …` tabs merged into **THREE** filterable Tables, each with an `Issue` column +
  `Detail` + `Open`: **`Audit - Coding`** (Data Entry · Missing Project · FW Misplaced · Sub No
  Project · Cost Code) · **`Audit - PO`** (Unused PO · **new Missing PO** = a real COGS bill —
  not sub, not expense-only — with NO PO, last 90 days) · **`Audit - Bills`** (Not Approved ·
  Duplicates). One orchestrator `build_audits()` (+ `_cost_code_findings`, `_missing_po_bills`)
  replaces `build_audit_sheets`/`build_unused_po_sheet`/`build_cost_code_sheet` (all deleted);
  finding logic unchanged, only rendering. New `url` cell-kind lets one sheet mix bill + PO
  deep-links. Verified offline (every issue routes to the right sheet, `validate_xlsx` clean,
  6/6 tests). **First live `sync-ap` confirms the merged output.**
- **Cost Code: aggregate (*4) allowed on concrete vendors WHEN the memo says so (2026-08-25, owner).**
  Ready-mix vendors haul aggregate (pea gravel / sand / base), so `*4` is legit on a concrete vendor
  ONLY when the line memo reads as aggregate (`aggregate_memo` - pea gravel, sand, gravel, select
  fill, flex base, crushed, asphalt, dirt…). A `*4` with no aggregate memo still flags, as do rebar/
  lumber/pump/labor. Owner: NOT a blanket `*4` pass - memo-gated. Pump (`*51`) still flags per prior
  call. In `shared/cost_code_audit.flag_lines`.
- **`Audit - History` — persistent cost-code miscode log (2026-09-01, owner).** "I need this logged
  in the system - how often the clerk is making the mistakes, and what got fixed after refreshing."
  New module `cost_code_history.py` (pure state, no QBO) + a fourth audit sheet. State =
  `<companyhealth>/cost_code_history.json` (OUTSIDE the repo, never committed). `_cost_code_findings`
  now also returns the raw flags; each REAL run (dry-run bails first) opens/updates an entry
  (First/Last Seen + Times) per current miscode, flips any vanished OPEN entry to FIXED (with date),
  and re-opens a reappearing one. Key = `bill_id|cost_code`. Sheet shows all OPEN + FIXED-within-60d
  (older fixes kept off-sheet); caption = open/new/fixed counts + rolling new/run rate (one clerk, so
  the aggregate IS the clerk rate). Verified: state machine unit-tested (new/fixed/reopen, persistence,
  same-key amount sum, corrupt-file resilience); integration test renders the sheet, passes
  `validate_xlsx` + `xlsx_verify.assert_clean`, and confirms a recode flips OPEN→FIXED.
- **Cost Code: skip non-job fees / read the memo (2026-08-25, owner).** Credit-card, finance, bank,
  and late fees legitimately post to an EXPENSE ACCOUNT, not a cost code - `shared/cost_code_audit`
  now skips any line whose memo/bill#/account reads as one (`is_nonjob`), and a concrete vendor's
  account line (no cost code) only flags when its MEMO reads as a real concrete purchase
  (`concrete_memo`), not as overhead. Cut the live count 94→68. **Open question:** ready-mix vendors
  billing PUMP (`*51`) - likely legitimate (they provide the pump), so concrete-vendor flagging may
  need to allow `*51/*52`; pending owner call.

- **No realm in the terminal (2026-08-06, owner).** The auth step printed
  `ok. company_id=<realm>`; removed — it now just prints `ok.`, consistent with `sync-ar`
  (which never echoes the realm). Swept the same leak out of `qbo_bill_tracker.py`,
  `qbo-export`, `health-dashboard`, and `wip/qbo_bulk_close.py`. **Rule: never echo the QBO
  company_id / realm to stdout or a tee'd log.**

## OPEN ISSUES

- **Parity check pending a live run.** The folded sections should reproduce the retired
  scripts' findings on the same data. Population is a superset (broader sub detection, full
  pop for dups), so no regression is expected — confirm on the first `sync-ap` and restore
  from git if a gap appears.
- **Audit history window.** The retired scripts could scan all-time; the audit runs over the
  tracker's population (open + paid-since-`PAID_CUTOFF_DATE`), so pre-cutoff **paid** bills are
  outside its window. Revisit only if the owner wants a full-history coding sweep.
- **FW audit is not active-scoped.** It flags FW miscodes on any pulled bill regardless of
  the job's WIP status (no WIP-master dependency on the tracker). Add active-only filtering
  (promote `load_wip_master` to `shared/`) only if the owner asks.

## TO DO

- (none queued)
