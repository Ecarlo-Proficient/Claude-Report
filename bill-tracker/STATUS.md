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
