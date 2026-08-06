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
