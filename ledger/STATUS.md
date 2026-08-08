# ledger/ — STATUS

Progression record for the canonical project database. Update in the SAME commit as any
change to this tool (repo rule). Tool-scope only — business/dollar analyses live in the vault.

## DONE / FINALIZED
- **`schema.sql`** — the 6-table spine (`project`, `cost_code`, `budget_line`, `cost_line`,
  `billing_event`, `wip_snapshot`) + `v_wip_latest` view. Portable across SQLite and Postgres
  (natural keys, ISO-text timestamps, 0/1 booleans, `DROP VIEW`+`CREATE VIEW`, `ON CONFLICT`).
- **`load_wip_master.py`** — lands the FINAL WIP master Test tabs into `project` + `wip_snapshot`.
  - CP←`Test - CP`, RP←`Test - RP`, MFD←`Test-Master`; each project read once from its richest tab.
  - Rows filtered to `^(MFD|CP|RP)\d+(-FTW)?$` → all legend/total/section rows excluded.
  - Excel opened **read-only**; upserts idempotent by `project_no` / `(project_no, report_date)`.
  - `--dry-run` (write nothing) and `--show N` (sample after load).
- **Verified against the 2026-08-07 master:** 170 projects (Commercial 48, Residential 119,
  Multi Family 3), report date parsed, re-run leaves 0 duplicate snapshot keys, `v_wip_latest`
  and division/category rollups query correctly.
- `docs/ARCHITECTURE.md` updated (new "Ledger" section + folder map + folder-map diagram).
- **`dashboard.py` + `static/`** — local web dashboard over the ledger (Phase-1 UI / "Rung 1").
  - Portfolio KPIs, division rollup, searchable/filterable/sortable projects table, click-into-job
    detail, click-to-copy cells, CSV export.
  - **Customize panel:** theme (auto/light/dark), accent, font, text size, density, width, widget
    toggles, per-column visibility — saved per person in `localStorage`.
  - Read-only on the DB; binds 127.0.0.1 only; stdlib server (no Flask). No new dependencies.
  - Verified live in the browser (light + dark), 170 projects, detail + settings + sort all working.
  - NOTE: the app-preview sandbox can't run this server (it needs the DB + `shared/` outside
    `.preview`); run it directly with `python3 ledger/dashboard.py`. A `.claude/launch.json` entry
    (`ledger-dashboard`) exists but launch.json is untracked/local.

## IN PROGRESS
- (none)

## TO DO
- **Phase 2 connectors** fill the granular tables — one at a time, no rewrites of the existing
  tools' outputs:
  - `cost_code` + `cost_line` from `bill-tracker` (cost code via `cost_leaf()`; append-only by bill+line id).
  - `budget_line` from the takeoff/ETC extractor (`shared/takeoff_etc.py`).
  - `billing_event` from `invoice-sync` (draw period from PrivateNote).
- Once granular tables exist, add computed WIP views (over/under-billing, budget-vs-actual by
  cost code) so those numbers come from the spine, not Excel columns.
- Postgres deployment decision (Synology container vs. small cloud box) — schema is ready either way.
- Optional: a read-only dashboard over `v_wip_latest` (Phase 3) — DB first, UI later.

## OPEN ISSUES / NOTES
- MFD rows come from `Test-Master`, which has **no STATUS column** → MFD `status` loads as NULL
  (MFD closures are manual anyway). RP/CP status comes from their own tabs.
- `wip_snapshot` stores the master's already-computed figures verbatim (source of truth = the
  sheet, per the "do not generate" instruction). When Phase-2 granular data exists, snapshots can
  be reconciled against a spine-computed WIP as a cross-check.
- The RP tab has no retainage / over-under / earned columns (those are computed on the master);
  for RP those snapshot fields load as NULL by design.
