# Automation Suite — System Map

> **Maintenance rule (binding):** any commit that adds/removes a script, changes what a script
> reads or writes, or rewires a data flow MUST update this file in the same commit.
> Claude sessions: check this file whenever `git status` shows script changes at session end.
> GitHub renders the Mermaid blocks natively — view this file on github.com.
>
> **Presentation view:** [`architecture.html`](architecture.html) is the designed, full-system
> picture (open it in a browser after pulling). Refresh it when structure meaningfully changes;
> THIS file is the always-current source of truth.

Last updated: 2026-08-08 (ledger/ cont'd: shared/qbo_costs.py — cost_leaf MOVED out of project-pnl
(imported back, byte-compatible) so the ledger shares the ONE cost-code resolver; load_costs.py pulls
QBO Bills+Purchases → complete cost_line by cost code incl. subs, reconciles to wip_snapshot,
--selftest proves it offline. cost_line fleshed out + v_cost_by_project/v_cost_by_code views.
CLAUDE.md updated (cost_leaf location + ledger subsystem bullet). — earlier same day —
ledger/: new — the canonical project database. schema.sql defines the
6-table spine [project · cost_code · budget_line · cost_line · billing_event · wip_snapshot],
portable across SQLite and Postgres; load_wip_master.py lands the final WIP master Test tabs into
project + wip_snapshot, read-only on Excel, idempotent. load_bill_tracker.py adds an AP + lien feed
(ap_bill_line) from Bill Tracker.xlsx — deliberately NOT cost truth (subs excluded). dashboard.py +
static/ add a local read-only web UI (KPIs · Needs-attention · AP & liens · division rollup ·
projects · job detail · copy/CSV · Customize panel). Phase 1 of "own the spine, keep the systems as
peripherals." bill-tracker: FULL pull incl. subs → subs to the QBO Audit sheet, which gains an
FW-misplacement + sub-missing-project section and folds in the retired
duplicate/item-no-project/sub-bill audit scripts; cost codes captured audit-only. project-pnl:
"Open Project in QBO" header link → project home page on all three templates)

---

## The folder map

```
shared/                the ONLY importable common code
├─ qbo_vault.py        QBO Keychain blob — one Touch ID unlocks all keys
├─ paths.py            per-machine output paths (machine.env at REPO ROOT)
├─ qbo_api.py          QBO auth + retrying GET, query_all, P&L walkers, PROJ_RE
├─ qbo_costs.py        cost_leaf (the ONE cost-code resolver) + iter_cost_lines — shared w/ ledger
├─ notion_client.py    thin Notion API client (create/query/update pages) — used by ledger/sync_actions
├─ cost_lines.py       cost-line category (Concrete/Labor/Materials) + bill-line combine
├─ draws.py            CP draw (AIA G702/G703) discovery + parsing (wip ↔ health)
├─ takeoff_etc.py      blank ETC → takeoff cost sheet (rp_wip_reader ↔ schedule preview)
├─ pnl_paths.py        resolve a project's P&L workbook + "last pulled" mtime (ledger ↔ project-pnl)
└─ setup_qbo.py        vault admin CLI (--status/--test/--rotate/--purge)

invoice-sync/          QBO → Notion AR sync + Teams cards   (was automation-worker/)
bill-tracker/          AP bills (FULL pull incl. subs) → Excel tracker + QBO Audit sheet (6 checks) + job_coding_audit drill
statement-reconciler/  vendor statement PDFs ↔ QBO open bills
wip/                   ALL WIP tooling: wip_writer.py (shared engine) + CP/RP readers + close scripts
ledger/                canonical project DB: schema.sql spine + loaders (WIP · Bill Tracker · costs) + dashboard
project-pnl/           per-project P&L workbooks → OneDrive
debt-schedule/         equipment debt workbook + loan_sync (writes beside itself)
health-dashboard/      local company-health xlsx (private, chmod 600)
qbo-export/            one-row-per-line-item txn export → OneDrive inbox
job-auditor/           DESIGN + prototypes: audits proposal scope vs takeoff cost
                       bands on the live job folders (read-only; proposes, never
                       applies) — see its SPEC.md / STATUS.md
one-offs/              occasional / not-yet-developed tools (never the repo root)
synology/              NAS file-tree audit (always --exclude the sensitive path)
docker/                invoice-sync container package (v1.1.0)
docs/                  this map + system references
```

**The rules (full text in CLAUDE.md):** one folder = one tool · `shared/` is the only
importable common code · tools never import tools · one-offs live in `one-offs/` ·
`machine.env` stays at the repo root.

**Auth in one line:** every QBO script goes through the shared Keychain vault (single
Touch ID per run); QBO is production-only; writers are gated (`--commit` / `CONFIRM=Y`) —
everything else is read-only against QBO.

---

## AR — the invoice sync (busiest pipeline)

`invoice-sync/` · manual `sync-ar` · Docker v1.1.0 on Synology (15-min loop)

```mermaid
flowchart LR
    classDef src fill:#dbe6f0,stroke:#4A6B8A,color:#1f2937
    classDef tool fill:#f6f5f1,stroke:#4b5563,color:#1f2937
    classDef out fill:#dfeae2,stroke:#3E7A5C,color:#1f2937

    QBO[("QBO\nopen invoices + CDC")]:::src
    SYNC["run_invoice_sync.py"]:::tool
    MFD[("Notion\nMFD Invoice DB")]:::out
    RES[("Notion\nRes/Com Invoice DB")]:::out
    TEAMS[("Teams\nMFD paid / short-pay cards")]:::out
    XL["export_invoices_xlsx.py\n+ aging_sheet.py\n+ draw_chain.py"]:::tool
    BTX[("Bill Tracker.xlsx\nOneDrive · READ-ONLY")]:::src
    OD[("OneDrive\nOpen_Invoices.xlsx\ntab 1 Open Invoices\ntab 2 AR Aging")]:::out

    QBO --> SYNC
    SYNC -- "route by project-# prefix" --> MFD
    SYNC --> RES
    SYNC -- "sweep paid · archive QBO-deleted" --> MFD
    SYNC --> TEAMS
    SYNC --> XL --> OD
    BTX -- "MFD/CP vendor unpaid-bill status" --> XL
```

Support cast (same folder): `doctor.py` diagnostics · `verify_invoices.py` /
`verify_excel_export.py` read-only audits · `sync_view.py` visual runner ·
`setup_keychain.py` Notion/Teams secrets.

**The AR Aging tab** (`aging_sheet.py`, added 2026-08-05) is the owner's
at-a-glance collections view: QBO-style Current / 1-30 / 31-60 / 61-90 / 90+
buckets aged by due date, all three divisions in one filterable table, invoices
grouped under the parent client and collapsed by default, the collections
clerk's Notion `Quick Status` note carried across, and litigation invoices
excluded. It **reads** `Bill Tracker.xlsx` (the AP tool's output file, never its
code — repo rule 3) for what is still owed to vendors.

**The MFD/CP draw-funding chain** is what those vendor columns actually answer.
The GC funds draw N → we pay draw N's vendor bills → those vendors issue
unconditional waivers → the GC releases draw N+1. So an unpaid draw is gated by
its **predecessor**, not by its own bills. `draw_chain.py` orders each project's
draws (excluding non-draw invoices like retainage releases, and keying by
contract so parallel contracts don't interleave), and the verdict separates
`PAY BILLS → unlock` — prev draw funded but our vendors still owed, ours to fix
— from `Waiting GC on prev`, where the hold-up is upstream. Projects running
parallel contracts report `Multi-contract` rather than a guess, because bills
carry a project #, not a contract.

> **Run order: AP → AR.** That read is the only edge between the two pipelines,
> and it makes AR downstream of AP. `sync-all` runs **the bill tracker first,
> then the invoice sync** (~5 min total: AP ≈ 3.5 min, AR ≈ 1-2 min) — it was
> AR-then-AP until 2026-08-05, which would leave the vendor columns reporting
> the previous AP run. **This is a DAG, not a cycle:** the bill tracker pulls its
> bills *and* its invoices straight from QBO and never reads Notion or
> `Open_Invoices.xlsx`. AR never triggers AP; if the tracker is stale the tab
> subtitle turns red and the run logs a warning, and a missing file yields `?`
> rather than a false "Vendors Paid".

---

## AP — bill tracker & statement reconciler

`bill-tracker/` · manual `sync-ap` &nbsp;|&nbsp; `statement-reconciler/`

```mermaid
flowchart LR
    classDef src fill:#dbe6f0,stroke:#4A6B8A,color:#1f2937
    classDef tool fill:#f6f5f1,stroke:#4b5563,color:#1f2937
    classDef out fill:#dfeae2,stroke:#3E7A5C,color:#1f2937

    QBO[("QBO\nALL bills (incl. subs) + invoices")]:::src
    NAS[("Synology NAS\nvendor statement PDFs")]:::src
    GL[("General List xlsx\nSynology · READ-ONLY")]:::src
    BT["excel_bill_sync.py\nBills/Inventory/Liens + 6 Audit- Table sheets"]:::tool
    JCA["job_coding_audit.py\non-demand per-job drill"]:::tool
    SR["statement_reconciler.py"]:::tool
    BX[("Bill Tracker.xlsx\nOneDrive/Automations-\ndisplay = non-sub · audit = incl. subs")]:::out
    RX[("reconciliation xlsx\n→ back to NAS")]:::out

    QBO --> BT --> BX
    QBO -. "audit-job" .-> JCA
    GL -- "RP draw matching" --> BT
    QBO --> SR
    NAS --> SR --> RX
```

Full pull (2026-08-06): the tracker pulls every bill incl. subs. Subs are kept off the
Bills/Inventory/Liens sheets but flow to the audit, now **six `Audit - …` sheets**, each a
proper Excel Table (filter/sort): Not Approved · Data Entry · Missing Project · Duplicates ·
**FW Misplaced** · Sub No Project. The old `duplicate_bill_audit` / `item_no_project_audit`
/ `sub_bill_audit` scripts were folded in and retired; `job_coding_audit.py` remains as the
interactive `audit-job` drill. Cost codes (QBO Item name) are captured for the audit only —
never a display column.

---

## WIP — readers & close scripts (all in `wip/`)

```mermaid
flowchart LR
    classDef src fill:#dbe6f0,stroke:#4A6B8A,color:#1f2937
    classDef tool fill:#f6f5f1,stroke:#4b5563,color:#1f2937
    classDef out fill:#dfeae2,stroke:#3E7A5C,color:#1f2937
    classDef gate fill:#f3e0d3,stroke:#B9541E,color:#7c2d12

    FOLDERS[("Project folders\ntakeoffs · draws (G702)")]:::src
    GL[("General List xlsx\nAlpha + Small Jobs — READ-ONLY")]:::src
    QBO[("QBO\nvia shared/qbo_api")]:::src
    RPFIX[("Owner's RP WIP workbook\n'RP WIP' sheet — verified lines")]:::src
    WMTAB[("'WIP Master' tab\nMFD contract/ETC")]:::src
    TKETC["shared/takeoff_etc.py\nfind_takeoff_etc: blank ETC → takeoff\ncost sheet (SL+PR / FW / BID)"]:::tool

    ENGINE["wip_writer.py\nthe SHARED report ENGINE:\nCpRow · COLS · write_test_cp ·\nformatting · change audit ·\nedit-tracking · QC\n(guarded by wip_excel_guard.py)"]:::tool
    CPR["cp_wip_reader.py\nCP READER: folder scan / draws /\nproposal PDF → 'Test - CP'"]:::tool
    RPR["rp_wip_reader.py\nRP READER: owner's file → 'Test - RP'"]:::tool
    MASTER["master_wip_test.py\nORCHESTRATOR: MFD + CP + RP → 'Test-Master'\n(+ change audit)"]:::tool
    TEST[("WIP - MASTER new.xlsx\nTest tabs ONLY (SharePoint)")]:::out
    CLOSE["qbo_close_list.py →\nqbo_bulk_close.py"]:::tool
    QW[("QBO WRITE — gated\nCONFIRM=Y · MFD always excluded")]:::gate

    FOLDERS --> CPR
    RPFIX --> RPR
    FOLDERS --> TKETC
    TKETC -.->|"blank ETC only\n(estimator entry wins)"| RPR
    QBO --> CPR & RPR & MASTER
    WMTAB --> MASTER
    CPR -.->|"scan reused by"| MASTER
    RPR -.->|"classify/write reused by"| MASTER
    CPR & RPR & MASTER ==>|"import the engine"| ENGINE
    ENGINE --> TEST
    AUDIT["wip_audit.py\n--audit: pre-write provenance report\nadds/removes + reason · contract/ETC source"]:::out
    MASTER -.->|"--audit (inspect, NO write)"| AUDIT
    QBO --> CLOSE --> QW
```

**The three readers import ONE engine (`wip_writer.py`), never each other** (2026-08-04).
`wip_writer` owns everything that turns `CpRow`s into a formatted, audited, edit-tracked
tab — `CpRow`, `COLS`, `write_test_cp`, the change audit, the QC check. The readers only
gather their division's numbers: `cp_wip_reader` (CP folder scan / draws / proposal PDF),
`rp_wip_reader` (the owner's RP file), and `master_wip_test` (MFD off the 'WIP Master' tab,
and orchestrates all three). This replaced the old shape where the engine lived inside
`cp_wip_reader` and every tool did `import cp_wip_reader` — a tool importing a tool, which
buried MFD/RP logic in a file named "cp" and let the layout drift.

Over/under-billing and job-borrow are computed columns in Excel, not in these scripts.
RP v2 (2026-07-13): the General List is the RP source — each RP job auto-splits into
`RP####` (slab) + `RP####-FTW` (flatwork) lines; CP jobs in the list stay standalone.
`master_wip_test --rp-from-file <xlsx>` (2026-07-29) replaces the GL pipeline for the
RP section with the owner's verified RP WIP workbook (sections from its band rows,
duplicates deduped, CP lines excluded); billed/costs still refresh from QBO per line.

**Pre-write audit (`--audit`, 2026-08-07).** `master_wip_test --audit` runs the full
pipeline (QBO, ETC fallback, classify) and writes `wip_audit.py`'s inspect-only workbook
(`~/Downloads/WIP Audit.xlsx` by default) instead of touching the WIP report — one row per
job with its Δ vs the current report (ADDED / REMOVED / SAME) + the reason, and CONTRACT /
ETC with the exact source (owner's RP-file cell, takeoff file+cell, or blank). It answers
"where did each non-QBO value come from, and why did this job appear/disappear" before any
write. READ-ONLY on every source (safe even with the WIP file open). Provenance is captured
at read time (`rp_wip_reader` tags each row's `audit_contract_src` / `audit_etc_src`).

**Blank ETC → takeoff fallback (2026-08-07).** When the estimator left an ETC cell
blank in the RP file, `rp_wip_reader.classify_from_file` calls
`shared/takeoff_etc.find_takeoff_etc` (moved out of `one-offs/rp_schedule_wip_preview`
the moment a second tool needed it — the verified SL+PR / FW / commercial-BID extractor)
to read the budget from the job's takeoff cost sheet. The estimator's manual entry always
wins — the fallback runs only on blanks. Provenance is colour-coded on 'Test - RP' (font):
**BLUE** = estimator's manual entry, **ORANGE** = machine-read from the takeoff (verify).
A job whose folder has no takeoff cost sheet stays blank and is flagged. Both this reader
and `master_wip_test` go through `classify_from_file`, so the fallback applies identically.
Test tabs are styled to match the real 'WIP Master' sheet — **frozen, see CLAUDE.md
rail 5a**: Tahoma 8, no-cents currency, and its two-line left-aligned title block
(company prefix read from `WIP Master`!B1 at runtime; no merge-and-center banner).
ONE commentary column (`NOTES` = owner's ACTION text · script notes · must-fix flags);
`TYPE` is Tract/Custom only. All tabs open Active-only. `qbo_links_only` is the DEFAULT for every tab (2026-07-29): NO
file/Synology links anywhere — the only hyperlinks are the QBO deep links on Billed
(customer page) and Costs (project P&L report), helpers in `shared/qbo_api.py`.
`--rp-from-file` also rewrites 'Test - RP' with the same RP rows in the master
layout (revised contract/ETC across the board) plus a CATEGORY column decided from
the DATA — GOOD only when QBO shows costs/billing, else NOT STARTED (still on the
schedule) or FTW BACKLOG — with a legend block, the
owner's ACTION notes in NOTES, and his colour marks (green verified / red changed /
orange verify) re-applied to the $ cells — marked Billed/Costs values survive the
QBO refresh. Rows whose numbers don't reconcile render red (`needs_review`).

---

## Ledger — the canonical project database (`ledger/`, Phase 1)

```mermaid
flowchart LR
    classDef src fill:#dbe6f0,stroke:#4A6B8A,color:#1f2937
    classDef tool fill:#f6f5f1,stroke:#4b5563,color:#1f2937
    classDef out fill:#dfeae2,stroke:#3E7A5C,color:#1f2937
    classDef future fill:#efeaf3,stroke:#6b5b95,color:#312e40

    TEST[("WIP - MASTER new.xlsx\nTest tabs — the FINAL WIP\n(read-only source)")]:::src
    BT[("Bill Tracker.xlsx\nBills + Inventory\n(read-only source)")]:::src
    SCHEMA["schema.sql\nspine: project · cost_code · budget_line ·\ncost_line · billing_event · wip_snapshot ·\nap_bill_line · customer · sales_touch  (SQLite + Postgres)"]:::tool
    LOADER["load_wip_master.py\nCP←Test-CP · RP←Test-RP · MFD←Test-Master\nfilter to real project #s · idempotent upsert"]:::tool
    APLOAD["load_bill_tracker.py\nAP pay status + lien clock → ap_bill_line\n(NOT cost truth — subs excluded)"]:::tool
    DB[("ledger.sqlite3\nproject + wip_snapshot + ap_bill_line\n→ v_wip_latest · v_ap_by_project")]:::out
    DASH["dashboard.py + static/\nlocal web UI (127.0.0.1) — READ-ONLY except one write.\nTabs: My view · Overview · Costs (code→jobs pivot) · Draws\n(race-through) · Liens · Vendors · Sales (CRM pipeline)"]:::tool
    BROWSER[("Browser\nhttp://127.0.0.1:8787")]:::out
    QBO[("QBO\nBills + Purchases\n(read-only pull, Touch ID)")]:::src
    QCOSTS["shared/qbo_costs.py\ncost_leaf + iter_cost_lines\n(the ONE resolver — shared with project-pnl)"]:::tool
    COSTLOAD["load_costs.py\ncost_line by cost code · incl. subs ·\nreconciles to wip_snapshot · --selftest"]:::tool
    FUTURE["later: budget_line (takeoff by code)\n· billing_event (AR / draws)"]:::future

    TEST --> LOADER
    BT --> APLOAD
    QBO --> COSTLOAD
    QCOSTS -.->|"cost_leaf"| COSTLOAD
    SCHEMA -.->|"applied on connect"| LOADER
    LOADER ==>|"project · wip_snapshot"| DB
    APLOAD ==>|"ap_bill_line"| DB
    COSTLOAD ==>|"cost_line · cost_code"| DB
    DB ==>|"read-only"| DASH --> BROWSER
    DASH -.->|"the ONE write: waiver mark"| DB
    SYNCACT["sync_actions.py\naction items → Notion pages\n(shared/notion_client)"]:::tool
    NOTION[("Notion 'Ledger Actions' DB\nthe folder-memory per action")]:::out
    DB --> SYNCACT --> NOTION
    NOTION -.->|"Status readback"| DB
    NCL[("Notion 'Customer List'\nCRM leads/clients + touch notes\n(read-only source)")]:::src
    CUSTLOAD["load_customers.py\ncustomer + sales_touch (CRM)\nnotes → touch log · created/last-edited-by\nread-only · --selftest"]:::tool
    NCL --> CUSTLOAD ==>|"customer · sales_touch"| DB
    FUTURE -.-> DB
```

`shared/qbo_costs.py` is the same `cost_leaf` resolver **project-pnl** uses (moved there 2026-08-08
so a second tool could share it), so cost-code figures tie between the ledger and the P&L export.

**The idea:** own the spine, keep the systems as peripherals. QBO stays the books,
JobTread stays the ops shell, Excel goes back to being an export — but the reconciled
shape of a JOB lives in `ledger/`, in one place we own. `schema.sql` is the whole
6-table spine and is **portable across SQLite (the zero-install spike) and Postgres
(the real deploy) unchanged** — natural keys only, ISO-text timestamps, 0/1 booleans,
`ON CONFLICT` upserts.

**Phase 1 (done): `load_wip_master.py`** reads the FINAL WIP master (the Test tabs) and
lands `project` + `wip_snapshot` — no QBO generation, the sheet is the source. Each project
is read from its richest tab exactly once (CP←`Test - CP`, RP←`Test - RP`, MFD←`Test-Master`),
rows filtered to real project #s (`^(MFD|CP|RP)\d+(-FTW)?$`) so every legend/total/section row
drops out. Excel is opened **read-only**; upserts are idempotent by `project_no` /
`(project_no, report_date)`. `v_wip_latest` joins each project to its most-recent snapshot —
the portfolio rollup that's rebuilt in Excel today becomes one query.

**The dashboard (`dashboard.py` + `static/`)** is the browser UI over that query — a stdlib web
server (no Flask) that reads the ledger **read-only** and binds to `127.0.0.1` only. It serves
portfolio KPIs, a **Needs-attention** widget (generic exposure rules — underbilled / overbilled /
over-budget / borrowing-cash — that click-filter the table), the division rollup, and a searchable /
filterable / sortable projects table with click-into-job detail, click-to-copy cells, and CSV export. A **Customize** panel (theme, accent,
font, text size, density, width, widget + column visibility) is saved per person in `localStorage`.
Run it with `python3 ledger/dashboard.py` (the preview sandbox can't — it needs the DB + `shared/`
outside `.preview`). This is Rung 1 of turning the terminal DB into a platform; Postgres + a shared
server is Rung 2, when a second person needs to log in.

The **Liens** tab is a single filtered worklist: clickable stage tiles (Past-due / ≤7d / ≤15d / ≤30d /
Notice-sent / Lien-filed) filter one table keyed **CP # · Draw # · Name/Address · Invoice # · Amount**.
The job-detail panel links to **project-pnl**: `shared/pnl_paths.py` finds the project's
`Project_PnL_<proj>.xlsx` (+ its "last pulled" mtime); the dashboard can **open** it (`open`) or
**generate/refresh** it by shelling out to `project-pnl/run_pnl.sh` (a subprocess — tools never import
tools — gated behind a `confirm`; QBO stays read-only, only the .xlsx is written). This is the ledger's
first reach OUT to a peripheral tool; the "own the spine" inverse (project-pnl reading `cost_line` from
the ledger) is still ahead.

**On/off, on demand (owner: no always-on).** The dashboard is launched, not resident. `open_ledger.command`
starts it if down + opens the browser; `build_ledger_app.command` osacompiles a stay-open **Project
Ledger.app** whose Dock icon IS the on/off indicator — Quit / logout / shutdown / sleep all stop the
server (`pkill -f ledger/dashboard.py`). The My-view freshness strip flags a source **⟳ Sync recommended**
when it's stale > 48 **business**-hours (`businessHoursSince` — weekends don't age the data).

**AP + liens (`load_bill_tracker.py` → `ap_bill_line`).** The line-level `Bills`/`Inventory`
sheets of `Bill Tracker.xlsx` load into `ap_bill_line` — vendor, project, account, open balance,
pay status, and the Texas lien clock per bill. This is **not** the cost ledger: Bill Tracker's
display sheets exclude subs, so it runs 25–98% short of the QBO WIP cost per job (measured). Job
cost stays in `wip_snapshot`; what this uniquely adds is AP pay status + lien deadlines, surfaced
as the dashboard's **AP & liens** widget (open-AP stats + a lien watchlist ordered by urgency) and
an AP line in each job's detail. Full-replace by `source` each run; `project_no` is a soft link
(no FK), so off-WIP/closed jobs are kept.

**Complete costs (`load_costs.py` → `cost_line`, built; awaiting the first live pull).** Pulls QBO
Bills + Purchases and writes one `cost_line` per expense line — attributed to a project by its line
`CustomerRef`, keyed to the cost code by `shared/qbo_costs.cost_leaf` (the SAME resolver project-pnl
uses, moved to `shared/` so the two can't drift). Subs are included, and it **reconciles** to
`wip_snapshot.costs_to_date` per project. `--selftest` proves the whole pipeline offline (no QBO);
a real load needs one Touch ID. Scoped full-replace by `source='qbo'` (idempotent, drops deleted
txns). This is what Bill Tracker couldn't be — the complete, cost-code-keyed cost ledger.

**CRM / sales pipeline (`load_customers.py` → `customer` + `sales_touch`).** The pre-project spine:
the ledger owns the client/lead master too, not just the job. Reads the Notion "Customer List" data
source **read-only** and lands one `customer` row per page (identity + current pipeline stage +
Notion's own `Created by` / `Last edited by` system fields = who sourced it / who worked it last —
honest per-rep attribution with no manual Owner property to maintain) plus one `sales_touch` row per
"History of interactions" line in the page body (the outreach touch log, date parsed when present).
`v_sales_pipeline` (counts by stage) and `v_sales_by_rep` (activity by last editor) turn "what has the
outreach rep done" into a query instead of a scraped spreadsheet. Idempotent full-replace by
`source='notion_customer_list'`; `--selftest` proves the parse+load offline (no Notion). Auth reuses
the shared `NotionClient` token (the same one `sync_actions.py` uses). Not yet joined to `project` —
leads become jobs downstream — but it puts sales activity in the same database as the WIP.

**Still later:** `budget_line` (takeoff budget by cost code → budget-vs-actual from the spine) and
`billing_event` (AR / draws). Once both exist, over/under-billing computes from the spine, not Excel.
The dashboard **Sales tab** (built 2026-08-09) surfaces this read-only via `_fetch_sales` →
`v_sales_pipeline` / `v_sales_by_rep` / `sales_touch`: pipeline funnel, activity by rep (last-editor
attribution, the invoice-sync bot shown as "Automation (sync)"), warm-account cards with each
account's full touch log, and a searchable/filterable all-customers table linking out to Notion.

---

## Finance exports (read-only pulls → workbooks)

```mermaid
flowchart LR
    classDef src fill:#dbe6f0,stroke:#4A6B8A,color:#1f2937
    classDef tool fill:#f6f5f1,stroke:#4b5563,color:#1f2937
    classDef out fill:#dfeae2,stroke:#3E7A5C,color:#1f2937
    classDef gate fill:#f3e0d3,stroke:#B9541E,color:#7c2d12

    QBO[("QBO")]:::src
    WIPM[("WIP - MASTER new.xlsx\nTest-Master tab (SharePoint)")]:::src
    TKO[("Takeoffs (Synology)\nCP 'Cost Code' · 'CONCRETE YARDS' · RP 'Cost Gral'")]:::src
    G702[("Pay applications (Synology)\n<CP job>/Draws/Draw #N/*Pay App.xls")]:::src
    PNL["project-pnl/\nproject_pnl_export.py"]:::tool
    LOAN["debt-schedule/\nloan_sync.py"]:::tool
    HEALTH["health-dashboard/\nqbo_health.py"]:::tool
    EXP["qbo-export/\nqbo_export.py"]:::tool
    RECODE["one-offs/\nqbo_recode_review.py"]:::tool
    LOANS["one-offs/\nloans_to_subs_audit.py"]:::tool

    P1[("OneDrive\nPROJECT P&Ls")]:::out
    P2[("Equipment_Debt_Schedule_v2.xlsx\nbeside the script")]:::out
    P3[("health xlsx\nprivate · chmod 600")]:::out
    P4[("OneDrive\n-Inbox- Project Report Exports")]:::out
    P5[("QBO WRITE — gated\nxlsx audit · Approved=Y · --commit")]:::gate
    P7[("OneDrive QBO Audits xlsx\n+ QBO WRITE — gated\nConfirm Sub-Account · --commit")]:::gate

    QBO --> PNL --> P1
    WIPM -->|"Contract/ETC/COs/STATUS auto-pull\n(typed override still wins)"| PNL
    TKO -->|"cost-code BUDGET + concrete yards\n(Budget vs Actual · Labor · Concrete)"| PNL
    G702 -->|"CONTRACT + approved COs\n(shared/draws.py — beats the WIP master)"| PNL
    QBO --> LOAN --> P2
    QBO --> HEALTH --> P3
    QBO --> EXP --> P4
    QBO --> RECODE --> P5
    QBO --> LOANS --> P7
```

**`one-offs/schedule_report.py`** (read-only) — standalone weekly crew-schedule stage
Gantt. The model logic lives in **`shared/schedule.py`** (current-week discovery/parse,
address→pricing match, job×day×stage model, stage→colour) — shared with company_tracker,
which folds the same Gantt into the one workbook + dashboard. Pricing is broadened: the
**General List** resolves address→project# across all jobs (+ bid prices on the priced
sheets) and the **WIP master** overlays clean active contracts, with a project#→contract
cross-lookup and range/token-overlap fuzzy matching. Week anchors on today (next week's
pre-loaded schedule doesn't hijack it). Reads
`…/OPERATIONS/SCHEDULE/<yr>/<month>/Schedule M-D-YY.xlsx` for the latest Mon–Fri week →
one row per job, one column per day, each cell the STAGE that day (Pour / Wreck / Forms /
…) coloured by stage. Output `~/Documents/CompanyHealth/Weekly Schedule.xlsx` +
`Weekly Schedule.html` (chmod 600).

**`one-offs/money_out_register.py`** (read-only QBO) — check register for tracking
uncashed / outstanding checks. QBO exposes every check written (BillPayment Check +
Purchase Check: check #, payee, amount, bank) but NOT cleared status, so the register
carries a user-owned `CLEARED?` (Y/N) column and is STATEFUL — a refresh preserves your
marks (merged by txn id) and prunes long-cleared checks. Output `~/Documents/CompanyHealth/
Money Out Register.xlsx` (chmod 600); the aged-&gt;30-days unmarked checks are the chase
list, surfaced on the company dashboard.

**`one-offs/sub_loc_report.py`** (read-only) — subcontractor line-of-credit float
model. Sub bills (memo ~ "sub") paid → LOC draws (BillPayment date, allocated across
each bill's line projects); client Payments → repayments. Matched **by draw period**:
MFD/CP reimburse per `(project, month)` — the sub's work month (from the bill memo's
`Period …`) against the draw's month (from the invoice memo's `Draw #N (Period …)`),
so a May draw can't offset June sub costs; RP invoices are lump per scope → matched by
project. Chronological FIFO; a same-period GC draw arriving first is genuine prefunding.
Output `~/Documents/CompanyHealth/Sub LOC Report.xlsx` (chmod 600): Summary (company
peak = LOC truly needed + avg draw→repay float), **By Division** KPI (per-division peak /
float / outstanding), Ledger (each draw stamped with its reimbursing invoice + client-paid
date), Per-Project.

### Money Bleeds — company-health exceptions (2026-07-16)

```mermaid
flowchart LR
    classDef src fill:#dbe6f0,stroke:#4A6B8A,color:#1f2937
    classDef tool fill:#f6f5f1,stroke:#4b5563,color:#1f2937
    classDef out fill:#dfeae2,stroke:#3E7A5C,color:#1f2937

    QBO[("QBO invoices + PurchaseOrders\nvia shared/qbo_api")]:::src
    WM[("WIP - MASTER new.xlsx\n'WIP Master' tab (MFD actives)\n'Test - RP' tab (RP classify)")]:::src
    MFDV[("Multi Family volume\n…/&lt;client&gt;/&lt;MFD#&gt;/PM MISC/DRAWS")]:::src
    CPV[("Common volume\nCP project folders → latest draw G702\nvia shared/draws.py")]:::src
    BT[("Bill Tracker.xlsx\nAP bills ↔ GC invoice (sync-ap)\nread-only")]:::src
    MB["health-dashboard/\nmoney_bleeds.py"]:::tool
    OUT[("Money Bleeds.xlsx\n~/Documents/CompanyHealth\nprivate · chmod 600")]:::out

    QBO --> MB
    WM --> MB
    MFDV --> MB
    CPV --> MB
    BT --> MB
    MB --> OUT
```

**`health-dashboard/company_tracker.py` + `company_dashboard.py`** (read-only, no QBO) —
the consolidation. `company_dashboard.py` holds the shared readers and `build_sections()`
— the ONE metric model (MONEY IN / MONEY OUT / POSITION) computed from the tracker
workbooks (Money Bleeds, Sub LOC, Money Out Register, health_dashboard, WIP master
Test-Master): AR/backlog/retainage/lien in; AP/POs/checks/LOC/burn out; cash/runway/
coverage/margins/over-under position. `company_tracker.py` folds that model into a single
**`Company Tracker.xlsx`** (Summary + Money In/Out/Position tabs + a **Weekly Schedule**
Gantt tab via `shared/schedule.py`, an **RP Billing Status** tab via `shared/rp_billing.py`, hero numbers, aging + LOC data bars, semantic colour)
AND renders `Company Dashboard.html` (with the Gantt section) from the same model, so the
workbook and page never disagree. The source Excels are the data layer; this is the
one-workbook-a-HTML-breaks-down deliverable.

Read-only exception checks: **draws with no invoice** (MFD: latest numbered draw folder
vs latest QBO invoice date; CP: latest draw's G702 earned-less-retainage vs cumulative
QBO invoiced), the **Texas lien-notice clock** on every open construction invoice
(commercial 15th-of-3rd-month, residential 15th-of-2nd; work month = invoice month;
OK rows hidden; retainage + equipment-lease/note invoices split off their own sheets),
**RP wrap-up** (slab lines 100% in the General List but not fully billed, from the
Test - RP tab), **unused POs** (open QBO purchase orders ≥30 days old with no bill
linked), and **open bills (AP)** (read from the Bill Tracker, never recomputed —
grouped by AR state × the sub's lien clock). CP draw discovery + G702 parsing live in
**`shared/draws.py`** (shared with `wip/cp_wip_reader.py` — tools never import tools).
The rich Excel (grouped bands, data bars, color scales, colored tabs) is a deliberate
exception to the plain-Excel rule — the user asked for an at-a-glance watchboard.

project-pnl reads the WIP master's **Test-Master** tab (the readers' unified MFD+CP+RP
table) to pre-fill **Original Contract / ETC + Approved COs** (original = total − COs;
revised rows are live formulas) and honor a **Closed** status (WIP close-out: % complete
forced to 100%). Company overhead default is **10% of revenue** (MFD alt: 9% on costs —
also drives the MFD draw-coverage columns). Its Transactions sheet groups job costs
**concrete → labor → materials** via `shared/cost_lines.py` — non-labor bill lines
combine per (bill × account); labor lines are never combined. The **Budget vs Actual**
**Labor** and **Concrete** sheets (CP) are the PM/ops manager's main view (the user
2026-07-29): rows are that trade's cost codes (`*6` labor, `*1` concrete), columns are the
draw windows, and every code expands (outline ±) to the bills behind it with the amount in
the draw column it landed in — so a code row is visibly the sum of the rows beneath it.
Sales tax is billed on its own vendor line and is pulled OUT of the comparison (the takeoff
budget is pre-tax), then summed at the bottom with each contributing line; Concrete adds
yards and $/yd against the takeoff's implied rate, excluding lump vendor bills that carry
no yardage. CP **contract price and approved COs come from the latest G702 pay application**
(`shared/draws.py` `read_pay_app` — legacy .xls needs `xlrd`), overriding both the WIP
master and any hand-typed cell; the P&L names the source. The **Budget vs Actual**
sheet (CP + RP) joins the takeoff's cost-code budget (CP: the 'Cost Code(s)' sheet;
RP: 'Cost Gral', FW codes → the -FTW project) against QBO cost-code actuals — jobs
with change orders carry a "may be inaccurate" banner until the CO template ships its
cost line. Sheet order: P&L · Transactions · Budget vs Actual · Next Draw · draws ·
POs · Reconciliations · Cash Flow.

---

### JobTread migration tools (read-only unless noted)

**`one-offs/jobtread_migration_setup.py`** (read-only) — the "what do we still need to
add to JobTread" workbook. Two tabs: **Active Jobs to Add** (jobs on the daily schedule
with no JobTread job) and **Bidding to Add** (Notion **Bid List** RP rows whose
`Lead Status` is not terminal — not Sold / Lost / GC Not Awarded / No Response / No
Opportunity — and not already in JobTread). Reads the schedule, JobTread (full job
sweep), and Notion via the invoice-sync integration token (Keychain
`proficient-automation-worker/notion`). Output `~/Downloads/JobTread Migration Setup.xlsx`.
Overlaps `rp_jobtread_coverage.py` on the schedule half — coverage answers *does it have
an approved proposal*, this answers *does it exist at all* + the bid pipeline.

**`one-offs/jobtread_bloat_report.py`** (read-only QBO + JobTread) — open JobTread jobs
vs reality. Per open job, joins QBO (last invoice date, AR balance) + the daily schedule:
`CLOSE — paid & idle` (AR ≈ $0, idle > 90d, unscheduled), `CLOSE? — no QBO, stale`
(no invoices, created > 120d, unscheduled), `DONE? unpaid` (idle but money owed — an AR
chase, not bloat), `ACTIVE`/`NEW`. `-FTW` jobs fall back to the base project's QBO record
(flatwork often bills under the base #). Output
`~/Downloads/JobTread Bloat - Close Candidates.xlsx`.

**`one-offs/jobtread_close_jobs.py`** (**writes to JobTread**, audit-gated) — closes the
approved bloat. `--export` turns the bloat report into an APPROVE workbook with a
`CLOSE? (Y/N)` column (pre-filled Y only for the high-confidence bucket); the user edits it;
`--apply` dry-runs; `--apply --commit` performs it. **A job is closed by setting
`closedOn`** (its `status` is not directly settable) — verified reversible: clearing
`closedOn` restores the prior status, so `--reopen` is a true undo. Nothing is ever
deleted. MFD excluded by default. Every change logged before/after to
`~/Library/Logs/Proficient/jobtread_close_*.jsonl`.

**`one-offs/cable_calculator.py`** (read-only) — the PT-cable cut-list + cost engine
lifted out of the RP takeoff's hidden `'0'` sheet, mapped cell-for-cell and validated
against 64 active takeoffs (count, LF, every cut-list row and cost tie exactly). Inputs
live in **two** INFORMATION blocks (`I/K/M` 25-69 and `N/P/R` 24-69); cost rows are found
**by label** ("Cable take off" / "Per cable") because their row and rate vary per file.
The takeoff only sorts+sums typed pairs — it derives nothing — so this is a verification
harness for the JobTread migration, not an estimator tool.


**`one-offs/rp_wip_update.py`** (read-only sources, writes the OWNER'S file) — the RP WIP is
no longer a generated report; it is the owner's working file in OneDrive
(`RP WIP TO FIX_Final.xlsx`, sheet `RP WIP`). This OPENS that file and refreshes only the
machine columns (SCHEDULE ✓/✗, QBO billed/costs, GP%), appending schedule lines it doesn't
have. **The owner's colour marks are authoritative and are never overwritten** — theme 9
(orange) = ops manager must verify, `00B050` = verified, `FF0000` = he changed it. Dry run
by default; `--commit` writes back to OneDrive.

**`one-offs/jobtread_schedule_writer.py`** (**writes to JobTread**, audit-gated) — the daily
Excel schedule → JobTread dated tasks, using the estimator's own stage names (Wreck/Clean,
Set Forms, Trench, Grade/Backout, Drill Piers, Pour, Tension Cables). UPSERTs by (job, task
name): existing task → update dates only if changed; none → create. Never deletes, never
touches an unmapped task. Dry run by default.

**`one-offs/rp_wip_simple.py`** (read-only, writes the WIP master's `Test - RP` tab) — the
stripped-down RP WIP: contract · ETC · billed · costs · GP%, plus SCHEDULE / GENERAL LISTA
marks and an ACTION column. ETC is computed from the takeoff's **cost code rows**, never the
pier subtotal cell (that cell means slab+piers on some takeoffs and piers-only on others);
FW is never added to a non-FTW ETC. Sections: CP-on-RP-schedule, dropped-but-unbilled,
⚠ FTW-with-costs (not backlog), FTW backlog. **Writes no `file://` hyperlinks** — Excel
resolves those through ScopedBookmarkAgent at load and beachballs on a network share.

**Schedule file choice (all four tools):** `shared/schedule.schedule_on_or_before()` — never
reads a schedule dated in the future. The team pre-loads tomorrow's board, and a plain
"highest filename wins" pick silently jumps to it. `--as-of YYYY-MM-DD` overrides.

## Who writes to QBO (the short list)

| Script | What it writes | Gate |
|---|---|---|
| `one-offs/qbo_recode_review.py --apply` | line Customer:Project + Class on job-cost lines | xlsx audit, `Approved=Y` rows only, then `--commit` |
| `one-offs/loans_to_subs_audit.py --apply` | line `AccountRef` (parent `Loans to Sub-Contractors` → per-sub sub-account) on Bill/Purchase/VendorCredit | xlsx audit, `Confirm Sub-Account` filled, then `--commit`; skips stale-SyncToken / closed-period / not-still-on-parent |
| `wip/qbo_bulk_close.py` | closes customers/projects | `CONFIRM=Y`; always excludes MFD (manual close) |

Everything else is read-only against QBO. All other "writes" land in Excel, Notion, Teams,
or JobTread.

## Who writes to JobTread

| Script | What it writes | Gate |
|---|---|---|
| `one-offs/jobtread_close_jobs.py --apply` | a job's `closedOn` date (= closes it; clearing it reopens) | xlsx audit, `CLOSE?=Y` rows only, then `--commit`; excludes MFD; never deletes; `--reopen` undoes |
| `one-offs/jobtread_schedule_writer.py --commit` | task name + start/end date + task type on a job | dry run by default; upsert only, never deletes |

**What the JobTread API CANNOT do** (verified 2026-07-29/30, do not build on these):
`createDocument` (proposals/COs/invoices) is rejected outright — *"A job location name or
address is required"* even on jobs with a complete location. `updateCostItem(costCodeId)` on a
QBO-synced bill **returns success and changes nothing**. Bill UPDATES sync in neither
direction; the QBO push fires on **approval**, so corrections are free while `draft`/`pending`
and brutal once approved+paid. JobTread is the operational shell — QBO stays the cost truth.

The JobTread grant key (`JT_GRANT_KEY`, shared vault) has **write** access — grants carry
no read/write scope, they inherit the user's permissions. Verified 2026-07-27 by a
create → read-back → delete round-trip on the CP000 placeholder job.

---

## Machine notes

- Output paths (OneDrive folders, `~/Documents/CompanyHealth/`) resolve through
  **`shared/paths.py`**: process env > `machine.env` (REPO ROOT, gitignored, per-machine) >
  owner's original defaults. New machines: `cp machine.env.example machine.env`, then
  **`python3 shared/paths.py`** — it prints where each path resolved from and flags anything
  missing or not writable. Never hardcode a machine path in a script.
- **Migration note (old clones):** after pulling the restructure, untracked files linger in the
  old `automation-worker/` folder. Move `.env` and `state/` into `invoice-sync/`, then delete
  the leftover folder. Until then, `invoice-sync/config.py` falls back to the legacy location
  automatically (the CDC watermark is never lost).
- Logs stay at `~/Library/Logs/Proficient/automation-worker/` (historical name kept on purpose)
  and always outside the repo. The Keychain service `proficient-automation-worker` also keeps
  its name — renaming it would orphan stored secrets.
- Each developer has their own Intuit app + own Keychain vault. One app connection = one machine.
