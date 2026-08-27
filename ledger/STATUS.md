# ledger/ — STATUS

Progression record for the canonical project database. Update in the SAME commit as any
change to this tool (repo rule). Tool-scope only — business/dollar analyses live in the vault.

## DONE / FINALIZED
- **UX overhaul for efficiency (owner, 2026-08-27: "visual improvements ... efficiency ... i need
  console first ... graph is cluttered/glitchy ... my view and overview are basically the same,
  one overview ... fold the bill tracker audit sheet in for Accounting fixes, filterable ... open
  invoices, two views: amounts + aging").**
  - **Console is the first nav group.** **My view + Overview merged** into one Overview - freshness +
    Resync + action items now sit on top of the Portfolio KPIs / Needs-attention / divisions /
    projects (the duplicate "Working on" list dropped; renderHome's active-projects block guards off).
  - **Graph tab removed** - the vault org-map was an illegible hairball of internal notes; gone from
    nav + HTML. (The `vault_graph.py` backend + `/api/graph` are now dead code, left for a later prune.)
  - **NEW Accounting tab** (Vendor group): the three Bill Tracker audit sheets folded into one
    filterable list. `_fetch_accounting_audits` reads `Audit - Coding/PO/Bills` from Bill Tracker.xlsx
    live (bill-tracker computes them with the full bill data - subs + cost codes - that ap_bill_line
    lacks; the Open column is an `=HYPERLINK` formula, parsed for the QBO URL). Stat tiles by theme
    (Coding/Bills/PO), a chip per audit type with counts (Not Approved · Missing Project · FW Misplaced
    · Sub No Project · Duplicate · Missing/Unused PO · Cost Code · Data Entry · ...), search + division
    filter, QBO bill links. `/api/accounting`. Verified: 1876 findings, filter chips isolate each type.
    - **Row layout reworked (owner, 2026-08-27: "how much space is wasted ... i need to read the
      description line").** The 9-column stretched table wasted width and truncated the Detail column
      (the flagged-reason) off the right edge. Replaced with a compact 2-line list row (`.acct-item`):
      line 1 = issue pill · vendor · project · cost-code chips; line 2 = bill # · date · the FULL
      flagged-reason description, wrapping across the whole row so it never truncates; amount + one QBO
      link in a fixed right rail. Front-end only (app.js/index.html/style.css), so a browser reload picks
      it up. Verified live: description reads in full, filters + QBO deep links intact.
  - **Open invoices: Amounts | Aging toggle.** **Amounts** = a clean flat list (Client · Project ·
    Invoice # · Date · Open balance · Invoice total + a TOTAL row), shares the tab's filters + sort,
    click a row for the invoice detail. **Aging** = the existing buckets + lien clock. Amounts is the
    default; the aging-only Flatten/Collapse buttons hide in Amounts.
  - Front end reload for the UI; the Accounting tab added `/api/accounting`, so a restart picks that up.
- **One efficient sync + the P&L actually works now (owner, 2026-08-27: "the most efficient way to update
  this ledger ... the P&L is not functioning due to it needing data ... the payments section is not showing
  recent payments ... simple to sync and efficient").** Three problems, one root theme - two loaders never
  ran in the sync people actually use:
  - **Payments was wired into NO sync** (`load_payments` was in neither the dashboard pipelines nor
    `reload_ledger.sh`), so the Payments tab was frozen at the last manual run. Fix: a **`payments` pipeline**
    (rides the reload chain like every other loader) + a step in `reload_ledger.sh`. Rolling **12-month**
    window because `load_payments` DELETE+reloads its window, so the window IS the history depth (a shorter
    window silently drops older payments). Refreshed: 848 payments, recent ones included.
  - **Costs (the P&L's cost side) only ran on the dashboard Resync, never on terminal `sync-all`**, so a
    terminal-only user's P&L cost went stale. Fix: `load_costs --active --since <90d>` added to
    `reload_ledger.sh` too. So **both** paths (dashboard Resync AND `sync-all`) now leave the ledger fully
    fresh; the Console shows each source's last-load so you can SEE what's stale (`_freshness` gained
    Payments).
  - **The P&L read every job as a total loss** even with fresh costs, because `v_wip_latest` had
    `total_contract_price` / `percent_complete` NULL for ~all jobs. Root cause: those are **Excel formulas**
    on the WIP Test tabs, and openpyxl reads a formula cell as None whenever the cached values were stripped
    (every script write to the tabs does that). Fix: `load_wip_master._fill_derived` **computes** the derived
    WIP columns from the input columns (the same formulas the sheet uses: C=A+B, F=D+E, K=I/F, ...), only
    ever filling a blank. After reload: contract populated 3 -> 156 jobs, % complete 0 -> 112; the Project
    P&L went from all-losses to **$23.6M earned / net +$1.12M (4.8%)**, split MF 5.4% · CP 4.6% · RP 0.7%.
  - Backend needs a **restart** (dashboard.py + the loaders changed); the front end (Console Payments card +
    Resync wording) is a reload. Known follow-up: each QBO loader authenticates separately - a shared-session
    runner would cut that, and %complete >100% on over-budget jobs is passed through uncapped (matches Excel).
- **Console: queue multiple syncs (run in order) + a "Sync AP + AR" button (owner, 2026-08-25: "can it not
  run two at the same time ... i also need progress bar ... sync AR and AP together, both need the same
  info").** Runs are single-locked on the server (concurrent QBO pulls + ledger DELETE/INSERT would corrupt
  each other), so a 2nd Run used to just 409 into a fleeting toast. Added a **client-side queue**:
  `runPipeline` enqueues when one is running; `finishSync -> _drainQueue` starts the next automatically.
  Cards show **Running… / Queued ✕** (click a queued card to drop it), the shared progress bar names the
  running pipeline + "N queued". New **Sync AP + AR** button runs AP then AR back to back (AP first - AR's
  aging reads AP's Bill Tracker output). Verified: enqueue while running doesn't re-POST, dedups, AP orders
  before AR, buttons reflect state. Frontend-only.
- **Vendor Center jump auto-expands the vendor's bills (owner, 2026-08-25).** Clicking a vendor jumps to
  Bills filtered to them, but bills open collapsed by default, so it was an extra click.
  `jumpToVendorBills` now deletes that vendor's group key(s) from `billsCollapsed` (via `billGroupKey`
  under the current Group-by) before the final render, so their bills show immediately. Verified: ABATIX
  jump shows all 6 bills expanded.
- **Console: each sync now says what it does (owner, 2026-08-25: "I need to know what the syncs do ...
  how it grabs info and where it goes/update").** Added a plain-language description line per pipeline card
  (`PIPELINE_DESC` in `renderConsole`): what it grabs (QBO / SharePoint Excel / Notion), where it writes
  (Bill Tracker.xlsx / Notion + AR Aging / the ledger), and which tabs it feeds. Frontend-only.
- **REVERTED the chromeless app-mode browser (owner, 2026-08-25: "revert the chrome thing, it opens a new
  window I don't like").** `ledger/app/ledger_app.py` is back to `webbrowser.open(URL)` (the default
  browser). The 127.0.0.1 address bar is browser chrome; leaving it rather than forcing a new Chrome
  window. No rebuild needed to keep the current behavior.
- **Payments "Unlocks (AP)": scope by DIVISION - draw for CP/MFD, whole job for RP (owner, 2026-08-25:
  "use draw period ... do the same fix for MFD and NOT RP").** A $460K Tri-C check showed it unlocked $1.1M
  of AP because `payUnlockBills` matched open vendor bills by `project_no` (every open bill on CP800), so
  every payment on a job showed the entire backlog. Fix: **CP/MFD are STAGED** (each draw = a scope with
  its own costs + invoice), so match open bills to the DRAW the payment paid (`bill.invoice_no` ↔ the
  application's `invoice_no`, via `payOpenBillsByDraw`); **RP** is regular work (costs UP FRONT, invoiced
  ONCE at the end - bills aren't tied to a draw), so use the whole project's open AP (`payOpenBillsByProject`).
  `payUnlockBills(p, drawIdx, projIdx)` branches per application via `_payIsRP` (project# prefix / division).
  Verified: CP790 ~$11k·3 (draw) vs ~$86k·20 project; MFD192 ~$303k·11 (draw) vs ~$921k·65; RP6901-FTW
  ~$20k·3 = whole job (·N = bill count). Frontend-only (BILLS already carry `invoice_no`).
- **Payments "Net after AP" column (owner, 2026-08-25: "a new column that shows me the net income after
  ap is paid").** New rightmost column = **Amount Paid − Unlocks (AP)**: what's left of a payment once the
  vendor bills it funds are paid. Reuses the same `payUnlockBills` sum (so it's division-correct: draw for
  CP/MFD, whole job for RP). **Negative = red** (the AP owed on that draw/job exceeds the payment - it
  doesn't cover itself). No portfolio total on purpose: summing per-payment nets would double-count AP that
  two RP payments on the same job both "unlock". Frontend-only (app.js) - reload, no restart.
- **Invoice detail sidebar - read a draw/invoice without QBO (owner, 2026-08-25: "need memo on invoices if i
  click it give me the details ... for draws i want that same info in the sidebar ... i like the qbo links
  but also hate using qbo").** New `#invDetail` side panel (`openInvoiceDetail`): the invoice's **Memo**
  (headline) + Billing (amount / open / paid / status), Dates & terms (invoice / due / days-past-due / terms /
  draw period / paid date), and Lien (notice deadline / status / notice type) - all from the `billing_event`
  row already loaded, no QBO call. **Invoices tab:** the row (and the invoice #) opens it; **Draws tab:** the
  invoice # opens it (row-click still expands the vendor bills). The invoice # is now a **native-detail link**
  with a small **↗** QBO icon beside it, and the panel keeps an "Open in QuickBooks" link - QBO stays one
  click away but is no longer the default (shared `invNoCell`). Days-past-due shows only while OPEN (a paid
  draw shows its Paid date instead). Backend: `_fetch_draws` attaches the full invoice (`d.inv`) and
  `_fetch_open_invoices` adds `paid_date`/`draw_period` - **needs a restart** (the fetch changed); the front
  end is a reload. Verified light + dark, both tabs, no console errors.
- **Dark-mode legibility pass (owner, 2026-08-25: "hard to see").** Brightened the dark tokens:
  `--text-dim` #98a2b3 -> #aeb9c9 (secondary text is everywhere), `--border` #2a323d -> #35404f (faint
  separators), a touch more surface separation, and a brighter `--row-hover`. The accent is a per-user
  inline setting (default #3E7A5C, a muted green that read low-contrast on dark), so instead of fighting
  it, links (`a.qbo-link` / `.unlock-link`) are brightened toward white **in dark mode only**
  (`color-mix(accent 58%, #fff)`) - keeps the chosen hue, lifts the contrast. Light theme untouched.
- **Customer Center: avg days to pay + future cash-in forecast (owner, 2026-08-25).** New backend
  `_client_pay_speed(con)` computes per-client average days from INVOICE date to PAID date over paid
  `billing_event` rows (guarded 0-400 days), plus a portfolio average, exposed on the open-invoices payload
  as `pay_speed`. Customer Center gains an **Avg days to pay** column (dimmed `~Nd` portfolio fallback when a
  client has no paid history) and three **forecast KPI tiles** - cumulative expected cash-in **≤30d / ≤60d /
  ≤90d**, each open invoice projected to `invoice_date + client_avg_days`. Verified: portfolio avg 37d, 75
  clients with history, forecast tiles populate.
- **Open Invoices: freeze ONLY the column header + kill the white scrollbar (owner, 2026-08-25).** The
  frozen-header bounded box was pinning the aging KPI tiles too. Moved `#invStats` INSIDE `.inv-scroll`
  (above the table) so the tiles scroll away with the list while the sticky column header still freezes;
  filters stay fixed above (their dropdowns need to escape the scroll). Enlarged the box (tiles no longer
  sit above it). Added **dark, subtle scrollbars** app-wide (`.table-scroll` / `.panel-body` / `.msel-menu`)
  - the default light track read as a "white bar" in dark mode. Verified: tiles scroll away, header pins at
  container-top 0, filters persist.
- **Payments tab: fixed blank + weeks/months breakdown (owner, 2026-08-25).** The tab rendered its KPI
  strip but a BLANK table. Root cause: the Pay Bills tab (added earlier this session) reused
  `id="payTable"`, colliding with the Payments table's own `id="payTable"`, so `buildHead("#payTable")`
  filled the hidden Pay Bills table. Renamed the Pay Bills table to `payBillsTable`. Then added the
  owner's ask: a **Flat / Weeks / Months** segmented toggle (default Months) that bands payments by period
  with a per-period cash-in total (e.g. "Aug 2026 · $1.2M · 31 payments"), plus a **Date** column. Click a
  payment still expands the invoices it paid; **Unlocks (AP)** still shows the vendor bills (costs) that
  cash frees up. Verified live: 847 payments, 13 month bands / 53 week bands, Pay Bills unaffected (745
  rows).
- **`sync-all` now reloads the ledger too - one command (owner, 2026-08-24: "the ledger cannot be another
  thing I HAVE to update ... I sync-all and expect the ledger to sync too, why not one?").** Root cause of
  a stale invoice (34318 showed open in the dashboard though paid in QBO): the terminal `sync-all` ran only
  the PRODUCERS (QBO -> Bill Tracker.xlsx, QBO -> Notion), never the ledger LOADERS - so the dashboard's
  ledger stayed on the last load while Notion/Excel were current. (Confirmed the split: Notion had 34318
  Paid; the ledger was the Aug-19 snapshot.) Fix: new **`ledger/reload_ledger.sh`** runs the loader half
  (WIP · bills · invoices `--no-qbo` · customers - and (added 2026-08-27, see
  below) costs + payments; continue-on-error with a summary. Wired into the machine's `sync-all` as step **3/3**
  (`~/.zshrc`, not repo-tracked; skipped on `--dry-run`). So `sync-all` = AP -> AR -> **ledger reload ->
  dashboard**. Verified: reload ran clean (exit 0); 34318 now `Paid · $0` in the ledger. Same loaders the
  dashboard's Resync("reload") runs, so terminal + dashboard can't drift.
- **Open Invoices: lien-deadline filter + dropped the aging blurb (owner, 2026-08-21).** Removed the
  paragraph describing the aging buckets (the tiles are self-explanatory). Added a **Lien deadline**
  filter on the computed lien-notice CLOCK (`lien_due_state`, from `shared/lien_clock`), separate from the
  existing Notion **Lien status** filter (renamed for clarity): **Past due (missed)** = `PAST`,
  **Upcoming (due soon)** = `URGENT` + `WATCH`, **Notice sent** = `SENT`. "Upcoming" covers CP draws, CP
  **retainage** (RET-banded 30-day clock) and RP - the clock runs for every division. Verified live: Past
  due -> 38 (CP 22 / RP 12 / MFD 4), Upcoming -> 20 spanning all three divisions; no console errors.
- **Open Invoices: multi-select filters + project sub-grouping + a client statement to copy/paste
  (owner, 2026-08-21).** Four asks, all frontend:
  - **Client + Project # are now multi-selects** (the generic `buildMSel` component, search + Select
    all/None), replacing the free-text boxes. `invPasses` uses `mselPasses`; Customer Center's
    jump-to-client sets the Client set. `_invMSelSig` guards the rebuild so a toggle keeps its search.
  - **Sub-group each client's invoices by PROJECT** (default) when the client has >1 project - an
    indented sub-band per project (project # · name · subtotal · count). A **Flatten / Group by project**
    button toggles back to the original flat list. Grand total is unaffected either way.
  - **Frozen header**: the invoices `.table-scroll` gets `.inv-scroll` (bounded height) so the base
    `.grid` sticky thead freezes as you scroll (verified: pinned at offset 0 after a scroll).
  - **"Copy for client" statement** (a right-side panel, the "different view"): a flat, searchable list of
    the filtered open invoices. **Nothing is selected to start** (owner 2026-08-21: auto-checking all was
    hard to work with) - a **search box** narrows by project / address / invoice # / client, and
    **Select all / None** act on the filtered rows; tick individual rows too. Live "N of M selected" +
    "Copy table (N)" (disabled at 0). The Client column shows only in the multi-client view; a single-client
    statement fits without horizontal scroll. **Copy writes BOTH formats** via `ClipboardItem` - `text/html` (a bordered table
    that pastes into an email) and `text/plain` TSV (pastes into Excel as cells), fallback to plain copy.
    Client-facing columns only: Client · Project (# + name) · Invoice # · Invoice date · Due date · Days
    past due · Amount due + Total due; **internal columns (lien clock, litigation) are deliberately left
    off**. Verified live: NEWLEAF (5 projects) sub-bands render, uncheck updates the total, TSV/HTML build
    correctly and sum to the shown total; no console errors. Frontend-only - a browser reload picks it up.
- **Project P&L: the invoices behind billed-to-date (owner, 2026-08-21: "in project P&L I need my billed
  to date, I need to see all the invoices the project has").** The P&L expansion (`buildPnlGroup`, shown
  in both the P&L-tab inline expand and the project detail panel) now lists **every AR invoice (draw) the
  project has** right under "Billed to GC (AR)" - oldest first, each row = invoice date (mm/dd/yyyy) ·
  invoice # (QBO deep link) · amount · Paid/Open (Open = the GC still owes; tooltip carries the paid date
  or the open AR balance). Server: `_project_pnl` returns an `invoices` list from `billing_event` (the AR
  ledger, paid + open) alongside the existing billed total; the amounts sum to billed-to-date. Verified
  live: CP790 shows 8 draws (6 Paid, 2 Open) summing exactly to its billed figure; no console errors.
- **Pay Bills: default = ALL open bills + the same rich multi-select filters as Bills (owner, 2026-08-21:
  "why put a default view for non-approved bills? I need to see all open bills as default and filter down
  just like bills - client, approved, liens, project, division, same multi-select").** Default `Show` is
  now **Open bills** (every bill with an open balance, 649), not Approved-only. Added **Client · Vendor ·
  Division · Approved · Lien** multi-selects (search + Select all/None), matching the Bills tab; Project is
  the search box. Rather than a 4th copy of the msel builder, added ONE **generic** component
  (`buildMSel` / `mselBulk` / `mselLabelUpdate` / `mselPasses`, driven by a caller-owned `store` + onChange)
  and a `PAY_MSEL` config; the older Bills/Liens builders predate it and stay as-is. Filters are built once
  per data change (signature guard) so a checkbox toggle keeps its open search box. Verified live: default
  649 open; Approved→397, +CP→101, clear→649; client search 113→1 and survives a toggle; menu open/close +
  Clear both reset cleanly; no console errors.
- **Dates → numeric mm/dd/yyyy everywhere in the dashboard (owner, 2026-08-21: "draws shows invoices
  with full date, I just need it mm/dd/yyyy for all formatting everywhere").** Replaced the long
  weekday+abbr-month style. `fmtDate(v, withTime)` now emits `08/07/2026` (was `Fri, Aug 7, 2026`) and
  `fmtDateShort(v)` emits `06/23/2026` (was 2-digit `06/23/26`); the `_DOW`/`_MON` arrays are removed. Both
  helpers cover every displayed date (meta line, Draws, WIP report date, detail panels, Bills/Pay dates).
  Also added a **Bill date** column to the Pay Bills **pay list** (the pay table already had it) so the
  generated check run carries dates. Core "never year-first" rule is unchanged - this is a style switch
  from readable to compact numeric. Excel outputs are a separate surface and were NOT touched (WIP header
  stays frozen to `MON DD, YYYY`). Updated memory [[dates-never-year-first]]. Verified live: Draws shows
  `05/05/2026` with zero weekday format left; meta line + pay list dates all mm/dd/yyyy.
- **Checklist filters: working search + Select all / None (owner, 2026-08-21: "you are expecting me to
  select through countless vendors ... select all/deselect so I can just select the few I need ... when
  I type a vendor name it doesn't filter down").** Two fixes to the `.msel` multi-select component
  (Bills vendor/client, Liens client/vendor):
  - **Search now filters the list.** Same trap the author already fixed for `.msel-menu`: `.msel-opt`
    had `display:flex`, which beat the `hidden` attribute the search sets, so typing narrowed nothing.
    Added `.msel-opt[hidden] { display:none }`. Typing a name now collapses the checklist to matches.
  - **Select all / None** toolbar on the big (searchable) checklists, acting on the VISIBLE
    (search-filtered) options in place - so you type a name, hit None on the rest, then check the few
    you want. A live "N shown / selected" count sits in the toolbar. Vendor also keeps a
    "Reset to default (hide pumps)" link.
  - **Toggling a box no longer rebuilds the menu** (label/count update in place), so the search box and
    scroll position survive a click (they were resetting before). Verified live: vendor search 89→5 for
    "ready mix", None hides the visible set, single toggle is exact (4→5→4), search survives a toggle.
- **Lazy-tab refresh fix (found while building Pay Bills).** `setTab(savedTab)` runs at init BEFORE
  `load()`, and `render()` does not re-dispatch the tabs that read the main `/api/data` globals
  (`wip` / `payments` / `paybills`) - so a fresh refresh sitting on one of them rendered EMPTY until you
  interacted. `load()` now calls `_renderLazyTab(activeTab)` after `render()`. Verified: a refresh with
  the saved tab = Pay Bills now shows all 397 rows immediately (was "No AP data").
- **Pay Bills tab - a dedicated check-run worksheet (owner, 2026-08-21: "mark bills for payment, hit
  save → generate the list + amount to be paid ... ability to change the amount / partial paid ...
  a dedicated paying bills page ... still see invoice paid, what invoice / client it goes to, lien
  notice sent/filed"; kept OFF the Bills tab on purpose - "too crowded / accidentally clicked").**
  New **Pay Bills** tab in the Vendor group (Vendor Center · Bills · **Pay Bills** · Sub LOC · Liens).
  - **Boundary (safety):** it is a PLANNING WORKSHEET, not a payment. It records intent (which bills,
    how much) as a LOCAL overlay and generates the pay list; it NEVER pays QBO or moves money. QBO
    stays source of truth - the owner records the real payment there and the bill clears here on the
    next `sync-ap`. So, unlike the lien tag, pay marks are NOT mirrored to the workbook.
  - **Store:** `shared/bill_marks.py` gains a SEPARATE `pay_mark(bill_id, amount, updated_at)` table +
    `read_pay_marks` / `set_pay_marks` (one txn: selected→upsert, unselected→delete) / `clear_pay_marks`.
    Keyed by the QBO bill id, absent-safe, in the ledger DB. `amount` NULL = pay the full open balance;
    a number = a partial. Only the dashboard reads/writes it (excel_bill_sync never touches it).
  - **Backend:** `_fetch_ap` attaches `pay_selected` / `pay_amount` to each bill; `POST /api/pay-run`
    (`_save_pay_run`, batch) and `POST /api/pay-run/clear` (`_clear_pay_run`).
  - **Frontend (`renderPayBills`):** mark a bill (checkbox), edit its **Pay $** (defaults to the full
    open balance, editable down to a partial - amber when it is not the full amount). Filters: search,
    division, Show (Approved & open · All open · In this pay run), GC-funded only, Select all shown,
    Clear run. Columns carry every metric the owner asked for: Vendor · Client · Project # · Bill # ·
    Date · Open bal · **Pay $** · GC draw (AR invoice) status · Invoice # (QBO deep link) · Lien.
    A sticky **Save pay run** bar (mirrors the lien save bar) shows the live count + total; Save
    persists and **generates the Pay list** grouped by vendor with subtotals + a GRAND TOTAL, plus a
    **CSV export** to take to QBO / the bank. State model: `paySaved` recomputed from BILLS every
    render (server truth) + `payDraft` overlay for unsaved edits; auto-refresh + unload are guarded
    while a draft is unsaved. Verified live on :8791: 2891 bills / 649 open / 397 approved+open,
    select→save→**reload persists** (partial `$100` kept, dirty=0), clear empties the run, CSV quotes
    commas correctly, no console errors.
- **Systems tab — the process registry, live in the ledger (2026-08-19).** The systems & process
  registry (`AI Brain_Vault/02_processes/`, eight domain files) now renders as a tab instead of a
  daily markdown digest. Requested by the owner: "we just need to have this in the Project Ledger,
  my systems and processes live view."
  - **`registry_view.py`** — parses the eight domain markdown row tables
    (`ID | Process | Owner | Operators & touchers | Record | Automation | Cadence | H | State | Life`).
    Only tables whose header's first cell is `ID` are treated as registry tables (the files carry
    other pipe tables). Strips markdown, resolves the three axes into `health_key` / `state_kind` +
    `confirmed_on` / `life_key`, and flags retired rows (struck ID, or State/Life `retired`).
    The `Life` column is absent from older domain files — a row without one is **live**.
    Standalone self-check: `python3 ledger/registry_view.py`.
  - **`GET /api/processes`** — re-parses on EVERY request. No cache, no ledger table, no write-back:
    the vault stays the source of truth and this is only a window onto it. Read-only, per the scope
    the owner picked.
  - **The tab** — six KPIs (processes · broken · fragile · running clean · unconfirmed · agreed-but-
    not-live), domain chips, and filters for search / owner / health / state / life / show-retired,
    over one table grouped by domain. Health dot is the only saturated colour on a row; the life tag
    shows only when a row is NOT live, because "agreed but never built" is what the registry exists
    to catch.
  - **Path** — `shared/paths.vault_dir()` / `process_registry_dir()`, override `ACB_VAULT_DIR`.
    READ-ONLY from this repo. No name ever enters the UI: owners are the role handles as stored, and
    the roster is not read.
  - **Verified live in the browser** (dark, 1223px pane): 81 rows / 8 domains parsed, 73 active +
    8 retired; every filter and chip re-counts correctly; no console errors; table fits with no
    horizontal scroll. **Liveness proven** against a scratch copy of the registry — edited a row's
    health and state in the markdown and both the row and the rollup counts moved on the next fetch
    with no restart. **Missing vault** degrades to a one-line message on the tab while `/api/data`
    and every other tab keep working.
  - Dashboard build **v1.1.0**.
  - **The 06:38 daily digest scheduled task is disabled** (not deleted); `02_processes/digest-log.md`
    is kept as history.

- **Graph tab — the org as a map (2026-08-25).** Requested by the owner: "add the graph view of my
  org the same way that we have it for obsidian, any other graphs you have made to help explain our
  systems don't create just import and finalize in ledger." One self-contained canvas viewer (no
  libraries — the dashboard has zero external scripts and stays that way) renders two things:
  - **Org map** — every vault note is a node, every `[[wikilink]]` an edge (live, ~80 notes /
    ~388 links). Force-directed like Obsidian's graph: nodes coloured by top folder
    (hub / company / processes / systems / integrations / tools / tasks), sized by link count.
    Default framing centres on the MEDIAN node at the 85th-percentile radius so a few flung-out
    nodes can't shrink the core to dots; **Fit** frames every node exactly. Hover highlights a
    node + its neighbours (rest dim); click opens a side panel (group, link count, neighbour list);
    search dims non-matches and centres the first hit; drag a node, scroll to zoom, drag to pan.
  - **System diagrams** — the mermaid flowcharts already authored in `docs/ARCHITECTURE.md`
    (AR · AP · WIP · Ledger · exports · money-bleeds) are IMPORTED (not redrawn): the exact nodes
    and arrows are parsed from the mermaid source and laid out top-down/left-right (longest-path
    layers + barycentre crossing-reduction) with directional arrowheads and edge labels.
  - **`vault_graph.py`** — walks the vault (`rglob('*.md')`), resolves wikilinks by path-id then
    bare stem, and parses the mermaid blocks. **`ROSTER.md` is excluded explicitly** (plus
    dotfolders / `~$` temp) so no name can reach the graph — the vault is role-handle-only and the
    roster is the one names file. Standalone self-check: `python3 ledger/vault_graph.py`.
  - **`GET /api/graph`** — re-parses the vault + `ARCHITECTURE.md` on EVERY request. No cache, no
    ledger table, no write-back; a missing vault degrades to a one-line message on the tab.
  - **Path** — `shared/paths.vault_dir()`, override `ACB_VAULT_DIR`, READ-ONLY. Diagrams read from
    the repo's own `docs/ARCHITECTURE.md`.
  - **Render loop** idles when nothing moves (redraws only on change or while the sim settles) and
    stops entirely when you leave the tab. Colours are theme tokens (`--graph-*`, light + dark).
  - **Verified live in the browser** (dark + light, 1221px pane): org map 80 nodes / 388 links and
    all 6 diagrams render; deterministic parse (CLI == server); hover / click / search / mode-switch
    work; no console errors; ROSTER absent from the node set (80 notes, not 81).

- **WIP Review tab — the pending WIP update as accept/merge (2026-08-25).** Requested by the
  owner: "make the wip update give me the report to accept/merge ... i can accept costs/billed to
  date and have pm answer on the rest ... easy to see the before and afters ... showing me the
  changes and once i approve/disapprove each change it then syncs to the RP Test for RP, CP Test
  for CP and the Master Test for all three divisions." A **Financials → WIP Review** tab shows the
  pending update as a per-job, per-field **WAS → NOW** diff and writes only what the owner approves
  to the three Test tabs.
  - **The split (the whole point).** **Accept · QuickBooks** = Costs / Billed / Retainage - facts,
    **checked by default**. **PM answers** = Original Contract / Approved COs / Original ETC / CO
    Costs - **unchecked** until confirmed. Each changed field has its own approve/disapprove
    checkbox; approve writes the new value, leave it off to keep the current tab value. Bulk:
    Approve all QBO · Approve all shown · Clear all, plus per-job Approve. Added jobs get an
    include/exclude toggle; removed jobs are shown (they drop on the next write).
  - **`wip/wip_review_common.py`** — the ONE place the reviewable field set, the tab-header lookup
    (working tabs carry the ORIGINAL CONTRACT / APPROVED COs breakout; the lean Test-Master folds
    them into TOTAL CONTRACT PRICE / ESTIMATED TOTAL COSTS), the diff, and the revert live, so the
    three tabs can't diff or merge differently. A disapproved field carries its **`revert`** (the
    exact "was" the owner saw) so every tab reverts the SAME source number - reverting the source
    fields (base_contract, co_revenue) also fixes Test-Master's derived totals for free.
  - **Faithful reuse, no drift.** Each tool keeps its exact compute+write; two thin modes were
    added: `--emit-review <json>` (compute as usual, diff the tab, dump JSON, **no write**) and
    `--apply-review <json>` (compute, revert disapproved fields, then write normally). CP →
    `cp_wip_reader` (Test - CP, active+completed), RP → `rp_wip_reader` (Test - RP, the owner's RP
    file), MFD → `master_wip_test` (Test-Master; a **fast early path** emits MFD only, skipping the
    CP/RP scan). Master's `--apply-review` writes Test-Master for all three divisions and is guarded
    to NOT double-write Test - RP (rp_wip_reader owns that write, decisions applied).
  - **Dashboard** talks to the tools ONLY by subprocess + JSON (repo rule): `POST /api/wip/review`
    runs the three emits (gated, QBO Touch ID each), `GET /api/wip/review` merges the three JSONs,
    `POST /api/wip/merge` saves `decisions.json` and runs the three guarded writes. JSON lives in
    `~/Library/Application Support/Proficient/wip-review/` (outside the repo). The single sync-lock
    is shared with the Console so a review/merge can't overlap a sync (concurrent QBO + tab writes
    would corrupt).
  - **Verified:** CP emit (51 jobs / 34 changed) and RP emit (119 / 116) from real runs; MFD
    snapshot mapping (folded columns) correct; merged endpoint (173 jobs, 3 divisions); UI before/
    after + QBO/PM split + toggles + bulk + filters, light + dark, no console errors; decisions
    payload carries `revert` for disapproved; `apply_decisions` unit-verified (revert to carried
    value, keep approved fresh, drop rejected adds); CP `--apply-review --dry-run` ran clean.
  - **Known v1 gaps (fast-follow):** review + merge each re-pull QBO (≈3 Touch IDs apiece; a shared
    cache would halve it); approved QBO fields write the LATEST pull (may differ by pennies from the
    review if QBO moved between); a REMOVED job can't yet be KEPT from the UI (it drops).

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
- **"Needs attention" widget** (generic, data-driven exposure rules — not stored findings):
  Underbilled / Overbilled / Over budget (costs>ETC) / Borrowing cash. Each chip shows a count +
  subtotal and click-filters the projects table (composes with the dropdown filters); "Clear filter"
  resets it. Toggleable like the other widgets. Table cues: % complete as an inline bar (red past
  100%), pure-job-borrow in red, underbillings in green. Verified live (Underbilled → 10 rows).
  - NOTE: the app-preview sandbox can't run this server (it needs the DB + `shared/` outside
    `.preview`); run it directly with `python3 ledger/dashboard.py`. A `.claude/launch.json` entry
    (`ledger-dashboard`) exists but launch.json is untracked/local.

- **`ap_bill_line` + `load_bill_tracker.py`** — AP + lien feed from `Bill Tracker.xlsx`
  (Bills + Inventory display sheets → 2,814 lines, $5.0M open AP, 448 on the lien clock).
  - Read-only on Excel; full-replace by `source='bill_tracker'` (idempotent, mirrors the file).
  - `v_ap_by_project` view; dashboard **AP & liens** widget (open-AP stats + lien watchlist ordered
    by urgency, red past-due pills) + AP line in each job's detail. Verified live.
  - **Deliberately NOT cost_line:** Bill Tracker excludes subs (measured 25–98% short of WIP cost
    per job), so it can't state job cost. Job cost stays in wip_snapshot; complete cost_line waits
    for the qbo-export pull. 284 AP project#s are off-WIP (closed/older) — kept, no FK on project_no.

- **`shared/qbo_costs.py` + `load_costs.py`** — complete cost load, by cost code, incl. subs.
  - `cost_leaf` MOVED out of project-pnl into `shared/qbo_costs.py` (project-pnl imports it back —
    byte-compatible, compiles clean). The ONE resolver both tools share, so they can't drift.
    Engine adds `is_cost_code`, `cost_code_meta`, `build_account_map`, `pull_expense_txns`,
    `cost_lines_from_txns` (network-free, unit-testable), `iter_cost_lines`.
  - `schema.sql`: `cost_line` fleshed out (txn_type, account, vendor, description, source, loaded_at)
    + `v_cost_by_project` / `v_cost_by_code` views.
  - `load_costs.py`: pulls QBO Bills + Purchases → `cost_line` keyed by cost code; scoped full-replace
    (idempotent); `--active/--division/--project/--since/--dry-run`; **`--selftest` proves the whole
    pipeline OFFLINE** (fabricated txns → codes resolved → cost_line written → reconciles $25k=$25k).
    Reconciles loaded cost vs `wip_snapshot.costs_to_date` per project after each load.
  - CLAUDE.md updated (cost_leaf now in shared/qbo_costs; ledger subsystem bullet added).
  - **Run against live QBO 2026-08-08** (owner, `--active`): cost_line populated across active
    projects; **90 of 96 active projects with a WIP cost reconcile within 5%** of
    wip_snapshot.costs_to_date. Residual = per-job attribution differences (a handful of RP/CP
    jobs where QBO-sourced cost ≠ the WIP figure) — surfaced to the owner for review; specific
    dollar findings stay OUT of the repo (scope rule).

- **`customer` + `sales_touch` + `load_customers.py`** — the CRM / sales-pipeline fold (2026-08-09).
  - `schema.sql`: `customer` (one row per Notion Customer List page — identity + pipeline stage +
    Notion `Created by` / `Last edited by` = who sourced / who worked it last, the honest per-rep
    attribution, NO manual Owner property) and `sales_touch` (one row per "History of interactions"
    body line, touch_date parsed) + `v_sales_pipeline` / `v_sales_by_rep` views. Portable (no
    SQLite-only date funcs in views).
  - `shared/notion_client.py`: added `block_children()` (page-body reader) — the shared client now
    reads bodies, not just query/create/update.
  - `load_customers.py`: reads the Customer List **read-only** via `shared.notion_client`; parses
    props + the touch log; idempotent full-replace by `source='notion_customer_list'`;
    `--dry-run/--show/--limit/--all-notes`; **`--selftest` proves parse+load OFFLINE** (no Notion).
    Body fetched only for worked rows (status past Lead/Follow up) by default.
  - `machine.env`: `ACB_CUSTOMER_LIST_DS_ID` added (local, gitignored). Auth reuses the shared
    Notion token (Keychain `proficient-automation-worker/notion`) — verified it can read the list.
  - **Run live 2026-08-09:** 622 customers + 168 touches landed; spine untouched (project 170,
    wip 170, ap 2814, cost 6009). `v_sales_by_rep` for the outreach rep = 141 worked / 112 contacted
    / 12 interested; touch-log dates parse (e.g. "Quote sent 07/15/26" → 2026-07-15).
  - `docs/ARCHITECTURE.md` + README + CLAUDE.md ledger bullet updated in the same commit.
  - Not joined to `project` yet (leads→jobs downstream).

- **Dashboard Sales tab** (2026-08-09) — `/api/data` now carries a `sales` section (`_fetch_sales`):
  pipeline funnel, activity-by-rep (last-editor attribution; the invoice-sync bot relabeled
  "Automation (sync)" via `_rep_label`), warm-account cards with each account's full touch log +
  a stale flag (>21d), and a searchable/filterable all-customers table linking out to Notion. New
  **Sales** tab in `index.html` + `renderSales()` in `app.js` (reuses the existing kpi/bar/table
  helpers, no new deps) + warm-card CSS. **Verified live** against the loaded DB (622 customers /
  168 touches): pipeline, rep table, all 30 warm accounts with touch logs, and the customer table
  all render in the browser (dark + light). Read-only — the tab never writes; edits stay in Notion.
  - QoL (2026-08-09): **Set as default** button + **boxed default** — `baseDefaults()` baseline the
    user snapshots via "Set as default"; Reset + fresh browsers restore to it (shipped default width
    now `boxed`). **Automation accounts kept out of the sales scoreboard** — Notion bots (bare UUID)
    and any name in `ACB_SALES_AUTOMATION_REPS` (machine.env, gitignored — no names in the repo) are
    excluded from Activity-by-rep and shown as "Automation" in the customer table (fixes the raw-UUID
    + the import-account-as-rep). Clickable KPI tiles + pipeline rows → filter the customer table;
    client names are Notion links. Verified live: reps = 5 real people, Interested tile → 30 shown.
- **Dashboard cost-code drill** — `/api/data` now carries a `cost` section (`_fetch_costs`):
  portfolio by-code, per-project by-code, and per-project rollup attached to each project row.
  New **"Costs by code"** widget (portfolio table with % bars), a **QBO Costs / Subs** toggleable
  column pair, and a **"Costs (QBO, by code)"** group in each job's detail showing total loaded /
  subs / WIP costs_to_date (the reconciliation) + the full code breakdown. Verified live — the
  biggest MFD job loads within ~0.3% of its WIP costs_to_date; portfolio spans ~80 cost codes.
- **Costs by code grouped as cost TYPE (parent) → job TYPE (sub)** — the JobTread model: the number
  meaning (Concrete/Labor/Rebar…) is the parent that ALL material rolls up to, the prefix
  (Slab/Paving/Flatwork…) is the collapsible sub, with the cost code shown. Account-based lines land
  under their cost-type parent as an "(account)" sub. Grouping computed server-side in `_fetch_costs`
  via `shared/qbo_costs.job_type_name` + `cost_code_meta`. Verified live. **This mirrors an intended
  QBO restructure** (today QBO cost codes are standalone items routing to categories) — that future
  change is an owner/ops decision, tracked in the vault, not here.
- **Layout: metrics up, bills down (owner).** Widget order is now KPIs → attention → costs → margins
  → division → **AP & liens (bills) moved to position 6, off the top** → projects. The Costs widget
  leads with a **cost-mix** proportional bar + legend (how much each cost type takes, % wise —
  Concrete/Labor/Rebar… as one glance) above the grouped tree.
- **Multi-tab app (owner: "deliberate tabs only a page would contain").** A tab bar splits the app
  into **Overview** (the glance: KPIs · attention · cost-mix · margins · division · projects) plus
  three deep pages: **Costs** (the cost-type→job-type tree + a **code→jobs pivot** — click a code,
  see every job that spent on it), **Liens** (the full collections worklist — see the redesign below),
  and **Vendors** (spend by vendor, jobs, of-which-subs — from `cost_line.vendor`; new `by_vendor`
  in the API).
  Active tab persists in localStorage; bills now live only on the Liens tab.
- **Markup + margin (owner).** Derived per job: **planned markup** (contract÷ETC, on cost),
  **planned margin** (GP÷contract, on revenue), **actual markup** (billed÷QBO cost) — as toggleable
  columns, in the Margins widget stats, and in each job's detail. Markup and margin are kept
  distinct on purpose (markup on cost ≠ margin on revenue).
- **Margins & burn** — derived from the loaded QBO costs (client-side; the WIP only ever showed
  *billed* margin, never actual-cost margin). Budget burn (cost ÷ ETC), margin-to-date (billed −
  cost) + margin %, and subs-share as toggleable columns; a portfolio "Margins & burn" widget
  (stats + an over-budget watchlist ordered by burn, worst first); and a "Margin (QBO actual)"
  group in each job's detail. Verified live.
- **Budget-adherence rule (`isOverBudget`)** — the ONE thing the dashboard flags "over budget" with,
  encoding the owner/ops-manager ruling: **flatwork (`-FTW`) budgets are a soft reference, not a
  strict target like slab** (flatwork = sub sent, charged by labor). A `-FTW` job is flagged over
  budget only when its size (max contract/ETC) ≥ ~$15k; slab / CP / MFD stay strict. Used by the
  "Over budget" attention chip and the Margins over-budget watchlist — cut the over-budget list from
  the raw burn>1 count down to the genuinely-actionable jobs (small flatwork false-alarms removed,
  big flatwork + slab + CP kept). Threshold `FTW_BUDGET_FLOOR` is a documented judgment knob. Apply
  the same tolerance when `budget_line` (budget-vs-actual by code) is built.

- **`open_ledger.command` launcher + CompanyHealth co-location (owner).** Double-click launcher
  (self-locating) that starts the server if down, then opens the browser — one-click, no terminal.
  A thin copy lives in the owner-private `~/Documents/CompanyHealth/` ("Open Project Ledger.command")
  so the ledger opens from the same private cockpit as Company Dashboard + Tracker. Tool code stays
  in the repo; only the front-door launcher sits in CompanyHealth.

- **Draws tab — the draw race-through + the waiver input (owner).** `load_bill_tracker.py` now also
  captures `matched_invoice` / `invoice_status` / `gc_paid_date` / `pay_date` / `bt_key` (ap_bill_line
  migrated in place — rows reload from Excel, lossless). `dashboard.py _fetch_draws` rolls bills up BY
  DRAW (dedup lines → bills), computes the stage — **Awaiting GC funding → Fund in, pay vendors →
  Ready to turn in** (green once every bill is paid; see the later entry — the collect-waivers stage
  was removed) — sorted worklist, 40 most-recent shown of 337. The
  **Draws** tab renders each draw with its bills and a **waiver checkbox** per bill.
- **The ledger's FIRST write surface: `waiver` table + `POST /api/waiver`.** The one place the app
  writes — the owner marks "unconditional waiver in hand." Everything else stays read-only (`mode=ro`);
  the waiver write opens a scoped writable connection, binds 127.0.0.1 only, keyed by
  hash(matched_invoice+vendor+bill) so it survives Bill Tracker reloads. Verified end-to-end: UI
  checkbox → POST → DB persisted (received + timestamp), reverts on failure.

- **"My view" home tab (default) + big-picture→zoom UX pass (owner).**
  - **My view** (default landing): a **data-freshness** strip (sync-ap / sync-ar / WIP-master file
    mtimes + ledger loaded_at, via `_freshness`), clickable **action items** (liens past-due / due-≤7d,
    draws collect-waivers / ready, over-budget, underbilled → each jumps to the right tab+filter), and
    **Working on** = active projects with a division filter. (sync-ar mtime shows "not found" until the
    Open_Invoices path is confirmed.)
  - **Collapse everything by default** — cost tree parents + draw cards start collapsed; expand to zoom.
    Preserved across the live auto-refresh.
  - **Search** on Draws (draw/project/vendor) and Vendors; **division filter** on Draws (MFD/CP only).
  - **Division drill**: click a By-division rollup row → that division's active projects.
  - **Live**: soft auto-refresh every 90s (preserves expand state + active tab).
  - **Vendor TYPE** column (replaces the subs-$ column): "Sub" vs "Supplier: <material>" (Concrete /
    Rebar / …), derived from each vendor's cost mix.
  - **RP is not draws** — the Draws rollup excludes RP (residential bills at completion/milestones, not
    formal draws): 337 → 49 draws (MFD/CP only). Owner ruling.

- **Notion link-out — action items → Notion pages (owner: "folder memory like Notion").**
  MVP wired end-to-end for **draws ready to turn in**. `shared/notion_client.py` (the clean
  invoice-sync client, graduated to shared — invoice-sync keeps its tool-local copy on purpose);
  `ledger/sync_actions.py` finds ready draws (funded + all paid + all waivers) and upserts a page in
  the Notion **"Ledger Actions"** DB keyed by a stable Action Key, reads Status back → the local
  `action` table (new). The dashboard shows a **📄 Notion · <status>** link on tracked draws; the
  ledger stays the RADAR, the thread/notes/done live in the Notion page. `--dry-run` proves it
  offline. `ACB_ACTIONS_DS_ID` in machine.env (gitignored). Ledger stays read-only; Notion writes
  are scoped to the Actions DB.
  - **ONE manual step for the automated path:** share the "Ledger Actions" DB with the Notion
    integration ("Automation Integrator") — open the DB → ••• → Connections → add it — so
    `sync_actions.py` (keychain token) can write unattended. Until then the 404 is expected.
  - Proof page created via the connected workspace (CP585 Draw #4). **Demo data to clean up:** the
    CP585 draw's 2 waivers were test-marked to make it "ready"; uncheck them (and delete/close the
    demo Notion page) — they are not real.

- **Liens tab redesign — clickable stage tiles → one filtered table (owner).** Replaced the stacked
  per-status bucket sections with **clickable stage widgets** ("All on the clock" + Past-due / ≤7d /
  ≤15d / ≤30d / Notice-sent / Lien-filed, each count + $ open, urgency-coloured edge) that filter a
  **single table below** (the `.attn`-tile pattern, mirroring Overview's "Needs attention"). Columns
  reordered to the owner's spec — **CP # · Draw # · Name/Address · Invoice # · Amount** — with Vendor
  trailing and urgency shown as the row's left edge so **CP # stays first**. Invoice # (the vendor
  bill_ref) is a mono chip and Amount is bold — the two the owner said "get lost". Added a search box
  (CP #/draw/name/vendor/invoice). Backend: `_fetch_ap` now also selects `matched_invoice` +
  `invoice_no`; front-end derives Draw # from `invoice_no` (falls back to parsing the draw label) and
  Name from the WIP name (falls back to the draw label). Row-click still opens the job detail. No new
  scripts / data-flow (same `ap_bill_line` → dashboard), so ARCHITECTURE.md unchanged. Verified live:
  header order, tile filtering (Past-due → 110 rows all past-due), search (7 SUNRISE rows), row-click
  detail, no console errors.
- **P&L link — job detail ↔ project-pnl (owner picked A+B: "open + generate, show when last pulled").**
  New `shared/pnl_paths.py` resolves a project's `Project_PnL_<proj>.xlsx` with the SAME rules
  project-pnl writes with (CP → Common-drive awarded folder's `Profit and Loss/`, else OneDrive
  `PROJECT P&Ls/<proj>/`) and returns `{exists, path, mtime, note}` — `mtime` = the **"last pulled"**
  time. Dashboard endpoints (all guarded by `_PROJ_RE`): **`GET /api/pnl`** (find + mtime),
  **`POST /api/pnl/open`** (macOS `open` on the resolved workbook — only ever `Project_PnL_<proj>.xlsx`),
  **`POST /api/pnl/generate`** (runs `project-pnl/run_pnl.sh <proj>` as a **subprocess**, not an import —
  gated behind a `confirm` flag; logs to `~/Library/Logs/Proficient/ledger-pnl/`), **`GET /api/pnl/status`**
  (running/done/error + elapsed; a daemon thread reaps the process). Job detail shows a **P&L
  (project-pnl)** group: *Last pulled <ago · timestamp>* + **Open** + **Generate / Refresh** (confirm
  dialog warns about the QBO pull + Touch ID). Generate is the ONE place the dashboard triggers a QBO
  pull + a file write — QBO stays read-only inside project-pnl; the .xlsx write is the gated action.
  Verified live: find (MFD325 → mtime; CP745 → "not generated yet" + CP note), the `confirm`-required
  gate (400 without it), status idle, and the panel rendering both states. **Generate's first LIVE run
  is owner-driven** (Touch ID on the Mac) — the non-QBO plumbing is verified; the QBO run was not
  triggered from here.
  - **Follow-up — dedupe the resolver:** project-pnl still has its own
    `_resolve_project_out_dir`/`_find_awarded_cp_folder`; it should import `shared/pnl_paths.py`
    (same move as `cost_leaf`→`shared/qbo_costs.py`). Deferred to avoid editing the 326 KB export
    script while a concurrent session is in `shared/`.
  - **Not built — option C** (project-pnl reads cost_line from the ledger instead of re-pulling QBO,
    the "own-the-spine" data-source refactor). Still the strategic direction; larger, separate job.

- **Dock on/off switch — `Project Ledger.app` (owner: on-demand, "no always-on", real indicator + clean Quit).**
  The owner rejected an always-on launchd agent and wanted a Dock indicator + one-click open + a clean
  off switch. **`ledger/app/`** holds a real Cocoa app (`ledger_app.py` via **PyObjC + py2app**);
  `build_ledger_app.command` installs the toolkit once, optionally makes an icon from `app_icon.png`,
  and py2app-builds (alias mode) → **`~/Applications/Project Ledger.app`**. The app runs the SAME
  `dashboard.py` server as a **child process**; its **Dock icon present = ON, gone = OFF** (the
  indicator); **Cmd-Q / Quit / log out / shut down / real system sleep** all stop it cleanly
  (`applicationShouldTerminate` + `NSWorkspaceWillSleepNotification` → `pkill -f ledger/dashboard.py`).
  Repo path is baked via the plist `LSEnvironment` (works wherever the .app lives); the child is
  spawned with a **cleaned env** (strip py2app's `PYTHONPATH`/`PYTHONHOME` so the server sees its own
  site-packages, e.g. `requests`). Never runs at login.
  - **Why not the simpler routes (learned the hard way):** an `osacompile` applet **can't stay open**
    from the CLI (it quit instantly and killed the server); a hand-built foreground app **can't Quit
    cleanly** (⌘Q hangs). A proper PyObjC/py2app Cocoa app is the only one that does **both**. Isolation-
    tested each claim before building.
  - **Server hardening (used by the app + the terminal launcher):** `dashboard.py --background`
    double-forks + `setsid` (detaches so a launcher can't reap it; PPID→launchd), and it now quits
    cleanly on **SIGTERM** (`signal.default_int_handler`). `open_ledger.command` uses `--background`.
  - Verified live end-to-end: launch → server 200 + Dock name "Project Ledger" + browser opens; Quit →
    app gone, server 000, child gone. Sleep-stop uses the documented sleep notification (not triggered
    from here). Build artifacts (`ledger/app/{build,dist}`, `app_icon.icns`) are gitignored.
- **"Recommended to sync" freshness flag (owner, weekend-aware).** The My-view data-freshness cards now
  flag a source with a **⟳ Sync recommended** badge (+ amber border, + a "N recommended to sync" note)
  when it is stale **> 48 business-hours** — weekends don't age the data (a Friday load isn't "stale"
  Monday). New client-side `businessHoursSince()` sums only Mon–Fri slices; threshold `STALE_BUSINESS_H
  = 48`. Purely front-end (uses the existing `meta.freshness`); no server change. Verified: unit tests
  (Fri→Mon = 16 business-h vs 64 raw → not flagged; Wed→Fri = 52 → flagged) + the badge/note render.

- **Draws tab — clickable stage tiles + unambiguous "who paid whom" wording (owner).** The 3 stat
  tiles (Ready to turn in / Collect waivers / Pay vendors) are now **clickable → filter the draw list**
  by stage (toggle; "Show all" clears; active tile outlined), mirroring the Liens tab. The owner asked
  what "Paid" meant — so pill text is now direction-explicit via a display map `DRAW_STAGE_LABEL`
  ("Fund in — pay vendors" → **"GC funded → pay vendors"**, "Paid — collect waivers" → **"Vendors paid →
  collect waivers"**), tile subs spell it out ("GC funded — vendors not paid yet" / "vendors paid —
  waivers pending"), and the hint states the one-way flow (GC funds you IN → you pay vendors OUT →
  waivers back). **Internal stage keys are unchanged** (matched in `_STAGE_ORDER`, `DRAW_STAGE_CLASS`,
  `renderHome`, the waiver-toggle recompute) — display-only. Front-end only; verified live (Pay vendors
  → 49→15, active highlight, clearer pills, no console errors). New `kpi-click` CSS.

- **Owner's date format everywhere + Refresh feedback + a judgeable % bar (owner QoL).**
  - **Dates NEVER year-first** (binding owner preference): new `fmtDate(v, withTime)` renders every
    displayed date as **weekday, abbr-month day, year** — "Mon, Aug 10, 2026" (+ 12h time) — applied to
    the meta line (report/loaded), Data-freshness cards, P&L "last pulled", detail report_date, and the
    draw pay/GC-funded cells. Parses ISO components into a LOCAL date (no `new Date("YYYY-MM-DD")` UTC
    off-by-one). ISO stays only in filenames/keys. (Memory: dates-never-year-first.)
  - **Refresh** now gives feedback (button → "Refreshing…", then a toast "Refreshed · ledger loaded
    <date>") + honest tooltip — it re-reads the ledger instantly; it does NOT re-pull QBO (that's a
    sync). Fixed the wiring (onclick had passed the MouseEvent as `isAuto`).
  - **`% complete` bar** is now full-width (`.pct-bar`, min 130px) with a **visible track**, so the fill
    level is judgeable at a glance; over-100% caps full + red (163.1% reads clearly). Verified live.

- **AR money-IN per draw + the Draws-tab table redesign (owner: "connect systems, don't re-pull QBO").**
  Owner wanted the Draws view to show what the GC pays HIM (in) next to what he pays vendors (out) — and
  corrected an initial QBO-pull approach: the **Invoice Tracker** (Notion, `invoice-sync`) already mirrors
  every QBO invoice and keeps paid ones 12 months, so the ledger reads THAT, not QBO again.
  `QBO → invoice-sync → Invoice Tracker (Notion) → ledger.billing_event`.
  - **`load_invoices.py`** reads both Invoice Tracker DBs (Res/Com `265b…`, MFD `0f8e…`) via
    `shared/notion_client` (the shared token — **no QBO, no Touch ID**) → `billing_event`, keyed by
    **Invoice #** = `ap_bill_line.invoice_no`, so the draw↔invoice join is exact. `billing_event` schema
    fleshed out (doc_number, project, division, customer, memo, amount=TotalAmt, balance, status
    Unpaid/Partially Paid/Paid, source, loaded_at); empty-table migration = safe drop+recreate. Read-only
    on Notion; full-replace by `source='invoice_tracker'`; `--dry-run` coverage; `--selftest` proves the
    parse + status + draw↔invoice join OFFLINE. **Run live: 310 invoices → 41 of 49 draws matched;
    $14.7M billed · $4.5M still open.** Also fixed the freshness sync-ar path (`Collections/Open_Invoices.xlsx`).
  - **`_fetch_draws`** now joins `billing_event` by Invoice # → each draw carries `billed` (net in),
    `ar_status`, `ar_open`, `ar_date`, `customer`.
  - **Draws tab is now a TABLE** (`renderDraws` rewrite): one row per draw — **caret · Project # · Draw
    memo · Billed (in, green) · Invoice # · Date · Paid out (+paid/N) · Stage** — **green row when fully
    done** (Ready to turn in); **click a row → its bills open underneath** (vendor · bill # · amount ·
    paid · GC-funded · waiver checkbox) with a caption. `drawsExpanded` state; waiver POST preserved.
    Verified live: 49 rows, 41 with billed-in, expand works, dates month-first, no console errors.

- **Draws/Liens UI polish batch (owner).** Draws **grouped by project #** (project header row with
  name + N draws + $in/$out totals; dropped the redundant per-row Project # column; memo strips the
  project #). Tables **fit the boxed width** (default is boxed — owner dislikes full width): Name/memo
  columns truncate with ellipsis, and the draw **stage pill shows the short action** ("Pay vendors" /
  "Collect waivers", full "GC funded → …" on hover) so the Stage column stops clipping. **Per-field
  search boxes** on Draws (Project # · Vendor · Invoice # · Division) and Liens (CP # · Vendor · Invoice
  # · Name/address) — a row must match every filled field (AND); Vendors/Sales stay single. **QBO deep
  links on invoices:** the draws' Invoice # links to `app.qbo.intuit.com/app/invoice?txnId=<id>` (the
  Invoice Id comes from the Invoice Tracker load, on `billing_event`). Verified live.
  - **QBO deep links on BILLS (AP) — DONE.** The earlier "no bill id" assumption was wrong: the Bill
    Tracker's **"Open" column is `=HYPERLINK("…/app/bill?txnId=<id>","↗")`** for every bill. `data_only`
    reads only the cached "↗" glyph, so `load_bill_tracker` now does a **second read-only, formula pass**
    (`read_bill_links`, scanning each row for `app/bill?txnId=`) and stores the link in the new
    `ap_bill_line.qbo_link` column (migration drops+recreates on the missing column; matched **2825/2825**).
    `_fetch_ap` and `_fetch_draws` carry `qbo_link`; the Liens **Invoice #** chip (448/448) and the draw
    **Bill #** cell are now `app/bill` links (new-tab, click-stops-row-expand). A shared `qboLinkCell()`
    helper + `a.invno.qbo-link` hover style. Verified live on both tabs, no console errors.
  - **Draw "done" is now PAID, not paid+waivers (owner).** The `waiver` table is empty (the owner
    doesn't mark waivers in-tool), so the old `waivers == n` gate made **"Ready to turn in" (green)
    unreachable** — every fully-paid draw sat on amber **"Paid — collect waivers"** forever. Per the
    owner ("don't say collect waiver if all bills have been paid, just make it green"), the stage is
    now **`paid == n` → green "Ready to turn in"** and the "Paid — collect waivers" stage is **removed**
    (backend stage logic + `_STAGE_ORDER`; frontend `DRAW_STAGE_CLASS/LABEL/SHORT`, the 3 stat tiles
    now **Ready to turn in · Pay vendors · Awaiting GC**, the My-view action item, and the explainer).
    Per-bill **waiver checkboxes stay** (they persist to `waiver` and update the expanded caption for
    the owner's records) but **no longer gate the color**. Live: 18 ready / 15 pay / 16 awaiting = 49.
  - **NOTE — money-IN blank on some draws is a source gap, not a bug.** Billed-(in) comes from the
    Invoice Tracker (`billing_event`, by Invoice #); a draw whose AR invoice isn't in the tracker shows
    "—" (e.g. Briarwood invoices 33942/34103 are absent from the tracker though older ones are present —
    a data-entry gap in the Invoice Tracker, not a recency cutoff). No QBO fallback by design (owner).
  - **UI fit + readability polish (owner).** (a) The draws STAGE pill was clipping off the right edge —
    the short label was reverted to **"Ready to turn in"** (the verbose "Paid — ready to turn in" pushed
    the column out of view; the full text stays on hover). (b) The Sales **All-customers** table clipped
    its last column (Touches) because long client names blew out col 1 — the Client name now truncates
    (`#salesTable` col 1 `max-width:240px`, ellipsis), so all six columns fit. (c) Pipeline **bars were
    near-invisible** (faint fill, sub-pixel for small counts) — added a **border** to `.bar .bar-fill`
    (+`.over`) and a **7px min-width** on the funnel fill so counts like 1/2/9 show. (d) Sales dates
    (customers table + warm list) now use `fmtDate` — no more ISO `2026-05-29`. Verified live both themes.
  - **Per-rep activity drill on the Sales tab (owner: "Devan report — weekly/daily").** Click a rep in
    "Activity by rep" → a drill (defaults to the busiest-by-touches rep, i.e. the outreach person) with:
    This-week/Last-week/Today/All-time touch tiles (with a vs-last-week trend), a **12-week touch
    timeline** (bordered bars — surfaces the drop-off), a **recent-touch log** (dated notes), **follow-ups
    due**, **going stale (21d+ no contact)**, and their **pipeline by stage**. Backend: `_fetch_sales`
    now sends `touch_log` (every dated touch joined to its customer, rep = `_rep_label(last_edited_by)`)
    and enriches `customers` with `follow_up_date` + `main_status`; all bucketing (Monday-anchored weeks,
    local dates) is client-side in `renderRepActivity()`. Rep is a **runtime value, never hard-coded** —
    the drill works for any rep and the code carries no personal names (names come live from Notion).
    Verified live: auto-features the outreach rep, weekly trend renders, rep-switch works, no console errors.
  - **P&L folded INTO the dashboard — live compute (owner: "super database, one place").** The job
    detail now SHOWS the P&L (was only a link out to the Excel): `_project_pnl(con, proj)` assembles it
    from the spine — **Earned Revenue = contract × %complete** (WIP), **costs from `cost_line`** (QBO
    truth, incl subs, itemized by cost code), **overhead 10% of revenue** (MFD alt = 9% of costs), net
    margin + %; **billed (AR)** shown alongside. Conventions match `project_pnl_export.py` so they
    reconcile. `GET /api/pnl/pl?proj=`; `buildPnlGroup` renders the numbers + a "Costs by code" list
    (uncoded flagged red, subs marked) with the project-pnl **Excel demoted to "Detailed export"**.
    **Cross-platform open (owner's "trick"):** `_os_open()` opens files/folders with the host OS command
    (`open`/`os.startfile`/`xdg-open`) so the same dashboard works on Mac OR Windows; `_pnl_open` gained
    an **Open folder** action (`?folder=1`) — CP resolves onto the Synology Common drive, RP/MFD onto
    OneDrive, per `pnl_paths`. Verified live (CP800 net 2.8%, MFD 9%-on-costs), no console errors.
    See [[ledger-super-database]] memory.
  - **Portfolio "P&L" tab — Phase 2 (owner).** New **P&L** tab (nav after Overview): company totals
    (earned · cost · overhead · net% · billed), a **by-division** table, and a **by-job** table that's
    sortable (click a header) + filterable (project # / division), **worst margin first** so the money
    losers surface (RP6440 −53%, RP6901 −504% — the reconcile-mismatch jobs). `_portfolio_pnl(con)`
    batches the same math as `_project_pnl` into 3 aggregate reads; **active = status Active OR NULL
    (MFD)**, Closed/Complete excluded (137 jobs, company net ~7.5%). `GET /api/pnl/portfolio`, lazy-loaded
    on tab open (`renderPnl`), cache invalidated on reload; rows click → the job detail. Verified live,
    sort/filter/row-click work, no console errors.

  - **Source-folder links — Phase 3 (owner: "put source links").** Each job's detail now has an
    **Open job folder ↗** button (source docs/takeoffs/photos on the file server), cross-platform via
    `_os_open`. New `shared/pnl_paths.job_folder(proj, builder)`: **CP → the Synology awarded folder**
    (matched by #), **RP → the builder folder** under `…/Residential` (matched on `project.builder_or_gc`
    — the exact address folder is inside; full match needs `rp_wip_reader`'s General-List index), **MFD →
    OneDrive** P&L folder fallback (moves a lot). `POST /api/job/open`. Reachable from the P&L tab via the
    row → detail (no per-row clutter). Verified: resolver returns the right paths, button renders.

  - **In-app data sync — Phase 4 (owner: "full-fledged app, no terminal").** A **⟳ Resync** button in
    the Data-freshness panel runs every ledger loader **from the UI** with a live **progress bar** —
    no terminal. Backend: `_run_sync()` (background thread) runs the 5 loaders in order (WIP master →
    QBO costs → Bill Tracker → Invoices → Customers) as subprocesses, recording per-step state;
    `POST /api/sync` (confirm-gated, single-run lock), `GET /api/sync/status` (state · current step ·
    elapsed) for the bar. WIP first (creates `project`); **QBO costs prompts Touch ID** on the Mac.
    Front-end pauses the 90s auto-refresh while syncing (loaders drop/rebuild tables), reloads on done,
    surfaces the failing step + points at `~/Library/Logs/Proficient/ledger-sync`. **"Last ran" is
    plainly visible** — the freshness cards already show each source's timestamp + relative age.
    Verified: state machine (happy + error paths, monkeypatched — no real QBO pull), endpoint gates,
    button/bar render. The real full run is owner-triggered (Touch ID).

  - **QC pass — DONE (the owner's final step).** (a) **Aesthetic (all 8 tabs): pass** — one font,
    color only encodes meaning, new views reuse `.kpi`/`.grid`/pills so they read native
    ([[ledger-ui-aesthetic]]). (b) **P&L reconciliation: pass** — per-project detail = portfolio row =
    company total (verified across CP/MFD/RP incl. the negative-margin jobs). (c) **Correctness review
    (subagent)** found the per-row math clean + no injection/arbitrary-path; **3 real bugs FIXED:**
    (1) `_sync_start` TOCTOU — now **claims `state="running"` inside the lock** so two POSTs can't launch
    overlapping syncs; (2) `pollSync` retried forever on a down server — now **caps at 5 fails** +
    re-enables the button + handles idle-mid-poll (server restart); (3) three different "active"
    definitions — added **`isActive(r)` = status in (Active, blank)** used by the Overview KPI,
    working-list, and active filter, so every "active" count now **agrees at 137 (MFD included)**, matching
    `_portfolio_pnl`. Also a cosmetic empty-company net guard. Verified live, no JS errors.

  - **Draws: AR pay status is its own column + QBO gap-fallback (owner).** (a) The **Status** column
    (Paid green / Unpaid · Partially Paid amber) was split out of the Billed-(in) cell — the amount no
    longer has the status crammed onto it. (b) **QBO gap-fallback** — a handful of older CP/MFD draws
    (8 now; matches "41 of 49 draws matched") had no billed-in because their AR invoice was never entered
    in the Invoice Tracker (NOT swept — the tracker keeps paid invoices). `load_invoices.fill_gaps_from_qbo`
    finds those gaps and pulls ONLY them from QBO by `DocNumber` (`source='qbo_fallback'`), so the tracker
    stays authoritative and QBO fills only its holes (owner's pick over manual entry). Adds **one Touch ID**
    to `load_invoices` (skip with `--no-qbo`); `--dry-run` reports the gap count without pulling; `--selftest`
    stays offline. Verified: mapping (Paid/Partial/Unpaid + project extract), gap-finder = 8, selftest, dry-run.
    The real fill is owner-triggered (Touch ID) via Resync or `load_invoices`.

  - **Excel-style cell selection + running sum (owner: "everyone here uses this").** Click a number
    cell to select it; **drag** for a rectangular range, **Cmd/Ctrl+click** to add cells, **Shift+click**
    to extend. A floating status bar (bottom-right, like Excel) shows **Sum / Count / Avg** with a Copy
    button; Esc or a click-away clears. Number cells only (dates, %, labels, and the trailing "2/7"
    paid-count are skipped); the `$` prefix shows only when every selected cell is money. Works across
    every `.grid` table app-wide, survives re-renders (detached cells drop from the selection), and does
    NOT hijack normal clicks - text cells still open the row. Pure front-end (`initCellSelect` in app.js
    + `.sumbar`/`.cell-sel` CSS). Verified: sum reconciles, drag/Cmd/Shift/Esc/click-away all work, no
    console errors.

  - **Draws bills: Bill date column + a Medium width (owner).** The expanded per-draw bills table gained
    a **Bill date** column (after Bill #, via `fmtDate(b.bill_date)`). Added a third width option
    **Medium = 1500px** (between Boxed 1180 and Full 100%) in the Customize panel, and made it the new
    default - the owner wanted a width between boxed-narrow and full-wide ("chinese"). Verified live.

  - **QBO deep links are company-scoped now (owner).** The bill/invoice links used the bare
    `app/bill?txnId=` / `app/invoice?txnId=` form, which resolves the txn inside whatever Intuit company
    the browser is on - a real "wrong company" risk with the affiliated entities. Matched invoice-sync's
    fix: `load_costs` stashes the realm in a new local **`meta`** table (`qbo_realm`, **never printed**)
    after it authenticates; `fetch_data` sends it in the payload; the front-end **`qboUrl(kind,txnId)`** /
    **`qboBillHref(bareUrl)`** build the `/app/login?pagereq=…&deeplinkcompanyid=<realm>` form for invoice
    AND bill links (liens chip + draw bills), falling back to bare until a sync writes the realm. Verified
    end-to-end with a throwaway realm (payload read + rendered lien chip company-scoped + bare fallback);
    the real realm lands on the next Resync (`load_costs`). ARCHITECTURE.md updated.

  - **Resync is fast now - incremental cost pull (owner: "the ledger takes a while to update").** The
    Resync re-pulled the ENTIRE QBO cost history every time (all Bills + Purchases). `load_costs` gained
    an **incremental** mode: with a `--since` window it **skips the scoped DELETE and only upserts** by
    (txn,line), so a windowed pull adds/updates recent lines and **keeps older ones** (verified: a 2024
    line survives incremental, and a full run - no `--since` - still drops+replaces to reap QBO deletions).
    The Resync now runs `load_costs --active --since <90d>` (`WHERE TxnDate >= since` shrinks the QBO pull
    from minutes to seconds). Run a full `load_costs` occasionally to backfill older-than-window + reap
    deletions. **NEXT (owner's real point): the ledger as the CONSOLE / control plane** - a tools view
    that lists every pipeline (the 5 loaders + their upstream producers: AR sync, AP sync, WIP readers)
    with last-run + logs, and runs any one or the whole chain from the UI (generalize `_SYNC_STEPS` into
    a tool registry). See [[ledger-super-database]].

  - **Ledger as the CONSOLE / control plane - Phase B DONE (owner: "we aren't managing them via the
    ledger").** New **Console tab**: the sync engine is now a **pipeline registry** (`_pipelines()` +
    `_resolve_steps`) instead of a hardcoded loader list. Each pipeline = producer(s) + loader; a card
    shows its steps (producers flagged amber), last-run, and a **Run** button. Keys: **reload** (the
    My-view Resync - LOADERS ONLY, read-only, the safe default), **all** (Full refresh - producers too),
    per-pipeline (`ar`/`ap`/`costs`/`crm`/`wip`), and **wip-draft**. Run scope per the owner: **AR**
    (`run_invoice_sync`) and **AP** (`qbo_bill_tracker`) run the full chain incl. the real producer;
    **Costs/CRM** are loader-only; **WIP** loads the current draft, and generating a new **draft WIP for
    PM review** (the readers -> Test tabs) is a **separate, confirm-gated** button, kept OUT of any
    refresh. The run machinery (`runPipeline`/`pollSync`/`finishSync`) is parameterized so My-view + the
    Console reuse it; `GET /api/pipelines`, `POST /api/sync {pipeline}`. Verified: resolution
    (reload=no producers, all=+producers, wip-draft separate), endpoints, UI renders, no console errors.
    **Producers fire REAL syncs (Notion/Teams/Excel + Touch ID) - owner validates those with a real click.**
    NEXT: per-run log viewer in the UI; parallelize independent Notion loads; collapse the two QBO Touch IDs.

  - **P&L opens INLINE, not the side panel (owner: "turns the head unnaturally + dense info").** On the
    P&L tab, clicking a job now **expands its live P&L under the row** (full width, room for the numbers)
    instead of the right-side slide-over. Reuses `buildPnlGroup` (the same live compute + costs-by-code +
    export/folder actions), constrained to a readable width; caret + highlight on the open row; the
    duplicate "P&L" heading is hidden. `pnlExpanded` Set. Verified: expands inline, side panel does NOT
    open, no console errors. (Other detail entry points still use the slide-over; only the P&L changed.)

  - **FIX: Console AP producer pointed at the LEGACY script.** The Console's AP pipeline ran
    `bill-tracker/qbo_bill_tracker.py`, which is legacy (writes `_legacy_bill_payment_tracker.xlsx`). The
    real AP sync that produces the `Bill Tracker.xlsx` the ledger reads (OneDrive `Automations-/`) is
    **`bill-tracker/excel_bill_sync.py`** (what `run_tracker.sh` runs). Fixed the AP producer to
    `excel_bill_sync.py`. Caught by reading the code while building the cheat sheet.
  - **Cheat sheet / FAQ (owner).** New `docs/CHEATSHEET.md` - a task-oriented "how do I run it" for every
    tool from the terminal (leads with: most syncs are now one click in the Console). Worked example: the
    **statement reconciler inbox sweep** - `python3 statement-reconciler/statement_reconciler.py --inbox`
    scans the Synology **Accounting** share (`/Volumes/Accounting/Automations/Vendor Statements`), `--yes`
    for unattended, single-file + `--vendor` forms too. Commands are repo-root-relative (no `/Users` paths).

  - **Draws: "All paid" stage + a Client filter (owner).** "Ready to turn in" was shown even when the GC
    had ALREADY paid us AND vendors were paid - fully settled, nothing to turn in. Split the terminal
    state on `gc_paid_in` (AR Paid / balance <= 0): **"All paid"** (green, done) when the GC has paid you
    AND all vendors are paid; **"Collect from GC"** (amber, internal key still "Ready to turn in") when
    vendors are paid but the GC's AR is still open (you fronted it). Live now: 17 All paid, 1 Collect from
    GC (CP745, GC partially paid). Tiles + explainer updated; the **green row is now only "All paid."**
    New **Client** filter on the Draws tab - a name-substring match on the draw customer/label, so
    "Firestone" catches every Firestone job (CP672 / CP745 / CP961). Verified live, no console errors.

  - **Bill Tracker rolled into the dashboard - a new `Bills` tab with Notion-style saved views (owner).**
    The whole Bill Tracker (`ap_bill_line`, all 2,868 bills) now lives in the ledger. `_fetch_ap` gained a
    full `bills` list (per-bill `open_balance`, safe to sum); the front end fronts it with **named views
    instead of a filter panel** - each view = a predicate + default sort with a **live count** on its chip:
    **Open AP** (761) · **GC-funded · unpaid · 2mo+** (168 - the owner's example: approved + `Invoice paid`
    + still owed + bill ≥ 2 calendar months old = the pay-these-first / lien-risk list) · **Lien risk** (382)
    · **To approve** (311) · **Awaiting invoice** (473) · **No project #** (401) · **All bills** (2,868).
    Rows: division-tinted project chip (RP green / CP blue / MFD violet) · company-scoped QBO bill link ·
    bill date + age · amount · open balance · status pills (pay · invoice · lien · not-approved-only, color
    the only signal). **Group by** division/project/vendor/draw adds subtotalled headers; a division
    split-bar shows where the view's open $ sits; money cells feed the existing select-and-sum bar. The
    active view is remembered per person (`localStorage`). Row render capped at 1,500 with an honest note
    (totals still cover the full set). Read-only over `ap_bill_line`; the file is still produced by
    `excel_bill_sync.py` (`sync-ap`). Verified live against the loaded DB - all view counts, grouping,
    search, QBO scoping, and the example view checked in-browser, no console errors.

  - **Bills tab REDESIGN - Excel-dense, per-field filters (owner, 2026-08-14).** The views-only first
    cut was "almost useless": too tall (KPI cards + description + split-bar ate the viewport, cut off
    after ~7 rows), no default worth seeing, and a single search box instead of real filters. Reworked to
    read like the workbook:
    - **Default = open bills, grouped by Vendor A→Z, oldest bill first** (no config needed - that's the
      owner's daily scroll: vendors alphabetical, oldest→newest, statuses visible).
    - **Every field is its own filter dropdown** (Vendor · Division · Pay status · Invoice · Approved ·
      Lien - `buildBillFilters` populates each from the data, `Any lien risk` shortcut), AND-combined, on
      top of the quick-preset chips. The free-text search is GONE. A **Clear filters** button appears when
      any is set.
    - **Density:** the KPI cards, the view description, and the division split-bar were removed; a single
      inline quick-stat (`$ open · N lien risk`) sits in the header. `.bills-grid` tightens padding and
      keeps every row on ONE line - status is now compact **colored text** (Unpaid / No project / Due ≤7d
      / Not appr), not stacked pills. ~2-3x the rows visible.
    - **Group by** (Vendor default · None · Division · Project · Draw) orders groups **alphabetically**
      (A→Z) with a per-group subtotal; **Sort** (Oldest · Newest · Vendor · Most owed · Biggest · Lien)
      applies within groups. Money cells still feed the select-and-sum bar. Verified at **medium width
      (1500px)** per the owner - all 7 columns fit, filters/sort/group/clear all exercised, no console
      errors. Cap raised to 2,000 rows with the honest note.

  - **Bills tab: collapsible groups, split status columns, invoice slide-over (owner, 2026-08-17).**
    Five follow-ups after the dense redesign:
    - **Collapse/expand.** Each group header has a **caret** (▾/▸) and is click-to-toggle; a
      **Collapse all / Expand all** button in the filter bar toggles every group at once (`billsCollapsed`
      Set + `billGroupKeys`; the button hides when Group=None and resets when the grouping changes).
    - **Each status is its OWN column** now - **Paid · Invoice · Lien · Appr** - not merged. The owner's
      point: a merged cell hides a *missing* value; with one column each, a blank unambiguously means
      "none for this bill" (dim `–`). Compact colored text per column; header tooltips explain each.
    - **Invoice status column** shows whether the AR invoice/draw was paid (from `invoice_status`).
    - **Row click → invoice slides in on the right** (`#billDetail` slide-over): a **Bill (money out)**
      block + an **Invoice / draw (money in)** block with the real AR status, amount, GC-still-owes,
      GC-funded date, and **QuickBooks deep links to BOTH the bill and the invoice** (company-scoped).
      Needed a backend change: `_fetch_ap` now joins each bill to `billing_event` on Invoice # and
      attaches `inv_qbo_id` / `inv_ar_status` / `inv_amount` / `inv_balance` / `inv_date` (works for RP
      too, since billing_event is not RP-excluded). 1,595 of 2,891 bills carry an invoice link.
    - **Date is `MM/DD/YY`** (01/01/26) in the grid via `fmtDateShort` - still month-first (obeys the
      never-year-first rule); the detail panel keeps the readable weekday `fmtDate`. See
      [[dates-never-year-first]]. Verified live at 1500px (10 columns, caret + collapse-all, both QBO
      links open, real invoice status in the panel), no console errors.

  - **Reused the bill/invoice panel where it's a real win (owner said "apply where it's meaningful, not
    bloat", 2026-08-17).** Two surgical drill-throughs, no new patterns invented:
    - **Liens rows now open the bill/invoice slide-over** (was: the project detail). A lien IS a bill, so
      from a past-due row you see the bill (paid? approved? open $), its invoice/draw (GC paid? still
      owes?), and QBO links to BOTH - the collections decision in one place. `findBillForLien(r)` matches
      the lien-watch row back to the full enriched `BILLS` entry (by QBO bill id, else vendor+bill#+inv#);
      falls back to project detail if unmatched.
    - **Vendors rows → the Bills tab, pre-filtered to that vendor** (`jumpToVendorBills`, All-bills view).
      Spend ranking → "show me their bills." Only wired when the vendor actually has bills.
    - **Deliberately left alone** (would be bloat): Draws (already grouped + inline bill expansion + stage
      tiles + per-field filters), the Overview projects table (already has the detail panel + filters +
      sortable cols), P&L/Costs/Sales (well-suited already), and the Liens free-text search boxes (typing
      a CP#/address beats a giant dropdown there). Verified both drills live, no console errors.

  - **Bills grid: draggable column widths + wrap-when-squished (owner, 2026-08-17).** The grid is now
    `table-layout: fixed` with a `<colgroup>`; each header carries a **resize grip** on its right divider
    (visible border between headers) - drag it to set that column's width, which **persists per person**
    (`localStorage` `proficient-ledger-billcols`, `startBillColResize` updates the `<col>` + table width
    live, saves on mouseup). Cells now **wrap** (`white-space: normal; overflow-wrap: break-word`) so a
    squished column reflows instead of clipping; the vendor/name ellipsis was dropped. **Specificity
    gotcha (noted):** the base `table.grid td` rule (0,1,2) outranks `.bills-grid td` (0,1,1), so the
    bills overrides had to be written `table.bills-grid td/th` or the base `white-space: nowrap` would
    block wrapping. Verified live: drag resizes + saves, long vendor names wrap 1→3 lines at 60px, dense
    single-line at defaults, no console errors.

  - **Edit-back: lien marks from the site (owner, 2026-08-17). Phase 1 (overlay) DONE; Phase 2 (workbook
    mirror) PENDING.** The owner wanted to mark a bill (e.g. "Notice Sent") on the website and have it
    reflect in the workbook. Design chosen: **overlay + mirror on sync** (not a direct dashboard write to
    the live OneDrive .xlsx - file-lock/corruption risk, and it'd fight the AP sync for ownership).
    - **New writable overlay: `bill_mark`** (2nd write surface after `waiver`). Keyed by the **QBO bill id
      (= the workbook's hidden `_Key`)** so a mark survives every `ap_bill_line` reload AND joins back to
      the workbook. Lives in `shared/bill_marks.py` (the ONE module both the dashboard-writer and the
      AP-reader use - no tool imports another tool). Absent-safe (no ledger → `{}`).
    - **Dashboard:** `POST /api/bill-mark` `{bill_id, lien}` (lien ∈ Notice Sent / Lien Filed / ✓ Released,
      or '' to clear); `_fetch_ap` merges marks over the loaded `lien_status` **instantly** (bills +
      lien_watch) and exposes `bill_id` / `lien_marked`. The bill detail panel gained a **Mark lien**
      control (3 buttons + clear); it POSTs, re-pulls authoritative data, re-opens. Works from the Liens
      tab too (same panel). Verified live on a DB copy: mark → grid + panel + `bill_mark` row all update;
      clear reverts; real ledger untouched; no console errors.
    - **PENDING - Phase 2 (the workbook mirror):** `excel_bill_sync.py` must call
      `bill_marks.read_lien_marks()` + `bill_marks.resolve_lien(...)` at its Lien-cell preservation step so
      the mark lands in `Bill Tracker.xlsx` on the next `sync-ap` (then the existing `_Key` preservation
      keeps it). NOT done yet because that file had **uncommitted changes from the other session** - editing
      it would entangle their work (can't stage just my hunks). Do it once that file is clean. Until then
      the mark is live on the site + persists, but the workbook won't show it. `resolve_lien` is already
      written and unit-checked, so Phase 2 is a ~2-line call at one spot.

  - **Sub LOC tracker - a live tab (owner, 2026-08-17).** How much we've FRONTED to subs before the client
    repaid us: the working-capital float + the LOC to size to. The proven model (`one-offs/sub_loc_report.py`,
    validated to a real $1.70M peak) was **extracted into `shared/sub_loc.py`** (the engine: read-only QBO
    pull -> per-project, per-draw-period, chronological FIFO) so both the Excel report AND a new ledger
    loader use ONE copy (the one-off was refactored to import it, 675->334 lines, no duplication).
    - **`ledger/load_sub_loc.py`** runs the engine, full-replaces `sub_loc_event` (the DRAW/REPAY timeline
      + running LOC balance) and upserts `sub_loc_run` (the summary). `--selftest` proves it offline (no
      QBO); one Touch ID per real run; company_id never printed. Console pipeline **"Sub LOC (QBO float)"**.
    - **Dashboard `Sub LOC` tab** (`_fetch_sub_loc`): headline KPIs - **Fronted, still out** (today's float)
      · **Peak LOC needed** (+ date; size the LOC to this - research says a LOC ~10-20% of revenue, but the
      real need is the peak) · **Avg draw->repay** days · **Prefunded**. Then by division, a **repayment
      feed** ("a client payment paid off these fronted subs" = the owner's "this got paid -> settled $X"),
      and by project. Verified live on a DB copy (peak $162k / outstanding $62k sample), real ledger
      untouched, no console errors.
    - **QC (owner-requested):** high-effort review found 5, all fixed - guarded the `sub_loc_event` query so
      a partial load can't blank the whole dashboard; killed the model duplication (the refactor above);
      dropped an unused `--show`; un-scoped `.cell.open-amt` so the float emphasis actually renders; docs
      updated (this + ARCHITECTURE). Aesthetic pass: reuses `.kpi`/`.grid`, color only on the green
      settlement amounts, one font. See [[sub-loc-model]].

  - **BUG FIX: Console "last ran" was blank for AR (owner, 2026-08-18).** After a `sync-all` the AP card
    showed a time but AR didn't. Cause: `_freshness` derived the AR "last ran" from `Open_Invoices.xlsx`
    at OneDrive candidate paths, but the owner's live AR mirror is **`Collections/Invoice Tracker.xlsx`**
    (the old file is now `backup_dont_use Open_Invoices.xlsx`) - so it never resolved. Fix: resolve the AR
    source as `ACB_INVOICE_TRACKER_XLSX` (default `Collections/Invoice Tracker.xlsx`) first, and in the
    Console **fall back to the ledger `loaded_at`** when no source file is found so a card is never blank.
    Also added AR / CRM / Sub LOC to the ledger-freshness loop, so all six pipeline cards now show a time.
    Verified live: AR resolves to 2026-08-18T10:08 (this morning), all cards populated, no console errors.
    (AR data itself was never affected - `billing_event` had loaded fine; this was a display bug.)

  - **Sub LOC tab refinements (owner, 2026-08-18).** Made the tab work the way the owner asked:
    - **By project FIRST** (the priority), rows **clickable → a drill-over** showing that project's still-open
      subs **grouped by the draw they sit under**, each draw's **status + billed/still-owed**, each sub bill
      **linking to QuickBooks**, and an **"All transactions in QuickBooks"** link (the `customerdetail?nameId=`
      page - the standard "everything under this customer/project" view, per the owner's research ask).
    - **Repayment feed bucketed** This week / This month / **Prior months (collapsed)**; **By division**
      collapses too (both default-collapsed to keep By-project front and centre).
    - Engine (`shared/sub_loc.py`): draws now carry the **QBO bill id**, invoices carry **amount/balance/
      status + the QBO customer id**, and `attach_open_by_project` builds the per-project open-subs-by-draw
      structure. Loader stores it as `open_by_project` JSON on `sub_loc_run`; `_fetch_sub_loc` serves it.
    - **Migration guard (caught in test):** `CREATE TABLE IF NOT EXISTS` won't add the new column to an
      existing `sub_loc_run`, so the loader now `ALTER TABLE ... ADD COLUMN open_by_project` if missing -
      without it a re-run on a pre-drill-down ledger crashed. Verified live on a DB copy (MFD300 → two draw
      groups Unpaid/Partially-Paid with sub bills + both QBO link kinds), no console errors. See [[sub-loc-model]].

  - **Bills lien marks: STAGE + Save (owner, 2026-08-18).** The marks were auto-saving on every click (no
    data loss - the owner's 8 Notice-Sent marks were all in `bill_mark`), but the owner expected the Save
    workflow he'd designed. Now: a mark **stages** (optimistic on the panel + grid, nothing written); a
    bottom-centre **Save bar** shows "N unsaved lien mark(s)" with **Save** (batch-POSTs to `/api/bill-mark`,
    then reloads authoritative) and **Discard** (reverts). **`beforeunload` guards** leaving the page with
    unsaved marks, and the 90s auto-refresh pauses while marks are pending so the optimistic view holds.
    (The workbook mirror is still Phase 2, unchanged.) Verified live on a DB copy: stage → bar + leave-warn,
    Save persists + clears, Discard reverts, no console errors.

  - **Open Invoices tab - the AR aging, with a Lien column (owner, 2026-08-18).** "All open invoices … the
    way the aging shows, it's perfect … a new column of Lien that talks to the Notion Lien Tracker."
    - **`load_invoices.py` enriched:** each Invoice Tracker page now also captures **Due Date · Net Terms ·
      Notion Aging Bucket · Litigation**, and resolves its **`Lien` relation** against the Notion **Lien
      Tracker** (`load_lien_index` → one `{page_id: Status}` query; `LIEN_DS` = `2c5b…`, override
      `ACB_LIEN_TRACKER_DS_ID`) to the matching **lien Status** (most-escalated when an invoice has >1 lien).
      Read-only add-on: any read error / DB-not-shared degrades the lien column to blank, the AR load still
      runs. New `billing_event` cols (`due_date, net_terms, aging_bucket, litigation, lien_status,
      lien_notice`) with an **ALTER-based migration** for existing DBs. `--selftest` extended (due date +
      escalated-lien resolution) and green.
    - **`dashboard.py::_fetch_open_invoices`:** open invoices (balance>0) aged by **DUE DATE** into
      Current/1-30/31-60/61-90/90+ using the **same thresholds as `invoice-sync/aging_sheet.py`** (live-
      computed; falls back to Notion's stored bucket only when there's no due date). Sorted client → due.
    - **Dashboard `Invoices` tab** (`renderOpenInvoices`): an AR-aging grid - client-banded + collapsible,
      the open balance **tinted green→red in its one bucket column**, a **per-bucket grand total**, and the
      green→red **bucket tiles double as a filter**. Filters: client · project · division · **lien** ·
      litigation · sort. **Lien column** shows the Notion status (Mailed / Lien filed / Ready→mail / …).
      Division dots, MM/DD/YY dates, a ⚖ litigation marker, and a **QBO deep link on every Invoice #**.
      Colors are theme-aware (blended toward the foreground → legible in light + dark).
    - Verified live on a synthetic DB (7 open invoices spanning all 5 buckets, 3 clients, liens +
      litigation): tiles + grand total reconcile, every filter/collapse works, QBO links company-scoped,
      no console errors, colors pass in both themes. Lien column needs the Lien Tracker DB shared with the
      automation integration to populate (blank until then). See [[open-invoices-tab-spec]].
    - **Kept connected (owner follow-on):** the relation->Status logic was extracted to
      **`shared/lien_status.py`** (repo rule: a file a second tool needs moves to shared/), and
      invoice-sync's AR Aging Excel now carries the **same "Lien status" column** beside its deadline
      clock - the site and the workbook read the one resolver, so they can't drift. See invoice-sync/STATUS.md.

  - **Bills tab: collapse-by-default + scannable amounts (owner, 2026-08-18).** The Bills tab now opens
    with every vendor group **collapsed** (like Costs/Draws - `load()` seeds `billsCollapsed` from the
    current grouping; the group-by dropdown re-collapses under the new key). Each group header shows the
    **open $ + bill count at the SAME size/weight as the vendor name**, right after it, so the amount reads
    at a glance (was a small dim sub-line). Fixed a resulting bug: the "Showing N of M - narrow with a
    filter" note keyed off `rendered < rows.length`, which collapse makes true, so it printed "Showing 0 of
    6"; it now fires only when the 2000-row CAP actually truncated (a `capped` flag). Same header treatment
    applied to the Open Invoices group rows for consistency. Layout note: the flex sits on an inner div,
    not the `<td>` (flex on a `<td>` drops table-cell layout and collapses the colspan). Verified live.

  - **Dashboard version badge (owner, 2026-08-18).** `dashboard.py` now carries `LEDGER_VERSION`
    (surfaced in `meta.version`, rendered as a small `vX.Y.Z` pill beside the "Project Ledger" title) so
    the owner can confirm which build is live. Bump on every user-visible release. **`1.0.0`** = the
    Open Invoices tab + lien columns. Loaded live against the real ledger: 341 invoices, $5.27M open
    aged across all five buckets. NOTE: the lien column came back blank because the automation
    integration can't read the Lien Tracker DB yet (`NotionError`) - the owner shares that DB with the
    integration to light it up (the loader degrades gracefully and says so).

  - **v1.0.1 - Open Invoices client shows the parent GC (owner, 2026-08-18).** The Client column showed
    the project-level `Customer (raw)` ("MFD177 - MERRITT PARK") instead of the parent client. Fixed:
    `load_invoices` now resolves each invoice's `Customer` **relation** to the parent title via the new
    shared **`shared/notion_customers.py`** (the ONE parent-client resolver, also used by invoice-sync's
    Excel so both name the client identically - repo rule). Reads the Res/Com + MFD customer lists
    (`ACB_INVOICE_RESCOM_CUST_DS_ID` `19db…` / `ACB_INVOICE_MFD_CUST_DS_ID` `34bb…`, override via env),
    merges them, falls back to `Customer (raw)` when a relation doesn't resolve; a customer list the
    token can't read is skipped (those invoices show the project name). Selftest extended. Loaded live:
    all 98 open invoices now show a real client (51 distinct GCs, 0 project-shaped names).

  - **Open Invoices refinements (owner, 2026-08-19).** A batch of feedback on the Invoices tab:
    - **Lien column now shows the computed CLOCK + the Notion status.** `_fetch_open_invoices` adds
      `lien_due_label`/`lien_due_state` from **`shared/lien_clock`** (the SAME clock the AR Aging Excel
      uses, so the site and workbook agree - "when the lien is due"); the frontend renders it colored by
      urgency, with the **Notion Lien Tracker status beside it as a Notion-style PILL** (grey pill +
      colored dot) so it reads as "from Notion". (Pill is blank until the Lien Tracker DB is shared.)
    - **Litigation excluded by DEFAULT** (was Include); the filter box turns **red** whenever it's
      hiding/limiting rows so it's obvious a filter is in place. 76 of 98 shown by default (22 in litigation).
    - **Removed the Due column** (the payment-due date; days-past-due moved to the Date cell's tooltip).
    - **Vertical rules between the aging columns**; the client group-header band is now **neutral grey**
      (was the green accent tint the owner flagged).
    - **Sort now orders the client GROUPS** by the chosen key (Oldest-due-first really puts the oldest
      client on top), fixing the "says Oldest due first but shows Client A-Z" mismatch.
    - Verified live (v1.1.0): pill + clock render, litigation red/excluded, no Due column, group order.
    - STILL PENDING (next): Project # -> QBO project page (needs the QBO customer id per project;
      `shared/qbo_costs` already computes it from `CustomerRef.value` - persist it to `cost_line`).

  - **Project P&L batch (owner, 2026-08-19).** Four asks on the P&L tab:
    - **Costs card removed from the Console** - it's now `hidden` (kept in the reload/all chains so
      Resync still refreshes costs for the P&L, which reads `cost_line`); it just isn't a standalone
      button. Costs belong in a future "Company P&L" view (NOT built - owner "don't build, just remove").
    - **Renamed the tab + heading "P&L" -> "Project P&L"** (reserving "Company P&L" for the period view).
    - **"P&L updated" column** = when each project's project-pnl Excel was last generated
      (`pnl_paths.find_pnl(p).mtime`, batched into `_portfolio_pnl` - ~4 stats/project, cached client-side);
      shows "6d ago" / "not generated".
    - **Client column + client filter** on the by-job table (client = the resolved GC from
      `billing_event.customer`, fallback `v_wip_latest.builder_or_gc`). 134 of 137 active jobs resolve a client.
    - Verified live (v1.1.0): heading, Client + P&L-updated columns, client filter (BEDROCK -> its jobs),
      costs card gone from the Console, reload still runs load_costs, no console errors.

  - **Project # -> QBO project page (owner, 2026-08-19).** Clicking a project # (Open Invoices AND
    Project P&L) opens that project's QBO **customerdetail** page (all its transactions), via the existing
    `qboCustomerUrl`. The QBO customer id per project = **`CustomerRef.value`**, which `shared/qbo_costs`
    already computes but `load_costs` dropped - now **persisted to `cost_line.customer_id`** (schema column +
    ALTER migration + added to the insert cols). The dashboard maps `{project -> customer_id}` from
    `cost_line` and attaches `cust_id` to every invoice/P&L row; the link is absent-safe (plain text until a
    costs pull populates it). Populated by re-running load_costs (Resync) - one QBO pull, Touch ID.

  - **Grouped two-level nav + Customers page + Payments stub (owner, 2026-08-19).** The flat tab bar became a
    **two-level grouped nav** (data-driven `NAV_GROUPS` in app.js; `buildGroupBar`/`buildSubTabs`; `setTab`
    highlights the group + renders the second row): **My view · Overview · Financials** (Project P&L · Costs)
    **· Customers** (Customers · Invoices · Draws · Payments · Sales Outreach) **· Vendor Center** (Vendors ·
    Bills · Sub LOC · Liens) **· IT** (Systems · Graph · Console). A group opens its landing (first) tab; single-page
    groups hide the sub-row. The **Vendors** page is renamed **"Vendor Center"** and **Sales -> "Sales
    Outreach"**. New **Customers** page - top clients by open AR (from the open invoices; click a client ->
    Invoices filtered to them; 51 clients / $5.27M). **Payments** is a clear stub (received payments + the AP
    due out under a fully-paid invoice - needs a payment feed; built next). Verified live: all 15 tabs route,
    group/sub highlight, no console errors. See [[ledger-expansion-backlog]].

  - **Payments page (owner, 2026-08-19).** Every invoice the GC has paid (full or part): received =
    gross - open balance; when **fully paid**, **AP due out** = the open `ap_bill_line` on that project (the
    vendor cash to send out now). `load_invoices` now captures the tracker's **Paid Date** (billing_event
    `paid_date` col + ALTER migration); `_fetch_payments` joins billing_event (balance < amount) to the AP
    open-per-project and the `cost_line` customer_id (project# -> QBO). `ap_due_total` sums DISTINCT paid
    jobs (a job with 2 paid invoices counts once). Frontend `renderPayments`: KPIs (received / count / paid-in-
    full / AP due out) + a table (Paid · Client · Project# → QBO · Invoice# → QBO · Invoiced · Received ·
    Status · AP due out in red). Verified live: 275 payments, $12.6M received, 248 paid in full, $4.13M AP
    due out; no console errors. (Per-invoice payment status, not individual QBO Payment txns - fine for v1.)

  - **Payments REDESIGN → payment-as-transaction (owner, 2026-08-19: "just the payment as a transaction then
    see below as grouped the invoices it pays, it's simple").** Rebuilt off the actual QBO **Payment** objects,
    not invoice rows (billing_event is invoice-level, so a cheque paying N draws was not reconstructable there -
    it is only a real thing in the Payment's `Line[].LinkedTxn`). NEW loader **`load_payments.py`** pulls Payment
    txns (rolling window, default 12 months) → two spine tables **`payment`** (the transaction: date · customer ·
    total · ref#/method · unapplied) + **`payment_application`** (one row per invoice a payment paid, with the
    applied amount). Each linked invoice is resolved to its invoice # / project via `billing_event`, then - for the
    ones that aged out of the tracker (billing_event holds only ~351 open/recent) - **by a second QBO pull of those
    invoices by Id** (`shared.qbo_api.extract_proj` → project/division). Idempotent full-replace; read-only on QBO;
    `--selftest` offline. `_fetch_payments` returns each payment with its `applications[]` nested; `renderPayments`
    renders a collapsible **payment header** (green, → QBO customer) with the **invoice rows grouped beneath**
    (invoice# → QBO · project# → QBO · division · applied $), default expanded, Collapse-all. Verified live:
    **846 payments · $38.67M received · 1,415 invoice links (100% resolved)**; multi-invoice payments group
    correctly (one cheque → up to 32 invoices); no console errors. Supersedes the per-invoice v1 above (AP-due-out
    dropped per "it's simple").
    **GC normalization (owner "Yes normalize for sure!").** QBO records a payment against the bare leaf customer,
    which is often the project sub-customer (`RP6676-FTW`, `CP790 - HUNTER RANCH AMENITY`) - not the client. The
    loader now pulls the QBO Customer hierarchy once (2,012 customers) and walks each payer's `ParentRef` to the
    TOP parent = the GC, stored as `payment.parent_customer` / `parent_customer_id` (ALTER-migrated). The header
    shows the GC (→ QBO), the project stays in the invoice sub-rows. All 846 payers resolved to a real GC (e.g.
    RP6676-FTW → LONESTAR GREEN HOMES, CP790 → DL MEACHAM LP); root-walk stops at the deepest KNOWN ancestor so an
    inactive parent (absent from the active-only pull) never crashes it.

  - **Payments v3: columnar transaction + "Unlocks (AP)" side menu (owner, 2026-08-19: "client, Payment Ref #,
    Payment Type, Amount Paid · grouped: Invoices paid, the total open and the amount applied · then see what bills
    that unlocks via the side menu by talking to bills").** Each payment is now a proper ROW - **Client · Payment
    Ref # · Payment Type · Amount Paid · Unlocks (AP)** - default COLLAPSED (scannable); click to reveal a nested
    grid of the invoices it paid (**Invoice # · Total open · Amount applied**). `payment_application` gained
    `invoice_open` (the invoice's QBO `Balance`, captured in both resolution paths; ALTER-migrated). **Payment Type
    fix:** `PaymentMethodRef` returns an id ONLY (no name) - added `_payment_method_map` (pull the PaymentMethod
    entity) so the column fills (Check 527 · ACH 167 · QB Payments-Bank 109 · Credit Card 10 · Direct Deposit/Wire/
    Zelle/Cash). **Side menu** `openPaymentBills` (new `#payBills` panel): click a payment's "Unlocks (AP)" and the
    open vendor bills on that payment's project(s) slide in, grouped by job (Vendor · Bill # · Open · Status, → QBO)
    - the money-IN → money-OUT link, read live from the already-loaded Bills tab data. Verified live: DL MEACHAM's
    $33,516 Check unlocks $85,758 across 20 open bills on CP790; no console errors.

  - **Customer Center: top clients PER DIVISION (owner, 2026-08-19: "biggest client useless... give me the top
    clients per division instead").** `renderCustomers` now groups open-AR clients under Commercial / Residential /
    Multi Family bands (each a header with the division's open $ + client count), clients ranked by open AR within.
    The useless "Biggest client" KPI is replaced by **per-division open-AR tiles**. Verified live: Commercial $1.22M/
    15 clients, Residential $1.12M/35, Multi Family $3.12M/3.

  - **Nav rename: group = singular, page = "... Center" (owner, 2026-08-19: "name the parent header of Vendor =
    Vendor, the page Vendor Center is where all the vendors are listed. SAME FOR CUSTOMERS").** Group headers are now
    **Customer** and **Vendor** (singular); the first page under each is **Customer Center** / **Vendor Center**.
    NAV_GROUPS labels + TAB_LABELS + the index.html page heading updated.

  - **QBO field reference logged (owner, 2026-08-19: "are you logging all the qbo codes... so we evolve and not
    regress").** The id-only-Ref gotcha + Payment anatomy added to CLAUDE.md "QBO API gotchas"; a matching
    "API field reference - QBO codes → what they mean for us" section added to the vault's
    `04_integrations/quickbooks-online.md` (PaymentMethod values, Payment=money-in, leaf-vs-GC customer, invoice Balance).

  - **Lien clock: two stages - notice, then the AFFIDAVIT deadline (owner, 2026-08-19: "if mailed at or before
    deadline mark done and then move ... lien due ... research when Texas decides a notice cannot fork over to a
    lien due to time passing").** Researched (Tex. Prop. Code Ch. 53, post-SB219): the real cutoff is the **lien
    affidavit filing deadline = 15th of the 4th month (MFD/CP) / 3rd month (RP) after the work month** (§53.052) -
    a FIXED date one month past the notice deadline, NOT 30 days from the mailing (the "~30 days" is coincidental).
    `shared/lien_clock.py` gained `AFFIDAVIT_MONTHS`, `affidavit_deadline`, and `status_stage`; `lien_state` now
    takes `lien_status=` (the Notion Lien Tracker status). When it's **Mailed/Sent**, the clock advances from the
    notice deadline to the affidavit deadline (label prefixed **LIEN**, e.g. `LIEN Sep 15, 2026 · 27d` or
    `LIEN PAST DUE · Aug 14, 2026` once the window closes); a **Lien/Paid/Released** status ends the clock; anything
    else stays on the notice clock. Backward-compatible - callers passing no status are unchanged (money_bleeds uses
    `notice_deadline` directly and is untouched). Wired in BOTH `ledger/dashboard.py` (Open Invoices) and
    `invoice-sync/export_invoices_xlsx.py` (AR Aging Excel) so the site and workbook agree. Verified live: 5 Mailed
    invoices show their affidavit date (4 past due = window closed, RP7466 has 27d left); retainage/Lien/None paths
    unchanged; no console errors.

  - **Bills: red line means an actual notice/lien, + a month→day date filter (owner, 2026-08-19: "need
    date filter by month first then specific day like Excel" · "the red line, why only put it in notice due?
    it's when there is a notice or lien filed that we should make red").** (1) The red row line (`.risk`
    left-border) now keys off **`BILL_LIEN_ACTIVE = {Notice Sent, Lien Filed}`** - an actual notice sent or lien
    filed - NOT the computed deadline states (`BILL_LIEN_RISK`, which still drives the separate "Lien risk"
    view). Verified live: 69 Notice-Sent bills all red, 249 Notice-PAST-due bills none red. The deadline urgency
    still shows in the lien CELL colour. (2) NEW **Month + Day filters** lead the Bills filter bar: pick a month
    (newest first) and the Day list populates with just that month's days (Excel-style drill-down; Day disabled
    until a month is picked). Wired into `billPassesFilters` by `bill_date` prefix; Clear resets both. Verified:
    July 2026 → 432 bills, then Jul 1 → 15. **Month upgraded to MULTI-SELECT (owner 2026-08-20: "multi select
    the months ... all of June and back ... checkbox ... auto selects the prior months ... option to remove those
    priors").** The Month field is now a checkbox dropdown (`.msel`): checking a month adds it AND every OLDER
    month (`billMonths` Set; "June and back"), and individual priors can be unchecked; button shows "Jun 2026 +8".
    The Day drill stays, enabled only when exactly one month is selected. Verified live: Jun → 9 checked, remove
    May → 8, Dec 2024 (oldest) → 1 bill + Day on.

  - **Sub LOC: dashboard-shaped, + repayment line items (owner, 2026-08-19: "by division top ungrouped ·
    repayment feed needs to show me the line items in side menu · by project collapse all only show top 5 most
    in the hole then expand more (x count) ... a dash not a huge long list").** (1) **By division** moved to the
    TOP, flat/ungrouped (its own plain widget, always shown - no more collapsed section). (2) **By project** now
    shows only the **top 5 most in the hole** (outstanding desc) with an **Expand more (N)** toggle (373 hidden →
    "Show top 5"). (3) **Repayment feed rows are clickable → a side panel lists the fronted subs that payment paid
    down** (FIFO oldest-first: Sub · Bill # → QBO · Fronted date · Applied $ · Fully/Partial). NEW DATA: the FIFO
    engine (`shared/sub_loc.py`) records a `settled[]` list per REPAY event → `sub_loc_event.settled` (JSON,
    ALTER-migrated) → `_fetch_sub_loc` parses it → `openSublocRepay` renders it. Verified live: INV 34482 (DL
    MEACHAM, CP790) settled $25,428 across 11 subs; 181 of 308 repayments carry line items (the rest are pure
    prefunding/surplus); no console errors.
    **Affordance polish (owner: "green bar too long · hard to know if i click the project # it shows costs").**
    The By-project project # is now an **accent link** (`.row-open`, with a `›` on row hover) so the drill-in is
    obvious; caption says "click a project #". And the Excel-style money-cell SELECTION now **hugs the value**
    (`.cell-sel .cell` gets the box, not the full wide `<td>`) so a selected amount reads as a tight box, not a
    long green bar - app-wide.

  - **Bills show the CLIENT + a Client filter/group (owner, 2026-08-20: "put Project:Customer instead of the
    project name ... need a filter for customer to see all open bills for both client then dive deeper per
    project").** The Bills PROJECT cell now shows the **chip + the CLIENT** (e.g. `CP790 · DL MEACHAM LP`) instead
    of the job name (job name moved to the tooltip). Client resolved per project in `_fetch_ap` via a new
    `proj_customer` map: **payments' QBO-hierarchy GC first** (`payment_application` → `payment.parent_customer`,
    most reliable), else a `billing_event` customer that isn't project-shaped; attached as `b.client` (2,354/2,891
    bills). NEW **Client filter** (`#bfCustomer`, leads the field group) + a **Customer group-by** so you can see
    all of a client's open bills across projects, then group/drill per project. Verified live: CP790 → DL MEACHAM
    LP, filter → 98 bills, group-by-client works; no console errors.
    **Client resolved for EVERY project (owner 2026-08-20: "Project tells you the customer ... Customer:Project,
    then reverse it ... it's all blank ... you are missing where it's at").** The first pass only resolved paid/
    invoiced projects (left ~537 blank). Fixed at the source: every project is a QBO **sub-customer** whose name
    carries the project # (Customer:Project), so `load_payments` now **reverses that hierarchy for all customers**
    → a new **`project_customer`** table (project_no → client + client_id), 1,514 projects mapped. `_fetch_ap`
    reads it FIRST (payments/billing_event only fill gaps). Coverage 2,354 → **2,455 of 2,891**; the remaining 436
    are genuinely project-less bills (397, no project = no client) or multi-project bills (37). **NOTE: needs a
    `load_payments` re-run to populate `project_customer` AND a dashboard-server restart to serve `b.client`.**

  - **Bills Vendor filter → multi-select, pumps excluded by default (owner, 2026-08-20: "vendor filter to be
    able to select ... by default don't show MCP/Core Concrete Pumping, still have info there in case we want to
    remove filter ... say 'all vendors except pumps'").** Vendor is now a checkbox dropdown (`.msel`, like Month)
    with a search box (100+ vendors) + a "Show all (incl. pumps)" link. Checked = shown. On first build it defaults
    to **hiding every pump vendor** (matched by `/pump/i`: Core, MCP, Nelson, Five Star) - the data stays, just
    filtered; check them back or "Show all" to include. Button reads **"All vendors except pumps"** at the default,
    "All except N" when changed, "All vendors" when nothing hidden. `billVendorHidden` Set drives `billPassesFilters`;
    "Clear filters" resets to the default (pumps hidden) and only lights up when the vendor pick deviates. Verified
    live: default 2,569 of 2,891 (322 pump bills hidden), check Core back → "All except 3" and its bills return.

  - **Project P&L generation shows a live status line (owner, 2026-08-20: "see the terminal status of where
    it's at not just Generating and the seconds ... don't pull full log, just one line that is updating").** The
    generation subprocess runs with `PYTHONUNBUFFERED=1` so project-pnl's progress flushes live to its per-project
    log; `_pnl_status` returns `status` = **`_pnl_last_line(log)`** (tail only, ANSI + box-drawing stripped, blank/
    border lines skipped - one clean line, capped 130 chars). Frontend: the poll (now every 1.5s) shows
    `Generating… (Ns) · <that line>` instead of just the seconds, falling back to "Touch ID may be waiting" until
    the log has content. Unit-verified (`_pnl_last_line` extracts `✓ costs pulled: 128 lines` from a mixed log);
    live line needs a real generation to watch.

  - **All applicable Bills filters are now multi-select (owner, 2026-08-20: "the same for all filters that are
    applicable / multi select makes sense to add").** One generic checkbox component (`BILL_MSEL` + `buildBillMSel`
    / `billMSelPasses`) converts **Client · Division · Pay status · Invoice · Approved · Lien** from single-selects
    to multi-select dropdowns (include model: empty = all, checked = only those). Client carries a search box (113
    values); Lien shows the short labels (Past due / Notice sent / …); blanks/`(multiple)` are labeled. Day stays a
    single drill (dependent on one month); Vendor keeps its bespoke pump-exclude default; Month keeps its cascade.
    Only one menu opens at a time; "Clear filters" resets every multi-select. Verified live: RP+MFD → "2 selected",
    1,758 bills, no CP rows; per-menu search/labels/clear all work; no console errors.

  - **Lien-mark REVIEW summary + a Synology lien-folder link (owner, 2026-08-20: "Lien marks saved - press and
    it shows me all the marks i made ... first do the lien summary so i know what i changed before saving ... link
    the lien marks to a new folder in my synology").** Pressing the save-bar text ("N unsaved lien marks" / "Lien
    marks saved") opens a **review panel** (`openLienReview`, new `#lienReview` slide-over): an **Unsaved changes**
    section showing each staged mark as **old → new** (with Save/Discard right there), and an **On file** section
    of the marks currently in effect (`lien_marked` bills). So nothing saves blind. Plus an **"Open lien folder ↗"**
    button → `POST /api/lien/folder` (`_LIEN_FOLDER`, machine.env `LIEN_FOLDER`, default
    `/Volumes/Accounting/LIENS & MONTHLY NOTICES/Vendor Liens/2026`) opens the Synology folder cross-platform via
    `_os_open`. **Organized BY VENDOR, auto-created on mark (owner's pick, 2026-08-20):** `?vendor=` drills to
    `2026/<VENDOR>` and **creates it if missing** (`_ensure_lien_vendor_dir` - only the subfolder, only if the base
    exists, so it never shadows an unmounted share; path-traversal guarded). Saving a **Notice Sent / Lien Filed**
    mark auto-creates that vendor's folder (`_set_bill_mark` takes the vendor; `saveBillMarks` passes it); the bill
    detail has a per-vendor **"Open lien folder ↗"** button. Verified live: staged 2 → review shows `– → Notice
    sent` / `– → Lien filed` + 8 on file; POST with a vendor creates + opens `2026/<vendor>`. Route is a **POST**
    (like `/api/job/open`); needs a dashboard-server restart. Next option (not built): attach the actual PDF path
    to each mark.

  - **Liens page: real columns (all divisions, not CP) + multi-select filters (owner, 2026-08-20: "Liens ... Vendor
    first, Date, amount, Invoice #, Client, project #, Invoice associated, Invoice Payment status ... you put cp# as
    if this page is only cp# ... need filters too not just search").** The watchlist is now built from the ENRICHED
    bills (`_fetch_ap` computes `lien_watch` from `bills` after the client/AR-invoice join), so it carries client,
    bill date, and the AR invoice + its pay status. Columns are now **Vendor · Date · Amount · Invoice # · Client ·
    Project # · Invoice associated (the AR draw → QBO) · Invoice pay status (did the client pay it)** - `Project #`,
    not `CP #`, and it's ALL divisions (verified: MFD325/RP present). NEW **multi-select filters** (Client · Vendor ·
    Division · Invoice status - the same checkbox UI as Bills, `LIEN_MSEL`/`buildLienMSel`/`lienMSelPasses`) plus a
    Project # search box; the old 4 search boxes are gone.
    **SCOPED to sent/filed only (owner 2026-08-20: "i don't need on the lien clock on that page, this page is
    solely for what has actually been sent/filed").** `_fetch_ap` now also returns **`ap["liens"]`** = bills whose
    mark is **Notice Sent / Lien Filed** (Filed first), and the page renders THAT, not the deadline clock. Header
    "Liens filed & notices sent"; KPIs Notices sent / Liens filed / Open $ at stake; tiles Lien filed · Notice sent.
    Verified live: **88 sent/filed (5 filed, 83 sent), client 88/88, date 88/88, pay-status 78/88** (the 10 blanks
    are invoices not in the tracker/billing_event); 0 blank client/date; no console errors. NOTE: the "date/client/
    pay-status blank" the owner saw was the **un-restarted server** serving the pre-enrichment `lien_watch` - the
    data was always resolvable (client is 100% on sent/filed). Backend change - needs the dashboard restart.

  - **Client on the Overview projects list + job name/address in Project P&L (owner, 2026-08-20: "need client
    here in overview, it's important to always see client, in project P&L get the job name/address as well").**
    Extracted the client resolver to a shared **`_project_customer_map(con)`** (was inline in `_fetch_ap`) so
    Bills, Liens, the projects list, and the P&L all name the client identically. `fetch_data` attaches `client`
    to every project row (152/170); the Overview table gets an **always-on Client column** (`always:true`, survives
    a customized column set) + client in the project search. `_portfolio_pnl` now selects `project_name` and uses
    the shared resolver → the P&L "By job" table leads **Project · Name / address · Division · Client** (name
    170/170, client 157/170). Verified live on :8787: CP585 → Green Road Construction; CP800 → TOPAZ AT LIGHT
    FARMS / Tri-C; no console errors.

  - **WIP report tab under Financials (owner, 2026-08-20: "give me the wip report here now under financials, use
    the master test sheet for reference").** New **WIP report** tab (Financials group, between Project P&L and
    Costs). `renderWip` renders the company WIP schedule straight from the ledger's `wip_snapshot` (already served
    in `data.projects` / `ALL` - no backend change), columns + order mirroring **Test-Master** exactly: Project # ·
    Name · Bonded · Total Contract · Est. Total Costs · Original Profit · GP % · Costs to Date · Cost to Complete ·
    % Complete · Revenues Earned · Profit Earned · Billed · Over · Under · Left to Bill · Future Profit · Pure Job
    Borrow. Grouped by division (band header) with a **subtotal per division + a GRAND TOTAL**; **Active only**
    toggle (default on). Read-only (the master workbook stays where WIP is edited). Verified live on :8787: 137
    active jobs, report Aug 7 2026, the top MFD job's contract / 14.5% GP / 98.5% complete tie to the sheet;
    grand total ~$29.7M contract; no console errors. Frontend-only - no server restart needed.
    - **Enhancements (owner, 2026-08-20: "conditional formatting ... a greater visual representation of job
      performance + ability to change column size. also freeze the header. drop bonded in this version only").**
      (a) **Bonded dropped** from `WIP_COLS` - DASHBOARD ONLY; the Excel Test tabs keep it. (b) **Frozen header**:
      the WIP `.table-scroll` gets a `.wip-scroll` class (bounded `max-height`, `overflow:auto`) so the base
      `.grid` sticky thead truly freezes as the schedule scrolls (verified: header pins at container-top offset 0
      after a 900px scroll). (c) **Resizable columns**: reuses the Bills grip pattern - a `<colgroup>` + a
      `.col-resize` grip per header, `startWipColResize`, widths persist in `localStorage` (`proficient-ledger-wipcols`);
      `.wip-grid` is now `table-layout:fixed`. (d) **Conditional formatting** (`_wipCond`, color only to encode) with
      sign conventions taken straight from `wip/wip_writer.py`: GP% thin/negative red · healthy green · **>30% amber**
      (owner's "missing cost" flag); **% Complete** an in-cell data bar; **Overbillings** green (holding the GC's
      cash); **Underbillings** red scaled by contract (financing the job); **Pure Job Borrow** amber→red (a cash
      drain); **Future Profit** red when negative; **Profit Earned** red when negative. A compact color **legend**
      sits above the table. Verified live on :8791: 137 GP tints, 130 data bars, 7 overbill / 8 underbill / 5
      job-borrow / 18 future-profit tints fire on real rows; no console errors. Frontend-only.

  - **Project P&L: status column + sort dropdown + status filter (owner, 2026-08-19).** `_portfolio_pnl` now
    returns ALL jobs with a `status` + `active` flag (company/division TOTALS stay active-only). Frontend:
    a **Status column** (green Active / dim Closed·Complete), a **status filter** (default Active only; All;
    Closed only), and a real **sort dropdown** (worst/best margin · most earned · most cost · biggest contract
    · project #) alongside the clickable headers. Verified live: 170 rows (137 active shown by default, 33
    Closed under the filter), sort re-orders, no console errors. STATUS SOURCE: **CP = Test-CP** (Active/
    Closed/Complete), **MFD = Test-Master** (blank -> Active by construction). **RP: all 119 read "Active"
    because the RP source (Test-RP) is active-only by construction - closed RP jobs drop OUT of the WIP master
    entirely, so there is no "Closed" RP row to load.** CORRECTION (owner, 2026-08-19: "was already done...
    schedule is mounted?"): the daily SCHEDULE *is* mounted (`/Volumes/Common/OPERATIONS/SCHEDULE`,
    `shared/schedule.py` reads it) and its per-job on-schedule mark IS already in the ledger as
    `wip_snapshot.mark_schedule` (✓ 77 / ✗ 42) - it was never "not loaded". So the finer RP signal (on the
    crew schedule vs off) already exists; a true Active/Closed for RP would mean loading closed RP jobs from a
    source that carries them (not the active-only Test-RP). See [[ledger-expansion-backlog]].

## IN PROGRESS
- **Lien-mark workbook mirror (Phase 2 above):** add the `bill_marks.resolve_lien` call to
  `bill-tracker/excel_bill_sync.py`'s Lien preservation once that file has no foreign uncommitted changes.
- Owner to validate the producer Runs (AR/AP) and the draft-WIP button with a real click (real syncs + Touch ID).

## TO DO
- **Sync-diff view (owner building an HTML prototype separately, 2026-08-20):** a "what changed since the
  last sync" view. Owner is drafting the HTML on their side and will hand it over to merge into the dashboard
  later - parked here so it is not lost; no action until the owner shares it.
- **Investigate the ~6 active reconcile mismatches** (QBO cost ≠ WIP figure, e.g. RP6901/RP6440):
  likely a base/-FTW split or a stale WIP cost — owner review; keep dollar specifics out of the repo.
- `budget_line` from the takeoff/ETC extractor by cost code (`shared/takeoff_etc.py` is project-total
  today — needs per-code) → enables budget-vs-actual from the spine.
- Postgres deployment decision (Synology container vs. small cloud box) — schema is ready either way.
- Optional: a read-only dashboard over `v_wip_latest` (Phase 3) — DB first, UI later.

- **Console: Project P&L workbooks, per division (2026-08-25).** A card with
  three buttons - Active CP / Active RP / Active MFD - each regenerating every
  ACTIVE job of that division via `project-pnl/project_pnl_export.py active
  <div>`. It is a GENERATOR, not a loader: its steps live under `actions`, not
  `steps`, so `_resolve_steps("reload"/"all")` can never sweep it in - a Full
  refresh must not kick off 138 P&L runs. Same isolation trick the WIP draft
  uses. Writes Excel only; the ledger is untouched. `pnl` alone and any
  unknown `pnl-<x>` resolve to zero steps and are rejected by `_sync_start`.

## OPEN ISSUES / NOTES
- MFD rows come from `Test-Master`, which has **no STATUS column** → MFD `status` loads as NULL
  (MFD closures are manual anyway). RP/CP status comes from their own tabs.
- `wip_snapshot` stores the master's already-computed figures verbatim (source of truth = the
  sheet, per the "do not generate" instruction). When Phase-2 granular data exists, snapshots can
  be reconciled against a spine-computed WIP as a cross-check.
- The RP tab has no retainage / over-under / earned columns (those are computed on the master);
  for RP those snapshot fields load as NULL by design.
- **P&L active-scope (accepted, from the QC):** `_project_pnl`/`_pnl_pl` compute a P&L for ANY project
  (so a Closed job's P&L is still openable), but `_portfolio_pnl` counts **active-only** (status Active
  or blank/MFD). So "sum of every per-project P&L" won't equal the company total when a non-active job
  still carries costs — intended (active-only company view), flagged so it isn't mistaken for drift.
- **In-app Resync scope:** runs the ledger loaders (fresh QBO + Notion, re-reads the current WIP/Bill
  Tracker Excel). The WIP master + `Bill Tracker.xlsx` are still produced by their own flows (the owner
  / `sync-ap`); chaining those upstream syncs into the button is a possible future step.
