-- schema.sql — the canonical "project ledger" for Proficient Concrete.
--
-- WHAT THIS IS
-- The single spine every system feeds. QBO stays the books, JobTread stays the
-- ops shell, Excel goes back to being an export — but the *reconciled shape of a
-- JOB* (identity + budget + costs + billing + the computed WIP position) lives
-- here, in one place, owned by us instead of by a vendor's data model.
--
-- PORTABLE DDL — runs on BOTH SQLite (the Phase-1 spike) and PostgreSQL (the real
-- deployment) UNCHANGED. The choices that keep it portable:
--   * natural primary keys only            (no SERIAL / AUTOINCREMENT)
--   * timestamps stored as ISO-8601 TEXT   (Postgres users may retype as TIMESTAMPTZ)
--   * booleans stored as INTEGER 0/1        (Postgres users may retype as BOOLEAN)
--   * DROP VIEW + CREATE VIEW               (CREATE VIEW IF NOT EXISTS isn't portable)
--   * INSERT ... ON CONFLICT upserts        (works on SQLite >=3.24 and Postgres >=9.5)
-- SQLite only: turn on FK enforcement per-connection with  PRAGMA foreign_keys = ON;
--
-- WHAT FILLS WHAT
--   project, wip_snapshot ......... the WIP-master loader (load_wip_master.py) — TODAY
--   cost_code, budget_line,
--   cost_line, billing_event ...... the QBO connectors (bill-tracker, invoice-sync,
--                                   the WIP readers) — Phase 2, one connector at a time
-- ============================================================================


-- ── project : the aggregate root. One row per real job. ─────────────────────
-- project_no is THE join key across every system. -FTW is its own project
-- (RP7186 and RP7186-FTW are two distinct rows — never family-matched).
CREATE TABLE IF NOT EXISTS project (
    project_no      TEXT PRIMARY KEY,           -- 'RP7358', 'RP7279-FTW', 'MFD177', 'CP672'
    division        TEXT NOT NULL,              -- 'Multi Family' | 'Commercial' | 'Residential'
    is_ftw          INTEGER NOT NULL DEFAULT 0, -- 1 when project_no ends with -FTW
    name            TEXT,
    type            TEXT,                        -- RP only: 'Tract' | 'Custom'  (never repurposed)
    builder_or_gc   TEXT,                        -- RP builder; CP/MFD GC when known
    bonded          INTEGER,                     -- 1 / 0 / NULL
    rp_category     TEXT,                        -- RP band: GOOD / NOT STARTED / FTW WITH COSTS / ...
    updated_at      TEXT NOT NULL               -- ISO-8601 of the last load that touched this row
);


-- ── cost_code : job-type prefix + number, a FIRST-CLASS dimension ───────────
-- Cost codes stop being a string smuggled through QBO item names and become a
-- real thing with a foreign key. (Populated in Phase 2 by the QBO connectors.)
CREATE TABLE IF NOT EXISTS cost_code (
    code            TEXT PRIMARY KEY,            -- 'SL1', 'PV6', 'FW3', ...
    prefix          TEXT,                        -- SL / PV / FW / PR / WL / CS / MS
    description     TEXT
);


-- ── budget_line : the plan (takeoff / ETC) by cost code ─────────────────────
CREATE TABLE IF NOT EXISTS budget_line (
    project_no      TEXT NOT NULL REFERENCES project(project_no),
    cost_code       TEXT NOT NULL REFERENCES cost_code(code),
    etc_amount      NUMERIC NOT NULL,
    source          TEXT,                        -- 'takeoff' | 'estimator' | ...
    updated_at      TEXT NOT NULL,
    PRIMARY KEY (project_no, cost_code)
);


-- ── cost_line : actual spend from QBO, ONE ROW PER expense line ─────────────
-- The COMPLETE cost ledger (incl. subs), keyed by cost code via shared/qbo_costs
-- cost_leaf() — the same resolver project-pnl uses, so the two can't drift.
-- Keyed by (txn id, line id) so re-running can never double-count. Filled by
-- ledger/load_costs.py from a QBO pull. cost_code is set only for coded lines
-- (SL/PV/…); account-based lines carry `account` with cost_code NULL.
CREATE TABLE IF NOT EXISTS cost_line (
    qbo_txn_id      TEXT NOT NULL,               -- QBO Bill / Purchase Id
    qbo_line_id     TEXT NOT NULL,               -- line within the txn
    txn_type        TEXT,                        -- 'Bill' | 'Expense'
    project_no      TEXT NOT NULL REFERENCES project(project_no),
    cost_code       TEXT REFERENCES cost_code(code),  -- via cost_leaf(); NULL = account-based line
    account         TEXT,                        -- resolved account/category name
    amount          NUMERIC NOT NULL,
    txn_date        TEXT,
    is_sub          INTEGER NOT NULL DEFAULT 0,
    vendor          TEXT,
    description     TEXT,
    source          TEXT NOT NULL DEFAULT 'qbo',
    loaded_at       TEXT NOT NULL,
    PRIMARY KEY (qbo_txn_id, qbo_line_id)
);


-- ── billing_event : AR invoices / draws (append-only, idempotent) ───────────
CREATE TABLE IF NOT EXISTS billing_event (
    qbo_txn_id      TEXT PRIMARY KEY,            -- QBO Invoice Id
    doc_number      TEXT,                        -- Invoice # (the draw's invoice_no — join key to ap_bill_line)
    project_no      TEXT,                        -- from CustomerRef.name (soft link; AR invoices may be off-WIP)
    division        TEXT,
    customer        TEXT,                        -- the GC / client billed
    memo            TEXT,                        -- invoice memo (the draw memo)
    amount          NUMERIC,                     -- TotalAmt — gross billed = the NET the GC pays (money IN)
    balance         NUMERIC,                     -- open balance (0 = the GC has paid this draw)
    txn_date        TEXT,                        -- invoice date
    status          TEXT,                        -- Paid | Open (derived from balance)
    draw_period     TEXT,                        -- from PrivateNote (the QBO custom field is unreachable)
    source          TEXT NOT NULL DEFAULT 'qbo_invoice',
    loaded_at       TEXT NOT NULL
);


-- ── wip_snapshot : the COMPUTED WIP position, one row per (project, date) ────
-- This is exactly what a Test tab of "WIP - MASTER new.xlsx" holds. The loader
-- lands it here so the monthly WIP stops living in a spreadsheet. Re-running for
-- the same report_date REPLACES the row (idempotent).
CREATE TABLE IF NOT EXISTS wip_snapshot (
    project_no                TEXT NOT NULL REFERENCES project(project_no),
    report_date               TEXT NOT NULL,     -- 'YYYY-MM-DD'
    status                    TEXT,              -- Active / Closed / Complete
    original_contract         NUMERIC,
    approved_cos              NUMERIC,
    total_contract_price      NUMERIC,
    original_estimated_cost   NUMERIC,
    co_costs                  NUMERIC,
    estimated_total_costs     NUMERIC,           -- ETC (the budget)
    original_profit           NUMERIC,
    gross_profit_pct          NUMERIC,
    costs_to_date             NUMERIC,
    cost_to_complete          NUMERIC,
    percent_complete          NUMERIC,
    revenues_earned_to_date   NUMERIC,
    profit_earned_to_date     NUMERIC,
    billed_to_date            NUMERIC,
    overbillings              NUMERIC,
    underbillings             NUMERIC,
    retainage_held            NUMERIC,
    left_to_bill              NUMERIC,
    future_profit_to_earn     NUMERIC,
    pure_job_borrow           NUMERIC,
    mark_schedule             TEXT,              -- RP cross-check marks (✓ / ✗)
    mark_general_list         TEXT,
    mark_jobtread             TEXT,
    notes                     TEXT,
    source_tab                TEXT,              -- which Test tab this row came from
    loaded_at                 TEXT NOT NULL,
    PRIMARY KEY (project_no, report_date)
);


-- ── ap_bill_line : vendor bills from Bill Tracker (AP + lien tracking) ───────
-- This is NOT the cost ledger. Bill Tracker's display sheets EXCLUDE subs, so it
-- cannot state a job's true cost (subs are most of it — that lives in wip_snapshot
-- and, later, the QBO-complete cost_line). What it uniquely carries is AP pay
-- status and the Texas lien clock per bill. project_no is a SOFT link (no FK): AP
-- tracks every bill, including jobs not in the WIP-derived project table. Loaded by
-- a full replace each run (source = 'bill_tracker'), so it mirrors the current file.
CREATE TABLE IF NOT EXISTS ap_bill_line (
    line_uid      TEXT PRIMARY KEY,   -- 'bt:<sheet>:<excel_row>' — stable within a file version
    project_no    TEXT,               -- soft link to project(project_no); may be NULL / off-WIP
    division      TEXT,
    vendor        TEXT,
    bill_ref      TEXT,               -- vendor Bill #
    bill_date     TEXT,
    account       TEXT,               -- QBO account/category (NOT the SL/PV cost code)
    description   TEXT,
    line_amount   NUMERIC,
    bill_total    NUMERIC,
    open_balance  NUMERIC,
    pay_status    TEXT,
    approved      TEXT,
    lien_status   TEXT,               -- 'Notice due in ≤15d' / 'Notice PAST due' / ...
    matched_invoice TEXT,             -- the DRAW this bill is on (AR invoice it authorizes)
    invoice_status  TEXT,             -- Bill Tracker pipeline: Invoice paid / Awaiting Invoice / ...
    invoice_no      TEXT,
    gc_paid_date    TEXT,             -- when the client/GC funded that draw
    pay_date        TEXT,             -- when we paid the vendor
    bt_key          TEXT,             -- Bill Tracker _Key (stable-ish bill id)
    qbo_link        TEXT,             -- QBO deep link to the bill (from the "Open" ↗ =HYPERLINK column)
    source_sheet  TEXT,               -- Bills | Inventory
    source        TEXT NOT NULL DEFAULT 'bill_tracker',
    loaded_at     TEXT NOT NULL
);

-- ── waiver : an owner input - unconditional waiver received per bill ─────────
-- The ledger is read-only except the owner's own marks; this is one such surface
-- (owner marks "waiver in hand" so a draw can be turned in to unlock the next).
-- Keyed by (draw, vendor, bill) so it survives ap_bill_line reloads.
CREATE TABLE IF NOT EXISTS waiver (
    waiver_key      TEXT PRIMARY KEY,  -- hash of matched_invoice + vendor + bill_ref
    matched_invoice TEXT,
    vendor          TEXT,
    bill_ref        TEXT,
    received        INTEGER NOT NULL DEFAULT 0,
    received_date   TEXT,
    note            TEXT,
    updated_at      TEXT NOT NULL
);

-- ── bill_mark : an owner input - the lien tag set on the dashboard per bill ───
-- The other writable overlay (see shared/bill_marks.py). Keyed by the QBO bill id
-- (= Bill Tracker's hidden _Key), so a mark survives ap_bill_line reloads AND joins
-- cleanly back to the workbook: the dashboard writes it here instantly, and the next
-- excel_bill_sync run mirrors it into the workbook's Lien cell. lien '' = cleared.
CREATE TABLE IF NOT EXISTS bill_mark (
    bill_id     TEXT PRIMARY KEY,      -- QBO bill TxnId (the workbook _Key)
    lien        TEXT,                  -- 'Notice Sent' | 'Lien Filed' | '✓ Released' | '' (cleared)
    updated_at  TEXT NOT NULL
);

-- ── v_ap_by_project : open AP + bill counts per project ─────────────────────
DROP VIEW IF EXISTS v_ap_by_project;
CREATE VIEW v_ap_by_project AS
SELECT project_no,
       COUNT(*)                                                    AS bill_lines,
       SUM(CASE WHEN COALESCE(open_balance,0) > 0 THEN 1 ELSE 0 END) AS open_lines,
       SUM(COALESCE(open_balance,0))                               AS open_balance,
       SUM(COALESCE(line_amount,0))                                AS billed_amount
FROM ap_bill_line
GROUP BY project_no;

-- ── v_cost_by_project : loaded QBO cost per project (reconcile vs WIP) ──────
DROP VIEW IF EXISTS v_cost_by_project;
CREATE VIEW v_cost_by_project AS
SELECT project_no,
       SUM(amount)                                  AS costs_loaded,
       SUM(CASE WHEN is_sub = 1 THEN amount ELSE 0 END) AS sub_costs,
       COUNT(*)                                      AS lines
FROM cost_line
GROUP BY project_no;

-- ── v_cost_by_code : per-project cost-code drill (budget-vs-actual base) ────
DROP VIEW IF EXISTS v_cost_by_code;
CREATE VIEW v_cost_by_code AS
SELECT project_no,
       COALESCE(cost_code, account, '(unclassified)') AS code,
       cost_code,
       SUM(amount)                                    AS actual,
       COUNT(*)                                       AS lines
FROM cost_line
GROUP BY project_no, COALESCE(cost_code, account, '(unclassified)');

-- ── action : ledger action items mirrored to Notion (the folder-memory link) ─
-- sync_actions.py upserts a Notion page per action item and reads its Status
-- back here; the dashboard shows the Notion link + status. The ledger stays the
-- RADAR — the work/thread/done lives in the Notion page.
CREATE TABLE IF NOT EXISTS action (
    action_key      TEXT PRIMARY KEY,   -- 'draw:<invoice>' | 'lien:<proj>:<bill>' | 'overbudget:<proj>' | ...
    type            TEXT,               -- draw | lien | overbudget | underbilled
    project_no      TEXT,
    title           TEXT,
    amount          NUMERIC,
    status          TEXT,               -- Open | Working | Done  (mirrored FROM Notion)
    notion_page_id  TEXT,
    notion_url      TEXT,
    synced_at       TEXT
);

-- ── v_wip_latest : each project joined to its most-recent snapshot ──────────
-- The "one pane of glass" query. This is the thing currently rebuilt in Excel
-- every month — here it is one definition, computed once.
DROP VIEW IF EXISTS v_wip_latest;
CREATE VIEW v_wip_latest AS
SELECT s.*,
       p.division,
       p.is_ftw,
       p.name          AS project_name,
       p.type          AS project_type,
       p.builder_or_gc,
       p.bonded,
       p.rp_category
FROM wip_snapshot s
JOIN project p ON p.project_no = s.project_no
WHERE s.report_date = (
    SELECT MAX(x.report_date) FROM wip_snapshot x WHERE x.project_no = s.project_no
);


-- ============================================================================
-- CRM / SALES PIPELINE — the pre-project spine (leads → clients)
-- The ledger owns the customer master too, not just the job. QBO/JobTread/Excel
-- still run the job; Notion's "Customer List" is just the outreach feed. Filled by
-- load_customers.py (read-only on Notion, idempotent full-replace).
-- ============================================================================

-- ── customer : one row per Notion "Customer List" page ──────────────────────
-- Identity + current pipeline stage + who sourced it (created_by) and who last
-- worked it (last_edited_by — the honest per-rep attribution, from Notion's own
-- system fields, no manual Owner property to maintain). customer_key = the Notion
-- page id (dashless), stable across reloads.
CREATE TABLE IF NOT EXISTS customer (
    customer_key     TEXT PRIMARY KEY,
    name             TEXT NOT NULL,
    division         TEXT,               -- Residential | Commercial | Multi-Family (joined if multi)
    main_status      TEXT,               -- Active | Qualified | Pending Approval | Inactive
    sales_status     TEXT,               -- Lead | Follow up | Contacted | Interested | No response | Closed - Won | Closed - Lost
    last_contacted   TEXT,               -- ISO date
    follow_up_date   TEXT,               -- ISO date
    referral         TEXT,               -- multi-select, joined with ', '
    primary_contact  TEXT,
    primary_email    TEXT,
    primary_phone    TEXT,
    created_by       TEXT,               -- who sourced the record  (Notion "Created by")
    last_edited_by   TEXT,               -- who worked it last       (Notion "Last edited by") = rep attribution
    last_edited_time TEXT,               -- ISO datetime
    n_touches        INTEGER NOT NULL DEFAULT 0,  -- interaction-log lines parsed from the page body
    notion_url       TEXT,
    source           TEXT NOT NULL DEFAULT 'notion_customer_list',
    loaded_at        TEXT NOT NULL
);

-- ── sales_touch : one row per "History of interactions" log line ────────────
-- The outreach touch log lifted out of the Notion page body into queryable rows.
-- seq preserves on-page order; touch_date is parsed when the line carries a date.
-- Wholly owned by load_customers.py (full-replaced each run).
CREATE TABLE IF NOT EXISTS sales_touch (
    customer_key    TEXT NOT NULL REFERENCES customer(customer_key),
    seq             INTEGER NOT NULL,
    touch_date      TEXT,               -- ISO date if the line carried one, else NULL
    note            TEXT NOT NULL,
    PRIMARY KEY (customer_key, seq)
);

-- ── v_sales_pipeline : customer counts by pipeline stage ────────────────────
DROP VIEW IF EXISTS v_sales_pipeline;
CREATE VIEW v_sales_pipeline AS
SELECT COALESCE(sales_status, '(none)') AS sales_status,
       COUNT(*)                          AS customers,
       SUM(n_touches)                    AS touches
FROM customer
GROUP BY COALESCE(sales_status, '(none)');

-- ── v_sales_by_rep : who worked what — attribution by last editor ───────────
DROP VIEW IF EXISTS v_sales_by_rep;
CREATE VIEW v_sales_by_rep AS
SELECT COALESCE(last_edited_by, '(unknown)')                          AS rep,
       COUNT(*)                                                       AS worked,
       SUM(CASE WHEN sales_status = 'Contacted'    THEN 1 ELSE 0 END) AS contacted,
       SUM(CASE WHEN sales_status = 'Interested'   THEN 1 ELSE 0 END) AS interested,
       SUM(CASE WHEN sales_status = 'Closed - Won' THEN 1 ELSE 0 END) AS won
FROM customer
GROUP BY COALESCE(last_edited_by, '(unknown)');

-- ── meta : tiny key/value for local runtime facts (NOT business data) ────────
-- e.g. 'qbo_realm' = the company realm id, written by a loader that authenticates
-- to QBO, so the dashboard can build COMPANY-SCOPED deep links (open the txn in the
-- right Intuit company) without itself touching QBO/Keychain. Local DB only; the
-- realm is never printed to a terminal or a log (owner rule).
CREATE TABLE IF NOT EXISTS meta (
    key        TEXT PRIMARY KEY,
    value      TEXT,
    loaded_at  TEXT
);
