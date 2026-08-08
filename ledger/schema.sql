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


-- ── cost_line : actual spend from QBO bills (append-only, idempotent) ────────
-- Keyed by (bill id, line id) so re-running a connector can never double-count.
-- This is the (TxnId, ...) idempotency rule promoted from convention to schema.
CREATE TABLE IF NOT EXISTS cost_line (
    qbo_txn_id      TEXT NOT NULL,               -- QBO Bill Id
    qbo_line_id     TEXT NOT NULL,               -- line within the bill
    project_no      TEXT NOT NULL REFERENCES project(project_no),
    cost_code       TEXT REFERENCES cost_code(code),  -- via cost_leaf(); NULL = account-based line
    amount          NUMERIC NOT NULL,
    txn_date        TEXT,
    is_sub          INTEGER NOT NULL DEFAULT 0,
    PRIMARY KEY (qbo_txn_id, qbo_line_id)
);


-- ── billing_event : AR invoices / draws (append-only, idempotent) ───────────
CREATE TABLE IF NOT EXISTS billing_event (
    qbo_txn_id      TEXT PRIMARY KEY,            -- QBO Invoice Id
    project_no      TEXT NOT NULL REFERENCES project(project_no),
    amount          NUMERIC NOT NULL,            -- gross billed
    txn_date        TEXT,
    draw_period     TEXT                         -- from PrivateNote (the QBO custom field is unreachable)
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
    source_sheet  TEXT,               -- Bills | Inventory
    source        TEXT NOT NULL DEFAULT 'bill_tracker',
    loaded_at     TEXT NOT NULL
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
