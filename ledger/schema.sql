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
    customer_id     TEXT,                        -- QBO CustomerRef.value (the project's customer id) → customerdetail deep link
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
    due_date        TEXT,                        -- invoice due date (ages the AR; from the tracker / QBO DueDate)
    net_terms       TEXT,                        -- Net 15/30/45/... (display)
    aging_bucket    TEXT,                        -- Notion-computed bucket, fallback only when due_date is null
    litigation      INTEGER NOT NULL DEFAULT 0,  -- 1 = in litigation (shown, but flagged off the clean aging)
    lien_status     TEXT,                        -- related Lien Tracker Status (Mailed/Lien/Paid/...) or NULL
    lien_notice     TEXT,                        -- related Lien Tracker Notice Type (RP/CP/MFD Notice, Intent to Lien...)
    paid_date       TEXT,                        -- Paid Date from the tracker (when the GC paid; drives the Payments tab)
    draw_period     TEXT,                        -- from PrivateNote (the QBO custom field is unreachable)
    note            TEXT,                        -- Notion Quick Status: the collections one-liner ("GC paying Fri")
    source          TEXT NOT NULL DEFAULT 'qbo_invoice',
    loaded_at       TEXT NOT NULL
);


-- ── project_customer : project # → its CLIENT (the GC), from the QBO hierarchy ──
-- Every project is a QBO sub-customer whose name carries the project #, nested under the
-- GC (Customer:Project). Reversing that gives project → client for EVERY project, not just
-- the ones that happen to have a payment/invoice. Filled by load_payments.py (which already
-- pulls the whole customer list) as a full replace; read by the Bills tab.
CREATE TABLE IF NOT EXISTS project_customer (
    project_no  TEXT PRIMARY KEY,
    client      TEXT,                        -- the GC = top of this project's QBO customer chain
    client_id   TEXT,                        -- the GC's QBO customer id (deep link)
    loaded_at   TEXT NOT NULL
);


-- ── payment : a received customer payment - money IN as ONE transaction ──────
-- A QBO Payment object: the GC hands over cash once, and that one payment can
-- settle several invoices. This is the TRANSACTION (total, date, who paid); the
-- invoices it lands on live in payment_application below. Loaded by a full
-- replace each run (load_payments.py, source = 'qbo_payment') over a rolling
-- window, so the table mirrors the last pull. Not the books - QBO is.
CREATE TABLE IF NOT EXISTS payment (
    qbo_txn_id      TEXT PRIMARY KEY,            -- QBO Payment Id
    txn_date        TEXT,                        -- payment date (when the money came in)
    customer        TEXT,                        -- CustomerRef.name (raw QBO leaf - often the project sub-customer)
    customer_id     TEXT,                        -- CustomerRef.value (QBO deep link to the leaf)
    parent_customer     TEXT,                    -- the GC: QBO customer hierarchy walked to its top parent
    parent_customer_id  TEXT,                    -- the top parent's id (deep link to the GC)
    total_amt       NUMERIC,                     -- TotalAmt - the whole payment
    unapplied_amt   NUMERIC,                     -- UnappliedAmt - credit not yet on any invoice
    method          TEXT,                        -- PaymentMethodRef.name (Check / ACH / ...) if present
    ref_no          TEXT,                        -- PaymentRefNum (check #) if present
    source          TEXT NOT NULL DEFAULT 'qbo_payment',
    loaded_at       TEXT NOT NULL
);


-- ── payment_application : which invoice(s) a payment paid, and how much ───────
-- One row per (payment, invoice) link (QBO Payment Line -> LinkedTxn of type
-- Invoice). amount = the slice of the payment applied to that invoice. invoice_no
-- / project_no / division are resolved against billing_event by the invoice's QBO
-- id (NULL when the invoice isn't in the tracker window - still shown by id).
CREATE TABLE IF NOT EXISTS payment_application (
    payment_txn_id  TEXT NOT NULL,               -- -> payment.qbo_txn_id
    invoice_txn_id  TEXT,                         -- LinkedTxn.TxnId (QBO Invoice Id)
    invoice_no      TEXT,                         -- resolved DocNumber (billing_event join)
    project_no      TEXT,
    division        TEXT,
    amount          NUMERIC,                      -- money applied to THIS invoice from THIS payment
    invoice_open    NUMERIC,                      -- the invoice's current open balance (0 = fully settled)
    PRIMARY KEY (payment_txn_id, invoice_txn_id)
);
CREATE INDEX IF NOT EXISTS ix_payapp_payment ON payment_application (payment_txn_id);
CREATE INDEX IF NOT EXISTS ix_payapp_invoice ON payment_application (invoice_no);

-- ── bill_payment : money OUT to vendors — the QBO BillPayment transaction ────
-- A cheque/ACH can pay several bills at once (Line[].LinkedTxn -> Bill), so the
-- payment lives on top and the bills it covered are its lines. The vendor page
-- reads this ON DEMAND per vendor (never the bulk load). Loaded by
-- load_bill_payments.py (read-only QBO pull, this-year window, DELETE+reload).
CREATE TABLE IF NOT EXISTS bill_payment (
    qbo_txn_id  TEXT PRIMARY KEY,                 -- QBO BillPayment Id
    txn_date    TEXT,                             -- when we paid the vendor
    vendor      TEXT,                             -- VendorRef.name
    vendor_id   TEXT,                             -- VendorRef.value (QBO deep link)
    total_amt   NUMERIC,                          -- TotalAmt — the whole payment
    pay_type    TEXT,                             -- PayType: Check | CreditCard
    ref_no      TEXT,                             -- DocNumber (cheque #) if present
    source      TEXT NOT NULL DEFAULT 'qbo_billpayment',
    loaded_at   TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS ix_billpay_vendor ON bill_payment (vendor);
CREATE TABLE IF NOT EXISTS bill_payment_line (
    payment_id  TEXT NOT NULL,                    -- -> bill_payment.qbo_txn_id
    bill_id     TEXT,                             -- the Bill this payment applied to (QBO Bill Id)
    amount      NUMERIC,                          -- money applied to THIS bill from THIS payment
    PRIMARY KEY (payment_id, bill_id)
);
-- ── vendor_ap : open AP per vendor, straight from QBO (Bill.Balance > 0) ──────
-- The Vendor Center's "Open $" / "Open bills". Pulled by load_bill_payments from
-- QBO Bills so it covers EVERY vendor (incl. subs, which the Bill Tracker sheets
-- drop) and uses QBO vendor names (which match cost_line, so the join is clean).
CREATE TABLE IF NOT EXISTS vendor_ap (
    vendor      TEXT PRIMARY KEY,                 -- QBO VendorRef.name
    vendor_id   TEXT,
    open_bal    NUMERIC,                          -- sum of open Bill balances
    open_bills  INTEGER,                          -- how many bills still open
    loaded_at   TEXT NOT NULL
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

-- ── sub_loc_event / sub_loc_run : the subcontractor LOC float model ──────────
-- How much we have FRONTED to subs before the client repaid us (money out to subs, netted
-- against client payments per project+draw-period, FIFO, chronological). The engine is
-- shared/sub_loc.py (read-only QBO pull); loaded by ledger/load_sub_loc.py, full-replaced
-- each run. `sub_loc_event` = the DRAW (sub paid) / REPAY (client paid) timeline with the
-- running LOC balance; `sub_loc_run` = the single latest run's summary (peak = the LOC to
-- size; outstanding = still fronted right now). NOT a source of truth - a derived snapshot.
CREATE TABLE IF NOT EXISTS sub_loc_event (
    seq         INTEGER PRIMARY KEY,   -- chronological order within the load
    event_date  TEXT NOT NULL,         -- ISO date
    type        TEXT NOT NULL,         -- 'DRAW' (sub paid, money out) | 'REPAY' (client paid, money in)
    project     TEXT,
    division    TEXT,                  -- MFD | CP | RP | Other
    party       TEXT,                  -- the sub (DRAW) or the client (REPAY)
    out_amt     NUMERIC DEFAULT 0,     -- money OUT: the sub payment
    in_amt      NUMERIC DEFAULT 0,     -- money IN applied to the LOC (repayment)
    lag_days    INTEGER,               -- REPAY: days from the oldest draw it settled
    balance     NUMERIC,               -- running LOC balance AFTER this event
    note        TEXT,                  -- 'prefunded $x' / 'surplus $x'
    invoice     TEXT,                  -- REPAY: the client invoice # that paid
    reimb       TEXT,                  -- DRAW: JSON [[invoice, iso_date], ...] that later settled it
    settled     TEXT,                  -- REPAY: JSON [{party, bill_id, bill_ref, draw_date, amount, fully, lag_days}] the subs it paid down
    loaded_at   TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS sub_loc_run (
    id            INTEGER PRIMARY KEY CHECK (id = 1),   -- single latest run
    window_start  TEXT,
    window_end    TEXT,
    peak          NUMERIC,             -- high-water LOC balance = the LOC you truly need
    peak_date     TEXT,
    outstanding   NUMERIC,             -- fronted-but-uncollected RIGHT NOW
    total_drawn   NUMERIC,
    total_repaid  NUMERIC,
    prefunded     NUMERIC,             -- client cash that arrived before the sub was paid (zero-LOC)
    avg_lag       NUMERIC,             -- amount-weighted draw->repay days
    n_draws       INTEGER,
    divisions     TEXT,                -- JSON: per-division {peak, drawn, repaid, outstanding, avg_lag}
    projects      TEXT,                -- JSON: per-project rows (drawn/repaid/outstanding/avg_lag)
    open_by_project TEXT,              -- JSON: {project: {cust_id, open, groups:[{period, draw, subs[]}]}}
    loaded_at     TEXT NOT NULL        -- for the project drill-down (open subs grouped by draw)
);

-- ── health_snapshot : the company-health metric layer (QBO-only numbers) ─────
-- The few numbers the Health tab can NOT derive from the ledger tables: bank
-- balances (cash), the retainage GL accounts, the P&L blocks (MTD/YTD/prior -
-- margins + break-even inputs), the 13-week cash flow (burn/runway), and the
-- recurring-obligations register (FIN-12). One row per payload, JSON body
-- (same derived-snapshot pattern as sub_loc_run). Full-replaced by
-- ledger/load_health.py each run; as_of = the pull time. NOT a source of
-- truth - QBO is; cash moves only when uploads are entered, so the tab shows
-- as_of on every cash figure.
CREATE TABLE IF NOT EXISTS health_snapshot (
    key        TEXT PRIMARY KEY,   -- 'bank_accounts' | 'retainage' | 'pl_blocks' | 'weekly_flow' | 'recurring'
    payload    TEXT,               -- JSON
    as_of      TEXT,               -- ISO datetime of the QBO pull
    loaded_at  TEXT NOT NULL
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

-- ── performance indexes for the ON-DEMAND slice reads ────────────────────────
-- The dashboard never loads a whole table into the browser - it reads small slices
-- per vendor / project / draw when you open something. These indexes keep those
-- slice reads O(log n) as the tables grow, instead of a full table scan. (Placed
-- after all tables so every referenced table exists when the schema is applied.)
CREATE INDEX IF NOT EXISTS ix_apbill_vendor    ON ap_bill_line  (vendor);          -- vendor page bills
CREATE INDEX IF NOT EXISTS ix_apbill_draw      ON ap_bill_line  (matched_invoice); -- draws roll-up
CREATE INDEX IF NOT EXISTS ix_costline_sub     ON cost_line     (is_sub, project_no); -- subs on a draw
CREATE INDEX IF NOT EXISTS ix_costline_proj    ON cost_line     (project_no);       -- project cost slices
CREATE INDEX IF NOT EXISTS ix_sublocevent_proj ON sub_loc_event (project);          -- sub-LOC source drill-down
