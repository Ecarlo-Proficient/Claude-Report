/* ── Ledger dashboard front-end ──────────────────────────────────────────────
   Vanilla JS, no build step, no external libraries. Fetches /api/data and
   renders KPIs + division rollup + a searchable/sortable projects table with a
   click-into-job detail panel. Appearance is customizable and saved per person
   in localStorage.
--------------------------------------------------------------------------- */
"use strict";

const $  = (sel, el = document) => el.querySelector(sel);
const $$ = (sel, el = document) => [...el.querySelectorAll(sel)];

// ── Column catalog ────────────────────────────────────────────────────────
const COLUMNS = [
  { key: "project_no",            label: "Project #",     type: "text",   align: "left", always: true },
  { key: "division",              label: "Division",      type: "text",   align: "left", def: true },
  { key: "project_name",          label: "Name",          type: "text",   align: "left", def: true },
  { key: "client",                label: "Client",        type: "text",   align: "left", always: true },   // always shown (owner: "always see client")
  { key: "status",                label: "Status",        type: "status", align: "left", def: true },
  { key: "rp_category",           label: "Category",      type: "text",   align: "left" },
  { key: "builder_or_gc",         label: "Builder / GC",  type: "text",   align: "left" },
  { key: "total_contract_price",  label: "Contract",      type: "money",  def: true },
  { key: "estimated_total_costs", label: "ETC",           type: "money",  def: true },
  { key: "costs_to_date",         label: "Costs",         type: "money",  def: true },
  { key: "percent_complete",      label: "% Complete",    type: "pct",    def: true },
  { key: "billed_to_date",        label: "Billed",        type: "money",  def: true },
  { key: "left_to_bill",          label: "Left to Bill",  type: "money",  def: true },
  { key: "overbillings",          label: "Over",          type: "money" },
  { key: "underbillings",         label: "Under",         type: "money" },
  { key: "retainage_held",        label: "Retainage",     type: "money" },
  { key: "gross_profit_pct",      label: "GP %",          type: "pct" },
  { key: "original_profit",       label: "Orig. Profit",  type: "money" },
  { key: "costs_loaded",          label: "QBO Costs",     type: "money" },
  { key: "sub_costs",             label: "Subs (QBO)",    type: "money" },
  { key: "budget_burn",           label: "Budget Burn",   type: "pct" },
  { key: "markup_pct",            label: "Markup %",      type: "pct" },
  { key: "margin_pct",            label: "Margin %",      type: "pct" },
  { key: "qbo_margin",            label: "Margin $ (QBO)", type: "money" },
  { key: "actual_markup_pct",     label: "Actual Markup %", type: "pct" },
  { key: "subs_pct",              label: "Subs %",        type: "pct" },
];

// Derived per-job metrics from the REAL QBO costs — computed once at load time.
// Only for jobs that actually have costs loaded; others stay null (blank).
// margin here = billed − QBO cost (a billed-basis margin-to-date, labeled as such).
function deriveMetrics(r) {
  const cost = r.costs_loaded, etc = r.estimated_total_costs, billed = r.billed_to_date,
        subs = r.sub_costs, contract = r.total_contract_price;
  const has = cost !== null && cost !== undefined;
  r.budget_burn   = has && etc ? cost / etc : null;
  r.qbo_margin    = has && billed != null ? billed - cost : null;
  r.qbo_margin_pct = (r.qbo_margin != null && billed) ? r.qbo_margin / billed : null;
  r.subs_pct      = has && cost && subs != null ? subs / cost : null;
  // PLANNED markup (on cost) vs PLANNED margin (on revenue) — never the same number.
  r.markup_pct    = (etc && contract != null) ? (contract - etc) / etc : null;
  r.margin_pct    = (contract && contract !== 0) ? ((contract - num(etc)) / contract) : null;
  // ACTUAL markup from real QBO cost (how much we marked the true cost up).
  r.actual_markup_pct = (has && cost) ? (billed - cost) / cost : null;
}

// ── Settings ──────────────────────────────────────────────────────────────
const LS_KEY = "proficient-ledger-settings-v1";
const DEFAULTS = {
  theme: "light", accent: "#3E7A5C", font: "system", fontSize: 15,   // light by default (owner 2026-09-02); a saved choice still wins   // 15px base (owner 2026-09-01: older Excel readers)
  density: "comfortable", width: "medium",
  widgets: { kpis: true, attention: true, ap: true, costs: true, margins: true, divisions: true, projects: true },
  columns: COLUMNS.filter(c => c.always || c.def).map(c => c.key),
};

// Lien state → urgency css class (most urgent first), for the AP watchlist.
const LIEN_CLASS = {
  "Notice PAST due": "past", "Notice due in ≤7d": "d7", "Notice due in ≤15d": "d15",
  "Notice due in ≤30d": "d30", "Notice Sent": "info", "Lien Filed": "info",
};

// Budget adherence — the ONE rule the whole dashboard flags "over budget" with.
// Flatwork (-FTW) budgets are a SOFT reference, not a strict target: the ops
// manager just sends a sub and charges by the labor it took (flatwork is simple
// next to slab), so the estimator's FTW budget is a starting point, not a
// must-hit number the way slab is — EXCEPT on a big flatwork job (~$15k+), which
// does need to hold its budget. Slab / CP / MFD stay strict.
const FTW_BUDGET_FLOOR = 15000;
function budgetCost(r) { return r.costs_loaded != null ? r.costs_loaded : r.costs_to_date; }
function isOverBudget(r) {
  const etc = num(r.estimated_total_costs);
  if (!(etc > 0) || num(budgetCost(r)) <= etc) return false;
  if (r.is_ftw) {                                  // flatwork: soft budget…
    const size = Math.max(num(r.total_contract_price), etc);
    if (size < FTW_BUDGET_FLOOR) return false;     // …unless it's a big flatwork job
  }
  return true;
}

// Generic, data-driven exposure rules (applied to whatever data is loaded).
const RULES = [
  { key: "underbilled", label: "Underbilled", warn: false,
    hint: "earned ahead of billed — could invoice",
    test: r => num(r.underbillings) > 0, amt: r => num(r.underbillings) },
  { key: "overbilled", label: "Overbilled", warn: false,
    hint: "billed ahead of earned",
    test: r => num(r.overbillings) > 0, amt: r => num(r.overbillings) },
  { key: "overbudget", label: "Over budget", warn: true,
    hint: "cost over ETC (flatwork budgets soft under $15k)",
    test: isOverBudget,
    amt: r => num(budgetCost(r)) - num(r.estimated_total_costs) },
  { key: "borrow", label: "Borrowing cash", warn: true,
    hint: "pure job borrow > 0",
    test: r => num(r.pure_job_borrow) > 0, amt: r => num(r.pure_job_borrow) },
];
let settings = loadSettings();

const LS_DEF = "proficient-ledger-defaults-v1";   // the user's saved "default view" baseline
function baseDefaults() {
  // What Reset returns to, and what a fresh browser opens with: the user's own
  // saved default (via "Set as default") if present, else the shipped DEFAULTS.
  try {
    const d = JSON.parse(localStorage.getItem(LS_DEF));
    if (d) return { ...structuredClone(DEFAULTS), ...d,
                    widgets: { ...DEFAULTS.widgets, ...(d.widgets || {}) },
                    columns: Array.isArray(d.columns) && d.columns.length ? d.columns : DEFAULTS.columns };
  } catch { /* ignore */ }
  return structuredClone(DEFAULTS);
}
function loadSettings() {
  const base = baseDefaults();
  try {
    const s = JSON.parse(localStorage.getItem(LS_KEY));
    if (!s) return base;
    return { ...base, ...s,
             widgets: { ...base.widgets, ...(s.widgets || {}) },
             columns: Array.isArray(s.columns) && s.columns.length ? s.columns : base.columns };
  } catch { return base; }
}
function saveSettings() { localStorage.setItem(LS_KEY, JSON.stringify(settings)); }

const FONTS = { system: "var(--font-system)", inter: "var(--font-inter)",
                serif: "var(--font-serif)", mono: "var(--font-mono)" };

function applySettings() {
  const root = document.documentElement;
  // theme
  const dark = settings.theme === "dark" ||
    (settings.theme === "auto" && matchMedia("(prefers-color-scheme: dark)").matches);
  root.setAttribute("data-theme", dark ? "dark" : "light");
  // typographic + layout vars
  root.style.setProperty("--font", FONTS[settings.font] || FONTS.system);
  root.style.setProperty("--fs", settings.fontSize + "px");
  root.style.setProperty("--accent", settings.accent);
  root.style.setProperty("--row-pad", settings.density === "compact" ? "5px 10px" : "10px 12px");
  root.style.setProperty("--maxw", settings.width === "boxed" ? "1180px"
    : settings.width === "medium" ? "1500px" : "100%");   // medium sits between boxed and full
  // widgets
  $("#widget-kpis").hidden      = !settings.widgets.kpis;
  $("#widget-attention").hidden = !settings.widgets.attention;
  $("#widget-costs").hidden     = !settings.widgets.costs;
  $("#widget-margins").hidden   = !settings.widgets.margins;
  $("#widget-divisions").hidden = !settings.widgets.divisions;
  $("#widget-projects").hidden  = !settings.widgets.projects;
}

// ── Formatting ────────────────────────────────────────────────────────────
const isNum = v => typeof v === "number" && !Number.isNaN(v);
function money(v) {
  if (v != null && v !== "" && !Number.isNaN(Number(v)) && Math.abs(Number(v)) > 0 && Math.abs(Number(v)) < 0.5) { const n = Number(v); const s = "$" + Math.abs(n).toFixed(2); return n < 0 ? `(${s})` : s; }   // cents, never a misleading $0
  if (v === null || v === undefined || v === "") return "—";
  const n = Number(v); if (Number.isNaN(n)) return "—";
  const s = "$" + Math.round(Math.abs(n)).toLocaleString();
  return n < 0 ? "(" + s + ")" : s;      // negatives like Excel: ($28,067), coloured red by .neg (owner 2026-09-01)
}
function pct(v) {
  if (v === null || v === undefined || v === "") return "—";
  const n = Number(v); if (Number.isNaN(n)) return "—";
  return (n * 100).toFixed(1) + "%";
}
function fmt(col, v) {
  if (col.type === "money") return money(v);
  if (col.type === "pct")   return pct(v);
  return (v === null || v === undefined || v === "") ? "—" : String(v);
}
// raw value for copy / CSV (numbers stay numeric so they paste clean into Excel)
function raw(col, v) {
  if (v === null || v === undefined) return "";
  return (col.type === "money" || col.type === "pct") && isNum(Number(v)) && v !== "" ? Number(v) : String(v);
}

// ── State ─────────────────────────────────────────────────────────────────
let ALL = [];
let AP = { summary: {}, lien_watch: [], liens: [], by_project: {}, bills: [] };
let BILLS = [];   // full ap_bill_line list for the Bill Tracker tab
let COST = { by_code: [], by_project_code: {}, by_project: {}, by_cost_type: [], by_vendor: [], loaded_total: 0 };
let DRAWS = { draws: [], total: 0 };
let SALES = { pipeline: [], by_rep: [], warm: [], customers: [], totals: {} };
let SUBLOC = { summary: null, divisions: {}, projects: [], open_by_project: {}, repays: [], events: [] };
let OI = { as_of: null, buckets: ["Current", "1-30", "31-60", "61-90", "90+"], invoices: [] };  // open AR invoices (aging tab)
let PAY = { payments: [], total_received: 0, count: 0, invoices_paid: 0 };   // received payments, each with the invoices it paid
let paymentsExpanded = new Set();   // payment ids expanded to show their invoices (default: all collapsed - scannable list)
let paymentsPeriodsExpanded = new Set();  // month/week bands expanded (default none = all collapsed, owner 2026-08-31)
let paymentsGroupBy = "month";      // 'none' | 'week' | 'month' - cash-in broken down by period (owner 2026-08-25)
const _MON3 = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
// Period key (sortable, newest-first) + a human label for a payment's date.
function payPeriod(dateStr, mode) {
  const m = String(dateStr || "").match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return { key: "0000", label: "No date" };
  const [, Y, Mo, D] = m;
  if (mode === "month") return { key: `${Y}-${Mo}`, label: `${_MON3[+Mo - 1]} ${Y}` };
  // week: roll back to Monday (local date math, no UTC drift)
  const d = new Date(+Y, +Mo - 1, +D); const dow = (d.getDay() + 6) % 7; d.setDate(d.getDate() - dow);
  const wk = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
  return { key: wk, label: `Week of ${fmtDateShort(wk)}` };
}
let costCollapsed = new Set();   // collapsed cost-type parents (default: all collapsed)
let drawsCollapsed = new Set();  // collapsed draw cards (default: all collapsed)
let drawsExpanded = new Set();   // draws whose bills are expanded in the table (default: none)
let drawVendorExpanded = new Set();   // (draw|vendor) groups expanded inside a draw's bills (default: none)
// Waiver tracking is PARKED (owner 2026-08-27: "i don't do waiver ... a feature for future PMs").
// The engine stays intact - /api/waiver, the waiver table, setWaiver() - flip this to true to
// bring the per-bill "Waiver in hand" column + caption back. It never gated a draw's stage/color.
const WAIVERS_ENABLED = false;

// ── Tabs (two-level grouped nav) ─────────────────────────────────────────────
// Parent groups on the top row; the active group's tabs on the second row. The whole
// structure lives here, so adding/moving a tab is a one-line edit (owner 2026-08-19).
const NAV_GROUPS = [   // order (owner 2026-08-31): Overview · Customer · Vendor · Financials · Health · IT · Console (Console last)
  { id: "overview",   label: "Overview",      tabs: ["overview"] },
  { id: "customers",  label: "Customer",      tabs: ["customers", "invoices", "draws", "payments", "sales"] },
  { id: "vendors",    label: "Vendor",        tabs: ["vendors", "bills", "accounting", "subloc"] },   // Pay Bills + Liens fold under Bills (owner 2026-09-02)
  { id: "financials", label: "Financials",    tabs: ["pnl", "wip", "wipreview", "costs"] },
  { id: "health",     label: "Health",        tabs: ["health"] },
  { id: "it",         label: "IT",            tabs: ["systems"] },
  { id: "console",    label: "Console",       tabs: ["console"] },
];
const TAB_LABELS = {
  overview: "Overview", health: "Health", pnl: "Project P&L", wip: "WIP report", wipreview: "WIP Review", costs: "Costs",
  customers: "Customer Center", invoices: "Invoices", draws: "Funding", payments: "Payments", sales: "Sales Outreach",
  vendors: "Vendor Center", bills: "Bills", paybills: "Pay Bills", accounting: "Audit", subloc: "Sub LOC", liens: "Liens",
  systems: "Systems", console: "Console",
};
const HIDDEN_TAB_GROUP = { paybills: "vendors", liens: "vendors" };   // pages without a sub-tab: opened from Bills
const groupOf = t => NAV_GROUPS.find(g => g.tabs.includes(t)) || NAV_GROUPS.find(g => g.id === HIDDEN_TAB_GROUP[t]) || NAV_GROUPS[0];
function buildGroupBar() {
  const bar = $("#groupbar"); if (!bar) return; bar.innerHTML = "";
  for (const g of NAV_GROUPS) {
    const b = document.createElement("button"); b.className = "tab"; b.dataset.group = g.id; b.textContent = g.label;
    b.onclick = () => setTab(g.tabs[0]);   // a group opens its landing (first) tab
    bar.appendChild(b);
  }
}
function buildSubTabs(g, active) {
  const bar = $("#subtabbar"); if (!bar) return;
  if (!g || g.tabs.length <= 1) { bar.hidden = true; bar.innerHTML = ""; return; }   // single-page group: no second row
  bar.hidden = false; bar.innerHTML = "";
  for (const t of g.tabs) {
    const b = document.createElement("button"); b.className = "subtab" + (t === active ? " active" : "");
    b.textContent = TAB_LABELS[t] || t; b.onclick = () => setTab(t);
    bar.appendChild(b);
  }
}
let activeTab = "overview";
function setTab(t) {
  activeTab = t;
  try { localStorage.setItem("proficient-ledger-tab", t); } catch { /* ignore */ }
  { const rv = $("#recordView"); if (rv) rv.hidden = true; }   // leaving a record view when a tab is picked
  $$(".tab-page").forEach(p => { p.hidden = p.dataset.tab !== t; });
  const g = groupOf(t);
  $$("#groupbar .tab").forEach(b => b.classList.toggle("active", b.dataset.group === g.id));
  buildSubTabs(g, t);
  if (t === "pnl") renderPnl();     // portfolio P&L is computed server-side, lazy-loaded
  if (t === "health") loadHealth(); // company health is computed server-side, lazy-loaded
  if (t === "wip") renderWip();
  if (t === "wipreview") loadWipReview();
  if (t === "console") renderConsole();
  if (t === "systems") loadSystems();
  if (t === "accounting") loadAccounting();
  if (t === "customers") renderCustomers();
  if (t === "payments") renderPayments();
  if (t === "paybills") renderPayBills();
  if (typeof _csClear === "function") _csClear();   // drop any cell selection when the tab changes
  window.scrollTo(0, 0);
}
let PNL = null;                       // cached /api/pnl/portfolio result (invalidated on reload)
let pnlSort = { key: "net", dir: 1 }; // net ascending = worst margin first
let pnlExpanded = new Set();          // P&L jobs expanded inline (instead of the side panel)
const nameOf = pn => (ALL.find(r => r.project_no === pn) || {}).project_name || "";
// A project is "active" if its WIP status is Active — OR blank, which only happens for
// MFD (Test-Master carries no STATUS column, so its jobs are active by construction).
// Matches _portfolio_pnl on the server so every "active" count agrees, MFD included.
const isActive = r => ["", "active"].includes((r.status || "").toLowerCase());
let meta = {};
let sortKey = "total_contract_price";
let sortDir = -1;   // -1 desc, 1 asc
let activeRule = null;   // key of a RULES entry currently filtering the table
let activeLien = null;   // lien stage currently filtering the Liens table (null = all)
let activeDrawStage = null; // draw stage currently filtering the Draws list (null = all)
let activeRep = null;    // rep whose activity drill is shown (null = auto: the outreach rep)

// ── Load ──────────────────────────────────────────────────────────────────
async function load(isAuto) {
  let data;
  try { data = await (await fetch(isAuto ? "/api/data" : "/api/data?light=1")).json(); }
  catch (e) { return showError("Could not reach the server: " + e); }
  if (data.error) return showError(data.error);
  $("#errorBanner").hidden = true;
  ALL = data.projects || [];
  ALL.forEach(deriveMetrics);
  AP = data.ap || { summary: {}, lien_watch: [], liens: [], by_project: {}, bills: [] };
  BILLS = AP.bills || [];
  COST = data.cost || { by_code: [], by_project_code: {}, by_project: {}, by_cost_type: [], by_vendor: [], loaded_total: 0 };
  DRAWS = data.draws || { draws: [], total: 0 };
  // heavy blobs: present on a full/auto load; on the light initial load they arrive via loadHeavy()
  if (data.sales !== undefined) SALES = data.sales || { pipeline: [], by_rep: [], warm: [], customers: [], totals: {} };
  if (data.sub_loc !== undefined) SUBLOC = data.sub_loc || { summary: null, divisions: {}, projects: [], open_by_project: {}, repays: [], events: [] };
  OI = data.open_invoices || { as_of: null, buckets: ["Current", "1-30", "31-60", "61-90", "90+"], invoices: [] };
  if (data.payments !== undefined) PAY = data.payments || { payments: [], total_received: 0, count: 0, invoices_paid: 0 };
  PNL = null;   // recompute the portfolio P&L on next open (data just changed)
  HEALTH = null;   // same for the Health tab - its sections derive from the same tables
  // Big-picture first: collapse everything by default; the user expands to zoom in.
  // On a live auto-refresh, preserve what the user has already expanded.
  if (!isAuto) {
    costCollapsed = new Set((COST.by_cost_type || []).map(g => g.parent));
    drawsCollapsed = new Set((DRAWS.draws || []).map(d => d.matched_invoice));
    // Bills open COLLAPSED by default (owner 2026-08-18) - scan vendor + amount, expand on demand.
    const bgrp = $("#billGroup") ? $("#billGroup").value : "vendor";
    billsCollapsed = bgrp === "none" ? new Set() : new Set((BILLS || []).map(b => billGroupKey(b, bgrp)));
  }
  meta = data.meta || {};
  { const v = $("#appVersion"); if (v) v.textContent = meta.version ? "v" + meta.version : ""; }
  // The meta line lives in the Data freshness head (the title bar it sat in is gone, 2026-09-01).
  $("#metaLine").textContent =
    `${meta.project_count} projects · WIP report ${meta.report_date ? fmtDate(meta.report_date) : "–"}` +
    (meta.loaded_at ? ` · ledger loaded ${fmtDate(meta.loaded_at, true)}` : "");
  buildFilterOptions();
  render();
  _renderLazyTab(activeTab);   // wip/payments/paybills read main-load globals but aren't in render();
                               // re-dispatch the active one now that data is in (fixes a fresh refresh on it)
  if (!isAuto) loadHeavy();    // phase 2: pull the deferred heavy tab blobs in the background
}

// Phase 2 of the FIRST load: the heavy tab blobs (bills ~2.7 MB, sub_loc, payments, sales)
// are fetched in the background so first paint isn't blocked on the whole ~5 MB. When they
// land, fill the globals and re-render (render() builds every tab's DOM, incl. hidden ones).
async function loadHeavy() {
  let h;
  try { h = await (await fetch("/api/data?heavy=1")).json(); } catch (e) { return; }
  if (!h || h.error) return;
  AP.bills = h.ap_bills || [];
  BILLS = AP.bills;
  if (h.sub_loc) SUBLOC = h.sub_loc;
  if (h.payments) PAY = h.payments;
  if (h.sales) SALES = h.sales;
  const bgrp = $("#billGroup") ? $("#billGroup").value : "vendor";   // bills open collapsed by default
  billsCollapsed = bgrp === "none" ? new Set() : new Set((BILLS || []).map(b => billGroupKey(b, bgrp)));
  buildFilterOptions();
  render();
  _renderLazyTab(activeTab);
}
// Lazy tabs dispatched by setTab (not render()) that read the /api/data globals. pnl/systems/console
// fetch their OWN data on open, so they self-refresh; these three read ALL / PAY / BILLS synchronously.
function _renderLazyTab(t) {
  const map = { wip: renderWip, payments: renderPayments, paybills: renderPayBills,
                health: () => loadHealth(true) };
  if (map[t]) map[t]();
}
function showError(msg) {
  const b = $("#errorBanner"); b.hidden = false; b.textContent = msg;
  $("#metaLine").textContent = "not loaded";
}

// Manual "Refresh" — re-reads the ledger DB and re-renders, WITH feedback so it's
// obvious it did something (the silent 90s auto-refresh does the same in the
// background). It does NOT re-pull QBO/Excel — that's a sync (the loaders); the
// "Data freshness" strip flags when a sync is worth running.
async function manualRefresh() {
  const btn = $("#btnRefresh"); const orig = btn.textContent;
  btn.disabled = true; btn.textContent = "Refreshing…";
  try { await load(true); }                 // load(true) = keep what you've expanded
  finally { btn.disabled = false; btn.textContent = orig; }
  toast(meta.loaded_at ? `Refreshed · ledger loaded ${fmtDate(meta.loaded_at, true)}` : "Refreshed");
}

// ── In-app runs: the Console (and My-view Resync) run a pipeline via the sync engine,
// with a live progress bar. Pauses the silent auto-refresh while running (loaders
// drop/rebuild tables). Producer steps + QBO costs prompt Touch ID on this Mac.
let syncing = false;
let runningPipeline = null;   // key of the sync currently running
let syncQueue = [];           // [{ key, els }] confirmed syncs waiting - they run ONE AT A TIME
                              // (concurrent QBO pulls + ledger DELETE/INSERT would corrupt each other).
// Run a pipeline key ('reload' = safe loaders-only default, 'all' = full chain incl
// producers, 'ar'/'ap'/'costs'/'crm'/'wip', 'wip-draft'), driving the given progress
// elements. `els` = { btn, prog, fill, step }. If a sync is already running, this QUEUES it.
async function runPipeline(pipeline, confirmMsg, els) {
  if (confirmMsg && !confirm(confirmMsg)) return;
  if (syncing) {
    if (pipeline === runningPipeline || syncQueue.some(q => q.key === pipeline)) { toast("That sync is already running or queued."); return; }
    syncQueue.push({ key: pipeline, els });
    toast(`Queued - runs when the current sync finishes (${syncQueue.length} waiting).`);
    if (activeTab === "console") renderConsole();
    return;
  }
  _startPipeline(pipeline, els);
}
async function _startPipeline(pipeline, els) {
  let res;
  try { res = await (await fetch("/api/sync", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ pipeline, confirm: true }) })).json(); }
  catch (e) { toast("Could not start: " + e); _drainQueue(); return; }
  if (res.error) { toast(res.error); _drainQueue(); return; }
  syncing = true; runningPipeline = pipeline; if (els.btn) els.btn.disabled = true;
  els.prog.hidden = false; els.fill.classList.remove("err"); els.fill.style.width = "0%";
  if (activeTab === "console") renderConsole();
  pollSync(res.steps || [], els, pipeline);
}
// Start the next queued sync once the current one is fully done (sequential).
function _drainQueue() {
  if (syncing || !syncQueue.length) return;
  const next = syncQueue.shift();
  const label = (PIPELINES || []).find(p => p.key === next.key)?.label || next.key;
  toast(`Starting queued sync: ${label}`);
  _startPipeline(next.key, next.els);
}

async function startResync() {
  runPipeline("reload",
    "Reload all data now?\n\nRuns EVERY loader into the ledger - WIP, Costs (QBO), Bill Tracker, Invoices, Payments (QBO), Customers, Sub LOC - so the Project P&L and Payments come back current. Read-only on the sources; a couple of QBO steps may prompt Touch ID; takes a minute or two.",
    { btn: $("#btnResync"), prog: $("#syncProgress"), fill: $("#syncBarFill"), step: $("#syncStep") });
}

// AP + AR back to back - both pull the same QBO info and AR's aging reads AP's Bill Tracker output,
// so they belong together (owner 2026-08-25). AP runs first (the run-order rule), AR queued behind it.
async function runApAr() {
  if (!confirm("Sync AP + AR now?\n\nRuns the bill tracker (AP) FIRST, then the invoice sync (AR) - both pull from QuickBooks and prompt Touch ID, and AR's aging reads AP's output. Loads both into the ledger.")) return;
  const els = _consoleEls();
  if (syncing) {
    if (runningPipeline !== "ap" && !syncQueue.some(q => q.key === "ap")) syncQueue.push({ key: "ap", els });
    if (runningPipeline !== "ar" && !syncQueue.some(q => q.key === "ar")) syncQueue.push({ key: "ar", els });
    toast(`Queued AP + AR (${syncQueue.length} waiting).`);
    if (activeTab === "console") renderConsole();
  } else {
    syncQueue.push({ key: "ar", els });   // AR waits behind AP
    _startPipeline("ap", els);
  }
}

function pollSync(steps, els, pipeline) {
  const { btn, prog, fill, step } = els;
  const total = steps.length || 1;
  const plabel = (PIPELINES || []).find(p => p.key === pipeline)?.label || pipeline || "";
  let fails = 0;
  const tick = () => fetch("/api/sync/status").then(r => r.json()).then(s => {
    fails = 0;
    const done = (s.steps || []).filter(x => x.state === "done").length;
    const cur = (s.steps || [])[s.current];
    fill.style.width = Math.round(done / total * 100) + "%";
    if (s.state === "running") {
      const q = syncQueue.length ? ` · ${syncQueue.length} queued` : "";
      step.textContent = `${plabel ? plabel + " - " : ""}${cur ? cur.label : "..."} - step ${Math.min(done + 1, total)} of ${total}${s.elapsed ? ` - ${s.elapsed}s` : ""}${q}`;
      setTimeout(tick, 1500);
    } else if (s.state === "done") {
      fill.style.width = "100%"; step.textContent = "Done - reloading the app...";
      finishSync(els, "Done - data refreshed.", true);
    } else if (s.state === "error") {
      const bad = (s.steps || []).find(x => x.state === "error");
      fill.classList.add("err");
      step.textContent = `Failed at: ${bad ? bad.label : "a step"} - see the log (~/Library/Logs/Proficient/ledger-sync).`;
      finishSync(els, "Run failed - " + (bad ? bad.label : ""), false);
    } else {   // idle mid-poll: the app restarted; a step may still be running in the background
      finishSync(els, "Status lost (did the app restart?) - check Data freshness.", false);
    }
  }).catch(() => {   // server unreachable - cap the retries so the button can't hang disabled forever
    if (++fails >= 5) { finishSync(els, "Lost contact with the app - check Data freshness.", false); return; }
    setTimeout(tick, 2500);
  });
  setTimeout(tick, 800);
}

async function finishSync(els, msg, reload) {
  syncing = false; runningPipeline = null; if (els.btn) els.btn.disabled = false;
  if (reload) { try { await load(true); if (typeof renderConsole === "function" && activeTab === "console") renderConsole(); } catch { /* ignore */ } }
  if (msg) toast(msg);
  // Keep the bar up if another run is queued (it starts right away); else tidy it after a beat.
  if (!syncQueue.length && !/(failed|lost|Lost)/.test(msg)) setTimeout(() => { els.prog.hidden = true; els.fill.style.width = "0%"; }, 2600);
  _drainQueue();
}

// ── Console tab: the control plane. Lists each pipeline (from /api/pipelines) with its
// steps, last-run, and a Run button (a pipeline's Run also fires its real producer).
let PIPELINES = null;
const _consoleEls = () => ({ prog: $("#consoleProgress"), fill: $("#consoleBarFill"), step: $("#consoleStep") });
// Plain-language "what this sync does" per pipeline (owner 2026-08-25): what it grabs, where it
// writes, and which tabs it feeds. Two hops for AP/AR (QBO -> the working system -> the ledger).
const PIPELINE_DESC = {
  ap: "Grabs every vendor bill + purchase from QuickBooks and matches each to the GC draw that authorizes paying it -> writes Bill Tracker.xlsx (OneDrive) -> loads the bills + lien clock into the ledger. Feeds Bills · Pay Bills · Liens.",
  ar: "Grabs your open AR invoices (the draws) from QuickBooks -> updates the Notion Invoice Tracker + the AR Aging Excel (sweeps paid ones to Paid, posts MFD pay events to Teams) -> loads them into the ledger. Feeds Invoices · Draws · Customers · Payments.",
  wip: "Reads the WIP master's Test tabs (SharePoint Excel) -> loads the project list + WIP snapshot (contract, % complete, over/under-billing) into the ledger. Feeds Overview · WIP report · Project P&L.",
  costs: "Pulls the last 90 days of job costs from QuickBooks (incl. subs), keyed by cost code, into the ledger. Prompts Touch ID. Feeds the Costs tab + the Project P&L margins.",
  payments: "Pulls QuickBooks Payments (money IN) - who paid, which invoices/draws it cleared, the payment method - into the ledger, a rolling 12-month year. Feeds the Payments tab.",
  crm: "Pulls the Notion Customer List into the ledger - leads/clients + the per-rep outreach touch log. Feeds Customer Center · Sales Outreach.",
  subloc: "Pulls QBO payments to subs to model each sub's float (line-item, actual pay dates, chronological FIFO). Feeds the Sub LOC tab.",
  pnl: "Regenerates the per-project P&L workbooks for every ACTIVE job of one division (OneDrive PROJECT P&Ls; CP lands in the job's Synology folder). Reads QBO + the takeoff budget and writes Excel - it does NOT change the ledger. One division at a time; a full division takes a while, so pick the one you need.",
};
async function renderConsole() {
  const box = $("#consoleList"); if (!box) return;
  if (!PIPELINES) {
    try { PIPELINES = (await (await fetch("/api/pipelines")).json()).pipelines || []; }
    catch { box.textContent = "Console unavailable."; return; }
  }
  const fr = meta.freshness || { sources: {}, ledger: {} };
  // "last ran" per pipeline: prefer when the SOURCE last synced (file mtime), but ALWAYS
  // fall back to when the ledger last LOADED that feed (loaded_at) so a card is never blank
  // just because the source file isn't on this machine (the AR mirror often isn't) - that was
  // the "AP showed, AR didn't" bug. costs/crm/subloc pull straight from QBO/Notion (no file).
  const src = fr.sources || {}, led = fr.ledger || {};
  const lastRun = {
    ap: src["sync-ap"] || led["AP (Bill Tracker)"],
    ar: src["sync-ar"] || led["AR (invoices)"],
    wip: src["WIP master"] || led["WIP"],
    costs: led["Costs (QBO)"],
    payments: led["Payments"],
    crm: led["CRM (customers)"],
    subloc: led["Sub LOC"],
  };
  box.innerHTML = "";
  for (const p of PIPELINES) {
    const card = document.createElement("div"); card.className = "pl-card";
    const head = document.createElement("div"); head.className = "pl-head";
    const nm = document.createElement("span"); nm.className = "pl-name"; nm.textContent = p.label; head.appendChild(nm);
    const lr = lastRun[p.key];
    const when = document.createElement("span"); when.className = "pl-when";
    when.textContent = lr ? `last ${timeAgo(lr)}` : ""; if (lr) when.title = fmtDate(lr, true);
    head.appendChild(when);
    card.appendChild(head);
    if (PIPELINE_DESC[p.key]) { const d = document.createElement("p"); d.className = "pl-desc"; d.textContent = PIPELINE_DESC[p.key]; card.appendChild(d); }
    const steps = document.createElement("div"); steps.className = "pl-steps";
    for (const s of p.steps) {
      const chip = document.createElement("span"); chip.className = "pl-step" + (s.side ? " producer" : "");
      chip.textContent = s.label + (s.side ? " · producer" : ""); steps.appendChild(chip);
    }
    card.appendChild(steps);
    const acts = document.createElement("div"); acts.className = "pl-acts";
    const actionsOnly = (p.actions || []).length > 0 && !p.steps.length;
    const runBtn = document.createElement("button"); runBtn.className = "btn small";
    const sides = p.steps.filter(s => s.side).map(s => s.label);
    const msg = sides.length
      ? `Run the ${p.label} pipeline?\n\nThis fires a REAL sync (${sides.join(", ")}) - writes to the source (Notion / Teams / Excel) and prompts Touch ID - then loads it into the ledger.`
      : `Run the ${p.label} loader?\n\nReads the current source into the ledger (read-only on the source).`;
    if (p.key === runningPipeline) { runBtn.textContent = "Running…"; runBtn.disabled = true; card.classList.add("pl-running"); }
    else if (syncQueue.some(q => q.key === p.key)) {   // click a queued card to drop it from the queue
      runBtn.textContent = "Queued ✕"; runBtn.classList.add("subtle"); card.classList.add("pl-queued");
      runBtn.onclick = () => { syncQueue = syncQueue.filter(q => q.key !== p.key); toast("Removed from the queue"); renderConsole(); };
    } else { runBtn.textContent = "Run"; runBtn.onclick = () => runPipeline(p.key, msg, { ..._consoleEls(), btn: runBtn }); }
    if (!actionsOnly) acts.appendChild(runBtn);
    // Pipelines that expose per-variant actions (P&L by division) render one
    // button each INSTEAD of a generic Run - "pnl" alone resolves to no steps.
    for (const a of (p.actions || [])) {
      const aBtn = document.createElement("button"); aBtn.className = "btn small";
      if (a.key === runningPipeline) { aBtn.textContent = `${a.label}…`; aBtn.disabled = true; card.classList.add("pl-running"); }
      else if (syncQueue.some(q => q.key === a.key)) {
        aBtn.textContent = `${a.label} ✕`; aBtn.classList.add("subtle"); card.classList.add("pl-queued");
        aBtn.onclick = () => { syncQueue = syncQueue.filter(q => q.key !== a.key); toast("Removed from the queue"); renderConsole(); };
      } else {
        aBtn.textContent = a.label;
        aBtn.onclick = () => runPipeline(a.key,
          `Regenerate the ${a.label} P&L workbooks?\n\nRuns every ACTIVE ${a.label.replace("Active ", "")} job against QuickBooks and rewrites its workbook. Writes Excel only - the ledger is untouched. This can take several minutes.`,
          { ..._consoleEls(), btn: aBtn });
      }
      acts.appendChild(aBtn);
    }
    if (p.draft) {
      const dBtn = document.createElement("button"); dBtn.className = "btn small subtle"; dBtn.textContent = p.draft.label;
      dBtn.onclick = () => runPipeline("wip-draft",
        `${p.draft.label}?\n\nGenerates the DRAFT WIP (Test tabs) for PMs to review - it does NOT implement anything into the live report. Reads Excel + QBO; prompts Touch ID.`,
        { ..._consoleEls(), btn: dBtn });
      acts.appendChild(dBtn);
    }
    card.appendChild(acts);
    box.appendChild(card);
  }
}

// ── Filters ───────────────────────────────────────────────────────────────
function buildFilterOptions() {
  fillSelect("#fDivision", uniq(ALL.map(r => r.division)));
  fillSelect("#fStatus",   uniq(ALL.map(r => r.status)));
  fillSelect("#fCategory", uniq(ALL.map(r => r.rp_category)));
  if ($("#salesStage")) fillSelect("#salesStage", uniq((SALES.customers || []).map(c => c.sales_status)));
  if ($("#salesDivision")) fillSelect("#salesDivision", uniq((SALES.customers || []).map(c => c.division)));
}
const uniq = arr => [...new Set(arr.filter(Boolean))].sort();
function fillSelect(sel, values) {
  const el = $(sel); const keep = el.firstElementChild;
  el.innerHTML = ""; el.appendChild(keep);
  for (const v of values) { const o = document.createElement("option"); o.value = v; o.textContent = v; el.appendChild(o); }
}
function currentFilters() {
  return {
    q: $("#search").value.trim().toLowerCase(),
    division: $("#fDivision").value,
    status: $("#fStatus").value,
    category: $("#fCategory").value,
    activeOnly: $("#fActive").checked,
  };
}
function filtered() {
  const f = currentFilters();
  const rule = activeRule ? RULES.find(x => x.key === activeRule) : null;
  let rows = ALL.filter(r => {
    if (rule && !rule.test(r)) return false;
    if (f.division && r.division !== f.division) return false;
    if (f.status && r.status !== f.status) return false;
    if (f.category && r.rp_category !== f.category) return false;
    if (f.activeOnly && !isActive(r)) return false;
    if (f.q) {
      const hay = [r.project_no, r.project_name, r.builder_or_gc, r.client].filter(Boolean).join(" ").toLowerCase();
      if (!hay.includes(f.q)) return false;
    }
    return true;
  });
  const col = COLUMNS.find(c => c.key === sortKey) || COLUMNS[0];
  rows.sort((a, b) => {
    const av = a[sortKey], bv = b[sortKey];
    const na = av === null || av === undefined || av === "";
    const nb = bv === null || bv === undefined || bv === "";
    if (na && nb) return 0;
    if (na) return 1;   // nulls ALWAYS last, regardless of sort direction
    if (nb) return -1;
    return cmpVal(av, bv, col.type) * sortDir;
  });
  return rows;
}
function cmpVal(a, b, type) {
  if (type === "money" || type === "pct") return Number(a) - Number(b);
  return String(a).localeCompare(String(b));
}

// ── Render ────────────────────────────────────────────────────────────────
function visibleColumns() {
  return COLUMNS.filter(c => c.always || settings.columns.includes(c.key));
}
function render() {
  renderHome();
  renderKPIs(); renderAttention(); renderCosts(); renderMargins(); renderDivisions();
  renderProjects(); renderLiens(); renderVendors(); renderFunding(); renderBills(); renderOpenInvoices(); renderSubLoc(); renderSales(); renderCustomers();
}

function timeAgo(iso) {
  if (!iso) return "not found";
  const t = Date.parse(iso.length <= 16 ? iso + ":00" : iso);
  if (isNaN(t)) return iso;
  const mins = Math.floor((Date.now() - t) / 60000);
  if (mins < 1) return "just now";
  if (mins < 60) return mins + "m ago";
  const hrs = Math.floor(mins / 60); if (hrs < 24) return hrs + "h ago";
  return Math.floor(hrs / 24) + "d ago";
}

// The owner's date format (owner 2026-08-21: "mm/dd/yyyy for all formatting everywhere").
// Numeric month-day-year, zero-padded, 4-digit year - NEVER year-first. Add 12h time only
// when asked. This is THE date format for the whole dashboard (see also fmtDateShort).
function fmtDate(v, withTime) {
  if (!v) return "–";
  const m = String(v).trim().match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2}))?/);
  if (!m) return String(v);                      // not an ISO date → leave as-is
  const [, Y, Mo, D, H, Mi] = m;
  let out = `${Mo}/${D}/${Y}`;                    // mm/dd/yyyy, already zero-padded by the ISO source
  if (withTime && H != null) {
    let hr = +H; const ap = hr >= 12 ? "PM" : "AM"; hr = hr % 12 || 12;
    out += ` · ${hr}:${Mi} ${ap}`;
  }
  return out;
}

// Hours elapsed since `thenMs`, counting only Mon–Fri (weekends don't age the
// data — nobody syncs on the weekend, so a Friday load isn't "stale" on Monday).
// Steps day-by-day, adding only weekday slices. Used for the sync recommendation.
function businessHoursSince(thenMs, nowMs) {
  if (!(thenMs > 0) || nowMs <= thenMs) return 0;
  let total = 0, cur = thenMs;
  while (cur < nowMs) {
    const d = new Date(cur);
    const nextMidnight = new Date(d.getFullYear(), d.getMonth(), d.getDate() + 1).getTime();
    const sliceEnd = Math.min(nextMidnight, nowMs);
    const dow = d.getDay();                       // 0 Sun … 6 Sat
    if (dow !== 0 && dow !== 6) total += sliceEnd - cur;
    cur = sliceEnd;
  }
  return total / 3600e3;
}
const STALE_BUSINESS_H = 48;                       // > 2 business days → recommend a sync

function renderSyncPill() {
  const pill = $("#syncPill"), txt = $("#syncPillText"); if (!pill || !txt) return;
  const fr = meta.freshness || { ledger: {}, sources: {} }, S = fr.sources || {}, L = fr.ledger || {};
  const feeds = [["AP bills", S["sync-ap"]], ["AR", S["sync-ar"]], ["WIP master", S["WIP master"]], ["Costs", L["Costs (QBO)"]], ["Invoices", L["AR (invoices)"]],
                 ["Payments", L["Payments"]], ["Customers", L["CRM (customers)"]], ["Sub LOC", L["Sub LOC"]], ["Health", L["Health (QBO)"]]];
  let newest = null, stale = [];
  for (const [n, w] of feeds) { if (!w) { stale.push(n + " never"); continue; } const t = Date.parse(w.length <= 16 ? w + ":00" : w); if (isNaN(t)) continue;
    if (!newest || t > newest) newest = t; if (businessHoursSince(t, Date.now()) > STALE_BUSINESS_H) stale.push(`${n} ${timeAgo(w)}`); }
  if (typeof syncing !== "undefined" && syncing) { pill.className = "sync-pill busy"; txt.textContent = "Syncing…"; pill.title = "A sync is running - see the progress on Overview"; return; }
  pill.className = "sync-pill " + (stale.length ? "stale" : "ok");
  txt.textContent = newest ? `Synced ${fmtDate(new Date(newest - new Date().getTimezoneOffset() * 60000).toISOString().slice(0, 19), true)}` : "Not synced";
  if (stale.length) txt.textContent += ` · ${stale.length} stale`;
  pill.title = feeds.map(([n, w]) => `${n}: ${w ? fmtDate(w, true) + " (" + timeAgo(w) + ")" : "never"}`).join("\n") + (stale.length ? "\n\nStale (over 48 business hours): " + stale.join(", ") : "\n\nEvery feed is within 48 business hours");
}
function renderHome() {
  renderSyncPill();
  // ── data freshness ──
  const fr = meta.freshness || { ledger: {}, sources: {} };
  // Every feed _freshness() computes - the three source files and each ledger load - so AR, payments,
  // CRM, Sub LOC and Health staleness are visible too (owner 2026-09-02: "4 of 8 feeds shown").
  const S = fr.sources || {}, L = fr.ledger || {};
  const items = [
    ["sync-ap (AP bills)", S["sync-ap"]],
    ["sync-ar (AR)", S["sync-ar"]],
    ["WIP master", S["WIP master"]],
    ["Costs loaded (QBO)", L["Costs (QBO)"]],
    ["Invoices loaded (AR)", L["AR (invoices)"]],
    ["Payments loaded (QBO)", L["Payments"]],
    ["Customers loaded (Notion)", L["CRM (customers)"]],
    ["Sub LOC computed", L["Sub LOC"]],
    ["Health pulled (QBO)", L["Health (QBO)"]],
  ];
  const box = $("#homeFresh"); box.innerHTML = "";
  let needSync = 0;
  for (const [label, when] of items) {
    const el = document.createElement("div"); el.className = "fresh";
    let stale = false;
    if (when) {
      const t = Date.parse(when.length <= 16 ? when + ":00" : when);
      if (!isNaN(t) && businessHoursSince(t, Date.now()) > STALE_BUSINESS_H) stale = true;
    }
    if (stale) { el.classList.add("stale"); needSync++; }
    el.innerHTML = `<div class="f-label"></div><div class="f-when"></div><div class="f-ago"></div>`;
    el.querySelector(".f-label").textContent = label;
    el.querySelector(".f-when").textContent = when ? fmtDate(when, true) : "never";
    el.querySelector(".f-ago").textContent = when ? timeAgo(when) : "not loaded yet";
    if (stale) {
      const b = document.createElement("div"); b.className = "sync-rec"; b.textContent = "⟳ Sync recommended";
      el.appendChild(b);
    }
    box.appendChild(el);
  }
  const freshNote = $("#homeFreshNote");
  if (freshNote) freshNote.textContent = needSync
    ? `- ${needSync} recommended to sync (over 48h, weekends aside)` : "";   // each card says its own as-of; no blanket "all current" (owner 2026-09-01)
  // ── action items (click → jump to the work) ──
  const pastDue = (AP.lien_watch || []).filter(r => r.lien_status === "Notice PAST due").length;
  const dueSoon = (AP.lien_watch || []).filter(r => r.lien_status === "Notice due in ≤7d").length;
  const readyDraws = (DRAWS.draws || []).filter(d => d.stage === "Ready to turn in").length;
  const overB = ALL.filter(isOverBudget).length;
  const underB = ALL.filter(r => num(r.underbillings) > 0).length;
  const goRule = (key) => { setTab("overview"); activeRule = key; renderAttention(); renderProjects(); $("#btnClearRule").hidden = false; window.scrollTo(0, 0); };
  const acts = [
    ["Bills past lien date", pastDue, true, () => setTab("liens"),
      "Unpaid bills YOU owe (AP) whose vendor/supplier lien-notice deadline has passed — they can lien the project. Pay to clear. This is money OUT, not your AR."],
    ["Lien date in ≤7d", dueSoon, true, () => setTab("liens"),
      "Unpaid bills you owe that are within 7 days of the vendor's lien-notice deadline."],
    ["Draws ready to turn in", readyDraws, false, () => setTab("draws"),
      "Draws with every bill paid — turn in to unlock the next draw."],
    ["Over budget", overB, true, () => goRule("overbudget"),
      "Jobs where cost-to-date has passed the ETC budget."],
    ["Underbilled (can invoice)", underB, false, () => goRule("underbilled"),
      "Jobs earning ahead of what's been billed — you could invoice more."],
  ];
  const ar = $("#homeActions"); ar.innerHTML = "";
  for (const [label, n, warn, go, tip] of acts) {
    const el = document.createElement("div"); el.className = "action" + (warn && n ? " warn" : "") + (n ? "" : " none");
    el.innerHTML = `<span class="a-n"></span><span class="a-lab"></span>`;
    el.querySelector(".a-n").textContent = n;
    el.querySelector(".a-lab").textContent = label;
    if (tip) el.title = tip;
    if (n) el.onclick = go;
    ar.appendChild(el);
  }
  // ── working on (active projects) - only if the section exists. It was merged into the
  //    Overview's Projects section (which has search + filters), so on the merged tab this
  //    block simply no-ops. ──
  const sel = $("#homeDivision");
  if (!sel || !$("#homeWorkingTable")) return;
  if (sel && sel.options.length <= 1) for (const d of uniq(ALL.map(r => r.division))) { const o = document.createElement("option"); o.value = d; o.textContent = d; sel.appendChild(o); }
  const div = sel ? sel.value : "";
  const active = ALL.filter(r => isActive(r) && (!div || r.division === div))
    .sort((a, b) => num(b.total_contract_price) - num(a.total_contract_price));
  $("#homeWorkingNote").textContent = `(${active.length} active)`;
  const cols = [["Project", "left"], ["Division", "left"], ["Name", "left"], ["Contract", "right"], ["% Complete", "right"], ["Costs", "right"]];
  const thead = $("#homeWorkingTable thead"), tbody = $("#homeWorkingTable tbody");
  thead.innerHTML = ""; tbody.innerHTML = "";
  const htr = document.createElement("tr");
  for (const [c, al] of cols) { const th = document.createElement("th"); if (al === "left") th.className = "left"; th.textContent = c; htr.appendChild(th); }
  thead.appendChild(htr);
  for (const r of active.slice(0, 60)) {
    const tr = document.createElement("tr");
    tr.onclick = (e) => { if (!e.target.closest(".cell")) openDetail(r); };
    tr.appendChild(leftText(r.project_no));
    tr.appendChild(leftText(r.division || ""));
    tr.appendChild(leftText(r.project_name || ""));
    const cc = document.createElement("td"); cc.appendChild(cellFor({ key: "total_contract_price", type: "money" }, r.total_contract_price)); tr.appendChild(cc);
    const pc = document.createElement("td"); pc.appendChild(cellFor({ key: "percent_complete", type: "pct" }, r.percent_complete)); tr.appendChild(pc);
    const co = document.createElement("td"); co.appendChild(cellFor({ key: "costs_to_date", type: "money" }, r.costs_to_date)); tr.appendChild(co);
    tbody.appendChild(tr);
  }
}

const DRAW_STAGE_CLASS = {
  "Fund in — pay vendors": "d7",
  "Awaiting GC funding": "info",
  "Ready to turn in": "d7",   // vendors paid, GC still owes -> amber (collect)
  "All paid": "ready",        // GC paid + vendors paid -> green (done)
};
// Clearer, direction-explicit pill text (who paid whom). Display only — the internal
// stage keys above are unchanged (they're matched in several places).
const DRAW_STAGE_LABEL = {
  "Fund in — pay vendors": "GC funded → pay vendors",
  "Awaiting GC funding": "Awaiting GC funding",
  "Ready to turn in": "Vendors paid → collect the rest from the GC",
  "All paid": "All paid - the GC paid you and the vendors are paid",
};
// Company-scoped QBO deep link. The BARE app/invoice?txnId= form resolves the txn in
// whatever Intuit company the browser is on - with more than one company logged in it
// opens the WRONG company's txn. Routing through /app/login with deeplinkcompanyid pins
// the company first (Intuit's own "copy link" form). Falls back to bare until the realm
// loads (populated in the DB by load_costs; never printed). kind = 'invoice' | 'bill'.
function qboUrl(kind, txnId) {
  if (!txnId) return null;
  const realm = meta && meta.qbo_realm;
  if (realm) return `https://qbo.intuit.com/app/login?pagereq=${encodeURIComponent(kind + "?txnId=" + txnId)}&deeplinkcompanyid=${encodeURIComponent(realm)}`;
  return `https://qbo.intuit.com/app/${kind}?txnId=${encodeURIComponent(txnId)}`;
}
const qboInvoiceUrl = id => qboUrl("invoice", id);
// ap_bill_line.qbo_link holds a BARE bill URL (from the Bill Tracker's own hyperlink);
// pull the txnId out and rebuild it company-scoped.
function qboBillHref(link) {
  if (!link) return null;
  const m = String(link).match(/txnId=(\d+)/i);
  return m ? qboUrl("bill", m[1]) : link;
}
// A left-aligned <td> whose text opens a QBO deep link in a new tab when `url` is
// set; a plain cell otherwise. Used for bill/invoice numbers across the tables.
// Ref # cell: the NUMBER copies to the clipboard on click (owner 2026-08-28: "sometimes i just want
// to copy the ref# and not take me to qbo"); the trailing ↗ is the QBO link.
function qboLinkCell(text, url, title) {
  const td = document.createElement("td"); td.className = "left";
  const label = text || "—";
  if (text) {
    const s = document.createElement("span"); s.className = "refcopy"; s.textContent = label;
    s.title = "Click to copy " + label;
    s.onclick = e => { e.stopPropagation(); copy(label); toast("Copied " + label); };
    td.appendChild(s);
    if (url) {
      const a = document.createElement("a"); a.href = url; a.target = "_blank"; a.rel = "noopener";
      a.className = "qbo-ico"; a.textContent = "qb"; a.title = title || "Open in QuickBooks";
      a.onclick = e => e.stopPropagation(); td.appendChild(a);
    }
  } else { td.textContent = label; }
  return td;
}
// Short pill text (keeps the table narrow); the full "who paid whom" is the tooltip.
const DRAW_STAGE_SHORT = {
  "Fund in — pay vendors": "Pay vendors",
  "Awaiting GC funding": "Awaiting GC",
  "Ready to turn in": "Collect from GC",
  "All paid": "All paid",
};
// Draws filters: the SAME searchable multi-select used on Pay Bills / Invoices (owner
// 2026-08-27: "consistent throughout the ledger ... a selectable box that drills down the
// more you type, with select/deselect all"). Vendor is multi-valued (a draw spans many
// bills), so it carries its own pass; the rest key one value per draw.
const DIV_LABEL = { MFD: "Multi Family", CP: "Commercial", RP: "Residential" };
function _drawDiv(d) { const m = String(d.project_no || "").toUpperCase().match(/^(MFD|CP|RP)/); return m ? m[1] : ""; }
const drawMSel = {};
let _drawMSelSig = null;
const DRAW_MSEL = [
  { id: "dfClient", all: "All clients",   get: d => d.customer || "",   search: true, lbl: v => v || "(no client)" },
  { id: "dfProj",   all: "All projects",  get: d => d.project_no || "", search: true, lbl: v => v || "(none)" },
  { id: "dfVendor", all: "All vendors",   get: v => v, vendors: true,   search: true, lbl: v => v || "(none)" },
  { id: "dfInv",    all: "All invoices",  get: d => d.invoice_no || "", search: true, lbl: v => v || "(none)" },
  { id: "dfDiv",    all: "All divisions", get: d => _drawDiv(d),                      lbl: v => DIV_LABEL[v] || v || "(none)" },
];
function buildDrawFilters() {
  const draws = DRAWS.draws || [];
  const sig = String(draws.length);
  if (sig === _drawMSelSig && $("#dfClientMenu") && $("#dfClientMenu").querySelector(".msel-opt")) return;
  _drawMSelSig = sig;
  const vendors = [...new Set(draws.flatMap(d => (d.bills || []).map(b => b.vendor || "").filter(Boolean)))];
  for (const cfg of DRAW_MSEL) buildMSel(cfg, cfg.vendors ? vendors : draws, drawMSel, renderDraws);
}
function drawMselPasses(d) {
  for (const cfg of DRAW_MSEL) {
    const s = drawMSel[cfg.id]; if (!s || !s.size) continue;
    if (cfg.vendors) { if (!(d.bills || []).some(b => s.has(b.vendor || ""))) return false; }
    else if (!s.has(cfg.get(d))) return false;
  }
  return true;
}
// The draw period lives in the matched-invoice text (the ledger's draw_period field is empty -
// it's a QBO PrivateNote value that isn't loaded). Pull a short label: "August 2026", "Draw #4",
// or the period end date; the full "(Period: start - end)" range goes in the tooltip.
function drawPeriod(mi) {
  const s = String(mi || "");
  let m = s.match(/\b([A-Z][a-z]+) Draw (\d{4})/); if (m) return m[1] + " " + m[2];
  m = s.match(/\bDraw\s*#?\s*(\d+)/i); if (m) return "Draw #" + m[1];
  m = s.match(/Period:\s*[\d/]+\s*-\s*([\d/]+)/i); if (m) return m[1];
  return "";
}
function drawPeriodFull(mi) { const m = String(mi || "").match(/Period:\s*([\d/]+\s*-\s*[\d/]+)/i); return m ? "Period " + m[1].replace(/\s+/g, " ") : ""; }
// Resolve each PROJECT to one canonical client (GC) for grouping. billing_event customers are
// inconsistent - a project's draws can carry the GC ("JPI Construction, LLC") on some and a
// project-prefixed sub-customer ("MFD325 - BRIARWOOD") on others - so per project we PREFER a clean GC
// name and strip the project prefix otherwise. Built once per render so a project never splits.
let _drawClientByProj = {};
function _buildDrawClientMap(draws) {
  const m = {};
  for (const d of draws) {
    const p = d.project_no || ""; const raw = (d.customer || "").trim(); if (!raw) continue;
    const prefixed = /^(MFD|CP|RP)\d+/i.test(raw);
    const name = prefixed ? (raw.replace(/^(MFD|CP|RP)\d+(-FTW)?\s*[-–]\s*/i, "").trim() || raw) : raw;
    if (!m[p] || (!prefixed && m[p].prefixed)) m[p] = { name, prefixed };   // clean GC name wins
  }
  _drawClientByProj = {}; for (const p in m) _drawClientByProj[p] = m[p].name;
}
function _drawCustomer(d) { return _drawClientByProj[d.project_no || ""] || "(no client)"; }
// When a filter is active, the tab's description says WHAT it's filtering; with no filter it stays the
// generic blurb (owner 2026-08-28: "change the desc to show what it's filtering ... All = generic").
const _DRAW_FLABEL = { dfClient: "client", dfProj: "project", dfVendor: "vendor", dfInv: "invoice", dfDiv: "division" };
function drawFilterSummary(shownCount) {
  const parts = [];
  for (const cfg of DRAW_MSEL) {
    const s = drawMSel[cfg.id];    // ≤2: name the values; more: "3 vendors" (the values, not the field label)
    if (s && s.size) parts.push(s.size <= 2 ? [...s].map(cfg.lbl).join(", ") : `${s.size} ${_DRAW_FLABEL[cfg.id]}s`);
  }
  if (activeDrawStage) parts.push(DRAW_STAGE_SHORT[activeDrawStage] || activeDrawStage);
  if (!parts.length) return "";
  return `${shownCount} draw${shownCount === 1 ? "" : "s"} · ${parts.join(" · ")}`;
}
// Swap a tab's `.hint` between its generic blurb and a live "Showing: ..." filter summary.
function _setHintFilter(tab, summary) {
  const h = document.querySelector(`.tab-page[data-tab="${tab}"] .hint`); if (!h) return;
  if (!h.dataset.generic) h.dataset.generic = h.innerHTML;   // capture the generic blurb once
  if (summary) { h.innerHTML = `<b>Showing:</b> ${_ge(summary)} <span class="hint-clear-note">- clear the filters for the full list</span>`; h.classList.add("hint-filtered"); }
  else { h.innerHTML = h.dataset.generic; h.classList.remove("hint-filtered"); }
}

// ── Funding by project (owner 2026-09-02: fold Draws into the project page). One row per job:
// the next draw the GC owes, what blocks it (unpaid bills on EARLIER draws - the funding chain),
// and the latest draw's vendors paid. Click a row -> the project page, where the work happens.
const _isPaidBill = b => !!b.pay_date || (b.pay_status || "").toLowerCase().startsWith("bill paid") || (num(b.open) <= 0.005 && !!b.pay_status);
let fundingStage = null;
function _fundingRows() {
  const draws = (DRAWS.draws || []).filter(d => drawMselPasses(d) && (!drawDate || drawDate.passes(d.ar_date || d.recency)));
  const byP = new Map();
  for (const d of draws) { const k = d.project_no || "(none)"; if (!byP.has(k)) byP.set(k, []); byP.get(k).push(d); }
  const rows = [];
  for (const [pn, list] of byP) {
    list.sort((a, b) => (a.no_draw ? 1 : 0) - (b.no_draw ? 1 : 0) || String(a.ar_date || a.recency || "").localeCompare(String(b.ar_date || b.recency || "")));
    const real = list.filter(d => !d.no_draw);
    const next = real.find(d => num(d.ar_open) > 0.005) || null;
    let blockers = 0, blockAmt = 0;
    if (next) for (const d of real) { if (d === next || String(d.ar_date || "") > String(next.ar_date || "")) continue;
      for (const b of d.bills) if (b.gates && !_isPaidBill(b)) { blockers++; blockAmt += num(b.open); } }
    const latest = real.length ? real[real.length - 1] : list[list.length - 1];
    const gate = latest ? latest.bills.filter(b => b.gates) : [];
    const paidCt = gate.filter(_isPaidBill).length;
    const gcOwes = real.reduce((s, d) => s + num(d.ar_open), 0);
    const status = !real.length ? "No draw yet" : !next ? "Settled" : blockers ? "Blocked - pay vendors first" : "Ready to collect";
    const client = (_drawClientByProj[pn] || "") || list.map(d => d.customer).find(Boolean) || ((invData().invoices || []).find(i => i.project_no === pn) || {}).customer
                 || (((ALL || []).find(r => r.project_no === pn) || {}).builder_or_gc) || "";
    rows.push({ pn, client, n: real.length, gcOwes, next, blockers, blockAmt, latest, paidCt, gateN: gate.length, status,
                unpaidLatest: gate.filter(b => !_isPaidBill(b)).reduce((s, b) => s + num(b.open), 0) });
  }
  rows.sort((a, b) => b.gcOwes - a.gcOwes || a.pn.localeCompare(b.pn, undefined, { numeric: true }));
  return rows;
}
function renderFunding() {
  buildDrawFilters();
  _buildDrawClientMap(DRAWS.draws || []);   // the project -> GC map the old Draws view built (clean GC name wins)
  if (!drawDate) drawDate = dateFilter("dfDate", () => (DRAWS.draws || []).map(d => d.ar_date || d.recency), renderFunding);
  drawDate.build();
  const all = _fundingRows();
  const shown = fundingStage ? all.filter(r => r.status === fundingStage) : all;
  $("#drawsNote").textContent = (DRAWS.draws || []).length ? `(${shown.length} of ${all.length} projects · GC owes ${money(shown.reduce((s, r) => s + r.gcOwes, 0))})` : "(no draw data - run load_bill_tracker.py)";
  _setHintFilter("draws", drawFilterSummary(shown.length));
  const stats = $("#drawsStats"); stats.innerHTML = "";
  for (const [st, sub] of [["Ready to collect", "nothing blocks the next draw"], ["Blocked - pay vendors first", "earlier-draw bills unpaid"], ["Settled", "GC has paid every draw"], ["No draw yet", "bills in, nothing invoiced"]]) {
    const n = all.filter(r => r.status === st).length, amt = all.filter(r => r.status === st).reduce((s, r) => s + r.gcOwes, 0);
    const k = document.createElement("div"); k.className = "kpi kpi-click" + (fundingStage === st ? " kpi-fc" : "") + (st.startsWith("Blocked") && n ? " pnl-kpi-neg" : "");
    k.innerHTML = `<div class="k-label"></div><div class="k-value"></div><div class="k-sub"></div>`;
    k.querySelector(".k-label").textContent = st; k.querySelector(".k-value").textContent = String(n); k.querySelector(".k-sub").textContent = amt ? `${money(amt)} owed · ${sub}` : sub;
    k.onclick = () => { fundingStage = fundingStage === st ? null : st; renderFunding(); }; stats.appendChild(k);
  }
  { const b = $("#btnClearDrawStage"); if (b) { b.hidden = !fundingStage; b.onclick = () => { fundingStage = null; renderFunding(); }; } }
  const host = $("#drawList"); host.innerHTML = "";
  if (!shown.length) { const p = document.createElement("div"); p.className = "bills-cap"; p.textContent = all.length ? "No projects match." : "No draw data yet."; host.appendChild(p); return; }
  const scroll = document.createElement("div"); scroll.className = "table-scroll";
  const table = document.createElement("table"); table.className = "grid clickable"; const thead = document.createElement("thead"), tbody = document.createElement("tbody");
  const htr = document.createElement("tr");
  for (const [c, al] of [["Project", "left"], ["Client", "left"], ["Draws", "right"], ["GC owes", "right"], ["Next draw", "left"], ["Blocked by", "left"], ["Latest draw vendors", "left"], ["Status", "left"]]) { const th = document.createElement("th"); if (al === "left") th.className = "left"; th.textContent = c; htr.appendChild(th); }
  thead.appendChild(htr);
  for (const r of shown) {
    const tr = document.createElement("tr"); tr.style.cursor = "pointer"; tr.title = "Open the project page";
    tr.onclick = (e) => { if (e.target.closest("a") || e.target.closest(".cell")) return; openProjectPage(r.pn); };
    const pc = leftText(r.pn); const nm = nameOf(r.pn); if (nm) { const s = document.createElement("span"); s.className = "g-sub"; s.textContent = " · " + nm; pc.appendChild(s); } tr.appendChild(pc);
    tr.appendChild(leftText(r.client || "–"));
    tr.appendChild(rightText(String(r.n)));
    const oc = document.createElement("td"); oc.className = "right ip-amt"; oc.textContent = r.gcOwes > 0.005 ? money(r.gcOwes) : "–"; if (r.gcOwes > 0.005) oc.style.color = "var(--neg)"; tr.appendChild(oc);
    tr.appendChild(leftText(r.next ? `${r.next.invoice_no || ""} · ${money(r.next.ar_open)}${r.next.ar_date ? " · " + fmtDateShort(r.next.ar_date) : ""}` : (r.n ? "all paid" : "–")));
    const bc = document.createElement("td"); bc.className = "left"; if (r.blockers) { bc.textContent = `${r.blockers} bill${r.blockers === 1 ? "" : "s"} · ${money(r.blockAmt)}`; bc.style.color = "var(--neg)"; bc.style.fontWeight = "600"; } else bc.textContent = r.next ? "nothing" : "–"; tr.appendChild(bc);
    const vc = document.createElement("td"); vc.className = "left"; const sp = document.createElement("span"); sp.className = "ip-paid " + (r.gateN && r.paidCt === r.gateN ? "ok" : "due"); sp.textContent = r.gateN ? `${r.paidCt}/${r.gateN} paid${r.unpaidLatest > 0.005 ? " · " + money(r.unpaidLatest) + " to pay" : ""}` : "–"; vc.appendChild(sp); tr.appendChild(vc);
    const sc = leftText(r.status); if (r.status.startsWith("Blocked")) sc.style.color = "var(--neg)"; else if (r.status === "Ready to collect") sc.style.color = "var(--pos)"; tr.appendChild(sc);
    tbody.appendChild(tr);
  }
  table.appendChild(thead); table.appendChild(tbody); scroll.appendChild(table); host.appendChild(scroll);
}
function renderDraws() {
  buildDrawFilters();                              // (re)build the multi-selects when the draw set changes
  if (!drawDate) drawDate = dateFilter("dfDate", () => (DRAWS.draws || []).map(d => d.ar_date || d.recency), renderDraws);
  drawDate.build();
  const all = (DRAWS.draws || []).filter(d => drawMselPasses(d) && drawDate.passes(d.ar_date || d.recency));   // Client · Project · Vendor · Invoice · Division · Date (AND)
  const shown = activeDrawStage ? all.filter(d => d.stage === activeDrawStage) : all;
  _setHintFilter("draws", drawFilterSummary(shown.length));   // count + what's filtered (generic when nothing selected)
  $("#drawsNote").textContent = (DRAWS.draws || []).length
    ? `(${shown.length} shown of ${DRAWS.total} · most recent first)`
    : "(no draw data — run load_bill_tracker.py)";
  // Clickable stage tiles → filter the draw list. Counts come from `all` (all stages);
  // subs spell out the money direction (GC pays us in → we pay vendors out → waivers).
  const stats = [
    ["All paid", "All paid", "GC paid you + vendors paid"],
    ["Collect from GC", "Ready to turn in", "vendors paid, GC still owes"],
    ["Pay vendors", "Fund in — pay vendors", "GC funded, vendors not paid yet"],   // pump bills don't gate this
    ["Awaiting GC", "Awaiting GC funding", "not funded by the GC yet"],
    ["No draw yet", "No draw yet", "bills in, no draw invoice yet"],   // e.g. a job still 'Awaiting Invoice' in the tracker
  ];
  const sr = $("#drawsStats"); sr.innerHTML = "";
  for (const [label, stageKey, sub] of stats) {
    const n = all.filter(d => d.stage === stageKey).length;
    const el = document.createElement("div");
    el.className = "kpi kpi-click" + (activeDrawStage === stageKey ? " active" : "") + (n ? "" : " none");
    el.innerHTML = `<div class="k-label"></div><div class="k-value"></div><div class="k-sub"></div>`;
    el.querySelector(".k-label").textContent = label;
    el.querySelector(".k-value").textContent = String(n);
    el.querySelector(".k-sub").textContent = sub;
    if (n) el.onclick = () => { activeDrawStage = activeDrawStage === stageKey ? null : stageKey; renderFunding(); };
    sr.appendChild(el);
  }
  { const cb = $("#btnClearDrawStage"); if (cb) cb.hidden = !activeDrawStage; }
  // One row per draw (table). Click a row → its bills open underneath. Green = done =
  // every bill PAID (waivers are tracked per bill but don't gate the color). "Billed
  // (in)" = the GC pays you; "Paid out" = you pay vendors — money-in vs money-out.
  const box = $("#drawList"); box.innerHTML = "";
  if (!shown.length) { box.innerHTML = '<p class="hint" style="padding:14px 18px">No draws match.</p>'; return; }
  // Grouped by CUSTOMER (GC), then by project # within each; newest draw first within a project.
  _buildDrawClientMap(DRAWS.draws || []);   // resolve each project to one canonical GC (stable across filters)
  const grouped = [...shown].sort((a, b) =>
    _drawCustomer(a).localeCompare(_drawCustomer(b)) ||
    (a.project_no || "").localeCompare(b.project_no || "") ||
    String(b.ar_date || b.recency || "").localeCompare(String(a.ar_date || a.recency || "")));
  const scroll = document.createElement("div"); scroll.className = "table-scroll";
  const table = document.createElement("table"); table.className = "grid draws-table";
  const thead = document.createElement("thead"), tbody = document.createElement("tbody");
  const cols = [["", "left"], ["Draw memo", "left"], ["Period", "left"], ["Billed (in)", "right"], ["Status", "left"],
                ["Invoice #", "left"], ["Date", "left"], ["Paid out", "right"], ["Net", "right"], ["Stage", "left"]];
  const htr = document.createElement("tr");
  for (const [c, al] of cols) { const th = document.createElement("th"); if (al === "left") th.className = "left"; th.textContent = c; htr.appendChild(th); }
  thead.appendChild(htr);
  let curCust = null, curProj = null;
  for (const d of grouped) {
    const cust = _drawCustomer(d);
    if (cust !== curCust) {                               // customer (GC) group header
      curCust = cust; curProj = null;
      const cg = grouped.filter(x => _drawCustomer(x) === cust);
      const cIn = cg.reduce((t, x) => t + (x.billed || 0), 0);
      const cOut = cg.reduce((t, x) => t + (x.total_gate != null ? x.total_gate : (x.total || 0)), 0);
      const nProj = new Set(cg.map(x => x.project_no)).size;
      const ctr = document.createElement("tr"); ctr.className = "draw-cust";
      const ctd = document.createElement("td"); ctd.colSpan = cols.length;
      const cs = document.createElement("span"); cs.className = "g-cust"; cs.textContent = cust;
      const csub = document.createElement("span"); csub.className = "g-sub";
      csub.textContent = ` · ${nProj} project${nProj > 1 ? "s" : ""} · ${cg.length} draw${cg.length > 1 ? "s" : ""} · ${money(cIn)} in / ${money(cOut)} out / ${money(cIn - cOut)} net`;
      ctd.appendChild(cs); ctd.appendChild(csub); ctr.appendChild(ctd); tbody.appendChild(ctr);
    }
    if (d.project_no !== curProj) {                       // project group header (within the customer)
      curProj = d.project_no;
      const g = grouped.filter(x => x.project_no === curProj && _drawCustomer(x) === cust);
      const gIn = g.reduce((t, x) => t + (x.billed || 0), 0);
      const gOut = g.reduce((t, x) => t + (x.total_gate != null ? x.total_gate : (x.total || 0)), 0);
      const gtr = document.createElement("tr"); gtr.className = "draw-group";
      const gtd = document.createElement("td"); gtd.colSpan = cols.length;
      const sp = document.createElement("span"); sp.className = "g-proj"; sp.textContent = curProj || "—";
      const sub = document.createElement("span"); sub.className = "g-sub";
      const nm = nameOf(curProj);
      sub.textContent = `${nm ? " · " + nm : ""} · ${g.length} draw${g.length > 1 ? "s" : ""} · ${money(gIn)} in / ${money(gOut)} out / ${money(gIn - gOut)} net`;
      gtr.style.cursor = "pointer"; gtr.title = "Open the project page"; gtr.onclick = () => openProjectPage(curProj);
      gtd.appendChild(sp); gtd.appendChild(sub); gtr.appendChild(gtd); tbody.appendChild(gtr);
    }
    const done = d.stage === "All paid";   // green row only when fully settled (GC paid + vendors paid)
    const open = drawsExpanded.has(d.matched_invoice);
    const tr = document.createElement("tr"); tr.className = "draw-row" + (done ? " done" : "");
    tr.style.cursor = "pointer";
    tr.onclick = (e) => { if (e.target.closest(".cell") || e.target.closest("a")) return;
      open ? drawsExpanded.delete(d.matched_invoice) : drawsExpanded.add(d.matched_invoice); renderFunding(); };
    const cc = document.createElement("td"); cc.className = "left draw-caret"; cc.textContent = open ? "▾" : "▸"; tr.appendChild(cc);
    const memo = (d.label || "").replace(/^\s*\S+\s*—\s*/, "").replace(/^\s*(MFD|CP|RP)\d+(-FTW)?\s*-\s*/i, "").trim() || d.label || "—";
    tr.appendChild(leftText(memo));
    const per = drawPeriod(d.matched_invoice); const perCell = leftText(per || "—");
    if (per) perCell.title = drawPeriodFull(d.matched_invoice) || per; tr.appendChild(perCell);
    const bt = document.createElement("td");
    if (d.billed != null) { const mc = moneyCell(d.billed); mc.classList.add("draw-in"); bt.appendChild(mc); }
    else bt.appendChild(document.createTextNode("—"));
    tr.appendChild(bt);
    // AR pay status — its own column (green Paid / amber still-owed), not crammed onto the amount
    const stt = document.createElement("td"); stt.className = "left";
    if (d.ar_status) { const s = document.createElement("span"); s.className = d.ar_status === "Paid" ? "ar-paid" : "ar-open"; s.textContent = d.ar_status; stt.appendChild(s); }
    else stt.appendChild(document.createTextNode("—"));
    tr.appendChild(stt);
    // Invoice #: click the number for the memo + details in the sidebar (owner: "same info in
    // the sidebar for draws"), the ↗ for QuickBooks. Row click still expands the vendor bills.
    tr.appendChild(invNoCell(d.inv || (d.invoice_no ? { doc_number: d.invoice_no, qbo_txn_id: d.ar_qbo_id } : null)));
    tr.appendChild(leftText(fmtDate(d.ar_date || d.recency)));
    // Pay-out excludes MCP/CORE concrete pumping (we don't pay them) - the gating total/count.
    const payTotal = d.total_gate != null ? d.total_gate : d.total;
    const ot = document.createElement("td"); const mo = moneyCell(payTotal); mo.classList.add("draw-out"); ot.appendChild(mo);
    const pc = document.createElement("span"); pc.className = "paidcnt"; pc.textContent = ` ${d.paid_gate != null ? d.paid_gate : d.paid}/${d.n_gate != null ? d.n_gate : d.n}`; ot.appendChild(pc); tr.appendChild(ot);
    // Net = billed in minus the vendor bills we actually pay (excl. pump; not what's been paid) - the margin
    const nt = document.createElement("td");
    if (d.billed != null) { const net = (d.billed || 0) - payTotal; const nc = moneyCell(net); if (net < 0) nc.style.color = "var(--neg)"; nc.title = "billed in minus vendor bills we pay (excl. pump)"; nt.appendChild(nc); }
    else nt.appendChild(document.createTextNode("—"));
    tr.appendChild(nt);
    const st = document.createElement("td"); st.className = "left";
    const pill = document.createElement("span"); pill.className = "lien " + (DRAW_STAGE_CLASS[d.stage] || "info"); pill.textContent = DRAW_STAGE_SHORT[d.stage] || DRAW_STAGE_LABEL[d.stage] || d.stage; pill.title = DRAW_STAGE_LABEL[d.stage] || d.stage; st.appendChild(pill);
    if (d.action && d.action.url) { const a = document.createElement("a"); a.className = "notion-link"; a.href = d.action.url; a.target = "_blank"; a.rel = "noopener"; a.textContent = " 📄"; a.title = "Notion · " + (d.action.status || "Open"); st.appendChild(a); }
    tr.appendChild(st);
    tbody.appendChild(tr);
    if (open) {
      const br = document.createElement("tr"); br.className = "draw-bills-row";
      const btd = document.createElement("td"); btd.colSpan = cols.length; btd.appendChild(buildBillsTable(d));
      br.appendChild(btd); tbody.appendChild(br);
    }
  }
  table.appendChild(thead); table.appendChild(tbody); scroll.appendChild(table); box.appendChild(scroll);
}

function buildBillsTable(d) {
  const wrap = document.createElement("div"); wrap.className = "bills-sub";
  const payTotal = d.total_gate != null ? d.total_gate : d.total;   // excl MCP/CORE pumping (not paid by us)
  const subsT = d.subs_total || 0;
  const totalOut = payTotal + subsT;               // materials + subs (labor) = the full picture
  const net = (d.billed || 0) - totalOut;
  // Report header: a summary line + a Copy button (paste the draw + its subs to people). Inline - no
  // side panel (owner 2026-08-28); built from the draw data already loaded, so no extra fetch / bloat.
  const rh = document.createElement("div"); rh.className = "draw-rep-head";
  const sm = document.createElement("div"); sm.className = "draw-rep-sum";
  sm.innerHTML = `<b>${_ge(_drawCustomer(d))}</b> · ${_ge(d.project_no || "")}`
    + (drawPeriod(d.matched_invoice) ? " · " + _ge(drawPeriod(d.matched_invoice)) : "")
    + ` &nbsp;·&nbsp; Billed in <b>${_ge(d.billed != null ? money(d.billed) : "—")}</b> · Materials <b>${_ge(money(payTotal))}</b>`
    + (subsT ? ` · Subs <b>${_ge(money(subsT))}</b> · Total out <b>${_ge(money(totalOut))}</b>` : "")
    + ` · Net <b>${_ge(d.billed != null ? money(net) : "—")}</b> · ${_ge(DRAW_STAGE_SHORT[d.stage] || d.stage || "")}`;
  if (d.project_no) { const pp = document.createElement("button"); pp.type = "button"; pp.className = "btn small primary"; pp.textContent = "Project page"; pp.style.marginLeft = "8px"; pp.onclick = (e) => { e.stopPropagation(); openProjectPage(d.project_no); }; sm.appendChild(pp); }
  const cpBtn = document.createElement("button"); cpBtn.type = "button"; cpBtn.className = "btn small"; cpBtn.textContent = "Copy report";
  cpBtn.title = "Copy this draw + its subs as a table to paste to people";
  cpBtn.onclick = (e) => { e.stopPropagation(); copyDrawReport(d, cpBtn); };
  rh.appendChild(sm); rh.appendChild(cpBtn); wrap.appendChild(rh);
  const cap = document.createElement("div"); cap.className = "bills-cap";
  cap.textContent = `${d.n} bills · ${money(payTotal)} to pay · ${(d.paid_gate != null ? d.paid_gate : d.paid)}/${(d.n_gate != null ? d.n_gate : d.n)} paid`
    + (d.total - payTotal > 0.5 ? ` · ${money(d.total - payTotal)} pump (not paid by us)` : "")
    + (WAIVERS_ENABLED ? ` · ${d.waivers}/${d.n} waivers in` : "");
  wrap.appendChild(cap);
  const scroll = document.createElement("div"); scroll.className = "table-scroll";
  const table = document.createElement("table"); table.className = "grid";
  const thead = document.createElement("thead"), tbody = document.createElement("tbody");
  const cols = [["Vendor", "left"], ["Bill #", "left"], ["Bill date", "left"], ["Amount", "right"], ["Paid", "left"], ["GC funded", "left"]];
  if (WAIVERS_ENABLED) cols.push(["Waiver in hand", "left"]);
  const htr = document.createElement("tr");
  for (const [c, al] of cols) { const th = document.createElement("th"); if (al === "left") th.className = "left"; th.textContent = c; htr.appendChild(th); }
  thead.appendChild(htr);
  // Sub-group by vendor: each vendor is a header row with its total (biggest first), collapsed;
  // open it to see that vendor's bills underneath (owner 2026-08-27: "totals first, open for bills").
  const byVendor = new Map();
  for (const b of d.bills) { const v = b.vendor || "—"; if (!byVendor.has(v)) byVendor.set(v, []); byVendor.get(v).push(b); }
  const vendors = [...byVendor.keys()].sort((a, b) =>
    byVendor.get(b).reduce((t, x) => t + (x.amount || 0), 0) - byVendor.get(a).reduce((t, x) => t + (x.amount || 0), 0));
  for (const v of vendors) {
    const bills = byVendor.get(v);
    const vtot = bills.reduce((t, x) => t + (x.amount || 0), 0);
    const vpaid = bills.filter(x => x.pay_date).length;
    const vkey = d.matched_invoice + "|" + v;
    const vopen = drawVendorExpanded.has(vkey);
    const vtr = document.createElement("tr"); vtr.className = "vgroup"; vtr.style.cursor = "pointer";
    vtr.onclick = () => { vopen ? drawVendorExpanded.delete(vkey) : drawVendorExpanded.add(vkey); renderFunding(); };
    const vtd = document.createElement("td"); vtd.colSpan = cols.length; vtd.className = "left";
    const car = document.createElement("span"); car.className = "vg-caret"; car.textContent = vopen ? "▾ " : "▸ ";
    const nm = document.createElement("span"); nm.className = "vg-name"; nm.textContent = v;
    const mt = document.createElement("span"); mt.className = "vg-meta";
    mt.textContent = ` · ${bills.length} bill${bills.length > 1 ? "s" : ""} · ${money(vtot)} · ${vpaid}/${bills.length} paid`;
    vtd.appendChild(car); vtd.appendChild(nm); vtd.appendChild(mt);
    // MCP/CORE concrete pumping: shown for completeness but NOT paid by us - excluded from the
    // "Pay vendors" stage AND from the pay-out total/count/net (owner 2026-08-28).
    if (bills.length && bills[0].gates === false) {
      const tag = document.createElement("span"); tag.className = "vg-tag"; tag.textContent = "not paid by us";
      vtd.appendChild(tag);
    }
    vtr.appendChild(vtd); tbody.appendChild(vtr);
    if (!vopen) continue;
    for (const b of bills) {
      const tr = document.createElement("tr"); tr.className = "vbill";
      tr.appendChild(leftText(""));                         // vendor cell blank - grouped in the header above
      tr.appendChild(qboLinkCell(b.bill_ref, qboBillHref(b.qbo_link), "Open this bill in QuickBooks"));
      tr.appendChild(leftText(fmtDate(b.bill_date)));
      const av = document.createElement("td"); av.appendChild(moneyCell(b.amount)); tr.appendChild(av);
      tr.appendChild(leftText(b.pay_date ? "✓ " + fmtDate(b.pay_date) : "—"));
      tr.appendChild(leftText(b.gc_paid ? "✓ " + fmtDate(b.gc_paid) : "—"));
      if (WAIVERS_ENABLED) {
        const wtd = document.createElement("td"); wtd.className = "left";
        const lab = document.createElement("label"); lab.className = "chk";
        const cb = document.createElement("input"); cb.type = "checkbox"; cb.checked = !!b.waiver;
        cb.onchange = () => setWaiver(d, b, cb);
        lab.appendChild(cb); lab.appendChild(document.createTextNode(b.waiver ? " in hand" : " mark"));
        wtd.appendChild(lab); tr.appendChild(wtd);
      }
      tbody.appendChild(tr);
    }
  }
  table.appendChild(thead); table.appendChild(tbody); scroll.appendChild(table); wrap.appendChild(scroll);
  if ((d.subs || []).length) wrap.appendChild(_drawSubsTable(d));   // the labor side (subs), so it's the full picture
  return wrap;
}
// Subs (labor) on a draw - is_sub cost lines matched by project + draw period, grouped by sub. Not in
// ap_bill_line (the Bill Tracker excludes subs from the display sheets), so this is the labor side.
function _drawSubsTable(d) {
  const wrap = document.createElement("div"); wrap.className = "draw-subs";
  const cap = document.createElement("div"); cap.className = "bills-cap";
  cap.textContent = `Subs (labor) · ${money(d.subs_total)} · ${d.subs.length} sub${d.subs.length > 1 ? "s" : ""} · matched by project + draw period`;
  wrap.appendChild(cap);
  const scroll = document.createElement("div"); scroll.className = "table-scroll";
  const table = document.createElement("table"); table.className = "grid";
  const thead = document.createElement("thead"), tbody = document.createElement("tbody");
  const htr = document.createElement("tr");
  for (const [c, al] of [["Sub", "left"], ["Bills", "right"], ["Total", "right"]]) { const th = document.createElement("th"); if (al === "left") th.className = "left"; th.textContent = c; htr.appendChild(th); }
  thead.appendChild(htr);
  for (const s of d.subs) {
    const tr = document.createElement("tr");
    tr.appendChild(leftText(s.vendor || "–"));
    tr.appendChild(rightText(String(s.n || 0)));
    const tc = document.createElement("td"); tc.appendChild(moneyCell(s.total)); tr.appendChild(tc);
    tbody.appendChild(tr);
  }
  table.appendChild(thead); table.appendChild(tbody); scroll.appendChild(table); wrap.appendChild(scroll);
  return wrap;
}
// Copy a draw + its subs as a table (TSV for spreadsheets + HTML for email) to paste to people.
async function copyDrawReport(d, btn) {
  const payTotal = d.total_gate != null ? d.total_gate : d.total;
  const subsT = d.subs_total || 0, totalOut = payTotal + subsT;
  const net = (d.billed || 0) - totalOut;
  const summary = [
    ["Client", _drawCustomer(d)], ["Project", d.project_no || ""],
    ["Period", drawPeriod(d.matched_invoice) || ""], ["Invoice #", d.invoice_no || ""],
    ["Billed in", d.billed != null ? money(d.billed) : ""], ["Materials out", money(payTotal)],
    ["Subs (labor)", money(subsT)], ["Total out", money(totalOut)],
    ["Net", d.billed != null ? money(net) : ""], ["Status", DRAW_STAGE_SHORT[d.stage] || d.stage || ""],
  ];
  const billHead = ["Vendor (material)", "Bill #", "Bill date", "Amount", "Paid", "GC funded"];
  const billRows = (d.bills || []).slice().sort((a, b) => (b.amount || 0) - (a.amount || 0)).map(b => [
    b.vendor || "", b.bill_ref || "", fmtDate(b.bill_date) || "",
    money(b.amount) + (b.gates === false ? " (pump)" : ""), b.pay_date ? fmtDate(b.pay_date) : "", b.gc_paid ? fmtDate(b.gc_paid) : ""]);
  const subHead = ["Sub (labor)", "Bills", "Total"];
  const subRows = (d.subs || []).map(s => [s.vendor || "", String(s.n || 0), money(s.total)]);
  const clean = v => String(v == null ? "" : v).replace(/[\t\r\n]+/g, " ").trim();
  const tsvTable = rows => rows.map(r => r.map(clean).join("\t")).join("\n");
  const tsv = tsvTable(summary) + "\n\nMATERIALS\n" + tsvTable([billHead, ...billRows])
    + (subRows.length ? "\n\nSUBS (LABOR)\n" + tsvTable([subHead, ...subRows]) : "");
  const esc = v => _ge(String(v == null ? "" : v));
  const htmlTable = (head, rows) => "<table><thead><tr>" + head.map(h => `<th>${esc(h)}</th>`).join("") + "</tr></thead><tbody>"
    + rows.map(r => "<tr>" + r.map(v => `<td>${esc(v)}</td>`).join("") + "</tr>").join("") + "</tbody></table>";
  const html = "<table>" + summary.map(r => `<tr><th style="text-align:left">${esc(r[0])}</th><td>${esc(r[1])}</td></tr>`).join("") + "</table><br>"
    + htmlTable(billHead, billRows) + (subRows.length ? "<br>" + htmlTable(subHead, subRows) : "");
  let ok = false;
  try {
    if (window.ClipboardItem && navigator.clipboard && navigator.clipboard.write) {
      await navigator.clipboard.write([new ClipboardItem({ "text/plain": new Blob([tsv], { type: "text/plain" }), "text/html": new Blob([html], { type: "text/html" }) })]); ok = true;
    } else if (navigator.clipboard && navigator.clipboard.writeText) { await navigator.clipboard.writeText(tsv); ok = true; }
  } catch (e) {
    try { const ta = document.createElement("textarea"); ta.value = tsv; ta.style.cssText = "position:fixed;opacity:0"; document.body.appendChild(ta); ta.select(); ok = document.execCommand("copy"); ta.remove(); } catch (_) { /* ignore */ }
  }
  if (btn) { const t = btn.textContent; btn.disabled = true; btn.textContent = ok ? "Copied ✓" : "Copy failed"; setTimeout(() => { btn.textContent = t; btn.disabled = false; }, 1400); }
  toast(ok ? `Draw report copied (${(d.bills || []).length} subs)` : "Copy failed");
}
async function setWaiver(draw, bill, cb) {
  const received = cb.checked;
  try {
    const res = await fetch("/api/waiver", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ matched_invoice: draw.matched_invoice, vendor: bill.vendor, bill_ref: bill.bill_ref, received }),
    });
    const j = await res.json();
    if (!j.ok) throw new Error(j.error || "write failed");
    bill.waiver = received;                       // update local state
    draw.waivers = draw.bills.filter(b => b.waiver).length;   // caption only — doesn't gate the stage
    toast(received ? "Waiver marked in hand" : "Waiver cleared");
    renderFunding();
  } catch (e) {
    cb.checked = !received;                        // revert on failure
    toast("Could not save: " + e.message);
  }
}

function renderMargins() {
  // portfolio, over jobs that actually have QBO costs loaded
  const loaded = ALL.filter(r => r.costs_loaded != null);
  const cost = loaded.reduce((t, r) => t + num(r.costs_loaded), 0);
  const billed = loaded.reduce((t, r) => t + num(r.billed_to_date), 0);
  const subs = loaded.reduce((t, r) => t + num(r.sub_costs), 0);
  const margin = billed - cost;
  const overBudget = loaded.filter(isOverBudget).length;
  const cSum = loaded.reduce((t, r) => t + num(r.total_contract_price), 0);
  const eSum = loaded.reduce((t, r) => t + num(r.estimated_total_costs), 0);
  const pct1 = (n, d) => d ? (n / d * 100).toFixed(1) + "%" : "—";
  $("#marginNote").textContent = loaded.length ? `(${loaded.length} jobs with QBO costs)` : "(no cost data — run load_costs.py)";
  const stats = [
    ["Planned markup", pct1(cSum - eSum, eSum), "contract vs ETC (on cost)"],
    ["Planned margin", pct1(cSum - eSum, cSum), "GP ÷ contract"],
    ["Actual markup", pct1(billed - cost, cost), "billed ÷ QBO cost"],
    ["Margin to date", money(margin), "billed − QBO cost"],
    ["Subs share", cost ? (subs / cost * 100).toFixed(0) + "%" : "—", "of QBO cost"],
    ["Over budget", String(overBudget), "cost > ETC (flatwork soft <$15k)"],
  ];
  const sr = $("#marginStats"); sr.innerHTML = "";
  for (const [label, value, sub] of stats) {
    const el = document.createElement("div"); el.className = "kpi";
    el.innerHTML = `<div class="k-label"></div><div class="k-value"></div><div class="k-sub"></div>`;
    el.querySelector(".k-label").textContent = label;
    el.querySelector(".k-value").textContent = value;
    el.querySelector(".k-sub").textContent = sub;
    sr.appendChild(el);
  }
  // OVER-BUDGET jobs — cost past the full ETC, with the flatwork-budget tolerance
  // applied (small -FTW jobs excluded; slab / CP / MFD / big flatwork kept). Worst first.
  const watch = loaded
    .filter(isOverBudget)
    .sort((a, b) => b.budget_burn - a.budget_burn).slice(0, 10);
  const cols = [["Project", "left"], ["Name", "left"], ["Burn", "right"], ["QBO Cost", "right"], ["ETC", "right"]];
  const thead = $("#marginTable thead"), tbody = $("#marginTable tbody");
  thead.innerHTML = ""; tbody.innerHTML = "";
  const htr = document.createElement("tr");
  for (const [c, al] of cols) { const th = document.createElement("th"); if (al === "left") th.className = "left"; th.textContent = c; htr.appendChild(th); }
  thead.appendChild(htr);
  for (const r of watch) {
    const tr = document.createElement("tr");
    tr.onclick = (e) => { if (!e.target.closest(".cell")) openDetail(r); };
    tr.appendChild(leftText(r.project_no));
    tr.appendChild(leftText(r.project_name || ""));
    const bt = document.createElement("td");
    const bS = document.createElement("span"); bS.className = "cell neg"; bS.textContent = pct(r.budget_burn);
    bS.title = "Click to copy"; bS.onclick = (e) => { e.stopPropagation(); copy(String(r.budget_burn)); };
    bt.appendChild(bS); tr.appendChild(bt);
    const cv = document.createElement("td"); cv.appendChild(moneyCell(r.costs_loaded)); tr.appendChild(cv);
    const ev = document.createElement("td"); ev.appendChild(moneyCell(r.estimated_total_costs)); tr.appendChild(ev);
    tbody.appendChild(tr);
  }
}

function renderCosts() {
  // Grouped: cost TYPE (parent) → job TYPE (sub) — the JobTread model. Material
  // rolls to ONE cost-type parent; the job-type sub shows the split for budget.
  const groups = COST.by_cost_type || [];
  const total = COST.loaded_total || groups.reduce((t, g) => t + (g.actual || 0), 0);
  $("#costCount").textContent = total
    ? `($${Math.round(total).toLocaleString()} loaded from QuickBooks${loadedAt("Costs (QBO)") ? " " + fmtDate(loadedAt("Costs (QBO)"), true) : ""} · where the money goes)`
    : "(no cost data — run load_costs.py)";
  renderCostMix(groups, total);
  const cols = [["Cost type  ▸  job type", "left"], ["Code", "left"], ["Actual", "right"], ["% of total", "right"], ["Lines", "right"]];
  const thead = $("#costTreeTable thead"), tbody = $("#costTreeTable tbody");
  if (!thead) return;
  thead.innerHTML = ""; tbody.innerHTML = "";
  const htr = document.createElement("tr");
  for (const [c, al] of cols) { const th = document.createElement("th"); if (al === "left") th.className = "left"; th.textContent = c; htr.appendChild(th); }
  thead.appendChild(htr);
  const max = groups.reduce((m, g) => Math.max(m, g.actual || 0), 0) || 1;
  for (const g of groups) {
    const collapsed = costCollapsed.has(g.parent);
    // ── parent row: the cost TYPE (material) ──
    const ptr = document.createElement("tr"); ptr.className = "cost-parent";
    const nameTd = document.createElement("td"); nameTd.className = "left";
    const tw = document.createElement("span"); tw.className = "tw"; tw.textContent = (collapsed ? "▸ " : "▾ ") + g.parent;
    nameTd.appendChild(tw); ptr.appendChild(nameTd);
    ptr.appendChild(leftText(""));
    const at = document.createElement("td");
    const bar = document.createElement("span"); bar.className = "cell bar";
    const fill = document.createElement("span"); fill.className = "bar-fill"; fill.style.width = ((g.actual || 0) / max * 100) + "%";
    const txt = document.createElement("span"); txt.className = "bar-txt"; txt.textContent = money(g.actual);
    bar.appendChild(fill); bar.appendChild(txt); bar.title = "Click to copy"; bar.onclick = (e) => { e.stopPropagation(); copy(String(Math.round(g.actual || 0))); };
    at.appendChild(bar); ptr.appendChild(at);
    ptr.appendChild(rightText(total ? ((g.actual || 0) / total * 100).toFixed(1) + "%" : "—"));
    ptr.appendChild(rightText(String(g.lines || 0)));
    ptr.onclick = (e) => { if (!e.target.closest(".cell")) { collapsed ? costCollapsed.delete(g.parent) : costCollapsed.add(g.parent); renderCosts(); } };
    tbody.appendChild(ptr);
    // ── sub rows: the job TYPE split ──
    if (!collapsed) for (const s of g.subs) {
      const str = document.createElement("tr"); str.className = "cost-sub";
      if (s.code) str.onclick = (e) => { if (!e.target.closest(".cell")) showCodeJobs(s.code, `${g.parent} ▸ ${s.sub}`); };
      const sn = document.createElement("td"); sn.className = "left";
      const si = document.createElement("span"); si.className = "sub-name"; si.textContent = s.sub; sn.appendChild(si); str.appendChild(sn);
      const ct = document.createElement("td"); ct.className = "left";
      if (s.code) { const chip = document.createElement("span"); chip.className = "codechip"; chip.textContent = s.code; ct.appendChild(chip); }
      else ct.appendChild(document.createTextNode("—"));
      str.appendChild(ct);
      const av = document.createElement("td"); av.appendChild(moneyCell(s.actual)); str.appendChild(av);
      str.appendChild(rightText(total ? ((s.actual || 0) / total * 100).toFixed(1) + "%" : "—"));
      str.appendChild(rightText(String(s.lines || 0)));
      tbody.appendChild(str);
    }
  }
}
function rightText(v) { const td = document.createElement("td"); const s = document.createElement("span"); s.textContent = v; td.appendChild(s); return td; }

// Cost mix — "how much each cost type takes, % wise" as one proportional bar + legend.
const MIX_PALETTE = ["#4A6B8A", "#3E7A5C", "#B9541E", "#6b5b95", "#b8860b", "#1f7a4d",
                     "#B4341E", "#4478a0", "#7a5c3e", "#5c8a6b", "#8a4a6b", "#997a3d"];
function renderCostMix(groups, total) {
  const box = $("#costMix"); box.innerHTML = "";
  if (!total || !groups.length) return;
  const bar = document.createElement("div"); bar.className = "mixbar";
  const legend = document.createElement("div"); legend.className = "mixlegend";
  groups.forEach((g, i) => {
    const p = (g.actual || 0) / total * 100;
    const color = MIX_PALETTE[i % MIX_PALETTE.length];
    const seg = document.createElement("span"); seg.className = "mixseg";
    seg.style.width = p + "%"; seg.style.background = color;
    seg.title = `${g.parent}: ${p.toFixed(1)}% (${money(g.actual)})`;
    bar.appendChild(seg);
    if (i < 7 && p >= 0.5) {
      const key = document.createElement("span"); key.className = "mixkey";
      const dot = document.createElement("span"); dot.className = "mixdot"; dot.style.background = color;
      key.appendChild(dot);
      key.appendChild(document.createTextNode(`${g.parent} ${p.toFixed(1)}%`));
      legend.appendChild(key);
    }
  });
  box.appendChild(bar); box.appendChild(legend);
}

// Cost-code pivot: click a code in the tree → every job that spent on it.
function showCodeJobs(code, label) {
  if (!code) return;
  const rows = [];
  for (const [proj, codes] of Object.entries(COST.by_project_code || {}))
    for (const c of codes) if (c.cost_code === code) rows.push({ project: proj, actual: c.actual, lines: c.lines });
  rows.sort((a, b) => (b.actual || 0) - (a.actual || 0));
  const tot = rows.reduce((t, r) => t + (r.actual || 0), 0);
  $("#codeJobsTitle").textContent = `${label} (${code}) — ${money(tot)} across ${rows.length} job${rows.length === 1 ? "" : "s"}`;
  const cols = [["Project", "left"], ["Name", "left"], ["Amount", "right"], ["% of code", "right"], ["Lines", "right"]];
  const thead = $("#codeJobsTable thead"), tbody = $("#codeJobsTable tbody");
  thead.innerHTML = ""; tbody.innerHTML = "";
  const htr = document.createElement("tr");
  for (const [c, al] of cols) { const th = document.createElement("th"); if (al === "left") th.className = "left"; th.textContent = c; htr.appendChild(th); }
  thead.appendChild(htr);
  const known = new Set(ALL.map(r => r.project_no));
  for (const r of rows) {
    const tr = document.createElement("tr");
    if (known.has(r.project)) tr.onclick = (e) => { if (!e.target.closest(".cell")) openDetail(ALL.find(x => x.project_no === r.project)); };
    tr.appendChild(leftText(r.project));
    tr.appendChild(leftText(nameOf(r.project)));
    const av = document.createElement("td"); av.appendChild(moneyCell(r.actual)); tr.appendChild(av);
    tr.appendChild(rightText(tot ? ((r.actual || 0) / tot * 100).toFixed(1) + "%" : "—"));
    tr.appendChild(rightText(String(r.lines || 0)));
    tbody.appendChild(tr);
  }
  $("#codeJobsWidget").hidden = false;
  $("#codeJobsWidget").scrollIntoView({ behavior: "smooth", block: "nearest" });
}

const LIEN_ORDER = ["Notice PAST due", "Notice due in ≤7d", "Notice due in ≤15d",
                    "Notice due in ≤30d", "Notice Sent", "Lien Filed"];
// Short tile labels for the clickable stage widgets (full status = pill / tooltip).
const LIEN_SHORT = {
  "Notice PAST due": "Past due", "Notice due in ≤7d": "Due ≤7d",
  "Notice due in ≤15d": "Due ≤15d", "Notice due in ≤30d": "Due ≤30d",
  "Notice Sent": "Notice sent", "Lien Filed": "Lien filed",
};

// Pull the draw # and property name out of a matched_invoice label like
// "34449 — CP745 - Firestone Forever…" → { draw:"34449", name:"Firestone…" }.
// Used only as a fallback — invoice_no and the WIP name win when present.
function splitDraw(mi) {
  if (!mi) return { draw: "", name: "" };
  const head = String(mi).split("\n")[0].trim();
  const m = head.match(/^\s*([^—]*?)\s*—\s*(.*)$/);
  if (!m) return { draw: head, name: "" };
  const rest = m[2].trim();                       // "CP745 - Firestone Forever…"
  const dash = rest.indexOf(" - ");
  return { draw: m[1].trim(), name: dash >= 0 ? rest.slice(dash + 3).trim() : rest };
}

// Liens-page multi-select filters (same checkbox UI as the Bills tab), self-contained. Built from
// the lien watchlist (all divisions), so you filter, not just search.
const lienMSel = {};   // { id: Set }
let _lienMSelSig = null;   // data signature; the msels rebuild only when it changes (not on a toggle)
const LIEN_MSEL = [
  { id: "lfClient", all: "All clients",     get: r => r.client || "",         search: true, lbl: v => v || "(no client)" },
  { id: "lfVendor", all: "All vendors",     get: r => r.vendor || "",         search: true, lbl: v => v || "(none)" },
  { id: "lfDiv",    all: "All divisions",   get: r => r.division || "",        lbl: v => v || "(none)" },
  { id: "lfPay",    all: "Any inv. status", get: r => r.inv_ar_status || "",   lbl: v => v || "(no invoice)" },
];
function _lienLabelUpdate(cfg) {
  const btn = $("#" + cfg.id + "Btn"), menu = $("#" + cfg.id + "Menu"); if (!btn) return;
  const s = lienMSel[cfg.id] || new Set();
  btn.textContent = !s.size ? cfg.all : (s.size === 1 ? cfg.lbl([...s][0]) : s.size + " selected");
  btn.classList.toggle("on", s.size > 0);
  const cnt = menu ? menu.querySelector(".msel-count") : null; if (cnt) cnt.textContent = `${s.size} selected`;
}
function _lienBulk(cfg, sel) {   // Select all / None over the VISIBLE (search-filtered) options
  const menu = $("#" + cfg.id + "Menu"); if (!menu) return;
  const s = lienMSel[cfg.id] || (lienMSel[cfg.id] = new Set());
  for (const lab of menu.querySelectorAll(".msel-opt")) {
    if (lab.hidden) continue;
    const v = lab.dataset.val;
    if (sel) s.add(v); else s.delete(v);
    const cb = lab.querySelector("input"); if (cb) cb.checked = sel;
  }
  _lienLabelUpdate(cfg); renderLiens();
}
function buildLienMSel(cfg, watch) {
  const menu = $("#" + cfg.id + "Menu"), btn = $("#" + cfg.id + "Btn");
  if (!menu || !btn) return;
  const s = lienMSel[cfg.id] || (lienMSel[cfg.id] = new Set());
  const vals = [...new Set(watch.map(cfg.get))].sort((a, b) => cfg.lbl(a).localeCompare(cfg.lbl(b)));
  for (const v of [...s]) if (!vals.includes(v)) s.delete(v);
  menu.innerHTML = "";
  if (cfg.search) { const q = document.createElement("input"); q.type = "search"; q.className = "msel-search"; q.placeholder = "Search";
    q.oninput = () => { const t = q.value.toLowerCase(); for (const lab of menu.querySelectorAll(".msel-opt")) lab.hidden = t && !lab.textContent.toLowerCase().includes(t); }; menu.appendChild(q);
    const tools = document.createElement("div"); tools.className = "msel-tools";
    const all = document.createElement("button"); all.type = "button"; all.className = "msel-tool"; all.textContent = "Select all"; all.onclick = () => _lienBulk(cfg, true);
    const none = document.createElement("button"); none.type = "button"; none.className = "msel-tool"; none.textContent = "None"; none.onclick = () => _lienBulk(cfg, false);
    const cnt = document.createElement("span"); cnt.className = "msel-count";
    tools.appendChild(all); tools.appendChild(none); tools.appendChild(cnt); menu.appendChild(tools); }
  { const clr = document.createElement("button"); clr.type = "button"; clr.className = "msel-clear"; clr.textContent = "Clear"; clr.onclick = () => { s.clear(); buildLienMSel(cfg, watch); renderLiens(); }; menu.appendChild(clr); }
  for (const v of vals) {
    const lab = document.createElement("label"); lab.className = "msel-opt"; lab.dataset.val = v;
    const cb = document.createElement("input"); cb.type = "checkbox"; cb.checked = s.has(v);
    cb.onchange = () => { if (cb.checked) s.add(v); else s.delete(v); _lienLabelUpdate(cfg); renderLiens(); };
    lab.appendChild(cb); lab.appendChild(document.createTextNode(" " + cfg.lbl(v)));
    menu.appendChild(lab);
  }
  _lienLabelUpdate(cfg);
}
function lienMSelPasses(r) { for (const cfg of LIEN_MSEL) { const s = lienMSel[cfg.id]; if (s && s.size && !s.has(cfg.get(r))) return false; } return true; }
function renderLiens() {
  // This page is ONLY what's actually been SENT or FILED (owner 2026-08-20) - not the deadline clock.
  const watch = AP.liens || [];
  $("#liensNote").textContent = watch.length ? `(${watch.length} sent + filed)` : "(none sent or filed yet)";
  const openOf = rows => rows.reduce((t, r) => t + num(r.open_balance), 0);
  const sent = watch.filter(r => r.lien_status === "Notice Sent");
  const filed = watch.filter(r => r.lien_status === "Lien Filed");
  // ── summary KPIs ──
  const stats = [
    ["Notices sent", String(sent.length), money(openOf(sent)) + " open"],
    ["Liens filed", String(filed.length), money(openOf(filed)) + " open"],
    ["Open $ at stake", money(openOf(watch)), `across ${watch.length} bill${watch.length === 1 ? "" : "s"}`],
  ];
  const sr = $("#liensStats"); sr.innerHTML = "";
  for (const [label, value, sub] of stats) {
    const el = document.createElement("div"); el.className = "kpi";
    el.innerHTML = `<div class="k-label"></div><div class="k-value"></div><div class="k-sub"></div>`;
    el.querySelector(".k-label").textContent = label;
    el.querySelector(".k-value").textContent = value;
    el.querySelector(".k-sub").textContent = sub;
    sr.appendChild(el);
  }
  // ── clickable tiles → filter to Notice Sent / Lien Filed ──
  const byStatus = {};
  for (const r of watch) (byStatus[r.lien_status] || (byStatus[r.lien_status] = [])).push(r);
  const filters = $("#lienFilters"); filters.innerHTML = "";
  const tile = (key, label, n, open, cls, active) => {
    const el = document.createElement("div");
    el.className = "attn" + (cls ? " u-" + cls : "") + (active ? " active" : "") + (n ? "" : " none");
    el.innerHTML = `<span class="a-count"></span><span class="a-label"></span><span class="a-sub"></span>`;
    el.querySelector(".a-count").textContent = n;
    el.querySelector(".a-label").textContent = label;
    el.querySelector(".a-sub").textContent = money(open) + " open";
    if (n || key === null) el.onclick = () => { activeLien = key; renderLiens(); };
    filters.appendChild(el);
  };
  tile(null, "All sent + filed", watch.length, openOf(watch), "", activeLien === null);
  for (const status of ["Lien Filed", "Notice Sent"]) {
    const rows = byStatus[status]; if (!rows || !rows.length) continue;
    tile(status, LIEN_SHORT[status] || status, rows.length, openOf(rows), LIEN_CLASS[status] || "info", activeLien === status);
  }
  $("#btnClearLien").hidden = activeLien === null;

  // ── multi-select filters (built from the sent/filed list) + the active stage + a Project # search ──
  // Rebuild only when the underlying data changes - NOT on every render - so a checkbox toggle keeps the
  // open search box + scroll position instead of rebuilding the menu under the user (owner 2026-08-21). A
  // toggle re-renders with the SAME watch, so its signature is unchanged; a data load makes a new one.
  const lienSig = `${watch.length}:${watch[0] ? watch[0].bill_id : ""}:${watch.length ? watch[watch.length - 1].bill_id : ""}`;
  if (lienSig !== _lienMSelSig || !$("#lfVendorMenu") || !$("#lfVendorMenu").querySelector(".msel-opt")) {
    _lienMSelSig = lienSig;
    for (const cfg of LIEN_MSEL) buildLienMSel(cfg, watch);
  }
  const qProj = ($("#lienFProj") ? $("#lienFProj").value : "").trim().toLowerCase();
  const known = new Set(ALL.map(r => r.project_no));
  const base = activeLien ? (byStatus[activeLien] || []) : watch;
  const shown = base.filter(r => {
    if (qProj && !`${r.project_no || ""} ${r.invoice_no || ""}`.toLowerCase().includes(qProj)) return false;
    if (!lienMSelPasses(r)) return false;
    return true;
  });

  // Vendor first, then Date · Amount · Invoice # (the bill) · Client · Project # (ALL divisions,
  // not CP) · the AR invoice it's associated with + whether the client PAID that invoice. Urgency
  // is the coloured row edge.
  const payShort = st => !st ? null : (/paid/i.test(st) && !/unpaid|partial/i.test(st) ? ["Paid", "st-ok"]
    : (/partial/i.test(st) ? ["Partial", "st-warn"] : ["Unpaid", "st-warn"]));
  const cols = [["Vendor", "left"], ["Date", "left"], ["Amount", "right"], ["Invoice #", "left"],
                ["Client", "left"], ["Project #", "left"], ["Invoice associated", "left"], ["Invoice pay status", "left"]];
  const thead = $("#lienTable thead"), tbody = $("#lienTable tbody");
  thead.innerHTML = ""; tbody.innerHTML = "";
  const htr = document.createElement("tr");
  for (const [c, al] of cols) { const th = document.createElement("th"); if (al === "left") th.className = "left"; th.textContent = c; htr.appendChild(th); }
  thead.appendChild(htr);
  for (const r of shown) {
    const tr = document.createElement("tr");
    tr.className = "lien-row u-" + (LIEN_CLASS[r.lien_status] || "info");
    tr.title = r.lien_status || "";
    tr.onclick = (e) => {                       // → the bill + its invoice/draw, with QBO links
      if (e.target.closest(".cell") || e.target.closest("a")) return;
      const fb = findBillForLien(r);
      if (fb) openBillDetail(fb);
      else if (r.project_no && known.has(r.project_no)) openDetail(ALL.find(x => x.project_no === r.project_no));
    };
    tr.appendChild(leftText(r.vendor || "–"));                                    // Vendor
    tr.appendChild(leftText(r.bill_date ? fmtDateShort(r.bill_date) : "–"));      // Date
    const amt = document.createElement("td"); const mc = moneyCell(r.open_balance); mc.classList.add("lien-amt"); amt.appendChild(mc); tr.appendChild(amt);   // Amount (open)
    // Invoice # = the bill's own number → QBO bill
    const inv = document.createElement("td"); inv.className = "left"; let chip;
    if (r.qbo_link) { chip = document.createElement("a"); chip.href = qboBillHref(r.qbo_link); chip.target = "_blank"; chip.rel = "noopener"; chip.title = "Open this bill in QuickBooks"; chip.onclick = (e) => e.stopPropagation(); chip.className = "invno qbo-link"; }
    else { chip = document.createElement("span"); chip.className = "invno"; }
    chip.textContent = r.bill_ref || "–"; inv.appendChild(chip); tr.appendChild(inv);
    tr.appendChild(leftText(r.client || "–"));                                    // Client
    tr.appendChild(leftText(r.project_no || "–"));                                // Project #
    // Invoice associated = the AR draw invoice → QBO invoice
    const ia = document.createElement("td"); ia.className = "left";
    if (r.invoice_no && r.inv_qbo_id) { const a = document.createElement("a"); a.href = qboInvoiceUrl(r.inv_qbo_id); a.target = "_blank"; a.rel = "noopener"; a.className = "qbo-link"; a.textContent = r.invoice_no; a.title = "Open this invoice in QuickBooks"; a.onclick = (e) => e.stopPropagation(); ia.appendChild(a); }
    else ia.appendChild(document.createTextNode(r.invoice_no || "–"));
    tr.appendChild(ia);
    // Invoice pay status = did the CLIENT pay that AR invoice (the lien tell: they paid, you didn't)
    const ps = document.createElement("td"); ps.className = "left";
    const pj = payShort(r.inv_ar_status);
    if (pj) ps.appendChild(stText(pj[0], pj[1], r.inv_ar_status)); else ps.appendChild(document.createTextNode("–"));
    tr.appendChild(ps);
    tbody.appendChild(tr);
  }
  if (!shown.length) {
    const tr = document.createElement("tr"); const td = document.createElement("td");
    td.colSpan = cols.length; td.className = "left"; td.style.color = "var(--text-dim)";
    td.textContent = watch.length ? "No bills match these filters." : "No AP data - run load_bill_tracker.py.";
    tr.appendChild(td); tbody.appendChild(tr);
  }
}

let vendorTypeExpanded = new Set();   // vendor TYPE groups expanded (default none = all collapsed, owner 2026-08-31)
let _vendorGroupKeys = [];
function _updateVendorExpandBtn() {
  const b = $("#vendorExpandAll"); if (!b) return;
  const grouped = _vendorGroupKeys.length > 0;
  b.style.display = grouped ? "" : "none";
  const allExp = grouped && _vendorGroupKeys.every(k => vendorTypeExpanded.has(k));
  b.textContent = allExp ? "Collapse all" : "Expand all";
}
function _vendorToggleAll() {
  const allExp = _vendorGroupKeys.length && _vendorGroupKeys.every(k => vendorTypeExpanded.has(k));
  if (allExp) vendorTypeExpanded.clear(); else _vendorGroupKeys.forEach(k => vendorTypeExpanded.add(k));
  renderVendors();
}
function renderVendors() {
  const q = ($("#vendorSearch") ? $("#vendorSearch").value : "").trim().toLowerCase();
  let vends = COST.by_vendor || [];
  if (q) vends = vends.filter(v => (v.vendor || "").toLowerCase().includes(q));
  const grouped = $("#vendorGroupType") && $("#vendorGroupType").checked;
  const totalOpen = vends.reduce((t, v) => t + (v.open_bal || 0), 0);
  $("#vendorsNote").textContent = (COST.by_vendor || []).length
    ? `(${vends.length} vendors · ${money(totalOpen)} open)`
    : "(no cost data — run load_costs.py)";
  const cols = [["Vendor", "left"], ["Type", "left"], ["Jobs", "right"], ["Open bills (QBO)", "right"], ["Open $ (QBO)", "right"]];   // labelled: QuickBooks open AP, subs included
  const thead = $("#vendorTable thead"), tbody = $("#vendorTable tbody");
  thead.innerHTML = ""; tbody.innerHTML = "";
  const htr = document.createElement("tr");
  for (const [c, al] of cols) { const th = document.createElement("th"); if (al === "left") th.className = "left"; th.textContent = c; htr.appendChild(th); }
  thead.appendChild(htr);
  const gType = v => (v.vtype || "—").split(":")[0].trim();   // Sub | Service | Supplier
  const rows = [...vends].sort(grouped
    ? (a, b) => gType(a).localeCompare(gType(b)) || (b.open_bal || 0) - (a.open_bal || 0)
    : (a, b) => (b.open_bal || 0) - (a.open_bal || 0));       // default: most owed first
  const vendorRow = v => {
    const tr = document.createElement("tr");
    tr.classList.add("row-click"); tr.title = "Open this vendor's page";
    tr.onclick = (e) => { if (e.target.closest(".cell")) return; openVendorPage(v.vendor); };
    tr.appendChild(leftText(v.vendor));
    const ty = document.createElement("td"); ty.className = "left";
    const pill = document.createElement("span"); pill.className = "vtype" + (v.vtype === "Sub" ? " sub" : (v.vtype === "Service" ? " service" : ""));
    pill.textContent = v.vtype || "—"; ty.appendChild(pill); tr.appendChild(ty);
    tr.appendChild(rightText(String(v.jobs || 0)));
    tr.appendChild(rightText(v.open_bills ? String(v.open_bills) : "–"));
    const oc = document.createElement("td");
    if (v.open_bal > 0.5) oc.appendChild(moneyCell(v.open_bal)); else oc.appendChild(document.createTextNode("–"));
    tr.appendChild(oc);
    return tr;
  };
  if (grouped) {   // type groups, COLLAPSED by default; open a type to see its vendors (owner 2026-08-31)
    const order = [], byType = new Map();
    for (const v of rows) { const t = gType(v); if (!byType.has(t)) { byType.set(t, []); order.push(t); } byType.get(t).push(v); }
    _vendorGroupKeys = order;
    for (const t of order) {
      const gv = byType.get(t), expanded = vendorTypeExpanded.has(t);
      const gopen = gv.reduce((s, x) => s + (x.open_bal || 0), 0);
      const gtr = document.createElement("tr"); gtr.className = "draw-cust"; gtr.style.cursor = "pointer";
      gtr.title = expanded ? "Click to collapse" : "Click to expand";
      const gtd = document.createElement("td"); gtd.colSpan = cols.length;
      const caret = document.createElement("span"); caret.className = "bg-caret"; caret.textContent = expanded ? "▾ " : "▸ ";
      const gs = document.createElement("span"); gs.className = "g-cust"; gs.textContent = t;
      const cell = document.createElement("div"); cell.className = "bg-cell"; const left = document.createElement("span"); left.className = "bg-left"; left.appendChild(caret); left.appendChild(gs); cell.appendChild(left);
      bandMetrics(cell, [[gv.length, "vendors"], [money(gopen), "open"]]);
      gtd.appendChild(cell); gtr.appendChild(gtd);
      gtr.onclick = () => { if (vendorTypeExpanded.has(t)) vendorTypeExpanded.delete(t); else vendorTypeExpanded.add(t); renderVendors(); };
      tbody.appendChild(gtr);
      if (expanded) for (const v of gv.slice(0, 300)) tbody.appendChild(vendorRow(v));
    }
  } else {
    _vendorGroupKeys = [];
    for (const v of rows.slice(0, 300)) tbody.appendChild(vendorRow(v));
  }
  _updateVendorExpandBtn();
}
function leftText(v) { const td = document.createElement("td"); td.className = "left"; const s = document.createElement("span"); s.textContent = v; td.appendChild(s); return td; }
// Every grouped band (Bills vendors, Invoices clients, Payments periods, WIP divisions, the invoice
// and project pages' vendors...) lays its figures out the SAME way: the name on the left, then fixed
// metric columns, each a value with its label under it - so bands line up down the page and every
// number says what it is (owner 2026-09-02: "a column I can rely on ... ALL headers like this").
function bandMetrics(cell, metrics) {
  const ms = (metrics || []).filter(Boolean);
  cell.classList.add("bg-grid"); cell.style.setProperty("--nm", String(ms.length));
  for (const [v, label, cls] of ms) {
    const m = document.createElement("span"); m.className = "bg-m" + (cls ? " " + cls : "");
    const b = document.createElement("b"); b.textContent = v == null ? "–" : String(v);
    const s = document.createElement("small"); s.textContent = label || "";
    m.appendChild(b); m.appendChild(s); cell.appendChild(m);
  }
  return cell;
}

// ── Bill Tracker (the full ap_bill_line) ──────────────────────────────────────
// An Excel-dense table you scroll like the workbook. Default: open bills, grouped
// by vendor A→Z, oldest bill first. The chips are quick presets; every field below
// is its own filter dropdown (not a search box). Group + sort are yours to change.
let activeBillView = "open";
let billsCollapsed = new Set();   // group keys the owner has collapsed (caret / Collapse-all)
let billGroupKeys = [];           // group keys currently on screen (drives the Collapse/Expand-all button)
const BILL_LIEN_RISK = new Set(["Notice PAST due", "Notice due in ≤7d", "Lien Filed"]);
// A row goes RED only when a notice/lien ACTUALLY EXISTS on the bill (owner 2026-08-19:
// "it's when there is a notice or lien filed that we should make red") - NOT merely because a
// computed deadline passed. The deadline urgency still shows in the lien CELL colour; the red
// row line means an actual notice was sent or a lien was filed against the job.
const BILL_LIEN_ACTIVE = new Set(["Notice Sent", "Lien Filed"]);
// Compact numeric date for the dense grid: MM/DD/YY (01/01/26). Still month-first (never
// year-first). bill_date is ISO yyyy-mm-dd.
function fmtDateShort(v) {   // mm/dd/yyyy (owner 2026-08-21: 4-digit year everywhere)
  const m = String(v || "").trim().match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? `${m[2]}/${m[3]}/${m[1]}` : (v ? String(v) : "–");
}
// Per-column widths for the Bills grid - drag the divider between headers to resize;
// a squished column wraps its text instead of clipping. Widths persist per person.
const BILL_COL_DEFAULTS = { "Vendor": 210, "Project": 300, "Bill #": 100, "Date": 110, "Amount": 100,
  "Open": 100, "Paid": 90, "Invoice": 120, "Lien": 130, "Appr": 70 };
function loadBillColWidths() {
  try { return { ...BILL_COL_DEFAULTS, ...JSON.parse(localStorage.getItem("proficient-ledger-billcols") || "{}") }; }
  catch { return { ...BILL_COL_DEFAULTS }; }
}
let billColW = loadBillColWidths();
function saveBillColWidths() { try { localStorage.setItem("proficient-ledger-billcols", JSON.stringify(billColW)); } catch { /* ignore */ } }
function startBillColResize(e, idx, label) {
  e.preventDefault(); e.stopPropagation();
  const table = $("#billTable"); const cg = table.querySelector("colgroup"); if (!cg) return;
  const col = cg.children[idx]; const startX = e.clientX; const startW = parseFloat(col.style.width) || col.offsetWidth;
  document.body.classList.add("col-resizing");
  const onMove = (ev) => {
    const w = Math.max(48, Math.round(startW + (ev.clientX - startX)));
    col.style.width = w + "px"; billColW[label] = w;
    let s = 0; for (const c of cg.children) s += parseFloat(c.style.width) || 0; table.style.width = s + "px";
  };
  const onUp = () => { document.removeEventListener("mousemove", onMove); document.removeEventListener("mouseup", onUp);
    document.body.classList.remove("col-resizing"); saveBillColWidths(); };
  document.addEventListener("mousemove", onMove); document.addEventListener("mouseup", onUp);
}
// Age helpers. bill_date is ISO (yyyy-mm-dd), so lexical order == chronological.
function billYm(b) { const m = String(b.bill_date || "").match(/^(\d{4})-(\d{2})/); return m ? (+m[1]) * 12 + (+m[2] - 1) : null; }
function billMonthsOld(b) { const ym = billYm(b); if (ym == null) return null; const n = new Date(); return (n.getFullYear() * 12 + n.getMonth()) - ym; }
const bOpen = b => num(b.open_balance);
const BILL_VIEWS = [   // quick presets - a base predicate the dropdowns then narrow
  { id: "open",     name: "Open AP",                  pred: b => bOpen(b) > 0 },
  { id: "paynow",   name: "GC-funded · unpaid · 2mo+", pred: b => b.approved === "approved" && b.invoice_status === "Invoice paid" && bOpen(b) > 0 && billMonthsOld(b) >= 2 },
  { id: "lien",     name: "Lien risk",                pred: b => BILL_LIEN_RISK.has(b.lien_status) && bOpen(b) > 0 },
  { id: "approve",  name: "To approve",               pred: b => b.approved !== "approved" && bOpen(b) > 0 },
  { id: "awaiting", name: "Awaiting invoice",         pred: b => b.invoice_status === "Awaiting Invoice" },
  { id: "noproj",   name: "No project #",             pred: b => !b.project_no || b.invoice_status === "No project #" },
  { id: "all",      name: "All bills",                pred: () => true },
];
const BILL_ROW_CAP = 2000;   // render ceiling (noted in-table when hit) - keeps the DOM snappy
function billView() { return BILL_VIEWS.find(v => v.id === activeBillView) || BILL_VIEWS[0]; }
function divClass(d) { const s = String(d || "").toUpperCase(); return s === "RP" ? "rp" : s === "CP" ? "cp" : s === "MFD" ? "mfd" : ""; }

// Compact colored status TEXT (Excel-legible, single line) - a <span> or null.
function stText(text, cls, title) { const s = document.createElement("span"); s.className = "st " + (cls || ""); s.textContent = text; if (title) s.title = title; return s; }
function payText(b) { const v = b.pay_status || ""; if (!v) return null;
  if (/unpaid/i.test(v)) return stText("Unpaid", "st-warn");
  if (/partial/i.test(v)) return stText("Partial", "st-warn", v);
  if (/paid/i.test(v)) return stText("Paid", "st-ok", v);   // fronted/collected in the tooltip
  return stText(v, "st-dim"); }
const BILL_INV_SHORT = { "Invoice paid": ["Inv paid", "st-ok"], "Awaiting Payment": ["Await pay", "st-warn"],
  "Awaiting Invoice": ["No invoice", "st-dim"], "No project #": ["No project", "st-bad"], "Partial paid": ["Partial", "st-warn"] };
function invText(b) { const v = b.invoice_status || ""; if (!v) return null; const m = BILL_INV_SHORT[v] || [v, "st-dim"]; return stText(m[0], m[1], v); }
function lienText(b) { const v = b.lien_status; if (!v || !LIEN_CLASS[v]) return null; return stText(LIEN_SHORT[v] || v, "st-lien-" + LIEN_CLASS[v], v); }
// Approved gets its OWN column (Yes/No) - not merged - so a blank never hides a missing value.
function apprText(b) { const v = b.approved || ""; if (!v) return null;
  return v === "approved" ? stText("Yes", "st-ok", "Approved for payment") : stText("No", "st-warn", "Not approved for payment"); }
// A dim placeholder for a genuinely empty status cell (so blank = "none for this bill", unambiguous).
function dimDash() { const s = document.createElement("span"); s.className = "st-none"; s.textContent = "–"; return s; }
function statusCell(node) { const td = document.createElement("td"); td.className = "left status-col"; td.appendChild(node || dimDash()); return td; }

// ── the six per-field filter dropdowns (each a component, not a search box) ──
// Categorical filters, all MULTI-select (checkboxes): empty selection = all; checked = show only
// those (owner 2026-08-20: "the same for all filters where multi select makes sense"). ONE generic
// component. Vendor + Month are bespoke (pump default / month cascade); Day stays a single drill.
const billMSel = {};   // { id: Set(selected raw values) }
const BILL_MSEL = [
  { id: "bfCustomer", all: "All clients",    get: b => b.client || "",        search: true, lbl: v => v || "(no client)" },
  { id: "bfProject",  all: "All projects",   get: b => b.project_no || "",    search: true, lbl: v => v || "(no project #)" },
  { id: "bfDivision", all: "All divisions",  get: b => b.division || "",       lbl: v => v || "(no division)" },
  { id: "bfPay",      all: "Any pay status", get: b => b.pay_status || "",     lbl: v => v || "(none)" },
  { id: "bfInv",      all: "Any invoice",    get: b => b.invoice_status || "", lbl: v => v || "(none)" },
  { id: "bfAppr",     all: "Any approval",   get: b => b.approved || "",       lbl: v => v === "approved" ? "Approved" : (v === "not approved" ? "Not approved" : (v || "(blank)")) },
  { id: "bfLien",     all: "Any lien",       get: b => b.lien_status || "",    lbl: v => v ? (LIEN_SHORT[v] || v) : "(no lien clock)" },
];
function _billMSelVals(cfg) { return [...new Set((BILLS || []).map(cfg.get))].sort((a, b) => cfg.lbl(a).localeCompare(cfg.lbl(b))); }
// Toggle updates the label IN PLACE (no rebuild) so an active search survives (owner 2026-08-21).
function toggleBillMSel(id, val, checked) {
  const s = billMSel[id] || (billMSel[id] = new Set());
  if (checked) s.add(val); else s.delete(val);
  _mselLabelUpdate(BILL_MSEL.find(c => c.id === id)); renderBills();
}
function _mselLabelUpdate(cfg) {
  if (!cfg) return;
  const btn = $("#" + cfg.id + "Btn"), menu = $("#" + cfg.id + "Menu"); if (!btn) return;
  const s = billMSel[cfg.id] || new Set();
  if (!s.size) btn.textContent = cfg.all;
  else if (s.size === 1) btn.textContent = cfg.lbl([...s][0]);
  else btn.textContent = s.size + " selected";
  btn.classList.toggle("on", s.size > 0);
  btn.title = s.size ? [...s].map(cfg.lbl).join(", ") : "";
  const cnt = menu ? menu.querySelector(".msel-count") : null; if (cnt) cnt.textContent = `${s.size} selected`;
}
// Select all / None over the VISIBLE (search-filtered) options, in place.
function _billMSelBulk(cfg, sel) {
  const menu = $("#" + cfg.id + "Menu"); if (!menu) return;
  const s = billMSel[cfg.id] || (billMSel[cfg.id] = new Set());
  for (const lab of menu.querySelectorAll(".msel-opt")) {
    if (lab.hidden) continue;
    const v = lab.dataset.val;
    if (sel) s.add(v); else s.delete(v);
    const cb = lab.querySelector("input"); if (cb) cb.checked = sel;
  }
  _mselLabelUpdate(cfg); renderBills();
}
function buildBillMSel(cfg) {
  const menu = $("#" + cfg.id + "Menu"), btn = $("#" + cfg.id + "Btn");
  if (!menu || !btn) return;
  const s = billMSel[cfg.id] || (billMSel[cfg.id] = new Set());
  const vals = _billMSelVals(cfg);
  for (const v of [...s]) if (!vals.includes(v)) s.delete(v);   // drop values gone from the data
  menu.innerHTML = "";
  if (cfg.search) { const q = document.createElement("input"); q.type = "search"; q.className = "msel-search"; q.placeholder = "Search";
    q.oninput = () => { const t = q.value.toLowerCase(); for (const lab of menu.querySelectorAll(".msel-opt")) lab.hidden = t && !lab.textContent.toLowerCase().includes(t); }; menu.appendChild(q);
    const tools = document.createElement("div"); tools.className = "msel-tools";
    const all = document.createElement("button"); all.type = "button"; all.className = "msel-tool"; all.textContent = "Select all"; all.title = "Select every option the search lists"; all.onclick = () => _billMSelBulk(cfg, true);
    const none = document.createElement("button"); none.type = "button"; none.className = "msel-tool"; none.textContent = "None"; none.title = "Deselect every option the search lists"; none.onclick = () => _billMSelBulk(cfg, false);
    const cnt = document.createElement("span"); cnt.className = "msel-count";
    tools.appendChild(all); tools.appendChild(none); tools.appendChild(cnt); menu.appendChild(tools); }
  { const clr = document.createElement("button"); clr.type = "button"; clr.className = "msel-clear"; clr.textContent = "Clear";
    clr.onclick = () => { s.clear(); buildBillMSel(cfg); renderBills(); }; menu.appendChild(clr); }
  for (const v of vals) {
    const lab = document.createElement("label"); lab.className = "msel-opt"; lab.dataset.val = v;
    const cb = document.createElement("input"); cb.type = "checkbox"; cb.checked = s.has(v);
    cb.onchange = () => toggleBillMSel(cfg.id, v, cb.checked);
    lab.appendChild(cb); lab.appendChild(document.createTextNode(" " + cfg.lbl(v)));
    menu.appendChild(lab);
  }
  _mselLabelUpdate(cfg);
}
function billMSelPasses(b) {
  for (const cfg of BILL_MSEL) { const s = billMSel[cfg.id]; if (s && s.size && !s.has(cfg.get(b))) return false; }
  return true;
}
function buildBillFilters() {
  for (const cfg of BILL_MSEL) buildBillMSel(cfg);
  if (!billDate) billDate = dateFilter("bfDate", () => (BILLS || []).map(b => b.bill_date), renderBills);
  billDate.build();
  buildBillVendorFilter();
}
// Month MULTI-select (checkboxes) + a day drill. Clicking a month checks it AND all OLDER
// months ("June and back"); individual priors can then be unchecked (owner 2026-08-20). The
// Day select drills into a single month (Excel-style), enabled only when exactly one is chosen.

// ── Date filter shared by Bills and Draws (owner 2026-09-02): two modes. MONTH = every month ticked
// by default with Select all / Deselect all; DATE = a from / to pair with the native calendar (pick a
// day or type it) - "as of July 25" is just a To date. Switching modes re-renders at once.
function dateFilter(id, getDates, onChange) {
  const st = { mode: "month", months: null, from: "", to: "" };   // months: null = all
  const host = $("#" + id); if (!host) return st;
  host.innerHTML = `<span class="seg tiny"><button type="button" class="seg-btn on" data-m="month">Month</button><button type="button" class="seg-btn" data-m="date">Date</button></span>
    <span class="datef-month msel" id="${id}Msel"><button type="button" class="msel-btn" id="${id}Btn">All months</button><div class="msel-menu" id="${id}Menu" hidden></div></span>
    <span class="datef-range" hidden><input type="date" id="${id}From" title="From (leave blank for no lower bound)"> <span class="dim">to</span> <input type="date" id="${id}To" title="To - a statement 'as of' a day is just this box"></span>
    <button type="button" class="btn small datef-reset" id="${id}Reset" title="Back to all dates" hidden>Reset</button>`;
  const btn = $("#" + id + "Btn"), menu = $("#" + id + "Menu"), from = $("#" + id + "From"), to = $("#" + id + "To");
  const paintMode = () => { host.querySelectorAll(".seg-btn").forEach(b => b.classList.toggle("on", b.dataset.m === st.mode));
    host.querySelector(".datef-month").hidden = st.mode !== "month"; host.querySelector(".datef-range").hidden = st.mode !== "date"; };
  host.querySelectorAll(".seg-btn").forEach(b => b.onclick = () => { st.mode = b.dataset.m; paintMode(); onChange(); });
  btn.onclick = (e) => { e.stopPropagation(); const open = menu.hidden; document.querySelectorAll(".msel-menu").forEach(m => m.hidden = true); menu.hidden = !open; if (open) _placeMenu(btn, menu); };
  from.onchange = () => { st.from = from.value; onChange(); }; to.onchange = () => { st.to = to.value; onChange(); };
  $("#" + id + "Reset").onclick = () => { st.clear(); onChange(); };
  st.build = () => {
    const asc = [...new Set(getDates().map(x => String(x || "").slice(0, 7)).filter(x => /^\d{4}-\d{2}$/.test(x)))].sort();
    if (st.months) for (const m of [...st.months]) if (!asc.includes(m)) st.months.delete(m);
    const sel = st.months === null ? new Set(asc) : st.months;
    menu.innerHTML = "";
    const tools = document.createElement("div"); tools.className = "msel-tools";
    const all = document.createElement("button"); all.type = "button"; all.className = "msel-tool"; all.textContent = "Select all"; all.onclick = () => { st.months = null; st.build(); onChange(); };
    const none = document.createElement("button"); none.type = "button"; none.className = "msel-tool"; none.textContent = "Deselect all"; none.onclick = () => { st.months = new Set(); st.build(); onChange(); };
    const cnt = document.createElement("span"); cnt.className = "msel-count"; cnt.textContent = `${sel.size} of ${asc.length}`;
    tools.appendChild(all); tools.appendChild(none); tools.appendChild(cnt); menu.appendChild(tools);
    for (const ym of [...asc].reverse()) {
      const lab = document.createElement("label"); lab.className = "msel-opt";
      const cb = document.createElement("input"); cb.type = "checkbox"; cb.checked = sel.has(ym);
      cb.onchange = () => { const s2 = new Set(st.months === null ? asc : st.months); if (cb.checked) s2.add(ym); else s2.delete(ym);
        st.months = (s2.size === asc.length || s2.size === 0) ? null : s2;   // unticking the last one goes back to ALL (owner 2026-09-02)
        st.build(); onChange(); };
      lab.appendChild(cb); lab.appendChild(document.createTextNode(" " + billMonthLabel(ym))); menu.appendChild(lab);
    }
    if (st.months === null) btn.textContent = "All months";
    else if (!sel.size) btn.textContent = "No months";
    else { const newest = [...sel].sort().reverse()[0]; btn.textContent = sel.size === 1 ? billMonthLabel(newest) : `${billMonthLabel(newest)} +${sel.size - 1}`; }
    btn.classList.toggle("on", st.months !== null);
    paintMode();
    { const rs = $("#" + id + "Reset"); if (rs) rs.hidden = !st.active(); }
  };
  st.passes = (dateStr) => {
    const ds = String(dateStr || "").slice(0, 10);
    if (st.mode === "date") { if (st.from && ds < st.from) return false; if (st.to && ds > st.to) return false; return true; }
    return st.months === null || st.months.has(ds.slice(0, 7));
  };
  st.active = () => st.mode === "date" ? !!(st.from || st.to) : st.months !== null;
  const _onChange = onChange; onChange = () => { st.build(); _onChange(); };   // every change repaints the control (Reset visibility, label)
  st.label = () => st.mode === "date" ? [st.from ? "from " + fmtDate(st.from) : "", st.to ? "to " + fmtDate(st.to) : ""].filter(Boolean).join(" ") : btn.textContent;
  st.clear = () => { st.months = null; st.from = ""; st.to = ""; from.value = ""; to.value = ""; st.build(); };
  return st;
}
let billDate = null, drawDate = null;
// Dropdown menus live inside cards that clip (`.widget { overflow: hidden }`), so an open menu is
// pinned to the viewport at its button instead - it can never be cut off by the card, and it gets
// as much height as the screen below (or above) the button allows. Closed on scroll / resize.
function _placeMenu(btn, menu) {
  const r = btn.getBoundingClientRect();
  menu.style.position = "fixed"; menu.style.top = ""; menu.style.bottom = ""; menu.style.left = Math.max(8, Math.min(r.left, window.innerWidth - menu.offsetWidth - 8)) + "px";
  const below = window.innerHeight - r.bottom - 12, above = r.top - 12;
  if (below >= 220 || below >= above) { menu.style.top = (r.bottom + 4) + "px"; menu.style.maxHeight = Math.max(160, Math.min(420, below)) + "px"; }
  else { menu.style.bottom = (window.innerHeight - r.top + 4) + "px"; menu.style.maxHeight = Math.max(160, Math.min(420, above)) + "px"; }
}
(function () {
  const closeAll = () => document.querySelectorAll(".msel-menu:not([hidden])").forEach(m => { m.hidden = true; });
  window.addEventListener("resize", closeAll);
  window.addEventListener("scroll", closeAll, true);   // any scrolling container - the pinned menu would drift otherwise
})();
let billMonths = new Set();   // (legacy - the Month/Date control above replaced the month click-older filter)
const _BMONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
function billMonthLabel(ym) { const [y, m] = ym.split("-"); return `${_BMONTHS[+m - 1]} ${y}`; }
function _billMonthsAsc() {   // every month present in the data, oldest → newest
  return [...new Set((BILLS || []).map(b => String(b.bill_date || "").slice(0, 7)).filter(s => /^\d{4}-\d{2}$/.test(s)))].sort();
}
// Vendor MULTI-select. The concrete-pump vendors are excluded BY DEFAULT (owner 2026-08-20) - the
// data stays, just filtered; check them back (or "Show all") to include them. Checked = shown.
let billVendorHidden = new Set();     // vendor names currently hidden
let billVendorDefault = new Set();    // the default-hidden set (the pumps) - to detect a non-default pick
let billVendorInit = false;
function _billVendors() { return [...new Set((BILLS || []).map(b => b.vendor).filter(Boolean))].sort((a, b) => a.localeCompare(b)); }
function _billPumpVendors() { return _billVendors().filter(v => /pump/i.test(v)); }
function _vendorNonDefault() {
  if (billVendorHidden.size !== billVendorDefault.size) return true;
  for (const v of billVendorHidden) if (!billVendorDefault.has(v)) return true;
  return false;
}
// A single checkbox toggle updates the label + count IN PLACE - no rebuild - so the search
// box and scroll position survive (owner 2026-08-21: checking a vendor must not reset the search).
function toggleBillVendor(v, checked) {
  if (checked) billVendorHidden.delete(v); else billVendorHidden.add(v);
  _vendorLabelUpdate(); renderBills();
}
function _vendorLabelUpdate() {
  const btn = $("#bfVendorBtn"), menu = $("#bfVendorMenu"); if (!btn) return;
  const vendors = _billVendors(); const shown = vendors.length - billVendorHidden.size;
  const isDefaultPumps = billVendorDefault.size > 0 && !_vendorNonDefault();
  if (!billVendorHidden.size) btn.textContent = "All vendors";
  else if (isDefaultPumps) btn.textContent = "All vendors except pumps";
  else btn.textContent = `${shown} of ${vendors.length} vendors`;
  btn.classList.toggle("on", _vendorNonDefault());
  btn.title = billVendorHidden.size ? "Hidden: " + [...billVendorHidden].join(", ") : "";
  const cnt = menu ? menu.querySelector(".msel-count") : null; if (cnt) cnt.textContent = `${shown} shown`;
}
// Select all / None act on the VISIBLE (search-filtered) options only, in place, so you can
// type a name, hit None on everything, then check just the few you want (owner 2026-08-21).
function _vendorBulk(show) {
  const menu = $("#bfVendorMenu"); if (!menu) return;
  for (const lab of menu.querySelectorAll(".msel-opt")) {
    if (lab.hidden) continue;
    const v = lab.dataset.vendor;
    if (show) billVendorHidden.delete(v); else billVendorHidden.add(v);
    const cb = lab.querySelector("input"); if (cb) cb.checked = show;
  }
  _vendorLabelUpdate(); renderBills();
}
function buildBillVendorFilter() {
  const menu = $("#bfVendorMenu"), btn = $("#bfVendorBtn");
  if (!menu || !btn) return;
  const vendors = _billVendors();
  if (!billVendorInit && vendors.length) {          // first build → default hides the pumps
    billVendorDefault = new Set(_billPumpVendors());
    billVendorHidden = new Set(billVendorDefault);
    billVendorInit = true;
  }
  for (const v of [...billVendorHidden]) if (!vendors.includes(v)) billVendorHidden.delete(v);   // drop vendors gone from data
  menu.innerHTML = "";
  { const s = document.createElement("input"); s.type = "search"; s.className = "msel-search"; s.placeholder = "Search vendors";
    s.oninput = () => { const q = s.value.toLowerCase(); for (const lab of menu.querySelectorAll(".msel-opt")) lab.hidden = q && !lab.textContent.toLowerCase().includes(q); }; menu.appendChild(s); }
  { const tools = document.createElement("div"); tools.className = "msel-tools";
    const all = document.createElement("button"); all.type = "button"; all.className = "msel-tool"; all.textContent = "Select all"; all.title = "Show every vendor the search lists"; all.onclick = () => _vendorBulk(true);
    const none = document.createElement("button"); none.type = "button"; none.className = "msel-tool"; none.textContent = "None"; none.title = "Hide every vendor the search lists - then check the few you want"; none.onclick = () => _vendorBulk(false);
    const cnt = document.createElement("span"); cnt.className = "msel-count";
    tools.appendChild(all); tools.appendChild(none); tools.appendChild(cnt); menu.appendChild(tools); }
  if (billVendorDefault.size) { const r = document.createElement("button"); r.type = "button"; r.className = "msel-clear"; r.textContent = "Reset to default (hide pumps)";
    r.onclick = () => { billVendorHidden = new Set(billVendorDefault); buildBillVendorFilter(); renderBills(); }; menu.appendChild(r); }
  for (const v of vendors) {
    const lab = document.createElement("label"); lab.className = "msel-opt"; lab.dataset.vendor = v;
    const cb = document.createElement("input"); cb.type = "checkbox"; cb.checked = !billVendorHidden.has(v);
    cb.onchange = () => toggleBillVendor(v, cb.checked);
    lab.appendChild(cb); lab.appendChild(document.createTextNode(" " + v));
    menu.appendChild(lab);
  }
  _vendorLabelUpdate();
}
function billFilterValues() {
  const f = {};
  f["#bfDate"] = billDate && billDate.active() ? "1" : "";
  f["#bfVendor"] = _vendorNonDefault() ? "1" : "";                                   // vendor deviates from the pump default
  f["#bfMSel"] = BILL_MSEL.some(c => (billMSel[c.id] || {}).size) ? "1" : "";        // any categorical multi-select active
  return f;                                                                          // (drives the "Clear filters" button)
}
function billPassesFilters(b, f) {
  if (billDate && !billDate.passes(b.bill_date)) return false;   // Month | Date (from / to)
  if (billVendorHidden.has(b.vendor || "")) return false;      // vendor multi-select (pumps hidden by default)
  if (!billMSelPasses(b)) return false;                        // Client / Division / Pay / Invoice / Approved / Lien
  return true;
}
function billClearFilters() {
  for (const cfg of BILL_MSEL) (billMSel[cfg.id] || (billMSel[cfg.id] = new Set())).clear();
  if (billDate) billDate.clear();
  billVendorHidden = new Set(billVendorDefault);   // back to the default (pumps hidden), not "show everything"
  buildBillFilters();
  renderBills();
}
// Sort comparators. Rows are sorted BEFORE grouping, so within each group the order
// holds (default oldest→newest); groups themselves render alphabetically (A→Z).
function billLienRank(b) { const i = LIEN_ORDER.indexOf(b.lien_status); return i < 0 ? 99 : i; }
const BILL_SORTS = {
  oldest: (a, b) => String(a.bill_date || "9999").localeCompare(String(b.bill_date || "9999")),
  newest: (a, b) => String(b.bill_date || "").localeCompare(String(a.bill_date || "")),
  vendor: (a, b) => (a.vendor || "").localeCompare(b.vendor || "") || String(a.bill_date || "").localeCompare(String(b.bill_date || "")),
  owed:   (a, b) => bOpen(b) - bOpen(a),
  amount: (a, b) => num(b.line_amount) - num(a.line_amount),
  lien:   (a, b) => (billLienRank(a) - billLienRank(b)) || String(a.bill_date || "").localeCompare(String(b.bill_date || "")),
};

function renderBills() {
  const bills = BILLS || [];
  const vc = $("#billViews"); if (!vc) return;
  renderBillSaveBar();
  if (!$("#bfVendor") || !$("#bfVendor").options.length) buildBillFilters();
  // quick-preset chips with live counts
  vc.innerHTML = "";
  for (const v of BILL_VIEWS) {
    const n = bills.filter(v.pred).length;
    const chip = document.createElement("button");
    chip.className = "view-chip" + (v.id === activeBillView ? " active" : "");
    const nm = document.createElement("span"); nm.className = "vc-name"; nm.textContent = v.name;
    const ct = document.createElement("span"); ct.className = "vc-count"; ct.textContent = String(n);
    chip.appendChild(nm); chip.appendChild(ct);
    chip.onclick = () => { activeBillView = v.id; try { localStorage.setItem("proficient-ledger-billview", v.id); } catch { /* ignore */ } renderBills(); };
    vc.appendChild(chip);
  }
  const view = billView();
  const f = billFilterValues();
  { const cb = $("#bfClear"); if (cb) cb.hidden = !Object.values(f).some(x => x); }

  // filter (view predicate AND every dropdown), then sort
  let rows = bills.filter(b => view.pred(b) && billPassesFilters(b, f));
  const sortKey = $("#billSort") ? $("#billSort").value : "oldest";
  rows = [...rows].sort(BILL_SORTS[sortKey] || BILL_SORTS.oldest);

  const openSum = rows.reduce((t, b) => t + bOpen(b), 0);
  const lienN = rows.filter(b => BILL_LIEN_RISK.has(b.lien_status)).length;
  $("#billsNote").textContent = bills.length ? `(${rows.length.toLocaleString()} of ${bills.length.toLocaleString()})` : "(no AP data - run load_bill_tracker.py)";
  { const qs = $("#billsQuickStat"); if (qs) qs.textContent = bills.length ? `${money(openSum)} open · ${lienN} lien risk` : ""; }
  { const pr = $("#btnPayRunGo"); if (pr) { const n = (BILLS || []).filter(b => b.pay_selected).length; pr.textContent = n ? `Pay run (${n}) →` : "Pay run →"; pr.classList.toggle("on", n > 0); } }

  // table. Each status is its OWN column (Paid / Invoice / Lien / Appr) so a blank in
  // one never hides a missing value by being merged with the others.
  const group = $("#billGroup") ? $("#billGroup").value : "vendor";
  const thead = $("#billTable thead"), tbody = $("#billTable tbody");
  thead.innerHTML = "";
  thead.hidden = false; tbody.innerHTML = "";
  const cols = [["Vendor", "left", ""], ["Project", "left", ""], ["Bill #", "left", "Bill number - opens the bill in QuickBooks"],
                ["Date", "left", "Bill date (MM/DD/YY)"], ["Amount", "right", "Bill amount"], ["Open", "right", "Open balance we still owe"],
                ["Paid", "left", "Did we pay the vendor?"], ["Invoice", "left", "Was the AR invoice (draw) paid by the GC?"],
                ["Lien", "left", "Texas lien-notice clock"], ["Appr", "left", "Approved for payment?"]];
  // Fixed layout + a <colgroup> so column widths are exact and draggable; each header
  // carries a resize grip on its right divider, and the table width tracks the sum.
  const table = $("#billTable");
  { const oldCg = table.querySelector("colgroup"); if (oldCg) oldCg.remove(); }
  const colgroup = document.createElement("colgroup");
  const htr = document.createElement("tr");
  let wsum = 0;
  cols.forEach(([c, al, tip], i) => {
    const th = document.createElement("th"); if (al === "left") th.className = "left"; th.textContent = c; if (tip) th.title = tip;
    const grip = document.createElement("div"); grip.className = "col-resize"; grip.title = "Drag to resize this column";
    grip.addEventListener("mousedown", (e) => startBillColResize(e, i, c));
    th.appendChild(grip); htr.appendChild(th);
    const w = Math.max(48, billColW[c] || 100); const col = document.createElement("col"); col.style.width = w + "px";
    colgroup.appendChild(col); wsum += w;
  });
  table.insertBefore(colgroup, table.firstChild);
  table.style.width = wsum + "px";
  thead.appendChild(htr);

  if (!rows.length) {
    const tr = document.createElement("tr"); const td = document.createElement("td");
    td.colSpan = cols.length; td.className = "left"; td.style.color = "var(--text-dim)"; td.style.padding = "14px 12px";
    if (!bills.length) td.textContent = "No AP data - run load_bill_tracker.py.";
    else {   // say what is hiding them: the view pill (e.g. Open AP hides paid bills) is easy to miss (owner 2026-09-02)
      const byFilters = bills.filter(b => billPassesFilters(b, f)).length;
      td.textContent = byFilters ? `No bills match - ${byFilters} bill${byFilters === 1 ? "" : "s"} pass${byFilters === 1 ? "es" : ""} the filters but ${byFilters === 1 ? "is" : "are"} hidden by the "${view.name}" view above. Pick "All bills" to see ${byFilters === 1 ? "it" : "them"}.` : "No bills match these filters.";
    }
    tr.appendChild(td); tbody.appendChild(tr); billGroupKeys = []; updateBillCollapseBtn(group); return;
  }

  let rendered = 0, capped = false;
  const pushRow = b => { if (rendered >= BILL_ROW_CAP) { capped = true; return false; } tbody.appendChild(billRow(b, cols.length)); rendered++; return true; };
  if (group === "none") {
    billGroupKeys = [];
    for (const b of rows) if (!pushRow(b)) break;
  } else {
    const groups = new Map();
    for (const b of rows) { const k = billGroupKey(b, group); if (!groups.has(k)) groups.set(k, []); groups.get(k).push(b); }
    const order = [...groups.keys()].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));   // A→Z
    billGroupKeys = order;
    thead.hidden = order.length > 0 && order.every(k => billsCollapsed.has(k));   // headers only once a group is open (owner 2026-09-01)
    outer:
    for (const k of order) {
      const g = groups.get(k);
      const collapsed = billsCollapsed.has(k);
      const gOpen = g.reduce((t, x) => t + bOpen(x), 0);
      const gtr = document.createElement("tr"); gtr.className = "bill-group"; gtr.style.cursor = "pointer";
      gtr.title = collapsed ? "Click to expand" : "Click to collapse";
      const gtd = document.createElement("td"); gtd.colSpan = cols.length;
      // Flex lives on an inner div, NOT the td: display:flex on a <td> drops table-cell
      // layout and the colspan collapses to content width.
      const cell = document.createElement("div"); cell.className = "bg-cell";
      const left = document.createElement("span"); left.className = "bg-left";
      const caret = document.createElement("span"); caret.className = "bg-caret"; caret.textContent = collapsed ? "▸" : "▾";
      const key = document.createElement("span"); key.className = "bg-key"; key.textContent = billGroupLabel(k, group);
      left.appendChild(caret); left.appendChild(key);
      // Open $ + bill count at the SAME size/weight as the vendor (owner 2026-08-18) so the
      // amount is scannable down the collapsed list; right-aligned in the row.
      cell.appendChild(left);
      bandMetrics(cell, [[money(gOpen), "open", gOpen > 0.005 ? "neg" : ""], [g.length, "bills"], [g.filter(b => b.pay_date).length, "paid"], [g.filter(b => BILL_LIEN_RISK.has(b.lien_status)).length || "–", "lien risk"]]);
      gtd.appendChild(cell); gtr.appendChild(gtd);
      gtr.onclick = () => { if (billsCollapsed.has(k)) billsCollapsed.delete(k); else billsCollapsed.add(k); renderBills(); };
      tbody.appendChild(gtr);
      if (!collapsed) for (const b of g) if (!pushRow(b)) break outer;
    }
  }
  updateBillCollapseBtn(group);
  // Only when the 2000-row CAP actually truncated the render - never for rows merely
  // hidden by a collapsed group (else "Showing 0 of N" fires on the collapse-by-default view).
  if (capped) {
    const tr = document.createElement("tr"); const td = document.createElement("td");
    td.colSpan = cols.length; td.className = "left"; td.style.color = "var(--text-dim)"; td.style.padding = "10px 12px";
    td.textContent = `Showing ${rendered.toLocaleString()} of ${rows.length.toLocaleString()} - narrow with a filter. (Open total above covers all ${rows.length.toLocaleString()}.)`;
    tr.appendChild(td); tbody.appendChild(tr);
  }
}
function billGroupKey(b, group) {
  if (group === "division") return b.division || "–";
  if (group === "project_no") return b.project_no || "–";
  if (group === "vendor") return b.vendor || "–";
  if (group === "client") return b.client || "–";
  if (group === "matched_invoice") return b.invoice_no || b.matched_invoice || "–";
  return "–";
}
function billGroupLabel(k, group) {
  if (group === "project_no" && k !== "–") { const nm = nameOf(k); return nm ? `${k} · ${nm}` : k; }
  if (group === "client") return k === "–" ? "No client on file" : k;
  if (group === "division") return k === "–" ? "No division" : k;
  if (group === "matched_invoice") return k === "–" ? "No draw" : "Draw " + k;
  return k;
}
function billRow(b) {
  const tr = document.createElement("tr");
  tr.className = "bill-row" + (BILL_LIEN_ACTIVE.has(b.lien_status) ? " risk" : "");
  tr.style.cursor = "pointer";
  // Click the row (not a link / not a selectable money cell) → the invoice slides in on the right.
  tr.onclick = (e) => { if (e.target.closest(".cell") || e.target.closest("a")) return; openBillDetail(b); };
  // Vendor
  const vtd = document.createElement("td"); vtd.className = "left";
  const vs = document.createElement("span"); vs.className = "bill-vendor"; vs.textContent = b.vendor || "–"; vs.title = b.vendor || "";
  vtd.appendChild(vs); tr.appendChild(vtd);
  // Project (division chip + CLIENT - easier to scan than the job name; job name is in the tooltip)
  const ptd = document.createElement("td"); ptd.className = "left";
  if (b.project_no) {
    const chip = document.createElement("span"); const dc = divClass(b.division);
    chip.className = "divchip" + (dc ? " " + dc : ""); chip.textContent = b.project_no; ptd.appendChild(chip);
    const nm = nameOf(b.project_no); const disp = b.client || nm;
    if (disp) { const s = document.createElement("span"); s.className = "bill-name"; s.textContent = disp;
      s.title = b.client ? (nm ? `${b.client} · ${nm}` : b.client) : nm; ptd.appendChild(s); }
  } else { ptd.appendChild(document.createTextNode("–")); }
  tr.appendChild(ptd);
  // Bill # (QBO deep link)
  tr.appendChild(qboLinkCell(b.bill_ref, qboBillHref(b.qbo_link), "Open this bill in QuickBooks"));
  // Date (MM/DD/YY) + age badge once a bill is 2+ months old
  const dtd = document.createElement("td"); dtd.className = "left bill-date";
  const ds = document.createElement("span"); ds.textContent = fmtDateShort(b.bill_date); ds.title = fmtDate(b.bill_date); dtd.appendChild(ds);
  if (String(b.bill_date || "").slice(0, 10) > new Date(Date.now() - new Date().getTimezoneOffset() * 60000).toISOString().slice(0, 10)) {
    const w = document.createElement("span"); w.className = "vg-tag future"; w.textContent = "future date"; w.title = "This bill is dated after today - almost certainly a typo on the bill date in QuickBooks (wrong year)"; dtd.appendChild(w); }
  const mo = billMonthsOld(b);
  if (mo != null && mo >= 2) { const a = document.createElement("span"); a.className = "bill-age"; a.textContent = mo + "mo"; dtd.appendChild(a); }
  tr.appendChild(dtd);
  // Amount
  const atd = document.createElement("td"); atd.appendChild(moneyCell(b.line_amount)); tr.appendChild(atd);
  // Open balance
  const otd = document.createElement("td"); const oc = moneyCell(b.open_balance);
  oc.classList.add(bOpen(b) > 0 ? "open-amt" : "open-zero"); otd.appendChild(oc); tr.appendChild(otd);
  // Four SEPARATE status columns - Paid / Invoice / Lien / Appr (blank = none, unambiguous)
  tr.appendChild(statusCell(payText(b)));
  tr.appendChild(statusCell(invText(b)));
  tr.appendChild(statusCell(lienText(b)));
  tr.appendChild(statusCell(apprText(b)));
  return tr;
}
// Collapse/expand-all button label + visibility (grouped views only).
function updateBillCollapseBtn(group) {
  const cb = $("#bfCollapse"); if (!cb) return;
  cb.hidden = (group === "none") || !billGroupKeys.length;
  const allC = billGroupKeys.length && billGroupKeys.every(k => billsCollapsed.has(k));
  cb.textContent = allC ? "Expand all" : "Collapse all";
}
function billToggleAll() {
  const allC = billGroupKeys.length && billGroupKeys.every(k => billsCollapsed.has(k));
  if (allC) billsCollapsed.clear(); else billGroupKeys.forEach(k => billsCollapsed.add(k));
  renderBills();
}
// Match a lien-worklist row back to its full bill in BILLS (which carries pay status,
// approval, and the joined invoice fields) so a lien row can open the same rich panel.
// Prefer the QBO bill id (unique); fall back to vendor + bill # + invoice #.
function findBillForLien(r) {
  const rt = String(r.qbo_link || "").match(/txnId=(\d+)/i);
  return (BILLS || []).find(b => {
    const bt = String(b.qbo_link || "").match(/txnId=(\d+)/i);
    if (bt && rt) return bt[1] === rt[1];
    return (b.vendor || "") === (r.vendor || "") && (b.bill_ref || "") === (r.bill_ref || "")
        && (b.invoice_no || "") === (r.invoice_no || "");
  });
}
// From the Vendors spend tab → the Bills tab, pre-filtered to that vendor (all their bills).
function jumpToVendorBills(vendor) {
  activeBillView = "all";
  setTab("bills");
  renderBills();                              // ensure the filter menus are populated
  billVendorHidden = new Set(_billVendors().filter(v => v !== vendor));   // show ONLY this vendor (exclude the rest)
  buildBillVendorFilter();
  // Expand this vendor's group(s) so their bills show immediately - no extra click (owner 2026-08-25).
  const grp = $("#billGroup") ? $("#billGroup").value : "vendor";
  for (const b of (BILLS || [])) if ((b.vendor || "") === vendor) billsCollapsed.delete(billGroupKey(b, grp));
  renderBills();
  window.scrollTo(0, 0);
}

// Vendor page (QBO-style, ON DEMAND) - one vendor's bills, fetched per vendor via /api/vendor (never
// in the bulk load). Each bill shows its project, or "multiple" -> click the bill to see every line
// item + project #. Filter by pay status. Owner 2026-08-28: "vendor center open into its own vendor
// page like qbo ... see the bill its paying and the project ... if multiple say multiple, click for lines".
let _vendorData = null, _vendorType = "all", _vendorView = "bills";   // bills | payments
const _vendorBillOpen = new Set();
async function openVendorPage(vendor) {
  openRecord(vendor, "loading…");
  const body = $("#recordBody"); body.innerHTML = "";
  _vendorData = null; _vendorType = "all"; _vendorView = "bills"; _vendorBillOpen.clear();
  let data;
  try { data = await (await fetch("/api/vendor?v=" + encodeURIComponent(vendor))).json(); }
  catch (e) { body.textContent = "could not load this vendor"; return; }
  if (!data || !data.ok) { body.textContent = (data && data.error) || "no data for this vendor"; return; }
  _vendorData = data;
  renderVendorPage();
}
function renderVendorPage() {
  const d = _vendorData; if (!d) return;
  const body = $("#recordBody"); body.innerHTML = "";
  // Bills | Payments view toggle
  const vseg = document.createElement("div"); vseg.className = "seg vendor-seg";
  for (const [k, lbl] of [["bills", `Bills (${d.count})`], ["payments", `Payments (${d.pay_count || 0})`]]) {
    const b = document.createElement("button"); b.type = "button"; b.className = "seg-btn" + (_vendorView === k ? " on" : ""); b.textContent = lbl;
    b.onclick = () => { _vendorView = k; renderVendorPage(); }; vseg.appendChild(b);
  }
  body.appendChild(vseg);
  if (_vendorView === "payments") {
    $("#recordSub").textContent = `${d.pay_count || 0} payments · ${money(d.pay_total || 0)} paid out this year`;
    return _renderVendorPayments(d, body);
  }
  // Two systems, two labels (owner 2026-09-02: the list said one "open", the page another): QuickBooks
  // open AP covers every vendor incl. subs; the Bill Tracker excludes subs.
  $("#recordSub").textContent = (d.qbo_open != null ? `open ${money(d.qbo_open)} (QuickBooks, ${d.qbo_open_bills || 0} bills) · ` : "")
    + `Bill Tracker: ${d.count} bills · ${money(d.total)} billed · ${money(d.open)} open · ${d.paid_ct} paid` + (d.count ? "" : " (subs are not in the Bill Tracker)");
  const seg = document.createElement("div"); seg.className = "seg vendor-seg";
  for (const [k, lbl] of [["all", "All"], ["open", "Open"], ["paid", "Paid"]]) {
    const b = document.createElement("button"); b.type = "button"; b.className = "seg-btn" + (_vendorType === k ? " on" : ""); b.textContent = lbl;
    b.onclick = () => { _vendorType = k; renderVendorPage(); }; seg.appendChild(b);
  }
  body.appendChild(seg);
  const bills = (d.bills || []).filter(b => _vendorType === "all" || (_vendorType === "paid" ? b.paid : !b.paid));
  if (!(d.bills || []).length && (d.qbo_bills || []).length) return _renderVendorQboBills(d, body);   // a sub: show its QBO bills instead
  if (!bills.length) { const p = document.createElement("div"); p.className = "bills-cap"; p.textContent = "No bills match this filter."; body.appendChild(p); return; }
  const scroll = document.createElement("div"); scroll.className = "table-scroll";
  const table = document.createElement("table"); table.className = "grid"; const thead = document.createElement("thead"), tbody = document.createElement("tbody");
  const htr = document.createElement("tr");
  for (const [c, al] of [["", "left"], ["Date", "left"], ["Bill #", "left"], ["Project", "left"], ["Amount", "right"], ["Status", "left"]]) { const th = document.createElement("th"); if (al === "left") th.className = "left"; th.textContent = c; htr.appendChild(th); }
  thead.appendChild(htr);
  bills.forEach((b, i) => {
    const key = (b.bill_ref || "?") + "|" + (b.bill_date || "") + "|" + i;
    const open = _vendorBillOpen.has(key), multi = b.lines.length > 1;
    const tr = document.createElement("tr"); tr.className = "vp-bill"; if (multi) tr.style.cursor = "pointer";
    tr.onclick = (e) => { if (e.target.closest("a") || e.target.closest(".refcopy") || e.target.closest(".cell")) return; if (!multi) return; open ? _vendorBillOpen.delete(key) : _vendorBillOpen.add(key); renderVendorPage(); };
    const cc = document.createElement("td"); cc.className = "left draw-caret"; cc.textContent = multi ? (open ? "▾" : "▸") : ""; tr.appendChild(cc);
    tr.appendChild(leftText(fmtDateShort(b.bill_date)));
    tr.appendChild(qboLinkCell(b.bill_ref, qboBillHref(b.qbo_link), "Open this bill in QuickBooks"));
    const pcell = document.createElement("td"); pcell.className = "left";
    if (b.project === "multiple") { const s = document.createElement("span"); s.className = "vp-multi"; s.textContent = "multiple"; s.title = "Click the bill to see all line items + project #s"; pcell.appendChild(s); }
    else pcell.appendChild(document.createTextNode(b.project || "–"));
    tr.appendChild(pcell);
    const amt = document.createElement("td"); amt.appendChild(moneyCell(b.amount)); tr.appendChild(amt);
    const st = document.createElement("td"); st.className = "left";
    const stp = document.createElement("span"); stp.className = b.paid ? "ar-paid" : "ar-open"; stp.textContent = b.paid ? ("Paid " + fmtDateShort(b.pay_date)) : (b.pay_status || "Open"); st.appendChild(stp); tr.appendChild(st);
    tbody.appendChild(tr);
    if (multi && open) { const lr = document.createElement("tr"); const ltd = document.createElement("td"); ltd.colSpan = 6; ltd.appendChild(_vendorLines(b)); lr.appendChild(ltd); tbody.appendChild(lr); }
  });
  table.appendChild(thead); table.appendChild(tbody); scroll.appendChild(table); body.appendChild(scroll);
}
// A vendor with no Bill Tracker rows (a sub, or a vendor the tracker doesn't carry): its bills straight
// from the QBO cost lines already in the ledger - date, bill #, project(s), memo, amount, qb link.
function _renderVendorQboBills(d, body) {
  const cap = document.createElement("div"); cap.className = "bills-cap";
  cap.textContent = `${d.qbo_bills.length} bills from QuickBooks (job-costed lines) - this vendor has no Bill Tracker rows, so pay status comes from QuickBooks' open balance above.`;
  body.appendChild(cap);
  const scroll = document.createElement("div"); scroll.className = "table-scroll";
  const table = document.createElement("table"); table.className = "grid"; const thead = document.createElement("thead"), tbody = document.createElement("tbody");
  const htr = document.createElement("tr");
  for (const [c, al] of [["Date", "left"], ["Bill #", "left"], ["Project", "left"], ["Memo", "left"], ["Amount", "right"]]) { const th = document.createElement("th"); if (al === "left") th.className = "left"; th.textContent = c; htr.appendChild(th); }
  thead.appendChild(htr);
  for (const b of d.qbo_bills) {
    const tr = document.createElement("tr");
    tr.appendChild(leftText(fmtDateShort(b.date)));
    tr.appendChild(qboLinkCell(b.doc_number || "–", qboUrl(b.txn_type === "Expense" ? "expense" : "bill", b.txn_id), "Open this bill in QuickBooks"));
    tr.appendChild(leftText(b.projects.length > 1 ? b.projects.join(", ") : (b.projects[0] || "–")));
    const m = leftText(b.memo || "–"); m.title = b.memo || ""; m.className += " inv-memo"; tr.appendChild(m);
    const amt = document.createElement("td"); amt.appendChild(moneyCell(b.amount)); tr.appendChild(amt);
    tbody.appendChild(tr);
  }
  table.appendChild(thead); table.appendChild(tbody); scroll.appendChild(table); body.appendChild(scroll);
}
function _vendorLines(b) {
  const wrap = document.createElement("div"); wrap.className = "bills-sub";
  const cap = document.createElement("div"); cap.className = "bills-cap"; cap.textContent = `${b.lines.length} line items on bill ${b.bill_ref || ""}${b.projects.length > 1 ? " · " + b.projects.filter(p => p !== "(multiple)").join(", ") : ""}`; wrap.appendChild(cap);
  const scroll = document.createElement("div"); scroll.className = "table-scroll";
  const table = document.createElement("table"); table.className = "grid"; const thead = document.createElement("thead"), tbody = document.createElement("tbody");
  const htr = document.createElement("tr");
  for (const [c, al] of [["Project #", "left"], ["Description", "left"], ["Amount", "right"]]) { const th = document.createElement("th"); if (al === "left") th.className = "left"; th.textContent = c; htr.appendChild(th); }
  thead.appendChild(htr);
  for (const ln of b.lines) { const tr = document.createElement("tr");
    tr.appendChild(leftText(ln.project_no || "–"));
    tr.appendChild(leftText(ln.description || "–"));
    tr.appendChild(rightText(money(ln.amount)));
    tbody.appendChild(tr);
  }
  table.appendChild(thead); table.appendChild(tbody); scroll.appendChild(table); wrap.appendChild(scroll);
  return wrap;
}
// Vendor payments view: the QBO BillPayments (money out) this year, from the local bill_payment table.
function _renderVendorPayments(d, body) {
  const pays = d.payments || [];
  if (!pays.length) { const p = document.createElement("div"); p.className = "bills-cap"; p.textContent = "No bill payments recorded this year (run the AP / bill-payments sync to pull them)."; body.appendChild(p); return; }
  const scroll = document.createElement("div"); scroll.className = "table-scroll";
  const table = document.createElement("table"); table.className = "grid"; const thead = document.createElement("thead"), tbody = document.createElement("tbody");
  const htr = document.createElement("tr");
  for (const [c, al] of [["Date", "left"], ["Ref / cheque #", "left"], ["Type", "left"], ["Bills paid", "right"], ["Amount", "right"]]) { const th = document.createElement("th"); if (al === "left") th.className = "left"; th.textContent = c; htr.appendChild(th); }
  thead.appendChild(htr);
  for (const p of pays) {
    const tr = document.createElement("tr");
    tr.appendChild(leftText(fmtDateShort(p.txn_date)));
    tr.appendChild(qboLinkCell(p.ref_no || "–", null, ""));   // ref # copyable
    tr.appendChild(leftText(p.pay_type || "–"));
    tr.appendChild(rightText(String(p.n_bills || 0)));
    const amt = document.createElement("td"); amt.appendChild(moneyCell(p.total_amt)); tr.appendChild(amt);
    tbody.appendChild(tr);
  }
  table.appendChild(thead); table.appendChild(tbody); scroll.appendChild(table); body.appendChild(scroll);
}
// Lien marks are STAGED, then Saved (the owner marks several, then commits once). A mark
// updates the panel + grid optimistically and shows the Save bar; nothing is written until
// Save. Leaving the page (or Discard) is guarded so a marking session is never lost.
const pendingBillMarks = new Map();   // bill_id -> {lien, prevLien, prevMarked}
function setBillLien(b, lien) {
  if (!b.bill_id) { toast("This bill has no QBO bill link - can't mark it."); return; }
  if (!pendingBillMarks.has(b.bill_id)) pendingBillMarks.set(b.bill_id, { prevLien: b.lien_status, prevMarked: b.lien_marked });
  pendingBillMarks.get(b.bill_id).lien = lien;
  b.lien_status = lien || "";        // optimistic; Save reconciles the computed value on clear
  b.lien_marked = !!lien;
  renderBillSaveBar(); renderBills(); openBillDetail(b);
}
function renderBillSaveBar() {
  const bar = $("#billSaveBar"); if (!bar) return;
  const n = pendingBillMarks.size;
  bar.classList.toggle("dirty", n > 0);
  bar.hidden = n === 0;                   // only while something is unsaved (owner 2026-09-02: "why is this here?"); saved marks are reviewed from the button by Clear filters
  $("#billSaveText").textContent = n ? `${n} unsaved lien mark${n > 1 ? "s" : ""} - review` : "";
  { const rv = $("#btnLienReview"); if (rv) rv.hidden = n > 0; }
  { const b = $("#btnSaveBillMarks"); if (b) b.disabled = n === 0; }
  { const d = $("#btnDiscardBillMarks"); if (d) d.hidden = n === 0; }
}
async function saveBillMarks() {
  const entries = [...pendingBillMarks.entries()]; if (!entries.length) return;
  const btn = $("#btnSaveBillMarks"); if (btn) { btn.disabled = true; btn.textContent = "Saving…"; }
  try {
    for (const [bill_id, p] of entries) {
      const bx = (BILLS || []).find(x => x.bill_id === bill_id);      // vendor → auto-create its lien folder on save
      const res = await fetch("/api/bill-mark", { method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ bill_id, lien: p.lien || "", vendor: bx ? (bx.vendor || "") : "" }) });
      const j = await res.json(); if (!j.ok) throw new Error(j.error || "write failed");
    }
    pendingBillMarks.clear();
    toast(`Saved ${entries.length} lien mark${entries.length > 1 ? "s" : ""}`);
    await load(true);                                       // authoritative merged state
  } catch (e) { toast("Save failed: " + e.message); }
  finally { if (btn) { btn.disabled = false; btn.textContent = "Save"; } renderBillSaveBar(); }
}
function discardBillMarks() {
  for (const [bill_id, p] of pendingBillMarks) { const b = (BILLS || []).find(x => x.bill_id === bill_id);
    if (b) { b.lien_status = p.prevLien; b.lien_marked = p.prevMarked; } }
  pendingBillMarks.clear(); renderBillSaveBar(); renderBills();
  const bd = $("#billDetail"); if (bd && !bd.hidden) closePanels();
  toast("Discarded unsaved marks");
}

// ══════════════ PAY BILLS (a dedicated check-run worksheet) ══════════════
// Mark bills for a check run, set a partial amount, Save → generate the pay list.
// It records INTENT only - it never pays QuickBooks or moves money (the owner records
// the real payment in QBO; the bill clears here on the next AP sync). Kept off the Bills
// tab on purpose (owner 2026-08-21: that tab is too crowded / an accidental click risk).
// State: paySaved = the server's current run ({bill_id -> amount|null}); payDraft = edits
// since load (only touched bills). Effective = draft wins, else saved, else unselected.
let paySaved = new Map();   // server truth: {bill_id -> amount|null}, recomputed from BILLS each render
let payDraft = new Map();   // unsaved overlay: {bill_id -> {selected, amount}} for touched bills
const PAY_ROW_CAP = 1500;

function _payMarkable() { return (BILLS || []).filter(b => b.bill_id); }   // has a QBO bill id → markable
function _payRecomputeSaved() {   // mirror the current server run from the latest BILLS (idempotent, cheap)
  paySaved = new Map();
  for (const b of _payMarkable()) if (b.pay_selected) paySaved.set(b.bill_id, b.pay_amount == null ? null : num(b.pay_amount));
}
function payState(b) {                    // effective {selected, amount} for a bill
  if (payDraft.has(b.bill_id)) return payDraft.get(b.bill_id);
  if (paySaved.has(b.bill_id)) return { selected: true, amount: paySaved.get(b.bill_id) };
  return { selected: false, amount: null };
}
function payAmountOf(b) { const a = payState(b).amount; return a == null ? num(b.open_balance) : num(a); }
function paySelectedBills() { return _payMarkable().filter(b => payState(b).selected && payAmountOf(b) > 0); }
function payDirtyCount() {
  let n = 0;
  for (const [id, d] of payDraft) {
    const savedSel = paySaved.has(id);
    const savedAmt = savedSel ? paySaved.get(id) : null;
    const dAmt = d.amount == null ? null : num(d.amount);
    const same = (!!d.selected === savedSel) && (!d.selected || dAmt === (savedAmt == null ? null : num(savedAmt)));
    if (!same) n++;
  }
  return n;
}
// ── Generic multi-select checklist (search + Select all/None; a toggle updates in place so
// the search box survives). The caller owns a `store` ({id -> Set}) and passes an onChange
// render callback. cfg: { id, all, get, lbl, search }. DOM: `${id}Btn` pill + `${id}Menu` panel.
// (Pay Bills uses this; the older Bills/Liens builders predate it and stay as they are.)
function _mselVals(items, cfg) { return [...new Set(items.map(cfg.get))].sort((a, b) => cfg.lbl(a).localeCompare(cfg.lbl(b))); }
function mselLabelUpdate(cfg, store) {
  const btn = $("#" + cfg.id + "Btn"), menu = $("#" + cfg.id + "Menu"); if (!btn) return;
  const s = store[cfg.id] || new Set();
  btn.textContent = !s.size ? cfg.all : (s.size === 1 ? cfg.lbl([...s][0]) : s.size + " selected");
  btn.classList.toggle("on", s.size > 0);
  btn.title = s.size ? [...s].map(cfg.lbl).join(", ") : "";
  const cnt = menu ? menu.querySelector(".msel-count") : null; if (cnt) cnt.textContent = `${s.size} selected`;
}
function mselBulk(cfg, store, sel, onChange) {   // Select all / None over the VISIBLE (search-filtered) options
  const menu = $("#" + cfg.id + "Menu"); if (!menu) return;
  const s = store[cfg.id] || (store[cfg.id] = new Set());
  for (const lab of menu.querySelectorAll(".msel-opt")) {
    if (lab.hidden) continue; const v = lab.dataset.val;
    if (sel) s.add(v); else s.delete(v);
    const cb = lab.querySelector("input"); if (cb) cb.checked = sel;
  }
  mselLabelUpdate(cfg, store); onChange();
}
function buildMSel(cfg, items, store, onChange) {
  const menu = $("#" + cfg.id + "Menu"), btn = $("#" + cfg.id + "Btn"); if (!menu || !btn) return;
  const s = store[cfg.id] || (store[cfg.id] = new Set());
  const vals = _mselVals(items, cfg);
  for (const v of [...s]) if (!vals.includes(v)) s.delete(v);   // drop values gone from the data
  menu.innerHTML = "";
  if (cfg.search) { const q = document.createElement("input"); q.type = "search"; q.className = "msel-search"; q.placeholder = "Search";
    q.oninput = () => { const t = q.value.toLowerCase(); for (const lab of menu.querySelectorAll(".msel-opt")) lab.hidden = t && !lab.textContent.toLowerCase().includes(t); }; menu.appendChild(q);
    const tools = document.createElement("div"); tools.className = "msel-tools";
    const all = document.createElement("button"); all.type = "button"; all.className = "msel-tool"; all.textContent = "Select all"; all.onclick = () => mselBulk(cfg, store, true, onChange);
    const none = document.createElement("button"); none.type = "button"; none.className = "msel-tool"; none.textContent = "None"; none.onclick = () => mselBulk(cfg, store, false, onChange);
    const cnt = document.createElement("span"); cnt.className = "msel-count";
    tools.appendChild(all); tools.appendChild(none); tools.appendChild(cnt); menu.appendChild(tools); }
  { const clr = document.createElement("button"); clr.type = "button"; clr.className = "msel-clear"; clr.textContent = "Clear";
    clr.onclick = () => { s.clear(); buildMSel(cfg, items, store, onChange); onChange(); }; menu.appendChild(clr); }
  for (const v of vals) {
    const lab = document.createElement("label"); lab.className = "msel-opt"; lab.dataset.val = v;
    const cb = document.createElement("input"); cb.type = "checkbox"; cb.checked = s.has(v);
    cb.onchange = () => { if (cb.checked) s.add(v); else s.delete(v); mselLabelUpdate(cfg, store); onChange(); };
    lab.appendChild(cb); lab.appendChild(document.createTextNode(" " + cfg.lbl(v)));
    menu.appendChild(lab);
  }
  mselLabelUpdate(cfg, store);
}
function mselPasses(item, cfgs, store) {
  for (const cfg of cfgs) { const s = store[cfg.id]; if (s && s.size && !s.has(cfg.get(item))) return false; }
  return true;
}

// Pay Bills multi-select filters - the same rich set as the Bills tab (owner 2026-08-21:
// "filter down just like bills: client, approved, liens, project, division - same multi-select").
const payMSel = {};
let _payMSelSig = null;
const PAY_MSEL = [
  { id: "pfClient", all: "All clients",   get: b => b.client || "",      search: true, lbl: v => v || "(no client)" },
  { id: "pfVendor", all: "All vendors",   get: b => b.vendor || "",      search: true, lbl: v => v || "(none)" },
  { id: "pfDiv",    all: "All divisions", get: b => b.division || "",     lbl: v => v || "(none)" },
  { id: "pfAppr",   all: "Any approval",  get: b => b.approved || "",     lbl: v => v === "approved" ? "Approved" : (v === "not approved" ? "Not approved" : (v || "(blank)")) },
  { id: "pfLien",   all: "Any lien",      get: b => b.lien_status || "",  lbl: v => v ? (LIEN_SHORT[v] || v) : "(no lien clock)" },
];
function buildPayFilters() { for (const cfg of PAY_MSEL) buildMSel(cfg, _payMarkable(), payMSel, renderPayBills); }

const payFunded = b => (b.inv_ar_status === "Paid") || (b.invoice_status === "Invoice paid");
function _payFilterPass(b) {
  const st = payState(b);
  const show = $("#pfShow") ? $("#pfShow").value : "open";
  if (show === "run") { if (!st.selected) return false; }
  else if (num(b.open_balance) <= 0 && !st.selected) return false;   // "Open bills" (default): still owed
  const q = ($("#pfSearch") ? $("#pfSearch").value : "").trim().toLowerCase();
  if (q && !`${b.project_no || ""} ${b.bill_ref || ""} ${b.invoice_no || ""} ${b.vendor || ""} ${b.client || ""}`.toLowerCase().includes(q)) return false;
  if ($("#pfFunded") && $("#pfFunded").checked && !payFunded(b)) return false;
  if (!mselPasses(b, PAY_MSEL, payMSel)) return false;              // Client / Vendor / Division / Approved / Lien
  return true;
}
function payArCell(b) {
  const v = b.inv_ar_status || "";
  if (v) { const cls = v === "Paid" ? "st-ok" : /partial/i.test(v) ? "st-warn" : "st-bad";
    return stText(/partial/i.test(v) ? "Partial" : v, cls, "GC draw (AR invoice) status"); }
  return invText(b);
}
function _paySetSelected(b, sel) {
  const cur = payState(b);
  payDraft.set(b.bill_id, { selected: sel, amount: sel ? cur.amount : null });
  renderPayBills();
}
function _paySetAmount(b, val) {           // live: update draft + save bar only (keep input focus)
  let amt = val === "" ? null : Math.max(0, Math.round(num(val)));
  if (amt != null && amt === Math.round(num(b.open_balance))) amt = null;   // exactly the full balance → "full"
  payDraft.set(b.bill_id, { selected: true, amount: amt });
  renderPaySaveBar(); renderPayList();
}
function renderPayBills() {
  _payRecomputeSaved();   // always reflect the latest server run; payDraft holds unsaved edits on top
  const thead = $("#payBillsTable thead"), tbody = $("#payBillsTable tbody"); if (!thead || !tbody) return;
  // Build the multi-select filter menus once per data change (NOT on every render), so a checkbox
  // toggle keeps its open search box - a toggle re-renders with the same bill set, same signature.
  const paySig = String(_payMarkable().length);
  if (paySig !== _payMSelSig || !($("#pfClientMenu") && $("#pfClientMenu").querySelector(".msel-opt"))) {
    _payMSelSig = paySig; buildPayFilters();
  }
  let rows = _payMarkable().filter(_payFilterPass);
  rows.sort((a, b) => (a.vendor || "").localeCompare(b.vendor || "") || String(a.bill_date || "").localeCompare(String(b.bill_date || "")));
  const cols = ["Pay", "Vendor", "Client", "Project #", "Bill #", "Date", "Open bal", "Pay $", "GC draw", "Invoice #", "Lien"];
  thead.innerHTML = ""; tbody.innerHTML = "";
  const htr = document.createElement("tr");
  cols.forEach((c, i) => { const th = document.createElement("th");
    if (![0, 6, 7].includes(i)) th.className = "left"; th.textContent = c; htr.appendChild(th); });
  thead.appendChild(htr);
  if (!rows.length) { const tr = document.createElement("tr"); const td = document.createElement("td");
    td.colSpan = cols.length; td.className = "left"; td.style.color = "var(--text-dim)"; td.style.padding = "14px 12px";
    td.textContent = BILLS && BILLS.length ? "No bills match - widen or clear the filters." : "No AP data - run load_bill_tracker.py.";
    tr.appendChild(td); tbody.appendChild(tr); renderPaySaveBar(); renderPayList(); return; }
  let capped = false;
  rows.forEach((b, i) => {
    if (i >= PAY_ROW_CAP) { capped = true; return; }
    const st = payState(b), sel = st.selected;
    const tr = document.createElement("tr"); tr.className = "pay-row" + (sel ? " on" : "");
    // Pay checkbox
    const c0 = document.createElement("td"); c0.style.textAlign = "center";
    const cb = document.createElement("input"); cb.type = "checkbox"; cb.checked = sel; cb.title = "Include this bill in the pay run";
    cb.onchange = () => _paySetSelected(b, cb.checked); c0.appendChild(cb); tr.appendChild(c0);
    // Vendor / Client / Project / Bill # / Date
    const cV = document.createElement("td"); cV.className = "left"; const vs = document.createElement("span"); vs.className = "bill-vendor"; vs.textContent = b.vendor || "–"; cV.appendChild(vs); tr.appendChild(cV);
    const cC = document.createElement("td"); cC.className = "left"; cC.textContent = b.client || "–"; if (!b.client) cC.style.color = "var(--text-dim)"; tr.appendChild(cC);
    const cP = document.createElement("td"); cP.className = "left"; cP.textContent = b.project_no || "–"; tr.appendChild(cP);
    tr.appendChild(qboLinkCell(b.bill_ref, qboBillHref(b.qbo_link), "Open this bill in QuickBooks"));
    const cD = document.createElement("td"); cD.className = "left"; cD.textContent = fmtDateShort(b.bill_date); tr.appendChild(cD);
    // Open balance
    const cO = document.createElement("td"); cO.className = "right"; cO.appendChild(moneyCell(b.open_balance)); tr.appendChild(cO);
    // Pay $ (editable; disabled unless selected)
    const cPay = document.createElement("td"); cPay.className = "right";
    const inp = document.createElement("input"); inp.type = "number"; inp.min = "0"; inp.step = "1"; inp.className = "pay-amt";
    inp.value = String(Math.round(payAmountOf(b))); inp.disabled = !sel;
    inp.oninput = () => { _paySetAmount(b, inp.value);
      inp.classList.toggle("partial", inp.value !== "" && num(inp.value) !== Math.round(num(b.open_balance))); };
    if (sel && st.amount != null) inp.classList.add("partial");   // an explicit custom amount (not the full balance)
    cPay.appendChild(inp); tr.appendChild(cPay);
    // GC draw status / Invoice # / Lien
    const cAr = document.createElement("td"); cAr.className = "left"; const ar = payArCell(b); if (ar) cAr.appendChild(ar); else cAr.textContent = "–"; tr.appendChild(cAr);
    tr.appendChild(_payInvNoCell(b));
    const cL = document.createElement("td"); cL.className = "left"; const lt = lienText(b); if (lt) cL.appendChild(lt); else cL.textContent = "–"; tr.appendChild(cL);
    tbody.appendChild(tr);
  });
  if (capped) { const tr = document.createElement("tr"); const td = document.createElement("td");
    td.colSpan = cols.length; td.className = "left"; td.style.color = "var(--text-dim)"; td.style.padding = "8px 12px";
    td.textContent = `Showing the first ${PAY_ROW_CAP} - narrow the filters to see the rest.`; tr.appendChild(td); tbody.appendChild(tr); }
  renderPaySaveBar(); renderPayList();
}
function _payInvNoCell(b) {
  const td = document.createElement("td"); td.className = "left";
  if (b.invoice_no && b.inv_qbo_id) { const a = document.createElement("a"); a.href = qboInvoiceUrl(b.inv_qbo_id);
    a.target = "_blank"; a.rel = "noopener"; a.className = "qbo-link"; a.textContent = b.invoice_no; a.title = "Open this invoice in QuickBooks"; td.appendChild(a); }
  else { const s = document.createElement("span"); s.textContent = b.invoice_no || "–"; if (!b.invoice_no) s.style.color = "var(--text-dim)"; td.appendChild(s); }
  return td;
}
function renderPaySaveBar() {
  const sel = paySelectedBills();
  const total = sel.reduce((t, b) => t + payAmountOf(b), 0);
  const n = payDirtyCount();
  const bar = $("#paySaveBar"); if (bar) bar.classList.toggle("dirty", n > 0);
  const txt = $("#paySaveText");
  if (txt) txt.textContent = n
    ? `${n} unsaved change${n > 1 ? "s" : ""} · run: ${sel.length} bill${sel.length !== 1 ? "s" : ""}, ${money(total)}`
    : (sel.length ? `Pay run: ${sel.length} bill${sel.length !== 1 ? "s" : ""} · ${money(total)}` : "No bills in the pay run");
  { const s = $("#btnSavePayRun"); if (s) s.disabled = n === 0; }
  { const d = $("#btnDiscardPayRun"); if (d) d.hidden = n === 0; }
  { const qs = $("#payQuickStat"); if (qs) qs.textContent = sel.length ? `${sel.length} selected · ${money(total)} to pay` : ""; }
  { const note = $("#payNote"); if (note) note.textContent = `(${_payMarkable().filter(b => num(b.open_balance) > 0).length} open bills)`; }
}
function renderPayList() {
  const w = $("#payListWidget"); if (!w) return;
  const sel = paySelectedBills();
  if (!sel.length) { w.hidden = true; return; }
  w.hidden = false;
  const byV = new Map();
  for (const b of sel) { const v = b.vendor || "–"; if (!byV.has(v)) byV.set(v, []); byV.get(v).push(b); }
  const vendors = [...byV.keys()].sort((a, b) => a.localeCompare(b));
  const grand = sel.reduce((t, b) => t + payAmountOf(b), 0);
  const note = $("#payListNote");
  if (note) note.textContent = `(${vendors.length} vendor${vendors.length !== 1 ? "s" : ""} · ${sel.length} bill${sel.length !== 1 ? "s" : ""} · ${money(grand)})`;
  const thead = $("#payListTable thead"), tbody = $("#payListTable tbody");
  thead.innerHTML = ""; tbody.innerHTML = "";
  const cols = ["Vendor", "Bill #", "Bill date", "Project #", "Client", "Invoice #", "GC draw", "Pay $"];
  const htr = document.createElement("tr");
  cols.forEach((c, i) => { const th = document.createElement("th"); if (i !== cols.length - 1) th.className = "left"; th.textContent = c; htr.appendChild(th); });
  thead.appendChild(htr);
  for (const v of vendors) {
    const list = byV.get(v); const sub = list.reduce((t, b) => t + payAmountOf(b), 0);
    const gtr = document.createElement("tr"); gtr.className = "bill-group";
    const gtd = document.createElement("td"); gtd.colSpan = cols.length;
    const cell = document.createElement("div"); cell.className = "bg-cell";
    const key = document.createElement("span"); key.className = "bg-key"; key.textContent = v;
    cell.appendChild(key);
    bandMetrics(cell, [[list.length, "bills"], [money(sub), "total"]]);
    gtd.appendChild(cell); gtr.appendChild(gtd); tbody.appendChild(gtr);
    for (const b of list) {
      const tr = document.createElement("tr"); tr.className = "pay-row";
      const cV = document.createElement("td"); cV.className = "left"; cV.textContent = b.vendor || "–"; tr.appendChild(cV);
      tr.appendChild(qboLinkCell(b.bill_ref, qboBillHref(b.qbo_link), "Open this bill in QuickBooks"));
      const cDt = document.createElement("td"); cDt.className = "left"; cDt.textContent = fmtDateShort(b.bill_date); tr.appendChild(cDt);
      const cP = document.createElement("td"); cP.className = "left"; cP.textContent = b.project_no || "–"; tr.appendChild(cP);
      const cC = document.createElement("td"); cC.className = "left"; cC.textContent = b.client || "–"; tr.appendChild(cC);
      tr.appendChild(_payInvNoCell(b));
      const cAr = document.createElement("td"); cAr.className = "left"; const ar = payArCell(b); if (ar) cAr.appendChild(ar); else cAr.textContent = "–"; tr.appendChild(cAr);
      const cPay = document.createElement("td"); cPay.className = "right"; cPay.appendChild(moneyCell(payAmountOf(b))); tr.appendChild(cPay);
      tbody.appendChild(tr);
    }
  }
  const gt = document.createElement("tr"); gt.className = "wip-total";
  cols.forEach((c, i) => { const td = document.createElement("td");
    if (i === 0) { td.className = "left"; td.textContent = "GRAND TOTAL"; }
    else if (i === cols.length - 1) { td.className = "right"; td.appendChild(moneyCell(grand)); }
    else td.className = "left";
    gt.appendChild(td); });
  tbody.appendChild(gt);
}
async function savePayRun() {
  const items = [...payDraft.entries()].map(([bill_id, d]) => ({
    bill_id, selected: !!d.selected, amount: d.selected ? (d.amount == null ? null : num(d.amount)) : null }));
  if (!items.length) return;
  const btn = $("#btnSavePayRun"); if (btn) { btn.disabled = true; btn.textContent = "Saving…"; }
  try {
    const res = await fetch("/api/pay-run", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ items }) });
    const j = await res.json(); if (!j.ok) throw new Error(j.error || "write failed");
    toast(`Pay run saved · ${j.count} bill${j.count !== 1 ? "s" : ""}`);
    payDraft = new Map(); await load(true); renderPayBills();
  } catch (e) { toast("Save failed: " + e.message); }
  finally { if (btn) { btn.textContent = "Save pay run"; } renderPaySaveBar(); }
}
function discardPayRun() { payDraft = new Map(); renderPayBills(); toast("Discarded unsaved changes"); }
async function clearPayRun() {
  if (!paySelectedBills().length && !payDraft.size) { toast("The pay run is already empty"); return; }
  if (!confirm("Empty the whole pay run? This unmarks every bill in it.")) return;
  try {
    const res = await fetch("/api/pay-run/clear", { method: "POST" });
    const j = await res.json(); if (!j.ok) throw new Error(j.error || "clear failed");
    toast(`Cleared ${j.cleared} bill${j.cleared !== 1 ? "s" : ""} from the pay run`);
    payDraft = new Map(); await load(true); renderPayBills();
  } catch (e) { toast("Clear failed: " + e.message); }
}
function paySelectAllShown() {
  const rows = _payMarkable().filter(_payFilterPass).filter(b => num(b.open_balance) > 0);
  for (const b of rows) { const cur = payState(b); if (!cur.selected) payDraft.set(b.bill_id, { selected: true, amount: cur.amount }); }
  renderPayBills();
}
function exportPayList() {
  const sel = paySelectedBills().slice().sort((a, b) => (a.vendor || "").localeCompare(b.vendor || "") || String(a.bill_date || "").localeCompare(String(b.bill_date || "")));
  if (!sel.length) { toast("Nothing to export - mark some bills first"); return; }
  const head = ["Vendor", "Bill #", "Bill date", "Division", "Project #", "Client", "Invoice #", "GC draw status", "Open balance", "Pay amount"];
  const lines = [head];
  for (const b of sel) lines.push([b.vendor || "", b.bill_ref || "", b.bill_date || "", b.division || "", b.project_no || "",
    b.client || "", b.invoice_no || "", b.inv_ar_status || b.invoice_status || "", Math.round(num(b.open_balance)), Math.round(payAmountOf(b))]);
  const csv = lines.map(r => r.map(c => { const s = String(c == null ? "" : c); return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; }).join(",")).join("\r\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob); const a = document.createElement("a");
  a.href = url; a.download = "pay-run.csv"; document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
// Press the save-bar text → a review of the lien marks: what you STAGED (old → new, so nothing
// saves blind) and what's already ON FILE, plus a jump to the Synology lien folder.
function openLienReview() {
  const staged = [...pendingBillMarks.entries()].map(([id, p]) => ({ b: (BILLS || []).find(x => x.bill_id === id), p }));
  const saved = (BILLS || []).filter(b => b.lien_marked && b.lien_status && !pendingBillMarks.has(b.bill_id));
  $("#lienReviewSub").textContent = staged.length ? `${staged.length} unsaved · review, then Save`
    : (saved.length ? `${saved.length} on file` : "no lien marks yet");
  const body = $("#lienReviewBody"); body.innerHTML = "";
  const lienTxt = v => v ? (LIEN_SHORT[v] || v) : "–";
  const projCell = b => (b && b.project_no ? `${b.project_no}${b.client ? " · " + b.client : ""}` : "–");
  const section = (title, note) => { const g = document.createElement("div"); g.className = "dgroup";
    g.appendChild(el2("h4", null, title));
    if (note) { const p = el2("p", "hint", note); p.style.margin = "2px 0 6px"; g.appendChild(p); }
    body.appendChild(g); return g; };
  const grid = (cols) => { const t = document.createElement("table"); t.className = "sub-grid";
    t.innerHTML = "<thead><tr>" + cols.map(c => `<th class='left'>${c}</th>`).join("") + "</tr></thead>";
    const tb = document.createElement("tbody"); t.appendChild(tb); return { t, tb }; };
  if (staged.length) {
    const g = section("Unsaved changes", "Review each change, then Save. Discard drops them all.");
    const { t, tb } = grid(["Vendor", "Project · client", "Bill #", "Change"]);
    for (const { b, p } of staged) {
      const tr = document.createElement("tr");
      tr.appendChild(leftText(b ? (b.vendor || "–") : "(bill not on screen)"));
      tr.appendChild(leftText(projCell(b)));
      tr.appendChild(leftText(b ? (b.bill_ref || "–") : "–"));
      const ch = document.createElement("td"); ch.className = "left";
      ch.innerHTML = `<span class="dim">${lienTxt(p.prevLien)}</span> → <b>${p.lien ? lienTxt(p.lien) : "cleared"}</b>`;
      tr.appendChild(ch); tb.appendChild(tr);
    }
    g.appendChild(t);
    const acts = document.createElement("div"); acts.className = "pnl-actions";
    const sv = document.createElement("button"); sv.className = "btn"; sv.textContent = `Save ${staged.length} mark${staged.length > 1 ? "s" : ""}`;
    sv.onclick = () => { closePanels(); saveBillMarks(); }; acts.appendChild(sv);
    const dc = document.createElement("button"); dc.className = "btn subtle"; dc.textContent = "Discard"; dc.onclick = () => { discardBillMarks(); openLienReview(); }; acts.appendChild(dc);
    g.appendChild(acts);
  }
  if (saved.length) {
    const g = section(`On file (${saved.length})`, "Lien marks currently in effect on your bills.");
    const { t, tb } = grid(["Vendor", "Project · client", "Bill #", "Mark"]);
    saved.sort((a, b) => (a.vendor || "").localeCompare(b.vendor || ""));
    for (const b of saved) {
      const tr = document.createElement("tr");
      tr.appendChild(leftText(b.vendor || "–"));
      tr.appendChild(leftText(projCell(b)));
      tr.appendChild(leftText(b.bill_ref || "–"));
      const m = document.createElement("td"); m.className = "left"; m.appendChild(stText(lienTxt(b.lien_status), "st-lien-" + (LIEN_CLASS[b.lien_status] || "info"))); tr.appendChild(m);
      tb.appendChild(tr);
    }
    g.appendChild(t);
  }
  if (!staged.length && !saved.length) body.appendChild(el2("p", "hint", "No lien marks yet. Open a bill and mark Notice Sent / Lien Filed / Released."));
  { const g = section("Lien documents", "");
    const a = document.createElement("button"); a.className = "btn"; a.textContent = "Open lien folder ↗";
    a.title = "Open the Synology Vendor Liens folder (where the notice / lien PDFs are filed)";
    a.onclick = () => fetch("/api/lien/folder", { method: "POST" }).then(r => r.json()).then(j => toast(j.error ? "Couldn't open: " + j.error : "Opened the lien folder"));
    g.appendChild(a);
    const p = el2("p", "hint", "Vendor Liens / 2026 on the Accounting share."); p.style.marginTop = "6px"; g.appendChild(p);
  }
  openPanel("#lienReview");
}
// Click a bill row → the invoice slides in on the right: bill (money out) + its AR
// invoice / draw (money in), with QuickBooks deep links to both.
function openBillDetail(b) {
  $("#billDetailTitle").textContent = b.vendor || "Bill";
  const projLbl = b.project_no ? (nameOf(b.project_no) ? `${b.project_no} · ${nameOf(b.project_no)}` : b.project_no) : "No project";
  $("#billDetailSub").textContent = `Bill ${b.bill_ref || "–"} · ${projLbl}`;
  const body = $("#billDetailBody"); body.innerHTML = "";
  const grp = (label) => { const g = document.createElement("div"); g.className = "dgroup";
    const h = document.createElement("h4"); h.textContent = label; g.appendChild(h); body.appendChild(g); return g; };
  const row = (g, k, v, cls) => { const r = document.createElement("div"); r.className = "drow";
    const dk = document.createElement("span"); dk.className = "dk"; dk.textContent = k;
    const dv = document.createElement("span"); dv.className = "dv" + (cls ? " " + cls : "");
    if (v instanceof Node) dv.appendChild(v); else dv.textContent = (v == null || v === "") ? "–" : v;
    r.appendChild(dk); r.appendChild(dv); g.appendChild(r); };
  const linkBtn = (label, url) => { const a = document.createElement("a"); a.className = "btn"; a.href = url;
    a.target = "_blank"; a.rel = "noopener"; a.textContent = label; return a; };

  const gb = grp("Bill  ·  money out");
  row(gb, "Vendor", b.vendor);
  row(gb, "Bill #", b.bill_ref);
  row(gb, "Bill date", fmtDate(b.bill_date));
  row(gb, "This line", money(b.line_amount));                                   // the tracker row is ONE line of the bill
  if (b.bill_total != null && Math.abs(num(b.bill_total) - num(b.line_amount)) > 0.5) row(gb, "Bill total (what the qb link opens)", money(b.bill_total));
  row(gb, "Open balance", money(b.open_balance), bOpen(b) > 0 ? "neg" : "");
  row(gb, "Paid the vendor?", b.pay_status);
  row(gb, "Approved?", b.approved === "approved" ? "Yes" : (b.approved ? "No" : ""));
  row(gb, "Lien clock", b.lien_status);
  // Lien mark: the owner sets Notice Sent / Lien Filed / Released here. Saves to the ledger
  // instantly and mirrors into the workbook's Lien cell on the next sync-ap.
  { const r = document.createElement("div"); r.className = "drow lien-mark-row";
    const dk = document.createElement("span"); dk.className = "dk"; dk.textContent = "Mark lien";
    const dv = document.createElement("span"); dv.className = "dv lien-mark-ctl";
    if (!b.bill_id) { const s = document.createElement("span"); s.className = "dim"; s.textContent = "no QBO bill link"; dv.appendChild(s); }
    else {
      for (const [label, val] of [["Notice Sent", "Notice Sent"], ["Lien Filed", "Lien Filed"], ["Released", "✓ Released"]]) {
        const active = b.lien_marked && b.lien_status === val;
        const btn = document.createElement("button");
        btn.className = "btn small lien-mark-btn" + (active ? " active" : "");
        btn.textContent = label; btn.title = active ? "Click to clear this mark" : ("Mark " + label);
        btn.onclick = () => setBillLien(b, active ? "" : val);
        dv.appendChild(btn);
      }
    }
    r.appendChild(dk); r.appendChild(dv); gb.appendChild(r); }
  if (b.bill_id) { const hint = document.createElement("p"); hint.className = "hint lien-mark-hint";
    hint.textContent = pendingBillMarks.has(b.bill_id)
      ? "Unsaved - hit Save in the bar at the bottom to commit. It mirrors to the workbook on the next AP sync."
      : "Pick a tag to stage it, then Save (bar appears at the bottom). Mirrors to the workbook on the next AP sync.";
    gb.appendChild(hint); }
  { const acts = document.createElement("div"); acts.className = "pnl-actions";
    const bl = qboBillHref(b.qbo_link); if (bl) acts.appendChild(linkBtn("Open bill in QuickBooks ↗", bl));
    if (b.vendor) { const lf = document.createElement("button"); lf.className = "btn subtle"; lf.textContent = "Open lien folder ↗";
      lf.title = "Open this vendor's lien folder on the Accounting share (created if missing)";
      lf.onclick = () => fetch("/api/lien/folder?vendor=" + encodeURIComponent(b.vendor), { method: "POST" }).then(r => r.json()).then(j => toast(j.error ? "Couldn't open: " + j.error : `Opened ${b.vendor}'s lien folder`));
      acts.appendChild(lf); }
    if (acts.childNodes.length) gb.appendChild(acts); }

  const gi = grp("Invoice / draw  ·  money in");
  row(gi, "Invoice #", b.invoice_no);
  const drawMemo = (b.matched_invoice || "").split("\n")[0].trim();
  row(gi, "Draw", drawMemo || b.matched_invoice);
  row(gi, "Invoice status", b.invoice_status);
  if (b.inv_ar_status) row(gi, "GC paid the invoice?", b.inv_ar_status,
    /paid/i.test(b.inv_ar_status) && !/unpaid|partial/i.test(b.inv_ar_status) ? "pos" : "neg");
  if (b.inv_amount != null) row(gi, "Invoice amount", money(b.inv_amount));
  if (b.inv_balance != null) row(gi, "GC still owes", money(b.inv_balance), (b.inv_balance || 0) > 0.005 ? "neg" : "pos");
  if (b.gc_paid_date) row(gi, "GC funded", fmtDate(b.gc_paid_date));
  { const acts = document.createElement("div"); acts.className = "pnl-actions";
    if (b.inv_qbo_id) acts.appendChild(linkBtn("Open invoice in QuickBooks ↗", qboInvoiceUrl(b.inv_qbo_id)));
    if (acts.childNodes.length) gi.appendChild(acts);
    else { const p = document.createElement("p"); p.className = "hint"; p.style.margin = "2px 0 0";
      p.textContent = b.invoice_no ? "Invoice not matched in the ledger yet - no direct link." : "No invoice on this bill yet.";
      gi.appendChild(p); } }

  openPanel("#billDetail");
}

// ── Sub LOC (subcontractor float we front before the GC pays) ───────────────
// From load_sub_loc.py (shared/sub_loc engine). By project FIRST (click a row → its open
// subs grouped by the draw they sit under); the feed is bucketed this week / this month /
// prior (prior collapsed); By division collapses. QBO links: each sub bill, and a project's
// customerdetail page (all its transactions).
let sublocCollapsed = new Set(["feed-prior"]);   // the feed's prior-months bucket, collapsed by default
let sublocProjExpanded = false;   // By project shows the top few most-in-the-hole; expand for the rest
const SUBLOC_PROJ_TOP = 5;
const _slBuildG = (sel, cols) => { const th = $(sel + " thead"), tb = $(sel + " tbody"); th.innerHTML = ""; tb.innerHTML = "";
  const htr = document.createElement("tr"); for (const [c, al] of cols) { const h = document.createElement("th"); if (al === "left") h.className = "left"; h.textContent = c; htr.appendChild(h); } th.appendChild(htr); return tb; };
const _slMcell = v => { const td = document.createElement("td"); td.appendChild(moneyCell(v)); return td; };
const _slEmpty = (tb, n, msg) => { const tr = document.createElement("tr"); const td = document.createElement("td"); td.colSpan = n; td.className = "left"; td.style.color = "var(--text-dim)"; td.style.padding = "12px"; td.textContent = msg; tr.appendChild(td); tb.appendChild(tr); };
// Company-scoped QBO customer page = all of a project's transactions (customerdetail?nameId=).
function qboCustomerUrl(custId) {
  if (!custId) return null; const realm = meta && meta.qbo_realm; const page = "customerdetail?nameId=" + custId;
  return realm ? `https://qbo.intuit.com/app/login?pagereq=${encodeURIComponent(page)}&deeplinkcompanyid=${encodeURIComponent(realm)}`
               : `https://qbo.intuit.com/app/${page}`;
}
function applySublocSections() {
  for (const head of $$(".sec-head")) {
    const sec = head.closest(".widget"); if (!sec) continue;
    const collapsed = sublocCollapsed.has(head.dataset.sec);
    const caret = head.querySelector(".sec-caret"); if (caret) caret.textContent = collapsed ? "▸" : "▾";
    const body = sec.querySelector(".sec-body"); if (body) body.hidden = collapsed;
  }
}
// ══ OPEN INVOICES (AR aging) ════════════════════════════════════════════════
// The GC's side of the ledger: what they still owe you, aged by DUE DATE into the
// same Current/1-30/31-60/61-90/90+ buckets as the AR Aging workbook, each carrying
// the matching Notion Lien Tracker status. Read-only; Invoice # deep-links to QBO.
let invExpanded = new Set();      // customer groups the owner has EXPANDED (default: none = all collapsed, owner 2026-08-31)
let invGroupKeys = [];            // customer groups on screen (drives Collapse/Expand-all)
let invBucketFilter = null;       // aging bucket clicked in the stats row (null = all)
let invSubGroup = true;           // sub-group a client's invoices by project (default) vs one flat list
const invMSel = {};               // Client / Project # multi-select filters (owner 2026-08-21)
let _invMSelSig = null;
const INV_MSEL = [
  { id: "ifClient", all: "All clients",  get: i => i.customer || "",   search: true, lbl: v => v || "(no client)" },
  { id: "ifProj",   all: "All projects", get: i => i.project_no || "", search: true, lbl: v => v || "(no project)" },
];
const AGING_HEX = ["#2E7D32", "#7CB342", "#D68910", "#C0552B", "#922B21"];  // green→red (matches aging_sheet.py)

// Notion Lien Tracker status → [label, dot color]. Rendered as a Notion-style status pill
// (grey pill + a colored dot) so it reads as "this came from the Notion Lien Tracker".
const OI_LIEN = {
  "Lien":            ["Lien filed",   "#C0392B"],   // red
  "Mailed":          ["Mailed",       "#2E77BC"],   // blue
  "Ready to Mail":   ["Ready to mail", "#D68910"],  // orange
  "In progress":     ["In progress",  "#2E77BC"],   // blue
  "Ready to Review": ["Review",       "#D68910"],   // orange
  "Not started":     ["Not started",  "#9AA1AC"],   // grey
  "Did Not Send":    ["Skipped",      "#9AA1AC"],   // grey
  "Paid":            ["Paid",         "#3E9B57"],   // green
  "Closed":          ["Closed",       "#9AA1AC"],   // grey
};
function oiLienNode(inv) {
  const v = inv.lien_status; if (!v) return null;
  const m = OI_LIEN[v] || [v, "#9AA1AC"];
  const pill = document.createElement("span"); pill.className = "notion-pill";
  const dot = document.createElement("span"); dot.className = "np-dot"; dot.style.background = m[1];
  const lbl = document.createElement("span"); lbl.textContent = m[0];
  pill.appendChild(dot); pill.appendChild(lbl);
  pill.title = "Notion Lien Tracker: " + v + (inv.lien_notice ? " · " + inv.lien_notice : "");
  return pill;
}
// Computed Texas lien-notice CLOCK (when a lien is due) - from shared/lien_clock in the backend
// payload, the SAME clock the AR Aging Excel uses so the two never disagree.
function oiLienClock(inv) {
  const v = inv.lien_due_label; if (!v) return null;
  const st = inv.lien_due_state || "";
  const cls = st === "PAST" ? "lc-past" : (st === "URGENT" ? "lc-urgent"
            : (st === "WATCH" || st === "RETAINAGE") ? "lc-watch" : st === "SENT" ? "lc-sent" : "lc-ok");
  const s = document.createElement("span"); s.className = "lien-clock " + cls;
  s.textContent = v; s.title = "Texas lien-notice deadline (computed)";
  return s;
}
const oiBal = i => num(i.balance);

// The Division + Lien filter selects, built once from the data (preserving the pick).
function buildInvFilters() {
  const invs = OI.invoices || [];
  const specs = [
    { sel: "#ifDivision", get: i => i.division || "", all: "All divisions" },
    { sel: "#ifLien",     get: i => i.lien_status || "", all: "Any lien", none: "No lien on file" },
  ];
  for (const s of specs) {
    const el = $(s.sel); if (!el) continue;
    const prev = el.value;
    const vals = [...new Set(invs.map(s.get).filter(v => v !== ""))].sort((a, b) => a.localeCompare(b));
    el.innerHTML = "";
    const a0 = document.createElement("option"); a0.value = ""; a0.textContent = s.all; el.appendChild(a0);
    if (s.none) { const o = document.createElement("option"); o.value = "__none__"; o.textContent = s.none; el.appendChild(o); }
    for (const v of vals) { const o = document.createElement("option"); o.value = v; o.textContent = v; el.appendChild(o); }
    el.value = prev; if (el.value !== prev) el.value = "";
  }
}

// Lien-notice CLOCK buckets (the computed deadline, not the Notion status). "upcoming" = urgent
// OR watch, which covers CP draws, CP retainage (RET-banded), and RP - all divisions the clock runs.
const LIENCLK = { past: s => s === "PAST", upcoming: s => s === "URGENT" || s === "WATCH", sent: s => s === "SENT" };
// Invoice MONTH filter (owner 2026-09-02: "show all the boxes selected so i can deselect ... we need a
// select / deselect all ... days is useless, remove"). null = every month (all boxes ticked); a Set = the
// ticked months. Unticking the first month turns the full list into a Set minus that month; ticking the
// last missing one goes back to null.
let invMonthSel = null;
function _invMonthsAsc() {
  return [...new Set(((invData().invoices) || []).map(i => String(i.txn_date || "").slice(0, 7)).filter(s => /^\d{4}-\d{2}$/.test(s)))].sort();
}
function _invMonthSet(asc) { return invMonthSel === null ? new Set(asc) : invMonthSel; }
function toggleInvMonth(ym, checked) {
  const asc = _invMonthsAsc(), s = new Set(_invMonthSet(asc));
  if (checked) s.add(ym); else s.delete(ym);
  invMonthSel = (s.size === asc.length || s.size === 0) ? null : s;   // unticking the last one goes back to ALL
  buildInvDateFilter(); renderOpenInvoices();
}
function buildInvDateFilter() {
  const menu = $("#ifMonthMenu"), btn = $("#ifMonthBtn");
  if (!menu || !btn) return;
  const asc = _invMonthsAsc();
  if (invMonthSel) { for (const m of [...invMonthSel]) if (!asc.includes(m)) invMonthSel.delete(m); }
  const sel = _invMonthSet(asc);
  menu.innerHTML = "";
  { const tools = document.createElement("div"); tools.className = "msel-tools";
    const all = document.createElement("button"); all.type = "button"; all.className = "msel-tool"; all.textContent = "Select all";
    all.onclick = () => { invMonthSel = null; buildInvDateFilter(); renderOpenInvoices(); };
    const none = document.createElement("button"); none.type = "button"; none.className = "msel-tool"; none.textContent = "Deselect all";
    none.onclick = () => { invMonthSel = new Set(); buildInvDateFilter(); renderOpenInvoices(); };
    const cnt = document.createElement("span"); cnt.className = "msel-count"; cnt.textContent = `${sel.size} of ${asc.length}`;
    tools.appendChild(all); tools.appendChild(none); tools.appendChild(cnt); menu.appendChild(tools); }
  for (const ym of [...asc].reverse()) {
    const lab = document.createElement("label"); lab.className = "msel-opt";
    const cb = document.createElement("input"); cb.type = "checkbox"; cb.checked = sel.has(ym);
    cb.onchange = () => toggleInvMonth(ym, cb.checked);
    lab.appendChild(cb); lab.appendChild(document.createTextNode(" " + billMonthLabel(ym)));
    menu.appendChild(lab);
  }
  if (invMonthSel === null) btn.textContent = "All months";
  else if (!sel.size) btn.textContent = "No months";
  else { const newest = [...sel].sort().reverse()[0]; btn.textContent = sel.size === 1 ? billMonthLabel(newest) : `${billMonthLabel(newest)} +${sel.size - 1}`; }
  btn.classList.toggle("on", invMonthSel !== null);
}
function invDatePasses(i) {
  return invMonthSel === null || invMonthSel.has(String(i.txn_date || "").slice(0, 7));
}
// Quick find (⌘F / Ctrl+F on this tab): every word must match somewhere in invoice # · memo · amount ·
// project # · client · note · status; a word starting with "-" must NOT match (filter it out).
let invQuick = "";
function _invHay(i) {
  const amt = num(i.amount), bal = oiBal(i);
  return [i.doc_number, i.memo, i.project_no, i.customer, i.note, i.status, i.division,
          amt != null ? String(Math.round(amt)) : "", amt != null ? money(amt) : "",
          bal != null ? String(Math.round(bal)) : "", bal != null ? money(bal) : ""].join(" ").toLowerCase();
}
function invQuickPasses(i) {
  const terms = invQuick.toLowerCase().split(/\s+/).filter(Boolean);
  if (!terms.length) return true;
  const hay = _invHay(i), hayNum = hay.replace(/[$,]/g, "");
  for (const raw of terms) {
    const neg = raw.length > 1 && raw[0] === "-", term = neg ? raw.slice(1) : raw, tn = term.replace(/[$,]/g, "");
    const hit = hay.includes(term) || (tn && hayNum.includes(tn));
    if (neg ? hit : !hit) return false;
  }
  return true;
}
function invPasses(i, f) {
  if (!invDatePasses(i)) return false;                   // Month (invoice date)
  if (!invQuickPasses(i)) return false;                  // quick find (⌘F)
  if (!mselPasses(i, INV_MSEL, invMSel)) return false;   // Client / Project # multi-selects
  if (f.div && (i.division || "") !== f.div) return false;
  if (f.lienclk && LIENCLK[f.lienclk] && !LIENCLK[f.lienclk](i.lien_due_state || "")) return false;
  if (f.lien === "__none__" ? !!i.lien_status : (f.lien && (i.lien_status || "") !== f.lien)) return false;
  if (f.litig === "ex" && i.litigation) return false;
  if (f.litig === "only" && !i.litigation) return false;
  if (invBucketFilter != null && i.bucket_index !== invBucketFilter) return false;
  return true;
}

const INV_SORTS = {
  due:    (a, b) => String(a.due_date || "9999").localeCompare(String(b.due_date || "9999")) || String(a.doc_number || "").localeCompare(String(b.doc_number || "")),
  owed:   (a, b) => oiBal(b) - oiBal(a),
  client: (a, b) => (a.customer || "~").localeCompare(b.customer || "~") || String(a.due_date || "9999").localeCompare(String(b.due_date || "9999")),
};

let invView = "amounts";   // "amounts" | "aging" - the Open-invoices view toggle (owner 2026-08-27)
let invScope = "open";     // "open" | "all" - open-only vs every invoice incl. paid (owner 2026-08-31)
let OI_ALL = null;         // on-demand cache of ALL invoices; fetched the first time scope flips to "all"
function invData() { return (invScope === "all" && OI_ALL) ? OI_ALL : OI; }
// avg days-to-pay for a client, from the active dataset's pay_speed (falls back to the portfolio avg).
function invClientAvgDays(c) { const ps = (invData().pay_speed) || {}; const s = (ps.by_client || {})[(c || "").toLowerCase()]; return (s && s.avg_days != null) ? s.avg_days : null; }
// Flip open-only ↔ all. "All" is fetched on demand the first time (kept off the bulk load), then cached.
async function _setInvScope(scope, seg, btn) {
  invScope = scope;
  seg.querySelectorAll(".seg-btn").forEach(x => x.classList.toggle("on", x === btn));
  if (scope === "all" && !OI_ALL) {
    { const n = $("#invNote"); if (n) n.textContent = "(loading all invoices…)"; }
    try { OI_ALL = await (await fetch("/api/invoices/all")).json(); }
    catch (e) {
      OI_ALL = null; toast("Could not load all invoices"); invScope = "open";
      seg.querySelectorAll(".seg-btn").forEach(x => x.classList.toggle("on", x.dataset.scope === "open"));
    }
  }
  renderOpenInvoices();
}

// The AMOUNTS view: invoices GROUPED BY CLIENT like QBO's AR (owner 2026-08-31). Each client is a
// header showing how many invoices, what's open, and how fast they pay (avg days-to-pay → a guess at
// when it lands); its rows show project / invoice # / date / open / total / the Notion collections
// note. Honors the Open-only↔All scope toggle. Click a row for details, or the invoice # for QBO.
function renderInvAmounts(all, f) {
  const host = $("#invTable"), thead = host.querySelector("thead"), tbody = host.querySelector("tbody");
  { const st = $("#invStats"); if (st) st.innerHTML = ""; }   // aging tiles belong to the Aging view
  const rows = all.filter(i => invPasses(i, f));
  const sortFn = INV_SORTS[($("#ifSort") || {}).value || "due"] || INV_SORTS.due;
  // group by client, tally open + billed, then order clients by most-open-first
  const groups = new Map();
  for (const i of rows) {
    const c = i.customer || "–";
    let g = groups.get(c); if (!g) { g = { client: c, open: 0, billed: 0, rows: [] }; groups.set(c, g); }
    g.open += oiBal(i); g.billed += num(i.amount); g.rows.push(i);
  }
  const clients = [...groups.values()].sort((a, b) => b.open - a.open || (a.client || "~").localeCompare(b.client || "~"));
  for (const g of clients) g.rows.sort(sortFn);
  const totOpen = rows.reduce((t, i) => t + oiBal(i), 0);
  const totBilled = rows.reduce((t, i) => t + num(i.amount), 0);
  $("#invNote").textContent = all.length
    ? `(${rows.length.toLocaleString()} of ${all.length.toLocaleString()} · ${money(totOpen)} open · ${clients.length} client${clients.length === 1 ? "" : "s"})` : "(no AR data)";
  { const anyMsel = INV_MSEL.some(c => (invMSel[c.id] || {}).size);
    const cb = $("#ifClear"); if (cb) cb.hidden = !(anyMsel || invMonthSel !== null || invQuick || f.div || f.lien || f.lienclk || f.litig !== "ex"); }
  const cols = [["Project", "left"], ["Invoice #", "left"], ["Date", "left"], ["Due", "left"], ["Memo", "left"], ["Open balance", "right"], ["Invoice total", "right"], ["Last action", "left"], ["Next follow-up", "left"], ["Collections note", "left"]];
  thead.innerHTML = ""; const htr = document.createElement("tr");
  for (const [c, al] of cols) { const th = document.createElement("th"); th.className = al; th.textContent = c; htr.appendChild(th); } thead.appendChild(htr);
  tbody.innerHTML = "";
  if (!rows.length) {
    const tr = document.createElement("tr"), td = document.createElement("td");
    td.colSpan = cols.length; td.className = "left"; td.style.cssText = "padding:14px;color:var(--text-dim)";
    td.textContent = all.length ? "No invoices match these filters." : "No AR data - run load_invoices.py.";
    tr.appendChild(td); tbody.appendChild(tr); invGroupKeys = []; updateInvCollapseBtn(); return;
  }
  invGroupKeys = clients.map(g => g.client);   // drives the Collapse/Expand-all button
  thead.hidden = !clients.some(g => invExpanded.has(g.client));   // headers only when a client is open (owner 2026-09-01)
  const amtRow = (i) => {
    const paid = oiBal(i) <= 0.005;
    const tr = document.createElement("tr"); tr.style.cursor = "pointer"; if (paid) tr.classList.add("inv-paid");
    tr.title = "Click for the invoice memo + details";
    tr.onclick = (e) => { if (e.target.closest("a")) return; openInvoicePage(i); };
    const pc = document.createElement("td"); pc.className = "left";
    if (i.division) { const dot = document.createElement("span"); dot.className = "divdot " + divClass(i.division); dot.title = i.division; pc.appendChild(dot); }
    pc.appendChild(document.createTextNode(i.project_no || "–")); tr.appendChild(pc);
    tr.appendChild(invNoCell(i));
    tr.appendChild(leftText(fmtDateShort(i.txn_date)));
    // Due + how late (the collections question), then the invoice MEMO in full (owner 2026-09-02:
    // "i need to see the memo ... every single data point for meeting")
    const due = document.createElement("td"); due.className = "left";
    if (i.due_date) { due.textContent = fmtDateShort(i.due_date);
      if (!paid && i.days_past_due != null && i.days_past_due > 0) { const l = document.createElement("span"); l.className = "inv-late"; l.textContent = ` ${i.days_past_due}d late`; due.appendChild(l); } }
    else { due.textContent = "–"; due.classList.add("dim"); }
    tr.appendChild(due);
    const mc = document.createElement("td"); mc.className = "left inv-memo";
    if (i.memo) { mc.textContent = i.memo; mc.title = i.memo; } else { mc.textContent = "–"; mc.classList.add("dim"); }
    tr.appendChild(mc);
    const ob = document.createElement("td"); ob.className = "right";
    if (paid) { ob.textContent = "–"; ob.classList.add("dim"); }
    else { ob.textContent = money(oiBal(i)); if (i.days_past_due != null && i.days_past_due > 0) { ob.style.color = "var(--neg)"; ob.title = i.days_past_due + " days past due"; } }
    tr.appendChild(ob);
    tr.appendChild(rightText(money(i.amount)));
    // The two dates collections runs on (Invoice Tracker "Last Action Date" / "Next Follow-Up"); an
    // overdue follow-up reads red (owner 2026-09-02).
    { const la = document.createElement("td"); la.className = "left"; la.textContent = i.last_action_date ? fmtDateShort(i.last_action_date) : "–"; if (!i.last_action_date) la.classList.add("dim"); tr.appendChild(la);
      const nf = document.createElement("td"); nf.className = "left";
      if (i.next_followup) { nf.textContent = fmtDateShort(i.next_followup); const dd = Date.parse(i.next_followup); if (!paid && !isNaN(dd) && dd < Date.now() - 86400000) { nf.style.color = "var(--neg)"; nf.style.fontWeight = "600"; nf.title = "follow-up date has passed"; } }
      else { nf.textContent = "–"; nf.classList.add("dim"); }
      tr.appendChild(nf); }
    // Collections note = Notion Quick Status; a paid row leads with its paid date
    const nc = document.createElement("td"); nc.className = "left inv-note";
    if (paid && i.paid_date) { const p = document.createElement("span"); p.className = "st ok"; p.textContent = "Paid " + fmtDateShort(i.paid_date); nc.appendChild(p); }
    if (i.note) { if (nc.childNodes.length) nc.appendChild(document.createTextNode(" ")); const n = document.createElement("span"); n.className = "note-txt"; n.textContent = i.note; n.title = i.note; nc.appendChild(n); }
    // The note's Notion page (the Invoice Tracker) - one click to read the whole thread or update it -
    // and how fresh the note is (owner 2026-09-02: "the collections note with what notion page shows if clicked").
    if (i.notion_url) {
      const a = document.createElement("a"); a.className = "notion-link"; a.href = i.notion_url; a.target = "_blank"; a.rel = "noopener";
      a.textContent = "Notion"; a.title = "Open this invoice's page in the Notion Invoice Tracker" + (i.notion_edited ? ` (last edited ${fmtDate(i.notion_edited, true)})` : "");
      a.onclick = e => e.stopPropagation(); nc.appendChild(a);
      if (i.notion_edited) { const ed = document.createElement("span"); ed.className = "note-edited"; ed.textContent = "edited " + fmtDateShort(i.notion_edited); nc.appendChild(ed); }
    }
    if (!nc.childNodes.length) { nc.textContent = "–"; nc.classList.add("dim"); }
    tr.appendChild(nc);   // (was never appended before 2026-09-02 - the column rendered blank)
    return tr;
  };
  for (const g of clients) {
    const expanded = invExpanded.has(g.client);   // collapsed by default; open a client to see its invoices
    // client header (like QBO's customer group): caret, who, how many, open $, and how fast they pay
    const hr = document.createElement("tr"); hr.className = "inv-client"; hr.style.cursor = "pointer";
    hr.title = expanded ? "Click to collapse" : "Click to expand";
    const htd = document.createElement("td"); htd.colSpan = cols.length;
    const caret = document.createElement("span"); caret.className = "bg-caret"; caret.textContent = expanded ? "▾ " : "▸ ";
    const nm = document.createElement("span"); nm.className = "g-cust"; nm.textContent = g.client;
    const ad = invClientAvgDays(g.client);
    const sub = document.createElement("span"); sub.className = "g-sub"; sub.hidden = true;   // (the metrics grid replaced the text run)
    const cellG = document.createElement("div"); cellG.className = "bg-cell"; const leftG = document.createElement("span"); leftG.className = "bg-left"; leftG.appendChild(caret); leftG.appendChild(nm); cellG.appendChild(leftG);
    bandMetrics(cellG, [[g.rows.length, "invoices"], [money(g.open), "open", g.open > 0.005 ? "neg" : ""], [money(g.billed), "billed"], [ad != null ? ad + "d" : "–", "avg days to pay"]]);
    htd.appendChild(cellG); hr.appendChild(htd);
    hr.onclick = () => { if (invExpanded.has(g.client)) invExpanded.delete(g.client); else invExpanded.add(g.client); renderOpenInvoices(); };
    tbody.appendChild(hr);
    if (!expanded) continue;   // collapsed: skip the invoice rows
    // sub-group a client's invoices by PROJECT (owner 2026-08-31: "too mixed up") when the toggle is
    // on and there's more than one project; otherwise a flat list. Same pattern as the Aging view.
    const sortKey = ($("#ifSort") || {}).value || "due";
    const projs = [...new Set(g.rows.map(x => x.project_no || "(no project)"))];
    if (invSubGroup && projs.length > 1) {
      const inP = p => g.rows.filter(x => (x.project_no || "(no project)") === p);
      const pTotal = p => inP(p).reduce((t, x) => t + oiBal(x), 0);
      const pMinDue = p => inP(p).reduce((m, x) => (x.due_date && (!m || x.due_date < m)) ? x.due_date : m, null) || "9999";
      const pCmp = { due: (a, b) => pMinDue(a).localeCompare(pMinDue(b)) || a.localeCompare(b, undefined, { numeric: true }),
        owed: (a, b) => pTotal(b) - pTotal(a), client: (a, b) => a.localeCompare(b, undefined, { numeric: true }) }[sortKey] || null;
      const porder = pCmp ? [...projs].sort(pCmp) : projs;
      for (const p of porder) { tbody.appendChild(invSubBand(p, nameOf(p), pTotal(p), inP(p).length, cols.length)); for (const i of inP(p)) tbody.appendChild(amtRow(i)); }
    } else {
      for (const i of g.rows) tbody.appendChild(amtRow(i));
    }
  }
  const tr = document.createElement("tr"); tr.className = "inv-total-row";
  const td0 = document.createElement("td"); td0.className = "left"; td0.colSpan = 5; td0.textContent = "TOTAL"; tr.appendChild(td0);
  tr.appendChild(rightText(money(totOpen)));
  tr.appendChild(rightText(money(totBilled)));
  tr.appendChild(document.createElement("td")); tr.appendChild(document.createElement("td")); tr.appendChild(document.createElement("td"));
  tbody.appendChild(tr);
  updateInvCollapseBtn();
}

function renderOpenInvoices() {
  const host = $("#invTable"); if (!host) return;
  const D = invData();
  const buckets = D.buckets || ["Current", "1-30", "31-60", "61-90", "90+"];
  const all = D.invoices || [];
  { const h = $("#invHeading"); if (h) h.textContent = invScope === "all" ? "All invoices" : "Open invoices"; }
  if (!$("#ifDivision") || !$("#ifDivision").options.length) buildInvFilters();

  const fv = sel => ($(sel) ? $(sel).value : "");
  const f = { div: fv("#ifDivision"), lien: fv("#ifLien"), lienclk: fv("#ifLienClock"), litig: fv("#ifLitig") || "ex" };  // Client/Project # are msels now
  // Client + Project # multi-selects: build once per data change (signature guard) so a toggle keeps its search.
  const invSig = String(all.length);
  if (invSig !== _invMSelSig || !($("#ifClientMenu") && $("#ifClientMenu").querySelector(".msel-opt"))) {
    _invMSelSig = invSig; for (const cfg of INV_MSEL) buildMSel(cfg, all, invMSel, renderOpenInvoices);
  }
  buildInvDateFilter();
  // Litigation is EXCLUDED by default; flag the box red whenever it's hiding/limiting rows so it's
  // obvious to the eye that a filter is in place (owner 2026-08-19).
  { const el = $("#ifLitig"); if (el) el.classList.toggle("filter-on", (el.value || "ex") !== "all"); }

  // Two views over the same filtered invoices (owner 2026-08-27): AMOUNTS = a clean list of what's
  // owed; AGING = the buckets + lien clock. Both group by client, sub-group by project, and collapse -
  // so the Collapse/Expand-all and Group-by-project buttons show in BOTH (owner 2026-08-31).
  { const fl = $("#ifSubGroup"), cl = $("#ifCollapse");
    if (fl) fl.style.display = ""; if (cl) cl.style.display = ""; }
  if (invView === "amounts") { renderInvAmounts(all, f); return; }

  // Aging tiles double as the bucket filter. Their totals ignore the bucket pick (so the
  // full aging picture always shows) but DO honor the other filters.
  const forTiles = all.filter(i => {
    const save = invBucketFilter; invBucketFilter = null;
    const ok = invPasses(i, f); invBucketFilter = save; return ok;
  });
  const bTot = buckets.map(() => 0); let bGrand = 0;
  for (const i of forTiles) { bTot[i.bucket_index] += oiBal(i); bGrand += oiBal(i); }
  const stats = $("#invStats"); stats.innerHTML = "";
  const mkTile = (label, val, idx, hex, active) => {
    const el = document.createElement("div");
    el.className = "attn ag-tile" + (active ? " active" : "") + (val > 0.005 || idx == null ? "" : " none");
    if (hex) el.style.borderLeftColor = hex;
    el.innerHTML = `<span class="a-count"></span><span class="a-label"></span>`;
    el.querySelector(".a-count").textContent = money(val);
    el.querySelector(".a-label").textContent = label;
    el.onclick = () => { invBucketFilter = (idx == null || invBucketFilter === idx) ? null : idx; renderOpenInvoices(); };
    return el;
  };
  stats.appendChild(mkTile("All open", bGrand, null, "", invBucketFilter == null));
  buckets.forEach((b, k) => stats.appendChild(mkTile(b === "Current" ? "Current" : b + " days", bTot[k], k, AGING_HEX[k], invBucketFilter === k)));

  let rows = all.filter(i => invPasses(i, f));
  rows = [...rows].sort(INV_SORTS[fv("#ifSort") || "due"] || INV_SORTS.due);

  const shown = rows.reduce((t, i) => t + oiBal(i), 0);
  $("#invNote").textContent = all.length
    ? `(${rows.length.toLocaleString()} of ${all.length.toLocaleString()} · ${money(shown)} open)`
    : "(no AR data - run load_invoices.py)";
  { const el = $("#invAsOf"); if (el) el.textContent = D.as_of ? "aged as of " + fmtDate(D.as_of) : ""; }
  { const anyMsel = INV_MSEL.some(c => (invMSel[c.id] || {}).size);
    const cb = $("#ifClear"); if (cb) cb.hidden = !(anyMsel || invMonthSel !== null || invQuick || f.div || f.lien || f.lienclk || f.litig !== "ex" || invBucketFilter != null); }

  const thead = host.querySelector("thead"), tbody = host.querySelector("tbody");
  thead.innerHTML = ""; tbody.innerHTML = "";
  const cols = [["Client", "left"], ["Project", "left"], ["Invoice #", "left"], ["Date", "left"],
                ["Net", "left"], ["Lien", "left"], ...buckets.map(b => [b, "right ag"])];
  const htr = document.createElement("tr");
  cols.forEach(([c, al]) => { const th = document.createElement("th"); th.className = al; th.textContent = c; htr.appendChild(th); });
  thead.appendChild(htr);

  if (!rows.length) {
    const tr = document.createElement("tr"); const td = document.createElement("td");
    td.colSpan = cols.length; td.className = "left"; td.style.color = "var(--text-dim)"; td.style.padding = "14px 12px";
    td.textContent = all.length ? "No open invoices match these filters." : "No AR data - run load_invoices.py.";
    tr.appendChild(td); tbody.appendChild(tr); invGroupKeys = []; updateInvCollapseBtn(); return;
  }

  // group by client (banding + collapse); per-bucket grand total at the bottom
  const groups = new Map();
  for (const i of rows) { const k = i.customer || "(no client)"; if (!groups.has(k)) groups.set(k, []); groups.get(k).push(i); }
  // Order the CLIENT GROUPS by the chosen sort (not just A-Z), so "Oldest due first" really
  // puts the client with the oldest invoice on top, "Most owed" the biggest balance, etc.
  const sortKey = fv("#ifSort") || "due";
  const gMinDue = k => groups.get(k).reduce((m, i) => (i.due_date && (!m || i.due_date < m)) ? i.due_date : m, null) || "9999";
  const gTotal = k => groups.get(k).reduce((t, i) => t + oiBal(i), 0);
  const groupCmp = {
    due:    (a, b) => gMinDue(a).localeCompare(gMinDue(b)) || a.localeCompare(b, undefined, { numeric: true }),
    owed:   (a, b) => gTotal(b) - gTotal(a),
    client: (a, b) => a.localeCompare(b, undefined, { numeric: true }),
  };
  const order = [...groups.keys()].sort(groupCmp[sortKey] || groupCmp.due);
  invGroupKeys = order;
  const grand = buckets.map(() => 0);

  for (const k of order) {
    const g = groups.get(k);
    const collapsed = !invExpanded.has(k);   // collapsed by default; expanded only if the owner opened it
    const gOpen = g.reduce((t, x) => t + oiBal(x), 0);
    const gtr = document.createElement("tr"); gtr.className = "bill-group"; gtr.style.cursor = "pointer";
    gtr.title = collapsed ? "Click to expand" : "Click to collapse";
    const gtd = document.createElement("td"); gtd.colSpan = cols.length;
    const cell = document.createElement("div"); cell.className = "bg-cell";   // flex on the div, not the td
    const left = document.createElement("span"); left.className = "bg-left";
    const caret = document.createElement("span"); caret.className = "bg-caret"; caret.textContent = collapsed ? "▸" : "▾";
    const key = document.createElement("span"); key.className = "bg-key"; key.textContent = k;
    left.appendChild(caret); left.appendChild(key);
    cell.appendChild(left);
    bandMetrics(cell, [[money(gOpen), "open", gOpen > 0.005 ? "neg" : ""], [g.length, "invoices"]]);
    gtd.appendChild(cell); gtr.appendChild(gtd);
    gtr.onclick = () => { if (invExpanded.has(k)) invExpanded.delete(k); else invExpanded.add(k); renderOpenInvoices(); };
    tbody.appendChild(gtr);
    for (const i of g) grand[i.bucket_index] += oiBal(i);   // grand total counts every invoice, even collapsed
    if (collapsed) continue;
    // Sub-group a client's invoices by PROJECT when the toggle is on and there's >1 project
    // (owner 2026-08-21); otherwise the original flat list. Grand total is unaffected either way.
    const projs = [...new Set(g.map(x => x.project_no || "(no project)"))];
    if (invSubGroup && projs.length > 1) {
      const inP = p => g.filter(x => (x.project_no || "(no project)") === p);
      const pMinDue = p => inP(p).reduce((m, x) => (x.due_date && (!m || x.due_date < m)) ? x.due_date : m, null) || "9999";
      const pTotal = p => inP(p).reduce((t, x) => t + oiBal(x), 0);
      const pCmp = { due: (a, b) => pMinDue(a).localeCompare(pMinDue(b)) || a.localeCompare(b, undefined, { numeric: true }),
        owed: (a, b) => pTotal(b) - pTotal(a),
        client: (a, b) => a.localeCompare(b, undefined, { numeric: true }) }[sortKey] || null;
      const porder = pCmp ? [...projs].sort(pCmp) : projs;
      for (const p of porder) {
        const pg = inP(p);
        tbody.appendChild(invSubBand(p, nameOf(p), pTotal(p), pg.length, cols.length));
        for (const i of pg) tbody.appendChild(invRow(i, buckets));
      }
    } else {
      for (const i of g) tbody.appendChild(invRow(i, buckets));
    }
  }
  const ttr = document.createElement("tr"); ttr.className = "ag-total";
  const lead = document.createElement("td"); lead.className = "left"; lead.colSpan = 6; lead.textContent = "Total open"; ttr.appendChild(lead);
  buckets.forEach((b, k) => {
    const td = document.createElement("td"); td.className = "right ag";
    if (grand[k] > 0.005) { td.textContent = money(grand[k]); td.classList.add("ag" + k); }
    ttr.appendChild(td);
  });
  tbody.appendChild(ttr);
  updateInvCollapseBtn();
}

function invRow(i, buckets) {
  const tr = document.createElement("tr");
  const cli = document.createElement("td"); cli.className = "left dim"; cli.textContent = i.customer || "–"; tr.appendChild(cli);
  const proj = document.createElement("td"); proj.className = "left";
  if (i.division) { const dot = document.createElement("span"); dot.className = "divdot " + divClass(i.division); dot.title = i.division; proj.appendChild(dot); }
  const purl = qboCustomerUrl(i.cust_id);   // project # → QBO project page (all its transactions)
  if (purl && i.project_no) {
    const a = document.createElement("a"); a.href = purl; a.target = "_blank"; a.rel = "noopener"; a.className = "qbo-link";
    a.textContent = i.project_no; a.title = "Open this project in QuickBooks (all transactions)"; a.onclick = e => e.stopPropagation();
    proj.appendChild(a);
  } else { proj.appendChild(document.createTextNode(i.project_no || "–")); }
  tr.appendChild(proj);
  tr.appendChild(invNoCell(i));
  const dt = document.createElement("td"); dt.className = "left"; dt.textContent = fmtDateShort(i.txn_date);
  if (i.days_past_due != null && i.days_past_due > 0) dt.title = i.days_past_due + " days past due (due " + fmtDateShort(i.due_date) + ")";
  tr.appendChild(dt);
  const net = document.createElement("td"); net.className = "left dim"; net.textContent = i.net_terms || "–"; tr.appendChild(net);
  // Lien cell: the computed notice-deadline CLOCK (when a lien is due) + the Notion status pill.
  const lien = document.createElement("td"); lien.className = "left status-col lien-cell";
  const clock = oiLienClock(i); if (clock) lien.appendChild(clock);
  const pill = oiLienNode(i); if (pill) lien.appendChild(pill);
  if (!clock && !pill) lien.appendChild(dimDash());
  tr.appendChild(lien);
  buckets.forEach((b, k) => {
    const td = document.createElement("td"); td.className = "right ag";
    if (k === i.bucket_index) {
      td.classList.add("ag" + k);
      td.appendChild(document.createTextNode(money(oiBal(i))));
      if (i.litigation) { td.title = "In litigation"; const f = document.createElement("span"); f.className = "litig"; f.textContent = " ⚖"; td.appendChild(f); }
    }
    tr.appendChild(td);
  });
  tr.style.cursor = "pointer";
  tr.title = "Click for the invoice memo + details (no QuickBooks)";
  tr.onclick = (e) => { if (e.target.closest("a")) return; openInvoicePage(i); };
  return tr;
}

// Invoice number cell: the NUMBER opens the native detail (memo + fields, no QBO); the small
// ↗ opens QuickBooks for when you actually need it (owner 2026-08-25: "i like the qbo links but
// hate using qbo"). `inv` carries the billing_event fields (doc_number, memo, amount, …).
function invNoCell(inv) {
  const td = document.createElement("td"); td.className = "left";
  const docn = inv && (inv.doc_number || inv.invoice_no);
  if (!docn) { td.appendChild(dimDash()); return td; }
  const link = document.createElement("span"); link.className = "inv-detail-link"; link.textContent = docn;
  link.title = "Invoice memo + details (no QuickBooks)";
  link.onclick = (e) => { e.stopPropagation(); openInvoiceDetail(inv); };
  td.appendChild(link);
  const qurl = qboInvoiceUrl(inv.qbo_txn_id);
  if (qurl) {
    const a = document.createElement("a"); a.href = qurl; a.target = "_blank"; a.rel = "noopener";
    a.className = "qbo-ico"; a.textContent = "qb"; a.title = "Open this invoice in QuickBooks";
    a.onclick = (e) => e.stopPropagation(); td.appendChild(a);
  }
  return td;
}

// The invoice's memo + every field in the side panel - read a draw/invoice without opening QBO.
// Works for an Invoices row and a Draws row alike (both carry the same billing_event fields).

// ── The PROJECT page (owner 2026-09-02): everything about one job in one place - section 1 "how
// it's doing" (WIP + live P&L + the trail), section 2 "how we get funded" (draws in order, GC paid,
// vendors x/y paid, the funding-chain math, pay-to-unlock checkboxes on the existing pay run, export),
// then bills / links. Opened from any project # in the app. Read-only except the pay-run marks.
let _pp = null;
async function openProjectPage(pn) {
  if (!pn || !/^(MFD|CP|RP)\d/i.test(String(pn))) { if (pn) toast(`"${pn}" is not a project # - nothing to open`); return; }   // e.g. the "(multiple)" bucket
  pn = String(pn).toUpperCase();
  const r0 = (ALL || []).find(x => x.project_no === pn) || {};
  openRecord(pn + (r0.project_name ? " · " + r0.project_name : ""), [r0.division, r0.status ? "WIP status " + r0.status : ""].filter(Boolean).join(" · "));
  const body = $("#recordBody"); body.innerHTML = ""; skeletonInto(body, 6);
  let d;
  try { d = await (await fetch(`/api/project/page?no=${encodeURIComponent(pn)}`)).json(); }
  catch (e) { body.textContent = "could not load this project"; return; }
  if (!d || !d.ok) { body.textContent = (d && d.error) || "no data for this project"; return; }
  _pp = { d, pn, open: new Set(), openV: new Set(), filter: "unpaid" };   // open = draws, openV = vendor groups (collapsed by default, always)
  body.innerHTML = "";
  if (!r0.project_name && d.project && d.project.name) $("#recordTitle").textContent = `${pn} · ${d.project.name}`;
  const sec = (title, note) => { const w = document.createElement("section"); w.className = "widget ip-sec";
    const h = document.createElement("div"); h.className = "widget-head"; h.innerHTML = `<h2>${_ge(title)} <span class="count">${_ge(note || "")}</span></h2>`; w.appendChild(h); body.appendChild(w); return w; };
  const kpi = (host, items) => { const strip = document.createElement("div"); strip.className = "kpi-row ip-strip";
    for (const [l, v, sub, cls] of items) { const k = document.createElement("div"); k.className = "kpi" + (cls ? " " + cls : "");
      k.innerHTML = `<div class="k-label"></div><div class="k-value"></div><div class="k-sub"></div>`;
      k.querySelector(".k-label").textContent = l; k.querySelector(".k-value").textContent = v; k.querySelector(".k-sub").textContent = sub || ""; strip.appendChild(k); }
    host.appendChild(strip); };
  // ── 1. how it's doing ──
  const p = d.pnl || {};
  const s1 = sec("How it's doing", `WIP master report ${r0.report_date ? fmtDate(r0.report_date) : "–"} · QuickBooks costs loaded ${loadedAt("Costs (QBO)") ? fmtDate(loadedAt("Costs (QBO)"), true) : "–"}`);
  const gp = num(r0.total_contract_price) - num(r0.estimated_total_costs);
  kpi(s1, [
    ["Contract", money(p.contract || r0.total_contract_price), r0.approved_cos ? `incl. COs ${money(r0.approved_cos)}` : ""],
    ["ETC (budget)", money(r0.estimated_total_costs), r0.total_contract_price ? `planned GP ${money(gp)} · ${(gp / num(r0.total_contract_price) * 100).toFixed(1)}%` : ""],
    ["Costs to date (QuickBooks)", money(p.cost), r0.costs_to_date != null ? `WIP report ${money(r0.costs_to_date)}` : "", num(p.cost) > num(r0.estimated_total_costs) && r0.estimated_total_costs ? "pnl-kpi-neg" : ""],
    ["Billed to date", money(p.billed), r0.billed_to_date != null ? `WIP (gross) ${money(r0.billed_to_date)}` : ""],
    ["Earned revenue", money(p.earned), `${((p.pct_complete || 0) * 100).toFixed(1)}% complete`],
    ["Net margin (live P&L)", money(p.net), p.net_pct != null ? `${(p.net_pct * 100).toFixed(1)}% of earned · overhead ${p.overhead_basis || ""}` : "", num(p.net) < 0 ? "pnl-kpi-neg" : "pnl-kpi-pos"],
  ]);
  const acts1 = document.createElement("div"); acts1.className = "ip-actions";
  const trailBtn = document.createElement("button"); trailBtn.className = "btn small"; trailBtn.textContent = "Show every dollar"; trailBtn.onclick = () => openTrail(pn); acts1.appendChild(trailBtn);
  if (r0.project_no) { const dr = document.createElement("button"); dr.className = "btn small"; dr.textContent = "WIP row detail"; dr.onclick = () => openDetail(r0); acts1.appendChild(dr); }
  s1.appendChild(acts1);
  const plWrap = document.createElement("div"); plWrap.className = "ip-top pp-pnl"; plWrap.appendChild(buildPnlGroup(pn)); s1.appendChild(plWrap);   // 3 columns (owner: save vertical space)
  // ── 2. how we get funded ──
  const F = d.funding || {}, nx = F.next_draw;
  const s2 = sec("How we get funded", `${d.draws.length} draw${d.draws.length === 1 ? "" : "s"} · GC owes ${money(d.draws.reduce((s, x) => s + num(x.ar_open), 0))}`);
  const unlock = document.createElement("div"); unlock.className = "pp-unlock" + (nx ? "" : " ok");
  if (nx) {
    const blk = F.blockers || [];
    unlock.innerHTML = `<div class="pp-unlock-h">Next money in: <b>${_ge(nx.label.split(" — ")[0])}${nx.draw_no ? " · Draw #" + nx.draw_no : ""}</b> · GC owes <b>${_ge(money(nx.ar_open))}</b>${nx.ar_date ? " · invoiced " + _ge(fmtDate(nx.ar_date)) : ""}</div>`
      + (blk.length ? (F.blockers_total > 0.005
            ? `<div class="pp-unlock-b">Blocked by <b>${blk.length}</b> unpaid bill${blk.length === 1 ? "" : "s"} on earlier draws · <b>${_ge(money(F.blockers_total))}</b> to pay (their unconditional waivers release this draw)</div>`
            : `<div class="pp-unlock-b"><b>${blk.length}</b> bill${blk.length === 1 ? "" : "s"} on earlier draws show no payment date yet ($0 open) - confirm they are paid and collect the waivers, then this draw is clear on our side</div>`)
                    : `<div class="pp-unlock-b ok">No unpaid bills on earlier draws - nothing on our side blocks this draw${F.own_unpaid > 0.005 ? `; ${_ge(money(F.own_unpaid))} of its own bills still to pay once funded` : ""}.</div>`);
  } else unlock.innerHTML = `<div class="pp-unlock-h ok">Nothing outstanding - the GC has paid every draw on file.</div>`;
  s2.appendChild(unlock);
  const tools = document.createElement("div"); tools.className = "ip-tools";
  const seg = document.createElement("div"); seg.className = "seg";
  for (const [k, lbl] of [["unpaid", "Unpaid bills"], ["all", "All bills"]]) { const b = document.createElement("button"); b.type = "button"; b.className = "seg-btn" + (k === "unpaid" ? " on" : ""); b.textContent = lbl;
    b.onclick = () => { _pp.filter = k; seg.querySelectorAll(".seg-btn").forEach(x => x.classList.toggle("on", x === b)); _renderPpDraws(); }; seg.appendChild(b); }
  // three views (owner 2026-09-02): Draws = bands only · Vendors = every draw open, vendor totals collapsed · Bills = everything open
  const tog = document.createElement("div"); tog.className = "seg"; tog.id = "ppToggle";
  for (const [k, lbl, title] of [["draws", "Draws", "Just the draw bands"], ["vendors", "Vendors", "Every draw open to its vendor totals"], ["bills", "Bills", "Every vendor open to its bills"]]) {
    const b = document.createElement("button"); b.type = "button"; b.className = "seg-btn" + (k === "draws" ? " on" : ""); b.dataset.v = k; b.textContent = lbl; b.title = title;
    b.onclick = () => { const keys = d.draws.map(x => x.matched_invoice);
      if (k === "draws") { _pp.open = new Set(); _pp.openV = new Set(); }
      else if (k === "vendors") { _pp.open = new Set(keys); _pp.openV = new Set(); }
      else { _pp.open = new Set(keys); _pp.openV = new Set(["*"]); }
      _renderPpDraws(); };
    tog.appendChild(b);
  }
  const mk = document.createElement("button"); mk.type = "button"; mk.className = "btn small"; mk.textContent = "Mark blockers to pay"; mk.title = "Tick every unpaid bill on the draws before the next one the GC owes - they go on the pay run";
  mk.onclick = () => _ppMarkBlockers();
  const ex = document.createElement("button"); ex.type = "button"; ex.className = "btn small"; ex.textContent = "Export pay list"; ex.title = "Excel report of the bills ticked to pay on this job, grouped by draw";
  ex.onclick = () => _ppExport();
  const pb = document.createElement("button"); pb.type = "button"; pb.className = "btn small subtle"; pb.textContent = "Open Pay Bills"; pb.onclick = () => setTab("paybills");
  tools.appendChild(seg); tools.appendChild(tog); tools.appendChild(mk); tools.appendChild(ex); tools.appendChild(pb); s2.appendChild(tools);
  const sel = document.createElement("div"); sel.id = "ppSelected"; sel.className = "pp-selected"; s2.appendChild(sel);   // what is ticked, at a glance, before any export
  const host = document.createElement("div"); host.id = "ppDraws"; s2.appendChild(host);
  _renderPpDraws();
  // ── 3. bills + links ──
  const s3 = sec("Bills and links", "");
  const acts3 = document.createElement("div"); acts3.className = "ip-actions";
  const bb = document.createElement("button"); bb.className = "btn small"; bb.textContent = "Bills on this job"; bb.title = "The Bill Tracker filtered to this project";
  bb.onclick = () => { if (typeof billMSel === "object") { for (const c of BILL_MSEL) billMSel[c.id] = new Set(); billMSel["bfProject"] = new Set([pn]); } activeBillView = "all"; setTab("bills"); if (typeof buildBillFilters === "function") buildBillFilters(); renderBills(); };
  acts3.appendChild(bb);
  const ib = document.createElement("button"); ib.className = "btn small"; ib.textContent = "Invoices on this job"; ib.onclick = () => { invMSel["ifProj"] = new Set([pn]); _invMSelSig = null; setTab("invoices"); renderOpenInvoices(); }; acts3.appendChild(ib);
  const cid = (COST.by_project && COST.by_project[pn] && COST.by_project[pn].customer_id) || ((invData().invoices || []).find(i => i.project_no === pn) || {}).cust_id;
  const qurl = qboCustomerUrl(cid); if (qurl) { const a = document.createElement("a"); a.className = "btn small"; a.href = qurl; a.target = "_blank"; a.rel = "noopener"; a.textContent = "Project in QuickBooks ↗"; acts3.appendChild(a); }
  s3.appendChild(acts3);
}
function _ppBillShown(b) { const paid = !!b.paid; return _pp.filter === "all" || !paid; }
function _renderPpSelected() {
  const box = $("#ppSelected"); if (!box) return;
  const picked = _ppAllSelected();
  box.innerHTML = "";
  if (!picked.length) { box.classList.add("empty"); box.textContent = "Nothing ticked to pay on this job yet - tick bills below, or Mark blockers to pay."; return; }
  box.classList.remove("empty");
  const tot = picked.reduce((s, x) => s + num(x.b.open), 0);
  const h = document.createElement("div"); h.className = "pp-sel-h"; h.textContent = `Selected to pay: ${picked.length} bill${picked.length === 1 ? "" : "s"} · ${money(tot)}`; box.appendChild(h);
  const ul = document.createElement("div"); ul.className = "pp-sel-list";
  const byV = new Map(); for (const x of picked) { const k = x.b.vendor || "?"; if (!byV.has(k)) byV.set(k, []); byV.get(k).push(x); }
  for (const [v, xs] of [...byV].sort((a, b) => a[0].localeCompare(b[0]))) {
    const li = document.createElement("div"); li.className = "pp-sel-v";
    li.innerHTML = `<b>${_ge(v)}</b> · ${xs.length} bill${xs.length === 1 ? "" : "s"} · ${_ge(money(xs.reduce((s, x) => s + num(x.b.open), 0)))} <span class="dim">(${xs.map(x => _ge(x.b.bill_ref || "?") + " on " + _ge(x.dr.invoice_no ? "inv " + x.dr.invoice_no : "no draw")).join(", ")})</span>`;
    ul.appendChild(li);
  }
  box.appendChild(ul);
  const clr = document.createElement("button"); clr.type = "button"; clr.className = "btn small subtle"; clr.textContent = "Untick all on this job"; clr.onclick = () => _ppSetPay(picked.map(x => x.b), false); box.appendChild(clr);
}
function _ppBillsOf(dr) { return [...(dr.bills || []), ...(dr.sub_bills || [])]; }
function _ppAllSelected() { const out = []; for (const dr of _pp.d.draws) for (const b of _ppBillsOf(dr)) if (b.pay_selected) out.push({ dr, b }); return out; }
function _ppCheck(bills, label, title) {   // a select-all checkbox for a group of bills (unpaid, payable by us)
  const payable = bills.filter(b => !b.paid && b.gates && b.bill_id);
  const cb = document.createElement("input"); cb.type = "checkbox"; cb.className = "pp-grp-cb";
  cb.checked = payable.length > 0 && payable.every(b => b.pay_selected);
  cb.indeterminate = !cb.checked && payable.some(b => b.pay_selected);
  cb.disabled = !payable.length; cb.title = title || `Tick every unpaid bill in ${label}`;
  cb.onclick = (e) => e.stopPropagation();
  cb.onchange = () => _ppSetPay(payable, cb.checked);
  return cb;
}
function _renderPpDraws() {
  const { d, host } = { d: _pp.d, host: $("#ppDraws") }; if (!host) return; host.innerHTML = "";
  _renderPpSelected();
  const keys = d.draws.map(x => x.matched_invoice); const allOpen = keys.length && keys.every(k => _pp.open.has(k));
  { const tg = $("#ppToggle"); if (tg) { const v = !_pp.open.size ? "draws" : (allOpen && _pp.openV.has("*")) ? "bills" : (allOpen && !_pp.openV.size) ? "vendors" : ""; tg.querySelectorAll(".seg-btn").forEach(b => b.classList.toggle("on", b.dataset.v === v)); } }
  if (!d.draws.length) { const p = document.createElement("div"); p.className = "bills-cap"; p.textContent = "No draws or bills on this job in the Bill Tracker yet."; host.appendChild(p); return; }
  const nxInv = d.funding && d.funding.next_draw ? d.funding.next_draw.invoice_no : null;
  { const hd = document.createElement("div"); hd.className = "pp-draw-h pp-cols";
    hd.innerHTML = `<span></span><span>Draw</span><span>Invoiced</span><span>Billed</span><span>GC</span><span>Materials</span><span>Labor</span><span>Pay run</span><span>Stage</span>`; host.appendChild(hd); }
  const billRow = (dr, b, isSub) => {
    const tr = document.createElement("tr"); if (b.paid) tr.classList.add("inv-paid");
    const pc = document.createElement("td"); pc.className = "left";
    if (!b.paid && b.gates && b.bill_id) { const cb = document.createElement("input"); cb.type = "checkbox"; cb.checked = !!b.pay_selected; cb.title = "Put this bill on the pay run (Pay Bills) - local intent, never QuickBooks";
      cb.onclick = (e) => e.stopPropagation(); cb.onchange = () => _ppSetPay([b], cb.checked); pc.appendChild(cb); }
    else if (!b.gates) { const s = document.createElement("span"); s.className = "vg-tag"; s.textContent = "GC pays"; s.title = "Concrete pumping - paid by the GC directly"; pc.appendChild(s); }
    tr.appendChild(pc);
    tr.appendChild(leftText(""));   // vendor is the group header
    tr.appendChild(qboLinkCell(b.bill_ref || "–", isSub ? qboUrl(b.txn_type === "Expense" ? "expense" : "bill", b.bill_id) : qboBillHref(b.qbo_link), "Open this bill in QuickBooks"));
    tr.appendChild(leftText(fmtDateShort(b.bill_date)));
    { const mc = leftText(isSub ? (b.memo || (b.codes || []).join(", ") || "–") : "–"); mc.className += " inv-memo"; mc.title = isSub ? (b.memo || "") : ""; tr.appendChild(mc); }
    const ac = document.createElement("td"); ac.className = "ip-amt"; ac.appendChild(moneyCell(b.amount)); tr.appendChild(ac);
    const oc = document.createElement("td"); oc.className = "right ip-amt"; oc.textContent = num(b.open) > 0.005 ? money(b.open) : "–"; if (num(b.open) > 0.005) oc.style.color = "var(--neg)"; else oc.classList.add("dim"); tr.appendChild(oc);
    const st = document.createElement("td"); st.className = "left"; const pill = document.createElement("span"); pill.className = b.paid ? "ar-paid" : "ar-open";
    pill.textContent = isSub ? (b.paid ? "Paid " + fmtDateShort(b.pay_date) : "No payment on file") : (b.pay_date ? "Paid " + fmtDateShort(b.pay_date) : (b.paid ? (b.pay_status || "Paid") + " (no date)" : (b.pay_status || "Open")));
    if (isSub && !b.paid) pill.title = "No QuickBooks bill payment applied to this bill in the loaded window (this year)"; st.appendChild(pill); tr.appendChild(st);
    { const wc = document.createElement("td"); wc.className = "left";
      if (dr.no_draw || isSub) wc.textContent = isSub ? "–" : "–";
      else { const cb = document.createElement("input"); cb.type = "checkbox"; cb.checked = !!b.waiver; cb.title = "Tick when the vendor's unconditional waiver is in hand";
        cb.onclick = (e) => e.stopPropagation(); cb.onchange = () => { setWaiver(dr, b, cb); b.waiver = cb.checked; }; wc.appendChild(cb); if (!b.waiver && !b.paid) { const s = document.createElement("span"); s.className = "inv-late"; s.textContent = " needed"; wc.appendChild(s); } }
      tr.appendChild(wc); }
    return tr;
  };
  const sectionRow = (tbody, dr, label, bills, sect) => {   // level 2: the section total (Materials $X / Labor $Y) with its own select-all
    const paidCt = bills.filter(b => b.paid).length, tot = bills.reduce((s, b) => s + num(b.amount), 0), owed = bills.reduce((s, b) => s + (b.paid ? 0 : num(b.open)), 0);
    const sr = document.createElement("tr"); sr.className = "pp-sect";
    const cbTd = document.createElement("td"); cbTd.className = "left"; cbTd.appendChild(_ppCheck(bills, label, `Tick every unpaid ${label.toLowerCase()} bill on this draw`)); sr.appendChild(cbTd);
    const td = document.createElement("td"); td.className = "left"; td.colSpan = 8;
    td.innerHTML = `<b>${_ge(label)}</b> <span class="ip-paid ${paidCt === bills.length ? "ok" : "due"}">${_ge(money(tot))} · ${paidCt}/${bills.length} paid${owed > 0.005 ? " · " + _ge(money(owed)) + " to pay" : ""}</span>` + (sect === "labor" ? ` <span class="dim">QuickBooks bills dated in the draw period · paid = a bill payment applied this year</span>` : ` <span class="dim">Bill Tracker</span>`);
    sr.appendChild(td); tbody.appendChild(sr);
  };
  const vendorGroups = (tbody, dr, bills, isSub, sect) => {   // level 2 vendors (collapsed) -> level 3 bills
    const byV = new Map(); for (const b of bills) { const k = b.vendor || "?"; if (!byV.has(k)) byV.set(k, []); byV.get(k).push(b); }
    for (const [v, list] of [...byV].sort((a, b) => a[0].localeCompare(b[0]))) {
      const vkey = `${dr.matched_invoice}|${sect}|${v}`, vopen = _pp.openV.has("*") || _pp.openV.has(vkey);
      const paidCt = list.filter(b => b.paid).length, tot = list.reduce((s, b) => s + num(b.amount), 0), owed = list.reduce((s, b) => s + (b.paid ? 0 : num(b.open)), 0);
      const gtr = document.createElement("tr"); gtr.className = "bill-subgroup pp-vendor"; gtr.style.cursor = "pointer"; gtr.title = vopen ? "Click to collapse" : "Click to see the bills";
      const cbTd = document.createElement("td"); cbTd.className = "left"; cbTd.appendChild(_ppCheck(list, v)); gtr.appendChild(cbTd);
      const vt = document.createElement("td"); vt.className = "left"; vt.colSpan = 8;
      const cell = document.createElement("div"); cell.className = "bg-cell"; const caret = document.createElement("span"); caret.className = "bg-caret"; caret.textContent = vopen ? "▾ " : "▸ ";
      const k = document.createElement("span"); k.className = "sg-key"; k.textContent = v;
      const leftP = document.createElement("span"); leftP.className = "bg-left"; leftP.appendChild(caret); leftP.appendChild(k); cell.appendChild(leftP);
      bandMetrics(cell, [[money(tot), "total"], [list.length, "bills"], [`${paidCt}/${list.length}`, "paid", paidCt === list.length ? "ok" : "due"], [owed > 0.005 ? money(owed) : "–", "to pay", owed > 0.005 ? "neg" : ""]]);
      vt.appendChild(cell); gtr.appendChild(vt);
      gtr.onclick = (e) => { if (e.target.closest("input")) return; if (_pp.openV.has("*")) { _pp.openV = new Set(); for (const d2 of _pp.d.draws) for (const s2 of ["materials", "labor"]) for (const b2 of (s2 === "labor" ? (d2.sub_bills || []) : d2.bills)) _pp.openV.add(`${d2.matched_invoice}|${s2}|${b2.vendor || "?"}`); }
        if (_pp.openV.has(vkey)) _pp.openV.delete(vkey); else _pp.openV.add(vkey); _renderPpDraws(); };
      tbody.appendChild(gtr);
      if (vopen) for (const b of list) tbody.appendChild(billRow(dr, b, isSub));
    }
  };
  for (const dr of d.draws) {
    const key = dr.matched_invoice, open = _pp.open.has(key);
    const band = document.createElement("div"); band.className = "pp-draw" + (dr.invoice_no && dr.invoice_no === nxInv ? " next" : "");
    const head = document.createElement("div"); head.className = "pp-draw-h"; head.style.cursor = "pointer";
    const paidAll = dr.vendors_total > 0 && dr.vendors_paid === dr.vendors_total;
    const nSub = (dr.sub_bills || []).length;
    const cell2 = (a, b, cls) => `<span class="pp-c ${cls || ""}"><span class="pp-c1">${a}</span><span class="pp-c2">${b}</span></span>`;
    head.innerHTML = `<span class="bg-caret">${open ? "▾" : "▸"}</span>
      <span class="pp-lab">${_ge(dr.no_draw ? "No draw yet" : "Invoice " + (dr.invoice_no || dr.label.split(" — ")[0]))}${dr.draw_no ? `<small class="pp-drawno">Draw #${dr.draw_no}</small>` : ""}</span>
      <span class="pp-dt">${_ge(dr.ar_date ? fmtDate(dr.ar_date) : "–")}</span>
      <span class="pp-billed">${dr.no_draw ? "–" : _ge(money(dr.billed))}</span>
      <span class="pp-gc ${dr.no_draw ? "" : dr.gc_paid ? "ok" : "due"}">${dr.no_draw ? "–" : dr.gc_paid ? "paid" : "owes " + _ge(money(dr.ar_open))}</span>
      ${cell2(_ge(money(dr.gate_amt)), `${dr.vendors_paid}/${dr.vendors_total} paid${dr.unpaid_amt > 0.005 ? " · " + _ge(money(dr.unpaid_amt)) + " to pay" : ""}`, "ip-paid " + (paidAll ? "ok" : "due"))}
      ${nSub ? cell2(_ge(money(dr.subs_amt)), `${dr.subs_paid_ct}/${nSub} paid${dr.subs_unpaid_amt > 0.005 ? " · " + _ge(money(dr.subs_unpaid_amt)) + " to pay" : ""}`, "ip-paid " + (dr.subs_paid_ct === nSub ? "ok" : "due")) : cell2("$0", dr.no_draw ? "" : "none in period", "dim")}
      <span class="pp-allcell"></span>
      <span class="pp-stage">${_ge(dr.stage)}</span>`;
    // draw-level select-all: every unpaid bill we pay on this draw (materials + labor)
    const allCb = _ppCheck(_ppBillsOf(dr), "this draw", "Tick every unpaid bill on this draw (materials and labor)");
    const allWrap = document.createElement("label"); allWrap.className = "pp-all"; allWrap.title = allCb.title; allWrap.onclick = (e) => e.stopPropagation();
    allWrap.appendChild(allCb); allWrap.appendChild(document.createTextNode(" all unpaid")); head.querySelector(".pp-allcell").appendChild(allWrap);
    head.onclick = () => { if (_pp.open.has(key)) _pp.open.delete(key); else _pp.open.add(key); _renderPpDraws(); };
    band.appendChild(head);
    if (open) {
      const mat = dr.bills.filter(_ppBillShown), subs = (dr.sub_bills || []).filter(_ppBillShown);
      if (!mat.length && !subs.length) { const p = document.createElement("div"); p.className = "bills-cap"; p.textContent = _pp.filter === "unpaid" ? "Every bill on this draw is paid." : "No bills."; band.appendChild(p); }
      else {
        const table = document.createElement("table"); table.className = "grid pp-bills"; const thead = document.createElement("thead"), tbody = document.createElement("tbody");
        thead.innerHTML = "<tr><th class='left'>Pay</th><th class='left'>Vendor</th><th class='left'>Bill #</th><th class='left'>Date</th><th class='left'>Memo</th><th class='right'>Amount</th><th class='right'>Open</th><th class='left'>Status</th><th class='left'>Waiver received</th></tr>";
        if (mat.length) { sectionRow(tbody, dr, "Materials", mat, "materials"); vendorGroups(tbody, dr, mat, false, "materials"); }
        if (subs.length) { sectionRow(tbody, dr, "Labor (subs)", subs, "labor"); vendorGroups(tbody, dr, subs, true, "labor"); }
        thead.hidden = ![...mat, ...subs].some(b => _pp.openV.has("*") || _pp.openV.has(`${dr.matched_invoice}|${dr.sub_bills && dr.sub_bills.includes(b) ? "labor" : "materials"}|${b.vendor || "?"}`));   // headers only once bills show
        table.appendChild(thead); table.appendChild(tbody); band.appendChild(table);
      }
    }
    host.appendChild(band);
  }
}
async function _ppSetPay(bills, selected) {
  const items = bills.filter(b => b.bill_id).map(b => ({ bill_id: b.bill_id, selected, amount: null }));
  if (!items.length) { toast("These bills have no QuickBooks link to key the pay run on"); return; }
  try {
    const r = await (await fetch("/api/pay-run", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ items }) })).json();
    if (r && r.ok) { for (const b of bills) b.pay_selected = selected; toast(selected ? `${items.length} bill${items.length === 1 ? "" : "s"} added to the pay run` : "Removed from the pay run"); }
    else toast("Could not save: " + ((r && r.error) || "unknown"));
  } catch (e) { toast("Could not save: " + e); }
  _renderPpDraws();
}
function _ppMarkBlockers() {
  const F = _pp.d.funding || {}, ids = new Set((F.blockers || []).map(b => b.bill_id).filter(Boolean));
  const bills = []; for (const dr of _pp.d.draws) for (const b of _ppBillsOf(dr)) if (ids.has(b.bill_id)) bills.push(b);
  if (!bills.length) { toast("No blockers to mark"); return; }
  _ppSetPay(bills, true);
}
function _ppExport() {
  const d = _pp.d, rows = [];
  for (const dr of d.draws) for (const b of _ppBillsOf(dr)) if (b.pay_selected) rows.push([dr.no_draw ? "No draw yet" : "Invoice " + (dr.invoice_no || "") + (dr.draw_no ? " · Draw #" + dr.draw_no : ""), b.vendor, b.bill_ref, b.bill_date, num(b.amount), num(b.open), b.pay_date ? "Paid " + fmtDate(b.pay_date) : (b.pay_status || (b.paid ? "Paid" : "Open")), dr.sub_bills && dr.sub_bills.includes(b) ? "labor" : (b.waiver ? "received" : "needed")]);
  if (!rows.length) { toast("Nothing ticked to pay on this job yet - tick bills (or Mark blockers) first"); return; }
  if (!confirm(`Export ${rows.length} bill${rows.length === 1 ? "" : "s"} ticked to pay (${money(rows.reduce((s, r) => s + num(r[5]), 0))} open) as the pay-list report?\n\nThe list above the draws shows exactly what is ticked.`)) return;
  const nx = (d.funding || {}).next_draw, toPay = rows.reduce((s, r) => s + num(r[5]), 0);
  const footer = nx ? [{ label: `Unlocks ${nx.invoice_no || ""}${nx.draw_no ? " · Draw #" + nx.draw_no : ""}`, value: num(nx.ar_open) },
                       { label: "Net = unlock - to pay", value: num(nx.ar_open) - toPay, cls: num(nx.ar_open) - toPay >= 0 ? "pos" : "neg" }] : [];
  const body = { name: `Pay list ${_pp.pn}`, sheet: "Pay list", title: `${_pp.pn} - bills to pay to unlock the next draw${nx ? " " + money(nx.ar_open) + " - Invoice " + (nx.invoice_no || "") + (nx.draw_no ? " · Draw #" + nx.draw_no : "") : ""}`, footer,
    subtitle: `${rows.length} bills ticked on the pay run · exported ${fmtDate(new Date(Date.now() - new Date().getTimezoneOffset() * 60000).toISOString().slice(0, 19), true)}`,
    columns: [{ label: "Draw" }, { label: "Vendor" }, { label: "Bill #" }, { label: "Bill date" }, { label: "Amount", type: "money" }, { label: "Open", type: "money" }, { label: "Status" }, { label: "Waiver" }],
    rows, group_by: 0, fmt: rows.map((r, i) => ({ r: i, c: 5, cls: r[5] > 0 ? "neg" : "pos" })) };
  toast("Building the Excel report…");
  fetch("/api/export/xlsx", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) }).then(r => r.json())
    .then(r => toast(r && r.ok ? `Pay list saved to Downloads (${r.rows} bills) - opened in Finder` : "Export failed: " + ((r && r.error) || "unknown"))).catch(e => toast("Export failed: " + e));
}
// ── The invoice as a PAGE (owner 2026-09-02): QBO details on top, the draw's bills grouped by vendor
// with pay status (+ subs from QuickBooks), then the Notion collections log. Full-width record view.
// ── invoice page: the bills block (vendors + subs), re-rendered on filter / collapse ──
let _ip = null;
function _ipGroups() { const d = _ip.d; return [...(d.vendors || []).map(v => "v:" + v.vendor), ...(d.subs || []).map(v => "s:" + v.vendor)]; }
function _ipBillPasses(b) { const paid = !!b.pay_date || num(b.open) <= 0.005; return _ip.filter === "all" || (_ip.filter === "paid" ? paid : !paid); }
function _copyIpBills() {
  const { d } = _ip, i = d.invoice;
  const vendors = (d.vendors || []).map(v => ({ ...v, bills: v.bills.filter(_ipBillPasses) })).filter(v => v.bills.length);
  const lines = [`Invoice ${i.doc_number} · ${i.customer || ""} · ${i.project_no || ""} ${nameOf(i.project_no) || ""} · ${_ip.filter === "all" ? "all bills" : _ip.filter + " bills"} · as of ${fmtDate(new Date().toISOString().slice(0, 10))}`,
                 ["Vendor", "Bill date", "Bill #", "Amount", "Open", "Pay status", "Approved", "Lien", "GC paid us"].join("\t")];
  let tAmt = 0, tOpen = 0;
  for (const v of vendors) {
    let vAmt = 0, vOpen = 0;
    for (const b of v.bills) { vAmt += num(b.amount); vOpen += num(b.open);
      lines.push([v.vendor, fmtDate(b.bill_date), b.bill_ref || "", Math.round(num(b.amount) * 100) / 100, Math.round(num(b.open) * 100) / 100,
                  b.pay_date ? "Paid " + fmtDate(b.pay_date) : (b.pay_status || "Open"), b.approved || "", b.lien_status || "", b.gc_paid ? fmtDate(b.gc_paid) : ""].join("\t")); }
    lines.push([v.vendor + " total", "", "", Math.round(vAmt * 100) / 100, Math.round(vOpen * 100) / 100, "", "", "", ""].join("\t"));
    tAmt += vAmt; tOpen += vOpen;
  }
  lines.push(["TOTAL", "", "", Math.round(tAmt * 100) / 100, Math.round(tOpen * 100) / 100, "", "", "", ""].join("\t"));
  copy(lines.join("\n"));
}
function _renderIpBills() {
  const { d, host } = _ip; host.innerHTML = "";
  const i = d.invoice;
  const groups = _ipGroups(); const allOpen = groups.length && groups.every(g => _ip.open.has(g));
  { const tg = $("#ipToggle"); if (tg) tg.textContent = allOpen ? "Collapse all" : "Expand all"; }
  const vendors = (d.vendors || []).map(v => ({ ...v, bills: v.bills.filter(_ipBillPasses) })).filter(v => v.bills.length);
  if (!(d.vendors || []).length) { const p = document.createElement("div"); p.className = "bills-cap"; p.textContent = "No Bill Tracker bills carry this invoice # yet (the AP sync matches bills to draws; RP jobs bill at completion)."; host.appendChild(p); }
  else if (!vendors.length) { const p = document.createElement("div"); p.className = "bills-cap"; p.textContent = _ip.filter === "paid" ? "No paid bills on this draw yet." : "No unpaid bills on this draw - every vendor is paid."; host.appendChild(p); }
  else {
    const scroll = document.createElement("div"); scroll.className = "table-scroll";
    const table = document.createElement("table"); table.className = "grid"; const thead = document.createElement("thead"), tbody = document.createElement("tbody");
    const htr = document.createElement("tr");
    for (const [c, al] of [["Vendor / bill", "left"], ["Date", "left"], ["Bill #", "left"], ["Amount", "right"], ["Open", "right"], ["Pay status", "left"], ["Approved", "left"], ["Lien", "left"], ["GC paid us", "left"]]) { const th = document.createElement("th"); if (al === "left") th.className = "left"; th.textContent = c; htr.appendChild(th); }
    thead.appendChild(htr);
    thead.hidden = !vendors.some(v => _ip.open.has("v:" + v.vendor));   // headers only once a group is open
    for (const v of vendors) {
      const key = "v:" + v.vendor, open = _ip.open.has(key);
      const tot = v.bills.reduce((s, b) => s + num(b.amount), 0), opn = v.bills.reduce((s, b) => s + num(b.open), 0), paidCt = v.bills.filter(b => b.pay_date).length;
      const paidAmt = v.bills.reduce((s, b) => s + (b.pay_date ? num(b.amount) : Math.max(0, num(b.amount) - num(b.open))), 0);
      const gtr = document.createElement("tr"); gtr.className = "bill-group"; gtr.style.cursor = "pointer"; gtr.title = open ? "Click to collapse" : "Click to expand";
      const gtd = document.createElement("td"); gtd.colSpan = 9;
      const cell = document.createElement("div"); cell.className = "bg-cell";
      const caret = document.createElement("span"); caret.className = "bg-caret"; caret.textContent = open ? "▾ " : "▸ ";
      const k = document.createElement("span"); k.className = "bg-key"; k.textContent = v.vendor;
      const leftV = document.createElement("span"); leftV.className = "bg-left"; leftV.appendChild(caret); leftV.appendChild(k); cell.appendChild(leftV);
      bandMetrics(cell, [[`${paidCt}/${v.bills.length}`, "paid", paidCt === v.bills.length ? "ok" : "due"], [money(paidAmt), "paid $"], [money(tot), "total"], [opn > 0.005 ? money(opn) : "–", "still owed", opn > 0.005 ? "neg" : ""]]);
      gtd.appendChild(cell); gtr.appendChild(gtd);
      gtr.onclick = () => { if (_ip.open.has(key)) _ip.open.delete(key); else _ip.open.add(key); _renderIpBills(); };
      tbody.appendChild(gtr);
      if (!open) continue;
      for (const b of v.bills) {
        const tr = document.createElement("tr");
        const nm = leftText(""); if (!b.gates) { const s = document.createElement("span"); s.className = "vg-tag"; s.textContent = "not paid by us"; s.title = "Concrete pumping - the GC pays this vendor directly"; nm.appendChild(s); } tr.appendChild(nm);
        tr.appendChild(leftText(fmtDateShort(b.bill_date)));
        tr.appendChild(qboLinkCell(b.bill_ref || "–", qboBillHref(b.qbo_link), "Open this bill in QuickBooks"));
        const ac = document.createElement("td"); ac.className = "ip-amt"; ac.appendChild(moneyCell(b.amount)); tr.appendChild(ac);
        const oc = document.createElement("td"); oc.className = "right ip-amt"; oc.textContent = num(b.open) > 0.005 ? money(b.open) : "–"; if (num(b.open) > 0.005) oc.style.color = "var(--neg)"; else oc.classList.add("dim"); tr.appendChild(oc);
        const st = document.createElement("td"); st.className = "left"; const pill = document.createElement("span"); pill.className = b.pay_date ? "ar-paid" : "ar-open"; pill.textContent = b.pay_date ? "Paid " + fmtDateShort(b.pay_date) : (b.pay_status || "Open"); st.appendChild(pill); tr.appendChild(st);
        tr.appendChild(leftText(b.approved || "–"));
        tr.appendChild(leftText(b.lien_status || "–"));
        tr.appendChild(leftText(b.gc_paid ? fmtDateShort(b.gc_paid) : "–"));
        tbody.appendChild(tr);
      }
    }
    table.appendChild(thead); table.appendChild(tbody); scroll.appendChild(table); host.appendChild(scroll);
  }
  if ((d.subs || []).length && _ip.filter === "all") {
    const cap = document.createElement("div"); cap.className = "bills-cap"; cap.textContent = `Subs (labor) on ${i.project_no} dated ${fmtDate(d.period.start)} – ${fmtDate(d.period.end)}, from QuickBooks - pay status is not tracked per sub bill, so they sit outside the Paid / Unpaid filter.`; host.appendChild(cap);
    const scroll = document.createElement("div"); scroll.className = "table-scroll";
    const table = document.createElement("table"); table.className = "grid"; const thead = document.createElement("thead"), tbody = document.createElement("tbody");
    const htr = document.createElement("tr"); for (const [c, al] of [["Sub / bill", "left"], ["Date", "left"], ["Bill #", "left"], ["Description", "left"], ["Code", "left"], ["Amount", "right"]]) { const th = document.createElement("th"); if (al === "left") th.className = "left"; th.textContent = c; htr.appendChild(th); } thead.appendChild(htr);
    thead.hidden = !d.subs.some(v => _ip.open.has("s:" + v.vendor));
    for (const v of d.subs) {
      const key = "s:" + v.vendor, open = _ip.open.has(key);
      const gtr = document.createElement("tr"); gtr.className = "bill-group"; gtr.style.cursor = "pointer"; const gtd = document.createElement("td"); gtd.colSpan = 6;
      const cell = document.createElement("div"); cell.className = "bg-cell"; const caret = document.createElement("span"); caret.className = "bg-caret"; caret.textContent = open ? "▾ " : "▸ ";
      const k = document.createElement("span"); k.className = "bg-key"; k.textContent = v.vendor;
      const leftS = document.createElement("span"); leftS.className = "bg-left"; leftS.appendChild(caret); leftS.appendChild(k); cell.appendChild(leftS);
      bandMetrics(cell, [[v.lines.length, "lines"], [money(v.total), "total"]]);
      gtd.appendChild(cell); gtr.appendChild(gtd);
      gtr.onclick = () => { if (_ip.open.has(key)) _ip.open.delete(key); else _ip.open.add(key); _renderIpBills(); };
      tbody.appendChild(gtr);
      if (!open) continue;
      for (const l of v.lines) { const tr = document.createElement("tr"); tr.appendChild(leftText("")); tr.appendChild(leftText(fmtDateShort(l.date)));
        tr.appendChild(qboLinkCell(l.doc_number || "–", qboUrl(l.txn_type === "Expense" ? "expense" : "bill", l.txn_id), "Open in QuickBooks"));
        const dc = leftText(l.description || "–"); dc.className += " inv-memo"; dc.title = l.description || ""; tr.appendChild(dc);
        const cc = document.createElement("td"); cc.className = "left"; if (l.cost_code) { const chip = document.createElement("span"); chip.className = "codechip"; chip.textContent = l.cost_code; cc.appendChild(chip); } tr.appendChild(cc);
        const ac = document.createElement("td"); ac.appendChild(moneyCell(l.amount)); tr.appendChild(ac); tbody.appendChild(tr); }
    }
    table.appendChild(thead); table.appendChild(tbody); scroll.appendChild(table); host.appendChild(scroll);
  }
}
async function openInvoicePage(inv) {
  const docn = inv.doc_number || inv.invoice_no || "";
  openRecord(`Invoice ${docn}`, [inv.customer, [inv.project_no, nameOf(inv.project_no)].filter(Boolean).join(" "), inv.division].filter(Boolean).join(" · "));   // the ONE identity line (owner: no repeats)
  const body = $("#recordBody"); body.innerHTML = ""; skeletonInto(body, 6);
  let d;
  try { d = await (await fetch(`/api/invoice/page?no=${encodeURIComponent(docn)}`)).json(); }
  catch (e) { body.textContent = "could not load this invoice"; return; }
  if (!d || !d.ok) { body.textContent = (d && d.error) || "no data for this invoice"; return; }
  body.innerHTML = "";
  const i = d.invoice;
  const sec = (title, note) => { const w = document.createElement("section"); w.className = "widget ip-sec";
    const h = document.createElement("div"); h.className = "widget-head"; h.innerHTML = `<h2>${_ge(title)} <span class="count">${_ge(note || "")}</span></h2>`; w.appendChild(h); body.appendChild(w); return w; };
  const kv = (host, rows) => { const g = document.createElement("div"); g.className = "ip-kv";
    for (const [k, v, cls] of rows) { if (v == null || v === "" || v === "—") continue;
      const r = document.createElement("div"); r.className = "drow"; const dk = document.createElement("span"); dk.className = "dk"; dk.textContent = k;
      const dv = document.createElement("span"); dv.className = "dv" + (cls ? " " + cls : ""); dv.textContent = v; dv.title = "Click to copy"; dv.onclick = () => copy(String(v));
      r.appendChild(dk); r.appendChild(dv); g.appendChild(r); } host.appendChild(g); return g; };
  // ── 1. the invoice, as QuickBooks has it ──
  const amt = num(i.amount), bal = i.balance == null ? null : num(i.balance), paidAmt = (amt != null && bal != null) ? amt - bal : null;
  const isOpen = !(bal != null && bal <= 0.005) && (i.status || "").toLowerCase() !== "paid";
  const s1 = sec("Invoice · QuickBooks", "");
  const top = document.createElement("div"); top.className = "ip-top"; s1.appendChild(top);
  const memoBox = document.createElement("div"); memoBox.className = "ip-memo" + (i.memo ? "" : " dim"); memoBox.textContent = i.memo || "(no memo on this invoice)"; memoBox.title = "Invoice memo"; top.appendChild(memoBox);
  const grid = document.createElement("div"); grid.className = "ip-grid"; top.appendChild(grid);
  const col = (title, rows) => { const c = document.createElement("div"); c.className = "dgroup"; const h = document.createElement("h4"); h.textContent = title; c.appendChild(h); kv(c, rows); grid.appendChild(c); };
  col("Billing", [["Invoice #", docn], ["Amount billed", money(amt), "ip-big"], ["Open balance", money(bal), "ip-big" + (bal != null && bal > 0.005 ? " neg" : "")], paidAmt != null ? ["Paid", money(paidAmt)] : null, ["Status", i.status || (bal > 0.005 ? "Open" : "Paid")]].filter(Boolean));
  col("Dates & terms", [["Invoice date", i.txn_date ? fmtDate(i.txn_date) : null], ["Due date", i.due_date ? fmtDate(i.due_date) : null],
    isOpen && i.days_past_due != null ? ["Days past due", i.days_past_due > 0 ? i.days_past_due + " days" : "current", i.days_past_due > 0 ? "neg" : ""] : null,
    ["Terms", i.net_terms], ["Draw period", d.period && d.period.start ? `${fmtDate(d.period.start)} – ${fmtDate(d.period.end)}` : i.draw_period], i.paid_date ? ["Paid date", fmtDate(i.paid_date)] : null].filter(Boolean));
  col("Lien", [["Notice deadline", i.lien_due_label], ["Lien status", i.lien_status], ["Notice type", i.lien_notice], ["Litigation", i.litigation ? "yes" : null]]);
  col("Collections", [["Last action", i.last_action_date ? fmtDate(i.last_action_date) : null], ["Next follow-up", i.next_followup ? fmtDate(i.next_followup) : null], ["Note", i.note]]);
  const acts = document.createElement("div"); acts.className = "ip-actions";
  const qurl = qboInvoiceUrl(i.qbo_txn_id); if (qurl) { const a = document.createElement("a"); a.className = "btn small"; a.href = qurl; a.target = "_blank"; a.rel = "noopener"; a.textContent = "Open in QuickBooks ↗"; acts.appendChild(a); }
  if (i.project_no) { const b = document.createElement("button"); b.className = "btn small primary"; b.textContent = "Project page"; b.onclick = () => openProjectPage(i.project_no); acts.appendChild(b); }
  const purl = qboCustomerUrl(i.cust_id); if (purl) { const a = document.createElement("a"); a.className = "btn small"; a.href = purl; a.target = "_blank"; a.rel = "noopener"; a.textContent = "Project in QuickBooks ↗"; acts.appendChild(a); }
  if (i.notion_url) { const a = document.createElement("a"); a.className = "btn small"; a.href = i.notion_url; a.target = "_blank"; a.rel = "noopener"; a.textContent = "Open in Notion ↗"; acts.appendChild(a); }
  s1.appendChild(acts);
  // ── 2. the bills on this draw, grouped by vendor, with pay status ──
  const T = d.totals || {};
  const s2 = sec("Bills on this draw · Bill Tracker, subs from QuickBooks", `${T.bills_paid || 0}/${T.bills || 0} paid · ${money((T.materials || 0) - (T.materials_open || 0))} / ${money(T.materials)}` + ((T.materials_open || 0) > 0.005 ? ` · ${money(T.materials_open)} still owed` : ""));
  { const cnt = s2.querySelector(".count"); if (cnt) cnt.classList.add("ip-paid", (T.bills_paid || 0) === (T.bills || 0) ? "ok" : "due"); }
  const strip = document.createElement("div"); strip.className = "kpi-row ip-strip";
  for (const [l, v, sub] of [["Billed to the GC", money(T.billed), i.status || ""], ["Materials we pay", money(T.materials_we_pay), T.materials !== T.materials_we_pay ? `${money(T.materials)} incl. pump vendors the GC pays` : ""],
                             ["Subs (labor)", money(T.subs), d.period && d.period.start ? "in the draw period" : "no period on the memo"], ["Net after bills + subs", money(T.net), T.net != null && T.net < 0 ? "bills exceed the draw" : ""]]) {
    const k = document.createElement("div"); k.className = "kpi" + (l.startsWith("Net") && T.net != null && T.net < 0 ? " pnl-kpi-neg" : "");
    k.innerHTML = `<div class="k-label"></div><div class="k-value"></div><div class="k-sub"></div>`;
    k.querySelector(".k-label").textContent = l; k.querySelector(".k-value").textContent = v; k.querySelector(".k-sub").textContent = sub; strip.appendChild(k);
  }
  s2.appendChild(strip);
  // vendor groups collapsed by default, Collapse/Expand all, and an All | Open | Paid filter
  // (owner 2026-09-02: "group the vendors so it's collapsed by default, show a toggle like invoices ...
  // have a filter to show paid and unpaid bills"). Re-rendered in place on every toggle.
  _ip = { d, open: new Set(), filter: "all", host: document.createElement("div") };
  const tools = document.createElement("div"); tools.className = "ip-tools";
  const seg = document.createElement("div"); seg.className = "seg";
  for (const [k, lbl] of [["all", "All bills"], ["open", "Unpaid"], ["paid", "Paid"]]) {
    const b = document.createElement("button"); b.type = "button"; b.className = "seg-btn" + (k === "all" ? " on" : ""); b.dataset.k = k; b.textContent = lbl;
    b.onclick = () => { _ip.filter = k; seg.querySelectorAll(".seg-btn").forEach(x => x.classList.toggle("on", x === b)); _renderIpBills(); }; seg.appendChild(b);
  }
  const tog = document.createElement("button"); tog.type = "button"; tog.className = "btn small"; tog.id = "ipToggle"; tog.textContent = "Expand all";
  tog.onclick = () => { const groups = _ipGroups(); const allOpen = groups.every(g => _ip.open.has(g)); _ip.open = new Set(allOpen ? [] : groups); _renderIpBills(); };
  const cp = document.createElement("button"); cp.type = "button"; cp.className = "btn small"; cp.textContent = "Copy bills"; cp.title = "Copy the bills shown (this filter), grouped by vendor with subtotals - paste into Excel or an email";
  cp.onclick = () => _copyIpBills();
  tools.appendChild(seg); tools.appendChild(tog); tools.appendChild(cp); s2.appendChild(tools);
  s2.appendChild(_ip.host);
  _renderIpBills();
  // ── 3. the Notion collections log ──
  const s3 = sec("Collections log · Notion Invoice Tracker", i.notion_edited ? `page edited ${fmtDate(i.notion_edited, true)}` : "");
  if (i.notion_url) invNotionSection(s3, i.notion_url, true);   // body + timestamp + comments only
  else { const p = document.createElement("div"); p.className = "bills-cap"; p.textContent = "This invoice came from QuickBooks directly - it has no Invoice Tracker page yet."; s3.appendChild(p); }
}
function openInvoiceDetail(inv) {
  if (!inv) return;
  const docn = inv.doc_number || inv.invoice_no || "—";
  $("#invDetailTitle").textContent = "Invoice " + docn;
  $("#invDetailSub").textContent = [inv.customer, inv.project_no, inv.division].filter(Boolean).join(" · ");
  const body = $("#invDetailBody"); body.innerHTML = "";
  // Memo first - the headline the owner asked for.
  const memo = (inv.memo == null ? "" : String(inv.memo)).trim();
  { const g = document.createElement("div"); g.className = "dgroup";
    const h = document.createElement("h4"); h.textContent = "Memo"; g.appendChild(h);
    const n = document.createElement("div"); n.className = "dnote" + (memo ? "" : " dim");
    n.textContent = memo || "(no memo on this invoice)"; g.appendChild(n); body.appendChild(g); }
  // Collections: the Notion Quick Status note, when it was last touched, and the tracker page itself
  // (owner 2026-09-02: "i need to do collections and need every single data point for meeting").
  { const g = document.createElement("div"); g.className = "dgroup";
    const h = document.createElement("h4"); h.textContent = "Collections"; g.appendChild(h);
    const n = document.createElement("div"); n.className = "dnote" + (inv.note ? "" : " dim");
    n.textContent = inv.note || "(no collections note in the Invoice Tracker)"; g.appendChild(n);
    for (const [lab, val] of [["Last action", inv.last_action_date ? fmtDate(inv.last_action_date) : null], ["Next follow-up", inv.next_followup ? fmtDate(inv.next_followup) : null]]) {
      if (!val) continue; const row = document.createElement("div"); row.className = "drow";
      const dk = document.createElement("span"); dk.className = "dk"; dk.textContent = lab; const dv = document.createElement("span"); dv.className = "dv"; dv.textContent = val;
      row.appendChild(dk); row.appendChild(dv); g.appendChild(row); }
    if (inv.notion_url) {
      const row = document.createElement("div"); row.className = "drow";
      const a = document.createElement("a"); a.className = "notion-link"; a.style.marginLeft = "0"; a.href = inv.notion_url; a.target = "_blank"; a.rel = "noopener";
      a.textContent = "Open in Notion"; a.title = "This invoice's page in the Invoice Tracker - the full note thread";
      const ed = document.createElement("span"); ed.className = "dv dim"; ed.textContent = inv.notion_edited ? "page edited " + fmtDate(inv.notion_edited, true) : "";
      row.appendChild(a); row.appendChild(ed); g.appendChild(row);
    }
    body.appendChild(g); }
  if (inv.notion_url) invNotionSection(body, inv.notion_url);   // the whole page - properties, body, comments
  const amt = num(inv.amount), bal = inv.balance == null ? null : num(inv.balance);
  const paid = (amt != null && bal != null) ? amt - bal : null;
  const isOpen = !(bal != null && bal <= 0.005) && (inv.status || "").toLowerCase() !== "paid";
  // Days-past-due is only meaningful while the invoice is still OPEN. A paid draw shows its
  // Paid date instead (a paid invoice isn't "past due").
  let dpd = isOpen ? inv.days_past_due : null;
  if (isOpen && dpd == null && inv.due_date) { const dd = Math.floor((Date.now() - Date.parse(inv.due_date)) / 86400000); if (!isNaN(dd)) dpd = dd; }
  const groups = [
    ["Billing", [
      ["Amount billed", money(amt), false],
      ["Open balance", money(bal), bal != null && bal > 0.005],
      paid != null ? ["Paid", money(paid), false] : null,
      ["Status", inv.status || (bal != null && bal > 0.005 ? "Open" : "Paid"), false],
    ]],
    ["Dates & terms", [
      ["Invoice date", inv.txn_date ? fmtDate(inv.txn_date) : null, false],
      ["Due date", inv.due_date ? fmtDate(inv.due_date) : null, false],
      dpd != null ? ["Days past due", dpd > 0 ? dpd + " days" : "current", dpd > 0] : null,
      ["Terms", inv.net_terms, false],
      ["Draw period", inv.draw_period, false],
      inv.paid_date ? ["Paid date", fmtDate(inv.paid_date), false] : null,
    ]],
    ["Lien", [
      ["Notice deadline", inv.lien_due_label, false],
      ["Lien status", inv.lien_status, false],
      ["Notice type", inv.lien_notice, false],
    ]],
  ];
  for (const [title, rows] of groups) {
    const present = rows.filter(r => r && r[1] != null && r[1] !== "" && r[1] !== "—");
    if (!present.length) continue;
    const g = document.createElement("div"); g.className = "dgroup";
    const h = document.createElement("h4"); h.textContent = title; g.appendChild(h);
    for (const [label, val, neg] of present) {
      const row = document.createElement("div"); row.className = "drow";
      const dk = document.createElement("span"); dk.className = "dk"; dk.textContent = label;
      const dv = document.createElement("span"); dv.className = "dv" + (neg ? " neg" : "");
      dv.textContent = val; dv.title = "Click to copy"; dv.onclick = () => copy(String(val));
      row.appendChild(dk); row.appendChild(dv); g.appendChild(row);
    }
    body.appendChild(g);
  }
  const qurl = qboInvoiceUrl(inv.qbo_txn_id);
  if (qurl) {
    const g = document.createElement("div"); g.className = "dgroup inv-qbo";
    const a = document.createElement("a"); a.href = qurl; a.target = "_blank"; a.rel = "noopener";
    a.className = "qbo-link"; a.textContent = "Open in QuickBooks ↗"; a.title = "Only if you need it";
    g.appendChild(a); body.appendChild(g);
  }
  openPanel("#invDetail");
}

function updateInvCollapseBtn() {
  const btn = $("#ifCollapse"); if (!btn) return;
  const allExp = invGroupKeys.length && invGroupKeys.every(k => invExpanded.has(k));
  btn.textContent = allExp ? "Collapse all" : "Expand all";
}
function invToggleAll() {
  const allExp = invGroupKeys.length && invGroupKeys.every(k => invExpanded.has(k));
  if (allExp) invExpanded.clear(); else invGroupKeys.forEach(k => invExpanded.add(k));
  renderOpenInvoices();
}
// The whole Notion page inside the drawer (owner 2026-09-02: "i need all the Notion page contents,
// all of it so i don't need to open notion"). Fetched on open via /api/invoice/notion (server-cached 60 s).
const _npCache = {};
function invNotionSection(body, url, bodyOnly) {
  // bodyOnly (the invoice PAGE, owner 2026-09-02: "just the page body and the timestamp"): no property
  // list - the page already shows those fields above; the drawer keeps the full property view.
  const g = document.createElement("div"); g.className = "dgroup np";
  if (!bodyOnly) { const h = document.createElement("h4"); h.textContent = "Notion page"; g.appendChild(h); }
  const box = document.createElement("div"); box.className = "np-box"; box.textContent = "Loading the Notion page…"; g.appendChild(box);
  body.appendChild(g);
  const draw = (d) => {
    box.innerHTML = "";
    if (!d || !d.ok) { box.textContent = "Could not load the page" + (d && d.error ? ": " + d.error : "") + "."; box.classList.add("dim"); return; }
    const meta = document.createElement("div"); meta.className = "np-meta"; meta.textContent = `as in Notion · page edited ${fmtDate(d.last_edited.replace(" ", "T"), true)}`; box.appendChild(meta);
    const props = bodyOnly ? [] : (d.properties || []).filter(p => p.value !== "" && p.value != null && p.type !== "title");
    if (props.length) {
      const pl = document.createElement("div"); pl.className = "np-props";
      for (const p of props) {
        const row = document.createElement("div"); row.className = "drow";
        const dk = document.createElement("span"); dk.className = "dk"; dk.textContent = p.name;
        const dv = document.createElement("span"); dv.className = "dv np-val"; dv.textContent = p.value; dv.title = "Click to copy"; dv.onclick = () => copy(String(p.value));
        if (p.type === "url" && /^https?:/.test(p.value)) { const a = document.createElement("a"); a.href = p.value; a.target = "_blank"; a.rel = "noopener"; a.textContent = p.value; dv.textContent = ""; dv.appendChild(a); }
        row.appendChild(dk); row.appendChild(dv); pl.appendChild(row);
      }
      box.appendChild(pl);
    }
    const blocks = d.blocks || [];
    if (blocks.length) {
      const bh = document.createElement("div"); bh.className = "np-sub"; bh.textContent = "Page body"; box.appendChild(bh);
      const bl = document.createElement("div"); bl.className = "np-body";
      for (const b of blocks) {
        const el = document.createElement("div"); el.className = "np-b np-" + b.type; el.style.marginLeft = (b.depth * 14) + "px";
        if (b.type === "divider") { el.className += " np-divider"; }
        else if (b.type === "to_do") { el.textContent = (b.checked ? "☑ " : "☐ ") + b.text; }
        else if (b.type === "bulleted_list_item") { el.textContent = "• " + b.text; }
        else if (b.type === "numbered_list_item") { el.textContent = "· " + b.text; }
        else if (b.url) { const a = document.createElement("a"); a.href = b.url; a.target = "_blank"; a.rel = "noopener"; a.textContent = b.text || b.type; el.appendChild(a); }
        else { el.textContent = b.text; }
        if (b.at && b.type === "paragraph" && b.text) el.title = "written " + b.at;
        bl.appendChild(el);
      }
      box.appendChild(bl);
    }
    const cm = d.comments || [];
    if (cm.length) {
      const ch = document.createElement("div"); ch.className = "np-sub"; ch.textContent = `Comments (${cm.length})`; box.appendChild(ch);
      for (const c of cm) {
        const el = document.createElement("div"); el.className = "np-comment";
        const who = document.createElement("div"); who.className = "np-who"; who.textContent = [c.by, c.at ? fmtDate(c.at.replace(" ", "T"), true) : ""].filter(Boolean).join(" · ");
        const tx = document.createElement("div"); tx.textContent = c.text; el.appendChild(who); el.appendChild(tx); box.appendChild(el);
      }
    }
    if (!props.length && !blocks.length && !cm.length) { box.textContent = bodyOnly ? "Nothing written on the page body yet, and no comments." : "The page has no content beyond the fields above."; box.classList.add("dim"); }
  };
  if (_npCache[url]) { draw(_npCache[url]); return; }
  fetch("/api/invoice/notion?url=" + encodeURIComponent(url)).then(r => r.json()).then(d => { if (d && d.ok) _npCache[url] = d; draw(d); }).catch(e => draw({ ok: false, error: String(e) }));
}

// ── Saved views for the Invoices tab (owner 2026-09-02: "give me ability to save custom views") ──
const INV_VIEWS_KEY = "proficient-ledger-invviews";
function _invViewsLoad() { try { return JSON.parse(localStorage.getItem(INV_VIEWS_KEY) || "{}") || {}; } catch { return {}; } }
function _invViewsSave(v) { try { localStorage.setItem(INV_VIEWS_KEY, JSON.stringify(v)); } catch { /* ignore */ } }
function invStateCapture() {
  const fv = sel => ($(sel) ? $(sel).value : "");
  return { months: invMonthSel === null ? null : [...invMonthSel],
           msel: Object.fromEntries(INV_MSEL.map(c => [c.id, [...(invMSel[c.id] || [])]])),
           div: fv("#ifDivision"), lienclk: fv("#ifLienClock"), lien: fv("#ifLien"), litig: fv("#ifLitig") || "ex",
           sort: fv("#ifSort") || "due", scope: invScope, view: invView, quick: invQuick,
           subgroup: invSubGroup, bucket: invBucketFilter };
}
function invStateApply(st) {
  if (!st) return;
  invMonthSel = st.months === null || st.months === undefined ? null : new Set(st.months);
  for (const c of INV_MSEL) invMSel[c.id] = new Set((st.msel || {})[c.id] || []);
  _invMSelSig = null;                                                  // rebuild the client / project menus
  const setv = (sel, v) => { const el = $(sel); if (el) el.value = v || ""; };
  setv("#ifDivision", st.div); setv("#ifLienClock", st.lienclk); setv("#ifLien", st.lien); setv("#ifLitig", st.litig || "ex"); setv("#ifSort", st.sort || "due");
  invQuick = st.quick || ""; { const q = $("#ifQuick"); if (q) q.value = invQuick; }
  invSubGroup = st.subgroup !== false; invBucketFilter = st.bucket == null ? null : st.bucket;
  const clickSeg = (segSel, attr, val) => { const b = document.querySelector(`${segSel} .seg-btn[data-${attr}="${val}"]`); if (b && !b.classList.contains("on")) b.click(); };
  clickSeg("#invViewSeg", "view", st.view || "amounts");
  clickSeg("#invScopeSeg", "scope", st.scope || "open");             // "all" fetches the paid ones on demand
  renderOpenInvoices();
}
function buildInvViews() {
  const vs = $("#ifViews"); if (!vs) return;
  const views = _invViewsLoad(), cur = vs.value;
  vs.innerHTML = ""; const o0 = document.createElement("option"); o0.value = ""; o0.textContent = "Custom"; vs.appendChild(o0);
  for (const name of Object.keys(views).sort((a, b) => a.localeCompare(b))) { const o = document.createElement("option"); o.value = name; o.textContent = name; vs.appendChild(o); }
  vs.value = cur; if (vs.value !== cur) vs.value = "";
  { const d = $("#ifDelView"); if (d) d.hidden = !vs.value; }
}
function invSaveView() {
  const vs = $("#ifViews");
  const name = (prompt("Name this view (same name overwrites):", vs && vs.value ? vs.value : "") || "").trim();
  if (!name) return;
  const views = _invViewsLoad(); views[name] = invStateCapture(); _invViewsSave(views);
  buildInvViews(); if (vs) { vs.value = name; } { const d = $("#ifDelView"); if (d) d.hidden = false; }
}
function invDeleteView() {
  const vs = $("#ifViews"); if (!vs || !vs.value) return;
  if (!confirm(`Delete the view "${vs.value}"?`)) return;
  const views = _invViewsLoad(); delete views[vs.value]; _invViewsSave(views);
  vs.value = ""; buildInvViews();
}
function invApplyView(name) {
  { const d = $("#ifDelView"); if (d) d.hidden = !name; }
  if (!name) return;
  invStateApply(_invViewsLoad()[name]);
}
function invClearFilters() {
  ["#ifDivision", "#ifLien", "#ifLienClock"].forEach(s => { const el = $(s); if (el) el.value = ""; });
  for (const cfg of INV_MSEL) invMSel[cfg.id] = new Set();   // clear Client + Project # multi-selects
  _invMSelSig = null;                                        // force the menus to rebuild (reset checks + label)
  const lt = $("#ifLitig"); if (lt) lt.value = "ex";         // baseline = litigation excluded
  invBucketFilter = null;
  invMonthSel = null; invQuick = ""; { const q = $("#ifQuick"); if (q) q.value = ""; }   // month + quick find too
  renderOpenInvoices();
}
// A project sub-band inside a client group (indented, lighter than the client band).
function invSubBand(proj, name, open, count, colspan) {
  const tr = document.createElement("tr"); tr.className = "bill-subgroup"; tr.style.cursor = "pointer"; tr.title = "Open the project page";
  tr.onclick = (e) => { e.stopPropagation(); if (proj && proj !== "(no project)") openProjectPage(proj); };
  const td = document.createElement("td"); td.colSpan = colspan;
  const cell = document.createElement("div"); cell.className = "bg-cell";
  const key = document.createElement("span"); key.className = "sg-key"; key.textContent = proj + (name ? " · " + name : "");
  cell.appendChild(key);
  bandMetrics(cell, [[money(open), "open", open > 0.005 ? "neg" : ""], [count, "invoices"]]);
  td.appendChild(cell); tr.appendChild(td);
  return tr;
}
function invSubGroupToggle() { invSubGroup = !invSubGroup; const b = $("#ifSubGroup"); if (b) b.textContent = invSubGroup ? "Flatten" : "Group by project"; renderOpenInvoices(); }

// ── Client statement: a clean, copy/paste-able table of the filtered open invoices ──
// A "different view" the owner opens, picks which invoices to include (all checked by
// default), and copies for a client - into Excel as cells or into an email as a table.
let stmtRows = [];                     // snapshot of the invoices shown when the panel opened
let stmtOn = new Set();                // keys the owner SELECTED (default NONE - opt in, owner 2026-08-21)
const invKey = i => String(i.qbo_txn_id || i.doc_number || `${i.project_no}|${i.txn_date}|${i.balance}`);
function _invRows() {                   // the same filtered set the table shows (msels + selects + bucket)
  const fv = sel => ($(sel) ? $(sel).value : "");
  const f = { div: fv("#ifDivision"), lien: fv("#ifLien"), litig: fv("#ifLitig") || "ex" };
  return (OI.invoices || []).filter(i => invPasses(i, f));
}
function _stmtChecked() { return stmtRows.filter(i => stmtOn.has(invKey(i))); }
function _stmtRowEls() { const b = $("#invStmtBody"); return b ? [...b.querySelectorAll("tr.stmt-row")] : []; }
function _stmtTotalUpdate() {
  const t = _stmtChecked().reduce((s, i) => s + oiBal(i), 0);
  { const el = $("#stmtTotalVal"); if (el) el.textContent = money(t); }
  { const el = $("#stmtCount"); if (el) el.textContent = `${stmtOn.size} of ${stmtRows.length} selected`; }
  { const b = $("#btnCopyStmt"); if (b) { b.textContent = stmtOn.size ? `Copy table (${stmtOn.size})` : "Copy table"; b.disabled = !stmtOn.size; } }
}
function _stmtSearch() {                 // narrow the list by project / invoice # / client
  const q = ($("#stmtSearch") ? $("#stmtSearch").value : "").trim().toLowerCase();
  for (const tr of _stmtRowEls()) tr.hidden = !!q && !(tr.dataset.s || "").includes(q);
}
function _stmtSelectVisible(on) {        // Select all / None over the rows the search currently shows
  for (const tr of _stmtRowEls()) {
    if (tr.hidden) continue;
    const cb = tr.querySelector("input");
    if (on) stmtOn.add(tr.dataset.key); else stmtOn.delete(tr.dataset.key);
    if (cb) cb.checked = on; tr.classList.toggle("on", on);
  }
  _stmtTotalUpdate();
}
function openInvStatement() {
  stmtRows = _invRows().slice().sort(INV_SORTS.due);
  stmtOn = new Set();
  if (!stmtRows.length) { toast("No invoices in the current filter to copy"); return; }
  const clients = [...new Set(stmtRows.map(i => i.customer || "(no client)"))];
  const multi = clients.length > 1;
  $("#invStmtTitle").textContent = multi ? `Open invoices · ${clients.length} clients` : clients[0];
  $("#invStmtSub").textContent = `${stmtRows.length} open invoice${stmtRows.length > 1 ? "s" : ""}${OI.as_of ? " · as of " + fmtDate(OI.as_of) : ""}`;
  const body = $("#invStmtBody"); body.innerHTML = "";
  // Controls: search + Select all / None (over the filtered rows) + a live count. Nothing is
  // selected to start (owner 2026-08-21: auto-selecting all was hard to work with).
  const ctrl = document.createElement("div"); ctrl.className = "stmt-ctrl";
  const search = document.createElement("input"); search.type = "search"; search.id = "stmtSearch"; search.className = "stmt-search";
  search.placeholder = "Search project / address / invoice #" + (multi ? " / client" : ""); search.oninput = _stmtSearch;
  const all = document.createElement("button"); all.type = "button"; all.className = "btn small"; all.textContent = "Select all"; all.title = "Select every invoice the search shows"; all.onclick = () => _stmtSelectVisible(true);
  const none = document.createElement("button"); none.type = "button"; none.className = "btn small"; none.textContent = "None"; none.onclick = () => _stmtSelectVisible(false);
  const cnt = document.createElement("span"); cnt.className = "stmt-count"; cnt.id = "stmtCount";
  ctrl.appendChild(search); ctrl.appendChild(all); ctrl.appendChild(none); ctrl.appendChild(cnt); body.appendChild(ctrl);
  body.appendChild(el2("p", "hint", "Search to narrow, then Select all or tick the ones to send. Internal columns (lien, litigation) are left off; Copy pastes into Excel as cells or into an email as a table."));
  const tbl = document.createElement("table"); tbl.className = "grid stmt-grid";
  const thead = document.createElement("thead"), tbody = document.createElement("tbody");
  const cols = multi ? ["", "Client", "Project", "Invoice #", "Invoice date", "Due date", "Past due", "Amount due"]
    : ["", "Project", "Invoice #", "Invoice date", "Due date", "Past due", "Amount due"];
  const amtIdx = cols.length - 1;
  const htr = document.createElement("tr");
  cols.forEach((c, idx) => { const th = document.createElement("th"); if (idx !== amtIdx) th.className = "left"; th.textContent = c; htr.appendChild(th); });
  thead.appendChild(htr);
  for (const i of stmtRows) {
    const p = i.project_no || "(no project)";
    const tr = document.createElement("tr"); tr.className = "stmt-row"; tr.dataset.key = invKey(i);
    tr.dataset.s = `${i.customer || ""} ${p} ${nameOf(p)} ${i.doc_number || ""}`.toLowerCase();
    const c0 = document.createElement("td"); const cb = document.createElement("input"); cb.type = "checkbox"; cb.checked = false;
    cb.onchange = () => { if (cb.checked) stmtOn.add(tr.dataset.key); else stmtOn.delete(tr.dataset.key); tr.classList.toggle("on", cb.checked); _stmtTotalUpdate(); };
    c0.appendChild(cb); tr.appendChild(c0);
    if (multi) tr.appendChild(leftText(i.customer || "–"));
    tr.appendChild(leftText(p + (nameOf(p) ? " · " + nameOf(p) : "")));
    tr.appendChild(leftText(i.doc_number || "–"));
    tr.appendChild(leftText(fmtDateShort(i.txn_date)));
    tr.appendChild(leftText(fmtDateShort(i.due_date)));
    const dpd = (i.days_past_due != null && i.days_past_due > 0) ? i.days_past_due + "d" : "–";
    const dc = leftText(dpd); if (i.days_past_due > 0) dc.style.color = "var(--neg)"; tr.appendChild(dc);
    tr.appendChild(rightText(money(oiBal(i))));
    tbody.appendChild(tr);
  }
  const ttr = document.createElement("tr"); ttr.className = "ag-total";
  const lead = document.createElement("td"); lead.className = "left"; lead.colSpan = amtIdx; lead.textContent = "Total selected"; ttr.appendChild(lead);
  const tv = document.createElement("td"); tv.className = "right"; tv.id = "stmtTotalVal"; ttr.appendChild(tv); tbody.appendChild(ttr);
  tbl.appendChild(thead); tbl.appendChild(tbody); body.appendChild(tbl);
  _stmtTotalUpdate();
  openPanel("#invStatement");
}
const _esc = s => String(s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
async function copyInvStatement() {
  const rows = _stmtChecked();
  if (!rows.length) { toast("Nothing checked to copy"); return; }
  const byClient = new Map();
  for (const i of rows) { const c = i.customer || "(no client)"; if (!byClient.has(c)) byClient.set(c, []); byClient.get(c).push(i); }
  const H = ["Client", "Project", "Invoice #", "Invoice date", "Due date", "Days past due", "Amount due"];
  const tsv = [H.join("\t")];
  let html = '<table border="1" cellspacing="0" cellpadding="5" style="border-collapse:collapse;font-family:Arial,Helvetica,sans-serif;font-size:13px">';
  html += "<thead><tr>" + H.map(h => `<th style="text-align:left;background:#f2f2f2">${_esc(h)}</th>`).join("") + "</tr></thead><tbody>";
  let grand = 0;
  for (const [c, list] of byClient) {
    list.sort((a, b) => (a.project_no || "").localeCompare(b.project_no || "", undefined, { numeric: true }) || String(a.txn_date || "").localeCompare(String(b.txn_date || "")));
    for (const i of list) {
      const proj = (i.project_no || "") + (nameOf(i.project_no) ? " " + nameOf(i.project_no) : "");
      const dpd = (i.days_past_due != null && i.days_past_due > 0) ? String(i.days_past_due) : "";
      const amt = Math.round(oiBal(i)); grand += oiBal(i);
      tsv.push([c, proj, i.doc_number || "", fmtDateShort(i.txn_date), fmtDateShort(i.due_date), dpd, amt].join("\t"));
      const htmlCells = [c, proj, i.doc_number || "", fmtDateShort(i.txn_date), fmtDateShort(i.due_date), dpd, "$" + amt.toLocaleString()];
      html += "<tr>" + htmlCells.map((x, idx) => `<td style="text-align:${idx === 6 ? "right" : "left"}">${_esc(String(x))}</td>`).join("") + "</tr>";
    }
  }
  tsv.push(["", "", "", "", "", "Total due", Math.round(grand)].join("\t"));
  html += `<tr><td colspan="6" style="text-align:right;font-weight:bold">Total due</td><td style="text-align:right;font-weight:bold">$${Math.round(grand).toLocaleString()}</td></tr></tbody></table>`;
  try {
    if (navigator.clipboard && window.ClipboardItem) {
      await navigator.clipboard.write([new ClipboardItem({
        "text/html": new Blob([html], { type: "text/html" }),
        "text/plain": new Blob([tsv.join("\n")], { type: "text/plain" }) })]);
      toast(`Copied ${rows.length} invoice${rows.length > 1 ? "s" : ""} - paste into Excel or an email`);
    } else { throw new Error("no ClipboardItem"); }
  } catch { copy(tsv.join("\n")); }
}

// ══ CUSTOMER CENTER ══════════════════════════════════════════════════════════
// Top clients by OPEN AR (what they still owe), grouped PER DIVISION - so you see
// who the big clients are in Commercial vs Residential vs Multi Family, not just
// one "biggest" overall. Click a client to jump to Invoices filtered to them.
const CUST_DIV_ORDER = ["Commercial", "Residential", "Multi Family"];
function renderCustomers() {
  const invs = OI.invoices || [];
  const byDiv = new Map();               // division -> Map(client -> {open, n, oldest})
  for (const i of invs) {
    const div = i.division || "(no division)";
    const c = i.customer || "(no client)";
    if (!byDiv.has(div)) byDiv.set(div, new Map());
    const m = byDiv.get(div);
    const e = m.get(c) || { client: c, open: 0, n: 0, oldest: null };
    e.open += oiBal(i); e.n += 1;
    if (i.due_date && (!e.oldest || i.due_date < e.oldest)) e.oldest = i.due_date;
    m.set(c, e);
  }
  // stable division order: the three known ones first, then any extras by open $ desc
  const order = [...byDiv.keys()].sort((a, b) => {
    const ia = CUST_DIV_ORDER.indexOf(a), ib = CUST_DIV_ORDER.indexOf(b);
    if (ia !== -1 || ib !== -1) return (ia === -1 ? 99 : ia) - (ib === -1 ? 99 : ib);
    return a.localeCompare(b);
  });
  const divOpen = d => [...byDiv.get(d).values()].reduce((t, e) => t + e.open, 0);
  const clients = new Set(); let total = 0;
  for (const m of byDiv.values()) for (const e of m.values()) { clients.add(e.client); total += e.open; }
  { const n = $("#custNote"); if (n) n.textContent = clients.size ? `(${clients.size} clients · ${money(total)} open)` : "(no AR data - load invoices)"; }
  // ── Payment speed → future cash-in forecast (owner 2026-08-25) ──
  const paySpeed = OI.pay_speed || { by_client: {}, all_avg: null };
  const clientAvgDays = c => { const s = paySpeed.by_client[(c || "").toLowerCase()]; return (s && s.avg_days != null) ? s.avg_days : paySpeed.all_avg; };
  const _DAYMS = 86400000, _now = Date.now();
  const expectedMs = i => { const m = String(i.txn_date || "").match(/^(\d{4})-(\d{2})-(\d{2})/); const a = clientAvgDays(i.customer);
    if (!m || a == null) return null; const d = new Date(+m[1], +m[2] - 1, +m[3]); d.setDate(d.getDate() + a); return d.getTime(); };
  let f30 = 0, f60 = 0, f90 = 0;   // cumulative expected receipts within N days (by each client's own pay pattern)
  for (const i of invs) { const ms = expectedMs(i); if (ms == null) continue; const out = (ms - _now) / _DAYMS; const bal = oiBal(i);
    if (out <= 30) { f30 += bal; f60 += bal; f90 += bal; } else if (out <= 60) { f60 += bal; f90 += bal; } else if (out <= 90) { f90 += bal; } }
  { const stats = $("#custStats"); if (stats) { stats.innerHTML = "";
      const tiles = [["Open AR", money(total)], ["Clients", String(clients.size)]];
      for (const d of order) tiles.push([d, money(divOpen(d))]);       // per-division open AR (replaces the useless "biggest client")
      if (paySpeed.all_avg != null) tiles.push(["Cash-in ≤30d", money(f30), "fc"], ["≤60d", money(f60), "fc"], ["≤90d", money(f90), "fc"]);
      for (const [l, v, cls] of tiles) {
        const k = el2("div", "kpi" + (cls ? " kpi-" + cls : "")); k.appendChild(el2("div", "k-label", l)); k.appendChild(el2("div", "k-value", v)); stats.appendChild(k); } } }
  const NCOL = 5;
  const tb = buildHead("#custTable", [["Client", "left"], ["Open AR", "right"], ["Open invoices", "right"], ["Oldest due", "left"], ["Avg days to pay", "right"]]);
  if (!tb) return; tb.innerHTML = "";
  if (!order.length) { const tr = document.createElement("tr"); const td = document.createElement("td"); td.colSpan = NCOL; td.className = "left"; td.style.color = "var(--text-dim)"; td.style.padding = "14px 12px"; td.textContent = "No open AR - run load_invoices.py."; tr.appendChild(td); tb.appendChild(tr); return; }
  for (const div of order) {
    const rows = [...byDiv.get(div).values()].sort((a, b) => b.open - a.open);   // top clients first
    // division band header (spans the row) - open AR + client count in this division
    const gtr = document.createElement("tr"); gtr.className = "bill-group";
    const gtd = document.createElement("td"); gtd.colSpan = NCOL;
    const cell = document.createElement("div"); cell.className = "bg-cell";
    const key = document.createElement("span"); key.className = "bg-key"; key.textContent = div;
    cell.appendChild(key);
    bandMetrics(cell, [[money(divOpen(div)), "open", "neg"], [rows.length, "clients"]]);
    gtd.appendChild(cell); gtr.appendChild(gtd); tb.appendChild(gtr);
    for (const r of rows) {
      const tr = document.createElement("tr"); tr.style.cursor = "pointer"; tr.title = "See this client's open invoices";
      tr.onclick = () => { invMSel.ifClient = new Set([r.client]); invMSel.ifProj = new Set(); _invMSelSig = null;   // filter Invoices to this client (msel)
        const df = $("#ifDivision"); if (df) df.value = ""; setTab("invoices"); renderOpenInvoices(); };
      tr.appendChild(leftText(r.client)); tr.appendChild(rightText(money(r.open)));
      tr.appendChild(rightText(String(r.n))); tr.appendChild(leftText(r.oldest ? fmtDateShort(r.oldest) : "–"));
      // Avg days to pay (from this client's paid history); dim the portfolio fallback so it reads as an estimate.
      const sp = paySpeed.by_client[(r.client || "").toLowerCase()];
      const ad = document.createElement("td"); ad.className = "right";
      if (sp && sp.avg_days != null) { ad.textContent = `${sp.avg_days}d`; ad.title = `avg over ${sp.n} paid invoice${sp.n === 1 ? "" : "s"}`; }
      else if (paySpeed.all_avg != null) { ad.textContent = `~${paySpeed.all_avg}d`; ad.className = "right dim"; ad.title = "no paid history for this client - portfolio average"; }
      else ad.textContent = "–";
      tr.appendChild(ad);
      tb.appendChild(tr);
    }
  }
}

// ══ PAYMENTS ═════════════════════════════════════════════════════════════════
// Each row is ONE payment transaction (money IN): Client · Payment Ref # · Payment Type ·
// Amount Paid. Expand it to see the invoices it paid (invoice # · total open · amount applied).
// "Unlocks (AP)" ties the money-in to money-out: the open vendor bills matched to the DRAW(S)
// this payment paid (bill.invoice_no) - NOT every bill on the project - click it and the side
// panel lists them (talks to the Bills tab data).
// Sourced from QBO Payment objects (load_payments.py → payment / payment_application).
// Open vendor bills indexed by the DRAW (AR invoice) they're matched to - NOT by project.
// A payment pays a specific draw; only the bills tied to that draw are what it unlocks, not the
// whole project's AP backlog (owner 2026-08-25: "use draw period, you are grabbing all costs").
function payOpenBillsByDraw() {   // open bills keyed by the DRAW (AR invoice) they're matched to
  const idx = {};
  for (const b of (BILLS || [])) {
    const draw = b.invoice_no || b.matched_invoice;
    if (num(b.open_balance) > 0.005 && draw) (idx[String(draw)] ||= []).push(b);
  }
  return idx;
}
function payOpenBillsByProject() {   // open bills keyed by project (for RP - see the division rule below)
  const idx = {};
  for (const b of (BILLS || [])) { if (num(b.open_balance) > 0.005 && b.project_no) (idx[b.project_no] ||= []).push(b); }
  return idx;
}
// Division rule (owner 2026-08-25). CP/MFD are STAGED: each draw is its own scope with its own costs
// and its own invoice, so a payment unlocks ONLY the bills matched to the draw it paid. RP is regular
// work: costs go in UP FRONT and the job is invoiced ONCE at the end, so bills aren't tied to a draw -
// use the whole project's open AP for RP.
const _payIsRP = (proj, div) => /^RP/i.test(proj || "") || String(div || "").toLowerCase().startsWith("res");
function payUnlockBills(p, drawIdx, projIdx) {
  const bills = new Set();
  for (const a of (p.applications || [])) {
    if (_payIsRP(a.project_no, a.division)) { for (const b of (projIdx[a.project_no] || [])) bills.add(b); }
    else if (a.invoice_no) { for (const b of (drawIdx[String(a.invoice_no)] || [])) bills.add(b); }
  }
  return [...bills];
}
function renderPayments() {
  const body = $("#payBody"); if (!body) return;
  const pays = PAY.payments || [];
  { const n = $("#payNote"); if (n) n.textContent = pays.length ? `(${pays.length} payments · ${money(PAY.total_received)} received)` : "(no payment data - run load_payments.py)"; }
  body.innerHTML = "";
  const drawIdx = payOpenBillsByDraw(), projIdx = payOpenBillsByProject();
  const stats = document.createElement("div"); stats.className = "kpi-row";
  for (const [l, v] of [["Received", money(PAY.total_received)], ["Payments", String(pays.length)],
                        ["Invoices paid", String(PAY.invoices_paid || 0)]]) {
    const k = el2("div", "kpi"); k.appendChild(el2("div", "k-label", l)); k.appendChild(el2("div", "k-value", v)); stats.appendChild(k);
  }
  body.appendChild(stats);
  const head = document.createElement("div"); head.className = "list-head";
  const hint = document.createElement("p"); hint.className = "hint"; hint.style.margin = "0";
  hint.innerHTML = "Each row is a <b>payment received</b>. Click it to see the invoices (draws) it paid. <b>Unlocks (AP)</b> is the open vendor bills this payment funds: for staged <b>CP/MFD</b> draws, only the bills on the draw it paid; for <b>RP</b> (costs up front, billed once), the whole job's open AP. <b>Net after AP</b> = amount paid − that AP: what's left once those vendors are paid (red = the AP exceeds the payment).";
  head.appendChild(hint);
  const actions = document.createElement("div"); actions.className = "list-actions";
  const seg = document.createElement("div"); seg.className = "seg"; seg.title = "Break cash-in down by period";
  for (const [val, lbl] of [["none", "Flat"], ["week", "Weeks"], ["month", "Months"]]) {
    const b = document.createElement("button"); b.type = "button"; b.className = "seg-btn" + (paymentsGroupBy === val ? " on" : ""); b.textContent = lbl;
    b.onclick = () => { paymentsGroupBy = val; renderPayments(); };
    seg.appendChild(b);
  }
  actions.appendChild(seg);
  if (pays.length) {
    const btn = document.createElement("button"); btn.className = "btn small subtle";
    if (paymentsGroupBy !== "none") {   // grouped: the button expands/collapses the month/week BANDS (top level)
      const keys = [...new Set(pays.map(p => payPeriod(p.txn_date, paymentsGroupBy).key))];
      const allExp = keys.length && keys.every(k => paymentsPeriodsExpanded.has(k));
      btn.textContent = allExp ? "Collapse all" : "Expand all";
      btn.onclick = () => { if (allExp) paymentsPeriodsExpanded.clear(); else keys.forEach(k => paymentsPeriodsExpanded.add(k)); renderPayments(); };
    } else {                            // flat: the button expands/collapses each payment's invoices
      const allExpanded = pays.every(p => paymentsExpanded.has(p.qbo_txn_id));
      btn.textContent = allExpanded ? "Collapse all" : "Expand all";
      btn.onclick = () => { if (allExpanded) paymentsExpanded.clear(); else pays.forEach(p => paymentsExpanded.add(p.qbo_txn_id)); renderPayments(); };
    }
    actions.appendChild(btn);
  }
  head.appendChild(actions);
  body.appendChild(head);
  const wrap = document.createElement("div"); wrap.className = "table-scroll";
  const table = document.createElement("table"); table.className = "grid"; table.id = "payTable";
  table.innerHTML = "<thead></thead><tbody></tbody>"; wrap.appendChild(table); body.appendChild(wrap);
  const cols = [["Client", "left"], ["Project", "left"], ["Date", "left"], ["Payment Ref #", "left"], ["Payment Type", "left"], ["Amount Paid", "right"], ["Unlocks (AP)", "right"], ["Net after AP", "right"]];
  const tb = buildHead("#payTable", cols);
  if (!tb) return; tb.innerHTML = "";
  if (!pays.length) {
    const tr = document.createElement("tr"); const td = document.createElement("td"); td.colSpan = cols.length;
    td.className = "left"; td.style.color = "var(--text-dim)"; td.style.padding = "14px 12px";
    td.textContent = "No payments loaded - run python3 ledger/load_payments.py (pulls QBO Payment transactions).";
    tr.appendChild(td); tb.appendChild(tr); return;
  }
  // Newest first; optionally banded by week/month with a per-period cash-in total (owner 2026-08-25).
  const sorted = [...pays].sort((a, b) => String(b.txn_date || "").localeCompare(String(a.txn_date || "")));
  const perTot = {}, perN = {};
  if (paymentsGroupBy !== "none") for (const p of sorted) { const k = payPeriod(p.txn_date, paymentsGroupBy).key; perTot[k] = (perTot[k] || 0) + num(p.total_amt); perN[k] = (perN[k] || 0) + 1; }
  let curPeriod = null;
  for (const p of sorted) {
    if (paymentsGroupBy !== "none") {
      const per = payPeriod(p.txn_date, paymentsGroupBy);
      if (per.key !== curPeriod) {
        curPeriod = per.key;
        const pExp = paymentsPeriodsExpanded.has(per.key);   // bands COLLAPSED by default (owner 2026-08-31)
        const gtr = document.createElement("tr"); gtr.className = "bill-group"; gtr.style.cursor = "pointer";
        gtr.title = pExp ? "Click to collapse" : "Click to expand";
        const gtd = document.createElement("td"); gtd.colSpan = cols.length;
        const cell = document.createElement("div"); cell.className = "bg-cell";
        const left = document.createElement("span"); left.className = "bg-left";
        const caret = document.createElement("span"); caret.className = "bg-caret"; caret.textContent = pExp ? "▾" : "▸";
        const key = document.createElement("span"); key.className = "bg-key"; key.textContent = per.label;
        left.appendChild(caret); left.appendChild(key);
        cell.appendChild(left);
        bandMetrics(cell, [[money(perTot[per.key]), "received", "pos"], [perN[per.key], "payments"]]);
        gtd.appendChild(cell); gtr.appendChild(gtd); tb.appendChild(gtr);
        gtr.onclick = () => { if (paymentsPeriodsExpanded.has(per.key)) paymentsPeriodsExpanded.delete(per.key); else paymentsPeriodsExpanded.add(per.key); renderPayments(); };
      }
      if (!paymentsPeriodsExpanded.has(per.key)) continue;   // collapsed band → skip its payment rows
    }
    const expanded = paymentsExpanded.has(p.qbo_txn_id);
    // ── the payment transaction: Client · Ref # · Type · Amount Paid · Unlocks (AP) ──
    const tr = document.createElement("tr"); tr.className = "pay-row"; tr.style.cursor = "pointer";
    tr.title = expanded ? "Click to hide the invoices" : "Click to see the invoices this payment paid";
    // client (with caret + GC link)
    const cc = document.createElement("td"); cc.className = "left";
    const caret = document.createElement("span"); caret.className = "bg-caret"; caret.textContent = expanded ? "▾" : "▸"; cc.appendChild(caret);
    const payer = p.parent_customer || p.customer || "–";
    const curl = qboCustomerUrl(p.parent_customer_id || p.customer_id);
    if (curl) { const a = document.createElement("a"); a.href = curl; a.target = "_blank"; a.rel = "noopener"; a.className = "qbo-link"; a.textContent = payer; a.title = "Open this customer in QuickBooks"; a.onclick = (e) => e.stopPropagation(); cc.appendChild(a); }
    else cc.appendChild(document.createTextNode(payer));
    tr.appendChild(cc);
    // the project # this payment pays (owner 2026-09-02): one project, or "multiple" - click the row for the lines
    { const projs = [...new Set((p.applications || []).map(a => a.project_no).filter(Boolean))];
      const pc = document.createElement("td"); pc.className = "left";
      if (projs.length > 1) { const s = document.createElement("span"); s.className = "vp-multi"; s.textContent = "multiple"; s.title = projs.join(", ") + " - click the row to see each invoice"; pc.appendChild(s); }
      else { pc.textContent = projs[0] || "–"; if (!projs.length) pc.classList.add("dim"); }
      tr.appendChild(pc); }
    tr.appendChild(leftText(fmtDateShort(p.txn_date)));
    tr.appendChild(leftText(p.ref_no || "–"));
    tr.appendChild(leftText(p.method || "–"));
    tr.appendChild(rightText(money(p.total_amt)));
    // Unlocks (AP): open vendor bills on this payment's project(s) → click opens the side panel
    const uc = document.createElement("td"); uc.className = "right";
    const bills = payUnlockBills(p, drawIdx, projIdx);
    const apSum = bills.reduce((t, b) => t + num(b.open_balance), 0);
    if (bills.length) {
      const link = document.createElement("span"); link.className = "unlock-link";
      link.textContent = `${money(apSum)} · ${bills.length}`;
      link.title = "Open vendor bills on the draw(s) this payment paid - the AP it funds";
      link.onclick = (e) => { e.stopPropagation(); openPaymentBills(p, bills); };
      uc.appendChild(link);
    } else uc.appendChild(document.createTextNode("–"));
    tr.appendChild(uc);
    // Net after AP: money in − the AP this payment funds = what's left once those vendors are paid.
    // Negative = the AP owed on this draw/job exceeds what came in (this payment doesn't cover it).
    const net = num(p.total_amt) - apSum;
    const nc = document.createElement("td"); nc.className = "right"; nc.style.fontWeight = "600";
    nc.textContent = money(net);
    nc.title = bills.length
      ? `${money(p.total_amt)} in − ${money(apSum)} AP = ${money(net)} left after paying those vendors`
      : "No AP tied to this payment - the full amount is net";
    if (net < -0.005) nc.style.color = "var(--neg)";
    tr.appendChild(nc);
    tr.onclick = () => { if (paymentsExpanded.has(p.qbo_txn_id)) paymentsExpanded.delete(p.qbo_txn_id); else paymentsExpanded.add(p.qbo_txn_id); renderPayments(); };
    tb.appendChild(tr);
    // ── grouped invoices this payment paid: Invoice # · Total open · Amount applied ──
    if (expanded) {
      const sr = document.createElement("tr"); sr.className = "pay-invoices";
      const std = document.createElement("td"); std.colSpan = cols.length;
      const box = document.createElement("table"); box.className = "sub-grid";
      const th = document.createElement("thead"); th.innerHTML = "<tr><th class='left'>Invoice #</th><th class='left'>Project</th><th class='left'>Invoice date</th><th class='left'>Memo</th><th class='right'>Total open</th><th class='right'>Amount applied</th></tr>";
      box.appendChild(th);
      const bod = document.createElement("tbody");
      if (!p.applications.length) {
        const r = document.createElement("tr"); const c = document.createElement("td"); c.colSpan = 6; c.className = "left dim";
        c.textContent = (p.unapplied_amt || 0) > 0.005 ? "Unapplied - a credit on account, not yet on an invoice." : "No invoice links on this payment.";
        r.appendChild(c); bod.appendChild(r);
      }
      for (const a of p.applications) {
        const r = document.createElement("tr");
        r.appendChild(qboLinkCell(a.invoice_no || ("inv " + a.invoice_txn_id), a.invoice_no ? qboInvoiceUrl(a.invoice_txn_id) : null, "Open this invoice in QuickBooks"));
        r.appendChild(leftText(a.project_no || "–"));
        r.appendChild(leftText(a.invoice_date ? fmtDateShort(a.invoice_date) : "–"));
        { const mc = leftText(a.memo || "–"); mc.className += " inv-memo"; mc.title = a.memo || ""; r.appendChild(mc); }
        const oc = document.createElement("td"); oc.className = "right";
        if (a.invoice_open == null) oc.appendChild(document.createTextNode("–"));
        else if (a.invoice_open > 0.005) { oc.textContent = money(a.invoice_open); oc.style.color = "var(--neg)"; }
        else { oc.textContent = "paid"; oc.className = "right dim"; }
        r.appendChild(oc);
        r.appendChild(rightText(money(a.amount)));
        bod.appendChild(r);
      }
      box.appendChild(bod); std.appendChild(box); sr.appendChild(std); tb.appendChild(sr);
    }
  }
}

// The AP bills a payment unlocks: open vendor bills on the same project(s), grouped by project.
// Read from the already-loaded Bills tab data - money IN (this payment) → money OUT (these bills).
function openPaymentBills(p, bills) {
  $("#payBillsTitle").textContent = (p.parent_customer || p.customer || "Payment");
  const projs = [...new Set((p.applications || []).map(a => a.project_no).filter(Boolean))];
  const sum = bills.reduce((t, b) => t + num(b.open_balance), 0);
  $("#payBillsSub").textContent = `${fmtDateShort(p.txn_date)} · ${money(p.total_amt)} in · unlocks ${money(sum)} AP across ${projs.length} job${projs.length === 1 ? "" : "s"}`;
  const body = $("#payBillsBody"); body.innerHTML = "";
  const intro = document.createElement("p"); intro.className = "hint";
  intro.textContent = "Open vendor bills tied to the draw(s) this payment paid - the AP this cash-in actually funds (not the whole job).";
  body.appendChild(intro);
  const byProj = new Map();
  for (const b of bills) { const k = b.project_no || "–"; if (!byProj.has(k)) byProj.set(k, []); byProj.get(k).push(b); }
  for (const [proj, list] of byProj) {
    list.sort((a, b) => num(b.open_balance) - num(a.open_balance));
    const g = document.createElement("div"); g.className = "dgroup";
    const h = document.createElement("h4");
    h.textContent = `${proj}${nameOf(proj) ? " · " + nameOf(proj) : ""} · ${money(list.reduce((t, b) => t + num(b.open_balance), 0))} open`;
    g.appendChild(h);
    const t = document.createElement("table"); t.className = "sub-grid";
    t.innerHTML = "<thead><tr><th class='left'>Vendor</th><th class='left'>Bill #</th><th class='right'>Open</th><th class='left'>Status</th></tr></thead>";
    const tbb = document.createElement("tbody");
    for (const b of list) {
      const r = document.createElement("tr");
      const vend = qboBillHref(b.qbo_link);
      if (vend) { const vtd = document.createElement("td"); vtd.className = "left"; const a = document.createElement("a"); a.href = vend; a.target = "_blank"; a.rel = "noopener"; a.className = "qbo-link"; a.textContent = b.vendor || "–"; a.title = "Open bill in QuickBooks"; vtd.appendChild(a); r.appendChild(vtd); }
      else r.appendChild(leftText(b.vendor || "–"));
      r.appendChild(leftText(b.bill_ref || "–"));
      const oc = document.createElement("td"); oc.className = "right"; oc.textContent = money(b.open_balance); oc.style.color = "var(--neg)"; r.appendChild(oc);
      r.appendChild(leftText(b.pay_status || "–"));
      tbb.appendChild(r);
    }
    t.appendChild(tbb); g.appendChild(t); body.appendChild(g);
  }
  openPanel("#payBills");
}

function renderSubLoc() {
  const s = SUBLOC.summary;
  const note = $("#sublocNote");
  const clearAll = () => { if ($("#sublocStats")) $("#sublocStats").innerHTML = ""; if ($("#sublocHint")) $("#sublocHint").textContent = "";
    for (const t of ["#sublocProjTable", "#sublocDivTable"]) { const el = $(t); if (el) { el.querySelector("thead").innerHTML = ""; el.querySelector("tbody").innerHTML = ""; } }
    if ($("#sublocFeed")) $("#sublocFeed").innerHTML = ""; };
  if (!s) { if (note) note.textContent = "(not loaded - run the Sub LOC pipeline in Console, or python3 ledger/load_sub_loc.py)"; clearAll(); return; }
  if (note) note.textContent = `(window ${fmtDateShort(s.window_start)}–${fmtDateShort(s.window_end)} · loaded ${s.loaded_at ? fmtDate(s.loaded_at, true) : "–"})`;
  const stats = [
    ["Fronted, still out", money(s.outstanding), "sub $ paid, not yet collected"],
    ["Peak LOC needed", money(s.peak), s.peak_date ? "high-water " + fmtDate(s.peak_date) : "high-water"],
    ["Avg draw→repay", (s.avg_lag != null ? Math.round(s.avg_lag) : "–") + " days", "days our cash is out"],
    ["Prefunded", money(s.prefunded), "GC paid before we paid the sub"],
  ];
  const sr = $("#sublocStats"); sr.innerHTML = "";
  for (const [label, value, sub] of stats) {
    const el = document.createElement("div"); el.className = "kpi";
    el.innerHTML = `<div class="k-label"></div><div class="k-value"></div><div class="k-sub"></div>`;
    el.querySelector(".k-label").textContent = label; el.querySelector(".k-value").textContent = value; el.querySelector(".k-sub").textContent = sub;
    sr.appendChild(el);
  }
  $("#sublocHint").innerHTML = "Cash you <b>front to subs before the GC pays you</b> for that work. " +
    "<b>Fronted, still out</b> is today's float; <b>Peak</b> is the high-water mark - <b>size your line of credit to it</b> " +
    "(rule of thumb is a LOC around 10–20% of revenue, but your real need is the peak). Matched per project + draw period, " +
    "FIFO, so a client payment pays off the oldest fronted subs first. Read-only from QBO via <code>load_sub_loc.py</code>.";
  applySublocSections();

  // By project: TOP few most in the hole (outstanding desc); expand for the rest so the tab
  // stays a dashboard, not a long list. Click a row for its open subs grouped by draw.
  { const tb = _slBuildG("#sublocProjTable", [["Project", "left"], ["Fronted (out)", "right"], ["Drawn", "right"], ["Repaid", "right"], ["Avg days", "right"]]);
    const obp = SUBLOC.open_by_project || {};
    const all = [...(SUBLOC.projects || [])].sort((a, b) => (b.outstanding || 0) - (a.outstanding || 0));   // most in the hole first
    const shown = sublocProjExpanded ? all : all.slice(0, SUBLOC_PROJ_TOP);
    for (const p of shown) { const tr = document.createElement("tr");
      const drillable = !!obp[p.project];
      if (drillable) { tr.style.cursor = "pointer"; tr.title = "See this project's open subs by draw";
        tr.onclick = (e) => { if (e.target.closest(".cell")) return; openSublocDetail(p.project); }; }
      // project # is the visible click affordance (accent link + a › on hover) when it drills in
      const pc = document.createElement("td"); pc.className = "left";
      if (drillable) { const lk = document.createElement("span"); lk.className = "row-open"; lk.textContent = p.project; pc.appendChild(lk); }
      else pc.appendChild(document.createTextNode(p.project || "–"));
      tr.appendChild(pc);
      const o = _slMcell(p.outstanding); if ((p.outstanding || 0) > 0.005) o.querySelector(".cell").classList.add("open-amt"); tr.appendChild(o);
      tr.appendChild(_slMcell(p.drawn)); tr.appendChild(_slMcell(p.repaid));
      tr.appendChild(rightText(p.avg_lag != null ? Math.round(p.avg_lag) + "d" : "–")); tb.appendChild(tr); }
    if (!all.length) _slEmpty(tb, 5, "Nothing fronted in this window.");
    const more = $("#sublocProjMore"); if (more) { more.innerHTML = "";
      if (all.length > SUBLOC_PROJ_TOP) {
        const btn = document.createElement("button"); btn.className = "btn small subtle";
        btn.textContent = sublocProjExpanded ? `Show top ${SUBLOC_PROJ_TOP}` : `Expand more (${all.length - SUBLOC_PROJ_TOP})`;
        btn.onclick = () => { sublocProjExpanded = !sublocProjExpanded; renderSubLoc(); };
        more.appendChild(btn);
      }
    }
  }
  // By division (flat, at the top of the tab)
  { const tb = _slBuildG("#sublocDivTable", [["Division", "left"], ["Fronted (out)", "right"], ["Peak", "right"], ["Drawn", "right"], ["Repaid", "right"], ["Avg days", "right"]]);
    const divs = SUBLOC.divisions || {}; const order = ["MFD", "CP", "RP", "Other"]; const rank = k => { const i = order.indexOf(k); return i < 0 ? 99 : i; };
    const keys = Object.keys(divs).sort((a, b) => rank(a) - rank(b));
    for (const k of keys) { const d = divs[k]; const tr = document.createElement("tr");
      tr.appendChild(leftText(k)); const o = _slMcell(d.outstanding); if ((d.outstanding || 0) > 0.005) o.querySelector(".cell").classList.add("open-amt"); tr.appendChild(o);
      tr.appendChild(_slMcell(d.peak)); tr.appendChild(_slMcell(d.drawn)); tr.appendChild(_slMcell(d.repaid));
      tr.appendChild(rightText(d.avg_lag != null ? Math.round(d.avg_lag) + "d" : "–")); tb.appendChild(tr); }
    if (!keys.length) _slEmpty(tb, 6, "No sub float in this window.");
  }
  renderSublocFeed();
}
// The per-project LOC event chain (the "where this came from" source report) is fetched ON DEMAND
// and cached - it never rides in the bulk load, so the tab stays light no matter how many exist.
const _sublocSrcCache = {};
async function _sublocLoadSource(project, body) {
  const sec = document.createElement("div"); sec.className = "dgroup";
  const h = document.createElement("h4"); h.textContent = "Where this came from - every LOC transaction"; sec.appendChild(h);
  const cap = document.createElement("div"); cap.className = "bills-cap"; cap.textContent = "loading the source…"; sec.appendChild(cap);
  body.appendChild(sec);
  let data = _sublocSrcCache[project];
  if (!data) {
    try { data = await (await fetch("/api/subloc/project?p=" + encodeURIComponent(project))).json(); _sublocSrcCache[project] = data; }
    catch (e) { cap.textContent = "could not load the source events"; return; }
  }
  const evs = (data && data.events) || [];
  if (!evs.length) { cap.textContent = "No LOC events recorded for this project."; return; }
  const outT = evs.reduce((t, e) => t + (e.out_amt || 0), 0), inT = evs.reduce((t, e) => t + (e.in_amt || 0), 0);
  cap.textContent = `${evs.length} transactions · ${money(outT)} fronted out · ${money(inT)} reimbursed in`;
  const scroll = document.createElement("div"); scroll.className = "table-scroll";
  const table = document.createElement("table"); table.className = "grid"; const thead = document.createElement("thead"), tbody = document.createElement("tbody");
  const htr = document.createElement("tr");
  for (const [c, al] of [["Date", "left"], ["Event", "left"], ["Party / invoice", "left"], ["Out", "right"], ["In", "right"], ["LOC balance", "right"], ["Lag", "right"]]) { const th = document.createElement("th"); if (al === "left") th.className = "left"; th.textContent = c; htr.appendChild(th); }
  thead.appendChild(htr);
  for (const e of evs) {
    const tr = document.createElement("tr");
    tr.appendChild(leftText(fmtDateShort(e.event_date)));
    tr.appendChild(leftText(e.type || "–"));
    tr.appendChild(leftText(((e.invoice ? "INV " + e.invoice : "") + (e.party ? (e.invoice ? " · " : "") + e.party : "")) || "–"));
    tr.appendChild(rightText(e.out_amt ? money(e.out_amt) : "–"));
    tr.appendChild(rightText(e.in_amt ? money(e.in_amt) : "–"));
    tr.appendChild(rightText(e.balance != null ? money(e.balance) : "–"));
    tr.appendChild(rightText(e.lag_days != null ? e.lag_days + "d" : "–"));
    tbody.appendChild(tr);
  }
  table.appendChild(thead); table.appendChild(tbody); scroll.appendChild(table); sec.appendChild(scroll);
}
function renderSublocFeed() {
  const box = $("#sublocFeed"); if (!box) return; box.innerHTML = "";
  const repays = SUBLOC.repays || (SUBLOC.events || []).filter(e => e.type === "REPAY" && (e.in_amt || 0) > 0.005);
  if (!repays.length) { const p = document.createElement("p"); p.className = "hint"; p.style.margin = "10px 18px"; p.textContent = "No client repayments matched to fronted subs yet."; box.appendChild(p); return; }
  const now = new Date(); const dow = (now.getDay() + 6) % 7;   // Monday = 0
  const weekStart = new Date(now.getFullYear(), now.getMonth(), now.getDate() - dow);
  const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);
  const parse = d => { const m = String(d || "").match(/^(\d{4})-(\d{2})-(\d{2})/); return m ? new Date(+m[1], +m[2] - 1, +m[3]) : null; };
  const b = { week: [], month: [], prior: [] };
  for (const e of repays) { const d = parse(e.event_date);
    if (d && d >= weekStart) b.week.push(e); else if (d && d >= monthStart) b.month.push(e); else b.prior.push(e); }
  for (const [key, label, rows, collapsible] of [["feed-week", "This week", b.week, false],
       ["feed-month", "This month", b.month, false], ["feed-prior", "Prior months", b.prior, true]]) {
    if (!rows.length && key !== "feed-week") continue;
    rows.sort((x, y) => String(y.event_date).localeCompare(String(x.event_date)));
    const total = rows.reduce((t, e) => t + (e.in_amt || 0), 0);
    const collapsed = sublocCollapsed.has(key);
    const hd = document.createElement("div"); hd.className = "feed-bucket-head" + (collapsible ? " clickable" : "");
    const caret = document.createElement("span"); caret.className = "fb-caret"; caret.textContent = collapsible ? (collapsed ? "▸ " : "▾ ") : "";
    const lab = document.createElement("span"); lab.className = "fb-label"; lab.textContent = label;
    const mt = document.createElement("span"); mt.className = "fb-meta"; mt.textContent = rows.length ? `  ${rows.length} · ${money(total)} settled` : "  none";
    hd.appendChild(caret); hd.appendChild(lab); hd.appendChild(mt);
    if (collapsible) hd.onclick = () => { if (sublocCollapsed.has(key)) sublocCollapsed.delete(key); else sublocCollapsed.add(key); renderSublocFeed(); };
    box.appendChild(hd);
    if (rows.length && !(collapsible && collapsed)) {
      const scroll = document.createElement("div"); scroll.className = "table-scroll";
      const table = document.createElement("table"); table.className = "grid"; const thead = document.createElement("thead"), tbody = document.createElement("tbody");
      const htr = document.createElement("tr");
      for (const [c, al] of [["Date", "left"], ["Client paid (invoice)", "left"], ["Project", "left"], ["Settled subs", "right"], ["Lag", "right"], ["LOC balance after", "right"]]) { const th = document.createElement("th"); if (al === "left") th.className = "left"; th.textContent = c; htr.appendChild(th); }
      thead.appendChild(htr);
      for (const e of rows) { const tr = document.createElement("tr");
        const hasItems = e.settled && e.settled.length;
        if (hasItems) { tr.style.cursor = "pointer"; tr.title = "See the fronted subs this payment paid off";
          tr.onclick = (ev) => { if (ev.target.closest(".cell")) return; openSublocRepay(e); }; }
        tr.appendChild(leftText(fmtDateShort(e.event_date)));
        tr.appendChild(leftText((e.invoice ? "INV " + e.invoice : "–") + (e.party ? " · " + e.party : "")));
        tr.appendChild(leftText(e.project || "–"));
        const st = _slMcell(e.in_amt); st.querySelector(".cell").classList.add("st-ok"); tr.appendChild(st);
        tr.appendChild(rightText(e.lag_days != null ? e.lag_days + "d" : "–"));
        tr.appendChild(_slMcell(e.balance)); tbody.appendChild(tr); }
      table.appendChild(thead); table.appendChild(tbody); scroll.appendChild(table); box.appendChild(scroll);
    }
  }
}
// Project drill-over: open subs grouped by the draw (AR invoice) they sit under, each draw's
// status/details, each sub bill linking to QuickBooks, and a project → all-transactions link.
function openSublocDetail(project) {
  const p = (SUBLOC.open_by_project || {})[project]; if (!p) return;
  const nm = nameOf(project);
  $("#sublocDetailTitle").textContent = project + (nm ? " · " + nm : "");
  $("#sublocDetailSub").textContent = `${money(p.open)} fronted, still out`;
  const body = $("#sublocDetailBody"); body.innerHTML = "";
  { const acts = document.createElement("div"); acts.className = "pnl-actions"; const cu = qboCustomerUrl(p.cust_id);
    if (cu) { const a = document.createElement("a"); a.className = "btn"; a.href = cu; a.target = "_blank"; a.rel = "noopener"; a.textContent = "All transactions in QuickBooks ↗"; acts.appendChild(a); }
    if (acts.childNodes.length) body.appendChild(acts); }
  for (const g of (p.groups || [])) {
    const gd = document.createElement("div"); gd.className = "dgroup";
    const h = document.createElement("h4"); const dw = g.draw; const per = g.period || "no draw period";
    h.textContent = dw && dw.doc ? `Draw ${dw.doc} · ${per}` : per; gd.appendChild(h);
    const meta2 = document.createElement("div"); meta2.className = "drow";
    const dk = document.createElement("span"); dk.className = "dk"; dk.textContent = "Draw status";
    const dv = document.createElement("span"); dv.className = "dv";
    if (dw) { const st = document.createElement("span"); st.className = "st " + (dw.status === "Paid" ? "st-ok" : dw.status === "Unpaid" ? "st-bad" : "st-warn"); st.textContent = dw.status; dv.appendChild(st);
      dv.appendChild(document.createTextNode(`  ${money(dw.total)} billed${(dw.balance || 0) > 0.005 ? ` · ${money(dw.balance)} still owed` : ""}`)); }
    else dv.textContent = "not invoiced to the GC yet";
    meta2.appendChild(dk); meta2.appendChild(dv); gd.appendChild(meta2);
    const cap = document.createElement("div"); cap.className = "bills-cap"; cap.textContent = `${g.subs.length} sub${g.subs.length > 1 ? "s" : ""} · ${money(g.open)} still fronted`; gd.appendChild(cap);
    const scroll = document.createElement("div"); scroll.className = "table-scroll";
    const table = document.createElement("table"); table.className = "grid"; const thead = document.createElement("thead"), tbody = document.createElement("tbody");
    const htr = document.createElement("tr");
    for (const [c, al] of [["Sub", "left"], ["Bill #", "left"], ["Paid", "left"], ["Open $", "right"]]) { const th = document.createElement("th"); if (al === "left") th.className = "left"; th.textContent = c; htr.appendChild(th); }
    thead.appendChild(htr);
    for (const sub of g.subs) { const tr = document.createElement("tr");
      tr.appendChild(leftText(sub.party || "–"));
      tr.appendChild(qboLinkCell(sub.bill_ref || (sub.bill_id ? "#" + sub.bill_id : "–"), sub.bill_id ? qboUrl("bill", sub.bill_id) : null, "Open this bill in QuickBooks"));
      tr.appendChild(leftText(fmtDateShort(sub.date)));
      const o = _slMcell(sub.open); o.querySelector(".cell").classList.add("open-amt"); tr.appendChild(o); tbody.appendChild(tr); }
    table.appendChild(thead); table.appendChild(tbody); scroll.appendChild(table); gd.appendChild(scroll); body.appendChild(gd);
  }
  _sublocLoadSource(project, body);   // append the full LOC event chain (the source), fetched on demand
  openPanel("#sublocDetail");
}
// A client payment (repayment) → the fronted sub payments it paid down, FIFO oldest-first (the line items).
function openSublocRepay(e) {
  $("#sublocDetailTitle").textContent = (e.invoice ? "INV " + e.invoice : "Client payment") + (e.party ? " · " + e.party : "");
  const nm = nameOf(e.project);
  $("#sublocDetailSub").textContent = `${fmtDateShort(e.event_date)} · ${e.project || "–"}${nm ? " · " + nm : ""} · ${money(e.in_amt)} settled${e.lag_days != null ? " · " + e.lag_days + "d lag" : ""}`;
  const body = $("#sublocDetailBody"); body.innerHTML = "";
  const items = e.settled || [];
  const intro = document.createElement("p"); intro.className = "hint";
  intro.textContent = "The fronted sub payments this client payment paid down (FIFO, oldest first). Bill # opens QuickBooks.";
  body.appendChild(intro);
  const scroll = document.createElement("div"); scroll.className = "table-scroll";
  const table = document.createElement("table"); table.className = "sub-grid";
  table.innerHTML = "<thead><tr><th class='left'>Sub</th><th class='left'>Bill #</th><th class='left'>Fronted</th><th class='right'>Applied</th><th class='left'>Status</th></tr></thead>";
  const tb = document.createElement("tbody");
  for (const s of items) {
    const tr = document.createElement("tr");
    tr.appendChild(leftText(s.party || "–"));
    const bt = document.createElement("td"); bt.className = "left";
    if (s.bill_id) { const a = document.createElement("a"); a.href = qboUrl("bill", s.bill_id); a.target = "_blank"; a.rel = "noopener"; a.className = "qbo-link"; a.textContent = s.bill_ref || ("#" + s.bill_id); a.title = "Open this bill in QuickBooks"; bt.appendChild(a); }
    else bt.appendChild(document.createTextNode(s.bill_ref || "–"));
    tr.appendChild(bt);
    tr.appendChild(leftText(s.draw_date ? fmtDateShort(s.draw_date) : "–"));
    const at = document.createElement("td"); at.className = "right"; at.textContent = money(s.amount); tr.appendChild(at);
    tr.appendChild(leftText(s.fully ? "Fully collected" : "Partial"));
    tb.appendChild(tr);
  }
  if (!items.length) { const tr = document.createElement("tr"); const td = document.createElement("td"); td.colSpan = 5; td.className = "left dim"; td.textContent = "No matched sub line items on this payment."; tr.appendChild(td); tb.appendChild(tr); }
  table.appendChild(tb); scroll.appendChild(table); body.appendChild(scroll);
  openPanel("#sublocDetail");
}

// ── Sales / CRM (read-only from the Notion Customer List) ───────────────────
function daysAgo(iso) {
  if (!iso) return null;
  const t = Date.parse(iso.length <= 10 ? iso + "T00:00:00" : iso);
  return isNaN(t) ? null : Math.floor((Date.now() - t) / 864e5);
}
function buildHead(tableSel, cols) {
  const thead = $(tableSel + " thead"); thead.innerHTML = "";
  const htr = document.createElement("tr");
  for (const [c, al] of cols) { const th = document.createElement("th"); if (al === "left") th.className = "left"; th.textContent = c; htr.appendChild(th); }
  thead.appendChild(htr);
  return $(tableSel + " tbody");
}
function setSalesFilter(stage, div) {
  if ($("#salesStage")) $("#salesStage").value = stage;
  if ($("#salesDivision")) $("#salesDivision").value = div;
  renderSales();
  const t = $("#salesTable"); if (t) t.scrollIntoView({ behavior: "smooth", block: "start" });
}
function renderSales() {
  const S = SALES || {}, t = S.totals || {};
  const loaded = (S.customers || []).length > 0;
  $("#salesNote").textContent = loaded
    ? `(${t.customers || 0} customers · ${t.touches || 0} touches logged)`
    : "(no CRM data — run load_customers.py)";

  // ── KPI stats (clickable → filter/jump) ──
  const stats = [
    ["Customers", String(t.customers || 0), "in the list — click to clear filters", () => setSalesFilter("", "")],
    ["Interested", String(t.interested || 0), "warm — click to see them", () => setSalesFilter("Interested", "")],
    ["Touches logged", String(t.touches || 0), "interaction-log lines", null],
    ["Sales reps", String((S.by_rep || []).length), "working the list", () => { const el = $("#salesRepTable"); if (el) el.scrollIntoView({ behavior: "smooth", block: "start" }); }],
  ];
  const sr = $("#salesStats"); sr.innerHTML = "";
  for (const [label, value, sub, onClick] of stats) {
    const el = document.createElement("div"); el.className = "kpi" + (onClick ? " clickable" : "");
    el.innerHTML = `<div class="k-label"></div><div class="k-value"></div><div class="k-sub"></div>`;
    el.querySelector(".k-label").textContent = label;
    el.querySelector(".k-value").textContent = value;
    el.querySelector(".k-sub").textContent = sub;
    if (onClick) { el.onclick = onClick; el.title = "Click to filter"; }
    sr.appendChild(el);
  }

  // ── pipeline funnel (click a stage → filter the customer table) ──
  const pipe = S.pipeline || [];
  const maxc = pipe.reduce((m, p) => Math.max(m, p.customers || 0), 0) || 1;
  let tb = buildHead("#salesFunnel", [["Stage", "left"], ["Customers", "left"], ["Touches", "right"]]);
  tb.innerHTML = "";
  for (const p of pipe) {
    const tr = document.createElement("tr");
    tr.title = "Click to filter customers by this stage";
    tr.onclick = () => setSalesFilter(p.sales_status, $("#salesDivision") ? $("#salesDivision").value : "");
    tr.appendChild(leftText(p.sales_status));
    const st = document.createElement("td");
    const bar = document.createElement("span"); bar.className = "cell bar";
    const fill = document.createElement("span"); fill.className = "bar-fill"; fill.style.width = ((p.customers || 0) / maxc * 100) + "%";
    if (p.customers > 0) fill.style.minWidth = "7px";   // keep tiny counts (1, 2) visible
    const txt = document.createElement("span"); txt.className = "bar-txt"; txt.textContent = p.customers || 0;
    bar.appendChild(fill); bar.appendChild(txt);
    st.appendChild(bar); tr.appendChild(st);
    tr.appendChild(rightText(String(p.touches || 0)));
    tb.appendChild(tr);
  }

  // ── activity by rep ──
  tb = buildHead("#salesRepTable", [["Rep", "left"], ["Worked", "right"], ["Contacted", "right"], ["Interested", "right"], ["Won", "right"]]);
  tb.innerHTML = "";
  for (const r of (S.by_rep || [])) {
    const tr = document.createElement("tr");
    if (r.rep === activeRep) tr.className = "rep-active";
    tr.title = "Click for this rep's daily/weekly activity";
    tr.onclick = () => { activeRep = r.rep; renderSales(); };
    tr.appendChild(leftText(r.rep));
    tr.appendChild(rightText(String(r.worked || 0)));
    tr.appendChild(rightText(String(r.contacted || 0)));
    tr.appendChild(rightText(String(r.interested || 0)));
    tr.appendChild(rightText(String(r.won || 0)));
    tb.appendChild(tr);
  }
  renderRepActivity();

  // ── warm accounts with their touch log ──
  const warm = S.warm || [];
  $("#salesWarmNote").textContent = warm.length ? `(${warm.length} interested)` : "";
  const box = $("#salesWarm"); box.innerHTML = "";
  if (!warm.length) {
    const p = document.createElement("p"); p.className = "hint"; p.textContent = "No accounts in the Interested stage.";
    box.appendChild(p);
  }
  for (const a of warm) {
    const card = document.createElement("div"); card.className = "warm-acct";
    const head = document.createElement("div"); head.className = "warm-head";
    const nm = document.createElement("span"); nm.className = "warm-name"; nm.textContent = a.name; head.appendChild(nm);
    if (a.division) { const dv = document.createElement("span"); dv.className = "vtype"; dv.textContent = a.division; head.appendChild(dv); }
    const d = daysAgo(a.last_contacted);
    const when = document.createElement("span"); when.className = "warm-when" + (d !== null && d > 21 ? " stale" : "");
    when.textContent = a.last_contacted ? `last ${fmtDate(a.last_contacted)}${d !== null ? ` · ${d}d ago` : ""}` : "no contact date";
    head.appendChild(when);
    if (a.last_edited_by) { const by = document.createElement("span"); by.className = "warm-by"; by.textContent = a.last_edited_by; head.appendChild(by); }
    if (a.notion_url) { const lk = document.createElement("a"); lk.className = "warm-link"; lk.href = a.notion_url; lk.target = "_blank"; lk.rel = "noopener"; lk.textContent = "Notion ↗"; head.appendChild(lk); }
    card.appendChild(head);
    const ul = document.createElement("ul"); ul.className = "warm-log";
    if ((a.touches || []).length) {
      for (const tch of a.touches) {
        const li = document.createElement("li");
        if (tch.date) { const dspan = document.createElement("span"); dspan.className = "d"; dspan.textContent = tch.date; li.appendChild(dspan); }
        li.appendChild(document.createTextNode(tch.note));
        ul.appendChild(li);
      }
    } else {
      const li = document.createElement("li"); li.className = "warm-empty"; li.textContent = "No notes logged yet."; ul.appendChild(li);
    }
    card.appendChild(ul);
    box.appendChild(card);
  }

  // ── all-customers table (search + stage + division filters) ──
  const q = ($("#salesSearch") ? $("#salesSearch").value : "").trim().toLowerCase();
  const stage = $("#salesStage") ? $("#salesStage").value : "";
  const div = $("#salesDivision") ? $("#salesDivision").value : "";
  let rows = S.customers || [];
  if (stage) rows = rows.filter(c => c.sales_status === stage);
  if (div) rows = rows.filter(c => (c.division || "") === div);
  if (q) rows = rows.filter(c => [c.name, c.last_edited_by, c.sales_status, c.division].filter(Boolean).join(" ").toLowerCase().includes(q));
  $("#salesCustNote").textContent = (S.customers || []).length ? `(${rows.length} shown)` : "";
  tb = buildHead("#salesTable", [["Client", "left"], ["Division", "left"], ["Stage", "left"], ["Last contacted", "left"], ["Worked by", "left"], ["Touches", "right"]]);
  tb.innerHTML = "";
  for (const c of rows.slice(0, 400)) {
    const tr = document.createElement("tr");
    if (c.notion_url) { tr.onclick = () => window.open(c.notion_url, "_blank", "noopener"); tr.title = "Open in Notion"; }
    const nameTd = document.createElement("td"); nameTd.className = "left";
    if (c.notion_url) {
      const a = document.createElement("a"); a.className = "row-link"; a.href = c.notion_url; a.target = "_blank"; a.rel = "noopener";
      a.textContent = c.name; a.onclick = (e) => e.stopPropagation();
      nameTd.appendChild(a);
    } else { const s = document.createElement("span"); s.textContent = c.name; nameTd.appendChild(s); }
    tr.appendChild(nameTd);
    tr.appendChild(leftText(c.division || "—"));
    tr.appendChild(leftText(c.sales_status || "—"));
    tr.appendChild(leftText(c.last_contacted ? fmtDate(c.last_contacted) : "—"));
    tr.appendChild(leftText(c.last_edited_by || "—"));
    tr.appendChild(rightText(String(c.n_touches || 0)));
    tb.appendChild(tr);
  }
  if (!rows.length) {
    const tr = document.createElement("tr"); const td = document.createElement("td");
    td.colSpan = 6; td.className = "left"; td.style.color = "var(--text-dim)";
    td.textContent = (S.customers || []).length ? "No customers match this filter." : "No CRM data — run load_customers.py.";
    tr.appendChild(td); tb.appendChild(tr);
  }
}

// Per-rep activity drill (defaults to the outreach rep = most touches). Everything is
// derived from the live payload; the rep is a runtime value, never hard-coded.
function renderRepActivity() {
  const box = $("#repActivity"); if (!box) return;
  const S = SALES || {};
  const log = S.touch_log || [], custs = S.customers || [], reps = S.by_rep || [];
  // auto-pick the busiest-by-touches rep if none is selected (or the selection is gone)
  if (!activeRep || !reps.some(r => r.rep === activeRep)) {
    const cnt = {}; for (const t of log) cnt[t.rep] = (cnt[t.rep] || 0) + 1;
    const ranked = reps.map(r => r.rep).filter(r => cnt[r]).sort((a, b) => cnt[b] - cnt[a]);
    activeRep = ranked[0] || (reps[0] && reps[0].rep) || null;
  }
  const note = $("#repActNote"); box.innerHTML = "";
  if (!activeRep) { if (note) note.textContent = ""; box.appendChild(el2("p", "hint", "No rep activity — run load_customers.py.")); return; }
  const rep = activeRep;
  if (note) note.textContent = `— ${rep} · daily & weekly`;
  const mine = log.filter(t => t.rep === rep);
  const myc = custs.filter(c => c.last_edited_by === rep);

  // local date helpers (Monday-anchored weeks, no UTC drift)
  const monday = s => { const m = String(s || "").match(/^(\d{4})-(\d{2})-(\d{2})/); if (!m) return null;
    const d = new Date(+m[1], +m[2] - 1, +m[3]); d.setDate(d.getDate() - ((d.getDay() + 6) % 7)); d.setHours(0, 0, 0, 0); return d; };
  const ymd = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const todayStr = ymd(today), thisMon = monday(todayStr), thisMonStr = ymd(thisMon);
  const lastMon = new Date(thisMon); lastMon.setDate(lastMon.getDate() - 7); const lastMonStr = ymd(lastMon);
  const wk = t => { const md = monday(t.date); return md ? ymd(md) : null; };
  const thisWeek = mine.filter(t => wk(t) === thisMonStr).length;
  const lastWeek = mine.filter(t => wk(t) === lastMonStr).length;
  const todayN = mine.filter(t => t.date === todayStr).length;
  const wins = myc.filter(c => c.sales_status === "Closed - Won").length;
  const lastActive = mine.length ? mine.map(t => t.date).sort().slice(-1)[0] : null;

  // header
  const head = el2("div", "rep-head"); const h = el2("div");
  h.appendChild(el2("h3", null, rep));
  h.appendChild(el2("span", "rep-sub", `${mine.length} touches · ${myc.length} customers · ${wins} won · last active ${lastActive ? fmtDate(lastActive) : "—"}`));
  head.appendChild(h); box.appendChild(head);

  // summary tiles
  const trend = lastWeek ? Math.round((thisWeek - lastWeek) / lastWeek * 100) : null;
  const kr = el2("div", "kpi-row");
  [["This week", String(thisWeek), trend == null ? "touches logged" : `${trend >= 0 ? "▲" : "▼"} ${Math.abs(trend)}% vs last week`],
   ["Last week", String(lastWeek), "touches logged"],
   ["Today", String(todayN), "touches logged"],
   ["All-time", String(mine.length), "touches logged"]].forEach(([l, v, s]) => {
    const k = el2("div", "kpi"); k.appendChild(el2("div", "k-label", l)); k.appendChild(el2("div", "k-value", v)); k.appendChild(el2("div", "k-sub", s)); kr.appendChild(k);
  });
  box.appendChild(kr);

  const cols = el2("div", "rep-cols"); const left = el2("div", "rep-col"), right = el2("div", "rep-col");
  cols.appendChild(left); cols.appendChild(right); box.appendChild(cols);

  // WEEKLY timeline — last 12 weeks including zero weeks (so a drop-off shows)
  left.appendChild(el2("h4", null, "Weekly touches — last 12 weeks"));
  const weeks = []; for (let i = 11; i >= 0; i--) { const d = new Date(thisMon); d.setDate(d.getDate() - i * 7); weeks.push(ymd(d)); }
  const wc = {}; for (const t of mine) { const k = wk(t); if (k) wc[k] = (wc[k] || 0) + 1; }
  const wmax = Math.max(1, ...weeks.map(w => wc[w] || 0));
  const wt = el2("div", "rep-weeks");
  for (const w of weeks) {
    const n = wc[w] || 0; const row = el2("div", "rep-week");
    row.appendChild(el2("span", "rw-lab", fmtDate(w).replace(/^\w+, /, "")));
    const bar = el2("span", "rw-bar"); const fill = el2("span", "bar-fill"); fill.style.width = (n / wmax * 100) + "%"; if (n > 0) fill.style.minWidth = "7px";
    bar.appendChild(fill); row.appendChild(bar); row.appendChild(el2("span", "rw-val", String(n)));
    wt.appendChild(row);
  }
  left.appendChild(wt);

  // RECENT touches (what was actually done)
  left.appendChild(el2("h4", null, "Recent touches"));
  const rlog = el2("div", "rep-log");
  const recent = [...mine].sort((a, b) => a.date < b.date ? 1 : -1).slice(0, 15);
  if (!recent.length) rlog.appendChild(el2("p", "hint", "No touches logged."));
  for (const t of recent) {
    const line = el2("div", "rep-touch");
    line.appendChild(el2("span", "rt-date", fmtDate(t.date).replace(/^\w+, /, "").replace(/, \d{4}$/, "")));
    const body = el2("span", "rt-body"); body.appendChild(el2("span", "rt-cust", t.customer || "—"));
    if (t.note) body.appendChild(el2("span", "rt-note", t.note));
    line.appendChild(body); rlog.appendChild(line);
  }
  left.appendChild(rlog);

  const open = c => c.sales_status && c.sales_status.indexOf("Closed") !== 0;
  const itemName = c => { if (c.notion_url) { const a = el2("a", "ri-name", c.name); a.href = c.notion_url; a.target = "_blank"; a.rel = "noopener"; return a; } return el2("span", "ri-name", c.name); };

  // FOLLOW-UPS due (open + follow_up_date on/before today)
  right.appendChild(el2("h4", null, "Follow-ups due"));
  const due = myc.filter(c => open(c) && c.follow_up_date && c.follow_up_date <= todayStr).sort((a, b) => a.follow_up_date < b.follow_up_date ? -1 : 1);
  const dueBox = el2("div", "rep-list");
  if (!due.length) dueBox.appendChild(el2("p", "hint", "Nothing due."));
  for (const c of due.slice(0, 10)) { const li = el2("div", "rep-item"); li.appendChild(itemName(c)); li.appendChild(el2("span", "ri-meta", `${fmtDate(c.follow_up_date)} · ${c.sales_status}`)); dueBox.appendChild(li); }
  right.appendChild(dueBox);

  // GOING STALE (open + no contact in 21+ days)
  right.appendChild(el2("h4", null, "Going stale — 21d+ no contact"));
  const stale = myc.filter(c => open(c) && c.last_contacted && daysAgo(c.last_contacted) > 21).sort((a, b) => daysAgo(b.last_contacted) - daysAgo(a.last_contacted));
  const staleBox = el2("div", "rep-list");
  if (!stale.length) staleBox.appendChild(el2("p", "hint", "Nothing stale."));
  for (const c of stale.slice(0, 10)) { const li = el2("div", "rep-item"); li.appendChild(itemName(c)); li.appendChild(el2("span", "ri-meta", `${daysAgo(c.last_contacted)}d · ${c.sales_status}`)); staleBox.appendChild(li); }
  right.appendChild(staleBox);

  // PIPELINE (their customers by stage)
  right.appendChild(el2("h4", null, "Their pipeline"));
  const byStage = {}; for (const c of myc) { const s = c.sales_status || "(none)"; byStage[s] = (byStage[s] || 0) + 1; }
  const pipeBox = el2("div", "rep-list");
  ["Lead", "Follow up", "Contacted", "Interested", "No response", "Closed - Won", "Closed - Lost", "(none)"].filter(s => byStage[s]).forEach(s => {
    const li = el2("div", "rep-item"); li.appendChild(el2("span", "ri-name", s)); li.appendChild(el2("span", "ri-meta", String(byStage[s]))); pipeBox.appendChild(li);
  });
  right.appendChild(pipeBox);
}
// tiny DOM helper (local to the sales drill)
function el2(tag, cls, txt) { const e = document.createElement(tag); if (cls) e.className = cls; if (txt != null) e.textContent = txt; return e; }

// Portfolio P&L tab — every active job's live P&L + division/company totals. Computed
// server-side (/api/pnl/portfolio), lazy-loaded on first open, recomputed after a reload.
// ══ WIP REPORT ═══════════════════════════════════════════════════════════════
// The company Work-in-Progress schedule, straight from the ledger's wip_snapshot (loaded from the
// WIP master's Test tabs). Columns + order mirror the Test-Master reference; grouped by division
// with subtotals and a grand total. Read-only - the master workbook stays where you EDIT the WIP.
// Bonded is intentionally NOT here - dropped from the dashboard WIP view (the user
// 2026-08-20); the Excel Test tabs keep it. `cf` marks a column that carries a
// job-performance conditional format (see _wipCond).
const WIP_COLS = [
  { k: "project_no", label: "Project #", t: "text" },
  { k: "project_name", label: "Name", t: "text" },
  { k: "total_contract_price", label: "Total Contract", t: "money" },
  { k: "estimated_total_costs", label: "Est. Total Costs", t: "money" },
  { k: "original_profit", label: "Original Profit", t: "money" },
  { k: "gross_profit_pct", label: "GP %", t: "pct", cf: "gp" },
  { k: "costs_to_date", label: "Costs to Date", t: "money" },
  { k: "cost_to_complete", label: "Cost to Complete", t: "money" },
  { k: "percent_complete", label: "% Complete", t: "pct", cf: "pctbar" },
  { k: "revenues_earned_to_date", label: "Revenues Earned", t: "money" },
  { k: "profit_earned_to_date", label: "Profit Earned", t: "money", cf: "neg0" },
  { k: "billed_to_date", label: "Billed", t: "money" },
  { k: "overbillings", label: "Overbillings", t: "money", cf: "over" },
  { k: "underbillings", label: "Underbillings", t: "money", cf: "under" },
  { k: "left_to_bill", label: "Left to Bill", t: "money" },
  { k: "future_profit_to_earn", label: "Future Profit", t: "money", cf: "future" },
  { k: "pure_job_borrow", label: "Pure Job Borrow", t: "money", cf: "borrow" },
];
const WIP_DIV_ORDER = ["Multi Family", "Commercial", "Residential"];
function _wipFmt(c, v) {
  if (c.t === "pct") return v == null || v === "" ? "–" : (v * 100).toFixed(1) + "%";
  return v == null || v === "" ? "–" : String(v);
}

// ── WIP column widths (drag a header divider; persists per person) ──────────
const WIP_COL_DEFAULTS = { "Project #": 88, "Name": 190, "Total Contract": 122,
  "Est. Total Costs": 122, "Original Profit": 118, "GP %": 74, "Costs to Date": 116,
  "Cost to Complete": 128, "% Complete": 96, "Revenues Earned": 128, "Profit Earned": 116,
  "Billed": 112, "Overbillings": 116, "Underbillings": 118, "Left to Bill": 112,
  "Future Profit": 116, "Pure Job Borrow": 128 };
function loadWipColWidths() {
  try { return { ...WIP_COL_DEFAULTS, ...JSON.parse(localStorage.getItem("proficient-ledger-wipcols") || "{}") }; }
  catch { return { ...WIP_COL_DEFAULTS }; }
}
let wipColW = loadWipColWidths();
function saveWipColWidths() { try { localStorage.setItem("proficient-ledger-wipcols", JSON.stringify(wipColW)); } catch { /* ignore */ } }
function startWipColResize(e, idx, label) {
  e.preventDefault(); e.stopPropagation();
  const table = $("#wipTable"); const cg = table.querySelector("colgroup"); if (!cg) return;
  const col = cg.children[idx]; const startX = e.clientX; const startW = parseFloat(col.style.width) || col.offsetWidth;
  document.body.classList.add("col-resizing");
  const onMove = (ev) => {
    const w = Math.max(48, Math.round(startW + (ev.clientX - startX)));
    col.style.width = w + "px"; wipColW[label] = w;
    let s = 0; for (const c of cg.children) s += parseFloat(c.style.width) || 0; table.style.width = s + "px";
  };
  const onUp = () => { document.removeEventListener("mousemove", onMove); document.removeEventListener("mouseup", onUp);
    document.body.classList.remove("col-resizing"); saveWipColWidths(); };
  document.addEventListener("mousemove", onMove); document.addEventListener("mouseup", onUp);
}

// ── Conditional formatting: encode job health with color, never decoration ──
// Returns {bg, fg, bold, bar, title} for one cell, or null. Sign conventions come
// straight from the WIP writer's formulas (wip/wip_writer.py):
//   OVERBILLINGS  = MAX(Billed − Earned, 0)  → holding the GC's cash (good, green)
//   UNDERBILLINGS = MAX(Earned − Billed, 0)  → earned but not billed (financing, red)
//   PURE JOB BORROW = MAX(CostToComplete − LeftToBill, 0) → cash drain (red)
//   FUTURE PROFIT = Original − Earned profit  → negative means eroded (red)
//   GP %: thin/negative red, healthy green, >30% amber (owner's "missing cost" flag)
function _mix(varName, pct) { return `color-mix(in srgb, ${varName} ${pct}%, transparent)`; }
function _wipCond(kind, r, key) {
  const contract = Math.max(num(r.total_contract_price), 1);
  if (kind === "gp") {
    const v = r.gross_profit_pct; if (v == null || v === "") return null;
    if (v < 0.05) return { bg: _mix("var(--neg)", 20), fg: "var(--neg)", bold: true, title: "Margin very thin / negative" };
    if (v < 0.12) return { bg: _mix("#b8860b", 18), title: "Below-target margin" };
    if (v > 0.30) return { bg: _mix("#b8860b", 20), fg: "#8a6508", title: "Unusually high GP% - verify for a missing cost" };
    return { bg: _mix("var(--pos)", 15), title: "Healthy margin" };
  }
  if (kind === "pctbar") {
    const v = r.percent_complete; if (v == null || v === "") return null;
    const p = Math.max(0, Math.min(100, v * 100));
    if (v > 1.0005) return { bar: 100, bg: _mix("var(--neg)", 18), fg: "var(--neg)", bold: true,
                             title: (v * 100).toFixed(1) + "% - costs to date exceed the ETC (over budget)" };
    return { bar: p, title: p.toFixed(1) + "% complete" };
  }
  if (kind === "over") {           // overbilled = holding cash = positive
    const v = num(r.overbillings); if (v <= 0) return null;
    const a = 8 + Math.min(20, (v / contract) * 120);
    return { bg: _mix("var(--pos)", a), title: "Billed ahead of earned - holding the GC's cash" };
  }
  if (kind === "under") {          // underbilled = financing the job = red flag
    const v = num(r.underbillings); if (v <= 0) return null;
    const ratio = v / contract, a = 8 + Math.min(24, ratio * 140);
    return { bg: _mix("var(--neg)", a), fg: ratio > 0.08 ? "var(--neg)" : null, bold: ratio > 0.08,
      title: "Earned ahead of billed - unbilled work you are financing" };
  }
  if (kind === "borrow") {         // pure job borrow = cash the job pulls to finish
    const v = num(r.pure_job_borrow); if (v <= 0) return null;
    const ratio = v / contract;
    if (ratio < 0.05) return { bg: _mix("#b8860b", 16), title: "This job borrows some cash to finish" };
    return { bg: _mix("var(--neg)", 8 + Math.min(22, ratio * 130)), fg: "var(--neg)", bold: true,
      title: "Cost to complete exceeds what is left to bill - a cash drain" };
  }
  if (kind === "future") {         // remaining profit to earn
    const v = r.future_profit_to_earn; if (v == null || v === "") return null;
    if (v < 0) return { bg: _mix("var(--neg)", 20), fg: "var(--neg)", bold: true, title: "Expected profit eroded below what is already earned" };
    if (v > 0) return { bg: _mix("var(--pos)", 10), title: "Profit still ahead to earn" };
    return null;
  }
  if (kind === "neg0") {           // any money col that is a red flag when negative
    if (num(r[key]) < 0) return { bg: _mix("var(--neg)", 20), fg: "var(--neg)", bold: true, title: "Negative - losing money to date" };
    return null;
  }
  return null;
}

function renderWip() {
  const thead = $("#wipTable thead"), tbody = $("#wipTable tbody"); if (!thead || !tbody) return;
  const activeOnly = $("#wipActive") ? $("#wipActive").checked : true;
  const isActive = r => { const s = (r.status || "").trim().toLowerCase(); return s === "" || s === "active"; };
  const rows = (ALL || []).filter(r => !activeOnly || isActive(r));
  if ($("#wipNote")) $("#wipNote").textContent = `(${rows.length} job${rows.length === 1 ? "" : "s"} · report ${meta && meta.report_date ? fmtDate(meta.report_date) : "–"})`;
  if ($("#wipHint")) $("#wipHint").innerHTML = "The company Work-in-Progress schedule, live from the ledger (loaded from the WIP master's Test tabs). Columns mirror <b>Test-Master</b>; grouped by division with subtotals. Read-only here - edit the WIP in the master workbook.";
  const byDiv = new Map();
  for (const r of rows) { const d = r.division || "Other"; if (!byDiv.has(d)) byDiv.set(d, []); byDiv.get(d).push(r); }
  const order = [...byDiv.keys()].sort((a, b) => { const ia = WIP_DIV_ORDER.indexOf(a), ib = WIP_DIV_ORDER.indexOf(b); return (ia === -1 ? 99 : ia) - (ib === -1 ? 99 : ib) || a.localeCompare(b); });
  thead.innerHTML = ""; tbody.innerHTML = "";
  // Fixed layout + a <colgroup> so widths are exact and draggable; the sticky
  // thead (base .grid rule) then freezes as the bounded container scrolls.
  const table = $("#wipTable");
  { const oldCg = table.querySelector("colgroup"); if (oldCg) oldCg.remove(); }
  const colgroup = document.createElement("colgroup");
  const htr = document.createElement("tr");
  let wsum = 0;
  WIP_COLS.forEach((c, i) => {
    const th = document.createElement("th"); if (c.t === "text") th.className = "left"; th.textContent = c.label;
    const grip = document.createElement("div"); grip.className = "col-resize"; grip.title = "Drag to resize this column";
    grip.addEventListener("mousedown", (e) => startWipColResize(e, i, c.label));
    th.appendChild(grip); htr.appendChild(th);
    const w = Math.max(48, wipColW[c.label] || 110); const col = document.createElement("col"); col.style.width = w + "px";
    colgroup.appendChild(col); wsum += w;
  });
  table.insertBefore(colgroup, table.firstChild);
  table.style.width = wsum + "px";
  thead.appendChild(htr);
  if (!rows.length) { const tr = document.createElement("tr"); const td = document.createElement("td"); td.colSpan = WIP_COLS.length; td.className = "left"; td.style.color = "var(--text-dim)"; td.style.padding = "14px 12px"; td.textContent = "No WIP data - run load_wip_master.py."; tr.appendChild(td); tbody.appendChild(tr); return; }
  const sumRow = (label, list, cls) => {
    const tr = document.createElement("tr"); tr.className = cls;
    WIP_COLS.forEach((c, i) => { const td = document.createElement("td");
      if (i === 0) { td.className = "left"; td.textContent = label; td.colSpan = 2; }   // label spans Project # + Name (was cut at 88px)
      else if (i === 1) return;
      else if (c.t === "money") { td.className = "right"; td.appendChild(moneyCell(list.reduce((t, r) => t + num(r[c.k]), 0))); }
      else td.className = c.t === "text" ? "left" : "right";
      tr.appendChild(td); });
    return tr;
  };
  for (const div of order) {
    const list = byDiv.get(div);
    const gtr = document.createElement("tr"); gtr.className = "bill-group";
    const gtd = document.createElement("td"); gtd.colSpan = WIP_COLS.length;
    const cell = document.createElement("div"); cell.className = "bg-cell";
    const key = document.createElement("span"); key.className = "bg-key"; key.textContent = div;
    cell.appendChild(key);
    const sumOf = k2 => list.reduce((t2, r) => t2 + num(r[k2]), 0);
    bandMetrics(cell, [[list.length, "jobs"], [money(sumOf("total_contract_price")), "contract"], [money(sumOf("costs_to_date")), "costs to date"], [money(sumOf("billed_to_date")), "billed"], [money(sumOf("underbillings")), "underbilled", sumOf("underbillings") > 0 ? "neg" : ""]]);
    gtd.appendChild(cell); gtr.appendChild(gtd); tbody.appendChild(gtr);
    for (const r of list) {
      const tr = document.createElement("tr"); tr.style.cursor = "pointer"; tr.title = "Open this project";
      tr.onclick = (e) => { if (e.target.closest(".cell")) return; openDetail(r); };
      for (const c of WIP_COLS) {
        const td = document.createElement("td");
        if (c.t === "money") td.appendChild(moneyCell(r[c.k]));
        else if (c.t === "text") { td.className = "left"; const s = document.createElement("span"); s.textContent = _wipFmt(c, r[c.k]); s.title = _wipFmt(c, r[c.k]); td.appendChild(s); }
        else { td.className = "right"; td.textContent = _wipFmt(c, r[c.k]); }
        if (c.cf) {
          const cond = _wipCond(c.cf, r, c.k);
          if (cond) {
            if (cond.bar != null) { td.classList.add("wip-bar"); td.style.setProperty("--bar", cond.bar + "%"); }
            if (cond.bg) td.style.background = cond.bg;
            if (cond.fg) td.style.color = cond.fg;
            if (cond.bold) td.style.fontWeight = "700";
            if (cond.title) td.title = cond.title;
          }
        }
        tr.appendChild(td);
      }
      tbody.appendChild(tr);
    }
    tbody.appendChild(sumRow(div + " total", list, "wip-subtotal"));
  }
  tbody.appendChild(sumRow("GRAND TOTAL", rows, "wip-total"));
}
function renderPnl() {
  if (!PNL) {
    const n = $("#pnlNote"); if (n) n.textContent = "computing…";
    skeletonInto($("#pnlJobTable") ? $("#pnlJobTable").querySelector("tbody") : null, 6);
    fetch("/api/pnl/portfolio").then(r => r.json()).then(d => { PNL = d.error ? { rows: [], by_division: [], company: {} } : d; renderPnl(); })
      .catch(() => { const e = $("#pnlNote"); if (e) e.textContent = "unavailable"; });
    return;
  }
  const rows = PNL.rows || [], divs = PNL.by_division || [], comp = PNL.company || {};
  const pctTxt = p => (p == null ? "—" : (p * 100).toFixed(1) + "%");
  $("#pnlNote").textContent = rows.length ? `(${comp.n || 0} active jobs · ${money(comp.earned)} earned)` : "(no P&L data - load WIP + costs)";

  // company totals
  const tiles = [["Earned revenue", money(comp.earned)], ["Costs", money(comp.cost)],
    ["Overhead", money(comp.overhead)], ["Net margin", `${money(comp.net)} · ${pctTxt(comp.net_pct)}`, comp.net == null ? "" : (comp.net >= 0 ? "pos" : "neg")],
    ["Billed (AR)", money(comp.billed)]];
  const tr = $("#pnlTotals"); tr.innerHTML = "";
  for (const [l, v, cls] of tiles) {
    const k = el2("div", "kpi" + (cls ? " pnl-kpi-" + cls : ""));
    k.appendChild(el2("div", "k-label", l)); k.appendChild(el2("div", "k-value", v)); tr.appendChild(k);
  }

  // by division
  let tb = buildHead("#pnlDivTable", [["Division", "left"], ["Jobs", "right"], ["Earned", "right"], ["Cost", "right"], ["Overhead", "right"], ["Net", "right"], ["Net %", "right"]]);
  tb.innerHTML = "";
  for (const d of divs) {
    const row = document.createElement("tr");
    row.appendChild(leftText(d.division)); row.appendChild(rightText(String(d.n)));
    row.appendChild(rightText(money(d.earned))); row.appendChild(rightText(money(d.cost)));
    row.appendChild(rightText(money(d.overhead))); row.appendChild(rightText(money(d.net)));
    const pt = document.createElement("td"); pt.className = d.net >= 0 ? "pos" : "neg"; pt.textContent = pctTxt(d.net_pct); row.appendChild(pt);
    tb.appendChild(row);
  }

  // by job — filterable + sortable (headers), click → detail
  const fProj = ($("#pnlFProj") ? $("#pnlFProj").value : "").trim().toLowerCase();
  const fDiv = $("#pnlFDivision") ? $("#pnlFDivision").value : "";
  const fClient = ($("#pnlFClient") ? $("#pnlFClient").value : "").trim().toLowerCase();
  const fStatus = $("#pnlFStatus") ? $("#pnlFStatus").value : "active";   // default: active jobs only
  let shown = rows.filter(r => (!fProj || r.proj.toLowerCase().includes(fProj))
    && (!fDiv || r.division === fDiv)
    && (!fClient || (r.client || "").toLowerCase().includes(fClient))
    && (fStatus === "all" ? true : (fStatus === "closed" ? !r.active : r.active)));
  shown.sort((a, b) => { const k = pnlSort.key, av = a[k], bv = b[k];
    if (av == null) return 1; if (bv == null) return -1;
    return (av < bv ? -1 : av > bv ? 1 : 0) * pnlSort.dir; });
  const cols = [["proj", "Project", "left"], ["name", "Name / address", "left"], ["division", "Division", "left"], ["client", "Client", "left"],
    ["status", "Status", "left"], ["contract", "Contract", "right"], ["pct_complete", "%", "right"], ["earned", "Earned", "right"],
    ["cost", "Cost", "right"], ["overhead", "Overhead", "right"], ["net", "Net", "right"],
    ["net_pct", "Net %", "right"], ["pnl_mtime", "P&L updated", "right"]];
  const thead = $("#pnlJobTable thead"), tbody = $("#pnlJobTable tbody"); thead.innerHTML = ""; tbody.innerHTML = "";
  const htr = document.createElement("tr");
  for (const [key, label, al] of cols) {
    const th = document.createElement("th"); if (al === "left") th.className = "left";
    th.textContent = label + (pnlSort.key === key ? (pnlSort.dir < 0 ? " ▾" : " ▴") : "");
    th.style.cursor = "pointer";
    th.onclick = () => { if (pnlSort.key === key) pnlSort.dir *= -1; else { pnlSort.key = key; pnlSort.dir = (key === "proj" || key === "name" || key === "division" || key === "client" || key === "status") ? 1 : -1; } renderPnl(); };
    htr.appendChild(th);
  }
  thead.appendChild(htr);
  const known = new Set(ALL.map(r => r.project_no));
  for (const r of shown) {
    const row = document.createElement("tr");
    const open = pnlExpanded.has(r.proj);
    if (known.has(r.proj)) {
      row.style.cursor = "pointer"; if (open) row.className = "pnl-open";
      // expand the P&L inline (full width, room for the dense numbers) - no side panel
      row.onclick = (e) => { if (e.target.closest(".cell")) return; open ? pnlExpanded.delete(r.proj) : pnlExpanded.add(r.proj); renderPnl(); };
    }
    const pcell = document.createElement("td"); pcell.className = "left";
    if (known.has(r.proj)) pcell.appendChild(document.createTextNode(open ? "▾ " : "▸ "));
    const ppurl = qboCustomerUrl(r.cust_id);   // project # → QBO project page (all its transactions)
    if (ppurl) {
      const a = document.createElement("a"); a.href = ppurl; a.target = "_blank"; a.rel = "noopener"; a.className = "qbo-link";
      a.textContent = r.proj; a.title = "Open this project in QuickBooks (all transactions)"; a.onclick = e => e.stopPropagation();
      pcell.appendChild(a);
    } else { pcell.appendChild(document.createTextNode(r.proj)); }
    row.appendChild(pcell);
    { const c = leftText(r.name || "–"); c.style.color = r.name ? "" : "var(--text-dim)"; c.title = r.name || ""; row.appendChild(c); }
    row.appendChild(leftText(r.division));
    { const c = leftText(r.client || "–"); c.style.color = r.client ? "" : "var(--text-dim)"; row.appendChild(c); }
    { const s = document.createElement("td"); s.className = "left"; s.appendChild(stText(r.status || "Active", r.active ? "st-ok" : "st-dim")); row.appendChild(s); }
    row.appendChild(rightText(money(r.contract))); row.appendChild(rightText(((r.pct_complete || 0) * 100).toFixed(0) + "%"));
    row.appendChild(rightText(money(r.earned))); row.appendChild(rightText(money(r.cost)));
    row.appendChild(rightText(money(r.overhead))); row.appendChild(rightText(money(r.net)));
    const pt = document.createElement("td"); pt.className = r.net >= 0 ? "pos" : "neg"; pt.textContent = pctTxt(r.net_pct); row.appendChild(pt);
    // P&L updated = when this project's project-pnl Excel was last generated (owner 2026-08-19).
    const upd = document.createElement("td"); upd.className = "right"; upd.style.color = "var(--text-dim)"; upd.style.fontSize = ".88em";
    if (r.pnl_mtime) { upd.textContent = timeAgo(r.pnl_mtime); upd.title = "P&L Excel generated " + fmtDate(r.pnl_mtime, true); }
    else { upd.textContent = "not generated"; upd.style.opacity = ".55"; }
    row.appendChild(upd);
    tbody.appendChild(row);
    if (open) {
      const er = document.createElement("tr"); er.className = "pnl-expand-row";
      const td = document.createElement("td"); td.colSpan = cols.length; td.className = "pnl-expand";
      td.appendChild(buildPnlGroup(r.proj));
      er.appendChild(td); tbody.appendChild(er);
    }
  }
  if (!shown.length) { const row = document.createElement("tr"); const td = document.createElement("td"); td.colSpan = cols.length; td.className = "left"; td.style.color = "var(--text-dim)"; td.textContent = rows.length ? "No jobs match this filter." : "No P&L data yet."; row.appendChild(td); tbody.appendChild(row); }
}

function renderAttention() {
  const row = $("#attnRow"); row.innerHTML = "";
  for (const rule of RULES) {
    const hits = ALL.filter(rule.test);
    const total = hits.reduce((t, r) => t + rule.amt(r), 0);
    const el = document.createElement("div");
    el.className = "attn" + (rule.warn ? " warn" : "") + (activeRule === rule.key ? " active" : "") + (hits.length ? "" : " none");
    el.innerHTML = `<span class="a-count"></span><span class="a-label"></span><span class="a-sub"></span>`;
    el.querySelector(".a-count").textContent = hits.length;
    el.querySelector(".a-label").textContent = rule.label;
    el.querySelector(".a-sub").textContent = hits.length ? money(total) : rule.hint;
    el.title = rule.hint + (hits.length ? " — click to filter the table" : "");
    if (hits.length) el.onclick = () => {
      activeRule = activeRule === rule.key ? null : rule.key;
      renderAttention(); renderProjects();
      $("#btnClearRule").hidden = !activeRule;
    };
    row.appendChild(el);
  }
  $("#btnClearRule").hidden = !activeRule;
}

function renderKPIs() {
  const rows = ALL;
  const sum = k => rows.reduce((t, r) => t + (isNum(r[k]) ? r[k] : Number(r[k]) || 0), 0);
  const contract = sum("total_contract_price"), costs = sum("costs_to_date"),
        billed = sum("billed_to_date"), left = sum("left_to_bill"),
        over = sum("overbillings"), under = sum("underbillings"),
        active = rows.filter(isActive).length;
  const net = over - under;
  // Every card says where its number comes from and as of when (owner 2026-09-01: an Excel-only
  // reader who does not trust a figure until he sees its source). These six are the WIP master's
  // own columns, summed - not QuickBooks (that total is on the Cost mix widget, labelled).
  const wipSrc = srcText("WIP master", meta.report_date, "report");
  const cards = [
    ["Total Contract", money(contract), `${rows.length} jobs`, wipSrc, "sum of Total Contract Price across the WIP rows"],
    ["Costs to Date (WIP report)", money(costs), contract ? `${(costs / contract * 100).toFixed(0)}% of contract` : "", wipSrc, "sum of the WIP master's Costs to Date column - the report-date cut, not live QuickBooks"],
    ["Billed to Date", money(billed), contract ? `${(billed / contract * 100).toFixed(0)}% of contract` : "", wipSrc, "sum of Billed to Date (gross, incl. retainage)"],
    ["Left to Bill", money(left), "", wipSrc, "sum of Left to Bill = contract - billed"],
    ["Net Over/(Under)", money(net), net >= 0 ? "overbilled" : "underbilled", wipSrc, "overbillings minus underbillings"],
    ["Active Jobs", String(active), `of ${rows.length}`, wipSrc, "STATUS = Active on the Test tabs (blank = MFD, active by construction)"],
  ];
  const row = $("#kpiRow"); row.innerHTML = "";
  for (const [label, value, sub, src, how] of cards) {
    const el = document.createElement("div"); el.className = "kpi";
    el.innerHTML = `<div class="k-label"></div><div class="k-value"></div><div class="k-sub"></div>`;
    el.querySelector(".k-label").textContent = label;
    el.querySelector(".k-value").textContent = value;
    el.querySelector(".k-sub").textContent = sub;
    el.appendChild(srcChip(src, how));
    row.appendChild(el);
  }
}
// ── provenance chips: "<system> · <as-of>" under a figure, the formula in words on hover ──
function srcText(system, when, kind) {
  if (!when) return system;
  return `${system} · ${kind === "report" ? "report " + fmtDate(when) : "loaded " + fmtDate(when, true)}`;
}
function srcChip(text, how) {
  const s = document.createElement("div"); s.className = "k-src"; s.textContent = text || ""; if (how) s.title = how; return s;
}
// Grey shimmer lines where content is about to land (owner 2026-09-01: no blank cards while loading).
function skeletonInto(host, n) {
  if (!host) return;
  const isTbody = host.tagName === "TBODY";
  host.innerHTML = "";
  for (let k = 0; k < (n || 5); k++) {
    if (isTbody) { const tr = document.createElement("tr"); const td = document.createElement("td"); td.colSpan = 12; td.className = "left"; const s = document.createElement("div"); s.className = "skel"; s.style.width = (55 + ((k * 17) % 40)) + "%"; td.appendChild(s); tr.appendChild(td); host.appendChild(tr); }
    else { const s = document.createElement("div"); s.className = "skel"; s.style.width = (55 + ((k * 17) % 40)) + "%"; host.appendChild(s); }
  }
}
function loadedAt(feed) { return ((meta.freshness || {}).ledger || {})[feed] || null; }

function renderDivisions() {
  const groups = {};
  for (const r of ALL) {
    const d = r.division || "—";
    const g = groups[d] || (groups[d] = { jobs: 0, contract: 0, costs: 0, billed: 0, over: 0, under: 0 });
    g.jobs++; g.contract += num(r.total_contract_price); g.costs += num(r.costs_to_date);
    g.billed += num(r.billed_to_date); g.over += num(r.overbillings); g.under += num(r.underbillings);
  }
  const cols = ["Division", "Jobs", "Contract", "Costs", "Billed", "Over", "Under"];
  const thead = $("#divTable thead"), tbody = $("#divTable tbody");
  thead.innerHTML = ""; tbody.innerHTML = "";
  const htr = document.createElement("tr");
  cols.forEach((c, i) => { const th = document.createElement("th"); th.textContent = c; if (i === 0) th.className = "left"; htr.appendChild(th); });
  thead.appendChild(htr);
  const order = Object.keys(groups).sort((a, b) => groups[b].contract - groups[a].contract);
  for (const d of order) {
    const g = groups[d];
    const tr = document.createElement("tr");
    tr.style.cursor = "pointer";
    tr.title = "Show this division's active projects";
    tr.onclick = (e) => { if (!e.target.closest(".cell")) drillDivision(d); };
    [textCell(d, true), textCell(String(g.jobs)), moneyCell(g.contract), moneyCell(g.costs),
     moneyCell(g.billed), moneyCell(g.over), moneyCell(g.under)].forEach((c, i) => {
      const td = document.createElement("td"); if (i === 0) td.className = "left"; td.appendChild(c); tr.appendChild(td);
    });
    tbody.appendChild(tr);
  }
}
// Big-picture → zoom: click a division rollup row → its active projects.
function drillDivision(div) {
  $("#fDivision").value = div;
  $("#fActive").checked = true;
  renderProjects();
  $("#widget-projects").scrollIntoView({ behavior: "smooth", block: "start" });
  toast(`${div} — active projects`);
}
const num = v => (isNum(v) ? v : Number(v) || 0);

function renderProjects() {
  const rows = filtered();
  $("#projCount").textContent = `(${rows.length})`;
  const cols = visibleColumns();
  const thead = $("#projTable thead"), tbody = $("#projTable tbody");
  thead.innerHTML = ""; tbody.innerHTML = "";
  // header
  const htr = document.createElement("tr");
  for (const c of cols) {
    const th = document.createElement("th");
    if (c.align === "left") th.className = "left";
    th.textContent = c.label;
    if (sortKey === c.key) { const a = document.createElement("span"); a.className = "arrow"; a.textContent = sortDir === 1 ? " ▲" : " ▼"; th.appendChild(a); }
    th.onclick = () => { if (sortKey === c.key) sortDir = -sortDir; else { sortKey = c.key; sortDir = c.type === "text" || c.type === "status" ? 1 : -1; } renderProjects(); };
    htr.appendChild(th);
  }
  thead.appendChild(htr);
  // body
  for (const r of rows) {
    const tr = document.createElement("tr");
    tr.onclick = (e) => { if (!e.target.closest(".cell")) openDetail(r); };
    for (const c of cols) {
      const td = document.createElement("td");
      if (c.align === "left") td.className = "left";
      td.appendChild(cellFor(c, r[c.key]));
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }
}

// ── cell builders ─────────────────────────────────────────────────────────
function cellFor(col, value) {
  if (col.type === "status") return statusPill(value);
  if (col.key === "percent_complete") return pctBar(col, value);
  const span = document.createElement("span");
  span.className = "cell";
  span.textContent = fmt(col, value);
  const n = Number(value);
  const hasNum = value !== null && value !== "" && !Number.isNaN(n);
  if (col.type === "money" && hasNum && n < 0) span.classList.add("neg");
  if (col.key === "pure_job_borrow" && hasNum && n > 0) span.classList.add("neg");
  if (col.key === "underbillings" && hasNum && n > 0) span.classList.add("pos");
  if (col.key === "budget_burn" && hasNum && n > 1) span.classList.add("neg");
  if (col.key === "qbo_margin_pct" && hasNum && n < 0.05) span.classList.add("neg");
  span.title = "Click to copy";
  span.onclick = (e) => { e.stopPropagation(); copy(String(raw(col, value))); };
  return span;
}
function pctBar(col, value) {
  const wrap = document.createElement("span");
  wrap.className = "cell bar pct-bar";
  const n = Number(value);
  const has = value !== null && value !== "" && !Number.isNaN(n);
  const fill = document.createElement("span");
  fill.className = "bar-fill" + (has && n > 1 ? " over" : "");
  fill.style.width = (has ? Math.max(0, Math.min(100, n * 100)) : 0) + "%";
  const txt = document.createElement("span"); txt.className = "bar-txt"; txt.textContent = fmt(col, value);
  wrap.appendChild(fill); wrap.appendChild(txt);
  wrap.title = "Click to copy";
  wrap.onclick = (e) => { e.stopPropagation(); copy(String(raw(col, value))); };
  return wrap;
}
function statusPill(v) {
  const s = document.createElement("span");
  s.className = "pill " + (v || "").toLowerCase();
  s.textContent = v || "—";
  return s;
}
function textCell(v, left) { const s = document.createElement("span"); s.textContent = v; const w = document.createElement("span"); w.appendChild(s); return w; }
function moneyCell(v) { const s = document.createElement("span"); s.className = "cell" + (num(v) < 0 ? " neg" : ""); s.textContent = money(v); s.onclick = () => copy(String(Math.round(num(v)))); s.title = "Click to copy"; return s; }
function addRow(tbody, cells) { const tr = document.createElement("tr"); cells.forEach((c, i) => { const td = document.createElement("td"); if (i === 0) td.className = "left"; td.appendChild(c); tr.appendChild(td); }); tbody.appendChild(tr); }

// ── Detail panel ──────────────────────────────────────────────────────────
const DETAIL_GROUPS = [
  ["Identity", [["division", "Division", "text"], ["project_type", "Type", "text"], ["builder_or_gc", "Builder / GC", "text"], ["rp_category", "Category", "text"], ["status", "Status", "text"], ["report_date", "Report date", "text"]]],
  ["Contract", [["original_contract", "Original contract", "money"], ["approved_cos", "Approved COs", "money"], ["total_contract_price", "Total contract price", "money"]]],
  ["Budget", [["original_estimated_cost", "Original estimated cost", "money"], ["co_costs", "CO costs", "money"], ["estimated_total_costs", "Estimated total costs (ETC)", "money"], ["original_profit", "Original profit", "money"], ["gross_profit_pct", "Gross profit %", "pct"]]],
  ["Costs", [["costs_to_date", "Costs to date", "money"], ["cost_to_complete", "Cost to complete", "money"], ["percent_complete", "Percent complete", "pct"]]],
  ["Earned", [["revenues_earned_to_date", "Revenues earned", "money"], ["profit_earned_to_date", "Profit earned", "money"]]],
  ["Billing", [["billed_to_date", "Billed to date", "money"], ["overbillings", "Overbillings", "money"], ["underbillings", "Underbillings", "money"], ["retainage_held", "Retainage held", "money"], ["left_to_bill", "Left to bill", "money"], ["future_profit_to_earn", "Future profit to earn", "money"], ["pure_job_borrow", "Pure job borrow", "money"]]],
  ["Cross-checks", [["mark_schedule", "Schedule", "text"], ["mark_general_list", "General list", "text"], ["mark_jobtread", "JobTread", "text"]]],
];
let detailRow = null;

// P&L (project-pnl) link — shows when the workbook was last pulled, opens it, and
// (on an explicit confirm) runs project-pnl to (re)generate it. The generate call is
// the ONLY place the dashboard triggers a QBO pull + a file write; it is gated by a
// confirm dialog here and a `confirm` flag the server also requires.
function buildPnlGroup(proj) {
  const g = document.createElement("div"); g.className = "dgroup";
  const h = document.createElement("h4"); h.textContent = "P&L"; g.appendChild(h);
  if (!(_pp && _pp.pn === proj && $("#recordView") && !$("#recordView").hidden)) {   // not when already ON the project page
    const b = document.createElement("button"); b.className = "btn small primary"; b.textContent = "Open project page"; b.style.marginBottom = "8px";
    b.onclick = (e) => { e.stopPropagation(); closePanels(); openProjectPage(proj); }; g.appendChild(b); }

  // ── live computed P&L (folded in from the spine; reconciles with project-pnl) ──
  const pl = document.createElement("div"); pl.className = "pnl-live"; pl.textContent = "computing…"; g.appendChild(pl);
  fetch(`/api/pnl/pl?proj=${encodeURIComponent(proj)}`).then(r => r.json()).then(d => {
    pl.innerHTML = "";
    if (d.error) { pl.textContent = d.error; return; }
    const rowP = (k, v, cls) => {
      const r = document.createElement("div"); r.className = "drow" + (cls ? " " + cls : "");
      const a = document.createElement("span"); a.className = "dk"; a.textContent = k;
      const b = document.createElement("span"); b.className = "dv"; b.textContent = v;
      r.appendChild(a); r.appendChild(b); pl.appendChild(r); return r;
    };
    if (!d.has_wip) rowP("Revenue basis", "no WIP snapshot", "pnl-sub");
    rowP("Contract", money(d.contract));
    rowP("% complete", ((d.pct_complete || 0) * 100).toFixed(1) + "%");
    rowP("Earned revenue", money(d.earned));
    rowP("Costs to date", money(d.cost));
    rowP(`Overhead (${d.overhead_basis})`, "(" + money(d.overhead) + ")");
    const nr = rowP("Net margin", `${money(d.net)} · ${d.net_pct == null ? "—" : (d.net_pct * 100).toFixed(1) + "%"}`, "pnl-net");
    nr.classList.add(d.net >= 0 ? "pos" : "neg");
    rowP("Billed to GC (AR)", money(d.billed), "pnl-sub");
    // The make-up of billed-to-date: every AR invoice (draw) the project has, paid or open
    // (owner 2026-08-21: "I need to see all the invoices the project has"). Oldest first.
    if (d.invoices && d.invoices.length) {
      const cap = document.createElement("div"); cap.className = "pnl-cap"; cap.textContent = `Invoices - all draws (${d.invoices.length})`; pl.appendChild(cap);
      const tbl = document.createElement("div"); tbl.className = "pnl-invoices";
      for (const iv of d.invoices) {
        const r = document.createElement("div"); r.className = "pnl-inv";
        const dt = document.createElement("span"); dt.className = "pi-date"; dt.textContent = fmtDateShort(iv.txn_date);
        const no = document.createElement("span"); no.className = "pi-no";
        if (iv.doc_number && iv.qbo_txn_id) { const a = document.createElement("a"); a.href = qboInvoiceUrl(iv.qbo_txn_id); a.target = "_blank"; a.rel = "noopener"; a.className = "qbo-link"; a.textContent = iv.doc_number; a.title = "Open this invoice in QuickBooks"; no.appendChild(a); }
        else no.textContent = iv.doc_number || "–";
        const am = document.createElement("span"); am.className = "pi-amt"; am.appendChild(moneyCell(iv.amount));
        const paid = num(iv.balance) <= 0;
        const st = document.createElement("span"); st.className = "pi-st";
        st.appendChild(stText(paid ? "Paid" : "Open", paid ? "st-ok" : "st-warn",
          paid ? (iv.paid_date ? "GC paid " + fmtDateShort(iv.paid_date) : "GC paid") : ("Open AR balance " + money(iv.balance))));
        r.appendChild(dt); r.appendChild(no); r.appendChild(am); r.appendChild(st); tbl.appendChild(r);
      }
      pl.appendChild(tbl);
    }
    if (d.by_code && d.by_code.length) {
      const cap = document.createElement("div"); cap.className = "pnl-cap"; cap.textContent = "Costs by code"; pl.appendChild(cap);
      const tbl = document.createElement("div"); tbl.className = "pnl-codes";
      for (const c of d.by_code.slice(0, 10)) {
        const r = document.createElement("div"); r.className = "pnl-code" + (c.code === "(uncoded)" ? " uncoded" : "");
        const nm = document.createElement("span"); nm.className = "pc-code"; nm.textContent = c.code + (c.is_sub ? " · sub" : "");
        const am = document.createElement("span"); am.className = "pc-amt"; am.textContent = money(c.amount);
        r.appendChild(nm); r.appendChild(am); tbl.appendChild(r);
      }
      pl.appendChild(tbl);
    }
  }).catch(() => { pl.textContent = "P&L unavailable."; });

  // ── source job folder (Synology CP/RP · OneDrive MFD) — the owner's "source link" ──
  const src = document.createElement("div"); src.className = "pnl-actions";
  const jobBtn = document.createElement("button"); jobBtn.className = "btn small"; jobBtn.textContent = "Open job folder ↗";
  jobBtn.title = "Open this job's folder on the file server (docs · takeoffs · photos)";
  jobBtn.onclick = () => fetch(`/api/job/open?proj=${encodeURIComponent(proj)}`, { method: "POST" })
    .then(r => r.json()).then(x => toast(x.error ? x.error : "Opening job folder…"));
  src.appendChild(jobBtn); g.appendChild(src);

  // ── detailed export (project-pnl Excel) — open / generate ──
  const cap2 = document.createElement("div"); cap2.className = "pnl-cap"; cap2.textContent = "Detailed export (project-pnl)"; g.appendChild(cap2);
  const row = document.createElement("div"); row.className = "drow";
  const dk = document.createElement("span"); dk.className = "dk"; dk.textContent = "Last pulled";
  const dv = document.createElement("span"); dv.className = "dv"; dv.textContent = "checking…";
  row.appendChild(dk); row.appendChild(dv); g.appendChild(row);
  const acts = document.createElement("div"); acts.className = "pnl-actions";
  const openBtn = document.createElement("button"); openBtn.className = "btn small"; openBtn.textContent = "Open Excel"; openBtn.disabled = true;
  const genBtn = document.createElement("button"); genBtn.className = "btn small"; genBtn.textContent = "Generate / Refresh";
  acts.appendChild(openBtn); acts.appendChild(genBtn); g.appendChild(acts);
  const msg = document.createElement("div"); msg.className = "pnl-msg"; g.appendChild(msg);

  const refresh = () => fetch(`/api/pnl?proj=${encodeURIComponent(proj)}`).then(r => r.json()).then(d => {
    if (d.error) { dv.textContent = "—"; return; }
    if (d.exists) {
      dv.textContent = `${timeAgo(d.mtime)} · ${fmtDate(d.mtime, true)}`;
      openBtn.disabled = false;
      openBtn.onclick = () => fetch(`/api/pnl/open?proj=${encodeURIComponent(proj)}`, { method: "POST" })
        .then(r => r.json()).then(x => toast(x.error ? x.error : "Opening P&L…"));
    } else {
      dv.textContent = "not generated yet"; openBtn.disabled = true;
    }
    msg.textContent = d.note || "";
  }).catch(() => { dv.textContent = "unavailable"; });

  genBtn.onclick = () => {
    if (!confirm(`Generate the P&L for ${proj}?\n\nThis runs project-pnl against QBO — a Touch ID prompt will appear on this Mac — and can take a minute or two.`)) return;
    genBtn.disabled = true; genBtn.textContent = "Generating…";
    msg.textContent = "Running project-pnl — watch for the Touch ID prompt on this Mac.";
    fetch(`/api/pnl/generate`, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ proj, confirm: true }) })
      .then(r => r.json()).then(d => {
        if (d.error) { msg.textContent = "Error: " + d.error; genBtn.disabled = false; genBtn.textContent = "Generate / Refresh"; return; }
        pollPnl(proj, genBtn, msg, refresh);
      }).catch(e => { msg.textContent = "Error: " + e; genBtn.disabled = false; genBtn.textContent = "Generate / Refresh"; });
  };
  refresh();
  return g;
}

function pollPnl(proj, genBtn, msg, refresh) {
  const finish = (t) => { msg.textContent = t; genBtn.disabled = false; genBtn.textContent = "Generate / Refresh"; refresh(); };
  const tick = () => fetch(`/api/pnl/status?proj=${encodeURIComponent(proj)}`).then(r => r.json()).then(s => {
    if (s.state === "running") {
      const where = s.status ? " · " + s.status : " · Touch ID may be waiting";
      msg.textContent = `Generating… (${s.elapsed || 0}s)${where}`;
      msg.title = s.status || "";
      setTimeout(tick, 1500);
    }
    else if (s.state === "done") { finish("Done — P&L refreshed."); }
    else if (s.state === "error") { msg.textContent = "Failed: " + (s.detail || "see the log"); genBtn.disabled = false; genBtn.textContent = "Generate / Refresh"; }
    else { finish(""); }
  }).catch(() => setTimeout(tick, 3000));
  setTimeout(tick, 1500);
}

function openDetail(r) {
  detailRow = r;
  $("#detailTitle").textContent = `${r.project_no} — ${r.project_name || ""}`;
  $("#detailSub").textContent = `${r.division || ""}${r.source_tab ? " · " + r.source_tab : ""}`;
  const body = $("#detailBody"); body.innerHTML = "";
  { const g = document.createElement("div"); g.className = "dgroup"; const b = document.createElement("button"); b.className = "btn primary"; b.textContent = "Open project page";
    b.title = "Everything about this job in one place - how it's doing, how we get funded, bills, the trail"; b.onclick = () => { closePanels(); openProjectPage(r.project_no); }; g.appendChild(b); body.appendChild(g); }
  const typ = k => ({ money: "money", pct: "pct" }[k] ? { type: k } : { type: "text" });
  for (const [title, fields] of DETAIL_GROUPS) {
    const rows = fields.filter(([k]) => r[k] !== null && r[k] !== undefined && r[k] !== "");
    if (!rows.length) continue;
    const g = document.createElement("div"); g.className = "dgroup";
    const h = document.createElement("h4"); h.textContent = title === "Identity" ? title : `${title}  ·  WIP master, report ${fmtDate(r.report_date)}`; g.appendChild(h);
    for (const [k, label, type] of rows) {
      const row = document.createElement("div"); row.className = "drow";
      const dk = document.createElement("span"); dk.className = "dk"; dk.textContent = label;
      const dv = document.createElement("span"); dv.className = "dv";
      dv.textContent = k === "report_date" ? fmtDate(r[k]) : fmt({ type }, r[k]);
      dv.title = "Click to copy";
      dv.onclick = () => copy(String(raw({ type }, r[k])));
      row.appendChild(dk); row.appendChild(dv); g.appendChild(row);
    }
    body.appendChild(g);
  }
  body.appendChild(buildPnlGroup(r.project_no));
  if (window.openTrail) {   // the money trail (trail.js): every QBO line behind Costs / Billed, with the running total
    const g = document.createElement("div"); g.className = "dgroup"; const b = document.createElement("button"); b.className = "btn";
    b.textContent = "Show every dollar"; b.title = "Every QBO line behind Costs to date and Billed to date, with a running total against the budget";
    b.onclick = () => { closePanels(); openTrail(r.project_no); }; g.appendChild(b); body.appendChild(g);
  }
  const ap = AP.by_project && AP.by_project[r.project_no];
  if (ap) {
    const g = document.createElement("div"); g.className = "dgroup";
    const h = document.createElement("h4"); h.textContent = `AP / Liens  ·  Bill Tracker${((meta.freshness || {}).sources || {})["sync-ap"] ? ", " + fmtDate(((meta.freshness || {}).sources || {})["sync-ap"], true) : ""} (excludes subs)`; g.appendChild(h);
    for (const [label, val] of [["Open AP balance", money(ap.open_balance)], ["Open bills", String(ap.open_lines)]]) {
      const row = document.createElement("div"); row.className = "drow";
      const dk = document.createElement("span"); dk.className = "dk"; dk.textContent = label;
      const dv = document.createElement("span"); dv.className = "dv"; dv.textContent = val;
      dv.title = "Click to copy"; dv.onclick = () => copy(val);
      row.appendChild(dk); row.appendChild(dv); g.appendChild(row);
    }
    body.appendChild(g);
  }
  const cbp = COST.by_project_code && COST.by_project_code[r.project_no];
  const cload = COST.by_project && COST.by_project[r.project_no];
  if (cbp && cbp.length) {
    const g = document.createElement("div"); g.className = "dgroup";
    const h = document.createElement("h4"); h.textContent = `Costs by code  ·  QuickBooks${loadedAt("Costs (QBO)") ? ", loaded " + fmtDate(loadedAt("Costs (QBO)"), true) : ""}`; g.appendChild(h);
    const summary = [];
    if (cload) {
      summary.push(["Total loaded", money(cload.costs_loaded)]);
      if (cload.sub_costs) summary.push(["of which subs", money(cload.sub_costs)]);
      if (r.costs_to_date != null) {
        summary.push(["WIP master costs to date (report " + fmtDate(r.report_date) + ")", money(r.costs_to_date)]);
        // The reconciliation the loader prints to the terminal, on screen: match within 5% or a red gap.
        const gap = num(cload.costs_loaded) - num(r.costs_to_date), tol = Math.abs(num(r.costs_to_date)) * 0.05;
        summary.push([Math.abs(gap) <= tol ? "Difference (match, within 5%)" : "Difference (QuickBooks minus WIP)", (gap >= 0 ? "+" : "") + money(gap), Math.abs(gap) > tol]);
      }
    }
    for (const [label, val, bad] of summary) {
      const row = document.createElement("div"); row.className = "drow";
      const dk = document.createElement("span"); dk.className = "dk"; dk.textContent = label;
      const dv = document.createElement("span"); dv.className = "dv" + (bad ? " neg" : ""); dv.textContent = val; dv.title = "Click to copy"; dv.onclick = () => copy(val);
      row.appendChild(dk); row.appendChild(dv); g.appendChild(row);
    }
    for (const c of cbp) {
      const row = document.createElement("div"); row.className = "drow";
      const dk = document.createElement("span"); dk.className = "dk";
      const chip = document.createElement("span"); chip.className = "codechip"; chip.textContent = c.code;
      dk.appendChild(chip); if (c.lines > 1) dk.appendChild(document.createTextNode(` ·${c.lines}`));
      const dv = document.createElement("span"); dv.className = "dv"; dv.textContent = money(c.actual);
      dv.title = "Click to copy"; dv.onclick = () => copy(String(Math.round(c.actual || 0)));
      row.appendChild(dk); row.appendChild(dv); g.appendChild(row);
    }
    body.appendChild(g);
  }
  if (r.costs_loaded != null) {
    const g = document.createElement("div"); g.className = "dgroup";
    const h = document.createElement("h4"); h.textContent = "Margin  ·  computed here from the WIP contract / ETC and QuickBooks costs"; g.appendChild(h);
    const mrows = [
      ["Planned markup ((contract − ETC) ÷ ETC)", pct(r.markup_pct)],   // label now matches the arithmetic (was "contract ÷ ETC")
      ["Planned margin (GP ÷ contract)", pct(r.margin_pct)],
      ["Budget burn (cost ÷ ETC)", pct(r.budget_burn)],
      ["Actual markup (billed ÷ QBO cost)", pct(r.actual_markup_pct)],
      ["Margin to date (billed − cost)", money(r.qbo_margin)],
      ["Margin % of billed", pct(r.qbo_margin_pct)],
      ["Subs share of cost", pct(r.subs_pct)],
    ];
    for (const [label, val] of mrows) {
      const row = document.createElement("div"); row.className = "drow";
      const dk = document.createElement("span"); dk.className = "dk"; dk.textContent = label;
      const dv = document.createElement("span"); dv.className = "dv"; dv.textContent = val; dv.title = "Click to copy"; dv.onclick = () => copy(val);
      row.appendChild(dk); row.appendChild(dv); g.appendChild(row);
    }
    body.appendChild(g);
  }
  if (r.notes) {
    const g = document.createElement("div"); g.className = "dgroup";
    const h = document.createElement("h4"); h.textContent = "Notes"; g.appendChild(h);
    const n = document.createElement("div"); n.className = "dnote"; n.textContent = r.notes; g.appendChild(n);
    body.appendChild(g);
  }
  openPanel("#detail");
}
function detailAsText() {
  if (!detailRow) return "";
  const r = detailRow; const lines = [`${r.project_no} — ${r.project_name || ""}`];
  for (const [title, fields] of DETAIL_GROUPS) {
    const present = fields.filter(([k]) => r[k] !== null && r[k] !== undefined && r[k] !== "");
    if (!present.length) continue;
    lines.push("", title.toUpperCase());
    for (const [k, label, type] of present) lines.push(`  ${label}: ${k === "report_date" ? fmtDate(r[k]) : fmt({ type }, r[k])}`);
  }
  if (r.notes) lines.push("", "NOTES", "  " + r.notes);
  return lines.join("\n");
}

// ── Panels ────────────────────────────────────────────────────────────────
function openPanel(sel) { $("#overlay").hidden = false; $(sel).hidden = false; }
function closePanels() { $("#overlay").hidden = true; $("#detail").hidden = true; $("#settings").hidden = true;
  { const bd = $("#billDetail"); if (bd) bd.hidden = true; } { const sd = $("#sublocDetail"); if (sd) sd.hidden = true; }
  { const pb = $("#payBills"); if (pb) pb.hidden = true; } { const lr = $("#lienReview"); if (lr) lr.hidden = true; }
  { const st = $("#invStatement"); if (st) st.hidden = true; } { const iv = $("#invDetail"); if (iv) iv.hidden = true; }
  { const vd = $("#vendorDetail"); if (vd) vd.hidden = true; } }

// Full-page record view (app-style, like JobTread) - takes over the main content area instead of a
// narrow side slide-over, so wide detail has room to read (owner 2026-08-28: "side view squishes too
// much"). Opening hides the tab-pages; Back restores the tab you came from (activeTab is unchanged).
function openRecord(title, sub) {
  $$(".tab-page").forEach(p => { p.hidden = true; });
  $("#recordView").hidden = false;
  $("#recordTitle").textContent = title || "";
  $("#recordSub").textContent = sub || "";
  window.scrollTo(0, 0);
}
function closeRecord() {
  const rv = $("#recordView"); if (rv) rv.hidden = true;
  $$(".tab-page").forEach(p => { p.hidden = p.dataset.tab !== activeTab; });
  window.scrollTo(0, 0);
}

// ── Copy + CSV + toast ────────────────────────────────────────────────────
let toastTimer = null;
function toast(msg) {
  const t = $("#toast"); t.textContent = msg; t.hidden = false;
  clearTimeout(toastTimer); toastTimer = setTimeout(() => (t.hidden = true), 1400);
}
async function copy(text) {
  try { await navigator.clipboard.writeText(text); toast("Copied: " + text.slice(0, 40)); }
  catch { const ta = document.createElement("textarea"); ta.value = text; document.body.appendChild(ta); ta.select(); document.execCommand("copy"); ta.remove(); toast("Copied"); }
}
// Export CSV = the table you are looking at, with its filters (owner 2026-09-02: "export the current
// visible table respecting filters"). Each tab names its rows + columns; anything else falls back to
// the Overview projects table.
function _csvTable() {
  const nz = v => v == null ? "" : v;
  if (activeTab === "invoices") {
    const rows = _invRows();
    return { name: "invoices", rows, cols: [
      ["Client", i => i.customer], ["Project", i => i.project_no], ["Invoice #", i => i.doc_number], ["Date", i => i.txn_date],
      ["Due", i => i.due_date], ["Days past due", i => i.days_past_due], ["Memo", i => i.memo], ["Open balance", i => oiBal(i)],
      ["Invoice total", i => num(i.amount)], ["Status", i => i.status], ["Last action", i => i.last_action_date],
      ["Next follow-up", i => i.next_followup], ["Collections note", i => i.note], ["Lien", i => i.lien_status],
      ["Notice deadline", i => i.lien_due_label], ["Notion page", i => i.notion_url]].map(([l, g]) => [l, r => nz(g(r))]) };
  }
  if (activeTab === "bills") {
    const view = billView(), f = billFilterValues();
    const rows = (BILLS || []).filter(b => view.pred(b) && billPassesFilters(b, f));
    return { name: "bills", rows, cols: [
      ["Vendor", b => b.vendor], ["Project", b => b.project_no], ["Division", b => b.division], ["Bill #", b => b.bill_ref],
      ["Bill date", b => b.bill_date], ["This line", b => num(b.line_amount)], ["Bill total", b => num(b.bill_total)],
      ["Open balance", b => num(b.open_balance)], ["Pay status", b => b.pay_status], ["Paid", b => b.pay_date],
      ["Invoice", b => b.invoice_no], ["Invoice status", b => b.invoice_status], ["GC paid", b => b.gc_paid_date],
      ["Approved", b => b.approved], ["Lien", b => b.lien_status], ["Description", b => b.description]].map(([l, g]) => [l, r => nz(g(r))]) };
  }
  if (activeTab === "wip") {
    const activeOnly = $("#wipActive") ? $("#wipActive").checked : true;
    const rows = (ALL || []).filter(r => !activeOnly || ["", "active"].includes((r.status || "").toLowerCase()));
    return { name: "wip", rows, cols: WIP_COLS.map(c => [c.label, r => nz(r[c.k])]) };
  }
  const cols = visibleColumns();
  return { name: "projects", rows: filtered(), cols: cols.map(c => [c.label, r => raw(c, r[c.key])]) };
}
// Export dialog (owner 2026-09-02): Excel (a grouped report that keeps the state colours) or CSV, and
// "how would you like it grouped" - defaults to the grouping on screen (Bills: the Group by select).
const EXPORT_GROUPS = {
  bills: [["", "None"], ["Vendor", "Vendor"], ["Project", "Project"], ["Client", "Client"], ["Division", "Division"], ["Invoice", "Draw / invoice"]],
  invoices: [["", "None"], ["Client", "Client"], ["Project", "Project"], ["Status", "Status"]],
  wip: [["", "None"], ["Division", "Division"]],
  projects: [["", "None"], ["Division", "Division"], ["Client", "Client"]],
};
function exportCSV() {
  const spec = _csvTable();
  const groups = EXPORT_GROUPS[spec.name] || [["", "None"]];
  let def = "";
  if (spec.name === "bills") { const g = ($("#billGroup") || {}).value; def = { vendor: "Vendor", project_no: "Project", client: "Client", division: "Division", matched_invoice: "Invoice" }[g] || ""; }
  if (spec.name === "invoices") def = "Client";
  const ov = document.createElement("div"); ov.className = "xdlg-ov";
  ov.innerHTML = `<div class="xdlg" role="dialog"><h3>Export ${_ge(spec.rows.length.toLocaleString())} ${_ge(spec.name)} rows</h3>
    <label>Format <select id="xdFmt"><option value="xlsx">Excel (.xlsx) - keeps the colour coding, groups collapse</option><option value="csv">CSV - plain rows</option></select></label>
    <label>Group by <select id="xdGrp">${groups.map(([v, l]) => `<option value="${_ge(v)}"${v === def ? " selected" : ""}>${_ge(l)}</option>`).join("")}</select></label>
    <div class="xdlg-actions"><button class="btn" id="xdCancel">Cancel</button><button class="btn primary" id="xdGo">Export</button></div></div>`;
  document.body.appendChild(ov);
  const close = () => ov.remove();
  ov.onclick = (e) => { if (e.target === ov) close(); };
  $("#xdCancel").onclick = close;
  $("#xdGo").onclick = () => { const fmt = $("#xdFmt").value, grp = $("#xdGrp").value; close(); if (fmt === "xlsx") exportXlsx(spec, grp); else exportCsvPlain(spec, grp); };
}
function _stateFmt(name, row, label) {   // the on-screen colour for a cell, as a state class for Excel
  if (name === "bills") {
    if (label === "Open balance" && num(row.open_balance) > 0) return "neg";
    if (label === "Paid" && row.pay_date) return "pos";
    if (label === "Lien" && BILL_LIEN_RISK.has(row.lien_status)) return "warn";
    if (label === "Approved" && row.approved && row.approved !== "approved") return "warn";
    if (label === "Invoice status" && row.invoice_status === "Invoice paid") return "pos";
  }
  if (name === "invoices") {
    if (label === "Open balance" && oiBal(row) > 0) return "neg";
    if (label === "Days past due" && num(row.days_past_due) > 0) return "neg";
    if (label === "Status" && (row.status || "").toLowerCase() === "paid") return "pos";
    if (label === "Lien" && row.lien_status) return "warn";
  }
  return null;
}
async function exportXlsx(spec, grp) {
  const { name, rows, cols } = spec;
  const moneyLabels = new Set(["This line", "Bill total", "Open balance", "Invoice total", "Amount", ...WIP_COLS.filter(c => c.t === "money").map(c => c.label)]);
  const columns = cols.map(([l]) => ({ label: l, type: moneyLabels.has(l) ? "money" : "text" }));
  const gi = grp ? cols.findIndex(([l]) => l === grp) : -1;
  const data = rows.map(r => cols.map(([, g]) => g(r)));
  const fmt = [];
  rows.forEach((r, ri) => cols.forEach(([l], ci) => { const c = _stateFmt(name, r, l); if (c) fmt.push({ r: ri, c: ci, cls: c }); }));
  const filt = name === "bills" ? [(billView().name), billDate && billDate.active() ? "date " + billDate.label() : "", ...BILL_MSEL.map(c => (billMSel[c.id] || {}).size ? [...billMSel[c.id]].join("/") : "").filter(Boolean)].filter(Boolean).join(" · ")
             : name === "invoices" ? [invScope === "all" ? "all invoices" : "open invoices", invQuick ? `find "${invQuick}"` : ""].filter(Boolean).join(" · ") : "";
  const body = { name: `${name === "bills" ? "Bill Tracker" : name === "invoices" ? "Invoices" : name} ${grp ? "by " + grp : ""}`.trim(),
    sheet: name, title: `${name === "bills" ? "Bill Tracker" : name === "invoices" ? "Open invoices" : "Ledger"}${filt ? " - " + filt : ""}`,
    subtitle: `${rows.length.toLocaleString()} rows as shown · exported ${fmtDate(new Date(Date.now() - new Date().getTimezoneOffset() * 60000).toISOString().slice(0, 19), true)}${grp ? " · grouped by " + grp : ""}`,   // local time, not UTC
    columns, rows: data, group_by: gi >= 0 ? gi : null, fmt };
  toast("Building the Excel report…");
  try {
    const r = await (await fetch("/api/export/xlsx", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) })).json();
    if (r && r.ok) toast(`Excel report saved to Downloads (${r.rows} rows) - opened in Finder`); else toast("Export failed: " + ((r && r.error) || "unknown"));
  } catch (e) { toast("Export failed: " + e); }
}
function exportCsvPlain(spec, grp) {
  const { name, rows, cols } = spec;
  const esc = v => { const s = String(v ?? ""); return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; };
  const head = cols.map(([l]) => esc(l)).join(",");
  const gi = grp ? cols.findIndex(([l]) => l === grp) : -1;
  let ordered = rows;
  if (gi >= 0) {   // grouped CSV = sorted by the group (stable) with a subtotal line per group
    const key = r => String(cols[gi][1](r) ?? "");
    const seen = new Map(); rows.forEach((r, i) => { const k = key(r); if (!seen.has(k)) seen.set(k, i); });
    ordered = [...rows].sort((a, b) => seen.get(key(a)) - seen.get(key(b)));
  }
  const moneyLabels = new Set(["This line", "Bill total", "Open balance", "Invoice total"]);
  const lines = []; let cur = null, subs = null;
  const flush = () => { if (cur == null) return; lines.push(cols.map(([l], ci) => ci === gi ? esc(cur + " total") : (moneyLabels.has(l) ? esc(Math.round((subs[ci] || 0) * 100) / 100) : "")).join(",")); };
  for (const r of ordered) {
    if (gi >= 0) { const k = String(cols[gi][1](r) ?? ""); if (k !== cur) { flush(); cur = k; subs = {}; } cols.forEach(([l], ci) => { if (moneyLabels.has(l)) subs[ci] = (subs[ci] || 0) + (num(cols[ci][1](r)) || 0); }); }
    lines.push(cols.map(([, g]) => esc(g(r))).join(","));
  }
  flush();
  const body = lines.join("\n");
  const blob = new Blob(["\ufeff" + head + "\n" + body], { type: "text/csv" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `ledger_${name}_${new Date().toISOString().slice(0, 10)}.csv`;
  a.click(); URL.revokeObjectURL(a.href);
  toast(`Exported ${rows.length} ${name} rows (as filtered)`);
}

// ── Settings UI ───────────────────────────────────────────────────────────
function syncSettingsUI() {
  $("#setTheme").value = settings.theme;
  $("#setAccent").value = settings.accent;
  $("#setFont").value = settings.font;
  $("#setFontSize").value = settings.fontSize;
  $("#fsVal").textContent = settings.fontSize + "px";
  $("#setDensity").value = settings.density;
  $("#setWidth").value = settings.width;
  $("#wKpis").checked = settings.widgets.kpis;
  $("#wAttention").checked = settings.widgets.attention;
  $("#wCosts").checked = settings.widgets.costs;
  $("#wMargins").checked = settings.widgets.margins;
  $("#wDivisions").checked = settings.widgets.divisions;
  $("#wProjects").checked = settings.widgets.projects;
  const cc = $("#colChooser"); cc.innerHTML = "";
  for (const c of COLUMNS) {
    const lab = document.createElement("label");
    const cb = document.createElement("input"); cb.type = "checkbox";
    cb.checked = c.always || settings.columns.includes(c.key);
    cb.disabled = !!c.always;
    cb.onchange = () => {
      const set = new Set(settings.columns);
      cb.checked ? set.add(c.key) : set.delete(c.key);
      settings.columns = COLUMNS.filter(x => set.has(x.key) || x.always).map(x => x.key);
      saveSettings(); renderProjects();
    };
    lab.appendChild(cb); lab.appendChild(document.createTextNode(" " + c.label));
    cc.appendChild(lab);
  }
}
function wireSettings() {
  const on = (sel, ev, fn) => $(sel).addEventListener(ev, fn);
  on("#setTheme", "change", e => { settings.theme = e.target.value; saveSettings(); applySettings(); });
  on("#setAccent", "input", e => { settings.accent = e.target.value; saveSettings(); applySettings(); });
  on("#setFont", "change", e => { settings.font = e.target.value; saveSettings(); applySettings(); });
  on("#setFontSize", "input", e => { settings.fontSize = +e.target.value; $("#fsVal").textContent = settings.fontSize + "px"; saveSettings(); applySettings(); });
  on("#setDensity", "change", e => { settings.density = e.target.value; saveSettings(); applySettings(); });
  on("#setWidth", "change", e => { settings.width = e.target.value; saveSettings(); applySettings(); });
  on("#wKpis", "change", e => { settings.widgets.kpis = e.target.checked; saveSettings(); applySettings(); });
  on("#wAttention", "change", e => { settings.widgets.attention = e.target.checked; saveSettings(); applySettings(); });
  on("#wCosts", "change", e => { settings.widgets.costs = e.target.checked; saveSettings(); applySettings(); });
  on("#wMargins", "change", e => { settings.widgets.margins = e.target.checked; saveSettings(); applySettings(); });
  on("#wDivisions", "change", e => { settings.widgets.divisions = e.target.checked; saveSettings(); applySettings(); });
  on("#wProjects", "change", e => { settings.widgets.projects = e.target.checked; saveSettings(); applySettings(); });
  on("#btnReset", "click", () => { settings = baseDefaults(); saveSettings(); applySettings(); syncSettingsUI(); render(); toast("Reset to your default"); });
  on("#btnSetDefault", "click", () => { localStorage.setItem(LS_DEF, JSON.stringify(settings)); toast("Saved as your default view"); });
}

// ── Wire up ───────────────────────────────────────────────────────────────
// == Excel-style cell selection + running sum (click / drag / Cmd+click / Shift+click) ==
// Select number cells across any table; a status bar (bottom-right) shows Sum / Count /
// Avg, like Excel. Number cells only (dates, %, labels are skipped). Non-number cells
// keep their normal click (open the row) and clear the selection.
const _cs = { cells: new Set(), anchor: null, dragging: false, swallow: false, bar: null, sumText: "" };

function _csNum(td) {
  let t = (td.textContent || "").trim();
  t = t.replace(/\s+\d+\/\d+$/, "");                     // drop a trailing "2/7" paid-count
  if (!t || /[a-z]/i.test(t) || /%/.test(t)) return null; // letters (dates/labels) or percent -> not summable
  const c = t.replace(/[$,\s]/g, "").replace(/^\((.*)\)$/, "-$1");   // ($123) -> -123
  return /^-?\d+(\.\d+)?$/.test(c) ? parseFloat(c) : null;
}
function _csApply(list) {
  _cs.cells.forEach(td => td.classList.remove("cell-sel"));
  _cs.cells = new Set(list);
  _cs.cells.forEach(td => td.classList.add("cell-sel"));
}
function _csToggle(td) { if (_cs.cells.has(td)) { _cs.cells.delete(td); td.classList.remove("cell-sel"); } else { _cs.cells.add(td); td.classList.add("cell-sel"); } }
function _csClear() { _cs.cells.forEach(td => td.classList.remove("cell-sel")); _cs.cells.clear(); _cs.anchor = null; _csUpdate(); }

function _csRange(a, b) {                                  // rectangular range within one table
  const table = a.closest("table");
  if (b.closest("table") !== table) return [b];
  const rows = [...table.querySelectorAll("tr")];
  const ra = rows.indexOf(a.parentElement), rb = rows.indexOf(b.parentElement);
  const c0 = Math.min(a.cellIndex, b.cellIndex), c1 = Math.max(a.cellIndex, b.cellIndex);
  const out = [];
  for (let r = Math.min(ra, rb); r <= Math.max(ra, rb); r++) {
    const cells = rows[r] ? rows[r].children : [];
    for (let c = c0; c <= c1; c++) if (cells[c]) out.push(cells[c]);
  }
  return out;
}

function _csUpdate() {
  if (!_cs.bar) return;
  [..._cs.cells].forEach(td => { if (!td.isConnected) { _cs.cells.delete(td); } });   // drop cells lost to a re-render
  const nums = [], money = [];
  _cs.cells.forEach(td => { const n = _csNum(td); if (n !== null) { nums.push(n); money.push(/\$/.test(td.textContent || "")); } });
  if (!nums.length) { _cs.bar.hidden = true; _cs.sumText = ""; return; }
  const sum = nums.reduce((t, n) => t + n, 0), isMoney = money.every(Boolean);
  const fmt = n => (isMoney ? "$" : "") + n.toLocaleString(undefined, { maximumFractionDigits: 2 });
  _cs.sumText = fmt(sum);
  _cs.bar.hidden = false;
  _cs.bar.querySelector(".sb-sum").textContent = _cs.sumText;
  _cs.bar.querySelector(".sb-meta").textContent = `Count ${nums.length} · Avg ${fmt(sum / nums.length)}`;
}

function initCellSelect() {
  const bar = document.createElement("div"); bar.className = "sumbar"; bar.hidden = true;
  bar.innerHTML = '<span class="sb-label">Σ</span><span class="sb-sum"></span>'
    + '<span class="sb-meta"></span><button class="sb-copy" title="Copy the sum">Copy</button>'
    + '<button class="sb-x" title="Clear (Esc)">✕</button>';
  document.body.appendChild(bar); _cs.bar = bar;
  bar.querySelector(".sb-x").onclick = _csClear;
  bar.querySelector(".sb-copy").onclick = () => { if (_cs.sumText && navigator.clipboard) { navigator.clipboard.writeText(_cs.sumText); toast("Copied " + _cs.sumText); } };

  const cellAt = e => { const td = e.target.closest("table.grid td"); return (td && !e.target.closest("a,button,input,label,select")) ? td : null; };

  document.addEventListener("mousedown", e => {
    const td = cellAt(e); if (!td) return;
    const mod = e.metaKey || e.ctrlKey;
    if (!mod && !e.shiftKey && _csNum(td) === null) return;   // plain click on a non-number -> leave it (row can open)
    e.preventDefault();
    if (e.shiftKey && _cs.anchor) _csApply(_csRange(_cs.anchor, td));
    else if (mod) { _csToggle(td); _cs.anchor = td; }
    else { _csApply([td]); _cs.anchor = td; _cs.dragging = true; }
    _cs.swallow = true; _csUpdate();
  }, true);

  document.addEventListener("mouseover", e => {
    if (!_cs.dragging || !_cs.anchor) return;
    const td = e.target.closest("table.grid td"); if (!td) return;
    _csApply(_csRange(_cs.anchor, td)); _csUpdate();
  }, true);

  document.addEventListener("mouseup", () => { _cs.dragging = false; }, true);

  document.addEventListener("click", e => {
    if (_cs.swallow) { _cs.swallow = false; e.stopPropagation(); e.preventDefault(); return; }   // a selecting click must not also open a row
    if (_cs.cells.size && !e.target.closest(".sumbar")) _csClear();                              // click-away clears
  }, true);

  document.addEventListener("keydown", e => {
    if (e.key === "Escape" && _cs.cells.size) _csClear();
    else if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === "c" && _cs.sumText && navigator.clipboard && _cs.cells.size) navigator.clipboard.writeText(_cs.sumText);
  });
}


// ── Systems & processes ─────────────────────────────────────────────────────
// The registry lives in the vault as markdown (02_processes/*.md); the server
// parses it per request, so this is a live view of those files, not a copy.
// Read-only by design - the vault owns the truth, we only render it.
// Replaced the daily markdown digest (the owner, 2026-08-19).
let REG = null;             // cached /api/processes payload
let sysDomain = null;       // domain code currently filtered to (null = all)

const HEALTH_LABEL = { red: "Broken", yellow: "Fragile", green: "Running", none: "Not started" };

// ── Health tab: the company-health metric layer ─────────────────────────────
// Sections come PREFORMATTED from /api/healthtab (one model, server-side - the
// fold-in of the retired Company Tracker). This side only draws them: hero
// cards, metric rows (click -> jump to the tab that holds the detail), aging /
// division bars, the Recurring & Debt register, and the break-even audit trail.
let HEALTH = null;

async function loadHealth(force) {
  const box = $("#healthSections"); if (!box) return;
  if (HEALTH && HEALTH.ok && !force) { renderHealth(); return; }
  const note = $("#healthNote");
  if (note && !HEALTH) { note.textContent = "loading…"; skeletonInto($("#healthSections"), 8); }
  try { HEALTH = await (await fetch("/api/healthtab")).json(); }
  catch (e) { if (note) note.textContent = "could not load"; return; }
  renderHealth();
}

function renderHealth() {
  const box = $("#healthSections"), note = $("#healthNote");
  if (!box || !HEALTH) return;
  if (!HEALTH.ok) {
    if (note) note.textContent = "unavailable";
    box.innerHTML = ""; box.appendChild(el2("p", "hint", HEALTH.error || "The ledger is not ready."));
    return;
  }
  if (note) note.textContent = HEALTH.as_of_label
    ? `QBO metrics as of ${HEALTH.as_of_label}` : "QBO metrics not pulled yet — derived rows only";
  box.innerHTML = "";
  for (const sec of HEALTH.sections || []) {
    const w = el2("section", "widget health-sec tone-" + (sec.tone || "n"));
    const head = el2("div", "widget-head"); head.appendChild(el2("h2", null, sec.title)); w.appendChild(head);
    if (sec.heroes && sec.heroes.length) {
      const kr = el2("div", "kpi-row");
      for (const [lab, val, cls] of sec.heroes) {
        const k = el2("div", "kpi");
        k.appendChild(el2("div", "k-label", lab));
        k.appendChild(el2("div", "k-value hv-" + (cls || "n"), val));
        kr.appendChild(k);
      }
      w.appendChild(kr);
    }
    if (sec.rows && sec.rows.length) {
      const t = el2("table", "grid health-grid"), tb = el2("tbody");
      for (const [metric, val, detail, cls, target] of sec.rows) {
        const tr = el2("tr", target ? "h-click" : null);
        const tdm = el2("td", "left h-metric", metric);
        const tdv = el2("td", "h-val hv-" + (cls || "n"), val);
        const tdd = el2("td", "left h-detail", detail || "");
        if (target) {
          tr.title = "Open " + (TAB_LABELS[target] || target);
          tr.onclick = () => setTab(target);
          tdd.appendChild(el2("span", "h-jump", " ↗"));
        }
        tr.appendChild(tdm); tr.appendChild(tdv); tr.appendChild(tdd);
        tb.appendChild(tr);
      }
      t.appendChild(tb);
      const sc = el2("div", "table-scroll"); sc.appendChild(t); w.appendChild(sc);
    }
    for (const [header, segs] of sec.bars || []) {
      if (segs && segs.length) w.appendChild(healthBar(header, segs));
    }
    if (sec.note) w.appendChild(el2("p", "hint", sec.note));
    box.appendChild(w);
  }
  renderHealthRecurring();
  renderHealthAudit();
}

// One proportional bar + legend. Class tokens (bk0..bk4 age, dv-* division)
// come from the server; colour encodes age or division, never decoration.
function healthBar(header, segs) {
  const total = segs.reduce((s, x) => s + (x[1] || 0), 0) || 1;
  const wrap = el2("div", "hbar-wrap");
  wrap.appendChild(el2("div", "hbar-head", header));
  const bar = el2("div", "hbar");
  for (const [label, value, cls] of segs) {
    const seg = el2("span", "hseg " + (cls || ""));
    seg.style.width = Math.max(1.5, value / total * 100) + "%";
    seg.title = `${label}: ${money(value)}`;
    bar.appendChild(seg);
  }
  wrap.appendChild(bar);
  const leg = el2("div", "hbar-legend");
  for (const [label, value, cls, sub] of segs) {
    const li = el2("span", "hleg");
    li.appendChild(el2("span", "hdot " + (cls || "")));
    li.appendChild(el2("span", null, `${label} ${money(value)}${sub ? " · " + sub : ""}`));
    leg.appendChild(li);
  }
  wrap.appendChild(leg);
  return wrap;
}

function renderHealthRecurring() {
  const widget = $("#healthRecWidget"); if (!widget) return;
  const rec = HEALTH && HEALTH.recurring;
  widget.hidden = !rec;
  if (!rec) return;
  const kpis = $("#healthRecKpis"); kpis.innerHTML = "";
  const tile = (lab, val, sub) => {
    const k = el2("div", "kpi");
    k.appendChild(el2("div", "k-label", lab));
    k.appendChild(el2("div", "k-value", val));
    if (sub) k.appendChild(el2("div", "k-sub", sub));
    kpis.appendChild(k);
  };
  tile("Fixed overhead / month", money(rec.fixed_overhead_month), "last full month");
  tile("Debt service / month", money(rec.debt_service_month), "liability balance drops");
  tile("Total monthly obligation", money(rec.total_monthly_obligation), "the nut to cover");
  const alerts = rec.alerts || [];
  $("#healthRecNote").textContent = alerts.length ? `${alerts.length} to review` : "all steady";
  const abox = $("#healthRecAlerts"); abox.innerHTML = "";
  for (const a of alerts) {
    const row = el2("div", "hrec-alert");
    row.appendChild(el2("span", "acct-pill " + (a.status === "STOPPED" ? "neg" : a.status === "NEW" ? "info" : "warn"), a.status));
    row.appendChild(el2("span", "hrec-name", a.name));
    row.appendChild(el2("span", "hrec-note", a.note || ""));
    abox.appendChild(row);
  }
  const thead = $("#healthRecTable thead"), tbody = $("#healthRecTable tbody");
  thead.innerHTML = ""; tbody.innerHTML = "";
  const htr = el2("tr");
  ["Obligation", "Kind", "Last paid", "Last amount", "Prior", "Status", "YTD"].forEach((c, i) => {
    const th = el2("th", i === 0 || i === 1 || i === 2 || i === 5 ? "left" : null, c); htr.appendChild(th);
  });
  thead.appendChild(htr);
  const band = label => {
    const tr = el2("tr", "bill-group"); const td = el2("td", "left", label);
    td.colSpan = 7; tr.appendChild(td); tbody.appendChild(tr);
  };
  const row = (r, kind) => {
    const tr = el2("tr");
    tr.appendChild(el2("td", "left", r.name));
    tr.appendChild(el2("td", "left", kind));
    tr.appendChild(el2("td", "left", r.last_month || "–"));
    tr.appendChild(el2("td", "h-num", money(r.last_amount)));
    tr.appendChild(el2("td", "h-num", r.prior_amount ? money(r.prior_amount) : "–"));
    tr.appendChild(el2("td", "left " + (r.status === "STOPPED" ? "hv-r" : r.status === "steady" ? "" : "hv-a"),
                       r.status === "steady" ? "steady" : `${r.status}${r.note ? " · " + r.note : ""}`));
    tr.appendChild(el2("td", "h-num", money(r.ytd)));
    tbody.appendChild(tr);
  };
  band("Overhead (P&L expense accounts)");
  for (const r of rec.overhead || []) row(r, r.kind);
  band("Debt service (liability balance drops — MCA rows shown, never counted)");
  for (const r of rec.debt || []) row(r, r.refinancing ? "MCA (excluded)" : "debt");
}

function renderHealthAudit() {
  const widget = $("#healthAuditWidget"); if (!widget) return;
  const rows = (HEALTH && HEALTH.be_audit) || [];
  widget.hidden = !rows.length;
  if (!rows.length) return;
  const thead = $("#healthAuditTable thead"), tbody = $("#healthAuditTable tbody");
  thead.innerHTML = ""; tbody.innerHTML = "";
  const htr = el2("tr");
  ["Item", "Value", "Where it came from"].forEach((c, i) => htr.appendChild(el2("th", i === 1 ? null : "left", c)));
  thead.appendChild(htr);
  for (const [item, value, src] of rows) {
    if (String(item).startsWith("──")) {   // divider row; the CAVEAT divider carries its text in src
      const label = String(item).replace(/─/g, "").trim();
      const tr = el2("tr", "bill-group");
      const td = el2("td", "left", src ? `${label} - ${src}` : label);
      td.colSpan = 3; tr.appendChild(td); tbody.appendChild(tr); continue;
    }
    const tr = el2("tr");
    tr.appendChild(el2("td", "left", item));
    tr.appendChild(el2("td", "h-num", value));
    tr.appendChild(el2("td", "left h-detail", src));
    tbody.appendChild(tr);
  }
}

async function loadSystems(force) {
  if (REG && !force) { renderSystems(); return; }
  const note = $("#sysNote");
  if (note) note.textContent = "(reading the vault…)";
  try { REG = await (await fetch("/api/processes")).json(); }
  catch { REG = { ok: false, rows: [], domains: [], error: "could not reach the ledger server" }; }
  renderSystems();
}

function sysFiltered() {
  const rows = (REG && REG.rows) || [];
  const q = ($("#sysSearch") ? $("#sysSearch").value : "").trim().toLowerCase();
  const owner = $("#sysOwner") ? $("#sysOwner").value : "";
  const health = $("#sysHealth") ? $("#sysHealth").value : "";
  const state = $("#sysState") ? $("#sysState").value : "";
  const life = $("#sysLife") ? $("#sysLife").value : "";
  const showRetired = $("#sysRetired") ? $("#sysRetired").checked : false;
  return rows.filter(r => {
    if (r.retired && !showRetired) return false;
    if (sysDomain && r.domain_code !== sysDomain) return false;
    if (owner && r.owner !== owner) return false;
    if (health && r.health_key !== health) return false;
    if (state && r.state_kind !== state) return false;
    if (life && r.life_key !== life) return false;
    if (q) {
      const hay = [r.id, r.process, r.owner, r.touchers, r.record, r.automation, r.cadence, r.domain]
        .join(" ").toLowerCase();
      if (!hay.includes(q)) return false;
    }
    return true;
  });
}

function renderSystems() {
  const note = $("#sysNote"), tb = $("#sysTable tbody"), th = $("#sysTable thead");
  if (!tb || !th) return;
  if (REG && REG.ok === false) {
    note.textContent = "(unavailable)";
    th.innerHTML = ""; tb.innerHTML = "";
    const tr = document.createElement("tr"), td = document.createElement("td");
    td.className = "left"; td.colSpan = 8;
    td.textContent = REG.error || "The process registry could not be read.";
    tr.appendChild(td); tb.appendChild(tr);
    return;
  }
  if (!REG) return;

  // Owner list is rebuilt from the payload, preserving the current pick.
  const sel = $("#sysOwner");
  if (sel && sel.options.length <= 1) {
    for (const o of (REG.owners || [])) {
      const opt = document.createElement("option");
      opt.value = o.owner; opt.textContent = `${o.owner} (${o.count})`;
      sel.appendChild(opt);
    }
  }

  const c = REG.counts || { health: {}, state: {}, life: {} };
  const h = c.health || {}, st = c.state || {}, lf = c.life || {};
  const stalled = (lf.agreed || 0) + (lf.building || 0);
  const stats = [
    ["Processes", String(c.active || 0), `${c.retired || 0} retired`],
    ["Broken", String(h.red || 0), "running and going wrong"],
    ["Fragile", String(h.yellow || 0), "one person or one step from failing"],
    ["Running clean", String(h.green || 0), "automated or reliable"],
    ["Unconfirmed", String((st.proposed || 0) + (st.inferred || 0)),
      `${st.proposed || 0} proposed · ${st.inferred || 0} inferred`],
    ["Agreed, not live", String(stalled), "decided but never built"],
  ];
  const sr = $("#sysStats"); sr.innerHTML = "";
  for (const [label, value, sub] of stats) {
    const el = document.createElement("div"); el.className = "kpi";
    el.innerHTML = `<div class="k-label"></div><div class="k-value"></div><div class="k-sub"></div>`;
    el.querySelector(".k-label").textContent = label;
    el.querySelector(".k-value").textContent = value;
    el.querySelector(".k-sub").textContent = sub;
    sr.appendChild(el);
  }

  // Domain chips - one per registry file, plus All.
  const dom = $("#sysDomains"); dom.innerHTML = "";
  const mk = (code, label, count) => {
    const b = document.createElement("button");
    b.className = "sys-chip" + ((sysDomain === code) ? " active" : "");
    b.innerHTML = `<span class="sys-chip-label"></span><span class="sys-chip-count"></span>`;
    b.querySelector(".sys-chip-label").textContent = label;
    b.querySelector(".sys-chip-count").textContent = String(count);
    b.onclick = () => { sysDomain = code; renderSystems(); };
    dom.appendChild(b);
  };
  const live = (REG.rows || []).filter(r => !r.retired);
  mk(null, "All", live.length);
  for (const d of (REG.domains || [])) {
    mk(d.code, d.code || d.title, d.rows.filter(r => !r.retired).length);
  }

  const rows = sysFiltered();
  // Denominator follows the retired toggle, so the count never reads "81 of 73".
  const universe = ($("#sysRetired") && $("#sysRetired").checked) ? (c.total || 0) : (c.active || 0);
  note.textContent = `(${rows.length} of ${universe}${REG.source ? " · live from the vault" : ""})`;

  const cols = ["", "ID", "Process", "Owner", "Also touches", "Record", "Automation", "Cadence", "State"];
  th.innerHTML = ""; tb.innerHTML = "";
  const htr = document.createElement("tr");
  for (const c2 of cols) {
    const el = document.createElement("th");
    el.className = "left"; el.textContent = c2; htr.appendChild(el);
  }
  th.appendChild(htr);

  if (!rows.length) {
    const tr = document.createElement("tr"), td = document.createElement("td");
    td.className = "left"; td.colSpan = cols.length; td.textContent = "No processes match these filters.";
    tr.appendChild(td); tb.appendChild(tr); return;
  }

  let lastDomain = null;
  for (const r of rows) {
    if (r.domain !== lastDomain) {          // a section header per domain
      lastDomain = r.domain;
      const gr = document.createElement("tr"); gr.className = "sys-group";
      const gd = document.createElement("td"); gd.className = "left"; gd.colSpan = cols.length;
      gd.textContent = `${r.domain_code} · ${r.domain}`;
      gr.appendChild(gd); tb.appendChild(gr);
    }
    const tr = document.createElement("tr");
    if (r.retired) tr.classList.add("sys-retired");

    const hd = document.createElement("td");
    const dot = document.createElement("span");
    dot.className = "sys-dot " + (r.health_key || "none");
    dot.title = HEALTH_LABEL[r.health_key] || "";
    hd.appendChild(dot); tr.appendChild(hd);

    const idc = document.createElement("td"); idc.className = "left sys-id";
    idc.textContent = r.id; tr.appendChild(idc);

    const pc = document.createElement("td"); pc.className = "left sys-process";
    pc.textContent = r.process;
    if (r.life_key && r.life_key !== "live" && !r.retired) {
      const tag = document.createElement("span");
      tag.className = "sys-life " + r.life_key;
      tag.textContent = r.life_key;
      tag.title = "Decided, but not running yet";
      pc.appendChild(tag);
    }
    tr.appendChild(pc);

    for (const k of ["owner", "touchers", "record", "automation", "cadence"]) {
      tr.appendChild(leftText(r[k] || ""));
    }

    const sc = document.createElement("td"); sc.className = "left";
    const pill = document.createElement("span");
    pill.className = "sys-state " + (r.state_kind || "unknown");
    pill.textContent = r.state_kind === "confirmed" && r.confirmed_on
      ? `confirmed ${fmtDateShort(r.confirmed_on)}` : (r.state_kind || "");
    if (r.confirmed_on) pill.title = "Confirmed by the owner on " + fmtDate(r.confirmed_on);
    sc.appendChild(pill); tr.appendChild(sc);

    tb.appendChild(tr);
  }
}

// ── Graph tab: the org knowledge graph + imported system diagrams ─────────────
// A self-contained canvas graph, no libraries. One renderer, two layouts:
//   org map  -> force-directed (Obsidian-style): nodes = vault notes, edges = [[links]]
//   diagrams -> layered flow, imported from docs/ARCHITECTURE.md mermaid (arrows = data flow)
// Data is fetched once from /api/graph (parsed live server-side from the vault + docs).
let GRAPH = null;            // cached /api/graph payload
let graphMode = "org";       // "org" or a diagram key
let GV = null;               // live view state for the current mode
let _graphRAF = 0;
let _gDraw = true;           // redraw-needed flag: paint only on change or while the sim moves (idle-cheap)
const _gmark = () => { _gDraw = true; };

const GRAPH_GROUPS = ["hub", "01_company", "02_processes", "03_systems", "04_integrations", "05_tools", "tasks"];
const GRAPH_GROUP_LABEL = {
  hub: "Hubs", "01_company": "Company & people", "02_processes": "Processes",
  "03_systems": "Systems", "04_integrations": "Integrations", "05_tools": "Tools", tasks: "Tasks",
};
const _ge = s => String(s == null ? "" : s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const _graphFont = () => getComputedStyle(document.body).fontFamily || "system-ui, sans-serif";

function _pal() {
  // Theme-aware colours, read fresh each draw so the dark-mode toggle just works.
  const cs = getComputedStyle(document.documentElement);
  const v = (n, d) => (cs.getPropertyValue(n).trim() || d);
  return {
    bg: v("--graph-bg", v("--surface", "#ffffff")),
    edge: v("--graph-edge", "#c8d0da"),
    text: v("--text", "#1f2937"),
    dim: v("--text-dim", "#8a97a6"),
    stroke: v("--graph-node-stroke", "#ffffff"),
    hi: v("--accent", "#3b82f6"),
    box: v("--graph-box", "#eef1f5"),
    boxStroke: v("--border", "#d5dbe2"),
    c: [1, 2, 3, 4, 5, 6, 7, 8].map(i => v("--graph-c" + i, "#8aa0b6")),
  };
}
const _groupColor = (pal, group) => pal.c[(GRAPH_GROUPS.indexOf(group) + pal.c.length) % pal.c.length];
const _graphCanvas = () => document.getElementById("graphCanvas");

function _graphSize() {
  const cv = _graphCanvas(); if (!cv) return null;
  const wrap = cv.parentElement, dpr = window.devicePixelRatio || 1;
  const w = wrap.clientWidth, h = wrap.clientHeight;
  if (!w || !h) return null;                          // tab hidden -> size later
  if (cv._cw !== w || cv._ch !== h || cv._dpr !== dpr) {
    cv.width = Math.round(w * dpr); cv.height = Math.round(h * dpr);
    cv.style.width = w + "px"; cv.style.height = h + "px";
    cv._cw = w; cv._ch = h; cv._dpr = dpr;
  }
  return { cv, ctx: cv.getContext("2d"), w, h, dpr };
}

async function loadGraph(force) {
  const note = document.getElementById("graphNote");
  if (GRAPH && !force) { buildGraphModes(); setGraphMode(graphMode); return; }
  if (note) note.textContent = "loading…";
  try { GRAPH = await (await fetch("/api/graph")).json(); }
  catch (e) { if (note) note.textContent = "could not load the graph"; return; }
  buildGraphModes();
  const keys = ["org", ...((GRAPH.diagrams && GRAPH.diagrams.diagrams) || []).map(d => d.key)];
  if (!keys.includes(graphMode)) graphMode = "org";
  setGraphMode(graphMode);
}

function buildGraphModes() {
  const bar = document.getElementById("graphModes"); if (!bar) return;
  bar.innerHTML = "";
  const mk = (mode, label, sub) => {
    const b = document.createElement("button");
    b.className = "graph-mode" + (mode === graphMode ? " active" : "");
    b.dataset.mode = mode;
    b.innerHTML = `<span class="gm-label">${_ge(label)}</span>` + (sub ? `<span class="gm-sub">${_ge(sub)}</span>` : "");
    b.onclick = () => setGraphMode(mode);
    bar.appendChild(b);
  };
  const org = GRAPH.org || { nodes: [], links: [] };
  mk("org", "Org map", `${(org.nodes || []).length} notes`);
  for (const d of ((GRAPH.diagrams && GRAPH.diagrams.diagrams) || [])) {
    // The diagram title leads with a domain code, then a dash and the long description; the
    // short lead word makes a tidy chip, full title on hover. The split set MUST include the
    // em dash: the ARCHITECTURE.md headings are authored with one (same as registry_view).
    const short = (d.title.split(/[-–—(]/)[0] || d.title).trim();
    const b = document.createElement("button");
    b.className = "graph-mode" + (d.key === graphMode ? " active" : "");
    b.dataset.mode = d.key; b.title = d.title;
    b.innerHTML = `<span class="gm-label">${_ge(short)}</span><span class="gm-sub">${d.nodes.length} · ${d.edges.length}</span>`;
    b.onclick = () => setGraphMode(d.key);
    bar.appendChild(b);
  }
}

function setGraphMode(mode) {
  graphMode = mode;
  document.querySelectorAll("#graphModes .graph-mode").forEach(b => b.classList.toggle("active", b.dataset.mode === mode));
  buildGV(mode);
  const note = document.getElementById("graphNote");
  if (note) note.textContent = GV.err ? GV.err : `${GV.nodes.length} nodes · ${GV.links.length} links`;
  buildGraphLegend();
  hideGraphInfo();
  const s = document.getElementById("graphSearch"); if (s) s.value = "";
  layoutGV();
  GV._fitted = false;
  startGraphLoop();
}

function buildGV(mode) {
  if (mode === "org") {
    const src = (GRAPH.org && GRAPH.org.ok) ? GRAPH.org : { nodes: [], links: [] };
    const nodes = src.nodes.map(n => ({
      id: n.id, label: n.label, group: n.group, deg: n.deg || 0,
      r: 4 + Math.sqrt(n.deg || 0) * 1.9, kind: "dot",
    }));
    GV = { layout: "force", directed: false, nodes, links: (src.links || []).slice(),
           err: (GRAPH.org && GRAPH.org.error) || "" };
  } else {
    const d = ((GRAPH.diagrams && GRAPH.diagrams.diagrams) || []).find(x => x.key === mode)
      || { nodes: [], edges: [], direction: "LR" };
    const nodes = d.nodes.map(n => ({ id: n.id, label: n.label, cluster: n.cluster || "", kind: "box" }));
    GV = { layout: "layered", directed: true, dir: d.direction || "LR", nodes,
           links: (d.edges || []).map(e => ({ source: e.source, target: e.target, label: e.label || "" })), err: "" };
  }
  GV.view = { x: 0, y: 0, scale: 1 };
  GV.alpha = GV.layout === "force" ? 1 : 0;
  GV.hover = GV.sel = GV.drag = null; GV.pan = null; GV.hits = null; GV._userMoved = false;
  GV.byId = {}; GV.adj = {};
  GV.nodes.forEach(n => { n.x = n.y = 0; GV.byId[n.id] = n; GV.adj[n.id] = new Set(); });
  GV.links.forEach(l => { if (GV.adj[l.source] && GV.adj[l.target]) { GV.adj[l.source].add(l.target); GV.adj[l.target].add(l.source); } });
  if (GV.layout === "layered") measureBoxes();
}

function _wrapLabel(ctx, text, maxW, maxLines) {
  const words = String(text || "").split(/\s+/).filter(Boolean);
  const lines = []; let cur = "";
  for (const w of words) {
    const t = cur ? cur + " " + w : w;
    if (ctx.measureText(t).width > maxW && cur) { lines.push(cur); cur = w; if (lines.length === maxLines - 1) break; }
    else cur = t;
  }
  if (cur && lines.length < maxLines) lines.push(cur);
  const used = lines.join(" ").split(/\s+/).length;
  if (used < words.length && lines.length) {              // ran out of lines -> ellipsis
    let last = lines[lines.length - 1];
    while (ctx.measureText(last + " …").width > maxW && last.length) last = last.slice(0, -1);
    lines[lines.length - 1] = last + " …";
  }
  return lines.length ? lines : [""];
}

function measureBoxes() {
  const sz = _graphSize(); const ctx = sz ? sz.ctx : _graphCanvas().getContext("2d");
  ctx.font = "600 12px " + _graphFont();
  for (const n of GV.nodes) {
    n.lines = _wrapLabel(ctx, n.label, 150, 3);
    let w = 0; for (const ln of n.lines) w = Math.max(w, ctx.measureText(ln).width);
    n.w = Math.min(184, Math.max(56, w + 22));
    n.h = n.lines.length * 15 + 14;
  }
}

function layoutGV() { if (GV.layout === "layered") layeredLayout(); else forceInit(); }

function forceInit() {
  const n = GV.nodes.length || 1, R = 60 + n * 3.5;
  GV.nodes.forEach((nd, i) => {
    const a = (i / n) * Math.PI * 2 * 1.618;            // golden-angle spiral start
    const rad = R * Math.sqrt((i + 0.5) / n);
    nd.x = Math.cos(a) * rad; nd.y = Math.sin(a) * rad; nd.vx = nd.vy = 0; nd.fixed = false;
  });
  GV.alpha = 1;
}

function forceStep() {
  const nodes = GV.nodes, n = nodes.length; if (!n) return;
  const k = 80, temp = 36 * GV.alpha, grav = 0.034;
  for (const a of nodes) { a.dx = 0; a.dy = 0; }
  for (let i = 0; i < n; i++) {
    const a = nodes[i];
    for (let j = i + 1; j < n; j++) {
      const b = nodes[j];
      let dx = a.x - b.x, dy = a.y - b.y, d = Math.hypot(dx, dy);
      if (d < 0.02) { dx = (i - j) * 0.1 + 0.05; dy = 0.05; d = Math.hypot(dx, dy); }
      const m = (k * k) / d / d, ux = dx * m, uy = dy * m;   // repulsion ~ k^2/d
      a.dx += ux; a.dy += uy; b.dx -= ux; b.dy -= uy;
      const minD = a.r + b.r + 5;                            // soft collision: keep circles apart
      if (d < minD) { const push = (minD - d) / d * 0.4, px = dx * push, py = dy * push;
        a.dx += px; a.dy += py; b.dx -= px; b.dy -= py; }
    }
  }
  for (const l of GV.links) {
    const a = GV.byId[l.source], b = GV.byId[l.target]; if (!a || !b) continue;
    let dx = a.x - b.x, dy = a.y - b.y, d = Math.hypot(dx, dy) || 0.02;
    const m = (d / k), ux = (dx / d) * d * m, uy = (dy / d) * d * m;   // attraction ~ d^2/k
    a.dx -= ux; a.dy -= uy; b.dx += ux; b.dy += uy;
  }
  for (const a of nodes) {
    a.dx -= a.x * grav; a.dy -= a.y * grav;                 // gentle gravity keeps it centred
    if (a === (GV.drag && GV.drag.node)) continue;
    const len = Math.hypot(a.dx, a.dy);
    if (len > 0) { const s = Math.min(len, temp) / len; a.x += a.dx * s; a.y += a.dy * s; }
  }
  GV.alpha *= 0.985;
}

function layeredLayout() {
  const nodes = GV.nodes, byId = GV.byId;
  const out = {}, indeg = {};
  nodes.forEach(n => { out[n.id] = []; indeg[n.id] = 0; });
  for (const l of GV.links) if (byId[l.source] && byId[l.target]) { out[l.source].push(l.target); indeg[l.target]++; }
  // Longest-path ranks via Kahn; nodes stuck in a cycle keep rank 0 (rare in these DAGs).
  const rank = {}, ind = {}; nodes.forEach(n => { rank[n.id] = 0; ind[n.id] = indeg[n.id]; });
  let q = nodes.filter(n => indeg[n.id] === 0).map(n => n.id); const seen = new Set();
  while (q.length) {
    const id = q.shift(); if (seen.has(id)) continue; seen.add(id);
    for (const t of out[id]) { rank[t] = Math.max(rank[t], rank[id] + 1); if (--ind[t] <= 0 && !seen.has(t)) q.push(t); }
  }
  const layers = {}; nodes.forEach(n => (layers[rank[n.id]] ||= []).push(n));
  const ranks = Object.keys(layers).map(Number).sort((a, b) => a - b);
  ranks.forEach((r, i) => layers[r].forEach((n, j) => { n._rk = i; n._ord = j; }));
  // Barycentre sweeps to reduce crossings.
  for (let pass = 0; pass < 6; pass++) {
    for (const r of ranks) {
      const layer = layers[r];
      layer.forEach(n => {
        const nb = [...GV.adj[n.id]].map(id => byId[id]).filter(m => m && m._rk !== n._rk);
        n._bc = nb.length ? nb.reduce((s, m) => s + m._ord, 0) / nb.length : n._ord;
      });
      layer.sort((a, b) => a._bc - b._bc);
      layer.forEach((n, j) => { n._ord = j; });
    }
  }
  const horizontal = /^[LR]/.test(GV.dir || "LR");         // LR/RL lay ranks along x
  const maxW = Math.max(60, ...nodes.map(n => n.w || 60));
  const maxH = Math.max(30, ...nodes.map(n => n.h || 30));
  const gapMain = (horizontal ? maxW : maxH) + 74;
  const gapCross = (horizontal ? maxH : maxW) + 22;
  for (const r of ranks) {
    const layer = layers[r], span = (layer.length - 1) * gapCross;
    layer.forEach((n, j) => {
      const main = n._rk * gapMain, cross = j * gapCross - span / 2;
      if (horizontal) { n.x = main; n.y = cross; } else { n.x = cross; n.y = main; }
    });
  }
}

// Comfortable default framing for the force map: centre on the MEDIAN node and
// scale to the core (85th-percentile radius), so a few flung-out nodes can't shrink
// the whole graph to dots. "Fit" (fitGraph) still frames every node exactly.
function frameGraph() {
  const sz = _graphSize(); if (!sz || !GV || !GV.nodes.length) return false;
  const med = arr => arr.slice().sort((a, b) => a - b)[arr.length >> 1];
  const cx = med(GV.nodes.map(n => n.x)), cy = med(GV.nodes.map(n => n.y));
  const ds = GV.nodes.map(n => Math.hypot(n.x - cx, n.y - cy)).sort((a, b) => a - b);
  const core = ds[Math.floor(ds.length * 0.85)] || ds[ds.length - 1] || 1;
  const scale = Math.max(0.14, Math.min(1.8, 0.42 * Math.min(sz.w, sz.h) / core));
  GV.view.scale = scale; GV.view.x = sz.w / 2 - cx * scale; GV.view.y = sz.h / 2 - cy * scale;
  GV._fitted = true;
  return true;
}

function fitGraph() {
  const sz = _graphSize(); if (!sz || !GV || !GV.nodes.length) return false;
  let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
  for (const n of GV.nodes) {
    const hw = n.kind === "box" ? n.w / 2 : n.r, hh = n.kind === "box" ? n.h / 2 : n.r;
    x0 = Math.min(x0, n.x - hw); y0 = Math.min(y0, n.y - hh);
    x1 = Math.max(x1, n.x + hw); y1 = Math.max(y1, n.y + hh);
  }
  const pad = 46, gw = Math.max(1, x1 - x0), gh = Math.max(1, y1 - y0);
  const scale = Math.min((sz.w - pad * 2) / gw, (sz.h - pad * 2) / gh, 2.2);
  GV.view.scale = Math.max(0.12, scale);
  GV.view.x = sz.w / 2 - ((x0 + x1) / 2) * GV.view.scale;
  GV.view.y = sz.h / 2 - ((y0 + y1) / 2) * GV.view.scale;
  GV._fitted = true;
  return true;
}

function _borderPoint(n, tx, ty) {
  const dx = tx - n.x, dy = ty - n.y, d = Math.hypot(dx, dy) || 1, ux = dx / d, uy = dy / d;
  if (n.kind === "box") {
    const hw = n.w / 2 + 2, hh = n.h / 2 + 2;
    const s = 1 / Math.max(Math.abs(ux) / hw, Math.abs(uy) / hh);
    return [n.x + ux * s, n.y + uy * s];
  }
  return [n.x + ux * (n.r + 1), n.y + uy * (n.r + 1)];
}

function drawGraph() {
  const sz = _graphSize(); if (!sz) return;
  if (!GV._fitted) { if (GV.layout === "force") frameGraph(); else fitGraph(); }
  const { ctx, w, h, dpr } = sz, pal = _pal(), v = GV.view;
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  ctx.fillStyle = pal.bg; ctx.fillRect(0, 0, w, h);
  ctx.translate(v.x, v.y); ctx.scale(v.scale, v.scale);
  const focus = GV.hover || GV.sel, near = focus ? GV.adj[focus.id] : null;
  const hits = GV.hits;
  // edges
  ctx.lineWidth = 1 / v.scale; ctx.lineCap = "round";
  for (const l of GV.links) {
    const a = GV.byId[l.source], b = GV.byId[l.target]; if (!a || !b) continue;
    const on = focus && (l.source === focus.id || l.target === focus.id);
    ctx.globalAlpha = focus && !on ? 0.12 : (GV.directed ? 0.55 : 0.5);
    ctx.strokeStyle = on ? pal.hi : pal.edge;
    const [ax, ay] = GV.directed ? _borderPoint(a, b.x, b.y) : [a.x, a.y];
    const [bx, by] = GV.directed ? _borderPoint(b, a.x, a.y) : [b.x, b.y];
    ctx.beginPath(); ctx.moveTo(ax, ay); ctx.lineTo(bx, by); ctx.stroke();
    if (GV.directed) {
      const ang = Math.atan2(by - ay, bx - ax), s = 8;
      ctx.fillStyle = on ? pal.hi : pal.edge;
      ctx.beginPath();
      ctx.moveTo(bx, by);
      ctx.lineTo(bx - s * Math.cos(ang - 0.4), by - s * Math.sin(ang - 0.4));
      ctx.lineTo(bx - s * Math.cos(ang + 0.4), by - s * Math.sin(ang + 0.4));
      ctx.closePath(); ctx.fill();
    }
  }
  ctx.globalAlpha = 1;
  const showLabels = document.getElementById("graphLabels") && document.getElementById("graphLabels").checked;
  // nodes
  for (const n of GV.nodes) {
    const spotlight = focus === n || (near && near.has(n.id)) || (hits && hits.has(n.id));
    const dim = (focus && n !== focus && !(near && near.has(n.id))) || (hits && !hits.has(n.id));
    ctx.globalAlpha = dim ? 0.22 : 1;
    if (n.kind === "box") _drawBox(ctx, n, pal, showLabels, focus === n);
    else _drawDot(ctx, n, pal, v.scale, showLabels && (spotlight || n.r > 11 || v.scale > 1.15), focus === n);
  }
  ctx.globalAlpha = 1;
}

function _drawDot(ctx, n, pal, scale, label, isFocus) {
  ctx.beginPath(); ctx.arc(n.x, n.y, n.r, 0, Math.PI * 2);
  ctx.fillStyle = _groupColor(pal, n.group);
  ctx.fill();
  ctx.lineWidth = (isFocus ? 2.4 : 1.2) / scale; ctx.strokeStyle = isFocus ? pal.hi : pal.stroke; ctx.stroke();
  if (label) {
    ctx.font = `${isFocus ? "600 " : ""}12px ${_graphFont()}`;
    ctx.fillStyle = isFocus ? pal.text : pal.dim;
    ctx.textAlign = "left"; ctx.textBaseline = "middle";
    ctx.fillText(n.label, n.x + n.r + 4, n.y);
  }
}

function _roundRect(ctx, x, y, w, h, r) {
  ctx.beginPath(); ctx.moveTo(x + r, y);
  ctx.arcTo(x + w, y, x + w, y + h, r); ctx.arcTo(x + w, y + h, x, y + h, r);
  ctx.arcTo(x, y + h, x, y, r); ctx.arcTo(x, y, x + w, y, r); ctx.closePath();
}

function _drawBox(ctx, n, pal, label, isFocus) {
  const x = n.x - n.w / 2, y = n.y - n.h / 2;
  _roundRect(ctx, x, y, n.w, n.h, 7);
  ctx.fillStyle = pal.box; ctx.fill();
  ctx.lineWidth = isFocus ? 2.4 : 1.2; ctx.strokeStyle = isFocus ? pal.hi : pal.boxStroke; ctx.stroke();
  if (label !== false) {
    ctx.font = "600 12px " + _graphFont();
    ctx.fillStyle = pal.text; ctx.textAlign = "center"; ctx.textBaseline = "middle";
    const lines = n.lines || [n.label], y0 = n.y - (lines.length - 1) * 7.5;
    lines.forEach((ln, i) => ctx.fillText(ln, n.x, y0 + i * 15));
  }
}

function buildGraphLegend() {
  const el = document.getElementById("graphLegend"); if (!el) return;
  el.innerHTML = "";
  if (GV.layout === "force") {
    const pal = _pal(), counts = {};
    GV.nodes.forEach(n => { counts[n.group] = (counts[n.group] || 0) + 1; });
    for (const g of GRAPH_GROUPS) {
      if (!counts[g]) continue;
      const row = document.createElement("div"); row.className = "gl-row";
      row.innerHTML = `<span class="gl-dot" style="background:${_groupColor(pal, g)}"></span>` +
        `<span>${_ge(GRAPH_GROUP_LABEL[g] || g)}</span><span class="gl-n">${counts[g]}</span>`;
      el.appendChild(row);
    }
  } else {
    const row = document.createElement("div"); row.className = "gl-row gl-note";
    row.innerHTML = `<span>Arrows show data flow.</span>`;
    el.appendChild(row);
  }
}

function showGraphInfo(n) {
  const el = document.getElementById("graphInfo"); if (!el) return;
  const nb = [...(GV.adj[n.id] || [])].map(id => GV.byId[id]).filter(Boolean)
    .sort((a, b) => (b.deg || 0) - (a.deg || 0));
  const sub = GV.layout === "force"
    ? `${_ge(GRAPH_GROUP_LABEL[n.group] || n.group)} · ${n.deg} link${n.deg === 1 ? "" : "s"}`
    : (n.cluster ? _ge(n.cluster) : `${nb.length} connection${nb.length === 1 ? "" : "s"}`);
  const list = nb.slice(0, 14).map(m => `<li>${_ge(m.label)}</li>`).join("");
  el.innerHTML = `<button class="gi-x" title="Close">×</button>` +
    `<div class="gi-title">${_ge(n.label)}</div><div class="gi-sub">${sub}</div>` +
    (list ? `<div class="gi-h">Connected to</div><ul class="gi-list">${list}</ul>` : "") +
    (nb.length > 14 ? `<div class="gi-more">+${nb.length - 14} more</div>` : "");
  el.querySelector(".gi-x").onclick = () => { GV.sel = null; hideGraphInfo(); _gmark(); };
  el.hidden = false;
}
function hideGraphInfo() { const el = document.getElementById("graphInfo"); if (el) el.hidden = true; }

function _graphEventWorld(e) {
  const cv = _graphCanvas(), rect = cv.getBoundingClientRect();
  const cx = e.clientX - rect.left, cy = e.clientY - rect.top;
  return { cx, cy, wx: (cx - GV.view.x) / GV.view.scale, wy: (cy - GV.view.y) / GV.view.scale };
}
function _pickNode(wx, wy) {
  // topmost-ish: iterate in reverse so later-drawn wins ties
  for (let i = GV.nodes.length - 1; i >= 0; i--) {
    const n = GV.nodes[i];
    if (n.kind === "box") { if (Math.abs(wx - n.x) <= n.w / 2 && Math.abs(wy - n.y) <= n.h / 2) return n; }
    else if (Math.hypot(wx - n.x, wy - n.y) <= n.r + 4) return n;
  }
  return null;
}

function startGraphLoop() { stopGraphLoop(); _gDraw = true; const tick = () => {
  if (!GV || activeTab !== "graph") { _graphRAF = 0; return; }
  if (GV.layout === "force" && GV.alpha > 0.02) { forceStep(); if (!GV._userMoved) frameGraph(); _gDraw = true; }  // keep the view framed as the layout settles
  if (_gDraw) { drawGraph(); _gDraw = false; }   // idle when nothing moves; the tab-switch stops the loop entirely
  _graphRAF = requestAnimationFrame(tick);
}; _graphRAF = requestAnimationFrame(tick); }
function stopGraphLoop() { if (_graphRAF) cancelAnimationFrame(_graphRAF); _graphRAF = 0; }

function wireGraphCanvas() {
  const cv = _graphCanvas(); if (!cv) return;
  cv.addEventListener("wheel", e => {
    e.preventDefault(); if (!GV) return; GV._userMoved = true; _gmark();
    const { cx, cy } = _graphEventWorld(e);
    const f = Math.exp(-e.deltaY * 0.0016), ns = Math.min(3.2, Math.max(0.12, GV.view.scale * f));
    const k = ns / GV.view.scale;
    GV.view.x = cx - (cx - GV.view.x) * k; GV.view.y = cy - (cy - GV.view.y) * k; GV.view.scale = ns;
  }, { passive: false });
  cv.addEventListener("pointerdown", e => {
    if (!GV) return; GV._userMoved = true; _gmark(); cv.setPointerCapture(e.pointerId);
    const { cx, cy, wx, wy } = _graphEventWorld(e);
    const n = _pickNode(wx, wy);
    if (n) { GV.drag = { node: n, moved: false }; n.fixed = true; if (GV.layout === "force") GV.alpha = Math.max(GV.alpha, 0.4); }
    else GV.pan = { x: cx, y: cy, vx: GV.view.x, vy: GV.view.y, moved: false };
  });
  cv.addEventListener("pointermove", e => {
    if (!GV) return; _gmark();
    const { cx, cy, wx, wy } = _graphEventWorld(e);
    if (GV.drag) { GV.drag.node.x = wx; GV.drag.node.y = wy; GV.drag.moved = true; if (GV.layout === "force") GV.alpha = Math.max(GV.alpha, 0.3); return; }
    if (GV.pan) { GV.view.x = GV.pan.vx + (cx - GV.pan.x); GV.view.y = GV.pan.vy + (cy - GV.pan.y); GV.pan.moved = true; cv.style.cursor = "grabbing"; return; }
    const n = _pickNode(wx, wy);
    GV.hover = n; cv.style.cursor = n ? "pointer" : "grab";
  });
  const end = e => {
    if (!GV) return; _gmark();
    if (GV.drag) { const nd = GV.drag.node; if (GV.layout === "force") nd.fixed = false; if (!GV.drag.moved) { GV.sel = nd; showGraphInfo(nd); } GV.drag = null; }
    else if (GV.pan) { if (!GV.pan.moved) { GV.sel = null; hideGraphInfo(); } GV.pan = null; cv.style.cursor = "grab"; }
  };
  cv.addEventListener("pointerup", end);
  cv.addEventListener("pointerleave", () => { if (GV && !GV.drag && !GV.pan) { GV.hover = null; _gmark(); } });
}

function graphSearch(q) {
  if (!GV) return; q = (q || "").trim().toLowerCase(); _gmark();
  if (!q) { GV.hits = null; return; }
  GV.hits = new Set(GV.nodes.filter(n => (n.label || "").toLowerCase().includes(q) || (n.id || "").toLowerCase().includes(q)).map(n => n.id));
  const first = GV.nodes.find(n => GV.hits.has(n.id));
  if (first) { GV.sel = first; showGraphInfo(first); GV._userMoved = true; const sz = _graphSize(); if (sz) { GV.view.x = sz.w / 2 - first.x * GV.view.scale; GV.view.y = sz.h / 2 - first.y * GV.view.scale; } }
}

// ── WIP Review: pending WIP update as before/after → approve/disapprove → write ─
// Compute runs the WIP pipeline (3 QBO pulls) and diffs each Test tab; the owner
// approves QBO facts and answers on the PM fields; Sync writes the approved values
// to Test - CP / Test - RP / Test-Master. All read-only until Sync. State is
// wrDecisions[PN][field] = approved(bool) for CHANGED fields; wrDrop = rejected ADDs.
let WR = null;
let wrDecisions = {};
let wrDrop = new Set();
let wrPoll = null;

const WR_DIV_ORDER = ["Commercial", "Residential", "Multi-Family"];
const wrChanged = c => c.filter(f => f.changed);
function wrDelta(f) {
  const a = f.was == null ? null : Number(f.was), b = f.now == null ? null : Number(f.now);
  if (a == null && b == null) return null;
  if (a == null) return { txt: "new", dir: 0 };
  if (b == null) return { txt: "cleared", dir: 0 };
  const d = b - a;
  return { txt: (d >= 0 ? "+" : "") + money(d), dir: d > 0 ? 1 : d < 0 ? -1 : 0 };
}

async function loadWipReview(force) {
  const note = $("#wrNote"), body = $("#wrBody");
  if (!body) return;
  if (wrPoll) return;                                  // a compute/sync run is in flight
  try { WR = await (await fetch("/api/wip/review")).json(); }
  catch (e) { if (note) note.textContent = "could not load"; return; }
  if (!WR.ready) {
    if (note) note.textContent = "";
    $("#wrFilters").hidden = true; $("#wrSync").hidden = true; $("#wrStats").innerHTML = "";
    body.innerHTML = `<div class="wr-empty"><p>No pending review yet.</p>
      <p class="hint">Hit <b>Compute pending update</b> to run the WIP pipeline and see every
      change before anything is written. It pulls Billed/Costs from QuickBooks (a few Touch ID
      prompts) and takes a few minutes.</p></div>`;
    return;
  }
  wrInitDecisions();
  renderWipReview();
}

function wrInitDecisions() {
  // Fresh review → default marks: QBO facts approved, PM fields left for an answer.
  wrDecisions = {}; wrDrop = new Set();
  for (const r of WR.records) {
    if (r.status === "SAME") continue;
    const marks = {};
    for (const f of wrChanged(r.fields)) marks[f.key] = (f.block === "qbo");
    wrDecisions[r.project_num] = marks;
  }
}

function renderWipReview() {
  const note = $("#wrNote"), body = $("#wrBody");
  const g = WR.generated || {}, at = Object.values(g).map(x => x.at).filter(Boolean).sort().pop();
  if (note) note.textContent = at ? `computed ${fmtDate(at, true)}` : "";
  $("#wrFilters").hidden = false; $("#wrSync").hidden = false;
  renderWrStats();
  const div = $("#wrDivision").value, st = $("#wrStatus").value;
  const q = ($("#wrSearch").value || "").trim().toLowerCase();
  const changedOnly = $("#wrChangedOnly").checked;
  body.innerHTML = "";
  let shown = 0;
  for (const dv of WR_DIV_ORDER) {
    if (div && div !== dv) continue;
    let recs = WR.records.filter(r => r.division === dv);
    if (changedOnly) recs = recs.filter(r => r.status !== "SAME");
    if (st) recs = recs.filter(r => r.status === st);
    if (q) recs = recs.filter(r => (r.project_num + " " + r.name).toLowerCase().includes(q));
    if (!recs.length) continue;
    const gen = g[dv];
    const head = document.createElement("div");
    head.className = "wr-div-head";
    head.innerHTML = `<span>${dv}</span><span class="wr-div-sub">${gen ? gen.tab : ""} · ${recs.length} shown</span>`;
    body.appendChild(head);
    for (const r of recs) { body.appendChild(wrJobCard(r)); shown++; }
  }
  if (!shown) body.innerHTML = `<div class="wr-empty"><p>Nothing matches the filters.</p></div>`;
  wrUpdateApproveCount();
}

function renderWrStats() {
  const el = $("#wrStats"); if (!el) return;
  const c = WR.counts || {};
  const tiles = [
    ["Jobs in update", c.jobs || 0, ""],
    ["Changed", c.changed || 0, "amber"],
    ["Reversed", c.reversed || 0, c.reversed ? "red" : ""],
    ["Added", c.added || 0, ""],
    ["Removed", c.removed || 0, ""],
  ];
  el.innerHTML = "";
  for (const [label, val, cls] of tiles) {
    const k = document.createElement("div"); k.className = "kpi" + (cls ? " wr-kpi-" + cls : "");
    k.innerHTML = `<div class="k-label">${label}</div><div class="k-value">${val}</div>`;
    el.appendChild(k);
  }
}

function wrJobCard(r) {
  const card = document.createElement("div");
  card.className = "wr-job wr-" + r.status.toLowerCase();
  const badge = `<span class="wr-badge ${r.status.toLowerCase()}">${r.status}</span>`;
  const removed = r.status === "REMOVED", added = r.status === "ADDED";
  let head = `<div class="wr-job-head"><span class="wr-pn">${_ge(r.project_num)}</span>`
    + `<span class="wr-name">${_ge(r.name)}</span>${badge}`;
  if (added) {
    const inc = !wrDrop.has(r.project_num);
    head += `<label class="wr-inc"><input type="checkbox" class="wr-inc-cb" ${inc ? "checked" : ""}> add this job</label>`;
  } else if (removed) {
    head += `<span class="wr-drop-note">will drop off the tab</span>`;
  } else {
    head += `<button class="btn tiny wr-job-all" type="button">Approve job</button>`;
  }
  head += `</div>`;
  card.innerHTML = head;
  if (r.flags) { const fl = document.createElement("div"); fl.className = "wr-flags"; fl.textContent = r.flags; card.appendChild(fl); }
  const carried = (r.fields || []).filter(f => f.carried);
  if (carried.length) {
    const c = document.createElement("div"); c.className = "wr-carried";
    c.textContent = "Kept from the tab (no source document this run): " + carried.map(f => `${f.label} ${money(f.now)}`).join(" · ");
    card.appendChild(c);
  }
  const changed = wrChanged(r.fields);
  for (const block of ["qbo", "pm"]) {
    const fs = changed.filter(f => f.block === block);
    if (!fs.length) continue;
    const wrap = document.createElement("div"); wrap.className = "wr-block";
    wrap.innerHTML = `<div class="wr-block-title ${block}">${block === "qbo" ? "Accept · QuickBooks" : "PM answers"}</div>`;
    for (const f of fs) wrap.appendChild(wrFieldRow(r, f, removed));
    card.appendChild(wrap);
  }
  if (added) card.querySelector(".wr-inc-cb").onchange = e => {
    if (e.target.checked) wrDrop.delete(r.project_num); else wrDrop.add(r.project_num);
    card.classList.toggle("wr-excluded", !e.target.checked); wrUpdateApproveCount();
  };
  const allBtn = card.querySelector(".wr-job-all");
  if (allBtn) allBtn.onclick = () => { for (const f of changed) if (!f.reversed) wrSet(r.project_num, f.key, true); renderWipReview(); };
  return card;
}

function wrFieldRow(r, f, removed) {
  const row = document.createElement("label");
  row.className = "wr-field";
  const d = wrDelta(f);
  const approved = removed ? false : !!(wrDecisions[r.project_num] && wrDecisions[r.project_num][f.key]);
  // Direction colour: up is neutral; a PM value going DOWN is amber (REVERSED - needs a named
  // document), a contract going down is red; a QBO decrease is marked, not coloured.
  let dcls = "";
  if (d && d.dir < 0) dcls = f.block === "pm" ? (f.key === "orig_contract" ? "down" : "amber") : "";
  if (f.reversed) row.classList.add("wr-reversed");
  const src = f.source ? `<span class="wr-src" title="${_ge(f.source_path || f.source)}">${_ge(f.source)}</span>` : "";
  const note = f.note ? `<span class="wr-note">${_ge(f.note)}</span>` : "";
  const mark = f.reversed ? `<span class="wr-mark rev">REVERSED</span>` : f.decreased ? `<span class="wr-mark dec">decreased</span>` : "";
  row.innerHTML =
    `<span class="wr-fl">${_ge(f.label)}</span>`
    + `<span class="wr-was">${money(f.was)}</span><span class="wr-arrow">→</span>`
    + `<span class="wr-now">${money(f.now)}${mark}</span>`
    + (d ? `<span class="wr-delta ${dcls}">${d.txt}</span>` : `<span class="wr-delta"></span>`)
    + ((src || note) ? `<span class="wr-srcline">${src}${note}</span>` : "");
  if (!removed) {
    const cb = document.createElement("input");
    cb.type = "checkbox"; cb.className = "wr-check"; cb.checked = approved;
    cb.onchange = () => { wrSet(r.project_num, f.key, cb.checked); row.classList.toggle("on", cb.checked); wrUpdateApproveCount(); };
    row.appendChild(cb);
    row.classList.toggle("on", approved);
  } else {
    row.classList.add("wr-ro");
  }
  return row;
}

function wrSet(pn, key, val) { (wrDecisions[pn] = wrDecisions[pn] || {})[key] = val; }

function wrUpdateApproveCount() {
  let n = 0;
  for (const pn in wrDecisions) for (const k in wrDecisions[pn]) if (wrDecisions[pn][k]) n++;
  const btn = $("#wrSync");
  if (btn) btn.textContent = n ? `Sync ${n} approved →` : "Sync approved →";
}

function wrBulk(mode) {
  // mode: 'qbo' | 'all' | 'clear' — over the CURRENTLY VISIBLE jobs only.
  const div = $("#wrDivision").value, st = $("#wrStatus").value;
  const q = ($("#wrSearch").value || "").trim().toLowerCase();
  const changedOnly = $("#wrChangedOnly").checked;
  for (const r of WR.records) {
    if (r.status === "SAME") continue;
    if (div && r.division !== div) continue;
    if (changedOnly && r.status === "SAME") continue;
    if (st && r.status !== st) continue;
    if (q && !(r.project_num + " " + r.name).toLowerCase().includes(q)) continue;
    for (const f of wrChanged(r.fields)) {
      if (mode === "clear") wrSet(r.project_num, f.key, false);
      else if (mode === "all" && !f.reversed) wrSet(r.project_num, f.key, true);   // a reversal is never bulk-approved
      else if (mode === "qbo" && f.block === "qbo") wrSet(r.project_num, f.key, true);
    }
  }
  renderWipReview();
}

async function runWipReview() {
  if (WR && WR.ready && !confirm("Recompute the pending WIP update?\n\nThis re-runs the WIP pipeline (CP folders, RP file, MFD) and pulls Billed/Costs from QuickBooks - expect a few Touch ID prompts and a few minutes. Nothing is written.")) return;
  else if (!(WR && WR.ready) && !confirm("Compute the pending WIP update?\n\nRuns the WIP pipeline and pulls Billed/Costs from QuickBooks (a few Touch ID prompts, a few minutes). Nothing is written - you review the changes first.")) return;
  const r = await fetch("/api/wip/review", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ confirm: true }) });
  if (r.status === 409) { alert("A sync is already running - let it finish first."); return; }
  if (!r.ok) { alert("Could not start the review."); return; }
  wrRunProgress("Computing the pending update", () => loadWipReview(true));
}

async function syncWipReview() {
  let n = 0; for (const pn in wrDecisions) for (const k in wrDecisions[pn]) if (wrDecisions[pn][k]) n++;
  const dropped = wrDrop.size;
  if (!confirm(`Write approved changes to the WIP master?\n\n${n} approved change(s) will be written to Test - CP, Test - RP and Test-Master. Unchecked changes keep the current tab value${dropped ? `; ${dropped} added job(s) will be left off` : ""}.\n\nThis writes the production WIP workbook (guarded) and pulls QuickBooks again (Touch ID).`)) return;
  const decisions = wrBuildDecisions();
  const r = await fetch("/api/wip/merge", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ confirm: true, decisions }) });
  if (r.status === 409) { alert("A sync is already running - let it finish first."); return; }
  if (!r.ok) { alert("Could not start the sync."); return; }
  wrRunProgress("Writing the approved changes", () => loadWipReview(true));
}

function wrBuildDecisions() {
  // Carry each disapproved field's "was" so every tab reverts the SAME number.
  const wasOf = {};
  for (const r of WR.records) {
    wasOf[r.project_num] = wasOf[r.project_num] || {};
    for (const f of r.fields) wasOf[r.project_num][f.key] = f.was;
  }
  const fields = {};
  for (const pn in wrDecisions) {
    const m = {};
    for (const k in wrDecisions[pn]) {
      const approved = !!wrDecisions[pn][k];
      m[k] = { approved, revert: approved ? null : (wasOf[pn] ? wasOf[pn][k] : null) };
    }
    fields[pn] = m;
  }
  return { fields, drop_added: [...wrDrop] };
}

function wrRunProgress(label, onDone) {
  const body = $("#wrBody"), note = $("#wrNote");
  $("#wrFilters").hidden = true; $("#wrSync").hidden = true; $("#wrStats").innerHTML = "";
  $("#wrCompute").disabled = true; $("#wrSync").disabled = true;
  body.innerHTML = `<div class="wr-run"><div class="wr-run-label">${_ge(label)}…</div>
    <div class="wr-steps" id="wrSteps"></div>
    <div class="pl-bar"><div class="pl-fill" id="wrFill"></div></div>
    <div class="hint" id="wrRunHint">Running - this can take a few minutes; Touch ID prompts appear on the Mac.</div></div>`;
  if (wrPoll) clearInterval(wrPoll);
  wrPoll = setInterval(async () => {
    let s; try { s = await (await fetch("/api/sync/status")).json(); } catch { return; }
    const steps = s.steps || [];
    const done = steps.filter(x => x.state === "done").length;
    const fill = $("#wrFill"); if (fill) fill.style.width = steps.length ? Math.round(done / steps.length * 100) + "%" : "0%";
    const box = $("#wrSteps");
    if (box) box.innerHTML = steps.map(x => `<div class="wr-step ${x.state}">${x.state === "done" ? "✓" : x.state === "error" ? "✕" : x.state === "running" ? "▶" : "·"} ${_ge(x.label)}</div>`).join("");
    if (s.state !== "running") {
      clearInterval(wrPoll); wrPoll = null;
      $("#wrCompute").disabled = false; $("#wrSync").disabled = false;
      if (s.state === "error") { if (note) note.textContent = "run failed - see the log"; }
      onDone();
    }
  }, 1500);
}

// ── Accounting fixes: the Bill Tracker audits, filterable by audit type ───────
let ACCT = null;            // cached /api/accounting payload
let acctIssue = null;       // the audit-type filter currently active (null = all)
const ACCT_VENDOR_MSEL = { id: "acctVendor", all: "All vendors", get: f => f.vendor || "", search: true, lbl: v => v || "(no vendor)" };
const acctMSel = {}; let _acctVendorSig = null;
let acctSel = new Set();    // selected finding keys (f._k) for copy-as-table
let _acctVisible = [];      // the currently-filtered rows ("Copy all" copies these)
let acctSort = [];          // multi-column sort: [{key, dir}] - click a header to add/cycle asc/desc/off
async function loadAccounting(force) {
  const note = $("#acctNote"), table = $("#acctTable");
  if (!table) return;
  if (ACCT && !force) { renderAccounting(); return; }
  if (note) note.textContent = "loading…";
  skeletonInto(table.tBodies[0] || table, 8);
  try { ACCT = await (await fetch("/api/accounting")).json(); }
  catch (e) { if (note) note.textContent = "could not load"; return; }
  (ACCT.findings || []).forEach((f, i) => { f._k = i; });   // stable key for selection across filters
  acctSel = new Set();
  renderAccounting();
}

function _acctPillClass(issue) {
  const i = (issue || "").toLowerCase();
  if (i.includes("not approved") || i.includes("missing project") || i.includes("no project")) return "warn";
  if (i.includes("duplicate")) return "neg";
  return "info";
}

// Plain-language "what's shown", not the filter-widget labels (owner 2026-08-28: "tell the user
// what it's filtering"). Leads with the result count, then the actual values (no "Issue:"/"Class:").
function acctFilterDesc(shown, total) {
  const parts = [];
  if (acctIssue) parts.push(acctIssue);
  const dv = $("#acctDivision") ? $("#acctDivision").value : ""; if (dv) parts.push(dv);
  const q = ($("#acctSearch").value || "").trim(); if (q) parts.push(`matching "${q}"`);
  if (!parts.length) return "";
  return `${shown} of ${total} bills · ${parts.join(" · ")}`;
}

// Multi-column sort for the Audit table (owner 2026-08-28: "sort by date or vendor ... both ways at
// the same time"). Click a header to add it; click again to flip asc/desc; again to drop it. Columns
// stack in click order, so Vendor-then-Date sorts by vendor, then by date within each vendor.
const ACCT_SORT_KEYS = {
  "Issue": f => f.issue, "Vendor": f => f.vendor, "Bill #": f => f.bill_no,
  "Date": f => f.date || "", "Project": f => f.project, "Class": f => f.division,
  "Cost": f => f.cost_code, "Amount": f => (f.amount == null ? -Infinity : f.amount),
  "Line memo": f => f.memo, "Why flagged": f => f.detail,
};
function _acctCmp(a, b) {
  for (const s of acctSort) {
    const acc = ACCT_SORT_KEYS[s.key]; if (!acc) continue;
    const av = acc(a), bv = acc(b);
    const c = (typeof av === "number" || typeof bv === "number")
      ? (av || 0) - (bv || 0) : String(av || "").localeCompare(String(bv || ""));
    if (c) return c * s.dir;
  }
  return 0;
}
function _acctToggleSort(key) {
  const i = acctSort.findIndex(s => s.key === key);
  if (i < 0) acctSort.push({ key, dir: 1 });
  else if (acctSort[i].dir === 1) acctSort[i].dir = -1;
  else acctSort.splice(i, 1);
  renderAccounting();
}

function renderAccounting() {
  const note = $("#acctNote"), stats = $("#acctStats"), filt = $("#acctFilters"), table = $("#acctTable");
  if (!table) return;
  const thead = table.querySelector("thead"), tbody = table.querySelector("tbody");
  if (!ACCT || !ACCT.ok) {
    if (note) note.textContent = ACCT && ACCT.error ? "unavailable" : "";
    if (stats) stats.innerHTML = ""; if (filt) filt.innerHTML = ""; thead.innerHTML = "";
    tbody.innerHTML = ACCT && ACCT.error ? `<tr><td class="left" style="padding:14px;color:var(--text-dim)">${_ge(ACCT.error)}</td></tr>` : "";
    return;
  }
  const all = ACCT.findings || [];
  if (note) note.textContent = `${all.length} to fix`;
  const counts = ACCT.counts || {};
  // stat tiles: total + per themed group
  const groups = {};
  for (const f of all) groups[f.group] = (groups[f.group] || 0) + 1;
  stats.innerHTML = "";
  const tile = (label, val, cls) => { const k = document.createElement("div"); k.className = "kpi" + (cls ? " " + cls : ""); k.innerHTML = `<div class="k-label">${_ge(label)}</div><div class="k-value">${val}</div>`; stats.appendChild(k); };
  tile("All fixes", all.length, all.length ? "wr-kpi-amber" : "");
  for (const g of ["Coding", "Bills", "PO"]) if (groups[g]) tile(g, groups[g]);
  // filter chips by audit type (with counts)
  filt.innerHTML = "";
  const chip = (label, key, n) => {
    const b = document.createElement("button"); b.className = "acct-chip" + (acctIssue === key ? " active" : "");
    b.innerHTML = `${_ge(label)} <span class="ac-n">${n}</span>`;
    b.onclick = () => { acctIssue = acctIssue === key ? null : key; renderAccounting(); };
    filt.appendChild(b);
  };
  chip("All", null, all.length);
  // two labelled sections (owner 2026-09-02): the BILL audits (Coding + Bills sheets) and the PO audits
  const issGroup = {}; for (const f of all) issGroup[f.issue] = f.group;
  const sect = (label) => { const s = document.createElement("span"); s.className = "acct-sect"; s.textContent = label; filt.appendChild(s); };
  const byCount = Object.keys(counts).sort((a, b) => counts[b] - counts[a]);
  const billIss = byCount.filter(i => issGroup[i] !== "PO"), poIss = byCount.filter(i => issGroup[i] === "PO");
  if (billIss.length) { sect("Bills"); for (const iss of billIss) chip(iss, iss, counts[iss]); }
  if (poIss.length) { sect("POs"); for (const iss of poIss) chip(iss, iss, counts[iss]); }
  // vendor checkbox filter (same multi-select as the other tabs)
  { const sig = String(all.length); if (sig !== _acctVendorSig || !($("#acctVendorMenu") && $("#acctVendorMenu").querySelector(".msel-opt"))) { _acctVendorSig = sig; buildMSel(ACCT_VENDOR_MSEL, all, acctMSel, renderAccounting); } }
  // division filter
  const dsel = $("#acctDivision");
  if (dsel && dsel.options.length <= 1) for (const d of [...new Set(all.map(f => f.division).filter(Boolean))].sort()) { const o = document.createElement("option"); o.value = d; o.textContent = d; dsel.appendChild(o); }
  const dv = dsel ? dsel.value : "", q = ($("#acctSearch").value || "").trim().toLowerCase();
  const rows = all.filter(f => (!acctIssue || f.issue === acctIssue) && (!dv || f.division === dv) && mselPasses(f, [ACCT_VENDOR_MSEL], acctMSel)
    && (!q || (f.vendor + " " + f.project + " " + f.bill_no + " " + (f.memo || "") + " " + f.detail).toLowerCase().includes(q)));
  _setHintFilter("accounting", acctFilterDesc(rows.length, all.length));   // count + what's filtered (generic when All)
  if (acctSort.length) rows.sort(_acctCmp);   // multi-column sort (applied before the render cap)
  _acctVisible = rows;
  // fixed meta widths (px) so one long outlier can't blow a column wide (the old wasted
  // space); the two text columns (null width) share the rest and wrap - nothing truncates.
  const cols = [["Issue", "left audit-soft", 126], ["Vendor", "left audit-soft", 148], ["Bill #", "left", 78],
    ["📎", "left", 52], ["Date", "left", 104], ["Project", "left", 122], ["Class", "left", 104], ["Cost", "left", 64], ["Amount", "right", 92],
    ["Line memo", "left audit-soft", null], ["Why flagged", "left audit-soft", null]];
  thead.innerHTML = ""; const htr = document.createElement("tr");
  const chTh = document.createElement("th"); chTh.className = "left acct-check"; chTh.style.width = "32px";
  const selAll = document.createElement("input"); selAll.type = "checkbox"; selAll.id = "acctSelAll"; selAll.title = "Select all shown";
  selAll.onchange = () => { if (selAll.checked) rows.forEach(f => acctSel.add(f._k)); else rows.forEach(f => acctSel.delete(f._k)); renderAccounting(); };
  chTh.appendChild(selAll); htr.appendChild(chTh);
  for (const [c, cls, w] of cols) {
    const th = document.createElement("th"); th.className = cls; if (w) th.style.width = w + "px";
    const si = acctSort.findIndex(s => s.key === c);
    if (ACCT_SORT_KEYS[c]) {
      th.classList.add("acct-sortable");
      th.textContent = c + (si >= 0 ? (acctSort[si].dir === 1 ? " ▲" : " ▼") + (acctSort.length > 1 ? (si + 1) : "") : "");
      th.title = "Click to sort; click again to reverse; a third click clears it. Sort by more than one column - they stack in click order.";
      th.onclick = () => _acctToggleSort(c);
    } else th.textContent = c;
    htr.appendChild(th);
  }
  thead.appendChild(htr);
  tbody.innerHTML = "";
  if (!rows.length) {
    const tr = document.createElement("tr"), td = document.createElement("td");
    td.colSpan = cols.length + 1; td.className = "left"; td.style.cssText = "padding:14px;color:var(--text-dim)";
    td.textContent = all.length ? "Nothing matches the filters." : "No audit findings - everything's clean.";
    tr.appendChild(td); tbody.appendChild(tr); _acctUpdateSelAll(); _acctUpdateCopyBtn(); _acctUpdateDownloadBtn(); return;
  }
  const ACCT_CAP = 250;   // render cap - all ~1900 rows (each w/ a checkbox + scan button) crashed the tab
  const frag = document.createDocumentFragment();
  for (const f of rows.slice(0, ACCT_CAP)) {
    const tr = document.createElement("tr");
    const chTd = document.createElement("td"); chTd.className = "left acct-check";
    const cb = document.createElement("input"); cb.type = "checkbox"; cb.checked = acctSel.has(f._k);
    cb.onchange = () => { if (cb.checked) acctSel.add(f._k); else acctSel.delete(f._k); _acctUpdateSelAll(); _acctUpdateCopyBtn(); _acctUpdateDownloadBtn(); };
    chTd.appendChild(cb); tr.appendChild(chTd);
    const ic = document.createElement("td"); ic.className = "left audit-soft";
    { const src = document.createElement("span"); src.className = "acct-src " + (f.group === "PO" ? "po" : "bill"); src.textContent = f.group === "PO" ? "PO" : "Bill"; src.title = f.group === "PO" ? "A purchase-order finding (Audit - PO)" : `A bill finding (Audit - ${f.group})`; ic.appendChild(src); }
    const pill = document.createElement("span"); pill.className = "acct-pill " + _acctPillClass(f.issue); pill.textContent = f.issue; ic.appendChild(pill); tr.appendChild(ic);
    const vc = leftText(f.vendor || "–"); vc.classList.add("audit-soft"); tr.appendChild(vc);
    tr.appendChild(qboLinkCell(f.bill_no, f.url, "Open this bill in QuickBooks"));
    const sc = document.createElement("td"); sc.className = "left";
    if (f.att > 0) {
      const b = document.createElement("button"); b.type = "button"; b.className = "acct-scan";
      b.textContent = f.att > 1 ? ("📎" + f.att) : "📎";   // 📎 / 📎N
      b.title = f.att > 1 ? (f.att + " scans - click to choose") : "Open the bill scan (no QBO)";
      b.onclick = (e) => { e.stopPropagation(); openBillScan(f, b); };
      sc.appendChild(b);
    }
    tr.appendChild(sc);
    tr.appendChild(leftText(f.date ? fmtDateShort(f.date) : "–"));
    const pc = leftText(f.project || "–"); pc.title = f.project || ""; tr.appendChild(pc);
    tr.appendChild(leftText(f.division || "–"));   // Class (QBO division)
    tr.appendChild(leftText(f.cost_code || "–"));
    tr.appendChild(rightText(f.amount != null ? money(f.amount) : ""));
    const mc = document.createElement("td"); mc.className = "left audit-soft"; mc.textContent = f.memo || "–"; if (!f.memo) mc.classList.add("audit-dim"); tr.appendChild(mc);
    const dc = document.createElement("td"); dc.className = "left audit-soft"; dc.textContent = f.detail || ""; tr.appendChild(dc);
    frag.appendChild(tr);
  }
  tbody.appendChild(frag);
  if (rows.length > ACCT_CAP) {
    const tr = document.createElement("tr"), td = document.createElement("td");
    td.colSpan = cols.length + 1; td.className = "left"; td.style.cssText = "padding:12px 14px;color:var(--text-dim)";
    td.textContent = `Showing the first ${ACCT_CAP} of ${rows.length} - narrow with a chip, division, or the search above (Copy still takes all ${rows.length}).`;
    tr.appendChild(td); tbody.appendChild(tr);
  }
  _acctUpdateSelAll(); _acctUpdateCopyBtn(); _acctUpdateDownloadBtn();
}

// Copy-as-table: the columns copied (headers + values), minus the checkbox and 📎 columns.
const ACCT_COPY_COLS = [["Issue", f => f.issue], ["Vendor", f => f.vendor], ["Bill #", f => f.bill_no],
  ["Date", f => f.date ? fmtDateShort(f.date) : ""], ["Project", f => f.project], ["Class", f => f.division],
  ["Cost", f => f.cost_code], ["Amount", f => f.amount != null ? money(f.amount) : ""],
  ["Line memo", f => f.memo], ["Why flagged", f => f.detail]];

function _acctUpdateCopyBtn() {
  const b = $("#btnAcctCopy"); if (!b || b.disabled) return;
  b.textContent = acctSel.size ? `Copy ${acctSel.size}` : `Copy all (${_acctVisible.length})`;
}
function _acctUpdateSelAll() {
  const sa = $("#acctSelAll"); if (!sa) return;
  const n = _acctVisible.filter(f => acctSel.has(f._k)).length;
  sa.checked = _acctVisible.length > 0 && n === _acctVisible.length;
  sa.indeterminate = n > 0 && n < _acctVisible.length;
}

// Put the rows on the clipboard as BOTH tab-separated text (pastes into Excel/Sheets as
// cells) and an HTML table (pastes into email/Word as a formatted table) - header first.
async function copyAcctTable(rows) {
  const heads = ACCT_COPY_COLS.map(c => c[0]);
  const body = rows.map(f => ACCT_COPY_COLS.map(c => { const v = c[1](f); return v == null ? "" : String(v); }));
  const clean = v => v.replace(/[\t\r\n]+/g, " ").trim();
  const tsv = [heads, ...body].map(r => r.map(clean).join("\t")).join("\n");
  const html = "<table><thead><tr>" + heads.map(h => `<th>${_ge(h)}</th>`).join("") + "</tr></thead><tbody>"
    + body.map(r => "<tr>" + r.map(v => `<td>${_ge(v)}</td>`).join("") + "</tr>").join("") + "</tbody></table>";
  try {
    if (window.ClipboardItem && navigator.clipboard && navigator.clipboard.write) {
      await navigator.clipboard.write([new ClipboardItem({
        "text/plain": new Blob([tsv], { type: "text/plain" }),
        "text/html": new Blob([html], { type: "text/html" }) })]);
    } else if (navigator.clipboard && navigator.clipboard.writeText) {
      await navigator.clipboard.writeText(tsv);
    } else { throw new Error("no clipboard"); }
    return true;
  } catch (e) {
    try { const ta = document.createElement("textarea"); ta.value = tsv; ta.style.cssText = "position:fixed;opacity:0"; document.body.appendChild(ta); ta.select(); const ok = document.execCommand("copy"); ta.remove(); return ok; }
    catch (_) { return false; }
  }
}

async function _acctDoCopy() {
  const rows = acctSel.size ? ((ACCT && ACCT.findings) || []).filter(f => acctSel.has(f._k)) : _acctVisible;
  if (!rows.length) return;
  const ok = await copyAcctTable(rows);
  const b = $("#btnAcctCopy"); if (!b) return;
  b.disabled = true; b.textContent = ok ? `Copied ${rows.length} ✓` : "Copy failed";
  setTimeout(() => { b.disabled = false; _acctUpdateCopyBtn(); }, 1400);
}

// Download the selected rows' bill scans to a folder + open it, so the owner can drag them into
// the message to the responsible party (owner 2026-08-31: "download attachments ... want to show
// the attachments"). Targets the selection (or all shown when nothing's ticked); only rows that
// HAVE a scan count. The backend saves each scan named to match the copied table, then reveals
// the folder. Batch-capped so it stays "a folder to attach", not a bulk export.
let _acctDownloading = false;
function _acctDownloadTarget() {
  const tgt = acctSel.size ? ((ACCT && ACCT.findings) || []).filter(f => acctSel.has(f._k)) : _acctVisible;
  return tgt.filter(f => f.att > 0);
}
function _acctUpdateDownloadBtn() {
  const b = $("#btnAcctDownload"); if (!b || _acctDownloading) return;
  const n = _acctDownloadTarget().length, over = n > 60;   // 60 = the batch cap (a folder to attach, not an export)
  b.disabled = n === 0 || over;
  b.textContent = (n && !over) ? `Download scans (${n})` : "Download scans";
  b.title = over ? `Too many (${n}) - tick up to 60 rows, or filter smaller, to download their scans`
    : (n ? "Download these rows' bill scans into a folder and open it - drag them into your message"
         : "Tick rows (or filter) to download their bill scans");
}
async function _acctDoDownload() {
  const b = $("#btnAcctDownload"); if (!b || b.disabled || _acctDownloading) return;
  const withScans = _acctDownloadTarget();
  if (!withScans.length) { toast("None of those rows have a scan to download"); return; }
  if (withScans.length > 60) { toast(`Too many (${withScans.length}) - tick up to 60 rows to download at once`); return; }
  const bills = withScans.map(f => { const m = /txnId=(\d+)/.exec(f.url || ""); return m ? { txnId: m[1], bill_no: f.bill_no, vendor: f.vendor, type: "Bill" } : null; }).filter(Boolean);
  if (!bills.length) { toast("Could not resolve those bills"); return; }
  _acctDownloading = true; b.disabled = true; b.textContent = `Downloading ${bills.length}…`;
  try {
    const r = await (await fetch("/api/attachment/download", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ bills }) })).json();
    if (r && r.ok && r.count) {
      b.textContent = `Downloaded ${r.count} ✓`;
      toast(`${r.count} scan${r.count === 1 ? "" : "s"} from ${r.bills} bill${r.bills === 1 ? "" : "s"} → folder opened · auto-clears in 24h`);
    } else if (r && r.ok) {
      b.textContent = "Download scans"; toast("No scans found on those bills");
    } else {
      b.textContent = "Download scans"; toast((r && r.error) || "Download failed");
    }
  } catch (e) {
    b.textContent = "Download scans"; toast("Download failed");
  }
  setTimeout(() => { _acctDownloading = false; _acctUpdateDownloadBtn(); }, 1800);
}

// Resolve a bill's scan link(s) on click - the dashboard fetches FRESH (minutes-lived)
// QBO download links, so the file opens without going into QuickBooks. One scan opens
// straight away; several show a chooser.
async function openBillScan(f, el) {
  const m = /txnId=(\d+)/.exec(f.url || ""); if (!m) return;
  const orig = el.textContent; el.textContent = "…"; el.disabled = true;
  try {
    const r = await (await fetch(`/api/attachment?bill=${m[1]}`)).json();
    el.disabled = false; el.textContent = orig;
    const files = (r && r.files) || [];
    if (!r || !r.ok || !files.length) { el.title = (r && r.error) || "No scan available"; el.classList.add("scan-empty"); return; }
    if (files.length === 1) window.open(files[0].url, "_blank", "noopener");
    else showScanMenu(files, el);
  } catch (e) { el.disabled = false; el.textContent = orig; }
}

function showScanMenu(files, el) {
  _closeScanMenu();
  const menu = document.createElement("div"); menu.className = "scan-menu"; menu.id = "scanMenu";
  for (const f of files) {
    const a = document.createElement("a"); a.href = f.url; a.target = "_blank"; a.rel = "noopener";
    a.textContent = f.name || "attachment"; a.onclick = () => setTimeout(_closeScanMenu, 0); menu.appendChild(a);
  }
  document.body.appendChild(menu);
  const r = el.getBoundingClientRect();
  menu.style.left = Math.max(8, Math.min(r.left, window.innerWidth - menu.offsetWidth - 12)) + "px";
  menu.style.top = (r.bottom + 4) + "px";
  setTimeout(() => document.addEventListener("click", _closeScanMenu, { once: true }), 0);
}
function _closeScanMenu() { const m = $("#scanMenu"); if (m) m.remove(); }

function init() {
  applySettings();
  syncSettingsUI();
  wireSettings();
  ["#search", "#fDivision", "#fStatus", "#fCategory", "#fActive"].forEach(sel =>
    $(sel).addEventListener("input", renderProjects));
  // Draws filters are multi-selects now (built by buildDrawFilters, toggled via _mselWraps below).
  { const el = $("#vendorSearch"); if (el) el.addEventListener("input", renderVendors); }
  { const el = $("#vendorGroupType"); if (el) el.addEventListener("change", renderVendors); }
  { const el = $("#vendorExpandAll"); if (el) el.onclick = _vendorToggleAll; }
  { const el = $("#lienFProj"); if (el) el.addEventListener("input", renderLiens); }   // the other lien filters are multi-selects now
  { const el = $("#wipActive"); if (el) el.addEventListener("change", renderWip); }
  { const el = $("#billSort"); if (el) el.addEventListener("change", renderBills); }
  // Month + Vendor + every categorical multi-select: the button toggles the checkbox menu; a click outside closes it.
  const _mselWraps = [["#bfDateBtn", "#bfDateMenu", "#bfDateMsel"], ["#dfDateBtn", "#dfDateMenu", "#dfDateMsel"], ["#bfVendorBtn", "#bfVendorMenu", "#bfVendorMsel"],
    ...BILL_MSEL.map(c => [`#${c.id}Btn`, `#${c.id}Menu`, `#${c.id}Msel`]),
    ...LIEN_MSEL.map(c => [`#${c.id}Btn`, `#${c.id}Menu`, `#${c.id}Msel`]),
    ...PAY_MSEL.map(c => [`#${c.id}Btn`, `#${c.id}Menu`, `#${c.id}Msel`]),
    ...INV_MSEL.map(c => [`#${c.id}Btn`, `#${c.id}Menu`, `#${c.id}Msel`]), ["#ifMonthBtn", "#ifMonthMenu", "#ifMonthMsel"], ["#acctVendorBtn", "#acctVendorMenu", "#acctVendorMsel"],
    ...DRAW_MSEL.map(c => [`#${c.id}Btn`, `#${c.id}Menu`, `#${c.id}Msel`])];
  const _closeMsels = (except) => { for (const [, mId] of _mselWraps) { const m = $(mId); if (m && mId !== except) m.hidden = true; } };
  for (const [btnId, menuId, wrapId] of _mselWraps) {
    const btn = $(btnId), menu = $(menuId);
    if (btn && menu) {
      btn.addEventListener("click", (e) => { e.stopPropagation(); const open = menu.hidden; _closeMsels(menuId); menu.hidden = !open; if (open) _placeMenu(btn, menu); });   // one menu open at a time; pinned to the viewport so no card clips it
      document.addEventListener("click", (e) => { if (!menu.hidden && !e.target.closest(wrapId)) menu.hidden = true; });
    }
  }

  { const el = $("#billGroup"); if (el) el.addEventListener("change", () => {
    const grp = el.value;   // re-collapse under the new grouping (collapse stays the default)
    billsCollapsed = grp === "none" ? new Set() : new Set((BILLS || []).map(b => billGroupKey(b, grp)));
    renderBills(); }); }
  { const el = $("#bfClear"); if (el) el.onclick = billClearFilters; }
  { const el = $("#bfCollapse"); if (el) el.onclick = billToggleAll; }
  ["#ifDivision", "#ifLien", "#ifLienClock", "#ifLitig", "#ifSort"].forEach(sel => { const el = $(sel); if (el) el.addEventListener("change", renderOpenInvoices); });
  { const el = $("#ifClear"); if (el) el.onclick = invClearFilters; }
  // Quick find: type to filter (short debounce); ⌘F / Ctrl+F on the Invoices tab jumps here; Esc clears.
  { const q = $("#ifQuick"); let tq = null;
    if (q) { q.addEventListener("input", () => { clearTimeout(tq); tq = setTimeout(() => { invQuick = q.value.trim(); renderOpenInvoices(); }, 120); });
      q.addEventListener("keydown", e => { if (e.key === "Escape") { q.value = ""; invQuick = ""; renderOpenInvoices(); q.blur(); e.stopPropagation(); } }); }
    document.addEventListener("keydown", e => {
      if (!(e.metaKey || e.ctrlKey) || e.key.toLowerCase() !== "f" || e.altKey) return;
      if (activeTab !== "invoices" || !$("#ifQuick")) return;
      if (document.querySelector(".panel:not([hidden])")) return;          // a side panel is open - leave the browser's find alone
      e.preventDefault(); const el = $("#ifQuick"); el.focus(); el.select();
    }); }
  // Saved views: the current filters + sort + scope + quick find under a name (localStorage, per person).
  buildInvViews();
  { const sv = $("#ifSaveView"); if (sv) sv.onclick = invSaveView; }
  { const dv = $("#ifDelView"); if (dv) dv.onclick = invDeleteView; }
  { const vs = $("#ifViews"); if (vs) vs.onchange = () => invApplyView(vs.value); }
  { const el = $("#ifCollapse"); if (el) el.onclick = invToggleAll; }
  { const el = $("#ifSubGroup"); if (el) el.onclick = invSubGroupToggle; }
  { const el = $("#ifStatement"); if (el) el.onclick = openInvStatement; }
  { const el = $("#btnCopyStmt"); if (el) el.onclick = copyInvStatement; }
  { const seg = $("#invViewSeg"); if (seg) for (const b of seg.querySelectorAll(".seg-btn")) b.onclick = () => { invView = b.dataset.view; seg.querySelectorAll(".seg-btn").forEach(x => x.classList.toggle("on", x === b)); renderOpenInvoices(); }; }
  { const seg = $("#invScopeSeg"); if (seg) for (const b of seg.querySelectorAll(".seg-btn")) b.onclick = () => _setInvScope(b.dataset.scope, seg, b); }
  { const el = $("#btnCloseStmt"); if (el) el.onclick = closePanels; }
  { const el = $("#btnCloseBillDetail"); if (el) el.onclick = closePanels; }
  { const el = $("#btnClosePayBills"); if (el) el.onclick = closePanels; }
  { const el = $("#btnCloseInvDetail"); if (el) el.onclick = closePanels; }
  { const el = $("#btnCloseLienReview"); if (el) el.onclick = closePanels; }
  { const el = $("#billSaveText"); if (el) el.onclick = openLienReview; }   // press the unsaved count → review them
  { const el = $("#btnLienReview"); if (el) el.onclick = openLienReview; }   // saved marks on file, reviewed on demand
  { const el = $("#btnCloseSublocDetail"); if (el) el.onclick = closePanels; }
  { const el = $("#btnCloseVendorDetail"); if (el) el.onclick = closePanels; }
  { const el = $("#recordBack"); if (el) el.onclick = closeRecord; }
  { const el = $("#btnSaveBillMarks"); if (el) el.onclick = saveBillMarks; }
  { const el = $("#btnDiscardBillMarks"); if (el) el.onclick = discardBillMarks; }
  // Pay Bills (check-run worksheet)
  { const el = $("#pfSearch"); if (el) el.addEventListener("input", renderPayBills); }
  ["#pfShow", "#pfFunded"].forEach(sel => { const el = $(sel); if (el) el.addEventListener("change", renderPayBills); });
  { const el = $("#pfSelectAll"); if (el) el.onclick = paySelectAllShown; }
  { const el = $("#pfClearRun"); if (el) el.onclick = clearPayRun; }
  { const el = $("#pfExport"); if (el) el.onclick = exportPayList; }
  { const el = $("#btnSavePayRun"); if (el) el.onclick = savePayRun; }
  { const el = $("#btnDiscardPayRun"); if (el) el.onclick = discardPayRun; }
  window.addEventListener("beforeunload", (e) => { if (pendingBillMarks.size || payDraft.size) { e.preventDefault(); e.returnValue = ""; } });
  $$(".sec-head").forEach(h => h.onclick = () => { const k = h.dataset.sec;
    if (sublocCollapsed.has(k)) sublocCollapsed.delete(k); else sublocCollapsed.add(k); applySublocSections(); });
  try { const bv = localStorage.getItem("proficient-ledger-billview"); if (bv && BILL_VIEWS.some(v => v.id === bv)) activeBillView = bv; } catch { /* ignore */ }
  ["#salesSearch", "#salesStage", "#salesDivision"].forEach(sel => { const el = $(sel); if (el) el.addEventListener("input", renderSales); });
  ["#pnlFProj", "#pnlFClient"].forEach(sel => { const el = $(sel); if (el) el.addEventListener("input", renderPnl); });
  ["#pnlFDivision", "#pnlFStatus"].forEach(sel => { const el = $(sel); if (el) el.addEventListener("change", renderPnl); });
  { const el = $("#pnlSortSel"); if (el) el.addEventListener("change", () => {
      const m = { worst: { key: "net", dir: 1 }, best: { key: "net", dir: -1 }, earned: { key: "earned", dir: -1 },
        cost: { key: "cost", dir: -1 }, contract: { key: "contract", dir: -1 }, name: { key: "proj", dir: 1 } }[el.value];
      if (m) pnlSort = m; renderPnl(); }); }
  { const el = $("#btnClearLien"); if (el) el.onclick = () => { activeLien = null; renderLiens(); }; }
  { const el = $("#btnClearDrawStage"); if (el) el.onclick = () => { activeDrawStage = null; renderFunding(); }; }
  { const el = $("#homeDivision"); if (el) el.addEventListener("input", renderHome); }
  $("#btnExport").onclick = exportCSV;
  $("#btnRefresh").onclick = manualRefresh;
  { const el = $("#btnResync"); if (el) el.onclick = startResync; }
  { const el = $("#btnResyncTop"); if (el) el.onclick = startResync; }   // the same reload, from the strip that is always on screen
  { const p = $("#syncPill"); if (p) p.onclick = () => { setTab("overview"); const w = $("#homeFresh"); if (w) w.scrollIntoView({ behavior: "smooth", block: "center" }); }; }
  setInterval(renderSyncPill, 5000);   // reflects "Syncing…" while a run is in flight and the age as time passes
  // Bills: the secondary filters live behind "More filters" (owner 2026-09-01: "10 dropdowns + 8 pills
  // above the fold"); remembered per person, and forced open while one of them is active.
  { const btn = $("#bfMoreBtn"), more = $("#bfMore");
    if (btn && more) {
      let on = false; try { on = localStorage.getItem("proficient-ledger-billmore") === "1"; } catch { /* ignore */ }
      const paint = () => { more.hidden = !on; btn.classList.toggle("on", on); btn.textContent = on ? "Fewer filters" : "More filters"; };
      btn.onclick = () => { on = !on; try { localStorage.setItem("proficient-ledger-billmore", on ? "1" : "0"); } catch { /* ignore */ } paint(); };
      window._billMoreOpenIfActive = () => { if (!on && BILL_MSEL.some(c => (billMSel[c.id] || {}).size && more.contains($("#" + c.id + "Msel")))) { on = true; paint(); } };
      paint();
    } }
  { const pr = $("#btnPayRunGo"); if (pr) pr.onclick = () => setTab("paybills"); }
  { const lr = $("#btnLienRegGo"); if (lr) lr.onclick = () => setTab("liens"); }
  for (const id of ["btnBackBillsPay", "btnBackBillsLien"]) { const b = $("#" + id); if (b) b.onclick = () => setTab("bills"); }
  { const el = $("#btnCostsFull"); if (el) el.onclick = () => runPipeline("costs-full",
      "Reload ALL job costs from QuickBooks?\n\nEvery project, all history - a full replace, so bills that were deleted or re-coded in QuickBooks drop out (the 90-day Resync never removes them). Read-only on QuickBooks; one Touch ID; 30 to 40 minutes.",
      { btn: $("#btnCostsFull"), prog: $("#syncProgress"), fill: $("#syncBarFill"), step: $("#syncStep") }); }
  { const el = $("#btnSyncApAr"); if (el) el.onclick = runApAr; }
  { const el = $("#btnFullRefresh"); if (el) el.onclick = () => runPipeline("all",
      "Full refresh - run EVERY pipeline?\n\nRuns the source producers (AR sync -> Notion/Teams, AP sync -> Bill Tracker.xlsx) AND the loaders, in order. Real writes; expect multiple Touch ID prompts; takes a few minutes.",
      { ..._consoleEls(), btn: el }); }
  $("#btnClearRule").onclick = () => { activeRule = null; renderAttention(); renderProjects(); };
  $("#btnSettings").onclick = () => openPanel("#settings");
  { const el = $("#btnSysReload"); if (el) el.onclick = () => loadSystems(true); }
  for (const id of ["#sysSearch", "#sysOwner", "#sysHealth", "#sysState", "#sysLife", "#sysRetired"]) {
    const el = $(id); if (el) el.addEventListener("input", renderSystems);
  }
  { const el = $("#btnGraphReload"); if (el) el.onclick = () => loadGraph(true); }
  { const el = $("#graphFit"); if (el) el.onclick = () => { if (GV) { GV._userMoved = false; fitGraph(); _gmark(); } }; }
  { const el = $("#graphSearch"); if (el) el.addEventListener("input", e => graphSearch(e.target.value)); }
  { const el = $("#btnHealthPull"); if (el) el.onclick = () => runPipeline("healthpull",
      "Pull QBO health metrics now?\n\nOne loader: bank balances, P&L blocks, 13 weeks of cash flow, and the recurring-obligations register - read-only against QuickBooks, Touch ID on this Mac, under a minute. Everything else on the Health tab is already live from the ledger.",
      { btn: el, prog: $("#healthProg"), fill: $("#healthFill"), step: $("#healthStep") }); }
  { const el = $("#btnAcctReload"); if (el) el.onclick = () => loadAccounting(true); }
  { const el = $("#btnAcctCopy"); if (el) el.onclick = _acctDoCopy; }
  { const el = $("#btnAcctDownload"); if (el) el.onclick = _acctDoDownload; }
  for (const id of ["#acctSearch", "#acctDivision"]) { const el = $(id); if (el) el.addEventListener("input", () => { if (ACCT && ACCT.ok) renderAccounting(); }); }
  { const el = $("#wrCompute"); if (el) el.onclick = runWipReview; }
  { const el = $("#wrSync"); if (el) el.onclick = syncWipReview; }
  { const el = $("#wrApproveQbo"); if (el) el.onclick = () => wrBulk("qbo"); }
  { const el = $("#wrApproveAll"); if (el) el.onclick = () => wrBulk("all"); }
  { const el = $("#wrClearAll"); if (el) el.onclick = () => wrBulk("clear"); }
  for (const id of ["#wrSearch", "#wrDivision", "#wrStatus", "#wrChangedOnly"]) {
    const el = $(id); if (el) el.addEventListener("input", () => { if (WR && WR.ready) renderWipReview(); });
  }
  buildGroupBar();   // generate the two-level nav (top groups; sub-tabs render on setTab)
  let savedTab = "overview";
  try { savedTab = localStorage.getItem("proficient-ledger-tab") || "overview"; } catch { /* ignore */ }
  if (savedTab === "home" || savedTab === "graph") savedTab = "overview";   // merged/removed tabs
  setTab(savedTab);
  initCellSelect();   // Excel-style click/drag cell selection + running-sum bar
  setInterval(() => { if (!syncing && !pendingBillMarks.size && !payDraft.size) load(true); }, 90000);   // soft auto-refresh (paused during a resync or while lien / pay-run marks are unsaved)
  $("#btnCloseSettings").onclick = closePanels;
  $("#btnCloseDetail").onclick = closePanels;
  $("#btnCopyDetail").onclick = () => copy(detailAsText());
  $("#overlay").onclick = closePanels;
  document.addEventListener("keydown", e => { if (e.key === "Escape") closePanels(); });
  matchMedia("(prefers-color-scheme: dark)").addEventListener("change", () => { if (settings.theme === "auto") applySettings(); });
  load();
}
init();
