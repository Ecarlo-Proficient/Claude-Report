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
  theme: "auto", accent: "#3E7A5C", font: "system", fontSize: 14,
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
  if (v === null || v === undefined || v === "") return "—";
  const n = Number(v); if (Number.isNaN(n)) return "—";
  const s = "$" + Math.round(Math.abs(n)).toLocaleString();
  return n < 0 ? "-" + s : s;
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
let SUBLOC = { summary: null, divisions: {}, projects: [], open_by_project: {}, events: [] };
let OI = { as_of: null, buckets: ["Current", "1-30", "31-60", "61-90", "90+"], invoices: [] };  // open AR invoices (aging tab)
let PAY = { payments: [], total_received: 0, count: 0, invoices_paid: 0 };   // received payments, each with the invoices it paid
let paymentsExpanded = new Set();   // payment ids expanded to show their invoices (default: all collapsed - scannable list)
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

// ── Tabs (two-level grouped nav) ─────────────────────────────────────────────
// Parent groups on the top row; the active group's tabs on the second row. The whole
// structure lives here, so adding/moving a tab is a one-line edit (owner 2026-08-19).
const NAV_GROUPS = [
  { id: "home",       label: "My view",       tabs: ["home"] },
  { id: "overview",   label: "Overview",      tabs: ["overview"] },
  { id: "financials", label: "Financials",    tabs: ["pnl", "wip", "costs"] },
  { id: "customers",  label: "Customer",      tabs: ["customers", "invoices", "draws", "payments", "sales"] },
  { id: "vendors",    label: "Vendor",        tabs: ["vendors", "bills", "paybills", "subloc", "liens"] },
  { id: "it",         label: "IT",            tabs: ["systems", "console"] },
];
const TAB_LABELS = {
  home: "My view", overview: "Overview", pnl: "Project P&L", wip: "WIP report", costs: "Costs",
  customers: "Customer Center", invoices: "Invoices", draws: "Draws", payments: "Payments", sales: "Sales Outreach",
  vendors: "Vendor Center", bills: "Bills", paybills: "Pay Bills", subloc: "Sub LOC", liens: "Liens",
  systems: "Systems", console: "Console",
};
const groupOf = t => NAV_GROUPS.find(g => g.tabs.includes(t)) || NAV_GROUPS[0];
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
  $$(".tab-page").forEach(p => { p.hidden = p.dataset.tab !== t; });
  const g = groupOf(t);
  $$("#groupbar .tab").forEach(b => b.classList.toggle("active", b.dataset.group === g.id));
  buildSubTabs(g, t);
  if (t === "pnl") renderPnl();     // portfolio P&L is computed server-side, lazy-loaded
  if (t === "wip") renderWip();
  if (t === "console") renderConsole();
  if (t === "systems") loadSystems();
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
  try { data = await (await fetch("/api/data")).json(); }
  catch (e) { return showError("Could not reach the server: " + e); }
  if (data.error) return showError(data.error);
  $("#errorBanner").hidden = true;
  ALL = data.projects || [];
  ALL.forEach(deriveMetrics);
  AP = data.ap || { summary: {}, lien_watch: [], liens: [], by_project: {}, bills: [] };
  BILLS = AP.bills || [];
  COST = data.cost || { by_code: [], by_project_code: {}, by_project: {}, by_cost_type: [], by_vendor: [], loaded_total: 0 };
  DRAWS = data.draws || { draws: [], total: 0 };
  SALES = data.sales || { pipeline: [], by_rep: [], warm: [], customers: [], totals: {} };
  SUBLOC = data.sub_loc || { summary: null, divisions: {}, projects: [], open_by_project: {}, events: [] };
  OI = data.open_invoices || { as_of: null, buckets: ["Current", "1-30", "31-60", "61-90", "90+"], invoices: [] };
  PAY = data.payments || { payments: [], total_received: 0, count: 0, invoices_paid: 0 };
  PNL = null;   // recompute the portfolio P&L on next open (data just changed)
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
  $("#metaLine").textContent =
    `${meta.project_count} projects · report ${meta.report_date ? fmtDate(meta.report_date) : "—"}` +
    (meta.loaded_at ? ` · loaded ${fmtDate(meta.loaded_at, true)}` : "");
  buildFilterOptions();
  render();
  _renderLazyTab(activeTab);   // wip/payments/paybills read main-load globals but aren't in render();
                               // re-dispatch the active one now that data is in (fixes a fresh refresh on it)
}
// Lazy tabs dispatched by setTab (not render()) that read the /api/data globals. pnl/systems/console
// fetch their OWN data on open, so they self-refresh; these three read ALL / PAY / BILLS synchronously.
function _renderLazyTab(t) {
  const map = { wip: renderWip, payments: renderPayments, paybills: renderPayBills };
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
// Run a pipeline key ('reload' = safe loaders-only default, 'all' = full chain incl
// producers, 'ar'/'ap'/'costs'/'crm'/'wip', 'wip-draft'), driving the given progress
// elements. `els` = { btn, prog, fill, step }.
async function runPipeline(pipeline, confirmMsg, els) {
  if (confirmMsg && !confirm(confirmMsg)) return;
  let res;
  try { res = await (await fetch("/api/sync", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ pipeline, confirm: true }) })).json(); }
  catch (e) { toast("Could not start: " + e); return; }
  if (res.error) { toast(res.error); return; }
  syncing = true; if (els.btn) els.btn.disabled = true;
  els.prog.hidden = false; els.fill.classList.remove("err"); els.fill.style.width = "0%";
  pollSync(res.steps || [], els);
}

async function startResync() {
  runPipeline("reload",
    "Reload all data now?\n\nRuns every loader (WIP, QBO costs [Touch ID], Bill Tracker, Invoices, Customers) and reads the current sources into the ledger. Read-only on the sources; takes about a minute.",
    { btn: $("#btnResync"), prog: $("#syncProgress"), fill: $("#syncBarFill"), step: $("#syncStep") });
}

function pollSync(steps, els) {
  const { btn, prog, fill, step } = els;
  const total = steps.length || 1;
  let fails = 0;
  const tick = () => fetch("/api/sync/status").then(r => r.json()).then(s => {
    fails = 0;
    const done = (s.steps || []).filter(x => x.state === "done").length;
    const cur = (s.steps || [])[s.current];
    fill.style.width = Math.round(done / total * 100) + "%";
    if (s.state === "running") {
      step.textContent = `${cur ? cur.label : "..."} - step ${Math.min(done + 1, total)} of ${total}${s.elapsed ? ` - ${s.elapsed}s` : ""}`;
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
  syncing = false; if (els.btn) els.btn.disabled = false;
  if (reload) { try { await load(true); if (typeof renderConsole === "function" && activeTab === "console") renderConsole(); } catch { /* ignore */ } }
  if (msg) toast(msg);
  if (!/(failed|lost|Lost)/.test(msg)) setTimeout(() => { els.prog.hidden = true; els.fill.style.width = "0%"; }, 2600);
}

// ── Console tab: the control plane. Lists each pipeline (from /api/pipelines) with its
// steps, last-run, and a Run button (a pipeline's Run also fires its real producer).
let PIPELINES = null;
const _consoleEls = () => ({ prog: $("#consoleProgress"), fill: $("#consoleBarFill"), step: $("#consoleStep") });
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
    const steps = document.createElement("div"); steps.className = "pl-steps";
    for (const s of p.steps) {
      const chip = document.createElement("span"); chip.className = "pl-step" + (s.side ? " producer" : "");
      chip.textContent = s.label + (s.side ? " · producer" : ""); steps.appendChild(chip);
    }
    card.appendChild(steps);
    const acts = document.createElement("div"); acts.className = "pl-acts";
    const runBtn = document.createElement("button"); runBtn.className = "btn small"; runBtn.textContent = "Run";
    const sides = p.steps.filter(s => s.side).map(s => s.label);
    const msg = sides.length
      ? `Run the ${p.label} pipeline?\n\nThis fires a REAL sync (${sides.join(", ")}) - writes to the source (Notion / Teams / Excel) and prompts Touch ID - then loads it into the ledger.`
      : `Run the ${p.label} loader?\n\nReads the current source into the ledger (read-only on the source).`;
    runBtn.onclick = () => runPipeline(p.key, msg, { ..._consoleEls(), btn: runBtn });
    acts.appendChild(runBtn);
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
  renderProjects(); renderLiens(); renderVendors(); renderDraws(); renderBills(); renderOpenInvoices(); renderSubLoc(); renderSales(); renderCustomers();
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

function renderHome() {
  // ── data freshness ──
  const fr = meta.freshness || { ledger: {}, sources: {} };
  const items = [
    ["sync-ap (AP bills)", (fr.sources || {})["sync-ap"]],
    ["sync-ar (AR)", (fr.sources || {})["sync-ar"]],
    ["WIP master", (fr.sources || {})["WIP master"]],
    ["Costs loaded (QBO)", (fr.ledger || {})["Costs (QBO)"]],
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
    el.querySelector(".f-when").textContent = fmtDate(when, true);
    el.querySelector(".f-ago").textContent = timeAgo(when);
    if (stale) {
      const b = document.createElement("div"); b.className = "sync-rec"; b.textContent = "⟳ Sync recommended";
      el.appendChild(b);
    }
    box.appendChild(el);
  }
  const freshNote = $("#homeFreshNote");
  if (freshNote) freshNote.textContent = needSync
    ? `— ${needSync} recommended to sync (>48h, weekends aside)` : "— all current";
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
  // ── working on (active projects) ──
  const sel = $("#homeDivision");
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
function qboLinkCell(text, url, title) {
  const td = document.createElement("td"); td.className = "left";
  const label = text || "—";
  if (url) {
    const a = document.createElement("a");
    a.href = url; a.target = "_blank"; a.rel = "noopener"; a.className = "qbo-link";
    a.textContent = label; if (title) a.title = title;
    a.onclick = e => e.stopPropagation();
    td.appendChild(a);
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
function renderDraws() {
  const fv = sel => ($(sel) ? $(sel).value : "").trim().toLowerCase();
  const fProj = fv("#drawFProj"), fVend = fv("#drawFVendor"), fInv = fv("#drawFInv"), fClient = fv("#drawFClient");
  const div = $("#drawDivision") ? $("#drawDivision").value : "";
  const all = (DRAWS.draws || []).filter(d => {                 // each filled field must match (AND)
    if (div && !String(d.project_no || "").toUpperCase().startsWith(div)
            && !(d.label || "").toUpperCase().includes("— " + div)) return false;
    if (fProj && !String(d.project_no || "").toLowerCase().includes(fProj)
              && !(d.label || "").toLowerCase().includes(fProj)) return false;
    if (fInv && !String(d.invoice_no || "").toLowerCase().includes(fInv)) return false;
    if (fVend && !(d.bills || []).some(b => (b.vendor || "").toLowerCase().includes(fVend))) return false;
    // Client = the GC / project name (e.g. "Firestone" catches every Firestone job)
    if (fClient && !(String(d.customer || "") + " " + (d.label || "")).toLowerCase().includes(fClient)) return false;
    return true;
  });
  const shown = activeDrawStage ? all.filter(d => d.stage === activeDrawStage) : all;
  $("#drawsNote").textContent = (DRAWS.draws || []).length
    ? `(${shown.length} shown of ${DRAWS.total} · most recent first)`
    : "(no draw data — run load_bill_tracker.py)";
  // Clickable stage tiles → filter the draw list. Counts come from `all` (all stages);
  // subs spell out the money direction (GC pays us in → we pay vendors out → waivers).
  const stats = [
    ["All paid", "All paid", "GC paid you + vendors paid"],
    ["Collect from GC", "Ready to turn in", "vendors paid, GC still owes"],
    ["Pay vendors", "Fund in — pay vendors", "GC funded, vendors not paid yet"],
    ["Awaiting GC", "Awaiting GC funding", "not funded by the GC yet"],
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
    if (n) el.onclick = () => { activeDrawStage = activeDrawStage === stageKey ? null : stageKey; renderDraws(); };
    sr.appendChild(el);
  }
  { const cb = $("#btnClearDrawStage"); if (cb) cb.hidden = !activeDrawStage; }
  // One row per draw (table). Click a row → its bills open underneath. Green = done =
  // every bill PAID (waivers are tracked per bill but don't gate the color). "Billed
  // (in)" = the GC pays you; "Paid out" = you pay vendors — money-in vs money-out.
  const box = $("#drawList"); box.innerHTML = "";
  if (!shown.length) { box.innerHTML = '<p class="hint" style="padding:14px 18px">No draws match.</p>'; return; }
  // Grouped by project #: sort by project, newest draw first within each project.
  const grouped = [...shown].sort((a, b) =>
    (a.project_no || "").localeCompare(b.project_no || "") ||
    String(b.ar_date || b.recency || "").localeCompare(String(a.ar_date || a.recency || "")));
  const scroll = document.createElement("div"); scroll.className = "table-scroll";
  const table = document.createElement("table"); table.className = "grid draws-table";
  const thead = document.createElement("thead"), tbody = document.createElement("tbody");
  const cols = [["", "left"], ["Draw memo", "left"], ["Billed (in)", "right"], ["Status", "left"],
                ["Invoice #", "left"], ["Date", "left"], ["Paid out", "right"], ["Stage", "left"]];
  const htr = document.createElement("tr");
  for (const [c, al] of cols) { const th = document.createElement("th"); if (al === "left") th.className = "left"; th.textContent = c; htr.appendChild(th); }
  thead.appendChild(htr);
  let curProj = null;
  for (const d of grouped) {
    if (d.project_no !== curProj) {                       // project group header
      curProj = d.project_no;
      const g = grouped.filter(x => x.project_no === curProj);
      const gIn = g.reduce((t, x) => t + (x.billed || 0), 0);
      const gOut = g.reduce((t, x) => t + (x.total || 0), 0);
      const gtr = document.createElement("tr"); gtr.className = "draw-group";
      const gtd = document.createElement("td"); gtd.colSpan = cols.length;
      const sp = document.createElement("span"); sp.className = "g-proj"; sp.textContent = curProj || "—";
      const sub = document.createElement("span"); sub.className = "g-sub";
      const nm = nameOf(curProj);
      sub.textContent = `${nm ? " · " + nm : ""} · ${g.length} draw${g.length > 1 ? "s" : ""} · ${money(gIn)} in / ${money(gOut)} out`;
      gtd.appendChild(sp); gtd.appendChild(sub); gtr.appendChild(gtd); tbody.appendChild(gtr);
    }
    const done = d.stage === "All paid";   // green row only when fully settled (GC paid + vendors paid)
    const open = drawsExpanded.has(d.matched_invoice);
    const tr = document.createElement("tr"); tr.className = "draw-row" + (done ? " done" : "");
    tr.style.cursor = "pointer";
    tr.onclick = (e) => { if (e.target.closest(".cell") || e.target.closest("a")) return;
      open ? drawsExpanded.delete(d.matched_invoice) : drawsExpanded.add(d.matched_invoice); renderDraws(); };
    const cc = document.createElement("td"); cc.className = "left draw-caret"; cc.textContent = open ? "▾" : "▸"; tr.appendChild(cc);
    const memo = (d.label || "").replace(/^\s*\S+\s*—\s*/, "").replace(/^\s*(MFD|CP|RP)\d+(-FTW)?\s*-\s*/i, "").trim() || d.label || "—";
    tr.appendChild(leftText(memo));
    const bt = document.createElement("td");
    if (d.billed != null) { const mc = moneyCell(d.billed); mc.classList.add("draw-in"); bt.appendChild(mc); }
    else bt.appendChild(document.createTextNode("—"));
    tr.appendChild(bt);
    // AR pay status — its own column (green Paid / amber still-owed), not crammed onto the amount
    const stt = document.createElement("td"); stt.className = "left";
    if (d.ar_status) { const s = document.createElement("span"); s.className = d.ar_status === "Paid" ? "ar-paid" : "ar-open"; s.textContent = d.ar_status; stt.appendChild(s); }
    else stt.appendChild(document.createTextNode("—"));
    tr.appendChild(stt);
    const invtd = document.createElement("td"); invtd.className = "left";
    if (d.invoice_no && d.ar_qbo_id) {
      const a = document.createElement("a"); a.href = qboInvoiceUrl(d.ar_qbo_id);
      a.target = "_blank"; a.rel = "noopener"; a.className = "qbo-link"; a.textContent = d.invoice_no;
      a.title = "Open invoice in QuickBooks"; a.onclick = (e) => e.stopPropagation();
      invtd.appendChild(a);
    } else { const s = document.createElement("span"); s.textContent = d.invoice_no || "—"; invtd.appendChild(s); }
    tr.appendChild(invtd);
    tr.appendChild(leftText(fmtDate(d.ar_date || d.recency)));
    const ot = document.createElement("td"); const mo = moneyCell(d.total); mo.classList.add("draw-out"); ot.appendChild(mo);
    const pc = document.createElement("span"); pc.className = "paidcnt"; pc.textContent = ` ${d.paid}/${d.n}`; ot.appendChild(pc); tr.appendChild(ot);
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
  const cap = document.createElement("div"); cap.className = "bills-cap";
  cap.textContent = `${d.n} bills · ${money(d.total)} to vendors · ${d.paid}/${d.n} paid · ${d.waivers}/${d.n} waivers in`;
  wrap.appendChild(cap);
  const scroll = document.createElement("div"); scroll.className = "table-scroll";
  const table = document.createElement("table"); table.className = "grid";
  const thead = document.createElement("thead"), tbody = document.createElement("tbody");
  const cols = [["Vendor", "left"], ["Bill #", "left"], ["Bill date", "left"], ["Amount", "right"], ["Paid", "left"], ["GC funded", "left"], ["Waiver in hand", "left"]];
  const htr = document.createElement("tr");
  for (const [c, al] of cols) { const th = document.createElement("th"); if (al === "left") th.className = "left"; th.textContent = c; htr.appendChild(th); }
  thead.appendChild(htr);
  for (const b of d.bills) {
    const tr = document.createElement("tr");
    tr.appendChild(leftText(b.vendor || "—"));
    tr.appendChild(qboLinkCell(b.bill_ref, qboBillHref(b.qbo_link), "Open this bill in QuickBooks"));
    tr.appendChild(leftText(fmtDate(b.bill_date)));
    const av = document.createElement("td"); av.appendChild(moneyCell(b.amount)); tr.appendChild(av);
    tr.appendChild(leftText(b.pay_date ? "✓ " + fmtDate(b.pay_date) : "—"));
    tr.appendChild(leftText(b.gc_paid ? "✓ " + fmtDate(b.gc_paid) : "—"));
    const wtd = document.createElement("td"); wtd.className = "left";
    const lab = document.createElement("label"); lab.className = "chk";
    const cb = document.createElement("input"); cb.type = "checkbox"; cb.checked = !!b.waiver;
    cb.onchange = () => setWaiver(d, b, cb);
    lab.appendChild(cb); lab.appendChild(document.createTextNode(b.waiver ? " in hand" : " mark"));
    wtd.appendChild(lab); tr.appendChild(wtd);
    tbody.appendChild(tr);
  }
  table.appendChild(thead); table.appendChild(tbody); scroll.appendChild(table); wrap.appendChild(scroll);
  return wrap;
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
    renderDraws();
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
    ? `($${Math.round(total).toLocaleString()} · where the money goes)`
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

function renderVendors() {
  const q = ($("#vendorSearch") ? $("#vendorSearch").value : "").trim().toLowerCase();
  let vends = COST.by_vendor || [];
  if (q) vends = vends.filter(v => (v.vendor || "").toLowerCase().includes(q));
  const totalSpend = vends.reduce((t, v) => t + (v.spend || 0), 0);
  $("#vendorsNote").textContent = (COST.by_vendor || []).length
    ? `(${vends.length} vendors · $${Math.round(totalSpend).toLocaleString()})`
    : "(no cost data — run load_costs.py)";
  const cols = [["Vendor", "left"], ["Type", "left"], ["Spend", "right"], ["Jobs", "right"], ["Lines", "right"]];
  const thead = $("#vendorTable thead"), tbody = $("#vendorTable tbody");
  thead.innerHTML = ""; tbody.innerHTML = "";
  const htr = document.createElement("tr");
  for (const [c, al] of cols) { const th = document.createElement("th"); if (al === "left") th.className = "left"; th.textContent = c; htr.appendChild(th); }
  thead.appendChild(htr);
  const max = vends.reduce((m, v) => Math.max(m, v.spend || 0), 0) || 1;
  const billVendors = new Set((BILLS || []).map(b => b.vendor));   // who has bills we can drill into
  for (const v of vends.slice(0, 150)) {
    const tr = document.createElement("tr");
    if (billVendors.has(v.vendor)) {   // click the row → the Bills tab, filtered to this vendor
      tr.classList.add("row-click"); tr.title = "See this vendor's bills";
      tr.onclick = (e) => { if (e.target.closest(".cell")) return; jumpToVendorBills(v.vendor); };
    }
    tr.appendChild(leftText(v.vendor));
    const ty = document.createElement("td"); ty.className = "left";
    const pill = document.createElement("span"); pill.className = "vtype" + (v.vtype === "Sub" ? " sub" : "");
    pill.textContent = v.vtype || "—"; ty.appendChild(pill); tr.appendChild(ty);
    const st = document.createElement("td");
    const bar = document.createElement("span"); bar.className = "cell bar";
    const fill = document.createElement("span"); fill.className = "bar-fill"; fill.style.width = ((v.spend || 0) / max * 100) + "%";
    const txt = document.createElement("span"); txt.className = "bar-txt"; txt.textContent = money(v.spend);
    bar.appendChild(fill); bar.appendChild(txt); bar.title = "Click to copy"; bar.onclick = () => copy(String(Math.round(v.spend || 0)));
    st.appendChild(bar); tr.appendChild(st);
    tr.appendChild(rightText(String(v.jobs || 0)));
    tr.appendChild(rightText(String(v.lines || 0)));
    tbody.appendChild(tr);
  }
}
function leftText(v) { const td = document.createElement("td"); td.className = "left"; const s = document.createElement("span"); s.textContent = v; td.appendChild(s); return td; }

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
  buildBillDateFilter();
  buildBillVendorFilter();
}
// Month MULTI-select (checkboxes) + a day drill. Clicking a month checks it AND all OLDER
// months ("June and back"); individual priors can then be unchecked (owner 2026-08-20). The
// Day select drills into a single month (Excel-style), enabled only when exactly one is chosen.
let billMonths = new Set();   // selected 'YYYY-MM' (empty = all months)
const _BMONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
function billMonthLabel(ym) { const [y, m] = ym.split("-"); return `${_BMONTHS[+m - 1]} ${y}`; }
function _billMonthsAsc() {   // every month present in the data, oldest → newest
  return [...new Set((BILLS || []).map(b => String(b.bill_date || "").slice(0, 7)).filter(s => /^\d{4}-\d{2}$/.test(s)))].sort();
}
function toggleBillMonth(ym, checked) {
  const asc = _billMonthsAsc();
  if (checked) { for (const m of asc) if (m <= ym) billMonths.add(m); }   // this month + everything older
  else billMonths.delete(ym);                                             // remove just this prior month
  buildBillDateFilter(); renderBills();
}
function buildBillDateFilter() {
  const menu = $("#bfMonthMenu"), btn = $("#bfMonthBtn"), dayEl = $("#bfDay");
  if (!menu || !btn || !dayEl) return;
  const asc = _billMonthsAsc();
  for (const m of [...billMonths]) if (!asc.includes(m)) billMonths.delete(m);   // drop months no longer in the data
  // the checkbox menu, newest first, with a Clear
  menu.innerHTML = "";
  { const clr = document.createElement("button"); clr.type = "button"; clr.className = "msel-clear";
    clr.textContent = "Clear"; clr.onclick = () => { billMonths.clear(); buildBillDateFilter(); renderBills(); }; menu.appendChild(clr); }
  for (const ym of [...asc].reverse()) {
    const lab = document.createElement("label"); lab.className = "msel-opt";
    const cb = document.createElement("input"); cb.type = "checkbox"; cb.checked = billMonths.has(ym);
    cb.onchange = () => toggleBillMonth(ym, cb.checked);
    lab.appendChild(cb); lab.appendChild(document.createTextNode(" " + billMonthLabel(ym)));
    menu.appendChild(lab);
  }
  if (!billMonths.size) btn.textContent = "All months";
  else { const newest = [...billMonths].sort().reverse()[0]; btn.textContent = billMonths.size === 1 ? billMonthLabel(newest) : `${billMonthLabel(newest)} +${billMonths.size - 1}`; }
  btn.classList.toggle("on", billMonths.size > 0);
  // day drill: only when EXACTLY one month is selected
  const prevD = dayEl.value; dayEl.innerHTML = "";
  { const o = document.createElement("option"); o.value = ""; o.textContent = "All days"; dayEl.appendChild(o); }
  if (billMonths.size === 1) {
    const ym = [...billMonths][0];
    const days = [...new Set((BILLS || []).map(b => String(b.bill_date || "")).filter(d => d.slice(0, 7) === ym && /^\d{4}-\d{2}-\d{2}/.test(d)).map(d => d.slice(0, 10)))].sort();
    for (const d of days) { const o = document.createElement("option"); o.value = d; o.textContent = `${_BMONTHS[+ym.slice(5, 7) - 1]} ${+d.slice(8, 10)}`; dayEl.appendChild(o); }
    dayEl.disabled = false;
  } else { dayEl.disabled = true; }
  dayEl.value = prevD; if (dayEl.value !== prevD) dayEl.value = "";
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
  f["#bfMonth"] = billMonths.size ? [...billMonths] : "";
  f["#bfDay"] = $("#bfDay") ? $("#bfDay").value : "";
  f["#bfVendor"] = _vendorNonDefault() ? "1" : "";                                   // vendor deviates from the pump default
  f["#bfMSel"] = BILL_MSEL.some(c => (billMSel[c.id] || {}).size) ? "1" : "";        // any categorical multi-select active
  return f;                                                                          // (drives the "Clear filters" button)
}
function billPassesFilters(b, f) {
  const mo = f["#bfMonth"]; if (mo && mo.length && !mo.includes(String(b.bill_date || "").slice(0, 7))) return false;
  const dy = f["#bfDay"];   if (dy && String(b.bill_date || "").slice(0, 10) !== dy) return false;
  if (billVendorHidden.has(b.vendor || "")) return false;      // vendor multi-select (pumps hidden by default)
  if (!billMSelPasses(b)) return false;                        // Client / Division / Pay / Invoice / Approved / Lien
  return true;
}
function billClearFilters() {
  for (const cfg of BILL_MSEL) (billMSel[cfg.id] || (billMSel[cfg.id] = new Set())).clear();
  billMonths.clear();
  billVendorHidden = new Set(billVendorDefault);   // back to the default (pumps hidden), not "show everything"
  { const d = $("#bfDay"); if (d) d.value = ""; }
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

  // table. Each status is its OWN column (Paid / Invoice / Lien / Appr) so a blank in
  // one never hides a missing value by being merged with the others.
  const group = $("#billGroup") ? $("#billGroup").value : "vendor";
  const thead = $("#billTable thead"), tbody = $("#billTable tbody");
  thead.innerHTML = ""; tbody.innerHTML = "";
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
    td.textContent = bills.length ? "No bills match these filters." : "No AP data - run load_bill_tracker.py.";
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
      const amt = document.createElement("span"); amt.className = "bg-amt";
      amt.textContent = `${money(gOpen)} open · ${g.length} bill${g.length > 1 ? "s" : ""}`;
      cell.appendChild(left); cell.appendChild(amt); gtd.appendChild(cell); gtr.appendChild(gtd);
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
  bar.classList.toggle("dirty", n > 0);   // always present on the Bills tab; amber when unsaved
  $("#billSaveText").textContent = n ? `${n} unsaved lien mark${n > 1 ? "s" : ""}` : "Lien marks saved";
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
    const amt = document.createElement("span"); amt.className = "bg-amt"; amt.textContent = `${list.length} bill${list.length !== 1 ? "s" : ""} · ${money(sub)}`;
    cell.appendChild(key); cell.appendChild(amt); gtd.appendChild(cell); gtr.appendChild(gtd); tbody.appendChild(gtr);
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
  row(gb, "Amount", money(b.line_amount));
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
let invCollapsed = new Set();     // customer groups the owner has collapsed
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
function invPasses(i, f) {
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

function renderOpenInvoices() {
  const host = $("#invTable"); if (!host) return;
  const buckets = OI.buckets || ["Current", "1-30", "31-60", "61-90", "90+"];
  const all = OI.invoices || [];
  if (!$("#ifDivision") || !$("#ifDivision").options.length) buildInvFilters();

  const fv = sel => ($(sel) ? $(sel).value : "");
  const f = { div: fv("#ifDivision"), lien: fv("#ifLien"), lienclk: fv("#ifLienClock"), litig: fv("#ifLitig") || "ex" };  // Client/Project # are msels now
  // Client + Project # multi-selects: build once per data change (signature guard) so a toggle keeps its search.
  const invSig = String(all.length);
  if (invSig !== _invMSelSig || !($("#ifClientMenu") && $("#ifClientMenu").querySelector(".msel-opt"))) {
    _invMSelSig = invSig; for (const cfg of INV_MSEL) buildMSel(cfg, all, invMSel, renderOpenInvoices);
  }
  // Litigation is EXCLUDED by default; flag the box red whenever it's hiding/limiting rows so it's
  // obvious to the eye that a filter is in place (owner 2026-08-19).
  { const el = $("#ifLitig"); if (el) el.classList.toggle("filter-on", (el.value || "ex") !== "all"); }

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
  { const el = $("#invAsOf"); if (el) el.textContent = OI.as_of ? "aged as of " + fmtDate(OI.as_of) : ""; }
  { const anyMsel = INV_MSEL.some(c => (invMSel[c.id] || {}).size);
    const cb = $("#ifClear"); if (cb) cb.hidden = !(anyMsel || f.div || f.lien || f.lienclk || f.litig !== "ex" || invBucketFilter != null); }

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
    const collapsed = invCollapsed.has(k);
    const gOpen = g.reduce((t, x) => t + oiBal(x), 0);
    const gtr = document.createElement("tr"); gtr.className = "bill-group"; gtr.style.cursor = "pointer";
    gtr.title = collapsed ? "Click to expand" : "Click to collapse";
    const gtd = document.createElement("td"); gtd.colSpan = cols.length;
    const cell = document.createElement("div"); cell.className = "bg-cell";   // flex on the div, not the td
    const left = document.createElement("span"); left.className = "bg-left";
    const caret = document.createElement("span"); caret.className = "bg-caret"; caret.textContent = collapsed ? "▸" : "▾";
    const key = document.createElement("span"); key.className = "bg-key"; key.textContent = k;
    left.appendChild(caret); left.appendChild(key);
    const amt = document.createElement("span"); amt.className = "bg-amt";
    amt.textContent = `${money(gOpen)} open · ${g.length} invoice${g.length > 1 ? "s" : ""}`;
    cell.appendChild(left); cell.appendChild(amt); gtd.appendChild(cell); gtr.appendChild(gtd);
    gtr.onclick = () => { if (invCollapsed.has(k)) invCollapsed.delete(k); else invCollapsed.add(k); renderOpenInvoices(); };
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
  tr.appendChild(qboLinkCell(i.doc_number, qboInvoiceUrl(i.qbo_txn_id), "Open this invoice in QuickBooks"));
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
  return tr;
}

function updateInvCollapseBtn() {
  const btn = $("#ifCollapse"); if (!btn) return;
  const allC = invGroupKeys.length && invGroupKeys.every(k => invCollapsed.has(k));
  btn.textContent = allC ? "Expand all" : "Collapse all";
}
function invToggleAll() {
  const allC = invGroupKeys.length && invGroupKeys.every(k => invCollapsed.has(k));
  if (allC) invCollapsed.clear(); else invGroupKeys.forEach(k => invCollapsed.add(k));
  renderOpenInvoices();
}
function invClearFilters() {
  ["#ifDivision", "#ifLien", "#ifLienClock"].forEach(s => { const el = $(s); if (el) el.value = ""; });
  for (const cfg of INV_MSEL) invMSel[cfg.id] = new Set();   // clear Client + Project # multi-selects
  _invMSelSig = null;                                        // force the menus to rebuild (reset checks + label)
  const lt = $("#ifLitig"); if (lt) lt.value = "ex";         // baseline = litigation excluded
  invBucketFilter = null;
  renderOpenInvoices();
}
// A project sub-band inside a client group (indented, lighter than the client band).
function invSubBand(proj, name, open, count, colspan) {
  const tr = document.createElement("tr"); tr.className = "bill-subgroup";
  const td = document.createElement("td"); td.colSpan = colspan;
  const cell = document.createElement("div"); cell.className = "bg-cell";
  const key = document.createElement("span"); key.className = "sg-key"; key.textContent = proj + (name ? " · " + name : "");
  const amt = document.createElement("span"); amt.className = "bg-amt"; amt.textContent = `${money(open)} · ${count} inv`;
  cell.appendChild(key); cell.appendChild(amt); td.appendChild(cell); tr.appendChild(td);
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
    const amt = document.createElement("span"); amt.className = "bg-amt";
    amt.textContent = `${money(divOpen(div))} open · ${rows.length} client${rows.length === 1 ? "" : "s"}`;
    cell.appendChild(key); cell.appendChild(amt); gtd.appendChild(cell); gtr.appendChild(gtd); tb.appendChild(gtr);
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
  hint.innerHTML = "Each row is a <b>payment received</b>. Click it to see the invoices (draws) it paid. <b>Unlocks (AP)</b> is the open vendor bills this payment funds: for staged <b>CP/MFD</b> draws, only the bills on the draw it paid; for <b>RP</b> (costs up front, billed once), the whole job's open AP.";
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
    const allExpanded = pays.every(p => paymentsExpanded.has(p.qbo_txn_id));
    const btn = document.createElement("button"); btn.className = "btn small subtle";
    btn.textContent = allExpanded ? "Collapse all" : "Expand all";
    btn.onclick = () => { if (allExpanded) paymentsExpanded.clear(); else pays.forEach(p => paymentsExpanded.add(p.qbo_txn_id)); renderPayments(); };
    actions.appendChild(btn);
  }
  head.appendChild(actions);
  body.appendChild(head);
  const wrap = document.createElement("div"); wrap.className = "table-scroll";
  const table = document.createElement("table"); table.className = "grid"; table.id = "payTable";
  table.innerHTML = "<thead></thead><tbody></tbody>"; wrap.appendChild(table); body.appendChild(wrap);
  const cols = [["Client", "left"], ["Date", "left"], ["Payment Ref #", "left"], ["Payment Type", "left"], ["Amount Paid", "right"], ["Unlocks (AP)", "right"]];
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
        const gtr = document.createElement("tr"); gtr.className = "bill-group";
        const gtd = document.createElement("td"); gtd.colSpan = cols.length;
        const cell = document.createElement("div"); cell.className = "bg-cell";
        const key = document.createElement("span"); key.className = "bg-key"; key.textContent = per.label;
        const amt = document.createElement("span"); amt.className = "bg-amt"; amt.textContent = `${money(perTot[per.key])} · ${perN[per.key]} payment${perN[per.key] === 1 ? "" : "s"}`;
        cell.appendChild(key); cell.appendChild(amt); gtd.appendChild(cell); gtr.appendChild(gtd); tb.appendChild(gtr);
      }
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
    tr.appendChild(leftText(fmtDateShort(p.txn_date)));
    tr.appendChild(leftText(p.ref_no || "–"));
    tr.appendChild(leftText(p.method || "–"));
    tr.appendChild(rightText(money(p.total_amt)));
    // Unlocks (AP): open vendor bills on this payment's project(s) → click opens the side panel
    const uc = document.createElement("td"); uc.className = "right";
    const bills = payUnlockBills(p, drawIdx, projIdx);
    if (bills.length) {
      const sum = bills.reduce((t, b) => t + num(b.open_balance), 0);
      const link = document.createElement("span"); link.className = "unlock-link";
      link.textContent = `${money(sum)} · ${bills.length}`;
      link.title = "Open vendor bills on the draw(s) this payment paid - the AP it funds";
      link.onclick = (e) => { e.stopPropagation(); openPaymentBills(p, bills); };
      uc.appendChild(link);
    } else uc.appendChild(document.createTextNode("–"));
    tr.appendChild(uc);
    tr.onclick = () => { if (paymentsExpanded.has(p.qbo_txn_id)) paymentsExpanded.delete(p.qbo_txn_id); else paymentsExpanded.add(p.qbo_txn_id); renderPayments(); };
    tb.appendChild(tr);
    // ── grouped invoices this payment paid: Invoice # · Total open · Amount applied ──
    if (expanded) {
      const sr = document.createElement("tr"); sr.className = "pay-invoices";
      const std = document.createElement("td"); std.colSpan = cols.length;
      const box = document.createElement("table"); box.className = "sub-grid";
      const th = document.createElement("thead"); th.innerHTML = "<tr><th class='left'>Invoice #</th><th class='right'>Total open</th><th class='right'>Amount applied</th></tr>";
      box.appendChild(th);
      const bod = document.createElement("tbody");
      if (!p.applications.length) {
        const r = document.createElement("tr"); const c = document.createElement("td"); c.colSpan = 3; c.className = "left dim";
        c.textContent = (p.unapplied_amt || 0) > 0.005 ? "Unapplied - a credit on account, not yet on an invoice." : "No invoice links on this payment.";
        r.appendChild(c); bod.appendChild(r);
      }
      for (const a of p.applications) {
        const r = document.createElement("tr");
        r.appendChild(qboLinkCell(a.invoice_no || ("inv " + a.invoice_txn_id), a.invoice_no ? qboInvoiceUrl(a.invoice_txn_id) : null, "Open this invoice in QuickBooks"));
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
function renderSublocFeed() {
  const box = $("#sublocFeed"); if (!box) return; box.innerHTML = "";
  const repays = (SUBLOC.events || []).filter(e => e.type === "REPAY" && (e.in_amt || 0) > 0.005);
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
      if (i === 0) { td.className = "left"; td.textContent = label; }
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
    const amt = document.createElement("span"); amt.className = "bg-amt"; amt.textContent = `${list.length} job${list.length === 1 ? "" : "s"} · ${money(list.reduce((t, r) => t + num(r.total_contract_price), 0))} contract`;
    cell.appendChild(key); cell.appendChild(amt); gtd.appendChild(cell); gtr.appendChild(gtd); tbody.appendChild(gtr);
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
  const cards = [
    ["Total Contract", money(contract), `${rows.length} jobs`],
    ["Costs to Date", money(costs), contract ? `${(costs / contract * 100).toFixed(0)}% of contract` : ""],
    ["Billed to Date", money(billed), contract ? `${(billed / contract * 100).toFixed(0)}% of contract` : ""],
    ["Left to Bill", money(left), ""],
    ["Net Over/(Under)", money(net), net >= 0 ? "overbilled" : "underbilled"],
    ["Active Jobs", String(active), `of ${rows.length}`],
  ];
  const row = $("#kpiRow"); row.innerHTML = "";
  for (const [label, value, sub] of cards) {
    const el = document.createElement("div"); el.className = "kpi";
    el.innerHTML = `<div class="k-label"></div><div class="k-value"></div><div class="k-sub"></div>`;
    el.querySelector(".k-label").textContent = label;
    el.querySelector(".k-value").textContent = value;
    el.querySelector(".k-sub").textContent = sub;
    row.appendChild(el);
  }
}

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
function moneyCell(v) { const s = document.createElement("span"); s.className = "cell"; s.textContent = money(v); s.onclick = () => copy(String(Math.round(num(v)))); s.title = "Click to copy"; return s; }
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
  const typ = k => ({ money: "money", pct: "pct" }[k] ? { type: k } : { type: "text" });
  for (const [title, fields] of DETAIL_GROUPS) {
    const rows = fields.filter(([k]) => r[k] !== null && r[k] !== undefined && r[k] !== "");
    if (!rows.length) continue;
    const g = document.createElement("div"); g.className = "dgroup";
    const h = document.createElement("h4"); h.textContent = title; g.appendChild(h);
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
  const ap = AP.by_project && AP.by_project[r.project_no];
  if (ap) {
    const g = document.createElement("div"); g.className = "dgroup";
    const h = document.createElement("h4"); h.textContent = "AP / Liens (Bill Tracker)"; g.appendChild(h);
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
    const h = document.createElement("h4"); h.textContent = "Costs (QBO, by code)"; g.appendChild(h);
    const summary = [];
    if (cload) {
      summary.push(["Total loaded", money(cload.costs_loaded)]);
      if (cload.sub_costs) summary.push(["of which subs", money(cload.sub_costs)]);
      if (r.costs_to_date != null) summary.push(["WIP costs_to_date", money(r.costs_to_date)]);
    }
    for (const [label, val] of summary) {
      const row = document.createElement("div"); row.className = "drow";
      const dk = document.createElement("span"); dk.className = "dk"; dk.textContent = label;
      const dv = document.createElement("span"); dv.className = "dv"; dv.textContent = val; dv.title = "Click to copy"; dv.onclick = () => copy(val);
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
    const h = document.createElement("h4"); h.textContent = "Margin (QBO actual)"; g.appendChild(h);
    const mrows = [
      ["Planned markup (contract ÷ ETC)", pct(r.markup_pct)],
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
  { const st = $("#invStatement"); if (st) st.hidden = true; } }

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
function exportCSV() {
  const cols = visibleColumns();
  const rows = filtered();
  const esc = v => { const s = String(v ?? ""); return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; };
  const head = cols.map(c => esc(c.label)).join(",");
  const body = rows.map(r => cols.map(c => esc(raw(c, r[c.key]))).join(",")).join("\n");
  const blob = new Blob([head + "\n" + body], { type: "text/csv" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `ledger_${meta.report_date || "export"}.csv`;
  a.click(); URL.revokeObjectURL(a.href);
  toast(`Exported ${rows.length} rows`);
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

function init() {
  applySettings();
  syncSettingsUI();
  wireSettings();
  ["#search", "#fDivision", "#fStatus", "#fCategory", "#fActive"].forEach(sel =>
    $(sel).addEventListener("input", renderProjects));
  ["#drawFClient", "#drawFProj", "#drawFVendor", "#drawFInv", "#drawDivision"].forEach(sel => { const el = $(sel); if (el) el.addEventListener("input", renderDraws); });
  { const el = $("#vendorSearch"); if (el) el.addEventListener("input", renderVendors); }
  { const el = $("#lienFProj"); if (el) el.addEventListener("input", renderLiens); }   // the other lien filters are multi-selects now
  { const el = $("#wipActive"); if (el) el.addEventListener("change", renderWip); }
  { const el = $("#billSort"); if (el) el.addEventListener("change", renderBills); }
  // Month + Vendor + every categorical multi-select: the button toggles the checkbox menu; a click outside closes it.
  const _mselWraps = [["#bfMonthBtn", "#bfMonthMenu", "#bfMonthMsel"], ["#bfVendorBtn", "#bfVendorMenu", "#bfVendorMsel"],
    ...BILL_MSEL.map(c => [`#${c.id}Btn`, `#${c.id}Menu`, `#${c.id}Msel`]),
    ...LIEN_MSEL.map(c => [`#${c.id}Btn`, `#${c.id}Menu`, `#${c.id}Msel`]),
    ...PAY_MSEL.map(c => [`#${c.id}Btn`, `#${c.id}Menu`, `#${c.id}Msel`]),
    ...INV_MSEL.map(c => [`#${c.id}Btn`, `#${c.id}Menu`, `#${c.id}Msel`])];
  const _closeMsels = (except) => { for (const [, mId] of _mselWraps) { const m = $(mId); if (m && mId !== except) m.hidden = true; } };
  for (const [btnId, menuId, wrapId] of _mselWraps) {
    const btn = $(btnId), menu = $(menuId);
    if (btn && menu) {
      btn.addEventListener("click", (e) => { e.stopPropagation(); const open = menu.hidden; _closeMsels(menuId); menu.hidden = !open; });   // one menu open at a time
      document.addEventListener("click", (e) => { if (!menu.hidden && !e.target.closest(wrapId)) menu.hidden = true; });
    }
  }
  { const d = $("#bfDay"); if (d) d.addEventListener("change", renderBills); }
  { const el = $("#billGroup"); if (el) el.addEventListener("change", () => {
    const grp = el.value;   // re-collapse under the new grouping (collapse stays the default)
    billsCollapsed = grp === "none" ? new Set() : new Set((BILLS || []).map(b => billGroupKey(b, grp)));
    renderBills(); }); }
  { const el = $("#bfClear"); if (el) el.onclick = billClearFilters; }
  { const el = $("#bfCollapse"); if (el) el.onclick = billToggleAll; }
  ["#ifDivision", "#ifLien", "#ifLienClock", "#ifLitig", "#ifSort"].forEach(sel => { const el = $(sel); if (el) el.addEventListener("change", renderOpenInvoices); });
  { const el = $("#ifClear"); if (el) el.onclick = invClearFilters; }
  { const el = $("#ifCollapse"); if (el) el.onclick = invToggleAll; }
  { const el = $("#ifSubGroup"); if (el) el.onclick = invSubGroupToggle; }
  { const el = $("#ifStatement"); if (el) el.onclick = openInvStatement; }
  { const el = $("#btnCopyStmt"); if (el) el.onclick = copyInvStatement; }
  { const el = $("#btnCloseStmt"); if (el) el.onclick = closePanels; }
  { const el = $("#btnCloseBillDetail"); if (el) el.onclick = closePanels; }
  { const el = $("#btnClosePayBills"); if (el) el.onclick = closePanels; }
  { const el = $("#btnCloseLienReview"); if (el) el.onclick = closePanels; }
  { const el = $("#billSaveText"); if (el) el.onclick = openLienReview; }   // press "Lien marks saved" → review them
  { const el = $("#btnCloseSublocDetail"); if (el) el.onclick = closePanels; }
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
  { const el = $("#btnClearDrawStage"); if (el) el.onclick = () => { activeDrawStage = null; renderDraws(); }; }
  { const el = $("#homeDivision"); if (el) el.addEventListener("input", renderHome); }
  $("#btnExport").onclick = exportCSV;
  $("#btnRefresh").onclick = manualRefresh;
  { const el = $("#btnResync"); if (el) el.onclick = startResync; }
  { const el = $("#btnFullRefresh"); if (el) el.onclick = () => runPipeline("all",
      "Full refresh - run EVERY pipeline?\n\nRuns the source producers (AR sync -> Notion/Teams, AP sync -> Bill Tracker.xlsx) AND the loaders, in order. Real writes; expect multiple Touch ID prompts; takes a few minutes.",
      { ..._consoleEls(), btn: el }); }
  $("#btnClearRule").onclick = () => { activeRule = null; renderAttention(); renderProjects(); };
  $("#btnSettings").onclick = () => openPanel("#settings");
  { const el = $("#btnSysReload"); if (el) el.onclick = () => loadSystems(true); }
  for (const id of ["#sysSearch", "#sysOwner", "#sysHealth", "#sysState", "#sysLife", "#sysRetired"]) {
    const el = $(id); if (el) el.addEventListener("input", renderSystems);
  }
  buildGroupBar();   // generate the two-level nav (top groups; sub-tabs render on setTab)
  let savedTab = "home";
  try { savedTab = localStorage.getItem("proficient-ledger-tab") || "home"; } catch { /* ignore */ }
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
