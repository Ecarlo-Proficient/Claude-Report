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

// Derived per-job metrics from the REAL QBO costs - computed once at load time.
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
  // PLANNED markup (on cost) vs PLANNED margin (on revenue) - never the same number.
  r.markup_pct    = (etc && contract != null) ? (contract - etc) / etc : null;
  r.margin_pct    = (contract && contract !== 0) ? ((contract - num(etc)) / contract) : null;
  // ACTUAL markup from real QBO cost (how much we marked the true cost up).
  r.actual_markup_pct = (has && cost) ? (billed - cost) / cost : null;
}

// ── Settings ──────────────────────────────────────────────────────────────
const LS_KEY = "proficient-ledger-settings-v1";
// Two settings, one owner (2026-09-13: accent, font, text size, width, widget sizes and per-table
// columns are gone - "one look, one owner"). Light by default (owner 2026-09-02); a saved choice wins.
const DEFAULTS = { theme: "light", density: "comfortable" };

// Lien state → urgency css class (most urgent first), for the AP watchlist.
const LIEN_CLASS = {
  "Notice PAST due": "past", "Notice due in ≤7d": "d7", "Notice due in ≤15d": "d15",
  "Notice due in ≤30d": "d30", "Notice Sent": "info", "Lien Filed": "info",
};

// Budget adherence - the ONE rule the whole dashboard flags "over budget" with.
// Flatwork (-FTW) budgets are a SOFT reference, not a strict target: the ops
// manager just sends a sub and charges by the labor it took (flatwork is simple
// next to slab), so the estimator's FTW budget is a starting point, not a
// must-hit number the way slab is - EXCEPT on a big flatwork job (~$15k+), which
// does need to hold its budget. Slab / CP / MFD stay strict.
const FTW_BUDGET_FLOOR = 15000;
function budgetCost(r) { return r.costs_loaded != null ? r.costs_loaded : r.costs_to_date; }
function isOverBudget(r) {
  if (r.over_budget_accepted) return false;      // a standing ruling (job_rulings.json): the owner knows why - not a flag
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
    hint: "earned ahead of billed - could invoice",
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

function loadSettings() {
  try {
    const s = JSON.parse(localStorage.getItem(LS_KEY)) || {};
    return { theme: ["auto", "light", "dark"].includes(s.theme) ? s.theme : DEFAULTS.theme,
             density: s.density === "compact" ? "compact" : DEFAULTS.density };
  } catch { return { ...DEFAULTS }; }
}
function saveSettings() { localStorage.setItem(LS_KEY, JSON.stringify(settings)); }

function applySettings() {
  const root = document.documentElement;
  const dark = settings.theme === "dark" ||
    (settings.theme === "auto" && matchMedia("(prefers-color-scheme: dark)").matches);
  root.setAttribute("data-theme", dark ? "dark" : "light");
  root.style.setProperty("--row-pad", settings.density === "compact" ? "5px 10px" : "10px 12px");
  root.style.setProperty("--row-pad-dense", settings.density === "compact" ? "3px 8px" : "5px 9px");   // the dense grids follow the same setting
}


// ── Groups: every row that has a parent collapses under it (owner 2026-09-08: "EVERY SINGLE row that has a
// parent must be grouped - total data first, the detail on a click"). One delegated click handler and one
// MutationObserver decorate every group header the renderers emit - a caret, a remembered open/closed state
// (per header text, in localStorage), children hidden while closed. Headers that already run their own
// toggle (an `onclick` of their own) are left alone. Kinds and their default state:
const GRP_KINDS = [
  { sel: "tr.bill-group", kind: "band", open: false },        // vendor / client / division bands in tables
  { sel: "tr.vp-pay", kind: "sib", open: false, until: "tr.vp-pay" },   // vendor page Payments: a payment over the bills it paid (owner 2026-09-22)
  { sel: "tr.qa-pay", kind: "sib", open: false, until: "tr:not(.qa-child)" },   // QBO changes: a check over the bills it lost (owner 2026-09-24: "it's just too much")
  { sel: "tr.sys-group", kind: "band", open: false },         // Systems: a domain
  { sel: "tr.tr-msum", kind: "band", open: false },           // the money trail: a month
  { sel: ".pnl-codegrp", kind: "sib", open: false, until: ".pnl-codegrp" },   // P&L: a job type over its cost codes
  { sel: ".pnl-invcap", kind: "sib", open: false, until: ".pnl-cap" },
  { sel: ".pnl-codecap", kind: "sib", open: false, until: ".pnl-cap" },       // P&L: "Costs by code" folded like the invoices (owner 2026-09-23)         // P&L: "Invoices - all draws" over the list (closed by default, owner 2026-09-23)
  { sel: ".wr-div-head", kind: "sib", open: true, until: ".wr-div-head" },    // WIP Review: a division (approval work - open)
  { sel: "tr.wr-lines-head", kind: "sib", open: true, until: "tr:not(.wr-line)" },   // WIP Review slide: the QBO lines under a Costs / Billed change (owner 2026-09-23)
  { sel: ".dgroup > h4", kind: "parent", open: true },        // record panels: a titled block (content - open)
  { sel: ".ip-sec.fold-sec > .widget-head", kind: "parent", open: false },   // project page: Audit findings + Change log start folded (owner 2026-09-23)
  { sel: ".warm-head", kind: "parent", open: false },         // Sales: an account card
  { sel: ".pl-head", kind: "parent", open: true },            // Console: a pipeline card (has the buttons - open)
];
const GRP_LS = "proficient-ledger-groups-v1";
let _grpState = (() => { try { return JSON.parse(localStorage.getItem(GRP_LS)) || {}; } catch { return {}; } })();
const GRP_HEADER_TR = "tr.bill-group, tr.sys-group, tr.tr-msum, tr.tr-total, tr.pp-sect, tr.bill-subgroup, tr.inv-client, tr.sumRow";
function _grpChildren(h, def) {
  if (def.kind === "band") { const out = []; let e = h.nextElementSibling; while (e && !e.matches(GRP_HEADER_TR)) { out.push(e); e = e.nextElementSibling; } return out; }
  if (def.kind === "sib") { const out = []; let e = h.nextElementSibling; while (e && !e.matches(def.until)) { out.push(e); e = e.nextElementSibling; } return out; }
  return [...h.parentElement.children].filter(c => c !== h);   // parent: everything in the card but the head
}
function _grpKeyOf(h, def) {   // keyed on the NAME (not the caret, not the amounts that change every sync) so the choice sticks
  if (h.dataset && h.dataset.grpkey) return `${def.sel}|${h.dataset.grpkey}`;   // a row with its own stable key (payments: the QBO id)
  const nm = h.querySelector(".bg-key, .sg-key, .g-cust, .warm-name, .pl-name") || h.querySelector("td, span") || h;
  const t = (nm.textContent || "").replace(/[▸▾\u25B6\u25BC\uFE0E]/g, "").replace(/\s+/g, " ").trim().slice(0, 80);
  return `${def.sel}|${typeof activeTab === "undefined" ? "" : activeTab}|${t}`;
}
function _grpApply(h, def, open) {
  h.classList.toggle("grp-closed", !open);
  const c = h.querySelector(":scope .grp-caret") || h.querySelector(".grp-caret"); if (c) c.textContent = open ? "\u25BC\uFE0E" : "\u25B6\uFE0E";   // full-size triangles, text style (owner 2026-09-23: "the carrot bigger")
  for (const el of _grpChildren(h, def)) el.hidden = !open;
}
function _grpDecorate(root) {
  for (const def of GRP_KINDS) {
    const list = root.matches && root.matches(def.sel) ? [root] : [];
    list.push(...root.querySelectorAll(def.sel));
    for (const h of list) {
      if (h.dataset.grp) continue;
      if (h.onclick || (h.closest("tr") && h.closest("tr").onclick && def.kind === "band")) { h.dataset.grp = "own"; continue; }   // renders its own toggle
      h.dataset.grp = def.kind; h.dataset.grpSel = def.sel; h.classList.add("grp-h");
      const caret = document.createElement("span"); caret.className = "grp-caret";
      let slot = null;   // the caret goes INSIDE the name, never as a new grid / flex item that would shift the columns
      for (const q of [".bg-left", ".bg-key", ".sg-key", ".g-cust", ".warm-name", ".pl-name", "td", "h2"]) { slot = h.querySelector(q); if (slot) break; }
      if (!slot) slot = (h.firstElementChild && h.firstElementChild.tagName === "SPAN") ? h.firstElementChild : h;
      slot.insertBefore(caret, slot.firstChild);
      const k = _grpKeyOf(h, def); const open = (k in _grpState) ? !!_grpState[k] : def.open;
      _grpApply(h, def, open);
    }
  }
}
document.addEventListener("click", (e) => {
  const h = e.target.closest(".grp-h"); if (!h || h.dataset.grp === "own") return;
  if (e.target.closest("a, input, button, select, textarea, label, .qbo-ico, .att-btn")) return;
  const def = GRP_KINDS.find(d => d.sel === h.dataset.grpSel); if (!def) return;
  const open = h.classList.contains("grp-closed");
  _grpApply(h, def, open);
  _grpState[_grpKeyOf(h, def)] = open;
  try { localStorage.setItem(GRP_LS, JSON.stringify(_grpState)); } catch { /* private window */ }
});
if (window.MutationObserver && !window.__grpObs) {
  window.__grpObs = new MutationObserver((muts) => { for (const m of muts) for (const n of m.addedNodes) if (n.nodeType === 1) _grpDecorate(n); });
  window.__grpObs.observe(document.body, { childList: true, subtree: true });
  _grpDecorate(document.body);
}

// ── Formatting ────────────────────────────────────────────────────────────
const isNum = v => typeof v === "number" && !Number.isNaN(v);
function money(v) {
  if (v != null && v !== "" && !Number.isNaN(Number(v)) && Math.abs(Number(v)) > 0 && Math.abs(Number(v)) < 0.5) { const n = Number(v); const s = "$" + Math.abs(n).toFixed(2); return n < 0 ? `(${s})` : s; }   // cents, never a misleading $0
  if (v === null || v === undefined || v === "") return "–";
  const n = Number(v); if (Number.isNaN(n)) return "–";
  const s = "$" + Math.round(Math.abs(n)).toLocaleString();
  return n < 0 ? "(" + s + ")" : s;      // negatives like Excel: ($28,067), coloured red by .neg (owner 2026-09-01)
}
function pct(v) {
  if (v === null || v === undefined || v === "") return "–";
  const n = Number(v); if (Number.isNaN(n)) return "–";
  return (n * 100).toFixed(1) + "%";
}
function fmt(col, v) {
  if (col.type === "money") return money(v);
  if (col.type === "pct")   return pct(v);
  return (v === null || v === undefined || v === "") ? "–" : String(v);
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

// ── Nav (owner 2026-09-23: "remove Vendors and Customers from Company and reinstate them as their own"):
// Projects = the WIP as the page · Vendors = Bill Tracker / Vendor Center · Customers = Invoice Tracker / Customer
// Center / Payments received / Sales pipeline · Company = Money / QBO Audit · a tools group only the gear reaches.
// Every old tab name is an alias that lands on its page (and scrolls to its section where it became one), so every
// deep link in the app still works.
const NAV_GROUPS = [
  { id: "projects",  label: "Projects",  tabs: ["projects"] },
  { id: "vendors",   label: "Vendors",   tabs: ["bills", "vendorcenter"] },
  { id: "customers", label: "Customers", tabs: ["invoices", "customercenter", "payments", "sales"] },
  { id: "company",   label: "Company",   tabs: ["money", "billaudit", "qboaudit", "checkdrift", "uncleared"] },   // each audit is its own page (owner 2026-09-23 / 09-24)
  { id: "tools",     label: "Tools",     tabs: ["wipreview", "review", "console", "systems"], hidden: true },   // from the gear, not the bar
];
const TAB_LABELS = {
  projects: "Projects", bills: "Bill Tracker", vendorcenter: "Vendor Center", invoices: "Invoice Tracker", customercenter: "Customer Center",
  payments: "Payments received", sales: "Sales pipeline", money: "Money", billaudit: "Bills to fix", qboaudit: "QBO changes", checkdrift: "Checks QBO changed", uncleared: "Uncleared checks",
  wipreview: "WIP Review", review: "WIP review", console: "Console", systems: "Systems", paybills: "Pay run", liens: "Lien register",
};
const HIDDEN_TAB_GROUP = { paybills: "vendors", liens: "vendors" };   // pages without a sub-tab (opened from the Bill Tracker): the Vendors group stays lit
// old tab -> the page it lives on now (where the old tab became a section, its id is the old name and setTab scrolls to it)
const TAB_ALIAS = { overview: "projects", home: "projects", wip: "projects", pnl: "projects", draws: "projects",
                    clients: "invoices", customers: "customercenter", vendors: "vendorcenter", accounting: "billaudit",
                    health: "money", subloc: "money", costs: "money", graph: "systems",
                    rpreview: "review" };   // the RP review became the per-division WIP review (2026-09-15)
const KNOWN_TABS = new Set([...NAV_GROUPS.flatMap(g => g.tabs), ...Object.keys(HIDDEN_TAB_GROUP)]);
const groupOf = t => NAV_GROUPS.find(g => g.tabs.includes(t)) || NAV_GROUPS.find(g => g.id === HIDDEN_TAB_GROUP[t]) || NAV_GROUPS[0];
function buildGroupBar() {
  const bar = $("#groupbar"); if (!bar) return; bar.innerHTML = "";
  for (const g of NAV_GROUPS) {
    if (g.hidden) continue;
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
let activeTab = "projects";
function setTab(t) {
  if (typeof _ppLeaveBlocked === "function" && _ppLeaveBlocked()) return;   // an unsaved pay run on the project page: Save or Discard first
  const sec = TAB_ALIAS[t] ? t : null;          // an old tab name: land on its page, then scroll to its section
  if (sec) t = TAB_ALIAS[t];
  if (!KNOWN_TABS.has(t)) t = "projects";
  activeTab = t;
  try { localStorage.setItem("proficient-ledger-tab", t); } catch { /* ignore */ }
  { const rv = $("#recordView"); if (rv) { if (!rv.hidden) _recSave(null); rv.hidden = true; } }   // leaving a record view when a tab is picked (a refresh no longer reopens it)
  $$(".tab-page").forEach(p => { p.hidden = p.dataset.tab !== t; });
  const g = groupOf(t);
  $$("#groupbar .tab").forEach(b => b.classList.toggle("active", b.dataset.group === g.id));
  buildSubTabs(g, t);
  { const ch = $("#companyHead"); if (ch) ch.hidden = g.id !== "company"; }
  if (t === "projects") renderPnl();          // the P&L by job fold (server-computed, cached until the next load)
  if (t === "customercenter") renderCustomers();
  if (t === "payments") renderPayments();
  if (t === "billaudit") loadAccounting();
  if (t === "qboaudit") loadQboAudit();
  if (t === "checkdrift") loadCheckDrift();
  if (t === "uncleared") loadUncleared();
  if (t === "money") { loadHealth(); renderPnl(); }
  if (t === "wipreview") loadWipReview();
  if (t === "review") loadReview();
  if (t === "console") renderConsole();
  if (t === "systems") loadSystems();
  if (t === "paybills") renderPayBills();
  if (typeof _csClear === "function") _csClear();   // drop any cell selection when the tab changes
  window.scrollTo(0, 0);
  if (sec) { const el = document.getElementById("sec-" + sec); if (el) { if (el.classList.contains("fold")) foldSet(el, true); requestAnimationFrame(() => el.scrollIntoView({ behavior: "smooth", block: "start" })); } }
}
// ── Folds: a section that opens on its head (Payments, Sales, Audit, Costs by code, Sub LOC, P&L by job).
// Closed by default, remembered per section, and opened by any deep link that lands on it.
const FOLD_LS = "proficient-ledger-folds-v1";
let _folds = (() => { try { return JSON.parse(localStorage.getItem(FOLD_LS)) || {}; } catch { return {}; } })();
function foldSet(el, open) {
  el.classList.toggle("closed", !open);
  const c = el.querySelector(":scope > .fold-head .fold-caret"); if (c) c.textContent = open ? "▾" : "▸";
  _folds[el.id] = open; try { localStorage.setItem(FOLD_LS, JSON.stringify(_folds)); } catch { /* ignore */ }
}
function initFolds() {
  for (const el of $$(".subpage.fold")) foldSet(el, _folds[el.id] === true);
  document.addEventListener("click", e => { const h = e.target.closest(".fold-head"); if (!h) return; const el = h.parentElement; foldSet(el, el.classList.contains("closed")); });
}
let PNL = null;                       // cached /api/pnl/portfolio result (invalidated on reload)
let pnlSort = { key: "net", dir: 1 }; // net ascending = worst margin first
let pnlExpanded = new Set();          // P&L jobs expanded inline (instead of the side panel)
const nameOf = pn => (ALL.find(r => r.project_no === pn) || {}).project_name || "";
// A project is "active" if its WIP status is Active - OR blank, which only happens for
// MFD (Test-Master carries no STATUS column, so its jobs are active by construction).
// Matches _portfolio_pnl on the server so every "active" count agrees, MFD included.
const isActive = r => ["", "active"].includes((r.status || "").toLowerCase());
let meta = {};
let activeLien = null;   // lien stage currently filtering the Liens table (null = all)
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
    const bgrp = $("#billGroup") ? $("#billGroup").value : "none";
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
  const bgrp = $("#billGroup") ? $("#billGroup").value : "none";   // a flat list by default (owner 2026-09-22); a chosen grouping opens collapsed
  billsCollapsed = bgrp === "none" ? new Set() : new Set((BILLS || []).map(b => billGroupKey(b, bgrp)));
  buildFilterOptions();
  render();
  _renderLazyTab(activeTab);
}
// Lazy tabs dispatched by setTab (not render()) that read the /api/data globals. pnl/systems/console
// fetch their OWN data on open, so they self-refresh; these three read ALL / PAY / BILLS synchronously.
function _renderLazyTab(t) {
  const map = { clients: renderPayments, paybills: renderPayBills, money: () => loadHealth(true) };
  if (map[t]) map[t]();
}
function showError(msg) {
  const b = $("#errorBanner"); b.hidden = false; b.textContent = msg;
  $("#metaLine").textContent = "not loaded";
}

// Manual "Refresh" - re-reads the ledger DB and re-renders, WITH feedback so it's
// obvious it did something (the silent 90s auto-refresh does the same in the
// background). It does NOT re-pull QBO/Excel - that's a sync (the loaders); the
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
  if (typeof els.after === "function") { try { els.after(reload); } catch { /* ignore */ } }   // the page that started the run re-reads its own feed
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
  if ($("#salesStage")) fillSelect("#salesStage", uniq((SALES.customers || []).map(c => c.sales_status)));
  if ($("#salesDivision")) fillSelect("#salesDivision", uniq((SALES.customers || []).map(c => c.division)));
}
const uniq = arr => [...new Set(arr.filter(Boolean))].sort();
function fillSelect(sel, values) {
  const el = $(sel); const keep = el.firstElementChild;
  el.innerHTML = ""; el.appendChild(keep);
  for (const v of values) { const o = document.createElement("option"); o.value = v; o.textContent = v; el.appendChild(o); }
}
// ── Render ────────────────────────────────────────────────────────────────
function render() {
  renderSync();
  renderKPIs(); renderCosts(); renderDivisions();
  renderProjects(); renderLiens(); renderVendors(); renderBills(); renderOpenInvoices(); renderSubLoc(); renderSales(); renderCustomers();
  renderCompanyHead();
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
// data - nobody syncs on the weekend, so a Friday load isn't "stale" on Monday).
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
// The ONE place load / sync times live (owner 2026-09-23: "remove ALL loads/sync data in the actual ledger and simply keep
// that status where it belongs in the top, break it out, show qbo status"). Grouped by system, a dot per feed
// (green = within 48 business hours, amber = older, grey = never), the time and how long ago.
const SYNC_GROUPS = [
  ["QuickBooks", [["Mirror (every QBO read)", "S", "QBO mirror"], ["Costs", "L", "Costs (QBO)"], ["Invoices (AR)", "L", "AR (invoices)"],
                  ["Payments received", "L", "Payments"], ["Company health", "L", "Health (QBO)"], ["Sub LOC", "L", "Sub LOC"]]],
  ["Bill Tracker (AP)", [["Bill Tracker file", "S", "sync-ap"], ["Loaded into the ledger", "L", "AP (Bill Tracker)"]]],
  ["Invoice Tracker (AR)", [["Invoice file", "S", "sync-ar"]]],
  ["WIP", [["WIP master file", "S", "WIP master"], ["Loaded into the ledger", "L", "WIP"]]],
  ["Notion", [["Customer List (CRM)", "L", "CRM (customers)"]]],
];
function toggleSyncPop() {
  let pop = $("#syncPop");
  if (pop && !pop.hidden) { pop.hidden = true; return; }
  if (!pop) { pop = document.createElement("div"); pop.id = "syncPop"; pop.className = "sync-pop"; document.body.appendChild(pop);
    document.addEventListener("click", (e) => { if (!pop.hidden && !pop.contains(e.target) && e.target.closest("#syncPill") == null) pop.hidden = true; });
    document.addEventListener("keydown", (e) => { if (e.key === "Escape") pop.hidden = true; }); }
  const fr = meta.freshness || { ledger: {}, sources: {} }, S = fr.sources || {}, L = fr.ledger || {};
  const at = w => { if (!w) return null; const t = Date.parse(w.length <= 16 ? w + ":00" : w); return isNaN(t) ? null : t; };
  let html = `<div class="sp-head"><b>Data status</b><span>green = current · amber = over 2 business days</span></div>`;
  for (const [sys, feeds] of SYNC_GROUPS) {
    html += `<div class="sp-sys">${_ge(sys)}</div>`;
    for (const [lbl, kind, key] of feeds) {
      const w = (kind === "S" ? S : L)[key], t = at(w);
      const cls = !t ? "none" : businessHoursSince(t, Date.now()) > STALE_BUSINESS_H ? "stale" : "ok";
      html += `<div class="sp-row"><span class="sp-dot ${cls}"></span><span class="sp-l">${_ge(lbl)}</span>`
        + `<span class="sp-t">${w ? _ge(fmtDate(w, true)) : "never"}</span><span class="sp-ago">${w ? _ge(timeAgo(w)) : ""}</span></div>`;
    }
  }
  html += `<div class="sp-foot"><button class="btn small" id="spOpenGear" type="button">Sync settings and run buttons</button></div>`;
  pop.innerHTML = html; pop.hidden = false;
  const r = $("#syncPill").getBoundingClientRect();
  pop.style.top = Math.round(r.bottom + 6 + window.scrollY) + "px";
  pop.style.right = Math.max(12, Math.round(document.documentElement.clientWidth - r.right)) + "px";
  $("#spOpenGear").onclick = () => { pop.hidden = true; openPanel("#settings"); };
}
function renderSync() {
  renderSyncPill();
  // ── data freshness (the gear panel) ──
  const fr = meta.freshness || { ledger: {}, sources: {} };
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
  const box = $("#homeFresh"); if (!box) return; box.innerHTML = "";
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
    if (stale) { const b = document.createElement("div"); b.className = "sync-rec"; b.textContent = "⟳ Sync recommended"; el.appendChild(b); }
    box.appendChild(el);
  }
  const freshNote = $("#homeFreshNote");
  if (freshNote) freshNote.textContent = needSync ? `- ${needSync} recommended to sync (over 48h, weekends aside)` : "";
}

const DRAW_STAGE_CLASS = {
  "Fund in - pay vendors": "d7",
  "Awaiting GC funding": "info",
  "Ready to turn in": "d7",   // vendors paid, GC still owes -> amber (collect)
  "All paid": "ready",        // GC paid + vendors paid -> green (done)
};
// Clearer, direction-explicit pill text (who paid whom). Display only - the internal
// stage keys above are unchanged (they're matched in several places).
const DRAW_STAGE_LABEL = {
  "Fund in - pay vendors": "GC funded → pay vendors",
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
// Every transaction row gets the same 📎 (owner 2026-09-03: "every transaction needs an attachment
// that we can view on the ledger and open"): `attBtn(type, id, n)` shows the count from the ledger's
// attachment index and opens the viewer - fresh QBO links are fetched on click (they expire in minutes).
function attBtn(type, id, n, title, ctx) {   // ctx = {items: [{type, id, n, title}], index} - the bills to flip through in the viewer
  const b = document.createElement("button"); b.type = "button"; b.className = "att-btn" + (n ? "" : " none");
  b.textContent = n > 1 ? "📎" + n : "📎"; b.title = n ? `${n} attachment${n === 1 ? "" : "s"} in QuickBooks - click to view the bill and its scan` : "No attachment on file in QuickBooks";
  if (id && (n || (ctx && ctx.items && ctx.items.length))) { b.disabled = false; b.onclick = (e) => { e.stopPropagation(); if (ctx && ctx.items && ctx.items.length) openBillViewer(ctx.items, ctx.index || 0); else openBillViewer([{ type, id: String(id), n, title: title || "" }], 0); }; }
  else b.disabled = true;
  return b;
}
// The bill VIEWER (owner 2026-09-16: "a bigger version of the bill info so we see all and can flip through multiple
// bills without clicking out of it"): the bill's header, every line (description · cost code · project · amount), the
// Bill Tracker's pay / invoice / approval / lien state, and the scan beside it; ← → flip through the list handed in.
let _bv = null;
async function openBillViewer(items, index) {
  _bv = { items, i: Math.max(0, Math.min(index || 0, items.length - 1)) };
  let ov = $("#attViewer");
  if (!ov) { ov = document.createElement("div"); ov.id = "attViewer"; ov.className = "xdlg-ov"; document.body.appendChild(ov); ov.onclick = (e) => { if (e.target === ov) { ov.remove(); _bv = null; } }; }
  if (!window.__bvKeys) { window.__bvKeys = true; document.addEventListener("keydown", (e) => { if (!_bv || !$("#attViewer")) return; if (e.key === "ArrowRight") _bvGo(1); else if (e.key === "ArrowLeft") _bvGo(-1); else if (e.key === "Escape") { $("#attViewer").remove(); _bv = null; } else return; e.preventDefault(); }); }
  _bvRender();
}
function _bvGo(step) { if (!_bv) return; const n = _bv.i + step; if (n < 0 || n >= _bv.items.length) return; _bv.i = n; _bvRender(); }
async function _bvRender() {
  const ov = $("#attViewer"); if (!ov || !_bv) return;
  const it = _bv.items[_bv.i], many = _bv.items.length > 1;
  ov.innerHTML = `<div class="xdlg att-dlg bv-dlg" role="dialog">
    <div class="att-head"><button class="btn small" id="bvPrev" ${_bv.i === 0 ? "disabled" : ""} title="Previous bill (←)">←</button><span class="bv-pos">${many ? `${_bv.i + 1} of ${_bv.items.length}` : ""}</span><button class="btn small" id="bvNext" ${_bv.i >= _bv.items.length - 1 ? "disabled" : ""} title="Next bill (→)">→</button>
      <h3 id="bvTitle">${_ge(it.title || (it.type + " " + it.id))}</h3><span class="dim" id="attStatus">loading…</span><button class="btn small" id="attClose">Close</button></div>
    <div class="bv-body"><div class="bv-info" id="bvInfo"><div class="tr-note">Loading the bill…</div></div><div class="bv-scan"><div class="att-list" id="attList"></div><div class="att-view" id="attView"><div class="tr-note">Fetching the scan from QuickBooks…</div></div></div></div></div>`;
  $("#attClose").onclick = () => { ov.remove(); _bv = null; };
  $("#bvPrev").onclick = () => _bvGo(-1); $("#bvNext").onclick = () => _bvGo(1);
  const mine = _bv.i;
  // the bill's info from the ledger, and the scan from QuickBooks, side by side
  fetch(`/api/bill/info?id=${encodeURIComponent(it.id)}`).then(r => r.json()).then(b => {
    if (!_bv || _bv.i !== mine) return;
    const host = $("#bvInfo"); if (!host) return; host.innerHTML = "";
    if (!b || !b.ok) { host.innerHTML = `<div class="tr-note">${_ge((b && b.error) || "no bill info")}</div>`; return; }
    $("#bvTitle").textContent = `${b.vendor || ""} · bill ${b.bill_ref || it.id}`;
    const kv = document.createElement("div"); kv.className = "bv-kv";
    const add = (k, v, cls) => { if (v == null || v === "") return; const r = document.createElement("div"); r.className = "drow"; const a = document.createElement("span"); a.className = "dk"; a.textContent = k; const c = document.createElement("span"); c.className = "dv" + (cls ? " " + cls : ""); if (v instanceof Node) c.appendChild(v); else c.textContent = v; r.appendChild(a); r.appendChild(c); kv.appendChild(r); };
    add("Vendor", b.vendor); add("Bill #", b.bill_ref); add("Date", b.date ? fmtDate(b.date) : null); add("Amount", money(b.total), "bv-big");
    if (b.open != null) add("Open", money(b.open), num(b.open) > 0.005 ? "neg" : "pos");
    add("Paid", b.pay_date ? "Paid " + fmtDate(b.pay_date) : (b.pay_status || (b.is_sub ? "see QuickBooks (sub bill)" : null)));
    if (b.invoice_no) { const s = document.createElement("span"); s.textContent = `Invoice ${b.invoice_no}`; if (b.invoice) { const paid = (b.invoice.balance || 0) <= 0.005; s.appendChild(document.createTextNode(" · ")); s.appendChild(stText(paid ? "GC paid" : "GC owes " + money(b.invoice.balance), paid ? "st-ok" : "st-warn")); } add("On invoice", s); }
    if (b.invoice_status) add("Tracker", b.invoice_status);
    if (b.approved) add("Approved", b.approved === "approved" ? "Yes" : b.approved);
    if (b.lien_status) add("Lien", b.lien_status);
    add("Project", (b.projects || []).join(", ") || null); add("Client", (b.clients || []).join(", ") || null);
    if (b.memo) add("Memo", b.memo);
    host.appendChild(kv);
    const t = document.createElement("table"); t.className = "grid bv-lines";
    t.innerHTML = "<thead><tr><th class='left'>Line item</th><th class='left'>Cost code</th><th class='left'>Project</th><th class='right'>Amount</th></tr></thead>";
    const tb = document.createElement("tbody");
    for (const ln of (b.lines || [])) { const tr = document.createElement("tr");
      tr.appendChild(leftText(ln.description || "–"));
      { const cc = document.createElement("td"); cc.className = "left"; if (ln.cost_code) { const ch = document.createElement("span"); ch.className = "codechip"; ch.textContent = ln.cost_code; cc.appendChild(ch); } else { cc.textContent = ln.account ? ln.account.split(":").pop().trim() : "–"; cc.classList.add("dim"); } tr.appendChild(cc); }
      tr.appendChild(leftText(ln.project_no || "–"));
      tr.appendChild(rightText(money(ln.amount)));
      tb.appendChild(tr); }
    const tot = document.createElement("tr"); tot.className = "bv-tot"; tot.innerHTML = `<td class="left" colspan="3">${(b.lines || []).length} line${(b.lines || []).length === 1 ? "" : "s"}</td><td class="right">${_ge(money((b.lines || []).reduce((s, l) => s + num(l.amount), 0)))}</td>`; tb.appendChild(tot);
    t.appendChild(tb); host.appendChild(t);
    const a = document.createElement("a"); a.className = "btn small"; a.href = qboUrl(b.txn_type === "Expense" ? "expense" : "bill", it.id); a.target = "_blank"; a.rel = "noopener"; a.textContent = "Open in QuickBooks ↗"; host.appendChild(a);
  }).catch(() => { const host = $("#bvInfo"); if (host) host.innerHTML = `<div class="tr-note">could not load the bill</div>`; });
  let r;
  try { r = await (await fetch(`/api/attachment?id=${encodeURIComponent(it.id)}&type=${encodeURIComponent(it.type || "Bill")}`)).json(); }
  catch (e) { r = { ok: false, error: String(e) }; }
  if (!_bv || _bv.i !== mine) return;
  const files = (r && r.files) || [];
  const st = $("#attStatus"), list = $("#attList"), view = $("#attView"); if (!st) return;
  if (!r || !r.ok || !files.length) { st.textContent = (r && r.error) || (it.n ? "The file(s) counted here were deleted in QuickBooks since the last sync - Resync to refresh the count" : "No scan on file"); view.innerHTML = `<div class="tr-note">${_ge(st.textContent)}</div>`; return; }
  st.textContent = `${files.length} file${files.length === 1 ? "" : "s"} · links expire in a few minutes`;
  const show = (f, btn) => { list.querySelectorAll(".att-file").forEach(x => x.classList.toggle("on", x === btn));
    const isImg = /\.(png|jpe?g|gif|webp|heic)(\?|$)/i.test(f.name || ""); const isPdf = /\.pdf(\?|$)/i.test(f.name || "");
    view.innerHTML = `<div class="att-tools"><b>${_ge(f.name || "attachment")}</b><a class="btn small" href="${_ge(f.url)}" target="_blank" rel="noopener">Open in a new tab ↗</a><a class="btn small" href="${_ge(f.url)}" download>Download</a></div>`
      + (isImg ? `<img class="att-img" src="${_ge(f.url)}" alt="">` : `<iframe class="att-frame" src="${_ge(f.url)}${isPdf ? "#toolbar=1" : ""}" title="attachment"></iframe>`); };
  files.forEach((f, k) => { const b = document.createElement("button"); b.type = "button"; b.className = "att-file"; b.textContent = f.name || ("file " + (k + 1)); b.onclick = () => show(f, b); list.appendChild(b); if (k === 0) show(f, b); });
  list.hidden = files.length < 2;
}
async function openAttachmentViewer(type, id, title, expected) {
  let ov = $("#attViewer");
  if (!ov) { ov = document.createElement("div"); ov.id = "attViewer"; ov.className = "xdlg-ov"; document.body.appendChild(ov); ov.onclick = (e) => { if (e.target === ov) ov.remove(); }; }
  ov.innerHTML = `<div class="xdlg att-dlg" role="dialog"><div class="att-head"><h3>${_ge(title || (type + " " + id))}</h3><span class="dim" id="attStatus">fetching fresh links from QuickBooks…</span><button class="btn small" id="attClose">Close</button></div>
    <div class="att-body"><div class="att-list" id="attList"></div><div class="att-view" id="attView"><div class="tr-note">Loading…</div></div></div></div>`;
  $("#attClose").onclick = () => ov.remove();
  let r;
  try { r = await (await fetch(`/api/attachment?id=${encodeURIComponent(id)}&type=${encodeURIComponent(type)}`)).json(); }
  catch (e) { r = { ok: false, error: String(e) }; }
  const files = (r && r.files) || [];
  const st = $("#attStatus"), list = $("#attList"), view = $("#attView"); if (!st) return;
  if (!r || !r.ok || !files.length) { st.textContent = (r && r.error) || (expected ? "The file(s) counted here were deleted in QuickBooks since the last sync - Resync to refresh the count" : "No attachment on file"); view.innerHTML = `<div class="tr-note">${_ge(st.textContent)}</div>`; return; }
  st.textContent = `${files.length} file${files.length === 1 ? "" : "s"} · links expire in a few minutes` + (expected && files.length < expected ? ` · ${expected - files.length} counted file${expected - files.length === 1 ? " was" : "s were"} deleted in QuickBooks since the last sync` : "");
  const show = (f, btn) => { list.querySelectorAll(".att-file").forEach(x => x.classList.toggle("on", x === btn));
    const isImg = /\.(png|jpe?g|gif|webp|heic)(\?|$)/i.test(f.name || ""); const isPdf = /\.pdf(\?|$)/i.test(f.name || "");
    view.innerHTML = `<div class="att-tools"><b>${_ge(f.name || "attachment")}</b><a class="btn small" href="${_ge(f.url)}" target="_blank" rel="noopener">Open in a new tab ↗</a><a class="btn small" href="${_ge(f.url)}" download>Download</a></div>`
      + (isImg ? `<img class="att-img" src="${_ge(f.url)}" alt="">` : `<iframe class="att-frame" src="${_ge(f.url)}${isPdf ? "#toolbar=1" : ""}" title="attachment"></iframe>`); };
  files.forEach((f, k) => { const b = document.createElement("button"); b.type = "button"; b.className = "att-file"; b.textContent = f.name || ("file " + (k + 1)); b.onclick = () => show(f, b); list.appendChild(b); if (k === 0) show(f, b); });
}
function qboLinkCell(text, url, title) {
  const td = document.createElement("td"); td.className = "left";
  const label = text || "-";
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
  "Fund in - pay vendors": "Pay vendors",
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
// A draw's headline: the MONTH the invoice bills ("September 2026" - MFD draws are monthly and carry
// no number), or "Draw #N" when the GC numbers them (CP), else "No draw yet" / "Draw". Prefers the
// backend parse (draw_month / draw_no), falls back to the memo (owner 2026-09-10: the tab was showing
// "Draw #2026", the year, and no period).
function drawTitle(dr) {
  if (!dr) return "Draw";
  if (dr.no_draw) return "No draw yet";
  if (dr.draw_month) return dr.draw_month;
  const p = drawPeriod(dr.matched_invoice); if (p && !/^Draw #/.test(p)) return p;   // month from the memo
  if (dr.draw_no) return "Draw #" + dr.draw_no;
  return "Draw";
}
// The span a draw covers, "08/02/2026 – 09/04/2026" (mm/dd/yyyy), from the backend period or the
// "(Period: … - …)" in the memo. "" when neither carries one.
function drawSpan(dr) {
  if (!dr) return "";
  if (dr.period_start && dr.period_end) return fmtDate(dr.period_start) + " – " + fmtDate(dr.period_end);
  const m = String(dr.matched_invoice || "").match(/Period:\s*([\d/]+)\s*-\s*([\d/]+)/i);
  return m ? m[1] + " – " + m[2] : "";
}
// " · September 2026" / " · Draw #4" suffix for a one-line label; "" when the draw has no specific name.
function drawTag(dr) { const t = drawTitle(dr); return (t && t !== "Draw" && t !== "No draw yet") ? " · " + t : ""; }
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
// Swap a tab's `.hint` between its generic blurb and a live "Showing: ..." filter summary.
function _setHintFilter(tab, summary) {
  const h = document.querySelector(`.tab-page[data-tab="${tab}"] .hint, [data-sec="${tab}"] .hint`); if (!h) return;
  if (!h.dataset.generic) h.dataset.generic = h.innerHTML;   // capture the generic blurb once
  if (summary) { h.innerHTML = `<b>Showing:</b> ${_ge(summary)} <span class="hint-clear-note">- clear the filters for the full list</span>`; h.classList.add("hint-filtered"); }
  else { h.innerHTML = h.dataset.generic; h.classList.remove("hint-filtered"); }
}

// ── Funding by project (owner 2026-09-02: fold Draws into the project page). One row per job:
// the next draw the GC owes, what blocks it (unpaid bills on EARLIER draws - the funding chain),
// and the latest draw's vendors paid. Click a row -> the project page, where the work happens.
const _isPaidBill = b => !!b.pay_date || (b.pay_status || "").toLowerCase().startsWith("bill paid") || (num(b.open) <= 0.005 && !!b.pay_status);
function _fundingRows() {
  const draws = DRAWS.draws || [];
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
    draw.waivers = draw.bills.filter(b => b.waiver).length;   // caption only - doesn't gate the stage
    toast(received ? "Waiver marked in hand" : "Waiver cleared");
    renderProjects();
  } catch (e) {
    cb.checked = !received;                        // revert on failure
    toast("Could not save: " + e.message);
  }
}


function renderCosts() {
  // Grouped: cost TYPE (parent) → job TYPE (sub) - the JobTread model. Material
  // rolls to ONE cost-type parent; the job-type sub shows the split for budget.
  const groups = COST.by_cost_type || [];
  const total = COST.loaded_total || groups.reduce((t, g) => t + (g.actual || 0), 0);
  $("#costCount").textContent = total
    ? `($${Math.round(total).toLocaleString()} from QuickBooks · where the money goes)`
    : "(no cost data - run load_costs.py)";
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
    ptr.appendChild(rightText(total ? ((g.actual || 0) / total * 100).toFixed(1) + "%" : "-"));
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
      else ct.appendChild(document.createTextNode("-"));
      str.appendChild(ct);
      const av = document.createElement("td"); av.appendChild(moneyCell(s.actual)); str.appendChild(av);
      str.appendChild(rightText(total ? ((s.actual || 0) / total * 100).toFixed(1) + "%" : "-"));
      str.appendChild(rightText(String(s.lines || 0)));
      tbody.appendChild(str);
    }
  }
}
function rightText(v) { const td = document.createElement("td"); const s = document.createElement("span"); s.textContent = v; td.appendChild(s); return td; }

// Cost mix - "how much each cost type takes, % wise" as one proportional bar + legend.
// cost-mix categories: none of these may be the accent, --pos or --neg (those colours MEAN something elsewhere)
const MIX_PALETTE = ["#4A6B8A", "#3f8f8a", "#B9541E", "#6b5b95", "#b8860b", "#5b7fbf",
                     "#b07a9a", "#4478a0", "#7a5c3e", "#5c8a6b", "#8a4a6b", "#997a3d"];
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
  $("#codeJobsTitle").textContent = `${label} (${code}) - ${money(tot)} across ${rows.length} job${rows.length === 1 ? "" : "s"}`;
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
    tr.appendChild(rightText(tot ? ((r.actual || 0) / tot * 100).toFixed(1) + "%" : "-"));
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
// "34449 - CP745 - Firestone Forever…" → { draw:"34449", name:"Firestone…" }.
// Used only as a fallback - invoice_no and the WIP name win when present.

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
    : "(no cost data - run load_costs.py)";
  const cols = [["Vendor", "left"], ["Type", "left"], ["Jobs", "right"], ["Total spend (QBO)", "right"], ["Open bills (QBO)", "right"], ["Open $ (QBO)", "right"]];   // labelled: QuickBooks open AP, subs included
  const thead = $("#vendorTable thead"), tbody = $("#vendorTable tbody");
  thead.innerHTML = ""; tbody.innerHTML = "";
  const htr = document.createElement("tr");
  for (const [c, al] of cols) { const th = document.createElement("th"); if (al === "left") th.className = "left"; th.textContent = c; htr.appendChild(th); }
  thead.appendChild(htr);
  const gType = v => (v.vtype || "-").split(":")[0].trim();   // Sub | Service | Supplier
  const vs = ($("#vendorSort") && $("#vendorSort").value) || "open";   // open $ (default) | name A-Z | total spend
  const cmp = vs === "name" ? ((a, b) => (a.vendor || "").localeCompare(b.vendor || ""))
    : vs === "spend" ? ((a, b) => (b.spend || 0) - (a.spend || 0) || (a.vendor || "").localeCompare(b.vendor || ""))
    : ((a, b) => (b.open_bal || 0) - (a.open_bal || 0) || (a.vendor || "").localeCompare(b.vendor || ""));   // default: most owed first
  const rows = [...vends].sort(grouped ? ((a, b) => gType(a).localeCompare(gType(b)) || cmp(a, b)) : cmp);
  const vendorRow = v => {
    const tr = document.createElement("tr");
    tr.classList.add("row-click"); tr.title = "Open this vendor's page";
    tr.onclick = (e) => { if (e.target.closest(".cell")) return; openVendorPage(v.vendor); };
    tr.appendChild(leftText(v.vendor));
    const ty = document.createElement("td"); ty.className = "left";
    const pill = document.createElement("span"); pill.className = "vtype" + (v.vtype === "Sub" ? " sub" : (v.vtype === "Service" ? " service" : ""));
    pill.textContent = v.vtype || "-"; ty.appendChild(pill); tr.appendChild(ty);
    tr.appendChild(rightText(String(v.jobs || 0)));
    { const sc = document.createElement("td"); sc.appendChild(moneyCell(v.spend || 0)); tr.appendChild(sc); }
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
  for (const k of cell.querySelectorAll(".bg-key, .sg-key, .g-cust")) if (!k.title) k.title = k.textContent || "";   // names truncate with an ellipsis in the grid - the full name on hover
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
// Compact numeric date for the dense grid: mm/dd/yyyy (owner 2026-08-21). Still month-first (never
// year-first). bill_date is ISO yyyy-mm-dd.
function fmtDateShort(v) {   // mm/dd/yyyy (owner 2026-08-21: 4-digit year everywhere)
  const m = String(v || "").trim().match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? `${m[2]}/${m[3]}/${m[1]}` : (v ? String(v) : "–");
}
// Per-column widths for the Bills grid - drag the divider between headers to resize;
// a squished column wraps its text instead of clipping. Widths persist per person.
const BILL_COL_DEFAULTS = { "Vendor": 210, "Project": 300, "Bill #": 100, "Memo": 260, "Date": 110, "Amount": 100,
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
// ── Excel-style column filters (owner 2026-09-22: "the bill tracker should just show the bills as a list just like the
// excel, and give me the ability to filter/search down just like i do with excel table filtering"). One mechanism for
// every bill table: a funnel in each header opens that column's distinct values (with counts, computed over the rows
// the OTHER columns leave - Excel's behaviour), a search box, Select all / None / Clear. State per table, per session.
const _hf = {};                                     // tableKey -> { colKey -> Set(values) }
function _ym(d) {   // "YYYY-MM" from an ISO date, or from a typed m/d/yyyy the workbook let through; "" when neither
  const t = String(d || ""); if (/^\d{4}-\d{2}/.test(t)) return t.slice(0, 7);
  const m = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/); return m ? `${m[3]}-${m[1].padStart(2, "0")}` : "";
}
const HF_BILL_COLS = {                              // colKey -> [getter, label of a value]
  vendor:  [b => b.vendor || "", v => v || "(no vendor)"],
  project: [b => b.project_no || "", v => v || "(no project #)"],
  client:  [b => b.client || "", v => v || "(no client)"],
  bill:    [b => b.bill_ref || "", v => v || "(no bill #)"],
  invoice: [b => b.invoice_no || "", v => v || "(no invoice)"],
  date:    [b => _ym(b.bill_date), v => v ? v.slice(5, 7) + "/" + v.slice(0, 4) : "(no date)"],
  open:    [b => (bOpen(b) > 0.005 ? "Open" : "Paid off"), v => v],
  pay:     [b => b.pay_status || "", v => v || "(none)"],
  inv:     [b => b.invoice_status || "", v => v || "(none)"],
  lien:    [b => b.lien_status || "", v => v ? (LIEN_SHORT[v] || v) : "(no lien clock)"],
  appr:    [b => b.approved || "", v => v === "approved" ? "Approved" : (v === "not approved" ? "Not approved" : (v || "(blank)"))],
};
function hfState(tableKey) { return _hf[tableKey] || (_hf[tableKey] = {}); }
function hfPasses(tableKey, b, except) {
  const st = hfState(tableKey);
  for (const k in st) { if (k === except || !st[k].size) continue; const get = HF_BILL_COLS[k][0]; if (!st[k].has(get(b))) return false; }
  return true;
}
function hfActive(tableKey) { const st = hfState(tableKey); return Object.keys(st).some(k => st[k].size); }
function hfClear(tableKey) { const st = hfState(tableKey); for (const k in st) st[k].clear(); }
let _hfOpen = null;   // the one open menu
function hfCloseMenu() { if (_hfOpen) { _hfOpen.remove(); _hfOpen = null; } }
document.addEventListener("click", (e) => { if (_hfOpen && !e.target.closest(".hf-menu, .hf-btn")) hfCloseMenu(); });
document.addEventListener("keydown", (e) => { if (e.key === "Escape") hfCloseMenu(); });
window.addEventListener("scroll", (e) => { if (_hfOpen && !(e.target && e.target.nodeType === 1 && _hfOpen.contains(e.target))) hfCloseMenu(); }, true);   // the page scrolled (not the list itself)
// Every multi-select menu (filter bar, the date months, the vendor page) closes on a click OUTSIDE it or on Esc - and
// never on a tick inside it (owner 2026-09-22: "i expect the filter to STAY open when i select"). One closer, app-wide.
document.addEventListener("click", (e) => {
  if (e.target.closest(".msel, .datef, .hf-menu, .hf-btn, .vp-projwrap")) return;
  document.querySelectorAll(".msel-menu:not(.hf-menu):not(.vp-projlist)").forEach(m => { if (!m.hidden) m.hidden = true; });
});
document.addEventListener("keydown", (e) => { if (e.key === "Escape") document.querySelectorAll(".msel-menu:not(.hf-menu):not(.vp-projlist)").forEach(m => { if (!m.hidden) m.hidden = true; }); });
function hfDecorate(th, tableKey, colKey, rowsFn, rerender) {   // rowsFn() = the rows before THIS column's filter (after every other filter)
  if (!HF_BILL_COLS[colKey]) return;
  const [get, lbl] = HF_BILL_COLS[colKey];
  const st = hfState(tableKey); const sel = st[colKey] || (st[colKey] = new Set());
  const btn = document.createElement("button"); btn.type = "button"; btn.className = "hf-btn" + (sel.size ? " on" : "");
  btn.innerHTML = '<svg viewBox="0 0 16 16" width="12" height="12" aria-hidden="true"><path d="M1.5 2.5h13l-5 6v4.5l-3 1.5V8.5z" fill="currentColor"/></svg>';   // the Excel funnel
  btn.title = sel.size ? `Filtered: ${sel.size} value${sel.size === 1 ? "" : "s"} - click to change` : "Filter this column";
  btn.onclick = (e) => {
    e.stopPropagation(); e.preventDefault();
    if (_hfOpen && _hfOpen.isConnected && !_hfOpen.hidden && _hfOpen._for === btn) { hfCloseMenu(); return; }
    hfCloseMenu();
    document.querySelectorAll(".msel-menu:not(.hf-menu)").forEach(m => { m.hidden = true; });   // one menu at a time
    const rows = rowsFn();
    const counts = new Map(); for (const b of rows) { const v = get(b); counts.set(v, (counts.get(v) || 0) + 1); }
    for (const v of [...sel]) if (!counts.has(v)) counts.set(v, 0);   // a picked value that no longer appears still shows, so it can be unpicked
    const vals = [...counts.keys()].sort((a, b) => colKey === "date" ? b.localeCompare(a) : lbl(a).localeCompare(lbl(b), undefined, { numeric: true }));
    const menu = document.createElement("div"); menu.className = "msel-menu hf-menu wide"; menu._for = btn;
    const q = document.createElement("input"); q.type = "search"; q.className = "msel-search"; q.placeholder = "Search values";
    q.oninput = () => { const t = q.value.toLowerCase(); for (const lab of menu.querySelectorAll(".msel-opt")) lab.hidden = t && !lab.textContent.toLowerCase().includes(t); }; menu.appendChild(q);
    const tools = document.createElement("div"); tools.className = "msel-tools";
    const visible = () => [...menu.querySelectorAll(".msel-opt")].filter(l => !l.hidden).map(l => l.dataset.val);
    const all = document.createElement("button"); all.type = "button"; all.className = "msel-tool"; all.textContent = "Select all"; all.onclick = () => { visible().forEach(v => sel.add(v)); rerender(); hfCloseMenu(); };
    const none = document.createElement("button"); none.type = "button"; none.className = "msel-tool"; none.textContent = "None"; none.onclick = () => { visible().forEach(v => sel.delete(v)); rerender(); hfCloseMenu(); };
    const clr = document.createElement("button"); clr.type = "button"; clr.className = "msel-tool"; clr.textContent = "Clear"; clr.onclick = () => { sel.clear(); rerender(); hfCloseMenu(); };
    const cnt = document.createElement("span"); cnt.className = "msel-count"; cnt.textContent = `${vals.length} value${vals.length === 1 ? "" : "s"}`;
    tools.appendChild(all); tools.appendChild(none); tools.appendChild(clr); tools.appendChild(cnt); menu.appendChild(tools);
    for (const v of vals) {
      const lab = document.createElement("label"); lab.className = "msel-opt"; lab.dataset.val = v;
      const cb = document.createElement("input"); cb.type = "checkbox"; cb.checked = sel.has(v);
      cb.onchange = () => { if (cb.checked) sel.add(v); else sel.delete(v); rerender(); };
      const t = document.createElement("span"); t.className = "hf-val"; t.textContent = lbl(v);
      const c = document.createElement("span"); c.className = "hf-cnt"; c.textContent = String(counts.get(v) || 0);
      lab.appendChild(cb); lab.appendChild(t); lab.appendChild(c); menu.appendChild(lab);
    }
    const r = btn.getBoundingClientRect(); menu.style.position = "fixed"; menu.style.top = (r.bottom + 4) + "px"; menu.style.left = Math.min(r.left, window.innerWidth - 300) + "px"; menu.hidden = false;
    document.body.appendChild(menu); _hfOpen = menu; q.focus();
  };
  th.appendChild(btn); th.classList.add("hf-th");
}
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
function dateFilter(id, getDates, onChange, prev) {   // prev = an earlier state to keep (the host was re-rendered)
  const st = { mode: prev ? prev.mode : "month", months: prev && prev.months ? new Set(prev.months) : null, from: prev ? prev.from : "", to: prev ? prev.to : "" };   // months: null = all; copied, never shared
  const host = $("#" + id); if (!host) return st;
  host.innerHTML = `<span class="seg tiny"><button type="button" class="seg-btn on" data-m="month">Month</button><button type="button" class="seg-btn" data-m="date">Date</button></span>
    <span class="datef-month msel" id="${id}Msel"><button type="button" class="msel-btn" id="${id}Btn">All months</button><div class="msel-menu" id="${id}Menu" hidden></div></span>
    <span class="datef-range" hidden><input type="date" id="${id}From" title="From (leave blank for no lower bound)"> <span class="dim">to</span> <input type="date" id="${id}To" title="To - a statement 'as of' a day is just this box"></span>
    <button type="button" class="btn small datef-reset" id="${id}Reset" title="Back to all dates" hidden>Reset</button>`;
  const btn = $("#" + id + "Btn"), menu = $("#" + id + "Menu"), from = $("#" + id + "From"), to = $("#" + id + "To");
  from.value = st.from; to.value = st.to;
  const paintMode = () => { host.querySelectorAll(".seg-btn").forEach(b => b.classList.toggle("on", b.dataset.m === st.mode));
    host.querySelector(".datef-month").hidden = st.mode !== "month"; host.querySelector(".datef-range").hidden = st.mode !== "date"; };
  host.querySelectorAll(".seg-btn").forEach(b => b.onclick = () => { st.mode = b.dataset.m; paintMode(); onChange(); });
  btn.onclick = (e) => { e.stopPropagation(); const open = menu.hidden; document.querySelectorAll(".msel-menu").forEach(m => m.hidden = true); menu.hidden = !open; if (open) _placeMenu(btn, menu); };
  from.onchange = () => { st.from = from.value; onChange(); }; to.onchange = () => { st.to = to.value; onChange(); };
  $("#" + id + "Reset").onclick = () => { st.clear(); onChange(); };
  st.build = () => {
    const asc = [...new Set(getDates().map(x => String(x || "").slice(0, 7)).filter(x => /^\d{4}-\d{2}$/.test(x)))].sort();
    // months picked on another view of the same page stay picked (the vendor page swaps bill dates for payment dates);
    // the list shows the ones this view has
    const sel = st.months === null ? new Set(asc) : new Set([...st.months].filter(m => asc.includes(m)));
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
        const here = [...s2].filter(m => asc.includes(m)).length;
        st.months = (here === asc.length || here === 0) ? null : s2;   // unticking the last one goes back to ALL (owner 2026-09-02)
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
let billDate = null;
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
  const closeAll = () => document.querySelectorAll(".msel-menu:not([hidden]):not(.hf-menu)").forEach(m => { m.hidden = true; });   // the funnel menu has its own closer (it may scroll itself)
  window.addEventListener("resize", closeAll);
  window.addEventListener("scroll", closeAll, true);   // any scrolling container - the pinned menu would drift otherwise
})();
const _BMONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
function billMonthLabel(ym) { const [y, m] = ym.split("-"); return `${_BMONTHS[+m - 1]} ${y}`; }
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
let _billQ = "";   // the Bill Tracker search box (owner 2026-09-22)
function billClearFilters() {
  for (const cfg of BILL_MSEL) (billMSel[cfg.id] || (billMSel[cfg.id] = new Set())).clear();
  hfClear("bills"); _billQ = ""; { const q = $("#bfQuick"); if (q) q.value = ""; }
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
  vendor: (a, b) => (a.vendor || "").localeCompare(b.vendor || "") || String(b.bill_date || "").localeCompare(String(a.bill_date || "")),   // A-Z, then newest first - the Bill Tracker's own order (owner 2026-09-23)
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

  // filter (view predicate AND every dropdown AND the search AND the Excel-style header filters), then sort
  const qOK = !_billQ.trim() ? () => true : b => _vq(_billQ, [b.vendor, b.project_no, nameOf(b.project_no), b.client, b.bill_ref, b.memo, b.invoice_no, fmtDateShort(b.bill_date), b.bill_date,
    Math.round(num(b.line_amount)), money(b.line_amount), Math.round(bOpen(b)), b.pay_status, b.invoice_status, b.lien_status, b.approved, b.division]);
  const baseRows = bills.filter(b => view.pred(b) && billPassesFilters(b, f) && qOK(b));
  let rows = baseRows.filter(b => hfPasses("bills", b));
  { const hc = $("#bfHfClear"); if (hc) { hc.hidden = !hfActive("bills"); } }
  const sortKey = $("#billSort") ? $("#billSort").value : "vendor";
  rows = [...rows].sort(BILL_SORTS[sortKey] || BILL_SORTS.vendor);

  const openSum = rows.reduce((t, b) => t + bOpen(b), 0);
  const lienN = rows.filter(b => BILL_LIEN_RISK.has(b.lien_status)).length;
  $("#billsNote").textContent = bills.length ? `(${rows.length.toLocaleString()} of ${bills.length.toLocaleString()})` : "(no AP data - run load_bill_tracker.py)";
  { const qs = $("#billsQuickStat"); if (qs) qs.textContent = bills.length ? `${money(openSum)} open · ${lienN} lien risk` : ""; }
  { const pr = $("#btnPayRunGo"); if (pr) { const n = (BILLS || []).filter(b => b.pay_selected).length; pr.textContent = n ? `Pay run (${n}) →` : "Pay run →"; pr.classList.toggle("on", n > 0); } }

  // table. Each status is its OWN column (Paid / Invoice / Lien / Appr) so a blank in
  // one never hides a missing value by being merged with the others.
  const group = $("#billGroup") ? $("#billGroup").value : "none";   // a flat list like the Excel, unless the owner picks a grouping
  const thead = $("#billTable thead"), tbody = $("#billTable tbody");
  thead.innerHTML = "";
  thead.hidden = false; tbody.innerHTML = "";
  const cols = [["Vendor", "left", ""], ["Project", "left", ""], ["Bill #", "left", "Bill number - opens the bill in QuickBooks"],
                ["Memo", "left", "The bill's memo in QuickBooks (project, address, client - what AP typed)"],
                ["Invoice #", "left", "The draw (AR invoice) this bill is matched to - opens the invoice page; GC paid / GC owes is the live QuickBooks state of that invoice"],
                ["Date", "left", "Bill date - the funnel filters by month"], ["Amount", "right", "Bill amount"], ["Open", "right", "Open balance we still owe"],
                ["Paid", "left", "Did we pay the vendor?"], ["Invoice", "left", "Was the AR invoice (draw) paid by the GC?"],
                ["Lien", "left", "Texas lien-notice clock"], ["Appr", "left", "Approved for payment?"]];
  // Fixed layout + a <colgroup> so column widths are exact and draggable; each header
  // carries a resize grip on its right divider, and the table width tracks the sum.
  const table = $("#billTable");
  { const oldCg = table.querySelector("colgroup"); if (oldCg) oldCg.remove(); }
  const colgroup = document.createElement("colgroup");
  const htr = document.createElement("tr");
  let wsum = 0;
  const HF_KEYS = { "Vendor": "vendor", "Project": "project", "Bill #": "bill", "Invoice #": "invoice", "Date": "date", "Open": "open", "Paid": "pay", "Invoice": "inv", "Lien": "lien", "Appr": "appr" };
  cols.forEach(([c, al, tip], i) => {
    const th = document.createElement("th"); if (al === "left") th.className = "left"; th.textContent = c; if (tip) th.title = tip;
    if (HF_KEYS[c]) hfDecorate(th, "bills", HF_KEYS[c], () => baseRows.filter(b => hfPasses("bills", b, HF_KEYS[c])), renderBills);
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
    else if (_billQ.trim() || hfActive("bills")) td.textContent = _billQ.trim() ? `No bills match "${_billQ.trim()}"${hfActive("bills") ? " with the column filters" : ""}.` : "No bills match the column filters.";
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
  const vs = document.createElement("span"); vs.className = "bill-vendor lnk"; vs.textContent = b.vendor || "–"; vs.title = (b.vendor || "") + " - open this vendor's page";
  vs.onclick = (e) => { e.stopPropagation(); openVendorPage(b.vendor); };   // dial in by vendor (owner 2026-09-16)
  vtd.appendChild(vs); tr.appendChild(vtd);
  // Project (division chip + CLIENT - easier to scan than the job name; job name is in the tooltip)
  const ptd = document.createElement("td"); ptd.className = "left";
  if (b.project_no) {
    const chip = document.createElement("span"); const dc = divClass(b.division);
    chip.className = "divchip" + (dc ? " " + dc : ""); chip.textContent = b.project_no; ptd.appendChild(chip);
    const nm = nameOf(b.project_no); const disp = b.client || nm;
    if (disp) { const s = document.createElement("span"); s.className = "bill-name" + (b.client ? " lnk" : ""); s.textContent = disp;
      s.title = (b.client ? (nm ? `${b.client} · ${nm}` : b.client) : nm) + (b.client ? " - open this client's page" : "");
      if (b.client) s.onclick = (e) => { e.stopPropagation(); openClientPage(b.client); }; ptd.appendChild(s); }
  } else { ptd.appendChild(document.createTextNode("–")); }
  tr.appendChild(ptd);
  // Bill # (QBO deep link)
  tr.appendChild(qboLinkCell(b.bill_ref, qboBillHref(b.qbo_link), "Open this bill in QuickBooks"));
  { const bid = b.bill_id || (qboBillHref(b.qbo_link) || "").replace(/.*txnId=(\d+).*/, "$1");
    if (bid) { const same = (BILLS || []).filter(y => y.vendor === b.vendor && y.bill_id).map(y => ({ type: "Bill", id: String(y.bill_id), n: y.att || 0, title: `${y.vendor || ""} · bill ${y.bill_ref || ""}` }));
      const _ab = attBtn("Bill", bid, b.att, `${b.vendor || ""} · bill ${b.bill_ref || ""}`, { items: same, index: Math.max(0, same.findIndex(y => y.id === String(bid))) }); _ab.style.marginLeft = "6px"; tr.lastElementChild.appendChild(_ab); } }
  { const mt = document.createElement("td"); mt.className = "left bill-memo"; const ms = document.createElement("span"); ms.textContent = b.memo || ""; ms.title = b.memo || ""; mt.appendChild(ms); if (!b.memo) mt.classList.add("dim"); tr.appendChild(mt); }   // the QBO memo (owner 2026-09-22)
  tr.appendChild(_billInvCell(b));   // the draw / AR invoice this bill is matched to, and whether the GC paid it (owner 2026-09-16)
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
// The invoice a bill is matched to (the draw) + the live AR state of that invoice: "GC paid" or "GC owes $X" (owner
// 2026-09-16: "i need to see the invoice it's associated to and if it's been paid"). The number opens the invoice page.
function _billInvCell(b) {
  const td = document.createElement("td"); td.className = "left bill-inv";
  if (!b.invoice_no) { td.appendChild(dimDash()); td.title = "Not matched to a draw / invoice in the Bill Tracker yet"; return td; }
  const link = document.createElement("span"); link.className = "inv-detail-link"; link.textContent = b.invoice_no; link.title = "Open this invoice's page (the draw, its bills, the collections log)";
  link.onclick = (e) => { e.stopPropagation(); openInvoicePage({ doc_number: b.invoice_no, customer: b.inv_customer || b.client, project_no: b.project_no, division: b.division, qbo_txn_id: b.inv_qbo_id }); };
  td.appendChild(link);
  if (b.inv_ar_status != null || b.inv_balance != null) {
    const paid = b.inv_ar_status === "Paid" || num(b.inv_balance) <= 0.005;
    td.appendChild(stText(paid ? "GC paid" : "GC owes " + money(b.inv_balance), paid ? "st-ok" : "st-warn",
      paid ? "The GC has paid this invoice" + (b.inv_date ? " (invoiced " + fmtDateShort(b.inv_date) + ")" : "") : `Invoice ${b.invoice_no} is still open in QuickBooks - ${money(b.inv_balance)} of ${money(b.inv_amount)}`));
  }
  return td;
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

// Vendor page (QBO-style, ON DEMAND) - one vendor's bills, fetched per vendor via /api/vendor (never
// in the bulk load). Each bill shows its project, or "multiple" -> click the bill to see every line
// item + project #. Filter by pay status. Owner 2026-08-28: "vendor center open into its own vendor
// page like qbo ... see the bill its paying and the project ... if multiple say multiple, click for lines".
let _vendorData = null, _vendorType = "all", _vendorView = "bills";   // bills | payments
let _vendorInv = "any";   // the vendor page's invoice filter (any | gcpaid | gcowes | none)
let _vendorQ = "";   // the vendor page search - one box, filters whichever view is up (owner 2026-09-22: "need ability to search on both pages")
let _vendorDate = null;   // the vendor page Date filter (Month | Date from-to), same component as the Bill Tracker's; state survives re-renders
let _vendorProjOpen = false, _vendorProjIdx = -1;   // the Project box's suggestion list: open? which row is highlighted (ArrowDown / ArrowUp, Enter picks)
let _vendorProj = "";   // the standalone Project box beside it (owner 2026-09-22: "make project a standalone box") - project # or job name, both views
const _vq = (q, parts) => { q = (q || "").trim().toLowerCase(); if (!q) return true; const hay = parts.filter(x => x != null && x !== "").map(x => String(x).toLowerCase()).join(" \u0001 "); return q.split(/\s+/).every(w => hay.includes(w)); };
const _vendorBillOpen = new Set();
async function openVendorPage(vendor) {
  if (_ppLeaveBlocked()) return;   // unsaved pay ticks on the project page: Save or Discard first
  openRecord(vendor, "loading…"); skeletonInto($("#recordBody"), 6);
  _recSave({ k: "vendor", id: vendor, view: "bills" });
  const body = $("#recordBody");
  _vendorData = null; _vendorType = "all"; _vendorView = "bills"; _vendorInv = "any"; _vendorQ = ""; _vendorProj = ""; _vendorDate = null; _vendorBillOpen.clear();
  hfClear("vendorBills"); _stubSel.clear(); _vendorProjOpen = false; _vendorProjIdx = -1;   // nothing carries over from the last vendor
  let data;
  try { data = await (await fetch("/api/vendor?v=" + encodeURIComponent(vendor))).json(); }
  catch (e) { body.innerHTML = ""; body.textContent = "could not load this vendor"; return; }
  body.innerHTML = "";
  if (!data || !data.ok) { body.textContent = (data && data.error) || "no data for this vendor"; return; }
  _vendorData = data;
  renderVendorPage();
}
let _vpSeq = 0;   // bumps on every vendor render; an async step that finds it moved on stops (no double table, no painting over another page)
function _vendorPageShowing(d) { const rv = $("#recordView"); return !!(d && _vendorData === d && rv && !rv.hidden && ($("#recordTitle") || {}).textContent === d.vendor); }
function renderVendorPage() {
  const d = _vendorData; if (!d) return;
  const seq = ++_vpSeq;
  // a filter STAYS OPEN while the owner ticks values (owner 2026-09-22): the page re-renders on every change, so
  // remember which menus were open before the body is rebuilt and reopen them after
  const _dateMenuWasOpen = !!($("#vpDateMenu") && !$("#vpDateMenu").hidden);
  const body = $("#recordBody"); body.innerHTML = "";
  // Bills | Payments view toggle
  const vseg = document.createElement("div"); vseg.className = "seg vendor-seg";
  for (const [k, lbl] of [["bills", `Bills (${d.count})`], ["payments", `Payments (${d.pay_count || 0})`]]) {
    const b = document.createElement("button"); b.type = "button"; b.className = "seg-btn" + (_vendorView === k ? " on" : ""); b.textContent = lbl;
    b.onclick = () => { _vendorView = k; _recSave({ k: "vendor", id: d.vendor, view: k }); renderVendorPage(); }; vseg.appendChild(b);
  }
  { const wrap = document.createElement("span"); wrap.className = "vp-projwrap";
    const pj = document.createElement("input"); pj.type = "search"; pj.id = "vpProj"; pj.className = "msel-search vendor-proj"; pj.value = _vendorProj; pj.placeholder = "Project #"; pj.autocomplete = "off";
    pj.title = "Narrow to a project # - type, then ArrowDown and Enter to pick one of this vendor's projects; both views";
    // this vendor's projects (bills + payments), the list the box suggests from
    const projs = new Map();
    for (const b of (BILLS || []).filter(b => (b.vendor || "") === d.vendor)) if (b.project_no) projs.set(b.project_no, b.client || nameOf(b.project_no) || "");
    for (const p of (d.payments || [])) for (const no of (p.projects || [])) if (!projs.has(no)) projs.set(no, nameOf(no) || "");
    const matches = () => { const t = _vendorProj.trim(); return [...projs].filter(([no, nm]) => !t || _vq(t, [no, nm, nameOf(no)])).sort((a, b) => a[0].localeCompare(b[0], undefined, { numeric: true })).slice(0, 12); };
    const refocus = () => { const n = $("#vpProj"); if (n) { n.focus(); const pos = n.value.length; try { n.setSelectionRange(pos, pos); } catch {} } };
    const pick = (no) => { _vendorProj = no; _vendorProjOpen = false; _vendorProjIdx = -1; renderVendorPage(); refocus(); };
    pj.oninput = () => { _vendorProj = pj.value; _vendorProjOpen = true; _vendorProjIdx = -1; clearTimeout(pj._t); pj._t = setTimeout(() => { const pos = pj.selectionStart; renderVendorPage(); const n = $("#vpProj"); if (n) { n.focus(); try { n.setSelectionRange(pos, pos); } catch {} } }, 160); };
    pj.onfocus = () => { if (!_vendorProjOpen) { _vendorProjOpen = true; renderVendorPage(); refocus(); } };
    pj.onkeydown = (e) => {
      const m = matches();
      if (e.key === "ArrowDown") { e.preventDefault(); _vendorProjOpen = true; _vendorProjIdx = Math.min(m.length - 1, _vendorProjIdx + 1); renderVendorPage(); refocus(); }
      else if (e.key === "ArrowUp") { e.preventDefault(); _vendorProjIdx = Math.max(-1, _vendorProjIdx - 1); renderVendorPage(); refocus(); }
      else if (e.key === "Enter") { e.preventDefault(); if (_vendorProjIdx >= 0 && m[_vendorProjIdx]) pick(m[_vendorProjIdx][0]); else if (m.length === 1) pick(m[0][0]); else { _vendorProjOpen = false; renderVendorPage(); refocus(); } }
      else if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); if (_vendorProjOpen) { _vendorProjOpen = false; renderVendorPage(); refocus(); } else { pj.value = ""; _vendorProj = ""; renderVendorPage(); refocus(); } }
    };
    pj.onblur = () => { setTimeout(() => { if (_vendorProjOpen && document.activeElement !== $("#vpProj") && !document.activeElement.closest(".vp-projlist")) { _vendorProjOpen = false; const l = document.querySelector(".vp-projlist"); if (l) l.remove(); } }, 150); };
    wrap.appendChild(pj);
    if (_vendorProjOpen) { const m = matches(); if (m.length) {
      const list = document.createElement("div"); list.className = "msel-menu vp-projlist"; list.hidden = false;
      m.forEach(([no, nm], i) => { const row = document.createElement("div"); row.className = "vp-projopt" + (i === _vendorProjIdx ? " on" : ""); row.tabIndex = -1;
        const a = document.createElement("b"); a.textContent = no; const b = document.createElement("span"); b.className = "dim"; b.textContent = nm ? " · " + nm : "";
        row.appendChild(a); row.appendChild(b); row.onmousedown = (e) => { e.preventDefault(); pick(no); }; list.appendChild(row); });
      wrap.appendChild(list); } }
    vseg.appendChild(wrap); }
  { const fld = document.createElement("span"); fld.className = "fld vp-datefld"; const lb = document.createElement("span"); lb.textContent = "Date"; fld.appendChild(lb);
    const host = document.createElement("span"); host.className = "datef"; host.id = "vpDate"; fld.appendChild(host); vseg.appendChild(fld); }
  { const q = document.createElement("input"); q.type = "search"; q.id = "vpSearch"; q.className = "msel-search vendor-search"; q.value = _vendorQ;
    q.placeholder = _vendorView === "payments" ? "Search payments (⌘F) - ref, date, client, project, bill #, amount" : "Search bills (⌘F) - anything on the row: client, bill #, memo, invoice #, date, amount";
    q.oninput = () => { _vendorQ = q.value; clearTimeout(q._t); q._t = setTimeout(() => { const pos = q.selectionStart; renderVendorPage(); const n = $("#vpSearch"); if (n) { n.focus(); try { n.setSelectionRange(pos, pos); } catch {} } }, 160); };
    q.onkeydown = (e) => { if (e.key === "Escape") { q.value = ""; _vendorQ = ""; renderVendorPage(); } };
    vseg.appendChild(q); }
  body.appendChild(vseg);
  { const dates = () => _vendorView === "payments" ? (d.payments || []).map(p => p.txn_date) : (BILLS || []).filter(b => (b.vendor || "") === d.vendor).map(b => b.bill_date);
    _vendorDate = dateFilter("vpDate", dates, renderVendorPage, _vendorDate); _vendorDate.build();   // after the bar is in the document - the control binds to #vpDate
    if (_dateMenuWasOpen) { const m = $("#vpDateMenu"), b = $("#vpDateBtn"); if (m && b) { m.hidden = false; _placeMenu(b, m); } } }
  if (_vendorView === "payments") {
    $("#recordSub").textContent = `${d.pay_count || 0} payments · ${money(d.pay_total || 0)} paid out this year`;
    return _renderVendorPayments(d, body, seq);
  }
  // Two systems, two labels (owner 2026-09-02: the list said one "open", the page another): QuickBooks
  // open AP covers every vendor incl. subs; the Bill Tracker excludes subs.
  $("#recordSub").textContent = (d.qbo_open != null ? `open ${money(d.qbo_open)} (QuickBooks, ${d.qbo_open_bills || 0} bills) · ` : "")
    + `Bill Tracker: ${d.count} bills · ${money(d.total)} billed · ${money(d.open)} open · ${d.paid_ct} paid` + (d.count ? "" : " (subs are not in the Bill Tracker)");
  // The Bill Tracker WITHIN the vendor page (owner 2026-09-16: "basically the bill tracker within here ... vendors where i can
  // dial in by vendor"): the same rows and columns as the tracker, filtered to this vendor - every bill with its project,
  // client, the invoice it sits on and whether the GC paid it, our pay status, lien and approval.
  const rowsAll = (BILLS || []).filter(b => (b.vendor || "") === d.vendor);
  if (!rowsAll.length && (d.qbo_bills || []).length) return _renderVendorQboBills(d, body);   // a sub: show its QBO bills instead
  const isPaid = b => bOpen(b) <= 0.005;
  const invState = b => !b.invoice_no ? "none" : (b.inv_ar_status === "Paid" || (b.inv_balance != null && num(b.inv_balance) <= 0.005)) ? "gcpaid" : "gcowes";
  const tools = document.createElement("div"); tools.className = "cp-tools";
  const seg = document.createElement("div"); seg.className = "seg big";
  for (const [k, lbl] of [["all", `All bills · ${rowsAll.length}`], ["open", `Unpaid · ${rowsAll.filter(b => !isPaid(b)).length}`], ["paid", `Paid · ${rowsAll.filter(isPaid).length}`]]) {
    const b = document.createElement("button"); b.type = "button"; b.className = "seg-btn" + (_vendorType === k ? " on" : ""); b.textContent = lbl;
    b.onclick = () => { _vendorType = k; renderVendorPage(); }; seg.appendChild(b);
  }
  tools.appendChild(seg);
  const iseg = document.createElement("div"); iseg.className = "seg"; iseg.title = "The invoice (draw) each bill is matched to, and whether the GC has paid it";
  for (const [k, lbl] of [["any", "Any invoice"], ["gcpaid", "GC paid"], ["gcowes", "GC owes"], ["none", "No invoice yet"]]) {
    const b = document.createElement("button"); b.type = "button"; b.className = "seg-btn" + (_vendorInv === k ? " on" : ""); b.textContent = lbl;
    b.onclick = () => { _vendorInv = k; renderVendorPage(); }; iseg.appendChild(b);
  }
  tools.appendChild(iseg);
  { const hc = document.createElement("button"); hc.type = "button"; hc.className = "btn small"; hc.textContent = "Clear column filters"; hc.hidden = !hfActive("vendorBills"); hc.onclick = () => { hfClear("vendorBills"); renderVendorPage(); }; tools.appendChild(hc); }
  const bt = document.createElement("button"); bt.type = "button"; bt.className = "btn small"; bt.textContent = "Open in Bill Tracker"; bt.title = "The Bill Tracker with its vendor filter set to this vendor - every other filter is there";
  bt.onclick = () => { if (_ppLeaveBlocked()) return; billVendorHidden = new Set(_billVendors().filter(v => v !== d.vendor)); activeBillView = "all"; closeRecord(); setTab("bills"); buildBillVendorFilter(); renderBills(); };
  tools.appendChild(bt);
  body.appendChild(tools);
  let rows = rowsAll.filter(b => _vendorType === "all" || (_vendorType === "paid" ? isPaid(b) : !isPaid(b)));
  if (_vendorInv !== "any") rows = rows.filter(b => invState(b) === _vendorInv);
  if (_vendorDate && _vendorDate.active()) rows = rows.filter(b => _vendorDate.passes(b.bill_date));   // Month | Date (from / to)
  if (_vendorProj.trim()) rows = rows.filter(b => _vq(_vendorProj, [b.project_no, nameOf(b.project_no)]));   // the standalone Project box
  if (_vendorQ.trim()) rows = rows.filter(b => _vq(_vendorQ, [b.project_no, nameOf(b.project_no), b.client, b.bill_ref, b.memo, b.invoice_no, fmtDateShort(b.bill_date), b.bill_date,
    Math.round(num(b.line_amount)), money(b.line_amount), Math.round(bOpen(b)), b.lien_status, b.division, isPaid(b) ? "paid" : "unpaid", invState(b) === "gcpaid" ? "gc paid" : invState(b) === "gcowes" ? "gc owes" : "no invoice"]));
  rows.sort((a, b) => String(b.bill_date || "").localeCompare(String(a.bill_date || "")) || String(a.bill_ref || "").localeCompare(String(b.bill_ref || "")));
  const baseRows = rows;                                   // before the header filters - what each funnel lists
  rows = rows.filter(b => hfPasses("vendorBills", b));
  if (!rows.length) { const p = document.createElement("div"); p.className = "bills-cap"; p.textContent = rowsAll.length ? (_vendorQ.trim() ? `No bills match "${_vendorQ.trim()}".` : "No bills match this filter.") : "No Bill Tracker rows for this vendor."; body.appendChild(p); return; }
  const scroll = document.createElement("div"); scroll.className = "table-scroll";
  const table = document.createElement("table"); table.className = "grid vp-grid"; const thead = document.createElement("thead"), tbody = document.createElement("tbody");
  const cols = [["Project", "left", "Project + client - the client opens the client's page"], ["Bill #", "left", "Bill number - opens the bill in QuickBooks"],
                ["Memo", "left", "The bill's memo in QuickBooks"],
                ["Invoice #", "left", "The draw (AR invoice) this bill is matched to - opens the invoice page; GC paid / GC owes is the live QuickBooks state of that invoice"],
                ["Date", "left", "Bill date - the funnel filters by month"], ["Amount", "right", "Bill amount"], ["Open", "right", "Open balance we still owe"],
                ["Paid", "left", "Did we pay the vendor?"], ["Invoice", "left", "The Bill Tracker's invoice pipeline status"], ["Lien", "left", "Texas lien-notice clock"], ["Appr", "left", "Approved for payment?"]];
  { const HF_KEYS = { "Project": "project", "Bill #": "bill", "Invoice #": "invoice", "Date": "date", "Open": "open", "Paid": "pay", "Invoice": "inv", "Lien": "lien", "Appr": "appr" };
    const htr = document.createElement("tr");
    for (const [c, al, tip] of cols) { const th = document.createElement("th"); th.className = al; th.title = tip; th.textContent = c;
      if (HF_KEYS[c]) hfDecorate(th, "vendorBills", HF_KEYS[c], () => baseRows.filter(b => hfPasses("vendorBills", b, HF_KEYS[c])), renderVendorPage);
      htr.appendChild(th); }
    thead.appendChild(htr); }
  const row = b => { const tr = billRow(b); tr.removeChild(tr.firstElementChild); return tr; };   // the tracker's row without the vendor column - it is the vendor's page
  // a flat list like the Excel (owner 2026-09-22: "i don't want to see it grouped by anything") - the funnels do the narrowing
  for (const b of rows) tbody.appendChild(row(b));
  table.appendChild(thead); table.appendChild(tbody); scroll.appendChild(table); body.appendChild(scroll);
  const cap = document.createElement("div"); cap.className = "bills-cap";
  cap.textContent = `${rows.length} bill${rows.length === 1 ? "" : "s"} · ${money(rows.reduce((t, b) => t + num(b.line_amount), 0))} billed · ${money(rows.reduce((t, b) => t + bOpen(b), 0))} open · Invoice # = the draw the Bill Tracker matched the bill to; GC paid / GC owes = that invoice's live QuickBooks balance · click a row for the bill's detail, the invoice # for the invoice page, the client for the client's page.`;
  body.appendChild(cap);
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
    { const lc = qboLinkCell(b.doc_number || "–", qboUrl(b.txn_type === "Expense" ? "expense" : "bill", b.txn_id), "Open this bill in QuickBooks"); if (b.att) { const ab = attBtn(b.txn_type === "Expense" ? "Purchase" : "Bill", b.txn_id, b.att, `bill ${b.doc_number || ""}`); ab.style.marginLeft = "6px"; lc.appendChild(ab); } tr.appendChild(lc); }
    tr.appendChild(leftText(b.projects.length > 1 ? b.projects.join(", ") : (b.projects[0] || "–")));
    const m = leftText(b.memo || "–"); m.title = b.memo || ""; m.className += " inv-memo"; tr.appendChild(m);
    const amt = document.createElement("td"); amt.appendChild(moneyCell(b.amount)); tr.appendChild(amt);
    tbody.appendChild(tr);
  }
  table.appendChild(thead); table.appendChild(tbody); scroll.appendChild(table); body.appendChild(scroll);
}
function _vpCapped(host, nodes, n, what) {   // the first n nodes, then a "+N more" that reveals the rest in place
  nodes.slice(0, n).forEach(x => host.appendChild(x));
  if (nodes.length <= n) return;
  const rest = document.createElement("div"); rest.hidden = true; nodes.slice(n).forEach(x => rest.appendChild(x));
  const more = document.createElement("button"); more.type = "button"; more.className = "btn tiny subtle vp-more"; more.textContent = `+${nodes.length - n} more ${what}`;
  more.onclick = (e) => { e.stopPropagation(); rest.hidden = !rest.hidden; more.textContent = rest.hidden ? `+${nodes.length - n} more ${what}` : "show fewer"; };
  host.appendChild(more); host.appendChild(rest);
}
function _vpProjCell(projs, title) {   // a cell of project #s, each opening its project page
  const td = document.createElement("td"); td.className = "left vp-projs";
  if (!projs.length) { td.textContent = "–"; td.classList.add("dim"); return td; }
  projs.forEach((p, i) => { if (i) td.appendChild(document.createTextNode(", ")); const a = document.createElement("a"); a.href = "#"; a.className = "vp-proj"; a.textContent = p; a.title = title || "Open this project's page";
    a.onclick = (e) => { e.preventDefault(); e.stopPropagation(); openProjectPage(p); }; td.appendChild(a); });
  return td;
}
function _vendorLines(b) {
  const wrap = document.createElement("div"); wrap.className = "bills-sub";
  const cap = document.createElement("div"); cap.className = "bills-cap"; cap.textContent = `${b.lines.length} line items on bill ${b.bill_ref || ""}${b.projects.length > 1 ? " · " + b.projects.filter(p => p !== "(multiple)").join(", ") : ""}`; wrap.appendChild(cap);
  const scroll = document.createElement("div"); scroll.className = "table-scroll";
  const table = document.createElement("table"); table.className = "grid"; const thead = document.createElement("thead"), tbody = document.createElement("tbody");
  const htr = document.createElement("tr");
  for (const [c, al] of [["Project #", "left"], ["Client", "left"], ["Description", "left"], ["Amount", "right"]]) { const th = document.createElement("th"); if (al === "left") th.className = "left"; th.textContent = c; htr.appendChild(th); }
  thead.appendChild(htr);
  for (const ln of b.lines) { const tr = document.createElement("tr");
    tr.appendChild(_vpProjCell(ln.project_no ? [ln.project_no] : []));
    tr.appendChild(leftText(ln.client || "–"));
    tr.appendChild(leftText(ln.description || "–"));
    tr.appendChild(rightText(money(ln.amount)));
    tbody.appendChild(tr);
  }
  table.appendChild(thead); table.appendChild(tbody); scroll.appendChild(table); wrap.appendChild(scroll);
  return wrap;
}
// Vendor payments view: the QBO BillPayments (money out) this year, from the local bill_payment table.
// + the bill payment STUB (owner 2026-09-22): "Print stub" per payment -> the PDF in the vendor's folder on the
// Accounting share, one file per payment; a column picker (QBO's six on by default); and the print HISTORY -
// every stub as printed, judged against QBO now (current / changed / voided / deleted), so a payment
// QuickBooks no longer shows keeps its stubs. History + registry come from /api/bill-payment/stubs.
const STUB_COLS_LS = "ledger.stubColumns";
let _stubCols = (() => { try { return JSON.parse(localStorage.getItem(STUB_COLS_LS)) || null; } catch { return null; } })();
let _stubHist = { vendor: null, prints: [], columns: [] };
async function _loadStubHistory(vendor) {
  try { const j = await (await fetch("/api/bill-payment/stubs?vendor=" + encodeURIComponent(vendor))).json();
    _stubHist = { vendor, prints: (j && j.prints) || [], columns: (j && j.columns) || [] }; }
  catch { _stubHist = { vendor, prints: [], columns: [] }; }
  if (!_stubCols && _stubHist.columns.length) _stubCols = _stubHist.columns.filter(c => c.default).map(c => c.key);
}
function _stubPill(status) {
  const s = document.createElement("span"); s.className = "stub-pill " + (status || "current");
  s.textContent = { current: "current", changed: "changed in QBO", voided: "voided in QBO", deleted: "deleted in QBO", unchecked: "not checked" }[status] || status;
  s.title = { current: "QuickBooks still shows this payment exactly as printed", changed: "The payment was edited in QuickBooks after this stub was printed - the stub is what went out",
    voided: "The payment was voided in QuickBooks after this stub was printed", deleted: "The payment no longer exists in QuickBooks - this stub is the record of it",
    unchecked: "No QuickBooks mirror on this machine, so this print could not be compared to QuickBooks" }[status] || "";
  return s;
}
function _stubLink(h) {   // one printed stub: opens the PDF as it went out
  const a = document.createElement("a"); a.className = "stub-link"; a.target = "_blank"; a.rel = "noopener";
  a.href = "/api/bill-payment/stub/file?id=" + encodeURIComponent(h.id);
  a.textContent = fmtDate(h.printed_at, true); a.title = h.file_path || "";
  if (h.file_exists === false) { a.classList.add("dim"); a.title = "PDF not on disk (Accounting share unmounted?) - " + (h.file_path || ""); }
  return a;
}
async function _printStub(paymentId, btn) {
  const was = btn.textContent; btn.disabled = true; btn.textContent = "Printing…";
  try {
    const res = await fetch("/api/bill-payment/stub", { method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ payment_id: paymentId, columns: _stubCols || undefined }) });
    const j = await res.json();
    if (!j.ok) throw new Error(j.error || "print failed");
    toast("Stub printed - " + (j.file || "").split("/").pop());
    window.open("/api/bill-payment/stub/file?id=" + encodeURIComponent(j.id), "_blank", "noopener");
    const d = _vendorData; await _loadStubHistory(_stubHist.vendor); if (_vendorPageShowing(d)) renderVendorPage();   // never paint over a page the owner opened meanwhile
  } catch (e) { toast("Could not print the stub: " + e.message); btn.disabled = false; btn.textContent = was; }
}
function _stubColumnPicker() {   // the registry as checkboxes; the choice sticks (localStorage) and applies to the next print
  const wrap = document.createElement("details"); wrap.className = "stub-cols";
  const sum = document.createElement("summary"); sum.textContent = "Stub columns"; wrap.appendChild(sum);
  const box = document.createElement("div"); box.className = "stub-cols-box";
  for (const c of _stubHist.columns) {
    const lab = document.createElement("label"); const cb = document.createElement("input"); cb.type = "checkbox";
    cb.checked = (_stubCols || []).includes(c.key);
    cb.onchange = () => { const order = _stubHist.columns.map(x => x.key); const set = new Set(_stubCols || []);
      if (cb.checked) set.add(c.key); else set.delete(c.key);
      _stubCols = order.filter(k => set.has(k)); if (!_stubCols.length) { _stubCols = [c.key]; cb.checked = true; }
      try { localStorage.setItem(STUB_COLS_LS, JSON.stringify(_stubCols)); } catch {} };
    lab.appendChild(cb); lab.appendChild(document.createTextNode(" " + c.label + (c.default ? "" : " (optional)"))); box.appendChild(lab);
  }
  const reset = document.createElement("button"); reset.type = "button"; reset.className = "btn tiny subtle"; reset.textContent = "QuickBooks default";
  reset.onclick = () => { _stubCols = _stubHist.columns.filter(c => c.default).map(c => c.key); try { localStorage.setItem(STUB_COLS_LS, JSON.stringify(_stubCols)); } catch {} renderVendorPage(); };
  box.appendChild(reset); wrap.appendChild(box);
  return wrap;
}
let _stubSel = new Set();   // payment ids ticked for a multi-print (owner 2026-09-22: "a box to check in case I want to select multiple payments")
async function _printStubsSelected(btn) {
  const ids = [..._stubSel]; if (!ids.length) return;
  btn.disabled = true; const was = btn.textContent;
  let done = 0, failed = [];
  for (const pid of ids) {
    btn.textContent = `Printing ${done + 1} of ${ids.length}…`;
    try {
      const res = await fetch("/api/bill-payment/stub", { method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ payment_id: pid, columns: _stubCols || undefined }) });
      const j = await res.json(); if (!j.ok) throw new Error(j.error || "print failed");
      done++;
    } catch (e) { failed.push(pid + ": " + e.message); }
  }
  toast(`${done} stub${done === 1 ? "" : "s"} printed to the vendor folder` + (failed.length ? ` · ${failed.length} failed` : ""), 5000);
  if (failed.length) console.warn("stub prints failed", failed);
  _stubSel.clear(); btn.textContent = was;
  const d = _vendorData; await _loadStubHistory(_stubHist.vendor); if (_vendorPageShowing(d)) renderVendorPage();
}
async function _renderVendorPayments(d, body, seq) {
  if (_stubHist.vendor !== d.vendor) { await _loadStubHistory(d.vendor); _stubSel.clear(); }
  if (seq !== _vpSeq || _vendorView !== "payments" || !_vendorPageShowing(d)) return;   // a newer render (or another page) took over while the history loaded
  const pays = (d.payments || []).filter(p => (!_vendorDate || !_vendorDate.active() || _vendorDate.passes(p.txn_date))
    && (!_vendorProj.trim() || _vq(_vendorProj, [...(p.projects || []), ...(p.projects || []).map(nameOf)]))
    && _vq(_vendorQ, [p.ref_no, fmtDateShort(p.txn_date), p.txn_date, p.pay_type === "CreditCard" ? "credit card" : p.pay_type, ...(p.clients || []), ...(p.projects || []),
    ...(p.bills || []).map(b => b.bill_ref), ...(p.bills || []).map(b => b.bill_date ? fmtDateShort(b.bill_date) : ""), Math.round(num(p.total_amt)), money(p.total_amt), p.voided ? "voided" : "", p.memo]));
  { const visible = new Set(pays.map(p => String(p.qbo_txn_id))); for (const id of [..._stubSel]) if (!visible.has(id)) _stubSel.delete(id); }   // a pick you can't see is not a pick
  const byPay = new Map(); for (const h of _stubHist.prints) { if (!byPay.has(h.payment_id)) byPay.set(h.payment_id, []); byPay.get(h.payment_id).push(h); }
  // toolbar: the column picker + the multi-print button (lives on the selection)
  const bar = document.createElement("div"); bar.className = "stub-bar";
  if (_stubHist.columns.length) bar.appendChild(_stubColumnPicker());
  const selBtn = document.createElement("button"); selBtn.type = "button"; selBtn.className = "btn small stub-sel-btn";
  const selLabel = () => { const n = _stubSel.size; selBtn.textContent = n ? `Print ${n} stub${n === 1 ? "" : "s"}` : "Print stubs for selected"; selBtn.disabled = !n; };
  selLabel(); selBtn.onclick = () => _printStubsSelected(selBtn); bar.appendChild(selBtn);
  body.appendChild(bar);
  if (!pays.length) { const p = document.createElement("div"); p.className = "bills-cap"; p.textContent = (d.payments || []).length ? (_vendorQ.trim() ? `No payments match "${_vendorQ.trim()}".` : "No payments match these filters.") : "No bill payments recorded this year (run the AP / bill-payments sync to pull them)."; body.appendChild(p); _renderStubOrphans(body, d.payments || [], byPay); return; }
  const scroll = document.createElement("div"); scroll.className = "table-scroll";
  const table = document.createElement("table"); table.className = "grid vp-paytable"; const thead = document.createElement("thead"), tbody = document.createElement("tbody");
  const htr = document.createElement("tr");
  // GROUPED (owner 2026-09-22): the payment is the row - ref, date, client, amount - and the bills it paid open under it
  // (GRP_KINDS "tr.vp-pay": caret in the ref, remembered state, children hidden while closed).
  { const th = document.createElement("th"); th.className = "left vp-ck"; const all = document.createElement("input"); all.type = "checkbox"; all.title = "Select every payment (voided ones can't be printed)";
    const printable = pays.filter(p => !p.voided).map(p => String(p.qbo_txn_id));
    all.checked = printable.length > 0 && printable.every(id => _stubSel.has(id));
    all.onchange = () => { if (all.checked) printable.forEach(id => _stubSel.add(id)); else printable.forEach(id => _stubSel.delete(id)); renderVendorPage(); };
    th.appendChild(all); htr.appendChild(th); }
  for (const [c, al] of [["Ref / cheque #", "left"], ["Date", "left"], ["Type", "left"], ["Client", "left"], ["Client invoice", "left"], ["Amount", "right"], ["Stub", "left"]]) { const th = document.createElement("th"); if (al === "left") th.className = "left"; th.textContent = c; htr.appendChild(th); }
  thead.appendChild(htr);
  for (const p of pays) {
    const pid = String(p.qbo_txn_id);
    const tr = document.createElement("tr"); tr.className = "vp-pay" + (p.voided ? " vp-void" : ""); tr.dataset.grpkey = "pay:" + pid;
    { const td = document.createElement("td"); td.className = "left vp-ck"; const cb = document.createElement("input"); cb.type = "checkbox";
      cb.checked = _stubSel.has(pid); cb.disabled = !!p.voided; cb.title = p.voided ? "Voided - nothing to print" : "Select for a multi-print";
      cb.onchange = () => { if (cb.checked) _stubSel.add(pid); else _stubSel.delete(pid); selLabel(); const all = thead.querySelector("input"); if (all) all.checked = pays.filter(x => !x.voided).every(x => _stubSel.has(String(x.qbo_txn_id))); };
      td.appendChild(cb); tr.appendChild(td); }
    { const rc = document.createElement("td"); rc.className = "left"; const key = document.createElement("span"); key.className = "bg-key"; key.textContent = p.ref_no || "–"; rc.appendChild(key);   // the caret lands inside .bg-key
      if (p.voided) { const v = document.createElement("span"); v.className = "stub-pill voided vp-voidpill"; v.textContent = "VOIDED"; v.title = p.memo || "Voided in QuickBooks"; rc.appendChild(v); }
      if (p.att) { const ab = attBtn("BillPayment", p.qbo_txn_id, p.att, `payment ${p.ref_no || ""}`); ab.style.marginLeft = "6px"; rc.appendChild(ab); }
      if (p.voided && p.memo && !/^voided\.?$/i.test(p.memo.trim())) { const m = document.createElement("div"); m.className = "dim vp-memo"; m.textContent = p.memo; rc.appendChild(m); }
      tr.appendChild(rc); }
    tr.appendChild(leftText(fmtDateShort(p.txn_date)));
    tr.appendChild(leftText(p.pay_type === "CreditCard" ? "Credit card" : (p.pay_type || "–")));
    { const cl = p.clients || []; const cc = leftText(cl.slice(0, 4).join(", ") + (cl.length > 4 ? ` +${cl.length - 4} more` : "") || (p.voided ? "" : "–")); if (!cl.length) cc.classList.add("dim"); cc.title = cl.join(", "); tr.appendChild(cc); }
    tr.appendChild(_vpInvTally(p.bills || []));
    { const amt = document.createElement("td"); if (p.voided) { const z = document.createElement("span"); z.className = "cell dim"; z.textContent = "voided"; amt.appendChild(z); } else amt.appendChild(moneyCell(p.total_amt)); tr.appendChild(amt); }
    { const sc = document.createElement("td"); sc.className = "left vp-stub"; const hist = byPay.get(pid) || [];
      if (!p.voided) { const b = document.createElement("button"); b.type = "button"; b.className = "btn tiny"; b.textContent = hist.length ? "Print again" : "Print stub";
        b.title = "Payment on top, the bills it paid below - the PDF lands in this vendor's folder under Accounting / Accounts Payable / Bill Payment Stubs";
        b.onclick = (e) => { e.stopPropagation(); _printStub(pid, b); }; sc.appendChild(b); }
      if (hist.length) { const hl = document.createElement("div"); hl.className = "stub-hist";
        const lines = hist.map(h => { const ln = document.createElement("div"); ln.className = "stub-histline"; ln.appendChild(_stubLink(h)); ln.appendChild(_stubPill(h.status)); return ln; });
        _vpCapped(hl, lines, 3, "prints"); sc.appendChild(hl); }
      tr.appendChild(sc); }
    tbody.appendChild(tr);
    // the expansion: one row per bill this payment paid (bill # -> QuickBooks, its project, its client, the amount applied)
    const bl = p.bills || [];
    if (!bl.length && !p.voided) { const br = document.createElement("tr"); br.className = "vp-bill"; const td = document.createElement("td"); td.colSpan = 8; td.className = "left dim"; td.textContent = p.n_bills ? `${p.n_bills} bill${p.n_bills === 1 ? "" : "s"} (details not loaded - run the AP sync)` : "no bills on this payment"; br.appendChild(td); tbody.appendChild(br); }
    for (const b of bl) {
      const br = document.createElement("tr"); br.className = "vp-bill";
      br.appendChild(document.createElement("td"));
      { const td = document.createElement("td"); td.className = "left"; const a = document.createElement("a"); a.href = qboUrl("bill", b.bill_id); a.target = "_blank"; a.rel = "noopener"; a.className = "qbo-link"; a.textContent = b.bill_ref || ("bill " + b.bill_id); a.title = "Open this bill in QuickBooks"; td.appendChild(a); br.appendChild(td); }
      { const dc = leftText(b.bill_date ? fmtDateShort(b.bill_date) : "–"); dc.title = "Bill date"; if (!b.bill_date) dc.classList.add("dim"); br.appendChild(dc); }   // the bill's own date (owner 2026-09-23)
      br.appendChild(_vpProjCell(b.projects || []));
      { const cl = b.clients || []; const cc = leftText(cl.join(", ") || "–"); if (!cl.length) cc.classList.add("dim"); br.appendChild(cc); }
      br.appendChild(_vpInvCell(b));
      { const td = document.createElement("td"); td.appendChild(moneyCell(b.amount)); br.appendChild(td); }
      br.appendChild(document.createElement("td"));
      tbody.appendChild(br);
    }
  }
  table.appendChild(thead); table.appendChild(tbody); scroll.appendChild(table); body.appendChild(scroll);
  _renderStubOrphans(body, d.payments || [], byPay);   // "no longer in the list" = not in the vendor's payments at all, never "filtered out"
}
// The client invoice a bill is billed through and whether the client paid it (owner 2026-09-23: "also show the
// invoice paid status"). Live from QuickBooks' invoice when the ledger has it, else the Bill Tracker's paid date.
const VP_INV_OTHER = {
  awaiting: ["Not invoiced yet", "On the Bill Tracker as Awaiting Invoice - not billed to the client yet"],
  noproject: ["No project #", "The bill carries no project #, so no client invoice"],
  untracked: ["Not tracked", "Not on the Bill Tracker (older bills, subs, overhead) - no client invoice on file"],
};
function _vpInvCell(b) {
  const td = document.createElement("td"); td.className = "left vp-inv";
  if (!b.invoice_state || VP_INV_OTHER[b.invoice_state]) {
    const [t, tip] = VP_INV_OTHER[b.invoice_state] || ["–", "No client invoice on file"];
    if (b.invoice_state === "awaiting") { const pill = document.createElement("span"); pill.className = "stub-pill vp-inv-open"; pill.textContent = t; td.appendChild(pill); }
    else { td.textContent = t; td.classList.add("dim"); }
    td.title = tip; return td;
  }
  const pill = document.createElement("span"); pill.className = "stub-pill " + (b.invoice_state === "paid" ? "vp-inv-paid" : "vp-inv-open");
  pill.textContent = b.invoice_state === "paid" ? "Paid" + (b.invoice_paid_on ? " " + fmtDateShort(b.invoice_paid_on) : "") : "Not paid";
  if (b.invoice_no) { const n = document.createElement("span"); n.textContent = "#" + b.invoice_no + " "; td.appendChild(n); }
  td.appendChild(pill); td.title = b.invoice_state === "paid" ? "The client paid this invoice" : "The client has not paid this invoice yet";
  return td;
}
function _vpInvTally(bills) {
  const withInv = bills.filter(b => b.invoice_state === "paid" || b.invoice_state === "open" || b.invoice_state === "awaiting"), paid = withInv.filter(b => b.invoice_state === "paid").length;
  const td = document.createElement("td"); td.className = "left vp-inv";
  if (!withInv.length) { td.textContent = bills.length ? "–" : ""; td.classList.add("dim"); return td; }
  const pill = document.createElement("span"); pill.className = "stub-pill " + (paid === withInv.length ? "vp-inv-paid" : "vp-inv-open");
  pill.textContent = paid === withInv.length ? (withInv.length === 1 ? "Paid" : `All ${paid} paid`) : `${paid} of ${withInv.length} paid`;
  td.appendChild(pill); return td;
}
// Stubs printed for payments the ledger's payment list no longer carries (deleted or voided in QuickBooks, or outside
// this year's window): the history keeps the stub as printed, so they are listed here, never lost.
function _renderStubOrphans(body, pays, byPay) {
  const have = new Set(pays.map(p => String(p.qbo_txn_id)));
  const orphans = [...byPay.entries()].filter(([pid]) => !have.has(pid));
  if (!orphans.length) return;
  const h = document.createElement("h3"); h.className = "stub-orph-h"; h.textContent = "Printed stubs for payments no longer in this list"; body.appendChild(h);
  const cap = document.createElement("div"); cap.className = "bills-cap"; cap.textContent = "Changed, voided or deleted in QuickBooks after printing (or outside this year's window). Each stub opens as it went out."; body.appendChild(cap);
  const scroll = document.createElement("div"); scroll.className = "table-scroll";
  const table = document.createElement("table"); table.className = "grid"; const thead = document.createElement("thead"), tbody = document.createElement("tbody");
  const htr = document.createElement("tr");
  for (const [c, al] of [["Paid", "left"], ["Ref / cheque #", "left"], ["Type", "left"], ["Bills", "right"], ["Amount as printed", "right"], ["QBO now", "left"], ["Prints", "left"]]) { const th = document.createElement("th"); if (al === "left") th.className = "left"; th.textContent = c; htr.appendChild(th); }
  thead.appendChild(htr);
  for (const [pid, hist] of orphans) {
    const h0 = hist[0]; const tr = document.createElement("tr");
    tr.appendChild(leftText(fmtDateShort(h0.txn_date)));
    tr.appendChild(qboLinkCell(h0.ref || "–", null, ""));
    tr.appendChild(leftText(h0.method || "–"));
    { const td = document.createElement("td"); td.textContent = h0.n_bills != null ? String(h0.n_bills) : "–"; tr.appendChild(td); }
    { const td = document.createElement("td"); td.appendChild(moneyCell(h0.total)); tr.appendChild(td); }
    { const td = document.createElement("td"); td.className = "left"; td.appendChild(_stubPill(h0.status));
      if (h0.status === "changed" && h0.qbo_total_now != null) { const s = document.createElement("span"); s.className = "dim"; s.textContent = " now " + money(h0.qbo_total_now); td.appendChild(s); }
      if (h0.status === "deleted" && h0.deleted_at) { const s = document.createElement("span"); s.className = "dim"; s.textContent = " " + fmtDateShort(h0.deleted_at); td.appendChild(s); }
      tr.appendChild(td); }
    { const td = document.createElement("td"); td.className = "left vp-stub"; const hl = document.createElement("div"); hl.className = "stub-hist";
      const lines = hist.map(h => { const ln = document.createElement("div"); ln.className = "stub-histline"; ln.appendChild(_stubLink(h)); return ln; });
      _vpCapped(hl, lines, 3, "prints"); td.appendChild(hl);
      const b = document.createElement("button"); b.type = "button"; b.className = "btn tiny subtle"; b.textContent = "Print again"; b.title = "Re-print from the mirror's copy of the payment (kept even after a QuickBooks delete)";
      b.onclick = (e) => { e.stopPropagation(); _printStub(pid, b); }; td.appendChild(b);
      tr.appendChild(td); }
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
    tr.appendChild(qboLinkCell(b.bill_ref, qboBillHref(b.qbo_link), "Open this bill in QuickBooks")); if (b.att) { const _ab = attBtn("Bill", b.bill_id || (qboBillHref(b.qbo_link) || "").replace(/.*txnId=(\d+).*/, "$1"), b.att, `${b.vendor || ""} · bill ${b.bill_ref || ""}`); _ab.style.marginLeft = "6px"; tr.lastElementChild.appendChild(_ab); }
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
      tr.appendChild(qboLinkCell(b.bill_ref, qboBillHref(b.qbo_link), "Open this bill in QuickBooks")); if (b.att) { const _ab = attBtn("Bill", b.bill_id || (qboBillHref(b.qbo_link) || "").replace(/.*txnId=(\d+).*/, "$1"), b.att, `${b.vendor || ""} · bill ${b.bill_ref || ""}`); _ab.style.marginLeft = "6px"; tr.lastElementChild.appendChild(_ab); }
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
  const cols = [["Pick", "left"], ["Project", "left"], ["Invoice #", "left"], ["Date", "left"], ["Due", "left"], ["Memo", "left"], ["Open balance", "right"], ["Invoice total", "right"], ["Last action", "left"], ["Next follow-up", "left"], ["Collections note", "left"]];
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
    tr.onclick = (e) => { if (e.target.closest("a") || e.target.closest("input")) return; openInvoicePage(i); };
    { const kc = document.createElement("td"); kc.className = "left inv-pick"; const cb = document.createElement("input"); cb.type = "checkbox"; cb.checked = invPick.has(invKey(i)); cb.title = "Pick this invoice for the collections report";
      cb.onclick = (e) => e.stopPropagation(); cb.onchange = () => { if (cb.checked) invPick.add(invKey(i)); else invPick.delete(invKey(i)); _invPickUpdate(); }; kc.appendChild(cb); tr.appendChild(kc); }
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
    const ob = document.createElement("td"); ob.className = "right amt-box";
    if (paid) { ob.textContent = "–"; ob.classList.add("dim"); }
    else { ob.textContent = money(oiBal(i)); if (i.days_past_due != null && i.days_past_due > 0) { ob.style.color = "var(--neg)"; ob.title = i.days_past_due + " days past due"; } }
    tr.appendChild(ob);
    { const tc = rightText(money(i.amount)); tc.classList.add("amt-box", "amt-box-soft"); tr.appendChild(tc); }
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
    const hr = document.createElement("tr"); hr.className = "inv-client" + (expanded ? " on" : ""); hr.style.cursor = "pointer";
    hr.title = expanded ? "Click to collapse" : "Click to expand";
    const htd = document.createElement("td"); htd.colSpan = cols.length;
    const caret = document.createElement("span"); caret.className = "bg-caret"; caret.textContent = expanded ? "▾ " : "▸ ";
    { const gcb = document.createElement("input"); gcb.type = "checkbox"; gcb.className = "inv-pick-all"; gcb.title = "Pick every invoice of this client for the collections report";
      gcb.checked = g.rows.every(x => invPick.has(invKey(x))); gcb.indeterminate = !gcb.checked && g.rows.some(x => invPick.has(invKey(x)));
      gcb.onclick = (e) => e.stopPropagation(); gcb.onchange = () => { for (const x of g.rows) { if (gcb.checked) invPick.add(invKey(x)); else invPick.delete(invKey(x)); } renderOpenInvoices(); }; caret.appendChild(gcb); }
    const nm = document.createElement("span"); nm.className = "g-cust"; nm.textContent = g.client;
    const ad = invClientAvgDays(g.client);
    const sub = document.createElement("span"); sub.className = "g-sub"; sub.hidden = true;   // (the metrics grid replaced the text run)
    const cellG = document.createElement("div"); cellG.className = "bg-cell"; const leftG = document.createElement("span"); leftG.className = "bg-left"; leftG.appendChild(caret); leftG.appendChild(nm); cellG.appendChild(leftG);
    bandMetrics(cellG, [[g.rows.length, "invoices"], [money(g.open), "open", (g.open > 0.005 ? "neg" : "") + " boxed"], [money(g.billed), "billed"], [ad != null ? ad + "d" : "–", "avg days to pay"]]);
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
      const pickRow = (i) => { const r = amtRow(i); if (invPick.has(invKey(i))) r.classList.add("picked"); return r; };
      for (const p of porder) {
        const pg = inP(p);
        if (pg.length === 1 && /^RP/i.test(p)) { tbody.appendChild(pickRow(pg[0])); continue; }   // one RP invoice on one job: no project band (owner 2026-09-08)
        tbody.appendChild(invSubBand(p, nameOf(p), pTotal(p), pg.length, cols.length));
        pg.forEach((i, ix) => { const r = pickRow(i); r.classList.add("in-proj"); if (ix === pg.length - 1) r.classList.add("proj-last"); tbody.appendChild(r); });
      }
    } else {
      for (const i of g.rows) { const r = amtRow(i); if (invPick.has(invKey(i))) r.classList.add("picked"); tbody.appendChild(r); }
    }
  }
  const tr = document.createElement("tr"); tr.className = "inv-total-row";
  const td0 = document.createElement("td"); td0.className = "left"; td0.colSpan = 6; td0.textContent = "TOTAL"; tr.appendChild(td0);
  tr.appendChild(rightText(money(totOpen)));
  tr.appendChild(rightText(money(totBilled)));
  tr.appendChild(document.createElement("td")); tr.appendChild(document.createElement("td")); tr.appendChild(document.createElement("td"));
  tbody.appendChild(tr);
  updateInvCollapseBtn(); _invPickUpdate();
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
  { const el = $("#invAsOf"); if (el) el.textContent = (D.as_of ? "aged today " + fmtDate(D.as_of) : ""); }   // both views - load times live in the sync pill
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
  { const el = $("#invAsOf"); if (el) el.textContent = (D.as_of ? "aged today " + fmtDate(D.as_of) : ""); }
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
    const gtr = document.createElement("tr"); gtr.className = "bill-group" + (collapsed ? "" : " on"); gtr.style.cursor = "pointer";   // .on = the client you are in (accent); the rest stay neutral (owner 2026-09-08)
    gtr.title = collapsed ? "Click to expand" : "Click to collapse";
    const gtd = document.createElement("td"); gtd.colSpan = cols.length;
    const cell = document.createElement("div"); cell.className = "bg-cell";   // flex on the div, not the td
    const left = document.createElement("span"); left.className = "bg-left";
    const caret = document.createElement("span"); caret.className = "bg-caret"; caret.textContent = collapsed ? "▸" : "▾";
    const key = document.createElement("span"); key.className = "bg-key"; key.textContent = k;
    left.appendChild(caret); left.appendChild(key);
    cell.appendChild(left);
    bandMetrics(cell, [[money(gOpen), "open", (gOpen > 0.005 ? "neg" : "") + " boxed"], [g.length, "invoices"]]);
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
        if (pg.length === 1 && /^RP/i.test(p)) { tbody.appendChild(invRow(pg[0], buckets)); continue; }   // one RP invoice on one job: just the invoice, no project band (owner 2026-09-08)
        tbody.appendChild(invSubBand(p, nameOf(p), pTotal(p), pg.length, cols.length));
        pg.forEach((i, ix) => { const r = invRow(i, buckets); r.classList.add("in-proj"); if (ix === pg.length - 1) r.classList.add("proj-last"); tbody.appendChild(r); });   // the project reads as one BOX
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
  if (invPick.has(invKey(i))) tr.classList.add("picked");
  { const cells = tr.querySelectorAll("td.ag"); for (const c of cells) if (c.textContent.trim() && c.textContent.trim() !== "–") c.classList.add("amt-box"); }   // the open amount in a black line (owner 2026-09-08)
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
  if (inv.att) { const ab = attBtn("Invoice", inv.qbo_txn_id, inv.att, `${inv.customer || ""} · invoice ${docn}`); ab.style.marginLeft = "6px"; td.appendChild(ab); }
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
// WHY the live QuickBooks figure and the WIP report differ - the first honest explanation is the lines dated after
// the report date (owner: "mistrusting on where numbers are pulled from"); what is left is coding / scope differences.
function _whySince(p, kind) {
  const sw = p.since_wip || {}; if (!sw.report_date) return "";
  const rd = fmtDateShort(sw.report_date);
  if (kind === "cost") { const gap = num(p.cost) - num(sw.wip_cost); if (Math.abs(gap) < 1) return ` (${rd}) · same`;
    const after = num(sw.cost_amount); const rest = gap - after;
    return ` (${rd}) · QuickBooks ${gap >= 0 ? "+" : "-"}${money(Math.abs(gap))}: ${sw.cost_lines} line${sw.cost_lines === 1 ? "" : "s"} dated after the report ${money(after)}${Math.abs(rest) >= 1 ? `, ${rest >= 0 ? "+" : "-"}${money(Math.abs(rest))} coding / scope` : ""}`; }
  const inv = num(sw.invoices); return inv ? ` (${rd}) · ${inv} invoice${inv === 1 ? "" : "s"} ${money(sw.billed_amount)} dated after the report` : ` (${rd})`;
}
async function openProjectPage(pn) {
  if (!pn || !/^(MFD|CP|RP)\d/i.test(String(pn))) { if (pn) toast(`"${pn}" is not a project # - nothing to open`); return; }   // e.g. the "(multiple)" bucket
  pn = String(pn).toUpperCase();
  if (_ppLeaveBlocked()) return;   // unsaved pay ticks on the page that is open: Save or Discard first
  const r0 = (ALL || []).find(x => x.project_no === pn) || {};
  _recSave({ k: "project", id: pn });
  openRecord(pn + (r0.project_name ? " · " + r0.project_name : ""), [r0.division, r0.status ? "WIP status " + r0.status : ""].filter(Boolean).join(" · "));
  const body = $("#recordBody"); body.innerHTML = ""; skeletonInto(body, 6);
  let d;
  try { d = await (await fetch(`/api/project/page?no=${encodeURIComponent(pn)}`)).json(); }
  catch (e) { body.textContent = "could not load this project"; return; }
  if (!d || !d.ok) { body.textContent = (d && d.error) || "no data for this project"; return; }
  // view = "all" (every bill on the job, with the draw each sits under) | one draw's key. The Coverage table and the
  // draw boxes stay on top whatever is open (owner 2026-09-15). filter = all | unpaid (all by default - "i need to be
  // able to see all bills"). sort = vendor (A-Z) | amount (band total) | code | date. payMode = the pay-run controls,
  // off until "Pay bills" is clicked; payDraft = ticks not yet saved (bill_id -> selected).
  _pp = { d, pn, view: "all", openV: new Set(), filter: "all", sort: "vendor", payMode: false, payDraft: new Map(), isRp: /^RP/i.test(pn), notesOpen: false };
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
  const s1 = sec("Profit & Loss", `projected = WIP report ${r0.report_date ? fmtDate(r0.report_date) : "–"} · actual = QuickBooks`);
  s1.appendChild(_ppProjectedVsActual(p, r0, pn));   // what we projected next to what actually happened (owner 2026-09-16: "put what we projected and what the actual are side by side")
  if ((d.rulings || []).length) {   // the owner's standing rulings (job_rulings.json): the why, so nobody re-flags it
    const rb = document.createElement("div"); rb.className = "pp-unlock pp-rulings";
    rb.innerHTML = `<div class="pp-unlock-h">Known - the owner ruled on this job</div>` + d.rulings.map(x =>
      `<div class="pp-unlock-b"><b>${_ge(String(x.kind || "note").toUpperCase())}</b> ${_ge(x.note || "")}`
      + (x.amount != null ? ` · <b>${_ge(money(x.amount))}</b>` : "") + (x.line ? ` · ${_ge(x.line)}` : "")
      + (x.source ? ` · ${_ge(x.source)}` : "") + (x.on ? ` · ruled ${_ge(fmtDate(x.on))}` : "") + `</div>`).join("");
    s1.appendChild(rb);
  }
  const acts1 = document.createElement("div"); acts1.className = "ip-actions";
  if (r0.project_no) { const dr = document.createElement("button"); dr.className = "btn small"; dr.textContent = "WIP row detail"; dr.onclick = () => openDetail(r0); acts1.appendChild(dr); }
  s1.appendChild(acts1);
  const plWrap = document.createElement("div"); plWrap.className = "ip-top pp-pnl"; plWrap.appendChild(buildPnlGroup(pn)); s1.appendChild(plWrap);   // 3 columns (owner: save vertical space)
  // ── 2. how we get funded ──
  const F = d.funding || {}, nx = F.next_draw;
  const isRp = /^RP/i.test(pn);
  const nInv = d.draws.filter(x => !x.no_draw).length, nNo = d.draws.filter(x => x.no_draw).length;   // the same count the P&L block shows, plus the not-yet-drawn bucket named
  const owes = d.draws.reduce((s, x) => s + num(x.ar_open), 0);
  const s2 = sec("How we get funded", (isRp ? `${nInv} scope${nInv === 1 ? "" : "s"} invoiced${nNo ? " + bills not yet invoiced" : ""} - an RP job has no draws: each invoice is a scope, its costs the bills dated up to it` : `${nInv} draw${nInv === 1 ? "" : "s"} invoiced${nNo ? ` + ${nNo} not yet drawn` : ""}`)
    + ` · GC owes ${money(owes)}`);
  // Colour encodes OUR side only (owner 2026-09-10: "why is the box red?" - it was red on every
  // awaiting-funding draw, even when nothing blocked it). Red only when we owe money on an earlier
  // draw; amber when earlier bills show $0 open / no pay date (confirm); green otherwise.
  const blk = nx ? (F.blockers || []) : [];
  const blocked = nx && (F.blockers_total || 0) > 0.005;
  const caution = nx && blk.length && !blocked;
  const unlock = document.createElement("div"); unlock.className = "pp-unlock" + (blocked ? "" : caution ? " caution" : " ok");
  if (nx) {
    unlock.innerHTML = `<div class="pp-unlock-h">Next money in: <b>${_ge(nx.label.split(/\s+[—–-]\s+/)[0])}${_ge(drawTag(nx))}</b> · GC owes <b>${_ge(money(nx.ar_open))}</b>${nx.ar_date ? " · invoiced " + _ge(fmtDate(nx.ar_date)) : ""}${drawSpan(nx) ? " · covers " + _ge(drawSpan(nx)) : ""}</div>`
      + (blk.length ? (blocked
            ? `<div class="pp-unlock-b">Blocked by <b>${blk.length}</b> unpaid bill${blk.length === 1 ? "" : "s"} on earlier draws · <b>${_ge(money(F.blockers_total))}</b> to pay (their unconditional waivers release this draw)</div>`
            : `<div class="pp-unlock-b"><b>${blk.length}</b> bill${blk.length === 1 ? "" : "s"} on earlier draws show no payment date yet ($0 open) - confirm they are paid and collect the waivers, then this draw is clear on our side</div>`)
                    : `<div class="pp-unlock-b ok">No unpaid bills on earlier draws - nothing on our side blocks this draw${F.own_unpaid > 0.005 ? `; ${_ge(money(F.own_unpaid))} of its own bills still to pay once funded` : ""}.</div>`);
  } else unlock.innerHTML = `<div class="pp-unlock-h ok">Nothing outstanding - the GC has paid every draw on file.</div>`;
  // owner 2026-09-10: an × to dismiss this message, and a button to bring it back (remembered per browser).
  const _fundHidden = () => { try { return localStorage.getItem("ppFundingHidden") === "1"; } catch (e) { return false; } };
  const reopen = document.createElement("button"); reopen.type = "button"; reopen.className = "btn small pp-funding-show"; reopen.textContent = "Show funding status";
  const xb = document.createElement("button"); xb.type = "button"; xb.className = "pp-x"; xb.title = "Hide the funding status"; xb.setAttribute("aria-label", "Hide the funding status"); xb.textContent = "×";
  xb.onclick = () => { try { localStorage.setItem("ppFundingHidden", "1"); } catch (e) { /* private mode */ } unlock.hidden = true; reopen.hidden = false; };
  reopen.onclick = () => { try { localStorage.removeItem("ppFundingHidden"); } catch (e) { /* private mode */ } unlock.hidden = false; reopen.hidden = true; };
  unlock.appendChild(xb);
  unlock.hidden = _fundHidden(); reopen.hidden = !unlock.hidden;
  s2.appendChild(unlock); s2.appendChild(reopen);
  // the tools row: the bills filter (big, with counts - owner 2026-09-15: "this toggle ... be more present"), the sort,
  // Expand / Collapse, and "Pay bills" off to the right - the pay-run controls stay hidden until it is clicked
  const tools = document.createElement("div"); tools.className = "ip-tools pp-tools"; _pp.toolsEl = tools;   // re-homed right above the bill list on every render (#9)
  const fseg = document.createElement("div"); fseg.className = "seg big"; fseg.id = "ppFilterSeg"; tools.appendChild(fseg);
  const sl = document.createElement("span"); sl.className = "dim pp-sortlab"; sl.textContent = "Sort by"; tools.appendChild(sl);
  const sseg = document.createElement("div"); sseg.className = "seg"; sseg.id = "ppSortSeg"; tools.appendChild(sseg);
  const exp = document.createElement("button"); exp.type = "button"; exp.className = "btn small"; exp.id = "ppExpand"; exp.onclick = () => { _pp.openV = _pp.openV.has("*") ? new Set() : new Set(["*"]); _renderPpDraws(); }; tools.appendChild(exp);
  const sp = document.createElement("span"); sp.className = "pp-spacer"; tools.appendChild(sp);
  const nb = document.createElement("button"); nb.type = "button"; nb.className = "btn small"; nb.id = "ppNotesBtn"; nb.onclick = () => { _pp.notesOpen = !_pp.notesOpen; _ppRenderNotes(); }; tools.appendChild(nb);   // owner 2026-09-16: "give me a button to add notes"
  const payBtn = document.createElement("button"); payBtn.type = "button"; payBtn.className = "btn small pp-paybtn"; payBtn.id = "ppPayBtn"; payBtn.onclick = _ppTogglePay; tools.appendChild(payBtn);
  s2.appendChild(tools);
  const notesBox = document.createElement("div"); notesBox.className = "pp-notes"; notesBox.id = "ppNotes"; notesBox.hidden = true; s2.appendChild(notesBox);
  const paybar = document.createElement("div"); paybar.className = "pp-paybar"; paybar.id = "ppPayBar"; paybar.hidden = true; s2.appendChild(paybar);
  const host = document.createElement("div"); host.id = "ppDraws"; s2.appendChild(host);
  _renderPpDraws();
  // ── 3. audit findings on this job (the Audit tab, filtered to this project) ──
  { const au = d.audits || [];
    const sA = sec("Audit findings on this job", au.length ? `${au.length} to fix · Bill Tracker audits` : "none open"); sA.classList.add("fold-sec");   // folded (owner 2026-09-23)
    if (!au.length) { const p0 = document.createElement("div"); p0.className = "bills-cap"; p0.textContent = "Nothing flagged on this job in the Bill Tracker audits."; sA.appendChild(p0); }
    else {
      const scroll = document.createElement("div"); scroll.className = "table-scroll"; scroll.style.padding = "0 18px 12px";
      const t = document.createElement("table"); t.className = "grid pp-audit";
      t.innerHTML = "<thead><tr><th class='left'>Bill #</th><th class='left'>Vendor</th><th class='left'>Date</th><th class='left'>Code</th><th class='right'>Amount</th><th class='left'>What the clerk wrote</th><th class='left'>Detail</th></tr></thead>";
      const tb = document.createElement("tbody");
      const byI = new Map(); for (const f of au) { if (!byI.has(f.issue)) byI.set(f.issue, []); byI.get(f.issue).push(f); }
      for (const [iss, list] of [...byI].sort((a, b) => b[1].length - a[1].length)) {
        const g = document.createElement("tr"); g.className = "bill-group"; const gd = document.createElement("td"); gd.colSpan = 7;
        const cell = document.createElement("div"); cell.className = "bg-cell"; const k = document.createElement("span"); k.className = "bg-key"; k.textContent = iss; cell.appendChild(k);
        bandMetrics(cell, [[list.length, "bills"], [money(list.reduce((s0, f) => s0 + num(f.amount), 0)), "amount"], [list[0].group || "", "audit"]]);
        gd.appendChild(cell); g.appendChild(gd); tb.appendChild(g);
        for (const f of list) {
          const tr = document.createElement("tr");
          const idm = (f.url || "").match(/txnId=(\d+)/);
          const lc = qboLinkCell(f.bill_no || "–", f.url, "Open this bill in QuickBooks"); if (f.att && idm) { const ab = attBtn("Bill", idm[1], f.att, `${f.vendor || ""} · bill ${f.bill_no || ""}`); ab.style.marginLeft = "6px"; lc.appendChild(ab); } tr.appendChild(lc);
          tr.appendChild(leftText(f.vendor || "–")); tr.appendChild(leftText(f.date ? fmtDateShort(f.date) : "–"));
          { const cc = document.createElement("td"); cc.className = "left"; if (f.cost_code) { const ch = document.createElement("span"); ch.className = "codechip"; ch.textContent = f.cost_code; cc.appendChild(ch); } else cc.textContent = "–"; tr.appendChild(cc); }
          { const ac = document.createElement("td"); ac.className = "ip-amt"; ac.appendChild(moneyCell(f.amount)); tr.appendChild(ac); }
          { const mc = leftText(f.memo || "–"); mc.className += " inv-memo"; mc.title = f.memo || ""; tr.appendChild(mc); }
          { const dc = leftText(f.detail || "–"); dc.className += " inv-memo"; dc.title = f.detail || ""; tr.appendChild(dc); }
          tb.appendChild(tr);
        }
      }
      t.appendChild(tb); scroll.appendChild(t); sA.appendChild(scroll);
      const go = document.createElement("div"); go.className = "ip-actions"; const gb = document.createElement("button"); gb.className = "btn small"; gb.textContent = "Open the Audit tab"; gb.onclick = () => { closeRecord(); setTab("accounting"); }; go.appendChild(gb); sA.appendChild(go);
    }
  }
  // ── 4. bills + links ──
  { const sL = sec("Change log · contract, COs, ETC, billed, costs", "every change the WIP writer made, with its source, plus the review answers"); sL.classList.add("fold-sec");   // folded (owner 2026-09-23)
    const box = document.createElement("div"); box.className = "ip-audit"; sL.appendChild(box); fillAuditInto(box, pn); }
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
// What we projected (the WIP master: contract, ETC) next to what actually happened (QuickBooks: billed, costs), row by
// row, the big numbers first (owner 2026-09-16: "remove the big blocks ... make the numbers that are important jump
// out like billed to date, costs to date, put what we projected and what the actual are side by side").
function _ppProjectedVsActual(p, r0, pn) {
  const wrap = document.createElement("div"); wrap.className = "pp-pva-wrap";
  const t = document.createElement("table"); t.className = "pp-pva";
  t.innerHTML = `<thead><tr><th class="left"></th><th class="right">Projected</th><th class="right">Actual</th><th class="left">Difference</th></tr></thead>`;
  const tb = document.createElement("tbody");
  const contract = num(p.contract || r0.total_contract_price), etc = num(r0.estimated_total_costs);
  const ret = num(p.retainage), netB = num(p.net_billed), cost = num(p.cost);
  // actual billed = the QuickBooks invoices (net + retainage) when they run past the WIP report's figure (invoices dated after the report)
  // gross billed = the QuickBooks invoices before retainage (server, the Excel P&L's basis); the WIP report's figure rides along
  const billedG = num(p.billed_gross), wipG = num(p.wip_billed_gross), billedNote = wipG && Math.abs(wipG - billedG) > 0.5 ? `WIP report ${money(wipG)}` : "";
  const rate = p.overhead_rate != null ? num(p.overhead_rate) : (/^MFD/i.test(pn) ? 0.09 : 0.10);
  const ohP = contract * rate, gpP = contract - etc, netP = gpP - ohP;
  // the REAL net profit = GROSS billed - costs - overhead (owner 2026-09-23: "use the gross total billed as the factor")
  const gpA = billedG - cost, ohA = num(p.overhead), netA = num(p.net);
  const pctOf = (v, base) => base ? ` (${(v / base * 100).toFixed(1)}%)` : "";
  // one row = label · projected · actual (bold) · the difference as a number, the % off, and a little bar of how far off
  // the projection it is (owner 2026-09-16: "real difference conditional format cell little graph in the cell")
  const row = (label, proj, act, delta, opts = {}) => {
    const tr = document.createElement("tr"); if (opts.cls) tr.className = opts.cls;
    const l = document.createElement("td"); l.className = "left pp-pva-lab"; l.textContent = label; if (opts.sub) l.title = opts.sub; tr.appendChild(l);   // one line a row: the note is the hover (owner 2026-09-23)
    const a = document.createElement("td"); a.className = "right pp-pva-proj"; a.textContent = proj == null ? "" : proj; tr.appendChild(a);
    const b = document.createElement("td"); b.className = "right pp-pva-act" + (opts.big ? " big" : ""); b.textContent = act; if (opts.actSub) { const s = document.createElement("small"); s.textContent = opts.actSub; b.appendChild(s); } tr.appendChild(b);
    const c = document.createElement("td"); c.className = "left pp-pva-diff";
    if (delta && delta.base) {   // delta = {v: actual - projected, base: projected, good: +1 (over is good) | -1 (over is bad) | 0 (neutral), word}
      const pct = delta.v / delta.base * 100, good = delta.good === 0 ? "" : ((delta.v >= 0) === (delta.good > 0) ? "pos" : "neg");
      c.classList.add(good || "neutral");
      const txt = document.createElement("span"); txt.className = "pp-diff-txt"; txt.textContent = `${delta.v >= 0 ? "+" : "−"}${money(Math.abs(delta.v))} · ${pct >= 0 ? "+" : ""}${pct.toFixed(1)}%${delta.word ? " · " + delta.word : ""}`;
      const bar = document.createElement("span"); bar.className = "pp-bar"; const fill = document.createElement("i"); fill.style.width = Math.min(100, Math.abs(pct)).toFixed(1) + "%"; bar.appendChild(fill); bar.title = `${Math.abs(pct).toFixed(1)}% off the projection`;
      c.appendChild(txt); c.appendChild(bar);
    } else if (delta && delta.word) { c.textContent = delta.word; if (delta.cls) c.classList.add(delta.cls); }
    tr.appendChild(c); tb.appendChild(tr);
  };
  const left = contract - billedG;
  row("Contract · Gross billed", money(contract), money(billedG), { v: billedG - contract, base: contract, good: 0, word: left > 0.5 ? `${money(left)} left to bill` : (left < -0.5 ? "billed past the contract" : "billed out") },
      { sub: "gross billed = every QuickBooks invoice before retainage" + (r0.approved_cos ? ` · contract incl. COs ${money(r0.approved_cos)}` : ""), actSub: billedNote });
  row("Retained", null, "(" + money(ret) + ")", null, { sub: "retainage the GC is holding - billed, not yet collected" });
  row("Net billed", null, money(netB), p.billed_gap ? { word: `WIP report shows ${money(p.billed_gap)} more - Resync`, cls: "warn" } : null, { sub: "gross billed - retained = what the GC pays now" });
  const over = cost - etc;
  row("ETC · Costs to date", money(etc), money(cost), etc ? { v: over, base: etc, good: -1, word: (over > 0.5 ? "over budget" : "under budget") + ` · ${(cost / etc * 100).toFixed(1)}% complete` } : null);
  row("Gross profit", money(gpP) + pctOf(gpP, contract), money(gpA) + pctOf(gpA, billedG), etc && billedG ? { v: gpA - gpP, base: Math.abs(gpP) || contract, good: 1 } : null, { sub: "projected: contract - ETC · actual: gross billed - costs" });
  row(`Overhead · ${+(rate * 100).toFixed(2)}% of contract`, "(" + money(ohP) + ")", "(" + money(ohA) + ")", null, { sub: /^MFD/i.test(pn) ? "MFD rate" : "company rate" });
  row("Net profit", money(netP) + pctOf(netP, contract), money(netA) + (p.net_pct != null ? ` (${(p.net_pct * 100).toFixed(1)}%)` : ""), etc && billedG ? { v: netA - netP, base: Math.abs(netP) || contract, good: 1 } : null, { cls: "pp-pva-total", sub: "gross billed - costs - overhead" });
  t.appendChild(tb); wrap.appendChild(t);
  return wrap;
}
// Notes on the job (owner 2026-09-16: "give me a button to add notes") - free text, stamped, an optional "about" prefilled
// with the draw / scope that is open. Stored in the ledger (project_note), never QuickBooks.
function _ppRenderNotes() {
  const box = $("#ppNotes"), btn = $("#ppNotesBtn"); if (!box || !btn) return;
  const notes = _pp.d.notes || [];
  btn.textContent = notes.length ? `Notes · ${notes.length}` : "Add a note"; btn.classList.toggle("on", _pp.notesOpen);
  box.hidden = !_pp.notesOpen; box.innerHTML = ""; if (!_pp.notesOpen) return;
  const cur = _pp.view === "all" ? null : _pp.d.draws.find(x => x.matched_invoice === _pp.view);
  const form = document.createElement("div"); form.className = "pp-note-form";
  const about = document.createElement("input"); about.type = "text"; about.className = "pp-note-about"; about.placeholder = "about (optional)"; about.value = cur ? _ppTitle(cur) : ""; about.maxLength = 200;
  const ta = document.createElement("textarea"); ta.className = "pp-note-text"; ta.placeholder = "Write the note…"; ta.rows = 2;
  const save = document.createElement("button"); save.type = "button"; save.className = "btn small primary"; save.textContent = "Save note";
  save.onclick = async () => { const text = ta.value.trim(); if (!text) { toast("Write the note first"); ta.focus(); return; } save.disabled = true;
    let r; try { r = await (await fetch("/api/project/note", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ project_no: _pp.pn, about: about.value.trim(), text }) })).json(); } catch (e) { r = { error: String(e) }; }
    save.disabled = false; if (!(r && r.ok)) { toast("NOT saved - " + ((r && r.error) || "could not reach the server"), 5000); return; }
    _pp.d.notes = [r.note, ...(_pp.d.notes || [])]; toast("Note saved ✓"); _ppRenderNotes(); };
  form.appendChild(about); form.appendChild(ta); form.appendChild(save); box.appendChild(form);
  if (!notes.length) { const p = document.createElement("div"); p.className = "dim pp-note-empty"; p.textContent = "No notes on this job yet."; box.appendChild(p); }
  for (const n of notes) {
    const row = document.createElement("div"); row.className = "pp-note";
    const h = document.createElement("div"); h.className = "pp-note-h"; h.innerHTML = `<b>${_ge(fmtDate(n.at, true))}</b>${n.about ? ` · ${_ge(n.about)}` : ""}`;
    const x = document.createElement("button"); x.type = "button"; x.className = "btn tiny subtle"; x.textContent = "delete"; x.onclick = async () => { if (!confirm("Delete this note?")) return;
      let r; try { r = await (await fetch("/api/project/note/delete", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ id: n.id }) })).json(); } catch (e) { r = { error: String(e) }; }
      if (!(r && r.ok)) { toast("could not delete"); return; } _pp.d.notes = (_pp.d.notes || []).filter(m => m.id !== n.id); _ppRenderNotes(); };
    h.appendChild(x); row.appendChild(h);
    const tx = document.createElement("div"); tx.className = "pp-note-t"; tx.textContent = n.text; row.appendChild(tx);
    box.appendChild(row);
  }
}
// ── project page helpers: what is shown, what is ticked, the pay draft ──
function _ppBillShown(b) { return _pp.filter === "all" || !b.paid; }
function _ppBillsOf(dr) { return [...(dr.bills || []), ...(dr.sub_bills || [])]; }
function _ppAllRows(draws) {   // every bill on the given draws as {dr, b, isSub}
  const out = [];
  for (const dr of draws) { for (const b of (dr.bills || [])) out.push({ dr, b, isSub: false }); for (const b of (dr.sub_bills || [])) out.push({ dr, b, isSub: true }); }
  return out;
}
function _ppTitle(dr) {   // RP = scopes, not draws (owner 2026-09-16); a scope is named by its invoice
  if (dr.scope) { if (dr.no_draw) return "Not yet invoiced"; const nos = dr.invoice_nos || (dr.invoice_no ? [dr.invoice_no] : []); return nos.length ? "Scope · Inv " + nos.join(", ") : "Scope"; }
  return drawTitle(dr);
}
function _ppSpan(dr) {   // the window a scope's costs come from ("through 05/12/2026", "05/13/2026 – 06/12/2026", "after 06/12/2026"); draws keep their memo period
  if (!dr.scope) return drawSpan(dr);
  const a = dr.period_start, b = dr.period_end;
  if (!a && !b) return "";
  if (a === "1900-01-01") return b && b !== "9999-12-31" ? "through " + fmtDate(b) : "";
  if (b === "9999-12-31") return "after " + fmtDate(new Date(new Date(a).getTime() - 86400000).toISOString().slice(0, 10));
  return fmtDate(a) + " – " + fmtDate(b);
}
function _ppStage(dr) { return dr.scope && dr.no_draw ? "Not yet invoiced" : (dr.stage || ""); }
function _ppAllLabel() { return _pp.isRp ? "All scopes" : "All draws"; }
function _ppUnit(n) { return _pp.isRp ? (n === 1 ? "scope" : "scopes") : (n === 1 ? "draw" : "draws"); }
function _ppInvLabel(dr) { const nos = dr.invoice_nos || (dr.invoice_no ? [dr.invoice_no] : []); return !nos.length ? "" : nos.length === 1 ? "Inv " + nos[0] : `${nos.length} invoices`; }
function _ppPayable(b) { return !b.paid && b.gates && b.bill_id; }
function _ppSel(b) { return _pp.payDraft.has(b.bill_id) ? _pp.payDraft.get(b.bill_id) : !!b.pay_selected; }   // effective tick = the draft over the saved run
function _ppAllSelected() { const out = []; for (const dr of _pp.d.draws) for (const b of _ppBillsOf(dr)) if (b.bill_id && _ppSel(b)) out.push({ dr, b }); return out; }
function _ppDraftSet(bills, selected) {   // tick / untick into the DRAFT - nothing reaches the server until Save (owner 2026-09-15)
  let n = 0;
  for (const b of bills) { if (!b.bill_id) continue; n++; if (!!b.pay_selected === !!selected) _pp.payDraft.delete(b.bill_id); else _pp.payDraft.set(b.bill_id, !!selected); }
  if (!n) { toast("These bills have no QuickBooks link to key the pay run on"); return; }
  const y = window.scrollY; _renderPpDraws(); window.scrollTo(0, y);
}
function _ppCheck(bills, label, title) {   // a select-all checkbox for a group of bills (unpaid, payable by us) - pay mode only
  const payable = bills.filter(_ppPayable);
  const cb = document.createElement("input"); cb.type = "checkbox"; cb.className = "pp-grp-cb";
  cb.checked = payable.length > 0 && payable.every(_ppSel);
  cb.indeterminate = !cb.checked && payable.some(_ppSel);
  cb.disabled = !payable.length; cb.title = title || `Tick every unpaid bill in ${label}`;
  cb.onclick = (e) => e.stopPropagation();
  cb.onchange = () => _ppDraftSet(payable, cb.checked);
  return cb;
}
function _ppLeaveBlocked() {   // an unsaved pay run never walks away with you (owner 2026-09-15: "don't let leave without saving or discarding")
  if (!(typeof _pp !== "undefined" && _pp && _pp.payDraft && _pp.payDraft.size)) return false;
  const rv = $("#recordView"); if (!rv || rv.hidden || !$("#ppDraws")) return false;
  const n = _pp.payDraft.size;
  toast(`${n} unsaved pay tick${n === 1 ? "" : "s"} on ${_pp.pn} - Save or Discard first`, 3200);
  const bar = $("#ppPayBar"); if (bar) { bar.classList.add("flash"); setTimeout(() => bar.classList.remove("flash"), 1000); bar.scrollIntoView({ block: "center", behavior: "smooth" }); }
  return true;
}
function _ppTogglePay() {
  if (_pp.payMode && _pp.payDraft.size) { _ppLeaveBlocked(); return; }
  _pp.payMode = !_pp.payMode;
  if (!_pp.payMode) _pp.payDraft = new Map();
  _renderPpDraws();
}
async function _ppSavePay() {
  const items = [..._pp.payDraft].map(([bill_id, selected]) => ({ bill_id, selected, amount: null }));
  if (!items.length) return;
  const btn = $("#ppPaySave"); if (btn) { btn.disabled = true; btn.textContent = "Saving…"; }
  let r;
  try { r = await (await fetch("/api/pay-run", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ items }) })).json(); }
  catch (e) { r = { error: String(e) }; }
  if (!(r && r.ok)) { if (btn) { btn.disabled = false; btn.textContent = "Save"; } toast("NOT saved - " + ((r && r.error) || "could not reach the server"), 5000); return; }
  for (const dr of _pp.d.draws) for (const b of _ppBillsOf(dr)) if (b.bill_id && _pp.payDraft.has(b.bill_id)) b.pay_selected = _pp.payDraft.get(b.bill_id);
  _pp.payDraft = new Map();
  toast(`Saved ✓ pay run · ${items.length} change${items.length === 1 ? "" : "s"} on ${_pp.pn}`, 3000);
  const y = window.scrollY; _renderPpDraws(); window.scrollTo(0, y);
}
function _ppDiscardPay() { const n = _pp.payDraft.size; _pp.payDraft = new Map(); const y = window.scrollY; _renderPpDraws(); window.scrollTo(0, y); if (n) toast("Discarded unsaved pay ticks"); }
function _ppPaintTools() {
  const d = _pp.d, all = _ppAllRows(d.draws);
  const nAll = all.length, nUnpaid = all.filter(x => !x.b.paid).length;
  const fseg = $("#ppFilterSeg");
  if (fseg) { fseg.innerHTML = "";
    for (const [k, lbl] of [["all", `All bills · ${nAll}`], ["unpaid", `Unpaid · ${nUnpaid}`]]) { const b = document.createElement("button"); b.type = "button"; b.className = "seg-btn" + (_pp.filter === k ? " on" : ""); b.textContent = lbl;
      b.title = k === "all" ? "Every bill on the job, paid or not, with the draw it sits under" : "Only what is still owed to vendors"; b.onclick = () => { _pp.filter = k; _renderPpDraws(); }; fseg.appendChild(b); } }
  const sseg = $("#ppSortSeg");
  if (sseg) { sseg.innerHTML = "";
    for (const [k, l, t] of [["vendor", "Name", "Vendors A to Z"], ["amount", "Total amount", "Vendors by what they billed on this job, largest first"], ["code", "Cost code", "Bands by cost code"], ["date", "Date", "Every bill in date order, no bands"]]) { const b = document.createElement("button"); b.type = "button"; b.className = "seg-btn" + (_pp.sort === k ? " on" : ""); b.textContent = l; b.title = t;
      b.onclick = () => { _pp.sort = k; if (k === "date") _pp.openV = new Set(["*"]); _renderPpDraws(); }; sseg.appendChild(b); } }
  const exp = $("#ppExpand"); if (exp) { exp.textContent = _pp.openV.has("*") ? "Collapse all" : "Expand all"; exp.hidden = _pp.sort === "date"; }
  const pb = $("#ppPayBtn");
  if (pb) { const n = _ppAllSelected().length; pb.classList.toggle("on", _pp.payMode);
    pb.textContent = _pp.payMode ? "Paying bills · Done" : (n ? `Pay bills · ${n} on the run` : "Pay bills");
    pb.title = _pp.payMode ? "Leave the pay-run controls (Save or Discard first)" : "Tick the bills to pay on this job - the pay-run controls appear"; }
}
function _ppPaintPayBar() {
  const bar = $("#ppPayBar"); if (!bar) return;
  bar.hidden = !_pp.payMode; bar.innerHTML = ""; if (!_pp.payMode) return;
  const picked = _ppAllSelected(), tot = picked.reduce((s, x) => s + num(x.b.open), 0), dirty = _pp.payDraft.size;
  bar.classList.toggle("dirty", dirty > 0);
  const st = document.createElement("span"); st.className = "pp-paystat";
  st.innerHTML = `<b>${picked.length}</b> bill${picked.length === 1 ? "" : "s"} ticked to pay · <b>${_ge(money(tot))}</b>` + (dirty ? ` · <span class="pp-dirty">${dirty} unsaved change${dirty === 1 ? "" : "s"}</span>` : ` · <span class="dim">saved</span>`);
  bar.appendChild(st);
  const mk = (lbl, cls, fn, title, id) => { const b = document.createElement("button"); b.type = "button"; b.className = "btn small " + (cls || ""); b.textContent = lbl; if (title) b.title = title; if (id) b.id = id; b.onclick = fn; bar.appendChild(b); return b; };
  mk("Mark blockers to pay", "", _ppMarkBlockers, "Tick every unpaid bill on the draws before the next one the GC owes - their waivers unlock it");
  mk("Export pay list", "", _ppExport, "Excel report of the bills ticked to pay on this job, grouped by draw");
  mk("Open Pay run", "subtle", () => { if (!_ppLeaveBlocked()) setTab("paybills"); }, "The check-run worksheet across every job");
  const sv = mk("Save", "primary", _ppSavePay, "Write the ticks to the pay run (local intent only - never QuickBooks)", "ppPaySave"); sv.disabled = !dirty;
  const dc = mk("Discard", "", _ppDiscardPay, "Drop the unsaved ticks"); dc.hidden = !dirty;
  if (picked.length) {   // what is ticked, by vendor - at a glance before any export
    const ul = document.createElement("div"); ul.className = "pp-sel-list";
    const byV = new Map(); for (const x of picked) { const k = x.b.vendor || "?"; if (!byV.has(k)) byV.set(k, []); byV.get(k).push(x); }
    for (const [v, xs] of [...byV].sort((a, b) => a[0].localeCompare(b[0]))) {
      const li = document.createElement("div"); li.className = "pp-sel-v";
      li.innerHTML = `<b>${_ge(v)}</b> · ${xs.length} bill${xs.length === 1 ? "" : "s"} · ${_ge(money(xs.reduce((s, x) => s + num(x.b.open), 0)))} <span class="dim">(${xs.map(x => _ge(x.b.bill_ref || "?") + " on " + _ge(x.dr.no_draw ? "no draw" : _ppTitle(x.dr))).join(", ")})</span>`;
      ul.appendChild(li);
    }
    bar.appendChild(ul);
  }
}
// The draw boxes + the Coverage table sit on top whatever is open; a click on a box or a row opens that draw's
// details underneath and the boxes stay (owner 2026-09-15). "All draws" = every bill on the job with the draw it
// sits under. Left / right arrow keys flip between them.
function _ppTabs(d, nxInv) {
  const strip = document.createElement("div"); strip.className = "pp-tabs";
  const mk = (key, lbl, sub, title) => { const b = document.createElement("button"); b.type = "button"; b.className = "pp-tab" + (_pp.view === key ? " on" : "");
    b.innerHTML = `<span>${_ge(lbl)}</span>${sub ? `<small>${_ge(sub)}</small>` : ""}`; b.title = title || ""; b.onclick = () => { _pp.view = key; _renderPpDraws(); _ppScrollDetail(); }; return b; };
  const n = d.draws.filter(x => !x.no_draw).length;
  strip.appendChild(mk("all", _ppAllLabel(), `${n} ${_ppUnit(n)} · ${_ppAllRows(d.draws).length} bills`, `Every bill on the job, with the ${_pp.isRp ? "scope" : "draw"} it sits under`));
  for (const dr of d.draws) {
    const span = _ppSpan(dr);
    const b = mk(dr.matched_invoice, _ppTitle(dr), dr.no_draw ? `${_ppBillsOf(dr).length} bills` : (dr.scope ? (span || _ppInvLabel(dr)) : _ppInvLabel(dr)), [span ? "Covers " + span : "", _ppStage(dr)].filter(Boolean).join(" · "));
    if (dr.invoice_no && dr.invoice_no === nxInv) b.classList.add("next");
    strip.appendChild(b);
  }
  const hint = document.createElement("span"); hint.className = "pp-tabs-hint dim"; hint.textContent = `← → flips ${_pp.isRp ? "scopes" : "draws"}`; strip.appendChild(hint);
  return strip;
}
function _ppScrollDetail() { /* the page stays put when a draw is picked (owner 2026-09-23: "when i click a draw do not make my screen go down") */ }
function _ppCoverage(d, nxInv) {
  const wrap = document.createElement("div"); wrap.className = "table-scroll pp-cov-wrap";
  const t = document.createElement("table"); t.className = "grid pp-cov" + (_pp.view === "all" ? " all-sel" : "");   // All draws = the whole table lit; one draw = its row (owner 2026-09-23)
  t.innerHTML = `<thead><tr><th class='left'>${_pp.isRp ? "Costs dated" : "Period covered"}</th><th class='left'>${_pp.isRp ? "Scope" : "Draw"}</th><th class='left'>Invoice</th><th class='left'>Date</th><th class='right' title='Billed incl. retainage - the project P&amp;L basis'>Billed (gross)</th><th class='left'>GC</th><th class='right'>Costs</th><th class='right'>Gross</th><th class='right'>Margin</th><th class='right'>Overhead</th><th class='right'>Net</th><th class='left'>Stage</th></tr></thead>`;
  const tb = document.createElement("tbody");
  const tot = { income: 0, costs: 0, gross: 0, overhead: 0, net: 0, unbilled: 0 };
  const rt = (v, cls) => { const td = document.createElement("td"); td.className = "right" + (cls ? " " + cls : ""); td.textContent = v; return td; };
  for (const dr of d.draws) {
    const p = dr.pl || {}; const costs = dr.no_draw ? num(dr.gate_amt) + num(dr.subs_amt) : num(p.costs);
    const tr = document.createElement("tr"); tr.className = "pp-cov-row" + (_pp.view === dr.matched_invoice ? " sel" : "");   // ONE highlight: the draw you picked (owner 2026-09-23 - the "next draw" tint read as a second selection) tr.title = "Open this draw underneath";
    tr.onclick = () => { _pp.view = dr.matched_invoice; _renderPpDraws(); _ppScrollDetail(); };
    { const td = leftText(_ppSpan(dr) || "–"); if (!_ppSpan(dr)) td.classList.add("dim"); tr.appendChild(td); }
    tr.appendChild(leftText(_ppTitle(dr)));
    { const nos = dr.invoice_nos || (dr.invoice_no ? [dr.invoice_no] : []); const td = leftText(dr.no_draw ? "–" : (nos.join(", ") || "–")); if (nos.length > 1) td.title = `${nos.length} invoices dated the same day - one draw`; tr.appendChild(td); }
    tr.appendChild(leftText(dr.ar_date ? fmtDateShort(dr.ar_date) : "–"));
    tr.appendChild(rt(dr.no_draw ? "–" : money(p.income)));
    { const td = document.createElement("td"); td.className = "left"; const sp = document.createElement("span"); sp.className = dr.no_draw ? "dim" : dr.gc_paid ? "ar-paid" : "ar-open"; sp.textContent = dr.no_draw ? "–" : dr.gc_paid ? "paid" : "owes " + money(dr.ar_open); td.appendChild(sp); tr.appendChild(td); }
    tr.appendChild(rt(money(costs)));
    if (dr.no_draw) { for (let i = 0; i < 4; i++) tr.appendChild(rt("–", "dim")); tot.unbilled += costs; }
    else {
      tr.appendChild(rt(money(p.gross), num(p.gross) < 0 ? "neg" : "")); tr.appendChild(rt(p.margin_pct != null ? (p.margin_pct * 100).toFixed(1) + "%" : "–", num(p.gross) < 0 ? "neg" : ""));
      tr.appendChild(rt(money(p.overhead))); tr.appendChild(rt(money(p.net), num(p.net) < 0 ? "neg" : "pos"));
      tot.income += num(p.income); tot.costs += costs; tot.gross += num(p.gross); tot.overhead += num(p.overhead); tot.net += num(p.net);
    }
    { const td = leftText(_ppStage(dr) || "–"); td.classList.add("dim"); tr.appendChild(td); }
    tb.appendChild(tr);
  }
  const tr = document.createElement("tr"); tr.className = "pp-cov-total pp-cov-row" + (_pp.view === "all" ? " sel" : ""); tr.title = "Every bill on the job underneath";
  tr.onclick = () => { _pp.view = "all"; _renderPpDraws(); _ppScrollDetail(); };
  tr.appendChild(leftText("")); tr.appendChild(leftText(_ppAllLabel())); tr.appendChild(leftText("")); tr.appendChild(leftText(""));
  tr.appendChild(rt(money(tot.income))); tr.appendChild(leftText(""));
  tr.appendChild(rt(money(tot.costs + tot.unbilled)));
  tr.appendChild(rt(money(tot.gross), tot.gross < 0 ? "neg" : "")); tr.appendChild(rt(tot.income ? (tot.gross / tot.income * 100).toFixed(1) + "%" : "–", tot.gross < 0 ? "neg" : ""));
  tr.appendChild(rt(money(tot.overhead))); tr.appendChild(rt(money(tot.net), tot.net < 0 ? "neg" : "pos"));
  { const td = leftText(tot.unbilled > 0.005 ? `${money(tot.unbilled)} of costs not ${_pp.isRp ? "invoiced" : "on a draw"} yet` : ""); td.classList.add("dim"); tr.appendChild(td); }
  tb.appendChild(tr); t.appendChild(tb); wrap.appendChild(t);
  const cap = document.createElement("div"); cap.className = "bills-cap"; cap.textContent = _pp.isRp
    ? "An RP job has no draws: each invoice is a scope · Costs dated = the bills (materials we pay + subs) dated after the previous invoice up to this one · Billed (gross) = the invoice incl. retainage, the project P&L's basis · Overhead = the scope's share · click a row to open that scope underneath."
    : "Period covered = the billing window stated on the invoice · Draw = the month (MFD) or number (CP) on the invoice; invoices of the same draw are one draw · Billed (gross) = the invoices incl. retainage, the project P&L's basis · Costs = materials we pay + labor dated in the draw period · Overhead = the draw's share · click a row to open that draw underneath.";
  wrap.appendChild(cap);
  return wrap;
}
if (!window.__ppKeys) { window.__ppKeys = true;   // ← → flip between All draws and each draw while a project page is open
  document.addEventListener("keydown", (e) => {
    if (e.key !== "ArrowLeft" && e.key !== "ArrowRight") return;
    if (typeof _pp === "undefined" || !_pp || !_pp.d) return;
    const rv = $("#recordView"); if (!rv || rv.hidden || !$("#ppDraws")) return;
    const t = e.target; if (t && (t.tagName === "INPUT" || t.tagName === "SELECT" || t.tagName === "TEXTAREA" || t.isContentEditable)) return;
    const keys = ["all", ..._pp.d.draws.map(x => x.matched_invoice)];
    let i = keys.indexOf(_pp.view || "all"); if (i < 0) i = 0;
    i = e.key === "ArrowRight" ? Math.min(keys.length - 1, i + 1) : Math.max(0, i - 1);
    _pp.view = keys[i]; _renderPpDraws(); e.preventDefault();
  });
}
function _renderPpDraws() {
  const y0 = window.scrollY;   // BEFORE emptying: an empty host shortens the page and the browser pulls the scroll up (owner 2026-09-23: "clicking on draws moves page up")
  const d = _pp.d, host = $("#ppDraws"); if (!host) return;
  host.style.minHeight = host.offsetHeight + "px";   // hold the height while it redraws
  host.innerHTML = "";
  _ppPaintPayBar();
  if (!d.draws.length) { const p = document.createElement("div"); p.className = "bills-cap"; p.textContent = "No invoices or bills on this job in the ledger yet."; return host.appendChild(p); }
  const nxInv = d.funding && d.funding.next_draw ? d.funding.next_draw.invoice_no : null;
  const keys = ["all", ...d.draws.map(x => x.matched_invoice)];
  if (!keys.includes(_pp.view)) _pp.view = "all";
  _recSave({ k: "project", id: _pp.pn, view: _pp.view });   // the open draw survives a refresh (owner 2026-09-23)
  const cur = _pp.view === "all" ? null : d.draws.find(x => x.matched_invoice === _pp.view);
  host.appendChild(_ppTabs(d, nxInv));
  host.appendChild(_ppCoverage(d, nxInv));
  const det = document.createElement("div"); det.className = "pp-detail"; det.id = "ppDetail";
  // the filter / sort / expand bar sits ON the bill list it changes (owner 2026-09-23: "you put the filters all the way
  // over here yet it's changing the data all the way below") - not above the draw table
  if (cur && !cur.no_draw) det.appendChild(_ppDrawIncome(cur));   // the draw's income, split like the P&L, above its bills (owner 2026-09-23)
  if (_pp.toolsEl) det.appendChild(_pp.toolsEl);
  // No head and no equation under the table any more (owner 2026-09-23: "remove this, we already have this info on top
  // in the table") - the Coverage table IS the P&L per draw and lights what is open; only a draw's push notes stay.
  if (cur) for (const [k, word] of [["pushed_in", "moved into this draw from the one before"], ["pushed_out", "moved on to the next draw"]]) { const p = cur[k]; if (!p) continue;
    const cap = document.createElement("div"); cap.className = "bills-cap pp-push"; cap.textContent = `${p.count} bill${p.count === 1 ? "" : "s"} · ${money(p.total)} ${word}: ${p.note || ""}`; det.appendChild(cap); }
  det.appendChild(_ppBillsTable(cur));
  host.appendChild(det);
  _ppPaintTools();
  _ppRenderNotes();
  host.style.minHeight = "";
  window.scrollTo(0, y0);
}
// One draw's P&L, laid out like the job P&L (owner 2026-09-23: "when i click a draw i need the income above the bills
// sections, split any retainage, just like the pnl"): gross billed - retained = net billed; gross billed - costs =
// gross profit; - overhead = net profit. Gross billed = the invoice lines before retainage (server _invoices_split).
function _ppDrawIncome(dr) {
  const p = dr.pl || {}, wrap = document.createElement("div"); wrap.className = "pp-drawpl";
  const t = document.createElement("table"); t.className = "pp-pva pp-drawpl-t";
  const tb = document.createElement("tbody");
  const r = (label, v, opts = {}) => {
    const tr = document.createElement("tr"); if (opts.cls) tr.className = opts.cls;
    const l = document.createElement("td"); l.className = "left pp-pva-lab"; l.textContent = label; if (opts.tip) l.title = opts.tip; tr.appendChild(l);
    const a = document.createElement("td"); a.className = "right pp-pva-act" + (opts.neg ? " neg" : ""); a.textContent = v; tr.appendChild(a);
    const n = document.createElement("td"); n.className = "left pp-pva-diff"; n.textContent = opts.note || ""; tr.appendChild(n);
    tb.appendChild(tr);
  };
  const gross = num(p.income), ret = num(p.retainage), netB = num(p.net_billed), costs = num(p.costs), gp = num(p.gross), oh = num(p.overhead), net = num(p.net);
  const pct = (v) => gross ? ` (${(v / gross * 100).toFixed(1)}%)` : "";
  r("Gross billed", money(gross), { tip: "the draw's invoices before retainage" });
  r("Retained", ret ? "(" + money(ret) + ")" : money(0), { note: ret ? "held by the GC until release" : "" });
  r("Net billed", money(netB), { note: dr.gc_paid ? "GC paid" : (num(dr.ar_open) > 0.005 ? `GC owes ${money(dr.ar_open)}` : "") });
  r("Costs", "(" + money(costs) + ")", { note: `materials ${money(p.materials)} · labor ${money(p.labor)}` });
  r("Gross profit", money(gp) + pct(gp), { neg: gp < 0 });
  r(`Overhead · ${p.overhead_basis || ""}`, "(" + money(oh) + ")");
  r("Net profit", money(net) + pct(net), { cls: "pp-pva-total", neg: net < 0, tip: "gross billed - costs - overhead" });
  t.appendChild(tb); wrap.appendChild(t);
  return wrap;
}
function _ppBillsTable(cur) {
  const d = _pp.d, rows = _ppAllRows(cur ? [cur] : d.draws).filter(x => _ppBillShown(x.b));
  const box = document.createElement("div"); box.className = "pp-billsbox";
  if (!rows.length) { const p = document.createElement("div"); p.className = "bills-cap"; p.textContent = _pp.filter === "unpaid" ? "Every bill here is paid." : "No bills."; box.appendChild(p); return box; }
  const pay = _pp.payMode, showDraw = !cur;
  const cols = [];
  if (pay) cols.push(["Pay", "left"]);
  cols.push(["Vendor / Bill #", "left"], ["Code", "left"], ["Amount", "right"], ["Date", "left"]);
  if (showDraw) cols.push(["Draw", "left"]);
  cols.push(["Where / status", "left"], ["Paid?", "left"], ["Description", "left"]);
  if (pay) cols.push(["Waiver", "left"]);
  const COLS = cols.length;
  const codeOf = b => (b.codes && b.codes.length) ? b.codes.join(", ") : "(uncoded)";
  const table = document.createElement("table"); table.className = "grid pp-bills"; const thead = document.createElement("thead"), tbody = document.createElement("tbody");
  thead.innerHTML = "<tr>" + cols.map(([c, al]) => `<th class="${al}">${_ge(c)}</th>`).join("") + "</tr>";
  const billRow = (x, inGrp) => {   // inGrp = under a vendor / code band: indented, and the vendor name is not repeated (owner 2026-09-08)
    const { dr, b, isSub } = x;
    const tr = document.createElement("tr"); if (b.paid) tr.classList.add("inv-paid"); if (inGrp) tr.classList.add("pp-in-grp"); if (pay && _ppPayable(b) && _ppSel(b)) tr.classList.add("pp-ticked");
    if (pay) { const pc = document.createElement("td"); pc.className = "left";
      if (_ppPayable(b)) { const cb = document.createElement("input"); cb.type = "checkbox"; cb.checked = _ppSel(b); cb.title = "Put this bill on the pay run (Pay Bills) - local intent, never QuickBooks";
        cb.onclick = (e) => e.stopPropagation(); cb.onchange = () => _ppDraftSet([b], cb.checked); pc.appendChild(cb); }
      else if (!b.gates) { const s = document.createElement("span"); s.className = "vg-tag"; s.textContent = "GC pays"; s.title = "Concrete pumping - paid by the GC directly"; pc.appendChild(s); }
      tr.appendChild(pc); }
    { const vb = document.createElement("td"); vb.className = "left pp-vb"; if (!(inGrp && _pp.sort !== "code")) { const v = document.createElement("span"); v.className = "pp-vend"; v.textContent = b.vendor || "–"; vb.appendChild(v); vb.appendChild(document.createTextNode(" ")); }
      const link = qboLinkCell(b.bill_ref || "–", isSub ? qboUrl(b.txn_type === "Expense" ? "expense" : "bill", b.bill_id) : qboBillHref(b.qbo_link), "Open this bill in QuickBooks");
      while (link.firstChild) vb.appendChild(link.firstChild);
      if (b.bill_id) {   // the viewer flips through this vendor's bills in the view you are looking at (owner 2026-09-16)
        const same = rows.filter(y => (y.b.vendor || "") === (b.vendor || "") && y.b.bill_id).map(y => ({ type: y.isSub && y.b.txn_type === "Expense" ? "Purchase" : "Bill", id: String(y.b.bill_id), n: y.b.att || 0, title: `${y.b.vendor || ""} · bill ${y.b.bill_ref || ""}` }));
        const ab = attBtn(isSub && b.txn_type === "Expense" ? "Purchase" : "Bill", b.bill_id, b.att, `${b.vendor || ""} · bill ${b.bill_ref || ""}`, { items: same, index: Math.max(0, same.findIndex(y => y.id === String(b.bill_id))) });
        ab.style.marginLeft = "6px"; vb.appendChild(ab); }
      if (!pay && !b.gates) { const s = document.createElement("span"); s.className = "vg-tag"; s.textContent = "GC pays"; s.title = "Concrete pumping - paid by the GC directly"; s.style.marginLeft = "6px"; vb.appendChild(s); }
      tr.appendChild(vb); }
    { const cc = document.createElement("td"); cc.className = "left"; for (const c of (b.codes || [])) { const chip = document.createElement("span"); chip.className = "codechip"; chip.textContent = c; cc.appendChild(chip); cc.appendChild(document.createTextNode(" ")); } if (!(b.codes || []).length) { cc.textContent = "–"; cc.classList.add("dim"); } tr.appendChild(cc); }
    const ac = document.createElement("td"); ac.className = "ip-amt"; ac.appendChild(moneyCell(b.amount)); tr.appendChild(ac);
    tr.appendChild(leftText(fmtDateShort(b.bill_date)));
    if (showDraw) { const dc = document.createElement("td"); dc.className = "left pp-draw-col"; const a = document.createElement("a"); a.href = "#"; a.textContent = dr.no_draw ? "No draw yet" : _ppTitle(dr); a.title = dr.no_draw ? "Not matched to a draw in the Bill Tracker yet" : `${_ppInvLabel(dr)}${drawSpan(dr) ? " · covers " + drawSpan(dr) : ""} - open this draw`;
      a.onclick = (e) => { e.preventDefault(); _pp.view = dr.matched_invoice; _renderPpDraws(); _ppScrollDetail(); }; dc.appendChild(a);
      if (!dr.no_draw && dr.invoice_no) { const s = document.createElement("small"); s.className = "dim"; s.textContent = " " + _ppInvLabel(dr); dc.appendChild(s); } tr.appendChild(dc); }
    { const wc = leftText(isSub ? "QuickBooks · labor" : [b.invoice_status, b.approved ? (b.approved === "approved" ? "approved" : b.approved) : null].filter(Boolean).join(" · ") || "Bill Tracker"); wc.classList.add("dim");
      if (b.pushed) { const s = document.createElement("span"); s.className = "vg-tag push"; s.textContent = b.pushed; s.title = b.pushed_note || "carried into this draw by agreement with the supplier"; wc.prepend(document.createTextNode(" ")); wc.prepend(s); }
      tr.appendChild(wc); }
    { const st = document.createElement("td"); st.className = "left"; const pill = document.createElement("span"); pill.className = b.paid ? "ar-paid" : "ar-open";
      pill.textContent = isSub ? (b.paid ? "Paid " + fmtDateShort(b.pay_date) : "No") : (b.pay_date ? "Paid " + fmtDateShort(b.pay_date) : (b.paid ? "Yes (no date)" : "No" + (num(b.open) > 0.005 ? " · " + money(b.open) + " open" : "")));
      if (isSub && !b.paid) pill.title = "No QuickBooks bill payment applied to this bill in the loaded window (this year)"; st.appendChild(pill); tr.appendChild(st); }
    { const dc = leftText(b.description || "–"); dc.className += " inv-memo"; dc.title = b.description || ""; if (!b.description) dc.classList.add("dim"); tr.appendChild(dc); }
    if (pay) { const wc = document.createElement("td"); wc.className = "left";
      if (dr.no_draw || dr.scope || isSub) wc.textContent = "–";
      else { const cb = document.createElement("input"); cb.type = "checkbox"; cb.checked = !!b.waiver; cb.title = "Tick when the vendor's unconditional waiver is in hand";
        cb.onclick = (e) => e.stopPropagation(); cb.onchange = () => { setWaiver(dr, b, cb); b.waiver = cb.checked; }; wc.appendChild(cb); if (!b.waiver && !b.paid) { const s = document.createElement("span"); s.className = "inv-late"; s.textContent = " needed"; wc.appendChild(s); } }
      tr.appendChild(wc); }
    return tr;
  };
  const iAmt = (pay ? 1 : 0) + 2;   // the Amount column's index - the totals line up under it
  const sectionRow = (label, xs, sect) => {   // a plain section label (owner 2026-09-16: no highlighted rows); the totals sit at the bottom
    const bills = xs.map(x => x.b);
    const sr = document.createElement("tr"); sr.className = "pp-sect";
    if (pay) { const cbTd = document.createElement("td"); cbTd.className = "left"; cbTd.appendChild(_ppCheck(bills, label, `Tick every unpaid ${label.toLowerCase()} bill shown`)); sr.appendChild(cbTd); }
    const td = document.createElement("td"); td.className = "left"; td.colSpan = COLS - (pay ? 1 : 0);
    // click the section to open every vendor in it at once (owner 2026-09-23: "give me ability to uncollapse the material and subs so i can see all")
    const dk = cur ? cur.matched_invoice : "all", kf = _pp.sort === "code" ? (x => codeOf(x.b)) : (x => x.b.vendor || "?");
    const keys = [...new Set(xs.map(kf))].map(v => `${dk}|${sect}|${_pp.sort}|${v}`);
    const allOpen = !flat && (_pp.openV.has("*") || keys.every(k => _pp.openV.has(k)));
    td.innerHTML = `<span class="pp-sect-lab">${flat ? "" : `<span class="bg-caret">${allOpen ? "▾" : "▸"} </span>`}${_ge(label)}</span>`
      + (flat ? "" : ` <span class="pp-sect-all">${allOpen ? "hide the bills" : "show every bill"}</span>`)
      + (sect === "labor" ? ` <span class="dim">QuickBooks bills · paid = a bill payment applied this year</span>` : ` <span class="dim">Bill Tracker</span>`);
    if (!flat) { sr.style.cursor = "pointer"; sr.title = allOpen ? `Fold every ${label.toLowerCase()} vendor` : `Open every ${label.toLowerCase()} vendor`;
      sr.onclick = (e) => { if (e.target.closest("input")) return;
        if (_pp.openV.has("*")) { _pp.openV = new Set(); for (const x2 of _ppAllRows(cur ? [cur] : d.draws)) _pp.openV.add(`${dk}|${x2.isSub ? "labor" : "materials"}|${_pp.sort}|${kf(x2)}`); }
        for (const k of keys) { if (allOpen) _pp.openV.delete(k); else _pp.openV.add(k); }
        const y = window.scrollY; _renderPpDraws(); window.scrollTo(0, y); }; }
    sr.appendChild(td); tbody.appendChild(sr);
  };
  const totalRow = (label, xs, grand) => {   // Materials total / Labor total / Total - amount under Amount, what is still to pay beside it, like a real sheet
    const bills = xs.map(x => x.b), paidCt = bills.filter(b => b.paid).length;
    const tot = bills.reduce((s, b) => s + num(b.amount), 0), owed = bills.reduce((s, b) => s + (b.paid ? 0 : num(b.open)), 0);
    const tr = document.createElement("tr"); tr.className = "pp-tot" + (grand ? " pp-grand" : "");
    const l = document.createElement("td"); l.className = "left"; l.colSpan = iAmt; l.textContent = label; tr.appendChild(l);
    const a = document.createElement("td"); a.className = "ip-amt pp-tot-amt"; a.textContent = money(tot); tr.appendChild(a);
    const r = document.createElement("td"); r.className = "left"; r.colSpan = COLS - iAmt - 1;
    r.innerHTML = `<span class="${paidCt === bills.length ? "ip-paid ok" : "dim"}">${paidCt}/${bills.length} paid</span>` + (owed > 0.005 ? ` · <span class="pp-topay">${_ge(money(owed))} to pay</span>` : "");
    tr.appendChild(r); tbody.appendChild(tr);
  };
  const groupedRows = (xs, sect) => {   // bands by vendor (A to Z or by total) or by cost code; collapsed until opened
    const keyOf = _pp.sort === "code" ? (x => codeOf(x.b)) : (x => x.b.vendor || "?");
    const byK = new Map(); for (const x of xs) { const k = keyOf(x); if (!byK.has(k)) byK.set(k, []); byK.get(k).push(x); }
    const bands = [...byK]; const totOf = list => list.reduce((s, x) => s + num(x.b.amount), 0);
    bands.sort(_pp.sort === "amount" ? ((a, b) => totOf(b[1]) - totOf(a[1]) || a[0].localeCompare(b[0])) : ((a, b) => a[0].localeCompare(b[0], undefined, { numeric: true })));
    for (const [v, list] of bands) {
      const vkey = `${cur ? cur.matched_invoice : "all"}|${sect}|${_pp.sort}|${v}`, vopen = _pp.openV.has("*") || _pp.openV.has(vkey);
      const bills = list.map(x => x.b), paidCt = bills.filter(b => b.paid).length, tot = totOf(list), owed = bills.reduce((s, b) => s + (b.paid ? 0 : num(b.open)), 0);
      const gtr = document.createElement("tr"); gtr.className = "bill-subgroup pp-vendor"; gtr.style.cursor = "pointer"; gtr.title = vopen ? "Click to collapse" : "Click to see the bills";
      if (pay) { const cbTd = document.createElement("td"); cbTd.className = "left"; cbTd.appendChild(_ppCheck(bills, v)); gtr.appendChild(cbTd); }
      const vt = document.createElement("td"); vt.className = "left"; vt.colSpan = COLS - (pay ? 1 : 0);
      const cell = document.createElement("div"); cell.className = "bg-cell"; const leftP = document.createElement("span"); leftP.className = "bg-left";
      const caret = document.createElement("span"); caret.className = "bg-caret"; caret.textContent = vopen ? "▾ " : "▸ "; const k = document.createElement("span"); k.className = "sg-key"; k.textContent = v; leftP.appendChild(caret); leftP.appendChild(k); cell.appendChild(leftP);
      const nDraws = new Set(list.map(x => x.dr.matched_invoice)).size;
      bandMetrics(cell, [[money(tot), "total"], [list.length, "bills"], showDraw ? [nDraws, nDraws === 1 ? "draw" : "draws"] : null, [`${paidCt}/${list.length}${paidCt === list.length ? " ✓" : ""}`, "paid", paidCt === list.length ? "ok" : "due"], [owed > 0.005 ? money(owed) : "–", "to pay", owed > 0.005 ? "neg" : ""]]);
      vt.appendChild(cell); gtr.appendChild(vt);
      gtr.onclick = (e) => { if (e.target.closest("input")) return;
        if (_pp.openV.has("*")) { _pp.openV = new Set(); for (const x2 of _ppAllRows(cur ? [cur] : d.draws)) _pp.openV.add(`${cur ? cur.matched_invoice : "all"}|${x2.isSub ? "labor" : "materials"}|${_pp.sort}|${keyOf(x2)}`); }
        if (_pp.openV.has(vkey)) _pp.openV.delete(vkey); else _pp.openV.add(vkey); const y = window.scrollY; _renderPpDraws(); window.scrollTo(0, y); };
      tbody.appendChild(gtr);
      if (vopen) for (const x of list.sort((p, q) => (p.b.bill_date || "").localeCompare(q.b.bill_date || ""))) tbody.appendChild(billRow(x, true));
    }
  };
  const flatRows = (xs) => { for (const x of [...xs].sort((p, q) => (p.b.bill_date || "").localeCompare(q.b.bill_date || ""))) tbody.appendChild(billRow(x, false)); };
  const mat = rows.filter(x => !x.isSub), subs = rows.filter(x => x.isSub);
  const flat = _pp.sort === "date";
  const gap = () => { const tr = document.createElement("tr"); tr.className = "pp-gap"; const td = document.createElement("td"); td.colSpan = COLS; tr.appendChild(td); tbody.appendChild(tr); };
  if (mat.length) { sectionRow("Materials", mat, "materials"); if (flat) flatRows(mat); else groupedRows(mat, "materials"); totalRow("Materials total", mat); gap(); }
  if (subs.length) { sectionRow("Labor (subs)", subs, "labor"); if (flat) flatRows(subs); else groupedRows(subs, "labor"); totalRow("Labor total", subs); gap(); }
  if (mat.length && subs.length) totalRow("Total", rows, true);
  thead.hidden = !flat && !_pp.openV.has("*") && ![..._pp.openV].some(k => k.startsWith(`${cur ? cur.matched_invoice : "all"}|`));   // headers only once bills show
  table.appendChild(thead); table.appendChild(tbody); box.appendChild(table);
  return box;
}
function _ppMarkBlockers() {
  const F = _pp.d.funding || {}, ids = new Set((F.blockers || []).map(b => b.bill_id).filter(Boolean));
  const bills = []; for (const dr of _pp.d.draws) for (const b of _ppBillsOf(dr)) if (ids.has(b.bill_id)) bills.push(b);
  if (!bills.length) { toast("No blockers to mark"); return; }
  _ppDraftSet(bills, true);
}
function _ppExport() {
  const d = _pp.d, rows = [];
  for (const dr of d.draws) for (const b of _ppBillsOf(dr)) if (b.bill_id && _ppSel(b)) rows.push([dr.no_draw ? (dr.scope ? "Not yet invoiced" : "No draw yet") : dr.scope ? _ppTitle(dr) : "Invoice " + (dr.invoice_no || "") + drawTag(dr), b.vendor, b.bill_ref, b.bill_date, num(b.amount), num(b.open), b.pay_date ? "Paid " + fmtDate(b.pay_date) : (b.pay_status || (b.paid ? "Paid" : "Open")), dr.sub_bills && dr.sub_bills.includes(b) ? "labor" : (b.waiver ? "received" : "needed")]);
  if (!rows.length) { toast("Nothing ticked to pay on this job yet - tick bills (or Mark blockers) first"); return; }
  const unsaved = _pp.payDraft.size ? `\n\n${_pp.payDraft.size} tick(s) are not saved yet - the export follows what is ticked on screen.` : "";
  if (!confirm(`Export ${rows.length} bill${rows.length === 1 ? "" : "s"} ticked to pay (${money(rows.reduce((s, r) => s + num(r[5]), 0))} open) as the pay-list report?${unsaved}`)) return;
  const nx = (d.funding || {}).next_draw, toPay = rows.reduce((s, r) => s + num(r[5]), 0);
  const footer = nx ? [{ label: `Unlocks ${nx.invoice_no || ""}${drawTag(nx)}`, value: num(nx.ar_open) },
                       { label: "Net = unlock - to pay", value: num(nx.ar_open) - toPay, cls: num(nx.ar_open) - toPay >= 0 ? "pos" : "neg" }] : [];
  const body = { name: `Pay list ${_pp.pn}`, sheet: "Pay list", title: `${_pp.pn} - bills to pay to unlock the next draw${nx ? " " + money(nx.ar_open) + " - Invoice " + (nx.invoice_no || "") + drawTag(nx) : ""}`, footer,
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
      const gtr = document.createElement("tr"); gtr.className = "bill-group"; gtr.style.cursor = "pointer"; gtr.title = open ? "Click to collapse" : "Click to expand";
      const gtd = document.createElement("td"); gtd.colSpan = 9;
      const cell = document.createElement("div"); cell.className = "bg-cell";
      const caret = document.createElement("span"); caret.className = "bg-caret"; caret.textContent = open ? "▾ " : "▸ ";
      const k = document.createElement("span"); k.className = "bg-key"; k.textContent = v.vendor;
      const leftV = document.createElement("span"); leftV.className = "bg-left"; leftV.appendChild(caret); leftV.appendChild(k); cell.appendChild(leftV);
      const allPaid = paidCt === v.bills.length;
      bandMetrics(cell, [[`${paidCt}/${v.bills.length}${allPaid ? " ✓" : ""}`, "paid", allPaid ? "ok" : "due"], [opn > 0.005 ? money(opn) : "–", "still owed", opn > 0.005 ? "neg" : ""], [money(tot), "total"]]);
      gtd.appendChild(cell); gtr.appendChild(gtd);
      gtr.onclick = () => { if (_ip.open.has(key)) _ip.open.delete(key); else _ip.open.add(key); _renderIpBills(); };
      tbody.appendChild(gtr);
      if (!open) continue;
      for (const b of v.bills) {
        const tr = document.createElement("tr");
        const nm = leftText(""); if (!b.gates) { const s = document.createElement("span"); s.className = "vg-tag"; s.textContent = "not paid by us"; s.title = "Concrete pumping - the GC pays this vendor directly"; nm.appendChild(s); } tr.appendChild(nm);
        tr.appendChild(leftText(fmtDateShort(b.bill_date)));
        { const lc = qboLinkCell(b.bill_ref || "–", qboBillHref(b.qbo_link), "Open this bill in QuickBooks"); if (b.att) { const ab = attBtn("Bill", b.bill_id, b.att, `${v.vendor} · bill ${b.bill_ref || ""}`); ab.style.marginLeft = "6px"; lc.appendChild(ab); } tr.appendChild(lc); }
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
        { const lc = qboLinkCell(l.doc_number || "–", qboUrl(l.txn_type === "Expense" ? "expense" : "bill", l.txn_id), "Open in QuickBooks"); if (l.att) { const ab = attBtn(l.txn_type === "Expense" ? "Purchase" : "Bill", l.txn_id, l.att, `${v.vendor} · bill ${l.doc_number || ""}`); ab.style.marginLeft = "6px"; lc.appendChild(ab); } tr.appendChild(lc); }
        const dc = leftText(l.description || "–"); dc.className += " inv-memo"; dc.title = l.description || ""; tr.appendChild(dc);
        const cc = document.createElement("td"); cc.className = "left"; if (l.cost_code) { const chip = document.createElement("span"); chip.className = "codechip"; chip.textContent = l.cost_code; cc.appendChild(chip); } tr.appendChild(cc);
        const ac = document.createElement("td"); ac.appendChild(moneyCell(l.amount)); tr.appendChild(ac); tbody.appendChild(tr); }
    }
    table.appendChild(thead); table.appendChild(tbody); scroll.appendChild(table); host.appendChild(scroll);
  }
}
async function openInvoicePage(inv) {
  if (_ppLeaveBlocked()) return;   // unsaved pay ticks on the project page: Save or Discard first
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
    for (const [k, v, cls] of rows) { if (v == null || v === "" || v === "–") continue;
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
  const docn = inv.doc_number || inv.invoice_no || "-";
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
    const present = rows.filter(r => r && r[1] != null && r[1] !== "" && r[1] !== "–");
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
// ── Collections report (owner 2026-09-02: "a summary like this ... the status of select open invoices ...
// to send off to financials to do cashflow forecast"). Pick invoices (or take everything shown), then
// Copy = one line per invoice in the owner's format
//   "CP831 - Town East - Invoice 34517 - $227k = <collections note>"
// or Excel = the same rows with the dates, days late, open, note, next follow-up, grouped by client.
const invPick = new Set();
function _invPickRows() { const shown = _invRows(); const picked = shown.filter(i => invPick.has(invKey(i))); return picked.length ? picked : shown; }
function _invPickUpdate() { const b = $("#ifCollect"); if (!b) return; const n = _invRows().filter(i => invPick.has(invKey(i))).length; b.textContent = n ? `Collections report (${n} picked)` : "Collections report"; b.classList.toggle("on", n > 0); }
function _kfmt(v) { const n = Math.abs(num(v)); const s = n >= 1e6 ? "$" + (n / 1e6).toFixed(n >= 1e7 ? 0 : 1) + "M" : n >= 1e3 ? "$" + Math.round(n / 1e3) + "k" : "$" + Math.round(n); return num(v) < 0 ? "(" + s + ")" : s; }
function _title(s) {   // "JPI-MERRIT PARK" -> "JPI-Merrit Park": short all-caps tokens (JPI, LLC, USA, TX) stay as they are
  return String(s || "").split(/(\s+|-|\/)/).map(w => (/^[A-Z0-9&.]{1,4}$/.test(w) ? w : w.toLowerCase().replace(/^([a-z])/, c => c.toUpperCase()))).join("");
}
function invCollectionsLines(rows) {
  const lines = [];
  for (const i of rows) {
    const nm = _title(nameOf(i.project_no) || "").trim();
    const status = (i.note || "").trim() || (i.next_followup ? "follow up " + fmtDateShort(i.next_followup) : "no note");
    lines.push(`${i.project_no || "(no project)"}${nm ? " - " + nm : ""} - Invoice ${i.doc_number || ""} - ${_kfmt(oiBal(i))} = ${status}` + (i.days_past_due > 0 ? ` (${i.days_past_due}d late)` : ""));
  }
  const tot = rows.reduce((s, i) => s + oiBal(i), 0);
  lines.push(`${rows.length} invoice${rows.length === 1 ? "" : "s"} - ${_kfmt(tot)} open - as of ${fmtDate(new Date(Date.now() - new Date().getTimezoneOffset() * 60000).toISOString().slice(0, 10))}`);
  return lines;
}
function openCollectionsReport() {
  const rows = _invPickRows(); if (!rows.length) { toast("Nothing to report - no invoices shown"); return; }
  const picked = _invRows().filter(i => invPick.has(invKey(i))).length;
  const ov = document.createElement("div"); ov.className = "xdlg-ov";
  const preview = invCollectionsLines(rows);
  ov.innerHTML = `<div class="xdlg xdlg-wide" role="dialog"><h3>Collections report - ${rows.length} invoice${rows.length === 1 ? "" : "s"}${picked ? " picked" : " shown (none picked)"}</h3>
    <pre class="xdlg-pre">${_ge(preview.join("\n"))}</pre>
    <div class="xdlg-actions"><button class="btn" id="xcCancel">Close</button><button class="btn" id="xcCopy">Copy these lines</button><button class="btn primary" id="xcXlsx">Excel for financials</button></div></div>`;
  document.body.appendChild(ov);
  const close = () => ov.remove(); ov.onclick = (e) => { if (e.target === ov) close(); };
  $("#xcCancel").onclick = close;
  $("#xcCopy").onclick = () => { copy(preview.join("\n")); toast(`Copied ${rows.length} lines`); close(); };
  $("#xcXlsx").onclick = () => { close(); invCollectionsXlsx(rows); };
}
function invCollectionsXlsx(rows) {
  const cols = [{ label: "Client" }, { label: "Project" }, { label: "Job" }, { label: "Invoice #" }, { label: "Invoice date" }, { label: "Due" }, { label: "Days late" },
                { label: "Open", type: "money" }, { label: "Invoice total", type: "money" }, { label: "Status / promise (collections note)" }, { label: "Next follow-up" }, { label: "Last action" }, { label: "Lien" }, { label: "Notion page" }];
  const data = rows.map(i => [i.customer || "", i.project_no || "", _title(nameOf(i.project_no) || ""), i.doc_number || "", i.txn_date || "", i.due_date || "", i.days_past_due > 0 ? i.days_past_due : 0,
                              oiBal(i), num(i.amount), i.note || "", i.next_followup || "", i.last_action_date || "", i.lien_status || "", i.notion_url || ""]);
  const fmt = []; rows.forEach((i, r) => { if (i.days_past_due > 0) fmt.push({ r, c: 6, cls: "neg" }); fmt.push({ r, c: 7, cls: oiBal(i) > 0 ? "neg" : "pos" }); if (i.note) fmt.push({ r, c: 9, cls: "warn" }); });
  const tot = rows.reduce((s, i) => s + oiBal(i), 0);
  const body = { name: "Collections report", sheet: "Collections", title: `Collections - ${rows.length} open invoice${rows.length === 1 ? "" : "s"} - ${money(tot)} open`,
    subtitle: `status per invoice for the cash-flow forecast · exported ${fmtDate(new Date(Date.now() - new Date().getTimezoneOffset() * 60000).toISOString().slice(0, 19), true)}`,
    columns: cols, rows: data, group_by: 0, fmt };
  toast("Building the Excel report…");
  fetch("/api/export/xlsx", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) }).then(r => r.json())
    .then(r => toast(r && r.ok ? `Collections report saved to Downloads (${r.rows} invoices) - opened in Finder` : "Export failed: " + ((r && r.error) || "unknown"))).catch(e => toast("Export failed: " + e));
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
      const htmlCells = [c, proj, i.doc_number || "", fmtDateShort(i.txn_date), fmtDateShort(i.due_date), dpd, money(amt)];
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
function _custForecastList(days, invs, expectedMs) {   // the invoices a Cash-in tile counts, with the date each is expected
  const rows = invs.map(i => ({ i, ms: expectedMs(i) })).filter(x => x.ms != null && (x.ms - Date.now()) / 86400000 <= days && oiBal(x.i) > 0.005).sort((a, b) => a.ms - b.ms);
  openRecord(`Cash-in within ${days} days`, `${rows.length} open invoices · projected from each client's average days-to-pay · not a QuickBooks figure`);
  const body = $("#recordBody"); body.innerHTML = "";
  const scroll = document.createElement("div"); scroll.className = "table-scroll"; scroll.style.padding = "0 18px 18px";
  const t = document.createElement("table"); t.className = "grid";
  t.innerHTML = "<thead><tr><th class='left'>Expected</th><th class='left'>Client</th><th class='left'>Project</th><th class='left'>Invoice</th><th class='left'>Invoiced</th><th class='right'>Open</th></tr></thead>";
  const tb = document.createElement("tbody"); let tot = 0;
  for (const { i, ms } of rows) { const tr = document.createElement("tr"); tr.style.cursor = "pointer"; tr.onclick = () => openInvoicePage(i);
    tr.appendChild(leftText(fmtDateShort(new Date(ms).toISOString().slice(0, 10)))); tr.appendChild(leftText(i.customer || "–")); tr.appendChild(leftText(i.project_no || "–")); tr.appendChild(leftText(i.doc_number || "–")); tr.appendChild(leftText(fmtDateShort(i.txn_date)));
    const oc = document.createElement("td"); oc.className = "right"; oc.appendChild(moneyCell(oiBal(i))); tr.appendChild(oc); tot += oiBal(i); tb.appendChild(tr); }
  const tr = document.createElement("tr"); tr.className = "ag-total"; const td = document.createElement("td"); td.className = "left"; td.colSpan = 5; td.textContent = "Total expected"; tr.appendChild(td); const oc = document.createElement("td"); oc.className = "right"; oc.appendChild(moneyCell(tot)); tr.appendChild(oc); tb.appendChild(tr);
  t.appendChild(tb); scroll.appendChild(t); body.appendChild(scroll);
}
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
      const tiles = [];   // Open AR + client count sit on the Company strip above (2026-09-13) - no double
      for (const d of order) tiles.push([d, money(divOpen(d))]);       // per-division open AR
      if (paySpeed.all_avg != null) tiles.push(["Cash-in ≤30d", money(f30), "fc"], ["≤60d", money(f60), "fc"], ["≤90d", money(f90), "fc"]);
      const arSrc = srcText("QuickBooks AR", loadedAt("AR (invoices)"), "loaded");
      for (const [l, v, cls] of tiles) {
        const k = el2("div", "kpi" + (cls ? " kpi-" + cls : "")); k.appendChild(el2("div", "k-label", l)); k.appendChild(el2("div", "k-value", v));
        if (cls === "fc") {   // a PROJECTION, not a QuickBooks figure - say so, and show which invoices it counts
          const days = parseInt((l.match(/(\d+)d/) || [])[1], 10) || 30;
          k.appendChild(srcChip("projection · client pay pattern", `Each open invoice placed at its client's average days-to-pay (${paySpeed.all_avg != null ? Math.round(paySpeed.all_avg) + "d portfolio average" : "portfolio average"} when the client has no history). Click to see the invoices counted.`));
          k.classList.add("kpi-click"); k.title = "Click to see the invoices this counts";
          k.onclick = () => _custForecastList(days, invs, expectedMs);
        } else k.appendChild(srcChip(arSrc, l === "Clients" ? "clients with an open invoice" : "open balance of the QuickBooks invoices loaded"));
        stats.appendChild(k); } } }
  const NCOL = 5;
  const tb = buildHead("#custTable", [["Client", "left"], ["Open AR", "right"], ["Open invoices", "right"], ["Oldest due", "left"], ["Avg days to pay", "right"]]);
  if (!tb) return; tb.innerHTML = "";
  if (!order.length) { const tr = document.createElement("tr"); const td = document.createElement("td"); td.colSpan = NCOL; td.className = "left"; td.style.color = "var(--text-dim)"; td.style.padding = "14px 12px"; td.textContent = "No open AR - run load_invoices.py."; tr.appendChild(td); tb.appendChild(tr); return; }
  for (const div of order) {
    const rows = [...byDiv.get(div).values()].sort((a, b) => b.open - a.open);   // top clients first
    // division band header (spans the row) - open AR + client count in this division
    // the division row uses the table's own columns (owner 2026-09-23: "put open AR under the open AR and just put the
    // clients next to division") - name + client count in Client, the total under Open AR, invoices under Open invoices
    const gtr = document.createElement("tr"); gtr.className = "bill-group cust-div";
    { const td = document.createElement("td"); td.className = "left";
      const key = document.createElement("span"); key.className = "bg-key"; key.textContent = div; td.appendChild(key);
      const n = document.createElement("span"); n.className = "cust-div-n"; n.textContent = ` ${rows.length} client${rows.length === 1 ? "" : "s"}`; td.appendChild(n);
      gtr.appendChild(td); }
    { const td = rightText(money(divOpen(div))); td.classList.add("cust-div-amt"); gtr.appendChild(td); }
    gtr.appendChild(rightText(String(rows.reduce((a, r) => a + (r.n || 0), 0))));
    gtr.appendChild(leftText("")); gtr.appendChild(rightText(""));
    tb.appendChild(gtr);
    for (const r of rows) {
      const tr = document.createElement("tr"); tr.style.cursor = "pointer"; tr.title = "Open this client's page - its invoices, paid and open";
      tr.onclick = () => openClientPage(r.client);   // dial in by client (owner 2026-09-16); "Open in Invoices" on the page applies the filter to the tracker
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

// ── The CLIENT page (owner 2026-09-16: "same thing for open invoices" - dial in by client, the invoice tracker within):
// the Open invoices grid filtered to one client, project bands, Open | All (paid included), the same row as the tracker
// (invoice # opens the invoice page), and "Open in Invoices" to carry the client into the tracker's own filters.
let _cp = null;   // { client, scope: "open" | "all", group }
async function openClientPage(client) {
  client = String(client || "").trim(); if (!client) return;
  if (_ppLeaveBlocked()) return;   // unsaved pay ticks on the project page: Save or Discard first
  openRecord(client, "loading…");
  _recSave({ k: "client", id: client });
  const anyOpen = (OI.invoices || []).some(i => (i.customer || "") === client && oiBal(i) > 0.005);
  _cp = { client, scope: anyOpen ? "open" : "all", group: true };   // nothing open -> start on every invoice, not an empty page
  renderClientPage();
}
function _cpData() { return (_cp.scope === "all" && OI_ALL) ? OI_ALL : OI; }
async function renderClientPage() {
  if (!_cp) return;
  const body = $("#recordBody"); body.innerHTML = "";
  if (_cp.scope === "all" && !OI_ALL) {
    skeletonInto(body, 4);
    try { OI_ALL = await (await fetch("/api/invoices/all")).json(); } catch (e) { OI_ALL = null; toast("Could not load all invoices"); _cp.scope = "open"; }
    body.innerHTML = "";
  }
  const D = _cpData(), buckets = D.buckets || ["Current", "1-30", "31-60", "61-90", "90+"];
  const invs = (D.invoices || []).filter(i => (i.customer || "") === _cp.client).sort((a, b) => String(b.txn_date || "").localeCompare(String(a.txn_date || "")));
  const openInvs = (OI.invoices || []).filter(i => (i.customer || "") === _cp.client);
  const open = openInvs.reduce((t, i) => t + oiBal(i), 0);
  const ps = ((OI.pay_speed || {}).by_client || {})[_cp.client.toLowerCase()];
  $("#recordSub").textContent = `${money(open)} open · ${openInvs.length} open invoice${openInvs.length === 1 ? "" : "s"}` + (ps && ps.avg_days != null ? ` · pays in ${ps.avg_days} days on average (${ps.n} paid)` : "")
    ;
  const tools = document.createElement("div"); tools.className = "cp-tools";
  const seg = document.createElement("div"); seg.className = "seg big";
  for (const [k, lbl] of [["open", `Open invoices · ${openInvs.length}`], ["all", "All invoices" + (OI_ALL ? ` · ${(OI_ALL.invoices || []).filter(i => (i.customer || "") === _cp.client).length}` : "")]]) {
    const b = document.createElement("button"); b.type = "button"; b.className = "seg-btn" + (_cp.scope === k ? " on" : ""); b.textContent = lbl; b.title = k === "all" ? "Every invoice on file for this client, paid ones included" : "What the client still owes";
    b.onclick = () => { _cp.scope = k; renderClientPage(); }; seg.appendChild(b);
  }
  tools.appendChild(seg);
  const grp = document.createElement("button"); grp.type = "button"; grp.className = "btn small"; grp.textContent = _cp.group ? "Flat list" : "Group by project"; grp.onclick = () => { _cp.group = !_cp.group; renderClientPage(); }; tools.appendChild(grp);
  const it = document.createElement("button"); it.type = "button"; it.className = "btn small"; it.textContent = "Open in Invoices"; it.title = "The Invoices tracker with its client filter set to this client - every other filter is there";
  it.onclick = () => { if (_ppLeaveBlocked()) return; invMSel.ifClient = new Set([_cp.client]); invMSel.ifProj = new Set(); _invMSelSig = null; const df = $("#ifDivision"); if (df) df.value = ""; closeRecord(); setTab("invoices"); renderOpenInvoices(); };
  tools.appendChild(it);
  const projs = [...new Set(invs.map(i => i.project_no).filter(Boolean))];
  if (projs.length) { const pb = document.createElement("span"); pb.className = "dim cp-projs"; pb.textContent = `${projs.length} project${projs.length === 1 ? "" : "s"}`; tools.appendChild(pb); }
  body.appendChild(tools);
  if (!invs.length) { const p = document.createElement("div"); p.className = "bills-cap"; p.textContent = _cp.scope === "open" ? "Nothing open for this client." : "No invoices on file for this client."; body.appendChild(p); return; }
  const scroll = document.createElement("div"); scroll.className = "table-scroll";
  const table = document.createElement("table"); table.className = "grid aging-grid cp-grid"; const thead = document.createElement("thead"), tbody = document.createElement("tbody");
  const cols = [["Project", "left"], ["Invoice #", "left"], ["Date", "left"], ["Amount", "right"], ["Status", "left"], ["Lien", "left"], ...buckets.map(b => [b, "right ag"])];
  thead.innerHTML = "<tr>" + cols.map(([c, al]) => `<th class="${al}">${_ge(c)}</th>`).join("") + "</tr>";
  const row = i => {   // the tracker's own row, minus the client column (it is the client's page), plus Amount + Status
    const tr = invRow(i, buckets); tr.removeChild(tr.firstElementChild);      // client
    const net = tr.children[3]; tr.removeChild(net);                          // net terms
    const amt = document.createElement("td"); amt.className = "right"; amt.appendChild(moneyCell(i.amount));
    const st = document.createElement("td"); st.className = "left status-col";
    const bal = oiBal(i), paid = bal <= 0.005;
    st.appendChild(stText(paid ? "Paid" + (i.paid_date ? " " + fmtDateShort(i.paid_date) : "") : (i.days_past_due > 0 ? `${i.days_past_due}d past due` : "Open"), paid ? "st-ok" : (i.days_past_due > 0 ? "st-warn" : "st-dim"),
      paid ? "The GC has paid this invoice" : (i.due_date ? "due " + fmtDateShort(i.due_date) : "")));
    if (i.litigation) st.appendChild(stText(" ⚖", "st-bad", "In litigation"));
    tr.insertBefore(st, tr.children[3]); tr.insertBefore(amt, tr.children[3]);
    return tr;
  };
  const grand = buckets.map(() => 0); for (const i of invs) if (oiBal(i) > 0.005) grand[i.bucket_index] += oiBal(i);
  if (_cp.group && projs.length > 1) {
    const byP = new Map(); for (const i of invs) { const k = i.project_no || "(no project)"; if (!byP.has(k)) byP.set(k, []); byP.get(k).push(i); }
    for (const [p, list] of [...byP].sort((a, b) => a[0].localeCompare(b[0], undefined, { numeric: true }))) {
      tbody.appendChild(invSubBand(p, nameOf(p), list.reduce((t, i) => t + oiBal(i), 0), list.length, cols.length));
      list.forEach((i, ix) => { const r = row(i); r.classList.add("in-proj"); if (ix === list.length - 1) r.classList.add("proj-last"); tbody.appendChild(r); });
    }
  } else for (const i of invs) tbody.appendChild(row(i));
  const ttr = document.createElement("tr"); ttr.className = "ag-total";
  const lead = document.createElement("td"); lead.className = "left"; lead.colSpan = 6; lead.textContent = `Total open · ${invs.length} invoice${invs.length === 1 ? "" : "s"} · ${money(invs.reduce((t, i) => t + num(i.amount), 0))} invoiced`; ttr.appendChild(lead);
  buckets.forEach((b, k) => { const td = document.createElement("td"); td.className = "right ag"; if (grand[k] > 0.005) { td.textContent = money(grand[k]); td.classList.add("ag" + k); } ttr.appendChild(td); });
  tbody.appendChild(ttr);
  table.appendChild(thead); table.appendChild(tbody); scroll.appendChild(table); body.appendChild(scroll);
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
const HINT_PAYMENTS = "Each row is a payment received. Click it to see the invoices (draws) it paid. Unlocks (AP) is the open vendor bills this payment funds: for staged CP/MFD draws, only the bills on the draw it paid; for RP (costs up front, billed once), the whole job's open AP. Net after AP = amount paid − that AP: what's left once those vendors are paid (red = the AP exceeds the payment).";
function renderPayments() {
  const body = $("#payBody"); if (!body) return;
  const pays = PAY.payments || [];
  { const n = $("#payNote"); if (n) n.textContent = pays.length ? `(${pays.length} payments · ${money(PAY.total_received)} received)` : "(no payment data - run load_payments.py)"; }
  body.innerHTML = "";
  const drawIdx = payOpenBillsByDraw(), projIdx = payOpenBillsByProject();
  const stats = document.createElement("div"); stats.className = "kpi-row";
  for (const [l, v] of [["Received (last 12 months)", money(PAY.total_received)], ["Payments", String(pays.length)],
                        ["Invoices paid", String(PAY.invoices_paid || 0)]]) {
    const k = el2("div", "kpi"); k.appendChild(el2("div", "k-label", l)); k.appendChild(el2("div", "k-value", v));
    k.appendChild(srcChip(srcText("QuickBooks payments", loadedAt("Payments"), "loaded"), "QBO Payment transactions (money in), a rolling 12-month window reloaded on every Resync")); stats.appendChild(k);
  }
  body.appendChild(stats);
  const head = document.createElement("div"); head.className = "list-head";
  head.title = HINT_PAYMENTS;   // the explanation lives on hover now (owner 2026-09-23: "remove 'each row is a payment', it is clutter")
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
        // the period row uses the table's own columns (the same fix as the Customer Center, owner 2026-09-23): the month and
        // its payment count under Client, the total under Amount Paid - never a floating mini-grid that misses the headers
        const iAmt = cols.findIndex(c => c[0] === "Amount Paid");
        for (let ci = 0; ci < cols.length; ci++) {
          const td = document.createElement("td"); td.className = ci === iAmt ? "right" : "left";
          if (ci === 0) { td.classList.add("pay-per");
            td.innerHTML = `<span class="bg-caret">${pExp ? "▾" : "▸"}</span> <span class="bg-key">${_ge(per.label)}</span> <span class="cust-div-n">${perN[per.key]} payment${perN[per.key] === 1 ? "" : "s"}</span>`; }
          else if (ci === iAmt) { td.classList.add("pay-per-amt"); td.textContent = money(perTot[per.key]); }
          gtr.appendChild(td);
        }
        tb.appendChild(gtr);
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
    { const rc = leftText(p.ref_no || "–"); if (p.att) { const ab = attBtn("Payment", p.qbo_txn_id, p.att, `${p.parent_customer || p.customer || ""} · payment ${p.ref_no || ""}`); ab.style.marginLeft = "6px"; rc.appendChild(ab); } tr.appendChild(rc); }
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
  if (note) note.textContent = `(window ${fmtDateShort(s.window_start)}–${fmtDateShort(s.window_end)})`;
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
    ? `(${t.customers || 0} customers · ${t.touches || 0} touches logged · Notion CRM)`
    : "(no CRM data - run load_customers.py)";

  // ── KPI stats (clickable → filter/jump) ──
  const stats = [
    ["Customers", String(t.customers || 0), "in the list - click to clear filters", () => setSalesFilter("", "")],
    ["Interested", String(t.interested || 0), "warm - click to see them", () => setSalesFilter("Interested", "")],
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
    tr.appendChild(leftText(c.division || "-"));
    tr.appendChild(leftText(c.sales_status || "-"));
    tr.appendChild(leftText(c.last_contacted ? fmtDate(c.last_contacted) : "-"));
    tr.appendChild(leftText(c.last_edited_by || "-"));
    tr.appendChild(rightText(String(c.n_touches || 0)));
    tb.appendChild(tr);
  }
  if (!rows.length) {
    const tr = document.createElement("tr"); const td = document.createElement("td");
    td.colSpan = 6; td.className = "left"; td.style.color = "var(--text-dim)";
    td.textContent = (S.customers || []).length ? "No customers match this filter." : "No CRM data - run load_customers.py.";
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
  if (!activeRep) { if (note) note.textContent = ""; box.appendChild(el2("p", "hint", "No rep activity - run load_customers.py.")); return; }
  const rep = activeRep;
  if (note) note.textContent = `- ${rep} · daily & weekly`;
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
  h.appendChild(el2("span", "rep-sub", `${mine.length} touches · ${myc.length} customers · ${wins} won · last active ${lastActive ? fmtDate(lastActive) : "-"}`));
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

  // WEEKLY timeline - last 12 weeks including zero weeks (so a drop-off shows)
  left.appendChild(el2("h4", null, "Weekly touches - last 12 weeks"));
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
    const body = el2("span", "rt-body"); body.appendChild(el2("span", "rt-cust", t.customer || "-"));
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
  right.appendChild(el2("h4", null, "Going stale - 21d+ no contact"));
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

// Portfolio P&L tab - every active job's live P&L + division/company totals. Computed
// server-side (/api/pnl/portfolio), lazy-loaded on first open, recomputed after a reload.
// ══ WIP REPORT ═══════════════════════════════════════════════════════════════
// The company Work-in-Progress schedule, straight from the ledger's wip_snapshot (loaded from the
// WIP master's Test tabs). Columns + order mirror the Test-Master reference; grouped by division
// with subtotals and a grand total. Read-only - the master workbook stays where you EDIT the WIP.
// Bonded is intentionally NOT here - dropped from the dashboard WIP view (the user
// 2026-08-20); the Excel Test tabs keep it. `cf` marks a column that carries a
// job-performance conditional format (see _wipCond).
const WIP_DIV_ORDER = ["Multi Family", "Commercial", "Residential"];

// ── WIP column widths (drag a header divider; persists per person) ──────────
const WIP_COL_DEFAULTS = { "Project #": 88, "Name": 190, "Total Contract": 122,
  "Est. Total Costs": 122, "Original Profit": 118, "GP %": 74, "Costs to Date": 116,
  "Cost to Complete": 128, "% Complete": 96, "Revenues Earned": 128, "Profit Earned": 116,
  "Billed": 112, "Overbillings": 116, "Underbillings": 118, "Left to Bill": 112,
  "Future Profit": 116, "Pure Job Borrow": 128 };

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
    if (v < 0.05) return { bg: _mix("var(--neg)", 20), fg: "var(--neg-text)", bold: true, title: "Margin very thin / negative" };
    if (v < 0.12) return { bg: _mix("#b8860b", 18), title: "Below-target margin" };
    if (v > 0.30) return { bg: _mix("#b8860b", 20), fg: "var(--warn-text)", title: "Unusually high GP% - verify for a missing cost" };
    return { bg: _mix("var(--pos)", 15), title: "Healthy margin" };
  }
  if (kind === "pctbar") {
    const v = r.percent_complete; if (v == null || v === "") return null;
    const p = Math.max(0, Math.min(100, v * 100));
    if (v > 1.0005) return { bar: 100, bg: _mix("var(--neg)", 18), fg: "var(--neg-text)", bold: true,
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
    return { bg: _mix("var(--neg)", a), fg: ratio > 0.08 ? "var(--neg-text)" : null, bold: ratio > 0.08,
      title: "Earned ahead of billed - unbilled work you are financing" };
  }
  if (kind === "borrow") {         // pure job borrow = cash the job pulls to finish
    const v = num(r.pure_job_borrow); if (v <= 0) return null;
    const ratio = v / contract;
    if (ratio < 0.05) return { bg: _mix("#b8860b", 16), title: "This job borrows some cash to finish" };
    return { bg: _mix("var(--neg)", 8 + Math.min(22, ratio * 130)), fg: "var(--neg-text)", bold: true,
      title: "Cost to complete exceeds what is left to bill - a cash drain" };
  }
  if (kind === "future") {         // remaining profit to earn
    const v = r.future_profit_to_earn; if (v == null || v === "") return null;
    if (v < 0) return { bg: _mix("var(--neg)", 20), fg: "var(--neg-text)", bold: true, title: "Expected profit eroded below what is already earned" };
    if (v > 0) return { bg: _mix("var(--pos)", 10), title: "Profit still ahead to earn" };
    return null;
  }
  if (kind === "neg0") {           // any money col that is a red flag when negative
    if (num(r[key]) < 0) return { bg: _mix("var(--neg)", 20), fg: "var(--neg-text)", bold: true, title: "Negative - losing money to date" };
    return null;
  }
  return null;
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
  const pctTxt = p => (p == null ? "–" : (p * 100).toFixed(1) + "%");
  $("#pnlNote").textContent = rows.length ? `(${comp.n || 0} active jobs · ${money(comp.billed)} net billed · WIP report ${meta.report_date ? fmtDate(meta.report_date) : "–"})` : "(no P&L data - load WIP + costs)";
  { const n2 = $("#pnlNote2"); if (n2) n2.textContent = $("#pnlNote").textContent; }
  renderCompanyHead();

  // company totals
  const tiles = [["Billed (gross)", money(comp.billed_gross)], ["Retainage held", money(comp.retainage)], ["Net billed", money(comp.billed)], ["Costs", money(comp.cost)],
    ["Overhead", money(comp.overhead)], ["Net", `${money(comp.net)} · ${pctTxt(comp.net_pct)}`, comp.net == null ? "" : (comp.net >= 0 ? "pos" : "neg")]];
  const tr = $("#pnlTotals"); tr.innerHTML = "";
  const pnlSrcOf = l => /gross|retainage/i.test(l) ? [srcText("WIP master", meta.report_date, "report"), "the WIP master's billed-to-date / retainage columns"]
    : /net billed/i.test(l) ? [srcText("QuickBooks invoices", loadedAt("AR (invoices)"), "loaded"), "every QuickBooks invoice on the active jobs, after retainage"]
    : /cost/i.test(l) ? [srcText("QuickBooks costs", loadedAt("Costs (QBO)"), "loaded"), "QuickBooks cost lines summed by job (line amounts, never bill totals)"]
    : /overhead/i.test(l) ? ["10% of contract (MFD 9%)", "the owner's overhead rule - a share of the contract, never of billed or cost"]
    : ["net billed - costs - overhead", "actuals, no earned-revenue proration"];
  for (const [l, v, cls] of tiles) {
    const k = el2("div", "kpi" + (cls ? " pnl-kpi-" + cls : ""));
    k.appendChild(el2("div", "k-label", l)); k.appendChild(el2("div", "k-value", v)); const [st, how] = pnlSrcOf(l); k.appendChild(srcChip(st, how)); tr.appendChild(k);
  }

  // by division
  let tb = buildHead("#pnlDivTable", [["Division", "left"], ["Jobs", "right"], ["Net billed", "right"], ["Cost", "right"], ["Overhead", "right"], ["Net", "right"], ["Net %", "right"]]);
  tb.innerHTML = "";
  for (const d of divs) {
    const row = document.createElement("tr");
    row.appendChild(leftText(d.division)); row.appendChild(rightText(String(d.n)));
    row.appendChild(rightText(money(d.billed))); row.appendChild(rightText(money(d.cost)));
    row.appendChild(rightText(money(d.overhead))); row.appendChild(rightText(money(d.net)));
    const pt = document.createElement("td"); pt.className = d.net >= 0 ? "pos" : "neg"; pt.textContent = pctTxt(d.net_pct); row.appendChild(pt);
    row.style.cursor = "pointer"; row.title = "Show this division's jobs below";
    row.onclick = () => { const f = $("#pnlFDivision"); if (f) { f.value = f.value === d.division ? "" : d.division; renderPnl(); } };
    tb.appendChild(row);
  }

  // by job - filterable + sortable (headers), click → detail
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
    ["status", "Status", "left"], ["contract", "Contract", "right"], ["pct_complete", "%", "right"], ["billed", "Net billed", "right"],
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
    row.appendChild(rightText(money(r.billed))); row.appendChild(rightText(money(r.cost)));
    row.appendChild(rightText(money(r.overhead))); row.appendChild(rightText(money(r.net)));
    const pt = document.createElement("td"); pt.className = r.net >= 0 ? "pos" : "neg"; pt.textContent = pctTxt(r.net_pct); row.appendChild(pt);
    // P&L updated = when this project's project-pnl Excel was last generated (owner 2026-08-19).
    const upd = document.createElement("td"); upd.className = "right"; upd.style.color = "var(--text-dim)"; upd.style.fontSize = ".88em";
    if (r.pnl_mtime) { upd.textContent = timeAgo(r.pnl_mtime); upd.title = "P&L Excel generated " + fmtDate(r.pnl_mtime, true); }
    else { upd.textContent = "not generated"; upd.classList.add("dim"); }   // dim colour, not a 55% ghost (AA audit 2026-09-08)
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


function renderKPIs() {
  const rows = ALL;
  const sum = k => rows.reduce((t, r) => t + (isNum(r[k]) ? r[k] : Number(r[k]) || 0), 0);
  const contract = sum("total_contract_price"), costs = sum("costs_to_date"),
        billed = sum("billed_to_date"), left = sum("left_to_bill"), under = sum("underbillings");
  const underN = rows.filter(r => num(r.underbillings) > 0).length, overN = rows.filter(isOverBudget).length,
        active = rows.filter(isActive).length;
  // Every card says where its number comes from and as of when (owner 2026-09-01). The WIP master's own
  // columns, summed; the two flags open the Needs-attention view of the grid below.
  const wipSrc = srcText("WIP master", meta.report_date, "report");
  const cards = [
    ["Total contract", money(contract), `${active} active of ${rows.length} jobs`, wipSrc, "sum of Total Contract Price across the WIP rows", null],
    ["Costs to date", money(costs), contract ? `${(costs / contract * 100).toFixed(0)}% of contract` : "", wipSrc, "sum of the WIP master's Costs to Date column - the report-date cut, not live QuickBooks", null],
    ["Billed to date", money(billed), contract ? `${(billed / contract * 100).toFixed(0)}% of contract` : "", wipSrc, "sum of Billed to Date (gross, incl. retainage)", null],
    ["Left to bill", money(left), "", wipSrc, "sum of Left to Bill = contract - billed", null],
    ["Underbilled (can invoice)", money(under), `${underN} job${underN === 1 ? "" : "s"} · financing the GC`, wipSrc, "earned ahead of billed - work you could invoice now; click to list the jobs", underN ? "attention" : null],
    ["Over budget", String(overN), overN ? "cost past the ETC · flatwork soft under $15k" : "no job past its ETC", srcText("QuickBooks costs", loadedAt("Costs (QBO)"), "loaded"), "cost to date over the ETC budget (a standing ruling takes a job off this list); click to list the jobs", overN ? "attention" : null],
  ];
  const row = $("#kpiRow"); if (!row) return; row.innerHTML = "";
  for (const [label, value, sub, src, how, go] of cards) {
    const el = document.createElement("div"); el.className = "kpi" + (go ? " kpi-click" : "") + (/under|over/i.test(label) && value !== "0" && value !== "$0" ? " kpi-neg" : "");
    el.innerHTML = `<div class="k-label"></div><div class="k-value"></div><div class="k-sub"></div>`;
    el.querySelector(".k-label").textContent = label;
    el.querySelector(".k-value").textContent = value;
    el.querySelector(".k-sub").textContent = sub;
    el.appendChild(srcChip(src, how));
    if (go) el.onclick = () => { projView = go; syncProjChips(); renderProjects(); $("#widget-projects").scrollIntoView({ behavior: "smooth", block: "start" }); };
    row.appendChild(el);
  }
}
// ── provenance chips: "<system> · <as-of>" under a figure, the formula in words on hover ──
// Load / sync times never print in the page (owner 2026-09-23: "remove ALL loads/sync data in the actual ledger and keep
// that status where it belongs in the top") - they live in the sync pill's breakdown. A WIP REPORT date stays: it
// says which report the numbers are from, not when they loaded.
function srcText(system, when, kind) {
  if (kind !== "report") return "";
  if (!when) return system;
  return `${system} · ${kind === "report" ? "report " + fmtDate(when) : "loaded " + fmtDate(when, true)}`;
}
function srcChip(text, how) {
  const s = document.createElement("div"); s.className = "k-src"; s.textContent = text || ""; if (how) s.title = how; if (!text) s.hidden = true; return s;
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
function syncedAt(src) { const v = ((meta.freshness || {}).sources || {})[src]; if (!v) return null; if (typeof v === "string") return v; return v.mtime || v.when || null; }   // e.g. "sync-ap" = the Bill Tracker workbook's file time

function renderDivisions() {
  { const h = document.querySelector("#widget-divisions .widget-head h2"); if (h) { let c = h.querySelector(".count"); if (!c) { c = document.createElement("span"); c.className = "count"; h.appendChild(c); }
      c.textContent = " " + srcText("WIP master", meta.report_date, "report") + " - contract, costs and billed are the report's own columns; the QuickBooks cost total is on Cost mix"; } }   // two costs on one page: say which is which
  const groups = {};
  for (const r of ALL) {
    const d = r.division || "-";
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
    tr.title = "Show this division's projects";
    tr.onclick = (e) => { if (!e.target.closest(".cell")) { projDiv = d; setTab("projects"); renderProjects(); } };
    [textCell(d, true), textCell(String(g.jobs)), moneyCell(g.contract), moneyCell(g.costs),
     moneyCell(g.billed), moneyCell(g.over), moneyCell(g.under)].forEach((c, i) => {
      const td = document.createElement("td"); if (i === 0) td.className = "left"; td.appendChild(c); tr.appendChild(td);
    });
    tbody.appendChild(tr);
  }
}
// Big-picture → zoom: click a division rollup row → its active projects.
const num = v => (isNum(v) ? v : Number(v) || 0);


// ── Projects: the WIP as the page (owner 2026-09-13) ─────────────────────
// One grid, every job. Active jobs by division (open), then "Completed recently" (closed but still
// carrying an open invoice or bill), then the settled closed jobs (folded). Columns = the WIP's own
// numbers plus what the job still owes / is owed: the ETC and open AP were dropped on purpose (owner:
// "drop ETC and open AP" - both are one click away on the project page). Row click opens the project page.
const PROJ_COLS = [
  { k: "project_no", label: "Project", t: "text" },
  { k: "project_name", label: "Name", t: "text" },
  { k: "client", label: "Client", t: "text" },
  { k: "total_contract_price", label: "Contract", t: "money" },
  { k: "costs_to_date", label: "Costs", t: "money" },
  { k: "billed_to_date", label: "Billed", t: "money" },
  { k: "percent_complete", label: "% compl.", t: "pct" },
  { k: "_overunder", label: "Over / (under)", t: "money" },
  { k: "_ar", label: "Open AR", t: "money" },
  { k: "_next", label: "Next", t: "text" },
];
let projView = "active";   // active | attention | completed | all
let projDiv = "";          // "" = every division
let projSort = { key: "total_contract_price", dir: -1 };
const PROJ_BANDS_LS = "proficient-ledger-projbands-v1";
let _projBands = (() => { try { return JSON.parse(localStorage.getItem(PROJ_BANDS_LS)) || {}; } catch { return {}; } })();
const isClosed = r => !isActive(r);
const needsAttention = r => isOverBudget(r) || num(r.underbillings) > 0 || num(r.pure_job_borrow) > 0;
function syncProjChips() {
  $$("#projView .seg-btn").forEach(b => b.classList.toggle("on", b.dataset.view === projView));
  $$("#projDiv .seg-btn").forEach(b => b.classList.toggle("on", b.dataset.div === projDiv));
}
function _projOpenAR() {   // project -> open AR (the QuickBooks invoices still unpaid)
  const m = {}; for (const i of (OI.invoices || [])) { const k = i.project_no || ""; if (!k) continue; m[k] = (m[k] || 0) + oiBal(i); } return m;
}
function _projNext() {     // project -> the funding position (CP/MFD draws; RP bills at completion, no draws)
  const m = {}; _buildDrawClientMap(DRAWS.draws || []);
  for (const f of _fundingRows()) m[f.pn] = f; return m;
}
function _projDerived() {
  const ar = _projOpenAR(), nx = _projNext();
  for (const r of ALL) {
    r._ar = ar[r.project_no] || 0;
    r._overunder = num(r.overbillings) - num(r.underbillings);
    const ap = AP.by_project && AP.by_project[r.project_no];
    r._apOpen = ap ? num(ap.open_balance) : 0;
    const f = nx[r.project_no];
    if (isClosed(r)) r._next = "Closed";
    else if (f) r._next = f.status === "Ready to collect" ? `Collect ${money(f.gcOwes)}${f.next && f.next.invoice_no ? " · #" + f.next.invoice_no : ""}`
      : f.status.startsWith("Blocked") ? `Blocked · pay ${money(f.blockAmt)} first`
      : f.status === "Settled" ? "Draws settled" : f.status === "No draw yet" ? "No draw yet" : f.status;
    else r._next = r._ar > 0.005 ? `Collect ${money(r._ar)}` : (String(r.project_no).startsWith("RP") ? "" : "");
    r._recent = isClosed(r) && (r._ar > 0.005 || r._apOpen > 0.005);   // closed but still settling
  }
}
function projRows() {
  const q = ($("#search") ? $("#search").value : "").trim().toLowerCase();
  return ALL.filter(r => {
    if (projDiv && r.division !== projDiv) return false;
    if (projView === "attention" && !needsAttention(r)) return false;
    if (projView === "completed" && !isClosed(r)) return false;
    if (q) { const hay = [r.project_no, r.project_name, r.builder_or_gc, r.client].filter(Boolean).join(" ").toLowerCase(); if (!hay.includes(q)) return false; }
    return true;
  });
}
function projSorted(list) {
  const c = PROJ_COLS.find(x => x.k === projSort.key) || PROJ_COLS[3];
  return [...list].sort((a, b) => {
    const av = a[c.k], bv = b[c.k];
    const na = av == null || av === "", nb = bv == null || bv === "";
    if (na && nb) return 0; if (na) return 1; if (nb) return -1;   // blanks always last
    return (c.t === "text" ? String(av).localeCompare(String(bv), undefined, { numeric: true }) : Number(av) - Number(bv)) * projSort.dir;
  });
}
function projBands(rows) {   // [{key, label, title, list, open}]
  const byDiv = list => WIP_DIV_ORDER.concat([...new Set(list.map(r => r.division || "Other"))].filter(d => !WIP_DIV_ORDER.includes(d)))
    .map(d => ({ key: d, label: d, list: list.filter(r => (r.division || "Other") === d), open: true })).filter(b => b.list.length);
  const recent = rows.filter(r => r._recent), settled = rows.filter(r => isClosed(r) && !r._recent);
  const done = [
    { key: "recent", label: "Completed recently", title: "Closed on the WIP but still carrying an open invoice or an open bill", list: recent, open: true },
    { key: "settled", label: "Completed · settled", title: "Closed on the WIP, nothing open either way", list: settled, open: false },
  ].filter(b => b.list.length);
  if (projView === "completed") return done;
  if (projView === "all") return byDiv(rows);
  if (projView === "attention") return byDiv(rows.filter(r => !isClosed(r))).concat(byDiv(rows.filter(isClosed)).map(b => ({ ...b, key: "closed:" + b.key, label: b.label + " · closed", open: false })));
  return byDiv(rows.filter(r => !isClosed(r))).concat(done);
}
function renderProjects() {
  const thead = $("#projTable thead"), tbody = $("#projTable tbody"); if (!thead || !tbody) return;
  _projDerived();
  const rows = projRows();
  { const c = $("#projCount"); if (c) c.textContent = `(${rows.length} of ${ALL.length})`; }
  { const sN = $("#projSrc"); if (sN) sN.textContent = `${srcText("WIP master", meta.report_date, "report")}`; }
  thead.innerHTML = ""; tbody.innerHTML = "";
  const htr = document.createElement("tr");
  for (const c of PROJ_COLS) {
    const th = document.createElement("th"); th.className = c.t === "text" ? "left" : "right"; th.textContent = c.label;
    if (projSort.key === c.k) { const a = document.createElement("span"); a.className = "arrow"; a.textContent = projSort.dir === 1 ? " ▲" : " ▼"; th.appendChild(a); }
    th.title = "Sort by " + c.label;
    th.onclick = () => { if (projSort.key === c.k) projSort.dir = -projSort.dir; else projSort = { key: c.k, dir: c.t === "text" ? 1 : -1 }; renderProjects(); };
    htr.appendChild(th);
  }
  thead.appendChild(htr);
  if (!ALL.length) { const tr = document.createElement("tr"); const td = document.createElement("td"); td.colSpan = PROJ_COLS.length; td.className = "left"; td.style.color = "var(--text-dim)"; td.style.padding = "14px 12px"; td.textContent = "No WIP data - run load_wip_master.py."; tr.appendChild(td); tbody.appendChild(tr); return; }
  if (!rows.length) { const tr = document.createElement("tr"); const td = document.createElement("td"); td.colSpan = PROJ_COLS.length; td.className = "left"; td.style.color = "var(--text-dim)"; td.style.padding = "14px 12px"; td.textContent = "No jobs match."; tr.appendChild(td); tbody.appendChild(tr); return; }
  for (const b of projBands(rows)) {
    const open = (b.key in _projBands) ? !!_projBands[b.key] : b.open;
    const gtr = document.createElement("tr"); gtr.className = "bill-group proj-band" + (open ? "" : " grp-closed");
    // One cell per column, so every subtotal sits UNDER its column (owner 2026-09-14: the merged
    // band strip did not line up with the grid). Name + count in the first cells; the number columns
    // carry the band's sums; % complete is the band's costs / ETC; Next stays blank.
    const sumOf = k => b.list.reduce((t, r) => t + num(r[k]), 0);
    const contract = sumOf("total_contract_price"), costs = sumOf("costs_to_date"), billed = sumOf("billed_to_date"),
          overUnder = sumOf("overbillings") - sumOf("underbillings"), ar = sumOf("_ar"), etc = sumOf("estimated_total_costs");
    const caret = document.createElement("span"); caret.className = "grp-caret"; caret.textContent = open ? "\u25BC\uFE0E" : "\u25B6\uFE0E";
    for (const c of PROJ_COLS) {
      const td = document.createElement("td");
      if (c.k === "project_no") { td.className = "left"; const key = document.createElement("span"); key.className = "bg-key"; key.appendChild(caret);
        key.appendChild(document.createTextNode(b.label)); if (b.title) key.title = b.title; td.appendChild(key); }
      else if (c.k === "project_name") { td.className = "left"; const n = document.createElement("span"); n.className = "bg-n"; n.textContent = `${b.list.length} job${b.list.length === 1 ? "" : "s"}`; td.appendChild(n);
        const dv = /resid/i.test(b.label) ? "RP" : /commer/i.test(b.label) ? "CP" : /multi/i.test(b.label) ? "MFD" : null;
        if (dv) { const rb = document.createElement("button"); rb.type = "button"; rb.className = "btn tiny proj-review"; rb.textContent = "WIP review →"; rb.title = `The weekly ${dv} review with the PM: every line, where each number came from, your answers`;
          rb.onclick = (e) => { e.stopPropagation(); openReviewDiv(dv); }; td.appendChild(rb); } }
      else if (c.k === "total_contract_price") { td.className = "right"; td.textContent = money(contract); }
      else if (c.k === "costs_to_date") { td.className = "right"; td.textContent = money(costs); }
      else if (c.k === "billed_to_date") { td.className = "right"; td.textContent = money(billed); }
      else if (c.k === "percent_complete") { td.className = "right"; td.textContent = etc > 0 ? pct(costs / etc) : "–"; if (etc > 0) td.title = "Band costs to date / band ETC"; }
      else if (c.k === "_overunder") { td.className = "right" + (overUnder < 0 ? " neg" : overUnder > 0 ? " pos" : ""); td.textContent = overUnder ? money(overUnder) : "–";
        td.title = overUnder < 0 ? "Net underbilled across the band" : overUnder > 0 ? "Net overbilled across the band" : ""; }
      else if (c.k === "_ar") { td.className = "right" + (ar > 0.005 ? " neg" : ""); td.textContent = ar > 0.005 ? money(ar) : "–"; }
      else td.className = "left";
      gtr.appendChild(td);
    }
    const kids = [];
    gtr.onclick = () => { const o = gtr.classList.contains("grp-closed"); gtr.classList.toggle("grp-closed", !o); caret.textContent = o ? "\u25BC\uFE0E" : "\u25B6\uFE0E"; kids.forEach(k => k.hidden = !o); _projBands[b.key] = o; try { localStorage.setItem(PROJ_BANDS_LS, JSON.stringify(_projBands)); } catch { /* ignore */ } };
    tbody.appendChild(gtr);
    for (const r of projSorted(b.list)) {
      const tr = document.createElement("tr"); tr.hidden = !open; tr.title = "Open this project";
      tr.onclick = (e) => { if (e.target.closest(".cell, a, button")) return; openProjectPage(r.project_no); };
      for (const c of PROJ_COLS) {
        const td = document.createElement("td"); const v = r[c.k];
        if (c.k === "project_no") { td.className = "left pn"; const sp = document.createElement("span"); sp.textContent = r.project_no; td.appendChild(sp);
          if (isOverBudget(r)) { const t = document.createElement("span"); t.className = "tag tag-ob"; t.textContent = "over budget"; t.title = `cost ${money(budgetCost(r))} vs ETC ${money(r.estimated_total_costs)}`; td.appendChild(t); }
          else if (r.over_budget_accepted) { const t = document.createElement("span"); t.className = "tag tag-known"; t.textContent = "known"; t.title = "A standing ruling: the owner settled this overrun once (job_rulings.json)"; td.appendChild(t); }
          if (num(r.pure_job_borrow) > 0) { const t = document.createElement("span"); t.className = "tag tag-borrow"; t.textContent = "borrowing"; t.title = "Cost to complete exceeds what is left to bill - " + money(r.pure_job_borrow); td.appendChild(t); } }
        else if (c.t === "text") { td.className = "left"; const sp = document.createElement("span"); sp.className = "nm"; sp.textContent = v == null || v === "" ? "–" : String(v); sp.title = sp.textContent; td.appendChild(sp); }
        else if (c.k === "percent_complete") { td.className = "right"; td.textContent = pct(v); const cond = _wipCond("pctbar", r); if (cond) { if (cond.bar != null) { td.classList.add("wip-bar"); td.style.setProperty("--bar", cond.bar + "%"); } if (cond.bg) td.style.background = cond.bg; if (cond.fg) td.style.color = cond.fg; if (cond.title) td.title = cond.title; } }
        else if (c.k === "_overunder") { td.className = "right"; const n = num(v); const sp = document.createElement("span"); sp.className = "cell" + (n < 0 ? " neg" : n > 0 ? " pos" : ""); sp.textContent = n ? money(n) : "–"; sp.title = n < 0 ? "Underbilled - earned ahead of billed, you are financing the job (click to copy)" : n > 0 ? "Overbilled - billed ahead of earned, holding the GC's cash (click to copy)" : ""; sp.onclick = () => copy(String(Math.round(n))); td.appendChild(sp);
          const cond = _wipCond(n < 0 ? "under" : "over", r); if (cond && cond.bg) td.style.background = cond.bg; }
        else if (c.k === "_ar") { td.className = "right"; if (num(v) > 0.005) { const mc = moneyCell(v); mc.classList.add("ar-open-amt"); td.appendChild(mc); } else td.textContent = "–"; }
        else if (c.k === "costs_to_date") { td.className = "right"; td.appendChild(moneyCell(v)); if (r.costs_loaded != null) td.title = `WIP report ${money(v)} · QuickBooks ${money(r.costs_loaded)}${loadedAt("Costs (QBO)") ? " (loaded " + fmtDate(loadedAt("Costs (QBO)"), true) + ")" : ""}`; }
        else { td.className = "right"; td.appendChild(moneyCell(v)); }
        tr.appendChild(td);
      }
      tbody.appendChild(tr); kids.push(tr);
    }
  }
}
// ── Company: the strip above Clients / Vendors / Money - the whole business in five numbers ──
function renderCompanyHead() {
  const host = $("#companyKpis"); if (!host) return; host.innerHTML = "";
  const invs = OI.invoices || []; const ar = invs.reduce((t, i) => t + oiBal(i), 0); const clients = new Set(invs.map(i => i.customer).filter(Boolean)).size;
  const ap = AP.summary || {};
  const pastDue = (AP.lien_watch || []).filter(r => r.lien_status === "Notice PAST due").length;
  const dueSoon = (AP.lien_watch || []).filter(r => r.lien_status === "Notice due in ≤7d").length;
  let collect = 0, collectAmt = 0; try { _buildDrawClientMap(DRAWS.draws || []); for (const f of _fundingRows()) if (f.status === "Ready to collect") { collect++; collectAmt += f.gcOwes; } } catch { /* no draws */ }
  const comp = (PNL && PNL.company) || null;
  const tiles = [
    ["Open AR", money(ar), `${clients} client${clients === 1 ? "" : "s"} · what they owe you`, srcText("QuickBooks", loadedAt("AR (invoices)"), "loaded"), "open balance of the QuickBooks invoices loaded", () => setTab("invoices"), false],
    ["Open AP", money(ap.open_balance), `${ap.open_lines || 0} open bill${ap.open_lines === 1 ? "" : "s"} · what you owe vendors`, srcText("Bill Tracker", syncedAt("sync-ap"), "loaded"), "open balance of the vendor bills in the Bill Tracker (subs excluded)", () => setTab("bills"), false],
    ["Lien notices", `${pastDue} past · ${dueSoon} soon`, "unpaid bills past the notice deadline, or within 7 days", srcText("Bill Tracker", syncedAt("sync-ap"), "loaded"), "bills you owe whose vendor lien-notice deadline has passed, or is within 7 days", () => setTab("liens"), pastDue > 0],
    ["Draws to collect", collect ? `${collect} · ${money(collectAmt)}` : "0", collect ? "nothing blocks them - collect from the GC" : "no draw is waiting on the GC", srcText("Bill Tracker + invoices", syncedAt("sync-ap"), "loaded"), "draws whose earlier-draw bills are paid: the GC owes the next one", () => { projView = "active"; syncProjChips(); setTab("projects"); }, false],
    ["Net P&L · active jobs", comp ? `${money(comp.net)} · ${comp.net_pct == null ? "–" : (comp.net_pct * 100).toFixed(1) + "%"}` : "computing…", comp ? `${comp.n || 0} jobs · net billed − costs − 10% of contract` : "", srcText("QuickBooks", loadedAt("Costs (QBO)"), "loaded"), "the live project P&L, active jobs only", () => setTab("money"), comp ? comp.net < 0 : false],
  ];
  for (const [label, value, sub, src, how, go, bad] of tiles) {
    const el = document.createElement("div"); el.className = "kpi kpi-click" + (bad ? " kpi-neg" : "");
    el.innerHTML = `<div class="k-label"></div><div class="k-value"></div><div class="k-sub"></div>`;
    el.querySelector(".k-label").textContent = label; el.querySelector(".k-value").textContent = value; el.querySelector(".k-sub").textContent = sub;
    el.appendChild(srcChip(src, how)); el.onclick = go; host.appendChild(el);
  }
  if (!comp && !PNL && typeof renderPnl === "function" && $("#pnlTotals")) renderPnl();   // fetch once; renderPnl re-paints this strip when it lands
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
  s.textContent = v || "-";
  return s;
}
function textCell(v, left) { const s = document.createElement("span"); s.textContent = v; const w = document.createElement("span"); w.appendChild(s); return w; }
function moneyCell(v) { const s = document.createElement("span"); s.className = "cell" + (num(v) < 0 ? " neg" : ""); s.textContent = money(v); s.onclick = () => copy(String(Math.round(num(v)))); s.title = "Click to copy"; return s; }

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

// P&L (project-pnl) link - shows when the workbook was last pulled, opens it, and
// (on an explicit confirm) runs project-pnl to (re)generate it. The generate call is
// the ONLY place the dashboard triggers a QBO pull + a file write; it is gated by a
// confirm dialog here and a `confirm` flag the server also requires.
function buildPnlGroup(proj) {
  const g = document.createElement("div"); g.className = "dgroup";
  const onPage = !!(_pp && _pp.pn === proj && $("#recordView") && !$("#recordView").hidden);   // on the project page the Profit & Loss table above already shows these rows
  if (!onPage) { const h = document.createElement("h4"); h.textContent = "P&L"; g.appendChild(h); }
  if (!onPage) {   // not when already ON the project page
    const b = document.createElement("button"); b.className = "btn small primary"; b.textContent = "Open project page"; b.style.marginBottom = "8px";
    b.onclick = (e) => { e.stopPropagation(); closePanels(); openProjectPage(proj); }; g.appendChild(b); }

  // ── live computed P&L (folded in from the spine; reconciles with project-pnl) ──
  const pl = document.createElement("div"); pl.className = "pnl-live"; pl.textContent = "computing…"; g.appendChild(pl);
  fetch(`/api/pnl/pl?proj=${encodeURIComponent(proj)}`).then(r => r.json()).then(d => {
    pl.innerHTML = "";
    if (d.error) { pl.textContent = d.error; return; }
    // the two folds (Invoices · Costs by code) sit side by side in one row, full width (owner 2026-09-23: less scrolling)
    let _foldCol = null, _foldsEl = null;
    const _folds = () => _foldsEl || (_foldsEl = pl.appendChild(Object.assign(document.createElement("div"), { className: "pnl-folds" })));
    const rowP = (k, v, cls) => {
      const r = document.createElement("div"); r.className = "drow" + (cls ? " " + cls : "");
      const a = document.createElement("span"); a.className = "dk"; a.textContent = k;
      const b = document.createElement("span"); b.className = "dv"; b.textContent = v;
      r.appendChild(a); r.appendChild(b); pl.appendChild(r); return r;
    };
    if (!d.has_wip && !onPage) rowP("Revenue basis", "no WIP snapshot", "pnl-sub");
    // the owner's reading order (2026-09-08): billed (gross) - retainage = net billed - costs - overhead = net
    if (!onPage) {
    rowP("Contract", money(d.contract));
    rowP("% complete", ((d.pct_complete || 0) * 100).toFixed(1) + "%", "pnl-sub");
    rowP("Billed to GC (gross)", money(d.billed_gross));
    rowP("Retainage held", "(" + money(d.retainage) + ")");
    rowP("Net billed", money(d.net_billed), "pnl-net").title = d.billed_src || "";
    if (d.billed_gap) { const w = rowP("WIP report shows more billed", money(d.billed_gap), "pnl-sub pnl-warn"); w.title = "The WIP report's billed-to-date is higher than the QuickBooks invoices loaded - Resync loads the invoice history."; }
    rowP("Costs to date", "(" + money(d.cost) + ")");
    rowP(`Overhead (${d.overhead_basis})`, "(" + money(d.overhead) + ")");
    const nr = rowP("Net", `${money(d.net)} · ${d.net_pct == null ? "–" : (d.net_pct * 100).toFixed(1) + "%"}`, "pnl-net");
    nr.classList.add(d.net >= 0 ? "pos" : "neg");
    }
    // The make-up of billed-to-date: every AR invoice (draw) the project has, paid or open
    // (owner 2026-08-21: "I need to see all the invoices the project has"). Oldest first.
    if (d.invoices && d.invoices.length) {
      // a fold, closed by default; date · invoice # · MEMO · amount · Paid <date> (owner 2026-09-23: "all draws collapsed by
      // default, include memo in between the invoice # and the amount. Paid needs a date after it")
      _foldCol = document.createElement("div"); _foldCol.className = "pnl-fold"; _folds().appendChild(_foldCol);
      const cap = document.createElement("div"); cap.className = "pnl-cap pnl-invcap"; cap.textContent = `Invoices - all draws (${d.invoices.length})`; cap.title = "Click to list the invoices"; _foldCol.appendChild(cap);
      const tbl = document.createElement("div"); tbl.className = "pnl-invoices";
      for (const iv of d.invoices) {
        const r = document.createElement("div"); r.className = "pnl-inv";
        const dt = document.createElement("span"); dt.className = "pi-date"; dt.textContent = fmtDateShort(iv.txn_date);
        const no = document.createElement("span"); no.className = "pi-no";
        if (iv.doc_number && iv.qbo_txn_id) { const a = document.createElement("a"); a.href = qboInvoiceUrl(iv.qbo_txn_id); a.target = "_blank"; a.rel = "noopener"; a.className = "qbo-link"; a.textContent = iv.doc_number; a.title = "Open this invoice in QuickBooks"; no.appendChild(a); }
        else no.textContent = iv.doc_number || "–";
        const memo = document.createElement("span"); memo.className = "pi-memo" + (iv.memo ? "" : " dim"); memo.textContent = iv.memo || "–"; memo.title = iv.memo || "";
        const am = document.createElement("span"); am.className = "pi-amt"; am.appendChild(moneyCell(iv.amount));
        const paid = num(iv.balance) <= 0;
        const st = document.createElement("span"); st.className = "pi-st";
        st.appendChild(stText(paid ? "Paid" + (iv.paid_date ? " " + fmtDateShort(iv.paid_date) : "") : "Open", paid ? "st-ok" : "st-warn",
          paid ? (iv.paid_date ? "GC paid " + fmtDateShort(iv.paid_date) : "GC paid (no payment date on file)") : ("Open AR balance " + money(iv.balance))));
        r.appendChild(dt); r.appendChild(no); r.appendChild(memo); r.appendChild(am); r.appendChild(st); tbl.appendChild(r);
      }
      _foldCol.appendChild(tbl);
    }
    if (d.by_code && d.by_code.length) {
      // grouped by JOB TYPE (owner 2026-09-08: "Slab > SL1 - Concrete"): the prefix is the pocket, each code
      // carries its cost-type name; uncoded / unknown prefixes land in "Other", last. Every code shows.
      // folded like the invoices, and beside them (owner 2026-09-23: "costs by cost code should be collapsed just like invoices")
      _foldCol = document.createElement("div"); _foldCol.className = "pnl-fold"; _folds().appendChild(_foldCol);
      const cap = document.createElement("div"); cap.className = "pnl-cap pnl-codecap"; cap.title = "Click to list the costs by code";
      cap.textContent = `Costs by code · by job type (${money(d.by_code.reduce((t, c) => t + num(c.amount), 0))})`; _foldCol.appendChild(cap);
      const tbl = document.createElement("div"); tbl.className = "pnl-codes";
      const groups = new Map();
      for (const c of d.by_code) { const g = c.job_type || c.prefix || "No cost code"; if (!groups.has(g)) groups.set(g, { total: 0, codes: [] }); const e = groups.get(g); e.total += num(c.amount); e.codes.push(c); }
      const ordered = [...groups].sort((a, b) => ((a[0] === "No cost code") - (b[0] === "No cost code")) || (b[1].total - a[1].total));
      for (const [g, e] of ordered) {
        const gh = document.createElement("div"); gh.className = "pnl-codegrp";
        const gn = document.createElement("span"); gn.textContent = g; const ga = document.createElement("span"); ga.textContent = money(e.total); gh.appendChild(gn); gh.appendChild(ga); tbl.appendChild(gh);
        e.codes.sort((x, y) => ((parseInt(x.number, 10) || 999) - (parseInt(y.number, 10) || 999)) || String(x.code).localeCompare(String(y.code)));
        for (const c of e.codes) {
          const r = document.createElement("div"); r.className = "pnl-code in-grp" + (c.uncoded ? " uncoded" : "");
          const nm = document.createElement("span"); nm.className = "pc-code"; nm.textContent = c.code;
          const n2 = document.createElement("span"); n2.className = "pc-name"; n2.textContent = (c.name ? "- " + c.name : "") + (c.is_sub ? " · sub" : "");
          const am = document.createElement("span"); am.className = "pc-amt"; am.textContent = money(c.amount);
          r.appendChild(nm); r.appendChild(n2); r.appendChild(am); tbl.appendChild(r);
        }
      }
      _foldCol.appendChild(tbl);
    }
  }).catch(() => { pl.textContent = "P&L unavailable."; });

  // ── source job folder (Synology CP/RP · OneDrive MFD) - the owner's "source link" ──
  const src = document.createElement("div"); src.className = "pnl-actions";
  const jobBtn = document.createElement("button"); jobBtn.className = "btn small"; jobBtn.textContent = "Open job folder ↗";
  jobBtn.title = "Open this job's folder on the file server (docs · takeoffs · photos)";
  jobBtn.onclick = () => fetch(`/api/job/open?proj=${encodeURIComponent(proj)}`, { method: "POST" })
    .then(r => r.json()).then(x => toast(x.error ? x.error : "Opening job folder…"));
  src.appendChild(jobBtn); g.appendChild(src);

  // ── detailed export (project-pnl Excel) - open / generate ──
  const cap2 = document.createElement("div"); cap2.className = "pnl-cap"; cap2.textContent = "Detailed export (project-pnl)"; g.appendChild(cap2);
  const row = document.createElement("div"); row.className = "drow pnl-pulled";   // the stamp sits in its own box right next to its label (owner 2026-09-08)
  const dk = document.createElement("span"); dk.className = "dk"; dk.textContent = "Last pulled";
  const dv = document.createElement("span"); dv.className = "dv pnl-stamp"; dv.textContent = "checking…";
  row.appendChild(dk); row.appendChild(dv); g.appendChild(row);
  const acts = document.createElement("div"); acts.className = "pnl-actions";
  const openBtn = document.createElement("button"); openBtn.className = "btn small"; openBtn.textContent = "Open Excel"; openBtn.disabled = true;
  const genBtn = document.createElement("button"); genBtn.className = "btn small"; genBtn.textContent = "Generate / Refresh";
  acts.appendChild(openBtn); acts.appendChild(genBtn); g.appendChild(acts);
  const msg = document.createElement("div"); msg.className = "pnl-msg"; g.appendChild(msg);

  const refresh = () => fetch(`/api/pnl?proj=${encodeURIComponent(proj)}`).then(r => r.json()).then(d => {
    if (d.error) { dv.textContent = "–"; return; }
    if (d.exists) {
      dv.textContent = `${timeAgo(d.mtime)} · ${fmtDate(d.mtime, true)}`;
      openBtn.disabled = false;
      openBtn.onclick = () => fetch(`/api/pnl/open?proj=${encodeURIComponent(proj)}`, { method: "POST" })
        .then(r => r.json()).then(x => toast(x.error ? x.error : "Opening P&L…"));
    } else {
      dv.textContent = "not generated yet"; openBtn.disabled = true;
    }
    msg.textContent = "";   // operator detail (mounts, fallbacks) stays off the page (owner 2026-09-16: "useless info")
  }).catch(() => { dv.textContent = "unavailable"; });

  genBtn.onclick = () => {
    if (!confirm(`Generate the P&L for ${proj}?\n\nThis runs project-pnl against QBO - a Touch ID prompt will appear on this Mac - and can take a minute or two.`)) return;
    genBtn.disabled = true; genBtn.textContent = "Generating…";
    msg.textContent = "Running project-pnl - watch for the Touch ID prompt on this Mac.";
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
    else if (s.state === "done") { finish("Done - P&L refreshed."); }
    else if (s.state === "error") { msg.textContent = "Failed: " + (s.detail || "see the log"); genBtn.disabled = false; genBtn.textContent = "Generate / Refresh"; }
    else { finish(""); }
  }).catch(() => setTimeout(tick, 3000));
  setTimeout(tick, 1500);
}

function openDetail(r) {
  detailRow = r;
  $("#detailTitle").textContent = `${r.project_no} - ${r.project_name || ""}`;
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
  // "Show every dollar" (the money trail) is parked everywhere for now (owner 2026-09-16: "not good ... remove for now"); trail.js stays
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
    const h = document.createElement("h4"); h.textContent = "Costs by code  ·  QuickBooks"; g.appendChild(h);
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
  const r = detailRow; const lines = [`${r.project_no} - ${r.project_name || ""}`];
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
// Browser Back works inside the app (owner 2026-09-15: "i can't use my back button without it refreshing
// all"): opening a record view or an RP job page pushes a history entry; popstate closes it in place.
let _histSkip = false;
function _pushView(state) { try { history.pushState(state, ""); } catch { /* ignore */ } }
function _popViewIfOwn(kind) {   // the in-app Back: pop our own entry so the browser's stack stays in step
  if (history.state && history.state.v === kind) { _histSkip = false; history.back(); return true; }
  return false;
}
window.addEventListener("popstate", () => {
  if (_histSkip) { _histSkip = false; return; }
  const rv = $("#recordView");
  if (rv && !rv.hidden) { if (typeof _ppLeaveBlocked === "function" && _ppLeaveBlocked()) { _pushView({ v: "record" }); return; } closeRecord(); return; }
  if (activeTab === "review" && typeof rrOpenLine !== "undefined" && rrOpenLine) { rrOpenLine = null; rrRenderCards(); window.scrollTo(0, 0); }
});

// A refresh reopens the page you were on (owner 2026-09-23: "why can't it remember where i was when i click
// refresh in the sub menus?"). The tab was always kept (localStorage); an opened vendor / project / client page
// and the vendor page's Bills | Payments view were not. Kept per browser tab (sessionStorage), cleared on Back.
const REC_SS = "proficient-ledger-record";
function _recSave(rec) { try { if (rec) sessionStorage.setItem(REC_SS, JSON.stringify(rec)); else sessionStorage.removeItem(REC_SS); } catch { /* private window */ } }
function _recLoad() { try { return JSON.parse(sessionStorage.getItem(REC_SS) || "null"); } catch { return null; } }
async function _recRestore(r) {
  if (!r || !r.id) return;
  if (r.k === "vendor") { await openVendorPage(r.id); if (r.view && r.view !== _vendorView && _vendorData) { _vendorView = r.view; renderVendorPage(); _recSave(r); } }
  else if (r.k === "project") { await openProjectPage(r.id); if (r.view && r.view !== "all" && _pp && _pp.pn === r.id) { _pp.view = r.view; _renderPpDraws(); } }
  else if (r.k === "client") await openClientPage(r.id);
}

function openRecord(title, sub) {
  if (typeof _ppLeaveBlocked === "function" && _ppLeaveBlocked()) return false;   // an unsaved pay run never walks away with you
  if (!($("#recordView") && !$("#recordView").hidden)) _pushView({ v: "record" });   // one entry per open, not per re-render
  $$(".tab-page").forEach(p => { p.hidden = true; });
  $("#recordView").hidden = false;
  $("#recordTitle").textContent = title || "";
  $("#recordSub").textContent = sub || "";
  window.scrollTo(0, 0);
}
function closeRecord() {
  if (typeof _ppLeaveBlocked === "function" && _ppLeaveBlocked()) return;
  _recSave(null);
  const rv = $("#recordView"); if (rv) rv.hidden = true;
  $$(".tab-page").forEach(p => { p.hidden = p.dataset.tab !== activeTab; });
  window.scrollTo(0, 0);
}

// ── Copy + CSV + toast ────────────────────────────────────────────────────
let toastTimer = null;
function toast(msg, ms) {
  const t = $("#toast"); t.textContent = msg; t.hidden = false;
  clearTimeout(toastTimer); toastTimer = setTimeout(() => (t.hidden = true), ms || 1400);
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
  _projDerived();
  const rows = projBands(projRows()).flatMap(b => projSorted(b.list).map(r => ({ ...r, _band: b.label })));
  return { name: "projects", rows, cols: [["Band", r => r._band], ["Division", r => r.division], ["Status", r => r.status]]
    .concat(PROJ_COLS.map(c => [c.label, r => c.t === "money" || c.t === "pct" ? (r[c.k] == null || r[c.k] === "" ? "" : Number(r[c.k])) : nz(r[c.k])])) };
}
// Export dialog (owner 2026-09-02): Excel (a grouped report that keeps the state colours) or CSV, and
// "how would you like it grouped" - defaults to the grouping on screen (Bills: the Group by select).
const EXPORT_GROUPS = {
  bills: [["", "None"], ["Vendor", "Vendor"], ["Project", "Project"], ["Client", "Client"], ["Division", "Division"], ["Invoice", "Draw / invoice"]],
  invoices: [["", "None"], ["Client", "Client"], ["Project", "Project"], ["Status", "Status"]],
  projects: [["", "None"], ["Band", "Band"], ["Division", "Division"], ["Client", "Client"]],
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
  const moneyLabels = new Set(["This line", "Bill total", "Open balance", "Invoice total", "Amount", ...PROJ_COLS.filter(c => c.t === "money").map(c => c.label)]);
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
  $("#setDensity").value = settings.density;
}
function wireSettings() {
  const on = (sel, ev, fn) => $(sel).addEventListener(ev, fn);
  on("#setTheme", "change", e => { settings.theme = e.target.value; saveSettings(); applySettings(); });
  on("#setDensity", "change", e => { settings.density = e.target.value; saveSettings(); applySettings(); });
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
    ? `QBO metrics as of ${HEALTH.as_of_label}` : "QBO metrics not pulled yet - derived rows only";
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
  band("Debt service (liability balance drops - MCA rows shown, never counted)");
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
    idc.textContent = r.id;
    // The one-page guide for this process (vault assets/processes) sits under the ID, the same
    // spot on every row - trailing the process text it was lost at the end of a long cell.
    if (r.guide && r.guide.length) {
      const g = document.createElement("a");
      g.className = "sys-guide"; g.textContent = "Guide";
      g.href = "/api/process-guide?id=" + encodeURIComponent(r.id) + "&fmt=" + r.guide[0];
      g.target = "_blank"; g.rel = "noopener";
      g.title = "Open the one-page guide for " + r.id;
      idc.appendChild(g);
    }
    tr.appendChild(idc);

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
    const short = (d.title.split(/[\u2014\u2013(-]/)[0] || d.title).trim();
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
let wrView = "slides";        // "slides" = one job at a time (default) · "list" = every job
let wrIdx = 0;                // the slide being shown
let wrTouched = {};           // pn -> "acc" | "skip" once the owner decided the whole job

const WR_DIV_ORDER = ["Commercial", "Residential", "Multi-Family"];
const WR_DIV_SHORT = { "Commercial": "CP", "Residential": "RP", "Multi-Family": "MFD" };
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
  if (!wrChosen && !force) { wrChooser(); return; }
  wrInitDecisions();
  renderWipReview();
}

// Opening WIP Review asks first (owner 2026-09-23: "it says last computed this day and time - would you like to
// recompute or continue with this compute? so that every time i click it it doesn't need to compute"). Once per
// page load; the stamps come per division, and a review computed before the WIP file's last save is called out.
let wrChosen = false;
function wrChooser() {
  const body = $("#wrBody"), g = WR.generated || {};
  $("#wrFilters").hidden = true; $("#wrSync").hidden = true; $("#wrStats").innerHTML = "";
  const ats = WR_DIV_ORDER.filter(d => g[d] && g[d].at).map(d => g[d].at);
  const oldest = ats.length ? ats.slice().sort()[0] : null;
  const behind = WR.wip_file_at && oldest && WR.wip_file_at > oldest;
  const today = new Date().toISOString().slice(0, 10);
  const lines = WR_DIV_ORDER.filter(d => g[d]).map(d => {
    const v = g[d], old = v.at && v.at.slice(0, 10) < today;
    return `<div class="wr-ch-row"><span class="wr-ch-div">${d === "Multi-Family" ? "MFD" : d === "Commercial" ? "CP" : "RP"}</span>
      <span class="${old ? "wr-stale" : ""}">${v.at ? fmtDate(v.at, true) : "–"}${old ? " · not today" : ""}</span></div>`;
  }).join("");
  body.innerHTML = `<div class="wr-chooser">
      <div class="wr-ch-title">Last computed</div>${lines}
      ${behind ? `<div class="wr-ch-warn">The WIP file was saved after this (${_ge(fmtDate(WR.wip_file_at, true))}) - recompute before you sync.</div>` : ""}
      <div class="wr-ch-actions">
        <button class="btn big ${behind ? "" : "primary"}" id="wrUse" type="button">Use this one</button>
        <button class="btn big ${behind ? "primary" : ""}" id="wrRecompute" type="button">Recompute (about 3 min)</button>
      </div></div>`;
  $("#wrUse").onclick = () => { wrChosen = true; wrInitDecisions(); renderWipReview(); };
  $("#wrRecompute").onclick = () => { wrChosen = true; runWipReview(true); };
}

// The owner's Accept / Keep choices survive a page refresh (owner 2026-09-23: "if i refresh will it save the
// original choice of the previous project?"). Kept in this browser, keyed to the review they were made on -
// a fresh Compute has a new stamp, so old choices never ride onto new numbers. Sync still writes from memory.
const WR_LS = "proficient-wip-review-choices-v1";
const wrStamp = () => JSON.stringify(Object.entries((WR && WR.generated) || {}).map(([d, v]) => [d, v.at]).sort());
function wrSaveChoices() {
  try { localStorage.setItem(WR_LS, JSON.stringify({ stamp: wrStamp(), decisions: wrDecisions, drop: [...wrDrop], touched: wrTouched, idx: wrIdx })); }
  catch { /* private window - choices live until the refresh */ }
}
function wrInitDecisions() {
  // Fresh review → default marks: QBO facts approved, PM fields left for an answer.
  wrDecisions = {}; wrDrop = new Set(); wrTouched = {}; wrIdx = 0;
  for (const r of WR.records) {
    if (r.status === "SAME") continue;
    const marks = {};
    for (const f of wrChanged(r.fields)) marks[f.key] = (f.block === "qbo");
    wrDecisions[r.project_num] = marks;
  }
  try {   // then the choices already made on THIS review, if the page was refreshed
    const saved = JSON.parse(localStorage.getItem(WR_LS) || "null");
    if (saved && saved.stamp === wrStamp()) {
      for (const pn in saved.decisions || {}) if (wrDecisions[pn]) Object.assign(wrDecisions[pn], saved.decisions[pn]);
      wrDrop = new Set(saved.drop || []); wrTouched = saved.touched || {}; wrIdx = saved.idx || 0;
    }
  } catch { /* nothing saved */ }
}

function renderWipReview() {
  const note = $("#wrNote"), body = $("#wrBody");
  const g = WR.generated || {};
  if (note) {   // one as-of PER DIVISION, the stale one in red with its reason (owner 2026-09-04: one date hid an 08/25 CP file)
    note.innerHTML = "";
    const parts = WR_DIV_ORDER.filter(dv => g[dv]).map(dv => { const v = g[dv]; const sp = document.createElement("span"); sp.className = "wr-asof" + (v.stale ? " wr-stale" : "");
      sp.textContent = `${dv === "Multi-Family" ? "MFD" : dv === "Commercial" ? "CP" : "RP"} ${v.at ? fmtDate(v.at, true) : "–"}${v.stale ? " STALE" : ""}`; if (v.reason) sp.title = v.reason; return sp; });
    note.append("computed "); parts.forEach((sp, i) => { if (i) note.append(" · "); note.appendChild(sp); });
    const src = (WR.sources || {}).Commercial; if (src && !src.mounted) { const w = document.createElement("span"); w.className = "wr-stale"; w.textContent = " · CP source (Common drive) not mounted"; w.title = "Mount smb://10.27.10.100/Common, then Compute again"; note.appendChild(w); }
  }
  $("#wrFilters").hidden = false; $("#wrSync").hidden = false;
  renderWrStats();
  if (wrView === "slides") { renderWrSlides(); return; }
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
    recs = recs.filter(r => wrHasBlock(r, wrBlockSel()));
    if (!recs.length) continue;
    const gen = g[dv];
    const head = document.createElement("div");
    head.className = "wr-div-head";
    head.innerHTML = `<span>${dv}</span><span class="wr-div-sub">${gen ? gen.tab : ""} · ${recs.length} shown${gen && gen.at ? " · computed " + _ge(fmtDate(gen.at, true)) : ""}${gen && gen.stale ? ` <span class="wr-stale" title="${_ge(gen.reason || "")}">STALE - ${_ge(gen.reason || "older than the other divisions")}</span>` : ""}</span>`;
    body.appendChild(head);
    for (const r of recs) { body.appendChild(wrJobCard(r)); shown++; }
  }
  if (!shown) body.innerHTML = `<div class="wr-empty"><p>Nothing matches the filters.</p></div>`;
  wrUpdateApproveCount();
}

// Show: QuickBooks + PM / QuickBooks only / PM questions only (owner 2026-09-23: "a filter for PM questions").
// A job is shown when it has a change in that block; the list card shows only that block.
const wrBlockSel = () => ($("#wrBlock") && $("#wrBlock").value) || "";
const wrHasBlock = (r, b) => !b || r.status === "REMOVED" || wrChanged(r.fields || []).some(f => f.block === b);
// "blank" and $0 are the same answer for a change-order field - never a question for a PM
const wrPmTrivial = f => (f.key === "approved_cos" || f.key === "co_costs") && !Number(f.was || 0) && !Number(f.now || 0);

// The PM questions as a Teams message, one block per division, for the jobs the filters show
// (owner 2026-09-23: "copy for all divisions to send via Teams so they can approve on their own time").
function wrCopyPmQuestions() {
  const div = $("#wrDivision").value, q = ($("#wrSearch").value || "").trim().toLowerCase();
  const m = v => v == null || v === "" ? "blank" : money(v);
  const out = [`WIP update - PM questions (${fmtDateShort(new Date().toISOString().slice(0, 10))})`,
    "Reply OK, or the right number and the document it comes from, for each line.", ""];
  let n = 0;
  for (const dv of WR_DIV_ORDER) {
    if (div && div !== dv) continue;
    const jobs = WR.records.filter(r => r.division === dv && r.status !== "SAME"
      && (!q || (r.project_num + " " + r.name).toLowerCase().includes(q)))
      .map(r => [r, wrChanged(r.fields || []).filter(f => f.block === "pm" && !wrPmTrivial(f))])
      .filter(([, fs]) => fs.length)
      .sort((a, b) => a[0].project_num.localeCompare(b[0].project_num));
    if (!jobs.length) continue;
    out.push(`${dv.toUpperCase()}`);
    for (const [r, fs] of jobs) {
      out.push(`${r.project_num} ${r.name || ""}`.trim());
      for (const f of fs) {
        const src = [f.source, f.note].filter(Boolean).join(" · ");
        out.push(`  - ${f.label}: ${m(f.was)} now, ${m(f.now)} proposed${f.reversed ? " (GOES DOWN)" : ""}${src ? " - from " + src : ""}`);
        n++;
      }
    }
    out.push("");
  }
  if (!n) { toast("No PM questions in what is shown"); return; }
  copy(out.join("\n").trim().replace(/\u2014/g, "-")).then(() => toast(`Copied ${n} PM question${n === 1 ? "" : "s"} - paste in Teams`));
}

// ── one job at a time ────────────────────────────────────────────────────────
function wrVisible() {
  // The same filters as the list, but ONE mixed stream (no division bands): what
  // changes first, then new jobs, then the ones coming off; job # inside each.
  const div = $("#wrDivision").value, st = $("#wrStatus").value;
  const q = ($("#wrSearch").value || "").trim().toLowerCase();
  const changedOnly = $("#wrChangedOnly").checked;
  const order = { CHANGED: 0, REVERSED: 0, ADDED: 1, REMOVED: 2, SAME: 3 };
  const blk = wrBlockSel();
  return WR.records.filter(r => (!div || r.division === div) && (!changedOnly || r.status !== "SAME")
      && (!st || r.status === st) && (!q || (r.project_num + " " + r.name).toLowerCase().includes(q)) && wrHasBlock(r, blk))
    .sort((a, b) => (order[a.status] - order[b.status]) || a.project_num.localeCompare(b.project_num));
}

function wrWhat(r) {
  // The one line the owner reads first, in plain words.
  if (r.status === "ADDED") return "New on the WIP";
  if (r.status === "REMOVED") return "Comes off the WIP";
  const ch = wrChanged(r.fields).map(f => f.label.toLowerCase());
  if (!ch.length) return "No change";
  return "Changes: " + ch.join(", ");
}

function wrWhy(r) {
  // Why, from the row's own notes/flags and the section it sits in - never invented.
  const bits = [];
  if (r.section) bits.push(r.section);
  if (r.flags) bits.push(r.flags);
  const rev = wrChanged(r.fields).filter(f => f.reversed);
  if (rev.length) bits.push("Went DOWN: " + rev.map(f => f.label.toLowerCase()).join(", ") + " - the source is named on the line");
  return bits.join(" · ");
}

function renderWrSlides() {
  const body = $("#wrBody");
  const recs = wrVisible();
  body.innerHTML = "";
  if (!recs.length) { body.innerHTML = `<div class="wr-empty"><p>Nothing matches the filters.</p></div>`; wrUpdateApproveCount(); return; }
  if (wrIdx >= recs.length) wrIdx = recs.length - 1;
  if (wrIdx < 0) wrIdx = 0;
  const r = recs[wrIdx];
  const decided = Object.keys(wrTouched).filter(pn => recs.some(x => x.project_num === pn)).length;
  const nav = document.createElement("div"); nav.className = "wr-slide-nav";
  nav.innerHTML = `<button class="btn small" id="wrPrev" type="button" ${wrIdx === 0 ? "disabled" : ""}>← Previous</button>
    <span>${wrIdx + 1} of ${recs.length} · ${decided} decided</span>
    <button class="btn small" id="wrNext" type="button" ${wrIdx >= recs.length - 1 ? "disabled" : ""}>Next →</button>`;
  body.appendChild(nav);
  body.appendChild(wrSlide(r));
  $("#wrPrev").onclick = () => { wrIdx--; renderWrSlides(); };
  $("#wrNext").onclick = () => { wrIdx++; renderWrSlides(); };
  wrUpdateApproveCount();
}

// Direction between "on the WIP now" and "after this update": up / down / level. A drop is
// coloured because on MFD/CP billed and costs only move up - a lower number is a finding.
function wrDir(was, now) {
  const a = Number(was) || 0, b = Number(now) || 0;
  if (b > a) return '<span class="wr-dir up" title="Goes up">▲</span>';
  if (b < a) return '<span class="wr-dir down" title="Goes down">▼</span>';
  return '<span class="wr-dir" title="No change">=</span>';
}

// The QBO lines behind a Costs / Billed change, grouped under its row (owner 2026-09-23: "i want to see the
// line transactions of what is adding/removing"). /api/wip/lines walks QBO's history (entered / edited /
// deleted, from the mirror) back to the moment the WIP number was true; the lines that differ since then
// ARE the change, and the head says whether they add up to it.
const asofRaw = gen => gen && gen.at ? gen.at : "";
const wrLinesWanted = f => f.changed && (f.key === "costs" || (f.key === "billed" && /quickbooks/i.test(f.source || "")));
const WR_TAG = { new: "+ new", edited: "edited", deleted: "− deleted" };
async function wrAttachLines(row, r, f, since) {
  if (!row) return;
  const q = new URLSearchParams({ no: r.project_num, field: f.key, was: f.was || 0, now: f.now || 0, since });
  let d;
  try { d = await (await fetch("/api/wip/lines?" + q)).json(); } catch { d = { ok: false, error: "no answer from the server" }; }
  if (!row.isConnected) return;
  const tr = (cls, cells) => { const e = document.createElement("tr"); e.className = cls; e.innerHTML = cells; return e; };
  const frag = document.createDocumentFragment();
  if (!d.ok) { frag.appendChild(tr("wr-lines-note", `<td colspan="5">Could not list the lines - ${_ge(d.error || "")}</td>`)); row.after(frag); return; }
  const n = d.lines.length;
  const tie = d.ties ? `<span class="wr-tie ok">adds up</span>`
    : `<span class="wr-tie gap">${money(Math.abs(d.gap))} of the change is not in QuickBooks' history</span>`;
  const sinceTxt = d.since ? ` · entered since ${fmtDateShort(d.since.slice(0, 10))}` : "";
  const head = tr("wr-lines-head", `<td colspan="5"><span class="wr-lh-key">${n} QuickBooks line${n === 1 ? "" : "s"}</span> · ${money(d.explained)}${sinceTxt} · ${tie}</td>`);
  head.dataset.grpkey = `${r.project_num}|${f.key}`;
  frag.appendChild(head);
  for (const l of d.lines) {
    const ref = l.doc_number ? (l.qbo_url ? `<a href="${_ge(l.qbo_url)}" target="_blank" rel="noopener" title="Open in QuickBooks">#${_ge(l.doc_number)}</a>` : `#${_ge(l.doc_number)}`) : "";
    // a bill dated well before it was entered is flagged - that is what hid CP742's JCP bills (dated 01/01, entered 09/03)
    const backdated = l.entered && l.date && (new Date(l.entered) - new Date(l.date)) > 45 * 864e5 ? `entered ${fmtDateShort(l.entered)}` : "";
    const what = [l.code, l.description, l.was != null ? `was ${money(l.was)}` : "", !l.code && !l.description ? l.memo : "", backdated].filter(Boolean).map(_ge).join(" · ");
    frag.appendChild(tr(`wr-line ${l.tag}`,
      `<td><span class="wr-line-tag">${WR_TAG[l.tag] || ""}</span> ${_ge(l.party || "")} ${ref}</td>`
      + `<td class="wr-line-date">${l.date ? fmtDateShort(l.date) : ""}</td>`
      + `<td class="dir">${wrDir(0, l.amount)}</td>`
      + `<td class="wr-line-amt">${money(l.amount)}</td>`
      + `<td class="src">${what}</td>`));
  }
  row.after(frag);
}

function wrSlide(r) {
  const card = document.createElement("div");
  card.className = "wr-slide wr-" + r.status.toLowerCase();
  const gen = (WR.generated || {})[r.division] || {};
  const asof = gen.at ? fmtDate(gen.at, true) : "";
  const removed = r.status === "REMOVED", added = r.status === "ADDED";
  const state = wrTouched[r.project_num];
  let html = `<div class="wr-slide-head"><span class="wr-slide-pn">${_ge(r.project_num)}</span>
      <span class="wr-slide-name">${_ge(r.name)}</span>
      <span class="wr-slide-div">${_ge(WR_DIV_SHORT[r.division] || r.division)} · ${_ge(r.tab || "")}</span></div>
    <div class="wr-slide-what">${_ge(wrWhat(r))}</div>`;
  const why = wrWhy(r); if (why) html += `<div class="wr-slide-why">${_ge(why)}</div>`;
  html += `<table class="wr-tbl"><thead><tr><th></th><th>On the WIP now${asof ? " (" + _ge(asof) + ")" : ""}</th><th class="dir"></th><th>After this update</th><th style="text-align:left">Where the new number comes from</th></tr></thead><tbody>`;
  for (const f of r.fields || []) {
    const chg = f.changed, cls = (chg ? "chg" : "") + (f.reversed ? " rev" : "");
    const src = chg ? (f.source || "") + (f.note ? (f.source ? " · " : "") + f.note : "") : "";
    html += `<tr class="${cls}" data-fkey="${_ge(f.key)}"><td>${_ge(f.label)}</td><td class="was">${money(f.was)}</td><td class="dir">${chg ? wrDir(f.was, f.now) : ""}</td><td class="now">${chg ? money(f.now) : ""}</td><td class="src">${_ge(src)}</td></tr>`;
  }
  html += `</tbody></table>`;
  html += `<div class="wr-slide-actions">
      <button class="btn big primary" id="wrAcc" type="button">${added ? "Add it" : removed ? "Take it off" : "Accept"}</button>
      <button class="btn big" id="wrSkip" type="button">${added ? "Leave it off" : removed ? "Keep it on" : "Keep as is"}</button>
      <span class="wr-slide-state ${state === "acc" ? "acc" : state === "skip" ? "skip" : ""}">${state === "acc" ? "✓ accepted" : state === "skip" ? "kept as is" : "not decided"}</span></div>
    <div class="wr-slide-keys">Keys: A accept · S keep · ← → move · Sync writes when you are done</div>`;
  card.innerHTML = html;
  for (const f of r.fields || []) if (wrLinesWanted(f)) wrAttachLines(card.querySelector(`.wr-tbl tr[data-fkey="${f.key}"]`), r, f, asofRaw(gen));
  card.querySelector("#wrAcc").onclick = () => wrDecideJob(r, true);
  card.querySelector("#wrSkip").onclick = () => wrDecideJob(r, false);
  return card;
}

function wrDecideJob(r, yes) {
  // Whole-job decision: every changed field approved (or none), an added job kept on
  // (or left off). A REMOVED job has no approvable fields - "keep it on" is honoured by
  // the writer through the decisions' revert values, so it is recorded here only.
  if (r.status === "ADDED") { if (yes) wrDrop.delete(r.project_num); else wrDrop.add(r.project_num); }
  for (const f of wrChanged(r.fields)) wrSet(r.project_num, f.key, yes);
  wrTouched[r.project_num] = yes ? "acc" : "skip";
  const recs = wrVisible();
  if (wrIdx < recs.length - 1) wrIdx++;
  renderWrSlides();
}

document.addEventListener("keydown", e => {
  if (wrView !== "slides" || !WR || !WR.ready) return;
  const page = document.querySelector('.tab-page[data-tab="wipreview"]');
  if (!page || page.hidden) return;
  if (/input|textarea|select/i.test((e.target && e.target.tagName) || "")) return;
  const recs = wrVisible(); if (!recs.length) return;
  if (e.key === "ArrowRight") { if (wrIdx < recs.length - 1) { wrIdx++; renderWrSlides(); } }
  else if (e.key === "ArrowLeft") { if (wrIdx > 0) { wrIdx--; renderWrSlides(); } }
  else if (e.key === "a" || e.key === "A") wrDecideJob(recs[wrIdx], true);
  else if (e.key === "s" || e.key === "S") wrDecideJob(recs[wrIdx], false);
  else return;
  e.preventDefault();
});

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
    head += `<button class="btn small wr-job-all" type="button">Approve job</button>`;
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
    if (wrBlockSel() && wrBlockSel() !== block) continue;   // the Show filter
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
  wrSaveChoices();
  let n = 0;
  for (const pn in wrDecisions) for (const k in wrDecisions[pn]) if (wrDecisions[pn][k]) n++;
  const btn = $("#wrSync");
  if (btn) btn.textContent = n ? `Sync ${n} approved →` : "Sync approved →";
}

function wrBulk(mode) {
  // mode: 'qbo' | 'all' | 'clear' - over the CURRENTLY VISIBLE jobs only.
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

async function runWipReview(asked) {
  if (asked) { /* the choice card already asked */ }
  else if (WR && WR.ready && !confirm("Recompute the pending WIP update?\n\nThis re-runs the WIP pipeline (CP folders, RP file, MFD) and pulls Billed/Costs from QuickBooks - expect a few Touch ID prompts and a few minutes. Nothing is written.")) return;
  else if (!(WR && WR.ready) && !confirm("Compute the pending WIP update?\n\nRuns the WIP pipeline and pulls Billed/Costs from QuickBooks (a few Touch ID prompts, a few minutes). Nothing is written - you review the changes first.")) return;
  const r = await fetch("/api/wip/review", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ confirm: true }) });
  if (r.status === 409) { alert("A sync is already running - let it finish first."); return; }
  if (!r.ok) { alert("Could not start the review."); return; }
  wrRunProgress("Computing the pending update", () => loadWipReview(true));
}

async function syncWipReview() {
  let n = 0; for (const pn in wrDecisions) for (const k in wrDecisions[pn]) if (wrDecisions[pn][k]) n++;
  const dropped = wrDrop.size;
  const stale = (WR && WR.stale) || [];
  if (stale.length && !confirm(`WARNING: the ${stale.join(" and ")} review is OLDER than the others (${stale.map(d => fmtDate(((WR.generated || {})[d] || {}).at || "", true)).join(", ")}).\n\nWriting it would put old numbers on the tab. Recompute first, or continue anyway?`)) return;
  if (!confirm(`Write approved changes to the WIP master?\n\n${n} approved change(s) will be written to Test - CP, Test - RP and Test-Master. Unchecked changes keep the current tab value${dropped ? `; ${dropped} added job(s) will be left off` : ""}.\n\nThis writes the production WIP workbook (guarded) and pulls QuickBooks again (Touch ID).`)) return;
  const decisions = wrBuildDecisions();
  const r = await fetch("/api/wip/merge", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ confirm: true, decisions }) });
  if (r.status === 409) { alert("A sync is already running - let it finish first."); return; }
  if (!r.ok) { alert("Could not start the sync."); return; }
  const started = wrLocalIso(new Date(Date.now() - 5000));
  wrRunProgress("Writing the approved changes", () => wrSynced(started));
}

// After a Sync, say what was WRITTEN - never re-show the review it was made from (owner 2026-09-23: "why when i
// update the wip it goes back and shows the page before i clicked sync? did my changes land?"). The review's
// before-values are the old tab now, so it is spent: the saved choices are cleared and the page offers Recompute.
const wrLocalIso = d => new Date(d.getTime() - d.getTimezoneOffset() * 60000).toISOString().slice(0, 19);
async function wrSynced(since) {
  const body = $("#wrBody"); if (!body) return;
  try { await load(true); } catch { /* the page still shows the summary */ }   // the ledger just reloaded the WIP - pull it into the page
  try { localStorage.removeItem(WR_LS); } catch { /* nothing saved */ }
  wrChosen = false;
  let st = {}, entries = [];
  try { st = await (await fetch("/api/sync/status")).json(); } catch { /* shown as unknown */ }
  try { entries = ((await (await fetch(`/api/wip/audit?since=${encodeURIComponent(since)}&limit=5000`)).json()).entries || []); } catch { entries = []; }
  const ok = st.state !== "error" && !(st.steps || []).some(x => x.state === "error");
  const moved = entries.filter(e => e.field !== "line" && e.old != null && e.actor === "sync");
  const lines = entries.filter(e => e.field === "line").length;
  const byTab = {}; for (const e of moved) byTab[e.tab] = (byTab[e.tab] || 0) + 1;
  const LBL = { costs: "Costs", billed: "Billed", retainage: "Retainage", contract: "Contract", approved_cos: "Approved COs", etc: "ETC", co_costs: "CO costs" };
  const byJob = {}; for (const e of moved.slice().reverse()) (byJob[e.project_no] = byJob[e.project_no] || []).push(e);
  const steps = (st.steps || []).map(x => `<div class="wr-step ${x.state}">${x.state === "done" ? "✓" : x.state === "error" ? "✕" : "·"} ${_ge(x.label)}</div>`).join("");
  const jobs = Object.keys(byJob).sort().map(pn => {
    const seen = new Set(), fs = byJob[pn].filter(e => !seen.has(e.field) && seen.add(e.field));   // a change logged on two tabs shows once
    return `<div class="wr-done-job"><b>${_ge(pn)}</b> ${fs.map(e => `${_ge(LBL[e.field] || e.field)} ${money(e.old)} → ${money(e.new)}`).join(" · ")}</div>`;
  }).join("");
  $("#wrFilters").hidden = true; $("#wrSync").hidden = true; $("#wrStats").innerHTML = "";
  body.innerHTML = `<div class="wr-done">
      <div class="wr-done-head ${ok ? "ok" : "bad"}"><span class="wr-done-icon">${ok ? "✓" : "✕"}</span>
        <div><div class="wr-done-title">${ok ? "WIP updated" : "Sync did not finish - see the steps"}</div>
        <div class="wr-done-when">${_ge(fmtDate(wrLocalIso(new Date()), true))}</div></div></div>
      <div class="wr-steps">${steps}</div>
      <div class="wr-done-sum">${moved.length} value${moved.length === 1 ? "" : "s"} changed${Object.keys(byTab).length ? " (" + Object.entries(byTab).map(([t, n]) => `${_ge(t)} ${n}`).join(" · ") + ")" : ""}${lines ? ` · ${lines} line${lines === 1 ? "" : "s"} added or removed` : ""}. Anything you left unchecked stayed as it was on the tab.</div>
      <div class="wr-done-jobs">${jobs || "<i>No value changed.</i>"}</div>
      <div class="wr-ch-actions"><button class="btn big primary" id="wrAfterRecompute" type="button">Recompute to see what is left</button></div>
    </div>`;
  $("#wrAfterRecompute").onclick = () => { wrChosen = true; runWipReview(true); };
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
    <div class="wr-now" id="wrNow"></div>
    <div class="hint" id="wrRunHint">Running - this can take a few minutes; Touch ID prompts appear on the Mac.</div></div>`;
  if (wrPoll) clearInterval(wrPoll);
  wrPoll = setInterval(async () => {
    let s; try { s = await (await fetch("/api/sync/status")).json(); } catch { return; }
    const steps = s.steps || [];
    const done = steps.filter(x => x.state === "done").length;
    const dt = s.detail, part = dt && dt.i && dt.n ? Math.min(dt.i / dt.n, 1) : 0;   // how far into the running step
    const fill = $("#wrFill"); if (fill) fill.style.width = steps.length ? Math.round((done + part) / steps.length * 100) + "%" : "0%";
    const now = $("#wrNow");   // one live line: the job and what is happening to it - replaced as it goes, gone when done
    if (now) now.innerHTML = dt && dt.project ? `<b>${_ge(dt.project)}</b> · ${_ge(dt.what || "")}${dt.i && dt.n ? ` <span class="wr-now-n">${dt.i} of ${dt.n}</span>` : ""}` : "";
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


// ── QBO Audit: the mirror's change log (owner 2026-09-23: "a QBO Audit section where I am watching the usual audit and
// this now with the deletions"). /api/qboaudit = every record QuickBooks deleted / edited / restored / added since the
// previous refresh, with QuickBooks' own time and plain-word flags (shared/qbo_mirror.change_flags); the mirror keeps
// the before record. Born of 2026-09-18, when the audit log blamed the owner for deletions a connected app made.
let QA = null, qaDays = 30, qaFlag = null, qaShowOk = false;
const QA_ENTITY = { Bill: "Bill", BillPayment: "Bill payment", Purchase: "Expense", Invoice: "Invoice", Payment: "Payment received", JournalEntry: "Journal entry",
  VendorCredit: "Vendor credit", CreditMemo: "Credit memo", Deposit: "Deposit", Transfer: "Transfer", PurchaseOrder: "Purchase order", Estimate: "Estimate",
  Customer: "Customer", Vendor: "Vendor", Account: "Account", Item: "Item", Class: "Class", Term: "Term", PaymentMethod: "Payment method" };
const QA_QBO_KIND = { Bill: "bill", BillPayment: "billpayment", Purchase: "expense", Invoice: "invoice", Payment: "recvpayment", JournalEntry: "journal", VendorCredit: "vendorcredit", CreditMemo: "creditmemo", Deposit: "deposit", Transfer: "transfer", PurchaseOrder: "purchaseorder", Estimate: "estimate" };
const qaFlagCls = f => /deleted|unapplied|reopened|voided/.test(f) ? "neg" : (/restored/.test(f) ? "ok" : "warn");
async function loadQboAudit(force) {
  const note = $("#qaNote"), table = $("#qaTable"); if (!table) return;
  if (QA && QA.ok && QA.days === qaDays && !force) { renderQboAudit(); return; }
  if (note) note.textContent = "loading…";
  skeletonInto(table.tBodies[0] || table, 6);
  try { QA = await (await fetch(`/api/qboaudit?days=${qaDays}`)).json(); } catch (e) { QA = { ok: false, error: String(e) }; }
  renderQboAudit();
}
async function qaSetOk(c, ok) {   // a ledger write (qbo_change_ok) - never QuickBooks; the row leaves the list in place, no scroll
  try {
    const r = await (await fetch("/api/qboaudit/ok", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ ids: [c.id], ok }) })).json();
    if (!r.ok) throw new Error(r.error || "failed");
    c.ok_at = ok ? r.marked_at : null;
    for (const k of (QA.changes || [])) if (k.parent_id === c.id) k.ok_at = c.ok_at;
    if (ok && c.flags) for (const f of c.flags) if (QA.flag_counts && QA.flag_counts[f]) QA.flag_counts[f] -= 1;
    if (!ok && c.flags) for (const f of c.flags) if (QA.flag_counts) QA.flag_counts[f] = (QA.flag_counts[f] || 0) + 1;
    const y = window.scrollY; renderQboAudit(); window.scrollTo(0, y);
  } catch (e) { toast ? toast("Could not save: " + e.message) : alert("Could not save: " + e.message); }
}
function renderQboAudit() {
  const note = $("#qaNote"), stats = $("#qaStats"), filt = $("#qaFilters"), table = $("#qaTable"); if (!table) return;
  const thead = table.querySelector("thead"), tbody = table.querySelector("tbody");
  $$("#qaDays .seg-btn").forEach(b => b.classList.toggle("on", Number(b.dataset.days) === qaDays));
  if (!QA || !QA.ok) {
    if (note) note.textContent = QA && QA.error ? "unavailable" : "";
    stats.innerHTML = ""; filt.innerHTML = ""; thead.innerHTML = "";
    tbody.innerHTML = QA && QA.error ? `<tr><td class="left" style="padding:14px;color:var(--text-dim)">${_ge(QA.error)}</td></tr>` : "";
    return;
  }
  const everything = QA.changes || [], fc = QA.flag_counts || {};
  const kids = {}; for (const c of everything) if (c.parent_id) (kids[c.parent_id] = kids[c.parent_id] || []).push(c);   // bills a check lost, under the check
  const okd = everything.filter(c => !c.parent_id && c.ok_at);   // the owner's "that's OK" - off the list unless asked for
  const all = everything.filter(c => !c.parent_id && (qaShowOk || !c.ok_at)), flagged = all.filter(c => c.flags.length && !c.ok_at);
  if (note) note.textContent = `${all.length} change${all.length === 1 ? "" : "s"} in ${qaDays} days · ${flagged.length} flagged · mirror refreshed ${QA.last_refresh ? fmtDate(QA.last_refresh, true) : "never"}`;
  stats.innerHTML = "";
  const tile = (label, val, sub, bad) => { const k = document.createElement("div"); k.className = "kpi" + (bad ? " kpi-neg" : ""); k.innerHTML = `<div class="k-label"></div><div class="k-value"></div><div class="k-sub"></div>`;
    k.querySelector(".k-label").textContent = label; k.querySelector(".k-value").textContent = val; k.querySelector(".k-sub").textContent = sub || ""; stats.appendChild(k); };
  const dels = all.filter(c => c.kind === "deleted" && !c.parent_id);
  tile("Deleted", dels.length, dels.length ? money(dels.reduce((s, c) => s + num(c.total_before), 0)) : "nothing deleted", dels.length > 0);
  tile("Paid bills deleted", fc["deleted paid bill"] || 0, "", !!fc["deleted paid bill"]);
  tile("Checks that lost bills", fc["payment unapplied"] || 0, "fix on Checks QBO changed", !!fc["payment unapplied"]);
  tile("Voided", fc["voided"] || 0, "", false);
  filt.innerHTML = "";
  const chip = (label, key, n) => { const b = document.createElement("button"); b.className = "acct-chip" + (qaFlag === key ? " active" : ""); b.innerHTML = `${_ge(label)} <span class="ac-n">${n}</span>`;
    b.onclick = () => { qaFlag = qaFlag === key ? null : key; renderQboAudit(); }; filt.appendChild(b); };
  chip("All flagged", null, flagged.length);
  for (const f of Object.keys(fc).sort((a, b) => fc[b] - fc[a])) chip(f, f, fc[f]);
  if (okd.length) { const b = document.createElement("button"); b.className = "acct-chip" + (qaShowOk ? " active" : ""); b.title = "Changes you marked OK - hidden until you show them";
    b.innerHTML = `Show OK'd <span class="ac-n">${okd.length}</span>`; b.onclick = () => { qaShowOk = !qaShowOk; const y = window.scrollY; renderQboAudit(); window.scrollTo(0, y); }; filt.appendChild(b); }
  const q = (($("#qaSearch") || {}).value || "").trim().toLowerCase();
  const flaggedOnly = !!($("#qaFlaggedOnly") && $("#qaFlaggedOnly").checked);
  const rows = all.filter(c => (!flaggedOnly || c.flags.length) && (!qaFlag || c.flags.includes(qaFlag))
    && (!q || [c.ref_name, c.doc_number, QA_ENTITY[c.entity], c.entity, c.rec_id, c.kind, c.flags.join(" "), c.total_before, c.total_after, c.txn_date].join(" ").toLowerCase().includes(q)));
  const cols = [["When (QuickBooks time)", "left"], ["What", "left"], ["Type", "left"], ["No.", "left"], ["Vendor / client", "left"], ["Txn date", "left"], ["Before", "right"], ["After", "right"], ["Open now", "right"], ["Record", "left"], ["", "left"]];
  thead.innerHTML = "<tr>" + cols.map(([c, al]) => `<th class="${al}">${_ge(c)}</th>`).join("") + "</tr>";
  tbody.innerHTML = "";
  if (!rows.length) { tbody.innerHTML = `<tr><td colspan="${cols.length}" class="left" style="padding:14px;color:var(--text-dim)">${all.length ? "Nothing matches." : "No changes in this window - refresh the mirror to check again."}</td></tr>`; return; }
  const frag = document.createDocumentFragment();
  const qaRow = (c, child) => {
    const tr = document.createElement("tr");
    if (child) tr.className = "qa-child";
    else if (kids[c.id]) { tr.className = "qa-pay"; tr.dataset.grpkey = c.rec_id + "|" + c.id; }
    tr.appendChild(leftText(c.changed_at ? fmtDate(c.changed_at, true) : "–"));
    { const td = document.createElement("td"); td.className = "left qa-what";
      if (!c.flags.length) { const s = document.createElement("span"); s.className = "st st-dim"; s.textContent = c.kind; td.appendChild(s); }
      for (const f of c.flags) { const s = document.createElement("span"); s.className = "qa-flag " + qaFlagCls(f); s.textContent = f; td.appendChild(s); }
      const nk = (kids[c.id] || []).length;
      const fixCheck = c.entity === "BillPayment" && (c.repair || c.flags.includes("payment unapplied")) ? c.doc_number : (c.knocked_out && c.knocked_out.check);
      if (nk || fixCheck) {                          // one short line under the flag: how many bills it reopened + the fix (on ONE page: Checks QBO changed)
        const d = document.createElement("div"); d.className = "qa-sub";
        if (nk) d.append(`${nk} bill${nk === 1 ? "" : "s"} reopened`);
        if (fixCheck && typeof loadCheckDrift === "function") {
          if (nk) d.append(" · ");
          const a = document.createElement("a"); a.href = "#"; a.textContent = c.entity === "BillPayment" ? "fix ↗" : `fix check #${fixCheck} ↗`;
          a.onclick = (e) => { e.preventDefault(); e.stopPropagation(); const q = $("#cdSearch"); if (q) q.value = String(fixCheck); cdFilter = "all"; setTab("checkdrift"); };
          d.appendChild(a); }
        td.appendChild(d); }
      tr.appendChild(td); }
    tr.appendChild(leftText(QA_ENTITY[c.entity] || c.entity));
    { const url = c.deleted_now || !QA_QBO_KIND[c.entity] ? null : qboUrl(QA_QBO_KIND[c.entity], c.rec_id);
      const td = qboLinkCell(c.doc_number || (c.entity === "Purchase" ? "(expense)" : "–"), url, "Open in QuickBooks"); if (!c.doc_number) td.classList.add("dim"); tr.appendChild(td); }
    { const td = document.createElement("td"); td.className = "left";
      if (c.ref_name && c.ref_type === "vendor" && typeof openVendorPage === "function") { const a = document.createElement("a"); a.href = "#"; a.textContent = c.ref_name; a.title = "Open the vendor page"; a.onclick = (e) => { e.preventDefault(); openVendorPage(c.ref_name); }; td.appendChild(a); }
      else if (c.ref_name && c.ref_type === "customer" && typeof openClientPage === "function") { const a = document.createElement("a"); a.href = "#"; a.textContent = c.ref_name; a.title = "Open the client page"; a.onclick = (e) => { e.preventDefault(); openClientPage(c.ref_name); }; td.appendChild(a); }
      else { td.textContent = c.ref_name || "–"; if (!c.ref_name) td.classList.add("dim"); }
      tr.appendChild(td); }
    tr.appendChild(leftText(c.txn_date ? fmtDateShort(c.txn_date) : "–"));
    // before / after = what the flag is about: the money applied for a payment, the open balance for a reopen, else the total
    let vb = c.total_before, va = c.total_after, what = "total";
    if (c.flags.includes("ck # changed") && c.flags.length === 1 && c.doc_before != null) {   // the NUMBER changed, not the money: show the numbers
      for (const v of [c.doc_before, c.doc_number]) { const td = document.createElement("td"); td.className = "right"; td.title = "check #"; td.textContent = v ? "#" + v : "(none)"; if (!v) td.classList.add("dim"); tr.appendChild(td); }
      vb = va = undefined; what = null;
    }
    else if (c.flags.includes("payment unapplied")) { vb = c.applied_before; va = c.applied_after; what = "applied to bills"; }
    else if (c.flags.includes("reopened")) { vb = c.balance_before; va = c.balance_after; what = "open balance"; }
    if (what !== null) for (const v of [vb, va]) { const td = document.createElement("td"); td.className = "right"; td.title = what; if (v == null) { td.textContent = "–"; td.classList.add("dim"); } else td.appendChild(moneyCell(v)); tr.appendChild(td); }
    { const td = document.createElement("td"); td.className = "right"; const v = c.kind === "deleted" ? null : c.balance_after; if (v == null) { td.textContent = "–"; td.classList.add("dim"); } else td.textContent = money(v); tr.appendChild(td); }
    { const td = document.createElement("td"); td.className = "left"; const s = document.createElement("span"); s.className = "st " + (c.deleted_now ? "st-bad" : "st-dim");
      s.textContent = c.deleted_now ? "deleted in QuickBooks" : "on file"; if (c.has_before) s.title = "The record as it was before this change is kept in the mirror (the repair material)"; td.appendChild(s); tr.appendChild(td); }
    { const td = document.createElement("td"); td.className = "left qa-okcell";   // the owner's review: OK = seen, fine, off the list next run
      if (!child && c.ok_at) {
        const s = document.createElement("span"); s.className = "st st-dim"; s.textContent = `OK'd ${fmtDateShort(c.ok_at)}`; td.appendChild(s);
        const u = document.createElement("a"); u.href = "#"; u.className = "qa-undo"; u.textContent = "undo"; u.onclick = (e) => { e.preventDefault(); e.stopPropagation(); qaSetOk(c, false); }; td.appendChild(u);
      } else if (!child && c.flags.length) {
        const b = document.createElement("button"); b.className = "btn small qa-ok"; b.textContent = "OK"; b.title = "Seen it, it's fine - hide it from the next run";
        b.onclick = (e) => { e.stopPropagation(); qaSetOk(c, true); }; td.appendChild(b);
      }
      tr.appendChild(td); }
    return tr;
  };
  for (const c of rows.slice(0, 600)) {
    frag.appendChild(qaRow(c, false));
    for (const k of kids[c.id] || []) frag.appendChild(qaRow(k, true));
  }
  tbody.appendChild(frag);
  if (rows.length > 600) { const tr = document.createElement("tr"); const td = document.createElement("td"); td.colSpan = cols.length; td.className = "left dim"; td.style.padding = "10px 14px"; td.textContent = `${rows.length - 600} more - narrow the search`; tr.appendChild(td); tbody.appendChild(tr); }
}

// cents: these are the amounts typed back into the check, to the penny
const qaCents = v => v == null || v === "" || Number.isNaN(Number(v)) ? "–" : "$" + Number(v).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });

// ── Checks QBO changed (owner 2026-09-24: "qbo freaks out when a bill paid gets changed and does automatic changes
// without ever consenting or warning ... subs it happens to the most"). /api/checkdrift = ledger/check_drift.py over
// the mirror, every year on file: a check whose money floats (applied to nothing), the bill it paid now open again
// (edited: 47436 / UC015) or its re-entered copy open (deleted: 48314 / UC41), the freed loan credit, and floating
// money QuickBooks dropped onto a LATER bill (48314 -> UC044 $135.40). One bill belongs to one check. Read-only.
let CD = null, cdFilter = "subs";
const CD_PAT_CLS = { "paid bill deleted": "neg", "paid bill edited": "neg", "check rewritten": "neg", "put on a later bill": "warn", "credit dropped": "warn" };
const CD_FILTERS = [["subs", "Subs", r => r.sub], ["all", "Everyone", () => true],
  ["floating", "Money floating", r => r.floating > 1], ["owed", "Bill reads as owed again", r => r.reopened.length || r.copies.length],
  ["late", "Put on a later bill", r => r.late.length]];
async function loadCheckDrift(force) {
  const note = $("#cdNote"), table = $("#cdTable"); if (!table) return;
  if (CD && CD.ok && !force) { renderCheckDrift(); return; }
  if (note) note.textContent = "loading…";
  skeletonInto(table.tBodies[0] || table, 6);
  try { CD = await (await fetch("/api/checkdrift")).json(); } catch (e) { CD = { ok: false, error: String(e) }; }
  renderCheckDrift();
}
function renderCheckDrift() {
  const note = $("#cdNote"), stats = $("#cdStats"), filt = $("#cdFilters"), table = $("#cdTable"); if (!table) return;
  const thead = table.querySelector("thead"), tbody = table.querySelector("tbody");
  if (!CD || !CD.ok) {
    if (note) note.textContent = CD && CD.error ? "unavailable" : "";
    stats.innerHTML = ""; filt.innerHTML = ""; thead.innerHTML = "";
    tbody.innerHTML = CD && CD.error ? `<tr><td class="left" style="padding:14px;color:var(--text-dim)">${_ge(CD.error)}</td></tr>` : "";
    return;
  }
  const all = CD.rows || [], sm = CD.summary || {};
  if (note) note.textContent = `${all.length} check${all.length === 1 ? "" : "s"} · ${sm.subs || 0} subs · every year on file · mirror refreshed ${CD.last_refresh ? fmtDate(CD.last_refresh, true) : "never"}`;
  stats.innerHTML = "";
  const stat = (label, val, sub, bad) => { const k = document.createElement("div"); k.className = "kpi" + (bad ? " kpi-neg" : ""); k.innerHTML = `<div class="k-label"></div><div class="k-value"></div><div class="k-sub"></div>`;
    k.querySelector(".k-label").textContent = label; k.querySelector(".k-value").textContent = val; k.querySelector(".k-sub").textContent = sub || ""; stats.appendChild(k); };
  stat("Money floating", money(sm.floating || 0), `${sm.floating_n || 0} checks · subs ${money(sm.floating_subs || 0)}`, (sm.floating || 0) > 0);
  stat("Bills owed again", money((sm.reopened || 0) + (sm.copies || 0)), `${(sm.reopened_n || 0) + (sm.copies_n || 0)} bills - don't pay twice`, ((sm.reopened_n || 0) + (sm.copies_n || 0)) > 0);
  stat("Put on a later bill", money(sm.late || 0), `${sm.late_n || 0} checks - that week went out short`, (sm.late_n || 0) > 0);
  filt.innerHTML = "";
  for (const [key, label, fn] of CD_FILTERS) {
    const b = document.createElement("button"); b.className = "acct-chip" + (cdFilter === key ? " active" : "");
    b.innerHTML = `${_ge(label)} <span class="ac-n">${all.filter(fn).length}</span>`;
    b.onclick = () => { cdFilter = key; const y = window.scrollY; renderCheckDrift(); window.scrollTo(0, y); }; filt.appendChild(b);
  }
  const fn0 = (CD_FILTERS.find(f => f[0] === cdFilter) || CD_FILTERS[0])[2];
  const fn = fn0;
  const q = (($("#cdSearch") || {}).value || "").trim().toLowerCase().split(/\s+/).filter(Boolean);
  const hay = r => [r.vendor, r.check, r.txn_date, r.total, r.floating, r.pattern, ...r.reopened.map(b => b.doc_number + " " + b.balance),
    ...r.copies.map(b => b.doc_number + " " + b.balance), ...r.late.map(b => b.doc_number + " " + b.amount), ...r.cause, ...r.todo].join(" ").toLowerCase();
  const base = all.filter(r => fn(r) && (!q.length || q.every(w => hay(r).includes(w))));
  const rows = base.filter(r => hfPasses("checkdrift", r));
  const cols = [["Vendor", "left", "vendor"], ["Check #", "left"], ["Check date", "left"], ["Check total", "right"], ["Applied now", "right"],
    ["Floating", "right"], ["Bill it paid (open again)", "left"], ["Put on a later bill", "left"], ["What happened", "left"], ["Changed", "left"]];
  thead.innerHTML = "";
  { const tr = document.createElement("tr");
    for (const [c, al, hk] of cols) { const th = document.createElement("th"); th.className = al; th.textContent = c;
      if (hk) hfDecorate(th, "checkdrift", hk, () => base.filter(r => hfPasses("checkdrift", r, hk)), renderCheckDrift); tr.appendChild(th); }
    thead.appendChild(tr); }
  tbody.innerHTML = "";
  if (!rows.length) { tbody.innerHTML = `<tr><td colspan="${cols.length}" class="left" style="padding:14px;color:var(--text-dim)">${all.length ? "Nothing matches." : "No check QuickBooks changed - every paid check still sits on its bills."}</td></tr>`; return; }
  const frag = document.createDocumentFragment();
  const billTxt = b => `#${b.doc_number} ${fmtDateShort(b.txn_date)} ${qaCents(b.balance)}` + (b.exact ? "" : " (possible)");
  for (const r of rows) {
    const tr = document.createElement("tr"); tr.className = "vp-pay";
    { const td = document.createElement("td"); td.className = "left"; const sp = document.createElement("span");
      if (typeof openVendorPage === "function") { const a = document.createElement("a"); a.href = "#"; a.textContent = r.vendor || "–"; a.title = "Open the vendor page"; a.onclick = (e) => { e.preventDefault(); e.stopPropagation(); openVendorPage(r.vendor); }; sp.appendChild(a); }
      else sp.textContent = r.vendor || "–";
      td.appendChild(sp);
      if (r.sub) { const s = document.createElement("span"); s.className = "st st-dim"; s.style.marginLeft = "6px"; s.textContent = "sub"; td.appendChild(s); }
      tr.appendChild(td); }
    tr.appendChild(qboLinkCell(r.check || "–", qboUrl("billpayment", r.payment_id), "Open the check in QuickBooks"));
    tr.appendChild(leftText(r.txn_date ? fmtDateShort(r.txn_date) : "–"));
    for (const v of [r.total, r.applied]) { const td = document.createElement("td"); td.className = "right"; td.textContent = qaCents(v); tr.appendChild(td); }
    { const td = document.createElement("td"); td.className = "right"; td.textContent = r.floating > 1 ? qaCents(r.floating) : "–"; if (r.floating > 1) td.classList.add("neg"); else td.classList.add("dim"); tr.appendChild(td); }
    { const owed = [...r.reopened, ...r.copies];
      const td = leftText(!owed.length ? "–" : owed.length > 3 ? `${owed.length} bills · ${qaCents(owed.reduce((t, b) => t + num(b.amount || b.balance), 0))}` : owed.map(billTxt).join(" · "));
      if (!owed.length) td.classList.add("dim");
      tr.appendChild(td); }
    { const td = leftText(r.late.length ? r.late.map(b => `#${b.doc_number} ${qaCents(b.amount)}`).join(" · ") : "–"); if (!r.late.length) td.classList.add("dim"); tr.appendChild(td); }
    { const td = document.createElement("td"); td.className = "left"; const s = document.createElement("span"); s.className = "qa-flag " + (CD_PAT_CLS[r.pattern] || "st-dim"); s.textContent = r.pattern; td.appendChild(s); tr.appendChild(td); }
    tr.appendChild(leftText(r.changed_local ? fmtDate(r.changed_local, true) : "–"));   // Central, not QBO's Pacific stamp
    frag.appendChild(tr);
    frag.appendChild(cdFixRow(r, cols.length));
  }
  tbody.appendChild(frag);
}
// the fix under a check (a vp-pay group: collapsed by default, the caret opens it)
function cdFixRow(r, span) {
  const tr = document.createElement("tr"); tr.className = "qa-repair";
  const td = document.createElement("td"); td.colSpan = span; td.className = "left";
  const h = document.createElement("div"); h.className = "qa-repair-head";
  h.textContent = `Check #${r.check}: ${qaCents(r.total)}, ${qaCents(r.applied)} applied now` + (r.cause.length ? ` · ${r.cause.join(" · ")}` : "");
  td.appendChild(h);
  const ol = document.createElement("ol"); ol.className = "cd-fix";
  for (const step of r.todo) { const li = document.createElement("li"); li.textContent = step; ol.appendChild(li); }
  td.appendChild(ol);
  const t = document.createElement("table"); t.className = "qa-repair-tbl";
  t.innerHTML = `<thead><tr><th class="left">What</th><th class="left">No.</th><th class="left">Date</th><th class="right">Amount</th><th class="left"></th></tr></thead>`;
  const tb = document.createElement("tbody");
  const line = (what, no, url, date, amt, note, cls) => { const x = document.createElement("tr");
    x.appendChild(leftText(what)); x.appendChild(qboLinkCell(no || "–", url, "Open in QuickBooks")); x.appendChild(leftText(date ? fmtDateShort(date) : "–"));
    { const c = document.createElement("td"); c.className = "right"; c.textContent = qaCents(amt); x.appendChild(c); }
    { const c = document.createElement("td"); c.className = "left"; if (note) { const s = document.createElement("span"); s.className = "qa-flag " + (cls || "warn"); s.textContent = note; c.appendChild(s); } x.appendChild(c); }
    tb.appendChild(x); };
  for (const a of r.applied_lines) line(a.type === "VendorCredit" ? "Applied: credit" : "Applied now", a.doc_number || a.id, a.type === "Bill" ? qboUrl("bill", a.id) : (a.type === "VendorCredit" ? qboUrl("vendorcredit", a.id) : null), a.txn_date, a.amount, "", "");
  for (const b of r.reopened) line("Paid bill, open again", b.doc_number, qboUrl("bill", b.id), b.txn_date, b.amount || b.balance, b.exact ? "re-apply" : "possible - check it is the one", b.exact ? "neg" : "warn");
  for (const b of r.copies) line("Re-entered copy, open", b.doc_number, qboUrl("bill", b.id), b.txn_date, b.balance, b.exact ? "don't pay - re-apply" : "possible - check it is the one", b.exact ? "neg" : "warn");
  for (const c of r.freed) line("Credit freed", c.doc_number || c.id, qboUrl("vendorcredit", c.id), c.txn_date, c.balance, "put back on the check", "warn");
  for (const b of r.late) line("Put on a later bill", b.doc_number, qboUrl("bill", b.id), b.txn_date, b.amount, "take off - that week went out short", "neg");
  t.appendChild(tb); td.appendChild(t);
  tr.appendChild(td); return tr;
}

// ── Uncleared checks (owner 2026-09-24: "a list of all unmatched checks and/or any checks that haven't been deposited").
// /api/uncleared = ledger/load_uncleared_checks.py: QuickBooks' TransactionList `cleared=Uncleared` filter (the column
// is hidden, the filter works). In QuickBooks "not matched" and "not deposited" are the same flag. Each account says
// how far its bank-feed matching has got; Joint Checks is never feed-matched and stays apart. Read-only.
let UC = null, ucFilter = "old";
const UC_FILTERS = [["old", "Older than 30 days", c => c.feed_matched && c.days > 30], ["bank", "All bank checks", c => c.feed_matched],
  ["recent", "Last 30 days", c => c.feed_matched && c.days <= 30], ["other", "Never feed-matched (Joint Checks …)", c => !c.feed_matched]];
async function loadUncleared(force) {
  const note = $("#ucNote"), table = $("#ucTable"); if (!table) return;
  if (UC && UC.ok && !force) { renderUncleared(); return; }
  if (note) note.textContent = "loading…";
  skeletonInto(table.tBodies[0] || table, 6);
  try { UC = await (await fetch("/api/uncleared")).json(); } catch (e) { UC = { ok: false, error: String(e) }; }
  renderUncleared();
}
function renderUncleared() {
  const note = $("#ucNote"), stats = $("#ucStats"), filt = $("#ucFilters"), table = $("#ucTable"); if (!table) return;
  const thead = table.querySelector("thead"), tbody = table.querySelector("tbody");
  if (!UC || !UC.ok) {
    if (note) note.textContent = ""; stats.innerHTML = ""; filt.innerHTML = ""; thead.innerHTML = "";
    tbody.innerHTML = `<tr><td class="left" style="padding:14px;color:var(--text-dim)">${_ge((UC && UC.error) || "")}</td></tr>`;
    return;
  }
  const all = UC.checks || [];
  if (note) note.textContent = `pulled ${UC.loaded_at ? fmtDate(UC.loaded_at, true) : "never"}`;
  stats.innerHTML = "";
  const stat = (label, val, sub, bad) => { const k = document.createElement("div"); k.className = "kpi" + (bad ? " kpi-neg" : ""); k.innerHTML = `<div class="k-label"></div><div class="k-value"></div><div class="k-sub"></div>`;
    k.querySelector(".k-label").textContent = label; k.querySelector(".k-value").textContent = val; k.querySelector(".k-sub").textContent = sub || ""; stats.appendChild(k); };
  const sum = xs => xs.reduce((t, c) => t + num(c.amount), 0);
  const old = all.filter(UC_FILTERS[0][2]), bank = all.filter(UC_FILTERS[1][2]), other = all.filter(UC_FILTERS[3][2]);
  stat("Older than 30 days", money(sum(old)), `${old.length} checks - not cashed, or cashed and never matched`, old.length > 0);
  stat("All bank checks", money(sum(bank)), `${bank.length} checks`, false);
  for (const a of UC.accounts || []) if (a.feed_matched) stat(a.account.replace(/\*+/g, " "), a.matched_through ? fmtDateShort(a.matched_through) : "–", "matched through", false);
  stat("Never feed-matched", money(sum(other)), `${other.length} checks · Joint Checks and others`, false);
  filt.innerHTML = "";
  for (const [key, label, fn] of UC_FILTERS) {
    const b = document.createElement("button"); b.className = "acct-chip" + (ucFilter === key ? " active" : "");
    b.innerHTML = `${_ge(label)} <span class="ac-n">${all.filter(fn).length}</span>`;
    b.onclick = () => { ucFilter = key; const y = window.scrollY; renderUncleared(); window.scrollTo(0, y); }; filt.appendChild(b);
  }
  const fn = (UC_FILTERS.find(f => f[0] === ucFilter) || UC_FILTERS[0])[2];
  const q = (($("#ucSearch") || {}).value || "").trim().toLowerCase().split(/\s+/).filter(Boolean);
  const base = all.filter(c => fn(c) && (!q.length || q.every(w => [c.payee, c.check_no, c.account, c.amount, c.txn_date, fmtDateShort(c.txn_date)].join(" ").toLowerCase().includes(w))));
  const rows = base.filter(c => hfPasses("uncleared", c));
  const cols = [["Check #", "left"], ["Date", "left"], ["Payee", "left", "vendor"], ["Bank account", "left"], ["Amount", "right"], ["Days out", "right"], ["Type", "left"]];
  thead.innerHTML = "";
  { const tr = document.createElement("tr");
    for (const [c, al, hk] of cols) { const th = document.createElement("th"); th.className = al; th.textContent = c;
      if (hk) hfDecorate(th, "uncleared", hk, () => base.filter(r => hfPasses("uncleared", r, hk)), renderUncleared); tr.appendChild(th); }
    thead.appendChild(tr); }
  tbody.innerHTML = "";
  if (!rows.length) { tbody.innerHTML = `<tr><td colspan="${cols.length}" class="left" style="padding:14px;color:var(--text-dim)">${all.length ? "Nothing matches." : "No uncleared checks."}</td></tr>`; return; }
  const frag = document.createDocumentFragment();
  for (const c of rows.slice(0, 1500)) {
    const tr = document.createElement("tr");
    tr.appendChild(qboLinkCell(c.check_no || "–", qboUrl(c.txn_type === "Check" ? "check" : "billpayment", c.qbo_txn_id), "Open in QuickBooks"));
    tr.appendChild(leftText(c.txn_date ? fmtDateShort(c.txn_date) : "–"));
    { const td = document.createElement("td"); td.className = "left";
      if (c.payee && typeof openVendorPage === "function") { const a = document.createElement("a"); a.href = "#"; a.textContent = c.payee; a.title = "Open the vendor page"; a.onclick = (e) => { e.preventDefault(); openVendorPage(c.payee); }; td.appendChild(a); }
      else td.textContent = c.payee || "–";
      tr.appendChild(td); }
    tr.appendChild(leftText(c.account || "–"));
    { const td = document.createElement("td"); td.className = "right"; td.appendChild(moneyCell(c.amount)); tr.appendChild(td); }
    { const td = document.createElement("td"); td.className = "right" + (c.feed_matched && c.days > 30 ? " neg" : ""); td.textContent = c.days == null ? "–" : String(c.days); tr.appendChild(td); }
    tr.appendChild(leftText(c.txn_type === "Check" ? "Check" : "Bill payment"));
    frag.appendChild(tr);
  }
  tbody.appendChild(frag);
}

// ── Accounting fixes: the Bill Tracker audits, filterable by audit type ───────
let ACCT = null;            // cached /api/accounting payload
let acctIssue = null;       // the audit-type filter currently active (null = all)
// Four groups by WHO fixes it (owner 2026-09-23: "audit needs a ton of work and simplification") - the group is the
// main filter (one bold line, counts); inside it, one collapsible row per issue with its bills under it.
const ACCT_GROUPS = [
  { id: "coding",   label: "Coding",                who: "fix in QuickBooks - the bill clerk" },
  { id: "approval", label: "Approval",              who: "not approved - PMs / owner" },
  { id: "po",       label: "Purchase orders",       who: "purchasing" },
  { id: "tracker",  label: "Tracker vs QuickBooks", who: "on the Bill Tracker, not in QuickBooks" },
];
const acctGroupOf = f => /not in qbo/i.test(f.issue || "") ? "tracker" : /not approved/i.test(f.issue || "") ? "approval" : f.group === "PO" ? "po" : "coding";
let acctGroup = "coding";
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
  const parts = [(ACCT_GROUPS.find(g => g.id === acctGroup) || {}).label];
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
  "Date": f => f.date || "", "Project": f => f.project, "Class": f => f.qbo_class || f.division,
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
  if (note) note.textContent = `${all.length} to fix · Bill Tracker audits`;
  // the four groups: the main filter, pronounced (owner 2026-09-23: "make the filter portion more pronounced so i know what im seeing")
  if (stats) stats.innerHTML = "";
  const gbox = $("#acctGroups"); gbox.innerHTML = "";
  const gCount = {}; for (const f of all) { const g = acctGroupOf(f); gCount[g] = (gCount[g] || 0) + 1; }
  for (const g of ACCT_GROUPS) {
    const b = document.createElement("button"); b.type = "button"; b.className = "acct-group" + (acctGroup === g.id ? " on" : "");
    b.innerHTML = `<span class="ag-n">${gCount[g.id] || 0}</span><span class="ag-l">${_ge(g.label)}</span><span class="ag-w">${_ge(g.who)}</span>`;
    b.onclick = () => { acctGroup = g.id; acctIssue = null; renderAccounting(); };
    gbox.appendChild(b);
  }
  const inGroup = all.filter(f => acctGroupOf(f) === acctGroup);
  // the issues inside the group, as small chips (only when there is more than one)
  filt.innerHTML = "";
  const iCount = {}; for (const f of inGroup) iCount[f.issue] = (iCount[f.issue] || 0) + 1;
  const issues = Object.keys(iCount).sort((a, b) => iCount[b] - iCount[a]);
  if (acctIssue && !iCount[acctIssue]) acctIssue = null;
  if (issues.length > 1) {
    const chip = (label, key, n) => {
      const b = document.createElement("button"); b.className = "acct-chip" + (acctIssue === key ? " active" : "");
      b.innerHTML = `${_ge(label)} <span class="ac-n">${n}</span>`;
      b.onclick = () => { acctIssue = acctIssue === key ? null : key; renderAccounting(); };
      filt.appendChild(b);
    };
    chip("All " + (ACCT_GROUPS.find(g => g.id === acctGroup) || {}).label.toLowerCase(), null, inGroup.length);
    for (const iss of issues) chip(iss, iss, iCount[iss]);
  }
  // vendor checkbox filter (same multi-select as the other tabs)
  { const sig = String(all.length); if (sig !== _acctVendorSig || !($("#acctVendorMenu") && $("#acctVendorMenu").querySelector(".msel-opt"))) { _acctVendorSig = sig; buildMSel(ACCT_VENDOR_MSEL, all, acctMSel, renderAccounting); } }
  // division filter
  const dsel = $("#acctDivision");
  if (dsel && dsel.options.length <= 1) for (const d of [...new Set(all.map(f => f.division).filter(Boolean))].sort()) { const o = document.createElement("option"); o.value = d; o.textContent = d; dsel.appendChild(o); }
  const dv = dsel ? dsel.value : "", q = ($("#acctSearch").value || "").trim().toLowerCase();
  const rows = inGroup.filter(f => (!acctIssue || f.issue === acctIssue) && (!dv || f.division === dv) && mselPasses(f, [ACCT_VENDOR_MSEL], acctMSel)
    && (!q || (f.vendor + " " + f.project + " " + f.bill_no + " " + (f.memo || "") + " " + f.detail).toLowerCase().includes(q)));
  _setHintFilter("accounting", acctFilterDesc(rows.length, all.length));   // count + what's filtered (generic when All)
  if (acctSort.length) rows.sort(_acctCmp);   // multi-column sort (applied before the render cap)
  _acctVisible = rows;
  // fixed meta widths (px) so one long outlier can't blow a column wide (the old wasted
  // space); the two text columns (null width) share the rest and wrap - nothing truncates.
  const cols = [["Vendor", "left audit-soft", 160], ["Bill #", "left", 118],
    ["📎", "left", 44], ["Date", "left", 112], ["Project", "left", 140], ["Class (QuickBooks)", "left", 124], ["Cost", "left", 64], ["Amount", "right", 92],
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
  const ACCT_CAP = 150;   // render cap PER issue - all ~1900 rows (each w/ a checkbox + scan button) crashed the tab
  const frag = document.createDocumentFragment();
  // one collapsible band per issue (GRP_KINDS tr.bill-group), its bills under it - the issue is said ONCE, not squeezed into every row
  // every issue gets its band; the render cap applies PER issue (a total cap hid whole issues past the first 250 rows)
  const byIssue = new Map();
  for (const f of rows) { if (!byIssue.has(f.issue)) byIssue.set(f.issue, []); byIssue.get(f.issue).push(f); }
  const allByIssue = {}; for (const [iss, l] of byIssue) allByIssue[iss] = l.length;
  const bands = [...byIssue.entries()].sort((a, b) => b[1].length - a[1].length);
  for (const [iss, full] of bands) {
  const list = full.slice(0, ACCT_CAP);
  { const hr = document.createElement("tr"); hr.className = "bill-group acct-band"; hr.dataset.grpkey = "acct:" + iss;
    const td = document.createElement("td"); td.colSpan = cols.length + 1; td.className = "left";
    const k = document.createElement("span"); k.className = "bg-key"; k.textContent = iss; td.appendChild(k);
    const n = document.createElement("span"); n.className = "acct-band-n"; n.textContent = `${allByIssue[iss]} bill${allByIssue[iss] === 1 ? "" : "s"}`; td.appendChild(n);
    hr.appendChild(td); frag.appendChild(hr); }
  for (const f of list) {
    const tr = document.createElement("tr");
    const chTd = document.createElement("td"); chTd.className = "left acct-check";
    const cb = document.createElement("input"); cb.type = "checkbox"; cb.checked = acctSel.has(f._k);
    cb.onchange = () => { if (cb.checked) acctSel.add(f._k); else acctSel.delete(f._k); _acctUpdateSelAll(); _acctUpdateCopyBtn(); _acctUpdateDownloadBtn(); };
    chTd.appendChild(cb); tr.appendChild(chTd);
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
    { const cc = leftText(f.qbo_class || f.division || "–");   // what QuickBooks HAS on the line; red when it is not the job's division
      if (f.qbo_class && f.division && f.qbo_class !== f.division) { cc.classList.add("acct-class-bad"); cc.title = `QuickBooks says ${f.qbo_class} - ${f.project} is a ${f.division} job`; }
      tr.appendChild(cc); }
    tr.appendChild(leftText(f.cost_code || "–"));
    tr.appendChild(rightText(f.amount != null ? money(f.amount) : ""));
    const mc = document.createElement("td"); mc.className = "left audit-soft"; mc.textContent = f.memo || "–"; if (!f.memo) mc.classList.add("audit-dim"); tr.appendChild(mc);
    const dc = document.createElement("td"); dc.className = "left audit-soft"; dc.textContent = f.detail || ""; tr.appendChild(dc);
    frag.appendChild(tr);
  }
  if (full.length > ACCT_CAP) {
    const tr = document.createElement("tr"), td = document.createElement("td");
    td.colSpan = cols.length + 1; td.className = "left"; td.style.cssText = "padding:10px 14px;color:var(--text-dim)";
    td.textContent = `Showing the first ${ACCT_CAP} of ${full.length} - narrow with the division, vendor or search above (Copy still takes all ${full.length}).`;
    tr.appendChild(td); frag.appendChild(tr);
  }
  }
  tbody.appendChild(frag);
  _acctUpdateSelAll(); _acctUpdateCopyBtn(); _acctUpdateDownloadBtn();
}

// Copy-as-table: the columns copied (headers + values), minus the checkbox and 📎 columns.
const ACCT_COPY_COLS = [["Issue", f => f.issue], ["Vendor", f => f.vendor], ["Bill #", f => f.bill_no],
  ["Date", f => f.date ? fmtDateShort(f.date) : ""], ["Project", f => f.project], ["Class (QuickBooks)", f => f.qbo_class || f.division],
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

// "Copy for clerk" (owner 2026-09-08): one plain sentence per finding that says exactly what to change, ready to paste.
function _acctClerkLine(f) {
  const what = (f.qbo_class && f.division && f.qbo_class !== f.division)
    ? `Class is ${f.qbo_class} in QuickBooks - change it to ${f.division} (${f.project} is a ${f.division} job)`
    : (f.detail || f.issue || "").replace(/^Class\s+[A-Za-z /-]+\s*·\s*/, "");
  return [`Bill ${f.bill_no || "?"}`, f.vendor, f.date ? fmtDateShort(f.date) : null, f.memo ? `"${f.memo}"` : null, f.amount != null ? money(f.amount) : null,
          f.cost_code ? `code ${f.cost_code}` : null].filter(Boolean).join(" · ") + ` -> ${what}`;
}
async function _acctCopyClerk() {
  const rows = acctSel.size ? ((ACCT && ACCT.findings) || []).filter(f => acctSel.has(f._k)) : _acctVisible;
  if (!rows.length) return;
  const txt = rows.map(_acctClerkLine).join("\n");
  let ok = true; try { await navigator.clipboard.writeText(txt); } catch (e) { ok = false; }
  const b = $("#btnAcctClerk"); if (b) { b.disabled = true; b.textContent = ok ? `Copied ${rows.length} ✓` : "Copy failed"; setTimeout(() => { b.disabled = false; b.textContent = "Copy for clerk"; }, 1400); }
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
  if (true) { openAttachmentViewer(f.type === "Expense" || f.type === "Purchase" ? "Purchase" : "Bill", m[1], (f.vendor ? f.vendor + " · " : "") + "bill " + (f.bill_no || m[1])); return; }   // in-app viewer (2026-09-03)
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
  _placeMenu(el, menu);   // the shared placer: flips above the button and caps the height when there is no room below
  setTimeout(() => document.addEventListener("click", _closeScanMenu, { once: true }), 0);
}
function _closeScanMenu() { const m = $("#scanMenu"); if (m) m.remove(); }

function init() {
  applySettings();
  syncSettingsUI();
  wireSettings();
  { const q = $("#search"); let tq = null; if (q) q.addEventListener("input", () => { clearTimeout(tq); tq = setTimeout(renderProjects, 120); }); }
  $$("#projView .seg-btn").forEach(b => b.onclick = () => { projView = b.dataset.view; syncProjChips(); renderProjects(); });
  $$("#projDiv .seg-btn").forEach(b => b.onclick = () => { projDiv = b.dataset.div; syncProjChips(); renderProjects(); });
  { const el = $("#vendorSearch"); if (el) el.addEventListener("input", renderVendors); }
  { const el = $("#vendorGroupType"); if (el) el.addEventListener("change", renderVendors); }
  { const el = $("#vendorSort"); if (el) el.addEventListener("change", renderVendors); }
  { const el = $("#vendorExpandAll"); if (el) el.onclick = _vendorToggleAll; }
  { const el = $("#lienFProj"); if (el) el.addEventListener("input", renderLiens); }   // the other lien filters are multi-selects now
  { const el = $("#billSort"); if (el) el.addEventListener("change", renderBills); }
  // Month + Vendor + every categorical multi-select: the button toggles the checkbox menu; a click outside closes it.
  const _mselWraps = [["#bfDateBtn", "#bfDateMenu", "#bfDateMsel"], ["#bfVendorBtn", "#bfVendorMenu", "#bfVendorMsel"],
    ...BILL_MSEL.map(c => [`#${c.id}Btn`, `#${c.id}Menu`, `#${c.id}Msel`]),
    ...LIEN_MSEL.map(c => [`#${c.id}Btn`, `#${c.id}Menu`, `#${c.id}Msel`]),
    ...PAY_MSEL.map(c => [`#${c.id}Btn`, `#${c.id}Menu`, `#${c.id}Msel`]),
    ...INV_MSEL.map(c => [`#${c.id}Btn`, `#${c.id}Menu`, `#${c.id}Msel`]), ["#ifMonthBtn", "#ifMonthMenu", "#ifMonthMsel"], ["#acctVendorBtn", "#acctVendorMenu", "#acctVendorMsel"]];
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
  { const el = $("#bfHfClear"); if (el) el.onclick = () => { hfClear("bills"); renderBills(); }; }
  { const q = $("#bfQuick"); if (q) { let t = null; q.addEventListener("input", () => { _billQ = q.value; clearTimeout(t); t = setTimeout(renderBills, 160); });
    q.addEventListener("keydown", (e) => { if (e.key === "Escape") { q.value = ""; _billQ = ""; renderBills(); } }); } }
  { const el = $("#bfCollapse"); if (el) el.onclick = billToggleAll; }
  ["#ifDivision", "#ifLien", "#ifLienClock", "#ifLitig", "#ifSort"].forEach(sel => { const el = $(sel); if (el) el.addEventListener("change", renderOpenInvoices); });
  { const el = $("#ifClear"); if (el) el.onclick = invClearFilters; }
  { const el = $("#ifCollect"); if (el) el.onclick = openCollectionsReport; }
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
  // ⌘F / Ctrl+F = the broad search box: the vendor page's when it is open, else the Bill Tracker's (owner 2026-09-22)
  document.addEventListener("keydown", e => {
    if (!(e.metaKey || e.ctrlKey) || e.key.toLowerCase() !== "f" || e.altKey) return;
    if (document.querySelector(".panel:not([hidden])")) return;
    const vis = x => !!(x && x.offsetParent);                 // on screen right now (the Bill Tracker is a section of the Company view, not a tab of its own)
    const el = vis($("#vpSearch")) ? $("#vpSearch") : (vis($("#bfQuick")) ? $("#bfQuick") : null);   // vpSearch = the vendor PAGE box (vendorSearch is the Vendors list's)
    if (!el) return;
    e.preventDefault(); el.focus(); el.select();
  });
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
  { const el = $("#recordBack"); if (el) el.onclick = () => { if (!_popViewIfOwn("record")) closeRecord(); }; }
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
  window.addEventListener("beforeunload", (e) => { if (pendingBillMarks.size || payDraft.size || (typeof _pp !== "undefined" && _pp && _pp.payDraft && _pp.payDraft.size)) { e.preventDefault(); e.returnValue = ""; } });
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
  $("#btnExport").onclick = exportCSV;
  $("#btnRefresh").onclick = manualRefresh;
  { const el = $("#btnResync"); if (el) el.onclick = startResync; }
  { const p = $("#syncPill"); if (p) p.onclick = (e) => { e.stopPropagation(); toggleSyncPop(); }; }   // the pill opens the status breakdown (owner 2026-09-23)
  { const el = $("#btnGearWip"); if (el) el.onclick = () => { closePanels(); setTab("wipreview"); }; }
  { const el = $("#btnGearRp"); if (el) el.onclick = () => { closePanels(); setTab("review"); }; }
  { const el = $("#btnGearConsole"); if (el) el.onclick = () => { closePanels(); setTab("console"); }; }
  { const el = $("#btnGearSystems"); if (el) el.onclick = () => { closePanels(); setTab("systems"); }; }
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
  { const el = $("#qaSearch"); if (el) { el.addEventListener("input", renderQboAudit); el.addEventListener("keydown", e => { if (e.key === "Escape") { el.value = ""; renderQboAudit(); } }); } }
  { const el = $("#qaFlaggedOnly"); if (el) el.onchange = renderQboAudit; }
  $$("#qaDays .seg-btn").forEach(b => { b.onclick = () => { qaDays = Number(b.dataset.days); loadQboAudit(true); }; });
  { const el = $("#btnQaReload"); if (el) el.onclick = () => loadQboAudit(true); }
  { const el = $("#btnCdReload"); if (el) el.onclick = () => loadCheckDrift(true); }
  { const el = $("#btnUcRefresh"); if (el) el.onclick = () => runPipeline("uncleared", null, { btn: el, prog: $("#ucProg"), fill: $("#ucFill"), step: $("#ucStep"), after: () => loadUncleared(true) }); }
  { const el = $("#ucSearch"); if (el) { el.addEventListener("input", renderUncleared); el.addEventListener("keydown", e => { if (e.key === "Escape") { el.value = ""; renderUncleared(); } }); } }
  { const el = $("#btnCdRefresh"); if (el) el.onclick = () => runPipeline("mirror", null, { btn: el, prog: $("#cdProg"), fill: $("#cdFill"), step: $("#cdStep"), after: () => loadCheckDrift(true) }); }
  { const el = $("#cdSearch"); if (el) { el.addEventListener("input", renderCheckDrift); el.addEventListener("keydown", e => { if (e.key === "Escape") { el.value = ""; renderCheckDrift(); } }); } }
  { const el = $("#btnQaRefresh"); if (el) el.onclick = () => runPipeline("mirror", null, { btn: el, prog: $("#qaProg"), fill: $("#qaFill"), step: $("#qaStep"), after: () => loadQboAudit(true) }); }
  { const el = $("#btnAcctCopy"); if (el) el.onclick = _acctDoCopy; }
  { const el = $("#btnAcctClerk"); if (el) el.onclick = _acctCopyClerk; }
  { const el = $("#btnAcctDownload"); if (el) el.onclick = _acctDoDownload; }
  for (const id of ["#acctSearch", "#acctDivision"]) { const el = $(id); if (el) el.addEventListener("input", () => { if (ACCT && ACCT.ok) renderAccounting(); }); }
  { const el = $("#wrCompute"); if (el) el.onclick = runWipReview; }
  document.querySelectorAll("#wrView .seg-btn").forEach(b => b.onclick = () => {
    wrView = b.dataset.view; wrIdx = 0;
    document.querySelectorAll("#wrView .seg-btn").forEach(x => x.classList.toggle("on", x === b));
    if (WR && WR.ready) renderWipReview();
  });
  { const el = $("#wrSync"); if (el) el.onclick = syncWipReview; }
  { const el = $("#wrApproveQbo"); if (el) el.onclick = () => wrBulk("qbo"); }
  { const el = $("#wrApproveAll"); if (el) el.onclick = () => wrBulk("all"); }
  { const el = $("#wrClearAll"); if (el) el.onclick = () => wrBulk("clear"); }
  for (const id of ["#wrSearch", "#wrDivision", "#wrStatus", "#wrChangedOnly", "#wrBlock"]) {
    const el = $(id); if (el) el.addEventListener("input", () => { if (WR && WR.ready) renderWipReview(); });
  }
  if ($("#wrCopyPm")) $("#wrCopyPm").onclick = () => { if (WR && WR.ready) wrCopyPmQuestions(); };
  buildGroupBar();   // the two views (sub-tabs render on setTab)
  initFolds();
  syncProjChips();
  const bootRec = _recLoad();   // read BEFORE the first setTab (which clears it)
  let savedTab = "projects";
  try { savedTab = localStorage.getItem("proficient-ledger-tab") || "projects"; } catch { /* ignore */ }
  if (["wipreview", "rpreview", "review", "console", "systems"].includes(savedTab)) savedTab = "projects";   // tool pages never reopen on their own
  setTab(TAB_ALIAS[savedTab] || savedTab);
  initCellSelect();   // Excel-style click/drag cell selection + running-sum bar
  setInterval(() => { if (!syncing && !pendingBillMarks.size && !payDraft.size) load(true); }, 90000);   // soft auto-refresh (paused during a resync or while lien / pay-run marks are unsaved)
  $("#btnCloseSettings").onclick = closePanels;
  $("#btnCloseDetail").onclick = closePanels;
  $("#btnCopyDetail").onclick = () => copy(detailAsText());
  $("#overlay").onclick = closePanels;
  document.addEventListener("keydown", e => { if (e.key === "Escape") closePanels(); });
  matchMedia("(prefers-color-scheme: dark)").addEventListener("change", () => { if (settings.theme === "auto") applySettings(); });
  Promise.resolve(load()).then(() => _recRestore(bootRec)).catch(() => { /* the tab still shows */ });
}
init();


// ── WIP review: the weekly sit-down with the division PM (owner 2026-09-15) ─────────────────
// ONE page per division - RP, CP, MFD - the same system for all three ("that way we all can use
// the same system and update together"). One card per WIP line: the numbers on the WIP, WHERE each
// was grabbed from (a source pill + the file), and "Our numbers" - the owner's / PM's own space:
// prepopulated with the prepared value, a check or an X per number, a note for what could not be
// settled. Every answer is stamped with the MODE (Me = the owner alone, PM = the division PM is here,
// deciding together) and the time, and kept in rp_review_mark (+ a log). The prepared numbers are
// never edited from here. RP carries the crew schedule and pictures of the real files; CP and MFD
// have no schedule - their cards show the contract and approved COs from the draw G702, the ETC from
// the takeoff, costs and billed from QuickBooks (MFD's contract / ETC are typed on the master by design).
let RR = null, rrPoll = null, rrWired = false;
let rrMode = (() => { try { const m = localStorage.getItem("proficient-ledger-rrmode") || "me"; return m === "ops" ? "pm" : m; } catch { return "me"; } })();
let rrDiv = (() => { try { return localStorage.getItem("proficient-ledger-rrdiv") || "RP"; } catch { return "RP"; } })();
let rrSection = "current";
const RR_DUE_DAYS = 7;
const RR_DIV_NAME = { RP: "Residential", CP: "Commercial", MFD: "Multi Family" };
const rrWho = () => rrMode === "pm" ? "PM + you" : "You";
const rrWhoLow = () => rrMode === "pm" ? "PM + you" : "you";
const rrApi = () => `/api/review?div=${encodeURIComponent(rrDiv)}`;

async function loadReview(force) {
  const body = $("#rrBody"); if (!body) return;
  rrWire();
  if (rrPoll && !force) return;
  try { RR = await (await fetch(rrApi())).json(); }
  catch { body.innerHTML = `<div class="rr-empty">could not load</div>`; return; }
  { const t = $("#rrTitle"); if (t) t.textContent = `WIP review · ${rrDiv}`; }
  { const s = $("#rrSection"); if (s) s.hidden = rrDiv !== "RP"; }
  { const f = $("#rrFinalize"); if (f) f.hidden = rrDiv !== "RP"; }
  { const w = $("#rrWipReview"); if (w) w.hidden = rrDiv === "RP"; }
  if (rrDiv !== "RP") rrSection = "current";
  if (!RR.ready) {
    $("#rrFilters").hidden = true; $("#rrStats").innerHTML = "";
    { const note = $("#rrNote"); if (note) note.textContent = ""; }
    body.innerHTML = rrDiv === "RP"
      ? `<div class="rr-empty"><p>No RP review built yet.</p><p class="hint">Hit <b>Rebuild review data</b> in the gear: it reads the master's RP tab,
        the RP WIP file, every job folder, JobTread and the crew schedules, and pictures the source file for every number. A few minutes.</p></div>`
      : `<div class="rr-empty"><p>No ${rrDiv} update computed yet.</p><p class="hint">Hit <b>Rebuild review data</b> in the gear: it runs the ${rrDiv} WIP reader
        (${rrDiv === "CP" ? "the job folders on Common - draw G702s and takeoffs - plus QuickBooks" : "the master's MFD rows plus QuickBooks"}; Touch ID) and lays every line out here with the document each number came from. A few minutes.</p></div>`;
    return;
  }
  rrRender();
}
const loadRpReview = loadReview;   // the old name, for any deep link that still uses it

function rrWire() {
  if (rrWired) return; rrWired = true;
  $$("#rrMode .seg-btn").forEach(b => { b.onclick = () => { rrMode = b.dataset.mode; try { localStorage.setItem("proficient-ledger-rrmode", rrMode); } catch { /* ignore */ } rrPaintMode(); }; });
  $$("#rrDiv .seg-btn").forEach(b => { b.onclick = () => { rrDiv = b.dataset.div; try { localStorage.setItem("proficient-ledger-rrdiv", rrDiv); } catch { /* ignore */ } rrOpenLine = null; rrLastLine = null; rrJustSaved = null; rrPaintMode(); loadReview(true); }; });
  $$("#rrSection .seg-btn").forEach(b => { b.onclick = () => { rrSection = b.dataset.sec; rrOpenLine = null; $$("#rrSection .seg-btn").forEach(x => x.classList.toggle("on", x === b)); rrRenderCards(); }; });
  { const q = $("#rrSearch"); if (q) q.oninput = rrRenderCards; }
  { const c = $("#rrDueOnly"); if (c) c.onchange = rrRenderCards; }
  { const c = $("#rrProbOnly"); if (c) c.onchange = rrRenderCards; }
  { const b = $("#rrRebuild"); if (b) b.onclick = () => { closePanels(); setTab("review"); rrRebuild(); }; }
  { const b = $("#rrFinalize"); if (b) b.onclick = rrFinalize; }
  { const b = $("#rrWipReview"); if (b) b.onclick = () => setTab("wipreview"); }
  rrPaintMode();
}
function openReviewDiv(div) {   // from the Projects page: a division band's Review button
  if (div && RR_DIV_NAME[div]) { rrDiv = div; try { localStorage.setItem("proficient-ledger-rrdiv", rrDiv); } catch { /* ignore */ } rrOpenLine = null; rrLastLine = null; }
  setTab("review");
}

function rrPaintMode() {
  $$("#rrMode .seg-btn").forEach(b => b.classList.toggle("on", b.dataset.mode === rrMode));
  $$("#rrDiv .seg-btn").forEach(b => b.classList.toggle("on", b.dataset.div === rrDiv));
  // no badge - the toggle itself says who is here (owner 2026-09-15)
}

function rrMark(x, kind) { return (RR.marks || {})[`${x.line}|${kind}`] || null; }
function rrIsDue(m) { if (!m || !m.at) return true; return (Date.now() - new Date(m.at).getTime()) > RR_DUE_DAYS * 86400e3; }

function rrRender() {
  const note = $("#rrNote");
  if (note) {
    const m = RR.master || {};
    note.textContent = rrDiv === "RP"
      ? `built ${fmtDate(RR.built_at, true)} · master ${m.tab || ""} (saved ${fmtDate(m.mtime, true)})` + (m.renamed ? " · tab renamed - the sync writes Test - RP" : "")
      : `pending update computed ${fmtDate(RR.built_at, true)} · ${m.tab || ""} · ${RR_DIV_NAME[rrDiv]}`;
  }
  $("#rrFilters").hidden = false;
  rrRenderStats();
  rrRenderCards();
}

function rrRenderStats() {
  const el = $("#rrStats"); if (!el) return;
  const c = RR.counts || {}, cur = RR.current || [], fin = RR.finished || [];
  const okm = m => m && !rrIsDue(m) && (m.decision === "confirmed" || m.decision === "agree");
  const answered = cur.filter(x => okm(rrMark(x, "current"))).length + fin.filter(x => okm(rrMark(x, "finished"))).length;
  const due = cur.length + fin.length - answered;
  const tiles = rrDiv === "RP"
    ? [["Current lines", cur.length, ""], ["Finished lines", fin.length, ""], ["Confirmed this week", answered, answered ? "green" : ""],
       ["Due", due, due ? "amber" : ""], ["Typed on the master", c.typed || 0, c.typed ? "red" : ""], ["No approved JobTread proposal", cur.length - (c.in_jobtread || 0), "amber"]]
    : [[`${rrDiv} lines`, cur.length, ""], ["Changed by this update", c.changed || 0, c.changed ? "amber" : ""], ["Confirmed this week", answered, answered ? "green" : ""],
       ["Due", due, due ? "amber" : ""], [rrDiv === "MFD" ? "Typed on the master" : "Problems", rrDiv === "MFD" ? (c.typed || 0) : (c.problems || 0), (rrDiv === "MFD" ? c.typed : c.problems) ? "amber" : ""]];
  el.innerHTML = "";
  for (const [label, val, cls] of tiles) {
    const d = document.createElement("div"); d.className = "kpi" + (cls ? " wr-kpi-" + cls : "");
    d.innerHTML = `<div class="k-label">${_ge(label)}</div><div class="k-value">${val}</div>`; el.appendChild(d);
  }
}

function rrVisible() {
  const q = ($("#rrSearch")?.value || "").trim().toLowerCase();
  const dueOnly = $("#rrDueOnly")?.checked, probOnly = $("#rrProbOnly")?.checked;
  const kind = rrSection, list = kind === "current" ? (RR.current || []) : (RR.finished || []);
  return list.filter(x => {
    if (q && !`${x.line} ${x.name || ""} ${x.builder || ""}`.toLowerCase().includes(q)) return false;
    if (dueOnly && !rrIsDue(rrMark(x, kind))) return false;
    if (probOnly && kind === "current" && !x.problem) return false;
    return true;
  });
}


// ── The WIP change log (owner 2026-09-17: "an audit log of the etc/contract changing that can be
// pulled up easily just like qbo"): every contract / CO / ETC / billed / costs change the WIP writer
// made (with its source) plus every answer saved on the review page, newest first, per job.
const AUDIT_FIELD = { contract: "Contract", approved_cos: "Approved COs", etc: "ETC", co_costs: "CO costs", billed: "Billed to date", costs: "Costs to date", retainage: "Retainage", line: "Line" };
async function loadAuditLog(pn) {
  try { return await (await fetch(`/api/wip/audit?no=${encodeURIComponent(pn)}&limit=300`)).json(); } catch { return null; }
}
function auditTable(d) {
  const rows = [];
  for (const e of (d && d.entries) || []) {
    const isLine = e.field === "line";
    rows.push({ at: e.at, what: AUDIT_FIELD[e.field] || e.field, change: isLine ? _ge(e.note || (e.new ? "added" : "removed")) : `${e.old == null ? "blank" : money(e.old)} → <b>${e.new == null ? "blank" : money(e.new)}</b>`, src: e.source || "", who: e.actor || "", run: e.run || "" });
  }
  for (const a of (d && d.answers) || []) {
    const bits = [];
    if (a.decision) bits.push({ confirmed: "Confirmed", fix: "Needs a fix", agree: "Agreed done", keep: "Kept on the WIP", noted: "Noted", cleared: "Answer cleared" }[a.decision] || a.decision);
    if (a.contract_ok != null) bits.push(`contract ${a.contract_ok ? "✓" : "✗"}`);
    if (a.etc_ok != null) bits.push(`ETC ${a.etc_ok ? "✓" : "✗"}`);
    if (a.our_contract != null) bits.push(`their contract ${money(a.our_contract)}`);
    if (a.our_etc != null) bits.push(`their ETC ${money(a.our_etc)}`);
    rows.push({ at: a.at, what: "Review answer", change: _ge(bits.join(" · ")), src: a.note ? `"${_ge(a.note)}"` : "", who: a.mode === "ops" || a.mode === "pm" ? "PM + owner" : "owner", run: "review page" });
  }
  rows.sort((x, y) => String(y.at).localeCompare(String(x.at)));
  if (!rows.length) return `<div class="hint" style="margin:4px 0">No changes logged yet for this job.</div>`;
  return `<table class="rr-tl audit-tl"><thead><tr><th>When</th><th>What</th><th>Change</th><th>Source</th><th>Who</th><th>Run</th></tr></thead><tbody>` +
    rows.map(r => `<tr><td class="n">${fmtDate(r.at, true)}</td><td>${_ge(r.what)}</td><td>${r.change}</td><td>${typeof r.src === "string" && r.src.startsWith('"') ? r.src : _ge(r.src)}</td><td>${_ge(r.who)}</td><td class="d">${_ge(r.run)}</td></tr>`).join("") + `</tbody></table>`;
}
async function fillAuditInto(el, pn) {
  if (!el) return;
  el.innerHTML = `<div class="hint" style="margin:4px 0">loading the change log…</div>`;
  const d = await loadAuditLog(pn);
  el.innerHTML = d ? auditTable(d) : `<div class="hint" style="margin:4px 0">could not load the change log</div>`;
}

let rrOpenLine = null;   // the job whose full page is open (null = the list)
let rrLastLine = null;   // the job you last had open - the list scrolls back to it and marks the row (owner 2026-09-15)
const RR_DEC_LABEL = { confirmed: "Confirmed", fix: "Needs a fix", agree: "Agreed done", keep: "Kept on the WIP", noted: "Noted, not confirmed" };

function rrRenderCards() {
  const body = $("#rrBody"); if (!body || !RR) return;
  const kind = rrSection, list = rrVisible();
  const cnt = $("#rrCount"); if (cnt) cnt.textContent = `${list.length} shown`;
  body.innerHTML = "";
  if (rrOpenLine) rrLastLine = rrOpenLine;
  if (rrOpenLine) {
    const x = list.find(r => r.line === rrOpenLine) || (kind === "current" ? RR.current : RR.finished).find(r => r.line === rrOpenLine);
    if (x) { rrRenderPage(x, kind, list); return; }
    rrOpenLine = null;
  }
  if (!list.length) { body.innerHTML = `<div class="rr-empty">Nothing to show with these filters.</div>`; return; }
  const t = document.createElement("table"); t.className = "rr-list";
  const isRp = rrDiv === "RP";
  const cols = !isRp
    ? ["", "Job", "Name", "Client", "Contract", "Approved COs", "ETC", "Costs", "Billed", "This update", "Answer"]
    : kind === "current"
      ? ["", "Job", "Address", "Builder", "Contract", "ETC", "Costs", "Billed", "Last on schedule", "To settle", "Answer"]
      : ["", "Job", "Address", "Builder", "Contract", "Billed", "Costs", "Gross profit", "Net (10% OH)", "Last day", "Answer"];
  t.innerHTML = `<thead><tr>${cols.map(c => `<th>${c}</th>`).join("")}</tr></thead>`;
  const tb = document.createElement("tbody");
  for (const x of list) {
    const m = rrMark(x, kind), due = rrIsDue(m);
    const tr = document.createElement("tr"); tr.className = ((m && !due && (m.decision === "confirmed" || m.decision === "agree")) ? "rr-ok-row" : (x.problem ? "rr-prob-row" : "")) + (x.line === rrLastLine ? " rr-here" : "");
    if (x.line === rrLastLine) tr.title = "you were here";
    // the list's mark: ✓ only for Confirmed / Agreed; a fix or a keep shows "!"; a plain note shows nothing;
    // deselect the verdict and Save = unconfirmed (owner 2026-09-15)
    const who = m ? (m.mode === "pm" ? "PM + you" : "You") + " · " + fmtDate(m.at, true) : "";
    const verdictOk = m && (m.decision === "confirmed" || m.decision === "agree");
    const mark = !m ? "" : due ? `<span class="rr-check due" title="answered ${fmtDate(m.at)} - due again">↻</span>`
      : verdictOk ? `<span class="rr-check" title="${_ge(who)}">✓</span>`
      : (m.decision === "fix" || m.decision === "keep") ? `<span class="rr-check fix" title="${_ge(who)}">!</span>`
      : "";
    const ans = m ? `${_ge(RR_DEC_LABEL[m.decision] || m.decision)}${x.line === rrJustSaved ? ` <span class="rr-saved">saved ✓</span>` : ""}<div class="d">${m.mode === "pm" ? "PM + you" : "You"} · ${fmtDate(m.at, true)}</div>` : `<span class="rr-nojt">not yet</span>`;
    let mid;
    if (!isRp) {
      const pend = x.pending || [];
      const pendTxt = pend.length ? `${pend.length} · ${pend.map(f => f.label.toLowerCase()).join(", ")}` : (x.status === "ADDED" ? "new on the WIP" : "");
      mid = `<td class="n">${x.contract == null ? '<span class="miss">blank</span>' : money(x.contract)}</td><td class="n">${x.cos == null ? "–" : money(x.cos)}</td><td class="n">${x.etc == null ? '<span class="miss">blank</span>' : money(x.etc)}</td>`
        + `<td class="n">${money(x.costs)}</td><td class="n">${money(x.billed)}</td><td class="rr-list-flags">${_ge(pendTxt)}${(x.flags || []).length ? `<div class="d">${_ge(x.flags[0])}${x.flags.length > 1 ? "…" : ""}</div>` : ""}</td>`;
    } else if (kind === "current") {
      mid = `<td class="n">${x.contract == null ? '<span class="miss">blank</span>' : money(x.contract)}</td><td class="n">${x.etc == null ? '<span class="miss">blank</span>' : money(x.etc)}</td>`
        + `<td class="n">${money(x.costs)}</td><td class="n">${money(x.billed)}</td><td class="${x.stale ? "rr-stale" : ""}">${x.last_seen ? fmtDate(x.last_seen) : "–"}</td><td class="rr-list-flags">${(x.flags || []).length ? `${x.flags.length} · ${_ge(x.flags[0])}${x.flags.length > 1 ? "…" : ""}` : ""}</td>`;
    } else {
      const gp = (x.billed != null && x.costs != null) ? x.billed - x.costs : null, net = (gp != null && x.contract != null) ? gp - x.contract * 0.10 : null;
      mid = `<td class="n">${x.contract == null ? '<span class="miss">blank</span>' : money(x.contract)}</td><td class="n">${money(x.billed)}</td><td class="n">${money(x.costs)}</td><td class="n ${gp != null && gp < 0 ? "neg" : ""}">${gp == null ? "–" : money(gp)}${gp != null && x.billed ? `<div class="d">${(gp / x.billed * 100).toFixed(1)}%</div>` : ""}</td><td class="n ${net != null && net < 0 ? "neg" : ""}">${net == null ? "–" : money(net)}</td><td>${_ge(x.last_day || "–")}</td>`;
    }
    // answer right on the list (owner 2026-09-15: "we only click if we need details"): the verdict + Save;
    // numbers and notes stay on the job page
    const dec = kind === "current" ? [["confirmed", "Confirmed"], ["fix", "Needs a fix"]] : [["agree", "Agree - done"], ["keep", "Keep on the WIP"]];
    const quick = `<div class="rr-quick"><span class="seg rr-dec">${dec.map(([d, l]) => `<button type="button" class="seg-btn ${m && m.decision === d ? "on" : ""}" data-dec="${d}">${l}</button>`).join("")}</span><button type="button" class="btn tiny primary" data-save="1">Save</button></div>`;
    tr.innerHTML = `<td class="rr-markcell">${mark}</td><td class="wr-pn">${_ge(x.line)}</td><td>${_ge(x.name || "")}</td><td>${_ge(x.builder || "")}</td>${mid}<td class="rr-ans">${ans}${quick}</td>`;
    tr.onclick = (e) => { if (e.target.closest(".rr-quick")) return; rrOpenLine = x.line; _pushView({ v: "rr", line: x.line }); rrRenderCards(); window.scrollTo(0, 0); };
    tr.querySelectorAll(".rr-dec .seg-btn").forEach(b => { b.onclick = (e) => { e.stopPropagation(); const on = b.classList.contains("on"); tr.querySelectorAll(".rr-dec .seg-btn").forEach(o => o.classList.remove("on")); if (!on) b.classList.add("on"); }; });
    tr.querySelector("[data-save]").onclick = (e) => { e.stopPropagation(); const on = tr.querySelector(".rr-dec .seg-btn.on"); rrSaveQuick(tr, x, kind, on ? on.dataset.dec : "noted"); };
    tb.appendChild(tr);
  }
  t.appendChild(tb); body.appendChild(t);
  const here = t.querySelector("tr.rr-here"); if (here) requestAnimationFrame(() => here.scrollIntoView({ block: "center" }));
}

function rrRenderPage(x, kind, list) {
  const body = $("#rrBody");
  const i = list.findIndex(r => r.line === x.line);
  const nav = document.createElement("div"); nav.className = "rr-pagenav";
  nav.innerHTML = `<button type="button" class="btn small" id="rrBack">← Back to the list</button>
    <span class="rr-pagepos">${i >= 0 ? `${i + 1} of ${list.length}` : ""}</span>
    <button type="button" class="btn small" id="rrPrev" ${i <= 0 ? "disabled" : ""}>← Previous</button>
    <button type="button" class="btn small" id="rrNext" ${i < 0 || i >= list.length - 1 ? "disabled" : ""}>Next →</button>`;
  body.appendChild(nav);
  body.appendChild(rrDiv === "RP" ? rrCard(x, kind) : rrCardDiv(x));
  $("#rrBack").onclick = () => { if (!_popViewIfOwn("rr")) { rrOpenLine = null; rrRenderCards(); } };
  $("#rrPrev").onclick = () => { if (i > 0) { rrOpenLine = list[i - 1].line; try { history.replaceState({ v: "rr", line: rrOpenLine }, ""); } catch { /* ignore */ } rrRenderCards(); window.scrollTo(0, 0); } };
  $("#rrNext").onclick = () => { if (i < list.length - 1) { rrOpenLine = list[i + 1].line; try { history.replaceState({ v: "rr", line: rrOpenLine }, ""); } catch { /* ignore */ } rrRenderCards(); window.scrollTo(0, 0); } };
}

function rrPill(src) { return src ? `<span class="rr-src ${_ge(src.kind)}" title="${_ge(src.detail || "")}">${_ge(src.label)}</span>` : ""; }
function rrMoney(v) { return v == null ? `<span class="miss">blank</span>` : money(v); }

function rrPic(pic, caption) {
  if (!pic || !pic.img) return "";
  const src = "/api/rp/img/" + pic.img.split("/").map(encodeURIComponent).join("/");
  const sub = [pic.file ? _ge(pic.file) : "", pic.sheet ? "sheet " + _ge(pic.sheet) : (pic.page ? `page ${pic.page} of ${pic.pages}` : ""), pic.anchor ? _ge(pic.anchor) : "", pic.note ? `<i class="rr-picnote">${_ge(pic.note)}</i>` : ""].filter(Boolean).join(" · ");
  return `<figure class="rr-pic"><figcaption><b>${_ge(caption)}</b>${sub ? " · " + sub : ""}</figcaption><a href="${src}" target="_blank" rel="noopener" title="Open full size"><img src="${src}" loading="lazy" alt="${_ge(caption)}"></a></figure>`;
}

function rrSrcLine(sr) {   // the source's explanation, unless it is just the file name shown next to it
  const d = (sr && sr.detail) || "", f = (sr && sr.file) ? sr.file.split("/").pop() : "";
  return d && d !== f ? _ge(d) + " · " : "";
}

function rrTimeline(sc) {
  if (!sc || !sc.days) return `<div class="hint" style="margin:4px 0">Never on a crew schedule.</div>`;
  const rows = (sc.runs || []).map(r => `<tr><td class="n">${fmtDate(r.from)}${r.days > 1 ? ` to ${fmtDate(r.to)}` : ""}</td><td class="n">${r.days}</td><td>${_ge(r.section || "")}</td><td>${_ge(r.stage || "")}</td></tr>`).join("");
  return `<div class="rr-tl-sum">${sc.days} day${sc.days === 1 ? "" : "s"} on the crew schedule · first ${fmtDate(sc.first)} · last ${fmtDate(sc.last)}</div>
    <table class="rr-tl"><thead><tr><th>When</th><th>Days</th><th>Section</th><th>Task</th></tr></thead><tbody>${rows}</tbody></table>`;
}

async function rrReveal(path) {
  if (!path) return;
  let res; try { res = await (await fetch("/api/rp/reveal", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ path }) })).json(); }
  catch { toast("could not open"); return; }
  if (!res.ok) toast(res.error || "could not open");
}

// "Our numbers" + the verdict, the same block on every division's card
function rrOursHtml(x, kind, m, due, withCos) {
  const ourK = m && m.our_contract != null ? m.our_contract : x.contract, ourE = m && m.our_etc != null ? m.our_etc : x.etc, ourC = m && m.our_cos != null ? m.our_cos : x.cos;
  const okBtn = (f, val) => `<span class="rr-ok" data-f="${f}"><button type="button" class="yes ${val === 1 ? "on" : ""}" data-v="1" title="right">✓</button><button type="button" class="no ${val === 0 ? "on" : ""}" data-v="0" title="wrong">✗</button></span>`;
  const dec = kind === "current" ? [["confirmed", "Confirmed"], ["fix", "Needs a fix"]] : [["agree", "Agree - done"], ["keep", "Keep on the WIP"]];
  const stamp = m ? `<span class="rr-stamp ${m.mode === "pm" ? "rr-stamp-ops" : ""}"><b>${RR_DEC_LABEL[m.decision] || _ge(m.decision)}</b> · ${m.mode === "pm" ? "PM + you" : "You"} · ${fmtDate(m.at, true)}${due ? " · <i>due again</i>" : ""}</span>` : `<span class="rr-stamp"><i>not saved yet</i></span>`;
  const fmtN = v => v == null ? "" : Number(v).toLocaleString();
  return `<div class="rr-ours">
      <div class="rr-ours-row">
        <label class="rr-field">Contract<input type="text" data-f="our_contract" value="${fmtN(ourK)}"></label>${okBtn("contract_ok", m ? m.contract_ok : null)}
        ${withCos ? `<label class="rr-field">Approved COs<input type="text" data-f="our_cos" value="${fmtN(ourC)}"></label>${okBtn("cos_ok", m ? m.cos_ok : null)}` : ""}
        <label class="rr-field">ETC<input type="text" data-f="our_etc" value="${fmtN(ourE)}"></label>${okBtn("etc_ok", m ? m.etc_ok : null)}
        <label class="rr-field rr-note">Notes - what changed, what could not be settled<textarea data-f="note">${_ge(m?.note || "")}</textarea></label>
      </div>
      <div class="rr-actions"><span class="seg rr-dec">${dec.map(([d, l]) => `<button type="button" class="seg-btn ${m && m.decision === d ? "on" : ""}" data-dec="${d}">${l}</button>`).join("")}</span>
        <button type="button" class="btn small primary" data-save="1">Save</button>${m ? `<button type="button" class="btn small subtle" data-clear="1">Clear answer</button>` : ""}${stamp}</div>
    </div>`;
}
function rrWireOurs(card, x, kind) {
  card.querySelectorAll("[data-reveal]").forEach(b => { b.onclick = () => rrReveal(b.dataset.reveal); });
  { const b = card.querySelector("[data-project]"); if (b) b.onclick = () => openProjectPage(b.dataset.project); }
  if (rrOpenLine === x.line) fillAuditInto(card.querySelector("[data-audit]"), x.line);   // the full page only, not list rows
  { const b = card.querySelector("[data-wipreview]"); if (b) b.onclick = () => setTab("wipreview"); }
  card.querySelectorAll('input[data-f]').forEach(inp => { const base = inp.dataset.f === "our_contract" ? x.contract : inp.dataset.f === "our_cos" ? x.cos : x.etc;
    const paint = () => { const v = Number(String(inp.value).replace(/[$,]/g, "")); inp.classList.toggle("changed", inp.value.trim() !== "" && !Number.isNaN(v) && base != null && Math.abs(v - base) > 0.5); }; inp.oninput = paint; paint(); });
  card.querySelectorAll(".rr-ok button").forEach(b => { b.onclick = () => { const on = b.classList.contains("on"); b.parentElement.querySelectorAll("button").forEach(o => o.classList.remove("on")); if (!on) b.classList.add("on"); }; });
  card.querySelectorAll(".rr-dec .seg-btn").forEach(b => { b.onclick = () => { const on = b.classList.contains("on"); card.querySelectorAll(".rr-dec .seg-btn").forEach(o => o.classList.remove("on")); if (!on) b.classList.add("on"); }; });
  card.querySelector("[data-save]").onclick = () => { const on = card.querySelector(".rr-dec .seg-btn.on"); rrSave(card, x, kind, on ? on.dataset.dec : "noted"); };   // no verdict = "noted": numbers + note kept, the ✓ comes off
  { const c = card.querySelector("[data-clear]"); if (c) c.onclick = () => { if (confirm(`Clear the saved answer on ${x.line}?`)) rrSave(card, x, kind, ""); }; }
}

function rrCard(x, kind) {   // the RP card: schedule first, then the contract and the ETC with pictures of the real files
  const m = rrMark(x, kind), due = rrIsDue(m);
  const card = document.createElement("div");
  card.className = "rr-card " + ((m && !due && (m.decision === "confirmed" || m.decision === "agree")) ? "rr-done" : (x.problem ? "rr-prob" : "rr-due"));
  const jt = x.jt || {}, src = x.src || {}, pics = x.pics || {};
  const link = (v, href) => href ? `<a href="${_ge(href)}" target="_blank" rel="noopener" title="Open in QuickBooks">${rrMoney(v)}</a>` : rrMoney(v);
  const jtLink = jt.url ? `<a class="btn tiny" href="${_ge(jt.url)}" target="_blank" rel="noopener">Open in JobTread</a>` : `<span class="rr-nojt">not in JobTread</span>`;
  const folderBtn = x.folder ? `<button type="button" class="btn tiny" data-reveal="${_ge(x.folder)}" title="${_ge(x.folder)}">Open job folder</button>` : "";
  const pageBtn = `<button type="button" class="btn tiny" data-project="${_ge(x.line)}" title="This job's page in the ledger: invoices, draws, bills, costs">Project page</button>`;
  const qboBtn = x.billed_link ? `<a class="btn tiny" href="${_ge(x.billed_link)}" target="_blank" rel="noopener" title="The project in QuickBooks (invoices)">Open in QuickBooks</a>` : "";
  const qboPl = x.costs_link ? `<a class="btn tiny" href="${_ge(x.costs_link)}" target="_blank" rel="noopener" title="The project's P&L in QuickBooks (costs)">QBO P&amp;L</a>` : "";
  // one line of numbers, each with its source, then the profit line (owner 2026-09-15: "remove big
  // block just show the numbers and below it gross profit - oh 10% net"). GP = contract - ETC (the WIP's
  // original profit), overhead = 10% of the contract, net = GP - overhead.
  const cell = (l, v, pill) => `<span class="rr-n"><span class="l">${l}</span> <span class="v">${v}</span>${pill || ""}</span>`;
  const K = x.contract, E = x.etc;
  // current: GP = contract - ETC (the WIP's original profit); finished: GP = billed - costs (actuals)
  const fin = kind === "finished";
  const gp = fin ? ((x.billed != null && x.costs != null) ? x.billed - x.costs : null) : ((K != null && E != null) ? K - E : null);
  const oh = K != null ? K * 0.10 : null, net = (gp != null && oh != null) ? gp - oh : null;
  const base = fin ? x.billed : K;
  const pctTxt = (gp != null && base) ? ` (${(gp / base * 100).toFixed(1)}%)` : "";
  const profit = K != null ? `<div class="rr-profit">Gross profit${fin ? " (billed - costs)" : ""} <b class="${gp != null && gp < 0 ? "neg" : ""}">${gp == null ? "–" : money(gp)}</b>${pctTxt} · overhead 10% <b>${money(oh)}</b> · net <b class="${net != null && net < 0 ? "neg" : ""}">${net == null ? "–" : money(net)}</b></div>` : "";
  const jtTxt = jt.price != null ? `${money(jt.price)} / ${money(jt.cost)}` : `<span class="miss">${jt.exists ? "no approved proposal" : "not in JobTread"}</span>`;
  const nums = (kind === "current"
    ? cell("Contract", rrMoney(K), rrPill(src.contract)) + cell("ETC", rrMoney(E), rrPill(src.etc)) + cell("Costs", link(x.costs, x.costs_link), rrPill(src.costs)) + cell("Billed", link(x.billed, x.billed_link), rrPill(src.billed)) + cell("JobTread", jtTxt, jt.price != null ? `<span class="rr-src jt" title="approved ${fmtDate(jt.date)}${jt.scope_note ? " · " + _ge(jt.scope_note) : ""}">${fmtDate(jt.date)}</span>` : "")
    : cell("Contract", rrMoney(K), `<span class="rr-src master">RP WIP file</span>`) + cell("ETC", rrMoney(E), `<span class="rr-src master">RP WIP file</span>`) + cell("Billed", rrMoney(x.billed), `<span class="rr-src qbo" title="as of 09/09/2026">QuickBooks</span>`) + cell("Costs", rrMoney(x.costs), `<span class="rr-src qbo" title="as of 09/09/2026">QuickBooks</span>`) + cell("JobTread", jtTxt, ""))
    + profit;
  const flags = kind === "current" ? (x.flags || []) : [x.why].filter(Boolean);
  // 1. the schedule - where the project came from
  const schedBlock = `<div class="rr-sec"><div class="rr-sec-title">1 · On the crew schedule</div>${rrTimeline(x.schedule)}
    <div class="rr-pics">${rrPic(pics.sched_first, "First day " + (pics.sched_first ? fmtDate(pics.sched_first.date) : ""))}${pics.sched_last && (!pics.sched_first || pics.sched_last.date !== pics.sched_first.date) ? rrPic(pics.sched_last, "Last day " + fmtDate(pics.sched_last.date)) : ""}</div></div>`;
  // 2. the contract - the file and the page it sits on
  const fileBtn = (f, label) => f ? `<button type="button" class="btn tiny" data-reveal="${_ge(f)}" title="Open the folder and highlight this file">${_ge(label)}</button>` : "";
  let contractBlock = "", etcBlock = "", moreBlock = "";
  if (kind === "current") {
    contractBlock = `<div class="rr-sec"><div class="rr-sec-title">2 · Contract ${rrMoney(x.contract)} ${rrPill(src.contract)}</div>
      <div class="rr-sec-line">${rrSrcLine(src.contract)}${src.contract?.file ? `<b>${_ge(src.contract.file.split("/").pop())}</b> ${fileBtn(src.contract.file, "Show file in folder")}` : ""}</div>
      ${rrPic(pics.contract, "Where the contract sits")}</div>`;
    etcBlock = `<div class="rr-sec"><div class="rr-sec-title">3 · ETC ${rrMoney(x.etc)} ${rrPill(src.etc)}</div>
      <div class="rr-sec-line">${rrSrcLine(src.etc)}${src.etc?.file ? `<b>${_ge(src.etc.file.split("/").pop())}</b> ${fileBtn(src.etc.file, "Show file in folder")}` : ""}</div>
      ${rrPic(pics.etc, "Where the ETC sits")}</div>`;
    moreBlock = (pics.master || pics.rpfile) ? `<details class="rr-more"><summary>The WIP master row and the RP file row</summary>${rrPic(pics.master, "WIP master")}${rrPic(pics.rpfile, "RP WIP file")}</details>` : "";
  } else {
    contractBlock = `<div class="rr-sec"><div class="rr-sec-title">2 · Why it left the WIP</div><div class="rr-sec-line">${_ge(x.why || "")}</div>${rrPic(pics.removed, "Removed log")}</div>`;
  }
  card.innerHTML = `<div class="rr-head"><span class="wr-pn">${_ge(x.line)}</span><span class="wr-name">${_ge(x.name || "")}</span><span class="wr-name">· ${_ge(x.builder || "")}</span>${x.rp_status ? `<span class="wr-badge changed" title="status in the RP WIP file">${_ge(x.rp_status)}</span>` : ""}<span class="rr-links">${pageBtn}${folderBtn}${qboBtn}${qboPl}${jtLink}</span></div>
    <div class="rr-strip">${nums}</div>
    ${flags.length ? `<div class="rr-flags">${flags.map(f => `<div>${_ge(f)}</div>`).join("")}</div>` : ""}
    ${schedBlock}${contractBlock}${etcBlock}${moreBlock}
    <div class="rr-sec"><div class="rr-sec-title">${kind === "current" ? "4" : "3"} · Change log</div><div class="rr-audit" data-audit="${_ge(x.line)}"></div></div>
    ${rrOursHtml(x, kind, m, due, false)}`;
  rrWireOurs(card, x, kind);
  return card;
}

// The CP / MFD card: no schedule (owner 2026-09-15: "CP will obv not have the schedule ... it's just the change
// order location, the takeoff contract etc"). The contract and approved COs from the draw G702, the ETC from the
// takeoff, costs and billed from QuickBooks - each with the document it came from and a Finder button; then the
// pending update (was -> now, where the new number comes from) when this update changes the line.
function rrCardDiv(x) {
  const kind = "current", m = rrMark(x, kind), due = rrIsDue(m);
  const card = document.createElement("div");
  card.className = "rr-card " + ((m && !due && m.decision === "confirmed") ? "rr-done" : (x.problem ? "rr-prob" : "rr-due"));
  const src = x.src || {};
  const pageBtn = `<button type="button" class="btn tiny" data-project="${_ge(x.line)}" title="This job's page in the ledger: invoices, draws, bills, costs">Project page</button>`;
  const wrBtn = (x.pending || []).length ? `<button type="button" class="btn tiny" data-wipreview="1" title="The WIP Review tab writes the approved changes to the tabs">Write in WIP Review</button>` : "";
  const cell = (l, v, pill) => `<span class="rr-n"><span class="l">${l}</span> <span class="v">${v}</span>${pill || ""}</span>`;
  const K = x.contract, C = x.cos, E = x.etc, tot = K != null ? K + (C || 0) : null;
  const rate = rrDiv === "MFD" ? 0.09 : 0.10;
  const gp = (tot != null && E != null) ? tot - E - (x.co_costs || 0) : null;
  const oh = tot != null ? tot * rate : null, net = (gp != null && oh != null) ? gp - oh : null;
  const pctTxt = (gp != null && tot) ? ` (${(gp / tot * 100).toFixed(1)}%)` : "";
  const profit = tot != null ? `<div class="rr-profit">Contract incl. COs <b>${money(tot)}</b> · gross profit (contract - ETC${x.co_costs ? " - CO costs" : ""}) <b class="${gp != null && gp < 0 ? "neg" : ""}">${gp == null ? "–" : money(gp)}</b>${pctTxt} · overhead ${Math.round(rate * 100)}% <b>${money(oh)}</b> · net <b class="${net != null && net < 0 ? "neg" : ""}">${net == null ? "–" : money(net)}</b></div>` : "";
  const nums = cell("Contract", rrMoney(K), rrPill(src.contract)) + cell("Approved COs", C == null ? "–" : money(C), rrPill(src.cos)) + cell("ETC", rrMoney(E), rrPill(src.etc))
    + cell("Costs", rrMoney(x.costs), rrPill(src.costs)) + cell("Billed", rrMoney(x.billed), rrPill(src.billed)) + (x.retainage ? cell("Retainage held", money(x.retainage), "") : "") + profit;
  const fileBtn = (f, label) => f ? `<button type="button" class="btn tiny" data-reveal="${_ge(f)}" title="Open the folder and highlight this file">${_ge(label)}</button>` : "";
  const secBlock = (n, title, v, sr) => `<div class="rr-sec"><div class="rr-sec-title">${n} · ${title} ${v} ${rrPill(sr)}</div>
      <div class="rr-sec-line">${rrSrcLine(sr)}${sr && sr.file ? `<b>${_ge(sr.file.split("/").pop())}</b> ${fileBtn(sr.file, "Show file in folder")}` : (sr && sr.detail ? "" : "<i>no document named</i>")}</div></div>`;
  const pend = x.pending || [];
  const pendBlock = pend.length ? `<div class="rr-sec"><div class="rr-sec-title">This update changes the line</div>
      <table class="rr-pend"><thead><tr><th></th><th class="n">On the WIP now</th><th class="dir"></th><th class="n">After this update</th><th>Where the new number comes from</th></tr></thead><tbody>
      ${pend.map(f => `<tr class="${f.reversed ? "rev" : ""}"><td>${_ge(f.label)}</td><td class="n was">${money(f.was)}</td><td class="dir">${wrDir(f.was, f.now)}</td><td class="n now">${money(f.now)}${f.reversed ? ' <span class="wr-mark rev">REVERSED</span>' : f.decreased ? ' <span class="wr-mark dec">decreased</span>' : ""}</td><td class="src">${_ge(f.source || "")}${f.note ? (f.source ? " · " : "") + `<i>${_ge(f.note)}</i>` : ""}</td></tr>`).join("")}
      </tbody></table></div>` : `<div class="rr-sec"><div class="rr-sec-line">This update leaves the line as it is on the WIP.</div></div>`;
  const badge = x.status && x.status !== "SAME" ? `<span class="wr-badge ${_ge(String(x.status).toLowerCase())}">${_ge(x.status)}</span>` : "";
  card.innerHTML = `<div class="rr-head"><span class="wr-pn">${_ge(x.line)}</span><span class="wr-name">${_ge(x.name || "")}</span><span class="wr-name">· ${_ge(x.builder || "")}</span>${badge}<span class="rr-links">${pageBtn}${wrBtn}</span></div>
    <div class="rr-strip">${nums}</div>
    ${(x.flags || []).length ? `<div class="rr-flags">${x.flags.map(f => `<div>${_ge(f)}</div>`).join("")}</div>` : ""}
    ${(x.notes || []).length ? `<div class="rr-sec-line rr-notes">${x.notes.map(_ge).join(" · ")}</div>` : ""}
    ${secBlock(1, "Contract", rrMoney(K), src.contract)}
    ${secBlock(2, "Approved change orders", C == null ? "–" : money(C), src.cos)}
    ${secBlock(3, "ETC", rrMoney(E), src.etc)}
    ${pendBlock}
    ${rrOursHtml(x, kind, m, due, true)}`;
  rrWireOurs(card, x, kind);
  return card;
}

let rrJustSaved = null;   // the line whose answer was just written and read back - tagged on the list

async function rrSave(card, x, kind, decision) {
  const val = f => { const el = card.querySelector(`[data-f="${f}"]`); return el ? el.value : ""; };
  const ok = f => { const on = card.querySelector(`.rr-ok[data-f="${f}"] button.on`); return on ? on.dataset.v : ""; };
  const body = { project_no: x.line, kind, decision, mode: rrMode, note: val("note"), our_contract: val("our_contract"), our_etc: val("our_etc"), our_cos: val("our_cos"), contract_ok: ok("contract_ok"), etc_ok: ok("etc_ok"), cos_ok: ok("cos_ok") };
  const saveBtn = card.querySelector("[data-save]"); if (saveBtn) { saveBtn.disabled = true; saveBtn.textContent = "Saving…"; }
  let res;
  try { res = await (await fetch("/api/review/mark", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) })).json(); }
  catch { res = { error: "no answer from the server" }; }
  if (!res || !res.ok) { if (saveBtn) { saveBtn.disabled = false; saveBtn.textContent = "Save"; } toast("NOT saved - " + ((res && res.error) || "could not reach the server"), 5000); return; }
  // Read it back from the server (owner 2026-09-15: "show me that it was saved instead of trusting that
  // i clicked it") - the list is rebuilt from what the database returns, not from what was typed.
  let fresh = null;
  try { fresh = await (await fetch(rrApi())).json(); } catch { fresh = null; }
  const key = `${x.line}|${kind}`;
  const stored = fresh && fresh.marks ? fresh.marks[key] : undefined;
  if (fresh && fresh.marks) RR.marks = fresh.marks;
  const verified = decision ? (stored && stored.at === res.at) : (stored === undefined);
  if (!verified) { if (saveBtn) { saveBtn.disabled = false; saveBtn.textContent = "Save"; } toast(`NOT saved - ${x.line} did not read back from the database`, 6000); return; }
  const label = RR_DEC_LABEL[decision] || (decision ? decision : "cleared");
  toast(`Saved ✓ ${x.line} · ${label} · ${rrWhoLow()} · ${fmtDate(res.at, true)}`, 4000);
  rrJustSaved = decision ? x.line : null;
  rrRenderStats();
  if (decision) { if (!_popViewIfOwn("rr")) { rrOpenLine = null; rrRenderCards(); window.scrollTo(0, 0); } }   // back to the list: the row shows the answer read back
  else { const fresh2 = rrDiv === "RP" ? rrCard(x, kind) : rrCardDiv(x); card.replaceWith(fresh2); }
}

async function rrSaveQuick(tr, x, kind, decision) {
  const m = rrMark(x, kind) || {};
  const keep = (mv, xv) => mv != null ? mv : (xv != null ? xv : "");
  const body = { project_no: x.line, kind, decision, mode: rrMode, note: m.note || "", our_contract: keep(m.our_contract, x.contract), our_etc: keep(m.our_etc, x.etc), our_cos: keep(m.our_cos, x.cos), contract_ok: m.contract_ok != null ? m.contract_ok : "", etc_ok: m.etc_ok != null ? m.etc_ok : "", cos_ok: m.cos_ok != null ? m.cos_ok : "" };
  const btn = tr.querySelector("[data-save]"); if (btn) { btn.disabled = true; btn.textContent = "Saving…"; }
  let res; try { res = await (await fetch("/api/review/mark", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) })).json(); } catch { res = { error: "no answer from the server" }; }
  if (!res || !res.ok) { if (btn) { btn.disabled = false; btn.textContent = "Save"; } toast("NOT saved - " + ((res && res.error) || "could not reach the server"), 5000); return; }
  let fresh = null; try { fresh = await (await fetch(rrApi())).json(); } catch { fresh = null; }
  const stored = fresh && fresh.marks ? fresh.marks[`${x.line}|${kind}`] : undefined;
  if (fresh && fresh.marks) RR.marks = fresh.marks;
  if (!(stored && stored.at === res.at)) { if (btn) { btn.disabled = false; btn.textContent = "Save"; } toast(`NOT saved - ${x.line} did not read back from the database`, 6000); return; }
  toast(`Saved ✓ ${x.line} · ${RR_DEC_LABEL[decision] || decision} · ${rrWhoLow()} · ${fmtDate(res.at, true)}`, 4000);
  rrJustSaved = x.line; rrLastLine = x.line; rrRenderStats(); rrRenderCards();
}

async function rrFinalize() {
  let res; try { res = await (await fetch("/api/rp/finalize", { method: "POST" })).json(); } catch { toast("could not build the package"); return; }
  if (!res.ok) { toast(res.error || "could not build the package"); return; }
  const p = res.package, body = $("#rrBody"); rrOpenLine = null;
  const li = arr => arr.length ? `<ul>${arr.map(a => `<li>${_ge(typeof a === "string" ? a : a.line + (a.field ? ` · ${a.field} ${money(a.from)} → ${money(a.to)}` : "") + (a.note ? ` · "${a.note}"` : ""))}</li>`).join("")}</ul>` : `<div class="hint" style="margin:2px 0 8px">none</div>`;
  body.innerHTML = `<div class="rr-final"><div class="rr-pagenav"><button type="button" class="btn small" id="rrFinalBack">← Back to the list</button><span class="rr-pagepos">Finalize package · ${fmtDate(p.generated, true)} · ${p.answered} line(s) answered</span></div>
    <h3>Numbers you changed (go to the WIP master through the guarded WIP writer)</h3>${li(p.changes)}
    <h3>Confirmed as they are (${p.confirmed.length})</h3>${li(p.confirmed)}
    <h3>Need a fix (${p.fixes.length})</h3>${li(p.fixes)}
    <h3>Finished - agreed done (${p.finished_agreed.length}) · keep on the WIP (${p.finished_keep.length})</h3>${li(p.finished_agreed)}${li(p.finished_keep)}
    <h3>Notes (${p.notes.length})</h3>${li(p.notes.map(n => ({ line: n.line, note: n.note })))}
    ${p.master_renamed ? `<p class="rr-warn">The master's RP tab is named "${_ge(p.master_tab)}"; the WIP writer writes "Test - RP". Rename it back before anything is written.</p>` : ""}
    <p class="hint" style="margin:10px 0 0">The package is on disk (finalize.json, decisions_rp.json, answers.txt). Nothing has been written to the master from here.</p></div>`;
  $("#rrFinalBack").onclick = () => rrRenderCards();
}

async function rrRebuild() {
  const what = rrDiv === "RP" ? "Rebuild the RP review? It re-reads the master, the RP file, every job folder, JobTread (Touch ID) and the schedules. A few minutes."
    : rrDiv === "CP" ? "Recompute the CP update? It reads every CP job folder on Common (draw G702s, takeoffs) and pulls costs / billed from QuickBooks (Touch ID). A few minutes; nothing is written."
    : "Recompute the MFD update? It reads the master's MFD rows and pulls costs / billed from QuickBooks (Touch ID). Nothing is written.";
  if (!confirm(what)) return;
  let res; try { res = await (await fetch("/api/review/refresh", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ confirm: true, div: rrDiv }) })).json(); }
  catch { toast("could not start"); return; }
  if (!res.ok) { toast(res.error || "could not start"); return; }
  const body = $("#rrBody"); $("#rrRebuild").disabled = true;
  body.innerHTML = `<div class="wr-run"><div class="wr-run-label">Building the ${rrDiv} review…</div><div class="pl-bar"><div class="pl-fill" id="rrFill"></div></div><div class="hint" id="rrRunHint">${rrDiv === "RP" ? "Reading the folders on Common and JobTread - a Touch ID prompt appears on the Mac." : "Running the WIP reader - a Touch ID prompt appears on the Mac."}</div></div>`;
  if (rrPoll) clearInterval(rrPoll);
  rrPoll = setInterval(async () => {
    let s; try { s = await (await fetch("/api/sync/status")).json(); } catch { return; }
    const fill = $("#rrFill"); if (fill) fill.style.width = s.state === "running" ? "60%" : "100%";
    if (s.state !== "running") { clearInterval(rrPoll); rrPoll = null; $("#rrRebuild").disabled = false; if (s.state === "error") toast("build failed - see the sync log"); loadReview(true); }
  }, 1500);
}
