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
  density: "comfortable", width: "boxed",
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
  root.style.setProperty("--maxw", settings.width === "boxed" ? "1180px" : "100%");
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
let AP = { summary: {}, lien_watch: [], by_project: {} };
let COST = { by_code: [], by_project_code: {}, by_project: {}, by_cost_type: [], by_vendor: [], loaded_total: 0 };
let DRAWS = { draws: [], total: 0 };
let SALES = { pipeline: [], by_rep: [], warm: [], customers: [], totals: {} };
let costCollapsed = new Set();   // collapsed cost-type parents (default: all collapsed)
let drawsCollapsed = new Set();  // collapsed draw cards (default: all collapsed)

// ── Tabs ────────────────────────────────────────────────────────────────────
let activeTab = "overview";
function setTab(t) {
  activeTab = t;
  try { localStorage.setItem("proficient-ledger-tab", t); } catch { /* ignore */ }
  $$(".tab-page").forEach(p => { p.hidden = p.dataset.tab !== t; });
  $$(".tab").forEach(b => b.classList.toggle("active", b.dataset.tabbtn === t));
  window.scrollTo(0, 0);
}
const nameOf = pn => (ALL.find(r => r.project_no === pn) || {}).project_name || "";
let meta = {};
let sortKey = "total_contract_price";
let sortDir = -1;   // -1 desc, 1 asc
let activeRule = null;   // key of a RULES entry currently filtering the table
let activeLien = null;   // lien stage currently filtering the Liens table (null = all)
let activeDrawStage = null; // draw stage currently filtering the Draws list (null = all)

// ── Load ──────────────────────────────────────────────────────────────────
async function load(isAuto) {
  let data;
  try { data = await (await fetch("/api/data")).json(); }
  catch (e) { return showError("Could not reach the server: " + e); }
  if (data.error) return showError(data.error);
  $("#errorBanner").hidden = true;
  ALL = data.projects || [];
  ALL.forEach(deriveMetrics);
  AP = data.ap || { summary: {}, lien_watch: [], by_project: {} };
  COST = data.cost || { by_code: [], by_project_code: {}, by_project: {}, by_cost_type: [], by_vendor: [], loaded_total: 0 };
  DRAWS = data.draws || { draws: [], total: 0 };
  SALES = data.sales || { pipeline: [], by_rep: [], warm: [], customers: [], totals: {} };
  // Big-picture first: collapse everything by default; the user expands to zoom in.
  // On a live auto-refresh, preserve what the user has already expanded.
  if (!isAuto) {
    costCollapsed = new Set((COST.by_cost_type || []).map(g => g.parent));
    drawsCollapsed = new Set((DRAWS.draws || []).map(d => d.matched_invoice));
  }
  meta = data.meta || {};
  $("#metaLine").textContent =
    `${meta.project_count} projects · report ${meta.report_date || "—"}` +
    (meta.loaded_at ? ` · loaded ${meta.loaded_at}` : "");
  buildFilterOptions();
  render();
}
function showError(msg) {
  const b = $("#errorBanner"); b.hidden = false; b.textContent = msg;
  $("#metaLine").textContent = "not loaded";
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
    if (f.activeOnly && (r.status || "").toLowerCase() !== "active") return false;
    if (f.q) {
      const hay = [r.project_no, r.project_name, r.builder_or_gc].filter(Boolean).join(" ").toLowerCase();
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
  renderProjects(); renderLiens(); renderVendors(); renderDraws(); renderSales();
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
    el.querySelector(".f-when").textContent = when ? when.replace("T", " ") : "—";
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
  const collectDraws = (DRAWS.draws || []).filter(d => d.stage === "Paid — collect waivers").length;
  const overB = ALL.filter(isOverBudget).length;
  const underB = ALL.filter(r => num(r.underbillings) > 0).length;
  const goRule = (key) => { setTab("overview"); activeRule = key; renderAttention(); renderProjects(); $("#btnClearRule").hidden = false; window.scrollTo(0, 0); };
  const acts = [
    ["Liens past due", pastDue, true, () => setTab("liens")],
    ["Lien due ≤7d", dueSoon, true, () => setTab("liens")],
    ["Draws: collect waivers", collectDraws, false, () => setTab("draws")],
    ["Draws ready to turn in", readyDraws, false, () => setTab("draws")],
    ["Over budget", overB, true, () => goRule("overbudget")],
    ["Underbilled (can invoice)", underB, false, () => goRule("underbilled")],
  ];
  const ar = $("#homeActions"); ar.innerHTML = "";
  for (const [label, n, warn, go] of acts) {
    const el = document.createElement("div"); el.className = "action" + (warn && n ? " warn" : "") + (n ? "" : " none");
    el.innerHTML = `<span class="a-n"></span><span class="a-lab"></span>`;
    el.querySelector(".a-n").textContent = n;
    el.querySelector(".a-lab").textContent = label;
    if (n) el.onclick = go;
    ar.appendChild(el);
  }
  // ── working on (active projects) ──
  const sel = $("#homeDivision");
  if (sel && sel.options.length <= 1) for (const d of uniq(ALL.map(r => r.division))) { const o = document.createElement("option"); o.value = d; o.textContent = d; sel.appendChild(o); }
  const div = sel ? sel.value : "";
  const active = ALL.filter(r => (r.status || "").toLowerCase() === "active" && (!div || r.division === div))
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
  "Fund in — pay vendors": "d7", "Paid — collect waivers": "d15",
  "Awaiting GC funding": "info", "Ready to turn in": "ready",
};
// Clearer, direction-explicit pill text (who paid whom). Display only — the internal
// stage keys above are unchanged (they're matched in several places).
const DRAW_STAGE_LABEL = {
  "Fund in — pay vendors": "GC funded → pay vendors",
  "Paid — collect waivers": "Vendors paid → collect waivers",
  "Awaiting GC funding": "Awaiting GC funding",
  "Ready to turn in": "Ready to turn in",
};
function renderDraws() {
  const f = { q: ($("#drawSearch") ? $("#drawSearch").value : "").trim().toLowerCase(),
              div: $("#drawDivision") ? $("#drawDivision").value : "" };
  const all = (DRAWS.draws || []).filter(d => {
    if (f.div && !String(d.project_no || "").toUpperCase().startsWith(f.div)
              && !(d.label || "").toUpperCase().includes("— " + f.div)) return false;
    if (f.q) {
      const hay = [d.label, d.project_no].concat((d.bills || []).map(b => b.vendor))
        .filter(Boolean).join(" ").toLowerCase();
      if (!hay.includes(f.q)) return false;
    }
    return true;
  });
  const shown = activeDrawStage ? all.filter(d => d.stage === activeDrawStage) : all;
  $("#drawsNote").textContent = (DRAWS.draws || []).length
    ? `(${shown.length} shown of ${DRAWS.total} · most recent first)`
    : "(no draw data — run load_bill_tracker.py)";
  // Clickable stage tiles → filter the draw list. Counts come from `all` (all stages);
  // subs spell out the money direction (GC pays us in → we pay vendors out → waivers).
  const stats = [
    ["Ready to turn in", "Ready to turn in", "all paid + waivers in"],
    ["Collect waivers", "Paid — collect waivers", "vendors paid — waivers pending"],
    ["Pay vendors", "Fund in — pay vendors", "GC funded — vendors not paid yet"],
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
  const box = $("#drawList"); box.innerHTML = "";
  for (const d of shown) {
    const collapsed = drawsCollapsed.has(d.matched_invoice);
    const sec = document.createElement("section"); sec.className = "widget draw";
    const head = document.createElement("div"); head.className = "widget-head draw-head"; head.style.cursor = "pointer";
    head.onclick = () => { collapsed ? drawsCollapsed.delete(d.matched_invoice) : drawsCollapsed.add(d.matched_invoice); renderDraws(); };
    const left = document.createElement("div");
    const h = document.createElement("h2"); h.textContent = (collapsed ? "▸ " : "▾ ") + (d.label || d.matched_invoice); left.appendChild(h);
    const meta2 = document.createElement("div"); meta2.className = "panel-sub";
    meta2.textContent = `${money(d.total)} · ${d.n} bills · ${d.paid}/${d.n} paid · ${d.waivers}/${d.n} waivers`;
    left.appendChild(meta2); head.appendChild(left);
    const pill = document.createElement("span"); pill.className = "lien " + (DRAW_STAGE_CLASS[d.stage] || "info"); pill.textContent = DRAW_STAGE_LABEL[d.stage] || d.stage;
    head.appendChild(pill);
    if (d.action && d.action.url) {
      const a = document.createElement("a"); a.className = "notion-link"; a.href = d.action.url;
      a.target = "_blank"; a.rel = "noopener"; a.textContent = "📄 Notion · " + (d.action.status || "Open");
      a.onclick = (e) => e.stopPropagation();
      head.appendChild(a);
    }
    sec.appendChild(head);
    if (!collapsed) {
      const scroll = document.createElement("div"); scroll.className = "table-scroll";
      const table = document.createElement("table"); table.className = "grid";
      const thead = document.createElement("thead"), tbody = document.createElement("tbody");
      const cols = [["Vendor", "left"], ["Bill #", "left"], ["Amount", "right"], ["Paid", "left"], ["GC funded", "left"], ["Waiver in hand", "left"]];
      const htr = document.createElement("tr");
      for (const [c, al] of cols) { const th = document.createElement("th"); if (al === "left") th.className = "left"; th.textContent = c; htr.appendChild(th); }
      thead.appendChild(htr);
      for (const b of d.bills) {
        const tr = document.createElement("tr");
        tr.appendChild(leftText(b.vendor || "—"));
        tr.appendChild(leftText(b.bill_ref || "—"));
        const av = document.createElement("td"); av.appendChild(moneyCell(b.amount)); tr.appendChild(av);
        tr.appendChild(leftText(b.pay_date ? "✓ " + b.pay_date : "—"));
        tr.appendChild(leftText(b.gc_paid ? "✓ " + b.gc_paid : "—"));
        const wtd = document.createElement("td"); wtd.className = "left";
        const lab = document.createElement("label"); lab.className = "chk";
        const cb = document.createElement("input"); cb.type = "checkbox"; cb.checked = !!b.waiver;
        cb.onchange = () => setWaiver(d, b, cb);
        lab.appendChild(cb); lab.appendChild(document.createTextNode(b.waiver ? " in hand" : " mark"));
        wtd.appendChild(lab); tr.appendChild(wtd);
        tbody.appendChild(tr);
      }
      table.appendChild(thead); table.appendChild(tbody); scroll.appendChild(table);
      sec.appendChild(scroll);
    }
    box.appendChild(sec);
  }
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
    draw.waivers = draw.bills.filter(b => b.waiver).length;
    if (draw.stage !== "Ready to turn in" && draw.funded && draw.paid === draw.n && draw.waivers === draw.n) draw.stage = "Ready to turn in";
    else if (draw.waivers < draw.n && draw.paid === draw.n && draw.funded) draw.stage = "Paid — collect waivers";
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

function renderLiens() {
  const s = AP.summary || {};
  const watch = AP.lien_watch || [];
  $("#liensNote").textContent = watch.length ? `(${watch.length} bills on the clock)` : "(no AP data — run load_bill_tracker.py)";
  const pastDue = watch.filter(r => r.lien_status === "Notice PAST due");
  // ── summary KPIs ──
  const stats = [
    ["Open AP", money(s.open_balance || 0), `${s.open_lines || 0} open bills`],
    ["On the lien clock", String(s.watch_count || 0), "need action"],
    ["Past due", String(pastDue.length), money(pastDue.reduce((t, r) => t + num(r.open_balance), 0)) + " open"],
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
  // ── clickable stage tiles (the widgets) → filter the one table below ──
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
  tile(null, "All on the clock", watch.length, watch.reduce((t, r) => t + num(r.open_balance), 0), "", activeLien === null);
  for (const status of LIEN_ORDER) {
    const rows = byStatus[status]; if (!rows || !rows.length) continue;
    tile(status, LIEN_SHORT[status] || status, rows.length,
         rows.reduce((t, r) => t + num(r.open_balance), 0), LIEN_CLASS[status] || "info", activeLien === status);
  }
  $("#btnClearLien").hidden = activeLien === null;

  // ── the one table below — filtered by the active stage + the search box ──
  const q = ($("#lienSearch") ? $("#lienSearch").value : "").trim().toLowerCase();
  const known = new Set(ALL.map(r => r.project_no));
  const base = activeLien ? (byStatus[activeLien] || []) : watch;
  const enriched = base.map(r => {
    const d = splitDraw(r.matched_invoice);
    return { r, draw: r.invoice_no || d.draw || "", name: nameOf(r.project_no) || d.name || "" };
  });
  const shown = q ? enriched.filter(({ r, draw, name }) =>
    [r.project_no, draw, name, r.vendor, r.bill_ref].filter(Boolean).join(" ").toLowerCase().includes(q)) : enriched;

  // CP # · Draw # · Name/Address · Invoice # · Amount lead (owner's order); the
  // vendor trails and urgency is the coloured row edge so CP # stays first.
  const cols = [["CP #", "left"], ["Draw #", "left"], ["Name / Address", "left"],
                ["Invoice #", "left"], ["Amount", "right"], ["Vendor", "left"]];
  const thead = $("#lienTable thead"), tbody = $("#lienTable tbody");
  thead.innerHTML = ""; tbody.innerHTML = "";
  const htr = document.createElement("tr");
  for (const [c, al] of cols) { const th = document.createElement("th"); if (al === "left") th.className = "left"; th.textContent = c; htr.appendChild(th); }
  thead.appendChild(htr);
  for (const { r, draw, name } of shown) {
    const tr = document.createElement("tr");
    tr.className = "lien-row u-" + (LIEN_CLASS[r.lien_status] || "info");
    tr.title = r.lien_status || "";
    if (r.project_no && known.has(r.project_no)) tr.onclick = (e) => { if (!e.target.closest(".cell")) openDetail(ALL.find(x => x.project_no === r.project_no)); };
    tr.appendChild(leftText(r.project_no || "—"));
    tr.appendChild(leftText(draw || "—"));
    tr.appendChild(leftText(name || "—"));
    const inv = document.createElement("td"); inv.className = "left";
    const chip = document.createElement("span"); chip.className = "invno"; chip.textContent = r.bill_ref || "—"; inv.appendChild(chip); tr.appendChild(inv);
    const amt = document.createElement("td"); const mc = moneyCell(r.open_balance); mc.classList.add("lien-amt"); amt.appendChild(mc); tr.appendChild(amt);
    tr.appendChild(leftText(r.vendor || "—"));
    tbody.appendChild(tr);
  }
  if (!shown.length) {
    const tr = document.createElement("tr"); const td = document.createElement("td");
    td.colSpan = cols.length; td.className = "left"; td.style.color = "var(--text-dim)";
    td.textContent = watch.length ? "No bills match this stage / search." : "No AP data — run load_bill_tracker.py.";
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
  for (const v of vends.slice(0, 150)) {
    const tr = document.createElement("tr");
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
    tr.appendChild(leftText(r.rep));
    tr.appendChild(rightText(String(r.worked || 0)));
    tr.appendChild(rightText(String(r.contacted || 0)));
    tr.appendChild(rightText(String(r.interested || 0)));
    tr.appendChild(rightText(String(r.won || 0)));
    tb.appendChild(tr);
  }

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
    when.textContent = a.last_contacted ? `last ${a.last_contacted}${d !== null ? ` · ${d}d ago` : ""}` : "no contact date";
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
    tr.appendChild(leftText(c.last_contacted || "—"));
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
        active = rows.filter(r => (r.status || "").toLowerCase() === "active").length;
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
  wrap.className = "cell bar";
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
  const h = document.createElement("h4"); h.textContent = "P&L (project-pnl)"; g.appendChild(h);
  const row = document.createElement("div"); row.className = "drow";
  const dk = document.createElement("span"); dk.className = "dk"; dk.textContent = "Last pulled";
  const dv = document.createElement("span"); dv.className = "dv"; dv.textContent = "checking…";
  row.appendChild(dk); row.appendChild(dv); g.appendChild(row);
  const acts = document.createElement("div"); acts.className = "pnl-actions";
  const openBtn = document.createElement("button"); openBtn.className = "btn small"; openBtn.textContent = "Open"; openBtn.disabled = true;
  const genBtn = document.createElement("button"); genBtn.className = "btn small"; genBtn.textContent = "Generate / Refresh";
  acts.appendChild(openBtn); acts.appendChild(genBtn); g.appendChild(acts);
  const msg = document.createElement("div"); msg.className = "pnl-msg"; g.appendChild(msg);

  const refresh = () => fetch(`/api/pnl?proj=${encodeURIComponent(proj)}`).then(r => r.json()).then(d => {
    if (d.error) { dv.textContent = "—"; return; }
    if (d.exists) {
      dv.textContent = `${timeAgo(d.mtime)} · ${(d.mtime || "").replace("T", " ")}`;
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
    if (s.state === "running") { msg.textContent = `Generating… (${s.elapsed || 0}s) — Touch ID may be waiting.`; setTimeout(tick, 2000); }
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
      dv.textContent = fmt({ type }, r[k]);
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
    for (const [k, label, type] of present) lines.push(`  ${label}: ${fmt({ type }, r[k])}`);
  }
  if (r.notes) lines.push("", "NOTES", "  " + r.notes);
  return lines.join("\n");
}

// ── Panels ────────────────────────────────────────────────────────────────
function openPanel(sel) { $("#overlay").hidden = false; $(sel).hidden = false; }
function closePanels() { $("#overlay").hidden = true; $("#detail").hidden = true; $("#settings").hidden = true; }

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
function init() {
  applySettings();
  syncSettingsUI();
  wireSettings();
  ["#search", "#fDivision", "#fStatus", "#fCategory", "#fActive"].forEach(sel =>
    $(sel).addEventListener("input", renderProjects));
  ["#drawSearch", "#drawDivision"].forEach(sel => { const el = $(sel); if (el) el.addEventListener("input", renderDraws); });
  { const el = $("#vendorSearch"); if (el) el.addEventListener("input", renderVendors); }
  { const el = $("#lienSearch"); if (el) el.addEventListener("input", renderLiens); }
  ["#salesSearch", "#salesStage", "#salesDivision"].forEach(sel => { const el = $(sel); if (el) el.addEventListener("input", renderSales); });
  { const el = $("#btnClearLien"); if (el) el.onclick = () => { activeLien = null; renderLiens(); }; }
  { const el = $("#btnClearDrawStage"); if (el) el.onclick = () => { activeDrawStage = null; renderDraws(); }; }
  { const el = $("#homeDivision"); if (el) el.addEventListener("input", renderHome); }
  $("#btnExport").onclick = exportCSV;
  $("#btnRefresh").onclick = load;
  $("#btnClearRule").onclick = () => { activeRule = null; renderAttention(); renderProjects(); };
  $("#btnSettings").onclick = () => openPanel("#settings");
  $$(".tab").forEach(b => { b.onclick = () => setTab(b.dataset.tabbtn); });
  let savedTab = "home";
  try { savedTab = localStorage.getItem("proficient-ledger-tab") || "home"; } catch { /* ignore */ }
  setTab(savedTab);
  setInterval(() => load(true), 90000);   // live: soft auto-refresh (preserves expand state + tab)
  $("#btnCloseSettings").onclick = closePanels;
  $("#btnCloseDetail").onclick = closePanels;
  $("#btnCopyDetail").onclick = () => copy(detailAsText());
  $("#overlay").onclick = closePanels;
  document.addEventListener("keydown", e => { if (e.key === "Escape") closePanels(); });
  matchMedia("(prefers-color-scheme: dark)").addEventListener("change", () => { if (settings.theme === "auto") applySettings(); });
  load();
}
init();
