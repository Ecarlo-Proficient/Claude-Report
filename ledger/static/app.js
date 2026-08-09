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
  { key: "qbo_margin",            label: "Margin (QBO)",  type: "money" },
  { key: "qbo_margin_pct",        label: "Margin %",      type: "pct" },
  { key: "subs_pct",              label: "Subs %",        type: "pct" },
];

// Derived per-job metrics from the REAL QBO costs — computed once at load time.
// Only for jobs that actually have costs loaded; others stay null (blank).
// margin here = billed − QBO cost (a billed-basis margin-to-date, labeled as such).
function deriveMetrics(r) {
  const cost = r.costs_loaded, etc = r.estimated_total_costs, billed = r.billed_to_date, subs = r.sub_costs;
  const has = cost !== null && cost !== undefined;
  r.budget_burn   = has && etc ? cost / etc : null;
  r.qbo_margin    = has && billed != null ? billed - cost : null;
  r.qbo_margin_pct = (r.qbo_margin != null && billed) ? r.qbo_margin / billed : null;
  r.subs_pct      = has && cost && subs != null ? subs / cost : null;
}

// ── Settings ──────────────────────────────────────────────────────────────
const LS_KEY = "proficient-ledger-settings-v1";
const DEFAULTS = {
  theme: "auto", accent: "#3E7A5C", font: "system", fontSize: 14,
  density: "comfortable", width: "full",
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

function loadSettings() {
  try {
    const s = JSON.parse(localStorage.getItem(LS_KEY));
    if (!s) return structuredClone(DEFAULTS);
    return { ...structuredClone(DEFAULTS), ...s,
             widgets: { ...DEFAULTS.widgets, ...(s.widgets || {}) },
             columns: Array.isArray(s.columns) && s.columns.length ? s.columns : DEFAULTS.columns };
  } catch { return structuredClone(DEFAULTS); }
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
  $("#widget-ap").hidden        = !settings.widgets.ap;
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
let COST = { by_code: [], by_project_code: {}, by_project: {}, by_cost_type: [], loaded_total: 0 };
let costCollapsed = new Set();   // collapsed cost-type parents in the Costs widget
let meta = {};
let sortKey = "total_contract_price";
let sortDir = -1;   // -1 desc, 1 asc
let activeRule = null;   // key of a RULES entry currently filtering the table

// ── Load ──────────────────────────────────────────────────────────────────
async function load() {
  let data;
  try { data = await (await fetch("/api/data")).json(); }
  catch (e) { return showError("Could not reach the server: " + e); }
  if (data.error) return showError(data.error);
  $("#errorBanner").hidden = true;
  ALL = data.projects || [];
  ALL.forEach(deriveMetrics);
  AP = data.ap || { summary: {}, lien_watch: [], by_project: {} };
  COST = data.cost || { by_code: [], by_project_code: {}, by_project: {}, by_cost_type: [], loaded_total: 0 };
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
function render() { renderKPIs(); renderAttention(); renderAP(); renderCosts(); renderMargins(); renderDivisions(); renderProjects(); }

function renderMargins() {
  // portfolio, over jobs that actually have QBO costs loaded
  const loaded = ALL.filter(r => r.costs_loaded != null);
  const cost = loaded.reduce((t, r) => t + num(r.costs_loaded), 0);
  const billed = loaded.reduce((t, r) => t + num(r.billed_to_date), 0);
  const subs = loaded.reduce((t, r) => t + num(r.sub_costs), 0);
  const margin = billed - cost;
  const overBudget = loaded.filter(isOverBudget).length;
  $("#marginNote").textContent = loaded.length ? `(${loaded.length} jobs with QBO costs)` : "(no cost data — run load_costs.py)";
  const stats = [
    ["Margin to date", money(margin), "billed − QBO cost"],
    ["Portfolio margin %", billed ? (margin / billed * 100).toFixed(1) + "%" : "—", "of billed"],
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
    ? `($${Math.round(total).toLocaleString()} · cost type ▸ job type)`
    : "(no cost data — run load_costs.py)";
  renderCostMix(groups, total);
  const cols = [["Cost type  ▸  job type", "left"], ["Code", "left"], ["Actual", "right"], ["% of total", "right"], ["Lines", "right"]];
  const thead = $("#costTable thead"), tbody = $("#costTable tbody");
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

function renderAP() {
  const s = AP.summary || {};
  const stats = [
    ["Open AP", money(s.open_balance || 0), `${s.open_lines || 0} open bills`],
    ["Lien deadlines", String(s.watch_count || 0), "bills on the clock"],
    ["Past due", String((AP.lien_watch || []).filter(r => r.lien_status === "Notice PAST due").length), "notice past due"],
  ];
  const sr = $("#apStats"); sr.innerHTML = "";
  for (const [label, value, sub] of stats) {
    const el = document.createElement("div"); el.className = "kpi";
    el.innerHTML = `<div class="k-label"></div><div class="k-value"></div><div class="k-sub"></div>`;
    el.querySelector(".k-label").textContent = label;
    el.querySelector(".k-value").textContent = value;
    el.querySelector(".k-sub").textContent = sub;
    sr.appendChild(el);
  }
  const watch = AP.lien_watch || [];
  $("#apCount").textContent = watch.length ? `(${watch.length} on the lien clock)` : "";
  const cols = [["Lien", "left"], ["Project", "left"], ["Vendor", "left"], ["Bill #", "left"], ["Open Bal", "right"]];
  const thead = $("#lienTable thead"), tbody = $("#lienTable tbody");
  thead.innerHTML = ""; tbody.innerHTML = "";
  const htr = document.createElement("tr");
  for (const [c, al] of cols) { const th = document.createElement("th"); if (al === "left") th.className = "left"; th.textContent = c; htr.appendChild(th); }
  thead.appendChild(htr);
  const known = new Set(ALL.map(r => r.project_no));
  for (const r of watch.slice(0, 60)) {
    const tr = document.createElement("tr");
    if (r.project_no && known.has(r.project_no)) tr.onclick = (e) => { if (!e.target.closest(".cell")) openDetail(ALL.find(x => x.project_no === r.project_no)); };
    const lienTd = document.createElement("td"); lienTd.className = "left";
    const pill = document.createElement("span"); pill.className = "lien " + (LIEN_CLASS[r.lien_status] || "info"); pill.textContent = r.lien_status;
    lienTd.appendChild(pill); tr.appendChild(lienTd);
    tr.appendChild(leftText(r.project_no || "—"));
    tr.appendChild(leftText(r.vendor || "—"));
    tr.appendChild(leftText(r.bill_ref || "—"));
    const ob = document.createElement("td"); ob.appendChild(moneyCell(r.open_balance)); tr.appendChild(ob);
    tbody.appendChild(tr);
  }
}
function leftText(v) { const td = document.createElement("td"); td.className = "left"; const s = document.createElement("span"); s.textContent = v; td.appendChild(s); return td; }

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
    addRow(tbody, [textCell(d, true), textCell(String(g.jobs)), moneyCell(g.contract),
      moneyCell(g.costs), moneyCell(g.billed), moneyCell(g.over), moneyCell(g.under)]);
  }
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
      ["Budget burn (cost ÷ ETC)", pct(r.budget_burn)],
      ["Margin to date (billed − cost)", money(r.qbo_margin)],
      ["Margin %", pct(r.qbo_margin_pct)],
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
  $("#wAp").checked = settings.widgets.ap;
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
  on("#wAp", "change", e => { settings.widgets.ap = e.target.checked; saveSettings(); applySettings(); });
  on("#wCosts", "change", e => { settings.widgets.costs = e.target.checked; saveSettings(); applySettings(); });
  on("#wMargins", "change", e => { settings.widgets.margins = e.target.checked; saveSettings(); applySettings(); });
  on("#wDivisions", "change", e => { settings.widgets.divisions = e.target.checked; saveSettings(); applySettings(); });
  on("#wProjects", "change", e => { settings.widgets.projects = e.target.checked; saveSettings(); applySettings(); });
  on("#btnReset", "click", () => { settings = structuredClone(DEFAULTS); saveSettings(); applySettings(); syncSettingsUI(); render(); toast("Reset to defaults"); });
}

// ── Wire up ───────────────────────────────────────────────────────────────
function init() {
  applySettings();
  syncSettingsUI();
  wireSettings();
  ["#search", "#fDivision", "#fStatus", "#fCategory", "#fActive"].forEach(sel =>
    $(sel).addEventListener("input", renderProjects));
  $("#btnExport").onclick = exportCSV;
  $("#btnRefresh").onclick = load;
  $("#btnClearRule").onclick = () => { activeRule = null; renderAttention(); renderProjects(); };
  $("#btnSettings").onclick = () => openPanel("#settings");
  $("#btnCloseSettings").onclick = closePanels;
  $("#btnCloseDetail").onclick = closePanels;
  $("#btnCopyDetail").onclick = () => copy(detailAsText());
  $("#overlay").onclick = closePanels;
  document.addEventListener("keydown", e => { if (e.key === "Escape") closePanels(); });
  matchMedia("(prefers-color-scheme: dark)").addEventListener("change", () => { if (settings.theme === "auto") applySettings(); });
  load();
}
init();
