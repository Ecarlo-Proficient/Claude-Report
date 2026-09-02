// trail.js - "Where every dollar went": every QBO line behind a project's Costs / Billed,
// with a running total (the red line against the budget). Full-page record view, read-only.
// Data: GET /api/trail?project=<PN>&kind=costs|billed|both (ledger/trail.py). Own formatters
// on purpose (this page is 15px, tabular, negatives in red parentheses like Excel).
(function () {
  "use strict";
  const $t = (s, r) => (r || document).querySelector(s);
  const esc = s => String(s == null ? "" : s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  const tm = v => { if (v == null || v === "" || Number.isNaN(Number(v))) return "–";
    const n = Number(v), s = "$" + Math.round(Math.abs(n)).toLocaleString(); return n < 0 ? `(${s})` : s; };
  const tmc = v => `<span class="tr-money${Number(v) < 0 ? " neg" : ""}">${tm(v)}</span>`;
  const d8 = v => { if (!v) return "–"; const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(v); return m ? `${m[2]}/${m[3]}/${m[1]}` : String(v); };
  const dt8 = v => { if (!v) return "–"; const d = new Date(v); if (isNaN(d)) return d8(v);
    let h = d.getHours(), ap = h >= 12 ? "PM" : "AM"; h = h % 12 || 12;
    return `${String(d.getMonth() + 1).padStart(2, "0")}/${String(d.getDate()).padStart(2, "0")}/${d.getFullYear()} ${h}:${String(d.getMinutes()).padStart(2, "0")} ${ap}`; };
  const monthKey = v => (v || "").slice(0, 7);
  const monthLabel = k => { const [y, m] = k.split("-"); const n = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][Number(m) - 1];
    return n ? `${n} ${y}` : k; };

  const state = { pn: null, data: null, kind: "both", q: "", sort: { k: "date", dir: 1 } };

  // ── open / close ───────────────────────────────────────────────────────
  window.openTrail = async function (pn) {
    state.pn = pn; state.data = null; state.q = ""; state.sort = { k: "date", dir: 1 };
    if (typeof openRecord === "function") openRecord(`${pn} · where every dollar went`, "loading…");
    const host = $t("#recordBody") || $t("#recordView .record-body");
    if (!host) return;
    host.innerHTML = `<div class="tr-note">Loading the lines…</div>`;
    try {
      const r = await fetch(`/api/trail?project=${encodeURIComponent(pn)}&kind=both`);
      if (r.status === 404) { host.innerHTML = `<div class="tr-note">No line data for ${esc(pn)} yet - run the costs / invoices sync (Overview › Data freshness › Resync).</div>`; return; }
      state.data = await r.json();
    } catch (e) { host.innerHTML = `<div class="tr-note">Line data not available (${esc(e.message || e)}).</div>`; return; }
    render();
  };

  // ── render ─────────────────────────────────────────────────────────────
  function visibleLines() {
    const d = state.data; if (!d) return [];
    let rows = d.lines.filter(l => state.kind === "both" || (state.kind === "costs" ? l.kind === "cost" : l.kind === "billed"));
    if (state.q) { const q = state.q.toLowerCase();
      rows = rows.filter(l => [l.party, l.doc_number, l.memo, l.description, l.cost_code, l.account, l.txn_id].some(x => x && String(x).toLowerCase().includes(q))); }
    const { k, dir } = state.sort;
    const val = l => k === "amount" || k === "running_total" ? Number(l[k] || 0) : String(l[k] || "");
    rows = rows.slice().sort((a, b) => { const x = val(a), y = val(b); return (x < y ? -1 : x > y ? 1 : 0) * dir || String(a.date).localeCompare(String(b.date)); });
    return rows;
  }

  function render() {
    const d = state.data, host = $t("#recordBody") || $t("#recordView .record-body"); if (!d || !host) return;
    const sub = $t("#recordSub"); if (sub) sub.textContent = `${d.name || ""} · QBO loaded ${dt8(d.as_of.qbo_loaded_at)} · WIP report ${d8(d.as_of.wip_report_date)}`;
    const t = d.totals, b = d.budget;
    const pair = (label, qbo, qboNote, wip, delta, gapNote) => {
      const ok = delta != null && Math.abs(delta) < 1;
      const cls = delta == null ? "" : ok ? "ok" : "gap";
      return `<div class="tr-pair ${cls}">
        <div class="tr-fig"><div class="tr-lab">${label} - QuickBooks</div><div class="tr-val">${tm(qbo)}</div><div class="tr-as">${qboNote}</div></div>
        <div class="tr-fig"><div class="tr-lab">${label} - WIP report ${d8(d.as_of.wip_report_date)}</div><div class="tr-val">${wip == null ? "–" : tm(wip)}</div>
          <div class="tr-as">${delta == null ? "no WIP figure" : ok ? "match" : `gap ${tm(delta)} (QuickBooks minus WIP)${gapNote}`}</div></div>
      </div>`;
    };
    const afterNote = t.costs_after_report ? ` - ${tm(t.costs_after_report)} of it is lines dated after the ${d8(d.as_of.wip_report_date)} report` : "";
    const unexpl = t.delta_costs_unexplained != null && Math.abs(t.delta_costs_unexplained) >= 1 && t.costs_after_report ? `, ${tm(t.delta_costs_unexplained)} not explained by dates` : "";
    const kindBtn = (k, lab) => `<button class="btn small tr-kind${state.kind === k ? " on" : ""}" data-kind="${k}">${lab}</button>`;
    host.innerHTML = `
      <div class="tr-chips">
        <span class="tr-chip">Contract ${tm(b.contract)}</span>
        <span class="tr-chip">ETC (budget) ${tm(b.etc)}</span>
        <span class="tr-chip">QBO loaded ${dt8(d.as_of.qbo_loaded_at)}</span>
        <span class="tr-chip">WIP report ${d8(d.as_of.wip_report_date)}</span>
      </div>
      <div class="tr-pairs">
        ${state.kind !== "billed" ? pair("Costs to date", t.costs, "sum of the cost lines below", t.wip_costs_to_date, t.delta_costs, afterNote + unexpl) : ""}
        ${state.kind !== "costs" ? pair("Billed to date", t.billed_gross, `invoices ${tm(t.billed)} + retainage held ${tm(t.wip_retainage_held)} (WIP billed is gross)`, t.wip_billed_to_date, t.delta_billed, "") : ""}
      </div>
      <section class="widget tr-chartbox"><div class="widget-head"><h2>The red line <span class="count">running total against the budget and the contract</span></h2></div><div id="trChart"></div></section>
      <section class="widget">
        <div class="widget-head tr-tools">
          <span class="seg">${kindBtn("both", "Both")}${kindBtn("costs", "Costs")}${kindBtn("billed", "Billed")}</span>
          <input class="tr-search" id="trSearch" placeholder="Find a vendor, bill #, memo, code…" value="${esc(state.q)}">
          <span class="tr-spacer"></span>
          <button class="btn small" id="trCopy" title="Copy the visible lines, tab-separated (paste into Excel)">Copy</button>
          <a class="btn small" id="trCsv" href="/api/trail?project=${encodeURIComponent(d.project)}&kind=${state.kind}&csv=1" download="${esc(d.project)}-trail.csv">CSV</a>
        </div>
        <div class="table-scroll"><table class="grid tr-table" id="trTable"><thead></thead><tbody></tbody></table></div>
      </section>`;
    host.querySelectorAll(".tr-kind").forEach(x => x.onclick = () => { state.kind = x.dataset.kind; render(); });
    const s = $t("#trSearch"); s.oninput = () => { state.q = s.value.trim(); renderTable(); };
    $t("#trCopy").onclick = () => copyVisible();
    renderChart();
    renderTable();
  }

  const COLS = [
    ["date", "Date", "left"], ["party", "Vendor / customer", "left"], ["doc_number", "Bill / inv #", "left"],
    ["memo", "Memo", "left"], ["description", "Description", "left"], ["cost_code", "Code", "left"],
    ["amount", "Amount", "right"], ["running_total", "Running total", "right"], ["qb", "", "right"], ["scan", "", "right"],
  ];
  function renderTable() {
    const rows = visibleLines(), table = $t("#trTable"); if (!table) return;
    const thead = table.tHead, tbody = table.tBodies[0];
    thead.innerHTML = ""; const htr = document.createElement("tr");
    for (const [k, lab, al] of COLS) {
      const th = document.createElement("th"); th.className = al; th.textContent = lab;
      if (lab) { th.classList.add("tr-sortable"); if (state.sort.k === k) th.textContent += state.sort.dir > 0 ? " ▲" : " ▼";
        th.onclick = () => { state.sort = { k, dir: state.sort.k === k ? -state.sort.dir : 1 }; renderTable(); }; }
      htr.appendChild(th);
    }
    thead.appendChild(htr);
    tbody.innerHTML = "";
    if (!rows.length) { tbody.innerHTML = `<tr><td colspan="${COLS.length}" class="left tr-note">No lines match.</td></tr>`; return; }
    let month = null, mSum = 0, mCount = 0, total = 0;
    const flushMonth = () => { if (month == null) return;
      const tr = document.createElement("tr"); tr.className = "tr-msum";
      tr.innerHTML = `<td class="left" colspan="6">${monthLabel(month)} · ${mCount} line${mCount === 1 ? "" : "s"}</td><td>${tmc(mSum)}</td><td></td><td></td><td></td>`;
      tbody.appendChild(tr); };
    const byDate = state.sort.k === "date";
    for (const l of rows) {
      const mk = monthKey(l.date);
      if (byDate && mk !== month) { flushMonth(); month = mk; mSum = 0; mCount = 0; }
      mSum += l.amount || 0; mCount++; total += l.amount || 0;
      const tr = document.createElement("tr"); tr.className = "tr-line " + (l.kind === "billed" ? "tr-billed" : "tr-cost");
      const tag = l.kind === "billed" ? `<span class="tr-tag inv">invoice</span>` : l.is_sub ? `<span class="tr-tag sub">sub</span>` : "";
      const part = l.bill_total != null && Math.abs(l.bill_total - l.amount) > 0.5 ? `<span class="tr-part" title="This line is part of a ${esc(tm(l.bill_total))} bill">part of ${esc(tm(l.bill_total))}</span>` : "";
      tr.innerHTML =
        `<td class="left">${esc(d8(l.date))}</td>` +
        `<td class="left">${esc(l.party || "–")} ${tag}</td>` +
        `<td class="left tr-doc" title="Click to copy">${esc(l.doc_number || "–")}</td>` +
        `<td class="left tr-txt" title="${esc(l.memo)}">${esc(l.memo || "")}</td>` +
        `<td class="left tr-txt" title="${esc(l.description)}">${esc(l.description || "")} ${part}</td>` +
        `<td class="left">${l.cost_code ? `<span class="codechip">${esc(l.cost_code)}</span>` : esc(l.account || "")}</td>` +
        `<td>${tmc(l.amount)}</td><td class="tr-run">${tmc(l.running_total)}</td>` +
        `<td>${l.qbo_url ? `<a class="qbo-ico" href="${esc(l.qbo_url)}" target="_blank" rel="noopener" title="Open in QuickBooks">qb</a>` : ""}</td>` +
        `<td>${l.kind === "cost" && l.has_attachment !== false && l.qbo_url ? `<button class="btn tiny tr-scan" title="Open the scan">📎</button>` : ""}</td>`;
      const doc = tr.querySelector(".tr-doc"); if (doc && l.doc_number && typeof copy === "function") doc.onclick = () => copy(l.doc_number);
      const sc = tr.querySelector(".tr-scan"); if (sc && typeof openBillScan === "function") sc.onclick = (e) => { e.stopPropagation(); openBillScan({ url: l.qbo_url }, sc); };
      tbody.appendChild(tr);
    }
    if (byDate) flushMonth();
    const tot = document.createElement("tr"); tot.className = "tr-total";
    tot.innerHTML = `<td class="left" colspan="6">Total of the ${rows.length} line${rows.length === 1 ? "" : "s"} shown</td><td>${tmc(total)}</td><td></td><td></td><td></td>`;
    tbody.appendChild(tot);
  }

  function copyVisible() {
    const rows = visibleLines();
    const head = ["Date", "Vendor/customer", "Bill/inv #", "Memo", "Description", "Code", "Amount", "Running total", "QBO link"].join("\t");
    const body = rows.map(l => [d8(l.date), l.party || "", l.doc_number || "", l.memo || "", l.description || "", l.cost_code || l.account || "",
      Math.round((l.amount || 0) * 100) / 100, Math.round((l.running_total || 0) * 100) / 100, l.qbo_url || ""].join("\t"));
    const txt = [head, ...body].join("\n");
    if (typeof copy === "function") copy(txt); else navigator.clipboard && navigator.clipboard.writeText(txt);
  }

  // ── the red line: cumulative cost vs ETC (and billed vs contract), inline SVG, no libraries ──
  function renderChart() {
    const box = $t("#trChart"), d = state.data; if (!box || !d) return;
    const series = [];
    const cum = kind => d.lines.filter(l => l.kind === kind).map(l => ({ x: l.date, y: l.running_total }));
    if (state.kind !== "billed") series.push({ name: "Costs to date", pts: cum("cost"), ref: d.budget.etc, refName: "ETC (budget)", cls: "cost" });
    if (state.kind !== "costs") series.push({ name: "Billed to date", pts: cum("billed"), ref: d.budget.contract, refName: "Contract", cls: "billed" });
    const all = series.flatMap(s => s.pts);
    if (!all.length) { box.innerHTML = `<div class="tr-note">No lines to draw.</div>`; return; }
    const xs = all.map(p => new Date(p.x).getTime()).filter(v => !isNaN(v));
    const x0 = Math.min(...xs), x1 = Math.max(...xs, x0 + 86400000);
    const ymax = Math.max(...all.map(p => p.y), ...series.map(s => s.ref || 0)) * 1.06 || 1;
    const W = 1000, H = 190, L = 70, R = 16, T = 14, B = 28;
    const X = t => L + (W - L - R) * ((t - x0) / (x1 - x0));
    const Y = v => T + (H - T - B) * (1 - v / ymax);
    const fmtK = v => v >= 1e6 ? `$${(v / 1e6).toFixed(1)}M` : v >= 1e3 ? `$${Math.round(v / 1e3)}k` : `$${Math.round(v)}`;
    let svg = `<svg viewBox="0 0 ${W} ${H}" class="tr-svg" role="img" aria-label="running totals against budget and contract">`;
    for (let i = 0; i <= 4; i++) { const v = ymax * i / 4; svg += `<line x1="${L}" x2="${W - R}" y1="${Y(v)}" y2="${Y(v)}" class="tr-grid"/><text x="${L - 8}" y="${Y(v) + 4}" class="tr-ax" text-anchor="end">${fmtK(v)}</text>`; }
    svg += `<text x="${L}" y="${H - 8}" class="tr-ax">${d8(new Date(x0).toISOString().slice(0, 10))}</text><text x="${W - R}" y="${H - 8}" class="tr-ax" text-anchor="end">${d8(new Date(x1).toISOString().slice(0, 10))}</text>`;
    for (const s of series) {
      if (s.ref) svg += `<line x1="${L}" x2="${W - R}" y1="${Y(s.ref)}" y2="${Y(s.ref)}" class="tr-ref ${s.cls}"/><text x="${W - R}" y="${Y(s.ref) - 5}" class="tr-reflab ${s.cls}" text-anchor="end">${esc(s.refName)} ${fmtK(s.ref)}</text>`;
      let path = "", prev = null, cross = null;
      for (const p of s.pts) { const t = new Date(p.x).getTime(); if (isNaN(t)) continue;
        if (prev) path += ` L${X(t)},${Y(prev.y)}`;            // step: the total holds until the next line lands
        path += `${path ? " L" : "M"}${X(t)},${Y(p.y)}`;
        if (s.ref && !cross && p.y > s.ref) cross = { x: X(t), y: Y(p.y), date: p.x, y0: p.y };
        prev = { t, y: p.y }; }
      svg += `<path d="${path}" class="tr-line ${s.cls}"/>`;
      if (cross) svg += `<circle cx="${cross.x}" cy="${cross.y}" r="5" class="tr-cross"/><text x="${cross.x + 8}" y="${cross.y - 8}" class="tr-crosslab">${esc(s.name)} passed ${esc(s.refName)} on ${d8(cross.date)}</text>`;
    }
    svg += `</svg>`;
    const legend = series.map(s => `<span class="tr-lg ${s.cls}">${esc(s.name)}</span>`).join("") + `<span class="tr-lg ref">red rule = budget / contract</span>`;
    box.innerHTML = svg + `<div class="tr-legend">${legend}</div>`;
  }
})();
