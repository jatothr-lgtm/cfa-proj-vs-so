/* Projection vs SO — loading, views and export. The rules live in calc.js. */

import {
  TABS, TAB_LABEL, DEFAULT_WAREHOUSES, CAT_A,
  classifySheet, parseNormalisation, extractTab, prepare, leaves, tree, classify, ymLabel,
} from "./calc.js";

const $ = (s) => document.querySelector(s);
const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

const state = {
  tabs: {},                 // tab → { records, missing, periods, file, sheet }
  warehouses: DEFAULT_WAREHOUSES,
  whSource: "built-in default",
  notes: [],                // load-time messages
  prepared: null,
  ui: { view: "data", mode: "all", wh: "", sort: "gap", q: "", matchMonth: true, open: { customer: null, item: null } },
};

/* ── formatting ──────────────────────────────────────────────────── */

const nf0 = new Intl.NumberFormat("en-IN", { maximumFractionDigits: 0 });
const nf1 = new Intl.NumberFormat("en-IN", { minimumFractionDigits: 1, maximumFractionDigits: 1 });
const kg = (v) => (Math.abs(v) < 10 && v !== 0 ? nf1.format(v) : nf0.format(v));
const signed = (v) => (v > 0.0005 ? "+" : v < -0.0005 ? "−" : "") + kg(Math.abs(v));
const ratio = (p, s) => (s > 0 ? nf0.format((p / s) * 100) + "%" : "—");
const BAND = { over: "▲ Over", under: "▼ Under", match: "● Match" };

function toast(msg, ms = 2600) {
  const t = $("#toast");
  t.textContent = msg;
  t.classList.add("show");
  clearTimeout(toast.h);
  toast.h = setTimeout(() => t.classList.remove("show"), ms);
}

/* ── loading ─────────────────────────────────────────────────────── */

async function loadFiles(files) {
  if (!window.XLSX) { toast("The Excel reader did not load — check the connection and reload."); return; }
  for (const f of files) {
    toast(`Reading ${f.name}…`, 60000);
    await new Promise((r) => setTimeout(r, 30));
    try {
      const wb = XLSX.read(await f.arrayBuffer(), { type: "array", cellDates: true, dense: true });
      let found = 0;
      const seen = new Set();
      for (const name of wb.SheetNames) {
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null });
        const c = classifySheet(name, rows);
        if (!c) continue;
        if (c.tab === "normalisation") {
          const wh = parseNormalisation(rows);
          if (wh.length) {
            state.warehouses = wh;
            state.whSource = `Normalisation tab of ${f.name}`;
            found++;
          } else {
            state.notes.push({ kind: "warn", msg: `${f.name} › ${name}: no "CFA (…)" rows found — using the built-in warehouse mapping.` });
          }
          continue;
        }
        if (seen.has(c.tab)) {
          state.notes.push({ kind: "warn", msg: `${f.name} has more than one ${TAB_LABEL[c.tab]} sheet — used the first, ignored "${name}".` });
          continue;
        }
        seen.add(c.tab);
        state.tabs[c.tab] = { ...extractTab(c.tab, rows, c.headerRow), file: f.name, sheet: name, headerRow: c.headerRow };
        found++;
      }
      if (!found) state.notes.push({ kind: "err", msg: `${f.name}: no projection, so or dispatches sheet recognised.` });
    } catch (e) {
      console.error(e);
      state.notes.push({ kind: "err", msg: `${f.name} could not be read: ${e.message}` });
    }
  }
  toast("Loaded", 1200);
  recompute();
  if (TABS.every((t) => state.tabs[t] && !state.tabs[t].missing.length)) setView("customer");
  else render();
}

function recompute() {
  const ok = TABS.filter((t) => state.tabs[t] && !state.tabs[t].missing.length);
  const usable = Object.fromEntries(ok.map((t) => [t, state.tabs[t]]));
  state.prepared = ok.length ? prepare(usable, state.warehouses, { matchMonth: state.ui.matchMonth }) : null;
  state.ui.open = { customer: null, item: null };
}

/* ── views ───────────────────────────────────────────────────────── */

function setView(v) {
  state.ui.view = v;
  document.querySelectorAll("#tabs button").forEach((b) => b.classList.toggle("is-active", b.dataset.view === v));
  $("#view-data").classList.toggle("is-active", v === "data");
  $("#view-compare").classList.toggle("is-active", v !== "data");
  render();
}

function render() {
  renderSlots();
  renderData();
  if (state.ui.view !== "data") renderCompare();
}

function renderSlots() {
  const hint = {
    projection: "Wareouse · Customer Group · Customer · Item Group · Item Parent · Item Name · Total KGs",
    so: "Origin · Customer Group · Customer · NEW MIS ITEM GROUP · Parent Item · Item Name · Stock Qty in KGS",
    dispatches: "From Warehouse · Sales Channel · Customer · New Mis Item Group · Item Parent · Item Name · Stock Qty In Kg · Invoice Status",
  };
  $("#slots").innerHTML = TABS.map((t) => {
    const s = state.tabs[t];
    let body;
    if (!s) body = `<div class="state muted">Not loaded</div>`;
    else {
      const months = [...s.periods.entries()].sort().map(([p, n]) => `${ymLabel(p)} <span class="muted">(${nf0.format(n)})</span>`).join(", ") || "<span class='muted'>no dates</span>";
      body = `<div class="state"><b>${esc(s.file)}</b> › ${esc(s.sheet)}<br>
        ${nf0.format(s.records.length)} rows${s.headerRow ? ` · header on row ${s.headerRow + 1}` : ""}<br>
        ${t === "projection" ? "Month" : t === "so" ? "Sales Order Date" : "Invoice Date"}: ${months}</div>
        ${s.missing.length ? `<div class="banner err">Missing column${s.missing.length > 1 ? "s" : ""}: <b>${s.missing.map(esc).join(", ")}</b> — this tab is not used.</div>` : ""}`;
    }
    return `<div class="slot ${s ? "has" : ""}"><div class="hd"><strong>${TAB_LABEL[t]}</strong>
      <span class="src ${s ? "fresh" : "empty"}">${s ? (s.missing.length ? "unusable" : "loaded") : "empty"}</span></div>
      <div class="hint">${hint[t]}</div>${body}</div>`;
  }).join("");
}

function renderData() {
  const out = [];
  if (state.notes.length) {
    out.push(state.notes.map((n) => `<div class="banner ${n.kind}">${esc(n.msg)}</div>`).join(""));
  }
  out.push(`<div class="grid g-2">`);

  out.push(`<div class="card"><h3>How SO is built</h3>
    <p class="note">Projection and SO are compared in kg for the CFAs below only.</p>
    <dl class="kv">
      <dt>Month proj</dt><dd><b>projection</b> › Total KGs</dd>
      <dt>MTD proj</dt><dd>Month proj ÷ days in the month × SO days — the per-day projection over the days SO covers</dd>
      <dt>SO days</dt><dd>distinct Sales Order Dates in the <b>so</b> tab, one count for every CFA and customer, ${CAT_A} included</dd>
      <dt>SO</dt><dd><b>so</b> › Stock Qty in KGS where Customer Group ≠ ${CAT_A}</dd>
      <dt>+</dt><dd><b>dispatches</b> › Stock Qty In Kg where Sales Channel = ${CAT_A}, Invoice Status ≠ Return</dd>
      <dt>${CAT_A}</dt><dd>compared as one customer: every ${CAT_A} customer on either side is added into one row</dd>
      <dt>Item Group</dt><dd>projection › Item Group · so › NEW MIS ITEM GROUP · dispatches › New Mis Item Group</dd>
      <dt>Items</dt><dd>matched on Item Code; labels taken from the projection first, then so, then dispatches</dd>
      <dt>Over / Under</dt><dd>decided per customer or per item on MTD proj: MTD proj &gt; SO is over, MTD proj &lt; SO is under</dd>
    </dl></div>`);

  out.push(`<div class="card"><h3>CFAs compared</h3><p class="note">From the ${esc(state.whSource)}. Any other warehouse is left out.</p>
    <div class="tblwrap"><table class="plain"><thead><tr><th>CFA</th><th class="l">projection › Wareouse</th><th class="l">so › Origin</th><th class="l">dispatches › From Warehouse</th></tr></thead>
    <tbody>${state.warehouses.map((w) => `<tr><td>${esc(w.code)}</td><td class="l">${esc(w.projection)}</td><td class="l">${esc(w.so)}</td><td class="l">${esc(w.dispatches)}</td></tr>`).join("")}</tbody></table></div></div>`);
  out.push(`</div>`);

  const p = state.prepared;
  if (p) {
    out.push(`<div class="card" style="margin-top:14px"><h3>Rows left out</h3>
      <p class="note">Every row that does not reach the comparison, and why.${p.filterMonth ? ` Month filter: ${p.projPeriods.map(ymLabel).join(", ")}.` : " Month filter is off."}</p>
      <div class="tblwrap"><table><thead><tr><th>Tab</th><th class="l">Reason</th><th>Rows</th><th>Kg</th></tr></thead><tbody>
      ${p.exclusions.length ? p.exclusions.map((e) => `<tr><td>${TAB_LABEL[e.tab]}</td><td class="l">${esc(e.reason)}</td><td class="num">${nf0.format(e.rows)}</td><td class="num">${kg(e.kg)}</td></tr>`).join("") : `<tr><td colspan="4" class="l muted">Nothing left out.</td></tr>`}
      </tbody></table></div>
      ${[...p.outside.entries()].map(([tab, m]) => `<details style="margin-top:10px"><summary class="small">${TAB_LABEL[tab]}: ${m.size} warehouse${m.size > 1 ? "s" : ""} outside the compared CFAs</summary>
        <div class="tblwrap" style="margin-top:6px"><table><tbody>${[...m.entries()].sort((a, b) => b[1] - a[1]).map(([l, v]) => `<tr><td class="l">${esc(l)}</td><td class="num">${kg(v)}</td></tr>`).join("")}</tbody></table></div></details>`).join("")}
      </div>`);
  }
  $("#dataBody").innerHTML = out.join("");
}

/* ── the comparison views ────────────────────────────────────────── */

function banners() {
  const out = [];
  const missing = TABS.filter((t) => !state.tabs[t] || state.tabs[t].missing.length);
  if (missing.length) {
    out.push(`<div class="banner warn"><b>${missing.map((t) => TAB_LABEL[t]).join(", ")} not loaded.</b>
      Its kg count as zero below, so ${missing.includes("projection") ? "everything reads under-projected" : missing.includes("dispatches") ? `${CAT_A} reads fully over-projected` : "every customer outside " + CAT_A + " reads over-projected"} — load it on the Data tab.</div>`);
  }
  const p = state.prepared;
  if (p) {
    const pr = p.prorate;
    out.push(pr.on
      ? `<div class="banner info"><b>MTD projection = ${pr.soDays} of ${pr.monthDays} days</b> (${nf0.format(pr.factor * 100)}% of the month):
          Month projection ÷ ${pr.monthDays} days in ${p.projPeriods.map(ymLabel).join(", ")} × ${pr.soDays} unique Sales Order Dates in the so tab. Gap and Over / Under compare MTD projection with SO.</div>`
      : `<div class="banner warn">${esc(pr.why)}</div>`);
  }
  if (p && p.projPeriods.length) {
    for (const t of ["so", "dispatches"]) {
      const s = state.tabs[t];
      if (!s || s.missing.length) continue;
      const have = [...s.periods.keys()];
      const inMonth = have.filter((m) => p.projPeriods.includes(m));
      const off = have.filter((m) => !p.projPeriods.includes(m));
      if (p.filterMonth && !inMonth.length && have.length) {
        out.push(`<div class="banner err"><b>${TAB_LABEL[t]} has no rows in ${p.projPeriods.map(ymLabel).join(", ")}</b> (it covers ${have.map(ymLabel).join(", ")}).
          Its contribution is zero${t === "dispatches" ? `, so ${CAT_A} reads fully over-projected` : ""}. Load the same month, or untick “Only the projection month”.</div>`);
      } else if (p.filterMonth && off.length) {
        out.push(`<div class="banner info">${TAB_LABEL[t]} also covers ${off.map(ymLabel).join(", ")}; those rows are left out to match the projection month.</div>`);
      } else if (!p.filterMonth && off.length) {
        out.push(`<div class="banner warn">Month filter is off: ${TAB_LABEL[t]} (${have.map(ymLabel).join(", ")}) is compared against a ${p.projPeriods.map(ymLabel).join(", ")} projection.</div>`);
      }
    }
  }
  return out.join("");
}

function currentLeaves(view) {
  if (!state.prepared) return [];
  let ls = leaves(state.prepared, view, state.ui.wh);
  const q = state.ui.q.trim().toLowerCase();
  if (q) ls = ls.filter((l) => l.path.some((x) => x.toLowerCase().includes(q)) || (l.code || "").toLowerCase().includes(q));
  return ls;
}

const nodeId = (path) => path.join("\u0001");

function branchIds(t) {
  const ids = new Set();
  const walk = (n, path) => n.kids.forEach((k) => { if (!k.leaf) { const p = [...path, k.label]; ids.add(nodeId(p)); walk(k, p); } });
  walk(t, []);
  return ids;
}

function renderCompare() {
  const { view, mode } = state.ui;
  $("#cmpTitle").textContent = view === "customer" ? "By customer" : "By item";
  $("#cmpSub").innerHTML = view === "customer"
    ? `Customer Group › Customer. MTD projection (per-day projection × SO days) against SO (so tab, plus ${CAT_A} from dispatches).`
    : `Item Group › Item Parent › Item Name, across all customers. Same SO as the customer view.`;
  $("#cmpBanners").innerHTML = banners();

  const all = currentLeaves(view);
  const by = { over: [], under: [], match: [] };
  for (const l of all) by[classify(l.gap)].push(l);
  document.querySelectorAll("#mode button").forEach((b) => {
    const m = b.dataset.mode;
    b.classList.toggle("on", m === mode);
    const n = m === "all" ? all.length : by[m].length;
    b.innerHTML = `${{ all: "All", over: "Over-projected", under: "Under-projected" }[m]}<span class="ct">${n}</span>`;
  });

  // Tiles always describe the whole view, whatever the mode.
  const sum = (a, f) => a.reduce((s, l) => s + l[f], 0);
  const P = sum(all, "proj"), PF = sum(all, "projFull"), S = sum(all, "so"), ST = sum(all, "soTab"), SD = sum(all, "soDis");
  const OV = sum(by.over, "gap"), UN = sum(by.under, "gap");
  const unit = view === "customer" ? "customer" : "item";
  $("#cmpTiles").innerHTML = state.prepared ? `
    <div class="tile"><div class="k">Month projection</div><div class="v">${kg(PF)}</div><div class="m">kg · Total KGs</div></div>
    <div class="tile"><div class="k">MTD projection</div><div class="v">${kg(P)}</div><div class="m">${state.prepared?.prorate.on ? `${state.prepared.prorate.soDays} of ${state.prepared.prorate.monthDays} days` : "= month (no SO days)"}</div></div>
    <div class="tile"><div class="k">SO</div><div class="v">${kg(S)}</div><div class="m">so tab ${kg(ST)} · ${CAT_A} dispatches ${kg(SD)}</div></div>
    <div class="tile"><div class="k">Net gap</div><div class="v">${signed(P - S)}</div><div class="m">MTD proj − SO · MTD ÷ SO ${ratio(P, S)}</div></div>
    <div class="tile"><div class="k">Over-projected</div><div class="v" style="color:var(--amber)">${signed(OV)}</div><div class="m">${by.over.length} ${unit}${by.over.length === 1 ? "" : "s"} with MTD proj &gt; SO</div></div>
    <div class="tile"><div class="k">Under-projected</div><div class="v" style="color:var(--red)">${signed(UN)}</div><div class="m">${by.under.length} ${unit}${by.under.length === 1 ? "" : "s"} with MTD proj &lt; SO</div></div>` : "";

  if (!state.prepared) {
    $("#cmpBody").innerHTML = `<div class="gate"><b>Nothing to compare yet</b><div class="why">Load the projection, so and dispatches tabs on the Data tab.</div></div>`;
    return;
  }
  const shown = mode === "all" ? all : by[mode];
  if (!shown.length) {
    $("#cmpBody").innerHTML = `<div class="gate"><b>No ${mode === "all" ? "rows" : mode + "-projected " + unit + "s"}</b><div class="why">${state.ui.q ? "Nothing matches the search." : "Every " + unit + " lands elsewhere."}</div></div>`;
    return;
  }

  const t = tree(shown, state.ui.sort, mode);
  if (!state.ui.open[view]) {
    // Start with the top level open, so groups show their customers / item
    // parents; a search opens everything so the matches are on screen.
    state.ui.open[view] = state.ui.q.trim() ? branchIds(t) : new Set(t.kids.map((k) => nodeId([k.label])));
  }
  const open = state.ui.open[view];
  const rows = [];
  const walk = (node, path) => {
    for (const k of node.kids) {
      const pth = [...path, k.label];
      const id = nodeId(pth);
      const isLeaf = !!k.leaf;
      const cls = classify(k.gap);
      const isOpen = open.has(id);
      let title = "";
      if (isLeaf && view === "customer" && k.label.startsWith(CAT_A)) {
        title = `Projection: ${[...k.leaf.custs.proj].join(", ") || "none"} · SO: ${k.leaf.custs.so.size} dispatch customer(s)`;
      }
      rows.push(`<tr class="lvl${k.depth}${isLeaf ? " leaf" : ""}">
        <td class="l"><span class="tw" style="padding-left:${k.depth * 20}px"${title ? ` title="${esc(title)}"` : ""}>
          ${isLeaf ? `<span class="caret-sp"></span>` : `<button class="caret${isOpen ? " open" : ""}" data-id="${esc(id)}" aria-label="${isOpen ? "Collapse" : "Expand"}">▸</button>`}
          <span class="t">${esc(k.label)}${isLeaf && k.leaf.code ? `<span class="code">${esc(k.leaf.code)}</span>` : ""}</span>
          ${!isLeaf ? `<span class="muted small">${k.n}</span>` : ""}</span></td>
        <td class="num muted">${kg(k.projFull)}</td>
        <td class="num">${kg(k.proj)}</td>
        <td class="num">${kg(k.so)}</td>
        ${view === "item" ? `<td class="num muted">${kg(k.soDis)}</td>` : ""}
        <td class="num gap-${cls}">${signed(k.gap)}</td>
        <td class="num">${ratio(k.proj, k.so)}</td>
        <td>${isLeaf ? `<span class="band band-${cls}">${BAND[cls]}</span>` : ""}</td>
      </tr>`);
      if (!isLeaf && isOpen) walk(k, pth);
    }
  };
  walk(t, []);
  const head = view === "customer"
    ? `<th class="l">Customer Group › Customer</th>`
    : `<th class="l">Item Group › Item Parent › Item Name</th>`;
  $("#cmpBody").innerHTML = `<div class="tblwrap"><table>
    <thead><tr>${head}<th title="Total KGs for the whole month">Month proj kg</th><th title="Month proj ÷ days in month × unique Sales Order Dates">MTD proj kg</th><th>SO kg</th>${view === "item" ? `<th title="Part of SO that came from ${CAT_A} dispatches">of which ${CAT_A}</th>` : ""}<th title="MTD proj − SO">Gap kg</th><th>MTD ÷ SO</th><th>Status</th></tr></thead>
    <tbody>${rows.join("")}</tbody>
    <tfoot><tr><td class="l">Total · ${shown.length} ${unit}${shown.length === 1 ? "" : "s"}</td><td class="num muted">${kg(t.projFull)}</td><td class="num">${kg(t.proj)}</td><td class="num">${kg(t.so)}</td>${view === "item" ? `<td class="num muted">${kg(t.soDis)}</td>` : ""}<td class="num">${signed(t.gap)}</td><td class="num">${ratio(t.proj, t.so)}</td><td></td></tr></tfoot>
  </table></div>`;
}

/* ── export ──────────────────────────────────────────────────────── */

function exportXlsx() {
  if (!state.prepared) { toast("Nothing loaded yet"); return; }
  const wb = XLSX.utils.book_new();
  const order = { over: (a, b) => b.gap - a.gap, under: (a, b) => a.gap - b.gap, all: (a, b) => Math.abs(b.gap) - Math.abs(a.gap) };
  for (const view of ["customer", "item"]) {
    const ls = leaves(state.prepared, view, state.ui.wh);
    const head = view === "customer"
      ? ["Customer Group", "Customer"]
      : ["Item Group", "Item Parent", "Item Name", "Item Code"];
    for (const mode of ["all", "over", "under"]) {
      const pick = mode === "all" ? ls : ls.filter((l) => classify(l.gap) === mode);
      const aoa = [[...head, "Month projection kg", "MTD projection kg", "SO kg", "SO from so tab kg", `SO from ${CAT_A} dispatches kg`, "Gap kg (MTD proj − SO)", "MTD proj ÷ SO", "Status"]];
      for (const l of [...pick].sort(order[mode])) {
        aoa.push([
          ...(view === "customer" ? l.path : [...l.path, l.code]),
          round(l.projFull), round(l.proj), round(l.so), round(l.soTab), round(l.soDis), round(l.gap),
          l.so > 0 ? round(l.proj / l.so, 4) : null,
          { over: "Over-projected", under: "Under-projected", match: "Match" }[classify(l.gap)],
        ]);
      }
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      ws["!cols"] = aoa[0].map((h, i) => ({ wch: i < head.length ? (i === head.length - 1 && view === "item" ? 18 : 34) : 16 }));
      const name = `${view === "customer" ? "Customer" : "Item"} - ${{ all: "All", over: "Over", under: "Under" }[mode]}`;
      XLSX.utils.book_append_sheet(wb, ws, name);
    }
  }
  const p = state.prepared;
  const info = [
    ["Projection vs SO"],
    ["CFA", state.ui.wh || "All CFAs"],
    ["Month filter", p.filterMonth ? p.projPeriods.map(ymLabel).join(", ") : "off"],
    ["MTD projection", p.prorate.on ? `Month projection ÷ ${p.prorate.monthDays} days × ${p.prorate.soDays} unique Sales Order Dates (${round(p.prorate.factor, 4)})` : p.prorate.why],
    [],
    ["Tab", "Left out because", "Rows", "Kg"],
    ...p.exclusions.map((e) => [TAB_LABEL[e.tab], e.reason, e.rows, round(e.kg)]),
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(info), "Notes");
  const tag = [p.projPeriods[0] || "", state.ui.wh || "all"].filter(Boolean).join("-");
  XLSX.writeFile(wb, `proj-vs-so-${tag}.xlsx`);
}

const round = (v, d = 3) => Math.round(v * 10 ** d) / 10 ** d;

/* ── wiring ──────────────────────────────────────────────────────── */

function fillWarehouses() {
  const sel = $("#wh");
  const cur = sel.value;
  sel.innerHTML = `<option value="">All CFAs</option>` + state.warehouses.map((w) => `<option value="${esc(w.code)}">CFA (${esc(w.code)})</option>`).join("");
  sel.value = state.warehouses.some((w) => w.code === cur) ? cur : "";
  state.ui.wh = sel.value;
}

document.querySelectorAll("#tabs button").forEach((b) => b.addEventListener("click", () => setView(b.dataset.view)));
document.querySelectorAll("#mode button").forEach((b) => b.addEventListener("click", () => { state.ui.mode = b.dataset.mode; renderCompare(); }));
$("#wh").addEventListener("change", (e) => { state.ui.wh = e.target.value; state.ui.open = { customer: null, item: null }; renderCompare(); });
$("#sort").addEventListener("change", (e) => { state.ui.sort = e.target.value; renderCompare(); });
$("#q").addEventListener("input", (e) => { state.ui.q = e.target.value; state.ui.open = { customer: null, item: null }; renderCompare(); });
$("#matchMonth").addEventListener("change", (e) => { state.ui.matchMonth = e.target.checked; recompute(); render(); });
$("#export").addEventListener("click", exportXlsx);
$("#expand").addEventListener("click", () => {
  const v = state.ui.view;
  state.ui.open[v] = branchIds(tree(state.ui.mode === "all" ? currentLeaves(v) : currentLeaves(v).filter((l) => classify(l.gap) === state.ui.mode)));
  renderCompare();
});
$("#collapse").addEventListener("click", () => { state.ui.open[state.ui.view] = new Set(); renderCompare(); });
$("#cmpBody").addEventListener("click", (e) => {
  const b = e.target.closest("button.caret");
  if (!b) return;
  const open = state.ui.open[state.ui.view];
  open.has(b.dataset.id) ? open.delete(b.dataset.id) : open.add(b.dataset.id);
  renderCompare();
});

const input = $("#file"), drop = $("#drop");
input.addEventListener("change", async () => {
  const files = [...input.files];
  input.value = "";
  if (files.length) { await loadFiles(files); fillWarehouses(); render(); }
});
["dragenter", "dragover"].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add("drag"); }));
["dragleave", "drop"].forEach((ev) => drop.addEventListener(ev, () => drop.classList.remove("drag")));
drop.addEventListener("drop", async (e) => {
  e.preventDefault();
  const files = [...(e.dataTransfer?.files || [])];
  if (files.length) { await loadFiles(files); fillWarehouses(); render(); }
});

fillWarehouses();
render();
