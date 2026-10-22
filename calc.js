/* Projection vs SO — parsing and the comparison itself. No DOM in here.
 *
 *   Month proj = projection tab, Total KGs
 *   MTD proj   = Month proj ÷ days in the month × SO days (distinct Sales Order Dates)
 *   SO         = so tab, Stock Qty in KGS, Customer Group ≠ Category A
 *              + dispatches tab, Stock Qty In Kg, Sales Channel = Category A,
 *                Invoice Status ≠ Return
 *
 * Only the CFAs in the Normalisation tab are compared; each tab spells the
 * warehouse its own way and the Normalisation tab says which spellings are one
 * CFA. Item Group is the MIS item group in all three tabs (the projection's
 * own "Item Group" column holds it; so / dispatches carry it as NEW MIS ITEM
 * GROUP). Items are matched on Item Code, so one SKU never splits into two
 * rows because the tabs label its parent differently.
 */

export const TABS = ["projection", "so", "dispatches"];

export const TAB_LABEL = { projection: "Projection", so: "SO", dispatches: "Dispatches" };

/* Used when the workbook carries no Normalisation tab. */
export const DEFAULT_WAREHOUSES = [
  { code: "GGN", projection: "Storage Beyond Sqfeet Gurgaon Finished Goods - CBSPL",   so: "CFA (GGN)", dispatches: "Storage Beyond Sqfeet Gurgaon Finished Goods - CBSPL" },
  { code: "BLR", projection: "Storage Beyond Sqfeet Bangalore Finished Goods - CBSPL", so: "CFA (BLR)", dispatches: "Storage Beyond Sqfeet Bangalore Finished Goods - CBSPL" },
  { code: "MUM", projection: "Storage Beyond Sqfeet Mumbai Finished Goods - CBSPL",    so: "CFA (MUM)", dispatches: "Storage Beyond Sqfeet Mumbai Finished Goods - CBSPL" },
];

export const CAT_A = "Category A";
const CAT_A_KEY = "category a";
const CAT_A_CUSTOMER = "Category A — all customers";

/* Column aliases per tab. The first alias is the name in the current extract. */
const COLS = {
  projection: {
    wh:     ["wareouse", "warehouse"],
    cgroup: ["customer group"],
    cust:   ["customer"],
    code:   ["item code"],
    group:  ["item group", "new mis item group"],
    parent: ["item parent", "parent item"],
    name:   ["item name"],
    qty:    ["total kgs", "total kg"],
    month:  ["month"],
    year:   ["year"],
  },
  so: {
    wh:     ["origin"],
    cgroup: ["customer group"],
    cust:   ["customer"],
    code:   ["item code"],
    group:  ["new mis item group"],
    parent: ["parent item", "item parent"],
    name:   ["item name"],
    qty:    ["stock qty in kgs", "stock qty in kg"],
    date:   ["sales order date"],
  },
  dispatches: {
    wh:     ["from warehouse"],
    cgroup: ["sales channel"],
    cust:   ["customer"],
    code:   ["item code"],
    group:  ["new mis item group"],
    parent: ["item parent", "parent item"],
    name:   ["item name"],
    qty:    ["stock qty in kg", "stock qty in kgs"],
    date:   ["invoice date"],
    status: ["invoice status"],
  },
};

const REQUIRED = {
  projection: ["wh", "cgroup", "cust", "group", "parent", "name", "qty"],
  so:         ["wh", "cgroup", "cust", "group", "parent", "name", "qty"],
  dispatches: ["wh", "cgroup", "cust", "group", "parent", "name", "qty", "status"],
};

/* What tells the tabs apart, tried in this order. */
const SIGNATURE = {
  dispatches: ["from warehouse", "sales channel"],
  so:         ["origin", "customer group", "stock qty in kgs"],
  projection: ["total kgs"],
};

const MONTHS = ["january","february","march","april","may","june","july","august","september","october","november","december"];

/* ── text helpers ─────────────────────────────────────────────────── */

const hkey = (s) => String(s ?? "").toLowerCase().replace(/[^a-z0-9]/g, "");
export const clean = (s) => s == null ? "" : String(s).replace(/[\s ]+/g, " ").trim();
const key = (s) => clean(s).toLowerCase();

const num = (v) => {
  if (typeof v === "number") return Number.isFinite(v) ? v : 0;
  const n = parseFloat(String(v ?? "").replace(/,/g, ""));
  return Number.isFinite(n) ? n : 0;
};

const pad = (n) => String(n).padStart(2, "0");

/* "YYYY-MM-DD" from a Date, an ISO / D-M-Y string or an Excel serial. */
export function ymd(v) {
  if (v == null || v === "") return null;
  if (v instanceof Date && !isNaN(v)) return `${v.getFullYear()}-${pad(v.getMonth() + 1)}-${pad(v.getDate())}`;
  if (typeof v === "number" && v > 20000 && v < 80000) {
    const d = new Date(Math.round((v - 25569) * 86400000));
    return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`;
  }
  const s = String(v).trim();
  let m = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if (m) return `${m[1]}-${pad(m[2])}-${pad(m[3])}`;
  m = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})/);
  if (m) return `${m[3]}-${pad(m[2])}-${pad(m[1])}`;
  return null;
}

/* "YYYY-MM" — the month of ymd(). */
export const ym = (v) => ymd(v)?.slice(0, 7) ?? null;

/* Calendar days in a "YYYY-MM" month. */
export const daysInMonth = (p) => {
  const [y, m] = p.split("-").map(Number);
  return new Date(y, m, 0).getDate();
};

export function ymLabel(p) {
  if (!p) return "—";
  const [y, m] = p.split("-");
  const n = MONTHS[+m - 1];
  return n ? `${n[0].toUpperCase()}${n.slice(1)} ${y}` : p;
}

/* ── sheet recognition ───────────────────────────────────────────── */

function findHeader(rows, wanted) {
  const need = wanted.map(hkey);
  for (let i = 0; i < Math.min(rows.length, 15); i++) {
    const cells = new Set((rows[i] || []).map(hkey));
    if (need.every((k) => cells.has(k))) return i;
  }
  return -1;
}

function colIndex(header, aliases) {
  const keys = header.map(hkey);
  for (const a of aliases) {
    const i = keys.indexOf(hkey(a));
    if (i !== -1) return i;
  }
  return -1;
}

/* Which input a sheet is, or null. `rows` is an array of row arrays. */
export function classifySheet(name, rows) {
  if (/normali[sz]ation/i.test(name)) return { tab: "normalisation", headerRow: 0 };
  for (const tab of ["dispatches", "so", "projection"]) {
    const h = findHeader(rows, SIGNATURE[tab]);
    if (h !== -1) return { tab, headerRow: h };
  }
  return null;
}

/* Warehouse rows of the Normalisation tab: the row whose so cell reads
 * "CFA (XXX)" gives one CFA, with the projection and dispatches spellings
 * beside it. Column positions come from the tab-name row at the top. */
export function parseNormalisation(rows) {
  let pos = null;
  for (let i = 0; i < Math.min(rows.length, 5) && !pos; i++) {
    const r = (rows[i] || []).map(key);
    const p = r.findIndex((c) => c === "projection");
    const s = r.findIndex((c) => c === "so");
    const d = r.findIndex((c) => c.startsWith("dispatch"));
    if (p !== -1 && s !== -1 && d !== -1) pos = { p, s, d };
  }
  if (!pos) return [];
  const out = [];
  for (const r of rows) {
    const so = clean(r?.[pos.s]);
    const m = so.match(/^CFA\s*\(\s*([^)]+?)\s*\)$/i);
    if (!m) continue;
    out.push({
      code: m[1].toUpperCase(),
      projection: clean(r[pos.p]),
      so,
      dispatches: clean(r[pos.d]),
    });
  }
  return out;
}

/* Pull one tab into plain records. Reports the columns it could not find. */
export function extractTab(tab, rows, headerRow) {
  const header = rows[headerRow] || [];
  const idx = {};
  for (const [f, aliases] of Object.entries(COLS[tab])) idx[f] = colIndex(header, aliases);
  const missing = REQUIRED[tab].filter((f) => idx[f] === -1).map((f) => COLS[tab][f][0]);
  const get = (r, f) => (idx[f] === -1 ? null : r[idx[f]]);

  const records = [];
  const periods = new Map();
  for (let i = headerRow + 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || r.every((c) => c == null || c === "")) continue;
    const rec = {
      wh: clean(get(r, "wh")),
      cgroup: clean(get(r, "cgroup")),
      cust: clean(get(r, "cust")),
      code: clean(get(r, "code")),
      group: clean(get(r, "group")),
      parent: clean(get(r, "parent")),
      name: clean(get(r, "name")),
      qty: num(get(r, "qty")),
      status: clean(get(r, "status")),
      period: null,
      day: null,
      total: false,
    };
    if (tab === "projection") {
      const mo = key(get(r, "month"));
      const yr = Math.round(num(get(r, "year")));
      const mi = MONTHS.indexOf(mo);
      if (mi !== -1 && yr >= 2000 && yr <= 2100) rec.period = `${yr}-${pad(mi + 1)}`;
      // The extract ends on a grand-total row; summing it would double everything.
      rec.total = [get(r, "month"), rec.wh, rec.cust].some((v) => /^(grand )?total$/i.test(clean(v)));
    } else {
      rec.day = ymd(get(r, "date"));
      rec.period = rec.day?.slice(0, 7) ?? null;
    }
    if (!rec.total && rec.period) periods.set(rec.period, (periods.get(rec.period) || 0) + 1);
    records.push(rec);
  }
  return { records, missing, periods };
}

/* ── the comparison ──────────────────────────────────────────────── */

/* Apply the rules once; everything downstream reads `rows`.
 *   tabs:       { projection, so, dispatches } as returned by extractTab
 *   warehouses: [{ code, projection, so, dispatches }]
 *   matchMonth: keep so / dispatches rows only in the projection's month(s)
 *   prorate:    scale the projection to the days the SO dump covers
 *
 * Each projection row carries both figures: `fullQty` is the month's Total KGs,
 * `qty` the MTD projection. Gap and Over / Under are always MTD projection − SO.
 *
 * Pro-rating: Projection = Total KGs ÷ days in the projection month × SO days,
 * where SO days = distinct Sales Order Dates in the so tab — counted over the
 * whole tab, before the Category A and warehouse rules, so every CFA and
 * customer (Category A included) is scaled by the same factor. With the month
 * filter on, only dates inside the projection month count.
 */
export function prepare(tabs, warehouses, { matchMonth = true, prorate = true } = {}) {
  const whMap = {};
  for (const t of TABS) {
    whMap[t] = new Map();
    for (const w of warehouses) if (w[t]) whMap[t].set(key(w[t]), w.code);
  }

  const projPeriods = new Set();
  for (const r of tabs.projection?.records || []) if (!r.total && r.period) projPeriods.add(r.period);
  const filterMonth = matchMonth && projPeriods.size > 0;

  const soDayset = new Set();
  for (const r of tabs.so?.records || []) {
    if (r.day && (!filterMonth || projPeriods.has(r.period))) soDayset.add(r.day);
  }
  const monthDays = [...projPeriods].reduce((s, p) => s + daysInMonth(p), 0);
  const pr = { on: false, factor: 1, soDays: soDayset.size, monthDays, why: "" };
  if (!prorate) pr.why = "Pro-rating is off — MTD projection equals the month projection.";
  else if (!tabs.so) pr.why = "No so tab, so there are no SO days — MTD projection falls back to the full month.";
  else if (!projPeriods.size) pr.why = "The projection has no readable Month / Year — MTD projection falls back to the full month.";
  else if (!soDayset.size) pr.why = "The so tab has no Sales Order Dates in the projection month — MTD projection falls back to the full month.";
  else Object.assign(pr, { on: true, factor: soDayset.size / monthDays });

  const ex = [];   // exclusion buckets: { tab, reason, rows, kg, kind }
  const bump = (tab, reason, kind, kg) => {
    let b = ex.find((e) => e.tab === tab && e.reason === reason);
    if (!b) ex.push((b = { tab, reason, kind, rows: 0, kg: 0 }));
    b.rows++; b.kg += kg;
  };
  const outside = new Map();  // tab → Map(label → kg) for warehouses not compared

  const rows = [];
  for (const tab of TABS) {
    const src = tabs[tab];
    if (!src) continue;
    const src3 = tab === "projection" ? "proj" : tab === "so" ? "so" : "dis";
    for (const r of src.records) {
      if (r.total) { bump(tab, "Grand-total row", "rule", r.qty); continue; }
      const catA = key(r.cgroup) === CAT_A_KEY;
      if (tab === "so" && catA) { bump(tab, "Customer Group = Category A (taken from dispatches instead)", "rule", r.qty); continue; }
      if (tab === "dispatches" && !catA) { bump(tab, "Sales Channel ≠ Category A (taken from the so tab instead)", "rule", r.qty); continue; }
      if (tab === "dispatches" && key(r.status) === "return") { bump(tab, "Invoice Status = Return", "rule", r.qty); continue; }
      const wh = whMap[tab].get(key(r.wh));
      if (!wh) {
        bump(tab, "Warehouse not in the Normalisation tab", "scope", r.qty);
        if (!outside.has(tab)) outside.set(tab, new Map());
        const m = outside.get(tab), lab = r.wh || "(blank)";
        m.set(lab, (m.get(lab) || 0) + r.qty);
        continue;
      }
      if (filterMonth && tab !== "projection") {
        if (!r.period) { bump(tab, "No readable date (kept)", "warn", r.qty); }
        else if (!projPeriods.has(r.period)) { bump(tab, `Outside the projection month (${ymLabel(r.period)})`, "month", r.qty); continue; }
      }
      rows.push({
        src: src3,
        wh,
        catA,
        cgroup: catA ? CAT_A : r.cgroup,
        cust: catA ? CAT_A_CUSTOMER : r.cust,
        rawCust: r.cust,
        code: r.code,
        group: r.group,
        parent: r.parent,
        name: r.name,
        qty: tab === "projection" ? r.qty * pr.factor : r.qty,
        fullQty: r.qty,
      });
    }
  }

  return { rows, exclusions: ex, outside, projPeriods: [...projPeriods].sort(), filterMonth, prorate: pr };
}

/* Labels for each item, keyed on Item Code: first non-blank of
 * projection → so → dispatches, level by level. */
function itemLabels(rows) {
  const rank = { proj: 0, so: 1, dis: 2 };
  const best = new Map();
  for (const r of rows) {
    const k = itemKey(r);
    let b = best.get(k);
    if (!b) best.set(k, (b = { group: null, parent: null, name: null, code: r.code }));
    for (const f of ["group", "parent", "name"]) {
      if (!r[f]) continue;
      if (!b[f] || rank[r.src] < b[f].rank) b[f] = { v: r[f], rank: rank[r.src] };
    }
  }
  const out = new Map();
  for (const [k, b] of best) {
    out.set(k, {
      group: b.group?.v || "(no MIS item group)",
      parent: b.parent?.v || "(no item parent)",
      name: b.name?.v || b.code || "(no item name)",
      code: b.code,
    });
  }
  return out;
}

const itemKey = (r) => (r.code ? "c:" + key(r.code) : "n:" + key(r.name));

/* Leaf rows for one view: customer leaves or item leaves, each with
 * proj / so (and so split by source) for the chosen warehouse. */
export function leaves(prepared, view, wh = "") {
  const rows = wh ? prepared.rows.filter((r) => r.wh === wh) : prepared.rows;
  const out = new Map();
  const display = new Map();   // key → first label seen, projection first
  const ordered = [...rows].sort((a, b) => (a.src === "proj" ? 0 : 1) - (b.src === "proj" ? 0 : 1));
  const labels = view === "item" ? itemLabels(prepared.rows) : null;

  for (const r of ordered) {
    let k, path;
    if (view === "customer") {
      const g = key(r.cgroup), c = key(r.cust);
      if (!display.has("g|" + g)) display.set("g|" + g, r.cgroup || "(no customer group)");
      if (!display.has("c|" + g + "|" + c)) display.set("c|" + g + "|" + c, r.cust || "(no customer)");
      k = g + "|" + c;
      path = [display.get("g|" + g), display.get("c|" + g + "|" + c)];
    } else {
      k = itemKey(r);
      const l = labels.get(k);
      path = [l.group, l.parent, l.name];
    }
    let leaf = out.get(k);
    if (!leaf) out.set(k, (leaf = { path, code: view === "item" ? labels.get(k).code : "", proj: 0, projFull: 0, so: 0, soTab: 0, soDis: 0, custs: { proj: new Set(), so: new Set() } }));
    if (r.src === "proj") { leaf.proj += r.qty; leaf.projFull += r.fullQty; leaf.custs.proj.add(r.rawCust); }
    else {
      leaf.so += r.qty;
      if (r.src === "so") leaf.soTab += r.qty; else leaf.soDis += r.qty;
      leaf.custs.so.add(r.rawCust);
    }
  }
  for (const l of out.values()) l.gap = l.proj - l.so;
  return [...out.values()];
}

const EPS = 0.0005;
export const classify = (gap) => (gap > EPS ? "over" : gap < -EPS ? "under" : "match");

/* Nest leaves into a tree by their path. Every node carries its totals. */
export function tree(leafRows, sortBy = "gap", mode = "all") {
  const root = { label: "Total", children: new Map(), proj: 0, projFull: 0, so: 0, soDis: 0, n: 0, depth: -1 };
  for (const l of leafRows) {
    let node = root;
    node.proj += l.proj; node.projFull += l.projFull; node.so += l.so; node.soDis += l.soDis; node.n++;
    l.path.forEach((label, d) => {
      const last = d === l.path.length - 1;
      const k = last ? label + "\u0000" + (l.code || "") : label;
      let c = node.children.get(k);
      if (!c) node.children.set(k, (c = { label, children: new Map(), proj: 0, projFull: 0, so: 0, soDis: 0, n: 0, depth: d, leaf: last ? l : null }));
      c.proj += l.proj; c.projFull += l.projFull; c.so += l.so; c.soDis += l.soDis; c.n++;
      node = c;
    });
  }
  const cmp = {
    gap:   mode === "under" ? (a, b) => a.gap - b.gap
         : mode === "over"  ? (a, b) => b.gap - a.gap
         :                    (a, b) => Math.abs(b.gap) - Math.abs(a.gap),
    proj:  (a, b) => b.proj - a.proj,
    so:    (a, b) => b.so - a.so,
    name:  (a, b) => a.label.localeCompare(b.label),
  }[sortBy];
  const fin = (n) => {
    n.gap = n.proj - n.so;
    n.kids = [...n.children.values()];
    n.kids.forEach(fin);
    n.kids.sort(cmp);
    delete n.children;
    return n;
  };
  return fin(root);
}
