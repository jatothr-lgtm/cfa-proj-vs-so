/* Projection vs SO — the Excel export.
 *
 * Tabs as in projection-analytics-dashboard:
 *
 *   All Conditions Applied  row-level data; filters, SO days and rows left
 *                           out sit to the right of it (after a gap column)
 *   Customer Pivot          Customer Group › Customer › Item Group ›
 *                           Item Parent › Item Name
 *   Item Pivot              Item Group › Item Parent › Item Name ›
 *                           Customer Group › Customer
 *   Over-Projected          Item Group › Item Parent › Item Name, MTD proj > SO
 *   Under-Projected         Item Group › Item Parent › Item Name, SO > MTD proj
 *
 * Written with xlsx-js-style (SheetJS 0.18.5 plus cell styles).
 *
 * Why the pivots are plain values: desktop Excel discarded every sheet that
 * used row outline grouping ("Replaced Part … XML error. Load error") — both
 * from this writer and from ExcelJS — while the row-level sheet, which has
 * none, always opened. So the pivots carry only what that sheet proves Excel
 * reads: styled text and numbers and a filter row. Hierarchy reads from the
 * repeated level columns, the Level column and the depth tints; filter Level
 * to collapse. The row-level sheet keeps its in-row formulas.
 *
 * Colours: indigo header, one tint per depth, bold total row, red → amber →
 * green on Mismatch % as in the reference; over / under in the dashboard's
 * blue / red (gap, ratios, Status).
 */

import { CAT_A, TAB_LABEL, leaves, tree, classify, ymLabel, rowLabeler } from "./calc.js?v=9";

const DATA_SHEET = "All Conditions Applied";
const PCT_Z = "0.0%", KG_Z = "#,##0.00", INT_Z = "#,##0";

/* ── cells ───────────────────────────────────────────────────────── */

const F = (f, z, v, s) => ({ f, z, v, s });          // formula + cached result
const N = (v, z, s) => ({ v, z, s });                // literal number
const T = (v, s) => ({ txt: v == null ? "" : String(v), s });

function buildSheet(grid) {
  const ws = {};
  let maxC = 0;
  grid.forEach((row, R) => row.forEach((val, C) => {
    if (C > maxC) maxC = C;
    const addr = XLSX.utils.encode_cell({ r: R, c: C });
    if (val == null || val === "") { ws[addr] = { t: "s", v: "" }; return; }
    if (typeof val === "object" && val.f !== undefined) {
      // The cell type follows the cached result; Excel recalculates on open.
      if (typeof val.v === "number" && Number.isFinite(val.v)) {
        ws[addr] = { t: "n", v: val.v, f: val.f };
        if (val.z) ws[addr].z = val.z;
      } else ws[addr] = { t: "s", v: val.v == null ? "" : String(val.v), f: val.f };
    } else if (typeof val === "object" && val.txt !== undefined) {
      ws[addr] = { t: "s", v: val.txt };
    } else if (typeof val === "object" && val.v !== undefined) {
      ws[addr] = { t: "n", v: val.v };
      if (val.z) ws[addr].z = val.z;
    } else if (typeof val === "number") ws[addr] = { t: "n", v: val };
    else ws[addr] = { t: "s", v: String(val) };
    if (val && typeof val === "object" && val.s) ws[addr].s = val.s;
  }));
  ws["!ref"] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: Math.max(0, grid.length - 1), c: maxC } });
  return ws;
}

function addSheet(wb, name, grid, widths) {
  const ws = buildSheet(grid);
  const nCols = (grid[0] || []).length;
  ws["!cols"] = widths;
  if (grid.length > 1) {
    ws["!autofilter"] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: grid.length - 1, c: Math.max(0, nCols - 1) } }) };
  }
  XLSX.utils.book_append_sheet(wb, ws, name.slice(0, 31));
  return ws;
}

/* ── styles ──────────────────────────────────────────────────────── */

const BRD = {
  top: { style: "thin", color: { rgb: "FFDDE1EC" } }, bottom: { style: "thin", color: { rgb: "FFDDE1EC" } },
  left: { style: "thin", color: { rgb: "FFDDE1EC" } }, right: { style: "thin", color: { rgb: "FFDDE1EC" } },
};
const S_HEAD = {
  fill: { fgColor: { rgb: "FF4F46E5" } }, font: { color: { rgb: "FFFFFFFF" }, bold: true, sz: 10 },
  alignment: { horizontal: "center", vertical: "center", wrapText: true }, border: BRD,
};
const S_TOTAL = { fill: { fgColor: { rgb: "FFDDE1EC" } }, font: { bold: true, color: { rgb: "FF16192B" } }, border: BRD };
const DEPTH_FILL = ["FFEEF0F6", "FFFFFFFF", "FFF7F8FC", "FFFFFFFF"];

const sDepth = (depth, bold) => ({
  fill: { fgColor: { rgb: DEPTH_FILL[Math.min(depth, 3)] } },
  font: { bold: !!bold, color: { rgb: depth ? "FF4A5060" : "FF16192B" } },
  border: BRD,
});
const sTint = (base, rgb, bold) => ({ ...base, font: { color: { rgb }, bold: bold === undefined ? base.font?.bold : bold } });
function sMismatch(base, pct) {
  const p = Number.isFinite(pct) ? pct : 0;
  if (p >= 0.5) return { ...base, fill: { fgColor: { rgb: "FFFBDDDD" } }, font: { color: { rgb: "FFB91C1C" }, bold: true } };
  if (p >= 0.25) return { ...base, fill: { fgColor: { rgb: "FFFDEEDA" } }, font: { color: { rgb: "FFB45309" }, bold: false } };
  return { ...base, fill: { fgColor: { rgb: "FFE7F7EC" } }, font: { color: { rgb: "FF15803D" }, bold: false } };
}
// Direction, as on the dashboard: over-projected blue, under-projected red.
const OVER = "FF1D4F91", UNDER = "FFC0392F", FLAT = "FF6B7186";
const GAP_RGB = (d) => (d > 0.0005 ? OVER : d < -0.0005 ? UNDER : FLAT);
const RATIO_RGB = (r) => (r == null || !Number.isFinite(r) ? FLAT : r > 1.05 ? OVER : r < 0.95 ? UNDER : "FF15803D");
// Plain ASCII on purpose: a formula Excel cannot parse makes it discard the
// whole sheet ("Replaced Part … Load error"), and the status formula is the
// verdict formula the row-level sheet already uses.
const STATUS = {
  over: { text: "Over-projected", fill: "FFE8EFF9", font: OVER },
  under: { text: "Under-projected", fill: "FFFDECEA", font: UNDER },
  match: { text: "On target", fill: "FFE8F5EE", font: "FF1F7D53" },
};
const sStatus = (base, cls) => ({ ...base, fill: { fgColor: { rgb: STATUS[cls].fill } }, font: { color: { rgb: STATUS[cls].font }, bold: true }, alignment: { horizontal: "center" } });

/* ── the row-level sheet ─────────────────────────────────────────── */

const DATA_FIELDS = [
  "cfa", "customer_group", "customer", "item_group", "item_parent", "item_name", "item_code",
  "row_type", "source_customer", "month_proj_kgs", "mtd_proj_kgs", "so_kgs", "cat_a_dispatch_kgs",
];
const DC = Object.fromEntries(DATA_FIELDS.map((f, i) => [f, null]));   // letters filled in below
const ROW_TYPE = { proj: "Projection", so: "SO", dis: `${CAT_A} dispatch` };

/* ── pivots ──────────────────────────────────────────────────────── */

/* Our trees (calc.tree) → one node shape: label, code, the four sums, children. */
function node(n, children, code) {
  return { label: n.label, code: code || n.leaf?.code || "", month: n.projFull, mtd: n.proj, so: n.so, dis: n.soDis, children };
}

/* A PIVOT sheet, written as values. One column per level with the ancestors
 * repeated, a Level column (L1 … Detail) and every parent directly above its
 * children, as in the reference — but no Excel outline grouping and no
 * cross-sheet formulas: the pivot sheets use only what the row-level sheet
 * already proves Excel opens (styled text and numbers, a filter row).
 * Filter the Level column to collapse a sheet to a level. */
function pivotSheet(roots, levels) {
  const grid = [[...levels.map((l) => l.label), "Item Code", "Level", "Month Proj kgs", "MTD Proj kgs", "SO kgs",
    `${CAT_A} Disp kgs`, "Gap kgs", "SO / Proj %", "Proj / SO %", "Proj vs SO %", "Mismatch %", "Status"].map((h) => T(h, S_HEAD))];

  const pct = (v, s) => (v == null || !Number.isFinite(v) ? T("", s) : N(v, PCT_Z, s));
  const numbers = (a, base) => {
    const gap = a.mtd - a.so;
    const den = Math.max(a.so, a.mtd);
    const projOverSo = a.so > 0 ? a.mtd / a.so : null;
    const cls = classify(gap);
    return [
      N(a.month, KG_Z, sTint(base, "FF8B9992")),
      N(a.mtd, KG_Z, base),
      N(a.so, KG_Z, base),
      N(a.dis, KG_Z, sTint(base, "FF8B9992")),
      N(gap, KG_Z, sTint(base, GAP_RGB(gap), true)),
      pct(a.mtd > 0 ? a.so / a.mtd : null, sTint(base, RATIO_RGB(projOverSo))),
      pct(projOverSo, sTint(base, RATIO_RGB(projOverSo))),
      pct(a.so > 0 ? gap / a.so : null, sTint(base, GAP_RGB(gap))),
      N(den > 0 ? Math.abs(gap) / den : 0, PCT_Z, sMismatch(base, den > 0 ? Math.abs(gap) / den : 0)),
      T(STATUS[cls].text, sStatus(base, cls)),
    ];
  };

  function emit(n, depth, path, code) {
    const base = sDepth(depth, n.children.length > 0);
    const pth = [...path, n.label];
    const itemCode = n.code || code;
    grid.push([
      ...levels.map((l, i) => T(i < pth.length ? pth[i] : "", base)),
      T(itemCode, sTint(base, "FF8B9992", false)),
      T(n.children.length ? `L${depth + 1}` : "Detail", base),
      ...numbers(n, base),
    ]);
    n.children.forEach((c) => emit(c, depth + 1, pth, itemCode));
  }
  roots.forEach((n) => emit(n, 0, [], ""));

  if (roots.length) {
    const t = roots.reduce((s, n) => ({ month: s.month + n.month, mtd: s.mtd + n.mtd, so: s.so + n.so, dis: s.dis + n.dis }), { month: 0, mtd: 0, so: 0, dis: 0 });
    grid.push([T("TOTAL", S_TOTAL), ...levels.slice(1).map(() => T("", S_TOTAL)), T("", S_TOTAL), T("Total", S_TOTAL), ...numbers(t, S_TOTAL)]);
  }
  return grid;
}

/* ── the workbook ────────────────────────────────────────────────── */

export function exportWorkbook({ prepared, wh, sort = "gap" }) {
  if (!window.XLSX?.utils) throw new Error("The Excel writer did not load — check the connection and reload");
  DATA_FIELDS.forEach((f, i) => { DC[f] = XLSX.utils.encode_col(i); });
  const wb = XLSX.utils.book_new();
  const label = rowLabeler(prepared, wh);
  const rows = wh ? prepared.rows.filter((r) => r.wh === wh) : prepared.rows;
  const pr = prepared.prorate;

  /* 1. Row-level data. */
  const base = { border: BRD, font: { color: { rgb: "FF16192B" } } };
  const derivedHeads = ["gap_kgs", "verdict"];
  const dataGrid = [[...DATA_FIELDS, ...derivedHeads].map((h) => T(h, S_HEAD))];
  const { mtd_proj_kgs: MT, so_kgs: SO } = DC;
  rows.forEach((r, i) => {
    const R = i + 2;
    const l = label(r);
    const month = r.src === "proj" ? r.fullQty : 0, mtd = r.src === "proj" ? r.qty : 0;
    const so = r.src === "proj" ? 0 : r.qty, dis = r.src === "dis" ? r.qty : 0;
    const gap = mtd - so;
    dataGrid.push([
      T(`CFA (${r.wh})`, base), T(l.cgroup, base), T(l.cust, base), T(l.group, base), T(l.parent, base), T(l.name, base), T(l.code, base),
      T(ROW_TYPE[r.src], base), T(r.rawCust, base),
      N(month, KG_Z, base), N(mtd, KG_Z, base), N(so, KG_Z, base), N(dis, KG_Z, base),
      F(`${MT}${R}-${SO}${R}`, KG_Z, gap, sTint(base, GAP_RGB(gap))),
      F(`IF(${MT}${R}>${SO}${R},"Over-projected",IF(${MT}${R}<${SO}${R},"Under-projected","On target"))`, null,
        gap > 0 ? "Over-projected" : gap < 0 ? "Under-projected" : "On target", base),
    ]);
  });
  const wide = { cfa: 11, customer_group: 18, customer: 30, item_group: 24, item_parent: 30, item_name: 46, item_code: 20, row_type: 18, source_customer: 30 };
  const dws = addSheet(wb, DATA_SHEET, dataGrid, [...DATA_FIELDS.map((f) => ({ wch: wide[f] || 14 })), ...derivedHeads.map(() => ({ wch: 15 }))]);

  /* 2. Filters, SO days and rows left out, parked to the right of the data
        so they need no tab of their own and touch neither formulas nor filter. */
  const period = prepared.projPeriods.map(ymLabel).join(", ") || "—";
  const prov = [
    ["FILTERS APPLIED", ""],
    ["CFA", wh ? `CFA (${wh})` : "All CFAs"],
    ["Projection month", period],
    ["Only the projection month", prepared.filterMonth ? "Yes" : "No"],
    ["SO days / days in month", pr.on ? `${pr.soDays} / ${pr.monthDays}` : `— / ${pr.monthDays || "—"}`],
    ["MTD projection", pr.on ? `Month projection ÷ ${pr.monthDays} × ${pr.soDays} unique Sales Order Dates (${Math.round(pr.factor * 1000) / 10}%)` : pr.why],
    ["SO", `so tab (Customer Group ≠ ${CAT_A}) + dispatches (Sales Channel = ${CAT_A}, Invoice Status ≠ Return)`],
    ["Rows in this sheet", String(rows.length)],
    ["Generated", new Date().toLocaleString("en-IN")],
    ["", ""],
    ["ROWS LEFT OUT", ""],
    ...prepared.exclusions.map((e) => [`${TAB_LABEL[e.tab]} · ${e.reason}`, `${e.rows.toLocaleString("en-IN")} rows · ${Math.round(e.kg).toLocaleString("en-IN")} kg`]),
  ];
  const PROV_C = DATA_FIELDS.length + derivedHeads.length + 1;   // one blank column gap
  prov.forEach((row, i) => row.forEach((v, j) => {
    const cell = { t: "s", v: String(v) };
    if (row[0] === "FILTERS APPLIED" || row[0] === "ROWS LEFT OUT") cell.s = S_HEAD;
    else if (j === 0 && v) cell.s = { font: { bold: true, color: { rgb: "FF4A5060" } } };
    dws[XLSX.utils.encode_cell({ r: i, c: PROV_C + j })] = cell;
  }));
  const dr = XLSX.utils.decode_range(dws["!ref"]);
  dr.e.c = Math.max(dr.e.c, PROV_C + 1);
  dr.e.r = Math.max(dr.e.r, prov.length - 1);
  dws["!ref"] = XLSX.utils.encode_range(dr);
  dws["!cols"][PROV_C] = { wch: 44 };
  dws["!cols"][PROV_C + 1] = { wch: 70 };

  /* 3. Pivots — every expandable row opens the next level down. */
  const cust = leaves(prepared, "customer", wh);
  const items = leaves(prepared, "item", wh);

  // Customer › Item: the dashboard's customer rows, each opened into its items.
  const itemNodes = (t) => t.kids.map((k) => (k.leaf ? node(k, [], k.leaf.code) : node(k, itemNodes(k))));
  const custRoots = tree(cust, sort, "all").kids.map((g) => node(g, g.kids.map((c) => node(c, itemNodes(tree(c.leaf.items, sort, "all"))))));

  // Item › Customer: each item opened into the customers who carry it.
  const byItem = new Map();
  const ik = (it) => (it.code ? "c:" + it.code.toLowerCase() : "n:" + it.path[2].toLowerCase());
  for (const l of cust) for (const it of l.items) {
    if (!byItem.has(ik(it))) byItem.set(ik(it), []);
    byItem.get(ik(it)).push({ path: l.path, code: "", proj: it.proj, projFull: it.projFull, so: it.so, soTab: it.soTab, soDis: it.soDis, gap: it.gap });
  }
  const custUnder = (leaf) => tree(byItem.get(ik({ code: leaf.code, path: leaf.path })) || [], sort, "all").kids
    .map((g) => node(g, g.kids.map((c) => node(c, []))));
  const itemTree = (ls, mode, withCustomers) => tree(ls, mode === "all" ? sort : "gap", mode).kids.map((g) =>
    node(g, g.kids.map((p) => node(p, p.kids.map((n) => node(n, withCustomers ? custUnder(n.leaf) : [], n.leaf.code))))));

  const P = [
    ["Customer Pivot", [["customer_group", "Customer Group"], ["customer", "Customer"], ["item_group", "Item Group"], ["item_parent", "Item Parent"], ["item_name", "Item Name"]], custRoots],
    ["Item Pivot", [["item_group", "Item Group"], ["item_parent", "Item Parent"], ["item_name", "Item Name"], ["customer_group", "Customer Group"], ["customer", "Customer"]], itemTree(items, "all", true)],
    ["Over-Projected", [["item_group", "Item Group"], ["item_parent", "Item Parent"], ["item_name", "Item Name"]], itemTree(items.filter((l) => classify(l.gap) === "over"), "over", false)],
    ["Under-Projected", [["item_group", "Item Group"], ["item_parent", "Item Parent"], ["item_name", "Item Name"]], itemTree(items.filter((l) => classify(l.gap) === "under"), "under", false)],
  ];
  const wch = { customer_group: 18, customer: 30, item_group: 24, item_parent: 30, item_name: 46 };
  for (const [name, lv, roots] of P) {
    const levels = lv.map(([key, lab]) => ({ key, label: lab }));
    const grid = pivotSheet(roots, levels);
    addSheet(wb, name, grid, [...levels.map((l) => ({ wch: wch[l.key] })), { wch: 18 }, { wch: 8 },
      { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 13 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 11 }, { wch: 16 }]);
  }

  const tag = [prepared.projPeriods[0] || "", wh || "all"].filter(Boolean).join("-");
  XLSX.writeFile(wb, `proj-vs-so-${tag}.xlsx`, { compression: true });
}
