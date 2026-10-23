/* Projection vs SO — the Excel export, in the format of
 * projection-analytics-dashboard's workbook:
 *
 *   All Conditions Applied  row-level data with per-row diff / % / verdict
 *                           formulas; a FILTERS APPLIED block to its right
 *   Customer Pivot          Customer Group › Customer › Item Group ›
 *                           Item Parent › Item Name
 *   Item Pivot              Item Group › Item Parent › Item Name ›
 *                           Customer Group › Customer
 *   Over-Projected          Item Group › Item Parent › Item Name, MTD proj > SO
 *   Under-Projected         Item Group › Item Parent › Item Name, SO > MTD proj
 *
 * As in the reference: written with xlsx-js-style; pivot leaves are SUMIFS on
 * the data sheet and parents SUM their direct children, each with its cached
 * result; label columns, then Level, then the figures; children collapsed
 * under Excel's outline + buttons; an autofilter over every sheet's full range.
 *
 * Formulas stay plain ASCII and every SUM takes at most 250 arguments
 * (larger parents nest), so nothing in them can stop Excel parsing a sheet.
 *
 * Colours: indigo header, one tint per depth, bold total row, red → amber →
 * green on Mismatch % as in the reference; over / under in the dashboard's
 * blue / red (gap, ratios, Status).
 */

import { DISPATCH_GROUPS, TAB_LABEL, leaves, tree, classify, ymLabel, rowLabeler } from "./calc.js?v=11";

const DATA_SHEET = "All Conditions Applied";
const PCT_Z = "0.0%", KG_Z = "#,##0.00";

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

function addSheet(wb, name, grid, widths, rowMeta) {
  const ws = buildSheet(grid);
  const nCols = (grid[0] || []).length;
  ws["!cols"] = widths;
  if (grid.length > 1) {
    ws["!autofilter"] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: grid.length - 1, c: Math.max(0, nCols - 1) } }) };
  }
  if (rowMeta) {
    ws["!rows"] = rowMeta;
    ws["!outline"] = { above: true };   // the summary row sits above its detail
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
  "target_month", "cfa", "customer_group", "customer", "item_group", "item_parent", "item_name", "item_code",
  "row_type", "source_customer", "month_proj_kgs", "mtd_proj_kgs", "so_kgs", "dispatch_kgs",
];
const DERIVED = ["diff_kgs", "so_over_proj_pct", "proj_over_so_pct", "proj_vs_so_pct", "mismatch_pct", "verdict"];
const DC = Object.fromEntries(DATA_FIELDS.map((f, i) => [f, null]));   // letters filled in below
const ROW_TYPE = { proj: "Projection", so: "SO", dis: "Dispatch" };

/* ── pivots ──────────────────────────────────────────────────────── */

/* Our trees (calc.tree) → one node shape: label, code, the four sums, children. */
function node(n, children, code) {
  return { label: n.label, code: code || n.leaf?.code || "", month: n.projFull, mtd: n.proj, so: n.so, dis: n.soDis, children };
}

/* SUMIFS criteria. Labels are matched by cell reference, as in the reference;
 * one Excel would read as a wildcard or comparison goes in as an escaped
 * literal. The item code has no column of its own, so it is always a literal. */
const unsafe = (s) => /[*?~]|^[<>=]/.test(s);
const lit = (s) => '"=' + s.replace(/[~*?]/g, "~$&").replace(/"/g, '""') + '"';

/* SUM of explicit cell references. Excel allows 255 arguments per function,
 * so a parent with more children sums them in nested chunks of 250. */
function sumOf(refs) {
  if (refs.length <= 250) return `SUM(${refs.join(",")})`;
  const parts = [];
  for (let i = 0; i < refs.length; i += 250) parts.push(`SUM(${refs.slice(i, i + 250).join(",")})`);
  return `SUM(${parts.join(",")})`;
}

/* A PIVOT sheet, N levels deep. Each node is followed by its children one
 * outline level deeper, all collapsed, so Excel draws a + beside every
 * expandable row. A node SUMs its direct child rows by explicit reference;
 * leaves are SUMIFS on their full path (plus the item code where there is
 * one). The grand total sums the root rows only. */
function pivotSheet(roots, levels, lastRow) {
  const col = (i) => XLSX.utils.encode_col(i);
  const q = (c) => `'${DATA_SHEET}'!$${c}$2:$${c}$${lastRow}`;
  const N_L = levels.length;
  const cMo = col(N_L + 1), cMtd = col(N_L + 2), cSo = col(N_L + 3), cDis = col(N_L + 4);

  const grid = [[...levels.map((l) => l.label), "Level", "Month Proj kgs", "MTD Proj kgs", "SO kgs",
    "Dispatch kgs", "Gap kgs", "SO / Proj %", "Proj / SO %", "Proj vs SO %", "Mismatch %", "Status"].map((h) => T(h, S_HEAD))];
  const meta = [{}];

  const derived = (R, a, base) => {
    const gap = a.mtd - a.so;
    const den = Math.max(a.so, a.mtd);
    const cls = a.mtd > a.so ? "over" : a.mtd < a.so ? "under" : "match";   // as the formula decides
    const projOverSo = a.so > 0 ? a.mtd / a.so : null;
    return [
      F(`${cMtd}${R}-${cSo}${R}`, KG_Z, gap, sTint(base, GAP_RGB(gap))),
      F(`IF(${cMtd}${R}=0,"",${cSo}${R}/${cMtd}${R})`, PCT_Z, a.mtd > 0 ? a.so / a.mtd : null, sTint(base, RATIO_RGB(projOverSo))),
      F(`IF(${cSo}${R}=0,"",${cMtd}${R}/${cSo}${R})`, PCT_Z, projOverSo, sTint(base, RATIO_RGB(projOverSo))),
      F(`IF(${cSo}${R}=0,"",(${cMtd}${R}-${cSo}${R})/${cSo}${R})`, PCT_Z, a.so > 0 ? gap / a.so : null, sTint(base, GAP_RGB(gap))),
      F(`IF(MAX(${cSo}${R},${cMtd}${R})=0,0,ABS(${cMtd}${R}-${cSo}${R})/MAX(${cSo}${R},${cMtd}${R}))`, PCT_Z, den > 0 ? Math.abs(gap) / den : 0, sMismatch(base, den > 0 ? Math.abs(gap) / den : 0)),
      F(`IF(${cMtd}${R}>${cSo}${R},"${STATUS.over.text}",IF(${cMtd}${R}<${cSo}${R},"${STATUS.under.text}","${STATUS.match.text}"))`, null, STATUS[cls].text, sStatus(base, cls)),
    ];
  };

  let R = 2;
  function emit(n, depth, path, code) {
    const base = sDepth(depth, n.children.length > 0);
    const myRow = R++;
    const pth = [...path, n.label];
    const itemCode = n.code || code;
    const placeholder = [];
    grid.push(placeholder);
    meta.push(depth ? { level: depth, hidden: true } : {});

    const childRows = n.children.map((c) => emit(c, depth + 1, pth, itemCode));
    const isLeaf = !n.children.length;
    const sum = (c) => sumOf(childRows.map((r) => `${c}${r}`));
    const crit = [
      ...pth.map((v, i) => `${q(DC[levels[i].key])},${unsafe(v) ? lit(v) : `$${col(i)}${myRow}`}`),
      ...(itemCode ? [`${q(DC.item_code)},${lit(itemCode)}`] : []),
    ].join(",");
    const agg = (c, field, val) => (isLeaf ? F(`SUMIFS(${q(DC[field])},${crit})`, KG_Z, val, base) : F(sum(c), KG_Z, val, base));

    placeholder.push(
      ...levels.map((l, i) => T(i < pth.length ? pth[i] : "", base)),
      T(isLeaf ? "Detail" : `L${depth + 1}`, base),
      agg(cMo, "month_proj_kgs", n.month),
      agg(cMtd, "mtd_proj_kgs", n.mtd),
      agg(cSo, "so_kgs", n.so),
      agg(cDis, "dispatch_kgs", n.dis),
      ...derived(myRow, n, base),
    );
    return myRow;
  }
  const rootRows = roots.map((n) => emit(n, 0, [], ""));

  if (rootRows.length) {
    const t = roots.reduce((s, n) => ({ month: s.month + n.month, mtd: s.mtd + n.mtd, so: s.so + n.so, dis: s.dis + n.dis }), { month: 0, mtd: 0, so: 0, dis: 0 });
    const rs = (c) => sumOf(rootRows.map((r) => `${c}${r}`));
    const TR = R;
    grid.push([
      T("TOTAL", S_TOTAL), ...levels.slice(1).map(() => T("", S_TOTAL)), T("Total", S_TOTAL),
      F(rs(cMo), KG_Z, t.month, S_TOTAL), F(rs(cMtd), KG_Z, t.mtd, S_TOTAL),
      F(rs(cSo), KG_Z, t.so, S_TOTAL), F(rs(cDis), KG_Z, t.dis, S_TOTAL),
      ...derived(TR, t, S_TOTAL),
    ]);
    meta.push({});
  }
  return { grid, meta };
}

/* ── the workbook ────────────────────────────────────────────────── */

export function exportWorkbook({ prepared, wh, sort = "gap" }) {
  if (!window.XLSX?.utils) throw new Error("The Excel writer did not load — check the connection and reload");
  DATA_FIELDS.forEach((f, i) => { DC[f] = XLSX.utils.encode_col(i); });
  const wb = XLSX.utils.book_new();
  const label = rowLabeler(prepared, wh);
  const rows = wh ? prepared.rows.filter((r) => r.wh === wh) : prepared.rows;
  const pr = prepared.prorate;

  /* 1. Row-level data — the only sheet with literal kg; everything else points here. */
  const lastRow = Math.max(2, rows.length + 1);
  const base = { border: BRD, font: { color: { rgb: "FF16192B" } } };
  const dataGrid = [[...DATA_FIELDS, ...DERIVED].map((h) => T(h, S_HEAD))];
  const { mtd_proj_kgs: MT, so_kgs: SO } = DC;
  rows.forEach((r, i) => {
    const R = i + 2;
    const l = label(r);
    const month = r.src === "proj" ? r.fullQty : 0, mtd = r.src === "proj" ? r.qty : 0;
    const so = r.src === "proj" ? 0 : r.qty, dis = r.src === "dis" ? r.qty : 0;
    const diff = mtd - so, den = Math.max(so, mtd);
    const verdict = mtd > so ? "Over-projected" : mtd < so ? "Under-projected" : "On target";
    dataGrid.push([
      T(r.period ? ymLabel(r.period) : "", base), T(`CFA (${r.wh})`, base), T(l.cgroup, base), T(l.cust, base),
      T(l.group, base), T(l.parent, base), T(l.name, base), T(l.code, base), T(ROW_TYPE[r.src], base), T(r.rawCust, base),
      N(month, KG_Z, base), N(mtd, KG_Z, base), N(so, KG_Z, base), N(dis, KG_Z, base),
      F(`${MT}${R}-${SO}${R}`, KG_Z, diff, sTint(base, GAP_RGB(diff))),
      F(`IF(${MT}${R}=0,"",${SO}${R}/${MT}${R})`, PCT_Z, mtd > 0 ? so / mtd : null, sTint(base, RATIO_RGB(so > 0 ? mtd / so : null))),
      F(`IF(${SO}${R}=0,"",${MT}${R}/${SO}${R})`, PCT_Z, so > 0 ? mtd / so : null, sTint(base, RATIO_RGB(so > 0 ? mtd / so : null))),
      F(`IF(${SO}${R}=0,"",(${MT}${R}-${SO}${R})/${SO}${R})`, PCT_Z, so > 0 ? diff / so : null, sTint(base, GAP_RGB(diff))),
      F(`IF(MAX(${SO}${R},${MT}${R})=0,0,ABS(${MT}${R}-${SO}${R})/MAX(${SO}${R},${MT}${R}))`, PCT_Z, den > 0 ? Math.abs(diff) / den : 0,
        sMismatch(base, den > 0 ? Math.abs(diff) / den : 0)),
      F(`IF(${MT}${R}>${SO}${R},"Over-projected",IF(${MT}${R}<${SO}${R},"Under-projected","On target"))`, null, verdict, base),
    ]);
  });
  const wide = { target_month: 14, cfa: 11, customer_group: 18, customer: 30, item_group: 24, item_parent: 30, item_name: 46, item_code: 20, row_type: 18, source_customer: 30 };
  const dws = addSheet(wb, DATA_SHEET, dataGrid, [...DATA_FIELDS.map((f) => ({ wch: wide[f] || 14 })), ...DERIVED.map(() => ({ wch: 15 }))]);

  /* 2. FILTERS APPLIED, parked to the right of the data after one blank
        column — the reference's block, with this dashboard's filters. */
  const period = prepared.projPeriods.map(ymLabel).join(", ") || "All";
  const sourceRows = prepared.rows.length + prepared.exclusions.filter((e) => e.kind !== "warn").reduce((a, e) => a + e.rows, 0);
  const prov = [
    ["FILTERS APPLIED", ""],
    ["Target Month", period],
    ["CFA", wh ? `CFA (${wh})` : "All"],
    ["Projection Status", "All"],
    ["Customer Group", "All"],
    ["Customer", "All"],
    ["Item Group", "All"],
    ["Item Parent", "All"],
    ["Item Name", "All"],
    ["Only the projection month", prepared.filterMonth ? "Yes" : "No"],
    ["SO from dispatches for", DISPATCH_GROUPS.join(", ")],
    ["SO days / days in month", pr.on ? `${pr.soDays} / ${pr.monthDays}` : `- / ${pr.monthDays || "-"}`],
    ["MTD projection", pr.on ? `Month projection / ${pr.monthDays} x ${pr.soDays} unique Sales Order Dates (${Math.round(pr.factor * 1000) / 10}%)` : pr.why],
    ["Rows after filters", String(rows.length)],
    ["Rows in source file", String(sourceRows)],
    ["Generated", new Date().toLocaleString("en-IN")],
    ["", ""],
    ["ROWS LEFT OUT", ""],
    ...prepared.exclusions.map((e) => [`${TAB_LABEL[e.tab]} - ${e.reason}`, `${e.rows} rows, ${Math.round(e.kg)} kg`]),
  ];
  const PROV_C = DATA_FIELDS.length + DERIVED.length + 1;   // one blank column gap
  prov.forEach((row, i) => row.forEach((v, j) => {
    const cell = { t: "s", v: String(v) };
    if (row[0] === "FILTERS APPLIED" || row[0] === "ROWS LEFT OUT") cell.s = S_HEAD;
    dws[XLSX.utils.encode_cell({ r: i, c: PROV_C + j })] = cell;
  }));
  const dr = XLSX.utils.decode_range(dws["!ref"]);
  dr.e.c = Math.max(dr.e.c, PROV_C + 1);
  dr.e.r = Math.max(dr.e.r, prov.length - 1);
  dws["!ref"] = XLSX.utils.encode_range(dr);
  dws["!cols"][PROV_C] = { wch: 26 };
  dws["!cols"][PROV_C + 1] = { wch: 60 };

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
    const { grid, meta } = pivotSheet(roots, levels, lastRow);
    addSheet(wb, name, grid, [...levels.map((l) => ({ wch: wch[l.key] })), { wch: 9 },
      { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 13 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 11 }, { wch: 16 }], meta);
  }

  const tag = [prepared.projPeriods[0] || "", wh || "all"].filter(Boolean).join("-");
  XLSX.writeFile(wb, `proj-vs-so-${tag}.xlsx`, { compression: true });
}
