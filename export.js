/* Projection vs SO — the Excel export, styled like the dashboard.
 *
 * ExcelJS (loaded on first export) rather than SheetJS, because the
 * community SheetJS build cannot write fills, fonts or outline levels.
 *
 * Three tabs:
 *   Summary   period, SO days, the tiles, the rules, rows left out
 *   Customer  Customer Group › Customer › Item Group › Item Parent › Item Name
 *   Item      Item Group › Item Parent › Item Name
 * The breakdowns are Excel outline levels (the 1–5 buttons). Over / Under is
 * the colour-coded Status column, filterable, rather than tabs of its own.
 *
 * Colours match the dashboard: over is blue, under is red, group rows shaded.
 */

import { CAT_A, TAB_LABEL, leaves, tree, classify, ymLabel } from "./calc.js";

const EXCELJS = "https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js";

let loading;
export function loadExcelJS() {
  if (window.ExcelJS) return Promise.resolve(window.ExcelJS);
  return (loading ||= new Promise((ok, fail) => {
    const s = document.createElement("script");
    s.src = EXCELJS;
    s.onload = () => (window.ExcelJS ? ok(window.ExcelJS) : fail(new Error("ExcelJS did not initialise")));
    s.onerror = () => { loading = null; fail(new Error("Could not load the Excel writer — check the connection")); };
    document.head.appendChild(s);
  }));
}

/* ── palette (the dashboard's light tokens, as ARGB) ─────────────── */

const C = {
  brand: "FF17603F", white: "FFFFFFFF",
  ink: "FF16241D", ink2: "FF5B6B63", ink3: "FF8B9992",
  line: "FFE2E7E3", shade0: "FFF1F4F2", shade1: "FFFAFBFA",
  over: "FF1D4F91", overBg: "FFE8EFF9",
  under: "FFC0392F", underBg: "FFFDECEA",
  match: "FF1F7D53", matchBg: "FFE8F5EE",
  warnBg: "FFFDF3E2", warn: "FFA87F00",
};
const FONT = "Calibri";
const fill = (argb) => ({ type: "pattern", pattern: "solid", fgColor: { argb } });
const hair = { style: "thin", color: { argb: C.line } };
const KG = "#,##0";                 // follows the reader's digit grouping (1,33,861 on en-IN)
const GAP = "+#,##0;-#,##0;0";
const PCT = "0%";
const STATUS = {
  over:  { text: "▲ Over",  font: C.over,  bg: C.overBg },
  under: { text: "▼ Under", font: C.under, bg: C.underBg },
  match: { text: "● Match", font: C.match, bg: C.matchBg },
};

function header(ws, cols) {
  ws.columns = cols.map((c) => ({ header: c.h, key: c.k, width: c.w }));
  const r = ws.getRow(1);
  r.height = 22;
  r.eachCell((cell, i) => {
    cell.font = { bold: true, color: { argb: C.white }, size: 10 };
    cell.fill = fill(C.brand);
    cell.alignment = { vertical: "middle", horizontal: i === 1 ? "left" : cols[i - 1].num ? "right" : "left", wrapText: true };
  });
  ws.views = [{ state: "frozen", ySplit: 1, xSplit: 1 }];
}

/* One data row: numbers formatted, gap coloured, status as a coloured tag. */
function dataRow(ws, values, { depth = 0, leaf = false, strong = false, gap = 0, showStatus = leaf, cols }) {
  const row = ws.addRow(values);
  const cls = classify(gap);
  row.outlineLevel = depth;
  // ExcelJS flags the deepest level "collapsed" (level >= outlineLevelRow);
  // nothing here is collapsed, so every row opens and the +/− buttons agree.
  Object.defineProperty(row, "collapsed", { get: () => false, configurable: true });
  cols.forEach((c, i) => {
    const cell = row.getCell(i + 1);
    cell.border = { bottom: hair };
    if (c.fmt) cell.numFmt = c.fmt;
    if (c.k === "name") cell.alignment = { horizontal: "left", indent: depth * 2, vertical: "middle" };
    if (c.k === "month") cell.font = { color: { argb: C.ink3 } };
    if (c.k === "dis" || c.k === "code") cell.font = { color: { argb: C.ink3 } };
    if (c.k === "gap") cell.font = { color: { argb: cls === "over" ? C.over : cls === "under" ? C.under : C.ink2 }, bold: !leaf };
    if (c.k === "status") {
      if (showStatus) {
        const s = STATUS[cls];
        cell.value = s.text;
        cell.font = { color: { argb: s.font }, bold: true };
        cell.fill = fill(s.bg);
        cell.alignment = { horizontal: "center" };
      } else cell.value = null;
    }
  });
  // Rows read like the dashboard: top-level groups shaded and bold, lower
  // groups in secondary ink, customers bold, items plain.
  if (!leaf && depth === 0) {
    row.eachCell({ includeEmpty: true }, (cell) => {
      cell.fill = fill(C.shade0);
      cell.font = { ...(cell.font || {}), bold: true };
    });
  } else if (!leaf) {
    row.getCell(1).font = { color: { argb: C.ink2 } };
  } else if (strong) {
    row.getCell(1).font = { bold: true, color: { argb: C.ink } };
  }
  return row;
}

function totalRow(ws, values, cols, gap) {
  const row = ws.addRow(values);
  row.eachCell({ includeEmpty: true }, (cell, i) => {
    const c = cols[i - 1];
    cell.font = { bold: true, color: { argb: c.k === "gap" ? (gap > 0 ? C.over : gap < 0 ? C.under : C.ink) : C.ink } };
    cell.fill = fill(C.shade1);
    cell.border = { top: { style: "medium", color: { argb: C.line } } };
    if (c.fmt) cell.numFmt = c.fmt;
  });
}

const ratio = (p, s) => (s > 0 ? p / s : null);

/* ── hierarchical sheets ─────────────────────────────────────────── */

function customerSheet(wb, name, custLeaves, mode, sort) {
  // outlineLevelRow = deepest level, so every row opens expanded with the 1–5 buttons above.
  const ws = wb.addWorksheet(name, { properties: { outlineLevelRow: 4, outlineProperties: { summaryBelow: false } } });
  const cols = [
    { h: "Customer Group › Customer › Item Group › Item Parent › Item Name", k: "name", w: 62 },
    { h: "Item Code", k: "code", w: 18 },
    { h: "Month proj kg", k: "month", w: 14, num: true, fmt: KG },
    { h: "MTD proj kg", k: "mtd", w: 14, num: true, fmt: KG },
    { h: "SO kg", k: "so", w: 14, num: true, fmt: KG },
    { h: "Gap kg (MTD − SO)", k: "gap", w: 15, num: true, fmt: GAP },
    { h: "MTD ÷ SO", k: "ratio", w: 10, num: true, fmt: PCT },
    { h: "Status", k: "status", w: 11 },
  ];
  header(ws, cols);
  const t = tree(custLeaves, sort, mode);
  const line = (n, code) => [n.label, code || "", n.projFull, n.proj, n.so, n.gap, ratio(n.proj, n.so), ""];
  for (const g of t.kids) {
    dataRow(ws, line(g), { depth: 0, gap: g.gap, cols });
    for (const c of g.kids) {
      dataRow(ws, line(c), { depth: 1, leaf: true, strong: true, gap: c.gap, cols });
      const sub = tree(c.leaf.items, sort, "all");
      const walk = (n, d) => n.kids.forEach((k) => {
        dataRow(ws, line(k, k.leaf?.code), { depth: d, leaf: !!k.leaf, gap: k.gap, cols });
        if (!k.leaf) walk(k, d + 1);
      });
      walk(sub, 2);
    }
  }
  totalRow(ws, [`Total · ${custLeaves.length} customer${custLeaves.length === 1 ? "" : "s"}`, "", t.projFull, t.proj, t.so, t.gap, ratio(t.proj, t.so), ""], cols, t.gap);
  return ws;
}

function itemSheet(wb, name, itemLeaves, mode, sort) {
  const ws = wb.addWorksheet(name, { properties: { outlineLevelRow: 2, outlineProperties: { summaryBelow: false } } });
  const cols = [
    { h: "Item Group › Item Parent › Item Name", k: "name", w: 62 },
    { h: "Item Code", k: "code", w: 18 },
    { h: "Month proj kg", k: "month", w: 14, num: true, fmt: KG },
    { h: "MTD proj kg", k: "mtd", w: 14, num: true, fmt: KG },
    { h: "SO kg", k: "so", w: 14, num: true, fmt: KG },
    { h: `of which ${CAT_A} kg`, k: "dis", w: 14, num: true, fmt: KG },
    { h: "Gap kg (MTD − SO)", k: "gap", w: 15, num: true, fmt: GAP },
    { h: "MTD ÷ SO", k: "ratio", w: 10, num: true, fmt: PCT },
    { h: "Status", k: "status", w: 11 },
  ];
  header(ws, cols);
  const t = tree(itemLeaves, sort, mode);
  const line = (n) => [n.label, n.leaf?.code || "", n.projFull, n.proj, n.so, n.soDis, n.gap, ratio(n.proj, n.so), ""];
  const walk = (n) => n.kids.forEach((k) => {
    dataRow(ws, line(k), { depth: k.depth, leaf: !!k.leaf, gap: k.gap, cols });
    if (!k.leaf) walk(k);
  });
  walk(t);
  totalRow(ws, [`Total · ${itemLeaves.length} item${itemLeaves.length === 1 ? "" : "s"}`, "", t.projFull, t.proj, t.so, t.soDis, t.gap, ratio(t.proj, t.so), ""], cols, t.gap);
  return ws;
}

/* ── summary ─────────────────────────────────────────────────────── */

function summarySheet(wb, prepared, custLeaves, whLabel) {
  const ws = wb.addWorksheet("Summary", { views: [{ showGridLines: false }] });
  ws.columns = [{ width: 30 }, { width: 18 }, { width: 18 }, { width: 50 }];
  const pr = prepared.prorate;
  const period = prepared.projPeriods.map(ymLabel).join(", ") || "—";

  const title = ws.addRow(["Projection vs SO"]);
  title.font = { name: FONT, bold: true, size: 16, color: { argb: C.ink } };
  title.height = 26;
  const sub = ws.addRow([`${period} · ${whLabel} · exported ${new Date().toLocaleString("en-IN")}`]);
  sub.font = { name: FONT, color: { argb: C.ink2 } };
  ws.addRow([]);

  // Period in short, as on the dashboard header.
  const days = ws.addRow(["SO days / days in month", pr.on ? `${pr.soDays} / ${pr.monthDays}` : `— / ${pr.monthDays || "—"}`, pr.on ? pr.factor : null]);
  days.getCell(1).font = { bold: true };
  days.getCell(2).font = { bold: true, size: 13, color: { argb: C.over } };
  days.getCell(3).numFmt = PCT;
  days.getCell(3).font = { color: { argb: C.ink2 } };
  days.getCell(4).value = pr.on ? "MTD projection = Month projection ÷ days in month × SO days" : pr.why;
  days.getCell(4).font = { color: { argb: C.ink2 }, italic: true };
  ws.addRow([]);

  const sum = (f, a = custLeaves) => a.reduce((s, l) => s + l[f], 0);
  const over = custLeaves.filter((l) => classify(l.gap) === "over");
  const under = custLeaves.filter((l) => classify(l.gap) === "under");
  const P = sum("proj"), S = sum("so");
  const tiles = [
    ["Month projection", sum("projFull"), "kg · Total KGs", null],
    ["MTD projection", P, pr.on ? `${pr.soDays} of ${pr.monthDays} days` : "= month", null],
    ["SO", S, `so tab ${Math.round(sum("soTab")).toLocaleString("en-IN")} · ${CAT_A} dispatches ${Math.round(sum("soDis")).toLocaleString("en-IN")}`, null],
    ["Net gap (MTD − SO)", P - S, `MTD ÷ SO ${S > 0 ? Math.round((P / S) * 100) + "%" : "—"}`, P - S],
    ["Over-projected", sum("gap", over), `${over.length} customer${over.length === 1 ? "" : "s"} with MTD proj > SO`, 1],
    ["Under-projected", sum("gap", under), `${under.length} customer${under.length === 1 ? "" : "s"} with MTD proj < SO`, -1],
  ];
  const th = ws.addRow(["Measure", "kg", "", "Note"]);
  th.eachCell((c) => { c.font = { bold: true, color: { argb: C.white } }; c.fill = fill(C.brand); });
  for (const [k, v, note, sign] of tiles) {
    const r = ws.addRow([k, v, "", note]);
    r.getCell(1).font = { bold: true };
    r.getCell(2).numFmt = sign == null ? KG : GAP;
    r.getCell(2).font = { bold: true, size: 12, color: { argb: sign == null ? C.ink : sign > 0 ? C.over : sign < 0 ? C.under : C.ink } };
    r.getCell(4).font = { color: { argb: C.ink2 } };
    r.eachCell({ includeEmpty: true }, (c) => { c.border = { bottom: hair }; });
  }
  ws.addRow([]);

  const rules = [
    ["How it is built", ""],
    ["Month projection", "projection › Total KGs"],
    ["MTD projection", "Month projection ÷ days in month × unique Sales Order Dates in the so tab"],
    ["SO", `so › Stock Qty in KGS (Customer Group ≠ ${CAT_A}) + dispatches › Stock Qty In Kg (Sales Channel = ${CAT_A}, Invoice Status ≠ Return)`],
    ["Over / Under", "Status column — per customer and per item: MTD proj > SO is ▲ Over (blue), < SO is ▼ Under (red). Filter it from the header."],
    ["Month filter", prepared.filterMonth ? period : "off — tabs compared as loaded"],
    ["Outline", "Use the 1–5 buttons at the top left of the Customer and Item tabs to collapse or open levels"],
  ];
  rules.forEach(([a, b], i) => {
    const r = ws.addRow([a, b]);
    if (!i) { r.font = { name: FONT, bold: true, size: 12 }; return; }
    r.getCell(1).font = { bold: true, color: { argb: C.ink2 } };
    r.getCell(1).alignment = { vertical: "top" };
    ws.mergeCells(r.number, 2, r.number, 4);
    r.getCell(2).alignment = { wrapText: true, vertical: "top" };
    if (b.length > 80) r.height = 30;   // merged cells do not grow on their own
  });
  ws.addRow([]);

  const xh = ws.addRow(["Rows left out", "Rows", "Kg", "Reason"]);
  xh.eachCell((c) => { c.font = { bold: true, color: { argb: C.white } }; c.fill = fill(C.brand); });
  xh.getCell(4).alignment = { indent: 1 };
  for (const e of prepared.exclusions) {
    const r = ws.addRow([TAB_LABEL[e.tab], e.rows, e.kg, e.reason]);
    r.getCell(2).numFmt = KG;
    r.getCell(3).numFmt = KG;
    r.getCell(4).alignment = { indent: 1 };
    if (e.kind === "month") r.eachCell((c) => { c.fill = fill(C.warnBg); });
    r.eachCell({ includeEmpty: true }, (c) => { c.border = { bottom: hair }; });
  }
}

/* ── the workbook ────────────────────────────────────────────────── */

export async function exportWorkbook({ prepared, wh, sort = "gap" }) {
  const ExcelJS = await loadExcelJS();
  const wb = new ExcelJS.Workbook();
  wb.creator = "Projection vs SO";
  wb.created = new Date();

  const cust = leaves(prepared, "customer", wh);
  const items = leaves(prepared, "item", wh);
  const whLabel = wh ? `CFA (${wh})` : "All CFAs";

  summarySheet(wb, prepared, cust, whLabel);
  customerSheet(wb, "Customer", cust, "all", sort);
  itemSheet(wb, "Item", items, "all", sort);

  // Without a workbook window (<bookViews>) Excel can open the file with no
  // visible sheet at all — a blank grey window, though the data is there.
  wb.views = [{ x: 0, y: 0, width: 28800, height: 16000, firstSheet: 0, activeTab: 1, visibility: "visible" }];

  // One typeface everywhere: a font set for colour or weight alone would
  // otherwise drop back to the application default. Print landscape, one
  // page wide.
  wb.eachSheet((ws) => {
    ws.eachRow({ includeEmpty: true }, (row) => row.eachCell({ includeEmpty: true }, (cell) => {
      cell.font = { name: FONT, size: 10, ...(cell.font || {}) };
    }));
    ws.pageSetup = { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0, margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 } };
    if (ws.name === "Summary") ws.properties.tabColor = { argb: C.brand };
    else {
      ws.pageSetup.printTitlesRow = "1:1";
      // Filter buttons on the header, e.g. Status = ▼ Under.
      // The range spans the data (not the Total row), so the dropdowns filter it.
      ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: ws.rowCount - 1, column: ws.columnCount } };
    }
  });

  const buf = await wb.xlsx.writeBuffer();
  const tag = [prepared.projPeriods[0] || "", wh || "all"].filter(Boolean).join("-");
  const a = document.createElement("a");
  a.href = URL.createObjectURL(new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
  a.download = `proj-vs-so-${tag}.xlsx`;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
}
