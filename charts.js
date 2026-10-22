/* Projection vs SO — charts. Inline SVG, no library.
 *
 *  groupBars  MTD projection vs SO per Customer Group / Item Group (paired bars)
 *  gapBars    the largest gaps, diverging from zero: over right, under left
 *
 * Colours come from CSS roles (--viz-mtd, --viz-so, --viz-over, --viz-under),
 * validated as pairs in light and dark. Text never wears a series colour.
 * Labels are data: everything user-supplied goes in through textContent.
 */

const NS = "http://www.w3.org/2000/svg";
const nf0 = new Intl.NumberFormat("en-IN", { maximumFractionDigits: 0 });

const el = (tag, attrs = {}, text) => {
  const n = document.createElementNS(NS, tag);
  for (const [k, v] of Object.entries(attrs)) if (v != null) n.setAttribute(k, v);
  if (text != null) n.textContent = text;
  return n;
};

/* Compact axis figures: 0 · 950 · 12K · 1.2L · 2.4Cr — Indian units. */
export function compact(v) {
  const a = Math.abs(v), s = v < 0 ? "−" : "";
  if (a >= 1e7) return s + trim(a / 1e7) + "Cr";
  if (a >= 1e5) return s + trim(a / 1e5) + "L";
  if (a >= 1e3) return s + trim(a / 1e3) + "K";
  return s + nf0.format(a);
}
const trim = (x) => (x >= 100 ? x.toFixed(0) : x >= 10 ? x.toFixed(1).replace(/\.0$/, "") : x.toFixed(2).replace(/0$/, "").replace(/\.0$/, ""));

/* Clean ticks: 0..max in 4–5 steps of 1/2/2.5/5 × 10^n. */
function ticks(max) {
  if (max <= 0) return [0];
  const raw = max / 4, p = 10 ** Math.floor(Math.log10(raw));
  const step = [1, 2, 2.5, 5, 10].map((m) => m * p).find((s) => s >= raw);
  const out = [];
  for (let v = 0; v <= max + step * 0.001; v += step) out.push(v);
  if (out[out.length - 1] < max) out.push(out[out.length - 1] + step);
  return out;
}

/* Text fitted to a width by measuring, never by clipping. */
function fit(textEl, maxW) {
  const full = textEl.textContent;
  if (textEl.getComputedTextLength() <= maxW) return;
  let lo = 0, hi = full.length;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    textEl.textContent = full.slice(0, mid) + "…";
    if (textEl.getComputedTextLength() <= maxW) lo = mid; else hi = mid - 1;
  }
  textEl.textContent = full.slice(0, lo) + "…";
}

/* A bar with a 4px rounded data-end and a square baseline end. */
function barPath(x0, x1, y, h) {
  const r = Math.min(4, Math.abs(x1 - x0), h / 2);
  if (Math.abs(x1 - x0) < 0.5) return "";
  if (x1 >= x0) {
    return `M${x0},${y}H${x1 - r}Q${x1},${y} ${x1},${y + r}V${y + h - r}Q${x1},${y + h} ${x1 - r},${y + h}H${x0}Z`;
  }
  return `M${x0},${y}H${x1 + r}Q${x1},${y} ${x1},${y + r}V${y + h - r}Q${x1},${y + h} ${x1 + r},${y + h}H${x0}Z`;
}

/* ── tooltip: one per page, values lead, labels follow ─────────────── */

let tip;
function tooltip() {
  if (tip) return tip;
  tip = document.createElement("div");
  tip.className = "viz-tip";
  tip.setAttribute("role", "status");
  document.body.appendChild(tip);
  return tip;
}
function showTip(evt, title, rows) {
  const t = tooltip();
  t.replaceChildren();
  const h = document.createElement("div");
  h.className = "viz-tip-t";
  h.textContent = title;
  t.appendChild(h);
  for (const r of rows) {
    const row = document.createElement("div");
    row.className = "viz-tip-r";
    if (r.color) {
      const k = document.createElement("i");
      k.style.background = r.color;
      row.appendChild(k);
    }
    const v = document.createElement("b");
    v.textContent = r.value;
    const l = document.createElement("span");
    l.textContent = r.label;
    row.append(v, l);
    t.appendChild(row);
  }
  t.classList.add("on");
  const rect = evt.currentTarget.getBoundingClientRect?.() || { left: evt.clientX, top: evt.clientY, width: 0, height: 0 };
  const x = evt.clientX ?? rect.left + rect.width / 2;
  const y = evt.clientY ?? rect.top;
  const w = t.offsetWidth, hgt = t.offsetHeight;
  t.style.left = Math.max(8, Math.min(window.innerWidth - w - 8, x + 14)) + "px";
  t.style.top = Math.max(8, y - hgt - 12) + "px";
}
const hideTip = () => tip?.classList.remove("on");

function bindTip(target, title, rows) {
  target.setAttribute("tabindex", "0");
  target.addEventListener("pointermove", (e) => showTip(e, title, rows));
  target.addEventListener("pointerleave", hideTip);
  target.addEventListener("focus", (e) => {
    const r = e.currentTarget.getBoundingClientRect();
    showTip({ currentTarget: e.currentTarget, clientX: r.right, clientY: r.top }, title, rows);
  });
  target.addEventListener("blur", hideTip);
}

const css = (name) => getComputedStyle(document.documentElement).getPropertyValue(name).trim();

/* ── paired bars: MTD projection vs SO per group ─────────────────── */

export function groupBars(host, groups, { animate = true, kg, signed } = {}) {
  host.replaceChildren();
  if (!groups.length) return;
  const W = Math.max(300, host.clientWidth);
  const narrow = W < 520;
  const labelW = narrow ? 110 : 168, gapW = narrow ? 64 : 84, padR = 8;
  const barH = 9, pairGap = 2, rowH = 34, top = 6, axisH = 22;
  const plotW = W - labelW - gapW - padR;
  const H = top + groups.length * rowH + axisH;
  const max = Math.max(...groups.map((g) => Math.max(g.proj, g.so)), 1);
  const tk = ticks(max), xmax = tk[tk.length - 1];
  const X = (v) => labelW + (v / xmax) * plotW;

  const svg = el("svg", { viewBox: `0 0 ${W} ${H}`, width: W, height: H, role: "img", class: animate ? "viz anim" : "viz" });
  svg.appendChild(el("title", {}, "MTD projection and SO by group, kg"));

  const grid = el("g", { class: "viz-grid" });
  for (const v of tk) {
    const x = X(v);
    grid.appendChild(el("line", { x1: x, x2: x, y1: top, y2: H - axisH + 2 }));
    grid.appendChild(el("text", { x, y: H - 6, "text-anchor": "middle", class: "viz-tick" }, compact(v)));
  }
  svg.appendChild(grid);
  svg.appendChild(el("line", { x1: X(0), x2: X(0), y1: top, y2: H - axisH + 2, class: "viz-base" }));

  const cMtd = css("--viz-mtd"), cSo = css("--viz-so");
  host.appendChild(svg);   // attached first so labels can be measured
  groups.forEach((g, i) => {
    const y = top + i * rowH + (rowH - (barH * 2 + pairGap)) / 2;
    const row = el("g", { class: "viz-row", style: `--i:${i}` });
    const lab = el("text", { x: labelW - 10, y: y + barH + 1, "text-anchor": "end", class: "viz-label", "dominant-baseline": "middle" }, g.label);
    row.appendChild(lab);
    row.appendChild(el("path", { d: barPath(X(0), X(g.proj), y, barH), class: "viz-bar", fill: "var(--viz-mtd)" }));
    row.appendChild(el("path", { d: barPath(X(0), X(g.so), y + barH + pairGap, barH), class: "viz-bar", fill: "var(--viz-so)" }));
    const gp = g.proj - g.so;
    row.appendChild(el("text", { x: W - padR, y: y + barH + 1, "text-anchor": "end", class: "viz-val", "dominant-baseline": "middle" }, signed(gp)));
    const hit = el("rect", { x: 0, y: top + i * rowH, width: W, height: rowH, class: "viz-hit" });
    row.appendChild(hit);
    svg.appendChild(row);
    fit(lab, labelW - 16);
    bindTip(hit, g.label, [
      { color: cMtd, value: kg(g.proj) + " kg", label: "MTD projection" },
      { color: cSo, value: kg(g.so) + " kg", label: "SO" },
      { value: signed(gp) + " kg", label: gp > 0 ? "over-projected" : gp < 0 ? "under-projected" : "match" },
    ]);
  });
}

/* ── diverging bars: the largest gaps ────────────────────────────── */

export function gapBars(host, items, { animate = true, kg, signed } = {}) {
  host.replaceChildren();
  if (!items.length) return;
  const W = Math.max(300, host.clientWidth);
  const narrow = W < 520;
  const labelW = narrow ? 120 : 210, padR = 8, valPad = 6;
  const barH = 14, rowH = 26, top = 4, axisH = 22;
  const H = top + items.length * rowH + axisH;
  const maxPos = Math.max(0, ...items.map((d) => d.gap));
  const maxNeg = Math.max(0, ...items.map((d) => -d.gap));
  // Room either side of zero in proportion to the data, plus space for the tip label.
  const lab = 58;
  const plotW = W - labelW - padR;
  const span = maxPos + maxNeg || 1;
  const usable = plotW - (maxPos ? lab : 0) - (maxNeg ? lab : 0);
  const zero = labelW + (maxNeg ? lab : 0) + (maxNeg / span) * usable;
  const X = (v) => zero + (v / span) * usable;

  const svg = el("svg", { viewBox: `0 0 ${W} ${H}`, width: W, height: H, role: "img", class: animate ? "viz anim" : "viz" });
  svg.appendChild(el("title", {}, "Largest gaps, MTD projection minus SO, kg"));
  svg.appendChild(el("line", { x1: zero, x2: zero, y1: top - 2, y2: H - axisH + 2, class: "viz-base" }));
  svg.appendChild(el("text", { x: zero, y: H - 6, "text-anchor": "middle", class: "viz-tick" }, "0"));
  if (maxNeg) svg.appendChild(el("text", { x: labelW + 2, y: H - 6, "text-anchor": "start", class: "viz-tick" }, "◀ under"));
  if (maxPos) svg.appendChild(el("text", { x: W - padR, y: H - 6, "text-anchor": "end", class: "viz-tick" }, "over ▶"));

  host.appendChild(svg);
  items.forEach((d, i) => {
    const y = top + i * rowH + (rowH - barH) / 2;
    const over = d.gap > 0;
    const row = el("g", { class: "viz-row " + (over ? "pos" : "neg"), style: `--i:${i}` });
    const name = el("text", { x: labelW - 10, y: y + barH / 2 + 1, "text-anchor": "end", class: "viz-label", "dominant-baseline": "middle" }, d.label);
    row.appendChild(name);
    row.appendChild(el("path", { d: barPath(zero, X(d.gap), y, barH), class: "viz-bar", fill: over ? "var(--viz-over)" : "var(--viz-under)" }));
    row.appendChild(el("text", {
      x: X(d.gap) + (over ? valPad : -valPad), y: y + barH / 2 + 1,
      "text-anchor": over ? "start" : "end", class: "viz-val", "dominant-baseline": "middle",
    }, signed(d.gap)));
    const hit = el("rect", { x: 0, y: top + i * rowH, width: W, height: rowH, class: "viz-hit" });
    row.appendChild(hit);
    svg.appendChild(row);
    fit(name, labelW - 16);
    bindTip(hit, d.full || d.label, [
      { color: css("--viz-mtd"), value: kg(d.proj) + " kg", label: "MTD projection" },
      { color: css("--viz-so"), value: kg(d.so) + " kg", label: "SO" },
      { color: over ? css("--viz-over") : css("--viz-under"), value: signed(d.gap) + " kg", label: over ? "over-projected" : "under-projected" },
    ]);
  });
}

window.addEventListener("scroll", hideTip, { passive: true });
