// Charts: data extraction from ranges, canvas rendering, and floating chart objects over the grid.
import { parseRangeName, rangeName, uid, clamp } from './util.js';
import { formatValue, isDateFormat } from './format.js';

export const CHART_TYPES = [
  { id: 'column', label: 'Clustered column' }, { id: 'stackedColumn', label: 'Stacked column' }, { id: 'bar', label: 'Clustered bar' }, { id: 'stackedBar', label: 'Stacked bar' },
  { id: 'line', label: 'Line' }, { id: 'lineMarkers', label: 'Line with markers' }, { id: 'area', label: 'Area' }, { id: 'stackedArea', label: 'Stacked area' },
  { id: 'pie', label: 'Pie' }, { id: 'doughnut', label: 'Doughnut' }, { id: 'scatter', label: 'Scatter (XY)' },
];
// Validated categorical palette (fixed order, never cycled past 8: later series fold into grey)
export const PALETTE = ['#2a78d6', '#eb6834', '#1baf7a', '#eda100', '#e87ba4', '#008300', '#6250d6', '#e34948'];
const OTHER = '#9a9a9a';
const INK = '#333', INK2 = '#666', GRID = '#e6e6e6', AXIS = '#b0b0b0';
const FONT = '12px "Segoe UI", system-ui, sans-serif';

export function seriesColor(i) { return i < PALETTE.length ? PALETTE[i] : OTHER; }

// Extract series/categories from the chart's range.
export function chartData(wb, chart) {
  const rg = parseRangeName(chart.range || '');
  if (!rg) return { series: [], categories: [], error: 'Invalid range' };
  const sheet = rg.sheet ? wb.getSheet(rg.sheet) : wb.sheets[0];
  if (!sheet) return { series: [], categories: [], error: 'Sheet not found' };
  const r2 = Math.min(rg.r2, rg.r1 + 5000), c2 = Math.min(rg.c2, rg.c1 + 200);
  const rows = [];
  for (let r = rg.r1; r <= r2; r++) { const row = []; for (let c = rg.c1; c <= c2; c++) row.push(sheet.getValue(r, c)); rows.push(row); }
  const nr = rows.length, nc = rows[0] ? rows[0].length : 0;
  if (!nr || !nc) return { series: [], categories: [] };
  const isNum = (v) => typeof v === 'number';
  const inRows = chart.seriesInRows != null ? chart.seriesInRows : nc > nr; // more columns than rows -> series laid out in rows
  const m = inRows ? transpose(rows) : rows; // now each column of m is a series candidate, rows are categories
  const H = m.length, W = m[0].length;
  const firstRowHeaders = W > 0 && m[0].some((v) => typeof v === 'string') && !(m[0].every(isNum));
  const firstColCats = H > (firstRowHeaders ? 1 : 0) && m.slice(firstRowHeaders ? 1 : 0).some((row) => typeof row[0] === 'string' || (typeof row[0] === 'number' && !isNum(row[1]) && W === 1 ? false : false)) ;
  let catCol = false;
  if (W > 1) { const body = m.slice(firstRowHeaders ? 1 : 0); const firstCol = body.map((r) => r[0]); if (firstCol.some((v) => typeof v === 'string')) catCol = true; else if (chart.type === 'scatter') catCol = true; else { const fmtCell = sheet.getCell(rg.r1 + (firstRowHeaders ? 1 : 0), rg.c1); const code = sheet.effectiveNumFmt(inRows ? rg.r1 : rg.r1 + (firstRowHeaders ? 1 : 0), inRows ? rg.c1 + (firstRowHeaders ? 1 : 0) : rg.c1); if (isDateFormat(code)) catCol = true; } }
  const start = firstRowHeaders ? 1 : 0;
  const body = m.slice(start);
  const categories = catCol ? body.map((row, i) => { const v = row[0]; if (v == null) return ''; if (typeof v === 'number') { const rr = inRows ? rg.r1 : rg.r1 + start + i, cc = inRows ? rg.c1 + start + i : rg.c1; return formatValue(v, sheet.effectiveNumFmt(rr, cc)).text; } return String(v); }) : body.map((_, i) => String(i + 1));
  let xValues = catCol ? body.map((row) => (typeof row[0] === 'number' ? row[0] : null)) : body.map((_, i) => i + 1);
  if (xValues.every((v) => v == null)) xValues = body.map((_, i) => i + 1); // text categories on a scatter chart: use 1..n
  const series = [];
  for (let j = catCol ? 1 : 0; j < W; j++) {
    const name = firstRowHeaders && m[0][j] != null ? String(m[0][j]) : `Series ${series.length + 1}`;
    const values = body.map((row) => (typeof row[j] === 'number' ? row[j] : null));
    const fr = inRows ? rg.r1 + j : rg.r1 + start, fc = inRows ? rg.c1 + start : rg.c1 + j;
    series.push({ name, values, color: seriesColor(series.length), numFmt: sheet.effectiveNumFmt(fr, fc) });
  }
  return { series, categories, xValues, sheet };
}
function transpose(m) { return m[0].map((_, j) => m.map((row) => row[j])); }

function niceTicks(min, max, count = 5) {
  if (min === max) { if (min === 0) { max = 1; } else { min = min > 0 ? 0 : min * 1.2; max = max > 0 ? max * 1.2 : 0; } }
  const range = max - min;
  const rough = range / count;
  const mag = Math.pow(10, Math.floor(Math.log10(rough)));
  const norm = rough / mag;
  const step = (norm <= 1 ? 1 : norm <= 2 ? 2 : norm <= 2.5 ? 2.5 : norm <= 5 ? 5 : 10) * mag;
  const lo = Math.floor(min / step) * step, hi = Math.ceil(max / step) * step;
  const ticks = []; for (let v = lo; v <= hi + step / 2; v += step) ticks.push(Number(v.toFixed(10)));
  return { lo, hi: ticks[ticks.length - 1], ticks, step };
}
function fmtNum(v, code) {
  if (code && code !== 'General') { const t = formatValue(v, code).text; if (t.length <= 14) return t; }
  if (Math.abs(v) >= 1e6) return (v / 1e6).toFixed(Math.abs(v) >= 1e7 ? 0 : 1) + 'M';
  if (Math.abs(v) >= 1e4) return (v / 1e3).toFixed(0) + 'k';
  return Number.isInteger(v) ? String(v) : String(Number(v.toPrecision(6)));
}
const rrect = (ctx, x, y, w, h, r) => { r = Math.min(r, Math.abs(w) / 2, Math.abs(h) / 2); ctx.beginPath(); ctx.moveTo(x + r, y); ctx.lineTo(x + w - r, y); ctx.quadraticCurveTo(x + w, y, x + w, y + r); ctx.lineTo(x + w, y + h - r); ctx.quadraticCurveTo(x + w, y + h, x + w - r, y + h); ctx.lineTo(x + r, y + h); ctx.quadraticCurveTo(x, y + h, x, y + h - r); ctx.lineTo(x, y + r); ctx.quadraticCurveTo(x, y, x + r, y); ctx.closePath(); };

// Render a chart into a canvas (css size w × h). Returns hit regions [{x,y,w,h,label}] for tooltips.
export function renderChart(canvas, chart, data, w, h) {
  const dpr = window.devicePixelRatio || 1;
  if (canvas.width !== Math.round(w * dpr) || canvas.height !== Math.round(h * dpr)) { canvas.width = Math.round(w * dpr); canvas.height = Math.round(h * dpr); canvas.style.width = w + 'px'; canvas.style.height = h + 'px'; }
  const ctx = canvas.getContext('2d');
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, w, h);
  ctx.font = FONT; ctx.textBaseline = 'middle';
  const hits = [];
  const pad = 10;
  let top = pad;
  const title = chart.title != null ? chart.title : (data.series.length === 1 ? data.series[0].name : '');
  if (title) { ctx.fillStyle = INK; ctx.font = '600 14px "Segoe UI", system-ui, sans-serif'; ctx.textAlign = 'center'; ctx.fillText(title, w / 2, top + 9, w - 2 * pad); top += 26; ctx.font = FONT; }
  if (data.error) { ctx.fillStyle = INK2; ctx.textAlign = 'center'; ctx.fillText(data.error, w / 2, h / 2); return hits; }
  const series = data.series;
  if (!series.length) { ctx.fillStyle = INK2; ctx.textAlign = 'center'; ctx.fillText('No numeric data in range', w / 2, h / 2); return hits; }
  const type = chart.type || 'column';
  const showLegend = chart.legend !== false && (series.length > 1 || type === 'pie' || type === 'doughnut');
  let bottom = h - pad;
  // legend (bottom)
  if (showLegend) {
    const items = (type === 'pie' || type === 'doughnut') ? data.categories.map((c, i) => ({ name: c, color: seriesColor(i) })) : series.map((s) => ({ name: s.name, color: s.color }));
    const widths = items.map((it) => ctx.measureText(it.name).width + 22);
    const lines = []; let cur = [], cw = 0;
    for (let i = 0; i < items.length; i++) { if (cw + widths[i] > w - 2 * pad && cur.length) { lines.push(cur); cur = []; cw = 0; } cur.push(i); cw += widths[i]; }
    if (cur.length) lines.push(cur);
    const maxLines = Math.min(lines.length, 3);
    bottom -= maxLines * 18;
    for (let li = 0; li < maxLines; li++) {
      const lw = lines[li].reduce((a, i) => a + widths[i], 0);
      let x = (w - lw) / 2; const y = bottom + 9 + li * 18;
      for (const i of lines[li]) { ctx.fillStyle = items[i].color; rrect(ctx, x, y - 5, 10, 10, 2); ctx.fill(); ctx.fillStyle = INK2; ctx.textAlign = 'left'; ctx.fillText(items[i].name, x + 14, y, widths[i] - 22); x += widths[i]; }
    }
    if (lines.length > maxLines) { ctx.fillStyle = INK2; ctx.textAlign = 'center'; ctx.fillText(`+${items.length - lines.slice(0, maxLines).flat().length} more`, w / 2, bottom + maxLines * 18 - 2); }
    bottom -= 6;
  }
  if (type === 'pie' || type === 'doughnut') return renderPie(ctx, chart, data, pad, top, w - pad, bottom, hits, type === 'doughnut');
  // ---- axis charts ----
  const horizontal = type === 'bar' || type === 'stackedBar';
  const stacked = type.startsWith('stacked');
  const isLine = type === 'line' || type === 'lineMarkers'; const isArea = type === 'area' || type === 'stackedArea'; const isScatter = type === 'scatter';
  const n = data.categories.length;
  // value range
  let vmin = Infinity, vmax = -Infinity;
  if (stacked) { for (let i = 0; i < n; i++) { let pos = 0, neg = 0; for (const s of series) { const v = s.values[i]; if (v == null) continue; if (v >= 0) pos += v; else neg += v; } vmin = Math.min(vmin, neg, 0); vmax = Math.max(vmax, pos, 0); } }
  else { for (const s of series) for (const v of s.values) if (v != null) { vmin = Math.min(vmin, v); vmax = Math.max(vmax, v); } if (!isScatter && !isLine) { vmin = Math.min(vmin, 0); vmax = Math.max(vmax, 0); } else { vmin = Math.min(vmin, 0 > vmin ? vmin : 0); vmax = Math.max(vmax, 0); } }
  if (!isFinite(vmin)) { vmin = 0; vmax = 1; }
  const vt = niceTicks(vmin, vmax, horizontal ? 6 : 5);
  // x range for scatter
  let xt = null;
  if (isScatter) { const xs = data.xValues.filter((v) => v != null); const xmin = xs.length ? Math.min(...xs) : 0, xmax = xs.length ? Math.max(...xs) : 1; xt = niceTicks(Math.min(xmin, 0 < xmin ? xmin : 0), xmax, 6); }
  const valueFmt = series[0].numFmt;
  // measure label widths
  const vLabels = vt.ticks.map((v) => fmtNum(v, valueFmt));
  const vLabelW = Math.max(...vLabels.map((s) => ctx.measureText(s).width));
  const catLabelW = Math.max(0, ...data.categories.map((s) => ctx.measureText(s).width));
  let left = pad + (horizontal ? Math.min(catLabelW, w * 0.3) : vLabelW) + 8;
  let plotBottom = bottom - 22;
  const right = w - pad - 6;
  if (!horizontal && !isScatter) { const slotW = (right - left) / Math.max(1, n); if (catLabelW > slotW - 4 && n > 1) plotBottom = bottom - Math.min(70, catLabelW * 0.7 + 14); }
  const plotTop = top + 6;
  const pw = right - left, ph = plotBottom - plotTop;
  if (pw < 20 || ph < 20) return hits;
  const vScale = (v) => horizontal ? left + (v - vt.lo) / (vt.hi - vt.lo) * pw : plotBottom - (v - vt.lo) / (vt.hi - vt.lo) * ph;
  // gridlines + value axis labels
  ctx.strokeStyle = GRID; ctx.lineWidth = 1; ctx.fillStyle = INK2; ctx.font = FONT;
  for (let i = 0; i < vt.ticks.length; i++) {
    const v = vt.ticks[i]; const p = Math.round(vScale(v)) + 0.5;
    ctx.beginPath(); if (horizontal) { ctx.moveTo(p, plotTop); ctx.lineTo(p, plotBottom); } else { ctx.moveTo(left, p); ctx.lineTo(right, p); } ctx.stroke();
    if (horizontal) { ctx.textAlign = 'center'; ctx.fillText(vLabels[i], p, plotBottom + 11); } else { ctx.textAlign = 'right'; ctx.fillText(vLabels[i], left - 6, p); }
  }
  // zero / base axis
  ctx.strokeStyle = AXIS; ctx.beginPath();
  const zero = Math.round(vScale(clamp(0, vt.lo, vt.hi))) + 0.5;
  if (horizontal) { ctx.moveTo(zero, plotTop); ctx.lineTo(zero, plotBottom); } else { ctx.moveTo(left, zero); ctx.lineTo(right, zero); }
  ctx.stroke();
  // category positions
  const slot = (horizontal ? ph : pw) / Math.max(1, isScatter ? 1 : n);
  const catPos = (i) => horizontal ? plotTop + slot * i : left + slot * i;
  // category labels
  if (!isScatter) {
    ctx.fillStyle = INK2;
    const step = Math.max(1, Math.ceil((horizontal ? 14 : (catLabelW > slot - 4 ? 14 : catLabelW + 8)) / slot));
    for (let i = 0; i < n; i += step) {
      const label = data.categories[i];
      if (horizontal) { ctx.textAlign = 'right'; ctx.fillText(label, left - 6, catPos(i) + slot / 2, left - pad - 4); }
      else if (catLabelW > slot - 4 && n > 1) { ctx.save(); ctx.translate(catPos(i) + slot / 2, plotBottom + 6); ctx.rotate(-Math.PI / 4); ctx.textAlign = 'right'; ctx.fillText(label, 0, 0, Math.min(90, catLabelW)); ctx.restore(); }
      else { ctx.textAlign = 'center'; ctx.fillText(label, catPos(i) + slot / 2, plotBottom + 11, slot - 2); }
    }
  } else {
    ctx.fillStyle = INK2; ctx.textAlign = 'center';
    for (const v of xt.ticks) { const px = left + (v - xt.lo) / (xt.hi - xt.lo) * pw; ctx.fillText(fmtNum(v), px, plotBottom + 11); ctx.strokeStyle = GRID; ctx.beginPath(); ctx.moveTo(Math.round(px) + 0.5, plotTop); ctx.lineTo(Math.round(px) + 0.5, plotBottom); ctx.stroke(); }
  }
  const label = (s, i, v) => `${data.categories[i] ? data.categories[i] + ' · ' : ''}${s.name}: ${formatValue(v, s.numFmt).text}`;
  if (isScatter) {
    for (const s of series) {
      ctx.fillStyle = s.color;
      for (let i = 0; i < n; i++) { const xv = data.xValues[i], yv = s.values[i]; if (xv == null || yv == null) continue; const px = left + (xv - xt.lo) / (xt.hi - xt.lo) * pw, py = vScale(yv); ctx.beginPath(); ctx.arc(px, py, 4, 0, Math.PI * 2); ctx.fill(); ctx.strokeStyle = '#fff'; ctx.lineWidth = 1.5; ctx.stroke(); hits.push({ x: px - 6, y: py - 6, w: 12, h: 12, label: `${s.name}: (${fmtNum(xv)}, ${formatValue(yv, s.numFmt).text})` }); }
    }
    return hits;
  }
  if (isLine || isArea) {
    const stackBase = new Array(n).fill(0);
    for (let si = 0; si < series.length; si++) {
      const s = series[si];
      const pts = [];
      for (let i = 0; i < n; i++) { const v = s.values[i]; if (v == null) { pts.push(null); continue; } const base = stacked ? stackBase[i] : 0; const yv = base + v; if (stacked) stackBase[i] = yv; pts.push({ x: catPos(i) + slot / 2, y: vScale(yv), y0: vScale(base), v }); }
      if (isArea) {
        ctx.fillStyle = s.color + '66'; ctx.beginPath(); let started = false;
        for (let i = 0; i < n; i++) { const p = pts[i]; if (!p) continue; if (!started) { ctx.moveTo(p.x, p.y0); started = true; } ctx.lineTo(p.x, p.y); }
        for (let i = n - 1; i >= 0; i--) { const p = pts[i]; if (!p) continue; ctx.lineTo(p.x, p.y0); }
        ctx.closePath(); ctx.fill();
      }
      ctx.strokeStyle = s.color; ctx.lineWidth = 2; ctx.lineJoin = 'round'; ctx.beginPath(); let pen = false;
      for (const p of pts) { if (!p) { pen = false; continue; } if (!pen) { ctx.moveTo(p.x, p.y); pen = true; } else ctx.lineTo(p.x, p.y); }
      ctx.stroke();
      const markers = type === 'lineMarkers' || n === 1 || (slot > 24 && !isArea);
      for (let i = 0; i < n; i++) { const p = pts[i]; if (!p) continue; if (markers) { ctx.fillStyle = s.color; ctx.beginPath(); ctx.arc(p.x, p.y, 4, 0, Math.PI * 2); ctx.fill(); ctx.strokeStyle = '#fff'; ctx.lineWidth = 1.5; ctx.stroke(); } hits.push({ x: p.x - Math.max(6, slot / 2), y: p.y - 8, w: Math.max(12, slot), h: 16, label: label(s, i, p.v) }); }
    }
    return hits;
  }
  // bars / columns
  const groupGap = slot * 0.2;
  const groupSize = slot - groupGap;
  const barCount = stacked ? 1 : series.length;
  const barSize = Math.max(1, (groupSize - (barCount - 1) * 2) / barCount);
  const posBase = new Array(n).fill(0), negBase = new Array(n).fill(0);
  for (let si = 0; si < series.length; si++) {
    const s = series[si];
    ctx.fillStyle = s.color;
    for (let i = 0; i < n; i++) {
      const v = s.values[i]; if (v == null) continue;
      let base = 0;
      if (stacked) { if (v >= 0) { base = posBase[i]; posBase[i] += v; } else { base = negBase[i]; negBase[i] += v; } }
      const a = vScale(base), b = vScale(base + v);
      const off = stacked ? groupGap / 2 : groupGap / 2 + si * (barSize + 2);
      let x, y, bw, bh;
      if (horizontal) { y = catPos(i) + off; bh = barSize; x = Math.min(a, b); bw = Math.abs(b - a); }
      else { x = catPos(i) + off; bw = barSize; y = Math.min(a, b); bh = Math.abs(b - a); }
      if (stacked && bh > 2 && bw > 2) { if (horizontal) { x += 1; bw -= 2; } else { y += 1; bh -= 2; } }
      // round only the data end (away from the baseline)
      const r = Math.min(4, barSize / 2);
      ctx.beginPath();
      if (horizontal) { const end = v >= 0 ? x + bw : x; const start = v >= 0 ? x : x + bw; if (v >= 0) { ctx.moveTo(start, y); ctx.lineTo(end - r, y); ctx.quadraticCurveTo(end, y, end, y + r); ctx.lineTo(end, y + bh - r); ctx.quadraticCurveTo(end, y + bh, end - r, y + bh); ctx.lineTo(start, y + bh); } else { ctx.moveTo(start, y); ctx.lineTo(end + r, y); ctx.quadraticCurveTo(end, y, end, y + r); ctx.lineTo(end, y + bh - r); ctx.quadraticCurveTo(end, y + bh, end + r, y + bh); ctx.lineTo(start, y + bh); } }
      else { const end = v >= 0 ? y : y + bh; const start = v >= 0 ? y + bh : y; if (v >= 0) { ctx.moveTo(x, start); ctx.lineTo(x, end + r); ctx.quadraticCurveTo(x, end, x + r, end); ctx.lineTo(x + bw - r, end); ctx.quadraticCurveTo(x + bw, end, x + bw, end + r); ctx.lineTo(x + bw, start); } else { ctx.moveTo(x, start); ctx.lineTo(x, end - r); ctx.quadraticCurveTo(x, end, x + r, end); ctx.lineTo(x + bw - r, end); ctx.quadraticCurveTo(x + bw, end, x + bw, end - r); ctx.lineTo(x + bw, start); } }
      ctx.closePath(); ctx.fill();
      if (chart.dataLabels && !stacked) { ctx.fillStyle = INK2; ctx.textAlign = 'center'; const t = formatValue(v, s.numFmt).text; if (horizontal) { ctx.textAlign = 'left'; ctx.fillText(t, x + bw + 4, y + bh / 2); } else ctx.fillText(t, x + bw / 2, (v >= 0 ? y : y + bh) + (v >= 0 ? -9 : 9)); ctx.fillStyle = s.color; }
      hits.push({ x, y, w: bw, h: bh, label: label(s, i, v) });
    }
  }
  return hits;
}

function renderPie(ctx, chart, data, left, top, right, bottom, hits, doughnut) {
  const s = data.series[0];
  const vals = s.values.map((v) => (v != null && v > 0 ? v : 0));
  const total = vals.reduce((a, b) => a + b, 0);
  const cx = (left + right) / 2, cy = (top + bottom) / 2;
  const R = Math.max(10, Math.min(right - left, bottom - top) / 2 - 16);
  if (!total) { ctx.fillStyle = INK2; ctx.textAlign = 'center'; ctx.fillText('No positive values', cx, cy); return hits; }
  let a0 = -Math.PI / 2;
  const slices = [];
  vals.forEach((v, i) => {
    if (!v) return;
    const a1 = a0 + (v / total) * Math.PI * 2;
    ctx.fillStyle = seriesColor(i);
    ctx.beginPath(); ctx.moveTo(cx, cy); ctx.arc(cx, cy, R, a0, a1); ctx.closePath(); ctx.fill();
    ctx.strokeStyle = '#fff'; ctx.lineWidth = 2; ctx.stroke();
    slices.push({ a0, a1, i, v });
    a0 = a1;
  });
  if (doughnut) { ctx.fillStyle = '#fff'; ctx.beginPath(); ctx.arc(cx, cy, R * 0.55, 0, Math.PI * 2); ctx.fill(); }
  // labels for large slices
  ctx.fillStyle = INK; ctx.textAlign = 'center';
  for (const sl of slices) {
    const frac = sl.v / total;
    if (frac > 0.05) { const mid = (sl.a0 + sl.a1) / 2; const rr = doughnut ? R * 0.78 : R * 0.65; const lx = cx + Math.cos(mid) * rr, ly = cy + Math.sin(mid) * rr; ctx.fillStyle = '#fff'; ctx.font = '600 11px "Segoe UI", system-ui, sans-serif'; ctx.fillText(Math.round(frac * 100) + '%', lx, ly); ctx.font = FONT; }
    hits.push({ arc: true, cx, cy, r: R, r0: doughnut ? R * 0.55 : 0, a0: sl.a0, a1: sl.a1, label: `${data.categories[sl.i]}: ${formatValue(sl.v, s.numFmt).text} (${(sl.v / total * 100).toFixed(1)}%)` });
  }
  return hits;
}

export function hitTest(hits, x, y) {
  for (let i = hits.length - 1; i >= 0; i--) {
    const h = hits[i];
    if (h.arc) { const dx = x - h.cx, dy = y - h.cy; const d = Math.hypot(dx, dy); if (d > h.r || d < h.r0) continue; let a = Math.atan2(dy, dx); while (a < h.a0) a += Math.PI * 2; while (a > h.a0 + Math.PI * 2) a -= Math.PI * 2; if (a >= h.a0 && a <= h.a1) return h; continue; }
    if (x >= h.x && x <= h.x + h.w && y >= h.y && y <= h.y + h.h) return h;
  }
  return null;
}

// ---------------- Floating chart objects ----------------
export class ChartLayer {
  constructor(app, overlay) {
    this.app = app; this.overlay = overlay;
    this.nodes = new Map(); // chart id -> { el, canvas, tip, hits, sig }
    this.selectedId = null;
    this.drag = null;
    document.addEventListener('mousemove', (e) => this.onMove(e));
    document.addEventListener('mouseup', (e) => this.onUp(e));
  }
  get sheet() { return this.app.sheet; }
  select(id) { if (this.selectedId === id) return; this.selectedId = id; for (const [cid, n] of this.nodes) n.el.classList.toggle('selected', cid === id); if (id) this.app.editor.cancel(); this.app.refreshToolbar(); }
  selected() { return this.sheet.charts.find((c) => c.id === this.selectedId) || null; }
  deleteSelected() { const ch = this.selected(); if (!ch) return false; this.app.transact('Delete chart', () => this.sheet.setCharts(this.sheet.charts.filter((c) => c.id !== ch.id))); this.select(null); return true; }
  // Called after every grid render: create/position/redraw chart elements for the active sheet
  sync() {
    const sheet = this.sheet; const grid = this.app.grid; const z = grid.zoom;
    const live = new Set();
    for (const ch of sheet.charts) {
      live.add(ch.id);
      let n = this.nodes.get(ch.id);
      if (!n) { n = this.createNode(ch); this.nodes.set(ch.id, n); }
      const o = grid.cellOrigin(ch.anchor.r, ch.anchor.c);
      const x = o.x + (ch.anchor.dx || 0) * z, y = o.y + (ch.anchor.dy || 0) * z;
      const w = Math.max(80, ch.w * z), h = Math.max(60, ch.h * z);
      const vis = o.visible && x + w > grid.layout.headerW && y + h > grid.layout.headerH;
      n.el.style.display = vis ? '' : 'none';
      if (!vis) continue;
      n.el.style.left = x + 'px'; n.el.style.top = y + 'px'; n.el.style.width = w + 'px'; n.el.style.height = h + 'px';
      n.el.classList.toggle('selected', ch.id === this.selectedId);
      const data = chartData(this.app.wb, ch);
      const sig = JSON.stringify([ch.type, ch.title, ch.legend, ch.dataLabels, ch.seriesInRows, w, h, data.series.map((s) => [s.name, s.values, s.numFmt]), data.categories, data.xValues, data.error]);
      if (sig !== n.sig) { n.sig = sig; n.hits = renderChart(n.canvas, ch, data, w, h); }
    }
    for (const [id, n] of this.nodes) if (!live.has(id)) { n.el.remove(); this.nodes.delete(id); if (this.selectedId === id) this.selectedId = null; }
  }
  createNode(ch) {
    const el = document.createElement('div'); el.className = 'chartObj'; el.dataset.id = ch.id;
    const canvas = document.createElement('canvas'); el.appendChild(canvas);
    const handle = document.createElement('div'); handle.className = 'handle'; el.appendChild(handle);
    const tip = document.createElement('div'); tip.className = 'tip'; el.appendChild(tip);
    const n = { el, canvas, tip, hits: [], sig: null };
    el.addEventListener('mousedown', (e) => {
      e.stopPropagation(); e.preventDefault();
      this.select(ch.id);
      const chart = this.sheet.charts.find((c) => c.id === ch.id); if (!chart) return;
      const rect = el.getBoundingClientRect();
      if (e.target === handle) this.drag = { mode: 'resize', id: ch.id, startX: e.clientX, startY: e.clientY, w: chart.w, h: chart.h };
      else this.drag = { mode: 'move', id: ch.id, startX: e.clientX, startY: e.clientY, left: rect.left, top: rect.top, offX: e.clientX - rect.left, offY: e.clientY - rect.top, moved: false };
    });
    el.addEventListener('dblclick', (e) => { e.stopPropagation(); const chart = this.sheet.charts.find((c) => c.id === ch.id); if (chart) this.app.dialogs.chart(chart); });
    el.addEventListener('contextmenu', (e) => { e.preventDefault(); e.stopPropagation(); this.select(ch.id); this.app.showContextMenu(this.app.chartContextItems(ch.id), e.clientX, e.clientY); });
    el.addEventListener('mousemove', (e) => {
      if (this.drag) return;
      const r = canvas.getBoundingClientRect();
      const h = hitTest(n.hits, e.clientX - r.left, e.clientY - r.top);
      if (h) { tip.textContent = h.label; tip.style.display = 'block'; const tx = Math.min(e.clientX - r.left + 12, r.width - tip.offsetWidth - 4); tip.style.left = Math.max(0, tx) + 'px'; tip.style.top = Math.max(0, e.clientY - r.top - 28) + 'px'; } else tip.style.display = 'none';
    });
    el.addEventListener('mouseleave', () => { tip.style.display = 'none'; });
    this.overlay.appendChild(el);
    return n;
  }
  onMove(e) {
    const d = this.drag; if (!d) return;
    const n = this.nodes.get(d.id); if (!n) return;
    const z = this.app.grid.zoom;
    if (d.mode === 'move') {
      const dx = e.clientX - d.startX, dy = e.clientY - d.startY;
      if (Math.abs(dx) + Math.abs(dy) > 2) d.moved = true;
      const ov = this.overlay.getBoundingClientRect();
      n.el.style.left = (d.left - ov.left + dx) + 'px'; n.el.style.top = (d.top - ov.top + dy) + 'px';
    } else {
      const chart = this.sheet.charts.find((c) => c.id === d.id); if (!chart) return;
      const w = Math.max(80, d.w + (e.clientX - d.startX) / z), h = Math.max(60, d.h + (e.clientY - d.startY) / z);
      n.el.style.width = w * z + 'px'; n.el.style.height = h * z + 'px'; d.newW = w; d.newH = h;
    }
  }
  onUp(e) {
    const d = this.drag; if (!d) return; this.drag = null;
    const chart = this.sheet.charts.find((c) => c.id === d.id); if (!chart) return;
    const z = this.app.grid.zoom;
    if (d.mode === 'move') {
      if (!d.moved) return;
      const ov = this.overlay.getBoundingClientRect();
      const x = parseFloat(this.nodes.get(d.id).el.style.left), y = parseFloat(this.nodes.get(d.id).el.style.top);
      const grid = this.app.grid;
      const hit = grid.cellAt(Math.max(grid.layout.headerW, x), Math.max(grid.layout.headerH, y));
      const r = hit.r ?? 0, c = hit.c ?? 0;
      const o = grid.cellOrigin(r, c);
      this.app.transact('Move chart', () => this.sheet.updateChart(chart.id, { anchor: { r, c, dx: Math.round((x - o.x) / z), dy: Math.round((y - o.y) / z) } }));
    } else if (d.newW) {
      this.app.transact('Resize chart', () => this.sheet.updateChart(chart.id, { w: Math.round(d.newW), h: Math.round(d.newH) }));
    }
  }
  exportPng(id) {
    const n = this.nodes.get(id); if (!n) return;
    const ch = this.sheet.charts.find((c) => c.id === id);
    const w = ch.w * 2, h = ch.h * 2; const c = document.createElement('canvas');
    const saveDpr = window.devicePixelRatio; try { Object.defineProperty(window, 'devicePixelRatio', { value: 2, configurable: true }); renderChart(c, ch, chartData(this.app.wb, ch), ch.w, ch.h); } finally { Object.defineProperty(window, 'devicePixelRatio', { value: saveDpr, configurable: true }); }
    c.toBlob((blob) => this.app.downloadBlob(blob, `${(ch.title || 'chart').replace(/[^\w-]+/g, '_')}.png`), 'image/png');
  }
}

export function newChart(type, range, anchor) {
  return { id: uid('ch'), type, range, title: null, legend: true, dataLabels: false, seriesInRows: null, anchor: { r: anchor.r, c: anchor.c, dx: 10, dy: 10 }, w: 480, h: 300 };
}
