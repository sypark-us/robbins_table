// Autofill: extend a source block into a destination range with series detection.
import { shiftFormula } from './formula.js';
import { MONTHS, WEEKDAYS, isDateFormat, serialToDate, dateToSerial } from './format.js';

const lower = (s) => String(s).toLowerCase();
const MONTH_SHORT = MONTHS.map((m) => m.slice(0, 3));
const DAY_SHORT = WEEKDAYS.map((d) => d.slice(0, 3));

function listIndex(list, v) { return typeof v === 'string' ? list.findIndex((x) => lower(x) === lower(v)) : -1; }
function matchCase(template, s) { if (template === template.toUpperCase() && template !== template.toLowerCase()) return s.toUpperCase(); if (template === template.toLowerCase()) return s.toLowerCase(); return s; }

// Describe a one-dimensional sequence of cells [{v, f, s, fmt}] and return fn(k) -> cell for the k-th element (k >= 0, k may exceed length).
function makeSeries(cells, opts = {}) {
  const n = cells.length;
  const vals = cells.map((c) => c.v);
  const allNumbers = vals.every((v) => typeof v === 'number') && n > 0;
  const allDates = allNumbers && cells.every((c) => c.isDate);
  // formulas: copy with shifted refs
  if (cells.every((c) => c.f)) return (k, dr, dc) => { const src = cells[k % n]; return { ...src, f: shiftFormula(src.f, dr(k), dc(k)) }; };
  if (allNumbers) {
    if (n === 1) {
      const step = opts.increment || allDates ? 1 : 0; // a single date fills by one day, like Excel
      return (k) => ({ ...cells[0], v: vals[0] + step * k });
    }
    if (allDates && n >= 2) {
      // detect month/year steps if days differ but day-of-month matches
      const a = serialToDate(vals[0]), b = serialToDate(vals[1]);
      const monthDiff = (b.y - a.y) * 12 + (b.m - a.m);
      const sameDay = a.d === b.d || (b.d >= 28 && a.d >= 28);
      if (monthDiff !== 0 && sameDay && vals.every((v, i) => i === 0 || Math.abs((serialToDate(v).y - a.y) * 12 + serialToDate(v).m - a.m - monthDiff * i) === 0)) {
        return (k) => { const total = a.m - 1 + monthDiff * k; const y = a.y + Math.floor(total / 12), m = ((total % 12) + 12) % 12 + 1; const dim = new Date(Date.UTC(y, m, 0)).getUTCDate(); return { ...cells[k % n], v: dateToSerial(y, m, Math.min(a.d, dim)) + (vals[0] % 1) }; };
      }
    }
    // linear trend (least squares like Excel's fill series)
    const step = n === 2 ? vals[1] - vals[0] : (() => { const xm = (n - 1) / 2, ym = vals.reduce((a, b) => a + b, 0) / n; let num = 0, den = 0; for (let i = 0; i < n; i++) { num += (i - xm) * (vals[i] - ym); den += (i - xm) ** 2; } return den ? num / den : 0; })();
    const consistent = vals.every((v, i) => i === 0 || Math.abs((v - vals[i - 1]) - step) < 1e-9);
    if (consistent) return (k) => ({ ...cells[k % n], v: round(vals[0] + step * k) });
    const ym = vals.reduce((a, b) => a + b, 0) / n, xm = (n - 1) / 2; const b0 = ym - step * xm;
    return (k) => ({ ...cells[k % n], v: round(b0 + step * k) });
  }
  // text lists (months, weekdays) and text with trailing number
  if (vals.every((v) => typeof v === 'string') && n > 0) {
    for (const list of [MONTHS, MONTH_SHORT, WEEKDAYS, DAY_SHORT]) {
      const idx = vals.map((v) => listIndex(list, v));
      if (idx.every((i) => i >= 0)) {
        const step = n >= 2 ? ((idx[1] - idx[0]) + list.length) % list.length || list.length : 1;
        if (idx.every((v, i) => i === 0 || ((v - idx[i - 1]) + list.length) % list.length === step % list.length)) return (k) => ({ ...cells[k % n], v: matchCase(vals[0], list[(idx[0] + step * k) % list.length]) });
      }
    }
    const parts = vals.map((v) => /^(.*?)(\d+)(\D*)$/.exec(v));
    if (parts.every(Boolean) && parts.every((p) => p[1] === parts[0][1] && p[3] === parts[0][3])) {
      const nums = parts.map((p) => parseInt(p[2], 10));
      const step = n >= 2 ? nums[1] - nums[0] : 1;
      if (nums.every((v, i) => i === 0 || v - nums[i - 1] === step)) {
        const width = parts[0][2].length;
        return (k) => { const val = nums[0] + step * k; const s = String(Math.abs(val)).padStart(parts[0][2].startsWith('0') ? width : 0, '0'); return { ...cells[k % n], v: `${parts[0][1]}${val < 0 ? '-' : ''}${s}${parts[0][3]}` }; };
      }
    }
  }
  // mixed: numbers inside a mixed block increment per occurrence when a single number present; otherwise repeat pattern
  return (k, dr, dc) => { const src = cells[k % n]; if (src.f) return { ...src, f: shiftFormula(src.f, dr(k), dc(k)) }; return { ...src }; };
}
const round = (v) => Math.abs(v) < 1e-10 ? 0 : Number(v.toPrecision(15));

// src: {r1,c1,r2,c2}; dst: full range including src (extends down/up/right/left). opts: {increment:boolean (ctrl), copyOnly:boolean}
export function autofill(sheet, src, dst, opts = {}) {
  const wb = sheet.wb;
  const vertical = dst.r2 - dst.r1 > src.r2 - src.r1;
  const styleOf = (r, c) => sheet.getStyleId(r, c);
  const cellInfo = (r, c) => { const cell = sheet.getCell(r, c); const s = styleOf(r, c); const fmt = wb.styles[s] && wb.styles[s].numFmt; return { v: cell && !cell.f ? cell.v : null, f: cell ? cell.f : null, s, note: cell ? cell.note : null, isDate: isDateFormat(fmt) || !!(cell && cell.autoFmt) }; };
  if (vertical) {
    const down = dst.r2 > src.r2;
    for (let c = src.c1; c <= src.c2; c++) {
      let cells = []; for (let r = src.r1; r <= src.r2; r++) cells.push(cellInfo(r, c));
      if (!down) cells = cells.reverse();
      const series = opts.copyOnly ? ((k, dr) => ({ ...cells[k % cells.length], f: cells[k % cells.length].f ? shiftFormula(cells[k % cells.length].f, dr(k), 0) : null })) : makeSeries(cells, opts);
      const n = cells.length;
      if (down) for (let r = src.r2 + 1, k = n; r <= dst.r2; r++, k++) write(sheet, r, c, series(k, (kk) => (r - (src.r1 + (kk % n))), () => 0));
      else for (let r = src.r1 - 1, k = n; r >= dst.r1; r--, k++) write(sheet, r, c, series(k, (kk) => (r - (src.r2 - (kk % n))), () => 0));
    }
  } else {
    const right = dst.c2 > src.c2;
    for (let r = src.r1; r <= src.r2; r++) {
      let cells = []; for (let c = src.c1; c <= src.c2; c++) cells.push(cellInfo(r, c));
      if (!right) cells = cells.reverse();
      const series = opts.copyOnly ? ((k, dr, dc) => ({ ...cells[k % cells.length], f: cells[k % cells.length].f ? shiftFormula(cells[k % cells.length].f, 0, dc(k)) : null })) : makeSeries(cells, opts);
      const n = cells.length;
      if (right) for (let c = src.c2 + 1, k = n; c <= dst.c2; c++, k++) write(sheet, r, c, series(k, () => 0, (kk) => (c - (src.c1 + (kk % n)))));
      else for (let c = src.c1 - 1, k = n; c >= dst.c1; c--, k++) write(sheet, r, c, series(k, () => 0, (kk) => (c - (src.c2 - (kk % n)))));
    }
  }
}
function write(sheet, r, c, cell) {
  sheet._setCellRaw(r, c, { v: cell.f ? null : cell.v, f: cell.f || null, s: cell.s || 0 });
}

// Fill a range with the content of its first row/column (Ctrl+D / Ctrl+R)
export function fillDirection(sheet, rg, dir) {
  if (dir === 'down') for (let c = rg.c1; c <= rg.c2; c++) { const src = sheet.getCell(rg.r1, c); for (let r = rg.r1 + 1; r <= rg.r2; r++) sheet._setCellRaw(r, c, src ? { ...src, f: src.f ? shiftFormula(src.f, r - rg.r1, 0) : null, v: src.f ? null : src.v } : null); }
  else if (dir === 'right') for (let r = rg.r1; r <= rg.r2; r++) { const src = sheet.getCell(r, rg.c1); for (let c = rg.c1 + 1; c <= rg.c2; c++) sheet._setCellRaw(r, c, src ? { ...src, f: src.f ? shiftFormula(src.f, 0, c - rg.c1) : null, v: src.f ? null : src.v } : null); }
  else if (dir === 'up') for (let c = rg.c1; c <= rg.c2; c++) { const src = sheet.getCell(rg.r2, c); for (let r = rg.r1; r < rg.r2; r++) sheet._setCellRaw(r, c, src ? { ...src, f: src.f ? shiftFormula(src.f, r - rg.r2, 0) : null, v: src.f ? null : src.v } : null); }
  else if (dir === 'left') for (let r = rg.r1; r <= rg.r2; r++) { const src = sheet.getCell(r, rg.c2); for (let c = rg.c1; c < rg.c2; c++) sheet._setCellRaw(r, c, src ? { ...src, f: src.f ? shiftFormula(src.f, 0, c - rg.c2) : null, v: src.f ? null : src.v } : null); }
}
