// Canvas grid renderer: virtualised rows/columns, frozen panes, selection, headers, mouse interaction.
import { formatValue, formatGeneral, isDateFormat } from './format.js';
import { colToName, clamp, key as cellKey, rangeContains, normRange, MAX_ROWS, MAX_COLS, mixColor, isDarkColor } from './util.js';
import { isErr, parse, shiftFormula } from './formula.js';
import { num, bool, criteria as makeCriteria } from './functions.js';
import { isRange, isArr } from './range.js';

const HEADER_H = 20, HEADER_W_MIN = 36, SCROLLBAR = 14;
const DEFAULT_FONT_PX = 14.67; // Calibri 11pt at 96 dpi
const FONT_FALLBACK = 'Calibri, "Segoe UI", Arial, sans-serif';
const REF_COLORS = ['#1a73e8', '#e8710a', '#9334e6', '#188038', '#d93025', '#007b83', '#b31412', '#a142f4'];
const BORDER_WIDTH = { thin: 1, hair: 1, dotted: 1, dashed: 1, medium: 2, mediumDashed: 2, thick: 3, double: 3 };

export class GridView {
  constructor(app, wrap) {
    this.app = app; this.wrap = wrap;
    this.canvas = wrap.querySelector('#grid'); this.ctx = this.canvas.getContext('2d');
    this.overlay = wrap.querySelector('#overlay');
    this.vscroll = wrap.querySelector('#vscroll'); this.hscroll = wrap.querySelector('#hscroll');
    this.refHighlights = []; this.copyRange = null; this.hover = null;
    this.layout = { rows: [], cols: [], headerW: 40, headerH: HEADER_H, frozenW: 0, frozenH: 0, W: 0, H: 0 };
    this.measureCache = new Map();
    this.drag = null; this._raf = 0; this._ignoreScroll = false; this._dash = 0; this._marqueeTimer = null;
    this.dpr = window.devicePixelRatio || 1;
    this.cfCache = new Map();
    new ResizeObserver(() => this.resize()).observe(wrap);
    this.vscroll.addEventListener('scroll', () => this.onScrollbar('v'));
    this.hscroll.addEventListener('scroll', () => this.onScrollbar('h'));
    this.canvas.addEventListener('wheel', (e) => this.onWheel(e), { passive: false });
    this.overlay.addEventListener('wheel', (e) => this.onWheel(e), { passive: false });
    this.canvas.addEventListener('mousedown', (e) => this.onMouseDown(e));
    this.canvas.addEventListener('dblclick', (e) => this.onDblClick(e));
    this.canvas.addEventListener('contextmenu', (e) => this.onContextMenu(e));
    this.canvas.addEventListener('mousemove', (e) => this.onHover(e));
    this.canvas.addEventListener('mouseleave', () => { this.setHover(null); });
    document.addEventListener('mousemove', (e) => this.onDragMove(e));
    document.addEventListener('mouseup', (e) => this.onDragEnd(e));
    this.resize();
  }
  get sheet() { return this.app.sheet; }
  get zoom() { return this.sheet.view.zoom || 1; }
  get scrollRow() { return this.sheet.view.scrollRow || 0; }
  set scrollRow(v) { this.sheet.view.scrollRow = clamp(v, this.sheet.freeze.rows, MAX_ROWS - 1); }
  get scrollCol() { return this.sheet.view.scrollCol || 0; }
  set scrollCol(v) { this.sheet.view.scrollCol = clamp(v, this.sheet.freeze.cols, MAX_COLS - 1); }

  resize() {
    const r = this.wrap.getBoundingClientRect();
    this.dpr = window.devicePixelRatio || 1;
    const W = Math.max(0, Math.floor(r.width) - SCROLLBAR), H = Math.max(0, Math.floor(r.height) - SCROLLBAR);
    this.canvas.style.width = W + 'px'; this.canvas.style.height = H + 'px';
    this.canvas.width = Math.round(W * this.dpr); this.canvas.height = Math.round(H * this.dpr);
    this.cssW = W; this.cssH = H;
    this.invalidate();
  }
  invalidate() { if (!this._raf) this._raf = requestAnimationFrame(() => { this._raf = 0; this.render(); }); }

  // ---------- layout ----------
  computeLayout() {
    const s = this.sheet, z = this.zoom, fz = s.freeze;
    const headerH = Math.round(HEADER_H * z);
    const approxLast = this.scrollRow + Math.ceil(this.cssH / (s.rowAxis.default * z)) + 1;
    const headerW = Math.round(Math.max(HEADER_W_MIN, 14 + String(approxLast + 1).length * 7.5) * z);
    const cols = [], rows = [];
    let x = headerW;
    for (let c = 0; c < fz.cols && c < MAX_COLS; c++) { const w = s.colAxis.size(c) * z; if (w > 0) cols.push({ c, x, w, frozen: true }); x += w; }
    const frozenW = x - headerW;
    if (this.scrollCol < fz.cols) this.sheet.view.scrollCol = fz.cols;
    for (let c = this.scrollCol; x < this.cssW && c < MAX_COLS; c++) { const w = s.colAxis.size(c) * z; if (w <= 0) continue; cols.push({ c, x, w }); x += w; }
    let y = headerH;
    for (let r = 0; r < fz.rows && r < MAX_ROWS; r++) { const h = s.rowAxis.size(r) * z; if (h > 0) rows.push({ r, y, h, frozen: true }); y += h; }
    const frozenH = y - headerH;
    if (this.scrollRow < fz.rows) this.sheet.view.scrollRow = fz.rows;
    for (let r = this.scrollRow; y < this.cssH && r < MAX_ROWS; r++) { const h = s.rowAxis.size(r) * z; if (h <= 0) continue; rows.push({ r, y, h }); y += h; }
    this.layout = { rows, cols, headerW, headerH, frozenW, frozenH, W: this.cssW, H: this.cssH, rowMap: new Map(rows.map((e) => [e.r, e])), colMap: new Map(cols.map((e) => [e.c, e])) };
  }
  visibleRowCount() { return this.layout.rows.filter((e) => !e.frozen && e.y + e.h <= this.cssH).length; }
  visibleColCount() { return this.layout.cols.filter((e) => !e.frozen && e.x + e.w <= this.cssW).length; }
  // Canvas coordinates of a cell's top-left (even when off screen); visible=false when hidden by frozen panes
  cellOrigin(r, c) {
    const s = this.sheet, z = this.zoom, L = this.layout, fz = s.freeze;
    let x, y, visible = true;
    if (c < fz.cols) x = L.headerW + s.colAxis.offset(c) * z;
    else { x = L.headerW + L.frozenW + (s.colAxis.offset(c) - s.colAxis.offset(this.scrollCol)) * z; if (c < this.scrollCol) visible = false; }
    if (r < fz.rows) y = L.headerH + s.rowAxis.offset(r) * z;
    else { y = L.headerH + L.frozenH + (s.rowAxis.offset(r) - s.rowAxis.offset(this.scrollRow)) * z; if (r < this.scrollRow) visible = false; }
    return { x, y, visible };
  }
  cellRect(r, c, useMerge = true) {
    const s = this.sheet;
    if (useMerge) { const m = s.mergeAt(r, c); if (m) { const a = this.cellRect(m.r1, m.c1, false), b = this.cellRect(m.r2, m.c2, false); const o1 = this.cellOrigin(m.r1, m.c1), o2 = this.cellOrigin(m.r2, m.c2); if (!a && !b && !(o1.visible && o2.visible)) return null; const x = o1.x, y = o1.y; const w = o2.x + s.colAxis.size(m.c2) * this.zoom - x, h = o2.y + s.rowAxis.size(m.r2) * this.zoom - y; return { x, y, w, h, merge: m }; } }
    const re = this.layout.rowMap.get(r), ce = this.layout.colMap.get(c);
    if (!re || !ce) return null;
    return { x: ce.x, y: re.y, w: ce.w, h: re.h };
  }
  // Hit test: returns {type:'corner'|'colHeader'|'rowHeader'|'cell'|'colResize'|'rowResize', r, c}
  cellAt(x, y) {
    const L = this.layout, s = this.sheet, z = this.zoom;
    const colAt = (px) => { for (const e of L.cols) if (px >= e.x && px < e.x + e.w) return e.c; const last = L.cols[L.cols.length - 1]; if (!last) return this.scrollCol; if (px >= last.x + last.w) return Math.min(MAX_COLS - 1, s.colAxis.indexAt(s.colAxis.offset(last.c) + (px - last.x) / z, MAX_COLS - 1)); return L.cols[0].c; };
    const rowAt = (py) => { for (const e of L.rows) if (py >= e.y && py < e.y + e.h) return e.r; const last = L.rows[L.rows.length - 1]; if (!last) return this.scrollRow; if (py >= last.y + last.h) return Math.min(MAX_ROWS - 1, s.rowAxis.indexAt(s.rowAxis.offset(last.r) + (py - last.y) / z, MAX_ROWS - 1)); return L.rows[0].r; };
    if (x < L.headerW && y < L.headerH) return { type: 'corner' };
    if (y < L.headerH) {
      for (const e of L.cols) if (Math.abs(x - (e.x + e.w)) <= 4) return { type: 'colResize', c: e.c };
      if (x < L.headerW) return { type: 'corner' };
      return { type: 'colHeader', c: colAt(x) };
    }
    if (x < L.headerW) {
      for (const e of L.rows) if (Math.abs(y - (e.y + e.h)) <= 4) return { type: 'rowResize', r: e.r };
      return { type: 'rowHeader', r: rowAt(y) };
    }
    return { type: 'cell', r: rowAt(y), c: colAt(x) };
  }

  // ---------- scrolling ----------
  updateScrollbars() {
    const s = this.sheet, z = this.zoom;
    const ub = s.usedBounds();
    const vis = Math.max(1, this.visibleRowCount());
    const lastRow = Math.min(MAX_ROWS - 1, Math.max(ub.r2 + 60, this.scrollRow + vis + 40, 100));
    const lastCol = Math.min(MAX_COLS - 1, Math.max(ub.c2 + 15, this.scrollCol + Math.max(1, this.visibleColCount()) + 10, 26));
    this._ignoreScroll = true;
    const vh = (s.rowAxis.offset(lastRow + 1) - s.rowAxis.offset(s.freeze.rows)) * z + (this.cssH - this.layout.headerH - this.layout.frozenH) - s.rowAxis.size(lastRow) * z;
    const hw = (s.colAxis.offset(lastCol + 1) - s.colAxis.offset(s.freeze.cols)) * z + (this.cssW - this.layout.headerW - this.layout.frozenW) - s.colAxis.size(lastCol) * z;
    this.vscroll.firstElementChild.style.height = Math.max(this.cssH + 1, vh) + 'px';
    this.hscroll.firstElementChild.style.width = Math.max(this.cssW + 1, hw) + 'px';
    const top = (s.rowAxis.offset(this.scrollRow) - s.rowAxis.offset(s.freeze.rows)) * z;
    const left = (s.colAxis.offset(this.scrollCol) - s.colAxis.offset(s.freeze.cols)) * z;
    if (Math.abs(this.vscroll.scrollTop - top) > 1) this.vscroll.scrollTop = top;
    if (Math.abs(this.hscroll.scrollLeft - left) > 1) this.hscroll.scrollLeft = left;
    requestAnimationFrame(() => { this._ignoreScroll = false; });
  }
  onScrollbar(which) {
    if (this._ignoreScroll) return;
    const s = this.sheet, z = this.zoom;
    if (which === 'v') { const pos = this.vscroll.scrollTop / z + s.rowAxis.offset(s.freeze.rows); const r = s.rowAxis.indexAt(pos, MAX_ROWS - 1); if (r !== this.scrollRow) { this.scrollRow = r; this.invalidate(); } }
    else { const pos = this.hscroll.scrollLeft / z + s.colAxis.offset(s.freeze.cols); const c = s.colAxis.indexAt(pos, MAX_COLS - 1); if (c !== this.scrollCol) { this.scrollCol = c; this.invalidate(); } }
    this.app.editor.updatePosition();
  }
  onWheel(e) {
    e.preventDefault();
    if (e.ctrlKey) { this.app.setZoom(this.zoom * (e.deltaY < 0 ? 1.1 : 1 / 1.1)); return; }
    const lines = e.deltaMode === 1 ? Math.abs(e.deltaY) : Math.max(1, Math.round(Math.abs(e.deltaY) / 40) * 3);
    const dirY = Math.sign(e.deltaY), dirX = Math.sign(e.deltaX);
    if (e.shiftKey && dirY) this.scrollBy(0, dirY * 3);
    else { if (dirY) this.scrollBy(dirY * lines, 0); if (dirX) this.scrollBy(0, dirX * 3); }
  }
  scrollBy(dr, dc) {
    const s = this.sheet;
    if (dr) { let r = this.scrollRow; let n = Math.abs(dr); while (n > 0) { r += Math.sign(dr); if (r < s.freeze.rows || r >= MAX_ROWS) { r = clamp(r, s.freeze.rows, MAX_ROWS - 1); break; } if (s.rowAxis.size(r) > 0) n--; } this.scrollRow = r; }
    if (dc) { let c = this.scrollCol; let n = Math.abs(dc); while (n > 0) { c += Math.sign(dc); if (c < s.freeze.cols || c >= MAX_COLS) { c = clamp(c, s.freeze.cols, MAX_COLS - 1); break; } if (s.colAxis.size(c) > 0) n--; } this.scrollCol = c; }
    this.invalidate(); this.app.editor.updatePosition();
  }
  // Ensure the cell is visible, scrolling minimally
  scrollTo(r, c) {
    const s = this.sheet, z = this.zoom, L = this.layout, fz = s.freeze;
    let changed = false;
    if (r >= fz.rows) {
      if (r < this.scrollRow) { this.scrollRow = r; changed = true; }
      else {
        const avail = this.cssH - L.headerH - L.frozenH;
        const bottom = (s.rowAxis.offset(r + 1) - s.rowAxis.offset(this.scrollRow)) * z;
        if (bottom > avail) { let sr = this.scrollRow; while (sr < r && (s.rowAxis.offset(r + 1) - s.rowAxis.offset(sr)) * z > avail) sr++; this.scrollRow = sr; changed = true; }
      }
    }
    if (c >= fz.cols) {
      if (c < this.scrollCol) { this.scrollCol = c; changed = true; }
      else {
        const avail = this.cssW - L.headerW - L.frozenW;
        const right = (s.colAxis.offset(c + 1) - s.colAxis.offset(this.scrollCol)) * z;
        if (right > avail) { let sc = this.scrollCol; while (sc < c && (s.colAxis.offset(c + 1) - s.colAxis.offset(sc)) * z > avail) sc++; this.scrollCol = sc; changed = true; }
      }
    }
    if (changed) { this.computeLayout(); this.invalidate(); }
    return changed;
  }

  // ---------- text helpers ----------
  fontFor(st, z = this.zoom) {
    const size = (st.size || 11) * 96 / 72 * z;
    return `${st.italic ? 'italic ' : ''}${st.bold ? 'bold ' : ''}${size.toFixed(2)}px ${st.font ? `"${st.font}", ` : ''}${FONT_FALLBACK}`;
  }
  measure(text, font) {
    let m = this.measureCache.get(font);
    if (!m) { m = new Map(); this.measureCache.set(font, m); if (this.measureCache.size > 60) this.measureCache.clear(); }
    let w = m.get(text);
    if (w === undefined) { this.ctx.font = font; w = this.ctx.measureText(text).width; if (m.size > 5000) m.clear(); m.set(text, w); }
    return w;
  }
  wrapLines(text, font, width) {
    const out = [];
    for (const para of String(text).split('\n')) {
      const words = para.split(' ');
      let line = '';
      for (const w of words) {
        const t = line ? line + ' ' + w : w;
        if (this.measure(t, font) <= width || !line) {
          if (this.measure(t, font) > width && !line) { // break long word
            let cur = '';
            for (const ch of w) { if (this.measure(cur + ch, font) > width && cur) { out.push(cur); cur = ch; } else cur += ch; }
            line = cur;
          } else line = t;
        } else { out.push(line); line = w; }
      }
      out.push(line);
    }
    return out;
  }
  // display info for a cell: text, color, alignment
  display(sheet, r, c, st, cell) {
    const v = sheet.getValue(r, c);
    if (v == null) return null;
    let text, color = null, kind;
    if (isErr(v) || (v && v.error)) { text = v.error; kind = 'error'; }
    else if (typeof v === 'boolean') { text = v ? 'TRUE' : 'FALSE'; kind = 'bool'; }
    else if (typeof v === 'number') { const f = formatValue(v, sheet.effectiveNumFmt(r, c)); text = f.text; color = f.color; kind = 'number'; }
    else { const code = st.numFmt; if (code && code.includes('@')) text = formatValue(String(v), code).text; else text = String(v); kind = 'text'; }
    return { text, color, kind, value: v };
  }

  // ---------- conditional formatting ----------
  cfFor(sheet, r, c, cache) {
    if (!sheet.cf.length) return null;
    let result = null;
    for (const rule of sheet.cf) {
      if (!rangeContains(rule, r, c)) continue;
      const v = sheet.getValue(r, c);
      let stats = cache.get(rule);
      if (!stats) { stats = this.cfStats(sheet, rule); cache.set(rule, stats); }
      let apply = false, extra = null;
      try {
        switch (rule.type) {
          case 'cellIs': { if (typeof v !== 'number' && typeof v !== 'string' && typeof v !== 'boolean') break; const a = this.cfValue(sheet, rule, rule.v1, r, c), b = this.cfValue(sheet, rule, rule.v2, r, c); const cmp = (x, y) => (typeof x === 'string' && typeof y === 'string') ? x.toLowerCase().localeCompare(y.toLowerCase()) : (typeof x === 'number' && typeof y === 'number') ? x - y : NaN; const d = cmp(v, a); if (isNaN(d) && rule.op !== 'ne' && rule.op !== 'eq') break; switch (rule.op) { case 'gt': apply = d > 0; break; case 'lt': apply = d < 0; break; case 'ge': apply = d >= 0; break; case 'le': apply = d <= 0; break; case 'eq': apply = d === 0; break; case 'ne': apply = d !== 0 && !isNaN(d) || (isNaN(d) && typeof v !== typeof a); break; case 'between': { const e = cmp(v, b); apply = !isNaN(e) && d >= 0 && e <= 0; break; } case 'notBetween': { const e = cmp(v, b); apply = !isNaN(e) && (d < 0 || e > 0); break; } } break; }
          case 'containsText': apply = v != null && String(this.app.fmt(sheet, r, c).text).toLowerCase().includes(String(rule.v1 || '').toLowerCase()); break;
          case 'notContainsText': apply = v != null && !String(this.app.fmt(sheet, r, c).text).toLowerCase().includes(String(rule.v1 || '').toLowerCase()); break;
          case 'beginsWith': apply = v != null && String(this.app.fmt(sheet, r, c).text).toLowerCase().startsWith(String(rule.v1 || '').toLowerCase()); break;
          case 'endsWith': apply = v != null && String(this.app.fmt(sheet, r, c).text).toLowerCase().endsWith(String(rule.v1 || '').toLowerCase()); break;
          case 'duplicate': apply = v != null && stats.counts.get(keyOf(v)) > 1; break;
          case 'unique': apply = v != null && stats.counts.get(keyOf(v)) === 1; break;
          case 'top': apply = typeof v === 'number' && stats.sorted.length > 0 && v >= stats.sorted[Math.max(0, Math.min(stats.sorted.length, rule.percent ? Math.ceil(stats.sorted.length * (rule.rank || 10) / 100) : (rule.rank || 10)) - 1)]; break;
          case 'bottom': apply = typeof v === 'number' && stats.sorted.length > 0 && v <= stats.sorted[stats.sorted.length - Math.max(1, Math.min(stats.sorted.length, rule.percent ? Math.ceil(stats.sorted.length * (rule.rank || 10) / 100) : (rule.rank || 10)))]; break;
          case 'aboveAverage': apply = typeof v === 'number' && v > stats.avg; break;
          case 'belowAverage': apply = typeof v === 'number' && v < stats.avg; break;
          case 'blank': apply = v == null || v === ''; break;
          case 'notBlank': apply = !(v == null || v === ''); break;
          case 'formula': { const res = this.cfValue(sheet, rule, rule.v1, r, c); apply = res === true || (typeof res === 'number' && res !== 0); break; }
          case 'colorScale': { if (typeof v !== 'number' || stats.min == null) break; const cs = rule.colors || ['#f8696b', '#ffeb84', '#63be7b']; const t = stats.max === stats.min ? 0.5 : (v - stats.min) / (stats.max - stats.min); let fill; if (cs.length >= 3) fill = t < 0.5 ? mixColor(cs[0], cs[1], t * 2) : mixColor(cs[1], cs[2], (t - 0.5) * 2); else fill = mixColor(cs[0], cs[cs.length - 1], t); result = { ...(result || {}), fill }; continue; }
          case 'dataBar': { if (typeof v !== 'number' || stats.min == null) break; const lo = Math.min(0, stats.min), hi = Math.max(0, stats.max); const ratio = hi === lo ? 0 : (v - lo) / (hi - lo); const zero = hi === lo ? 0 : (0 - lo) / (hi - lo); result = { ...(result || {}), bar: { from: Math.min(zero, ratio), to: Math.max(zero, ratio), color: rule.barColor || '#638ec6', neg: v < 0 } }; continue; }
        }
      } catch { apply = false; }
      if (apply && rule.style) { result = { ...(rule.style), ...(result || {}) }; if (rule.stopIfTrue) break; }
    }
    return result;
  }
  cfStats(sheet, rule) {
    const stats = { counts: new Map(), sorted: [], min: null, max: null, avg: 0 };
    const r2 = Math.min(rule.r2, sheet.usedBounds().r2), c2 = Math.min(rule.c2, sheet.usedBounds().c2);
    let sum = 0, n = 0;
    for (let r = rule.r1; r <= r2; r++) for (let c = rule.c1; c <= c2; c++) {
      const v = sheet.getValue(r, c); if (v == null) continue;
      stats.counts.set(keyOf(v), (stats.counts.get(keyOf(v)) || 0) + 1);
      if (typeof v === 'number') { stats.sorted.push(v); sum += v; n++; stats.min = stats.min == null ? v : Math.min(stats.min, v); stats.max = stats.max == null ? v : Math.max(stats.max, v); }
    }
    stats.sorted.sort((a, b) => b - a); stats.avg = n ? sum / n : 0;
    return stats;
  }
  cfValue(sheet, rule, expr, r, c) {
    if (expr == null || expr === '') return null;
    const s = String(expr);
    if (!s.startsWith('=')) { const n = Number(s); if (s.trim() !== '' && !isNaN(n)) return n; const low = s.toLowerCase(); if (low === 'true') return true; if (low === 'false') return false; return s.replace(/^"(.*)"$/, '$1'); }
    const dr = r - rule.r1, dc = c - rule.c1;
    const text = shiftFormula(s, dr, dc);
    let ast = this.cfCache.get(text);
    if (!ast) { try { ast = parse(text); } catch { return null; } if (this.cfCache.size > 2000) this.cfCache.clear(); this.cfCache.set(text, ast); }
    const engine = this.app.wb.engine;
    let v = engine.evalNode(ast, engine.makeContext(sheet, r, c, []));
    if (isRange(v)) v = v.isCell ? v.value() : v.toArray()[0][0];
    if (isArr(v)) v = v[0][0];
    return isErr(v) ? null : v;
  }

  // ---------- rendering ----------
  render() {
    const ctx = this.ctx, s = this.sheet, z = this.zoom, app = this.app;
    this.computeLayout();
    const L = this.layout;
    ctx.setTransform(this.dpr, 0, 0, this.dpr, 0, 0);
    ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, this.cssW, this.cssH);
    const cfCache = new Map();
    const styles = app.wb.styles;
    const mergeSet = new Set(); const mergesVisible = [];
    for (const m of s.merges) { const rr = this.cellRect(m.r1, m.c1); if (rr && rr.x < this.cssW && rr.y < this.cssH && rr.x + rr.w > L.headerW && rr.y + rr.h > L.headerH) { mergesVisible.push({ m, rect: rr }); for (let r = m.r1; r <= m.r2; r++) for (let c = m.c1; c <= m.c2; c++) mergeSet.add(cellKey(r, c)); } }
    // pass 1: fills, data bars
    const cellInfos = [];
    for (const re of L.rows) for (const ce of L.cols) {
      const k = cellKey(re.r, ce.c);
      const sid = s.getStyleId(re.r, ce.c);
      const st = styles[sid] || {};
      const cf = this.cfFor(s, re.r, ce.c, cfCache);
      const fill = (cf && cf.fill) || st.fill;
      if (mergeSet.has(k)) { cellInfos.push({ re, ce, st, cf, merged: true }); continue; }
      if (fill) { ctx.fillStyle = fill; ctx.fillRect(ce.x, re.y, ce.w, re.h); }
      if (cf && cf.bar) { const b = cf.bar; const x0 = ce.x + 1 + b.from * (ce.w - 2), x1 = ce.x + 1 + b.to * (ce.w - 2); ctx.fillStyle = b.neg ? '#ff555599' : b.color + '99'; ctx.fillRect(x0, re.y + 2, Math.max(1, x1 - x0), re.h - 4); }
      cellInfos.push({ re, ce, st, cf });
    }
    // gridlines
    if (s.showGridlines) {
      ctx.strokeStyle = '#e2e3e3'; ctx.lineWidth = 1; ctx.beginPath();
      for (const ce of L.cols) { const x = Math.round(ce.x + ce.w) - 0.5; ctx.moveTo(x, L.headerH); ctx.lineTo(x, this.cssH); }
      for (const re of L.rows) { const y = Math.round(re.y + re.h) - 0.5; ctx.moveTo(L.headerW, y); ctx.lineTo(this.cssW, y); }
      ctx.stroke();
    }
    // merged cells: cover gridlines with background
    for (const { m, rect } of mergesVisible) { const st = styles[s.getStyleId(m.r1, m.c1)] || {}; const cf = this.cfFor(s, m.r1, m.c1, cfCache); ctx.fillStyle = (cf && cf.fill) || st.fill || '#fff'; ctx.fillRect(rect.x, rect.y, rect.w - 1, rect.h - 1); }
    // pass 2: text
    const occupied = new Map(); // for overflow: key -> true when cell has content
    for (const info of cellInfos) { const cell = s.getCell(info.re.r, info.ce.c); if (cell && (cell.v != null || cell.f)) occupied.set(cellKey(info.re.r, info.ce.c), true); }
    for (const info of cellInfos) {
      if (info.merged) continue;
      const { re, ce, st, cf } = info;
      const cell = s.getCell(re.r, ce.c);
      if (!cell || (cell.v == null && !cell.f)) continue;
      this.drawCellText(s, re.r, ce.c, { x: ce.x, y: re.y, w: ce.w, h: re.h }, st, cf, cell, occupied, re.frozen, ce.frozen);
    }
    for (const { m, rect } of mergesVisible) { const cell = s.getCell(m.r1, m.c1); if (!cell || (cell.v == null && !cell.f)) continue; const st = styles[s.getStyleId(m.r1, m.c1)] || {}; this.drawCellText(s, m.r1, m.c1, rect, st, this.cfFor(s, m.r1, m.c1, cfCache), cell, null, false, false, true); }
    // pass 3: borders
    for (const info of cellInfos) {
      const { re, ce, st } = info;
      if (!st.border || info.merged) continue;
      this.drawBorders(st.border, ce.x, re.y, ce.w, re.h, z);
    }
    for (const { m, rect } of mergesVisible) { const st = styles[s.getStyleId(m.r1, m.c1)] || {}; if (st.border) this.drawBorders(st.border, rect.x, rect.y, rect.w, rect.h, z); }
    // decorations: notes, filter buttons, spill outline, validation arrows
    for (const info of cellInfos) {
      const { re, ce } = info; const cell = s.getCell(re.r, ce.c);
      if (cell && cell.note) { ctx.fillStyle = '#d93025'; ctx.beginPath(); ctx.moveTo(ce.x + ce.w - 7, re.y); ctx.lineTo(ce.x + ce.w - 1, re.y); ctx.lineTo(ce.x + ce.w - 1, re.y + 6); ctx.closePath(); ctx.fill(); }
    }
    if (s.filter) for (let c = s.filter.c1; c <= s.filter.c2; c++) { const rect = this.cellRect(s.filter.r1, c, false); if (!rect) continue; this.drawFilterButton(rect, s.filter.conds && s.filter.conds[c]); }
    const sel = s.view;
    const act = sel.active;
    const actCell = s.getCell(act.r, act.c);
    const spillOrigin = actCell && actCell.spill ? actCell : (actCell && actCell.spillOf ? s.getCell(...actCell.spillOf.split(',').map(Number)) : null);
    if (spillOrigin && spillOrigin.spill) { const sp = spillOrigin.spill; const a = this.cellOrigin(sp.r1, sp.c1), b = this.cellOrigin(sp.r2, sp.c2); ctx.strokeStyle = '#1a73e8'; ctx.lineWidth = 1; ctx.strokeRect(a.x + 0.5, a.y + 0.5, b.x + s.colAxis.size(sp.c2) * z - a.x - 1, b.y + s.rowAxis.size(sp.r2) * z - a.y - 1); }
    // selection
    this.drawSelection(ctx, s, z);
    // frozen pane separators
    ctx.strokeStyle = '#9e9e9e'; ctx.lineWidth = 1; ctx.beginPath();
    if (s.freeze.cols) { const x = Math.round(L.headerW + L.frozenW) - 0.5; ctx.moveTo(x, 0); ctx.lineTo(x, this.cssH); }
    if (s.freeze.rows) { const y = Math.round(L.headerH + L.frozenH) - 0.5; ctx.moveTo(0, y); ctx.lineTo(this.cssW, y); }
    ctx.stroke();
    this.drawHeaders(ctx, s, z);
    // validation dropdown arrow on the active cell
    const dv = s.validationAt(act.r, act.c);
    if (dv && dv.type === 'list' && dv.showDropdown !== false && !app.editor.isEditing) { const rect = this.cellRect(act.r, act.c); if (rect) { ctx.fillStyle = '#fff'; ctx.strokeStyle = '#999'; const bx = rect.x + rect.w + 1, by = rect.y, bw = Math.round(18 * z), bh = rect.h; ctx.fillRect(bx, by, bw, bh); ctx.strokeRect(bx + 0.5, by + 0.5, bw - 1, bh - 1); ctx.fillStyle = '#333'; ctx.beginPath(); ctx.moveTo(bx + bw / 2 - 4 * z, by + bh / 2 - 2 * z); ctx.lineTo(bx + bw / 2 + 4 * z, by + bh / 2 - 2 * z); ctx.lineTo(bx + bw / 2, by + bh / 2 + 3 * z); ctx.closePath(); ctx.fill(); this.dvButton = { x: bx, y: by, w: bw, h: bh }; } } else this.dvButton = null;
    this.updateScrollbars();
    app.charts.sync();
    app.editor.updatePosition();
    app.onRendered && app.onRendered();
  }

  drawCellText(s, r, c, rect, st, cf, cell, occupied, frozenRow, frozenCol, isMerge = false) {
    const ctx = this.ctx, z = this.zoom;
    const d = this.display(s, r, c, st, cell);
    if (!d || d.text === '') return;
    const fst = cf ? { ...st, ...(cf.bold ? { bold: true } : {}), ...(cf.italic ? { italic: true } : {}), ...(cf.color ? { color: cf.color } : {}) } : st;
    const font = this.fontFor(fst);
    ctx.font = font;
    let color = fst.color || '#000';
    if (d.color && !fst.color) color = d.color;
    if (cell.hyperlink && !fst.color) color = '#0563c1';
    const pad = 3 * z;
    let align = st.align && st.align !== 'general' ? st.align : (d.kind === 'number' ? 'right' : d.kind === 'text' ? 'left' : 'center');
    const indent = (st.indent || 0) * 8 * z;
    let text = d.text;
    let lines = null;
    const availW = rect.w - pad * 2 - indent;
    let tw = this.measure(text, font);
    if (d.kind === 'number' && tw > availW) {
      // shrink display of numbers: fewer digits, then ####
      if (!st.numFmt || st.numFmt === 'General') { for (let w = 10; w >= 5 && tw > availW; w--) { text = (d.value < 0 ? '-' : '') + formatGeneral(Math.abs(d.value), w); tw = this.measure(text, font); } }
      if (tw > availW) { const hashW = this.measure('#', font); text = '#'.repeat(Math.max(1, Math.floor(availW / hashW))); tw = this.measure(text, font); }
    }
    if (st.wrap && d.kind === 'text') lines = this.wrapLines(text, font, availW);
    else if (text.includes('\n')) lines = text.split('\n');
    const lineH = (fst.size || 11) * 96 / 72 * z * 1.2;
    // horizontal clipping region; text may overflow into empty neighbours
    let clipX = rect.x, clipW = rect.w;
    if (!isMerge && !lines && d.kind === 'text' && tw > availW && occupied) {
      const L = this.layout;
      if (align === 'left') { let c2 = c; let x2 = rect.x + rect.w; while (x2 < this.cssW && tw + pad * 2 + indent > x2 - rect.x) { c2++; const ce = L.colMap.get(c2); if (!ce || occupied.get(cellKey(r, c2)) || s.mergeAt(r, c2) || (frozenCol && !ce.frozen)) break; x2 = ce.x + ce.w; } clipW = x2 - rect.x; }
      else if (align === 'right') { let c2 = c; let x1 = rect.x; while (x1 > L.headerW && tw + pad * 2 > rect.x + rect.w - x1) { c2--; const ce = L.colMap.get(c2); if (!ce || occupied.get(cellKey(r, c2)) || s.mergeAt(r, c2) || (!frozenCol && ce.frozen)) break; x1 = ce.x; } clipX = x1; clipW = rect.x + rect.w - x1; }
      else { let c1 = c, c2 = c, x1 = rect.x, x2 = rect.x + rect.w; let guard = 0; while (tw + pad * 2 > x2 - x1 && guard++ < 50) { const l = L.colMap.get(c1 - 1), rr = L.colMap.get(c2 + 1); let grew = false; if (l && !occupied.get(cellKey(r, c1 - 1)) && !s.mergeAt(r, c1 - 1)) { c1--; x1 = l.x; grew = true; } if (rr && !occupied.get(cellKey(r, c2 + 1)) && !s.mergeAt(r, c2 + 1)) { c2++; x2 = rr.x + rr.w; grew = true; } if (!grew) break; } clipX = x1; clipW = x2 - x1; }
    }
    ctx.save();
    ctx.beginPath(); ctx.rect(clipX, rect.y, clipW - 1, rect.h - 1); ctx.clip();
    ctx.fillStyle = color; ctx.textBaseline = 'middle';
    const valign = st.valign || 'bottom';
    const drawLine = (txt, y) => {
      let x;
      const w = this.measure(txt, font);
      if (align === 'right') { x = rect.x + rect.w - pad - w - indent; ctx.textAlign = 'left'; }
      else if (align === 'center') { x = rect.x + rect.w / 2 - w / 2; ctx.textAlign = 'left'; }
      else { x = rect.x + pad + indent; ctx.textAlign = 'left'; }
      if (align === 'right' && w > clipW - pad * 2 && clipX < rect.x) x = Math.max(clipX + pad, x);
      ctx.fillText(txt, x, y);
      if (st.underline || cell.hyperlink) { ctx.fillRect(x, y + lineH * 0.35, w, Math.max(1, z)); }
      if (st.strike) ctx.fillRect(x, y, w, Math.max(1, z));
    };
    if (lines) {
      const totalH = lines.length * lineH;
      let y0 = valign === 'top' ? rect.y + pad + lineH / 2 : valign === 'middle' ? rect.y + rect.h / 2 - totalH / 2 + lineH / 2 : rect.y + rect.h - pad - totalH + lineH / 2;
      if (totalH > rect.h) y0 = valign === 'bottom' ? rect.y + rect.h - pad - totalH + lineH / 2 : rect.y + pad + lineH / 2;
      lines.forEach((ln, i) => drawLine(ln, y0 + i * lineH));
    } else {
      const y = valign === 'top' ? rect.y + pad + lineH / 2 : valign === 'middle' ? rect.y + rect.h / 2 : rect.y + rect.h - pad - lineH / 2 + 1;
      drawLine(text, Math.max(rect.y + lineH / 2 - 1, y));
    }
    ctx.restore();
  }

  drawBorders(b, x, y, w, h, z) {
    const ctx = this.ctx;
    const side = (spec, x1, y1, x2, y2) => {
      if (!spec) return;
      const lw = Math.max(1, (BORDER_WIDTH[spec.style] || 1) * (z >= 1 ? Math.round(z) : 1));
      ctx.strokeStyle = spec.color || '#000'; ctx.lineWidth = lw;
      ctx.setLineDash(spec.style === 'dashed' || spec.style === 'mediumDashed' ? [4 * z, 2 * z] : spec.style === 'dotted' || spec.style === 'hair' ? [1.5 * z, 1.5 * z] : []);
      const off = lw % 2 ? 0.5 : 0;
      ctx.beginPath(); ctx.moveTo(x1 + off, y1 + off); ctx.lineTo(x2 + off, y2 + off); ctx.stroke();
      if (spec.style === 'double') { ctx.lineWidth = 1; ctx.beginPath(); const dx = x1 === x2 ? 2 : 0, dy = y1 === y2 ? 2 : 0; ctx.moveTo(x1 - dx + 0.5, y1 - dy + 0.5); ctx.lineTo(x2 - dx + 0.5, y2 - dy + 0.5); ctx.stroke(); }
      ctx.setLineDash([]);
    };
    const R = Math.round(x + w) - 1, B = Math.round(y + h) - 1;
    side(b.top, Math.round(x), Math.round(y) - 1, R, Math.round(y) - 1);
    side(b.bottom, Math.round(x), B, R, B);
    side(b.left, Math.round(x) - 1, Math.round(y), Math.round(x) - 1, B);
    side(b.right, R, Math.round(y), R, B);
  }

  drawFilterButton(rect, active) {
    const ctx = this.ctx, z = this.zoom;
    const sz = Math.min(rect.h - 4, 16 * z);
    const x = rect.x + rect.w - sz - 2, y = rect.y + (rect.h - sz) / 2;
    ctx.fillStyle = active ? '#1a73e8' : '#fff'; ctx.strokeStyle = active ? '#1a73e8' : '#999'; ctx.lineWidth = 1;
    ctx.fillRect(x, y, sz, sz); ctx.strokeRect(x + 0.5, y + 0.5, sz - 1, sz - 1);
    ctx.fillStyle = active ? '#fff' : '#333';
    if (active) { ctx.beginPath(); ctx.moveTo(x + sz * 0.25, y + sz * 0.3); ctx.lineTo(x + sz * 0.75, y + sz * 0.3); ctx.lineTo(x + sz * 0.55, y + sz * 0.55); ctx.lineTo(x + sz * 0.55, y + sz * 0.78); ctx.lineTo(x + sz * 0.45, y + sz * 0.78); ctx.lineTo(x + sz * 0.45, y + sz * 0.55); ctx.closePath(); ctx.fill(); }
    else { ctx.beginPath(); ctx.moveTo(x + sz * 0.25, y + sz * 0.38); ctx.lineTo(x + sz * 0.75, y + sz * 0.38); ctx.lineTo(x + sz * 0.5, y + sz * 0.68); ctx.closePath(); ctx.fill(); }
  }
  filterButtonAt(x, y) {
    const s = this.sheet; if (!s.filter) return null;
    const z = this.zoom;
    for (let c = s.filter.c1; c <= s.filter.c2; c++) { const rect = this.cellRect(s.filter.r1, c, false); if (!rect) continue; const sz = Math.min(rect.h - 4, 16 * z); const bx = rect.x + rect.w - sz - 2, by = rect.y + (rect.h - sz) / 2; if (x >= bx && x <= bx + sz && y >= by && y <= by + sz) return { c, rect: { x: bx, y: by, w: sz, h: sz } }; }
    return null;
  }

  rangeRect(rg) {
    // bounding rect of a range in canvas coordinates (may extend off screen)
    const s = this.sheet, z = this.zoom;
    const a = this.cellOrigin(rg.r1, rg.c1);
    const r2 = Math.min(rg.r2, MAX_ROWS - 1), c2 = Math.min(rg.c2, MAX_COLS - 1);
    let x2, y2;
    if (rg.c2 >= MAX_COLS - 1) x2 = this.cssW + 10; else { const b = this.cellOrigin(r2, c2); x2 = b.x + s.colAxis.size(c2) * z; }
    if (rg.r2 >= MAX_ROWS - 1) y2 = this.cssH + 10; else { const b = this.cellOrigin(r2, c2); y2 = b.y + s.rowAxis.size(r2) * z; }
    if (rg.r2 - rg.r1 > 100000 && rg.r2 < MAX_ROWS - 1) y2 = Math.min(y2, this.cssH + 10);
    if (rg.c2 - rg.c1 > 10000 && rg.c2 < MAX_COLS - 1) x2 = Math.min(x2, this.cssW + 10);
    return { x: a.x, y: a.y, w: x2 - a.x, h: y2 - a.y };
  }
  clipToPane(ctx, r, c) {
    const L = this.layout, s = this.sheet;
    const x0 = c < s.freeze.cols ? L.headerW : L.headerW + L.frozenW;
    const y0 = r < s.freeze.rows ? L.headerH : L.headerH + L.frozenH;
    const x1 = c < s.freeze.cols ? L.headerW + L.frozenW : this.cssW;
    const y1 = r < s.freeze.rows ? L.headerH + L.frozenH : this.cssH;
    ctx.beginPath(); ctx.rect(x0, y0, x1 - x0, y1 - y0); ctx.clip();
  }
  // Draw an outline for a range, split across panes so frozen areas clip correctly
  strokeRange(ctx, rg, color, width, dash = null, fill = null) {
    const s = this.sheet;
    const panesR = [[0, s.freeze.rows - 1], [Math.max(s.freeze.rows, this.scrollRow), MAX_ROWS - 1]];
    const panesC = [[0, s.freeze.cols - 1], [Math.max(s.freeze.cols, this.scrollCol), MAX_COLS - 1]];
    for (const [pr1, pr2] of panesR) for (const [pc1, pc2] of panesC) {
      if (pr1 > pr2 || pc1 > pc2) continue;
      if (rg.r2 < pr1 || rg.r1 > pr2 || rg.c2 < pc1 || rg.c1 > pc2) continue;
      ctx.save(); this.clipToPane(ctx, pr1, pc1);
      const rect = this.rangeRect(rg);
      if (fill) { ctx.fillStyle = fill; ctx.fillRect(rect.x, rect.y, rect.w, rect.h); }
      if (color) { ctx.strokeStyle = color; ctx.lineWidth = width; ctx.setLineDash(dash || []); ctx.lineDashOffset = dash ? -this._dash : 0; const o = width / 2; ctx.strokeRect(rect.x - 1 + o, rect.y - 1 + o, rect.w + 1 - width, rect.h + 1 - width); ctx.setLineDash([]); }
      ctx.restore();
    }
  }
  drawSelection(ctx, s, z) {
    const view = s.view; const act = view.active;
    const editing = this.app.editor.isEditing;
    for (const rg of view.ranges) {
      const single = rg.r1 === rg.r2 && rg.c1 === rg.c2 && !(s.mergeAt(rg.r1, rg.c1));
      if (!single) this.strokeRange(ctx, rg, null, 0, null, 'rgba(26,115,232,0.14)');
    }
    // active cell: clear highlight inside
    const am = s.mergeAt(act.r, act.c);
    const activeRange = am ? am : { r1: act.r, c1: act.c, r2: act.r, c2: act.c };
    if (view.ranges.length === 1 && !(view.ranges[0].r1 === activeRange.r1 && view.ranges[0].r2 === activeRange.r2 && view.ranges[0].c1 === activeRange.c1 && view.ranges[0].c2 === activeRange.c2)) {
      const rect = this.cellRect(act.r, act.c);
      if (rect) { ctx.save(); this.clipToPane(ctx, act.r, act.c); ctx.clearRect(rect.x, rect.y, rect.w - 1, rect.h - 1); ctx.fillStyle = '#fff'; ctx.fillRect(rect.x, rect.y, rect.w - 1, rect.h - 1); const st = s.getStyle(act.r, act.c); const cell = s.getCell(act.r, act.c); const cf = this.cfFor(s, act.r, act.c, new Map()); if ((cf && cf.fill) || st.fill) { ctx.fillStyle = (cf && cf.fill) || st.fill; ctx.fillRect(rect.x, rect.y, rect.w - 1, rect.h - 1); } if (cell && (cell.v != null || cell.f)) this.drawCellText(s, act.r, act.c, rect, st, cf, cell, null, act.r < s.freeze.rows, act.c < s.freeze.cols, !!am); if (st.border) this.drawBorders(st.border, rect.x, rect.y, rect.w, rect.h, z); ctx.restore(); }
    }
    const outer = view.ranges[view.ranges.length - 1];
    this.strokeRange(ctx, outer, '#1a73e8', 1);
    if (!editing) {
      this.strokeRange(ctx, activeRange, '#1a73e8', 2);
      // fill handle
      const rect = this.rangeRect(outer);
      ctx.save(); this.clipToPane(ctx, outer.r2, outer.c2);
      const hs = Math.max(5, Math.round(6 * z));
      ctx.fillStyle = '#1a73e8'; ctx.fillRect(rect.x + rect.w - hs / 2 - 1, rect.y + rect.h - hs / 2 - 1, hs, hs);
      ctx.strokeStyle = '#fff'; ctx.lineWidth = 1; ctx.strokeRect(rect.x + rect.w - hs / 2 - 1.5, rect.y + rect.h - hs / 2 - 1.5, hs + 1, hs + 1);
      ctx.restore();
      this.fillHandleRect = { x: rect.x + rect.w - hs - 2, y: rect.y + rect.h - hs - 2, w: hs * 2 + 2, h: hs * 2 + 2 };
    } else this.fillHandleRect = null;
    if (this.drag && this.drag.mode === 'fill' && this.drag.target) this.strokeRange(ctx, this.drag.target, '#555', 1, [4, 2]);
    if (this.copyRange && this.copyRange.sheetId === s.id) { this.strokeRange(ctx, this.copyRange, '#1a73e8', 1.5, [5, 3]); this.ensureMarquee(); } else this.stopMarquee();
    for (let i = 0; i < this.refHighlights.length; i++) { const h = this.refHighlights[i]; if (h.sheet && h.sheet.toLowerCase() !== s.name.toLowerCase()) continue; if (!h.sheet && h.sheetId && h.sheetId !== s.id) continue; const color = REF_COLORS[i % REF_COLORS.length]; this.strokeRange(ctx, h, color, 2, null, color + '22'); }
  }
  ensureMarquee() { if (this._marqueeTimer) return; this._marqueeTimer = setInterval(() => { this._dash = (this._dash + 1) % 16; this.invalidate(); }, 120); }
  stopMarquee() { if (this._marqueeTimer) { clearInterval(this._marqueeTimer); this._marqueeTimer = null; } }

  drawHeaders(ctx, s, z) {
    const L = this.layout; const view = s.view;
    const selCols = new Set(), selRows = new Set(); let allCols = false, allRows = false;
    for (const rg of view.ranges) { if (rg.r1 === 0 && rg.r2 >= MAX_ROWS - 1) allCols = true; if (rg.c1 === 0 && rg.c2 >= MAX_COLS - 1) allRows = true; }
    const inSelCol = (c) => view.ranges.some((rg) => c >= rg.c1 && c <= rg.c2);
    const inSelRow = (r) => view.ranges.some((rg) => r >= rg.r1 && r <= rg.r2);
    const fullCol = (c) => view.ranges.some((rg) => c >= rg.c1 && c <= rg.c2 && rg.r1 === 0 && rg.r2 >= MAX_ROWS - 1);
    const fullRow = (r) => view.ranges.some((rg) => r >= rg.r1 && r <= rg.r2 && rg.c1 === 0 && rg.c2 >= MAX_COLS - 1);
    const font = `${(11 * z).toFixed(1)}px "Segoe UI", system-ui, sans-serif`;
    ctx.font = font; ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
    // column headers
    ctx.fillStyle = '#f3f3f3'; ctx.fillRect(0, 0, this.cssW, L.headerH);
    for (const ce of L.cols) {
      const sel = inSelCol(ce.c), full = fullCol(ce.c);
      if (full) ctx.fillStyle = '#1a73e8'; else if (sel) ctx.fillStyle = '#d3e3fd'; else ctx.fillStyle = '#f3f3f3';
      ctx.fillRect(ce.x, 0, ce.w, L.headerH);
      ctx.fillStyle = full ? '#fff' : sel ? '#0b57d0' : '#444';
      ctx.font = sel ? `600 ${font}` : font;
      if (ce.w > 12) ctx.fillText(colToName(ce.c), ce.x + ce.w / 2, L.headerH / 2 + 0.5, ce.w - 4);
      ctx.strokeStyle = '#c8c8c8'; ctx.beginPath(); ctx.moveTo(Math.round(ce.x + ce.w) - 0.5, 0); ctx.lineTo(Math.round(ce.x + ce.w) - 0.5, L.headerH); ctx.stroke();
      if (s.colAxis.hidden.has(ce.c + 1) || (ce.c + 1 < MAX_COLS && s.colAxis.size(ce.c + 1) === 0)) { ctx.strokeStyle = '#1a73e8'; ctx.lineWidth = 2; ctx.beginPath(); ctx.moveTo(Math.round(ce.x + ce.w), 0); ctx.lineTo(Math.round(ce.x + ce.w), L.headerH); ctx.stroke(); ctx.lineWidth = 1; }
    }
    // row headers
    ctx.fillStyle = '#f3f3f3'; ctx.fillRect(0, L.headerH, L.headerW, this.cssH - L.headerH);
    for (const re of L.rows) {
      const sel = inSelRow(re.r), full = fullRow(re.r);
      if (full) ctx.fillStyle = '#1a73e8'; else if (sel) ctx.fillStyle = '#d3e3fd'; else ctx.fillStyle = '#f3f3f3';
      ctx.fillRect(0, re.y, L.headerW, re.h);
      ctx.fillStyle = full ? '#fff' : sel ? '#0b57d0' : '#444';
      ctx.font = sel ? `600 ${font}` : font;
      if (re.h > 8) ctx.fillText(String(re.r + 1), L.headerW / 2, re.y + re.h / 2 + 0.5, L.headerW - 4);
      ctx.strokeStyle = '#c8c8c8'; ctx.beginPath(); ctx.moveTo(0, Math.round(re.y + re.h) - 0.5); ctx.lineTo(L.headerW, Math.round(re.y + re.h) - 0.5); ctx.stroke();
      if (s.rowAxis.hidden.has(re.r + 1)) { ctx.strokeStyle = '#1a73e8'; ctx.lineWidth = 2; ctx.beginPath(); ctx.moveTo(0, Math.round(re.y + re.h)); ctx.lineTo(L.headerW, Math.round(re.y + re.h)); ctx.stroke(); ctx.lineWidth = 1; }
    }
    ctx.strokeStyle = '#c8c8c8'; ctx.beginPath(); ctx.moveTo(L.headerW - 0.5, 0); ctx.lineTo(L.headerW - 0.5, this.cssH); ctx.moveTo(0, L.headerH - 0.5); ctx.lineTo(this.cssW, L.headerH - 0.5); ctx.stroke();
    // corner
    ctx.fillStyle = allCols && allRows ? '#d3e3fd' : '#f3f3f3'; ctx.fillRect(0, 0, L.headerW - 1, L.headerH - 1);
    ctx.fillStyle = '#b0b0b0'; ctx.beginPath(); ctx.moveTo(L.headerW - 4, L.headerH - 14); ctx.lineTo(L.headerW - 4, L.headerH - 4); ctx.lineTo(L.headerW - 14, L.headerH - 4); ctx.closePath(); ctx.fill();
  }

  // ---------- autofit ----------
  autofitCol(c) {
    const s = this.sheet; const ub = s.usedBounds();
    let best = 0;
    for (let r = 0; r <= ub.r2; r++) {
      const cell = s.getCell(r, c); if (!cell || (cell.v == null && !cell.f)) continue;
      if (s.mergeAt(r, c)) continue;
      const st = s.getStyle(r, c); const d = this.display(s, r, c, st, cell); if (!d) continue;
      const font = this.fontFor(st, 1);
      const lines = d.text.split('\n');
      for (const ln of lines) best = Math.max(best, this.measure(ln, font) + 8 + (st.indent || 0) * 8);
    }
    return best ? Math.min(1000, Math.ceil(best)) : s.colAxis.default;
  }
  autofitRow(r) {
    const s = this.sheet; const ub = s.usedBounds();
    let best = 0;
    for (let c = 0; c <= ub.c2; c++) {
      const cell = s.getCell(r, c); if (!cell || (cell.v == null && !cell.f)) continue;
      const st = s.getStyle(r, c); const d = this.display(s, r, c, st, cell); if (!d) continue;
      const font = this.fontFor(st, 1); const lineH = (st.size || 11) * 96 / 72 * 1.2;
      let n = d.text.split('\n').length;
      if (st.wrap && d.kind === 'text') n = Math.max(n, this.wrapLines(d.text, font, s.colAxis.size(c) - 6).length);
      best = Math.max(best, n * lineH + 5);
    }
    return best ? Math.min(409 * 96 / 72, Math.ceil(best)) : s.rowAxis.default;
  }

  // ---------- mouse ----------
  pos(e) { const r = this.canvas.getBoundingClientRect(); return { x: e.clientX - r.left, y: e.clientY - r.top }; }
  onMouseDown(e) {
    if (e.button === 1) return;
    const app = this.app; const { x, y } = this.pos(e);
    app.closeMenus(); app.charts.select(null);
    if (e.button === 2) return; // context menu handled separately
    const hit = this.cellAt(x, y);
    const s = this.sheet;
    e.preventDefault();
    // dropdown arrow on active cell (data validation list)
    if (this.dvButton && x >= this.dvButton.x && x <= this.dvButton.x + this.dvButton.w && y >= this.dvButton.y && y <= this.dvButton.y + this.dvButton.h) { app.dialogs.validationList(s.view.active.r, s.view.active.c); return; }
    const fb = this.filterButtonAt(x, y);
    if (fb) { app.dialogs.filterPopover(fb.c, fb.rect); return; }
    if (app.editor.isEditing && app.editor.isFormula() && hit.type === 'cell') {
      // reference picking
      if (app.editor.canInsertRef()) { this.drag = { mode: 'refpick', start: { r: hit.r, c: hit.c } }; app.editor.insertRef({ r1: hit.r, c1: hit.c, r2: hit.r, c2: hit.c }, false); return; }
    }
    if (app.editor.isEditing) { if (!app.editor.commit()) return; }
    if (this.fillHandleRect && x >= this.fillHandleRect.x && x <= this.fillHandleRect.x + this.fillHandleRect.w && y >= this.fillHandleRect.y && y <= this.fillHandleRect.y + this.fillHandleRect.h) {
      this.drag = { mode: 'fill', src: { ...s.view.ranges[s.view.ranges.length - 1] }, target: null }; this.canvas.style.cursor = 'crosshair'; return;
    }
    switch (hit.type) {
      case 'corner': app.selectAll(); break;
      case 'colResize': this.drag = { mode: 'colResize', c: hit.c, startX: x, startW: s.colAxis.size(hit.c) }; break;
      case 'rowResize': this.drag = { mode: 'rowResize', r: hit.r, startY: y, startH: s.rowAxis.size(hit.r) }; break;
      case 'colHeader': {
        if (e.shiftKey) app.selectRange({ r1: 0, c1: Math.min(s.view.anchor ? s.view.anchor.c : s.view.active.c, hit.c), r2: MAX_ROWS - 1, c2: Math.max(s.view.anchor ? s.view.anchor.c : s.view.active.c, hit.c) }, { keepActive: true });
        else app.selectRange({ r1: 0, c1: hit.c, r2: MAX_ROWS - 1, c2: hit.c }, { add: e.ctrlKey || e.metaKey, active: { r: this.scrollRow, c: hit.c } });
        this.drag = { mode: 'colSelect', startC: hit.c, add: e.ctrlKey || e.metaKey }; break;
      }
      case 'rowHeader': {
        if (e.shiftKey) app.selectRange({ r1: Math.min(s.view.anchor ? s.view.anchor.r : s.view.active.r, hit.r), c1: 0, r2: Math.max(s.view.anchor ? s.view.anchor.r : s.view.active.r, hit.r), c2: MAX_COLS - 1 }, { keepActive: true });
        else app.selectRange({ r1: hit.r, c1: 0, r2: hit.r, c2: MAX_COLS - 1 }, { add: e.ctrlKey || e.metaKey, active: { r: hit.r, c: this.scrollCol } });
        this.drag = { mode: 'rowSelect', startR: hit.r, add: e.ctrlKey || e.metaKey }; break;
      }
      case 'cell': {
        if (e.shiftKey) app.extendSelectionTo(hit.r, hit.c);
        else app.select(hit.r, hit.c, { add: e.ctrlKey || e.metaKey });
        this.drag = { mode: 'select', anchor: e.shiftKey ? (s.view.anchor || s.view.active) : { r: hit.r, c: hit.c }, add: e.ctrlKey || e.metaKey };
        const cell = s.getCell(hit.r, hit.c);
        if (cell && cell.hyperlink && (e.ctrlKey || e.metaKey)) app.openHyperlink(cell.hyperlink);
        break;
      }
    }
  }
  onDragMove(e) {
    const d = this.drag; if (!d) return;
    const { x, y } = this.pos(e); const s = this.sheet; const app = this.app; const z = this.zoom;
    if (d.mode === 'colResize') { const w = Math.max(0, d.startW + (x - d.startX) / z); this.previewColW = { c: d.c, w }; s.colAxis.setSize(d.c, Math.round(w)); this.invalidate(); return; }
    if (d.mode === 'rowResize') { const h = Math.max(0, d.startH + (y - d.startY) / z); this.previewRowH = { r: d.r, h }; s.rowAxis.setSize(d.r, Math.round(h)); this.invalidate(); return; }
    const hit = this.cellAt(clamp(x, this.layout.headerW, this.cssW - 1), clamp(y, this.layout.headerH, this.cssH - 1));
    this.autoScroll(x, y);
    if (d.mode === 'select' && hit.r != null) app.extendSelectionTo(hit.r, hit.c, d.anchor);
    else if (d.mode === 'colSelect' && hit.c != null) app.selectRange({ r1: 0, c1: Math.min(d.startC, hit.c), r2: MAX_ROWS - 1, c2: Math.max(d.startC, hit.c) }, { replaceLast: true, keepActive: true });
    else if (d.mode === 'rowSelect' && hit.r != null) app.selectRange({ r1: Math.min(d.startR, hit.r), c1: 0, r2: Math.max(d.startR, hit.r), c2: MAX_COLS - 1 }, { replaceLast: true, keepActive: true });
    else if (d.mode === 'fill' && hit.r != null) {
      const src = d.src; let t;
      const dr = hit.r < src.r1 ? hit.r - src.r1 : hit.r > src.r2 ? hit.r - src.r2 : 0;
      const dc = hit.c < src.c1 ? hit.c - src.c1 : hit.c > src.c2 ? hit.c - src.c2 : 0;
      if (Math.abs(dr) >= Math.abs(dc)) t = dr === 0 ? null : { r1: Math.min(src.r1, hit.r), r2: Math.max(src.r2, hit.r), c1: src.c1, c2: src.c2 };
      else t = { r1: src.r1, r2: src.r2, c1: Math.min(src.c1, hit.c), c2: Math.max(src.c2, hit.c) };
      d.target = t; this.invalidate();
    } else if (d.mode === 'refpick' && hit.r != null) { app.editor.insertRef(normRange(d.start.r, d.start.c, hit.r, hit.c), true); }
  }
  onDragEnd(e) {
    const d = this.drag; if (!d) return; this.drag = null; this.stopAutoScroll(); this.canvas.style.cursor = 'cell';
    const s = this.sheet, app = this.app;
    if (d.mode === 'colResize') { const w = s.colAxis.size(d.c); s.colAxis.setSize(d.c, d.startW); const cols = app.selectedWholeCols().includes(d.c) ? app.selectedWholeCols() : [d.c]; app.transact('Column width', () => cols.forEach((c) => s.setColWidth(c, w))); }
    else if (d.mode === 'rowResize') { const h = s.rowAxis.size(d.r); s.rowAxis.setSize(d.r, d.startH); const rows = app.selectedWholeRows().includes(d.r) ? app.selectedWholeRows() : [d.r]; app.transact('Row height', () => rows.forEach((r) => s.setRowHeight(r, h))); }
    else if (d.mode === 'fill' && d.target) { app.autofill(d.src, d.target, { increment: e.ctrlKey }); }
    else if (d.mode === 'refpick') { app.editor.focus(); }
    this.invalidate();
  }
  autoScroll(x, y) {
    const L = this.layout;
    let dr = 0, dc = 0;
    if (y > this.cssH - 10) dr = 1; else if (y < L.headerH + L.frozenH && y < this.lastY) dr = -1;
    if (x > this.cssW - 10) dc = 1; else if (x < L.headerW + L.frozenW && x < this.lastX) dc = -1;
    this.lastX = x; this.lastY = y;
    if (dr || dc) { if (!this._autoTimer) this._autoTimer = setInterval(() => { this.scrollBy(this._autoDr, this._autoDc); if (this.drag) this.onDragMove({ clientX: this._lastClient.x, clientY: this._lastClient.y }); }, 80); this._autoDr = dr; this._autoDc = dc; }
    else this.stopAutoScroll();
    const r = this.canvas.getBoundingClientRect(); this._lastClient = { x: x + r.left, y: y + r.top };
  }
  stopAutoScroll() { if (this._autoTimer) { clearInterval(this._autoTimer); this._autoTimer = null; } }
  onDblClick(e) {
    const { x, y } = this.pos(e); const hit = this.cellAt(x, y); const app = this.app; const s = this.sheet;
    if (hit.type === 'colResize') { const cols = app.selectedWholeCols().includes(hit.c) ? app.selectedWholeCols() : [hit.c]; app.transact('Autofit column', () => cols.forEach((c) => s.setColWidth(c, this.autofitCol(c)))); return; }
    if (hit.type === 'rowResize') { const rows = app.selectedWholeRows().includes(hit.r) ? app.selectedWholeRows() : [hit.r]; app.transact('Autofit row', () => rows.forEach((r) => s.setRowHeight(r, this.autofitRow(r)))); return; }
    if (hit.type === 'cell') { if (this.filterButtonAt(x, y)) return; app.editor.start('edit'); }
  }
  onContextMenu(e) {
    e.preventDefault();
    const { x, y } = this.pos(e); const hit = this.cellAt(x, y); const app = this.app; const s = this.sheet;
    if (app.editor.isEditing) app.editor.commit();
    if (hit.type === 'cell') { if (!s.view.ranges.some((rg) => rangeContains(rg, hit.r, hit.c))) app.select(hit.r, hit.c); app.showContextMenu(app.cellContextItems(), e.clientX, e.clientY); }
    else if (hit.type === 'colHeader' || hit.type === 'colResize') { if (!app.selectedWholeCols().includes(hit.c)) app.selectRange({ r1: 0, c1: hit.c, r2: MAX_ROWS - 1, c2: hit.c }, { active: { r: this.scrollRow, c: hit.c } }); app.showContextMenu(app.colContextItems(), e.clientX, e.clientY); }
    else if (hit.type === 'rowHeader' || hit.type === 'rowResize') { if (!app.selectedWholeRows().includes(hit.r)) app.selectRange({ r1: hit.r, c1: 0, r2: hit.r, c2: MAX_COLS - 1 }, { active: { r: hit.r, c: this.scrollCol } }); app.showContextMenu(app.rowContextItems(), e.clientX, e.clientY); }
  }
  onHover(e) {
    if (this.drag) return;
    const { x, y } = this.pos(e); const hit = this.cellAt(x, y);
    let cursor = 'cell';
    if (hit.type === 'colResize') cursor = 'col-resize'; else if (hit.type === 'rowResize') cursor = 'row-resize';
    else if (hit.type === 'colHeader' || hit.type === 'rowHeader') cursor = 'default';
    else if (this.fillHandleRect && x >= this.fillHandleRect.x && x <= this.fillHandleRect.x + this.fillHandleRect.w && y >= this.fillHandleRect.y && y <= this.fillHandleRect.y + this.fillHandleRect.h) cursor = 'crosshair';
    else if (this.filterButtonAt(x, y) || (this.dvButton && x >= this.dvButton.x && x <= this.dvButton.x + this.dvButton.w && y >= this.dvButton.y && y <= this.dvButton.y + this.dvButton.h)) cursor = 'pointer';
    else if (hit.type === 'cell') { const cell = this.sheet.getCell(hit.r, hit.c); if (cell && cell.hyperlink && (e.ctrlKey || e.metaKey)) cursor = 'pointer'; }
    this.canvas.style.cursor = cursor;
    this.setHover(hit.type === 'cell' ? hit : null, e);
  }
  setHover(hit, e) {
    const key = hit ? `${hit.r},${hit.c}` : null;
    if (key === this.hoverKey) return;
    this.hoverKey = key;
    this.app.onHoverCell(hit, e);
  }
}
function keyOf(v) { return typeof v === 'string' ? 's:' + v.toLowerCase() : typeof v === 'number' ? 'n:' + v : typeof v === 'boolean' ? 'b:' + v : 'o:' + String(v); }
