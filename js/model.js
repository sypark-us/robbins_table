// Workbook model: sheets, cells, styles, undo/redo, structural operations.
import { Engine } from './engine.js';
import { parseInput, isDateFormat } from './format.js';
import { adjustFormulaForInsertDelete, adjustFormulaForMove, shiftFormula, renameSheetInFormula, removeSheetInFormula, isFormula, rewriteRefs } from './formula.js';
import { key as cellKey, unkey, normRange, rangesIntersect, compareValues, MAX_ROWS, MAX_COLS, uid, deepClone, parseRangeName } from './util.js';

export const DEFAULT_COL_WIDTH = 80;
export const DEFAULT_ROW_HEIGHT = 20;
export const DEFAULT_STYLE = { font: 'Calibri', size: 11 };

// ---------------- Axis (row heights / column widths with hidden entries) ----------------
export class Axis {
  constructor(defaultSize) {
    this.default = defaultSize;
    this.sizes = new Map();
    this.hidden = new Set();
    this.filtered = new Set();
    this._specials = null; // sorted [{i, delta}] with cumulative deltas
  }
  invalidate() { this._specials = null; }
  size(i) { if (this.hidden.has(i) || this.filtered.has(i)) return 0; return this.sizes.get(i) ?? this.default; }
  rawSize(i) { return this.sizes.get(i) ?? this.default; }
  isHidden(i) { return this.hidden.has(i) || this.filtered.has(i); }
  setSize(i, s) { if (s === this.default) this.sizes.delete(i); else this.sizes.set(i, s); this.invalidate(); }
  _build() {
    const idx = new Set([...this.sizes.keys(), ...this.hidden, ...this.filtered]);
    const arr = [...idx].sort((a, b) => a - b);
    let cum = 0;
    this._specials = arr.map((i) => { cum += this.size(i) - this.default; return { i, cum }; });
  }
  // sum of sizes of entries [0, i)
  offset(i) {
    if (!this._specials) this._build();
    const sp = this._specials;
    let lo = 0, hi = sp.length; // find count of specials with index < i
    while (lo < hi) { const mid = (lo + hi) >> 1; if (sp[mid].i < i) lo = mid + 1; else hi = mid; }
    return i * this.default + (lo > 0 ? sp[lo - 1].cum : 0);
  }
  // index whose span contains pos (pos >= 0)
  indexAt(pos, max) {
    let lo = 0, hi = max;
    while (lo < hi) { const mid = (lo + hi + 1) >> 1; if (this.offset(mid) <= pos) lo = mid; else hi = mid - 1; }
    // skip hidden entries at that position (zero size) forward
    while (lo < max && this.size(lo) === 0 && this.offset(lo) <= pos) lo++;
    return lo;
  }
  // shift entries at or after index by count (insert) or remove [index, index-count) (delete)
  shift(index, count) {
    const remap = (map) => { const out = new Map(); for (const [i, v] of map) { if (i < index) out.set(i, v); else if (count > 0) out.set(i + count, v); else if (i >= index - count) out.set(i + count, v); } return out; };
    const remapSet = (set) => { const out = new Set(); for (const i of set) { if (i < index) out.add(i); else if (count > 0) out.add(i + count); else if (i >= index - count) out.add(i + count); } return out; };
    this.sizes = remap(this.sizes); this.hidden = remapSet(this.hidden); this.filtered = remapSet(this.filtered);
    this.invalidate();
  }
  clone() { const a = new Axis(this.default); a.sizes = new Map(this.sizes); a.hidden = new Set(this.hidden); a.filtered = new Set(this.filtered); return a; }
}

// ---------------- Sheet ----------------
export class Sheet {
  constructor(wb, name, id) {
    this.wb = wb; this.name = name; this.id = id || uid('s');
    this.cells = new Map();
    this.rowAxis = new Axis(DEFAULT_ROW_HEIGHT);
    this.colAxis = new Axis(DEFAULT_COL_WIDTH);
    this.colStyles = new Map(); this.rowStyles = new Map();
    this.merges = [];
    this.freeze = { rows: 0, cols: 0 };
    this.cf = [];
    this.validations = [];
    this.charts = [];
    this.filter = null;
    this.tabColor = null;
    this.showGridlines = true;
    this.view = { active: { r: 0, c: 0 }, ranges: [{ r1: 0, c1: 0, r2: 0, c2: 0 }], scrollRow: 0, scrollCol: 0, zoom: 1 };
    this._bounds = null;
  }
  index() { return this.wb.sheets.indexOf(this); }
  getCell(r, c) { return this.cells.get(cellKey(r, c)) || null; }
  getValue(r, c) {
    const cell = this.cells.get(cellKey(r, c));
    if (!cell) return null;
    if (cell.f && this.wb.engine.isDirty(this, r, c)) return this.wb.engine.evalCell(this, r, c);
    if (cell.spillOf) { const [or, oc] = unkey(cell.spillOf); if (this.wb.engine.isDirty(this, or, oc)) this.wb.engine.evalCell(this, or, oc); }
    return cell.v === undefined ? null : cell.v;
  }
  getStyleId(r, c) {
    const cell = this.cells.get(cellKey(r, c));
    if (cell && cell.s) return cell.s;
    return this.rowStyles.get(r) || this.colStyles.get(c) || 0;
  }
  getStyle(r, c) { return this.wb.styles[this.getStyleId(r, c)] || DEFAULT_STYLE; }
  effectiveNumFmt(r, c) {
    const st = this.getStyle(r, c);
    if (st.numFmt && st.numFmt !== 'General') return st.numFmt;
    const cell = this.cells.get(cellKey(r, c));
    if (cell && cell.autoFmt) return cell.autoFmt;
    return 'General';
  }
  colWidth(c) { return this.colAxis.size(c); }
  rowHeight(r) { return this.rowAxis.size(r); }
  isRowHidden(r) { return this.rowAxis.hidden.has(r); }
  isRowFiltered(r) { return this.rowAxis.filtered.has(r); }
  invalidateBounds() { this._bounds = null; }
  usedBounds() {
    if (this._bounds) return this._bounds;
    let r2 = -1, c2 = -1;
    for (const [k, cell] of this.cells) { if (cell.v == null && !cell.f && !cell.s) continue; const [r, c] = unkey(k); if (r > r2) r2 = r; if (c > c2) c2 = c; }
    this._bounds = { r1: 0, c1: 0, r2: Math.max(0, r2), c2: Math.max(0, c2), empty: r2 < 0 };
    return this._bounds;
  }
  // bounds of non-empty content (values/formulas only)
  dataBounds() {
    let r1 = Infinity, c1 = Infinity, r2 = -1, c2 = -1;
    for (const [k, cell] of this.cells) { if (cell.v == null && !cell.f) continue; const [r, c] = unkey(k); if (r < r1) r1 = r; if (c < c1) c1 = c; if (r > r2) r2 = r; if (c > c2) c2 = c; }
    if (r2 < 0) return null;
    return { r1, c1, r2, c2 };
  }
  mergeAt(r, c) { for (const m of this.merges) if (r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2) return m; return null; }

  // ---- low-level mutation (records undo) ----
  _setCellRaw(r, c, cell) {
    const k = cellKey(r, c);
    const old = this.cells.get(k) || null;
    this.wb._recordCell(this, k, old);
    if (old && old.spillOf) { const [or, oc] = unkey(old.spillOf); this.wb.engine.markDirty(this, or, oc); }
    if (old && old.spill) this.wb.engine.clearSpill(this, old, r, c);
    if (cell && (cell.v != null || cell.f || cell.s || cell.note || cell.hyperlink)) {
      const nc = { v: cell.v === undefined ? null : cell.v, f: cell.f || null, s: cell.s || 0 };
      if (cell.note) nc.note = cell.note;
      if (cell.hyperlink) nc.hyperlink = cell.hyperlink;
      this.cells.set(k, nc);
    } else this.cells.delete(k);
    this._bounds = null;
    this.wb.engine.cellChanged(this, r, c);
    this.wb._touch();
  }

  // ---- public editing API (call inside wb.transact) ----
  setInput(r, c, text) {
    // text typed by the user: formula, number, date, boolean or text
    const old = this.getCell(r, c);
    const s = old ? old.s : this.getStyleId(r, c);
    if (text == null || text === '') { if (old && (old.v != null || old.f)) this._setCellRaw(r, c, { v: null, f: null, s, note: old && old.note }); return; }
    if (isFormula(text)) { this._setCellRaw(r, c, { v: null, f: text, s, note: old && old.note }); return; }
    const style = this.wb.styles[s] || DEFAULT_STYLE;
    if (style.numFmt === '@') { this._setCellRaw(r, c, { v: text, f: null, s, note: old && old.note }); return; }
    const p = parseInput(text);
    let sid = s;
    if (p.format && (!style.numFmt || style.numFmt === 'General')) sid = this.wb.styleId({ ...style, numFmt: p.format });
    this._setCellRaw(r, c, { v: p.value, f: null, s: sid, note: old && old.note });
  }
  setValue(r, c, value) {
    const old = this.getCell(r, c);
    this._setCellRaw(r, c, { v: value, f: null, s: old ? old.s : this.getStyleId(r, c), note: old && old.note });
  }
  setFormula(r, c, f) {
    const old = this.getCell(r, c);
    this._setCellRaw(r, c, { v: null, f: f.startsWith('=') ? f : '=' + f, s: old ? old.s : this.getStyleId(r, c), note: old && old.note });
  }
  setNote(r, c, note) { const old = this.getCell(r, c) || { v: null, f: null, s: this.getStyleId(r, c) }; this._setCellRaw(r, c, { ...old, note: note || null }); }
  setHyperlink(r, c, url) { const old = this.getCell(r, c) || { v: null, f: null, s: this.getStyleId(r, c) }; this._setCellRaw(r, c, { ...old, hyperlink: url || null }); }
  clearRange(rg, what = 'contents') {
    for (let r = rg.r1; r <= Math.min(rg.r2, this.usedBounds().r2); r++) for (let c = rg.c1; c <= Math.min(rg.c2, this.usedBounds().c2); c++) {
      const cell = this.getCell(r, c);
      if (!cell) continue;
      if (what === 'all') this._setCellRaw(r, c, null);
      else if (what === 'contents') { if (cell.v != null || cell.f || cell.spillOf) this._setCellRaw(r, c, { v: null, f: null, s: cell.s, note: cell.note }); }
      else if (what === 'formats') this._setCellRaw(r, c, { ...cell, s: 0 });
      else if (what === 'notes') this._setCellRaw(r, c, { ...cell, note: null });
    }
    if (what === 'all' || what === 'formats') {
      const whole = rg.r1 === 0 && rg.r2 >= MAX_ROWS - 1, wholeR = rg.c1 === 0 && rg.c2 >= MAX_COLS - 1;
      if (whole) for (let c = rg.c1; c <= rg.c2 && c < MAX_COLS; c++) if (this.colStyles.has(c)) this.setColStyle(c, 0);
      if (wholeR) for (let r = rg.r1; r <= rg.r2 && r < MAX_ROWS; r++) if (this.rowStyles.has(r)) this.setRowStyle(r, 0);
    }
  }
  setColStyle(c, sid) { const old = this.colStyles.get(c) || 0; this.wb._record(() => this.setColStyle(c, old)); if (sid) this.colStyles.set(c, sid); else this.colStyles.delete(c); this.wb._touch(); }
  setRowStyle(r, sid) { const old = this.rowStyles.get(r) || 0; this.wb._record(() => this.setRowStyle(r, old)); if (sid) this.rowStyles.set(r, sid); else this.rowStyles.delete(r); this.wb._touch(); }

  // Apply a style patch to a range. patch: partial style object; values of null remove the property.
  applyStyle(rg, patch) {
    const wholeCols = rg.r1 === 0 && rg.r2 >= MAX_ROWS - 1;
    const wholeRows = rg.c1 === 0 && rg.c2 >= MAX_COLS - 1;
    const ub = this.usedBounds();
    if (wholeCols) {
      for (let c = rg.c1; c <= rg.c2; c++) this.setColStyle(c, this.wb.styleId(mergeStyle(this.wb.styles[this.colStyles.get(c) || 0], patch)));
      for (const [k, cell] of [...this.cells]) { const [r, c] = unkey(k); if (c >= rg.c1 && c <= rg.c2 && cell.s) this._setCellRaw(r, c, { ...cell, s: this.wb.styleId(mergeStyle(this.wb.styles[cell.s], patch)) }); }
      for (const [r, sid] of [...this.rowStyles]) { for (let c = rg.c1; c <= rg.c2; c++) if (!this.getCell(r, c)) this._setCellRaw(r, c, { v: null, f: null, s: this.wb.styleId(mergeStyle(this.wb.styles[sid], patch)) }); }
      return;
    }
    if (wholeRows) {
      for (let r = rg.r1; r <= rg.r2; r++) this.setRowStyle(r, this.wb.styleId(mergeStyle(this.wb.styles[this.rowStyles.get(r) || 0], patch)));
      for (const [k, cell] of [...this.cells]) { const [r, c] = unkey(k); if (r >= rg.r1 && r <= rg.r2 && cell.s) this._setCellRaw(r, c, { ...cell, s: this.wb.styleId(mergeStyle(this.wb.styles[cell.s], patch)) }); }
      for (const [c, sid] of [...this.colStyles]) { for (let r = rg.r1; r <= rg.r2; r++) if (!this.getCell(r, c)) this._setCellRaw(r, c, { v: null, f: null, s: this.wb.styleId(mergeStyle(this.wb.styles[sid], patch)) }); }
      return;
    }
    const r2 = Math.min(rg.r2, MAX_ROWS - 1), c2 = Math.min(rg.c2, MAX_COLS - 1);
    if ((r2 - rg.r1 + 1) * (c2 - rg.c1 + 1) > 2_000_000) throw new Error('Range too large to format');
    const cache = new Map();
    for (let r = rg.r1; r <= r2; r++) for (let c = rg.c1; c <= c2; c++) {
      const cell = this.getCell(r, c);
      const base = this.getStyleId(r, c);
      let sid = cache.get(base); if (sid === undefined) { sid = this.wb.styleId(mergeStyle(this.wb.styles[base], patch)); cache.set(base, sid); }
      if (cell ? cell.s === sid : sid === base) continue;
      this._setCellRaw(r, c, cell ? { ...cell, s: sid } : { v: null, f: null, s: sid });
    }
  }

  setColWidth(c, w) { const old = this.colAxis.rawSize(c); this.wb._record(() => this.setColWidth(c, old)); this.colAxis.setSize(c, Math.max(0, Math.round(w))); this.wb._touch(); }
  setRowHeight(r, h) { const old = this.rowAxis.rawSize(r); this.wb._record(() => this.setRowHeight(r, old)); this.rowAxis.setSize(r, Math.max(0, Math.round(h))); this.wb._touch(); }
  setColHidden(c, hidden) { const was = this.colAxis.hidden.has(c); if (was === hidden) return; this.wb._record(() => this.setColHidden(c, was)); if (hidden) this.colAxis.hidden.add(c); else this.colAxis.hidden.delete(c); this.colAxis.invalidate(); this.wb._touch(); }
  setRowHidden(r, hidden) { const was = this.rowAxis.hidden.has(r); if (was === hidden) return; this.wb._record(() => this.setRowHidden(r, was)); if (hidden) this.rowAxis.hidden.add(r); else this.rowAxis.hidden.delete(r); this.rowAxis.invalidate(); this.wb._touch(); }
  setFreeze(rows, cols) { const old = { ...this.freeze }; this.wb._record(() => this.setFreeze(old.rows, old.cols)); this.freeze = { rows, cols }; this.wb._touch(); }
  setMerges(merges) { const old = this.merges; this.wb._record(() => this.setMerges(old)); this.merges = merges; this.wb._touch(); }
  merge(rg) {
    rg = normRange(rg.r1, rg.c1, rg.r2, rg.c2);
    if (rg.r1 === rg.r2 && rg.c1 === rg.c2) return;
    const keep = this.merges.filter((m) => !rangesIntersect(m, rg));
    // keep only top-left content
    for (let r = rg.r1; r <= Math.min(rg.r2, this.usedBounds().r2); r++) for (let c = rg.c1; c <= Math.min(rg.c2, this.usedBounds().c2); c++) { if (r === rg.r1 && c === rg.c1) continue; const cell = this.getCell(r, c); if (cell && (cell.v != null || cell.f)) this._setCellRaw(r, c, { v: null, f: null, s: cell.s }); }
    this.setMerges([...keep, { r1: rg.r1, c1: rg.c1, r2: rg.r2, c2: rg.c2 }]);
  }
  unmerge(rg) { this.setMerges(this.merges.filter((m) => !rangesIntersect(m, rg))); }
  setCF(rules) { const old = this.cf; this.wb._record(() => this.setCF(old)); this.cf = rules; this.wb._touch(); }
  setValidations(v) { const old = this.validations; this.wb._record(() => this.setValidations(old)); this.validations = v; this.wb._touch(); }
  setCharts(ch) { const old = this.charts; this.wb._record(() => this.setCharts(old)); this.charts = ch; this.wb._touch(); }
  updateChart(id, patch) { this.setCharts(this.charts.map((c) => c.id === id ? { ...c, ...patch } : c)); }
  setFilter(f) { const old = this.filter; this.wb._record(() => this.setFilter(old)); this.filter = f; this.applyFilter(); this.wb._touch(); }
  setGridlines(on) { const old = this.showGridlines; this.wb._record(() => this.setGridlines(old)); this.showGridlines = on; this.wb._touch(); }
  setTabColor(col) { const old = this.tabColor; this.wb._record(() => this.setTabColor(old)); this.tabColor = col; this.wb._touch(); }
  validationAt(r, c) { for (const v of this.validations) if (r >= v.r1 && r <= v.r2 && c >= v.c1 && c <= v.c2) return v; return null; }

  // recompute filtered rows from this.filter
  applyFilter() {
    const f = this.filter;
    const filtered = new Set();
    if (f) {
      const r2 = Math.min(f.r2, this.usedBounds().r2);
      for (let r = f.r1 + 1; r <= r2; r++) {
        let hide = false;
        for (const [colStr, cond] of Object.entries(f.conds || {})) {
          const c = +colStr;
          const v = this.getValue(r, c);
          const text = this.wb.displayText(this, r, c);
          if (cond.values && !cond.values.includes(text)) { hide = true; break; }
          if (cond.text) {
            const hay = String(text).toLowerCase(), needle = String(cond.text.v).toLowerCase();
            const ok = { contains: hay.includes(needle), notContains: !hay.includes(needle), equals: hay === needle, notEquals: hay !== needle, begins: hay.startsWith(needle), ends: hay.endsWith(needle) }[cond.text.op];
            if (!ok) { hide = true; break; }
          }
          if (cond.num) {
            const n = typeof v === 'number' ? v : NaN;
            const { op, v1, v2 } = cond.num;
            const ok = !isNaN(n) && ({ gt: n > v1, ge: n >= v1, lt: n < v1, le: n <= v1, eq: n === v1, ne: n !== v1, between: n >= Math.min(v1, v2) && n <= Math.max(v1, v2) }[op]);
            if (!ok) { hide = true; break; }
          }
          if (cond.blank === 'blank' && !(v == null || v === '')) { hide = true; break; }
          if (cond.blank === 'nonblank' && (v == null || v === '')) { hide = true; break; }
        }
        if (hide) filtered.add(r);
      }
    }
    this.rowAxis.filtered = filtered; this.rowAxis.invalidate();
  }

  // ---- structural operations ----
  insertRows(at, n) { this.wb._structural(this, 'row', at, n); }
  deleteRows(at, n) { this.wb._structural(this, 'row', at, -n); }
  insertCols(at, n) { this.wb._structural(this, 'col', at, n); }
  deleteCols(at, n) { this.wb._structural(this, 'col', at, -n); }

  // Move a block of cells (cut/paste). Formulas referring to the block follow it.
  moveRange(src, dr, dc, dstSheet = this) {
    const cells = [];
    for (let r = src.r1; r <= src.r2; r++) for (let c = src.c1; c <= src.c2; c++) { const cell = this.getCell(r, c); cells.push({ r, c, cell: cell ? { ...cell } : null }); }
    const dst = { r1: src.r1 + dr, c1: src.c1 + dc, r2: src.r2 + dr, c2: src.c2 + dc };
    // rewrite all formulas in the workbook
    for (const sh of this.wb.sheets) sh._rewriteFormulas((f) => adjustFormulaForMove(f, sh.name, this.name, src, dr, dc, dstSheet !== this ? dstSheet.name : null));
    for (const { r, c } of cells) this._setCellRaw(r, c, null);
    for (let r = dst.r1; r <= dst.r2; r++) for (let c = dst.c1; c <= dst.c2; c++) if (!(dstSheet === this && r >= src.r1 && r <= src.r2 && c >= src.c1 && c <= src.c2)) dstSheet._setCellRaw(r, c, null);
    for (const { r, c, cell } of cells) if (cell) dstSheet._setCellRaw(r + dr, c + dc, cell);
    // merges inside the block move too
    const moved = this.merges.filter((m) => m.r1 >= src.r1 && m.r2 <= src.r2 && m.c1 >= src.c1 && m.c2 <= src.c2);
    if (moved.length) {
      this.setMerges(this.merges.filter((m) => !moved.includes(m)));
      dstSheet.setMerges([...dstSheet.merges.filter((m) => !rangesIntersect(m, dst)), ...moved.map((m) => ({ r1: m.r1 + dr, c1: m.c1 + dc, r2: m.r2 + dr, c2: m.c2 + dc }))]);
    }
    this.wb._needsRebuild = true;
  }

  _rewriteFormulas(fn) {
    for (const [k, cell] of [...this.cells]) {
      if (!cell.f) continue;
      const nf = fn(cell.f);
      if (nf !== cell.f) { const [r, c] = unkey(k); this._setCellRaw(r, c, { ...cell, f: nf }); }
    }
  }

  // Sort rows of a range by keys [{col, asc}]. Cells (with formats) move as rows.
  sortRange(rg, keys, hasHeader = false) {
    const r1 = rg.r1 + (hasHeader ? 1 : 0);
    const rows = [];
    for (let r = r1; r <= rg.r2; r++) {
      const row = [];
      for (let c = rg.c1; c <= rg.c2; c++) { const cell = this.getCell(r, c); row.push(cell ? { ...cell, v: cell.f ? null : cell.v } : null); }
      rows.push({ r, row, vals: keys.map((k) => this.getValue(r, k.col)) });
    }
    const idx = rows.map((_, i) => i);
    idx.sort((a, b) => { for (let k = 0; k < keys.length; k++) { let cmp = compareValues(rows[a].vals[k], rows[b].vals[k]); if (!keys[k].asc) { const ea = rows[a].vals[k] == null, eb = rows[b].vals[k] == null; if (ea || eb) cmp = ea && eb ? 0 : ea ? 1 : -1; else cmp = -cmp; } if (cmp) return cmp; } return a - b; });
    idx.forEach((srcIdx, i) => {
      const target = r1 + i; const src = rows[srcIdx];
      for (let j = 0; j < src.row.length; j++) {
        const cell = src.row[j]; const c = rg.c1 + j;
        if (!cell) { if (this.getCell(target, c)) this._setCellRaw(target, c, null); continue; }
        const nc = { ...cell };
        if (nc.f) nc.f = shiftFormula(nc.f, target - src.r, 0);
        this._setCellRaw(target, c, nc);
      }
    });
    this.wb._needsRebuild = true;
  }

  toJSON() {
    const cells = [];
    for (const [k, cell] of this.cells) {
      if (cell.spillOf && !cell.f && !cell.s && !cell.note) continue;
      const [r, c] = unkey(k);
      const e = [r, c, cell.f ? null : serializeValue(cell.v), cell.f || null, cell.s || 0];
      if (cell.note) e.push(cell.note); if (cell.hyperlink) { while (e.length < 6) e.push(null); e.push(cell.hyperlink); }
      cells.push(e);
    }
    return { name: this.name, id: this.id, cells, colWidths: [...this.colAxis.sizes], rowHeights: [...this.rowAxis.sizes], hiddenCols: [...this.colAxis.hidden], hiddenRows: [...this.rowAxis.hidden], colStyles: [...this.colStyles], rowStyles: [...this.rowStyles], merges: this.merges, freeze: this.freeze, cf: this.cf, validations: this.validations, charts: this.charts, filter: this.filter, tabColor: this.tabColor, showGridlines: this.showGridlines, view: this.view };
  }
  static fromJSON(wb, j) {
    const s = new Sheet(wb, j.name, j.id);
    for (const e of j.cells || []) { const cell = { v: e[3] ? null : deserializeValue(e[2]), f: e[3] || null, s: e[4] || 0 }; if (e[5]) cell.note = e[5]; if (e[6]) cell.hyperlink = e[6]; s.cells.set(cellKey(e[0], e[1]), cell); }
    s.colAxis.sizes = new Map(j.colWidths || []); s.rowAxis.sizes = new Map(j.rowHeights || []);
    s.colAxis.hidden = new Set(j.hiddenCols || []); s.rowAxis.hidden = new Set(j.hiddenRows || []);
    s.colStyles = new Map(j.colStyles || []); s.rowStyles = new Map(j.rowStyles || []);
    s.merges = j.merges || []; s.freeze = j.freeze || { rows: 0, cols: 0 }; s.cf = j.cf || []; s.validations = j.validations || []; s.charts = j.charts || []; s.filter = j.filter || null; s.tabColor = j.tabColor || null; s.showGridlines = j.showGridlines !== false;
    if (j.view) s.view = { ...s.view, ...j.view };
    return s;
  }
}

function serializeValue(v) { if (v && typeof v === 'object' && v.error) return { error: v.error }; return v; }
function deserializeValue(v) { if (v && typeof v === 'object' && v.error) return null; return v; }

export function mergeStyle(base, patch) {
  const out = { ...(base || {}) };
  for (const [k, v] of Object.entries(patch)) {
    if (v === null || v === undefined || v === false) delete out[k];
    else if (k === 'border' && typeof v === 'object') { out.border = { ...(out.border || {}) }; for (const [side, b] of Object.entries(v)) { if (b == null) delete out.border[side]; else out.border[side] = b; } if (!Object.keys(out.border).length) delete out.border; }
    else out[k] = v;
  }
  return out;
}

// ---------------- Workbook ----------------
export class Workbook {
  constructor() {
    this.sheets = [];
    this.styles = [{}];
    this.styleIndex = new Map([['{}', 0]]);
    this.names = new Map();
    this.engine = new Engine(this);
    this.undoStack = []; this.redoStack = [];
    this._tx = null;
    this._listeners = new Map();
    this._needsRebuild = false;
    this.name = 'Book1';
    this.activeSheetIndex = 0;
    this.dirty = false;
    this.maxUndo = 200;
  }
  static blank(sheetCount = 1) { const wb = new Workbook(); for (let i = 0; i < sheetCount; i++) wb._addSheetRaw(`Sheet${i + 1}`); return wb; }

  on(ev, fn) { if (!this._listeners.has(ev)) this._listeners.set(ev, new Set()); this._listeners.get(ev).add(fn); return () => this._listeners.get(ev).delete(fn); }
  emit(ev, data) { const s = this._listeners.get(ev); if (s) for (const fn of s) fn(data); }

  get activeSheet() { return this.sheets[Math.min(this.activeSheetIndex, this.sheets.length - 1)]; }
  getSheet(name) { const low = String(name).toLowerCase(); return this.sheets.find((s) => s.name.toLowerCase() === low) || null; }
  sheetById(id) { return this.sheets.find((s) => s.id === id) || null; }

  styleId(style) {
    const clean = {};
    for (const k of Object.keys(style).sort()) { const v = style[k]; if (v === null || v === undefined || v === false || v === '' ) continue; if (k === 'border' && (!v || !Object.keys(v).length)) continue; clean[k] = v; }
    const key = JSON.stringify(clean);
    let id = this.styleIndex.get(key);
    if (id === undefined) { id = this.styles.length; this.styles.push(clean); this.styleIndex.set(key, id); }
    return id;
  }

  // ---- names ----
  getName(name, sheet) {
    const up = name.toUpperCase();
    const local = sheet ? this.names.get(`${sheet.id}!${up}`) : null;
    return local || this.names.get(up) || null;
  }
  setName(name, ref, sheetScope = null) {
    const k = (sheetScope ? `${sheetScope.id}!` : '') + name.toUpperCase();
    const old = this.names.get(k) || null;
    this._record(() => { if (old) this.names.set(k, old); else this.names.delete(k); this.engine.rebuildAll(); this._touch(); });
    if (ref == null) this.names.delete(k); else this.names.set(k, { name, ref: ref.startsWith('=') ? ref : '=' + ref, sheetId: sheetScope ? sheetScope.id : null });
    this._needsRebuild = true; this._touch();
  }
  deleteName(name, sheetScope = null) { this.setName(name, null, sheetScope); }

  // ---- sheets ----
  _addSheetRaw(name, index = this.sheets.length) { const s = new Sheet(this, name); this.sheets.splice(index, 0, s); return s; }
  uniqueSheetName(base = 'Sheet') { let i = this.sheets.length + 1; let n = `${base}${i}`; while (this.getSheet(n)) n = `${base}${++i}`; return n; }
  addSheet(name = this.uniqueSheetName(), index = this.sheets.length) {
    if (this.getSheet(name)) throw new Error(`Sheet "${name}" already exists`);
    const s = this._addSheetRaw(name, index);
    this._record(() => this._removeSheetObj(s));
    this._touch(); this.emit('sheets'); return s;
  }
  _removeSheetObj(s) {
    const idx = this.sheets.indexOf(s); if (idx < 0) return;
    this.sheets.splice(idx, 1);
    this._record(() => { this.sheets.splice(Math.min(idx, this.sheets.length), 0, s); this._needsRebuild = true; this._touch(); this.emit('sheets'); });
    if (this.activeSheetIndex >= this.sheets.length) this.activeSheetIndex = this.sheets.length - 1;
    this._needsRebuild = true; this._touch(); this.emit('sheets');
  }
  removeSheet(s) {
    if (this.sheets.length <= 1) throw new Error('A workbook needs at least one sheet');
    for (const sh of this.sheets) if (sh !== s) sh._rewriteFormulas((f) => removeSheetInFormula(f, s.name));
    this._removeSheetObj(s);
  }
  renameSheet(s, newName) {
    newName = newName.trim();
    if (!newName || /[\\/?*[\]:]/.test(newName) || newName.length > 31) throw new Error('Invalid sheet name');
    if (this.getSheet(newName) && this.getSheet(newName) !== s) throw new Error(`Sheet "${newName}" already exists`);
    const old = s.name;
    for (const sh of this.sheets) sh._rewriteFormulas((f) => renameSheetInFormula(f, old, newName));
    for (const [k, n] of this.names) if (n.ref) { const nr = renameSheetInFormula(n.ref, old, newName); if (nr !== n.ref) { this._record(() => { n.ref = renameSheetInFormula(nr, newName, old); n.ast = null; }); n.ref = nr; n.ast = null; } }
    this._record(() => { s.name = old; this._touch(); this.emit('sheets'); });
    s.name = newName; this._needsRebuild = true; this._touch(); this.emit('sheets');
  }
  moveSheet(s, toIndex) {
    const from = this.sheets.indexOf(s); if (from < 0 || from === toIndex) return;
    this._record(() => this.moveSheet(s, from));
    this.sheets.splice(from, 1); this.sheets.splice(toIndex, 0, s);
    this._touch(); this.emit('sheets');
  }
  duplicateSheet(s) {
    const j = s.toJSON(); j.name = this.uniqueSheetName(s.name.replace(/ \(\d+\)$/, '') + ' ('); j.name = j.name.replace(/\($/, '') ; j.id = uid('s');
    let base = s.name, n = 2; while (this.getSheet(`${base} (${n})`)) n++; j.name = `${base} (${n})`;
    j.charts = deepClone(j.charts).map((c) => ({ ...c, id: uid('ch') }));
    const ns = Sheet.fromJSON(this, j);
    const idx = this.sheets.indexOf(s) + 1;
    this.sheets.splice(idx, 0, ns);
    this._record(() => this._removeSheetObj(ns));
    this._needsRebuild = true; this._touch(); this.emit('sheets');
    return ns;
  }
  setActiveSheet(i) { this.activeSheetIndex = Math.max(0, Math.min(i, this.sheets.length - 1)); this.emit('activeSheet'); }

  // ---- transactions / undo ----
  transact(label, fn) {
    if (this._tx) return fn(); // nested: join
    this._tx = { label, inverses: [], cellKeys: new Set(), touched: false };
    let result, error;
    try { result = fn(); } catch (e) { error = e; }
    const tx = this._tx; this._tx = null;
    const inverses = tx.inverses;
    if (this._needsRebuild) { this._needsRebuild = false; this.engine.rebuildAll(); } else this.engine.recalc();
    if (inverses.length) {
      if (!tx.isUndo && !tx.isRedo) { this.undoStack.push({ label, inverses }); if (this.undoStack.length > this.maxUndo) this.undoStack.shift(); this.redoStack.length = 0; }
      this.dirty = true;
    }
    if (tx.touched || inverses.length) this.emit('change', { label });
    if (error) throw error;
    return result;
  }
  _record(inverse) { if (this._tx) this._tx.inverses.push(inverse); }
  // Cell changes are recorded in sequence with structural inverses so that undo restores each cell
  // at the position it had when it was changed. Within a run of plain edits only the first change per cell is kept.
  _recordCell(sheet, k, old) {
    if (!this._tx) return;
    const kk = `${sheet.id}!${k}`;
    if (this._tx.cellKeys.has(kk)) return;
    this._tx.cellKeys.add(kk);
    const snap = old ? { ...old } : null;
    const [r, c] = unkey(k);
    this._tx.inverses.push(() => sheet._setCellRaw(r, c, snap));
  }
  // A structural change moves cells, so later cell edits must get fresh inverse entries.
  _structuralBoundary() { if (this._tx) this._tx.cellKeys = new Set(); }
  _touch() { if (this._tx) this._tx.touched = true; }
  _runInverse(entry, isUndo) {
    this._tx = { label: entry.label, inverses: [], cellKeys: new Set(), touched: false, isUndo, isRedo: !isUndo };
    try { for (let i = entry.inverses.length - 1; i >= 0; i--) entry.inverses[i](); } finally {
      const tx = this._tx; this._tx = null;
      const inverses = tx.inverses;
      if (this._needsRebuild) { this._needsRebuild = false; this.engine.rebuildAll(); } else this.engine.recalc();
      (isUndo ? this.redoStack : this.undoStack).push({ label: entry.label, inverses });
      this.dirty = true;
      this.emit('change', { label: entry.label, undo: isUndo });
    }
  }
  undo() { const e = this.undoStack.pop(); if (!e) return false; this._runInverse(e, true); return true; }
  redo() { const e = this.redoStack.pop(); if (!e) return false; this._runInverse(e, false); return true; }
  canUndo() { return this.undoStack.length > 0; }
  canRedo() { return this.redoStack.length > 0; }

  // ---- structural change implementation (rows/cols insert/delete) ----
  _structural(sheet, axis, index, count) {
    if (!count) return;
    const isRow = axis === 'row';
    const max = isRow ? MAX_ROWS : MAX_COLS;
    let inverse;
    if (count < 0) {
      // snapshot deleted cells
      const removed = [];
      const span = { lo: index, hi: index - count - 1 };
      for (const [k, cell] of sheet.cells) { const [r, c] = unkey(k); const i = isRow ? r : c; if (i >= span.lo && i <= span.hi) removed.push({ r, c, cell: { ...cell } }); }
      const removedSizes = []; const ax = isRow ? sheet.rowAxis : sheet.colAxis;
      for (let i = span.lo; i <= span.hi; i++) removedSizes.push({ i, size: ax.sizes.get(i), hidden: ax.hidden.has(i), style: (isRow ? sheet.rowStyles : sheet.colStyles).get(i) });
      const oldMerges = sheet.merges, oldCf = sheet.cf, oldVal = sheet.validations, oldCharts = sheet.charts, oldFilter = sheet.filter, oldFreeze = { ...sheet.freeze };
      const oldNames = [...this.names.entries()].map(([k, v]) => [k, { ...v }]);
      // collect formulas that will change so the inverse can restore them exactly
      const formulaBackup = [];
      for (const sh of this.sheets) for (const [k, cell] of sh.cells) if (cell.f) formulaBackup.push({ sh, k, f: cell.f });
      inverse = () => {
        this._structural(sheet, axis, index, -count);
        for (const { i, size, hidden, style } of removedSizes) { if (size != null) ax.sizes.set(i, size); if (hidden) ax.hidden.add(i); if (style) (isRow ? sheet.rowStyles : sheet.colStyles).set(i, style); }
        ax.invalidate();
        for (const { r, c, cell } of removed) sheet._setCellRaw(r, c, cell);
        for (const { sh, k, f } of formulaBackup) { const cell = sh.cells.get(k); if (cell && cell.f !== f) { const [r, c] = unkey(k); sh._setCellRaw(r, c, { ...cell, f }); } }
        sheet.merges = oldMerges; sheet.cf = oldCf; sheet.validations = oldVal; sheet.charts = oldCharts; sheet.filter = oldFilter; sheet.freeze = oldFreeze;
        this.names = new Map(oldNames.map(([k, v]) => [k, { ...v }]));
        sheet.applyFilter(); this._needsRebuild = true; this._touch();
      };
    } else {
      inverse = () => this._structural(sheet, axis, index, -count);
    }
    // 1. rewrite formulas everywhere (recorded as ordinary cell changes at their pre-shift positions)
    for (const sh of this.sheets) sh._rewriteFormulas((f) => adjustFormulaForInsertDelete(f, sh.name, sheet.name, axis, index, count));
    for (const [k, n] of this.names) { const nr = adjustFormulaForInsertDelete(n.ref, sheet.name, sheet.name, axis, index, count); if (nr !== n.ref) { n.ref = nr; n.ast = null; } }
    // the structural inverse runs before those cell restores during undo, so positions line up first
    this._record(inverse); this._structuralBoundary();
    // 2. move cells
    const moved = [];
    for (const [k, cell] of sheet.cells) {
      const [r, c] = unkey(k); const i = isRow ? r : c;
      if (i < index) continue;
      if (count < 0 && i <= index - count - 1) { moved.push({ k, cell: null, r, c }); continue; }
      const ni = i + count;
      moved.push({ k, cell, r, c, nr: isRow ? ni : r, nc: isRow ? c : ni });
    }
    for (const m of moved) sheet.cells.delete(m.k);
    for (const m of moved) if (m.cell && (isRow ? m.nr : m.nc) < max) { const cell = { ...m.cell }; delete cell.spill; delete cell.spillOf; if (!cell.f && cell.v == null && !cell.s && !cell.note) continue; sheet.cells.set(cellKey(m.nr, m.nc), cell); }
    // 3. axis sizes/hidden/styles
    const ax = isRow ? sheet.rowAxis : sheet.colAxis;
    ax.shift(index, count);
    const styles = isRow ? sheet.rowStyles : sheet.colStyles;
    const ns = new Map(); for (const [i, v] of styles) { if (i < index) ns.set(i, v); else if (count > 0) ns.set(i + count, v); else if (i >= index - count) ns.set(i + count, v); }
    if (isRow) sheet.rowStyles = ns; else sheet.colStyles = ns;
    // 4. merges / cf / validations / charts / filter / freeze
    const adj = (rg) => adjustRange(rg, isRow, index, count);
    sheet.merges = sheet.merges.map(adj).filter(Boolean).filter((m) => m.r1 !== m.r2 || m.c1 !== m.c2);
    sheet.cf = sheet.cf.map((rule) => { const rg = adj(rule); return rg ? { ...rule, ...rg } : null; }).filter(Boolean);
    sheet.validations = sheet.validations.map((v) => { const rg = adj(v); return rg ? { ...v, ...rg } : null; }).filter(Boolean);
    sheet.charts = sheet.charts.map((ch) => { const a = ch.anchor; if (!a) return ch; const i = isRow ? a.r : a.c; let ni = i; if (i >= index) ni = count > 0 ? i + count : Math.max(index, i + count); return { ...ch, anchor: { ...a, [isRow ? 'r' : 'c']: ni } }; });
    if (sheet.filter) { const rg = adj(sheet.filter); sheet.filter = rg ? { ...sheet.filter, ...rg, conds: isRow ? sheet.filter.conds : shiftConds(sheet.filter.conds, index, count) } : null; }
    const fz = isRow ? sheet.freeze.rows : sheet.freeze.cols;
    if (fz > index) { const nf = Math.max(index, fz + count); sheet.freeze = isRow ? { ...sheet.freeze, rows: nf } : { ...sheet.freeze, cols: nf }; }
    sheet.applyFilter();
    sheet._bounds = null;
    this._needsRebuild = true; this._touch();
  }

  // display text for a cell (used by filters, clipboard, etc.) - set by the UI layer via formatter hook
  displayText(sheet, r, c) { return this.formatter ? this.formatter(sheet, r, c).text : String(sheet.getValue(r, c) ?? ''); }

  toJSON() { return { app: 'officetable', version: 1, name: this.name, activeSheet: this.activeSheetIndex, styles: this.styles, names: [...this.names.entries()].map(([k, v]) => [k, { name: v.name, ref: v.ref, sheetId: v.sheetId }]), sheets: this.sheets.map((s) => s.toJSON()) }; }
  static fromJSON(j) {
    const wb = new Workbook();
    wb.name = j.name || 'Book1';
    wb.styles = j.styles && j.styles.length ? j.styles : [{}];
    wb.styleIndex = new Map(wb.styles.map((s, i) => [JSON.stringify(s), i]));
    wb.names = new Map((j.names || []).map(([k, v]) => [k, { ...v }]));
    for (const sj of j.sheets || []) wb.sheets.push(Sheet.fromJSON(wb, sj));
    if (!wb.sheets.length) wb._addSheetRaw('Sheet1');
    wb.activeSheetIndex = Math.min(j.activeSheet || 0, wb.sheets.length - 1);
    wb.engine.rebuildAll();
    for (const s of wb.sheets) s.applyFilter();
    return wb;
  }
}

function adjustRange(rg, isRow, index, count) {
  let a = isRow ? rg.r1 : rg.c1, b = isRow ? rg.r2 : rg.c2;
  const max = (isRow ? MAX_ROWS : MAX_COLS) - 1;
  if (b >= max && a === 0) return { ...rg };
  if (count > 0) { if (a >= index) a += count; if (b >= index) b += count; }
  else {
    const ds = index, de = index - count - 1;
    if (a >= ds && b <= de) return null;
    if (a > de) a += count; else if (a >= ds) a = ds;
    if (b > de) b += count; else if (b >= ds) b = ds - 1;
    if (b < a) return null;
  }
  if (a > max) return null; if (b > max) b = max;
  return isRow ? { ...rg, r1: a, r2: b } : { ...rg, c1: a, c2: b };
}
function shiftConds(conds, index, count) {
  const out = {};
  for (const [k, v] of Object.entries(conds || {})) { const c = +k; if (c < index) out[c] = v; else if (count > 0) out[c + count] = v; else if (c >= index - count) out[c + count] = v; }
  return out;
}

export { cellKey, unkey, parseRangeName };
