// OfficeTable application: wires model, grid, editor, clipboard, charts, dialogs and commands.
import { Workbook, Sheet, mergeStyle } from './model.js';
import { GridView } from './grid.js';
import { CellEditor } from './editor.js';
import { ClipboardManager } from './clipboard.js';
import { ChartLayer, newChart, CHART_TYPES } from './charts.js';
import { createDialogs, h } from './dialogs.js';
import { buildMenus, refreshToolbar, ICONS } from './menus.js';
import { autofill, fillDirection } from './fill.js';
import { formatValue, parseInput, isDateFormat, todaySerial, nowSerial, FORMAT_PRESETS } from './format.js';
import { exportXlsx, importXlsx } from './xlsx.js';
import { parseCSV, toCSV } from './csv.js';
import { parse, isFormula, shiftFormula, FormulaError, isErr } from './formula.js';
import { FUNCTIONS, num, bool, criteria } from './functions.js';
import { colToName, parseRangeName, rangeName, cellName, normRange, rangeContains, rangesIntersect, escapeHtml, debounce, MAX_ROWS, MAX_COLS, key as cellKey, unkey, quoteSheetName } from './util.js';
import * as formatMod from './format.js';
import * as formulaMod from './formula.js';
import * as modelMod from './model.js';
import * as xlsxMod from './xlsx.js';
import * as csvMod from './csv.js';
import * as chartsMod from './charts.js';

const STORAGE_KEY = 'officetable.autosave';
const SETTINGS_KEY = 'officetable.settings';

const app = {
  wb: null, grid: null, editor: null, clipboard: null, charts: null, dialogs: null,
  commands: {}, toolbarButtons: {}, settings: { autosave: true, showFormulas: false },
  pickingRef: false, painter: null, fileHandle: null,
  modules: { format: formatMod, formula: formulaMod, model: modelMod, xlsx: xlsxMod, csv: csvMod, charts: chartsMod, functions: FUNCTIONS },
  get sheet() { return this.wb.activeSheet; },
  get sel() { return this.sheet.view; },

  // ---------- display ----------
  fmt(sheet, r, c) {
    const cell = sheet.getCell(r, c);
    if (this.settings.showFormulas && cell && cell.f) return { text: cell.f, color: null };
    const v = sheet.getValue(r, c);
    if (v == null) return { text: '', color: null };
    if (typeof v === 'string' && !(sheet.getStyle(r, c).numFmt || '').includes('@')) return { text: v, color: null };
    return formatValue(v, sheet.effectiveNumFmt(r, c));
  },
  toast(msg, ms = 2200) { const t = document.getElementById('toast'); t.textContent = msg; t.classList.add('show'); clearTimeout(this._toastT); this._toastT = setTimeout(() => t.classList.remove('show'), ms); },
  setStatus(text) { document.getElementById('statusMode').textContent = text; },
  dialogOpen() { return this.dialogs.isOpen(); },
  focusGrid() { if (!this.editor.isEditing && !this.dialogOpen()) this.grid.canvas.focus({ preventScroll: true }); },

  // ---------- transactions ----------
  transact(label, fn) {
    let result;
    try { result = this.wb.transact(label, fn); }
    catch (e) { console.error(e); this.toast(e.message || String(e)); throw e; }
    this.afterChange();
    return result;
  },
  afterChange() { this.grid.invalidate(); this.refreshToolbar(); this.updateFormulaBar(); this.updateStatus(); this.scheduleAutosave(); this.updateTitle(); },
  refreshToolbar() { refreshToolbar(this); },
  updateTitle() { document.getElementById('saveState').textContent = this.wb.dirty ? '●' : ''; document.title = `${this.wb.name}${this.wb.dirty ? ' •' : ''} — OfficeTable`; },

  // ---------- selection ----------
  normalizeView() { const v = this.sheet.view; if (!v.ranges || !v.ranges.length) v.ranges = [{ r1: v.active.r, c1: v.active.c, r2: v.active.r, c2: v.active.c }]; },
  select(r, c, opts = {}) {
    r = Math.max(0, Math.min(MAX_ROWS - 1, r)); c = Math.max(0, Math.min(MAX_COLS - 1, c));
    const s = this.sheet; const v = s.view;
    const m = s.mergeAt(r, c);
    const rg = m ? { r1: m.r1, c1: m.c1, r2: m.r2, c2: m.c2 } : { r1: r, c1: c, r2: r, c2: c };
    if (opts.add) v.ranges = [...v.ranges, rg]; else v.ranges = [rg];
    v.active = { r, c }; v.anchor = { r, c };
    this.afterSelect(!opts.noScroll);
    if (this.painter && !opts.add) this.applyPainter();
  },
  selectRange(rg, opts = {}) {
    const s = this.sheet; const v = s.view;
    rg = normRange(rg.r1, rg.c1, rg.r2, rg.c2);
    if (opts.replaceLast) v.ranges = [...v.ranges.slice(0, -1), rg];
    else if (opts.add) v.ranges = [...v.ranges, rg];
    else v.ranges = [rg];
    if (opts.active) v.active = opts.active;
    else if (!opts.keepActive) v.active = { r: rg.r1 === 0 && rg.r2 >= MAX_ROWS - 1 ? Math.max(this.grid.scrollRow, 0) : rg.r1, c: rg.c1 === 0 && rg.c2 >= MAX_COLS - 1 ? Math.max(this.grid.scrollCol, 0) : rg.c1 };
    if (!rangeContains(rg, v.active.r, v.active.c) && !opts.keepActive) v.active = { r: Math.max(rg.r1, Math.min(rg.r2, v.active.r)), c: Math.max(rg.c1, Math.min(rg.c2, v.active.c)) };
    if (!opts.keepActive) v.anchor = { ...v.active };
    this.afterSelect(false);
  },
  extendSelectionTo(r, c, anchor) {
    const v = this.sheet.view; const a = anchor || v.anchor || v.active;
    let rg = normRange(a.r, a.c, r, c);
    // grow to include merged cells
    let grown = true; let guard = 0;
    while (grown && guard++ < 20) { grown = false; for (const m of this.sheet.merges) if (rangesIntersect(m, rg)) { const n = { r1: Math.min(rg.r1, m.r1), c1: Math.min(rg.c1, m.c1), r2: Math.max(rg.r2, m.r2), c2: Math.max(rg.c2, m.c2) }; if (n.r1 !== rg.r1 || n.c1 !== rg.c1 || n.r2 !== rg.r2 || n.c2 !== rg.c2) { rg = n; grown = true; } } }
    v.ranges = [...v.ranges.slice(0, -1), rg];
    this.afterSelect(false, { r, c });
  },
  selectAll() {
    const s = this.sheet; const v = s.view; const cur = v.ranges[0];
    const region = this.dataRegion();
    const isRegion = cur && region && cur.r1 === region.r1 && cur.c1 === region.c1 && cur.r2 === region.r2 && cur.c2 === region.c2;
    const whole = { r1: 0, c1: 0, r2: MAX_ROWS - 1, c2: MAX_COLS - 1 };
    if (region && !isRegion && rangeContains(region, v.active.r, v.active.c) && (region.r1 !== region.r2 || region.c1 !== region.c2)) v.ranges = [region]; else v.ranges = [whole];
    this.afterSelect(false);
  },
  afterSelect(scroll = true, scrollTo = null) {
    const v = this.sheet.view;
    if (scroll) this.grid.scrollTo(v.active.r, v.active.c);
    else if (scrollTo) this.grid.scrollTo(scrollTo.r, scrollTo.c);
    this.grid.invalidate(); this.updateFormulaBar(); this.refreshToolbar(); this.updateStatus();
  },
  selectedWholeCols() { const out = []; for (const rg of this.sheet.view.ranges) if (rg.r1 === 0 && rg.r2 >= MAX_ROWS - 1) for (let c = rg.c1; c <= Math.min(rg.c2, rg.c1 + 1000); c++) out.push(c); return out; },
  selectedWholeRows() { const out = []; for (const rg of this.sheet.view.ranges) if (rg.c1 === 0 && rg.c2 >= MAX_COLS - 1) for (let r = rg.r1; r <= Math.min(rg.r2, rg.r1 + 10000); r++) out.push(r); return out; },
  selectedCols() { const set = new Set(); for (const rg of this.sheet.view.ranges) for (let c = rg.c1; c <= Math.min(rg.c2, rg.c1 + 1000); c++) set.add(c); return [...set].sort((a, b) => a - b); },
  selectedRows() { const set = new Set(); for (const rg of this.sheet.view.ranges) for (let r = rg.r1; r <= Math.min(rg.r2, rg.r1 + 10000); r++) set.add(r); return [...set].sort((a, b) => a - b); },
  // contiguous block of data around the active cell (or the selection if multi-cell)
  dataRegion() {
    const s = this.sheet; const rg = s.view.ranges[0];
    if (rg.r1 !== rg.r2 || rg.c1 !== rg.c2) { const ub = s.usedBounds(); return { r1: rg.r1, c1: rg.c1, r2: Math.min(rg.r2, Math.max(rg.r1, ub.r2)), c2: Math.min(rg.c2, Math.max(rg.c1, ub.c2)) }; }
    const has = (r, c) => { const cell = s.getCell(r, c); return !!cell && (cell.v != null || cell.f); };
    let { r, c } = s.view.active;
    if (!has(r, c)) return rg;
    let r1 = r, r2 = r, c1 = c, c2 = c; let grown = true; let guard = 0;
    while (grown && guard++ < 10000) {
      grown = false;
      const rowHas = (rr) => { for (let cc = c1; cc <= c2; cc++) if (has(rr, cc)) return true; return false; };
      const colHas = (cc) => { for (let rr = r1; rr <= r2; rr++) if (has(rr, cc)) return true; return false; };
      if (r1 > 0 && rowHas(r1 - 1)) { r1--; grown = true; }
      if (rowHas(r2 + 1)) { r2++; grown = true; }
      if (c1 > 0 && colHas(c1 - 1)) { c1--; grown = true; }
      if (colHas(c2 + 1)) { c2++; grown = true; }
    }
    return { r1, c1, r2, c2 };
  },
  sortRegion() { const rg = this.sheet.view.ranges[0]; if (rg.r1 !== rg.r2 || rg.c1 !== rg.c2) return this.dataRegion(); return this.dataRegion(); },
  looksLikeHeader(rg) { const s = this.sheet; if (rg.r1 === rg.r2) return false; let text = 0, n = 0; for (let c = rg.c1; c <= Math.min(rg.c2, rg.c1 + 50); c++) { const v = s.getValue(rg.r1, c), v2 = s.getValue(rg.r1 + 1, c); if (v == null) continue; n++; if (typeof v === 'string' && typeof v2 !== 'string' && v2 != null) text++; } return n > 0 && text >= Math.ceil(n / 2); },
  moveActive(dir, opts = {}) {
    const s = this.sheet; const v = s.view; let { r, c } = v.active;
    const rg = v.ranges[0]; const multi = v.ranges.length === 1 && (rg.r1 !== rg.r2 || rg.c1 !== rg.c2) && opts.enter;
    const nextVisible = (axis, i, step) => { const ax = axis === 'row' ? s.rowAxis : s.colAxis; const max = axis === 'row' ? MAX_ROWS - 1 : MAX_COLS - 1; let j = i + step; while (j >= 0 && j <= max && ax.size(j) === 0) j += step; return Math.max(0, Math.min(max, j)); };
    const m = s.mergeAt(r, c);
    if (multi) {
      // move within the selection, wrapping
      if (dir === 'down') { r++; if (r > rg.r2) { r = rg.r1; c++; if (c > rg.c2) c = rg.c1; } }
      else if (dir === 'up') { r--; if (r < rg.r1) { r = rg.r2; c--; if (c < rg.c1) c = rg.c2; } }
      else if (dir === 'right') { c++; if (c > rg.c2) { c = rg.c1; r++; if (r > rg.r2) r = rg.r1; } }
      else if (dir === 'left') { c--; if (c < rg.c1) { c = rg.c2; r--; if (r < rg.r1) r = rg.r2; } }
      v.active = { r, c }; v.anchor = { r, c }; this.afterSelect(true); return;
    }
    switch (dir) {
      case 'down': r = nextVisible('row', m ? m.r2 : r, 1); break;
      case 'up': r = nextVisible('row', m ? m.r1 : r, -1); break;
      case 'right': c = nextVisible('col', m ? m.c2 : c, 1); break;
      case 'left': c = nextVisible('col', m ? m.c1 : c, -1); break;
      case 'home': c = 0; break;
      case 'end': { const ub = s.usedBounds(); c = ub.c2; break; }
      case 'ctrlHome': r = 0; c = 0; break;
      case 'ctrlEnd': { const db = s.dataBounds(); if (db) { r = db.r2; c = db.c2; } else { r = 0; c = 0; } break; }
      case 'pagedown': r = Math.min(MAX_ROWS - 1, r + Math.max(1, this.grid.visibleRowCount() - 1)); break;
      case 'pageup': r = Math.max(0, r - Math.max(1, this.grid.visibleRowCount() - 1)); break;
      case 'pageright': c = Math.min(MAX_COLS - 1, c + Math.max(1, this.grid.visibleColCount() - 1)); break;
      case 'pageleft': c = Math.max(0, c - Math.max(1, this.grid.visibleColCount() - 1)); break;
      case 'ctrlDown': r = this.jump(r, c, 1, 0); break;
      case 'ctrlUp': r = this.jump(r, c, -1, 0); break;
      case 'ctrlRight': c = this.jump(r, c, 0, 1); break;
      case 'ctrlLeft': c = this.jump(r, c, 0, -1); break;
    }
    this.select(r, c);
  },
  // Ctrl+arrow semantics: jump to the edge of the current data block
  jump(r, c, dr, dc) {
    const s = this.sheet; const has = (rr, cc) => { const cell = s.getCell(rr, cc); return !!cell && (cell.v != null || cell.f); };
    const maxR = MAX_ROWS - 1, maxC = MAX_COLS - 1;
    let nr = r + dr, nc = c + dc;
    if (nr < 0 || nr > maxR || nc < 0 || nc > maxC) return dr ? r : c;
    const db = s.dataBounds() || { r1: 0, c1: 0, r2: 0, c2: 0 };
    if (has(r, c) && has(nr, nc)) { while (nr + dr >= 0 && nr + dr <= maxR && nc + dc >= 0 && nc + dc <= maxC && has(nr + dr, nc + dc)) { nr += dr; nc += dc; } return dr ? nr : nc; }
    // move to next non-empty cell, or the edge
    const limitR = dr > 0 ? Math.max(db.r2, r) : 0, limitC = dc > 0 ? Math.max(db.c2, c) : 0;
    while (!has(nr, nc)) { if ((dr > 0 && nr >= limitR) || (dr < 0 && nr <= 0) || (dc > 0 && nc >= limitC) || (dc < 0 && nc <= 0)) { return dr ? (dr > 0 ? (has(nr, nc) ? nr : (db.r2 > r ? nr : maxR)) : 0) : (dc > 0 ? (has(nr, nc) ? nc : (db.c2 > c ? nc : maxC)) : 0); } nr += dr; nc += dc; }
    return dr ? nr : nc;
  },
  extendActive(dir) {
    const v = this.sheet.view; const rg = v.ranges[v.ranges.length - 1]; const a = v.anchor || v.active;
    // the moving corner is the one opposite to the anchor
    let r = a.r === rg.r1 ? rg.r2 : rg.r1, c = a.c === rg.c1 ? rg.c2 : rg.c1;
    const s = this.sheet;
    switch (dir) {
      case 'down': r = Math.min(MAX_ROWS - 1, r + 1); break; case 'up': r = Math.max(0, r - 1); break; case 'right': c = Math.min(MAX_COLS - 1, c + 1); break; case 'left': c = Math.max(0, c - 1); break;
      case 'ctrlDown': r = this.jump(r, c, 1, 0); break; case 'ctrlUp': r = this.jump(r, c, -1, 0); break; case 'ctrlRight': c = this.jump(r, c, 0, 1); break; case 'ctrlLeft': c = this.jump(r, c, 0, -1); break;
      case 'home': c = 0; break; case 'ctrlHome': r = 0; c = 0; break; case 'ctrlEnd': { const db = s.dataBounds(); if (db) { r = db.r2; c = db.c2; } break; }
      case 'pagedown': r = Math.min(MAX_ROWS - 1, r + this.grid.visibleRowCount()); break; case 'pageup': r = Math.max(0, r - this.grid.visibleRowCount()); break;
    }
    this.extendSelectionTo(r, c, a);
  },
  goTo(text) {
    const t = text.trim(); if (!t) return;
    let rg = parseRangeName(t);
    if (!rg) { const n = this.wb.getName(t, this.sheet); if (n) rg = parseRangeName(n.ref.replace(/^=/, '').replace(/\$/g, '')); }
    if (!rg) { this.toast(`"${t}" is not a valid reference or name`); return; }
    if (rg.sheet) { const sh = this.wb.getSheet(rg.sheet); if (!sh) { this.toast(`Sheet "${rg.sheet}" not found`); return; } this.switchSheet(this.wb.sheets.indexOf(sh)); }
    this.selectRange(rg, { active: { r: rg.r1, c: rg.c1 } }); this.grid.scrollTo(rg.r1, rg.c1);
  },
  selectionText() { const s = this.sheet; const { r, c } = s.view.active; const v = s.getValue(r, c); return v == null ? '' : String(this.fmt(s, r, c).text); },

  // ---------- formula bar / name box ----------
  updateFormulaBar() {
    if (this.editor.isEditing) return;
    const s = this.sheet; const { r, c } = s.view.active;
    document.getElementById('formulaInput').value = CellEditor.editText(this, s, r, c);
    const rg = s.view.ranges[s.view.ranges.length - 1];
    const nb = document.getElementById('nameBox');
    if (document.activeElement !== nb) {
      const names = [...this.wb.names.values()].filter((n) => { const p = parseRangeName(n.ref.replace(/^=/, '').replace(/\$/g, '')); return p && (!p.sheet || p.sheet.toLowerCase() === s.name.toLowerCase()) && p.r1 === rg.r1 && p.c1 === rg.c1 && p.r2 === rg.r2 && p.c2 === rg.c2; });
      nb.value = names.length ? names[0].name : (rg.r1 !== rg.r2 || rg.c1 !== rg.c2 ? `${rg.r2 - rg.r1 + 1}R x ${rg.c2 - rg.c1 + 1}C` : cellName(r, c));
      if (rg.r1 !== rg.r2 || rg.c1 !== rg.c2) { if (!names.length && !this.grid.drag) nb.value = rangeName(rg); }
    }
  },
  updateStatus() {
    const s = this.sheet; const el = document.getElementById('stats');
    let count = 0, nums = 0, sum = 0, min = Infinity, max = -Infinity; let cells = 0;
    for (const rg of s.view.ranges) {
      const r2 = Math.min(rg.r2, s.usedBounds().r2), c2 = Math.min(rg.c2, s.usedBounds().c2);
      for (let r = rg.r1; r <= r2; r++) { if (s.rowAxis.isHidden(r)) continue; for (let c = rg.c1; c <= c2; c++) { if (++cells > 200000) break; const v = s.getValue(r, c); if (v == null) continue; count++; if (typeof v === 'number') { nums++; sum += v; if (v < min) min = v; if (v > max) max = v; } } }
    }
    const f = (v) => formatValue(v, s.effectiveNumFmt(s.view.active.r, s.view.active.c) === 'General' ? 'General' : s.effectiveNumFmt(s.view.active.r, s.view.active.c)).text;
    el.innerHTML = count > 1 ? `<span>Count: <b>${count}</b></span>${nums ? `<span>Sum: <b>${escapeHtml(f(sum))}</b></span><span>Average: <b>${escapeHtml(f(sum / nums))}</b></span><span>Min: <b>${escapeHtml(f(min))}</b></span><span>Max: <b>${escapeHtml(f(max))}</b></span>` : ''}` : '';
    const ci = document.getElementById('calcInfo'); const n = this.wb.engine.state.size; ci.textContent = n ? `${n} formula${n === 1 ? '' : 's'}` : '';
  },
  onRendered() { /* hook for tests */ },
  onHoverCell(hit, e) {
    let tip = document.querySelector('.cellTip'); if (tip) tip.remove();
    if (!hit) return;
    const s = this.sheet; const cell = s.getCell(hit.r, hit.c);
    const dv = s.validationAt(hit.r, hit.c);
    const parts = [];
    if (cell && cell.note) parts.push(cell.note);
    if (cell && cell.hyperlink) parts.push(`${cell.hyperlink}\nCtrl+click to follow`);
    if (dv && dv.prompt && hit.r === s.view.active.r && hit.c === s.view.active.c) parts.push(dv.prompt);
    const v = s.getValue(hit.r, hit.c);
    if (isErr(v) && v.detail) parts.push(`${v.error}: ${v.detail}`);
    else if (isErr(v)) parts.push(({ '#DIV/0!': 'Division by zero', '#N/A': 'Value not available', '#NAME?': 'Unknown function or name', '#REF!': 'Invalid cell reference', '#VALUE!': 'Wrong type of argument', '#NUM!': 'Invalid numeric value', '#SPILL!': 'Spill range is blocked by other cells', '#CIRC!': 'Circular reference', '#ERROR!': 'Formula could not be parsed', '#CALC!': 'Calculation returned an empty array' })[v.error] || v.error);
    if (!parts.length) return;
    const rect = this.grid.cellRect(hit.r, hit.c); if (!rect) return;
    tip = h('div', { class: 'cellTip' }, parts.join('\n\n'));
    this.grid.overlay.appendChild(tip);
    tip.style.left = Math.min(rect.x + rect.w + 4, this.grid.cssW - 300) + 'px'; tip.style.top = Math.max(0, rect.y - 4) + 'px';
  },

  // ---------- styling helpers ----------
  applyStyle(patch) {
    const s = this.sheet;
    this.transact('Format', () => { for (const rg of s.view.ranges) s.applyStyle(rg, patch); });
    this.focusGrid();
  },
  toggleStyle(prop) { const s = this.sheet; const { r, c } = s.view.active; const cur = !!s.getStyle(r, c)[prop]; this.applyStyle({ [prop]: cur ? null : true }); },
  applyNumFmt(code) { this.applyStyle({ numFmt: code === 'General' ? null : code }); },
  adjustDecimals(delta) {
    const s = this.sheet; const { r, c } = s.view.active; let code = s.effectiveNumFmt(r, c);
    const v = s.getValue(r, c);
    if (code === 'General') { const t = typeof v === 'number' ? formatValue(v, 'General').text : '0'; const dec = (t.split('.')[1] || '').length; code = '0' + (dec ? '.' + '0'.repeat(dec) : ''); }
    const m = /^([^;]*?)(0)(\.(0+))?/.exec(code);
    if (!m) { this.toast('Cannot change decimals for this format'); return; }
    const cur = m[4] ? m[4].length : 0; const n = Math.max(0, Math.min(15, cur + delta));
    const re = /(0)(\.0+)?/g;
    const out = code.replace(re, (mm, z, d) => z + (n ? '.' + '0'.repeat(n) : ''));
    this.applyNumFmt(out);
  },
  applyBorderPreset(rg, preset, style = 'thin', color = '#000000') {
    const s = this.sheet; const spec = { style, color };
    const r2 = Math.min(rg.r2, Math.max(rg.r1, s.usedBounds().r2 + 50)), c2 = Math.min(rg.c2, Math.max(rg.c1, s.usedBounds().c2 + 20));
    if ((r2 - rg.r1 + 1) * (c2 - rg.c1 + 1) > 50000) { this.toast('Selection too large for borders'); return; }
    const set = (r, c, sides) => { const cell = s.getCell(r, c); const st = s.wb.styles[s.getStyleId(r, c)] || {}; const border = { ...(st.border || {}) }; for (const [k, v] of Object.entries(sides)) { if (v === null) delete border[k]; else border[k] = v; } const sid = s.wb.styleId(mergeStyle(st, { border: Object.keys(border).length ? border : null })); if (!Object.keys(border).length) { const clean = { ...st }; delete clean.border; const sid2 = s.wb.styleId(clean); s._setCellRaw(r, c, cell ? { ...cell, s: sid2 } : (sid2 ? { v: null, f: null, s: sid2 } : null)); } else s._setCellRaw(r, c, cell ? { ...cell, s: sid } : { v: null, f: null, s: sid }); };
    for (let r = rg.r1; r <= r2; r++) for (let c = rg.c1; c <= c2; c++) {
      const top = r === rg.r1, bottom = r === r2, left = c === rg.c1, right = c === c2;
      let sides = {};
      switch (preset) {
        case 'none': sides = { top: null, bottom: null, left: null, right: null }; break;
        case 'all': sides = { top: spec, bottom: spec, left: spec, right: spec }; break;
        case 'outline': sides = { ...(top ? { top: spec } : {}), ...(bottom ? { bottom: spec } : {}), ...(left ? { left: spec } : {}), ...(right ? { right: spec } : {}) }; break;
        case 'inside': sides = { ...(top ? {} : { top: spec }), ...(bottom ? {} : { bottom: spec }), ...(left ? {} : { left: spec }), ...(right ? {} : { right: spec }) }; break;
        case 'top': if (top) sides = { top: spec }; break;
        case 'bottom': if (bottom) sides = { bottom: spec }; break;
        case 'left': if (left) sides = { left: spec }; break;
        case 'right': if (right) sides = { right: spec }; break;
        case 'thickBottom': if (bottom) sides = { bottom: { style: 'thick', color } }; break;
        case 'doubleBottom': if (bottom) sides = { bottom: { style: 'double', color } }; break;
        case 'topBottom': sides = { ...(top ? { top: spec } : {}), ...(bottom ? { bottom: spec } : {}) }; break;
        case 'leftRight': sides = { ...(left ? { left: spec } : {}), ...(right ? { right: spec } : {}) }; break;
      }
      if (Object.keys(sides).length) set(r, c, sides);
    }
  },
  applyPainter() { const sid = this.painter; this.painter = null; const s = this.sheet; const st = this.wb.styles[sid] || {}; this.transact('Paste format', () => { for (const rg of s.view.ranges) { const r2 = Math.min(rg.r2, rg.r1 + 5000), c2 = Math.min(rg.c2, rg.c1 + 500); for (let r = rg.r1; r <= r2; r++) for (let c = rg.c1; c <= c2; c++) { const cell = s.getCell(r, c); s._setCellRaw(r, c, cell ? { ...cell, s: sid } : (sid ? { v: null, f: null, s: sid } : null)); } } }); this.toolbarButtons['format.painter'].classList.remove('active'); this.setStatus('Ready'); },
  autofill(src, dst, opts) { const s = this.sheet; this.transact('Autofill', () => autofill(s, src, dst, opts)); this.selectRange(dst, { active: s.view.active }); },

  // ---------- validation ----------
  validationListValues(sheet, dv) {
    const src = String(dv.v1 || '');
    if (src.startsWith('=')) { const rg = parseRangeName(src.slice(1).replace(/\$/g, '')); if (rg) { const sh = rg.sheet ? this.wb.getSheet(rg.sheet) : sheet; if (!sh) return []; const out = []; for (let r = rg.r1; r <= Math.min(rg.r2, rg.r1 + 1000); r++) for (let c = rg.c1; c <= Math.min(rg.c2, rg.c1 + 100); c++) { const v = sh.getValue(r, c); if (v != null && v !== '') out.push(this.fmt(sh, r, c).text); } return out; } return []; }
    return src.split(',').map((x) => x.trim()).filter((x) => x !== '');
  },
  validate(sheet, r, c, dv, value) {
    if (value == null || value === '') return dv.allowBlank !== false;
    const evalExpr = (expr) => { if (expr == null || expr === '') return null; const s = String(expr); if (!s.startsWith('=')) { const p = parseInput(s); return p.value; } try { const ast = parse(shiftFormula(s, r - dv.r1, c - dv.c1)); const eng = this.wb.engine; let v = eng.evalNode(ast, eng.makeContext(sheet, r, c, [])); if (v && v.isCell) v = v.value(); else if (v && v.toArray) v = v.toArray()[0][0]; if (Array.isArray(v)) v = v[0][0]; return v; } catch { return null; } };
    const cmp = (x) => { const a = evalExpr(dv.v1), b = evalExpr(dv.v2); if (typeof a !== 'number') return true; switch (dv.op) { case 'between': return typeof b === 'number' && x >= Math.min(a, b) && x <= Math.max(a, b); case 'notBetween': return typeof b === 'number' && (x < Math.min(a, b) || x > Math.max(a, b)); case 'eq': return x === a; case 'ne': return x !== a; case 'gt': return x > a; case 'lt': return x < a; case 'ge': return x >= a; case 'le': return x <= a; } return true; };
    switch (dv.type) {
      case 'list': { const vals = this.validationListValues(sheet, dv).map((v) => v.toLowerCase()); const text = typeof value === 'number' ? formatValue(value, 'General').text : String(value); return vals.includes(text.toLowerCase()) || vals.includes(String(value).toLowerCase()); }
      case 'whole': return typeof value === 'number' && Number.isInteger(value) && cmp(value);
      case 'decimal': return typeof value === 'number' && cmp(value);
      case 'date': return typeof value === 'number' && cmp(value);
      case 'textLength': return cmp(String(value).length);
      case 'custom': { const v = evalExpr(dv.v1); return v === true || (typeof v === 'number' && v !== 0); }
    }
    return true;
  },

  // ---------- find / replace ----------
  matches(text, needle, opts) { if (!needle) return false; const a = opts.matchCase ? text : text.toLowerCase(), b = opts.matchCase ? needle : needle.toLowerCase(); return opts.entire ? a === b : a.includes(b); },
  cellSearchText(sheet, r, c, opts) { const cell = sheet.getCell(r, c); if (!cell) return null; if (opts.inFormulas && cell.f) return cell.f; const d = this.fmt(sheet, r, c).text; return d === '' ? null : d; },
  *iterCells(opts, start = null) {
    const sheets = opts.allSheets ? this.wb.sheets : [this.sheet];
    const startIdx = start ? sheets.indexOf(start.sheet) : 0;
    for (let si = 0; si < sheets.length; si++) {
      const sheet = sheets[(startIdx + si) % sheets.length];
      const keys = [...sheet.cells.keys()].map(unkey).sort((a, b) => a[0] - b[0] || a[1] - b[1]);
      for (const [r, c] of keys) { if (start && si === 0 && (r < start.r || (r === start.r && c <= start.c))) continue; const text = this.cellSearchText(sheet, r, c, opts); if (text != null) yield { sheet, r, c, text }; }
    }
    if (start) { const sheet = start.sheet; const keys = [...sheet.cells.keys()].map(unkey).sort((a, b) => a[0] - b[0] || a[1] - b[1]); for (const [r, c] of keys) { if (r > start.r || (r === start.r && c > start.c)) break; const text = this.cellSearchText(sheet, r, c, opts); if (text != null) yield { sheet, r, c, text }; } }
  },
  findNext(needle, opts) { const a = this.sheet.view.active; for (const hit of this.iterCells(opts, { sheet: this.sheet, r: a.r, c: a.c })) if (this.matches(hit.text, needle, opts)) { this.gotoHit(hit); return true; } return false; },
  findAll(needle, opts) { const out = []; for (const hit of this.iterCells(opts)) if (this.matches(hit.text, needle, opts)) out.push(hit); return out; },
  gotoHit(hit) { if (hit.sheet !== this.sheet) this.switchSheet(this.wb.sheets.indexOf(hit.sheet)); this.select(hit.r, hit.c); },
  replaceIn(sheet, r, c, needle, repl, opts) { const cell = sheet.getCell(r, c); if (!cell) return false; const src = cell.f ? cell.f : (typeof cell.v === 'string' ? cell.v : this.fmt(sheet, r, c).text); if (!this.matches(src, needle, opts)) return false; let out; if (opts.entire) out = repl; else { const re = new RegExp(needle.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), opts.matchCase ? 'g' : 'gi'); out = src.replace(re, repl); } sheet.setInput(r, c, out); return true; },
  replaceCurrent(needle, repl, opts) { const s = this.sheet; const a = s.view.active; let done = false; const text = this.cellSearchText(s, a.r, a.c, opts); if (text != null && this.matches(text, needle, opts)) { this.transact('Replace', () => { done = this.replaceIn(s, a.r, a.c, needle, repl, opts); }); } this.findNext(needle, opts); return done; },
  replaceAll(needle, repl, opts) { let n = 0; const hits = this.findAll(needle, opts); if (!hits.length) return 0; this.transact('Replace all', () => { for (const hit of hits) if (this.replaceIn(hit.sheet, hit.r, hit.c, needle, repl, opts)) n++; }); return n; },

  // ---------- data tools ----------
  removeDuplicates(rg, cols, hasHeader) {
    const s = this.sheet; const seen = new Set(); const dupRows = [];
    for (let r = rg.r1 + (hasHeader ? 1 : 0); r <= rg.r2; r++) { const key = JSON.stringify(cols.map((c) => { const v = s.getValue(r, c); return typeof v === 'string' ? v.toLowerCase() : v; })); if (seen.has(key)) dupRows.push(r); else seen.add(key); }
    if (dupRows.length) this.transact('Remove duplicates', () => { for (let i = dupRows.length - 1; i >= 0; i--) { const r = dupRows[i]; const wholeRow = rg.c1 === 0 && rg.c2 >= MAX_COLS - 1; if (wholeRow) s.deleteRows(r, 1); else { s.moveRange({ r1: r + 1, c1: rg.c1, r2: Math.max(r + 1, s.usedBounds().r2), c2: rg.c2 }, -1, 0); } } });
    return dupRows.length;
  },
  textToColumns(delim) {
    const s = this.sheet; const rg = s.view.ranges[0]; const c = rg.c1; const r2 = Math.min(rg.r2, s.usedBounds().r2);
    this.transact('Text to columns', () => { for (let r = rg.r1; r <= r2; r++) { const v = s.getValue(r, c); if (typeof v !== 'string') continue; const parts = v.split(delim); parts.forEach((p, i) => s.setInput(r, c + i, p.trim())); } });
  },
  insertCellsShift(rg, dir) {
    const s = this.sheet; const ub = s.usedBounds();
    this.transact('Insert cells', () => {
      if (dir === 'down') { const h0 = rg.r2 - rg.r1 + 1; if (ub.r2 >= rg.r1) s.moveRange({ r1: rg.r1, c1: rg.c1, r2: ub.r2, c2: rg.c2 }, h0, 0); }
      else { const w0 = rg.c2 - rg.c1 + 1; if (ub.c2 >= rg.c1) s.moveRange({ r1: rg.r1, c1: rg.c1, r2: rg.r2, c2: ub.c2 }, 0, w0); }
    });
  },
  deleteCellsShift(rg, dir) {
    const s = this.sheet; const ub = s.usedBounds();
    this.transact('Delete cells', () => {
      s.clearRange(rg, 'all');
      if (dir === 'up') { const h0 = rg.r2 - rg.r1 + 1; if (ub.r2 > rg.r2) s.moveRange({ r1: rg.r2 + 1, c1: rg.c1, r2: ub.r2, c2: rg.c2 }, -h0, 0); }
      else { const w0 = rg.c2 - rg.c1 + 1; if (ub.c2 > rg.c2) s.moveRange({ r1: rg.r1, c1: rg.c2 + 1, r2: rg.r2, c2: ub.c2 }, 0, -w0); }
    });
  },
  autosum() {
    const s = this.sheet; const v = s.view; const rg = v.ranges[0];
    const isNum = (r, c) => typeof s.getValue(r, c) === 'number';
    if (rg.r1 !== rg.r2 || rg.c1 !== rg.c2) {
      // sum each column of the selection into the row below (or right if single row)
      this.transact('AutoSum', () => { if (rg.r1 === rg.r2) s.setFormula(rg.r1, rg.c2 + 1, `=SUM(${rangeName(rg)})`); else for (let c = rg.c1; c <= rg.c2; c++) { let r2 = rg.r2; while (r2 >= rg.r1 && s.getValue(r2, c) == null) r2--; const hasNumbers = (() => { for (let r = rg.r1; r <= r2; r++) if (isNum(r, c)) return true; return false; })(); if (!hasNumbers && !(rg.r2 - rg.r1 >= 1)) continue; s.setFormula(rg.r2 + (s.getValue(rg.r2, c) != null ? 1 : 0), c, `=SUM(${rangeName({ r1: rg.r1, c1: c, r2: s.getValue(rg.r2, c) != null ? rg.r2 : rg.r2 - 1, c2: c })})`); } });
      return;
    }
    const { r, c } = v.active;
    let r1 = r - 1; while (r1 >= 0 && isNum(r1, c)) r1--; r1++;
    if (r1 < r) { this.editor.start('edit', `=SUM(${rangeName({ r1, c1: c, r2: r - 1, c2: c })})`); return; }
    let c1 = c - 1; while (c1 >= 0 && isNum(r, c1)) c1--; c1++;
    if (c1 < c) { this.editor.start('edit', `=SUM(${rangeName({ r1: r, c1, r2: r, c2: c - 1 })})`); return; }
    this.editor.start('edit', '=SUM(');
  },

  // ---------- sheets ----------
  renderTabs() {
    const tabs = document.getElementById('tabs'); tabs.innerHTML = '';
    this.wb.sheets.forEach((s, i) => {
      const tab = h('div', { class: 'tab' + (i === this.wb.activeSheetIndex ? ' active' : ''), title: s.name }, s.name, s.tabColor ? h('div', { class: 'tabcolor', style: `background:${s.tabColor}` }) : null);
      tab.addEventListener('mousedown', (e) => { if (e.button === 0) { if (this.editor.isEditing && this.editor.isFormula()) { this.switchSheet(i, true); return; } if (this.editor.isEditing) this.editor.commit(); this.switchSheet(i); } });
      tab.addEventListener('dblclick', () => this.renameSheetInline(tab, s));
      tab.addEventListener('contextmenu', (e) => { e.preventDefault(); this.switchSheet(i); this.showContextMenu(this.tabContextItems(s), e.clientX, e.clientY); });
      tabs.appendChild(tab);
    });
    const active = tabs.children[this.wb.activeSheetIndex]; if (active) active.scrollIntoView({ inline: 'nearest' });
  },
  renameSheetInline(tab, s) {
    const inp = h('input', { type: 'text', value: s.name });
    tab.textContent = ''; tab.appendChild(inp); inp.focus(); inp.select();
    const done = (save) => { if (save && inp.value.trim() && inp.value.trim() !== s.name) { try { this.transact('Rename sheet', () => this.wb.renameSheet(s, inp.value)); } catch (e) { this.toast(e.message); } } this.renderTabs(); this.focusGrid(); };
    inp.addEventListener('keydown', (e) => { e.stopPropagation(); if (e.key === 'Enter') done(true); if (e.key === 'Escape') done(false); });
    inp.addEventListener('blur', () => done(true));
  },
  switchSheet(i, keepEditing = false) {
    if (i < 0 || i >= this.wb.sheets.length) return;
    if (!keepEditing && this.editor.isEditing) this.editor.commit();
    this.wb.setActiveSheet(i);
    this.charts.select(null);
    this.renderTabs(); this.normalizeView();
    this.grid.computeLayout(); this.grid.invalidate(); this.updateFormulaBar(); this.refreshToolbar(); this.updateStatus(); this.updateZoomUi();
    if (!keepEditing) this.focusGrid();
  },
  switchToSheetId(id) { const i = this.wb.sheets.findIndex((s) => s.id === id); if (i >= 0 && i !== this.wb.activeSheetIndex) this.switchSheet(i, true); },
  addSheet() { let s; this.transact('New sheet', () => { s = this.wb.addSheet(undefined, this.wb.activeSheetIndex + 1); }); this.switchSheet(this.wb.sheets.indexOf(s)); },
  tabContextItems(s) {
    const idx = this.wb.sheets.indexOf(s);
    return [
      { label: 'Insert sheet', run: () => this.addSheet() },
      { label: 'Delete sheet', disabled: this.wb.sheets.length <= 1, run: async () => { const used = !s.usedBounds().empty; if (used && !(await this.dialogs.confirm('Delete sheet', `Delete "${s.name}" and all its data?`))) return; this.transact('Delete sheet', () => this.wb.removeSheet(s)); this.switchSheet(Math.min(idx, this.wb.sheets.length - 1)); } },
      { label: 'Rename sheet', run: () => { const tab = document.getElementById('tabs').children[idx]; if (tab) this.renameSheetInline(tab, s); } },
      { label: 'Duplicate sheet', run: () => { let ns; this.transact('Duplicate sheet', () => { ns = this.wb.duplicateSheet(s); }); this.switchSheet(this.wb.sheets.indexOf(ns)); } },
      { sep: true },
      { label: 'Move left', disabled: idx === 0, run: () => { this.transact('Move sheet', () => this.wb.moveSheet(s, idx - 1)); this.switchSheet(idx - 1); } },
      { label: 'Move right', disabled: idx >= this.wb.sheets.length - 1, run: () => { this.transact('Move sheet', () => this.wb.moveSheet(s, idx + 1)); this.switchSheet(idx + 1); } },
      { sep: true },
      { label: 'Tab color…', run: () => { const tab = document.getElementById('tabs').children[idx]; this.dialogs.colorPicker(tab || document.getElementById('tabs'), s.tabColor, (c) => { this.transact('Tab color', () => s.setTabColor(c)); this.renderTabs(); }); } },
    ];
  },

  // ---------- context menus ----------
  cellContextItems() {
    const s = this.sheet; const { r, c } = s.view.active; const cell = s.getCell(r, c);
    return [{ cmd: 'edit.cut' }, { cmd: 'edit.copy' }, { cmd: 'edit.paste' }, { cmd: 'edit.pasteSpecial' }, { sep: true }, { cmd: 'insert.cellsDown' }, { cmd: 'edit.deleteCellsUp' }, { cmd: 'insert.rowsAbove' }, { cmd: 'insert.colsLeft' }, { cmd: 'edit.deleteRows' }, { cmd: 'edit.deleteCols' }, { sep: true }, { cmd: 'edit.clearContents' }, { cmd: 'edit.clearFormats' }, { sep: true }, { cmd: 'data.filter' }, { cmd: 'data.sortAsc' }, { cmd: 'data.sortDesc' }, { sep: true }, { label: cell && cell.note ? 'Edit note' : 'Insert note', run: () => this.dialogs.note(r, c) }, ...(cell && cell.note ? [{ label: 'Delete note', run: () => this.transact('Delete note', () => s.setNote(r, c, null)) }] : []), { cmd: 'insert.hyperlink' }, ...(cell && cell.hyperlink ? [{ label: 'Open hyperlink', run: () => this.openHyperlink(cell.hyperlink) }] : []), { sep: true }, { cmd: 'format.cells' }, { cmd: 'format.cf' }, { cmd: 'data.validation' }, { cmd: 'insert.name' }];
  },
  colContextItems() { return [{ cmd: 'edit.cut' }, { cmd: 'edit.copy' }, { cmd: 'edit.paste' }, { sep: true }, { cmd: 'insert.colsLeft' }, { cmd: 'insert.colsRight' }, { cmd: 'edit.deleteCols' }, { cmd: 'edit.clearContents' }, { sep: true }, { cmd: 'format.colWidth' }, { cmd: 'format.autofitCols' }, { cmd: 'view.hideCols' }, { cmd: 'view.unhideCols' }, { sep: true }, { cmd: 'data.sortAsc' }, { cmd: 'data.sortDesc' }, { cmd: 'format.cells' }]; },
  rowContextItems() { return [{ cmd: 'edit.cut' }, { cmd: 'edit.copy' }, { cmd: 'edit.paste' }, { sep: true }, { cmd: 'insert.rowsAbove' }, { cmd: 'insert.rowsBelow' }, { cmd: 'edit.deleteRows' }, { cmd: 'edit.clearContents' }, { sep: true }, { cmd: 'format.rowHeight' }, { cmd: 'format.autofitRows' }, { cmd: 'view.hideRows' }, { cmd: 'view.unhideRows' }, { sep: true }, { cmd: 'format.cells' }]; },
  chartContextItems(id) { const s = this.sheet; return [{ label: 'Edit chart…', run: () => { const ch = s.charts.find((x) => x.id === id); if (ch) this.dialogs.chart(ch); } }, { label: 'Chart type', sub: CHART_TYPES.map((t) => ({ label: t.label, run: () => this.transact('Chart type', () => s.updateChart(id, { type: t.id })) })) }, { label: 'Save as PNG', run: () => this.charts.exportPng(id) }, { sep: true }, { label: 'Delete chart', run: () => { this.charts.select(id); this.charts.deleteSelected(); } }]; },
  openHyperlink(url) { if (url.startsWith('#')) { this.goTo(url.slice(1)); return; } if (!/^[a-z][a-z0-9+.-]*:/i.test(url)) url = 'https://' + url; window.open(url, '_blank', 'noopener'); },
  onToolbarSelect(id, value) {
    if (id === 'fontName') this.applyStyle({ font: value === 'Calibri' ? null : value });
    else if (id === 'fontSize') this.applyStyle({ size: +value === 11 ? null : +value });
    else if (id === 'numFmtSel') { if (value === 'Custom…') this.dialogs.customNumber(); else { const p = FORMAT_PRESETS.find((x) => x.name === value); if (p) this.applyNumFmt(p.code); } }
  },

  // ---------- zoom ----------
  setZoom(z) { z = Math.max(0.25, Math.min(4, Math.round(z * 100) / 100)); this.sheet.view.zoom = z; this.grid.measureCache.clear(); this.grid.invalidate(); this.updateZoomUi(); this.scheduleAutosave(); },
  updateZoomUi() { const z = this.sheet.view.zoom || 1; document.getElementById('zoomRange').value = Math.round(z * 100); document.getElementById('zoomLabel').textContent = Math.round(z * 100) + '%'; },

  // ---------- files ----------
  newWorkbook(wb = Workbook.blank(1), name = null) {
    this.wb = wb; if (name) wb.name = name;
    this.wb.formatter = (s, r, c) => this.fmt(s, r, c);
    this.wb.on('change', () => { this.renderTabs(); });
    this.wb.on('sheets', () => this.renderTabs());
    document.getElementById('fileName').value = wb.name;
    this.editor.isEditing = false; this.editor.ta.style.display = 'none';
    this.grid.copyRange = null; this.clipboard.internal = null; this.painter = null;
    for (const n of this.charts.nodes.values()) n.el.remove(); this.charts.nodes.clear();
    this.renderTabs(); this.normalizeView(); this.grid.computeLayout(); this.afterChange(); this.updateZoomUi(); this.grid.invalidate();
    this.wb.dirty = false; this.updateTitle();
  },
  async openFile(file) {
    const name = file.name; const lower = name.toLowerCase();
    try {
      this.setStatus(`Opening ${name}…`);
      if (lower.endsWith('.xlsx') || lower.endsWith('.xlsm') || lower.endsWith('.xltx')) { const wb = await importXlsx(new Uint8Array(await file.arrayBuffer()), name); this.newWorkbook(wb); }
      else if (lower.endsWith('.json')) { const j = JSON.parse(await file.text()); if (j.app !== 'officetable') throw new Error('Not an OfficeTable JSON workbook'); this.newWorkbook(Workbook.fromJSON(j)); }
      else { const text = await file.text(); const rows = parseCSV(text, lower.endsWith('.tsv') ? '\t' : null); const wb = Workbook.blank(1); wb.name = name.replace(/\.[^.]+$/, ''); wb.sheets[0].name = wb.name.slice(0, 31).replace(/[\\/?*[\]:]/g, '_') || 'Sheet1'; wb.transact('import', () => rows.forEach((row, r) => row.forEach((v, c) => { if (v !== '') wb.sheets[0].setInput(r, c, v); }))); wb.undoStack.length = 0; this.newWorkbook(wb); this.transact('Autofit', () => { for (let c = 0; c <= Math.min(50, wb.sheets[0].usedBounds().c2); c++) wb.sheets[0].setColWidth(c, Math.max(wb.sheets[0].colAxis.default, Math.min(300, this.grid.autofitCol(c)))); }); wb.undoStack.length = 0; wb.dirty = false; }
      this.toast(`Opened ${name}`); this.setStatus('Ready'); this.updateTitle();
    } catch (e) { console.error(e); this.setStatus('Ready'); this.dialogs.alert('Could not open file', `${name}\n\n${e.message || e}`); }
  },
  async saveXlsx() {
    try { const bytes = await exportXlsx(this.wb); await this.saveBytes(new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), `${this.wb.name}.xlsx`, [{ description: 'Excel workbook', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] } }]); this.wb.dirty = false; this.updateTitle(); }
    catch (e) { if (e && e.name === 'AbortError') return; console.error(e); this.toast('Save failed: ' + (e.message || e)); }
  },
  async saveJson() { try { const blob = new Blob([JSON.stringify(this.wb.toJSON())], { type: 'application/json' }); await this.saveBytes(blob, `${this.wb.name}.officetable.json`, [{ description: 'OfficeTable workbook', accept: { 'application/json': ['.json'] } }]); this.wb.dirty = false; this.updateTitle(); } catch (e) { if (e && e.name !== 'AbortError') this.toast('Save failed: ' + e.message); } },
  async exportCsv(delim = ',') {
    const s = this.sheet; const db = s.dataBounds(); const rows = [];
    if (db) for (let r = 0; r <= db.r2; r++) { const row = []; for (let c = 0; c <= db.c2; c++) row.push(this.fmt(s, r, c).text); rows.push(row); }
    const blob = new Blob(['﻿' + toCSV(rows, delim)], { type: delim === '\t' ? 'text/tab-separated-values' : 'text/csv' });
    try { await this.saveBytes(blob, `${s.name}.${delim === '\t' ? 'tsv' : 'csv'}`); } catch (e) { if (e && e.name !== 'AbortError') this.toast('Export failed: ' + e.message); }
  },
  async saveBytes(blob, suggestedName, types) {
    if (window.showSaveFilePicker && !navigator.userAgent.includes('HeadlessChrome')) {
      const handle = await window.showSaveFilePicker({ suggestedName, types });
      const w = await handle.createWritable(); await w.write(blob); await w.close();
      if (suggestedName.endsWith('.xlsx')) { this.wb.name = handle.name.replace(/\.xlsx$/i, ''); document.getElementById('fileName').value = this.wb.name; }
      this.toast(`Saved ${handle.name}`); return;
    }
    this.downloadBlob(blob, suggestedName); this.toast(`Downloaded ${suggestedName}`);
  },
  downloadBlob(blob, name) { const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000); },
  print() {
    const s = this.sheet; const rg0 = s.view.ranges[0]; const multi = rg0.r1 !== rg0.r2 || rg0.c1 !== rg0.c2;
    const db = s.dataBounds(); if (!db && !multi) { this.toast('Nothing to print'); return; }
    const rg = multi ? rg0 : { r1: 0, c1: 0, r2: db.r2, c2: db.c2 };
    let html = `<html><head><title>${escapeHtml(this.wb.name)} - ${escapeHtml(s.name)}</title><style>body{font-family:Calibri,Arial,sans-serif;font-size:11pt}table{border-collapse:collapse}td{border:1px solid #bbb;padding:2px 5px;white-space:pre;vertical-align:bottom}h3{font-weight:normal;color:#555;font-size:10pt}</style></head><body><h3>${escapeHtml(this.wb.name)} — ${escapeHtml(s.name)}</h3><table>`;
    for (let r = rg.r1; r <= rg.r2; r++) {
      if (s.rowAxis.isHidden(r)) continue;
      html += `<tr style="height:${s.rowAxis.size(r)}px">`;
      for (let c = rg.c1; c <= rg.c2; c++) {
        if (s.colAxis.isHidden(c)) continue;
        const m = s.mergeAt(r, c); if (m && (m.r1 !== r || m.c1 !== c)) continue;
        const st = s.getStyle(r, c); const d = this.fmt(s, r, c); const v = s.getValue(r, c);
        const css = [`width:${s.colAxis.size(c)}px`, `text-align:${st.align && st.align !== 'general' ? st.align : typeof v === 'number' ? 'right' : typeof v === 'boolean' || isErr(v) ? 'center' : 'left'}`];
        if (st.bold) css.push('font-weight:bold'); if (st.italic) css.push('font-style:italic'); if (st.underline) css.push('text-decoration:underline'); if (st.color) css.push(`color:${st.color}`); if (st.fill) css.push(`background:${st.fill}`); if (st.font) css.push(`font-family:${st.font}`); if (st.size) css.push(`font-size:${st.size}pt`); if (st.wrap) css.push('white-space:pre-wrap');
        html += `<td${m ? ` rowspan="${m.r2 - m.r1 + 1}" colspan="${m.c2 - m.c1 + 1}"` : ''} style="${css.join(';')}">${escapeHtml(d.text)}</td>`;
      }
      html += '</tr>';
    }
    html += '</table></body></html>';
    const w = window.open('', '_blank'); if (!w) { this.toast('Pop-up blocked'); return; }
    w.document.write(html); w.document.close(); w.focus(); setTimeout(() => w.print(), 300);
  },

  // ---------- storage ----------
  scheduleAutosave: null,
  autosave() { if (!this.settings.autosave) return; try { localStorage.setItem(STORAGE_KEY, JSON.stringify(this.wb.toJSON())); } catch (e) { /* quota */ } },
  saveSettings() { try { localStorage.setItem(SETTINGS_KEY, JSON.stringify(this.settings)); } catch { /* ignore */ } },
  restore() {
    try { const s = localStorage.getItem(SETTINGS_KEY); if (s) Object.assign(this.settings, JSON.parse(s)); } catch { /* ignore */ }
    if (!this.settings.autosave) return false;
    try { const j = localStorage.getItem(STORAGE_KEY); if (!j) return false; const wb = Workbook.fromJSON(JSON.parse(j)); this.newWorkbook(wb); return true; } catch (e) { console.warn('restore failed', e); return false; }
  },

  // ---------- commands ----------
  run(name, e) {
    const cmd = this.commands[name]; if (!cmd) return;
    if (cmd.enabled && !cmd.enabled()) return;
    try { const r = cmd.run(e); if (r && r.catch) r.catch((err) => { console.error(err); this.toast(err.message || String(err)); }); } catch (err) { console.error(err); this.toast(err.message || String(err)); }
  },
  defineCommands() {
    const s = () => this.sheet; const sel = () => this.sheet.view.ranges; const C = this.commands;
    const def = (name, label, sc, run, extra = {}) => { C[name] = { label, sc, run, ...extra }; };
    const rowsOf = () => { const rg = sel()[0]; return { at: rg.r1, n: rg.r2 - rg.r1 + 1 }; };
    const colsOf = () => { const rg = sel()[0]; return { at: rg.c1, n: rg.c2 - rg.c1 + 1 }; };
    def('file.new', 'New workbook', 'Ctrl+N', async () => { if (this.wb.dirty && !(await this.dialogs.confirm('New workbook', 'Discard unsaved changes in the current workbook?'))) return; this.newWorkbook(Workbook.blank(1)); });
    def('file.open', 'Open…', 'Ctrl+O', () => { document.getElementById('fileInput').value = ''; document.getElementById('fileInput').click(); });
    def('file.saveXlsx', 'Save as .xlsx', 'Ctrl+S', () => this.saveXlsx());
    def('file.saveJson', 'Save as OfficeTable JSON (keeps charts)', null, () => this.saveJson());
    def('file.exportCsv', 'Export sheet as CSV', null, () => this.exportCsv(','));
    def('file.exportTsv', 'Export sheet as TSV', null, () => this.exportCsv('\t'));
    def('file.print', 'Print…', 'Ctrl+P', () => this.print());
    def('file.autosave', 'Keep a copy in this browser', null, () => { this.settings.autosave = !this.settings.autosave; this.saveSettings(); if (this.settings.autosave) this.autosave(); else localStorage.removeItem(STORAGE_KEY); this.toast(this.settings.autosave ? 'Autosave on' : 'Autosave off'); }, { checked: () => this.settings.autosave });
    def('file.clearStorage', 'Forget saved copy', null, () => { localStorage.removeItem(STORAGE_KEY); this.toast('Saved copy removed'); });
    def('edit.undo', 'Undo', 'Ctrl+Z', () => { if (this.editor.isEditing) this.editor.cancel(); if (this.wb.undo()) { this.normalizeView(); this.afterChange(); this.renderTabs(); } }, { enabled: () => this.wb.canUndo() });
    def('edit.redo', 'Redo', 'Ctrl+Y', () => { if (this.wb.redo()) { this.normalizeView(); this.afterChange(); this.renderTabs(); } }, { enabled: () => this.wb.canRedo() });
    def('edit.cut', 'Cut', 'Ctrl+X', () => this.clipboard.copy(true));
    def('edit.copy', 'Copy', 'Ctrl+C', () => this.clipboard.copy(false));
    def('edit.paste', 'Paste', 'Ctrl+V', () => this.clipboard.paste(null));
    def('edit.pasteSpecial', 'Paste special…', 'Ctrl+Alt+V', () => this.dialogs.pasteSpecial());
    def('edit.pasteValues', 'Paste values only', 'Ctrl+Shift+V', () => this.clipboard.paste(null, { mode: 'values' }));
    def('edit.fillDown', 'Fill down', 'Ctrl+D', () => this.transact('Fill down', () => sel().forEach((rg) => fillDirection(s(), { ...rg, r2: Math.min(rg.r2, rg.r1 + 100000) }, 'down'))));
    def('edit.fillRight', 'Fill right', 'Ctrl+R', () => this.transact('Fill right', () => sel().forEach((rg) => fillDirection(s(), { ...rg, c2: Math.min(rg.c2, rg.c1 + 10000) }, 'right'))));
    def('edit.clearContents', 'Clear contents', 'Delete', () => this.transact('Clear', () => sel().forEach((rg) => s().clearRange(rg, 'contents'))));
    def('edit.clearFormats', 'Clear formats', null, () => this.transact('Clear formats', () => sel().forEach((rg) => s().clearRange(rg, 'formats'))));
    def('edit.clearNotes', 'Clear notes', null, () => this.transact('Clear notes', () => sel().forEach((rg) => s().clearRange(rg, 'notes'))));
    def('edit.clearAll', 'Clear all', null, () => this.transact('Clear all', () => sel().forEach((rg) => { s().clearRange(rg, 'all'); s().unmerge(rg); })));
    def('edit.deleteRows', 'Delete rows', null, () => { const { at, n } = rowsOf(); this.transact('Delete rows', () => s().deleteRows(at, n)); this.select(at, this.sheet.view.active.c); });
    def('edit.deleteCols', 'Delete columns', null, () => { const { at, n } = colsOf(); this.transact('Delete columns', () => s().deleteCols(at, n)); this.select(this.sheet.view.active.r, at); });
    def('edit.deleteCellsUp', 'Delete cells, shift up', null, () => this.deleteCellsShift(sel()[0], 'up'));
    def('edit.deleteCellsLeft', 'Delete cells, shift left', null, () => this.deleteCellsShift(sel()[0], 'left'));
    def('edit.find', 'Find…', 'Ctrl+F', () => this.dialogs.findReplace(false));
    def('edit.replace', 'Replace…', 'Ctrl+H', () => this.dialogs.findReplace(true));
    def('edit.goto', 'Go to…', 'Ctrl+G', () => this.dialogs.goTo());
    def('edit.selectAll', 'Select all', 'Ctrl+A', () => this.selectAll());
    def('view.freezeSel', 'Freeze panes at selection', null, () => { const a = this.sheet.view.active; const fz = this.sheet.freeze; if (fz.rows || fz.cols) this.transact('Unfreeze', () => s().setFreeze(0, 0)); else this.transact('Freeze panes', () => s().setFreeze(Math.max(0, a.r - 0), Math.max(0, a.c))); this.grid.scrollRow = Math.max(this.grid.scrollRow, this.sheet.freeze.rows); this.grid.scrollCol = Math.max(this.grid.scrollCol, this.sheet.freeze.cols); }, { checked: () => !!(this.sheet.freeze.rows || this.sheet.freeze.cols) });
    def('view.freezeRow', 'Freeze top row', null, () => this.transact('Freeze', () => s().setFreeze(1, s().freeze.cols)));
    def('view.freezeCol', 'Freeze first column', null, () => this.transact('Freeze', () => s().setFreeze(s().freeze.rows, 1)));
    def('view.unfreeze', 'Unfreeze panes', null, () => this.transact('Unfreeze', () => s().setFreeze(0, 0)));
    def('view.gridlines', 'Gridlines', null, () => this.transact('Gridlines', () => s().setGridlines(!s().showGridlines)), { checked: () => this.sheet.showGridlines });
    def('view.formulas', 'Show formulas', 'Ctrl+`', () => { this.settings.showFormulas = !this.settings.showFormulas; this.grid.invalidate(); }, { checked: () => this.settings.showFormulas });
    def('view.zoomIn', 'Zoom in', 'Ctrl+Wheel', () => this.setZoom(this.grid.zoom * 1.1));
    def('view.zoomOut', 'Zoom out', null, () => this.setZoom(this.grid.zoom / 1.1));
    def('view.zoom100', 'Zoom 100%', null, () => this.setZoom(1));
    def('view.hideRows', 'Hide rows', 'Ctrl+9', () => this.transact('Hide rows', () => this.selectedRows().forEach((r) => s().setRowHidden(r, true))));
    def('view.hideCols', 'Hide columns', 'Ctrl+0', () => this.transact('Hide columns', () => this.selectedCols().forEach((c) => s().setColHidden(c, true))));
    def('view.unhideRows', 'Unhide rows', 'Ctrl+Shift+9', () => { const rg = sel()[0]; this.transact('Unhide rows', () => { for (let r = Math.max(0, rg.r1 - 1); r <= Math.min(rg.r2 + 1, MAX_ROWS - 1); r++) if (s().rowAxis.hidden.has(r)) s().setRowHidden(r, false); if (rg.r1 === 0 && rg.r2 >= MAX_ROWS - 1 && rg.c1 === 0) [...s().rowAxis.hidden].forEach((r) => s().setRowHidden(r, false)); }); });
    def('view.unhideCols', 'Unhide columns', 'Ctrl+Shift+0', () => { const rg = sel()[0]; this.transact('Unhide columns', () => { for (let c = Math.max(0, rg.c1 - 1); c <= Math.min(rg.c2 + 1, MAX_COLS - 1); c++) if (s().colAxis.hidden.has(c)) s().setColHidden(c, false); if (rg.c1 === 0 && rg.c2 >= MAX_COLS - 1 && rg.r1 === 0) [...s().colAxis.hidden].forEach((c) => s().setColHidden(c, false)); }); });
    def('insert.rowsAbove', 'Insert rows above', null, () => { const { at, n } = rowsOf(); this.transact('Insert rows', () => s().insertRows(at, n)); });
    def('insert.rowsBelow', 'Insert rows below', null, () => { const { at, n } = rowsOf(); this.transact('Insert rows', () => s().insertRows(at + n, n)); });
    def('insert.colsLeft', 'Insert columns left', null, () => { const { at, n } = colsOf(); this.transact('Insert columns', () => s().insertCols(at, n)); });
    def('insert.colsRight', 'Insert columns right', null, () => { const { at, n } = colsOf(); this.transact('Insert columns', () => s().insertCols(at + n, n)); });
    def('insert.cellsDown', 'Insert cells, shift down', null, () => this.insertCellsShift(sel()[0], 'down'));
    def('insert.cellsRight', 'Insert cells, shift right', null, () => this.insertCellsShift(sel()[0], 'right'));
    def('insert.cells', 'Insert cells…', 'Ctrl++', () => this.dialogs.insertCells().then((m) => { if (!m) return; if (m === 'rows') this.run('insert.rowsAbove'); else if (m === 'cols') this.run('insert.colsLeft'); else this.insertCellsShift(sel()[0], m); }));
    def('edit.deleteCells', 'Delete cells…', 'Ctrl+-', () => this.dialogs.deleteCells().then((m) => { if (!m) return; if (m === 'rows') this.run('edit.deleteRows'); else if (m === 'cols') this.run('edit.deleteCols'); else this.deleteCellsShift(sel()[0], m); }));
    def('insert.sheet', 'New sheet', 'Shift+F11', () => this.addSheet());
    def('insert.chart', 'Chart…', null, () => this.dialogs.chart());
    def('insert.function', 'Function…', 'Shift+F3', () => this.dialogs.insertFunction());
    def('insert.name', 'Define name…', null, () => this.dialogs.nameManager());
    def('insert.note', 'Note', 'Shift+F2', () => { const a = this.sheet.view.active; this.dialogs.note(a.r, a.c); });
    def('insert.hyperlink', 'Hyperlink…', 'Ctrl+K', () => this.dialogs.hyperlink());
    def('insert.date', 'Current date', 'Ctrl+;', () => { if (this.editor.isEditing) { this.editor.insertText(formatValue(todaySerial(), 'm/d/yyyy').text); return; } const a = this.sheet.view.active; this.transact('Insert date', () => { s().setValue(a.r, a.c, todaySerial()); if (!isDateFormat(s().effectiveNumFmt(a.r, a.c))) s().applyStyle({ r1: a.r, c1: a.c, r2: a.r, c2: a.c }, { numFmt: 'm/d/yyyy' }); }); });
    def('insert.time', 'Current time', 'Ctrl+Shift+;', () => { if (this.editor.isEditing) { this.editor.insertText(formatValue(nowSerial(), 'h:mm AM/PM').text); return; } const a = this.sheet.view.active; this.transact('Insert time', () => { const t = nowSerial(); s().setValue(a.r, a.c, t - Math.floor(t)); s().applyStyle({ r1: a.r, c1: a.c, r2: a.r, c2: a.c }, { numFmt: 'h:mm AM/PM' }); }); });
    def('format.cells', 'Format cells…', 'Ctrl+1', () => this.dialogs.formatCells());
    def('format.customNumber', 'Custom number format…', null, () => this.dialogs.customNumber());
    def('format.bold', 'Bold', 'Ctrl+B', () => this.toggleStyle('bold'), { checked: () => !!this.activeStyle().bold });
    def('format.italic', 'Italic', 'Ctrl+I', () => this.toggleStyle('italic'), { checked: () => !!this.activeStyle().italic });
    def('format.underline', 'Underline', 'Ctrl+U', () => this.toggleStyle('underline'), { checked: () => !!this.activeStyle().underline });
    def('format.strike', 'Strikethrough', 'Ctrl+5', () => this.toggleStyle('strike'), { checked: () => !!this.activeStyle().strike });
    for (const [n, v, label] of [['alignLeft', 'left', 'Align left'], ['alignCenter', 'center', 'Align center'], ['alignRight', 'right', 'Align right']]) def(`format.${n}`, label, null, () => this.applyStyle({ align: this.activeStyle().align === v ? null : v }), { checked: () => this.activeStyle().align === v });
    for (const [n, v, label] of [['valignTop', 'top', 'Align top'], ['valignMiddle', 'middle', 'Align middle'], ['valignBottom', 'bottom', 'Align bottom']]) def(`format.${n}`, label, null, () => this.applyStyle({ valign: v === 'bottom' ? null : v }), { checked: () => (this.activeStyle().valign || 'bottom') === v });
    def('format.wrap', 'Wrap text', null, () => { const on = !this.activeStyle().wrap; this.applyStyle({ wrap: on || null }); if (on) this.transact('Autofit', () => this.selectedRows().slice(0, 500).forEach((r) => s().setRowHeight(r, Math.max(s().rowAxis.rawSize(r), this.grid.autofitRow(r))))); }, { checked: () => !!this.activeStyle().wrap });
    def('format.merge', 'Merge cells', null, () => { const rg = sel()[0]; if (s().mergeAt(rg.r1, rg.c1) && s().merges.some((m) => m.r1 === rg.r1 && m.c1 === rg.c1 && m.r2 === rg.r2 && m.c2 === rg.c2)) this.transact('Unmerge', () => s().unmerge(rg)); else this.transact('Merge', () => s().merge({ ...rg, r2: Math.min(rg.r2, rg.r1 + 10000), c2: Math.min(rg.c2, rg.c1 + 1000) })); this.select(rg.r1, rg.c1); }, { checked: () => !!this.sheet.mergeAt(this.sheet.view.active.r, this.sheet.view.active.c) });
    def('format.mergeAcross', 'Merge across (each row)', null, () => { const rg = sel()[0]; this.transact('Merge across', () => { for (let r = rg.r1; r <= Math.min(rg.r2, rg.r1 + 5000); r++) s().merge({ r1: r, c1: rg.c1, r2: r, c2: Math.min(rg.c2, rg.c1 + 1000) }); }); });
    def('format.unmerge', 'Unmerge cells', null, () => this.transact('Unmerge', () => sel().forEach((rg) => s().unmerge(rg))));
    def('format.colWidth', 'Column width…', null, () => this.dialogs.colWidth());
    def('format.rowHeight', 'Row height…', null, () => this.dialogs.rowHeight());
    def('format.autofitCols', 'Autofit column width', null, () => this.transact('Autofit', () => this.selectedCols().slice(0, 500).forEach((c) => s().setColWidth(c, this.grid.autofitCol(c)))));
    def('format.autofitRows', 'Autofit row height', null, () => this.transact('Autofit', () => this.selectedRows().slice(0, 5000).forEach((r) => s().setRowHeight(r, this.grid.autofitRow(r)))));
    def('format.cf', 'Conditional formatting…', null, () => this.dialogs.conditionalFormatting());
    def('format.clearCf', 'Clear conditional formatting from selection', null, () => this.transact('Clear rules', () => s().setCF(s().cf.filter((rule) => !sel().some((rg) => rangesIntersect(rg, rule))))));
    def('format.painter', 'Format painter', null, () => { if (this.painter) { this.painter = null; this.toolbarButtons['format.painter'].classList.remove('active'); this.setStatus('Ready'); return; } const a = this.sheet.view.active; this.painter = s().getStyleId(a.r, a.c); this.toolbarButtons['format.painter'].classList.add('active'); this.setStatus('Select the cells to paint with this format'); });
    def('format.borders', 'Borders', null, (e) => { const b = this.toolbarButtons['format.borders']; const r = b ? b.getBoundingClientRect() : { left: e ? e.clientX : 100, bottom: e ? e.clientY : 100 }; const presets = [['all', 'All borders'], ['outline', 'Outside borders'], ['inside', 'Inside borders'], ['top', 'Top border'], ['bottom', 'Bottom border'], ['left', 'Left border'], ['right', 'Right border'], ['topBottom', 'Top and bottom'], ['thickBottom', 'Thick bottom'], ['doubleBottom', 'Double bottom'], ['none', 'No border']]; this.showContextMenu([...presets.map(([id, label]) => ({ label, run: () => this.transact('Borders', () => sel().forEach((rg) => this.applyBorderPreset(rg, id))) })), { sep: true }, { label: 'More borders…', run: () => this.dialogs.formatCells('border') }], r.left, r.bottom + 2); });
    def('format.currency', 'Currency format', 'Ctrl+Shift+4', () => this.applyNumFmt('$#,##0.00;[Red]-$#,##0.00'));
    def('format.percent', 'Percent format', 'Ctrl+Shift+5', () => this.applyNumFmt('0%'));
    def('format.comma', 'Thousands separator', 'Ctrl+Shift+1', () => this.applyNumFmt('#,##0.00'));
    def('format.decInc', 'Increase decimals', null, () => this.adjustDecimals(1));
    def('format.decDec', 'Decrease decimals', null, () => this.adjustDecimals(-1));
    def('format.indentInc', 'Increase indent', null, () => this.applyStyle({ indent: Math.min(15, (this.activeStyle().indent || 0) + 1) }));
    def('format.indentDec', 'Decrease indent', null, () => this.applyStyle({ indent: Math.max(0, (this.activeStyle().indent || 0) - 1) || null }));
    def('format.general', 'General format', 'Ctrl+Shift+~', () => this.applyNumFmt('General'));
    def('format.date', 'Date format', 'Ctrl+Shift+3', () => this.applyNumFmt('d-mmm-yy'));
    def('format.time', 'Time format', 'Ctrl+Shift+2', () => this.applyNumFmt('h:mm AM/PM'));
    def('data.sortAsc', 'Sort A → Z', null, () => this.quickSort(true));
    def('data.sortDesc', 'Sort Z → A', null, () => this.quickSort(false));
    def('data.sort', 'Custom sort…', null, () => this.dialogs.sort());
    def('data.filter', 'Filter', 'Ctrl+Shift+L', () => { if (s().filter) this.transact('Remove filter', () => s().setFilter(null)); else { const rg = this.dataRegion(); if (rg.r1 === rg.r2 && rg.c1 === rg.c2 && s().getValue(rg.r1, rg.c1) == null) { this.toast('Select a cell inside your data first'); return; } this.transact('Filter', () => s().setFilter({ r1: rg.r1, c1: rg.c1, r2: Math.max(rg.r2, rg.r1 + 1), c2: rg.c2, conds: {} })); } }, { checked: () => !!this.sheet.filter });
    def('data.clearFilter', 'Clear filter criteria', null, () => { if (s().filter) this.transact('Clear filter', () => s().setFilter({ ...s().filter, conds: {} })); });
    def('data.validation', 'Data validation…', null, () => this.dialogs.dataValidation());
    def('data.removeDuplicates', 'Remove duplicates…', null, () => this.dialogs.removeDuplicates());
    def('data.textToColumns', 'Text to columns…', null, () => this.dialogs.textToColumns());
    def('data.trim', 'Trim whitespace', null, () => this.transact('Trim', () => sel().forEach((rg) => { const r2 = Math.min(rg.r2, s().usedBounds().r2), c2 = Math.min(rg.c2, s().usedBounds().c2); for (let r = rg.r1; r <= r2; r++) for (let c = rg.c1; c <= c2; c++) { const cell = s().getCell(r, c); if (cell && typeof cell.v === 'string' && !cell.f) { const t = cell.v.replace(/\s+/g, ' ').trim(); if (t !== cell.v) s().setValue(r, c, t); } } })));
    def('data.names', 'Name manager…', null, () => this.dialogs.nameManager());
    def('data.recalc', 'Recalculate now', 'F9', () => { this.wb.engine.recalcAll(); this.grid.invalidate(); this.toast(`Recalculated ${this.wb.engine.lastRecalcCount} formulas`); });
    def('formulas.autosum', 'AutoSum', 'Alt+=', () => this.autosum());
    def('help.shortcuts', 'Keyboard shortcuts', 'F1', () => this.dialogs.help());
    def('help.functions', 'Function list', null, () => this.dialogs.functionList());
    def('help.about', 'About OfficeTable', null, () => this.dialogs.about());
  },
  activeStyle() { const a = this.sheet.view.active; return this.sheet.getStyle(a.r, a.c); },
  quickSort(asc) {
    const s = this.sheet; const rg = this.dataRegion(); const col = s.view.active.c;
    if (rg.r1 === rg.r2) { this.toast('Nothing to sort'); return; }
    const header = this.looksLikeHeader(rg);
    this.transact('Sort', () => s.sortRange(rg, [{ col: Math.min(Math.max(col, rg.c1), rg.c2), asc }], header));
    this.selectRange(rg, { active: s.view.active });
  },

  // ---------- keyboard ----------
  onKeyDown(e) {
    const tag = (document.activeElement && document.activeElement.tagName) || '';
    const inField = ['INPUT', 'TEXTAREA', 'SELECT'].includes(tag) && document.activeElement !== this.grid.canvas;
    if (this.dialogOpen()) return;
    if (inField && (document.activeElement.id === 'nameBox')) { if (e.key === 'Enter') { e.preventDefault(); this.onNameBoxEnter(); } if (e.key === 'Escape') { this.updateFormulaBar(); this.focusGrid(); } return; }
    if (inField) return; // editor and other inputs handle their own keys
    const ctrl = e.ctrlKey || e.metaKey;
    const k = e.key;
    const dir = { ArrowDown: 'down', ArrowUp: 'up', ArrowLeft: 'left', ArrowRight: 'right' }[k];
    if (dir) { e.preventDefault(); const d = ctrl ? 'ctrl' + dir[0].toUpperCase() + dir.slice(1) : dir; if (e.shiftKey) this.extendActive(d); else this.moveActive(d); return; }
    if (k === 'Tab') { e.preventDefault(); this.moveActive(e.shiftKey ? 'left' : 'right', { enter: true }); return; }
    if (k === 'Enter') { e.preventDefault(); if (this.grid.copyRange && this.clipboard.internal && this.clipboard.internal.cut) { this.clipboard.paste(null); return; } if (e.altKey) { this.editor.start('edit'); return; } this.moveActive(e.shiftKey ? 'up' : 'down', { enter: true }); return; }
    if (k === 'Home') { e.preventDefault(); if (e.shiftKey) this.extendActive(ctrl ? 'ctrlHome' : 'home'); else this.moveActive(ctrl ? 'ctrlHome' : 'home'); return; }
    if (k === 'End') { e.preventDefault(); if (e.shiftKey) this.extendActive('ctrlEnd'); else this.moveActive(ctrl ? 'ctrlEnd' : 'end'); return; }
    if (k === 'PageDown' || k === 'PageUp') { e.preventDefault(); if (ctrl) { this.switchSheet((this.wb.activeSheetIndex + (k === 'PageDown' ? 1 : this.wb.sheets.length - 1)) % this.wb.sheets.length); return; } const d = e.altKey ? (k === 'PageDown' ? 'pageright' : 'pageleft') : (k === 'PageDown' ? 'pagedown' : 'pageup'); if (e.shiftKey) this.extendActive(d); else this.moveActive(d); return; }
    if (k === 'Escape') { e.preventDefault(); if (this.grid.copyRange) { this.grid.copyRange = null; this.grid.invalidate(); } if (this.painter) this.run('format.painter'); this.charts.select(null); this.closeMenus(); return; }
    if (k === 'Delete') { e.preventDefault(); if (this.charts.deleteSelected()) return; this.run('edit.clearContents'); return; }
    if (k === 'Backspace') { e.preventDefault(); if (this.charts.deleteSelected()) return; this.editor.start('type', ''); return; }
    if (k === 'F2') { e.preventDefault(); if (e.shiftKey) this.run('insert.note'); else this.editor.start('edit'); return; }
    if (k === 'F1') { e.preventDefault(); this.run('help.shortcuts'); return; }
    if (k === 'F3' && e.shiftKey) { e.preventDefault(); this.run('insert.function'); return; }
    if (k === 'F9') { e.preventDefault(); this.run('data.recalc'); return; }
    if (k === 'F11' && e.shiftKey) { e.preventDefault(); this.run('insert.sheet'); return; }
    if (k === 'F4') { e.preventDefault(); return; }
    if (k === ' ' && ctrl) { e.preventDefault(); const rg = this.sheet.view.ranges[0]; this.selectRange({ r1: 0, c1: rg.c1, r2: MAX_ROWS - 1, c2: rg.c2 }, { keepActive: true }); return; }
    if (k === ' ' && e.shiftKey) { e.preventDefault(); const rg = this.sheet.view.ranges[0]; this.selectRange({ r1: rg.r1, c1: 0, r2: rg.r2, c2: MAX_COLS - 1 }, { keepActive: true }); return; }
    if (e.altKey && k === '=') { e.preventDefault(); this.run('formulas.autosum'); return; }
    if (ctrl) {
      const map = { a: 'edit.selectAll', b: 'format.bold', i: 'format.italic', u: 'format.underline', '5': 'format.strike', s: 'file.saveXlsx', o: 'file.open', n: 'file.new', p: 'file.print', f: 'edit.find', h: 'edit.replace', g: 'edit.goto', d: 'edit.fillDown', r: 'edit.fillRight', k: 'insert.hyperlink', '1': 'format.cells', '9': e.shiftKey ? 'view.unhideRows' : 'view.hideRows', '0': e.shiftKey ? 'view.unhideCols' : 'view.hideCols', ';': e.shiftKey ? 'insert.time' : 'insert.date', ':': 'insert.time', '`': 'view.formulas', '~': 'format.general', '!': 'format.comma', '@': 'format.time', '#': 'format.date', '$': 'format.currency', '%': 'format.percent', '+': 'insert.cells', '=': 'insert.cells', '-': 'edit.deleteCells', '_': 'edit.deleteCells' };
      const name = map[k.length === 1 ? (e.shiftKey && /[a-z]/i.test(k) ? k.toLowerCase() : k) : k];
      if (k === 'z' || k === 'Z') { e.preventDefault(); this.run(e.shiftKey ? 'edit.redo' : 'edit.undo'); return; }
      if (k === 'y' || k === 'Y') { e.preventDefault(); this.run('edit.redo'); return; }
      if (k === 'v' || k === 'V') { if (e.shiftKey) { e.preventDefault(); this.run('edit.pasteValues'); } else if (e.altKey) { e.preventDefault(); this.run('edit.pasteSpecial'); } return; } // plain Ctrl+V: 'paste' event
      if (k === 'c' || k === 'C' || k === 'x' || k === 'X') return; // handled by copy/cut events
      if (name) { e.preventDefault(); this.run(name); return; }
      if (e.shiftKey && (k === 'L' || k === 'l')) { e.preventDefault(); this.run('data.filter'); return; }
      return;
    }
    // printable character starts editing
    if (k.length === 1 && !e.altKey) { e.preventDefault(); this.editor.start('type', k); return; }
  },
  onNameBoxEnter() {
    const nb = document.getElementById('nameBox'); const t = nb.value.trim(); if (!t) return;
    const rg = parseRangeName(t);
    if (rg) { this.goTo(t); this.focusGrid(); return; }
    if (this.wb.getName(t, this.sheet)) { this.goTo(t); this.focusGrid(); return; }
    if (/^[A-Za-z_\\][A-Za-z0-9_.\\]*$/.test(t)) { const s = this.sheet; const sel = s.view.ranges[0]; const ref = `=${quoteSheetName(s.name)}!${rangeName(sel).replace(/([A-Z]+)(\d+)/g, '$$$1$$$2')}`; this.transact('Define name', () => this.wb.setName(t, ref)); this.toast(`Defined name "${t}" = ${ref}`); this.updateFormulaBar(); this.focusGrid(); return; }
    this.toast(`"${t}" is not a valid reference or name`);
  },

  // ---------- init ----------
  init() {
    this.wb = Workbook.blank(1);
    this.wb.formatter = (s, r, c) => this.fmt(s, r, c);
    this.dialogs = createDialogs(this);
    this.grid = new GridView(this, document.getElementById('gridwrap'));
    this.editor = new CellEditor(this);
    this.clipboard = new ClipboardManager(this);
    this.charts = new ChartLayer(this, this.grid.overlay);
    this.defineCommands();
    buildMenus(this);
    this.scheduleAutosave = debounce(() => this.autosave(), 800);
    this.grid.canvas.tabIndex = 0;
    document.addEventListener('keydown', (e) => this.onKeyDown(e));
    document.addEventListener('copy', (e) => { if (this.editor.isEditing || this.dialogOpen() || ['INPUT', 'TEXTAREA'].includes(document.activeElement.tagName)) return; e.preventDefault(); this.clipboard.copy(false); if (this.clipboard.lastHtml) { e.clipboardData.setData('text/html', this.clipboard.lastHtml); e.clipboardData.setData('text/plain', this.clipboard.internal.cells.map((row) => row.map((c) => c.text).join('\t')).join('\n')); } });
    document.addEventListener('cut', (e) => { if (this.editor.isEditing || this.dialogOpen() || ['INPUT', 'TEXTAREA'].includes(document.activeElement.tagName)) return; e.preventDefault(); this.clipboard.copy(true); if (this.clipboard.lastHtml) { e.clipboardData.setData('text/html', this.clipboard.lastHtml); e.clipboardData.setData('text/plain', this.clipboard.internal.cells.map((row) => row.map((c) => c.text).join('\t')).join('\n')); } });
    document.addEventListener('paste', (e) => { if (this.editor.isEditing || this.dialogOpen() || ['INPUT', 'TEXTAREA'].includes(document.activeElement.tagName)) return; e.preventDefault(); this.clipboard.paste(e); });
    document.getElementById('fileInput').addEventListener('change', (e) => { const f = e.target.files[0]; if (f) this.openFile(f); });
    document.getElementById('fileName').addEventListener('change', (e) => { this.wb.name = e.target.value.trim() || 'Book1'; e.target.value = this.wb.name; this.updateTitle(); this.scheduleAutosave(); });
    document.getElementById('fileName').addEventListener('keydown', (e) => { e.stopPropagation(); if (e.key === 'Enter') { e.target.blur(); this.focusGrid(); } });
    document.getElementById('addSheet').addEventListener('click', () => this.addSheet());
    document.getElementById('tabPrev').addEventListener('click', () => this.switchSheet(Math.max(0, this.wb.activeSheetIndex - 1)));
    document.getElementById('tabNext').addEventListener('click', () => this.switchSheet(Math.min(this.wb.sheets.length - 1, this.wb.activeSheetIndex + 1)));
    document.getElementById('zoomRange').addEventListener('input', (e) => this.setZoom(+e.target.value / 100));
    document.getElementById('zoomIn').addEventListener('click', () => this.setZoom(this.grid.zoom + 0.1));
    document.getElementById('zoomOut').addEventListener('click', () => this.setZoom(this.grid.zoom - 0.1));
    document.getElementById('fxBtn').addEventListener('click', () => this.dialogs.insertFunction());
    document.getElementById('fbCancel').addEventListener('click', () => this.editor.cancel());
    document.getElementById('fbEnter').addEventListener('click', () => { if (this.editor.isEditing) this.editor.commit('stay'); });
    document.getElementById('nameBox').addEventListener('keydown', (e) => { e.stopPropagation(); if (e.key === 'Enter') { e.preventDefault(); this.onNameBoxEnter(); } if (e.key === 'Escape') { this.updateFormulaBar(); this.focusGrid(); } });
    document.getElementById('nameBox').addEventListener('focus', (e) => e.target.select());
    // drag & drop files
    const drop = document.getElementById('dropOverlay');
    document.addEventListener('dragover', (e) => { if (e.dataTransfer && [...e.dataTransfer.types].includes('Files')) { e.preventDefault(); drop.classList.add('show'); } });
    document.addEventListener('dragleave', (e) => { if (e.relatedTarget == null) drop.classList.remove('show'); });
    document.addEventListener('drop', (e) => { drop.classList.remove('show'); if (e.dataTransfer && e.dataTransfer.files.length) { e.preventDefault(); this.openFile(e.dataTransfer.files[0]); } });
    window.addEventListener('beforeunload', (e) => { this.autosave(); if (this.wb.dirty && !this.settings.autosave) { e.preventDefault(); e.returnValue = ''; } });
    const restored = this.restore();
    if (!restored) this.newWorkbook(this.wb);
    this.focusGrid();
    this.setStatus('Ready');
  },
};

window.officeTable = app;
app.init();
export default app;
