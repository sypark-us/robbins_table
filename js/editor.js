// In-cell editor + formula bar: typing, formula reference picking, function autocomplete, validation.
import { parse, extractRefs, isFormula, tokenize } from './formula.js';
import { FUNCTIONS, FUNCTION_LIST, num, bool } from './functions.js';
import { formatValue, isDateFormat, parseInput, formatGeneral } from './format.js';
import { rangeName, quoteSheetName, escapeHtml, MAX_ROWS, MAX_COLS } from './util.js';
import { isRange, isArr } from './range.js';
import { isErr } from './formula.js';

const REF_CHARS = new Set(['=', '(', ',', ';', '+', '-', '*', '/', '^', '&', '<', '>', ':', '{', ' ']);

export class CellEditor {
  constructor(app) {
    this.app = app;
    this.ta = document.getElementById('cellEditor');
    this.fb = document.getElementById('formulaInput');
    this.ac = document.getElementById('autocomplete');
    this.tip = document.getElementById('fnTip');
    this.isEditing = false; this.mode = 'edit'; this.cell = null; this.origText = ''; this.source = 'cell';
    this.refSpan = null; this.refSheetId = null; this.acItems = []; this.acIndex = 0;
    for (const el of [this.ta, this.fb]) {
      el.addEventListener('keydown', (e) => this.onKey(e, el));
      el.addEventListener('input', () => this.onInput(el));
      el.addEventListener('click', () => { this.refSpan = null; this.refresh(el); });
      el.addEventListener('keyup', (e) => { if (['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(e.key)) { this.refSpan = null; this.refresh(el); } });
    }
    this.fb.addEventListener('focus', () => { if (!this.isEditing) this.start('fb'); else this.source = 'fb'; });
    this.ta.addEventListener('focus', () => { this.source = 'cell'; });
    this.ta.addEventListener('blur', () => { setTimeout(() => { if (this.isEditing && document.activeElement !== this.fb && document.activeElement !== this.ta && !this.ac.contains(document.activeElement) && !this.app.dialogOpen() && !this.app.pickingRef) { /* keep editing while clicking around grid */ } }, 0); });
    this.ac.addEventListener('mousedown', (e) => { e.preventDefault(); const it = e.target.closest('.ac'); if (it) { this.acIndex = +it.dataset.i; this.acceptAutocomplete(); } });
  }
  get sheet() { return this.app.sheet; }
  isFormula() { return this.isEditing && this.text().startsWith('='); }
  text() { return this.source === 'fb' ? this.fb.value : this.ta.value; }
  focus() { (this.source === 'fb' ? this.fb : this.ta).focus(); }
  activeEl() { return this.source === 'fb' ? this.fb : this.ta; }

  // Text shown when editing a cell (raw formula or the unformatted value)
  static editText(app, sheet, r, c) {
    const cell = sheet.getCell(r, c);
    if (!cell) return '';
    if (cell.f) return cell.f;
    if (cell.spillOf) return '';
    const v = cell.v;
    if (v == null) return '';
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (typeof v === 'number') {
      const code = sheet.effectiveNumFmt(r, c);
      if (isDateFormat(code)) { const hasTime = v % 1 !== 0; return formatValue(v, hasTime ? (Math.abs(v) < 1 ? 'h:mm:ss' : 'm/d/yyyy h:mm:ss') : 'm/d/yyyy').text; }
      if (/%/.test(code) && !/"[^"]*%[^"]*"/.test(code)) return formatValue(v, '0.############%').text;
      return String(Number(v.toPrecision(15)));
    }
    const p = parseInput(String(v));
    if (typeof p.value !== 'string' || String(v).startsWith("'")) return "'" + v;
    return String(v);
  }

  start(mode = 'edit', initial = null) {
    if (this.isEditing) { if (initial != null) { this.ta.value += initial; this.fb.value = this.ta.value; } return; }
    const app = this.app; const s = this.sheet; const { r, c } = s.view.active;
    const m = s.mergeAt(r, c); const cr = m ? m.r1 : r, cc = m ? m.c1 : c;
    this.cell = { r: cr, c: cc, sheetId: s.id };
    this.origText = CellEditor.editText(app, s, cr, cc);
    const text = initial != null ? initial : this.origText;
    this.isEditing = true; this.mode = mode === 'fb' ? 'edit' : mode; this.source = mode === 'fb' ? 'fb' : 'cell';
    this.ta.value = text; this.fb.value = text;
    const st = s.getStyle(cr, cc);
    this.ta.style.font = app.grid.fontFor(st);
    this.ta.style.textAlign = st.align && st.align !== 'general' ? st.align : 'left';
    this.ta.classList.toggle('wrap', !!st.wrap);
    this.ta.style.display = 'block';
    this.updatePosition();
    if (mode !== 'fb') { this.ta.focus(); this.ta.setSelectionRange(text.length, text.length); }
    this.refSpan = null;
    app.setStatus(mode === 'type' ? 'Enter' : 'Edit');
    this.refresh(this.activeEl());
    app.grid.invalidate();
  }
  updatePosition() {
    if (!this.isEditing || this.source !== 'cell' && !this.ta.style.display) return;
    const app = this.app; const s = this.sheet;
    if (!this.cell || s.id !== this.cell.sheetId) { this.ta.style.display = 'none'; return; }
    const rect = app.grid.cellRect(this.cell.r, this.cell.c);
    if (!rect) { this.ta.style.display = 'none'; return; }
    this.ta.style.display = 'block';
    const z = app.grid.zoom;
    const text = this.ta.value || '';
    const font = this.ta.style.font;
    const lines = text.split('\n');
    let w = rect.w;
    if (!this.ta.classList.contains('wrap')) { let maxW = 0; for (const ln of lines) maxW = Math.max(maxW, app.grid.measure(ln, font)); w = Math.max(rect.w, maxW + 14 * z); }
    const maxW = app.grid.cssW - rect.x - 2;
    w = Math.min(w, Math.max(rect.w, maxW));
    const lineH = parseFloat(font) * 1.25;
    let h = Math.max(rect.h, lines.length * lineH + 6);
    if (this.ta.classList.contains('wrap')) { this.ta.style.height = 'auto'; h = Math.max(rect.h, this.ta.scrollHeight + 2); }
    this.ta.style.left = (rect.x - 1) + 'px'; this.ta.style.top = (rect.y - 1) + 'px';
    this.ta.style.width = (w + 2) + 'px'; this.ta.style.height = (h + 2) + 'px';
    this.ta.style.lineHeight = lineH + 'px';
    this.positionPopups();
  }
  positionPopups() {
    const rect = this.ta.getBoundingClientRect(); const ov = this.app.grid.overlay.getBoundingClientRect();
    const fbRect = this.fb.getBoundingClientRect();
    const useFb = this.source === 'fb' || this.ta.style.display === 'none';
    const left = useFb ? fbRect.left - ov.left : rect.left - ov.left;
    const top = useFb ? 4 : rect.bottom - ov.top + 2;
    for (const el of [this.ac, this.tip]) { el.style.left = Math.max(0, Math.min(left, ov.width - 300)) + 'px'; el.style.top = top + 'px'; }
    if (this.ac.style.display === 'block' && this.tip.style.display === 'block') this.tip.style.top = (top + this.ac.offsetHeight + 2) + 'px';
  }
  cancel() {
    if (!this.isEditing) return;
    this.isEditing = false; this.ta.style.display = 'none'; this.ac.style.display = 'none'; this.tip.style.display = 'none';
    this.app.grid.refHighlights = [];
    const s = this.sheet; const { r, c } = s.view.active;
    this.fb.value = CellEditor.editText(this.app, s, r, c);
    this.app.setStatus('Ready'); this.app.focusGrid(); this.app.grid.invalidate();
    if (this.cell && this.cell.sheetId !== s.id) this.app.switchToSheetId(this.cell.sheetId);
  }
  // Commit the edit. Returns false if the formula is invalid or validation rejects the value.
  commit(moveDir = null) {
    if (!this.isEditing) return true;
    const app = this.app;
    const text = this.text();
    const target = this.cell;
    const sheet = app.wb.sheetById(target.sheetId) || this.sheet;
    if (isFormula(text)) {
      try { parse(text); } catch (e) {
        app.dialogs.alert('Formula error', `There is a problem with this formula:\n${e.message}\n\nIf you did not intend to type a formula, start the text with an apostrophe (').`).then(() => this.focus());
        return false;
      }
    }
    if (text !== this.origText) {
      const dv = sheet.validationAt(target.r, target.c);
      if (dv && !isFormula(text)) {
        const p = parseInput(text);
        const ok = app.validate(sheet, target.r, target.c, dv, p.value);
        if (!ok) {
          if (dv.errorStyle === 'stop' || !dv.errorStyle) { app.dialogs.alert(dv.errorTitle || 'Invalid value', dv.error || 'The value you entered is not valid. A user has restricted values that can be entered into this cell.').then(() => this.focus()); return false; }
          if (dv.errorStyle === 'warning') { app.dialogs.confirm(dv.errorTitle || 'Warning', (dv.error || 'The value you entered is not valid.') + '\n\nContinue?').then((yes) => { if (yes) { this.finish(sheet, target, text, moveDir); } else this.focus(); }); return false; }
          app.toast(dv.error || 'Value does not match the validation rule');
        }
      }
    }
    this.finish(sheet, target, text, moveDir);
    return true;
  }
  finish(sheet, target, text, moveDir) {
    const app = this.app;
    this.isEditing = false; this.ta.style.display = 'none'; this.ac.style.display = 'none'; this.tip.style.display = 'none';
    app.grid.refHighlights = [];
    if (this.cell && this.cell.sheetId !== app.sheet.id) app.switchToSheetId(this.cell.sheetId);
    if (text !== this.origText) {
      app.transact('Edit cell', () => {
        const ranges = app.sheet.view.ranges;
        if (moveDir === 'ctrl' && ranges.length) for (const rg of ranges) for (let r = rg.r1; r <= Math.min(rg.r2, rg.r1 + 5000); r++) for (let c = rg.c1; c <= Math.min(rg.c2, rg.c1 + 500); c++) sheet.setInput(r, c, text);
        else sheet.setInput(target.r, target.c, text);
      });
    }
    app.setStatus('Ready');
    app.focusGrid();
    if (moveDir && moveDir !== 'ctrl' && moveDir !== 'stay') app.moveActive(moveDir, { enter: true });
    else { app.updateFormulaBar(); app.grid.invalidate(); }
  }

  // ---- keyboard inside the editor ----
  onKey(e, el) {
    if (!this.isEditing) { if (el === this.fb && (e.key === 'Enter' || e.key === 'Tab')) { e.preventDefault(); this.app.focusGrid(); } return; }
    const acOpen = this.ac.style.display === 'block';
    if (acOpen && (e.key === 'ArrowDown' || e.key === 'ArrowUp')) { e.preventDefault(); this.acIndex = (this.acIndex + (e.key === 'ArrowDown' ? 1 : this.acItems.length - 1)) % this.acItems.length; this.renderAutocomplete(); return; }
    if (acOpen && (e.key === 'Tab' || (e.key === 'Enter' && this.acItems.length))) { e.preventDefault(); this.acceptAutocomplete(); return; }
    if (e.key === 'Escape') { e.preventDefault(); if (acOpen) { this.ac.style.display = 'none'; return; } this.cancel(); return; }
    if (e.key === 'Enter') {
      if (e.altKey) { e.preventDefault(); this.insertText('\n'); return; }
      e.preventDefault(); this.commit(e.ctrlKey || e.metaKey ? 'ctrl' : e.shiftKey ? 'up' : 'down'); return;
    }
    if (e.key === 'Tab') { e.preventDefault(); this.commit(e.shiftKey ? 'left' : 'right'); return; }
    if (e.key === 'F2') { e.preventDefault(); this.mode = this.mode === 'type' ? 'edit' : 'type'; this.app.setStatus(this.mode === 'type' ? 'Enter' : 'Edit'); return; }
    if (e.key === 'F4') { e.preventDefault(); this.toggleAbsolute(); return; }
    if ((e.key.startsWith('Arrow') || e.key === 'Home' || e.key === 'End' || e.key === 'PageUp' || e.key === 'PageDown') && this.mode === 'type' && el === this.ta && !e.altKey) {
      // Enter mode: arrow keys commit and move (or extend a picked reference)
      if (this.isFormula() && this.canInsertRef() && !e.ctrlKey) { e.preventDefault(); this.keyboardRef(e); return; }
      if (!e.ctrlKey && !e.shiftKey) { e.preventDefault(); const dir = { ArrowDown: 'down', ArrowUp: 'up', ArrowLeft: 'left', ArrowRight: 'right', Home: 'home', End: 'end', PageUp: 'pageup', PageDown: 'pagedown' }[e.key]; this.commit(dir); return; }
    }
    if (this.refSpan && e.key.startsWith('Arrow') && this.isFormula() && !e.ctrlKey && !e.altKey) { e.preventDefault(); this.keyboardRef(e); return; }
    if (e.key.startsWith('Arrow')) this.refSpan = null;
    if (e.ctrlKey && (e.key === 'z' || e.key === 'y')) e.stopPropagation(); // native undo inside editor
  }
  keyboardRef(e) {
    // move/extend the picked reference with arrow keys
    const s = this.sheet;
    let rg = this.refSpan && this.refSpan.range ? { ...this.refSpan.range } : null;
    const base = rg ? (e.shiftKey ? { r: rg.r2, c: rg.c2 } : { r: rg.r1, c: rg.c1 }) : { ...s.view.active };
    const d = { ArrowDown: [1, 0], ArrowUp: [-1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[e.key];
    if (!d) return;
    const nr = Math.max(0, Math.min(MAX_ROWS - 1, base.r + d[0])), nc = Math.max(0, Math.min(MAX_COLS - 1, base.c + d[1]));
    if (e.shiftKey && rg) rg = { r1: Math.min(rg.r1, nr), c1: Math.min(rg.c1, nc), r2: Math.max(rg.r1, nr), c2: Math.max(rg.c1, nc) };
    else rg = { r1: nr, c1: nc, r2: nr, c2: nc };
    this.insertRef(rg, !!this.refSpan);
    this.app.grid.scrollTo(e.shiftKey ? rg.r2 : rg.r1, e.shiftKey ? rg.c2 : rg.c1);
  }
  onInput(el) {
    if (!this.isEditing) { if (el === this.fb) this.start('fb'); else return; }
    this.source = el === this.fb ? 'fb' : 'cell';
    if (el === this.fb) this.ta.value = this.fb.value; else this.fb.value = this.ta.value;
    this.refSpan = null;
    this.refresh(el);
    this.updatePosition();
  }
  refresh(el) {
    this.updateRefHighlights();
    this.updateAutocomplete(el);
    this.updateFnTip(el);
    this.positionPopups();
  }
  updateRefHighlights() {
    const app = this.app;
    if (!this.isFormula()) { if (app.grid.refHighlights.length) { app.grid.refHighlights = []; app.grid.invalidate(); } return; }
    const refs = extractRefs(this.text());
    const seen = new Map(); const out = [];
    for (const r of refs) { const k = `${r.sheet || ''}!${r.r1},${r.c1},${r.r2},${r.c2}`; if (seen.has(k)) continue; seen.set(k, out.length); out.push({ ...r, sheetId: r.sheet ? null : this.cell.sheetId }); }
    app.grid.refHighlights = out; app.grid.invalidate();
  }
  insertText(t) { const el = this.activeEl(); const a = el.selectionStart, b = el.selectionEnd; el.value = el.value.slice(0, a) + t + el.value.slice(b); el.setSelectionRange(a + t.length, a + t.length); this.onInput(el); }
  // Can a reference be inserted at the caret (previous significant char is an operator)?
  canInsertRef() {
    if (!this.isFormula()) return false;
    if (this.refSpan) return true;
    const el = this.activeEl(); const pos = el.selectionStart;
    const before = el.value.slice(0, pos).replace(/\s+$/, '');
    if (!before) return false;
    return REF_CHARS.has(before[before.length - 1]);
  }
  insertRef(rg, replaceLast) {
    const el = this.activeEl();
    const picked = this.app.sheet;
    const sheetPrefix = picked.id !== this.cell.sheetId ? quoteSheetName(picked.name) + '!' : '';
    const text = sheetPrefix + rangeName(rg);
    let a, b;
    if (replaceLast && this.refSpan) { a = this.refSpan.start; b = this.refSpan.end; }
    else { a = el.selectionStart; b = el.selectionEnd; }
    el.value = el.value.slice(0, a) + text + el.value.slice(b);
    el.setSelectionRange(a + text.length, a + text.length);
    this.refSpan = { start: a, end: a + text.length, range: rg };
    if (el === this.fb) this.ta.value = this.fb.value; else this.fb.value = this.ta.value;
    this.refresh(el); this.updatePosition();
    this.app.pickingRef = true;
    setTimeout(() => { this.app.pickingRef = false; el.focus(); }, 0);
  }
  toggleAbsolute() {
    const el = this.activeEl(); if (!this.isFormula()) return;
    const pos = el.selectionStart; const src = el.value.slice(1);
    const toks = tokenize(src).filter((t) => t.t === 'ref');
    const t = toks.find((tk) => pos - 1 >= tk.pos && pos - 1 <= tk.pos + tk.s.length);
    if (!t) return;
    const m = /^(.*!)?(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?::(\$?)([A-Za-z]{1,3})(\$?)(\d+))?$/.exec(t.s);
    if (!m) return;
    const cycle = (ac, ar) => ac && ar ? ['', '$'] : !ac && ar ? ['$', ''] : ac && !ar ? ['', ''] : ['$', '$'];
    const [c1, r1] = cycle(m[2] === '$', m[4] === '$');
    let out = (m[1] || '') + c1 + m[3] + r1 + m[5];
    if (m[7]) { const [c2, r2] = cycle(m[6] === '$', m[8] === '$'); out += ':' + c2 + m[7] + r2 + m[9]; }
    el.value = '=' + src.slice(0, t.pos) + out + src.slice(t.pos + t.s.length);
    const np = Math.min(el.value.length, t.pos + 1 + out.length);
    el.setSelectionRange(np, np);
    this.onInput(el);
  }
  // ---- autocomplete ----
  currentWord(el) {
    if (!this.isFormula()) return null;
    const pos = el.selectionStart; const before = el.value.slice(0, pos);
    const m = /([A-Za-z_][A-Za-z0-9_.]*)$/.exec(before);
    if (!m) return null;
    const start = pos - m[1].length;
    if (start > 0 && (before[start - 1] === '!' || before[start - 1] === '"')) return null;
    if (/^[A-Za-z]{1,3}\d+$/.test(m[1])) return null;
    return { word: m[1], start, end: pos };
  }
  updateAutocomplete(el) {
    const cw = this.currentWord(el);
    if (!cw || cw.word.length < 1) { this.ac.style.display = 'none'; this.acItems = []; return; }
    const up = cw.word.toUpperCase();
    const names = [...new Set(FUNCTION_LIST.map((f) => f.name))];
    const items = names.filter((n) => n.startsWith(up)).slice(0, 12);
    if (!items.length || (items.length === 1 && items[0] === up && el.value[cw.end] === '(')) { this.ac.style.display = 'none'; this.acItems = []; return; }
    this.acItems = items; this.acIndex = 0; this.acWord = cw;
    this.renderAutocomplete();
  }
  renderAutocomplete() {
    const html = this.acItems.map((n, i) => { const f = FUNCTIONS.get(n); return `<div class="ac${i === this.acIndex ? ' sel' : ''}" data-i="${i}"><span><b>${escapeHtml(n)}</b><span style="color:var(--fg-dim)">(${escapeHtml(f.sig)})</span></span><span class="cat">${escapeHtml(f.cat)}</span></div>`; }).join('');
    this.ac.innerHTML = html; this.ac.style.display = 'block';
    const sel = this.ac.querySelector('.sel'); if (sel) sel.scrollIntoView({ block: 'nearest' });
    this.positionPopups();
  }
  acceptAutocomplete() {
    const el = this.activeEl(); const cw = this.currentWord(el) || this.acWord; const name = this.acItems[this.acIndex];
    if (!cw || !name) return;
    const after = el.value.slice(cw.end);
    const ins = name + (after.startsWith('(') ? '' : '(');
    el.value = el.value.slice(0, cw.start) + ins + after;
    const p = cw.start + ins.length; el.setSelectionRange(p, p);
    this.ac.style.display = 'none'; this.acItems = [];
    this.onInput(el);
  }
  // ---- function signature tip ----
  updateFnTip(el) {
    if (!this.isFormula()) { this.tip.style.display = 'none'; return; }
    const pos = el.selectionStart; const src = el.value;
    // find innermost unclosed function call before caret
    let depth = 0, argIdx = 0, fnName = null, q = false;
    const stack = [];
    for (let i = 1; i < pos; i++) {
      const ch = src[i];
      if (ch === '"') { q = !q; continue; } if (q) continue;
      if (ch === '(') { const m = /([A-Za-z_][A-Za-z0-9_.]*)\s*$/.exec(src.slice(1, i)); stack.push({ name: m ? m[1].toUpperCase() : null, arg: 0 }); }
      else if (ch === ')') stack.pop();
      else if ((ch === ',' || ch === ';') && stack.length) stack[stack.length - 1].arg++;
    }
    const top = stack.length ? stack[stack.length - 1] : null;
    if (!top || !top.name || !FUNCTIONS.has(top.name)) { this.tip.style.display = 'none'; return; }
    const f = FUNCTIONS.get(top.name);
    const args = f.sig.split(',').map((s) => s.trim()).filter(Boolean);
    const cur = Math.min(top.arg, args.length - 1);
    const html = `${escapeHtml(top.name)}(${args.map((a, i) => (i === cur || (i === args.length - 1 && a.includes('...') && top.arg >= i)) ? `<b>${escapeHtml(a)}</b>` : escapeHtml(a)).join(', ')})`;
    this.tip.innerHTML = html; this.tip.style.display = 'block';
  }
}
