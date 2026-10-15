// Modal dialogs and popovers.
import { FORMAT_PRESETS, formatValue, parseInput, isDateFormat } from './format.js';
import { FUNCTIONS, FUNCTION_LIST } from './functions.js';
import { CHART_TYPES, newChart } from './charts.js';
import { colToName, rangeName, parseRangeName, escapeHtml, uid, normRange, MAX_ROWS, MAX_COLS, cellName, compareValues } from './util.js';
import { parse } from './formula.js';

export const THEME_COLORS = ['#000000', '#ffffff', '#e7e6e6', '#44546a', '#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47'];
const SHADES = [0.5, 0.35, 0.25, -0.25, -0.5];
export const STANDARD_COLORS = ['#c00000', '#ff0000', '#ffc000', '#ffff00', '#92d050', '#00b050', '#00b0f0', '#0070c0', '#002060', '#7030a0'];
function shade(hex, t) { const n = parseInt(hex.slice(1), 16); let r = (n >> 16) & 255, g = (n >> 8) & 255, b = n & 255; const f = (c) => t > 0 ? c + (255 - c) * t : c * (1 + t); return '#' + [r, g, b].map((c) => Math.round(f(c)).toString(16).padStart(2, '0')).join(''); }

export function h(tag, attrs = {}, ...children) {
  const el = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (v == null || v === false) continue;
    if (k === 'class') el.className = v; else if (k === 'style') el.style.cssText = v; else if (k.startsWith('on')) el.addEventListener(k.slice(2).toLowerCase(), v); else if (k === 'html') el.innerHTML = v; else if (k === 'value') el.value = v; else if (k === 'checked' || k === 'disabled' || k === 'selected' || k === 'multiple') el[k] = !!v; else el.setAttribute(k, v);
  }
  for (const c of children.flat()) { if (c == null || c === false) continue; el.appendChild(typeof c === 'string' ? document.createTextNode(c) : c); }
  return el;
}
const opt = (v, label, selected) => h('option', { value: v, selected }, label ?? String(v));
function sel(values, current, attrs = {}) { return h('select', attrs, values.map((v) => Array.isArray(v) ? opt(v[0], v[1], v[0] === current) : opt(v, null, v === current))); }

export function createDialogs(app) {
  const back = document.getElementById('modalBack');
  let stack = [];
  function open({ title, body, buttons = [], width, onClose, focus }) {
    const dlg = h('div', { class: 'dialog', style: width ? `width:${width}px` : '' });
    const head = h('div', { class: 'dhead' }, h('span', {}, title), h('button', { class: 'close', title: 'Close', onClick: () => close(null) }, '✕'));
    const bodyEl = h('div', { class: 'dbody' }); if (typeof body === 'string') bodyEl.innerHTML = body; else bodyEl.appendChild(body);
    const foot = h('div', { class: 'dfoot' });
    let resolve; const promise = new Promise((r) => { resolve = r; });
    const entry = { dlg, resolve, onClose };
    function close(result) { if (!stack.includes(entry)) return; stack = stack.filter((e) => e !== entry); dlg.remove(); if (!stack.length) back.classList.remove('open'); else stack[stack.length - 1].dlg.style.display = ''; if (onClose) onClose(result); resolve(result); setTimeout(() => app.focusGrid(), 0); }
    for (const b of buttons) foot.appendChild(h('button', { class: (b.primary ? 'primary' : '') + (b.left ? ' left' : ''), onClick: async () => { const r = b.onClick ? await b.onClick() : b.value; if (r !== false) close(r === undefined ? b.value : r); } }, b.label));
    dlg.appendChild(head); dlg.appendChild(bodyEl); if (buttons.length) dlg.appendChild(foot);
    dlg.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); close(null); }
      else if (e.key === 'Enter' && !(e.target.tagName === 'TEXTAREA') && !(e.target.tagName === 'SELECT' && e.target.multiple) && !e.target.classList.contains('noenter')) { const p = buttons.find((b) => b.primary); if (p) { e.preventDefault(); foot.children[buttons.indexOf(p)].click(); } }
      e.stopPropagation();
    });
    for (const e of stack) e.dlg.style.display = 'none';
    stack.push(entry); back.appendChild(dlg); back.classList.add('open');
    setTimeout(() => { const f = focus || dlg.querySelector('input:not([type=checkbox]):not([type=radio]), select, textarea, button.primary'); if (f) { f.focus(); if (f.select && f.type === 'text') f.select(); } }, 0);
    promise.close = close; promise.el = dlg; return promise;
  }
  back.addEventListener('mousedown', (e) => { if (e.target === back && stack.length) { /* click outside: keep open */ } });

  const D = {
    isOpen: () => stack.length > 0,
    alert: (title, msg) => open({ title, body: h('div', { style: 'white-space:pre-wrap;max-width:460px' }, msg), buttons: [{ label: 'OK', primary: true, value: true }] }),
    confirm: (title, msg) => open({ title, body: h('div', { style: 'white-space:pre-wrap;max-width:460px' }, msg), buttons: [{ label: 'Cancel', value: false }, { label: 'OK', primary: true, value: true }] }).then((v) => !!v),
    prompt: (title, label, value = '', hint = '') => { const inp = h('input', { type: 'text', value, style: 'width:100%' }); return open({ title, body: h('div', { class: 'form' }, h('label', {}, label), inp, hint ? h('div', { class: 'full', style: 'color:var(--fg-dim);font-size:11px' }, hint) : null), buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, onClick: () => inp.value }], focus: inp }); },
    toastError: (e) => app.toast(e && e.message ? e.message : String(e)),

    // ---------------- Format cells ----------------
    formatCells(tab = 'number') {
      const s = app.sheet; const { r, c } = s.view.active; const st = { ...s.getStyle(r, c) };
      const value = s.getValue(r, c); const sample = typeof value === 'number' ? value : 1234.5678;
      const patch = {};
      const tabs = ['Number', 'Alignment', 'Font', 'Border', 'Fill'];
      let active = tab;
      const tabBar = h('div', { class: 'tabs' }); const panes = {};
      // --- Number
      const categories = ['General', 'Number', 'Currency', 'Accounting', 'Percentage', 'Scientific', 'Fraction', 'Date', 'Time', 'Text', 'Custom'];
      const curCode = s.effectiveNumFmt(r, c);
      const guessCat = (code) => code === 'General' ? 'General' : code === '@' ? 'Text' : /_\(/.test(code) ? 'Accounting' : /[$€£¥]/.test(code) ? 'Currency' : /%/.test(code) ? 'Percentage' : /E\+/.test(code) ? 'Scientific' : /\?\//.test(code) ? 'Fraction' : /[hs]|AM\/PM/i.test(code) && !/y|d/.test(code.replace(/\[.*?\]/g, '')) ? 'Time' : /y|d|m{3,}/.test(code) && isDateFormat(code) ? 'Date' : isDateFormat(code) ? 'Date' : /^[#0.,]+$/.test(code) || /^[#0.,]+;/.test(code) ? 'Number' : 'Custom';
      let numCode = curCode;
      const catList = h('div', { class: 'listbox', style: 'width:140px;max-height:300px' });
      const opts = h('div', { style: 'flex:1;min-width:300px' });
      const preview = h('div', { class: 'preview' });
      const codeInput = h('input', { type: 'text', style: 'width:100%;font-family:Consolas,monospace', onInput: () => { numCode = codeInput.value; updatePreview(); } });
      const updatePreview = () => { try { const f = formatValue(typeof value === 'string' ? value : sample, numCode); preview.textContent = f.text || ' '; preview.style.color = f.color || '#000'; } catch { preview.textContent = 'Invalid format'; } codeInput.value = numCode; patch.numFmt = numCode === 'General' ? null : numCode; };
      const decimals = h('input', { type: 'number', min: 0, max: 15, value: 2, style: 'width:60px' });
      const thousands = h('input', { type: 'checkbox', checked: true });
      const negStyle = sel([['minus', '-1,234.10'], ['red', '1,234.10 (red)'], ['paren', '(1,234.10)'], ['redparen', '(1,234.10) (red)']], 'minus');
      const symbol = sel(['$', '€', '£', '¥', '₩', 'none'], '$');
      const dateList = h('div', { class: 'listbox', style: 'max-height:200px' });
      const buildNumber = (cat) => {
        const d = Math.max(0, Math.min(15, +decimals.value || 0)); const dec = d ? '.' + '0'.repeat(d) : '';
        const base = (thousands.checked ? '#,##0' : '0') + dec;
        const neg = negStyle.value;
        const wrapNeg = (pos) => neg === 'minus' ? pos : neg === 'red' ? `${pos};[Red]${pos}` : neg === 'paren' ? `${pos};(${pos})` : `${pos};[Red](${pos})`;
        if (cat === 'Number') return wrapNeg(base) === base ? base : `${base};${neg === 'red' ? '[Red]-' + base : neg === 'paren' ? '(' + base + ')' : '[Red](' + base + ')'}`;
        const sym = symbol.value === 'none' ? '' : symbol.value;
        if (cat === 'Currency') { const pos = sym + '#,##0' + dec; return neg === 'minus' ? pos : neg === 'red' ? `${pos};[Red]-${pos}` : neg === 'paren' ? `${pos};(${pos})` : `${pos};[Red](${pos})`; }
        if (cat === 'Accounting') return `_(${sym}* #,##0${dec}_);_(${sym}* (#,##0${dec});_(${sym}* "-"${d ? '?'.repeat(d) : ''}_);_(@_)`;
        if (cat === 'Percentage') return '0' + dec + '%';
        if (cat === 'Scientific') return '0' + dec + 'E+00';
        return base;
      };
      const renderOpts = (cat) => {
        opts.innerHTML = '';
        const row = (label, el) => h('div', { class: 'row', style: 'display:flex;gap:8px;align-items:center;margin:6px 0' }, h('label', { style: 'min-width:110px;color:var(--fg-dim)' }, label), el);
        if (cat === 'General') { numCode = 'General'; opts.appendChild(h('p', { style: 'color:var(--fg-dim)' }, 'General format cells have no specific number format.')); }
        else if (cat === 'Text') { numCode = '@'; opts.appendChild(h('p', { style: 'color:var(--fg-dim)' }, 'Text cells are displayed exactly as typed, even when the cell contains a number.')); }
        else if (['Number', 'Currency', 'Accounting', 'Percentage', 'Scientific'].includes(cat)) {
          opts.appendChild(row('Decimal places', decimals));
          if (cat === 'Number') { opts.appendChild(row('Use 1000 separator', thousands)); opts.appendChild(row('Negative numbers', negStyle)); }
          if (cat === 'Currency') { opts.appendChild(row('Symbol', symbol)); opts.appendChild(row('Negative numbers', negStyle)); }
          if (cat === 'Accounting') opts.appendChild(row('Symbol', symbol));
          const upd = () => { numCode = buildNumber(cat); updatePreview(); };
          for (const el of [decimals, thousands, negStyle, symbol]) { el.oninput = upd; el.onchange = upd; }
          upd();
        } else if (cat === 'Fraction') {
          const fl = sel([['# ?/?', 'Up to one digit (1/4)'], ['# ??/??', 'Up to two digits (21/25)'], ['# ???/???', 'Up to three digits (312/943)'], ['# ?/2', 'As halves (1/2)'], ['# ?/4', 'As quarters (2/4)'], ['# ?/8', 'As eighths (4/8)'], ['# ??/16', 'As sixteenths (8/16)'], ['# ?/10', 'As tenths (3/10)'], ['# ??/100', 'As hundredths (30/100)']], numCode);
          fl.onchange = () => { numCode = fl.value; updatePreview(); }; opts.appendChild(row('Type', fl)); numCode = fl.value;
        } else if (cat === 'Date' || cat === 'Time') {
          const codes = cat === 'Date' ? ['m/d/yyyy', 'yyyy-mm-dd', 'd-mmm-yy', 'd-mmm', 'mmm-yy', 'mmmm d, yyyy', 'dddd, mmmm d, yyyy', 'm/d/yy', 'mm/dd/yyyy', 'dd.mm.yyyy', 'yyyy-mm-dd h:mm', 'm/d/yyyy h:mm', 'mmm d'] : ['h:mm', 'h:mm:ss', 'h:mm AM/PM', 'h:mm:ss AM/PM', 'hh:mm:ss', '[h]:mm:ss', 'mm:ss', 'mm:ss.0', 'h:mm:ss.000'];
          dateList.innerHTML = '';
          if (!codes.includes(numCode)) numCode = codes[0];
          for (const code of codes) { const li = h('div', { class: 'li' + (code === numCode ? ' sel' : ''), onClick: () => { numCode = code; [...dateList.children].forEach((x) => x.classList.toggle('sel', x.dataset.code === code)); updatePreview(); } }, h('span', {}, formatValue(sample > 59 ? sample : 45000.75, code).text), h('small', {}, code)); li.dataset.code = code; dateList.appendChild(li); }
          opts.appendChild(h('div', {}, h('label', { style: 'color:var(--fg-dim)' }, 'Type'), dateList));
        } else if (cat === 'Custom') { opts.appendChild(h('div', {}, h('label', { style: 'color:var(--fg-dim)' }, 'Type a format code (Excel syntax):'), codeInput, h('div', { style: 'color:var(--fg-dim);font-size:11px;margin-top:6px' }, 'Examples: 0.00  #,##0  0%  $#,##0.00;[Red]-$#,##0.00  yyyy-mm-dd  [h]:mm  0.0,"K"  "Qty: "0'))); }
        updatePreview();
      };
      let curCat = guessCat(curCode);
      if (curCat === 'Number' || curCat === 'Currency' || curCat === 'Percentage' || curCat === 'Scientific') { const m = /\.(0+)/.exec(curCode); decimals.value = m ? m[1].length : 0; thousands.checked = curCode.includes('#,##'); if (/\[Red\]\(/.test(curCode)) negStyle.value = 'redparen'; else if (/\[Red\]/.test(curCode)) negStyle.value = 'red'; else if (/;\(/.test(curCode)) negStyle.value = 'paren'; const sm = /[$€£¥₩]/.exec(curCode); if (sm) symbol.value = sm[0]; }
      for (const cat of categories) { const li = h('div', { class: 'li' + (cat === curCat ? ' sel' : ''), onClick: () => { curCat = cat; [...catList.children].forEach((x) => x.classList.toggle('sel', x.textContent === cat)); if (cat !== guessCat(curCode)) numCode = curCode; renderOpts(cat); } }, cat); catList.appendChild(li); }
      panes.Number = h('div', { class: 'tabpane' }, h('div', { style: 'display:flex;gap:12px' }, h('div', {}, h('label', { style: 'color:var(--fg-dim)' }, 'Category'), catList), opts), h('label', { style: 'color:var(--fg-dim)' }, 'Sample'), preview);
      renderOpts(curCat);
      // --- Alignment
      const hAlign = sel([['general', 'General'], ['left', 'Left'], ['center', 'Center'], ['right', 'Right']], st.align || 'general');
      const vAlign = sel([['bottom', 'Bottom'], ['middle', 'Middle'], ['top', 'Top']], st.valign || 'bottom');
      const wrap = h('input', { type: 'checkbox', checked: !!st.wrap }); const indent = h('input', { type: 'number', min: 0, max: 15, value: st.indent || 0 });
      const mergeCb = h('input', { type: 'checkbox', checked: !!s.mergeAt(r, c) });
      panes.Alignment = h('div', { class: 'tabpane' }, h('div', { class: 'form' }, h('label', {}, 'Horizontal'), hAlign, h('label', {}, 'Vertical'), vAlign, h('label', {}, 'Indent'), indent, h('label', {}, 'Wrap text'), wrap, h('label', {}, 'Merge cells'), mergeCb));
      // --- Font
      const fontName = sel(['Calibri', 'Arial', 'Aptos', 'Cambria', 'Consolas', 'Courier New', 'Georgia', 'Segoe UI', 'Tahoma', 'Times New Roman', 'Trebuchet MS', 'Verdana', st.font].filter((v, i, a) => v && a.indexOf(v) === i), st.font || 'Calibri');
      const fontSize = h('input', { type: 'number', min: 1, max: 409, value: st.size || 11 });
      const bold = h('input', { type: 'checkbox', checked: !!st.bold }), italic = h('input', { type: 'checkbox', checked: !!st.italic }), underline = h('input', { type: 'checkbox', checked: !!st.underline }), strike = h('input', { type: 'checkbox', checked: !!st.strike });
      const fontColor = colorField(st.color || null, 'Automatic');
      const fontPreview = h('div', { class: 'preview' }, 'AaBbCcYyZz 123');
      const updFont = () => { fontPreview.style.fontFamily = fontName.value; fontPreview.style.fontSize = (fontSize.value || 11) + 'pt'; fontPreview.style.fontWeight = bold.checked ? 'bold' : 'normal'; fontPreview.style.fontStyle = italic.checked ? 'italic' : 'normal'; fontPreview.style.textDecoration = (underline.checked ? 'underline ' : '') + (strike.checked ? 'line-through' : ''); fontPreview.style.color = fontColor.get() || '#000'; };
      for (const el of [fontName, fontSize, bold, italic, underline, strike]) el.onchange = updFont; fontColor.onChange = updFont; updFont();
      panes.Font = h('div', { class: 'tabpane' }, h('div', { class: 'form' }, h('label', {}, 'Font'), fontName, h('label', {}, 'Size (pt)'), fontSize, h('label', {}, 'Bold'), bold, h('label', {}, 'Italic'), italic, h('label', {}, 'Underline'), underline, h('label', {}, 'Strikethrough'), strike, h('label', {}, 'Color'), fontColor.el), fontPreview);
      // --- Border
      const bStyle = sel([['thin', 'Thin'], ['medium', 'Medium'], ['thick', 'Thick'], ['dashed', 'Dashed'], ['dotted', 'Dotted'], ['double', 'Double']], 'thin');
      const bColor = colorField('#000000', 'Black');
      let borderChoice = null;
      const bp = h('div', { class: 'borderpicker' });
      for (const [id, label] of [['none', 'None'], ['all', 'All'], ['outline', 'Outline'], ['inside', 'Inside'], ['top', 'Top'], ['bottom', 'Bottom'], ['left', 'Left'], ['right', 'Right'], ['thickBottom', 'Thick bottom'], ['doubleBottom', 'Double bottom'], ['topBottom', 'Top & bottom'], ['leftRight', 'Left & right']]) {
        bp.appendChild(h('button', { onClick: (e) => { borderChoice = id; [...bp.children].forEach((x) => x.classList.toggle('primary', x === e.currentTarget)); } }, label));
      }
      panes.Border = h('div', { class: 'tabpane' }, h('div', { class: 'form' }, h('label', {}, 'Line style'), bStyle, h('label', {}, 'Color'), bColor.el), h('h4', {}, 'Presets'), bp, h('p', { style: 'color:var(--fg-dim);font-size:11px' }, 'Pick a preset to change borders; leave unselected to keep existing borders.'));
      // --- Fill
      const fillColor = colorField(st.fill || null, 'No fill');
      panes.Fill = h('div', { class: 'tabpane' }, h('div', { class: 'form' }, h('label', {}, 'Background color'), fillColor.el));
      const body = h('div', { style: 'min-width:560px' }, tabBar, ...tabs.map((t) => panes[t]));
      const show = (t) => { active = t.toLowerCase(); [...tabBar.children].forEach((x) => x.classList.toggle('active', x.textContent === t)); for (const k of tabs) panes[k].classList.toggle('active', k === t); };
      for (const t of tabs) tabBar.appendChild(h('div', { class: 't', onClick: () => show(t) }, t));
      show(tabs.find((t) => t.toLowerCase() === tab) || 'Number');
      return open({ title: 'Format cells', body, buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, onClick: () => {
        const p = { numFmt: numCode === 'General' ? null : numCode, align: hAlign.value === 'general' ? null : hAlign.value, valign: vAlign.value === 'bottom' ? null : vAlign.value, wrap: wrap.checked || null, indent: +indent.value || null, font: fontName.value === 'Calibri' ? null : fontName.value, size: +fontSize.value === 11 ? null : +fontSize.value, bold: bold.checked || null, italic: italic.checked || null, underline: underline.checked || null, strike: strike.checked || null, color: fontColor.get(), fill: fillColor.get() };
        app.transact('Format cells', () => {
          for (const rg of s.view.ranges) { s.applyStyle(rg, p); if (borderChoice) app.applyBorderPreset(rg, borderChoice, bStyle.value, bColor.get() || '#000000'); }
          const rg0 = s.view.ranges[0]; const merged = !!s.mergeAt(rg0.r1, rg0.c1);
          if (mergeCb.checked && !merged) s.merge(rg0); else if (!mergeCb.checked && merged) s.unmerge(rg0);
        });
        return true;
      } }] });
    },

    // ---------------- Sort ----------------
    sort() {
      const s = app.sheet; const rg = app.sortRegion();
      const hasHeader = h('input', { type: 'checkbox', checked: app.looksLikeHeader(rg) });
      const levels = h('div');
      const colOptions = () => { const out = []; for (let c = rg.c1; c <= Math.min(rg.c2, rg.c1 + 200); c++) { const hv = s.getValue(rg.r1, c); out.push([c, `Column ${colToName(c)}${hasHeader.checked && hv != null ? ` (${String(hv).slice(0, 20)})` : ''}`]); } return out; };
      const addLevel = (col) => { const row = h('div', { class: 'row', style: 'display:flex;gap:6px;margin:4px 0;align-items:center' }, h('span', { style: 'width:60px;color:var(--fg-dim)' }, levels.children.length ? 'Then by' : 'Sort by'), sel(colOptions(), col ?? rg.c1), sel([['asc', 'A → Z (ascending)'], ['desc', 'Z → A (descending)']], 'asc'), h('button', { onClick: () => row.remove() }, '✕')); levels.appendChild(row); };
      addLevel(s.view.active.c);
      hasHeader.onchange = () => { for (const row of levels.children) { const cur = row.children[1].value; const ns = sel(colOptions(), +cur); row.replaceChild(ns, row.children[1]); } };
      const body = h('div', { style: 'min-width:460px' }, h('div', { class: 'row', style: 'display:flex;gap:8px;align-items:center;margin-bottom:8px' }, h('label', {}, 'Range: '), h('b', {}, rangeName(rg)), h('label', { style: 'margin-left:auto' }, hasHeader, ' My data has headers')), levels, h('button', { onClick: () => addLevel(), style: 'margin-top:6px' }, '+ Add level'));
      return open({ title: 'Sort', body, buttons: [{ label: 'Cancel', value: null }, { label: 'Sort', primary: true, onClick: () => { const keys = [...levels.children].map((row) => ({ col: +row.children[1].value, asc: row.children[2].value === 'asc' })); if (!keys.length) return false; app.transact('Sort', () => s.sortRange(rg, keys, hasHeader.checked)); app.selectRange(rg); return true; } }] });
    },

    // ---------------- Filter popover ----------------
    filterPopover(col, anchorRect) {
      const s = app.sheet; const f = s.filter; if (!f) return;
      const cond = (f.conds && f.conds[col]) || {};
      const uniq = new Map();
      const r2 = Math.min(f.r2, s.usedBounds().r2);
      for (let r = f.r1 + 1; r <= r2; r++) {
        // rows hidden by OTHER columns' filters are excluded from the list
        let hiddenByOther = false;
        if (s.rowAxis.filtered.has(r)) { const others = Object.entries(f.conds || {}).filter(([k]) => +k !== col); if (others.length) { const save = f.conds; f.conds = Object.fromEntries(others); const tmp = new Set(s.rowAxis.filtered); s.applyFilter(); hiddenByOther = s.rowAxis.filtered.has(r); f.conds = save; s.rowAxis.filtered = tmp; s.rowAxis.invalidate(); } }
        if (hiddenByOther) continue;
        const text = app.fmt(s, r, col).text; const key = text; if (!uniq.has(key)) uniq.set(key, { text: text === '' ? '(Blanks)' : text, value: s.getValue(r, col), count: 0 }); uniq.get(key).count++;
      }
      const entries = [...uniq.entries()].sort((a, b) => compareValues(a[1].value, b[1].value));
      const checked = new Set(cond.values ? cond.values : entries.map((e) => e[0]));
      const pop = h('div', { class: 'popover', style: 'width:260px' });
      const search = h('input', { type: 'search', placeholder: 'Search', style: 'width:100%' });
      const list = h('div', { class: 'list', style: 'max-height:180px' });
      const render = () => { list.innerHTML = ''; const q = search.value.toLowerCase(); for (const [k, e] of entries) { if (q && !e.text.toLowerCase().includes(q)) continue; const cb = h('input', { type: 'checkbox', checked: checked.has(k), onChange: () => { if (cb.checked) checked.add(k); else checked.delete(k); } }); list.appendChild(h('label', {}, cb, ' ', e.text, h('span', { style: 'color:var(--fg-dim)' }, ` (${e.count})`))); } };
      search.oninput = render; render();
      const textOp = sel([['', 'Text filter…'], ['contains', 'Contains'], ['notContains', 'Does not contain'], ['equals', 'Equals'], ['notEquals', 'Does not equal'], ['begins', 'Begins with'], ['ends', 'Ends with']], cond.text ? cond.text.op : '');
      const textVal = h('input', { type: 'text', value: cond.text ? cond.text.v : '', style: 'flex:1', placeholder: 'value' });
      const numOp = sel([['', 'Number filter…'], ['gt', 'Greater than'], ['ge', 'Greater or equal'], ['lt', 'Less than'], ['le', 'Less or equal'], ['eq', 'Equals'], ['ne', 'Not equal'], ['between', 'Between']], cond.num ? cond.num.op : '');
      const numV1 = h('input', { type: 'text', value: cond.num ? cond.num.v1 : '', style: 'width:70px', placeholder: 'value' }); const numV2 = h('input', { type: 'text', value: cond.num && cond.num.v2 != null ? cond.num.v2 : '', style: 'width:70px', placeholder: 'and' });
      const close = () => { pop.remove(); document.removeEventListener('mousedown', outside, true); };
      const outside = (e) => { if (!pop.contains(e.target)) close(); };
      const apply = () => {
        const nc = {};
        if (checked.size !== entries.length) nc.values = [...checked].map((k) => (k === '' ? '' : k));
        if (textOp.value && textVal.value !== '') nc.text = { op: textOp.value, v: textVal.value };
        if (numOp.value && numV1.value !== '') nc.num = { op: numOp.value, v1: parseFloat(numV1.value), v2: numV2.value !== '' ? parseFloat(numV2.value) : undefined };
        const conds = { ...(f.conds || {}) }; if (Object.keys(nc).length) conds[col] = nc; else delete conds[col];
        app.transact('Filter', () => s.setFilter({ ...f, conds })); close();
      };
      pop.append(
        h('div', { class: 'item', onClick: () => { app.transact('Sort', () => s.sortRange(f, [{ col, asc: true }], true)); close(); } }, 'Sort A → Z'),
        h('div', { class: 'item', onClick: () => { app.transact('Sort', () => s.sortRange(f, [{ col, asc: false }], true)); close(); } }, 'Sort Z → A'),
        h('div', { class: 'item', onClick: () => { const conds = { ...(f.conds || {}) }; delete conds[col]; app.transact('Clear filter', () => s.setFilter({ ...f, conds })); close(); } }, `Clear filter from column ${colToName(col)}`),
        h('hr'),
        h('div', { class: 'row' }, textOp, textVal), h('div', { class: 'row' }, numOp, numV1, numV2),
        search, list,
        h('div', { class: 'row' }, h('button', { onClick: () => { entries.forEach(([k]) => checked.add(k)); render(); } }, 'Select all'), h('button', { onClick: () => { checked.clear(); render(); } }, 'Clear')),
        h('div', { class: 'row right' }, h('button', { onClick: close }, 'Cancel'), h('button', { class: 'primary', onClick: apply }, 'OK')),
      );
      app.grid.overlay.appendChild(pop);
      pop.style.left = Math.max(0, Math.min(anchorRect.x + anchorRect.w - 260, app.grid.cssW - 270)) + 'px'; pop.style.top = (anchorRect.y + anchorRect.h + 2) + 'px';
      pop.addEventListener('keydown', (e) => { e.stopPropagation(); if (e.key === 'Escape') close(); if (e.key === 'Enter' && e.target !== search) apply(); });
      setTimeout(() => document.addEventListener('mousedown', outside, true), 0);
      search.focus();
    },

    // ---------------- Validation list popover ----------------
    validationList(r, c) {
      const s = app.sheet; const dv = s.validationAt(r, c); if (!dv) return;
      const values = app.validationListValues(s, dv);
      const rect = app.grid.cellRect(r, c); if (!rect) return;
      const pop = h('div', { class: 'popover dvlist' });
      const close = () => { pop.remove(); document.removeEventListener('mousedown', outside, true); };
      const outside = (e) => { if (!pop.contains(e.target)) close(); };
      for (const v of values) pop.appendChild(h('div', { class: 'item', onClick: () => { app.transact('Pick value', () => s.setInput(r, c, String(v))); close(); app.focusGrid(); } }, String(v)));
      if (!values.length) pop.appendChild(h('div', { class: 'item', style: 'color:var(--fg-dim)' }, '(no values)'));
      app.grid.overlay.appendChild(pop);
      pop.style.left = rect.x + 'px'; pop.style.top = (rect.y + rect.h) + 'px'; pop.style.minWidth = Math.max(120, rect.w + 18) + 'px';
      setTimeout(() => document.addEventListener('mousedown', outside, true), 0);
    },

    // ---------------- Find & replace ----------------
    findReplace(showReplace = false) {
      let panel = document.getElementById('findPanel');
      if (panel) panel.remove();
      const find = h('input', { type: 'text', style: 'flex:1', placeholder: 'Find what' });
      const repl = h('input', { type: 'text', style: 'flex:1', placeholder: 'Replace with' });
      const matchCase = h('input', { type: 'checkbox' }), entire = h('input', { type: 'checkbox' }), inFormulas = h('input', { type: 'checkbox', checked: true }), allSheets = h('input', { type: 'checkbox' });
      const results = h('div', { class: 'results hidden' });
      const status = h('span', { style: 'color:var(--fg-dim);font-size:11px;margin-right:auto' });
      const opts = () => ({ matchCase: matchCase.checked, entire: entire.checked, inFormulas: inFormulas.checked, allSheets: allSheets.checked });
      const doFind = () => { const n = app.findNext(find.value, opts()); status.textContent = n ? '' : 'Not found'; };
      const doFindAll = () => { const hits = app.findAll(find.value, opts()); results.innerHTML = ''; results.classList.toggle('hidden', !hits.length); status.textContent = `${hits.length} found`; for (const hit of hits.slice(0, 500)) results.appendChild(h('div', { onClick: () => app.gotoHit(hit) }, `${hit.sheet.name}!${cellName(hit.r, hit.c)}: ${hit.text.slice(0, 60)}`)); };
      panel = h('div', { id: 'findPanel', class: 'popover' },
        h('div', { class: 'row' }, find, h('button', { onClick: doFind }, 'Find next'), h('button', { onClick: doFindAll }, 'Find all')),
        h('div', { class: 'row' + (showReplace ? '' : ' hidden'), id: 'replaceRow' }, repl, h('button', { onClick: () => { const ok = app.replaceCurrent(find.value, repl.value, opts()); status.textContent = ok ? 'Replaced' : 'Not found'; } }, 'Replace'), h('button', { onClick: () => { const n = app.replaceAll(find.value, repl.value, opts()); status.textContent = `Replaced ${n}`; } }, 'Replace all')),
        h('div', { class: 'row', style: 'font-size:11px;flex-wrap:wrap' }, h('label', {}, matchCase, ' Match case'), h('label', {}, entire, ' Entire cell'), h('label', {}, inFormulas, ' Formulas'), h('label', {}, allSheets, ' All sheets')),
        results,
        h('div', { class: 'row right' }, status, h('button', { onClick: () => { document.getElementById('replaceRow').classList.toggle('hidden'); } }, 'Replace…'), h('button', { onClick: () => panel.remove() }, 'Close')));
      panel.addEventListener('keydown', (e) => { e.stopPropagation(); if (e.key === 'Escape') panel.remove(); if (e.key === 'Enter') { e.preventDefault(); if (e.target === repl) panel.querySelector('#replaceRow button').click(); else doFind(); } });
      app.grid.overlay.appendChild(panel);
      const selText = app.selectionText(); if (selText && selText.length < 60) find.value = selText;
      find.focus(); find.select();
    },

    // ---------------- Insert function ----------------
    insertFunction() {
      const search = h('input', { type: 'search', placeholder: 'Search functions', style: 'width:100%' });
      const cats = ['All', ...new Set(FUNCTION_LIST.map((f) => f.cat))];
      const catSel = sel(cats, 'All');
      const list = h('div', { class: 'listbox', style: 'height:260px' });
      const desc = h('div', { style: 'font-family:Consolas,monospace;font-size:12px;color:var(--fg-dim);min-height:36px;margin-top:6px;white-space:pre-wrap' });
      let selected = null;
      const names = [...new Set(FUNCTION_LIST.map((f) => f.name))].sort();
      const render = () => { list.innerHTML = ''; const q = search.value.toUpperCase(); for (const n of names) { const f = FUNCTIONS.get(n); if (catSel.value !== 'All' && f.cat !== catSel.value) continue; if (q && !n.includes(q) && !f.sig.toUpperCase().includes(q)) continue; const li = h('div', { class: 'li' + (selected === n ? ' sel' : ''), onClick: () => { selected = n; desc.textContent = `${n}(${f.sig})`; [...list.children].forEach((x) => x.classList.toggle('sel', x.dataset.n === n)); }, onDblClick: () => { selected = n; okBtn(); } }, h('span', {}, n), h('small', {}, f.cat)); li.dataset.n = n; list.appendChild(li); } };
      search.oninput = render; catSel.onchange = render; render();
      let okBtn = () => {};
      const p = open({ title: 'Insert function', body: h('div', { style: 'width:420px' }, h('div', { class: 'row', style: 'display:flex;gap:6px;margin-bottom:6px' }, search, catSel), list, desc), buttons: [{ label: 'Cancel', value: null }, { label: 'Insert', primary: true, onClick: () => { if (!selected) return false; return selected; } }], focus: search });
      okBtn = () => p.close(selected);
      return p.then((name) => { if (name) { const cur = app.editor.isEditing ? app.editor.text() : ''; if (app.editor.isEditing) app.editor.insertText(`${name}(`); else app.editor.start('edit', cur.startsWith('=') ? cur + `${name}(` : `=${name}(`); } });
    },

    // ---------------- Chart ----------------
    chart(existing = null) {
      const s = app.sheet;
      const rg = existing ? null : app.dataRegion();
      const type = sel(CHART_TYPES.map((t) => [t.id, t.label]), existing ? existing.type : 'column');
      const range = h('input', { type: 'text', value: existing ? existing.range : (rg ? `${s.name.includes(' ') ? `'${s.name}'` : s.name}!${rangeName(rg)}` : '') });
      const title = h('input', { type: 'text', value: existing && existing.title != null ? existing.title : '', placeholder: '(auto)' });
      const legend = h('input', { type: 'checkbox', checked: existing ? existing.legend !== false : true });
      const labels = h('input', { type: 'checkbox', checked: existing ? !!existing.dataLabels : false });
      const orient = sel([['auto', 'Auto'], ['cols', 'Columns'], ['rows', 'Rows']], existing ? (existing.seriesInRows == null ? 'auto' : existing.seriesInRows ? 'rows' : 'cols') : 'auto');
      const body = h('div', { class: 'form', style: 'width:420px' }, h('label', {}, 'Chart type'), type, h('label', {}, 'Data range'), range, h('label', {}, 'Title'), title, h('label', {}, 'Series in'), orient, h('label', {}, 'Legend'), legend, h('label', {}, 'Data labels'), labels);
      return open({ title: existing ? 'Edit chart' : 'Insert chart', body, buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, onClick: () => {
        const pr = parseRangeName(range.value); if (!pr) { app.toast('Enter a valid range like Sheet1!A1:C10'); return false; }
        const patch = { type: type.value, range: range.value.trim(), title: title.value.trim() || null, legend: legend.checked, dataLabels: labels.checked, seriesInRows: orient.value === 'auto' ? null : orient.value === 'rows' };
        if (existing) app.transact('Edit chart', () => s.updateChart(existing.id, patch));
        else { const anchor = { r: Math.max(s.view.active.r, app.grid.scrollRow), c: (pr.sheet == null || pr.sheet.toLowerCase() === s.name.toLowerCase()) ? pr.c2 + 1 : s.view.active.c }; const ch = { ...newChart(type.value, patch.range, anchor), ...patch }; app.transact('Insert chart', () => s.setCharts([...s.charts, ch])); app.charts.select(ch.id); }
        return true;
      } }] });
    },

    // ---------------- Conditional formatting ----------------
    conditionalFormatting() {
      const s = app.sheet;
      const list = h('div', { class: 'listbox', style: 'min-height:120px;max-height:300px' });
      let rules = s.cf.map((r) => ({ ...r }));
      let selIdx = -1;
      const describe = (r) => ({ cellIs: `Cell value ${({ gt: '>', lt: '<', ge: '≥', le: '≤', eq: '=', ne: '≠', between: 'between', notBetween: 'not between' })[r.op] || r.op} ${r.v1 ?? ''}${r.v2 != null && r.v2 !== '' ? ` and ${r.v2}` : ''}`, containsText: `Text contains "${r.v1}"`, notContainsText: `Text does not contain "${r.v1}"`, beginsWith: `Begins with "${r.v1}"`, endsWith: `Ends with "${r.v1}"`, duplicate: 'Duplicate values', unique: 'Unique values', top: `Top ${r.rank}${r.percent ? '%' : ''}`, bottom: `Bottom ${r.rank}${r.percent ? '%' : ''}`, aboveAverage: 'Above average', belowAverage: 'Below average', blank: 'Blank cells', notBlank: 'Non-blank cells', formula: `Formula: ${r.v1}`, colorScale: `Color scale (${(r.colors || []).length} colors)`, dataBar: 'Data bar' })[r.type] || r.type;
      const render = () => { list.innerHTML = ''; rules.forEach((r, i) => { const sample = h('div', { class: 'sample', style: `background:${r.type === 'colorScale' ? `linear-gradient(90deg,${(r.colors || ['#f8696b', '#63be7b']).join(',')})` : r.type === 'dataBar' ? `linear-gradient(90deg,${r.barColor || '#638ec6'} 60%,#fff 60%)` : (r.style && r.style.fill) || '#fff'};color:${(r.style && r.style.color) || '#000'};font-weight:${r.style && r.style.bold ? 'bold' : 'normal'};font-style:${r.style && r.style.italic ? 'italic' : 'normal'}` }, r.type === 'colorScale' || r.type === 'dataBar' ? '' : 'AaBb'); list.appendChild(h('div', { class: 'cf-rule' + (i === selIdx ? ' sel' : ''), style: i === selIdx ? 'background:var(--bg3)' : '', onClick: () => { selIdx = i; render(); }, onDblClick: () => edit(i) }, sample, h('div', { class: 'desc' }, describe(r)), h('div', { style: 'color:var(--fg-dim)' }, rangeName(r)))); }); };
      const edit = (idx) => this.cfRule(idx >= 0 ? rules[idx] : null).then((rule) => { if (!rule) return; if (idx >= 0) rules[idx] = rule; else rules.push(rule); selIdx = idx >= 0 ? idx : rules.length - 1; render(); });
      render();
      const body = h('div', { style: 'width:560px' }, h('div', { class: 'row', style: 'display:flex;gap:6px;margin-bottom:6px' }, h('button', { onClick: () => edit(-1) }, 'New rule…'), h('button', { onClick: () => { if (selIdx >= 0) edit(selIdx); } }, 'Edit rule…'), h('button', { onClick: () => { if (selIdx >= 0) { rules.splice(selIdx, 1); selIdx = -1; render(); } } }, 'Delete rule'), h('button', { onClick: () => { if (selIdx > 0) { [rules[selIdx - 1], rules[selIdx]] = [rules[selIdx], rules[selIdx - 1]]; selIdx--; render(); } } }, '▲'), h('button', { onClick: () => { if (selIdx >= 0 && selIdx < rules.length - 1) { [rules[selIdx + 1], rules[selIdx]] = [rules[selIdx], rules[selIdx + 1]]; selIdx++; render(); } } }, '▼')), list, h('p', { style: 'color:var(--fg-dim);font-size:11px' }, 'Rules are evaluated top to bottom; earlier rules win when several apply.'));
      return open({ title: 'Conditional formatting rules', body, buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, onClick: () => { app.transact('Conditional formatting', () => s.setCF(rules)); return true; } }] });
    },
    cfRule(rule) {
      const s = app.sheet; const rg = rule || s.view.ranges[0];
      const range = h('input', { type: 'text', value: rangeName(rg) });
      const type = sel([['cellIs', 'Cell value is…'], ['containsText', 'Text contains…'], ['notContainsText', 'Text does not contain…'], ['beginsWith', 'Text begins with…'], ['endsWith', 'Text ends with…'], ['duplicate', 'Duplicate values'], ['unique', 'Unique values'], ['top', 'Top N'], ['bottom', 'Bottom N'], ['aboveAverage', 'Above average'], ['belowAverage', 'Below average'], ['blank', 'Blank cells'], ['notBlank', 'Non-blank cells'], ['formula', 'Custom formula is true'], ['colorScale', 'Color scale'], ['dataBar', 'Data bar']], rule ? rule.type : 'cellIs');
      const op = sel([['gt', 'greater than'], ['ge', 'greater than or equal'], ['lt', 'less than'], ['le', 'less than or equal'], ['eq', 'equal to'], ['ne', 'not equal to'], ['between', 'between'], ['notBetween', 'not between']], rule && rule.op || 'gt');
      const v1 = h('input', { type: 'text', value: rule && rule.v1 != null ? rule.v1 : '', placeholder: 'value or =formula' });
      const v2 = h('input', { type: 'text', value: rule && rule.v2 != null ? rule.v2 : '', placeholder: 'and' });
      const rank = h('input', { type: 'number', value: rule && rule.rank || 10, min: 1 }); const percent = h('input', { type: 'checkbox', checked: !!(rule && rule.percent) });
      const st = (rule && rule.style) || { fill: '#ffc7ce', color: '#9c0006' };
      const bold = h('input', { type: 'checkbox', checked: !!st.bold }), italic = h('input', { type: 'checkbox', checked: !!st.italic });
      const color = colorField(st.color || null, 'Automatic'); const fill = colorField(st.fill || null, 'No fill');
      const presets = sel([['custom', 'Custom…'], ['lightRed', 'Light red fill with dark red text'], ['yellow', 'Yellow fill with dark yellow text'], ['green', 'Green fill with dark green text'], ['redText', 'Red text'], ['redBorderless', 'Light red fill'], ['bold', 'Bold']], 'custom');
      presets.onchange = () => { const map = { lightRed: { fill: '#ffc7ce', color: '#9c0006' }, yellow: { fill: '#ffeb9c', color: '#9c5700' }, green: { fill: '#c6efce', color: '#006100' }, redText: { color: '#ff0000' }, redBorderless: { fill: '#ffc7ce' }, bold: { bold: true } }; const p = map[presets.value]; if (!p) return; color.set(p.color || null); fill.set(p.fill || null); bold.checked = !!p.bold; italic.checked = false; };
      const c1 = colorField(rule && rule.colors ? rule.colors[0] : '#f8696b', null), c2 = colorField(rule && rule.colors && rule.colors.length === 3 ? rule.colors[1] : '#ffeb84', null), c3 = colorField(rule && rule.colors ? rule.colors[rule.colors.length - 1] : '#63be7b', null);
      const three = h('input', { type: 'checkbox', checked: !(rule && rule.colors && rule.colors.length === 2) });
      const barColor = colorField(rule && rule.barColor || '#638ec6', null);
      const rows = {
        op: h('div', { class: 'row', style: 'display:contents' }, h('label', {}, 'Operator'), op),
        v1: h('div', { style: 'display:contents' }, h('label', {}, 'Value'), v1), v2: h('div', { style: 'display:contents' }, h('label', {}, 'And'), v2),
        rank: h('div', { style: 'display:contents' }, h('label', {}, 'Rank'), h('div', {}, rank, ' ', h('label', {}, percent, ' percent'))),
        style: h('div', { style: 'display:contents' }, h('label', {}, 'Preset'), presets, h('label', {}, 'Bold / italic'), h('div', {}, bold, ' B ', italic, ' I'), h('label', {}, 'Text color'), color.el, h('label', {}, 'Fill'), fill.el),
        scale: h('div', { style: 'display:contents' }, h('label', {}, 'Minimum'), c1.el, h('label', {}, 'Midpoint'), h('div', {}, three, ' ', c2.el), h('label', {}, 'Maximum'), c3.el),
        bar: h('div', { style: 'display:contents' }, h('label', {}, 'Bar color'), barColor.el),
      };
      const form = h('div', { class: 'form', style: 'width:480px' }, h('label', {}, 'Apply to'), range, h('label', {}, 'Rule type'), type, ...Object.values(rows));
      const update = () => { const t = type.value; rows.op.style.display = t === 'cellIs' ? 'contents' : 'none'; rows.v1.style.display = ['cellIs', 'containsText', 'notContainsText', 'beginsWith', 'endsWith', 'formula'].includes(t) ? 'contents' : 'none'; rows.v2.style.display = t === 'cellIs' && (op.value === 'between' || op.value === 'notBetween') ? 'contents' : 'none'; rows.rank.style.display = t === 'top' || t === 'bottom' ? 'contents' : 'none'; rows.style.style.display = t === 'colorScale' || t === 'dataBar' ? 'none' : 'contents'; rows.scale.style.display = t === 'colorScale' ? 'contents' : 'none'; rows.bar.style.display = t === 'dataBar' ? 'contents' : 'none'; v1.placeholder = t === 'formula' ? '=A1>100 (relative to top-left cell)' : t === 'cellIs' ? 'value or =formula' : 'text'; };
      type.onchange = update; op.onchange = update; update();
      return open({ title: rule ? 'Edit rule' : 'New formatting rule', body: form, buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, onClick: () => {
        const pr = parseRangeName(range.value); if (!pr) { app.toast('Invalid range'); return false; }
        const t = type.value;
        const out = { id: rule ? rule.id : uid('cf'), r1: pr.r1, c1: pr.c1, r2: pr.r2, c2: pr.c2, type: t };
        if (t === 'cellIs') { out.op = op.value; out.v1 = v1.value; out.v2 = v2.value; }
        else if (['containsText', 'notContainsText', 'beginsWith', 'endsWith', 'formula'].includes(t)) out.v1 = v1.value;
        else if (t === 'top' || t === 'bottom') { out.rank = +rank.value || 10; out.percent = percent.checked; }
        if (t === 'colorScale') { out.colors = three.checked ? [c1.get() || '#f8696b', c2.get() || '#ffeb84', c3.get() || '#63be7b'] : [c1.get() || '#f8696b', c3.get() || '#63be7b']; out.style = null; }
        else if (t === 'dataBar') { out.barColor = barColor.get() || '#638ec6'; out.style = null; }
        else out.style = { bold: bold.checked || undefined, italic: italic.checked || undefined, color: color.get() || undefined, fill: fill.get() || undefined };
        if (t === 'formula' && out.v1 && !out.v1.startsWith('=')) out.v1 = '=' + out.v1;
        if (t === 'formula') { try { parse(out.v1); } catch (e) { app.toast('Invalid formula: ' + e.message); return false; } }
        return out;
      } }] });
    },

    // ---------------- Data validation ----------------
    dataValidation() {
      const s = app.sheet; const rg = s.view.ranges[0]; const existing = s.validationAt(s.view.active.r, s.view.active.c);
      const type = sel([['list', 'List'], ['whole', 'Whole number'], ['decimal', 'Decimal'], ['date', 'Date'], ['textLength', 'Text length'], ['custom', 'Custom formula']], existing ? existing.type : 'list');
      const op = sel([['between', 'between'], ['notBetween', 'not between'], ['eq', 'equal to'], ['ne', 'not equal to'], ['gt', 'greater than'], ['lt', 'less than'], ['ge', 'greater than or equal'], ['le', 'less than or equal']], existing && existing.op || 'between');
      const v1 = h('input', { type: 'text', value: existing ? existing.v1 ?? '' : '' }); const v2 = h('input', { type: 'text', value: existing ? existing.v2 ?? '' : '' });
      const allowBlank = h('input', { type: 'checkbox', checked: existing ? existing.allowBlank !== false : true }); const dropdown = h('input', { type: 'checkbox', checked: existing ? existing.showDropdown !== false : true });
      const errStyle = sel([['stop', 'Stop (reject)'], ['warning', 'Warning (ask)'], ['info', 'Information (allow)']], existing && existing.errorStyle || 'stop');
      const errTitle = h('input', { type: 'text', value: existing && existing.errorTitle || '' }); const errMsg = h('input', { type: 'text', value: existing && existing.error || '' }); const prompt = h('input', { type: 'text', value: existing && existing.prompt || '' });
      const range = h('input', { type: 'text', value: rangeName(existing || rg) });
      const l1 = h('label', {}, 'Value'), l2 = h('label', {}, 'And'), lop = h('label', {}, 'Data'), ldd = h('label', {}, 'In-cell dropdown');
      const update = () => { const t = type.value; const show = (el, on) => { el.style.display = on ? '' : 'none'; }; show(op, t !== 'list' && t !== 'custom'); show(lop, t !== 'list' && t !== 'custom'); show(v2, t !== 'list' && t !== 'custom' && (op.value === 'between' || op.value === 'notBetween')); show(l2, v2.style.display !== 'none'); show(dropdown, t === 'list'); show(ldd, t === 'list'); l1.textContent = t === 'list' ? 'Source' : t === 'custom' ? 'Formula' : op.value === 'between' || op.value === 'notBetween' ? 'Minimum' : 'Value'; v1.placeholder = t === 'list' ? 'Yes,No,Maybe  or  =Sheet1!$A$1:$A$5' : t === 'custom' ? '=ISNUMBER(A1)' : 'number or =formula'; };
      type.onchange = update; op.onchange = update; update();
      const body = h('div', { class: 'form', style: 'width:460px' }, h('label', {}, 'Apply to'), range, h('label', {}, 'Allow'), type, lop, op, l1, v1, l2, v2, h('label', {}, 'Ignore blank'), allowBlank, ldd, dropdown, h('label', {}, 'Input message'), prompt, h('label', {}, 'On invalid data'), errStyle, h('label', {}, 'Error title'), errTitle, h('label', {}, 'Error message'), errMsg);
      return open({ title: 'Data validation', body, buttons: [{ label: 'Clear', left: true, onClick: () => { app.transact('Clear validation', () => s.setValidations(s.validations.filter((v) => !existing || v.id !== existing.id))); return true; } }, { label: 'Cancel', value: null }, { label: 'OK', primary: true, onClick: () => {
        const pr = parseRangeName(range.value); if (!pr) { app.toast('Invalid range'); return false; }
        const v = { id: existing ? existing.id : uid('dv'), r1: pr.r1, c1: pr.c1, r2: pr.r2, c2: pr.c2, type: type.value, op: op.value, v1: v1.value, v2: v2.value, allowBlank: allowBlank.checked, showDropdown: dropdown.checked, errorStyle: errStyle.value, errorTitle: errTitle.value, error: errMsg.value, prompt: prompt.value };
        app.transact('Data validation', () => s.setValidations([...s.validations.filter((x) => !existing || x.id !== existing.id), v]));
        return true;
      } }] });
    },

    // ---------------- Name manager ----------------
    nameManager() {
      const wb = app.wb;
      const list = h('div', { class: 'listbox', style: 'min-height:100px;max-height:260px' });
      const render = () => { list.innerHTML = ''; for (const [k, n] of wb.names) list.appendChild(h('div', { class: 'li' }, h('span', {}, h('b', {}, n.name), n.sheetId ? h('small', {}, ` (${(wb.sheetById(n.sheetId) || {}).name})`) : null), h('span', {}, h('small', { style: 'font-family:Consolas,monospace' }, n.ref), ' ', h('button', { onClick: () => editName(n, k) }, 'Edit'), ' ', h('button', { onClick: () => { app.transact('Delete name', () => wb.names.delete(k) && wb._record(() => wb.names.set(k, n)) && (wb._needsRebuild = true)); render(); } }, 'Delete')))); if (!wb.names.size) list.appendChild(h('div', { class: 'li', style: 'color:var(--fg-dim)' }, 'No defined names')); };
      const editName = (n, k) => {
        const name = h('input', { type: 'text', value: n ? n.name : '', disabled: !!n }); const ref = h('input', { type: 'text', value: n ? n.ref : `=${app.sheet.name.includes(' ') ? `'${app.sheet.name}'` : app.sheet.name}!${rangeName(app.sheet.view.ranges[0]).replace(/([A-Z]+)(\d+)/g, '$$$1$$$2')}` });
        open({ title: n ? 'Edit name' : 'New name', body: h('div', { class: 'form', style: 'width:420px' }, h('label', {}, 'Name'), name, h('label', {}, 'Refers to'), ref), buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, onClick: () => { const nm = name.value.trim(); if (!/^[A-Za-z_\\][A-Za-z0-9_.\\]*$/.test(nm) || /^[A-Za-z]{1,3}\d+$/.test(nm) || nm.toUpperCase() === 'TRUE' || nm.toUpperCase() === 'FALSE') { app.toast('Invalid name: use letters, digits, underscores; not a cell reference'); return false; } try { parse(ref.value.startsWith('=') ? ref.value : '=' + ref.value); } catch (e) { app.toast('Invalid reference: ' + e.message); return false; } app.transact('Define name', () => wb.setName(nm, ref.value)); render(); return true; } }] });
      };
      render();
      return open({ title: 'Name manager', body: h('div', { style: 'width:520px' }, h('div', { style: 'margin-bottom:6px' }, h('button', { onClick: () => editName(null) }, 'New…')), list), buttons: [{ label: 'Close', primary: true, value: true }] });
    },

    goTo() { return D.prompt('Go to', 'Reference or name', '', 'Examples: B12, Sheet2!A1:C5, MyRange').then((v) => { if (v) app.goTo(v); }); },
    pasteSpecial() {
      let mode = 'all';
      const radio = (v, label) => h('label', { style: 'display:block;margin:4px 0' }, h('input', { type: 'radio', name: 'ps', checked: v === 'all', onChange: () => { mode = v; } }), ' ', label);
      const skip = h('input', { type: 'checkbox' });
      return open({ title: 'Paste special', body: h('div', {}, radio('all', 'All'), radio('values', 'Values only'), radio('formats', 'Formats only'), radio('formulas', 'Formulas (no formats)'), radio('transpose', 'Transpose'), h('hr'), h('label', {}, skip, ' Skip blanks')), buttons: [{ label: 'Cancel', value: null }, { label: 'Paste', primary: true, onClick: () => ({ mode, skipBlanks: skip.checked }) }] }).then((o) => { if (o) app.clipboard.paste(null, o); });
    },
    hyperlink() {
      const s = app.sheet; const { r, c } = s.view.active; const cell = s.getCell(r, c);
      const url = h('input', { type: 'text', value: cell && cell.hyperlink || '', placeholder: 'https://… or #Sheet2!A1' });
      const text = h('input', { type: 'text', value: cell && !cell.f && cell.v != null ? String(cell.v) : '' });
      return open({ title: 'Hyperlink', body: h('div', { class: 'form', style: 'width:420px' }, h('label', {}, 'Address'), url, h('label', {}, 'Text to display'), text), buttons: [{ label: 'Remove', left: true, onClick: () => { app.transact('Remove hyperlink', () => s.setHyperlink(r, c, null)); return true; } }, { label: 'Cancel', value: null }, { label: 'OK', primary: true, onClick: () => { app.transact('Hyperlink', () => { if (text.value && (!cell || cell.v == null) || (cell && !cell.f && String(cell.v) !== text.value)) s.setInput(r, c, text.value || url.value); s.setHyperlink(r, c, url.value.trim() || null); }); return true; } }] });
    },
    note(r, c) {
      const s = app.sheet; const cell = s.getCell(r, c);
      const rect = app.grid.cellRect(r, c);
      const ta = h('textarea', { style: 'width:240px;height:90px;background:#fffbe6;color:#000;border:1px solid #e0c060', value: cell && cell.note || '' });
      const pop = h('div', { class: 'popover', style: 'padding:4px' }, ta, h('div', { class: 'row right' }, h('button', { onClick: () => { app.transact('Delete note', () => s.setNote(r, c, null)); close(); } }, 'Delete'), h('button', { class: 'primary', onClick: () => save() }, 'OK')));
      const close = () => { pop.remove(); document.removeEventListener('mousedown', outside, true); app.focusGrid(); };
      const save = () => { const t = ta.value.trim(); if (t !== (cell && cell.note || '')) app.transact('Note', () => s.setNote(r, c, t || null)); close(); };
      const outside = (e) => { if (!pop.contains(e.target)) save(); };
      pop.addEventListener('keydown', (e) => { e.stopPropagation(); if (e.key === 'Escape') close(); });
      app.grid.overlay.appendChild(pop);
      pop.style.left = ((rect ? rect.x + rect.w : 100) + 4) + 'px'; pop.style.top = (rect ? rect.y : 100) + 'px';
      setTimeout(() => { document.addEventListener('mousedown', outside, true); ta.focus(); }, 0);
    },
    colWidth() { const s = app.sheet; const cols = app.selectedCols(); return D.prompt('Column width', 'Width in pixels', String(s.colAxis.rawSize(cols[0]))).then((v) => { const w = parseFloat(v); if (!isNaN(w) && w >= 0) app.transact('Column width', () => cols.forEach((c) => s.setColWidth(c, w))); }); },
    rowHeight() { const s = app.sheet; const rows = app.selectedRows(); return D.prompt('Row height', 'Height in pixels', String(s.rowAxis.rawSize(rows[0]))).then((v) => { const hh = parseFloat(v); if (!isNaN(hh) && hh >= 0) app.transact('Row height', () => rows.forEach((r) => s.setRowHeight(r, hh))); }); },
    customNumber() { const s = app.sheet; const { r, c } = s.view.active; return D.prompt('Custom number format', 'Format code', s.effectiveNumFmt(r, c), 'Excel syntax, e.g. #,##0.00;[Red]-#,##0.00  or  yyyy-mm-dd').then((v) => { if (v != null) app.applyNumFmt(v || 'General'); }); },
    removeDuplicates() {
      const s = app.sheet; const rg = app.dataRegion();
      const header = h('input', { type: 'checkbox', checked: app.looksLikeHeader(rg) });
      const cols = h('div', { class: 'list', style: 'max-height:200px;overflow:auto;border:1px solid var(--border);padding:4px' });
      const cbs = [];
      for (let c = rg.c1; c <= Math.min(rg.c2, rg.c1 + 100); c++) { const cb = h('input', { type: 'checkbox', checked: true }); cbs.push([c, cb]); cols.appendChild(h('label', { style: 'display:block' }, cb, ` Column ${colToName(c)}${s.getValue(rg.r1, c) != null ? ` (${String(s.getValue(rg.r1, c)).slice(0, 24)})` : ''}`)); }
      return open({ title: 'Remove duplicates', body: h('div', { style: 'width:380px' }, h('p', {}, `Range ${rangeName(rg)}. Rows with identical values in the selected columns are removed (first occurrence kept).`), h('label', {}, header, ' My data has headers'), cols), buttons: [{ label: 'Cancel', value: null }, { label: 'Remove', primary: true, onClick: () => { const use = cbs.filter(([, cb]) => cb.checked).map(([c]) => c); const n = app.removeDuplicates(rg, use, header.checked); app.toast(`${n} duplicate row${n === 1 ? '' : 's'} removed`); return true; } }] });
    },
    textToColumns() {
      const delim = sel([[',', 'Comma'], [';', 'Semicolon'], ['\t', 'Tab'], [' ', 'Space'], ['|', 'Pipe'], ['other', 'Other…']], ',');
      const other = h('input', { type: 'text', value: '-', style: 'width:50px' });
      return open({ title: 'Text to columns', body: h('div', { class: 'form' }, h('label', {}, 'Delimiter'), h('div', {}, delim, ' ', other)), buttons: [{ label: 'Cancel', value: null }, { label: 'Split', primary: true, onClick: () => { app.textToColumns(delim.value === 'other' ? other.value : delim.value); return true; } }] });
    },
    insertCells() {
      let mode = 'down';
      const radio = (v, label) => h('label', { style: 'display:block;margin:4px 0' }, h('input', { type: 'radio', name: 'ic', checked: v === 'down', onChange: () => { mode = v; } }), ' ', label);
      return open({ title: 'Insert cells', body: h('div', {}, radio('down', 'Shift cells down'), radio('right', 'Shift cells right'), radio('rows', 'Entire row'), radio('cols', 'Entire column')), buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, onClick: () => mode }] });
    },
    deleteCells() {
      let mode = 'up';
      const radio = (v, label) => h('label', { style: 'display:block;margin:4px 0' }, h('input', { type: 'radio', name: 'dc', checked: v === 'up', onChange: () => { mode = v; } }), ' ', label);
      return open({ title: 'Delete cells', body: h('div', {}, radio('up', 'Shift cells up'), radio('left', 'Shift cells left'), radio('rows', 'Entire row'), radio('cols', 'Entire column')), buttons: [{ label: 'Cancel', value: null }, { label: 'OK', primary: true, onClick: () => mode }] });
    },
    help() {
      const rows = [['Arrows / Tab / Enter', 'Move active cell'], ['Shift + arrows', 'Extend selection'], ['Ctrl + arrows', 'Jump to edge of data'], ['Ctrl+Home / Ctrl+End', 'First cell / last used cell'], ['Page Up / Down (Alt: sideways)', 'Scroll a page'], ['Ctrl+Space / Shift+Space', 'Select column / row'], ['Ctrl+A', 'Select region, then whole sheet'], ['F2', 'Edit cell (F2 again toggles enter mode)'], ['Esc', 'Cancel edit or copy'], ['Delete / Backspace', 'Clear contents / edit empty'], ['Alt+Enter', 'New line in cell'], ['Ctrl+Enter', 'Fill selection with entry'], ['F4 (editing)', 'Toggle $ absolute references'], ['Ctrl+C / X / V', 'Copy / cut / paste'], ['Ctrl+Shift+V', 'Paste values only'], ['Ctrl+Z / Ctrl+Y', 'Undo / redo'], ['Ctrl+D / Ctrl+R', 'Fill down / right'], ['Ctrl+B / I / U, Ctrl+5', 'Bold / italic / underline / strikethrough'], ['Ctrl+1', 'Format cells'], ['Ctrl+Shift+1 / 4 / 5', 'Number / currency / percent format'], ['Ctrl+Shift+3 / 2 / ~', 'Date / time / general format'], ['Ctrl+; / Ctrl+Shift+;', 'Insert date / time'], ['Ctrl+F / Ctrl+H / Ctrl+G', 'Find / replace / go to'], ['Ctrl+K', 'Hyperlink'], ['Shift+F2', 'Note'], ['Ctrl+Shift+L', 'Toggle filter'], ['Alt+=', 'AutoSum'], ['Ctrl+` ', 'Show formulas'], ['Ctrl+Plus / Ctrl+Minus', 'Insert / delete cells'], ['Ctrl+9 / Ctrl+0', 'Hide rows / columns'], ['Ctrl+Shift+9 / Ctrl+Shift+0', 'Unhide rows / columns'], ['Ctrl+PgUp / PgDn', 'Previous / next sheet'], ['Shift+F11', 'New sheet'], ['F9', 'Recalculate'], ['Ctrl+S / O / N / P', 'Save / open / new / print'], ['Ctrl + mouse wheel', 'Zoom'], ['Drag fill handle (Ctrl: force increment)', 'Autofill series']];
      return open({ title: 'Keyboard shortcuts', body: h('div', { class: 'help-grid', style: 'max-height:60vh;overflow:auto' }, rows.flatMap(([k, d]) => [h('kbd', {}, k), h('span', {}, d)])), buttons: [{ label: 'Close', primary: true, value: true }] });
    },
    functionList() {
      const cats = {}; for (const f of FUNCTION_LIST) (cats[f.cat] ||= new Set()).add(f.name);
      const body = h('div', { style: 'width:620px;max-height:60vh;overflow:auto;font-size:12px' }, Object.entries(cats).sort().map(([cat, names]) => h('div', {}, h('h4', {}, `${cat} (${names.size})`), h('div', { style: 'font-family:Consolas,monospace;line-height:1.7' }, [...names].sort().map((n) => h('span', { style: 'display:inline-block;margin-right:10px' }, h('b', {}, n), h('span', { style: 'color:var(--fg-dim)' }, `(${FUNCTIONS.get(n).sig})`)))))));
      return open({ title: `Functions (${new Set(FUNCTION_LIST.map((f) => f.name)).size})`, body, buttons: [{ label: 'Close', primary: true, value: true }] });
    },
    about() { return open({ title: 'About OfficeTable', body: h('div', { style: 'max-width:420px;line-height:1.5' }, h('p', {}, h('b', {}, 'OfficeTable'), ' — a spreadsheet that runs entirely in your browser. No server, no install; your data stays on your computer.'), h('p', {}, '300+ worksheet functions with dynamic arrays, number formats, charts, sort & filter, conditional formatting, data validation, and .xlsx / .csv import and export.'), h('p', {}, h('a', { href: 'https://github.com/sypark-us/robbins_table', target: '_blank' }, 'github.com/sypark-us/robbins_table'), ' · MIT license')), buttons: [{ label: 'Close', primary: true, value: true }] }); },

    // ---------------- Color picker popover ----------------
    colorPicker(anchorEl, current, onPick) {
      document.querySelectorAll('.colorpop').forEach((p) => p.remove());
      const pop = h('div', { class: 'colorpop' });
      const grid = h('div', { class: 'swatches' });
      const pick = (c) => { pop.remove(); document.removeEventListener('mousedown', outside, true); onPick(c); app.focusGrid(); };
      const outside = (e) => { if (!pop.contains(e.target)) { pop.remove(); document.removeEventListener('mousedown', outside, true); } };
      for (const c of THEME_COLORS) grid.appendChild(h('div', { class: 'sw', style: `background:${c}`, title: c, onClick: () => pick(c) }));
      for (const t of SHADES) for (const c of THEME_COLORS) { const sc = shade(c, t); grid.appendChild(h('div', { class: 'sw', style: `background:${sc}`, title: sc, onClick: () => pick(sc) })); }
      const std = h('div', { class: 'swatches' }, STANDARD_COLORS.map((c) => h('div', { class: 'sw', style: `background:${c}`, title: c, onClick: () => pick(c) })));
      const custom = h('input', { type: 'color', value: current && /^#[0-9a-f]{6}$/i.test(current) ? current : '#000000' });
      pop.append(h('div', { style: 'font-size:11px;color:var(--fg-dim);margin-bottom:4px' }, 'Theme colors'), grid, h('div', { style: 'font-size:11px;color:var(--fg-dim);margin:6px 0 4px' }, 'Standard colors'), std, h('div', { class: 'row' }, h('div', { class: 'sw none', style: 'width:18px;height:18px;border:1px solid #666;cursor:pointer', title: 'No color', onClick: () => pick(null) }), h('span', { style: 'cursor:pointer', onClick: () => pick(null) }, 'No color'), h('span', { style: 'margin-left:auto' }, 'Custom: '), custom, h('button', { onClick: () => pick(custom.value) }, 'OK')));
      document.body.appendChild(pop);
      const r = anchorEl.getBoundingClientRect();
      pop.style.left = Math.min(r.left, window.innerWidth - pop.offsetWidth - 8) + 'px'; pop.style.top = (r.bottom + 4) + 'px';
      setTimeout(() => document.addEventListener('mousedown', outside, true), 0);
    },
  };
  function colorField(initial, noneLabel) {
    let value = initial;
    const sw = h('span', { class: 'sw', style: `display:inline-block;width:22px;height:18px;border:1px solid #666;vertical-align:middle;border-radius:2px;background:${value || 'transparent'}` });
    const lbl = h('span', { style: 'margin-left:6px;color:var(--fg-dim)' }, value || noneLabel || '');
    const btn = h('button', { onClick: (e) => { e.preventDefault(); D.colorPicker(btn, value, (c) => { if (c === null && noneLabel === null) return; field.set(c); if (field.onChange) field.onChange(); }); } }, sw, lbl);
    const field = { el: btn, get: () => value, set: (c) => { value = c; sw.style.background = c || 'transparent'; lbl.textContent = c || noneLabel || ''; }, onChange: null };
    return field;
  }
  return D;
}
