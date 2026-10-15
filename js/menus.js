// Menu bar, toolbar and context menus, all driven by app.commands.
import { escapeHtml } from './util.js';
import { FORMAT_PRESETS } from './format.js';

const svg = (d, extra = '') => `<svg viewBox="0 0 24 24">${d}${extra}</svg>`;
export const ICONS = {
  new: svg('<path d="M6 3h8l4 4v14H6z"/><path d="M14 3v4h4"/><path d="M12 11v6M9 14h6"/>'),
  open: svg('<path d="M3 7h6l2 2h10v10H3z"/><path d="M3 11h18"/>'),
  save: svg('<path d="M5 3h11l3 3v15H5z"/><path d="M8 3v5h7V3"/><path d="M8 21v-6h8v6"/>'),
  undo: svg('<path d="M9 14 4 9l5-5"/><path d="M4 9h10a6 6 0 0 1 0 12h-3"/>'),
  redo: svg('<path d="m15 14 5-5-5-5"/><path d="M20 9H10a6 6 0 0 0 0 12h3"/>'),
  alignLeft: svg('<path d="M4 6h16M4 10h10M4 14h16M4 18h10"/>'),
  alignCenter: svg('<path d="M4 6h16M7 10h10M4 14h16M7 18h10"/>'),
  alignRight: svg('<path d="M4 6h16M10 10h10M4 14h16M10 18h10"/>'),
  valignTop: svg('<path d="M4 4h16"/><path d="M8 8h8v10H8z"/>'),
  valignMiddle: svg('<path d="M4 12h3M17 12h3"/><path d="M8 6h8v12H8z"/>'),
  valignBottom: svg('<path d="M4 20h16"/><path d="M8 6h8v10H8z"/>'),
  wrap: svg('<path d="M4 6h16M4 12h11a3 3 0 0 1 0 6h-3"/><path d="m14 15-2 3 2 3"/><path d="M4 18h5"/>'),
  merge: svg('<path d="M4 5h16v14H4z"/><path d="M4 12h16"/><path d="M12 5v14" stroke-dasharray="2 2"/>'),
  borders: svg('<path d="M4 4h16v16H4z"/><path d="M4 12h16M12 4v16" stroke-dasharray="2 2"/>'),
  fill: svg('<path d="m7 11 6-6 6 6-6 6z"/><path d="M4 20h16"/><path d="M3 15c0 1.5 1 2.5 2 2.5s2-1 2-2.5-2-3-2-3-2 1.5-2 3z"/>'),
  fontColor: svg('<path d="m6 17 6-13 6 13"/><path d="M8.5 12h7"/><path d="M4 21h16"/>'),
  sortAZ: svg('<path d="M6 4v16"/><path d="m3 17 3 3 3-3"/><path d="M12 10h6l-6 7h6"/><path d="M12 5h4"/>'),
  sortZA: svg('<path d="M6 4v16"/><path d="m3 17 3 3 3-3"/><path d="M12 4h6l-6 7h6"/><path d="M12 17h4"/>'),
  filter: svg('<path d="M4 5h16l-6 8v6l-4-2v-4z"/>'),
  chart: svg('<path d="M4 20h16"/><path d="M6 16v-5M11 16V7M16 16v-9"/>'),
  sum: svg('<path d="M17 5H7l6 7-6 7h10"/>'),
  freeze: svg('<path d="M4 4h16v16H4z"/><path d="M4 9h16M9 4v16"/>'),
  find: svg('<circle cx="10.5" cy="10.5" r="6"/><path d="m15 15 5 5"/>'),
  percent: svg('<path d="m6 18 12-12"/><circle cx="7.5" cy="7.5" r="2.5"/><circle cx="16.5" cy="16.5" r="2.5"/>'),
  comma: svg('<path d="M12 9a2 2 0 1 1 2 2c0 2-1 4-2 5"/>'),
  decInc: svg('<text x="2" y="16" font-size="9" fill="currentColor" stroke="none" font-family="Segoe UI">.00</text><path d="M14 13h6M17 10v6"/>'),
  decDec: svg('<text x="2" y="16" font-size="9" fill="currentColor" stroke="none" font-family="Segoe UI">.0</text><path d="M14 13h6"/>'),
  currency: svg('<path d="M12 3v18"/><path d="M16 7.5c0-1.5-1.8-2.5-4-2.5S8 6 8 7.5 9.5 10 12 10s4 1 4 2.5-1.8 2.5-4 2.5-4-1-4-2.5" transform="translate(0 2)"/>'),
  indentInc: svg('<path d="M4 6h16M10 10h10M10 14h10M4 18h16"/><path d="m4 9 3 3-3 3"/>'),
  indentDec: svg('<path d="M4 6h16M10 10h10M10 14h10M4 18h16"/><path d="m7 9-3 3 3 3"/>'),
  paint: svg('<path d="M5 4h12v4H5z"/><path d="M17 6h2v5H10v9"/>'),
  cut: svg('<circle cx="7" cy="17" r="3"/><circle cx="17" cy="17" r="3"/><path d="M9.5 14.5 18 4M14.5 14.5 6 4"/>'),
  copy: svg('<path d="M9 9h11v11H9z"/><path d="M5 15H4V4h11v1"/>'),
  paste: svg('<path d="M8 5H5v16h14V5h-3"/><path d="M9 3h6v4H9z"/>'),
  table: svg('<path d="M4 5h16v14H4z"/><path d="M4 10h16M4 15h16M10 5v14"/>'),
  note: svg('<path d="M4 5h16v11H9l-5 4z"/>'),
  link: svg('<path d="M10 14a4 4 0 0 1 0-5.6l2.4-2.4a4 4 0 1 1 5.6 5.6L17 12.6"/><path d="M14 10a4 4 0 0 1 0 5.6l-2.4 2.4a4 4 0 1 1-5.6-5.6L7 11.4"/>'),
  print: svg('<path d="M7 8V3h10v5"/><path d="M4 8h16v9h-3v4H7v-4H4z"/><path d="M7 13h10"/>'),
  fx: svg('<path d="M14 4c-3 0-3 2-3.5 5L9 20c-.3 2-1.5 2-3 1"/><path d="M8 9h7"/><path d="m14 13 5 6M19 13l-5 6"/>'),
};

export const MENUS = [
  { label: 'File', items: [{ cmd: 'file.new' }, { cmd: 'file.open' }, { sep: true }, { cmd: 'file.saveXlsx' }, { cmd: 'file.saveJson' }, { cmd: 'file.exportCsv' }, { cmd: 'file.exportTsv' }, { sep: true }, { cmd: 'file.print' }, { sep: true }, { cmd: 'file.autosave' }, { cmd: 'file.clearStorage' }] },
  { label: 'Edit', items: [{ cmd: 'edit.undo' }, { cmd: 'edit.redo' }, { sep: true }, { cmd: 'edit.cut' }, { cmd: 'edit.copy' }, { cmd: 'edit.paste' }, { cmd: 'edit.pasteSpecial' }, { cmd: 'edit.pasteValues' }, { sep: true }, { cmd: 'edit.fillDown' }, { cmd: 'edit.fillRight' }, { sep: true }, { label: 'Clear', sub: [{ cmd: 'edit.clearContents' }, { cmd: 'edit.clearFormats' }, { cmd: 'edit.clearNotes' }, { cmd: 'edit.clearAll' }] }, { label: 'Delete', sub: [{ cmd: 'edit.deleteRows' }, { cmd: 'edit.deleteCols' }, { cmd: 'edit.deleteCellsUp' }, { cmd: 'edit.deleteCellsLeft' }] }, { sep: true }, { cmd: 'edit.find' }, { cmd: 'edit.replace' }, { cmd: 'edit.goto' }, { cmd: 'edit.selectAll' }] },
  { label: 'View', items: [{ label: 'Freeze panes', sub: [{ cmd: 'view.freezeSel' }, { cmd: 'view.freezeRow' }, { cmd: 'view.freezeCol' }, { cmd: 'view.unfreeze' }] }, { cmd: 'view.gridlines' }, { cmd: 'view.formulas' }, { sep: true }, { cmd: 'view.zoomIn' }, { cmd: 'view.zoomOut' }, { cmd: 'view.zoom100' }, { sep: true }, { cmd: 'view.hideRows' }, { cmd: 'view.hideCols' }, { cmd: 'view.unhideRows' }, { cmd: 'view.unhideCols' }] },
  { label: 'Insert', items: [{ cmd: 'insert.rowsAbove' }, { cmd: 'insert.rowsBelow' }, { cmd: 'insert.colsLeft' }, { cmd: 'insert.colsRight' }, { label: 'Cells', sub: [{ cmd: 'insert.cellsDown' }, { cmd: 'insert.cellsRight' }] }, { sep: true }, { cmd: 'insert.sheet' }, { sep: true }, { cmd: 'insert.chart' }, { cmd: 'insert.function' }, { cmd: 'insert.name' }, { sep: true }, { cmd: 'insert.note' }, { cmd: 'insert.hyperlink' }, { cmd: 'insert.date' }, { cmd: 'insert.time' }, { cmd: 'insert.checkbox' }] },
  { label: 'Format', items: [{ cmd: 'format.cells' }, { label: 'Number', sub: FORMAT_PRESETS.map((p) => ({ label: p.name, run: (app) => app.applyNumFmt(p.code) })).concat([{ sep: true }, { cmd: 'format.customNumber' }]) }, { sep: true }, { cmd: 'format.bold' }, { cmd: 'format.italic' }, { cmd: 'format.underline' }, { cmd: 'format.strike' }, { label: 'Align', sub: [{ cmd: 'format.alignLeft' }, { cmd: 'format.alignCenter' }, { cmd: 'format.alignRight' }, { sep: true }, { cmd: 'format.valignTop' }, { cmd: 'format.valignMiddle' }, { cmd: 'format.valignBottom' }] }, { cmd: 'format.wrap' }, { cmd: 'format.merge' }, { cmd: 'format.mergeAcross' }, { cmd: 'format.unmerge' }, { sep: true }, { cmd: 'format.colWidth' }, { cmd: 'format.rowHeight' }, { cmd: 'format.autofitCols' }, { cmd: 'format.autofitRows' }, { sep: true }, { cmd: 'format.cf' }, { cmd: 'format.clearCf' }, { cmd: 'edit.clearFormats' }] },
  { label: 'Data', items: [{ cmd: 'data.sortAsc' }, { cmd: 'data.sortDesc' }, { cmd: 'data.sort' }, { sep: true }, { cmd: 'data.filter' }, { cmd: 'data.clearFilter' }, { sep: true }, { cmd: 'data.validation' }, { cmd: 'data.removeDuplicates' }, { cmd: 'data.textToColumns' }, { cmd: 'data.trim' }, { sep: true }, { cmd: 'data.names' }, { cmd: 'data.recalc' }] },
  { label: 'Help', items: [{ cmd: 'help.shortcuts' }, { cmd: 'help.functions' }, { cmd: 'help.about' }] },
];

export const TOOLBAR = [
  { cmd: 'file.new', icon: 'new' }, { cmd: 'file.open', icon: 'open' }, { cmd: 'file.saveXlsx', icon: 'save' }, { sep: true },
  { cmd: 'edit.undo', icon: 'undo' }, { cmd: 'edit.redo', icon: 'redo' }, { sep: true },
  { cmd: 'edit.cut', icon: 'cut' }, { cmd: 'edit.copy', icon: 'copy' }, { cmd: 'edit.paste', icon: 'paste' }, { cmd: 'format.painter', icon: 'paint' }, { sep: true },
  { select: 'fontName', title: 'Font', options: ['Calibri', 'Arial', 'Aptos', 'Cambria', 'Consolas', 'Courier New', 'Georgia', 'Segoe UI', 'Tahoma', 'Times New Roman', 'Trebuchet MS', 'Verdana'] },
  { select: 'fontSize', title: 'Font size', options: [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72] }, { sep: true },
  { cmd: 'format.bold', text: 'B', cls: 'text', style: 'font-weight:700' }, { cmd: 'format.italic', text: 'I', cls: 'text', style: 'font-style:italic;font-family:serif' }, { cmd: 'format.underline', text: 'U', cls: 'text', style: 'text-decoration:underline' }, { cmd: 'format.strike', text: 'S', cls: 'text', style: 'text-decoration:line-through' },
  { color: 'fontColor', icon: 'fontColor', title: 'Font color', prop: 'color' }, { color: 'fillColor', icon: 'fill', title: 'Fill color', prop: 'fill' }, { cmd: 'format.borders', icon: 'borders' }, { sep: true },
  { cmd: 'format.alignLeft', icon: 'alignLeft' }, { cmd: 'format.alignCenter', icon: 'alignCenter' }, { cmd: 'format.alignRight', icon: 'alignRight' }, { cmd: 'format.valignTop', icon: 'valignTop' }, { cmd: 'format.valignMiddle', icon: 'valignMiddle' }, { cmd: 'format.valignBottom', icon: 'valignBottom' }, { cmd: 'format.indentDec', icon: 'indentDec' }, { cmd: 'format.indentInc', icon: 'indentInc' }, { cmd: 'format.wrap', icon: 'wrap' }, { cmd: 'format.merge', icon: 'merge' }, { sep: true },
  { select: 'numFmtSel', title: 'Number format', options: FORMAT_PRESETS.map((p) => p.name).concat(['Custom…']) },
  { cmd: 'format.currency', icon: 'currency' }, { cmd: 'format.percent', icon: 'percent' }, { cmd: 'format.comma', icon: 'comma' }, { cmd: 'format.decInc', icon: 'decInc' }, { cmd: 'format.decDec', icon: 'decDec' }, { sep: true },
  { cmd: 'data.sortAsc', icon: 'sortAZ' }, { cmd: 'data.sortDesc', icon: 'sortZA' }, { cmd: 'data.filter', icon: 'filter' }, { sep: true },
  { cmd: 'insert.chart', icon: 'chart' }, { cmd: 'insert.function', icon: 'fx' }, { cmd: 'formulas.autosum', icon: 'sum' }, { cmd: 'view.freezeSel', icon: 'freeze' }, { cmd: 'insert.note', icon: 'note' }, { cmd: 'insert.hyperlink', icon: 'link' }, { cmd: 'edit.find', icon: 'find' },
];

export function buildMenus(app) {
  const bar = document.getElementById('menus');
  bar.innerHTML = '';
  for (const m of MENUS) {
    const el = document.createElement('div'); el.className = 'menu';
    const label = document.createElement('div'); label.className = 'label'; label.textContent = m.label; el.appendChild(label);
    const dd = document.createElement('div'); dd.className = 'dropdown'; el.appendChild(dd);
    label.addEventListener('mousedown', (e) => { e.preventDefault(); e.stopPropagation(); const open = el.classList.contains('open'); closeAll(); if (!open) { fillDropdown(app, dd, m.items); el.classList.add('open'); } });
    label.addEventListener('mouseenter', () => { if ([...bar.children].some((x) => x.classList.contains('open')) && !el.classList.contains('open')) { closeAll(); fillDropdown(app, dd, m.items); el.classList.add('open'); } });
    bar.appendChild(el);
  }
  document.addEventListener('mousedown', (e) => { if (!e.target.closest('.menu') && !e.target.closest('#contextMenu')) closeAll(); });
  const ctx = document.getElementById('contextMenu');
  function closeAll() { for (const el of bar.children) el.classList.remove('open'); ctx.classList.remove('open'); }
  app.closeMenus = closeAll;
  app.showContextMenu = (items, x, y) => {
    closeAll();
    fillDropdown(app, ctx, items);
    ctx.classList.add('open');
    ctx.style.left = '0px'; ctx.style.top = '0px';
    const w = ctx.offsetWidth, h = ctx.offsetHeight;
    ctx.style.left = Math.min(x, window.innerWidth - w - 4) + 'px';
    ctx.style.top = Math.min(y, window.innerHeight - h - 4) + 'px';
  };
  buildToolbar(app);
}

function fillDropdown(app, dd, items) {
  dd.innerHTML = '';
  for (const it of items) {
    if (it.sep) { const s = document.createElement('div'); s.className = 'sep'; dd.appendChild(s); continue; }
    if (it.title) { const t = document.createElement('div'); t.className = 'title'; t.textContent = it.title; dd.appendChild(t); continue; }
    const cmd = it.cmd ? app.commands[it.cmd] : null;
    if (it.cmd && !cmd) continue;
    const el = document.createElement('div'); el.className = 'item';
    const label = it.label || (cmd && cmd.label) || it.cmd;
    const checked = it.checked != null ? (typeof it.checked === 'function' ? it.checked(app) : it.checked) : cmd && cmd.checked ? cmd.checked() : null;
    const disabled = it.disabled != null ? (typeof it.disabled === 'function' ? it.disabled(app) : it.disabled) : cmd && cmd.enabled ? !cmd.enabled() : false;
    if (disabled) el.classList.add('disabled');
    el.innerHTML = `<span>${checked != null ? `<span class="check">${checked ? '✓' : ''}</span>` : ''}${escapeHtml(label)}</span>${cmd && cmd.sc ? `<span class="sc">${escapeHtml(cmd.sc)}</span>` : ''}`;
    if (it.sub) { el.classList.add('submenu'); const sub = document.createElement('div'); sub.className = 'dropdown'; el.appendChild(sub); el.addEventListener('mouseenter', () => fillDropdown(app, sub, it.sub)); }
    else el.addEventListener('mouseup', (e) => { e.stopPropagation(); app.closeMenus(); if (it.run) it.run(app); else if (it.cmd) app.run(it.cmd); });
    el.addEventListener('mousedown', (e) => e.preventDefault());
    dd.appendChild(el);
  }
}

function buildToolbar(app) {
  const tb = document.getElementById('toolbar');
  tb.innerHTML = '';
  app.toolbarButtons = {};
  for (const it of TOOLBAR) {
    if (it.sep) { const s = document.createElement('div'); s.className = 'sep'; tb.appendChild(s); continue; }
    if (it.select) {
      const sel = document.createElement('select'); sel.id = it.select; sel.title = it.title;
      for (const o of it.options) { const op = document.createElement('option'); op.value = o; op.textContent = o; sel.appendChild(op); }
      sel.addEventListener('mousedown', (e) => e.stopPropagation());
      sel.addEventListener('change', () => { app.onToolbarSelect(it.select, sel.value); app.focusGrid(); });
      tb.appendChild(sel); continue;
    }
    if (it.color) {
      const b = document.createElement('button'); b.className = 'tb color'; b.title = it.title; b.innerHTML = ICONS[it.icon] + `<span class="swatch" style="background:${it.prop === 'fill' ? '#ffff00' : '#ff0000'}"></span>`;
      b.dataset.color = it.prop === 'fill' ? '#ffff00' : '#ff0000';
      b.addEventListener('mousedown', (e) => e.preventDefault());
      b.addEventListener('click', () => { app.applyStyle({ [it.prop]: b.dataset.color }); });
      const arrow = document.createElement('button'); arrow.className = 'tb arrow'; arrow.textContent = '▾'; arrow.title = it.title + ' picker';
      arrow.addEventListener('mousedown', (e) => e.preventDefault());
      arrow.addEventListener('click', () => app.dialogs.colorPicker(arrow, b.dataset.color, (c) => { if (c) { b.dataset.color = c; b.querySelector('.swatch').style.background = c; } app.applyStyle({ [it.prop]: c || null }); }));
      const wrap = document.createElement('span'); wrap.className = 'split'; wrap.appendChild(b); wrap.appendChild(arrow); tb.appendChild(wrap); continue;
    }
    const cmd = app.commands[it.cmd]; if (!cmd) continue;
    const b = document.createElement('button'); b.className = 'tb' + (it.cls ? ' ' + it.cls : ''); b.title = cmd.label + (cmd.sc ? ` (${cmd.sc})` : '');
    if (it.icon) b.innerHTML = ICONS[it.icon]; else b.textContent = it.text || cmd.label;
    if (it.style) b.style.cssText += it.style;
    b.addEventListener('mousedown', (e) => e.preventDefault());
    b.addEventListener('click', (e) => { app.run(it.cmd, e); });
    tb.appendChild(b);
    app.toolbarButtons[it.cmd] = b;
  }
}

export function refreshToolbar(app) {
  const s = app.sheet; const { r, c } = s.view.active;
  const st = s.getStyle(r, c);
  for (const [name, b] of Object.entries(app.toolbarButtons)) {
    const cmd = app.commands[name];
    if (cmd.checked) b.classList.toggle('active', !!cmd.checked());
    if (cmd.enabled) b.disabled = !cmd.enabled();
  }
  const fn = document.getElementById('fontName'); if (fn) { const v = st.font || 'Calibri'; if (![...fn.options].some((o) => o.value === v)) { const op = document.createElement('option'); op.value = v; op.textContent = v; fn.appendChild(op); } fn.value = v; }
  const fs = document.getElementById('fontSize'); if (fs) { const v = String(st.size || 11); if (![...fs.options].some((o) => o.value === v)) { const op = document.createElement('option'); op.value = v; op.textContent = v; fs.appendChild(op); } fs.value = v; }
  const nf = document.getElementById('numFmtSel'); if (nf) { const code = s.effectiveNumFmt(r, c); const p = FORMAT_PRESETS.find((x) => x.code === code); nf.value = p ? p.name : (code === 'General' ? 'General' : 'Custom…'); nf.title = 'Number format: ' + code; }
}
