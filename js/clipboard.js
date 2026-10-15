// Copy / cut / paste, including interchange with other spreadsheets via text/html and TSV.
import { shiftFormula } from './formula.js';
import { parseInput, formatValue, isDateFormat } from './format.js';
import { escapeHtml, normRange, MAX_ROWS, MAX_COLS } from './util.js';
import { mergeStyle } from './model.js';

const MARK = 'data-officetable';

export class ClipboardManager {
  constructor(app) {
    this.app = app;
    this.internal = null; // { sheet, range, cells: [[cell]], cut, merges }
    this.lastHtml = null;
  }

  // Snapshot the current selection
  copy(cut = false) {
    const app = this.app; const sheet = app.sheet; const rg = app.sel.ranges[0];
    const r2 = Math.min(rg.r2, Math.max(rg.r1, sheet.usedBounds().r2 + 1)), c2 = Math.min(rg.c2, Math.max(rg.c1, sheet.usedBounds().c2 + 1));
    const range = { r1: rg.r1, c1: rg.c1, r2, c2 };
    const cells = [];
    const rowsIncluded = [];
    for (let r = range.r1; r <= range.r2; r++) {
      if (sheet.rowAxis.filtered.has(r)) continue; // filtered rows are skipped like Excel does for visible-only copy
      rowsIncluded.push(r);
      const row = [];
      for (let c = range.c1; c <= range.c2; c++) { const cell = sheet.getCell(r, c); const sid = sheet.getStyleId(r, c); row.push({ v: cell ? (cell.f ? sheet.getValue(r, c) : cell.v) : null, raw: cell && !cell.f ? cell.v : null, f: cell ? cell.f : null, s: sid, note: cell ? cell.note : null, hyperlink: cell ? cell.hyperlink : null, text: app.fmt(sheet, r, c).text }); }
      cells.push(row);
    }
    const merges = sheet.merges.filter((m) => m.r1 >= range.r1 && m.r2 <= range.r2 && m.c1 >= range.c1 && m.c2 <= range.c2).map((m) => ({ r1: m.r1 - range.r1, c1: m.c1 - range.c1, r2: m.r2 - range.r1, c2: m.c2 - range.c1 }));
    this.internal = { sheet, range, cells, cut, merges, rows: rowsIncluded };
    app.grid.copyRange = { ...range, sheetId: sheet.id, cut };
    app.grid.invalidate();
    const html = this.toHtml(this.internal);
    const text = cells.map((row) => row.map((c) => c.text.replace(/\t/g, ' ')).join('\t')).join('\n');
    this.lastHtml = html;
    this.writeSystem(text, html);
    app.setStatus(cut ? 'Select destination and press Enter or Ctrl+V' : `Copied ${cells.length} × ${cells[0] ? cells[0].length : 0}`);
  }

  async writeSystem(text, html) {
    try {
      if (navigator.clipboard && window.ClipboardItem) {
        await navigator.clipboard.write([new ClipboardItem({ 'text/plain': new Blob([text], { type: 'text/plain' }), 'text/html': new Blob([html], { type: 'text/html' }) })]);
        return;
      }
      if (navigator.clipboard) await navigator.clipboard.writeText(text);
    } catch { /* clipboard permission denied: internal clipboard still works */ }
  }

  toHtml(snap) {
    const wb = this.app.wb;
    const trs = snap.cells.map((row) => '<tr>' + row.map((c) => {
      const st = wb.styles[c.s] || {};
      const css = [];
      if (st.bold) css.push('font-weight:bold'); if (st.italic) css.push('font-style:italic'); if (st.underline) css.push('text-decoration:underline'); if (st.strike) css.push('text-decoration:line-through');
      if (st.color) css.push(`color:${st.color}`); if (st.fill) css.push(`background:${st.fill}`); if (st.align) css.push(`text-align:${st.align}`); if (st.font) css.push(`font-family:${st.font}`); if (st.size) css.push(`font-size:${st.size}pt`);
      const attrs = [`data-v="${escapeHtml(JSON.stringify(c.raw))}"`, c.f ? `data-f="${escapeHtml(c.f)}"` : '', `data-s="${escapeHtml(JSON.stringify(st))}"`, c.note ? `data-note="${escapeHtml(c.note)}"` : '', c.hyperlink ? `data-href="${escapeHtml(c.hyperlink)}"` : ''].filter(Boolean).join(' ');
      return `<td ${attrs}${css.length ? ` style="${css.join(';')}"` : ''}>${escapeHtml(c.text).replace(/\n/g, '<br>')}</td>`;
    }).join('') + '</tr>').join('');
    return `<table ${MARK}="1" data-merges="${escapeHtml(JSON.stringify(snap.merges))}"><tbody>${trs}</tbody></table>`;
  }

  // Called from the document 'paste' event (preferred, has access to clipboard data) or from a menu command.
  async paste(e, options = {}) {
    const app = this.app;
    let html = null, text = null;
    if (e && e.clipboardData) { html = e.clipboardData.getData('text/html') || null; text = e.clipboardData.getData('text/plain') || null; }
    else if (navigator.clipboard && navigator.clipboard.read) {
      try { const items = await navigator.clipboard.read(); for (const it of items) { if (it.types.includes('text/html')) html = await (await it.getType('text/html')).text(); if (it.types.includes('text/plain')) text = await (await it.getType('text/plain')).text(); } } catch { /* fall back */ }
      if (html == null && text == null && navigator.clipboard.readText) { try { text = await navigator.clipboard.readText(); } catch { /* ignore */ } }
    }
    // Decide source: internal snapshot if the system html is ours (or unavailable)
    let grid;
    const ours = html && html.includes(MARK);
    if (this.internal && (ours || (!html && !text) || (text && this.internal.cells.map((r) => r.map((c) => c.text.replace(/\t/g, ' ')).join('\t')).join('\n') === text.replace(/\r\n/g, '\n').replace(/\n$/, '')))) grid = this.fromInternal();
    else if (html && /<table/i.test(html)) grid = this.fromHtml(html);
    else if (text) grid = this.fromText(text);
    else if (this.internal) grid = this.fromInternal();
    if (!grid || !grid.cells.length) return false;
    this.applyGrid(grid, options);
    if (grid.cut && this.internal) { this.internal = null; app.grid.copyRange = null; }
    return true;
  }

  fromInternal() {
    const s = this.internal;
    return { cells: s.cells.map((row) => row.map((c) => ({ ...c }))), merges: s.merges, cut: s.cut, srcSheet: s.sheet, srcRange: s.range, internal: true };
  }
  fromText(text) {
    const lines = text.replace(/\r\n/g, '\n').replace(/\n$/, '').split('\n');
    const cells = lines.map((l) => l.split('\t').map((t) => ({ text: t, raw: undefined, f: null, s: null })));
    return { cells, merges: [] };
  }
  fromHtml(html) {
    const doc = new DOMParser().parseFromString(html, 'text/html');
    const table = doc.querySelector('table'); if (!table) return null;
    const ours = table.hasAttribute(MARK);
    const rows = [...table.querySelectorAll('tr')];
    const cells = []; const spans = []; // handle rowspan/colspan by filling placeholders
    const occupied = new Map();
    rows.forEach((tr, ri) => {
      const row = []; let ci = 0;
      for (const td of tr.children) {
        if (!/^t[dh]$/i.test(td.tagName)) continue;
        while (occupied.get(`${ri},${ci}`)) { row[ci] = { text: '', raw: null, f: null, s: null }; ci++; }
        const rs = +td.getAttribute('rowspan') || 1, cs = +td.getAttribute('colspan') || 1;
        const cell = { text: (td.innerText ?? td.textContent).replace(/ /g, ' ').replace(/\r?\n$/, ''), f: null, s: null, raw: undefined };
        if (ours) { try { cell.raw = JSON.parse(td.getAttribute('data-v')); } catch { cell.raw = undefined; } cell.f = td.getAttribute('data-f') || null; try { cell.style = JSON.parse(td.getAttribute('data-s') || '{}'); } catch { cell.style = null; } cell.note = td.getAttribute('data-note'); cell.hyperlink = td.getAttribute('data-href'); }
        else {
          const st = {}; const cssText = td.getAttribute('style') || '';
          const fw = /font-weight:\s*(bold|[6-9]00)/i.test(cssText) || td.querySelector('b,strong') || td.tagName === 'TH'; if (fw) st.bold = true;
          if (/font-style:\s*italic/i.test(cssText) || td.querySelector('i,em')) st.italic = true;
          const color = /(?:^|;)\s*color:\s*([^;]+)/i.exec(cssText); if (color) { const hx = cssToHex(color[1]); if (hx && hx !== '#000000') st.color = hx; }
          const bg = /background(?:-color)?:\s*([^;]+)/i.exec(cssText); if (bg) { const hx = cssToHex(bg[1]); if (hx && hx !== '#ffffff') st.fill = hx; }
          const ta = /text-align:\s*(left|center|right)/i.exec(cssText); if (ta) st.align = ta[1].toLowerCase();
          const nf = td.getAttribute('data-sheets-numberformat'); // Google Sheets
          const xnum = /mso-number-format:\s*"?([^;"]+)"?/i.exec(cssText); if (xnum && xnum[1] !== 'General' && xnum[1] !== '\\@') st.numFmt = xnum[1].replace(/\\/g, '');
          if (Object.keys(st).length) cell.style = st;
          const gv = td.getAttribute('data-sheets-value'); if (gv) { try { const j = JSON.parse(gv); if (j[3] != null) cell.raw = j[3]; else if (j[2] != null) cell.raw = j[2]; else if (j[4] != null) cell.raw = !!j[4]; } catch { /* ignore */ } }
          const gf = td.getAttribute('data-sheets-formula'); if (gf) cell.f = gf.replace(/R\[(-?\d+)\]C\[(-?\d+)\]/g, (m) => m); // keep as text; R1C1 unsupported -> will be stored as value below
          if (cell.f && /R\[?-?\d*\]?C/.test(cell.f)) cell.f = null;
        }
        row[ci] = cell;
        if (rs > 1 || cs > 1) spans.push({ r1: ri, c1: ci, r2: ri + rs - 1, c2: ci + cs - 1 });
        for (let a = 0; a < rs; a++) for (let b = 0; b < cs; b++) if (a || b) occupied.set(`${ri + a},${ci + b}`, true);
        ci += cs;
        for (let b = 1; b < cs; b++) row[ci - cs + b] = { text: '', raw: null, f: null, s: null };
      }
      cells.push(row);
    });
    const w = Math.max(...cells.map((r) => r.length));
    for (const r of cells) while (r.length < w) r.push({ text: '', raw: null, f: null, s: null });
    let merges = spans;
    if (ours) { try { merges = JSON.parse(table.getAttribute('data-merges') || '[]'); } catch { merges = []; } }
    return { cells, merges, internal: ours };
  }

  // Write a parsed grid into the sheet at the active cell (tiling over the selection when it is a multiple).
  applyGrid(grid, options) {
    const app = this.app; const sheet = app.sheet; const wb = app.wb;
    const mode = options.mode || 'all';
    const src = grid.cells; const h = src.length, w = src[0].length;
    let target = app.sel.ranges[0];
    const tw = target.c2 - target.c1 + 1, th = target.r2 - target.r1 + 1;
    const tiles = (tw % w === 0 && th % h === 0 && (tw > w || th > h)) ? { cols: tw / w, rows: th / h } : { cols: 1, rows: 1 };
    const origin = { r: target.r1, c: target.c1 };
    const transposed = mode === 'transpose';
    const outH = transposed ? w : h, outW = transposed ? h : w;
    const total = { r1: origin.r, c1: origin.c, r2: origin.r + outH * tiles.rows - 1, c2: origin.c + outW * tiles.cols - 1 };
    if (total.r2 >= MAX_ROWS || total.c2 >= MAX_COLS) { app.toast('Paste area exceeds the sheet'); return; }
    app.transact('Paste', () => {
      if (grid.cut && grid.internal && grid.srcSheet) {
        // move semantics: references to the moved block follow it
        const dr = origin.r - grid.srcRange.r1, dc = origin.c - grid.srcRange.c1;
        grid.srcSheet.moveRange(grid.srcRange, dr, dc, sheet);
        grid.srcSheet.unmerge(grid.srcRange);
        if (grid.merges.length) sheet.setMerges([...sheet.merges.filter((m) => !(m.r1 >= total.r1 && m.r2 <= total.r2 && m.c1 >= total.c1 && m.c2 <= total.c2)), ...grid.merges.map((m) => ({ r1: m.r1 + origin.r, c1: m.c1 + origin.c, r2: m.r2 + origin.r, c2: m.c2 + origin.c }))]);
        return;
      }
      for (let tr = 0; tr < tiles.rows; tr++) for (let tc = 0; tc < tiles.cols; tc++) {
        for (let i = 0; i < h; i++) for (let j = 0; j < w; j++) {
          const srcCell = src[i][j];
          const oi = transposed ? j : i, oj = transposed ? i : j;
          const r = origin.r + tr * outH + oi, c = origin.c + tc * outW + oj;
          const existing = sheet.getCell(r, c);
          if (options.skipBlanks && srcCell.text === '' && srcCell.raw == null && !srcCell.f) continue;
          // determine style
          let sid = existing ? existing.s : sheet.getStyleId(r, c);
          const srcStyle = srcCell.s != null && grid.internal && this.internal ? wb.styles[srcCell.s] : srcCell.style || (srcCell.s != null ? wb.styles[srcCell.s] : null);
          if (mode === 'all' || mode === 'formats' || mode === 'transpose') { if (srcStyle) sid = wb.styleId(srcStyle); else if (grid.internal) sid = 0; }
          if (mode === 'formats') { sheet._setCellRaw(r, c, { ...(existing || { v: null, f: null }), s: sid }); continue; }
          // determine content
          let v = null, f = null;
          if (srcCell.f && (mode === 'all' || mode === 'formulas' || mode === 'transpose')) {
            const srcR = grid.srcRange ? grid.srcRange.r1 + i : r, srcC = grid.srcRange ? grid.srcRange.c1 + j : c;
            f = grid.internal ? shiftFormula(srcCell.f, r - srcR, c - srcC) : srcCell.f;
          } else if (srcCell.f && mode === 'values') { v = srcCell.v !== undefined ? srcCell.v : srcCell.raw; if (v && v.error) v = null; }
          else if (srcCell.raw !== undefined && srcCell.raw !== null) v = srcCell.raw;
          else if (srcCell.raw === null && (grid.internal || srcCell.text === '')) v = null;
          else { const p = parseInput(srcCell.text); v = p.value; if (p.format && !(srcStyle && srcStyle.numFmt) && (mode === 'all' || mode === 'transpose')) { const base = wb.styles[sid] || {}; if (!base.numFmt || base.numFmt === 'General') sid = wb.styleId(mergeStyle(base, { numFmt: p.format })); } }
          if (mode === 'values' && existing) sid = existing.s;
          if (typeof v === 'string' && v.startsWith('=') && !f) f = v, v = null;
          const cell = { v, f, s: sid, note: mode === 'all' ? (srcCell.note || null) : existing && existing.note, hyperlink: mode === 'all' ? (srcCell.hyperlink || null) : existing && existing.hyperlink };
          sheet._setCellRaw(r, c, cell);
        }
        if ((mode === 'all') && grid.merges && grid.merges.length && !transposed) {
          const add = grid.merges.map((m) => ({ r1: m.r1 + origin.r + tr * outH, c1: m.c1 + origin.c + tc * outW, r2: m.r2 + origin.r + tr * outH, c2: m.c2 + origin.c + tc * outW }));
          sheet.setMerges([...sheet.merges.filter((m) => !add.some((a) => !(m.r2 < a.r1 || a.r2 < m.r1 || m.c2 < a.c1 || a.c2 < m.c1))), ...add]);
        }
      }
    });
    app.selectRange(total);
  }
}

function cssToHex(s) {
  s = s.trim().toLowerCase();
  if (/^#[0-9a-f]{6}$/.test(s)) return s;
  if (/^#[0-9a-f]{3}$/.test(s)) return '#' + s.slice(1).split('').map((ch) => ch + ch).join('');
  const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(s);
  if (m) return '#' + [m[1], m[2], m[3]].map((v) => (+v).toString(16).padStart(2, '0')).join('');
  const names = { black: '#000000', white: '#ffffff', red: '#ff0000', blue: '#0000ff', green: '#008000', yellow: '#ffff00', gray: '#808080', grey: '#808080', orange: '#ffa500', transparent: null, inherit: null, windowtext: '#000000' };
  return names[s] === undefined ? null : names[s];
}
