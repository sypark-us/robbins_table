// XLSX (Office Open XML SpreadsheetML) import and export.
import { zipRead, zipWrite, utf8 } from './zip.js';
import { parseXml, child, children, attr, textOf, x, escapeXml, XML_DECL, local } from './xml.js';
import { Workbook, Sheet, DEFAULT_COL_WIDTH, DEFAULT_ROW_HEIGHT } from './model.js';
import { BUILTIN_FORMATS, BUILTIN_FORMAT_IDS, isDateFormat } from './format.js';
import { shiftFormula, FormulaError } from './formula.js';
import { colToName, parseCellName, parseRangeName, rangeName, key as cellKey, unkey, uid, MAX_ROWS, MAX_COLS } from './util.js';

const pxToChars = (px) => Math.round(Math.max(0, (px - 5) / 7) * 100) / 100;
const charsToPx = (ch) => Math.round(ch * 7 + 5);
const pxToPt = (px) => Math.round(px * 0.75 * 100) / 100;
const ptToPx = (pt) => Math.round(pt / 0.75);
const toArgb = (hex) => 'FF' + hex.replace('#', '').toUpperCase().padStart(6, '0');
const fromArgb = (s) => s ? '#' + s.slice(-6).toLowerCase() : null;

const INDEXED_COLORS = ['000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080', '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF', '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF', '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99', '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696', '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333'];
const THEME_COLORS = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72'];

function colorOf(el) {
  if (!el) return null;
  if (attr(el, 'rgb')) return fromArgb(attr(el, 'rgb'));
  if (attr(el, 'indexed') != null) { const i = +attr(el, 'indexed'); return INDEXED_COLORS[i] ? '#' + INDEXED_COLORS[i].toLowerCase() : null; }
  if (attr(el, 'theme') != null) { const t = +attr(el, 'theme'); const base = THEME_COLORS[t]; if (!base) return null; const tint = parseFloat(attr(el, 'tint') || '0'); return applyTint('#' + base.toLowerCase(), tint); }
  if (attr(el, 'auto')) return null;
  return null;
}
function applyTint(hex, tint) {
  if (!tint) return hex;
  const n = parseInt(hex.slice(1), 16); let r = (n >> 16) & 255, g = (n >> 8) & 255, b = n & 255;
  const f = (c) => tint < 0 ? c * (1 + tint) : c + (255 - c) * tint;
  r = Math.round(f(r)); g = Math.round(f(g)); b = Math.round(f(b));
  return '#' + [r, g, b].map((v) => Math.max(0, Math.min(255, v)).toString(16).padStart(2, '0')).join('');
}

// ---------------- Export ----------------
export async function exportXlsx(wb) {
  const styles = new StyleTable(wb);
  const files = [];
  const sheetXmls = wb.sheets.map((s, i) => sheetXml(s, styles, i));
  files.push({ name: '[Content_Types].xml', data: XML_DECL + x('Types', { xmlns: 'http://schemas.openxmlformats.org/package/2006/content-types' }, [
    x('Default', { Extension: 'rels', ContentType: 'application/vnd.openxmlformats-package.relationships+xml' }),
    x('Default', { Extension: 'xml', ContentType: 'application/xml' }),
    x('Override', { PartName: '/xl/workbook.xml', ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml' }),
    ...wb.sheets.map((_, i) => x('Override', { PartName: `/xl/worksheets/sheet${i + 1}.xml`, ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml' })),
    x('Override', { PartName: '/xl/styles.xml', ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml' }),
    x('Override', { PartName: '/docProps/core.xml', ContentType: 'application/vnd.openxmlformats-package.core-properties+xml' }),
    x('Override', { PartName: '/docProps/app.xml', ContentType: 'application/vnd.openxmlformats-officedocument.extended-properties+xml' }),
  ]) });
  files.push({ name: '_rels/.rels', data: XML_DECL + x('Relationships', { xmlns: 'http://schemas.openxmlformats.org/package/2006/relationships' }, [
    x('Relationship', { Id: 'rId1', Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument', Target: 'xl/workbook.xml' }),
    x('Relationship', { Id: 'rId2', Type: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', Target: 'docProps/core.xml' }),
    x('Relationship', { Id: 'rId3', Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties', Target: 'docProps/app.xml' }),
  ]) });
  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  files.push({ name: 'docProps/core.xml', data: XML_DECL + `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${escapeXml(wb.name)}</dc:title><dc:creator>OfficeTable</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>` });
  files.push({ name: 'docProps/app.xml', data: XML_DECL + `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>OfficeTable</Application></Properties>` });
  const definedNames = [...wb.names.values()].map((n) => x('definedName', { name: n.name, localSheetId: n.sheetId ? wb.sheets.findIndex((s) => s.id === n.sheetId) : undefined }, escapeXml(n.ref.replace(/^=/, ''))));
  files.push({ name: 'xl/workbook.xml', data: XML_DECL + x('workbook', { xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main', 'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships' }, [
    x('bookViews', {}, [x('workbookView', { activeTab: wb.activeSheetIndex })]),
    x('sheets', {}, wb.sheets.map((s, i) => x('sheet', { name: s.name, sheetId: i + 1, 'r:id': `rId${i + 1}` }))),
    definedNames.length ? x('definedNames', {}, definedNames) : '',
    x('calcPr', { calcId: '191029', fullCalcOnLoad: '1' }),
  ]) });
  files.push({ name: 'xl/_rels/workbook.xml.rels', data: XML_DECL + x('Relationships', { xmlns: 'http://schemas.openxmlformats.org/package/2006/relationships' }, [
    ...wb.sheets.map((_, i) => x('Relationship', { Id: `rId${i + 1}`, Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet', Target: `worksheets/sheet${i + 1}.xml` })),
    x('Relationship', { Id: `rId${wb.sheets.length + 1}`, Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles', Target: 'styles.xml' }),
  ]) });
  sheetXmls.forEach((xml, i) => {
    files.push({ name: `xl/worksheets/sheet${i + 1}.xml`, data: xml });
    const rels = wb.sheets[i]._exportRels || [];
    if (rels.length) files.push({ name: `xl/worksheets/_rels/sheet${i + 1}.xml.rels`, data: XML_DECL + x('Relationships', { xmlns: 'http://schemas.openxmlformats.org/package/2006/relationships' }, rels.map((r) => x('Relationship', { Id: r.id, Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', Target: r.target, TargetMode: 'External' }))) });
    delete wb.sheets[i]._exportRels;
  });
  files.push({ name: 'xl/styles.xml', data: styles.toXml() });
  return zipWrite(files);
}

class StyleTable {
  constructor(wb) {
    this.wb = wb;
    this.numFmts = new Map(); this.nextFmt = 164;
    this.fonts = [JSON.stringify({})]; this.fontIndex = new Map([[this.fonts[0], 0]]);
    this.fills = ['none', 'gray125']; this.fillIndex = new Map();
    this.borders = [JSON.stringify({})]; this.borderIndex = new Map([[this.borders[0], 0]]);
    this.xfs = []; this.xfIndex = new Map();
    this.dxfs = [];
    this.map = new Map(); // our style id -> xf index
    this.xf({}); // default
  }
  fmtId(code) {
    if (!code || code === 'General') return 0;
    if (BUILTIN_FORMAT_IDS[code] != null) return BUILTIN_FORMAT_IDS[code];
    if (!this.numFmts.has(code)) this.numFmts.set(code, this.nextFmt++);
    return this.numFmts.get(code);
  }
  fontId(st) {
    const f = { name: st.font, sz: st.size, b: st.bold, i: st.italic, u: st.underline, strike: st.strike, color: st.color };
    const k = JSON.stringify(f);
    if (!this.fontIndex.has(k)) { this.fontIndex.set(k, this.fonts.length); this.fonts.push(k); }
    return this.fontIndex.get(k);
  }
  fillId(st) {
    if (!st.fill) return 0;
    const k = st.fill;
    if (!this.fillIndex.has(k)) { this.fillIndex.set(k, this.fills.length); this.fills.push(k); }
    return this.fillIndex.get(k);
  }
  borderId(st) {
    const k = JSON.stringify(st.border || {});
    if (!this.borderIndex.has(k)) { this.borderIndex.set(k, this.borders.length); this.borders.push(k); }
    return this.borderIndex.get(k);
  }
  xf(st) {
    const xf = { numFmtId: this.fmtId(st.numFmt), fontId: this.fontId(st), fillId: this.fillId(st), borderId: this.borderId(st), align: st.align, valign: st.valign, wrap: st.wrap, indent: st.indent, rotation: st.rotation };
    const k = JSON.stringify(xf);
    if (!this.xfIndex.has(k)) { this.xfIndex.set(k, this.xfs.length); this.xfs.push(xf); }
    return this.xfIndex.get(k);
  }
  idFor(styleId) {
    if (!styleId) return 0;
    if (!this.map.has(styleId)) this.map.set(styleId, this.xf(this.wb.styles[styleId] || {}));
    return this.map.get(styleId);
  }
  dxf(style) { this.dxfs.push(style); return this.dxfs.length - 1; }
  toXml() {
    const fontXml = (k) => { const f = JSON.parse(k); return x('font', {}, [f.b ? '<b/>' : '', f.i ? '<i/>' : '', f.u ? '<u/>' : '', f.strike ? '<strike/>' : '', x('sz', { val: f.sz || 11 }), f.color ? x('color', { rgb: toArgb(f.color) }) : '', x('name', { val: f.name || 'Calibri' })]); };
    const fillXml = (k) => k === 'none' || k === 'gray125' ? x('fill', {}, [x('patternFill', { patternType: k })]) : x('fill', {}, [x('patternFill', { patternType: 'solid' }, [x('fgColor', { rgb: toArgb(k) }), x('bgColor', { indexed: 64 })])]);
    const borderXml = (k) => { const b = JSON.parse(k); const side = (n) => b[n] ? x(n, { style: b[n].style || 'thin' }, [x('color', { rgb: toArgb(b[n].color || '#000000') })]) : x(n); return x('border', {}, [side('left'), side('right'), side('top'), side('bottom'), x('diagonal')]); };
    const xfXml = (xf) => { const al = (xf.align || xf.valign || xf.wrap || xf.indent || xf.rotation) ? x('alignment', { horizontal: xf.align && xf.align !== 'general' ? xf.align : undefined, vertical: xf.valign === 'middle' ? 'center' : xf.valign, wrapText: xf.wrap ? 1 : undefined, indent: xf.indent || undefined, textRotation: xf.rotation || undefined }) : ''; return x('xf', { numFmtId: xf.numFmtId, fontId: xf.fontId, fillId: xf.fillId, borderId: xf.borderId, xfId: 0, applyNumberFormat: xf.numFmtId ? 1 : undefined, applyFont: xf.fontId ? 1 : undefined, applyFill: xf.fillId ? 1 : undefined, applyBorder: xf.borderId ? 1 : undefined, applyAlignment: al ? 1 : undefined }, al ? [al] : null); };
    const dxfXml = (d) => x('dxf', {}, [(d.bold || d.italic || d.color) ? x('font', {}, [d.bold ? '<b/>' : '', d.italic ? '<i/>' : '', d.color ? x('color', { rgb: toArgb(d.color) }) : '']) : '', d.fill ? x('fill', {}, [x('patternFill', {}, [x('bgColor', { rgb: toArgb(d.fill) })])]) : '']);
    return XML_DECL + x('styleSheet', { xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main' }, [
      this.numFmts.size ? x('numFmts', { count: this.numFmts.size }, [...this.numFmts].map(([code, id]) => x('numFmt', { numFmtId: id, formatCode: code }))) : '',
      x('fonts', { count: this.fonts.length }, this.fonts.map(fontXml)),
      x('fills', { count: this.fills.length }, this.fills.map(fillXml)),
      x('borders', { count: this.borders.length }, this.borders.map(borderXml)),
      x('cellStyleXfs', { count: 1 }, [x('xf', { numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 })]),
      x('cellXfs', { count: this.xfs.length }, this.xfs.map(xfXml)),
      x('cellStyles', { count: 1 }, [x('cellStyle', { name: 'Normal', xfId: 0, builtinId: 0 })]),
      x('dxfs', { count: this.dxfs.length }, this.dxfs.map(dxfXml)),
    ]);
  }
}

const cellRef = (r, c) => colToName(c) + (r + 1);
const rangeRef = (rg) => rg.r1 === rg.r2 && rg.c1 === rg.c2 ? cellRef(rg.r1, rg.c1) : `${cellRef(rg.r1, rg.c1)}:${cellRef(rg.r2, rg.c2)}`;
const CF_OPS = { gt: 'greaterThan', lt: 'lessThan', ge: 'greaterThanOrEqual', le: 'lessThanOrEqual', eq: 'equal', ne: 'notEqual', between: 'between', notBetween: 'notBetween' };
const CF_OPS_INV = Object.fromEntries(Object.entries(CF_OPS).map(([k, v]) => [v, k]));

function sheetXml(sheet, styles, index) {
  const rows = new Map();
  for (const [k, cell] of sheet.cells) {
    const [r, c] = unkey(k);
    if (r >= MAX_ROWS || c >= MAX_COLS) continue;
    if (!rows.has(r)) rows.set(r, []);
    rows.get(r).push([c, cell]);
  }
  for (const [r] of sheet.rowAxis.sizes) if (!rows.has(r)) rows.set(r, []);
  for (const r of sheet.rowAxis.hidden) if (!rows.has(r)) rows.set(r, []);
  for (const [r] of sheet.rowStyles) if (!rows.has(r)) rows.set(r, []);
  const sortedRows = [...rows.keys()].sort((a, b) => a - b);
  const rowXml = sortedRows.map((r) => {
    const cells = rows.get(r).sort((a, b) => a[0] - b[0]).map(([c, cell]) => {
      const s = styles.idFor(cell.s || sheet.rowStyles.get(r) || sheet.colStyles.get(c) || 0);
      const attrs = { r: cellRef(r, c), s: s || undefined };
      if (cell.f) {
        const v = cell.v;
        const f = x('f', {}, escapeXml(cell.f.slice(1)));
        if (v == null) return x('c', attrs, [f]);
        if (typeof v === 'string') return x('c', { ...attrs, t: 'str' }, [f, x('v', {}, escapeXml(v))]);
        if (typeof v === 'boolean') return x('c', { ...attrs, t: 'b' }, [f, x('v', {}, v ? '1' : '0')]);
        if (v instanceof FormulaError || (v && v.error)) return x('c', { ...attrs, t: 'e' }, [f, x('v', {}, v.error)]);
        return x('c', attrs, [f, x('v', {}, numStr(v))]);
      }
      if (cell.spillOf) return s ? x('c', attrs) : '';
      const v = cell.v;
      if (v == null) return x('c', attrs);
      if (typeof v === 'string') return x('c', { ...attrs, t: 'inlineStr' }, [x('is', {}, [x('t', { 'xml:space': 'preserve' }, escapeXml(v))])]);
      if (typeof v === 'boolean') return x('c', { ...attrs, t: 'b' }, [x('v', {}, v ? '1' : '0')]);
      if (v && v.error) return x('c', { ...attrs, t: 'e' }, [x('v', {}, v.error)]);
      return x('c', attrs, [x('v', {}, numStr(v))]);
    }).filter(Boolean);
    const h = sheet.rowAxis.sizes.get(r);
    const rs = sheet.rowStyles.get(r);
    return x('row', { r: r + 1, ht: h != null ? pxToPt(h) : undefined, customHeight: h != null ? 1 : undefined, hidden: sheet.rowAxis.hidden.has(r) ? 1 : undefined, s: rs ? styles.idFor(rs) : undefined, customFormat: rs ? 1 : undefined }, cells);
  });
  const colEntries = new Set([...sheet.colAxis.sizes.keys(), ...sheet.colAxis.hidden, ...sheet.colStyles.keys()]);
  const cols = [...colEntries].sort((a, b) => a - b).map((c) => x('col', { min: c + 1, max: c + 1, width: pxToChars(sheet.colAxis.rawSize(c)), customWidth: 1, hidden: sheet.colAxis.hidden.has(c) ? 1 : undefined, style: sheet.colStyles.get(c) ? styles.idFor(sheet.colStyles.get(c)) : undefined }));
  const fz = sheet.freeze;
  const pane = (fz.rows || fz.cols) ? x('pane', { xSplit: fz.cols || undefined, ySplit: fz.rows || undefined, topLeftCell: cellRef(fz.rows, fz.cols), activePane: fz.rows && fz.cols ? 'bottomRight' : fz.rows ? 'bottomLeft' : 'topRight', state: 'frozen' }) : '';
  const active = sheet.view && sheet.view.active ? cellRef(sheet.view.active.r, sheet.view.active.c) : 'A1';
  const sheetView = x('sheetView', { tabSelected: index === sheet.wb.activeSheetIndex ? 1 : undefined, workbookViewId: 0, showGridLines: sheet.showGridlines ? undefined : 0, zoomScale: sheet.view && sheet.view.zoom && sheet.view.zoom !== 1 ? Math.round(sheet.view.zoom * 100) : undefined }, [pane, x('selection', { pane: pane ? (fz.rows && fz.cols ? 'bottomRight' : fz.rows ? 'bottomLeft' : 'topRight') : undefined, activeCell: active, sqref: active })]);
  const cfXml = sheet.cf.map((rule, i) => {
    const sq = rangeRef(rule);
    const dxfId = rule.style ? styles.dxf(rule.style) : undefined;
    const common = { priority: i + 1, dxfId };
    const fml = (v) => x('formula', {}, escapeXml(String(v ?? '').replace(/^=/, '')));
    let inner;
    switch (rule.type) {
      case 'cellIs': inner = x('cfRule', { type: 'cellIs', operator: CF_OPS[rule.op] || 'equal', ...common }, [fml(rule.v1), rule.v2 != null && rule.v2 !== '' ? fml(rule.v2) : '']); break;
      case 'containsText': inner = x('cfRule', { type: 'containsText', operator: 'containsText', text: rule.v1, ...common }, [fml(`NOT(ISERROR(SEARCH("${rule.v1}",${cellRef(rule.r1, rule.c1)})))`)]); break;
      case 'notContainsText': inner = x('cfRule', { type: 'notContainsText', operator: 'notContains', text: rule.v1, ...common }, [fml(`ISERROR(SEARCH("${rule.v1}",${cellRef(rule.r1, rule.c1)}))`)]); break;
      case 'beginsWith': inner = x('cfRule', { type: 'beginsWith', operator: 'beginsWith', text: rule.v1, ...common }, [fml(`LEFT(${cellRef(rule.r1, rule.c1)},LEN("${rule.v1}"))="${rule.v1}"`)]); break;
      case 'endsWith': inner = x('cfRule', { type: 'endsWith', operator: 'endsWith', text: rule.v1, ...common }, [fml(`RIGHT(${cellRef(rule.r1, rule.c1)},LEN("${rule.v1}"))="${rule.v1}"`)]); break;
      case 'duplicate': inner = x('cfRule', { type: 'duplicateValues', ...common }); break;
      case 'unique': inner = x('cfRule', { type: 'uniqueValues', ...common }); break;
      case 'top': inner = x('cfRule', { type: 'top10', rank: rule.rank || 10, percent: rule.percent ? 1 : undefined, ...common }); break;
      case 'bottom': inner = x('cfRule', { type: 'top10', bottom: 1, rank: rule.rank || 10, percent: rule.percent ? 1 : undefined, ...common }); break;
      case 'aboveAverage': inner = x('cfRule', { type: 'aboveAverage', ...common }); break;
      case 'belowAverage': inner = x('cfRule', { type: 'aboveAverage', aboveAverage: 0, ...common }); break;
      case 'blank': inner = x('cfRule', { type: 'containsBlanks', ...common }, [fml(`LEN(TRIM(${cellRef(rule.r1, rule.c1)}))=0`)]); break;
      case 'notBlank': inner = x('cfRule', { type: 'notContainsBlanks', ...common }, [fml(`LEN(TRIM(${cellRef(rule.r1, rule.c1)}))>0`)]); break;
      case 'formula': inner = x('cfRule', { type: 'expression', ...common }, [fml(rule.v1)]); break;
      case 'colorScale': { const cs = rule.colors || ['#f8696b', '#ffeb84', '#63be7b']; const n = cs.length; const cfvo = n === 3 ? [x('cfvo', { type: 'min' }), x('cfvo', { type: 'percentile', val: 50 }), x('cfvo', { type: 'max' })] : [x('cfvo', { type: 'min' }), x('cfvo', { type: 'max' })]; inner = x('cfRule', { type: 'colorScale', priority: i + 1 }, [x('colorScale', {}, [...cfvo, ...cs.map((c) => x('color', { rgb: toArgb(c) }))])]); break; }
      case 'dataBar': inner = x('cfRule', { type: 'dataBar', priority: i + 1 }, [x('dataBar', {}, [x('cfvo', { type: 'min' }), x('cfvo', { type: 'max' }), x('color', { rgb: toArgb(rule.barColor || '#638ec6') })])]); break;
      default: return '';
    }
    return x('conditionalFormatting', { sqref: sq }, [inner]);
  });
  const dvXml = sheet.validations.length ? x('dataValidations', { count: sheet.validations.length }, sheet.validations.map((v) => {
    const type = v.type === 'list' ? 'list' : v.type;
    const ops = { between: 'between', notBetween: 'notBetween', eq: 'equal', ne: 'notEqual', gt: 'greaterThan', lt: 'lessThan', ge: 'greaterThanOrEqual', le: 'lessThanOrEqual' };
    let f1 = v.v1 ?? '', f2 = v.v2;
    if (v.type === 'list' && f1 && !f1.startsWith('=')) f1 = '"' + f1.replace(/"/g, '') + '"';
    return x('dataValidation', { type, operator: v.type !== 'list' && v.type !== 'custom' ? (ops[v.op] || 'between') : undefined, allowBlank: v.allowBlank === false ? undefined : 1, showDropDown: v.type === 'list' && v.showDropdown === false ? 1 : undefined, showInputMessage: v.prompt ? 1 : undefined, showErrorMessage: 1, errorStyle: v.errorStyle === 'warning' ? 'warning' : v.errorStyle === 'info' ? 'information' : undefined, errorTitle: v.errorTitle || undefined, error: v.error || undefined, prompt: v.prompt || undefined, sqref: rangeRef(v) }, [f1 !== '' ? x('formula1', {}, escapeXml(String(f1).replace(/^=/, ''))) : '', f2 != null && f2 !== '' ? x('formula2', {}, escapeXml(String(f2).replace(/^=/, ''))) : '']);
  })) : '';
  const links = []; const rels = [];
  for (const [k, cell] of sheet.cells) if (cell.hyperlink) {
    const [r, c] = unkey(k);
    if (cell.hyperlink.startsWith('#')) links.push(x('hyperlink', { ref: cellRef(r, c), location: cell.hyperlink.slice(1), display: cell.hyperlink.slice(1) }));
    else { const id = `rId${rels.length + 1}`; rels.push({ id, target: cell.hyperlink }); links.push(x('hyperlink', { ref: cellRef(r, c), 'r:id': id, tooltip: cell.hyperlink })); }
  }
  sheet._exportRels = rels;
  return XML_DECL + x('worksheet', { xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main', 'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships' }, [
    sheet.tabColor ? x('sheetPr', {}, [x('tabColor', { rgb: toArgb(sheet.tabColor) })]) : '',
    x('sheetViews', {}, [sheetView]),
    x('sheetFormatPr', { defaultRowHeight: pxToPt(DEFAULT_ROW_HEIGHT), defaultColWidth: pxToChars(DEFAULT_COL_WIDTH) }),
    cols.length ? x('cols', {}, cols) : '',
    x('sheetData', {}, rowXml),
    sheet.filter ? x('autoFilter', { ref: rangeRef(sheet.filter) }) : '',
    sheet.merges.length ? x('mergeCells', { count: sheet.merges.length }, sheet.merges.map((m) => x('mergeCell', { ref: rangeRef(m) }))) : '',
    cfXml.join(''),
    dvXml,
    links.length ? x('hyperlinks', {}, links) : '',
    x('pageMargins', { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 }),
  ]);
}
function numStr(v) { if (Number.isInteger(v)) return String(v); const s = String(v); return s.includes('e') ? v.toPrecision(17).replace(/0+$/, '') : s; }

// ---------------- Import ----------------
export async function importXlsx(bytes, fileName = 'Book1') {
  const parts = await zipRead(bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes));
  const read = (name) => { const b = parts.get(name) || parts.get(name.replace(/^\//, '')); return b ? utf8.decode(b) : null; };
  const wbXmlText = read('xl/workbook.xml');
  if (!wbXmlText) throw new Error('Not a valid .xlsx file (xl/workbook.xml missing)');
  const wbXml = parseXml(wbXmlText);
  const relsXml = parseXml(read('xl/_rels/workbook.xml.rels') || '<Relationships/>');
  const rels = {};
  for (const r of children(relsXml, 'Relationship')) rels[attr(r, 'Id')] = attr(r, 'Target');
  const date1904 = attr(child(wbXml, 'workbookPr'), 'date1904') === '1' || attr(child(wbXml, 'workbookPr'), 'date1904') === 'true';
  // shared strings
  const sst = [];
  const sstText = read('xl/sharedStrings.xml');
  if (sstText) for (const si of children(parseXml(sstText), 'si')) sst.push(richText(si));
  // styles
  const st = parseStyles(read('xl/styles.xml'));
  const wb = new Workbook();
  wb.name = fileName.replace(/\.xlsx$/i, '');
  wb.styles = [{}]; wb.styleIndex = new Map([['{}', 0]]);
  const styleIds = st.xfs.map((xf) => wb.styleId(xf));
  const sheetEls = children(child(wbXml, 'sheets'), 'sheet');
  for (const sEl of sheetEls) {
    const rid = attr(sEl, 'id') || attr(sEl, 'r:id');
    let target = rels[rid] || '';
    if (target.startsWith('/')) target = target.slice(1); else if (!target.startsWith('xl/')) target = 'xl/' + target;
    const xmlText = read(target);
    const sheet = new Sheet(wb, attr(sEl, 'name'));
    wb.sheets.push(sheet);
    if (attr(sEl, 'state') === 'hidden') sheet.hidden = true;
    if (!xmlText) continue;
    const relText = read(target.replace(/worksheets\/(sheet[^/]+)\.xml$/, 'worksheets/_rels/$1.xml.rels'));
    const sheetRels = {};
    if (relText) for (const r of children(parseXml(relText), 'Relationship')) sheetRels[attr(r, 'Id')] = attr(r, 'Target');
    parseSheet(parseXml(xmlText), sheet, { sst, st, styleIds, date1904, sheetRels, wb });
  }
  const dn = child(wbXml, 'definedNames');
  if (dn) for (const n of children(dn, 'definedName')) {
    const name = attr(n, 'name'); const ref = textOf(n);
    if (!name || name.startsWith('_xlnm.')) { if (name === '_xlnm._FilterDatabase') continue; if (name && name.startsWith('_xlnm.')) continue; }
    const ls = attr(n, 'localSheetId');
    const sheet = ls != null ? wb.sheets[+ls] : null;
    wb.names.set((sheet ? `${sheet.id}!` : '') + name.toUpperCase(), { name, ref: '=' + ref, sheetId: sheet ? sheet.id : null });
  }
  const bv = child(child(wbXml, 'bookViews'), 'workbookView');
  wb.activeSheetIndex = Math.min(+(attr(bv, 'activeTab') || 0), wb.sheets.length - 1);
  if (!wb.sheets.length) wb._addSheetRaw('Sheet1');
  wb.engine.rebuildAll();
  for (const s of wb.sheets) s.applyFilter();
  return wb;
}

function richText(si) {
  const t = child(si, 't');
  if (t) return textOf(t);
  return children(si, 'r').map((r) => textOf(child(r, 't'))).join('');
}

function parseStyles(text) {
  const out = { numFmts: { ...BUILTIN_FORMATS }, fonts: [], fills: [], borders: [], xfs: [], dxfs: [] };
  if (!text) { out.xfs.push({}); return out; }
  const xml = parseXml(text);
  for (const nf of children(child(xml, 'numFmts'), 'numFmt')) out.numFmts[+attr(nf, 'numFmtId')] = attr(nf, 'formatCode');
  for (const f of children(child(xml, 'fonts'), 'font')) {
    const font = {};
    const name = attr(child(f, 'name'), 'val'); if (name) font.font = name;
    const sz = attr(child(f, 'sz'), 'val'); if (sz) font.size = parseFloat(sz);
    if (child(f, 'b') && attr(child(f, 'b'), 'val') !== '0') font.bold = true;
    if (child(f, 'i') && attr(child(f, 'i'), 'val') !== '0') font.italic = true;
    if (child(f, 'u') && attr(child(f, 'u'), 'val') !== 'none') font.underline = true;
    if (child(f, 'strike') && attr(child(f, 'strike'), 'val') !== '0') font.strike = true;
    const color = colorOf(child(f, 'color')); if (color && color !== '#000000') font.color = color;
    out.fonts.push(font);
  }
  for (const fl of children(child(xml, 'fills'), 'fill')) {
    const pf = child(fl, 'patternFill');
    const type = attr(pf, 'patternType');
    let fill = null;
    if (pf && type && type !== 'none') { fill = colorOf(child(pf, 'fgColor')) || colorOf(child(pf, 'bgColor')); if (type === 'gray125' && !colorOf(child(pf, 'fgColor'))) fill = null; }
    out.fills.push(fill);
  }
  for (const b of children(child(xml, 'borders'), 'border')) {
    const border = {};
    for (const side of ['left', 'right', 'top', 'bottom']) { const s = child(b, side); const style = attr(s, 'style'); if (style && style !== 'none') border[side] = { style, color: colorOf(child(s, 'color')) || '#000000' }; }
    out.borders.push(border);
  }
  for (const xf of children(child(xml, 'cellXfs'), 'xf')) {
    const s = {};
    const nf = +(attr(xf, 'numFmtId') || 0); if (nf && out.numFmts[nf] && out.numFmts[nf] !== 'General') s.numFmt = out.numFmts[nf];
    const font = out.fonts[+(attr(xf, 'fontId') || 0)]; if (font) Object.assign(s, font);
    const fill = out.fills[+(attr(xf, 'fillId') || 0)]; if (fill) s.fill = fill;
    const border = out.borders[+(attr(xf, 'borderId') || 0)]; if (border && Object.keys(border).length) s.border = border;
    const al = child(xf, 'alignment');
    if (al) { const h = attr(al, 'horizontal'); if (h && h !== 'general') s.align = h === 'centerContinuous' ? 'center' : h; const v = attr(al, 'vertical'); if (v) s.valign = v === 'center' ? 'middle' : v; if (attr(al, 'wrapText') === '1' || attr(al, 'wrapText') === 'true') s.wrap = true; const ind = attr(al, 'indent'); if (ind) s.indent = +ind; const rot = attr(al, 'textRotation'); if (rot) s.rotation = +rot; }
    if (s.font === 'Calibri') delete s.font; if (s.size === 11) delete s.size;
    out.xfs.push(s);
  }
  if (!out.xfs.length) out.xfs.push({});
  for (const d of children(child(xml, 'dxfs'), 'dxf')) {
    const s = {};
    const f = child(d, 'font'); if (f) { if (child(f, 'b')) s.bold = true; if (child(f, 'i')) s.italic = true; const c = colorOf(child(f, 'color')); if (c) s.color = c; }
    const fl = child(child(d, 'fill'), 'patternFill'); if (fl) { const c = colorOf(child(fl, 'bgColor')) || colorOf(child(fl, 'fgColor')); if (c) s.fill = c; }
    out.dxfs.push(s);
  }
  return out;
}

function parseSheet(xml, sheet, ctx) {
  const { sst, st, styleIds, date1904, sheetRels } = ctx;
  const pr = child(xml, 'sheetPr');
  const tab = child(pr, 'tabColor'); if (tab) sheet.tabColor = colorOf(tab);
  const view = child(child(xml, 'sheetViews'), 'sheetView');
  if (view) {
    if (attr(view, 'showGridLines') === '0' || attr(view, 'showGridLines') === 'false') sheet.showGridlines = false;
    const zoom = attr(view, 'zoomScale'); if (zoom) sheet.view.zoom = Math.max(0.25, Math.min(4, +zoom / 100));
    const pane = child(view, 'pane');
    if (pane && attr(pane, 'state') === 'frozen') sheet.freeze = { rows: +(attr(pane, 'ySplit') || 0), cols: +(attr(pane, 'xSplit') || 0) };
    const sel = children(view, 'selection').pop(); const ac = attr(sel, 'activeCell'); const p = ac && parseCellName(ac); if (p) { sheet.view.active = { r: p.r, c: p.c }; sheet.view.ranges = [{ r1: p.r, c1: p.c, r2: p.r, c2: p.c }]; }
  }
  const fmt = child(xml, 'sheetFormatPr');
  const defColW = attr(fmt, 'defaultColWidth') || attr(fmt, 'baseColWidth');
  if (defColW) sheet.colAxis.default = charsToPx(+defColW + (attr(fmt, 'defaultColWidth') ? 0 : 0.71));
  const defRowH = attr(fmt, 'defaultRowHeight'); if (defRowH && attr(fmt, 'customHeight') === '1') sheet.rowAxis.default = ptToPx(+defRowH);
  for (const col of children(child(xml, 'cols'), 'col')) {
    const min = +attr(col, 'min') - 1, max = Math.min(+attr(col, 'max') - 1, min + 2000);
    const w = attr(col, 'width'); const hidden = attr(col, 'hidden') === '1' || attr(col, 'hidden') === 'true';
    const sIdx = attr(col, 'style');
    for (let c = min; c <= max; c++) {
      if (w != null && attr(col, 'customWidth') !== '0') sheet.colAxis.sizes.set(c, charsToPx(parseFloat(w)));
      if (hidden) sheet.colAxis.hidden.add(c);
      if (sIdx && styleIds[+sIdx]) sheet.colStyles.set(c, styleIds[+sIdx]);
    }
  }
  sheet.colAxis.invalidate();
  const shared = new Map();
  const dateShift = date1904 ? 1462 : 0;
  for (const row of children(child(xml, 'sheetData'), 'row')) {
    const r = +attr(row, 'r') - 1;
    const ht = attr(row, 'ht'); if (ht && (attr(row, 'customHeight') === '1' || attr(row, 'customHeight') === 'true')) sheet.rowAxis.sizes.set(r, ptToPx(parseFloat(ht)));
    if (attr(row, 'hidden') === '1' || attr(row, 'hidden') === 'true') sheet.rowAxis.hidden.add(r);
    if (attr(row, 'customFormat') === '1' && attr(row, 's')) { const sid = styleIds[+attr(row, 's')]; if (sid) sheet.rowStyles.set(r, sid); }
    let colCursor = 0;
    for (const c of children(row, 'c')) {
      const ref = attr(c, 'r');
      const pos = ref ? parseCellName(ref) : { r, c: colCursor };
      if (!pos) continue;
      colCursor = pos.c + 1;
      const t = attr(c, 't') || 'n';
      const sIdx = attr(c, 's');
      const cell = { v: null, f: null, s: sIdx != null ? (styleIds[+sIdx] || 0) : 0 };
      const fEl = child(c, 'f');
      const vEl = child(c, 'v');
      const vText = vEl ? textOf(vEl) : null;
      if (fEl) {
        let ftext = textOf(fEl);
        const ft = attr(fEl, 't');
        if (ft === 'shared') {
          const si = attr(fEl, 'si');
          if (ftext) shared.set(si, { f: ftext, r: pos.r, c: pos.c });
          else { const m = shared.get(si); if (m) ftext = shiftFormula('=' + m.f, pos.r - m.r, pos.c - m.c).slice(1); }
        }
        if (ftext) cell.f = '=' + ftext.replace(/_xlfn\./g, '').replace(/_xlws\./g, '');
      }
      let value = null;
      if (t === 's') value = sst[+vText] ?? '';
      else if (t === 'inlineStr') value = richText(child(c, 'is'));
      else if (t === 'str') value = vText ?? '';
      else if (t === 'b') value = vText === '1' || vText === 'true';
      else if (t === 'e') value = { error: vText || '#VALUE!' };
      else if (vText != null && vText !== '') { value = Number(vText); if (isNaN(value)) value = vText; else if (dateShift && cell.s && isDateFormat(ctx.wb.styles[cell.s].numFmt)) value += dateShift; }
      if (!cell.f) cell.v = value && value.error ? null : value;
      if (cell.v != null || cell.f || cell.s) sheet.cells.set(cellKey(pos.r, pos.c), cell);
    }
  }
  sheet.rowAxis.invalidate();
  for (const m of children(child(xml, 'mergeCells'), 'mergeCell')) { const rg = parseRangeName(attr(m, 'ref')); if (rg) sheet.merges.push({ r1: rg.r1, c1: rg.c1, r2: rg.r2, c2: rg.c2 }); }
  const af = child(xml, 'autoFilter');
  if (af) { const rg = parseRangeName(attr(af, 'ref')); if (rg) { const conds = {}; for (const fc of children(af, 'filterColumn')) { const colId = +attr(fc, 'colId'); const filters = child(fc, 'filters'); if (filters) conds[rg.c1 + colId] = { values: children(filters, 'filter').map((f) => attr(f, 'val')) }; } sheet.filter = { r1: rg.r1, c1: rg.c1, r2: rg.r2, c2: rg.c2, conds }; } }
  for (const cf of children(xml, 'conditionalFormatting')) {
    const sq = (attr(cf, 'sqref') || '').split(/\s+/)[0]; const rg = parseRangeName(sq); if (!rg) continue;
    for (const rule of children(cf, 'cfRule')) {
      const type = attr(rule, 'type'); const dxf = attr(rule, 'dxfId'); const style = dxf != null ? st.dxfs[+dxf] : null;
      const fs = children(rule, 'formula').map((f) => '=' + textOf(f));
      const base = { id: uid('cf'), r1: rg.r1, c1: rg.c1, r2: rg.r2, c2: rg.c2, style: style || { fill: '#ffc7ce', color: '#9c0006' } };
      let out = null;
      switch (type) {
        case 'cellIs': out = { ...base, type: 'cellIs', op: CF_OPS_INV[attr(rule, 'operator')] || 'eq', v1: fs[0], v2: fs[1] }; break;
        case 'containsText': out = { ...base, type: 'containsText', v1: attr(rule, 'text') || '' }; break;
        case 'notContainsText': out = { ...base, type: 'notContainsText', v1: attr(rule, 'text') || '' }; break;
        case 'beginsWith': out = { ...base, type: 'beginsWith', v1: attr(rule, 'text') || '' }; break;
        case 'endsWith': out = { ...base, type: 'endsWith', v1: attr(rule, 'text') || '' }; break;
        case 'duplicateValues': out = { ...base, type: 'duplicate' }; break;
        case 'uniqueValues': out = { ...base, type: 'unique' }; break;
        case 'top10': out = { ...base, type: attr(rule, 'bottom') === '1' ? 'bottom' : 'top', rank: +(attr(rule, 'rank') || 10), percent: attr(rule, 'percent') === '1' }; break;
        case 'aboveAverage': out = { ...base, type: attr(rule, 'aboveAverage') === '0' ? 'belowAverage' : 'aboveAverage' }; break;
        case 'containsBlanks': out = { ...base, type: 'blank' }; break;
        case 'notContainsBlanks': out = { ...base, type: 'notBlank' }; break;
        case 'expression': out = { ...base, type: 'formula', v1: fs[0] }; break;
        case 'colorScale': { const cs = child(rule, 'colorScale'); const colors = children(cs, 'color').map(colorOf).filter(Boolean); out = { ...base, type: 'colorScale', colors: colors.length >= 2 ? colors : ['#f8696b', '#63be7b'], style: null }; break; }
        case 'dataBar': { const db = child(rule, 'dataBar'); out = { ...base, type: 'dataBar', barColor: colorOf(child(db, 'color')) || '#638ec6', style: null }; break; }
        default: break;
      }
      if (out) sheet.cf.push(out);
    }
  }
  const dvs = child(xml, 'dataValidations');
  const OPS_INV = { between: 'between', notBetween: 'notBetween', equal: 'eq', notEqual: 'ne', greaterThan: 'gt', lessThan: 'lt', greaterThanOrEqual: 'ge', lessThanOrEqual: 'le' };
  if (dvs) for (const dv of children(dvs, 'dataValidation')) {
    for (const sq of (attr(dv, 'sqref') || '').split(/\s+/)) {
      const rg = parseRangeName(sq); if (!rg) continue;
      const type = attr(dv, 'type') || 'any'; if (type === 'any') continue;
      let v1 = textOf(child(dv, 'formula1')); const v2 = textOf(child(dv, 'formula2'));
      if (type === 'list') { if (/^".*"$/.test(v1)) v1 = v1.slice(1, -1); else v1 = '=' + v1; } else { if (v1 && isNaN(Number(v1))) v1 = '=' + v1; }
      sheet.validations.push({ id: uid('dv'), r1: rg.r1, c1: rg.c1, r2: rg.r2, c2: rg.c2, type, op: OPS_INV[attr(dv, 'operator')] || 'between', v1, v2: v2 && isNaN(Number(v2)) ? '=' + v2 : v2, allowBlank: attr(dv, 'allowBlank') !== '0', showDropdown: attr(dv, 'showDropDown') !== '1', errorStyle: attr(dv, 'errorStyle') === 'warning' ? 'warning' : attr(dv, 'errorStyle') === 'information' ? 'info' : 'stop', errorTitle: attr(dv, 'errorTitle') || '', error: attr(dv, 'error') || '', prompt: attr(dv, 'prompt') || '' });
    }
  }
  for (const hl of children(child(xml, 'hyperlinks'), 'hyperlink')) {
    const p = parseCellName(attr(hl, 'ref') || ''); if (!p) continue;
    const rid = attr(hl, 'id'); const loc = attr(hl, 'location');
    const url = rid && sheetRels[rid] ? sheetRels[rid] : loc ? '#' + loc : null;
    if (!url) continue;
    const k = cellKey(p.r, p.c); const cell = sheet.cells.get(k) || { v: null, f: null, s: 0 }; cell.hyperlink = url; sheet.cells.set(k, cell);
  }
}
