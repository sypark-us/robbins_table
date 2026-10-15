// Shared helpers: A1 addressing, string utilities, misc.

export const MAX_ROWS = 1048576;
export const MAX_COLS = 16384;

export function colToName(c) {
  let s = '';
  c += 1;
  while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); }
  return s;
}

export function nameToCol(name) {
  let c = 0;
  for (let i = 0; i < name.length; i++) c = c * 26 + (name.charCodeAt(i) - 64);
  return c - 1;
}

export function cellName(r, c) { return colToName(c) + (r + 1); }

const A1_RE = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/;
export function parseCellName(s) {
  const m = A1_RE.exec(s.trim());
  if (!m) return null;
  const c = nameToCol(m[2].toUpperCase()), r = parseInt(m[4], 10) - 1;
  if (c < 0 || c >= MAX_COLS || r < 0 || r >= MAX_ROWS) return null;
  return { r, c, absC: m[1] === '$', absR: m[3] === '$' };
}

// Parse "A1", "A1:B5", "A:A", "1:1", "Sheet1!A1:B2". Returns {sheet, r1,c1,r2,c2} or null.
export function parseRangeName(s) {
  s = s.trim();
  let sheet = null;
  const bang = s.lastIndexOf('!');
  if (bang >= 0) {
    sheet = s.slice(0, bang);
    if (sheet.startsWith("'") && sheet.endsWith("'")) sheet = sheet.slice(1, -1).replace(/''/g, "'");
    s = s.slice(bang + 1);
  }
  const parts = s.split(':');
  if (parts.length === 1) {
    const a = parseCellName(parts[0]);
    if (!a) return null;
    return { sheet, r1: a.r, c1: a.c, r2: a.r, c2: a.c };
  }
  if (parts.length !== 2) return null;
  const a = parseCellName(parts[0]), b = parseCellName(parts[1]);
  if (a && b) return { sheet, r1: Math.min(a.r, b.r), c1: Math.min(a.c, b.c), r2: Math.max(a.r, b.r), c2: Math.max(a.c, b.c) };
  if (/^\$?[A-Za-z]{1,3}$/.test(parts[0]) && /^\$?[A-Za-z]{1,3}$/.test(parts[1])) {
    const c1 = nameToCol(parts[0].replace('$', '').toUpperCase()), c2 = nameToCol(parts[1].replace('$', '').toUpperCase());
    return { sheet, r1: 0, c1: Math.min(c1, c2), r2: MAX_ROWS - 1, c2: Math.max(c1, c2) };
  }
  if (/^\$?\d+$/.test(parts[0]) && /^\$?\d+$/.test(parts[1])) {
    const r1 = parseInt(parts[0].replace('$', ''), 10) - 1, r2 = parseInt(parts[1].replace('$', ''), 10) - 1;
    return { sheet, r1: Math.min(r1, r2), c1: 0, r2: Math.max(r1, r2), c2: MAX_COLS - 1 };
  }
  return null;
}

export function rangeName(rg, withSheet = false) {
  let s;
  if (rg.r1 === 0 && rg.r2 >= MAX_ROWS - 1) s = `${colToName(rg.c1)}:${colToName(rg.c2)}`;
  else if (rg.c1 === 0 && rg.c2 >= MAX_COLS - 1) s = `${rg.r1 + 1}:${rg.r2 + 1}`;
  else if (rg.r1 === rg.r2 && rg.c1 === rg.c2) s = cellName(rg.r1, rg.c1);
  else s = `${cellName(rg.r1, rg.c1)}:${cellName(rg.r2, rg.c2)}`;
  if (withSheet && rg.sheet) s = `${quoteSheetName(rg.sheet)}!${s}`;
  return s;
}

export function quoteSheetName(name) {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`;
}

export function normRange(r1, c1, r2, c2) {
  return { r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) };
}

export function rangesIntersect(a, b) {
  return !(a.r2 < b.r1 || b.r2 < a.r1 || a.c2 < b.c1 || b.c2 < a.c1);
}

export function rangeContains(rg, r, c) {
  return r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2;
}

export function key(r, c) { return r + ',' + c; }
export function unkey(k) { const i = k.indexOf(','); return [+k.slice(0, i), +k.slice(i + 1)]; }

export const clamp = (v, lo, hi) => Math.max(lo, Math.min(hi, v));

export function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (ch) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
}

export function debounce(fn, ms) {
  let t = null;
  return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); };
}

let uidCounter = 0;
export function uid(prefix = 'id') { return `${prefix}${Date.now().toString(36)}${(uidCounter++).toString(36)}`; }

export function deepClone(o) { return o == null ? o : JSON.parse(JSON.stringify(o)); }

export function isNumericString(s) { return /^\s*[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?\s*$/.test(s); }

// Color helpers
export function hexToRgb(hex) {
  hex = hex.replace('#', '');
  if (hex.length === 3) hex = hex.split('').map((ch) => ch + ch).join('');
  if (hex.length === 8) hex = hex.slice(2);
  const n = parseInt(hex, 16);
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
}
export function rgbToHex(r, g, b) {
  return '#' + [r, g, b].map((v) => clamp(Math.round(v), 0, 255).toString(16).padStart(2, '0')).join('');
}
export function mixColor(a, b, t) {
  const A = hexToRgb(a), B = hexToRgb(b);
  return rgbToHex(A[0] + (B[0] - A[0]) * t, A[1] + (B[1] - A[1]) * t, A[2] + (B[2] - A[2]) * t);
}
export function isDarkColor(hex) {
  const [r, g, b] = hexToRgb(hex);
  return (r * 299 + g * 587 + b * 114) / 1000 < 140;
}

export function compareValues(a, b) {
  // Excel sort order: numbers < text < booleans < errors < blanks
  const rank = (v) => v == null || v === '' ? 4 : typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : typeof v === 'boolean' ? 2 : 3;
  const ra = rank(a), rb = rank(b);
  if (ra !== rb) return ra - rb;
  if (ra === 0) return a - b;
  if (ra === 1) return a.localeCompare(b, undefined, { sensitivity: 'accent', numeric: false });
  if (ra === 2) return (a ? 1 : 0) - (b ? 1 : 0);
  return 0;
}
