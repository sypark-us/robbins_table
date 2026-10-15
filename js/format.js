// Number formatting: Excel-style format codes, date serial numbers, input parsing.

const DAY_MS = 86400000;
const EPOCH_1900 = Date.UTC(1899, 11, 30); // serial 0 (for serials >= 61)
const EPOCH_1900_EARLY = Date.UTC(1899, 11, 31); // Excel's fictitious 1900-02-29 shifts early dates by one

export const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
export const WEEKDAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

export function dateToSerial(y, m, d, h = 0, mi = 0, s = 0, ms = 0) {
  const t = Date.UTC(y, m - 1, d);
  const epoch = t < Date.UTC(1900, 2, 1) ? EPOCH_1900_EARLY : EPOCH_1900;
  return (t - epoch) / DAY_MS + (h * 3600000 + mi * 60000 + s * 1000 + ms) / DAY_MS;
}

export function jsDateToSerial(date, local = true) {
  if (local) return dateToSerial(date.getFullYear(), date.getMonth() + 1, date.getDate(), date.getHours(), date.getMinutes(), date.getSeconds(), date.getMilliseconds());
  return dateToSerial(date.getUTCFullYear(), date.getUTCMonth() + 1, date.getUTCDate(), date.getUTCHours(), date.getUTCMinutes(), date.getUTCSeconds(), date.getUTCMilliseconds());
}

export function serialToDate(serial) {
  const days = Math.floor(serial);
  const frac = serial - days;
  const epoch = days < 61 ? EPOCH_1900_EARLY : EPOCH_1900;
  const dt = new Date(epoch + days * DAY_MS);
  let totalMs = Math.round(frac * DAY_MS);
  if (totalMs >= DAY_MS) totalMs = DAY_MS - 1;
  const h = Math.floor(totalMs / 3600000), mi = Math.floor((totalMs % 3600000) / 60000), s = Math.floor((totalMs % 60000) / 1000), ms = totalMs % 1000;
  return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate(), h, mi, s, ms, weekday: dt.getUTCDay(), serial };
}

export function todaySerial() { const n = new Date(); return dateToSerial(n.getFullYear(), n.getMonth() + 1, n.getDate()); }
export function nowSerial() { return jsDateToSerial(new Date()); }

// ---------- Built-in formats ----------
export const BUILTIN_FORMATS = {
  0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??',
  14: 'm/d/yyyy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss', 22: 'm/d/yyyy h:mm',
  37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)', 39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)',
  45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mmss.0', 48: '##0.0E+0', 49: '@',
};
export const BUILTIN_FORMAT_IDS = Object.fromEntries(Object.entries(BUILTIN_FORMATS).map(([k, v]) => [v, +k]));

export const FORMAT_PRESETS = [
  { name: 'General', code: 'General' },
  { name: 'Number', code: '0.00' },
  { name: 'Number (thousands)', code: '#,##0.00' },
  { name: 'Integer', code: '0' },
  { name: 'Currency', code: '$#,##0.00;[Red]-$#,##0.00' },
  { name: 'Accounting', code: '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)' },
  { name: 'Percent', code: '0.00%' },
  { name: 'Percent (integer)', code: '0%' },
  { name: 'Scientific', code: '0.00E+00' },
  { name: 'Fraction', code: '# ?/?' },
  { name: 'Short date', code: 'm/d/yyyy' },
  { name: 'ISO date', code: 'yyyy-mm-dd' },
  { name: 'Long date', code: 'dddd, mmmm d, yyyy' },
  { name: 'Time', code: 'h:mm:ss AM/PM' },
  { name: 'Time (24h)', code: 'h:mm' },
  { name: 'Date & time', code: 'm/d/yyyy h:mm' },
  { name: 'Duration', code: '[h]:mm:ss' },
  { name: 'Text', code: '@' },
];

const COLOR_NAMES = { black: '#000000', blue: '#0000ff', cyan: '#00ffff', green: '#00a000', magenta: '#ff00ff', red: '#ff0000', white: '#ffffff', yellow: '#ffd700' };

// ---------- Format code parsing ----------
const parseCache = new Map();

export function parseFormat(code) {
  if (!code || code === 'General') code = 'General';
  let p = parseCache.get(code);
  if (p) return p;
  p = { code, sections: splitSections(code).map(parseSection) };
  parseCache.set(code, p);
  return p;
}

function splitSections(code) {
  const out = []; let cur = '', q = false, br = false;
  for (let i = 0; i < code.length; i++) {
    const ch = code[i];
    if (ch === '"' && !br) q = !q;
    else if (ch === '[' && !q) br = true;
    else if (ch === ']' && !q) br = false;
    else if (ch === '\\' && !q && !br) { cur += ch + (code[i + 1] || ''); i++; continue; }
    if (ch === ';' && !q && !br) { out.push(cur); cur = ''; } else cur += ch;
  }
  out.push(cur);
  return out.slice(0, 4);
}

function parseSection(src) {
  const sec = { tokens: [], color: null, cond: null, type: 'number', intPlaces: [], decPlaces: 0, hasDec: false, pct: false, grouping: false, scale: 0, exp: null, frac: null };
  let i = 0;
  const tokens = [];
  while (i < src.length) {
    const ch = src[i];
    if (ch === '[') {
      const j = src.indexOf(']', i);
      const inner = src.slice(i + 1, j < 0 ? src.length : j);
      i = j < 0 ? src.length : j + 1;
      const low = inner.toLowerCase();
      if (COLOR_NAMES[low]) sec.color = COLOR_NAMES[low];
      else if (/^color\s*\d+$/i.test(inner)) sec.color = indexedColor(parseInt(inner.replace(/\D/g, ''), 10));
      else if (/^[<>=!]+/.test(inner)) { const m = /^(<=|>=|<>|<|>|=)\s*(-?[\d.]+(?:e[-+]?\d+)?)$/i.exec(inner); if (m) sec.cond = { op: m[1], v: parseFloat(m[2]) }; }
      else if (/^(h+|m+|s+)$/i.test(inner)) tokens.push({ t: 'elapsed', ch: low[0], len: inner.length });
      else if (/^\$/.test(inner)) tokens.push({ t: 'lit', s: inner.split('-')[0].slice(1) });
      continue;
    }
    if (ch === '"') { const j = src.indexOf('"', i + 1); tokens.push({ t: 'lit', s: src.slice(i + 1, j < 0 ? src.length : j) }); i = j < 0 ? src.length : j + 1; continue; }
    if (ch === '\\') { tokens.push({ t: 'lit', s: src[i + 1] || '' }); i += 2; continue; }
    if (ch === '_') { tokens.push({ t: 'lit', s: ' ' }); i += 2; continue; }
    if (ch === '*') { i += 2; continue; }
    if (ch === '0' || ch === '#' || ch === '?') { tokens.push({ t: 'digit', ch }); i++; continue; }
    if (ch === '.') { tokens.push({ t: 'dec' }); i++; continue; }
    if (ch === ',') { tokens.push({ t: 'comma' }); i++; continue; }
    if (ch === '%') { tokens.push({ t: 'pct' }); i++; continue; }
    if (ch === '@') { tokens.push({ t: 'text' }); i++; continue; }
    if (ch === '/') { tokens.push({ t: 'slash' }); i++; continue; }
    if ((ch === 'E' || ch === 'e') && (src[i + 1] === '+' || src[i + 1] === '-')) { tokens.push({ t: 'exp', sign: src[i + 1] }); i += 2; continue; }
    if (/^general/i.test(src.slice(i))) { tokens.push({ t: 'general' }); i += 7; continue; }
    if (/^(am\/pm|a\/p)/i.test(src.slice(i))) { const m = /^(am\/pm|a\/p)/i.exec(src.slice(i)); tokens.push({ t: 'ampm', s: m[0] }); i += m[0].length; continue; }
    if (/[ymdhs]/i.test(ch)) {
      const m = /^([ymdhs])\1*/i.exec(src.slice(i));
      tokens.push({ t: 'date', ch: m[0][0].toLowerCase(), len: m[0].length }); i += m[0].length; continue;
    }
    tokens.push({ t: 'lit', s: ch }); i++;
  }
  const hasDate = tokens.some((t) => t.t === 'date' || t.t === 'elapsed' || t.t === 'ampm');
  const hasDigits = tokens.some((t) => t.t === 'digit');
  if (tokens.some((t) => t.t === 'general') && !hasDigits && !hasDate) sec.type = 'general';
  else if (hasDate) { sec.type = 'date'; resolveMonthMinutes(tokens); }
  else if (tokens.some((t) => t.t === 'text') && !hasDigits) sec.type = 'text';
  else if (!hasDigits) sec.type = 'literal';
  if (sec.type === 'number') analyzeNumber(sec, tokens);
  sec.tokens = tokens;
  return sec;
}

function resolveMonthMinutes(tokens) {
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i];
    if (t.t !== 'date' || t.ch !== 'm') continue;
    let prev = null, next = null;
    for (let j = i - 1; j >= 0; j--) if (tokens[j].t === 'date' || tokens[j].t === 'elapsed') { prev = tokens[j]; break; }
    for (let j = i + 1; j < tokens.length; j++) if (tokens[j].t === 'date' || tokens[j].t === 'elapsed') { next = tokens[j]; break; }
    if ((prev && (prev.ch === 'h' || prev.ch === 's') && t.len <= 2) || (next && next.ch === 's' && t.len <= 2)) t.minutes = true;
  }
  for (let i = 0; i < tokens.length; i++) {
    if (tokens[i].t === 'dec') {
      let n = 0; let j = i + 1;
      while (j < tokens.length && tokens[j].t === 'digit' && tokens[j].ch === '0') { n++; j++; }
      if (n) { tokens[i] = { t: 'secfrac', n }; tokens.splice(i + 1, n); }
    }
  }
}

function analyzeNumber(sec, tokens) {
  let afterDec = false, seenDigit = false, afterExp = false;
  let slashAt = -1;
  for (let i = 0; i < tokens.length; i++) if (tokens[i].t === 'slash') { slashAt = i; break; }
  if (slashAt >= 0) {
    const num = [], den = [], int = [];
    let j = slashAt - 1;
    while (j >= 0 && tokens[j].t === 'digit') { num.unshift(j); j--; }
    for (let k = 0; k <= j; k++) if (tokens[k].t === 'digit') int.push(k);
    let fixedDen = '';
    for (let k = slashAt + 1; k < tokens.length; k++) {
      if (tokens[k].t === 'digit') den.push(k);
      else if (tokens[k].t === 'lit' && /^\d$/.test(tokens[k].s) && den.length === 0) fixedDen += tokens[k].s;
      else break;
    }
    sec.frac = { int, num, den, fixedDen: fixedDen ? parseInt(fixedDen, 10) : 0, slashAt };
    sec.intPlaces = int;
    return;
  }
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i];
    if (t.t === 'digit') {
      if (afterExp) sec.exp.places.push(i);
      else if (afterDec) sec.decPlaces++;
      else sec.intPlaces.push(i);
      seenDigit = true;
    } else if (t.t === 'dec' && !afterDec && !afterExp) { afterDec = true; sec.hasDec = true; t.decPoint = true; }
    else if (t.t === 'pct') sec.pct = true;
    else if (t.t === 'comma') {
      const nextDigit = tokens.slice(i + 1).find((x) => x.t === 'digit' || x.t === 'dec' || x.t === 'exp');
      if (!afterDec && seenDigit && nextDigit && nextDigit.t === 'digit') sec.grouping = true;
      else if (seenDigit && !nextDigit) { sec.scale++; t.scale = true; }
      else { t.t = 'lit'; t.s = ','; }
    } else if (t.t === 'exp') { afterExp = true; sec.exp = { sign: t.sign, places: [] }; }
  }
}

function indexedColor(n) {
  const idx = ['#000000', '#ffffff', '#ff0000', '#00ff00', '#0000ff', '#ffff00', '#ff00ff', '#00ffff'];
  return idx[(n - 1) % 8] || '#000000';
}

// ---------- Formatting ----------

export function formatValue(value, code) {
  if (value == null || value === '') return { text: '', color: null };
  if (typeof value === 'object' && value.error) return { text: value.error, color: null };
  const fmt = parseFormat(code);
  const secs = fmt.sections;
  if (typeof value === 'boolean') return { text: value ? 'TRUE' : 'FALSE', color: null };
  if (typeof value === 'string') {
    const ts = secs.length === 4 ? secs[3] : secs.find((s) => s.type === 'text');
    if (!ts) return { text: value, color: null };
    return { text: renderText(ts, value), color: ts.color };
  }
  if (typeof value !== 'number' || !isFinite(value)) return { text: String(value), color: null };
  let sec, showMinus = false;
  const hasConds = secs.some((s) => s.cond);
  if (hasConds) {
    sec = secs.find((s) => s.cond && testCond(s.cond, value)) || secs.find((s) => !s.cond && s.type !== 'text') || secs[0];
    showMinus = value < 0;
  } else if (secs.length === 1 || (secs.length === 2 && secs[1].type === 'text')) { sec = secs[0]; showMinus = value < 0; }
  else if (value > 0 || (value === 0 && secs.length < 3)) sec = secs[0];
  else if (value < 0) sec = secs[1];
  else sec = secs[2] || secs[0];
  if (sec.type === 'text') sec = secs[0];
  let text;
  const a = Math.abs(value);
  if (sec.type === 'general') text = (showMinus ? '-' : '') + formatGeneral(a);
  else if (sec.type === 'date') text = formatDate(sec, value);
  else if (sec.type === 'literal') text = sec.tokens.map((t) => t.t === 'lit' ? t.s : '').join('');
  else {
    const body = formatNumber(sec, a);
    text = showMinus && body.replace(/[^0-9]/g, '').replace(/^0+/, '') !== '' ? '-' + body : body;
  }
  return { text, color: sec.color };
}

function testCond(cond, v) {
  switch (cond.op) {
    case '<': return v < cond.v; case '<=': return v <= cond.v; case '>': return v > cond.v; case '>=': return v >= cond.v;
    case '=': return v === cond.v; case '<>': return v !== cond.v; default: return false;
  }
}

function renderText(sec, s) {
  return sec.tokens.map((t) => t.t === 'text' ? s : t.t === 'lit' ? t.s : '').join('');
}

export function formatGeneral(n, maxWidth = 11) {
  if (n === 0) return '0';
  if (!isFinite(n)) return '#NUM!';
  const neg = n < 0; const a = Math.abs(n);
  let s;
  if (Number.isInteger(a) && a < 1e15) s = String(a);
  else if (Number.isInteger(a)) s = sciNotation(a, 5);
  else {
    const exp = Math.floor(Math.log10(a));
    if (exp >= maxWidth - 1 || exp < -4) {
      s = sciNotation(a, Math.max(1, Math.min(6, maxWidth - 5)));
    } else {
      const sig = Math.max(1, Math.min(15, maxWidth - 1 - Math.max(0, -exp)));
      s = a.toPrecision(Math.min(15, Math.max(sig, exp + 1)));
      if (s.includes('e')) s = a.toFixed(Math.max(0, sig - exp - 1));
      if (s.includes('.')) s = s.replace(/0+$/, '').replace(/\.$/, '');
      if (s.length > maxWidth && s.includes('.')) { s = a.toFixed(Math.max(0, maxWidth - 1 - String(Math.floor(a)).length)); s = s.replace(/0+$/, '').replace(/\.$/, ''); }
      if (s === '0') s = sciNotation(a, 5);
    }
  }
  if (s.length > maxWidth && !s.includes('E')) { const alt = sciNotation(a, Math.max(1, maxWidth - 5)); if (alt.length < s.length) s = alt; }
  return (neg ? '-' : '') + s;
}

function sciNotation(a, digits) {
  const s = a.toExponential(digits);
  let [m, e] = s.split('e');
  if (m.includes('.')) m = m.replace(/0+$/, '').replace(/\.$/, '');
  const en = parseInt(e, 10);
  return `${m}E${en < 0 ? '-' : '+'}${String(Math.abs(en)).padStart(2, '0')}`;
}

function digitsOf(v, places) {
  let s = v.toFixed(Math.min(100, places));
  if (s.includes('e')) s = BigInt(Math.round(v)).toString() + (places ? '.' + '0'.repeat(places) : '');
  const [i, d = ''] = s.split('.');
  return [i, d];
}

export function formatNumber(sec, v) {
  const tokens = sec.tokens;
  if (sec.frac) return formatFraction(sec, v);
  let val = v;
  if (sec.pct) val *= 100;
  if (sec.scale) val /= Math.pow(1000, sec.scale);
  let intStr, decStr, expStr = '';
  if (sec.exp) {
    const intCount = sec.intPlaces.length || 1;
    let e = val === 0 ? 0 : Math.floor(Math.log10(Math.abs(val)));
    e -= (intCount - 1);
    if (intCount > 1) e = Math.floor(e / intCount) * intCount;
    let mant = val / Math.pow(10, e);
    const [i0] = digitsOf(mant, sec.decPlaces);
    if (i0.length > intCount) { e += i0.length - intCount; mant = val / Math.pow(10, e); }
    [intStr, decStr] = digitsOf(mant, sec.decPlaces);
    const expDigits = sec.exp.places.length || 1;
    expStr = (e < 0 ? '-' : sec.exp.sign === '+' ? '+' : '') + String(Math.abs(e)).padStart(expDigits, '0');
  } else {
    [intStr, decStr] = digitsOf(val, sec.decPlaces);
  }
  const n = sec.intPlaces.length;
  let intDigits = intStr;
  if (intDigits === '0' && (n === 0 || sec.intPlaces.every((i) => tokens[i].ch !== '0'))) intDigits = '';
  const minInt = sec.intPlaces.filter((i) => tokens[i].ch === '0').length;
  if (intDigits.length < minInt) intDigits = intDigits.padStart(minInt, '0');
  let grouped = sec.grouping ? groupThousands(intDigits) : intDigits;
  const padQ = sec.intPlaces.filter((i) => tokens[i].ch === '?').length;
  if (intDigits.length < minInt + padQ) grouped = ' '.repeat(minInt + padQ - intDigits.length) + grouped;
  const parts = new Array(n).fill('');
  if (n > 0) {
    let pos = grouped.length;
    for (let k = n - 1; k >= 1; k--) {
      if (pos <= 0) break;
      let start = pos - 1;
      while (start > 0 && !/[\d\s]/.test(grouped[start])) start--;
      parts[k] = grouped.slice(start, pos); pos = start;
    }
    parts[0] = grouped.slice(0, pos);
  }
  const out = [];
  let placeIdx = 0, di = 0, afterDec = false, afterExp = false;
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i];
    if (t.t === 'digit') {
      if (afterExp) continue;
      if (!afterDec) out.push(parts[placeIdx++] ?? '');
      else {
        const ch = decStr[di++];
        if (ch === undefined) out.push(t.ch === '0' ? '0' : t.ch === '?' ? ' ' : '');
        else if (t.ch !== '0' && /^0*$/.test(decStr.slice(di - 1))) out.push(t.ch === '?' ? ' ' : '');
        else out.push(ch);
      }
    } else if (t.t === 'dec' && t.decPoint) {
      afterDec = true;
      const shown = decStr.replace(/0+$/, '');
      let anyForced = false;
      for (let k = i + 1; k < tokens.length && tokens[k].t === 'digit'; k++) if (tokens[k].ch !== '#') anyForced = true;
      out.push(shown.length || anyForced ? '.' : '');
    }
    else if (t.t === 'exp') { afterExp = true; out.push('E' + expStr); }
    else if (t.t === 'comma') { if (!t.scale && !sec.grouping) out.push(','); }
    else if (t.t === 'pct') out.push('%');
    else if (t.t === 'lit') out.push(t.s);
    else if (t.t === 'general') out.push(formatGeneral(v));
  }
  return out.join('');
}

function groupThousands(s) {
  let out = '';
  for (let i = 0; i < s.length; i++) {
    const fromEnd = s.length - i;
    out += s[i];
    if (fromEnd > 1 && (fromEnd - 1) % 3 === 0 && /\d/.test(s[i])) out += ',';
  }
  return out;
}

function formatFraction(sec, v) {
  const tokens = sec.tokens; const f = sec.frac;
  const hasInt = f.int.length > 0;
  let whole = hasInt ? Math.floor(v) : 0;
  const rem = v - whole;
  let num, den;
  if (f.fixedDen) { den = f.fixedDen; num = Math.round(rem * den); }
  else { const maxDen = Math.pow(10, f.den.length) - 1; [num, den] = bestFraction(rem, maxDen); }
  if (num === den && den !== 0) { whole += 1; num = 0; }
  if (!hasInt) { num += whole * den; whole = 0; }
  const out = [];
  let intDone = false, numDone = false, denDone = false;
  const noFrac = num === 0;
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i];
    if (f.int.includes(i)) { if (!intDone) { out.push(whole === 0 && !noFrac ? (tokens[i].ch === '0' ? '0' : ' '.repeat(f.int.length)) : String(whole).padStart(f.int.length, tokens[i].ch === '0' ? '0' : ' ')); intDone = true; } }
    else if (f.num.includes(i)) { if (!numDone) { out.push(noFrac && hasInt ? ' '.repeat(f.num.length) : String(num).padStart(f.num.length, ' ')); numDone = true; } }
    else if (f.den.includes(i)) { if (!denDone) { out.push(noFrac && hasInt ? ' '.repeat(f.den.length) : String(den).padEnd(f.den.length, ' ')); denDone = true; } }
    else if (t.t === 'slash') out.push(noFrac && hasInt ? ' ' : '/');
    else if (t.t === 'lit') { if (f.fixedDen && /^\d$/.test(t.s) && i > f.slashAt && !denDone) { out.push(noFrac && hasInt ? ' ' : String(den)); denDone = true; } else out.push(t.s); }
  }
  return out.join('').replace(/\s+$/, '');
}

function bestFraction(x, maxDen) {
  if (x === 0) return [0, 1];
  let bestN = Math.round(x), bestD = 1, bestErr = Math.abs(x - Math.round(x));
  let h1 = 1, h0 = 0, k1 = 0, k0 = 1, b = x;
  for (let i = 0; i < 64; i++) {
    const a = Math.floor(b);
    const h2 = a * h1 + h0, k2 = a * k1 + k0;
    if (k2 > maxDen) break;
    h0 = h1; h1 = h2; k0 = k1; k1 = k2;
    const err = Math.abs(x - h1 / k1);
    if (err < bestErr) { bestErr = err; bestN = h1; bestD = k1; }
    if (err < 1e-12 || b - a === 0) break;
    b = 1 / (b - a);
  }
  return [bestN, bestD];
}

export function formatDate(sec, serial) {
  if (serial < 0 || serial >= 2958466) return '#######';
  const tokens = sec.tokens;
  const hasAmPm = tokens.some((t) => t.t === 'ampm');
  let fracDigits = 0;
  for (const t of tokens) if (t.t === 'secfrac') fracDigits = Math.max(fracDigits, t.n);
  const hasSeconds = tokens.some((t) => (t.t === 'date' && t.ch === 's') || (t.t === 'elapsed' && t.ch === 's'));
  const hasMinutes = tokens.some((t) => (t.t === 'date' && t.minutes) || (t.t === 'elapsed' && t.ch === 'm'));
  const hasHours = tokens.some((t) => (t.t === 'date' && t.ch === 'h') || (t.t === 'elapsed' && t.ch === 'h'));
  const unit = fracDigits ? Math.pow(10, -fracDigits) / 86400 : hasSeconds ? 1 / 86400 : hasMinutes ? 1 / 1440 : hasHours ? 1 / 24 : 1;
  let rounded = Math.round(serial / unit + 1e-7) * unit;
  if (!hasHours && !hasMinutes && !hasSeconds) rounded = Math.floor(serial + 1e-9);
  const d = serialToDate(rounded);
  const out = [];
  for (const t of tokens) {
    if (t.t === 'date') {
      if (t.ch === 'y') out.push(t.len <= 2 ? String(d.y % 100).padStart(2, '0') : String(d.y));
      else if (t.ch === 'm') {
        if (t.minutes) out.push(t.len >= 2 ? String(d.mi).padStart(2, '0') : String(d.mi));
        else if (t.len === 1) out.push(String(d.m));
        else if (t.len === 2) out.push(String(d.m).padStart(2, '0'));
        else if (t.len === 3) out.push(MONTHS[d.m - 1].slice(0, 3));
        else if (t.len === 4) out.push(MONTHS[d.m - 1]);
        else out.push(MONTHS[d.m - 1][0]);
      } else if (t.ch === 'd') {
        if (t.len === 1) out.push(String(d.d));
        else if (t.len === 2) out.push(String(d.d).padStart(2, '0'));
        else if (t.len === 3) out.push(WEEKDAYS[d.weekday].slice(0, 3));
        else out.push(WEEKDAYS[d.weekday]);
      } else if (t.ch === 'h') {
        let h = d.h;
        if (hasAmPm) { h %= 12; if (h === 0) h = 12; }
        out.push(t.len >= 2 ? String(h).padStart(2, '0') : String(h));
      } else if (t.ch === 's') out.push(t.len >= 2 ? String(d.s).padStart(2, '0') : String(d.s));
    } else if (t.t === 'elapsed') {
      const totalSec = Math.floor(rounded * 86400 + 1e-6);
      let val;
      if (t.ch === 'h') val = Math.floor(totalSec / 3600);
      else if (t.ch === 'm') val = Math.floor(totalSec / 60);
      else val = totalSec;
      out.push(String(val).padStart(t.len, '0'));
    } else if (t.t === 'ampm') {
      const am = d.h < 12;
      out.push(t.s.toUpperCase() === 'A/P' ? (am ? 'A' : 'P') : (am ? 'AM' : 'PM'));
    } else if (t.t === 'secfrac') {
      const ms = Math.round(((rounded * 86400) % 1) * Math.pow(10, t.n)) % Math.pow(10, t.n);
      out.push('.' + String(ms).padStart(t.n, '0'));
    } else if (t.t === 'lit') out.push(t.s);
    else if (t.t === 'slash') out.push('/');
    else if (t.t === 'dec') out.push('.');
    else if (t.t === 'comma') out.push(',');
    else if (t.t === 'general') out.push(formatGeneral(serial));
  }
  return out.join('');
}

export function formatType(code) {
  const f = parseFormat(code);
  return f.sections[0].type;
}

export function isDateFormat(code) { return !!code && formatType(code) === 'date'; }

export function formatWith(value, code) { return formatValue(value, code).text; }

// ---------- Input parsing ----------
const MONTH_RE = '(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)';
const monthIndex = (s) => MONTHS.findIndex((m) => m.toLowerCase().startsWith(s.slice(0, 3).toLowerCase())) + 1;

function parseTimePart(s) {
  const m = /^(\d{1,2})(?::(\d{1,2}))?(?::(\d{1,2})(?:\.(\d+))?)?\s*(am|pm|a|p)?$/i.exec(s.trim());
  if (!m) return null;
  if (!m[2] && !m[5]) return null;
  let h = +m[1]; const mi = +(m[2] || 0), sec = +(m[3] || 0), ms = m[4] ? Math.round(parseFloat('0.' + m[4]) * 1000) : 0;
  if (mi > 59 || sec > 59) return null;
  if (m[5]) { if (h > 12 || h === 0) return null; const pm = /^p/i.test(m[5]); if (h === 12) h = pm ? 12 : 0; else if (pm) h += 12; }
  else if (h > 23) return null;
  let code = m[2] ? (m[3] ? 'h:mm:ss' : 'h:mm') : 'h';
  if (m[5]) code += ' AM/PM';
  if (m[4]) code = 'h:mm:ss.000';
  return { h, mi, s: sec, ms, code };
}

function validYMD(y, m, d) {
  if (m < 1 || m > 12 || d < 1) return false;
  return d <= new Date(Date.UTC(y, m, 0)).getUTCDate();
}
const fixYear = (txt) => { let y = +txt; if (txt.length === 2) y += y < 30 ? 2000 : 1900; return y; };

export function parseDateInput(text) {
  const s = text.trim();
  let m;
  const now = new Date();
  if ((m = /^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?:[ T](.+))?$/.exec(s))) {
    const y = +m[1], mo = +m[2], d = +m[3];
    if (!validYMD(y, mo, d)) return null;
    let serial = dateToSerial(y, mo, d), code = 'yyyy-mm-dd';
    if (m[4]) { const t = parseTimePart(m[4]); if (!t) return null; serial += (t.h * 3600 + t.mi * 60 + t.s + t.ms / 1000) / 86400; code = 'yyyy-mm-dd h:mm' + (t.s || t.ms ? ':ss' : ''); }
    return { serial, code };
  }
  if ((m = /^(\d{1,2})[-/.](\d{1,2})(?:[-/.](\d{2}|\d{4}))?(?:\s+(.+))?$/.exec(s))) {
    const mo = +m[1], d = +m[2];
    const y = m[3] ? fixYear(m[3]) : now.getFullYear();
    if (!validYMD(y, mo, d)) return null;
    let serial = dateToSerial(y, mo, d);
    let code = m[3] ? 'm/d/yyyy' : 'd-mmm';
    if (m[4]) { const t = parseTimePart(m[4]); if (!t) return null; serial += (t.h * 3600 + t.mi * 60 + t.s + t.ms / 1000) / 86400; code = 'm/d/yyyy h:mm'; }
    return { serial, code };
  }
  let re = new RegExp(`^${MONTH_RE}[\\s-]+(\\d{4})$`, 'i');
  if ((m = re.exec(s))) return { serial: dateToSerial(+m[2], monthIndex(m[1]), 1), code: 'mmm-yy' };
  re = new RegExp(`^(\\d{1,2})[\\s-]+${MONTH_RE}[\\s,-]*(\\d{2}|\\d{4})?$`, 'i');
  if ((m = re.exec(s))) {
    const d = +m[1], mo = monthIndex(m[2]);
    const y = m[3] ? fixYear(m[3]) : now.getFullYear();
    if (!validYMD(y, mo, d)) return null;
    return { serial: dateToSerial(y, mo, d), code: m[3] ? 'd-mmm-yy' : 'd-mmm' };
  }
  re = new RegExp(`^${MONTH_RE}[\\s-]+(\\d{1,2})(?:st|nd|rd|th)?[\\s,-]*(\\d{2}|\\d{4})?$`, 'i');
  if ((m = re.exec(s))) {
    const mo = monthIndex(m[1]), d = +m[2];
    const y = m[3] ? fixYear(m[3]) : now.getFullYear();
    if (!validYMD(y, mo, d)) return null;
    return { serial: dateToSerial(y, mo, d), code: m[3] ? 'mmm d, yyyy' : 'd-mmm' };
  }
  re = new RegExp(`^${MONTH_RE}[\\s-]+(\\d{2}|\\d{4})$`, 'i');
  if ((m = re.exec(s))) return { serial: dateToSerial(fixYear(m[2]), monthIndex(m[1]), 1), code: 'mmm-yy' };
  const t = parseTimePart(s);
  if (t) return { serial: (t.h * 3600 + t.mi * 60 + t.s + t.ms / 1000) / 86400, code: t.code };
  return null;
}

// Convert typed text into a stored value plus (optionally) a number format to apply.
export function parseInput(text) {
  if (text == null) return { value: null };
  if (typeof text !== 'string') return { value: text };
  const s = text.trim();
  if (s === '') return { value: null };
  if (text.startsWith("'")) return { value: text.slice(1) };
  if (/^(true|false)$/i.test(s)) return { value: s.toLowerCase() === 'true' };
  const m = /^([-+])?(\$|€|£|¥)?\s*(\()?([-+])?(\d{1,3}(?:,\d{3})+|\d+)?(\.\d*)?(?:[eE]([-+]?\d+))?(\))?\s*(%?)$/.exec(s);
  if (m && (m[5] != null || (m[6] != null && m[6].length > 1)) && (!!m[3] === !!m[8])) {
    const neg = m[1] === '-' || m[4] === '-' || !!m[3];
    const intPart = (m[5] || '0').replace(/,/g, '');
    let v = parseFloat(intPart + (m[6] || '') + (m[7] != null ? 'e' + m[7] : ''));
    if (neg) v = -v;
    let format = null;
    const decimals = m[6] ? m[6].length - 1 : 0;
    if (m[9]) { v /= 100; format = decimals > 0 ? '0.' + '0'.repeat(Math.min(decimals, 2)) + '%' : '0%'; }
    else if (m[2]) format = m[2] + '#,##0.00';
    else if (m[5] && m[5].includes(',')) format = decimals ? '#,##0.00' : '#,##0';
    else if (m[7] != null) format = '0.00E+00';
    if (isFinite(v)) return { value: v, format };
  }
  let f;
  if ((f = /^(-?\d+)\s+(\d+)\/(\d+)$/.exec(s)) && +f[3] !== 0) {
    const v = +f[1] + (f[1].startsWith('-') ? -1 : 1) * (+f[2] / +f[3]);
    return { value: v, format: '# ?/?' };
  }
  const d = parseDateInput(s);
  if (d) return { value: d.serial, format: d.code };
  return { value: text };
}
