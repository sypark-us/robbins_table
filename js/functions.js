// Worksheet function library.
import { FormulaError, ERR, isErr } from './formula.js';
import { RangeRef, isRange, isArr } from './range.js';
import { formatGeneral, formatValue, parseInput, serialToDate, dateToSerial, todaySerial, nowSerial, parseDateInput, MONTHS, WEEKDAYS } from './format.js';
import { compareValues, colToName, parseRangeName, cellName, MAX_ROWS, MAX_COLS, quoteSheetName } from './util.js';

export const FUNCTIONS = new Map();
export const FUNCTION_LIST = [];

function def(names, sig, opts, fn) {
  if (typeof opts === 'function') { fn = opts; opts = {}; }
  for (const name of [].concat(names)) {
    const entry = { name, sig, fn, lift: !!opts.lift, lazy: !!opts.lazy, volatile: !!opts.volatile, errors: !!opts.errors, cat: opts.cat || 'Other', minArgs: opts.min ?? 0, maxArgs: opts.max ?? Infinity };
    FUNCTIONS.set(name, entry);
    FUNCTION_LIST.push(entry);
  }
}

// ---------------- Coercion helpers ----------------
export function num(v) {
  if (typeof v === 'number') return v;
  if (v == null) return 0;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (isErr(v)) throw v;
  if (typeof v === 'string') {
    if (v.trim() === '') throw ERR.VALUE('empty text');
    const p = parseInput(v);
    if (typeof p.value === 'number') return p.value;
    if (typeof p.value === 'boolean') return p.value ? 1 : 0;
    throw ERR.VALUE(`"${v}" is not a number`);
  }
  if (isArr(v)) return num(v[0][0]);
  if (isRange(v)) return num(v.value());
  throw ERR.VALUE();
}
export function str(v) {
  if (typeof v === 'string') return v;
  if (v == null) return '';
  if (typeof v === 'number') return formatGeneral(v);
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (isErr(v)) throw v;
  if (isArr(v)) return str(v[0][0]);
  if (isRange(v)) return str(v.value());
  return String(v);
}
export function bool(v) {
  if (typeof v === 'boolean') return v;
  if (v == null) return false;
  if (typeof v === 'number') return v !== 0;
  if (isErr(v)) throw v;
  if (typeof v === 'string') { const u = v.trim().toUpperCase(); if (u === 'TRUE') return true; if (u === 'FALSE') return false; const n = Number(v); if (v.trim() !== '' && !isNaN(n)) return n !== 0; throw ERR.VALUE(); }
  if (isArr(v)) return bool(v[0][0]);
  if (isRange(v)) return bool(v.value());
  throw ERR.VALUE();
}
const int = (v) => Math.trunc(num(v));
const scalar = (v) => isRange(v) ? (v.isCell ? v.value() : v.toArray()) : v;
const isEmpty = (v) => v == null || v === '';

export function toMatrix(v) {
  if (isRange(v)) return v.toArray();
  if (isArr(v)) return v;
  return [[v]];
}
const dims = (m) => [m.length, m[0] ? m[0].length : 0];
const flat = (m) => m.flat();

// iterate scalar values of all args: cb(value, isDirect)
function each(args, cb) {
  for (const a of args) {
    if (isRange(a)) a.forEach((v) => cb(v, false));
    else if (isArr(a)) { for (const row of a) for (const v of row) cb(v, false); }
    else cb(a, true);
  }
}
// numbers for aggregate functions (SUM semantics)
function nums(args, opts = {}) {
  const out = [];
  each(args, (v, direct) => {
    if (isErr(v)) throw v;
    if (typeof v === 'number') out.push(v);
    else if (typeof v === 'boolean') { if (direct || opts.boolInRange) out.push(v ? 1 : 0); }
    else if (typeof v === 'string') {
      if (direct) out.push(num(v));
      else if (opts.textInRange) out.push(0);
    } else if (v == null && opts.blankAsZero) out.push(0);
  });
  return out;
}
function numsOrErr(fn) { return (args, ctx) => fn(nums(args), args, ctx); }

// ---------------- Criteria (COUNTIF family) ----------------
function wildcardToRegex(s) {
  let re = '';
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (ch === '~' && i + 1 < s.length) { re += escapeRe(s[++i]); }
    else if (ch === '*') re += '.*';
    else if (ch === '?') re += '.';
    else re += escapeRe(ch);
  }
  return new RegExp('^' + re + '$', 'i');
}
const escapeRe = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

export function criteria(crit) {
  if (isRange(crit)) crit = crit.value();
  if (isArr(crit)) crit = crit[0][0];
  if (crit == null) return (v) => v == null || v === '' || v === 0;
  if (typeof crit === 'number') return (v) => (typeof v === 'number' && v === crit) || (typeof v === 'string' && Number(v) === crit && v.trim() !== '');
  if (typeof crit === 'boolean') return (v) => v === crit;
  if (isErr(crit)) return (v) => isErr(v) && v.error === crit.error;
  let s = String(crit);
  let op = '=';
  const m = /^(<=|>=|<>|<|>|=)/.exec(s);
  if (m) { op = m[1]; s = s.slice(op.length); }
  if (s === '') {
    if (op === '=') return (v) => v == null || v === '';
    if (op === '<>') return (v) => !(v == null || v === '');
  }
  const p = parseInput(s);
  const target = p.value;
  if (typeof target === 'number' || typeof target === 'boolean') {
    const t = typeof target === 'boolean' ? target : target;
    return (v) => {
      let x = v;
      if (typeof target === 'number' && typeof v === 'string' && v.trim() !== '' && !isNaN(Number(v))) x = Number(v);
      if (typeof x !== typeof target) return op === '<>';
      switch (op) {
        case '=': return x === t; case '<>': return x !== t; case '<': return x < t; case '<=': return x <= t; case '>': return x > t; case '>=': return x >= t;
      }
      return false;
    };
  }
  // text comparisons
  const lower = s.toLowerCase();
  if (op === '=' || op === '<>') {
    const re = wildcardToRegex(s);
    const test = (v) => typeof v === 'string' ? re.test(v) : (v != null && typeof v !== 'object' && re.test(str(v)) && false);
    return op === '=' ? test : (v) => !test(v);
  }
  return (v) => {
    if (typeof v !== 'string') return false;
    const c = v.toLowerCase().localeCompare(lower);
    switch (op) { case '<': return c < 0; case '<=': return c <= 0; case '>': return c > 0; case '>=': return c >= 0; }
    return false;
  };
}

function ifsPairs(args, start) {
  const pairs = [];
  for (let i = start; i + 1 < args.length; i += 2) {
    const rg = args[i];
    if (!isRange(rg) && !isArr(rg)) throw ERR.VALUE('criteria range expected');
    pairs.push({ m: toMatrix(rg), test: criteria(args[i + 1]) });
  }
  return pairs;
}
function ifsMask(pairs, rows, cols) {
  const mask = [];
  for (let i = 0; i < rows; i++) {
    const row = [];
    for (let j = 0; j < cols; j++) {
      let ok = true;
      for (const p of pairs) { const v = p.m[i] ? p.m[i][j] : undefined; if (!p.test(v === undefined ? null : v)) { ok = false; break; } }
      row.push(ok);
    }
    mask.push(row);
  }
  return mask;
}

// ---------------- Math ----------------
const M = 'Math';
def('SUM', 'number1, [number2], ...', { cat: M, min: 1 }, numsOrErr((n) => n.reduce((a, b) => a + b, 0)));
def('SUMSQ', 'number1, [number2], ...', { cat: M, min: 1 }, numsOrErr((n) => n.reduce((a, b) => a + b * b, 0)));
def('PRODUCT', 'number1, [number2], ...', { cat: M, min: 1 }, numsOrErr((n) => n.length ? n.reduce((a, b) => a * b, 1) : 0));
def('ABS', 'number', { cat: M, lift: true, min: 1, max: 1 }, ([x]) => Math.abs(num(x)));
def('SIGN', 'number', { cat: M, lift: true, min: 1, max: 1 }, ([x]) => Math.sign(num(x)));
def('SQRT', 'number', { cat: M, lift: true, min: 1, max: 1 }, ([x]) => { const v = num(x); if (v < 0) throw ERR.NUM(); return Math.sqrt(v); });
def('SQRTPI', 'number', { cat: M, lift: true, min: 1, max: 1 }, ([x]) => { const v = num(x); if (v < 0) throw ERR.NUM(); return Math.sqrt(v * Math.PI); });
def('POWER', 'number, power', { cat: M, lift: true, min: 2, max: 2 }, ([x, p]) => checkNum(Math.pow(num(x), num(p))));
def('EXP', 'number', { cat: M, lift: true, min: 1, max: 1 }, ([x]) => checkNum(Math.exp(num(x))));
def('LN', 'number', { cat: M, lift: true, min: 1, max: 1 }, ([x]) => { const v = num(x); if (v <= 0) throw ERR.NUM(); return Math.log(v); });
def('LOG', 'number, [base]', { cat: M, lift: true, min: 1, max: 2 }, ([x, b]) => { const v = num(x), base = b == null ? 10 : num(b); if (v <= 0 || base <= 0 || base === 1) throw ERR.NUM(); return Math.log(v) / Math.log(base); });
def('LOG10', 'number', { cat: M, lift: true, min: 1, max: 1 }, ([x]) => { const v = num(x); if (v <= 0) throw ERR.NUM(); return Math.log10(v); });
def('PI', '', { cat: M, max: 0 }, () => Math.PI);
def('MOD', 'number, divisor', { cat: M, lift: true, min: 2, max: 2 }, ([a, b]) => { const x = num(a), y = num(b); if (y === 0) throw ERR.DIV0(); return x - y * Math.floor(x / y); });
def('QUOTIENT', 'numerator, denominator', { cat: M, lift: true, min: 2, max: 2 }, ([a, b]) => { const y = num(b); if (y === 0) throw ERR.DIV0(); return Math.trunc(num(a) / y); });
def('INT', 'number', { cat: M, lift: true, min: 1, max: 1 }, ([x]) => Math.floor(num(x)));
def('TRUNC', 'number, [num_digits]', { cat: M, lift: true, min: 1, max: 2 }, ([x, d]) => { const f = Math.pow(10, d == null ? 0 : int(d)); return Math.trunc(num(x) * f) / f; });
def('EVEN', 'number', { cat: M, lift: true, min: 1, max: 1 }, ([x]) => { const v = num(x); const c = Math.ceil(Math.abs(v)); const e = c % 2 === 0 ? c : c + 1; return v < 0 ? -e : e; });
def('ODD', 'number', { cat: M, lift: true, min: 1, max: 1 }, ([x]) => { const v = num(x); const c = Math.ceil(Math.abs(v)); const e = c % 2 === 1 ? c : c + 1; return v < 0 ? -e : e; });
function roundHalfAway(v, d) {
  const f = Math.pow(10, d);
  const x = v * f;
  const r = Math.sign(x) * Math.round(Math.abs(x) + 1e-9 * Math.abs(x) * 0 + Number.EPSILON * Math.abs(x) * 10);
  return r / f;
}
export function excelRound(v, d) {
  d = Math.trunc(d);
  const f = Math.pow(10, Math.abs(d));
  let x = d >= 0 ? v * f : v / f;
  const s = Math.sign(x); x = Math.abs(x);
  // correct binary representation error (e.g. 2.675 -> 2.67499999)
  const corrected = Number(x.toPrecision(15));
  const r = Math.floor(corrected + 0.5);
  return s * (d >= 0 ? r / f : r * f);
}
def('ROUND', 'number, num_digits', { cat: M, lift: true, min: 1, max: 2 }, ([x, d]) => excelRound(num(x), d == null ? 0 : num(d)));
def('ROUNDUP', 'number, num_digits', { cat: M, lift: true, min: 1, max: 2 }, ([x, d]) => { const v = num(x), dd = d == null ? 0 : int(d); const f = Math.pow(10, Math.abs(dd)); const y = dd >= 0 ? v * f : v / f; const r = Math.sign(y) * Math.ceil(Number(Math.abs(y).toPrecision(15))); return dd >= 0 ? r / f : r * f; });
def('ROUNDDOWN', 'number, num_digits', { cat: M, lift: true, min: 1, max: 2 }, ([x, d]) => { const v = num(x), dd = d == null ? 0 : int(d); const f = Math.pow(10, Math.abs(dd)); const y = dd >= 0 ? v * f : v / f; const r = Math.sign(y) * Math.floor(Number(Math.abs(y).toPrecision(15))); return dd >= 0 ? r / f : r * f; });
def('MROUND', 'number, multiple', { cat: M, lift: true, min: 2, max: 2 }, ([x, m]) => { const v = num(x), mm = num(m); if (mm === 0) return 0; if ((v < 0) !== (mm < 0)) throw ERR.NUM(); return Math.round(v / mm) * mm; });
def(['CEILING', 'CEILING.MATH', 'ISO.CEILING'], 'number, [significance]', { cat: M, lift: true, min: 1, max: 3 }, ([x, s]) => { const v = num(x); const sig = s == null ? 1 : Math.abs(num(s)); if (sig === 0) return 0; return Number((Math.ceil(v / sig - 1e-12) * sig).toPrecision(15)); });
def(['FLOOR', 'FLOOR.MATH'], 'number, [significance]', { cat: M, lift: true, min: 1, max: 3 }, ([x, s]) => { const v = num(x); const sig = s == null ? 1 : Math.abs(num(s)); if (sig === 0) throw ERR.DIV0(); return Number((Math.floor(v / sig + 1e-12) * sig).toPrecision(15)); });
def('FACT', 'number', { cat: M, lift: true, min: 1, max: 1 }, ([x]) => { const n = int(x); if (n < 0) throw ERR.NUM(); let r = 1; for (let i = 2; i <= n; i++) r *= i; return checkNum(r); });
def('FACTDOUBLE', 'number', { cat: M, lift: true, min: 1, max: 1 }, ([x]) => { const n = int(x); if (n < 0) throw ERR.NUM(); let r = 1; for (let i = n; i > 1; i -= 2) r *= i; return checkNum(r); });
function combin(n, k) { if (k < 0 || n < 0 || k > n) throw ERR.NUM(); k = Math.min(k, n - k); let r = 1; for (let i = 1; i <= k; i++) r = r * (n - k + i) / i; return Math.round(r); }
def('COMBIN', 'number, number_chosen', { cat: M, lift: true, min: 2, max: 2 }, ([n, k]) => combin(int(n), int(k)));
def('COMBINA', 'number, number_chosen', { cat: M, lift: true, min: 2, max: 2 }, ([n, k]) => combin(int(n) + int(k) - 1, int(k)));
def('PERMUT', 'number, number_chosen', { cat: M, lift: true, min: 2, max: 2 }, ([n, k]) => { const nn = int(n), kk = int(k); if (kk < 0 || nn < kk) throw ERR.NUM(); let r = 1; for (let i = 0; i < kk; i++) r *= (nn - i); return r; });
def('PERMUTATIONA', 'number, number_chosen', { cat: M, lift: true, min: 2, max: 2 }, ([n, k]) => Math.pow(int(n), int(k)));
const gcd2 = (a, b) => { while (b) [a, b] = [b, a % b]; return a; };
def('GCD', 'number1, [number2], ...', { cat: M, min: 1 }, (args) => { const n = nums(args).map((v) => Math.floor(Math.abs(v))); return n.reduce(gcd2, 0); });
def('LCM', 'number1, [number2], ...', { cat: M, min: 1 }, (args) => { const n = nums(args).map((v) => Math.floor(Math.abs(v))); return n.reduce((a, b) => (a && b) ? a * b / gcd2(a, b) : 0, 1); });
def('RAND', '', { cat: M, volatile: true, max: 0 }, () => Math.random());
def('RANDBETWEEN', 'bottom, top', { cat: M, volatile: true, lift: true, min: 2, max: 2 }, ([a, b]) => { const lo = Math.ceil(num(a)), hi = Math.floor(num(b)); if (hi < lo) throw ERR.NUM(); return lo + Math.floor(Math.random() * (hi - lo + 1)); });
def('RANDARRAY', '[rows], [columns], [min], [max], [whole_number]', { cat: M, volatile: true, max: 5 }, ([r, c, lo, hi, whole]) => {
  const rows = r == null ? 1 : int(r), cols = c == null ? 1 : int(c); const mn = lo == null ? 0 : num(lo), mx = hi == null ? 1 : num(hi); const w = whole != null && bool(whole);
  if (rows < 1 || cols < 1) throw ERR.VALUE();
  return Array.from({ length: rows }, () => Array.from({ length: cols }, () => w ? Math.floor(mn + Math.random() * (mx - mn + 1)) : mn + Math.random() * (mx - mn)));
});
for (const [n, f] of Object.entries({ SIN: Math.sin, COS: Math.cos, TAN: Math.tan, ASIN: Math.asin, ACOS: Math.acos, ATAN: Math.atan, SINH: Math.sinh, COSH: Math.cosh, TANH: Math.tanh, ASINH: Math.asinh, ACOSH: Math.acosh, ATANH: Math.atanh }))
  def(n, 'number', { cat: M, lift: true, min: 1, max: 1 }, ([x]) => checkNum(f(num(x))));
def('COT', 'number', { cat: M, lift: true, min: 1, max: 1 }, ([x]) => { const v = num(x); if (v === 0) throw ERR.DIV0(); return 1 / Math.tan(v); });
def('SEC', 'number', { cat: M, lift: true, min: 1, max: 1 }, ([x]) => 1 / Math.cos(num(x)));
def('CSC', 'number', { cat: M, lift: true, min: 1, max: 1 }, ([x]) => { const v = num(x); if (v === 0) throw ERR.DIV0(); return 1 / Math.sin(v); });
def('ATAN2', 'x_num, y_num', { cat: M, lift: true, min: 2, max: 2 }, ([x, y]) => { const a = num(x), b = num(y); if (a === 0 && b === 0) throw ERR.DIV0(); return Math.atan2(b, a); });
def('DEGREES', 'angle', { cat: M, lift: true, min: 1, max: 1 }, ([x]) => num(x) * 180 / Math.PI);
def('RADIANS', 'angle', { cat: M, lift: true, min: 1, max: 1 }, ([x]) => num(x) * Math.PI / 180);
def('SUMPRODUCT', 'array1, [array2], ...', { cat: M, min: 1 }, (args) => {
  const ms = args.map(toMatrix);
  const [r, c] = dims(ms[0]);
  for (const m of ms) { const [rr, cc] = dims(m); if (rr !== r || cc !== c) throw ERR.VALUE('array sizes differ'); }
  let total = 0;
  for (let i = 0; i < r; i++) for (let j = 0; j < c; j++) { let p = 1; for (const m of ms) { const v = m[i][j]; if (isErr(v)) throw v; p *= typeof v === 'number' ? v : typeof v === 'boolean' && ms.length === 1 ? +v : (typeof v === 'string' && v.trim() !== '' && !isNaN(Number(v)) && false) ? Number(v) : 0; } total += p; }
  return total;
});
def('SUMIF', 'range, criteria, [sum_range]', { cat: M, min: 2, max: 3 }, ([rg, crit, sr]) => {
  const m = toMatrix(rg), test = criteria(crit); const sm = sr == null ? m : toMatrix(sr);
  let t = 0;
  for (let i = 0; i < m.length; i++) for (let j = 0; j < m[i].length; j++) if (test(m[i][j])) { const v = sm[i] ? sm[i][j] : null; if (isErr(v)) throw v; if (typeof v === 'number') t += v; }
  return t;
});
def('SUMIFS', 'sum_range, criteria_range1, criteria1, ...', { cat: M, min: 3 }, ([sr, ...rest]) => {
  const sm = toMatrix(sr); const pairs = ifsPairs(rest, 0); const [r, c] = dims(sm); const mask = ifsMask(pairs, r, c);
  let t = 0; for (let i = 0; i < r; i++) for (let j = 0; j < c; j++) if (mask[i][j]) { const v = sm[i][j]; if (isErr(v)) throw v; if (typeof v === 'number') t += v; }
  return t;
});
def('SUBTOTAL', 'function_num, ref1, [ref2], ...', { cat: M, min: 2 }, ([fn, ...refs], ctx) => {
  const code = int(fn) % 100;
  const hiddenOk = int(fn) < 100;
  // ignore nested SUBTOTAL cells and (for 101+) hidden rows
  const vals = [];
  for (const rg of refs) {
    if (isRange(rg)) {
      const { r1, c1, r2, c2 } = rg.clipped();
      for (let r = r1; r <= r2; r++) {
        if (!hiddenOk && rg.sheet.isRowHidden && rg.sheet.isRowHidden(r)) continue;
        if (rg.sheet.isRowFiltered && rg.sheet.isRowFiltered(r)) continue;
        for (let c = c1; c <= c2; c++) { const cell = rg.sheet.getCell(r, c); if (cell && cell.f && /SUBTOTAL\s*\(/i.test(cell.f)) continue; vals.push(rg.sheet.getValue(r, c)); }
      }
    } else each([rg], (v) => vals.push(v));
  }
  const names = { 1: 'AVERAGE', 2: 'COUNT', 3: 'COUNTA', 4: 'MAX', 5: 'MIN', 6: 'PRODUCT', 7: 'STDEV', 8: 'STDEVP', 9: 'SUM', 10: 'VAR', 11: 'VARP' };
  const name = names[code]; if (!name) throw ERR.VALUE();
  return FUNCTIONS.get(name).fn([[vals]], ctx);
});
def('AGGREGATE', 'function_num, options, array, [k]', { cat: M, min: 3 }, ([fn, opt, ...refs], ctx) => {
  const code = int(fn), options = int(opt);
  const ignoreErr = [2, 3, 6, 7].includes(options);
  const vals = []; for (const rg of refs.slice(0, 1)) each([rg], (v) => { if (ignoreErr && isErr(v)) return; vals.push(v); });
  const names = { 1: 'AVERAGE', 2: 'COUNT', 3: 'COUNTA', 4: 'MAX', 5: 'MIN', 6: 'PRODUCT', 7: 'STDEV', 8: 'STDEVP', 9: 'SUM', 10: 'VAR', 11: 'VARP', 12: 'MEDIAN', 13: 'MODE', 14: 'LARGE', 15: 'SMALL', 16: 'PERCENTILE.INC', 17: 'QUARTILE.INC', 18: 'PERCENTILE.EXC', 19: 'QUARTILE.EXC' };
  const name = names[code]; if (!name) throw ERR.VALUE();
  const extra = refs.slice(1);
  return FUNCTIONS.get(name).fn([[vals], ...extra], ctx);
});
def('ROMAN', 'number, [form]', { cat: M, lift: true, min: 1, max: 2 }, ([x]) => { let n = int(x); if (n < 0 || n > 3999) throw ERR.VALUE(); const t = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']]; let s = ''; for (const [v, r] of t) while (n >= v) { s += r; n -= v; } return s; });
def('ARABIC', 'text', { cat: M, lift: true, min: 1, max: 1 }, ([x]) => { const s = str(x).toUpperCase().trim(); const map = { I: 1, V: 5, X: 10, L: 50, C: 100, D: 500, M: 1000 }; let t = 0; for (let i = 0; i < s.length; i++) { const v = map[s[i]]; if (!v) throw ERR.VALUE(); const nx = map[s[i + 1]] || 0; t += v < nx ? -v : v; } return t; });
def('BASE', 'number, radix, [min_length]', { cat: M, lift: true, min: 2, max: 3 }, ([x, r, l]) => { const n = int(x), radix = int(r); if (n < 0 || radix < 2 || radix > 36) throw ERR.NUM(); return n.toString(radix).toUpperCase().padStart(l == null ? 0 : int(l), '0'); });
def('DECIMAL', 'text, radix', { cat: M, lift: true, min: 2, max: 2 }, ([x, r]) => { const v = parseInt(str(x), int(r)); if (isNaN(v)) throw ERR.NUM(); return v; });
def('SERIESSUM', 'x, n, m, coefficients', { cat: M, min: 4, max: 4 }, ([x, n, m, coef]) => { const xv = num(x), nv = num(n), mv = num(m); return flat(toMatrix(coef)).map(num).reduce((acc, c, i) => acc + c * Math.pow(xv, nv + i * mv), 0); });

function checkNum(v) { if (typeof v !== 'number' || !isFinite(v) || isNaN(v)) throw ERR.NUM(); return v; }

// ---------------- Statistics ----------------
const S = 'Statistical';
def('AVERAGE', 'number1, [number2], ...', { cat: S, min: 1 }, numsOrErr((n) => { if (!n.length) throw ERR.DIV0(); return n.reduce((a, b) => a + b, 0) / n.length; }));
def('AVERAGEA', 'value1, [value2], ...', { cat: S, min: 1 }, (args) => { const n = nums(args, { boolInRange: true, textInRange: true }); if (!n.length) throw ERR.DIV0(); return n.reduce((a, b) => a + b, 0) / n.length; });
def('COUNT', 'value1, [value2], ...', { cat: S, min: 1, errors: true }, (args) => { let c = 0; each(args, (v, direct) => { if (typeof v === 'number') c++; else if (direct && (typeof v === 'boolean' || (typeof v === 'string' && typeof parseInput(v).value === 'number'))) c++; }); return c; });
def('COUNTA', 'value1, [value2], ...', { cat: S, min: 1, errors: true }, (args) => { let c = 0; each(args, (v) => { if (v != null) c++; }); return c; });
def('COUNTBLANK', 'range', { cat: S, min: 1, max: 1 }, ([rg]) => { let c = 0; each([rg], (v) => { if (v == null || v === '') c++; }); return c; });
def('COUNTIF', 'range, criteria', { cat: S, min: 2, max: 2 }, ([rg, crit]) => { const test = criteria(crit); let c = 0; each([rg], (v) => { if (test(v)) c++; }); return c; });
def('COUNTIFS', 'criteria_range1, criteria1, ...', { cat: S, min: 2 }, (args) => { const pairs = ifsPairs(args, 0); const [r, c] = dims(pairs[0].m); const mask = ifsMask(pairs, r, c); let n = 0; for (const row of mask) for (const v of row) if (v) n++; return n; });
def('AVERAGEIF', 'range, criteria, [average_range]', { cat: S, min: 2, max: 3 }, ([rg, crit, ar]) => {
  const m = toMatrix(rg), test = criteria(crit); const am = ar == null ? m : toMatrix(ar); let t = 0, n = 0;
  for (let i = 0; i < m.length; i++) for (let j = 0; j < m[i].length; j++) if (test(m[i][j])) { const v = am[i] ? am[i][j] : null; if (isErr(v)) throw v; if (typeof v === 'number') { t += v; n++; } }
  if (!n) throw ERR.DIV0(); return t / n;
});
def('AVERAGEIFS', 'average_range, criteria_range1, criteria1, ...', { cat: S, min: 3 }, ([ar, ...rest]) => {
  const am = toMatrix(ar); const pairs = ifsPairs(rest, 0); const [r, c] = dims(am); const mask = ifsMask(pairs, r, c); let t = 0, n = 0;
  for (let i = 0; i < r; i++) for (let j = 0; j < c; j++) if (mask[i][j]) { const v = am[i][j]; if (isErr(v)) throw v; if (typeof v === 'number') { t += v; n++; } }
  if (!n) throw ERR.DIV0(); return t / n;
});
def('MAX', 'number1, [number2], ...', { cat: S, min: 1 }, numsOrErr((n) => n.length ? Math.max(...n) : 0));
def('MIN', 'number1, [number2], ...', { cat: S, min: 1 }, numsOrErr((n) => n.length ? Math.min(...n) : 0));
def('MAXA', 'value1, [value2], ...', { cat: S, min: 1 }, (args) => { const n = nums(args, { boolInRange: true, textInRange: true }); return n.length ? Math.max(...n) : 0; });
def('MINA', 'value1, [value2], ...', { cat: S, min: 1 }, (args) => { const n = nums(args, { boolInRange: true, textInRange: true }); return n.length ? Math.min(...n) : 0; });
def('MAXIFS', 'max_range, criteria_range1, criteria1, ...', { cat: S, min: 3 }, ([mr, ...rest]) => { const mm = toMatrix(mr); const pairs = ifsPairs(rest, 0); const [r, c] = dims(mm); const mask = ifsMask(pairs, r, c); let best = null; for (let i = 0; i < r; i++) for (let j = 0; j < c; j++) if (mask[i][j] && typeof mm[i][j] === 'number') best = best == null ? mm[i][j] : Math.max(best, mm[i][j]); return best == null ? 0 : best; });
def('MINIFS', 'min_range, criteria_range1, criteria1, ...', { cat: S, min: 3 }, ([mr, ...rest]) => { const mm = toMatrix(mr); const pairs = ifsPairs(rest, 0); const [r, c] = dims(mm); const mask = ifsMask(pairs, r, c); let best = null; for (let i = 0; i < r; i++) for (let j = 0; j < c; j++) if (mask[i][j] && typeof mm[i][j] === 'number') best = best == null ? mm[i][j] : Math.min(best, mm[i][j]); return best == null ? 0 : best; });
def('MEDIAN', 'number1, [number2], ...', { cat: S, min: 1 }, numsOrErr((n) => { if (!n.length) throw ERR.NUM(); n.sort((a, b) => a - b); const m = n.length >> 1; return n.length % 2 ? n[m] : (n[m - 1] + n[m]) / 2; }));
def(['MODE', 'MODE.SNGL'], 'number1, [number2], ...', { cat: S, min: 1 }, numsOrErr((n) => { const counts = new Map(); let best = null, bc = 1; for (const v of n) { const c = (counts.get(v) || 0) + 1; counts.set(v, c); if (c > bc || (c === bc && best == null)) { bc = c; best = v; } } if (best == null || bc < 2) throw ERR.NA(); return best; }));
function variance(n, sample) { if (n.length < (sample ? 2 : 1)) throw ERR.DIV0(); const mean = n.reduce((a, b) => a + b, 0) / n.length; const ss = n.reduce((a, b) => a + (b - mean) ** 2, 0); return ss / (n.length - (sample ? 1 : 0)); }
def(['VAR', 'VAR.S'], 'number1, [number2], ...', { cat: S, min: 1 }, numsOrErr((n) => variance(n, true)));
def(['VARP', 'VAR.P'], 'number1, [number2], ...', { cat: S, min: 1 }, numsOrErr((n) => variance(n, false)));
def('VARA', 'value1, [value2], ...', { cat: S, min: 1 }, (args) => variance(nums(args, { boolInRange: true, textInRange: true }), true));
def('VARPA', 'value1, [value2], ...', { cat: S, min: 1 }, (args) => variance(nums(args, { boolInRange: true, textInRange: true }), false));
def(['STDEV', 'STDEV.S'], 'number1, [number2], ...', { cat: S, min: 1 }, numsOrErr((n) => Math.sqrt(variance(n, true))));
def(['STDEVP', 'STDEV.P'], 'number1, [number2], ...', { cat: S, min: 1 }, numsOrErr((n) => Math.sqrt(variance(n, false))));
def('STDEVA', 'value1, [value2], ...', { cat: S, min: 1 }, (args) => Math.sqrt(variance(nums(args, { boolInRange: true, textInRange: true }), true)));
def('STDEVPA', 'value1, [value2], ...', { cat: S, min: 1 }, (args) => Math.sqrt(variance(nums(args, { boolInRange: true, textInRange: true }), false)));
def('AVEDEV', 'number1, [number2], ...', { cat: S, min: 1 }, numsOrErr((n) => { if (!n.length) throw ERR.NUM(); const mean = n.reduce((a, b) => a + b, 0) / n.length; return n.reduce((a, b) => a + Math.abs(b - mean), 0) / n.length; }));
def('DEVSQ', 'number1, [number2], ...', { cat: S, min: 1 }, numsOrErr((n) => { if (!n.length) throw ERR.NUM(); const mean = n.reduce((a, b) => a + b, 0) / n.length; return n.reduce((a, b) => a + (b - mean) ** 2, 0); }));
def('GEOMEAN', 'number1, [number2], ...', { cat: S, min: 1 }, numsOrErr((n) => { if (!n.length || n.some((v) => v <= 0)) throw ERR.NUM(); return Math.exp(n.reduce((a, b) => a + Math.log(b), 0) / n.length); }));
def('HARMEAN', 'number1, [number2], ...', { cat: S, min: 1 }, numsOrErr((n) => { if (!n.length || n.some((v) => v <= 0)) throw ERR.NUM(); return n.length / n.reduce((a, b) => a + 1 / b, 0); }));
def('LARGE', 'array, k', { cat: S, min: 2, max: 2 }, ([a, k]) => { const n = nums([a]).sort((x, y) => y - x); const kk = int(k); if (kk < 1 || kk > n.length) throw ERR.NUM(); return n[kk - 1]; });
def('SMALL', 'array, k', { cat: S, min: 2, max: 2 }, ([a, k]) => { const n = nums([a]).sort((x, y) => x - y); const kk = int(k); if (kk < 1 || kk > n.length) throw ERR.NUM(); return n[kk - 1]; });
def(['RANK', 'RANK.EQ'], 'number, ref, [order]', { cat: S, min: 2, max: 3 }, ([x, rg, o]) => { const v = num(x); const n = nums([rg]); const asc = o != null && bool(o); if (!n.includes(v)) throw ERR.NA(); return 1 + n.filter((y) => asc ? y < v : y > v).length; });
def('RANK.AVG', 'number, ref, [order]', { cat: S, min: 2, max: 3 }, ([x, rg, o]) => { const v = num(x); const n = nums([rg]); const asc = o != null && bool(o); if (!n.includes(v)) throw ERR.NA(); const less = n.filter((y) => asc ? y < v : y > v).length; const eq = n.filter((y) => y === v).length; return less + (eq + 1) / 2; });
function percentileInc(n, p) { if (!n.length || p < 0 || p > 1) throw ERR.NUM(); n.sort((a, b) => a - b); const idx = p * (n.length - 1); const lo = Math.floor(idx), hi = Math.ceil(idx); return n[lo] + (n[hi] - n[lo]) * (idx - lo); }
function percentileExc(n, p) { n.sort((a, b) => a - b); const idx = p * (n.length + 1) - 1; if (idx < 0 || idx > n.length - 1) throw ERR.NUM(); const lo = Math.floor(idx), hi = Math.ceil(idx); return n[lo] + (n[hi] - n[lo]) * (idx - lo); }
def(['PERCENTILE', 'PERCENTILE.INC'], 'array, k', { cat: S, min: 2, max: 2 }, ([a, k]) => percentileInc(nums([a]), num(k)));
def('PERCENTILE.EXC', 'array, k', { cat: S, min: 2, max: 2 }, ([a, k]) => percentileExc(nums([a]), num(k)));
def(['QUARTILE', 'QUARTILE.INC'], 'array, quart', { cat: S, min: 2, max: 2 }, ([a, q]) => { const qq = int(q); if (qq < 0 || qq > 4) throw ERR.NUM(); return percentileInc(nums([a]), qq / 4); });
def('QUARTILE.EXC', 'array, quart', { cat: S, min: 2, max: 2 }, ([a, q]) => { const qq = int(q); if (qq < 1 || qq > 3) throw ERR.NUM(); return percentileExc(nums([a]), qq / 4); });
def(['PERCENTRANK', 'PERCENTRANK.INC'], 'array, x, [significance]', { cat: S, min: 2, max: 3 }, ([a, x, s]) => { const n = nums([a]).sort((p, q) => p - q); const v = num(x); if (v < n[0] || v > n[n.length - 1]) throw ERR.NA(); let i = 0; while (i < n.length - 1 && n[i + 1] <= v) i++; let r; if (n[i] === v) r = i / (n.length - 1); else r = (i + (v - n[i]) / (n[i + 1] - n[i])) / (n.length - 1); const sig = s == null ? 3 : int(s); return Math.floor(r * 10 ** sig) / 10 ** sig; });
function pairs(a, b) { const x = flat(toMatrix(a)), y = flat(toMatrix(b)); if (x.length !== y.length) throw ERR.NA(); const xs = [], ys = []; for (let i = 0; i < x.length; i++) if (typeof x[i] === 'number' && typeof y[i] === 'number') { xs.push(x[i]); ys.push(y[i]); } if (!xs.length) throw ERR.DIV0(); return [xs, ys]; }
const mean = (n) => n.reduce((a, b) => a + b, 0) / n.length;
function covar(xs, ys, sample) { const mx = mean(xs), my = mean(ys); let s = 0; for (let i = 0; i < xs.length; i++) s += (xs[i] - mx) * (ys[i] - my); return s / (xs.length - (sample ? 1 : 0)); }
def(['CORREL', 'PEARSON'], 'array1, array2', { cat: S, min: 2, max: 2 }, ([a, b]) => { const [xs, ys] = pairs(a, b); const d = Math.sqrt(variance(xs, false) * variance(ys, false)); if (d === 0) throw ERR.DIV0(); return covar(xs, ys, false) / d; });
def('RSQ', 'known_ys, known_xs', { cat: S, min: 2, max: 2 }, ([a, b]) => { const [xs, ys] = pairs(a, b); const d = Math.sqrt(variance(xs, false) * variance(ys, false)); if (d === 0) throw ERR.DIV0(); return (covar(xs, ys, false) / d) ** 2; });
def(['COVAR', 'COVARIANCE.P'], 'array1, array2', { cat: S, min: 2, max: 2 }, ([a, b]) => { const [xs, ys] = pairs(a, b); return covar(xs, ys, false); });
def('COVARIANCE.S', 'array1, array2', { cat: S, min: 2, max: 2 }, ([a, b]) => { const [xs, ys] = pairs(a, b); if (xs.length < 2) throw ERR.DIV0(); return covar(xs, ys, true); });
function linfit(ysArg, xsArg) { const [ys, xs] = pairs(ysArg, xsArg); const vx = variance(xs, false); if (vx === 0) throw ERR.DIV0(); const slope = covar(xs, ys, false) / vx; return { slope, intercept: mean(ys) - slope * mean(xs) }; }
def('SLOPE', 'known_ys, known_xs', { cat: S, min: 2, max: 2 }, ([y, x]) => linfit(y, x).slope);
def('INTERCEPT', 'known_ys, known_xs', { cat: S, min: 2, max: 2 }, ([y, x]) => linfit(y, x).intercept);
def(['FORECAST', 'FORECAST.LINEAR'], 'x, known_ys, known_xs', { cat: S, min: 3, max: 3 }, ([x, y, xs]) => { const f = linfit(y, xs); return f.intercept + f.slope * num(x); });
def('TREND', 'known_ys, [known_xs], [new_xs]', { cat: S, min: 1, max: 4 }, ([y, xs, nx]) => { const ym = toMatrix(y); const xsm = xs == null ? ym.map((_, i) => [i + 1]) : toMatrix(xs); const f = linfit(ym, xsm); const target = nx == null ? xsm : toMatrix(nx); return target.map((row) => row.map((v) => f.intercept + f.slope * num(v))); });
def('STEYX', 'known_ys, known_xs', { cat: S, min: 2, max: 2 }, ([y, x]) => { const [xs, ys] = pairs(x, y); const f = linfit(y, x); if (xs.length < 3) throw ERR.DIV0(); const ss = xs.reduce((a, xv, i) => a + (ys[i] - (f.intercept + f.slope * xv)) ** 2, 0); return Math.sqrt(ss / (xs.length - 2)); });
def('STANDARDIZE', 'x, mean, standard_dev', { cat: S, lift: true, min: 3, max: 3 }, ([x, m, s]) => { const sd = num(s); if (sd <= 0) throw ERR.NUM(); return (num(x) - num(m)) / sd; });
function erf(x) { const t = 1 / (1 + 0.3275911 * Math.abs(x)); const y = 1 - (((((1.061405429 * t - 1.453152027) * t) + 1.421413741) * t - 0.284496736) * t + 0.254829592) * t * Math.exp(-x * x); return x >= 0 ? y : -y; }
function normCdf(z) { return 0.5 * (1 + erf(z / Math.SQRT2)); }
function normPdf(z) { return Math.exp(-z * z / 2) / Math.sqrt(2 * Math.PI); }
function normInv(p) { // Acklam's algorithm
  if (p <= 0 || p >= 1) throw ERR.NUM();
  const a = [-3.969683028665376e+01, 2.209460984245205e+02, -2.759285104469687e+02, 1.383577518672690e+02, -3.066479806614716e+01, 2.506628277459239e+00];
  const b = [-5.447609879822406e+01, 1.615858368580409e+02, -1.556989798598866e+02, 6.680131188771972e+01, -1.328068155288572e+01];
  const c = [-7.784894002430293e-03, -3.223964580411365e-01, -2.400758277161838e+00, -2.549732539343734e+00, 4.374664141464968e+00, 2.938163982698783e+00];
  const d = [7.784695709041462e-03, 3.224671290700398e-01, 2.445134137142996e+00, 3.754408661907416e+00];
  const pl = 0.02425, ph = 1 - pl; let q, r;
  if (p < pl) { q = Math.sqrt(-2 * Math.log(p)); return (((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1); }
  if (p <= ph) { q = p - 0.5; r = q * q; return (((((a[0] * r + a[1]) * r + a[2]) * r + a[3]) * r + a[4]) * r + a[5]) * q / (((((b[0] * r + b[1]) * r + b[2]) * r + b[3]) * r + b[4]) * r + 1); }
  q = Math.sqrt(-2 * Math.log(1 - p)); return -(((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1);
}
def(['NORM.DIST', 'NORMDIST'], 'x, mean, standard_dev, cumulative', { cat: S, lift: true, min: 4, max: 4 }, ([x, m, s, c]) => { const sd = num(s); if (sd <= 0) throw ERR.NUM(); const z = (num(x) - num(m)) / sd; return bool(c) ? normCdf(z) : normPdf(z) / sd; });
def(['NORM.S.DIST', 'NORMSDIST'], 'z, [cumulative]', { cat: S, lift: true, min: 1, max: 2 }, ([z, c]) => (c == null || bool(c)) ? normCdf(num(z)) : normPdf(num(z)));
def(['NORM.INV', 'NORMINV'], 'probability, mean, standard_dev', { cat: S, lift: true, min: 3, max: 3 }, ([p, m, s]) => { const sd = num(s); if (sd <= 0) throw ERR.NUM(); return num(m) + sd * normInv(num(p)); });
def(['NORM.S.INV', 'NORMSINV'], 'probability', { cat: S, lift: true, min: 1, max: 1 }, ([p]) => normInv(num(p)));
def(['CONFIDENCE', 'CONFIDENCE.NORM'], 'alpha, standard_dev, size', { cat: S, lift: true, min: 3, max: 3 }, ([a, s, n]) => { const al = num(a); if (al <= 0 || al >= 1) throw ERR.NUM(); return normInv(1 - al / 2) * num(s) / Math.sqrt(int(n)); });
def(['BINOM.DIST', 'BINOMDIST'], 'number_s, trials, probability_s, cumulative', { cat: S, lift: true, min: 4, max: 4 }, ([k, n, p, c]) => { const kk = int(k), nn = int(n), pp = num(p); if (kk < 0 || kk > nn || pp < 0 || pp > 1) throw ERR.NUM(); const pmf = (i) => combin(nn, i) * pp ** i * (1 - pp) ** (nn - i); if (!bool(c)) return pmf(kk); let t = 0; for (let i = 0; i <= kk; i++) t += pmf(i); return t; });
def(['POISSON', 'POISSON.DIST'], 'x, mean, cumulative', { cat: S, lift: true, min: 3, max: 3 }, ([x, m, c]) => { const k = int(x), l = num(m); if (k < 0 || l < 0) throw ERR.NUM(); const pmf = (i) => Math.exp(-l) * l ** i / FUNCTIONS.get('FACT').fn([i]); if (!bool(c)) return pmf(k); let t = 0; for (let i = 0; i <= k; i++) t += pmf(i); return t; });
def(['EXPON.DIST', 'EXPONDIST'], 'x, lambda, cumulative', { cat: S, lift: true, min: 3, max: 3 }, ([x, l, c]) => { const xv = num(x), lv = num(l); if (xv < 0 || lv <= 0) throw ERR.NUM(); return bool(c) ? 1 - Math.exp(-lv * xv) : lv * Math.exp(-lv * xv); });
def('SKEW', 'number1, [number2], ...', { cat: S, min: 1 }, numsOrErr((n) => { const k = n.length; if (k < 3) throw ERR.DIV0(); const m = mean(n), s = Math.sqrt(variance(n, true)); if (s === 0) throw ERR.DIV0(); return k / ((k - 1) * (k - 2)) * n.reduce((a, v) => a + ((v - m) / s) ** 3, 0); }));
def('KURT', 'number1, [number2], ...', { cat: S, min: 1 }, numsOrErr((n) => { const k = n.length; if (k < 4) throw ERR.DIV0(); const m = mean(n), s = Math.sqrt(variance(n, true)); if (s === 0) throw ERR.DIV0(); return k * (k + 1) / ((k - 1) * (k - 2) * (k - 3)) * n.reduce((a, v) => a + ((v - m) / s) ** 4, 0) - 3 * (k - 1) ** 2 / ((k - 2) * (k - 3)); }));
def('FREQUENCY', 'data_array, bins_array', { cat: S, min: 2, max: 2 }, ([d, b]) => { const data = nums([d]); const bins = nums([b]).sort((x, y) => x - y); const out = new Array(bins.length + 1).fill(0); for (const v of data) { let i = bins.findIndex((bn) => v <= bn); if (i < 0) i = bins.length; out[i]++; } return out.map((v) => [v]); });

// ---------------- Text ----------------
const T = 'Text';
def('LEFT', 'text, [num_chars]', { cat: T, lift: true, min: 1, max: 2 }, ([t, n]) => { const k = n == null ? 1 : int(n); if (k < 0) throw ERR.VALUE(); return [...str(t)].slice(0, k).join(''); });
def('RIGHT', 'text, [num_chars]', { cat: T, lift: true, min: 1, max: 2 }, ([t, n]) => { const k = n == null ? 1 : int(n); if (k < 0) throw ERR.VALUE(); const a = [...str(t)]; return k === 0 ? '' : a.slice(-k).join(''); });
def('MID', 'text, start_num, num_chars', { cat: T, lift: true, min: 3, max: 3 }, ([t, s, n]) => { const st = int(s), k = int(n); if (st < 1 || k < 0) throw ERR.VALUE(); return [...str(t)].slice(st - 1, st - 1 + k).join(''); });
def('LEN', 'text', { cat: T, lift: true, min: 1, max: 1 }, ([t]) => [...str(t)].length);
def('LOWER', 'text', { cat: T, lift: true, min: 1, max: 1 }, ([t]) => str(t).toLowerCase());
def('UPPER', 'text', { cat: T, lift: true, min: 1, max: 1 }, ([t]) => str(t).toUpperCase());
def('PROPER', 'text', { cat: T, lift: true, min: 1, max: 1 }, ([t]) => str(t).toLowerCase().replace(/(^|[^a-zA-ZÀ-ɏ])([a-zà-ɏ])/g, (m, a, b) => a + b.toUpperCase()));
def('TRIM', 'text', { cat: T, lift: true, min: 1, max: 1 }, ([t]) => str(t).replace(/ +/g, ' ').trim());
def('CLEAN', 'text', { cat: T, lift: true, min: 1, max: 1 }, ([t]) => str(t).replace(/[\x00-\x1f]/g, ''));
def(['CONCAT', 'CONCATENATE'], 'text1, [text2], ...', { cat: T, min: 1 }, (args) => { let s = ''; each(args, (v) => { if (isErr(v)) throw v; s += str(v); }); return s; });
def('TEXTJOIN', 'delimiter, ignore_empty, text1, [text2], ...', { cat: T, min: 3 }, ([d, ie, ...rest]) => { const delim = str(d), ign = bool(ie); const parts = []; each(rest, (v) => { if (isErr(v)) throw v; if (ign && isEmpty(v)) return; parts.push(str(v)); }); return parts.join(delim); });
def('REPT', 'text, number_times', { cat: T, lift: true, min: 2, max: 2 }, ([t, n]) => { const k = int(n); if (k < 0) throw ERR.VALUE(); const s = str(t); if (s.length * k > 32767) throw ERR.VALUE(); return s.repeat(k); });
def('SUBSTITUTE', 'text, old_text, new_text, [instance_num]', { cat: T, lift: true, min: 3, max: 4 }, ([t, o, nw, inst]) => { const s = str(t), old = str(o), rep = str(nw); if (old === '') return s; if (inst == null) return s.split(old).join(rep); const k = int(inst); if (k < 1) throw ERR.VALUE(); let idx = -1; for (let i = 0; i < k; i++) { idx = s.indexOf(old, idx + 1); if (idx < 0) return s; } return s.slice(0, idx) + rep + s.slice(idx + old.length); });
def('REPLACE', 'old_text, start_num, num_chars, new_text', { cat: T, lift: true, min: 4, max: 4 }, ([t, s, n, nw]) => { const a = [...str(t)], st = int(s), k = int(n); if (st < 1 || k < 0) throw ERR.VALUE(); return a.slice(0, st - 1).join('') + str(nw) + a.slice(st - 1 + k).join(''); });
def('FIND', 'find_text, within_text, [start_num]', { cat: T, lift: true, min: 2, max: 3 }, ([f, w, s]) => { const st = s == null ? 1 : int(s); const hay = str(w); if (st < 1 || st > hay.length + 1) throw ERR.VALUE(); const i = hay.indexOf(str(f), st - 1); if (i < 0) throw ERR.VALUE(); return i + 1; });
def('SEARCH', 'find_text, within_text, [start_num]', { cat: T, lift: true, min: 2, max: 3 }, ([f, w, s]) => { const st = s == null ? 1 : int(s); const hay = str(w); if (st < 1 || st > hay.length + 1) throw ERR.VALUE(); const re = new RegExp(wildcardToRegex(str(f)).source.replace(/^\^|\$$/g, ''), 'i'); const m = re.exec(hay.slice(st - 1)); if (!m) throw ERR.VALUE(); return m.index + st; });
def('EXACT', 'text1, text2', { cat: T, lift: true, min: 2, max: 2 }, ([a, b]) => str(a) === str(b));
def('TEXT', 'value, format_text', { cat: T, lift: true, min: 2, max: 2 }, ([v, f]) => { if (isErr(v)) throw v; const code = str(f); const val = typeof v === 'string' ? (parseInput(v).value ?? v) : v; return formatValue(val, code).text; });
def('VALUE', 'text', { cat: T, lift: true, min: 1, max: 1 }, ([t]) => num(t));
def('NUMBERVALUE', 'text, [decimal_separator], [group_separator]', { cat: T, lift: true, min: 1, max: 3 }, ([t, d, g]) => { let s = str(t); if (g != null) s = s.split(str(g)).join(''); if (d != null) s = s.replace(str(d), '.'); return num(s.replace(/\s/g, '')); });
def('T', 'value', { cat: T, lift: true, min: 1, max: 1 }, ([v]) => typeof v === 'string' ? v : '');
def('N', 'value', { cat: T, lift: true, min: 1, max: 1, errors: true }, ([v]) => typeof v === 'number' ? v : typeof v === 'boolean' ? +v : isErr(v) ? v : 0);
def('CHAR', 'number', { cat: T, lift: true, min: 1, max: 1 }, ([n]) => { const k = int(n); if (k < 1 || k > 255) throw ERR.VALUE(); return String.fromCharCode(k); });
def('UNICHAR', 'number', { cat: T, lift: true, min: 1, max: 1 }, ([n]) => { const k = int(n); if (k < 1) throw ERR.VALUE(); return String.fromCodePoint(k); });
def('CODE', 'text', { cat: T, lift: true, min: 1, max: 1 }, ([t]) => { const s = str(t); if (!s) throw ERR.VALUE(); return s.charCodeAt(0); });
def('UNICODE', 'text', { cat: T, lift: true, min: 1, max: 1 }, ([t]) => { const s = str(t); if (!s) throw ERR.VALUE(); return s.codePointAt(0); });
def('DOLLAR', 'number, [decimals]', { cat: T, lift: true, min: 1, max: 2 }, ([n, d]) => { const dd = d == null ? 2 : int(d); const code = dd > 0 ? `$#,##0.${'0'.repeat(dd)};($#,##0.${'0'.repeat(dd)})` : '$#,##0;($#,##0)'; return formatValue(dd < 0 ? excelRound(num(n), dd) : num(n), code).text; });
def('FIXED', 'number, [decimals], [no_commas]', { cat: T, lift: true, min: 1, max: 3 }, ([n, d, nc]) => { const dd = d == null ? 2 : int(d); const code = (nc != null && bool(nc) ? '0' : '#,##0') + (dd > 0 ? '.' + '0'.repeat(dd) : ''); return formatValue(dd < 0 ? excelRound(num(n), dd) : num(n), code).text; });
def('TEXTBEFORE', 'text, delimiter, [instance_num], [match_mode]', { cat: T, lift: true, min: 2, max: 6 }, ([t, d, i, mm]) => { const s = str(t), del = str(d), ci = mm != null && bool(mm); const hay = ci ? s.toLowerCase() : s, needle = ci ? del.toLowerCase() : del; let k = i == null ? 1 : int(i); let idx = -1; if (k > 0) { for (let n = 0; n < k; n++) { idx = hay.indexOf(needle, idx + 1); if (idx < 0) throw ERR.NA(); } } else { idx = hay.length; for (let n = 0; n < -k; n++) { idx = hay.lastIndexOf(needle, idx - 1); if (idx < 0) throw ERR.NA(); } } return s.slice(0, idx); });
def('TEXTAFTER', 'text, delimiter, [instance_num], [match_mode]', { cat: T, lift: true, min: 2, max: 6 }, ([t, d, i, mm]) => { const s = str(t), del = str(d), ci = mm != null && bool(mm); const hay = ci ? s.toLowerCase() : s, needle = ci ? del.toLowerCase() : del; let k = i == null ? 1 : int(i); let idx = -1; if (k > 0) { for (let n = 0; n < k; n++) { idx = hay.indexOf(needle, idx + 1); if (idx < 0) throw ERR.NA(); } } else { idx = hay.length; for (let n = 0; n < -k; n++) { idx = hay.lastIndexOf(needle, idx - 1); if (idx < 0) throw ERR.NA(); } } return s.slice(idx + del.length); });
def('TEXTSPLIT', 'text, col_delimiter, [row_delimiter], [ignore_empty]', { cat: T, min: 2, max: 6 }, ([t, cd, rd, ie]) => { const s = str(t); const rows = rd == null ? [s] : s.split(str(rd)); const ign = ie != null && bool(ie); const out = rows.map((r) => { let parts = cd == null || cd === '' ? [r] : r.split(str(cd)); if (ign) parts = parts.filter((p) => p !== ''); return parts; }); const w = Math.max(...out.map((r) => r.length)); return out.map((r) => { while (r.length < w) r.push(ERR.NA()); return r; }); });
def('ENCODEURL', 'text', { cat: T, lift: true, min: 1, max: 1 }, ([t]) => encodeURIComponent(str(t)));
def('HYPERLINK', 'link_location, [friendly_name]', { cat: T, lift: true, min: 1, max: 2 }, ([l, f]) => f == null ? str(l) : scalar(f));

// ---------------- Logical ----------------
const L = 'Logical';
def('TRUE', '', { cat: L, max: 0 }, () => true);
def('FALSE', '', { cat: L, max: 0 }, () => false);
def('NOT', 'logical', { cat: L, lift: true, min: 1, max: 1 }, ([v]) => !bool(v));
function logicalAgg(args, init, op) { let acc = init, any = false; each(args, (v, direct) => { if (isErr(v)) throw v; if (typeof v === 'boolean' || typeof v === 'number') { acc = op(acc, !!v); any = true; } else if (direct && typeof v === 'string') { acc = op(acc, bool(v)); any = true; } }); if (!any) throw ERR.VALUE(); return acc; }
def('AND', 'logical1, [logical2], ...', { cat: L, min: 1 }, (args) => logicalAgg(args, true, (a, b) => a && b));
def('OR', 'logical1, [logical2], ...', { cat: L, min: 1 }, (args) => logicalAgg(args, false, (a, b) => a || b));
def('XOR', 'logical1, [logical2], ...', { cat: L, min: 1 }, (args) => logicalAgg(args, false, (a, b) => a !== b));
def('IF', 'logical_test, [value_if_true], [value_if_false]', { cat: L, lazy: true, min: 1, max: 3 }, (nodes, ctx) => {
  const cond = scalar(ctx.eval(nodes[0]));
  if (isErr(cond)) return cond;
  const pick = (c) => { const b = bool(c); const node = b ? nodes[1] : nodes[2]; if (!node || node.type === 'empty') return b ? true : false; return scalar(ctx.eval(node)); };
  if (isArr(cond)) {
    const t = nodes[1] && nodes[1].type !== 'empty' ? scalar(ctx.eval(nodes[1])) : true;
    const f = nodes[2] && nodes[2].type !== 'empty' ? scalar(ctx.eval(nodes[2])) : false;
    return broadcast([cond, t, f], ([c, a, b]) => { if (isErr(c)) return c; try { return bool(c) ? a : b; } catch (e) { return e; } });
  }
  try { return pick(cond); } catch (e) { if (isErr(e)) return e; throw e; }
});
def('IFS', 'logical_test1, value_if_true1, ...', { cat: L, lazy: true, min: 2 }, (nodes, ctx) => {
  for (let i = 0; i + 1 < nodes.length; i += 2) { const c = scalar(ctx.eval(nodes[i])); if (isErr(c)) return c; let b; try { b = bool(c); } catch (e) { return e; } if (b) return scalar(ctx.eval(nodes[i + 1])); }
  return ERR.NA();
});
def('IFERROR', 'value, value_if_error', { cat: L, lazy: true, min: 2, max: 2 }, (nodes, ctx) => { const v = scalar(ctx.eval(nodes[0])); if (isArr(v)) { const alt = scalar(ctx.eval(nodes[1])); return broadcast([v, alt], ([a, b]) => isErr(a) ? b : a); } return isErr(v) ? scalar(ctx.eval(nodes[1])) : v; });
def('IFNA', 'value, value_if_na', { cat: L, lazy: true, min: 2, max: 2 }, (nodes, ctx) => { const v = scalar(ctx.eval(nodes[0])); return isErr(v) && v.error === '#N/A' ? scalar(ctx.eval(nodes[1])) : v; });
def('SWITCH', 'expression, value1, result1, ..., [default]', { cat: L, lazy: true, min: 3 }, (nodes, ctx) => {
  const x = scalar(ctx.eval(nodes[0])); if (isErr(x)) return x;
  let i = 1;
  for (; i + 1 < nodes.length; i += 2) { const v = scalar(ctx.eval(nodes[i])); if (compareValues(x, v) === 0 && typeof x === typeof v) return scalar(ctx.eval(nodes[i + 1])); }
  if (i < nodes.length) return scalar(ctx.eval(nodes[i]));
  return ERR.NA();
});
def('LET', 'name1, value1, [name2, value2], ..., calculation', { cat: L, lazy: true, min: 3 }, (nodes, ctx) => {
  const scope = Object.create(ctx.scope || null);
  let i = 0;
  for (; i + 1 < nodes.length - 1; i += 2) {
    const n = nodes[i]; if (n.type !== 'name' && n.type !== 'ref') return ERR.NAME();
    const key = (n.type === 'name' ? n.name : refNodeName(n)).toUpperCase();
    scope[key] = ctx.evalWithScope(nodes[i + 1], scope);
  }
  return ctx.evalWithScope(nodes[nodes.length - 1], scope);
});
function refNodeName(n) { return cellName(n.r1, n.c1); }

// ---------------- Information ----------------
const I = 'Information';
def('ISBLANK', 'value', { cat: I, lift: true, min: 1, max: 1, errors: true }, ([v]) => v == null);
def('ISNUMBER', 'value', { cat: I, lift: true, min: 1, max: 1, errors: true }, ([v]) => typeof v === 'number');
def('ISTEXT', 'value', { cat: I, lift: true, min: 1, max: 1, errors: true }, ([v]) => typeof v === 'string');
def('ISNONTEXT', 'value', { cat: I, lift: true, min: 1, max: 1, errors: true }, ([v]) => typeof v !== 'string');
def('ISLOGICAL', 'value', { cat: I, lift: true, min: 1, max: 1, errors: true }, ([v]) => typeof v === 'boolean');
def('ISERROR', 'value', { cat: I, lift: true, min: 1, max: 1, errors: true }, ([v]) => isErr(v));
def('ISERR', 'value', { cat: I, lift: true, min: 1, max: 1, errors: true }, ([v]) => isErr(v) && v.error !== '#N/A');
def('ISNA', 'value', { cat: I, lift: true, min: 1, max: 1, errors: true }, ([v]) => isErr(v) && v.error === '#N/A');
def('ISEVEN', 'number', { cat: I, lift: true, min: 1, max: 1 }, ([v]) => Math.trunc(num(v)) % 2 === 0);
def('ISODD', 'number', { cat: I, lift: true, min: 1, max: 1 }, ([v]) => Math.abs(Math.trunc(num(v))) % 2 === 1);
def('ISREF', 'value', { cat: I, min: 1, max: 1, errors: true }, ([v]) => isRange(v));
def('ISFORMULA', 'reference', { cat: I, min: 1, max: 1 }, ([v]) => { if (!isRange(v)) throw ERR.VALUE(); const c = v.sheet.getCell(v.r1, v.c1); return !!(c && c.f); });
def('NA', '', { cat: I, max: 0 }, () => ERR.NA());
def('ERROR.TYPE', 'error_val', { cat: I, lift: true, min: 1, max: 1, errors: true }, ([v]) => { if (!isErr(v)) throw ERR.NA(); const i = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A'].indexOf(v.error); return i < 0 ? 8 : i + 1; });
def('TYPE', 'value', { cat: I, min: 1, max: 1, errors: true }, ([v]) => { v = scalar(v); if (isArr(v)) return 64; if (typeof v === 'number' || v == null) return 1; if (typeof v === 'string') return 2; if (typeof v === 'boolean') return 4; if (isErr(v)) return 16; return 1; });
def('SHEET', '[value]', { cat: I, max: 1 }, ([v], ctx) => { if (v == null) return ctx.sheet.index() + 1; if (isRange(v)) return v.sheet.index() + 1; const s = ctx.wb.getSheet(str(v)); if (!s) throw ERR.NA(); return s.index() + 1; });
def('SHEETS', '', { cat: I, max: 0 }, (_, ctx) => ctx.wb.sheets.length);
def('FORMULATEXT', 'reference', { cat: I, min: 1, max: 1 }, ([v]) => { if (!isRange(v)) throw ERR.VALUE(); const c = v.sheet.getCell(v.r1, v.c1); if (!c || !c.f) throw ERR.NA(); return c.f; });
def('CELL', 'info_type, [reference]', { cat: I, min: 1, max: 2 }, ([t, rg], ctx) => {
  const type = str(t).toLowerCase(); const r = isRange(rg) ? rg : new RangeRef(ctx.sheet, ctx.row, ctx.col, ctx.row, ctx.col);
  const v = r.value();
  switch (type) {
    case 'address': return `$${colToName(r.c1)}$${r.r1 + 1}`; case 'col': return r.c1 + 1; case 'row': return r.r1 + 1;
    case 'contents': return v == null ? 0 : v; case 'type': return v == null ? 'b' : typeof v === 'string' ? 'l' : 'v';
    case 'filename': return `[${ctx.wb.name || 'Book1'}]${r.sheet.name}`; case 'format': return r.sheet.getStyle(r.r1, r.c1).numFmt || 'G';
    case 'width': return Math.round(r.sheet.colWidth(r.c1) / 7); default: throw ERR.VALUE();
  }
});

// ---------------- Lookup & reference ----------------
const K = 'Lookup';
def('ROW', '[reference]', { cat: K, max: 1 }, ([rg], ctx) => { if (rg == null) return ctx.row + 1; if (!isRange(rg)) throw ERR.VALUE(); if (rg.rows === 1) return rg.r1 + 1; return Array.from({ length: rg.rows }, (_, i) => [rg.r1 + i + 1]); });
def('COLUMN', '[reference]', { cat: K, max: 1 }, ([rg], ctx) => { if (rg == null) return ctx.col + 1; if (!isRange(rg)) throw ERR.VALUE(); if (rg.cols === 1) return rg.c1 + 1; return [Array.from({ length: rg.cols }, (_, i) => rg.c1 + i + 1)]; });
def('ROWS', 'array', { cat: K, min: 1, max: 1 }, ([a]) => isRange(a) ? a.rows : toMatrix(a).length);
def('COLUMNS', 'array', { cat: K, min: 1, max: 1 }, ([a]) => isRange(a) ? a.cols : dims(toMatrix(a))[1]);
def('ADDRESS', 'row_num, column_num, [abs_num], [a1], [sheet_text]', { cat: K, lift: true, min: 2, max: 5 }, ([r, c, a, a1, sh]) => { const rr = int(r), cc = int(c), abs = a == null ? 1 : int(a); if (rr < 1 || cc < 1) throw ERR.VALUE(); const useA1 = a1 == null || bool(a1); let s; if (useA1) s = `${abs === 1 || abs === 3 ? '$' : ''}${colToName(cc - 1)}${abs === 1 || abs === 2 ? '$' : ''}${rr}`; else s = `R${abs === 1 || abs === 2 ? rr : `[${rr}]`}C${abs === 1 || abs === 3 ? cc : `[${cc}]`}`; return sh != null ? `${quoteSheetName(str(sh))}!${s}` : s; });
def('INDIRECT', 'ref_text, [a1]', { cat: K, volatile: true, min: 1, max: 2 }, ([t], ctx) => { const rg = parseRangeName(str(t)); if (!rg) throw ERR.REF(); const sheet = rg.sheet ? ctx.wb.getSheet(rg.sheet) : ctx.sheet; if (!sheet) throw ERR.REF(); return ctx.makeRange(sheet, rg.r1, rg.c1, rg.r2, rg.c2); });
def('OFFSET', 'reference, rows, cols, [height], [width]', { cat: K, min: 3, max: 5 }, ([rg, r, c, h, w], ctx) => { if (!isRange(rg)) throw ERR.VALUE(); const dr = int(r), dc = int(c); const hh = h == null ? rg.rows : int(h), ww = w == null ? rg.cols : int(w); if (hh < 1 || ww < 1) throw ERR.REF(); const r1 = rg.r1 + dr, c1 = rg.c1 + dc; if (r1 < 0 || c1 < 0 || r1 + hh > MAX_ROWS || c1 + ww > MAX_COLS) throw ERR.REF(); return ctx.makeRange(rg.sheet, r1, c1, r1 + hh - 1, c1 + ww - 1); });
def('INDEX', 'array, row_num, [column_num]', { cat: K, min: 2, max: 4 }, ([a, r, c], ctx) => {
  const rr = r == null ? 0 : int(r), cc = c == null ? 0 : int(c);
  if (isRange(a)) {
    if (rr < 0 || cc < 0 || rr > a.rows || cc > a.cols) throw ERR.REF();
    let r1 = a.r1, r2 = a.r2, c1 = a.c1, c2 = a.c2;
    if (rr > 0) { r1 = r2 = a.r1 + rr - 1; } else if (a.rows === 1 && cc === 0 && c == null) { /* INDEX(row,n) */ }
    if (cc > 0) { c1 = c2 = a.c1 + cc - 1; }
    if (rr > 0 && cc === 0 && a.rows === 1 && a.cols > 1 && c == null) { c1 = c2 = a.c1 + rr - 1; r1 = r2 = a.r1; }
    return ctx.makeRange(a.sheet, r1, c1, r2, c2);
  }
  const m = toMatrix(a); const [nr, nc] = dims(m);
  if (rr > nr || cc > nc || rr < 0 || cc < 0) throw ERR.REF();
  if (rr === 0 && cc === 0) return m;
  if (rr > 0 && cc === 0) return nr === 1 && c == null ? m[0][rr - 1] : (nc === 1 ? m[rr - 1][0] : [m[rr - 1]]);
  if (rr === 0 && cc > 0) return m.map((row) => [row[cc - 1]]);
  return m[rr - 1][cc - 1];
});
function matchValue(lookup, arr, type) {
  // arr: 1-D array of values. type: 1 (asc, <=), 0 exact (wildcards), -1 (desc, >=)
  if (type === 0) {
    const isText = typeof lookup === 'string';
    const re = isText && /[*?~]/.test(lookup) ? wildcardToRegex(lookup) : null;
    for (let i = 0; i < arr.length; i++) {
      const v = arr[i];
      if (re) { if (typeof v === 'string' && re.test(v)) return i; continue; }
      if (typeof v === typeof lookup && (typeof v === 'string' ? v.toLowerCase() === lookup.toLowerCase() : v === lookup)) return i;
    }
    return -1;
  }
  // binary-search semantics approximated by linear scan for last value <= lookup (type 1) or >= (type -1)
  let best = -1;
  for (let i = 0; i < arr.length; i++) {
    const v = arr[i];
    if (v == null || typeof v !== typeof lookup) continue;
    const c = compareValues(typeof v === 'string' ? v.toLowerCase() : v, typeof lookup === 'string' ? lookup.toLowerCase() : lookup);
    if (type === 1) { if (c <= 0) best = i; else break; }
    else { if (c >= 0) best = i; else break; }
  }
  return best;
}
def('MATCH', 'lookup_value, lookup_array, [match_type]', { cat: K, min: 2, max: 3 }, ([lv, arr, mt]) => { lv = scalar(lv); if (isErr(lv)) throw lv; const m = toMatrix(arr); const list = m.length === 1 ? m[0] : m.map((r) => r[0]); const i = matchValue(lv, list, mt == null ? 1 : int(mt)); if (i < 0) throw ERR.NA(); return i + 1; });
def('VLOOKUP', 'lookup_value, table_array, col_index_num, [range_lookup]', { cat: K, min: 3, max: 4 }, ([lv, tbl, ci, rl]) => { lv = scalar(lv); if (isErr(lv)) throw lv; const m = toMatrix(tbl); const col = int(ci); if (col < 1 || col > dims(m)[1]) throw ERR.REF(); const approx = rl == null || bool(rl); const i = matchValue(lv, m.map((r) => r[0]), approx ? 1 : 0); if (i < 0) throw ERR.NA(); const v = m[i][col - 1]; return v == null ? 0 : v; });
def('HLOOKUP', 'lookup_value, table_array, row_index_num, [range_lookup]', { cat: K, min: 3, max: 4 }, ([lv, tbl, ri, rl]) => { lv = scalar(lv); if (isErr(lv)) throw lv; const m = toMatrix(tbl); const row = int(ri); if (row < 1 || row > m.length) throw ERR.REF(); const approx = rl == null || bool(rl); const i = matchValue(lv, m[0], approx ? 1 : 0); if (i < 0) throw ERR.NA(); const v = m[row - 1][i]; return v == null ? 0 : v; });
def('LOOKUP', 'lookup_value, lookup_vector, [result_vector]', { cat: K, min: 2, max: 3 }, ([lv, vec, res]) => { lv = scalar(lv); const m = toMatrix(vec); let list, results; if (res == null) { if (m.length >= dims(m)[1]) { list = m.map((r) => r[0]); results = m.map((r) => r[r.length - 1]); } else { list = m[0]; results = m[m.length - 1]; } } else { list = m.length === 1 ? m[0] : m.map((r) => r[0]); const rm = toMatrix(res); results = rm.length === 1 ? rm[0] : rm.map((r) => r[0]); } const i = matchValue(lv, list, 1); if (i < 0) throw ERR.NA(); return results[i] ?? 0; });
def('XMATCH', 'lookup_value, lookup_array, [match_mode], [search_mode]', { cat: K, min: 2, max: 4 }, ([lv, arr, mm, sm]) => { const i = xmatch(scalar(lv), arr, mm == null ? 0 : int(mm), sm == null ? 1 : int(sm)); if (i < 0) throw ERR.NA(); return i + 1; });
function xmatch(lv, arr, mode, search) {
  if (isErr(lv)) throw lv;
  const m = toMatrix(arr); const list = m.length === 1 ? m[0] : m.map((r) => r[0]);
  const idxs = list.map((_, i) => i); if (search < 0) idxs.reverse();
  const cmp = (v) => { if (v == null || typeof v !== typeof lv) return null; return typeof v === 'string' ? v.toLowerCase().localeCompare(lv.toLowerCase()) : v < lv ? -1 : v > lv ? 1 : 0; };
  if (mode === 0 || mode === 2) {
    const re = mode === 2 && typeof lv === 'string' ? wildcardToRegex(lv) : null;
    for (const i of idxs) { const v = list[i]; if (re) { if (typeof v === 'string' && re.test(v)) return i; } else if (cmp(v) === 0) return i; }
    return -1;
  }
  let best = -1, bestV = null;
  for (const i of idxs) { const v = list[i]; const c = cmp(v); if (c === 0) return i; if (c == null) continue; if (mode === -1 && c < 0 && (bestV == null || compareValues(v, bestV) > 0)) { best = i; bestV = v; } if (mode === 1 && c > 0 && (bestV == null || compareValues(v, bestV) < 0)) { best = i; bestV = v; } }
  return best;
}
def('XLOOKUP', 'lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode]', { cat: K, lazy: true, min: 3, max: 6 }, (nodes, ctx) => {
  const lv = scalar(ctx.eval(nodes[0])); if (isErr(lv)) return lv;
  const arr = ctx.eval(nodes[1]), ret = ctx.eval(nodes[2]);
  const mm = nodes[4] && nodes[4].type !== 'empty' ? int(scalar(ctx.eval(nodes[4]))) : 0; const sm = nodes[5] && nodes[5].type !== 'empty' ? int(scalar(ctx.eval(nodes[5]))) : 1;
  let i; try { i = xmatch(lv, arr, mm, sm); } catch (e) { if (isErr(e)) return e; throw e; }
  if (i < 0) { if (nodes[3] && nodes[3].type !== 'empty') return scalar(ctx.eval(nodes[3])); return ERR.NA(); }
  const lm = toMatrix(arr); const vertical = lm.length > 1 || dims(lm)[1] === 1;
  if (isRange(ret)) return vertical ? ctx.makeRange(ret.sheet, ret.r1 + i, ret.c1, ret.r1 + i, ret.c2) : ctx.makeRange(ret.sheet, ret.r1, ret.c1 + i, ret.r2, ret.c1 + i);
  const rm = toMatrix(ret); return vertical ? (rm[i].length === 1 ? rm[i][0] : [rm[i]]) : rm.map((r) => [r[i]]);
});
def('CHOOSE', 'index_num, value1, [value2], ...', { cat: K, lazy: true, min: 2 }, (nodes, ctx) => { const i = scalar(ctx.eval(nodes[0])); if (isErr(i)) return i; const k = int(i); if (k < 1 || k >= nodes.length) return ERR.VALUE(); return ctx.eval(nodes[k]); });
def('TRANSPOSE', 'array', { cat: K, min: 1, max: 1 }, ([a]) => { const m = toMatrix(a); const [r, c] = dims(m); return Array.from({ length: c }, (_, j) => Array.from({ length: r }, (_, i) => m[i][j])); });
def('UNIQUE', 'array, [by_col], [exactly_once]', { cat: K, min: 1, max: 3 }, ([a, bc, eo]) => {
  let m = toMatrix(a); const byCol = bc != null && bool(bc); const once = eo != null && bool(eo);
  if (byCol) m = FUNCTIONS.get('TRANSPOSE').fn([m]);
  const keyOf = (row) => JSON.stringify(row.map((v) => typeof v === 'string' ? v.toLowerCase() : v == null ? null : isErr(v) ? v.error : v));
  const counts = new Map(); for (const row of m) { const k = keyOf(row); counts.set(k, (counts.get(k) || 0) + 1); }
  const seen = new Set(); let out = [];
  for (const row of m) { const k = keyOf(row); if (seen.has(k)) continue; seen.add(k); if (once && counts.get(k) !== 1) continue; out.push(row.map((v) => v == null ? 0 : v)); }
  if (!out.length) throw ERR.CALC();
  if (byCol) out = FUNCTIONS.get('TRANSPOSE').fn([out]);
  return out;
});
function sortMatrix(m, idx, order, byCol) {
  if (byCol) m = FUNCTIONS.get('TRANSPOSE').fn([m]);
  const keys = [].concat(idx), orders = [].concat(order);
  const sorted = m.slice().sort((a, b) => { for (let k = 0; k < keys.length; k++) { const col = keys[k] - 1; const c = compareValues(a[col], b[col]) * ((orders[k] ?? orders[0] ?? 1) < 0 ? -1 : 1); if (c) return c; } return 0; });
  return byCol ? FUNCTIONS.get('TRANSPOSE').fn([sorted]) : sorted;
}
def('SORT', 'array, [sort_index], [sort_order], [by_col]', { cat: K, min: 1, max: 4 }, ([a, si, so, bc]) => sortMatrix(toMatrix(a), si == null ? 1 : flat(toMatrix(si)).map(int), so == null ? 1 : flat(toMatrix(so)).map(int), bc != null && bool(bc)));
def('SORTBY', 'array, by_array1, [sort_order1], ...', { cat: K, min: 2 }, ([a, ...rest]) => {
  const m = toMatrix(a); const keys = [];
  for (let i = 0; i < rest.length; i += 2) { const km = toMatrix(rest[i]); const list = km.length === 1 ? km[0] : km.map((r) => r[0]); if (list.length !== m.length) throw ERR.VALUE(); keys.push({ list, order: rest[i + 1] == null ? 1 : int(rest[i + 1]) }); }
  const idx = m.map((_, i) => i).sort((x, y) => { for (const k of keys) { const c = compareValues(k.list[x], k.list[y]) * (k.order < 0 ? -1 : 1); if (c) return c; } return 0; });
  return idx.map((i) => m[i]);
});
def('FILTER', 'array, include, [if_empty]', { cat: K, lazy: true, min: 2, max: 3 }, (nodes, ctx) => {
  const m = toMatrix(ctx.eval(nodes[0])); const inc = toMatrix(ctx.eval(nodes[1]));
  const [r, c] = dims(m); const [ir, ic] = dims(inc);
  let out;
  if (ir === r && ic === 1) out = m.filter((_, i) => { const v = inc[i][0]; if (isErr(v)) throw v; return bool(v); });
  else if (ir === 1 && ic === c) { const keep = inc[0].map((v) => { if (isErr(v)) throw v; return bool(v); }); out = m.map((row) => row.filter((_, j) => keep[j])); if (!keep.some(Boolean)) out = []; }
  else return ERR.VALUE();
  if (!out.length) { if (nodes[2] && nodes[2].type !== 'empty') return scalar(ctx.eval(nodes[2])); return ERR.CALC(); }
  return out;
});
def('SEQUENCE', 'rows, [columns], [start], [step]', { cat: K, min: 1, max: 4 }, ([r, c, s, st]) => { const rows = int(r), cols = c == null ? 1 : int(c), start = s == null ? 1 : num(s), step = st == null ? 1 : num(st); if (rows < 1 || cols < 1) throw ERR.VALUE(); let k = 0; return Array.from({ length: rows }, () => Array.from({ length: cols }, () => start + step * k++)); });
def('CHOOSECOLS', 'array, col_num1, ...', { cat: K, min: 2 }, ([a, ...cs]) => { const m = toMatrix(a); const w = dims(m)[1]; const idx = cs.flatMap((c) => flat(toMatrix(c)).map(int)).map((i) => i < 0 ? w + i : i - 1); if (idx.some((i) => i < 0 || i >= w)) throw ERR.VALUE(); return m.map((row) => idx.map((i) => row[i])); });
def('CHOOSEROWS', 'array, row_num1, ...', { cat: K, min: 2 }, ([a, ...rs]) => { const m = toMatrix(a); const idx = rs.flatMap((c) => flat(toMatrix(c)).map(int)).map((i) => i < 0 ? m.length + i : i - 1); if (idx.some((i) => i < 0 || i >= m.length)) throw ERR.VALUE(); return idx.map((i) => m[i]); });
def('HSTACK', 'array1, [array2], ...', { cat: K, min: 1 }, (args) => { const ms = args.map(toMatrix); const h = Math.max(...ms.map((m) => m.length)); const out = []; for (let i = 0; i < h; i++) { const row = []; for (const m of ms) { const w = dims(m)[1]; for (let j = 0; j < w; j++) row.push(m[i] ? m[i][j] : ERR.NA()); } out.push(row); } return out; });
def('VSTACK', 'array1, [array2], ...', { cat: K, min: 1 }, (args) => { const ms = args.map(toMatrix); const w = Math.max(...ms.map((m) => dims(m)[1])); const out = []; for (const m of ms) for (const row of m) { const r = row.slice(); while (r.length < w) r.push(ERR.NA()); out.push(r); } return out; });
def('TAKE', 'array, rows, [columns]', { cat: K, min: 2, max: 3 }, ([a, r, c]) => { let m = toMatrix(a); const rr = r == null ? m.length : int(r); if (rr === 0) throw ERR.CALC(); m = rr > 0 ? m.slice(0, rr) : m.slice(rr); if (c != null) { const cc = int(c); if (cc === 0) throw ERR.CALC(); m = m.map((row) => cc > 0 ? row.slice(0, cc) : row.slice(cc)); } return m; });
def('DROP', 'array, rows, [columns]', { cat: K, min: 2, max: 3 }, ([a, r, c]) => { let m = toMatrix(a); const rr = r == null ? 0 : int(r); m = rr >= 0 ? m.slice(rr) : m.slice(0, m.length + rr); if (c != null) { const cc = int(c); m = m.map((row) => cc >= 0 ? row.slice(cc) : row.slice(0, row.length + cc)); } if (!m.length || !m[0].length) throw ERR.CALC(); return m; });
def('TOCOL', 'array, [ignore], [scan_by_column]', { cat: K, min: 1, max: 3 }, ([a, ig, sbc]) => { let m = toMatrix(a); if (sbc != null && bool(sbc)) m = FUNCTIONS.get('TRANSPOSE').fn([m]); const mode = ig == null ? 0 : int(ig); const vals = flat(m).filter((v) => !((mode === 1 || mode === 3) && v == null) && !((mode === 2 || mode === 3) && isErr(v))); return vals.map((v) => [v == null ? 0 : v]); });
def('TOROW', 'array, [ignore], [scan_by_column]', { cat: K, min: 1, max: 3 }, ([a, ig, sbc]) => [FUNCTIONS.get('TOCOL').fn([a, ig, sbc]).map((r) => r[0])]);
def('WRAPROWS', 'vector, wrap_count, [pad_with]', { cat: K, min: 2, max: 3 }, ([v, w, p]) => { const list = flat(toMatrix(v)); const n = int(w); if (n < 1) throw ERR.NUM(); const out = []; for (let i = 0; i < list.length; i += n) { const row = list.slice(i, i + n); while (row.length < n) row.push(p == null ? ERR.NA() : p); out.push(row); } return out; });
def('WRAPCOLS', 'vector, wrap_count, [pad_with]', { cat: K, min: 2, max: 3 }, ([v, w, p]) => FUNCTIONS.get('TRANSPOSE').fn([FUNCTIONS.get('WRAPROWS').fn([v, w, p])]));
def('EXPAND', 'array, rows, [columns], [pad_with]', { cat: K, min: 2, max: 4 }, ([a, r, c, p]) => { const m = toMatrix(a); const [h, w] = dims(m); const rr = r == null ? h : int(r), cc = c == null ? w : int(c); if (rr < h || cc < w) throw ERR.VALUE(); const pad = p == null ? ERR.NA() : p; return Array.from({ length: rr }, (_, i) => Array.from({ length: cc }, (_, j) => (i < h && j < w) ? m[i][j] : pad)); });

// ---------------- Date & time ----------------
const D = 'Date';
const toSerial = (v) => { if (typeof v === 'string') { const d = parseDateInput(v); if (d) return d.serial; } return num(v); };
def('TODAY', '', { cat: D, volatile: true, max: 0 }, () => todaySerial());
def('NOW', '', { cat: D, volatile: true, max: 0 }, () => nowSerial());
def('DATE', 'year, month, day', { cat: D, lift: true, min: 3, max: 3 }, ([y, m, d]) => { let yy = int(y); const mm = int(m), dd = int(d); if (yy < 1900) yy += 1900; const base = dateToSerial(yy, 1, 1); const dt = new Date(Date.UTC(yy, mm - 1, dd)); const s = dateToSerial(dt.getUTCFullYear(), dt.getUTCMonth() + 1, dt.getUTCDate()); if (s < 0 || isNaN(s)) throw ERR.NUM(); return s; });
def('TIME', 'hour, minute, second', { cat: D, lift: true, min: 3, max: 3 }, ([h, m, s]) => { const t = (int(h) * 3600 + int(m) * 60 + int(s)) / 86400; if (t < 0) throw ERR.NUM(); return t - Math.floor(t); });
def('YEAR', 'serial_number', { cat: D, lift: true, min: 1, max: 1 }, ([v]) => serialToDate(toSerial(v)).y);
def('MONTH', 'serial_number', { cat: D, lift: true, min: 1, max: 1 }, ([v]) => serialToDate(toSerial(v)).m);
def('DAY', 'serial_number', { cat: D, lift: true, min: 1, max: 1 }, ([v]) => serialToDate(toSerial(v)).d);
def('HOUR', 'serial_number', { cat: D, lift: true, min: 1, max: 1 }, ([v]) => serialToDate(toSerial(v)).h);
def('MINUTE', 'serial_number', { cat: D, lift: true, min: 1, max: 1 }, ([v]) => serialToDate(toSerial(v)).mi);
def('SECOND', 'serial_number', { cat: D, lift: true, min: 1, max: 1 }, ([v]) => serialToDate(toSerial(v)).s);
def('WEEKDAY', 'serial_number, [return_type]', { cat: D, lift: true, min: 1, max: 2 }, ([v, t]) => { const wd = serialToDate(toSerial(v)).weekday; const type = t == null ? 1 : int(t); if (type === 1 || type === 17) return wd + 1; if (type === 2 || type === 11) return wd === 0 ? 7 : wd; if (type === 3) return wd === 0 ? 6 : wd - 1; if (type >= 12 && type <= 16) { const start = type - 10; return ((wd - start + 7) % 7) + 1; } throw ERR.NUM(); });
def('WEEKNUM', 'serial_number, [return_type]', { cat: D, lift: true, min: 1, max: 2 }, ([v, t]) => { const s = Math.floor(toSerial(v)); const type = t == null ? 1 : int(t); if (type === 21) return isoWeek(s); const d = serialToDate(s); const jan1 = dateToSerial(d.y, 1, 1); const jan1wd = serialToDate(jan1).weekday; const startDay = type === 1 || type === 17 ? 0 : type === 2 || type === 11 ? 1 : type - 10; const offset = (jan1wd - startDay + 7) % 7; return Math.floor((s - jan1 + offset) / 7) + 1; });
function isoWeek(serial) { const d = serialToDate(serial); const dt = new Date(Date.UTC(d.y, d.m - 1, d.d)); const day = (dt.getUTCDay() + 6) % 7; dt.setUTCDate(dt.getUTCDate() - day + 3); const firstThu = new Date(Date.UTC(dt.getUTCFullYear(), 0, 4)); return 1 + Math.round(((dt - firstThu) / 86400000 - 3 + ((firstThu.getUTCDay() + 6) % 7)) / 7); }
def('ISOWEEKNUM', 'date', { cat: D, lift: true, min: 1, max: 1 }, ([v]) => isoWeek(Math.floor(toSerial(v))));
def('DATEVALUE', 'date_text', { cat: D, lift: true, min: 1, max: 1 }, ([v]) => { const d = parseDateInput(str(v)); if (!d) throw ERR.VALUE(); return Math.floor(d.serial); });
def('TIMEVALUE', 'time_text', { cat: D, lift: true, min: 1, max: 1 }, ([v]) => { const d = parseDateInput(str(v)); if (!d) throw ERR.VALUE(); return d.serial - Math.floor(d.serial); });
def('DAYS', 'end_date, start_date', { cat: D, lift: true, min: 2, max: 2 }, ([e, s]) => Math.floor(toSerial(e)) - Math.floor(toSerial(s)));
def('DAYS360', 'start_date, end_date, [method]', { cat: D, lift: true, min: 2, max: 3 }, ([s, e, m]) => { const a = serialToDate(toSerial(s)), b = serialToDate(toSerial(e)); const eu = m != null && bool(m); let d1 = a.d, d2 = b.d; if (eu) { if (d1 === 31) d1 = 30; if (d2 === 31) d2 = 30; } else { const lastFeb = (y, mo, d) => mo === 2 && d === new Date(Date.UTC(y, 2, 0)).getUTCDate(); if (lastFeb(a.y, a.m, a.d) && lastFeb(b.y, b.m, b.d)) d2 = 30; if (lastFeb(a.y, a.m, a.d)) d1 = 30; if (d2 === 31 && d1 >= 30) d2 = 30; if (d1 === 31) d1 = 30; } return (b.y - a.y) * 360 + (b.m - a.m) * 30 + (d2 - d1); });
def('DATEDIF', 'start_date, end_date, unit', { cat: D, lift: true, min: 3, max: 3 }, ([s, e, u]) => {
  const a = serialToDate(toSerial(s)), b = serialToDate(toSerial(e)); if (b.serial < a.serial) throw ERR.NUM();
  const unit = str(u).toUpperCase();
  const months = (b.y - a.y) * 12 + (b.m - a.m) - (b.d < a.d ? 1 : 0);
  switch (unit) {
    case 'D': return Math.floor(b.serial) - Math.floor(a.serial);
    case 'M': return months; case 'Y': return Math.floor(months / 12);
    case 'YM': return months % 12;
    case 'MD': { let d = b.d - a.d; if (d < 0) { const pm = new Date(Date.UTC(b.y, b.m - 1, 0)).getUTCDate(); d += pm; } return d; }
    case 'YD': { const yrs = Math.floor(months / 12); const start = dateToSerial(a.y + yrs, a.m, Math.min(a.d, 28)) + (a.d - Math.min(a.d, 28)); return Math.floor(b.serial) - Math.floor(start); }
    default: throw ERR.NUM();
  }
});
function addMonths(serial, n) { const d = serialToDate(serial); const total = d.m - 1 + n; const y = d.y + Math.floor(total / 12), m = ((total % 12) + 12) % 12 + 1; const dim = new Date(Date.UTC(y, m, 0)).getUTCDate(); return { y, m, d: Math.min(d.d, dim), dim }; }
def('EDATE', 'start_date, months', { cat: D, lift: true, min: 2, max: 2 }, ([s, n]) => { const r = addMonths(toSerial(s), int(n)); return dateToSerial(r.y, r.m, r.d); });
def('EOMONTH', 'start_date, months', { cat: D, lift: true, min: 2, max: 2 }, ([s, n]) => { const r = addMonths(toSerial(s), int(n)); return dateToSerial(r.y, r.m, r.dim); });
function weekendSet(w) { if (w == null) return new Set([0, 6]); if (typeof w === 'string' && /^[01]{7}$/.test(w)) { const set = new Set(); for (let i = 0; i < 7; i++) if (w[i] === '1') set.add((i + 1) % 7); return set; } const k = int(w); const map = { 1: [6, 0], 2: [0, 1], 3: [1, 2], 4: [2, 3], 5: [3, 4], 6: [4, 5], 7: [5, 6], 11: [0], 12: [1], 13: [2], 14: [3], 15: [4], 16: [5], 17: [6] }; if (!map[k]) throw ERR.NUM(); return new Set(map[k]); }
function holidaySet(h) { const set = new Set(); if (h != null) each([h], (v) => { if (typeof v === 'number') set.add(Math.floor(v)); else if (typeof v === 'string' && v) { const d = parseDateInput(v); if (d) set.add(Math.floor(d.serial)); } }); return set; }
function networkdays(s, e, weekend, holidays) { let a = Math.floor(s), b = Math.floor(e); const sign = b < a ? -1 : 1; if (b < a) [a, b] = [b, a]; let n = 0; for (let d = a; d <= b; d++) { if (weekend.has((d + 6) % 7) || holidays.has(d)) continue; n++; } return n * sign; }
def('NETWORKDAYS', 'start_date, end_date, [holidays]', { cat: D, min: 2, max: 3 }, ([s, e, h]) => networkdays(toSerial(scalar(s)), toSerial(scalar(e)), weekendSet(null), holidaySet(h)));
def('NETWORKDAYS.INTL', 'start_date, end_date, [weekend], [holidays]', { cat: D, min: 2, max: 4 }, ([s, e, w, h]) => networkdays(toSerial(scalar(s)), toSerial(scalar(e)), weekendSet(scalar(w)), holidaySet(h)));
function workday(s, n, weekend, holidays) { let d = Math.floor(s); const step = n < 0 ? -1 : 1; let left = Math.abs(n); while (left > 0) { d += step; if (weekend.has((d + 6) % 7) || holidays.has(d)) continue; left--; } return d; }
def('WORKDAY', 'start_date, days, [holidays]', { cat: D, min: 2, max: 3 }, ([s, n, h]) => workday(toSerial(scalar(s)), int(scalar(n)), weekendSet(null), holidaySet(h)));
def('WORKDAY.INTL', 'start_date, days, [weekend], [holidays]', { cat: D, min: 2, max: 4 }, ([s, n, w, h]) => workday(toSerial(scalar(s)), int(scalar(n)), weekendSet(scalar(w)), holidaySet(h)));
def('YEARFRAC', 'start_date, end_date, [basis]', { cat: D, lift: true, min: 2, max: 3 }, ([s, e, b]) => {
  let a = toSerial(s), z = toSerial(e); if (a > z) [a, z] = [z, a]; const basis = b == null ? 0 : int(b);
  const A = serialToDate(a), Z = serialToDate(z);
  if (basis === 0 || basis === 4) return FUNCTIONS.get('DAYS360').fn([a, z, basis === 4]) / 360;
  if (basis === 2) return (z - a) / 360; if (basis === 3) return (z - a) / 365;
  if (basis === 1) { const leap = (y) => (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0; let ylen; if (A.y === Z.y) ylen = leap(A.y) ? 366 : 365; else { let days = 0; for (let y = A.y; y <= Z.y; y++) days += leap(y) ? 366 : 365; ylen = days / (Z.y - A.y + 1); } return (z - a) / ylen; }
  throw ERR.NUM();
});

// ---------------- Financial ----------------
const F = 'Financial';
function pmt(rate, nper, pv, fv = 0, type = 0) { if (rate === 0) return -(pv + fv) / nper; const f = Math.pow(1 + rate, nper); return -(rate * (pv * f + fv)) / ((1 + rate * type) * (f - 1)); }
def('PMT', 'rate, nper, pv, [fv], [type]', { cat: F, lift: true, min: 3, max: 5 }, ([r, n, p, f, t]) => checkNum(pmt(num(r), num(n), num(p), f == null ? 0 : num(f), t == null ? 0 : num(t))));
def('PV', 'rate, nper, pmt, [fv], [type]', { cat: F, lift: true, min: 3, max: 5 }, ([r, n, p, f, t]) => { const rate = num(r), nper = num(n), pm = num(p), fv = f == null ? 0 : num(f), type = t == null ? 0 : num(t); if (rate === 0) return -(fv + pm * nper); const fac = Math.pow(1 + rate, nper); return -(fv + pm * (1 + rate * type) * (fac - 1) / rate) / fac; });
def('FV', 'rate, nper, pmt, [pv], [type]', { cat: F, lift: true, min: 3, max: 5 }, ([r, n, p, v, t]) => { const rate = num(r), nper = num(n), pm = num(p), pv = v == null ? 0 : num(v), type = t == null ? 0 : num(t); if (rate === 0) return -(pv + pm * nper); const fac = Math.pow(1 + rate, nper); return -(pv * fac + pm * (1 + rate * type) * (fac - 1) / rate); });
def('NPER', 'rate, pmt, pv, [fv], [type]', { cat: F, lift: true, min: 3, max: 5 }, ([r, p, v, f, t]) => { const rate = num(r), pm = num(p), pv = num(v), fv = f == null ? 0 : num(f), type = t == null ? 0 : num(t); if (rate === 0) { if (pm === 0) throw ERR.NUM(); return -(pv + fv) / pm; } const a = pm * (1 + rate * type) / rate; const x = (a - fv) / (pv + a); if (x <= 0) throw ERR.NUM(); return Math.log(x) / Math.log(1 + rate); });
def('RATE', 'nper, pmt, pv, [fv], [type], [guess]', { cat: F, lift: true, min: 3, max: 6 }, ([n, p, v, f, t, g]) => {
  const nper = num(n), pm = num(p), pv = num(v), fv = f == null ? 0 : num(f), type = t == null ? 0 : num(t); let rate = g == null ? 0.1 : num(g);
  const fn = (r) => r === 0 ? pv + pm * nper + fv : pv * Math.pow(1 + r, nper) + pm * (1 + r * type) * (Math.pow(1 + r, nper) - 1) / r + fv;
  for (let i = 0; i < 100; i++) { const y = fn(rate); const dy = (fn(rate + 1e-6) - y) / 1e-6; if (dy === 0) break; const nr = rate - y / dy; if (Math.abs(nr - rate) < 1e-10) return nr; rate = nr; }
  if (Math.abs(fn(rate)) < 1e-6) return rate; throw ERR.NUM();
});
def('IPMT', 'rate, per, nper, pv, [fv], [type]', { cat: F, lift: true, min: 4, max: 6 }, ([r, per, n, v, f, t]) => {
  const rate = num(r), p = int(per), nper = num(n), pv = num(v), fv = f == null ? 0 : num(f), type = t == null ? 0 : num(t);
  if (p < 1 || p > nper) throw ERR.NUM();
  const pm = pmt(rate, nper, pv, fv, type);
  const FV = FUNCTIONS.get('FV').fn;
  let interest;
  if (p === 1) interest = type === 1 ? 0 : -pv;
  else interest = type === 1 ? FV([rate, p - 2, pm, pv, 1]) - pm : FV([rate, p - 1, pm, pv, 0]);
  return interest * rate;
});
def('PPMT', 'rate, per, nper, pv, [fv], [type]', { cat: F, lift: true, min: 4, max: 6 }, ([r, per, n, v, f, t]) => { const rate = num(r), nper = num(n), pv = num(v), fv = f == null ? 0 : num(f), type = t == null ? 0 : num(t); const pm = pmt(rate, nper, pv, fv, type); return pm - FUNCTIONS.get('IPMT').fn([rate, per, nper, pv, fv, type]); });
def('CUMIPMT', 'rate, nper, pv, start_period, end_period, type', { cat: F, lift: true, min: 6, max: 6 }, ([r, n, v, s, e, t]) => { let total = 0; for (let p = int(s); p <= int(e); p++) total += FUNCTIONS.get('IPMT').fn([r, p, n, v, 0, t]); return total; });
def('CUMPRINC', 'rate, nper, pv, start_period, end_period, type', { cat: F, lift: true, min: 6, max: 6 }, ([r, n, v, s, e, t]) => { let total = 0; for (let p = int(s); p <= int(e); p++) total += FUNCTIONS.get('PPMT').fn([r, p, n, v, 0, t]); return total; });
def('NPV', 'rate, value1, [value2], ...', { cat: F, min: 2 }, ([r, ...vals]) => { const rate = num(r); const n = nums(vals); return n.reduce((acc, v, i) => acc + v / Math.pow(1 + rate, i + 1), 0); });
function irr(cash, guess = 0.1) { let r = guess; const npv = (x) => cash.reduce((a, v, i) => a + v / Math.pow(1 + x, i), 0); for (let i = 0; i < 200; i++) { const y = npv(r); const dy = (npv(r + 1e-7) - y) / 1e-7; if (!isFinite(dy) || dy === 0) break; const nr = r - y / dy; if (Math.abs(nr - r) < 1e-10) return nr; r = nr; } if (Math.abs(npv(r)) < 1e-6) return r; throw ERR.NUM(); }
def('IRR', 'values, [guess]', { cat: F, min: 1, max: 2 }, ([v, g]) => irr(nums([v]), g == null ? 0.1 : num(g)));
def('MIRR', 'values, finance_rate, reinvest_rate', { cat: F, min: 3, max: 3 }, ([v, fr, rr]) => { const cash = nums([v]); const f = num(fr), r = num(rr); const n = cash.length; const pos = cash.map((c) => c > 0 ? c : 0), neg = cash.map((c) => c < 0 ? c : 0); const npvPos = pos.reduce((a, c, i) => a + c / Math.pow(1 + r, i), 0), npvNeg = neg.reduce((a, c, i) => a + c / Math.pow(1 + f, i), 0); if (npvNeg === 0) throw ERR.DIV0(); return Math.pow(-npvPos * Math.pow(1 + r, n) / (npvNeg * (1 + f)), 1 / (n - 1)) - 1; });
def('XNPV', 'rate, values, dates', { cat: F, min: 3, max: 3 }, ([r, v, d]) => { const rate = num(r); const vals = nums([v]), dates = nums([d]); if (vals.length !== dates.length) throw ERR.NUM(); return vals.reduce((a, c, i) => a + c / Math.pow(1 + rate, (dates[i] - dates[0]) / 365), 0); });
def('XIRR', 'values, dates, [guess]', { cat: F, min: 2, max: 3 }, ([v, d, g]) => { const vals = nums([v]), dates = nums([d]); if (vals.length !== dates.length) throw ERR.NUM(); let r = g == null ? 0.1 : num(g); const f = (x) => vals.reduce((a, c, i) => a + c / Math.pow(1 + x, (dates[i] - dates[0]) / 365), 0); for (let i = 0; i < 200; i++) { const y = f(r); const dy = (f(r + 1e-7) - y) / 1e-7; if (!isFinite(dy) || dy === 0) break; const nr = r - y / dy; if (Math.abs(nr - r) < 1e-10) return nr; r = nr; } if (Math.abs(f(r)) < 1e-6) return r; throw ERR.NUM(); });
def('SLN', 'cost, salvage, life', { cat: F, lift: true, min: 3, max: 3 }, ([c, s, l]) => { const life = num(l); if (life === 0) throw ERR.DIV0(); return (num(c) - num(s)) / life; });
def('SYD', 'cost, salvage, life, per', { cat: F, lift: true, min: 4, max: 4 }, ([c, s, l, p]) => { const life = num(l), per = num(p); if (per < 1 || per > life) throw ERR.NUM(); return (num(c) - num(s)) * (life - per + 1) * 2 / (life * (life + 1)); });
def('DB', 'cost, salvage, life, period, [month]', { cat: F, lift: true, min: 4, max: 5 }, ([c, s, l, p, m]) => { const cost = num(c), salvage = num(s), life = num(l), period = int(p), month = m == null ? 12 : int(m); const rate = Math.round((1 - Math.pow(salvage / cost, 1 / life)) * 1000) / 1000; let total = 0, dep = cost * rate * month / 12; if (period === 1) return dep; total = dep; for (let i = 2; i <= period; i++) { if (i === life + 1) dep = (cost - total) * rate * (12 - month) / 12; else dep = (cost - total) * rate; total += dep; } return dep; });
def('DDB', 'cost, salvage, life, period, [factor]', { cat: F, lift: true, min: 4, max: 5 }, ([c, s, l, p, f]) => { const cost = num(c), salvage = num(s), life = num(l), period = num(p), factor = f == null ? 2 : num(f); let total = 0, dep = 0; for (let i = 1; i <= period; i++) { dep = Math.min((cost - total) * factor / life, cost - salvage - total); if (dep < 0) dep = 0; total += dep; } return dep; });
def('EFFECT', 'nominal_rate, npery', { cat: F, lift: true, min: 2, max: 2 }, ([r, n]) => { const k = int(n); if (num(r) <= 0 || k < 1) throw ERR.NUM(); return Math.pow(1 + num(r) / k, k) - 1; });
def('NOMINAL', 'effect_rate, npery', { cat: F, lift: true, min: 2, max: 2 }, ([r, n]) => { const k = int(n); if (num(r) <= 0 || k < 1) throw ERR.NUM(); return (Math.pow(1 + num(r), 1 / k) - 1) * k; });

// ---------------- Engineering ----------------
const E = 'Engineering';
function fromBase(s, base, bits) { s = str(s).trim().toUpperCase(); if (!s) return 0; const v = parseInt(s, base); if (isNaN(v) || s.length > bits / Math.log2(base)) throw ERR.NUM(); const max = Math.pow(base, Math.round(bits / Math.log2(base))); return v >= max / 2 ? v - max : v; }
function toBase(n, base, places, bits) { n = int(n); const digits = Math.round(bits / Math.log2(base)); const lim = Math.pow(base, digits - 1) * (base / 2); if (n < -lim || n >= lim) throw ERR.NUM(); const v = n < 0 ? n + Math.pow(base, digits) : n; let s = v.toString(base).toUpperCase(); if (places != null) { const p = int(places); if (s.length > p) throw ERR.NUM(); s = s.padStart(p, '0'); } return s; }
def('DEC2BIN', 'number, [places]', { cat: E, lift: true, min: 1, max: 2 }, ([n, p]) => toBase(n, 2, p, 10));
def('DEC2OCT', 'number, [places]', { cat: E, lift: true, min: 1, max: 2 }, ([n, p]) => toBase(n, 8, p, 30));
def('DEC2HEX', 'number, [places]', { cat: E, lift: true, min: 1, max: 2 }, ([n, p]) => toBase(n, 16, p, 40));
def('BIN2DEC', 'number', { cat: E, lift: true, min: 1, max: 1 }, ([s]) => fromBase(s, 2, 10));
def('OCT2DEC', 'number', { cat: E, lift: true, min: 1, max: 1 }, ([s]) => fromBase(s, 8, 30));
def('HEX2DEC', 'number', { cat: E, lift: true, min: 1, max: 1 }, ([s]) => fromBase(s, 16, 40));
def('BIN2HEX', 'number, [places]', { cat: E, lift: true, min: 1, max: 2 }, ([s, p]) => toBase(fromBase(s, 2, 10), 16, p, 40));
def('HEX2BIN', 'number, [places]', { cat: E, lift: true, min: 1, max: 2 }, ([s, p]) => toBase(fromBase(s, 16, 40), 2, p, 10));
def('BITAND', 'number1, number2', { cat: E, lift: true, min: 2, max: 2 }, ([a, b]) => Number(BigInt(int(a)) & BigInt(int(b))));
def('BITOR', 'number1, number2', { cat: E, lift: true, min: 2, max: 2 }, ([a, b]) => Number(BigInt(int(a)) | BigInt(int(b))));
def('BITXOR', 'number1, number2', { cat: E, lift: true, min: 2, max: 2 }, ([a, b]) => Number(BigInt(int(a)) ^ BigInt(int(b))));
def('BITLSHIFT', 'number, shift_amount', { cat: E, lift: true, min: 2, max: 2 }, ([a, b]) => { const s = int(b); return s >= 0 ? int(a) * Math.pow(2, s) : Math.floor(int(a) / Math.pow(2, -s)); });
def('BITRSHIFT', 'number, shift_amount', { cat: E, lift: true, min: 2, max: 2 }, ([a, b]) => { const s = int(b); return s >= 0 ? Math.floor(int(a) / Math.pow(2, s)) : int(a) * Math.pow(2, -s); });
def('DELTA', 'number1, [number2]', { cat: E, lift: true, min: 1, max: 2 }, ([a, b]) => num(a) === (b == null ? 0 : num(b)) ? 1 : 0);
def('GESTEP', 'number, [step]', { cat: E, lift: true, min: 1, max: 2 }, ([a, b]) => num(a) >= (b == null ? 0 : num(b)) ? 1 : 0);
def('ERF', 'lower_limit, [upper_limit]', { cat: E, lift: true, min: 1, max: 2 }, ([a, b]) => b == null ? erf(num(a)) : erf(num(b)) - erf(num(a)));
def('ERFC', 'x', { cat: E, lift: true, min: 1, max: 1 }, ([a]) => 1 - erf(num(a)));
const UNITS = { m: ['len', 1], km: ['len', 1000], cm: ['len', 0.01], mm: ['len', 0.001], mi: ['len', 1609.344], yd: ['len', 0.9144], ft: ['len', 0.3048], in: ['len', 0.0254], Nmi: ['len', 1852], g: ['mass', 1], kg: ['mass', 1000], lbm: ['mass', 453.59237], ozm: ['mass', 28.349523125], stone: ['mass', 6350.29318], sec: ['time', 1], s: ['time', 1], min: ['time', 60], hr: ['time', 3600], day: ['time', 86400], yr: ['time', 31557600], l: ['vol', 1], L: ['vol', 1], ml: ['vol', 0.001], gal: ['vol', 3.785411784], qt: ['vol', 0.946352946], pt: ['vol', 0.473176473], cup: ['vol', 0.2365882365], oz: ['vol', 0.0295735295625], tsp: ['vol', 0.00492892159375], tbs: ['vol', 0.01478676478125], m3: ['vol', 1000], 'm^3': ['vol', 1000], ft3: ['vol', 28.316846592], 'ft^3': ['vol', 28.316846592], J: ['energy', 1], kJ: ['energy', 1000], cal: ['energy', 4.1868], kcal: ['energy', 4186.8], Wh: ['energy', 3600], kWh: ['energy', 3600000], BTU: ['energy', 1055.05585262], eV: ['energy', 1.602176634e-19], W: ['power', 1], kW: ['power', 1000], HP: ['power', 745.69987158227], Pa: ['press', 1], atm: ['press', 101325], mmHg: ['press', 133.322], psi: ['press', 6894.757293168], bar: ['press', 100000], N: ['force', 1], lbf: ['force', 4.4482216152605], dyn: ['force', 1e-5], 'm/s': ['speed', 1], 'm/h': ['speed', 1 / 3600], 'km/h': ['speed', 1 / 3.6], mph: ['speed', 0.44704], kn: ['speed', 0.514444], bit: ['info', 1], byte: ['info', 8], ha: ['area', 10000], 'm2': ['area', 1], 'm^2': ['area', 1], 'km2': ['area', 1e6], 'ft2': ['area', 0.09290304], 'mi2': ['area', 2589988.110336], ac: ['area', 4046.8564224], uk_acre: ['area', 4046.8564224] };
def('CONVERT', 'number, from_unit, to_unit', { cat: E, lift: true, min: 3, max: 3 }, ([n, f, t]) => {
  const v = num(n); const from = str(f), to = str(t);
  const temp = { C: (x) => x + 273.15, cel: (x) => x + 273.15, F: (x) => (x - 32) * 5 / 9 + 273.15, fah: (x) => (x - 32) * 5 / 9 + 273.15, K: (x) => x, kel: (x) => x };
  const tempInv = { C: (k) => k - 273.15, cel: (k) => k - 273.15, F: (k) => (k - 273.15) * 9 / 5 + 32, fah: (k) => (k - 273.15) * 9 / 5 + 32, K: (k) => k, kel: (k) => k };
  if (temp[from] && temp[to]) return tempInv[to](temp[from](v));
  const prefixes = { Y: 1e24, Z: 1e21, E: 1e18, P: 1e15, T: 1e12, G: 1e9, M: 1e6, k: 1e3, h: 1e2, da: 10, d: 0.1, c: 0.01, m: 0.001, u: 1e-6, n: 1e-9, p: 1e-12, f: 1e-15, a: 1e-18, Ki: 1024, Mi: 1024 ** 2, Gi: 1024 ** 3, Ti: 1024 ** 4 };
  const resolve = (u) => { if (UNITS[u]) return UNITS[u]; for (const [p, mult] of Object.entries(prefixes)) if (u.startsWith(p) && UNITS[u.slice(p.length)]) { const base = UNITS[u.slice(p.length)]; return [base[0], base[1] * mult]; } throw ERR.NA(); };
  const a = resolve(from), b = resolve(to); if (a[0] !== b[0]) throw ERR.NA(); return v * a[1] / b[1];
});

// ---------------- Database-style helpers ----------------
def('DSUM', 'database, field, criteria', { cat: 'Database', min: 3, max: 3 }, (args) => dbAgg(args, 'SUM'));
def('DAVERAGE', 'database, field, criteria', { cat: 'Database', min: 3, max: 3 }, (args) => dbAgg(args, 'AVERAGE'));
def('DCOUNT', 'database, field, criteria', { cat: 'Database', min: 3, max: 3 }, (args) => dbAgg(args, 'COUNT'));
def('DCOUNTA', 'database, field, criteria', { cat: 'Database', min: 3, max: 3 }, (args) => dbAgg(args, 'COUNTA'));
def('DMAX', 'database, field, criteria', { cat: 'Database', min: 3, max: 3 }, (args) => dbAgg(args, 'MAX'));
def('DMIN', 'database, field, criteria', { cat: 'Database', min: 3, max: 3 }, (args) => dbAgg(args, 'MIN'));
def('DGET', 'database, field, criteria', { cat: 'Database', min: 3, max: 3 }, (args) => { const rows = dbRows(args); if (rows.length !== 1) throw rows.length ? ERR.NUM() : ERR.VALUE(); return rows[0]; });
function dbRows([db, field, crit]) {
  const m = toMatrix(db), cm = toMatrix(crit); const header = m[0].map((h) => str(h).toLowerCase());
  let col; const f = scalar(field); if (typeof f === 'number') col = int(f) - 1; else col = header.indexOf(str(f).toLowerCase()); if (col < 0 || col >= header.length) throw ERR.VALUE();
  const critCols = cm[0].map((h) => header.indexOf(str(h).toLowerCase()));
  const tests = cm.slice(1).map((row) => row.map((c, j) => c == null || c === '' ? null : { col: critCols[j], test: criteria(c) }).filter(Boolean));
  const out = [];
  for (let i = 1; i < m.length; i++) { const row = m[i]; if (tests.some((t) => t.every((x) => x.col >= 0 && x.test(row[x.col])))) out.push(row[col]); }
  return out;
}
function dbAgg(args, name) { const rows = dbRows(args); return FUNCTIONS.get(name).fn([[rows]]); }

// ---------------- Array broadcasting (used by the engine for lifted functions and operators) ----------------
export function broadcast(values, fn) {
  // values: array of scalars or 2D arrays; returns scalar if all scalar, otherwise 2D array
  const mats = values.map((v) => isRange(v) ? v.toArray() : v);
  if (!mats.some(isArr)) return fn(mats);
  let rows = 1, cols = 1;
  for (const m of mats) if (isArr(m)) { const [r, c] = dims(m); rows = r === 1 ? rows : rows === 1 ? r : Math.min(rows, r); cols = c === 1 ? cols : cols === 1 ? c : Math.min(cols, c); }
  // if sizes conflict (neither 1) result is the larger one with #N/A in the gaps
  let maxR = 1, maxC = 1; for (const m of mats) if (isArr(m)) { const [r, c] = dims(m); maxR = Math.max(maxR, r); maxC = Math.max(maxC, c); }
  const out = [];
  for (let i = 0; i < maxR; i++) {
    const row = [];
    for (let j = 0; j < maxC; j++) {
      const picked = []; let na = false;
      for (const m of mats) {
        if (!isArr(m)) { picked.push(m); continue; }
        const [r, c] = dims(m); const ii = r === 1 ? 0 : i, jj = c === 1 ? 0 : j;
        if (ii >= r || jj >= c) { na = true; break; }
        picked.push(m[ii][jj]);
      }
      row.push(na ? ERR.NA() : fn(picked));
    }
    out.push(row);
  }
  return out;
}

export function functionNames() { return [...FUNCTIONS.keys()].sort(); }
