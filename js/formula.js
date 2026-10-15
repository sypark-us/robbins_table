// Formula tokenizer, parser and reference rewriting.
import { MAX_ROWS, MAX_COLS, nameToCol, colToName, quoteSheetName } from './util.js';

export const ERRORS = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#SPILL!', '#CIRC!', '#ERROR!', '#CALC!'];

export class FormulaError {
  constructor(code, detail) { this.error = code; this.detail = detail; }
  toString() { return this.error; }
}
export const ERR = {
  NULL: () => new FormulaError('#NULL!'), DIV0: () => new FormulaError('#DIV/0!'), VALUE: (d) => new FormulaError('#VALUE!', d),
  REF: () => new FormulaError('#REF!'), NAME: (d) => new FormulaError('#NAME?', d), NUM: () => new FormulaError('#NUM!'),
  NA: () => new FormulaError('#N/A'), SPILL: () => new FormulaError('#SPILL!'), CIRC: () => new FormulaError('#CIRC!'),
  ERROR: (d) => new FormulaError('#ERROR!', d), CALC: () => new FormulaError('#CALC!'),
};
export const isErr = (v) => v instanceof FormulaError;

// ---------------- Tokenizer ----------------
// Token: { t: 'num'|'str'|'bool'|'err'|'ref'|'ident'|'func'|'op'|'ws'|'sheet', v, s (source text), pos }
// ref token fields: sheet, r1, c1, r2, c2, absR1, absC1, absR2, absC2, kind ('cell'|'range'|'col'|'row')

const CELL_RE = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?![A-Za-z0-9_.(])/;
const COLRANGE_RE = /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![A-Za-z0-9_.(])/;
const ROWRANGE_RE = /^(\$?)(\d{1,7}):(\$?)(\d{1,7})(?![A-Za-z0-9_.(])/;
const IDENT_RE = /^[A-Za-z_\\][A-Za-z0-9_.\\]*/;
const NUM_RE = /^(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?/;

function validCell(c, r) { return c >= 0 && c < MAX_COLS && r >= 0 && r < MAX_ROWS; }

export function tokenize(src) {
  const tokens = [];
  let i = 0;
  const n = src.length;
  let pendingSheet = null, pendingSheetPos = -1;
  const push = (tok) => { tokens.push(tok); };
  while (i < n) {
    const ch = src[i];
    const rest = src.slice(i, i + 64);
    let m;
    if (ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r') {
      let j = i; while (j < n && /\s/.test(src[j])) j++;
      push({ t: 'ws', s: src.slice(i, j), pos: i }); i = j; continue;
    }
    if (ch === '"') {
      let j = i + 1, out = '';
      while (j < n) { if (src[j] === '"') { if (src[j + 1] === '"') { out += '"'; j += 2; continue; } break; } out += src[j++]; }
      push({ t: 'str', v: out, s: src.slice(i, j + 1), pos: i }); i = j + 1; continue;
    }
    if (ch === "'") {
      // quoted sheet name
      let j = i + 1, out = '';
      while (j < n) { if (src[j] === "'") { if (src[j + 1] === "'") { out += "'"; j += 2; continue; } break; } out += src[j++]; }
      if (src[j + 1] === '!') { pendingSheet = out; pendingSheetPos = i; i = j + 2; continue; }
      push({ t: 'err', v: '#ERROR!', s: src.slice(i, j + 1), pos: i }); i = j + 1; continue;
    }
    if (ch === '#') {
      const e = ERRORS.find((er) => rest.toUpperCase().startsWith(er));
      if (e) { push({ t: 'err', v: e, s: src.slice(i, i + e.length), pos: i }); i += e.length; continue; }
      push({ t: 'op', v: '#', s: '#', pos: i }); i++; continue;
    }
    if ((m = NUM_RE.exec(rest)) && !/^\d+:\d+/.test(rest)) {
      push({ t: 'num', v: parseFloat(m[0]), s: m[0], pos: i }); i += m[0].length; continue;
    }
    if ((m = ROWRANGE_RE.exec(rest))) {
      const r1 = +m[2] - 1, r2 = +m[4] - 1;
      if (r1 < MAX_ROWS && r2 < MAX_ROWS) {
        push(refTok({ sheet: pendingSheet, r1: Math.min(r1, r2), r2: Math.max(r1, r2), c1: 0, c2: MAX_COLS - 1, absR1: m[1] === '$', absR2: m[3] === '$', absC1: true, absC2: true, kind: 'row' }, src.slice(pendingSheet != null ? pendingSheetPos : i, i + m[0].length), pendingSheet != null ? pendingSheetPos : i));
        pendingSheet = null; i += m[0].length; continue;
      }
    }
    if ((m = COLRANGE_RE.exec(rest))) {
      const c1 = nameToCol(m[2].toUpperCase()), c2 = nameToCol(m[4].toUpperCase());
      if (c1 < MAX_COLS && c2 < MAX_COLS) {
        push(refTok({ sheet: pendingSheet, r1: 0, r2: MAX_ROWS - 1, c1: Math.min(c1, c2), c2: Math.max(c1, c2), absC1: m[1] === '$', absC2: m[3] === '$', absR1: true, absR2: true, kind: 'col' }, src.slice(pendingSheet != null ? pendingSheetPos : i, i + m[0].length), pendingSheet != null ? pendingSheetPos : i));
        pendingSheet = null; i += m[0].length; continue;
      }
    }
    if ((m = CELL_RE.exec(rest))) {
      const c = nameToCol(m[2].toUpperCase()), r = +m[4] - 1;
      if (validCell(c, r)) {
        const start = pendingSheet != null ? pendingSheetPos : i;
        let end = i + m[0].length;
        const ref = { sheet: pendingSheet, r1: r, c1: c, r2: r, c2: c, absR1: m[3] === '$', absC1: m[1] === '$', absR2: m[3] === '$', absC2: m[1] === '$', kind: 'cell' };
        // range A1:B2 ?
        const m2 = /^:(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?![A-Za-z0-9_.(])/.exec(src.slice(end, end + 32));
        if (m2) {
          const c2 = nameToCol(m2[2].toUpperCase()), r2 = +m2[4] - 1;
          if (validCell(c2, r2)) {
            ref.kind = 'range';
            const absR2 = m2[3] === '$', absC2 = m2[1] === '$';
            // normalise corners while keeping abs flags attached to the right corner
            const corners = [{ r, c, absR: ref.absR1, absC: ref.absC1 }, { r: r2, c: c2, absR: absR2, absC: absC2 }];
            ref.r1 = Math.min(r, r2); ref.r2 = Math.max(r, r2); ref.c1 = Math.min(c, c2); ref.c2 = Math.max(c, c2);
            const top = corners.find((k) => k.r === ref.r1), bottom = corners.find((k) => k.r === ref.r2 && k !== top) || corners.find((k) => k.r === ref.r2);
            const left = corners.find((k) => k.c === ref.c1), right = corners.find((k) => k.c === ref.c2 && k !== left) || corners.find((k) => k.c === ref.c2);
            ref.absR1 = top.absR; ref.absR2 = bottom.absR; ref.absC1 = left.absC; ref.absC2 = right.absC;
            end += m2[0].length;
          }
        }
        push(refTok(ref, src.slice(start, end), start));
        pendingSheet = null; i = end; continue;
      }
    }
    if ((m = IDENT_RE.exec(rest))) {
      const word = m[0];
      const after = src.slice(i + word.length);
      if (after[0] === '!') { pendingSheet = word; pendingSheetPos = i; i += word.length + 1; continue; }
      if (/^\s*\(/.test(after)) { push({ t: 'func', v: word.toUpperCase(), s: word, pos: i }); i += word.length; continue; }
      const up = word.toUpperCase();
      if (up === 'TRUE' || up === 'FALSE') { push({ t: 'bool', v: up === 'TRUE', s: word, pos: i }); i += word.length; continue; }
      push({ t: 'ident', v: word, s: word, pos: i, sheet: pendingSheet }); pendingSheet = null; i += word.length; continue;
    }
    if (ch === '<' && src[i + 1] === '=') { push({ t: 'op', v: '<=', s: '<=', pos: i }); i += 2; continue; }
    if (ch === '>' && src[i + 1] === '=') { push({ t: 'op', v: '>=', s: '>=', pos: i }); i += 2; continue; }
    if (ch === '<' && src[i + 1] === '>') { push({ t: 'op', v: '<>', s: '<>', pos: i }); i += 2; continue; }
    if ('+-*/^&=<>%:,;(){}'.includes(ch)) { push({ t: 'op', v: ch, s: ch, pos: i }); i++; continue; }
    push({ t: 'bad', v: ch, s: ch, pos: i }); i++;
  }
  return tokens;
}

function refTok(ref, s, pos) { return { t: 'ref', ...ref, s, pos }; }

export function refToString(ref, opts = {}) {
  const sheet = ref.sheet != null && !opts.noSheet ? quoteSheetName(ref.sheet) + '!' : '';
  if (ref.kind === 'col') return `${sheet}${ref.absC1 ? '$' : ''}${colToName(ref.c1)}:${ref.absC2 ? '$' : ''}${colToName(ref.c2)}`;
  if (ref.kind === 'row') return `${sheet}${ref.absR1 ? '$' : ''}${ref.r1 + 1}:${ref.absR2 ? '$' : ''}${ref.r2 + 1}`;
  const a = `${ref.absC1 ? '$' : ''}${colToName(ref.c1)}${ref.absR1 ? '$' : ''}${ref.r1 + 1}`;
  if (ref.kind === 'cell' || (ref.r1 === ref.r2 && ref.c1 === ref.c2 && ref.kind !== 'range')) return sheet + a;
  const b = `${ref.absC2 ? '$' : ''}${colToName(ref.c2)}${ref.absR2 ? '$' : ''}${ref.r2 + 1}`;
  return `${sheet}${a}:${b}`;
}

// ---------------- Parser ----------------
// AST nodes:
//  {type:'num', v} {type:'str', v} {type:'bool', v} {type:'err', v} {type:'empty'}
//  {type:'ref', sheet, r1,c1,r2,c2, kind}
//  {type:'name', name, sheet}
//  {type:'func', name, args:[...]}
//  {type:'bin', op, left, right}  {type:'un', op, arg}  {type:'pct', arg}
//  {type:'array', rows:[[...]]}
//  {type:'range', left, right}   (dynamic range operator between non-literal refs)

export class ParseError extends Error {}

export function parse(src) {
  const text = src.startsWith('=') ? src.slice(1) : src;
  const toks = tokenize(text).filter((t) => t.t !== 'ws');
  let p = 0;
  const peek = () => toks[p];
  const next = () => toks[p++];
  const isOp = (v) => peek() && peek().t === 'op' && peek().v === v;
  const expect = (v) => { if (!isOp(v)) throw new ParseError(`Expected '${v}'`); next(); };

  function parseExpr() { return parseCompare(); }
  function parseCompare() {
    let left = parseConcat();
    while (peek() && peek().t === 'op' && ['=', '<>', '<', '>', '<=', '>='].includes(peek().v)) { const op = next().v; left = { type: 'bin', op, left, right: parseConcat() }; }
    return left;
  }
  function parseConcat() {
    let left = parseAdd();
    while (isOp('&')) { next(); left = { type: 'bin', op: '&', left, right: parseAdd() }; }
    return left;
  }
  function parseAdd() {
    let left = parseMul();
    while (isOp('+') || isOp('-')) { const op = next().v; left = { type: 'bin', op, left, right: parseMul() }; }
    return left;
  }
  function parseMul() {
    let left = parsePow();
    while (isOp('*') || isOp('/')) { const op = next().v; left = { type: 'bin', op, left, right: parsePow() }; }
    return left;
  }
  function parsePow() {
    let left = parseUnary();
    while (isOp('^')) { next(); left = { type: 'bin', op: '^', left, right: parseUnary() }; }
    return left;
  }
  function parseUnary() {
    if (isOp('-')) { next(); return { type: 'un', op: '-', arg: parseUnary() }; }
    if (isOp('+')) { next(); return parseUnary(); }
    return parsePostfix();
  }
  function parsePostfix() {
    let e = parseRangeOp();
    while (isOp('%')) { next(); e = { type: 'pct', arg: e }; }
    return e;
  }
  function parseRangeOp() {
    let left = parsePrimary();
    while (isOp(':')) { next(); const right = parsePrimary(); left = { type: 'range', left, right }; }
    return left;
  }
  function parsePrimary() {
    const t = peek();
    if (!t) throw new ParseError('Unexpected end of formula');
    if (t.t === 'num') { next(); return { type: 'num', v: t.v }; }
    if (t.t === 'str') { next(); return { type: 'str', v: t.v }; }
    if (t.t === 'bool') { next(); return { type: 'bool', v: t.v }; }
    if (t.t === 'err') { next(); return { type: 'err', v: t.v }; }
    if (t.t === 'ref') { next(); return { type: 'ref', sheet: t.sheet, r1: t.r1, c1: t.c1, r2: t.r2, c2: t.c2, kind: t.kind, absR1: t.absR1, absC1: t.absC1, absR2: t.absR2, absC2: t.absC2 }; }
    if (t.t === 'ident') { next(); return { type: 'name', name: t.v, sheet: t.sheet }; }
    if (t.t === 'func') {
      next(); expect('(');
      const args = [];
      if (isOp(')')) { next(); return { type: 'func', name: t.v, args }; }
      for (;;) {
        if (isOp(',') || isOp(';')) { args.push({ type: 'empty' }); next(); continue; }
        if (isOp(')')) { args.push({ type: 'empty' }); next(); break; }
        args.push(parseExpr());
        if (isOp(',') || isOp(';')) { next(); if (isOp(')')) { args.push({ type: 'empty' }); next(); break; } continue; }
        expect(')'); break;
      }
      return { type: 'func', name: t.v, args };
    }
    if (t.t === 'op' && t.v === '(') { next(); const e = parseExpr(); expect(')'); return e; }
    if (t.t === 'op' && t.v === '{') {
      next();
      const rows = [[]];
      for (;;) {
        const v = peek();
        let el;
        if (v && v.t === 'num') { next(); el = { type: 'num', v: v.v }; }
        else if (v && v.t === 'str') { next(); el = { type: 'str', v: v.v }; }
        else if (v && v.t === 'bool') { next(); el = { type: 'bool', v: v.v }; }
        else if (v && v.t === 'err') { next(); el = { type: 'err', v: v.v }; }
        else if (isOp('-')) { next(); const nn = next(); if (!nn || nn.t !== 'num') throw new ParseError('Bad array constant'); el = { type: 'num', v: -nn.v }; }
        else throw new ParseError('Bad array constant');
        rows[rows.length - 1].push(el);
        if (isOp(',')) { next(); continue; }
        if (isOp(';')) { next(); rows.push([]); continue; }
        expect('}'); break;
      }
      return { type: 'array', rows };
    }
    throw new ParseError(`Unexpected token '${t.s}'`);
  }
  const ast = parseExpr();
  if (p < toks.length) throw new ParseError(`Unexpected '${toks[p].s}'`);
  return ast;
}

// Collect all static references in a formula text (for highlighting / dependencies before evaluation).
export function extractRefs(text) {
  const src = text.startsWith('=') ? text.slice(1) : text;
  const offset = text.length - src.length;
  return tokenize(src).filter((t) => t.t === 'ref').map((t) => ({ sheet: t.sheet, r1: t.r1, c1: t.c1, r2: t.r2, c2: t.c2, kind: t.kind, start: t.pos + offset, end: t.pos + t.s.length + offset, text: t.s }));
}

// Walk an AST, calling fn on each node (pre-order).
export function walk(ast, fn) {
  if (!ast) return;
  fn(ast);
  switch (ast.type) {
    case 'bin': walk(ast.left, fn); walk(ast.right, fn); break;
    case 'un': case 'pct': walk(ast.arg, fn); break;
    case 'func': ast.args.forEach((a) => walk(a, fn)); break;
    case 'range': walk(ast.left, fn); walk(ast.right, fn); break;
    default: break;
  }
}

// ---------------- Reference rewriting ----------------

// Generic: rewrite every ref token in formula text via mapper(ref) -> ref | null (null => #REF!).
export function rewriteRefs(text, mapper) {
  if (!text || text[0] !== '=') return text;
  const toks = tokenize(text.slice(1));
  let out = '=';
  for (const t of toks) {
    if (t.t !== 'ref') { out += t.s; continue; }
    const r = mapper({ ...t });
    out += r ? refToString(r) : '#REF!';
  }
  return out;
}

// Copy semantics: shift relative parts of every reference by (dr, dc).
export function shiftFormula(text, dr, dc) {
  if (!dr && !dc) return text;
  return rewriteRefs(text, (ref) => {
    const r = { ...ref };
    if (ref.kind !== 'col') { if (!r.absR1) r.r1 += dr; if (!r.absR2) r.r2 += dr; }
    if (ref.kind !== 'row') { if (!r.absC1) r.c1 += dc; if (!r.absC2) r.c2 += dc; }
    if (r.r1 < 0 || r.r2 >= MAX_ROWS || r.c1 < 0 || r.c2 >= MAX_COLS) return null;
    if (r.r1 > r.r2) [r.r1, r.r2] = [r.r2, r.r1];
    if (r.c1 > r.c2) [r.c1, r.c2] = [r.c2, r.c1];
    return r;
  });
}

// Structural change: rows/cols inserted (count>0) or deleted (count<0) at index on sheet `targetSheet`.
// `formulaSheet` is the sheet the formula lives on (unqualified refs point there).
export function adjustFormulaForInsertDelete(text, formulaSheet, targetSheet, axis, index, count) {
  return rewriteRefs(text, (ref) => {
    const refSheet = ref.sheet != null ? ref.sheet : formulaSheet;
    if (refSheet.toLowerCase() !== targetSheet.toLowerCase()) return ref;
    const r = { ...ref };
    const isRow = axis === 'row';
    let a = isRow ? r.r1 : r.c1, b = isRow ? r.r2 : r.c2;
    const max = isRow ? MAX_ROWS - 1 : MAX_COLS - 1;
    const whole = isRow ? ref.kind === 'col' : ref.kind === 'row';
    if (whole) return ref;
    if (count > 0) {
      if (a >= index) a += count;
      if (b >= index) b += count;
      if (a > max || b > max) { if (a > max) return null; b = max; }
    } else {
      const delStart = index, delEnd = index - count - 1; // inclusive deleted span
      if (a >= delStart && b <= delEnd) return null; // fully deleted
      if (a > delEnd) a += count; else if (a >= delStart) a = delStart;
      if (b > delEnd) b += count; else if (b >= delStart) b = delStart - 1;
      if (b < a) return null;
    }
    if (isRow) { r.r1 = a; r.r2 = b; } else { r.c1 = a; r.c2 = b; }
    return r;
  });
}

// Cut/paste (move) semantics: refs fully inside the moved block follow it.
export function adjustFormulaForMove(text, formulaSheet, srcSheet, src, dr, dc, dstSheet) {
  return rewriteRefs(text, (ref) => {
    const refSheet = ref.sheet != null ? ref.sheet : formulaSheet;
    if (refSheet.toLowerCase() !== srcSheet.toLowerCase()) return ref;
    if (ref.kind === 'col' || ref.kind === 'row') return ref;
    if (ref.r1 >= src.r1 && ref.r2 <= src.r2 && ref.c1 >= src.c1 && ref.c2 <= src.c2) {
      const r = { ...ref, r1: ref.r1 + dr, r2: ref.r2 + dr, c1: ref.c1 + dc, c2: ref.c2 + dc };
      if (dstSheet && dstSheet.toLowerCase() !== refSheet.toLowerCase()) r.sheet = dstSheet;
      else if (dstSheet && ref.sheet != null) r.sheet = dstSheet;
      if (r.r1 < 0 || r.c1 < 0 || r.r2 >= MAX_ROWS || r.c2 >= MAX_COLS) return null;
      return r;
    }
    return ref;
  });
}

// Rename a sheet inside formulas.
export function renameSheetInFormula(text, oldName, newName) {
  return rewriteRefs(text, (ref) => ref.sheet != null && ref.sheet.toLowerCase() === oldName.toLowerCase() ? { ...ref, sheet: newName } : ref);
}

// Does the formula reference a sheet that is being deleted? Replace with #REF!.
export function removeSheetInFormula(text, name) {
  return rewriteRefs(text, (ref) => ref.sheet != null && ref.sheet.toLowerCase() === name.toLowerCase() ? null : ref);
}

// Convert between A1 and R1C1-like relative forms is not needed; but we need "same formula, relative to another cell" for CF/validation:
export function relativeFormula(text, fromR, fromC, toR, toC) { return shiftFormula(text, toR - fromR, toC - fromC); }

export function isFormula(v) { return typeof v === 'string' && v.length > 1 && v[0] === '='; }
