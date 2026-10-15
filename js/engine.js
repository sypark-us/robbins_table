// Formula evaluation engine: dependency graph, incremental recalculation, spilling.
import { parse, FormulaError, ERR, isErr, ParseError } from './formula.js';
import { RangeRef, isRange, isArr } from './range.js';
import { FUNCTIONS, num, str, broadcast } from './functions.js';
import { formatGeneral, parseInput } from './format.js';
import { key as cellKey, compareValues, MAX_ROWS, MAX_COLS } from './util.js';

const DATE_FUNC_FORMATS = { TODAY: 'm/d/yyyy', DATE: 'm/d/yyyy', EDATE: 'm/d/yyyy', EOMONTH: 'm/d/yyyy', DATEVALUE: 'm/d/yyyy', WORKDAY: 'm/d/yyyy', 'WORKDAY.INTL': 'm/d/yyyy', NOW: 'm/d/yyyy h:mm', TIME: 'h:mm AM/PM', TIMEVALUE: 'h:mm AM/PM' };

export class Engine {
  constructor(wb) {
    this.wb = wb;
    this.precedents = new Map(); // depId -> [{sheetId, r1,c1,r2,c2}]
    this.cellDeps = new Map();   // sheetId -> Map(key -> Set(depId))
    this.rangeDeps = new Map();  // sheetId -> [{r1,c1,r2,c2,depId}]
    this.volatile = new Set();
    this.state = new Map();      // depId -> 'dirty' | 'evaluating' | 'clean'
    this.changed = new Set();    // depIds of cells whose input changed since last recalc
    this.pendingSpillDependents = [];
    this.blockedSpills = new Set();
    this.calcMode = 'auto';
    this.suspended = 0;
    this.lastRecalcCount = 0;
  }

  depId(sheet, r, c) { return `${sheet.id}:${r},${c}`; }
  parseDepId(id) { const i = id.indexOf(':'); const sheet = this.wb.sheetById(id.slice(0, i)); const [r, c] = id.slice(i + 1).split(',').map(Number); return { sheet, r, c }; }

  // ---- called by the model ----
  // A cell's content changed (value or formula). Registers/unregisters the formula and marks dependents dirty.
  cellChanged(sheet, r, c) {
    const id = this.depId(sheet, r, c);
    const cell = sheet.cells.get(cellKey(r, c));
    this.changed.add(id);
    this.spillNeighbourChanged(sheet);
    if (!cell || !cell.f) this.blockedSpills.delete(id);
    if (cell && cell.f) {
      try { cell.ast = parse(cell.f); cell.parseError = null; } catch (e) { cell.ast = null; cell.parseError = e; }
      this.state.set(id, 'dirty');
      if (cell.ast && hasVolatile(cell.ast)) this.volatile.add(id); else this.volatile.delete(id);
    } else {
      if (this.state.has(id)) { this.removePrecedents(id); this.state.delete(id); this.volatile.delete(id); }
      if (cell && cell.spill) this.clearSpill(sheet, cell, r, c);
    }
  }

  // Rebuild all state from the workbook (after structural changes or load).
  rebuildAll() {
    this.precedents.clear(); this.cellDeps.clear(); this.rangeDeps.clear(); this.volatile.clear(); this.state.clear(); this.changed.clear(); this.blockedSpills.clear();
    for (const sheet of this.wb.sheets) {
      for (const [k, cell] of sheet.cells) {
        if (cell.spillOf && !cell.f) { cell.spillOf = null; cell.v = null; }
        if (cell.spill) cell.spill = null;
        if (cell.f) {
          const [r, c] = k.split(',').map(Number);
          const id = this.depId(sheet, r, c);
          try { cell.ast = parse(cell.f); cell.parseError = null; } catch (e) { cell.ast = null; cell.parseError = e; }
          this.state.set(id, 'dirty');
          if (cell.ast && hasVolatile(cell.ast)) this.volatile.add(id);
        }
      }
      for (const [k, cell] of [...sheet.cells]) if (!cell.f && cell.v == null && !cell.s && !cell.note && !cell.hyperlink) sheet.cells.delete(k);
    }
    this.recalcAll();
  }

  recalcAll() {
    for (const id of this.state.keys()) this.state.set(id, 'dirty');
    this.changed.clear();
    this.evaluateDirty();
  }

  // Incremental recalculation of everything affected by changes since the last call.
  recalc() {
    if (this.suspended > 0) return;
    const dirty = new Set();
    for (const id of this.state.keys()) if (this.state.get(id) === 'dirty') dirty.add(id);
    const queue = [];
    for (const id of this.changed) { const { sheet, r, c } = this.parseDepId(id); if (!sheet) continue; for (const d of this.dependentsOf(sheet.id, r, c)) queue.push(d); const cell = sheet.cells.get(cellKey(r, c)); if (cell && cell.spill) this.spillCells(cell).forEach(([rr, cc]) => { for (const d of this.dependentsOf(sheet.id, rr, cc)) queue.push(d); }); }
    for (const id of this.volatile) queue.push(id);
    for (const id of dirty) queue.push(id);
    while (queue.length) {
      const id = queue.pop();
      if (dirty.has(id)) continue;
      dirty.add(id);
      const { sheet, r, c } = this.parseDepId(id);
      if (!sheet) continue;
      for (const d of this.dependentsOf(sheet.id, r, c)) if (!dirty.has(d)) queue.push(d);
      const cell = sheet.cells.get(cellKey(r, c));
      if (cell && cell.spill) this.spillCells(cell).forEach(([rr, cc]) => { for (const d of this.dependentsOf(sheet.id, rr, cc)) if (!dirty.has(d)) queue.push(d); });
    }
    for (const id of dirty) if (this.state.has(id)) this.state.set(id, 'dirty');
    this.changed.clear();
    this.evaluateDirty();
  }

  evaluateDirty() {
    let count = 0;
    for (let round = 0; round < 10; round++) {
      let any = false;
      for (const [id, st] of this.state) {
        if (st !== 'dirty') continue;
        any = true;
        const { sheet, r, c } = this.parseDepId(id);
        if (!sheet) { this.state.delete(id); continue; }
        this.evalCell(sheet, r, c); count++;
      }
      // spilled values may have changed cells that other formulas depend on
      if (this.pendingSpillDependents.length) {
        const pend = this.pendingSpillDependents; this.pendingSpillDependents = [];
        for (const id of pend) if (this.state.has(id) && this.state.get(id) === 'clean') { this.state.set(id, 'dirty'); any = true; }
        continue;
      }
      if (!any) break;
    }
    this.lastRecalcCount = count;
  }

  isDirty(sheet, r, c) { const st = this.state.get(this.depId(sheet, r, c)); return st === 'dirty' || st === 'evaluating'; }
  markDirty(sheet, r, c) { const id = this.depId(sheet, r, c); if (this.state.has(id)) { this.state.set(id, 'dirty'); this.changed.add(id); } }
  // a cell inside (or near) a spill area changed: blocked spills may now be free, free spills may now be blocked
  spillNeighbourChanged(sheet) { for (const id of this.blockedSpills) { const { sheet: sh } = this.parseDepId(id); if (sh === sheet && this.state.has(id)) { this.state.set(id, 'dirty'); this.changed.add(id); } } }

  // ---- dependency index ----
  removePrecedents(id) {
    const prev = this.precedents.get(id);
    if (!prev) return;
    for (const p of prev) {
      if (p.r1 === p.r2 && p.c1 === p.c2) { const m = this.cellDeps.get(p.sheetId); const set = m && m.get(cellKey(p.r1, p.c1)); if (set) { set.delete(id); if (!set.size) m.delete(cellKey(p.r1, p.c1)); } }
      else { const list = this.rangeDeps.get(p.sheetId); if (list) { const i = list.findIndex((x) => x.depId === id && x.r1 === p.r1 && x.c1 === p.c1 && x.r2 === p.r2 && x.c2 === p.c2); if (i >= 0) list.splice(i, 1); } }
    }
    this.precedents.delete(id);
  }
  addPrecedents(id, list) {
    this.precedents.set(id, list);
    for (const p of list) {
      if (p.r1 === p.r2 && p.c1 === p.c2) {
        let m = this.cellDeps.get(p.sheetId); if (!m) { m = new Map(); this.cellDeps.set(p.sheetId, m); }
        const k = cellKey(p.r1, p.c1); let set = m.get(k); if (!set) { set = new Set(); m.set(k, set); } set.add(id);
      } else {
        let list2 = this.rangeDeps.get(p.sheetId); if (!list2) { list2 = []; this.rangeDeps.set(p.sheetId, list2); }
        list2.push({ r1: p.r1, c1: p.c1, r2: p.r2, c2: p.c2, depId: id });
      }
    }
  }
  dependentsOf(sheetId, r, c) {
    const out = [];
    const m = this.cellDeps.get(sheetId); const set = m && m.get(cellKey(r, c)); if (set) out.push(...set);
    const list = this.rangeDeps.get(sheetId);
    if (list) for (const rg of list) if (r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2) out.push(rg.depId);
    return out;
  }
  // All formula cells that reference anything on `sheet` within the rectangle
  dependentsOfRange(sheetId, r1, c1, r2, c2) {
    const out = new Set();
    const m = this.cellDeps.get(sheetId);
    if (m) { if ((r2 - r1 + 1) * (c2 - c1 + 1) <= m.size) { for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) { const s = m.get(cellKey(r, c)); if (s) s.forEach((d) => out.add(d)); } } else for (const [k, s] of m) { const [r, c] = k.split(',').map(Number); if (r >= r1 && r <= r2 && c >= c1 && c <= c2) s.forEach((d) => out.add(d)); } }
    const list = this.rangeDeps.get(sheetId);
    if (list) for (const rg of list) if (!(rg.r2 < r1 || r2 < rg.r1 || rg.c2 < c1 || c2 < rg.c1)) out.add(rg.depId);
    return out;
  }
  precedentsOf(sheet, r, c) { return this.precedents.get(this.depId(sheet, r, c)) || []; }

  // ---- evaluation ----
  evalCell(sheet, r, c) {
    const id = this.depId(sheet, r, c);
    const st = this.state.get(id);
    const cell = sheet.cells.get(cellKey(r, c));
    if (!cell || !cell.f) { this.state.delete(id); return cell ? cell.v : null; }
    if (st === 'clean') return cell.v;
    if (st === 'evaluating') return ERR.CIRC();
    this.state.set(id, 'evaluating');
    this.removePrecedents(id);
    const recorded = [];
    const ctx = this.makeContext(sheet, r, c, recorded);
    let result;
    if (!cell.ast) result = ERR.ERROR(cell.parseError ? cell.parseError.message : 'parse error');
    else {
      try { result = this.evalNode(cell.ast, ctx); } catch (e) { result = isErr(e) ? e : ERR.ERROR(e && e.message); if (!isErr(e) && typeof console !== 'undefined' && e && e.stack && !(e instanceof ParseError)) console.warn('formula error', cell.f, e); }
    }
    result = this.deref(result);
    this.addPrecedents(id, recorded);
    // number format inference for General cells
    cell.autoFmt = null;
    if (typeof result === 'number' && cell.ast) { const fmt = this.inferFormat(cell.ast, ctx); if (fmt) cell.autoFmt = fmt; }
    if (isArr(result)) {
      if (result.length === 1 && result[0].length === 1) { this.clearSpill(sheet, cell, r, c); cell.v = normalizeScalar(result[0][0]); }
      else this.applySpill(sheet, cell, r, c, result);
    } else { this.clearSpill(sheet, cell, r, c); cell.v = normalizeScalar(result); }
    this.state.set(id, 'clean');
    return cell.v;
  }

  deref(v) {
    if (isRange(v)) {
      if (v.isCell) return v.value();
      const a = v.toArray();
      return a;
    }
    return v;
  }

  makeContext(sheet, row, col, recorded, scope = null) {
    const engine = this;
    const ctx = {
      sheet, row, col, wb: this.wb, engine, scope,
      eval: (node) => engine.evalNode(node, ctx),
      evalWithScope: (node, sc) => engine.evalNode(node, engine.makeContext(sheet, row, col, recorded, sc)),
      makeRange: (sh, r1, c1, r2, c2) => { recorded.push({ sheetId: sh.id, r1, c1, r2, c2 }); return new RangeRef(sh, r1, c1, r2, c2); },
    };
    return ctx;
  }

  evalNode(node, ctx) {
    switch (node.type) {
      case 'num': return node.v;
      case 'str': return node.v;
      case 'bool': return node.v;
      case 'err': return new FormulaError(node.v);
      case 'empty': return null;
      case 'array': return node.rows.map((row) => row.map((el) => el.type === 'err' ? new FormulaError(el.v) : el.v));
      case 'ref': {
        const sheet = node.sheet != null ? this.wb.getSheet(node.sheet) : ctx.sheet;
        if (!sheet) return ERR.REF();
        return ctx.makeRange(sheet, node.r1, node.c1, node.r2, node.c2);
      }
      case 'name': {
        const up = node.name.toUpperCase();
        if (ctx.scope && up in ctx.scope) return ctx.scope[up];
        const def = this.wb.getName(node.name, ctx.sheet);
        if (!def) return ERR.NAME(node.name);
        try { const ast = def.ast || (def.ast = parse(def.ref.startsWith('=') ? def.ref : '=' + def.ref)); return this.evalNode(ast, ctx); } catch (e) { return isErr(e) ? e : ERR.NAME(node.name); }
      }
      case 'range': {
        const a = this.evalNode(node.left, ctx), b = this.evalNode(node.right, ctx);
        if (isErr(a)) return a; if (isErr(b)) return b;
        if (!isRange(a) || !isRange(b) || a.sheet !== b.sheet) return ERR.VALUE('range operator needs references');
        return ctx.makeRange(a.sheet, Math.min(a.r1, b.r1), Math.min(a.c1, b.c1), Math.max(a.r2, b.r2), Math.max(a.c2, b.c2));
      }
      case 'un': {
        const v = this.deref(this.evalNode(node.arg, ctx));
        return broadcast([v], ([x]) => { if (isErr(x)) return x; try { return -num(x); } catch (e) { return e; } });
      }
      case 'pct': {
        const v = this.deref(this.evalNode(node.arg, ctx));
        return broadcast([v], ([x]) => { if (isErr(x)) return x; try { return num(x) / 100; } catch (e) { return e; } });
      }
      case 'bin': {
        const l = this.deref(this.evalNode(node.left, ctx));
        if (isErr(l) && !isArr(l)) return l;
        const r = this.deref(this.evalNode(node.right, ctx));
        if (isErr(r) && !isArr(r)) return r;
        return broadcast([l, r], ([a, b]) => binop(node.op, a, b));
      }
      case 'func': return this.callFunction(node, ctx);
      default: return ERR.ERROR('bad node');
    }
  }

  callFunction(node, ctx) {
    const def = FUNCTIONS.get(node.name);
    if (!def) return ERR.NAME(node.name);
    if (node.args.length < def.minArgs || node.args.length > def.maxArgs) {
      // allow trailing empties
      const real = node.args.filter((a) => a.type !== 'empty').length;
      if (real < def.minArgs || node.args.length > def.maxArgs) return ERR.VALUE(`${node.name}: wrong number of arguments`);
    }
    if (def.lazy) {
      try { return def.fn(node.args, ctx); } catch (e) { if (isErr(e)) return e; throw e; }
    }
    const args = node.args.map((a) => a.type === 'empty' ? null : this.evalNode(a, ctx));
    if (!def.errors) for (const a of args) if (isErr(a)) return a;
    if (def.lift) {
      const vals = args.map((a) => isRange(a) ? (a.isCell ? a.value() : a.toArray()) : a);
      if (vals.some(isArr)) return broadcast(vals, (picked) => { if (!def.errors) for (const p of picked) if (isErr(p)) return p; return safeCall(def, picked, ctx); });
      return safeCall(def, vals, ctx);
    }
    return safeCall(def, args, ctx);
  }

  inferFormat(ast, ctx) {
    if (ast.type === 'func') return DATE_FUNC_FORMATS[ast.name] || null;
    if (ast.type === 'ref' && ast.r1 === ast.r2 && ast.c1 === ast.c2) {
      const sheet = ast.sheet != null ? this.wb.getSheet(ast.sheet) : ctx.sheet;
      if (!sheet) return null;
      const st = sheet.getStyle(ast.r1, ast.c1);
      if (st.numFmt && st.numFmt !== 'General') return st.numFmt;
      const cell = sheet.cells.get(cellKey(ast.r1, ast.c1));
      return cell && cell.autoFmt ? cell.autoFmt : null;
    }
    if (ast.type === 'bin' && (ast.op === '+' || ast.op === '-')) return this.inferFormat(ast.left, ctx) || this.inferFormat(ast.right, ctx);
    if (ast.type === 'un') return this.inferFormat(ast.arg, ctx);
    return null;
  }

  // ---- spilling ----
  spillCells(cell) {
    const out = [];
    if (!cell.spill) return out;
    const s = cell.spill;
    for (let r = s.r1; r <= s.r2; r++) for (let c = s.c1; c <= s.c2; c++) if (!(r === s.r1 && c === s.c1)) out.push([r, c]);
    return out;
  }
  clearSpill(sheet, cell, r, c) {
    if (!cell.spill) return;
    const originKey = cellKey(r, c);
    for (const [rr, cc] of this.spillCells(cell)) {
      const k = cellKey(rr, cc);
      const t = sheet.cells.get(k);
      if (t && t.spillOf === originKey) { t.spillOf = null; t.v = null; if (!t.s && !t.note && !t.hyperlink) sheet.cells.delete(k); sheet.invalidateBounds(); for (const d of this.dependentsOf(sheet.id, rr, cc)) this.pendingSpillDependents.push(d); }
    }
    cell.spill = null;
  }
  applySpill(sheet, cell, r, c, matrix) {
    const rows = matrix.length, cols = matrix[0].length;
    const originKey = cellKey(r, c);
    const r2 = r + rows - 1, c2 = c + cols - 1;
    // check target area is free
    let blocked = r2 >= MAX_ROWS || c2 >= MAX_COLS;
    if (!blocked) for (let rr = r; rr <= r2 && !blocked; rr++) for (let cc = c; cc <= c2; cc++) {
      if (rr === r && cc === c) continue;
      const t = sheet.cells.get(cellKey(rr, cc));
      if (t && ((t.v != null && t.spillOf !== originKey) || t.f)) { blocked = true; break; }
    }
    if (blocked) { this.clearSpill(sheet, cell, r, c); cell.v = ERR.SPILL(); this.blockedSpills.add(this.depId(sheet, r, c)); return; }
    this.blockedSpills.delete(this.depId(sheet, r, c));
    // clear parts of the old spill that are outside the new one
    if (cell.spill) for (const [rr, cc] of this.spillCells(cell)) if (rr > r2 || cc > c2) { const k = cellKey(rr, cc); const t = sheet.cells.get(k); if (t && t.spillOf === originKey) { t.spillOf = null; t.v = null; if (!t.s && !t.note && !t.hyperlink) sheet.cells.delete(k); for (const d of this.dependentsOf(sheet.id, rr, cc)) this.pendingSpillDependents.push(d); } }
    cell.spill = { r1: r, c1: c, r2, c2 };
    cell.v = normalizeScalar(matrix[0][0]);
    for (let i = 0; i < rows; i++) for (let j = 0; j < cols; j++) {
      if (i === 0 && j === 0) continue;
      const rr = r + i, cc = c + j, k = cellKey(rr, cc);
      let t = sheet.cells.get(k);
      if (!t) { t = { v: null, f: null, s: 0 }; sheet.cells.set(k, t); }
      const nv = normalizeScalar(matrix[i][j]);
      const changed = !sameValue(t.v, nv) || t.spillOf !== originKey;
      t.v = nv; t.spillOf = originKey;
      if (changed) for (const d of this.dependentsOf(sheet.id, rr, cc)) this.pendingSpillDependents.push(d);
    }
    sheet.invalidateBounds();
  }
}

function safeCall(def, args, ctx) {
  try {
    const v = def.fn(args, ctx);
    if (typeof v === 'number' && (!isFinite(v) || isNaN(v))) return ERR.NUM();
    return v === undefined ? null : v;
  } catch (e) { if (isErr(e)) return e; if (e instanceof RangeError) return ERR.NUM(); throw e; }
}

function normalizeScalar(v) {
  if (isRange(v)) return v.isCell ? normalizeScalar(v.value()) : ERR.VALUE();
  if (isArr(v)) return normalizeScalar(v[0] ? v[0][0] : null);
  if (typeof v === 'number' && (!isFinite(v) || isNaN(v))) return ERR.NUM();
  if (v === undefined) return null;
  return v;
}
function sameValue(a, b) { if (isErr(a) && isErr(b)) return a.error === b.error; return a === b; }

function hasVolatile(ast) {
  let v = false;
  (function walk(n) { if (!n || v) return; if (n.type === 'func') { const d = FUNCTIONS.get(n.name); if (d && d.volatile) { v = true; return; } n.args.forEach(walk); } else if (n.type === 'bin' || n.type === 'range') { walk(n.left); walk(n.right); } else if (n.type === 'un' || n.type === 'pct') walk(n.arg); })(ast);
  return v;
}

export function binop(op, a, b) {
  if (isErr(a)) return a; if (isErr(b)) return b;
  try {
    switch (op) {
      case '+': return checkNum(num(a) + num(b));
      case '-': return checkNum(num(a) - num(b));
      case '*': return checkNum(num(a) * num(b));
      case '/': { const d = num(b); if (d === 0) return ERR.DIV0(); return checkNum(num(a) / d); }
      case '^': { const x = num(a), y = num(b); if (x === 0 && y < 0) return ERR.DIV0(); return checkNum(Math.pow(x, y)); }
      case '&': return str(a) + str(b);
      case '=': return compare(a, b) === 0;
      case '<>': return compare(a, b) !== 0;
      case '<': return compare(a, b) < 0;
      case '>': return compare(a, b) > 0;
      case '<=': return compare(a, b) <= 0;
      case '>=': return compare(a, b) >= 0;
    }
  } catch (e) { if (isErr(e)) return e; throw e; }
  return ERR.VALUE();
}
function checkNum(v) { return (!isFinite(v) || isNaN(v)) ? ERR.NUM() : v; }
export function compare(a, b) {
  // blanks compare as 0 / "" depending on the other side
  if (a == null && b == null) return 0;
  if (a == null) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
  if (b == null) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
  if (typeof a === 'string' && typeof b === 'string') { const x = a.toLowerCase(), y = b.toLowerCase(); return x < y ? -1 : x > y ? 1 : 0; }
  return compareValues(a, b);
}

export { parse, isErr, FormulaError };
