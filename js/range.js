// A reference to a rectangular block of cells on a sheet (values are read lazily).
import { MAX_ROWS, MAX_COLS } from './util.js';

export class RangeRef {
  constructor(sheet, r1, c1, r2, c2) {
    this.sheet = sheet; this.r1 = r1; this.c1 = c1; this.r2 = r2; this.c2 = c2;
  }
  get rows() { return this.r2 - this.r1 + 1; }
  get cols() { return this.c2 - this.c1 + 1; }
  get isCell() { return this.r1 === this.r2 && this.c1 === this.c2; }
  value(i = 0, j = 0) { return this.sheet.getValue(this.r1 + i, this.c1 + j); }
  // Effective bounds, clipped to the used area of the sheet (so A:A does not iterate a million rows)
  clipped() {
    let r2 = this.r2, c2 = this.c2;
    if (this.rows >= MAX_ROWS || this.cols >= MAX_COLS || this.rows * this.cols > 4_000_000) {
      const u = this.sheet.usedBounds();
      if (this.rows >= MAX_ROWS) r2 = Math.max(this.r1, Math.min(this.r2, u.r2));
      if (this.cols >= MAX_COLS) c2 = Math.max(this.c1, Math.min(this.c2, u.c2));
      if (r2 === this.r2 && c2 === this.c2) { r2 = Math.min(this.r2, Math.max(this.r1, u.r2)); c2 = Math.min(this.c2, Math.max(this.c1, u.c2)); }
    }
    return { r1: this.r1, c1: this.c1, r2, c2 };
  }
  toArray() {
    const { r1, c1, r2, c2 } = this.clipped();
    const out = [];
    for (let r = r1; r <= r2; r++) { const row = []; for (let c = c1; c <= c2; c++) row.push(this.sheet.getValue(r, c)); out.push(row); }
    return out;
  }
  forEach(fn) {
    const { r1, c1, r2, c2 } = this.clipped();
    for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) fn(this.sheet.getValue(r, c), r, c);
  }
  sub(r1, c1, r2, c2) { return new RangeRef(this.sheet, this.r1 + r1, this.c1 + c1, this.r1 + r2, this.c1 + c2); }
  toString() { return `${this.sheet.name}!R${this.r1 + 1}C${this.c1 + 1}:R${this.r2 + 1}C${this.c2 + 1}`; }
}

export const isRange = (v) => v instanceof RangeRef;
export const isArr = (v) => Array.isArray(v);
