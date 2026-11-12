/**
 * Runtime values used by the formula evaluator, coercion rules, and helpers shared by the function library.
 *
 * Function implementations throw `Fail` (via the helpers below) to return an error value early; the
 * evaluator catches it and turns it into the error result, so library code reads linearly.
 */
import { CellError, ERR, type Scalar, type Matrix, type ErrorCode } from '../model/types';
import type { Sheet, Workbook } from '../model/workbook';
import type { Node } from './parser';
import { parseDateTime } from '../model/dates';
import { parseNumberInput } from '../model/numfmt';
import { MAX_ROWS, MAX_COLS } from '../model/address';

export interface RefV { k: 'ref'; sheet: Sheet; r1: number; c1: number; r2: number; c2: number; sheets?: Sheet[] }
export interface MultiRefV { k: 'refs'; areas: RefV[] }
export interface LambdaV { k: 'lambda'; params: string[]; body: Node; scope: Map<string, Value> | null }

export type Value = Scalar | Matrix | RefV | MultiRefV | LambdaV;

export class Fail extends Error {
  constructor(public err: CellError) { super(err.code); }
}

export function fail(e: CellError | ErrorCode): never { throw new Fail(typeof e === 'string' ? new CellError(e) : e); }

export const isRef = (v: Value): v is RefV => !!v && typeof v === 'object' && (v as RefV).k === 'ref';
export const isRefs = (v: Value): v is MultiRefV => !!v && typeof v === 'object' && (v as MultiRefV).k === 'refs';
export const isLambda = (v: Value): v is LambdaV => !!v && typeof v === 'object' && (v as LambdaV).k === 'lambda';
export const isMatrix = (v: Value): v is Matrix => Array.isArray(v);
export const isErr = (v: unknown): v is CellError => v instanceof CellError;

/** Evaluation context passed to functions. */
export interface Ctx {
  book: Workbook;
  sheet: Sheet;
  r: number;
  c: number;
  /** Reads a cell's current value (recalculating it first if it is dirty). */
  cell(sheet: Sheet, r: number, c: number): Scalar;
  /** Evaluates a node in this context (used by lazy functions). */
  evalNode(n: Node, scope?: Map<string, Value> | null): Value;
  /** Calls a LAMBDA value. */
  callLambda(fn: LambdaV, args: Value[]): Value;
  scope: Map<string, Value> | null;
  /** Marks the current formula as volatile (recalculated on every recalc). */
  volatile(): void;
  /** Used-range clipping for whole-column/row references. */
  usedBounds(sheet: Sheet): { r2: number; c2: number };
  /** Resolves a sheet/name reference text to a reference value (INDIRECT). */
  resolveRef(text: string, a1: boolean): Value;
  now: Date;
  random(): number;
}

// --------------------------------------------------------------------------------------------
// Matrices

export function dims(v: Value): { rows: number; cols: number } {
  if (isMatrix(v)) return { rows: v.length, cols: v[0]?.length ?? 0 };
  if (isRef(v)) return { rows: v.r2 - v.r1 + 1, cols: v.c2 - v.c1 + 1 };
  return { rows: 1, cols: 1 };
}

/** Converts a reference (clipped to the used range for whole rows/columns) or scalar to a matrix. */
export function toMatrix(v: Value, ctx: Ctx, clip = true): Matrix {
  if (isMatrix(v)) return v;
  if (isRef(v)) {
    let { r2, c2 } = v;
    if (clip) {
      const b = ctx.usedBounds(v.sheet);
      // Only open-ended references shrink: A1:A10 must stay ten rows, or arrays built from it
      // come out the wrong size and pair up with the wrong cells.
      if (v.r2 >= MAX_ROWS - 1) r2 = Math.min(r2, Math.max(v.r1, b.r2));
      if (v.c2 >= MAX_COLS - 1) c2 = Math.min(c2, Math.max(v.c1, b.c2));
    }
    const out: Matrix = [];
    for (let r = v.r1; r <= r2; r++) {
      const row: Scalar[] = [];
      for (let c = v.c1; c <= c2; c++) row.push(ctx.cell(v.sheet, r, c));
      out.push(row);
    }
    return out;
  }
  if (isRefs(v)) return fail(ERR.value);
  if (isLambda(v)) return fail(ERR.calc);
  return [[v]];
}

/** Single value from a value: 1x1 ranges/arrays collapse; larger take the top-left (used by @ and scalar contexts). */
export function toScalar(v: Value, ctx: Ctx): Scalar {
  if (isMatrix(v)) return v[0]?.[0] ?? null;
  if (isRef(v)) return ctx.cell(v.sheet, v.r1, v.c1);
  if (isRefs(v)) return fail(ERR.value);
  if (isLambda(v)) return fail(ERR.calc);
  return v;
}

/** Implicit intersection with the formula cell (legacy semantics for @). */
export function implicitIntersect(v: Value, ctx: Ctx): Scalar {
  if (isRef(v)) {
    if (v.r1 === v.r2 && v.c1 === v.c2) return ctx.cell(v.sheet, v.r1, v.c1);
    if (v.c1 === v.c2 && ctx.r >= v.r1 && ctx.r <= v.r2) return ctx.cell(v.sheet, ctx.r, v.c1);
    if (v.r1 === v.r2 && ctx.c >= v.c1 && ctx.c <= v.c2) return ctx.cell(v.sheet, v.r1, ctx.c);
    return fail(ERR.value);
  }
  return toScalar(v, ctx);
}

/** All values from arguments, flattened row by row. */
export function* iterValues(v: Value, ctx: Ctx): Generator<{ v: Scalar; fromRef: boolean }> {
  if (isRef(v)) {
    const sheets = v.sheets || [v.sheet];
    for (const sh of sheets) {
      const b = ctx.usedBounds(sh);
      const r2 = Math.min(v.r2, b.r2), c2 = Math.min(v.c2, b.c2);
      // Sparse iteration for big ranges
      const area = (r2 - v.r1 + 1) * (c2 - v.c1 + 1);
      if (area > 4096 && area > sh.cells.size * 4) {
        const keys: [number, number][] = [];
        sh.forEachIn({ r1: v.r1, c1: v.c1, r2, c2 }, (r, c) => keys.push([r, c]));
        keys.sort((a, b) => a[0] - b[0] || a[1] - b[1]);
        for (const [r, c] of keys) yield { v: ctx.cell(sh, r, c), fromRef: true };
      } else {
        for (let r = v.r1; r <= r2; r++) for (let c = v.c1; c <= c2; c++) yield { v: ctx.cell(sh, r, c), fromRef: true };
      }
    }
    return;
  }
  if (isRefs(v)) { for (const a of v.areas) yield* iterValues(a, ctx); return; }
  if (isMatrix(v)) { for (const row of v) for (const x of row) yield { v: x, fromRef: true }; return; }
  if (isLambda(v)) return fail(ERR.calc);
  yield { v, fromRef: false };
}

// --------------------------------------------------------------------------------------------
// Coercion

export function parseNumeric(s: string): number | null {
  const t = s.trim();
  if (!t) return null;
  const n = parseNumberInput(t);
  if (n) return n.value;
  const d = parseDateTime(t);
  if (d) return d.serial;
  return null;
}

/** Number coercion for direct arguments: text that looks like a number/date converts, else #VALUE!. */
export function num(v: Scalar): number {
  if (typeof v === 'number') return v;
  if (v == null) return 0;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (v instanceof CellError) return fail(v);
  const n = parseNumeric(v);
  if (n == null) return fail(ERR.value);
  return n;
}

export function str(v: Scalar): string {
  if (v == null) return '';
  if (typeof v === 'string') return v;
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (v instanceof CellError) return fail(v);
  return numToText(v);
}

export function numToText(n: number): string {
  if (Number.isInteger(n) && Math.abs(n) < 1e21) return String(n);
  const s = String(+n.toPrecision(15));
  if (s.includes('e')) { const [m, e] = s.split('e'); return `${m}E${+e < 0 ? '-' : '+'}${String(Math.abs(+e)).padStart(2, '0')}`; }
  return s;
}

export function bool(v: Scalar): boolean {
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (v == null) return false;
  if (v instanceof CellError) return fail(v);
  const u = v.toUpperCase();
  if (u === 'TRUE') return true;
  if (u === 'FALSE') return false;
  return fail(ERR.value);
}

export function int(v: Scalar): number { return Math.trunc(num(v)); }

/** Throws if the scalar is an error. */
export function check(v: Scalar): Scalar { if (v instanceof CellError) fail(v); return v; }

/**
 * Numbers for aggregate functions (SUM, AVERAGE...): numbers in references count, text and booleans in
 * references are skipped; direct arguments are coerced (text numbers, booleans).
 */
export function numbers(args: Value[], ctx: Ctx, opts: { countText0?: boolean; includeBoolsInRefs?: boolean; skipErrors?: boolean } = {}): number[] {
  const out: number[] = [];
  for (const a of args) {
    for (const { v, fromRef } of iterValues(a, ctx)) {
      if (typeof v === 'number') out.push(v);
      else if (v instanceof CellError) { if (!opts.skipErrors) fail(v); }
      else if (fromRef) {
        if (typeof v === 'boolean' && opts.includeBoolsInRefs) out.push(v ? 1 : 0);
        else if (typeof v === 'string' && opts.countText0) out.push(0);
        else if (typeof v === 'boolean' && opts.countText0) out.push(v ? 1 : 0);
      } else if (v != null) out.push(num(v));
    }
  }
  return out;
}

/** Compares two scalars with spreadsheet ordering: numbers < text < booleans; text case-insensitive. */
export function compare(a: Scalar, b: Scalar): number {
  const rank = (x: Scalar) => (x == null ? 0 : typeof x === 'number' ? 1 : typeof x === 'string' ? 2 : typeof x === 'boolean' ? 3 : 4);
  if (a == null && b == null) return 0;
  if (a == null) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
  if (b == null) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
  const ra = rank(a), rb = rank(b);
  if (ra !== rb) return ra - rb;
  if (typeof a === 'number') return a < (b as number) ? -1 : a > (b as number) ? 1 : 0;
  if (typeof a === 'string') { const x = a.toLowerCase(), y = (b as string).toLowerCase(); return x < y ? -1 : x > y ? 1 : 0; }
  if (typeof a === 'boolean') return a === b ? 0 : a ? 1 : -1;
  return 0;
}

export function equals(a: Scalar, b: Scalar): boolean { return compare(a, b) === 0; }

/** Wildcard pattern (* ? ~) to RegExp. */
export function wildcard(pattern: string, whole = true): RegExp {
  let s = '';
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i];
    if (ch === '~' && i + 1 < pattern.length) { s += escapeRe(pattern[++i]); continue; }
    if (ch === '*') s += '[\\s\\S]*';
    else if (ch === '?') s += '[\\s\\S]';
    else s += escapeRe(ch);
  }
  return new RegExp(whole ? `^${s}$` : s, 'i');
}

export function escapeRe(s: string): string { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

/**
 * Criteria matcher for SUMIF/COUNTIF/AVERAGEIF/...IFS and database functions:
 * ">5", "<>x", "=", "", "a*", 5, TRUE.
 */
export function criteria(c: Scalar): (v: Scalar) => boolean {
  if (typeof c === 'number') return (v) => (typeof v === 'number' ? v === c : typeof v === 'string' ? parseNumeric(v) === c : false);
  if (typeof c === 'boolean') return (v) => v === c;
  if (c == null) return (v) => v === 0 || v == null || v === '';
  if (c instanceof CellError) return (v) => v instanceof CellError && v.code === c.code;
  const m = /^(<=|>=|<>|=|<|>)?([\s\S]*)$/.exec(c)!;
  const op = m[1] || '=';
  const rhs = m[2];
  const n = rhs.trim() === '' ? null : parseNumeric(rhs);
  const err = /^#/.test(rhs) ? rhs.toUpperCase() : null;
  const bl = /^(TRUE|FALSE)$/i.test(rhs) ? rhs.toUpperCase() === 'TRUE' : null;
  if (op === '=' || op === '<>') {
    let test: (v: Scalar) => boolean;
    if (rhs === '') test = (v) => v == null || v === '';
    else if (n != null) test = (v) => (typeof v === 'number' ? v === n : typeof v === 'string' ? parseNumeric(v) === n && v.trim() !== '' : false);
    else if (bl != null) test = (v) => v === bl;
    else if (err) test = (v) => v instanceof CellError && v.code === err;
    else if (/[*?~]/.test(rhs)) { const re = wildcard(rhs); test = (v) => typeof v === 'string' && re.test(v); }
    else { const low = rhs.toLowerCase(); test = (v) => typeof v === 'string' && v.toLowerCase() === low; }
    if (op === '<>') {
      if (rhs === '') return (v) => !(v == null || v === '');
      return (v) => !test(v);
    }
    return test;
  }
  if (n != null) {
    return (v) => {
      if (typeof v !== 'number') return false;
      switch (op) { case '<': return v < n; case '<=': return v <= n; case '>': return v > n; default: return v >= n; }
    };
  }
  const low = rhs.toLowerCase();
  return (v) => {
    if (typeof v !== 'string') return false;
    const x = v.toLowerCase();
    switch (op) { case '<': return x < low; case '<=': return x <= low; case '>': return x > low; default: return x >= low; }
  };
}

/** Builds a result matrix by mapping f over equal-shaped inputs (broadcasting 1x1 and single rows/cols). */
export function broadcast(mats: Matrix[], f: (...xs: Scalar[]) => Scalar): Matrix {
  let rows = 1, cols = 1;
  for (const m of mats) { rows = Math.max(rows, m.length); cols = Math.max(cols, m[0]?.length ?? 1); }
  const out: Matrix = [];
  for (let r = 0; r < rows; r++) {
    const row: Scalar[] = [];
    for (let c = 0; c < cols; c++) {
      const xs: Scalar[] = [];
      let bad = false;
      for (const m of mats) {
        const mr = m.length === 1 ? 0 : r, mc = (m[0]?.length ?? 1) === 1 ? 0 : c;
        if (mr >= m.length || mc >= (m[0]?.length ?? 0)) { bad = true; break; }
        xs.push(m[mr][mc]);
      }
      if (bad) { row.push(ERR.na); continue; }
      try { row.push(f(...xs)); } catch (e) { if (e instanceof Fail) row.push(e.err); else throw e; }
    }
    out.push(row);
  }
  return out;
}

export function transpose(m: Matrix): Matrix {
  const rows = m.length, cols = m[0]?.length ?? 0;
  const out: Matrix = [];
  for (let c = 0; c < cols; c++) { const row: Scalar[] = []; for (let r = 0; r < rows; r++) row.push(m[r][c]); out.push(row); }
  return out;
}

export function flat(m: Matrix): Scalar[] { const out: Scalar[] = []; for (const r of m) for (const x of r) out.push(x); return out; }

export { ERR, CellError };
