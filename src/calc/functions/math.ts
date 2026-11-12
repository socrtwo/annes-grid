/**
 * Math & Trig functions: arithmetic, rounding, trigonometry, combinatorics, matrices, conditional sums,
 * SUBTOTAL and AGGREGATE.
 *
 * Scalar functions rely on the engine's automatic lifting for array arguments; aggregates declare
 * `arrays` and iterate their arguments.
 */
import { defFn, type FnDef } from './registry';
import {
  bool, fail, isRef, isRefs, isMatrix, iterValues, toMatrix, num, str, numbers, ERR, CellError,
  type Ctx, type Value,
} from '../values';
import type { Matrix, Scalar } from '../../model/types';
import {
  combin, factorial, mean, stdev, variance, median, modeSngl, kth, percentileInc, percentileExc, quartileInc, quartileExc,
  grid, resized, ifsValues, ifsArrays, sum as sumOf,
} from './stats';

type Args = Value[];
const M = 'Math & Trig' as const;
const sc = (a: Args, i: number): Scalar => a[i] as Scalar;
const n0 = (a: Args, i = 0): number => num(sc(a, i));
const optN = (a: Args, i: number, def: number): number => (i >= a.length || a[i] == null ? def : num(sc(a, i)));
const tr = Math.trunc;
const TRIG_LIMIT = 134217728; // 2^27

/** Snaps a quotient to 15 significant digits so 0.3/0.1 floors to 3, not 2. */
const snap = (x: number): number => (x === 0 || !Number.isFinite(x) ? x : +x.toPrecision(15));

/** Rounds |x| at `digits` decimals with the given integer rounding, keeping the sign. */
export function roundTo(x: number, digits: number, mode: 'near' | 'down' | 'up'): number {
  const d = tr(digits);
  if (x === 0 || !Number.isFinite(x)) return x;
  const sign = x < 0 ? -1 : 1, a = Math.abs(x);
  if (d < -308) return 0;
  if (d > 308) return x;
  const p = Math.pow(10, Math.abs(d));
  let y = d >= 0 ? a * p : a / p;
  if (!Number.isFinite(y) || y >= 1e16) return x;
  y = snap(y);
  const r = mode === 'near' ? Math.floor(y + 0.5) : mode === 'down' ? Math.floor(y) : Math.ceil(y);
  return sign * (d >= 0 ? r / p : r * p);
}

function trig(name: string, desc: string, f: (x: number) => number, opts: { limit?: boolean; zeroDiv?: boolean; domain?: (x: number) => boolean } = {}): FnDef {
  return {
    name, min: 1, max: 1, category: M, desc, params: ['number'],
    impl: (a) => {
      const x = n0(a);
      if (opts.limit && Math.abs(x) >= TRIG_LIMIT) return ERR.num;
      if (opts.domain && !opts.domain(x)) return ERR.num;
      if (opts.zeroDiv && x === 0) return ERR.div0;
      const r = f(x);
      return Number.isNaN(r) ? ERR.num : r;
    },
  };
}

function power(x: number, y: number): number {
  if (x === 0 && y === 0) return fail(ERR.num);
  if (x === 0 && y < 0) return fail(ERR.div0);
  const r = Math.pow(x, y);
  return Number.isNaN(r) ? fail(ERR.num) : r;
}

/** Integers for GCD/LCM. */
function intArgs(args: Args, ctx: Ctx): number[] {
  const out: number[] = [];
  for (const a of args) for (const { v, fromRef } of iterValues(a, ctx)) {
    if (v == null) { if (!fromRef) out.push(0); continue; }
    if (v instanceof CellError) return fail(v);
    if (typeof v === 'boolean') return fail(ERR.value);
    const n = typeof v === 'number' ? v : num(v);
    if (n < 0 || n >= 9007199254740992) return fail(ERR.num);
    out.push(Math.floor(n));
  }
  return out;
}

const gcd2 = (a: number, b: number): number => { while (b) { const t = a % b; a = b; b = t; } return a; };

function roman(n: number, mode: number): string {
  const chars = ['M', 'D', 'C', 'L', 'X', 'V', 'I'];
  const vals = [1000, 500, 100, 50, 10, 5, 1];
  const maxI = vals.length - 1;
  let out = '';
  let v = n;
  for (let i = 0; i <= maxI / 2; i++) {
    let idx = 2 * i;
    const digit = Math.floor(v / vals[idx]);
    if (digit % 5 === 4) {
      const idx2 = digit === 4 ? idx - 1 : idx - 2;
      let steps = 0;
      while (steps < mode && idx < maxI) {
        steps++;
        if (vals[idx2] - vals[idx + 1] <= v) idx++;
        else steps = mode;
      }
      out += chars[idx] + chars[idx2];
      v = v + vals[idx] - vals[idx2];
    } else {
      if (digit > 4) out += chars[idx - 1];
      out += chars[idx].repeat(digit % 5);
      v %= vals[idx];
    }
  }
  return out;
}

/** Strictly numeric matrix (MMULT/MDETERM/MINVERSE): blanks and text give #VALUE!. */
function numMatrix(v: Value, ctx: Ctx): number[][] {
  const m = toMatrix(v, ctx, false);
  return m.map((row) => row.map((x) => {
    if (x instanceof CellError) return fail(x);
    if (typeof x !== 'number') return fail(ERR.value);
    return x;
  }));
}

function det(a: number[][]): number {
  const n = a.length;
  const m = a.map((r) => r.slice());
  let d = 1;
  for (let c = 0; c < n; c++) {
    let p = c;
    for (let r = c + 1; r < n; r++) if (Math.abs(m[r][c]) > Math.abs(m[p][c])) p = r;
    if (m[p][c] === 0) return 0;
    if (p !== c) { [m[p], m[c]] = [m[c], m[p]]; d = -d; }
    d *= m[c][c];
    for (let r = c + 1; r < n; r++) {
      const f = m[r][c] / m[c][c];
      if (f) for (let k = c; k < n; k++) m[r][k] -= f * m[c][k];
    }
  }
  return d;
}

function inverse(a: number[][]): number[][] | null {
  const n = a.length;
  const m = a.map((r, i) => [...r, ...Array.from({ length: n }, (_, j) => (i === j ? 1 : 0))]);
  let scale = 0;
  for (const r of a) for (const x of r) scale = Math.max(scale, Math.abs(x));
  for (let c = 0; c < n; c++) {
    let p = c;
    for (let r = c + 1; r < n; r++) if (Math.abs(m[r][c]) > Math.abs(m[p][c])) p = r;
    if (Math.abs(m[p][c]) <= scale * 1e-15) return null;
    [m[p], m[c]] = [m[c], m[p]];
    const pv = m[c][c];
    for (let k = 0; k < 2 * n; k++) m[c][k] /= pv;
    for (let r = 0; r < n; r++) {
      if (r === c) continue;
      const f = m[r][c];
      if (f) for (let k = 0; k < 2 * n; k++) m[r][k] -= f * m[c][k];
    }
  }
  return m.map((r) => r.slice(n));
}

// --------------------------------------------------------------------------------------------
// SUBTOTAL / AGGREGATE

const NESTED = /(^|[^A-Za-z0-9_.])(SUBTOTAL|AGGREGATE)\s*\(/i;

interface Item { v: Scalar; fromRef: boolean }

function* aggItems(v: Value, ctx: Ctx, skipHidden: boolean, skipNested: boolean, skipFiltered: boolean): Generator<Item> {
  const areas = isRef(v) ? [v] : isRefs(v) ? v.areas : null;
  if (!areas) { yield* iterValues(v, ctx); return; }
  for (const a of areas) for (const sh of a.sheets || [a.sheet]) {
    const b = ctx.usedBounds(sh);
    const r2 = Math.min(a.r2, b.r2), c2 = Math.min(a.c2, b.c2);
    const fr = sh.filter?.range;
    for (let r = a.r1; r <= r2; r++) {
      if (sh.hiddenRows.has(r) && (skipHidden || (skipFiltered && fr && r >= fr.r1 && r <= fr.r2))) continue;
      for (let c = a.c1; c <= c2; c++) {
        if (skipNested) {
          const f = sh.get(r, c)?.f;
          if (f && NESTED.test(f)) continue;
        }
        yield { v: ctx.cell(sh, r, c), fromRef: true };
      }
    }
  }
}

function aggNumbers(items: Iterable<Item>, skipErrors: boolean): number[] {
  const out: number[] = [];
  for (const { v, fromRef } of items) {
    if (typeof v === 'number') out.push(v);
    else if (v instanceof CellError) { if (!skipErrors) fail(v); }
    else if (!fromRef && v != null) out.push(num(v));
  }
  return out;
}

function aggregate(fn: number, items: Item[], skipErrors: boolean, k?: number): number {
  if (fn === 2) {
    let n = 0;
    for (const { v, fromRef } of items) {
      if (typeof v === 'number') n++;
      else if (!fromRef && (typeof v === 'boolean' || (typeof v === 'string' && !Number.isNaN(tryNum(v))))) n++;
    }
    return n;
  }
  if (fn === 3) {
    let n = 0;
    for (const { v, fromRef } of items) {
      if (v instanceof CellError && skipErrors) continue;
      if (v != null || !fromRef) n++;
    }
    return n;
  }
  const xs = aggNumbers(items, skipErrors);
  switch (fn) {
    case 1: return mean(xs);
    case 4: return xs.length ? Math.max(...xs) : 0;
    case 5: return xs.length ? Math.min(...xs) : 0;
    case 6: { if (!xs.length) return 0; let p = 1; for (const x of xs) p *= x; return p; }
    case 7: return stdev(xs, true);
    case 8: return stdev(xs, false);
    case 9: return sumOf(xs);
    case 10: return variance(xs, true);
    case 11: return variance(xs, false);
    case 12: return median(xs);
    case 13: return modeSngl(xs);
    case 14: return kth(xs, k!, true);
    case 15: return kth(xs, k!, false);
    case 16: return percentileInc(xs, k!);
    case 17: return quartileInc(xs, k!);
    case 18: return percentileExc(xs, k!);
    case 19: return quartileExc(xs, k!);
    default: return fail(ERR.value);
  }
}

function tryNum(s: string): number { try { return num(s); } catch { return NaN; } }

// --------------------------------------------------------------------------------------------

const defs: FnDef[] = [
  { name: 'ABS', min: 1, max: 1, category: M, desc: 'Absolute value of a number.', params: ['number'], impl: (a) => Math.abs(n0(a)) },
  { name: 'SIGN', min: 1, max: 1, category: M, desc: 'Sign of a number: 1, 0 or -1.', params: ['number'], impl: (a) => Math.sign(n0(a)) },
  { name: 'PI', min: 0, max: 0, category: M, desc: 'The constant pi to 15 digits.', params: [], impl: () => Math.PI },
  { name: 'SQRT', min: 1, max: 1, category: M, desc: 'Positive square root.', params: ['number'], impl: (a) => { const x = n0(a); return x < 0 ? ERR.num : Math.sqrt(x); } },
  { name: 'SQRTPI', min: 1, max: 1, category: M, desc: 'Square root of a number multiplied by pi.', params: ['number'], impl: (a) => { const x = n0(a); return x < 0 ? ERR.num : Math.sqrt(x * Math.PI); } },
  { name: 'EXP', min: 1, max: 1, category: M, desc: 'e raised to the given power.', params: ['number'], impl: (a) => Math.exp(n0(a)) },
  { name: 'LN', min: 1, max: 1, category: M, desc: 'Natural logarithm.', params: ['number'], impl: (a) => { const x = n0(a); return x <= 0 ? ERR.num : Math.log(x); } },
  { name: 'LOG10', min: 1, max: 1, category: M, desc: 'Base-10 logarithm.', params: ['number'], impl: (a) => { const x = n0(a); return x <= 0 ? ERR.num : Math.log10(x); } },
  {
    name: 'LOG', min: 1, max: 2, category: M, desc: 'Logarithm of a number to a base (10 if omitted).', params: ['number', '[base]'],
    impl: (a) => {
      const x = n0(a), b = optN(a, 1, 10);
      if (x <= 0 || b <= 0) return ERR.num;
      if (b === 1) return ERR.div0;
      if (b === 10) return Math.log10(x);
      if (b === 2) return Math.log2(x);
      return Math.log(x) / Math.log(b);
    },
  },
  { name: 'POWER', min: 2, max: 2, category: M, desc: 'A number raised to a power.', params: ['number', 'power'], impl: (a) => power(n0(a), n0(a, 1)) },
  { name: 'DEGREES', min: 1, max: 1, category: M, desc: 'Converts radians to degrees.', params: ['angle'], impl: (a) => (n0(a) * 180) / Math.PI },
  { name: 'RADIANS', min: 1, max: 1, category: M, desc: 'Converts degrees to radians.', params: ['angle'], impl: (a) => (n0(a) * Math.PI) / 180 },
  trig('SIN', 'Sine of an angle in radians.', Math.sin, { limit: true }),
  trig('COS', 'Cosine of an angle in radians.', Math.cos, { limit: true }),
  trig('TAN', 'Tangent of an angle in radians.', Math.tan, { limit: true }),
  trig('CSC', 'Cosecant of an angle in radians.', (x) => 1 / Math.sin(x), { limit: true, zeroDiv: true }),
  trig('SEC', 'Secant of an angle in radians.', (x) => 1 / Math.cos(x), { limit: true }),
  trig('COT', 'Cotangent of an angle in radians.', (x) => 1 / Math.tan(x), { limit: true, zeroDiv: true }),
  trig('SINH', 'Hyperbolic sine.', Math.sinh),
  trig('COSH', 'Hyperbolic cosine.', Math.cosh),
  trig('TANH', 'Hyperbolic tangent.', Math.tanh),
  trig('CSCH', 'Hyperbolic cosecant.', (x) => 1 / Math.sinh(x), { limit: true, zeroDiv: true }),
  trig('SECH', 'Hyperbolic secant.', (x) => 1 / Math.cosh(x), { limit: true }),
  trig('COTH', 'Hyperbolic cotangent.', (x) => 1 / Math.tanh(x), { limit: true, zeroDiv: true }),
  trig('ASIN', 'Arcsine, in radians from -pi/2 to pi/2.', Math.asin, { domain: (x) => x >= -1 && x <= 1 }),
  trig('ACOS', 'Arccosine, in radians from 0 to pi.', Math.acos, { domain: (x) => x >= -1 && x <= 1 }),
  trig('ATAN', 'Arctangent, in radians from -pi/2 to pi/2.', Math.atan),
  trig('ACOT', 'Arccotangent, in radians from 0 to pi.', (x) => Math.PI / 2 - Math.atan(x)),
  trig('ASINH', 'Inverse hyperbolic sine.', Math.asinh),
  trig('ACOSH', 'Inverse hyperbolic cosine.', Math.acosh, { domain: (x) => x >= 1 }),
  trig('ATANH', 'Inverse hyperbolic tangent.', Math.atanh, { domain: (x) => x > -1 && x < 1 }),
  trig('ACOTH', 'Inverse hyperbolic cotangent.', (x) => 0.5 * Math.log((x + 1) / (x - 1)), { domain: (x) => Math.abs(x) > 1 }),
  {
    name: 'ATAN2', min: 2, max: 2, category: M, desc: 'Arctangent of the point (x_num, y_num), in radians.', params: ['x_num', 'y_num'],
    impl: (a) => { const x = n0(a), y = n0(a, 1); return x === 0 && y === 0 ? ERR.div0 : Math.atan2(y, x); },
  },
  // --- Rounding ----------------------------------------------------------------------------
  { name: 'ROUND', min: 2, max: 2, category: M, desc: 'Rounds a number to a number of digits.', params: ['number', 'num_digits'], impl: (a) => roundTo(n0(a), n0(a, 1), 'near') },
  { name: 'ROUNDDOWN', min: 2, max: 2, category: M, desc: 'Rounds a number toward zero.', params: ['number', 'num_digits'], impl: (a) => roundTo(n0(a), n0(a, 1), 'down') },
  { name: 'ROUNDUP', min: 2, max: 2, category: M, desc: 'Rounds a number away from zero.', params: ['number', 'num_digits'], impl: (a) => roundTo(n0(a), n0(a, 1), 'up') },
  { name: 'TRUNC', min: 1, max: 2, category: M, desc: 'Truncates a number by removing its fractional part (or digits beyond a precision).', params: ['number', '[num_digits]'], impl: (a) => roundTo(n0(a), optN(a, 1, 0), 'down') },
  { name: 'INT', min: 1, max: 1, category: M, desc: 'Rounds a number down to the nearest integer.', params: ['number'], impl: (a) => Math.floor(n0(a)) },
  {
    name: 'EVEN', min: 1, max: 1, category: M, desc: 'Rounds a number away from zero to the nearest even integer.', params: ['number'],
    impl: (a) => { const x = n0(a); const r = Math.ceil(snap(Math.abs(x)) / 2) * 2; return x < 0 ? -r : r; },
  },
  {
    name: 'ODD', min: 1, max: 1, category: M, desc: 'Rounds a number away from zero to the nearest odd integer.', params: ['number'],
    impl: (a) => { const x = n0(a); let r = Math.ceil(snap(Math.abs(x))); if (r % 2 === 0) r++; return x < 0 ? -r : r; },
  },
  {
    name: 'MROUND', min: 2, max: 2, category: M, desc: 'Rounds a number to the nearest multiple.', params: ['number', 'multiple'],
    impl: (a) => {
      const x = n0(a), m = n0(a, 1);
      if (m === 0 || x === 0) return 0;
      if ((x < 0) !== (m < 0)) return ERR.num;
      return snap(Math.floor(snap(x / m) + 0.5) * m);
    },
  },
  {
    name: 'CEILING', min: 2, max: 2, category: M, desc: 'Rounds a number up to the nearest multiple of significance.', params: ['number', 'significance'],
    impl: (a) => {
      const x = n0(a), s = n0(a, 1);
      if (s === 0 || x === 0) return 0;
      if (x > 0 && s < 0) return ERR.num;
      return snap(Math.ceil(snap(x / s)) * s);
    },
  },
  {
    name: 'FLOOR', min: 2, max: 2, category: M, desc: 'Rounds a number down to the nearest multiple of significance.', params: ['number', 'significance'],
    impl: (a) => {
      const x = n0(a), s = n0(a, 1);
      if (x === 0) return 0;
      if (s === 0) return ERR.div0;
      if (x > 0 && s < 0) return ERR.num;
      return snap(Math.floor(snap(x / s)) * s);
    },
  },
  {
    name: 'CEILING.MATH', min: 1, max: 3, category: M, desc: 'Rounds a number up to the nearest integer or multiple; mode controls negative numbers.', params: ['number', '[significance]', '[mode]'],
    impl: (a) => {
      const x = n0(a), s = Math.abs(optN(a, 1, 1)), mode = optN(a, 2, 0);
      if (s === 0 || x === 0) return 0;
      const q = snap(x / s);
      return snap((x < 0 && mode !== 0 ? Math.floor(q) : Math.ceil(q)) * s);
    },
  },
  {
    name: 'FLOOR.MATH', min: 1, max: 3, category: M, desc: 'Rounds a number down to the nearest integer or multiple; mode controls negative numbers.', params: ['number', '[significance]', '[mode]'],
    impl: (a) => {
      const x = n0(a), s = Math.abs(optN(a, 1, 1)), mode = optN(a, 2, 0);
      if (s === 0 || x === 0) return 0;
      const q = snap(x / s);
      return snap((x < 0 && mode !== 0 ? Math.ceil(q) : Math.floor(q)) * s);
    },
  },
  {
    name: 'CEILING.PRECISE', min: 1, max: 2, category: M, desc: 'Rounds a number up to the nearest multiple, regardless of sign.', params: ['number', '[significance]'],
    impl: (a) => { const x = n0(a), s = Math.abs(optN(a, 1, 1)); return s === 0 || x === 0 ? 0 : snap(Math.ceil(snap(x / s)) * s); },
  },
  {
    name: 'ISO.CEILING', min: 1, max: 2, category: M, desc: 'Rounds a number up to the nearest multiple, regardless of sign.', params: ['number', '[significance]'],
    impl: (a) => { const x = n0(a), s = Math.abs(optN(a, 1, 1)); return s === 0 || x === 0 ? 0 : snap(Math.ceil(snap(x / s)) * s); },
  },
  {
    name: 'FLOOR.PRECISE', min: 1, max: 2, category: M, desc: 'Rounds a number down to the nearest multiple, regardless of sign.', params: ['number', '[significance]'],
    impl: (a) => { const x = n0(a), s = Math.abs(optN(a, 1, 1)); return s === 0 || x === 0 ? 0 : snap(Math.floor(snap(x / s)) * s); },
  },
  // --- Integer arithmetic ------------------------------------------------------------------
  {
    name: 'MOD', min: 2, max: 2, category: M, desc: 'Remainder after division; the result has the sign of the divisor.', params: ['number', 'divisor'],
    impl: (a) => {
      const x = n0(a), d = n0(a, 1);
      if (d === 0) return ERR.div0;
      const q = Math.floor(snap(x / d));
      let r = x - d * q;
      // For a huge quotient the subtraction cancels to rubbish; the exact remainder is still available.
      if (!(Math.abs(r) < Math.abs(d)) || (r !== 0 && (r < 0) !== (d < 0))) { r = x % d; if (r !== 0 && (r < 0) !== (d < 0)) r += d; }
      return Math.abs(r) < Math.abs(d) * 1e-15 ? 0 : r;
    },
  },
  {
    name: 'QUOTIENT', min: 2, max: 2, category: M, desc: 'Integer portion of a division.', params: ['numerator', 'denominator'],
    impl: (a) => { const x = n0(a), d = n0(a, 1); return d === 0 ? ERR.div0 : tr(snap(x / d)); },
  },
  {
    name: 'GCD', min: 1, max: 255, category: M, arrays: 'all', desc: 'Greatest common divisor of integers.', params: ['number1', '[number2]', '...'],
    impl: (a, ctx) => intArgs(a, ctx).reduce(gcd2, 0),
  },
  {
    name: 'LCM', min: 1, max: 255, category: M, arrays: 'all', desc: 'Least common multiple of integers.', params: ['number1', '[number2]', '...'],
    impl: (a, ctx) => {
      const xs = intArgs(a, ctx);
      let l = 1;
      for (const x of xs) { if (x === 0) return 0; l = (l / gcd2(l, x)) * x; }
      return l;
    },
  },
  {
    name: 'FACT', min: 1, max: 1, category: M, desc: 'Factorial of a number.', params: ['number'],
    impl: (a) => { const x = Math.floor(n0(a)); return x < 0 || x > 170 ? ERR.num : factorial(x); },
  },
  {
    name: 'FACTDOUBLE', min: 1, max: 1, category: M, desc: 'Double factorial of a number.', params: ['number'],
    impl: (a) => {
      const x = Math.floor(n0(a));
      if (x < -1) return ERR.num;
      let r = 1;
      for (let i = x; i > 1; i -= 2) r *= i;
      return Number.isFinite(r) ? r : ERR.num;
    },
  },
  {
    name: 'COMBIN', min: 2, max: 2, category: M, desc: 'Number of combinations of k items chosen from n.', params: ['number', 'number_chosen'],
    impl: (a) => {
      const n = tr(n0(a)), k = tr(n0(a, 1));
      if (n < 0 || k < 0 || n < k) return ERR.num;
      return combin(n, k);
    },
  },
  {
    name: 'COMBINA', min: 2, max: 2, category: M, desc: 'Number of combinations with repetition of k items chosen from n.', params: ['number', 'number_chosen'],
    impl: (a) => {
      const n = tr(n0(a)), k = tr(n0(a, 1));
      if (n < 0 || k < 0) return ERR.num;
      if (k === 0) return 1;
      if (n === 0) return 0;
      return combin(n + k - 1, k);
    },
  },
  {
    name: 'MULTINOMIAL', min: 1, max: 255, category: M, arrays: 'all', desc: 'Ratio of the factorial of a sum to the product of factorials.', params: ['number1', '[number2]', '...'],
    impl: (a, ctx) => {
      const xs = numbers(a, ctx).map(Math.floor);
      let s = 0, r = 1;
      for (const x of xs) {
        if (x < 0) return ERR.num;
        s += x;
        r *= combin(s, x);
      }
      return Number.isFinite(r) ? r : ERR.num;
    },
  },
  // --- Number systems ----------------------------------------------------------------------
  {
    name: 'ROMAN', min: 1, max: 2, category: M, desc: 'Converts an Arabic number to Roman numerals as text.', params: ['number', '[form]'],
    impl: (a) => {
      const x = Math.floor(n0(a));
      let form = 0;
      if (a.length > 1 && a[1] != null) {
        const f = sc(a, 1);
        form = typeof f === 'boolean' ? (f ? 0 : 4) : Math.floor(num(f));
      }
      if (x < 0 || x > 3999 || form < 0 || form > 4) return ERR.value;
      return roman(x, form);
    },
  },
  {
    name: 'ARABIC', min: 1, max: 1, category: M, desc: 'Converts a Roman numeral to an Arabic number.', params: ['text'],
    impl: (a) => {
      let t = str(sc(a, 0)).trim().toUpperCase();
      if (t.length > 255) return ERR.value;
      let neg = false;
      if (t.startsWith('-')) { neg = true; t = t.slice(1); }
      const V: Record<string, number> = { I: 1, V: 5, X: 10, L: 50, C: 100, D: 500, M: 1000 };
      let total = 0, max = 0;
      for (let i = t.length - 1; i >= 0; i--) {
        const v = V[t[i]];
        if (!v) return ERR.value;
        if (v >= max) { total += v; max = v; } else total -= v;
      }
      return neg ? -total : total;
    },
  },
  {
    name: 'BASE', min: 2, max: 3, category: M, desc: 'Converts a number to text in the given radix (base).', params: ['number', 'radix', '[min_length]'],
    impl: (a) => {
      const x = tr(n0(a)), r = tr(n0(a, 1)), len = tr(optN(a, 2, 0));
      if (x < 0 || x >= 9007199254740992 || r < 2 || r > 36 || len < 0 || len > 255) return ERR.num;
      return x.toString(r).toUpperCase().padStart(len, '0');
    },
  },
  {
    name: 'DECIMAL', min: 2, max: 2, category: M, desc: 'Converts a text representation of a number in a given base to decimal.', params: ['text', 'radix'],
    impl: (a) => {
      const t = str(sc(a, 0)).trim().toUpperCase(), r = tr(n0(a, 1));
      if (r < 2 || r > 36 || t.length > 255) return ERR.num;
      let v = 0;
      for (const ch of t) {
        const d = parseInt(ch, 36);
        if (Number.isNaN(d) || d >= r) return ERR.num;
        v = v * r + d;
      }
      return v >= 9007199254740992 ? ERR.num : v;
    },
  },
  // --- Random ------------------------------------------------------------------------------
  { name: 'RAND', min: 0, max: 0, category: M, volatile: true, desc: 'Random number between 0 (inclusive) and 1 (exclusive).', params: [], impl: (_a, ctx) => ctx.random() },
  {
    name: 'RANDBETWEEN', min: 2, max: 2, category: M, volatile: true, desc: 'Random integer between two numbers, inclusive.', params: ['bottom', 'top'],
    impl: (a, ctx) => {
      const lo = Math.ceil(n0(a)), hi = Math.floor(n0(a, 1));
      if (lo > hi) return ERR.num;
      return lo + Math.floor(ctx.random() * (hi - lo + 1));
    },
  },
  {
    name: 'RANDARRAY', min: 0, max: 5, category: M, volatile: true, desc: 'Array of random numbers with optional size, bounds and whole-number setting.', params: ['[rows]', '[columns]', '[min]', '[max]', '[whole_number]'],
    impl: (a, ctx) => {
      const rows = tr(optN(a, 0, 1)), cols = tr(optN(a, 1, 1)), lo = optN(a, 2, 0), hi = optN(a, 3, 1);
      const whole = a.length > 4 && a[4] != null ? bool(sc(a, 4)) : false;
      if (rows < 0 || cols < 0 || lo > hi) return ERR.value;
      if (rows === 0 || cols === 0) return ERR.calc;
      if (rows * cols > 1048576 * 16) return ERR.num;
      if (whole && (!Number.isInteger(lo) || !Number.isInteger(hi))) return ERR.value;
      const out: Matrix = [];
      for (let r = 0; r < rows; r++) {
        const row: Scalar[] = [];
        for (let c = 0; c < cols; c++) row.push(whole ? lo + Math.floor(ctx.random() * (hi - lo + 1)) : lo + ctx.random() * (hi - lo));
        out.push(row);
      }
      return out;
    },
  },
  // --- Sums & products ---------------------------------------------------------------------
  { name: 'SUM', min: 1, max: 255, category: M, arrays: 'all', desc: 'Adds its arguments.', params: ['number1', '[number2]', '...'], impl: (a, ctx) => sumOf(numbers(a, ctx)) },
  {
    name: 'PRODUCT', min: 1, max: 255, category: M, arrays: 'all', desc: 'Multiplies its arguments.', params: ['number1', '[number2]', '...'],
    impl: (a, ctx) => { const xs = numbers(a, ctx); if (!xs.length) return 0; let p = 1; for (const x of xs) p *= x; return p; },
  },
  {
    name: 'SUMSQ', min: 1, max: 255, category: M, arrays: 'all', desc: 'Sum of the squares of the arguments.', params: ['number1', '[number2]', '...'],
    impl: (a, ctx) => sumOf(numbers(a, ctx).map((x) => x * x)),
  },
  {
    name: 'SUMIF', min: 2, max: 3, category: M, arrays: [0, 2], desc: 'Adds the cells that meet a condition.', params: ['range', 'criteria', '[sum_range]'],
    impl: (a, ctx) => {
      const rg = grid(a[0], ctx);
      const target = a.length > 2 && a[2] != null ? resized(a[2], rg.rows, rg.cols, ctx) : rg;
      return sumOf(ifsValues(target, [a[0], a[1]], ctx).vals);
    },
  },
  {
    name: 'SUMIFS', min: 3, max: 255, category: M, arrays: ifsArrays([0], 1), desc: 'Adds the cells that meet every one of several conditions.', params: ['sum_range', 'criteria_range1', 'criteria1', '[criteria_range2]', '[criteria2]', '...'],
    impl: (a, ctx) => (a.length % 2 === 0 ? ERR.value : sumOf(ifsValues(grid(a[0], ctx), a.slice(1), ctx).vals)),
  },
  {
    name: 'SUMPRODUCT', min: 1, max: 255, category: M, arrays: 'all', desc: 'Sum of the products of corresponding array entries.', params: ['array1', '[array2]', '...'],
    impl: (a, ctx) => {
      const mats = a.map((v) => toMatrix(v, ctx));
      const rows = mats[0].length, cols = mats[0][0]?.length ?? 0;
      for (const m of mats) if (m.length !== rows || (m[0]?.length ?? 0) !== cols) return ERR.value;
      let s = 0;
      for (let r = 0; r < rows; r++) for (let c = 0; c < cols; c++) {
        let p = 1;
        for (const m of mats) {
          const v = m[r][c];
          if (v instanceof CellError) return v;
          p *= typeof v === 'number' ? v : 0;
        }
        s += p;
      }
      return s;
    },
  },
  ...(['SUMX2MY2', 'SUMX2PY2', 'SUMXMY2'] as const).map((name): FnDef => ({
    name, min: 2, max: 2, category: M, arrays: [0, 1],
    desc: name === 'SUMX2MY2' ? 'Sum of the differences of squares of corresponding values.' : name === 'SUMX2PY2' ? 'Sum of the sums of squares of corresponding values.' : 'Sum of squares of the differences of corresponding values.',
    params: ['array_x', 'array_y'],
    impl: ([x, y], ctx) => {
      const xa = toMatrix(x, ctx).flat(), ya = toMatrix(y, ctx).flat();
      if (xa.length !== ya.length) return ERR.na;
      let s = 0;
      for (let i = 0; i < xa.length; i++) {
        const p = xa[i], q = ya[i];
        if (p instanceof CellError) return p;
        if (q instanceof CellError) return q;
        if (typeof p !== 'number' || typeof q !== 'number') continue;
        s += name === 'SUMX2MY2' ? p * p - q * q : name === 'SUMX2PY2' ? p * p + q * q : (p - q) ** 2;
      }
      return s;
    },
  })),
  {
    name: 'SERIESSUM', min: 4, max: 4, category: M, arrays: [3], desc: 'Sum of a power series.', params: ['x', 'n', 'm', 'coefficients'],
    impl: (a, ctx) => {
      const x = n0(a), n = n0(a, 1), m = n0(a, 2);
      let s = 0, i = 0;
      for (const { v } of iterValues(a[3], ctx)) {
        if (v instanceof CellError) return v;
        if (typeof v !== 'number') return ERR.value;
        s += v * Math.pow(x, n + i * m);
        i++;
      }
      return s;
    },
  },
  {
    name: 'PERCENTOF', min: 2, max: 2, category: M, arrays: [0, 1], desc: 'Share that a subset of values makes up of a total.', params: ['data_subset', 'data_all'],
    impl: ([p, t], ctx) => {
      const tot = sumOf(numbers([t], ctx));
      const part = sumOf(numbers([p], ctx));
      return tot === 0 ? ERR.div0 : part / tot;
    },
  },
  // --- Matrices ----------------------------------------------------------------------------
  {
    name: 'MDETERM', min: 1, max: 1, category: M, arrays: [0], desc: 'Determinant of a square matrix.', params: ['array'],
    impl: ([v], ctx) => { const m = numMatrix(v, ctx); return m.length !== (m[0]?.length ?? 0) ? ERR.value : det(m); },
  },
  {
    name: 'MINVERSE', min: 1, max: 1, category: M, arrays: [0], desc: 'Inverse of a square matrix.', params: ['array'],
    impl: ([v], ctx) => {
      const m = numMatrix(v, ctx);
      if (m.length !== (m[0]?.length ?? 0)) return ERR.value;
      return inverse(m) ?? ERR.num;
    },
  },
  {
    name: 'MMULT', min: 2, max: 2, category: M, arrays: [0, 1], desc: 'Matrix product of two arrays.', params: ['array1', 'array2'],
    impl: ([x, y], ctx) => {
      const a = numMatrix(x, ctx), b = numMatrix(y, ctx);
      const n = a.length, k = a[0]?.length ?? 0, m = b[0]?.length ?? 0;
      if (k !== b.length) return ERR.value;
      const out: Matrix = [];
      for (let i = 0; i < n; i++) {
        const row: Scalar[] = [];
        for (let j = 0; j < m; j++) { let s = 0; for (let l = 0; l < k; l++) s += a[i][l] * b[l][j]; row.push(s); }
        out.push(row);
      }
      return out;
    },
  },
  {
    name: 'MUNIT', min: 1, max: 1, category: M, desc: 'Identity matrix of the given size.', params: ['dimension'],
    impl: (a) => {
      const n = tr(n0(a));
      if (n < 1 || n > 16384) return ERR.value;
      return Array.from({ length: n }, (_, i) => Array.from({ length: n }, (_, j) => (i === j ? 1 : 0)));
    },
  },
  // --- SUBTOTAL / AGGREGATE ----------------------------------------------------------------
  {
    name: 'SUBTOTAL', min: 2, max: 255, category: M, arrays: [1, -1], desc: 'Subtotal of a list using a chosen function; ignores other subtotals (101-111 also skip hidden rows).', params: ['function_num', 'ref1', '[ref2]', '...'],
    impl: (a, ctx) => {
      const f = tr(n0(a));
      const code = f > 100 ? f - 100 : f;
      if (code < 1 || code > 11) return ERR.value;
      const items: Item[] = [];
      for (const v of a.slice(1)) for (const it of aggItems(v, ctx, f > 100, true, true)) items.push(it);
      return aggregate(code, items, false);
    },
  },
  {
    name: 'AGGREGATE', min: 3, max: 255, category: M, arrays: [2, -1], desc: 'Aggregate of a list using a chosen function, optionally ignoring hidden rows, errors and nested subtotals.', params: ['function_num', 'options', 'ref1', '[ref2_or_k]', '...'],
    impl: (a, ctx) => {
      const f = tr(n0(a));
      const opt = a[1] == null ? 0 : tr(n0(a, 1));
      if (f < 1 || f > 19 || opt < 0 || opt > 7) return ERR.value;
      const skipNested = opt <= 3, skipHidden = opt % 2 === 1, skipErrors = opt === 2 || opt === 3 || opt === 6 || opt === 7;
      if (f >= 14) {
        if (a.length < 4) return ERR.value;
        const items = [...aggItems(a[2], ctx, skipHidden, skipNested, false)];
        const kv = a[3];
        if (isRef(kv) || isMatrix(kv)) {
          const km = toMatrix(kv, ctx);
          if (km.length === 1 && km[0].length === 1) return aggregate(f, items, skipErrors, num(km[0][0]));
          return km.map((row) => row.map((k) => {
            try { return aggregate(f, items, skipErrors, num(k)); } catch (e) { if ((e as { err?: CellError }).err instanceof CellError) return (e as { err: CellError }).err; throw e; }
          }));
        }
        return aggregate(f, items, skipErrors, num(kv as Scalar));
      }
      const items: Item[] = [];
      for (const v of a.slice(2)) for (const it of aggItems(v, ctx, skipHidden, skipNested, false)) items.push(it);
      return aggregate(f, items, skipErrors);
    },
  },
];

defFn(defs);
