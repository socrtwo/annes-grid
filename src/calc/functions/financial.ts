/**
 * Financial functions: time value of money, depreciation, cash-flow analysis and securities
 * (coupon bonds, discounted securities, odd first/last periods, Treasury bills).
 *
 * Day-count bases follow the spreadsheet convention:
 *   0 = US (NASD) 30/360, 1 = actual/actual, 2 = actual/360, 3 = actual/365, 4 = European 30/360.
 * Coupon schedules are derived backwards from maturity (or from the first/last coupon for odd periods),
 * keeping end-of-month alignment when the anchor date is the last day of its month.
 */
import { defFn, type FnDef } from './registry';
import { num, bool, fail, numbers, iterValues, ERR, CellError, type Ctx, type Value } from '../values';
import { serialToParts, partsToSerial, daysInMonth, isLeap } from '../../model/dates';
import type { Scalar } from '../../model/types';

// --------------------------------------------------------------------------------------------
// Argument helpers

const S = (v: Value): Scalar => v as Scalar;
const has = (a: Value[], i: number): boolean => i < a.length && a[i] != null;
const optNum = (a: Value[], i: number, d: number): number => (has(a, i) ? num(S(a[i])) : d);
const optBool = (a: Value[], i: number, d: boolean): boolean => (has(a, i) ? bool(S(a[i])) : d);

function dateArg(v: Value): number {
  const n = num(S(v));
  if (n < 0) fail(ERR.num);
  return Math.trunc(n);
}

function basisArg(a: Value[], i: number): number {
  const b = Math.trunc(optNum(a, i, 0));
  if (b < 0 || b > 4) fail(ERR.num);
  return b;
}

function freqArg(v: Value): number {
  const f = Math.trunc(num(S(v)));
  if (f !== 1 && f !== 2 && f !== 4) fail(ERR.num);
  return f;
}

const typeArg = (a: Value[], i: number): number => (optNum(a, i, 0) !== 0 ? 1 : 0);

/** Fails with #NUM! for non-finite results. */
function fin(x: number): number {
  if (!Number.isFinite(x)) fail(ERR.num);
  return x;
}

// --------------------------------------------------------------------------------------------
// Calendar helpers

interface YMD { y: number; m: number; d: number }

class Cal {
  constructor(public d1904: boolean) {}
  ymd(serial: number): YMD { const p = serialToParts(serial, this.d1904); return { y: p.y, m: p.m, d: p.d }; }
  serial(y: number, m: number, d: number): number { return partsToSerial(y, m, d, this.d1904); }
  /** Same day n months later (clamped to month end), like EDATE. */
  edate(serial: number, months: number): number {
    const p = this.ymd(serial);
    let m = p.m - 1 + months;
    const y = p.y + Math.floor(m / 12);
    m = ((m % 12) + 12) % 12;
    return this.serial(y, m + 1, Math.min(p.d, daysInMonth(y, m + 1)));
  }
}

const calOf = (ctx: Ctx): Cal => new Cal(!!ctx.book?.calc?.date1904);

/**
 * Date that tracks an "original day" so that month arithmetic preserves end-of-month alignment,
 * and that knows how to count days under a given basis (30-day months for bases 0 and 4).
 */
class BDate {
  origDay = 1; day = 1; month = 1; year = 1900;
  lastDay = false; b30 = false; us = false;

  static of(cal: Cal, serial: number, basis: number): BDate {
    const p = cal.ymd(serial);
    const d = new BDate();
    d.origDay = p.d; d.month = p.m; d.year = p.y;
    d.lastDay = p.d >= daysInMonth(p.y, p.m);
    d.b30 = basis === 0 || basis === 4;
    d.us = basis === 0;
    d.setDay();
    return d;
  }

  clone(): BDate { const d = new BDate(); Object.assign(d, this); return d; }

  setDay(): void {
    const dim = daysInMonth(this.year, this.month);
    if (this.b30) {
      this.day = Math.min(this.origDay, 30);
      if (this.lastDay || this.day >= dim) this.day = 30;
    } else {
      this.day = this.lastDay ? dim : Math.min(this.origDay, dim);
    }
  }

  monthDays(): number { return this.b30 ? 30 : daysInMonth(this.year, this.month); }

  addMonths(n: number): this {
    let nm = n + this.month;
    if (nm > 12) {
      nm--;
      this.year += Math.trunc(nm / 12);
      this.month = (nm % 12) + 1;
    } else if (nm < 1) {
      this.year += Math.trunc(nm / 12) - 1;
      this.month = (nm % 12) + 12;
    } else this.month = nm;
    this.setDay();
    return this;
  }

  addYears(n: number): this { this.year += n; this.setDay(); return this; }
  setYear(y: number): this { this.year = y; this.setDay(); return this; }

  serial(cal: Cal): number {
    const dim = daysInMonth(this.year, this.month);
    const real = this.lastDay ? dim : Math.min(dim, this.origDay);
    return cal.serial(this.year, this.month, real);
  }

  cmp(o: BDate): number {
    return this.year !== o.year ? this.year - o.year : this.month !== o.month ? this.month - o.month : this.day - o.day;
  }
  lt(o: BDate): boolean { return this.cmp(o) < 0; }
  gt(o: BDate): boolean { return this.cmp(o) > 0; }

  private monthRangeDays(from: number, to: number): number {
    if (from > to) return 0;
    if (this.b30) return (to - from + 1) * 30;
    let n = 0;
    for (let m = from; m <= to; m++) n += daysInMonth(this.year, m);
    return n;
  }

  private yearRangeDays(from: number, to: number): number {
    if (from > to) return 0;
    if (this.b30) return (to - from + 1) * 360;
    let n = 0;
    for (let y = from; y <= to; y++) n += isLeap(y) ? 366 : 365;
    return n;
  }

  /** Day count between two dates under the dates' basis (always >= 0). */
  static diff(rFrom: BDate, rTo: BDate): number {
    if (rFrom.gt(rTo)) return BDate.diff(rTo, rFrom);
    const from = rFrom.clone(), to = rTo.clone();
    if (to.b30) {
      if (to.us) {
        if ((rFrom.month === 2 || rFrom.day < 30) && to.origDay === 31) to.day = 31;
        else if (to.month === 2 && to.lastDay) to.day = daysInMonth(to.year, 2);
      } else {
        if (from.month === 2 && from.day === 30) from.day = daysInMonth(from.year, 2);
        if (to.month === 2 && to.day === 30) to.day = daysInMonth(to.year, 2);
      }
    }
    let n = 0;
    if (from.year < to.year || (from.year === to.year && from.month < to.month)) {
      n = from.monthDays() - from.day + 1;
      from.origDay = from.day = 1;
      from.lastDay = false;
      from.addMonths(1);
      if (from.year < to.year) {
        n += from.monthRangeDays(from.month, 12);
        from.addMonths(13 - from.month);
        n += from.yearRangeDays(from.year, to.year - 1);
        from.addYears(to.year - from.year);
      }
      n += from.monthRangeDays(from.month, to.month - 1);
      from.addMonths(to.month - from.month);
    }
    n += to.day - from.day;
    return Math.max(n, 0);
  }
}

/** Days in a year for the basis (basis 1 uses the year of the given date only for coupon-free formulas). */
const yearDays = (basis: number): number => (basis === 3 ? 365 : 360);

/** Fraction of a year between two serial dates (YEARFRAC semantics). */
function yearFrac(cal: Cal, start: number, end: number, basis: number): number {
  if (start === end) return 0;
  if (start > end) [start, end] = [end, start];
  const a = cal.ymd(start), b = cal.ymd(end);
  let d1 = a.d, d2 = b.d;
  switch (basis) {
    case 0: {
      if (d1 === 31) d1--;
      if (d1 === 30 && d2 === 31) d2--;
      else if (a.m === 2 && d1 === daysInMonth(a.y, 2)) {
        d1 = 30;
        if (b.m === 2 && d2 === daysInMonth(b.y, 2)) d2 = 30;
      }
      return ((b.y - a.y) * 360 + (b.m - a.m) * 30 + (d2 - d1)) / 360;
    }
    case 1: {
      const diffYear = a.y !== b.y;
      if (diffYear && (b.y !== a.y + 1 || a.m < b.m || (a.m === b.m && a.d < b.d))) {
        let days = 0;
        for (let y = a.y; y <= b.y; y++) days += isLeap(y) ? 366 : 365;
        return (end - start) / (days / (b.y - a.y + 1));
      }
      let leap: boolean;
      if (!diffYear) leap = isLeap(a.y);
      else leap = (isLeap(a.y) && (a.m < 2 || (a.m === 2 && a.d <= 29))) || (isLeap(b.y) && (b.m > 2 || (b.m === 2 && b.d === 29)));
      return (end - start) / (leap ? 366 : 365);
    }
    case 2: return (end - start) / 360;
    case 3: return (end - start) / 365;
    default: {
      if (d1 === 31) d1 = 30;
      if (d2 === 31) d2 = 30;
      return ((b.y - a.y) * 360 + (b.m - a.m) * 30 + (d2 - d1)) / 360;
    }
  }
}

// --------------------------------------------------------------------------------------------
// Coupon schedule

interface Coup { cal: Cal; settle: BDate; mat: BDate; freq: number; basis: number }

function coup(cal: Cal, settle: number, mat: number, freq: number, basis: number): Coup {
  if (settle >= mat) fail(ERR.num);
  return { cal, settle: BDate.of(cal, settle, basis), mat: BDate.of(cal, mat, basis), freq, basis };
}

/**
 * A date as it falls in the calendar, for deciding which coupon date comes before or after settlement. The 30/360
 * day (where the 30th and the 31st are the same day) is for counting days only: compared that way, a settlement on
 * the 30th would count as being on a coupon date that falls on the 31st.
 */
const calendarKey = (d: BDate): number => {
  const dim = daysInMonth(d.year, d.month);
  return d.year * 10000 + d.month * 100 + (d.lastDay ? dim : Math.min(d.origDay, dim));
};

function couppcd(c: Coup): BDate {
  const settle = calendarKey(c.settle);
  const d = c.mat.clone().setYear(c.settle.year);
  if (calendarKey(d) < settle) d.addYears(1);
  while (calendarKey(d) > settle) d.addMonths(-12 / c.freq);
  return d;
}

function coupncd(c: Coup): BDate {
  const settle = calendarKey(c.settle);
  const d = c.mat.clone().setYear(c.settle.year);
  if (calendarKey(d) > settle) d.addYears(-1);
  while (calendarKey(d) <= settle) d.addMonths(12 / c.freq);
  return d;
}

const coupdaybs = (c: Coup): number => BDate.diff(couppcd(c), c.settle);

function coupdays(c: Coup): number {
  if (c.basis === 1) {
    const p = couppcd(c);
    return BDate.diff(p, p.clone().addMonths(12 / c.freq));
  }
  return yearDays(c.basis) / c.freq;
}

function coupdaysnc(c: Coup): number {
  if (c.basis !== 0 && c.basis !== 4) return BDate.diff(c.settle, coupncd(c));
  return coupdays(c) - coupdaybs(c);
}

function coupnum(c: Coup): number {
  const p = couppcd(c);
  const months = (c.mat.year - p.year) * 12 + c.mat.month - p.month;
  return (months * c.freq) / 12;
}

function bondPrice(c: Coup, rate: number, yld: number, red: number): number {
  const f = c.freq;
  const E = coupdays(c), A = coupdaybs(c), DSC = coupdaysnc(c), N = coupnum(c);
  const C = (100 * rate) / f;
  if (N <= 1) return (red + C) / (1 + (DSC / E) * (yld / f)) - (A / E) * C;
  const dsc = DSC / E, base = 1 + yld / f;
  let p = red / Math.pow(base, N - 1 + dsc);
  for (let k = 1; k <= N; k++) p += C / Math.pow(base, k - 1 + dsc);
  return p - C * (A / E);
}

function bondDuration(c: Coup, coupon: number, yld: number): number {
  const f = c.freq;
  const N = coupnum(c), dsc = coupdaysnc(c) / coupdays(c);
  const C = (100 * coupon) / f, base = 1 + yld / f;
  let num_ = 0, den = 0;
  for (let k = 1; k <= N; k++) {
    const t = k - 1 + dsc;
    const cf = C + (k === N ? 100 : 0);
    const pv = cf / Math.pow(base, t);
    num_ += t * pv;
    den += pv;
  }
  return num_ / den / f;
}

// --------------------------------------------------------------------------------------------
// Root finding

/**
 * Newton's method from `guess`, falling back to a bracketed bisection over [lo, hi] when Newton fails
 * to converge. Returns NaN when no root is found.
 */
function solve(f: (x: number) => number, guess: number, lo: number, hi: number, df?: (x: number) => number): number {
  const tol = 1e-12;
  let x = guess;
  for (let i = 0; i < 100; i++) {
    const y = f(x);
    if (!Number.isFinite(y)) break;
    if (y === 0) return x;
    let d: number;
    if (df) d = df(x);
    else { const h = Math.max(Math.abs(x) * 1e-7, 1e-9); d = (f(x + h) - f(x - h)) / (2 * h); }
    if (!Number.isFinite(d) || d === 0) break;
    const nx = x - y / d;
    if (!Number.isFinite(nx) || nx <= lo) break;
    if (Math.abs(nx - x) <= tol * Math.max(1, Math.abs(nx))) {
      const fy = f(nx);
      if (Number.isFinite(fy)) return nx;
      break;
    }
    x = nx;
  }
  return bisect(f, guess, lo, hi);
}

function bisect(f: (x: number) => number, guess: number, lo: number, hi: number): number {
  // Scan outward from the guess for a sign change.
  const pts: number[] = [];
  const g = Math.min(Math.max(guess, lo + 1e-9), hi);
  for (let i = 0; i <= 60; i++) {
    const step = Math.pow(1.4, i) * 1e-3 - 1e-3;
    pts.push(g + step, g - step);
  }
  const xs = [...new Set(pts.filter((x) => x > lo && x <= hi))].sort((a, b) => Math.abs(a - g) - Math.abs(b - g));
  let best: [number, number] | null = null;
  let px = g, py = f(g);
  if (py === 0) return g;
  // Walk in each direction separately.
  for (const dir of [1, -1]) {
    px = g; py = f(g);
    for (const x of xs) {
      if ((x - g) * dir <= 0) continue;
      const y = f(x);
      if (!Number.isFinite(y)) continue;
      if (Number.isFinite(py) && Math.sign(y) !== Math.sign(py)) { best = dir > 0 ? [px, x] : [x, px]; break; }
      px = x; py = y;
    }
    if (best) break;
  }
  if (!best) return NaN;
  let [a, b] = best;
  let fa = f(a);
  for (let i = 0; i < 200; i++) {
    const m = (a + b) / 2;
    const fm = f(m);
    if (fm === 0 || Math.abs(b - a) <= 1e-15 * Math.max(1, Math.abs(m))) return m;
    if (Math.sign(fm) === Math.sign(fa)) { a = m; fa = fm; } else b = m;
  }
  return (a + b) / 2;
}

// --------------------------------------------------------------------------------------------
// Annuity primitives

function pmtCalc(r: number, n: number, pv: number, fv: number, type: number): number {
  if (r === 0) return -(pv + fv) / n;
  const p = Math.pow(1 + r, n);
  return -(r * (fv + pv * p)) / ((1 + r * type) * (p - 1));
}

function fvCalc(r: number, n: number, pmt: number, pv: number, type: number): number {
  if (r === 0) return -(pv + pmt * n);
  const p = Math.pow(1 + r, n);
  return -(pv * p + (pmt * (1 + r * type) * (p - 1)) / r);
}

function pvCalc(r: number, n: number, pmt: number, fv: number, type: number): number {
  if (r === 0) return -(fv + pmt * n);
  const p = Math.pow(1 + r, n);
  return -(fv + (pmt * (1 + r * type) * (p - 1)) / r) / p;
}

function ipmtCalc(r: number, per: number, n: number, pv: number, fv: number, type: number): number {
  const pmt = pmtCalc(r, n, pv, fv, type);
  let ip: number;
  if (per === 1) ip = type ? 0 : -pv;
  else ip = type ? fvCalc(r, per - 2, pmt, pv, 1) - pmt : fvCalc(r, per - 1, pmt, pv, 0);
  return ip * r;
}

function ddbCalc(cost: number, salvage: number, life: number, period: number, factor: number): number {
  let rate = factor / life, oldV: number;
  if (rate >= 1) { rate = 1; oldV = period === 1 ? cost : 0; } else oldV = cost * Math.pow(1 - rate, period - 1);
  const newV = cost * Math.pow(1 - rate, period);
  const d = newV < salvage ? oldV - salvage : oldV - newV;
  return d < 0 ? 0 : d;
}

function interVdb(cost: number, salvage: number, life: number, life1: number, period: number, factor: number): number {
  let vdb = 0;
  const intEnd = Math.ceil(period - 1e-12);
  let sln = 0, salvageValue = cost - salvage, nowSln = false;
  for (let i = 1; i <= intEnd; i++) {
    let term: number;
    if (!nowSln) {
      const ddb = ddbCalc(cost, salvage, life, i, factor);
      sln = salvageValue / (life1 - (i - 1));
      if (sln > ddb) { term = sln; nowSln = true; } else { term = ddb; salvageValue -= ddb; }
    } else term = sln;
    if (i === intEnd) term *= period + 1 - intEnd;
    vdb += term;
  }
  return vdb;
}

const approxEq = (a: number, b: number): boolean => Math.abs(a - b) <= 1e-12 * Math.max(1, Math.abs(a), Math.abs(b));

/** Numbers from a cash-flow argument (references/arrays: numbers only; scalars coerced). */
function flows(v: Value, ctx: Ctx): number[] {
  return numbers([v], ctx);
}

/** Values paired with dates for XNPV/XIRR: every cell must be numeric. */
function pairedFlows(vals: Value, dates: Value, ctx: Ctx): { v: number[]; d: number[] } {
  const v: number[] = [], d: number[] = [];
  for (const { v: x } of iterValues(vals, ctx)) { if (x instanceof CellError) fail(x); if (typeof x !== 'number') return fail(ERR.value); v.push(x); }
  for (const { v: x } of iterValues(dates, ctx)) {
    if (x instanceof CellError) fail(x);
    const n = typeof x === 'number' ? x : num(x);
    if (n < 0) fail(ERR.num);
    d.push(Math.trunc(n));
  }
  if (v.length !== d.length || v.length < 2) fail(ERR.num);
  return { v, d };
}

const roundHalfAway = (x: number): number => Math.sign(x) * Math.round(Math.abs(x));

// --------------------------------------------------------------------------------------------
// Odd-period bonds

interface OddFirst { C: number; f: number; N: number; Nq: number; dsc: number; sumDC: number; sumA: number }

function oddFirstParts(cal: Cal, settle: number, mat: number, issue: number, first: number, rate: number, freq: number, basis: number): OddFirst {
  if (!(mat >= first && first > settle && settle >= issue)) fail(ERR.num);
  const m = 12 / freq;
  const fc = BDate.of(cal, first, basis);
  const iss = BDate.of(cal, issue, basis);
  const st = BDate.of(cal, settle, basis);
  const qs: BDate[] = [fc];
  for (let k = 1; ; k++) {
    const q = fc.clone().addMonths(-k * m);
    qs.push(q);
    if (!q.gt(iss) || k > 5000) break;
  }
  const NC = qs.length - 1;
  const len = (a: BDate, b: BDate) => (basis === 1 ? BDate.diff(a, b) : yearDays(basis) / freq);
  let sumDC = 0, sumA = 0;
  for (let i = 1; i <= NC; i++) {
    const start = qs[i], end = qs[i - 1];
    const NL = len(start, end);
    const s0 = start.lt(iss) ? iss : start;
    sumDC += BDate.diff(s0, end) / NL;
    if (st.gt(s0)) sumA += BDate.diff(s0, st.lt(end) ? st : end) / NL;
  }
  // Next quasi-coupon date after settlement.
  let j = 0;
  for (let k = NC; k >= 0; k--) if (qs[k].gt(st)) { j = k; break; }
  const E = len(qs[j + 1], qs[j]);
  const dsc = BDate.diff(st, qs[j]) / E;
  let N = 0;
  if (mat > first) N = coupnum(coup(cal, first, mat, freq, basis));
  return { C: (100 * rate) / freq, f: freq, N, Nq: j, dsc, sumDC, sumA };
}

function oddFirstPrice(o: OddFirst, yld: number, red: number): number {
  const base = 1 + yld / o.f;
  const t0 = o.Nq + o.dsc;
  let p = red / Math.pow(base, o.N + t0) + (o.C * o.sumDC) / Math.pow(base, t0);
  for (let k = 1; k <= o.N; k++) p += o.C / Math.pow(base, k + t0);
  return p - o.C * o.sumA;
}

interface OddLast { C: number; f: number; sumDC: number; sumA: number; sumDSC: number }

function oddLastParts(cal: Cal, settle: number, mat: number, last: number, rate: number, freq: number, basis: number): OddLast {
  if (!(mat > settle && settle > last)) fail(ERR.num);
  const m = 12 / freq;
  const li = BDate.of(cal, last, basis), md = BDate.of(cal, mat, basis), st = BDate.of(cal, settle, basis);
  const qs: BDate[] = [li];
  for (let k = 1; ; k++) {
    const q = li.clone().addMonths(k * m);
    qs.push(q);
    if (!q.lt(md) || k > 5000) break;
  }
  const len = (a: BDate, b: BDate) => (basis === 1 ? BDate.diff(a, b) : yearDays(basis) / freq);
  let sumDC = 0, sumA = 0, sumDSC = 0;
  for (let i = 1; i < qs.length; i++) {
    const start = qs[i - 1], end = qs[i];
    const NL = len(start, end);
    const e = end.gt(md) ? md : end;
    sumDC += BDate.diff(start, e) / NL;
    if (st.gt(start)) sumA += BDate.diff(start, st.lt(e) ? st : e) / NL;
    if (st.lt(e)) sumDSC += BDate.diff(st.gt(start) ? st : start, e) / NL;
  }
  return { C: (100 * rate) / freq, f: freq, sumDC, sumA, sumDSC };
}

// --------------------------------------------------------------------------------------------
// Definitions

const F = 'Financial' as const;
const def = (d: Omit<FnDef, 'category'>): FnDef => ({ ...d, category: F });

defFn([
  def({
    name: 'ACCRINT', min: 6, max: 8, desc: 'Accrued interest of a security that pays periodic interest.',
    params: ['issue', 'first_interest', 'settlement', 'rate', 'par', 'frequency', '[basis]', '[calc_method]'],
    impl: (a, ctx) => {
      const cal = calOf(ctx);
      const issue = dateArg(a[0]), first = dateArg(a[1]), settle = dateArg(a[2]);
      const rate = num(S(a[3])), par = num(S(a[4])), freq = freqArg(a[5]), basis = basisArg(a, 6);
      const calc = optBool(a, 7, true);
      if (rate <= 0 || par <= 0 || issue >= settle) fail(ERR.num);
      const start = !calc && settle > first ? first : issue;
      if (basis !== 1) {
        const days = basis === 0 || basis === 4 ? BDate.diff(BDate.of(cal, start, basis), BDate.of(cal, settle, basis)) : settle - start;
        return (par * rate * days) / yearDays(basis);
      }
      // Actual/actual: sum accrued days over each quasi-coupon period, normalised by its length.
      const m = 12 / freq;
      const fc = BDate.of(cal, first, basis);
      const s0 = BDate.of(cal, start, basis), s1 = BDate.of(cal, settle, basis);
      // Find the quasi-coupon period containing the start date.
      let k = 0;
      const q = (i: number) => fc.clone().addMonths(i * m);
      while (q(k).gt(s0)) k--;
      while (!q(k + 1).gt(s0)) k++;
      let sum = 0;
      for (let i = k; i < k + 100000; i++) {
        const a0 = q(i), a1 = q(i + 1);
        if (!a0.lt(s1)) break;
        const from = a0.lt(s0) ? s0 : a0, to = a1.lt(s1) ? a1 : s1;
        sum += BDate.diff(from, to) / BDate.diff(a0, a1);
      }
      return (par * rate * sum) / freq;
    },
  }),
  def({
    name: 'ACCRINTM', min: 4, max: 5, desc: 'Accrued interest of a security that pays interest at maturity.',
    params: ['issue', 'settlement', 'rate', 'par', '[basis]'],
    impl: (a, ctx) => {
      const issue = dateArg(a[0]), settle = dateArg(a[1]), rate = num(S(a[2])), par = num(S(a[3])), basis = basisArg(a, 4);
      if (rate <= 0 || par <= 0 || issue >= settle) fail(ERR.num);
      return par * rate * yearFrac(calOf(ctx), issue, settle, basis);
    },
  }),
  def({
    name: 'AMORDEGRC', min: 6, max: 7, desc: 'Depreciation for an accounting period using a degressive coefficient (French accounting system).',
    params: ['cost', 'date_purchased', 'first_period', 'salvage', 'period', 'rate', '[basis]'],
    impl: (a, ctx) => {
      let cost = num(S(a[0]));
      const bought = dateArg(a[1]), firstPer = dateArg(a[2]), salvage = num(S(a[3])), per = Math.trunc(num(S(a[4])));
      let rate = num(S(a[5]));
      const basis = basisArg(a, 6);
      if (basis === 2 || rate <= 0 || per < 0 || cost < 0 || salvage < 0 || salvage > cost || bought > firstPer) fail(ERR.num);
      const usePer = 1 / rate;
      if ((usePer > 0 && usePer < 1) || (usePer > 1 && usePer < 2) || (usePer > 2 && usePer < 3) || (usePer > 4 && usePer < 5)) fail(ERR.num);
      const coeff = usePer < 3 ? 1 : usePer < 5 ? 1.5 : usePer <= 6 ? 2 : 2.5;
      rate *= coeff;
      let nRate = roundHalfAway(yearFrac(calOf(ctx), bought, firstPer, basis) * rate * cost);
      cost -= nRate;
      let rest = cost - salvage;
      for (let n = 0; n < per; n++) {
        nRate = roundHalfAway(rate * cost);
        rest -= nRate;
        if (rest < 0) return per - n <= 1 ? roundHalfAway(cost * 0.5) : 0;
        cost -= nRate;
      }
      return nRate;
    },
  }),
  def({
    name: 'AMORLINC', min: 6, max: 7, desc: 'Linear depreciation for an accounting period (French accounting system).',
    params: ['cost', 'date_purchased', 'first_period', 'salvage', 'period', 'rate', '[basis]'],
    impl: (a, ctx) => {
      const cost = num(S(a[0])), bought = dateArg(a[1]), firstPer = dateArg(a[2]), salvage = num(S(a[3]));
      const per = Math.trunc(num(S(a[4]))), rate = num(S(a[5])), basis = basisArg(a, 6);
      if (basis === 2 || rate <= 0 || per < 0 || cost < 0 || salvage < 0 || salvage > cost || bought > firstPer) fail(ERR.num);
      const one = cost * rate;
      const delta = cost - salvage;
      const r0 = yearFrac(calOf(ctx), bought, firstPer, basis) * rate * cost;
      const full = Math.trunc((cost - salvage - r0) / one);
      let res = 0;
      if (per === 0) res = r0;
      else if (per <= full) res = one;
      else if (per === full + 1) res = delta - one * full - r0;
      return res > 0 ? res : 0;
    },
  }),
  ...(['COUPDAYBS', 'COUPDAYS', 'COUPDAYSNC', 'COUPNCD', 'COUPNUM', 'COUPPCD'] as const).map((name) => def({
    name, min: 3, max: 4,
    desc: {
      COUPDAYBS: 'Days from the start of the coupon period to the settlement date.',
      COUPDAYS: 'Days in the coupon period that contains the settlement date.',
      COUPDAYSNC: 'Days from the settlement date to the next coupon date.',
      COUPNCD: 'Next coupon date after the settlement date.',
      COUPNUM: 'Number of coupons payable between settlement and maturity.',
      COUPPCD: 'Previous coupon date on or before the settlement date.',
    }[name],
    params: ['settlement', 'maturity', 'frequency', '[basis]'],
    impl: (a, ctx) => {
      const c = coup(calOf(ctx), dateArg(a[0]), dateArg(a[1]), freqArg(a[2]), basisArg(a, 3));
      switch (name) {
        case 'COUPDAYBS': return coupdaybs(c);
        case 'COUPDAYS': return coupdays(c);
        case 'COUPDAYSNC': return coupdaysnc(c);
        case 'COUPNCD': return coupncd(c).serial(c.cal);
        case 'COUPNUM': return coupnum(c);
        default: return couppcd(c).serial(c.cal);
      }
    },
  })),
  def({
    name: 'CUMIPMT', min: 6, max: 6, desc: 'Cumulative interest paid on a loan between two periods.',
    params: ['rate', 'nper', 'pv', 'start_period', 'end_period', 'type'],
    impl: (a) => {
      const r = num(S(a[0])), n = num(S(a[1])), pv = num(S(a[2]));
      let start = Math.trunc(num(S(a[3])));
      const end = Math.trunc(num(S(a[4]))), type = num(S(a[5]));
      if (start < 1 || end < start || r <= 0 || end > n || n <= 0 || pv <= 0 || (type !== 0 && type !== 1)) fail(ERR.num);
      const pmt = pmtCalc(r, n, pv, 0, type);
      let ip = 0;
      if (start === 1) { if (type === 0) ip = -pv; start++; }
      for (let i = start; i <= end; i++) ip += type ? fvCalc(r, i - 2, pmt, pv, 1) - pmt : fvCalc(r, i - 1, pmt, pv, 0);
      return ip * r;
    },
  }),
  def({
    name: 'CUMPRINC', min: 6, max: 6, desc: 'Cumulative principal paid on a loan between two periods.',
    params: ['rate', 'nper', 'pv', 'start_period', 'end_period', 'type'],
    impl: (a) => {
      const r = num(S(a[0])), n = num(S(a[1])), pv = num(S(a[2]));
      let start = Math.trunc(num(S(a[3])));
      const end = Math.trunc(num(S(a[4]))), type = num(S(a[5]));
      if (start < 1 || end < start || r <= 0 || end > n || n <= 0 || pv <= 0 || (type !== 0 && type !== 1)) fail(ERR.num);
      const pmt = pmtCalc(r, n, pv, 0, type);
      let pp = 0;
      if (start === 1) { pp = type === 0 ? pmt + pv * r : pmt; start++; }
      for (let i = start; i <= end; i++) pp += type ? pmt - (fvCalc(r, i - 2, pmt, pv, 1) - pmt) * r : pmt - fvCalc(r, i - 1, pmt, pv, 0) * r;
      return pp;
    },
  }),
  def({
    name: 'DB', min: 4, max: 5, desc: 'Depreciation for a period using the fixed-declining balance method.',
    params: ['cost', 'salvage', 'life', 'period', '[month]'],
    impl: (a) => {
      const cost = num(S(a[0])), salvage = num(S(a[1])), life = num(S(a[2])), period = num(S(a[3]));
      const month = Math.trunc(optNum(a, 4, 12));
      if (month < 1 || month > 12 || life > 1200 || salvage < 0 || period > life + 1 || salvage > cost || cost <= 0 || life <= 0 || period <= 0) fail(ERR.num);
      if (month === 12 && period > life) fail(ERR.num);
      let off = 1 - Math.pow(salvage / cost, 1 / life);
      off = Math.floor(off * 1000 + 0.5) / 1000;
      const first = (cost * off * month) / 12;
      if (Math.floor(period) === 1) return first;
      let sum = first, d = 0;
      const max = Math.floor(Math.min(life, period));
      for (let i = 2; i <= max; i++) { d = (cost - sum) * off; sum += d; }
      if (period > life) d = ((cost - sum) * off * (12 - month)) / 12;
      return d;
    },
  }),
  def({
    name: 'DDB', min: 4, max: 5, desc: 'Depreciation for a period using the double-declining balance method or another factor.',
    params: ['cost', 'salvage', 'life', 'period', '[factor]'],
    impl: (a) => {
      const cost = num(S(a[0])), salvage = num(S(a[1])), life = num(S(a[2])), period = num(S(a[3])), factor = optNum(a, 4, 2);
      if (cost < 0 || salvage < 0 || factor <= 0 || salvage > cost || period <= 0 || life <= 0 || period > life) fail(ERR.num);
      return ddbCalc(cost, salvage, life, period, factor);
    },
  }),
  def({
    name: 'DISC', min: 4, max: 5, desc: 'Discount rate of a security.',
    params: ['settlement', 'maturity', 'pr', 'redemption', '[basis]'],
    impl: (a, ctx) => {
      const s = dateArg(a[0]), m = dateArg(a[1]), pr = num(S(a[2])), red = num(S(a[3])), basis = basisArg(a, 4);
      if (pr <= 0 || red <= 0 || s >= m) fail(ERR.num);
      return (red - pr) / red / yearFrac(calOf(ctx), s, m, basis);
    },
  }),
  def({
    name: 'DOLLARDE', min: 2, max: 2, desc: 'Converts a price written as a fraction (e.g. 1.02 for 1 2/16) to a decimal number.',
    params: ['fractional_dollar', 'fraction'],
    impl: (a) => {
      const x = num(S(a[0])), f = Math.trunc(num(S(a[1])));
      if (f < 0) fail(ERR.num);
      if (f === 0) fail(ERR.div0);
      const i = Math.trunc(x);
      return i + ((x - i) * Math.pow(10, Math.ceil(Math.log10(f)))) / f;
    },
  }),
  def({
    name: 'DOLLARFR', min: 2, max: 2, desc: 'Converts a decimal price to a fractional notation (e.g. 1.125 to 1.02 in sixteenths).',
    params: ['decimal_dollar', 'fraction'],
    impl: (a) => {
      const x = num(S(a[0])), f = Math.trunc(num(S(a[1])));
      if (f < 0) fail(ERR.num);
      if (f === 0) fail(ERR.div0);
      const i = Math.trunc(x);
      return i + ((x - i) * f) / Math.pow(10, Math.ceil(Math.log10(f)));
    },
  }),
  def({
    name: 'DURATION', min: 5, max: 6, desc: 'Macaulay duration of a security that pays periodic interest.',
    params: ['settlement', 'maturity', 'coupon', 'yld', 'frequency', '[basis]'],
    impl: (a, ctx) => {
      const s = dateArg(a[0]), m = dateArg(a[1]), cp = num(S(a[2])), y = num(S(a[3])), f = freqArg(a[4]), b = basisArg(a, 5);
      if (cp < 0 || y < 0) fail(ERR.num);
      return bondDuration(coup(calOf(ctx), s, m, f, b), cp, y);
    },
  }),
  def({
    name: 'EFFECT', min: 2, max: 2, desc: 'Effective annual interest rate from a nominal rate and compounding periods per year.',
    params: ['nominal_rate', 'npery'],
    impl: (a) => {
      const r = num(S(a[0])), n = Math.trunc(num(S(a[1])));
      if (r <= 0 || n < 1) fail(ERR.num);
      return Math.pow(1 + r / n, n) - 1;
    },
  }),
  def({
    name: 'FV', min: 3, max: 5, desc: 'Future value of an investment with constant payments and rate.',
    params: ['rate', 'nper', 'pmt', '[pv]', '[type]'],
    impl: (a) => fvCalc(num(S(a[0])), num(S(a[1])), num(S(a[2])), optNum(a, 3, 0), typeArg(a, 4)),
  }),
  def({
    name: 'FVSCHEDULE', min: 2, max: 2, arrays: [1], desc: 'Future value of a principal after applying a series of interest rates.',
    params: ['principal', 'schedule'],
    impl: (a, ctx) => {
      let p = num(S(a[0]));
      for (const { v } of iterValues(a[1], ctx)) {
        if (v instanceof CellError) fail(v);
        if (v == null) continue;
        if (typeof v !== 'number') return fail(ERR.value);
        p *= 1 + v;
      }
      return p;
    },
  }),
  def({
    name: 'INTRATE', min: 4, max: 5, desc: 'Interest rate for a fully invested security.',
    params: ['settlement', 'maturity', 'investment', 'redemption', '[basis]'],
    impl: (a, ctx) => {
      const s = dateArg(a[0]), m = dateArg(a[1]), inv = num(S(a[2])), red = num(S(a[3])), basis = basisArg(a, 4);
      if (inv <= 0 || red <= 0 || s >= m) fail(ERR.num);
      return (red - inv) / inv / yearFrac(calOf(ctx), s, m, basis);
    },
  }),
  def({
    name: 'IPMT', min: 4, max: 6, desc: 'Interest portion of a payment for a given period.',
    params: ['rate', 'per', 'nper', 'pv', '[fv]', '[type]'],
    impl: (a) => {
      const r = num(S(a[0])), per = num(S(a[1])), n = num(S(a[2])), pv = num(S(a[3]));
      if (per < 1 || per > n) fail(ERR.num);
      return ipmtCalc(r, per, n, pv, optNum(a, 4, 0), typeArg(a, 5));
    },
  }),
  def({
    name: 'IRR', min: 1, max: 2, arrays: [0], desc: 'Internal rate of return of a series of periodic cash flows.',
    params: ['values', '[guess]'],
    impl: (a, ctx) => {
      const v = flows(a[0], ctx);
      const guess = optNum(a, 1, 0.1);
      if (!v.some((x) => x > 0) || !v.some((x) => x < 0)) fail(ERR.num);
      const f = (r: number) => { let s = 0; for (let i = 0; i < v.length; i++) s += v[i] / Math.pow(1 + r, i); return s; };
      const df = (r: number) => { let s = 0; for (let i = 1; i < v.length; i++) s -= (i * v[i]) / Math.pow(1 + r, i + 1); return s; };
      return fin(solve(f, guess, -1, 1e10, df));
    },
  }),
  def({
    name: 'ISPMT', min: 4, max: 4, desc: 'Interest paid in a period of a loan with even principal payments.',
    params: ['rate', 'per', 'nper', 'pv'],
    impl: (a) => {
      const r = num(S(a[0])), per = num(S(a[1])), n = num(S(a[2])), pv = num(S(a[3]));
      if (n === 0) fail(ERR.div0);
      return pv * r * (per / n - 1);
    },
  }),
  def({
    name: 'MDURATION', min: 5, max: 6, desc: 'Modified duration of a security with an assumed par value of 100.',
    params: ['settlement', 'maturity', 'coupon', 'yld', 'frequency', '[basis]'],
    impl: (a, ctx) => {
      const s = dateArg(a[0]), m = dateArg(a[1]), cp = num(S(a[2])), y = num(S(a[3])), f = freqArg(a[4]), b = basisArg(a, 5);
      if (cp < 0 || y < 0) fail(ERR.num);
      return bondDuration(coup(calOf(ctx), s, m, f, b), cp, y) / (1 + y / f);
    },
  }),
  def({
    name: 'MIRR', min: 3, max: 3, arrays: [0], desc: 'Modified internal rate of return using separate finance and reinvestment rates.',
    params: ['values', 'finance_rate', 'reinvest_rate'],
    impl: (a, ctx) => {
      const v = flows(a[0], ctx), fr = num(S(a[1])), rr = num(S(a[2]));
      const n = v.length;
      let npvPos = 0, npvNeg = 0;
      for (let i = 0; i < n; i++) {
        if (v[i] > 0) npvPos += v[i] / Math.pow(1 + rr, i + 1);
        else if (v[i] < 0) npvNeg += v[i] / Math.pow(1 + fr, i + 1);
      }
      if (npvPos === 0 || npvNeg === 0 || n < 2) fail(ERR.div0);
      return fin(Math.pow((-npvPos * Math.pow(1 + rr, n)) / (npvNeg * (1 + fr)), 1 / (n - 1)) - 1);
    },
  }),
  def({
    name: 'NOMINAL', min: 2, max: 2, desc: 'Nominal annual interest rate from an effective rate and compounding periods per year.',
    params: ['effect_rate', 'npery'],
    impl: (a) => {
      const r = num(S(a[0])), n = Math.trunc(num(S(a[1])));
      if (r <= 0 || n < 1) fail(ERR.num);
      return (Math.pow(1 + r, 1 / n) - 1) * n;
    },
  }),
  def({
    name: 'NPER', min: 3, max: 5, desc: 'Number of periods for an investment with constant payments and rate.',
    params: ['rate', 'pmt', 'pv', '[fv]', '[type]'],
    impl: (a) => {
      const r = num(S(a[0])), pmt = num(S(a[1])), pv = num(S(a[2])), fv = optNum(a, 3, 0), t = typeArg(a, 4);
      if (r === 0) { if (pmt === 0) fail(ERR.num); return -(pv + fv) / pmt; }
      const k = pmt * (1 + r * t);
      return fin(Math.log((k - fv * r) / (k + pv * r)) / Math.log(1 + r));
    },
  }),
  def({
    name: 'NPV', min: 2, max: 255, arrays: [1, -1], desc: 'Net present value of periodic cash flows at a discount rate.',
    params: ['rate', 'value1', '[value2]', '...'],
    impl: (a, ctx) => {
      const r = num(S(a[0]));
      if (r === -1) fail(ERR.div0);
      const v = numbers(a.slice(1), ctx);
      let s = 0;
      for (let i = 0; i < v.length; i++) s += v[i] / Math.pow(1 + r, i + 1);
      return s;
    },
  }),
  def({
    name: 'ODDFPRICE', min: 8, max: 9, desc: 'Price per 100 face value of a security with an odd (short or long) first period.',
    params: ['settlement', 'maturity', 'issue', 'first_coupon', 'rate', 'yld', 'redemption', 'frequency', '[basis]'],
    impl: (a, ctx) => {
      const s = dateArg(a[0]), m = dateArg(a[1]), iss = dateArg(a[2]), fc = dateArg(a[3]);
      const rate = num(S(a[4])), y = num(S(a[5])), red = num(S(a[6])), f = freqArg(a[7]), b = basisArg(a, 8);
      if (rate < 0 || y < 0 || red <= 0) fail(ERR.num);
      return oddFirstPrice(oddFirstParts(calOf(ctx), s, m, iss, fc, rate, f, b), y, red);
    },
  }),
  def({
    name: 'ODDFYIELD', min: 8, max: 9, desc: 'Yield of a security with an odd (short or long) first period.',
    params: ['settlement', 'maturity', 'issue', 'first_coupon', 'rate', 'pr', 'redemption', 'frequency', '[basis]'],
    impl: (a, ctx) => {
      const s = dateArg(a[0]), m = dateArg(a[1]), iss = dateArg(a[2]), fc = dateArg(a[3]);
      const rate = num(S(a[4])), pr = num(S(a[5])), red = num(S(a[6])), f = freqArg(a[7]), b = basisArg(a, 8);
      if (rate < 0 || pr <= 0 || red <= 0) fail(ERR.num);
      const o = oddFirstParts(calOf(ctx), s, m, iss, fc, rate, f, b);
      return fin(solve((y) => oddFirstPrice(o, y, red) - pr, rate || 0.05, -f + 1e-10, 1e6));
    },
  }),
  def({
    name: 'ODDLPRICE', min: 7, max: 8, desc: 'Price per 100 face value of a security with an odd last period.',
    params: ['settlement', 'maturity', 'last_interest', 'rate', 'yld', 'redemption', 'frequency', '[basis]'],
    impl: (a, ctx) => {
      const s = dateArg(a[0]), m = dateArg(a[1]), li = dateArg(a[2]);
      const rate = num(S(a[3])), y = num(S(a[4])), red = num(S(a[5])), f = freqArg(a[6]), b = basisArg(a, 7);
      if (rate < 0 || y < 0 || red <= 0) fail(ERR.num);
      const o = oddLastParts(calOf(ctx), s, m, li, rate, f, b);
      return (red + o.C * o.sumDC) / (1 + o.sumDSC * (y / f)) - o.C * o.sumA;
    },
  }),
  def({
    name: 'ODDLYIELD', min: 7, max: 8, desc: 'Yield of a security with an odd last period.',
    params: ['settlement', 'maturity', 'last_interest', 'rate', 'pr', 'redemption', 'frequency', '[basis]'],
    impl: (a, ctx) => {
      const s = dateArg(a[0]), m = dateArg(a[1]), li = dateArg(a[2]);
      const rate = num(S(a[3])), pr = num(S(a[4])), red = num(S(a[5])), f = freqArg(a[6]), b = basisArg(a, 7);
      if (rate < 0 || pr <= 0 || red <= 0) fail(ERR.num);
      const o = oddLastParts(calOf(ctx), s, m, li, rate, f, b);
      const p = pr + o.C * o.sumA;
      return ((red + o.C * o.sumDC - p) / p) * (f / o.sumDSC);
    },
  }),
  def({
    name: 'PDURATION', min: 3, max: 3, desc: 'Number of periods an investment needs to reach a target value.',
    params: ['rate', 'pv', 'fv'],
    impl: (a) => {
      const r = num(S(a[0])), pv = num(S(a[1])), fv = num(S(a[2]));
      if (r <= 0 || pv <= 0 || fv <= 0) fail(ERR.num);
      return (Math.log(fv) - Math.log(pv)) / Math.log1p(r);
    },
  }),
  def({
    name: 'PMT', min: 3, max: 5, desc: 'Periodic payment for a loan with constant payments and rate.',
    params: ['rate', 'nper', 'pv', '[fv]', '[type]'],
    impl: (a) => {
      const r = num(S(a[0])), n = num(S(a[1]));
      if (n === 0) fail(ERR.num);
      return pmtCalc(r, n, num(S(a[2])), optNum(a, 3, 0), typeArg(a, 4));
    },
  }),
  def({
    name: 'PPMT', min: 4, max: 6, desc: 'Principal portion of a payment for a given period.',
    params: ['rate', 'per', 'nper', 'pv', '[fv]', '[type]'],
    impl: (a) => {
      const r = num(S(a[0])), per = num(S(a[1])), n = num(S(a[2])), pv = num(S(a[3])), fv = optNum(a, 4, 0), t = typeArg(a, 5);
      if (per < 1 || per > n) fail(ERR.num);
      return pmtCalc(r, n, pv, fv, t) - ipmtCalc(r, per, n, pv, fv, t);
    },
  }),
  def({
    name: 'PRICE', min: 6, max: 7, desc: 'Price per 100 face value of a security that pays periodic interest.',
    params: ['settlement', 'maturity', 'rate', 'yld', 'redemption', 'frequency', '[basis]'],
    impl: (a, ctx) => {
      const s = dateArg(a[0]), m = dateArg(a[1]), rate = num(S(a[2])), y = num(S(a[3])), red = num(S(a[4])), f = freqArg(a[5]), b = basisArg(a, 6);
      if (rate < 0 || y < 0 || red <= 0) fail(ERR.num);
      return bondPrice(coup(calOf(ctx), s, m, f, b), rate, y, red);
    },
  }),
  def({
    name: 'PRICEDISC', min: 4, max: 5, desc: 'Price per 100 face value of a discounted security.',
    params: ['settlement', 'maturity', 'discount', 'redemption', '[basis]'],
    impl: (a, ctx) => {
      const s = dateArg(a[0]), m = dateArg(a[1]), d = num(S(a[2])), red = num(S(a[3])), b = basisArg(a, 4);
      if (d <= 0 || red <= 0 || s >= m) fail(ERR.num);
      return red * (1 - d * yearFrac(calOf(ctx), s, m, b));
    },
  }),
  def({
    name: 'PRICEMAT', min: 5, max: 6, desc: 'Price per 100 face value of a security that pays interest at maturity.',
    params: ['settlement', 'maturity', 'issue', 'rate', 'yld', '[basis]'],
    impl: (a, ctx) => {
      const s = dateArg(a[0]), m = dateArg(a[1]), iss = dateArg(a[2]), rate = num(S(a[3])), y = num(S(a[4])), b = basisArg(a, 5);
      if (rate < 0 || y < 0 || s >= m) fail(ERR.num);
      const cal = calOf(ctx);
      const im = yearFrac(cal, iss, m, b), is = yearFrac(cal, iss, s, b), sm = yearFrac(cal, s, m, b);
      return ((1 + im * rate) / (1 + sm * y) - is * rate) * 100;
    },
  }),
  def({
    name: 'PV', min: 3, max: 5, desc: 'Present value of an investment with constant payments and rate.',
    params: ['rate', 'nper', 'pmt', '[fv]', '[type]'],
    impl: (a) => pvCalc(num(S(a[0])), num(S(a[1])), num(S(a[2])), optNum(a, 3, 0), typeArg(a, 4)),
  }),
  def({
    name: 'RATE', min: 3, max: 6, desc: 'Interest rate per period of an annuity.',
    params: ['nper', 'pmt', 'pv', '[fv]', '[type]', '[guess]'],
    impl: (a) => {
      const n = num(S(a[0])), pmt = num(S(a[1])), pv = num(S(a[2])), fv = optNum(a, 3, 0), t = typeArg(a, 4), guess = optNum(a, 5, 0.1);
      if (n <= 0) fail(ERR.num);
      const f = (r: number) => {
        if (r === 0) return pv + pmt * n + fv;
        const l = n * Math.log1p(r);
        return pv * Math.exp(l) + (pmt * (1 + r * t) * Math.expm1(l)) / r + fv;
      };
      const r = fin(solve(f, guess, -1, 1e10));
      return Math.abs(r) < 1e-15 ? 0 : r;
    },
  }),
  def({
    name: 'RECEIVED', min: 4, max: 5, desc: 'Amount received at maturity for a fully invested security.',
    params: ['settlement', 'maturity', 'investment', 'discount', '[basis]'],
    impl: (a, ctx) => {
      const s = dateArg(a[0]), m = dateArg(a[1]), inv = num(S(a[2])), d = num(S(a[3])), b = basisArg(a, 4);
      if (inv <= 0 || d <= 0 || s >= m) fail(ERR.num);
      const den = 1 - d * yearFrac(calOf(ctx), s, m, b);
      if (den <= 0) fail(ERR.num);
      return inv / den;
    },
  }),
  def({
    name: 'RRI', min: 3, max: 3, desc: 'Equivalent interest rate for the growth of an investment.',
    params: ['nper', 'pv', 'fv'],
    impl: (a) => {
      const n = num(S(a[0])), pv = num(S(a[1])), fv = num(S(a[2]));
      if (n <= 0 || pv === 0) fail(ERR.num);
      return fin(Math.pow(fv / pv, 1 / n) - 1);
    },
  }),
  def({
    name: 'SLN', min: 3, max: 3, desc: 'Straight-line depreciation of an asset for one period.',
    params: ['cost', 'salvage', 'life'],
    impl: (a) => {
      const life = num(S(a[2]));
      if (life === 0) fail(ERR.div0);
      return (num(S(a[0])) - num(S(a[1]))) / life;
    },
  }),
  def({
    name: 'SYD', min: 4, max: 4, desc: "Sum-of-years' digits depreciation of an asset for a period.",
    params: ['cost', 'salvage', 'life', 'per'],
    impl: (a) => {
      const cost = num(S(a[0])), salvage = num(S(a[1])), life = num(S(a[2])), per = num(S(a[3]));
      if (life <= 0 || per <= 0 || per > life) fail(ERR.num);
      return ((cost - salvage) * (life - per + 1) * 2) / (life * (life + 1));
    },
  }),
  def({
    name: 'TBILLEQ', min: 3, max: 3, desc: 'Bond-equivalent yield of a Treasury bill.',
    params: ['settlement', 'maturity', 'discount'],
    impl: (a, ctx) => {
      const s = dateArg(a[0]), m = dateArg(a[1]), d = num(S(a[2]));
      if (d <= 0 || s >= m || m > calOf(ctx).edate(s, 12)) fail(ERR.num);
      const dsm = m - s;
      const den = 360 - d * dsm;
      if (den === 0) fail(ERR.num);
      return (365 * d) / den;
    },
  }),
  def({
    name: 'TBILLPRICE', min: 3, max: 3, desc: 'Price per 100 face value of a Treasury bill.',
    params: ['settlement', 'maturity', 'discount'],
    impl: (a, ctx) => {
      const s = dateArg(a[0]), m = dateArg(a[1]), d = num(S(a[2]));
      if (d <= 0 || s >= m || m > calOf(ctx).edate(s, 12)) fail(ERR.num);
      const p = 100 * (1 - (d * (m - s)) / 360);
      if (p <= 0) fail(ERR.num);
      return p;
    },
  }),
  def({
    name: 'TBILLYIELD', min: 3, max: 3, desc: 'Yield of a Treasury bill.',
    params: ['settlement', 'maturity', 'pr'],
    impl: (a, ctx) => {
      const s = dateArg(a[0]), m = dateArg(a[1]), pr = num(S(a[2]));
      if (pr <= 0 || s >= m || m > calOf(ctx).edate(s, 12)) fail(ERR.num);
      return ((100 - pr) / pr) * (360 / (m - s));
    },
  }),
  def({
    name: 'VDB', min: 5, max: 7, desc: 'Depreciation over any span of periods using declining balance, switching to straight-line when larger.',
    params: ['cost', 'salvage', 'life', 'start_period', 'end_period', '[factor]', '[no_switch]'],
    impl: (a) => {
      let cost = num(S(a[0]));
      const salvage = num(S(a[1])), life = num(S(a[2])), start = num(S(a[3])), end = num(S(a[4]));
      const factor = optNum(a, 5, 2), noSwitch = optBool(a, 6, false);
      if (start < 0 || end < start || end > life || cost < 0 || salvage < 0 || salvage > cost || factor <= 0 || life <= 0) fail(ERR.num);
      const intStart = Math.floor(start + 1e-12), intEnd = Math.ceil(end - 1e-12);
      if (noSwitch) {
        let vdb = 0;
        for (let i = intStart + 1; i <= intEnd; i++) {
          let term = ddbCalc(cost, salvage, life, i, factor);
          if (i === intStart + 1) term *= Math.min(end, intStart + 1) - start;
          else if (i === intEnd) term *= end + 1 - intEnd;
          vdb += term;
        }
        return vdb;
      }
      let part = 0;
      if (!approxEq(start, intStart)) {
        const tmp = cost - interVdb(cost, salvage, life, life, intStart, factor);
        part += (start - intStart) * interVdb(tmp, salvage, life, life - intStart, 1, factor);
      }
      if (!approxEq(end, intEnd)) {
        const ts = intEnd - 1;
        const tmp = cost - interVdb(cost, salvage, life, life, ts, factor);
        part += (intEnd - end) * interVdb(tmp, salvage, life, life - ts, intEnd - ts, factor);
      }
      cost -= interVdb(cost, salvage, life, life, intStart, factor);
      return interVdb(cost, salvage, life, life - intStart, intEnd - intStart, factor) - part;
    },
  }),
  def({
    name: 'XIRR', min: 2, max: 3, arrays: [0, 1], desc: 'Internal rate of return for cash flows on specific dates.',
    params: ['values', 'dates', '[guess]'],
    impl: (a, ctx) => {
      const { v, d } = pairedFlows(a[0], a[1], ctx);
      const guess = optNum(a, 2, 0.1);
      if (d.some((x) => x < d[0])) fail(ERR.num);
      if (!v.some((x) => x > 0) || !v.some((x) => x < 0)) fail(ERR.num);
      const t = d.map((x) => (x - d[0]) / 365);
      const f = (r: number) => { let s = 0; for (let i = 0; i < v.length; i++) s += v[i] / Math.pow(1 + r, t[i]); return s; };
      const df = (r: number) => { let s = 0; for (let i = 0; i < v.length; i++) s -= (t[i] * v[i]) / Math.pow(1 + r, t[i] + 1); return s; };
      return fin(solve(f, guess, -1, 1e10, df));
    },
  }),
  def({
    name: 'XNPV', min: 3, max: 3, arrays: [1, 2], desc: 'Net present value of cash flows on specific dates.',
    params: ['rate', 'values', 'dates'],
    impl: (a, ctx) => {
      const r = num(S(a[0]));
      const { v, d } = pairedFlows(a[1], a[2], ctx);
      if (d.some((x) => x < d[0])) fail(ERR.num);
      if (r <= -1) fail(ERR.num);
      let s = 0;
      for (let i = 0; i < v.length; i++) s += v[i] / Math.pow(1 + r, (d[i] - d[0]) / 365);
      return s;
    },
  }),
  def({
    name: 'YIELD', min: 6, max: 7, desc: 'Yield of a security that pays periodic interest.',
    params: ['settlement', 'maturity', 'rate', 'pr', 'redemption', 'frequency', '[basis]'],
    impl: (a, ctx) => {
      const s = dateArg(a[0]), m = dateArg(a[1]), rate = num(S(a[2])), pr = num(S(a[3])), red = num(S(a[4])), f = freqArg(a[5]), b = basisArg(a, 6);
      if (rate < 0 || pr <= 0 || red <= 0) fail(ERR.num);
      const c = coup(calOf(ctx), s, m, f, b);
      if (coupnum(c) <= 1) {
        const E = coupdays(c), A = coupdaybs(c), DSR = coupdaysnc(c);
        const t1 = red / 100 + rate / f, t2 = pr / 100 + (A / E) * (rate / f);
        return ((t1 - t2) / t2) * ((f * E) / DSR);
      }
      return fin(solve((y) => bondPrice(c, rate, y, red) - pr, rate || 0.05, -f + 1e-10, 1e6));
    },
  }),
  def({
    name: 'YIELDDISC', min: 4, max: 5, desc: 'Annual yield of a discounted security.',
    params: ['settlement', 'maturity', 'pr', 'redemption', '[basis]'],
    impl: (a, ctx) => {
      const s = dateArg(a[0]), m = dateArg(a[1]), pr = num(S(a[2])), red = num(S(a[3])), b = basisArg(a, 4);
      if (pr <= 0 || red <= 0 || s >= m) fail(ERR.num);
      return (red / pr - 1) / yearFrac(calOf(ctx), s, m, b);
    },
  }),
  def({
    name: 'YIELDMAT', min: 5, max: 6, desc: 'Annual yield of a security that pays interest at maturity.',
    params: ['settlement', 'maturity', 'issue', 'rate', 'pr', '[basis]'],
    impl: (a, ctx) => {
      const s = dateArg(a[0]), m = dateArg(a[1]), iss = dateArg(a[2]), rate = num(S(a[3])), pr = num(S(a[4])), b = basisArg(a, 5);
      if (rate < 0 || pr <= 0 || s >= m) fail(ERR.num);
      const cal = calOf(ctx);
      const im = yearFrac(cal, iss, m, b), is = yearFrac(cal, iss, s, b), sm = yearFrac(cal, s, m, b);
      return ((1 + im * rate) / (pr / 100 + is * rate) - 1) / sm;
    },
  }),
]);
