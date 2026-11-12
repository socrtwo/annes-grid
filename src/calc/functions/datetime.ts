/** Date & Time functions. Serial numbers follow the workbook date system (1900 or 1904). */
import { defFn } from './registry';
import { fail, isErr, iterValues, num, str, int, ERR, type Ctx, type Value } from '../values';
import type { Scalar } from '../../model/types';
import { serialToParts, jsDateToSerial, parseDateTime, isLeap, daysInMonth } from '../../model/dates';

const MAX_SERIAL = 2958465; // 9999-12-31

const d1904 = (ctx: Ctx) => ctx.book.calc.date1904;
const S = (a: Value) => a as Scalar;

/** Date serial argument: coerced, truncated, must be within the valid range. */
function dateArg(v: Value, ctx: Ctx): number {
  const n = num(S(v));
  if (n < 0 || n >= MAX_SERIAL + 1 - (d1904(ctx) ? 1462 : 0)) return fail(ERR.num);
  return Math.floor(n);
}

/** Days since 1970-01-01 for a proleptic Gregorian date (month may overflow). */
function civilDays(y: number, m: number, d: number): number {
  return Math.round(Date.UTC(y, m - 1, 1) / 86400000) + d - 1;
}
const D1900 = civilDays(1899, 12, 31); // serial 0 (with the phantom 29 Feb 1900 after serial 59)
const D1904 = civilDays(1904, 1, 1);

/** Serial for y/m/d with month and day overflow following the spreadsheet calendar. */
export function ymdSerial(y: number, m: number, d: number, date1904: boolean): number {
  const total = y * 12 + (m - 1);
  const yy = Math.floor(total / 12), mm = total - yy * 12 + 1;
  if (date1904) return civilDays(yy, mm, 1) - D1904 + d - 1;
  // Month start serial; months after Feb 1900 include the phantom leap day.
  let first = civilDays(yy, mm, 1) - D1900;
  if (first > 59) first += 1;
  return first + d - 1;
}

interface YMD { y: number; m: number; d: number }

function partsOf(serial: number, ctx: Ctx): YMD {
  const s = Math.floor(serial);
  if (!d1904(ctx) && s === 0) return { y: 1900, m: 1, d: 0 };
  const p = serialToParts(s, d1904(ctx));
  return { y: p.y, m: p.m, d: p.d };
}

/** Weekday 0=Sunday..6=Saturday of a serial. */
function dow(serial: number, ctx: Ctx): number {
  const s = Math.floor(serial);
  if (d1904(ctx)) return (((s + 5) % 7) + 7) % 7; // 1904-01-01 was a Friday
  // 1900 system: serial 1 (1900-01-01) is treated as Sunday; the phantom day keeps the cycle.
  return (((s + 6) % 7) + 7) % 7;
}

function timeOfDay(serial: number): { h: number; m: number; s: number } {
  const frac = serial - Math.floor(serial);
  let secs = Math.round(frac * 86400);
  if (secs >= 86400) secs = 0;
  return { h: Math.floor(secs / 3600), m: Math.floor((secs % 3600) / 60), s: secs % 60 };
}

function textDate(v: Scalar, ctx: Ctx): number {
  if (isErr(v)) return fail(v);
  if (typeof v !== 'string') return fail(ERR.value);
  const p = parseDateTime(v, { date1904: d1904(ctx) });
  if (!p) return fail(ERR.value);
  return p.serial;
}

// --------------------------------------------------------------------------------------------
// Weekends and holidays

const WEEKEND_CODES: Record<number, number[]> = {
  1: [6, 0], 2: [0, 1], 3: [1, 2], 4: [2, 3], 5: [3, 4], 6: [4, 5], 7: [5, 6],
  11: [0], 12: [1], 13: [2], 14: [3], 15: [4], 16: [5], 17: [6],
};

/** Weekend mask indexed by JS weekday (0=Sunday). */
function weekendMask(v: Value | undefined, allowAll = false): boolean[] {
  const x = v === undefined ? 1 : S(v) === null ? 1 : S(v);
  const mask = [false, false, false, false, false, false, false];
  if (typeof x === 'string') {
    if (!/^[01]{7}$/.test(x)) return fail(ERR.value);
    // Every day a weekend: no working day can ever be reached (WORKDAY.INTL), but the count of them is simply 0.
    if (x === '1111111' && !allowAll) return fail(ERR.value);
    // String is Monday..Sunday
    for (let i = 0; i < 7; i++) if (x[i] === '1') mask[(i + 1) % 7] = true;
    return mask;
  }
  const code = int(x as Scalar);
  const days = WEEKEND_CODES[code];
  if (!days) return fail(ERR.num);
  for (const d of days) mask[d] = true;
  return mask;
}

function holidaySet(v: Value | undefined, ctx: Ctx): Set<number> {
  const out = new Set<number>();
  if (v === undefined || v === null) return out;
  for (const { v: x } of iterValues(v, ctx)) {
    if (x == null || x === '') continue;
    if (isErr(x)) fail(x);
    const n = typeof x === 'number' ? x : num(x);
    if (n < 0) fail(ERR.num);
    out.add(Math.floor(n));
  }
  return out;
}

function networkDays(start: number, end: number, mask: boolean[], hol: Set<number>, ctx: Ctx): number {
  const sign = start <= end ? 1 : -1;
  const a = Math.min(start, end), b = Math.max(start, end);
  let count = 0;
  const total = b - a + 1;
  const full = Math.floor(total / 7);
  const workPerWeek = mask.filter((x) => !x).length;
  count = full * workPerWeek;
  for (let s = a + full * 7; s <= b; s++) if (!mask[dow(s, ctx)]) count++;
  for (const h of hol) if (h >= a && h <= b && !mask[dow(h, ctx)]) count--;
  return sign * count;
}

function workday(start: number, days: number, mask: boolean[], hol: Set<number>, ctx: Ctx): number {
  if (mask.every((x) => x)) return fail(ERR.value);
  let s = start;
  const step = days >= 0 ? 1 : -1;
  let left = Math.abs(days);
  while (left > 0) {
    s += step;
    if (s < 0 || s > MAX_SERIAL) return fail(ERR.num);
    if (mask[dow(s, ctx)] || hol.has(s)) continue;
    left--;
  }
  return s;
}

// --------------------------------------------------------------------------------------------
// Week numbers

function weekNum(serial: number, type: number, ctx: Ctx): number {
  const { y } = partsOf(serial, ctx);
  if (type === 21) return isoWeek(serial, ctx);
  const startMap: Record<number, number> = { 1: 0, 17: 0, 2: 1, 11: 1, 12: 2, 13: 3, 14: 4, 15: 5, 16: 6 };
  const first = startMap[type];
  if (first === undefined) return fail(ERR.num);
  const jan1 = ymdSerial(y, 1, 1, d1904(ctx));
  const offset = (dow(jan1, ctx) - first + 7) % 7;
  return Math.floor((Math.floor(serial) - jan1 + offset) / 7) + 1;
}

function isoWeek(serial: number, ctx: Ctx): number {
  const s = Math.floor(serial);
  const wd = (dow(s, ctx) + 6) % 7; // Monday = 0
  const thursday = s - wd + 3;
  const { y } = partsOf(thursday, ctx);
  const jan1 = ymdSerial(y, 1, 1, d1904(ctx));
  return Math.floor((thursday - jan1) / 7) + 1;
}

// --------------------------------------------------------------------------------------------
// YEARFRAC / DAYS360

function days360(a: YMD, b: YMD, european: boolean): number {
  let d1 = a.d, d2 = b.d;
  if (european) {
    if (d1 === 31) d1 = 30;
    if (d2 === 31) d2 = 30;
  } else {
    const lastFeb = (p: YMD) => p.m === 2 && p.d === daysInMonth(p.y, 2);
    if (d1 === 31 || lastFeb(a)) d1 = 30;
    if (d2 === 31 && d1 >= 30) d2 = 30;
  }
  return (b.y - a.y) * 360 + (b.m - a.m) * 30 + (d2 - d1);
}

function yearFrac(s: number, e: number, basis: number, ctx: Ctx): number {
  if (s > e) [s, e] = [e, s];
  const a = partsOf(s, ctx), b = partsOf(e, ctx);
  switch (basis) {
    case 0: {
      let d1 = a.d, d2 = b.d;
      const lastFebA = a.m === 2 && a.d === daysInMonth(a.y, 2);
      const lastFebB = b.m === 2 && b.d === daysInMonth(b.y, 2);
      if (d1 === 31 && d2 === 31) { d1 = 30; d2 = 30; }
      else if (d1 === 31) d1 = 30;
      else if (d1 === 30 && d2 === 31) d2 = 30;
      else if (lastFebA && lastFebB) { d1 = 30; d2 = 30; }
      else if (lastFebA) d1 = 30;
      return ((b.y - a.y) * 360 + (b.m - a.m) * 30 + (d2 - d1)) / 360;
    }
    case 1: {
      const days = e - s;
      const within = a.y === b.y || (b.y === a.y + 1 && (a.m > b.m || (a.m === b.m && a.d >= b.d)));
      if (within) {
        let ylen = 365;
        if (a.y === b.y && isLeap(a.y)) ylen = 366;
        else if (feb29Between(a, b) || (b.m === 2 && b.d === 29)) ylen = 366;
        return days / ylen;
      }
      const years = b.y - a.y + 1;
      const total = civilDays(b.y + 1, 1, 1) - civilDays(a.y, 1, 1);
      return days / (total / years);
    }
    case 2: return (e - s) / 360;
    case 3: return (e - s) / 365;
    case 4: return days360(a, b, true) / 360;
  }
  return fail(ERR.num);
}

function feb29Between(a: YMD, b: YMD): boolean {
  const before = (p: YMD, q: YMD) => p.y < q.y || (p.y === q.y && (p.m < q.m || (p.m === q.m && p.d < q.d)));
  const mar1a = { y: a.y, m: 3, d: 1 };
  if (isLeap(a.y) && before(a, mar1a) && !before(b, mar1a)) return true;
  const mar1b = { y: b.y, m: 3, d: 1 };
  return isLeap(b.y) && !before(b, mar1b) && before(a, mar1b);
}

// --------------------------------------------------------------------------------------------

function addMonths(start: number, months: number, eom: boolean, ctx: Ctx): number {
  const p = partsOf(start, ctx);
  const total = p.y * 12 + (p.m - 1) + months;
  const y = Math.floor(total / 12), m = total - y * 12 + 1;
  if (y < (d1904(ctx) ? 1904 : 1900) || y > 9999) return fail(ERR.num);
  const dim = y === 1900 && m === 2 && !d1904(ctx) ? 29 : daysInMonth(y, m);
  const d = eom ? dim : Math.min(p.d, dim);
  const r = ymdSerial(y, m, d, d1904(ctx));
  if (r < 0 || r > MAX_SERIAL) return fail(ERR.num);
  return r;
}

defFn([
  {
    name: 'DATE', min: 3, max: 3, category: 'Date & Time', desc: 'Returns the serial number of a date from year, month and day.', params: ['year', 'month', 'day'],
    impl: ([y, m, d], ctx) => {
      let yy = int(S(y));
      const mm = int(S(m)), dd = int(S(d));
      if (yy < 0 || yy > 9999) return ERR.num;
      if (yy < 1900) yy += 1900;
      const r = ymdSerial(yy, mm, dd, d1904(ctx));
      if (r < 0 || r > MAX_SERIAL) return ERR.num;
      return r;
    },
  },
  {
    name: 'DATEDIF', min: 3, max: 3, category: 'Date & Time', desc: 'Difference between two dates in years, months or days.', params: ['start_date', 'end_date', 'unit'],
    impl: ([s, e, u], ctx) => {
      const a = dateArg(s, ctx), b = dateArg(e, ctx);
      if (a > b) return ERR.num;
      const pa = partsOf(a, ctx), pb = partsOf(b, ctx);
      const unit = str(S(u)).toUpperCase();
      let months = (pb.y - pa.y) * 12 + (pb.m - pa.m);
      if (pb.d < pa.d) months--;
      switch (unit) {
        case 'Y': return Math.floor(months / 12);
        case 'M': return months;
        case 'D': return b - a;
        case 'YM': return ((months % 12) + 12) % 12;
        case 'MD': {
          if (pb.d >= pa.d) return pb.d - pa.d;
          const pm = pb.m === 1 ? 12 : pb.m - 1, py = pb.m === 1 ? pb.y - 1 : pb.y;
          return daysInMonth(py, pm) - pa.d + pb.d;
        }
        case 'YD': {
          let y = pb.y;
          let st = ymdSerial(y, pa.m, Math.min(pa.d, daysInMonth(y, pa.m)), d1904(ctx));
          if (st > b) { y--; st = ymdSerial(y, pa.m, Math.min(pa.d, daysInMonth(y, pa.m)), d1904(ctx)); }
          return b - st;
        }
      }
      return ERR.num;
    },
  },
  {
    name: 'DATEVALUE', min: 1, max: 1, category: 'Date & Time', desc: 'Converts a date written as text to a serial number.', params: ['date_text'],
    impl: ([t], ctx) => {
      const n = textDate(S(t), ctx);
      if (n < 0 || n >= MAX_SERIAL + 1) return ERR.value;
      return Math.floor(n);
    },
  },
  { name: 'DAY', min: 1, max: 1, category: 'Date & Time', desc: 'Day of the month (1-31) of a date.', params: ['serial_number'], impl: ([v], ctx) => partsOf(dateArg(v, ctx), ctx).d },
  { name: 'MONTH', min: 1, max: 1, category: 'Date & Time', desc: 'Month (1-12) of a date.', params: ['serial_number'], impl: ([v], ctx) => partsOf(dateArg(v, ctx), ctx).m },
  { name: 'YEAR', min: 1, max: 1, category: 'Date & Time', desc: 'Year of a date.', params: ['serial_number'], impl: ([v], ctx) => partsOf(dateArg(v, ctx), ctx).y },
  {
    name: 'DAYS', min: 2, max: 2, category: 'Date & Time', desc: 'Number of days between two dates.', params: ['end_date', 'start_date'],
    impl: ([e, s], ctx) => dateArg(e, ctx) - dateArg(s, ctx),
  },
  {
    name: 'DAYS360', min: 2, max: 3, category: 'Date & Time', desc: 'Days between two dates on a 360-day year (US or European method).', params: ['start_date', 'end_date', '[method]'],
    impl: ([s, e, m], ctx) => {
      const a = dateArg(s, ctx), b = dateArg(e, ctx);
      const eu = m === undefined || m === null ? false : (typeof m === 'string' ? fail(ERR.value) : num(S(m)) !== 0);
      return days360(partsOf(a, ctx), partsOf(b, ctx), eu as boolean);
    },
  },
  {
    name: 'EDATE', min: 2, max: 2, category: 'Date & Time', desc: 'Date a number of months before or after a start date.', params: ['start_date', 'months'],
    impl: ([s, m], ctx) => addMonths(dateArg(s, ctx), int(S(m)), false, ctx),
  },
  {
    name: 'EOMONTH', min: 2, max: 2, category: 'Date & Time', desc: 'Last day of the month a number of months before or after a start date.', params: ['start_date', 'months'],
    impl: ([s, m], ctx) => addMonths(dateArg(s, ctx), int(S(m)), true, ctx),
  },
  { name: 'HOUR', min: 1, max: 1, category: 'Date & Time', desc: 'Hour (0-23) of a time value.', params: ['serial_number'], impl: ([v]) => { const n = num(S(v)); if (n < 0) return ERR.num; return timeOfDay(n).h; } },
  { name: 'MINUTE', min: 1, max: 1, category: 'Date & Time', desc: 'Minute (0-59) of a time value.', params: ['serial_number'], impl: ([v]) => { const n = num(S(v)); if (n < 0) return ERR.num; return timeOfDay(n).m; } },
  { name: 'SECOND', min: 1, max: 1, category: 'Date & Time', desc: 'Second (0-59) of a time value.', params: ['serial_number'], impl: ([v]) => { const n = num(S(v)); if (n < 0) return ERR.num; return timeOfDay(n).s; } },
  {
    name: 'ISOWEEKNUM', min: 1, max: 1, category: 'Date & Time', desc: 'ISO 8601 week number of the year for a date.', params: ['date'],
    impl: ([v], ctx) => isoWeek(dateArg(v, ctx), ctx),
  },
  {
    name: 'NETWORKDAYS', min: 2, max: 3, category: 'Date & Time', desc: 'Number of working days (Mon-Fri) between two dates, excluding holidays.', params: ['start_date', 'end_date', '[holidays]'], arrays: [2],
    impl: ([s, e, h], ctx) => networkDays(dateArg(s, ctx), dateArg(e, ctx), weekendMask(1), holidaySet(h, ctx), ctx),
  },
  {
    name: 'NETWORKDAYS.INTL', min: 2, max: 4, category: 'Date & Time', desc: 'Number of working days between two dates with custom weekend days.', params: ['start_date', 'end_date', '[weekend]', '[holidays]'], arrays: [3],
    impl: ([s, e, w, h], ctx) => networkDays(dateArg(s, ctx), dateArg(e, ctx), weekendMask(w, true), holidaySet(h, ctx), ctx),
  },
  {
    name: 'WORKDAY', min: 2, max: 3, category: 'Date & Time', desc: 'Date a number of working days before or after a start date.', params: ['start_date', 'days', '[holidays]'], arrays: [2],
    impl: ([s, d, h], ctx) => workday(dateArg(s, ctx), int(S(d)), weekendMask(1), holidaySet(h, ctx), ctx),
  },
  {
    name: 'WORKDAY.INTL', min: 2, max: 4, category: 'Date & Time', desc: 'Date a number of working days away, with custom weekend days.', params: ['start_date', 'days', '[weekend]', '[holidays]'], arrays: [3],
    impl: ([s, d, w, h], ctx) => workday(dateArg(s, ctx), int(S(d)), weekendMask(w), holidaySet(h, ctx), ctx),
  },
  {
    name: 'NOW', min: 0, max: 0, category: 'Date & Time', desc: 'Current date and time as a serial number.', params: [], volatile: true,
    impl: (_a, ctx) => { ctx.volatile(); return jsDateToSerial(ctx.now, d1904(ctx)); },
  },
  {
    name: 'TODAY', min: 0, max: 0, category: 'Date & Time', desc: 'Current date as a serial number.', params: [], volatile: true,
    impl: (_a, ctx) => { ctx.volatile(); return Math.floor(jsDateToSerial(ctx.now, d1904(ctx))); },
  },
  {
    name: 'TIME', min: 3, max: 3, category: 'Date & Time', desc: 'Returns the time fraction for an hour, minute and second.', params: ['hour', 'minute', 'second'],
    impl: ([h, m, s]) => {
      const hh = int(S(h)), mm = int(S(m)), ss = int(S(s));
      if ([hh, mm, ss].some((x) => x > 32767)) return ERR.num;
      const total = hh * 3600 + mm * 60 + ss;
      if (total < 0) return ERR.num;
      return (total % 86400) / 86400;
    },
  },
  {
    name: 'TIMEVALUE', min: 1, max: 1, category: 'Date & Time', desc: 'Converts a time written as text to a fraction of a day.', params: ['time_text'],
    impl: ([t], ctx) => { const n = textDate(S(t), ctx); return n - Math.floor(n); },
  },
  {
    name: 'WEEKDAY', min: 1, max: 2, category: 'Date & Time', desc: 'Day of the week of a date as a number.', params: ['serial_number', '[return_type]'],
    impl: ([v, t], ctx) => {
      const d = dow(dateArg(v, ctx), ctx); // 0 = Sunday
      const type = t === undefined ? 1 : int(S(t));
      switch (type) {
        case 1: case 17: return d + 1;
        case 2: case 11: return ((d + 6) % 7) + 1;
        case 3: return (d + 6) % 7;
        case 12: case 13: case 14: case 15: case 16: return ((d - (type - 10) + 7) % 7) + 1;
      }
      return ERR.num;
    },
  },
  {
    name: 'WEEKNUM', min: 1, max: 2, category: 'Date & Time', desc: 'Week number of the year for a date.', params: ['serial_number', '[return_type]'],
    impl: ([v, t], ctx) => weekNum(dateArg(v, ctx), t === undefined || t === null ? 1 : int(S(t)), ctx),
  },
  {
    name: 'YEARFRAC', min: 2, max: 3, category: 'Date & Time', desc: 'Fraction of a year between two dates under a day-count basis.', params: ['start_date', 'end_date', '[basis]'],
    impl: ([s, e, b], ctx) => {
      const basis = b === undefined || b === null ? 0 : int(S(b));
      if (basis < 0 || basis > 4) return ERR.num;
      return yearFrac(dateArg(s, ctx), dateArg(e, ctx), basis, ctx);
    },
  },
]);
