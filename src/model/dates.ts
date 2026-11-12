/**
 * Spreadsheet date serials (1900 date system, including the historical 29 Feb 1900 serial 60,
 * and the 1904 system). Serial 1 = 1900-01-01. Fractions are the time of day.
 */

const MS_DAY = 86400000;
// Serial 0 in the 1900 system corresponds to 1899-12-31; with the phantom leap day, serials >= 61
// map to Date.UTC(1899, 11, 30) + serial days.
const EPOCH_1900 = Date.UTC(1899, 11, 30);
const EPOCH_1904 = Date.UTC(1904, 0, 1);

export interface DateParts { y: number; m: number; d: number; hh: number; mm: number; ss: number; ms: number; dow: number }

export function serialToParts(serial: number, date1904 = false): DateParts {
  let s = serial;
  let ms: number;
  if (date1904) ms = EPOCH_1904 + s * MS_DAY;
  else {
    if (s < 61) {
      if (Math.floor(s) === 60) {
        const frac = s - 60;
        const t = timeParts(frac);
        return { y: 1900, m: 2, d: 29, ...t, dow: 3 };
      }
      // Serial 0 is shown as the zeroth of January 1900 (a time of day with no date), and weekdays before the
      // phantom day follow the serial, not the real calendar: serial 1 is a Sunday.
      if (s >= 0 && s < 61) {
        const day = Math.floor(s), dow = (day + 6) % 7;
        if (day === 0) return { y: 1900, m: 1, d: 0, ...timeParts(s), dow };
        const real = serialToPartsReal(s + 1);
        return { ...real, dow };
      }
      s += 1; // before the phantom day, serials are one day "late"
    }
    ms = EPOCH_1900 + s * MS_DAY;
  }
  // Round to the millisecond to avoid 23:59:59.999 artefacts.
  ms = Math.round(ms);
  const dt = new Date(ms);
  const dayStart = Date.UTC(dt.getUTCFullYear(), dt.getUTCMonth(), dt.getUTCDate());
  const t = timeParts((ms - dayStart) / MS_DAY);
  return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate(), ...t, dow: dt.getUTCDay() };
}

function serialToPartsReal(s: number): DateParts {
  const ms = Math.round(EPOCH_1900 + s * MS_DAY);
  const dt = new Date(ms);
  const dayStart = Date.UTC(dt.getUTCFullYear(), dt.getUTCMonth(), dt.getUTCDate());
  return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate(), ...timeParts((ms - dayStart) / MS_DAY), dow: dt.getUTCDay() };
}

function timeParts(frac: number): { hh: number; mm: number; ss: number; ms: number } {
  let total = Math.round(frac * MS_DAY);
  if (total >= MS_DAY) total = MS_DAY - 1;
  const hh = Math.floor(total / 3600000); total -= hh * 3600000;
  const mm = Math.floor(total / 60000); total -= mm * 60000;
  const ss = Math.floor(total / 1000); total -= ss * 1000;
  return { hh, mm, ss, ms: total };
}

/** Serial for a calendar date; month and day may overflow (DATE semantics). */
export function partsToSerial(y: number, m: number, d: number, date1904 = false): number {
  if (y >= 0 && y < 1900 && !date1904) y += 1900;
  const ms = Date.UTC(y, m - 1, 1) + (d - 1) * MS_DAY;
  if (date1904) return Math.round((ms - EPOCH_1904) / MS_DAY);
  // The 1900 date system counts a 29 February 1900 that never was (serial 60). Days counted from January or
  // February 1900 therefore run one behind the real calendar once they pass it: DATE(1900,2,29) is 60 and
  // DATE(1900,1,61) is 61. Counted back from March 1900 or later, the real day count already agrees.
  const months = y * 12 + (m - 1), ny = Math.floor(months / 12), nm = months - ny * 12 + 1;
  if (ny === 1900 && nm <= 2) return (nm === 1 ? 0 : 31) + d;
  return Math.round((ms - EPOCH_1900) / MS_DAY);
}

export function timeToSerial(h: number, m: number, s: number): number {
  return (h * 3600 + m * 60 + s) / 86400;
}

export function jsDateToSerial(dt: Date, date1904 = false): number {
  const day = partsToSerial(dt.getFullYear(), dt.getMonth() + 1, dt.getDate(), date1904);
  return day + timeToSerial(dt.getHours(), dt.getMinutes(), dt.getSeconds() + dt.getMilliseconds() / 1000);
}

export function serialToJsDate(serial: number, date1904 = false): Date {
  const p = serialToParts(serial, date1904);
  return new Date(p.y, p.m - 1, p.d, p.hh, p.mm, p.ss, p.ms);
}

export function isLeap(y: number): boolean { return (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0; }
export function daysInMonth(y: number, m: number): number { return [31, isLeap(y) ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m - 1]; }

export const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
export const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

const MONTH_LOOKUP: Record<string, number> = {};
MONTHS.forEach((m, i) => { MONTH_LOOKUP[m.toLowerCase()] = i + 1; MONTH_LOOKUP[m.slice(0, 3).toLowerCase()] = i + 1; });
MONTH_LOOKUP.sept = 9;

/**
 * Parses typed date/time text the way a spreadsheet does on entry. Returns the serial and a
 * display format suited to what was typed, or null.
 */
export function parseDateTime(text: string, opts: { dmy?: boolean; date1904?: boolean } = {}): { serial: number; format: string } | null {
  const t = text.trim();
  if (!t) return null;
  let m: RegExpExecArray | null;
  // Time only: 10:30, 10:30:15, 10:30 PM, 10 PM
  if ((m = /^(\d{1,4})(?::(\d{1,2}))?(?::(\d{1,2}(?:\.\d+)?))?\s*([ap]\.?m?\.?)?$/i.exec(t)) && (m[2] != null || m[4])) {
    let h = +m[1];
    const mi = m[2] != null ? +m[2] : 0, s = m[3] != null ? +m[3] : 0;
    const ap = m[4]?.[0]?.toLowerCase();
    if (ap) { if (h < 1 || h > 12) return null; if (ap === 'p' && h < 12) h += 12; if (ap === 'a' && h === 12) h = 0; }
    if (mi > 59 || s >= 60) return null;
    if (!ap && h > 9999) return null;
    const fmt = ap ? (m[3] != null ? 'h:mm:ss AM/PM' : 'h:mm AM/PM') : (m[3] != null ? 'h:mm:ss' : 'h:mm');
    return { serial: timeToSerial(h, mi, s), format: h >= 24 ? '[h]:mm:ss' : fmt };
  }
  // Date with optional time: 2024-03-14, 3/14/2024, 14/3/2024 (dmy), 3/14, 14-Mar-2024, Mar 14, 2024, March 2024
  const [datePart, ...rest] = t.split(/\s+(?=\d{1,2}:\d)/);
  let timeSerial = 0, timeFmt = '';
  if (rest.length) {
    const tt = parseDateTime(rest.join(' '), opts);
    if (!tt || tt.serial >= 1) return null;
    timeSerial = tt.serial; timeFmt = ' ' + tt.format;
  }
  const d = parseDatePart(datePart, opts);
  if (!d) return null;
  const serial = partsToSerial(d.y, d.m, d.d, opts.date1904) + timeSerial;
  if (serial < 0 || serial > 2958465) return null;
  return { serial, format: timeFmt ? 'm/d/yyyy h:mm' : d.format };
}

function validDate(y: number, m: number, d: number): boolean {
  // Dates before 1900 have no serial: they stay text (adding 1900 would silently make 1776 into 3676).
  if (y === 1900 && m === 2 && d === 29) return true; // the day the 1900 date system believes in (serial 60)
  return y >= 1900 && y <= 9999 && m >= 1 && m <= 12 && d >= 1 && d <= daysInMonth(y, m);
}

function fullYear(y: string): number {
  const n = +y;
  if (y.length <= 2) return n < 30 ? 2000 + n : 1900 + n;
  return n;
}

function parseDatePart(s: string, opts: { dmy?: boolean }): { y: number; m: number; d: number; format: string } | null {
  let m: RegExpExecArray | null;
  const nowY = new Date().getFullYear();
  if ((m = /^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/.exec(s))) {
    const y = +m[1], mo = +m[2], d = +m[3];
    return validDate(y, mo, d) ? { y, m: mo, d, format: 'yyyy-mm-dd' } : null;
  }
  if ((m = /^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2}|\d{4})$/.exec(s))) {
    let mo = +m[1], d = +m[2];
    if (opts.dmy) [mo, d] = [d, mo];
    const y = fullYear(m[3]);
    return validDate(y, mo, d) ? { y, m: mo, d, format: opts.dmy ? 'd/m/yyyy' : 'm/d/yyyy' } : null;
  }
  if ((m = /^(\d{1,2})[/](\d{1,2})$/.exec(s))) {
    let mo = +m[1], d = +m[2];
    if (opts.dmy) [mo, d] = [d, mo];
    if (validDate(nowY, mo, d)) return { y: nowY, m: mo, d, format: 'd-mmm' };
    // month/year: 3/24 -> Mar 2024 when not a valid day
    return null;
  }
  if ((m = /^(\d{1,2})[-\s]([A-Za-z]{3,9})\.?(?:[-\s,]+(\d{2}|\d{4}))?$/.exec(s))) {
    const mo = MONTH_LOOKUP[m[2].toLowerCase()];
    if (!mo) return null;
    const y = m[3] ? fullYear(m[3]) : nowY, d = +m[1];
    return validDate(y, mo, d) ? { y, m: mo, d, format: m[3] ? 'd-mmm-yy' : 'd-mmm' } : null;
  }
  if ((m = /^([A-Za-z]{3,9})\.?\s+(\d{1,2})(?:st|nd|rd|th)?(?:,?\s+(\d{2}|\d{4}))?$/.exec(s))) {
    const mo = MONTH_LOOKUP[m[1].toLowerCase()];
    if (!mo) return null;
    const y = m[3] ? fullYear(m[3]) : nowY, d = +m[2];
    return validDate(y, mo, d) ? { y, m: mo, d, format: m[3] ? 'mmm d, yyyy' : 'd-mmm' } : null;
  }
  if ((m = /^([A-Za-z]{3,9})[-\s,]+(\d{4})$/.exec(s))) {
    const mo = MONTH_LOOKUP[m[1].toLowerCase()];
    if (!mo) return null;
    return +m[2] >= 1900 ? { y: +m[2], m: mo, d: 1, format: 'mmm-yy' } : null;
  }
  return null;
}
