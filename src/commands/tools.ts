/**
 * Tools tab: data clean-up, transformation, comparison, analysis and generators.
 *
 * The transformations are pure exported functions (tested in tests/unit/tools.test.ts); the
 * commands apply them to the selection (clipped to the used range) as one undoable step.
 * Several tools come from other spreadsheets and data-cleaning add-ins; their heritage is noted
 * in `origin`.
 */
import { def, type MenuEntry } from './registry';
import type { App } from '../app';
import { Sheet, type Workbook } from '../model/workbook';
import type { Cell, CellStyle, Scalar, CFRule, ChartDef } from '../model/types';
import { CellError } from '../model/types';
import { MAX_COLS, MAX_ROWS, cellKey, colName, cellName, rangeName, type Range } from '../model/address';
import { selRanges, activeRange, activeCell, clipToUsed, dataRegion, hasHeaderRow, cellText, isWholeCols } from './ops';
import { parseNumberInput } from '../model/numfmt';
import { partsToSerial, daysInMonth, serialToParts, MONTHS, DAYS, timeToSerial } from '../model/dates';
import { shiftFormula } from '../calc/parser';
import { rangeRefText } from '../model/structure';
import { newId } from '../model/workbook';
import { openDialog, toast, confirmDialog, choiceDialog } from '../ui/dialog';
import { h } from '../ui/dom';
import { escHtml } from '../util/sanitize';

// ============================================================================================
// Pure text transformations

/** TRIM semantics (leading/trailing removed, inner runs collapsed to one space), also for non-breaking and other Unicode spaces. */
export function trimText(s: string): string {
  return s.replace(/[  -   　\t]/g, ' ').replace(/ {2,}/g, ' ').replace(/^ +| +$/g, '').replace(/ *\n */g, '\n');
}

/** Collapses runs of spaces inside the text but leaves the ends alone. */
export function collapseSpaces(s: string): string {
  return s.replace(/[  ]/g, ' ').replace(/ {2,}/g, ' ');
}

/** CLEAN semantics plus Unicode controls: removes characters 0-31, 127, 128-159, zero-width characters and the byte-order mark. */
export function cleanText(s: string, keepLineBreaks = false): string {
  const re = keepLineBreaks ? /[\u0000-\u0009\u000b-\u001f\u007f-\u009f​-‍⁠﻿­]/g : /[\u0000-\u001f\u007f-\u009f​-‍⁠﻿­]/g;
  return s.replace(re, '');
}

export type CaseMode = 'upper' | 'lower' | 'proper' | 'sentence' | 'toggle';

/** UPPER, lower, Proper (PROPER semantics: letter after a non-letter is capitalized), Sentence case and tOGGLE cASE. */
export function changeCase(s: string, mode: CaseMode): string {
  switch (mode) {
    case 'upper': return s.toUpperCase();
    case 'lower': return s.toLowerCase();
    case 'proper': return s.toLowerCase().replace(/(^|[^\p{L}\p{N}'’])(\p{L})/gu, (_m, a: string, b: string) => a + b.toUpperCase());
    case 'sentence': return s.toLowerCase().replace(/(^\s*|[.!?]\s+|\n\s*)(\p{L})/gu, (_m, a: string, b: string) => a + b.toUpperCase()).replace(/\bi\b/g, 'I');
    case 'toggle': return [...s].map((ch) => (ch === ch.toUpperCase() ? ch.toLowerCase() : ch.toUpperCase())).join('');
  }
}

/** Converts text that looks like a number ("1,234", " 12 ", "(45)", "12-", "$3.50", "50%", "1 234,5" is not guessed) to a number. */
export function textToNumber(s: string): number | null {
  let t = s.replace(/[   ]/g, ' ').trim();
  if (t.startsWith("'")) t = t.slice(1).trim();
  if (!t) return null;
  // Trailing minus (exported accounting systems): "123.45-"
  let neg = false;
  if (/^[^-]+-$/.test(t)) { neg = true; t = t.slice(0, -1).trim(); }
  // Spaces as thousands separators: "1 234 567"
  if (/^\d{1,3}( \d{3})+(\.\d+)?$/.test(t)) t = t.replace(/ /g, '');
  // Unicode minus
  t = t.replace(/^[−–]/, '-');
  // Currency words/codes around the number
  t = t.replace(/^(USD|EUR|GBP|CAD|AUD)\s*/i, '').replace(/\s*(USD|EUR|GBP|CAD|AUD)$/i, '');
  const p = parseNumberInput(t);
  if (!p) return null;
  return neg ? -p.value : p.value;
}

const MONTH_NAMES: Record<string, number> = {};
MONTHS.forEach((m, i) => { MONTH_NAMES[m.toLowerCase()] = i + 1; MONTH_NAMES[m.slice(0, 3).toLowerCase()] = i + 1; });
Object.assign(MONTH_NAMES, { sept: 9, janv: 1, févr: 2, fevr: 2, mars: 3, avr: 4, mai: 5, juin: 6, juil: 7, août: 8, aout: 8, déc: 12, ene: 1, abr: 4, ago: 8, dic: 12, okt: 10, dez: 12, mär: 3, maerz: 3 });

export interface DateFixOptions {
  /** Order of numeric day/month: 'auto' uses unambiguous values and falls back to mdy. */
  order?: 'auto' | 'mdy' | 'dmy' | 'ymd';
  /** Two-digit years up to this value are 20xx, above are 19xx (spreadsheet default 29). */
  pivot?: number;
  date1904?: boolean;
}

function year2(y: string, pivot: number): number {
  const n = +y;
  if (y.length > 2) return n;
  return n <= pivot ? 2000 + n : 1900 + n;
}

/**
 * Parses a date written as text in many common forms: ISO (2024-03-05), numeric with / . - or
 * spaces in M/D/Y, D/M/Y or Y/M/D order, 8-digit compact (20240305), month names ("Mar 5, 2024",
 * "5 March 2024", "5-Mar-24", "Tuesday, March 5th, 2024", "March 2024"), with an optional time.
 * Returns the date serial or null.
 */
export function parseTextDate(text: string, opts: DateFixOptions = {}): number | null {
  const pivot = opts.pivot ?? 29;
  let t = text.replace(/[  ]/g, ' ').trim().replace(/\s+/g, ' ');
  if (!t) return null;
  // Time part
  let time = 0;
  const tm = /[ T](\d{1,2}):(\d{2})(?::(\d{2})(?:\.\d+)?)?\s*([ap]\.?m\.?)?\s*(z|[+-]\d{2}:?\d{2})?$/i.exec(t);
  if (tm) {
    let hh = +tm[1];
    const ap = tm[4]?.[0].toLowerCase();
    if (ap === 'p' && hh < 12) hh += 12;
    if (ap === 'a' && hh === 12) hh = 0;
    if (hh > 23 || +tm[2] > 59) return null;
    time = timeToSerial(hh, +tm[2], +(tm[3] || 0));
    t = t.slice(0, tm.index).trim();
  }
  // Remove weekday names and ordinal suffixes
  t = t.replace(/^(sun|mon|tue|tues|wed|thu|thur|thurs|fri|sat)[a-z]*\.?,?\s+/i, '').replace(/(\d)(st|nd|rd|th)\b/gi, '$1').replace(/\bof\b/gi, ' ').replace(/\s+/g, ' ').trim();
  let y = 0, m = 0, d = 0;
  let r: RegExpExecArray | null;
  const ok = () => y >= 1 && y <= 9999 && m >= 1 && m <= 12 && d >= 1 && d <= daysInMonth(y, m);
  const order = opts.order || 'auto';
  if ((r = /^(\d{4})[-/. ](\d{1,2})[-/. ](\d{1,2})$/.exec(t))) { y = +r[1]; m = +r[2]; d = +r[3]; }
  else if ((r = /^(\d{4})(\d{2})(\d{2})$/.exec(t))) { y = +r[1]; m = +r[2]; d = +r[3]; }
  else if (order === 'ymd' && (r = /^(\d{2})[-/. ](\d{1,2})[-/. ](\d{1,2})$/.exec(t))) { y = year2(r[1], pivot); m = +r[2]; d = +r[3]; }
  else if ((r = /^(\d{1,2})[-/. ](\d{1,2})[-/. ](\d{2}|\d{4})$/.exec(t))) {
    const a = +r[1], b = +r[2];
    y = year2(r[3], pivot);
    let o = order === 'ymd' ? 'mdy' : order;
    if (o === 'auto') o = a > 12 && b <= 12 ? 'dmy' : 'mdy';
    if (o === 'dmy') { d = a; m = b; } else { m = a; d = b; }
  } else if ((r = /^(\d{1,2})[-/. ]?([a-zà-ü]{3,9})\.?[-/., ]+(\d{2}|\d{4})$/i.exec(t)) && MONTH_NAMES[r[2].toLowerCase()]) {
    d = +r[1]; m = MONTH_NAMES[r[2].toLowerCase()]; y = year2(r[3], pivot);
  } else if ((r = /^([a-zà-ü]{3,9})\.?[-/. ]?(\d{1,2}),?[-/. ]+(\d{2}|\d{4})$/i.exec(t)) && MONTH_NAMES[r[1].toLowerCase()]) {
    m = MONTH_NAMES[r[1].toLowerCase()]; d = +r[2]; y = year2(r[3], pivot);
  } else if ((r = /^([a-zà-ü]{3,9})\.?[-/., ]+(\d{4})$/i.exec(t)) && MONTH_NAMES[r[1].toLowerCase()]) {
    m = MONTH_NAMES[r[1].toLowerCase()]; d = 1; y = +r[2];
  } else if ((r = /^(\d{4})[-/. ]([a-zà-ü]{3,9})\.?[-/. ](\d{1,2})$/i.exec(t)) && MONTH_NAMES[r[2].toLowerCase()]) {
    y = +r[1]; m = MONTH_NAMES[r[2].toLowerCase()]; d = +r[3];
  } else return null;
  if (!ok()) return null;
  return partsToSerial(y, m, d, opts.date1904) + time;
}

export interface OcrOptions { joinLines?: boolean; dehyphenate?: boolean; digits?: boolean; letters?: boolean; ligatures?: boolean; quotes?: boolean }

/**
 * Typical repairs for text that came from OCR (scanned records, PDFs): ligatures, smart quotes,
 * soft hyphens, words broken across lines, single line breaks inside a paragraph, and O/0, l/1
 * confusions (letters inside numbers become digits; digits inside words become letters).
 */
export function ocrClean(s: string, o: OcrOptions = {}): string {
  const opt = { joinLines: true, dehyphenate: true, digits: true, letters: true, ligatures: true, quotes: true, ...o };
  let t = s.replace(/\r\n?/g, '\n').replace(/­/g, '');
  if (opt.ligatures) {
    t = t.replace(/ﬀ/g, 'ff').replace(/ﬁ/g, 'fi').replace(/ﬂ/g, 'fl').replace(/ﬃ/g, 'ffi').replace(/ﬄ/g, 'ffl').replace(/[ﬅﬆ]/g, 'st')
      .replace(/Ĳ/g, 'IJ').replace(/ĳ/g, 'ij').replace(/…/g, '...');
  }
  if (opt.quotes) t = t.replace(/[‘’‚‛′´`]/g, "'").replace(/[“”„‟″«»]/g, '"');
  if (opt.dehyphenate) t = t.replace(/(\p{L})[-‐]\s*\n\s*(\p{Ll})/gu, '$1$2');
  if (opt.joinLines) t = t.replace(/\n{2,}/g, '\u0000').replace(/[ \t]*\n[ \t]*/g, ' ').replace(/\u0000/g, '\n');
  if (opt.digits) {
    // Tokens that are numbers apart from look-alike letters: 1O5 -> 105, l23 -> 123, $1,2O0.00 -> $1,200.00
    t = t.replace(/(^|[\s(])([-+$€£]?)([0-9OoIl|][0-9OoIl|,.]*[0-9OoIl|])(%?)(?=$|[\s),;:])/g, (m, pre: string, sign: string, body: string, pct: string) => {
      if (!/\d/.test(body) || !/[OoIl|]/.test(body)) return m;
      const digits = (body.match(/\d/g) || []).length, letters = (body.match(/[OoIl|]/g) || []).length;
      if (letters > digits && !sign) return m;
      return pre + sign + body.replace(/[Oo]/g, '0').replace(/[Il|]/g, '1') + pct;
    });
  }
  if (opt.letters) {
    // Digits inside words: C0MPANY -> COMPANY, he1lo -> hello
    t = t.replace(/\b([A-Za-z]+)([01])([A-Za-z]+)\b/g, (_m, a: string, d: string, b: string) => {
      const upper = a.slice(-1) === a.slice(-1).toUpperCase() && b[0] === b[0].toUpperCase();
      const ch = d === '0' ? (upper ? 'O' : 'o') : (upper ? 'I' : 'l');
      return a + ch + b;
    });
  }
  return t.replace(/ {2,}/g, ' ').trim();
}

/** Removes accents and other diacritics (é -> e, ß -> ss, Ł -> L). */
export function removeAccents(s: string): string {
  return s.normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/ß/g, 'ss').replace(/Ł/g, 'L').replace(/ł/g, 'l').replace(/Ø/g, 'O').replace(/ø/g, 'o').replace(/Æ/g, 'AE').replace(/æ/g, 'ae').replace(/Œ/g, 'OE').replace(/œ/g, 'oe').replace(/Đ/g, 'D').replace(/đ/g, 'd');
}

export type SplitMode = { kind: 'delimiter'; delims: string; consecutive?: boolean; quote?: boolean } | { kind: 'positions'; at: number[] } | { kind: 'case' } | { kind: 'numberText' };

/** Splits text into parts by delimiters, fixed positions, case changes ("firstName" -> first, Name) or number/text boundaries ("12kg" -> 12, kg). */
export function splitText(s: string, mode: SplitMode): string[] {
  switch (mode.kind) {
    case 'delimiter': {
      const ds = mode.delims || ',';
      const out: string[] = [];
      let cur = '', inQ = false;
      for (let i = 0; i < s.length; i++) {
        const ch = s[i];
        if (mode.quote !== false && ch === '"') { if (inQ && s[i + 1] === '"') { cur += '"'; i++; } else inQ = !inQ; continue; }
        if (!inQ && ds.includes(ch)) {
          if (mode.consecutive && cur === '' && out.length && i > 0 && ds.includes(s[i - 1])) continue;
          out.push(cur); cur = '';
        } else cur += ch;
      }
      out.push(cur);
      return out;
    }
    case 'positions': {
      const at = [...new Set(mode.at.filter((n) => n > 0 && n < s.length))].sort((a, b) => a - b);
      const out: string[] = [];
      let p = 0;
      for (const x of at) { out.push(s.slice(p, x)); p = x; }
      out.push(s.slice(p));
      return out;
    }
    case 'case': return s.split(/(?<=\p{Ll})(?=\p{Lu})|(?<=\p{Lu})(?=\p{Lu}\p{Ll})|[\s_-]+/u).filter((x) => x !== '');
    case 'numberText': return s.match(/[-+]?\d+(?:[.,]\d+)*|[^\d\s]+(?:\s+[^\d\s]+)*/g)?.map((x) => x.trim()).filter(Boolean) || [s];
  }
}

/** Unpivots (melts) a table: the first idCols columns are kept; every other column becomes an Attribute/Value row. Blank values are skipped unless keepBlanks. */
export function unpivot(header: Scalar[], rows: Scalar[][], idCols: number, keepBlanks = false, names = ['Attribute', 'Value']): Scalar[][] {
  const out: Scalar[][] = [[...header.slice(0, idCols), names[0], names[1]]];
  for (const row of rows) {
    for (let c = idCols; c < header.length; c++) {
      const v = row[c] ?? null;
      if (!keepBlanks && (v == null || v === '')) continue;
      out.push([...row.slice(0, idCols).map((x) => x ?? null), header[c] ?? `Column${c + 1}`, v]);
    }
  }
  return out;
}

export function transpose<T>(m: T[][]): T[][] {
  const rows = m.length, cols = Math.max(0, ...m.map((r) => r.length));
  const out: T[][] = [];
  for (let c = 0; c < cols; c++) { const row: T[] = []; for (let r = 0; r < rows; r++) row.push(m[r][c]); out.push(row); }
  return out;
}

/** Fisher-Yates shuffle (returns a new array). */
export function shuffled<T>(xs: T[], rnd: () => number = Math.random): T[] {
  const a = [...xs];
  for (let i = a.length - 1; i > 0; i--) { const j = Math.floor(rnd() * (i + 1)); [a[i], a[j]] = [a[j], a[i]]; }
  return a;
}

// --------------------------------------------------------------------------------------------
// Fuzzy matching

export function levenshtein(a: string, b: string): number {
  if (a === b) return 0;
  if (!a.length) return b.length;
  if (!b.length) return a.length;
  let prev = Array.from({ length: b.length + 1 }, (_, i) => i);
  for (let i = 1; i <= a.length; i++) {
    const cur = [i];
    for (let j = 1; j <= b.length; j++) cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    prev = cur;
  }
  return prev[b.length];
}

/** 1 - distance / longer length (1 = identical). */
export function levenshteinSimilarity(a: string, b: string): number {
  const n = Math.max(a.length, b.length);
  return n ? 1 - levenshtein(a, b) / n : 1;
}

export function jaro(a: string, b: string): number {
  if (a === b) return 1;
  if (!a.length || !b.length) return 0;
  const win = Math.max(0, Math.floor(Math.max(a.length, b.length) / 2) - 1);
  const am = new Array(a.length).fill(false), bm = new Array(b.length).fill(false);
  let matches = 0;
  for (let i = 0; i < a.length; i++) {
    for (let j = Math.max(0, i - win); j < Math.min(b.length, i + win + 1); j++) {
      if (bm[j] || a[i] !== b[j]) continue;
      am[i] = bm[j] = true; matches++; break;
    }
  }
  if (!matches) return 0;
  let k = 0, trans = 0;
  for (let i = 0; i < a.length; i++) {
    if (!am[i]) continue;
    while (!bm[k]) k++;
    if (a[i] !== b[k]) trans++;
    k++;
  }
  return (matches / a.length + matches / b.length + (matches - trans / 2) / matches) / 3;
}

export function jaroWinkler(a: string, b: string, p = 0.1): number {
  const j = jaro(a, b);
  let l = 0;
  while (l < 4 && l < a.length && l < b.length && a[l] === b[l]) l++;
  return j + l * p * (1 - j);
}

export type FuzzyAlgo = 'levenshtein' | 'jaroWinkler' | 'tokens';

export function normalizeForMatch(s: string): string {
  return removeAccents(s).toLowerCase().replace(/[^\p{L}\p{N}]+/gu, ' ').trim();
}

export function fuzzyScore(a: string, b: string, algo: FuzzyAlgo = 'jaroWinkler'): number {
  const x = normalizeForMatch(a), y = normalizeForMatch(b);
  if (algo === 'levenshtein') return levenshteinSimilarity(x, y);
  if (algo === 'tokens') {
    const sx = x.split(' ').sort().join(' '), sy = y.split(' ').sort().join(' ');
    return Math.max(levenshteinSimilarity(sx, sy), jaroWinkler(sx, sy) * 0.98);
  }
  return jaroWinkler(x, y);
}

export function bestMatch(s: string, candidates: string[], algo: FuzzyAlgo = 'jaroWinkler'): { match: string; score: number; index: number } | null {
  let best: { match: string; score: number; index: number } | null = null;
  candidates.forEach((c, i) => {
    if (!c) return;
    const sc = fuzzyScore(s, c, algo);
    if (!best || sc > best.score) best = { match: c, score: sc, index: i };
  });
  return best;
}

// --------------------------------------------------------------------------------------------
// Numbers in words, Roman numerals

const ONES = ['zero', 'one', 'two', 'three', 'four', 'five', 'six', 'seven', 'eight', 'nine', 'ten', 'eleven', 'twelve', 'thirteen', 'fourteen', 'fifteen', 'sixteen', 'seventeen', 'eighteen', 'nineteen'];
const TENS = ['', '', 'twenty', 'thirty', 'forty', 'fifty', 'sixty', 'seventy', 'eighty', 'ninety'];
const SCALES = ['', 'thousand', 'million', 'billion', 'trillion', 'quadrillion'];

function under1000(n: number, and: boolean): string {
  const hnd = Math.floor(n / 100), rest = n % 100;
  const parts: string[] = [];
  if (hnd) parts.push(ONES[hnd] + ' hundred');
  if (rest) {
    const t = rest < 20 ? ONES[rest] : TENS[Math.floor(rest / 10)] + (rest % 10 ? '-' + ONES[rest % 10] : '');
    parts.push(hnd && and ? 'and ' + t : t);
  }
  return parts.join(' ');
}

/** Whole number in English words: 1234 -> "one thousand two hundred thirty-four" (and=true: "... two hundred and thirty-four"). */
export function intToWords(n: number, and = false): string {
  n = Math.floor(Math.abs(n));
  if (n === 0) return 'zero';
  const groups: number[] = [];
  while (n > 0) { groups.push(n % 1000); n = Math.floor(n / 1000); }
  if (groups.length > SCALES.length) throw new Error('Number is too large');
  const parts: string[] = [];
  for (let i = groups.length - 1; i >= 0; i--) {
    const g = groups[i];
    if (!g) continue;
    let w = under1000(g, and);
    if (and && i === 0 && g < 100 && groups.length > 1) w = 'and ' + w;
    parts.push(w + (SCALES[i] ? ' ' + SCALES[i] : ''));
  }
  return parts.join(' ');
}

const ORD_IRREG: Record<string, string> = { one: 'first', two: 'second', three: 'third', five: 'fifth', eight: 'eighth', nine: 'ninth', twelve: 'twelfth' };

export type WordsStyle = 'words' | 'check' | 'ordinal' | 'year' | 'legal';

const title = (s: string) => s.replace(/(^|[\s-])([a-z])/g, (_m, a: string, b: string) => a + b.toUpperCase()).replace(/ And /g, ' and ');

/**
 * Number in words. Styles: 'words' (one hundred twenty-three point four five), 'check'
 * (One Hundred Twenty-Three and 45/100), 'ordinal' (one hundred twenty-third), 'year'
 * (eighteen fifty, nineteen oh five, two thousand five), 'legal' (One Hundred and Twenty-Three).
 */
export function numberToWords(n: number, style: WordsStyle = 'words', currency = ''): string {
  if (!Number.isFinite(n)) throw new Error('Not a number');
  const neg = n < 0;
  const a = Math.abs(n);
  switch (style) {
    case 'check': {
      let whole = Math.floor(a), cents = Math.round((a - whole) * 100);
      if (cents === 100) { whole++; cents = 0; }
      return `${neg ? 'Minus ' : ''}${title(intToWords(whole))} and ${String(cents).padStart(2, '0')}/100${currency ? ' ' + currency : ''}`;
    }
    case 'legal': {
      let whole = Math.floor(a), cents = Math.round((a - whole) * 100);
      if (cents === 100) { whole++; cents = 0; }
      return `${neg ? 'Minus ' : ''}${title(intToWords(whole, true))}${cents ? ` and ${title(intToWords(cents))} Cents` : ''}${currency ? ' ' + currency : ''}`;
    }
    case 'ordinal': {
      const w = intToWords(Math.round(a));
      const m = /([a-z]+)$/.exec(w)!;
      const last = m[1];
      const ord = ORD_IRREG[last] || (last.endsWith('y') ? last.slice(0, -1) + 'ieth' : last + 'th');
      return (neg ? 'minus ' : '') + w.slice(0, m.index) + ord;
    }
    case 'year': {
      const y = Math.round(a);
      if (y < 1000 || y >= 10000 || (y >= 2000 && y < 2010)) return intToWords(y);
      const hi = Math.floor(y / 100), lo = y % 100;
      if (lo === 0) return `${intToWords(hi)} hundred`;
      return `${intToWords(hi)} ${lo < 10 ? 'oh ' + ONES[lo] : intToWords(lo)}`;
    }
    default: {
      const s = String(a);
      const [w, f] = s.includes('e') ? [String(Math.floor(a)), ''] : s.split('.');
      const head = intToWords(+w);
      const tail = f ? ' point ' + [...f].map((d) => ONES[+d]).join(' ') : '';
      return (neg ? 'minus ' : '') + head + tail;
    }
  }
}

const ROMAN: [number, string][] = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']];

export function toRoman(n: number): string {
  n = Math.round(n);
  if (n < 1 || n > 3999) throw new Error('Roman numerals cover 1 to 3999');
  let out = '';
  for (const [v, s] of ROMAN) while (n >= v) { out += s; n -= v; }
  return out;
}

export function fromRoman(s: string): number | null {
  const t = s.trim().toUpperCase();
  if (!/^[MDCLXVI]+$/.test(t)) return null;
  const val: Record<string, number> = { M: 1000, D: 500, C: 100, L: 50, X: 10, V: 5, I: 1 };
  let n = 0;
  for (let i = 0; i < t.length; i++) { const v = val[t[i]], nx = val[t[i + 1]] || 0; n += v < nx ? -v : v; }
  return toRoman(n) === t ? n : null;
}

// --------------------------------------------------------------------------------------------
// Extraction and statistics

export const EXTRACT_PATTERNS: Record<string, { label: string; re: RegExp }> = {
  emails: { label: 'Email addresses', re: /[\w.+-]+@[\w-]+(?:\.[\w-]+)+/g },
  urls: { label: 'Web addresses', re: /\b(?:https?:\/\/|www\.)[^\s<>"')]+/gi },
  numbers: { label: 'Numbers', re: /[-+]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?/g },
  phones: { label: 'Phone numbers', re: /(?:\+\d{1,3}[\s.-]?)?(?:\(\d{2,4}\)|\d{2,4})[\s.-]?\d{3,4}[\s.-]?\d{3,4}/g },
  hashtags: { label: 'Hashtags and mentions', re: /[#@][\p{L}\p{N}_]+/gu },
  dates: { label: 'Dates', re: /\b(?:\d{4}-\d{1,2}-\d{1,2}|\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4}|\d{1,2} (?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]* \d{4}|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]* \d{1,2},? \d{4})\b/gi },
  zip: { label: 'ZIP / postal codes', re: /\b\d{5}(?:-\d{4})?\b|\b[A-Z]\d[A-Z] ?\d[A-Z]\d\b|\b[A-Z]{1,2}\d[A-Z\d]? ?\d[A-Z]{2}\b/g },
};

export function extractMatches(s: string, re: RegExp): string[] {
  const g = new RegExp(re.source, re.flags.includes('g') ? re.flags : re.flags + 'g');
  return [...s.matchAll(g)].map((m) => m[0]);
}

export interface Describe { count: number; sum: number; mean: number; median: number; mode: number | null; stdev: number; variance: number; stderr: number; kurtosis: number | null; skewness: number | null; min: number; max: number; range: number }

/** Descriptive statistics with sample standard deviation and Excel's KURT/SKEW definitions. */
export function describe(xs: number[]): Describe {
  const n = xs.length;
  const sorted = [...xs].sort((a, b) => a - b);
  const sum = xs.reduce((a, b) => a + b, 0);
  const mean = n ? sum / n : 0;
  const ss = xs.reduce((a, x) => a + (x - mean) ** 2, 0);
  const variance = n > 1 ? ss / (n - 1) : 0;
  const sd = Math.sqrt(variance);
  const median = n ? (n % 2 ? sorted[(n - 1) / 2] : (sorted[n / 2 - 1] + sorted[n / 2]) / 2) : 0;
  const counts = new Map<number, number>();
  let mode: number | null = null, best = 1;
  for (const x of xs) { const c = (counts.get(x) || 0) + 1; counts.set(x, c); if (c > best) { best = c; mode = x; } }
  let skewness: number | null = null, kurtosis: number | null = null;
  if (n > 2 && sd > 0) skewness = (n / ((n - 1) * (n - 2))) * xs.reduce((a, x) => a + ((x - mean) / sd) ** 3, 0);
  if (n > 3 && sd > 0) {
    const s4 = xs.reduce((a, x) => a + ((x - mean) / sd) ** 4, 0);
    kurtosis = ((n * (n + 1)) / ((n - 1) * (n - 2) * (n - 3))) * s4 - (3 * (n - 1) ** 2) / ((n - 2) * (n - 3));
  }
  return { count: n, sum, mean, median, mode, stdev: sd, variance, stderr: n ? sd / Math.sqrt(n) : 0, kurtosis, skewness, min: n ? sorted[0] : 0, max: n ? sorted[n - 1] : 0, range: n ? sorted[n - 1] - sorted[0] : 0 };
}

export function pearson(xs: number[], ys: number[]): number | null {
  const n = Math.min(xs.length, ys.length);
  if (n < 2) return null;
  let mx = 0, my = 0;
  for (let i = 0; i < n; i++) { mx += xs[i]; my += ys[i]; }
  mx /= n; my /= n;
  let sxy = 0, sxx = 0, syy = 0;
  for (let i = 0; i < n; i++) { const dx = xs[i] - mx, dy = ys[i] - my; sxy += dx * dy; sxx += dx * dx; syy += dy * dy; }
  return sxx && syy ? sxy / Math.sqrt(sxx * syy) : null;
}

export interface ColumnProfile { name: string; numbers: number; text: number; booleans: number; errors: number; blanks: number; distinct: number; min: number | null; max: number | null; mean: number | null; top: { value: string; count: number }[]; type: string }

export function profileColumn(name: string, values: Scalar[]): ColumnProfile {
  const p: ColumnProfile = { name, numbers: 0, text: 0, booleans: 0, errors: 0, blanks: 0, distinct: 0, min: null, max: null, mean: null, top: [], type: 'Empty' };
  const counts = new Map<string, number>();
  let sum = 0;
  for (const v of values) {
    if (v == null || v === '') { p.blanks++; continue; }
    if (v instanceof CellError) p.errors++;
    else if (typeof v === 'number') { p.numbers++; sum += v; p.min = p.min == null ? v : Math.min(p.min, v); p.max = p.max == null ? v : Math.max(p.max, v); }
    else if (typeof v === 'boolean') p.booleans++;
    else p.text++;
    const k = v instanceof CellError ? v.code : String(v);
    counts.set(k, (counts.get(k) || 0) + 1);
  }
  p.distinct = counts.size;
  p.mean = p.numbers ? sum / p.numbers : null;
  p.top = [...counts].sort((a, b) => b[1] - a[1]).slice(0, 3).map(([value, count]) => ({ value, count }));
  const kinds: [string, number][] = [['Number', p.numbers], ['Text', p.text], ['Logical', p.booleans], ['Error', p.errors]];
  const nonBlank = values.length - p.blanks;
  const dom = kinds.sort((a, b) => b[1] - a[1])[0];
  p.type = !nonBlank ? 'Empty' : dom[1] === nonBlank ? dom[0] : `Mixed (mostly ${dom[0].toLowerCase()})`;
  return p;
}

/** Group-by summary: one row per distinct key with the aggregate of the value column. */
export function groupSummary(keys: Scalar[], values: Scalar[], fn: 'sum' | 'count' | 'average' | 'min' | 'max'): [string, number][] {
  const groups = new Map<string, number[]>();
  const order: string[] = [];
  keys.forEach((k, i) => {
    const key = k == null || k === '' ? '(blank)' : k instanceof CellError ? k.code : String(k);
    if (!groups.has(key)) { groups.set(key, []); order.push(key); }
    const v = values[i];
    if (fn === 'count') { if (v != null && v !== '') groups.get(key)!.push(1); } else if (typeof v === 'number') groups.get(key)!.push(v);
  });
  return order.map((k) => {
    const xs = groups.get(k)!;
    const sum = xs.reduce((a, b) => a + b, 0);
    const val = fn === 'sum' || fn === 'count' ? sum : fn === 'average' ? (xs.length ? sum / xs.length : 0) : fn === 'min' ? (xs.length ? Math.min(...xs) : 0) : (xs.length ? Math.max(...xs) : 0);
    return [k, val];
  });
}

/** Histogram bins: upper bounds and counts (values <= bound, above the previous bound). */
export function histogram(xs: number[], bins?: number): { bounds: number[]; counts: number[] } {
  if (!xs.length) return { bounds: [], counts: [] };
  const min = Math.min(...xs), max = Math.max(...xs);
  const k = Math.max(1, Math.min(100, bins || Math.ceil(Math.log2(xs.length) + 1)));
  const width = (max - min) / k || 1;
  const bounds = Array.from({ length: k }, (_, i) => +(min + width * (i + 1)).toPrecision(12));
  const counts = new Array(k).fill(0);
  for (const x of xs) { let i = Math.ceil((x - min) / width) - 1; if (i < 0) i = 0; if (i >= k) i = k - 1; counts[i]++; }
  return { bounds, counts };
}

// --------------------------------------------------------------------------------------------
// Random data

export const FIRST_NAMES = ['James', 'Mary', 'John', 'Patricia', 'Robert', 'Jennifer', 'Michael', 'Linda', 'William', 'Elizabeth', 'David', 'Barbara', 'Richard', 'Susan', 'Joseph', 'Jessica', 'Thomas', 'Sarah', 'Charles', 'Karen', 'Daniel', 'Nancy', 'Matthew', 'Lisa', 'Anthony', 'Margaret', 'Mark', 'Sandra', 'Paul', 'Ashley', 'Steven', 'Emily', 'Andrew', 'Donna', 'Joshua', 'Michelle', 'Kenneth', 'Carol', 'Kevin', 'Amanda', 'Brian', 'Melissa', 'George', 'Deborah', 'Aiden', 'Sofia', 'Mateo', 'Priya', 'Wei', 'Fatima', 'Hiroshi', 'Olga', 'Kwame', 'Ingrid', 'Diego', 'Amara'];
export const LAST_NAMES = ['Smith', 'Johnson', 'Williams', 'Brown', 'Jones', 'Garcia', 'Miller', 'Davis', 'Rodriguez', 'Martinez', 'Hernandez', 'Lopez', 'Gonzalez', 'Wilson', 'Anderson', 'Thomas', 'Taylor', 'Moore', 'Jackson', 'Martin', 'Lee', 'Perez', 'Thompson', 'White', 'Harris', 'Sanchez', 'Clark', 'Ramirez', 'Lewis', 'Robinson', 'Walker', 'Young', 'Allen', 'King', 'Wright', 'Scott', 'Torres', 'Nguyen', 'Hill', 'Flores', 'Patel', 'Kim', 'Chen', 'Okafor', 'Novak', 'Schmidt', 'Rossi', 'Dubois', 'Silva', 'Kowalski'];
export const CITIES = ['New York', 'Los Angeles', 'Chicago', 'Houston', 'Phoenix', 'Philadelphia', 'San Antonio', 'San Diego', 'Dallas', 'Austin', 'Toronto', 'Vancouver', 'London', 'Manchester', 'Dublin', 'Sydney', 'Melbourne', 'Auckland', 'Paris', 'Berlin', 'Madrid', 'Rome', 'Amsterdam', 'Stockholm', 'Tokyo', 'Seoul', 'Mumbai', 'Singapore', 'Cape Town', 'Mexico City', 'São Paulo', 'Buenos Aires'];
export const COUNTRIES = ['United States', 'Canada', 'Mexico', 'Brazil', 'Argentina', 'United Kingdom', 'Ireland', 'France', 'Germany', 'Spain', 'Italy', 'Netherlands', 'Sweden', 'Norway', 'Poland', 'India', 'China', 'Japan', 'South Korea', 'Australia', 'New Zealand', 'South Africa', 'Nigeria', 'Kenya', 'Egypt'];
export const COMPANIES = ['Acme Corp', 'Globex', 'Initech', 'Umbrella Supply', 'Stark Tools', 'Wayne Logistics', 'Northwind Traders', 'Contoso Foods', 'Blue Harbor', 'Summit Partners', 'Pioneer Labs', 'Evergreen Farms', 'Redwood Media', 'Silverline Bank', 'Brightside Health'];
export const PRODUCTS = ['Widget', 'Gadget', 'Sprocket', 'Gizmo', 'Bracket', 'Valve', 'Sensor', 'Cable', 'Adapter', 'Panel', 'Filter', 'Bearing', 'Switch', 'Module', 'Fitting'];

export function guid(rnd: () => number = Math.random): string {
  const hex = (n: number) => Array.from({ length: n }, () => Math.floor(rnd() * 16).toString(16)).join('');
  return `${hex(8)}-${hex(4)}-4${hex(3)}-${'89ab'[Math.floor(rnd() * 4)]}${hex(3)}-${hex(12)}`;
}

/** Normal deviate (Box-Muller). */
export function gaussian(rnd: () => number = Math.random): number {
  let u = 0;
  while (u === 0) u = rnd();
  return Math.sqrt(-2 * Math.log(u)) * Math.cos(2 * Math.PI * rnd());
}

// --------------------------------------------------------------------------------------------
// Date calculator

export function businessDays(a: number, b: number): number {
  // Inclusive count of Monday-Friday between two serials (NETWORKDAYS semantics).
  const sign = b >= a ? 1 : -1;
  let lo = Math.min(a, b), hi = Math.max(a, b);
  lo = Math.floor(lo); hi = Math.floor(hi);
  let n = 0;
  const dow = (s: number) => ((s % 7) + 6) % 7; // serial 1 (1900-01-01) is a Sunday -> 0 with 1900 bug ignored for modern dates
  for (let s = lo; s <= hi; s++) { const d = dow(s); if (d !== 0 && d !== 6) n++; }
  return n * sign;
}

/** Difference between two dates in whole years, months and days (DATEDIF Y, YM, MD). */
export function ymdDiff(a: number, b: number): { y: number; m: number; d: number } {
  if (b < a) [a, b] = [b, a];
  const p = serialToParts(a), q = serialToParts(b);
  let y = q.y - p.y, m = q.m - p.m, d = q.d - p.d;
  if (d < 0) { m--; const pm = q.m === 1 ? 12 : q.m - 1; const py = q.m === 1 ? q.y - 1 : q.y; d += daysInMonth(py, pm); }
  if (m < 0) { y--; m += 12; }
  return { y, m, d };
}

export function addToDate(serial: number, n: number, unit: 'days' | 'weeks' | 'months' | 'years' | 'workdays'): number {
  if (unit === 'days') return serial + n;
  if (unit === 'weeks') return serial + 7 * n;
  if (unit === 'workdays') {
    let s = Math.floor(serial), left = Math.abs(n);
    const step = n >= 0 ? 1 : -1;
    const dow = (x: number) => ((x % 7) + 6) % 7;
    while (left > 0) { s += step; const d = dow(s); if (d !== 0 && d !== 6) left--; }
    return s;
  }
  const p = serialToParts(serial);
  const months = unit === 'months' ? n : 12 * n;
  const total = p.y * 12 + (p.m - 1) + months;
  const y = Math.floor(total / 12), m = (total % 12) + 1;
  return partsToSerial(y, m, Math.min(p.d, daysInMonth(y, m)));
}

/** Calendar grid for a month: weeks of day numbers (0 = blank). weekStart 0 = Sunday, 1 = Monday. */
export function monthGrid(y: number, m: number, weekStart = 0): number[][] {
  const first = new Date(Date.UTC(y, m - 1, 1)).getUTCDay();
  const offset = (first - weekStart + 7) % 7;
  const days = daysInMonth(y, m);
  const weeks: number[][] = [];
  let week: number[] = new Array(offset).fill(0);
  for (let d = 1; d <= days; d++) { week.push(d); if (week.length === 7) { weeks.push(week); week = []; } }
  if (week.length) { while (week.length < 7) week.push(0); weeks.push(week); }
  return weeks;
}

/** Quick Analysis "Totals": formulas for a row below or a column to the right of a range. */
export function totalsFormulas(rg: Range, kind: 'sum' | 'average' | 'count' | 'percent' | 'running', dir: 'rows' | 'cols'): { r: number; c: number; f: string; numFmt?: string }[] {
  const out: { r: number; c: number; f: string; numFmt?: string }[] = [];
  const fn = kind === 'average' ? 'AVERAGE' : kind === 'count' ? 'COUNT' : 'SUM';
  const all = `$${colName(rg.c1)}$${rg.r1 + 1}:$${colName(rg.c2)}$${rg.r2 + 1}`;
  if (dir === 'rows') {
    const r = rg.r2 + 1;
    for (let c = rg.c1; c <= rg.c2; c++) {
      const col = `${colName(c)}${rg.r1 + 1}:${colName(c)}${rg.r2 + 1}`;
      if (kind === 'percent') out.push({ r, c, f: `SUM(${col})/SUM(${all})`, numFmt: '0.0%' });
      else if (kind === 'running') out.push({ r, c, f: `SUM($${colName(rg.c1)}$${rg.r1 + 1}:${colName(c)}$${rg.r2 + 1})` });
      else out.push({ r, c, f: `${fn}(${col})` });
    }
  } else {
    const c = rg.c2 + 1;
    for (let r = rg.r1; r <= rg.r2; r++) {
      const row = `${colName(rg.c1)}${r + 1}:${colName(rg.c2)}${r + 1}`;
      if (kind === 'percent') out.push({ r, c, f: `SUM(${row})/SUM(${all})`, numFmt: '0.0%' });
      else if (kind === 'running') out.push({ r, c, f: `SUM($${colName(rg.c1)}$${rg.r1 + 1}:$${colName(rg.c2)}${r + 1})` });
      else out.push({ r, c, f: `${fn}(${row})` });
    }
  }
  return out;
}

// ============================================================================================
// Sheet helpers shared by the Tools, AI and Developer commands

export type OutCell = Scalar | { v?: Scalar; f?: string; style?: Partial<CellStyle> };

/** Writes a value, formula or styled value into a cell, keeping the cell's existing style unless a style patch is given. */
export function writeOut(sheet: Sheet, r: number, c: number, o: OutCell): void {
  const cur = sheet.get(r, c);
  if (o == null || typeof o !== 'object' || o instanceof CellError) {
    if (o == null || o === '') { if (cur?.s) sheet.set(r, c, { s: cur.s }); else sheet.set(r, c, undefined); return; }
    sheet.set(r, c, { ...(cur?.s ? { s: cur.s } : {}), v: o });
    return;
  }
  const cell: Cell = {};
  let s = cur?.s ?? 0;
  if (o.style) s = sheet.book.derivedStyle(s, o.style);
  if (s) cell.s = s;
  if (o.f != null) cell.f = o.f.replace(/^=/, '');
  else if (o.v != null && o.v !== '') cell.v = o.v;
  sheet.set(r, c, cell.f == null && cell.v == null && !cell.s ? undefined : cell);
}

export const BOLD: Partial<CellStyle> = { bold: true };
export const HEAD: Partial<CellStyle> = { bold: true, fill: '#dde7f0', bb: { style: 'thin', color: '#7a8a99' } };

/** Adds a sheet after the active one, fills it with rows and activates it, as one undoable step. */
export function addSheetWithRows(app: App, base: string, rows: OutCell[][], opts: { colW?: number[]; freeze?: number; at?: { r: number; c: number }; after?: (s: Sheet) => void } = {}): Sheet {
  const book = app.book;
  let made!: Sheet;
  book.transact(base, () => {
    book.recordBook();
    made = new Sheet(book, uniqueSheet(book, base));
    book.sheets.splice(book.activeSheet + 1, 0, made);
    const r0 = opts.at?.r ?? 0, c0 = opts.at?.c ?? 0;
    rows.forEach((row, i) => row.forEach((v, j) => { if (v != null && v !== '') writeOut(made, r0 + i, c0 + j, v); }));
    opts.colW?.forEach((w, i) => { if (w) made.colW.set(c0 + i, w); });
    if (opts.freeze) made.freeze = { rows: opts.freeze, cols: 0 };
    opts.after?.(made);
    book.activeSheet = book.sheets.indexOf(made);
  });
  app.exec('sheet.activate', book.sheets.indexOf(made));
  return made;
}

export function uniqueSheet(book: Workbook, base: string): string {
  const clean = base.replace(/[\\/?*[\]:]/g, ' ').slice(0, 28).trim() || 'Sheet';
  if (!book.sheetByName(clean)) return clean;
  let i = 2;
  while (book.sheetByName(`${clean} (${i})`)) i++;
  return `${clean} (${i})`;
}

/** True when a cell may be changed (sheet protection, locked cells). */
export function canChange(sheet: Sheet, r: number, c: number): boolean {
  return !sheet.protection.on || sheet.styleOf(r, c).locked === false;
}

/** Refuses with a message on a protected sheet when any target cell is locked. Returns true when ok. */
export function checkWritable(app: App, ranges: Range[], sheet: Sheet = app.sheet): boolean {
  if (!sheet.protection.on) return true;
  for (const rg of ranges) {
    const c = clipToUsed(sheet, rg, 1);
    for (let r = c.r1; r <= Math.min(c.r2, c.r1 + 5000); r++) for (let cc = c.c1; cc <= c.c2; cc++) if (!canChange(sheet, r, cc)) { app.exec('review.protectionNotice'); toast('The cell or chart you are trying to change is on a protected sheet.'); return false; }
  }
  return true;
}

function structureCheck(app: App): boolean {
  if (app.book.structureLocked) { toast('The workbook structure is protected; new sheets cannot be added.'); return false; }
  return true;
}

/** The ranges a tool works on: the selection clipped to the used range. */
export function workRanges(app: App): Range[] {
  return selRanges(app).map((rg) => clipToUsed(app.sheet, rg));
}

/** The selection, or the current region around a single selected cell. */
export function workRegion(app: App): Range {
  return dataRegion(app);
}

/** Applies a value transform to every constant in the selection. Returns the number of changed cells. */
export function transformValues(app: App, label: string, fn: (v: Scalar, r: number, c: number) => Scalar | undefined, opts: { numbersToo?: boolean; numFmt?: (v: Scalar) => string | undefined } = {}): number {
  const sheet = app.sheet;
  if (!checkWritable(app, selRanges(app))) return 0;
  let n = 0, locked = 0;
  const seen = new Set<number>();
  app.book.transact(label, () => {
    for (const rg of workRanges(app)) {
      for (let r = rg.r1; r <= rg.r2; r++) {
        if (sheet.hiddenRows.has(r) && sheet.filter) continue;
        for (let c = rg.c1; c <= rg.c2; c++) {
          const k = cellKey(r, c);
          if (seen.has(k)) continue;
          seen.add(k);
          const cell = sheet.get(r, c);
          if (!cell || cell.f != null || cell.v == null || cell.sp != null) continue;
          if (typeof cell.v !== 'string' && !opts.numbersToo) continue;
          if (!canChange(sheet, r, c)) { locked++; continue; }
          const out = fn(cell.v, r, c);
          if (out === undefined || out === cell.v) continue;
          const next: Cell = { ...cell, v: out };
          delete next.rich;
          if (out == null || out === '') delete next.v;
          const fmt = opts.numFmt?.(out);
          if (fmt) next.s = sheet.book.derivedStyle(cell.s ?? 0, { numFmt: fmt });
          sheet.set(r, c, next.v == null && !next.s ? undefined : next);
          n++;
        }
      }
    }
  });
  toast(n ? `${label}: ${n.toLocaleString()} cell${n === 1 ? '' : 's'} changed${locked ? ` (${locked} locked cells skipped)` : ''}` : `${label}: nothing to change in the selection`);
  return n;
}

const colLabel = (sheet: Sheet, rg: Range, c: number, header: boolean) => {
  const t = header ? cellText(sheet, rg.r1, c) : '';
  return t ? `${colName(c)}: ${t}` : `Column ${colName(c)}`;
};

function columnOptions(sheet: Sheet, rg: Range, header: boolean): { value: string; label: string }[] {
  const out: { value: string; label: string }[] = [];
  for (let c = rg.c1; c <= rg.c2; c++) out.push({ value: String(c), label: colLabel(sheet, rg, c, header) });
  return out;
}

function numbersIn(sheet: Sheet, c: number, r1: number, r2: number): number[] {
  const xs: number[] = [];
  for (let r = r1; r <= r2; r++) { if (sheet.hiddenRows.has(r) && sheet.filter) continue; const v = sheet.value(r, c); if (typeof v === 'number') xs.push(v); }
  return xs;
}

/** Destination column for converters: the column right of the range must be empty, or the user agrees to overwrite. */
async function confirmOverwrite(sheet: Sheet, rg: Range): Promise<boolean> {
  for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) { const x = sheet.get(r, c); if (x && (x.v != null || x.f != null)) return confirmDialog('Replace Contents?', `There's already data in ${rangeName(rg)}. Do you want to replace it?`, 'Replace'); }
  return true;
}

function isEmptySheet(app: App): boolean {
  if (!app.sheet.usedRange(false)) { toast('The sheet is empty. Enter or paste some data first.'); return true; }
  return false;
}

/** Output of converters: in place, or into the column(s) to the right of each selected column. */
async function writeConverted(app: App, label: string, conv: (v: Scalar) => Scalar | undefined, dest: 'replace' | 'right', numFmt?: string): Promise<void> {
  if (dest === 'replace') { transformValues(app, label, conv, { numbersToo: true, numFmt: numFmt ? () => numFmt : undefined }); return; }
  const sheet = app.sheet;
  const rg = clipToUsed(sheet, activeRange(app));
  const target = { r1: rg.r1, c1: rg.c2 + 1, r2: rg.r2, c2: rg.c2 + 1 + (rg.c2 - rg.c1) };
  if (!checkWritable(app, [target])) return;
  if (!(await confirmOverwrite(sheet, target))) return;
  let n = 0;
  app.book.transact(label, () => {
    for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) {
      const v = sheet.value(r, c);
      if (v == null || v === '') continue;
      const out = conv(v);
      if (out === undefined) continue;
      writeOut(sheet, r, target.c1 + (c - rg.c1), numFmt ? { v: out, style: { numFmt } } : out);
      n++;
    }
  });
  toast(`${label}: ${n} value${n === 1 ? '' : 's'} written to ${rangeName(target)}`);
}

// ============================================================================================
// Commands

const T = 'Tools';

function caseMenu(): MenuEntry[] {
  return [
    { id: 'tools.changeCase', arg: 'upper', label: 'UPPERCASE' },
    { id: 'tools.changeCase', arg: 'lower', label: 'lowercase' },
    { id: 'tools.changeCase', arg: 'proper', label: 'Proper Case' },
    { id: 'tools.changeCase', arg: 'sentence', label: 'Sentence case.' },
    { id: 'tools.changeCase', arg: 'toggle', label: 'tOGGLE cASE' },
  ];
}

def([
  // ---------------------------------------------------------------- Clean Up
  {
    id: 'tools.trimSpaces', label: 'Trim Spaces', icon: 'format_letter_spacing', tab: T, group: 'Clean Up', edits: true, popular: true,
    tip: 'Remove leading, trailing and repeated spaces, including non-breaking spaces from web pages (like TRIM, in place)',
    keywords: 'trim whitespace nbsp clean spaces', origin: 'Data clean-up add-ins (TRIM in place)',
    run: (app) => { transformValues(app, 'Trim Spaces', (v) => trimText(String(v))); },
  },
  {
    id: 'tools.removeDupSpaces', label: 'Remove Repeated Spaces', icon: 'space_bar', tab: T, group: 'Clean Up', edits: true,
    tip: 'Collapse runs of spaces inside text to one space, leaving leading and trailing spaces alone',
    run: (app) => { transformValues(app, 'Remove Repeated Spaces', (v) => collapseSpaces(String(v))); },
  },
  {
    id: 'tools.cleanText', label: 'Remove Non-Printing Characters', icon: 'cleaning_services', tab: T, group: 'Clean Up', edits: true,
    tip: 'Remove line breaks, tabs, control and zero-width characters (like CLEAN, in place)', keywords: 'clean control characters line breaks zero width bom',
    run: async (app, arg?: string) => {
      let keep = arg === 'keepBreaks';
      if (arg == null) {
        const r = await openDialog({ title: 'Remove Non-Printing Characters', fields: [{ key: 'keep', label: 'Keep line breaks inside cells', type: 'checkbox', value: false }] });
        if (!r) return;
        keep = !!r.keep;
      }
      transformValues(app, 'Remove Non-Printing Characters', (v) => cleanText(String(v), keep));
    },
  },
  {
    id: 'tools.changeCase', label: 'Change Case', icon: 'match_case', tab: T, group: 'Clean Up', edits: true, kind: 'menu', popular: true,
    tip: 'UPPERCASE, lowercase, Proper Case, Sentence case or tOGGLE cASE, in place', keywords: 'upper lower proper capitalize title case',
    origin: 'Word processors (Change Case)',
    run: async (app, arg?: CaseMode) => {
      if (!arg) {
        const i = await choiceDialog('Change Case', 'Change the text in the selected cells to:', ['UPPERCASE', 'lowercase', 'Proper Case', 'Sentence case.', 'tOGGLE cASE']);
        if (i < 0) return;
        arg = (['upper', 'lower', 'proper', 'sentence', 'toggle'] as CaseMode[])[i];
      }
      const mode = arg;
      transformValues(app, 'Change Case', (v) => changeCase(String(v), mode));
    },
    menu: caseMenu,
  },
  {
    id: 'tools.textToNumbers', label: 'Convert Text to Numbers', icon: 'pin', tab: T, group: 'Clean Up', edits: true, popular: true,
    tip: 'Turn numbers stored as text (imported "1,234", "(45)", "12-", "$3.50", " 42 ") into real numbers', keywords: 'number stored as text convert value',
    run: (app) => {
      transformValues(app, 'Convert Text to Numbers', (v) => { if (typeof v !== 'string') return undefined; const n = textToNumber(v); return n == null ? undefined : n; });
    },
  },
  {
    id: 'tools.fixDates', label: 'Fix Dates', icon: 'event_repeat', tab: T, group: 'Clean Up', edits: true, popular: true,
    tip: 'Convert dates stored as text in many formats (ISO, 3/5/24, 05.03.2024, "Mar 5, 2024", 20240305) into real dates',
    keywords: 'text dates convert datevalue two digit year dmy mdy',
    run: async (app, arg?: DateFixOptions & { format?: string }) => {
      if (isEmptySheet(app)) return;
      let o = arg;
      if (!o) {
        const sample: string[] = [];
        for (const rg of workRanges(app)) for (let r = rg.r1; r <= rg.r2 && sample.length < 4; r++) for (let c = rg.c1; c <= rg.c2 && sample.length < 4; c++) { const v = app.sheet.value(r, c); if (typeof v === 'string' && v.trim()) sample.push(v); }
        const preview = (v: Record<string, any>) => sample.map((s) => {
          const n = parseTextDate(s, { order: v.order, pivot: Number(v.pivot) });
          const p = n == null ? null : serialToParts(n);
          return `<div>${escHtml(s)} → ${p ? `${p.y}-${String(p.m).padStart(2, '0')}-${String(p.d).padStart(2, '0')}` : '<i>not a date</i>'}</div>`;
        }).join('') || '<i>No text in the selection</i>';
        const r = await openDialog({
          title: 'Fix Dates', preview: true, fields: [
            { key: 'order', label: 'Numeric dates are written as', type: 'select', value: app.prefs.dateOrder === 'dmy' ? 'dmy' : 'auto', options: [{ value: 'auto', label: 'Detect (month first when unclear)' }, { value: 'mdy', label: 'Month/Day/Year (3/14/2024)' }, { value: 'dmy', label: 'Day/Month/Year (14/3/2024)' }, { value: 'ymd', label: 'Year/Month/Day (24/3/14)' }] },
            { key: 'pivot', label: 'Two-digit years up to this are 20xx', type: 'number', value: 29, min: 0, max: 99, half: true },
            { key: 'format', label: 'Display format', type: 'combo', value: 'yyyy-mm-dd', options: ['yyyy-mm-dd', 'm/d/yyyy', 'd/m/yyyy', 'd mmm yyyy', 'mmmm d, yyyy', 'dd.mm.yyyy'], half: true },
          ],
          onChange: (v, ctl) => ctl.setPreview(preview(v)),
        });
        if (!r) return;
        o = { order: r.order, pivot: Number(r.pivot ?? 29), format: r.format || 'yyyy-mm-dd' };
      }
      const opts = o;
      transformValues(app, 'Fix Dates', (v) => { if (typeof v !== 'string') return undefined; const n = parseTextDate(v, { ...opts, date1904: app.book.calc.date1904 }); return n == null ? undefined : n; }, { numFmt: () => opts.format || 'yyyy-mm-dd' });
    },
  },
  {
    id: 'tools.ocrCleanup', label: 'Clean Up OCR Text', icon: 'document_scanner', tab: T, group: 'Clean Up', edits: true,
    tip: 'Repair text from scans and PDFs: join broken lines, rejoin hyphenated words, fix O/0 and l/1 in numbers, ligatures and curly quotes',
    keywords: 'ocr scan pdf hyphen ligature quotes genealogy transcription', origin: 'Transcription and genealogy tools',
    run: async (app, arg?: OcrOptions) => {
      let o = arg;
      if (!o) {
        const r = await openDialog({
          title: 'Clean Up OCR Text', fields: [
            { key: 'joinLines', label: 'Join lines broken inside a paragraph', type: 'checkbox', value: true },
            { key: 'dehyphenate', label: 'Rejoin words hyphenated at line ends', type: 'checkbox', value: true },
            { key: 'digits', label: 'Letters inside numbers become digits (1O5 → 105, l23 → 123)', type: 'checkbox', value: true },
            { key: 'letters', label: 'Digits inside words become letters (C0MPANY → COMPANY)', type: 'checkbox', value: true },
            { key: 'ligatures', label: 'Expand ligatures (ﬁ → fi)', type: 'checkbox', value: true },
            { key: 'quotes', label: 'Straighten curly quotes', type: 'checkbox', value: true },
          ],
        });
        if (!r) return;
        o = r as OcrOptions;
      }
      const opts = o;
      transformValues(app, 'Clean Up OCR Text', (v) => ocrClean(String(v), opts));
    },
  },
  {
    id: 'tools.removeAccents', label: 'Remove Accents', icon: 'text_format', tab: T, group: 'Clean Up', edits: true,
    tip: 'Replace accented letters with plain ones (é → e, ß → ss) for matching and IDs', keywords: 'diacritics ascii fold',
    run: (app) => { transformValues(app, 'Remove Accents', (v) => removeAccents(String(v))); },
  },
  {
    id: 'tools.removeChars', label: 'Remove Characters…', icon: 'backspace', tab: T, group: 'Clean Up', edits: true,
    tip: 'Remove digits, letters, punctuation, symbols, non-ASCII characters or a custom set from text',
    run: async (app) => {
      const r = await openDialog({
        title: 'Remove Characters', fields: [
          { key: 'what', label: 'Remove', type: 'radio', value: 'digits', options: [{ value: 'digits', label: 'Digits' }, { value: 'letters', label: 'Letters' }, { value: 'punct', label: 'Punctuation and symbols' }, { value: 'nonascii', label: 'Non-ASCII characters' }, { value: 'custom', label: 'These characters:' }] },
          { key: 'chars', label: 'Characters', type: 'text', value: '', visibleIf: (v) => v.what === 'custom' },
          { key: 'first', label: 'Only the first … characters (0 = all)', type: 'number', value: 0, min: 0, half: true },
          { key: 'last', label: 'Only the last … characters (0 = all)', type: 'number', value: 0, min: 0, half: true },
        ],
      });
      if (!r) return;
      const res: Record<string, RegExp> = { digits: /\p{N}/gu, letters: /\p{L}/gu, punct: /[\p{P}\p{S}]/gu, nonascii: /[^\x20-\x7e\n]/g };
      const re = r.what === 'custom' ? new RegExp(`[${String(r.chars || '').replace(/[\\\]^-]/g, '\\$&')}]`, 'g') : res[r.what];
      const first = Number(r.first) || 0, last = Number(r.last) || 0;
      transformValues(app, 'Remove Characters', (v) => {
        const s = String(v);
        if (first) return s.slice(first);
        if (last) return s.slice(0, Math.max(0, s.length - last));
        return r.what === 'custom' && !r.chars ? s : s.replace(re, '');
      });
    },
  },
  {
    id: 'tools.fillBlanks', label: 'Fill Blank Cells', icon: 'format_color_fill', tab: T, group: 'Clean Up', edits: true, popular: true,
    tip: 'Fill empty cells in the selection with the value above or below, a fixed value, or a formula', keywords: 'fill down blanks empty gaps',
    run: async (app, arg?: { mode: 'above' | 'below' | 'value' | 'formula'; value?: string }) => {
      if (isEmptySheet(app)) return;
      let o = arg;
      if (!o) {
        const r = await openDialog({
          title: 'Fill Blank Cells', fields: [
            { key: 'mode', label: 'Fill each blank cell with', type: 'radio', value: 'above', options: [{ value: 'above', label: 'The value above' }, { value: 'below', label: 'The value below' }, { value: 'value', label: 'This value' }, { value: 'formula', label: 'This formula (references are relative to the blank cell)' }] },
            { key: 'value', label: 'Value or formula', type: 'text', value: '', placeholder: 'e.g. 0, N/A, or =R[-1]C+1 style =A1', visibleIf: (v) => v.mode === 'value' || v.mode === 'formula' },
          ],
        });
        if (!r) return;
        o = { mode: r.mode, value: r.value };
      }
      fillBlanks(app, o.mode, o.value ?? '');
    },
  },
  {
    id: 'tools.deleteBlankRows', label: 'Remove Blank Rows', icon: 'delete_sweep', tab: T, group: 'Clean Up', edits: true,
    tip: 'Move non-blank rows up inside the selected range, removing rows that are completely empty there', keywords: 'empty rows compact',
    run: (app) => {
      if (isEmptySheet(app)) return;
      const s = app.sheet;
      const rg = workRegion(app);
      if (!checkWritable(app, [rg])) return;
      const keep: { r: number; row: (Cell | undefined)[] }[] = [];
      let removed = 0;
      for (let r = rg.r1; r <= rg.r2; r++) {
        const row: (Cell | undefined)[] = [];
        let any = false;
        for (let c = rg.c1; c <= rg.c2; c++) { const x = s.get(r, c); row.push(x ? { ...x } : undefined); if (x && (x.f != null || (x.v != null && x.v !== ''))) any = true; }
        if (any) keep.push({ r, row }); else removed++;
      }
      if (!removed) { toast('There are no blank rows in the selection.'); return; }
      app.book.transact('Remove Blank Rows', () => {
        for (let i = 0; i <= rg.r2 - rg.r1; i++) for (let j = 0; j <= rg.c2 - rg.c1; j++) {
          const k = keep[i];
          s.set(rg.r1 + i, rg.c1 + j, k ? moved(k.row[j], rg.r1 + i - k.r, 0) : undefined);
        }
      });
      toast(`${removed} blank row${removed === 1 ? '' : 's'} removed`);
    },
  },

  // ---------------------------------------------------------------- Transform
  {
    id: 'tools.splitColumn', label: 'Split Column…', icon: 'vertical_split', tab: T, group: 'Transform', edits: true, popular: true,
    tip: 'Split a column into several by delimiter, fixed positions, case changes or number/text boundaries, with a preview',
    keywords: 'text to columns split delimiter fixed width', origin: 'Text to Columns; data tools in other spreadsheets',
    run: (app, arg?: { mode: SplitMode; keep?: boolean }) => splitColumn(app, arg),
  },
  {
    id: 'tools.mergeColumns', label: 'Merge Columns…', icon: 'join', tab: T, group: 'Transform', edits: true,
    tip: 'Join the selected columns row by row with a separator (without losing data, unlike merging cells)', keywords: 'concatenate join combine textjoin',
    run: (app, arg?: { sep: string; skipBlanks?: boolean; keep?: boolean }) => mergeColumns(app, arg),
  },
  {
    id: 'tools.unpivot', label: 'Unpivot Columns…', icon: 'table_rows', tab: T, group: 'Transform', edits: true,
    tip: 'Turn a wide table (one column per month, say) into a long one with Attribute and Value columns, on a new sheet',
    keywords: 'melt wide to long normalize reshape', origin: 'Query editors (Unpivot)',
    run: (app, arg?: { idCols: number; keepBlanks?: boolean }) => unpivotCmd(app, arg),
  },
  {
    id: 'tools.transposeInPlace', label: 'Transpose in Place', icon: 'flip', tab: T, group: 'Transform', edits: true,
    tip: 'Swap rows and columns of the selected range where it is', keywords: 'rotate transpose swap rows columns',
    run: (app) => transposeInPlace(app),
  },
  {
    id: 'tools.reverse', label: 'Reverse Order', icon: 'swap_vert', tab: T, group: 'Transform', edits: true, kind: 'split',
    tip: 'Reverse the order of the selected rows (or columns)', keywords: 'flip upside down reverse rows columns',
    run: (app, arg?: 'rows' | 'cols') => reverseRange(app, arg === 'cols' ? 'cols' : 'rows'),
    menu: () => [{ id: 'tools.reverse', arg: 'rows', label: 'Reverse Rows', icon: 'swap_vert' }, { id: 'tools.reverse', arg: 'cols', label: 'Reverse Columns', icon: 'swap_horiz' }],
  },
  {
    id: 'tools.shuffle', label: 'Shuffle Rows', icon: 'shuffle', tab: T, group: 'Transform', edits: true,
    tip: 'Put the selected rows in random order (a header row is kept in place)', keywords: 'random order sort randomize sample',
    origin: 'Spreadsheet add-ins (Randomize range)',
    run: (app) => shuffleRows(app),
  },
  {
    id: 'tools.addText', label: 'Add Prefix or Suffix…', icon: 'add_notes', tab: T, group: 'Transform', edits: true,
    tip: 'Add text before and/or after every value in the selection', keywords: 'prefix suffix prepend append',
    run: async (app, arg?: { prefix?: string; suffix?: string; blanks?: boolean }) => {
      let o = arg;
      if (!o) {
        const r = await openDialog({ title: 'Add Prefix or Suffix', fields: [{ key: 'prefix', label: 'Before each value', type: 'text', value: '', half: true }, { key: 'suffix', label: 'After each value', type: 'text', value: '', half: true }] });
        if (!r) return;
        o = r;
      }
      const p = o.prefix || '', s = o.suffix || '';
      transformValues(app, 'Add Prefix or Suffix', (v) => `${p}${typeof v === 'number' ? cellNumberText(v) : String(v)}${s}`, { numbersToo: true });
    },
  },
  {
    id: 'tools.roundValues', label: 'Round Values…', icon: 'decimal_decrease', tab: T, group: 'Transform', edits: true,
    tip: 'Round the stored numbers (not just the display) to a number of decimals or significant figures; formulas can be wrapped in ROUND',
    keywords: 'round precision decimals significant',
    run: async (app, arg?: { digits: number; mode: 'round' | 'up' | 'down' | 'sig'; formulas?: boolean }) => {
      let o = arg;
      if (!o) {
        const r = await openDialog({ title: 'Round Values', fields: [
          { key: 'mode', label: 'Rounding', type: 'select', value: 'round', options: [{ value: 'round', label: 'Round (half away from zero)' }, { value: 'up', label: 'Round up' }, { value: 'down', label: 'Round down' }, { value: 'sig', label: 'Significant figures' }] },
          { key: 'digits', label: 'Decimal places (or figures)', type: 'number', value: 2, min: -10, max: 15 },
          { key: 'formulas', label: 'Wrap formulas in ROUND()', type: 'checkbox', value: false },
        ] });
        if (!r) return;
        o = { digits: Number(r.digits) || 0, mode: r.mode, formulas: !!r.formulas };
      }
      roundValues(app, o.digits, o.mode, !!o.formulas);
    },
  },
  {
    id: 'tools.formulasToValues', label: 'Convert Formulas to Values', icon: 'functions', tab: T, group: 'Transform', edits: true,
    tip: 'Replace the formulas in the selection with their current results', keywords: 'paste values freeze formulas',
    run: (app) => {
      const s = app.sheet;
      if (!checkWritable(app, selRanges(app))) return;
      let n = 0;
      app.book.transact('Convert Formulas to Values', () => {
        for (const rg of workRanges(app)) s.forEachIn(rg, (r, c, cell) => {
          if (cell.f == null && cell.sp == null) return;
          const v = cell.v ?? null;
          const next: Cell = { ...(cell.s ? { s: cell.s } : {}), ...(v != null ? { v } : {}) };
          s.set(r, c, next.v == null && !next.s ? undefined : next);
          n++;
        });
      });
      toast(n ? `${n} formula${n === 1 ? '' : 's'} converted to values` : 'There are no formulas in the selection.');
    },
  },
  {
    id: 'tools.unmergeFill', label: 'Unmerge and Fill', icon: 'call_split', tab: T, group: 'Transform', edits: true,
    tip: 'Unmerge merged cells in the selection and copy the value into every cell that was part of the merge (good before sorting or filtering)',
    keywords: 'unmerge fill merged cells',
    run: (app) => {
      const s = app.sheet;
      const sel = selRanges(app);
      const merges = s.merges.filter((m) => sel.some((rg) => m.r1 <= rg.r2 && m.r2 >= rg.r1 && m.c1 <= rg.c2 && m.c2 >= rg.c1));
      if (!merges.length) { toast('There are no merged cells in the selection.'); return; }
      if (!checkWritable(app, merges)) return;
      app.book.transact('Unmerge and Fill', () => {
        app.book.recordSheet(s);
        for (const m of merges) {
          const top = s.get(m.r1, m.c1);
          for (let r = m.r1; r <= m.r2; r++) for (let c = m.c1; c <= m.c2; c++) {
            if (r === m.r1 && c === m.c1) continue;
            if (top?.f != null) s.set(r, c, { ...top, f: shiftFormula(top.f, r - m.r1, c - m.c1) });
            else s.set(r, c, top ? { ...top } : undefined);
          }
        }
        s.merges = s.merges.filter((m) => !merges.includes(m));
      });
      toast(`${merges.length} merged range${merges.length === 1 ? '' : 's'} unmerged and filled`);
    },
  },
  {
    id: 'tools.numberToWords', label: 'Number to Words', icon: 'abc', tab: T, group: 'Transform', edits: true, kind: 'menu',
    tip: 'Spell out numbers: plain words, check style (… and 45/100), legal style, ordinals or years (eighteen fifty)',
    keywords: 'spell number check amount words genealogy deed year ordinal', origin: 'Check-writing and genealogy document tools',
    run: async (app, arg?: WordsStyle | { style: WordsStyle; dest?: 'replace' | 'right'; currency?: string }) => {
      let o = typeof arg === 'object' ? arg : undefined;
      if (!o) {
        const r = await openDialog({ title: 'Number to Words', fields: [
          { key: 'style', label: 'Style', type: 'select', value: typeof arg === 'string' ? arg : 'words', options: [{ value: 'words', label: 'Words (one hundred twenty-three point four)' }, { value: 'check', label: 'Check (One Hundred Twenty-Three and 40/100)' }, { value: 'legal', label: 'Legal (One Hundred and Twenty-Three)' }, { value: 'ordinal', label: 'Ordinal (one hundred twenty-third)' }, { value: 'year', label: 'Year (eighteen fifty)' }] },
          { key: 'currency', label: 'Currency word after check amounts', type: 'text', value: '', placeholder: 'e.g. Dollars', visibleIf: (v) => v.style === 'check' || v.style === 'legal' },
          { key: 'dest', label: 'Put the words', type: 'radio', value: 'right', options: [{ value: 'right', label: 'In the column to the right' }, { value: 'replace', label: 'In place of the numbers' }] },
        ] });
        if (!r) return;
        o = { style: r.style, dest: r.dest, currency: r.currency };
      }
      const oo = o;
      await writeConverted(app, 'Number to Words', (v) => { const n = typeof v === 'number' ? v : typeof v === 'string' ? textToNumber(v) : null; return n == null ? undefined : numberToWords(n, oo.style, oo.currency); }, oo.dest || 'right');
    },
    menu: () => (['words', 'check', 'legal', 'ordinal', 'year'] as WordsStyle[]).map((s) => ({ id: 'tools.numberToWords', arg: s, label: { words: 'Words…', check: 'Check Amount…', legal: 'Legal Amount…', ordinal: 'Ordinal…', year: 'Year…' }[s] })),
  },
  {
    id: 'tools.romanNumerals', label: 'Roman Numerals', icon: 'looks_one', tab: T, group: 'Transform', edits: true, kind: 'menu',
    tip: 'Convert numbers to Roman numerals or Roman numerals back to numbers', keywords: 'roman numeral convert arabic',
    run: async (app, arg?: 'to' | 'from') => {
      if (!arg) { const i = await choiceDialog('Roman Numerals', 'Convert the selection:', ['Numbers to Roman numerals', 'Roman numerals to numbers']); if (i < 0) return; arg = i ? 'from' : 'to'; }
      if (arg === 'to') transformValues(app, 'Roman Numerals', (v) => (typeof v === 'number' && v >= 1 && v <= 3999 ? toRoman(v) : undefined), { numbersToo: true });
      else transformValues(app, 'Roman Numerals', (v) => (typeof v === 'string' ? fromRoman(v) ?? undefined : undefined));
    },
    menu: () => [{ id: 'tools.romanNumerals', arg: 'to', label: 'Numbers to Roman Numerals' }, { id: 'tools.romanNumerals', arg: 'from', label: 'Roman Numerals to Numbers' }],
  },
  {
    id: 'tools.extract', label: 'Extract with Pattern', icon: 'manage_search', tab: T, group: 'Transform', edits: true, kind: 'menu',
    tip: 'Copy the email addresses, web addresses, numbers, phone numbers, dates or a regular expression found in each cell into the next column',
    keywords: 'regex regular expression extract email url phone',
    run: async (app, arg?: string | { pattern: string }) => {
      let re: RegExp | null = null;
      if (typeof arg === 'string' && EXTRACT_PATTERNS[arg]) re = EXTRACT_PATTERNS[arg].re;
      else {
        const r = await openDialog({ title: 'Extract with Pattern', fields: [
          { key: 'kind', label: 'Find', type: 'select', value: 'emails', options: [...Object.entries(EXTRACT_PATTERNS).map(([k, p]) => ({ value: k, label: p.label })), { value: 'regex', label: 'Regular expression…' }] },
          { key: 'pattern', label: 'Regular expression', type: 'text', value: typeof arg === 'object' ? arg.pattern : '', placeholder: 'e.g. [A-Z]{2}\\d{4}', visibleIf: (v) => v.kind === 'regex' },
        ] });
        if (!r) return;
        try { re = r.kind === 'regex' ? new RegExp(r.pattern, 'gu') : EXTRACT_PATTERNS[r.kind].re; } catch (e) { toast(`That pattern is not valid: ${(e as Error).message}`); return; }
      }
      const pat = re!;
      await writeConverted(app, 'Extract with Pattern', (v) => { const m = extractMatches(String(v), pat); return m.length ? m.join(', ') : undefined; }, 'right');
    },
    menu: () => [...Object.entries(EXTRACT_PATTERNS).map(([k, p]) => ({ id: 'tools.extract', arg: k, label: p.label })), { separator: true }, { id: 'tools.extract', label: 'Regular Expression…' }],
  },

  // ---------------------------------------------------------------- Compare
  { id: 'tools.compareSheets', label: 'Compare Sheets…', icon: 'compare', tab: T, group: 'Compare', edits: true, tip: 'Compare two sheets cell by cell; a report sheet shows every difference highlighted', keywords: 'diff differences compare worksheets', origin: 'Spreadsheet comparison tools', run: (app, arg?: { a: string; b: string; formulas?: boolean }) => compareSheets(app, arg) },
  { id: 'tools.findDuplicates', label: 'Find Duplicates…', icon: 'content_copy', tab: T, group: 'Compare', edits: true, popular: true, tip: 'Highlight or list duplicate rows, judged by the key columns you choose', keywords: 'duplicate dedupe repeated rows', run: (app, arg?: any) => findDuplicates(app, arg) },
  { id: 'tools.fuzzyMatch', label: 'Fuzzy Match…', icon: 'join_inner', tab: T, group: 'Compare', edits: true, tip: 'For each value, find the closest value in another list (spelling variants, typos) with a similarity score', keywords: 'fuzzy lookup approximate match levenshtein jaro winkler similar', origin: 'Fuzzy lookup add-ins', run: (app, arg?: any) => fuzzyMatchCmd(app, arg) },

  // ---------------------------------------------------------------- Analyze
  { id: 'tools.dataProfile', label: 'Data Profile', icon: 'query_stats', tab: T, group: 'Analyze', edits: true, popular: true, tip: 'A new sheet describing each column: types, blanks, distinct values, minimum, maximum, mean and most common values', keywords: 'profile column quality summary', origin: 'Data-preparation tools (column profile)', run: (app) => dataProfile(app) },
  { id: 'tools.descriptiveStats', label: 'Descriptive Statistics', icon: 'functions', tab: T, group: 'Analyze', edits: true, tip: 'Mean, standard error, median, mode, standard deviation, variance, kurtosis, skewness, range, min, max, sum, count and 95% confidence for each numeric column', keywords: 'statistics mean median stdev analysis toolpak', run: (app) => descriptiveStats(app) },
  { id: 'tools.histogram', label: 'Histogram', icon: 'bar_chart', tab: T, group: 'Analyze', edits: true, tip: 'Frequency table and chart of a numeric column on a new sheet', keywords: 'frequency distribution bins', run: (app, arg?: { bins?: number }) => histogramCmd(app, arg) },
  { id: 'tools.correlation', label: 'Correlation Matrix', icon: 'grid_on', tab: T, group: 'Analyze', edits: true, tip: 'Pearson correlation between every pair of numeric columns, color-coded, on a new sheet', keywords: 'correlation pearson matrix relationship', run: (app) => correlationCmd(app) },
  { id: 'tools.groupSummary', label: 'Summarize by Category…', icon: 'category', tab: T, group: 'Analyze', edits: true, tip: 'Group rows by a column and total, count or average another column (a quick summary without a PivotTable)', keywords: 'group by categories subtotal summary', origin: 'Categories in tabular spreadsheet apps', run: (app, arg?: any) => groupSummaryCmd(app, arg) },

  // ---------------------------------------------------------------- Generate
  { id: 'tools.randomData', label: 'Random Data…', icon: 'casino', tab: T, group: 'Generate', edits: true, tip: 'Fill the selection with random numbers, dates, names, cities, choices from a list, yes/no values or unique IDs', keywords: 'random generator sample test data guid names', origin: 'Fill ▸ Random Number in other spreadsheets', run: (app, arg?: any) => randomData(app, arg) },
  { id: 'tools.sequence', label: 'Sequence…', icon: 'format_list_numbered', tab: T, group: 'Generate', edits: true, tip: 'Fill numbers, dates, weekdays, months or numbered codes (INV-0001…) from the active cell', keywords: 'series fill numbers increment', run: (app, arg?: any) => sequenceCmd(app, arg) },
  { id: 'tools.calendar', label: 'Calendar…', icon: 'calendar_month', tab: T, group: 'Generate', edits: true, tip: 'Make a month or year calendar on a new sheet', keywords: 'calendar month year planner', run: (app, arg?: any) => calendarCmd(app, arg) },
  { id: 'tools.dateCalculator', label: 'Date Calculator…', icon: 'date_range', tab: T, group: 'Generate', edits: false, noRecord: true, tip: 'Days, workdays, weeks and years between two dates, or a date plus or minus days, workdays, months or years; insert the result', keywords: 'date difference age days between add days', origin: 'Desktop calculator date mode', run: (app) => dateCalculator(app) },
]);

function cellNumberText(v: number): string { return String(+v.toPrecision(15)); }

// --------------------------------------------------------------------------------------------
// Command implementations

function fillBlanks(app: App, mode: 'above' | 'below' | 'value' | 'formula', value: string): void {
  const s = app.sheet;
  const ranges = selRanges(app).map((rg) => (rg.r1 === rg.r2 && rg.c1 === rg.c2 ? workRegion(app) : clipToUsed(s, rg)));
  if (!checkWritable(app, ranges)) return;
  let n = 0;
  const isBlank = (r: number, c: number) => { const x = s.get(r, c); return !x || ((x.v == null || x.v === '') && x.f == null && x.sp == null); };
  if (mode === 'formula' && !value.trim()) { toast('Enter a formula.'); return; }
  const first = ranges[0];
  app.book.transact('Fill Blank Cells', () => {
    for (const rg of ranges) {
      for (let c = rg.c1; c <= rg.c2; c++) {
        const rows = mode === 'below' ? Array.from({ length: rg.r2 - rg.r1 + 1 }, (_, i) => rg.r2 - i) : Array.from({ length: rg.r2 - rg.r1 + 1 }, (_, i) => rg.r1 + i);
        let last: Cell | undefined;
        for (const r of rows) {
          if (sheetRowHidden(s, r)) continue;
          if (!isBlank(r, c)) { last = s.get(r, c); continue; }
          if (!canChange(s, r, c)) continue;
          const cur = s.get(r, c);
          if (mode === 'above' || mode === 'below') {
            if (!last) continue;
            const v = last.v ?? null;
            if (v == null) continue;
            const st = cur?.s ?? last.s;
            s.set(r, c, st ? { s: st, v } : { v });
          } else if (mode === 'value') {
            const n2 = textToNumber(value);
            s.set(r, c, { ...(cur?.s ? { s: cur.s } : {}), v: n2 != null ? n2 : value });
          } else {
            // The formula is written for the first blank cell's position (top-left of the selection).
            s.set(r, c, { ...(cur?.s ? { s: cur.s } : {}), f: shiftFormula(value.replace(/^=/, ''), r - first.r1, c - first.c1) });
          }
          n++;
        }
      }
    }
  });
  toast(n ? `${n} blank cell${n === 1 ? '' : 's'} filled` : 'No blank cells were filled.');
}

const sheetRowHidden = (s: Sheet, r: number) => s.hiddenRows.has(r) && !!s.filter;

async function splitColumn(app: App, arg?: { mode: SplitMode; keep?: boolean }): Promise<void> {
  if (isEmptySheet(app)) return;
  const s = app.sheet;
  const sel = clipToUsed(s, activeRange(app));
  const rg = sel.r1 === sel.r2 && sel.c1 === sel.c2 ? { ...workRegion(app), c1: sel.c1, c2: sel.c1 } : sel;
  if (rg.c1 !== rg.c2) { toast('Select one column to split.'); return; }
  const col = rg.c1;
  const texts: { r: number; t: string }[] = [];
  for (let r = rg.r1; r <= rg.r2; r++) { const x = s.get(r, col); if (x && x.f == null && x.v != null && x.v !== '') texts.push({ r, t: cellText(s, r, col) }); }
  if (!texts.length) { toast('There is no text to split in that column.'); return; }
  let o = arg;
  if (!o) {
    const modeOf = (v: Record<string, any>): SplitMode => {
      if (v.kind === 'positions') return { kind: 'positions', at: String(v.positions || '').split(/[,\s]+/).map(Number).filter((n) => n > 0) };
      if (v.kind === 'case') return { kind: 'case' };
      if (v.kind === 'numberText') return { kind: 'numberText' };
      const d = (v.comma ? ',' : '') + (v.semicolon ? ';' : '') + (v.tab ? '\t' : '') + (v.space ? ' ' : '') + (v.other || '');
      return { kind: 'delimiter', delims: d || ',', consecutive: !!v.consecutive };
    };
    const preview = (v: Record<string, any>) => {
      const m = modeOf(v);
      const rows = texts.slice(0, 6).map((x) => splitText(x.t, m));
      return `<table class="ql-stats">${rows.map((r) => `<tr>${r.map((c) => `<td>${escHtml(c)}</td>`).join('')}</tr>`).join('')}</table>`;
    };
    const r = await openDialog({
      title: 'Split Column', wide: true, preview: true, fields: [
        { key: 'kind', label: 'Split by', type: 'radio', value: 'delimiter', options: [{ value: 'delimiter', label: 'Delimiters' }, { value: 'positions', label: 'Fixed positions' }, { value: 'case', label: 'Change from lowercase to uppercase (firstName → first | Name)' }, { value: 'numberText', label: 'Change between numbers and text (12kg → 12 | kg)' }] },
        { key: 'comma', label: 'Comma', type: 'checkbox', value: true, half: true, visibleIf: (v) => v.kind === 'delimiter' },
        { key: 'semicolon', label: 'Semicolon', type: 'checkbox', value: false, half: true, visibleIf: (v) => v.kind === 'delimiter' },
        { key: 'tab', label: 'Tab', type: 'checkbox', value: false, half: true, visibleIf: (v) => v.kind === 'delimiter' },
        { key: 'space', label: 'Space', type: 'checkbox', value: false, half: true, visibleIf: (v) => v.kind === 'delimiter' },
        { key: 'other', label: 'Other', type: 'text', value: '', half: true, visibleIf: (v) => v.kind === 'delimiter' },
        { key: 'consecutive', label: 'Treat consecutive delimiters as one', type: 'checkbox', value: false, half: true, visibleIf: (v) => v.kind === 'delimiter' },
        { key: 'positions', label: 'Split after characters (e.g. 3, 7)', type: 'text', value: '', visibleIf: (v) => v.kind === 'positions' },
        { key: 'trim', label: 'Trim spaces around each part', type: 'checkbox', value: true },
        { key: 'keep', label: 'Keep the original column (parts go to the right)', type: 'checkbox', value: true },
      ],
      onChange: (v, ctl) => ctl.setPreview(preview(v)),
    });
    if (!r) return;
    o = { mode: modeOf(r), keep: !!r.keep };
    (o as any).trim = !!r.trim;
  }
  const parts = texts.map((x) => ({ r: x.r, p: splitText(x.t, o!.mode).map((y) => ((o as any).trim === false ? y : y.trim())) }));
  const width = Math.max(...parts.map((x) => x.p.length));
  const c0 = o.keep === false ? col : col + 1;
  const target = { r1: rg.r1, c1: c0, r2: rg.r2, c2: c0 + width - 1 };
  const check = o.keep === false ? { ...target, c1: col + 1 } : target;
  if (check.c2 >= check.c1 && !(await confirmOverwrite(s, check))) return;
  if (!checkWritable(app, [target])) return;
  app.book.transact('Split Column', () => {
    for (const { r, p } of parts) for (let j = 0; j < width; j++) {
      const t = p[j] ?? '';
      const n = t !== '' ? textToNumber(t) : null;
      writeOut(s, r, c0 + j, t === '' ? null : n != null && !/^0\d/.test(t) ? n : t);
    }
  });
  toast(`Split into ${width} column${width === 1 ? '' : 's'}`);
}

async function mergeColumns(app: App, arg?: { sep: string; skipBlanks?: boolean; keep?: boolean }): Promise<void> {
  if (isEmptySheet(app)) return;
  const s = app.sheet;
  const rg = clipToUsed(s, activeRange(app));
  if (rg.c1 === rg.c2) { toast('Select two or more columns to merge.'); return; }
  let o = arg;
  if (!o) {
    const r = await openDialog({ title: 'Merge Columns', fields: [
      { key: 'sep', label: 'Separator', type: 'combo', value: ' ', options: [' ', ', ', '; ', '-', '/', '|', '\\n'] },
      { key: 'skipBlanks', label: 'Skip blank cells', type: 'checkbox', value: true },
      { key: 'keep', label: 'Keep the original columns (result goes to the right)', type: 'checkbox', value: true },
    ] });
    if (!r) return;
    o = { sep: String(r.sep ?? ' ').replace(/\\n/g, '\n'), skipBlanks: !!r.skipBlanks, keep: !!r.keep };
  }
  const target = o.keep === false ? { r1: rg.r1, c1: rg.c1, r2: rg.r2, c2: rg.c1 } : { r1: rg.r1, c1: rg.c2 + 1, r2: rg.r2, c2: rg.c2 + 1 };
  if (o.keep !== false && !(await confirmOverwrite(s, target))) return;
  if (!checkWritable(app, [rg, target])) return;
  app.book.transact('Merge Columns', () => {
    for (let r = rg.r1; r <= rg.r2; r++) {
      if (sheetRowHidden(s, r)) continue;
      const parts: string[] = [];
      for (let c = rg.c1; c <= rg.c2; c++) { const t = cellText(s, r, c); if (t !== '' || !o!.skipBlanks) parts.push(t); }
      if (o!.keep === false) for (let c = rg.c1 + 1; c <= rg.c2; c++) writeOut(s, r, c, null);
      writeOut(s, r, target.c1, parts.join(o!.sep) || null);
    }
  });
  toast(`Merged ${rg.c2 - rg.c1 + 1} columns into ${colName(target.c1)}`);
}

async function unpivotCmd(app: App, arg?: { idCols: number; keepBlanks?: boolean }): Promise<void> {
  if (isEmptySheet(app) || !structureCheck(app)) return;
  const s = app.sheet;
  const rg = workRegion(app);
  if (rg.c2 - rg.c1 < 1 || rg.r2 <= rg.r1) { toast('Select a table with a header row and at least two columns.'); return; }
  let o = arg;
  if (!o) {
    const r = await openDialog({ title: 'Unpivot Columns', help: `Range ${rangeName(rg)}. The first row is used as headers.`, fields: [
      { key: 'idCols', label: 'Identifier columns to keep (from the left)', type: 'number', value: 1, min: 0, max: rg.c2 - rg.c1 },
      { key: 'attr', label: 'Name for the attribute column', type: 'text', value: 'Attribute', half: true },
      { key: 'val', label: 'Name for the value column', type: 'text', value: 'Value', half: true },
      { key: 'keepBlanks', label: 'Include blank values', type: 'checkbox', value: false },
    ] });
    if (!r) return;
    o = { idCols: Math.max(0, Math.min(rg.c2 - rg.c1, Number(r.idCols) || 0)), keepBlanks: !!r.keepBlanks };
    (o as any).names = [r.attr || 'Attribute', r.val || 'Value'];
  }
  const header: Scalar[] = [], rows: Scalar[][] = [];
  for (let c = rg.c1; c <= rg.c2; c++) header.push(s.value(rg.r1, c));
  for (let r = rg.r1 + 1; r <= rg.r2; r++) { if (sheetRowHidden(s, r)) continue; const row: Scalar[] = []; for (let c = rg.c1; c <= rg.c2; c++) row.push(s.value(r, c)); rows.push(row); }
  const out = unpivot(header, rows, o.idCols, o.keepBlanks, (o as any).names);
  addSheetWithRows(app, 'Unpivoted', out.map((row, i) => (i === 0 ? row.map((v) => ({ v, style: HEAD })) : row)), { freeze: 1 });
  toast(`${out.length - 1} rows created`);
}

function blockOf(s: Sheet, rg: Range): (Cell | undefined)[][] {
  const m: (Cell | undefined)[][] = [];
  for (let r = rg.r1; r <= rg.r2; r++) { const row: (Cell | undefined)[] = []; for (let c = rg.c1; c <= rg.c2; c++) { const x = s.get(r, c); row.push(x ? { ...x } : undefined); } m.push(row); }
  return m;
}

function hasMergeIn(s: Sheet, rg: Range): boolean { return s.merges.some((m) => m.r1 <= rg.r2 && m.r2 >= rg.r1 && m.c1 <= rg.c2 && m.c2 >= rg.c1); }

function moved(cell: Cell | undefined, dr: number, dc: number): Cell | undefined {
  if (!cell) return undefined;
  const x: Cell = { ...cell };
  delete x.sp;
  if (x.f != null) x.f = shiftFormula(x.f, dr, dc);
  return x;
}

function transposeInPlace(app: App): void {
  if (isEmptySheet(app)) return;
  const s = app.sheet;
  const rg = workRegion(app);
  if (hasMergeIn(s, rg)) { toast("Transpose can't change merged cells. Unmerge them first."); return; }
  const tgt = { r1: rg.r1, c1: rg.c1, r2: rg.r1 + (rg.c2 - rg.c1), c2: rg.c1 + (rg.r2 - rg.r1) };
  if (!checkWritable(app, [rg, tgt])) return;
  const m = blockOf(s, rg);
  const t = transpose(m);
  app.book.transact('Transpose in Place', () => {
    for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) s.set(r, c, undefined);
    t.forEach((row, i) => row.forEach((cell, j) => { const r = tgt.r1 + i, c = tgt.c1 + j; s.set(r, c, moved(cell, r - (rg.r1 + j), c - (rg.c1 + i))); }));
  });
  app.grid.select(tgt.r1, tgt.c1, { range: tgt });
}

function reverseRange(app: App, dir: 'rows' | 'cols'): void {
  if (isEmptySheet(app)) return;
  const s = app.sheet;
  const rg = workRegion(app);
  if (hasMergeIn(s, rg)) { toast('Unmerge the cells in the selection first.'); return; }
  if (!checkWritable(app, [rg])) return;
  const m = blockOf(s, rg);
  app.book.transact(dir === 'rows' ? 'Reverse Rows' : 'Reverse Columns', () => {
    for (let i = 0; i < m.length; i++) for (let j = 0; j < m[0].length; j++) {
      const si = dir === 'rows' ? m.length - 1 - i : i, sj = dir === 'cols' ? m[0].length - 1 - j : j;
      s.set(rg.r1 + i, rg.c1 + j, moved(m[si][sj], i - si, j - sj));
    }
  });
}

function shuffleRows(app: App): void {
  if (isEmptySheet(app)) return;
  const s = app.sheet;
  const rg = workRegion(app);
  const top = hasHeaderRow(s, rg) ? rg.r1 + 1 : rg.r1;
  if (rg.r2 <= top) { toast('Select two or more rows to shuffle.'); return; }
  if (hasMergeIn(s, rg)) { toast('Unmerge the cells in the selection first.'); return; }
  if (!checkWritable(app, [rg])) return;
  const body = { ...rg, r1: top };
  const m = blockOf(s, body);
  const order = shuffled(m.map((_, i) => i));
  app.book.transact('Shuffle Rows', () => {
    order.forEach((src, i) => m[src].forEach((cell, j) => s.set(body.r1 + i, body.c1 + j, moved(cell, i - src, 0))));
  });
}

function roundValues(app: App, digits: number, mode: 'round' | 'up' | 'down' | 'sig', formulas: boolean): void {
  const s = app.sheet;
  if (!checkWritable(app, selRanges(app))) return;
  const fnName = mode === 'up' ? 'ROUNDUP' : mode === 'down' ? 'ROUNDDOWN' : 'ROUND';
  const rnd = (x: number) => {
    if (mode === 'sig') return x === 0 ? 0 : +x.toPrecision(Math.max(1, Math.min(15, digits)));
    const p = Math.pow(10, digits);
    const y = Math.abs(x) * p;
    const r = mode === 'up' ? Math.ceil(y - 1e-9) : mode === 'down' ? Math.floor(y + 1e-9) : Math.round(y + 1e-9);
    return Math.sign(x) * r / p;
  };
  let n = 0;
  app.book.transact('Round Values', () => {
    for (const rg of workRanges(app)) s.forEachIn(rg, (r, c, cell) => {
      if (!canChange(s, r, c)) return;
      if (cell.f != null) {
        if (!formulas || mode === 'sig' || /^ROUND(UP|DOWN)?\(/i.test(cell.f)) return;
        s.set(r, c, { ...cell, f: `${fnName}(${cell.f},${digits})` });
        n++;
      } else if (typeof cell.v === 'number') { const v = rnd(cell.v); if (v !== cell.v) { s.set(r, c, { ...cell, v }); n++; } }
    });
  });
  toast(n ? `${n} value${n === 1 ? '' : 's'} rounded` : 'Nothing to round.');
}

async function compareSheets(app: App, arg?: { a: string; b: string; formulas?: boolean }): Promise<void> {
  const book = app.book;
  if (book.sheets.length < 2) { toast('The workbook needs at least two sheets to compare.'); return; }
  if (!structureCheck(app)) return;
  let o = arg;
  if (!o) {
    const names = book.sheets.map((x) => x.name);
    const r = await openDialog({ title: 'Compare Sheets', fields: [
      { key: 'a', label: 'Original sheet', type: 'select', options: names, value: names[book.activeSheet], half: true },
      { key: 'b', label: 'Compare with', type: 'select', options: names, value: names[(book.activeSheet + 1) % names.length], half: true },
      { key: 'formulas', label: 'Also compare formulas (not just results)', type: 'checkbox', value: true },
    ], validate: (v) => (v.a === v.b ? 'Choose two different sheets.' : null) });
    if (!r) return;
    o = { a: r.a, b: r.b, formulas: !!r.formulas };
  }
  const A = book.sheetByName(o.a), B = book.sheetByName(o.b);
  if (!A || !B) { toast('Sheet not found.'); return; }
  const diffs = compareSheetCells(A, B, !!o.formulas);
  if (!diffs.length) { toast(`"${A.name}" and "${B.name}" have the same contents.`); return; }
  const rows: OutCell[][] = [[{ v: 'Cell', style: HEAD }, { v: A.name, style: HEAD }, { v: B.name, style: HEAD }, { v: 'Change', style: HEAD }]];
  for (const d of diffs.slice(0, 100000)) rows.push([cellName(d.r, d.c), d.a, d.b, d.kind]);
  addSheetWithRows(app, `Compare ${A.name}-${B.name}`, rows, { colW: [70, 180, 180, 90], freeze: 1, after: (sh) => {
    // A second block on the right mirrors the sheet layout with the differing cells highlighted.
    const used = { r2: Math.max(A.usedRange(false)?.r2 ?? 0, B.usedRange(false)?.r2 ?? 0), c2: Math.max(A.usedRange(false)?.c2 ?? 0, B.usedRange(false)?.c2 ?? 0) };
    const off = 5;
    writeOut(sh, 0, off, { v: `Layout of ${B.name} (differences highlighted)`, style: BOLD });
    const hl = sh.book.styleId({ fill: '#ffe08a' });
    const keyset = new Set(diffs.map((d) => cellKey(d.r, d.c)));
    for (let r = 0; r <= Math.min(used.r2, 2000); r++) for (let c = 0; c <= Math.min(used.c2, 200); c++) {
      const v = B.value(r, c);
      const diff = keyset.has(cellKey(r, c));
      if (v == null && !diff) continue;
      sh.set(r + 1, off + c, { v: diff ? `${fmtVal(A.value(r, c))} → ${fmtVal(v)}` : v, ...(diff ? { s: hl } : {}) });
    }
  } });
  toast(`${diffs.length.toLocaleString()} difference${diffs.length === 1 ? '' : 's'} found`);
}

const fmtVal = (v: Scalar) => (v == null ? '(blank)' : v instanceof CellError ? v.code : String(v));

export function compareSheetCells(A: Sheet, B: Sheet, formulas = true): { r: number; c: number; a: Scalar; b: Scalar; kind: string }[] {
  const keys = new Set<number>([...A.cells.keys(), ...B.cells.keys()]);
  const out: { r: number; c: number; a: Scalar; b: Scalar; kind: string }[] = [];
  const same = (x: Scalar, y: Scalar) => (x instanceof CellError && y instanceof CellError ? x.code === y.code : x === y || ((x == null || x === '') && (y == null || y === '')));
  for (const k of [...keys].sort((a, b) => a - b)) {
    const a = A.cells.get(k), b = B.cells.get(k);
    const va = a?.v ?? null, vb = b?.v ?? null;
    const r = Math.floor(k / 16384), c = k % 16384;
    if (!same(va, vb)) out.push({ r, c, a: a?.f != null && formulas ? '=' + a.f : va, b: b?.f != null && formulas ? '=' + b.f : vb, kind: va == null || va === '' ? 'Added' : vb == null || vb === '' ? 'Removed' : 'Changed' });
    else if (formulas && (a?.f ?? null) !== (b?.f ?? null)) out.push({ r, c, a: a?.f != null ? '=' + a.f : va, b: b?.f != null ? '=' + b.f : vb, kind: 'Formula' });
  }
  return out;
}

async function findDuplicates(app: App, arg?: { cols?: number[]; mode?: 'highlight' | 'highlightAll' | 'list'; header?: boolean; matchCase?: boolean; color?: string }): Promise<void> {
  if (isEmptySheet(app)) return;
  const s = app.sheet;
  const rg = workRegion(app);
  let o = arg;
  const header0 = hasHeaderRow(s, rg);
  if (!o) {
    const r = await openDialog({ title: 'Find Duplicates', help: `Range ${rangeName(rg)}`, fields: [
      { key: 'header', label: 'My data has headers', type: 'checkbox', value: header0 },
      { key: 'cols', label: 'Rows are duplicates when these columns match', type: 'list', multiple: true, rows: Math.min(8, rg.c2 - rg.c1 + 1), options: columnOptions(s, rg, header0), value: columnOptions(s, rg, header0).map((x) => x.value) },
      { key: 'mode', label: 'Then', type: 'radio', value: 'highlight', options: [{ value: 'highlight', label: 'Highlight repeats (not the first occurrence)' }, { value: 'highlightAll', label: 'Highlight every occurrence' }, { value: 'list', label: 'List the duplicate rows on a new sheet' }] },
      { key: 'color', label: 'Highlight color', type: 'color', value: '#ffc7ce', visibleIf: (v) => v.mode !== 'list', half: true },
      { key: 'matchCase', label: 'Match case', type: 'checkbox', value: false, half: true },
    ] });
    if (!r) return;
    o = { cols: (r.cols as string[]).map(Number), mode: r.mode, header: !!r.header, matchCase: !!r.matchCase, color: r.color };
  }
  const cols = o.cols?.length ? o.cols : Array.from({ length: rg.c2 - rg.c1 + 1 }, (_, i) => rg.c1 + i);
  const r1 = o.header ? rg.r1 + 1 : rg.r1;
  const keyOf = (r: number) => cols.map((c) => { const v = s.value(r, c); const t = v == null ? '' : String(v instanceof CellError ? v.code : v).trim(); return o!.matchCase ? t : t.toLowerCase(); }).join('\u0001');
  const groups = new Map<string, number[]>();
  for (let r = r1; r <= rg.r2; r++) { if (sheetRowHidden(s, r)) continue; const k = keyOf(r); if (k.replace(/\u0001/g, '') === '') continue; const g = groups.get(k) || []; g.push(r); groups.set(k, g); }
  const dupGroups = [...groups.values()].filter((g) => g.length > 1);
  if (!dupGroups.length) { toast('No duplicates found.'); return; }
  const rowsToMark = dupGroups.flatMap((g) => (o!.mode === 'highlightAll' ? g : g.slice(1)));
  if (o.mode === 'list') {
    if (!structureCheck(app)) return;
    const out: OutCell[][] = [[{ v: 'Row', style: HEAD }, { v: 'Group', style: HEAD }, ...Array.from({ length: rg.c2 - rg.c1 + 1 }, (_, i) => ({ v: o!.header ? s.value(rg.r1, rg.c1 + i) ?? colName(rg.c1 + i) : colName(rg.c1 + i), style: HEAD }))]];
    dupGroups.forEach((g, gi) => g.forEach((r) => out.push([r + 1, gi + 1, ...Array.from({ length: rg.c2 - rg.c1 + 1 }, (_, i) => s.value(r, rg.c1 + i))])));
    addSheetWithRows(app, 'Duplicates', out, { freeze: 1 });
  } else {
    if (!checkWritable(app, [rg])) return;
    app.book.transact('Find Duplicates', () => {
      for (const r of rowsToMark) for (let c = rg.c1; c <= rg.c2; c++) {
        const cur = s.get(r, c);
        s.set(r, c, { ...(cur || {}), s: app.book.derivedStyle(cur?.s ?? 0, { fill: o!.color || '#ffc7ce' }) });
      }
    });
  }
  toast(`${dupGroups.length} group${dupGroups.length === 1 ? '' : 's'} of duplicates; ${rowsToMark.length} row${rowsToMark.length === 1 ? '' : 's'} ${o.mode === 'list' ? 'listed' : 'highlighted'}`);
}

async function fuzzyMatchCmd(app: App, arg?: { lookup: string; list: string; algo?: FuzzyAlgo; threshold?: number }): Promise<void> {
  if (isEmptySheet(app)) return;
  const s = app.sheet;
  let o = arg;
  if (!o) {
    const sel = clipToUsed(s, activeRange(app));
    const r = await openDialog({ title: 'Fuzzy Match', help: 'For each value in the first range, finds the most similar value in the second. Results go to the two columns right of the first range.', fields: [
      { key: 'lookup', label: 'Values to match', type: 'text', value: rangeRefText(s, sel, false) },
      { key: 'list', label: 'Match against (another column or sheet)', type: 'text', value: '', placeholder: "e.g. Sheet2!A2:A500" },
      { key: 'algo', label: 'Method', type: 'select', value: 'jaroWinkler', options: [{ value: 'jaroWinkler', label: 'Jaro-Winkler (names, short text)' }, { value: 'levenshtein', label: 'Edit distance (typos)' }, { value: 'tokens', label: 'Words in any order (company names, addresses)' }] },
      { key: 'threshold', label: 'Minimum score to accept (0–1)', type: 'number', value: 0.8, min: 0, max: 1, step: 0.05 },
    ] });
    if (!r) return;
    o = { lookup: r.lookup, list: r.list, algo: r.algo, threshold: Number(r.threshold ?? 0.8) };
  }
  const { resolveRangeText } = await import('../model/structure');
  const a = resolveRangeText(app.book, o.lookup, s), b = resolveRangeText(app.book, o.list, s);
  if (!a || !b) { toast('Enter both ranges, such as A2:A100 and Sheet2!B2:B500.'); return; }
  const ra = clipToUsed(a.sheet, a.range), rb = clipToUsed(b.sheet, b.range);
  const cands: string[] = [];
  for (let r = rb.r1; r <= rb.r2; r++) for (let c = rb.c1; c <= rb.c2; c++) { const v = b.sheet.value(r, c); cands.push(v == null ? '' : String(v)); }
  const target = { r1: ra.r1, c1: ra.c2 + 1, r2: ra.r2, c2: ra.c2 + 2 };
  if (a.sheet !== app.sheet) { toast('The values to match must be on the active sheet.'); return; }
  if (!(await confirmOverwrite(s, target)) || !checkWritable(app, [target])) return;
  let matched = 0;
  app.book.transact('Fuzzy Match', () => {
    for (let r = ra.r1; r <= ra.r2; r++) {
      const v = s.value(r, ra.c1);
      if (v == null || v === '') continue;
      const m = bestMatch(String(v), cands, o!.algo);
      if (!m) continue;
      const ok = m.score >= (o!.threshold ?? 0.8);
      writeOut(s, r, target.c1, ok ? m.match : null);
      writeOut(s, r, target.c1 + 1, { v: Math.round(m.score * 1000) / 1000, style: { numFmt: '0.000', color: ok ? undefined : '#9c0006' } });
      if (ok) matched++;
    }
  });
  toast(`${matched} value${matched === 1 ? '' : 's'} matched; best match and score are in ${colName(target.c1)}:${colName(target.c2)}`);
}

function regionWithHeader(app: App): { s: Sheet; rg: Range; header: boolean; r1: number } {
  const s = app.sheet;
  const rg = workRegion(app);
  const header = hasHeaderRow(s, rg);
  return { s, rg, header, r1: header ? rg.r1 + 1 : rg.r1 };
}

const headName = (s: Sheet, rg: Range, header: boolean, c: number) => (header ? cellText(s, rg.r1, c) || colName(c) : `Column ${colName(c)}`);

function dataProfile(app: App): void {
  if (isEmptySheet(app) || !structureCheck(app)) return;
  const { s, rg, header, r1 } = regionWithHeader(app);
  const rows: OutCell[][] = [['Column', 'Type', 'Values', 'Numbers', 'Text', 'Logical', 'Errors', 'Blanks', 'Distinct', 'Minimum', 'Maximum', 'Mean', 'Most common'].map((v) => ({ v, style: HEAD }))];
  for (let c = rg.c1; c <= rg.c2; c++) {
    const vals: Scalar[] = [];
    for (let r = r1; r <= rg.r2; r++) if (!sheetRowHidden(s, r)) vals.push(s.value(r, c));
    const p = profileColumn(headName(s, rg, header, c), vals);
    rows.push([p.name, p.type, vals.length - p.blanks, p.numbers, p.text, p.booleans, p.errors, p.blanks, p.distinct, p.min, p.max, p.mean == null ? null : { v: p.mean, style: { numFmt: '#,##0.00' } }, p.top.map((t) => `${t.value} (${t.count})`).join('; ')]);
  }
  addSheetWithRows(app, `Profile of ${s.name}`, rows, { colW: [140, 130, 60, 64, 50, 56, 50, 50, 60, 80, 80, 80, 260], freeze: 1 });
}

function confidenceT(app: App, n: number): number {
  if (n < 2) return 0;
  try { const v = app.engine.evaluateScalar(`T.INV.2T(0.05,${n - 1})`, app.sheet, 0, 0); if (typeof v === 'number') return v; } catch { /* fall through */ }
  return 1.96;
}

function numericColumns(s: Sheet, rg: Range, r1: number): number[] {
  const cols: number[] = [];
  for (let c = rg.c1; c <= rg.c2; c++) if (numbersIn(s, c, r1, rg.r2).length) cols.push(c);
  return cols;
}

function descriptiveStats(app: App): void {
  if (isEmptySheet(app) || !structureCheck(app)) return;
  const { s, rg, header, r1 } = regionWithHeader(app);
  const cols = numericColumns(s, rg, r1);
  if (!cols.length) { toast('The selection has no numbers.'); return; }
  const labels = ['Mean', 'Standard Error', 'Median', 'Mode', 'Standard Deviation', 'Sample Variance', 'Kurtosis', 'Skewness', 'Range', 'Minimum', 'Maximum', 'Sum', 'Count', 'Confidence Level (95.0%)'];
  const stats = cols.map((c) => { const xs = numbersIn(s, c, r1, rg.r2); const d = describe(xs); return [d.mean, d.stderr, d.median, d.mode, d.stdev, d.variance, d.kurtosis, d.skewness, d.range, d.min, d.max, d.sum, d.count, d.count > 1 ? confidenceT(app, d.count) * d.stderr : null]; });
  const rows: OutCell[][] = [[{ v: 'Statistic', style: HEAD }, ...cols.map((c) => ({ v: headName(s, rg, header, c), style: HEAD }))]];
  labels.forEach((l, i) => rows.push([{ v: l, style: BOLD }, ...stats.map((st) => (st[i] == null ? '#N/A' : { v: st[i] as number, style: { numFmt: i === 12 ? '0' : '#,##0.0000' } }))]));
  addSheetWithRows(app, 'Descriptive Statistics', rows, { colW: [170, ...cols.map(() => 110)], freeze: 1 });
}

async function histogramCmd(app: App, arg?: { bins?: number }): Promise<void> {
  if (isEmptySheet(app) || !structureCheck(app)) return;
  const { s, rg, header, r1 } = regionWithHeader(app);
  const sel = activeRange(app);
  const c = sel.c1 >= rg.c1 && sel.c1 <= rg.c2 ? sel.c1 : rg.c1;
  const xs = numbersIn(s, c, r1, rg.r2);
  if (xs.length < 2) { toast('Select a column with at least two numbers.'); return; }
  let bins = arg?.bins;
  if (bins == null) {
    const r = await openDialog({ title: 'Histogram', help: `${xs.length} numbers in ${headName(s, rg, header, c)}`, fields: [{ key: 'bins', label: 'Number of bins (blank = automatic)', type: 'number', value: '', min: 1, max: 100 }] });
    if (!r) return;
    bins = r.bins ? Number(r.bins) : undefined;
  }
  const hst = histogram(xs, bins);
  let cum = 0;
  const rows: OutCell[][] = [['Bin (up to)', 'Frequency', 'Cumulative %'].map((v) => ({ v, style: HEAD }))];
  hst.bounds.forEach((b, i) => { cum += hst.counts[i]; rows.push([{ v: b, style: { numFmt: '#,##0.##' } }, hst.counts[i], { v: cum / xs.length, style: { numFmt: '0.0%' } }]); });
  addSheetWithRows(app, `Histogram ${headName(s, rg, header, c)}`, rows, { colW: [100, 80, 100], freeze: 1, after: (sh) => {
    const n = hst.bounds.length;
    const chart: ChartDef = { id: newId('ch'), type: 'column', title: `Histogram of ${headName(s, rg, header, c)}`, legend: 'none', gridlines: true,
      series: [{ name: 'Frequency', nameRef: rangeRefText(sh, { r1: 0, c1: 1, r2: 0, c2: 1 }), values: rangeRefText(sh, { r1: 1, c1: 1, r2: n, c2: 1 }), categories: rangeRefText(sh, { r1: 1, c1: 0, r2: n, c2: 0 }) }],
      at: { r: 1, c: 4, dx: 0, dy: 0, w: 480, h: 288 }, sourceRange: rangeRefText(sh, { r1: 0, c1: 0, r2: n, c2: 1 }) };
    sh.charts.push(chart);
  } });
}

function correlationCmd(app: App): void {
  if (isEmptySheet(app) || !structureCheck(app)) return;
  const { s, rg, header, r1 } = regionWithHeader(app);
  const cols = numericColumns(s, rg, r1);
  if (cols.length < 2) { toast('Select data with at least two numeric columns.'); return; }
  // Pairwise complete rows
  const col = (c: number) => { const xs: (number | null)[] = []; for (let r = r1; r <= rg.r2; r++) { if (sheetRowHidden(s, r)) continue; const v = s.value(r, c); xs.push(typeof v === 'number' ? v : null); } return xs; };
  const data = cols.map(col);
  const rows: OutCell[][] = [[{ v: '', style: HEAD }, ...cols.map((c) => ({ v: headName(s, rg, header, c), style: HEAD }))]];
  cols.forEach((ci, i) => {
    const row: OutCell[] = [{ v: headName(s, rg, header, ci), style: HEAD }];
    cols.forEach((_cj, j) => {
      if (j > i) { row.push(null); return; }
      const xs: number[] = [], ys: number[] = [];
      data[i].forEach((x, k) => { const y = data[j][k]; if (x != null && y != null) { xs.push(x); ys.push(y); } });
      const p = i === j ? 1 : pearson(xs, ys);
      row.push(p == null ? '#N/A' : { v: Math.round(p * 10000) / 10000, style: { numFmt: '0.0000' } });
    });
    rows.push(row);
  });
  addSheetWithRows(app, 'Correlation', rows, { colW: [140, ...cols.map(() => 100)], freeze: 1, after: (sh) => {
    const rule: CFRule = { id: newId('cf'), type: 'colorScale', ranges: [{ r1: 1, c1: 1, r2: cols.length, c2: cols.length }], colors: ['#5a8ac6', '#fcfcff', '#f8696b'], stops: [{ type: 'num', value: -1 }, { type: 'num', value: 0 }, { type: 'num', value: 1 }] };
    sh.cf.push(rule);
  } });
}

async function groupSummaryCmd(app: App, arg?: { key: number; value: number; fn: 'sum' | 'count' | 'average' | 'min' | 'max' }): Promise<void> {
  if (isEmptySheet(app) || !structureCheck(app)) return;
  const { s, rg, header, r1 } = regionWithHeader(app);
  let o = arg;
  if (!o) {
    const opts = columnOptions(s, rg, header);
    const r = await openDialog({ title: 'Summarize by Category', fields: [
      { key: 'key', label: 'Group rows by', type: 'select', options: opts, value: opts[0]?.value },
      { key: 'value', label: 'Summarize', type: 'select', options: opts, value: opts[Math.min(1, opts.length - 1)]?.value },
      { key: 'fn', label: 'Using', type: 'select', value: 'sum', options: [{ value: 'sum', label: 'Sum' }, { value: 'count', label: 'Count' }, { value: 'average', label: 'Average' }, { value: 'min', label: 'Minimum' }, { value: 'max', label: 'Maximum' }] },
    ] });
    if (!r) return;
    o = { key: Number(r.key), value: Number(r.value), fn: r.fn };
  }
  const keys: Scalar[] = [], vals: Scalar[] = [];
  for (let r = r1; r <= rg.r2; r++) { if (sheetRowHidden(s, r)) continue; keys.push(s.value(r, o.key)); vals.push(s.value(r, o.value)); }
  const g = groupSummary(keys, vals, o.fn);
  const fnLabel = { sum: 'Sum', count: 'Count', average: 'Average', min: 'Min', max: 'Max' }[o.fn];
  const rows: OutCell[][] = [[{ v: headName(s, rg, header, o.key), style: HEAD }, { v: `${fnLabel} of ${headName(s, rg, header, o.value)}`, style: HEAD }]];
  for (const [k, v] of g) rows.push([k, { v, style: { numFmt: o.fn === 'count' ? '0' : '#,##0.00' } }]);
  rows.push([{ v: 'Total', style: BOLD }, { f: o.fn === 'average' ? `AVERAGE(B2:B${g.length + 1})` : o.fn === 'min' ? `MIN(B2:B${g.length + 1})` : o.fn === 'max' ? `MAX(B2:B${g.length + 1})` : `SUM(B2:B${g.length + 1})`, style: { bold: true, numFmt: o.fn === 'count' ? '0' : '#,##0.00', bt: { style: 'thin' } } }]);
  addSheetWithRows(app, `By ${headName(s, rg, header, o.key)}`, rows, { colW: [180, 140], freeze: 1 });
}

function fillTarget(app: App): Range | null {
  const s = app.sheet;
  const rg = activeRange(app);
  if (isWholeCols(rg)) return { ...rg, r2: Math.max(rg.r1, (s.usedRange()?.r2 ?? 99)) };
  if (rg.c2 - rg.c1 > 200 || rg.r2 - rg.r1 > 100000) return null;
  return rg;
}

async function randomData(app: App, arg?: { kind: string; min?: number; max?: number; decimals?: number; mean?: number; sd?: number; list?: string; from?: string; to?: string; unique?: boolean; count?: number }): Promise<void> {
  let o = arg;
  const sel = activeRange(app);
  const single = sel.r1 === sel.r2 && sel.c1 === sel.c2;
  if (!o) {
    const r = await openDialog({ title: 'Random Data', help: single ? 'Fills down from the active cell.' : `Fills ${rangeName(sel)}.`, fields: [
      { key: 'kind', label: 'Fill with', type: 'select', value: 'integers', options: [
        { value: 'integers', label: 'Whole numbers' }, { value: 'decimals', label: 'Decimal numbers' }, { value: 'normal', label: 'Normal distribution' },
        { value: 'dates', label: 'Dates' }, { value: 'bool', label: 'TRUE / FALSE' }, { value: 'yesno', label: 'Yes / No' },
        { value: 'first', label: 'First names' }, { value: 'last', label: 'Last names' }, { value: 'full', label: 'Full names' }, { value: 'emails', label: 'Email addresses' },
        { value: 'cities', label: 'Cities' }, { value: 'countries', label: 'Countries' }, { value: 'companies', label: 'Company names' }, { value: 'products', label: 'Product names' },
        { value: 'list', label: 'Choices from my list' }, { value: 'guid', label: 'Unique IDs (GUID)' }, { value: 'password', label: 'Random passwords' }] },
      { key: 'min', label: 'From', type: 'number', value: 1, half: true, visibleIf: (v) => v.kind === 'integers' || v.kind === 'decimals' },
      { key: 'max', label: 'To', type: 'number', value: 100, half: true, visibleIf: (v) => v.kind === 'integers' || v.kind === 'decimals' },
      { key: 'decimals', label: 'Decimal places', type: 'number', value: 2, min: 0, max: 10, visibleIf: (v) => v.kind === 'decimals' || v.kind === 'normal' },
      { key: 'mean', label: 'Mean', type: 'number', value: 50, half: true, visibleIf: (v) => v.kind === 'normal' },
      { key: 'sd', label: 'Standard deviation', type: 'number', value: 10, half: true, visibleIf: (v) => v.kind === 'normal' },
      { key: 'from', label: 'From date', type: 'date', value: `${new Date().getFullYear()}-01-01`, half: true, visibleIf: (v) => v.kind === 'dates' },
      { key: 'to', label: 'To date', type: 'date', value: `${new Date().getFullYear()}-12-31`, half: true, visibleIf: (v) => v.kind === 'dates' },
      { key: 'list', label: 'Choices (comma-separated)', type: 'textarea', rows: 2, value: 'Red, Green, Blue', visibleIf: (v) => v.kind === 'list' },
      { key: 'unique', label: 'Unique values only', type: 'checkbox', value: false, visibleIf: (v) => v.kind === 'integers' },
      { key: 'count', label: 'How many rows', type: 'number', value: 20, min: 1, max: 100000, visibleIf: () => single },
    ] });
    if (!r) return;
    o = r as any;
  }
  const opt = o!;
  const s = app.sheet;
  let rg = single ? { r1: sel.r1, c1: sel.c1, r2: sel.r1 + Math.max(1, Number(opt.count) || 20) - 1, c2: sel.c1 } : fillTarget(app);
  if (!rg) { toast('Select a smaller range.'); return; }
  rg = { ...rg };
  if (!checkWritable(app, [rg])) return;
  if (!(await confirmOverwrite(s, rg))) return;
  const pick = <X,>(xs: X[]) => xs[Math.floor(Math.random() * xs.length)];
  const min = Number(opt.min ?? 1), max = Number(opt.max ?? 100), dec = Number(opt.decimals ?? 2);
  const n = (rg.r2 - rg.r1 + 1) * (rg.c2 - rg.c1 + 1);
  let pool: number[] | null = null;
  if (opt.kind === 'integers' && opt.unique) {
    const lo = Math.ceil(Math.min(min, max)), hi = Math.floor(Math.max(min, max));
    if (hi - lo + 1 < n) { toast(`Only ${hi - lo + 1} unique whole numbers exist between ${lo} and ${hi}.`); return; }
    const set = new Set<number>();
    while (set.size < n) set.add(lo + Math.floor(Math.random() * (hi - lo + 1)));
    pool = [...set];
  }
  const d0 = opt.from ? parseTextDate(opt.from) : null, d1 = opt.to ? parseTextDate(opt.to) : null;
  const list = String(opt.list || '').split(/\s*[,\n]\s*/).filter(Boolean);
  let i = 0;
  const gen = (): OutCell => {
    switch (opt.kind) {
      case 'integers': return pool ? pool[i++] : Math.floor(Math.min(min, max) + Math.random() * (Math.abs(max - min) + 1));
      case 'decimals': return +(Math.min(min, max) + Math.random() * Math.abs(max - min)).toFixed(dec);
      case 'normal': return +(Number(opt.mean ?? 0) + gaussian() * Number(opt.sd ?? 1)).toFixed(dec);
      case 'dates': { const a = d0 ?? 45292, b = d1 ?? a + 365; return { v: Math.floor(Math.min(a, b) + Math.random() * (Math.abs(b - a) + 1)), style: { numFmt: 'yyyy-mm-dd' } }; }
      case 'bool': return Math.random() < 0.5;
      case 'yesno': return Math.random() < 0.5 ? 'Yes' : 'No';
      case 'first': return pick(FIRST_NAMES);
      case 'last': return pick(LAST_NAMES);
      case 'full': return `${pick(FIRST_NAMES)} ${pick(LAST_NAMES)}`;
      case 'emails': return `${pick(FIRST_NAMES).toLowerCase()}.${pick(LAST_NAMES).toLowerCase()}@example.${pick(['com', 'org', 'net'])}`;
      case 'cities': return pick(CITIES);
      case 'countries': return pick(COUNTRIES);
      case 'companies': return pick(COMPANIES);
      case 'products': return pick(PRODUCTS);
      case 'list': return list.length ? pick(list) : null;
      case 'guid': return guid();
      case 'password': { const cs = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789!@#$%&*'; const a = new Uint32Array(14); crypto.getRandomValues(a); return [...a].map((x) => cs[x % cs.length]).join(''); }
    }
    return null;
  };
  app.book.transact('Random Data', () => { for (let r = rg!.r1; r <= rg!.r2; r++) for (let c = rg!.c1; c <= rg!.c2; c++) if (canChange(s, r, c)) writeOut(s, r, c, gen()); });
  if (single) app.grid.select(rg.r1, rg.c1, { range: rg });
}

async function sequenceCmd(app: App, arg?: { kind: string; start?: string; step?: number; count?: number; dir?: 'down' | 'right'; prefix?: string; pad?: number; suffix?: string }): Promise<void> {
  let o = arg;
  if (!o) {
    const r = await openDialog({ title: 'Sequence', help: 'Fills from the active cell.', fields: [
      { key: 'kind', label: 'Type', type: 'select', value: 'number', options: [{ value: 'number', label: 'Numbers' }, { value: 'code', label: 'Numbered codes (INV-0001)' }, { value: 'day', label: 'Dates (days)' }, { value: 'weekday', label: 'Workdays (Mon–Fri)' }, { value: 'month', label: 'Months' }, { value: 'year', label: 'Years' }, { value: 'time', label: 'Times (minutes)' }] },
      { key: 'start', label: 'Start', type: 'text', value: '1', half: true, placeholder: 'number or date' },
      { key: 'step', label: 'Step', type: 'number', value: 1, half: true },
      { key: 'count', label: 'How many', type: 'number', value: 10, min: 1, max: 100000, half: true },
      { key: 'dir', label: 'Direction', type: 'select', value: 'down', options: [{ value: 'down', label: 'Down a column' }, { value: 'right', label: 'Across a row' }], half: true },
      { key: 'prefix', label: 'Prefix', type: 'text', value: 'INV-', half: true, visibleIf: (v) => v.kind === 'code' },
      { key: 'pad', label: 'Digits', type: 'number', value: 4, min: 1, max: 12, half: true, visibleIf: (v) => v.kind === 'code' },
      { key: 'suffix', label: 'Suffix', type: 'text', value: '', visibleIf: (v) => v.kind === 'code' },
    ] });
    if (!r) return;
    o = r as any;
  }
  const opt = o!;
  const s = app.sheet;
  const { r: r0, c: c0 } = activeCell(app);
  // Never past the last column or row (a column index past the end would land in the next row).
  const count = Math.max(1, Math.min(100000, Number(opt.count) || 10, opt.dir === 'right' ? MAX_COLS - c0 : MAX_ROWS - r0));
  const step = Number(opt.step ?? 1) || 1;
  const rg = opt.dir === 'right' ? { r1: r0, c1: c0, r2: r0, c2: c0 + count - 1 } : { r1: r0, c1: c0, r2: r0 + count - 1, c2: c0 };
  if (!checkWritable(app, [rg]) || !(await confirmOverwrite(s, rg))) return;
  const startText = String(opt.start ?? '1');
  const isDate = ['day', 'weekday', 'month', 'year'].includes(opt.kind);
  const startDate = isDate ? parseTextDate(startText, { order: app.prefs.dateOrder === 'dmy' ? 'dmy' : 'auto' }) : null;
  if (isDate && startDate == null && opt.kind !== 'year') { toast('Enter a start date such as 2025-01-31.'); return; }
  const startNum = textToNumber(startText) ?? 1;
  app.book.transact('Sequence', () => {
    for (let i = 0; i < count; i++) {
      const r = opt.dir === 'right' ? r0 : r0 + i, c = opt.dir === 'right' ? c0 + i : c0;
      let out: OutCell;
      switch (opt.kind) {
        case 'code': out = `${opt.prefix || ''}${String(Math.round(startNum + i * step)).padStart(Number(opt.pad) || 1, '0')}${opt.suffix || ''}`; break;
        case 'day': out = { v: startDate! + i * step, style: { numFmt: 'yyyy-mm-dd' } }; break;
        case 'weekday': out = { v: addToDate(startDate!, i * step, 'workdays'), style: { numFmt: 'ddd yyyy-mm-dd' } }; break;
        case 'month': out = { v: addToDate(startDate!, i * step, 'months'), style: { numFmt: 'mmm yyyy' } }; break;
        case 'year': out = startDate != null ? { v: addToDate(startDate, i * step, 'years'), style: { numFmt: 'yyyy-mm-dd' } } : startNum + i * step; break;
        case 'time': out = { v: ((textToNumber(startText) ?? 0) + i * step) / 1440, style: { numFmt: 'h:mm' } }; break;
        default: out = +(startNum + i * step).toPrecision(15);
      }
      writeOut(s, r, c, out);
    }
  });
  app.grid.select(r0, c0, { range: rg });
}

async function calendarCmd(app: App, arg?: { year: number; month: number; weekStart?: number }): Promise<void> {
  if (!structureCheck(app)) return;
  let o = arg;
  const now = new Date();
  if (!o) {
    const r = await openDialog({ title: 'Calendar', fields: [
      { key: 'year', label: 'Year', type: 'number', value: now.getFullYear(), min: 1900, max: 9999, half: true },
      { key: 'month', label: 'Month', type: 'select', value: String(now.getMonth() + 1), half: true, options: [{ value: '0', label: 'Whole year' }, ...MONTHS.map((m, i) => ({ value: String(i + 1), label: m }))] },
      { key: 'weekStart', label: 'Weeks start on', type: 'select', value: '0', options: [{ value: '0', label: 'Sunday' }, { value: '1', label: 'Monday' }] },
    ] });
    if (!r) return;
    o = { year: Number(r.year), month: Number(r.month), weekStart: Number(r.weekStart) };
  }
  const year = o.year, ws = o.weekStart || 0;
  const months = o.month ? [o.month] : Array.from({ length: 12 }, (_, i) => i + 1);
  const rows: OutCell[][] = [];
  const dayHead = Array.from({ length: 7 }, (_, i) => ({ v: DAYS[(i + ws) % 7].slice(0, 3), style: { bold: true, hAlign: 'center' as const, fill: '#dde7f0' } }));
  for (const m of months) {
    rows.push([{ v: `${MONTHS[m - 1]} ${year}`, style: { bold: true, size: 14 } }]);
    rows.push(dayHead);
    for (const week of monthGrid(year, m, ws)) rows.push(week.map((d) => (d ? { v: partsToSerial(year, m, d), style: { numFmt: 'd', hAlign: 'center' as const, vAlign: 'top' as const } } : null)));
    rows.push([]);
  }
  addSheetWithRows(app, o.month ? `${MONTHS[o.month - 1].slice(0, 3)} ${year}` : `Calendar ${year}`, rows, { colW: Array(7).fill(72) });
}

async function dateCalculator(app: App): Promise<void> {
  const today = new Date();
  const iso = (d: Date) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  const toIso = (serial: number) => { const p = serialToParts(serial); return `${p.y}-${String(p.m).padStart(2, '0')}-${String(p.d).padStart(2, '0')}`; };
  const v0 = app.sheet.value(app.sheet.sel.r, app.sheet.sel.c);
  const start = typeof v0 === 'number' && v0 > 1 && v0 < 2958465 ? toIso(v0) : iso(today);
  let result: { value: number; isDate: boolean } | null = null;
  const compute = (v: Record<string, any>): string => {
    const a = parseTextDate(String(v.a || '')), b = parseTextDate(String(v.b || ''));
    if (v.mode === 'diff') {
      if (a == null || b == null) { result = null; return '<i>Enter two dates.</i>'; }
      const d = ymdDiff(a, b);
      const days = b - a;
      result = { value: days, isDate: false };
      return `<table class="ql-stats"><tr><td>Days</td><td><b>${days.toLocaleString()}</b></td></tr><tr><td>Workdays (Mon–Fri, inclusive)</td><td>${businessDays(a, b).toLocaleString()}</td></tr><tr><td>Weeks</td><td>${(days / 7).toFixed(2)}</td></tr><tr><td>Years, months, days</td><td>${d.y} y ${d.m} m ${d.d} d</td></tr><tr><td>Years (decimal)</td><td>${(days / 365.25).toFixed(3)}</td></tr></table>`;
    }
    if (a == null) { result = null; return '<i>Enter a start date.</i>'; }
    const n = (Number(v.n) || 0) * (v.mode === 'sub' ? -1 : 1);
    const out = addToDate(a, n, v.unit);
    result = { value: out, isDate: true };
    const dt = serialToParts(out);
    return `<p><b>${toIso(out)}</b> (${DAYS[dt.dow]}, ${MONTHS[dt.m - 1]} ${dt.d}, ${dt.y})</p>`;
  };
  const r = await openDialog({
    title: 'Date Calculator', preview: true, ok: 'Insert Result', cancel: 'Close',
    fields: [
      { key: 'mode', label: 'Calculate', type: 'radio', value: 'diff', options: [{ value: 'diff', label: 'Difference between dates' }, { value: 'add', label: 'Add to a date' }, { value: 'sub', label: 'Subtract from a date' }] },
      { key: 'a', label: 'Start date', type: 'date', value: start, half: true },
      { key: 'b', label: 'End date', type: 'date', value: iso(today), half: true, visibleIf: (v) => v.mode === 'diff' },
      { key: 'n', label: 'Amount', type: 'number', value: 30, half: true, visibleIf: (v) => v.mode !== 'diff' },
      { key: 'unit', label: 'Unit', type: 'select', value: 'days', half: true, options: [{ value: 'days', label: 'Days' }, { value: 'workdays', label: 'Workdays' }, { value: 'weeks', label: 'Weeks' }, { value: 'months', label: 'Months' }, { value: 'years', label: 'Years' }], visibleIf: (v) => v.mode !== 'diff' },
    ],
    onChange: (v, ctl) => ctl.setPreview(compute(v)),
  });
  if (!r || !result) return;
  const res: { value: number; isDate: boolean } = result;
  const { r: rr, c } = activeCell(app);
  if (app.isReadOnly() || !canChange(app.sheet, rr, c)) { toast('This cell cannot be changed.'); return; }
  app.book.transact('Date Calculator', () => writeOut(app.sheet, rr, c, res.isDate ? { v: res.value, style: { numFmt: 'yyyy-mm-dd' } } : res.value));
}

void h;
