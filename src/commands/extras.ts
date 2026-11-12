/**
 * Extras tab: tools of Anne's Grid's own, not found in the add-ins the Utilities and Plus tabs follow.
 * Groups: Genealogy, Publishing, Finance, Data Quality, Dates, Geography, Workflow. The pure functions
 * are exported for the unit tests.
 */
import type { App } from '../app';
import type { Cell, CellStyle, Scalar } from '../model/types';
import { CellError } from '../model/types';
import { cellKey, keyRow, keyCol, colName, cellName, rangeName, type Range } from '../model/address';
import { serialToParts, partsToSerial, daysInMonth, jsDateToSerial, DAYS, MONTHS } from '../model/dates';
import { isDateFormat } from '../model/numfmt';
import { selRanges, activeRange, activeCell, clipToUsed } from './ops';
import { defTools, X, selCells, selBlock, singleCell, hasContent, selectHits, sheetProtected, emptySheet, numOf, textOf, plural, report, shownText, argOrAsk, columnChoices, headerGuess, writeOut, checkWritable, changeCells, patchStyle, insertRows, copyText, valueRows, addSheetWithRows, HEAD, naturalCompare } from './utilShared';
import { menuOf } from './utilities';
import { transformValues, textToNumber, parseTextDate } from './tools';
import { openDialog, toast, promptDialog, alertDialog } from '../ui/dialog';
import { insertRule } from './conditional';
import { kvKeys, kvGet } from '../util/store';

const G_GEN = 'Genealogy';
const G_PUB = 'Publishing';
const G_FIN = 'Finance';
const G_QUALITY = 'Data Quality';
const G_DATES = 'Dates';
const G_GEO = 'Geography';
const G_FLOW = 'Workflow';

// ============================================================================================
// Genealogy

/** American Soundex (NARA rules): Robert → R163, Tymczak → T522, Pfister → P236. */
export function soundex(name: string): string {
  const s = name.toUpperCase().replace(/[^A-Z]/g, '');
  if (!s) return '';
  const code = (ch: string): string => ('BFPV'.includes(ch) ? '1' : 'CGJKQSXZ'.includes(ch) ? '2' : 'DT'.includes(ch) ? '3' : ch === 'L' ? '4' : 'MN'.includes(ch) ? '5' : ch === 'R' ? '6' : 'HW'.includes(ch) ? '-' : '0');
  let out = s[0]; let last = code(s[0]);
  for (let i = 1; i < s.length && out.length < 4; i++) {
    const c = code(s[i]);
    if (c === '-') continue; // H and W do not separate same codes
    if (c !== '0' && c !== last) out += c;
    last = c;
  }
  return (out + '000').slice(0, 4);
}

/** Daitch–Mokotoff would be long; this is a compact Metaphone-like key for English surnames. */
export function metaphoneKey(name: string): string {
  let s = name.toUpperCase().replace(/[^A-Z]/g, '');
  if (!s) return '';
  s = s.replace(/^(KN|GN|PN|AE|WR)/, (m) => m[1]).replace(/^X/, 'S').replace(/^WH/, 'W');
  s = s.replace(/MB$/, 'M');
  const out: string[] = [];
  for (let i = 0; i < s.length; i++) {
    const ch = s[i], prev = s[i - 1] || '', next = s[i + 1] || '', next2 = s[i + 2] || '';
    if (ch === prev && ch !== 'C') continue;
    switch (ch) {
      case 'A': case 'E': case 'I': case 'O': case 'U': if (i === 0) out.push(ch); break;
      case 'B': out.push('B'); break;
      case 'C': if (next === 'I' && next2 === 'A') out.push('X'); else if (next === 'H') { out.push('X'); i++; } else if ('IEY'.includes(next)) out.push('S'); else out.push('K'); break;
      case 'D': if (next === 'G' && 'EYI'.includes(next2)) { out.push('J'); i++; } else out.push('T'); break;
      case 'G': if (next === 'H' && !/[AEIOU]/.test(next2) && next2) { i++; } else if (next === 'N') { /* silent */ } else if ('IEY'.includes(next)) out.push('J'); else out.push('K'); break;
      case 'H': if (/[AEIOU]/.test(next) && !/[CSPTG]/.test(prev)) out.push('H'); break;
      case 'K': if (prev !== 'C') out.push('K'); break;
      case 'P': if (next === 'H') { out.push('F'); i++; } else out.push('P'); break;
      case 'Q': out.push('K'); break;
      case 'S': if (next === 'H') { out.push('X'); i++; } else if (next === 'I' && (next2 === 'O' || next2 === 'A')) out.push('X'); else out.push('S'); break;
      case 'T': if (next === 'I' && (next2 === 'O' || next2 === 'A')) out.push('X'); else if (next === 'H') { out.push('0'); i++; } else out.push('T'); break;
      case 'V': out.push('F'); break;
      case 'W': case 'Y': if (/[AEIOU]/.test(next)) out.push(ch); break;
      case 'X': out.push('KS'); break;
      case 'Z': out.push('S'); break;
      default: out.push(ch);
    }
  }
  return out.join('').slice(0, 6);
}

export interface GedcomDate { kind: 'exact' | 'about' | 'before' | 'after' | 'between' | 'estimated' | 'calculated' | 'from' | 'to' | 'fromTo' | 'unknown'; y?: number; m?: number; d?: number; y2?: number; m2?: number; d2?: number; text: string; dual?: number }

const MON = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];

/** Parses GEDCOM-style dates: "ABT 1850", "BEF 12 MAR 1799", "BET 1800 AND 1810", "FROM 1900 TO 1905", "1 FEB 1699/00". */
export function parseGedcomDate(text: string): GedcomDate {
  const t = text.trim().toUpperCase().replace(/\s+/g, ' ').replace(/\./g, '');
  if (!t) return { kind: 'unknown', text };
  const part = (s: string): { y?: number; m?: number; d?: number; dual?: number } | null => {
    const str = s.trim();
    const m1 = /^(?:(\d{1,2}) )?(?:([A-Z]{3,}) )?(\d{3,4})(?:\/(\d{1,2}))?$/.exec(str);
    if (!m1) {
      let mm: RegExpExecArray | null;
      if ((mm = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(str))) return { y: +mm[1], m: +mm[2], d: +mm[3] };
      if ((mm = /^([A-Z]{3,})\.? (\d{1,2}),? (\d{3,4})$/.exec(str)) && MON.includes(mm[1].slice(0, 3))) return { y: +mm[3], m: MON.indexOf(mm[1].slice(0, 3)) + 1, d: +mm[2] };
      if ((mm = /^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/.exec(str))) return { y: +mm[3], m: +mm[1], d: +mm[2] };
      if ((mm = /^(\d{1,2})[\/.-](\d{4})$/.exec(str))) return { y: +mm[2], m: +mm[1] };
      const serial = parseTextDate(str); if (serial != null) { const p = serialToParts(serial); return { y: p.y, m: p.m, d: p.d }; }
      return null;
    }
    const mi = m1[2] ? MON.indexOf(m1[2].slice(0, 3)) : -1;
    if (m1[2] && mi < 0) return null;
    const out: { y?: number; m?: number; d?: number; dual?: number } = { y: +m1[3] };
    if (mi >= 0) out.m = mi + 1; if (m1[1]) out.d = +m1[1];
    if (m1[4]) { const base = Math.floor(out.y! / 100) * 100; out.dual = base + +m1[4]; if (out.dual <= out.y!) out.dual += 100; if (out.dual - out.y! > 1) out.dual = out.y! + 1; }
    return out;
  };
  let m: RegExpExecArray | null;
  if ((m = /^BET (.+) AND (.+)$/.exec(t))) { const a = part(m[1]), b = part(m[2]); return a && b ? { kind: 'between', ...a, y2: b.y, m2: b.m, d2: b.d, text } : { kind: 'unknown', text }; }
  if ((m = /^FROM (.+) TO (.+)$/.exec(t))) { const a = part(m[1]), b = part(m[2]); return a && b ? { kind: 'fromTo', ...a, y2: b.y, m2: b.m, d2: b.d, text } : { kind: 'unknown', text }; }
  const mods: [RegExp, GedcomDate['kind']][] = [[/^(ABT|ABOUT|CIRCA|CA|C) (.+)$/, 'about'], [/^(BEF|BEFORE) (.+)$/, 'before'], [/^(AFT|AFTER) (.+)$/, 'after'], [/^(EST|ESTIMATED) (.+)$/, 'estimated'], [/^(CAL|CALCULATED) (.+)$/, 'calculated'], [/^FROM (.+)$/, 'from'], [/^TO (.+)$/, 'to']];
  for (const [re, kind] of mods) if ((m = re.exec(t))) { const p = part(m[m.length - 1]); return p ? { kind, ...p, text } : { kind: 'unknown', text }; }
  const p = part(t);
  return p ? { kind: 'exact', ...p, text } : { kind: 'unknown', text };
}

export interface Ymd { y: number; m: number; d: number }
/** Earliest and latest possible calendar dates of a GEDCOM date (for sorting and age ranges). */
export function gedcomRange(g: GedcomDate): [Ymd, Ymd] | null {
  if (g.y == null) return null;
  const lo = (y: number, m?: number, d?: number): Ymd => ({ y, m: m || 1, d: d || 1 });
  const hi = (y: number, m?: number, d?: number): Ymd => ({ y, m: m || 12, d: d || daysInMonth(y, m || 12) });
  const y = g.dual ?? g.y;
  switch (g.kind) {
    case 'exact': case 'calculated': return [lo(y, g.m, g.d), hi(y, g.m, g.d)];
    case 'about': case 'estimated': return [lo(y - 5, g.m, g.d), hi(y + 5, g.m, g.d)];
    case 'before': case 'to': return [lo(y - 20), addDays(lo(y, g.m, g.d), -1)];
    case 'after': case 'from': return [addDays(hi(y, g.m, g.d), 1), hi(y + 20)];
    case 'between': case 'fromTo': return g.y2 != null ? [lo(y, g.m, g.d), hi(g.y2, g.m2, g.d2)] : [lo(y), hi(y)];
    default: return null;
  }
}

/** Julian day number of a Gregorian calendar date (works for any year, unlike spreadsheet serials). */
export function jdn(p: Ymd): number { const a = Math.floor((14 - p.m) / 12), yy = p.y + 4800 - a, mm = p.m + 12 * a - 3; return p.d + Math.floor((153 * mm + 2) / 5) + 365 * yy + Math.floor(yy / 4) - Math.floor(yy / 100) + Math.floor(yy / 400) - 32045; }
export function addDays(p: Ymd, n: number): Ymd { return jdnToGregorian(jdn(p) + n); }
export function ymdText(p: Ymd): string { return `${String(p.y).padStart(4, '0')}-${String(p.m).padStart(2, '0')}-${String(p.d).padStart(2, '0')}`; }
/** A date as a cell: a serial with a date format from 1900 on, otherwise ISO text (spreadsheets have no serials before 1900). */
export function ymdCell(p: Ymd, date1904 = false): { v: Scalar; style?: Partial<CellStyle> } { return p.y >= 1900 ? { v: partsToSerial(p.y, p.m, p.d, date1904), style: { numFmt: 'yyyy-mm-dd' } } : { v: ymdText(p) }; }

/** Age in completed years (and months, days) between two calendar dates. */
export function ageBetweenYmd(a: Ymd, b: Ymd): { years: number; months: number; days: number } {
  let years = b.y - a.y, months = b.m - a.m, days = b.d - a.d;
  if (days < 0) { months--; const pm = b.m === 1 ? 12 : b.m - 1; days += daysInMonth(pm === 12 ? b.y - 1 : b.y, pm); }
  if (months < 0) { years--; months += 12; }
  return { years, months, days };
}
export function ageBetween(birth: number, at: number, date1904 = false): { years: number; months: number; days: number } {
  const a = serialToParts(Math.floor(birth), date1904), b = serialToParts(Math.floor(at), date1904);
  return ageBetweenYmd({ y: a.y, m: a.m, d: a.d }, { y: b.y, m: b.m, d: b.d });
}

/** Julian calendar date → Gregorian (for records before the 1752 British / 1582 Catholic change). */
export function julianToGregorian(y: number, m: number, d: number): { y: number; m: number; d: number } {
  const a = Math.floor((14 - m) / 12), yy = y + 4800 - a, mm = m + 12 * a - 3;
  const jdn = d + Math.floor((153 * mm + 2) / 5) + 365 * yy + Math.floor(yy / 4) - 32083;
  return jdnToGregorian(jdn);
}
export function gregorianToJulian(y: number, m: number, d: number): { y: number; m: number; d: number } {
  const a = Math.floor((14 - m) / 12), yy = y + 4800 - a, mm = m + 12 * a - 3;
  const jdn = d + Math.floor((153 * mm + 2) / 5) + 365 * yy + Math.floor(yy / 4) - Math.floor(yy / 100) + Math.floor(yy / 400) - 32045;
  const c = jdn + 32082, dd = Math.floor((4 * c + 3) / 1461), e = c - Math.floor(1461 * dd / 4), mo = Math.floor((5 * e + 2) / 153);
  return { d: e - Math.floor((153 * mo + 2) / 5) + 1, m: mo + 3 - 12 * Math.floor(mo / 10), y: dd - 4800 + Math.floor(mo / 10) };
}
function jdnToGregorian(jdn: number): { y: number; m: number; d: number } {
  const a = jdn + 32044, b = Math.floor((4 * a + 3) / 146097), c = a - Math.floor(146097 * b / 4), d = Math.floor((4 * c + 3) / 1461), e = c - Math.floor(1461 * d / 4), m = Math.floor((5 * e + 2) / 153);
  return { d: e - Math.floor((153 * m + 2) / 5) + 1, m: m + 3 - 12 * Math.floor(m / 10), y: 100 * b + d - 4800 + Math.floor(m / 10) };
}

/** Quaker (Friends) dates: before 1752 the first month was March ("1st mo" = March). */
export function quakerMonth(month: number, year: number): number { return year < 1752 ? ((month + 1) % 12) + 1 : month; }

/** "Smith, John" style key: last name first, for sorting an index of names. */
export function surnameFirst(name: string): string { const t = name.trim(); if (t.includes(',')) return t; const words = t.split(/\s+/); if (words.length < 2) return t; const particles = ['van', 'von', 'de', 'der', 'den', 'del', 'di', 'da', 'du', 'la', 'le', 'mac', 'mc', 'st.', 'ter', 'ten']; let i = words.length - 1; while (i > 1 && particles.includes(words[i - 1].toLowerCase())) i--; return `${words.slice(i).join(' ')}, ${words.slice(0, i).join(' ')}`; }

/** A cell's date as calendar parts: a formatted serial, a plausible serial, GEDCOM-style text or any text date (any year). */
function dateOf(app: App, v: Scalar, r: number, c: number): Ymd | null {
  const d1904 = app.book.calc.date1904;
  if (typeof v === 'number' && (isDateFormat(app.sheet.styleOf(r, c).numFmt) || (v > 20000 && v < 80000))) { const p = serialToParts(v, d1904); return { y: p.y, m: p.m, d: p.d }; }
  if (typeof v === 'string') { const g = parseGedcomDate(v); const rg = gedcomRange(g); return rg ? rg[0] : null; }
  return null;
}
/** The serial of a calendar date, or null before 1900. */
function serialOf(p: Ymd, date1904 = false): number | null { return p.y >= 1900 ? partsToSerial(p.y, p.m, p.d, date1904) : null; }
function todayYmd(): Ymd { const d = new Date(); return { y: d.getFullYear(), m: d.getMonth() + 1, d: d.getDate() }; }

defTools([
  { id: 'extra.soundex', label: 'Soundex Codes Beside Names', icon: 'spellcheck', tab: X, group: G_GEN, edits: true, popular: true, tip: 'The American Soundex code (R163) of each selected name in the column to the right, so spelling variants of a surname sort and match together', keywords: 'soundex surname variants genealogy', run: (app) => fillBeside(app, 'Soundex', (t) => soundex(t)) },
  { id: 'extra.metaphone', label: 'Phonetic Keys Beside Names', icon: 'record_voice_over', tab: X, group: G_GEN, edits: true, tip: 'A Metaphone-style sound key that catches more spelling variants than Soundex (Smith/Smyth/Schmidt)', run: (app) => fillBeside(app, 'Phonetic Key', (t) => metaphoneKey(t)) },
  {
    id: 'extra.gedcomDates', label: 'Parse Genealogy Dates (ABT, BEF, BET…)…', icon: 'event', tab: X, group: G_GEN, edits: true, popular: true, tip: '"ABT 1850", "BEF 12 MAR 1799", "BET 1800 AND 1810", "1 FEB 1699/00": get the earliest and latest possible date, the qualifier and the year as columns for sorting and filtering', keywords: 'gedcom dates about before between dual dating',
    run: (app) => { const s = app.sheet; const rg = clipToUsed(s, activeRange(app)); const header = headerGuess(s, rg); const target = { r1: rg.r1, r2: rg.r2, c1: rg.c2 + 1, c2: rg.c2 + 4 }; if (!checkWritable(app, [target])) return; let n = 0; app.book.transact('Parse Genealogy Dates', () => { for (let r = rg.r1; r <= rg.r2; r++) { if (r === rg.r1 && header) { ['Qualifier', 'Earliest', 'Latest', 'Year'].forEach((t, j) => writeOut(s, r, target.c1 + j, { v: t, style: HEAD })); continue; } const t = shownText(s, r, rg.c1); if (!t) continue; const g = parseGedcomDate(t); const range = gedcomRange(g); writeOut(s, r, target.c1, g.kind); if (range) { writeOut(s, r, target.c1 + 1, ymdCell(range[0], app.book.calc.date1904)); writeOut(s, r, target.c1 + 2, ymdCell(range[1], app.book.calc.date1904)); } if (g.y != null) writeOut(s, r, target.c1 + 3, g.dual ?? g.y); n++; } }); toast(`${plural(n, 'date')} parsed`); },
  },
  {
    id: 'extra.ageAtDate', label: 'Age Between Two Date Columns', icon: 'cake', tab: X, group: G_GEN, edits: true, popular: true, tip: 'Selected column = birth date, next column = death date or event date (empty = today): the age in years, and in years/months/days, in the columns to the right', keywords: 'age birth death years months days',
    run: (app) => { const s = app.sheet; const rg = clipToUsed(s, activeRange(app)); const header = headerGuess(s, rg); const target = { r1: rg.r1, r2: rg.r2, c1: rg.c2 + 1, c2: rg.c2 + 2 }; if (!checkWritable(app, [target])) return; const today = todayYmd(); let n = 0; app.book.transact('Age', () => { for (let r = rg.r1; r <= rg.r2; r++) { if (r === rg.r1 && header) { writeOut(s, r, target.c1, { v: 'Age', style: HEAD }); writeOut(s, r, target.c1 + 1, { v: 'Age (y m d)', style: HEAD }); continue; } const b = dateOf(app, s.value(r, rg.c1), r, rg.c1); if (b == null) continue; const e = rg.c2 > rg.c1 ? dateOf(app, s.value(r, rg.c1 + 1), r, rg.c1 + 1) ?? today : today; const a = ageBetweenYmd(b, e); writeOut(s, r, target.c1, a.years); writeOut(s, r, target.c1 + 1, `${a.years}y ${a.months}m ${a.days}d`); n++; } }); toast(`${plural(n, 'age')} computed`); },
  },
  {
    id: 'extra.julianGregorian', label: 'Convert Julian ⇄ Gregorian Dates', icon: 'swap_calls', tab: X, group: G_GEN, edits: true, kind: 'split', tip: 'Old Style (Julian) dates in records before 1752 (Britain and colonies) or 1582 become New Style, or the reverse',
    run: (app, arg?: 'toGregorian' | 'toJulian') => { const d1904 = app.book.calc.date1904; changeCells(app, 'Julian/Gregorian', (cell, r, c) => { const p = dateOf(app, cell.v ?? null, r, c); if (p == null) return undefined; const o = arg === 'toJulian' ? gregorianToJulian(p.y, p.m, p.d) : julianToGregorian(p.y, p.m, p.d); const out = ymdCell(o, d1904); return { ...cell, v: out.v, ...(out.style ? { s: app.book.derivedStyle(cell.s ?? 0, out.style) } : {}) }; }); },
    menu: () => [{ id: 'extra.julianGregorian', arg: 'toGregorian', label: 'Julian (Old Style) → Gregorian' }, { id: 'extra.julianGregorian', arg: 'toJulian', label: 'Gregorian → Julian' }],
  },
  { id: 'extra.quakerDates', label: 'Quaker Month Numbers to Month Names', icon: 'church', tab: X, group: G_GEN, edits: true, tip: '"12th 3mo 1720" style dates: the month number counted from March before 1752 becomes the calendar month', run: (app) => transformValues(app, 'Quaker Dates', (v) => { const m = /^(\d{1,2})(?:st|nd|rd|th)?\s+(?:(?:of\s+)?)(\d{1,2})\s*mo\.?\s+(\d{4})$/i.exec(String(v).trim()); if (!m) return undefined; const y = +m[3]; const mo = quakerMonth(+m[2], y); return `${+m[1]} ${MONTHS[mo - 1]} ${y}${y < 1752 && +m[2] >= 11 ? ' (dual-year possible)' : ''}`; }) },
  { id: 'extra.surnameFirst', label: 'Names to "Surname, Given" Order', icon: 'badge', tab: X, group: G_GEN, edits: true, tip: '"John van der Berg" becomes "van der Berg, John" (particles kept with the surname) for an index', run: (app) => transformValues(app, 'Surname First', (v) => surnameFirst(String(v))) },
  {
    id: 'extra.familyIndex', label: 'Build an Alphabetical Name Index', icon: 'sort_by_alpha', tab: X, group: G_GEN, edits: true, tip: 'A new sheet listing every distinct name of the selected column, surname first, with a letter heading per initial and the cells where each appears (a book index)',
    run: (app) => { const s = app.sheet; const rg = clipToUsed(s, activeRange(app)); const m = new Map<string, string[]>(); for (let r = rg.r1; r <= rg.r2; r++) { const t = shownText(s, r, rg.c1).trim(); if (!t) continue; const k = surnameFirst(t); (m.get(k) || m.set(k, []).get(k)!).push(String(r + 1)); } const keys = [...m.keys()].sort(naturalCompare); const rows: any[][] = []; let letter = ''; for (const k of keys) { const l = k[0].toUpperCase(); if (l !== letter) { letter = l; rows.push([{ v: letter, style: { bold: true, size: 14 } }, '']); } rows.push([k, m.get(k)!.join(', ')]); } if (!rows.length) { toast('No names in the selection.'); return; } addSheetWithRows(app, 'Index', [[{ v: 'Name', style: HEAD }, { v: 'Rows', style: HEAD }], ...rows], { colW: [240, 300], freeze: 1 }); },
  },
  { id: 'extra.menuGenealogy', label: 'Genealogy', icon: 'family_restroom', tab: X, group: G_GEN, kind: 'menu', noRecord: true, run: () => undefined, menu: () => menuOf(['extra.soundex', 'extra.metaphone', 'extra.gedcomDates', 'extra.ageAtDate', 'extra.julianGregorian', 'extra.quakerDates', 'extra.surnameFirst', 'extra.familyIndex', 'plus.splitNames', 'util.sortByLastWord', 'tools.fuzzyMatch', 'tools.ocrCleanup', 'tools.numberToWords']) },
]);

/** Writes fn(text) beside each selected cell. */
function fillBeside(app: App, label: string, fn: (t: string, v: Scalar) => Scalar | undefined, cols = 1): void {
  const s = app.sheet; const rg = clipToUsed(s, activeRange(app));
  const target = { r1: rg.r1, r2: rg.r2, c1: rg.c2 + 1, c2: rg.c2 + cols };
  if (!checkWritable(app, [target])) return;
  let n = 0;
  app.book.transact(label, () => { for (let r = rg.r1; r <= rg.r2; r++) { const t = shownText(s, r, rg.c1); if (!t) continue; const out = fn(t, s.value(r, rg.c1)); if (out === undefined) continue; writeOut(s, r, target.c1, out); n++; } });
  toast(`${label}: ${plural(n, 'cell')} written`);
}

// ============================================================================================
// Publishing

const SMALL_WORDS = new Set(['a', 'an', 'the', 'and', 'but', 'or', 'nor', 'for', 'so', 'yet', 'as', 'at', 'by', 'in', 'of', 'on', 'per', 'to', 'up', 'via', 'vs', 'v', 'from', 'into', 'with', 'off', 'out', 'over']);

/** Title case the way style guides want it: small words lower except first and last, after a colon capitalised, acronyms kept. */
export function titleCase(s: string): string {
  const words = s.trim().split(/(\s+|-)/);
  // Text typed all in capitals has no acronyms to protect: every word is recased.
  const allCaps = s === s.toUpperCase();
  let first = true; let afterColon = false;
  const out = words.map((w, i) => {
    if (/^\s+$/.test(w) || w === '-' || w === '') return w;
    const isLast = words.slice(i + 1).every((x) => /^\s*$/.test(x) || x === '-');
    const low = w.toLowerCase(); const bare = low.replace(/[^\w']/g, '');
    let res: string;
    if (!allCaps && w === w.toUpperCase() && /[A-Z]{2,}/.test(w) && w.length <= 5) res = w; // acronym
    else if (!first && !isLast && !afterColon && SMALL_WORDS.has(bare)) res = low;
    else res = w.charAt(0).toUpperCase() + w.slice(1).toLowerCase();
    first = false; afterColon = /:$/.test(w);
    return res;
  });
  return out.join('');
}

export function smartPunctuation(s: string): string {
  return s.replace(/(^|[\s(\["])"/g, '$1“').replace(/"/g, '”').replace(/(^|[\s(\["])'/g, '$1‘').replace(/'/g, '’').replace(/\s*---\s*/g, '—').replace(/(\w)--(\w)/g, '$1—$2').replace(/\s--\s/g, ' – ').replace(/(\d)\s*-\s*(\d)/g, '$1–$2').replace(/\.\.\./g, '…');
}
export function straightPunctuation(s: string): string { return s.replace(/[“”„]/g, '"').replace(/[‘’‚]/g, "'").replace(/—/g, '--').replace(/–/g, '-').replace(/…/g, '...'); }
const FRACTIONS: Record<string, string> = { '1/2': '½', '1/3': '⅓', '2/3': '⅔', '1/4': '¼', '3/4': '¾', '1/5': '⅕', '2/5': '⅖', '3/5': '⅗', '4/5': '⅘', '1/6': '⅙', '5/6': '⅚', '1/8': '⅛', '3/8': '⅜', '5/8': '⅝', '7/8': '⅞' };
export function fractionGlyphs(s: string): string { return s.replace(/(?<![\d/.])(\d+)\/(\d+)(?![\d/])/g, (m) => FRACTIONS[m] || m); }
export function thinSpaceThousands(s: string): string { return s.replace(/\b(\d{1,3})(,\d{3})+\b/g, (m) => m.replace(/,/g, ' ')); }
export function hangingPunctuationSafe(s: string): string { return s.replace(/\s+([,.;:!?])/g, '$1').replace(/([,.;:!?])(?=[^\s"'”’)\]])/g, '$1 ').replace(/\.\s\.\s\./g, '...'); }

/** Flesch reading ease, roughly, for one cell of prose. */
export function readingEase(text: string): number | null {
  const sentences = Math.max(1, (text.match(/[.!?]+(\s|$)/g) || []).length);
  const words = text.trim().split(/\s+/).filter(Boolean);
  if (words.length < 3) return null;
  const syllables = words.reduce((a, w) => a + Math.max(1, (w.toLowerCase().replace(/e$/, '').match(/[aeiouy]+/g) || []).length), 0);
  return Math.round((206.835 - 1.015 * (words.length / sentences) - 84.6 * (syllables / words.length)) * 10) / 10;
}

export function isbn10Check(d: string): string { let sum = 0; for (let i = 0; i < 9; i++) sum += (10 - i) * +d[i]; const c = (11 - (sum % 11)) % 11; return c === 10 ? 'X' : String(c); }
export function isbn13Check(d: string): string { let sum = 0; for (let i = 0; i < 12; i++) sum += (+d[i]) * (i % 2 ? 3 : 1); return String((10 - (sum % 10)) % 10); }
/** Validates ISBN-10/13 and converts between them; returns { valid, isbn10, isbn13 }. */
export function isbnInfo(raw: string): { valid: boolean; isbn10?: string; isbn13?: string; kind?: string } {
  const d = raw.toUpperCase().replace(/[^0-9X]/g, '');
  if (d.length === 10) { const ok = /^\d{9}[\dX]$/.test(d) && isbn10Check(d) === d[9]; const core = '978' + d.slice(0, 9); return { valid: ok, kind: 'ISBN-10', isbn10: d, isbn13: core + isbn13Check(core) }; }
  if (d.length === 13 && /^\d{13}$/.test(d)) { const ok = isbn13Check(d) === d[12]; const out: { valid: boolean; isbn10?: string; isbn13?: string; kind?: string } = { valid: ok, kind: 'ISBN-13', isbn13: d }; if (d.startsWith('978')) { const core = d.slice(3, 12); out.isbn10 = core + isbn10Check(core); } return out; }
  return { valid: false };
}
export function luhnValid(s: string): boolean { const d = s.replace(/\D/g, ''); if (d.length < 2) return false; let sum = 0; for (let i = 0; i < d.length; i++) { let n = +d[d.length - 1 - i]; if (i % 2) { n *= 2; if (n > 9) n -= 9; } sum += n; } return sum % 10 === 0; }
export function slugify(s: string): string { return s.normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, ''); }
export function ptToUnit(pt: number, unit: 'pica' | 'mm' | 'in' | 'px'): number { return unit === 'pica' ? pt / 12 : unit === 'mm' ? pt * 25.4 / 72 : unit === 'in' ? pt / 72 : pt * 96 / 72; }

defTools([
  { id: 'extra.titleCase', label: 'Title Case (Style-Guide Rules)', icon: 'title', tab: X, group: G_PUB, edits: true, popular: true, tip: 'Capitalize headlines the way editors do: small words (a, of, the, and) stay lower except first, last and after a colon; acronyms are kept', keywords: 'headline title case chicago ap', run: (app) => transformValues(app, 'Title Case', (v) => titleCase(String(v))) },
  { id: 'extra.smartPunctuation', label: 'Smart Quotes and Dashes', icon: 'format_quote', tab: X, group: G_PUB, edits: true, tip: 'Straight quotes become curly, -- becomes an em dash, number ranges get en dashes, ... becomes an ellipsis', run: (app) => transformValues(app, 'Smart Punctuation', (v) => smartPunctuation(String(v))) },
  { id: 'extra.straightPunctuation', label: 'Straight Quotes and Dashes', icon: 'format_quote', tab: X, group: G_PUB, edits: true, tip: 'The reverse: typographic quotes and dashes back to plain ASCII (for code, CSV and databases)', run: (app) => transformValues(app, 'Straight Punctuation', (v) => straightPunctuation(String(v))) },
  { id: 'extra.fractionGlyphs', label: 'Fraction Characters (1/2 → ½)', icon: 'pie_chart', tab: X, group: G_PUB, edits: true, run: (app) => transformValues(app, 'Fraction Glyphs', (v) => fractionGlyphs(String(v))) },
  { id: 'extra.thinSpaces', label: 'Thin Spaces as Thousands Separators', icon: 'space_bar', tab: X, group: G_PUB, edits: true, tip: '1,234,567 in text becomes 1 234 567 with thin spaces, as many style guides prefer', run: (app) => transformValues(app, 'Thin Spaces', (v) => thinSpaceThousands(String(v))) },
  { id: 'extra.fixSpacing', label: 'Fix Spacing Around Punctuation', icon: 'format_align_left', tab: X, group: G_PUB, edits: true, tip: 'No space before commas and periods, one space after, double spaces collapsed', run: (app) => transformValues(app, 'Fix Spacing', (v) => hangingPunctuationSafe(String(v).replace(/ {2,}/g, ' ')))  },
  { id: 'extra.readingEase', label: 'Reading Ease Score Beside Text', icon: 'menu_book', tab: X, group: G_PUB, edits: true, tip: 'A Flesch reading-ease score (higher = easier) for each selected cell of prose', run: (app) => fillBeside(app, 'Reading Ease', (t) => readingEase(t) ?? undefined) },
  {
    id: 'extra.isbn', label: 'Check and Convert ISBNs', icon: 'book', tab: X, group: G_PUB, edits: true, popular: true, tip: 'Beside each ISBN: whether its check digit is valid, and the ISBN-10 and ISBN-13 forms', keywords: 'isbn check digit convert 10 13 publishing',
    run: (app) => { const s = app.sheet; const rg = clipToUsed(s, activeRange(app)); const header = headerGuess(s, rg); const target = { r1: rg.r1, r2: rg.r2, c1: rg.c2 + 1, c2: rg.c2 + 3 }; if (!checkWritable(app, [target])) return; app.book.transact('ISBN', () => { for (let r = rg.r1; r <= rg.r2; r++) { if (r === rg.r1 && header) { ['Valid', 'ISBN-10', 'ISBN-13'].forEach((t, j) => writeOut(s, r, target.c1 + j, { v: t, style: HEAD })); continue; } const t = shownText(s, r, rg.c1); if (!t) continue; const info = isbnInfo(t); writeOut(s, r, target.c1, info.kind ? (info.valid ? 'Valid' : 'Bad check digit') : 'Not an ISBN'); if (info.isbn10) writeOut(s, r, target.c1 + 1, { v: info.isbn10, style: { numFmt: '@' } }); if (info.isbn13) writeOut(s, r, target.c1 + 2, { v: info.isbn13, style: { numFmt: '@' } }); } }); },
  },
  { id: 'extra.slugify', label: 'Make URL Slugs', icon: 'link', tab: X, group: G_PUB, edits: true, tip: '"Crème Brûlée & Co." becomes creme-brulee-co, for file names and web addresses', run: (app) => fillBeside(app, 'Slug', (t) => slugify(t)) },
  { id: 'extra.pointsConvert', label: 'Points ⇄ Picas, Millimeters, Inches, Pixels', icon: 'straighten', tab: X, group: G_PUB, edits: true, kind: 'split', tip: 'Typographic measurement conversion in place', run: (app, arg?: 'pica' | 'mm' | 'in' | 'px' | 'fromPica' | 'fromMm' | 'fromIn') => { const to = arg || 'mm'; changeCells(app, 'Convert Points', (cell) => { if (typeof cell.v !== 'number') return undefined; const v = cell.v; const out = to === 'fromPica' ? v * 12 : to === 'fromMm' ? v * 72 / 25.4 : to === 'fromIn' ? v * 72 : ptToUnit(v, to); return { ...cell, v: Math.round(out * 1000) / 1000 }; }); }, menu: () => [{ id: 'extra.pointsConvert', arg: 'pica', label: 'Points → picas' }, { id: 'extra.pointsConvert', arg: 'mm', label: 'Points → millimeters' }, { id: 'extra.pointsConvert', arg: 'in', label: 'Points → inches' }, { id: 'extra.pointsConvert', arg: 'px', label: 'Points → pixels (96 dpi)' }, { separator: true }, { id: 'extra.pointsConvert', arg: 'fromPica', label: 'Picas → points' }, { id: 'extra.pointsConvert', arg: 'fromMm', label: 'Millimeters → points' }, { id: 'extra.pointsConvert', arg: 'fromIn', label: 'Inches → points' }] },
  { id: 'extra.menuPublishing', label: 'Publishing', icon: 'auto_stories', tab: X, group: G_PUB, kind: 'menu', noRecord: true, run: () => undefined, menu: () => menuOf(['extra.titleCase', 'extra.smartPunctuation', 'extra.straightPunctuation', 'extra.fractionGlyphs', 'extra.thinSpaces', 'extra.fixSpacing', 'extra.readingEase', 'extra.isbn', 'extra.slugify', 'extra.pointsConvert', 'util.exportSelectionLatex', 'util.copyAsMarkdown', 'tools.ocrCleanup', 'plus.superscriptOrdinals']) },
]);

// ============================================================================================
// Finance

export function amortization(principal: number, annualRate: number, months: number, startSerial: number, date1904 = false): Scalar[][] {
  const r = annualRate / 12;
  const pay = r === 0 ? principal / months : principal * r / (1 - Math.pow(1 + r, -months));
  let bal = principal; const rows: Scalar[][] = [];
  const p0 = serialToParts(startSerial, date1904);
  for (let i = 1; i <= months; i++) {
    const interest = bal * r; const princ = Math.min(bal, pay - interest); bal = Math.max(0, bal - princ);
    const t = (p0.y * 12 + p0.m - 1) + i; const y = Math.floor(t / 12), m = t - y * 12 + 1;
    rows.push([i, partsToSerial(y, m, Math.min(p0.d, daysInMonth(y, m)), date1904), Math.round(pay * 100) / 100, Math.round(princ * 100) / 100, Math.round(interest * 100) / 100, Math.round(bal * 100) / 100]);
  }
  return rows;
}

export function agingBucket(days: number, edges = [0, 30, 60, 90]): string { if (days <= edges[0]) return 'Current'; for (let i = 1; i < edges.length; i++) if (days <= edges[i]) return `${edges[i - 1] + 1}–${edges[i]}`; return `Over ${edges[edges.length - 1]}`; }

/** Benford's law: expected and observed first-digit frequencies. */
export function benford(values: number[]): { digit: number; expected: number; observed: number; count: number }[] {
  const counts = new Array(10).fill(0); let total = 0;
  for (const v of values) { const d = Number.isFinite(v) && v !== 0 ? Math.abs(v).toExponential()[0] : ''; if (d && d !== '0') { counts[+d]++; total++; } }
  const out: { digit: number; expected: number; observed: number; count: number }[] = [];
  for (let d = 1; d <= 9; d++) out.push({ digit: d, expected: Math.log10(1 + 1 / d), observed: total ? counts[d] / total : 0, count: counts[d] });
  return out;
}

export function roundToIncrement(v: number, inc: number): number { return inc ? Math.round(v / inc) * inc : v; }
export function fiscalYearYmd(p: { y: number; m: number }, startMonth: number): { fy: number; quarter: number } { const shifted = (p.m - startMonth + 12) % 12; return { fy: startMonth === 1 ? p.y : p.y + (p.m >= startMonth ? 1 : 0), quarter: Math.floor(shifted / 3) + 1 }; }
export function fiscalYear(serial: number, startMonth: number, date1904 = false): { fy: number; quarter: number } { return fiscalYearYmd(serialToParts(serial, date1904), startMonth); }

defTools([
  {
    id: 'extra.runningBalance', label: 'Running Balance Column', icon: 'account_balance', tab: X, group: G_FIN, edits: true, popular: true, tip: 'Beside a column of amounts (or Debit and Credit columns), a running balance with formulas, starting from an opening balance', keywords: 'running total balance register ledger',
    run: async (app, arg?: { opening: number; mode: 'single' | 'debitCredit' }) => {
      const s = app.sheet; const rg = clipToUsed(s, activeRange(app)); const header = headerGuess(s, rg);
      const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Running Balance', fields: [{ key: 'opening', label: 'Opening balance', type: 'number', value: 0, step: 0.01 }, { key: 'mode', label: 'Amounts', type: 'radio', value: rg.c2 > rg.c1 ? 'debitCredit' : 'single', options: [{ value: 'single', label: 'One signed column (deposits +, payments −)' }, { value: 'debitCredit', label: 'Two columns: Debit (−) then Credit (+)' }] }] }); return r ? { opening: Number(r.opening) || 0, mode: r.mode } : null; });
      if (!o) return;
      const target = { r1: rg.r1, r2: rg.r2, c1: rg.c2 + 1, c2: rg.c2 + 1 }; if (!checkWritable(app, [target])) return;
      const first = rg.r1 + (header ? 1 : 0); const c = target.c1;
      app.book.transact('Running Balance', () => { if (header) writeOut(s, rg.r1, c, { v: 'Balance', style: HEAD }); for (let r = first; r <= rg.r2; r++) { const prev = r === first ? String(o.opening) : cellName(r - 1, c); const f = o.mode === 'single' ? `${prev}+${cellName(r, rg.c1)}` : `${prev}-N(${cellName(r, rg.c1)})+N(${cellName(r, rg.c1 + 1)})`; writeOut(s, r, c, { f, style: { numFmt: '#,##0.00;[Red]-#,##0.00' } }); } });
    },
  },
  {
    id: 'extra.amortization', label: 'Loan Amortization Schedule…', icon: 'payments', tab: X, group: G_FIN, edits: true, popular: true, tip: 'A payment-by-payment schedule (date, payment, principal, interest, balance) for a loan on a new sheet',
    run: async (app, arg?: { principal: number; rate: number; months: number; start?: string }) => { const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Loan Amortization', fields: [{ key: 'principal', label: 'Amount borrowed', type: 'number', value: 10000, step: 0.01, half: true }, { key: 'rate', label: 'Annual interest rate (%)', type: 'number', value: 6, step: 0.01, half: true }, { key: 'months', label: 'Months', type: 'number', value: 36, min: 1, max: 600, half: true }, { key: 'start', label: 'First payment', type: 'date', value: new Date().toISOString().slice(0, 10), half: true }] }); return r ? { principal: Number(r.principal), rate: Number(r.rate) / 100, months: Number(r.months), start: r.start } : null; }); if (!o) return; const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(o.start || ''); const start = m ? partsToSerial(+m[1], +m[2], +m[3], app.book.calc.date1904) : Math.floor(jsDateToSerial(new Date(), app.book.calc.date1904)); const rows = amortization(o.principal, o.rate, o.months, start, app.book.calc.date1904); const fmt = (v: Scalar, j: number) => (j === 1 ? { v, style: { numFmt: 'yyyy-mm-dd' } } : j >= 2 ? { v, style: { numFmt: '#,##0.00' } } : v); report(app, 'Amortization', ['#', 'Date', 'Payment', 'Principal', 'Interest', 'Balance'], rows.map((r) => r.map(fmt) as any), { colW: [50, 100, 100, 100, 100, 110] }); const total = rows.reduce((a, r) => a + Number(r[4]), 0); toast(`Payment ${(rows[0][2] as number).toFixed(2)} per month; total interest ${total.toFixed(2)}`); },
  },
  {
    id: 'extra.agingBuckets', label: 'Aging Buckets from Dates', icon: 'hourglass_bottom', tab: X, group: G_FIN, edits: true, tip: 'Beside each due date: days overdue as of today and the bucket (Current, 1–30, 31–60, 61–90, Over 90)', keywords: 'receivables aging overdue buckets',
    run: (app) => { const s = app.sheet; const rg = clipToUsed(s, activeRange(app)); const header = headerGuess(s, rg); const target = { r1: rg.r1, r2: rg.r2, c1: rg.c2 + 1, c2: rg.c2 + 2 }; if (!checkWritable(app, [target])) return; const today = jdn(todayYmd()); app.book.transact('Aging', () => { for (let r = rg.r1; r <= rg.r2; r++) { if (r === rg.r1 && header) { writeOut(s, r, target.c1, { v: 'Days', style: HEAD }); writeOut(s, r, target.c1 + 1, { v: 'Bucket', style: HEAD }); continue; } const d = dateOf(app, s.value(r, rg.c1), r, rg.c1); if (d == null) continue; const days = today - jdn(d); writeOut(s, r, target.c1, days); writeOut(s, r, target.c1 + 1, agingBucket(days)); } }); },
  },
  { id: 'extra.splitDebitCredit', label: 'Split Signed Amounts into Debit and Credit', icon: 'call_split', tab: X, group: G_FIN, edits: true, kind: 'split', tip: 'A signed amount column becomes two columns (negatives under Debit, positives under Credit), or two columns become one signed column', run: (app, arg?: 'split' | 'join') => { const s = app.sheet; const rg = clipToUsed(s, activeRange(app)); const header = headerGuess(s, rg); if (arg === 'join') { const target = { r1: rg.r1, r2: rg.r2, c1: rg.c2 + 1, c2: rg.c2 + 1 }; if (!checkWritable(app, [target])) return; app.book.transact('Join Debit/Credit', () => { for (let r = rg.r1; r <= rg.r2; r++) { if (r === rg.r1 && header) { writeOut(s, r, target.c1, { v: 'Amount', style: HEAD }); continue; } const d = numOf(s.value(r, rg.c1)) || 0, c = numOf(s.value(r, rg.c1 + 1)) || 0; if (d || c) writeOut(s, r, target.c1, c - d); } }); return; } const target = { r1: rg.r1, r2: rg.r2, c1: rg.c2 + 1, c2: rg.c2 + 2 }; if (!checkWritable(app, [target])) return; app.book.transact('Split Debit/Credit', () => { for (let r = rg.r1; r <= rg.r2; r++) { if (r === rg.r1 && header) { writeOut(s, r, target.c1, { v: 'Debit', style: HEAD }); writeOut(s, r, target.c1 + 1, { v: 'Credit', style: HEAD }); continue; } const v = numOf(s.value(r, rg.c1)); if (v == null) continue; if (v < 0) writeOut(s, r, target.c1, -v); else writeOut(s, r, target.c1 + 1, v); } }); }, menu: () => [{ id: 'extra.splitDebitCredit', arg: 'split', label: 'Signed amount → Debit and Credit' }, { id: 'extra.splitDebitCredit', arg: 'join', label: 'Debit and Credit → signed amount' }] },
  { id: 'extra.percentChange', label: 'Percent Change Column', icon: 'trending_up', tab: X, group: G_FIN, edits: true, tip: 'Beside a column of numbers, the change from the previous row as a percentage (formulas)', run: (app) => { const s = app.sheet; const rg = clipToUsed(s, activeRange(app)); const header = headerGuess(s, rg); const c = rg.c2 + 1; if (!checkWritable(app, [{ r1: rg.r1, r2: rg.r2, c1: c, c2: c }])) return; app.book.transact('Percent Change', () => { if (header) writeOut(s, rg.r1, c, { v: 'Change %', style: HEAD }); for (let r = rg.r1 + (header ? 2 : 1); r <= rg.r2; r++) writeOut(s, r, c, { f: `IF(${cellName(r - 1, rg.c1)}=0,"",(${cellName(r, rg.c1)}-${cellName(r - 1, rg.c1)})/ABS(${cellName(r - 1, rg.c1)}))`, style: { numFmt: '0.0%;[Red]-0.0%' } }); }); } },
  { id: 'extra.indexTo100', label: 'Index to 100 Column', icon: 'ssid_chart', tab: X, group: G_FIN, edits: true, tip: 'Each value as a percentage of the first value (first = 100), for comparing series', run: (app) => { const s = app.sheet; const rg = clipToUsed(s, activeRange(app)); const header = headerGuess(s, rg); const c = rg.c2 + 1; const first = rg.r1 + (header ? 1 : 0); if (!checkWritable(app, [{ r1: rg.r1, r2: rg.r2, c1: c, c2: c }])) return; app.book.transact('Index to 100', () => { if (header) writeOut(s, rg.r1, c, { v: 'Index', style: HEAD }); for (let r = first; r <= rg.r2; r++) writeOut(s, r, c, { f: `${cellName(r, rg.c1)}/${cellName(first, rg.c1, true, true)}*100`, style: { numFmt: '0.0' } }); }); } },
  { id: 'extra.movingAverage', label: 'Moving Average Column…', icon: 'show_chart', tab: X, group: G_FIN, edits: true, run: async (app, arg?: { n: number }) => { const o = await argOrAsk(arg, async () => { const t = await promptDialog('Moving Average', 'Window (rows)', '3', { type: 'number' }); return t ? { n: Math.max(1, Number(t) || 3) } : null; }); if (!o) return; const s = app.sheet; const rg = clipToUsed(s, activeRange(app)); const header = headerGuess(s, rg); const c = rg.c2 + 1; if (!checkWritable(app, [{ r1: rg.r1, r2: rg.r2, c1: c, c2: c }])) return; const first = rg.r1 + (header ? 1 : 0); app.book.transact('Moving Average', () => { if (header) writeOut(s, rg.r1, c, { v: `MA(${o.n})`, style: HEAD }); for (let r = first + o.n - 1; r <= rg.r2; r++) writeOut(s, r, c, { f: `AVERAGE(${cellName(r - o.n + 1, rg.c1)}:${cellName(r, rg.c1)})`, style: { numFmt: '#,##0.00' } }); }); } },
  { id: 'extra.roundToIncrement', label: 'Round to the Nearest 0.05, 0.25, 5…', icon: 'decimal_decrease', tab: X, group: G_FIN, edits: true, run: async (app, arg?: { inc: number }) => { const o = await argOrAsk(arg, async () => { const t = await promptDialog('Round to Increment', 'Increment (0.05 for cash rounding, 5, 100…)', '0.05', { type: 'number' }); return t ? { inc: Number(t) || 0.05 } : null; }); if (!o) return; changeCells(app, 'Round to Increment', (cell) => (typeof cell.v === 'number' ? { ...cell, v: Math.round(roundToIncrement(cell.v, o.inc) * 1e10) / 1e10 } : undefined)); } },
  { id: 'extra.fiscalYear', label: 'Fiscal Year and Quarter Columns…', icon: 'calendar_today', tab: X, group: G_FIN, edits: true, run: async (app, arg?: { startMonth: number }) => { const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Fiscal Year', fields: [{ key: 'startMonth', label: 'Fiscal year starts in', type: 'select', value: '7', options: MONTHS.map((m, i) => ({ value: String(i + 1), label: m })) }] }); return r ? { startMonth: Number(r.startMonth) } : null; }); if (!o) return; const s = app.sheet; const rg = clipToUsed(s, activeRange(app)); const header = headerGuess(s, rg); const target = { r1: rg.r1, r2: rg.r2, c1: rg.c2 + 1, c2: rg.c2 + 2 }; if (!checkWritable(app, [target])) return; app.book.transact('Fiscal Year', () => { for (let r = rg.r1; r <= rg.r2; r++) { if (r === rg.r1 && header) { writeOut(s, r, target.c1, { v: 'FY', style: HEAD }); writeOut(s, r, target.c1 + 1, { v: 'FQ', style: HEAD }); continue; } const d = dateOf(app, s.value(r, rg.c1), r, rg.c1); if (d == null) continue; const f = fiscalYearYmd(d, o.startMonth); writeOut(s, r, target.c1, f.fy); writeOut(s, r, target.c1 + 1, `Q${f.quarter}`); } }); } },
  {
    id: 'extra.benford', label: "Benford's Law Check", icon: 'bar_chart', tab: X, group: G_FIN, edits: true, popular: true, tip: 'Compare the first-digit distribution of the selected numbers with Benford\'s law (an audit screen for invented figures), on a new sheet', keywords: 'benford audit fraud first digit',
    run: (app) => { const nums = selCells(app, { existingOnly: true }).map((x) => x.cell!.v).filter((v): v is number => typeof v === 'number' && v !== 0); if (nums.length < 20) { toast('Select at least 20 numbers.'); return; } const b = benford(nums); const chi = b.reduce((a, x) => a + Math.pow(x.count - x.expected * nums.length, 2) / (x.expected * nums.length), 0); report(app, 'Benford', ['First digit', 'Expected', 'Observed', 'Count', 'Difference'], b.map((x) => [x.digit, { v: x.expected, style: { numFmt: '0.0%' } }, { v: x.observed, style: { numFmt: '0.0%' } }, x.count, { v: x.observed - x.expected, style: { numFmt: '+0.0%;-0.0%' } }] as any)); toast(`${nums.length} numbers; chi-square ${chi.toFixed(1)} (over 15.5 suggests a poor fit)`); },
  },
  { id: 'extra.luhn', label: 'Validate Card/ID Numbers (Luhn)', icon: 'credit_card', tab: X, group: G_FIN, edits: true, tip: 'Beside each number, whether it passes the Luhn checksum used by payment cards, IMEIs and many ID numbers', run: (app) => fillBeside(app, 'Luhn', (t) => (luhnValid(t) ? 'Valid' : 'Invalid')) },
  { id: 'extra.menuFinance', label: 'Finance', icon: 'account_balance', tab: X, group: G_FIN, kind: 'menu', noRecord: true, run: () => undefined, menu: () => menuOf(['extra.runningBalance', 'extra.amortization', 'extra.agingBuckets', 'extra.splitDebitCredit', 'extra.percentChange', 'extra.indexTo100', 'extra.movingAverage', 'extra.roundToIncrement', 'extra.fiscalYear', 'extra.benford', 'extra.luhn', 'tools.numberToWords', 'util.negativeRed', 'data.whatIf']) },
]);

// ============================================================================================
// Data Quality

/** OpenRefine-style fingerprint: lowercase, punctuation removed, accents stripped, tokens sorted and deduplicated. */
export function fingerprint(s: string): string { return [...new Set(s.normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/[^\p{L}\p{N}\s]/gu, ' ').split(/\s+/).filter(Boolean))].sort().join(' '); }
export function ngramFingerprint(s: string, n = 2): string { const t = s.normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/[^\p{L}\p{N}]/gu, ''); const grams = new Set<string>(); for (let i = 0; i <= t.length - n; i++) grams.add(t.slice(i, i + n)); return [...grams].sort().join(''); }

/** Clusters values whose fingerprints match; each cluster lists its variants with counts, most common first. */
export function clusterValues(values: string[], key: (s: string) => string = fingerprint): { key: string; variants: { text: string; n: number }[] }[] {
  const m = new Map<string, Map<string, number>>();
  for (const v of values) { const t = v.trim(); if (!t) continue; const k = key(t); const inner = m.get(k) || m.set(k, new Map()).get(k)!; inner.set(t, (inner.get(t) || 0) + 1); }
  const caps = (t: string) => (t.match(/\p{Lu}/gu) || []).length;
  return [...m].filter(([, inner]) => inner.size > 1).map(([k, inner]) => ({ key: k, variants: [...inner].map(([text, n]) => ({ text, n })).sort((a, b) => b.n - a.n || caps(b.text) - caps(a.text) || a.text.localeCompare(b.text)) })).sort((a, b) => b.variants.length - a.variants.length);
}

export function iqrOutliers(xs: number[]): { lo: number; hi: number } | null { if (xs.length < 4) return null; const s = [...xs].sort((a, b) => a - b); const q = (p: number) => { const i = (s.length - 1) * p; const f = Math.floor(i); return s[f] + (s[Math.min(f + 1, s.length - 1)] - s[f]) * (i - f); }; const q1 = q(0.25), q3 = q(0.75), iqr = q3 - q1; return { lo: q1 - 1.5 * iqr, hi: q3 + 1.5 * iqr }; }
export function crc32(s: string): string { let c: number; const table: number[] = []; for (let n = 0; n < 256; n++) { c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; table[n] = c; } const bytes = new TextEncoder().encode(s); let crc = 0 ^ -1; for (const b of bytes) crc = (crc >>> 8) ^ table[(crc ^ b) & 0xff]; return ((crc ^ -1) >>> 0).toString(16).padStart(8, '0'); }
export async function sha256Hex(s: string): Promise<string> { const d = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(s)); return [...new Uint8Array(d)].map((b) => b.toString(16).padStart(2, '0')).join(''); }
export function base64Encode(s: string): string { return btoa(String.fromCharCode(...new TextEncoder().encode(s))); }
export function base64Decode(s: string): string | null { try { return new TextDecoder().decode(Uint8Array.from(atob(s.trim()), (ch) => ch.charCodeAt(0))); } catch { return null; } }
export function normalizeUnicode(s: string, form: 'NFC' | 'NFD' | 'NFKC' | 'NFKD'): string { return s.normalize(form); }

defTools([
  {
    id: 'extra.clusterValues', label: 'Cluster Similar Values (Fingerprints)…', icon: 'workspaces', tab: X, group: G_QUALITY, edits: true, popular: true, tip: 'Find groups of values that are the same thing spelled differently ("New York", "new york ", "York, New") and replace them all with the most common spelling, or list the clusters', keywords: 'cluster fingerprint openrefine standardize variants',
    run: async (app, arg?: { method: 'fingerprint' | 'ngram'; apply: boolean }) => {
      const s = app.sheet; const cells = selCells(app, { existingOnly: true }).filter((x) => typeof x.cell?.v === 'string');
      const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Cluster Similar Values', fields: [{ key: 'method', label: 'Method', type: 'radio', value: 'fingerprint', options: [{ value: 'fingerprint', label: 'Fingerprint (case, punctuation, accents and word order ignored)' }, { value: 'ngram', label: '2-gram fingerprint (also catches small typos)' }] }, { key: 'apply', label: 'Replace every variant with the most common spelling', type: 'checkbox', value: false }] }); return r ? { method: r.method, apply: !!r.apply } : null; });
      if (!o) return;
      const clusters = clusterValues(cells.map((x) => x.cell!.v as string), o.method === 'ngram' ? (t) => ngramFingerprint(t) : fingerprint);
      if (!clusters.length) { toast('No clusters of similar values.'); return; }
      if (o.apply) { const map = new Map<string, string>(); for (const cl of clusters) for (const v of cl.variants) map.set(v.text, cl.variants[0].text); let n = 0; app.book.transact('Standardize Values', () => { for (const x of cells) { const t = (x.cell!.v as string).trim(); const to = map.get(t); if (to != null && to !== x.cell!.v) { s.set(x.r, x.c, { ...x.cell!, v: to }); n++; } } }); toast(`${plural(n, 'cell')} standardized in ${plural(clusters.length, 'cluster')}`); return; }
      report(app, 'Clusters', ['Cluster', 'Most common', 'Variants', 'Count'], clusters.map((cl, i) => [i + 1, cl.variants[0].text, cl.variants.map((v) => `${v.text} (${v.n})`).join(' | '), cl.variants.reduce((a, v) => a + v.n, 0)]), { colW: [60, 200, 420, 70] });
    },
  },
  {
    id: 'extra.outliers', label: 'Highlight Outliers (IQR Rule)', icon: 'scatter_plot', tab: X, group: G_QUALITY, edits: true, tip: 'Color the numbers that fall outside 1.5 × the interquartile range of the selection',
    run: (app) => { const s = app.sheet; const cells = selCells(app, { existingOnly: true }).filter((x) => typeof x.cell?.v === 'number'); const b = iqrOutliers(cells.map((x) => x.cell!.v as number)); if (!b) { toast('Select at least four numbers.'); return; } const hits = cells.filter((x) => (x.cell!.v as number) < b.lo || (x.cell!.v as number) > b.hi); if (!hits.length) { toast('No outliers.'); return; } app.book.transact('Outliers', () => { for (const x of hits) s.set(x.r, x.c, { ...x.cell!, s: app.book.derivedStyle(x.cell!.s ?? 0, { fill: '#ffcdd2' }) }); }); toast(`${plural(hits.length, 'outlier')} colored (below ${b.lo.toFixed(2)} or above ${b.hi.toFixed(2)})`); },
  },
  { id: 'extra.zScores', label: 'Z-Score Column', icon: 'functions', tab: X, group: G_QUALITY, edits: true, tip: 'How many standard deviations each value is from the column mean (formulas)', run: (app) => { const s = app.sheet; const rg = clipToUsed(s, activeRange(app)); const header = headerGuess(s, rg); const c = rg.c2 + 1; const first = rg.r1 + (header ? 1 : 0); if (!checkWritable(app, [{ r1: rg.r1, r2: rg.r2, c1: c, c2: c }])) return; const all = `${cellName(first, rg.c1, true, true)}:${cellName(rg.r2, rg.c1, true, true)}`; app.book.transact('Z-Scores', () => { if (header) writeOut(s, rg.r1, c, { v: 'z', style: HEAD }); for (let r = first; r <= rg.r2; r++) writeOut(s, r, c, { f: `IF(ISNUMBER(${cellName(r, rg.c1)}),(${cellName(r, rg.c1)}-AVERAGE(${all}))/STDEV.S(${all}),"")`, style: { numFmt: '0.00' } }); }); } },
  { id: 'extra.hashColumn', label: 'Hash or Checksum Beside Each Cell', icon: 'tag', tab: X, group: G_QUALITY, edits: true, kind: 'split', tip: 'CRC-32 or SHA-256 of each selected value, for matching records without exposing them', run: async (app, arg?: 'crc32' | 'sha256') => { if (arg === 'sha256') { const s = app.sheet; const rg = clipToUsed(s, activeRange(app)); const target = { r1: rg.r1, r2: rg.r2, c1: rg.c2 + 1, c2: rg.c2 + 1 }; if (!checkWritable(app, [target])) return; const jobs: [number, string][] = []; for (let r = rg.r1; r <= rg.r2; r++) { const t = shownText(s, r, rg.c1); if (t) jobs.push([r, t]); } const hashes = await Promise.all(jobs.map((j) => sha256Hex(j[1]))); app.book.transact('SHA-256', () => jobs.forEach((j, i) => writeOut(s, j[0], target.c1, hashes[i]))); return; } fillBeside(app, 'CRC-32', (t) => crc32(t)); }, menu: () => [{ id: 'extra.hashColumn', arg: 'crc32', label: 'CRC-32' }, { id: 'extra.hashColumn', arg: 'sha256', label: 'SHA-256' }] },
  { id: 'extra.base64', label: 'Base64 Encode / Decode', icon: 'code', tab: X, group: G_QUALITY, edits: true, kind: 'split', run: (app, arg?: 'encode' | 'decode') => transformValues(app, 'Base64', (v) => (arg === 'decode' ? base64Decode(String(v)) ?? undefined : base64Encode(String(v)))), menu: () => [{ id: 'extra.base64', arg: 'encode', label: 'Encode' }, { id: 'extra.base64', arg: 'decode', label: 'Decode' }] },
  { id: 'extra.unicodeNormalize', label: 'Normalize Unicode (NFC / NFKC)', icon: 'translate', tab: X, group: G_QUALITY, edits: true, kind: 'split', tip: 'Make equal-looking text equal: composed accents (NFC), or compatibility forms that also fold ligatures and full-width characters (NFKC)', run: (app, arg?: 'NFC' | 'NFD' | 'NFKC' | 'NFKD') => transformValues(app, 'Normalize Unicode', (v) => normalizeUnicode(String(v), arg || 'NFC')), menu: () => (['NFC', 'NFKC', 'NFD', 'NFKD'] as const).map((f) => ({ id: 'extra.unicodeNormalize', arg: f, label: f })) },
  { id: 'extra.showInvisible', label: 'Reveal Invisible Characters', icon: 'visibility', tab: X, group: G_QUALITY, edits: true, tip: 'Select the cells that contain non-breaking spaces, zero-width characters, tabs, trailing spaces or other characters you cannot see', run: (app) => { const re = /[ ​-‍⁠﻿\t\r­ -  ]|^\s|\s$/; selectHits(app, selCells(app, { existingOnly: true }).filter((x) => typeof x.cell?.v === 'string' && re.test(x.cell.v)).map((x) => ({ r: x.r, c: x.c })), 'cells with invisible characters'); } },
  { id: 'extra.mixedTypes', label: 'Find Columns with Mixed Types', icon: 'rule', tab: X, group: G_QUALITY, edits: true, tip: 'Columns that mix numbers and text (often numbers stored as text), with the odd cells listed', run: (app) => { const s = app.sheet; const rg = selBlock(app); const header = headerGuess(s, rg); const rows: Scalar[][] = []; for (let c = rg.c1; c <= rg.c2; c++) { let nums = 0, texts = 0; const odd: string[] = []; for (let r = rg.r1 + (header ? 1 : 0); r <= rg.r2; r++) { const v = s.value(r, c); if (typeof v === 'number') nums++; else if (typeof v === 'string' && v !== '') { texts++; if (textToNumber(v) != null) odd.push(cellName(r, c)); } } if (nums && texts) rows.push([header ? shownText(s, rg.r1, c) : colName(c), nums, texts, odd.length, odd.slice(0, 30).join(', ')]); } if (!rows.length) { toast('Every column has one kind of value.'); return; } report(app, 'Mixed Types', ['Column', 'Numbers', 'Text', 'Numbers as text', 'Cells'], rows, { colW: [160, 80, 80, 120, 400] }); } },
  { id: 'extra.menuQuality', label: 'Data Quality', icon: 'verified', tab: X, group: G_QUALITY, kind: 'menu', noRecord: true, run: () => undefined, menu: () => menuOf(['extra.clusterValues', 'extra.outliers', 'extra.zScores', 'extra.mixedTypes', 'extra.showInvisible', 'extra.unicodeNormalize', 'extra.hashColumn', 'extra.base64', 'tools.dataProfile', 'tools.findDuplicates', 'tools.fuzzyMatch', 'util.findProblems']) },
]);

// ============================================================================================
// Dates

/** Easter Sunday (Gregorian, Anonymous/Meeus algorithm). */
export function easter(y: number): { m: number; d: number } { const a = y % 19, b = Math.floor(y / 100), c = y % 100, d = Math.floor(b / 4), e = b % 4, f = Math.floor((b + 8) / 25), g = Math.floor((b - f + 1) / 3), h = (19 * a + b - d - g + 15) % 30, i = Math.floor(c / 4), k = c % 4, l = (32 + 2 * e + 2 * i - h - k) % 7, m = Math.floor((a + 11 * h + 22 * l) / 451); const month = Math.floor((h + l - 7 * m + 114) / 31), day = ((h + l - 7 * m + 114) % 31) + 1; return { m: month, d: day }; }

function nthWeekday(y: number, m: number, dow: number, n: number, date1904 = false): number { const first = partsToSerial(y, m, 1, date1904); const fdow = serialToParts(first, date1904).dow; let d = 1 + ((dow - fdow + 7) % 7) + (n - 1) * 7; if (n < 0) { const last = daysInMonth(y, m); const ldow = serialToParts(partsToSerial(y, m, last, date1904), date1904).dow; d = last - ((ldow - dow + 7) % 7); } return partsToSerial(y, m, d, date1904); }
function observed(serial: number, date1904 = false): number { const dow = serialToParts(serial, date1904).dow; return dow === 6 ? serial - 1 : dow === 0 ? serial + 1 : serial; }

/** United States federal holidays for a year, as observed. */
export function usHolidays(y: number, date1904 = false): { name: string; serial: number }[] {
  const fixed = (m: number, d: number) => observed(partsToSerial(y, m, d, date1904), date1904);
  const out = [{ name: "New Year's Day", serial: fixed(1, 1) }, { name: 'Martin Luther King Jr. Day', serial: nthWeekday(y, 1, 1, 3, date1904) }, { name: "Washington's Birthday", serial: nthWeekday(y, 2, 1, 3, date1904) }, { name: 'Memorial Day', serial: nthWeekday(y, 5, 1, -1, date1904) }];
  if (y >= 2021) out.push({ name: 'Juneteenth', serial: fixed(6, 19) });
  out.push({ name: 'Independence Day', serial: fixed(7, 4) }, { name: 'Labor Day', serial: nthWeekday(y, 9, 1, 1, date1904) }, { name: 'Columbus Day', serial: nthWeekday(y, 10, 1, 2, date1904) }, { name: 'Veterans Day', serial: fixed(11, 11) }, { name: 'Thanksgiving Day', serial: nthWeekday(y, 11, 4, 4, date1904) }, { name: 'Christmas Day', serial: fixed(12, 25) });
  return out.sort((a, b) => a.serial - b.serial);
}

export function weekNumberIsoYmd(p: Ymd): number { const d = new Date(Date.UTC(p.y, p.m - 1, p.d)); if (p.y < 100) d.setUTCFullYear(p.y); const day = d.getUTCDay() || 7; d.setUTCDate(d.getUTCDate() + 4 - day); const yStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1)); return Math.ceil(((d.getTime() - yStart.getTime()) / 86400000 + 1) / 7); }
export function weekNumberIso(serial: number, date1904 = false): number { const p = serialToParts(serial, date1904); return weekNumberIsoYmd({ y: p.y, m: p.m, d: p.d }); }

defTools([
  {
    id: 'extra.dateColumns', label: 'Date Parts Columns (Weekday, Week, Quarter…)', icon: 'date_range', tab: X, group: G_DATES, edits: true, popular: true, tip: 'Beside each date: weekday name, ISO week number, quarter, month name, day of year and whether it is a weekend',
    run: (app) => { const s = app.sheet; const rg = clipToUsed(s, activeRange(app)); const header = headerGuess(s, rg); const target = { r1: rg.r1, r2: rg.r2, c1: rg.c2 + 1, c2: rg.c2 + 6 }; if (!checkWritable(app, [target])) return; const d1904 = app.book.calc.date1904; app.book.transact('Date Parts', () => { for (let r = rg.r1; r <= rg.r2; r++) { if (r === rg.r1 && header) { ['Weekday', 'ISO week', 'Quarter', 'Month', 'Day of year', 'Weekend'].forEach((t, j) => writeOut(s, r, target.c1 + j, { v: t, style: HEAD })); continue; } const p = dateOf(app, s.value(r, rg.c1), r, rg.c1); if (p == null) continue; const dow = (jdn(p) + 1) % 7; const doy = jdn(p) - jdn({ y: p.y, m: 1, d: 1 }) + 1; [DAYS[dow], weekNumberIsoYmd(p), `Q${Math.floor((p.m - 1) / 3) + 1}`, MONTHS[p.m - 1], doy, dow === 0 || dow === 6 ? 'Yes' : 'No'].forEach((v, j) => writeOut(s, r, target.c1 + j, v)); void d1904; } }); },
  },
  {
    id: 'extra.holidays', label: 'Holiday List for a Year…', icon: 'celebration', tab: X, group: G_DATES, edits: true, tip: 'United States federal holidays (as observed) plus Easter and related days for a year, on a new sheet, ready for WORKDAY and NETWORKDAYS',
    run: async (app, arg?: { year: number }) => { const o = await argOrAsk(arg, async () => { const t = await promptDialog('Holiday List', 'Year', String(new Date().getFullYear()), { type: 'number' }); return t ? { year: Number(t) } : null; }); if (!o) return; const d1904 = app.book.calc.date1904; const e = easter(o.year); const es = partsToSerial(o.year, e.m, e.d, d1904); const rows = [...usHolidays(o.year, d1904).map((hh) => [hh.name, hh.serial, 'Federal']), ['Good Friday', es - 2, 'Christian'], ['Easter Sunday', es, 'Christian'], ['Easter Monday', es + 1, 'Christian'], ['Ash Wednesday', es - 46, 'Christian'], ['Pentecost', es + 49, 'Christian']].sort((a, b) => Number(a[1]) - Number(b[1])); report(app, `Holidays ${o.year}`, ['Holiday', 'Date', 'Kind'], rows.map((r) => [r[0], { v: r[1], style: { numFmt: 'ddd d mmm yyyy' } }, r[2]] as any), { colW: [220, 150, 100] }); },
  },
  { id: 'extra.easter', label: 'Easter Date Beside Years', icon: 'egg', tab: X, group: G_DATES, edits: true, run: (app) => { const d1904 = app.book.calc.date1904; const s = app.sheet; const rg = clipToUsed(s, activeRange(app)); const target = { r1: rg.r1, r2: rg.r2, c1: rg.c2 + 1, c2: rg.c2 + 1 }; if (!checkWritable(app, [target])) return; app.book.transact('Easter', () => { for (let r = rg.r1; r <= rg.r2; r++) { const y = numOf(s.value(r, rg.c1)); if (y == null || y < 1583 || y > 4099) continue; const e = easter(Math.trunc(y)); writeOut(s, r, target.c1, { v: partsToSerial(Math.trunc(y), e.m, e.d, d1904), style: { numFmt: 'yyyy-mm-dd' } }); } }); } },
  {
    id: 'extra.businessDayAdjust', label: 'Move Dates off Weekends', icon: 'event_available', tab: X, group: G_DATES, edits: true, kind: 'split', tip: 'Dates that fall on a Saturday or Sunday move to the next (or previous) weekday',
    run: (app, arg?: 'next' | 'prev') => { const d1904 = app.book.calc.date1904; changeCells(app, 'Move off Weekends', (cell, r, c) => { if (typeof cell.v !== 'number' || !isDateFormat(app.sheet.styleOf(r, c).numFmt)) return undefined; const dow = serialToParts(Math.floor(cell.v), d1904).dow; if (dow !== 0 && dow !== 6) return undefined; const shift = arg === 'prev' ? (dow === 0 ? -2 : -1) : (dow === 6 ? 2 : 1); return { ...cell, v: cell.v + shift }; }); },
    menu: () => [{ id: 'extra.businessDayAdjust', arg: 'next', label: 'To the following Monday' }, { id: 'extra.businessDayAdjust', arg: 'prev', label: 'To the preceding Friday' }],
  },
  { id: 'extra.daysUntil', label: 'Days Until / Since Each Date', icon: 'timelapse', tab: X, group: G_DATES, edits: true, tip: 'A live formula beside each date: TODAY() minus the date (negative = in the future)', run: (app) => { const s = app.sheet; const rg = clipToUsed(s, activeRange(app)); const header = headerGuess(s, rg); const c = rg.c2 + 1; if (!checkWritable(app, [{ r1: rg.r1, r2: rg.r2, c1: c, c2: c }])) return; app.book.transact('Days Since', () => { if (header) writeOut(s, rg.r1, c, { v: 'Days', style: HEAD }); for (let r = rg.r1 + (header ? 1 : 0); r <= rg.r2; r++) { const p = dateOf(app, s.value(r, rg.c1), r, rg.c1); if (p == null) continue; if (typeof s.value(r, rg.c1) === 'number') writeOut(s, r, c, { f: `TODAY()-${cellName(r, rg.c1)}`, style: { numFmt: '0' } }); else writeOut(s, r, c, jdn(todayYmd()) - jdn(p)); } }); } },
  {
    id: 'extra.timeline', label: 'Timeline Sheet from Dates and Events', icon: 'timeline', tab: X, group: G_DATES, edits: true, tip: 'Two selected columns (date, event) become a sorted timeline with the gap in years and days between events',
    run: (app) => { const s = app.sheet; const rg = clipToUsed(s, activeRange(app)); const header = headerGuess(s, rg); const items: { p: Ymd; j: number; t: string }[] = []; for (let r = rg.r1 + (header ? 1 : 0); r <= rg.r2; r++) { const p = dateOf(app, s.value(r, rg.c1), r, rg.c1); if (p == null) continue; items.push({ p, j: jdn(p), t: rg.c2 > rg.c1 ? shownText(s, r, rg.c1 + 1) : '' }); } if (!items.length) { toast('No dates found in the first selected column.'); return; } items.sort((a, b) => a.j - b.j); const d1904 = app.book.calc.date1904; const rows = items.map((it, i) => { const gap = i ? it.j - items[i - 1].j : 0; return [ymdCell(it.p, d1904), it.p.y, it.t, i ? `${Math.floor(gap / 365.25)}y ${Math.round(gap % 365.25)}d` : '', i ? gap : null]; }); report(app, 'Timeline', ['Date', 'Year', 'Event', 'Since previous', 'Days'], rows as any, { colW: [110, 60, 320, 120, 70] }); },
  },
  { id: 'extra.menuDates', label: 'Dates', icon: 'event', tab: X, group: G_DATES, kind: 'menu', noRecord: true, run: () => undefined, menu: () => menuOf(['extra.dateColumns', 'extra.holidays', 'extra.easter', 'extra.businessDayAdjust', 'extra.daysUntil', 'extra.timeline', 'extra.ageAtDate', 'extra.gedcomDates', 'util.shiftDates', 'util.splitDates', 'tools.fixDates', 'tools.dateCalculator', 'tools.calendar']) },
]);

// ============================================================================================
// Geography

/** "41°24'12.2"N 2°10'26.5"E" or "41 24 12 N" → decimal degrees. */
export function dmsToDecimal(s: string): number | null {
  const t = s.trim().replace(/[′’']/g, "'").replace(/[″”"]/g, '"');
  const m = /^(-?\d+(?:\.\d+)?)\s*[°º\s]\s*(\d+(?:\.\d+)?)?\s*['\s]?\s*(\d+(?:\.\d+)?)?\s*"?\s*([NSEW])?$/i.exec(t);
  if (!m) { const n = Number(t); return isNaN(n) ? null : n; }
  let v = Math.abs(+m[1]) + (+(m[2] || 0)) / 60 + (+(m[3] || 0)) / 3600;
  if (+m[1] < 0 || /[SW]/i.test(m[4] || '')) v = -v;
  return Math.round(v * 1e6) / 1e6;
}
export function decimalToDms(v: number, axis: 'lat' | 'lon'): string { const a = Math.abs(v); let d = Math.floor(a), m = Math.floor((a - d) * 60), sec = Math.round(((a - d) * 60 - m) * 60 * 10) / 10; /* seconds that round up to 60 carry into minutes and degrees */ if (sec >= 60) { sec = 0; m++; } if (m >= 60) { m = 0; d++; } const hemi = axis === 'lat' ? (v < 0 ? 'S' : 'N') : (v < 0 ? 'W' : 'E'); return `${d}°${m}'${sec}"${hemi}`; }
/** Great-circle distance in kilometers. */
export function haversineKm(lat1: number, lon1: number, lat2: number, lon2: number): number { const R = 6371.0088; const toR = (x: number) => (x * Math.PI) / 180; const dLat = toR(lat2 - lat1), dLon = toR(lon2 - lon1); const a = Math.sin(dLat / 2) ** 2 + Math.cos(toR(lat1)) * Math.cos(toR(lat2)) * Math.sin(dLon / 2) ** 2; return 2 * R * Math.asin(Math.sqrt(a)); }
export function bearingDeg(lat1: number, lon1: number, lat2: number, lon2: number): number { const toR = (x: number) => (x * Math.PI) / 180; const y = Math.sin(toR(lon2 - lon1)) * Math.cos(toR(lat2)); const x = Math.cos(toR(lat1)) * Math.sin(toR(lat2)) - Math.sin(toR(lat1)) * Math.cos(toR(lat2)) * Math.cos(toR(lon2 - lon1)); return (Math.atan2(y, x) * 180 / Math.PI + 360) % 360; }
export function compassPoint(deg: number): string { const pts = ['N', 'NNE', 'NE', 'ENE', 'E', 'ESE', 'SE', 'SSE', 'S', 'SSW', 'SW', 'WSW', 'W', 'WNW', 'NW', 'NNW']; return pts[Math.round(deg / 22.5) % 16]; }

defTools([
  { id: 'extra.dmsToDecimal', label: 'Degrees-Minutes-Seconds to Decimal', icon: 'explore', tab: X, group: G_GEO, edits: true, popular: true, tip: '41°24\'12"N becomes 41.403333, in place', run: (app) => transformValues(app, 'DMS to Decimal', (v) => { if (typeof v === 'number') return undefined; const d = dmsToDecimal(String(v)); return d == null ? undefined : d; }) },
  { id: 'extra.decimalToDms', label: 'Decimal Degrees to Degrees-Minutes-Seconds', icon: 'explore', tab: X, group: G_GEO, edits: true, kind: 'split', run: (app, arg?: 'lat' | 'lon') => transformValues(app, 'Decimal to DMS', (v) => (typeof v === 'number' ? decimalToDms(v, arg || 'lat') : undefined), { numbersToo: true }), menu: () => [{ id: 'extra.decimalToDms', arg: 'lat', label: 'As latitude (N/S)' }, { id: 'extra.decimalToDms', arg: 'lon', label: 'As longitude (E/W)' }] },
  {
    id: 'extra.distance', label: 'Distance and Bearing Between Coordinates', icon: 'straighten', tab: X, group: G_GEO, edits: true, popular: true, tip: 'Four selected columns (lat1, lon1, lat2, lon2) give the great-circle distance in km and miles and the compass bearing in the columns to the right', keywords: 'haversine distance coordinates gps bearing',
    run: (app) => { const s = app.sheet; const rg = clipToUsed(s, activeRange(app)); if (rg.c2 - rg.c1 < 3) { toast('Select four columns: latitude 1, longitude 1, latitude 2, longitude 2.'); return; } const header = headerGuess(s, rg); const target = { r1: rg.r1, r2: rg.r2, c1: rg.c2 + 1, c2: rg.c2 + 3 }; if (!checkWritable(app, [target])) return; app.book.transact('Distance', () => { for (let r = rg.r1; r <= rg.r2; r++) { if (r === rg.r1 && header) { ['km', 'miles', 'Bearing'].forEach((t, j) => writeOut(s, r, target.c1 + j, { v: t, style: HEAD })); continue; } const v = [0, 1, 2, 3].map((j) => { const x = s.value(r, rg.c1 + j); return typeof x === 'number' ? x : dmsToDecimal(textOf(x)); }); if (v.some((x) => x == null)) continue; const km = haversineKm(v[0]!, v[1]!, v[2]!, v[3]!); const b = bearingDeg(v[0]!, v[1]!, v[2]!, v[3]!); writeOut(s, r, target.c1, { v: Math.round(km * 1000) / 1000, style: { numFmt: '#,##0.0' } }); writeOut(s, r, target.c1 + 1, { v: Math.round(km * 0.621371 * 1000) / 1000, style: { numFmt: '#,##0.0' } }); writeOut(s, r, target.c1 + 2, `${Math.round(b)}° ${compassPoint(b)}`); } }); },
  },
  { id: 'extra.mapLinks', label: 'Map Links from Coordinates or Addresses', icon: 'map', tab: X, group: G_GEO, edits: true, tip: 'Beside each coordinate pair (or address text), a link that opens the place on OpenStreetMap', run: (app) => { const s = app.sheet; const rg = clipToUsed(s, activeRange(app)); const target = { r1: rg.r1, r2: rg.r2, c1: rg.c2 + 1, c2: rg.c2 + 1 }; if (!checkWritable(app, [target])) return; app.book.transact('Map Links', () => { for (let r = rg.r1; r <= rg.r2; r++) { const a = s.value(r, rg.c1), b = rg.c2 > rg.c1 ? s.value(r, rg.c1 + 1) : null; let url: string | null = null; if (typeof a === 'number' && typeof b === 'number') url = `https://www.openstreetmap.org/?mlat=${a}&mlon=${b}#map=14/${a}/${b}`; else if (typeof a === 'string' && a.trim()) url = `https://www.openstreetmap.org/search?query=${encodeURIComponent(a.trim())}`; if (!url) continue; s.set(r, target.c1, { v: 'Map', link: url, s: app.book.styleId({ color: '#0563c1', underline: 'single' }) }); } }); } },
  { id: 'extra.menuGeo', label: 'Geography', icon: 'public', tab: X, group: G_GEO, kind: 'menu', noRecord: true, run: () => undefined, menu: () => menuOf(['extra.dmsToDecimal', 'extra.decimalToDms', 'extra.distance', 'extra.mapLinks', 'util.convertUnits', 'view.mapView']) },
]);

// ============================================================================================
// Workflow

defTools([
  {
    id: 'extra.auditStamp', label: 'Stamp with User and Time', icon: 'approval', tab: X, group: G_FLOW, edits: true, popular: true, keys: ['Mod-Alt-s'], tip: 'Write your name and the current date and time into the active cell (or beside the selection) as a fixed audit mark',
    run: (app, arg?: 'cell' | 'right') => { const s = app.sheet; const stamp = `${app.prefs.author || 'User'} ${new Date().toISOString().slice(0, 16).replace('T', ' ')}`; app.book.transact('Audit Stamp', () => { if (arg === 'right') { const rg = clipToUsed(s, activeRange(app)); for (let r = rg.r1; r <= rg.r2; r++) if (hasContent(s.get(r, rg.c1))) writeOut(s, r, rg.c2 + 1, stamp); } else { const { r, c } = activeCell(app); writeOut(s, r, c, stamp); } }); },
  },
  { id: 'extra.todoColumn', label: 'Add a To-Do Check Box Column', icon: 'checklist', tab: X, group: G_FLOW, edits: true, tip: 'A "Done" column of check boxes beside the selected rows, with a rule that strikes through finished rows', run: (app) => { const s = app.sheet; const rg = selBlock(app); const header = headerGuess(s, rg); const c = rg.c2 + 1; if (!checkWritable(app, [{ r1: rg.r1, r2: rg.r2, c1: c, c2: c }])) return; app.book.transact('To-Do Column', () => { app.book.recordSheet(s); if (header) writeOut(s, rg.r1, c, { v: 'Done', style: HEAD }); for (let r = rg.r1 + (header ? 1 : 0); r <= rg.r2; r++) s.set(r, c, { v: false, s: app.book.derivedStyle(s.styleIdOf(r, c), { checkbox: true, hAlign: 'center' }) }); insertRule(s, { type: 'expression', formulas: [`$${colName(c)}${rg.r1 + (header ? 1 : 0) + 1}=TRUE`], ranges: [{ r1: rg.r1 + (header ? 1 : 0), r2: rg.r2, c1: rg.c1, c2: c }], style: { strike: true, color: '#9e9e9e' } }); }); app.tab.cf.invalidate(); app.grid.invalidate(); } },
  {
    id: 'extra.compareSnapshot', label: 'What Changed Since the Last Snapshot', icon: 'difference', tab: X, group: G_FLOW, edits: true, popular: true, tip: 'Compare the workbook with its most recent snapshot (File ▸ Take Snapshot) and list every changed cell', keywords: 'changes since snapshot diff version',
    run: async (app) => { const { docKey } = await import('./file'); const { Workbook } = await import('../model/workbook'); const { compareSheetCells } = await import('./tools'); const key = docKey(app.tab); const keys = (await kvKeys(`snap:${key}:`)).sort(); if (!keys.length) { toast('No snapshots of this workbook yet (File ▸ Take Snapshot).'); return; } const snap = await kvGet<{ json: any; when: number; label: string }>(keys[keys.length - 1]); if (!snap) return; const old = Workbook.fromJSON(snap.json); const rows: Scalar[][] = []; for (const s of app.book.sheets) { const o = old.sheetByName(s.name); if (!o) { rows.push([s.name, '', 'New sheet', '', '']); continue; } for (const d of compareSheetCells(o, s, true)) rows.push([s.name, cellName(d.r, d.c), d.kind, textOf(d.a), textOf(d.b)]); } for (const o of old.sheets) if (!app.book.sheetByName(o.name)) rows.push([o.name, '', 'Sheet removed', '', '']); if (!rows.length) { toast(`No changes since the snapshot of ${new Date(snap.when).toLocaleString()}.`); return; } report(app, 'Changes', ['Sheet', 'Cell', 'Change', `Snapshot ${new Date(snap.when).toLocaleString()}`, 'Now'], rows, { colW: [110, 60, 140, 220, 220] }); },
  },
  { id: 'extra.changeLog', label: 'Undo History as a Sheet', icon: 'history', tab: X, group: G_FLOW, edits: true, tip: 'The labels of the editing steps still in the undo history, newest first', run: (app) => { const rows = app.book.undoStack.map((t, i) => [app.book.undoStack.length - i, t.label, t.cells.size, t.sheets.size ? 'Yes' : '', t.book ? 'Yes' : ''] as Scalar[]).reverse(); if (!rows.length) { toast('The undo history is empty.'); return; } report(app, 'History', ['#', 'Step', 'Cells changed', 'Sheet settings', 'Workbook settings'], rows); } },
  { id: 'extra.selectionToQat', label: 'Put the Last Tool on the Quick Access Bar', icon: 'bolt', tab: X, group: G_FLOW, noRecord: true, run: async (app) => { const { toolLog } = await import('./utilShared'); const l = toolLog.last; if (!l) { toast('No tool used yet.'); return; } if (!app.prefs.qat.includes(l.id)) { app.prefs.qat.push(l.id); app.savePrefs(); app.ui.applyPrefs(); } toast(`${l.label} is on the Quick Access Bar`); } },
  { id: 'extra.zebraPrint', label: 'Zebra Rows for Printing', icon: 'view_agenda', tab: X, group: G_FLOW, edits: true, tip: 'A light conditional-format banding that survives sorting and filtering (MOD(ROW(),2))', run: (app) => { const s = app.sheet; const rg = selBlock(app); app.book.transact('Zebra Rows', () => { app.book.recordSheet(s); insertRule(s, { type: 'expression', formulas: ['MOD(ROW(),2)=0'], ranges: [rg], style: { fill: '#f3f6f9' } }); }); app.tab.cf.invalidate(); app.grid.invalidate(); } },
  { id: 'extra.lockFinished', label: 'Lock the Rows Marked Done', icon: 'lock', tab: X, group: G_FLOW, edits: true, tip: 'Rows whose last selected cell is TRUE or "Done" become locked (protect the sheet to enforce it)', run: (app) => { const s = app.sheet; const rg = selBlock(app); let n = 0; app.book.transact('Lock Finished', () => { for (let r = rg.r1; r <= rg.r2; r++) { const v = s.value(r, rg.c2); const done = v === true || /^(done|yes|x|✓|complete[d]?)$/i.test(textOf(v)); if (!done) continue; for (let c = rg.c1; c <= rg.c2; c++) s.set(r, c, { ...(s.get(r, c) || {}), s: app.book.derivedStyle(s.styleIdOf(r, c), { locked: undefined, fill: '#eeeeee' }) }); n++; } for (let r = rg.r1; r <= rg.r2; r++) { const v = s.value(r, rg.c2); if (v === true) continue; } }); toast(`${plural(n, 'row')} locked (unlock the others and protect the sheet)`); } },
  { id: 'extra.menuWorkflow', label: 'Workflow', icon: 'task_alt', tab: X, group: G_FLOW, kind: 'menu', noRecord: true, run: () => undefined, menu: () => menuOf(['extra.auditStamp', 'extra.todoColumn', 'extra.lockFinished', 'extra.zebraPrint', 'extra.compareSnapshot', 'extra.changeLog', 'extra.selectionToQat', 'file.snapshot', 'file.versions', 'util.myTools', 'util.lastTool']) },
]);

void serialOf; void jsDateToSerial; void cellKey; void keyRow; void keyCol; void rangeName; void selRanges; void singleCell; void sheetProtected; void emptySheet; void patchStyle; void insertRows; void copyText; void valueRows; void alertDialog; void CellError; void textToNumber;
export type ExtraCell = Cell; export type ExtraStyle = CellStyle; export type ExtraRange = Range;
