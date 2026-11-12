/**
 * Delimited text (CSV / TSV / tab-delimited TXT): an RFC 4180 parser and writer, delimiter
 * detection, encoding handling and the type detection used by the Text Import Wizard.
 */
import { Workbook, Sheet } from '../model/workbook';
import { CellError, errorFromCode, type Cell, type Scalar } from '../model/types';
import { cellKey, type Range } from '../model/address';
import { formatValue, parseNumberInput } from '../model/numfmt';
import { parseDateTime } from '../model/dates';
import { type CsvOptions } from './index';
import { decodeAuto } from './encoding';

const CANDIDATES = [',', ';', '\t', '|'];

/** Guesses the field delimiter from the first lines of a file (quote-aware). */
export function detectDelimiter(text: string): string {
  const sample = text.replace(/^﻿/, '').slice(0, 64 * 1024);
  let best = ',', bestScore = -1;
  for (const d of CANDIDATES) {
    const counts: number[] = [];
    let n = 0, inQ = false;
    for (let i = 0; i < sample.length && counts.length < 30; i++) {
      const ch = sample[i];
      if (ch === '"') { if (inQ && sample[i + 1] === '"') { i++; continue; } inQ = !inQ; continue; }
      if (inQ) continue;
      if (ch === d) n++;
      else if (ch === '\n' || ch === '\r') {
        if (ch === '\r' && sample[i + 1] === '\n') i++;
        counts.push(n); n = 0;
      }
    }
    if (n || !counts.length) counts.push(n);
    const lines = counts.filter((_, i) => i < counts.length);
    const freq = new Map<number, number>();
    for (const c of lines) if (c > 0) freq.set(c, (freq.get(c) || 0) + 1);
    if (!freq.size) continue;
    let mode = 0, modeN = 0;
    for (const [c, k] of freq) if (k > modeN || (k === modeN && c > mode)) { mode = c; modeN = k; }
    const score = (modeN / lines.length) * 1000 + Math.min(mode, 50) + (d === ',' ? 0.5 : d === '\t' ? 0.4 : 0);
    if (score > bestScore) { bestScore = score; best = d; }
  }
  return best;
}

/** Parses delimited text into rows of fields (RFC 4180: quoted fields, "" escapes, embedded newlines). */
export function parseCsv(text: string, opts: CsvOptions = {}): string[][] {
  const src = text.replace(/^﻿/, '');
  const d = opts.delimiter || detectDelimiter(src);
  const q = opts.quote === undefined ? '"' : opts.quote;
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let i = 0;
  const n = src.length;
  let atFieldStart = true;
  while (i < n) {
    const ch = src[i];
    if (q && ch === q && atFieldStart) {
      // Quoted field
      i++;
      let s = '';
      for (;;) {
        const j = src.indexOf(q, i);
        if (j < 0) { s += src.slice(i); i = n; break; }
        s += src.slice(i, j);
        if (src[j + 1] === q) { s += q; i = j + 2; continue; }
        i = j + 1;
        break;
      }
      field += s;
      // Anything between the closing quote and the next delimiter is kept literally.
      while (i < n && src[i] !== d && src[i] !== '\n' && src[i] !== '\r') field += src[i++];
      atFieldStart = false;
      continue;
    }
    if (ch === d) { row.push(field); field = ''; atFieldStart = true; i++; continue; }
    if (ch === '\n' || ch === '\r') {
      row.push(field); field = ''; rows.push(row); row = []; atFieldStart = true;
      i += ch === '\r' && src[i + 1] === '\n' ? 2 : 1;
      continue;
    }
    field += ch; atFieldStart = false; i++;
  }
  if (field !== '' || !atFieldStart || row.length) { row.push(field); rows.push(row); }
  return rows;
}

/** Converts one field to a typed cell value, with the number format that suits it. */
export function convertField(s: string, opts: CsvOptions = {}, colType: 'general' | 'text' | 'date' | 'skip' = 'general'): { v: Scalar; f?: string; fmt?: string } {
  if (colType === 'text') return { v: s === '' ? null : s, fmt: '@' };
  if (s === '') return { v: null };
  const t = s.trim();
  const dmy = opts.dateOrder === 'dmy';
  if (colType === 'date') {
    const d = parseDate(t, opts);
    return d ? { v: d.serial, fmt: d.format } : { v: s };
  }
  if (opts.detectTypes === false) return { v: s };
  if (t.length > 1 && t[0] === '=') return { v: null, f: t.slice(1) };
  if (t.length > 1 && t[0] === "'" && t[1] === '=') return { v: t.slice(1) };
  const up = t.toUpperCase();
  if (up === 'TRUE' || up === 'FALSE') return { v: up === 'TRUE' };
  const err = t[0] === '#' ? errorFromCode(t) : null;
  if (err) return { v: err };
  // Numbers: normalize locale separators first.
  let numText = t;
  const dec = opts.decimal || '.';
  const thousands = opts.thousands ?? (dec === ',' ? '.' : ',');
  // Group separators are removed only where they group digits in threes ("1.234.567,8"), so
  // dates such as 15.03.2024 are not read as numbers.
  const grouped = (sep: string) => numText.split(sep).slice(1).every((g) => /^\d{3}/.test(g)) && /\d$/.test(numText.split(sep)[0]);
  if (dec === ',') {
    numText = numText.replace(/(\d)[\s  ](?=\d{3}(?!\d))/g, '$1');
    if (thousands && numText.includes(thousands)) numText = grouped(thousands) ? numText.split(thousands).join('') : '\u0000';
    if (/^[^,]*,[^,]*$/.test(numText)) numText = numText.replace(',', '.');
    else if (numText.includes(',')) numText = '\u0000';
  } else if (thousands && thousands !== ',' && numText.includes(thousands)) {
    numText = grouped(thousands) ? numText.split(thousands).join('') : '\u0000';
  }
  if (/[\d]/.test(numText)) {
    const n = parseNumberInput(numText);
    if (n) return { v: n.value, fmt: n.format };
  }
  if (/[-/.:]|\b[a-z]{3}/i.test(t) && /\d/.test(t)) {
    const d = parseDate(t, opts);
    if (d) return { v: d.serial, fmt: d.format };
  }
  void dmy;
  return { v: s };
}

function parseDate(t: string, opts: CsvOptions): { serial: number; format: string } | null {
  if (opts.dateOrder === 'ymd') {
    const m = /^(\d{2,4})[-/.](\d{1,2})[-/.](\d{1,2})(.*)$/.exec(t);
    if (m) {
      const y = m[1].length === 2 ? (+m[1] < 30 ? 2000 + +m[1] : 1900 + +m[1]) : +m[1];
      return parseDateTime(`${y}-${m[2]}-${m[3]}${m[4]}`);
    }
  }
  return parseDateTime(t, { dmy: opts.dateOrder === 'dmy' });
}

/** Builds a workbook (one sheet) from parsed rows. */
export function rowsToBook(rows: string[][], opts: CsvOptions = {}, sheetName = 'Sheet1'): Workbook {
  const book = new Workbook(false);
  const sheet = new Sheet(book, sheetName.replace(/[\\/?*[\]:]/g, '_').slice(0, 31) || 'Sheet1');
  book.sheets.push(sheet);
  fillSheet(sheet, rows, opts);
  book.clearHistory();
  return book;
}

export function fillSheet(sheet: Sheet, rows: string[][], opts: CsvOptions = {}, r0 = 0, c0 = 0): void {
  const book = sheet.book;
  const styleCache = new Map<string, number>();
  const sid = (fmt: string) => { let s = styleCache.get(fmt); if (s == null) { s = book.styleId({ numFmt: fmt }); styleCache.set(fmt, s); } return s; };
  const types = opts.columnTypes || [];
  const header = opts.headerRows || 0;
  rows.forEach((row, ri) => {
    let cOut = 0;
    row.forEach((field, ci) => {
      const ct = types[ci] || 'general';
      if (ct === 'skip') return;
      const c = c0 + cOut++;
      const conv = ri < header ? { v: field === '' ? null : field as Scalar } as ReturnType<typeof convertField> : convertField(field, opts, ct);
      if (conv.v == null && conv.f == null && !conv.fmt) return;
      const cell: Cell = {};
      if (conv.v != null) cell.v = conv.v;
      if (conv.f != null) cell.f = conv.f;
      if (conv.fmt && conv.fmt !== 'General') cell.s = sid(conv.fmt);
      if (cell.v == null && cell.f == null && conv.fmt !== '@') return;
      sheet.cells.set(cellKey(r0 + ri, c), cell);
    });
  });
}

export function importCsv(data: ArrayBuffer, name: string, opts: CsvOptions = {}): Workbook {
  let { text } = decodeAuto(new Uint8Array(data), opts.encoding);
  const o: CsvOptions = { ...opts };
  // Excel's "sep=;" first line names the delimiter.
  const sep = /^﻿?sep=(.)\r?\n/i.exec(text);
  if (sep) { text = text.slice(sep[0].length); if (!o.delimiter) o.delimiter = sep[1]; }
  if (!o.delimiter) o.delimiter = detectDelimiter(text);
  const rows = parseCsv(text, o);
  if (o.decimal == null || o.dateOrder == null) {
    const loc = detectLocale(rows, o.delimiter);
    if (o.decimal == null) { o.decimal = loc.decimal; if (o.thousands == null) o.thousands = loc.thousands; }
    if (o.dateOrder == null && loc.dateOrder) o.dateOrder = loc.dateOrder;
  }
  const base = name.replace(/^.*[\\/]/, '').replace(/\.[^.]*$/, '') || 'Sheet1';
  return rowsToBook(rows, o, base);
}

/**
 * Guesses the number and date conventions of delimited data: decimal comma ("1.234,56", or
 * "3,5" in a file not delimited by commas) versus decimal point, and day-month versus month-day
 * dates (a first part above 12 means day first; dotted dates are day first).
 */
export function detectLocale(rows: string[][], delimiter = ','): { decimal: '.' | ','; thousands: string; dateOrder?: 'mdy' | 'dmy' } {
  let commaDec = 0, pointDec = 0, dmy = 0, mdy = 0, dotted = 0;
  const lim = Math.min(rows.length, 2000);
  for (let i = 0; i < lim; i++) {
    for (const f of rows[i]) {
      const t = f.trim();
      if (!t || t.length > 40) continue;
      let m: RegExpExecArray | null;
      if (/^[-+]?[$€£¥]?\s?\d{1,3}(\.\d{3})+,\d+\s?[€$£%]?$/.test(t) || (delimiter !== ',' && /^[-+]?\d+,\d+\s?[€%]?$/.test(t) && !/^\d{1,3},\d{3}$/.test(t))) commaDec++;
      else if (/^[-+]?[$€£¥]?\d{1,3}(,\d{3})+\.\d+%?$/.test(t) || (/^[-+]?\d+\.\d+%?$/.test(t) && !/^\d{1,3}\.\d{3}$/.test(t))) pointDec++;
      if ((m = /^(\d{1,2})([./-])(\d{1,2})\2(\d{2}|\d{4})(?:\s|$)/.exec(t))) {
        if (+m[1] > 12 && +m[3] <= 12) dmy++;
        else if (+m[3] > 12 && +m[1] <= 12) mdy++;
        if (m[2] === '.') dotted++;
      }
    }
  }
  const decimal: '.' | ',' = commaDec > pointDec ? ',' : '.';
  let dateOrder: 'mdy' | 'dmy' | undefined;
  if (dmy > mdy) dateOrder = 'dmy';
  else if (mdy > dmy) dateOrder = 'mdy';
  else if (dotted || decimal === ',') dateOrder = 'dmy';
  return { decimal, thousands: decimal === ',' ? '.' : ',', dateOrder };
}

// --------------------------------------------------------------------------------------------
// Export

/** Text of a cell as written to delimited files: formatted for non-General formats, else raw. */
export function cellText(sheet: Sheet, r: number, c: number, opts: CsvOptions = {}): string {
  const cell = sheet.get(r, c);
  const v = cell?.v;
  if (v == null) return '';
  if (v instanceof CellError) return v.code;
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'string') return v;
  const fmt = sheet.styleOf(r, c).numFmt;
  let s = fmt && fmt !== 'General' ? formatValue(v, fmt, { date1904: sheet.book.calc.date1904 }).text : numText(v);
  if (opts.decimal === ',' && (!fmt || fmt === 'General')) s = s.replace('.', ',');
  return s;
}

function numText(n: number): string {
  if (!Number.isFinite(n)) return '#NUM!';
  const s = String(n);
  if (/e/.test(s) && Math.abs(n) < 1e21 && Math.abs(n) >= 1e-7) return n.toFixed(20).replace(/\.?0+$/, '');
  return s;
}

export function sheetRows(sheet: Sheet, range: Range | null | undefined, opts: CsvOptions = {}): string[][] {
  const rg = range || sheet.usedRange(false);
  if (!rg) return [];
  const rows: string[][] = [];
  for (let r = rg.r1; r <= rg.r2; r++) {
    const row: string[] = [];
    for (let c = rg.c1; c <= rg.c2; c++) {
      const t = cellText(sheet, r, c, opts);
      // Text that begins with = would be read back (here and in other spreadsheets) as a formula: it is written with the
      // apostrophe that marks text, which the import below takes off again.
      const cell = sheet.get(r, c);
      row.push(typeof cell?.v === 'string' && cell.f == null && t.startsWith('=') ? "'" + t : t);
    }
    rows.push(row);
  }
  return rows;
}

export function writeCsv(rows: string[][], opts: CsvOptions = {}): string {
  const d = opts.delimiter || ',';
  const q = opts.quote === undefined ? '"' : opts.quote;
  const eol = opts.lineEnding || '\r\n';
  const needs = (s: string) => s.includes(d) || (q && s.includes(q)) || s.includes('\n') || s.includes('\r') || /^\s|\s$/.test(s);
  return rows.map((row) => row.map((f) => (q && needs(f) ? q + f.split(q).join(q + q) + q : f)).join(d)).join(eol) + (rows.length ? eol : '');
}

const CP1252: Record<number, number> = {
  0x20ac: 0x80, 0x201a: 0x82, 0x0192: 0x83, 0x201e: 0x84, 0x2026: 0x85, 0x2020: 0x86, 0x2021: 0x87, 0x02c6: 0x88, 0x2030: 0x89, 0x0160: 0x8a,
  0x2039: 0x8b, 0x0152: 0x8c, 0x017d: 0x8e, 0x2018: 0x91, 0x2019: 0x92, 0x201c: 0x93, 0x201d: 0x94, 0x2022: 0x95, 0x2013: 0x96, 0x2014: 0x97,
  0x02dc: 0x98, 0x2122: 0x99, 0x0161: 0x9a, 0x203a: 0x9b, 0x0153: 0x9c, 0x017e: 0x9e, 0x0178: 0x9f,
};

/** Encodes text in the requested encoding (utf-8, utf-16le/be, windows-1252, iso-8859-1). */
export function encodeText(text: string, encoding = 'utf-8', bom = false): Uint8Array {
  const enc = encoding.toLowerCase().replace(/_/g, '-');
  if (enc === 'utf-16le' || enc === 'utf-16' || enc === 'unicode') {
    const out = new Uint8Array(2 + text.length * 2);
    out[0] = 0xff; out[1] = 0xfe;
    for (let i = 0; i < text.length; i++) { const c = text.charCodeAt(i); out[2 + i * 2] = c & 255; out[3 + i * 2] = c >> 8; }
    return out;
  }
  if (enc === 'utf-16be') {
    const out = new Uint8Array(2 + text.length * 2);
    out[0] = 0xfe; out[1] = 0xff;
    for (let i = 0; i < text.length; i++) { const c = text.charCodeAt(i); out[2 + i * 2] = c >> 8; out[3 + i * 2] = c & 255; }
    return out;
  }
  if (enc === 'windows-1252' || enc === 'cp1252' || enc === 'iso-8859-1' || enc === 'latin1' || enc === 'ascii' || enc === 'us-ascii') {
    const out = new Uint8Array(text.length);
    const ascii = enc === 'ascii' || enc === 'us-ascii';
    for (let i = 0; i < text.length; i++) {
      const c = text.charCodeAt(i);
      out[i] = ascii ? (c < 128 ? c : 63) : c < 256 && !(enc.includes('1252') && c >= 0x80 && c < 0xa0) ? c : CP1252[c] ?? 63;
    }
    return out;
  }
  const body = new TextEncoder().encode(text);
  if (!bom) return body;
  const out = new Uint8Array(body.length + 3);
  out.set([0xef, 0xbb, 0xbf]); out.set(body, 3);
  return out;
}

export function exportCsv(book: Workbook, sheetIndex = 0, opts: CsvOptions = {}, range?: Range): Blob {
  const sheet = book.sheets[Math.max(0, Math.min(sheetIndex, book.sheets.length - 1))];
  const rows = sheetRows(sheet, range, opts);
  const text = writeCsv(rows, opts);
  const bytes = encodeText(text, opts.encoding || 'utf-8', !!opts.bom);
  const type = opts.delimiter === '\t' ? 'text/tab-separated-values' : 'text/csv';
  return new Blob([bytes as BlobPart], { type: `${type};charset=${opts.encoding || 'utf-8'}` });
}
