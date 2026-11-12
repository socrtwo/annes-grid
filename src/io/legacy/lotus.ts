/**
 * Lotus 1-2-3 worksheets, and the Lotus-format files of Symphony and Microsoft Works:
 *
 *   WKS   1-2-3 Release 1A (1983)            BOF 0x0404, 2-byte body
 *   WRK   Symphony 1.0                       BOF 0x0405
 *   WK1   1-2-3 Release 2.x, Symphony 1.1+   BOF 0x0406
 *   WK3   1-2-3 Release 3.x (1989)           BOF 0x1000-0x1002, 26-byte body; several sheets
 *   WK4   1-2-3 Release 4/5 for Windows      BOF 0x1003-0x1004
 *   123   1-2-3 97 and Millennium Edition    BOF 0x1005-0x1006
 *   WKS   Microsoft Works for DOS (Lotus records plus 0x54xx records), Works for Windows 3/4
 *         (BOF record 0x00FF, cells carry a style number)
 *
 * Release 1A/2 records follow Lotus's public "Worksheet File Format" specification: values,
 * labels with their alignment prefixes, formulas (decompiled and translated, see lotusfmla.ts),
 * cell formats, column widths, hidden columns and named ranges. Release 3 and later files are read
 * for values, formulas, sheet names, column widths, hidden columns, named ranges and cell notes.
 *
 * Every record is bounds-checked; in lenient mode a damaged record is skipped and reading continues.
 */
import { Workbook, Sheet } from '../../model/workbook';
import { CellError, type Cell, type CellStyle } from '../../model/types';
import { cellKey, keyRow, keyCol, colName, MAX_COLS, MAX_ROWS } from '../../model/address';
import { lotusLabel } from '../../heritage/lotus';
import { u16, i16, u32, f64, f64hi, f80, f80Special, snum, lnum, cstr, codepage, lmbcs, walkRecords, charsToPx, type Decoder, type LRec } from './bin';
import { lotusFormula, type Dialect } from './lotusfmla';

export interface LegacyOptions {
  /** Keep going after damaged records (used by the recovery pipeline). */
  lenient?: boolean;
  /** Code page for 8-bit text; defaults to the program's usual one. */
  codepage?: number;
}

import { lotusVariant, type LotusVariant } from './signatures';
export { lotusVariant, LOTUS_VARIANT_LABEL, type LotusVariant } from './signatures';

// ------------------------------------------------------------------------------------------
// Cell formats (Appendix A of the specification)

/** Number format for a 1-2-3 cell format byte; null for General, undefined for "use the sheet default". */
export function lotusNumFmt(fmt: number): string | null | undefined {
  const type = (fmt >> 4) & 7, d = fmt & 15;
  const dec = d ? '.' + '0'.repeat(d) : '';
  switch (type) {
    case 0: return `0${dec}`;
    case 1: return `0${dec}E+00`;
    case 2: return `$#,##0${dec}_);($#,##0${dec})`;
    case 3: return `0${dec}%`;
    case 4: return `#,##0${dec}_);(#,##0${dec})`;
    case 7:
      switch (d) {
        case 0: return null;           // +/- bar graph
        case 1: return null;           // General
        case 2: return 'dd-mmm-yy';     // D1
        case 3: return 'dd-mmm';        // D2
        case 4: return 'mmm-yy';        // D3
        case 5: return null;           // Text (shows the formula)
        case 6: return ';;;';          // Hidden
        case 7: return 'h:mm:ss AM/PM'; // D6
        case 8: return 'h:mm AM/PM';    // D7
        case 9: return 'mm/dd/yy';      // D4 (Long International date)
        case 10: return 'mm/dd';        // D5 (Short International date)
        case 11: return 'hh:mm:ss';     // D8
        case 12: return 'hh:mm';        // D9
        case 15: return undefined;     // Default
      }
      return null;
  }
  return null;
}

/**
 * Number format of a Works for Windows cell style: the low five bits are the kind (0 fixed,
 * 1 scientific, 2 currency, 3 percent, 4 comma, 5 general, 6 date, 10 leading zeros), the top
 * three bits the decimals.
 */
export function worksWinNumFmt(b: number): string | null {
  const kind = b & 0x1f, d = b >> 5;
  const dec = d ? '.' + '0'.repeat(d) : '';
  switch (kind) {
    case 0: return `0${dec}`;
    case 1: return `0${dec}E+00`;
    case 2: return `$#,##0${dec}_);($#,##0${dec})`;
    case 3: return `0${dec}%`;
    case 4: return `#,##0${dec}_);(#,##0${dec})`;
    case 6: return 'm/d/yyyy';
    case 10: return '0'.repeat(d + 1);
  }
  return null;
}

const ALIGN: Record<string, CellStyle['hAlign']> = { left: 'left', right: 'right', center: 'center', fill: 'fill' };

// ------------------------------------------------------------------------------------------
// Reader

interface PendingString { sheet: Sheet; r: number; c: number }

class LotusReader {
  book = new Workbook(false);
  sheets: Sheet[] = [];
  sheetNames: string[] = [];
  warnings: string[] = [];
  styleCache = new Map<string, number>();
  defaultFmt = 0xff;
  untranslated = 0;
  lastCell: { sheet: Sheet; r: number; c: number } | null = null;
  pendingString: PendingString | null = null;
  constructor(public variant: LotusVariant, public dec: Decoder, public lenient: boolean) {}

  warn(m: string): void { if (!this.warnings.includes(m) && this.warnings.length < 100) this.warnings.push(m); }

  sheet(i: number): Sheet {
    while (this.sheets.length <= i) {
      const k = this.sheets.length;
      const nm = this.sheetNames[k] || (this.variant === 'wk3' || this.variant === 'wk4' || this.variant === '123' ? colName(k) : 'Sheet1');
      const sh = new Sheet(this.book, nm);
      sh.defaultColW = charsToPx(9);
      this.sheets.push(sh);
    }
    return this.sheets[i];
  }

  style(st: CellStyle): number {
    const key = JSON.stringify(st);
    let id = this.styleCache.get(key);
    if (id == null) { id = Object.keys(st).length ? this.book.styleId(st) : 0; this.styleCache.set(key, id); }
    return id;
  }

  fmtStyle(fmt: number | null, align?: CellStyle['hAlign'], extra?: CellStyle): number {
    const st: CellStyle = { ...(extra || {}) };
    if (fmt != null) {
      let nf = lotusNumFmt(fmt);
      if (nf === undefined) nf = lotusNumFmt(this.defaultFmt) ?? null;
      if (nf) st.numFmt = nf;
    }
    if (align) st.hAlign = align;
    return this.style(st);
  }

  put(sheet: Sheet, r: number, c: number, cell: Cell): void {
    if (r < 0 || c < 0 || r >= MAX_ROWS || c >= MAX_COLS) return;
    if (!cell.s) delete cell.s;
    sheet.setComputed(r, c, cell);
    this.lastCell = { sheet, r, c };
  }

  label(sheet: Sheet, r: number, c: number, raw: string, fmt: number | null, extra?: CellStyle): void {
    const lab = lotusLabel(raw);
    let text = raw, align: CellStyle['hAlign'] | undefined;
    if (lab) { text = lab.text; align = lab.align === 'left' ? undefined : ALIGN[lab.align]; }
    else if (raw.length === 1 && "'\"^\\|".includes(raw)) text = '';
    else if (raw[0] === '|') text = raw.slice(1); // non-printing row marker
    this.put(sheet, r, c, { v: text, s: this.fmtStyle(fmt, align, extra) });
  }

  formula(sheet: Sheet, si: number, r: number, c: number, code: Uint8Array, dialect: Dialect, cached: Cell['v'], fmt: number | null, extra?: CellStyle): Cell {
    const res = lotusFormula(code, { dialect, row: r, col: c, sheet: si, dec: this.dec }, this.sheetNames.length ? this.sheetNames : this.sheets.map((s) => s.name));
    const cell: Cell = { v: cached, s: this.fmtStyle(fmt, undefined, extra) };
    if (res.f) cell.f = res.f;
    else {
      this.untranslated++;
      const where = `${this.sheets.length > 1 ? sheet.name + '!' : ''}${colName(c)}${r + 1}`;
      this.warn(`Formula in ${where} could not be translated (${res.error || 'unsupported'}); its last value is kept.`);
      sheet.comments.set(cellKey(r, c), { author: 'Import', text: res.lotus ? `1-2-3 formula, not translated (${res.error}): ${res.lotus}` : `The 1-2-3 formula here could not be read (${res.error}); the value is the one 1-2-3 last calculated.`, time: Date.now(), note: true });
    }
    this.put(sheet, r, c, cell);
    return cell;
  }
}

const NA = () => new CellError('#N/A');
const ERR = () => new CellError('#VALUE!');

/** Cached value of a 1A/2.x formula or number: NA and ERR are NaN patterns. */
function wk1Number(d: Uint8Array, o: number): Cell['v'] {
  const v = f64(d, o);
  if (Number.isFinite(v)) return v;
  const hi = u32(d, o + 4), lo = u32(d, o);
  if ((hi & 0x7ff00000) === 0x7ff00000 && (hi & 0xfffff) === 0 && lo === 0) return hi & 0x80000000 ? NA() : ERR();
  return null; // string result (Symphony) or unknown
}

function f80Value(d: Uint8Array, o: number): Cell['v'] {
  if (!f80Special(d, o)) return f80(d, o);
  const se = u16(d, o + 8), hi = u32(d, o + 4), lo = u32(d, o);
  if (hi === 0xc0000000 && lo === 0 && se & 0x8000) return NA();
  return null;
}

export function importLotus(data: ArrayBuffer | Uint8Array, name = '', warnings: string[] = [], opts: LegacyOptions = {}): Workbook {
  const b = data instanceof Uint8Array ? data : new Uint8Array(data);
  const variant = lotusVariant(b);
  if (!variant) throw new Error('This is not a Lotus 1-2-3 worksheet.');
  const r3 = variant === 'wk3' || variant === 'wk4' || variant === '123';
  const dec = opts.codepage ? codepage(opts.codepage) : r3 ? lmbcs : codepage(variant === 'works-win' ? 1252 : 437);
  const rd = new LotusReader(variant, dec, !!opts.lenient);
  const { recs, truncated, skipped } = walkRecords(b, {
    lenient: !!opts.lenient, stop: (op) => op === 0x0001,
    // Release 3+ files carry pictures and charts in long records; resynchronize after a damaged one.
    plausible: r3 ? (op, len) => op < 0x0c00 && len < 0x4000 : undefined,
  });
  if (skipped) rd.warn('Part of the file (pictures or charts) could not be read.');
  if (truncated) {
    if (!opts.lenient) throw new Error('The worksheet file is truncated.');
    rd.warn('The file ends early; data after the break is lost.');
  }
  const guard = (rec: LRec, fn: () => void) => {
    try { fn(); } catch (e: any) {
      if (!opts.lenient) throw e;
      rd.warn(`A damaged record (type 0x${rec.op.toString(16)}) was skipped.`);
    }
  };
  if (variant === 'wk3' || variant === 'wk4' || variant === '123') readWk3(rd, recs, guard);
  else readWk1(rd, recs, guard);

  const book = rd.book;
  if (!rd.sheets.length) rd.sheet(0);
  // A formula whose cached result is "not a number" with no string after it holds an error.
  for (const sh of rd.sheets) for (const c of sh.cells.values()) if (c.f != null && c.v == null) c.v = ERR();
  const seen = new Set<string>();
  for (const sh of rd.sheets) {
    let nm = sh.name.replace(/[\\/?*[\]:]/g, '_').slice(0, 31) || 'Sheet';
    let k = 2;
    while (seen.has(nm.toLowerCase())) nm = `${sh.name.slice(0, 27)} (${k++})`;
    seen.add(nm.toLowerCase());
    sh.name = nm;
    book.sheets.push(sh);
  }
  book.activeSheet = 0;
  if (name) book.props.title = name.replace(/\.[^.]+$/, '');
  if (rd.untranslated) rd.warn(`${rd.untranslated} formula${rd.untranslated > 1 ? 's' : ''} could not be translated and keep their last values (see the cell notes).`);
  warnings.push(...rd.warnings);
  return book;
}

// ------------------------------------------------------------------------------------------
// Release 1A / 2.x / Symphony / Works

function readWk1(rd: LotusReader, recs: LRec[], guard: (r: LRec, fn: () => void) => void): void {
  const sh = rd.sheet(0);
  const works = rd.variant === 'works-dos';
  const win = rd.variant === 'works-win';
  // Works for Windows: fonts (0x5456) and cell styles (0x545A) referenced by number.
  const wfonts: { flags: number; name: string }[] = [];
  const wstyles: Uint8Array[] = [];
  const wstyle = (i: number): CellStyle => {
    const st: CellStyle = {};
    const s = wstyles[i];
    if (!s) return st;
    const f = wfonts[s[4]];
    if (f) {
      if (f.flags & 1) st.bold = true;
      if (f.flags & 2) st.italic = true;
      if (f.flags & 4) st.underline = 'single';
    }
    const nf = worksWinNumFmt(s[0]);
    if (nf) st.numFmt = nf;
    const al = s[1] & 0x0c;
    if (al === 0x04) st.hAlign = 'left'; else if (al === 0x08) st.hAlign = 'center'; else if (al === 0x0c) st.hAlign = 'right';
    return st;
  };
  for (const rec of recs) {
    const d = rec.data;
    guard(rec, () => {
      switch (rec.op) {
        case 0x07: if (d.length >= 8) { rd.defaultFmt = d[4]; const w = u16(d, 6); if (w > 0 && w < 256) sh.defaultColW = charsToPx(w); } break;
        case 0x08: if (d.length >= 3) { const c = u16(d, 0), w = d[2]; if (c < MAX_COLS && w) sh.colW.set(c, charsToPx(w)); } break;
        case 0x0b: if (d.length >= 24) addName(rd, cstr(d, 0, rd.dec, 16), 0, u16(d, 16), u16(d, 18), u16(d, 20), u16(d, 22)); break;
        case 0x64: // HIDCOL: 256-bit map of hidden columns
          for (let i = 0; i < Math.min(32, d.length); i++) for (let k = 0; k < 8; k++) if (d[i] & (1 << k)) sh.hiddenCols.add(i * 8 + k);
          break;
        case 0x0c: if (!win && d.length >= 5) { const s = rd.fmtStyle(d[0]); if (s) rd.put(sh, u16(d, 3), u16(d, 1), { s }); } break;
        case 0x0d: if (!win && d.length >= 7) rd.put(sh, u16(d, 3), u16(d, 1), { v: i16(d, 5), s: rd.fmtStyle(d[0]) }); break;
        case 0x0e:
          if (win) { if (d.length >= 14) rd.put(sh, u16(d, 2), u16(d, 0), { v: wk1Number(d, 6), s: rd.style(wstyle(u16(d, 4))) }); }
          else if (d.length >= 13) rd.put(sh, u16(d, 3), u16(d, 1), { v: wk1Number(d, 5), s: rd.fmtStyle(d[0]) });
          break;
        case 0x0f:
          if (win) { if (d.length >= 7) rd.label(sh, u16(d, 2), u16(d, 0), cstr(d, 6, rd.dec), null, wstyle(u16(d, 4))); }
          else if (d.length >= 6) {
            // Quattro for DOS writes the label prefix, then a length byte, without a terminating zero.
            const quattro = rd.variant === 'wq1' || rd.variant === 'wq2';
            const raw = quattro && d.length >= 7 && d[6] < 0x20 && 7 + d[6] <= d.length && d[5] !== 0
              ? String.fromCharCode(d[5]) + rd.dec(d.subarray(7, 7 + d[6])) : cstr(d, 5, rd.dec);
            rd.label(sh, u16(d, 3), u16(d, 1), raw, d[0]);
          }
          break;
        case 0x10: {
          if (win) {
            if (d.length < 16) break;
            const c = u16(d, 0), r = u16(d, 2), size = u16(d, 14);
            const cell = rd.formula(sh, 0, r, c, d.subarray(16, 16 + size), 'wk1', wk1Number(d, 6), null, wstyle(u16(d, 4)));
            rd.pendingString = cell.v == null ? { sheet: sh, r, c } : null;
          } else {
            if (d.length < 15) break;
            const c = u16(d, 1), r = u16(d, 3), size = u16(d, 13);
            const cell = rd.formula(sh, 0, r, c, d.subarray(15, 15 + size), 'wk1', wk1Number(d, 5), d[0]);
            rd.pendingString = cell.v == null ? { sheet: sh, r, c } : null;
          }
          break;
        }
        case 0x33: { // STRING: value of the preceding string formula
          const p = rd.pendingString;
          if (p && d.length >= 5) {
            const cur = p.sheet.get(p.r, p.c);
            if (cur) cur.v = cstr(d, 5, rd.dec);
          }
          rd.pendingString = null;
          break;
        }
        case 0x5402: { // Works for DOS: font attributes of the preceding cell
          const lc = rd.lastCell;
          if (works && lc && d.length >= 2) {
            const cur = lc.sheet.get(lc.r, lc.c);
            if (cur) {
              const base = { ...(rd.book.styles[cur.s || 0] || {}) };
              if (d[1] & 0x20) base.bold = true;
              if (d[1] & 0x40) base.underline = 'single';
              if (d[0] & 0x10) base.italic = true;
              cur.s = rd.style(base);
              if (!cur.s) delete cur.s;
            }
          }
          break;
        }
        case 0x5456: if (win && d.length > 4) wfonts.push({ flags: u16(d, 0), name: cstr(d, 2, rd.dec, 32) }); break;
        case 0x545a: if (win) wstyles.push(d); break;
        case 0x545b: if (win && d.length >= 10) rd.put(sh, u16(d, 2), u16(d, 0), { v: f64hi(d, 6), s: rd.style(wstyle(u16(d, 4))) }); break;
      }
    });
  }
}

function addName(rd: LotusReader, nm: string, sheet: number, c1: number, r1: number, c2: number, r2: number, sheet2 = sheet): void {
  let name = nm.trim();
  // 1-2-3 allows spaces and some punctuation in range names; names here are identifiers.
  if (/^[A-Za-z_\\][^\x00-\x1f]*$/.test(name) && /[ \-&!#$%]/.test(name)) name = name.replace(/[ \-&!#$%]/g, '_');
  if (!name || !/^[A-Za-z_\\][A-Za-z0-9_.]*$/.test(name) || /^[A-Za-z]{1,3}\d+$/.test(name)) {
    if (name) rd.warn(`The range name "${name}" is not a valid name here and was left out.`);
    return;
  }
  if (rd.book.names.some((n) => n.name.toLowerCase() === name.toLowerCase())) return;
  if (c1 > 16383 || c2 > 16383 || r1 > 1048575 || r2 > 1048575) return;
  const sn = (i: number) => rd.sheetNames[i] || rd.sheets[i]?.name || (rd.variant === 'wk3' || rd.variant === 'wk4' || rd.variant === '123' ? colName(i) : 'Sheet1');
  const q = (s: string) => (/^[A-Za-z_][A-Za-z0-9_]*$/.test(s) && !/^[A-Za-z]{1,3}\d+$/.test(s) ? s : `'${s.replace(/'/g, "''")}'`);
  const a = `$${colName(c1)}$${r1 + 1}`, b = `$${colName(c2)}$${r2 + 1}`;
  const prefix = sheet === sheet2 ? q(sn(sheet)) : `${q(sn(sheet))}:${q(sn(sheet2))}`;
  rd.book.names.push({ name: name.replace(/^\\/, '_'), ref: `${prefix}!${a === b ? a : `${a}:${b}`}`, sheet: null });
}

// ------------------------------------------------------------------------------------------
// Release 3 and later

/** Release 3 cell records: the fewest bytes each needs (row, sheet and column, then the value). */
const WK3_CELL_MIN: Record<number, number> = { 0x14: 4, 0x15: 4, 0x16: 4, 0x17: 14, 0x18: 6, 0x19: 4, 0x25: 8, 0x27: 12, 0x28: 4 };

function readWk3(rd: LotusReader, recs: LRec[], guard: (r: LRec, fn: () => void) => void): void {
  // Pass 1: sheet names, so formulas can name other sheets.
  let maxSheet = 0;
  for (const rec of recs) {
    const d = rec.data;
    if ((rec.op === 0x1b || rec.op === 0x23) && d.length > 4 && u16(d, 0) === 14000) {
      const idx = u16(d, 2);
      if (idx < 256) rd.sheetNames[idx] = cstr(d, 4, rd.dec, 64);
    }
    if (rec.op >= 0x14 && rec.op <= 0x1a && d.length >= 4) maxSheet = Math.max(maxSheet, d[2]);
    if ((rec.op === 0x25 || rec.op === 0x26 || rec.op === 0x27 || rec.op === 0x28) && d.length >= 4) maxSheet = Math.max(maxSheet, d[2]);
  }
  for (let i = 0; i <= maxSheet; i++) if (!rd.sheetNames[i]) rd.sheetNames[i] = colName(i);
  for (let i = 0; i <= maxSheet; i++) rd.sheet(i);
  // Release 3 cell formats (record 0x13), one record per formatted row: kind 0 lists 4-byte cell
  // attributes (the first byte is the 1-2-3 format byte; bit 7 of the fourth announces a repeat
  // count byte), kind 1 holds the sheet default, kind 2 repeats another row.
  const rowFmt = new Map<string, number[]>();
  const sheetDefault = new Map<number, number>();
  const at = (d: Uint8Array) => ({ r: u16(d, 0), s: d[2], c: d[3] });
  for (const rec of recs) {
    const d = rec.data;
    guard(rec, () => {
      // A cell record too short to hold its address (and, for numbers, its value) is damaged: it is skipped.
      const least = WK3_CELL_MIN[rec.op];
      if (least && d.length < least) { rd.warn(`A damaged record (type 0x${rec.op.toString(16)}) was skipped.`); return; }
      switch (rec.op) {
        case 0x07: { // column widths: sheet, window, 2 bytes, then (column, width) pairs
          if (d.length < 4 || d[1] !== 0) break;
          const sh = rd.sheet(d[0]);
          for (let p = 4; p + 1 < d.length; p += 2) if (d[p + 1]) sh.colW.set(d[p], charsToPx(d[p + 1]));
          break;
        }
        case 0x08: { // hidden columns
          if (d.length < 4 || d[1] !== 0) break;
          const sh = rd.sheet(d[0]);
          for (let p = 4; p < d.length; p++) sh.hiddenCols.add(d[p]);
          break;
        }
        case 0x09: { // named ranges (0x0A holds the program's internal ones)
          if (d.length < 26) break;
          const nm = cstr(d, 2, rd.dec, 16);
          addName(rd, nm, d[20], d[21], u16(d, 18), d[25], u16(d, 22), d[24]);
          break;
        }
        case 0x13: {
          if (d.length < 5) break;
          const sh = d[0], kind = d[1], row = u16(d, 2);
          if (kind === 1) { sheetDefault.set(sh, d[4]); break; }
          if (kind === 2 && d.length >= 8) { const src = rowFmt.get(`${d[4]},${u16(d, 6)}`); if (src) rowFmt.set(`${sh},${row}`, src); break; }
          if (kind !== 0) break;
          const cols: number[] = [];
          for (let p = 4; p + 4 <= d.length && cols.length < 256;) {
            const f = d[p], rep = d[p + 3] & 0x80 ? d[p + 4] : 0;
            for (let k = 0; k <= rep; k++) cols.push(f);
            p += rep || d[p + 3] & 0x80 ? 5 : 4;
          }
          rowFmt.set(`${sh},${row}`, cols);
          break;
        }
        case 0x14: { const a = at(d); rd.put(rd.sheet(a.s), a.r, a.c, { v: ERR() }); break; }
        case 0x15: { const a = at(d); rd.put(rd.sheet(a.s), a.r, a.c, { v: NA() }); break; }
        case 0x16: { const a = at(d); rd.label(rd.sheet(a.s), a.r, a.c, cstr(d, 4, rd.dec), null); break; }
        case 0x17: { const a = at(d); rd.put(rd.sheet(a.s), a.r, a.c, { v: f80Value(d, 4) }); break; }
        case 0x18: { const a = at(d); rd.put(rd.sheet(a.s), a.r, a.c, { v: snum(u16(d, 4)) }); break; }
        case 0x25: { const a = at(d); rd.put(rd.sheet(a.s), a.r, a.c, { v: lnum(u32(d, 4)) }); break; }
        case 0x27: { const a = at(d); rd.put(rd.sheet(a.s), a.r, a.c, { v: wk1Number(d, 4) }); break; }
        case 0x19: case 0x28: {
          const a = at(d);
          const sh = rd.sheet(a.s);
          const wk3 = rec.op === 0x19;
          const cached = wk3 ? f80Value(d, 4) : wk1Number(d, 4);
          const cell = rd.formula(sh, a.s, a.r, a.c, d.subarray(wk3 ? 14 : 12), wk3 ? 'wk3' : 'wk4', cached, null);
          rd.pendingString = cell.v == null ? { sheet: sh, r: a.r, c: a.c } : null;
          break;
        }
        case 0x1a: {
          const a = at(d);
          const cur = rd.sheet(a.s).get(a.r, a.c);
          if (cur) cur.v = cstr(d, 4, rd.dec);
          rd.pendingString = null;
          break;
        }
        case 0x26: { // cell note
          const a = at(d);
          if (d.length < 6) break;
          const text = cstr(d, 5, rd.dec).replace(/\r\n?/g, '\n');
          if (text) rd.sheet(a.s).comments.set(cellKey(a.r, a.c), { author: '', text, time: Date.now(), note: true });
          break;
        }
      }
    });
  }
  // Apply the formats to the cells read.
  if (rowFmt.size || sheetDefault.size) {
    rd.sheets.forEach((sh, si) => {
      const dflt = sheetDefault.get(si);
      if (dflt != null) rd.defaultFmt = dflt;
      for (const [k, cell] of sh.cells) {
        const r = keyRow(k), c = keyCol(k);
        const f = rowFmt.get(`${si},${r}`)?.[c];
        if (f == null || cell.v == null || typeof cell.v === 'string') continue;
        const nf = lotusNumFmt(f);
        const useDefault = nf === undefined;
        const fmt = useDefault ? (dflt != null ? lotusNumFmt(dflt) ?? null : null) : nf;
        if (!fmt) continue;
        const st = { ...(rd.book.styles[cell.s || 0] || {}), numFmt: fmt };
        cell.s = rd.style(st);
      }
    });
  }
}
