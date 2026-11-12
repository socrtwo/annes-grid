/**
 * dBASE, FoxBASE, FoxPro, Visual FoxPro and Clipper tables (.dbf), with memo files
 * (.dbt for dBASE, .fpt for FoxPro) when they are supplied.
 *
 * The table becomes one sheet: field names in the first row (frozen and filterable), one record
 * per row. Character, numeric, float, date, logical, memo, integer, currency, double, datetime,
 * timestamp and autoincrement fields are converted to values with matching number formats.
 * Deleted records are left out. Text is decoded with the table's language driver (code page), or
 * with the more plausible of code page 437 and Windows-1252 when the table does not say.
 */
import { Workbook, Sheet } from '../../model/workbook';
import { MAX_COLS, MAX_ROWS } from '../../model/address';
import type { CellStyle } from '../../model/types';
import { u16, u32, i32, u32be, codepage, charsToPx, type Decoder } from './bin';
import type { LegacyOptions } from './lotus';

export interface DbfOptions extends LegacyOptions {
  /** Memo file (.dbt or .fpt) belonging to the table. */
  memo?: Uint8Array;
}

interface Field { name: string; raw: Uint8Array; type: string; len: number; dec: number; off: number; nullBit: number; varBit: number }

const LDID: Record<number, number> = {
  0x01: 437, 0x02: 850, 0x03: 1252, 0x04: 10000, 0x08: 865, 0x09: 437, 0x0a: 850, 0x0b: 437, 0x0d: 437, 0x0e: 850, 0x0f: 437,
  0x10: 850, 0x11: 437, 0x12: 850, 0x13: 932, 0x14: 850, 0x15: 437, 0x16: 850, 0x17: 865, 0x18: 437, 0x19: 437, 0x1a: 850,
  0x1b: 437, 0x1c: 863, 0x1d: 850, 0x1f: 852, 0x22: 852, 0x23: 852, 0x24: 860, 0x25: 850, 0x26: 866, 0x37: 850, 0x40: 852,
  0x4d: 936, 0x4e: 949, 0x4f: 950, 0x50: 874, 0x57: 1252, 0x58: 1252, 0x59: 1252, 0x64: 852, 0x65: 866, 0x66: 865, 0x67: 861,
  0x6a: 737, 0x6b: 857, 0x6c: 863, 0x78: 950, 0x79: 949, 0x7a: 936, 0x7b: 932, 0x7c: 874, 0x7d: 1255, 0x7e: 1256, 0x86: 737,
  0x87: 852, 0x88: 857, 0xc8: 1250, 0xc9: 1251, 0xca: 1254, 0xcb: 1253, 0xcc: 1257,
};

export const DBF_VERSIONS: Record<number, string> = {
  0x02: 'FoxBASE / dBASE II', 0x03: 'dBASE III / FoxPro', 0x04: 'dBASE IV', 0x05: 'dBASE 5', 0x30: 'Visual FoxPro', 0x31: 'Visual FoxPro (autoincrement)',
  0x32: 'Visual FoxPro (varchar)', 0x43: 'dBASE IV SQL table', 0x63: 'dBASE IV SQL system', 0x83: 'dBASE III with memo', 0x8b: 'dBASE IV with memo',
  0x8e: 'dBASE IV with SQL table', 0xcb: 'dBASE IV SQL table with memo', 0xf5: 'FoxPro 2 with memo', 0xfb: 'FoxBASE',
};

/** How much a decoding reads as letters rather than box-drawing characters and other symbols. */
function letterScore(t: string): number {
  let score = 0;
  for (const ch of t) {
    const c = ch.charCodeAt(0);
    if (c < 0x80) continue;
    if (/\p{L}/u.test(ch)) score += 1;
    else score -= 3;
    if (c >= 0x2500 && c <= 0x25ff) score -= 3;
    if (c >= 0x370 && c <= 0x3ff) score -= 1; // Greek letters in code page 437 are rarely the intent
  }
  return score;
}

/** Picks the code page that reads as the most letters; a declared page wins unless it reads as noise. */
function chooseDecoder(sample: Uint8Array, declared?: number): Decoder {
  let hi = 0;
  for (const x of sample) if (x >= 0x80) hi++;
  if (declared && !hi) return codepage(declared);
  if (!hi) return codepage(1252);
  const pages = [1252, 437, 850, 1251, 866, 1250, 852];
  const script = (cp: number) => ([1251, 866, 855].includes(cp) ? 'cyr' : [1250, 852].includes(cp) ? 'ce' : 'latin');
  let best = declared ?? 1252, bestScore = -Infinity;
  // The declared page gets the same benefit as the others of its script, or a short sample always loses to them.
  const boost = (sc: number) => (sc > 0 ? sc * 1.1 + 1 : sc);
  const declaredScore = declared ? boost(letterScore(codepage(declared)(sample))) : -Infinity;
  for (const cp of pages) {
    let sc = letterScore(codepage(cp)(sample));
    // A table that names a code page was written for that script, even when the page itself is wrong.
    if (declared && script(cp) === script(declared)) sc = sc > 0 ? sc * 1.1 + 1 : sc;
    if (sc > bestScore) { bestScore = sc; best = cp; }
  }
  if (declared && declaredScore >= bestScore * 0.8 && declaredScore > 0) return codepage(declared);
  return codepage(best);
}

export function importDbf(data: ArrayBuffer | Uint8Array, name = '', warnings: string[] = [], opts: DbfOptions = {}): Workbook {
  const b = data instanceof Uint8Array ? data : new Uint8Array(data);
  if (b.length < 32) throw new Error('The table file is too short.');
  const ver = b[0];
  const vfp = ver === 0x30 || ver === 0x31 || ver === 0x32;
  const dbase7 = (ver & 0x07) === 0x04 && b.length > 68 && b[32 + 32] !== 0x0d && u16(b, 8) > 32 + 48;
  let nrec: number, hlen: number, rlen: number;
  const fields: Field[] = [];
  let declared: number | undefined;
  if (ver === 0x02) {
    // dBASE II: record count, date, record length, then up to 32 descriptors of 16 bytes.
    nrec = u16(b, 1); rlen = u16(b, 6); hlen = 521;
    let off = 1;
    for (let p = 8; p + 16 <= 520 && b[p] !== 0x0d; p += 16) {
      const f: Field = { name: '', raw: nameBytes(b, p, 11), type: String.fromCharCode(b[p + 11]), len: b[p + 12], dec: b[p + 15], off, nullBit: -1, varBit: -1 };
      off += f.len; fields.push(f);
    }
  } else {
    nrec = u32(b, 4); hlen = u16(b, 8); rlen = u16(b, 10);
    declared = LDID[b[29]];
    const size = dbase7 ? 48 : 32;
    let off = 1;
    for (let p = dbase7 ? 68 : 32; p + size <= Math.min(hlen, b.length) && b[p] !== 0x0d; p += size) {
      const tp = dbase7 ? 32 : 11;
      // dBASE III-5 and FoxPro: type at 11, length at 16, decimals at 17; dBASE 7: 32, 33, 34.
      const lo = dbase7 ? 33 : 16;
      const f: Field = { name: '', raw: nameBytes(b, p, tp), type: String.fromCharCode(b[p + tp]), len: b[p + lo], dec: b[p + lo + 1], off, nullBit: -1, varBit: -1 };
      // Clipper and FoxPro: a character field longer than 255 keeps the high byte of its length in the decimals byte.
      if (f.type === 'C' && !dbase7) { f.len |= f.dec << 8; f.dec = 0; }
      if (vfp && u32(b, p + 12) > 0) f.off = u32(b, p + 12);
      off = Math.max(off, f.off + f.len);
      fields.push(f);
    }
    // Visual FoxPro: nullable and variable-length fields have bits in the _NullFlags field.
    if (vfp) {
      let bit = 0;
      for (const f of fields) {
        const flags = b[32 + fields.indexOf(f) * 32 + 18];
        if (f.type === '0') continue;
        if (flags & 0x02) f.nullBit = bit++;
        if (f.type === 'V' || f.type === 'Q') f.varBit = bit++;
      }
    }
  }
  if (!fields.length) throw new Error('The table has no fields.');
  if (rlen < 1) throw new Error('The table header is damaged (record length 0).');
  // A record cannot be shorter than its fields: a damaged length would otherwise read one overlapping "record" per byte.
  rlen = Math.max(rlen, ...fields.map((f) => f.off + f.len));
  const available = Math.max(0, Math.floor((b.length - hlen) / rlen));
  if (nrec > available) {
    if (nrec - available > 1) warnings.push(`The table says it has ${nrec} records but the file holds ${available}; the rest is lost.`);
    nrec = available;
  } else if (nrec < available) {
    // A header that was not updated after records were added: read on while records look valid.
    let extra = 0;
    for (let k = nrec; k < available; k++) { const m = b[hlen + k * rlen]; if (m !== 0x20 && m !== 0x2a) break; extra++; }
    if (extra) { warnings.push(`The table header counts ${nrec} records, but ${extra} more follow; they were read too.`); nrec += extra; }
  }
  const sample = b.subarray(hlen, Math.min(b.length, hlen + rlen * 200));
  const dec = chooseDecoder(sample, declared ?? (vfp ? 1252 : undefined));
  const nameDec = chooseDecoder(concat(fields.map((f) => f.raw)), declared);
  for (const f of fields) f.name = nameDec(f.raw).trim();
  const nullField = fields.find((f) => f.type === '0');
  const visible = fields.filter((f) => f.type !== '0');

  const book = new Workbook(false);
  const sh = new Sheet(book, (name.replace(/^.*[\\/]/, '').replace(/\.[^.]+$/, '') || 'Table').slice(0, 31));
  book.sheets.push(sh);
  const bold = book.styleId({ bold: true, bb: { style: 'thin' } });
  const styleCache = new Map<string, number>();
  const sid = (st: CellStyle) => { const k = JSON.stringify(st); let id = styleCache.get(k); if (id == null) { id = book.styleId(st); styleCache.set(k, id); } return id; };
  visible.forEach((f, c) => {
    if (c >= MAX_COLS) return;
    sh.setComputed(0, c, { v: f.name, s: bold });
    const w = Math.max(f.name.length, f.type === 'D' ? 10 : f.type === 'T' || f.type === '@' ? 19 : f.type === 'M' ? 30 : Math.min(f.len, 60)) + 1;
    sh.colW.set(c, charsToPx(w));
  });
  const memo = opts.memo ? memoReader(opts.memo, dec) : null;
  let memoMissing = false, deleted = 0, row = 1;
  for (let k = 0; k < nrec && row < MAX_ROWS; k++) {
    const base = hlen + k * rlen;
    if (b[base] === 0x2a) { deleted++; continue; }
    if (b[base] === 0x1a) break;
    const nulls = nullField ? b.subarray(base + nullField.off, base + nullField.off + nullField.len) : null;
    const isNull = (bit: number) => !!nulls && bit >= 0 && !!(nulls[bit >> 3] & (1 << (bit & 7)));
    visible.forEach((f, c) => {
      if (c >= MAX_COLS) return;
      if (isNull(f.nullBit)) return;
      const raw = b.subarray(base + f.off, base + f.off + f.len);
      const v = fieldValue(f, raw, dec, memo, () => { memoMissing = true; }, isNull(f.varBit), vfp);
      if (v == null || v.v === '' || v.v == null) return;
      const cell: any = { v: v.v };
      if (v.fmt) cell.s = sid({ numFmt: v.fmt });
      sh.setComputed(row, c, cell);
    });
    row++;
  }
  sh.freeze = { rows: 1, cols: 0 };
  if (row > 1 && visible.length) sh.filter = { range: { r1: 0, c1: 0, r2: row - 1, c2: Math.min(visible.length, MAX_COLS) - 1 }, columns: {} };
  if (deleted) warnings.push(`${deleted} record${deleted > 1 ? 's' : ''} marked as deleted ${deleted > 1 ? 'were' : 'was'} left out.`);
  if (memoMissing) warnings.push('The table has memo fields, but its memo file (.dbt or .fpt) was not opened with it; memo text is missing.');
  book.props.title = sh.name;
  book.activeSheet = 0;
  return book;
}

function concat(parts: Uint8Array[]): Uint8Array {
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length + 1, 0));
  let o = 0;
  for (const p of parts) { out.set(p, o); o += p.length + 1; out[o - 1] = 0x20; }
  return out;
}

function nameBytes(b: Uint8Array, o: number, n: number): Uint8Array {
  let e = o;
  while (e < o + n && e < b.length && b[e]) e++;
  return b.subarray(o, e);
}

/** dBASE dates count days from 1 Jan 4713 BC (Julian day numbers); the sheet counts from 30 Dec 1899. */
const JULIAN_1899_12_30 = 2415019;
function serialFromYmd(y: number, m: number, d: number): number { return (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000; }

function fieldValue(f: Field, raw: Uint8Array, dec: Decoder, memo: ((block: number) => string | null) | null, onMissingMemo: () => void, varFlag: boolean, vfp: boolean): { v: any; fmt?: string } | null {
  const text = () => dec(raw);
  switch (f.type) {
    case 'C': {
      let r = raw;
      if (varFlag) { const n = raw[raw.length - 1]; if (n < raw.length) r = raw.subarray(0, n); }
      return { v: dec(r).replace(/[\s\0]+$/, '') };
    }
    case 'V': { const n = varFlag ? raw[raw.length - 1] : raw.length; return { v: dec(raw.subarray(0, Math.min(n, raw.length))).replace(/[\s\0]+$/, '') }; }
    case 'N': case 'F': {
      const t = text().replace(/\0/g, '').trim();
      if (!t || /^[*?.]+$/.test(t)) return null;
      const n = Number(t.replace(',', '.'));
      if (!Number.isFinite(n)) return { v: t };
      return { v: n, fmt: f.dec ? `0.${'0'.repeat(Math.min(f.dec, 15))}` : undefined };
    }
    case 'D': {
      const t = text().trim();
      const m = /^(\d{4})(\d{2})(\d{2})$/.exec(t);
      if (!m) return t ? { v: t } : null;
      const y = +m[1], mo = +m[2], d = +m[3];
      if (!y || !mo || !d) return null;
      return { v: serialFromYmd(y, mo, d), fmt: 'mm/dd/yyyy' };
    }
    case 'L': {
      const ch = String.fromCharCode(raw[0] || 0x3f).toUpperCase();
      if ('TY'.includes(ch)) return { v: true };
      if ('FN'.includes(ch)) return { v: false };
      return null;
    }
    case 'M': case 'G': case 'P': case 'B': {
      if (f.type === 'B' && vfp) return raw.length < 8 ? null : { v: dv(raw).getFloat64(0, true) };
      let block: number;
      if (raw.length === 4) block = u32(raw, 0);
      else { const t = text().trim(); if (!t) return null; block = parseInt(t, 10); }
      if (!block) return null;
      if (f.type === 'G' || f.type === 'P') return { v: f.type === 'G' ? '(OLE object)' : '(picture)' };
      if (!memo) { onMissingMemo(); return { v: '(memo)' }; }
      const s = memo(block);
      return s == null ? null : { v: s };
    }
    case 'I': return { v: i32(raw, 0) };
    case '+': return { v: (u32be(raw, 0) ^ 0x80000000) | 0 };
    case 'Y': {
      const lo = u32(raw, 0), hi = i32(raw, 4);
      return { v: (hi * 4294967296 + lo) / 10000, fmt: '#,##0.0000' };
    }
    case 'O': {
      // dBASE 7 double: big-endian with the sign bit flipped so that byte order sorts.
      if (raw.length < 8) return null;
      const c = Uint8Array.from(raw.subarray(0, 8));
      if (c[0] & 0x80) c[0] ^= 0x80; else for (let i = 0; i < 8; i++) c[i] = ~c[i] & 0xff;
      return { v: new DataView(c.buffer).getFloat64(0, false) };
    }
    case 'T': case '@': {
      let day: number, ms: number;
      if (raw.length < 8) return null;
      if (f.type === 'T') { day = u32(raw, 0); ms = u32(raw, 4); }
      else { const c = Uint8Array.from(raw.subarray(0, 8)); if (c[0] & 0x80) c[0] ^= 0x80; const t = new DataView(c.buffer).getFloat64(0, false); day = Math.floor(t / 86400000); ms = t - day * 86400000; }
      if (!day) return null;
      return { v: day - JULIAN_1899_12_30 + ms / 86400000, fmt: 'mm/dd/yyyy hh:mm:ss' };
    }
    case 'Q': case 'W': return { v: '(binary)' };
  }
  const t = text().trim();
  return t ? { v: t } : null;
}

function dv(raw: Uint8Array): DataView { return new DataView(raw.buffer, raw.byteOffset, raw.byteLength); }

/** Returns a reader for memo blocks: FoxPro .fpt (big-endian block headers) or dBASE .dbt. */
export function memoReader(m: Uint8Array, dec: Decoder): (block: number) => string | null {
  const fptBlock = (m[6] << 8) | m[7];
  const dbt4Block = u16(m, 20);
  const looksFpt = fptBlock >= 1 && fptBlock <= 0x4000 && m.length >= 512 && u32be(m, 0) * fptBlock <= m.length + fptBlock;
  return (block: number) => {
    if (looksFpt) {
      const p = block * fptBlock;
      if (p + 8 > m.length) return null;
      const len = u32be(m, p + 4);
      return dec(m.subarray(p + 8, Math.min(m.length, p + 8 + len))).replace(/\0+$/, '');
    }
    const bs = dbt4Block && dbt4Block % 512 === 0 && dbt4Block <= 0x8000 ? dbt4Block : 512;
    const p = block * bs;
    if (p >= m.length) return null;
    if (m[p] === 0xff && m[p + 1] === 0xff && m[p + 2] === 0x08 && m[p + 3] === 0x00) {
      const len = u32(m, p + 4);
      return dec(m.subarray(p + 8, Math.min(m.length, p + len))).replace(/\0+$/, '');
    }
    let e = p;
    while (e < m.length && !(m[e] === 0x1a && (m[e + 1] === 0x1a || e + 1 === m.length))) e++;
    return dec(m.subarray(p, e)).replace(/\x8d\n/g, '\n').replace(/\r\n?/g, '\n').replace(/\0+$/, '');
  };
}
