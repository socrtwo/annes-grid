/**
 * Byte-level helpers shared by the legacy importers: little-endian readers, the 80-bit x87
 * extended float used by Lotus 1-2-3 Release 3 and later, DOS code pages and a record walker for
 * the "type, length, body" layout that Lotus, Quattro and Works files share.
 */
import { cp1252 } from '../recover/salvage';
import { decodeWith } from '../encoding';

export const u16 = (b: Uint8Array, o: number): number => (o + 2 <= b.length ? b[o] | (b[o + 1] << 8) : 0);
export const i16 = (b: Uint8Array, o: number): number => { const v = u16(b, o); return v & 0x8000 ? v - 0x10000 : v; };
export const u32 = (b: Uint8Array, o: number): number => (o + 4 <= b.length ? (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0 : 0);
export const i32 = (b: Uint8Array, o: number): number => u32(b, o) | 0;
export const u16be = (b: Uint8Array, o: number): number => (o + 2 <= b.length ? (b[o] << 8) | b[o + 1] : 0);
export const u32be = (b: Uint8Array, o: number): number => (o + 4 <= b.length ? ((b[o] << 24) | (b[o + 1] << 16) | (b[o + 2] << 8) | b[o + 3]) >>> 0 : 0);

export function f64(b: Uint8Array, o: number): number {
  if (o + 8 > b.length) return NaN;
  return new DataView(b.buffer, b.byteOffset + o, 8).getFloat64(0, true);
}

export function f64be(b: Uint8Array, o: number): number {
  if (o + 8 > b.length) return NaN;
  return new DataView(b.buffer, b.byteOffset + o, 8).getFloat64(0, false);
}

/** IEEE double from its upper 32 bits (the lower half is zero). */
export function f64hi(b: Uint8Array, o: number): number {
  const dv = new DataView(new ArrayBuffer(8));
  dv.setUint32(4, u32(b, o), true);
  return dv.getFloat64(0, true);
}

/** 80-bit x87 extended precision: 64-bit mantissa with an explicit integer bit, then sign and a 15-bit exponent. */
export function f80(b: Uint8Array, o: number): number {
  if (o + 10 > b.length) return NaN;
  const lo = u32(b, o), hi = u32(b, o + 4), se = u16(b, o + 8);
  const sign = se & 0x8000 ? -1 : 1;
  const exp = se & 0x7fff;
  if (exp === 0x7fff) return NaN;
  if (exp === 0 && hi === 0 && lo === 0) return 0;
  const mant = hi * 4294967296 + lo;
  return sign * mant * Math.pow(2, exp - 16383 - 63);
}

/** True when an x87 value is a NaN or infinity (Lotus uses these for strings, NA and ERR). */
export const f80Special = (b: Uint8Array, o: number): boolean => (u16(b, o + 8) & 0x7fff) === 0x7fff;

/** Lotus "small number" (2 bytes): an integer, or a small integer times a factor from a table. */
const SNUM_FACTORS = [5000, 500, 0.05, 0.005, 0.0005, 0.00005, 0.0625, 0.015625];
export function snum(v: number): number {
  const s = v & 0x8000 ? v - 0x10000 : v;
  if (s & 1) return SNUM_FACTORS[(s >> 1) & 7] * (s >> 4);
  return s >> 1;
}

/** Lotus 1-2-3 Release 4+ compressed number (4 bytes): mantissa, decimal exponent and two sign bits. */
export function lnum(v: number): number {
  const m = v >>> 6;
  const e = v & 0x0f;
  const p = Math.pow(10, e);
  let x = v & 0x10 ? m / p : m * p;
  if (v & 0x20) x = -x;
  return x;
}

// ------------------------------------------------------------------------------------------
// Code pages

export type Decoder = (b: Uint8Array) => string;

const WIN_PAGES: Record<number, string> = {
  874: 'windows-874', 932: 'shift_jis', 936: 'gbk', 949: 'euc-kr', 950: 'big5', 1250: 'windows-1250', 1251: 'windows-1251',
  1253: 'windows-1253', 1254: 'windows-1254', 1255: 'windows-1255', 1256: 'windows-1256', 1257: 'windows-1257', 1258: 'windows-1258',
  852: 'ibm852', 866: 'ibm866', 65001: 'utf-8',
};

/** A decoder for a code page number (437, 850, 1252, 10000 = Mac Roman, ...). Unknown pages fall back to Windows-1252. */
export function codepage(cp: number): Decoder {
  // The DOS code pages and Mac OS Roman come from the shared text-encoding module.
  if (cp === 437) return (b) => decodeWith(b, 'cp437');
  if (cp === 850 || cp === 865 || cp === 863 || cp === 860 || cp === 861) return (b) => decodeWith(b, 'cp850');
  if (cp === 852) return (b) => decodeWith(b, 'cp852');
  if (cp === 10000 || cp === 32768) return (b) => decodeWith(b, 'macintosh');
  const label = WIN_PAGES[cp];
  if (label) {
    try { const d = new TextDecoder(label); return (b) => d.decode(b); } catch { /* unsupported in this runtime */ }
  }
  return (b) => { let s = ''; for (let i = 0; i < b.length; i++) s += cp1252(b[i]); return s; };
}

/** Text of a zero-terminated byte string. */
export function cstr(b: Uint8Array, o: number, dec: Decoder, max = b.length - o): string {
  let e = o;
  const lim = Math.min(b.length, o + Math.max(0, max));
  while (e < lim && b[e] !== 0) e++;
  return dec(b.subarray(o, e));
}

// ------------------------------------------------------------------------------------------
// Record walking (16-bit type, 16-bit length)

export interface LRec { op: number; off: number; data: Uint8Array }

/**
 * Splits a stream into records. In lenient mode a record whose length runs past the end is
 * truncated instead of stopping the walk; `stop` ends the walk at an end-of-file record.
 */
export function walkRecords(b: Uint8Array, opts: { start?: number; stop?: (op: number) => boolean; lenient?: boolean; max?: number; plausible?: (op: number, len: number) => boolean } = {}): { recs: LRec[]; truncated: boolean; skipped: number } {
  const recs: LRec[] = [];
  let p = opts.start ?? 0;
  let truncated = false, skipped = 0;
  const max = opts.max ?? 5_000_000;
  const ok = (q: number, depth: number): boolean => {
    if (q === b.length) return true;
    if (q + 4 > b.length) return false;
    const op = u16(b, q), len = u16(b, q + 2);
    if (q + 4 + len > b.length || (opts.plausible && !opts.plausible(op, len))) return false;
    return depth <= 0 || ok(q + 4 + len, depth - 1);
  };
  while (p + 4 <= b.length && recs.length < max) {
    const op = u16(b, p), len = u16(b, p + 2);
    if (p + 4 + len > b.length) {
      // A record that runs past the end: look for the next plausible record (resynchronize).
      let q = -1;
      if (opts.plausible) for (let k = p + 1; k + 4 <= b.length && k < p + 65536; k++) if (ok(k, 3)) { q = k; break; }
      if (q > 0) { skipped += q - p; p = q; continue; }
      truncated = true;
      if (opts.lenient) recs.push({ op, off: p, data: b.subarray(p + 4) });
      break;
    }
    recs.push({ op, off: p, data: b.subarray(p + 4, p + 4 + len) });
    p += 4 + len;
    if (opts.stop?.(op)) break;
  }
  return { recs, truncated, skipped };
}

// ------------------------------------------------------------------------------------------
// LMBCS (Lotus Multi-Byte Character Set), the text encoding of 1-2-3 Release 3 and later

let cp850: Decoder | null = null;
const GROUP_PAGES: Record<number, number> = { 0x02: 851, 0x03: 1255, 0x04: 1256, 0x05: 1251, 0x06: 852, 0x08: 1254, 0x0b: 874 };
/**
 * Decodes LMBCS: bytes below 0x80 are ASCII, 0x80-0xFF are code page 850 ("group 1"); a group
 * byte 0x01-0x13 selects another code page for the next character, 0x0F maps the next byte down
 * to a control character, and 0x14 introduces a UTF-16 code unit.
 */
export function lmbcs(b: Uint8Array): string {
  cp850 ??= codepage(850);
  let s = '';
  for (let i = 0; i < b.length; i++) {
    const x = b[i];
    if (x >= 0x20 && x < 0x80) { s += String.fromCharCode(x); continue; }
    if (x >= 0x80) { s += cp850(b.subarray(i, i + 1)); continue; }
    if (x === 0x09 || x === 0x0a || x === 0x0d) { s += String.fromCharCode(x); continue; }
    const n = b[i + 1];
    if (n === undefined) break;
    if (x === 0x14) { s += String.fromCharCode((n << 8) | (b[i + 2] ?? 0)); i += 2; continue; }
    if (x >= 0x10 && x <= 0x13) { i += 2; s += '?'; continue; } // double-byte Asian groups
    i++;
    if (x === 0x0f) { s += String.fromCharCode(n >= 0x20 && n < 0x40 ? n - 0x20 : n); continue; }
    if (x === 0x01) { s += n >= 0x80 ? cp850(b.subarray(i, i + 1)) : String.fromCharCode(n); continue; }
    const page = GROUP_PAGES[x];
    s += n < 0x80 ? String.fromCharCode(n) : page ? codepage(page)(b.subarray(i, i + 1)) : '?';
  }
  return s;
}

/** Pixels for a column width given in characters of the default font (the convention all these programs use). */
export const charsToPx = (chars: number): number => Math.max(0, Math.round(chars * 7 + 5));
