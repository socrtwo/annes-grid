/**
 * SQLite 3 database files, read-only: every table becomes a sheet (header row of column names,
 * then one row per record, in rowid order), views are listed in a note.
 *
 * This is a small reader of the documented SQLite file format (public domain specification at
 * sqlite.org/fileformat.html), not the SQLite library: it walks table b-trees (interior and
 * leaf pages, overflow chains) and decodes records (integers, floats, text in the database
 * encoding, blobs). WITHOUT ROWID tables are read from their index b-trees. It does not run
 * SQL, so views and virtual tables have no rows here, and changes still sitting in a separate
 * write-ahead log (-wal file) are not seen.
 */
import { Workbook, Sheet } from '../model/workbook';
import { cellKey, MAX_COLS, MAX_ROWS } from '../model/address';
import type { Scalar } from '../model/types';

type Val = number | bigint | string | Uint8Array | null;

class Db {
  dv: DataView;
  pageSize: number;
  usable: number;
  dec: TextDecoder;
  constructor(public u8: Uint8Array) {
    this.dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    const ps = this.dv.getUint16(16);
    this.pageSize = ps === 1 ? 65536 : ps;
    if (this.pageSize < 512 || (this.pageSize & (this.pageSize - 1))) throw new Error('The SQLite page size is invalid.');
    this.usable = this.pageSize - u8[20];
    const enc = this.dv.getUint32(56);
    this.dec = new TextDecoder(enc === 2 ? 'utf-16le' : enc === 3 ? 'utf-16be' : 'utf-8');
  }
  pageOffset(n: number): number { return (n - 1) * this.pageSize; }
  hasPage(n: number): boolean { return n >= 1 && this.pageOffset(n) + this.pageSize <= this.u8.length; }

  varint(p: number): [number, number] {
    let v = 0;
    for (let i = 0; i < 8; i++) {
      const b = this.u8[p + i];
      v = v * 128 + (b & 0x7f);
      if (!(b & 0x80)) return [v, i + 1];
    }
    return [v * 256 + this.u8[p + 8], 9];
  }

  /** Reads a cell payload of length `len` whose local part starts at p, following overflow pages. */
  payload(p: number, len: number, leafTable: boolean): Uint8Array {
    const U = this.usable;
    const X = leafTable ? U - 35 : Math.floor(((U - 12) * 64) / 255) - 23;
    if (len <= X) return this.u8.subarray(p, p + len);
    const M = Math.floor(((U - 12) * 32) / 255) - 23;
    const K = M + ((len - M) % (U - 4));
    const local = K <= X ? K : M;
    const out = new Uint8Array(len);
    out.set(this.u8.subarray(p, p + local));
    let got = local;
    let next = this.dv.getUint32(p + local);
    let guard = 0;
    while (got < len && next && this.hasPage(next) && guard++ < 1e6) {
      const o = this.pageOffset(next);
      const n = Math.min(len - got, U - 4);
      out.set(this.u8.subarray(o + 4, o + 4 + n), got);
      got += n;
      next = this.dv.getUint32(o);
    }
    return out;
  }

  record(b: Uint8Array): Val[] {
    const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
    const rv = (p: number): [number, number] => {
      let v = 0;
      for (let i = 0; i < 8; i++) { const x = b[p + i]; v = v * 128 + (x & 0x7f); if (!(x & 0x80)) return [v, i + 1]; }
      return [v * 256 + b[p + 8], 9];
    };
    const [hlen, hn] = rv(0);
    const types: number[] = [];
    for (let p = hn; p < hlen;) { const [t, n] = rv(p); types.push(t); p += n; }
    const out: Val[] = [];
    let p = hlen;
    for (const t of types) {
      switch (t) {
        case 0: out.push(null); break;
        case 1: out.push(dv.getInt8(p)); p += 1; break;
        case 2: out.push(dv.getInt16(p)); p += 2; break;
        case 3: out.push((dv.getInt8(p) << 16) | dv.getUint16(p + 1)); p += 3; break;
        case 4: out.push(dv.getInt32(p)); p += 4; break;
        case 5: out.push(dv.getInt16(p) * 2 ** 32 + dv.getUint32(p + 2)); p += 6; break;
        case 6: { const v = dv.getBigInt64(p); out.push(v >= -(2n ** 53n) && v <= 2n ** 53n ? Number(v) : v); p += 8; break; }
        case 7: out.push(dv.getFloat64(p)); p += 8; break;
        case 8: out.push(0); break;
        case 9: out.push(1); break;
        case 10: case 11: out.push(null); break;
        default: {
          const n = t >= 12 && t % 2 === 0 ? (t - 12) / 2 : (t - 13) / 2;
          const bytes = b.subarray(p, p + n);
          out.push(t % 2 === 0 ? bytes : this.dec.decode(bytes));
          p += n;
        }
      }
    }
    return out;
  }

  /** Visits the rows of a table b-tree in key order: fn(rowid, record). */
  walkTable(root: number, fn: (rowid: number, rec: Val[]) => boolean | void): void {
    const seen = new Set<number>();
    const visit = (pg: number): boolean => {
      if (!this.hasPage(pg) || seen.has(pg)) return true;
      seen.add(pg);
      const base = this.pageOffset(pg);
      const h = pg === 1 ? base + 100 : base;
      const type = this.u8[h];
      const n = this.dv.getUint16(h + 3);
      const ptrs = h + (type === 2 || type === 5 ? 12 : 8);
      for (let i = 0; i < n; i++) {
        const c = base + this.dv.getUint16(ptrs + i * 2);
        if (type === 13) {
          const [len, a] = this.varint(c);
          const [rowid, b] = this.varint(c + a);
          if (fn(rowid, this.record(this.payload(c + a + b, len, true))) === false) return false;
        } else if (type === 5) {
          if (!visit(this.dv.getUint32(c))) return false;
        } else if (type === 10 || type === 2) {
          // Index b-tree (WITHOUT ROWID tables): interior cells hold a record too.
          const off = type === 2 ? 4 : 0;
          if (type === 2 && !visit(this.dv.getUint32(c))) return false;
          const [len, a] = this.varint(c + off);
          if (fn(0, this.record(this.payload(c + off + a, len, false))) === false) return false;
        } else return true;
      }
      if (type === 5 || type === 2) return visit(this.dv.getUint32(h + 8));
      return true;
    };
    visit(root);
  }
}

/** Column names (and the rowid alias column) from a CREATE TABLE statement. */
export function parseCreateTable(sql: string): { cols: string[]; rowidCol: number; withoutRowid: boolean; pk: string[] } {
  const open = sql.indexOf('(');
  const close = sql.lastIndexOf(')');
  const body = open >= 0 && close > open ? sql.slice(open + 1, close) : '';
  const parts: string[] = [];
  let depth = 0, cur = '', q = '';
  for (let i = 0; i < body.length; i++) {
    const ch = body[i];
    if (q) { cur += ch; if (ch === q) q = ''; continue; }
    // Comments are not part of the definition (and an apostrophe inside one is not a string).
    if (ch === '-' && body[i + 1] === '-') { const e = body.indexOf('\n', i); i = e < 0 ? body.length : e; cur += ' '; continue; }
    if (ch === '/' && body[i + 1] === '*') { const e = body.indexOf('*/', i + 2); i = e < 0 ? body.length : e + 1; cur += ' '; continue; }
    if (ch === '"' || ch === "'" || ch === '`') { q = ch; cur += ch; continue; }
    if (ch === '[') { q = ']'; cur += ch; continue; }
    if (ch === '(') depth++;
    if (ch === ')') depth--;
    if (ch === ',' && depth === 0) { parts.push(cur.trim()); cur = ''; continue; }
    cur += ch;
  }
  if (cur.trim()) parts.push(cur.trim());
  const cols: string[] = [];
  let rowidCol = -1;
  let pk: string[] = [];
  const intCols = new Map<string, number>();
  const unq = (s: string) => s.replace(/^["'`[]|["'`\]]$/g, '');
  for (const p of parts) {
    const tblPk = /^(?:CONSTRAINT\s+\S+\s+)?PRIMARY\s+KEY\s*\(([^)]*)\)/i.exec(p);
    if (tblPk) { pk = tblPk[1].split(',').map((x) => unq(x.trim().split(/\s+/)[0])); continue; }
    if (/^(CONSTRAINT|UNIQUE|CHECK|FOREIGN\s+KEY)\b/i.test(p)) continue;
    const m = /^("(?:[^"]|"")*"|`[^`]*`|\[[^\]]*\]|'[^']*'|\S+)\s*(.*)$/s.exec(p);
    if (!m) continue;
    const name = unq(m[1]).replace(/""/g, '"');
    // A column declared exactly INTEGER with PRIMARY KEY among its constraints is the rowid: its value is not in the record.
    const isInteger = /^INTEGER\b(?!\s*\()/i.test(m[2].trim());
    if (isInteger && /\bPRIMARY\s+KEY\b(?!\s+DESC)/i.test(m[2])) rowidCol = cols.length;
    if (isInteger) intCols.set(name.toLowerCase(), cols.length);
    if (/\bPRIMARY\s+KEY\b/i.test(m[2])) pk = [name];
    cols.push(name);
  }
  // PRIMARY KEY (id) as a table constraint makes an INTEGER column the rowid just the same.
  if (rowidCol < 0 && pk.length === 1 && intCols.has(pk[0].toLowerCase())) rowidCol = intCols.get(pk[0].toLowerCase())!;
  const withoutRowid = /\)\s*WITHOUT\s+ROWID\s*;?\s*$/i.test(sql);
  return { cols, rowidCol: withoutRowid ? -1 : rowidCol, withoutRowid, pk };
}

function toScalar(v: Val): Scalar {
  if (v == null) return null;
  if (typeof v === 'bigint') return v.toString();
  if (v instanceof Uint8Array) return `[BLOB ${v.length} bytes]`;
  return v;
}

export function importSqlite(data: ArrayBuffer, warnings: string[] = []): Workbook {
  const u8 = new Uint8Array(data);
  if (String.fromCharCode(...u8.subarray(0, 15)) !== 'SQLite format 3') throw new Error('This is not a SQLite 3 database.');
  const db = new Db(u8);
  if (u8[18] === 2 || u8[19] === 2) warnings.push('This database uses a write-ahead log. Changes still in its -wal file are not shown.');
  const schema: { type: string; name: string; root: number; sql: string }[] = [];
  db.walkTable(1, (_r, rec) => { schema.push({ type: String(rec[0] ?? ''), name: String(rec[1] ?? ''), root: Number(rec[3] ?? 0), sql: String(rec[4] ?? '') }); });
  const book = new Workbook(false);
  const head = book.styleId({ bold: true });
  const views: string[] = [];
  for (const t of schema) {
    if (t.type === 'view') { views.push(t.name); continue; }
    if (t.type !== 'table' || !t.root || /^sqlite_/i.test(t.name)) continue;
    if (/^CREATE\s+VIRTUAL\b/i.test(t.sql)) { warnings.push(`Virtual table "${t.name}" has no stored rows and was skipped.`); continue; }
    const def = parseCreateTable(t.sql);
    let nm = t.name.replace(/[\\/?*[\]:]/g, '_').slice(0, 31) || 'Table';
    while (book.sheetByName(nm)) nm = nm.slice(0, 28) + '_' + (book.sheets.length + 1);
    const sheet = new Sheet(book, nm);
    book.sheets.push(sheet);
    const cols = def.cols.slice(0, MAX_COLS);
    cols.forEach((c, i) => sheet.cells.set(cellKey(0, i), { v: c, s: head }));
    sheet.freeze = { rows: 1, cols: 0 };
    let r = 1, truncated = false;
    let order: number[] | null = null;
    if (def.withoutRowid && def.pk.length) {
      // Index records hold the primary key columns first, then the others in table order.
      const pkIdx = def.pk.map((p) => def.cols.findIndex((c) => c.toLowerCase() === p.toLowerCase())).filter((i) => i >= 0);
      order = [...pkIdx, ...def.cols.map((_c, i) => i).filter((i) => !pkIdx.includes(i))];
    }
    try {
      db.walkTable(t.root, (rowid, rec) => {
        if (r >= MAX_ROWS) { truncated = true; return false; }
        for (let i = 0; i < rec.length && i < MAX_COLS; i++) {
          const c = order ? order[i] ?? i : i;
          let v = rec[i];
          if (c === def.rowidCol && v == null) v = rowid;
          const s = toScalar(v);
          if (s != null && s !== '') sheet.cells.set(cellKey(r, c), { v: s });
        }
        if (def.rowidCol >= 0 && rec.length <= def.rowidCol) sheet.cells.set(cellKey(r, def.rowidCol), { v: rowid });
        r++;
      });
    } catch (e) {
      warnings.push(`Table "${t.name}" could not be read completely: ${e instanceof Error ? e.message : e}`);
    }
    if (truncated) warnings.push(`Table "${t.name}" has more rows than a sheet holds; the first ${MAX_ROWS - 1} were opened.`);
    cols.forEach((c, i) => sheet.colW.set(i, Math.max(64, Math.min(240, c.length * 8 + 16))));
  }
  if (views.length) warnings.push(`Views are not opened (they are saved queries): ${views.join(', ')}.`);
  if (!book.sheets.length) { book.sheets.push(new Sheet(book, 'Sheet1')); warnings.push('The database has no tables.'); }
  book.clearHistory();
  return book;
}
