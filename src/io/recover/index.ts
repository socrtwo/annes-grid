/**
 * Damaged-workbook recovery.
 *
 * The pipeline follows the author's Word Recovery / Immortal Unzipper design (socrtwo), adapted
 * from Anne's Pen's document recovery to spreadsheets. Methods run in order and the result
 * with the most recovered cells wins:
 *
 *   a. identify   S2 file identification: the real type regardless of extension, embedded or
 *                 concatenated foreign data, encrypted packages.
 *   b. open       the normal (lenient) importers.
 *   c. zip        byte-level ZIP salvage: local headers instead of the central directory,
 *                 ImmortalInflate for damaged deflate data, CRC-32 checks.
 *   d. xml        XML repair of every part (healing, spreadsheet-specific fixes, regenerated
 *                 package structure), re-zipped and opened with the normal importer.
 *   e. rawsheet   regex scan of worksheet XML fragments for cells, built directly.
 *   f. biff       lenient BIFF reader for .xls, with record resynchronization and a raw-byte
 *                 scan when the compound file is damaged.
 *   g. ods        OpenDocument content.xml salvage, healing and a raw table-cell scan.
 *   h. text       delimited text, HTML tables and SpreadsheetML 2003 fallbacks.
 *   i. strings    printable-strings scan, as a last resort (text only).
 *
 * recoverWorkbook never throws for damaged input: when nothing works, the report has book null.
 */
import JSZip from 'jszip';
import { Workbook, Sheet } from '../../model/workbook';
import { cellKey, colIndex, colName, MAX_COLS, MAX_ROWS } from '../../model/address';
import { CellError, errorFromCode, type Cell, type Scalar } from '../../model/types';
import { partsToSerial } from '../../model/dates';
import { scan, identifyAt, recommend, DESCRIPTIONS, type Identification, type Segment, type S2Program } from './fileid';
import { crc32 } from './inflate';
import { isOle2, readCfb } from './doc97';
import { scanZipBytes, repairXmlPart, stripInvalidChars, balanceXml, isWellFormed, findDamageCut, stringsScan } from './salvage';
import { readBiff, scanBofs, XLS_ENCRYPTED_MSG } from './xls';

export type MethodStatus = 'success' | 'partial' | 'failed' | 'skipped';
export type MethodId = 'identify' | 'open' | 'zip' | 'xml' | 'rawsheet' | 'biff' | 'ods' | 'text' | 'strings';

export interface MethodResult {
  id: MethodId;
  name: string;
  /** What the method does (shown dimmed in the report). */
  desc: string;
  status: MethodStatus;
  detail: string;
  cells?: number;
  sheets?: number;
}

export type SpreadsheetKind = 'xlsx' | 'xlsm' | 'ods' | 'fods' | 'xls' | 'csv' | 'html' | 'xml2003' | 'qgw' | 'encrypted' | 'other' | 'unknown';

export interface RepairedPackage {
  name: string;
  blob: Blob;
  kind: 'xlsx' | 'ods';
  repairedParts: string[];
  salvagedParts: string[];
  lostParts: string[];
}

export interface RecoveryReport {
  fileName: string;
  /** What the bytes really are. */
  kind: SpreadsheetKind;
  /** Format id for the importer registry (xlsx, ods, xls, csv, ...). */
  format: string;
  book: Workbook | null;
  methods: MethodResult[];
  identification: Identification;
  /** Pieces of other file types found inside the file (concatenated or embedded data). */
  foreign: Segment[];
  /** S2 tools that specialize in this kind of file. */
  recommended: S2Program[];
  repaired?: RepairedPackage;
  summary: string;
  /** Some recovered data was decoded after a damaged area and should be checked. */
  unverified?: boolean;
  /** Only loose text could be recovered (strings scan). */
  textOnly?: boolean;
  bestMethod: MethodId | null;
  notes: string[];
  warnings: string[];
  stats: { sheets: number; cells: number; formulas: number };
}

export interface RecoverOptions {
  /** Run every repair method even when the file opens normally (Open and Repair). */
  forceRepair?: boolean;
  onProgress?: (msg: string) => void;
}

const METHOD_INFO: Record<MethodId, [string, string]> = {
  identify: ['File identification', 'Identify the file by its signature (S2 File Identifier), regardless of its extension; find embedded or concatenated data and encrypted packages.'],
  open: ['Normal open', 'Open the file with the regular importer in lenient mode.'],
  zip: ['ZIP salvage', 'Walk the PK\\x03\\x04 local headers instead of the central directory, inflate each part with ImmortalInflate (keeps everything before the damage) and check CRC-32 values.'],
  xml: ['XML repair and rebuild', 'Heal or rebuild damaged XML, cut worksheets at the last complete row, renumber rows and cells, repair shared strings and styles, regenerate the package structure, then open the rebuilt package.'],
  rawsheet: ['Raw worksheet scan', 'Scan worksheet XML fragments for cells, values, formulas and inline strings, and build the workbook directly.'],
  biff: ['Excel 97-2003 record salvage', 'Read the BIFF record stream leniently: resynchronize after damaged records and, when the compound file is damaged, scan the raw bytes for BIFF substreams.'],
  ods: ['OpenDocument salvage', 'Heal content.xml and open it; otherwise scan it for table cells.'],
  text: ['Text, HTML and XML 2003 fallbacks', 'Read delimited text, HTML tables and Excel 2003 XML (SpreadsheetML) tolerantly.'],
  strings: ['Strings scan (last resort)', 'Collect readable text runs from the raw bytes; tabular lines are split into columns. Formatting and formulas are lost.'],
};

const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const ODS_MIME = 'application/vnd.oasis.opendocument.spreadsheet';
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

const KIND_DESC: Record<string, string> = {
  xlsx: 'Microsoft Excel workbook (XLSX)', xlsm: 'Microsoft Excel macro-enabled workbook (XLSM)', ods: 'OpenDocument spreadsheet (ODS)',
  fods: 'Flat OpenDocument spreadsheet (FODS)', xls: 'Microsoft Excel 97-2003 workbook (XLS)', csv: 'Delimited text (CSV/TSV)',
  html: 'HTML document (tables)', xml2003: 'Excel 2003 XML spreadsheet (SpreadsheetML)', qgw: 'Anne\'s Grid workbook',
  encrypted: 'Password-protected (encrypted) Office file', unknown: 'Unrecognized data',
};

// Importers being written alongside this module; only the ones that exist are bundled.
const LOADERS = import.meta.glob(['../xlsx-import.ts', '../ods.ts', '../csv.ts', '../html.ts', '../other.ts', '../native.ts']) as Record<string, () => Promise<any>>;
async function loadMod(name: string): Promise<any | null> {
  const l = LOADERS[`../${name}.ts`];
  if (!l) return null;
  try { return await l(); } catch { return null; }
}
/** Importers return a Workbook or { book, warnings }. */
function asBook(r: any, warnings: string[]): Workbook | null {
  if (!r) return null;
  if (r instanceof Workbook) return r;
  if (r.book instanceof Workbook) { if (Array.isArray(r.warnings)) warnings.push(...r.warnings); return r.book; }
  return null;
}

// ------------------------------------------------------------------------------------------
// Small helpers

const dec = (u8: Uint8Array): string => {
  if (u8[0] === 0xff && u8[1] === 0xfe) return new TextDecoder('utf-16le').decode(u8.subarray(2));
  if (u8[0] === 0xfe && u8[1] === 0xff) return new TextDecoder('utf-16be').decode(u8.subarray(2));
  const t = new TextDecoder('utf-8', { fatal: false }).decode(u8);
  return t.charCodeAt(0) === 0xfeff ? t.slice(1) : t;
};
const enc = (s: string) => new TextEncoder().encode(s);
const ab = (u8: Uint8Array): ArrayBuffer => u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength) as ArrayBuffer;
const plural = (n: number, w: string) => `${n.toLocaleString('en-US')} ${n === 1 ? w : /[^aeiou]y$/.test(w) ? w.slice(0, -1) + 'ies' : w + 's'}`;
const xmlEsc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const safeChar = (cp: number) => { try { return String.fromCodePoint(cp); } catch { return ''; } };
export const unent = (s: string) => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'")
  .replace(/&#x([0-9a-fA-F]+);/g, (_m, h: string) => safeChar(parseInt(h, 16))).replace(/&#(\d+);/g, (_m, d: string) => safeChar(parseInt(d, 10))).replace(/&amp;/g, '&')
  .replace(/_x([0-9A-Fa-f]{4})_/g, (_m, h: string) => String.fromCharCode(parseInt(h, 16)));

function attrs(s: string): Record<string, string> {
  const out: Record<string, string> = {};
  for (const m of s.matchAll(/([\w:.-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g)) out[m[1]] = m[2] ?? m[3] ?? '';
  return out;
}
const localAttr = (a: Record<string, string>, local: string): string | undefined => {
  if (a[local] != null) return a[local];
  for (const k of Object.keys(a)) if (k.endsWith(':' + local)) return a[k];
  return undefined;
};

function bookStats(book: Workbook | null): { sheets: number; cells: number; formulas: number } {
  if (!book) return { sheets: 0, cells: 0, formulas: 0 };
  let cells = 0, formulas = 0;
  for (const sh of book.sheets) for (const c of sh.cells.values()) { if (c.v != null && c.v !== '' || c.f) cells++; if (c.f) formulas++; }
  return { sheets: book.sheets.length, cells, formulas };
}

function parseA1(ref: string): { r: number; c: number } | null {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/.exec(ref.trim());
  if (!m) return null;
  const c = colIndex(m[1]), r = +m[2] - 1;
  return c >= 0 && c < MAX_COLS && r >= 0 && r < MAX_ROWS ? { r, c } : null;
}

/** Shift the relative A1 references of a formula (shared-formula followers). Strings and quoted names are left alone. */
export function shiftFormula(f: string, dr: number, dc: number): string {
  return f.split(/("(?:[^"]|"")*"|'(?:[^']|'')*')/).map((part, i) => i % 2 ? part : part.replace(
    /(^|[^A-Za-z0-9_.$])(\$?)([A-Z]{1,3})(\$?)(\d{1,7})(?![\d(A-Za-z_])/g,
    (m, pre: string, ac: string, col: string, ar: string, row: string) => {
      let c = colIndex(col), r = +row - 1;
      if (c < 0) return m;
      if (!ac) c += dc;
      if (!ar) r += dr;
      if (c < 0 || r < 0 || c > 16383 || r > 1048575) return `${pre}#REF!`;
      return `${pre}${ac}${colName(c)}${ar}${r + 1}`;
    })).join('');
}

/** Tag balance check for small fragments (faster than a DOM parse per cell). */
function balanced(frag: string): boolean {
  const stack: string[] = [];
  for (const m of frag.matchAll(/<(\/?)([A-Za-z_][\w.:-]*)[^<>]*?(\/?)>|<|>/g)) {
    if (m[0] === '<') return false;
    if (m[0] === '>') continue;
    if (m[1]) { if (stack.pop() !== m[2]) return false; }
    else if (!m[3]) stack.push(m[2]);
  }
  return stack.length === 0;
}

// ------------------------------------------------------------------------------------------
// Identification

function looksTextual(u8: Uint8Array): boolean {
  const n = Math.min(u8.length, 4096);
  if (!n) return false;
  let bad = 0;
  for (let i = 0; i < n; i++) { const c = u8[i]; if (c === 0) return false; if (c < 9 || (c > 13 && c < 32)) bad++; }
  return bad / n < 0.02;
}

function hasAscii(u8: Uint8Array, s: string, limit = u8.length): boolean {
  const pat = enc(s);
  const end = Math.min(u8.length, limit) - pat.length;
  outer: for (let i = 0; i <= end; i++) {
    if (u8[i] !== pat[0]) continue;
    for (let j = 1; j < pat.length; j++) if (u8[i + j] !== pat[j]) continue outer;
    return true;
  }
  return false;
}

function detectKind(u8: Uint8Array, fileName: string): { kind: SpreadsheetKind; id: Identification } {
  const lower = fileName.toLowerCase();
  const res = (kind: SpreadsheetKind, confidence: Identification['confidence'] = 'high', ext: string = kind): { kind: SpreadsheetKind; id: Identification } =>
    ({ kind, id: { ext, description: KIND_DESC[kind] || DESCRIPTIONS[ext] || ext, confidence } });
  if (!u8.length) return res('unknown', 'low', 'bin');
  const base = identifyAt(u8);
  const zipLike = base.ext === 'xlsx' || base.ext === 'ods' || base.ext === 'ooxml' || base.ext === 'zip' || ['docx', 'pptx', 'odt', 'odp', 'odg', 'epub', 'jar', 'apk'].includes(base.ext);
  if (zipLike) {
    if (base.ext === 'xlsx') return res(hasAscii(u8, 'xl/vbaProject.bin') || /\.xlsm$/.test(lower) ? 'xlsm' : 'xlsx');
    if (base.ext === 'ods') return res('ods');
    if (base.ext === 'ooxml' || base.ext === 'zip') {
      if (hasAscii(u8, 'xl/worksheets/') || /\.xl[st][xmb]$/.test(lower)) return res('xlsx', 'medium');
      if (hasAscii(u8, 'content.xml') && /\.o[dt]s$/.test(lower)) return res('ods', 'medium');
      if (/\.o[dt]s$/.test(lower)) return res('ods', 'low');
      return { kind: 'other', id: base };
    }
    return { kind: 'other', id: base };
  }
  if (base.ext === 'xls' || base.ext === 'ole' || base.ext === 'doc' || base.ext === 'ppt' || base.ext === 'msg' || base.ext === 'vsd') {
    let streams: string[] = [];
    try { streams = [...(readCfb(u8)?.streams.keys() || [])]; } catch { /* damaged */ }
    if (streams.includes('EncryptedPackage') || streams.includes('EncryptionInfo')) return res('encrypted', 'high', 'ooxml');
    if (base.ext === 'xls') return res('xls');
    if (base.ext === 'ole' && (/\.xl[stw]$/.test(lower) || scanBofs(u8.subarray(0, Math.min(u8.length, 1 << 20))).length)) return res('xls', 'medium');
    return { kind: 'other', id: base };
  }
  // Raw BIFF2-4 worksheets and bare BIFF streams.
  if (u8[0] === 0x09 && (u8[1] === 0x00 || u8[1] === 0x02 || u8[1] === 0x04 || u8[1] === 0x08) && u8[2] >= 4 && u8[2] <= 20 && u8[3] === 0) return res('xls');
  if (looksTextual(u8) || base.ext === 'xml' || base.ext === 'html' || base.ext === 'txt') {
    const head = dec(u8.subarray(0, 4096));
    const low = head.toLowerCase();
    if (/^\s*\{/.test(head) && /"sheets"\s*:/.test(head)) return res('qgw');
    if (low.includes('urn:schemas-microsoft-com:office:spreadsheet') || (/<workbook\b/.test(low) && /<worksheet\b/.test(low))) return res('xml2003');
    if (low.includes('office:document') || low.includes('opendocument.spreadsheet')) return res('fods');
    if (/<!doctype html|<html|<table/.test(low)) return res('html');
    if (/^\s*</.test(head)) return { kind: 'other', id: base };
    return res('csv', 'medium', 'csv');
  }
  if (/\.xl[stw]$/.test(lower) && scanBofs(u8).length) return res('xls', 'low');
  if (/\.xls[xmb]$/.test(lower) && hasAscii(u8, 'PK\x03\x04')) return res('xlsx', 'low');
  if (/\.ods$/.test(lower) && hasAscii(u8, 'PK\x03\x04')) return res('ods', 'low');
  if (/\.xls[xm]$/.test(lower)) return res('xlsx', 'low', 'xlsx');
  if (/\.xls$/.test(lower)) return res('xls', 'low', 'xls');
  if (/\.ods$/.test(lower)) return res('ods', 'low', 'ods');
  return res('unknown', 'low', base.ext);
}

const SPREADSHEET_KINDS: SpreadsheetKind[] = ['xlsx', 'xlsm', 'ods', 'fods', 'xls', 'csv', 'html', 'xml2003', 'qgw', 'encrypted'];

// ------------------------------------------------------------------------------------------
// ZIP salvage

interface Salvage {
  parts: Map<string, Uint8Array>;
  corrupt: Set<string>;
  salvaged: string[];
  zipOk: boolean;
  zipError: string;
}

async function salvageZip(u8: Uint8Array, deep: boolean): Promise<Salvage> {
  const s: Salvage = { parts: new Map(), corrupt: new Set(), salvaged: [], zipOk: false, zipError: '' };
  const damaged = new Set<string>();
  try {
    const zip = await JSZip.loadAsync(u8);
    s.zipOk = true;
    for (const f of Object.values(zip.files)) {
      if (f.dir) continue;
      try {
        const bytes = await f.async('uint8array');
        const want = (f as any)._data?.crc32;
        if (typeof want === 'number' && want !== 0 && crc32(bytes) !== want >>> 0) { damaged.add(f.name); s.corrupt.add(f.name); }
        s.parts.set(f.name, bytes);
      } catch { damaged.add(f.name); }
    }
  } catch (e: any) { s.zipError = e?.message || String(e); }
  if (deep || !s.zipOk || damaged.size) {
    const scanned = scanZipBytes(u8);
    for (const [name, e] of scanned) {
      if (!e.isCorrupt && damaged.has(name)) s.corrupt.delete(name);
      if (e.isCorrupt && (!s.parts.has(name) || damaged.has(name))) s.corrupt.add(name);
      if (!s.parts.has(name) || damaged.has(name)) { s.parts.set(name, e.data); s.salvaged.push(name); }
    }
  }
  return s;
}

// ------------------------------------------------------------------------------------------
// XLSX: worksheet preparation (damage cut), cell scanning, part repair

interface SheetPrep { main: string; tail: string; cut: boolean }

const isSheetPart = (n: string) => /^xl\/worksheets\/[^/]+\.xml$/i.test(n);
const sheetNum = (n: string) => +(/(\d+)\.xml$/i.exec(n)?.[1] ?? 1e6);

/** Keep the consistent prefix of a damaged worksheet; the rest is "unverified". */
function prepareSheet(text: string, corrupt: boolean): SheetPrep {
  let xml = stripInvalidChars(text);
  let cut = -1;
  if (corrupt || !isWellFormedFast(xml)) {
    cut = findDamageCut(xml, '</row>');
    if (cut < 0 && !/<\/sheetData>/.test(xml)) {
      const last = xml.lastIndexOf('</row>');
      if (last >= 0) cut = last + 6;
      else { const sd = /<sheetData[^>]*>/.exec(xml); if (sd) cut = sd.index + sd[0].length; }
    }
  }
  if (cut > 0) return { main: xml.slice(0, cut), tail: xml.slice(cut), cut: true };
  return { main: xml, tail: '', cut: false };
}

/** Cheap structural test before the DOM parser: balanced root and no stray control characters. */
function isWellFormedFast(xml: string): boolean {
  if (!/<\/(?:\w+:)?worksheet>\s*$/.test(xml)) return false;
  return isWellFormed(xml);
}

interface ScannedCell { r: number; c: number; cell: Cell }

/** Regex scan of worksheet XML (complete or fragmentary) for cells. */
export function scanSheetCells(xml: string, sst: string[], stats?: { missingStrings: number; dropped: number }): ScannedCell[] {
  const out: ScannedCell[] = [];
  const shared = new Map<string, { f: string; r: number; c: number }>();
  const re = /<row\b([^>]*?)(\/?)>|<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
  let row = -1, prevCol = -1;
  let m: RegExpExecArray | null;
  while ((m = re.exec(xml)) !== null) {
    if (m[0].startsWith('<row')) {
      const a = attrs(m[1]);
      const r = a.r ? parseInt(a.r, 10) - 1 : NaN;
      row = Number.isFinite(r) && r > row && r < MAX_ROWS ? r : row + 1;
      prevCol = -1;
      continue;
    }
    const inner = m[4] ?? '';
    const nested = inner.search(/<c[\s>/]/);
    if (nested >= 0) { re.lastIndex = m.index + m[0].length - inner.length - 4 + nested; if (stats) stats.dropped++; continue; }
    const a = attrs(m[3]);
    let r = row < 0 ? 0 : row, c = prevCol + 1;
    const p = a.r ? parseA1(a.r) : null;
    if (p && p.c > prevCol && (row < 0 || p.r === row)) c = p.c;
    if (p && row < 0) { r = p.r; row = r; }
    if (c >= MAX_COLS) continue;
    prevCol = c;
    const t = a.t || 'n';
    const vm = /<v>([\s\S]*?)<\/v>/.exec(inner);
    const fm = /<f\b([^>]*?)(?:\/>|>([\s\S]*?)<\/f>)/.exec(inner);
    let v: Scalar = null;
    const raw = vm ? vm[1] : null;
    if (t === 's') {
      if (raw != null) {
        const s = sst[parseInt(raw, 10)];
        if (s === undefined) { if (stats) stats.missingStrings++; } else v = s;
      }
    } else if (t === 'inlineStr') {
      const is = /<is>([\s\S]*?)(?:<\/is>|$)/.exec(inner);
      v = is ? siText(is[1]) : raw != null ? unent(raw) : null;
    } else if (t === 'str') v = raw != null ? unent(raw) : '';
    else if (t === 'b') v = raw != null ? raw.trim() === '1' || raw.trim() === 'true' : null;
    else if (t === 'e') v = raw != null ? errorFromCode(unent(raw).trim()) || new CellError('#N/A') : null;
    else if (t === 'd') v = raw != null ? isoToSerial(raw) ?? unent(raw) : null;
    else if (raw != null && raw.trim() !== '') {
      const n = Number(raw);
      if (!Number.isFinite(n)) { if (stats) stats.dropped++; continue; }
      v = n;
    }
    const cell: Cell = { v };
    if (fm) {
      const fa = attrs(fm[1]);
      const text = fm[2] != null ? unent(fm[2]).trim() : '';
      if (text) {
        cell.f = text.replace(/^=/, '');
        if (fa.t === 'shared' && fa.si != null) shared.set(fa.si, { f: cell.f, r, c });
        if (fa.t === 'array' && fa.ref) {
          const [x, y] = fa.ref.split(':');
          const p1 = parseA1(x), p2 = parseA1(y || x);
          if (p1 && p2) cell.arr = { r1: p1.r, c1: p1.c, r2: p2.r, c2: p2.c };
        }
      } else if (fa.t === 'shared' && fa.si != null) {
        const master = shared.get(fa.si);
        if (master) cell.f = shiftFormula(master.f, r - master.r, c - master.c);
      }
    }
    if (cell.v == null && !cell.f) continue;
    out.push({ r, c, cell });
  }
  return out;
}

function isoToSerial(s: string): number | null {
  const m = /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?)?/.exec(s.trim());
  if (!m) return null;
  let v = partsToSerial(+m[1], +m[2], +m[3]);
  if (m[4]) v += (+m[4] * 3600 + +m[5] * 60 + +(m[6] || 0)) / 86400;
  return v;
}

/** Text of a shared-string item: its <t> runs (phonetic <rPh> runs excluded), else its stripped text. */
function siText(inner: string): string {
  const body = inner.replace(/<rPh\b[\s\S]*?(?:<\/rPh>|$)/g, '');
  const runs = [...body.matchAll(/<t\b[^>]*?(?:\/>|>([\s\S]*?)(?:<\/t>|$))/g)].map((m) => m[1] ?? '');
  if (runs.length) return unent(runs.join('').replace(/<[^>]*>?/g, ''));
  return unent(body.replace(/<[^>]*>?/g, ''));
}

/** Shared strings from a possibly damaged part: complete items plus a truncated last item. */
export function parseSharedStrings(xml: string): { items: string[]; truncated: boolean } {
  const items: string[] = [];
  let truncated = false;
  const chunks = stripInvalidChars(xml).split(/<si\b[^>]*>/);
  for (let i = 1; i < chunks.length; i++) {
    const ch = chunks[i];
    const end = ch.indexOf('</si>');
    if (end < 0) {
      // Truncated (last) item, or an item whose end tag was lost: keep the text it has.
      const text = siText(ch.replace(/<\/sst>[\s\S]*$/, ''));
      if (i === chunks.length - 1) truncated = true;
      items.push(text);
    } else items.push(siText(ch.slice(0, end)));
  }
  return { items, truncated };
}

function rebuildSharedStrings(items: string[]): string {
  return `${XML_DECL}<sst xmlns="${MAIN_NS}" count="${items.length}" uniqueCount="${items.length}">${items.map((t) => `<si><t xml:space="preserve">${xmlEsc(t)}</t></si>`).join('')}</sst>`;
}

/**
 * Spreadsheet-specific worksheet repair: close sheetData/worksheet, drop broken cells, and
 * renumber rows and cells whose r attributes are missing, duplicated or out of order.
 */
export function fixSheetXml(xml: string): { xml: string; fixes: number; dropped: number; closed: boolean } {
  let s = stripInvalidChars(xml);
  let fixes = 0, dropped = 0, closed = false;
  const sdOpen = /<sheetData\b[^>]*?(\/?)>/.exec(s);
  if (!sdOpen) {
    const firstRow = s.search(/<row\b/);
    if (firstRow < 0) {
      if (!/<worksheet\b/.test(s)) return { xml: `${XML_DECL}<worksheet xmlns="${MAIN_NS}"><sheetData/></worksheet>`, fixes, dropped, closed: true };
      return { xml: s, fixes, dropped, closed };
    }
    s = `${s.slice(0, firstRow)}<sheetData>${s.slice(firstRow)}`;
    return { ...fixSheetXml(s), closed: true };
  }
  let head = s.slice(0, sdOpen.index);
  if (!/<worksheet\b/.test(head)) { head = `${XML_DECL}<worksheet xmlns="${MAIN_NS}" xmlns:r="${REL_NS}">`; closed = true; }
  if (sdOpen[1]) return { xml: s, fixes, dropped, closed };
  const bodyStart = sdOpen.index + sdOpen[0].length;
  const close = s.indexOf('</sheetData>', bodyStart);
  const body = close >= 0 ? s.slice(bodyStart, close) : s.slice(bodyStart);
  let after = close >= 0 ? s.slice(close + 12) : '</worksheet>';
  if (close < 0) closed = true;
  if (!/<\/(?:\w+:)?worksheet>\s*$/.test(after)) { after = after.replace(/<[^>]*$/, '') + '</worksheet>'; closed = true; }

  const rows: string[] = [];
  let prevRow = 0;
  const rowRe = /<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g;
  let m: RegExpExecArray | null;
  while ((m = rowRe.exec(body)) !== null) {
    const inner = m[2] ?? '';
    const nestedRow = inner.search(/<row\b/);
    if (nestedRow >= 0) { rowRe.lastIndex = m.index + m[0].length - inner.length - 6 + nestedRow; dropped++; continue; }
    const a = attrs(m[1]);
    let r = parseInt(a.r ?? '', 10);
    if (!Number.isFinite(r) || r <= prevRow || r > MAX_ROWS) { r = prevRow + 1; fixes++; }
    prevRow = r;
    const cells: string[] = [];
    let prevCol = -1;
    const cellRe = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
    let cm: RegExpExecArray | null;
    while ((cm = cellRe.exec(inner)) !== null) {
      const ci = cm[2] ?? '';
      const nested = ci.search(/<c[\s>/]/);
      if (nested >= 0) { cellRe.lastIndex = cm.index + cm[0].length - ci.length - 4 + nested; dropped++; continue; }
      const ca = attrs(cm[1]);
      if (ci && !balanced(ci)) { dropped++; continue; }
      const t = ca.t || 'n';
      const v = /<v>([\s\S]*?)<\/v>/.exec(ci)?.[1];
      if (v != null && t === 'n' && v.trim() !== '' && !Number.isFinite(Number(v))) { dropped++; continue; }
      if (v != null && t === 's' && !/^\s*\d+\s*$/.test(v)) { dropped++; continue; }
      const p = ca.r ? parseA1(ca.r) : null;
      let c = p && p.c > prevCol ? p.c : prevCol + 1;
      if (!p || p.c <= prevCol || p.r !== r - 1) fixes++;
      if (c >= MAX_COLS) { dropped++; continue; }
      prevCol = c;
      const rest = Object.entries(ca).filter(([k]) => k !== 'r').map(([k, val]) => ` ${k}="${xmlEsc(val)}"`).join('');
      cells.push(`<c r="${colName(c)}${r}"${rest}${ci ? `>${ci}</c>` : '/>'}`);
    }
    const rattrs = Object.entries(a).filter(([k]) => k !== 'r' && k !== 'spans').map(([k, val]) => ` ${k}="${xmlEsc(val)}"`).join('');
    rows.push(`<row r="${r}"${rattrs}>${cells.join('')}</row>`);
  }
  return { xml: `${head}<sheetData>${rows.join('')}</sheetData>${after}`, fixes, dropped, closed };
}

const MINIMAL_STYLES = `${XML_DECL}<styleSheet xmlns="${MAIN_NS}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;

function stylesUsable(xml: string, maxS: number): boolean {
  const xfs = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(xml);
  if (!xfs) return maxS < 0;
  const n = (xfs[1].match(/<xf\b/g) || []).length;
  const has = (tag: string) => new RegExp(`<${tag}\\b[^>]*>[\\s\\S]*?<${tag.slice(0, -1)}\\b`).test(xml);
  return n > maxS && has('fonts') && has('fills') && has('borders');
}

interface RelEntry { id: string; type: string; target: string; external: boolean; raw: string }
function parseRels(xml: string): RelEntry[] {
  return [...xml.matchAll(/<Relationship\b([^>]*?)\/?>/g)].map((m) => {
    const a = attrs(m[1]);
    return { id: a.Id || '', type: a.Type || '', target: a.Target || '', external: a.TargetMode === 'External', raw: m[0].replace(/>$/, '/>').replace(/\/\/>$/, '/>') };
  });
}
function relsXml(rels: { id: string; type: string; target: string; external?: boolean }[]): string {
  return `${XML_DECL}<Relationships xmlns="${PKG_REL_NS}">${rels.map((r) => `<Relationship Id="${xmlEsc(r.id)}" Type="${xmlEsc(r.type)}" Target="${xmlEsc(r.target)}"${r.external ? ' TargetMode="External"' : ''}/>`).join('')}</Relationships>`;
}
function resolveTarget(relsPart: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const dir = relsPart.replace(/_rels\/[^/]*$/, '');
  const segs = (dir + target).split('/');
  const out: string[] = [];
  for (const s of segs) { if (s === '..') out.pop(); else if (s && s !== '.') out.push(s); }
  return out.join('/');
}

const CT: [RegExp, string][] = [
  [/^xl\/worksheets\/[^/]+\.xml$/i, 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'],
  [/^xl\/chartsheets\/[^/]+\.xml$/i, 'application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml'],
  [/^xl\/sharedStrings\.xml$/i, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'],
  [/^xl\/styles\.xml$/i, 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'],
  [/^xl\/theme\/[^/]+\.xml$/i, 'application/vnd.openxmlformats-officedocument.theme+xml'],
  [/^xl\/comments[^/]*\.xml$/i, 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml'],
  [/^xl\/drawings\/[^/]+\.xml$/i, 'application/vnd.openxmlformats-officedocument.drawing+xml'],
  [/^xl\/tables\/[^/]+\.xml$/i, 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml'],
  [/^xl\/charts\/chart[^/]*\.xml$/i, 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'],
  [/^xl\/persons\/[^/]+\.xml$/i, 'application/vnd.ms-excel.person+xml'],
  [/^xl\/threadedComments\/[^/]+\.xml$/i, 'application/vnd.ms-excel.threadedcomments+xml'],
  [/^docProps\/core\.xml$/i, 'application/vnd.openxmlformats-package.core-properties+xml'],
  [/^docProps\/app\.xml$/i, 'application/vnd.openxmlformats-officedocument.extended-properties+xml'],
  [/^xl\/vbaProject\.bin$/i, 'application/vnd.ms-office.vbaProject'],
];
const DEFAULT_CT: Record<string, string> = {
  rels: 'application/vnd.openxmlformats-package.relationships+xml', xml: 'application/xml', vml: 'application/vnd.openxmlformats-officedocument.vmlDrawing',
  png: 'image/png', jpeg: 'image/jpeg', jpg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', tif: 'image/tiff', tiff: 'image/tiff', emf: 'image/x-emf', wmf: 'image/x-wmf', svg: 'image/svg+xml', bin: 'application/vnd.openxmlformats-officedocument.oleObject',
};

function contentTypes(names: string[], macro: boolean): string {
  const exts = new Set(names.map((n) => n.slice(n.lastIndexOf('.') + 1).toLowerCase()).filter((e) => DEFAULT_CT[e]));
  exts.add('rels'); exts.add('xml');
  const over: string[] = [];
  for (const n of names) {
    if (/^xl\/workbook\.xml$/i.test(n)) { over.push(`<Override PartName="/${n}" ContentType="${macro ? 'application/vnd.ms-excel.sheet.macroEnabled.main+xml' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'}"/>`); continue; }
    const hit = CT.find(([re]) => re.test(n));
    if (hit) over.push(`<Override PartName="/${n}" ContentType="${hit[1]}"/>`);
  }
  return `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${[...exts].map((e) => `<Default Extension="${e}" ContentType="${DEFAULT_CT[e]}"/>`).join('')}${over.join('')}</Types>`;
}

interface XlsxRepair {
  parts: Map<string, Uint8Array>;
  repaired: string[];
  lost: string[];
  sheetOrder: { part: string; name: string; state?: string }[];
  stylesReplaced: boolean;
}

/** Workbook sheet list (name, r:id, state) and the workbook relationships, tolerant of damage. */
function workbookSheets(parts: Map<string, Uint8Array>): { part: string; name: string; state?: string }[] {
  const wb = parts.get('xl/workbook.xml');
  const relsB = parts.get('xl/_rels/workbook.xml.rels');
  const rels = relsB ? parseRels(dec(relsB)) : [];
  const out: { part: string; name: string; state?: string }[] = [];
  if (wb) {
    for (const m of dec(wb).matchAll(/<(?:\w+:)?sheet\b([^>]*?)\/?>/g)) {
      const a = attrs(m[1]);
      const rid = localAttr(a, 'id');
      const rel = rels.find((r) => r.id === rid);
      const part = rel ? resolveTarget('xl/_rels/workbook.xml.rels', rel.target) : '';
      out.push({ part, name: unent(a.name || ''), state: a.state });
    }
  }
  return out;
}

function repairXlsx(src: Map<string, Uint8Array>, corrupt: Set<string>, preps: Map<string, SheetPrep>, force: boolean): XlsxRepair {
  const parts = new Map(src);
  const repaired: string[] = [], lost: string[] = [];
  const sheetParts = [...parts.keys()].filter(isSheetPart).sort((a, b) => sheetNum(a) - sheetNum(b) || a.localeCompare(b));
  // Stale calculation chains only cause trouble after a repair.
  if (parts.delete('xl/calcChain.xml')) repaired.push('xl/calcChain.xml (removed; recalculated on open)');

  // 1. Worksheets
  let maxS = -1;
  for (const name of sheetParts) {
    const prep = preps.get(name)!;
    const original = dec(parts.get(name)!);
    const needs = force || prep.cut || corrupt.has(name) || !isWellFormedFast(original);
    let xml = prep.main;
    if (needs) {
      const f = fixSheetXml(prep.cut ? prep.main : original);
      xml = f.xml;
      const steps: string[] = [];
      if (prep.cut) steps.push('cut at the last complete row before the damage');
      if (f.closed) steps.push('sheetData and worksheet closed');
      if (f.fixes) steps.push(`${plural(f.fixes, 'row/cell reference')} fixed`);
      if (f.dropped) steps.push(`${plural(f.dropped, 'broken cell')} dropped`);
      if (!isWellFormed(xml)) {
        const r = repairXmlPart(name, xml, true);
        if (r) { xml = r.xml; steps.push(...r.steps); }
        else { lost.push(name); parts.delete(name); continue; }
      }
      if (xml !== original) { parts.set(name, enc(xml)); repaired.push(`${name} (${steps.join(', ') || 'normalized'})`); }
    }
    for (const m of xml.matchAll(/<c\b[^>]*?\ss="(\d+)"/g)) maxS = Math.max(maxS, +m[1]);
  }

  // 2. Shared strings
  const sstB = parts.get('xl/sharedStrings.xml');
  if (sstB) {
    const text = dec(sstB);
    const ok = !corrupt.has('xl/sharedStrings.xml') && isWellFormed(text);
    const { items, truncated } = parseSharedStrings(text);
    const declared = /<sst\b[^>]*\buniqueCount="(\d+)"/.exec(text)?.[1];
    if (!ok || force) {
      parts.set('xl/sharedStrings.xml', enc(rebuildSharedStrings(items)));
      repaired.push(`xl/sharedStrings.xml (rebuilt with ${plural(items.length, 'string')}${truncated ? '; the last one was truncated' : ''})`);
    } else if (declared != null && +declared !== items.length) {
      parts.set('xl/sharedStrings.xml', enc(text.replace(/(<sst\b[^>]*?)\s(?:count|uniqueCount)="\d+"/g, '$1').replace(/<sst\b/, `<sst count="${items.length}" uniqueCount="${items.length}"`)));
      repaired.push('xl/sharedStrings.xml (string count corrected)');
    }
  }

  // 3. Styles
  const stB = parts.get('xl/styles.xml');
  let stylesReplaced = false;
  let styles = stB ? dec(stB) : '';
  if (styles && (!isWellFormed(styles) || corrupt.has('xl/styles.xml'))) {
    const r = repairXmlPart('xl/styles.xml', styles, true);
    styles = r && stylesUsable(r.xml, maxS) ? r.xml : '';
    if (styles) { parts.set('xl/styles.xml', enc(styles)); repaired.push(`xl/styles.xml (${r!.steps.join(', ') || 'healed'})`); }
  }
  if (!styles || !stylesUsable(styles, maxS)) {
    parts.set('xl/styles.xml', enc(MINIMAL_STYLES));
    stylesReplaced = true;
    repaired.push(`xl/styles.xml (${stB ? 'unusable; ' : ''}replaced with a minimal stylesheet; cell formatting was dropped, values kept)`);
    for (const name of [...parts.keys()].filter(isSheetPart)) {
      const x = dec(parts.get(name)!);
      parts.set(name, enc(x.replace(/(<c\b[^>]*?)\ss="\d+"/g, '$1').replace(/(<row\b[^>]*?)\ss="\d+"/g, '$1').replace(/(<row\b[^>]*?)\scustomFormat="[^"]*"/g, '$1').replace(/(<col\b[^>]*?)\sstyle="\d+"/g, '$1')));
    }
  }

  // 4. Other XML parts: heal, or drop (dangling references are removed below).
  for (const [name, bytes] of [...parts]) {
    if (!/\.(xml|rels|vml)$/i.test(name) || isSheetPart(name) || name === 'xl/sharedStrings.xml' || name === 'xl/styles.xml') continue;
    const text = dec(bytes);
    if (!force && !corrupt.has(name) && isWellFormed(text)) continue;
    if (isWellFormed(text) && !corrupt.has(name)) continue;
    const r = repairXmlPart(name, text, true);
    if (r && isWellFormed(r.xml)) { if (r.xml !== text) { parts.set(name, enc(r.xml)); repaired.push(`${name} (${r.steps.join(', ') || 'healed'})`); } }
    else { parts.delete(name); lost.push(name); }
  }

  // 5. Workbook part and its relationships
  const liveSheets = [...parts.keys()].filter(isSheetPart).sort((a, b) => sheetNum(a) - sheetNum(b) || a.localeCompare(b));
  let order = workbookSheets(parts).filter((s) => s.part && parts.has(s.part));
  const wbOk = parts.has('xl/workbook.xml') && isWellFormed(dec(parts.get('xl/workbook.xml')!));
  const relsOk = parts.has('xl/_rels/workbook.xml.rels') && isWellFormed(dec(parts.get('xl/_rels/workbook.xml.rels')!));
  const orphans = liveSheets.filter((p) => !order.some((s) => s.part === p));
  if (!wbOk || !relsOk || orphans.length || order.length !== workbookSheets(parts).length) {
    // Keep the names we know; append orphan sheet parts; regenerate both parts consistently.
    if (!order.length && parts.has('xl/workbook.xml')) {
      // Rels lost: pair <sheet> entries with sheet parts by position.
      const names = [...dec(parts.get('xl/workbook.xml')!).matchAll(/<(?:\w+:)?sheet\b([^>]*?)\/?>/g)].map((m) => attrs(m[1]));
      order = liveSheets.slice(0, names.length).map((p, i) => ({ part: p, name: unent(names[i].name || ''), state: names[i].state }));
    }
    for (const p of liveSheets) if (!order.some((s) => s.part === p)) order.push({ part: p, name: '', state: undefined });
    const used = new Set<string>();
    order.forEach((s, i) => {
      let nm = (s.name || `Sheet${i + 1}`).replace(/[\\/?*[\]:]/g, '_').slice(0, 31);
      let k = 2;
      while (used.has(nm.toLowerCase())) nm = `${(s.name || 'Sheet').slice(0, 27)} (${k++})`;
      used.add(nm.toLowerCase()); s.name = nm;
    });
    const oldRels = relsOk ? parseRels(dec(parts.get('xl/_rels/workbook.xml.rels')!)) : [];
    const keep = oldRels.filter((r) => !/\/(worksheet|calcChain)$/.test(r.type) && (r.external || parts.has(resolveTarget('xl/_rels/workbook.xml.rels', r.target))));
    const rels: { id: string; type: string; target: string; external?: boolean }[] = keep.map((r) => ({ id: r.id, type: r.type, target: r.target, external: r.external }));
    const has = (t: string) => rels.some((r) => r.type.endsWith('/' + t));
    if (parts.has('xl/styles.xml') && !has('styles')) rels.push({ id: 'rIdQs', type: `${REL_NS}/styles`, target: 'styles.xml' });
    if (parts.has('xl/sharedStrings.xml') && !has('sharedStrings')) rels.push({ id: 'rIdQss', type: `${REL_NS}/sharedStrings`, target: 'sharedStrings.xml' });
    const theme = [...parts.keys()].find((n) => /^xl\/theme\/theme\d*\.xml$/.test(n));
    if (theme && !has('theme')) rels.push({ id: 'rIdQt', type: `${REL_NS}/theme`, target: theme.slice(3) });
    const sheetsXml = order.map((s, i) => {
      const id = `rIdQw${i + 1}`;
      rels.push({ id, type: `${REL_NS}/worksheet`, target: s.part.slice(3) });
      return `<sheet name="${xmlEsc(s.name)}" sheetId="${i + 1}"${s.state && s.state !== 'visible' ? ` state="${s.state}"` : ''} r:id="${id}"/>`;
    }).join('');
    let wbXml = wbOk ? dec(parts.get('xl/workbook.xml')!) : '';
    if (wbXml && /<sheets\b[\s\S]*?<\/sheets>|<sheets\s*\/>/.test(wbXml)) {
      wbXml = wbXml.replace(/<sheets\b[\s\S]*?<\/sheets>|<sheets\s*\/>/, `<sheets>${sheetsXml}</sheets>`);
      if (!/xmlns:r=/.test(wbXml)) wbXml = wbXml.replace(/<workbook\b/, `<workbook xmlns:r="${REL_NS}"`);
      // Defined names may point at sheets that no longer exist; the importer tolerates them, but external/pivot links must resolve.
      wbXml = wbXml.replace(/<(externalReferences|pivotCaches)\b[\s\S]*?<\/\1>/g, (m) => (keep.length === oldRels.length ? m : ''));
    } else {
      wbXml = `${XML_DECL}<workbook xmlns="${MAIN_NS}" xmlns:r="${REL_NS}"><bookViews><workbookView/></bookViews><sheets>${sheetsXml}</sheets></workbook>`;
    }
    parts.set('xl/workbook.xml', enc(wbXml));
    parts.set('xl/_rels/workbook.xml.rels', enc(relsXml(rels)));
    repaired.push(`xl/workbook.xml and its relationships (${wbOk && relsOk ? 'sheet list rebuilt' : 'regenerated'} from ${plural(order.length, 'worksheet part')})`);
  }

  // 6. Dangling relationships in every other .rels part, and the references to them.
  for (const relName of [...parts.keys()].filter((n) => n.endsWith('.rels') && n !== '_rels/.rels' && n !== 'xl/_rels/workbook.xml.rels')) {
    const owner = relName.replace(/_rels\/([^/]+)\.rels$/, '$1');
    const list = parseRels(dec(parts.get(relName)!));
    if (!parts.has(owner)) { parts.delete(relName); continue; }
    const dead = list.filter((r) => !r.external && !parts.has(resolveTarget(relName, r.target)));
    if (!dead.length) continue;
    parts.set(relName, enc(relsXml(list.filter((r) => !dead.includes(r)))));
    let ox = dec(parts.get(owner)!);
    for (const r of dead) {
      const id = r.id.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      ox = ox.replace(new RegExp(`<(\\w+:)?(drawing|legacyDrawing|legacyDrawingHF|picture|tablePart|control|oleObject)\\b[^>]*?\\b\\w+:id="${id}"[^>]*?(\\/>|>[\\s\\S]*?<\\/\\1?\\2>)`, 'g'), '');
    }
    ox = ox.replace(/<tableParts\b[^>]*>\s*<\/tableParts>/g, '').replace(/<tableParts\b([^>]*?)count="\d+"([^>]*)>([\s\S]*?)<\/tableParts>/g, (_m, a: string, b: string, inner: string) => `<tableParts${a}count="${(inner.match(/<tablePart\b/g) || []).length}"${b}>${inner}</tableParts>`);
    parts.set(owner, enc(ox));
    repaired.push(`${relName} (${plural(dead.length, 'dangling relationship')} removed: ${dead.map((d) => d.target).join(', ')})`);
  }

  // 7. Package structure
  const macro = parts.has('xl/vbaProject.bin');
  const ctOld = parts.get('[Content_Types].xml');
  const ctNew = contentTypes([...parts.keys()].filter((n) => n !== '[Content_Types].xml'), macro);
  if (!ctOld || force || !isWellFormed(dec(ctOld)) || [...parts.keys()].some((n) => (isSheetPart(n) || n === 'xl/workbook.xml') && !dec(ctOld).includes(`/${n}"`))) {
    parts.set('[Content_Types].xml', enc(ctNew));
    repaired.push(`[Content_Types].xml (${ctOld ? 'rebuilt' : 'regenerated'})`);
  }
  const rootRels = parts.get('_rels/.rels');
  if (!rootRels || !isWellFormed(dec(rootRels)) || !/xl\/workbook\.xml/.test(dec(rootRels))) {
    const extra = [parts.has('docProps/core.xml') ? { id: 'rId2', type: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', target: 'docProps/core.xml' } : null,
      parts.has('docProps/app.xml') ? { id: 'rId3', type: `${REL_NS}/extended-properties`, target: 'docProps/app.xml' } : null].filter(Boolean) as { id: string; type: string; target: string }[];
    parts.set('_rels/.rels', enc(relsXml([{ id: 'rId1', type: `${REL_NS}/officeDocument`, target: 'xl/workbook.xml' }, ...extra])));
    repaired.push(`_rels/.rels (${rootRels ? 'rebuilt' : 'regenerated'})`);
  }
  return { parts, repaired, lost, sheetOrder: order.length ? order : workbookSheets(parts), stylesReplaced };
}

async function zipParts(parts: Map<string, Uint8Array>, ods = false): Promise<Uint8Array> {
  const zip = new JSZip();
  if (ods) zip.file('mimetype', ODS_MIME, { compression: 'STORE' });
  const order = [...parts.keys()].filter((n) => !(ods && n === 'mimetype')).sort((a, b) => (a === '[Content_Types].xml' ? -1 : b === '[Content_Types].xml' ? 1 : 0));
  for (const n of order) zip.file(n, parts.get(n)!);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } });
}

/** Build a workbook directly from scanned worksheet XML. */
function rawXlsxBook(parts: Map<string, Uint8Array>, preps: Map<string, SheetPrep>, stats: { missingStrings: number; dropped: number }): Workbook | null {
  const sstB = parts.get('xl/sharedStrings.xml');
  const sst = sstB ? parseSharedStrings(dec(sstB)).items : [];
  const known = workbookSheets(parts).filter((s) => s.part && preps.has(s.part));
  const order = [...known];
  for (const p of [...preps.keys()].sort((a, b) => sheetNum(a) - sheetNum(b))) if (!order.some((s) => s.part === p)) order.push({ part: p, name: '' });
  if (!order.length) return null;
  const book = new Workbook(false);
  const used = new Set<string>();
  order.forEach((s, i) => {
    let nm = (s.name || `Sheet${i + 1}`).slice(0, 31);
    let k = 2;
    while (used.has(nm.toLowerCase())) nm = `${(s.name || 'Sheet').slice(0, 27)} (${k++})`;
    used.add(nm.toLowerCase());
    const sh = new Sheet(book, nm);
    if (s.state === 'hidden' || s.state === 'veryHidden') sh.visibility = s.state;
    const prep = preps.get(s.part)!;
    for (const { r, c, cell } of scanSheetCells(prep.main, sst, stats)) sh.setComputed(r, c, cell);
    // Merges and frozen panes survive when the end of the part does.
    for (const m of prep.main.matchAll(/<mergeCell\b[^>]*?ref="([A-Z]+\d+):([A-Z]+\d+)"/g)) {
      const a = parseA1(m[1]), b = parseA1(m[2]);
      if (a && b) sh.merges.push({ r1: a.r, c1: a.c, r2: b.r, c2: b.c });
    }
    const pane = /<pane\b([^>]*?)\/?>/.exec(prep.main);
    if (pane) { const a = attrs(pane[1]); if (a.state === 'frozen' || a.state === 'frozenSplit') sh.freeze = { rows: +(a.ySplit || 0), cols: +(a.xSplit || 0) }; }
    book.sheets.push(sh);
  });
  if (/<workbookPr\b[^>]*date1904="(1|true)"/.test(parts.has('xl/workbook.xml') ? dec(parts.get('xl/workbook.xml')!) : '')) book.calc.date1904 = true;
  return book;
}

// ------------------------------------------------------------------------------------------
// OpenDocument

function odfFormula(f: string): string {
  let s = f.replace(/^(?:of|oooc|msoxl):/, '').replace(/^=/, '');
  s = s.replace(/\[\$?([^\].]*?)\.(\$?[A-Z]+\$?\d+)(?::\$?([^\].]*?)\.(\$?[A-Z]+\$?\d+))?\]/g, (_m, sh1: string, a: string, sh2: string | undefined, b: string | undefined) => {
    const pre = sh1 ? `${quoteName(sh1.replace(/^'|'$/g, '').replace(/^\$/, ''))}!` : '';
    return b ? `${pre}${a}:${b}` : `${pre}${a}`;
  });
  return s.replace(/;/g, ',');
}
const quoteName = (n: string) => (/^[A-Za-z_][\w.]*$/.test(n) ? n : `'${n.replace(/'/g, "''")}'`);

function odsDuration(s: string): number | null {
  const m = /^-?P(?:(\d+)D)?T?(?:(\d+)H)?(?:(\d+)M)?(?:([\d.]+)S)?$/.exec(s);
  if (!m) return null;
  return (+(m[1] || 0)) + ((+(m[2] || 0)) * 3600 + (+(m[3] || 0)) * 60 + (+(m[4] || 0))) / 86400;
}

function odsCellText(inner: string): string {
  const paras = [...inner.matchAll(/<text:p\b[^>]*?(?:\/>|>([\s\S]*?)(?:<\/text:p>|$))/g)].map((m) => m[1] ?? '');
  const src = paras.length ? paras : [inner];
  return src.map((p) => unent(p.replace(/<text:s\b([^>]*?)\/>/g, (_m, a: string) => ' '.repeat(+(attrs(a)['text:c'] || 1))).replace(/<text:tab\s*\/>/g, '\t').replace(/<text:line-break\s*\/>/g, '\n').replace(/<[^>]*>?/g, ''))).join('\n');
}

/** Raw scan of OpenDocument table XML (content.xml or a flat .fods). */
export function scanOdsTables(xml: string): Workbook | null {
  const book = new Workbook(false);
  const tables = xml.split(/<table:table\b(?![-\w])/).slice(1);
  tables.forEach((chunk, ti) => {
    const head = /^([^>]*)>/.exec(chunk);
    const a = attrs(head?.[1] || '');
    const name = unent(a['table:name'] || `Sheet${ti + 1}`).slice(0, 31);
    const body = chunk.slice(head ? head[0].length : 0).split('</table:table>')[0];
    const sh = new Sheet(book, book.sheetByName(name) ? `${name} (${ti + 1})` : name);
    let r = 0;
    for (const rm of body.matchAll(/<table:table-row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/table:table-row>)/g)) {
      const ra = attrs(rm[1]);
      const rrep = Math.max(1, +(ra['table:number-rows-repeated'] || 1));
      const cells: { c: number; cell: Cell }[] = [];
      let c = 0;
      for (const cm of (rm[2] || '').matchAll(/<table:(table-cell|covered-table-cell)\b([^>]*?)(?:\/>|>([\s\S]*?)<\/table:\1>)/g)) {
        const ca = attrs(cm[2]);
        const crep = Math.max(1, +(ca['table:number-columns-repeated'] || 1));
        const type = ca['office:value-type'];
        let v: Scalar = null;
        if (type === 'float' || type === 'percentage' || type === 'currency') v = Number(ca['office:value']);
        else if (type === 'date') v = isoToSerial(ca['office:date-value'] || '');
        else if (type === 'time') v = odsDuration(ca['office:time-value'] || '');
        else if (type === 'boolean') v = ca['office:boolean-value'] === 'true';
        else if (cm[3]) v = odsCellText(cm[3]);
        if (typeof v === 'number' && !Number.isFinite(v)) v = null;
        if (v === '') v = null;
        const cell: Cell = { v };
        if (ca['table:formula']) cell.f = odfFormula(unent(ca['table:formula']));
        if (v != null || cell.f) for (let k = 0; k < Math.min(crep, 256); k++) cells.push({ c: c + k, cell: { ...cell } });
        c += crep;
        const span = +(ca['table:number-columns-spanned'] || 1), rspan = +(ca['table:number-rows-spanned'] || 1);
        if (cm[1] === 'table-cell' && (span > 1 || rspan > 1)) sh.merges.push({ r1: r, c1: c - crep, r2: r + rspan - 1, c2: c - crep + span - 1 });
      }
      const times = cells.length ? Math.min(rrep, 1000) : 0;
      for (let k = 0; k < times; k++) for (const x of cells) if (r + k < MAX_ROWS && x.c < MAX_COLS) sh.setComputed(r + k, x.c, { ...x.cell });
      r += rrep;
      if (r >= MAX_ROWS) break;
    }
    book.sheets.push(sh);
  });
  return book.sheets.length ? book : null;
}

// ------------------------------------------------------------------------------------------
// Text formats

function detectDelim(lines: string[]): string {
  const cands = ['\t', ';', ',', '|'];
  let best = ',', bestScore = 0;
  for (const d of cands) {
    const counts = lines.slice(0, 50).map((l) => l.split(d).length - 1);
    const nonzero = counts.filter((n) => n > 0);
    if (!nonzero.length) continue;
    const mode = [...new Set(nonzero)].sort((a, b) => nonzero.filter((x) => x === b).length - nonzero.filter((x) => x === a).length)[0];
    const score = nonzero.filter((x) => x === mode).length / counts.length * (1 + mode / 100);
    if (score > bestScore) { bestScore = score; best = d; }
  }
  return best;
}

function splitDelimited(text: string, delim: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [], cur = '', q = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (q) {
      if (ch === '"') { if (text[i + 1] === '"') { cur += '"'; i++; } else q = false; }
      else cur += ch;
    } else if (ch === '"' && cur === '') q = true;
    else if (ch === delim) { row.push(cur); cur = ''; }
    else if (ch === '\n' || ch === '\r') { if (ch === '\r' && text[i + 1] === '\n') i++; row.push(cur); rows.push(row); row = []; cur = ''; }
    else cur += ch;
  }
  if (cur || row.length) { row.push(cur); rows.push(row); }
  return rows;
}

function typed(s: string): Scalar {
  const t = s.trim();
  if (!t) return null;
  if (/^[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?$/.test(t)) return Number(t);
  if (/^(TRUE|FALSE)$/i.test(t)) return t.toUpperCase() === 'TRUE';
  return s;
}

function gridBook(rows: string[][], name: string): Workbook {
  const book = new Workbook(false);
  const sh = new Sheet(book, name);
  rows.forEach((row, r) => row.forEach((s, c) => { const v = typed(s); if (v != null && r < MAX_ROWS && c < MAX_COLS) sh.setComputed(r, c, { v }); }));
  book.sheets.push(sh);
  return book;
}

function htmlBook(text: string): Workbook | null {
  const doc = new DOMParser().parseFromString(text, 'text/html');
  const tables = [...doc.querySelectorAll('table')];
  if (!tables.length) return null;
  const book = new Workbook(false);
  tables.forEach((t, i) => {
    const sh = new Sheet(book, `Table${i + 1}`);
    const taken = new Set<number>();
    [...t.querySelectorAll('tr')].forEach((tr, r) => {
      let c = 0;
      for (const td of [...tr.children].filter((e) => /^(TD|TH)$/.test(e.tagName))) {
        while (taken.has(cellKey(r, c))) c++;
        const v = typed((td.textContent || '').replace(/\s+/g, ' ').trim());
        if (v != null) sh.setComputed(r, c, { v });
        const cs = +(td.getAttribute('colspan') || 1), rs = +(td.getAttribute('rowspan') || 1);
        if (cs > 1 || rs > 1) { sh.merges.push({ r1: r, c1: c, r2: r + rs - 1, c2: c + cs - 1 }); for (let a = 0; a < rs; a++) for (let b = 0; b < cs; b++) taken.add(cellKey(r + a, c + b)); }
        c += cs;
      }
    });
    book.sheets.push(sh);
  });
  return book;
}

function r1c1ToA1(f: string, r: number, c: number): string {
  return f.split(/("(?:[^"]|"")*")/).map((p, i) => i % 2 ? p : p.replace(/(^|[^A-Za-z0-9_])R(\[-?\d+\]|\d+)?C(\[-?\d+\]|\d+)?(?![A-Za-z0-9_(])/g, (_m, pre: string, rr?: string, cc?: string) => {
    const conv = (x: string | undefined, base: number): [number, boolean] => !x ? [base, false] : x.startsWith('[') ? [base + +x.slice(1, -1), false] : [+x - 1, true];
    const [row, ar] = conv(rr, r), [col, ac] = conv(cc, c);
    if (row < 0 || col < 0) return `${pre}#REF!`;
    return `${pre}${ac ? '$' : ''}${colName(col)}${ar ? '$' : ''}${row + 1}`;
  })).join('');
}

/** Excel 2003 XML (SpreadsheetML), healed first when damaged. */
function xml2003Book(text: string): Workbook | null {
  const xml = isWellFormed(text) ? text : balanceXml(text);
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length) return null;
  const sheets = [...doc.getElementsByTagName('*')].filter((e) => e.localName === 'Worksheet');
  if (!sheets.length) return null;
  const book = new Workbook(false);
  const ss = (e: Element, n: string) => e.getAttribute(`ss:${n}`) ?? e.getAttribute(n);
  sheets.forEach((ws, i) => {
    const sh = new Sheet(book, (ss(ws, 'Name') || `Sheet${i + 1}`).slice(0, 31));
    let r = -1;
    for (const row of [...ws.getElementsByTagName('*')].filter((e) => e.localName === 'Row')) {
      const ri = ss(row, 'Index'); r = ri ? +ri - 1 : r + 1;
      let c = -1;
      for (const cell of [...row.children].filter((e) => e.localName === 'Cell')) {
        const ci = ss(cell, 'Index'); c = ci ? +ci - 1 : c + 1;
        const data = [...cell.children].find((e) => e.localName === 'Data');
        const type = data ? ss(data, 'Type') : null;
        const t = data?.textContent ?? '';
        let v: Scalar = null;
        if (type === 'Number') v = Number(t);
        else if (type === 'Boolean') v = t === '1' || t.toLowerCase() === 'true';
        else if (type === 'DateTime') v = isoToSerial(t) ?? t;
        else if (type === 'Error') v = errorFromCode(t) || new CellError('#N/A');
        else if (data) v = t;
        const cellObj: Cell = { v: typeof v === 'number' && !Number.isFinite(v) ? null : v };
        const f = ss(cell, 'Formula');
        if (f) cellObj.f = r1c1ToA1(f.replace(/^=/, ''), r, c);
        if (cellObj.v != null || cellObj.f) sh.setComputed(r, c, cellObj);
        const across = +(ss(cell, 'MergeAcross') || 0), down = +(ss(cell, 'MergeDown') || 0);
        if (across || down) sh.merges.push({ r1: r, c1: c, r2: r + down, c2: c + across });
        c += across;
      }
    }
    book.sheets.push(sh);
  });
  return book;
}

/** Strings-scan lines: tabular lines become columns. */
function linesBook(lines: string[]): Workbook {
  const delim = detectDelim(lines);
  const withDelim = lines.filter((l) => l.includes(delim)).length;
  const rows = withDelim >= Math.max(2, lines.length * 0.6) ? lines.map((l) => (delim === ',' || delim === ';' ? splitDelimited(l, delim)[0] || [l] : l.split(delim))) : lines.map((l) => [l]);
  const book = gridBook(rows, 'Recovered text');
  // Keep every line as text (no type guessing on fragments that might be partial numbers).
  return book;
}

// ------------------------------------------------------------------------------------------
// Orchestrator

interface Candidate { id: MethodId; book: Workbook; weight: number; corrupt?: boolean }

export async function recoverWorkbook(data: ArrayBuffer, name: string, opts: RecoverOptions = {}): Promise<RecoveryReport> {
  const progress = (m: string) => { try { opts.onProgress?.(m); } catch { /* UI callback errors are not ours */ } };
  const fileName = name || 'workbook';
  const base = fileName.replace(/\.[^.]+$/, '');
  const report: RecoveryReport = {
    fileName, kind: 'unknown', format: '', book: null, methods: [], identification: { ext: 'bin', description: 'Unrecognized data', confidence: 'low' },
    foreign: [], recommended: [], summary: '', bestMethod: null, notes: [], warnings: [], stats: { sheets: 0, cells: 0, formulas: 0 },
  };
  const M = (id: MethodId): MethodResult => { const m: MethodResult = { id, name: METHOD_INFO[id][0], desc: METHOD_INFO[id][1], status: 'skipped', detail: '' }; report.methods.push(m); return m; };
  const cands: Candidate[] = [];
  const addCand = (m: MethodResult, id: MethodId, book: Workbook | null, weight: number, corrupt = false) => {
    if (!book) return;
    if (!book.sheets.length) book.sheets.push(new Sheet(book, 'Sheet1'));
    const st = bookStats(book);
    m.cells = st.cells; m.sheets = st.sheets;
    cands.push({ id, book, weight, corrupt });
  };
  try {
    const raw = new Uint8Array(data);
    // ---- a) identify
    progress('Identifying the file…');
    const mId = M('identify');
    if (!raw.length) {
      mId.status = 'failed'; mId.detail = 'The file is empty (0 bytes).';
      report.summary = 'The file is empty (0 bytes); there is nothing to recover.';
      return finish(report, cands, base);
    }
    let segs: Segment[] = [];
    try { segs = scan(raw).segments; } catch { segs = []; }
    let u8: Uint8Array = raw;
    let det = detectKind(raw, fileName);
    if (!SPREADSHEET_KINDS.includes(det.kind) && segs.length > 1) {
      for (const s of segs) {
        const bytes = s.ext === 'ole' || s.ext === 'xls' ? raw.subarray(s.offset) : s.bytes;
        const d = detectKind(bytes, fileName);
        if (SPREADSHEET_KINDS.includes(d.kind)) { det = d; u8 = bytes; report.notes.push(`Repair ran on the ${d.id.description} found at byte ${s.offset.toLocaleString('en-US')}.`); break; }
      }
    } else if (segs.length > 1 && (det.kind === 'xlsx' || det.kind === 'xlsm' || det.kind === 'ods')) {
      u8 = segs[0].bytes;
    }
    report.kind = det.kind;
    report.identification = det.id;
    report.format = det.kind === 'xlsm' ? 'xlsm' : det.kind === 'encrypted' || det.kind === 'other' || det.kind === 'unknown' ? '' : det.kind === 'xml2003' ? 'xml' : det.kind;
    if (segs.length > 1) {
      report.foreign = segs.filter((s) => !(u8.byteOffset - raw.byteOffset <= s.offset && s.offset < u8.byteOffset - raw.byteOffset + u8.length)).map((s) => ({ ...s, recommended: recommend(s.ext) }));
      if (report.foreign.length) report.notes.push(`The file contains ${plural(segs.length, 'separate piece')}; ${plural(report.foreign.length, 'piece')} of another kind ${report.foreign.length === 1 ? 'was' : 'were'} set aside.`);
    }
    const ext = /\.([A-Za-z0-9]+)$/.exec(fileName)?.[1]?.toLowerCase() || '';
    const extKind = ext === 'xlsx' || ext === 'xlsm' || ext === 'xltx' ? 'xlsx' : ext === 'xls' ? 'xls' : ext === 'ods' ? 'ods' : ext;
    const realBase = det.kind === 'xlsm' ? 'xlsx' : det.kind;
    if (ext && SPREADSHEET_KINDS.includes(det.kind) && extKind && realBase !== extKind && !(extKind === 'csv' && det.kind === 'csv') && ['xlsx', 'xls', 'ods'].includes(extKind)) {
      report.notes.push(`The file is named .${ext} but is really a ${det.id.description}; it was treated as what it is.`);
    }
    mId.status = det.kind === 'unknown' ? 'partial' : 'success';
    mId.detail = `${det.id.description} (${det.id.confidence} confidence)${segs.length > 1 ? `; ${plural(segs.length, 'piece')} found in the file` : ''}.`;
    report.recommended = recommend(det.kind === 'xlsm' ? 'xlsx' : det.kind === 'fods' ? 'ods' : det.kind === 'encrypted' ? 'xlsx' : det.kind === 'other' || det.kind === 'unknown' ? det.id.ext : det.kind);
    if (det.kind === 'encrypted') {
      mId.detail = 'This is a password-protected (encrypted) Office file: an encrypted package inside a compound file.';
      report.notes.push('The workbook is encrypted with a password. Open it with the password in the program that created it, remove the password (File > Info > Protect Workbook > Encrypt with Password), and save it again.');
      for (const id of ['open', 'zip', 'xml', 'rawsheet', 'biff', 'ods', 'text', 'strings'] as MethodId[]) { const m = M(id); m.detail = 'Not possible: the content is encrypted.'; }
      report.summary = 'This workbook is password-protected (encrypted); Anne\'s Grid cannot decrypt it.';
      return finish(report, cands, base);
    }
    if (det.kind === 'other') {
      report.notes.push(`This is not a spreadsheet: it looks like a ${det.id.description}.${report.recommended[0] ? ` ${report.recommended[0].name} may be able to repair it.` : ''}`);
    }
    const kind = det.kind;
    const zipKind = kind === 'xlsx' || kind === 'xlsm' || kind === 'ods' || (kind === 'unknown' && hasAscii(u8, 'PK\x03\x04'));
    const textKind = kind === 'csv' || kind === 'html' || kind === 'xml2003' || kind === 'fods' || kind === 'qgw';

    // ---- b) normal open
    progress('Opening the file normally…');
    const mOpen = M('open');
    let openOk = false;
    {
      const w: string[] = [];
      let book: Workbook | null = null;
      let err = '';
      try {
        if (kind === 'xlsx' || kind === 'xlsm') {
          const mod = await loadMod('xlsx-import');
          if (!mod?.importXlsx) err = 'The XLSX importer is not available.';
          else book = asBook(await mod.importXlsx(ab(u8), fileName, w), w);
        } else if (kind === 'ods') {
          const mod = await loadMod('ods');
          if (!mod?.importOds) err = 'The ODS importer is not available.';
          else book = asBook(await mod.importOds(ab(u8), w), w);
        } else if (kind === 'fods') {
          const mod = await loadMod('ods');
          book = mod?.importFods ? asBook(await mod.importFods(dec(u8), w), w) : null;
          if (!book) err = 'The flat ODS importer is not available.';
        } else if (kind === 'xls') {
          const r = readBiff(u8, { lenient: false });
          book = r.book; w.push(...r.warnings);
        } else if (kind === 'csv') {
          const mod = await loadMod('csv');
          if (mod?.importCsv) book = asBook(await mod.importCsv(ab(u8), fileName, {}), w);
          else { const t = dec(u8); book = gridBook(splitDelimited(t, detectDelim(t.split(/\r?\n/))), base.slice(0, 31) || 'Sheet1'); }
        } else if (kind === 'html') {
          const mod = await loadMod('html');
          book = mod?.importHtml ? asBook(await mod.importHtml(dec(u8), w), w) : htmlBook(dec(u8));
        } else if (kind === 'xml2003') {
          const t = dec(u8);
          if (!isWellFormed(t)) throw new Error('The XML is not well-formed.');
          const mod = await loadMod('other');
          book = mod?.importXml2003 ? asBook(await mod.importXml2003(t, w), w) : xml2003Book(t);
        } else if (kind === 'qgw') {
          const mod = await loadMod('native');
          const t = dec(u8);
          book = mod?.importJson ? asBook(await mod.importJson(t), w) : Workbook.fromJSON(JSON.parse(t));
        } else { mOpen.detail = 'No importer for this kind of file.'; }
      } catch (e: any) {
        err = e?.message === XLS_ENCRYPTED_MSG ? XLS_ENCRYPTED_MSG : (e?.message || String(e));
      }
      if (book) {
        addCand(mOpen, 'open', book, 1.05);
        openOk = (mOpen.cells ?? 0) > 0 || !zipKind;
        if (zipKind && w.length) openOk = false;
        mOpen.status = w.length ? 'partial' : 'success';
        mOpen.detail = `Opened: ${plural(mOpen.sheets || 0, 'sheet')}, ${plural(mOpen.cells || 0, 'cell')}${w.length ? `; ${w.length} warning${w.length === 1 ? '' : 's'}: ${w.slice(0, 3).join(' ')}` : ''}.`;
        report.warnings.push(...w);
        if (!(mOpen.cells ?? 0) && zipKind) { mOpen.status = 'partial'; openOk = false; }
      } else if (err) { mOpen.status = 'failed'; mOpen.detail = err; }
      if (err === XLS_ENCRYPTED_MSG) {
        report.notes.push('This .xls is protected with a password to open (encrypted). Remove the password in Excel and save it again.');
        report.summary = XLS_ENCRYPTED_MSG;
        for (const id of ['zip', 'xml', 'rawsheet', 'biff', 'ods', 'text', 'strings'] as MethodId[]) { const m = M(id); m.detail = 'Not possible: the workbook is encrypted.'; }
        return finish(report, cands, base);
      }
    }
    let salv: Salvage | null = null;
    // A lenient importer can "open" a damaged package and quietly lose data: check the parts first.
    if (zipKind && cands.some((c) => c.id === 'open')) {
      salv = await salvageZip(u8, !!opts.forceRepair);
      const bad = [...salv.corrupt];
      for (const [n, b] of salv.parts) if (!bad.includes(n) && /\.(xml|rels)$/i.test(n) && b.length < 8 << 20 && !isWellFormed(dec(b))) bad.push(n);
      if (bad.length) {
        openOk = false;
        const oc = cands.find((c) => c.id === 'open')!;
        oc.weight = 0.5; oc.corrupt = true;
        mOpen.status = 'partial';
        mOpen.detail += ` However, ${plural(bad.length, 'part')} ${bad.length === 1 ? 'is' : 'are'} damaged (${bad.slice(0, 4).join(', ')}), so the repair methods ran too.`;
      }
    }
    const skipRest = openOk && !opts.forceRepair;
    const notNeeded = 'Not needed: the file opened normally.';

    // ---- c) ZIP salvage
    const mZip = M('zip');
    if (!zipKind) mZip.detail = 'Not a ZIP-based package.';
    else if (skipRest) mZip.detail = notNeeded;
    else {
      progress('Salvaging the ZIP package…');
      if (!salv) salv = await salvageZip(u8, !!opts.forceRepair);
      if (!salv.parts.size) { mZip.status = 'failed'; mZip.detail = `No package parts could be found${salv.zipError ? ` (${salv.zipError})` : ''}.`; }
      else {
        mZip.status = salv.corrupt.size ? 'partial' : 'success';
        mZip.detail = `${plural(salv.parts.size, 'part')} recovered${salv.zipOk ? '' : ' from local headers (the central directory is missing or damaged)'}${salv.salvaged.length && salv.zipOk ? `; ${plural(salv.salvaged.length, 'part')} re-read from raw bytes` : ''}${salv.corrupt.size ? `; damaged or truncated: ${[...salv.corrupt].slice(0, 6).join(', ')}` : '; every CRC-32 matches'}.`;
      }
    }
    const isOdsPkg = !!salv && (salv.parts.has('content.xml') && ![...salv.parts.keys()].some(isSheetPart));

    // ---- d) XML repair and e) raw sheet scan (XLSX)
    const mXml = M('xml');
    const mRaw = M('rawsheet');
    const preps = new Map<string, SheetPrep>();
    if (!salv || isOdsPkg || !salv.parts.size) {
      const why = skipRest ? notNeeded : !zipKind ? 'Not an Office Open XML package.' : isOdsPkg ? 'Not an Office Open XML package (OpenDocument).' : 'No package parts to repair.';
      mXml.detail = why; mRaw.detail = why;
    } else {
      for (const n of [...salv.parts.keys()].filter(isSheetPart)) preps.set(n, prepareSheet(dec(salv.parts.get(n)!), salv.corrupt.has(n)));
      const cutParts = [...preps].filter(([, p]) => p.cut);
      if (!preps.size) { mXml.status = 'failed'; mXml.detail = 'No worksheet parts (xl/worksheets/sheet*.xml) were recovered.'; mRaw.status = 'failed'; mRaw.detail = mXml.detail; }
      else {
        progress('Repairing the XML parts…');
        let rep: XlsxRepair | null = null;
        try { rep = repairXlsx(salv.parts, salv.corrupt, preps, !!opts.forceRepair); }
        catch (e: any) { mXml.status = 'failed'; mXml.detail = `The package could not be rebuilt (${e?.message || e}).`; }
        if (rep) {
          const pkg = await zipParts(rep.parts);
          report.repaired = {
            name: `${base} (repaired).${kind === 'xlsm' ? 'xlsm' : 'xlsx'}`, blob: new Blob([pkg as BlobPart], { type: XLSX_MIME }), kind: 'xlsx',
            repairedParts: rep.repaired, salvagedParts: salv.salvaged, lostParts: rep.lost,
          };
          if (rep.stylesReplaced) report.notes.push('The stylesheet could not be repaired; cell values were kept but cell formatting (including date formats) was dropped.');
          const mod = await loadMod('xlsx-import');
          const w: string[] = [];
          let book: Workbook | null = null;
          if (mod?.importXlsx) {
            progress('Opening the repaired package…');
            try { book = asBook(await mod.importXlsx(ab(pkg), fileName, w), w); }
            catch (e: any) { mXml.detail = `The repaired package was rebuilt but could not be opened (${e?.message || e}).`; }
          } else mXml.detail = 'The package was rebuilt, but the XLSX importer is not available to open it.';
          const changed = rep.repaired.length + rep.lost.length;
          if (book) {
            addCand(mXml, 'xml', book, 1.03, salv.corrupt.size > 0);
            mXml.status = (mXml.cells ?? 0) > 0 ? (rep.lost.length || cutParts.length ? 'partial' : 'success') : 'failed';
            mXml.detail = `${changed ? `Repaired ${plural(rep.repaired.length, 'part')}` : 'No repairs were needed'}${rep.lost.length ? `; could not repair: ${rep.lost.join(', ')}` : ''}; opened ${plural(mXml.cells || 0, 'cell')} in ${plural(mXml.sheets || 0, 'sheet')}.`;
          } else mXml.status = 'failed';
        }
        // e) raw scan
        progress('Scanning worksheet XML for cells…');
        const st = { missingStrings: 0, dropped: 0 };
        const book = rawXlsxBook(salv.parts, preps, st);
        if (book) {
          addCand(mRaw, 'rawsheet', book, 1.0, salv.corrupt.size > 0);
          mRaw.status = (mRaw.cells ?? 0) > 0 ? (st.missingStrings || st.dropped || cutParts.length ? 'partial' : 'success') : 'failed';
          mRaw.detail = `Found ${plural(mRaw.cells || 0, 'cell')} in ${plural(mRaw.sheets || 0, 'sheet')}${st.missingStrings ? `; ${plural(st.missingStrings, 'cell')} refer to shared strings that were lost` : ''}${st.dropped ? `; ${plural(st.dropped, 'damaged cell')} skipped` : ''}.`;
        } else { mRaw.status = 'failed'; mRaw.detail = 'No worksheet XML was found.'; }
        if (cutParts.length) report.notes.push(`${cutParts.map(([n]) => n).join(', ')} ${cutParts.length === 1 ? 'is' : 'are'} damaged: the rows up to the last complete row before the damage were kept.`);
      }
    }

    // ---- f) BIFF salvage
    const mBiff = M('biff');
    if (kind !== 'xls') mBiff.detail = 'Not an Excel 97-2003 file.';
    else if (skipRest) mBiff.detail = notNeeded;
    else {
      progress('Reading Excel 97-2003 records leniently…');
      try {
        const r = readBiff(u8, { lenient: true });
        const s = r.stats;
        addCand(mBiff, 'biff', r.book, 1.0, s.resyncs > 0 || s.source === 'scan');
        report.warnings.push(...r.warnings);
        mBiff.status = s.cells ? (s.resyncs || s.source === 'scan' || r.warnings.length ? 'partial' : 'success') : 'failed';
        mBiff.detail = `BIFF${s.biff} via ${s.source === 'cfb' ? 'the compound file' : s.source === 'scan' ? 'a raw-byte scan (the compound file is damaged)' : 'the raw stream'}: ${plural(s.cells, 'cell')}, ${plural(s.formulasDecoded, 'formula')} decoded of ${s.formulas}, ${plural(s.sharedStrings, 'shared string')}${s.resyncs ? `; resynchronized ${plural(s.resyncs, 'time')} (${plural(s.skippedBytes, 'byte')} skipped)` : ''}.`;
      } catch (e: any) {
        mBiff.status = 'failed';
        mBiff.detail = e?.message || String(e);
        if (e?.message === XLS_ENCRYPTED_MSG) report.notes.push('This .xls is encrypted with a password.');
      }
    }

    // ---- g) ODS
    const mOds = M('ods');
    if (kind !== 'ods' && kind !== 'fods' && !isOdsPkg) mOds.detail = 'Not an OpenDocument spreadsheet.';
    else if (skipRest) mOds.detail = notNeeded;
    else if (kind === 'fods') {
      progress('Scanning flat OpenDocument XML…');
      const book = scanOdsTables(stripInvalidChars(dec(u8)));
      if (book) { addCand(mOds, 'ods', book, 1.0); mOds.status = (mOds.cells ?? 0) ? 'partial' : 'failed'; mOds.detail = `Found ${plural(mOds.cells || 0, 'cell')} in ${plural(mOds.sheets || 0, 'table')}.`; }
      else { mOds.status = 'failed'; mOds.detail = 'No tables found.'; }
    } else {
      if (!salv) salv = await salvageZip(u8, true);
      const cb = salv.parts.get('content.xml');
      if (!cb) { mOds.status = 'failed'; mOds.detail = 'content.xml was not recovered.'; }
      else {
        progress('Repairing OpenDocument content…');
        let text = stripInvalidChars(dec(cb));
        let tail = '';
        if (salv.corrupt.has('content.xml') || !isWellFormed(text)) {
          let cut = findDamageCut(text, '</table:table-row>');
          if (cut < 0 && !text.includes('</office:spreadsheet>')) { const l = text.lastIndexOf('</table:table-row>'); if (l > 0) cut = l + 18; }
          if (cut > 0) { tail = text.slice(cut); text = text.slice(0, cut); }
        }
        const steps: string[] = [];
        const fixed = isWellFormed(text) ? { xml: text, steps: [] as string[] } : repairXmlPart('content.xml', text, true);
        let book: Workbook | null = null;
        if (fixed) {
          steps.push(...fixed.steps);
          const parts = new Map(salv.parts);
          parts.set('content.xml', enc(fixed.xml));
          for (const [n, b] of [...parts]) if (n !== 'content.xml' && /\.xml$/.test(n) && !isWellFormed(dec(b))) { const r = repairXmlPart(n, dec(b), true); if (r) parts.set(n, enc(r.xml)); else parts.delete(n); }
          if (!parts.has('META-INF/manifest.xml')) {
            const entries = [...parts.keys()].filter((n) => n !== 'mimetype' && !n.startsWith('META-INF/')).map((n) => `<manifest:file-entry manifest:full-path="${xmlEsc(n)}" manifest:media-type="${n.endsWith('.xml') ? 'text/xml' : ''}"/>`);
            parts.set('META-INF/manifest.xml', enc(`<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2"><manifest:file-entry manifest:full-path="/" manifest:media-type="${ODS_MIME}"/>${entries.join('')}</manifest:manifest>`));
            steps.push('manifest regenerated');
          }
          const pkg = await zipParts(parts, true);
          report.repaired = { name: `${base} (repaired).ods`, blob: new Blob([pkg as BlobPart], { type: ODS_MIME }), kind: 'ods', repairedParts: steps.length || tail ? [`content.xml (${[tail ? 'cut at the last complete row before the damage' : '', ...steps].filter(Boolean).join(', ')})`] : [], salvagedParts: salv.salvaged, lostParts: [] };
          const mod = await loadMod('ods');
          if (mod?.importOds) {
            const w: string[] = [];
            try { book = asBook(await mod.importOds(ab(pkg), w), w); } catch { book = null; }
          }
        }
        if (book) addCand(mOds, 'ods', book, 1.02, salv.corrupt.size > 0);
        const rawBook = scanOdsTables(text);
        const rawCells = bookStats(rawBook).cells;
        if (rawBook && rawCells > (mOds.cells ?? 0)) { const m2 = { ...mOds }; addCand(m2, 'ods', rawBook, 1.0, salv.corrupt.size > 0); mOds.cells = m2.cells; mOds.sheets = m2.sheets; }
        if (tail) { const tb = scanOdsTables(`<table:table table:name="Unverified">${tail}`); if (tb && bookStats(tb).cells) report.notes.push(`${plural(bookStats(tb).cells, 'cell')} decoded after the damaged area of content.xml ${bookStats(tb).cells === 1 ? 'was' : 'were'} not added because they cannot be verified.`); }
        mOds.status = (mOds.cells ?? 0) ? (tail || salv.corrupt.size || steps.length ? 'partial' : 'success') : 'failed';
        mOds.detail = (mOds.cells ?? 0) ? `Recovered ${plural(mOds.cells || 0, 'cell')} in ${plural(mOds.sheets || 0, 'sheet')}${steps.length ? ` (${steps.join(', ')})` : ''}.` : 'No table cells could be recovered from content.xml.';
      }
    }

    // ---- h) text fallbacks
    const mText = M('text');
    if (!textKind && kind !== 'unknown') mText.detail = 'Not a text-based format.';
    else if (skipRest) mText.detail = notNeeded;
    else {
      progress('Reading the file as text…');
      const t = stripInvalidChars(dec(u8));
      let book: Workbook | null = null;
      try {
        if (kind === 'xml2003') book = xml2003Book(t);
        else if (kind === 'html') book = htmlBook(t);
        else if (kind === 'fods') book = null;
        else if (kind === 'qgw') { try { book = Workbook.fromJSON(JSON.parse(t)); } catch { book = null; } }
        else if (kind === 'csv' || (kind === 'unknown' && looksTextual(u8))) { const lines = t.split(/\r?\n/); book = gridBook(splitDelimited(t, detectDelim(lines)), base.slice(0, 31) || 'Sheet1'); }
      } catch { book = null; }
      if (book) { addCand(mText, 'text', book, 1.0); mText.status = (mText.cells ?? 0) ? 'success' : 'failed'; mText.detail = `Found ${plural(mText.cells || 0, 'cell')} in ${plural(mText.sheets || 0, 'sheet')}.`; }
      else { mText.status = 'failed'; mText.detail = kind === 'unknown' ? 'The data is not readable text.' : 'No tabular data was found.'; }
    }

    // ---- i) strings scan
    const mStr = M('strings');
    const bestCells = Math.max(0, ...cands.map((c) => bookStats(c.book).cells));
    if (bestCells > 0) mStr.detail = 'Not needed: structured data was recovered.';
    else {
      progress('Scanning for readable text…');
      let text = '';
      if (salv && salv.parts.size) {
        const pieces: string[] = [];
        const sst = salv.parts.get('xl/sharedStrings.xml');
        if (sst) pieces.push(...parseSharedStrings(dec(sst)).items);
        for (const [n, b] of salv.parts) if (/\.xml$/.test(n) && n !== 'xl/sharedStrings.xml' && !/styles|theme|rels|Content_Types/.test(n)) {
          pieces.push(...unent(dec(b).replace(/<\/(row|text:p|table:table-row|si)>/g, '\n').replace(/<\/(c|table:table-cell)>/g, '\t').replace(/<[^>]*>?/g, '')).split('\n').map((l) => l.replace(/\t+$/, '')).filter((l) => l.trim()));
        }
        text = pieces.join('\n');
      }
      if (!text) text = stringsScan(u8);
      const lines = text.split('\n').map((l) => l.replace(/\s+$/, '')).filter((l) => l.trim().length > 0).slice(0, 200000);
      if (lines.length) {
        const book = linesBook(lines);
        addCand(mStr, 'strings', book, 0.3);
        mStr.status = 'partial';
        mStr.detail = `${plural(lines.length, 'line')} of text recovered; formatting, formulas and cell positions are lost.`;
      } else { mStr.status = 'failed'; mStr.detail = 'No readable text was found.'; }
    }

    // ---- unverified data decoded after damage (xlsx)
    const best = pickBest(cands);
    if (best && (best.id === 'xml' || best.id === 'rawsheet') && salv) {
      const sstB = salv.parts.get('xl/sharedStrings.xml');
      const sst = sstB ? parseSharedStrings(dec(sstB)).items : [];
      const names = workbookSheets(salv.parts);
      for (const [part, p] of preps) {
        if (!p.tail) continue;
        // Cells decoded after the damage that claim rows above the cut are LZ77 copies of earlier data.
        const lastMain = Math.max(-1, ...scanSheetCells(p.main.slice(-20000), sst).map((x) => x.r));
        const cells = scanSheetCells(p.tail, sst).filter((x) => x.r > lastMain);
        if (!cells.length) continue;
        const src = names.find((s) => s.part === part)?.name || part.replace(/^.*\//, '').replace(/\.xml$/, '');
        let nm = `${src.slice(0, 18)} (unverified)`;
        let k = 2;
        while (best.book.sheetByName(nm)) nm = `${src.slice(0, 15)} (unverified ${k++})`;
        const sh = new Sheet(best.book, nm);
        for (const { r, c, cell } of cells) sh.setComputed(r, c, cell);
        best.book.sheets.push(sh);
        report.unverified = true;
        report.notes.push(`${plural(cells.length, 'cell')} decoded after the damaged area of ${part} ${cells.length === 1 ? 'is' : 'are'} on the sheet "${nm}". They may be wrong or misplaced; check them against another copy.`);
      }
    }
    if (best?.corrupt) report.unverified = true;
    return finish(report, cands, base);
  } catch (e: any) {
    report.notes.push(`Recovery stopped unexpectedly: ${e?.message || e}.`);
    return finish(report, cands, base);
  }
}

function pickBest(cands: Candidate[]): Candidate | null {
  let best: Candidate | null = null, bestScore = -1;
  for (const c of cands) {
    const st = bookStats(c.book);
    const score = (st.cells + st.formulas * 0.05) * c.weight;
    if (score > bestScore) { bestScore = score; best = c; }
  }
  return best;
}

function finish(report: RecoveryReport, cands: Candidate[], base: string): RecoveryReport {
  const best = pickBest(cands);
  const st = bookStats(best?.book ?? null);
  if (best && (st.cells > 0 || best.id === 'open')) {
    report.book = best.book;
    report.bestMethod = best.id;
    report.textOnly = best.id === 'strings';
    report.stats = st;
    if (!best.book.props.title) best.book.props.title = base;
    best.book.activeSheet = Math.max(0, best.book.sheets.findIndex((s) => s.visibility === 'visible'));
    best.book.clearHistory();
    const via = report.methods.find((m) => m.id === best.id)?.name.toLowerCase() || best.id;
    if (!report.summary) {
      report.summary = best.id === 'open' && !report.methods.some((m) => m.status === 'partial' && m.id !== 'identify')
        ? `The file opened normally: ${plural(st.sheets, 'sheet')}, ${plural(st.cells, 'cell')}.`
        : best.id === 'strings'
          ? `Only loose text could be recovered (${plural(st.cells, 'line')}). Formatting, formulas and cell positions are lost.`
          : `Recovered ${plural(st.sheets, 'sheet')} and ${plural(st.cells, 'cell')}${st.formulas ? ` (${plural(st.formulas, 'formula')})` : ''} using ${via}.`;
    }
    if (report.unverified && !report.summary.includes('verify')) report.summary += ' Some data comes from a damaged area; verify it.';
  } else if (!report.summary) {
    report.summary = 'Nothing could be recovered from this file.';
  }
  return report;
}

/** Plain-text version of a report (for Copy report). */
export function reportText(r: RecoveryReport): string {
  const lines = [`Recovery report: ${r.fileName}`, r.summary, `Identified as: ${r.identification.description} (${r.identification.confidence} confidence)`];
  if (r.notes.length) lines.push('', ...r.notes.map((n) => `- ${n}`));
  lines.push('', 'Methods:');
  for (const m of r.methods) lines.push(`  [${m.status}] ${m.name}${r.bestMethod === m.id ? ' (used)' : ''}: ${m.detail}`);
  if (r.repaired) {
    lines.push('', `Repaired package: ${r.repaired.name}`);
    if (r.repaired.salvagedParts.length) lines.push(`  Salvaged from raw bytes: ${r.repaired.salvagedParts.join(', ')}`);
    if (r.repaired.repairedParts.length) lines.push(`  Repaired: ${r.repaired.repairedParts.join('; ')}`);
    if (r.repaired.lostParts.length) lines.push(`  Lost: ${r.repaired.lostParts.join(', ')}`);
  }
  if (r.foreign.length) lines.push('', 'Other data found in the file:', ...r.foreign.map((s) => `  ${s.description}: ${s.length} bytes at byte ${s.offset}`));
  if (r.book) {
    lines.push('', 'Recovered sheets:');
    for (const sh of r.book.sheets) { const u = sh.usedRange(false); lines.push(`  ${sh.name}: ${[...sh.cells.values()].filter((c) => (c.v != null && c.v !== '') || c.f).length} cells${u ? ` (${colName(u.c1)}${u.r1 + 1}:${colName(u.c2)}${u.r2 + 1})` : ''}`); }
  }
  return lines.join('\n');
}

export { readBiff, importXls } from './xls';
