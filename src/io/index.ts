/**
 * File format registry, sniffing and the import/export API, plus small helpers shared by the
 * converter modules (XML escaping and DOM access, text decoding, ZIP directory listing, base64).
 *
 * Converter modules are loaded lazily so the app only pays for the formats it actually opens.
 */
import type { Workbook } from '../model/workbook';
import type { Range } from '../model/address';
import { sniff, extOf, article, type SniffResult, type SniffCategory } from './sniff';

export { sniff, registerSniffer, type SniffResult, type SniffRule, type SniffCategory } from './sniff';

// --------------------------------------------------------------------------------------------
// Registry

export interface FormatDef {
  id: string;
  label: string;
  ext: string[];
  mime: string;
  canImport: boolean;
  canExport: boolean;
  binary: boolean;
  note?: string;
}

export interface CsvOptions {
  delimiter?: string;
  quote?: string;
  encoding?: string;
  headerRows?: number;
  decimal?: '.' | ',';
  thousands?: string;
  dateOrder?: 'mdy' | 'dmy' | 'ymd';
  detectTypes?: boolean;
  columnTypes?: ('general' | 'text' | 'date' | 'skip')[];
  bom?: boolean;
  lineEnding?: '\n' | '\r\n';
}

const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

export const FORMATS: FormatDef[] = [
  { id: 'xlsx', label: 'Excel Workbook (.xlsx)', ext: ['xlsx'], mime: XLSX_MIME, canImport: true, canExport: true, binary: true },
  { id: 'xlsm', label: 'Excel Macro-Enabled Workbook (.xlsm)', ext: ['xlsm'], mime: 'application/vnd.ms-excel.sheet.macroEnabled.12', canImport: true, canExport: true, binary: true, note: 'VBA projects are kept as-is but not run' },
  { id: 'xltx', label: 'Excel Template (.xltx)', ext: ['xltx', 'xltm'], mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.template', canImport: true, canExport: true, binary: true },
  { id: 'ods', label: 'OpenDocument Spreadsheet (.ods)', ext: ['ods', 'ots'], mime: 'application/vnd.oasis.opendocument.spreadsheet', canImport: true, canExport: true, binary: true },
  { id: 'fods', label: 'Flat OpenDocument Spreadsheet (.fods)', ext: ['fods'], mime: 'application/vnd.oasis.opendocument.spreadsheet-flat-xml', canImport: true, canExport: true, binary: false },
  { id: 'csv', label: 'CSV (Comma delimited) (.csv)', ext: ['csv'], mime: 'text/csv', canImport: true, canExport: true, binary: false },
  { id: 'tsv', label: 'Tab-separated values (.tsv)', ext: ['tsv', 'tab'], mime: 'text/tab-separated-values', canImport: true, canExport: true, binary: false },
  { id: 'txt', label: 'Text (Tab delimited) (.txt)', ext: ['txt', 'text'], mime: 'text/plain', canImport: true, canExport: true, binary: false },
  { id: 'html', label: 'Web Page (.html)', ext: ['html', 'htm', 'xhtml'], mime: 'text/html', canImport: true, canExport: true, binary: false },
  { id: 'json', label: 'Anne\'s Grid JSON (.json)', ext: ['json'], mime: 'application/json', canImport: true, canExport: true, binary: false },
  { id: 'qgw', label: 'Anne\'s Grid Workbook (.qgw)', ext: ['qgw', 'qlw'], mime: 'application/x-quillgrid', canImport: true, canExport: true, binary: true }, // rename:keep — Anne's Grid was named QuillLattice; its .qlw/.qlwe files still open
  { id: 'qgwe', label: 'Encrypted Anne\'s Grid Workbook (.qgwe)', ext: ['qgwe', 'qlwe'], mime: 'application/x-quillgrid-encrypted', canImport: true, canExport: true, binary: true, note: 'AES-256 with a password' }, // rename:keep — Anne's Grid was named QuillLattice; its .qlw/.qlwe files still open
  { id: 'xls', label: 'Excel 97-2003 Workbook (.xls)', ext: ['xls', 'xlt', 'xlw'], mime: 'application/vnd.ms-excel', canImport: true, canExport: false, binary: true, note: 'Opens values, formulas and basic formatting; save as .xlsx to keep it' },
  { id: 'sylk', label: 'SYLK (.slk)', ext: ['slk', 'sylk'], mime: 'application/x-sylk', canImport: true, canExport: false, binary: false },
  { id: 'dif', label: 'Data Interchange Format (.dif)', ext: ['dif'], mime: 'application/x-dif', canImport: true, canExport: false, binary: false },
  { id: 'xml', label: 'XML Spreadsheet 2003 (.xml)', ext: ['xml'], mime: 'application/xml', canImport: true, canExport: false, binary: false },
  // Modern, XML and text spreadsheet formats (import only); detection in sniff-b.ts.
  { id: 'xlsb', label: 'Excel Binary Workbook (.xlsb)', ext: ['xlsb'], mime: 'application/vnd.ms-excel.sheet.binary.macroEnabled.12', canImport: true, canExport: false, binary: true },
  { id: 'sxc', label: 'OpenOffice.org 1.0 Spreadsheet (.sxc)', ext: ['sxc', 'stc'], mime: 'application/vnd.sun.xml.calc', canImport: true, canExport: false, binary: true },
  { id: 'gnumeric', label: 'Gnumeric Spreadsheet (.gnumeric)', ext: ['gnumeric', 'gnm'], mime: 'application/x-gnumeric', canImport: true, canExport: false, binary: true },
  { id: 'ksp', label: 'KSpread / Calligra Sheets (.ksp)', ext: ['ksp'], mime: 'application/x-kspread', canImport: true, canExport: false, binary: true },
  { id: 'numbers', label: 'Apple Numbers (.numbers)', ext: ['numbers'], mime: 'application/vnd.apple.numbers', canImport: true, canExport: false, binary: true },
  { id: 'uos', label: 'Uniform Office Format Spreadsheet (.uos)', ext: ['uos', 'uof'], mime: 'application/vnd.uof.spreadsheet', canImport: true, canExport: false, binary: false },
  { id: 'ndjson', label: 'JSON Lines (.ndjson, .jsonl)', ext: ['ndjson', 'jsonl'], mime: 'application/x-ndjson', canImport: true, canExport: false, binary: false },
  { id: 'prn', label: 'Formatted Text (Space delimited) (.prn)', ext: ['prn'], mime: 'text/plain', canImport: true, canExport: false, binary: false },
  { id: 'sqlite', label: 'SQLite database (.sqlite, .db)', ext: ['sqlite', 'sqlite3', 'db3', 'db'], mime: 'application/vnd.sqlite3', canImport: true, canExport: false, binary: true, note: 'Tables open read-only as sheets' },
  { id: 'mht', label: 'Single File Web Page (.mht)', ext: ['mht', 'mhtml'], mime: 'multipart/related', canImport: true, canExport: false, binary: false },
  { id: 'gz', label: 'Gzip-compressed file (.gz)', ext: ['gz', 'gzip'], mime: 'application/gzip', canImport: true, canExport: false, binary: true },
  { id: 'md', label: 'Markdown table (.md)', ext: ['md', 'markdown'], mime: 'text/markdown', canImport: false, canExport: true, binary: false },
  // Legacy spreadsheets (import only)
  { id: 'lotus', label: 'Lotus 1-2-3 (.wk1, .wks, .wk3, .wk4, .123)', ext: ['wk1', 'wks', 'wk3', 'wk4', '123', 'wk2', 'wrk', 'wr1'], mime: 'application/vnd.lotus-1-2-3', canImport: true, canExport: false, binary: true },
  { id: 'quattro', label: 'Quattro Pro (.wq1, .wq2, .wb1, .wb2, .wb3, .qpw)', ext: ['wq1', 'wq2', 'wb1', 'wb2', 'wb3', 'qpw', 'wkq'], mime: 'application/x-quattro-pro', canImport: true, canExport: false, binary: true },
  { id: 'works', label: 'Microsoft Works spreadsheet (.wks, .xlr)', ext: ['xlr'], mime: 'application/vnd.ms-works', canImport: true, canExport: false, binary: true },
  { id: 'dbf', label: 'dBASE / FoxPro table (.dbf)', ext: ['dbf'], mime: 'application/x-dbf', canImport: true, canExport: false, binary: true },
  { id: 'visicalc', label: 'VisiCalc (.vc)', ext: ['vc'], mime: 'application/x-visicalc', canImport: true, canExport: false, binary: false },
  { id: 'cwk', label: 'AppleWorks / ClarisWorks spreadsheet (.cwk)', ext: ['cwk', 'cws'], mime: 'application/x-appleworks', canImport: true, canExport: false, binary: true, note: 'Values and text only' },
  { id: 'supercalc', label: 'SuperCalc (.cal)', ext: ['cal'], mime: 'application/x-supercalc', canImport: true, canExport: false, binary: true, note: 'Text and numbers recovered as values' },
  { id: 'multiplan', label: 'Multiplan (.mp)', ext: ['mp'], mime: 'application/x-multiplan', canImport: true, canExport: false, binary: true, note: 'Text and numbers recovered as values' },
  // Other kinds of documents: their tables (and text) become sheets
  { id: 'docx', label: 'Word document tables (.docx)', ext: ['docx', 'docm', 'dotx', 'dotm'], mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', canImport: true, canExport: false, binary: true },
  { id: 'odt', label: 'OpenDocument text tables (.odt)', ext: ['odt', 'ott', 'fodt', 'sxw'], mime: 'application/vnd.oasis.opendocument.text', canImport: true, canExport: false, binary: true },
  { id: 'rtf', label: 'Rich Text Format tables (.rtf)', ext: ['rtf'], mime: 'application/rtf', canImport: true, canExport: false, binary: false },
  { id: 'doc', label: 'Word 97-2003 document tables (.doc)', ext: ['doc', 'dot'], mime: 'application/msword', canImport: true, canExport: false, binary: true },
  { id: 'wordtext', label: 'Other word processor documents (text)', ext: ['wpd', 'wps', 'wri', 'sdw', 'hwp', 'lwp', 'abw', 'pages', 'epub'], mime: 'application/octet-stream', canImport: true, canExport: false, binary: true, note: 'Text only, one paragraph per row' },
  { id: 'pptx', label: 'PowerPoint slides (.pptx)', ext: ['pptx', 'pptm', 'potx', 'ppsx'], mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', canImport: true, canExport: false, binary: true },
  { id: 'odp', label: 'OpenDocument slides (.odp)', ext: ['odp', 'otp', 'fodp', 'sxi'], mime: 'application/vnd.oasis.opendocument.presentation', canImport: true, canExport: false, binary: true },
  { id: 'ppt', label: 'PowerPoint 97-2003 slides (.ppt)', ext: ['ppt', 'pps', 'pot'], mime: 'application/vnd.ms-powerpoint', canImport: true, canExport: false, binary: true, note: 'Slide text' },
  { id: 'eml', label: 'E-mail message (.eml)', ext: ['eml', 'emlx'], mime: 'message/rfc822', canImport: true, canExport: false, binary: false },
  { id: 'mbox', label: 'Mailbox (.mbox)', ext: ['mbox', 'mbx'], mime: 'application/mbox', canImport: true, canExport: false, binary: false },
  { id: 'msg', label: 'Outlook message (.msg)', ext: ['msg'], mime: 'application/vnd.ms-outlook', canImport: true, canExport: false, binary: true },
];

/** Formats whose files are not spreadsheets: opening one imports its tables, text or messages. */
export const DOCUMENT_FORMATS = ['docx', 'odt', 'rtf', 'doc', 'wordtext', 'pptx', 'odp', 'ppt', 'eml', 'mbox', 'msg'];
/** Legacy spreadsheet formats read by the modules in ./legacy. */
export const LEGACY_FORMATS = ['lotus', 'quattro', 'works', 'dbf', 'visicalc', 'cwk', 'supercalc', 'multiplan'];

export function formatForName(name: string): FormatDef | undefined {
  const lower = name.toLowerCase();
  const ext = lower.includes('.') ? lower.slice(lower.lastIndexOf('.') + 1) : lower;
  return FORMATS.find((f) => f.ext.includes(ext));
}

export function formatById(id: string): FormatDef | undefined {
  return FORMATS.find((f) => f.id === id);
}

// --------------------------------------------------------------------------------------------
// Sniffing

export function sniffFormat(data: ArrayBuffer, name: string): FormatDef | undefined {
  const r = sniff(data, name);
  return formatForSniff(r, name);
}

/** The importer for a sniff result; text that nothing recognized opens as delimited text. */
export function formatForSniff(r: SniffResult, name: string): FormatDef | undefined {
  const byId = formatById(r.format);
  if (byId) {
    // ZIP or compound files whose content is ambiguous: a matching extension decides between siblings.
    const byName = formatForName(name);
    if (byName && byName.id !== byId.id && ['xlsx', 'xlsm', 'xltx'].includes(byId.id) && ['xlsx', 'xlsm', 'xltx'].includes(byName.id)) return byName;
    return byId;
  }
  if (r.format === 'unknown') {
    const byName = formatForName(name);
    if (byName?.canImport) return byName;
    return undefined;
  }
  return undefined;
}

// --------------------------------------------------------------------------------------------
// Import / export

export interface ImportResult {
  book: Workbook; format: string; warnings: string[]; recovery?: any; encrypted?: boolean;
  /** What the content was identified as. */
  sniff?: SniffResult;
  /** Short message to show when the file was opened by its content (extension wrong or missing). */
  note?: string;
}

/** Error thrown when a password is needed (or was wrong). `needsPassword` lets the UI prompt. */
export class PasswordError extends Error {
  needsPassword = true;
  constructor(msg: string, public wrong = false) { super(msg); }
}

const ZIP_FORMATS = ['xlsx', 'xlsm', 'xltx', 'ods', 'qgw'];

export interface ImportOptions {
  repair?: boolean;
  password?: string;
  csv?: CsvOptions;
  /** Import as this format instead of the detected one. */
  as?: string;
}

export async function importFile(data: ArrayBuffer, name: string, opts: ImportOptions = {}): Promise<ImportResult> {
  const sn = sniff(data, name);
  const fmt = opts.as ? formatById(opts.as) : formatForSniff(sn, name);
  const warnings: string[] = [];
  const recover = async (cause: unknown): Promise<ImportResult> => {
    let report: any = null;
    try {
      const mod: any = await import('./recover');
      report = await mod.recoverWorkbook(data, name, { forceRepair: !!opts.repair });
    } catch (e) {
      if (cause) throw new Error(friendlyOpenError(errMsg(cause)));
      throw e;
    }
    if (!report || !report.book) {
      if (cause) throw new Error(friendlyOpenError(errMsg(cause)));
      throw new Error('No recoverable content was found in this file.');
    }
    const w: string[] = [];
    if (cause) w.push(`Normal open failed: ${friendlyOpenError(errMsg(cause))}`);
    if (Array.isArray(report.warnings)) w.push(...report.warnings);
    if (Array.isArray(report.notes)) w.push(...report.notes);
    report.book.clearHistory?.();
    return { book: report.book, format: report.format || fmt?.id || 'xlsx', warnings: w, recovery: report, sniff: sn };
  };

  if (opts.repair) return recover(null);
  if (!fmt || !fmt.canImport) {
    if (sn.format !== 'unknown' && sn.confidence >= 0.5) throw new Error(`This is ${article(sn.label)} ${sn.label}. Anne\'s Grid recognizes it but cannot open this kind of file.`);
    throw new Error(`Anne\'s Grid does not recognize the contents of ${name.replace(/^.*[\\/]/, '') || 'this file'}. Try Open and Repair to recover any text in it.`);
  }
  const note = opts.as ? undefined : sn.note;
  try {
    const book = await importAs(fmt.id, data, name, opts, warnings);
    book.clearHistory();
    return { book, format: fmt.id, warnings, encrypted: fmt.id === 'qgwe' || undefined, sniff: sn, note };
  } catch (e) {
    if (e instanceof PasswordError) throw e;
    if (fmt.id === 'xls' || ZIP_FORMATS.includes(fmt.id)) {
      if (fmt.id === 'xls' && isEncryptedOoxml(new Uint8Array(data))) throw new PasswordError('This workbook is encrypted with a password by Excel. Remove the password in Excel (File ▸ Info ▸ Protect Workbook) and open it again.');
      return recover(e);
    }
    if (LEGACY_FORMATS.includes(fmt.id)) {
      // Damaged legacy file: read it again leniently, then fall back to the recovery pipeline.
      try {
        const w: string[] = [];
        const book = await importLegacy(fmt.id, data, name, w, true);
        if (book && book.sheets.some((s) => s.cells.size)) {
          book.clearHistory();
          return { book, format: fmt.id, warnings: [`The file is damaged: ${friendlyOpenError(errMsg(e))}`, ...w], sniff: sn, note };
        }
      } catch { /* fall through */ }
      return recover(e);
    }
    throw new Error(friendlyOpenError(errMsg(e)));
  }
}

async function importLegacy(id: string, data: ArrayBuffer, name: string, warnings: string[], lenient: boolean): Promise<Workbook> {
  switch (id) {
    case 'lotus': return (await import('./legacy/lotus')).importLotus(data, name, warnings, { lenient });
    case 'works': {
      const u8 = new Uint8Array(data);
      if (u8[0] === 0xd0 && u8[1] === 0xcf) return importBiff(data, name, lenient, warnings);
      return (await import('./legacy/lotus')).importLotus(data, name, warnings, { lenient });
    }
    case 'quattro': return (await import('./legacy/quattro')).importQuattro(data, name, warnings, { lenient });
    case 'dbf': return (await import('./legacy/dbf')).importDbf(data, name, warnings, { lenient });
    case 'visicalc': return (await import('./legacy/visicalc')).importVisiCalc(decodeText(data), name, warnings);
    case 'cwk': return (await import('./legacy/textscan')).importAppleWorks(data, name, warnings);
    case 'supercalc': case 'multiplan': return (await import('./legacy/textscan')).importValuesScan(data, name, warnings, id === 'supercalc' ? 'SuperCalc' : 'Multiplan');
  }
  throw new Error(`Cannot open ${formatById(id)?.label || id} files.`);
}

async function importBiff(data: ArrayBuffer, name: string, lenient: boolean, warnings: string[]): Promise<Workbook> {
  const mod: any = await import('./recover/xls');
  if (lenient) {
    const r = mod.readBiff(new Uint8Array(data), { lenient: true });
    warnings.push(...r.warnings);
    return r.book;
  }
  const book = await mod.importXls(data, name, warnings);
  if (!book) throw new Error('The .xls file could not be read.');
  return book;
}

async function importAs(id: string, data: ArrayBuffer, name: string, opts: { password?: string; csv?: CsvOptions }, warnings: string[]): Promise<Workbook> {
  if (LEGACY_FORMATS.includes(id)) return importLegacy(id, data, name, warnings, false);
  if (DOCUMENT_FORMATS.includes(id)) return (await import('./documents')).importDocument(id, data, name, warnings);
  switch (id) {
    case 'xlsx': case 'xlsm': case 'xltx': return (await import('./xlsx-import')).importXlsx(data, name, warnings);
    case 'ods': return (await import('./ods')).importOds(data, warnings);
    case 'fods': return (await import('./ods')).importFods(decodeText(data), warnings);
    case 'csv': case 'tsv': case 'txt': {
      const csv = await import('./csv');
      const o: CsvOptions = { ...(opts.csv || {}) };
      if (!o.delimiter && id === 'tsv') o.delimiter = '\t';
      return csv.importCsv(data, name, o);
    }
    case 'html': return (await import('./html')).importHtml(decodeText(data), warnings);
    case 'json': {
      const text = decodeText(data);
      if (/"format"\s*:\s*"quillgrid"|"qgw"\s*:/.test(text.slice(0, 4096))) return (await import('./native')).importJson(text);
      return (await import('./json-records')).importJsonRecords(text, name, warnings);
    }
    case 'qgw': return (await import('./native')).importQlw(data);
    case 'qgwe': {
      if (!opts.password) throw new PasswordError('This workbook is protected with a password.');
      return (await import('./native')).importQlwe(data, opts.password);
    }
    case 'xls': return importBiff(data, name, false, warnings);
    case 'sylk': return (await import('./other')).importSylk(decodeText(data));
    case 'dif': return (await import('./other')).importDif(decodeText(data));
    case 'xml': return (await import('./other')).importXml2003(decodeText(data), warnings);
    case 'xlsb': return (await import('./xlsb')).importXlsb(data, warnings);
    case 'sxc': return (await import('./sxc')).importSxc(data, warnings);
    case 'gnumeric': return (await import('./gnumeric')).importGnumeric(data, warnings);
    case 'ksp': return (await import('./kspread')).importKspread(data, warnings);
    case 'numbers': return (await import('./numbers')).importNumbers(data, warnings);
    case 'uos': return (await import('./uof')).importUof(decodeText(data), warnings);
    case 'ndjson': return (await import('./json-records')).importNdjson(decodeText(data), name, warnings);
    case 'prn': return (await import('./fixedwidth')).importFixedWidth(data, name, opts.csv || {}, warnings);
    case 'sqlite': return (await import('./sqlite')).importSqlite(data, warnings);
    case 'mht': return (await import('./mhtml')).importMhtml(data, warnings);
    case 'gz': {
      const inner = await (await import('./gzip')).gunzip(data);
      const innerName = name.replace(/\.(gz|gzip)$/i, '') || 'file';
      const fmt = sniffFormat(inner.slice().buffer, innerName);
      if (!fmt || fmt.id === 'gz') throw new Error('The compressed file does not contain a spreadsheet.');
      return importAs(fmt.id, inner.slice().buffer, innerName, opts, warnings);
    }
  }
  throw new Error(`Cannot open ${formatById(id)?.label || id} files.`);
}

/** Category of a file's content (spreadsheet, document, presentation, email). */
export function sniffCategory(data: ArrayBuffer, name: string): SniffCategory { return sniff(data, name).category; }

/** The extension of a file name, lower case, without the dot. */
export { extOf };

export async function exportFile(book: Workbook, formatId: string, opts: { sheetIndex?: number; password?: string; csv?: CsvOptions; selectionOnly?: Range; /** Receives notes about anything the format could not hold. */ warnings?: string[] } = {}): Promise<Blob> {
  const fmt = formatById(formatId);
  if (!fmt || !fmt.canExport) throw new Error(`Cannot save as ${formatId}`);
  const sheetIndex = opts.sheetIndex ?? book.activeSheet;
  switch (fmt.id) {
    case 'xlsx': case 'xlsm': case 'xltx': return (await import('./xlsx-export')).exportXlsx(book, { kind: fmt.id as 'xlsx', warnings: opts.warnings });
    case 'ods': return (await import('./ods')).exportOds(book, opts.warnings);
    case 'fods': return (await import('./ods')).exportFods(book);
    case 'csv': case 'tsv': case 'txt': {
      const o: CsvOptions = { ...(opts.csv || {}) };
      if (!o.delimiter) o.delimiter = fmt.id === 'csv' ? ',' : '\t';
      return (await import('./csv')).exportCsv(book, sheetIndex, o, opts.selectionOnly);
    }
    case 'html': return (await import('./html')).exportHtml(book, { sheetIndex: opts.sheetIndex, range: opts.selectionOnly });
    case 'md': return (await import('./html')).exportMarkdown(book, sheetIndex, opts.selectionOnly);
    case 'json': return (await import('./native')).exportJson(book);
    case 'qgw': return (await import('./native')).exportQlw(book);
    case 'qgwe': {
      if (!opts.password) throw new PasswordError('A password is required to save an encrypted workbook.');
      return (await import('./native')).exportQlwe(book, opts.password);
    }
  }
  throw new Error(`Cannot save as ${formatId}`);
}

export { parseCsv, detectDelimiter } from './csv';

/** Translates low-level ZIP/XML/inflate errors into plain language. */
export function friendlyOpenError(msg: string): string {
  if (/size mismatch|invalid (distance|code|block|stored)|too far back|incorrect header|unexpected end|crc32|corrupt/i.test(msg)) return 'compressed data inside the file is damaged.';
  if (/end of central directory|is this a zip|zip file/i.test(msg)) return 'the file\'s table of contents (ZIP central directory) is missing or damaged.';
  if (/workbook\.xml|workbook part/i.test(msg)) return 'the main workbook part could not be read.';
  if (/xml parse error|parsererror|not well-formed/i.test(msg)) return 'part of the file contains malformed XML.';
  if (/password|decrypt|operation-specific reason|OperationError/i.test(msg)) return 'the password is incorrect or the file is damaged.';
  return msg;
}

function errMsg(e: unknown): string { return e instanceof Error ? e.message : String(e); }

function isEncryptedOoxml(u8: Uint8Array): boolean {
  const pat = 'EncryptedPackage';
  const lim = Math.min(u8.length, 1 << 20);
  outer: for (let i = 0; i + pat.length * 2 < lim; i++) {
    for (let j = 0; j < pat.length; j++) if (u8[i + j * 2] !== pat.charCodeAt(j) || u8[i + j * 2 + 1] !== 0) continue outer;
    return true;
  }
  return false;
}

// --------------------------------------------------------------------------------------------
// Shared helpers: text

function latin1(u8: Uint8Array): string {
  let s = '';
  for (let i = 0; i < u8.length; i++) s += String.fromCharCode(u8[i]);
  return s;
}

/** Decodes bytes as text: BOM detection (UTF-8/16), then UTF-8, falling back to Windows-1252. */
export function decodeText(data: ArrayBuffer | Uint8Array, encoding?: string): string {
  const u8 = data instanceof Uint8Array ? data : new Uint8Array(data);
  if (u8[0] === 0xff && u8[1] === 0xfe) return new TextDecoder('utf-16le').decode(u8.subarray(2));
  if (u8[0] === 0xfe && u8[1] === 0xff) return new TextDecoder('utf-16be').decode(u8.subarray(2));
  if (u8[0] === 0xef && u8[1] === 0xbb && u8[2] === 0xbf) return new TextDecoder('utf-8').decode(u8.subarray(3));
  if (encoding && !/^utf-?8$/i.test(encoding)) {
    try { return new TextDecoder(encoding).decode(u8); } catch { /* unsupported label: fall through */ }
  }
  // BOM-less UTF-16 (common for "Unicode text" exports): many zero bytes at odd/even positions.
  if (!encoding && u8.length >= 4) {
    let zo = 0, ze = 0;
    const n = Math.min(u8.length, 2000);
    for (let i = 0; i < n; i++) if (u8[i] === 0) { if (i & 1) zo++; else ze++; }
    if (zo > n / 4 && ze < n / 50) return new TextDecoder('utf-16le').decode(u8);
    if (ze > n / 4 && zo < n / 50) return new TextDecoder('utf-16be').decode(u8);
  }
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(u8);
  } catch {
    return new TextDecoder('windows-1252').decode(u8);
  }
}

// --------------------------------------------------------------------------------------------
// Shared helpers: XML

// eslint-disable-next-line no-control-regex
const INVALID_XML = /[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/g;

export function stripInvalidXml(s: string): string { return s.replace(INVALID_XML, ''); }

/** Escapes text for XML content or attribute values and strips characters invalid in XML 1.0. */
export function esc(s: unknown): string {
  return stripInvalidXml(String(s ?? ''))
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}

export const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

const ENT: Record<string, string> = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };
/** Decodes XML entities and character references. */
export function unescXml(s: string): string {
  if (s.indexOf('&') < 0) return s;
  return s.replace(/&(#x[0-9a-fA-F]+|#\d+|amp|lt|gt|quot|apos);/g, (_m, e: string) => {
    if (e[0] === '#') {
      const cp = e[1] === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
      try { return String.fromCodePoint(cp); } catch { return ''; }
    }
    return ENT[e];
  });
}

/** Parses an XML string; throws on malformed XML. */
export function parseXml(text: string, what = 'XML'): Document {
  const doc = new DOMParser().parseFromString(text.replace(/^﻿/, ''), 'application/xml');
  const err = doc.getElementsByTagName('parsererror')[0];
  if (err && (doc.documentElement.localName === 'parsererror' || err.parentNode === doc.documentElement)) {
    throw new Error(`XML parse error in ${what}: ${(err.textContent || '').slice(0, 200)}`);
  }
  return doc;
}

/** Direct element children with the given local name (namespace/prefix agnostic). */
export function kids(el: Element | Document | null | undefined, local?: string): Element[] {
  if (!el) return [];
  const out: Element[] = [];
  const parent = (el as Document).documentElement && (el as Document).nodeType === 9 ? [(el as Document).documentElement] : null;
  const list = parent || Array.from((el as Element).children || []);
  for (const c of list) if (!local || c.localName === local) out.push(c);
  return out;
}

export function kid(el: Element | Document | null | undefined, local: string): Element | null {
  if (!el) return null;
  if ((el as Document).nodeType === 9) { const d = (el as Document).documentElement; return d && d.localName === local ? d : null; }
  const ch = (el as Element).children;
  for (let i = 0; i < ch.length; i++) if (ch[i].localName === local) return ch[i];
  return null;
}

/** Child at a path of local names: path(el, 'a', 'b') = first b inside first a. */
export function path(el: Element | null | undefined, ...names: string[]): Element | null {
  let cur: Element | null = el || null;
  for (const n of names) { if (!cur) return null; cur = kid(cur, n); }
  return cur;
}

/** All descendants with the given local name. */
export function desc(el: Element | Document | null | undefined, local: string): Element[] {
  if (!el) return [];
  return Array.from(el.getElementsByTagNameNS('*', local));
}

/** Attribute by name; a bare local name also matches prefixed attributes (r:id matches 'id'). */
export function attr(el: Element | null | undefined, name: string): string | null {
  if (!el) return null;
  const v = el.getAttribute(name);
  if (v != null) return v;
  const at = el.attributes;
  for (let i = 0; i < at.length; i++) if (at[i].localName === name) return at[i].value;
  return null;
}

/** r:id style relationship attribute. */
export function relId(el: Element | null | undefined, name = 'id'): string | null {
  if (!el) return null;
  const at = el.attributes;
  for (let i = 0; i < at.length; i++) {
    const a = at[i];
    if (a.localName === name && (a.prefix || a.namespaceURI)) return a.value;
  }
  return el.getAttribute('r:' + name);
}

export function numAttr(el: Element | null | undefined, name: string): number | null {
  const v = attr(el, name);
  if (v == null || v === '') return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

export function boolAttr(el: Element | null | undefined, name: string, dflt = false): boolean {
  const v = attr(el, name);
  if (v == null) return dflt;
  return v === '1' || v.toLowerCase() === 'true' || v.toLowerCase() === 'on';
}

// --------------------------------------------------------------------------------------------
// Shared helpers: binary

/** Lists entry names in a ZIP archive from the central directory (local headers as a fallback). */
export function zipNames(u8: Uint8Array): string[] {
  const names: string[] = [];
  const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
  const dec = new TextDecoder('utf-8');
  let eocd = -1;
  for (let i = u8.length - 22; i >= Math.max(0, u8.length - 65557); i--) {
    if (u8[i] === 0x50 && u8[i + 1] === 0x4b && u8[i + 2] === 5 && u8[i + 3] === 6) { eocd = i; break; }
  }
  if (eocd >= 0) {
    const count = dv.getUint16(eocd + 10, true);
    let p = dv.getUint32(eocd + 16, true);
    for (let k = 0; k < count && p + 46 <= u8.length; k++) {
      if (dv.getUint32(p, true) !== 0x02014b50) break;
      const nl = dv.getUint16(p + 28, true), el = dv.getUint16(p + 30, true), cl = dv.getUint16(p + 32, true);
      names.push(dec.decode(u8.subarray(p + 46, p + 46 + nl)));
      p += 46 + nl + el + cl;
    }
    if (names.length) return names;
  }
  // Scan local file headers.
  for (let p = 0; p + 30 <= u8.length;) {
    if (dv.getUint32(p, true) !== 0x04034b50) { p++; continue; }
    const nl = dv.getUint16(p + 26, true), el = dv.getUint16(p + 28, true), cs = dv.getUint32(p + 18, true);
    names.push(dec.decode(u8.subarray(p + 30, p + 30 + nl)));
    p += 30 + nl + el + (cs > 0 && cs < u8.length ? cs : 0);
    if (names.length > 5000) break;
  }
  return names;
}

export function bytesToBase64(bytes: Uint8Array): string {
  let s = '';
  const CH = 0x8000;
  for (let i = 0; i < bytes.length; i += CH) s += String.fromCharCode.apply(null, Array.from(bytes.subarray(i, i + CH)));
  return btoa(s);
}

export function base64ToBytes(b64: string): Uint8Array {
  const s = atob(b64.replace(/\s+/g, ''));
  const out = new Uint8Array(s.length);
  for (let i = 0; i < s.length; i++) out[i] = s.charCodeAt(i);
  return out;
}

const MIME_EXT: Record<string, string> = {
  png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp',
  tif: 'image/tiff', tiff: 'image/tiff', emf: 'image/x-emf', wmf: 'image/x-wmf', ico: 'image/x-icon',
};
export function mimeForExt(ext: string): string { return MIME_EXT[ext.toLowerCase()] || 'application/octet-stream'; }
export function extForMime(mime: string): string {
  const m = mime.toLowerCase();
  if (m === 'image/jpeg') return 'jpeg';
  for (const [e, t] of Object.entries(MIME_EXT)) if (t === m) return e;
  return 'png';
}

export function parseDataUrl(src: string): { mime: string; bytes: Uint8Array } | null {
  const m = /^data:([^;,]+)?((?:;[^;,]*)*?)(;base64)?,(.*)$/s.exec(src);
  if (!m) return null;
  const mime = m[1] || 'application/octet-stream';
  if (m[3]) return { mime, bytes: base64ToBytes(m[4]) };
  return { mime, bytes: new TextEncoder().encode(decodeURIComponent(m[4])) };
}

/** Size in px of a PNG/JPEG/GIF image, when it can be read from the header. */
export function imageSize(b: Uint8Array): { w: number; h: number } | null {
  if (b[0] === 0x89 && b[1] === 0x50 && b.length > 24) return { w: (b[16] << 24 | b[17] << 16 | b[18] << 8 | b[19]) >>> 0, h: (b[20] << 24 | b[21] << 16 | b[22] << 8 | b[23]) >>> 0 };
  if (b[0] === 0x47 && b[1] === 0x49 && b.length > 10) return { w: b[6] | (b[7] << 8), h: b[8] | (b[9] << 8) };
  if (b[0] === 0xff && b[1] === 0xd8) {
    let p = 2;
    while (p + 9 < b.length) {
      if (b[p] !== 0xff) { p++; continue; }
      const mk = b[p + 1];
      const len = (b[p + 2] << 8) | b[p + 3];
      if (mk >= 0xc0 && mk <= 0xcf && mk !== 0xc4 && mk !== 0xc8 && mk !== 0xcc) return { h: (b[p + 5] << 8) | b[p + 6], w: (b[p + 7] << 8) | b[p + 8] };
      p += 2 + len;
    }
  }
  return null;
}

// --------------------------------------------------------------------------------------------
// Shared helpers: formulas

/** Removes Excel's future-function prefixes (_xlfn., _xlws., _xlpm.) outside string literals. */
export function stripXlfn(f: string): string {
  if (f.indexOf('_xl') < 0) return f;
  return mapOutsideStrings(f, (s) => s.replace(/_xlfn\._xlws\.|_xlfn\.|_xlws\.|_xlpm\./gi, ''));
}

/** Applies fn to the parts of a formula that are not inside "string literals" or 'quoted names'. */
export function mapOutsideStrings(f: string, fn: (s: string) => string): string {
  let out = '', i = 0, start = 0;
  while (i < f.length) {
    const ch = f[i];
    if (ch === '"' || ch === "'") {
      out += fn(f.slice(start, i));
      let j = i + 1;
      while (j < f.length) { if (f[j] === ch) { if (f[j + 1] === ch) { j += 2; continue; } break; } j++; }
      out += f.slice(i, j + 1);
      i = start = j + 1;
      continue;
    }
    i++;
  }
  return out + fn(f.slice(start));
}

// --------------------------------------------------------------------------------------------
// Shared helpers: colors

/** Normalizes a CSS-ish color to '#rrggbb' (or null). */
export function normColor(c: string | null | undefined): string | null {
  if (!c) return null;
  const s = c.trim().toLowerCase();
  let m = /^#?([0-9a-f]{6})$/.exec(s);
  if (m) return '#' + m[1];
  m = /^#([0-9a-f])([0-9a-f])([0-9a-f])$/.exec(s);
  if (m) return '#' + m[1] + m[1] + m[2] + m[2] + m[3] + m[3];
  m = /^#?([0-9a-f]{2})([0-9a-f]{6})$/.exec(s); // AARRGGBB
  if (m) return '#' + m[2];
  m = /^rgba?\(\s*(\d+)[\s,]+(\d+)[\s,]+(\d+)/.exec(s);
  if (m) return '#' + [m[1], m[2], m[3]].map((x) => Math.min(255, +x).toString(16).padStart(2, '0')).join('');
  const named = CSS_NAMED[s];
  return named ? '#' + named : null;
}

const CSS_NAMED: Record<string, string> = {
  black: '000000', white: 'ffffff', red: 'ff0000', green: '008000', lime: '00ff00', blue: '0000ff', yellow: 'ffff00', cyan: '00ffff', aqua: '00ffff',
  magenta: 'ff00ff', fuchsia: 'ff00ff', gray: '808080', grey: '808080', silver: 'c0c0c0', maroon: '800000', olive: '808000', purple: '800080',
  teal: '008080', navy: '000080', orange: 'ffa500', pink: 'ffc0cb', brown: 'a52a2a', gold: 'ffd700', lightgray: 'd3d3d3', lightgrey: 'd3d3d3',
  darkgray: 'a9a9a9', darkgrey: 'a9a9a9', darkred: '8b0000', darkblue: '00008b', darkgreen: '006400', lightblue: 'add8e6', lightgreen: '90ee90',
  lightyellow: 'ffffe0', whitesmoke: 'f5f5f5', gainsboro: 'dcdcdc',
};
