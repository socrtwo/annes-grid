/**
 * SpreadsheetML (.xlsx / .xlsm / .xltx) writer.
 *
 * Produces a package that opens cleanly in Excel and LibreOffice: element order inside each part
 * follows the schema sequence, every string is escaped and stripped of characters XML 1.0 forbids,
 * and styles are deduplicated (fonts, fills, borders, number formats) from the style ids in use.
 */
import JSZip from 'jszip';
import { Workbook, type Sheet } from '../model/workbook';
import { renameSheet } from '../model/structure';
import {
  CellError, type Cell, type CellStyle, type Border, type CFRule, type DataValidation, type ChartDef, type ChartSeries, type ChartType,
  type DrawingObject, type TableDef, type Scalar,
} from '../model/types';
import { cellKey, keyRow, keyCol, colName, cellName, rangeName, parseRange, absRangeName, quoteSheet, splitSheetRef, MAX_COLS, MAX_ROWS, type Range } from '../model/address';
import { BUILTIN_FORMATS } from '../model/numfmt';
import { tokenize, type Token } from '../calc/parser';
import { esc, XML_HEADER, parseDataUrl, extForMime, imageSize, normColor } from './index';
import { PROTECT_ACTIONS, PAPER } from './xlsx-import';

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const NS_XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const CT = 'application/vnd.openxmlformats-officedocument.';

const EMU = 9525;

// --------------------------------------------------------------------------------------------
// Formula translation

const XLFN = new Set(`XLOOKUP XMATCH SORTBY UNIQUE SEQUENCE RANDARRAY LET LAMBDA MAP REDUCE SCAN MAKEARRAY BYROW BYCOL ISOMITTED
TEXTSPLIT TEXTBEFORE TEXTAFTER VSTACK HSTACK TOCOL TOROW WRAPROWS WRAPCOLS TAKE DROP EXPAND CHOOSEROWS CHOOSECOLS IFS SWITCH MAXIFS
MINIFS CONCAT TEXTJOIN XOR IFNA AGGREGATE FORECAST.LINEAR FORECAST.ETS FORECAST.ETS.CONFINT FORECAST.ETS.SEASONALITY FORECAST.ETS.STAT
STDEV.S STDEV.P VAR.S VAR.P NORM.DIST NORM.INV NORM.S.DIST NORM.S.INV T.DIST T.DIST.2T T.DIST.RT T.INV T.INV.2T T.TEST CHISQ.DIST
CHISQ.DIST.RT CHISQ.INV CHISQ.INV.RT CHISQ.TEST F.DIST F.DIST.RT F.INV F.INV.RT F.TEST BINOM.DIST BINOM.INV BINOM.DIST.RANGE EXPON.DIST
GAMMA.DIST GAMMA.INV GAMMALN.PRECISE GAMMA LOGNORM.DIST LOGNORM.INV NEGBINOM.DIST POISSON.DIST WEIBULL.DIST BETA.DIST BETA.INV
HYPGEOM.DIST CONFIDENCE.NORM CONFIDENCE.T COVARIANCE.P COVARIANCE.S MODE.SNGL MODE.MULT PERCENTILE.INC PERCENTILE.EXC QUARTILE.INC
QUARTILE.EXC RANK.EQ RANK.AVG PERCENTRANK.INC PERCENTRANK.EXC Z.TEST CEILING.MATH CEILING.PRECISE FLOOR.MATH FLOOR.PRECISE ISO.CEILING
ERF.PRECISE ERFC.PRECISE NETWORKDAYS.INTL WORKDAY.INTL DAYS ISOWEEKNUM ISFORMULA FORMULATEXT SHEET SHEETS ACOT ACOTH COT COTH CSC CSCH
SEC SECH ARABIC BASE DECIMAL COMBINA PERMUTATIONA BITAND BITOR BITXOR BITLSHIFT BITRSHIFT ENCODEURL FILTERXML WEBSERVICE NUMBERVALUE
UNICHAR UNICODE PDURATION RRI GAUSS PHI SKEW.P STOCKHISTORY ARRAYTOTEXT VALUETOTEXT IMAGE GROUPBY PIVOTBY PERCENTOF REGEXTEST
REGEXEXTRACT REGEXREPLACE TRIMRANGE ANCHORARRAY SINGLE FIELDVALUE IMCOSH IMCOT IMCSC IMCSCH IMSEC IMSECH IMSINH IMTAN`.split(/\s+/));
const XLWS = new Set(['FILTER', 'SORT']);

/** Converts our formula text to the stored OOXML form (adds _xlfn./_xlws./_xlpm. prefixes). */
export function toFileFormula(f: string): string {
  let toks: Token[];
  try { toks = tokenize(f); } catch { return f; }
  // LET / LAMBDA parameter names get _xlpm.
  const params = new Set<string>();
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (t.type !== 'func') continue;
    const up = t.text.toUpperCase();
    if (up !== 'LET' && up !== 'LAMBDA') continue;
    let depth = 0, arg = 0;
    const argToks: Token[][] = [[]];
    for (let j = i + 1; j < toks.length; j++) {
      const x = toks[j];
      if (x.type === 'lparen' || x.type === 'lbrace') { depth++; if (depth === 1) continue; }
      if (x.type === 'rparen' || x.type === 'rbrace') { depth--; if (depth === 0) break; }
      if (x.type === 'comma' && depth === 1) { arg++; argToks.push([]); continue; }
      if (x.type !== 'ws') argToks[arg].push(x);
    }
    argToks.forEach((a, k) => {
      const isParam = up === 'LET' ? k % 2 === 0 && k < argToks.length - 1 : k < argToks.length - 1;
      if (isParam && a.length === 1 && a[0].type === 'name') params.add(a[0].text.toUpperCase());
    });
  }
  let out = '', last = 0;
  for (const t of toks) {
    let rep: string | null = null;
    if (t.type === 'func') {
      const up = t.text.toUpperCase();
      if (/^_xl/i.test(up)) continue;
      if (XLWS.has(up)) rep = '_xlfn._xlws.' + t.text;
      else if (XLFN.has(up)) rep = '_xlfn.' + t.text;
      else if (params.has(up)) rep = '_xlpm.' + t.text;
    } else if (t.type === 'name' && params.has(t.text.toUpperCase())) rep = '_xlpm.' + t.text;
    if (rep == null) continue;
    out += f.slice(last, t.pos) + rep;
    last = t.end;
  }
  return out + f.slice(last);
}

// --------------------------------------------------------------------------------------------
// Text helpers

/** Encodes characters XML cannot carry (and literal _xHHHH_ sequences) the OOXML way. */
function encX(s: string): string {
  // eslint-disable-next-line no-control-regex
  return s.replace(/_x[0-9A-Fa-f]{4}_/g, (m) => '_x005F' + m).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, (c) => `_x${c.charCodeAt(0).toString(16).toUpperCase().padStart(4, '0')}_`);
}
const tEl = (s: string) => {
  const e = esc(encX(s));
  return /^\s|\s$|\n/.test(s) ? `<t xml:space="preserve">${e}</t>` : `<t>${e}</t>`;
};
const argb = (c: string | undefined | null) => { const h = normColor(c || ''); return h ? 'FF' + h.slice(1).toUpperCase() : 'FF000000'; };
const num = (n: number) => (Object.is(n, -0) ? '0' : String(n));
const pxToWidth = (px: number) => Math.trunc((px / 7) * 256) / 256;
const pt = (px: number) => Math.round(px * 0.75 * 100) / 100;

const CLONE_FONTS: Record<string, string> = { Carlito: 'Calibri', Caladea: 'Cambria' };

// --------------------------------------------------------------------------------------------
// Styles

const BUILTIN_IDS = new Map<string, number>(Object.entries(BUILTIN_FORMATS).map(([k, v]) => [v, +k]));

class Styles {
  numFmts = new Map<string, number>();
  fonts: string[] = [];
  fontIdx = new Map<string, number>();
  fills: string[] = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
  fillIdx = new Map<string, number>();
  borders: string[] = ['<border><left/><right/><top/><bottom/><diagonal/></border>'];
  borderIdx = new Map<string, number>();
  xfs: string[] = [];
  xfIdx = new Map<string, number>();
  byStyle = new Map<number, number>();
  dxfs: string[] = [];
  defFont: { name: string; size: number };

  constructor(private book: Workbook) {
    const body = book.theme.bodyFont || 'Calibri';
    this.defFont = { name: CLONE_FONTS[body] || body, size: 11 };
    this.font({});
    this.xfs.push('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>');
    this.xfIdx.set(this.xfs[0], 0);
    this.byStyle.set(0, 0);
  }

  numFmtId(code: string | undefined): number {
    if (!code || code === 'General') return 0;
    const b = BUILTIN_IDS.get(code);
    if (b != null) return b;
    let id = this.numFmts.get(code);
    if (id == null) { id = 164 + this.numFmts.size; this.numFmts.set(code, id); }
    return id;
  }

  fontXml(st: Partial<CellStyle>, dxf = false): string {
    let x = '';
    if (st.bold) x += '<b/>';
    if (st.italic) x += '<i/>';
    if (st.strike) x += '<strike/>';
    if (st.underline) x += st.underline === 'single' ? '<u/>' : `<u val="${st.underline}"/>`;
    if (st.vertAlign) x += `<vertAlign val="${st.vertAlign}"/>`;
    if (dxf) { x += st.color ? `<color rgb="${argb(st.color)}"/>` : ''; return `<font>${x}</font>`; }
    x += `<sz val="${st.size || this.defFont.size}"/>`;
    x += st.color ? `<color rgb="${argb(st.color)}"/>` : '<color theme="1"/>';
    const name = st.font || this.defFont.name;
    x += `<name val="${esc(name)}"/><family val="2"/>`;
    if (!st.font) x += '<scheme val="minor"/>';
    return `<font>${x}</font>`;
  }

  font(st: Partial<CellStyle>): number {
    const x = this.fontXml(st);
    let i = this.fontIdx.get(x);
    if (i == null) { i = this.fonts.length; this.fonts.push(x); this.fontIdx.set(x, i); }
    return i;
  }

  fillXml(st: Partial<CellStyle>, dxf = false): string | null {
    if (st.gradient) {
      return `<fill><gradientFill degree="${st.gradient.angle ?? 90}"><stop position="0"><color rgb="${argb(st.gradient.from)}"/></stop><stop position="1"><color rgb="${argb(st.gradient.to)}"/></stop></gradientFill></fill>`;
    }
    if (st.pattern && st.pattern.type && st.pattern.type !== 'solid' && st.pattern.type !== 'none') {
      const fg = st.pattern.color ? `<fgColor rgb="${argb(st.pattern.color)}"/>` : '<fgColor indexed="64"/>';
      const bg = st.fill ? `<bgColor rgb="${argb(st.fill)}"/>` : '<bgColor indexed="65"/>';
      return `<fill><patternFill patternType="${esc(st.pattern.type)}">${fg}${bg}</patternFill></fill>`;
    }
    if (st.fill && st.fill !== 'none' && st.fill !== 'transparent') {
      return dxf
        ? `<fill><patternFill patternType="solid"><fgColor rgb="${argb(st.fill)}"/><bgColor rgb="${argb(st.fill)}"/></patternFill></fill>`
        : `<fill><patternFill patternType="solid"><fgColor rgb="${argb(st.fill)}"/><bgColor indexed="64"/></patternFill></fill>`;
    }
    return null;
  }

  fill(st: Partial<CellStyle>): number {
    const x = this.fillXml(st);
    if (!x) return 0;
    let i = this.fillIdx.get(x);
    if (i == null) { i = this.fills.length; this.fills.push(x); this.fillIdx.set(x, i); }
    return i;
  }

  borderXml(st: Partial<CellStyle>): string | null {
    if (!st.bl && !st.br && !st.bt && !st.bb && !st.diagUp && !st.diagDown) return null;
    const side = (n: string, b?: Border) => (b ? `<${n} style="${b.style}"><color ${b.color ? `rgb="${argb(b.color)}"` : 'indexed="64"'}/></${n}>` : `<${n}/>`);
    const diag = st.diagUp || st.diagDown;
    return `<border${st.diagUp ? ' diagonalUp="1"' : ''}${st.diagDown ? ' diagonalDown="1"' : ''}>${side('left', st.bl)}${side('right', st.br)}${side('top', st.bt)}${side('bottom', st.bb)}${side('diagonal', diag)}</border>`;
  }

  border(st: Partial<CellStyle>): number {
    const x = this.borderXml(st);
    if (!x) return 0;
    let i = this.borderIdx.get(x);
    if (i == null) { i = this.borders.length; this.borders.push(x); this.borderIdx.set(x, i); }
    return i;
  }

  alignXml(st: Partial<CellStyle>): string {
    const a: string[] = [];
    if (st.hAlign && st.hAlign !== 'general') a.push(`horizontal="${st.hAlign}"`);
    if (st.vAlign && st.vAlign !== 'bottom') a.push(`vertical="${st.vAlign === 'middle' ? 'center' : st.vAlign}"`);
    if (st.rotate) a.push(`textRotation="${st.rotate === 255 ? 255 : st.rotate < 0 ? 90 - st.rotate : Math.min(90, st.rotate)}"`);
    if (st.wrap) a.push('wrapText="1"');
    if (st.indent) a.push(`indent="${Math.max(0, Math.min(250, Math.round(st.indent)))}"`);
    if (st.shrink) a.push('shrinkToFit="1"');
    if (st.rtl) a.push('readingOrder="2"');
    return a.length ? `<alignment ${a.join(' ')}/>` : '';
  }

  xf(styleId: number): number {
    const cached = this.byStyle.get(styleId);
    if (cached != null) return cached;
    const st = this.book.styles[styleId] || {};
    const nf = this.numFmtId(st.numFmt), fo = this.font(st), fi = this.fill(st), bo = this.border(st);
    const al = this.alignXml(st);
    const pr = st.locked === false || st.hidden ? `<protection${st.locked === false ? ' locked="0"' : ''}${st.hidden ? ' hidden="1"' : ''}/>` : '';
    const attrs = `numFmtId="${nf}" fontId="${fo}" fillId="${fi}" borderId="${bo}" xfId="0"` +
      (nf ? ' applyNumberFormat="1"' : '') + (fo ? ' applyFont="1"' : '') + (fi ? ' applyFill="1"' : '') + (bo ? ' applyBorder="1"' : '') +
      (al ? ' applyAlignment="1"' : '') + (pr ? ' applyProtection="1"' : '');
    const x = al || pr ? `<xf ${attrs}>${al}${pr}</xf>` : `<xf ${attrs}/>`;
    let i = this.xfIdx.get(x);
    if (i == null) { i = this.xfs.length; this.xfs.push(x); this.xfIdx.set(x, i); }
    this.byStyle.set(styleId, i);
    return i;
  }

  dxf(st: Partial<CellStyle> | undefined): number {
    const s = st || {};
    let x = '';
    const f = this.fontXml(s, true);
    if (f !== '<font></font>') x += f;
    if (s.numFmt) x += `<numFmt numFmtId="${this.numFmtId(s.numFmt) || 164 + this.numFmts.size}" formatCode="${esc(s.numFmt)}"/>`;
    const fill = this.fillXml(s, true);
    if (fill) x += fill;
    const al = this.alignXml(s);
    if (al) x += al;
    const b = this.borderXml(s);
    if (b) x += b;
    this.dxfs.push(`<dxf>${x}</dxf>`);
    return this.dxfs.length - 1;
  }

  xml(): string {
    const nf = [...this.numFmts].map(([code, id]) => `<numFmt numFmtId="${id}" formatCode="${esc(code)}"/>`);
    return XML_HEADER + `<styleSheet xmlns="${NS_MAIN}">` +
      (nf.length ? `<numFmts count="${nf.length}">${nf.join('')}</numFmts>` : '') +
      `<fonts count="${this.fonts.length}">${this.fonts.join('')}</fonts>` +
      `<fills count="${this.fills.length}">${this.fills.join('')}</fills>` +
      `<borders count="${this.borders.length}">${this.borders.join('')}</borders>` +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      `<cellXfs count="${this.xfs.length}">${this.xfs.join('')}</cellXfs>` +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      `<dxfs count="${this.dxfs.length}">${this.dxfs.join('')}</dxfs>` +
      '<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>' +
      '</styleSheet>';
  }
}

// --------------------------------------------------------------------------------------------
// Shared strings

class SST {
  list: string[] = [];
  idx = new Map<string, number>();
  count = 0;
  add(s: string): number {
    this.count++;
    let i = this.idx.get(s);
    if (i == null) { i = this.list.length; this.list.push(`<si>${tEl(s)}</si>`); this.idx.set(s, i); }
    return i;
  }
  addRich(runs: NonNullable<Cell['rich']>): number {
    this.count++;
    const x = runs.map((r) => {
      const s = r.s || {};
      let p = '';
      if (s.font) p += `<rFont val="${esc(s.font)}"/>`;
      if (s.bold) p += '<b/>';
      if (s.italic) p += '<i/>';
      if (s.strike) p += '<strike/>';
      if (s.color) p += `<color rgb="${argb(s.color)}"/>`;
      if (s.size) p += `<sz val="${s.size}"/>`;
      if (s.underline) p += s.underline === 'single' ? '<u/>' : `<u val="${s.underline}"/>`;
      if (s.vertAlign) p += `<vertAlign val="${s.vertAlign}"/>`;
      return `<r>${p ? `<rPr>${p}</rPr>` : ''}${tEl(r.t)}</r>`;
    }).join('');
    this.list.push(`<si>${x}</si>`);
    return this.list.length - 1;
  }
  xml(): string {
    return XML_HEADER + `<sst xmlns="${NS_MAIN}" count="${this.count}" uniqueCount="${this.list.length}">${this.list.join('')}</sst>`;
  }
}

// --------------------------------------------------------------------------------------------
// Package bookkeeping

class Rels {
  items: string[] = [];
  add(type: string, target: string, external = false): string {
    const id = `rId${this.items.length + 1}`;
    this.items.push(`<Relationship Id="${id}" Type="${type}" Target="${esc(target)}"${external ? ' TargetMode="External"' : ''}/>`);
    return id;
  }
  xml(): string { return XML_HEADER + `<Relationships xmlns="${NS_PKG_REL}">${this.items.join('')}</Relationships>`; }
}

interface ExportCtx {
  book: Workbook;
  zip: JSZip;
  styles: Styles;
  sst: SST;
  overrides: string[];
  defaults: Map<string, string>;
  tableId: number;
  drawingN: number;
  chartN: number;
  imageN: number;
  commentN: number;
  usesDynamic: boolean;
  warnings: string[];
}

export interface XlsxExportOptions { kind?: 'xlsx' | 'xlsm' | 'xltx'; warnings?: string[] }

export async function exportXlsx(book: Workbook, opts: XlsxExportOptions = {}): Promise<Blob> {
  const kind = opts.kind || 'xlsx';
  // Sheet names the format cannot hold (over 31 characters, / \\ ? * [ ] :, a quote at either end, or two that then
  // collide) are changed on a copy of the workbook, through the same renaming that keeps every reference in step.
  const names = xlsxSheetNames(book.sheets.map((s) => s.name));
  if (names.some((n, i) => n !== book.sheets[i].name)) {
    book = Workbook.fromJSON(book.toJSON());
    // Two passes, so that a new name never meets a sheet that still carries it.
    book.sheets.forEach((s, i) => { if (s.name !== names[i]) renameSheet(book, s, `\u0001${i}`); });
    book.sheets.forEach((s, i) => { if (s.name !== names[i]) renameSheet(book, s, names[i]); });
    opts.warnings?.push('Some sheet names were changed to names an Excel workbook can hold.');
  }
  const zip = new JSZip();
  const ctx: ExportCtx = {
    book, zip, styles: new Styles(book), sst: new SST(), overrides: [], defaults: new Map([['rels', 'application/vnd.openxmlformats-package.relationships+xml'], ['xml', 'application/xml']]),
    tableId: 0, drawingN: 0, chartN: 0, imageN: 0, commentN: 0, usesDynamic: false, warnings: opts.warnings || [],
  };
  const wbRels = new Rels();
  // Visible/active sheet sanity: Excel needs at least one visible sheet, and the active one visible.
  const visible = book.sheets.map((s) => s.visibility === 'visible');
  if (!visible.some(Boolean)) visible[0] = true;
  let active = Math.min(Math.max(0, book.activeSheet), book.sheets.length - 1);
  if (!visible[active]) active = visible.indexOf(true);

  const sheetEntries: string[] = [];
  book.sheets.forEach((sheet, i) => {
    const path = `xl/worksheets/sheet${i + 1}.xml`;
    writeSheet(ctx, sheet, i, path, i === active);
    const rid = wbRels.add(REL + 'worksheet', `worksheets/sheet${i + 1}.xml`);
    const state = !visible[i] ? (sheet.visibility === 'veryHidden' ? ' state="veryHidden"' : ' state="hidden"') : '';
    sheetEntries.push(`<sheet name="${esc(safeSheetName(sheet.name, i))}" sheetId="${i + 1}"${state} r:id="${rid}"/>`);
    ctx.overrides.push(`<Override PartName="/${path}" ContentType="${CT}spreadsheetml.worksheet+xml"/>`);
  });

  // Workbook-level parts
  wbRels.add(REL + 'theme', 'theme/theme1.xml');
  zip.file('xl/theme/theme1.xml', themeXml(book));
  ctx.overrides.push(`<Override PartName="/xl/theme/theme1.xml" ContentType="${CT}theme+xml"/>`);
  wbRels.add(REL + 'styles', 'styles.xml');
  ctx.overrides.push(`<Override PartName="/xl/styles.xml" ContentType="${CT}spreadsheetml.styles+xml"/>`);
  if (ctx.sst.list.length) {
    wbRels.add(REL + 'sharedStrings', 'sharedStrings.xml');
    zip.file('xl/sharedStrings.xml', ctx.sst.xml());
    ctx.overrides.push(`<Override PartName="/xl/sharedStrings.xml" ContentType="${CT}spreadsheetml.sharedStrings+xml"/>`);
  }
  if (ctx.usesDynamic) {
    wbRels.add(REL + 'sheetMetadata', 'metadata.xml');
    zip.file('xl/metadata.xml', METADATA_XML);
    ctx.overrides.push(`<Override PartName="/xl/metadata.xml" ContentType="${CT}spreadsheetml.sheetMetadata+xml"/>`);
  }
  const vba: Uint8Array | undefined = (book as any).vbaProject;
  if (kind === 'xlsm' && vba instanceof Uint8Array) {
    wbRels.add('http://schemas.microsoft.com/office/2006/relationships/vbaProject', 'vbaProject.bin');
    zip.file('xl/vbaProject.bin', vba);
    ctx.defaults.set('bin', 'application/vnd.ms-office.vbaProject');
  }
  zip.file('xl/styles.xml', ctx.styles.xml());
  zip.file('xl/_rels/workbook.xml.rels', wbRels.xml());
  zip.file('xl/workbook.xml', workbookXml(book, sheetEntries, active));
  const mainCt = kind === 'xlsm' ? 'application/vnd.ms-excel.sheet.macroEnabled.main+xml' : kind === 'xltx' ? `${CT}spreadsheetml.template.main+xml` : `${CT}spreadsheetml.sheet.main+xml`;
  ctx.overrides.unshift(`<Override PartName="/xl/workbook.xml" ContentType="${mainCt}"/>`);

  // Document properties
  const rootRels = new Rels();
  rootRels.add(REL + 'officeDocument', 'xl/workbook.xml');
  rootRels.add('http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', 'docProps/core.xml');
  rootRels.add(REL + 'extended-properties', 'docProps/app.xml');
  zip.file('docProps/core.xml', coreXml(book));
  zip.file('docProps/app.xml', appXml(book));
  ctx.overrides.push('<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>');
  ctx.overrides.push(`<Override PartName="/docProps/app.xml" ContentType="${CT}extended-properties+xml"/>`);
  const custom = Object.entries(book.props.custom || {});
  if (custom.length) {
    rootRels.add(REL + 'custom-properties', 'docProps/custom.xml');
    zip.file('docProps/custom.xml', XML_HEADER + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
      custom.map(([k, v], i) => `<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="${i + 2}" name="${esc(k)}"><vt:lpwstr>${esc(v)}</vt:lpwstr></property>`).join('') + '</Properties>');
    ctx.overrides.push(`<Override PartName="/docProps/custom.xml" ContentType="${CT}custom-properties+xml"/>`);
  }
  zip.file('_rels/.rels', rootRels.xml());
  const defaults = [...ctx.defaults].map(([e, t]) => `<Default Extension="${e}" ContentType="${t}"/>`).join('');
  zip.file('[Content_Types].xml', XML_HEADER + `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${defaults}${ctx.overrides.join('')}</Types>`);

  const mime = kind === 'xlsm' ? 'application/vnd.ms-excel.sheet.macroEnabled.12' : kind === 'xltx' ? `${CT}spreadsheetml.template` : `${CT}spreadsheetml.sheet`;
  const data = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } });
  return new Blob([data as BlobPart], { type: mime });
}

/** The sheet names as an .xlsx file can hold them: legal, at most 31 characters, and unique whatever the letter case. */
export function xlsxSheetNames(names: string[]): string[] {
  const taken = new Set<string>();
  return names.map((name, i) => {
    const base = safeSheetName(name, i);
    let n = base;
    for (let k = 2; taken.has(n.toLowerCase()); k++) { const tail = ` (${k})`; n = base.slice(0, 31 - tail.length) + tail; }
    taken.add(n.toLowerCase());
    return n;
  });
}

function safeSheetName(name: string, i: number): string {
  const n = name.replace(/[\\/?*[\]:]/g, '_').replace(/^'+|'+$/g, '').slice(0, 31);
  return n || `Sheet${i + 1}`;
}

const isoDate = (t: number) => new Date(Number.isFinite(t) && t > 0 ? t : Date.now()).toISOString().replace(/\.\d+Z$/, 'Z');

function coreXml(book: Workbook): string {
  const p = book.props;
  const el = (n: string, v: string) => (v ? `<${n}>${esc(v)}</${n}>` : '');
  return XML_HEADER + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
    el('dc:title', p.title) + el('dc:subject', p.subject) + el('dc:creator', p.author) + el('cp:keywords', p.keywords) + el('dc:description', p.comments) +
    el('cp:lastModifiedBy', p.lastModifiedBy || p.author) + el('cp:category', p.category) + el('cp:contentStatus', p.status) +
    `<dcterms:created xsi:type="dcterms:W3CDTF">${isoDate(p.created)}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${isoDate(p.modified)}</dcterms:modified>` +
    '</cp:coreProperties>';
}

function appXml(book: Workbook): string {
  const p = book.props;
  return XML_HEADER + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
    '<Application>Anne\'s Grid</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>' +
    (p.manager ? `<Manager>${esc(p.manager)}</Manager>` : '') + (p.company ? `<Company>${esc(p.company)}</Company>` : '') +
    '<LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged></Properties>';
}

function workbookXml(book: Workbook, sheets: string[], active: number): string {
  const names: string[] = [];
  const sheetRef = (i: number) => quoteSheet(safeSheetName(book.sheets[i].name, i)) + '!';
  for (const n of book.names) {
    if (!n.name || !n.ref) continue;
    const local = n.sheet != null ? book.sheets.findIndex((s) => s.name === n.sheet) : -1;
    names.push(`<definedName name="${esc(n.name)}"${n.comment ? ` comment="${esc(n.comment)}"` : ''}${local >= 0 ? ` localSheetId="${local}"` : ''}${n.hidden ? ' hidden="1"' : ''}>${esc(toFileFormula(n.ref.replace(/^=/, '')))}</definedName>`);
  }
  book.sheets.forEach((s, i) => {
    if (s.filter) names.push(`<definedName name="_xlnm._FilterDatabase" localSheetId="${i}" hidden="1">${esc(sheetRef(i) + absRangeName(s.filter.range))}</definedName>`);
    if (s.print.area) {
      const parts = s.print.area.split(',').map((a) => parseRange(splitSheetRef(a).ref)).filter((r): r is Range => !!r);
      if (parts.length) names.push(`<definedName name="_xlnm.Print_Area" localSheetId="${i}">${esc(parts.map((r) => sheetRef(i) + absRangeName(r)).join(','))}</definedName>`);
    }
    const titles: string[] = [];
    if (s.print.titleCols) { const m = /^\$?([A-Z]+)(?::\$?([A-Z]+))?$/i.exec(s.print.titleCols.trim()); if (m) titles.push(`${sheetRef(i)}$${m[1]}:$${m[2] || m[1]}`); }
    if (s.print.titleRows) { const m = /^\$?(\d+)(?::\$?(\d+))?$/.exec(s.print.titleRows.trim()); if (m) titles.push(`${sheetRef(i)}$${m[1]}:$${m[2] || m[1]}`); }
    if (titles.length) names.push(`<definedName name="_xlnm.Print_Titles" localSheetId="${i}">${esc(titles.join(','))}</definedName>`);
  });
  const c = book.calc;
  const calc = `<calcPr calcId="191029"${c.mode === 'manual' ? ' calcMode="manual"' : c.mode === 'autoNoTables' ? ' calcMode="autoNoTable"' : ''}${c.iterative ? ` iterate="1" iterateCount="${c.maxIterations}" iterateDelta="${c.maxChange}"` : ''}${c.precisionAsDisplayed ? ' fullPrecision="0"' : ''}${c.recalcBeforeSave === false ? ' calcOnSave="0"' : ''} fullCalcOnLoad="1"/>`;
  return XML_HEADER + `<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_R}">` +
    (book.readOnlyRecommended ? '<fileSharing readOnlyRecommended="1"/>' : '') +
    `<workbookPr${c.date1904 ? ' date1904="1"' : ''} defaultThemeVersion="164011"/>` +
    (book.structureLocked ? `<workbookProtection lockStructure="1"${book.structurePassword && /^[0-9A-F]{4}$/i.test(book.structurePassword) ? ` workbookPassword="${book.structurePassword}"` : ''}/>` : '') +
    `<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="12300"${active ? ` activeTab="${active}"` : ''}/></bookViews>` +
    `<sheets>${sheets.join('')}</sheets>` +
    (names.length ? `<definedNames>${names.join('')}</definedNames>` : '') + calc + '</workbook>';
}

// --------------------------------------------------------------------------------------------
// Worksheet

function errCode(e: CellError): string {
  return ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#SPILL!', '#CALC!', '#GETTING_DATA'].includes(e.code) ? e.code : '#VALUE!';
}

function writeSheet(ctx: ExportCtx, sheet: Sheet, index: number, path: string, active: boolean): void {
  const { styles, sst } = ctx;
  const rels = new Rels();
  const x: string[] = [];
  x.push(XML_HEADER + `<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_R}">`);

  // sheetPr
  const fit = sheet.print.fitWide != null || sheet.print.fitTall != null;
  const prInner = (sheet.tabColor ? `<tabColor rgb="${argb(sheet.tabColor)}"/>` : '') + (fit ? '<pageSetUpPr fitToPage="1"/>' : '');
  if (prInner || sheet.codeName) x.push(`<sheetPr${sheet.codeName ? ` codeName="${esc(sheet.codeName)}"` : ''}>${prInner}</sheetPr>`.replace('<sheetPr></sheetPr>', '<sheetPr/>'));

  // Spill extents: anchor key -> range
  const spill = new Map<number, Range>();
  for (const [k, cell] of sheet.cells) {
    if (cell.sp == null || cell.f != null) continue;
    const r = keyRow(k), c = keyCol(k);
    const g = spill.get(cell.sp);
    if (!g) spill.set(cell.sp, { r1: keyRow(cell.sp), c1: keyCol(cell.sp), r2: r, c2: c });
    else { g.r2 = Math.max(g.r2, r); g.c2 = Math.max(g.c2, c); }
  }

  const used = sheet.usedRange();
  x.push(`<dimension ref="${used ? rangeName(used) : 'A1'}"/>`);

  // sheetViews
  const sv: string[] = [];
  if (active) sv.push('tabSelected="1"');
  if (!sheet.showGrid) sv.push('showGridLines="0"');
  if (!sheet.showHeadings) sv.push('showRowColHeaders="0"');
  if (!sheet.showZeros) sv.push('showZeros="0"');
  if (sheet.showFormulas) sv.push('showFormulas="1"');
  if (sheet.rtl) sv.push('rightToLeft="1"');
  if (sheet.zoom && sheet.zoom !== 1) sv.push(`zoomScale="${Math.round(Math.max(0.1, Math.min(4, sheet.zoom)) * 100)}" zoomScaleNormal="${Math.round(Math.max(0.1, Math.min(4, sheet.zoom)) * 100)}"`);
  if (sheet.scroll && (sheet.scroll.r || sheet.scroll.c) && !(sheet.freeze.rows || sheet.freeze.cols)) sv.push(`topLeftCell="${cellName(sheet.scroll.r, sheet.scroll.c)}"`);
  sv.push('workbookViewId="0"');
  let pane = '', paneName = '';
  const fr = sheet.freeze || { rows: 0, cols: 0 };
  if (fr.rows || fr.cols) {
    paneName = fr.rows && fr.cols ? 'bottomRight' : fr.rows ? 'bottomLeft' : 'topRight';
    pane = `<pane${fr.cols ? ` xSplit="${fr.cols}"` : ''}${fr.rows ? ` ySplit="${fr.rows}"` : ''} topLeftCell="${cellName(fr.rows, fr.cols)}" activePane="${paneName}" state="frozen"/>`;
  } else if (sheet.split && (sheet.split.x || sheet.split.y)) {
    paneName = sheet.split.x && sheet.split.y ? 'bottomRight' : sheet.split.y ? 'bottomLeft' : 'topRight';
    pane = `<pane${sheet.split.x ? ` xSplit="${Math.round(sheet.split.x * 15)}"` : ''}${sheet.split.y ? ` ySplit="${Math.round(sheet.split.y * 15)}"` : ''} activePane="${paneName}"/>`;
  }
  const sel = sheet.sel;
  const sqref = (sel.ranges.length ? sel.ranges : [{ r1: sel.r, c1: sel.c, r2: sel.r, c2: sel.c }]).map(rangeName).join(' ');
  const selection = `<selection${paneName ? ` pane="${paneName}"` : ''} activeCell="${cellName(sel.r, sel.c)}" sqref="${sqref}"/>`;
  x.push(`<sheetViews><sheetView ${sv.join(' ')}>${pane}${selection}</sheetView></sheetViews>`);

  // sheetFormatPr
  const maxRowLevel = Math.max(0, ...sheet.rowLevel.values());
  const maxColLevel = Math.max(0, ...sheet.colLevel.values());
  x.push(`<sheetFormatPr${sheet.defaultColW !== 64 ? ` defaultColWidth="${pxToWidth(sheet.defaultColW)}"` : ''} defaultRowHeight="${pt(sheet.defaultRowH)}"${sheet.defaultRowH !== 20 ? ' customHeight="1"' : ''}${maxRowLevel ? ` outlineLevelRow="${maxRowLevel}"` : ''}${maxColLevel ? ` outlineLevelCol="${maxColLevel}"` : ''}/>`);

  // cols
  const colSet = new Set<number>([...sheet.colW.keys(), ...sheet.hiddenCols, ...sheet.colLevel.keys(), ...sheet.colStyle.keys()]);
  if (colSet.size) {
    const cols = [...colSet].filter((c) => c >= 0 && c < MAX_COLS).sort((a, b) => a - b);
    const desc = (c: number) => {
      const w = sheet.colW.get(c);
      const s = sheet.colStyle.get(c);
      return `${w != null ? w : ''}|${sheet.hiddenCols.has(c) ? 1 : 0}|${sheet.colLevel.get(c) || 0}|${s ? styles.xf(s) : 0}`;
    };
    const out: string[] = [];
    let i = 0;
    while (i < cols.length) {
      const start = cols[i];
      const d = desc(start);
      let end = start;
      while (i + 1 < cols.length && cols[i + 1] === end + 1 && desc(cols[i + 1]) === d) { i++; end++; }
      i++;
      const [w, hid, lvl, xf] = d.split('|');
      const px = w ? +w : sheet.defaultColW;
      out.push(`<col min="${start + 1}" max="${end + 1}" width="${pxToWidth(px)}"${xf !== '0' ? ` style="${xf}"` : ''}${hid === '1' ? ' hidden="1"' : ''}${w ? ' customWidth="1"' : ''}${lvl !== '0' ? ` outlineLevel="${lvl}"` : ''}/>`);
    }
    x.push(`<cols>${out.join('')}</cols>`);
  }

  // Table header overrides (header cells must hold the column names as text)
  const headerText = new Map<number, string>();
  for (const t of sheet.tables) if (t.opts.header !== false) tableColumnNames(t).forEach((name, i) => headerText.set(cellKey(t.range.r1, t.range.c1 + i), name));

  // sheetData
  const keySet = new Set<number>(sheet.cells.keys());
  for (const k of headerText.keys()) keySet.add(k);
  const keys = [...keySet].sort((a, b) => a - b);
  const rowSet = new Set<number>();
  for (const r of sheet.rowH.keys()) rowSet.add(r);
  for (const r of sheet.hiddenRows) rowSet.add(r);
  for (const r of sheet.rowLevel.keys()) rowSet.add(r);
  for (const r of sheet.rowStyle.keys()) rowSet.add(r);
  const extraRows = [...rowSet].filter((r) => r >= 0 && r < MAX_ROWS).sort((a, b) => a - b);
  const sd: string[] = ['<sheetData>'];
  let ki = 0, ei = 0;
  let longWarned = false;
  const rowOpen = (r: number) => {
    const h = sheet.rowH.get(r);
    const s = sheet.rowStyle.get(r);
    const lvl = sheet.rowLevel.get(r);
    return `<row r="${r + 1}"${s ? ` s="${styles.xf(s)}" customFormat="1"` : ''}${h != null ? ` ht="${pt(h)}" customHeight="1"` : ''}${sheet.hiddenRows.has(r) ? ' hidden="1"' : ''}${lvl ? ` outlineLevel="${lvl}"` : ''}>`;
  };
  while (ki < keys.length || ei < extraRows.length) {
    const rk = ki < keys.length ? keyRow(keys[ki]) : Infinity;
    const re = ei < extraRows.length ? extraRows[ei] : Infinity;
    const r = Math.min(rk, re);
    if (re === r) ei++;
    const parts: string[] = [rowOpen(r)];
    while (ki < keys.length && keyRow(keys[ki]) === r) {
      const k = keys[ki++];
      const c = keyCol(k);
      const cell = sheet.cells.get(k) || {};
      const ref = colName(c) + (r + 1);
      const sAttr = cell.s ? ` s="${styles.xf(cell.s)}"` : '';
      const header = headerText.get(k);
      if (header != null && cell.f == null) { parts.push(`<c r="${ref}"${sAttr} t="s"><v>${sst.add(header)}</v></c>`); continue; }
      let v = cell.v;
      if (typeof v === 'string' && v.length > 32767) { v = v.slice(0, 32767); if (!longWarned) { ctx.warnings.push('Text longer than 32,767 characters was shortened.'); longWarned = true; } }
      if (cell.f != null && cell.sp == null) {
        let fAttr = '', cm = '';
        if (cell.arr) fAttr = ` t="array" ref="${rangeName(cell.arr)}"`;
        else if (spill.has(k)) { fAttr = ` t="array" ref="${rangeName(spill.get(k)!)}"`; cm = ' cm="1"'; ctx.usesDynamic = true; }
        const f = `<f${fAttr}>${esc(toFileFormula(cell.f))}</f>`;
        if (v == null) parts.push(`<c r="${ref}"${sAttr}${cm}>${f}</c>`);
        else if (typeof v === 'number') parts.push(`<c r="${ref}"${sAttr}${cm}>${f}<v>${Number.isFinite(v) ? num(v) : 0}</v></c>`);
        else if (typeof v === 'boolean') parts.push(`<c r="${ref}"${sAttr} t="b"${cm}>${f}<v>${v ? 1 : 0}</v></c>`);
        else if (v instanceof CellError) parts.push(`<c r="${ref}"${sAttr} t="e"${cm}>${f}<v>${esc(errCode(v))}</v></c>`);
        else parts.push(`<c r="${ref}"${sAttr} t="str"${cm}>${f}<v>${esc(encX(String(v)))}</v></c>`);
        continue;
      }
      if (v == null) { parts.push(`<c r="${ref}"${sAttr}/>`); continue; }
      if (typeof v === 'number') parts.push(Number.isFinite(v) ? `<c r="${ref}"${sAttr}><v>${num(v)}</v></c>` : `<c r="${ref}"${sAttr} t="e"><v>#NUM!</v></c>`);
      else if (typeof v === 'boolean') parts.push(`<c r="${ref}"${sAttr} t="b"><v>${v ? 1 : 0}</v></c>`);
      else if (v instanceof CellError) parts.push(`<c r="${ref}"${sAttr} t="e"><v>${esc(errCode(v))}</v></c>`);
      else {
        const idx = cell.rich && cell.rich.length && cell.sp == null ? sst.addRich(cell.rich) : sst.add(String(v));
        parts.push(`<c r="${ref}"${sAttr} t="s"><v>${idx}</v></c>`);
      }
    }
    parts.push('</row>');
    sd.push(parts.join(''));
  }
  sd.push('</sheetData>');
  x.push(sd.join(''));

  // sheetProtection
  const p = sheet.protection;
  if (p && p.on) {
    const allow = new Set(p.allow || []);
    let a = ' sheet="1"';
    const h = p.passwordHash || '';
    if (h.includes(':')) {
      const [alg, salt, spin, hash] = h.split(':');
      a += ` algorithmName="${esc(alg)}" hashValue="${esc(hash)}" saltValue="${esc(salt)}" spinCount="${esc(spin)}"`;
    } else if (/^[0-9A-F]{4}$/i.test(h)) a += ` password="${h.toUpperCase()}"`;
    for (const k of PROTECT_ACTIONS) a += ` ${k}="${allow.has(k) ? 0 : 1}"`;
    if (!allow.has('selectLocked')) a += ' selectLockedCells="1"';
    if (!allow.has('selectUnlocked')) a += ' selectUnlockedCells="1"';
    x.push(`<sheetProtection${a}/>`);
  }

  // autoFilter
  if (sheet.filter) x.push(autoFilterXml(sheet.filter));

  // mergeCells
  const merges = sheet.merges.filter((m) => m.r2 > m.r1 || m.c2 > m.c1);
  if (merges.length) x.push(`<mergeCells count="${merges.length}">${merges.map((m) => `<mergeCell ref="${rangeName(m)}"/>`).join('')}</mergeCells>`);

  // conditionalFormatting
  let prio = 1;
  for (const rule of [...sheet.cf].sort((a, b) => (a.priority ?? 0) - (b.priority ?? 0))) {
    const xml = cfRuleXml(ctx, rule, prio);
    if (!xml) continue;
    prio++;
    x.push(`<conditionalFormatting sqref="${rule.ranges.map(rangeName).join(' ')}">${xml}</conditionalFormatting>`);
  }

  // dataValidations
  const dvs = sheet.validations.filter((v) => v.ranges.length);
  if (dvs.length) x.push(`<dataValidations count="${dvs.length}">${dvs.map(dvXml).join('')}</dataValidations>`);

  // hyperlinks
  const links: string[] = [];
  for (const k of keys) {
    const cell = sheet.cells.get(k);
    if (!cell || !cell.link) continue;
    const ref = cellName(keyRow(k), keyCol(k));
    const tip = cell.linkTip ? ` tooltip="${esc(cell.linkTip)}"` : '';
    if (cell.link.startsWith('#')) links.push(`<hyperlink ref="${ref}" location="${esc(cell.link.slice(1))}" display="${esc(cell.link.slice(1))}"${tip}/>`);
    else {
      // Only the first # separates the address from the place in the document.
      const hash = cell.link.indexOf('#');
      const url = hash < 0 ? cell.link : cell.link.slice(0, hash), loc = hash < 0 ? '' : cell.link.slice(hash + 1);
      const id = rels.add(REL + 'hyperlink', url, true);
      links.push(`<hyperlink ref="${ref}" r:id="${id}"${loc ? ` location="${esc(loc)}"` : ''}${tip}/>`);
    }
  }
  if (links.length) x.push(`<hyperlinks>${links.join('')}</hyperlinks>`);

  // print
  const pr = sheet.print;
  if (pr.gridlines || pr.headings || pr.centerH || pr.centerV) {
    x.push(`<printOptions${pr.centerH ? ' horizontalCentered="1"' : ''}${pr.centerV ? ' verticalCentered="1"' : ''}${pr.headings ? ' headings="1"' : ''}${pr.gridlines ? ' gridLines="1"' : ''}/>`);
  }
  const m = pr.margins;
  x.push(`<pageMargins left="${m.left}" right="${m.right}" top="${m.top}" bottom="${m.bottom}" header="${m.header}" footer="${m.footer}"/>`);
  const paperId = Object.entries(PAPER).find(([, n]) => n === pr.paper)?.[0];
  x.push(`<pageSetup${paperId ? ` paperSize="${paperId}"` : ''}${pr.scale && pr.scale !== 100 ? ` scale="${Math.max(10, Math.min(400, Math.round(pr.scale)))}"` : ''}${pr.firstPage != null ? ` firstPageNumber="${pr.firstPage}"` : ''}${fit ? ` fitToWidth="${pr.fitWide ?? 1}" fitToHeight="${pr.fitTall ?? 0}"` : ''}${pr.order === 'overThenDown' ? ' pageOrder="overThenDown"' : ''} orientation="${pr.orientation === 'landscape' ? 'landscape' : 'portrait'}"${pr.blackAndWhite ? ' blackAndWhite="1"' : ''}${pr.draft ? ' draft="1"' : ''}${pr.firstPage != null ? ' useFirstPageNumber="1"' : ''}/>`);
  const hdr = joinHeaderFooter(pr.header), ftr = joinHeaderFooter(pr.footer);
  if (hdr || ftr) x.push(`<headerFooter>${hdr ? `<oddHeader>${esc(hdr)}</oddHeader>` : ''}${ftr ? `<oddFooter>${esc(ftr)}</oddFooter>` : ''}</headerFooter>`);
  const rb = [...new Set(sheet.rowBreaks)].filter((b) => b > 0).sort((a, b) => a - b);
  if (rb.length) x.push(`<rowBreaks count="${rb.length}" manualBreakCount="${rb.length}">${rb.map((b) => `<brk id="${b}" max="16383" man="1"/>`).join('')}</rowBreaks>`);
  const cb = [...new Set(sheet.colBreaks)].filter((b) => b > 0).sort((a, b) => a - b);
  if (cb.length) x.push(`<colBreaks count="${cb.length}" manualBreakCount="${cb.length}">${cb.map((b) => `<brk id="${b}" max="1048575" man="1"/>`).join('')}</colBreaks>`);

  // drawing (images, shapes, charts)
  if (sheet.drawings.length || sheet.charts.length) {
    const d = writeDrawing(ctx, sheet, index);
    if (d) x.push(`<drawing r:id="${rels.add(REL + 'drawing', `../drawings/drawing${d}.xml`)}"/>`);
  }
  // legacy comments
  if (sheet.comments.size) {
    const n = ++ctx.commentN;
    writeComments(ctx, sheet, n, index);
    rels.add(REL + 'comments', `../comments${n}.xml`);
    x.push(`<legacyDrawing r:id="${rels.add(REL + 'vmlDrawing', `../drawings/vmlDrawing${n}.vml`)}"/>`);
  }
  // tables
  if (sheet.tables.length) {
    const ids = sheet.tables.map((t) => {
      const n = ++ctx.tableId;
      ctx.zip.file(`xl/tables/table${n}.xml`, tableXml(t, n, sheet));
      ctx.overrides.push(`<Override PartName="/xl/tables/table${n}.xml" ContentType="${CT}spreadsheetml.table+xml"/>`);
      return rels.add(REL + 'table', `../tables/table${n}.xml`);
    });
    x.push(`<tableParts count="${ids.length}">${ids.map((id) => `<tablePart r:id="${id}"/>`).join('')}</tableParts>`);
  }
  // sparklines
  if (sheet.sparklines.length) x.push(sparklinesXml(sheet));

  x.push('</worksheet>');
  ctx.zip.file(path, x.join(''));
  if (rels.items.length) ctx.zip.file(`xl/worksheets/_rels/sheet${index + 1}.xml.rels`, rels.xml());
}

const HF_BACK: [RegExp, string][] = [[/&\[Page\]/g, '&P'], [/&\[Pages\]/g, '&N'], [/&\[Date\]/g, '&D'], [/&\[Time\]/g, '&T'], [/&\[Path\]/g, '&Z'], [/&\[File\]/g, '&F'], [/&\[Tab\]/g, '&A'], [/&\[Picture\]/g, '&G']];
function hfCodes(s: string): string { let o = s; for (const [re, c] of HF_BACK) o = o.replace(re, c); return o; }
function joinHeaderFooter(h: { left: string; center: string; right: string }): string {
  return (h.left ? '&L' + hfCodes(h.left) : '') + (h.center ? '&C' + hfCodes(h.center) : '') + (h.right ? '&R' + hfCodes(h.right) : '');
}

function autoFilterXml(f: NonNullable<Sheet['filter']>): string {
  const cols = Object.entries(f.columns).map(([id, c]) => {
    let inner = '';
    if (c.values || c.blanks) inner = `<filters${c.blanks ? ' blank="1"' : ''}>${(c.values || []).map((v) => `<filter val="${esc(v)}"/>`).join('')}</filters>`;
    else if (c.top10) inner = `<top10${c.top10.top ? '' : ' top="0"'}${c.top10.percent ? ' percent="1"' : ''} val="${c.top10.n}"/>`;
    else if (c.custom && c.custom.length) inner = `<customFilters${c.and ? ' and="1"' : ''}>${c.custom.slice(0, 2).map((x) => `<customFilter operator="${esc(x.op)}" val="${esc(x.value)}"/>`).join('')}</customFilters>`;
    else if (c.dynamic) inner = `<dynamicFilter type="${c.dynamic}"/>`;
    return inner ? `<filterColumn colId="${id}">${inner}</filterColumn>` : '';
  }).join('');
  let sort = '';
  if (f.sort) {
    const body = { r1: f.range.r1 + 1, c1: f.range.c1, r2: f.range.r2, c2: f.range.c2 };
    if (body.r2 >= body.r1) {
      const col = f.range.c1 + f.sort.col;
      sort = `<sortState ref="${rangeName(body)}"><sortCondition${f.sort.desc ? ' descending="1"' : ''} ref="${rangeName({ r1: body.r1, c1: col, r2: body.r2, c2: col })}"/></sortState>`;
    }
  }
  return `<autoFilter ref="${rangeName(f.range)}">${cols}${sort}</autoFilter>`;
}

function cfvoXml(s: NonNullable<CFRule['stops']>[number]): string {
  const v = s.value != null && s.type !== 'min' && s.type !== 'max' ? ` val="${esc(String(s.value))}"` : '';
  return `<cfvo type="${s.type}"${v}/>`;
}

const strLit = (s: string) => '"' + s.replace(/"/g, '""') + '"';

function cfRuleXml(ctx: ExportCtx, r: CFRule, priority: number): string | null {
  if (!r.ranges.length) return null;
  const tl = cellName(r.ranges[0].r1, r.ranges[0].c1);
  const a: string[] = [`type="${r.type}"`];
  const needsDxf = !['colorScale', 'dataBar', 'iconSet'].includes(r.type);
  if (needsDxf) a.push(`dxfId="${ctx.styles.dxf(r.style)}"`);
  a.push(`priority="${priority}"`);
  if (r.stopIfTrue) a.push('stopIfTrue="1"');
  let formulas = (r.formulas || []).map((f) => f.replace(/^=/, ''));
  const t = r.text ?? '';
  switch (r.type) {
    case 'cellIs': a.push(`operator="${r.operator || 'equal'}"`); break;
    case 'top10': if (r.percent) a.push('percent="1"'); if (r.bottom) a.push('bottom="1"'); a.push(`rank="${r.rank ?? 10}"`); break;
    case 'aboveAverage': if (r.above === false) a.push('aboveAverage="0"'); if (r.equalAverage) a.push('equalAverage="1"'); if (r.stdDev) a.push(`stdDev="${r.stdDev}"`); break;
    case 'containsText': a.push('operator="containsText"', `text="${esc(t)}"`); if (!formulas.length) formulas = [`NOT(ISERROR(SEARCH(${strLit(t)},${tl})))`]; break;
    case 'notContainsText': a.push('operator="notContains"', `text="${esc(t)}"`); if (!formulas.length) formulas = [`ISERROR(SEARCH(${strLit(t)},${tl}))`]; break;
    case 'beginsWith': a.push('operator="beginsWith"', `text="${esc(t)}"`); if (!formulas.length) formulas = [`LEFT(${tl},LEN(${strLit(t)}))=${strLit(t)}`]; break;
    case 'endsWith': a.push('operator="endsWith"', `text="${esc(t)}"`); if (!formulas.length) formulas = [`RIGHT(${tl},LEN(${strLit(t)}))=${strLit(t)}`]; break;
    case 'containsBlanks': if (!formulas.length) formulas = [`LEN(TRIM(${tl}))=0`]; break;
    case 'notContainsBlanks': if (!formulas.length) formulas = [`LEN(TRIM(${tl}))>0`]; break;
    case 'containsErrors': if (!formulas.length) formulas = [`ISERROR(${tl})`]; break;
    case 'notContainsErrors': if (!formulas.length) formulas = [`NOT(ISERROR(${tl}))`]; break;
    case 'timePeriod': {
      const tp = r.timePeriod || 'today';
      a.push(`timePeriod="${tp}"`);
      if (!formulas.length) formulas = [TIME_PERIOD[tp]?.replace(/A1/g, tl) || `FLOOR(${tl},1)=TODAY()`];
      break;
    }
  }
  let inner = formulas.slice(0, 3).map((f) => `<formula>${esc(toFileFormula(f))}</formula>`).join('');
  if (r.type === 'cellIs' && !inner) return null;
  if (r.type === 'expression' && !inner) return null;
  if (r.type === 'colorScale') {
    const colors = r.colors && r.colors.length >= 2 ? r.colors.slice(0, 3) : ['#f8696b', '#63be7b'];
    let stops = r.stops && r.stops.length === colors.length ? r.stops : colors.length === 3 ? [{ type: 'min' as const }, { type: 'percentile' as const, value: 50 }, { type: 'max' as const }] : [{ type: 'min' as const }, { type: 'max' as const }];
    stops = stops.slice(0, colors.length);
    inner = `<colorScale>${stops.map(cfvoXml).join('')}${colors.map((c) => `<color rgb="${argb(c)}"/>`).join('')}</colorScale>`;
  } else if (r.type === 'dataBar') {
    const stops = r.stops && r.stops.length >= 2 ? r.stops.slice(0, 2) : [{ type: 'min' as const }, { type: 'max' as const }];
    inner = `<dataBar${r.showValue === false ? ' showValue="0"' : ''}>${stops.map(cfvoXml).join('')}<color rgb="${argb(r.barColor || '#638ec6')}"/></dataBar>`;
  } else if (r.type === 'iconSet') {
    const set = r.iconSet || '3TrafficLights1';
    const n = parseInt(set, 10) || 3;
    const stops = r.stops && r.stops.length === n ? r.stops : Array.from({ length: n }, (_, i) => ({ type: 'percent' as const, value: Math.round((i * 100) / n) }));
    inner = `<iconSet${set !== '3TrafficLights1' ? ` iconSet="${esc(set)}"` : ''}${r.reverseIcons ? ' reverse="1"' : ''}${r.showValue === false ? ' showValue="0"' : ''}>${stops.map(cfvoXml).join('')}</iconSet>`;
  }
  return `<cfRule ${a.join(' ')}>${inner}</cfRule>`;
}

const TIME_PERIOD: Record<string, string> = {
  today: 'FLOOR(A1,1)=TODAY()', yesterday: 'FLOOR(A1,1)=TODAY()-1', tomorrow: 'FLOOR(A1,1)=TODAY()+1',
  last7Days: 'AND(TODAY()-FLOOR(A1,1)<=6,FLOOR(A1,1)<=TODAY())',
  thisWeek: 'AND(TODAY()-ROUNDDOWN(A1,0)<=WEEKDAY(TODAY())-1,ROUNDDOWN(A1,0)-TODAY()<=7-WEEKDAY(TODAY()))',
  lastWeek: 'AND(TODAY()-ROUNDDOWN(A1,0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(A1,0)<(WEEKDAY(TODAY())+7))',
  nextWeek: 'AND(ROUNDDOWN(A1,0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN(A1,0)-TODAY()<(15-WEEKDAY(TODAY())))',
  thisMonth: 'AND(MONTH(A1)=MONTH(TODAY()),YEAR(A1)=YEAR(TODAY()))',
  lastMonth: 'AND(MONTH(A1)=MONTH(EDATE(TODAY(),0-1)),YEAR(A1)=YEAR(EDATE(TODAY(),0-1)))',
  nextMonth: 'AND(MONTH(A1)=MONTH(EDATE(TODAY(),0+1)),YEAR(A1)=YEAR(EDATE(TODAY(),0+1)))',
};

function dvXml(v: DataValidation): string {
  const a: string[] = [];
  if (v.type !== 'any') a.push(`type="${v.type}"`);
  if (v.errorStyle && v.errorStyle !== 'stop') a.push(`errorStyle="${v.errorStyle}"`);
  if (v.operator && v.operator !== 'between' && v.type !== 'list' && v.type !== 'custom' && v.type !== 'any') a.push(`operator="${v.operator}"`);
  if (v.allowBlank) a.push('allowBlank="1"');
  if (v.type === 'list' && v.showDropdown === false) a.push('showDropDown="1"');
  if (v.inputTitle || v.inputMessage) a.push('showInputMessage="1"');
  if (v.type !== 'any') a.push('showErrorMessage="1"');
  if (v.errorTitle) a.push(`errorTitle="${esc(v.errorTitle.slice(0, 32))}"`);
  if (v.errorMessage) a.push(`error="${esc(v.errorMessage.slice(0, 255))}"`);
  if (v.inputTitle) a.push(`promptTitle="${esc(v.inputTitle.slice(0, 32))}"`);
  if (v.inputMessage) a.push(`prompt="${esc(v.inputMessage.slice(0, 255))}"`);
  a.push(`sqref="${v.ranges.map(rangeName).join(' ')}"`);
  const f1 = v.formula1 != null && v.formula1 !== '' ? `<formula1>${esc(toFileFormula(v.formula1.replace(/^=/, '')))}</formula1>` : '';
  const f2 = v.formula2 != null && v.formula2 !== '' ? `<formula2>${esc(toFileFormula(v.formula2.replace(/^=/, '')))}</formula2>` : '';
  return `<dataValidation ${a.join(' ')}>${f1}${f2}</dataValidation>`;
}

/** Unique, non-empty column names for a table (Excel requires them to match the header cells). */
function tableColumnNames(t: TableDef): string[] {
  const ncols = t.range.c2 - t.range.c1 + 1;
  const seen = new Set<string>();
  return Array.from({ length: ncols }, (_, i) => {
    let n = (t.columns[i]?.name || `Column${i + 1}`).replace(/[\r\n]+/g, ' ');
    const base = n;
    let k = 2;
    while (seen.has(n.toLowerCase())) n = base + k++;
    seen.add(n.toLowerCase());
    return n;
  });
}

function tableXml(t: TableDef, id: number, sheet: Sheet): string {
  const name = (t.name || `Table${id}`).replace(/[^\p{L}\p{N}_.\\]/gu, '_').replace(/^(\p{N})/u, '_$1');
  const ncols = t.range.c2 - t.range.c1 + 1;
  const cols = Array.from({ length: ncols }, (_, i) => t.columns[i] || { name: `Column${i + 1}` });
  const names = tableColumnNames(t);
  const colXml = cols.map((c, i) => {
    const n = names[i];
    let a = `id="${i + 1}" name="${esc(n)}"`;
    let inner = '';
    if (c.calcFormula) inner += `<calculatedColumnFormula>${esc(toFileFormula(c.calcFormula.replace(/^=/, '')))}</calculatedColumnFormula>`;
    if (t.opts.totals) {
      const fn = c.totalFn;
      const totalCell = sheet.get(t.range.r2, t.range.c1 + i);
      if (fn && fn !== 'none' && fn !== 'label') {
        const known = ['sum', 'min', 'max', 'average', 'count', 'countNums', 'stdDev', 'var'];
        if (known.includes(fn)) a += ` totalsRowFunction="${fn}"`;
        else if (c.totalFormula) { a += ' totalsRowFunction="custom"'; inner += `<totalsRowFormula>${esc(toFileFormula(c.totalFormula.replace(/^=/, '')))}</totalsRowFormula>`; }
      } else if (c.totalFormula && fn !== 'label') {
        a += ' totalsRowFunction="custom"';
        inner += `<totalsRowFormula>${esc(toFileFormula(c.totalFormula.replace(/^=/, '')))}</totalsRowFormula>`;
      } else if (totalCell && typeof totalCell.v === 'string' && totalCell.f == null) a += ` totalsRowLabel="${esc(totalCell.v)}"`;
    }
    return inner ? `<tableColumn ${a}>${inner}</tableColumn>` : `<tableColumn ${a}/>`;
  }).join('');
  const dataEnd = t.opts.totals ? t.range.r2 - 1 : t.range.r2;
  const af = t.opts.header !== false && t.opts.filterButton !== false ? `<autoFilter ref="${rangeName({ ...t.range, r2: Math.max(t.range.r1, dataEnd) })}"/>` : '';
  return XML_HEADER + `<table xmlns="${NS_MAIN}" id="${id}" name="${esc(name)}" displayName="${esc(name)}" ref="${rangeName(t.range)}"${t.opts.header === false ? ' headerRowCount="0"' : ''}${t.opts.totals ? ' totalsRowCount="1"' : ' totalsRowShown="0"'}>` +
    af + `<tableColumns count="${ncols}">${colXml}</tableColumns>` +
    `<tableStyleInfo name="${esc(t.style || 'TableStyleMedium2')}" showFirstColumn="${t.opts.firstCol ? 1 : 0}" showLastColumn="${t.opts.lastCol ? 1 : 0}" showRowStripes="${t.opts.banded ? 1 : 0}" showColumnStripes="${t.opts.bandedCols ? 1 : 0}"/></table>`;
}

function sparklinesXml(sheet: Sheet): string {
  const groups = sheet.sparklines.map((s) => {
    const type = s.type === 'column' ? ' type="column"' : s.type === 'winloss' ? ' type="stacked"' : '';
    let data = s.data.replace(/^=/, '');
    if (!splitSheetRef(data).sheet) data = quoteSheet(sheet.name) + '!' + data;
    return `<x14:sparklineGroup${type} displayEmptyCellsAs="gap"${s.markers ? ' markers="1"' : ''}${s.high ? ' high="1"' : ''}${s.low ? ' low="1"' : ''}${s.negative ? ' negative="1"' : ''}>` +
      `<x14:colorSeries rgb="${argb(s.color || '#376092')}"/><x14:colorNegative rgb="FFD00000"/><x14:colorAxis rgb="FF000000"/><x14:colorMarkers rgb="${argb(s.color || '#376092')}"/><x14:colorFirst rgb="FFD00000"/><x14:colorLast rgb="FFD00000"/><x14:colorHigh rgb="FFD00000"/><x14:colorLow rgb="FFD00000"/>` +
      `<x14:sparklines><x14:sparkline><xm:f>${esc(data)}</xm:f><xm:sqref>${cellName(s.cell.r, s.cell.c)}</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup>`;
  }).join('');
  return '<extLst><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">' +
    `<x14:sparklineGroups>${groups}</x14:sparklineGroups></ext></extLst>`;
}

// --------------------------------------------------------------------------------------------
// Comments (legacy notes with a VML drawing)

function writeComments(ctx: ExportCtx, sheet: Sheet, n: number, sheetIndex: number): void {
  const authors: string[] = [];
  const aid = (a: string) => { let i = authors.indexOf(a); if (i < 0) { i = authors.length; authors.push(a); } return i; };
  const list: string[] = [];
  const shapes: string[] = [];
  let k2 = 0;
  const entries = [...sheet.comments].sort((a, b) => a[0] - b[0]);
  for (const [k, cm] of entries) {
    const r = keyRow(k), c = keyCol(k);
    let text = cm.text || '';
    for (const rep of cm.replies || []) text += `\n${rep.author ? rep.author + ': ' : ''}${rep.text}`;
    const author = cm.author || '';
    const runs = author
      ? `<r><rPr><b/><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><family val="2"/></rPr>${tEl(author + ':')}</r><r><rPr><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><family val="2"/></rPr>${tEl('\n' + text)}</r>`
      : `<r><rPr><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><family val="2"/></rPr>${tEl(text)}</r>`;
    list.push(`<comment ref="${cellName(r, c)}" authorId="${aid(author)}"><text>${runs}</text></comment>`);
    const id = 1024 * (sheetIndex + 1) + ++k2;
    const w = cm.w || 144, h = cm.h || 79;
    shapes.push(`<v:shape id="_x0000_s${id}" type="#_x0000_t202" style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:${pt(w)}pt;height:${pt(h)}pt;z-index:${k2};visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto">` +
      '<v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>' +
      `<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>${c + 1}, 15, ${Math.max(0, r - 1)}, 10, ${c + 3}, 15, ${r + 3}, 4</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>${r}</x:Row><x:Column>${c}</x:Column></x:ClientData></v:shape>`);
  }
  ctx.zip.file(`xl/comments${n}.xml`, XML_HEADER + `<comments xmlns="${NS_MAIN}"><authors>${authors.map((a) => `<author>${esc(a)}</author>`).join('')}</authors><commentList>${list.join('')}</commentList></comments>`);
  ctx.overrides.push(`<Override PartName="/xl/comments${n}.xml" ContentType="${CT}spreadsheetml.comments+xml"/>`);
  ctx.defaults.set('vml', `${CT}vmlDrawing`);
  ctx.zip.file(`xl/drawings/vmlDrawing${n}.vml`, '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">' +
    `<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="${sheetIndex + 1}"/></o:shapelayout>` +
    '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>' +
    shapes.join('') + '</xml>');
}

// --------------------------------------------------------------------------------------------
// Drawings and charts

function cellAt(sheet: Sheet, r: number, c: number, dx: number, dy: number): { r: number; c: number; dx: number; dy: number } {
  // Normalize offsets that run past the cell
  while (dx >= sheet.colWidth(c) && sheet.colWidth(c) > 0 && c < MAX_COLS - 1) { dx -= sheet.colWidth(c); c++; }
  while (dy >= sheet.rowHeight(r) && sheet.rowHeight(r) > 0 && r < MAX_ROWS - 1) { dy -= sheet.rowHeight(r); r++; }
  return { r, c, dx: Math.max(0, dx), dy: Math.max(0, dy) };
}

function anchorXml(sheet: Sheet, at: DrawingObject['at'], body: string): string {
  const from = cellAt(sheet, at.r, at.c, at.dx || 0, at.dy || 0);
  const to = cellAt(sheet, from.r, from.c, from.dx + Math.max(1, at.w), from.dy + Math.max(1, at.h));
  const pos = (p: typeof from) => `<xdr:col>${p.c}</xdr:col><xdr:colOff>${Math.round(p.dx * EMU)}</xdr:colOff><xdr:row>${p.r}</xdr:row><xdr:rowOff>${Math.round(p.dy * EMU)}</xdr:rowOff>`;
  return `<xdr:twoCellAnchor editAs="oneCell"><xdr:from>${pos(from)}</xdr:from><xdr:to>${pos(to)}</xdr:to>${body}<xdr:clientData/></xdr:twoCellAnchor>`;
}

function writeDrawing(ctx: ExportCtx, sheet: Sheet, _sheetIndex: number): number | null {
  const n = ++ctx.drawingN;
  const rels = new Rels();
  const parts: string[] = [];
  let shapeId = 1;
  const xfrm = (o: { at: DrawingObject['at']; rotate?: number }) => `<a:xfrm${o.rotate ? ` rot="${Math.round(o.rotate * 60000)}"` : ''}><a:off x="0" y="0"/><a:ext cx="${Math.round(o.at.w * EMU)}" cy="${Math.round(o.at.h * EMU)}"/></a:xfrm>`;
  for (const d of sheet.drawings) {
    const id = ++shapeId;
    if (d.kind === 'image' || d.kind === 'signature' || (d.src && d.src.startsWith('data:image/') && d.kind !== 'shape' && d.kind !== 'textbox')) {
      const img = d.src ? parseDataUrl(d.src) : null;
      if (!img || !/^image\/(png|jpeg|gif|bmp|tiff|x-emf|x-wmf)$/.test(img.mime)) { ctx.warnings.push('Some pictures could not be saved in .xlsx format.'); continue; }
      const ext = extForMime(img.mime);
      const file = `image${++ctx.imageN}.${ext}`;
      ctx.zip.file(`xl/media/${file}`, img.bytes);
      ctx.defaults.set(ext, img.mime);
      const rid = rels.add(REL + 'image', `../media/${file}`);
      if (!d.at.w || !d.at.h) { const sz = imageSize(img.bytes); if (sz) d.at = { ...d.at, w: d.at.w || sz.w, h: d.at.h || sz.h }; }
      parts.push(anchorXml(sheet, d.at, `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${id}" name="Picture ${id - 1}"${d.alt ? ` descr="${esc(d.alt)}"` : ''}/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>` +
        `<xdr:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr>${xfrm(d)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>`));
      continue;
    }
    if (d.kind === 'shape' || d.kind === 'textbox' || d.kind === 'wordart') {
      const fill = d.fill === 'none' ? '<a:noFill/>' : d.fill ? `<a:solidFill><a:srgbClr val="${argb(d.fill).slice(2)}"/></a:solidFill>` : d.kind === 'textbox' ? '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>' : '<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>';
      const ln = d.stroke === 'none' ? '<a:ln><a:noFill/></a:ln>' : d.stroke || d.strokeWidth ? `<a:ln w="${Math.round((d.strokeWidth || 1) * EMU)}"><a:solidFill><a:srgbClr val="${argb(d.stroke || '#000000').slice(2)}"/></a:solidFill></a:ln>` : '';
      const f = d.font || {};
      const rpr = `<a:rPr lang="en-US" sz="${Math.round((f.size || 11) * 100)}"${f.bold ? ' b="1"' : ''}${f.italic ? ' i="1"' : ''}>${f.color ? `<a:solidFill><a:srgbClr val="${argb(f.color).slice(2)}"/></a:solidFill>` : ''}${f.font ? `<a:latin typeface="${esc(f.font)}"/>` : ''}</a:rPr>`;
      const paras = (d.text || '').split('\n').map((p) => (p ? `<a:p><a:r>${rpr}<a:t>${esc(p)}</a:t></a:r></a:p>` : `<a:p><a:endParaRPr lang="en-US"/></a:p>`)).join('');
      const geom = PRESET_SHAPES.has(d.shape || '') ? d.shape : 'rect';
      parts.push(anchorXml(sheet, d.at, `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${id}" name="${d.kind === 'textbox' ? 'TextBox' : 'Shape'} ${id - 1}"${d.alt ? ` descr="${esc(d.alt)}"` : ''}/><xdr:cNvSpPr${d.kind === 'textbox' ? ' txBox="1"' : ''}/></xdr:nvSpPr>` +
        `<xdr:spPr>${xfrm(d)}<a:prstGeom prst="${geom}"><a:avLst/></a:prstGeom>${fill}${ln}</xdr:spPr>` +
        `<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="square" rtlCol="0" anchor="t"/><a:lstStyle/>${paras || '<a:p><a:endParaRPr lang="en-US"/></a:p>'}</xdr:txBody></xdr:sp>`));
      continue;
    }
    ctx.warnings.push(`Drawing objects of kind "${d.kind}" are not saved in .xlsx files.`);
  }
  for (const ch of sheet.charts) {
    const id = ++shapeId;
    const cn = ++ctx.chartN;
    ctx.zip.file(`xl/charts/chart${cn}.xml`, chartXml(ctx.book, sheet, ch));
    ctx.overrides.push(`<Override PartName="/xl/charts/chart${cn}.xml" ContentType="${CT}drawingml.chart+xml"/>`);
    const rid = rels.add(REL + 'chart', `../charts/chart${cn}.xml`);
    parts.push(anchorXml(sheet, ch.at, `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="Chart ${id - 1}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
      `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="${NS_C}"><c:chart xmlns:c="${NS_C}" r:id="${rid}"/></a:graphicData></a:graphic></xdr:graphicFrame>`));
  }
  if (!parts.length) { ctx.drawingN--; return null; }
  ctx.zip.file(`xl/drawings/drawing${n}.xml`, XML_HEADER + `<xdr:wsDr xmlns:xdr="${NS_XDR}" xmlns:a="${NS_A}" xmlns:r="${NS_R}">${parts.join('')}</xdr:wsDr>`);
  if (rels.items.length) ctx.zip.file(`xl/drawings/_rels/drawing${n}.xml.rels`, rels.xml());
  ctx.overrides.push(`<Override PartName="/xl/drawings/drawing${n}.xml" ContentType="${CT}drawing+xml"/>`);
  return n;
}

const PRESET_SHAPES = new Set(['rect', 'roundRect', 'ellipse', 'triangle', 'rtTriangle', 'diamond', 'pentagon', 'hexagon', 'octagon', 'star5', 'star4', 'star6', 'rightArrow', 'leftArrow', 'upArrow', 'downArrow', 'leftRightArrow', 'line', 'parallelogram', 'trapezoid', 'heart', 'cloud', 'smileyFace', 'can', 'cube', 'plus', 'flowChartProcess', 'flowChartDecision', 'wedgeRectCallout', 'wedgeRoundRectCallout', 'chevron', 'homePlate', 'donut', 'noSmoking', 'blockArc', 'lightningBolt', 'sun', 'moon', 'arc', 'bracketPair', 'bracePair']);

const PALETTE = ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '264478', '9E480E', '636363', '997300'];

/** Values for a chart range reference, for the cached data Excel and viewers show before recalculating. */
function refValues(book: Workbook, sheet: Sheet, ref: string): { values: Scalar[]; full: string } | null {
  const { sheet: sn, ref: body } = splitSheetRef(ref);
  const sh = sn ? book.sheetByName(sn) : sheet;
  const rg = parseRange(body);
  if (!sh || !rg) return null;
  const values: Scalar[] = [];
  const cells = (rg.r2 - rg.r1 + 1) * (rg.c2 - rg.c1 + 1);
  if (cells > 10000) return { values: [], full: quoteSheet(sh.name) + '!' + absRangeName(rg) };
  for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) values.push(sh.value(r, c));
  return { values, full: quoteSheet(sh.name) + '!' + absRangeName(rg) };
}

function literal(ref: string): Scalar[] | null {
  const m = /^\{(.*)\}$/s.exec(ref.trim());
  if (!m) return null;
  return m[1].split(/[,;]/).map((x) => { const t = x.trim(); if (/^".*"$/.test(t)) return t.slice(1, -1).replace(/""/g, '"'); const n = Number(t); return Number.isFinite(n) ? n : t; });
}

function numData(book: Workbook, sheet: Sheet, ref: string | undefined, tag: string): string {
  if (!ref) return '';
  const lit = literal(ref);
  const pts = (vals: Scalar[]) => vals.map((v, i) => (typeof v === 'number' ? `<c:pt idx="${i}"><c:v>${num(v)}</c:v></c:pt>` : '')).join('');
  if (lit) return `<c:${tag}><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="${lit.length}"/>${pts(lit)}</c:numLit></c:${tag}>`;
  const rv = refValues(book, sheet, ref);
  if (!rv) return `<c:${tag}><c:numRef><c:f>${esc(toFileFormula(ref))}</c:f></c:numRef></c:${tag}>`;
  return `<c:${tag}><c:numRef><c:f>${esc(rv.full)}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${rv.values.length}"/>${pts(rv.values)}</c:numCache></c:numRef></c:${tag}>`;
}

function strData(book: Workbook, sheet: Sheet, ref: string | undefined, tag: string): string {
  if (!ref) return '';
  const lit = literal(ref);
  const pts = (vals: Scalar[]) => vals.map((v, i) => (v == null ? '' : `<c:pt idx="${i}"><c:v>${esc(String(v))}</c:v></c:pt>`)).join('');
  if (lit) return `<c:${tag}><c:strLit><c:ptCount val="${lit.length}"/>${pts(lit)}</c:strLit></c:${tag}>`;
  const rv = refValues(book, sheet, ref);
  if (!rv) return `<c:${tag}><c:strRef><c:f>${esc(toFileFormula(ref))}</c:f></c:strRef></c:${tag}>`;
  return `<c:${tag}><c:strRef><c:f>${esc(rv.full)}</c:f><c:strCache><c:ptCount val="${rv.values.length}"/>${pts(rv.values)}</c:strCache></c:strRef></c:${tag}>`;
}

const richTitle = (t: string) => `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/>${t.split('\n').map((p) => `<a:p><a:r><a:t>${esc(p)}</a:t></a:r></a:p>`).join('')}</c:rich></c:tx><c:overlay val="0"/></c:title>`;

type Family = 'bar' | 'line' | 'pie' | 'doughnut' | 'area' | 'scatter' | 'radar';
function family(t: ChartType): { fam: Family; dir?: 'col' | 'bar'; grouping?: string; markers?: boolean; filled?: boolean; lines?: boolean } {
  switch (t) {
    case 'column': case 'histogram': case 'pareto': case 'waterfall': case 'boxWhisker': case 'funnel': case 'treemap': case 'sunburst': case 'combo': return { fam: 'bar', dir: 'col', grouping: 'clustered' };
    case 'stackedColumn': return { fam: 'bar', dir: 'col', grouping: 'stacked' };
    case 'percentColumn': return { fam: 'bar', dir: 'col', grouping: 'percentStacked' };
    case 'bar': return { fam: 'bar', dir: 'bar', grouping: 'clustered' };
    case 'stackedBar': return { fam: 'bar', dir: 'bar', grouping: 'stacked' };
    case 'percentBar': return { fam: 'bar', dir: 'bar', grouping: 'percentStacked' };
    case 'line': case 'stock': return { fam: 'line', grouping: 'standard', markers: false };
    case 'lineMarkers': return { fam: 'line', grouping: 'standard', markers: true };
    case 'stackedLine': return { fam: 'line', grouping: 'stacked', markers: false };
    case 'area': return { fam: 'area', grouping: 'standard' };
    case 'stackedArea': return { fam: 'area', grouping: 'stacked' };
    case 'percentArea': return { fam: 'area', grouping: 'percentStacked' };
    case 'pie': return { fam: 'pie' };
    case 'doughnut': return { fam: 'doughnut' };
    case 'scatter': case 'bubble': return { fam: 'scatter', lines: false };
    case 'scatterLines': return { fam: 'scatter', lines: true };
    case 'radar': return { fam: 'radar' };
    case 'filledRadar': return { fam: 'radar', filled: true };
  }
  return { fam: 'bar', dir: 'col', grouping: 'clustered' };
}

export function chartXml(book: Workbook, sheet: Sheet, ch: ChartDef): string {
  const qualify = (ref: string | undefined) => {
    if (!ref) return ref;
    const r = ref.replace(/^=/, '');
    if (r.startsWith('{') || splitSheetRef(r).sheet) return r;
    return quoteSheet(sheet.name) + '!' + r;
  };
  // Group series by plot type
  const groups = new Map<ChartType, { s: ChartSeries; i: number }[]>();
  ch.series.forEach((s, i) => {
    const t = ch.type === 'combo' ? s.type || 'column' : ch.type;
    if (!groups.has(t)) groups.set(t, []);
    groups.get(t)!.push({ s, i });
  });
  if (!groups.size) groups.set(ch.type === 'combo' ? 'column' : ch.type, []);
  const dLbls = ch.dataLabels ? '<c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>' : '';
  let usesCatAx = false, usesScatter = false, horizontal = false, needsAxes = false;
  const plots: string[] = [];
  let trendDone = false;
  for (const [t, list] of groups) {
    const f = family(t);
    const ser = list.map(({ s, i }) => {
      const color = (s.color ? argb(s.color).slice(2) : null) || (ch.colors?.[i] ? argb(ch.colors[i]).slice(2) : PALETTE[i % PALETTE.length]);
      const tx = s.nameRef ? `<c:tx>${strData(book, sheet, qualify(s.nameRef), 'tx').replace(/^<c:tx>|<\/c:tx>$/g, '')}</c:tx>` : s.name ? `<c:tx><c:v>${esc(s.name)}</c:v></c:tx>` : '';
      const head = `<c:idx val="${i}"/><c:order val="${i}"/>${tx}`;
      const fillSp = `<c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></c:spPr>`;
      const lineSp = `<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:round/></a:ln></c:spPr>`;
      const cats = s.categories ? strData(book, sheet, qualify(s.categories), 'cat') : '';
      const vals = numData(book, sheet, qualify(s.values), 'val');
      let trend = '';
      if (ch.trendline && !trendDone && (f.fam === 'bar' || f.fam === 'line' || f.fam === 'scatter' || f.fam === 'area')) {
        const tt = { linear: 'linear', exponential: 'exp', polynomial: 'poly', movingAverage: 'movingAvg' }[ch.trendline];
        trend = `<c:trendline><c:trendlineType val="${tt}"/>${tt === 'poly' ? '<c:order val="2"/>' : tt === 'movingAvg' ? '<c:period val="2"/>' : ''}<c:dispRSqr val="0"/><c:dispEq val="0"/></c:trendline>`;
        trendDone = true;
      }
      switch (f.fam) {
        case 'bar': return `<c:ser>${head}${fillSp}<c:invertIfNegative val="0"/>${trend}${cats}${vals}</c:ser>`;
        case 'line': return `<c:ser>${head}${lineSp}${f.markers ? '<c:marker><c:symbol val="circle"/><c:size val="5"/></c:marker>' : '<c:marker><c:symbol val="none"/></c:marker>'}${trend}${cats}${vals}<c:smooth val="0"/></c:ser>`;
        case 'area': return `<c:ser>${head}${fillSp}${trend}${cats}${vals}</c:ser>`;
        case 'pie': case 'doughnut': return `<c:ser>${head}${cats}${vals}</c:ser>`;
        case 'radar': return `<c:ser>${head}${f.filled ? fillSp : lineSp}${f.filled ? '' : '<c:marker><c:symbol val="none"/></c:marker>'}${cats}${vals}</c:ser>`;
        case 'scatter': {
          const sp = f.lines ? lineSp : `<c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr>`;
          const xv = s.categories ? numData(book, sheet, qualify(s.categories), 'xVal') : '';
          const yv = numData(book, sheet, qualify(s.values), 'yVal');
          return `<c:ser>${head}${sp}<c:marker><c:symbol val="circle"/><c:size val="5"/><c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></c:spPr></c:marker>${trend}${xv}${yv}<c:smooth val="0"/></c:ser>`;
        }
      }
      return '';
    }).join('');
    const ax = '<c:axId val="500000001"/><c:axId val="500000002"/>';
    switch (f.fam) {
      case 'bar':
        usesCatAx = true; needsAxes = true; if (f.dir === 'bar') horizontal = true;
        plots.push(`<c:barChart><c:barDir val="${f.dir}"/><c:grouping val="${f.grouping}"/><c:varyColors val="0"/>${ser}${dLbls}<c:gapWidth val="150"/>${f.grouping !== 'clustered' ? '<c:overlap val="100"/>' : ''}${ax}</c:barChart>`);
        break;
      case 'line':
        usesCatAx = true; needsAxes = true;
        plots.push(`<c:lineChart><c:grouping val="${f.grouping}"/><c:varyColors val="0"/>${ser}${dLbls}<c:marker val="1"/>${ax}</c:lineChart>`);
        break;
      case 'area':
        usesCatAx = true; needsAxes = true;
        plots.push(`<c:areaChart><c:grouping val="${f.grouping}"/><c:varyColors val="0"/>${ser}${dLbls}${ax}</c:areaChart>`);
        break;
      case 'pie':
        plots.push(`<c:pieChart><c:varyColors val="1"/>${ser}${dLbls}<c:firstSliceAng val="0"/></c:pieChart>`);
        break;
      case 'doughnut':
        plots.push(`<c:doughnutChart><c:varyColors val="1"/>${ser}${dLbls}<c:firstSliceAng val="0"/><c:holeSize val="50"/></c:doughnutChart>`);
        break;
      case 'radar':
        usesCatAx = true; needsAxes = true;
        plots.push(`<c:radarChart><c:radarStyle val="${f.filled ? 'filled' : 'marker'}"/><c:varyColors val="0"/>${ser}${dLbls}${ax}</c:radarChart>`);
        break;
      case 'scatter':
        usesScatter = true; needsAxes = true;
        plots.push(`<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${ser}${dLbls}${ax}</c:scatterChart>`);
        break;
    }
  }
  let axes = '';
  if (needsAxes) {
    const tick = '<c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>';
    const scaling = `<c:scaling>${ch.logScale ? '<c:logBase val="10"/>' : ''}<c:orientation val="minMax"/>${ch.yMax != null ? `<c:max val="${ch.yMax}"/>` : ''}${ch.yMin != null ? `<c:min val="${ch.yMin}"/>` : ''}</c:scaling>`;
    const grid = ch.gridlines !== false ? '<c:majorGridlines/>' : '';
    const xt = ch.xTitle ? richTitle(ch.xTitle) : '', yt = ch.yTitle ? richTitle(ch.yTitle) : '';
    const xPos = horizontal ? 'l' : 'b', yPos = horizontal ? 'b' : 'l';
    if (usesScatter && !usesCatAx) {
      axes = `<c:valAx><c:axId val="500000001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/>${xt}<c:numFmt formatCode="General" sourceLinked="1"/>${tick}<c:crossAx val="500000002"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>` +
        `<c:valAx><c:axId val="500000002"/>${scaling}<c:delete val="0"/><c:axPos val="l"/>${grid}${yt}<c:numFmt formatCode="General" sourceLinked="1"/>${tick}<c:crossAx val="500000001"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>`;
    } else {
      axes = `<c:catAx><c:axId val="500000001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${xPos}"/>${xt}<c:numFmt formatCode="General" sourceLinked="1"/>${tick}<c:crossAx val="500000002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>` +
        `<c:valAx><c:axId val="500000002"/>${scaling}<c:delete val="0"/><c:axPos val="${yPos}"/>${grid}${yt}<c:numFmt formatCode="General" sourceLinked="1"/>${tick}<c:crossAx val="500000001"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`;
    }
  }
  const legendPos = ({ right: 'r', bottom: 'b', top: 't', left: 'l' } as Record<string, string>)[ch.legend || 'right'];
  const legend = ch.legend === 'none' ? '' : `<c:legend><c:legendPos val="${legendPos || 'r'}"/><c:overlay val="0"/></c:legend>`;
  return XML_HEADER + `<c:chartSpace xmlns:c="${NS_C}" xmlns:a="${NS_A}" xmlns:r="${NS_R}"><c:roundedCorners val="0"/>${ch.style ? `<c:style val="${Math.max(1, Math.min(48, ch.style))}"/>` : ''}<c:chart>` +
    (ch.title ? richTitle(ch.title) + '<c:autoTitleDeleted val="0"/>' : '<c:autoTitleDeleted val="1"/>') +
    `<c:plotArea><c:layout/>${plots.join('')}${axes}</c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>` +
    '<c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/><c:pageSetup/></c:printSettings></c:chartSpace>';
}

// --------------------------------------------------------------------------------------------
// Theme and metadata

function themeXml(book: Workbook): string {
  const c = book.theme.colors.map((x) => (normColor(x) || '#000000').slice(1).toUpperCase());
  const names = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
  const clr = names.map((n, i) => (i === 0 ? `<a:dk1><a:sysClr val="windowText" lastClr="${c[0] || '000000'}"/></a:dk1>` : i === 1 ? `<a:lt1><a:sysClr val="window" lastClr="${c[1] || 'FFFFFF'}"/></a:lt1>` : `<a:${n}><a:srgbClr val="${c[i] || '000000'}"/></a:${n}>`)).join('');
  const head = CLONE_FONTS[book.theme.headingFont] || book.theme.headingFont || 'Calibri Light';
  const body = CLONE_FONTS[book.theme.bodyFont] || book.theme.bodyFont || 'Calibri';
  const solid = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
  const ln = (w: number) => `<a:ln w="${w}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>`;
  return XML_HEADER + `<a:theme xmlns:a="${NS_A}" name="${esc(book.theme.name || 'Office Theme')}"><a:themeElements>` +
    `<a:clrScheme name="${esc(book.theme.name || 'Office')}">${clr}</a:clrScheme>` +
    `<a:fontScheme name="${esc(book.theme.name || 'Office')}"><a:majorFont><a:latin typeface="${esc(head)}"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="${esc(body)}"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>` +
    `<a:fmtScheme name="Office"><a:fillStyleLst>${solid}${solid}${solid}</a:fillStyleLst><a:lnStyleLst>${ln(6350)}${ln(12700)}${ln(19050)}</a:lnStyleLst>` +
    '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>' +
    `<a:bgFillStyleLst>${solid}${solid}${solid}</a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`;
}

const METADATA_XML = XML_HEADER + `<metadata xmlns="${NS_MAIN}" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">` +
  '<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes>' +
  '<futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>' +
  '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>';
