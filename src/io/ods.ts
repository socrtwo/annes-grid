/**
 * OpenDocument Spreadsheet import/export: .ods (ZIP package) and .fods (flat XML).
 *
 * Covers cell values and types, formulas (OpenFormula "of:=" syntax with [.A1] references,
 * converted to and from our A1 syntax), cell styles (fonts, fills, borders, alignment, number
 * formats as data styles, best effort), column widths, row heights, hidden rows/columns, outline
 * groups, merged cells, comments (annotations), named ranges, print ranges, autofilter ranges,
 * frozen panes (settings.xml), sheet visibility/tab color/protection and pictures.
 */
import JSZip from 'jszip';
import { Workbook, Sheet, newId } from '../model/workbook';
import { CellError, errorFromCode, type Cell, type CellStyle, type Border, type DrawingObject } from '../model/types';
import { cellKey, keyRow, keyCol, colName, colIndex, quoteSheet, splitSheetRef, parseRange, MAX_COLS, MAX_ROWS, type Range } from '../model/address';
import { isDateFormat, isPercentFormat } from '../model/numfmt';
import { partsToSerial, serialToParts } from '../model/dates';
import { tokenize, parseRefText, type Token } from '../calc/parser';
import { parseXml, kids, kid, desc, esc, XML_HEADER, normColor, decodeText, parseDataUrl, extForMime, mimeForExt, bytesToBase64, stripXlfn } from './index';
import { readValidations, applyValidations, noteValidation, noteStyleMap, readConditionalFormats, styleMapRules, readChart } from './ods-extras';

const NS = {
  office: 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
  style: 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
  text: 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
  table: 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
  draw: 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0',
  fo: 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0',
  xlink: 'http://www.w3.org/1999/xlink',
  dc: 'http://purl.org/dc/elements/1.1/',
  meta: 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0',
  number: 'urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0',
  svg: 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0',
  of: 'urn:oasis:names:tc:opendocument:xmlns:of:1.2',
  config: 'urn:oasis:names:tc:opendocument:xmlns:config:1.0',
  tableooo: 'http://openoffice.org/2009/table',
  calcext: 'urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0',
  loext: 'urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0',
  manifest: 'urn:oasis:names:tc:opendocument:xmlns:manifest:1.0',
} as const;
type Pfx = keyof typeof NS;

/** Namespaced attribute with a prefixed-name fallback. */
function a(el: Element | null | undefined, ns: Pfx, local: string): string | null {
  if (!el) return null;
  return el.getAttributeNS(NS[ns], local) ?? el.getAttribute(`${ns}:${local}`);
}

// --------------------------------------------------------------------------------------------
// Units

/** ODF length -> px at 96 dpi. */
export function lengthToPx(v: string | null | undefined): number | null {
  if (!v) return null;
  const m = /^\s*(-?[\d.]+)\s*(in|cm|mm|pt|pc|px)?\s*$/i.exec(v);
  if (!m) return null;
  const n = parseFloat(m[1]);
  switch ((m[2] || 'px').toLowerCase()) {
    case 'in': return n * 96;
    case 'cm': return (n * 96) / 2.54;
    case 'mm': return (n * 96) / 25.4;
    case 'pt': return (n * 4) / 3;
    case 'pc': return n * 16;
    default: return n;
  }
}
const pxToIn = (px: number) => `${Math.round((px / 96) * 10000) / 10000}in`;

// --------------------------------------------------------------------------------------------
// Formulas

function quoteOdfSheet(name: string): string {
  return /^[A-Za-z_][A-Za-z0-9_]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`;
}

/** Converts one ODF reference body (inside [ ]) to A1 syntax. */
export function odfRefToA1(body: string): string {
  const parts: string[] = [];
  let cur = '', q = false;
  for (const ch of body) {
    if (ch === "'") q = !q;
    if (ch === ':' && !q) { parts.push(cur); cur = ''; continue; }
    cur += ch;
  }
  parts.push(cur);
  let sheet: string | null = null, sheet2: string | null = null;
  const cells = parts.map((p, i) => {
    let s = p.trim();
    let qq = false, dot = -1;
    for (let k = 0; k < s.length; k++) { if (s[k] === "'") qq = !qq; else if (s[k] === '.' && !qq) dot = k; }
    let sh = '';
    if (dot >= 0) { sh = s.slice(0, dot); s = s.slice(dot + 1); }
    sh = sh.replace(/^\$/, '');
    if (sh.startsWith("'")) sh = sh.slice(1, -1).replace(/''/g, "'");
    if (sh) { if (i === 0) sheet = sh; else if (sh !== sheet) sheet2 = sh; }
    return s;
  });
  if (cells.some((c) => /#REF!/i.test(c) || !c)) return '#REF!';
  // Whole columns/rows written as A$1:A$1048576
  if (cells.length === 2) {
    const m1 = /^(\$?)([A-Z]+)\$?1$/i.exec(cells[0]), m2 = /^(\$?)([A-Z]+)\$?1048576$/i.exec(cells[1]);
    if (m1 && m2) cells.splice(0, 2, m1[1] + m1[2], m2[1] + m2[2]);
    const r1 = /^\$?A(\$?)(\d+)$/i.exec(cells[0]), r2 = /^\$?(?:XFD|AMJ)(\$?)(\d+)$/i.exec(cells[1]);
    if (r1 && r2) cells.splice(0, 2, r1[1] + r1[2], r2[1] + r2[2]);
  }
  const pre = sheet ? quoteSheet(sheet) + (sheet2 ? ':' + quoteSheet(sheet2) : '') + '!' : '';
  return pre + cells.join(':');
}

/** OpenFormula ("of:=SUM([.A1:.B2];2)") -> our formula text ("SUM(A1:B2,2)"). */
export function odfToFormula(f: string): string {
  let s = f.trim().replace(/^(?:of|oooc|msoxl|ooow)\s*:/i, '');
  if (s.startsWith('=')) s = s.slice(1);
  let out = '', braces = 0;
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (ch === '"') {
      let j = i + 1;
      while (j < s.length) { if (s[j] === '"') { if (s[j + 1] === '"') { j += 2; continue; } break; } j++; }
      out += s.slice(i, j + 1); i = j; continue;
    }
    if (ch === '[') {
      let j = i + 1, q = false;
      while (j < s.length && (s[j] !== ']' || q)) { if (s[j] === "'") q = !q; j++; }
      out += odfRefToA1(s.slice(i + 1, j)); i = j; continue;
    }
    if (ch === '{') braces++;
    if (ch === '}') braces--;
    if (ch === ';') { out += ','; continue; }
    if (ch === '|' && braces > 0) { out += ';'; continue; }
    if (ch === '~') { out += ','; continue; }
    out += ch;
  }
  return stripXlfn(out.replace(/\bCOM\.MICROSOFT\./gi, '').replace(/\bORG\.LIBREOFFICE\./gi, '').replace(/\bORG\.OPENOFFICE\./gi, ''));
}

function a1ToOdfRef(t: Token): string {
  const r = t.ref!;
  if (r.invalid) return '[.#REF!]';
  const sh = r.sheet ? '$' + quoteOdfSheet(r.sheet) + '.' : '.';
  const cell = (row: number, col: number, ar: boolean, ac: boolean) => `${ac ? '$' : ''}${colName(col)}${ar ? '$' : ''}${row + 1}`;
  switch (r.kind) {
    case 'cell': return `[${sh}${cell(r.r1, r.c1, r.ar1, r.ac1)}]`;
    case 'cols': return `[${sh}${r.ac1 ? '$' : ''}${colName(r.c1)}$1:.${r.ac2 ? '$' : ''}${colName(r.c2)}$1048576]`;
    case 'rows': return `[${sh}$A${r.ar1 ? '$' : ''}${r.r1 + 1}:.$XFD${r.ar2 ? '$' : ''}${r.r2 + 1}]`;
    default: {
      const sh2 = r.sheet2 ? '$' + quoteOdfSheet(r.sheet2) + '.' : '.';
      return `[${sh}${cell(r.r1, r.c1, r.ar1, r.ac1)}:${sh2}${cell(r.r2, r.c2, r.ar2, r.ac2)}]`;
    }
  }
}

/**
 * Resolves a structured reference (Table1[Col], Table1[[#Headers],[A]:[B]], [@Col]) to an A1
 * range, for formats that have no table references.
 */
export function resolveStructuredRef(book: Workbook, text: string, sheet?: Sheet, row?: number): string | null {
  const m = /^([A-Za-z_\\][\w.]*)?\[(.*)\]$/s.exec(text.trim());
  if (!m) return null;
  let table: { t: import('../model/types').TableDef; sh: Sheet } | null = null;
  for (const sh of book.sheets) for (const t of sh.tables) {
    if (m[1] ? t.name.toLowerCase() === m[1].toLowerCase() : sh === sheet && row != null && row >= t.range.r1 && row <= t.range.r2 && t.range.c1 >= 0) table = table || { t, sh };
  }
  if (!table) return null;
  const { t, sh } = table;
  let inner = m[2].trim();
  let thisRow = false;
  if (inner.startsWith('@')) { thisRow = true; inner = inner.slice(1).trim(); }
  const items: string[] = [];
  if (inner.includes('[')) {
    const re = /\[((?:'.|[^\]])*)\]|(:)/g;
    let mm: RegExpExecArray | null;
    while ((mm = re.exec(inner))) items.push(mm[2] ? ':' : mm[1]);
  } else if (inner) items.push(inner);
  const unesc = (x: string) => x.replace(/'(.)/g, '$1').trim().toLowerCase();
  const header = t.opts.header !== false ? 1 : 0, totals = t.opts.totals ? 1 : 0;
  let r1 = t.range.r1 + header, r2 = t.range.r2 - totals;
  const specials = items.filter((x) => x.startsWith('#')).map((x) => x.toLowerCase());
  if (thisRow || specials.includes('#this row')) { if (row == null) return null; r1 = r2 = row; }
  else if (specials.length) {
    const rows: number[] = [];
    for (const sp of specials) {
      if (sp === '#all') rows.push(t.range.r1, t.range.r2);
      else if (sp === '#headers' && header) rows.push(t.range.r1);
      else if (sp === '#totals' && totals) rows.push(t.range.r2);
      else if (sp === '#data') rows.push(t.range.r1 + header, t.range.r2 - totals);
    }
    if (!rows.length) return null;
    r1 = Math.min(...rows); r2 = Math.max(...rows);
  }
  const names = t.columns.map((c) => c.name.trim().toLowerCase());
  const colItems = items.filter((x) => !x.startsWith('#'));
  let c1 = t.range.c1, c2 = t.range.c2;
  if (colItems.length) {
    const idx = colItems.filter((x) => x !== ':').map((x) => names.indexOf(unesc(x)));
    if (idx.some((i) => i < 0)) return null;
    c1 = t.range.c1 + Math.min(...idx); c2 = t.range.c1 + Math.max(...idx);
  }
  const ref = r1 === r2 && c1 === c2 ? `$${colName(c1)}$${r1 + 1}` : `$${colName(c1)}$${r1 + 1}:$${colName(c2)}$${r2 + 1}`;
  return sh === sheet ? ref : quoteSheet(sh.name) + '!' + ref;
}

/** Our formula text -> OpenFormula with the of:= prefix. */
export function formulaToOdf(f: string, ctx?: { book: Workbook; sheet: Sheet; r: number }): string {
  let toks: Token[];
  try { toks = tokenize(f); } catch { return 'of:=' + f; }
  let out = '', braces = 0, last = 0;
  for (const t of toks) {
    if (t.type === 'eof') break;
    out += f.slice(last, t.pos);
    last = t.end;
    switch (t.type) {
      case 'ref': out += a1ToOdfRef(t); break;
      case 'sref': {
        const res = ctx ? resolveStructuredRef(ctx.book, t.text, ctx.sheet, ctx.r) : null;
        const ref = res ? parseRefText(res) : null;
        out += ref ? a1ToOdfRef({ ...t, ref }) : t.text;
        break;
      }
      case 'lbrace': braces++; out += '{'; break;
      case 'rbrace': braces--; out += '}'; break;
      case 'comma': out += ';'; break;
      case 'semi': out += braces > 0 ? '|' : ';'; break;
      default: out += f.slice(t.pos, t.end);
    }
  }
  out += f.slice(last);
  return 'of:=' + out;
}

// --------------------------------------------------------------------------------------------
// Import

interface StyleDefs {
  cell: Map<string, Element>;
  col: Map<string, Element>;
  row: Map<string, Element>;
  table: Map<string, Element>;
  data: Map<string, Element>;
  fonts: Map<string, string>;
  resolved: Map<string, number>;
  defaultStyle: Element | null;
  base: CellStyle | null;
}

function collectStyles(docs: (Document | null)[]): StyleDefs {
  const d: StyleDefs = { cell: new Map(), col: new Map(), row: new Map(), table: new Map(), data: new Map(), fonts: new Map(), resolved: new Map(), defaultStyle: null, base: null };
  for (const doc of docs) {
    if (!doc) continue;
    for (const ff of desc(doc, 'font-face')) {
      const n = a(ff, 'style', 'name');
      const fam = a(ff, 'svg', 'font-family');
      if (n) d.fonts.set(n, (fam || n).replace(/^['"]|['"]$/g, ''));
    }
    for (const ds of desc(doc, 'default-style')) if (a(ds, 'style', 'family') === 'table-cell') d.defaultStyle = ds;
    for (const st of desc(doc, 'style')) {
      if (st.namespaceURI !== NS.style && st.prefix !== 'style') continue;
      const name = a(st, 'style', 'name');
      if (!name) continue;
      const fam = a(st, 'style', 'family');
      if (fam === 'table-cell') d.cell.set(name, st);
      else if (fam === 'table-column') d.col.set(name, st);
      else if (fam === 'table-row') d.row.set(name, st);
      else if (fam === 'table') d.table.set(name, st);
    }
    for (const kind of ['number-style', 'percentage-style', 'currency-style', 'date-style', 'time-style', 'boolean-style', 'text-style']) {
      for (const ds of desc(doc, kind)) { const n = a(ds, 'style', 'name'); if (n) d.data.set(n, ds); }
    }
  }
  return d;
}

function odfBorder(v: string | null): Border | undefined {
  if (!v || v === 'none' || /hidden/.test(v)) return undefined;
  const w = lengthToPx(/([\d.]+(?:pt|in|cm|mm|px))/.exec(v)?.[1]) ?? 1;
  const style: Border['style'] = /double/.test(v) ? 'double' : /dashed/.test(v) ? 'dashed' : /dotted/.test(v) ? 'dotted' : w > 2.8 ? 'thick' : w > 1.6 ? 'medium' : 'thin';
  const color = normColor(/#[0-9a-f]{6}/i.exec(v)?.[0]) || undefined;
  return { style, ...(color ? { color } : {}) };
}

/** Data style element -> spreadsheet format code (best effort). */
function dataStyleCode(defs: StyleDefs, name: string, depth = 0): string | undefined {
  const ds = defs.data.get(name);
  if (!ds || depth > 3) return undefined;
  const kind = ds.localName;
  let code = '';
  const quote = (t: string) => (t === '' ? '' : /^[\s$\-+/():,.%€£¥]+$/.test(t) ? t.replace(/%/g, '"%"') : `"${t.replace(/"/g, '')}"`);
  const timeKind = kind === 'time-style';
  let sawHour = false;
  for (const ch of kids(ds)) {
    const long = a(ch, 'number', 'style') === 'long';
    switch (ch.localName) {
      case 'number': {
        const dp = +(a(ch, 'number', 'decimal-places') ?? 0);
        const minInt = +(a(ch, 'number', 'min-integer-digits') ?? 1);
        const grp = a(ch, 'number', 'grouping') === 'true';
        let int = minInt > 0 ? '0'.repeat(minInt) : '#';
        if (grp) int = '#,##' + (int.length >= 1 ? int.slice(-1) : '0');
        code += int + (dp > 0 ? '.' + '0'.repeat(dp) : '');
        break;
      }
      case 'scientific-number': {
        const dp = +(a(ch, 'number', 'decimal-places') ?? 2);
        const ed = +(a(ch, 'number', 'min-exponent-digits') ?? 2);
        code += '0' + (dp ? '.' + '0'.repeat(dp) : '') + 'E+' + '0'.repeat(Math.max(1, ed));
        break;
      }
      case 'fraction': {
        const nd = +(a(ch, 'number', 'min-numerator-digits') ?? 1), dd = +(a(ch, 'number', 'min-denominator-digits') ?? 1);
        const dv = a(ch, 'number', 'denominator-value');
        code += '# ' + '?'.repeat(nd) + '/' + (dv || '?'.repeat(dd));
        break;
      }
      case 'text': {
        const t = ch.textContent || '';
        code += kind === 'percentage-style' && t === '%' ? '%' : quote(t);
        break;
      }
      case 'currency-symbol': {
        const sym = ch.textContent || '$';
        code += sym === '$' ? '$' : `[$${sym}]`;
        break;
      }
      case 'text-content': code += '@'; break;
      case 'year': code += long ? 'yyyy' : 'yy'; break;
      case 'month': code += a(ch, 'number', 'textual') === 'true' ? (long ? 'mmmm' : 'mmm') : long ? 'mm' : 'm'; break;
      case 'day': code += long ? 'dd' : 'd'; break;
      case 'day-of-week': code += long ? 'dddd' : 'ddd'; break;
      case 'hours': code += timeKind && a(ds, 'number', 'truncate-on-overflow') === 'false' && !sawHour ? '[h]' : long ? 'hh' : 'h'; sawHour = true; break;
      case 'minutes': code += long ? 'mm' : 'm'; break;
      case 'seconds': { const dp = +(a(ch, 'number', 'decimal-places') ?? 0); code += (long ? 'ss' : 's') + (dp ? '.' + '0'.repeat(dp) : ''); break; }
      case 'am-pm': code += 'AM/PM'; break;
      case 'boolean': return undefined;
    }
  }
  const color = kids(ds, 'text-properties')[0];
  const col = color ? a(color, 'fo', 'color') : null;
  if (col && /^#ff0000$/i.test(col)) code = '[Red]' + code;
  // style:map conditions: positive/zero part comes from the mapped style.
  const maps = kids(ds, 'map');
  if (maps.length) {
    const parts = maps.map((m) => dataStyleCode(defs, a(m, 'style', 'apply-style-name') || '', depth + 1)).filter((x): x is string => !!x);
    if (parts.length) code = [...parts, code].join(';');
  }
  return code || undefined;
}

function cellStyleOf(defs: StyleDefs, name: string, depth = 0): CellStyle {
  const el = defs.cell.get(name);
  if (!el || depth > 8) return {};
  const parent = a(el, 'style', 'parent-style-name');
  const st: CellStyle = parent ? { ...cellStyleOf(defs, parent, depth + 1) } : {};
  const tp = kid(el, 'text-properties');
  if (tp) {
    const fw = a(tp, 'fo', 'font-weight');
    if (fw) { if (fw === 'bold' || +fw >= 600) st.bold = true; else delete st.bold; }
    const fs = a(tp, 'fo', 'font-style');
    if (fs) { if (fs === 'italic' || fs === 'oblique') st.italic = true; else delete st.italic; }
    const color = normColor(a(tp, 'fo', 'color'));
    if (color) st.color = color;
    const size = lengthToPx(a(tp, 'fo', 'font-size'));
    if (size) st.size = Math.round(size * 0.75 * 2) / 2;
    const fn = a(tp, 'style', 'font-name');
    const ff = a(tp, 'fo', 'font-family');
    if (fn) st.font = defs.fonts.get(fn) || fn; else if (ff) st.font = ff.replace(/^['"]|['"]$/g, '');
    const ul = a(tp, 'style', 'text-underline-style');
    if (ul) { if (ul !== 'none') st.underline = a(tp, 'style', 'text-underline-type') === 'double' ? 'double' : 'single'; else delete st.underline; }
    const lt = a(tp, 'style', 'text-line-through-style');
    if (lt) { if (lt !== 'none') st.strike = true; else delete st.strike; }
    const pos = a(tp, 'style', 'text-position');
    if (pos && /^super/.test(pos)) st.vertAlign = 'superscript'; else if (pos && /^sub/.test(pos)) st.vertAlign = 'subscript';
  }
  const cp = kid(el, 'table-cell-properties');
  if (cp) {
    const bg = a(cp, 'fo', 'background-color');
    if (bg) { const c = normColor(bg); if (c) st.fill = c; else if (bg === 'transparent') delete st.fill; }
    const all = odfBorder(a(cp, 'fo', 'border'));
    const side = (k: 'top' | 'bottom' | 'left' | 'right') => odfBorder(a(cp, 'fo', `border-${k}`)) || all;
    const bt = side('top'), bb = side('bottom'), bl = side('left'), br = side('right');
    if (bt) st.bt = bt; if (bb) st.bb = bb; if (bl) st.bl = bl; if (br) st.br = br;
    const du = odfBorder(a(cp, 'style', 'diagonal-bl-tr')); if (du) st.diagUp = du;
    const dd = odfBorder(a(cp, 'style', 'diagonal-tl-br')); if (dd) st.diagDown = dd;
    if (a(cp, 'fo', 'wrap-option') === 'wrap') st.wrap = true;
    const va = a(cp, 'style', 'vertical-align');
    if (va === 'top' || va === 'bottom') st.vAlign = va; else if (va === 'middle') st.vAlign = 'middle';
    const rot = +(a(cp, 'style', 'rotation-angle') || 0);
    if (rot) st.rotate = rot > 90 && rot <= 180 ? 90 : rot > 180 ? Math.max(-90, rot - 360) : rot;
    if (a(cp, 'style', 'shrink-to-fit') === 'true') st.shrink = true;
    const prot = a(cp, 'style', 'cell-protect');
    if (prot) {
      if (prot === 'none') st.locked = false; else delete st.locked;
      if (/formula-hidden|hidden-and-protected/.test(prot)) st.hidden = true;
    }
  }
  const pp = kid(el, 'paragraph-properties');
  if (pp) {
    const ta = a(pp, 'fo', 'text-align');
    if (ta) st.hAlign = ta === 'center' ? 'center' : ta === 'end' || ta === 'right' ? 'right' : ta === 'justify' ? 'justify' : 'left';
    const ml = lengthToPx(a(pp, 'fo', 'margin-left'));
    if (ml && ml > 1) st.indent = Math.round(ml / 9);
  }
  const ds = a(el, 'style', 'data-style-name');
  if (ds) { const code = dataStyleCode(defs, ds); if (code && code !== 'General') st.numFmt = code; }
  return st;
}

/** Document defaults (default-style plus the "Default" cell style), dropped from every cell style. */
function baseStyle(defs: StyleDefs): CellStyle {
  if (defs.base) return defs.base;
  let base: CellStyle = {};
  if (defs.defaultStyle) {
    const tmp = 'q:default-style';
    defs.cell.set(tmp, defs.defaultStyle);
    base = cellStyleOf(defs, tmp);
    defs.cell.delete(tmp);
  }
  defs.base = { ...base, ...cellStyleOf(defs, 'Default') };
  return defs.base;
}

function styleIdFor(book: Workbook, defs: StyleDefs, name: string | null): number {
  if (!name || name === 'Default') return 0;
  let id = defs.resolved.get(name);
  if (id == null) {
    const st = cellStyleOf(defs, name);
    const base = baseStyle(defs);
    for (const k of ['font', 'size', 'color', 'numFmt'] as const) if (st[k] != null && st[k] === base[k]) delete st[k];
    if (st.color === '#000000') delete st.color;
    if (st.vAlign === 'bottom') delete st.vAlign;
    id = book.styleId(st);
    defs.resolved.set(name, id);
  }
  return id;
}

function odfText(el: Element): string {
  let out = '';
  const walk = (n: Node) => {
    if (n.nodeType === 3) { out += n.nodeValue || ''; return; }
    if (n.nodeType !== 1) return;
    const e = n as Element;
    switch (e.localName) {
      case 's': out += ' '.repeat(+(a(e, 'text', 'c') || 1)); return;
      case 'tab': out += '\t'; return;
      case 'line-break': out += '\n'; return;
      case 'annotation': case 'note': return;
    }
    e.childNodes.forEach(walk);
  };
  walk(el);
  return out;
}

function parseIsoDuration(s: string): number | null {
  const m = /^-?P(?:(\d+)D)?T?(?:(\d+)H)?(?:(\d+)M)?(?:([\d.]+)S)?$/.exec(s.trim());
  if (!m) return null;
  const v = (+(m[1] || 0)) + ((+(m[2] || 0)) * 3600 + (+(m[3] || 0)) * 60 + (+(m[4] || 0))) / 86400;
  return s.trim().startsWith('-') ? -v : v;
}

/** Splits a list of cell range addresses on white space, but not inside a quoted sheet name ($'My Sheet'.A1). */
export function splitOdfAddresses(s: string): string[] {
  const out: string[] = [];
  let cur = '', q = false;
  for (const ch of s) {
    if (ch === "'") q = !q;
    if (!q && /\s/.test(ch)) { if (cur) out.push(cur); cur = ''; } else cur += ch;
  }
  if (cur) out.push(cur);
  return out;
}

function parseIsoDate(s: string, date1904: boolean): number | null {
  const m = /^(-?\d{4,})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?)?/.exec(s.trim());
  if (!m) return null;
  // Before 1900 there is no calendar serial to borrow: count days from the 1899-12-30 epoch (a time or small number
  // carrying a date format is written this way), instead of letting the year be read as 1900 + y.
  const day = +m[1] < 1900 && !date1904 ? Math.round((Date.UTC(+m[1], +m[2] - 1, +m[3]) - Date.UTC(1899, 11, 30)) / 86400000) : partsToSerial(+m[1], +m[2], +m[3], date1904);
  return day + ((+(m[4] || 0)) * 3600 + (+(m[5] || 0)) * 60 + (+(m[6] || 0))) / 86400;
}

export interface OdfDocs { content: Document; styles: Document | null; settings: Document | null; meta: Document | null; files?: JSZip }

export async function importOds(data: ArrayBuffer, warnings: string[] = []): Promise<Workbook> {
  const zip = await JSZip.loadAsync(data);
  const read = async (p: string) => { const f = zip.file(p); return f ? parseXml(decodeText(await f.async('uint8array')), p) : null; };
  const content = await read('content.xml');
  if (!content) throw new Error('content.xml is missing from this OpenDocument file');
  const soft = async (p: string) => { try { return await read(p); } catch { warnings.push(`${p} is damaged and was skipped.`); return null; } };
  return readOdf({ content, styles: await soft('styles.xml'), settings: await soft('settings.xml'), meta: await soft('meta.xml'), files: zip }, warnings);
}

export async function importFods(text: string, warnings: string[] = []): Promise<Workbook> {
  const doc = parseXml(text, 'flat ODS');
  return readOdf({ content: doc, styles: doc, settings: doc, meta: doc }, warnings);
}

export async function readOdf(docs: OdfDocs, warnings: string[]): Promise<Workbook> {
  const book = new Workbook(false);
  const defs = collectStyles([docs.styles !== docs.content ? docs.styles : null, docs.content]);
  const body = desc(docs.content, 'spreadsheet')[0];
  if (!body) throw new Error('This OpenDocument file has no spreadsheet body');
  // Null date: 1899-12-30 is the 1900 system; 1904-01-01 the 1904 system.
  const nd = desc(body, 'null-date')[0];
  if (nd && /^1904-01-01/.test(a(nd, 'table', 'date-value') || '')) book.calc.date1904 = true;
  const dateSerial = (s: string) => parseIsoDate(s, book.calc.date1904);
  const validity = readValidations(body, odfToFormula);
  const iter = desc(body, 'iteration')[0];
  if (iter && a(iter, 'table', 'status') === 'enable') {
    book.calc.iterative = true;
    const st = a(iter, 'table', 'steps'); if (st) book.calc.maxIterations = +st;
    const mc = a(iter, 'table', 'minimum-difference'); if (mc) book.calc.maxChange = +mc;
  }

  const tables = kids(body, 'table');
  for (const t of tables) {
    let name = a(t, 'table', 'name') || `Sheet${book.sheets.length + 1}`;
    while (book.sheetByName(name)) name += '_';
    const sheet = new Sheet(book, name);
    book.sheets.push(sheet);
    const ts = defs.table.get(a(t, 'table', 'style-name') || '');
    const tp = ts && kid(ts, 'table-properties');
    if (tp) {
      if (a(tp, 'table', 'display') === 'false') sheet.visibility = 'hidden';
      const tc = normColor(a(tp, 'tableooo', 'tab-color') || a(tp, 'table', 'tab-color'));
      if (tc) sheet.tabColor = tc;
      if (/^rl/.test(a(tp, 'style', 'writing-mode') || '')) sheet.rtl = true;
    }
    if (a(t, 'table', 'protected') === 'true') sheet.protection = { on: true, allow: ['selectLocked', 'selectUnlocked'] };
    const pr = a(t, 'table', 'print-ranges');
    if (pr) sheet.print.area = splitOdfAddresses(pr).map((x) => splitSheetRef(odfRefToA1(x)).ref).join(',');
    await readTable(book, sheet, t, defs, docs, dateSerial, warnings);
    readTableExtras(sheet, t, defs, validity, warnings);
    // Sheet-local names
    for (const ne of kids(kid(t, 'named-expressions'))) readName(book, ne, sheet.name);
  }
  if (!book.sheets.length) book.sheets.push(new Sheet(book, 'Sheet1'));
  for (const ne of kids(kid(body, 'named-expressions'))) readName(book, ne, null);
  // Database ranges with filter buttons -> autofilter
  for (const dr of desc(body, 'database-range')) {
    if (a(dr, 'table', 'display-filter-buttons') !== 'true') continue;
    const ref = odfRefToA1(a(dr, 'table', 'target-range-address') || '');
    const { sheet: sn, ref: body2 } = splitSheetRef(ref);
    const sh = sn ? book.sheetByName(sn) : book.sheets[0];
    const rg = parseRange(body2);
    if (sh && rg) sh.filter = { range: rg, columns: {} };
  }
  readSettings(book, docs.settings);
  readMeta(book, docs.meta);
  book.clearHistory();
  return book;
}

function readName(book: Workbook, ne: Element, sheet: string | null): void {
  const name = a(ne, 'table', 'name');
  if (!name) return;
  if (/^_xlnm\./i.test(name)) {
    const sh = sheet ? book.sheetByName(sheet) : null;
    const addr = splitOdfAddresses(a(ne, 'table', 'cell-range-address') || '').map(odfRefToA1);
    if (sh && /print_area/i.test(name)) sh.print.area = addr.map((x) => splitSheetRef(x).ref).join(',');
    return;
  }
  let ref = '';
  if (ne.localName === 'named-range') {
    const addr = a(ne, 'table', 'cell-range-address') || '';
    ref = splitOdfAddresses(addr).map(odfRefToA1).join(',');
  } else if (ne.localName === 'named-expression') ref = odfToFormula(a(ne, 'table', 'expression') || '');
  if (ref) book.names.push({ name, ref, sheet });
}

async function readTable(book: Workbook, sheet: Sheet, t: Element, defs: StyleDefs, docs: OdfDocs, dateSerial: (s: string) => number | null, warnings: string[]): Promise<void> {
  // Columns: collect runs first, then make the most common width the sheet default.
  const runs: { start: number; n: number; w: number | null; hidden: boolean; level: number; dsid: number; brk: boolean }[] = [];
  let col = 0;
  const colDefault: number[] = [];
  const walkCols = (el: Element, level: number) => {
    for (const ch of kids(el)) {
      if (ch.localName === 'table-column') {
        const n = Math.min(+(a(ch, 'table', 'number-columns-repeated') || 1), MAX_COLS - col);
        const cp = kid(defs.col.get(a(ch, 'table', 'style-name') || '') || null, 'table-column-properties');
        const w = lengthToPx(a(cp, 'style', 'column-width'));
        runs.push({
          start: col, n, w: w != null ? Math.round(w) : null, level,
          hidden: a(ch, 'table', 'visibility') === 'collapse' || a(ch, 'table', 'visibility') === 'filter',
          dsid: styleIdFor(book, defs, a(ch, 'table', 'default-cell-style-name')),
          brk: a(cp, 'fo', 'break-before') === 'page',
        });
        col += n;
      } else if (ch.localName === 'table-column-group') walkCols(ch, level + 1);
      else if (ch.localName === 'table-header-columns' || ch.localName === 'table-columns') walkCols(ch, level);
    }
  };
  walkCols(t, 0);
  const wCount = new Map<number, number>();
  for (const rn of runs) if (rn.w != null) wCount.set(rn.w, (wCount.get(rn.w) || 0) + rn.n);
  let modeW = 0, modeN = 0;
  for (const [w, n] of wCount) if (n > modeN) { modeW = w; modeN = n; }
  if (modeW > 0) sheet.defaultColW = modeW;
  for (const rn of runs) {
    const lim = Math.min(rn.n, 1024);
    if (rn.brk && rn.start > 0) sheet.colBreaks.push(rn.start);
    for (let i = 0; i < lim; i++) {
      const c = rn.start + i;
      if (rn.w != null && rn.w !== sheet.defaultColW) sheet.colW.set(c, rn.w);
      if (rn.hidden) sheet.hiddenCols.add(c);
      if (rn.level) sheet.colLevel.set(c, rn.level);
      if (rn.dsid) { colDefault[c] = rn.dsid; sheet.colStyle.set(c, rn.dsid); }
    }
  }

  // Rows
  let row = 0;
  let replicated = 0;
  const walkRows = async (el: Element, level: number) => {
    for (const ch of kids(el)) {
      if (row >= MAX_ROWS) return;
      if (ch.localName === 'table-row') {
        const rep = Math.max(1, +(a(ch, 'table', 'number-rows-repeated') || 1));
        const rs = defs.row.get(a(ch, 'table', 'style-name') || '');
        const rp = kid(rs, 'table-row-properties');
        const h = lengthToPx(a(rp, 'style', 'row-height'));
        const custom = a(rp, 'style', 'use-optimal-row-height') === 'false';
        const brk = a(rp, 'fo', 'break-before') === 'page';
        const hidden = a(ch, 'table', 'visibility') === 'collapse' || a(ch, 'table', 'visibility') === 'filter';
        const rowDefault = styleIdFor(book, defs, a(ch, 'table', 'default-cell-style-name'));
        const cells = kids(ch).filter((x) => x.localName === 'table-cell' || x.localName === 'covered-table-cell');
        const hasContent = cells.some((x) => x.childNodes.length > 0 || a(x, 'office', 'value-type') || a(x, 'table', 'formula'));
        // A row without values can still carry merged cells, cell formats and validation.
        const hasLayout = !hasContent && cells.some((x) => a(x, 'table', 'number-columns-spanned') || a(x, 'table', 'number-rows-spanned') || a(x, 'table', 'style-name') || a(x, 'table', 'content-validation-name'));
        const times = hasContent ? Math.min(rep, Math.max(1, 10000 - replicated)) : rep > 1000 ? 0 : rep;
        for (let k = 0; k < times && row + k < MAX_ROWS; k++) {
          const r = row + k;
          if (h && custom) sheet.rowH.set(r, Math.round(h));
          if (hidden) sheet.hiddenRows.add(r);
          if (level) sheet.rowLevel.set(r, level);
          if (brk && r > 0 && k === 0) sheet.rowBreaks.push(r);
          if (hasContent || rowDefault || hasLayout) await readRow(book, sheet, r, cells, defs, docs, dateSerial, colDefault, rowDefault, warnings);
        }
        if (hasContent && rep > 1) replicated += times;
        row += rep;
      } else if (ch.localName === 'table-row-group') await walkRows(ch, level + 1);
      else if (ch.localName === 'table-header-rows' || ch.localName === 'table-rows') await walkRows(ch, level);
    }
  };
  await walkRows(t, 0);

  // Shapes anchored to the page
  for (const sh of kids(kid(t, 'shapes'))) await readFrame(sheet, sh, null, docs, warnings);
}

async function readRow(book: Workbook, sheet: Sheet, r: number, cells: Element[], defs: StyleDefs, docs: OdfDocs, dateSerial: (s: string) => number | null, colDefault: number[], rowDefault: number, warnings: string[]): Promise<void> {
  let c = 0;
  for (const el of cells) {
    if (c >= MAX_COLS) break;
    const rep = Math.max(1, +(a(el, 'table', 'number-columns-repeated') || 1));
    if (el.localName === 'covered-table-cell') {
      // Covered cells can still hold annotations or content; keep values only.
      c += rep;
      continue;
    }
    const styleName = a(el, 'table', 'style-name');
    const sid = styleName ? styleIdFor(book, defs, styleName) : rowDefault || 0;
    const valName = a(el, 'table', 'content-validation-name');
    if (valName) noteValidation(sheet, valName, r, c, rep);
    if (styleName && kids(defs.cell.get(styleName) || null, 'map').length) noteStyleMap(sheet, styleName, r, c, rep);
    const vt = a(el, 'office', 'value-type') || a(el, 'calcext', 'value-type');
    const formula = a(el, 'table', 'formula');
    const paras = kids(el, 'p');
    const text = paras.map(odfText).join('\n');
    let v: Cell['v'] = null;
    switch (vt) {
      case 'float': case 'percentage': case 'currency': { const n = Number(a(el, 'office', 'value')); v = Number.isFinite(n) ? n : null; break; }
      case 'date': v = dateSerial(a(el, 'office', 'date-value') || ''); break;
      case 'time': v = parseIsoDuration(a(el, 'office', 'time-value') || ''); break;
      case 'boolean': v = a(el, 'office', 'boolean-value') === 'true'; break;
      case 'string': v = a(el, 'office', 'string-value') ?? text; break;
      case 'error': v = errorFromCode(text) || new CellError('#VALUE!'); break;
      default: v = text || null;
    }
    const cell: Cell = {};
    if (v != null && v !== '') cell.v = v;
    if (formula && !(vt === 'boolean' && /^(?:of:)?=?(TRUE|FALSE)\(\)$/i.test(formula))) {
      cell.f = odfToFormula(formula);
      const isErr = a(el, 'calcext', 'value-type') === 'error' || (vt === 'string' && /^(#[A-Z/0!?]+|Err:\d+)$/.test(text));
      if (isErr) cell.v = errorFromCode(text.trim()) || new CellError('#VALUE!');
    }
    let effSid = sid || colDefault[c] || 0;
    // Dates, times and percentages without a data style still display as such.
    if ((vt === 'time' || vt === 'date' || vt === 'percentage') && cell.v != null && !book.styles[effSid]?.numFmt) {
      const fmt = vt === 'time' ? (typeof cell.v === 'number' && cell.v >= 1 ? '[h]:mm:ss' : 'h:mm:ss') : vt === 'date' ? (typeof cell.v === 'number' && cell.v % 1 ? 'yyyy-mm-dd hh:mm:ss' : 'yyyy-mm-dd') : '0%';
      effSid = book.derivedStyle(effSid, { numFmt: fmt });
      if (effSid !== colDefault[c]) cell.s = effSid;
    }
    if (effSid && effSid !== colDefault[c]) cell.s = effSid;
    else if (effSid && (cell.v != null || cell.f)) cell.s = effSid;
    const link = desc(el, 'a').find((x) => x.localName === 'a' && a(x, 'xlink', 'href'));
    if (link) cell.link = a(link, 'xlink', 'href')!;
    const cs = +(a(el, 'table', 'number-columns-spanned') || 1), rs = +(a(el, 'table', 'number-rows-spanned') || 1);
    if (cs > 1 || rs > 1) sheet.merges.push({ r1: r, c1: c, r2: Math.min(MAX_ROWS - 1, r + rs - 1), c2: Math.min(MAX_COLS - 1, c + cs - 1) });
    const mc = +(a(el, 'table', 'number-matrix-columns-spanned') || 0), mr = +(a(el, 'table', 'number-matrix-rows-spanned') || 0);
    if (formula && (mc > 1 || mr > 1)) cell.arr = { r1: r, c1: c, r2: r + Math.max(1, mr) - 1, c2: c + Math.max(1, mc) - 1 };
    const ann = kid(el, 'annotation');
    const write = cell.v != null || cell.f != null || cell.link || (cell.s && rep <= 100);
    const count = write ? (cell.v != null || cell.f != null ? Math.min(rep, MAX_COLS - c) : Math.min(rep, 100)) : 0;
    for (let k = 0; k < count; k++) {
      const key2 = cellKey(r, c + k);
      const prev = sheet.cells.get(key2);
      const next: Cell = k === 0 ? cell : { ...cell };
      if (prev?.sp != null && next.f == null) next.sp = prev.sp;
      sheet.cells.set(key2, next);
    }
    if (cell.arr) {
      // Cells inside a matrix are results of the anchor
      const key = cellKey(r, c);
      for (let rr = cell.arr.r1; rr <= cell.arr.r2; rr++) for (let cc = cell.arr.c1; cc <= cell.arr.c2; cc++) {
        if (rr === r && cc === c) continue;
        const k2 = cellKey(rr, cc);
        const x = sheet.cells.get(k2);
        if (x && !x.f) x.sp = key;
        else if (!x) sheet.cells.set(k2, { sp: key });
      }
    }
    if (ann) {
      const author = desc(ann, 'creator')[0]?.textContent || '';
      const date = Date.parse(desc(ann, 'date')[0]?.textContent || '') || 0;
      const body = kids(ann, 'p').map(odfText).join('\n');
      sheet.comments.set(cellKey(r, c), { author, text: body, time: date, note: true });
    }
    for (const fr of kids(el).filter((x) => x.localName === 'frame' || x.localName === 'custom-shape' || x.localName === 'rect' || x.localName === 'ellipse')) {
      await readFrame(sheet, fr, { r, c }, docs, warnings);
    }
    c += rep;
  }
}

async function readFrame(sheet: Sheet, fr: Element, at: { r: number; c: number } | null, docs: OdfDocs, warnings: string[]): Promise<void> {
  const w = Math.round(lengthToPx(a(fr, 'svg', 'width')) || 100), h = Math.round(lengthToPx(a(fr, 'svg', 'height')) || 100);
  const x = lengthToPx(a(fr, 'svg', 'x')) || 0, y = lengthToPx(a(fr, 'svg', 'y')) || 0;
  let pos: DrawingObject['at'];
  if (at) pos = { r: at.r, c: at.c, dx: Math.round(x), dy: Math.round(y), w, h };
  else {
    let c = 0, r = 0, xx = x, yy = y;
    while (xx >= sheet.colWidth(c) && c < 5000) { xx -= sheet.colWidth(c); c++; }
    while (yy >= sheet.rowHeight(r) && r < 100000) { yy -= sheet.rowHeight(r); r++; }
    pos = { r, c, dx: Math.round(xx), dy: Math.round(yy), w, h };
  }
  const obj = kid(fr, 'object');
  if (obj || kid(fr, 'object-ole')) {
    let chart: Partial<import('../model/types').ChartDef> | null = null;
    try {
      const inline = obj && kid(obj, 'document');
      const href = (obj && a(obj, 'xlink', 'href')) || '';
      let cdoc: Document | Element | null = inline;
      if (!cdoc && docs.files && href) {
        const f = docs.files.file(href.replace(/^\.\//, '').replace(/\/$/, '') + '/content.xml');
        if (f) cdoc = parseXml(decodeText(await f.async('uint8array')), href);
      }
      if (cdoc) chart = readChart(cdoc, chartRef);
    } catch { chart = null; }
    if (chart) { sheet.charts.push({ id: newId('ch'), at: pos, ...chart } as import('../model/types').ChartDef); return; }
    if (!warnings.includes('Some embedded objects in OpenDocument files are not opened.')) warnings.push('Some embedded objects in OpenDocument files are not opened.');
    return;
  }
  const img = kid(fr, 'image');
  if (img) {
    const href = a(img, 'xlink', 'href') || '';
    let src = '';
    if (docs.files && href && !/^[a-z]+:/i.test(href)) {
      const f = docs.files.file(href.replace(/^\.\//, ''));
      if (f) src = `data:${mimeForExt(href.slice(href.lastIndexOf('.') + 1))};base64,${bytesToBase64(await f.async('uint8array'))}`;
    } else {
      const bin = kid(img, 'binary-data');
      if (bin) src = `data:image/png;base64,${(bin.textContent || '').replace(/\s+/g, '')}`;
    }
    if (src) sheet.drawings.push({ id: newId('img'), kind: 'image', at: pos, src, ...(a(fr, 'draw', 'name') ? { alt: a(fr, 'draw', 'name')! } : {}) });
    return;
  }
  const tb = kid(fr, 'text-box');
  if (tb || fr.localName !== 'frame') {
    const text = desc(fr, 'p').map(odfText).join('\n');
    sheet.drawings.push({ id: newId('shp'), kind: tb ? 'textbox' : 'shape', at: pos, shape: fr.localName === 'ellipse' ? 'ellipse' : 'rect', ...(text ? { text } : {}) });
    return;
  }
}

/** First range of an ODF chart range address list ("Data.B2:Data.B4 Data.B1") as an A1 reference. */
function chartRef(addr: string): string | null {
  const first = addr.trim().split(/\s+/)[0];
  if (!first) return null;
  const r = odfRefToA1(first);
  return /#REF!/.test(r) ? null : r;
}

/** Validity rules, conditional formats and style:map conditions of one table. */
function readTableExtras(sheet: Sheet, t: Element, defs: StyleDefs, validity: ReturnType<typeof readValidations>, warnings: string[]): void {
  applyValidations(sheet, validity);
  const addr = (s: string): Range[] => splitOdfAddresses(s).map((x) => parseRange(splitSheetRef(odfRefToA1(x)).ref)).filter((x): x is Range => !!x);
  const base = baseStyle(defs);
  const styleOf = (name: string): Partial<CellStyle> => {
    const st: Partial<CellStyle> = cellStyleOf(defs, name);
    for (const k of ['font', 'size', 'color', 'numFmt'] as const) if (st[k] != null && st[k] === base[k]) delete st[k];
    return st;
  };
  const conv = (f: string) => (/\[/.test(f) ? odfToFormula(f) : f.replace(/;/g, ','));
  const cf = readConditionalFormats(t, addr, styleOf, conv, warnings);
  if (cf.length) sheet.cf.push(...cf);
  else {
    const maps = (style: string) => kids(defs.cell.get(style) || null, 'map').map((m) => ({ condition: a(m, 'style', 'condition') || '', apply: a(m, 'style', 'apply-style-name') || '' }));
    sheet.cf.push(...styleMapRules(sheet, maps, styleOf, conv));
  }
}

function readSettings(book: Workbook, doc: Document | null): void {
  if (!doc) return;
  const view = desc(doc, 'config-item-set').find((s) => a(s, 'config', 'name') === 'ooo:view-settings');
  if (!view) return;
  const item = (el: Element, name: string) => kids(el, 'config-item').find((i) => a(i, 'config', 'name') === name)?.textContent ?? null;
  const entry = desc(view, 'config-item-map-entry')[0];
  if (!entry) return;
  const active = item(entry, 'ActiveTable');
  const tablesMap = kids(entry, 'config-item-map-named').find((m) => a(m, 'config', 'name') === 'Tables');
  for (const te of kids(tablesMap, 'config-item-map-entry')) {
    const sh = book.sheetByName(a(te, 'config', 'name') || '');
    if (!sh) continue;
    const hm = item(te, 'HorizontalSplitMode'), vm = item(te, 'VerticalSplitMode');
    const hp = +(item(te, 'HorizontalSplitPosition') || 0), vp = +(item(te, 'VerticalSplitPosition') || 0);
    if (hm === '2' || vm === '2') sh.freeze = { rows: vm === '2' ? vp : 0, cols: hm === '2' ? hp : 0 };
    const z = +(item(te, 'ZoomValue') || 0);
    if (z >= 10 && z <= 400) sh.zoom = z / 100;
    const sg = item(te, 'ShowGrid');
    if (sg === 'false') sh.showGrid = false;
    const cx = +(item(te, 'CursorPositionX') || 0), cy = +(item(te, 'CursorPositionY') || 0);
    if (cx || cy) sh.sel = { r: cy, c: cx, ranges: [{ r1: cy, c1: cx, r2: cy, c2: cx }] };
  }
  const glob = item(entry, 'ShowGrid');
  if (glob === 'false') for (const s of book.sheets) s.showGrid = false;
  if (active) { const i = book.sheets.findIndex((s) => s.name === active); if (i >= 0) book.activeSheet = i; }
}

function readMeta(book: Workbook, doc: Document | null): void {
  if (!doc) return;
  const meta = desc(doc, 'meta').find((m) => m.namespaceURI === NS.office || m.prefix === 'office');
  if (!meta) return;
  const t = (n: string) => (kids(meta, n)[0]?.textContent || '').trim();
  const p = book.props;
  p.title = t('title'); p.subject = t('subject'); p.comments = t('description');
  p.author = t('initial-creator') || t('creator'); p.lastModifiedBy = t('creator');
  p.keywords = kids(meta, 'keyword').map((k) => k.textContent || '').join(', ');
  const c = Date.parse(t('creation-date')), m = Date.parse(t('date'));
  if (Number.isFinite(c)) p.created = c;
  if (Number.isFinite(m)) p.modified = m;
  for (const u of kids(meta, 'user-defined')) { const n = a(u, 'meta', 'name'); if (n) p.custom[n] = u.textContent || ''; }
}

// --------------------------------------------------------------------------------------------
// Export

const ROOT_NS = Object.entries(NS).filter(([k]) => k !== 'manifest').map(([k, v]) => `xmlns:${k}="${v}"`).join(' ');

function borderOdf(b: Border): string {
  const w = b.style === 'thick' ? '2.5pt' : b.style === 'medium' || b.style === 'mediumDashed' ? '1.75pt' : b.style === 'double' ? '2.01pt' : b.style === 'hair' ? '0.05pt' : '0.75pt';
  const st = b.style === 'double' ? 'double' : /dash/i.test(b.style) ? 'dashed' : b.style === 'dotted' || b.style === 'hair' ? 'dotted' : 'solid';
  return `${w} ${st} ${b.color || '#000000'}`;
}

/** Spreadsheet format code -> ODF data style XML (best effort, first section). */
function dataStyleXml(name: string, code: string): string {
  const sec = code.split(/;(?=(?:[^"]*"[^"]*")*[^"]*$)/)[0];
  const nsAttr = (n: string) => `style:name="${name}"`;
  if (sec === '@') return `<number:text-style ${nsAttr(name)}><number:text-content/></number:text-style>`;
  if (isDateFormat(code)) {
    const parts: string[] = [];
    const s = sec.replace(/\[(?!h\]|hh\]|m\]|mm\]|s\]|ss\])[^\]]*\]/gi, '');
    const re = /(\[h+\]|\[m+\]|\[s+\]|y+|m+|d+|h+|s+(?:\.0+)?|AM\/PM|A\/P|"[^"]*"|\\.|.)/gi;
    const toks = s.match(re) || [];
    let lastWasHour = false;
    const isMinute = (i: number) => {
      if (lastWasHour) return true;
      for (let j = i + 1; j < toks.length; j++) { if (/^s/i.test(toks[j])) return true; if (/^[ymdh]/i.test(toks[j])) return false; }
      return false;
    };
    let hasDate = false, elapsed = false;
    toks.forEach((tk, i) => {
      const t = tk.toLowerCase();
      if (/^\[h+\]$/.test(t)) { elapsed = true; parts.push('<number:hours/>'); lastWasHour = true; return; }
      if (/^y+$/.test(t)) { hasDate = true; parts.push(`<number:year${t.length > 2 ? ' number:style="long"' : ''}/>`); lastWasHour = false; return; }
      if (/^m+$/.test(t)) {
        if (isMinute(i)) parts.push(`<number:minutes${t.length > 1 ? ' number:style="long"' : ''}/>`);
        else { hasDate = true; parts.push(t.length >= 3 ? `<number:month number:textual="true"${t.length > 3 ? ' number:style="long"' : ''}/>` : `<number:month${t.length > 1 ? ' number:style="long"' : ''}/>`); }
        lastWasHour = false; return;
      }
      if (/^d+$/.test(t)) { hasDate = true; parts.push(t.length >= 3 ? `<number:day-of-week${t.length > 3 ? ' number:style="long"' : ''}/>` : `<number:day${t.length > 1 ? ' number:style="long"' : ''}/>`); lastWasHour = false; return; }
      if (/^h+$/.test(t)) { parts.push(`<number:hours${t.length > 1 ? ' number:style="long"' : ''}/>`); lastWasHour = true; return; }
      if (/^s+/.test(t)) { const dp = (t.split('.')[1] || '').length; parts.push(`<number:seconds number:style="long"${dp ? ` number:decimal-places="${dp}"` : ''}/>`); lastWasHour = false; return; }
      if (t === 'am/pm' || t === 'a/p') { parts.push('<number:am-pm/>'); return; }
      const lit = tk.startsWith('"') ? tk.slice(1, -1) : tk.startsWith('\\') ? tk.slice(1) : tk;
      if (lit) {
        const prev = parts[parts.length - 1];
        if (prev && prev.startsWith('<number:text>')) parts[parts.length - 1] = `<number:text>${prev.slice(13, -14)}${esc(lit)}</number:text>`;
        else parts.push(`<number:text>${esc(lit)}</number:text>`);
      }
    });
    const tag = hasDate ? 'date-style' : 'time-style';
    return `<number:${tag} ${nsAttr(name)}${elapsed ? ' number:truncate-on-overflow="false"' : ''}>${parts.join('')}</number:${tag}>`;
  }
  const clean = sec.replace(/\[[^\]]*\]/g, (m) => (/^\[\$/.test(m) ? m : '')).replace(/_./g, '').replace(/\*./g, '');
  const prefix = /^(?:"([^"]*)"|\[\$([^\]-]*)[^\]]*\]|([$€£¥]))/.exec(clean.trim());
  const pre = prefix ? (prefix[1] ?? prefix[2] ?? prefix[3] ?? '') : '';
  const suffixM = /(?:"([^"]*)"|([$€£¥]))\s*$/.exec(clean.trim());
  const suf = suffixM && !prefix ? (suffixM[1] ?? suffixM[2] ?? '') : '';
  const numPart = clean.replace(/"[^"]*"|\[\$[^\]]*\]/g, '').replace(/[$€£¥]/g, '');
  const dp = (/\.([0#?]+)/.exec(numPart)?.[1] || '').length;
  const grouping = /#,##|0,0/.test(numPart);
  const minInt = /0[.,%E]|0$/i.test(numPart.replace(/,/g, '')) ? 1 : 0;
  const txt = (t: string) => (t ? `<number:text>${esc(t)}</number:text>` : '');
  if (/E[+-]/i.test(numPart)) {
    const ed = (/E[+-](0+)/i.exec(numPart)?.[1] || '00').length;
    return `<number:number-style ${nsAttr(name)}>${txt(pre)}<number:scientific-number number:decimal-places="${dp}" number:min-integer-digits="1" number:min-exponent-digits="${ed}"/>${txt(suf)}</number:number-style>`;
  }
  if (/\?\/[?\d]/.test(numPart)) {
    const m = /(\?+)\/(\?+|\d+)/.exec(numPart)!;
    const den = /^\d+$/.test(m[2]) ? ` number:denominator-value="${m[2]}"` : '';
    return `<number:number-style ${nsAttr(name)}><number:fraction number:min-integer-digits="0" number:min-numerator-digits="${m[1].length}" number:min-denominator-digits="${m[2].length}"${den}/></number:number-style>`;
  }
  const number = `<number:number number:decimal-places="${dp}" number:min-integer-digits="${minInt}"${grouping ? ' number:grouping="true"' : ''}/>`;
  if (isPercentFormat(code)) return `<number:percentage-style ${nsAttr(name)}>${txt(pre)}${number}<number:text>%</number:text></number:percentage-style>`;
  return `<number:number-style ${nsAttr(name)}>${txt(pre)}${number}${txt(suf)}</number:number-style>`;
}

interface OdsOut { files: Map<string, Uint8Array>; warnings: string[] }

function textP(s: string, link?: string): string {
  return s.split('\n').map((line) => {
    let body = esc(line).replace(/ {2,}/g, (m) => ` <text:s text:c="${m.length - 1}"/>`).replace(/\t/g, '<text:tab/>');
    if (body.startsWith(' ')) body = '<text:s/>' + body.slice(1);
    if (link) body = `<text:a xlink:href="${esc(link)}" xlink:type="simple">${body}</text:a>`;
    return `<text:p>${body}</text:p>`;
  }).join('');
}

function isoFromSerial(serial: number, date1904: boolean): string {
  const p = serialToParts(serial, date1904);
  const pad = (n: number, w = 2) => String(n).padStart(w, '0');
  const t = p.hh || p.mm || p.ss || p.ms ? `T${pad(p.hh)}:${pad(p.mm)}:${pad(p.ss)}${p.ms ? '.' + pad(p.ms, 3) : ''}` : '';
  return `${pad(p.y, 4)}-${pad(p.m)}-${pad(p.d)}${t}`;
}

function durationFromSerial(v: number): string {
  const total = Math.round(Math.abs(v) * 86400 * 1000) / 1000;
  const h = Math.floor(total / 3600), m = Math.floor((total % 3600) / 60), s = Math.round((total % 60) * 1000) / 1000;
  return `${v < 0 ? '-' : ''}PT${String(h).padStart(2, '0')}H${String(m).padStart(2, '0')}M${String(s).padStart(2, '0')}S`;
}

function buildOds(book: Workbook, out: OdsOut, flat: boolean): { content: string; styles: string; meta: string; settings: string } {
  const fonts = new Set<string>();
  const body = book.theme.bodyFont === 'Carlito' ? 'Calibri' : book.theme.bodyFont || 'Calibri';
  fonts.add(body);
  const dataStyles = new Map<string, string>();
  const dsXml: string[] = [];
  const cellStyleName = new Map<number, string>();
  const cellStyleXml: string[] = [];
  const dataStyleName = (code: string) => {
    let n = dataStyles.get(code);
    if (!n) { n = `N${dataStyles.size + 100}`; dataStyles.set(code, n); dsXml.push(dataStyleXml(n, code)); }
    return n;
  };
  const cellStyle = (sid: number): string | null => {
    if (!sid) return null;
    let n = cellStyleName.get(sid);
    if (n) return n;
    const st = book.styles[sid] || {};
    n = `ce${cellStyleName.size + 1}`;
    cellStyleName.set(sid, n);
    const tcp: string[] = [];
    if (st.fill) tcp.push(`fo:background-color="${st.fill}"`);
    if (st.bt) tcp.push(`fo:border-top="${borderOdf(st.bt)}"`);
    if (st.bb) tcp.push(`fo:border-bottom="${borderOdf(st.bb)}"`);
    if (st.bl) tcp.push(`fo:border-left="${borderOdf(st.bl)}"`);
    if (st.br) tcp.push(`fo:border-right="${borderOdf(st.br)}"`);
    if (st.diagUp) tcp.push(`style:diagonal-bl-tr="${borderOdf(st.diagUp)}"`);
    if (st.diagDown) tcp.push(`style:diagonal-tl-br="${borderOdf(st.diagDown)}"`);
    if (st.wrap) tcp.push('fo:wrap-option="wrap"');
    if (st.shrink) tcp.push('style:shrink-to-fit="true"');
    if (st.vAlign) tcp.push(`style:vertical-align="${st.vAlign === 'middle' || st.vAlign === 'distributed' || st.vAlign === 'justify' ? 'middle' : st.vAlign}"`);
    if (st.rotate && st.rotate !== 255) tcp.push(`style:rotation-angle="${st.rotate < 0 ? 360 + st.rotate : st.rotate}"`);
    if (st.locked === false || st.hidden) tcp.push(`style:cell-protect="${st.locked === false ? (st.hidden ? 'formula-hidden' : 'none') : 'hidden-and-protected'}"`);
    const ha = st.hAlign && st.hAlign !== 'general' ? (st.hAlign === 'center' || st.hAlign === 'centerContinuous' ? 'center' : st.hAlign === 'right' ? 'end' : st.hAlign === 'justify' || st.hAlign === 'distributed' ? 'justify' : 'start') : null;
    if (ha) tcp.push('style:text-align-source="fix"', 'style:repeat-content="false"');
    const pp: string[] = [];
    if (ha) pp.push(`fo:text-align="${ha}"`);
    if (st.indent) pp.push(`fo:margin-left="${pxToIn(st.indent * 9)}"`);
    const tp: string[] = [];
    if (st.bold) tp.push('fo:font-weight="bold" style:font-weight-asian="bold" style:font-weight-complex="bold"');
    if (st.italic) tp.push('fo:font-style="italic" style:font-style-asian="italic" style:font-style-complex="italic"');
    if (st.color) tp.push(`fo:color="${st.color}"`);
    if (st.size) tp.push(`fo:font-size="${st.size}pt" style:font-size-asian="${st.size}pt" style:font-size-complex="${st.size}pt"`);
    if (st.font) { fonts.add(st.font); tp.push(`style:font-name="${esc(st.font)}"`); }
    if (st.underline) tp.push(`style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"${st.underline.startsWith('double') ? ' style:text-underline-type="double"' : ''}`);
    if (st.strike) tp.push('style:text-line-through-style="solid" style:text-line-through-type="single"');
    if (st.vertAlign) tp.push(`style:text-position="${st.vertAlign === 'superscript' ? 'super' : 'sub'} 58%"`);
    const ds = st.numFmt && st.numFmt !== 'General' ? ` style:data-style-name="${dataStyleName(st.numFmt)}"` : '';
    cellStyleXml.push(`<style:style style:name="${n}" style:family="table-cell" style:parent-style-name="Default"${ds}>` +
      (tcp.length ? `<style:table-cell-properties ${tcp.join(' ')}/>` : '') + (pp.length ? `<style:paragraph-properties ${pp.join(' ')}/>` : '') + (tp.length ? `<style:text-properties ${tp.join(' ')}/>` : '') + '</style:style>');
    return n;
  };

  const colStyles = new Map<string, string>();
  const rowStyles = new Map<string, string>();
  const tableStyles = new Map<string, string>();
  const autoXml: string[] = [];
  const colStyle = (w: number, brk: boolean) => {
    const k = `${w}|${brk}`;
    let n = colStyles.get(k);
    if (!n) { n = `co${colStyles.size + 1}`; colStyles.set(k, n); autoXml.push(`<style:style style:name="${n}" style:family="table-column"><style:table-column-properties fo:break-before="${brk ? 'page' : 'auto'}" style:column-width="${pxToIn(w)}"/></style:style>`); }
    return n;
  };
  const rowStyle = (h: number, custom: boolean, brk: boolean) => {
    const k = `${h}|${custom}|${brk}`;
    let n = rowStyles.get(k);
    if (!n) { n = `ro${rowStyles.size + 1}`; rowStyles.set(k, n); autoXml.push(`<style:style style:name="${n}" style:family="table-row"><style:table-row-properties style:row-height="${pxToIn(h)}" fo:break-before="${brk ? 'page' : 'auto'}" style:use-optimal-row-height="${custom ? 'false' : 'true'}"/></style:style>`); }
    return n;
  };
  const tableStyle = (s: Sheet) => {
    const k = `${s.visibility === 'visible'}|${s.tabColor || ''}|${s.rtl}`;
    let n = tableStyles.get(k);
    if (!n) {
      n = `ta${tableStyles.size + 1}`; tableStyles.set(k, n);
      autoXml.push(`<style:style style:name="${n}" style:family="table" style:master-page-name="Default"><style:table-properties table:display="${s.visibility === 'visible'}" style:writing-mode="${s.rtl ? 'rl-tb' : 'lr-tb'}"${s.tabColor ? ` tableooo:tab-color="${s.tabColor}"` : ''}/></style:style>`);
    }
    return n;
  };

  let imageN = 0;
  const tablesXml: string[] = [];
  const dbRanges: string[] = [];
  const visibleIdx = book.sheets.findIndex((s) => s.visibility === 'visible');
  book.sheets.forEach((sheet, si) => {
    const x: string[] = [];
    const vis = si === visibleIdx ? { ...sheet, visibility: 'visible' as const } as Sheet : sheet;
    const pa = sheet.print.area ? ` table:print-ranges="${esc(sheet.print.area.split(',').map((p) => { const rg = parseRange(splitSheetRef(p).ref); return rg ? `$${quoteOdfSheet(sheet.name)}.$${colName(rg.c1)}$${rg.r1 + 1}:.$${colName(rg.c2)}$${rg.r2 + 1}` : ''; }).filter(Boolean).join(' '))}"` : '';
    x.push(`<table:table table:name="${esc(sheet.name)}" table:style-name="${tableStyle(vis)}"${sheet.protection.on ? ' table:protected="true"' : ''}${pa}>`);
    // Extra per-cell content: drawings anchored at cells
    const frames = new Map<number, string[]>();
    for (const d of sheet.drawings) {
      const k = cellKey(d.at.r, d.at.c);
      let fx = '';
      const size = `svg:x="${pxToIn(d.at.dx || 0)}" svg:y="${pxToIn(d.at.dy || 0)}" svg:width="${pxToIn(d.at.w)}" svg:height="${pxToIn(d.at.h)}"`;
      if (d.src && d.src.startsWith('data:')) {
        const img = parseDataUrl(d.src);
        if (!img) continue;
        const name = `Pictures/image${++imageN}.${extForMime(img.mime)}`;
        if (flat) fx = `<draw:frame draw:z-index="${imageN}" draw:name="${esc(d.alt || `Image ${imageN}`)}" ${size}><draw:image><office:binary-data>${bytesToBase64(img.bytes)}</office:binary-data></draw:image></draw:frame>`;
        else { out.files.set(name, img.bytes); fx = `<draw:frame draw:z-index="${imageN}" draw:name="${esc(d.alt || `Image ${imageN}`)}" ${size}><draw:image xlink:href="${name}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame>`; }
      } else if (d.kind === 'textbox' || d.kind === 'shape') {
        fx = `<draw:frame draw:z-index="0" ${size}><draw:text-box>${textP(d.text || '')}</draw:text-box></draw:frame>`;
      }
      if (fx) { if (!frames.has(k)) frames.set(k, []); frames.get(k)!.push(fx); }
    }
    if (sheet.charts.length && !out.warnings.includes('Charts are not saved in OpenDocument files.')) out.warnings.push('Charts are not saved in OpenDocument files.');
    for (const [has, note] of [[sheet.validations.length > 0, 'Data validation is not saved in OpenDocument files.'], [sheet.cf.length > 0, 'Conditional formatting is not saved in OpenDocument files.']] as const) if (has && !out.warnings.includes(note)) out.warnings.push(note);

    // Extent
    const used = sheet.usedRange();
    let maxR = used ? used.r2 : -1, maxC = used ? used.c2 : -1;
    for (const m of sheet.merges) { maxR = Math.max(maxR, m.r2); maxC = Math.max(maxC, m.c2); }
    for (const k of frames.keys()) { maxR = Math.max(maxR, keyRow(k)); maxC = Math.max(maxC, keyCol(k)); }
    for (const k of sheet.comments.keys()) { maxR = Math.max(maxR, keyRow(k)); maxC = Math.max(maxC, keyCol(k)); }
    for (const r of [...sheet.rowH.keys(), ...sheet.hiddenRows]) maxR = Math.max(maxR, r);
    for (const c of [...sheet.colW.keys(), ...sheet.hiddenCols]) maxC = Math.max(maxC, c);
    maxR = Math.min(maxR, MAX_ROWS - 1); maxC = Math.min(maxC, MAX_COLS - 1);

    // Columns
    const colDesc = (c: number) => {
      const st = colStyle(sheet.colW.get(c) ?? sheet.defaultColW, sheet.colBreaks.includes(c));
      const ds = sheet.colStyle.get(c) ? cellStyle(sheet.colStyle.get(c)!) : null;
      return `table:style-name="${st}"${sheet.hiddenCols.has(c) ? ' table:visibility="collapse"' : ''} table:default-cell-style-name="${ds || 'Default'}"`;
    };
    const colGroup: string[] = [];
    let c = 0;
    const lastCol = Math.max(maxC, 0);
    while (c <= lastCol) {
      const d = colDesc(c);
      let n = 1;
      while (c + n <= lastCol && colDesc(c + n) === d && (sheet.colLevel.get(c + n) || 0) === (sheet.colLevel.get(c) || 0)) n++;
      colGroup.push(`<table:table-column ${d}${n > 1 ? ` table:number-columns-repeated="${n}"` : ''}/>`);
      c += n;
    }
    colGroup.push(`<table:table-column table:style-name="${colStyle(sheet.defaultColW, false)}" table:number-columns-repeated="${Math.max(1, Math.min(1024, MAX_COLS) - lastCol - 1)}" table:default-cell-style-name="Default"/>`);
    x.push(colGroup.join(''));

    // Rows
    const covered = new Map<number, true>();
    const spanAt = new Map<number, Range>();
    for (const m of sheet.merges) {
      spanAt.set(cellKey(m.r1, m.c1), m);
      for (let r = m.r1; r <= m.r2; r++) for (let cc = m.c1; cc <= m.c2; cc++) if (r !== m.r1 || cc !== m.c1) covered.set(cellKey(r, cc), true);
    }
    // Spill extents for dynamic arrays
    const spill = new Map<number, Range>();
    for (const [k, cell] of sheet.cells) {
      if (cell.sp == null || cell.f != null) continue;
      const g = spill.get(cell.sp);
      const r = keyRow(k), cc = keyCol(k);
      if (!g) spill.set(cell.sp, { r1: keyRow(cell.sp), c1: keyCol(cell.sp), r2: r, c2: cc });
      else { g.r2 = Math.max(g.r2, r); g.c2 = Math.max(g.c2, cc); }
    }
    // Row index -> sorted cell keys
    const byRow = new Map<number, number[]>();
    for (const k of sheet.cells.keys()) { const r = keyRow(k); if (!byRow.has(r)) byRow.set(r, []); byRow.get(r)!.push(k); }
    const rowsXml: string[] = [];
    const rowAttrs = (r: number) => {
      const h = sheet.rowH.get(r);
      const st = rowStyle(h ?? sheet.defaultRowH, h != null, sheet.rowBreaks.includes(r));
      const ds = sheet.rowStyle.get(r) ? cellStyle(sheet.rowStyle.get(r)!) : null;
      return `table:style-name="${st}"${sheet.hiddenRows.has(r) ? ' table:visibility="collapse"' : ''}${ds ? ` table:default-cell-style-name="${ds}"` : ''}`;
    };
    const mergeRows = new Set<number>();
    for (const m of sheet.merges) for (let rr = m.r1; rr <= Math.min(m.r2, m.r1 + 100000); rr++) mergeRows.add(rr);
    const extraKeys = [...frames.keys(), ...sheet.comments.keys()];
    const extraRows = new Set<number>(extraKeys.map(keyRow));
    const special = (rr: number) => byRow.has(rr) || mergeRows.has(rr) || extraRows.has(rr);
    let r = 0;
    while (r <= maxR) {
      const keys = byRow.get(r);
      if (!special(r)) {
        const ra = rowAttrs(r);
        let n = 1;
        while (r + n <= maxR && !special(r + n) && rowAttrs(r + n) === ra) n++;
        rowsXml.push(`<table:table-row ${ra}${n > 1 ? ` table:number-rows-repeated="${n}"` : ''}><table:table-cell table:number-columns-repeated="${Math.max(1, lastCol + 1)}"/></table:table-row>`);
        r += n;
        continue;
      }
      const cellsXml: string[] = [];
      const cols = new Set<number>((keys || []).map(keyCol));
      if (mergeRows.has(r)) for (const m of sheet.merges) if (r >= m.r1 && r <= m.r2) for (let cc = m.c1; cc <= m.c2; cc++) cols.add(cc);
      if (extraRows.has(r)) for (const k of extraKeys) if (keyRow(k) === r) cols.add(keyCol(k));
      const sorted = [...cols].sort((p, q) => p - q);
      let cur = 0;
      for (const cc of sorted) {
        if (cc > cur) cellsXml.push(`<table:table-cell${cc - cur > 1 ? ` table:number-columns-repeated="${cc - cur}"` : ''}/>`);
        cur = cc + 1;
        const k = cellKey(r, cc);
        const cell = sheet.cells.get(k);
        const comment = sheet.comments.get(k);
        const extra = (frames.get(k) || []).join('');
        const ann = comment ? `<office:annotation><dc:creator>${esc(comment.author)}</dc:creator>${comment.time ? `<dc:date>${new Date(comment.time).toISOString().replace(/Z$/, '')}</dc:date>` : ''}${textP([comment.text, ...(comment.replies || []).map((rp) => `${rp.author}: ${rp.text}`)].join('\n'))}</office:annotation>` : '';
        const tag = covered.has(k) ? 'table:covered-table-cell' : 'table:table-cell';
        const at: string[] = [];
        const sn = cell?.s ? cellStyle(cell.s) : null;
        if (sn) at.push(`table:style-name="${sn}"`);
        const m = spanAt.get(k);
        if (m && tag === 'table:table-cell') {
          if (m.c2 > m.c1) at.push(`table:number-columns-spanned="${m.c2 - m.c1 + 1}"`);
          if (m.r2 > m.r1) at.push(`table:number-rows-spanned="${m.r2 - m.r1 + 1}"`);
        }
        let content = '';
        const v = cell?.v;
        if (cell?.f != null && cell.sp == null) {
          at.push(`table:formula="${esc(formulaToOdf(cell.f, { book, sheet, r }))}"`);
          const area = cell.arr || spill.get(k);
          if (area) at.push(`table:number-matrix-columns-spanned="${area.c2 - area.c1 + 1}" table:number-matrix-rows-spanned="${area.r2 - area.r1 + 1}"`);
        }
        const fmt = cell?.s ? book.styles[cell.s]?.numFmt : undefined;
        if (typeof v === 'number' && Number.isFinite(v)) {
          if (fmt && isDateFormat(fmt)) {
            const timeOnly = !/[yd]/i.test(fmt.replace(/"[^"]*"/g, '')) && !/(^|[^h:])m{3,}/i.test(fmt);
            if (timeOnly) at.push(`office:value-type="time" office:time-value="${durationFromSerial(v)}"`);
            else at.push(`office:value-type="date" office:date-value="${isoFromSerial(v, book.calc.date1904)}"`);
          } else if (fmt && isPercentFormat(fmt)) at.push(`office:value-type="percentage" office:value="${v}"`);
          else at.push(`office:value-type="float" office:value="${v}"`);
          content = textP(String(v));
        } else if (typeof v === 'boolean') {
          at.push(`office:value-type="boolean" office:boolean-value="${v}"`);
          content = textP(v ? 'TRUE' : 'FALSE');
        } else if (v instanceof CellError) {
          if (cell?.f == null) at.push(`table:formula="of:=${esc(v.code === '#N/A' ? 'NA()' : v.code)}"`);
          at.push('office:value-type="string" office:string-value="" calcext:value-type="error"');
          content = textP(v.code);
        } else if (typeof v === 'string') {
          at.push('office:value-type="string"');
          content = textP(v, cell?.link);
        } else if (cell?.link) content = textP(cell.link, cell.link);
        if (cell?.link && typeof v !== 'string' && content && !content.includes('text:a')) content = textP(String(v ?? cell.link), cell.link);
        const inner = ann + extra + content;
        cellsXml.push(inner ? `<${tag}${at.length ? ' ' + at.join(' ') : ''}>${inner}</${tag}>` : `<${tag}${at.length ? ' ' + at.join(' ') : ''}/>`);
      }
      if (cur <= lastCol) cellsXml.push(`<table:table-cell${lastCol + 1 - cur > 1 ? ` table:number-columns-repeated="${lastCol + 1 - cur}"` : ''}/>`);
      rowsXml.push(`<table:table-row ${rowAttrs(r)}>${cellsXml.join('')}</table:table-row>`);
      r++;
    }
    if (!rowsXml.length) rowsXml.push(`<table:table-row table:style-name="${rowStyle(sheet.defaultRowH, false, false)}"><table:table-cell/></table:table-row>`);
    // Outline groups for rows
    x.push(groupRows(rowsXml, sheet, maxR));
    // Sheet-local names
    const local = book.names.filter((n) => n.sheet === sheet.name);
    if (local.length) x.push(`<table:named-expressions>${local.map((n) => nameXml(n.name, n.ref, sheet.name)).join('')}</table:named-expressions>`);
    x.push('</table:table>');
    tablesXml.push(x.join(''));
    if (sheet.filter) {
      const f = sheet.filter.range;
      dbRanges.push(`<table:database-range table:name="__Anonymous_Sheet_DB__${si}" table:target-range-address="${esc(`$${quoteOdfSheet(sheet.name)}.$${colName(f.c1)}$${f.r1 + 1}:.$${colName(f.c2)}$${f.r2 + 1}`)}" table:display-filter-buttons="true"/>`);
    }
  });
  const globalNames = book.names.filter((n) => n.sheet == null);
  const named = globalNames.length ? `<table:named-expressions>${globalNames.map((n) => nameXml(n.name, n.ref, book.sheets[0]?.name || 'Sheet1')).join('')}</table:named-expressions>` : '';
  const calc = book.calc.date1904 || book.calc.iterative
    ? `<table:calculation-settings>${book.calc.date1904 ? '<table:null-date table:date-value="1904-01-01"/>' : ''}${book.calc.iterative ? `<table:iteration table:status="enable" table:steps="${book.calc.maxIterations}" table:minimum-difference="${book.calc.maxChange}"/>` : ''}</table:calculation-settings>`
    : '';
  const fontDecls = `<office:font-face-decls>${[...fonts].map((f) => `<style:font-face style:name="${esc(f)}" svg:font-family="${esc(/\s/.test(f) ? `'${f}'` : f)}"/>`).join('')}</office:font-face-decls>`;
  const automatic = `<office:automatic-styles>${dsXml.join('')}${autoXml.join('')}${cellStyleXml.join('')}</office:automatic-styles>`;
  const spreadsheet = `<office:body><office:spreadsheet>${calc}${tablesXml.join('')}${named}${dbRanges.length ? `<table:database-ranges>${dbRanges.join('')}</table:database-ranges>` : ''}</office:spreadsheet></office:body>`;
  const commonStyles = `<office:styles><style:default-style style:family="table-cell"><style:paragraph-properties style:tab-stop-distance="0.5in"/><style:text-properties style:font-name="${esc(body)}" fo:font-size="11pt" fo:language="en" fo:country="US"/></style:default-style><number:number-style style:name="N0"><number:number number:min-integer-digits="1"/></number:number-style><style:style style:name="Default" style:family="table-cell" style:data-style-name="N0"/></office:styles>`;
  const masters = `<office:automatic-styles><style:page-layout style:name="pm1"><style:page-layout-properties style:writing-mode="lr-tb"/></style:page-layout></office:automatic-styles><office:master-styles><style:master-page style:name="Default" style:page-layout-name="pm1"><style:header style:display="false"/><style:footer style:display="false"/></style:master-page></office:master-styles>`;
  const p = book.props;
  const metaInner = `<office:meta><meta:generator>Anne\'s Grid</meta:generator>${p.title ? `<dc:title>${esc(p.title)}</dc:title>` : ''}${p.subject ? `<dc:subject>${esc(p.subject)}</dc:subject>` : ''}${p.comments ? `<dc:description>${esc(p.comments)}</dc:description>` : ''}${p.author ? `<meta:initial-creator>${esc(p.author)}</meta:initial-creator>` : ''}${p.lastModifiedBy || p.author ? `<dc:creator>${esc(p.lastModifiedBy || p.author)}</dc:creator>` : ''}${p.keywords ? p.keywords.split(/\s*,\s*/).map((k) => `<meta:keyword>${esc(k)}</meta:keyword>`).join('') : ''}<meta:creation-date>${new Date(p.created || Date.now()).toISOString().replace(/Z$/, '')}</meta:creation-date><dc:date>${new Date(p.modified || Date.now()).toISOString().replace(/Z$/, '')}</dc:date>${Object.entries(p.custom || {}).map(([k, v]) => `<meta:user-defined meta:name="${esc(k)}">${esc(v)}</meta:user-defined>`).join('')}</office:meta>`;
  const settingsInner = settingsXml(book);
  if (flat) return { content: `${fontDecls}${commonStyles.replace('</office:styles>', '</office:styles>')}${automatic.replace('</office:automatic-styles>', '<style:page-layout style:name="pm1"><style:page-layout-properties style:writing-mode="lr-tb"/></style:page-layout></office:automatic-styles>')}<office:master-styles><style:master-page style:name="Default" style:page-layout-name="pm1"><style:header style:display="false"/><style:footer style:display="false"/></style:master-page></office:master-styles>${spreadsheet}`, styles: '', meta: metaInner, settings: settingsInner };
  return {
    content: XML_HEADER + `<office:document-content ${ROOT_NS} office:version="1.3">${fontDecls}${automatic}${spreadsheet}</office:document-content>`,
    styles: XML_HEADER + `<office:document-styles ${ROOT_NS} office:version="1.3">${fontDecls}${commonStyles}${masters}</office:document-styles>`,
    meta: XML_HEADER + `<office:document-meta ${ROOT_NS} office:version="1.3">${metaInner}</office:document-meta>`,
    settings: XML_HEADER + `<office:document-settings ${ROOT_NS} office:version="1.3">${settingsInner}</office:document-settings>`,
  };
}

function groupRows(rows: string[], sheet: Sheet, _maxR: number): string {
  if (!sheet.rowLevel.size) return rows.join('');
  // rows[] here is per emitted element, which may cover several rows; rebuild with indices.
  const out: string[] = [];
  let level = 0, r = 0;
  for (const x of rows) {
    const rep = +(/table:number-rows-repeated="(\d+)"/.exec(x)?.[1] || 1);
    const lvl = sheet.rowLevel.get(r) || 0;
    while (level < lvl) { out.push('<table:table-row-group>'); level++; }
    while (level > lvl) { out.push('</table:table-row-group>'); level--; }
    out.push(x);
    r += rep;
  }
  while (level > 0) { out.push('</table:table-row-group>'); level--; }
  return out.join('');
}

function nameXml(name: string, ref: string, baseSheet: string): string {
  const { sheet, ref: body } = splitSheetRef(ref.replace(/^=/, ''));
  const rg = parseRange(body);
  const base = esc(`$${quoteOdfSheet(sheet || baseSheet)}.$A$1`);
  if (rg && !/[(),]/.test(body)) {
    const sh = '$' + quoteOdfSheet(sheet || baseSheet);
    const abs = (r: number, c: number) => `$${colName(c)}$${r + 1}`;
    const addr = rg.r1 === rg.r2 && rg.c1 === rg.c2 ? `${sh}.${abs(rg.r1, rg.c1)}` : `${sh}.${abs(rg.r1, rg.c1)}:.${abs(rg.r2, rg.c2)}`;
    return `<table:named-range table:name="${esc(name)}" table:base-cell-address="${base}" table:cell-range-address="${esc(addr)}"/>`;
  }
  return `<table:named-expression table:name="${esc(name)}" table:base-cell-address="${base}" table:expression="${esc(formulaToOdf(ref.replace(/^=/, '')))}"/>`;
}

function settingsXml(book: Workbook): string {
  const it = (n: string, t: string, v: string | number | boolean) => `<config:config-item config:name="${n}" config:type="${t}">${v}</config:config-item>`;
  const tables = book.sheets.map((s) => {
    const f = s.freeze || { rows: 0, cols: 0 };
    return `<config:config-item-map-entry config:name="${esc(s.name)}">` +
      it('CursorPositionX', 'int', s.sel.c) + it('CursorPositionY', 'int', s.sel.r) +
      it('HorizontalSplitMode', 'short', f.cols ? 2 : 0) + it('VerticalSplitMode', 'short', f.rows ? 2 : 0) +
      it('HorizontalSplitPosition', 'int', f.cols) + it('VerticalSplitPosition', 'int', f.rows) +
      it('ActiveSplitRange', 'short', f.rows ? 2 : f.cols ? 3 : 2) + it('PositionLeft', 'int', 0) + it('PositionRight', 'int', f.cols) +
      it('PositionTop', 'int', 0) + it('PositionBottom', 'int', f.rows) + it('ZoomType', 'short', 0) + it('ZoomValue', 'int', Math.round((s.zoom || 1) * 100)) +
      it('PageViewZoomValue', 'int', 60) + it('ShowGrid', 'boolean', s.showGrid) + '</config:config-item-map-entry>';
  }).join('');
  const active = book.sheets[Math.min(book.activeSheet, book.sheets.length - 1)]?.name || '';
  return '<office:settings><config:config-item-set config:name="ooo:view-settings"><config:config-item-map-indexed config:name="Views"><config:config-item-map-entry>' +
    it('ViewId', 'string', 'view1') + `<config:config-item-map-named config:name="Tables">${tables}</config:config-item-map-named>` +
    it('ActiveTable', 'string', esc(active)) + it('ShowGrid', 'boolean', true) + it('HasColumnRowHeaders', 'boolean', true) + it('HasSheetTabs', 'boolean', true) +
    '</config:config-item-map-entry></config:config-item-map-indexed></config:config-item-set></office:settings>';
}

export async function exportOds(book: Workbook, warnings: string[] = []): Promise<Blob> {
  const out: OdsOut = { files: new Map(), warnings };
  const x = buildOds(book, out, false);
  const zip = new JSZip();
  zip.file('mimetype', 'application/vnd.oasis.opendocument.spreadsheet', { compression: 'STORE' });
  zip.file('content.xml', x.content);
  zip.file('styles.xml', x.styles);
  zip.file('meta.xml', x.meta);
  zip.file('settings.xml', x.settings);
  const entries = ['<manifest:file-entry manifest:full-path="/" manifest:version="1.3" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>'];
  for (const p of ['content.xml', 'styles.xml', 'meta.xml', 'settings.xml']) entries.push(`<manifest:file-entry manifest:full-path="${p}" manifest:media-type="text/xml"/>`);
  for (const [p, bytes] of out.files) {
    zip.file(p, bytes);
    entries.push(`<manifest:file-entry manifest:full-path="${p}" manifest:media-type="${mimeForExt(p.slice(p.lastIndexOf('.') + 1))}"/>`);
  }
  zip.file('META-INF/manifest.xml', XML_HEADER + `<manifest:manifest xmlns:manifest="${NS.manifest}" manifest:version="1.3">${entries.join('')}</manifest:manifest>`);
  const data = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', mimeType: 'application/vnd.oasis.opendocument.spreadsheet' });
  return new Blob([data as BlobPart], { type: 'application/vnd.oasis.opendocument.spreadsheet' });
}

export function exportFods(book: Workbook, warnings: string[] = []): Blob {
  const out: OdsOut = { files: new Map(), warnings };
  const x = buildOds(book, out, true);
  const xml = XML_HEADER + `<office:document ${ROOT_NS} office:version="1.3" office:mimetype="application/vnd.oasis.opendocument.spreadsheet">${x.meta}${x.settings}${x.content}</office:document>`;
  return new Blob([xml], { type: 'application/vnd.oasis.opendocument.spreadsheet-flat-xml' });
}

export { colIndex };
