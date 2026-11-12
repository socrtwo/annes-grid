/**
 * HTML table import/export and Markdown table export.
 */
import { Workbook, Sheet } from '../model/workbook';
import { CellError, type Cell, type CellStyle, type Border } from '../model/types';
import { cellKey, keyRow, keyCol, type Range } from '../model/address';
import { formatValue } from '../model/numfmt';
import { convertField } from './csv';
import { isScriptUrl, safeNum } from '../util/sanitize';
import { normColor } from './index';

// --------------------------------------------------------------------------------------------
// Import

export interface HtmlImportOptions { allTables?: boolean }

function cssBorder(v: string | undefined): Border | undefined {
  if (!v || /\bnone\b|^0(px|pt)?(\s|$)|hidden/.test(v.trim())) return undefined;
  const w = parseFloat(/([\d.]+)px/.exec(v)?.[1] || /([\d.]+)pt/.exec(v)?.[1] || '1');
  const style: Border['style'] = /double/.test(v) ? 'double' : /dashed/.test(v) ? 'dashed' : /dotted/.test(v) ? 'dotted' : w >= 3 ? 'thick' : w >= 2 ? 'medium' : 'thin';
  const color = normColor(/(#[0-9a-f]{3,6}|rgba?\([^)]*\)|\b[a-z]+\b)\s*$/i.exec(v.trim())?.[1] || '') || undefined;
  return { style, ...(color ? { color } : {}) };
}

/** Simple style sheet rules (tag, .class and tag.class selectors) from <style> elements. */
interface CssRule { tag: string | null; cls: string | null; text: string; order: number }

export function parseCssRules(doc: Document): CssRule[] {
  const rules: CssRule[] = [];
  let order = 0;
  for (const st of Array.from(doc.querySelectorAll('style'))) {
    const css = (st.textContent || '').replace(/\/\*[\s\S]*?\*\//g, '').replace(/<!--|-->/g, '');
    for (const m of css.matchAll(/([^{}@]+)\{([^{}]*)\}/g)) {
      for (const sel of m[1].split(',')) {
        const t = sel.trim();
        const mm = /^([a-zA-Z][\w-]*)?(?:\.([\w-]+))?$/.exec(t);
        if (!mm || (!mm[1] && !mm[2])) continue;
        rules.push({ tag: mm[1] ? mm[1].toLowerCase() : null, cls: mm[2] || null, text: m[2], order: order++ });
      }
    }
  }
  return rules;
}

/** CSS text that applies to an element: matching rules (tag, then class), then its style attribute. */
function cssFor(el: Element, rules: CssRule[]): string {
  const tag = el.tagName.toLowerCase();
  const classes = (el.getAttribute('class') || '').split(/\s+/).filter(Boolean);
  const hits = rules.filter((r) => (!r.tag || r.tag === tag) && (!r.cls || classes.includes(r.cls)) && !(r.tag && !['td', 'th', 'tr', 'table', 'col'].includes(r.tag) && !r.cls));
  hits.sort((x, y) => ((x.cls ? 1 : 0) - (y.cls ? 1 : 0)) || x.order - y.order);
  return hits.map((h) => h.text).join(';') + ';' + (el.getAttribute('style') || '');
}

/** The value of a (possibly non-standard) property in CSS text; the last declaration wins. */
function cssProp(text: string, name: string): string | undefined {
  let v: string | undefined;
  for (const m of text.matchAll(new RegExp(`(?:^|;)\\s*${name}\\s*:\\s*("(?:[^"\\\\]|\\\\.)*"|'[^']*'|[^;]*)`, 'gi'))) v = m[1].trim();
  return v;
}

/** Number format codes from HTML producers: Excel's mso-number-format, LibreOffice's sdnum, Google Sheets' JSON. */
function htmlNumFmt(code: string | undefined | null): string | undefined {
  if (!code) return undefined;
  let c = code.trim().replace(/^["']|["']$/g, '').replace(/\\(.)/g, '$1');
  if (!c || /^General$/i.test(c)) return undefined;
  if (c === '@' || /^\\?@$/.test(c)) return '@';
  // LibreOffice and some producers write date and time letters in upper case.
  if (/[YDHS]/.test(c) && !/"/.test(c)) c = c.replace(/[YMDHS]+/g, (t) => t.toLowerCase()).replace(/am\/pm/gi, 'AM/PM');
  return c;
}

function styleFromElement(el: HTMLElement, css: string): CellStyle {
  const st: CellStyle = {};
  const tmp = el.ownerDocument.createElement('td');
  tmp.style.cssText = css;
  const s = tmp.style;
  const fw = s.fontWeight;
  if (fw === 'bold' || fw === 'bolder' || +fw >= 600 || el.tagName === 'TH') st.bold = true;
  if (s.fontStyle === 'italic' || s.fontStyle === 'oblique') st.italic = true;
  const deco = (s.textDecoration || '') + ' ' + (cssProp(css, 'text-decoration-line') || '');
  if (/underline/.test(deco)) st.underline = /double/.test(cssProp(css, 'text-underline-style') || '') ? 'double' : 'single';
  if (/line-through/.test(deco)) st.strike = true;
  const color = normColor(s.color);
  if (color && color !== '#000000') st.color = color;
  const bg = normColor(s.backgroundColor || cssProp(css, 'background') || el.getAttribute('bgcolor') || '');
  if (bg && bg !== '#ffffff') st.fill = bg;
  const ta = (s.textAlign || el.getAttribute('align') || '').toLowerCase();
  if (ta === 'center' || ta === 'right' || ta === 'left' || ta === 'justify') st.hAlign = ta;
  const va = (s.verticalAlign || el.getAttribute('valign') || '').toLowerCase();
  if (va === 'top') st.vAlign = 'top'; else if (va === 'middle') st.vAlign = 'middle';
  const fs = s.fontSize;
  if (fs) { const m = /([\d.]+)(px|pt)/.exec(fs); if (m) st.size = Math.round((m[2] === 'px' ? +m[1] * 0.75 : +m[1]) * 2) / 2; }
  if (s.fontFamily) st.font = s.fontFamily.split(',')[0].replace(/["']/g, '').trim();
  const ws = s.whiteSpace || cssProp(css, 'white-space') || '';
  if (ws === 'normal' || ws === 'pre-wrap') st.wrap = true;
  const all = cssBorder(s.border || cssProp(css, 'border'));
  const side = (k: string) => cssBorder((s as any)[`border${k}`] || cssProp(css, `border-${k.toLowerCase()}`)) || all;
  const bt = side('Top'), bb = side('Bottom'), bl = side('Left'), br = side('Right');
  if (bt) st.bt = bt; if (bb) st.bb = bb; if (bl) st.bl = bl; if (br) st.br = br;
  const ind = cssProp(css, 'mso-char-indent-count'); if (ind && +ind > 0) st.indent = Math.round(+ind);
  const rot = cssProp(css, 'mso-rotate'); if (rot && +rot) st.rotate = Math.max(-90, Math.min(90, +rot));
  // Inline wrappers that cover the whole cell (<b><i><font color>...): walk single-child chains.
  let node: Element = el;
  for (let depth = 0; depth < 6; depth++) {
    const kids = Array.from(node.children).filter((k) => !/^(COMMENT|A)$/.test(k.tagName) || (k.tagName === 'A' && !k.classList.contains('comment-indicator')));
    const only = kids.length === 1 ? kids[0] : null;
    if (!only || (only.textContent || '').trim() !== cellText(el)) break;
    const t = only.tagName;
    if (/^(B|STRONG)$/.test(t)) st.bold = true;
    else if (/^(I|EM)$/.test(t)) st.italic = true;
    else if (t === 'U') st.underline = 'single';
    else if (/^(S|STRIKE|DEL)$/.test(t)) st.strike = true;
    else if (t === 'SUP') st.vertAlign = 'superscript';
    else if (t === 'SUB') st.vertAlign = 'subscript';
    else if (t === 'FONT') {
      const c = normColor(only.getAttribute('color')); if (c && c !== '#000000') st.color = c;
      const f = only.getAttribute('face'); if (f) st.font = f.split(',')[0].trim();
      const sz = +(only.getAttribute('size') || 0); if (sz >= 1 && sz <= 7) st.size = [8, 10, 12, 14, 18, 24, 36][sz - 1];
    } else if (t === 'SPAN') {
      const inner = styleFromElement(only as HTMLElement, only.getAttribute('style') || '');
      for (const k of ['bold', 'italic', 'underline', 'strike', 'color', 'font', 'size'] as const) if (inner[k] !== undefined) (st as any)[k] = inner[k];
    } else if (t !== 'A' && t !== 'P' && t !== 'DIV') break;
    node = only;
  }
  const fmt = htmlNumFmt(cssProp(css, 'mso-number-format'));
  if (fmt) st.numFmt = fmt;
  return st;
}

function cellText(el: HTMLElement): string {
  let out = '';
  const walk = (n: Node) => {
    if (n.nodeType === 3) { out += (n.nodeValue || '').replace(/[ \t\r\n]+/g, ' '); return; }
    if (n.nodeType !== 1) return;
    const e = n as Element;
    const tag = e.tagName;
    if (tag === 'BR') { out += '\n'; return; }
    if (tag === 'STYLE' || tag === 'SCRIPT' || tag === 'COMMENT') return;
    if (tag === 'A' && e.classList.contains('comment-indicator')) return;
    const block = /^(P|DIV|LI|TR|H[1-6])$/.test(tag);
    if (block && out && !out.endsWith('\n')) out += '\n';
    n.childNodes.forEach(walk);
    if (block && !out.endsWith('\n')) out += '\n';
  };
  el.childNodes.forEach(walk);
  return out.replace(/ *\n */g, '\n').replace(/^\s+|\s+$/g, '');
}

/** Google Sheets / LibreOffice data-sheets-value: {"1": type, "2": text, "3": number, "4": boolean}. */
function sheetsValue(el: Element): Cell['v'] | undefined {
  const raw = el.getAttribute('data-sheets-value');
  if (!raw) return undefined;
  try {
    const j = JSON.parse(raw);
    switch (j['1']) {
      case 2: return j['2'] ?? '';
      case 3: return typeof j['3'] === 'number' ? j['3'] : undefined;
      case 4: return !!j['4'];
    }
  } catch { /* not JSON */ }
  return undefined;
}

export function importHtml(text: string, warnings: string[] = [], opts: HtmlImportOptions = { allTables: true }): Workbook {
  const doc = new DOMParser().parseFromString(text, 'text/html');
  const book = new Workbook(false);
  const rules = parseCssRules(doc);
  let tables = Array.from(doc.querySelectorAll('table')).filter((t) => !t.parentElement?.closest('table'));
  if (!opts.allTables) tables = tables.slice(0, 1);
  const styleCache = new Map<string, number>();
  const sid = (st: CellStyle) => { const k = JSON.stringify(st); let s = styleCache.get(k); if (s == null) { s = book.styleId(st); styleCache.set(k, s); } return s; };
  let r1c1: ((f: string, r: number, c: number) => string) | null = null;
  tables.forEach((table, ti) => {
    const cap = table.querySelector('caption')?.textContent?.trim();
    // LibreOffice puts the sheet name in a heading just before each table.
    const before = ti > 0 ? tables[ti - 1] : null;
    const hs = Array.from(doc.querySelectorAll('h1,h2,h3')).filter((h) => (h.compareDocumentPosition(table) & 4) && (!before || (before.compareDocumentPosition(h) & 4)));
    const heading = hs.length ? (hs[hs.length - 1].textContent || '').replace(/^Sheet \d+:\s*/, '').trim() : '';
    let name = (cap || heading || (tables.length > 1 ? `Table${ti + 1}` : 'Sheet1')).replace(/[\\/?*[\]:]/g, '_').slice(0, 31);
    while (book.sheetByName(name)) name = name.slice(0, 28) + `_${ti + 1}`;
    const sheet = new Sheet(book, name);
    book.sheets.push(sheet);
    const occupied = new Set<number>();
    const rows = Array.from(table.rows);
    rows.forEach((tr, r) => {
      let c = 0;
      const h = parseFloat(tr.style.height || tr.getAttribute('height') || '');
      if (h > 0) sheet.rowH.set(r, Math.round(/pt/.test(tr.style.height) ? (h * 4) / 3 : h));
      for (const td of Array.from(tr.cells)) {
        while (occupied.has(cellKey(r, c))) c++;
        const rs = Math.max(1, Math.min(1000, parseInt(td.getAttribute('rowspan') || '1', 10) || 1));
        const cs = Math.max(1, Math.min(1000, parseInt(td.getAttribute('colspan') || '1', 10) || 1));
        for (let i = 0; i < rs; i++) for (let j = 0; j < cs; j++) occupied.add(cellKey(r + i, c + j));
        if (rs > 1 || cs > 1) sheet.merges.push({ r1: r, c1: c, r2: r + rs - 1, c2: c + cs - 1 });
        const css = cssFor(td, rules);
        const st = styleFromElement(td, css);
        const sdnum = td.getAttribute('sdnum');
        if (sdnum && !st.numFmt) { const f = htmlNumFmt(sdnum.split(';').slice(2).join(';')); if (f) st.numFmt = f; }
        const gnf = td.getAttribute('data-sheets-numberformat');
        if (gnf && !st.numFmt) { try { const f = htmlNumFmt(JSON.parse(gnf)['2']); if (f) st.numFmt = f; } catch { /* ignore */ } }
        const txt = cellText(td);
        const xnum = td.getAttribute('x:num');
        const sdval = td.getAttribute('sdval');
        const gv = sheetsValue(td);
        let conv: { v: Cell['v']; f?: string; fmt?: string };
        if (xnum != null && xnum !== '' && Number.isFinite(Number(xnum))) conv = { v: Number(xnum) };
        else if (sdval != null && sdval !== '' && Number.isFinite(Number(sdval))) conv = { v: Number(sdval) };
        else if (gv !== undefined) conv = { v: gv };
        else if (st.numFmt === '@' || td.getAttribute('x:str') != null) conv = { v: txt || null };
        else conv = txt.includes('\n') ? { v: txt } : convertField(txt);
        if (conv.f) conv = { v: txt };
        if (td.getAttribute('x:bool') != null) conv = { v: /^(true|1)$/i.test(td.getAttribute('x:bool') || txt) };
        if (td.getAttribute('x:err') != null) conv = { v: new CellError((td.getAttribute('x:err') || txt) as CellError['code']) };
        if (conv.fmt && !st.numFmt) st.numFmt = conv.fmt;
        const cell: Cell = {};
        if (conv.v != null && conv.v !== '') cell.v = conv.v;
        const fx = td.getAttribute('x:fmla');
        const gf = td.getAttribute('data-sheets-formula');
        if (fx && fx.length > 1) cell.f = fx.replace(/^=/, '');
        else if (gf && gf.length > 1) {
          if (!r1c1) r1c1 = r1c1ToA1Lite;
          cell.f = r1c1(gf.replace(/^=/, ''), r, c);
        }
        if (txt.includes('\n')) st.wrap = true;
        const s = sid(st);
        if (s) cell.s = s;
        const a = Array.from(td.querySelectorAll('a[href]')).find((x) => !x.classList.contains('comment-indicator'));
        const href = a?.getAttribute('href');
        if (href && !isScriptUrl(href)) cell.link = href;
        const cm = td.querySelector('comment');
        if (cm && cm.textContent?.trim()) sheet.comments.set(cellKey(r, c), { author: '', text: cm.textContent.trim(), time: 0, note: true });
        if (cell.v != null || cell.s || cell.link || cell.f) sheet.cells.set(cellKey(r, c), cell);
        const w = parseFloat(td.style.width || td.getAttribute('width') || '');
        if (w > 0 && cs === 1 && !/%/.test(td.style.width || td.getAttribute('width') || '') && !sheet.colW.has(c)) sheet.colW.set(c, Math.round(/pt/.test(td.style.width) ? (w * 4) / 3 : w));
        c += cs;
      }
    });
    // <col width> and <colgroup span width>
    let ci = 0;
    for (const g of Array.from(table.querySelectorAll('colgroup, col'))) {
      if (g.tagName === 'COLGROUP' && g.querySelector('col')) continue;
      const span = Math.max(1, parseInt(g.getAttribute('span') || '1', 10) || 1);
      const w = parseFloat((g as HTMLElement).style.width || g.getAttribute('width') || '');
      for (let k = 0; k < span && ci < 16384; k++, ci++) if (w > 0) sheet.colW.set(ci, Math.round(/pt/.test((g as HTMLElement).style.width) ? (w * 4) / 3 : w));
    }
  });
  if (!book.sheets.length) {
    // No table: put text lines in column A.
    const sheet = new Sheet(book, 'Sheet1');
    book.sheets.push(sheet);
    const body = doc.body ? cellText(doc.body) : '';
    body.split('\n').filter((l) => l.trim()).forEach((l, i) => sheet.cells.set(cellKey(i, 0), { v: l.trim() }));
    if (body.trim()) warnings.push('The page has no tables; its text was placed in column A.');
  }
  book.clearHistory();
  return book;
}

/** R1C1 references (RC[-2], R1C1, R[1]C) to A1, relative to row r and column c (zero-based). */
function r1c1ToA1Lite(f: string, r: number, c: number): string {
  const col = (n: number) => { let s = ''; let x = n + 1; while (x > 0) { const m = (x - 1) % 26; s = String.fromCharCode(65 + m) + s; x = Math.floor((x - 1) / 26); } return s; };
  return f.replace(/(?<![A-Za-z0-9_.$"])R(\[-?\d+\]|\d+)?C(\[-?\d+\]|\d+)?(?![A-Za-z0-9_(\[])/g, (_m, rs?: string, cs?: string) => {
    const part = (spec: string | undefined, base: number) => (!spec ? { n: base, abs: false } : spec.startsWith('[') ? { n: base + parseInt(spec.slice(1, -1), 10), abs: false } : { n: parseInt(spec, 10) - 1, abs: true });
    const R = part(rs, r), C = part(cs, c);
    if (R.n < 0 || C.n < 0) return '#REF!';
    return `${C.abs ? '$' : ''}${col(C.n)}${R.abs ? '$' : ''}${R.n + 1}`;
  });
}

// --------------------------------------------------------------------------------------------
// Export

const escHtml = (s: string) => s.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

function borderCss(b: Border): string {
  const w = b.style === 'thick' ? 3 : b.style === 'medium' || b.style === 'mediumDashed' ? 2 : 1;
  const st = b.style === 'double' ? 'double' : /dash/i.test(b.style) ? 'dashed' : b.style === 'dotted' || b.style === 'hair' ? 'dotted' : 'solid';
  return `${b.style === 'double' ? 3 : w}px ${st} ${b.color || '#000000'}`;
}

export function styleCss(st: CellStyle, isNumber = false): string {
  const c: string[] = [];
  if (st.font) c.push(`font-family:'${st.font.replace(/'/g, '')}'`);
  if (st.size) c.push(`font-size:${st.size}pt`);
  if (st.bold) c.push('font-weight:bold');
  if (st.italic) c.push('font-style:italic');
  const deco = [st.underline ? 'underline' : '', st.strike ? 'line-through' : ''].filter(Boolean).join(' ');
  if (deco) c.push(`text-decoration:${deco}`);
  if (st.color) c.push(`color:${st.color}`);
  if (st.fill) c.push(`background-color:${st.fill}`);
  const ha = st.hAlign && st.hAlign !== 'general' ? (st.hAlign === 'centerContinuous' ? 'center' : st.hAlign === 'distributed' || st.hAlign === 'fill' ? 'justify' : st.hAlign) : isNumber ? 'right' : '';
  if (ha) c.push(`text-align:${ha}`);
  if (st.vAlign) c.push(`vertical-align:${st.vAlign === 'distributed' || st.vAlign === 'justify' ? 'middle' : st.vAlign}`);
  if (st.wrap) c.push('white-space:pre-wrap'); else c.push('white-space:pre');
  if (st.indent) c.push(`padding-left:${st.indent * 9}px`);
  if (st.bt) c.push(`border-top:${borderCss(st.bt)}`);
  if (st.bb) c.push(`border-bottom:${borderCss(st.bb)}`);
  if (st.bl) c.push(`border-left:${borderCss(st.bl)}`);
  if (st.br) c.push(`border-right:${borderCss(st.br)}`);
  return c.join(';');
}

function displayText(sheet: Sheet, r: number, c: number): { text: string; color?: string; isNumber?: boolean } {
  const cell = sheet.get(r, c);
  const v = cell?.v ?? null;
  if (v == null) return { text: '' };
  if (v instanceof CellError) return { text: v.code };
  const st = sheet.styleOf(r, c);
  const f = formatValue(v, st.numFmt, { date1904: sheet.book.calc.date1904 });
  return { text: f.text, color: f.color, isNumber: f.isNumber };
}

function sheetTable(sheet: Sheet, range?: Range | null): string {
  const rg = range || sheet.usedRange();
  if (!rg) return '<table></table>';
  const covered = new Set<number>();
  const spans = new Map<number, Range>();
  for (const m of sheet.merges) {
    if (m.r1 < rg.r1 || m.c1 < rg.c1 || m.r1 > rg.r2 || m.c1 > rg.c2) continue;
    spans.set(cellKey(m.r1, m.c1), m);
    for (let r = m.r1; r <= m.r2; r++) for (let c = m.c1; c <= m.c2; c++) if (r !== m.r1 || c !== m.c1) covered.add(cellKey(r, c));
  }
  const out: string[] = ['<table style="border-collapse:collapse">', '<colgroup>'];
  for (let c = rg.c1; c <= rg.c2; c++) if (!sheet.hiddenCols.has(c)) out.push(`<col style="width:${safeNum(sheet.colWidth(c), 0, 0, 1e5)}px">`);
  out.push('</colgroup>');
  for (let r = rg.r1; r <= rg.r2; r++) {
    if (sheet.hiddenRows.has(r)) continue;
    const h = safeNum(sheet.rowH.get(r), 0, 0, 1e5);
    out.push(h ? `<tr style="height:${h}px">` : '<tr>');
    for (let c = rg.c1; c <= rg.c2; c++) {
      if (sheet.hiddenCols.has(c)) continue;
      const k = cellKey(r, c);
      if (covered.has(k)) continue;
      const m = spans.get(k);
      const span = m ? `${m.r2 > m.r1 ? ` rowspan="${Math.min(m.r2, rg.r2) - m.r1 + 1}"` : ''}${m.c2 > m.c1 ? ` colspan="${Math.min(m.c2, rg.c2) - m.c1 + 1}"` : ''}` : '';
      const d = displayText(sheet, r, c);
      const st = { ...sheet.styleOf(r, c) };
      if (d.color) st.color = d.color;
      const css = styleCss(st, d.isNumber);
      const cell = sheet.get(r, c);
      let body = escHtml(d.text).replace(/\n/g, '<br>');
      if (cell?.link && !isScriptUrl(cell.link)) body = `<a href="${escHtml(cell.link)}">${body}</a>`;
      out.push(`<td${span}${css ? ` style="${escHtml(css)}"` : ''}>${body}</td>`);
    }
    out.push('</tr>');
  }
  out.push('</table>');
  return out.join('\n');
}

export function exportHtml(book: Workbook, opts: { sheetIndex?: number; range?: Range } = {}): Blob {
  const sheets = opts.sheetIndex != null ? [book.sheets[opts.sheetIndex]].filter(Boolean) : book.sheets.filter((s) => s.visibility === 'visible');
  const title = book.props.title || (sheets.length === 1 ? sheets[0].name : 'Workbook');
  const body = sheets.map((s) => (sheets.length > 1 ? `<h2>${escHtml(s.name)}</h2>\n` : '') + sheetTable(s, opts.range)).join('\n');
  const html = `<!DOCTYPE html>\n<html>\n<head>\n<meta charset="utf-8">\n<meta name="generator" content="Anne\'s Grid">\n<title>${escHtml(title)}</title>\n<style>body{font-family:Calibri,Carlito,Arial,sans-serif;font-size:11pt}table{border-collapse:collapse;margin-bottom:1.5em}td{padding:1px 4px;vertical-align:bottom;overflow:hidden}</style>\n</head>\n<body>\n${body}\n</body>\n</html>\n`;
  return new Blob([html], { type: 'text/html;charset=utf-8' });
}

// --------------------------------------------------------------------------------------------
// Markdown

export function sheetToMarkdown(sheet: Sheet, range?: Range | null): string {
  const rg = range || sheet.usedRange(false);
  if (!rg) return '';
  const rows: string[][] = [];
  const numeric: number[] = new Array(rg.c2 - rg.c1 + 1).fill(0);
  for (let r = rg.r1; r <= rg.r2; r++) {
    if (sheet.hiddenRows.has(r)) continue;
    const row: string[] = [];
    for (let c = rg.c1; c <= rg.c2; c++) {
      if (sheet.hiddenCols.has(c)) continue;
      const d = displayText(sheet, r, c);
      if (d.isNumber && r > rg.r1) numeric[c - rg.c1]++;
      row.push(d.text.replace(/\\/g, '\\\\').replace(/\|/g, '\\|').replace(/\r?\n/g, '<br>'));
    }
    rows.push(row);
  }
  if (!rows.length) return '';
  const cols = rows[0].length;
  const visibleCols = [];
  for (let c = rg.c1; c <= rg.c2; c++) if (!sheet.hiddenCols.has(c)) visibleCols.push(c - rg.c1);
  const dataRows = rows.length - 1;
  const sep = visibleCols.map((ci) => (dataRows && numeric[ci] > dataRows / 2 ? '---:' : '---'));
  const line = (r: string[]) => '| ' + Array.from({ length: cols }, (_, i) => r[i] ?? '').join(' | ') + ' |';
  return [line(rows[0]), '| ' + sep.join(' | ') + ' |', ...rows.slice(1).map(line)].join('\n') + '\n';
}

export function exportMarkdown(book: Workbook, sheetIndex = 0, range?: Range): Blob {
  const sheet = book.sheets[Math.max(0, Math.min(sheetIndex, book.sheets.length - 1))];
  return new Blob([sheetToMarkdown(sheet, range)], { type: 'text/markdown;charset=utf-8' });
}

export { keyRow, keyCol };
