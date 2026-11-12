/**
 * Utilities tab, part 4: Web Data, Info & List, File & System, Import, Export, Start, My Tools & Shortcuts,
 * Options, Find and Run a Tool, Start Last Tool Again (ASAP Utilities-style tools, re-implemented; see
 * utilities.ts).
 */
import type { App } from '../app';
import { Sheet, Workbook } from '../model/workbook';
import type { Cell, Scalar } from '../model/types';
import { CellError } from '../model/types';
import { cellKey, keyRow, keyCol, colName, cellName, rangeName, type Range } from '../model/address';
import { selRanges, activeRange, activeCell, clipToUsed, rangeToTsv } from './ops';
import { defTools, externalLinkIn, U, selCells, selBlock, singleCell, hasContent, sheetProtected, structureLocked, emptySheet, numOf, textOf, plural, report, pickSheets, sheetsByNames, shownText, argOrAsk, writeOut, checkWritable, canChange, changeCells, copyText, saveTextFile, pickFiles, textRows, valueRows, toCsvField, escapeHtml, fileStem, nowStamp, toolLog, writeBlock, addSheetWithRows, HEAD } from './utilShared';
import { menuOf } from './utilities';
import { transformValues } from './tools';
import { openDialog, toast, promptDialog, choiceDialog, alertDialog, confirmDialog } from '../ui/dialog';
import { getCommand, allCommands, searchCommands, type MenuEntry } from './registry';
import { workbookStats } from './review';
import { importFile, exportFile, FORMATS } from '../io';
import { sheetToMarkdown } from '../io/html';
import { kvGet, kvSet } from '../util/store';

const G_WEB = 'Web Data';
const G_INFO = 'Info & List';
const G_FILE = 'File & System';
const G_IMPORT = 'Import';
const G_EXPORT = 'Export';
const G_START = 'Start';
const G_MY = 'My Tools & Shortcuts';

// ============================================================================================
// Web Data

const ENTITIES: Record<string, string> = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ', copy: '©', reg: '®', trade: '™', hellip: '…', mdash: '—', ndash: '–', lsquo: '‘', rsquo: '’', ldquo: '“', rdquo: '”', euro: '€', pound: '£', yen: '¥', cent: '¢', deg: '°', middot: '·', bull: '•', laquo: '«', raquo: '»', times: '×', divide: '÷', frac12: '½', frac14: '¼', frac34: '¾', eacute: 'é', egrave: 'è', agrave: 'à', aacute: 'á', ouml: 'ö', uuml: 'ü', auml: 'ä', szlig: 'ß', ccedil: 'ç', ntilde: 'ñ', iacute: 'í', oacute: 'ó', uacute: 'ú' };

export function decodeEntities(s: string): string {
  return s.replace(/&(#x[0-9a-f]+|#\d+|[a-z]+[0-9]*);/gi, (m, e: string) => {
    if (e[0] === '#') { const n = e[1].toLowerCase() === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10); return isNaN(n) ? m : String.fromCodePoint(n); }
    return ENTITIES[e.toLowerCase()] ?? m;
  });
}
export function encodeEntities(s: string): string { return s.replace(/[&<>"']/g, (ch) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]!)).replace(/[^\x00-\x7f]/gu, (ch) => `&#${ch.codePointAt(0)};`); }
export function stripTags(s: string): string { return decodeEntities(s.replace(/<br\s*\/?>/gi, '\n').replace(/<\/(p|div|li|tr|h[1-6])>/gi, '\n').replace(/<\/?[a-zA-Z!][^>]*>/g, '')).replace(/[ \t]+\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim(); }
export function urlDecode(s: string): string { try { return decodeURIComponent(s.replace(/\+/g, ' ')); } catch { return s; } }
export function urlEncode(s: string): string { return encodeURIComponent(s); }
export function urlParts(s: string): { protocol: string; host: string; path: string; query: string; hash: string } | null { try { const u = new URL(s.trim()); return { protocol: u.protocol.replace(':', ''), host: u.hostname, path: u.pathname, query: u.search.replace(/^\?/, ''), hash: u.hash.replace(/^#/, '') }; } catch { return null; } }
export const URL_RE = /\b((?:https?|ftp):\/\/[^\s<>"']+|www\.[^\s<>"']+)/gi;
export const EMAIL_RE = /\b[\w.+-]+@[\w-]+(?:\.[\w-]+)+\b/g;

/** An HTML table for a block of cells, with optional inline styles. */
export function htmlTable(sheet: Sheet, rg: Range, opts: { styles?: boolean; header?: boolean; caption?: string } = {}): string {
  const rows = textRows(sheet, rg);
  const out: string[] = ['<table' + (opts.styles ? ' style="border-collapse:collapse"' : '') + '>'];
  if (opts.caption) out.push(`<caption>${escapeHtml(opts.caption)}</caption>`);
  rows.forEach((row, i) => {
    const tag = opts.header && i === 0 ? 'th' : 'td';
    out.push('<tr>' + row.map((t, j) => { let st = ''; if (opts.styles) { const s = sheet.styleOf(rg.r1 + i, rg.c1 + j); const parts: string[] = ['border:1px solid #ccc', 'padding:2px 6px']; if (s.bold) parts.push('font-weight:bold'); if (s.italic) parts.push('font-style:italic'); if (s.fill) parts.push(`background:${s.fill}`); if (s.color) parts.push(`color:${s.color}`); if (typeof sheet.value(rg.r1 + i, rg.c1 + j) === 'number' || s.hAlign === 'right') parts.push('text-align:right'); st = ` style="${parts.join(';')}"`; } return `<${tag}${st}>${escapeHtml(t).replace(/\n/g, '<br>')}</${tag}>`; }).join('') + '</tr>');
  });
  out.push('</table>');
  return out.join('\n');
}

defTools([
  { id: 'util.stripHtml', label: 'Remove HTML Tags', icon: 'code_off', tab: U, group: G_WEB, edits: true, popular: true, tip: 'Strip <tags> from the selected text and decode &amp; style entities; <br> and </p> become line breaks', keywords: 'html tags strip clean web', origin: 'ASAP Utilities (Remove HTML tags)', run: (app) => transformValues(app, 'Remove HTML Tags', (v) => stripTags(String(v))) },
  { id: 'util.decodeEntities', label: 'Decode HTML Entities', icon: 'translate', tab: U, group: G_WEB, edits: true, tip: '&amp; becomes &, &eacute; becomes é, &#8217; becomes ’', origin: 'ASAP Utilities', run: (app) => transformValues(app, 'Decode Entities', (v) => decodeEntities(String(v))) },
  { id: 'util.encodeEntities', label: 'Encode as HTML Entities', icon: 'code', tab: U, group: G_WEB, edits: true, tip: 'Make text safe for HTML: < > & " and non-ASCII characters become entities', run: (app) => transformValues(app, 'Encode Entities', (v) => encodeEntities(String(v))) },
  { id: 'util.urlDecode', label: 'Decode URL Encoding', icon: 'link', tab: U, group: G_WEB, edits: true, tip: '%20 becomes a space, %C3%A9 becomes é', origin: 'ASAP Utilities (URL decode)', run: (app) => transformValues(app, 'URL Decode', (v) => urlDecode(String(v))) },
  { id: 'util.urlEncode', label: 'URL-Encode Text', icon: 'link', tab: U, group: G_WEB, edits: true, run: (app) => transformValues(app, 'URL Encode', (v) => urlEncode(String(v))) },
  {
    id: 'util.activateLinks', label: 'Turn Web and Email Addresses into Links', icon: 'add_link', tab: U, group: G_WEB, edits: true, popular: true, tip: 'Cells that contain a web address or an email address become clickable hyperlinks', keywords: 'hyperlink activate urls mailto', origin: 'ASAP Utilities / Kutools (Convert Hyperlinks)',
    run: (app) => { const s = app.sheet; let n = 0; app.book.transact('Activate Links', () => { for (const x of selCells(app, { existingOnly: true })) { const cell = x.cell!; if (cell.link || typeof cell.v !== 'string') continue; const t = cell.v.trim(); let link: string | null = null; if (/^(https?:\/\/|ftp:\/\/|www\.)\S+$/i.test(t)) link = /^www\./i.test(t) ? 'http://' + t : t; else if (/^[\w.+-]+@[\w-]+(?:\.[\w-]+)+$/.test(t)) link = 'mailto:' + t; if (!link) continue; s.set(x.r, x.c, { ...cell, link, s: app.book.derivedStyle(cell.s ?? 0, { color: '#0563c1', underline: 'single' }) }); n++; } }); toast(n ? `${plural(n, 'link')} activated` : 'No web or email addresses found.'); },
  },
  {
    id: 'util.extractUrls', label: 'Extract Link Addresses', icon: 'link', tab: U, group: G_WEB, edits: true, kind: 'split', tip: 'Write each cell\'s hyperlink address into the next column, or replace the cell text with it', origin: 'ASAP Utilities / Kutools (Convert Hyperlinks)',
    run: (app, arg?: 'right' | 'replace') => { const s = app.sheet; let n = 0; app.book.transact('Extract Links', () => { for (const x of selCells(app, { existingOnly: true })) { const l = x.cell?.link; if (!l) continue; if (arg === 'replace') s.set(x.r, x.c, { ...(x.cell!.s ? { s: x.cell!.s } : {}), v: l }); else writeOut(s, x.r, x.c + 1, l); n++; } }); toast(n ? `${plural(n, 'address')} extracted` : 'No hyperlinks in the selection.'); },
    menu: () => [{ id: 'util.extractUrls', arg: 'right', label: 'Into the column to the right' }, { id: 'util.extractUrls', arg: 'replace', label: 'Replace the cell text with the address' }],
  },
  { id: 'util.linksFromTwoColumns', label: 'Make Links from Text and Address Columns', icon: 'link', tab: U, group: G_WEB, edits: true, tip: 'Selected column = the text to show; the column to its right = the address; the result is a hyperlink in the first column', run: (app) => { const s = app.sheet; const rg = clipToUsed(s, activeRange(app)); let n = 0; app.book.transact('Links from Columns', () => { for (let r = rg.r1; r <= rg.r2; r++) { const addr = shownText(s, r, rg.c1 + 1).trim(); if (!addr) continue; const cur = s.get(r, rg.c1) || {}; s.set(r, rg.c1, { ...cur, v: cur.v ?? addr, link: /^(https?|ftp|mailto|#)/i.test(addr) ? addr : 'http://' + addr, s: app.book.derivedStyle(cur.s ?? 0, { color: '#0563c1', underline: 'single' }) }); n++; } }); toast(`${plural(n, 'link')} made`); } },
  { id: 'util.urlParts', label: 'Split Web Addresses into Parts', icon: 'call_split', tab: U, group: G_WEB, edits: true, tip: 'Protocol, host, path, query and fragment of each selected URL in the columns to the right', run: (app) => { const s = app.sheet; const rg = clipToUsed(s, activeRange(app)); const target = { r1: rg.r1, r2: rg.r2, c1: rg.c2 + 1, c2: rg.c2 + 5 }; if (!checkWritable(app, [target])) return; let n = 0; app.book.transact('URL Parts', () => { for (let r = rg.r1; r <= rg.r2; r++) { const p = urlParts(shownText(s, r, rg.c1) || s.get(r, rg.c1)?.link || ''); if (!p) continue; [p.protocol, p.host, p.path, p.query, p.hash].forEach((v, i) => writeOut(s, r, target.c1 + i, v)); n++; } }); toast(`${plural(n, 'address')} split`); } },
  {
    id: 'util.copyAsHtml', label: 'Copy Selection as an HTML Table', icon: 'html', tab: U, group: G_WEB, noRecord: true, popular: true, tip: 'The selection as <table> markup on the clipboard, ready for a web page, wiki or email', keywords: 'html table export clipboard', origin: 'ASAP Utilities (Export selection as HTML)',
    run: async (app, arg?: { styles?: boolean }) => { const s = app.sheet; const rg = clipToUsed(s, activeRange(app)); await copyText(htmlTable(s, rg, { styles: arg?.styles !== false, header: true }), 'HTML table'); },
  },
  { id: 'util.copyAsMarkdown', label: 'Copy Selection as a Markdown Table', icon: 'article', tab: U, group: G_WEB, noRecord: true, run: async (app) => { const s = app.sheet; await copyText(sheetToMarkdown(s, clipToUsed(s, activeRange(app))), 'Markdown table'); } },
  { id: 'util.copyAsWiki', label: 'Copy Selection as a Wiki Table', icon: 'article', tab: U, group: G_WEB, noRecord: true, tip: 'MediaWiki table markup', run: async (app) => { const s = app.sheet; const rows = textRows(s, clipToUsed(s, activeRange(app))); const text = ['{| class="wikitable"', ...rows.map((row, i) => (i === 0 ? '! ' + row.join(' !! ') : '|-\n| ' + row.join(' || '))), '|}'].join('\n'); await copyText(text, 'Wiki table'); } },
  { id: 'util.copyAsBbcode', label: 'Copy Selection as a BBCode Table', icon: 'article', tab: U, group: G_WEB, noRecord: true, run: async (app) => { const s = app.sheet; const rows = textRows(s, clipToUsed(s, activeRange(app))); await copyText('[table]\n' + rows.map((row) => '[tr]' + row.map((t) => `[td]${t}[/td]`).join('') + '[/tr]').join('\n') + '\n[/table]', 'BBCode table'); } },
  { id: 'util.openLinks', label: 'Open All Selected Links', icon: 'open_in_new', tab: U, group: G_WEB, noRecord: true, tip: 'Open every hyperlink in the selection in the browser (up to 20)', origin: 'ASAP Utilities', run: (app) => { const links = selCells(app, { existingOnly: true }).map((x) => x.cell!.link).filter((l): l is string => !!l).slice(0, 20); if (!links.length) { toast('No links in the selection.'); return; } for (const l of links) app.followLink(l); } },
  { id: 'util.menuWeb', label: 'Web Data', icon: 'public', tab: U, group: G_WEB, kind: 'menu', noRecord: true, run: () => undefined, menu: () => menuOf(['util.stripHtml', 'util.decodeEntities', 'util.encodeEntities', 'util.urlDecode', 'util.urlEncode', null, 'util.activateLinks', 'util.extractUrls', 'util.linksFromTwoColumns', 'util.urlParts', 'util.openLinks', 'insert.removeLink', null, 'util.copyAsHtml', 'util.copyAsMarkdown', 'util.copyAsWiki', 'util.copyAsBbcode', 'data.fromWeb']) },
]);

// ============================================================================================
// Info & List

function selectionInfo(app: App): [string, string][] {
  const s = app.sheet;
  let cells = 0, filled = 0, numbers = 0, text = 0, formulas = 0, errors = 0, blanks = 0, sum = 0, min = Infinity, max = -Infinity, chars = 0, words = 0, dates = 0;
  const distinct = new Set<string>();
  for (const x of selCells(app, { includeFiltered: true })) {
    cells++;
    const cell = x.cell;
    if (!hasContent(cell)) { blanks++; continue; }
    filled++;
    if (cell!.f != null) formulas++;
    const v = cell!.v;
    if (typeof v === 'number') { numbers++; sum += v; if (v < min) min = v; if (v > max) max = v; if (/[dmy]/i.test(s.styleOf(x.r, x.c).numFmt || '')) dates++; }
    else if (typeof v === 'string') { text++; chars += v.length; words += v.trim() ? v.trim().split(/\s+/).length : 0; }
    else if (v instanceof CellError) errors++;
    distinct.add(shownText(s, x.r, x.c).toLowerCase());
  }
  const rg = selRanges(app);
  return [['Areas', String(rg.length)], ['Address', rg.map(rangeName).join(', ')], ['Cells', cells.toLocaleString()], ['Filled', filled.toLocaleString()], ['Blank', blanks.toLocaleString()], ['Numbers', numbers.toLocaleString()], ['Dates', dates.toLocaleString()], ['Text', text.toLocaleString()], ['Formulas', formulas.toLocaleString()], ['Errors', errors.toLocaleString()], ['Distinct values', distinct.size.toLocaleString()], ['Sum', numbers ? String(+sum.toPrecision(15)) : '—'], ['Average', numbers ? String(+(sum / numbers).toPrecision(12)) : '—'], ['Minimum', numbers ? String(min) : '—'], ['Maximum', numbers ? String(max) : '—'], ['Characters', chars.toLocaleString()], ['Words', words.toLocaleString()]];
}

function narrowColumns(s: Sheet): { c: number; text: string }[] {
  const out: { c: number; text: string }[] = [];
  const u = s.usedRange(false); if (!u) return out;
  for (let c = u.c1; c <= u.c2; c++) {
    if (s.hiddenCols.has(c)) continue;
    const w = s.colWidth(c) || s.defaultColW;
    for (let r = u.r1; r <= Math.min(u.r2, u.r1 + 2000); r++) { const v = s.value(r, c); if (typeof v !== 'number') continue; const st = s.styleOf(r, c); const t = shownText(s, r, c); const px = t.length * ((st.size || 11) * 0.6) + 6; if (px > w) { out.push({ c, text: `${cellName(r, c)}: "${t}" needs about ${Math.round(px)} px, column is ${Math.round(w)} px` }); break; } }
  }
  return out;
}

defTools([
  { id: 'util.selectionInfo', label: 'Selection Information', icon: 'info', tab: U, group: G_INFO, noRecord: true, popular: true, tip: 'Counts, sums, distinct values, characters and words of the selected cells', origin: 'ASAP Utilities (Information about selected cells)', run: async (app) => { const { h } = await import('../ui/dom'); const table = h('table', { class: 'ql-info-table' }, ...selectionInfo(app).map(([k, v]) => h('tr', null, h('th', null, k), h('td', null, v)))); await alertDialog('Selection Information', table); } },
  { id: 'util.workbookInfo', label: 'Workbook and Sheet Statistics', icon: 'analytics', tab: U, group: G_INFO, noRecord: true, run: (app) => app.exec('review.workbookStats') },
  { id: 'util.listNames', label: 'List All Defined Names', icon: 'badge', tab: U, group: G_INFO, edits: true, tip: 'A sheet with every defined name, its reference, scope and current value', origin: 'ASAP Utilities (List all named ranges)', run: (app) => { const names = app.book.names; if (!names.length) { toast('The workbook has no defined names.'); return; } const rows = names.map((n) => { let v: Scalar = null; try { v = app.engine.evaluateScalar(n.ref.replace(/^=/, ''), app.sheet, 0, 0); } catch { v = null; } return [n.name, n.ref, n.sheet || 'Workbook', n.hidden ? 'Hidden' : '', v instanceof CellError ? v.code : v, n.comment || ''] as Scalar[]; }); report(app, 'Names', ['Name', 'Refers to', 'Scope', 'Hidden', 'Value', 'Comment'], rows, { colW: [160, 220, 110, 60, 120, 200] }); } },
  {
    id: 'util.listDuplicates', label: 'List Duplicate Values with Counts', icon: 'content_copy', tab: U, group: G_INFO, edits: true, tip: 'A sheet listing each value that occurs more than once in the selection, how often, and where', origin: 'ASAP Utilities (Count duplicates)',
    run: (app) => { const s = app.sheet; const m = new Map<string, { text: string; n: number; where: string[] }>(); for (const x of selCells(app, { existingOnly: true })) { if (!hasContent(x.cell)) continue; const t = shownText(s, x.r, x.c); const k = t.toLowerCase(); const e = m.get(k) || m.set(k, { text: t, n: 0, where: [] }).get(k)!; e.n++; if (e.where.length < 20) e.where.push(cellName(x.r, x.c)); } const rows = [...m.values()].filter((e) => e.n > 1).sort((a, b) => b.n - a.n).map((e) => [e.text, e.n, e.where.join(', ')] as Scalar[]); if (!rows.length) { toast('No duplicates in the selection.'); return; } report(app, 'Duplicates', ['Value', 'Count', 'Cells'], rows, { colW: [220, 70, 400] }); },
  },
  {
    id: 'util.valueCounts', label: 'List Each Value with Its Count', icon: 'bar_chart', tab: U, group: G_INFO, edits: true, popular: true, tip: 'A frequency table of the selection: every distinct value, how many times it occurs, and its share', keywords: 'frequency count unique values tally', origin: 'ASAP Utilities / Kutools (Count by value)',
    run: (app) => { const s = app.sheet; const m = new Map<string, { text: string; n: number }>(); let total = 0; for (const x of selCells(app, { existingOnly: true })) { if (!hasContent(x.cell)) continue; const t = shownText(s, x.r, x.c); const k = t.toLowerCase(); (m.get(k) || m.set(k, { text: t, n: 0 }).get(k)!).n++; total++; } if (!total) { toast('No values in the selection.'); return; } const rows = [...m.values()].sort((a, b) => b.n - a.n || a.text.localeCompare(b.text)).map((e) => [e.text, e.n, { v: e.n / total, style: { numFmt: '0.0%' } }]); report(app, 'Counts', ['Value', 'Count', 'Share'], rows as any); },
  },
  { id: 'util.listFonts', label: 'List Fonts and Colors Used', icon: 'font_download', tab: U, group: G_INFO, edits: true, tip: 'Which fonts, sizes, fill and font colors the workbook uses, and in how many cells', origin: 'ASAP Utilities', run: (app) => { const fonts = new Map<string, number>(), fills = new Map<string, number>(), colors = new Map<string, number>(), fmts = new Map<string, number>(); for (const s of app.book.sheets) for (const cell of s.cells.values()) { const st = app.book.styles[cell.s ?? 0] || {}; const f = `${st.font || app.prefs.defaultFont || 'Default'} ${st.size || app.prefs.defaultSize || 11}pt`; fonts.set(f, (fonts.get(f) || 0) + 1); if (st.fill) fills.set(st.fill, (fills.get(st.fill) || 0) + 1); if (st.color) colors.set(st.color, (colors.get(st.color) || 0) + 1); if (st.numFmt) fmts.set(st.numFmt, (fmts.get(st.numFmt) || 0) + 1); } const rows: any[][] = []; for (const [k, n] of [...fonts].sort((a, b) => b[1] - a[1])) rows.push(['Font', k, n]); for (const [k, n] of [...fills].sort((a, b) => b[1] - a[1])) rows.push(['Fill color', { v: k, style: { fill: k } }, n]); for (const [k, n] of [...colors].sort((a, b) => b[1] - a[1])) rows.push(['Font color', { v: k, style: { color: k } }, n]); for (const [k, n] of [...fmts].sort((a, b) => b[1] - a[1])) rows.push(['Number format', k, n]); report(app, 'Fonts and Colors', ['Kind', 'Value', 'Cells'], rows); } },
  { id: 'util.listProperties', label: 'List Document Properties', icon: 'description', tab: U, group: G_INFO, edits: true, origin: 'ASAP Utilities', run: (app) => { const p = app.book.props; const st = workbookStats(app.book); const rows: Scalar[][] = [['Title', p.title], ['Subject', p.subject], ['Author', p.author], ['Manager', p.manager], ['Company', p.company], ['Category', p.category], ['Keywords', p.keywords], ['Status', p.status], ['Comments', p.comments], ['Created', new Date(p.created).toLocaleString()], ['Modified', new Date(p.modified).toLocaleString()], ['Last modified by', p.lastModifiedBy], ['File name', app.tab.fileName || '(unsaved)'], ['Location', app.tab.uri || ''], ['Sheets', st.sheets], ['Cells with data', st.cells], ['Formulas', st.formulas], ['Defined names', st.names], ['Charts', st.charts], ['Comments and notes', st.comments + st.notes], ...Object.entries(p.custom || {}).map(([k, v]) => [`Custom: ${k}`, v] as Scalar[])]; report(app, 'Properties', ['Property', 'Value'], rows, { colW: [160, 360] }); } },
  { id: 'util.listHyperlinks', label: 'List All Hyperlinks', icon: 'link', tab: U, group: G_INFO, edits: true, origin: 'ASAP Utilities', run: (app) => { const rows: Scalar[][] = []; for (const s of app.book.sheets) for (const [k, cell] of s.cells) if (cell.link) rows.push([s.name, cellName(keyRow(k), keyCol(k)), shownText(s, keyRow(k), keyCol(k)), cell.link, cell.linkTip || '']); if (!rows.length) { toast('The workbook has no hyperlinks.'); return; } report(app, 'Hyperlinks', ['Sheet', 'Cell', 'Text', 'Address', 'Tip'], rows, { colW: [100, 60, 200, 320, 160] }); } },
  { id: 'util.listMerged', label: 'List All Merged Cells', icon: 'call_merge', tab: U, group: G_INFO, edits: true, run: (app) => { const rows: Scalar[][] = []; for (const s of app.book.sheets) for (const m of s.merges) rows.push([s.name, rangeName(m), m.r2 - m.r1 + 1, m.c2 - m.c1 + 1, shownText(s, m.r1, m.c1)]); if (!rows.length) { toast('No merged cells.'); return; } report(app, 'Merged Cells', ['Sheet', 'Range', 'Rows', 'Columns', 'Text'], rows); } },
  {
    id: 'util.findProblems', label: 'Find Problems (Errors, Broken Links, Narrow Columns)', icon: 'troubleshoot', tab: U, group: G_INFO, edits: true, popular: true, tip: 'Check every sheet for error values, #REF! in formulas and names, links to missing sheets, columns too narrow for their numbers (####), and circular references', keywords: 'errors ref broken links narrow columns check', origin: 'ASAP Utilities (Detect and list problems)',
    run: (app) => {
      const rows: Scalar[][] = [];
      const names = new Set(app.book.sheets.map((s) => s.name.toLowerCase()));
      for (const s of app.book.sheets) {
        for (const [k, cell] of s.cells) {
          const addr = cellName(keyRow(k), keyCol(k));
          if (cell.v instanceof CellError) rows.push([s.name, addr, cell.v.code === '#REF!' ? 'Broken reference' : 'Error value', cell.v.code, cell.f != null ? '=' + cell.f : '']);
          if (cell.f != null) { for (const m of cell.f.replace(/"(?:[^"]|"")*"/g, '').matchAll(/(?:'([^']+)'|([A-Za-z_][\w. ]*))!/g)) { const nm = (m[1] || m[2] || '').toLowerCase(); if (nm && !names.has(nm) && !/^\[/.test(nm)) rows.push([s.name, addr, 'Link to a missing sheet', m[1] || m[2], '=' + cell.f]); } if (externalLinkIn(cell.f)) rows.push([s.name, addr, 'Link to another workbook', externalLinkIn(cell.f), '=' + cell.f]); }
        }
        for (const nc of narrowColumns(s)) rows.push([s.name, colName(nc.c), 'Column too narrow for its numbers (####)', nc.text, '']);
      }
      for (const n of app.book.names) if (/#REF!/i.test(n.ref)) rows.push(['(names)', n.name, 'Name with a broken reference', n.ref, '']);
      if (!rows.length) { toast('No problems found.'); return; }
      report(app, 'Problems', ['Sheet', 'Where', 'Problem', 'Details', 'Formula'], rows, { colW: [100, 80, 240, 220, 300] });
    },
  },
  { id: 'util.screenInfo', label: 'Screen and App Information', icon: 'devices', tab: U, group: G_INFO, noRecord: true, tip: 'Window size, zoom, pixel density, platform and language, for support questions', origin: 'ASAP Utilities (Show screen settings)', run: async (app) => { const { h } = await import('../ui/dom'); const { platformName } = await import('../platform/native'); const rows: [string, string][] = [['Window', `${window.innerWidth} × ${window.innerHeight} px`], ['Screen', `${screen.width} × ${screen.height} px`], ['Pixel ratio', String(window.devicePixelRatio)], ['Platform', platformName()], ['Browser', navigator.userAgent.slice(0, 120)], ['Language', navigator.language], ['Sheet zoom', `${Math.round(app.sheet.zoom * 100)}%`], ['UI scale', String(app.prefs.uiScale)], ['Theme', app.prefs.theme], ['Touch mode', app.prefs.touchMode ? 'On' : 'Off'], ['Commands', String(allCommands().length)]]; await alertDialog('Screen and App Information', h('table', { class: 'ql-info-table' }, ...rows.map(([k, v]) => h('tr', null, h('th', null, k), h('td', null, v))))); } },
  { id: 'util.menuInfo', label: 'Info & List', icon: 'info', tab: U, group: G_INFO, kind: 'menu', noRecord: true, run: () => undefined, menu: () => menuOf(['util.selectionInfo', 'util.workbookInfo', 'util.sheetsReport', 'util.listProperties', 'util.screenInfo', null, 'util.valueCounts', 'util.listDuplicates', 'util.countSumByColor', null, 'util.formulasReport', 'util.listNames', 'util.commentsReport', 'util.listHyperlinks', 'util.listMerged', 'util.objectsReport', 'util.listFonts', 'util.sheetIndex', null, 'util.findProblems', 'tools.dataProfile', 'review.accessibility']) },
]);

// ============================================================================================
// File & System

async function blobForBook(book: Workbook, format: string, sheetIndex?: number): Promise<Blob> { return exportFile(book, format, { sheetIndex: sheetIndex ?? book.activeSheet }); }

function bookWithSheets(app: App, sheets: Sheet[]): Workbook {
  const json = app.book.toJSON();
  const keep = new Set(sheets.map((s) => s.name));
  json.sheets = (json.sheets || []).filter((j: any) => keep.has(j.name));
  json.activeSheet = 0;
  return Workbook.fromJSON(json);
}

defTools([
  {
    id: 'util.saveBackup', label: 'Save a Backup Copy with a Timestamp', icon: 'backup', tab: U, group: G_FILE, noRecord: true, popular: true, tip: 'Save a copy of the workbook as "name 2026-10-26 1432.xlsx" (or the current format) without changing the open file', keywords: 'backup copy timestamp version', origin: 'ASAP Utilities (Save file and create backup)',
    run: async (app) => { const { saveAs } = await import('../platform/native'); const fmt = FORMATS.find((f) => f.id === (app.tab.format || 'xlsx') && f.canExport) || FORMATS[0]; const blob = await blobForBook(app.book, fmt.id); const name = `${fileStem(app.tab.fileName || app.tab.title || 'Workbook')} ${nowStamp()}.${fmt.ext[0]}`; const r = await saveAs(blob, name); if (r) toast(`Backup saved as ${r.name}`); },
  },
  { id: 'util.saveAndClose', label: 'Save and Close', icon: 'save', tab: U, group: G_FILE, noRecord: true, keys: ['Mod-Alt-w'], run: async (app) => { const ok = await app.exec('file.save'); if (ok !== false) app.exec('file.close'); } },
  { id: 'util.closeOthers', label: 'Close All Other Workbooks', icon: 'tab_close', tab: U, group: G_FILE, noRecord: true, origin: 'ASAP Utilities', run: async (app) => { const keep = app.tab; for (const d of [...app.docs]) { if (d === keep) continue; app.activate(d); await app.exec('file.close'); } app.activate(keep); } },
  { id: 'util.copyPath', label: 'Copy the File Name and Path', icon: 'content_copy', tab: U, group: G_FILE, noRecord: true, origin: 'ASAP Utilities / Kutools (Copy Full Path)', run: async (app) => { const t = app.tab.uri && !app.tab.uri.startsWith('fsa:') ? app.tab.uri : app.tab.fileName || app.tab.title; await copyText(t, 'Path'); } },
  { id: 'util.insertFileName', label: 'Insert the File Name in the Active Cell', icon: 'description', tab: U, group: G_FILE, edits: true, kind: 'split', origin: 'ASAP Utilities / Kutools (Insert Workbook Information)', run: (app, arg?: 'name' | 'path' | 'sheet' | 'user' | 'date') => { const { r, c } = activeCell(app); const v = arg === 'path' ? (app.tab.uri && !app.tab.uri.startsWith('fsa:') ? app.tab.uri : app.tab.fileName || app.tab.title) : arg === 'sheet' ? app.sheet.name : arg === 'user' ? app.prefs.author || '' : arg === 'date' ? new Date().toLocaleString() : app.tab.fileName || app.tab.title; app.book.transact('Insert Info', () => writeOut(app.sheet, r, c, v)); }, menu: () => [{ id: 'util.insertFileName', arg: 'name', label: 'File name' }, { id: 'util.insertFileName', arg: 'path', label: 'Full path' }, { id: 'util.insertFileName', arg: 'sheet', label: 'Sheet name' }, { id: 'util.insertFileName', arg: 'user', label: 'User name' }, { id: 'util.insertFileName', arg: 'date', label: 'Current date and time' }] },
  { id: 'util.fileNameHeader', label: 'File Name and Date in the Header and Footer', icon: 'vertical_split', tab: U, group: G_FILE, edits: true, origin: 'ASAP Utilities', run: async (app) => { const { setPrint } = await import('./pagelayout'); setPrint(app, 'Header and Footer', (p) => { p.header = { left: '&F', center: '', right: '&D' }; p.footer = { left: '&A', center: 'Page &P of &N', right: '&T' }; }); toast('Header: file name and date; footer: sheet name, page numbers and time'); } },
  {
    id: 'util.removeAllMacros', label: 'Remove All Macros from the Workbook', icon: 'code_off', tab: U, group: G_FILE, edits: true, tip: 'Delete the JavaScript macro modules stored in the workbook (and any preserved VBA project)', origin: 'ASAP Utilities',
    run: async (app, arg) => { const n = app.book.modules.length + ((app.book as any).vbaProject ? 1 : 0); if (!n) { toast('The workbook has no macros.'); return; } if (arg === undefined && !(await confirmDialog('Remove All Macros', `Remove ${n} macro module${n === 1 ? '' : 's'} from the workbook?`, 'Remove'))) return; app.book.transact('Remove Macros', () => { app.book.recordBook(); app.book.modules = []; delete (app.book as any).vbaProject; }); toast('Macros removed'); },
  },
  {
    id: 'util.cleanWorkbook', label: 'Clean Up the Workbook (Reduce File Size)', icon: 'cleaning_services', tab: U, group: G_FILE, edits: true, popular: true, tip: 'Trim the used range of every sheet, drop unused styles and names with errors, and report what was removed', keywords: 'reduce file size clean excess formatting bloat', origin: 'ASAP Utilities / Inquire (Clean Excess Cell Formatting)',
    run: async (app) => { const { resetLastCell } = await import('./utilities'); let cells = 0; const names = app.book.names.filter((n) => /#REF!/i.test(n.ref)).length; const styles = app.book.cellStyles.filter((cs) => !cs.builtin && cs.name !== 'Normal').length; app.book.transact('Clean Up Workbook', () => { app.book.recordBook(); for (const s of app.book.sheets) cells += resetLastCell(s).cells; app.book.names = app.book.names.filter((n) => !/#REF!/i.test(n.ref)); app.book.cellStyles = app.book.cellStyles.filter((cs) => cs.builtin || cs.name === 'Normal'); }); app.grid.updateExtent(); app.grid.layout(); await alertDialog('Clean Up the Workbook', `Removed ${plural(cells, 'empty formatted cell')}, ${plural(names, 'broken name')} and ${plural(styles, 'custom cell style')}. Save the workbook to shrink the file.`); },
  },
  {
    id: 'util.listFolder', label: 'List the Files in a Folder…', icon: 'folder_open', tab: U, group: G_FILE, edits: true, popular: true, tip: 'Pick a folder (or files) and get a sheet with each file\'s name, extension, size and modification date', keywords: 'directory listing file names folder', origin: 'ASAP Utilities / Kutools (Filename List)',
    run: async (app) => {
      const rows: Scalar[][] = [];
      const anyWin = window as any;
      if (anyWin.showDirectoryPicker) {
        try {
          const dir = await anyWin.showDirectoryPicker();
          const walk = async (handle: any, path: string, depth: number) => { for await (const [name, h] of handle.entries()) { if (h.kind === 'file') { const f = await h.getFile(); rows.push([path + name, name.replace(/^.*\./, name.includes('.') ? '' : ''), f.size, { v: f.lastModified / 86400000 + 25569, style: { numFmt: 'yyyy-mm-dd hh:mm' } } as any, path || '.']); } else if (depth < 3) await walk(h, path + name + '/', depth + 1); } };
          await walk(dir, '', 0);
        } catch { return; }
      } else {
        const files = await pickFiles('*/*', true);
        for (const f of files) rows.push([f.name, f.name.includes('.') ? f.name.replace(/^.*\./, '') : '', f.size, { v: f.lastModified / 86400000 + 25569, style: { numFmt: 'yyyy-mm-dd hh:mm' } } as any, (f as any).webkitRelativePath || '']);
      }
      if (!rows.length) { toast('No files listed.'); return; }
      rows.sort((a, b) => String(a[0]).localeCompare(String(b[0])));
      report(app, 'Files', ['File', 'Extension', 'Size (bytes)', 'Modified', 'Folder'], rows, { colW: [260, 80, 100, 130, 200] });
    },
  },
  { id: 'util.openContainingFolder', label: 'Open the Folder That Contains This File', icon: 'folder', tab: U, group: G_FILE, noRecord: true, origin: 'ASAP Utilities / Kutools (Open Containing Folder)', run: async (app) => { const uri = app.tab.uri; if (!uri || uri.startsWith('fsa:') || uri.startsWith('ios:')) { toast('The file was opened through a file picker; its folder cannot be opened from here.'); return; } const { openExternal } = await import('../platform/native'); const folder = uri.replace(/[\\/][^\\/]*$/, ''); openExternal(folder.startsWith('file:') ? folder : 'file://' + folder); } },
  { id: 'util.exportSheetsFiles', label: 'Save Each Sheet as a Separate File…', icon: 'folder_zip', tab: U, group: G_FILE, noRecord: true, popular: true, tip: 'Export the chosen sheets one by one as .xlsx, .csv, .ods, .html, .json or .pdf files', keywords: 'split workbook sheets files export each', origin: 'ASAP Utilities / Kutools (Split Workbook)',
    run: async (app, arg?: { names?: string[]; format?: string }) => {
      const exportable = FORMATS.filter((f) => f.canExport && ['xlsx', 'csv', 'tsv', 'txt', 'ods', 'html', 'json', 'qgw'].includes(f.id));
      const pick = arg ? { sheets: sheetsByNames(app.book, arg.names), values: arg } : await pickSheets(app, 'Save Each Sheet as a File', { extra: [{ key: 'format', label: 'Format', type: 'select', value: 'xlsx', options: exportable.map((f) => ({ value: f.id, label: f.label })) }] });
      if (!pick || !pick.sheets.length) return;
      const fmt = FORMATS.find((f) => f.id === ((pick.values as any).format || 'xlsx'))!;
      const { saveAs } = await import('../platform/native');
      let n = 0;
      for (const s of pick.sheets) { const nb = bookWithSheets(app, [s]); const blob = await exportFile(nb, fmt.id, { sheetIndex: 0 }); const r = await saveAs(blob, `${s.name}.${fmt.ext[0]}`); if (!r) break; n++; }
      toast(`${plural(n, 'file')} saved`);
    },
  },
  { id: 'util.menuFile', label: 'File & System', icon: 'folder', tab: U, group: G_FILE, kind: 'menu', noRecord: true, run: () => undefined, menu: () => menuOf(['util.saveBackup', 'util.saveAndClose', 'util.closeOthers', 'file.saveAll', 'file.versions', null, 'util.copyPath', 'util.insertFileName', 'util.fileNameHeader', 'util.openContainingFolder', 'util.listFolder', null, 'util.cleanWorkbook', 'util.removeAllMacros', 'util.breakExternalLinks', 'file.inspect', 'file.properties', null, 'util.exportSheetsFiles', 'util.sheetsToNewWorkbook']) },
]);

// ============================================================================================
// Import

async function importMany(app: App, files: File[], mode: 'sheets' | 'stack' | 'books', opts: { nameColumn?: boolean } = {}): Promise<void> {
  if (!files.length) return;
  const book = app.book;
  if (mode === 'books') { for (const f of files) { const res = await importFile(await f.arrayBuffer(), f.name).catch((e) => { toast(`${f.name}: ${(e as Error).message}`); return null; }); if (res) app.createTab(res.book, { title: fileStem(f.name), fileName: f.name }); } return; }
  if (structureLocked(app)) return;
  const results: { name: string; book: Workbook }[] = [];
  for (const f of files) { try { const res = await importFile(await f.arrayBuffer(), f.name); results.push({ name: f.name, book: res.book }); } catch (e) { toast(`${f.name}: ${(e as Error).message}`); } }
  if (!results.length) return;
  book.transact('Import Files', () => {
    book.recordBook();
    if (mode === 'sheets') {
      let at = book.sheets.length - 1;
      for (const r of results) for (const sh of r.book.sheets) { const copy = new Sheet(book, book.uniqueSheetName(fileStem(r.name).slice(0, 25))); copy.loadJSON(sh.toJSON()); copy.name = r.book.sheets.length === 1 ? uniqueName(book, fileStem(r.name)) : uniqueName(book, `${fileStem(r.name)} ${sh.name}`); book.sheets.splice(++at, 0, copy); }
    } else {
      const dest = new Sheet(book, book.uniqueSheetName('Combined'));
      book.sheets.push(dest);
      let row = 0; let header: string[] | null = null;
      for (const r of results) for (const sh of r.book.sheets) {
        const u = sh.usedRange(false); if (!u) continue;
        const rows = textRows(sh, u, false);
        const first = rows[0] || [];
        const isHeader = first.every((t) => t === '' || isNaN(Number(t)));
        if (header == null && isHeader) { header = first; if (opts.nameColumn) header = ['File', ...header]; header.forEach((t, j) => writeOut(dest, 0, j, { v: t, style: HEAD })); row = 1; }
        const body = isHeader && header ? rows.slice(1) : rows;
        for (const line of body) { const vals: Scalar[] = line.map((t, j) => { const v = sh.value(u.r1 + rows.indexOf(line), u.c1 + j); return v; }); const outRow = opts.nameColumn ? [r.name, ...vals] : vals; outRow.forEach((v, j) => { if (v != null && v !== '') writeOut(dest, row, j, v as any); }); row++; }
      }
      dest.freeze = { rows: header ? 1 : 0, cols: 0 };
      book.activeSheet = book.sheets.length - 1;
    }
  });
  app.ui.rebuild(); app.exec('sheet.activate', book.activeSheet); app.grid.updateExtent();
  toast(`${plural(results.length, 'file')} imported`);
}

function uniqueName(book: Workbook, base: string): string { const clean = base.replace(/[\\/?*[\]:]/g, ' ').trim().slice(0, 31) || 'Sheet'; if (!book.sheetByName(clean)) return clean; let i = 2; while (book.sheetByName(`${clean.slice(0, 27)} (${i})`)) i++; return `${clean.slice(0, 27)} (${i})`; }

const IMPORT_ACCEPT = '.xlsx,.xlsm,.xls,.ods,.fods,.csv,.tsv,.txt,.dbf,.json,.html,.htm,.xml,.slk,.dif,.qgw,.wk1,.wks,.wk3,.wk4,.123,.wb1,.wb2,.wb3,.wq1';

defTools([
  { id: 'util.importAsSheets', label: 'Import Files as New Sheets…', icon: 'library_add', tab: U, group: G_IMPORT, edits: true, popular: true, tip: 'Pick several workbooks, CSV, text or dBASE files; each one (and each of its sheets) becomes a sheet in this workbook', keywords: 'merge files import many csv dbf workbooks', origin: 'ASAP Utilities (Import several files) / Kutools (Combine)', run: async (app) => importMany(app, await pickFiles(IMPORT_ACCEPT, true), 'sheets') },
  { id: 'util.importStacked', label: 'Import and Stack Files into One Sheet…', icon: 'view_agenda', tab: U, group: G_IMPORT, edits: true, popular: true, tip: 'Append the rows of several files (same layout) under each other on one sheet, keeping one header and adding a File column', keywords: 'append combine csv files one sheet', origin: 'ASAP Utilities / Ablebits (Combine Sheets) / Power Tools', run: async (app) => importMany(app, await pickFiles(IMPORT_ACCEPT, true), 'stack', { nameColumn: true }) },
  { id: 'util.openMany', label: 'Open Several Files at Once…', icon: 'file_open', tab: U, group: G_IMPORT, noRecord: true, run: async (app) => importMany(app, await pickFiles(IMPORT_ACCEPT, true), 'books') },
  { id: 'util.importDbf', label: 'Import a dBASE File (.dbf)…', icon: 'dataset', tab: U, group: G_IMPORT, edits: true, origin: 'ASAP Utilities (Import dBASE)', run: async (app) => importMany(app, await pickFiles('.dbf', true), 'sheets') },
  { id: 'util.importDelimited', label: 'Import a Delimited Text File with a Custom Separator…', icon: 'text_snippet', tab: U, group: G_IMPORT, edits: true, tip: 'Pipe, semicolon, tab or any separator', origin: 'ASAP Utilities', run: async (app, arg?: { sep: string }) => { const files = await pickFiles('.txt,.csv,.tsv,.dat,.log,*/*', false); if (!files.length) return; const o = await argOrAsk(arg, async () => { const sep = await promptDialog('Separator', 'Field separator (\\t for tab)', '|'); return sep == null ? null : { sep }; }); if (!o) return; const sep = o.sep === '\\t' ? '\t' : o.sep; const { parseCsv } = await import('../io/csv'); const text = await files[0].text(); const rows = parseCsv(text, { delimiter: sep } as any); addSheetWithRows(app, fileStem(files[0].name), rows.map((r) => r.map((t) => { const n = Number(t); return t !== '' && !isNaN(n) && /^-?\d/.test(t) ? n : t; }))); } },
  { id: 'util.pasteAsNewSheet', label: 'Paste the Clipboard as a New Sheet', icon: 'content_paste', tab: U, group: G_IMPORT, edits: true, tip: 'Tab- or comma-separated text from the clipboard becomes a new sheet', run: async (app) => { const { readClipboardText } = await import('../platform/native'); const text = await readClipboardText(); if (!text) { toast('The clipboard has no text.'); return; } const { parseCsv, detectDelimiter } = await import('../io/csv'); const rows = parseCsv(text, { delimiter: detectDelimiter(text) } as any); addSheetWithRows(app, 'Pasted', rows.map((r) => r.map((t) => { const n = Number(t); return t !== '' && !isNaN(n) && /^-?\d/.test(t) ? n : t; }))); } },
  { id: 'util.menuImport', label: 'Import', icon: 'download', tab: U, group: G_IMPORT, kind: 'menu', noRecord: true, run: () => undefined, menu: () => menuOf(['util.importAsSheets', 'util.importStacked', 'util.openMany', 'util.importDbf', 'util.importDelimited', 'util.pasteAsNewSheet', 'util.listFolder', null, 'data.fromText', 'data.fromWeb', 'data.fromWorkbook', 'data.fromClipboard', 'dev.importJson', 'file.openRepair']) },
]);

// ============================================================================================
// Export

function selectionOrSheet(app: App): Range | null { const s = app.sheet; if (singleCell(app)) return s.usedRange(false); return clipToUsed(s, activeRange(app)); }

async function exportSelection(app: App, kind: 'csv' | 'tsv' | 'custom' | 'json' | 'html' | 'md' | 'sql' | 'xml' | 'latex' | 'fixed', sep?: string): Promise<void> {
  const s = app.sheet; const rg = selectionOrSheet(app); if (!rg) { toast('Nothing to export.'); return; }
  const rows = textRows(s, rg); const vals = valueRows(s, rg);
  const header = rows[0] || [];
  const stem = `${fileStem(app.tab.fileName || app.tab.title || 'Export')}-${s.name}`;
  let text = ''; let ext: string = kind; let mime = 'text/plain';
  if (kind === 'csv' || kind === 'tsv' || kind === 'custom') { const d = kind === 'csv' ? ',' : kind === 'tsv' ? '\t' : (sep || ';'); text = rows.map((r) => r.map((t) => toCsvField(t, d)).join(d)).join('\r\n'); ext = kind === 'custom' ? 'txt' : kind; mime = kind === 'csv' ? 'text/csv' : 'text/plain'; }
  else if (kind === 'json') { const body = vals.slice(1).map((r) => Object.fromEntries(r.map((v, j) => [header[j] || colName(rg.c1 + j), v instanceof CellError ? v.code : v]))); text = JSON.stringify(body, null, 2); mime = 'application/json'; }
  else if (kind === 'html') { text = `<!DOCTYPE html>\n<html><head><meta charset="utf-8"><title>${escapeHtml(s.name)}</title></head><body>\n${htmlTable(s, rg, { styles: true, header: true })}\n</body></html>`; mime = 'text/html'; }
  else if (kind === 'md') { text = sheetToMarkdown(s, rg); }
  else if (kind === 'sql') { const table = s.name.replace(/\W+/g, '_'); const cols = header.map((h) => (h || 'col').replace(/\W+/g, '_')); text = `CREATE TABLE ${table} (${cols.map((c, j) => `${c} ${vals.slice(1).every((r) => typeof r[j] === 'number' || r[j] == null) ? 'NUMERIC' : 'TEXT'}`).join(', ')});\n` + vals.slice(1).map((r) => `INSERT INTO ${table} (${cols.join(', ')}) VALUES (${r.map((v) => (v == null || v === '' ? 'NULL' : typeof v === 'number' ? String(v) : `'${String(v instanceof CellError ? v.code : v).replace(/'/g, "''")}'`)).join(', ')});`).join('\n'); }
  else if (kind === 'xml') { text = `<?xml version="1.0" encoding="UTF-8"?>\n<rows>\n` + vals.slice(1).map((r) => '  <row>' + r.map((v, j) => { const tag = (header[j] || `c${j + 1}`).replace(/[^\w-]+/g, '_').replace(/^(\d)/, '_$1'); return `<${tag}>${escapeHtml(v == null ? '' : String(v instanceof CellError ? v.code : v))}</${tag}>`; }).join('') + '</row>').join('\n') + '\n</rows>'; mime = 'application/xml'; }
  else if (kind === 'latex') { const n = header.length; text = `\\begin{tabular}{${'l'.repeat(n)}}\n\\hline\n` + rows.map((r, i) => r.map((t) => t.replace(/([&%$#_{}])/g, '\\$1')).join(' & ') + ' \\\\' + (i === 0 ? '\n\\hline' : '')).join('\n') + '\n\\hline\n\\end{tabular}'; ext = 'tex'; }
  else if (kind === 'fixed') { const widths = header.map((_h, j) => Math.max(...rows.map((r) => (r[j] || '').length))); text = rows.map((r) => r.map((t, j) => (typeof vals[rows.indexOf(r)]?.[j] === 'number' ? t.padStart(widths[j]) : t.padEnd(widths[j]))).join('  ')).join('\n'); ext = 'txt'; }
  await saveTextFile(text, `${stem}.${ext}`, mime);
}

defTools([
  { id: 'util.exportSelectionCsv', label: 'Export Selection as CSV', icon: 'csv', tab: U, group: G_EXPORT, noRecord: true, popular: true, tip: 'Only the selected cells (or the sheet) as a CSV file', origin: 'ASAP Utilities (Export selection as CSV)', run: (app) => exportSelection(app, 'csv') },
  { id: 'util.exportSelectionTsv', label: 'Export Selection as Tab-Separated Text', icon: 'text_snippet', tab: U, group: G_EXPORT, noRecord: true, run: (app) => exportSelection(app, 'tsv') },
  { id: 'util.exportSelectionCustom', label: 'Export Selection with a Custom Separator…', icon: 'text_snippet', tab: U, group: G_EXPORT, noRecord: true, origin: 'ASAP Utilities', run: async (app, arg?: string) => { const sep = await argOrAsk(arg, () => promptDialog('Export with a Custom Separator', 'Separator (\\t for tab)', ';')); if (sep == null) return; await exportSelection(app, 'custom', sep === '\\t' ? '\t' : sep); } },
  { id: 'util.exportSelectionJson', label: 'Export Selection as JSON Records', icon: 'data_object', tab: U, group: G_EXPORT, noRecord: true, tip: 'An array of objects, one per row, keyed by the header row', origin: 'ASAP Utilities', run: (app) => exportSelection(app, 'json') },
  { id: 'util.exportSelectionHtml', label: 'Export Selection as an HTML Page', icon: 'html', tab: U, group: G_EXPORT, noRecord: true, run: (app) => exportSelection(app, 'html') },
  { id: 'util.exportSelectionMd', label: 'Export Selection as Markdown', icon: 'article', tab: U, group: G_EXPORT, noRecord: true, run: (app) => exportSelection(app, 'md') },
  { id: 'util.exportSelectionSql', label: 'Export Selection as SQL Statements', icon: 'database', tab: U, group: G_EXPORT, noRecord: true, tip: 'CREATE TABLE and INSERT statements named after the sheet and header row', run: (app) => exportSelection(app, 'sql') },
  { id: 'util.exportSelectionXml', label: 'Export Selection as XML', icon: 'code', tab: U, group: G_EXPORT, noRecord: true, run: (app) => exportSelection(app, 'xml') },
  { id: 'util.exportSelectionLatex', label: 'Export Selection as a LaTeX Table', icon: 'functions', tab: U, group: G_EXPORT, noRecord: true, origin: 'Publishing workflows', run: (app) => exportSelection(app, 'latex') },
  { id: 'util.exportSelectionFixed', label: 'Export Selection as Fixed-Width Text', icon: 'view_column', tab: U, group: G_EXPORT, noRecord: true, tip: 'Columns padded to equal widths, for plain-text reports and listings', run: (app) => exportSelection(app, 'fixed') },
  { id: 'util.copyValuesOnly', label: 'Copy Values Only (Plain Text)', icon: 'content_copy', tab: U, group: G_EXPORT, noRecord: true, keys: ['Mod-Alt-Shift-c'], tip: 'The selection as tab-separated values on the clipboard, without formatting or formulas', origin: 'ASAP Utilities (Copy clean values)', run: async (app) => { const s = app.sheet; await copyText(rangeToTsv(s, clipToUsed(s, activeRange(app))), 'Values'); } },
  { id: 'util.exportChartsImages', label: 'Save Each Chart as an Image', icon: 'image', tab: U, group: G_EXPORT, noRecord: true, tip: 'Every chart on the sheet as a PNG file', origin: 'ASAP Utilities / Kutools (Export Graphics)', run: async (app) => { const s = app.sheet; if (!s.charts.length) { toast('No charts on this sheet.'); return; } for (const ch of s.charts) { app.grid.objects.select(ch.id); const r = await app.exec('chart.saveImage'); if (r === false) break; } } },
  { id: 'util.exportSheetPdf', label: 'Export the Sheet as PDF', icon: 'picture_as_pdf', tab: U, group: G_EXPORT, noRecord: true, run: (app) => app.exec('file.exportPdf') },
  { id: 'util.menuExport', label: 'Export', icon: 'upload', tab: U, group: G_EXPORT, kind: 'menu', noRecord: true, run: () => undefined, menu: () => menuOf(['util.exportSelectionCsv', 'util.exportSelectionTsv', 'util.exportSelectionCustom', 'util.exportSelectionFixed', 'util.exportSelectionJson', 'util.exportSelectionXml', 'util.exportSelectionSql', 'util.exportSelectionHtml', 'util.exportSelectionMd', 'util.exportSelectionLatex', null, 'util.copyValuesOnly', 'util.copyAsHtml', 'util.copyAsMarkdown', 'edit.copyAsPicture', 'util.exportChartsImages', null, 'util.exportSheetPdf', 'util.exportSheetsFiles', 'sheet.exportCsv', 'file.exportAs']) },
]);

// ============================================================================================
// Start, My Tools & Shortcuts, Options

const CHEER = ['Rows are just columns that lay down for a rest.', 'A spreadsheet a day keeps the auditor away.', 'Every blank cell is a decision not yet made.', '#REF! is the sound of a cell that lost its way.', 'Freeze panes: because headers deserve to stay in view.', 'Pivot tables: turning questions into answers since 1994.', 'The fill handle is the most productive square inch in computing.', 'Remember to save. Then save again.', 'Behind every clean table is a Trim Spaces command.', 'Circular references: when a cell depends on itself for validation.'];

async function myTools(app: App): Promise<string[]> { return (await kvGet<string[]>('util:myTools')) || app.prefs.pinnedCommands || []; }
async function setMyTools(app: App, ids: string[]): Promise<void> { await kvSet('util:myTools', ids); app.prefs.pinnedCommands = ids; app.savePrefs(); }
let myToolsCache: string[] = [];
void kvGet<string[]>('util:myTools').then((v) => { if (v) myToolsCache = v; });

const UTIL_OPTIONS = { confirmDeletes: true, reportSheets: true };
void kvGet<typeof UTIL_OPTIONS>('util:options').then((v) => { if (v) Object.assign(UTIL_OPTIONS, v); });

defTools([
  { id: 'util.calculator', label: 'Calculator…', icon: 'calculate', tab: U, group: G_START, noRecord: true, tip: 'Evaluate an expression with the formula engine and insert or copy the result', origin: 'ASAP Utilities (Start ▸ Calculator) / Kutools', run: async (app) => { const r = await openDialog({ title: 'Calculator', fields: [{ key: 'expr', label: 'Expression (worksheet functions allowed)', type: 'text', value: '', autofocus: true, placeholder: '=ROUND(17.5*1.08,2)' }], ok: 'Insert into Cell', buttons: [{ label: 'Copy', onClick: async (v, ctl) => { const t = calc(app, String(v.expr || '')); await copyText(t, 'Result'); ctl.close(); } }] }); if (!r) return; const t = calc(app, String(r.expr || '')); const { r: rr, c } = activeCell(app); app.book.transact('Calculator', () => writeOut(app.sheet, rr, c, numOf(t) ?? t)); } },
  { id: 'util.notepad', label: 'Scratchpad…', icon: 'edit_note', tab: U, group: G_START, noRecord: true, tip: 'A plain text pad that is kept between sessions (notes, snippets, to-do items)', origin: 'ASAP Utilities (Start ▸ Notepad)', run: async () => { const cur = (await kvGet<string>('util:scratch')) || ''; const r = await openDialog({ title: 'Scratchpad', wide: true, fields: [{ key: 't', label: '', type: 'textarea', rows: 14, value: cur, autofocus: true }], ok: 'Keep' }); if (r) await kvSet('util:scratch', String(r.t || '')); } },
  { id: 'util.clipboardView', label: 'View the Clipboard', icon: 'content_paste_search', tab: U, group: G_START, noRecord: true, origin: 'ASAP Utilities (Start ▸ Clipboard viewer)', run: (app) => app.ui.openPane('clipboard') },
  { id: 'util.cheer', label: 'A Random Spreadsheet Saying', icon: 'sentiment_satisfied', tab: U, group: G_START, noRecord: true, origin: 'ASAP Utilities (Start ▸ random funny messages)', run: () => { toast(CHEER[Math.floor(Math.random() * CHEER.length)], { ms: 6000 }); } },
  { id: 'util.menuStart', label: 'Start', icon: 'rocket_launch', tab: U, group: G_START, kind: 'menu', noRecord: true, run: () => undefined, menu: () => menuOf(['util.calculator', 'util.notepad', 'util.clipboardView', 'util.openContainingFolder', 'util.listFolder', 'util.cheer']) },

  {
    id: 'util.myTools', label: 'My Tools', icon: 'star', tab: U, group: G_MY, kind: 'menu', noRecord: true, popular: true, tip: 'Your favorite tools in one menu; add any command with Add to My Tools, and give them shortcuts with Customize Keyboard', origin: 'ASAP Utilities (My Tools & Shortcuts)',
    run: (app, id?: string) => { if (id) app.exec(id); },
    menu: () => { const items = myToolsCache.map((id) => getCommand(id)).filter((c): c is NonNullable<typeof c> => !!c).map((c): MenuEntry => ({ id: 'util.myTools', arg: c.id, label: c.label, icon: c.icon })); return [...(items.length ? items : [{ label: '(empty: use Add to My Tools)', id: 'util.addToMyTools' }]), { separator: true }, { id: 'util.addToMyTools', label: 'Add the Last Used Tool…', icon: 'add' }, { id: 'util.editMyTools', label: 'Edit My Tools…', icon: 'edit' }, { id: 'file.customizeKeyboard', label: 'Customize Keyboard…', icon: 'keyboard' }]; },
  },
  {
    id: 'util.addToMyTools', label: 'Add to My Tools…', icon: 'star', tab: U, group: G_MY, noRecord: true, tip: 'Add the last tool you used, or any command, to the My Tools menu',
    run: async (app, arg?: string) => { const list = await myTools(app); const id = await argOrAsk(arg, async () => { const q = await openDialog({ title: 'Add to My Tools', fields: [{ key: 'q', label: 'Search for a command', type: 'text', value: toolLog.last?.label || '', autofocus: true }, { key: 'pick', label: 'Matches', type: 'list', rows: 8, options: (toolLog.last ? [toolLog.last.label] : []).concat(searchCommands(toolLog.last?.label || 'a', 30).map((c) => c.label)) }], onChange: (v, ctl, k) => { if (k === 'q') { const { setListOptions } = require_dialog(); setListOptions(ctl, 'pick', searchCommands(String(v.q || ''), 30).map((c) => c.label)); } } }); if (!q) return null; const label = Array.isArray(q.pick) ? q.pick[0] : q.pick; const c = allCommands().find((x) => x.label === label) || searchCommands(String(q.q || ''), 1)[0]; return c?.id || null; }); if (!id || !getCommand(id)) return; if (list.includes(id)) { toast('Already in My Tools.'); return; } list.push(id); await setMyTools(app, list); myToolsCache = list; toast(`${getCommand(id)!.label} added to My Tools`); },
  },
  { id: 'util.editMyTools', label: 'Edit My Tools…', icon: 'edit', tab: U, group: G_MY, noRecord: true, run: async (app) => { const list = await myTools(app); const r = await openDialog({ title: 'Edit My Tools', help: 'One command id per line, in menu order (Developer ▸ Copy Command ID List shows every id).', fields: [{ key: 'ids', label: 'Commands', type: 'textarea', rows: 10, value: list.join('\n') }] }); if (!r) return; const ids = String(r.ids || '').split(/\r?\n/).map((x) => x.trim()).filter((x) => x && getCommand(x)); await setMyTools(app, ids); myToolsCache = ids; toast(`${ids.length} tools in My Tools`); } },
  { id: 'util.findTool', label: 'Find and Run a Tool…', icon: 'search', tab: U, group: G_MY, noRecord: true, popular: true, keys: ['Mod-Alt-f'], tip: 'Search every command by name and description and run it', origin: 'ASAP Utilities (Find and run a tool)', run: (app) => app.ui.commandPalette() },
  { id: 'util.lastTool', label: 'Start Last Tool Again', icon: 'replay', tab: U, group: G_MY, noRecord: true, popular: true, keys: ['Mod-Alt-r'], tip: 'Run the last Utilities, Plus or Extras tool again with the same settings', origin: 'ASAP Utilities (Start last tool again)', run: (app) => { const l = toolLog.last; if (!l) { toast('No tool has been used yet.'); return; } app.exec(l.id, l.arg); } },
  { id: 'util.usageReport', label: 'Tool Usage Report', icon: 'insights', tab: U, group: G_MY, edits: true, tip: 'Which tools you have used this session, how often, and a rough estimate of the clicks saved', origin: 'ASAP Utilities (Time and money saved report)', run: (app) => { const rows = Object.entries(toolLog.counts).sort((a, b) => b[1] - a[1]).map(([id, n]) => [getCommand(id)?.label || id, n, n * 12, `${Math.round(n * 12 * 2 / 60)} min`] as Scalar[]); if (!rows.length) { toast('No tools used yet this session.'); return; } report(app, 'Tool Usage', ['Tool', 'Runs', 'Clicks saved (est.)', 'Time saved (est.)'], rows); } },
  { id: 'util.options', label: 'Utilities Options…', icon: 'settings', tab: U, group: G_MY, noRecord: true, origin: 'ASAP Utilities Options', run: async (app) => { const r = await openDialog({ title: 'Utilities Options', fields: [{ key: 'confirmDeletes', label: 'Ask before tools delete rows, columns, sheets or comments', type: 'checkbox', value: UTIL_OPTIONS.confirmDeletes }, { key: 'reportSheets', label: 'Put reports on new sheets (otherwise show them in a window where possible)', type: 'checkbox', value: UTIL_OPTIONS.reportSheets }, { key: 'info', type: 'info', html: '<p class="ql-help">Shortcuts for every tool: File ▸ Options ▸ Customize Keyboard. Tools on the Quick Access Bar: long-press any button.</p>' }] }); if (!r) return; Object.assign(UTIL_OPTIONS, { confirmDeletes: !!r.confirmDeletes, reportSheets: !!r.reportSheets }); await kvSet('util:options', UTIL_OPTIONS); void app; } },
  { id: 'util.about', label: 'About the Utilities', icon: 'help', tab: U, group: G_MY, noRecord: true, run: async () => { const n = allCommands().filter((c) => c.tab === 'Utilities' || c.tab === 'Plus' || c.tab === 'Extras').length; await alertDialog('About the Utilities, Plus and Extras tabs', `${n} tools in three tabs. Utilities follows the categories of the ASAP Utilities add-in; Plus gathers tools known from Kutools, Ablebits Ultimate Suite, Power Tools, XLTools and Inquire; Extras holds tools of Anne's Grid's own. All are original re-implementations: see docs/PLUGINS.md. Every tool can go on the Quick Access Bar, a custom tab, a keyboard shortcut or into a macro.`); } },
  { id: 'util.tipOfDay', label: 'Tip of the Day', icon: 'lightbulb', tab: U, group: G_MY, noRecord: true, run: async () => { const tips = allCommands().filter((c) => (c.tab === 'Utilities' || c.tab === 'Plus' || c.tab === 'Extras') && c.tip); const day = Math.floor(Date.now() / 86400000); const c = tips[day % tips.length]; await alertDialog(`Tip of the Day: ${c.label}`, `${c.tip}\n\nFind it on the ${c.tab} tab (${c.group}), or search for it with Find and Run a Tool.`); } },
]);

function calc(app: App, expr: string): string { const f = expr.replace(/^=/, '').trim(); if (!f) return ''; try { const v = app.engine.evaluateScalar(f, app.sheet, app.sheet.sel.r, app.sheet.sel.c); return v instanceof CellError ? v.code : v == null ? '' : String(v); } catch (e) { return (e as Error).message; } }
function require_dialog(): { setListOptions: typeof import('../ui/dialog').setListOptions } { return dialogModule; }
import * as dialogModule from '../ui/dialog';

void cellKey; void selBlock; void sheetProtected; void emptySheet; void textOf; void checkWritable; void canChange; void changeCells; void writeBlock; void choiceDialog;
export type UtilCell = Cell;
