/**
 * Tables: Create Table / Format as Table, the Table Design contextual tab (name, resize, style options,
 * totals row, convert to range) and table row/column editing.
 */
import { def, type MenuEntry } from './registry';
import type { App } from '../app';
import type { TableDef, CellStyle, Cell } from '../model/types';
import type { Sheet, Workbook, ThemeDef } from '../model/workbook';
import { newId } from '../model/workbook';
import { type Range, rangeName, parseRange, cellKey, keyRow, keyCol, intersects } from '../model/address';
import { insertCellsShift, resolveRangeText, rangeRefText } from '../model/structure';
import { TABLE_STYLE_NAMES, tableLook, tableCellLook, DEFAULT_TABLE_OPTS } from '../model/tableStyles';
import { activeRange, clipToUsed, currentRegion, hasHeaderRow, isWholeCols, isWholeRows } from './ops';
import { openDialog, toast, promptDialog, confirmDialog } from '../ui/dialog';
import { formatValue } from '../model/numfmt';
import { galleryDialog } from './objects';
import { mapRefs } from '../calc/parser';

// --------------------------------------------------------------------------------------------
// Pure helpers

export function tableAt(sheet: Sheet, r: number, c: number): TableDef | null {
  return sheet.tables.find((t) => r >= t.range.r1 && r <= t.range.r2 && c >= t.range.c1 && c <= t.range.c2) || null;
}

/** Table names live in the same namespace as defined names. */
export function uniqueTableName(book: Workbook, base = 'Table'): string {
  const taken = new Set<string>();
  for (const s of book.sheets) for (const t of s.tables) taken.add(t.name.toLowerCase());
  for (const n of book.names) taken.add(n.name.toLowerCase());
  let i = 1;
  while (taken.has(`${base}${i}`.toLowerCase())) i++;
  return `${base}${i}`;
}

export function validTableName(book: Workbook, name: string, self?: TableDef): string | null {
  if (!name) return 'Enter a name.';
  if (name.length > 255) return 'The name is too long.';
  if (!/^[A-Za-z_\\][A-Za-z0-9_.\\]*$/.test(name)) return 'A table name must start with a letter, underscore or backslash and contain no spaces or symbols.';
  if (/^[A-Za-z]{1,3}\d+$/.test(name) || /^R\d*C\d*$/i.test(name) || /^[RC]$/i.test(name)) return 'A table name cannot look like a cell reference.';
  for (const s of book.sheets) for (const t of s.tables) if (t !== self && t.name.toLowerCase() === name.toLowerCase()) return 'That name is already used by another table.';
  if (book.names.some((n) => n.name.toLowerCase() === name.toLowerCase())) return 'That name is already used by a defined name.';
  return null;
}

/** Header text for a cell (header names are always text). */
function headerText(sheet: Sheet, r: number, c: number): string {
  const cell = sheet.get(r, c);
  if (!cell || cell.v == null) return '';
  if (typeof cell.v === 'number') return formatValue(cell.v, sheet.styleOf(r, c).numFmt, { date1904: sheet.book.calc.date1904 }).text;
  return String(cell.v).trim();
}

/** Unique, non-blank column names (blank → ColumnN, duplicates get a number). */
export function uniqueColumnNames(names: string[]): string[] {
  const used = new Set<string>();
  return names.map((n, i) => {
    let base = String(n || '').replace(/[\r\n]+/g, ' ').trim() || `Column${i + 1}`;
    let nm = base, k = 2;
    while (used.has(nm.toLowerCase())) nm = `${base}${k++}`;
    used.add(nm.toLowerCase());
    if (!n) base = nm;
    return nm;
  });
}

/** Escapes a column name for a structured reference. */
export function colRef(name: string): string {
  const esc = name.replace(/(['#[\]])/g, "'$1");
  return /^[A-Za-z0-9_]+$/.test(name) ? esc : `[${esc}]`;
}

export function tableDataRows(t: TableDef): Range {
  return { r1: t.range.r1 + (t.opts.header ? 1 : 0), c1: t.range.c1, r2: t.range.r2 - (t.opts.totals ? 1 : 0), c2: t.range.c2 };
}

/**
 * Creates a table over rg (inside a transaction). Without headers, a header row is inserted above the data
 * (shifting the table's columns down) and filled with Column1, Column2...
 */
export function createTable(sheet: Sheet, rg: Range, opts: { headers: boolean; style?: string; name?: string } = { headers: true }): TableDef {
  const book = sheet.book;
  for (const t of sheet.tables) if (intersects(t.range, rg)) throw new Error('A table can\'t overlap another table.');
  for (const p of sheet.pivots) if (p.lastRange && intersects(p.lastRange, rg)) throw new Error('A table can\'t overlap a PivotTable.');
  if (sheet.merges.some((m) => intersects(m, rg))) throw new Error('A table can\'t contain merged cells. Unmerge them first.');
  book.recordSheet(sheet);
  let range = { ...rg };
  const width = rg.c2 - rg.c1 + 1;
  let names: string[];
  if (opts.headers) {
    names = uniqueColumnNames(Array.from({ length: width }, (_, i) => headerText(sheet, rg.r1, rg.c1 + i)));
    if (range.r2 === range.r1) range.r2++; // a header needs at least one data row
  } else {
    names = uniqueColumnNames(Array.from({ length: width }, () => ''));
    insertCellsShift(sheet, { r1: rg.r1, c1: rg.c1, r2: rg.r1, c2: rg.c2 }, 'down');
    range = { r1: rg.r1, c1: rg.c1, r2: rg.r2 + 1, c2: rg.c2 };
  }
  // Header cells hold the (text) names.
  names.forEach((n, i) => { const cur = sheet.get(range.r1, range.c1 + i); sheet.set(range.r1, range.c1 + i, { ...(cur?.s ? { s: cur.s } : {}), v: n }); });
  // An AutoFilter on the same data becomes the table's filter buttons.
  if (sheet.filter && intersects(sheet.filter.range, range)) { sheet.filter = null; }
  const t: TableDef = { id: newId('tbl'), name: opts.name || uniqueTableName(book), range, columns: names.map((name) => ({ name })), style: opts.style || 'TableStyleMedium2', opts: { ...DEFAULT_TABLE_OPTS } };
  sheet.tables.push(t);
  return t;
}

/** Re-reads header cell text into the table's column names (after the user typed over a header). */
export function syncHeaderNames(sheet: Sheet, t: TableDef): void {
  if (!t.opts.header) return;
  const names = uniqueColumnNames(t.columns.map((_, i) => headerText(sheet, t.range.r1, t.range.c1 + i)));
  names.forEach((n, i) => { if (t.columns[i]) t.columns[i].name = n; });
}

/** Rewrites structured references to a table into A1 references (used by Convert to Range). */
export function structuredToA1(f: string, sheet: Sheet, t: TableDef, home: Sheet, row?: number): string {
  const nameRe = t.name.replace(/[.\\]/g, (m) => '\\' + m);
  const dr = tableDataRows(t);
  const colIdx = (n: string) => t.columns.findIndex((c) => c.name.toLowerCase() === n.replace(/'(.)/g, '$1').toLowerCase());
  const ref = (rg: Range) => rangeRefText(sheet, { r1: rg.r1, c1: rg.c1, r2: rg.r2, c2: rg.c2 }, true, home).replace(/^(\$?[A-Z]+\$?\d+):\1$/, '$1');
  const body = (spec: string): string | null => {
    const inner = spec.slice(1, -1).trim();
    let r1 = dr.r1, r2 = dr.r2, c1 = t.range.c1, c2 = t.range.c2;
    const items: string[] = [];
    let thisRow = false;
    if (inner.startsWith('@')) { thisRow = true; const rest = inner.slice(1).trim().replace(/^\[|\]$/g, ''); if (rest) items.push(rest); }
    else if (!inner.startsWith('[')) { if (inner) items.push(inner); }
    else { const re = /\[((?:[^\]']|'.)*)\]|:/g; let m: RegExpExecArray | null; let range = false; while ((m = re.exec(inner))) { if (m[0] === ':') { range = true; continue; } if (range) { items[items.length - 1] += '\u0000' + m[1]; range = false; } else items.push(m[1]); } }
    for (const it of items) {
      const low = it.toLowerCase();
      if (low === '#all') { r1 = t.range.r1; r2 = t.range.r2; }
      else if (low === '#data') { /* default */ }
      else if (low === '#headers') { r1 = r2 = t.range.r1; }
      else if (low === '#totals') { r1 = r2 = t.range.r2; }
      else if (low === '#this row') thisRow = true;
      else { const [a, b] = it.split('\u0000'); const ia = colIdx(a), ib = b != null ? colIdx(b) : ia; if (ia < 0 || ib < 0) return null; c1 = t.range.c1 + Math.min(ia, ib); c2 = t.range.c1 + Math.max(ia, ib); }
    }
    if (thisRow) { if (row == null) return null; r1 = r2 = row; }
    return ref({ r1, c1, r2, c2 });
  };
  // Table[...] and, inside the table, [@Col] / [Col] without a table name.
  let out = f.replace(new RegExp(`\\b${nameRe}(\\[(?:[^\\[\\]]|\\[(?:[^\\]']|'.)*\\])*\\])?`, 'gi'), (m, spec: string | undefined) => body(spec || '[]') ?? m);
  if (home === sheet && row != null && row >= t.range.r1 && row <= t.range.r2) out = out.replace(/(^|[^A-Za-z0-9_\]'])(\[@?(?:[^\[\]]|\[(?:[^\]']|'.)*\])*\])/g, (m, pre: string, spec: string) => { const b = body(spec); return b ? pre + b : m; });
  return out;
}

/** Bakes the table's look into cell formats and removes the table (Convert to Range). */
export function convertTableToRange(sheet: Sheet, t: TableDef, theme: ThemeDef): void {
  const book = sheet.book;
  book.recordSheet(sheet);
  const rg = t.range;
  for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) {
    const look = tableCellLook(t, r, c, theme);
    if (!look) continue;
    const cur = sheet.get(r, c);
    const base = book.styles[cur?.s ?? 0] || {};
    const patch: Partial<CellStyle> = {};
    if (look.fill && !base.fill) patch.fill = look.fill;
    if (look.color && !base.color) patch.color = look.color;
    if (look.bold && base.bold == null) patch.bold = true;
    if (look.borderBottom && !base.bb) patch.bb = { style: 'thin', color: look.borderBottom };
    if (look.borderTop && !base.bt) patch.bt = { style: 'thin', color: look.borderTop };
    if (!Object.keys(patch).length) continue;
    const s = book.derivedStyle(cur?.s ?? 0, patch);
    sheet.set(r, c, { ...(cur || {}), s });
  }
  // Formulas that use structured references to this table now use cell references.
  for (const sh of book.sheets) for (const [k, cell] of [...sh.cells]) {
    if (cell.f == null || !new RegExp(`\\b${t.name}\\[|\\[@|\\[#`, 'i').test(cell.f)) continue;
    const nf = structuredToA1(cell.f, sheet, t, sh, sh === sheet ? keyRow(k) : undefined);
    if (nf !== cell.f) sh.set(keyRow(k), keyCol(k), { ...cell, f: nf });
  }
  if (t.filter) for (let r = rg.r1; r <= rg.r2; r++) sheet.hiddenRows.delete(r);
  sheet.tables = sheet.tables.filter((x) => x !== t);
}

const TOTAL_FNS: [string, string, number][] = [
  ['none', 'None', 0], ['average', 'Average', 101], ['count', 'Count', 103], ['countNums', 'Count Numbers', 102], ['max', 'Max', 104], ['min', 'Min', 105],
  ['sum', 'Sum', 109], ['stdDev', 'StdDev', 107], ['var', 'Var', 110],
];

function totalFormula(t: TableDef, col: number, fn: string): string | null {
  const f = TOTAL_FNS.find((x) => x[0] === fn);
  if (!f || !f[2]) return null;
  return `SUBTOTAL(${f[2]},${t.name}[${colRef(t.columns[col].name)}])`;
}

function isNumericColumn(sheet: Sheet, t: TableDef, i: number): boolean {
  const d = tableDataRows(t);
  let nums = 0, texts = 0;
  for (let r = d.r1; r <= d.r2; r++) { const v = sheet.value(r, t.range.c1 + i); if (typeof v === 'number') nums++; else if (typeof v === 'string' && v) texts++; }
  return nums > 0 && nums >= texts;
}

/** Turns the total row on or off (inside a transaction). */
export function setTotalsRow(sheet: Sheet, t: TableDef, on: boolean): void {
  if (t.opts.totals === on) return;
  sheet.book.recordSheet(sheet);
  if (on) {
    const r = t.range.r2 + 1;
    let busy = false;
    for (let c = t.range.c1; c <= t.range.c2; c++) { const x = sheet.get(r, c); if (x && (x.v != null || x.f != null)) busy = true; }
    if (busy) insertCellsShift(sheet, { r1: r, c1: t.range.c1, r2: r, c2: t.range.c2 }, 'down');
    t.range = { ...t.range, r2: r };
    t.opts = { ...t.opts, totals: true };
    let lastNum = -1;
    t.columns.forEach((_, i) => { if (isNumericColumn(sheet, { ...t, opts: { ...t.opts } }, i)) lastNum = i; });
    t.columns.forEach((col, i) => {
      const c = t.range.c1 + i;
      if (i === 0 && lastNum !== 0) { sheet.set(r, c, { v: 'Total' }); col.totalFn = undefined; }
      else if (i === lastNum || (col.totalFn && col.totalFn !== 'none')) { const fn = col.totalFn && col.totalFn !== 'none' ? col.totalFn : 'sum'; col.totalFn = fn; const f = totalFormula(t, i, fn); sheet.set(r, c, f ? { f } : undefined); }
      else sheet.set(r, c, undefined);
    });
  } else {
    const r = t.range.r2;
    for (let c = t.range.c1; c <= t.range.c2; c++) { const cur = sheet.get(r, c); sheet.set(r, c, cur?.s ? { s: cur.s } : undefined); }
    t.range = { ...t.range, r2: Math.max(t.range.r1 + (t.opts.header ? 1 : 0), r - 1) };
    t.opts = { ...t.opts, totals: false };
  }
}

/** Turns the header row on or off (inside a transaction). */
export function setHeaderRow(sheet: Sheet, t: TableDef, on: boolean): void {
  if (t.opts.header === on) return;
  sheet.book.recordSheet(sheet);
  if (!on) {
    syncHeaderNames(sheet, t);
    for (let c = t.range.c1; c <= t.range.c2; c++) { const cur = sheet.get(t.range.r1, c); sheet.set(t.range.r1, c, cur?.s ? { s: cur.s } : undefined); }
    t.range = { ...t.range, r1: Math.min(t.range.r1 + 1, t.range.r2) };
    t.opts = { ...t.opts, header: false };
  } else {
    let r = t.range.r1 - 1;
    let free = r >= 0;
    if (free) for (let c = t.range.c1; c <= t.range.c2; c++) { const x = sheet.get(r, c); if (x && (x.v != null || x.f != null)) free = false; }
    if (!free || sheet.tables.some((o) => o !== t && intersects(o.range, { r1: r, c1: t.range.c1, r2: r, c2: t.range.c2 }))) {
      insertCellsShift(sheet, { r1: t.range.r1, c1: t.range.c1, r2: t.range.r1, c2: t.range.c2 }, 'down');
      r = t.range.r1;
      t.range = { ...t.range, r2: t.range.r2 + 1 };
    } else t.range = { ...t.range, r1: r };
    t.columns.forEach((col, i) => sheet.set(r, t.range.c1 + i, { v: col.name }));
    t.opts = { ...t.opts, header: true };
  }
}

/** Renames a table and updates formulas that refer to it. */
export function renameTable(book: Workbook, t: TableDef, name: string): void {
  const old = t.name;
  book.recordBook(); // defined names are rewritten too, and they are not part of a cell or sheet record
  const re = new RegExp(`(^|[^A-Za-z0-9_.\\\\])${old.replace(/[.\\]/g, (m) => '\\' + m)}(?=\\[|\\b)`, 'gi');
  for (const sh of book.sheets) for (const [k, cell] of [...sh.cells]) {
    if (cell.f == null) continue;
    const nf = cell.f.replace(re, (_m, pre: string) => pre + name);
    if (nf !== cell.f) sh.set(keyRow(k), keyCol(k), { ...cell, f: nf });
  }
  for (const n of book.names) n.ref = n.ref.replace(re, (_m, pre: string) => pre + name);
  t.name = name;
}

// --------------------------------------------------------------------------------------------
// UI helpers

function currentTable(app: App): { sheet: Sheet; t: TableDef } | null {
  const s = app.sheet;
  const t = tableAt(s, s.sel.r, s.sel.c);
  return t ? { sheet: s, t } : null;
}

function needTable(app: App): { sheet: Sheet; t: TableDef } | null {
  const ct = currentTable(app);
  if (!ct) toast('Select a cell in a table first.');
  return ct;
}

function protectedSheet(app: App, sheet: Sheet = app.sheet): boolean {
  if (sheet.protection.on) { toast('The sheet is protected. Unprotect it to change tables.'); app.exec('review.protectionNotice'); return true; }
  return false;
}

/** Inline CSS showing a table style in a gallery chip. */
export function tableStylePreviewCss(name: string, theme: ThemeDef): string {
  const l = tableLook(name, theme);
  const hd = l.headerFill || '#ffffff', b1 = l.band1 || '#ffffff', b2 = l.band2 || '#ffffff';
  return `display:block;width:62px;height:40px;font-size:0;border:1px solid ${l.border || l.innerBorder || '#d0d0d0'};background:linear-gradient(to bottom, ${hd} 0 25%, ${b1} 25% 50%, ${b2} 50% 75%, ${b1} 75% 100%)`;
}

/** Small SVG preview of a table style. */
export function tableStylePreviewSvg(name: string, theme: ThemeDef): string {
  const l = tableLook(name, theme);
  const rows = 5, cols = 4, cw = 18, rh = 9;
  let s = `<svg xmlns="http://www.w3.org/2000/svg" width="${cols * cw + 2}" height="${rows * rh + 2}" viewBox="0 0 ${cols * cw + 2} ${rows * rh + 2}"><rect x="1" y="1" width="${cols * cw}" height="${rows * rh}" fill="#ffffff"/>`;
  for (let r = 0; r < rows; r++) {
    const fill = r === 0 ? l.headerFill : r % 2 ? l.band1 : l.band2;
    if (fill) s += `<rect x="1" y="${1 + r * rh}" width="${cols * cw}" height="${rh}" fill="${fill}"/>`;
    const line = r === 0 ? l.border : l.innerBorder;
    if (line && r > 0) s += `<line x1="1" x2="${1 + cols * cw}" y1="${1 + r * rh}" y2="${1 + r * rh}" stroke="${line}" stroke-width="0.8"/>`;
    for (let c = 0; c < cols; c++) s += `<rect x="${4 + c * cw}" y="${4 + r * rh}" width="${cw - 7}" height="2.5" fill="${r === 0 ? l.headerColor || '#000' : l.text || '#7f7f7f'}" opacity="${r === 0 ? 0.9 : 0.5}"/>`;
  }
  if (l.border) s += `<rect x="1" y="1" width="${cols * cw}" height="${rows * rh}" fill="none" stroke="${l.border}"/>`;
  return s + '</svg>';
}

const styleLabel = (n: string) => n.replace(/^TableStyle(Light|Medium|Dark)(\d+)$/, '$1 $2');

async function pickTableStyle(app: App, title: string, current?: string): Promise<string | null> {
  const theme = app.book.theme;
  return galleryDialog(title, TABLE_STYLE_NAMES.map((n) => ({ value: n, label: styleLabel(n), html: tableStylePreviewSvg(n, theme), group: n.includes('Light') ? 'Light' : n.includes('Medium') ? 'Medium' : 'Dark' })), { current, size: 'small' });
}

/** The Create Table dialog; creates the table with the given style. */
export async function createTableFlow(app: App, style?: string): Promise<TableDef | null> {
  if (protectedSheet(app)) return null;
  const s = app.sheet;
  let rg = activeRange(app);
  if (rg.r1 === rg.r2 && rg.c1 === rg.c2) rg = currentRegion(s, rg.r1, rg.c1);
  else if (isWholeCols(rg) || isWholeRows(rg)) rg = clipToUsed(s, rg);
  const guess = hasHeaderRow(s, rg);
  const r = await openDialog({
    title: 'Create Table',
    fields: [
      { key: 'range', label: 'Where is the data for your table?', type: 'text', value: '=' + rangeName({ ...rg }).replace(/([A-Z]+)(\d+)/g, '$$$1$$$2'), autofocus: true },
      { key: 'headers', label: 'My table has headers', type: 'checkbox', value: guess },
    ],
    validate: (v) => (resolveRangeText(app.book, String(v.range).replace(/^=/, '').replace(/\$/g, ''), s) ? null : 'The reference isn\'t valid.'),
  });
  if (!r) return null;
  const res = resolveRangeText(app.book, String(r.range).replace(/^=/, '').replace(/\$/g, ''), s)!;
  if (res.sheet !== s) { toast('The table must be on the active sheet.'); return null; }
  let t: TableDef | null = null;
  try {
    app.book.transact('Create Table', () => { t = createTable(s, res.range, { headers: !!r.headers, style: style || 'TableStyleMedium2' }); });
  } catch (e) { toast((e as Error).message); return null; }
  if (t) app.grid?.select((t as TableDef).range.r1, (t as TableDef).range.c1, { range: (t as TableDef).range, keepActive: false });
  return t;
}

function setOpt(app: App, key: keyof TableDef['opts'], label: string): void {
  const ct = needTable(app);
  if (!ct || protectedSheet(app)) return;
  const { sheet, t } = ct;
  app.book.transact(label, () => {
    if (key === 'totals') setTotalsRow(sheet, t, !t.opts.totals);
    else if (key === 'header') setHeaderRow(sheet, t, !t.opts.header);
    else { sheet.book.recordSheet(sheet); t.opts = { ...t.opts, [key]: !t.opts[key] }; }
  });
}

const optCmd = (id: string, label: string, key: keyof TableDef['opts'], icon: string, tip: string) => ({
  id, label, icon, tab: 'Table Design', group: 'Table Style Options', kind: 'toggle' as const, edits: true, tip,
  run: (app: App) => setOpt(app, key, label), isActive: (app: App) => !!currentTable(app)?.t.opts[key], isEnabled: (app: App) => !!currentTable(app),
});

function tableRowsInSelection(app: App, t: TableDef): number[] {
  const d = tableDataRows(t);
  const out = new Set<number>();
  for (const rg of app.sheet.sel.ranges) for (let r = Math.max(rg.r1, d.r1); r <= Math.min(rg.r2, d.r2); r++) out.add(r);
  return [...out].sort((a, b) => a - b);
}
function tableColsInSelection(app: App, t: TableDef): number[] {
  const out = new Set<number>();
  for (const rg of app.sheet.sel.ranges) for (let c = Math.max(rg.c1, t.range.c1); c <= Math.min(rg.c2, t.range.c2); c++) out.add(c);
  return [...out].sort((a, b) => a - b);
}

/** Inserts n table rows at row r (within the table's columns), filling calculated columns. */
export function insertTableRows(sheet: Sheet, t: TableDef, r: number, n = 1): void {
  sheet.book.recordSheet(sheet);
  insertCellsShift(sheet, { r1: r, c1: t.range.c1, r2: r + n - 1, c2: t.range.c2 }, 'down');
  t.range = { ...t.range, r2: t.range.r2 + n };
  // Copy formats from the row above and fill calculated columns.
  t.columns.forEach((col, i) => {
    const c = t.range.c1 + i;
    const above = sheet.get(r - 1, c);
    for (let k = 0; k < n; k++) {
      const cell: Cell = {};
      if (above?.s && r - 1 > t.range.r1) cell.s = above.s;
      if (col.calcFormula) cell.f = col.calcFormula;
      if (cell.s || cell.f) sheet.set(r + k, c, cell);
    }
  });
}

// --------------------------------------------------------------------------------------------
// Commands

def([
  {
    id: 'insert.table', label: 'Table', icon: 'table', tab: 'Insert', group: 'Tables', keys: ['Mod-t', 'Mod-l'], edits: true, popular: true, keywords: 'create table list',
    run: (app) => { const ct = currentTable(app); if (ct) { toast(`This cell is already in ${ct.t.name}.`); return; } return createTableFlow(app); },
  },
  {
    id: 'table.formatAsTable', label: 'Format as Table', icon: 'table_view', tab: 'Home', group: 'Styles', kind: 'menu', edits: true, popular: true, keywords: 'table style banded',
    run: async (app, arg) => {
      const ct = currentTable(app);
      const style = typeof arg === 'string' && TABLE_STYLE_NAMES.includes(arg) ? arg : await pickTableStyle(app, 'Format as Table', ct?.t.style);
      if (!style) return;
      if (ct) { if (protectedSheet(app)) return; app.book.transact('Table Style', () => { app.book.recordSheet(ct.sheet); ct.t.style = style; }); return; }
      await createTableFlow(app, style);
    },
    menu: (app) => {
      const theme = app.book.theme;
      const pick = (fam: string, n: number[]): MenuEntry[] => n.map((i) => { const nm = `TableStyle${fam}${i}`; return { id: 'table.formatAsTable', arg: nm, label: styleLabel(nm), swatch: tableLook(nm, theme).headerFill || tableLook(nm, theme).band1 || '#ffffff' }; });
      return [{ heading: 'Light' }, ...pick('Light', [1, 2, 9, 15]), { heading: 'Medium' }, ...pick('Medium', [1, 2, 3, 4, 9, 16, 23]), { heading: 'Dark' }, ...pick('Dark', [1, 2, 9]), { separator: true }, { id: 'table.formatAsTable', label: 'All Table Styles…', icon: 'table_view' }, { id: 'table.newStyle', label: 'New Table Style…', icon: 'add' }];
    },
  },
  {
    id: 'table.newStyle', label: 'New Table Style', icon: 'add', tab: 'Home', group: 'Styles', edits: true,
    run: () => toast('Custom table styles aren\'t available yet. Pick a built-in style, then adjust cells with Format Cells.'),
  },
  {
    id: 'table.name', label: 'Table Name', icon: 'badge', tab: 'Table Design', group: 'Properties', edits: true, isEnabled: (app) => !!currentTable(app),
    value: (app) => currentTable(app)?.t.name ?? null,
    run: async (app, arg) => {
      const ct = needTable(app);
      if (!ct || protectedSheet(app)) return;
      const name = typeof arg === 'string' ? arg : await promptDialog('Table Name', 'Name', ct.t.name, { help: 'Formulas that use the old name are updated.' } as any);
      if (name == null || name === ct.t.name) return;
      const err = validTableName(app.book, name.trim(), ct.t);
      if (err) { toast(err); return; }
      app.book.transact('Rename Table', () => { app.book.recordSheet(ct.sheet); renameTable(app.book, ct.t, name.trim()); });
    },
  },
  {
    id: 'table.resize', label: 'Resize Table', icon: 'aspect_ratio', tab: 'Table Design', group: 'Properties', edits: true, isEnabled: (app) => !!currentTable(app),
    run: async (app) => {
      const ct = needTable(app);
      if (!ct || protectedSheet(app)) return;
      const { sheet, t } = ct;
      const r = await openDialog({
        title: 'Resize Table', help: 'The headers must stay in the same row, and the new range must overlap the table.',
        fields: [{ key: 'range', label: 'Select the new data range for your table', type: 'text', value: '=' + rangeName(t.range).replace(/([A-Z]+)(\d+)/g, '$$$1$$$2'), autofocus: true }],
        validate: (v) => {
          const rg = parseRange(String(v.range).replace(/^=/, '').replace(/\$/g, '').replace(/^.*!/, ''));
          if (!rg) return 'The reference isn\'t valid.';
          if (rg.r1 !== t.range.r1) return 'The header row must stay in the same row.';
          if (!intersects(rg, t.range)) return 'The new range must overlap the table.';
          if (sheet.tables.some((o) => o !== t && intersects(o.range, rg))) return 'A table can\'t overlap another table.';
          if (rg.r2 <= rg.r1 && t.opts.header) return 'The table needs at least one data row.';
          return null;
        },
      });
      if (!r) return;
      const rg = parseRange(String(r.range).replace(/^=/, '').replace(/\$/g, '').replace(/^.*!/, ''))!;
      app.book.transact('Resize Table', () => {
        sheet.book.recordSheet(sheet);
        const old = t.range;
        const cols = [];
        for (let c = rg.c1; c <= rg.c2; c++) {
          const prev = c >= old.c1 && c <= old.c2 ? t.columns[c - old.c1] : null;
          cols.push(prev ? { ...prev } : { name: t.opts.header ? headerText(sheet, rg.r1, c) : '' });
        }
        const names = uniqueColumnNames(cols.map((x) => x.name));
        cols.forEach((x, i) => { x.name = names[i]; if (t.opts.header && !(rg.c1 + i >= old.c1 && rg.c1 + i <= old.c2)) sheet.set(rg.r1, rg.c1 + i, { v: names[i] }); });
        // A total row moves to the new last row.
        if (t.opts.totals && rg.r2 !== old.r2) { for (let c = old.c1; c <= old.c2; c++) { const cur = sheet.get(old.r2, c); if (cur && rg.r2 > old.r2) sheet.set(rg.r2, c, cur); sheet.set(old.r2, c, cur?.s ? { s: cur.s } : undefined); } }
        t.columns = cols;
        t.range = rg;
      });
    },
  },
  { id: 'table.summarizePivot', label: 'Summarize with PivotTable', icon: 'pivot_table_chart', tab: 'Table Design', group: 'Tools', edits: true, run: (app) => { const ct = needTable(app); if (ct) app.exec('pivot.create', { source: ct.t.name }); } },
  {
    id: 'table.removeDuplicates', label: 'Remove Duplicates', icon: 'layers_clear', tab: 'Table Design', group: 'Tools', edits: true,
    run: (app) => { const ct = needTable(app); if (!ct) return; const d = tableDataRows(ct.t); app.grid?.select(t0(ct.t), d.c1, { range: { ...d, r1: ct.t.range.r1 }, keepActive: true }); app.exec('data.removeDuplicates'); },
  },
  {
    id: 'table.convertToRange', label: 'Convert to Range', icon: 'grid_on', tab: 'Table Design', group: 'Tools', edits: true,
    run: async (app) => {
      const ct = needTable(app);
      if (!ct || protectedSheet(app)) return;
      if (!(await confirmDialog('Convert to Range', `Do you want to convert ${ct.t.name} to a normal range? The table's look is kept as cell formatting.`, 'Yes', 'No'))) return;
      app.book.transact('Convert to Range', () => convertTableToRange(ct.sheet, ct.t, app.book.theme));
    },
  },
  optCmd('table.headerRow', 'Header Row', 'header', 'table_rows', 'Show the header row with the column names'),
  optCmd('table.totalRow', 'Total Row', 'totals', 'functions', 'Add a row of totals under the table'),
  optCmd('table.bandedRows', 'Banded Rows', 'banded', 'density_medium', 'Shade every other row'),
  optCmd('table.firstColumn', 'First Column', 'firstCol', 'format_bold', 'Bold the first column'),
  optCmd('table.lastColumn', 'Last Column', 'lastCol', 'format_bold', 'Bold the last column'),
  optCmd('table.bandedColumns', 'Banded Columns', 'bandedCols', 'view_week', 'Shade every other column'),
  optCmd('table.filterButton', 'Filter Button', 'filterButton', 'filter_list', 'Show filter buttons in the header row'),
  {
    id: 'table.style', label: 'Table Styles', icon: 'table_view', tab: 'Table Design', group: 'Table Styles', kind: 'gallery', edits: true,
    run: async (app, arg) => {
      const ct = needTable(app);
      if (!ct || protectedSheet(app)) return;
      const style = typeof arg === 'string' && TABLE_STYLE_NAMES.includes(arg) ? arg : await pickTableStyle(app, 'Table Styles', ct.t.style);
      if (!style) return;
      app.book.transact('Table Style', () => { app.book.recordSheet(ct.sheet); ct.t.style = style; });
    },
    menu: (app) => {
      const cur = currentTable(app)?.t.style;
      const theme = app.book.theme;
      return TABLE_STYLE_NAMES.map((n) => ({ id: 'table.style', arg: n, label: styleLabel(n), checked: n === cur, preview: tableStylePreviewCss(n, theme) }));
    },
  },
  {
    id: 'table.totalFunction', label: 'Total Row Function', icon: 'functions', tab: 'Table Design', group: 'Table Style Options', kind: 'menu', edits: true, isEnabled: (app) => !!currentTable(app)?.t.opts.totals,
    run: async (app, arg: { col?: number; fn: string } | string) => {
      const ct = needTable(app);
      if (!ct || protectedSheet(app)) return;
      const { sheet, t } = ct;
      if (!t.opts.totals) { toast('Turn on the Total Row first.'); return; }
      const a = typeof arg === 'string' ? { fn: arg } : arg || { fn: 'sum' };
      const col = a.col ?? Math.max(0, Math.min(t.columns.length - 1, sheet.sel.c - t.range.c1));
      let formula: string | null = null;
      if (a.fn === 'more') {
        const f = await promptDialog('Total Row: More Functions', 'Formula', `=SUBTOTAL(109,${t.name}[${colRef(t.columns[col].name)}])`);
        if (f == null) return;
        formula = f.replace(/^=/, '');
      } else formula = totalFormula(t, col, a.fn);
      app.book.transact('Total Row Function', () => {
        sheet.book.recordSheet(sheet);
        t.columns[col].totalFn = a.fn === 'more' ? 'custom' : a.fn === 'none' ? undefined : a.fn;
        t.columns[col].totalFormula = a.fn === 'more' ? formula || undefined : undefined;
        const cur = sheet.get(t.range.r2, t.range.c1 + col);
        sheet.set(t.range.r2, t.range.c1 + col, formula ? { ...(cur?.s ? { s: cur.s } : {}), f: formula } : cur?.s ? { s: cur.s } : undefined);
      });
    },
    menu: (app) => {
      const ct = currentTable(app);
      const col = ct ? Math.max(0, Math.min(ct.t.columns.length - 1, app.sheet.sel.c - ct.t.range.c1)) : 0;
      const cur = ct?.t.columns[col]?.totalFn || 'none';
      return [...TOTAL_FNS.map(([k, l]) => ({ id: 'table.totalFunction', arg: { col, fn: k }, label: l, checked: k === cur })), { separator: true }, { id: 'table.totalFunction', arg: { col, fn: 'more' }, label: 'More Functions…' }];
    },
  },
  {
    id: 'table.calculatedColumn', label: 'Fill Calculated Column', icon: 'calculate', tab: 'Table Design', group: 'Tools', edits: true, isEnabled: (app) => !!currentTable(app),
    run: (app) => {
      const ct = needTable(app);
      if (!ct || protectedSheet(app)) return;
      const { sheet, t } = ct;
      const { r, c } = sheet.sel;
      const f = sheet.get(r, c)?.f;
      if (!f) { toast('Enter a formula in a table cell first; it is then filled down the whole column.'); return; }
      const d = tableDataRows(t);
      app.book.transact('Calculated Column', () => {
        sheet.book.recordSheet(sheet);
        t.columns[c - t.range.c1].calcFormula = f;
        for (let rr = d.r1; rr <= d.r2; rr++) if (rr !== r) { const cur = sheet.get(rr, c); sheet.set(rr, c, { ...(cur?.s ? { s: cur.s } : {}), f }); }
      });
    },
  },
  {
    id: 'table.insertRowAbove', label: 'Insert Table Rows Above', icon: 'table_rows', tab: 'Table Design', group: 'Rows & Columns', edits: true, isEnabled: (app) => !!currentTable(app),
    run: (app) => {
      const ct = needTable(app);
      if (!ct || protectedSheet(app)) return;
      const rows = tableRowsInSelection(app, ct.t);
      const at = rows[0] ?? tableDataRows(ct.t).r1;
      app.book.transact('Insert Table Rows', () => insertTableRows(ct.sheet, ct.t, at, Math.max(1, rows.length)));
    },
  },
  {
    id: 'table.insertRowBelow', label: 'Insert Table Rows Below', icon: 'table_rows', tab: 'Table Design', group: 'Rows & Columns', edits: true, isEnabled: (app) => !!currentTable(app),
    run: (app) => {
      const ct = needTable(app);
      if (!ct || protectedSheet(app)) return;
      const rows = tableRowsInSelection(app, ct.t);
      const at = (rows[rows.length - 1] ?? tableDataRows(ct.t).r2) + 1;
      app.book.transact('Insert Table Rows', () => insertTableRows(ct.sheet, ct.t, at, Math.max(1, rows.length)));
    },
  },
  {
    id: 'table.insertColumnLeft', label: 'Insert Table Columns to the Left', icon: 'view_column', tab: 'Table Design', group: 'Rows & Columns', edits: true, isEnabled: (app) => !!currentTable(app),
    run: (app) => insertCols(app, false),
  },
  {
    id: 'table.insertColumnRight', label: 'Insert Table Columns to the Right', icon: 'view_column', tab: 'Table Design', group: 'Rows & Columns', edits: true, isEnabled: (app) => !!currentTable(app),
    run: (app) => insertCols(app, true),
  },
  {
    id: 'table.deleteRows', label: 'Delete Table Rows', icon: 'delete_sweep', tab: 'Table Design', group: 'Rows & Columns', edits: true, isEnabled: (app) => !!currentTable(app),
    run: (app) => {
      const ct = needTable(app);
      if (!ct || protectedSheet(app)) return;
      const { sheet, t } = ct;
      const rows = tableRowsInSelection(app, t);
      if (!rows.length) { toast('Select data rows of the table.'); return; }
      const d = tableDataRows(t);
      app.book.transact('Delete Table Rows', () => {
        sheet.book.recordSheet(sheet);
        for (const r of [...rows].reverse()) {
          if (t.range.r2 - (t.opts.totals ? 1 : 0) <= t.range.r1 + (t.opts.header ? 1 : 0)) {
            // Keep one (empty) data row.
            for (let c = t.range.c1; c <= t.range.c2; c++) sheet.set(r, c, undefined);
            continue;
          }
          insertCellsShift(sheet, { r1: r, c1: t.range.c1, r2: r, c2: t.range.c2 }, 'up');
          t.range = { ...t.range, r2: t.range.r2 - 1 };
        }
      });
      void d;
    },
  },
  {
    id: 'table.deleteColumns', label: 'Delete Table Columns', icon: 'delete_sweep', tab: 'Table Design', group: 'Rows & Columns', edits: true, isEnabled: (app) => !!currentTable(app),
    run: (app) => {
      const ct = needTable(app);
      if (!ct || protectedSheet(app)) return;
      const { sheet, t } = ct;
      const cols = tableColsInSelection(app, t);
      app.book.transact('Delete Table Columns', () => {
        sheet.book.recordSheet(sheet);
        for (const c of [...cols].reverse()) {
          insertCellsShift(sheet, { r1: t.range.r1, c1: c, r2: t.range.r2, c2: c }, 'left');
          t.columns.splice(c - t.range.c1, 1);
          t.range = { ...t.range, c2: t.range.c2 - 1 };
        }
        if (!t.columns.length || t.range.c2 < t.range.c1) sheet.tables = sheet.tables.filter((x) => x !== t);
      });
    },
  },
  {
    id: 'table.selectColumn', label: 'Select Table Column Data', icon: 'view_column', tab: 'Table Design', group: 'Select', noRecord: true, isEnabled: (app) => !!currentTable(app),
    run: (app, arg) => {
      const ct = needTable(app);
      if (!ct) return;
      const d = tableDataRows(ct.t);
      const c = app.sheet.sel.c;
      const withHeader = arg === 'all';
      app.grid?.select(d.r1, c, { range: { r1: withHeader ? ct.t.range.r1 : d.r1, c1: c, r2: withHeader ? ct.t.range.r2 : d.r2, c2: c } });
    },
  },
  {
    id: 'table.selectRow', label: 'Select Table Row', icon: 'table_rows', tab: 'Table Design', group: 'Select', noRecord: true, isEnabled: (app) => !!currentTable(app),
    run: (app) => { const ct = needTable(app); if (!ct) return; const r = app.sheet.sel.r; app.grid?.select(r, ct.t.range.c1, { range: { r1: r, c1: ct.t.range.c1, r2: r, c2: ct.t.range.c2 } }); },
  },
  {
    id: 'table.selectAll', label: 'Select Entire Table', icon: 'select_all', tab: 'Table Design', group: 'Select', noRecord: true, isEnabled: (app) => !!currentTable(app),
    run: (app) => { const ct = needTable(app); if (!ct) return; app.grid?.select(ct.t.range.r1, ct.t.range.c1, { range: { ...ct.t.range } }); },
  },
  {
    id: 'table.syncHeaders', label: 'Update Column Names from Headers', icon: 'sync', tab: 'Table Design', group: 'Properties', edits: true, isEnabled: (app) => !!currentTable(app),
    run: (app) => { const ct = needTable(app); if (!ct) return; app.book.transact('Update Column Names', () => { app.book.recordSheet(ct.sheet); syncHeaderNames(ct.sheet, ct.t); }); },
  },
]);

function t0(t: TableDef): number { return t.range.r1; }

function insertCols(app: App, right: boolean): void {
  const ct = needTable(app);
  if (!ct || protectedSheet(app)) return;
  const { sheet, t } = ct;
  const cols = tableColsInSelection(app, t);
  const n = Math.max(1, cols.length);
  const at = right ? (cols[cols.length - 1] ?? t.range.c2) + 1 : cols[0] ?? t.range.c1;
  app.book.transact('Insert Table Columns', () => {
    sheet.book.recordSheet(sheet);
    insertCellsShift(sheet, { r1: t.range.r1, c1: at, r2: t.range.r2, c2: at + n - 1 }, 'right');
    const names = new Set(t.columns.map((c) => c.name.toLowerCase()));
    const add = [];
    for (let i = 0; i < n; i++) { let k = t.columns.length + i + 1, nm = `Column${k}`; while (names.has(nm.toLowerCase())) nm = `Column${++k}`; names.add(nm.toLowerCase()); add.push({ name: nm }); }
    t.columns.splice(at - t.range.c1, 0, ...add);
    t.range = { ...t.range, c2: t.range.c2 + n };
    if (t.opts.header) add.forEach((c, i) => sheet.set(t.range.r1, at + i, { v: c.name }));
  });
}

void cellKey; void mapRefs;
