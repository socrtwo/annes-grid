/**
 * Helpers shared by the Utilities, Plus and Extras tabs: the three "plug-in" tabs whose commands clone
 * the tools of the well-known spreadsheet add-ins (ASAP Utilities, Kutools, Ablebits Ultimate Suite,
 * Power Tools, XLTools, Inquire) and add some of our own. Everything that touches the workbook goes
 * through the same selection, protection and undo conventions as the rest of the command library.
 */
import type { App } from '../app';
import { Sheet, type Workbook } from '../model/workbook';
import type { Cell, CellStyle, Scalar } from '../model/types';
import { CellError } from '../model/types';
import { cellKey, keyRow, keyCol, colName, cellName, rangeName, parseRange, MAX_ROWS, MAX_COLS, type Range } from '../model/address';
import { selRanges, activeRange, activeCell, clipToUsed, currentRegion, cellText } from './ops';
import { shiftCells } from '../model/structure';
import { checkWritable, canChange, writeOut, addSheetWithRows, HEAD, type OutCell } from './tools';
import { def, type Command } from './registry';
import { openDialog, toast } from '../ui/dialog';

export const U = 'Utilities';
export const P = 'Plus';
export const X = 'Extras';

// --------------------------------------------------------------------------------------------
// Tool log: "Start Last Tool Again" and the usage report

export interface ToolRun { id: string; arg?: any; label: string }
export const toolLog: { last: ToolRun | null; counts: Record<string, number>; since: number } = { last: null, counts: {}, since: Date.now() };

/** Registers commands of the plug-in tabs; each run is logged for Start Last Tool Again and the usage report. */
export function defTools(list: Command[]): void {
  for (const c of list) {
    const run = c.run;
    c.run = (app, arg) => {
      if (!c.noRecord || c.kind === 'menu' || c.kind === 'split') {
        if (!c.id.startsWith('util.last') && c.id !== 'util.findTool' && c.id !== 'util.myTools') { toolLog.last = { id: c.id, arg, label: c.label }; toolLog.counts[c.id] = (toolLog.counts[c.id] || 0) + 1; }
      }
      return run(app, arg);
    };
  }
  def(list);
}

// --------------------------------------------------------------------------------------------
// Selection helpers

export interface SelCell { r: number; c: number; cell: Cell | undefined }

/** Every distinct cell position of the selection, clipped to the used range; rows hidden by a filter are skipped. */
export function selCells(app: App, opts: { sheet?: Sheet; existingOnly?: boolean; pad?: number; includeFiltered?: boolean; /** For tools that fill: every selected cell, used or not (whole rows and columns still stop at the data). */ fill?: boolean } = {}): SelCell[] {
  const sheet = opts.sheet || app.sheet;
  const out: SelCell[] = [];
  const seen = new Set<number>();
  for (const rg0 of selRanges(app)) {
    const bounded = rg0.r2 - rg0.r1 <= 100000 && rg0.c2 - rg0.c1 <= 1000;
    const rg = opts.fill && bounded ? rg0 : clipToUsed(sheet, rg0, opts.pad || 0);
    for (let r = rg.r1; r <= rg.r2; r++) {
      if (!opts.includeFiltered && sheet.hiddenRows.has(r) && sheet.filter) continue;
      for (let c = rg.c1; c <= rg.c2; c++) {
        const k = cellKey(r, c);
        if (seen.has(k)) continue;
        seen.add(k);
        const cell = sheet.get(r, c);
        if (opts.existingOnly && !cell) continue;
        out.push({ r, c, cell });
      }
    }
  }
  return out;
}

/** The selected area as one block: the active range clipped to the used range, or the current region around a lone cell. */
export function selBlock(app: App): Range {
  const rg = activeRange(app);
  if (rg.r1 === rg.r2 && rg.c1 === rg.c2) return currentRegion(app.sheet, rg.r1, rg.c1);
  return clipToUsed(app.sheet, rg);
}

/** True when the selection is a single cell (no range). */
export function singleCell(app: App): boolean { const rg = activeRange(app); return rg.r1 === rg.r2 && rg.c1 === rg.c2; }

export const hasContent = (cell: Cell | undefined): boolean => !!cell && (cell.f != null || (cell.v != null && cell.v !== ''));

/** Selects the given cells (merged into row runs). Returns the number selected; says so when there are none. */
export function selectHits(app: App, hits: { r: number; c: number }[], what = 'cells'): number {
  if (!hits.length) { toast(`No ${what} were found.`); return 0; }
  hits.sort((a, b) => a.r - b.r || a.c - b.c);
  const runs: Range[] = [];
  for (const hh of hits) { const last = runs[runs.length - 1]; if (last && last.r1 === hh.r && last.c2 + 1 === hh.c) last.c2 = hh.c; else runs.push({ r1: hh.r, c1: hh.c, r2: hh.r, c2: hh.c }); }
  // Merge vertically identical runs into blocks for compact selections
  const blocks: Range[] = [];
  for (const rn of runs) { const last = blocks[blocks.length - 1]; if (last && last.c1 === rn.c1 && last.c2 === rn.c2 && last.r2 + 1 === rn.r1) last.r2 = rn.r1; else blocks.push({ ...rn }); }
  const s = app.sheet;
  s.sel.ranges = blocks.slice(0, 5000);
  s.sel.r = blocks[0].r1; s.sel.c = blocks[0].c1;
  app.grid.scrollIntoView(s.sel.r, s.sel.c);
  app.grid.invalidate();
  app.ui.status(`${hits.length.toLocaleString()} ${what} selected`);
  return hits.length;
}

/** Selects the cells of the selection (or the used range) that satisfy pred. */
export function selectWhere(app: App, pred: (cell: Cell | undefined, r: number, c: number) => boolean, opts: { what?: string; wholeSheet?: boolean; existingOnly?: boolean } = {}): number {
  const s = app.sheet;
  const hits: { r: number; c: number }[] = [];
  if (opts.wholeSheet || singleCell(app)) {
    const u = s.usedRange();
    if (!u) { toast('The sheet is empty.'); return 0; }
    if (opts.existingOnly !== false) s.forEachIn(u, (r, c, cell) => { if (pred(cell, r, c)) hits.push({ r, c }); });
    else for (let r = u.r1; r <= u.r2; r++) for (let c = u.c1; c <= u.c2; c++) if (pred(s.get(r, c), r, c)) hits.push({ r, c });
  } else {
    for (const x of selCells(app, { existingOnly: opts.existingOnly !== false })) if (pred(x.cell, x.r, x.c)) hits.push({ r: x.r, c: x.c });
  }
  return selectHits(app, hits, opts.what || 'cells');
}

export function selectRange(app: App, rg: Range, keepActive = false): void {
  app.grid.select(keepActive ? app.sheet.sel.r : rg.r1, keepActive ? app.sheet.sel.c : rg.c1, { range: rg, keepActive });
}

/** Rows of the selection as [first, last] spans (merged, ascending). */
export function rowSpans(app: App, clip = true): [number, number][] {
  const s = app.sheet;
  const out: [number, number][] = selRanges(app).map((rg) => { const c = clip ? clipToUsed(s, rg) : rg; return [c.r1, c.r2]; });
  return mergeSpans(out);
}

export function colSpans(app: App, clip = true): [number, number][] {
  const s = app.sheet;
  const out: [number, number][] = selRanges(app).map((rg) => { const c = clip ? clipToUsed(s, rg) : rg; return [c.c1, c.c2]; });
  return mergeSpans(out);
}

export function mergeSpans(spans: [number, number][]): [number, number][] {
  const sorted = spans.map((x) => [Math.min(x[0], x[1]), Math.max(x[0], x[1])] as [number, number]).sort((a, b) => a[0] - b[0]);
  const out: [number, number][] = [];
  for (const sp of sorted) { const l = out[out.length - 1]; if (l && sp[0] <= l[1] + 1) l[1] = Math.max(l[1], sp[1]); else out.push([...sp]); }
  return out;
}

export function spanList(spans: [number, number][]): number[] { const out: number[] = []; for (const [a, b] of spans) for (let i = a; i <= b; i++) out.push(i); return out; }

export function rowIsBlank(s: Sheet, r: number, c1 = 0, c2 = MAX_COLS - 1): boolean {
  const u = s.usedRange(false);
  if (!u || r < u.r1 || r > u.r2) return true;
  for (let c = Math.max(c1, u.c1); c <= Math.min(c2, u.c2); c++) if (hasContent(s.get(r, c))) return false;
  return true;
}

export function colIsBlank(s: Sheet, c: number, r1 = 0, r2 = MAX_ROWS - 1): boolean {
  const u = s.usedRange(false);
  if (!u || c < u.c1 || c > u.c2) return true;
  for (let r = Math.max(r1, u.r1); r <= Math.min(r2, u.r2); r++) if (hasContent(s.get(r, c))) return false;
  return true;
}

// --------------------------------------------------------------------------------------------
// Protection and structure

/**
 * The workbook a formula links to ([Book.xlsx] in [Book.xlsx]Sheet1!A1), or null. Brackets inside text and
 * the brackets of a table reference (Table1[Qty]) are not links: a link's bracket is followed by a sheet name and "!".
 */
export function externalLinkIn(formula: string): string | null {
  const m = /(?:^|[^\w\]])'?(\[[^\[\]]+\])[^!'\[\]]*'?!/.exec(formula.replace(/"(?:[^"]|"")*"/g, '""'));
  return m ? m[1] : null;
}

export function sheetProtected(app: App, action: string, sheet: Sheet = app.sheet): boolean {
  if (sheet.protection.on && !sheet.protection.allow.includes(action)) { toast(`The sheet "${sheet.name}" is protected. Unprotect it (Review ▸ Protect Sheet) to make this change.`); return true; }
  return false;
}

export function structureLocked(app: App): boolean {
  if (app.book.structureLocked) { toast('The workbook structure is protected (Review ▸ Protect Workbook).'); return true; }
  return false;
}

export function emptySheet(app: App, sheet: Sheet = app.sheet): boolean {
  if (!sheet.usedRange(false)) { toast('The sheet is empty. Enter or paste some data first.'); return true; }
  return false;
}

/** Inserts (n > 0) or deletes (n < 0) rows at `at`, keeping formulas, names, charts and so on in step. */
export function insertRows(sheet: Sheet, at: number, n: number): void { shiftCells(sheet, 'row', at, n); }
export function insertCols(sheet: Sheet, at: number, n: number): void { shiftCells(sheet, 'col', at, n); }

/** Deletes rows given as descending-safe spans (any order). */
export function deleteRowSpans(sheet: Sheet, spans: [number, number][]): number {
  let n = 0;
  for (const [a, b] of mergeSpans(spans).reverse()) { shiftCells(sheet, 'row', a, -(b - a + 1)); n += b - a + 1; }
  return n;
}

export function deleteColSpans(sheet: Sheet, spans: [number, number][]): number {
  let n = 0;
  for (const [a, b] of mergeSpans(spans).reverse()) { shiftCells(sheet, 'col', a, -(b - a + 1)); n += b - a + 1; }
  return n;
}

// --------------------------------------------------------------------------------------------
// Values

export const numOf = (v: Scalar | undefined): number | null => (typeof v === 'number' ? v : typeof v === 'string' && v.trim() !== '' && !isNaN(Number(v)) ? Number(v) : null);
export const textOf = (v: Scalar | undefined): string => (v == null ? '' : v instanceof CellError ? v.code : String(v));

export function plural(n: number, one: string, many = one + 's'): string { return `${n.toLocaleString()} ${n === 1 ? one : many}`; }

/** Applies fn to the constants (and optionally formulas' texts) of the selection as one undoable step. */
export function changeCells(app: App, label: string, fn: (cell: Cell, r: number, c: number) => Cell | undefined | null, opts: { formulas?: boolean; blanks?: boolean; sheet?: Sheet } = {}): number {
  const s = opts.sheet || app.sheet;
  if (!checkWritable(app, selRanges(app), s)) return 0;
  let n = 0, locked = 0;
  app.book.transact(label, () => {
    for (const x of selCells(app, { sheet: s })) {
      const cell = x.cell;
      if (!cell && !opts.blanks) continue;
      if (cell && cell.sp != null) continue;
      if (cell && cell.f != null && !opts.formulas) continue;
      if (!canChange(s, x.r, x.c)) { locked++; continue; }
      const out = fn(cell || {}, x.r, x.c);
      if (out === undefined) continue;
      s.set(x.r, x.c, out || undefined);
      n++;
    }
  });
  toast(n ? `${label}: ${plural(n, 'cell')} changed${locked ? ` (${locked} locked skipped)` : ''}` : `${label}: nothing to change in the selection`);
  return n;
}

/** Sets the style of every selected cell (existing or not) through a patch. */
export function patchStyle(app: App, label: string, patch: Partial<CellStyle> | ((st: CellStyle, r: number, c: number) => Partial<CellStyle> | null), opts: { sheet?: Sheet; onlyExisting?: boolean } = {}): number {
  const s = opts.sheet || app.sheet;
  if (sheetProtected(app, 'formatCells', s)) return 0;
  let n = 0;
  app.book.transact(label, () => {
    for (const x of selCells(app, { sheet: s, existingOnly: !!opts.onlyExisting })) {
      const st = s.styleOf(x.r, x.c);
      const p = typeof patch === 'function' ? patch(st, x.r, x.c) : patch;
      if (!p) continue;
      const id = s.book.derivedStyle(s.styleIdOf(x.r, x.c), p);
      if (id === (x.cell?.s ?? s.styleIdOf(x.r, x.c)) && x.cell?.s != null) continue;
      s.set(x.r, x.c, { ...(x.cell || {}), s: id });
      n++;
    }
  });
  return n;
}

/** Puts a report on a new sheet after the active one (header row styled). */
export function report(app: App, name: string, header: string[], rows: OutCell[][], opts: { colW?: number[]; freeze?: boolean } = {}): Sheet {
  const all: OutCell[][] = [header.map((t) => ({ v: t, style: HEAD })), ...rows];
  return addSheetWithRows(app, name, all, { colW: opts.colW || header.map((_h, i) => (i === 0 ? 140 : 120)), freeze: opts.freeze === false ? 0 : 1 });
}

export { writeOut, HEAD, addSheetWithRows, checkWritable, canChange, type OutCell };

// --------------------------------------------------------------------------------------------
// Sheets

export function visibleSheets(book: Workbook): Sheet[] { return book.sheets.filter((s) => s.visibility === 'visible'); }

/** Asks which sheets to work on (all visible by default). Returns null when cancelled. */
export async function pickSheets(app: App, title: string, opts: { others?: boolean; help?: string; dflt?: 'all' | 'active' | 'none'; extra?: import('../ui/dialog').Field[] } = {}): Promise<{ sheets: Sheet[]; values: Record<string, any> } | null> {
  const book = app.book;
  const list = opts.others ? book.sheets.filter((s) => s !== app.sheet) : book.sheets;
  if (!list.length) { toast('There are no other sheets.'); return null; }
  const dflt = opts.dflt === 'none' ? [] : opts.dflt === 'active' ? [app.sheet.name] : list.map((s) => s.name);
  const r = await openDialog({ title, help: opts.help, fields: [{ key: 'sheets', label: 'Sheets', type: 'list', multiple: true, rows: Math.min(10, Math.max(4, list.length)), options: list.map((s) => s.name + (s.visibility !== 'visible' ? ' (hidden)' : '')), value: dflt.map((n) => { const s = book.sheetByName(n)!; return s.name + (s.visibility !== 'visible' ? ' (hidden)' : ''); }) }, ...(opts.extra || [])] });
  if (!r) return null;
  const chosen: string[] = Array.isArray(r.sheets) ? r.sheets : r.sheets ? [r.sheets] : [];
  const sheets = chosen.map((n) => book.sheetByName(n.replace(/ \(hidden\)$/, ''))!).filter(Boolean);
  return { sheets, values: r };
}

export function sheetsByNames(book: Workbook, names: string[] | undefined): Sheet[] {
  if (!names) return book.sheets.slice();
  return names.map((n) => book.sheetByName(n)!).filter(Boolean);
}

/** Resolves text like "Sheet2!A1:C10" or "B2:B9" to a sheet and a range. */
export function rangeFromText(app: App, text: string): { sheet: Sheet; range: Range } | null {
  const t = text.trim();
  const m = /^(?:'([^']+)'|([^!]+))!(.+)$/.exec(t);
  if (m) { const sh = app.book.sheetByName(m[1] || m[2]); const rg = parseRange(m[3]); return sh && rg ? { sheet: sh, range: rg } : null; }
  const rg = parseRange(t);
  return rg ? { sheet: app.sheet, range: rg } : null;
}

// --------------------------------------------------------------------------------------------
// Clipboard and files

export async function copyText(text: string, what = 'Text'): Promise<boolean> {
  const { writeClipboardText } = await import('../platform/native');
  const ok = await writeClipboardText(text).catch(() => false);
  toast(ok ? `${what} copied to the clipboard` : `${what} could not be copied.`);
  return ok;
}

export async function saveTextFile(text: string, name: string, mime = 'text/plain'): Promise<void> {
  const { saveAs } = await import('../platform/native');
  await saveAs(new Blob([text], { type: mime }), name);
}

export async function pickFiles(accept: string, multiple = true): Promise<File[]> {
  const { pickFile } = await import('../ui/dom');
  return pickFile(accept, multiple);
}

// --------------------------------------------------------------------------------------------
// Formatting and text

export function cellStyleOf(app: App, r: number, c: number, sheet: Sheet = app.sheet): CellStyle { return sheet.styleOf(r, c); }

export function sameStyle(a: CellStyle, b: CellStyle, keys?: (keyof CellStyle)[]): boolean {
  const ks = keys || (Array.from(new Set([...Object.keys(a), ...Object.keys(b)])) as (keyof CellStyle)[]);
  return ks.every((k) => JSON.stringify(a[k] ?? null) === JSON.stringify(b[k] ?? null));
}

export const normColor = (c: string | undefined | null): string => (c || '').toLowerCase();

/** The text shown in a cell (number format applied). */
export function shownText(sheet: Sheet, r: number, c: number): string { return cellText(sheet, r, c, true); }

export function describeRange(rg: Range): string { return rangeName(rg); }

export { cellKey, keyRow, keyCol, colName, cellName, rangeName, MAX_ROWS, MAX_COLS, selRanges, activeRange, activeCell, clipToUsed, currentRegion };
export type { Range, Cell, CellStyle, Scalar, Sheet, Workbook, App };

/** Takes arg when a command was run from a macro, otherwise asks. */
export async function argOrAsk<T>(arg: T | undefined, ask: () => Promise<T | null>): Promise<T | null> {
  return arg !== undefined && arg !== null ? arg : ask();
}

/** Column letters or header text for choosing columns in dialogs. */
export function columnChoices(sheet: Sheet, rg: Range, header: boolean): { value: string; label: string }[] {
  const out: { value: string; label: string }[] = [];
  for (let c = rg.c1; c <= rg.c2; c++) { const t = header ? shownText(sheet, rg.r1, c) : ''; out.push({ value: String(c), label: t ? `${colName(c)}: ${t}` : `Column ${colName(c)}` }); }
  return out;
}

export function headerGuess(sheet: Sheet, rg: Range): boolean {
  if (rg.r2 <= rg.r1) return false;
  let text = 0, other = 0;
  for (let c = rg.c1; c <= rg.c2; c++) { const v = sheet.value(rg.r1, c); if (typeof v === 'string' && v !== '') text++; else if (v != null && v !== '') other++; }
  if (!text || other) return false;
  for (let c = rg.c1; c <= rg.c2; c++) if (typeof sheet.value(rg.r1 + 1, c) === 'number') return true;
  return sheet.styleOf(rg.r1, rg.c1).bold === true && !sheet.styleOf(rg.r1 + 1, rg.c1).bold;
}

export function escapeHtml(s: string): string { return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }

export function toCsvField(s: string, sep = ','): string { return /["\n\r]/.test(s) || s.includes(sep) ? `"${s.replace(/"/g, '""')}"` : s; }

/** Rows of displayed text for a block. */
export function textRows(sheet: Sheet, rg: Range, formatted = true): string[][] {
  const out: string[][] = [];
  for (let r = rg.r1; r <= rg.r2; r++) { const row: string[] = []; for (let c = rg.c1; c <= rg.c2; c++) row.push(cellText(sheet, r, c, formatted)); out.push(row); }
  return out;
}

export function valueRows(sheet: Sheet, rg: Range): Scalar[][] {
  const out: Scalar[][] = [];
  for (let r = rg.r1; r <= rg.r2; r++) { const row: Scalar[] = []; for (let c = rg.c1; c <= rg.c2; c++) row.push(sheet.value(r, c)); out.push(row); }
  return out;
}

export function fileStem(name: string): string { return name.replace(/^.*[\\/]/, '').replace(/\.[^.]+$/, ''); }

export function nowStamp(): string { const d = new Date(); const p = (n: number) => String(n).padStart(2, '0'); return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}${p(d.getMinutes())}`; }

/** Counts cells with content inside a range (sparse). */
export function countContent(sheet: Sheet, rg: Range): number { let n = 0; sheet.forEachIn(rg, (_r, _c, cell) => { if (hasContent(cell)) n++; }); return n; }

/** Writes a block of values/formulas with the top-left at (r0, c0). */
export function writeBlock(sheet: Sheet, r0: number, c0: number, rows: OutCell[][]): void {
  rows.forEach((row, i) => row.forEach((v, j) => writeOut(sheet, r0 + i, c0 + j, v)));
}

export const isWholeRowSel = (rg: Range): boolean => rg.c1 === 0 && rg.c2 >= MAX_COLS - 1;
export const isWholeColSel = (rg: Range): boolean => rg.r1 === 0 && rg.r2 >= MAX_ROWS - 1;

// --------------------------------------------------------------------------------------------
// Blocks of cells: copying, reordering

import { shiftFormula } from '../calc/parser';

export function blockOf(s: Sheet, rg: Range): (Cell | undefined)[][] {
  const m: (Cell | undefined)[][] = [];
  for (let r = rg.r1; r <= rg.r2; r++) { const row: (Cell | undefined)[] = []; for (let c = rg.c1; c <= rg.c2; c++) { const x = s.get(r, c); row.push(x ? { ...x } : undefined); } m.push(row); }
  return m;
}

/** A copy of a cell as it would be after moving by (dr, dc): relative references follow. */
export function movedCell(cell: Cell | undefined, dr: number, dc: number): Cell | undefined {
  if (!cell) return undefined;
  const x: Cell = { ...cell };
  delete x.sp;
  if (x.f != null) x.f = shiftFormula(x.f, dr, dc);
  return x;
}

export function hasMergeIn(s: Sheet, rg: Range): boolean { return s.merges.some((m) => m.r1 <= rg.r2 && m.r2 >= rg.r1 && m.c1 <= rg.c2 && m.c2 >= rg.c1); }

/** Rewrites the rows of a block in a new order (order[i] = source row index). Call inside a transaction. */
export function reorderRows(s: Sheet, rg: Range, order: number[]): void {
  const m = blockOf(s, rg);
  order.forEach((src, i) => m[src].forEach((cell, j) => s.set(rg.r1 + i, rg.c1 + j, movedCell(cell, i - src, 0))));
  // Row heights and hidden flags travel with their rows.
  const heights = order.map((src) => s.rowH.get(rg.r1 + src));
  const hidden = order.map((src) => s.hiddenRows.has(rg.r1 + src));
  order.forEach((_src, i) => { const r = rg.r1 + i; if (heights[i] != null) s.rowH.set(r, heights[i]!); else s.rowH.delete(r); if (hidden[i]) s.hiddenRows.add(r); else s.hiddenRows.delete(r); });
}

export function reorderCols(s: Sheet, rg: Range, order: number[]): void {
  const m = blockOf(s, rg);
  order.forEach((src, j) => { for (let i = 0; i < m.length; i++) s.set(rg.r1 + i, rg.c1 + j, movedCell(m[i][src], 0, j - src)); });
  const widths = order.map((src) => s.colW.get(rg.c1 + src));
  const hidden = order.map((src) => s.hiddenCols.has(rg.c1 + src));
  order.forEach((_src, j) => { const c = rg.c1 + j; if (widths[j] != null) s.colW.set(c, widths[j]!); else s.colW.delete(c); if (hidden[j]) s.hiddenCols.add(c); else s.hiddenCols.delete(c); });
}

/** Sort key text for "natural" ordering (numbers inside text compare by value). */
export function naturalKey(s: string): (string | number)[] {
  return s.toLowerCase().split(/(\d+(?:\.\d+)?)/).filter((x) => x !== '').map((x) => (/^\d/.test(x) ? Number(x) : x));
}

export function naturalCompare(a: string, b: string): number {
  const ka = naturalKey(a), kb = naturalKey(b);
  for (let i = 0; i < Math.min(ka.length, kb.length); i++) {
    const x = ka[i], y = kb[i];
    if (typeof x === 'number' && typeof y === 'number') { if (x !== y) return x - y; }
    else { const c = String(x).localeCompare(String(y)); if (c) return c; }
  }
  return ka.length - kb.length;
}

/** Centimeters/inches/points to pixels at 96 dpi and back. */
export const PX_PER_CM = 96 / 2.54;
export const PX_PER_IN = 96;
export const PX_PER_PT = 96 / 72;
