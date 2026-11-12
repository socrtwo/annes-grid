/**
 * Utilities tab, part 1: Display & Navigation, Select, Sheets. The tab clones the tool catalog of the
 * ASAP Utilities add-in for Excel (nominative use: the tools are re-implemented from their public
 * descriptions; nothing is copied), grouped under the same category names so that anyone used to that
 * add-in finds the tools where they expect them. Each tool accepts an `arg` so macros can run it
 * without the dialog.
 */
import type { App } from '../app';
import { Sheet } from '../model/workbook';
import type { Cell, Scalar } from '../model/types';
import { CellError } from '../model/types';
import { cellKey, keyRow, keyCol, colName, cellName, rangeName, parseRange, MAX_ROWS, MAX_COLS, type Range } from '../model/address';
import { isDateFormat } from '../model/numfmt';
import { renameSheet, validSheetName } from '../model/structure';
import { selRanges, activeRange, activeCell, clipToUsed, currentRegion, isWholeCols, isWholeRows } from './ops';
import { defTools, U, selCells, selBlock, singleCell, hasContent, selectHits, selectWhere, selectRange, rowSpans, colSpans, sheetProtected, structureLocked, emptySheet, numOf, textOf, plural, report, visibleSheets, pickSheets, sheetsByNames, normColor, sameStyle, shownText, argOrAsk, columnChoices, headerGuess, naturalCompare, writeOut, HEAD, toolLog, copyText } from './utilShared';
import { openDialog, toast, confirmDialog, promptDialog, choiceDialog } from '../ui/dialog';
import { getCommand, type MenuEntry } from './registry';
import { setPrint } from './pagelayout';
import { hashPassword, verifyPassword, PROTECT_OPTIONS } from './review';
import { criteria as criteriaFn } from '../calc/values';

const G_DISPLAY = 'Display & Navigation';
const G_SELECT = 'Select';
const G_SHEETS = 'Sheets';

// ============================================================================================
// Display & Navigation

const zoomMemory = { a: 1, b: 0.75 };

function setSheetFlags(app: App, sheets: Sheet[], label: string, fn: (s: Sheet) => void): void {
  app.book.transact(label, () => { for (const s of sheets) { app.book.recordSheet(s); fn(s); } });
  app.grid.layout();
}

export function resetLastCell(sheet: Sheet): { cells: number; range: Range | null } {
  const content = sheet.usedRange(false);
  let n = 0;
  for (const [k, cell] of [...sheet.cells]) {
    if (hasContent(cell) || cell.link || cell.rich) continue;
    const r = keyRow(k), c = keyCol(k);
    if (!content || r > content.r2 || c > content.c2) { sheet.set(r, c, undefined); n++; }
  }
  if (content) {
    for (const r of [...sheet.rowH.keys()]) if (r > content.r2 && !sheet.hiddenRows.has(r)) sheet.rowH.delete(r);
    for (const c of [...sheet.colW.keys()]) if (c > content.c2 && !sheet.hiddenCols.has(c)) sheet.colW.delete(c);
  }
  return { cells: n, range: sheet.usedRange() };
}

defTools([
  {
    id: 'util.displaySettings', label: 'Display Settings…', icon: 'tune', tab: U, group: G_DISPLAY, noRecord: true, popular: true,
    tip: 'Turn gridlines, headings, zero values, formulas, page breaks, the formula bar and the Quick Access Bar on or off, and set the zoom, for this sheet or every sheet, in one dialog',
    keywords: 'vision control view options gridlines headings zoom', origin: 'ASAP Utilities (Vision Control)',
    run: async (app, arg?: Record<string, any>) => {
      const s = app.sheet;
      const o = await argOrAsk(arg, () => openDialog({ title: 'Display Settings', fields: [
        { key: 'grid', label: 'Gridlines', type: 'checkbox', value: s.showGrid, half: true }, { key: 'head', label: 'Row and column headings', type: 'checkbox', value: s.showHeadings, half: true },
        { key: 'zeros', label: 'Zero values', type: 'checkbox', value: s.showZeros, half: true }, { key: 'formulas', label: 'Show formulas instead of results', type: 'checkbox', value: s.showFormulas, half: true },
        { key: 'breaks', label: 'Page breaks', type: 'checkbox', value: !!(app.grid as any).showPageBreaks, half: true }, { key: 'bar', label: 'Formula bar', type: 'checkbox', value: app.prefs.showFormulaBar, half: true },
        { key: 'qat', label: 'Quick Access Bar', type: 'checkbox', value: !app.prefs.qatHidden, half: true }, { key: 'zoom', label: 'Zoom (%)', type: 'number', value: Math.round(s.zoom * 100), min: 10, max: 400, half: true },
        { key: 'all', label: 'Apply to every sheet', type: 'checkbox', value: false },
      ] }));
      if (!o) return;
      const sheets = o.all ? app.book.sheets : [s];
      setSheetFlags(app, sheets, 'Display Settings', (sh) => { sh.showGrid = !!o.grid; sh.showHeadings = !!o.head; sh.showZeros = !!o.zeros; sh.showFormulas = !!o.formulas; if (o.zoom) sh.zoom = Math.max(0.1, Math.min(4, Number(o.zoom) / 100)); });
      (app.grid as any).showPageBreaks = !!o.breaks;
      app.prefs.showFormulaBar = !!o.bar; app.prefs.qatHidden = !o.qat; app.savePrefs(); app.ui.applyPrefs(); app.refresh();
    },
  },
  {
    id: 'util.toggleZoom', label: 'Toggle Zoom', icon: 'zoom_in_map', tab: U, group: G_DISPLAY, noRecord: true, keys: ['Mod-Alt-z'],
    tip: 'Switch between two zoom levels (100% and 75% at first; set them with Toggle Zoom Levels…)', origin: 'ASAP Utilities',
    run: (app) => { const z = app.sheet.zoom; const target = Math.abs(z - zoomMemory.a) < 0.01 ? zoomMemory.b : zoomMemory.a; app.grid.setZoom(target); },
  },
  {
    id: 'util.toggleZoomLevels', label: 'Toggle Zoom Levels…', icon: 'settings_overscan', tab: U, group: G_DISPLAY, noRecord: true, tip: 'Set the two zoom levels that Toggle Zoom switches between',
    run: async (app, arg?: { a: number; b: number }) => {
      const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Toggle Zoom Levels', fields: [{ key: 'a', label: 'First zoom (%)', type: 'number', value: Math.round(zoomMemory.a * 100), min: 10, max: 400, half: true }, { key: 'b', label: 'Second zoom (%)', type: 'number', value: Math.round(zoomMemory.b * 100), min: 10, max: 400, half: true }] }); return r ? { a: Number(r.a), b: Number(r.b) } : null; });
      if (!o) return;
      zoomMemory.a = Math.max(10, Math.min(400, o.a)) / 100; zoomMemory.b = Math.max(10, Math.min(400, o.b)) / 100;
      app.grid.setZoom(zoomMemory.a);
    },
  },
  {
    id: 'util.zoomAllSheets', label: 'Same Zoom on All Sheets', icon: 'zoom_out_map', tab: U, group: G_DISPLAY, noRecord: true, tip: 'Give every sheet the zoom of the active sheet', origin: 'ASAP Utilities',
    run: (app) => { const z = app.sheet.zoom; for (const s of app.book.sheets) s.zoom = z; app.grid.layout(); toast(`All sheets set to ${Math.round(z * 100)}%`); },
  },
  {
    id: 'util.sameViewAllSheets', label: 'Same View Settings on All Sheets', icon: 'select_window', tab: U, group: G_DISPLAY, tip: 'Copy gridlines, headings, zero values, zoom and gridline color from the active sheet to every sheet', origin: 'ASAP Utilities',
    run: (app) => { const a = app.sheet; setSheetFlags(app, app.book.sheets.filter((s) => s !== a), 'Same View Settings', (s) => { s.showGrid = a.showGrid; s.showHeadings = a.showHeadings; s.showZeros = a.showZeros; s.zoom = a.zoom; s.gridColor = a.gridColor; s.showFormulas = a.showFormulas; }); toast('View settings copied to every sheet'); },
  },
  { id: 'util.toggleCalcMode', label: 'Toggle Automatic/Manual Calculation', icon: 'calculate', tab: U, group: G_DISPLAY, kind: 'toggle', noRecord: true, tip: 'Switch recalculation between automatic and manual', origin: 'ASAP Utilities', run: (app) => { const auto = app.book.calc.mode === 'auto'; app.book.calc.mode = auto ? 'manual' : 'auto'; if (!auto) app.calculate(); toast(auto ? 'Calculation is now manual (F9 recalculates)' : 'Calculation is now automatic'); }, isActive: (app) => app.book.calc.mode === 'manual' },
  { id: 'util.firstSheet', label: 'Go to First Sheet', icon: 'first_page', tab: U, group: G_DISPLAY, noRecord: true, keys: ['Mod-Alt-PageUp'], run: (app) => { const v = visibleSheets(app.book); if (v.length) app.exec('sheet.activate', app.book.sheets.indexOf(v[0])); } },
  { id: 'util.lastSheet', label: 'Go to Last Sheet', icon: 'last_page', tab: U, group: G_DISPLAY, noRecord: true, keys: ['Mod-Alt-PageDown'], run: (app) => { const v = visibleSheets(app.book); if (v.length) app.exec('sheet.activate', app.book.sheets.indexOf(v[v.length - 1])); } },
  { id: 'util.previousSheetVisited', label: 'Back to Previous Sheet', icon: 'undo', tab: U, group: G_DISPLAY, noRecord: true, tip: 'Return to the sheet that was active before this one', run: (app) => { const prev = sheetHistory[sheetHistory.length - 2]; if (prev == null) { toast('No previous sheet yet.'); return; } const i = app.book.sheets.findIndex((s) => s.id === prev); if (i >= 0) { sheetHistory.pop(); app.exec('sheet.activate', i); } } },
  {
    id: 'util.resetLastCell', label: 'Reset the Last Cell (Trim Used Range)', icon: 'crop', tab: U, group: G_DISPLAY, edits: true, popular: true,
    tip: 'Remove the empty formatted cells beyond your data so the scrollbars, Ctrl+End and the file size reflect the real sheet', keywords: 'used range last cell ghost formatting file size', origin: 'ASAP Utilities (Reset Excel\'s last cell)',
    run: (app, arg?: { all?: boolean }) => {
      const sheets = arg?.all ? app.book.sheets : [app.sheet];
      let total = 0;
      app.book.transact('Reset the Last Cell', () => { for (const s of sheets) { app.book.recordSheet(s); total += resetLastCell(s).cells; } });
      app.grid.updateExtent(); app.grid.layout();
      toast(total ? `${plural(total, 'empty formatted cell')} removed; the last cell is now ${app.sheet.usedRange() ? cellName(app.sheet.usedRange()!.r2, app.sheet.usedRange()!.c2) : 'A1'}` : 'The used range already ends at the data.');
    },
  },
  { id: 'util.resetLastCellAll', label: 'Reset the Last Cell on All Sheets', icon: 'crop_free', tab: U, group: G_DISPLAY, edits: true, run: (app) => app.exec('util.resetLastCell', { all: true }) },
  { id: 'util.scrollToTop', label: 'Scroll to the Top-Left (Keep Selection)', icon: 'north_west', tab: U, group: G_DISPLAY, noRecord: true, run: (app) => { app.grid.scroller.scrollTop = 0; app.grid.scroller.scrollLeft = 0; app.grid.layout(); } },
  { id: 'util.centerActiveCell', label: 'Center the Active Cell on Screen', icon: 'center_focus_strong', tab: U, group: G_DISPLAY, noRecord: true, run: (app) => { const { r, c } = activeCell(app); const g = app.grid; const vis = g.visible(); const dr = Math.floor((vis.r2 - vis.r1) / 2), dc = Math.floor((vis.c2 - vis.c1) / 2); g.scrollIntoView(Math.max(0, r - dr), Math.max(0, c - dc)); g.scrollIntoView(Math.min(MAX_ROWS - 1, r + dr), Math.min(MAX_COLS - 1, c + dc)); } },
  { id: 'util.hideAllObjects', label: 'Hide/Show All Objects', icon: 'hide_image', tab: U, group: G_DISPLAY, kind: 'toggle', noRecord: true, tip: 'Temporarily hide the charts, pictures and shapes of the sheet (they are not deleted)', origin: 'ASAP Utilities', run: (app) => { const g = app.grid as any; g.objectsHidden = !g.objectsHidden; document.body.classList.toggle('ql-objects-hidden', !!g.objectsHidden); app.grid.invalidate(); toast(g.objectsHidden ? 'Objects hidden (run again to show them)' : 'Objects shown'); }, isActive: (app) => !!(app.grid as any).objectsHidden },
  {
    id: 'util.windowList', label: 'Workbook List…', icon: 'view_list', tab: U, group: G_DISPLAY, noRecord: true, tip: 'Switch to another open workbook from a list that shows its sheet count and whether it has unsaved changes',
    run: async (app) => { const i = await choiceDialog('Workbooks', 'Switch to:', app.docs.map((d) => `${d.displayName}${d.dirty ? ' •' : ''} (${d.book.sheets.length} sheet${d.book.sheets.length === 1 ? '' : 's'})`)); if (i >= 0) app.activate(app.docs[i]); },
  },
  { id: 'util.menuDisplay', label: 'Display & Navigation', icon: 'visibility', tab: U, group: G_DISPLAY, kind: 'menu', noRecord: true, run: () => undefined, menu: () => menuOf(['util.displaySettings', 'util.toggleZoom', 'util.toggleZoomLevels', 'util.zoomAllSheets', 'util.sameViewAllSheets', null, 'view.gridlines', 'view.headings', 'view.showFormulas', 'view.showZeros', 'view.showPageBreaks', 'view.formulaBar', 'view.fullScreen', 'util.toggleCalcMode', 'util.hideAllObjects', null, 'util.firstSheet', 'util.lastSheet', 'util.previousSheetVisited', 'sheet.goto', 'util.windowList', 'view.switchWindows', null, 'util.resetLastCell', 'util.resetLastCellAll', 'util.scrollToTop', 'util.centerActiveCell', 'nav.activeCell', 'nav.end']) },
]);

const sheetHistory: string[] = [];
export function noteSheetVisit(id: string): void { if (sheetHistory[sheetHistory.length - 1] !== id) { sheetHistory.push(id); if (sheetHistory.length > 50) sheetHistory.shift(); } }

/** Menu entries for a list of command ids (null = separator); unknown ids are skipped. */
export function menuOf(ids: (string | null)[]): MenuEntry[] {
  const out: MenuEntry[] = [];
  for (const id of ids) {
    if (id === null) { if (out.length && !out[out.length - 1].separator) out.push({ separator: true }); continue; }
    const c = getCommand(id);
    if (c) out.push({ id, label: c.label, icon: c.icon });
  }
  return out;
}

// ============================================================================================
// Select

export type ContentTest = 'equals' | 'contains' | 'starts' | 'ends' | 'greater' | 'less' | 'between' | 'blank' | 'nonblank' | 'number' | 'text' | 'date' | 'bool' | 'error' | 'formula' | 'constant' | 'longer' | 'shorter' | 'regex' | 'negative' | 'positive' | 'duplicate' | 'unique';
export type FormatTest = 'bold' | 'italic' | 'underline' | 'strike' | 'fill' | 'fontColor' | 'locked' | 'unlocked' | 'merged' | 'wrapped' | 'comment' | 'link' | 'numFmt' | 'font' | 'size' | 'any';

export interface SelectCriteria { content?: ContentTest; value?: string; value2?: string; matchCase?: boolean; format?: FormatTest; color?: string; text?: string }

/** Whether a cell passes the content/format criteria of Select by Content. */
export function cellMatches(sheet: Sheet, r: number, c: number, cell: Cell | undefined, crit: SelectCriteria, ctx: { dupes?: Map<string, number> } = {}): boolean {
  const v = cell?.v;
  const shown = cell ? shownText(sheet, r, c) : '';
  const txt = crit.matchCase ? shown : shown.toLowerCase();
  const needle = crit.matchCase ? crit.value || '' : (crit.value || '').toLowerCase();
  const n = numOf(v);
  const a = Number(crit.value), b = Number(crit.value2);
  const date = typeof v === 'number' && isDateFormat(sheet.styleOf(r, c).numFmt);
  switch (crit.content) {
    case undefined: break;
    case 'equals': if (!(txt === needle || (n != null && !isNaN(a) && n === a))) return false; break;
    case 'contains': if (!txt.includes(needle)) return false; break;
    case 'starts': if (!txt.startsWith(needle)) return false; break;
    case 'ends': if (!txt.endsWith(needle)) return false; break;
    case 'greater': if (n == null || isNaN(a) || n <= a) return false; break;
    case 'less': if (n == null || isNaN(a) || n >= a) return false; break;
    case 'between': if (n == null || isNaN(a) || isNaN(b) || n < Math.min(a, b) || n > Math.max(a, b)) return false; break;
    case 'blank': if (hasContent(cell)) return false; break;
    case 'nonblank': if (!hasContent(cell)) return false; break;
    case 'number': if (typeof v !== 'number' || date) return false; break;
    case 'text': if (typeof v !== 'string' || v === '') return false; break;
    case 'date': if (!date) return false; break;
    case 'bool': if (typeof v !== 'boolean') return false; break;
    case 'error': if (!(v instanceof CellError)) return false; break;
    case 'formula': if (cell?.f == null) return false; break;
    case 'constant': if (!cell || cell.f != null || cell.sp != null || v == null) return false; break;
    case 'longer': if (shown.length <= (Number(crit.value) || 0)) return false; break;
    case 'shorter': if (!shown || shown.length >= (Number(crit.value) || 0)) return false; break;
    case 'regex': try { if (!new RegExp(crit.value || '', crit.matchCase ? 'u' : 'iu').test(shown)) return false; } catch { return false; } break;
    case 'negative': if (typeof v !== 'number' || v >= 0) return false; break;
    case 'positive': if (typeof v !== 'number' || v <= 0) return false; break;
    case 'duplicate': if (!hasContent(cell) || (ctx.dupes?.get(txt) || 0) < 2) return false; break;
    case 'unique': if (!hasContent(cell) || (ctx.dupes?.get(txt) || 0) !== 1) return false; break;
  }
  if (crit.format && crit.format !== 'any') {
    const st = sheet.styleOf(r, c);
    const col = normColor(crit.color);
    switch (crit.format) {
      case 'bold': if (!st.bold) return false; break;
      case 'italic': if (!st.italic) return false; break;
      case 'underline': if (!st.underline) return false; break;
      case 'strike': if (!st.strike) return false; break;
      case 'fill': if (col ? normColor(st.fill) !== col : !st.fill) return false; break;
      case 'fontColor': if (col ? normColor(st.color) !== col : !st.color) return false; break;
      case 'locked': if (st.locked === false) return false; break;
      case 'unlocked': if (st.locked !== false) return false; break;
      case 'merged': if (!sheet.mergeAt(r, c)) return false; break;
      case 'wrapped': if (!st.wrap) return false; break;
      case 'comment': if (!sheet.comments.has(cellKey(r, c))) return false; break;
      case 'link': if (!cell?.link) return false; break;
      case 'numFmt': if ((st.numFmt || 'General') !== (crit.text || 'General')) return false; break;
      case 'font': if ((st.font || '').toLowerCase() !== (crit.text || '').toLowerCase()) return false; break;
      case 'size': if (String(st.size || '') !== String(crit.text || '')) return false; break;
    }
  }
  return true;
}

const CONTENT_OPTIONS: { value: ContentTest | ''; label: string }[] = [
  { value: '', label: '(any content)' }, { value: 'equals', label: 'Equal to' }, { value: 'contains', label: 'Contains' }, { value: 'starts', label: 'Starts with' }, { value: 'ends', label: 'Ends with' },
  { value: 'greater', label: 'Greater than' }, { value: 'less', label: 'Less than' }, { value: 'between', label: 'Between' }, { value: 'longer', label: 'Text longer than … characters' }, { value: 'shorter', label: 'Text shorter than … characters' }, { value: 'regex', label: 'Matches a regular expression' },
  { value: 'blank', label: 'Blank' }, { value: 'nonblank', label: 'Not blank' }, { value: 'number', label: 'Numbers' }, { value: 'text', label: 'Text' }, { value: 'date', label: 'Dates' }, { value: 'bool', label: 'TRUE/FALSE' }, { value: 'error', label: 'Errors' }, { value: 'formula', label: 'Formulas' }, { value: 'constant', label: 'Constants' },
  { value: 'negative', label: 'Negative numbers' }, { value: 'positive', label: 'Positive numbers' }, { value: 'duplicate', label: 'Duplicate values' }, { value: 'unique', label: 'Unique values' },
];
const FORMAT_OPTIONS: { value: FormatTest; label: string }[] = [
  { value: 'any', label: '(any formatting)' }, { value: 'bold', label: 'Bold' }, { value: 'italic', label: 'Italic' }, { value: 'underline', label: 'Underlined' }, { value: 'strike', label: 'Strikethrough' }, { value: 'fill', label: 'Fill color (below, or any)' }, { value: 'fontColor', label: 'Font color (below, or any)' },
  { value: 'locked', label: 'Locked' }, { value: 'unlocked', label: 'Unlocked' }, { value: 'merged', label: 'Merged' }, { value: 'wrapped', label: 'Wrap text' }, { value: 'comment', label: 'Has a comment or note' }, { value: 'link', label: 'Has a hyperlink' }, { value: 'numFmt', label: 'Number format (below)' }, { value: 'font', label: 'Font name (below)' }, { value: 'size', label: 'Font size (below)' },
];
const NEEDS_VALUE: ContentTest[] = ['equals', 'contains', 'starts', 'ends', 'greater', 'less', 'between', 'longer', 'shorter', 'regex'];

function dupeCounts(sheet: Sheet, cells: { r: number; c: number; cell: Cell | undefined }[], matchCase: boolean): Map<string, number> {
  const m = new Map<string, number>();
  for (const x of cells) { if (!hasContent(x.cell)) continue; const t = shownText(sheet, x.r, x.c); const k = matchCase ? t : t.toLowerCase(); m.set(k, (m.get(k) || 0) + 1); }
  return m;
}

/** Cells of the selection (or the sheet) that match the criteria. */
export function findMatching(app: App, crit: SelectCriteria, scope: 'selection' | 'sheet' = 'selection'): { r: number; c: number }[] {
  const s = app.sheet;
  const cells: { r: number; c: number; cell: Cell | undefined }[] = [];
  if (scope === 'sheet' || singleCell(app)) { const u = s.usedRange(); if (u) for (let r = u.r1; r <= u.r2; r++) for (let c = u.c1; c <= u.c2; c++) cells.push({ r, c, cell: s.get(r, c) }); }
  else cells.push(...selCells(app));
  const ctx = crit.content === 'duplicate' || crit.content === 'unique' ? { dupes: dupeCounts(s, cells, !!crit.matchCase) } : {};
  return cells.filter((x) => cellMatches(s, x.r, x.c, x.cell, crit, ctx)).map((x) => ({ r: x.r, c: x.c }));
}

async function selectByContentDialog(app: App, title: string, extra: import('../ui/dialog').Field[] = []): Promise<(SelectCriteria & { scope: 'selection' | 'sheet' } & Record<string, any>) | null> {
  const r = await openDialog({ title, fields: [
    { key: 'content', label: 'Content', type: 'select', value: 'nonblank', options: CONTENT_OPTIONS },
    { key: 'value', label: 'Value', type: 'text', value: '', visibleIf: (v) => NEEDS_VALUE.includes(v.content), half: true },
    { key: 'value2', label: 'And', type: 'text', value: '', visibleIf: (v) => v.content === 'between', half: true },
    { key: 'matchCase', label: 'Match case', type: 'checkbox', value: false },
    { key: 'format', label: 'Formatting', type: 'select', value: 'any', options: FORMAT_OPTIONS },
    { key: 'color', label: 'Color (empty = any color)', type: 'text', value: '', placeholder: '#ffff00', visibleIf: (v) => v.format === 'fill' || v.format === 'fontColor' },
    { key: 'text', label: 'Format / font / size', type: 'text', value: '', visibleIf: (v) => v.format === 'numFmt' || v.format === 'font' || v.format === 'size' },
    { key: 'scope', label: 'Look in', type: 'radio', value: singleCell(app) ? 'sheet' : 'selection', options: [{ value: 'selection', label: 'The selection' }, { value: 'sheet', label: 'The whole sheet' }] },
    ...extra,
  ] });
  if (!r) return null;
  return { ...r, content: r.content || undefined, format: r.format === 'any' ? undefined : r.format } as any;
}

/** The selection before the last Select tool ran (Reselect Previous Selection). */
const selectionStack: { sheet: string; r: number; c: number; ranges: Range[] }[] = [];
let remembered: { sheet: string; r: number; c: number; ranges: Range[] } | null = null;
function pushSelection(app: App): void { const s = app.sheet; selectionStack.push({ sheet: s.id, r: s.sel.r, c: s.sel.c, ranges: s.sel.ranges.map((x) => ({ ...x })) }); if (selectionStack.length > 20) selectionStack.shift(); }
function restoreSelection(app: App, sel: { sheet: string; r: number; c: number; ranges: Range[] }): void {
  const i = app.book.sheets.findIndex((x) => x.id === sel.sheet);
  if (i < 0) { toast('That sheet no longer exists.'); return; }
  if (i !== app.book.activeSheet) app.exec('sheet.activate', i);
  app.sheet.sel = { r: sel.r, c: sel.c, ranges: sel.ranges.map((x) => ({ ...x })) };
  app.grid.scrollIntoView(sel.r, sel.c); app.grid.invalidate(); app.refresh();
}

function selectSame(app: App, what: 'format' | 'fill' | 'fontColor' | 'value' | 'numFmt' | 'font', label: string): void {
  const s = app.sheet; const { r, c } = activeCell(app);
  const st = s.styleOf(r, c); const v = s.value(r, c); const shown = shownText(s, r, c);
  pushSelection(app);
  selectWhere(app, (cell, rr, cc) => {
    const t = s.styleOf(rr, cc);
    switch (what) {
      case 'format': return sameStyle(t, st);
      case 'fill': return normColor(t.fill) === normColor(st.fill) && !!(t.fill || st.fill) === !!st.fill;
      case 'fontColor': return normColor(t.color) === normColor(st.color);
      case 'numFmt': return (t.numFmt || 'General') === (st.numFmt || 'General');
      case 'font': return (t.font || '') === (st.font || '') && (t.size || 0) === (st.size || 0);
      default: return hasContent(cell) && (cell!.v === v || shownText(s, rr, cc) === shown);
    }
  }, { what: label, wholeSheet: true, existingOnly: what === 'value' });
}

function extendSelection(app: App, how: 'lastRow' | 'lastCol' | 'end' | 'colEnd' | 'rowEnd' | 'colStart' | 'rowStart'): void {
  const s = app.sheet; const u = s.usedRange(false);
  if (!u) { toast('The sheet is empty.'); return; }
  const rg = activeRange(app); const { r, c } = activeCell(app);
  let next: Range = { ...rg };
  if (how === 'lastRow') next.r2 = Math.max(rg.r1, u.r2);
  else if (how === 'lastCol') next.c2 = Math.max(rg.c1, u.c2);
  else if (how === 'end') next = { r1: rg.r1, c1: rg.c1, r2: Math.max(rg.r1, u.r2), c2: Math.max(rg.c1, u.c2) };
  else if (how === 'colEnd') { let last = r; for (let rr = r; rr <= u.r2; rr++) if (hasContent(s.get(rr, c))) last = rr; next = { r1: r, c1: rg.c1, r2: last, c2: rg.c2 }; }
  else if (how === 'rowEnd') { let last = c; for (let cc = c; cc <= u.c2; cc++) if (hasContent(s.get(r, cc))) last = cc; next = { r1: rg.r1, c1: c, r2: rg.r2, c2: last }; }
  else if (how === 'colStart') next = { r1: 0, c1: rg.c1, r2: r, c2: rg.c2 };
  else next = { r1: rg.r1, c1: 0, r2: rg.r2, c2: c };
  pushSelection(app);
  app.grid.select(r, c, { range: next, keepActive: true });
}

function invertSelection(app: App): void {
  const s = app.sheet; const u = s.usedRange();
  if (!u) { toast('The sheet is empty.'); return; }
  const inSel = (r: number, c: number) => selRanges(app).some((rg) => r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2);
  const hits: { r: number; c: number }[] = [];
  for (let r = u.r1; r <= u.r2; r++) for (let c = u.c1; c <= u.c2; c++) if (!inSel(r, c)) hits.push({ r, c });
  pushSelection(app);
  selectHits(app, hits, 'cells outside the selection');
}

/** Rows/columns matched by a criterion in a key column/row: select, hide or delete. */
async function conditionalLines(app: App, arg: any, axis: 'row' | 'col'): Promise<void> {
  const s = app.sheet;
  if (emptySheet(app)) return;
  const rg = selBlock(app);
  const header = headerGuess(s, rg);
  const o = await argOrAsk(arg, () => openDialog({ title: axis === 'row' ? 'Conditional Row Select, Hide or Delete' : 'Conditional Column Select, Hide or Delete', help: axis === 'row' ? `Rows ${rg.r1 + 1}–${rg.r2 + 1} are checked (select a range first to limit them).` : `Columns ${colName(rg.c1)}–${colName(rg.c2)} are checked.`, fields: [
    ...(axis === 'row' ? [{ key: 'key', label: 'Look at the value in column', type: 'select' as const, value: String(rg.c1), options: [{ value: '-1', label: 'Any cell of the row' }, ...columnChoices(s, rg, header)] }] : [{ key: 'key', label: 'Look at the value in row', type: 'number' as const, value: rg.r1 + 1, min: 1 }, { key: 'anyRow', label: 'Any cell of the column', type: 'checkbox' as const, value: false }]),
    { key: 'content', label: 'Condition', type: 'select', value: 'contains', options: CONTENT_OPTIONS.filter((x) => x.value) },
    { key: 'value', label: 'Value', type: 'text', value: '', visibleIf: (v) => NEEDS_VALUE.includes(v.content), half: true },
    { key: 'value2', label: 'And', type: 'text', value: '', visibleIf: (v) => v.content === 'between', half: true },
    { key: 'matchCase', label: 'Match case', type: 'checkbox', value: false },
    { key: 'action', label: 'Then', type: 'radio', value: 'select', options: [{ value: 'select', label: axis === 'row' ? 'Select the rows' : 'Select the columns' }, { value: 'hide', label: 'Hide them' }, { value: 'delete', label: 'Delete them' }, { value: 'color', label: 'Color them yellow' }] },
    { key: 'skipHeader', label: 'Skip the first row (header)', type: 'checkbox', value: header, visibleIf: () => axis === 'row' },
  ] }));
  if (!o) return;
  const crit: SelectCriteria = { content: o.content, value: o.value, value2: o.value2, matchCase: !!o.matchCase };
  const cellsOf = (line: number): { r: number; c: number }[] => {
    if (axis === 'row') { const k = Number(o.key); return k >= 0 ? [{ r: line, c: k }] : Array.from({ length: rg.c2 - rg.c1 + 1 }, (_x, i) => ({ r: line, c: rg.c1 + i })); }
    const k = Number(o.key) - 1; return o.anyRow ? Array.from({ length: rg.r2 - rg.r1 + 1 }, (_x, i) => ({ r: rg.r1 + i, c: line })) : [{ r: k, c: line }];
  };
  const all: { r: number; c: number; cell: Cell | undefined }[] = [];
  const first = axis === 'row' ? rg.r1 + (o.skipHeader ? 1 : 0) : rg.c1, last = axis === 'row' ? rg.r2 : rg.c2;
  for (let line = first; line <= last; line++) for (const p of cellsOf(line)) all.push({ r: p.r, c: p.c, cell: s.get(p.r, p.c) });
  const ctx = crit.content === 'duplicate' || crit.content === 'unique' ? { dupes: dupeCounts(s, all, !!crit.matchCase) } : {};
  const lines: number[] = [];
  for (let line = first; line <= last; line++) if (cellsOf(line).some((p) => cellMatches(s, p.r, p.c, s.get(p.r, p.c), crit, ctx))) lines.push(line);
  if (!lines.length) { toast(`No ${axis === 'row' ? 'rows' : 'columns'} match.`); return; }
  const spans = mergeSpansOf(lines);
  if (o.action === 'select') { pushSelection(app); s.sel.ranges = spans.map(([a, b]) => (axis === 'row' ? { r1: a, c1: 0, r2: b, c2: MAX_COLS - 1 } : { r1: 0, c1: a, r2: MAX_ROWS - 1, c2: b })); s.sel.r = axis === 'row' ? lines[0] : rg.r1; s.sel.c = axis === 'row' ? rg.c1 : lines[0]; app.grid.invalidate(); app.ui.status(`${plural(lines.length, axis === 'row' ? 'row' : 'column')} selected`); return; }
  if (o.action === 'hide') { if (sheetProtected(app, axis === 'row' ? 'formatRows' : 'formatColumns')) return; app.book.transact('Hide', () => { app.book.recordSheet(s); for (const l of lines) (axis === 'row' ? s.hiddenRows : s.hiddenCols).add(l); }); app.grid.layout(); toast(`${plural(lines.length, axis === 'row' ? 'row' : 'column')} hidden`); return; }
  if (o.action === 'color') { app.book.transact('Color Matching Lines', () => { for (const l of lines) { const cells = axis === 'row' ? Array.from({ length: rg.c2 - rg.c1 + 1 }, (_x, i) => ({ r: l, c: rg.c1 + i })) : Array.from({ length: rg.r2 - rg.r1 + 1 }, (_x, i) => ({ r: rg.r1 + i, c: l })); for (const p of cells) s.set(p.r, p.c, { ...(s.get(p.r, p.c) || {}), s: app.book.derivedStyle(s.styleIdOf(p.r, p.c), { fill: '#ffff00' }) }); } }); return; }
  if (sheetProtected(app, axis === 'row' ? 'deleteRows' : 'deleteColumns')) return;
  if (arg === undefined && !(await confirmDialog('Delete', `Delete ${plural(lines.length, axis === 'row' ? 'row' : 'column')}? This can be undone.`, 'Delete'))) return;
  const { deleteRowSpans, deleteColSpans } = await import('./utilShared');
  app.book.transact(axis === 'row' ? 'Delete Rows' : 'Delete Columns', () => { if (axis === 'row') deleteRowSpans(s, spans); else deleteColSpans(s, spans); });
  app.grid.updateExtent(); app.grid.layout();
  toast(`${plural(lines.length, axis === 'row' ? 'row' : 'column')} deleted`);
}

function mergeSpansOf(lines: number[]): [number, number][] { const out: [number, number][] = []; for (const l of [...lines].sort((a, b) => a - b)) { const last = out[out.length - 1]; if (last && last[1] + 1 === l) last[1] = l; else out.push([l, l]); } return out; }

function everyNth(app: App, lines: number[], axis: 'row' | 'col', action: string, label: string): void {
  const s = app.sheet;
  if (!lines.length) { toast('Nothing matched.'); return; }
  const spans = mergeSpansOf(lines);
  if (action === 'select') { pushSelection(app); s.sel.ranges = spans.map(([a, b]) => (axis === 'row' ? { r1: a, c1: 0, r2: b, c2: MAX_COLS - 1 } : { r1: 0, c1: a, r2: MAX_ROWS - 1, c2: b })); s.sel.r = axis === 'row' ? lines[0] : 0; s.sel.c = axis === 'row' ? 0 : lines[0]; app.grid.invalidate(); return; }
  if (action === 'hide') { app.book.transact(label, () => { app.book.recordSheet(s); for (const l of lines) (axis === 'row' ? s.hiddenRows : s.hiddenCols).add(l); }); app.grid.layout(); return; }
  if (action === 'delete') { if (sheetProtected(app, axis === 'row' ? 'deleteRows' : 'deleteColumns')) return; import('./utilShared').then(({ deleteRowSpans, deleteColSpans }) => { app.book.transact(label, () => { if (axis === 'row') deleteRowSpans(s, spans); else deleteColSpans(s, spans); }); app.grid.updateExtent(); app.grid.layout(); toast(`${plural(lines.length, axis === 'row' ? 'row' : 'column')} deleted`); }); return; }
}

defTools([
  {
    id: 'util.selectByContent', label: 'Select Cells by Content and Format…', icon: 'filter_alt', tab: U, group: G_SELECT, noRecord: true, popular: true,
    tip: 'Select cells by what they contain (equal to, contains, between, blank, numbers, text, dates, errors, duplicates…) and how they are formatted (bold, fill color, locked, merged…)',
    keywords: 'select cells criteria content format color', origin: 'ASAP Utilities (Select cells based on content, formatting and more)',
    run: async (app, arg?: SelectCriteria & { scope?: 'selection' | 'sheet' }) => {
      const o = await argOrAsk(arg, () => selectByContentDialog(app, 'Select Cells by Content and Format'));
      if (!o) return;
      pushSelection(app);
      selectHits(app, findMatching(app, o, o.scope || 'selection'), 'matching cells');
    },
  },
  { id: 'util.condRows', label: 'Conditional Row Select, Hide or Delete…', icon: 'table_rows', tab: U, group: G_SELECT, edits: true, popular: true, tip: 'Find the rows whose value in a column meets a condition, then select, hide, color or delete them', keywords: 'delete rows containing hide rows where', origin: 'ASAP Utilities', run: (app, arg) => conditionalLines(app, arg, 'row') },
  { id: 'util.condCols', label: 'Conditional Column Select, Hide or Delete…', icon: 'view_column', tab: U, group: G_SELECT, edits: true, tip: 'Find the columns whose value in a row meets a condition, then select, hide, color or delete them', origin: 'ASAP Utilities', run: (app, arg) => conditionalLines(app, arg, 'col') },
  { id: 'util.selectSameFormat', label: 'Select Cells with the Same Formatting', icon: 'format_paint', tab: U, group: G_SELECT, noRecord: true, tip: 'Select every cell on the sheet formatted exactly like the active cell', origin: 'ASAP Utilities', run: (app) => selectSame(app, 'format', 'cells with the same formatting') },
  { id: 'util.selectSameFill', label: 'Select Cells with the Same Fill Color', icon: 'format_color_fill', tab: U, group: G_SELECT, noRecord: true, origin: 'ASAP Utilities', run: (app) => selectSame(app, 'fill', 'cells with the same fill') },
  { id: 'util.selectSameFontColor', label: 'Select Cells with the Same Font Color', icon: 'format_color_text', tab: U, group: G_SELECT, noRecord: true, origin: 'ASAP Utilities', run: (app) => selectSame(app, 'fontColor', 'cells with the same font color') },
  { id: 'util.selectSameValue', label: 'Select Cells with the Same Value', icon: 'join_full', tab: U, group: G_SELECT, noRecord: true, tip: 'Select every cell on the sheet with the value of the active cell', origin: 'ASAP Utilities', run: (app) => selectSame(app, 'value', 'cells with the same value') },
  { id: 'util.selectSameNumFmt', label: 'Select Cells with the Same Number Format', icon: 'pin', tab: U, group: G_SELECT, noRecord: true, origin: 'ASAP Utilities', run: (app) => selectSame(app, 'numFmt', 'cells with the same number format') },
  { id: 'util.selectSameFont', label: 'Select Cells with the Same Font', icon: 'font_download', tab: U, group: G_SELECT, noRecord: true, run: (app) => selectSame(app, 'font', 'cells with the same font') },
  { id: 'util.selectErrors', label: 'Select Cells with Errors', icon: 'error', tab: U, group: G_SELECT, noRecord: true, run: (app) => { pushSelection(app); selectWhere(app, (cell) => !!cell && cell.v instanceof CellError, { what: 'cells with errors' }); } },
  { id: 'util.selectNumbers', label: 'Select Cells with Numbers', icon: '123', tab: U, group: G_SELECT, noRecord: true, run: (app) => { pushSelection(app); selectWhere(app, (cell, r, c) => !!cell && typeof cell.v === 'number' && !isDateFormat(app.sheet.styleOf(r, c).numFmt), { what: 'number cells' }); } },
  { id: 'util.selectText', label: 'Select Cells with Text', icon: 'abc', tab: U, group: G_SELECT, noRecord: true, run: (app) => { pushSelection(app); selectWhere(app, (cell) => !!cell && typeof cell.v === 'string' && cell.v !== '', { what: 'text cells' }); } },
  { id: 'util.selectDates', label: 'Select Cells with Dates', icon: 'event', tab: U, group: G_SELECT, noRecord: true, run: (app) => { pushSelection(app); selectWhere(app, (cell, r, c) => !!cell && typeof cell.v === 'number' && isDateFormat(app.sheet.styleOf(r, c).numFmt), { what: 'date cells' }); } },
  { id: 'util.selectBooleans', label: 'Select Cells with TRUE/FALSE', icon: 'toggle_on', tab: U, group: G_SELECT, noRecord: true, run: (app) => { pushSelection(app); selectWhere(app, (cell) => !!cell && typeof cell.v === 'boolean', { what: 'logical cells' }); } },
  { id: 'util.selectBlanks', label: 'Select Blank Cells', icon: 'check_box_outline_blank', tab: U, group: G_SELECT, noRecord: true, run: (app) => { pushSelection(app); selectWhere(app, (cell) => !hasContent(cell), { what: 'blank cells', existingOnly: false }); } },
  { id: 'util.selectNonBlanks', label: 'Select Non-Blank Cells', icon: 'check_box', tab: U, group: G_SELECT, noRecord: true, run: (app) => { pushSelection(app); selectWhere(app, (cell) => hasContent(cell), { what: 'non-blank cells' }); } },
  { id: 'util.selectNegative', label: 'Select Negative Numbers', icon: 'remove', tab: U, group: G_SELECT, noRecord: true, run: (app) => { pushSelection(app); selectWhere(app, (cell) => !!cell && typeof cell.v === 'number' && cell.v < 0, { what: 'negative numbers' }); } },
  { id: 'util.selectPositive', label: 'Select Positive Numbers', icon: 'add', tab: U, group: G_SELECT, noRecord: true, run: (app) => { pushSelection(app); selectWhere(app, (cell) => !!cell && typeof cell.v === 'number' && cell.v > 0, { what: 'positive numbers' }); } },
  { id: 'util.selectBold', label: 'Select Bold Cells', icon: 'format_bold', tab: U, group: G_SELECT, noRecord: true, run: (app) => { pushSelection(app); selectWhere(app, (_cell, r, c) => !!app.sheet.styleOf(r, c).bold, { what: 'bold cells', existingOnly: false }); } },
  { id: 'util.selectItalic', label: 'Select Italic Cells', icon: 'format_italic', tab: U, group: G_SELECT, noRecord: true, run: (app) => { pushSelection(app); selectWhere(app, (_cell, r, c) => !!app.sheet.styleOf(r, c).italic, { what: 'italic cells', existingOnly: false }); } },
  { id: 'util.selectLocked', label: 'Select Locked Cells', icon: 'lock', tab: U, group: G_SELECT, noRecord: true, run: (app) => { pushSelection(app); selectWhere(app, (_cell, r, c) => app.sheet.styleOf(r, c).locked !== false, { what: 'locked cells', existingOnly: false }); } },
  { id: 'util.selectUnlocked', label: 'Select Unlocked Cells', icon: 'lock_open', tab: U, group: G_SELECT, noRecord: true, run: (app) => { pushSelection(app); selectWhere(app, (_cell, r, c) => app.sheet.styleOf(r, c).locked === false, { what: 'unlocked cells', existingOnly: false }); } },
  { id: 'util.selectMerged', label: 'Select Merged Cells', icon: 'call_merge', tab: U, group: G_SELECT, noRecord: true, run: (app) => { const s = app.sheet; if (!s.merges.length) { toast('There are no merged cells on this sheet.'); return; } pushSelection(app); s.sel.ranges = s.merges.map((m) => ({ ...m })); s.sel.r = s.merges[0].r1; s.sel.c = s.merges[0].c1; app.grid.invalidate(); app.ui.status(`${plural(s.merges.length, 'merged range')} selected`); } },
  { id: 'util.selectHyperlinks', label: 'Select Cells with Hyperlinks', icon: 'link', tab: U, group: G_SELECT, noRecord: true, run: (app) => { pushSelection(app); selectWhere(app, (cell) => !!cell?.link, { what: 'cells with hyperlinks' }); } },
  { id: 'util.selectComments', label: 'Select Cells with Comments or Notes', icon: 'comment', tab: U, group: G_SELECT, noRecord: true, run: (app) => { pushSelection(app); selectWhere(app, (_cell, r, c) => app.sheet.comments.has(cellKey(r, c)), { what: 'cells with comments', existingOnly: false }); } },
  { id: 'util.selectWrapped', label: 'Select Cells with Wrapped Text', icon: 'wrap_text', tab: U, group: G_SELECT, noRecord: true, run: (app) => { pushSelection(app); selectWhere(app, (_cell, r, c) => !!app.sheet.styleOf(r, c).wrap, { what: 'wrapped cells', existingOnly: false }); } },
  { id: 'util.selectValidation', label: 'Select Cells with Data Validation', icon: 'rule', tab: U, group: G_SELECT, noRecord: true, run: (app) => app.exec('nav.selectValidation') },
  {
    id: 'util.selectDuplicates', label: 'Select Duplicate Values', icon: 'content_copy', tab: U, group: G_SELECT, noRecord: true, kind: 'split', tip: 'Select the cells whose value occurs more than once in the selection (or the sheet)', origin: 'Kutools (Select Duplicate & Unique Cells)',
    run: (app, arg?: 'first' | 'all' | 'unique') => {
      const cells = singleCell(app) ? (() => { const u = app.sheet.usedRange(); const out: { r: number; c: number; cell: Cell | undefined }[] = []; if (u) app.sheet.forEachIn(u, (r, c, cell) => out.push({ r, c, cell })); return out; })() : selCells(app, { existingOnly: true });
      const counts = dupeCounts(app.sheet, cells, false);
      const seen = new Set<string>();
      const hits = cells.filter((x) => { if (!hasContent(x.cell)) return false; const k = shownText(app.sheet, x.r, x.c).toLowerCase(); const n = counts.get(k) || 0; if (arg === 'unique') return n === 1; if (n < 2) return false; if (arg === 'first') { if (seen.has(k)) return false; seen.add(k); } else if (arg !== 'all') { if (!seen.has(k)) { seen.add(k); return false; } } return true; });
      pushSelection(app);
      selectHits(app, hits.map((x) => ({ r: x.r, c: x.c })), arg === 'unique' ? 'unique values' : 'duplicate values');
    },
    menu: () => [{ id: 'util.selectDuplicates', arg: 'rest', label: 'Duplicates (except the first occurrence)' }, { id: 'util.selectDuplicates', arg: 'all', label: 'Duplicates (including the first occurrence)' }, { id: 'util.selectDuplicates', arg: 'first', label: 'First occurrence of each duplicate' }, { id: 'util.selectDuplicates', arg: 'unique', label: 'Unique values (occur once)' }],
  },
  {
    id: 'util.selectMaxMin', label: 'Select Cells with the Largest or Smallest Value', icon: 'swap_vert', tab: U, group: G_SELECT, noRecord: true, kind: 'split', origin: 'Kutools (Select Cells with Max & Min Value)',
    run: (app, arg?: 'max' | 'min') => {
      const mode = arg || 'max';
      const cells = selCells(app, { existingOnly: true }).filter((x) => typeof x.cell?.v === 'number');
      if (!cells.length) { toast('No numbers in the selection.'); return; }
      const best = cells.reduce((a, x) => (mode === 'max' ? Math.max(a, x.cell!.v as number) : Math.min(a, x.cell!.v as number)), mode === 'max' ? -Infinity : Infinity);
      pushSelection(app);
      selectHits(app, cells.filter((x) => x.cell!.v === best).map((x) => ({ r: x.r, c: x.c })), `${mode === 'max' ? 'largest' : 'smallest'} value cells`);
    },
    menu: () => [{ id: 'util.selectMaxMin', arg: 'max', label: 'Largest value' }, { id: 'util.selectMaxMin', arg: 'min', label: 'Smallest value' }],
  },
  { id: 'util.extendToLastRow', label: 'Extend Selection to the Last Used Row', icon: 'vertical_align_bottom', tab: U, group: G_SELECT, noRecord: true, keys: ['Mod-Alt-Shift-ArrowDown'], origin: 'ASAP Utilities', run: (app) => extendSelection(app, 'lastRow') },
  { id: 'util.extendToLastCol', label: 'Extend Selection to the Last Used Column', icon: 'keyboard_tab', tab: U, group: G_SELECT, noRecord: true, keys: ['Mod-Alt-Shift-ArrowRight'], origin: 'ASAP Utilities', run: (app) => extendSelection(app, 'lastCol') },
  { id: 'util.extendToEnd', label: 'Extend Selection to the Last Used Cell', icon: 'south_east', tab: U, group: G_SELECT, noRecord: true, origin: 'ASAP Utilities', run: (app) => extendSelection(app, 'end') },
  { id: 'util.selectToColumnEnd', label: 'Select from Here to the End of the Column Data', icon: 'arrow_downward', tab: U, group: G_SELECT, noRecord: true, tip: 'Select from the active cell down to the last non-blank cell of its column', run: (app) => extendSelection(app, 'colEnd') },
  { id: 'util.selectToRowEnd', label: 'Select from Here to the End of the Row Data', icon: 'arrow_forward', tab: U, group: G_SELECT, noRecord: true, run: (app) => extendSelection(app, 'rowEnd') },
  { id: 'util.selectToColumnStart', label: 'Select from Here to the Top', icon: 'arrow_upward', tab: U, group: G_SELECT, noRecord: true, run: (app) => extendSelection(app, 'colStart') },
  { id: 'util.selectToRowStart', label: 'Select from Here to the First Column', icon: 'arrow_back', tab: U, group: G_SELECT, noRecord: true, run: (app) => extendSelection(app, 'rowStart') },
  { id: 'util.selectUsedRange', label: 'Select the Used Range', icon: 'select_all', tab: U, group: G_SELECT, noRecord: true, run: (app) => { const u = app.sheet.usedRange(); if (!u) { toast('The sheet is empty.'); return; } pushSelection(app); selectRange(app, u); } },
  { id: 'util.selectPrintArea', label: 'Select the Print Area', icon: 'print', tab: U, group: G_SELECT, noRecord: true, run: (app) => { const a = app.sheet.print.area; const rg = a ? parseRange(a.split(',')[0].replace(/\$/g, '')) : null; if (!rg) { toast('No print area is set on this sheet.'); return; } pushSelection(app); selectRange(app, rg); } },
  { id: 'util.invertSelection', label: 'Invert the Selection', icon: 'flip_to_back', tab: U, group: G_SELECT, noRecord: true, tip: 'Select every used cell that is not selected now', origin: 'ASAP Utilities', run: (app) => invertSelection(app) },
  {
    id: 'util.deselect', label: 'Deselect Cells…', icon: 'deselect', tab: U, group: G_SELECT, noRecord: true, tip: 'Remove a range (or the active area) from a multi-area selection', origin: 'ASAP Utilities',
    run: async (app, arg?: string) => {
      const s = app.sheet;
      const text = await argOrAsk(arg, () => promptDialog('Deselect Cells', 'Range to remove from the selection', rangeName(activeRange(app))));
      if (!text) return;
      const cut = parseRange(text.replace(/\$/g, ''));
      if (!cut) { toast('That is not a valid range.'); return; }
      const { subtractRange } = await import('./dataops');
      const next = s.sel.ranges.flatMap((rg) => subtractRange(rg, cut));
      if (!next.length) { toast('That would leave nothing selected.'); return; }
      pushSelection(app);
      s.sel.ranges = next; s.sel.r = next[0].r1; s.sel.c = next[0].c1; app.grid.invalidate();
    },
  },
  {
    id: 'util.selectRandom', label: 'Select Random Cells or Rows…', icon: 'casino', tab: U, group: G_SELECT, noRecord: true, tip: 'Pick a random sample of the selected cells or rows', origin: 'ASAP Utilities / Kutools (Select Range Randomly)',
    run: async (app, arg?: { n: number; what: 'cells' | 'rows' | 'cols' }) => {
      const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Select Random', fields: [{ key: 'n', label: 'How many', type: 'number', value: 5, min: 1 }, { key: 'what', label: 'Pick random', type: 'radio', value: 'rows', options: [{ value: 'cells', label: 'Cells' }, { value: 'rows', label: 'Rows of the selection' }, { value: 'cols', label: 'Columns of the selection' }] }] }); return r ? { n: Number(r.n), what: r.what } : null; });
      if (!o) return;
      const rg = selBlock(app);
      const pool = o.what === 'cells' ? selCells(app).map((x) => ({ r: x.r, c: x.c })) : o.what === 'rows' ? Array.from({ length: rg.r2 - rg.r1 + 1 }, (_x, i) => rg.r1 + i) : Array.from({ length: rg.c2 - rg.c1 + 1 }, (_x, i) => rg.c1 + i);
      for (let i = pool.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); [pool[i], pool[j]] = [pool[j], pool[i]]; }
      const pick = pool.slice(0, Math.max(1, Math.min(o.n, pool.length)));
      pushSelection(app);
      if (o.what === 'cells') selectHits(app, pick as { r: number; c: number }[], 'random cells');
      else { const s = app.sheet; s.sel.ranges = (pick as number[]).map((l) => (o.what === 'rows' ? { r1: l, c1: rg.c1, r2: l, c2: rg.c2 } : { r1: rg.r1, c1: l, r2: rg.r2, c2: l })); s.sel.r = s.sel.ranges[0].r1; s.sel.c = s.sel.ranges[0].c1; app.grid.invalidate(); app.ui.status(`${pick.length} random ${o.what} selected`); }
    },
  },
  {
    id: 'util.selectEveryNth', label: 'Select, Hide or Delete Every Nth Row or Column…', icon: 'view_agenda', tab: U, group: G_SELECT, edits: true, popular: true, tip: 'Every 2nd (3rd, …) row or column of the selection: select, hide or delete them', keywords: 'alternate rows interval every other', origin: 'ASAP Utilities / Kutools (Select Interval Rows & Columns)',
    run: async (app, arg?: { axis: 'row' | 'col'; every: number; start: number; action: string }) => {
      const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Every Nth Row or Column', fields: [{ key: 'axis', label: 'Work on', type: 'radio', value: 'row', options: [{ value: 'row', label: 'Rows' }, { value: 'col', label: 'Columns' }] }, { key: 'every', label: 'Every', type: 'number', value: 2, min: 1, half: true }, { key: 'start', label: 'Starting with the', type: 'number', value: 2, min: 1, half: true, suffix: 'th' }, { key: 'action', label: 'Then', type: 'radio', value: 'select', options: [{ value: 'select', label: 'Select them' }, { value: 'hide', label: 'Hide them' }, { value: 'delete', label: 'Delete them' }] }] }); return r ? { axis: r.axis, every: Number(r.every) || 1, start: Number(r.start) || 1, action: r.action } : null; });
      if (!o) return;
      const rg = selBlock(app);
      const lines: number[] = [];
      const first = o.axis === 'row' ? rg.r1 : rg.c1, last = o.axis === 'row' ? rg.r2 : rg.c2;
      for (let l = first + o.start - 1; l <= last; l += o.every) lines.push(l);
      everyNth(app, lines, o.axis, o.action, 'Every Nth');
    },
  },
  { id: 'util.rememberSelection', label: 'Remember the Selection', icon: 'bookmark_add', tab: U, group: G_SELECT, noRecord: true, tip: 'Store the current selection so Reselect Remembered Selection can bring it back', origin: 'ASAP Utilities', run: (app) => { const s = app.sheet; remembered = { sheet: s.id, r: s.sel.r, c: s.sel.c, ranges: s.sel.ranges.map((x) => ({ ...x })) }; toast(`Selection ${s.sel.ranges.map(rangeName).join(', ')} remembered`); } },
  { id: 'util.reselectRemembered', label: 'Reselect Remembered Selection', icon: 'bookmark', tab: U, group: G_SELECT, noRecord: true, run: (app) => { if (!remembered) { toast('Nothing has been remembered yet.'); return; } restoreSelection(app, remembered); } },
  { id: 'util.reselectPrevious', label: 'Reselect Previous Selection', icon: 'history', tab: U, group: G_SELECT, noRecord: true, keys: ['Mod-Alt-Backspace'], tip: 'Go back to the selection you had before the last Select tool ran', origin: 'ASAP Utilities', run: (app) => { const prev = selectionStack.pop(); if (!prev) { toast('No previous selection.'); return; } restoreSelection(app, prev); } },
  { id: 'util.selectRowsWithBlanks', label: 'Select Rows with a Blank Cell', icon: 'table_rows', tab: U, group: G_SELECT, noRecord: true, tip: 'Rows of the selection that have at least one empty cell', run: (app) => { const rg = selBlock(app); const s = app.sheet; const lines: number[] = []; for (let r = rg.r1; r <= rg.r2; r++) { let blank = false; for (let c = rg.c1; c <= rg.c2 && !blank; c++) if (!hasContent(s.get(r, c))) blank = true; if (blank) lines.push(r); } if (!lines.length) { toast('No rows with blanks.'); return; } pushSelection(app); s.sel.ranges = mergeSpansOf(lines).map(([a, b]) => ({ r1: a, c1: rg.c1, r2: b, c2: rg.c2 })); s.sel.r = lines[0]; s.sel.c = rg.c1; app.grid.invalidate(); app.ui.status(`${plural(lines.length, 'row')} selected`); } },
  { id: 'util.selectLongText', label: 'Select Cells with Text Longer Than…', icon: 'short_text', tab: U, group: G_SELECT, noRecord: true, run: async (app, arg?: number) => { const n = await argOrAsk(arg, async () => { const t = await promptDialog('Select Long Text', 'Longer than (characters)', '255', { type: 'number' }); return t == null ? null : Number(t); }); if (n == null) return; pushSelection(app); selectWhere(app, (cell, r, c) => hasContent(cell) && shownText(app.sheet, r, c).length > n, { what: `cells longer than ${n} characters` }); } },
  { id: 'util.menuSelect', label: 'Select', icon: 'select', tab: U, group: G_SELECT, kind: 'menu', noRecord: true, run: () => undefined, menu: () => menuOf(['util.selectByContent', 'util.condRows', 'util.condCols', 'util.selectEveryNth', 'util.selectRandom', null, 'util.selectSameFormat', 'util.selectSameFill', 'util.selectSameFontColor', 'util.selectSameValue', 'util.selectSameNumFmt', 'util.selectSameFont', null, 'nav.selectFormulas', 'nav.selectConstants', 'util.selectErrors', 'util.selectNumbers', 'util.selectText', 'util.selectDates', 'util.selectBooleans', 'util.selectBlanks', 'util.selectNonBlanks', 'util.selectNegative', 'util.selectPositive', 'util.selectDuplicates', 'util.selectMaxMin', 'util.selectLongText', null, 'util.selectBold', 'util.selectItalic', 'util.selectLocked', 'util.selectUnlocked', 'util.selectMerged', 'util.selectWrapped', 'util.selectHyperlinks', 'util.selectComments', 'util.selectValidation', 'nav.selectCF', null, 'util.extendToLastRow', 'util.extendToLastCol', 'util.extendToEnd', 'util.selectToColumnEnd', 'util.selectToRowEnd', 'util.selectToColumnStart', 'util.selectToRowStart', 'util.selectUsedRange', 'util.selectPrintArea', 'nav.currentRegion', 'nav.selectVisible', 'util.selectRowsWithBlanks', null, 'util.invertSelection', 'util.deselect', 'util.rememberSelection', 'util.reselectRemembered', 'util.reselectPrevious']) },
]);

// ============================================================================================
// Sheets

function addSheet(app: App, name: string, after: number): Sheet {
  const book = app.book;
  const s = new Sheet(book, name);
  book.sheets.splice(after + 1, 0, s);
  return s;
}

const INDEX_MARK = 'QlSheetIndex';

/** A sheet's content for a copy: without its id, so the copy keeps its own (undo finds a sheet by id). */
function copyJSON(src: Sheet): any { const j: any = src.toJSON(); delete j.id; return j; }

export function uniqueSheetName(book: import('../model/workbook').Workbook, base: string, taken: Set<string> = new Set()): string {
  const clean = (validSheetName(base) ? base : base.replace(/[\\/?*[\]:]/g, ' ').trim().slice(0, 31)) || 'Sheet';
  const has = (n: string) => !!book.sheetByName(n) || taken.has(n.toLowerCase());
  if (!has(clean)) return clean;
  let i = 2;
  while (has(`${clean.slice(0, 27)} (${i})`)) i++;
  return `${clean.slice(0, 27)} (${i})`;
}

function sheetValuesOfSelection(app: App): string[] {
  const out: string[] = [];
  for (const x of selCells(app, { existingOnly: true })) { const t = shownText(app.sheet, x.r, x.c).trim(); if (t && !out.includes(t)) out.push(t); }
  return out;
}

async function sortSheets(app: App, arg?: { by: 'name' | 'nameDesc' | 'natural' | 'color' | 'reverse' | 'custom'; names?: string[] }): Promise<void> {
  if (structureLocked(app)) return;
  const book = app.book;
  const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Sort Sheets', fields: [{ key: 'by', label: 'Order', type: 'radio', value: 'natural', options: [{ value: 'natural', label: 'By name (natural: Sheet2 before Sheet10)' }, { value: 'name', label: 'By name A→Z' }, { value: 'nameDesc', label: 'By name Z→A' }, { value: 'color', label: 'By tab color' }, { value: 'reverse', label: 'Reverse the current order' }, { value: 'custom', label: 'Custom order (list below)' }] }, { key: 'names', label: 'Custom order, one sheet name per line', type: 'textarea', rows: 6, value: book.sheets.map((s) => s.name).join('\n'), visibleIf: (v) => v.by === 'custom' }] }); return r ? { by: r.by, names: String(r.names || '').split(/\r?\n/).map((x) => x.trim()).filter(Boolean) } : null; });
  if (!o) return;
  const active = book.sheet;
  let sorted = book.sheets.slice();
  if (o.by === 'name') sorted.sort((a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: 'base' }));
  else if (o.by === 'nameDesc') sorted.sort((a, b) => b.name.localeCompare(a.name, undefined, { sensitivity: 'base' }));
  else if (o.by === 'natural') sorted.sort((a, b) => naturalCompare(a.name, b.name));
  else if (o.by === 'color') sorted.sort((a, b) => (a.tabColor || 'zzz').localeCompare(b.tabColor || 'zzz') || naturalCompare(a.name, b.name));
  else if (o.by === 'reverse') sorted.reverse();
  else { const wanted = (o.names || []).map((n) => book.sheetByName(n)).filter((s): s is Sheet => !!s); sorted = [...wanted, ...book.sheets.filter((s) => !wanted.includes(s))]; }
  book.transact('Sort Sheets', () => { book.recordBook(); book.sheets = sorted; book.activeSheet = sorted.indexOf(active); });
  app.ui.rebuild(); app.refresh();
  toast('Sheets sorted');
}

async function deleteSheets(app: App, pick: (s: Sheet) => boolean, label: string, arg: any): Promise<void> {
  if (structureLocked(app)) return;
  const book = app.book;
  const victims = book.sheets.filter(pick);
  if (!victims.length) { toast(`No sheets to delete (${label.toLowerCase()}).`); return; }
  if (victims.length >= book.sheets.length) { toast('A workbook must keep at least one sheet.'); return; }
  if (arg === undefined && !(await confirmDialog(label, `Delete ${plural(victims.length, 'sheet')}: ${victims.map((s) => s.name).join(', ')}?`, 'Delete'))) return;
  const { invalidateSheetRefs } = await import('../model/structure');
  book.transact(label, () => { book.recordBook(); for (const v of victims) { invalidateSheetRefs(book, v.name); book.sheets = book.sheets.filter((s) => s !== v); } book.activeSheet = Math.max(0, Math.min(book.activeSheet, book.sheets.length - 1)); });
  app.ui.rebuild(); app.exec('sheet.activate', book.activeSheet);
  toast(`${plural(victims.length, 'sheet')} deleted`);
}

function sheetIndexRows(app: App, withLinks: boolean): import('./utilShared').OutCell[][] {
  const rows: import('./utilShared').OutCell[][] = [];
  app.book.sheets.forEach((s, i) => {
    const u = s.usedRange(false);
    rows.push([withLinks ? { v: s.name, f: undefined } : s.name, i + 1, s.visibility === 'visible' ? 'Visible' : 'Hidden', u ? rangeName(u) : '(empty)', u ? (u.r2 - u.r1 + 1) : 0, u ? (u.c2 - u.c1 + 1) : 0, s.cells.size]);
  });
  return rows;
}

defTools([
  {
    id: 'util.insertSheets', label: 'Insert Multiple Sheets…', icon: 'library_add', tab: U, group: G_SHEETS, edits: true, popular: true,
    tip: 'Insert several sheets at once, named from a list, a numbered pattern or the selected cells', keywords: 'add sheets names list create', origin: 'ASAP Utilities / Kutools (Create Sequence Worksheets)',
    run: async (app, arg?: { names?: string[]; count?: number; pattern?: string; fromSelection?: boolean; copyActive?: boolean }) => {
      if (structureLocked(app)) return;
      const book = app.book;
      const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Insert Multiple Sheets', fields: [{ key: 'mode', label: 'Names', type: 'radio', value: 'list', options: [{ value: 'list', label: 'From this list (one per line)' }, { value: 'sel', label: 'From the selected cells' }, { value: 'pattern', label: 'Numbered pattern' }] }, { key: 'names', label: 'Sheet names', type: 'textarea', rows: 5, value: '', visibleIf: (v) => v.mode === 'list' }, { key: 'pattern', label: 'Pattern ({n} = number)', type: 'text', value: 'Sheet {n}', visibleIf: (v) => v.mode === 'pattern', half: true }, { key: 'count', label: 'How many', type: 'number', value: 3, min: 1, max: 200, visibleIf: (v) => v.mode === 'pattern', half: true }, { key: 'copy', label: 'Make each one a copy of the active sheet', type: 'checkbox', value: false }] }); if (!r) return null; return { names: r.mode === 'list' ? String(r.names || '').split(/\r?\n/).map((x) => x.trim()).filter(Boolean) : undefined, fromSelection: r.mode === 'sel', pattern: r.mode === 'pattern' ? r.pattern : undefined, count: Number(r.count) || 1, copyActive: !!r.copy }; });
      if (!o) return;
      let names = o.names || [];
      if (o.fromSelection) names = sheetValuesOfSelection(app);
      if (o.pattern) names = Array.from({ length: o.count || 1 }, (_x, i) => o.pattern!.replace(/\{n\}/g, String(i + 1)).replace(/\{nn\}/g, String(i + 1).padStart(2, '0')));
      if (!names.length) { toast('No sheet names were given.'); return; }
      const taken = new Set<string>();
      const src = app.sheet;
      book.transact('Insert Multiple Sheets', () => {
        book.recordBook();
        let at = book.activeSheet;
        for (const n of names) {
          const name = uniqueSheetName(book, n, taken); taken.add(name.toLowerCase());
          const s = o.copyActive ? (() => { const c = new Sheet(book, name); c.loadJSON(copyJSON(src)); c.name = name; return c; })() : new Sheet(book, name);
          if (o.copyActive) { s.name = name; }
          book.sheets.splice(++at, 0, s);
        }
      });
      app.ui.rebuild(); app.exec('sheet.activate', book.activeSheet + 1);
      toast(`${plural(names.length, 'sheet')} inserted`);
    },
  },
  { id: 'util.sortSheets', label: 'Sort Sheets…', icon: 'sort_by_alpha', tab: U, group: G_SHEETS, edits: true, popular: true, tip: 'Put the sheets in alphabetical, natural, reverse, tab-color or custom order', origin: 'ASAP Utilities / Kutools (Sort Sheets)', run: (app, arg) => sortSheets(app, arg) },
  { id: 'util.reverseSheets', label: 'Reverse Sheet Order', icon: 'swap_horiz', tab: U, group: G_SHEETS, edits: true, run: (app) => sortSheets(app, { by: 'reverse' }) },
  { id: 'util.deleteEmptySheets', label: 'Delete Empty Sheets', icon: 'delete_sweep', tab: U, group: G_SHEETS, edits: true, tip: 'Delete every sheet with no cells, charts or objects', origin: 'ASAP Utilities / Kutools', run: (app, arg) => deleteSheets(app, (s) => s.cells.size === 0 && !s.charts.length && !s.drawings.length && !s.comments.size, 'Delete Empty Sheets', arg) },
  { id: 'util.deleteHiddenSheets', label: 'Delete Hidden Sheets', icon: 'visibility_off', tab: U, group: G_SHEETS, edits: true, origin: 'Kutools (Delete All Hidden Sheets)', run: (app, arg) => deleteSheets(app, (s) => s.visibility !== 'visible', 'Delete Hidden Sheets', arg) },
  { id: 'util.deleteOtherSheets', label: 'Delete All Sheets Except This One', icon: 'tab_unselected', tab: U, group: G_SHEETS, edits: true, origin: 'ASAP Utilities / Kutools (Delete All Inactive Sheets)', run: (app, arg) => deleteSheets(app, (s) => s !== app.sheet, 'Delete Other Sheets', arg) },
  {
    id: 'util.deleteSheetsByName', label: 'Delete Sheets by Name…', icon: 'find_replace', tab: U, group: G_SHEETS, edits: true, tip: 'Delete the sheets whose name contains text or matches a pattern (* and ? wildcards)',
    run: async (app, arg?: { pattern: string }) => {
      const o = await argOrAsk(arg, async () => { const t = await promptDialog('Delete Sheets by Name', 'Name contains (use * and ? as wildcards)', ''); return t ? { pattern: t } : null; });
      if (!o) return;
      const re = new RegExp('^' + o.pattern.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.') + '$', 'i');
      const plain = o.pattern.toLowerCase();
      await deleteSheets(app, (s) => re.test(s.name) || s.name.toLowerCase().includes(plain), 'Delete Sheets by Name', arg);
    },
  },
  {
    id: 'util.hideSheets', label: 'Hide Sheets…', icon: 'visibility_off', tab: U, group: G_SHEETS, edits: true, tip: 'Hide several sheets at once', origin: 'ASAP Utilities',
    run: async (app, arg?: { names: string[] }) => {
      const pick = arg ? { sheets: sheetsByNames(app.book, arg.names) } : await pickSheets(app, 'Hide Sheets', { others: true, dflt: 'none' });
      if (!pick) return;
      const keep = app.book.sheets.filter((s) => s.visibility === 'visible' && !pick.sheets.includes(s));
      if (!keep.length) { toast('At least one sheet must stay visible.'); return; }
      app.book.transact('Hide Sheets', () => { for (const s of pick.sheets) { app.book.recordSheet(s); s.visibility = 'hidden'; } });
      if (pick.sheets.includes(app.sheet)) app.exec('sheet.activate', app.book.sheets.indexOf(keep[0]));
      app.ui.rebuild();
    },
  },
  { id: 'util.hideOtherSheets', label: 'Hide All Sheets Except This One', icon: 'tab', tab: U, group: G_SHEETS, edits: true, origin: 'ASAP Utilities', run: (app) => { app.book.transact('Hide Other Sheets', () => { for (const s of app.book.sheets) if (s !== app.sheet) { app.book.recordSheet(s); s.visibility = 'hidden'; } }); app.ui.rebuild(); } },
  { id: 'util.unhideAllSheets', label: 'Unhide All Sheets', icon: 'visibility', tab: U, group: G_SHEETS, edits: true, popular: true, origin: 'ASAP Utilities / Kutools', run: (app) => { let n = 0; app.book.transact('Unhide All Sheets', () => { for (const s of app.book.sheets) if (s.visibility !== 'visible') { app.book.recordSheet(s); s.visibility = 'visible'; n++; } }); app.ui.rebuild(); toast(n ? `${plural(n, 'sheet')} unhidden` : 'No hidden sheets.'); } },
  {
    id: 'util.renameSheetsFromCells', label: 'Rename Sheets from a Cell…', icon: 'drive_file_rename_outline', tab: U, group: G_SHEETS, edits: true, tip: 'Name each sheet after the text in one of its cells (A1 by default)', origin: 'ASAP Utilities',
    run: async (app, arg?: { cell: string; names?: string[] }) => {
      const o = await argOrAsk<{ cell: string; names?: string[] }>(arg, async () => { const t = await promptDialog('Rename Sheets from a Cell', 'Cell whose text becomes the sheet name', 'A1'); return t ? { cell: t } : null; });
      if (!o) return;
      const { parseCell } = await import('../model/address');
      const p = parseCell(o.cell.replace(/\$/g, ''));
      if (!p) { toast('That is not a valid cell.'); return; }
      const book = app.book; let n = 0;
      book.transact('Rename Sheets from a Cell', () => { for (const s of sheetsByNames(book, o.names)) { const t = shownText(s, p.r, p.c).trim(); if (!t || t === s.name) continue; const name = uniqueSheetName(book, t); renameSheet(book, s, name); n++; } });
      app.ui.rebuild(); toast(`${plural(n, 'sheet')} renamed`);
    },
  },
  {
    id: 'util.renameSheetsReplace', label: 'Find and Replace in Sheet Names…', icon: 'find_replace', tab: U, group: G_SHEETS, edits: true, origin: 'ASAP Utilities / Kutools (Rename Multiple Worksheets)',
    run: async (app, arg?: { find: string; replace: string; prefix?: string; suffix?: string; caseMode?: string }) => {
      const o = await argOrAsk(arg, () => openDialog({ title: 'Rename Sheets', fields: [{ key: 'find', label: 'Find', type: 'text', value: '', half: true }, { key: 'replace', label: 'Replace with', type: 'text', value: '', half: true }, { key: 'prefix', label: 'Add before every name', type: 'text', value: '', half: true }, { key: 'suffix', label: 'Add after every name', type: 'text', value: '', half: true }, { key: 'caseMode', label: 'Change case', type: 'select', value: '', options: [{ value: '', label: '(keep)' }, { value: 'upper', label: 'UPPERCASE' }, { value: 'lower', label: 'lowercase' }, { value: 'proper', label: 'Proper Case' }] }] }));
      if (!o) return;
      const book = app.book; let n = 0;
      book.transact('Rename Sheets', () => {
        for (const s of book.sheets) {
          let name = o.find ? s.name.split(o.find).join(o.replace || '') : s.name;
          name = `${o.prefix || ''}${name}${o.suffix || ''}`;
          if (o.caseMode === 'upper') name = name.toUpperCase(); else if (o.caseMode === 'lower') name = name.toLowerCase(); else if (o.caseMode === 'proper') name = name.toLowerCase().replace(/(^|\s)(\p{L})/gu, (_m, a: string, b: string) => a + b.toUpperCase());
          if (name === s.name || !name.trim()) continue;
          renameSheet(book, s, uniqueSheetName(book, name)); n++;
        }
      });
      app.ui.rebuild(); toast(`${plural(n, 'sheet')} renamed`);
    },
  },
  {
    id: 'util.copySheetTimes', label: 'Copy This Sheet Several Times…', icon: 'content_copy', tab: U, group: G_SHEETS, edits: true, origin: 'ASAP Utilities / Kutools (Copy Worksheets)',
    run: async (app, arg?: { count: number; pattern?: string }) => {
      if (structureLocked(app)) return;
      const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Copy Sheet Several Times', fields: [{ key: 'count', label: 'Copies', type: 'number', value: 3, min: 1, max: 100 }, { key: 'pattern', label: 'Names ({name} = this sheet, {n} = number)', type: 'text', value: '{name} ({n})' }] }); return r ? { count: Number(r.count) || 1, pattern: r.pattern } : null; });
      if (!o) return;
      const book = app.book; const src = app.sheet; const taken = new Set<string>();
      book.transact('Copy Sheet', () => { book.recordBook(); let at = book.activeSheet; for (let i = 1; i <= o.count; i++) { const name = uniqueSheetName(book, (o.pattern || '{name} ({n})').replace(/\{name\}/g, src.name).replace(/\{n\}/g, String(i + 1)), taken); taken.add(name.toLowerCase()); const c = new Sheet(book, name); c.loadJSON(copyJSON(src)); c.name = name; book.sheets.splice(++at, 0, c); } });
      app.ui.rebuild(); app.refresh(); toast(`${plural(o.count, 'copy', 'copies')} made`);
    },
  },
  {
    id: 'util.sheetsToNewWorkbook', label: 'Copy Sheets to a New Workbook…', icon: 'open_in_new', tab: U, group: G_SHEETS, noRecord: true, tip: 'Open a new workbook that holds copies of the chosen sheets', origin: 'ASAP Utilities / Ablebits (Copy Sheets)',
    run: async (app, arg?: { names: string[] }) => {
      const pick = arg ? { sheets: sheetsByNames(app.book, arg.names) } : await pickSheets(app, 'Copy Sheets to a New Workbook', { dflt: 'active' });
      if (!pick || !pick.sheets.length) return;
      const { Workbook } = await import('../model/workbook');
      const json = app.book.toJSON();
      const keep = new Set(pick.sheets.map((s) => s.name));
      json.sheets = (json.sheets || []).filter((j: any) => keep.has(j.name));
      json.activeSheet = 0;
      const nb = Workbook.fromJSON(json);
      const t = app.createTab(nb, { title: `${pick.sheets.length === 1 ? pick.sheets[0].name : app.tab.title + ' (sheets)'}` });
      t.dirty = true;
    },
  },
  { id: 'util.moveSheetFirst', label: 'Move Sheet to the Beginning', icon: 'first_page', tab: U, group: G_SHEETS, edits: true, run: (app) => moveSheetTo(app, 0) },
  { id: 'util.moveSheetLast', label: 'Move Sheet to the End', icon: 'last_page', tab: U, group: G_SHEETS, edits: true, run: (app) => moveSheetTo(app, app.book.sheets.length - 1) },
  { id: 'util.moveSheetLeft', label: 'Move Sheet Left', icon: 'chevron_left', tab: U, group: G_SHEETS, edits: true, keys: ['Mod-Alt-Shift-ArrowLeft'], run: (app) => moveSheetTo(app, Math.max(0, app.book.activeSheet - 1)) },
  { id: 'util.moveSheetRight', label: 'Move Sheet Right', icon: 'chevron_right', tab: U, group: G_SHEETS, edits: true, keys: ['Mod-Alt-Shift-ArrowRight'], run: (app) => moveSheetTo(app, Math.min(app.book.sheets.length - 1, app.book.activeSheet + 1)) },
  {
    id: 'util.sheetIndex', label: 'Create a Sheet Index (Table of Contents)', icon: 'toc', tab: U, group: G_SHEETS, edits: true, popular: true,
    tip: 'A new first sheet listing every sheet with a link to it, its used range and size', keywords: 'table of contents index links sheets', origin: 'ASAP Utilities / Kutools (Create List of Sheet Names)',
    run: (app) => {
      if (structureLocked(app)) return;
      const book = app.book;
      // Only an index this tool made earlier is replaced; a sheet of the user's that happens to be called Index is kept and listed.
      const sheets = book.sheets.filter((s) => s.codeName !== INDEX_MARK);
      book.transact('Sheet Index', () => {
        book.recordBook();
        book.sheets = book.sheets.filter((s) => s.codeName !== INDEX_MARK);
        const idx = new Sheet(book, uniqueSheetName(book, 'Index'));
        idx.codeName = INDEX_MARK;
        book.sheets.unshift(idx);
        const head = ['Sheet', '#', 'Visibility', 'Used range', 'Rows', 'Columns', 'Cells'];
        head.forEach((t, j) => writeOut(idx, 0, j, { v: t, style: HEAD }));
        sheets.forEach((s, i) => {
          const u = s.usedRange(false);
          idx.set(i + 1, 0, { v: s.name, link: `#'${s.name.replace(/'/g, "''")}'!A1`, s: book.styleId({ color: '#0563c1', underline: 'single' }) });
          writeOut(idx, i + 1, 1, i + 1); writeOut(idx, i + 1, 2, s.visibility === 'visible' ? 'Visible' : 'Hidden'); writeOut(idx, i + 1, 3, u ? rangeName(u) : '(empty)');
          writeOut(idx, i + 1, 4, u ? u.r2 - u.r1 + 1 : 0); writeOut(idx, i + 1, 5, u ? u.c2 - u.c1 + 1 : 0); writeOut(idx, i + 1, 6, s.cells.size);
        });
        idx.colW.set(0, 180); idx.colW.set(3, 110); idx.freeze = { rows: 1, cols: 0 };
        book.activeSheet = 0;
      });
      app.ui.rebuild(); app.exec('sheet.activate', 0);
      toast(`Index of ${plural(sheets.length, 'sheet')} created`);
    },
  },
  {
    id: 'util.listSheetNames', label: 'Insert the Sheet Names in Cells', icon: 'format_list_bulleted', tab: U, group: G_SHEETS, edits: true, tip: 'Write the names of all sheets downward from the active cell', origin: 'ASAP Utilities',
    run: (app) => { const { r, c } = activeCell(app); const s = app.sheet; app.book.transact('List Sheet Names', () => app.book.sheets.forEach((sh, i) => writeOut(s, r + i, c, sh.name))); },
  },
  { id: 'util.insertSheetName', label: 'Insert the Sheet Name in the Active Cell', icon: 'label', tab: U, group: G_SHEETS, edits: true, kind: 'split', tip: 'As text, or as a formula that follows renames', origin: 'ASAP Utilities / Kutools (Insert Workbook Information)', run: (app, arg?: 'text' | 'formula') => { const { r, c } = activeCell(app); app.book.transact('Insert Sheet Name', () => writeOut(app.sheet, r, c, arg === 'formula' ? { f: 'MID(CELL("filename",A1),FIND("]",CELL("filename",A1))+1,255)' } : app.sheet.name)); }, menu: () => [{ id: 'util.insertSheetName', arg: 'text', label: 'As text' }, { id: 'util.insertSheetName', arg: 'formula', label: 'As a formula' }] },
  {
    id: 'util.protectSheets', label: 'Protect Several Sheets…', icon: 'lock', tab: U, group: G_SHEETS, edits: true, popular: true, tip: 'Protect the chosen sheets with one password and the same allowed actions', origin: 'ASAP Utilities / Kutools (Protect Worksheet)',
    run: async (app, arg?: { names?: string[]; password?: string; allow?: string[] }) => {
      const pick = arg ? { sheets: sheetsByNames(app.book, arg.names), values: arg } : await pickSheets(app, 'Protect Sheets', { extra: [{ key: 'password', label: 'Password (optional)', type: 'password', value: '' }, { key: 'again', label: 'Confirm password', type: 'password', value: '' }, { key: 'h', label: 'Allow all users to:', type: 'heading' }, ...PROTECT_OPTIONS.map((o) => ({ key: 'a_' + o.key, label: o.label, type: 'checkbox' as const, value: o.key.startsWith('select'), half: true }))] });
      if (!pick) return;
      const v = pick.values as Record<string, any>;
      if (!arg && v.password !== v.again) { toast('The passwords do not match.'); return; }
      const pw = String((arg ? arg.password : v.password) || '');
      const hash = pw ? await hashPassword(pw) : undefined;
      const allow = arg?.allow || PROTECT_OPTIONS.filter((o) => v['a_' + o.key]).map((o) => o.key);
      app.book.transact('Protect Sheets', () => { for (const s of pick.sheets) { app.book.recordSheet(s); s.protection = { on: true, allow, ...(hash ? { passwordHash: hash } : {}) }; } });
      toast(`${plural(pick.sheets.length, 'sheet')} protected`);
    },
  },
  {
    id: 'util.unprotectSheets', label: 'Unprotect Several Sheets…', icon: 'lock_open', tab: U, group: G_SHEETS, edits: true, origin: 'ASAP Utilities / Kutools (Unprotect Worksheet)',
    run: async (app, arg?: { names?: string[]; password?: string }) => {
      const locked = app.book.sheets.filter((s) => s.protection.on);
      if (!locked.length) { toast('No sheet is protected.'); return; }
      const pick = arg ? { sheets: sheetsByNames(app.book, arg.names).filter((s) => s.protection.on), values: arg } : await pickSheets(app, 'Unprotect Sheets', { extra: [{ key: 'password', label: 'Password (if any)', type: 'password', value: '' }] });
      if (!pick) return;
      const pw = String((pick.values as any).password || '');
      let n = 0, failed = 0;
      for (const s of pick.sheets) { if (!s.protection.on) continue; if (s.protection.passwordHash && !(await verifyPassword(pw, s.protection.passwordHash))) { failed++; continue; } app.book.transact('Unprotect Sheet', () => { app.book.recordSheet(s); s.protection = { on: false, allow: s.protection.allow }; }); n++; }
      toast(`${plural(n, 'sheet')} unprotected${failed ? `; ${failed} kept (wrong password)` : ''}`);
    },
  },
  {
    id: 'util.pageSetupToSheets', label: 'Copy Page Setup to Other Sheets…', icon: 'print', tab: U, group: G_SHEETS, edits: true, tip: 'Give other sheets the margins, orientation, paper, scaling, headers and footers of this sheet', origin: 'ASAP Utilities / Kutools (Copy Page Setup)',
    run: async (app, arg?: { names?: string[] }) => {
      const pick = arg ? { sheets: sheetsByNames(app.book, arg.names) } : await pickSheets(app, 'Copy Page Setup', { others: true });
      if (!pick) return;
      const src = app.sheet.print;
      app.book.transact('Copy Page Setup', () => { for (const s of pick.sheets) { if (s === app.sheet) continue; app.book.recordSheet(s); const { area, titleRows, titleCols, ...rest } = JSON.parse(JSON.stringify(src)); void area; void titleRows; void titleCols; s.print = { ...s.print, ...rest }; } });
      toast(`Page setup copied to ${plural(pick.sheets.filter((s) => s !== app.sheet).length, 'sheet')}`);
    },
  },
  {
    id: 'util.headerFooterAll', label: 'Same Header and Footer on All Sheets', icon: 'vertical_split', tab: U, group: G_SHEETS, edits: true, origin: 'ASAP Utilities',
    run: (app) => { const src = app.sheet.print; for (const s of app.book.sheets) if (s !== app.sheet) setPrint(app, 'Header and Footer', (p) => { p.header = { ...src.header }; p.footer = { ...src.footer }; }, s); toast('Header and footer copied to every sheet'); },
  },
  { id: 'util.freezeAllSheets', label: 'Freeze Panes on All Sheets Like This One', icon: 'view_compact', tab: U, group: G_SHEETS, edits: true, origin: 'Kutools (Freeze Panes Multiple Worksheets)', run: (app) => { const f = { ...app.sheet.freeze }; app.book.transact('Freeze All Sheets', () => { for (const s of app.book.sheets) { app.book.recordSheet(s); s.freeze = { ...f }; } }); app.grid.layout(); } },
  { id: 'util.unfreezeAllSheets', label: 'Unfreeze Panes on All Sheets', icon: 'view_comfy', tab: U, group: G_SHEETS, edits: true, origin: 'Kutools', run: (app) => { app.book.transact('Unfreeze All Sheets', () => { for (const s of app.book.sheets) { app.book.recordSheet(s); s.freeze = { rows: 0, cols: 0 }; } }); app.grid.layout(); } },
  {
    id: 'util.tabColors', label: 'Tab Colors for Several Sheets…', icon: 'palette', tab: U, group: G_SHEETS, edits: true, tip: 'Give the chosen sheets one tab color, rainbow colors, or no color',
    run: async (app, arg?: { names?: string[]; color?: string | 'rainbow' | 'none' }) => {
      const pick = arg ? { sheets: sheetsByNames(app.book, arg.names), values: arg } : await pickSheets(app, 'Tab Colors', { extra: [{ key: 'mode', label: 'Color', type: 'radio', value: 'rainbow', options: [{ value: 'rainbow', label: 'A different color per sheet' }, { value: 'one', label: 'This color:' }, { value: 'none', label: 'No color' }] }, { key: 'color', label: 'Color', type: 'color', value: '#4a90c2', visibleIf: (v) => v.mode === 'one' }] });
      if (!pick) return;
      const v = pick.values as any; const mode = arg ? (arg.color === 'rainbow' || arg.color === 'none' ? arg.color : 'one') : v.mode; const one = arg ? arg.color : v.color;
      const rainbow = ['#e57373', '#ffb74d', '#fff176', '#81c784', '#4fc3f7', '#7986cb', '#ba68c8', '#a1887f'];
      app.book.transact('Tab Colors', () => pick.sheets.forEach((s, i) => { app.book.recordSheet(s); s.tabColor = mode === 'none' ? null : mode === 'rainbow' ? rainbow[i % rainbow.length] : (one as string); }));
      app.ui.rebuild();
    },
  },
  {
    id: 'util.sheetsReport', label: 'Sheet Overview Report', icon: 'summarize', tab: U, group: G_SHEETS, edits: true, tip: 'A new sheet listing each sheet with its used range, row and column counts, cell count, formulas, comments, charts and protection',
    run: (app) => {
      const rows = app.book.sheets.map((s, i) => { const u = s.usedRange(false); let f = 0; for (const cell of s.cells.values()) if (cell.f != null) f++; return [s.name, i + 1, s.visibility === 'visible' ? 'Visible' : 'Hidden', u ? rangeName(u) : '(empty)', u ? u.r2 - u.r1 + 1 : 0, u ? u.c2 - u.c1 + 1 : 0, s.cells.size, f, s.comments.size, s.charts.length + s.drawings.length, s.merges.length, s.protection.on ? 'Yes' : 'No'] as Scalar[]; });
      report(app, 'Sheets', ['Sheet', '#', 'Visibility', 'Used range', 'Rows', 'Columns', 'Cells', 'Formulas', 'Comments', 'Objects', 'Merges', 'Protected'], rows);
    },
  },
  { id: 'util.menuSheets', label: 'Sheets', icon: 'tab', tab: U, group: G_SHEETS, kind: 'menu', noRecord: true, run: () => undefined, menu: () => menuOf(['util.insertSheets', 'util.copySheetTimes', 'util.sheetsToNewWorkbook', 'sheet.moveCopy', null, 'util.sortSheets', 'util.reverseSheets', 'util.moveSheetFirst', 'util.moveSheetLeft', 'util.moveSheetRight', 'util.moveSheetLast', null, 'util.renameSheetsFromCells', 'util.renameSheetsReplace', 'util.tabColors', null, 'util.hideSheets', 'util.hideOtherSheets', 'util.unhideAllSheets', 'util.deleteEmptySheets', 'util.deleteHiddenSheets', 'util.deleteOtherSheets', 'util.deleteSheetsByName', null, 'util.protectSheets', 'util.unprotectSheets', 'util.pageSetupToSheets', 'util.headerFooterAll', 'util.freezeAllSheets', 'util.unfreezeAllSheets', null, 'util.sheetIndex', 'util.listSheetNames', 'util.insertSheetName', 'util.sheetsReport', 'util.combineSheets', 'util.splitToSheets', 'util.exportSheetsFiles']) },
]);

function moveSheetTo(app: App, to: number): void {
  if (structureLocked(app)) return;
  const book = app.book; const from = book.activeSheet;
  if (from === to) return;
  book.transact('Move Sheet', () => { book.recordBook(); const [s] = book.sheets.splice(from, 1); book.sheets.splice(to, 0, s); book.activeSheet = to; });
  app.ui.rebuild(); app.refresh();
}

void sheetIndexRows; void criteriaFn; void clipToUsed; void currentRegion; void isWholeCols; void isWholeRows; void rowSpans; void colSpans; void textOf; void selRanges; void copyText; void toolLog; void keyRow; void keyCol; void colName; void cellName; void addSheet;

// Remember the sheets visited so "Back to Previous Sheet" can return to them.
{
  const act = getCommand('sheet.activate');
  if (act) { const run = act.run; act.run = (app, arg) => { noteSheetVisit(app.sheet.id); const r = run(app, arg); noteSheetVisit(app.sheet.id); return r; }; }
}
