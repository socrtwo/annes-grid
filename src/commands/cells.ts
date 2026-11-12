/** Cells group (insert, delete, format: sizes, visibility, organize sheets) and sheet-tab commands. */
import { def } from './registry';
import type { App } from '../app';
import { Sheet } from '../model/workbook';
import { MAX_COLS, MAX_ROWS, colName, keyRow, keyCol, type Range } from '../model/address';
import { shiftCells, insertCellsShift, renameSheet, invalidateSheetRefs, validSheetName } from '../model/structure';
import { selRanges, activeRange, activeCell, isWholeCols, isWholeRows, clipToUsed, filteredOut } from './ops';
import { openDialog, toast, colorMenu, confirmDialog, promptDialog, choiceDialog } from '../ui/dialog';
import { fontCss } from '../grid/render';
import { formatValue } from '../model/numfmt';

function protectedSheet(app: App, action: string): boolean {
  const s = app.sheet;
  if (s.protection.on && !s.protection.allow.includes(action)) { toast('The sheet is protected. Unprotect it (Review ▸ Unprotect Sheet) to make this change.'); return true; }
  return false;
}

function protectedStructure(app: App): boolean {
  if (app.book.structureLocked) { toast('The workbook structure is protected. Use Review ▸ Protect Workbook to change it.'); return true; }
  return false;
}

/** Rows (or columns) spanned by the selection, merged into runs, last first (so indexes stay valid). */
function spans(app: App, axis: 'row' | 'col'): [number, number][] {
  const out: [number, number][] = [];
  for (const rg of selRanges(app)) out.push(axis === 'row' ? [rg.r1, Math.min(rg.r2, MAX_ROWS - 1)] : [rg.c1, Math.min(rg.c2, MAX_COLS - 1)]);
  out.sort((a, b) => a[0] - b[0]);
  const merged: [number, number][] = [];
  for (const s of out) { const l = merged[merged.length - 1]; if (l && s[0] <= l[1] + 1) l[1] = Math.max(l[1], s[1]); else merged.push([...s]); }
  return merged.reverse();
}

function insertWhole(app: App, axis: 'row' | 'col'): void {
  if (protectedSheet(app, axis === 'row' ? 'insertRows' : 'insertColumns')) return;
  const s = app.sheet;
  const used = s.usedRange();
  const sp = spans(app, axis);
  const need = sp.reduce((a, [x, y]) => a + y - x + 1, 0);
  if (used && (axis === 'row' ? used.r2 + need >= MAX_ROWS : used.c2 + need >= MAX_COLS)) { toast('To prevent possible loss of data, cells cannot be shifted off the sheet.'); return; }
  app.book.transact(axis === 'row' ? 'Insert Rows' : 'Insert Columns', () => { for (const [a, b] of sp) shiftCells(s, axis, a, b - a + 1); });
  app.grid.updateExtent();
}

function deleteWhole(app: App, axis: 'row' | 'col'): void {
  if (protectedSheet(app, axis === 'row' ? 'deleteRows' : 'deleteColumns')) return;
  const s = app.sheet;
  app.book.transact(axis === 'row' ? 'Delete Rows' : 'Delete Columns', () => {
    for (const [a, b] of spans(app, axis)) {
      const filtering = !!s.filter || s.tables.some((t) => !!t.filter);
      if (axis === 'col' || !filtering) { shiftCells(s, axis, a, -(b - a + 1)); continue; }
      // In a filtered list only the rows that are showing are deleted, from the bottom up.
      let end = -1;
      for (let r = Math.min(b, (s.usedRange()?.r2 ?? 0) + 1); r >= a - 1; r--) {
        const keep = r < a || filteredOut(s, r);
        if (!keep && end < 0) end = r;
        if (keep && end >= 0) { shiftCells(s, 'row', r + 1, -(end - r)); end = -1; }
      }
    }
  });
  const { r, c } = activeCell(app);
  app.grid.select(r, c);
}

def([
  {
    id: 'cells.insert', label: 'Insert', icon: 'add_box', tab: 'Home', group: 'Cells', kind: 'split', popular: true, edits: true,
    run: (app, what) => {
      if (typeof what === 'string') return app.exec(what);
      const rg = activeRange(app);
      if (isWholeRows(rg)) return insertWhole(app, 'row');
      if (isWholeCols(rg)) return insertWhole(app, 'col');
      app.exec('cells.insertDialog');
    },
    menu: () => [{ id: 'cells.insertDialog', label: 'Insert Cells…', icon: 'add_box' }, { id: 'cells.insertRows', label: 'Insert Sheet Rows', icon: 'table_rows' }, { id: 'cells.insertCols', label: 'Insert Sheet Columns', icon: 'view_column' }, { separator: true }, { id: 'sheet.insert', label: 'Insert Sheet', icon: 'add' }],
  },
  {
    id: 'cells.delete', label: 'Delete', icon: 'disabled_by_default', tab: 'Home', group: 'Cells', kind: 'split', popular: true, edits: true,
    run: (app, what) => {
      if (typeof what === 'string') return app.exec(what);
      const rg = activeRange(app);
      if (isWholeRows(rg)) return deleteWhole(app, 'row');
      if (isWholeCols(rg)) return deleteWhole(app, 'col');
      app.exec('cells.deleteDialog');
    },
    menu: () => [{ id: 'cells.deleteDialog', label: 'Delete Cells…', icon: 'disabled_by_default' }, { id: 'cells.deleteRows', label: 'Delete Sheet Rows', icon: 'table_rows' }, { id: 'cells.deleteCols', label: 'Delete Sheet Columns', icon: 'view_column' }, { separator: true }, { id: 'sheet.delete', label: 'Delete Sheet', icon: 'delete' }],
  },
  { id: 'cells.insertRows', label: 'Insert Sheet Rows', icon: 'table_rows', tab: 'Home', group: 'Cells', keys: ['Mod-Shift-='], popular: true, edits: true, run: (app, arg) => { const rg = activeRange(app); if (arg !== 'force' && !isWholeRows(rg) && !isWholeCols(rg) && app.lastCommand?.id === 'cells.insertRows' && false) return; insertWhole(app, isWholeCols(rg) && !isWholeRows(rg) ? 'col' : 'row'); } },
  { id: 'cells.insertCols', label: 'Insert Sheet Columns', icon: 'view_column', tab: 'Home', group: 'Cells', popular: true, edits: true, run: (app) => insertWhole(app, 'col') },
  { id: 'cells.deleteRows', label: 'Delete Sheet Rows', icon: 'table_rows', tab: 'Home', group: 'Cells', keys: ['Mod--'], popular: true, edits: true, run: (app) => { const rg = activeRange(app); deleteWhole(app, isWholeCols(rg) && !isWholeRows(rg) ? 'col' : 'row'); } },
  { id: 'cells.deleteCols', label: 'Delete Sheet Columns', icon: 'view_column', tab: 'Home', group: 'Cells', edits: true, run: (app) => deleteWhole(app, 'col') },
  {
    id: 'cells.insertDialog', label: 'Insert Cells…', icon: 'add_box', tab: 'Home', group: 'Cells', edits: true,
    run: async (app) => {
      const r = await openDialog({ title: 'Insert', fields: [{ key: 'how', type: 'radio', label: 'Insert', value: 'down', options: [{ value: 'right', label: 'Shift cells right' }, { value: 'down', label: 'Shift cells down' }, { value: 'row', label: 'Entire row' }, { value: 'col', label: 'Entire column' }] }] });
      if (!r) return;
      if (r.how === 'row') return insertWhole(app, 'row');
      if (r.how === 'col') return insertWhole(app, 'col');
      if (protectedSheet(app, 'insertRows')) return;
      app.book.transact('Insert Cells', () => insertCellsShift(app.sheet, activeRange(app), r.how));
    },
  },
  {
    id: 'cells.deleteDialog', label: 'Delete Cells…', icon: 'disabled_by_default', tab: 'Home', group: 'Cells', edits: true,
    run: async (app) => {
      const r = await openDialog({ title: 'Delete', fields: [{ key: 'how', type: 'radio', label: 'Delete', value: 'up', options: [{ value: 'left', label: 'Shift cells left' }, { value: 'up', label: 'Shift cells up' }, { value: 'row', label: 'Entire row' }, { value: 'col', label: 'Entire column' }] }] });
      if (!r) return;
      if (r.how === 'row') return deleteWhole(app, 'row');
      if (r.how === 'col') return deleteWhole(app, 'col');
      if (protectedSheet(app, 'deleteRows')) return;
      app.book.transact('Delete Cells', () => insertCellsShift(app.sheet, activeRange(app), r.how));
    },
  },
  {
    id: 'cells.format', label: 'Format', icon: 'format_shapes', tab: 'Home', group: 'Cells', kind: 'menu', edits: true, run: () => undefined,
    menu: (app) => [
      { heading: 'Cell Size' },
      { id: 'format.rowHeightDialog', label: 'Row Height…', icon: 'height' }, { id: 'format.autofitRows', label: 'AutoFit Row Height', icon: 'fit_screen' },
      { id: 'format.colWidthDialog', label: 'Column Width…', icon: 'width' }, { id: 'format.autofitCols', label: 'AutoFit Column Width', icon: 'fit_width' }, { id: 'format.defaultWidth', label: 'Default Width…', icon: 'width_normal' },
      { heading: 'Visibility' },
      { label: 'Hide & Unhide', icon: 'visibility', submenu: [{ id: 'format.hideRows', label: 'Hide Rows' }, { id: 'format.hideCols', label: 'Hide Columns' }, { id: 'sheet.hide', label: 'Hide Sheet' }, { id: 'format.unhideRows', label: 'Unhide Rows' }, { id: 'format.unhideCols', label: 'Unhide Columns' }, { id: 'sheet.unhide', label: 'Unhide Sheet…' }] },
      { heading: 'Organize Sheets' },
      { id: 'sheet.rename', label: 'Rename Sheet', icon: 'edit' }, { id: 'sheet.moveCopy', label: 'Move or Copy Sheet…', icon: 'drive_file_move' }, { id: 'sheet.tabColor', label: 'Tab Color', icon: 'palette' },
      { heading: 'Protection' },
      { id: 'review.protectSheet', label: app.sheet.protection.on ? 'Unprotect Sheet…' : 'Protect Sheet…', icon: 'lock' }, { id: 'format.lockCell', label: 'Lock Cell', icon: 'lock_open', checked: app.sheet.styleOf(app.sheet.sel.r, app.sheet.sel.c).locked !== false },
      { id: 'fmt.cellsDialog', label: 'Format Cells…', icon: 'table_view' },
    ],
  },
]);

// --------------------------------------------------------------------------------------------
// Sizes and visibility

function selCols(app: App): number[] {
  const out = new Set<number>();
  for (const rg of selRanges(app)) for (let c = rg.c1; c <= Math.min(rg.c2, rg.c1 + 2000); c++) out.add(c);
  return [...out];
}
function selRows(app: App): number[] {
  const out = new Set<number>();
  for (const rg of selRanges(app)) for (let r = rg.r1; r <= Math.min(rg.r2, rg.r1 + 100000); r++) out.add(r);
  return [...out];
}

/** Width in px that fits a column's contents (AutoFit). */
export function fitWidth(app: App, sheet: Sheet, c: number, rows?: Range): number {
  const ctx = app.grid.ctx;
  const base = sheet.book.styles[0] || {};
  let w = 0;
  const rg = rows || { r1: 0, c1: c, r2: MAX_ROWS - 1, c2: c };
  sheet.forEachIn({ r1: rg.r1, c1: c, r2: rg.r2, c2: c }, (r, _c, cell) => {
    if (cell.v == null || sheet.mergeAt(r, c) || sheet.hiddenRows.has(r)) return;
    const st = sheet.styleOf(r, c);
    if (st.wrap) return;
    ctx.font = fontCss(st, 1, base);
    const text = sheet.showFormulas && cell.f ? '=' + cell.f : formatValue(cell.v, st.numFmt, { date1904: sheet.book.calc.date1904 }).text;
    w = Math.max(w, ctx.measureText(text).width + 8 + (st.indent || 0) * 9);
  });
  return Math.ceil(Math.min(Math.max(w, 12), 255 * 7));
}

export function fitHeight(app: App, sheet: Sheet, r: number): number {
  const ctx = app.grid.ctx;
  const base = sheet.book.styles[0] || {};
  let hgt = 0;
  sheet.forEachIn({ r1: r, c1: 0, r2: r, c2: MAX_COLS - 1 }, (_r, c, cell) => {
    if (sheet.mergeAt(r, c)) return;
    const st = sheet.styleOf(r, c);
    const size = (st.size || base.size || 11) * 4 / 3;
    let lines = 1;
    const text = cell.v == null ? '' : formatValue(cell.v, st.numFmt).text;
    if (st.wrap && text) {
      ctx.font = fontCss(st, 1, base);
      const width = sheet.colWidth(c) - 6;
      let line = '';
      lines = 0;
      for (const para of text.split('\n')) { line = ''; lines++; for (const word of para.split(/(\s+)/)) { if (ctx.measureText(line + word).width > width && line) { lines++; line = word.trimStart(); } else line += word; } }
    } else if (text.includes('\n') && st.wrap) lines = text.split('\n').length;
    if (st.rotate && st.rotate !== 255) { ctx.font = fontCss(st, 1, base); hgt = Math.max(hgt, Math.abs(Math.sin((st.rotate * Math.PI) / 180)) * ctx.measureText(text).width + size); }
    else if (st.rotate === 255) hgt = Math.max(hgt, text.length * size * 1.05 + 6);
    else hgt = Math.max(hgt, lines * size * 1.2 + 4);
  });
  return Math.ceil(Math.max(hgt, sheet.defaultRowH));
}

def([
  {
    id: 'format.colWidth', label: 'Set Column Width', icon: 'width', tab: 'Home', group: 'Cells', internal: true, edits: true,
    run: (app, arg: { cols: number[]; px: number }) => {
      if (protectedSheet(app, 'formatColumns')) return;
      const s = app.sheet;
      app.book.transact('Column Width', () => {
        app.book.recordSheet(s);
        for (const c of arg.cols) {
          if (arg.px <= 0) { s.hiddenCols.add(c); continue; }
          s.hiddenCols.delete(c);
          if (Math.round(arg.px) === s.defaultColW) s.colW.delete(c); else s.colW.set(c, Math.round(arg.px));
        }
      });
    },
  },
  {
    id: 'format.rowHeight', label: 'Set Row Height', icon: 'height', tab: 'Home', group: 'Cells', internal: true, edits: true,
    run: (app, arg: { rows: number[]; px: number }) => {
      if (protectedSheet(app, 'formatRows')) return;
      const s = app.sheet;
      app.book.transact('Row Height', () => {
        app.book.recordSheet(s);
        for (const r of arg.rows) {
          if (arg.px <= 0) { s.hiddenRows.add(r); continue; }
          s.hiddenRows.delete(r);
          if (Math.round(arg.px) === s.defaultRowH) s.rowH.delete(r); else s.rowH.set(r, Math.round(arg.px));
        }
      });
    },
  },
  {
    id: 'format.colWidthDialog', label: 'Column Width…', icon: 'width', tab: 'Home', group: 'Cells', edits: true,
    run: async (app) => {
      const cols = selCols(app);
      const cur = app.sheet.colWidth(cols[0]);
      const v = await promptDialog('Column Width', 'Column width (characters)', String(+((cur - 5) / 7).toFixed(2)), { type: 'number', min: 0, max: 255, step: 0.01 } as any);
      if (v == null) return;
      const chars = Math.max(0, Math.min(255, Number(v)));
      app.exec('format.colWidth', { cols, px: chars ? Math.round(chars * 7 + 5) : 0 });
    },
  },
  {
    id: 'format.rowHeightDialog', label: 'Row Height…', icon: 'height', tab: 'Home', group: 'Cells', edits: true,
    run: async (app) => {
      const rows = selRows(app);
      const cur = app.sheet.rowHeight(rows[0]);
      const v = await promptDialog('Row Height', 'Row height (points)', String(+(cur * 0.75).toFixed(2)), { type: 'number', min: 0, max: 409, step: 0.25 } as any);
      if (v == null) return;
      app.exec('format.rowHeight', { rows, px: Math.round(Math.max(0, Math.min(409, Number(v))) / 0.75) });
    },
  },
  {
    id: 'format.autofitCols', label: 'AutoFit Column Width', icon: 'fit_width', tab: 'Home', group: 'Cells', edits: true,
    run: (app, arg?: { cols?: number[] }) => {
      const s = app.sheet;
      const cols = arg?.cols || selCols(app);
      const rg = activeRange(app);
      const rows = isWholeCols(rg) || arg?.cols ? undefined : rg;
      app.book.transact('AutoFit Column Width', () => {
        app.book.recordSheet(s);
        for (const c of cols.slice(0, 2000)) { const w = fitWidth(app, s, c, rows); s.hiddenCols.delete(c); if (w === s.defaultColW) s.colW.delete(c); else s.colW.set(c, w); }
      });
    },
  },
  {
    id: 'format.autofitRows', label: 'AutoFit Row Height', icon: 'fit_screen', tab: 'Home', group: 'Cells', edits: true,
    run: (app, arg?: { rows?: number[]; quiet?: boolean }) => {
      const s = app.sheet;
      const rows = (arg?.rows || selRows(app)).filter((r) => r <= (s.usedRange()?.r2 ?? 0));
      app.book.transact('AutoFit Row Height', () => {
        app.book.recordSheet(s);
        for (const r of rows.slice(0, 20000)) { const hgt = fitHeight(app, s, r); if (hgt <= s.defaultRowH) s.rowH.delete(r); else s.rowH.set(r, hgt); }
      });
    },
  },
  {
    id: 'format.defaultWidth', label: 'Default Width…', icon: 'width_normal', tab: 'Home', group: 'Cells', edits: true,
    run: async (app) => {
      const s = app.sheet;
      const v = await promptDialog('Standard Width', 'Standard column width (characters)', String(+((s.defaultColW - 5) / 7).toFixed(2)));
      if (v == null) return;
      app.book.transact('Default Width', () => { app.book.recordSheet(s); s.defaultColW = Math.max(5, Math.round(Number(v) * 7 + 5)); });
    },
  },
  { id: 'format.hideRows', label: 'Hide Rows', icon: 'visibility_off', tab: 'Home', group: 'Cells', keys: ['Mod-9'], edits: true, run: (app) => app.exec('format.rowHeight', { rows: selRows(app), px: 0 }) },
  { id: 'format.hideCols', label: 'Hide Columns', icon: 'visibility_off', tab: 'Home', group: 'Cells', keys: ['Mod-0'], edits: true, run: (app) => app.exec('format.colWidth', { cols: selCols(app), px: 0 }) },
  {
    id: 'format.unhideRows', label: 'Unhide Rows', icon: 'visibility', tab: 'Home', group: 'Cells', keys: ['Mod-Shift-9', 'Mod-('], edits: true,
    run: (app) => {
      const s = app.sheet;
      let rows = selRows(app);
      // A single selected row next to hidden ones: include the hidden neighbors (e.g. row 1 hidden, A2 selected).
      if (rows.length === 1) { const r = rows[0]; const near: number[] = []; for (let x = r - 1; x >= 0 && s.hiddenRows.has(x); x--) near.push(x); rows = [...rows, ...near]; }
      app.book.transact('Unhide Rows', () => { app.book.recordSheet(s); for (const r of rows) s.hiddenRows.delete(r); });
    },
  },
  {
    id: 'format.unhideCols', label: 'Unhide Columns', icon: 'visibility', tab: 'Home', group: 'Cells', keys: ['Mod-Shift-0', 'Mod-)'], edits: true,
    run: (app) => {
      const s = app.sheet;
      let cols = selCols(app);
      if (cols.length === 1) { const c = cols[0]; const near: number[] = []; for (let x = c - 1; x >= 0 && s.hiddenCols.has(x); x--) near.push(x); cols = [...cols, ...near]; }
      app.book.transact('Unhide Columns', () => { app.book.recordSheet(s); for (const c of cols) s.hiddenCols.delete(c); });
    },
  },
  {
    id: 'format.lockCell', label: 'Lock Cell', icon: 'lock_open', tab: 'Home', group: 'Cells', kind: 'toggle', edits: true,
    isActive: (app) => app.sheet.styleOf(app.sheet.sel.r, app.sheet.sel.c).locked !== false,
    run: async (app) => { const { applyStyle } = await import('./ops'); const locked = app.sheet.styleOf(app.sheet.sel.r, app.sheet.sel.c).locked !== false; applyStyle(app, { locked: locked ? false : undefined }, locked ? 'Unlock Cells' : 'Lock Cells'); },
  },
]);

// --------------------------------------------------------------------------------------------
// Sheets

def([
  {
    id: 'sheet.activate', label: 'Activate Sheet', icon: 'tab', tab: 'Sheet', internal: true, noRecord: false,
    run: (app, i: number | string) => {
      const book = app.book;
      const idx = typeof i === 'string' ? book.sheets.findIndex((s) => s.name === i) : i;
      if (idx < 0 || idx >= book.sheets.length) return;
      if (app.grid.editing && !app.grid.editor.canPoint()) app.grid.commitEdit('none');
      book.activeSheet = idx;
      app.grid.objects.select(null);
      app.grid.axisVersion = '';
      app.grid.extent = { rows: 200, cols: 40 };
      app.grid.updateExtent();
      const s = book.sheet;
      app.grid.scroller.scrollTop = Math.max(0, (app.grid.rows.offset(s.scroll.r) - app.grid.rows.offset(s.freeze.rows)) * s.zoom);
      app.grid.scroller.scrollLeft = Math.max(0, (app.grid.cols.offset(s.scroll.c) - app.grid.cols.offset(s.freeze.cols)) * s.zoom);
      app.grid.invalidate();
      app.refresh();
    },
  },
  { id: 'sheet.next', label: 'Next Sheet', icon: 'chevron_right', tab: 'Sheet', keys: ['Mod-PageDown'], run: (app) => { const b = app.book; for (let i = b.activeSheet + 1; i < b.sheets.length; i++) if (b.sheets[i].visibility === 'visible') return app.exec('sheet.activate', i); } },
  { id: 'sheet.prev', label: 'Previous Sheet', icon: 'chevron_left', tab: 'Sheet', keys: ['Mod-PageUp'], run: (app) => { const b = app.book; for (let i = b.activeSheet - 1; i >= 0; i--) if (b.sheets[i].visibility === 'visible') return app.exec('sheet.activate', i); } },
  {
    id: 'sheet.insert', label: 'Insert Sheet', icon: 'add', tab: 'Sheet', keys: ['Shift-F11', 'Alt-Shift-F1'], popular: true, edits: true,
    run: (app) => {
      if (protectedStructure(app)) return;
      const book = app.book;
      book.transact('Insert Sheet', () => {
        book.recordBook();
        const s = new Sheet(book, book.uniqueSheetName());
        book.sheets.splice(book.activeSheet + 1, 0, s);
        book.activeSheet = book.activeSheet + 1;
      });
      app.exec('sheet.activate', book.activeSheet);
    },
  },
  {
    id: 'sheet.delete', label: 'Delete Sheet', icon: 'delete', tab: 'Sheet', edits: true,
    run: async (app) => {
      if (protectedStructure(app)) return;
      const book = app.book;
      if (book.sheets.filter((s) => s.visibility === 'visible').length <= 1) { toast('A workbook must contain at least one visible worksheet.'); return; }
      const s = book.sheet;
      if (s.cells.size && !(await confirmDialog('Delete Sheet', `Permanently delete "${s.name}"? Formulas that refer to it will show #REF!.`, 'Delete'))) return;
      book.transact('Delete Sheet', () => {
        book.recordBook();
        book.sheets.splice(book.activeSheet, 1);
        invalidateSheetRefs(book, s.name);
        book.activeSheet = Math.max(0, Math.min(book.activeSheet, book.sheets.length - 1));
        while (book.sheets[book.activeSheet].visibility !== 'visible' && book.activeSheet > 0) book.activeSheet--;
      });
      app.exec('sheet.activate', book.activeSheet);
    },
  },
  {
    id: 'sheet.rename', label: 'Rename Sheet', icon: 'edit', tab: 'Sheet', edits: true,
    run: async (app, i?: number) => {
      if (protectedStructure(app)) return;
      const book = app.book;
      const s = book.sheets[typeof i === 'number' ? i : book.activeSheet];
      const name = await promptDialog('Rename Sheet', 'Sheet name', s.name);
      if (name == null || name === s.name) return;
      const err = validSheetName(name);
      if (err) { toast(err); return; }
      if (book.sheets.some((x) => x !== s && x.name.toLowerCase() === name.toLowerCase())) { toast('That name is already taken. Try a different one.'); return; }
      book.transact('Rename Sheet', () => renameSheet(book, s, name));
    },
  },
  {
    id: 'sheet.duplicate', label: 'Duplicate Sheet', icon: 'content_copy', tab: 'Sheet', edits: true,
    run: (app) => app.exec('sheet.move', { from: app.book.activeSheet, to: app.book.activeSheet + 1, copy: true }),
  },
  {
    id: 'sheet.move', label: 'Move Sheet', icon: 'drive_file_move', tab: 'Sheet', internal: true, edits: true,
    run: (app, arg: { from: number; to: number; copy?: boolean }) => {
      if (protectedStructure(app)) return;
      const book = app.book;
      book.transact(arg.copy ? 'Copy Sheet' : 'Move Sheet', () => {
        book.recordBook();
        const src = book.sheets[arg.from];
        if (arg.copy) {
          const j = JSON.parse(JSON.stringify(src.toJSON()));
          delete j.id;
          let n = 2, name = `${src.name} (${n})`;
          while (book.sheetByName(name)) name = `${src.name} (${++n})`;
          j.name = name;
          const copy = Sheet.fromJSON(book, j);
          copy.tables = copy.tables.map((t) => ({ ...t, id: 't' + Math.random().toString(36).slice(2), name: `${t.name}${n}` }));
          book.sheets.splice(Math.min(arg.to, book.sheets.length), 0, copy);
          book.activeSheet = book.sheets.indexOf(copy);
        } else {
          const [s] = book.sheets.splice(arg.from, 1);
          book.sheets.splice(Math.min(arg.to, book.sheets.length), 0, s);
          book.activeSheet = book.sheets.indexOf(s);
        }
      });
      app.exec('sheet.activate', book.activeSheet);
    },
  },
  {
    id: 'sheet.moveCopy', label: 'Move or Copy Sheet…', icon: 'drive_file_move', tab: 'Sheet', edits: true,
    run: async (app) => {
      const book = app.book;
      const targets = app.docs.map((d) => ({ value: d.id, label: d.displayName }));
      const r = await openDialog({ title: 'Move or Copy', fields: [
        { key: 'book', label: 'To book', type: 'select', options: [...targets, { value: 'new', label: '(new book)' }], value: app.tab.id },
        { key: 'before', label: 'Before sheet', type: 'list', rows: 8, options: [...book.sheets.map((s, i) => ({ value: String(i), label: s.name })), { value: String(book.sheets.length), label: '(move to end)' }], value: String(book.activeSheet + 1) },
        { key: 'copy', label: 'Create a copy', type: 'checkbox', value: false },
      ] });
      if (!r) return;
      const to = Number(Array.isArray(r.before) ? r.before[0] : r.before);
      if (r.book === app.tab.id) return app.exec('sheet.move', { from: book.activeSheet, to: r.copy || to <= book.activeSheet ? to : to - 1, copy: r.copy });
      // To another (or a new) workbook
      const src = book.sheet;
      const json = JSON.parse(JSON.stringify(src.toJSON()));
      delete json.id;
      const target = r.book === 'new' ? app.createTab() : app.docs.find((d) => d.id === r.book)!;
      const tb = target.book;
      tb.transact('Copy Sheet', () => {
        tb.recordBook();
        // styles are per-workbook: re-intern them
        json.cells = json.cells.map(([k, c]: [number, any]) => [k, c.s ? { ...c, s: tb.styleId(book.styles[c.s] || {}) } : c]);
        let name = json.name; let n = 2; while (tb.sheetByName(name)) name = `${json.name} (${n++})`; json.name = name;
        const sh = Sheet.fromJSON(tb, json);
        if (r.book === 'new' && tb.sheets.length === 1 && !tb.sheets[0].cells.size) tb.sheets = [sh]; else tb.sheets.push(sh);
        tb.activeSheet = tb.sheets.indexOf(sh);
      });
      if (!r.copy) { app.activate(app.docs.find((d) => d.book === book)!); app.exec('sheet.delete'); }
      app.activate(target);
    },
  },
  {
    id: 'sheet.tabColor', label: 'Tab Color', icon: 'palette', tab: 'Sheet', edits: true,
    run: async (app) => {
      const c = await colorMenu(app.ui.sheetBar.querySelector('.ql-sheettab.ql-active') as HTMLElement || app.ui.sheetBar, { none: 'No Color' });
      if (c === undefined) return;
      app.book.transact('Tab Color', () => { app.book.recordBook(); app.sheet.tabColor = c; });
    },
  },
  {
    id: 'sheet.hide', label: 'Hide Sheet', icon: 'visibility_off', tab: 'Sheet', edits: true,
    run: (app) => {
      if (protectedStructure(app)) return;
      const book = app.book;
      if (book.sheets.filter((s) => s.visibility === 'visible').length <= 1) { toast('A workbook must contain at least one visible worksheet.'); return; }
      book.transact('Hide Sheet', () => { book.recordBook(); book.sheet.visibility = 'hidden'; });
      const next = book.sheets.findIndex((s) => s.visibility === 'visible');
      app.exec('sheet.activate', next);
    },
  },
  {
    id: 'sheet.unhide', label: 'Unhide Sheet…', icon: 'visibility', tab: 'Sheet', edits: true,
    run: async (app) => {
      if (protectedStructure(app)) return;
      const book = app.book;
      const hidden = book.sheets.filter((s) => s.visibility === 'hidden');
      if (!hidden.length) { toast('There are no hidden sheets.'); return; }
      const r = await openDialog({ title: 'Unhide', fields: [{ key: 'names', label: 'Unhide sheet', type: 'list', multiple: true, rows: 8, options: hidden.map((s) => s.name), value: [hidden[0].name] }] });
      if (!r?.names?.length) return;
      book.transact('Unhide Sheet', () => { book.recordBook(); for (const s of hidden) if (r.names.includes(s.name)) s.visibility = 'visible'; });
      app.exec('sheet.activate', book.sheets.findIndex((s) => s.name === r.names[0]));
    },
  },
  { id: 'sheet.selectAll', label: 'Select All Sheets', icon: 'select_all', tab: 'Sheet', run: () => toast('Grouped sheets: use Fill ▸ Across Worksheets to copy cells to every sheet.') },
  {
    id: 'sheet.fillAcross', label: 'Fill Across Worksheets…', icon: 'layers', tab: 'Sheet', edits: true,
    run: async (app) => {
      const book = app.book;
      const others = book.sheets.filter((s) => s !== app.sheet);
      if (!others.length) { toast('There is only one sheet.'); return; }
      const r = await openDialog({ title: 'Fill Across Worksheets', fields: [{ key: 'to', label: 'Sheets', type: 'list', multiple: true, rows: 8, options: others.map((s) => s.name), value: others.map((s) => s.name) }, { key: 'what', label: 'Fill', type: 'radio', options: [{ value: 'all', label: 'All' }, { value: 'contents', label: 'Contents' }, { value: 'formats', label: 'Formats' }], value: 'all' }] });
      if (!r?.to?.length) return;
      const rg = clipToUsed(app.sheet, activeRange(app));
      book.transact('Fill Across Worksheets', () => {
        for (const sh of others.filter((s) => r.to.includes(s.name))) {
          for (let rr = rg.r1; rr <= rg.r2; rr++) for (let cc = rg.c1; cc <= rg.c2; cc++) {
            const src = app.sheet.get(rr, cc);
            const cur = sh.get(rr, cc);
            if (r.what === 'formats') sh.patch(rr, cc, { s: src?.s });
            else if (r.what === 'contents') sh.set(rr, cc, src ? { ...(cur?.s ? { s: cur.s } : {}), ...(src.f != null ? { f: src.f } : { v: src.v }) } : (cur?.s ? { s: cur.s } : undefined));
            else sh.set(rr, cc, src ? { ...src, sp: undefined } : undefined);
          }
        }
      });
    },
  },
  {
    id: 'sheet.exportCsv', label: 'Save This Sheet as CSV…', icon: 'download', tab: 'Sheet', noRecord: true,
    run: async (app) => {
      const { exportFile } = await import('../io');
      const { saveAs } = await import('../platform/native');
      const blob = await exportFile(app.book, 'csv', { sheetIndex: app.book.activeSheet });
      await saveAs(blob, `${app.sheet.name}.csv`);
    },
  },
  {
    id: 'sheet.goto', label: 'Activate Sheet…', icon: 'tab', tab: 'Sheet', noRecord: true,
    run: async (app) => {
      const i = await choiceDialog('Activate', 'Choose a sheet', app.book.sheets.filter((s) => s.visibility === 'visible').map((s) => s.name));
      if (i >= 0) app.exec('sheet.activate', app.book.sheets.indexOf(app.book.sheets.filter((s) => s.visibility === 'visible')[i]));
    },
  },
]);

void colName; void keyRow; void keyCol;
