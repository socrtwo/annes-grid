/**
 * Edit commands: cell entry, undo/redo, clipboard (with Paste Special), fill, clear, find & replace, go to,
 * selection helpers and navigation.
 */
import { def, getCommand } from './registry';
import type { App } from '../app';
import { newId, type Sheet } from '../model/workbook';
import type { Cell, CellComment, Scalar, DataValidation } from '../model/types';
import { CellError, ERR } from '../model/types';
import { cellKey, keyRow, keyCol, MAX_COLS, MAX_ROWS, parseRange, splitSheetRef, normRange, rangeName, cellName, type Range } from '../model/address';
import { shiftFormula, mapRefs } from '../calc/parser';
import { moveBlock, resolveRangeText } from '../model/structure';
import { hasLotusSyntax } from '../heritage/lotus';
import { selRanges, activeRange, activeCell, setInput, parseInput, clearRange, fillRange, forEachSelected, clipToUsed, rangeToTsv, parseTsv, currentRegion, isWholeCols, isWholeRows, addrText, type ClearWhat } from './ops';
import { openDialog, toast, showMenu, alertDialog, choiceDialog } from '../ui/dialog';
import { readClipboardText, writeClipboardText } from '../platform/native';
import { formatValue } from '../model/numfmt';
import { compare, criteria } from '../calc/values';
import { escHtml } from '../util/sanitize';

// --------------------------------------------------------------------------------------------
// Cell entry and validation

function validationFor(sheet: Sheet, r: number, c: number): DataValidation | undefined {
  return sheet.validations.find((v) => v.ranges.some((rg) => r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2));
}

/** Checks a value against a data validation rule. */
export function validates(app: App, sheet: Sheet, r: number, c: number, v: Scalar, dv: DataValidation): boolean {
  if (dv.type === 'any') return true;
  if (v == null || v === '') return dv.allowBlank !== false;
  const eng = app.tab.engine;
  const base = { r: dv.ranges[0].r1, c: dv.ranges[0].c1 };
  const val = (f?: string): Scalar => (f == null ? null : eng.evaluateScalar(f, sheet, r, c, base));
  if (dv.type === 'list') {
    const items = listItems(app, sheet, dv, r, c);
    return items.some((x) => compare(x, v) === 0 || (typeof v === 'number' && String(v) === String(x)));
  }
  if (dv.type === 'custom') {
    // The rule judges the value being entered, so evaluate it with that value in the cell.
    const k = cellKey(r, c);
    const cur = sheet.cells.get(k);
    sheet.cells.set(k, { ...(cur?.s ? { s: cur.s } : {}), v });
    let x: Scalar;
    try { x = val(dv.formula1); } finally { if (cur) sheet.cells.set(k, cur); else sheet.cells.delete(k); }
    return x === true || (typeof x === 'number' && x !== 0);
  }
  let x: number;
  if (dv.type === 'textLength') x = String(v).length;
  else if (typeof v === 'number') x = v;
  else return false;
  if (dv.type === 'whole' && !Number.isInteger(x)) return false;
  const a = Number(val(dv.formula1)), b = Number(val(dv.formula2));
  switch (dv.operator || 'between') {
    case 'between': return x >= a && x <= b;
    case 'notBetween': return x < a || x > b;
    case 'equal': return x === a;
    case 'notEqual': return x !== a;
    case 'greaterThan': return x > a;
    case 'lessThan': return x < a;
    case 'greaterThanOrEqual': return x >= a;
    case 'lessThanOrEqual': return x <= a;
  }
  return true;
}

export function listItems(app: App, sheet: Sheet, dv: DataValidation, r: number, c: number): Scalar[] {
  const f = (dv.formula1 || '').trim();
  if (f.startsWith('"') && f.endsWith('"')) return f.slice(1, -1).split(/\s*,\s*/).map((x) => x);
  if (!f.startsWith('=') && f.includes(',') && !/[!$:]/.test(f)) return f.split(/\s*,\s*/);
  const m = app.tab.engine.evaluateMatrix(f.replace(/^=/, ''), sheet, r, c);
  return m.flat().filter((x) => x != null && x !== '');
}

def([
  {
    id: 'cell.canEdit', label: 'Can Edit Cell', tab: 'Home', internal: true, noRecord: true,
    run: (app, arg: { r: number; c: number }) => {
      const s = app.sheet;
      if (s.protection.on) {
        const st = s.styleOf(arg.r, arg.c);
        // Locked cells stay editable inside an Allow Edit Ranges range (review.canEditRange shows the notice or password prompt).
        if (st.locked !== false && app.exec('review.canEditRange', arg) !== true) return false;
      }
      const cell = s.get(arg.r, arg.c);
      if (cell?.sp != null && cell.f == null && !cell.arr) {
        // Part of a legacy array: cannot change part of an array.
        const anchor = s.cells.get(cell.sp);
        if (anchor?.arr) { toast("You can't change part of an array."); return false; }
      }
      return true;
    },
  },
  {
    id: 'cell.enter', label: 'Enter Cell Value', icon: 'keyboard_return', tab: 'Home', group: 'Editing', internal: true, edits: true,
    run: (app, arg: { r: number; c: number; text: string; array?: boolean; fill?: boolean }) => {
      const s = app.sheet;
      let text = arg.text;
      // AutoCorrect replacements for plain text
      if (app.prefs.autoCorrect && !text.startsWith('=') && !(app.prefs.lotusFormulas && /^[@+-]/.test(text) && hasLotusSyntax(text)) && /\S/.test(text)) {
        for (const [k, v] of Object.entries(app.prefs.replacements)) if (text.includes(k)) text = text.split(k).join(v);
      }
      const dv = validationFor(s, arg.r, arg.c);
      if (dv && !arg.array) {
        const p = parseInput(app, text, arg.r, arg.c);
        const v = p.cell.f != null ? null : p.cell.v ?? null;
        if (p.cell.f == null && !validates(app, s, arg.r, arg.c, v, dv)) {
          const msg = dv.errorMessage || 'This value doesn’t match the data validation restrictions defined for this cell.';
          if ((dv.errorStyle || 'stop') === 'stop') { void alertDialog(dv.errorTitle || 'Invalid Entry', msg); return false; }
          if (!confirm(`${dv.errorTitle || (dv.errorStyle === 'warning' ? 'Warning' : 'Information')}\n\n${msg}\n\nContinue?`)) return false;
        }
      }
      const book = app.book;
      book.transact('Typing', () => {
        if (arg.array) {
          const rg = activeRange(app);
          const f = text.replace(/^\{?=/, '').replace(/\}$/, '');
          for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) if (r !== rg.r1 || c !== rg.c1) s.set(r, c, s.get(r, c)?.s ? { s: s.get(r, c)!.s } : undefined);
          const cur = s.get(rg.r1, rg.c1);
          s.set(rg.r1, rg.c1, { ...(cur?.s ? { s: cur.s } : {}), f, arr: rg });
          return;
        }
        if (arg.fill) {
          // Ctrl+Enter: the same entry in every selected cell (formulas adjust relative references).
          const p0 = parseInput(app, text, arg.r, arg.c);
          forEachSelected(app, (r, c) => {
            if (p0.cell.f != null) {
              const cur = s.get(r, c);
              s.set(r, c, { ...(cur?.s ? { s: cur.s } : {}), f: shiftFormula(p0.cell.f, r - arg.r, c - arg.c) });
            } else setInput(app, s, r, c, text);
          }, { clip: selRanges(app).some((x) => isWholeCols(x) || isWholeRows(x)) }); // whole rows and columns stop at the data
          return;
        }
        setInput(app, s, arg.r, arg.c, text);
        if (app.prefs.flashFill && text && !text.startsWith('=')) setTimeout(() => { if (getCommand('data.flashFillSuggest')) app.exec('data.flashFillSuggest', { r: arg.r, c: arg.c }); }, 0);
        // Extend list formats and formulas: a new row at the bottom of a table grows the table.
        for (const t of s.tables) {
          if (t.opts.header && arg.r === t.range.r1 && arg.c >= t.range.c1 && arg.c <= t.range.c2) {
            const name = String(s.value(arg.r, arg.c) ?? '').trim();
            const col = t.columns[arg.c - t.range.c1];
            if (col && name && col.name !== name) { book.recordSheet(s); col.name = name; }
          }
          if (arg.r === t.range.r2 + 1 && arg.c >= t.range.c1 && arg.c <= t.range.c2 && !t.opts.totals) {
            book.recordSheet(s);
            t.range = { ...t.range, r2: arg.r };
            for (let c = t.range.c1; c <= t.range.c2; c++) {
              if (c === arg.c) continue;
              const above = s.get(arg.r - 1, c);
              const colDef = t.columns[c - t.range.c1];
              if (colDef?.calcFormula) s.set(arg.r, c, { f: colDef.calcFormula, ...(above?.s ? { s: above.s } : {}) });
              else if (above?.s && !s.get(arg.r, c)) s.set(arg.r, c, { s: above.s });
            }
          }
          if (arg.c === t.range.c2 + 1 && arg.r >= t.range.r1 && arg.r <= t.range.r2 && t.opts.header && arg.r === t.range.r1) {
            book.recordSheet(s);
            t.range = { ...t.range, c2: arg.c };
            t.columns.push({ name: String(s.value(arg.r, arg.c) ?? `Column${t.columns.length + 1}`) });
          }
          // A formula typed into an empty table column becomes a calculated column.
          if (arg.r > t.range.r1 && arg.r <= t.range.r2 && arg.c >= t.range.c1 && arg.c <= t.range.c2 && text.startsWith('=')) {
            const col = arg.c;
            let empty = true;
            for (let r = t.range.r1 + (t.opts.header ? 1 : 0); r <= t.range.r2 - (t.opts.totals ? 1 : 0); r++) if (r !== arg.r && (s.get(r, col)?.v != null || s.get(r, col)?.f != null)) { empty = false; break; }
            if (empty) {
              const f = s.get(arg.r, arg.c)!.f!;
              t.columns[col - t.range.c1].calcFormula = f;
              for (let r = t.range.r1 + (t.opts.header ? 1 : 0); r <= t.range.r2 - (t.opts.totals ? 1 : 0); r++) if (r !== arg.r) { const cur = s.get(r, col); s.set(r, col, { ...(cur?.s ? { s: cur.s } : {}), f: shiftFormula(f, r - arg.r, 0) }); }
            }
          }
        }
      });
      return true;
    },
  },
]);

// --------------------------------------------------------------------------------------------
// Undo / redo / repeat

def([
  { id: 'edit.undo', label: 'Undo', icon: 'undo', tab: 'Home', group: 'Undo', keys: ['Mod-z'], popular: true, noRecord: true, isEnabled: (app) => app.book.canUndo(), run: (app) => { if (app.grid.editing) { app.grid.cancelEdit(); return; } const l = app.book.undo(); if (l) { app.ui.status(`Undo: ${l}`); app.grid.updateExtent(); } } },
  { id: 'edit.redo', label: 'Redo', icon: 'redo', tab: 'Home', group: 'Undo', keys: ['Mod-y', 'Mod-Shift-z'], popular: true, noRecord: true, isEnabled: (app) => app.book.canRedo(), run: (app) => { const l = app.book.redo(); if (l) app.ui.status(`Redo: ${l}`); } },
  { id: 'edit.repeat', label: 'Repeat', icon: 'repeat', tab: 'Home', group: 'Undo', keys: ['F4'], noRecord: true, run: (app) => { if (app.lastCommand && app.lastCommand.id !== 'cell.enter') app.exec(app.lastCommand.id, app.lastCommand.arg); } },
]);

// --------------------------------------------------------------------------------------------
// Clipboard

interface ClipData {
  bookId: string;
  sheet: Sheet;
  range: Range;
  cells: Map<number, Cell>; // keyed by relative (dr, dc)
  comments: Map<number, CellComment>;
  colW: number[];
  rowH: number[];
  validations: DataValidation[];
  cut: boolean;
  text: string;
}

let clip: ClipData | null = null;

function snapshotRange(sheet: Sheet, rg: Range): Pick<ClipData, 'cells' | 'comments' | 'colW' | 'rowH' | 'validations'> {
  const cells = new Map<number, Cell>();
  const comments = new Map<number, CellComment>();
  sheet.forEachIn(rg, (r, c, cell) => {
    const copy = JSON.parse(JSON.stringify(cell, (k, v) => (v instanceof CellError ? { '#err': v.code } : v)));
    if (copy.v && typeof copy.v === 'object' && '#err' in copy.v) copy.v = new CellError(copy.v['#err']);
    cells.set(cellKey(r - rg.r1, c - rg.c1), copy);
  });
  for (const [k, cm] of sheet.comments) { const r = keyRow(k), c = keyCol(k); if (r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2) comments.set(cellKey(r - rg.r1, c - rg.c1), JSON.parse(JSON.stringify(cm))); }
  const colW: number[] = [], rowH: number[] = [];
  for (let c = rg.c1; c <= Math.min(rg.c2, rg.c1 + 200); c++) colW.push(sheet.colWidth(c) || sheet.defaultColW);
  for (let r = rg.r1; r <= Math.min(rg.r2, rg.r1 + 200); r++) rowH.push(sheet.rowHeight(r) || sheet.defaultRowH);
  const validations = sheet.validations.filter((v) => v.ranges.some((x) => x.r1 <= rg.r2 && x.r2 >= rg.r1 && x.c1 <= rg.c2 && x.c2 >= rg.c1));
  return { cells, comments, colW, rowH, validations: JSON.parse(JSON.stringify(validations)) };
}

function rangeToHtml(sheet: Sheet, rg: Range): string {
  const r2 = Math.min(rg.r2, sheet.usedRange()?.r2 ?? rg.r1), c2 = Math.min(rg.c2, sheet.usedRange()?.c2 ?? rg.c1);
  const rows: string[] = [];
  for (let r = rg.r1; r <= Math.max(r2, rg.r1); r++) {
    const tds: string[] = [];
    for (let c = rg.c1; c <= Math.max(c2, rg.c1); c++) {
      const cell = sheet.get(r, c);
      const st = sheet.styleOf(r, c);
      const css = [st.bold ? 'font-weight:bold' : '', st.italic ? 'font-style:italic' : '', st.color ? `color:${st.color}` : '', st.fill ? `background:${st.fill}` : '', st.hAlign && st.hAlign !== 'general' ? `text-align:${st.hAlign}` : typeof cell?.v === 'number' ? 'text-align:right' : ''].filter(Boolean).join(';');
      const txt = cell ? formatValue(cell.v ?? null, st.numFmt, { date1904: sheet.book.calc.date1904 }).text : '';
      tds.push(`<td${css ? ` style="${css}"` : ''}>${escHtml(txt)}</td>`);
    }
    rows.push(`<tr>${tds.join('')}</tr>`);
  }
  return `<table>${rows.join('')}</table>`;
}

async function copyOrCut(app: App, cut: boolean, ev?: ClipboardEvent): Promise<void> {
  const s = app.sheet;
  const g = app.grid;
  const obj = g.objects.selectedObject();
  if (obj) { (app as any).objectClip = JSON.parse(JSON.stringify(obj)); if (cut) app.exec('object.delete', { id: obj.def.id }); toast(cut ? 'Cut' : 'Copied'); return; }
  if (selRanges(app).length > 1) { toast("This action won't work on multiple selections."); return; }
  const rg = clipToUsed(s, activeRange(app));
  const text = rangeToTsv(s, rg);
  const html = rangeToHtml(s, rg);
  clip = { bookId: app.tab.id, sheet: s, range: rg, cut, text, ...snapshotRange(s, rg) };
  app.tab.marquee = { sheet: s, range: activeRange(app), cut };
  app.clipboard.unshift({ text, html, time: Date.now(), source: `${s.name}!${addrText(rg)}` });
  app.clipboard.splice(app.prefs.clipboardHistory);
  if (ev?.clipboardData) { ev.clipboardData.setData('text/plain', text); ev.clipboardData.setData('text/html', html); }
  else {
    try {
      if ((navigator.clipboard as any)?.write && typeof ClipboardItem !== 'undefined') await navigator.clipboard.write([new ClipboardItem({ 'text/plain': new Blob([text], { type: 'text/plain' }), 'text/html': new Blob([html], { type: 'text/html' }) })]);
      else await writeClipboardText(text);
    } catch { await writeClipboardText(text); }
  }
  g.invalidate();
}

export interface PasteOpts {
  what?: 'all' | 'formulas' | 'values' | 'formats' | 'comments' | 'validation' | 'allExceptBorders' | 'colWidths' | 'formulasAndNumFmt' | 'valuesAndNumFmt' | 'allMergingCF' | 'noBorders';
  op?: 'none' | 'add' | 'subtract' | 'multiply' | 'divide';
  skipBlanks?: boolean;
  transpose?: boolean;
  link?: boolean;
  keepWidths?: boolean;
}

/** Pastes from the internal clipboard (formulas, formats, comments...) into the active range. */
function pasteInternal(app: App, cd: ClipData, o: PasteOpts): void {
  const s = app.sheet;
  const book = app.book;
  const dst = activeRange(app);
  const srcRows = cd.range.r2 - cd.range.r1 + 1, srcCols = cd.range.c2 - cd.range.c1 + 1;
  const outRows = o.transpose ? srcCols : srcRows, outCols = o.transpose ? srcRows : srcCols;
  // Tile when the destination is an exact multiple of the source.
  const dRows = dst.r2 - dst.r1 + 1, dCols = dst.c2 - dst.c1 + 1;
  const tileR = dRows % outRows === 0 && dRows > outRows ? dRows / outRows : 1;
  const tileC = dCols % outCols === 0 && dCols > outCols ? dCols / outCols : 1;
  if (dst.r1 + outRows * tileR > MAX_ROWS || dst.c1 + outCols * tileC > MAX_COLS) { toast('The paste area extends beyond the sheet.'); return; }
  // Cut + paste in the same workbook moves the cells (references follow).
  if (cd.cut && !o.transpose && (o.what || 'all') === 'all' && cd.bookId === app.tab.id && app.book.sheets.includes(cd.sheet)) {
    book.transact('Move', () => moveBlock(cd.sheet, cd.range, dst.r1 - cd.range.r1, dst.c1 - cd.range.c1, s));
    clip = null;
    app.tab.marquee = null;
    app.grid.select(dst.r1, dst.c1, { range: { r1: dst.r1, c1: dst.c1, r2: dst.r1 + outRows - 1, c2: dst.c1 + outCols - 1 } });
    return;
  }
  const what = o.what || 'all';
  book.transact('Paste', () => {
    for (let tr = 0; tr < tileR; tr++) for (let tc = 0; tc < tileC; tc++) {
      for (let i = 0; i < srcRows; i++) for (let j = 0; j < srcCols; j++) {
        const src = cd.cells.get(cellKey(i, j));
        const r = dst.r1 + tr * outRows + (o.transpose ? j : i);
        const c = dst.c1 + tc * outCols + (o.transpose ? i : j);
        if (o.skipBlanks && (!src || (src.v == null && src.f == null))) continue;
        if (s.hiddenRows.has(r) && s.filter) continue;
        const cur = s.get(r, c);
        let next: Cell = { ...(cur || {}) };
        const srcR = cd.range.r1 + i, srcC = cd.range.c1 + j;
        if (o.link) {
          const ref = (cd.sheet === s ? '' : `'${cd.sheet.name.replace(/'/g, "''")}'!`) + cellName(srcR, srcC, true, true);
          next = { ...(cur?.s ? { s: cur.s } : {}), f: ref };
          s.set(r, c, next);
          continue;
        }
        const takeContent = ['all', 'formulas', 'values', 'allExceptBorders', 'formulasAndNumFmt', 'valuesAndNumFmt', 'allMergingCF', 'noBorders'].includes(what);
        const takeFormat = ['all', 'formats', 'allExceptBorders', 'allMergingCF', 'noBorders'].includes(what);
        // With an operation, a source cell that is not a plain number leaves the destination's content as it is.
        const opSkip = !!o.op && o.op !== 'none' && !(src && typeof src.v === 'number' && src.f == null);
        if (takeContent && !opSkip) {
          delete next.v; delete next.f; delete next.rich; delete next.sp; delete next.arr; delete next.link;
          if (src) {
            if (src.f != null && what !== 'values' && what !== 'valuesAndNumFmt') {
              let f = shiftFormula(src.f, r - srcR, c - srcC);
              if (o.transpose) f = transposeRefs(src.f, srcR, srcC, r, c);
              next.f = f;
            } else if (src.v != null) next.v = src.v;
            if (src.rich && what !== 'values') next.rich = src.rich;
            if (src.link && what !== 'values') next.link = src.link;
            // Operation: combine with the existing value
            if (o.op && o.op !== 'none' && typeof src.v === 'number' && src.f == null) {
              const base = typeof cur?.v === 'number' ? cur.v : 0;
              if (cur?.f != null) next.f = `(${cur.f})${{ add: '+', subtract: '-', multiply: '*', divide: '/' }[o.op]}${src.v}`;
              else next.v = o.op === 'add' ? base + src.v : o.op === 'subtract' ? base - src.v : o.op === 'multiply' ? base * src.v : src.v === 0 ? ERR.div0 : base / src.v;
              if (cur?.f != null) delete next.v;
            }
          }
        }
        if (takeFormat || what === 'formulasAndNumFmt' || what === 'valuesAndNumFmt') {
          const srcStyle = src?.s ?? cd.sheet.rowStyle.get(srcR) ?? cd.sheet.colStyle.get(srcC) ?? 0;
          if (what === 'formulasAndNumFmt' || what === 'valuesAndNumFmt') {
            const nf = book.styles[srcStyle]?.numFmt;
            next.s = book.derivedStyle(next.s ?? 0, { numFmt: nf }) || undefined;
          } else if (what === 'allExceptBorders' || what === 'noBorders') {
            const st = { ...(book.styles[srcStyle] || {}) };
            const cs = book.styles[cur?.s ?? 0] || {};
            for (const k of ['bt', 'bb', 'bl', 'br', 'diagUp', 'diagDown'] as const) { delete st[k]; if (cs[k]) (st as any)[k] = cs[k]; }
            next.s = book.styleId(st) || undefined;
          } else next.s = srcStyle || undefined;
        }
        if (what === 'comments' || what === 'all') { const cm = cd.comments.get(cellKey(i, j)); if (cm) { book.recordSheet(s); s.comments.set(cellKey(r, c), JSON.parse(JSON.stringify(cm))); } }
        if (next.s === 0) delete next.s;
        s.set(r, c, next);
      }
    }
    if (what === 'validation' || what === 'all') {
      for (const v of cd.validations) {
        const inter = v.ranges.map((x) => ({ r1: Math.max(x.r1, cd.range.r1), c1: Math.max(x.c1, cd.range.c1), r2: Math.min(x.r2, cd.range.r2), c2: Math.min(x.c2, cd.range.c2) })).filter((x) => x.r1 <= x.r2 && x.c1 <= x.c2);
        if (!inter.length) continue;
        book.recordSheet(s);
        s.validations.push({ ...v, id: `dv${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`, ranges: inter.map((x) => ({ r1: dst.r1 + x.r1 - cd.range.r1, c1: dst.c1 + x.c1 - cd.range.c1, r2: dst.r1 + x.r2 - cd.range.r1, c2: dst.c1 + x.c2 - cd.range.c1 })) });
      }
    }
    if (what === 'colWidths' || o.keepWidths) {
      book.recordSheet(s);
      cd.colW.forEach((w, j) => { if (w !== s.defaultColW) s.colW.set(dst.c1 + j, w); else s.colW.delete(dst.c1 + j); });
    }
    if (what === 'all') {
      // Merged cells inside the source come along
      const merges = cd.sheet.merges.filter((m) => m.r1 >= cd.range.r1 && m.r2 <= cd.range.r2 && m.c1 >= cd.range.c1 && m.c2 <= cd.range.c2);
      if (merges.length && !o.transpose) {
        book.recordSheet(s);
        for (const m of merges) s.merges.push({ r1: m.r1 - cd.range.r1 + dst.r1, c1: m.c1 - cd.range.c1 + dst.c1, r2: m.r2 - cd.range.r1 + dst.r1, c2: m.c2 - cd.range.c1 + dst.c1 });
      }
    }
  });
  app.grid.select(dst.r1, dst.c1, { range: { r1: dst.r1, c1: dst.c1, r2: dst.r1 + outRows * tileR - 1, c2: dst.c1 + outCols * tileC - 1 }, keepActive: false });
  if (cd.cut) { clip = null; app.tab.marquee = null; }
}

function transposeRefs(f: string, srcR: number, srcC: number, r: number, c: number): string {
  // Relative references are re-anchored so the pasted formula points at the transposed position of its target.
  return mapRefs(f, (ref) => {
    const n = { ...ref };
    if (!ref.ar1) { const dr = ref.r1 - srcR, dc = ref.c1 - srcC; n.r1 = r + dc; if (!ref.ac1) n.c1 = c + dr; }
    else if (!ref.ac1) n.c1 = c + (ref.r1 - srcR);
    if (!ref.ar2) { const dr = ref.r2 - srcR, dc = ref.c2 - srcC; n.r2 = r + dc; if (!ref.ac2) n.c2 = c + dr; }
    if (n.r1 < 0 || n.c1 < 0 || n.r2 < 0 || n.c2 < 0) return { ...n, invalid: true };
    return n;
  });
}

/** Pastes external text (TSV/CSV) or an HTML table into the active cell. */
async function pasteExternal(app: App, text: string, html: string | null, o: PasteOpts): Promise<void> {
  const s = app.sheet;
  const dst = activeRange(app);
  let rows: string[][];
  let styles: (Cell | null)[][] | null = null;
  if (html && /<table/i.test(html) && (o.what || 'all') !== 'values') {
    try {
      const { importHtml } = await import('../io/html');
      const b = importHtml(html, [], { allTables: false });
      const sh = b.sheets[0];
      const u = sh.usedRange();
      if (u) {
        rows = [];
        styles = [];
        for (let r = u.r1; r <= u.r2; r++) {
          const row: string[] = [], srow: (Cell | null)[] = [];
          for (let c = u.c1; c <= u.c2; c++) { const cell = sh.get(r, c); row.push(cell?.v == null ? '' : String(cell.v)); srow.push(cell ? { ...cell, s: cell.s ? app.book.styleId(b.styles[cell.s]) : undefined } : null); }
          rows.push(row); styles.push(srow);
        }
      } else rows = parseTsv(text);
    } catch { rows = parseTsv(text); }
  } else {
    // Text with tabs → columns; a single line with commas stays one cell (use Text to Columns).
    rows = parseTsv(text.replace(/\r?\n$/, ''));
  }
  if (!rows.length) return;
  if (o.transpose) {
    const t: string[][] = [];
    rows.forEach((row, i) => row.forEach((v, j) => { (t[j] ||= [])[i] = v; }));
    rows = t; styles = null;
  }
  app.book.transact('Paste', () => {
    rows.forEach((row, i) => row.forEach((v, j) => {
      const r = dst.r1 + i, c = dst.c1 + j;
      if (r >= MAX_ROWS || c >= MAX_COLS) return;
      if (o.skipBlanks && v === '') return;
      setInput(app, s, r, c, v);
      const st = styles?.[i]?.[j]?.s;
      if (st && (o.what || 'all') !== 'values') s.patch(r, c, { s: app.book.derivedStyle(s.get(r, c)?.s ?? 0, app.book.styles[st]) });
    }));
  });
  const h = rows.length, w = Math.max(...rows.map((r) => r.length));
  app.grid.select(dst.r1, dst.c1, { range: { r1: dst.r1, c1: dst.c1, r2: dst.r1 + h - 1, c2: dst.c1 + w - 1 } });
  if (rows.length > 2 && rows.every((r) => r.length === 1) && rows.some((r) => /[,;|]/.test(r[0]))) toast('Pasted into one column.', { action: 'Text to Columns', onAction: () => app.exec('data.textToColumns') });
}

async function doPaste(app: App, o: PasteOpts, ev?: ClipboardEvent): Promise<void> {
  if ((app as any).objectClip && !ev?.clipboardData?.getData('text/plain')) { app.exec('object.paste', (app as any).objectClip); return; }
  let text: string | null = null, html: string | null = null;
  if (ev?.clipboardData) {
    text = ev.clipboardData.getData('text/plain');
    html = ev.clipboardData.getData('text/html') || null;
    // Pasted images become pictures
    const file = [...(ev.clipboardData.files || [])].find((f) => f.type.startsWith('image/'));
    if (file && !text) { app.exec('insert.pictureFile', file); return; }
  } else {
    try {
      if ((navigator.clipboard as any)?.read) {
        // Some browsers wait for a permission prompt that never appears: give up after a moment.
        const items = await Promise.race([navigator.clipboard.read(), new Promise<never>((_, rej) => setTimeout(() => rej(new Error('timeout')), 800))]);
        for (const it of items) {
          if (it.types.includes('text/html')) html = await (await it.getType('text/html')).text();
          if (it.types.includes('text/plain')) text = await (await it.getType('text/plain')).text();
          const img = it.types.find((t) => t.startsWith('image/'));
          if (img && !text) { app.exec('insert.pictureFile', new File([await it.getType(img)], 'pasted.png', { type: img })); return; }
        }
      }
    } catch { /* fall back */ }
    if (text == null) text = await Promise.race([readClipboardText(), new Promise<null>((res) => setTimeout(() => res(null), 800))]);
  }
  // Our own copy: use the internal clipboard (formulas, formats, comments).
  if (clip && (text == null || normalizeNl(text) === normalizeNl(clip.text))) { pasteInternal(app, clip, o); return; }
  if (text == null && clip) { pasteInternal(app, clip, o); return; }
  if (!text) { toast('The clipboard is empty, or the browser did not allow reading it. Use Ctrl+V.'); return; }
  await pasteExternal(app, text, html, o);
}

const normalizeNl = (s: string) => s.replace(/\r\n/g, '\n').replace(/\n+$/, '');

def([
  { id: 'edit.cut', label: 'Cut', icon: 'content_cut', tab: 'Home', group: 'Clipboard', keys: ['Mod-x', 'Shift-Delete'], popular: true, run: (app, arg) => copyOrCut(app, true, arg?.event) },
  { id: 'edit.copy', label: 'Copy', icon: 'content_copy', tab: 'Home', group: 'Clipboard', keys: ['Mod-c', 'Mod-Insert'], popular: true, run: (app, arg) => copyOrCut(app, false, arg?.event) },
  {
    id: 'edit.copyAsPicture', label: 'Copy as Picture', icon: 'photo_camera', tab: 'Home', group: 'Clipboard',
    run: async (app) => {
      const g = app.grid;
      const rc = g.rangeRect(activeRange(app));
      const c = document.createElement('canvas');
      c.width = Math.max(1, Math.round(rc.w * g.dpr)); c.height = Math.max(1, Math.round(rc.h * g.dpr));
      c.getContext('2d')!.drawImage(g.canvas, Math.round(rc.x * g.dpr), Math.round(rc.y * g.dpr), c.width, c.height, 0, 0, c.width, c.height);
      c.toBlob(async (b) => { if (!b) return; try { await navigator.clipboard.write([new ClipboardItem({ 'image/png': b })]); toast('Picture copied'); } catch { toast('This browser does not allow copying pictures.'); } });
    },
  },
  {
    id: 'edit.paste', label: 'Paste', icon: 'content_paste', tab: 'Home', group: 'Clipboard', keys: ['Mod-v', 'Shift-Insert'], kind: 'split', popular: true, edits: true,
    run: (app, arg) => (typeof arg === 'string' ? app.exec(arg) : doPaste(app, {}, arg?.event)),
    menu: () => [
      { heading: 'Paste' },
      { id: 'edit.paste', label: 'Paste', icon: 'content_paste' }, { id: 'edit.pasteFormulas', label: 'Formulas', icon: 'function' }, { id: 'edit.pasteFormulasNumFmt', label: 'Formulas & Number Formatting', icon: 'functions' },
      { id: 'edit.pasteKeepWidths', label: 'Keep Source Column Widths', icon: 'width' }, { id: 'edit.pasteNoBorders', label: 'No Borders', icon: 'border_clear' }, { id: 'edit.pasteTranspose', label: 'Transpose', icon: 'swap_horiz' },
      { heading: 'Paste Values' },
      { id: 'edit.pasteValues', label: 'Values', icon: '123' }, { id: 'edit.pasteValuesNumFmt', label: 'Values & Number Formatting', icon: 'pin' },
      { heading: 'Other Paste Options' },
      { id: 'edit.pasteFormats', label: 'Formatting', icon: 'format_paint' }, { id: 'edit.pasteLink', label: 'Paste Link', icon: 'link' }, { id: 'edit.copyAsPicture', label: 'Copy as Picture', icon: 'photo_camera' },
      { separator: true }, { id: 'edit.pasteSpecial', label: 'Paste Special…', icon: 'content_paste_go' },
    ],
  },
  { id: 'edit.pasteValues', label: 'Paste Values', icon: '123', tab: 'Home', group: 'Clipboard', keys: ['Mod-Shift-v'], popular: true, edits: true, run: (app) => doPaste(app, { what: 'values' }) },
  { id: 'edit.pasteValuesNumFmt', label: 'Paste Values & Number Formatting', icon: 'pin', tab: 'Home', group: 'Clipboard', edits: true, run: (app) => doPaste(app, { what: 'valuesAndNumFmt' }) },
  { id: 'edit.pasteFormulas', label: 'Paste Formulas', icon: 'function', tab: 'Home', group: 'Clipboard', edits: true, run: (app) => doPaste(app, { what: 'formulas' }) },
  { id: 'edit.pasteFormulasNumFmt', label: 'Paste Formulas & Number Formatting', icon: 'functions', tab: 'Home', group: 'Clipboard', edits: true, run: (app) => doPaste(app, { what: 'formulasAndNumFmt' }) },
  { id: 'edit.pasteFormats', label: 'Paste Formatting', icon: 'format_paint', tab: 'Home', group: 'Clipboard', edits: true, run: (app) => doPaste(app, { what: 'formats' }) },
  { id: 'edit.pasteTranspose', label: 'Paste Transposed', icon: 'swap_horiz', tab: 'Home', group: 'Clipboard', edits: true, run: (app) => doPaste(app, { transpose: true }) },
  { id: 'edit.pasteLink', label: 'Paste Link', icon: 'link', tab: 'Home', group: 'Clipboard', edits: true, run: (app) => doPaste(app, { link: true }) },
  { id: 'edit.pasteNoBorders', label: 'Paste without Borders', icon: 'border_clear', tab: 'Home', group: 'Clipboard', edits: true, run: (app) => doPaste(app, { what: 'noBorders' }) },
  { id: 'edit.pasteColumnWidths', label: 'Paste Column Widths', icon: 'width', tab: 'Home', group: 'Clipboard', edits: true, run: (app) => doPaste(app, { what: 'colWidths' }) },
  { id: 'edit.pasteKeepWidths', label: 'Paste Keeping Source Column Widths', icon: 'width', tab: 'Home', group: 'Clipboard', edits: true, run: (app) => doPaste(app, { keepWidths: true }) },
  { id: 'edit.pasteComments', label: 'Paste Comments and Notes', icon: 'comment', tab: 'Home', group: 'Clipboard', edits: true, run: (app) => doPaste(app, { what: 'comments' }) },
  { id: 'edit.pasteValidation', label: 'Paste Validation', icon: 'rule', tab: 'Home', group: 'Clipboard', edits: true, run: (app) => doPaste(app, { what: 'validation' }) },
  { id: 'edit.pasteText', label: 'Paste Text', icon: 'content_paste', tab: 'Home', group: 'Clipboard', internal: true, edits: true, run: (app, text: string) => pasteExternal(app, String(text || ''), null, {}) },
  {
    id: 'edit.pasteSpecial', label: 'Paste Special…', icon: 'content_paste_go', tab: 'Home', group: 'Clipboard', keys: ['Mod-Alt-v'], popular: true, edits: true,
    run: async (app) => {
      const r = await openDialog({ title: 'Paste Special', wide: true, fields: [
        { key: 'what', label: 'Paste', type: 'radio', value: 'all', options: [{ value: 'all', label: 'All' }, { value: 'formulas', label: 'Formulas' }, { value: 'values', label: 'Values' }, { value: 'formats', label: 'Formats' }, { value: 'comments', label: 'Comments and Notes' }, { value: 'validation', label: 'Validation' }, { value: 'allExceptBorders', label: 'All except borders' }, { value: 'colWidths', label: 'Column widths' }, { value: 'formulasAndNumFmt', label: 'Formulas and number formats' }, { value: 'valuesAndNumFmt', label: 'Values and number formats' }], half: true },
        { key: 'op', label: 'Operation', type: 'radio', value: 'none', options: [{ value: 'none', label: 'None' }, { value: 'add', label: 'Add' }, { value: 'subtract', label: 'Subtract' }, { value: 'multiply', label: 'Multiply' }, { value: 'divide', label: 'Divide' }], half: true },
        { key: 'skipBlanks', label: 'Skip blanks', type: 'checkbox', value: false, half: true },
        { key: 'transpose', label: 'Transpose', type: 'checkbox', value: false, half: true },
      ], buttons: [{ label: 'Paste Link', onClick: (_v, ctl) => { ctl.close(); void doPaste(app, { link: true }); } }] });
      if (!r) return;
      await doPaste(app, { what: r.what, op: r.op, skipBlanks: r.skipBlanks, transpose: r.transpose });
    },
  },
  {
    id: 'edit.painter', label: 'Format Painter', icon: 'format_paint', tab: 'Home', group: 'Clipboard', kind: 'toggle', popular: true, edits: true,
    isActive: (app) => !!(app as any).painter,
    run: (app, arg) => {
      const a = app as any;
      if (a.painter && arg !== 'apply') { a.painter = null; app.ui.status(''); return; }
      if (arg === 'apply' && a.painter) {
        const src = a.painter as { range: Range; sheet: Sheet; sticky: boolean };
        const dst = activeRange(app);
        const s = app.sheet;
        app.book.transact('Format Painter', () => {
          const h = src.range.r2 - src.range.r1 + 1, w = src.range.c2 - src.range.c1 + 1;
          for (let r = dst.r1; r <= Math.max(dst.r2, dst.r1 + h - 1); r++) for (let c = dst.c1; c <= Math.max(dst.c2, dst.c1 + w - 1); c++) {
            const sr = src.range.r1 + ((r - dst.r1) % h), sc = src.range.c1 + ((c - dst.c1) % w);
            const sid = src.sheet.styleIdOf(sr, sc);
            const sidLocal = src.sheet.book === app.book ? sid : app.book.styleId(src.sheet.book.styles[sid] || {});
            s.patch(r, c, { s: sidLocal || undefined });
          }
          // Copy conditional formats too
          for (const cf of src.sheet.cf) {
            if (!cf.ranges.some((x) => x.r1 <= src.range.r2 && x.r2 >= src.range.r1 && x.c1 <= src.range.c2 && x.c2 >= src.range.c1)) continue;
            s.book.recordSheet(s);
            s.cf.push({ ...JSON.parse(JSON.stringify(cf)), id: newId('cf'), ranges: [{ r1: dst.r1, c1: dst.c1, r2: Math.max(dst.r2, dst.r1 + h - 1), c2: Math.max(dst.c2, dst.c1 + w - 1) }] });
          }
        });
        if (!src.sticky) { a.painter = null; app.ui.status(''); }
        return;
      }
      a.painter = { range: { ...activeRange(app) }, sheet: app.sheet, sticky: arg === 'sticky' };
      app.ui.status(arg === 'sticky' ? 'Format Painter (locked): select cells to paint; Esc to stop' : 'Format Painter: select the cells to format');
      // The next selection applies the format.
      const off = app.onRefresh(() => {
        if (!a.painter) { off(); return; }
        const cur = activeRange(app);
        const src = a.painter.range;
        if (cur.r1 !== src.r1 || cur.c1 !== src.c1 || cur.r2 !== src.r2 || cur.c2 !== src.c2 || app.sheet !== a.painter.sheet) {
          if ((app.grid as any).drag) return;
          app.exec('edit.painter', 'apply');
          if (!a.painter) off();
          else a.painter.range = a.painter.range;
        }
      });
    },
  },
]);

// --------------------------------------------------------------------------------------------
// Fill, clear, move

def([
  {
    id: 'edit.fillDrag', label: 'Fill (drag)', icon: 'south', tab: 'Home', group: 'Editing', internal: true, edits: true,
    run: (app, arg: { src: Range; target: Range; copy?: boolean }) => {
      app.book.transact('Fill', () => fillRange(app, app.sheet, arg.src, arg.target, !!arg.copy));
      app.grid.select(arg.target.r1, arg.target.c1, { range: arg.target, keepActive: true, noScroll: true });
      app.ui.status('Auto Fill: Ctrl while dragging copies cells instead of a series');
    },
  },
  {
    id: 'edit.fillDown', label: 'Fill Down', icon: 'south', tab: 'Home', group: 'Editing', keys: ['Mod-d'], popular: true, edits: true,
    run: (app, arg) => {
      const s = app.sheet;
      let rg = activeRange(app);
      if (arg?.auto) {
        // Double-click on the fill handle: fill down as far as the neighboring column has data.
        const col = rg.c1 > 0 && s.get(rg.r2 + 1, rg.c1 - 1)?.v != null ? rg.c1 - 1 : rg.c2 + 1;
        let r = rg.r2;
        while (r + 1 < MAX_ROWS && (s.get(r + 1, col)?.v != null || s.get(r + 1, col)?.f != null)) r++;
        if (r === rg.r2) return;
        app.exec('edit.fillDrag', { src: rg, target: { ...rg, r2: r } });
        return;
      }
      if (rg.r1 === rg.r2) { if (rg.r1 === 0) return; rg = { ...rg, r1: rg.r1 - 1 }; }
      // A whole column is filled down to the end of the data, not to the last row of the sheet.
      if (rg.r2 >= MAX_ROWS - 1) rg = { ...rg, r2: Math.max(rg.r1 + 1, clipToUsed(s, rg).r2) };
      app.book.transact('Fill Down', () => fillRange(app, s, { ...rg, r2: rg.r1 }, rg, true));
    },
  },
  {
    id: 'edit.fillRight', label: 'Fill Right', icon: 'east', tab: 'Home', group: 'Editing', keys: ['Mod-r'], edits: true,
    run: (app) => {
      let rg = activeRange(app);
      if (rg.c1 === rg.c2) { if (rg.c1 === 0) return; rg = { ...rg, c1: rg.c1 - 1 }; }
      app.book.transact('Fill Right', () => fillRange(app, app.sheet, { ...rg, c2: rg.c1 }, rg, true));
    },
  },
  { id: 'edit.fillUp', label: 'Fill Up', icon: 'north', tab: 'Home', group: 'Editing', edits: true, run: (app) => { const rg = activeRange(app); if (rg.r1 === rg.r2) return; app.book.transact('Fill Up', () => fillRange(app, app.sheet, { ...rg, r1: rg.r2 }, rg, true)); } },
  { id: 'edit.fillLeft', label: 'Fill Left', icon: 'west', tab: 'Home', group: 'Editing', edits: true, run: (app) => { const rg = activeRange(app); if (rg.c1 === rg.c2) return; app.book.transact('Fill Left', () => fillRange(app, app.sheet, { ...rg, c1: rg.c2 }, rg, true)); } },
  {
    id: 'edit.fillSeries', label: 'Series…', icon: 'trending_up', tab: 'Home', group: 'Editing', edits: true,
    run: async (app) => {
      const rg = activeRange(app);
      const s = app.sheet;
      const first = s.value(rg.r1, rg.c1);
      const r = await openDialog({ title: 'Series', fields: [
        { key: 'dir', label: 'Series in', type: 'radio', options: [{ value: 'cols', label: 'Columns' }, { value: 'rows', label: 'Rows' }], value: rg.r2 - rg.r1 >= rg.c2 - rg.c1 ? 'cols' : 'rows', half: true },
        { key: 'type', label: 'Type', type: 'radio', options: [{ value: 'linear', label: 'Linear' }, { value: 'growth', label: 'Growth' }, { value: 'date', label: 'Date' }, { value: 'auto', label: 'AutoFill' }], value: 'linear', half: true },
        { key: 'unit', label: 'Date unit', type: 'select', options: [{ value: 'day', label: 'Day' }, { value: 'weekday', label: 'Weekday' }, { value: 'month', label: 'Month' }, { value: 'year', label: 'Year' }], value: 'day', visibleIf: (v) => v.type === 'date' },
        { key: 'trend', label: 'Trend (best fit from the selected values)', type: 'checkbox', value: false },
        { key: 'step', label: 'Step value', type: 'number', value: 1, half: true },
        { key: 'stop', label: 'Stop value', type: 'text', value: '', half: true },
      ] });
      if (!r) return;
      const start = typeof first === 'number' ? first : 0;
      const step = Number(r.step) || 0;
      const stop = r.stop === '' ? null : Number(r.stop);
      const cols = r.dir === 'cols';
      const lines = cols ? [rg.c1, rg.c2] : [rg.r1, rg.r2];
      const len = cols ? rg.r2 - rg.r1 + 1 : rg.c2 - rg.c1 + 1;
      if (r.type === 'auto' || r.trend) {
        app.book.transact('Series', () => {
          for (let line = lines[0]; line <= lines[1]; line++) {
            let n = 0;
            for (let i = 0; i < len; i++) if ((cols ? s.value(rg.r1 + i, line) : s.value(line, rg.c1 + i)) != null) n = i + 1; else break;
            n = Math.max(1, Math.min(n, r.trend ? n : 2));
            const src = cols ? { r1: rg.r1, c1: line, r2: rg.r1 + n - 1, c2: line } : { r1: line, c1: rg.c1, r2: line, c2: rg.c1 + n - 1 };
            fillRange(app, s, src, cols ? { ...src, r2: rg.r2 } : { ...src, c2: rg.c2 });
          }
        });
        return;
      }
      app.book.transact('Series', () => {
        for (let line = lines[0]; line <= lines[1]; line++) {
          const base = cols ? s.value(rg.r1, line) : s.value(line, rg.c1);
          let v = typeof base === 'number' ? base : start;
          const first = v;
          for (let i = 1; i < len; i++) {
            if (r.type === 'linear') v += step;
            else if (r.type === 'growth') v *= step;
            // Months and years are counted from the first date each time, so a shortened month end does not carry on.
            else v = r.unit === 'month' || r.unit === 'year' ? dateStep(first, step * i, r.unit) : dateStep(v, step, r.unit);
            if (stop != null && (step >= 0 ? v > stop : v < stop)) break;
            const rr = cols ? rg.r1 + i : line, cc = cols ? line : rg.c1 + i;
            const cur = s.get(rr, cc);
            s.set(rr, cc, { ...(cur?.s ? { s: cur.s } : s.get(rg.r1, rg.c1)?.s ? { s: s.get(rg.r1, rg.c1)!.s } : {}), v: +v.toPrecision(15) });
          }
        }
      });
    },
  },
  {
    id: 'edit.fillJustify', label: 'Justify', icon: 'format_align_justify', tab: 'Home', group: 'Editing', edits: true, tip: 'Reflow the text of the selected column cells to fit the selection width',
    run: (app) => {
      const s = app.sheet;
      const rg = activeRange(app);
      const words: string[] = [];
      for (let r = rg.r1; r <= rg.r2; r++) { const v = s.value(r, rg.c1); if (typeof v === 'string') words.push(...v.split(/\s+/).filter(Boolean)); }
      let width = 0;
      for (let c = rg.c1; c <= rg.c2; c++) width += s.colWidth(c);
      const maxChars = Math.max(4, Math.floor(width / 7));
      const lines: string[] = [];
      let line = '';
      for (const w of words) { if ((line + ' ' + w).trim().length > maxChars && line) { lines.push(line); line = w; } else line = (line + ' ' + w).trim(); }
      if (line) lines.push(line);
      app.book.transact('Justify', () => {
        for (let r = rg.r1; r <= Math.max(rg.r2, rg.r1 + lines.length - 1); r++) s.patch(r, rg.c1, { v: lines[r - rg.r1] ?? undefined });
      });
    },
  },
  {
    id: 'edit.fill', label: 'Fill', icon: 'arrow_downward', tab: 'Home', group: 'Editing', kind: 'menu', edits: true,
    run: () => undefined,
    menu: () => [{ id: 'edit.fillDown', label: 'Down', icon: 'south' }, { id: 'edit.fillRight', label: 'Right', icon: 'east' }, { id: 'edit.fillUp', label: 'Up', icon: 'north' }, { id: 'edit.fillLeft', label: 'Left', icon: 'west' }, { id: 'sheet.fillAcross', label: 'Across Worksheets…', icon: 'layers' }, { id: 'edit.fillSeries', label: 'Series…', icon: 'trending_up' }, { id: 'edit.fillJustify', label: 'Justify', icon: 'format_align_justify' }, { id: 'data.flashFill', label: 'Flash Fill', icon: 'bolt' }],
  },
  {
    id: 'edit.moveRange', label: 'Move Cells', icon: 'open_with', tab: 'Home', group: 'Editing', internal: true, edits: true,
    run: async (app, arg: { src: Range; dr: number; dc: number; copy?: boolean }) => {
      const s = app.sheet;
      const dst = { r1: arg.src.r1 + arg.dr, c1: arg.src.c1 + arg.dc, r2: arg.src.r2 + arg.dr, c2: arg.src.c2 + arg.dc };
      if (!arg.copy) {
        let occupied = false;
        s.forEachIn(dst, (r, c, cell) => { if ((cell.v != null || cell.f != null) && !(r >= arg.src.r1 && r <= arg.src.r2 && c >= arg.src.c1 && c <= arg.src.c2)) occupied = true; });
        if (occupied) { const c = await choiceDialog('Replace Contents?', "There's already data here. Do you want to replace it?", ['OK']); if (c !== 0) return; }
        app.book.transact('Move', () => moveBlock(s, arg.src, arg.dr, arg.dc));
      } else {
        app.book.transact('Copy', () => {
          const snap = snapshotRange(s, arg.src);
          for (const [k, cell] of snap.cells) {
            const i = keyRow(k), j = keyCol(k);
            const n: Cell = { ...cell };
            delete n.sp;
            if (n.f != null) n.f = shiftFormula(n.f, arg.dr, arg.dc);
            s.set(arg.src.r1 + i + arg.dr, arg.src.c1 + j + arg.dc, n);
          }
        });
      }
      app.grid.select(dst.r1, dst.c1, { range: dst });
    },
  },
]);

function dateStep(v: number, step: number, unit: string): number {
  const d = new Date(Date.UTC(1899, 11, 30) + v * 86400000);
  if (unit === 'day') return v + step;
  if (unit === 'weekday') { let n = v; let k = Math.abs(step); while (k > 0) { n += Math.sign(step) || 1; const dw = new Date(Date.UTC(1899, 11, 30) + n * 86400000).getUTCDay(); if (dw !== 0 && dw !== 6) k--; } return n; }
  // Months and years keep the day of the month, shortened to the month's last day where it has fewer (31 Jan + 1 month
  // is the end of February, not early March).
  const months = d.getUTCFullYear() * 12 + d.getUTCMonth() + (unit === 'year' ? step * 12 : step);
  const y = Math.floor(months / 12), m = months - y * 12;
  const day = Math.min(d.getUTCDate(), new Date(Date.UTC(y, m + 1, 0)).getUTCDate());
  return Math.round((Date.UTC(y, m, day) - Date.UTC(1899, 11, 30)) / 86400000) + (v - Math.floor(v));
}

const clearCmd = (id: string, label: string, icon: string, what: ClearWhat, keys?: string[]) => ({
  id, label, icon, tab: 'Home', group: 'Editing', keys, edits: true,
  run: (app: App) => app.book.transact(label, () => { for (const rg of selRanges(app)) clearRange(app, app.sheet, rg, what); }),
});

def([
  clearCmd('edit.clearAll', 'Clear All', 'clear_all', 'all'),
  clearCmd('edit.clearFormats', 'Clear Formats', 'format_clear', 'formats'),
  clearCmd('edit.clearContents', 'Clear Contents', 'backspace', 'contents'),
  clearCmd('edit.clearComments', 'Clear Comments and Notes', 'comments_disabled', 'comments'),
  clearCmd('edit.clearLinks', 'Clear Hyperlinks', 'link_off', 'links'),
  clearCmd('edit.removeLinks', 'Remove Hyperlinks', 'link_off', 'linksAndFormats'),
  {
    id: 'edit.clear', label: 'Clear', icon: 'ink_eraser', tab: 'Home', group: 'Editing', kind: 'menu', edits: true, run: () => undefined,
    menu: () => [{ id: 'edit.clearAll', label: 'Clear All', icon: 'clear_all' }, { id: 'edit.clearFormats', label: 'Clear Formats', icon: 'format_clear' }, { id: 'edit.clearContents', label: 'Clear Contents', icon: 'backspace' }, { id: 'edit.clearComments', label: 'Clear Comments and Notes', icon: 'comments_disabled' }, { id: 'edit.clearLinks', label: 'Clear Hyperlinks', icon: 'link_off' }, { id: 'edit.removeLinks', label: 'Remove Hyperlinks', icon: 'link_off' }],
  },
]);

// --------------------------------------------------------------------------------------------
// AutoSum

function autoSum(app: App, fn: string): void {
  const s = app.sheet;
  const rg = activeRange(app);
  const isNum = (r: number, c: number) => typeof s.value(r, c) === 'number';
  if (rg.r1 === rg.r2 && rg.c1 === rg.c2) {
    const { r, c } = activeCell(app);
    // Prefer the numbers above, else to the left.
    let r1 = r - 1;
    while (r1 >= 0 && (s.value(r1, c) == null)) r1--;
    let ref = '';
    if (r1 >= 0 && isNum(r1, c)) { let top = r1; while (top - 1 >= 0 && (isNum(top - 1, c) || s.get(top - 1, c)?.f)) top--; ref = rangeName({ r1: top, c1: c, r2: r1, c2: c }); }
    else {
      let c1 = c - 1;
      while (c1 >= 0 && s.value(r, c1) == null) c1--;
      if (c1 >= 0 && isNum(r, c1)) { let left = c1; while (left - 1 >= 0 && isNum(r, left - 1)) left--; ref = rangeName({ r1: r, c1: left, r2: r, c2: c1 }); }
    }
    app.grid.startEdit({ mode: 'edit', text: `=${fn}(${ref})` });
    if (ref) { const inp = app.grid.editor.ta; inp.setSelectionRange(fn.length + 2, fn.length + 2 + ref.length); }
    return;
  }
  // A selected block: totals below each column (and to the right when a blank column is included).
  app.book.transact('AutoSum', () => {
    const lastRowEmpty = [...Array(rg.c2 - rg.c1 + 1)].every((_, j) => s.value(rg.r2, rg.c1 + j) == null);
    const targetRow = lastRowEmpty ? rg.r2 : rg.r2 + 1;
    const dataR2 = lastRowEmpty ? rg.r2 - 1 : rg.r2;
    for (let c = rg.c1; c <= rg.c2; c++) {
      if (![...Array(dataR2 - rg.r1 + 1)].some((_, i) => isNum(rg.r1 + i, c))) continue;
      const cur = s.get(targetRow, c);
      s.set(targetRow, c, { ...(cur?.s ? { s: cur.s } : {}), f: `${fn}(${rangeName({ r1: rg.r1, c1: c, r2: dataR2, c2: c })})` });
    }
  });
}

def([
  {
    id: 'edit.autoSum', label: 'AutoSum', icon: 'functions', tab: 'Home', group: 'Editing', keys: ['Alt-='], kind: 'split', popular: true, edits: true,
    run: (app, fn) => autoSum(app, typeof fn === 'string' ? fn : 'SUM'),
    menu: () => [{ id: 'edit.autoSum', arg: 'SUM', label: 'Sum' }, { id: 'edit.autoSum', arg: 'AVERAGE', label: 'Average' }, { id: 'edit.autoSum', arg: 'COUNT', label: 'Count Numbers' }, { id: 'edit.autoSum', arg: 'MAX', label: 'Max' }, { id: 'edit.autoSum', arg: 'MIN', label: 'Min' }, { separator: true }, { id: 'formulas.insertFunction', label: 'More Functions…' }],
  },
]);

// --------------------------------------------------------------------------------------------
// Find & Replace, Go To

interface FindOpts { what: string; replace?: string; within: 'sheet' | 'workbook'; by: 'rows' | 'cols'; lookIn: 'formulas' | 'values' | 'comments'; matchCase: boolean; whole: boolean; regex: boolean }

/** The search text as a regular expression source: wildcards * and ?, with ~* and ~? for the literal characters. */
function findPattern(o: FindOpts): string {
  return o.regex ? o.what : o.what.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/~\*/g, '\u0001').replace(/~\?/g, '\u0002').replace(/\*/g, '.*').replace(/\?/g, '.').replace(/\u0001/g, '\\*').replace(/\u0002/g, '\\?');
}

function* findIter(app: App, o: FindOpts): Generator<{ sheet: Sheet; r: number; c: number }> {
  const sheets = o.within === 'workbook' ? app.book.sheets : [app.sheet];
  const flags = o.matchCase ? '' : 'i';
  const pat = findPattern(o);
  const re = new RegExp(o.whole ? `^(?:${pat})$` : pat, flags);
  for (const sh of sheets) {
    const keys = [...sh.cells.keys()].sort((a, b) => (o.by === 'rows' ? a - b : keyCol(a) - keyCol(b) || keyRow(a) - keyRow(b)));
    if (o.lookIn === 'comments') {
      for (const [k, cm] of [...sh.comments].sort((a, b) => a[0] - b[0])) if (re.test(cm.text)) yield { sheet: sh, r: keyRow(k), c: keyCol(k) };
      continue;
    }
    for (const k of keys) {
      const cell = sh.cells.get(k)!;
      const r = keyRow(k), c = keyCol(k);
      const text = o.lookIn === 'formulas' && cell.f != null ? '=' + cell.f : cell.v == null ? '' : formatValue(cell.v, sh.styleOf(r, c).numFmt, { date1904: app.book.calc.date1904 }).text;
      if (text && re.test(text)) yield { sheet: sh, r, c };
    }
  }
}

async function findDialog(app: App, replace: boolean): Promise<void> {
  const last = (app as any).lastFind as FindOpts | undefined;
  const sel = activeRange(app);
  let results: { sheet: Sheet; r: number; c: number }[] = [];
  let idx = -1;
  const opts = (v: Record<string, any>): FindOpts => ({ what: v.what, replace: v.repl, within: v.within, by: v.by, lookIn: v.lookIn, matchCase: v.matchCase, whole: v.whole, regex: v.regex });
  const goTo = (i: number) => {
    const hit = results[i];
    if (!hit) return;
    const si = app.book.sheets.indexOf(hit.sheet);
    if (si !== app.book.activeSheet) app.exec('sheet.activate', si);
    app.grid.select(hit.r, hit.c);
  };
  await openDialog({
    title: replace ? 'Find and Replace' : 'Find', wide: true, ok: false, cancel: 'Close',
    fields: [
      { key: 'what', label: 'Find what', type: 'text', value: last?.what || (sel.r1 === sel.r2 && sel.c1 === sel.c2 ? '' : ''), autofocus: true },
      ...(replace ? [{ key: 'repl', label: 'Replace with', type: 'text' as const, value: last?.replace || '' }] : []),
      { key: 'within', label: 'Within', type: 'select', options: [{ value: 'sheet', label: 'Sheet' }, { value: 'workbook', label: 'Workbook' }], value: last?.within || 'sheet', half: true },
      { key: 'by', label: 'Search', type: 'select', options: [{ value: 'rows', label: 'By Rows' }, { value: 'cols', label: 'By Columns' }], value: last?.by || 'rows', half: true },
      { key: 'lookIn', label: 'Look in', type: 'select', options: [{ value: 'formulas', label: 'Formulas' }, { value: 'values', label: 'Values' }, { value: 'comments', label: 'Comments and notes' }], value: last?.lookIn || 'formulas', half: true },
      { key: 'matchCase', label: 'Match case', type: 'checkbox', value: !!last?.matchCase, half: true },
      { key: 'whole', label: 'Match entire cell contents', type: 'checkbox', value: !!last?.whole, half: true },
      { key: 'regex', label: 'Regular expression', type: 'checkbox', value: !!last?.regex, half: true },
      { key: 'res', type: 'info', html: '<div class="ql-find-results"></div>' },
    ],
    buttons: [
      { label: 'Find All', onClick: (v, ctl) => {
        if (!v.what) return;
        (app as any).lastFind = opts(v);
        results = [...findIter(app, opts(v))];
        idx = -1;
        const box = ctl.body.querySelector('.ql-find-results') as HTMLElement;
        box.innerHTML = '';
        const t = document.createElement('table'); t.className = 'ql-stats';
        results.slice(0, 500).forEach((rr, i) => { const tr = document.createElement('tr'); tr.innerHTML = `<td>${escHtml(rr.sheet.name)}</td><td>${cellName(rr.r, rr.c)}</td><td>${escHtml(String(rr.sheet.value(rr.r, rr.c) ?? ''))}</td>`; tr.style.cursor = 'pointer'; tr.onclick = () => { idx = i; goTo(i); }; t.appendChild(tr); });
        box.append(`${results.length} cell(s) found`, t);
      } },
      { label: 'Find Next', primary: !replace, onClick: (v) => {
        if (!v.what) return;
        (app as any).lastFind = opts(v);
        results = [...findIter(app, opts(v))];
        if (!results.length) { toast(`Cannot find "${v.what}"`); return; }
        const cur = activeCell(app);
        const curSheet = app.sheet;
        const after = results.findIndex((x) => x.sheet === curSheet && (v.by === 'rows' ? x.r > cur.r || (x.r === cur.r && x.c > cur.c) : x.c > cur.c || (x.c === cur.c && x.r > cur.r)));
        idx = after >= 0 ? after : 0;
        goTo(idx);
      } },
      ...(replace ? [
        { label: 'Replace', onClick: (v: Record<string, any>) => {
          const o = opts(v);
          const cur = activeCell(app);
          const s = app.sheet;
          const hit = [...findIter(app, { ...o, within: 'sheet' })].find((x) => x.r === cur.r && x.c === cur.c);
          if (hit) replaceIn(app, s, cur.r, cur.c, o);
          const next = [...findIter(app, o)];
          if (next.length) { results = next; const i = next.findIndex((x) => x.sheet !== s || x.r > cur.r || (x.r === cur.r && x.c > cur.c)); goTo(i >= 0 ? i : 0); }
        } },
        { label: 'Replace All', onClick: (v: Record<string, any>) => {
          const o = opts(v);
          const hits = [...findIter(app, o)];
          app.book.transact('Replace All', () => { for (const x of hits) replaceIn(app, x.sheet, x.r, x.c, o); });
          toast(`All done. We made ${hits.length} replacement(s).`);
        } },
      ] : []),
    ],
  });
}

function replaceIn(app: App, sheet: Sheet, r: number, c: number, o: FindOpts): void {
  const cell = sheet.get(r, c);
  if (!cell) return;
  // A formula found by its result is left alone: replacing in the result would turn it into a constant.
  if (cell.f != null && o.lookIn !== 'formulas') return;
  const re = new RegExp(o.whole ? `^(?:${findPattern(o)})$` : findPattern(o), o.matchCase ? 'g' : 'gi');
  const src = cell.f != null ? '=' + cell.f : cell.v == null ? '' : String(cell.v);
  // Outside regular-expression mode the replacement is literal text ($& and $1 mean nothing special).
  const out = o.regex ? src.replace(re, o.replace || '') : src.replace(re, () => o.replace || '');
  if (out === src) return;
  app.book.transact('Replace', () => setInput(app, sheet, r, c, out));
}

def([
  { id: 'edit.find', label: 'Find…', icon: 'search', tab: 'Home', group: 'Editing', keys: ['Mod-f', 'Shift-F5'], popular: true, noRecord: true, run: (app) => findDialog(app, false) },
  { id: 'edit.replace', label: 'Replace…', icon: 'find_replace', tab: 'Home', group: 'Editing', keys: ['Mod-h'], popular: true, run: (app) => findDialog(app, true) },
  {
    id: 'edit.findNext', label: 'Find Next', icon: 'keyboard_arrow_down', tab: 'Home', group: 'Editing', keys: ['Shift-F4'], noRecord: true,
    run: (app) => {
      const o = (app as any).lastFind as FindOpts | undefined;
      if (!o) return app.exec('edit.find');
      const res = [...findIter(app, o)];
      if (!res.length) { toast(`Cannot find "${o.what}"`); return; }
      const cur = activeCell(app);
      const i = res.findIndex((x) => x.sheet !== app.sheet || x.r > cur.r || (x.r === cur.r && x.c > cur.c));
      const hit = res[i >= 0 ? i : 0];
      const si = app.book.sheets.indexOf(hit.sheet);
      if (si !== app.book.activeSheet) app.exec('sheet.activate', si);
      app.grid.select(hit.r, hit.c);
    },
  },
  {
    id: 'edit.findSelect', label: 'Find & Select', icon: 'search', tab: 'Home', group: 'Editing', kind: 'menu', run: () => undefined,
    menu: () => [{ id: 'edit.find', label: 'Find…', icon: 'search' }, { id: 'edit.replace', label: 'Replace…', icon: 'find_replace' }, { id: 'nav.goto', label: 'Go To…', icon: 'move_down' }, { id: 'nav.gotoSpecial', label: 'Go To Special…', icon: 'filter_center_focus' }, { separator: true }, { id: 'nav.selectFormulas', label: 'Formulas', icon: 'function' }, { id: 'nav.selectComments', label: 'Notes', icon: 'sticky_note_2' }, { id: 'nav.selectCF', label: 'Conditional Formatting', icon: 'format_color_fill' }, { id: 'nav.selectConstants', label: 'Constants', icon: 'pin' }, { id: 'nav.selectValidation', label: 'Data Validation', icon: 'rule' }, { separator: true }, { id: 'nav.selectObjects', label: 'Select Objects', icon: 'touch_app' }, { id: 'view.selectionPane', label: 'Selection Pane…', icon: 'layers' }],
  },
]);

// --------------------------------------------------------------------------------------------
// Navigation and selection

function selectCells(app: App, pred: (cell: Cell | undefined, r: number, c: number) => boolean, within?: Range): void {
  const s = app.sheet;
  const area = within && (within.r1 !== within.r2 || within.c1 !== within.c2) ? within : s.usedRange() || { r1: 0, c1: 0, r2: 0, c2: 0 };
  const hits: Range[] = [];
  s.forEachIn(area, (r, c, cell) => { if (pred(cell, r, c)) hits.push({ r1: r, c1: c, r2: r, c2: c }); });
  if (pred === undefined) return;
  if (!hits.length) { toast('No cells were found.'); return; }
  // merge horizontally adjacent hits into row runs
  hits.sort((a, b) => a.r1 - b.r1 || a.c1 - b.c1);
  const runs: Range[] = [];
  for (const h of hits) { const last = runs[runs.length - 1]; if (last && last.r1 === h.r1 && last.c2 + 1 === h.c1) last.c2 = h.c1; else runs.push({ ...h }); }
  s.sel.ranges = runs.slice(0, 5000);
  s.sel.r = runs[0].r1; s.sel.c = runs[0].c1;
  app.grid.scrollIntoView(s.sel.r, s.sel.c);
  app.grid.invalidate();
  app.ui.status(`${hits.length} cell(s) selected`);
}

def([
  {
    id: 'nav.goto', label: 'Go To…', icon: 'move_down', tab: 'Home', group: 'Editing', keys: ['Mod-g', 'F5'], popular: true,
    run: async (app, target?: string) => {
      let ref = target;
      if (!ref) {
        const names = app.book.names.map((n) => n.name);
        const r = await openDialog({ title: 'Go To', fields: [{ key: 'list', label: 'Go to', type: 'list', options: names, rows: 6 }, { key: 'ref', label: 'Reference', type: 'text', value: '', autofocus: true }], onChange: (v, ctl, k) => { if (k === 'list') ctl.set('ref', Array.isArray(v.list) ? v.list[0] : v.list); }, buttons: [{ label: 'Special…', onClick: (_v, ctl) => { ctl.close(); app.exec('nav.gotoSpecial'); } }] });
        if (!r) return;
        ref = r.ref || (Array.isArray(r.list) ? r.list[0] : r.list);
      }
      if (!ref) return;
      const t = String(ref).trim();
      // Table name
      for (const [i, sh] of app.book.sheets.entries()) {
        const tb = sh.tables.find((x) => x.name.toLowerCase() === t.toLowerCase());
        if (tb) { app.exec('sheet.activate', i); app.grid.select(tb.range.r1, tb.range.c1, { range: tb.range }); return; }
      }
      const res = resolveRangeText(app.book, t, app.sheet);
      if (!res) {
        const { sheet: sn } = splitSheetRef(t);
        if (sn && !app.book.sheetByName(sn)) toast(`There is no sheet named "${sn}".`); else toast('Reference isn’t valid.');
        return;
      }
      const si = app.book.sheets.indexOf(res.sheet);
      if (si !== app.book.activeSheet) app.exec('sheet.activate', si);
      app.grid.select(res.range.r1, res.range.c1, { range: res.range });
    },
  },
  {
    id: 'nav.gotoSpecial', label: 'Go To Special…', icon: 'filter_center_focus', tab: 'Home', group: 'Editing',
    run: async (app) => {
      const r = await openDialog({ title: 'Go To Special', fields: [{ key: 'w', label: 'Select', type: 'radio', value: 'blanks', options: [
        { value: 'comments', label: 'Notes and comments' }, { value: 'constants', label: 'Constants' }, { value: 'formulas', label: 'Formulas' }, { value: 'blanks', label: 'Blanks' },
        { value: 'region', label: 'Current region' }, { value: 'array', label: 'Current array' }, { value: 'objects', label: 'Objects' }, { value: 'rowDiff', label: 'Row differences' }, { value: 'colDiff', label: 'Column differences' },
        { value: 'precedents', label: 'Precedents' }, { value: 'dependents', label: 'Dependents' }, { value: 'last', label: 'Last cell' }, { value: 'visible', label: 'Visible cells only' },
        { value: 'cf', label: 'Conditional formats' }, { value: 'dv', label: 'Data validation' }, { value: 'errors', label: 'Errors' }, { value: 'numbers', label: 'Numbers' }, { value: 'text', label: 'Text' }, { value: 'logicals', label: 'Logicals' },
      ] }] });
      if (!r) return;
      const s = app.sheet;
      const within = activeRange(app);
      const sel = within.r1 !== within.r2 || within.c1 !== within.c2 ? within : undefined;
      const { r: ar, c: ac } = activeCell(app);
      switch (r.w) {
        case 'comments': selectCells(app, (_c, rr, cc) => s.comments.has(cellKey(rr, cc)), sel); break;
        case 'constants': selectCells(app, (c) => !!c && c.f == null && c.sp == null && c.v != null, sel); break;
        case 'formulas': selectCells(app, (c) => !!c && c.f != null, sel); break;
        case 'errors': selectCells(app, (c) => !!c && c.v instanceof CellError, sel); break;
        case 'numbers': selectCells(app, (c) => !!c && typeof c.v === 'number', sel); break;
        case 'text': selectCells(app, (c) => !!c && typeof c.v === 'string', sel); break;
        case 'logicals': selectCells(app, (c) => !!c && typeof c.v === 'boolean', sel); break;
        case 'blanks': {
          const area0 = sel || s.usedRange();
          if (!area0) return;
          // Blanks are looked for within the data: beyond it every cell is blank.
          const area = clipToUsed(s, area0);
          const runs: Range[] = [];
          for (let rr = area.r1; rr <= area.r2; rr++) for (let cc = area.c1; cc <= area.c2; cc++) { const v = s.get(rr, cc); if (!v || (v.v == null && v.f == null)) { const last = runs[runs.length - 1]; if (last && last.r1 === rr && last.c2 + 1 === cc) last.c2 = cc; else runs.push({ r1: rr, c1: cc, r2: rr, c2: cc }); } }
          if (!runs.length) { toast('No cells were found.'); return; }
          if (runs.length > 5000) { runs.length = 5000; toast('Only the first 5,000 blank areas were selected.'); }
          s.sel.ranges = runs; s.sel.r = runs[0].r1; s.sel.c = runs[0].c1; app.grid.invalidate();
          break;
        }
        case 'region': app.grid.select(ar, ac, { range: currentRegion(s, ar, ac), keepActive: true }); break;
        case 'array': { const sp = app.tab.engine.spillRange(s, ar, ac) || s.get(ar, ac)?.arr; if (sp) app.grid.select(ar, ac, { range: sp, keepActive: true }); else toast('No array found.'); break; }
        case 'objects': app.exec('nav.selectObjects'); break;
        case 'last': { const u = s.usedRange(); if (u) app.grid.select(u.r2, u.c2); break; }
        case 'visible': { const rg = activeRange(app); const runs: Range[] = []; for (let rr = rg.r1; rr <= Math.min(rg.r2, (s.usedRange()?.r2 ?? rg.r1)); rr++) { if (s.hiddenRows.has(rr)) continue; const last = runs[runs.length - 1]; if (last && last.r2 + 1 === rr) last.r2 = rr; else runs.push({ r1: rr, c1: rg.c1, r2: rr, c2: rg.c2 }); } s.sel.ranges = runs.length ? runs : [rg]; app.grid.invalidate(); break; }
        case 'cf': selectCells(app, (_c, rr, cc) => s.cf.some((x) => x.ranges.some((g) => rr >= g.r1 && rr <= g.r2 && cc >= g.c1 && cc <= g.c2)), sel); break;
        case 'dv': selectCells(app, (_c, rr, cc) => s.validations.some((x) => x.ranges.some((g) => rr >= g.r1 && rr <= g.r2 && cc >= g.c1 && cc <= g.c2)), sel); break;
        case 'precedents': { const ps = app.tab.engine.precedentsOf(s, ar, ac).filter((p) => p.sheet === s); if (!ps.length) { toast('No precedents.'); return; } s.sel.ranges = ps.map((p) => ({ r1: p.r1, c1: p.c1, r2: Math.min(p.r2, (s.usedRange()?.r2 ?? p.r2)), c2: Math.min(p.c2, (s.usedRange()?.c2 ?? p.c2)) })); app.grid.invalidate(); break; }
        case 'dependents': { const ds = app.tab.engine.dependentsOfCell(s, ar, ac).filter((p) => p.sheet === s); if (!ds.length) { toast('No dependents.'); return; } s.sel.ranges = ds.map((p) => ({ r1: p.r, c1: p.c, r2: p.r, c2: p.c })); app.grid.invalidate(); break; }
        case 'rowDiff': case 'colDiff': {
          const rg = sel || currentRegion(s, ar, ac);
          selectCells(app, (_c, rr, cc) => {
            const ref = r.w === 'rowDiff' ? s.value(rr, ac) : s.value(ar, cc);
            return compare(s.value(rr, cc), ref) !== 0;
          }, rg);
          break;
        }
      }
      app.refresh();
    },
  },
  { id: 'nav.selectFormulas', label: 'Select Formulas', icon: 'function', tab: 'Home', group: 'Editing', run: (app) => selectCells(app, (c) => !!c && c.f != null) },
  { id: 'nav.selectComments', label: 'Select Notes', icon: 'sticky_note_2', tab: 'Home', group: 'Editing', run: (app) => selectCells(app, (_c, r, c) => app.sheet.comments.has(cellKey(r, c))) },
  { id: 'nav.selectConstants', label: 'Select Constants', icon: 'pin', tab: 'Home', group: 'Editing', run: (app) => selectCells(app, (c) => !!c && c.f == null && c.sp == null && c.v != null) },
  { id: 'nav.selectCF', label: 'Select Conditional Formatting', icon: 'format_color_fill', tab: 'Home', group: 'Editing', run: (app) => selectCells(app, (_c, r, c) => app.sheet.cf.some((x) => x.ranges.some((g) => r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2))) },
  { id: 'nav.selectValidation', label: 'Select Data Validation', icon: 'rule', tab: 'Home', group: 'Editing', run: (app) => selectCells(app, (_c, r, c) => app.sheet.validations.some((x) => x.ranges.some((g) => r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2))) },
  { id: 'nav.selectObjects', label: 'Select Objects', icon: 'touch_app', tab: 'Home', group: 'Editing', run: (app) => { const s = app.sheet; const first = s.charts[0]?.id || s.drawings[0]?.id; if (first) app.grid.objects.select(first); else toast('There are no objects on this sheet.'); } },
  { id: 'nav.select', label: 'Select Range', icon: 'select', tab: 'Navigation', internal: true, run: (app, arg: { ranges: Range[]; r: number; c: number }) => { if (!arg?.ranges) return; app.sheet.sel = { r: arg.r, c: arg.c, ranges: arg.ranges.map(normRange) }; app.grid.scrollIntoView(arg.r, arg.c); app.grid.invalidate(); } },
  { id: 'nav.selectAll', label: 'Select All', icon: 'select_all', tab: 'Home', group: 'Editing', keys: ['Mod-a', 'Mod-Shift-Space'], noRecord: true, run: (app) => {
    const s = app.sheet; const { r, c } = activeCell(app);
    const region = currentRegion(s, r, c);
    const cur = activeRange(app);
    // First press selects the current region (when inside data), second selects the whole sheet.
    if ((region.r1 !== region.r2 || region.c1 !== region.c2) && !(cur.r1 === region.r1 && cur.c1 === region.c1 && cur.r2 === region.r2 && cur.c2 === region.c2)) app.grid.select(r, c, { range: region, keepActive: true });
    else app.grid.select(r, c, { range: { r1: 0, c1: 0, r2: MAX_ROWS - 1, c2: MAX_COLS - 1 }, keepActive: true, noScroll: true });
  } },
  { id: 'nav.currentRegion', label: 'Select Current Region', icon: 'crop_free', tab: 'Home', group: 'Editing', keys: ['Mod-Shift-8', 'Mod-*'], noRecord: true, run: (app) => { const { r, c } = activeCell(app); app.grid.select(r, c, { range: currentRegion(app.sheet, r, c), keepActive: true }); } },
  { id: 'nav.selectColumn', label: 'Select Entire Column', icon: 'view_column', tab: 'Home', group: 'Editing', keys: ['Mod-Space'], noRecord: true, run: (app) => { const rg = activeRange(app); app.grid.selectCols(rg.c1, rg.c2, false); } },
  { id: 'nav.selectRow', label: 'Select Entire Row', icon: 'table_rows', tab: 'Home', group: 'Editing', keys: ['Shift-Space'], noRecord: true, run: (app) => { const rg = activeRange(app); app.grid.selectRows(rg.r1, rg.r2, false); } },
  { id: 'nav.selectVisible', label: 'Select Visible Cells', icon: 'visibility', tab: 'Home', group: 'Editing', keys: ['Alt-;'], noRecord: true, run: (app) => { const s = app.sheet; const rg = activeRange(app); const runs: Range[] = []; for (let rr = rg.r1; rr <= Math.min(rg.r2, (s.usedRange()?.r2 ?? rg.r1)); rr++) { if (s.hiddenRows.has(rr)) continue; const last = runs[runs.length - 1]; if (last && last.r2 + 1 === rr) last.r2 = rr; else runs.push({ r1: rr, c1: rg.c1, r2: rr, c2: rg.c2 }); } s.sel.ranges = runs.length ? runs : [rg]; app.grid.invalidate(); } },
  { id: 'nav.edge', label: 'Jump to Data Edge', icon: 'last_page', tab: 'Navigation', internal: true, noRecord: true, run: (app, arg: { dir: 'up' | 'down' | 'left' | 'right' }) => { const d = { up: [-1, 0], down: [1, 0], left: [0, -1], right: [0, 1] }[arg.dir]; app.grid.moveActive(d[0], d[1], true); } },
  { id: 'nav.home', label: 'Go to Beginning of Sheet', icon: 'first_page', tab: 'Navigation', keys: ['Mod-Home'], noRecord: true, run: (app) => app.grid.select(app.sheet.freeze.rows, app.sheet.freeze.cols) },
  { id: 'nav.end', label: 'Go to Last Used Cell', icon: 'last_page', tab: 'Navigation', keys: ['Mod-End'], noRecord: true, run: (app) => { const u = app.sheet.usedRange(false); if (u) app.grid.select(u.r2, u.c2); } },
  { id: 'nav.activeCell', label: 'Show Active Cell', icon: 'my_location', tab: 'Navigation', keys: ['Mod-Backspace'], noRecord: true, run: (app) => { const { r, c } = activeCell(app); app.grid.scrollIntoView(r, c); } },
  { id: 'nav.followLink', label: 'Follow Hyperlink', icon: 'open_in_new', tab: 'Navigation', noRecord: true, run: (app) => { const { r, c } = activeCell(app); const l = app.sheet.get(r, c)?.link; if (l) app.followLink(l); else toast('The active cell has no hyperlink.'); } },
]);

// --------------------------------------------------------------------------------------------
// Quick entry shortcuts

def([
  { id: 'edit.insertDate', label: 'Insert Current Date', icon: 'today', tab: 'Home', group: 'Editing', keys: ['Mod-;'], edits: true, run: (app) => { const { r, c } = activeCell(app); const d = new Date(); app.book.transact('Insert Date', () => setInput(app, app.sheet, r, c, `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`)); } },
  { id: 'edit.insertTime', label: 'Insert Current Time', icon: 'schedule', tab: 'Home', group: 'Editing', keys: ['Mod-Shift-;'], edits: true, run: (app) => { const { r, c } = activeCell(app); const d = new Date(); app.book.transact('Insert Time', () => setInput(app, app.sheet, r, c, `${d.getHours()}:${String(d.getMinutes()).padStart(2, '0')}`)); } },
  { id: 'edit.copyAbove', label: 'Copy Formula from Cell Above', icon: 'vertical_align_bottom', tab: 'Home', group: 'Editing', keys: ["Mod-'"], edits: true, run: (app) => { const { r, c } = activeCell(app); if (!r) return; const above = app.sheet.get(r - 1, c); app.grid.startEdit({ mode: 'edit', text: above?.f != null ? '=' + above.f : String(above?.v ?? '') }); } },
  { id: 'edit.copyValueAbove', label: 'Copy Value from Cell Above', icon: 'vertical_align_bottom', tab: 'Home', group: 'Editing', keys: ['Mod-Shift-"'], edits: true, run: (app) => { const { r, c } = activeCell(app); if (!r) return; const v = app.sheet.value(r - 1, c); app.grid.startEdit({ mode: 'edit', text: v == null ? '' : String(v) }); } },
  { id: 'edit.editCell', label: 'Edit Active Cell', icon: 'edit', tab: 'Home', group: 'Editing', noRecord: true, run: (app) => app.grid.startEdit({ mode: 'edit' }) },
  { id: 'edit.pickFromList', label: 'Pick from Drop-down List', icon: 'list', tab: 'Home', group: 'Editing', keys: ['Alt-ArrowDown'], run: (app) => app.exec('data.pickFromList') },
]);

void criteria; void showMenu; void isWholeCols; void isWholeRows; void parseRange; void ERR;
