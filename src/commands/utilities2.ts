/**
 * Utilities tab, part 2: Range, Fill, Columns & Rows, Objects & Comments (ASAP Utilities-style tools,
 * re-implemented; see utilities.ts).
 */
import type { App } from '../app';
import { Sheet } from '../model/workbook';
import type { Cell, CellComment, Scalar } from '../model/types';
import { CellError } from '../model/types';
import { cellKey, keyRow, keyCol, colName, cellName, rangeName, parseRange, MAX_ROWS, MAX_COLS, type Range } from '../model/address';
import { shiftFormula } from '../calc/parser';
import { selRanges, activeRange, activeCell, clipToUsed, isWholeCols, isWholeRows, rangeToTsv } from './ops';
import { defTools, U, selCells, selBlock, singleCell, hasContent, selectHits, sheetProtected, emptySheet, numOf, textOf, plural, report, pickSheets, sheetsByNames, normColor, shownText, argOrAsk, columnChoices, headerGuess, naturalCompare, writeOut, checkWritable, canChange, blockOf, movedCell, hasMergeIn, reorderRows, reorderCols, rowIsBlank, colIsBlank, deleteRowSpans, deleteColSpans, insertRows, insertCols, mergeSpans, PX_PER_CM, PX_PER_IN, PX_PER_PT, copyText, rangeFromText, valueRows, writeBlock } from './utilShared';
import { menuOf } from './utilities';
import { openDialog, toast, confirmDialog, promptDialog } from '../ui/dialog';
import { setComment } from './review';
import { parseTextDate } from './tools';
import { jsDateToSerial } from '../model/dates';

const G_RANGE = 'Range';
const G_FILL = 'Fill';
const G_COLROW = 'Columns & Rows';
const G_OBJECTS = 'Objects & Comments';

// ============================================================================================
// Range

export interface FindReplaceOptions { find: string; replace: string; matchCase?: boolean; wholeCell?: boolean; regex?: boolean; inValues?: boolean; inFormulas?: boolean; inComments?: boolean; inLinks?: boolean; inSheetNames?: boolean; scope?: 'selection' | 'sheet' | 'workbook' }

export function replaceText(s: string, o: FindReplaceOptions): string | null {
  if (!o.find) return null;
  if (o.wholeCell) {
    const same = o.matchCase ? s === o.find : s.toLowerCase() === o.find.toLowerCase();
    if (o.regex) { try { if (new RegExp(`^(?:${o.find})$`, o.matchCase ? 'u' : 'iu').test(s)) return s.replace(new RegExp(`^(?:${o.find})$`, o.matchCase ? 'u' : 'iu'), o.replace); } catch { return null; } return null; }
    return same ? o.replace : null;
  }
  let re: RegExp;
  try { re = o.regex ? new RegExp(o.find, o.matchCase ? 'gu' : 'giu') : new RegExp(o.find.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), o.matchCase ? 'g' : 'gi'); } catch { return null; }
  if (!re.test(s)) return null;
  re.lastIndex = 0;
  return o.regex ? s.replace(re, o.replace) : s.replace(re, () => o.replace);
}

/** Advanced find and replace over values, formulas, comments, hyperlinks and sheet names. Returns counts per kind. */
export function findReplaceAll(app: App, o: FindReplaceOptions): Record<string, number> {
  const book = app.book;
  const counts: Record<string, number> = { values: 0, formulas: 0, comments: 0, links: 0, sheets: 0 };
  const sheets = o.scope === 'workbook' ? book.sheets : [app.sheet];
  book.transact('Find and Replace', () => {
    for (const s of sheets) {
      const inScope = (r: number, c: number) => o.scope !== 'selection' || selRanges(app).some((rg) => r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2);
      for (const [k, cell] of [...s.cells]) {
        const r = keyRow(k), c = keyCol(k);
        if (!inScope(r, c) || !canChange(s, r, c)) continue;
        if (o.inValues !== false && cell.f == null && typeof cell.v === 'string') { const t = replaceText(cell.v, o); if (t != null) { const next: Cell = { ...cell, v: t }; delete next.rich; s.set(r, c, next); counts.values++; } }
        if (o.inFormulas && cell.f != null) { const t = replaceText(cell.f, o); if (t != null) { s.set(r, c, { ...cell, f: t }); counts.formulas++; } }
        if (o.inLinks && cell.link) { const t = replaceText(cell.link, o); if (t != null) { s.set(r, c, { ...s.get(r, c)!, link: t }); counts.links++; } }
      }
      if (o.inComments) for (const [k, cm] of s.comments) { const r = keyRow(k), c = keyCol(k); if (!inScope(r, c)) continue; const t = replaceText(cm.text, o); if (t != null) { book.recordSheet(s); s.comments.set(k, { ...cm, text: t }); counts.comments++; } }
    }
    if (o.inSheetNames && o.scope === 'workbook') { const { renameSheet } = requireStructure(); for (const s of book.sheets) { const t = replaceText(s.name, o); if (t != null && t.trim() && !book.sheetByName(t)) { renameSheet(book, s, t.slice(0, 31)); counts.sheets++; } } }
  });
  return counts;
}

// structure.ts is imported statically elsewhere; a tiny indirection keeps this file's imports short.
import * as structure from '../model/structure';
function requireStructure(): typeof structure { return structure; }

export function sortOrderBy(values: Scalar[], key: 'length' | 'lastWord' | 'natural' | 'firstNumber' | 'random', desc = false): number[] {
  const idx = values.map((_v, i) => i);
  const txt = (v: Scalar) => textOf(v);
  const keyOf = (v: Scalar): number | string => {
    const t = txt(v);
    if (key === 'length') return t.length;
    if (key === 'lastWord') return t.trim().split(/\s+/).pop()?.toLowerCase() || '';
    if (key === 'firstNumber') { const m = /-?\d+(?:[.,]\d+)?/.exec(t); return m ? Number(m[0].replace(',', '.')) : Infinity; }
    return t;
  };
  if (key === 'random') { for (let i = idx.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); [idx[i], idx[j]] = [idx[j], idx[i]]; } return idx; }
  idx.sort((a, b) => {
    const ka = keyOf(values[a]), kb = keyOf(values[b]);
    let x = key === 'natural' ? naturalCompare(String(ka), String(kb)) : typeof ka === 'number' && typeof kb === 'number' ? ka - kb : String(ka).localeCompare(String(kb));
    if (!x) x = a - b;
    return desc ? -x : x;
  });
  return idx;
}

async function advancedSort(app: App, arg?: { key: 'length' | 'lastWord' | 'natural' | 'firstNumber' | 'random'; col?: number; desc?: boolean; header?: boolean }): Promise<void> {
  if (emptySheet(app)) return;
  const s = app.sheet; const rg = selBlock(app);
  const header = headerGuess(s, rg);
  const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Advanced Sort', fields: [{ key: 'col', label: 'Sort by column', type: 'select', value: String(rg.c1), options: columnChoices(s, rg, header) }, { key: 'key', label: 'Order by', type: 'select', value: 'natural', options: [{ value: 'natural', label: 'Natural order (file2 before file10)' }, { value: 'length', label: 'Text length' }, { value: 'lastWord', label: 'Last word (surname)' }, { value: 'firstNumber', label: 'First number inside the text' }, { value: 'random', label: 'Random' }] }, { key: 'desc', label: 'Descending', type: 'checkbox', value: false }, { key: 'header', label: 'The first row is a header', type: 'checkbox', value: header }] }); return r ? { key: r.key, col: Number(r.col), desc: !!r.desc, header: !!r.header } : null; });
  if (!o) return;
  if (hasMergeIn(s, rg)) { toast('Unmerge the cells in the range first.'); return; }
  if (!checkWritable(app, [rg]) || sheetProtected(app, 'sort')) return;
  const body = { ...rg, r1: rg.r1 + (o.header ? 1 : 0) };
  if (body.r2 <= body.r1) { toast('Select two or more rows to sort.'); return; }
  const col = o.col ?? rg.c1;
  const vals: Scalar[] = []; for (let r = body.r1; r <= body.r2; r++) vals.push(s.value(r, col));
  const order = sortOrderBy(vals, o.key, !!o.desc);
  app.book.transact('Advanced Sort', () => reorderRows(s, body, order));
  toast('Sorted');
}

function swapRanges(app: App, a: Range, b: Range, sa: Sheet, sb: Sheet): boolean {
  if (a.r2 - a.r1 !== b.r2 - b.r1 || a.c2 - a.c1 !== b.c2 - b.c1) { toast('The two ranges must be the same size.'); return false; }
  if (hasMergeIn(sa, a) || hasMergeIn(sb, b)) { toast('Unmerge the cells first.'); return false; }
  if (sa === sb && a.r1 <= b.r2 && a.r2 >= b.r1 && a.c1 <= b.c2 && a.c2 >= b.c1) { toast('The two ranges overlap, so they cannot be swapped.'); return false; }
  if (!checkWritable(app, [a], sa) || !checkWritable(app, [b], sb)) return false;
  const A = blockOf(sa, a), B = blockOf(sb, b);
  app.book.transact('Swap Ranges', () => {
    for (let i = 0; i <= a.r2 - a.r1; i++) for (let j = 0; j <= a.c2 - a.c1; j++) {
      sb.set(b.r1 + i, b.c1 + j, movedCell(A[i][j], b.r1 - a.r1, b.c1 - a.c1));
      sa.set(a.r1 + i, a.c1 + j, movedCell(B[i][j], a.r1 - b.r1, a.c1 - b.c1));
    }
  });
  return true;
}

const PALETTE = ['#ffd54f', '#81d4fa', '#a5d6a7', '#f8bbd0', '#ce93d8', '#ffab91', '#80cbc4', '#e6ee9c', '#b0bec5', '#ffe082', '#90caf9', '#c5e1a5', '#f48fb1', '#b39ddb', '#ffcc80', '#9fa8da'];

/** Temporary highlights: the style each cell had before Highlight Cells, so Remove Highlights can restore it. */
const highlights = new Map<string, Map<number, number | undefined>>();

function countSumByColor(sheet: Sheet, cells: { r: number; c: number; cell: Cell | undefined }[], by: 'fill' | 'font'): { color: string; count: number; sum: number; numbers: number }[] {
  const m = new Map<string, { color: string; count: number; sum: number; numbers: number }>();
  for (const x of cells) {
    const st = sheet.styleOf(x.r, x.c);
    const color = normColor(by === 'fill' ? st.fill : st.color) || '(none)';
    let e = m.get(color); if (!e) { e = { color, count: 0, sum: 0, numbers: 0 }; m.set(color, e); }
    if (!hasContent(x.cell)) continue;
    e.count++;
    if (typeof x.cell?.v === 'number') { e.sum += x.cell.v; e.numbers++; }
  }
  return [...m.values()].sort((a, b) => b.count - a.count);
}

defTools([
  {
    id: 'util.findReplaceAdvanced', label: 'Advanced Find and Replace…', icon: 'find_replace', tab: U, group: G_RANGE, edits: true, popular: true,
    tip: 'Find and replace in values, formulas, comments, hyperlinks and sheet names, on this sheet or the whole workbook, with whole-cell and regular-expression options',
    keywords: 'replace all regex comments hyperlinks sheet names workbook', origin: 'ASAP Utilities / Ablebits (Advanced Find and Replace)',
    run: async (app, arg?: FindReplaceOptions) => {
      const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Advanced Find and Replace', fields: [
        { key: 'find', label: 'Find', type: 'text', value: '', autofocus: true, half: true }, { key: 'replace', label: 'Replace with', type: 'text', value: '', half: true },
        { key: 'scope', label: 'Look in', type: 'radio', value: singleCell(app) ? 'sheet' : 'selection', options: [{ value: 'selection', label: 'Selection' }, { value: 'sheet', label: 'This sheet' }, { value: 'workbook', label: 'Whole workbook' }] },
        { key: 'inValues', label: 'Values', type: 'checkbox', value: true, half: true }, { key: 'inFormulas', label: 'Formulas', type: 'checkbox', value: true, half: true }, { key: 'inComments', label: 'Comments and notes', type: 'checkbox', value: false, half: true }, { key: 'inLinks', label: 'Hyperlink addresses', type: 'checkbox', value: false, half: true }, { key: 'inSheetNames', label: 'Sheet names (workbook scope)', type: 'checkbox', value: false, half: true },
        { key: 'matchCase', label: 'Match case', type: 'checkbox', value: false, half: true }, { key: 'wholeCell', label: 'Match the entire cell', type: 'checkbox', value: false, half: true }, { key: 'regex', label: 'Regular expression ($1 in the replacement)', type: 'checkbox', value: false },
      ] }); return r as FindReplaceOptions | null; });
      if (!o || !o.find) return;
      const counts = findReplaceAll(app, o);
      const total = Object.values(counts).reduce((a, b) => a + b, 0);
      toast(total ? `Replaced in ${Object.entries(counts).filter(([, n]) => n).map(([k, n]) => `${n} ${k}`).join(', ')}` : 'Nothing was found.');
    },
  },
  {
    id: 'util.replaceFromList', label: 'Replace Values from a List…', icon: 'swap_horizontal_circle', tab: U, group: G_RANGE, edits: true, popular: true,
    tip: 'Replace many values at once using a two-column table of old and new values (whole cells or inside the text)', keywords: 'mass replace mapping translate codes lookup table', origin: 'ASAP Utilities (Replace values in selection with values from a list)',
    run: async (app, arg?: { list: string; whole?: boolean; matchCase?: boolean }) => {
      const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Replace Values from a List', help: 'The list is a two-column range: the first column holds the values to find, the second what to put instead (e.g. Sheet2!A1:B20).', fields: [{ key: 'list', label: 'List range (old, new)', type: 'text', value: '', placeholder: 'Sheet2!A2:B50' }, { key: 'whole', label: 'Replace whole cells only (otherwise inside the text too)', type: 'checkbox', value: true }, { key: 'matchCase', label: 'Match case', type: 'checkbox', value: false }] }); return r ? { list: r.list, whole: !!r.whole, matchCase: !!r.matchCase } : null; });
      if (!o) return;
      const res = rangeFromText(app, o.list);
      if (!res || res.range.c2 - res.range.c1 < 1) { toast('The list must be a range with at least two columns.'); return; }
      const pairs: [string, string][] = [];
      for (let r = res.range.r1; r <= res.range.r2; r++) { const a = shownText(res.sheet, r, res.range.c1), b = shownText(res.sheet, r, res.range.c1 + 1); if (a !== '') pairs.push([o.matchCase ? a : a.toLowerCase(), b]); }
      if (!pairs.length) { toast('The list is empty.'); return; }
      const map = new Map(pairs);
      pairs.sort((x, y) => y[0].length - x[0].length);
      let n = 0;
      app.book.transact('Replace from List', () => {
        for (const x of selCells(app, { existingOnly: true })) {
          const cell = x.cell!; if (cell.f != null || cell.sp != null || !canChange(app.sheet, x.r, x.c)) continue;
          const t = typeof cell.v === 'string' ? cell.v : shownText(app.sheet, x.r, x.c);
          if (o.whole) { const hit = map.get(o.matchCase ? t : t.toLowerCase()); if (hit === undefined) continue; const nv = numOf(hit); s_set(app.sheet, x, nv != null && typeof cell.v === 'number' ? nv : hit); n++; }
          else { let out = t; for (const [a, b] of pairs) out = out.split(o.matchCase ? a : new RegExp(a.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi') as any).join(b); if (out !== t) { s_set(app.sheet, x, out); n++; } }
        }
      });
      toast(`${plural(n, 'cell')} replaced`);
    },
  },
  { id: 'util.sortAdvanced', label: 'Advanced Sort…', icon: 'sort', tab: U, group: G_RANGE, edits: true, tip: 'Sort rows by text length, last word, natural order (file2 before file10), the first number in the text, or at random', keywords: 'sort length surname natural random', origin: 'ASAP Utilities / Ablebits (Advanced Sort)', run: (app, arg) => advancedSort(app, arg) },
  { id: 'util.sortNatural', label: 'Sort in Natural Order', icon: 'sort_by_alpha', tab: U, group: G_RANGE, edits: true, tip: 'Sort so that item2 comes before item10', run: (app) => advancedSort(app, { key: 'natural', col: selBlock(app).c1, header: headerGuess(app.sheet, selBlock(app)) }) },
  { id: 'util.sortByLength', label: 'Sort by Text Length', icon: 'short_text', tab: U, group: G_RANGE, edits: true, run: (app) => advancedSort(app, { key: 'length', col: selBlock(app).c1, header: headerGuess(app.sheet, selBlock(app)) }) },
  { id: 'util.sortByLastWord', label: 'Sort by Last Word (Surname)', icon: 'person', tab: U, group: G_RANGE, edits: true, origin: 'Genealogy and mailing-list practice', run: (app) => advancedSort(app, { key: 'lastWord', col: selBlock(app).c1, header: headerGuess(app.sheet, selBlock(app)) }) },
  {
    id: 'util.swapRanges', label: 'Swap Two Ranges…', icon: 'swap_horiz', tab: U, group: G_RANGE, edits: true, popular: true, tip: 'Exchange the contents and formatting of the selection with another range of the same size (even on another sheet)', keywords: 'swap exchange ranges', origin: 'Kutools / Ablebits (Swap Ranges)',
    run: async (app, arg?: { a?: string; b: string }) => {
      const a0 = clipToUsed(app.sheet, activeRange(app));
      const o = await argOrAsk<{ a?: string; b: string }>(arg, async () => { const t = await promptDialog('Swap Two Ranges', `Swap ${rangeName(a0)} with`, ''); return t ? { b: t } : null; });
      if (!o) return;
      const A = o.a ? rangeFromText(app, o.a) : { sheet: app.sheet, range: a0 };
      const B = rangeFromText(app, o.b);
      if (!A || !B) { toast('That range is not valid.'); return; }
      if (swapRanges(app, A.range, B.range, A.sheet, B.sheet)) toast('Ranges swapped');
    },
  },
  {
    id: 'util.colorDuplicates', label: 'Color Duplicate Values', icon: 'palette', tab: U, group: G_RANGE, edits: true, popular: true, tip: 'Give each set of duplicate values its own fill color, so matching rows can be spotted at a glance', keywords: 'duplicates color groups highlight', origin: 'ASAP Utilities (Color each duplicate with a unique color)',
    run: (app, arg?: { col?: number }) => {
      const s = app.sheet; const rg = selBlock(app);
      const col = arg?.col ?? (rg.c1 === rg.c2 ? rg.c1 : -1);
      const cells = col >= 0 ? Array.from({ length: rg.r2 - rg.r1 + 1 }, (_x, i) => ({ r: rg.r1 + i, c: col, cell: s.get(rg.r1 + i, col) })) : selCells(app);
      const groups = new Map<string, { r: number; c: number }[]>();
      for (const x of cells) { if (!hasContent(x.cell)) continue; const k = shownText(s, x.r, x.c).toLowerCase(); (groups.get(k) || groups.set(k, []).get(k)!).push({ r: x.r, c: x.c }); }
      const dup = [...groups.values()].filter((g) => g.length > 1);
      if (!dup.length) { toast('No duplicate values in the selection.'); return; }
      app.book.transact('Color Duplicates', () => dup.forEach((g, i) => { for (const p of g) { const cells2 = col >= 0 ? Array.from({ length: rg.c2 - rg.c1 + 1 }, (_x, j) => ({ r: p.r, c: rg.c1 + j })) : [p]; for (const q of cells2) s.set(q.r, q.c, { ...(s.get(q.r, q.c) || {}), s: app.book.derivedStyle(s.styleIdOf(q.r, q.c), { fill: PALETTE[i % PALETTE.length] }) }); } }));
      toast(`${plural(dup.length, 'duplicate group')} colored`);
    },
  },
  {
    id: 'util.mergeToOneCell', label: 'Join Selected Cells into One…', icon: 'join_inner', tab: U, group: G_RANGE, edits: true, tip: 'Put the text of all selected cells into the first one, separated by a space, comma, line break or your own text (the other cells are cleared)', keywords: 'concatenate merge keep contents combine', origin: 'ASAP Utilities / Kutools (Combine without losing data)',
    run: async (app, arg?: { sep: string; byRow?: boolean; skipBlanks?: boolean }) => {
      const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Join Selected Cells', fields: [{ key: 'sep', label: 'Separator', type: 'combo', value: ', ', options: [', ', ' ', '; ', '\\n (line break)', ' | ', '-'] }, { key: 'byRow', label: 'Join each row separately (into its first cell)', type: 'checkbox', value: false }, { key: 'skipBlanks', label: 'Skip blank cells', type: 'checkbox', value: true }] }); return r ? { sep: String(r.sep).startsWith('\\n') ? '\n' : String(r.sep), byRow: !!r.byRow, skipBlanks: !!r.skipBlanks } : null; });
      if (!o) return;
      const s = app.sheet; const rg = clipToUsed(s, activeRange(app));
      if (!checkWritable(app, [rg])) return;
      app.book.transact('Join Cells', () => {
        const lines = o.byRow ? Array.from({ length: rg.r2 - rg.r1 + 1 }, (_x, i) => ({ r1: rg.r1 + i, r2: rg.r1 + i })) : [{ r1: rg.r1, r2: rg.r2 }];
        for (const ln of lines) {
          const parts: string[] = [];
          for (let r = ln.r1; r <= ln.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) { const t = shownText(s, r, c); if (t !== '' || !o.skipBlanks) parts.push(t); if (!(r === ln.r1 && c === rg.c1)) { const cur = s.get(r, c); s.set(r, c, cur?.s ? { s: cur.s } : undefined); } }
          const first = s.get(ln.r1, rg.c1);
          s.set(ln.r1, rg.c1, { ...(first?.s ? { s: first.s } : {}), v: parts.join(o.sep) });
        }
      });
    },
  },
  {
    id: 'util.centerAcrossInstead', label: 'Convert Merged Cells to Center Across Selection', icon: 'format_align_center', tab: U, group: G_RANGE, edits: true, tip: 'Unmerge the merged cells in the selection and center their text across the same columns instead, which keeps sorting, filtering and copying working', origin: 'ASAP Utilities',
    run: (app) => {
      const s = app.sheet; const sel = selRanges(app);
      const merges = s.merges.filter((m) => m.r1 === m.r2 && sel.some((rg) => m.r1 <= rg.r2 && m.r2 >= rg.r1 && m.c1 <= rg.c2 && m.c2 >= rg.c1));
      if (!merges.length) { toast('No single-row merged cells in the selection.'); return; }
      app.book.transact('Center Across Selection', () => { app.book.recordSheet(s); for (const m of merges) { for (let c = m.c1; c <= m.c2; c++) s.set(m.r1, c, { ...(s.get(m.r1, c) || {}), s: app.book.derivedStyle(s.styleIdOf(m.r1, c), { hAlign: 'centerContinuous' }) }); } s.merges = s.merges.filter((m) => !merges.includes(m)); });
      toast(`${plural(merges.length, 'merged range')} converted`);
    },
  },
  { id: 'util.lockNonEmpty', label: 'Lock Cells with Content, Unlock the Rest', icon: 'lock', tab: U, group: G_RANGE, edits: true, tip: 'In the selection, lock every cell that has a value or formula and unlock the blanks, then protect the sheet to make a form', origin: 'ASAP Utilities', run: (app) => { let n = 0; app.book.transact('Lock Filled Cells', () => { for (const x of selCells(app)) { const lock = hasContent(x.cell); app.sheet.set(x.r, x.c, { ...(x.cell || {}), s: app.book.derivedStyle(app.sheet.styleIdOf(x.r, x.c), { locked: lock ? undefined : false }) }); n++; } }); toast(`${plural(n, 'cell')} updated; use Review ▸ Protect Sheet to enforce it`); } },
  { id: 'util.lockFormulas', label: 'Lock Formula Cells Only', icon: 'function', tab: U, group: G_RANGE, edits: true, tip: 'Lock the formulas on this sheet and unlock everything else', origin: 'ASAP Utilities', run: (app) => { const s = app.sheet; const u = s.usedRange(); if (!u) return; app.book.transact('Lock Formulas', () => { for (let r = u.r1; r <= u.r2; r++) for (let c = u.c1; c <= u.c2; c++) { const cell = s.get(r, c); const lock = !!cell && cell.f != null; const cur = s.styleOf(r, c).locked !== false; if (lock === cur) continue; s.set(r, c, { ...(cell || {}), s: app.book.derivedStyle(s.styleIdOf(r, c), { locked: lock ? undefined : false }) }); } }); toast('Formula cells locked, the rest unlocked'); } },
  { id: 'util.unlockSelection', label: 'Unlock the Selected Cells', icon: 'lock_open', tab: U, group: G_RANGE, edits: true, run: (app) => { app.book.transact('Unlock Cells', () => { for (const x of selCells(app)) app.sheet.set(x.r, x.c, { ...(x.cell || {}), s: app.book.derivedStyle(app.sheet.styleIdOf(x.r, x.c), { locked: false }) }); }); } },
  { id: 'util.lockSelection', label: 'Lock the Selected Cells', icon: 'lock', tab: U, group: G_RANGE, edits: true, run: (app) => { app.book.transact('Lock Cells', () => { for (const x of selCells(app)) app.sheet.set(x.r, x.c, { ...(x.cell || {}), s: app.book.derivedStyle(app.sheet.styleIdOf(x.r, x.c), { locked: undefined }) }); }); } },
  {
    id: 'util.copyMultiSelection', label: 'Copy a Multi-Area Selection', icon: 'content_copy', tab: U, group: G_RANGE, noRecord: true, tip: 'Copy several selected areas to the clipboard as one block of text, one area under the other', origin: 'ASAP Utilities / Kutools (Copy Ranges)',
    run: async (app) => { const s = app.sheet; const text = selRanges(app).map((rg) => rangeToTsv(s, clipToUsed(s, rg))).join('\n'); await copyText(text, `${selRanges(app).length} area${selRanges(app).length === 1 ? '' : 's'}`); },
  },
  {
    id: 'util.highlightCells', label: 'Highlight Selection (Temporary)', icon: 'highlight', tab: U, group: G_RANGE, edits: true, tip: 'Mark the selected cells yellow while you work; Remove Highlights puts their old formatting back', origin: 'ASAP Utilities (Mark cells)',
    run: (app, arg?: { color?: string }) => { const s = app.sheet; const store = highlights.get(s.id) || highlights.set(s.id, new Map()).get(s.id)!; app.book.transact('Highlight', () => { for (const x of selCells(app)) { const k = cellKey(x.r, x.c); if (!store.has(k)) store.set(k, x.cell?.s); s.set(x.r, x.c, { ...(x.cell || {}), s: app.book.derivedStyle(s.styleIdOf(x.r, x.c), { fill: arg?.color || '#ffff00' }) }); } }); },
  },
  { id: 'util.removeHighlights', label: 'Remove Highlights', icon: 'format_color_reset', tab: U, group: G_RANGE, edits: true, run: (app) => { const s = app.sheet; const store = highlights.get(s.id); if (!store?.size) { toast('No highlights on this sheet.'); return; } app.book.transact('Remove Highlights', () => { for (const [k, sid] of store) { const r = keyRow(k), c = keyCol(k); const cell = s.get(r, c); if (!cell && sid == null) continue; const next: Cell = { ...(cell || {}) }; if (sid) next.s = sid; else delete next.s; s.set(r, c, Object.keys(next).length ? next : undefined); } }); highlights.delete(s.id); } },
  {
    id: 'util.countSumByColor', label: 'Count and Sum by Color…', icon: 'format_color_fill', tab: U, group: G_RANGE, edits: true, popular: true, tip: 'How many selected cells have each fill (or font) color, and the sum of their numbers; optionally as a report sheet', keywords: 'count by color sum by color', origin: 'ASAP Utilities / Ablebits (Count & Sum by Color)',
    run: async (app, arg?: { by: 'fill' | 'font'; sheet?: boolean }) => {
      const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Count and Sum by Color', fields: [{ key: 'by', label: 'Group by', type: 'radio', value: 'fill', options: [{ value: 'fill', label: 'Fill color' }, { value: 'font', label: 'Font color' }] }, { key: 'sheet', label: 'Write the result to a new sheet', type: 'checkbox', value: false }] }); return r ? { by: r.by, sheet: !!r.sheet } : null; });
      if (!o) return;
      const rows = countSumByColor(app.sheet, selCells(app), o.by);
      if (o.sheet) { report(app, 'By Color', ['Color', 'Cells', 'Numbers', 'Sum', 'Average'], rows.map((x) => [{ v: x.color, style: x.color !== '(none)' ? (o.by === 'fill' ? { fill: x.color } : { color: x.color }) : {} }, x.count, x.numbers, x.sum, x.numbers ? x.sum / x.numbers : null])); return; }
      const { alertDialog } = await import('../ui/dialog');
      const { h } = await import('../ui/dom');
      const table = h('table', { class: 'ql-info-table' }, h('tr', null, h('th', null, 'Color'), h('th', null, 'Cells'), h('th', null, 'Sum')), ...rows.map((x) => h('tr', null, h('td', null, h('span', { style: `display:inline-block;width:14px;height:14px;border:1px solid #888;vertical-align:middle;background:${x.color === '(none)' ? 'transparent' : x.color}` }), ' ' + x.color), h('td', null, String(x.count)), h('td', null, x.numbers ? String(+x.sum.toPrecision(12)) : ''))));
      await alertDialog('Count and Sum by Color', table);
    },
  },
  {
    id: 'util.deleteNames', label: 'Delete Defined Names…', icon: 'label_off', tab: U, group: G_RANGE, edits: true, tip: 'Delete all defined names, only the broken ones (#REF!), or only hidden ones', origin: 'ASAP Utilities',
    run: async (app, arg?: 'all' | 'broken' | 'hidden') => {
      const which = await argOrAsk(arg, async () => { const { choiceDialog } = await import('../ui/dialog'); const i = await choiceDialog('Delete Defined Names', `The workbook has ${app.book.names.length} names. Delete:`, ['All names', 'Names with #REF! errors', 'Hidden names']); return i < 0 ? null : (['all', 'broken', 'hidden'] as const)[i]; });
      if (!which) return;
      const victims = app.book.names.filter((n) => which === 'all' || (which === 'broken' ? /#REF!/i.test(n.ref) : !!n.hidden));
      if (!victims.length) { toast('No names to delete.'); return; }
      app.book.transact('Delete Names', () => { app.book.recordBook(); app.book.names = app.book.names.filter((n) => !victims.includes(n)); });
      toast(`${plural(victims.length, 'name')} deleted`);
    },
  },
  {
    id: 'util.compareRanges', label: 'Compare Two Ranges…', icon: 'compare', tab: U, group: G_RANGE, edits: true, tip: 'Color the cells of the selection that differ from (or match) another range of the same size, or list the differences', keywords: 'compare columns differences same different', origin: 'Kutools (Compare Cells) / Ablebits (Compare Two Sheets)',
    run: async (app, arg?: { other: string; mode: 'different' | 'same' | 'list' }) => {
      const a = clipToUsed(app.sheet, activeRange(app));
      const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Compare Two Ranges', fields: [{ key: 'other', label: `Compare ${rangeName(a)} with`, type: 'text', value: '' }, { key: 'mode', label: 'Then', type: 'radio', value: 'different', options: [{ value: 'different', label: 'Color the cells that differ' }, { value: 'same', label: 'Color the cells that are the same' }, { value: 'list', label: 'List the differences on a new sheet' }] }] }); return r ? { other: r.other, mode: r.mode } : null; });
      if (!o) return;
      const B = rangeFromText(app, o.other); if (!B) { toast('That range is not valid.'); return; }
      const s = app.sheet;
      const diffs: { r: number; c: number; a: string; b: string }[] = [];
      const same: { r: number; c: number }[] = [];
      for (let i = 0; i <= a.r2 - a.r1; i++) for (let j = 0; j <= a.c2 - a.c1; j++) { const x = shownText(s, a.r1 + i, a.c1 + j), y = shownText(B.sheet, B.range.r1 + i, B.range.c1 + j); if (x === y) same.push({ r: a.r1 + i, c: a.c1 + j }); else diffs.push({ r: a.r1 + i, c: a.c1 + j, a: x, b: y }); }
      if (o.mode === 'list') { report(app, 'Differences', ['Cell', 'This range', 'Other range'], diffs.map((d) => [cellName(d.r, d.c), d.a, d.b])); return; }
      const targets = o.mode === 'different' ? diffs : same;
      if (!targets.length) { toast(o.mode === 'different' ? 'The ranges are identical.' : 'No cells are the same.'); return; }
      app.book.transact('Compare Ranges', () => { for (const t of targets) s.set(t.r, t.c, { ...(s.get(t.r, t.c) || {}), s: app.book.derivedStyle(s.styleIdOf(t.r, t.c), { fill: o.mode === 'different' ? '#ffcdd2' : '#c8e6c9' }) }); });
      toast(`${plural(targets.length, 'cell')} colored`);
    },
  },
  {
    id: 'util.moveToNewSheet', label: 'Move Selection to a New Sheet', icon: 'drive_file_move', tab: U, group: G_RANGE, edits: true, tip: 'Cut the selected block and paste it at A1 of a new sheet; formulas pointing at it follow',
    run: (app) => { const s = app.sheet; const rg = clipToUsed(s, activeRange(app)); const book = app.book; let dest!: Sheet; book.transact('Move to New Sheet', () => { book.recordBook(); dest = new Sheet(book, book.uniqueSheetName('Moved')); book.sheets.splice(book.activeSheet + 1, 0, dest); structure.moveBlock(s, rg, -rg.r1, -rg.c1, dest); }); app.ui.rebuild(); app.exec('sheet.activate', book.sheets.indexOf(dest)); },
  },
  { id: 'util.menuRange', label: 'Range', icon: 'select_all', tab: U, group: G_RANGE, kind: 'menu', noRecord: true, run: () => undefined, menu: () => menuOf(['util.findReplaceAdvanced', 'util.replaceFromList', null, 'util.sortAdvanced', 'util.sortNatural', 'util.sortByLength', 'util.sortByLastWord', 'data.sortByColor', 'tools.shuffle', null, 'util.swapRanges', 'util.moveToNewSheet', 'util.copyMultiSelection', 'edit.pasteValues', 'edit.pasteFormats', 'tools.transposeInPlace', null, 'util.colorDuplicates', 'util.countSumByColor', 'util.compareRanges', 'util.highlightCells', 'util.removeHighlights', 'tools.findDuplicates', 'data.removeDuplicates', null, 'util.mergeToOneCell', 'util.centerAcrossInstead', 'tools.unmergeFill', null, 'util.lockNonEmpty', 'util.lockFormulas', 'util.lockSelection', 'util.unlockSelection', null, 'formulas.nameManager', 'formulas.createFromSelection', 'util.deleteNames']) },
]);

function s_set(s: Sheet, x: { r: number; c: number; cell: Cell | undefined }, v: Scalar): void { const next: Cell = { ...(x.cell || {}), v }; delete next.rich; s.set(x.r, x.c, next); }

// ============================================================================================
// Fill

function fillTarget(app: App): Range | null {
  const rg = activeRange(app);
  if (rg.r2 - rg.r1 > 100000 || rg.c2 - rg.c1 > 1000) { toast('Select a smaller range to fill.'); return null; }
  return rg;
}

defTools([
  {
    id: 'util.fillWith', label: 'Fill Selection with a Value or Formula…', icon: 'format_color_fill', tab: U, group: G_FILL, edits: true, popular: true, tip: 'Put one value in every selected cell, or one formula whose references are adjusted for each cell', keywords: 'fill same value every cell formula', origin: 'ASAP Utilities',
    run: async (app, arg?: { text: string; blanksOnly?: boolean }) => {
      const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Fill Selection', fields: [{ key: 'text', label: 'Value or formula (relative to the first cell)', type: 'text', value: '', autofocus: true }, { key: 'blanksOnly', label: 'Fill blank cells only', type: 'checkbox', value: false }] }); return r ? { text: String(r.text ?? ''), blanksOnly: !!r.blanksOnly } : null; });
      if (!o) return;
      const rg = fillTarget(app); if (!rg || !checkWritable(app, [rg])) return;
      const s = app.sheet; const { parseInput } = await import('./ops');
      let n = 0;
      app.book.transact('Fill Selection', () => {
        for (const x of selCells(app, { fill: true })) {
          if (x.r > rg.r2 || x.c > rg.c2) continue;
          if (o.blanksOnly && hasContent(x.cell)) continue;
          if (!canChange(s, x.r, x.c)) continue;
          if (o.text.startsWith('=')) { s.set(x.r, x.c, { ...(x.cell?.s ? { s: x.cell.s } : {}), f: shiftFormula(o.text.slice(1), x.r - rg.r1, x.c - rg.c1) }); }
          else { const p = parseInput(app, o.text, x.r, x.c); const next: Cell = { ...(x.cell?.s ? { s: x.cell.s } : {}), ...p.cell }; if (p.format && !s.styleOf(x.r, x.c).numFmt) next.s = app.book.derivedStyle(x.cell?.s ?? 0, { numFmt: p.format }); s.set(x.r, x.c, next); }
          n++;
        }
      });
      toast(`${plural(n, 'cell')} filled`);
    },
  },
  {
    id: 'util.fillCycle', label: 'Fill with a Repeating List…', icon: 'repeat', tab: U, group: G_FILL, edits: true, tip: 'Fill the selection by repeating a list of values (Mon, Tue, Wed…; A, B, C; red, green) down or across', keywords: 'cycle repeat pattern list', origin: 'ASAP Utilities (Fill with cyclic values)',
    run: async (app, arg?: { items: string[]; dir?: 'down' | 'right' }) => {
      const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Fill with a Repeating List', fields: [{ key: 'items', label: 'Values, one per line (or comma-separated)', type: 'textarea', rows: 5, value: '' }, { key: 'dir', label: 'Direction', type: 'radio', value: 'down', options: [{ value: 'down', label: 'Down the columns' }, { value: 'right', label: 'Across the rows' }] }] }); if (!r) return null; const items = String(r.items || '').split(/\r?\n|,/).map((x) => x.trim()).filter(Boolean); return { items, dir: r.dir }; });
      if (!o || !o.items.length) return;
      const rg = fillTarget(app); if (!rg || !checkWritable(app, [rg])) return;
      const s = app.sheet; const { parseInput } = await import('./ops');
      const r2 = Math.min(rg.r2, rg.r1 + 100000), c2 = Math.min(rg.c2, rg.c1 + 1000);
      app.book.transact('Fill with List', () => { let i = 0; const put = (r: number, c: number) => { const p = parseInput(app, o.items[i++ % o.items.length], r, c); const cur = s.get(r, c); s.set(r, c, { ...(cur?.s ? { s: cur.s } : {}), ...p.cell }); }; if (o.dir === 'right') { for (let r = rg.r1; r <= r2; r++) { i = 0; for (let c = rg.c1; c <= c2; c++) put(r, c); } } else { for (let c = rg.c1; c <= c2; c++) { i = 0; for (let r = rg.r1; r <= r2; r++) put(r, c); } } });
    },
  },
  { id: 'util.fillRowNumbers', label: 'Fill with Row Numbers (1, 2, 3…)', icon: 'format_list_numbered', tab: U, group: G_FILL, edits: true, tip: 'Number the rows of the selection from 1, restarting in each column', origin: 'ASAP Utilities', run: (app) => { const rg = fillTarget(app); if (!rg || !checkWritable(app, [rg])) return; const s = app.sheet; app.book.transact('Fill Row Numbers', () => { for (let c = rg.c1; c <= Math.min(rg.c2, rg.c1 + 1000); c++) for (let r = rg.r1; r <= Math.min(rg.r2, rg.r1 + 100000); r++) writeOut(s, r, c, r - rg.r1 + 1); }); } },
  { id: 'util.fillLetters', label: 'Fill with Letters (A, B, C…)', icon: 'abc', tab: U, group: G_FILL, edits: true, tip: 'Letter series down the selection: A…Z, AA, AB…', run: (app) => { const rg = fillTarget(app); if (!rg || !checkWritable(app, [rg])) return; const s = app.sheet; app.book.transact('Fill Letters', () => { for (let c = rg.c1; c <= Math.min(rg.c2, rg.c1 + 1000); c++) for (let r = rg.r1; r <= Math.min(rg.r2, rg.r1 + 100000); r++) writeOut(s, r, c, colName(r - rg.r1)); }); } },
  { id: 'util.fillFirstCell', label: 'Fill Each Column with Its First Cell', icon: 'vertical_align_bottom', tab: U, group: G_FILL, edits: true, tip: 'Copy the first cell of each selected column into the rest of the column (formulas adjusted)', origin: 'ASAP Utilities', run: (app) => { const rg = fillTarget(app); if (!rg || !checkWritable(app, [rg])) return; const s = app.sheet; app.book.transact('Fill with First Cell', () => { for (let c = rg.c1; c <= Math.min(rg.c2, rg.c1 + 1000); c++) { const top = s.get(rg.r1, c); for (let r = rg.r1 + 1; r <= Math.min(rg.r2, rg.r1 + 100000); r++) s.set(r, c, movedCell(top, r - rg.r1, 0)); } }); } },
  {
    id: 'util.fillDownToLast', label: 'Fill Down to the Last Row of the Data', icon: 'south', tab: U, group: G_FILL, edits: true, popular: true, keys: ['Mod-Alt-d'], tip: 'Copy the active cell (or the selected row of cells) down as far as the neighboring column has data, like double-clicking the fill handle', origin: 'ASAP Utilities (Quick fill down)',
    run: (app) => {
      const s = app.sheet; const rg = activeRange(app);
      const u = s.usedRange(false); if (!u) return;
      const neighbors = [rg.c1 - 1, rg.c2 + 1].filter((c) => c >= 0);
      let last = rg.r1;
      for (const c of neighbors) { let r = rg.r1; while (r + 1 <= u.r2 && hasContent(s.get(r + 1, c))) r++; last = Math.max(last, r); }
      if (last <= rg.r1) { toast('The neighboring columns have no data below.'); return; }
      const target = { ...rg, r2: last };
      if (!checkWritable(app, [target])) return;
      app.book.transact('Fill Down', () => { for (let c = rg.c1; c <= rg.c2; c++) { const top = s.get(rg.r1, c); for (let r = rg.r1 + 1; r <= last; r++) s.set(r, c, movedCell(top, r - rg.r1, 0)); } });
      app.grid.select(rg.r1, rg.c1, { range: target });
    },
  },
  { id: 'util.fillToday', label: 'Fill with Today\'s Date', icon: 'today', tab: U, group: G_FILL, edits: true, tip: 'The current date as a fixed value (not a formula) in every selected cell', origin: 'ASAP Utilities', run: (app) => { const rg = fillTarget(app); if (!rg || !checkWritable(app, [rg])) return; const d = new Date(); const serial = Math.floor(jsDateToSerial(new Date(d.getFullYear(), d.getMonth(), d.getDate()), app.book.calc.date1904)); app.book.transact('Fill Date', () => { for (const x of selCells(app, { fill: true })) if (x.r <= rg.r2 && x.c <= rg.c2) writeOut(app.sheet, x.r, x.c, { v: serial, style: { numFmt: 'yyyy-mm-dd' } }); }); } },
  { id: 'util.fillNow', label: 'Fill with the Current Date and Time', icon: 'schedule', tab: U, group: G_FILL, edits: true, run: (app) => { const rg = fillTarget(app); if (!rg || !checkWritable(app, [rg])) return; const serial = jsDateToSerial(new Date(), app.book.calc.date1904); app.book.transact('Fill Date and Time', () => { for (const x of selCells(app, { fill: true })) if (x.r <= rg.r2 && x.c <= rg.c2) writeOut(app.sheet, x.r, x.c, { v: serial, style: { numFmt: 'yyyy-mm-dd hh:mm' } }); }); } },
  {
    id: 'util.fillCounts', label: 'Fill the Next Column with Occurrence Counts', icon: 'tag', tab: U, group: G_FILL, edits: true, tip: 'Beside each selected value, how many times it occurs in the selection (a COUNTIF without the formula)', origin: 'ASAP Utilities',
    run: (app) => { const s = app.sheet; const rg = clipToUsed(s, activeRange(app)); const target = { r1: rg.r1, r2: rg.r2, c1: rg.c2 + 1, c2: rg.c2 + 1 }; if (!checkWritable(app, [target])) return; const counts = new Map<string, number>(); for (let r = rg.r1; r <= rg.r2; r++) { const k = shownText(s, r, rg.c1).toLowerCase(); if (k) counts.set(k, (counts.get(k) || 0) + 1); } app.book.transact('Occurrence Counts', () => { for (let r = rg.r1; r <= rg.r2; r++) { const k = shownText(s, r, rg.c1).toLowerCase(); if (k) writeOut(s, r, target.c1, counts.get(k) || 0); } }); },
  },
  { id: 'util.menuFill', label: 'Fill', icon: 'format_color_fill', tab: U, group: G_FILL, kind: 'menu', noRecord: true, run: () => undefined, menu: () => menuOf(['util.fillWith', 'util.fillCycle', 'util.fillRowNumbers', 'util.fillLetters', 'util.fillFirstCell', 'util.fillDownToLast', 'util.fillToday', 'util.fillNow', 'util.fillCounts', null, 'tools.fillBlanks', 'tools.sequence', 'tools.randomData', 'edit.fillSeries', 'edit.fillDown', 'edit.fillRight', 'sheet.fillAcross']) },
]);

// ============================================================================================
// Columns & Rows

function insertEvery(app: App, axis: 'row' | 'col', every: number, count: number, startAfter: number): void {
  const s = app.sheet; const rg = selBlock(app);
  if (sheetProtected(app, axis === 'row' ? 'insertRows' : 'insertColumns')) return;
  const first = axis === 'row' ? rg.r1 : rg.c1, last = axis === 'row' ? rg.r2 : rg.c2;
  const spots: number[] = [];
  for (let l = first + startAfter; l <= last; l += every) spots.push(l);
  if (!spots.length) { toast('Nothing to insert in that range.'); return; }
  app.book.transact(axis === 'row' ? 'Insert Rows Every Nth' : 'Insert Columns Every Nth', () => { for (const at of spots.reverse()) (axis === 'row' ? insertRows : insertCols)(s, at, count); });
  app.grid.updateExtent(); app.grid.layout();
  toast(`${plural(spots.length * count, axis === 'row' ? 'row' : 'column')} inserted`);
}

function toPx(v: number, unit: string): number { return unit === 'cm' ? v * PX_PER_CM : unit === 'in' ? v * PX_PER_IN : unit === 'pt' ? v * PX_PER_PT : unit === 'chars' ? v * 7 + 5 : v; }

function moveLines(app: App, axis: 'row' | 'col', delta: number): void {
  const s = app.sheet; const rg = clipToUsed(s, activeRange(app), 1);
  const u = s.usedRange() || rg;
  const a = axis === 'row' ? rg.r1 : rg.c1, b = axis === 'row' ? rg.r2 : rg.c2;
  const n = b - a + 1;
  const target = delta < 0 ? a - 1 : b + 1;
  if (target < 0) { toast('Already at the edge.'); return; }
  const block: Range = axis === 'row' ? { r1: a, r2: b, c1: 0, c2: Math.max(u.c2, rg.c2) } : { c1: a, c2: b, r1: 0, r2: Math.max(u.r2, rg.r2) };
  const other: Range = axis === 'row' ? { ...block, r1: target, r2: target } : { ...block, c1: target, c2: target };
  if (hasMergeIn(s, block) || hasMergeIn(s, other)) { toast('Unmerge the cells first.'); return; }
  const A = blockOf(s, block), B = blockOf(s, other);
  app.book.transact(axis === 'row' ? 'Move Rows' : 'Move Columns', () => {
    // The neighboring line takes the block's old edge position; the block shifts by one.
    if (axis === 'row') {
      const newOther = delta < 0 ? b : a;
      B[0].forEach((cell, j) => s.set(newOther, block.c1 + j, movedCell(cell, newOther - target, 0)));
      A.forEach((row, i) => row.forEach((cell, j) => s.set(a + delta + i, block.c1 + j, movedCell(cell, delta, 0))));
      const h = s.rowH.get(target), hid = s.hiddenRows.has(target);
      const hs = Array.from({ length: n }, (_x, i) => s.rowH.get(a + i)); const hids = Array.from({ length: n }, (_x, i) => s.hiddenRows.has(a + i));
      for (let i = 0; i < n; i++) { const r = a + delta + i; if (hs[i] != null) s.rowH.set(r, hs[i]!); else s.rowH.delete(r); if (hids[i]) s.hiddenRows.add(r); else s.hiddenRows.delete(r); }
      if (h != null) s.rowH.set(newOther, h); else s.rowH.delete(newOther); if (hid) s.hiddenRows.add(newOther); else s.hiddenRows.delete(newOther);
    } else {
      const newOther = delta < 0 ? b : a;
      B.forEach((row, i) => s.set(block.r1 + i, newOther, movedCell(row[0], 0, newOther - target)));
      A.forEach((row, i) => row.forEach((cell, j) => s.set(block.r1 + i, a + delta + j, movedCell(cell, 0, delta))));
      const w = s.colW.get(target), hid = s.hiddenCols.has(target);
      const ws = Array.from({ length: n }, (_x, j) => s.colW.get(a + j)); const hids = Array.from({ length: n }, (_x, j) => s.hiddenCols.has(a + j));
      for (let j = 0; j < n; j++) { const c = a + delta + j; if (ws[j] != null) s.colW.set(c, ws[j]!); else s.colW.delete(c); if (hids[j]) s.hiddenCols.add(c); else s.hiddenCols.delete(c); }
      if (w != null) s.colW.set(newOther, w); else s.colW.delete(newOther); if (hid) s.hiddenCols.add(newOther); else s.hiddenCols.delete(newOther);
    }
  });
  const sel = axis === 'row' ? { ...rg, r1: rg.r1 + delta, r2: rg.r2 + delta } : { ...rg, c1: rg.c1 + delta, c2: rg.c2 + delta };
  app.grid.select(sel.r1, sel.c1, { range: isWholeRows(activeRange(app)) || isWholeCols(activeRange(app)) ? (axis === 'row' ? { r1: sel.r1, r2: sel.r2, c1: 0, c2: MAX_COLS - 1 } : { c1: sel.c1, c2: sel.c2, r1: 0, r2: MAX_ROWS - 1 }) : sel });
  app.grid.layout();
}

/** Reshapes a one-column (or one-row) list into a table of `width` columns, or a table into one column. */
export function reshape(values: Scalar[], width: number): Scalar[][] {
  const out: Scalar[][] = [];
  for (let i = 0; i < values.length; i += width) out.push(values.slice(i, i + width));
  if (out.length && out[out.length - 1].length < width) while (out[out.length - 1].length < width) out[out.length - 1].push(null);
  return out;
}

defTools([
  {
    id: 'util.insertRowsEvery', label: 'Insert Blank Rows Every Nth Row…', icon: 'playlist_add', tab: U, group: G_COLROW, edits: true, popular: true, tip: 'Insert one or more empty rows after every N rows of the selection', keywords: 'insert rows between every other', origin: 'ASAP Utilities / Kutools (Insert Blank Rows & Columns)',
    run: async (app, arg?: { every: number; count: number; startAfter?: number }) => {
      const o = await argOrAsk<{ every: number; count: number; startAfter?: number }>(arg, async () => { const r = await openDialog({ title: 'Insert Blank Rows Every Nth Row', fields: [{ key: 'every', label: 'After every … rows', type: 'number', value: 1, min: 1, half: true }, { key: 'count', label: 'Insert … blank rows', type: 'number', value: 1, min: 1, max: 100, half: true }] }); return r ? { every: Number(r.every) || 1, count: Number(r.count) || 1 } : null; });
      if (!o) return;
      insertEvery(app, 'row', o.every, o.count, o.startAfter ?? o.every);
    },
  },
  {
    id: 'util.insertColsEvery', label: 'Insert Blank Columns Every Nth Column…', icon: 'view_column', tab: U, group: G_COLROW, edits: true, origin: 'ASAP Utilities / Kutools',
    run: async (app, arg?: { every: number; count: number; startAfter?: number }) => {
      const o = await argOrAsk<{ every: number; count: number; startAfter?: number }>(arg, async () => { const r = await openDialog({ title: 'Insert Blank Columns Every Nth Column', fields: [{ key: 'every', label: 'After every … columns', type: 'number', value: 1, min: 1, half: true }, { key: 'count', label: 'Insert … blank columns', type: 'number', value: 1, min: 1, max: 50, half: true }] }); return r ? { every: Number(r.every) || 1, count: Number(r.count) || 1 } : null; });
      if (!o) return;
      insertEvery(app, 'col', o.every, o.count, o.startAfter ?? o.every);
    },
  },
  {
    id: 'util.insertRowAtChange', label: 'Insert a Blank Row Where a Value Changes…', icon: 'horizontal_rule', tab: U, group: G_COLROW, edits: true, popular: true, tip: 'Separate groups: insert an empty row (or a page break) each time the value in a column changes', keywords: 'group separator rows change of value page break', origin: 'ASAP Utilities / Kutools (Distinguish Differences)',
    run: async (app, arg?: { col: number; what: 'row' | 'break' | 'border'; count?: number }) => {
      if (emptySheet(app)) return;
      const s = app.sheet; const rg = selBlock(app); const header = headerGuess(s, rg);
      const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Insert Where a Value Changes', fields: [{ key: 'col', label: 'Watch the value in column', type: 'select', value: String(rg.c1), options: columnChoices(s, rg, header) }, { key: 'what', label: 'At each change insert', type: 'radio', value: 'row', options: [{ value: 'row', label: 'A blank row' }, { value: 'break', label: 'A page break' }, { value: 'border', label: 'A bottom border' }] }, { key: 'count', label: 'Rows to insert', type: 'number', value: 1, min: 1, max: 20, visibleIf: (v) => v.what === 'row' }] }); return r ? { col: Number(r.col), what: r.what, count: Number(r.count) || 1 } : null; });
      if (!o) return;
      const spots: number[] = [];
      for (let r = rg.r1 + (header ? 2 : 1); r <= rg.r2; r++) if (shownText(s, r, o.col) !== shownText(s, r - 1, o.col)) spots.push(r);
      if (!spots.length) { toast('The value never changes in that column.'); return; }
      if (o.what === 'break') { app.book.transact('Page Breaks', () => { app.book.recordSheet(s); s.rowBreaks = [...new Set([...s.rowBreaks, ...spots])].sort((a, b) => a - b); }); (app.grid as any).showPageBreaks = true; app.grid.invalidate(); toast(`${plural(spots.length, 'page break')} inserted`); return; }
      if (o.what === 'border') { app.book.transact('Borders at Changes', () => { for (const r of spots) for (let c = rg.c1; c <= rg.c2; c++) s.set(r - 1, c, { ...(s.get(r - 1, c) || {}), s: app.book.derivedStyle(s.styleIdOf(r - 1, c), { bb: { style: 'medium', color: '#333333' } }) }); }); return; }
      if (sheetProtected(app, 'insertRows')) return;
      app.book.transact('Insert Rows at Changes', () => { for (const at of spots.reverse()) insertRows(s, at, o.count || 1); });
      app.grid.updateExtent(); app.grid.layout();
      toast(`${plural(spots.length * (o.count || 1), 'row')} inserted`);
    },
  },
  {
    id: 'util.deleteEmptyRows', label: 'Delete Empty Rows', icon: 'delete_sweep', tab: U, group: G_COLROW, edits: true, popular: true, tip: 'Delete every row of the used range (or the selection) that is completely empty', keywords: 'remove blank rows', origin: 'ASAP Utilities / Kutools / Ablebits',
    run: async (app, arg?: { scope?: 'sheet' | 'selection' }) => {
      const s = app.sheet; const u = s.usedRange(false); if (!u) { toast('The sheet is empty.'); return; }
      const scope = arg?.scope || (singleCell(app) ? 'sheet' : 'selection');
      const rg = scope === 'sheet' ? u : clipToUsed(s, activeRange(app));
      const lines: number[] = []; for (let r = rg.r1; r <= rg.r2; r++) if (rowIsBlank(s, r)) lines.push(r);
      if (!lines.length) { toast('There are no empty rows.'); return; }
      if (sheetProtected(app, 'deleteRows')) return;
      app.book.transact('Delete Empty Rows', () => deleteRowSpans(s, mergeSpans(lines.map((l) => [l, l] as [number, number]))));
      app.grid.updateExtent(); app.grid.layout(); toast(`${plural(lines.length, 'empty row')} deleted`);
    },
  },
  {
    id: 'util.deleteEmptyCols', label: 'Delete Empty Columns', icon: 'delete_sweep', tab: U, group: G_COLROW, edits: true, origin: 'ASAP Utilities / Kutools / Ablebits',
    run: (app) => { const s = app.sheet; const u = s.usedRange(false); if (!u) { toast('The sheet is empty.'); return; } const rg = singleCell(app) ? u : clipToUsed(s, activeRange(app)); const lines: number[] = []; for (let c = rg.c1; c <= rg.c2; c++) if (colIsBlank(s, c)) lines.push(c); if (!lines.length) { toast('There are no empty columns.'); return; } if (sheetProtected(app, 'deleteColumns')) return; app.book.transact('Delete Empty Columns', () => deleteColSpans(s, mergeSpans(lines.map((l) => [l, l] as [number, number])))); app.grid.updateExtent(); app.grid.layout(); toast(`${plural(lines.length, 'empty column')} deleted`); },
  },
  { id: 'util.deleteHiddenRows', label: 'Delete Hidden Rows', icon: 'visibility_off', tab: U, group: G_COLROW, edits: true, tip: 'Remove the rows hidden by hand or by a filter (filtered-out rows), keeping only what is visible', origin: 'ASAP Utilities / Kutools (Delete Hidden Rows)', run: (app) => { const s = app.sheet; const u = s.usedRange(); const lines = [...s.hiddenRows].filter((r) => !u || r <= u.r2).sort((a, b) => a - b); if (!lines.length) { toast('No hidden rows.'); return; } if (sheetProtected(app, 'deleteRows')) return; app.book.transact('Delete Hidden Rows', () => { app.book.recordSheet(s); const spans = mergeSpans(lines.map((l) => [l, l] as [number, number])); for (const l of lines) s.hiddenRows.delete(l); deleteRowSpans(s, spans); s.hiddenRows.clear(); }); app.grid.updateExtent(); app.grid.layout(); toast(`${plural(lines.length, 'hidden row')} deleted`); } },
  { id: 'util.deleteHiddenCols', label: 'Delete Hidden Columns', icon: 'visibility_off', tab: U, group: G_COLROW, edits: true, origin: 'ASAP Utilities / Kutools', run: (app) => { const s = app.sheet; const u = s.usedRange(); const lines = [...s.hiddenCols].filter((c) => !u || c <= u.c2).sort((a, b) => a - b); if (!lines.length) { toast('No hidden columns.'); return; } if (sheetProtected(app, 'deleteColumns')) return; app.book.transact('Delete Hidden Columns', () => { app.book.recordSheet(s); const spans = mergeSpans(lines.map((l) => [l, l] as [number, number])); s.hiddenCols.clear(); deleteColSpans(s, spans); }); app.grid.updateExtent(); app.grid.layout(); toast(`${plural(lines.length, 'hidden column')} deleted`); } },
  { id: 'util.hideEmptyRows', label: 'Hide Empty Rows', icon: 'visibility_off', tab: U, group: G_COLROW, edits: true, origin: 'ASAP Utilities', run: (app) => { const s = app.sheet; const u = s.usedRange(false); if (!u) return; let n = 0; app.book.transact('Hide Empty Rows', () => { app.book.recordSheet(s); for (let r = u.r1; r <= u.r2; r++) if (rowIsBlank(s, r) && !s.hiddenRows.has(r)) { s.hiddenRows.add(r); n++; } }); app.grid.layout(); toast(`${plural(n, 'row')} hidden`); } },
  { id: 'util.hideEmptyCols', label: 'Hide Empty Columns', icon: 'visibility_off', tab: U, group: G_COLROW, edits: true, origin: 'ASAP Utilities', run: (app) => { const s = app.sheet; const u = s.usedRange(false); if (!u) return; let n = 0; app.book.transact('Hide Empty Columns', () => { app.book.recordSheet(s); for (let c = u.c1; c <= u.c2; c++) if (colIsBlank(s, c) && !s.hiddenCols.has(c)) { s.hiddenCols.add(c); n++; } }); app.grid.layout(); toast(`${plural(n, 'column')} hidden`); } },
  { id: 'util.unhideAll', label: 'Unhide All Rows and Columns', icon: 'visibility', tab: U, group: G_COLROW, edits: true, popular: true, origin: 'ASAP Utilities', run: (app) => { const s = app.sheet; const n = s.hiddenRows.size + s.hiddenCols.size; if (!n) { toast('Nothing is hidden.'); return; } app.book.transact('Unhide All', () => { app.book.recordSheet(s); s.hiddenRows.clear(); s.hiddenCols.clear(); }); app.grid.layout(); toast(`${n} hidden rows/columns shown`); } },
  {
    id: 'util.colWidthUnits', label: 'Set Column Width in cm, inches or points…', icon: 'width', tab: U, group: G_COLROW, edits: true, tip: 'Column widths of the selection in centimeters, inches, points or characters', origin: 'ASAP Utilities / Kutools (Adjust Cell Size)',
    run: async (app, arg?: { value: number; unit: string }) => {
      const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Column Width', fields: [{ key: 'value', label: 'Width', type: 'number', value: 2.5, min: 0, step: 0.1, half: true }, { key: 'unit', label: 'Unit', type: 'select', value: 'cm', options: ['cm', 'in', 'pt', 'chars', 'px'], half: true }] }); return r ? { value: Number(r.value), unit: r.unit } : null; });
      if (!o) return;
      if (sheetProtected(app, 'formatColumns')) return;
      const px = Math.max(0, Math.round(toPx(o.value, o.unit)));
      const s = app.sheet;
      app.book.transact('Column Width', () => { app.book.recordSheet(s); for (const rg of selRanges(app)) for (let c = rg.c1; c <= Math.min(rg.c2, rg.c1 + 2000); c++) { if (px) { s.colW.set(c, px); s.hiddenCols.delete(c); } else s.hiddenCols.add(c); } });
      app.grid.layout();
    },
  },
  {
    id: 'util.rowHeightUnits', label: 'Set Row Height in cm, inches or points…', icon: 'height', tab: U, group: G_COLROW, edits: true, origin: 'ASAP Utilities / Kutools',
    run: async (app, arg?: { value: number; unit: string }) => {
      const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Row Height', fields: [{ key: 'value', label: 'Height', type: 'number', value: 0.6, min: 0, step: 0.1, half: true }, { key: 'unit', label: 'Unit', type: 'select', value: 'cm', options: ['cm', 'in', 'pt', 'px'], half: true }] }); return r ? { value: Number(r.value), unit: r.unit } : null; });
      if (!o) return;
      if (sheetProtected(app, 'formatRows')) return;
      const px = Math.max(0, Math.round(toPx(o.value, o.unit)));
      const s = app.sheet;
      app.book.transact('Row Height', () => { app.book.recordSheet(s); for (const rg of selRanges(app)) for (let r = rg.r1; r <= Math.min(rg.r2, rg.r1 + 100000); r++) { if (px) { s.rowH.set(r, px); s.hiddenRows.delete(r); } else s.hiddenRows.add(r); } });
      app.grid.layout();
    },
  },
  { id: 'util.equalizeColWidths', label: 'Make Selected Columns the Same Width', icon: 'horizontal_distribute', tab: U, group: G_COLROW, edits: true, tip: 'Every selected column gets the width of the widest one', origin: 'ASAP Utilities', run: (app) => { if (sheetProtected(app, 'formatColumns')) return; const s = app.sheet; const cols = new Set<number>(); for (const rg of selRanges(app)) for (let c = rg.c1; c <= Math.min(rg.c2, rg.c1 + 2000); c++) cols.add(c); const w = Math.max(...[...cols].map((c) => s.colWidth(c) || s.defaultColW)); app.book.transact('Equal Column Widths', () => { app.book.recordSheet(s); for (const c of cols) s.colW.set(c, w); }); app.grid.layout(); } },
  { id: 'util.equalizeRowHeights', label: 'Make Selected Rows the Same Height', icon: 'vertical_distribute', tab: U, group: G_COLROW, edits: true, origin: 'ASAP Utilities', run: (app) => { if (sheetProtected(app, 'formatRows')) return; const s = app.sheet; const rows = new Set<number>(); for (const rg of selRanges(app)) for (let r = rg.r1; r <= Math.min(rg.r2, rg.r1 + 100000); r++) rows.add(r); const hgt = Math.max(...[...rows].map((r) => s.rowHeight(r) || s.defaultRowH)); app.book.transact('Equal Row Heights', () => { app.book.recordSheet(s); for (const r of rows) s.rowH.set(r, hgt); }); app.grid.layout(); } },
  { id: 'util.moveColLeft', label: 'Move Columns Left', icon: 'west', tab: U, group: G_COLROW, edits: true, keys: ['Mod-Alt-ArrowLeft'], tip: 'Swap the selected columns with the column to their left', origin: 'ASAP Utilities / Kutools', run: (app) => moveLines(app, 'col', -1) },
  { id: 'util.moveColRight', label: 'Move Columns Right', icon: 'east', tab: U, group: G_COLROW, edits: true, keys: ['Mod-Alt-ArrowRight'], run: (app) => moveLines(app, 'col', 1) },
  { id: 'util.moveRowUp', label: 'Move Rows Up', icon: 'north', tab: U, group: G_COLROW, edits: true, keys: ['Mod-Alt-ArrowUp'], run: (app) => moveLines(app, 'row', -1) },
  { id: 'util.moveRowDown', label: 'Move Rows Down', icon: 'south', tab: U, group: G_COLROW, edits: true, keys: ['Mod-Alt-ArrowDown'], run: (app) => moveLines(app, 'row', 1) },
  {
    id: 'util.reorderColumns', label: 'Reorder Columns…', icon: 'low_priority', tab: U, group: G_COLROW, edits: true, popular: true, tip: 'Put the columns of the selected table in the order you list (by header or letter); columns left out stay at the end', keywords: 'rearrange columns order move', origin: 'ASAP Utilities (Change the column order)',
    run: async (app, arg?: { order: string[] }) => {
      if (emptySheet(app)) return;
      const s = app.sheet; const rg = selBlock(app); const header = headerGuess(s, rg);
      const names = Array.from({ length: rg.c2 - rg.c1 + 1 }, (_x, j) => (header ? shownText(s, rg.r1, rg.c1 + j) : '') || colName(rg.c1 + j));
      const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Reorder Columns', help: 'Edit the list into the order you want, one column per line.', fields: [{ key: 'order', label: 'Columns', type: 'textarea', rows: Math.min(12, names.length + 1), value: names.join('\n') }] }); return r ? { order: String(r.order || '').split(/\r?\n/).map((x) => x.trim()).filter(Boolean) } : null; });
      if (!o) return;
      const idx: number[] = [];
      for (const n of o.order) { const i = names.findIndex((x, j) => x.toLowerCase() === n.toLowerCase() && !idx.includes(j)); if (i >= 0) idx.push(i); else { const ci = /^[A-Z]{1,3}$/i.test(n) ? names.findIndex((_x, j) => colName(rg.c1 + j) === n.toUpperCase()) : -1; if (ci >= 0 && !idx.includes(ci)) idx.push(ci); } }
      for (let j = 0; j < names.length; j++) if (!idx.includes(j)) idx.push(j);
      if (hasMergeIn(s, rg)) { toast('Unmerge the cells first.'); return; }
      if (!checkWritable(app, [rg])) return;
      app.book.transact('Reorder Columns', () => reorderCols(s, rg, idx));
    },
  },
  {
    id: 'util.sortColumnsByHeader', label: 'Sort Columns by Header', icon: 'sort_by_alpha', tab: U, group: G_COLROW, edits: true, tip: 'Rearrange the columns of the table alphabetically by their header text', origin: 'ASAP Utilities (Sort columns by header)',
    run: (app, arg?: { desc?: boolean }) => { if (emptySheet(app)) return; const s = app.sheet; const rg = selBlock(app); const names = Array.from({ length: rg.c2 - rg.c1 + 1 }, (_x, j) => shownText(s, rg.r1, rg.c1 + j)); const idx = names.map((_n, j) => j).sort((a, b) => naturalCompare(names[a], names[b]) || a - b); if (arg?.desc) idx.reverse(); if (hasMergeIn(s, rg)) { toast('Unmerge the cells first.'); return; } if (!checkWritable(app, [rg])) return; app.book.transact('Sort Columns', () => reorderCols(s, rg, idx)); },
  },
  {
    id: 'util.columnToTable', label: 'Reshape a Column into a Table…', icon: 'grid_on', tab: U, group: G_COLROW, edits: true, popular: true, tip: 'Turn one long column (name, street, city, name, street, city…) into rows of N columns, or a row into columns', keywords: 'reshape wrap list to table transform range', origin: 'ASAP Utilities / Kutools (Transform Range)',
    run: async (app, arg?: { width: number; dest?: 'right' | 'sheet' }) => {
      const s = app.sheet; const rg = clipToUsed(s, activeRange(app));
      const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Reshape a Column into a Table', fields: [{ key: 'width', label: 'Values per row', type: 'number', value: 3, min: 1, max: 1000 }, { key: 'dest', label: 'Put the table', type: 'radio', value: 'right', options: [{ value: 'right', label: 'To the right of the selection' }, { value: 'sheet', label: 'On a new sheet' }] }] }); return r ? { width: Number(r.width) || 1, dest: r.dest } : null; });
      if (!o) return;
      const vals: Scalar[] = []; for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) vals.push(s.value(r, c));
      const rows = reshape(vals, o.width);
      if (o.dest === 'sheet') { const { addSheetWithRows } = await import('./tools'); addSheetWithRows(app, 'Reshaped', rows); return; }
      const target = { r1: rg.r1, c1: rg.c2 + 2, r2: rg.r1 + rows.length - 1, c2: rg.c2 + 1 + o.width };
      if (!checkWritable(app, [target])) return;
      app.book.transact('Reshape', () => writeBlock(s, target.r1, target.c1, rows));
      app.grid.select(target.r1, target.c1, { range: target });
    },
  },
  {
    id: 'util.tableToColumn', label: 'Flatten a Range into One Column', icon: 'view_stream', tab: U, group: G_COLROW, edits: true, kind: 'split', tip: 'Write all values of the selection into one column (by rows or by columns) to the right of it, or on a new sheet', origin: 'ASAP Utilities / Kutools (Transform Range)',
    run: async (app, arg?: 'rows' | 'cols' | 'row') => {
      const s = app.sheet; const rg = clipToUsed(s, activeRange(app));
      const vals: Scalar[] = [];
      if (arg === 'cols') { for (let c = rg.c1; c <= rg.c2; c++) for (let r = rg.r1; r <= rg.r2; r++) vals.push(s.value(r, c)); }
      else for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) vals.push(s.value(r, c));
      const nonBlank = vals.filter((v) => v != null && v !== '');
      const target = arg === 'row' ? { r1: rg.r2 + 2, c1: rg.c1, r2: rg.r2 + 2, c2: rg.c1 + nonBlank.length - 1 } : { r1: rg.r1, c1: rg.c2 + 2, r2: rg.r1 + nonBlank.length - 1, c2: rg.c2 + 2 };
      if (!checkWritable(app, [target])) return;
      app.book.transact('Flatten', () => nonBlank.forEach((v, i) => writeOut(s, arg === 'row' ? target.r1 : target.r1 + i, arg === 'row' ? target.c1 + i : target.c1, v)));
      app.grid.select(target.r1, target.c1, { range: target });
    },
    menu: () => [{ id: 'util.tableToColumn', arg: 'rows', label: 'One column, row by row' }, { id: 'util.tableToColumn', arg: 'cols', label: 'One column, column by column' }, { id: 'util.tableToColumn', arg: 'row', label: 'One row (below the selection)' }],
  },
  {
    id: 'util.duplicateRows', label: 'Duplicate Rows…', icon: 'content_copy', tab: U, group: G_COLROW, edits: true, tip: 'Repeat each selected row N times, or as many times as a count column says (labels, tickets, packing lists)', keywords: 'repeat rows copies count', origin: 'ASAP Utilities / Kutools (Duplicate Rows)',
    run: async (app, arg?: { times?: number; col?: number }) => {
      if (emptySheet(app)) return;
      const s = app.sheet; const rg = selBlock(app); const header = headerGuess(s, rg);
      const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Duplicate Rows', fields: [{ key: 'mode', label: 'Repeat each row', type: 'radio', value: 'n', options: [{ value: 'n', label: 'This many times:' }, { value: 'col', label: 'As many times as the number in a column:' }] }, { key: 'times', label: 'Times', type: 'number', value: 2, min: 1, max: 1000, visibleIf: (v) => v.mode === 'n' }, { key: 'col', label: 'Count column', type: 'select', value: String(rg.c2), options: columnChoices(s, rg, header), visibleIf: (v) => v.mode === 'col' }] }); return r ? (r.mode === 'n' ? { times: Number(r.times) || 1 } : { col: Number(r.col) }) : null; });
      if (!o) return;
      if (sheetProtected(app, 'insertRows')) return;
      const first = rg.r1 + (header ? 1 : 0);
      const plan: { r: number; n: number }[] = [];
      for (let r = first; r <= rg.r2; r++) { const n = o.col != null ? Math.max(0, Math.floor(numOf(s.value(r, o.col)) ?? 1)) : o.times || 1; plan.push({ r, n }); }
      const total = plan.reduce((a, p) => a + Math.max(0, p.n - 1), 0);
      if (total > 200000) { toast('That would add too many rows.'); return; }
      app.book.transact('Duplicate Rows', () => {
        for (const p of plan.reverse()) {
          if (p.n <= 1) { if (p.n === 0) deleteRowSpans(s, [[p.r, p.r]]); continue; }
          insertRows(s, p.r + 1, p.n - 1);
          const row = blockOf(s, { r1: p.r, r2: p.r, c1: rg.c1, c2: rg.c2 })[0];
          for (let k = 1; k < p.n; k++) row.forEach((cell, j) => s.set(p.r + k, rg.c1 + j, movedCell(cell, k, 0)));
        }
      });
      app.grid.updateExtent(); app.grid.layout();
      toast(`${plural(total, 'row')} added`);
    },
  },
  {
    id: 'util.compactCells', label: 'Remove Blank Cells (Shift Up or Left)', icon: 'compress', tab: U, group: G_COLROW, edits: true, kind: 'split', tip: 'Close the gaps in each selected column (or row) by moving the non-blank cells up (or left)', origin: 'ASAP Utilities',
    run: (app, arg?: 'up' | 'left') => {
      const s = app.sheet; const rg = clipToUsed(s, activeRange(app));
      if (!checkWritable(app, [rg]) || hasMergeIn(s, rg)) { if (hasMergeIn(s, rg)) toast('Unmerge the cells first.'); return; }
      const m = blockOf(s, rg); let n = 0;
      app.book.transact('Remove Blank Cells', () => {
        if (arg === 'left') { for (let i = 0; i < m.length; i++) { const row = m[i].filter((c) => hasContent(c)); n += m[i].length - row.length; for (let j = 0; j < m[i].length; j++) s.set(rg.r1 + i, rg.c1 + j, j < row.length ? movedCell(row[j], 0, j - m[i].indexOf(row[j])) : undefined); } }
        else { for (let j = 0; j < m[0].length; j++) { const col = m.map((row) => row[j]); const keep = col.filter((c) => hasContent(c)); n += col.length - keep.length; for (let i = 0; i < col.length; i++) s.set(rg.r1 + i, rg.c1 + j, i < keep.length ? movedCell(keep[i], i - col.indexOf(keep[i]), 0) : undefined); } }
      });
      toast(`${plural(n, 'blank cell')} removed`);
    },
    menu: () => [{ id: 'util.compactCells', arg: 'up', label: 'Shift cells up' }, { id: 'util.compactCells', arg: 'left', label: 'Shift cells left' }],
  },
  { id: 'util.autofitAllSheets', label: 'AutoFit Columns on All Sheets', icon: 'fit_width', tab: U, group: G_COLROW, edits: true, origin: 'ASAP Utilities', run: (app) => { const cur = app.book.activeSheet; for (let i = 0; i < app.book.sheets.length; i++) { app.exec('sheet.activate', i); app.sheet.sel = { r: 0, c: 0, ranges: [{ r1: 0, c1: 0, r2: MAX_ROWS - 1, c2: MAX_COLS - 1 }] }; app.exec('format.autofitCols'); } app.exec('sheet.activate', cur); } },
  {
    id: 'util.joinRowsByKey', label: 'Combine Rows with the Same Key…', icon: 'merge_type', tab: U, group: G_COLROW, edits: true, popular: true, tip: 'Rows that share a value in the key column are merged into one: other columns are joined with a separator, summed, counted, or the first/last kept', keywords: 'consolidate duplicates combine rows group concatenate sum', origin: 'Kutools (Advanced Combine Rows) / Ablebits (Combine Rows)',
    run: async (app, arg?: { key: number; op: 'join' | 'sum' | 'first' | 'last' | 'count' | 'max' | 'min' | 'avg'; sep?: string }) => {
      if (emptySheet(app)) return;
      const s = app.sheet; const rg = selBlock(app); const header = headerGuess(s, rg);
      const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Combine Rows with the Same Key', fields: [{ key: 'key', label: 'Key column', type: 'select', value: String(rg.c1), options: columnChoices(s, rg, header) }, { key: 'op', label: 'Other columns', type: 'select', value: 'join', options: [{ value: 'join', label: 'Join the texts with a separator' }, { value: 'sum', label: 'Sum the numbers (join text)' }, { value: 'count', label: 'Count the rows' }, { value: 'avg', label: 'Average the numbers' }, { value: 'max', label: 'Largest value' }, { value: 'min', label: 'Smallest value' }, { value: 'first', label: 'Keep the first row\'s value' }, { value: 'last', label: 'Keep the last row\'s value' }] }, { key: 'sep', label: 'Separator', type: 'combo', value: ', ', options: [', ', '; ', ' ', ' | ', '\\n (line break)'], visibleIf: (v) => v.op === 'join' || v.op === 'sum' }] }); return r ? { key: Number(r.key), op: r.op, sep: String(r.sep).startsWith('\\n') ? '\n' : String(r.sep) } : null; });
      if (!o) return;
      if (!checkWritable(app, [rg]) || hasMergeIn(s, rg)) return;
      const first = rg.r1 + (header ? 1 : 0);
      const groups = new Map<string, number[]>();
      for (let r = first; r <= rg.r2; r++) { const k = shownText(s, r, o.key).toLowerCase(); (groups.get(k) || groups.set(k, []).get(k)!).push(r); }
      const out: Scalar[][] = [];
      for (const rows of groups.values()) {
        const row: Scalar[] = [];
        for (let c = rg.c1; c <= rg.c2; c++) {
          if (c === o.key) { row.push(s.value(rows[0], c)); continue; }
          const vals = rows.map((r) => s.value(r, c)).filter((v) => v != null && v !== '');
          const nums = vals.filter((v): v is number => typeof v === 'number');
          if (o.op === 'count') row.push(vals.length);
          else if (o.op === 'first') row.push(vals[0] ?? null);
          else if (o.op === 'last') row.push(vals[vals.length - 1] ?? null);
          else if (o.op === 'max') row.push(nums.length ? Math.max(...nums) : vals[0] ?? null);
          else if (o.op === 'min') row.push(nums.length ? Math.min(...nums) : vals[0] ?? null);
          else if (o.op === 'avg') row.push(nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : vals[0] ?? null);
          else if (o.op === 'sum' && nums.length === vals.length && vals.length) row.push(nums.reduce((a, b) => a + b, 0));
          else row.push(rows.map((r) => shownText(s, r, c)).filter((t) => t !== '').join(o.sep || ', '));
        }
        out.push(row);
      }
      app.book.transact('Combine Rows', () => { for (let r = first; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) { const cur = s.get(r, c); s.set(r, c, cur?.s ? { s: cur.s } : undefined); } writeBlock(s, first, rg.c1, out); /* the rows left over are cleared, not deleted: deleting sheet rows would take data beside the selection with them */ const spare: [number, number][] = []; for (let r = rg.r2; r >= first + out.length; r--) if (rowIsBlank(s, r)) spare.push([r, r]); if (spare.length && !app.sheet.protection.on) deleteRowSpans(s, spare); });
      app.grid.updateExtent(); app.grid.layout();
      toast(`${plural(rg.r2 - first + 1, 'row')} combined into ${out.length}`);
    },
  },
  { id: 'util.menuColRow', label: 'Columns & Rows', icon: 'table_rows', tab: U, group: G_COLROW, kind: 'menu', noRecord: true, run: () => undefined, menu: () => menuOf(['util.insertRowsEvery', 'util.insertColsEvery', 'util.insertRowAtChange', 'util.duplicateRows', null, 'util.deleteEmptyRows', 'util.deleteEmptyCols', 'util.deleteHiddenRows', 'util.deleteHiddenCols', 'util.selectEveryNth', 'util.condRows', 'tools.deleteBlankRows', 'util.compactCells', null, 'util.hideEmptyRows', 'util.hideEmptyCols', 'util.unhideAll', null, 'util.colWidthUnits', 'util.rowHeightUnits', 'util.equalizeColWidths', 'util.equalizeRowHeights', 'format.autofitCols', 'format.autofitRows', 'util.autofitAllSheets', null, 'util.moveColLeft', 'util.moveColRight', 'util.moveRowUp', 'util.moveRowDown', 'util.reorderColumns', 'util.sortColumnsByHeader', null, 'util.columnToTable', 'util.tableToColumn', 'util.joinRowsByKey', 'tools.mergeColumns', 'tools.splitColumn', 'tools.unpivot', 'tools.transposeInPlace', 'tools.reverse']) },
]);

// ============================================================================================
// Objects & Comments

function allComments(app: App, scope: 'sheet' | 'workbook'): { s: Sheet; k: number; cm: CellComment }[] {
  const out: { s: Sheet; k: number; cm: CellComment }[] = [];
  for (const s of scope === 'workbook' ? app.book.sheets : [app.sheet]) for (const [k, cm] of s.comments) out.push({ s, k, cm });
  return out;
}

function deleteObjects(app: App, kinds: ('chart' | 'image' | 'shape' | 'all')[], scope: 'sheet' | 'workbook', arg: any): void {
  const sheets = scope === 'workbook' ? app.book.sheets : [app.sheet];
  let n = 0;
  app.book.transact('Delete Objects', () => {
    for (const s of sheets) {
      if (sheetProtected(app, 'objects', s)) continue;
      app.book.recordSheet(s);
      if (kinds.includes('all') || kinds.includes('chart')) { n += s.charts.length; s.charts = []; }
      const keep = s.drawings.filter((d) => !(kinds.includes('all') || (kinds.includes('image') && d.kind === 'image') || (kinds.includes('shape') && d.kind !== 'image')));
      n += s.drawings.length - keep.length; s.drawings = keep;
    }
  });
  app.grid.objects.select(null); app.grid.invalidate();
  void arg;
  toast(n ? `${plural(n, 'object')} deleted` : 'No objects to delete.');
}

defTools([
  { id: 'util.deleteAllObjects', label: 'Delete All Objects on the Sheet', icon: 'delete', tab: U, group: G_OBJECTS, edits: true, tip: 'Remove every chart, picture, shape, text box and control from the active sheet', origin: 'ASAP Utilities / Kutools (Delete Illustrations & Objects)', run: (app, arg) => deleteObjects(app, ['all'], 'sheet', arg) },
  { id: 'util.deleteAllObjectsBook', label: 'Delete All Objects in the Workbook', icon: 'delete_forever', tab: U, group: G_OBJECTS, edits: true, origin: 'ASAP Utilities', run: (app, arg) => deleteObjects(app, ['all'], 'workbook', arg) },
  { id: 'util.deleteAllPictures', label: 'Delete All Pictures', icon: 'hide_image', tab: U, group: G_OBJECTS, edits: true, origin: 'ASAP Utilities', run: (app, arg) => deleteObjects(app, ['image'], 'sheet', arg) },
  { id: 'util.deleteAllShapes', label: 'Delete All Shapes and Text Boxes', icon: 'category', tab: U, group: G_OBJECTS, edits: true, origin: 'ASAP Utilities', run: (app, arg) => deleteObjects(app, ['shape'], 'sheet', arg) },
  { id: 'util.deleteAllCharts', label: 'Delete All Charts', icon: 'bar_chart', tab: U, group: G_OBJECTS, edits: true, origin: 'ASAP Utilities', run: (app, arg) => deleteObjects(app, ['chart'], 'sheet', arg) },
  {
    id: 'util.fitPicturesToCells', label: 'Fit Pictures to Their Cells', icon: 'fit_screen', tab: U, group: G_OBJECTS, edits: true, tip: 'Resize each picture on the sheet so it fits inside the cell it is anchored to', origin: 'ASAP Utilities / Kutools',
    run: (app) => { const s = app.sheet; if (sheetProtected(app, 'objects')) return; let n = 0; app.book.transact('Fit Pictures', () => { app.book.recordSheet(s); for (const d of s.drawings) { if (d.kind !== 'image') continue; const w = s.colWidth(d.at.c) || s.defaultColW, hgt = s.rowHeight(d.at.r) || s.defaultRowH; const k = d.lockAspect === false ? 1 : Math.min(w / Math.max(1, d.at.w), hgt / Math.max(1, d.at.h)); d.at = { ...d.at, dx: 0, dy: 0, w: d.lockAspect === false ? w : Math.max(1, Math.round(d.at.w * k)), h: d.lockAspect === false ? hgt : Math.max(1, Math.round(d.at.h * k)) }; n++; } }); app.grid.invalidate(); toast(n ? `${plural(n, 'picture')} fitted` : 'No pictures on this sheet.'); },
  },
  {
    id: 'util.objectsReport', label: 'List All Objects', icon: 'list', tab: U, group: G_OBJECTS, edits: true, tip: 'A sheet listing every chart, picture and shape in the workbook with its sheet, position and size', origin: 'ASAP Utilities',
    run: (app) => { const rows: Scalar[][] = []; for (const s of app.book.sheets) { for (const ch of s.charts) rows.push([s.name, 'Chart', ch.title || ch.type, cellName(ch.at.r, ch.at.c), ch.at.w, ch.at.h, ch.series.map((x) => x.values).join('; ')]); for (const d of s.drawings) rows.push([s.name, d.kind, d.alt || d.text || d.shape || '', cellName(d.at.r, d.at.c), d.at.w, d.at.h, d.src ? `${Math.round(d.src.length * 0.75 / 1024)} KB` : '']); } if (!rows.length) { toast('The workbook has no objects.'); return; } report(app, 'Objects', ['Sheet', 'Kind', 'Name', 'Cell', 'Width', 'Height', 'Details'], rows); },
  },
  {
    id: 'util.commentsToCells', label: 'Copy Comments into Cells', icon: 'comment', tab: U, group: G_OBJECTS, edits: true, kind: 'split', popular: true, tip: 'Write the comment or note of each selected cell into the cell to its right (or into the cell itself)', keywords: 'extract comments notes text', origin: 'ASAP Utilities / Kutools (Convert Comment & Cell)',
    run: (app, arg?: 'right' | 'replace') => { const s = app.sheet; let n = 0; app.book.transact('Comments to Cells', () => { for (const x of selCells(app, { includeFiltered: true })) { const cm = s.comments.get(cellKey(x.r, x.c)); if (!cm) continue; const text = [cm.text, ...(cm.replies || []).map((r) => r.text)].join('\n'); writeOut(s, x.r, arg === 'replace' ? x.c : x.c + 1, text); n++; } }); toast(n ? `${plural(n, 'comment')} copied` : 'No comments in the selection.'); },
    menu: () => [{ id: 'util.commentsToCells', arg: 'right', label: 'Into the cell to the right' }, { id: 'util.commentsToCells', arg: 'replace', label: 'Into the cell itself (replacing its value)' }],
  },
  {
    id: 'util.cellsToComments', label: 'Make Comments from Cell Values', icon: 'add_comment', tab: U, group: G_OBJECTS, edits: true, kind: 'split', tip: 'Attach each selected cell\'s text as a note to the cell to its left, or to the cell itself', origin: 'ASAP Utilities / Kutools',
    run: (app, arg?: 'left' | 'self' | 'append') => { const s = app.sheet; let n = 0; for (const x of selCells(app, { existingOnly: true })) { const t = shownText(s, x.r, x.c); if (!t) continue; const c = arg === 'left' ? Math.max(0, x.c - 1) : x.c; const cur = s.comments.get(cellKey(x.r, c)); const text = arg === 'append' && cur ? `${cur.text}\n${t}` : t; setComment(app.book, s, x.r, c, { author: app.prefs.author || 'Author', text, time: Date.now(), note: true }, 'Cells to Comments'); n++; } toast(`${plural(n, 'note')} written`); },
    menu: () => [{ id: 'util.cellsToComments', arg: 'self', label: 'As a note on the cell itself' }, { id: 'util.cellsToComments', arg: 'left', label: 'As a note on the cell to the left' }, { id: 'util.cellsToComments', arg: 'append', label: 'Append to the existing note' }],
  },
  {
    id: 'util.commentAuthor', label: 'Change the Author of All Comments…', icon: 'person_edit', tab: U, group: G_OBJECTS, edits: true, origin: 'ASAP Utilities / Kutools (Format Comment)',
    run: async (app, arg?: { author: string; scope?: 'sheet' | 'workbook' }) => { const o = await argOrAsk(arg, async () => { const a = await promptDialog('Comment Author', 'New author name for every comment and note', app.prefs.author || ''); return a == null ? null : { author: a, scope: 'workbook' as const }; }); if (!o) return; const list = allComments(app, o.scope || 'workbook'); app.book.transact('Comment Author', () => { for (const { s, k, cm } of list) { app.book.recordSheet(s); s.comments.set(k, { ...cm, author: o.author, replies: cm.replies?.map((r) => ({ ...r, author: o.author })) }); } }); toast(`${plural(list.length, 'comment')} updated`); },
  },
  {
    id: 'util.commentsFindReplace', label: 'Find and Replace in Comments…', icon: 'find_replace', tab: U, group: G_OBJECTS, edits: true, origin: 'Kutools (Find/Replace Comment Text)',
    run: async (app, arg?: { find: string; replace: string }) => { const o = await argOrAsk(arg, () => openDialog({ title: 'Find and Replace in Comments', fields: [{ key: 'find', label: 'Find', type: 'text', value: '', half: true }, { key: 'replace', label: 'Replace with', type: 'text', value: '', half: true }] })); if (!o?.find) return; const counts = findReplaceAll(app, { find: o.find, replace: o.replace || '', inValues: false, inComments: true, scope: 'workbook' }); toast(`${plural(counts.comments, 'comment')} changed`); },
  },
  { id: 'util.commentsResize', label: 'Resize All Comments to Fit', icon: 'aspect_ratio', tab: U, group: G_OBJECTS, edits: true, tip: 'Give every comment box a size that fits its text', origin: 'ASAP Utilities / Kutools (AutoFit Comment)', run: (app) => { const list = allComments(app, 'workbook'); app.book.transact('Resize Comments', () => { for (const { s, k, cm } of list) { app.book.recordSheet(s); const lines = cm.text.split('\n'); const w = Math.min(420, Math.max(120, Math.max(...lines.map((l) => l.length)) * 7 + 20)); const h = Math.min(400, Math.max(40, lines.reduce((a, l) => a + Math.ceil((l.length * 7 + 1) / w), 0) * 18 + 24)); s.comments.set(k, { ...cm, w, h }); } }); toast(`${plural(list.length, 'comment')} resized`); } },
  { id: 'util.resolveAllComments', label: 'Resolve All Comment Threads', icon: 'task_alt', tab: U, group: G_OBJECTS, edits: true, run: (app) => { const list = allComments(app, 'sheet').filter((x) => !x.cm.note && !x.cm.resolved); app.book.transact('Resolve Comments', () => { for (const { s, k, cm } of list) { app.book.recordSheet(s); s.comments.set(k, { ...cm, resolved: true }); } }); toast(`${plural(list.length, 'thread')} resolved`); } },
  { id: 'util.deleteCommentsBook', label: 'Delete All Comments in the Workbook', icon: 'comments_disabled', tab: U, group: G_OBJECTS, edits: true, origin: 'ASAP Utilities', run: async (app, arg) => { const list = allComments(app, 'workbook'); if (!list.length) { toast('No comments.'); return; } if (arg === undefined && !(await confirmDialog('Delete All Comments', `Delete ${plural(list.length, 'comment')} on every sheet?`, 'Delete'))) return; app.book.transact('Delete Comments', () => { for (const s of app.book.sheets) { app.book.recordSheet(s); s.comments.clear(); } }); app.grid.invalidate(); } },
  {
    id: 'util.commentsReport', label: 'List All Comments', icon: 'forum', tab: U, group: G_OBJECTS, edits: true, popular: true, tip: 'A sheet listing every comment and note with its sheet, cell, author, date and text', origin: 'ASAP Utilities / Kutools (Create Comment List)',
    run: (app) => { const rows: Scalar[][] = []; for (const { s, k, cm } of allComments(app, 'workbook')) { rows.push([s.name, cellName(keyRow(k), keyCol(k)), cm.note ? 'Note' : cm.resolved ? 'Comment (resolved)' : 'Comment', cm.author, cm.time ? new Date(cm.time).toISOString().slice(0, 16).replace('T', ' ') : '', cm.text, shownText(s, keyRow(k), keyCol(k))]); for (const rp of cm.replies || []) rows.push([s.name, cellName(keyRow(k), keyCol(k)), 'Reply', rp.author, rp.time ? new Date(rp.time).toISOString().slice(0, 16).replace('T', ' ') : '', rp.text, '']); } if (!rows.length) { toast('The workbook has no comments.'); return; } report(app, 'Comments', ['Sheet', 'Cell', 'Kind', 'Author', 'When', 'Text', 'Cell value'], rows, { colW: [100, 60, 110, 100, 120, 320, 140] }); },
  },
  { id: 'util.menuObjects', label: 'Objects & Comments', icon: 'comment', tab: U, group: G_OBJECTS, kind: 'menu', noRecord: true, run: () => undefined, menu: () => menuOf(['util.commentsToCells', 'util.cellsToComments', 'util.commentAuthor', 'util.commentsFindReplace', 'util.commentsResize', 'util.resolveAllComments', 'review.convertNotes', 'review.deleteAllComments', 'util.deleteCommentsBook', 'util.commentsReport', null, 'util.objectsReport', 'util.fitPicturesToCells', 'util.hideAllObjects', 'util.deleteAllPictures', 'util.deleteAllShapes', 'util.deleteAllCharts', 'util.deleteAllObjects', 'util.deleteAllObjectsBook']) },
]);

void parseRange; void CellError; void textOf; void selectHits; void pickSheets; void sheetsByNames; void parseTextDate; void valueRows;
