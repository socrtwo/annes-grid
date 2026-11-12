/**
 * Structural edits: insert/delete rows and columns, move ranges, rename sheets — with every reference
 * in the workbook (formulas, names, conditional formats, validation, tables, charts, merges, comments,
 * print areas, drawings) adjusted the way a spreadsheet does.
 */
import { cellKey, keyRow, keyCol, MAX_COLS, MAX_ROWS, parseRange, rangeName, splitSheetRef, quoteSheet, type Range } from './address';
import type { Sheet, Workbook } from './workbook';
import type { Cell } from './types';
import { mapRefs, shiftFormula, type RefToken } from '../calc/parser';

export type Axis = 'row' | 'col';

/** Adjusts a single coordinate range [a, b] for an insertion of n at index `at` (n < 0 deletes -n starting at `at`). */
function adjustSpan(a: number, b: number, at: number, n: number, max: number): [number, number] | null {
  if (n > 0) {
    if (a >= at) a += n;
    if (b >= at) b += n;
    if (a >= max) return null;
    return [a, Math.min(b, max - 1)];
  }
  const del = -n, end = at + del - 1;
  if (b < at) return [a, b];
  if (a > end) return [a - del, b - del];
  // overlap
  if (a >= at && b <= end) return null; // entirely deleted
  const na = a < at ? a : at;
  const nb = b > end ? b - del : at - 1;
  return nb < na ? null : [na, nb];
}

export function adjustRange(rg: Range, axis: Axis, at: number, n: number): Range | null {
  if (axis === 'row') {
    const s = adjustSpan(rg.r1, rg.r2, at, n, MAX_ROWS);
    return s ? { ...rg, r1: s[0], r2: s[1] } : null;
  }
  const s = adjustSpan(rg.c1, rg.c2, at, n, MAX_COLS);
  return s ? { ...rg, c1: s[0], c2: s[1] } : null;
}

/** Rewrites formula references to `target` sheet for an insert/delete, as seen from a formula on `home`. */
export function adjustFormula(f: string, home: Sheet, target: Sheet, axis: Axis, at: number, n: number): string {
  return mapRefs(f, (ref) => {
    const refSheet = ref.sheet ? home.book.sheetByName(ref.sheet) : home;
    if (refSheet !== target || ref.sheet2) return null;
    if ((axis === 'row' && ref.kind === 'cols') || (axis === 'col' && ref.kind === 'rows')) return null;
    // A range may be written backwards (A4:A2); it is adjusted in order and written back the way it was typed.
    const revR = ref.r1 > ref.r2, revC = ref.c1 > ref.c2;
    const rg = adjustRange({ r1: Math.min(ref.r1, ref.r2), c1: Math.min(ref.c1, ref.c2), r2: Math.max(ref.r1, ref.r2), c2: Math.max(ref.c1, ref.c2) }, axis, at, n);
    if (!rg) return { ...ref, invalid: true };
    const out = { r1: revR ? rg.r2 : rg.r1, r2: revR ? rg.r1 : rg.r2, c1: revC ? rg.c2 : rg.c1, c2: revC ? rg.c1 : rg.c2 };
    if (out.r1 === ref.r1 && out.r2 === ref.r2 && out.c1 === ref.c1 && out.c2 === ref.c2) return null;
    return { ...ref, ...out } as RefToken;
  });
}

/** Adjusts "Sheet!A1:B2" style reference text (names, chart series, print areas, validation lists). */
export function adjustRefText(text: string, home: Sheet, target: Sheet, axis: Axis, at: number, n: number): string {
  const lead = text.startsWith('=') ? '=' : '';
  return lead + adjustFormula(text.slice(lead.length), home, target, axis, at, n);
}

function remapKeyed<T>(m: Map<number, T>, axis: Axis, at: number, n: number): Map<number, T> {
  const out = new Map<number, T>();
  for (const [k, v] of m) {
    let r = keyRow(k), c = keyCol(k);
    if (axis === 'row') { if (r >= at) { if (n < 0 && r < at - n) continue; r += n; } if (r >= MAX_ROWS) continue; }
    else { if (c >= at) { if (n < 0 && c < at - n) continue; c += n; } if (c >= MAX_COLS) continue; }
    out.set(cellKey(r, c), v);
  }
  return out;
}

function remapIndex<T>(m: Map<number, T>, at: number, n: number, max: number): Map<number, T> {
  const out = new Map<number, T>();
  for (const [i, v] of m) {
    if (i < at) out.set(i, v);
    else if (n < 0 && i < at - n) continue;
    else if (i + n < max) out.set(i + n, v);
  }
  return out;
}

function remapSet(s: Set<number>, at: number, n: number, max: number): Set<number> {
  const out = new Set<number>();
  for (const i of s) {
    if (i < at) out.add(i);
    else if (n < 0 && i < at - n) continue;
    else if (i + n < max) out.add(i + n);
  }
  return out;
}

/**
 * Inserts (n > 0) or deletes (n < 0) whole rows or columns at index `at` of `sheet`.
 * Must be called inside book.transact().
 */
export function shiftCells(sheet: Sheet, axis: Axis, at: number, n: number): void {
  const book = sheet.book;
  book.recordBook();
  const max = axis === 'row' ? MAX_ROWS : MAX_COLS;
  // 1. Formulas everywhere that reference this sheet
  for (const sh of book.sheets) {
    for (const [k, cell] of sh.cells) {
      if (cell.f == null) continue;
      const nf = adjustFormula(cell.f, sh, sheet, axis, at, n);
      if (nf !== cell.f) sh.cells.set(k, { ...cell, f: nf });
    }
    // A rule's formula is written for the first cell of its range. When that cell is deleted but the range lives on, the
    // formula is first restated for the cell that becomes the new first cell, so it does not turn into #REF!.
    const rule = (f: string, ranges: Range[]): string => {
      const a = ranges[0];
      if (sh === sheet && n < 0 && a) {
        const lo = axis === 'row' ? a.r1 : a.c1, hi = axis === 'row' ? a.r2 : a.c2, end = at - n - 1;
        if (lo >= at && lo <= end && hi > end) f = axis === 'row' ? shiftFormula(f, end + 1 - lo, 0) : shiftFormula(f, 0, end + 1 - lo);
      }
      return adjustFormula(f, sh, sheet, axis, at, n);
    };
    for (const cf of sh.cf) cf.formulas = cf.formulas?.map((f) => rule(f, cf.ranges));
    for (const dv of sh.validations) {
      if (dv.formula1) dv.formula1 = rule(dv.formula1, dv.ranges);
      if (dv.formula2) dv.formula2 = rule(dv.formula2, dv.ranges);
    }
    // Reference text kept outside formulas follows the cells too.
    const refText = (t: string) => adjustRefText(t, sh, sheet, axis, at, n);
    const refList = (t: string) => t.split(',').map((x) => { const lead = /^\s*/.exec(x)![0]; return lead + refText(x.slice(lead.length)); }).join(',');
    for (const sc of sh.scenarios) if (sc.cells) sc.cells = refList(sc.cells);
    for (const er of sh.editRanges as { ref?: string }[]) if (typeof er.ref === 'string') er.ref = refList(er.ref);
    for (const d of sh.drawings) if (d.control) { if (d.control.link) d.control.link = refText(d.control.link); if (d.control.input) d.control.input = refText(d.control.input); }
    for (const ch of sh.charts) if (ch.sourceRange) ch.sourceRange = refText(ch.sourceRange);
    if (sh.print.titleRows) sh.print.titleRows = refText(sh.print.titleRows);
    if (sh.print.titleCols) sh.print.titleCols = refText(sh.print.titleCols);
    for (const ch of sh.charts) {
      for (const s of ch.series) {
        s.values = adjustRefText(s.values, sh, sheet, axis, at, n);
        if (s.categories) s.categories = adjustRefText(s.categories, sh, sheet, axis, at, n);
        if (s.nameRef) s.nameRef = adjustRefText(s.nameRef, sh, sheet, axis, at, n);
        if ((s as any).sizes) (s as any).sizes = adjustRefText((s as any).sizes, sh, sheet, axis, at, n);
      }
    }
    for (const sp of sh.sparklines) sp.data = adjustRefText(sp.data, sh, sheet, axis, at, n);
    for (const pv of sh.pivots) {
      if (/[!:]|^\$?[A-Z]{1,3}\$?\d/i.test(pv.source)) pv.source = adjustRefText(pv.source, sh, sheet, axis, at, n);
      if (pv.target.sheet === sheet.name) {
        const v = axis === 'row' ? pv.target.r : pv.target.c;
        const nv = v < at ? v : n < 0 && v < at - n ? at : v + n;
        if (axis === 'row') pv.target.r = nv; else pv.target.c = nv;
        // The area the report last filled moves with it; Refresh clears that area, so a stale one erases other data.
        if (pv.lastRange) pv.lastRange = adjustRange(pv.lastRange, axis, at, n) ?? undefined;
      }
    }
  }
  for (const nm of book.names) {
    const home = (nm.sheet && book.sheetByName(nm.sheet)) || book.sheets[0];
    nm.ref = adjustRefText(nm.ref, home, sheet, axis, at, n);
  }
  // 2. The sheet's own cells and per-row/col data
  sheet.cells = remapKeyed(sheet.cells, axis, at, n);
  sheet.comments = remapKeyed(sheet.comments, axis, at, n);
  if (axis === 'row') {
    sheet.rowH = remapIndex(sheet.rowH, at, n, max);
    sheet.rowLevel = remapIndex(sheet.rowLevel, at, n, max);
    sheet.rowStyle = remapIndex(sheet.rowStyle, at, n, max);
    sheet.hiddenRows = remapSet(sheet.hiddenRows, at, n, max);
    sheet.rowBreaks = [...remapSet(new Set(sheet.rowBreaks), at, n, max)];
    // Inserted rows inherit the formatting of the row above.
    if (n > 0 && at > 0) {
      const above = sheet.rowStyle.get(at - 1);
      const ah = sheet.rowH.get(at - 1);
      for (let i = 0; i < n; i++) {
        if (above != null) sheet.rowStyle.set(at + i, above);
        if (ah != null) sheet.rowH.set(at + i, ah);
      }
      for (const [k, cell] of [...sheet.cells]) {
        if (keyRow(k) !== at - 1 || !cell.s) continue;
        for (let i = 0; i < n; i++) sheet.cells.set(cellKey(at + i, keyCol(k)), { s: cell.s });
      }
    }
  } else {
    sheet.colW = remapIndex(sheet.colW, at, n, max);
    sheet.colLevel = remapIndex(sheet.colLevel, at, n, max);
    sheet.colStyle = remapIndex(sheet.colStyle, at, n, max);
    sheet.hiddenCols = remapSet(sheet.hiddenCols, at, n, max);
    sheet.colBreaks = [...remapSet(new Set(sheet.colBreaks), at, n, max)];
    if (n > 0 && at > 0) {
      const left = sheet.colStyle.get(at - 1);
      const lw = sheet.colW.get(at - 1);
      for (let i = 0; i < n; i++) {
        if (left != null) sheet.colStyle.set(at + i, left);
        if (lw != null) sheet.colW.set(at + i, lw);
      }
      for (const [k, cell] of [...sheet.cells]) {
        if (keyCol(k) !== at - 1 || !cell.s) continue;
        for (let i = 0; i < n; i++) sheet.cells.set(cellKey(keyRow(k), at + i), { s: cell.s });
      }
    }
  }
  // Spill caches are regenerated by recalc.
  for (const [k, cell] of [...sheet.cells]) if (cell.sp != null && cell.f == null) { if (cell.s) sheet.cells.set(k, { s: cell.s }); else sheet.cells.delete(k); }
  for (const [k, cell] of sheet.cells) if (cell.arr) { const a = adjustRange(cell.arr, axis, at, n); if (a) cell.arr = a; else delete cell.arr; void k; }
  const adj = (rg: Range) => adjustRange(rg, axis, at, n);
  sheet.merges = sheet.merges.map(adj).filter((x): x is Range => !!x && (x.r1 !== x.r2 || x.c1 !== x.c2));
  for (const cf of sheet.cf) cf.ranges = cf.ranges.map(adj).filter((x): x is Range => !!x);
  sheet.cf = sheet.cf.filter((x) => x.ranges.length);
  for (const dv of sheet.validations) dv.ranges = dv.ranges.map(adj).filter((x): x is Range => !!x);
  sheet.validations = sheet.validations.filter((x) => x.ranges.length);
  if (sheet.filter) { const r = adj(sheet.filter.range); if (r) { if (axis === 'col') remapFilterColumns(sheet.filter, sheet.filter.range, r, at, n); sheet.filter.range = r; } else sheet.filter = null; }
  sheet.tables = sheet.tables.filter((t) => { const r = adj(t.range); if (!r) return false; if (axis === 'col') { adjustTableColumns(t, at, n); if (t.filter) remapFilterColumns(t.filter, t.range, r, at, n); } t.range = r; return true; });
  const moveAnchor = (o: { at: { r: number; c: number } }) => {
    if (axis === 'row') { if (o.at.r >= at) o.at.r = Math.max(at, o.at.r + n); }
    else if (o.at.c >= at) o.at.c = Math.max(at, o.at.c + n);
  };
  sheet.charts.forEach(moveAnchor);
  sheet.drawings.forEach(moveAnchor);
  sheet.sparklines = sheet.sparklines.filter((sp) => {
    const v = axis === 'row' ? sp.cell.r : sp.cell.c;
    if (v < at) return true;
    if (n < 0 && v < at - n) return false;
    if (axis === 'row') sp.cell.r += n; else sp.cell.c += n;
    return true;
  });
  if (axis === 'row' && sheet.freeze.rows > at) sheet.freeze.rows = Math.max(at, sheet.freeze.rows + n);
  if (axis === 'col' && sheet.freeze.cols > at) sheet.freeze.cols = Math.max(at, sheet.freeze.cols + n);
  if (sheet.print.area) sheet.print.area = adjustRefText(sheet.print.area, sheet, sheet, axis, at, n);
  book.pending.structural = true;
}

/** Filter criteria and the sort column are stored by column offset within the list: they follow their columns. */
function remapFilterColumns(f: { columns: Record<number, unknown>; sort?: { col: number; desc: boolean } }, old: Range, now: Range, at: number, n: number): void {
  const move = (off: number): number | null => {
    const c = old.c1 + off;
    const nc = c < at ? c : n < 0 && c < at - n ? null : c + n;
    return nc == null || nc < now.c1 || nc > now.c2 ? null : nc - now.c1;
  };
  const cols: Record<number, unknown> = {};
  for (const [k, v] of Object.entries(f.columns)) { const o = move(Number(k)); if (o != null) cols[o] = v; }
  f.columns = cols;
  if (f.sort) { const o = move(f.sort.col); if (o == null) delete f.sort; else f.sort.col = o; }
}

function adjustTableColumns(t: { range: Range; columns: { name: string }[] }, at: number, n: number): void {
  const { c1, c2 } = t.range;
  if (n > 0) {
    // A column inserted at the table's first column or just past its last one moves or leaves the table; it does not join it.
    if (at <= c1 || at > c2) return;
    const names = new Set(t.columns.map((c) => c.name.toLowerCase()));
    const add = [];
    for (let i = 0; i < n; i++) { let k = t.columns.length + i + 1, nm = `Column${k}`; while (names.has(nm.toLowerCase())) nm = `Column${++k}`; names.add(nm.toLowerCase()); add.push({ name: nm }); }
    t.columns.splice(at - c1, 0, ...add);
  } else {
    // Remove exactly the table columns that fall inside the deleted span, wherever the span starts.
    const lo = Math.max(at, c1), hi = Math.min(at - n - 1, c2);
    if (lo <= hi) t.columns.splice(lo - c1, hi - lo + 1);
  }
}

/**
 * Inserts cells by shifting only a block (Insert ▸ Shift cells down/right), as opposed to whole rows/cols.
 * Formulas referencing the moved block move with it; other references are unchanged.
 */
export function insertCellsShift(sheet: Sheet, rg: Range, dir: 'down' | 'right' | 'up' | 'left'): void {
  const book = sheet.book;
  book.recordSheet(sheet);
  const used = sheet.usedRange() || rg;
  adjustStraddling(sheet, rg, dir);
  if (dir === 'down' || dir === 'up') {
    const h = rg.r2 - rg.r1 + 1;
    const block: Range = { r1: rg.r1, c1: rg.c1, r2: Math.max(used.r2, rg.r2), c2: rg.c2 };
    if (dir === 'down') moveBlock(sheet, block, h, 0);
    else { clearBlock(sheet, rg); moveBlock(sheet, { ...block, r1: rg.r2 + 1 }, -h, 0); }
  } else {
    const w = rg.c2 - rg.c1 + 1;
    const block: Range = { r1: rg.r1, c1: rg.c1, r2: rg.r2, c2: Math.max(used.c2, rg.c2) };
    if (dir === 'right') moveBlock(sheet, block, 0, w);
    else { clearBlock(sheet, rg); moveBlock(sheet, { ...block, c1: rg.c2 + 1 }, 0, -w); }
  }
}

/**
 * Before cells are shifted within a band of columns (or rows), ranges that only partly lie in the part that moves are
 * resized, and references to cells that are being deleted become #REF!. References wholly inside the part that moves
 * are left for moveBlock, which carries them along.
 */
function adjustStraddling(sheet: Sheet, rg: Range, dir: 'down' | 'right' | 'up' | 'left'): void {
  const book = sheet.book;
  const rows = dir === 'down' || dir === 'up';
  const size = rows ? rg.r2 - rg.r1 + 1 : rg.c2 - rg.c1 + 1;
  const at = rows ? rg.r1 : rg.c1, end = at + size - 1;
  const max = rows ? MAX_ROWS : MAX_COLS;
  const fix = (f: string, home: Sheet) => mapRefs(f, (ref) => {
    const refSheet = ref.sheet ? book.sheetByName(ref.sheet) : home;
    if (refSheet !== sheet || ref.sheet2 || ref.kind === 'cols' || ref.kind === 'rows') return null;
    const bandLo = rows ? Math.min(ref.c1, ref.c2) : Math.min(ref.r1, ref.r2), bandHi = rows ? Math.max(ref.c1, ref.c2) : Math.max(ref.r1, ref.r2);
    if (bandLo < (rows ? rg.c1 : rg.r1) || bandHi > (rows ? rg.c2 : rg.r2)) return null;
    const a = rows ? ref.r1 : ref.c1, b = rows ? ref.r2 : ref.c2;
    const rev = a > b;
    let lo = Math.min(a, b), hi = Math.max(a, b);
    if (dir === 'down' || dir === 'right') {
      if (!(lo < at && hi >= at)) return null;
      hi = Math.min(max - 1, hi + size);
    } else {
      if (hi < at || lo > end) return null;
      if (lo >= at && hi <= end) return { ...ref, invalid: true };
      if (lo < at) hi = hi > end ? hi - size : at - 1;
      else { lo = at; hi -= size; }
    }
    const p = rev ? [hi, lo] : [lo, hi];
    return (rows ? { ...ref, r1: p[0], r2: p[1] } : { ...ref, c1: p[0], c2: p[1] }) as RefToken;
  });
  book.recordBook();
  for (const sh of book.sheets) for (const [k, cell] of sh.cells) {
    if (cell.f == null) continue;
    const nf = fix(cell.f, sh);
    if (nf !== cell.f) sh.cells.set(k, { ...cell, f: nf });
  }
  for (const nm of book.names) {
    const home = (nm.sheet && book.sheetByName(nm.sheet)) || book.sheets[0];
    const lead = nm.ref.startsWith('=') ? '=' : '';
    nm.ref = lead + fix(nm.ref.slice(lead.length), home);
  }
}

function clearBlock(sheet: Sheet, rg: Range): void {
  for (const k of [...sheet.cells.keys()]) { const r = keyRow(k), c = keyCol(k); if (r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2) { sheet.book.recordCell(sheet, k); sheet.cells.delete(k); sheet.book.noteChange(sheet, k); } }
}

/**
 * Moves the cells of a block by (dr, dc) (cut/paste and drag-move semantics): the moved cells keep their
 * formulas' relative references pointing at the same cells (references INTO the block from anywhere follow
 * the block), and whatever was at the destination is overwritten.
 */
export function moveBlock(sheet: Sheet, src: Range, dr: number, dc: number, dest: Sheet = sheet): void {
  const book = sheet.book;
  // Formulas, names and objects on every sheet can point into the block, so the whole book is recorded.
  book.recordBook();
  const moved: [number, number, Cell][] = [];
  for (const [k, cell] of [...sheet.cells]) {
    const r = keyRow(k), c = keyCol(k);
    if (r >= src.r1 && r <= src.r2 && c >= src.c1 && c <= src.c2) { moved.push([r, c, cell]); sheet.cells.delete(k); }
  }
  const dst: Range = { r1: src.r1 + dr, c1: src.c1 + dc, r2: src.r2 + dr, c2: src.c2 + dc };
  if (dst.r1 < 0 || dst.c1 < 0 || dst.r2 >= MAX_ROWS || dst.c2 >= MAX_COLS) throw new Error('The move would place cells outside the sheet.');
  for (const k of [...dest.cells.keys()]) { const r = keyRow(k), c = keyCol(k); if (r >= dst.r1 && r <= dst.r2 && c >= dst.c1 && c <= dst.c2) dest.cells.delete(k); }
  // Moved formulas keep pointing at the same cells: their *relative* refs must be rewritten as if shifted back,
  // except refs into the moved block itself, which move along.
  for (const [r, c, cell] of moved) {
    let f = cell.f;
    if (f != null) f = rewriteMoved(f, sheet, dest, src, dr, dc, true);
    const next: Cell = { ...cell, ...(f != null ? { f } : {}) };
    if (next.sp != null) delete next.sp;
    dest.cells.set(cellKey(r + dr, c + dc), next);
  }
  // References into the block from all other formulas follow it.
  for (const sh of book.sheets) {
    for (const [k, cell] of sh.cells) {
      if (cell.f == null) continue;
      if (sh === dest && inRg(dst, keyRow(k), keyCol(k)) && moved.some(([r, c]) => cellKey(r + dr, c + dc) === k)) continue;
      const nf = rewriteMoved(cell.f, sh, dest, src, dr, dc, false, sheet);
      if (nf !== cell.f) sh.cells.set(k, { ...cell, f: nf });
    }
  }
  for (const nm of book.names) {
    const home = (nm.sheet && book.sheetByName(nm.sheet)) || book.sheets[0];
    const lead = nm.ref.startsWith('=') ? '=' : '';
    nm.ref = lead + rewriteMoved(nm.ref.slice(lead.length), home, dest, src, dr, dc, false, sheet);
  }
  // Charts, sparklines, PivotTable sources, rule formulas and print areas that point into the block follow it.
  const mv = (text: string, home: Sheet) => rewriteMoved(text, home, dest, src, dr, dc, false, sheet);
  for (const sh of book.sheets) {
    for (const ch of sh.charts) for (const se of ch.series) {
      se.values = mv(se.values, sh);
      if (se.categories) se.categories = mv(se.categories, sh);
      if (se.nameRef) se.nameRef = mv(se.nameRef, sh);
      if ((se as any).sizes) (se as any).sizes = mv((se as any).sizes, sh);
    }
    for (const sp of sh.sparklines) sp.data = mv(sp.data, sh);
    for (const pv of sh.pivots) if (/[!:]|^\$?[A-Z]{1,3}\$?\d/i.test(pv.source)) pv.source = mv(pv.source, sh);
    for (const cf of sh.cf) if (cf.formulas) cf.formulas = cf.formulas.map((f) => mv(f, sh));
    for (const dv of sh.validations) {
      if (dv.formula1) dv.formula1 = mv(dv.formula1, sh);
      if (dv.formula2) dv.formula2 = mv(dv.formula2, sh);
    }
    if (sh.print.area) sh.print.area = mv(sh.print.area, sh);
  }
  // Comments, merges, validation and formats inside the block move too.
  const movedComments: [number, any][] = [];
  for (const [k, cm] of [...sheet.comments]) if (inRg(src, keyRow(k), keyCol(k))) { movedComments.push([k, cm]); sheet.comments.delete(k); }
  for (const [k, cm] of movedComments) dest.comments.set(cellKey(keyRow(k) + dr, keyCol(k) + dc), cm);
  const mm = sheet.merges.filter((m) => m.r1 >= src.r1 && m.r2 <= src.r2 && m.c1 >= src.c1 && m.c2 <= src.c2);
  sheet.merges = sheet.merges.filter((m) => !mm.includes(m));
  dest.merges = dest.merges.filter((m) => !(m.r1 <= dst.r2 && m.r2 >= dst.r1 && m.c1 <= dst.c2 && m.c2 >= dst.c1));
  for (const m of mm) dest.merges.push({ r1: m.r1 + dr, c1: m.c1 + dc, r2: m.r2 + dr, c2: m.c2 + dc });
  book.pending.structural = true;
}

function inRg(rg: Range, r: number, c: number): boolean { return r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2; }

function rewriteMoved(f: string, home: Sheet, dest: Sheet, src: Range, dr: number, dc: number, isMovedCell: boolean, srcSheet: Sheet = home): string {
  return mapRefs(f, (ref) => {
    const refSheet = ref.sheet ? home.book.sheetByName(ref.sheet) : (isMovedCell ? srcSheet : home);
    const inside = refSheet === srcSheet && ref.r1 >= src.r1 && ref.r2 <= src.r2 && ref.c1 >= src.c1 && ref.c2 <= src.c2;
    const nr = { ...ref };
    if (inside) {
      nr.r1 += dr; nr.r2 += dr; nr.c1 += dc; nr.c2 += dc;
      if (dest !== srcSheet && (!isMovedCell || dest !== home)) nr.sheet = dest.name;
      if (isMovedCell && dest !== srcSheet && !ref.sheet) delete nr.sheet;
      return nr;
    }
    if (isMovedCell) {
      // A moved formula referring outside the block keeps pointing at the same cells.
      if (dest !== srcSheet && !ref.sheet) nr.sheet = srcSheet.name;
      return nr.sheet !== ref.sheet ? nr : null;
    }
    return null;
  });
}

/** Renames a sheet and updates every reference to it. Call inside book.transact(). */
export function renameSheet(book: Workbook, sheet: Sheet, newName: string): void {
  const old = sheet.name;
  if (old === newName) return;
  book.recordBook();
  const fix = (f: string, home: Sheet) => mapRefs(f, (ref) => {
    let changed = false;
    const nr = { ...ref };
    if (ref.sheet && ref.sheet.toLowerCase() === old.toLowerCase()) { nr.sheet = newName; changed = true; }
    if (ref.sheet2 && ref.sheet2.toLowerCase() === old.toLowerCase()) { nr.sheet2 = newName; changed = true; }
    void home;
    return changed ? nr : null;
  });
  // Catches sheet-qualified text the tokenizer leaves alone (3-D ranges, reference text); never inside "string literals".
  const re = new RegExp(`(^|[^\\w.'])(${escapeRe(quoteSheet(old))})!`, 'gi');
  const fixNames = (f: string) => f.split(/("(?:[^"]|"")*"?)/).map((part, i) => (i % 2 ? part : part.replace(re, (_m, pre) => `${pre}${quoteSheet(newName)}!`))).join('');
  for (const sh of book.sheets) {
    for (const [k, cell] of sh.cells) if (cell.f != null) { const nf = fixNames(fix(cell.f, sh)); if (nf !== cell.f) sh.cells.set(k, { ...cell, f: nf }); }
    for (const ch of sh.charts) for (const s of ch.series) { s.values = fixNames(s.values); if (s.categories) s.categories = fixNames(s.categories); if (s.nameRef) s.nameRef = fixNames(s.nameRef); }
    for (const sp of sh.sparklines) sp.data = fixNames(sp.data);
    for (const dv of sh.validations) { if (dv.formula1) dv.formula1 = fixNames(fix(dv.formula1, sh)); if (dv.formula2) dv.formula2 = fixNames(fix(dv.formula2, sh)); }
    for (const pv of sh.pivots) { pv.source = fixNames(pv.source); if (pv.target.sheet === old) pv.target.sheet = newName; }
    for (const cf of sh.cf) if (cf.formulas) cf.formulas = cf.formulas.map((f) => fixNames(fix(f, sh)));
    if (sh.print.area) sh.print.area = fixNames(sh.print.area);
  }
  for (const nm of book.names) { nm.ref = fixNames(fix(nm.ref.replace(/^=/, ''), sheet)); if (nm.sheet === old) nm.sheet = newName; }
  sheet.name = newName;
  book.pending.structural = true;
}

/** After deleting a sheet, references to it become #REF!. */
export function invalidateSheetRefs(book: Workbook, name: string): void {
  const low = name.toLowerCase();
  for (const sh of book.sheets) {
    for (const [k, cell] of sh.cells) {
      if (cell.f == null) continue;
      const nf = mapRefs(cell.f, (ref) => (ref.sheet && ref.sheet.toLowerCase() === low ? { ...ref, invalid: true } : null));
      if (nf !== cell.f) sh.cells.set(k, { ...cell, f: nf });
    }
  }
  book.names = book.names.filter((n) => !(n.sheet && n.sheet.toLowerCase() === low));
}

export function validSheetName(name: string): string | null {
  if (!name.trim()) return 'A sheet name cannot be blank.';
  if (name.length > 31) return 'A sheet name can have at most 31 characters.';
  if (/[\\/?*[\]:]/.test(name)) return 'A sheet name cannot contain \\ / ? * [ ] or :';
  if (/^'|'$/.test(name)) return 'A sheet name cannot begin or end with an apostrophe.';
  if (name.toLowerCase() === 'history') return '"History" is a reserved name.';
  return null;
}

function escapeRe(s: string): string { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

/** Parses "Sheet1!A1:B5" against a workbook. */
export function resolveRangeText(book: Workbook, text: string, fallback: Sheet): { sheet: Sheet; range: Range } | null {
  const { sheet: sn, ref } = splitSheetRef(text.replace(/^=/, ''));
  const sheet = sn ? book.sheetByName(sn) : fallback;
  if (!sheet) return null;
  const range = parseRange(ref);
  if (!range) {
    const nm = book.names.find((n) => n.name.toLowerCase() === ref.toLowerCase() && (!n.sheet || n.sheet === sheet.name));
    if (nm) return resolveRangeText(book, nm.ref, fallback);
    return null;
  }
  return { sheet, range };
}

export function rangeRefText(sheet: Sheet, rg: Range, abs = true, homeSheet?: Sheet): string {
  const body = abs ? rangeName(rg).replace(/([A-Z]+)(\d+)/g, '$$$1$$$2') : rangeName(rg);
  return (homeSheet && homeSheet === sheet ? '' : quoteSheet(sheet.name) + '!') + body;
}
