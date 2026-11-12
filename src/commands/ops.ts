/**
 * Shared editing helpers used by command modules: selection access, applying styles, entering values,
 * clearing, fill series, and conversions between cells and text.
 */
import type { App } from '../app';
import type { Sheet, Workbook } from '../model/workbook';
import type { Cell, CellStyle, Scalar, Matrix } from '../model/types';
import { CellError, errorFromCode } from '../model/types';
import { cellKey, keyRow, keyCol, MAX_COLS, MAX_ROWS, colName, cellName, type Range } from '../model/address';
import { parseNumberInput, isTextFormat, isDateFormat, formatValue } from '../model/numfmt';
import { parseDateTime, serialToParts, partsToSerial, daysInMonth } from '../model/dates';
import { shiftFormula } from '../calc/parser';
import { BUILTIN_LISTS } from '../prefs';
import { getFn } from '../calc/functions/registry';
import { lotusEntry, lotusLabel } from '../heritage/lotus';

export const selRanges = (app: App): Range[] => app.sheet.sel.ranges;
export const activeRange = (app: App): Range => { const s = app.sheet.sel; return s.ranges[s.ranges.length - 1]; };
export const activeCell = (app: App): { r: number; c: number } => ({ r: app.sheet.sel.r, c: app.sheet.sel.c });

export function isWholeCols(rg: Range): boolean { return rg.r1 === 0 && rg.r2 >= MAX_ROWS - 1; }
export function isWholeRows(rg: Range): boolean { return rg.c1 === 0 && rg.c2 >= MAX_COLS - 1; }

/** Clips huge selections (whole rows/columns, select-all) to the used range for per-cell work. */
export function clipToUsed(sheet: Sheet, rg: Range, pad = 0): Range {
  const u = sheet.usedRange() || { r1: 0, c1: 0, r2: 0, c2: 0 };
  return { r1: rg.r1, c1: rg.c1, r2: Math.min(rg.r2, Math.max(u.r2 + pad, rg.r1)), c2: Math.min(rg.c2, Math.max(u.c2 + pad, rg.c1)) };
}

/** Calls fn for each cell position in the selection (clipped to the used range). */
export function forEachSelected(app: App, fn: (r: number, c: number) => void, opts: { clip?: boolean } = {}): void {
  const sheet = app.sheet;
  const seen = new Set<number>();
  for (const rg0 of selRanges(app)) {
    const rg = opts.clip === false ? rg0 : clipToUsed(sheet, rg0);
    for (let r = rg.r1; r <= rg.r2; r++) {
      if (sheet.hiddenRows.has(r) && sheet.filter) continue;
      for (let c = rg.c1; c <= rg.c2; c++) {
        const k = cellKey(r, c);
        if (seen.has(k)) continue;
        seen.add(k);
        fn(r, c);
      }
    }
  }
}

/** Style of the active cell (used for toggle states). */
export function activeStyle(app: App): CellStyle {
  const { r, c } = activeCell(app);
  return app.sheet.styleOf(r, c);
}

/**
 * Applies a style patch to the selection. Whole rows/columns set row/column styles (so new cells inherit),
 * and existing cells in them are patched too.
 */
export function applyStyle(app: App, patch: Partial<CellStyle> | ((st: CellStyle, r: number, c: number) => Partial<CellStyle>), label = 'Format'): void {
  const book = app.book;
  const sheet = app.sheet;
  book.transact(label, () => {
    for (const rg of selRanges(app)) {
      if (isWholeCols(rg) || isWholeRows(rg)) {
        book.recordSheet(sheet);
        if (isWholeCols(rg) && isWholeRows(rg)) {
          // whole sheet: style every column and every existing cell
          for (let c = 0; c <= Math.max(rg.c2 >= MAX_COLS - 1 ? (sheet.usedRange()?.c2 ?? 0) + 50 : rg.c2, 0); c++) {
            const p = typeof patch === 'function' ? patch(book.styles[sheet.colStyle.get(c) ?? 0] || {}, -1, c) : patch;
            sheet.colStyle.set(c, book.derivedStyle(sheet.colStyle.get(c) ?? 0, p));
          }
        } else if (isWholeCols(rg)) {
          for (let c = rg.c1; c <= rg.c2; c++) {
            const p = typeof patch === 'function' ? patch(book.styles[sheet.colStyle.get(c) ?? 0] || {}, -1, c) : patch;
            sheet.colStyle.set(c, book.derivedStyle(sheet.colStyle.get(c) ?? 0, p));
          }
          for (const [r, sid] of sheet.rowStyle) for (let c = rg.c1; c <= rg.c2; c++) if (!sheet.get(r, c)) {
            const p = typeof patch === 'function' ? patch(book.styles[sid] || {}, r, c) : patch;
            sheet.cells.set(cellKey(r, c), { s: book.derivedStyle(sid, p) });
          }
        } else {
          for (let r = rg.r1; r <= rg.r2; r++) {
            const p = typeof patch === 'function' ? patch(book.styles[sheet.rowStyle.get(r) ?? 0] || {}, r, -1) : patch;
            sheet.rowStyle.set(r, book.derivedStyle(sheet.rowStyle.get(r) ?? 0, p));
          }
        }
        for (const [k, cell] of [...sheet.cells]) {
          const r = keyRow(k), c = keyCol(k);
          if (r < rg.r1 || r > rg.r2 || c < rg.c1 || c > rg.c2) continue;
          const base = cell.s ?? 0;
          const p = typeof patch === 'function' ? patch(book.styles[base] || {}, r, c) : patch;
          sheet.cells.set(k, { ...cell, s: book.derivedStyle(base, p) });
        }
      } else {
        for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) {
          if (filteredOut(sheet, r)) continue;
          const base = sheet.styleIdOf(r, c);
          const p = typeof patch === 'function' ? patch(book.styles[base] || {}, r, c) : patch;
          const ns = book.derivedStyle(base, p);
          const cur = sheet.get(r, c);
          if (ns === (cur?.s ?? 0) && cur) continue;
          sheet.patch(r, c, { s: ns || undefined });
        }
      }
    }
  });
}

export function toggleStyle(app: App, key: keyof CellStyle, label: string): void {
  const on = !activeStyle(app)[key];
  applyStyle(app, { [key]: on || undefined } as Partial<CellStyle>, label);
}

// --------------------------------------------------------------------------------------------
// Entering values

export interface ParsedInput { cell: Cell; format?: string; /** Alignment from a 1-2-3 label prefix. */ align?: 'left' | 'right' | 'center' | 'fill' }

/**
 * Converts typed text to a cell: formulas, numbers (with currency/percent/thousands), dates and times,
 * booleans, errors, and text (a leading apostrophe forces text).
 */
export function parseInput(app: App, text: string, r: number, c: number, sheet: Sheet = app.sheet): ParsedInput {
  const st = sheet.styleOf(r, c);
  if (isTextFormat(st.numFmt)) return { cell: text === '' ? {} : { v: text } };
  if (text === '') return { cell: {} };
  if (text.startsWith("'")) return { cell: { v: text.slice(1) } };
  if (app.prefs.lotusLabels) {
    const lb = lotusLabel(text);
    if (lb && lb.align !== 'left') return { cell: { v: lb.text }, align: lb.align };
  }
  if (app.prefs.lotusFormulas && /^[@+=-]/.test(text)) {
    const lf = lotusEntry(text, (n) => !!getFn(n), { sheetNames: sheet.book.sheets.map((x) => x.name) });
    if (lf) return { cell: { f: lf } };
  }
  if ((text.startsWith('=') || ((text.startsWith('+') || text.startsWith('-')) && text.length > 1 && !parseNumberInput(text))) && text.length > 1) {
    let f = text.startsWith('=') ? text.slice(1) : text;
    if (f.startsWith('=')) f = f.slice(1);
    return { cell: { f: f.trim() } };
  }
  if (/^(true|false)$/i.test(text.trim())) return { cell: { v: text.trim().toUpperCase() === 'TRUE' } };
  const err = errorFromCode(text.trim());
  if (err) return { cell: { v: err } };
  const n = parseNumberInput(text);
  if (n) {
    let value = n.value;
    // "Enable automatic percent entry": typing 5 into a % formatted cell means 5%.
    if (app.prefs.autoPercent && !n.format && st.numFmt && /%/.test(st.numFmt) && !/%/.test(text)) value /= 100;
    if (app.prefs.fixedDecimals && !/[.,%]/.test(text) && !n.format) value = value / Math.pow(10, app.prefs.fixedDecimals);
    return { cell: { v: value }, format: st.numFmt && st.numFmt !== 'General' ? undefined : n.format };
  }
  const d = parseDateTime(text, { dmy: app.prefs.dateOrder === 'dmy', date1904: sheet.book.calc.date1904 });
  if (d) return { cell: { v: d.serial }, format: st.numFmt && isDateFormat(st.numFmt) ? undefined : d.format };
  return { cell: { v: text } };
}

/** Writes parsed input into a cell, keeping its style and applying an inferred number format. */
export function setInput(app: App, sheet: Sheet, r: number, c: number, text: string): void {
  const p = parseInput(app, text, r, c, sheet);
  const cur = sheet.get(r, c);
  const next: Cell = { ...p.cell };
  let sid = cur?.s ?? sheet.rowStyle.get(r) ?? sheet.colStyle.get(c);
  if (p.format) sid = sheet.book.derivedStyle(sid ?? 0, { numFmt: p.format });
  if (p.align) sid = sheet.book.derivedStyle(sid ?? 0, { hAlign: p.align });
  // Typing a hyperlink-looking address makes a hyperlink (AutoCorrect: internet paths).
  if (typeof next.v === 'string' && app.prefs.autoCorrect && /^(https?:\/\/|www\.|mailto:)[^\s]+$/i.test(next.v)) {
    next.link = /^www\./i.test(next.v) ? 'https://' + next.v : next.v;
    sid = sheet.book.derivedStyle(sid ?? 0, { color: '#0563c1', underline: 'single' });
  } else if (cur?.link && next.v != null) next.link = cur.link;
  if (sid) next.s = sid;
  sheet.set(r, c, next);
}

// --------------------------------------------------------------------------------------------
// Clearing

/** True when the row is hidden by a filter (not merely hidden by hand): commands on a filtered list leave such rows alone. */
export function filteredOut(sheet: Sheet, r: number): boolean {
  return sheet.hiddenRows.has(r) && (!!sheet.filter || sheet.tables.some((t) => !!t.filter && r >= t.range.r1 && r <= t.range.r2));
}

export type ClearWhat = 'all' | 'contents' | 'formats' | 'comments' | 'links' | 'linksAndFormats';

export function clearRange(app: App, sheet: Sheet, rg: Range, what: ClearWhat): void {
  const book = sheet.book;
  const area = clipToUsed(sheet, rg);
  const keys: number[] = [];
  sheet.forEachIn(area, (r, c) => keys.push(cellKey(r, c)));
  for (const k of keys) {
    const r = keyRow(k), c = keyCol(k);
    const cell = sheet.cells.get(k);
    if (!cell) continue;
    if (filteredOut(sheet, r)) continue;
    if (cell.sp != null && cell.f == null && (what === 'contents' || what === 'all')) continue; // spilled values belong to the anchor
    switch (what) {
      case 'all': sheet.set(r, c, undefined); break;
      case 'contents': sheet.set(r, c, cell.s ? { s: cell.s } : undefined); break;
      case 'formats': { const { s: _s, ...rest } = cell; void _s; sheet.set(r, c, rest); break; }
      case 'links': { const { link: _l, ...rest } = cell; void _l; sheet.set(r, c, rest); break; }
      case 'linksAndFormats': { const { link: _l, s: _s, ...rest } = cell; void _l; void _s; sheet.set(r, c, rest); break; }
    }
  }
  if (what === 'all' || what === 'comments') {
    for (const k of [...sheet.comments.keys()]) { const r = keyRow(k), c = keyCol(k); if (r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2) { book.recordSheet(sheet); sheet.comments.delete(k); } }
  }
  if (what === 'all' || what === 'formats') {
    if (isWholeCols(rg) || isWholeRows(rg)) {
      book.recordSheet(sheet);
      if (isWholeCols(rg)) for (let c = rg.c1; c <= Math.min(rg.c2, MAX_COLS - 1); c++) sheet.colStyle.delete(c);
      if (isWholeRows(rg)) for (let r = rg.r1; r <= Math.min(rg.r2, MAX_ROWS - 1); r++) sheet.rowStyle.delete(r);
    }
    if (what === 'all') {
      const inside = (m: Range) => m.r1 >= rg.r1 && m.r2 <= rg.r2 && m.c1 >= rg.c1 && m.c2 <= rg.c2;
      if (sheet.merges.some(inside)) { book.recordSheet(sheet); sheet.merges = sheet.merges.filter((m) => !inside(m)); }
    }
  }
}

// --------------------------------------------------------------------------------------------
// Fill (fill handle, Fill Down/Right, Series)

/** Detects a series in source values and returns a generator for the n-th next value. */
export function seriesFrom(app: App, vals: { cell: Cell | undefined; r: number; c: number }[], copy: boolean): (i: number) => { cell: Cell | undefined; srcIndex: number } {
  const n = vals.length;
  const cells = vals.map((v) => v.cell);
  const nums = cells.map((c) => (c && c.f == null && typeof c.v === 'number' ? c.v : null));
  const allNums = nums.every((x) => x != null) && n > 0;
  // Formulas, or copy mode, repeat the source pattern (formulas shift relative refs).
  if (copy || cells.some((c) => c?.f != null) || (!allNums && !textSeriesInfo(app, cells))) {
    if (allNums && n === 1 && !copy) {
      // single number: copy (Excel copies single numbers; Ctrl+drag makes a series)
    }
    return (i) => ({ cell: cells[((i % n) + n) % n], srcIndex: ((i % n) + n) % n });
  }
  if (allNums) {
    if (n === 1) return (i) => ({ cell: copy ? cells[0] : cells[0], srcIndex: 0 });
    // Dates a whole number of months apart (the 1st of each month, month ends, the same day each year) continue by
    // the calendar: a straight line through them drifts, because months differ in length.
    const monthly = monthSeries(app, cells, nums as number[]);
    if (monthly) return (i) => { const src = ((i % n) + n) % n; return { cell: { ...(cells[src] || {}), v: monthly(n + i) }, srcIndex: src }; };
    // linear trend (best fit) for numeric series
    const xs = nums.map((_, i) => i), ys = nums as number[];
    const mx = (n - 1) / 2, my = ys.reduce((a, b) => a + b, 0) / n;
    let sxy = 0, sxx = 0;
    xs.forEach((x, i) => { sxy += (x - mx) * (ys[i] - my); sxx += (x - mx) ** 2; });
    const slope = sxx ? sxy / sxx : 0, icpt = my - slope * mx;
    return (i) => {
      const src = ((i % n) + n) % n;
      return { cell: { ...(cells[src] || {}), v: +(icpt + slope * (n + i)).toPrecision(15) }, srcIndex: src };
    };
  }
  const info = textSeriesInfo(app, cells)!;
  return (i) => {
    const src = ((i % n) + n) % n;
    return { cell: { ...(cells[src] || {}), v: info.next(n + i) }, srcIndex: src };
  };
}

/** For dates a constant number of months apart, on the same day of the month or all on month ends: the k-th date of the series. */
function monthSeries(app: App, cells: (Cell | undefined)[], nums: number[]): ((k: number) => number) | null {
  const d1904 = app.book.calc.date1904;
  if (nums.length < 2 || !cells.every((c) => isDateFormat(app.book.styles[c?.s ?? 0]?.numFmt))) return null;
  if (nums.some((v) => v < 61 || v > 2958465)) return null;
  const parts = nums.map((v) => serialToParts(Math.floor(v), d1904));
  const idx = parts.map((p) => p.y * 12 + (p.m - 1));
  const step = idx[1] - idx[0];
  if (step === 0 || idx.some((x, i) => i > 0 && x - idx[i - 1] !== step)) return null;
  const ends = parts.every((p) => p.d === daysInMonth(p.y, p.m));
  if (!ends && parts.some((p) => p.d !== parts[0].d)) return null;
  const frac = nums[0] - Math.floor(nums[0]);
  return (k) => {
    const m = idx[0] + step * k, y = Math.floor(m / 12), mo = m - y * 12 + 1;
    if (y < 1900 || y > 9999) return nums[nums.length - 1];
    const dim = daysInMonth(y, mo);
    return partsToSerial(y, mo, ends ? dim : Math.min(parts[0].d, dim), d1904) + frac;
  };
}

/** Text series: custom lists (days, months, quarters) and text with a trailing number (Item 1, Item 2). */
function textSeriesInfo(app: App, cells: (Cell | undefined)[]): { next: (i: number) => string } | null {
  const strs = cells.map((c) => (c && typeof c.v === 'string' && c.f == null ? c.v : null));
  if (strs.some((s) => s == null)) return null;
  const lists = [...BUILTIN_LISTS, ...app.prefs.customLists];
  for (const l of lists) {
    const low = l.items.map((x) => x.toLowerCase());
    const idx = strs.map((s) => low.indexOf(s!.toLowerCase()));
    if (idx.every((x) => x >= 0)) {
      const step = idx.length > 1 ? idx[1] - idx[0] : 1;
      const caseOf = (s: string, t: string) => (s === s.toUpperCase() ? t.toUpperCase() : s === s.toLowerCase() ? t.toLowerCase() : t);
      return { next: (i) => caseOf(strs[0]!, l.items[(((idx[0] + step * i) % l.items.length) + l.items.length) % l.items.length]) };
    }
  }
  const m = strs.map((s) => /^(.*?)(\d+)(\D*)$/.exec(s!));
  if (m.every((x) => x) && m.every((x) => x![1] === m[0]![1] && x![3] === m[0]![3])) {
    const ns = m.map((x) => +x![2]);
    const step = ns.length > 1 ? ns[1] - ns[0] : 1;
    const width = m[0]![2].length;
    return { next: (i) => `${m[0]![1]}${String(Math.max(0, ns[0] + step * i)).padStart(m[0]![2].startsWith('0') ? width : 0, '0')}${m[0]![3]}` };
  }
  return null;
}

/** Fills target from src (fill handle semantics). Must be called inside a transaction. */
export function fillRange(app: App, sheet: Sheet, src: Range, target: Range, copy = false): void {
  const down = target.r2 > src.r2, up = target.r1 < src.r1, right = target.c2 > src.c2, left = target.c1 < src.c1;
  if (!down && !up && !right && !left) {
    // shrinking: clear the part of src outside target
    for (let r = src.r1; r <= src.r2; r++) for (let c = src.c1; c <= src.c2; c++) if (r > target.r2 || c > target.c2) sheet.set(r, c, undefined);
    return;
  }
  const vertical = down || up;
  const lines = vertical ? range(src.c1, src.c2) : range(src.r1, src.r2);
  for (const line of lines) {
    const vals = vertical ? range(src.r1, src.r2).map((r) => ({ cell: sheet.get(r, line), r, c: line })) : range(src.c1, src.c2).map((c) => ({ cell: sheet.get(line, c), r: line, c }));
    const dirSign = down || right ? 1 : -1;
    const gen = seriesFrom(app, dirSign > 0 ? vals : [...vals].reverse(), copy);
    const n = vals.length;
    const count = vertical ? (down ? target.r2 - src.r2 : src.r1 - target.r1) : (right ? target.c2 - src.c2 : src.c1 - target.c1);
    for (let i = 0; i < count; i++) {
      const { cell, srcIndex } = gen(i);
      const r = vertical ? (down ? src.r2 + 1 + i : src.r1 - 1 - i) : line;
      const c = vertical ? line : (right ? src.c2 + 1 + i : src.c1 - 1 - i);
      if (sheet.hiddenRows.has(r) && sheet.filter) continue;
      const srcPos = dirSign > 0 ? vals[srcIndex] : [...vals].reverse()[srcIndex];
      let next: Cell | undefined = cell ? { ...cell } : undefined;
      if (next) { delete next.sp; delete next.arr; }
      if (next?.f != null) next.f = shiftFormula(next.f, r - srcPos.r, c - srcPos.c);
      if (next && next.f == null && next.v == null && !next.s) next = undefined;
      sheet.set(r, c, next);
    }
    void n;
  }
}

function range(a: number, b: number): number[] { const out: number[] = []; for (let i = a; i <= b; i++) out.push(i); return out; }

// --------------------------------------------------------------------------------------------
// Text conversions (clipboard, export)

export function cellText(sheet: Sheet, r: number, c: number, formatted = true): string {
  const cell = sheet.get(r, c);
  if (!cell) return '';
  const v = cell.v ?? null;
  if (!formatted) return v == null ? '' : v instanceof CellError ? v.code : String(v);
  return formatValue(v, sheet.styleOf(r, c).numFmt, { date1904: sheet.book.calc.date1904 }).text;
}

export function rangeToTsv(sheet: Sheet, rg: Range): string {
  const r2 = Math.min(rg.r2, (sheet.usedRange()?.r2 ?? rg.r1)), c2 = Math.min(rg.c2, (sheet.usedRange()?.c2 ?? rg.c1));
  const lines: string[] = [];
  for (let r = rg.r1; r <= Math.max(r2, rg.r1); r++) {
    if (sheet.hiddenRows.has(r) && sheet.filter) continue;
    const row: string[] = [];
    for (let c = rg.c1; c <= Math.max(c2, rg.c1); c++) {
      const t = cellText(sheet, r, c);
      row.push(/[\t\n"]/.test(t) ? `"${t.replace(/"/g, '""')}"` : t);
    }
    lines.push(row.join('\t'));
  }
  return lines.join('\r\n') + '\r\n';
}

export function parseTsv(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [], cur = '', inQ = false;
  const t = text.replace(/\r\n?/g, '\n');
  for (let i = 0; i < t.length; i++) {
    const ch = t[i];
    if (inQ) {
      if (ch === '"') { if (t[i + 1] === '"') { cur += '"'; i++; } else inQ = false; }
      else cur += ch;
    } else if (ch === '"' && cur === '') inQ = true;
    else if (ch === '\t') { row.push(cur); cur = ''; }
    else if (ch === '\n') { row.push(cur); rows.push(row); row = []; cur = ''; }
    else cur += ch;
  }
  if (cur !== '' || row.length) { row.push(cur); rows.push(row); }
  return rows;
}

export function matrixOf(sheet: Sheet, rg: Range): Matrix {
  const out: Matrix = [];
  for (let r = rg.r1; r <= rg.r2; r++) { const row: Scalar[] = []; for (let c = rg.c1; c <= rg.c2; c++) row.push(sheet.value(r, c)); out.push(row); }
  return out;
}

/** Current region (Ctrl+Shift+*): the contiguous block of non-empty cells around (r, c). */
export function currentRegion(sheet: Sheet, r: number, c: number): Range {
  const has = (rr: number, cc: number) => rr >= 0 && cc >= 0 && rr < MAX_ROWS && cc < MAX_COLS && (() => { const x = sheet.get(rr, cc); return !!x && (x.v != null && x.v !== '' || x.f != null); })();
  let rg: Range = { r1: r, c1: c, r2: r, c2: c };
  let grew = true;
  let guard = 0;
  while (grew && guard++ < 10000) {
    grew = false;
    const tryEdge = (rr1: number, cc1: number, rr2: number, cc2: number) => { for (let rr = rr1; rr <= rr2; rr++) for (let cc = cc1; cc <= cc2; cc++) if (has(rr, cc)) return true; return false; };
    if (rg.r1 > 0 && tryEdge(rg.r1 - 1, Math.max(0, rg.c1 - 1), rg.r1 - 1, rg.c2 + 1)) { rg.r1--; grew = true; }
    if (tryEdge(rg.r2 + 1, Math.max(0, rg.c1 - 1), rg.r2 + 1, rg.c2 + 1)) { rg.r2++; grew = true; }
    if (rg.c1 > 0 && tryEdge(Math.max(0, rg.r1 - 1), rg.c1 - 1, rg.r2 + 1, rg.c1 - 1)) { rg.c1--; grew = true; }
    if (tryEdge(Math.max(0, rg.r1 - 1), rg.c2 + 1, rg.r2 + 1, rg.c2 + 1)) { rg.c2++; grew = true; }
  }
  // Trim empty edges (the start cell may itself be blank)
  return rg;
}

/** The data region used by Sort, Filter, Table, PivotTable when a single cell is selected. */
export function dataRegion(app: App): Range {
  const rg = activeRange(app);
  if (rg.r1 === rg.r2 && rg.c1 === rg.c2) return currentRegion(app.sheet, rg.r1, rg.c1);
  return clipToUsed(app.sheet, rg);
}

/** Guesses whether the first row of a range is a header row. */
export function hasHeaderRow(sheet: Sheet, rg: Range): boolean {
  if (rg.r2 <= rg.r1) return false;
  let textTop = 0, numTop = 0, numBelow = 0;
  for (let c = rg.c1; c <= rg.c2; c++) {
    const a = sheet.value(rg.r1, c), b = sheet.value(rg.r1 + 1, c);
    if (typeof a === 'string') textTop++;
    if (typeof a === 'number') numTop++;
    if (typeof b === 'number') numBelow++;
  }
  const s1 = sheet.styleOf(rg.r1, rg.c1), s2 = sheet.styleOf(rg.r1 + 1, rg.c1);
  return (textTop > 0 && numTop === 0 && numBelow > 0) || (!!s1.bold && !s2.bold) || (textTop === rg.c2 - rg.c1 + 1 && numTop === 0);
}

export function addrText(rg: Range): string {
  const a = cellName(rg.r1, rg.c1), b = cellName(rg.r2, rg.c2);
  if (isWholeCols(rg)) return `${colName(rg.c1)}:${colName(rg.c2)}`;
  if (isWholeRows(rg)) return `${rg.r1 + 1}:${rg.r2 + 1}`;
  return a === b ? a : `${a}:${b}`;
}

export function uniqueName(book: Workbook, base: string, taken: (n: string) => boolean): string {
  let i = 1, n = `${base}${i}`;
  while (taken(n)) n = `${base}${++i}`;
  void book;
  return n;
}
