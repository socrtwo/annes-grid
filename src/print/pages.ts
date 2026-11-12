/**
 * Page layout for printing, Print Preview, Page Break Preview and Page Layout view: paper sizes,
 * print areas, print titles, scaling (fit to pages), automatic and manual page breaks, page order,
 * and header/footer codes.
 */
import type { Sheet } from '../model/workbook';
import type { PrintSettings } from '../model/types';
import { colIndex, type Range } from '../model/address';
import { resolveRangeText } from '../model/structure';

/** Paper sizes in inches (portrait width, height). */
export const PAPER_SIZES: Record<string, [number, number]> = {
  Letter: [8.5, 11], Legal: [8.5, 14], Tabloid: [11, 17], Ledger: [17, 11], Executive: [7.25, 10.5], Statement: [5.5, 8.5], Folio: [8.5, 13],
  A3: [11.69, 16.54], A4: [8.27, 11.69], A5: [5.83, 8.27], A6: [4.13, 5.83], B4: [9.84, 13.9], B5: [6.93, 9.84],
  'Envelope #10': [4.12, 9.5], 'Envelope DL': [4.33, 8.66], 'Envelope C5': [6.38, 9.02], 'Envelope Monarch': [3.87, 7.5],
  'Index Card 4x6': [4, 6], 'Photo 5x7': [5, 7],
};

export type MarginPreset = 'normal' | 'wide' | 'narrow';
export const MARGIN_PRESETS: Record<MarginPreset, { label: string; m: PrintSettings['margins'] }> = {
  normal: { label: 'Normal', m: { top: 0.75, bottom: 0.75, left: 0.7, right: 0.7, header: 0.3, footer: 0.3 } },
  wide: { label: 'Wide', m: { top: 1, bottom: 1, left: 1, right: 1, header: 0.5, footer: 0.5 } },
  narrow: { label: 'Narrow', m: { top: 0.75, bottom: 0.75, left: 0.25, right: 0.25, header: 0.3, footer: 0.3 } },
};

/** Size of the row/column headings when they print (px at 100%). */
export const HEADING_W = 34;
export const HEADING_H = 18;

/** Paper size in inches for the orientation. */
export function paperInches(p: PrintSettings): [number, number] {
  const [w, h] = PAPER_SIZES[p.paper] || PAPER_SIZES.Letter;
  return p.orientation === 'landscape' ? [Math.max(w, h), Math.min(w, h)] : [Math.min(w, h), Math.max(w, h)];
}

/** Lower-right cell reached by an object anchored at `at` (so objects count toward what prints). */
export function objectExtent(sheet: Sheet, at: { r: number; c: number; dx: number; dy: number; w: number; h: number }): { r: number; c: number } {
  let c = at.c, x = sheet.colWidth(c) - at.dx;
  while (x < at.w && c < at.c + 500) { c++; x += sheet.colWidth(c); }
  let r = at.r, y = sheet.rowHeight(r) - at.dy;
  while (y < at.h && r < at.r + 2000) { r++; y += sheet.rowHeight(r); }
  return { r, c };
}

/** The used range plus the cells that floating objects cover. */
export function contentRange(sheet: Sheet): Range | null {
  let u = sheet.usedRange();
  for (const o of [...sheet.charts, ...sheet.drawings]) {
    const e = objectExtent(sheet, o.at);
    const rg = { r1: o.at.r, c1: o.at.c, r2: e.r, c2: e.c };
    u = u ? { r1: Math.min(u.r1, rg.r1), c1: Math.min(u.c1, rg.c1), r2: Math.max(u.r2, rg.r2), c2: Math.max(u.c2, rg.c2) } : rg;
  }
  return u;
}

/** Every part of the print area (each prints on its own pages), clipped to the content; or the content. */
export function printAreas(sheet: Sheet, ignoreArea = false): Range[] {
  const content = contentRange(sheet);
  if (sheet.print.area && !ignoreArea) {
    const out: Range[] = [];
    for (const part of splitAreaText(sheet.print.area)) {
      const r = resolveRangeText(sheet.book, part, sheet);
      if (!r) continue;
      const rg = { ...r.range };
      // Whole rows or columns print only as far as there is content.
      if (content) { rg.r2 = Math.min(rg.r2, Math.max(content.r2, rg.r1)); rg.c2 = Math.min(rg.c2, Math.max(content.c2, rg.c1)); }
      else { rg.r2 = Math.min(rg.r2, rg.r1 + 199); rg.c2 = Math.min(rg.c2, rg.c1 + 49); }
      out.push(rg);
    }
    if (out.length) return out;
  }
  return content ? [content] : [];
}

/** Splits "A1:B2,Sheet1!C3:D4" at the commas outside quotes. */
export function splitAreaText(text: string): string[] {
  const out: string[] = [];
  let cur = '', q = false;
  for (const ch of text) {
    if (ch === "'") q = !q;
    if (ch === ',' && !q) { if (cur.trim()) out.push(cur.trim()); cur = ''; } else cur += ch;
  }
  if (cur.trim()) out.push(cur.trim());
  return out;
}

/** The area that prints first: the print area, or the used range. */
export function printArea(sheet: Sheet): Range | null {
  return printAreas(sheet)[0] || null;
}

/** Rows ("$1:$2", "1:2", "$3") or columns ("$A:$B", "A") to repeat, as an inclusive 0-based span. */
export function parseTitleSpan(text: string | undefined, axis: 'row' | 'col'): [number, number] | null {
  if (!text) return null;
  const t = text.replace(/^.*!/, '').replace(/\$/g, '').trim().toUpperCase();
  if (!t) return null;
  const parts = t.split(':');
  if (parts.length > 2) return null;
  const conv = (s: string): number => {
    if (axis === 'row') return /^\d+$/.test(s) ? parseInt(s, 10) - 1 : NaN;
    return /^[A-Z]{1,3}$/.test(s) ? colIndex(s) : NaN;
  };
  const a = conv(parts[0]), b = conv(parts[1] ?? parts[0]);
  if (!(a >= 0) || !(b >= 0)) return null;
  return [Math.min(a, b), Math.max(a, b)];
}

export interface PageSpec { r1: number; r2: number; c1: number; c2: number; area: number }

export interface SheetPages {
  /** Scale in percent. */
  scale: number;
  areas: Range[];
  /** Rows and columns where a new page starts (all areas). */
  rowBreaks: number[];
  colBreaks: number[];
  pages: PageSpec[];
  titleRows: [number, number] | null;
  titleCols: [number, number] | null;
}

function spanSize(sheet: Sheet, span: [number, number] | null, axis: 'row' | 'col', before: number): number {
  if (!span || before <= span[0]) return 0;
  let s = 0;
  for (let i = span[0]; i <= Math.min(span[1], before - 1); i++) s += axis === 'row' ? sheet.rowHeight(i) : sheet.colWidth(i);
  return s;
}

/** Scale (percent) that fits the areas on fitWide × fitTall pages; 100 when not fitting. */
export function fitScale(sheet: Sheet, area: Range | Range[], wIn?: number, hIn?: number): number {
  const p = sheet.print;
  const areas = Array.isArray(area) ? area : [area];
  if (wIn == null || hIn == null) [wIn, hIn] = paperInches(p);
  const availW = (wIn - p.margins.left - p.margins.right) * 96 - (p.headings ? HEADING_W : 0);
  const availH = (hIn - p.margins.top - p.margins.bottom) * 96 - (p.headings ? HEADING_H : 0);
  let s = 100;
  for (const a of areas) {
    let totalW = 0, totalH = 0;
    for (let c = a.c1; c <= a.c2; c++) totalW += sheet.colWidth(c);
    for (let r = a.r1; r <= a.r2; r++) totalH += sheet.rowHeight(r);
    // Repeated print titles take room again on every page after the first.
    const tr = parseTitleSpan(p.titleRows, 'row'), tc = parseTitleSpan(p.titleCols, 'col');
    totalW += spanSize(sheet, tc, 'col', a.c1) * Math.max(0, (p.fitWide || 1) - 1);
    totalH += spanSize(sheet, tr, 'row', a.r1) * Math.max(0, (p.fitTall || 1) - 1);
    if (p.fitWide && totalW > 0) s = Math.min(s, (availW * p.fitWide) / totalW * 100);
    if (p.fitTall && totalH > 0) s = Math.min(s, (availH * p.fitTall) / totalH * 100);
  }
  return Math.max(10, Math.min(400, Math.floor(s)));
}

/** Breaks along one axis: a new page starts where the next row/column no longer fits, or at a manual break. */
function axisBreaks(from: number, to: number, size: (i: number) => number, avail: number, manual: number[], titleSize: (start: number) => number): number[] {
  const out: number[] = [];
  const man = new Set(manual);
  let start = from;
  let acc = 0;
  for (let i = from; i <= to; i++) {
    const s = size(i);
    const room = avail - titleSize(start);
    if ((man.has(i) && i > from) || (acc + s > room && acc > 0)) { out.push(i); start = i; acc = 0; }
    acc += s;
  }
  return out;
}

/** Pages of a sheet (or of the given areas, e.g. the selection) in print order. */
export function layoutPages(sheet: Sheet, opts: { areas?: Range[]; ignoreArea?: boolean } = {}): SheetPages {
  const p = sheet.print;
  const areas = opts.areas || printAreas(sheet, opts.ignoreArea);
  const [w, hgt] = paperInches(p);
  if (!(p.fitWide || p.fitTall)) return layoutAt(sheet, areas, Math.max(10, Math.min(400, p.scale || 100))).pages;
  // The fit scale assumes rows and columns could be split between pages; they cannot, so each page wastes part of a
  // row or column and the result can run one page over. The scale comes down until the pages really fit.
  let scale = areas.length ? fitScale(sheet, areas, w, hgt) : 100;
  let l = layoutAt(sheet, areas, scale);
  while (l.over && scale > 10) l = layoutAt(sheet, areas, --scale);
  return l.pages;
}

function layoutAt(sheet: Sheet, areas: Range[], scale: number): { pages: SheetPages; over: boolean } {
  const p = sheet.print;
  const [w, hgt] = paperInches(p);
  let over = false;
  const k = scale / 100;
  const availW = ((w - p.margins.left - p.margins.right) * 96) / k - (p.headings ? HEADING_W : 0);
  const availH = ((hgt - p.margins.top - p.margins.bottom) * 96) / k - (p.headings ? HEADING_H : 0);
  const titleRows = parseTitleSpan(p.titleRows, 'row');
  const titleCols = parseTitleSpan(p.titleCols, 'col');
  const rowsAll: number[] = [], colsAll: number[] = [];
  const pages: PageSpec[] = [];
  areas.forEach((a, ai) => {
    // Manual page breaks are ignored while the sheet is scaled to fit a number of pages (as in Excel).
    const fit = !!(p.fitWide || p.fitTall);
    const rb = axisBreaks(a.r1, a.r2, (r) => sheet.rowHeight(r), availH, fit ? [] : sheet.rowBreaks, (start) => spanSize(sheet, titleRows, 'row', start));
    const cb = axisBreaks(a.c1, a.c2, (c) => sheet.colWidth(c), availW, fit ? [] : sheet.colBreaks, (start) => spanSize(sheet, titleCols, 'col', start));
    rowsAll.push(...rb); colsAll.push(...cb);
    const rs = [a.r1, ...rb], cs = [a.c1, ...cb];
    if ((p.fitTall && rs.length > p.fitTall) || (p.fitWide && cs.length > p.fitWide)) over = true;
    const rSpans = rs.map((r, i) => [r, (rs[i + 1] ?? a.r2 + 1) - 1] as [number, number]);
    const cSpans = cs.map((c, i) => [c, (cs[i + 1] ?? a.c2 + 1) - 1] as [number, number]);
    const push = (rr: [number, number], cc: [number, number]) => {
      // Pages whose rows or columns are all hidden do not print.
      let anyR = false, anyC = false;
      for (let r = rr[0]; r <= rr[1] && !anyR; r++) if (sheet.rowHeight(r) > 0) anyR = true;
      for (let c = cc[0]; c <= cc[1] && !anyC; c++) if (sheet.colWidth(c) > 0) anyC = true;
      if (anyR && anyC) pages.push({ r1: rr[0], r2: rr[1], c1: cc[0], c2: cc[1], area: ai });
    };
    if (p.order === 'overThenDown') { for (const rr of rSpans) for (const cc of cSpans) push(rr, cc); }
    else { for (const cc of cSpans) for (const rr of rSpans) push(rr, cc); }
  });
  return { pages: { scale, areas, rowBreaks: [...new Set(rowsAll)].sort((a, b) => a - b), colBreaks: [...new Set(colsAll)].sort((a, b) => a - b), pages, titleRows, titleCols }, over };
}

/** Automatic plus manual page breaks (row and column indexes where a new page starts). */
export function computePageBreaks(sheet: Sheet): { rows: number[]; cols: number[]; area: Range } {
  const l = layoutPages(sheet);
  return { rows: l.rowBreaks, cols: l.colBreaks, area: l.areas[0] || { r1: 0, c1: 0, r2: 0, c2: 0 } };
}

// --------------------------------------------------------------------------------------------
// Header and footer codes

export interface HFContext { page: number; pages: number; date: Date; file: string; path: string; sheet: string }
export interface HFRun { text: string; bold?: boolean; italic?: boolean; underline?: boolean; strike?: boolean; size?: number; font?: string; color?: string; sup?: boolean; sub?: boolean }

const BRACKET_CODES: Record<string, string> = { page: 'P', pages: 'N', date: 'D', time: 'T', file: 'F', path: 'Z', tab: 'A', picture: 'G' };

/**
 * Expands header/footer codes into formatted runs. Supports the readable codes (&[Page], &[Pages],
 * &[Date], &[Time], &[File], &[Path], &[Tab]) and the spreadsheet file codes (&P, &N, &D, &T, &F,
 * &Z, &A, &P+n, &P-n, &&, &B, &I, &U, &E, &S, &X, &Y, &"Font,Style", &nn size, &Krrggbb color).
 */
export function parseHeaderFooter(src: string, ctx: HFContext): HFRun[] {
  const runs: HFRun[] = [];
  let st: Omit<HFRun, 'text'> = {};
  let buf = '';
  const flush = () => { if (buf) { runs.push({ text: buf, ...st }); buf = ''; } };
  const setFmt = (patch: Partial<HFRun>) => { flush(); st = { ...st, ...patch }; };
  const value = (code: string, rest: string): { text: string; used: number } => {
    switch (code) {
      case 'P': {
        const m = /^([+-])(\d+)/.exec(rest);
        if (m) return { text: String(ctx.page + (m[1] === '+' ? 1 : -1) * parseInt(m[2], 10)), used: m[0].length };
        return { text: String(ctx.page), used: 0 };
      }
      case 'N': return { text: String(ctx.pages), used: 0 };
      case 'D': return { text: ctx.date.toLocaleDateString(), used: 0 };
      case 'T': return { text: ctx.date.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' }), used: 0 };
      case 'F': return { text: ctx.file, used: 0 };
      case 'Z': return { text: ctx.path, used: 0 };
      case 'A': return { text: ctx.sheet, used: 0 };
      default: return { text: '', used: 0 };
    }
  };
  let i = 0;
  while (i < src.length) {
    const ch = src[i];
    if (ch !== '&') { buf += ch; i++; continue; }
    const nx = src[i + 1];
    if (nx === undefined) { buf += '&'; i++; continue; }
    if (nx === '&') { buf += '&'; i += 2; continue; }
    if (nx === '[') {
      const end = src.indexOf(']', i);
      const name = end > 0 ? src.slice(i + 2, end).toLowerCase() : '';
      const code = BRACKET_CODES[name];
      if (code) {
        const v = value(code, src.slice(end + 1));
        buf += v.text;
        i = end + 1 + v.used;
        continue;
      }
      buf += '&'; i++; continue;
    }
    if (nx === '"') {
      const end = src.indexOf('"', i + 2);
      if (end < 0) { i = src.length; continue; }
      const [font, style = ''] = src.slice(i + 2, end).split(',');
      const s = style.toLowerCase();
      setFmt({ ...(font && font !== '-' ? { font } : {}), ...(s ? { bold: /bold/.test(s), italic: /italic|oblique/.test(s) } : {}) });
      i = end + 1;
      continue;
    }
    if (/\d/.test(nx)) {
      const m = /^\d+/.exec(src.slice(i + 1))!;
      setFmt({ size: parseInt(m[0], 10) });
      i += 1 + m[0].length;
      continue;
    }
    if (nx === 'K') {
      const m = /^[0-9A-Fa-f]{6}/.exec(src.slice(i + 2));
      if (m) { setFmt({ color: '#' + m[0].toLowerCase() }); i += 2 + 6; continue; }
      i += 2; continue;
    }
    const up = nx.toUpperCase();
    switch (up) {
      case 'B': setFmt({ bold: !st.bold }); break;
      case 'I': setFmt({ italic: !st.italic }); break;
      case 'U': case 'E': setFmt({ underline: !st.underline }); break;
      case 'S': setFmt({ strike: !st.strike }); break;
      case 'X': setFmt({ sup: !st.sup, sub: false }); break;
      case 'Y': setFmt({ sub: !st.sub, sup: false }); break;
      case 'L': case 'C': case 'R': break; // section markers (sections are stored separately)
      case 'G': break; // pictures are not printed in headers
      default: {
        if (['P', 'N', 'D', 'T', 'F', 'Z', 'A'].includes(up)) {
          const v = value(up, src.slice(i + 2));
          buf += v.text;
          i += 2 + v.used;
          continue;
        }
        buf += '&' + nx;
      }
    }
    i += 2;
  }
  flush();
  return runs;
}

/** Header/footer text with the codes expanded (formatting dropped). */
export function expandHeaderFooter(src: string, ctx: HFContext): string {
  return parseHeaderFooter(src, ctx).map((r) => r.text).join('');
}

/** Header/footer sections for a page (first page and even pages may differ). */
export function headerFooterFor(p: PrintSettings, pageIndex: number, pageNumber: number): { header: PrintSettings['header']; footer: PrintSettings['footer'] } {
  if (p.differentFirst && pageIndex === 0) return { header: p.firstHeader || { left: '', center: '', right: '' }, footer: p.firstFooter || { left: '', center: '', right: '' } };
  if (p.differentOddEven && pageNumber % 2 === 0) return { header: p.evenHeader || { left: '', center: '', right: '' }, footer: p.evenFooter || { left: '', center: '', right: '' } };
  return { header: p.header, footer: p.footer };
}
