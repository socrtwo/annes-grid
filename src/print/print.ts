/**
 * Printing and Print Preview.
 *
 * Pages are rendered as HTML (one .ql-print-page per sheet of paper, sized with @page) so the
 * system print dialog can print them or save them as PDF. Each page holds the sheet's cells as a
 * table (fonts, fills, borders, alignment, wrapping, merges, column widths and row heights, scaled
 * by the print scale or fit-to-page), repeated print titles, the header and footer, optional
 * gridlines and row/column headings, and the charts and pictures that fall on the page.
 */
import type { App, BookTab } from '../app';
import type { Sheet } from '../model/workbook';
import type { Cell, CellStyle, Border, CellComment, PrintSettings } from '../model/types';
import { CellError } from '../model/types';
import { cellKey, cellName, colName, keyRow, keyCol, MAX_COLS, MAX_ROWS, type Range } from '../model/address';
import { formatValue, formatGeneral } from '../model/numfmt';
import { tableCellLook } from '../model/tableStyles';
import { fontFamilyCss } from '../fonts';
import { renderChartSvg } from '../charts/render';
import { renderShapeSvg } from '../charts/shapes';
import { escHtml, safeColor, safeImageSrc, safeNum, sanitizeGeneratedHtml } from '../util/sanitize';
import { h, icon, clear, isPhone } from '../ui/dom';
import { openDialog, toast } from '../ui/dialog';
import { layoutPages, paperInches, parseHeaderFooter, headerFooterFor, objectExtent, contentRange, HEADING_W, HEADING_H, PAPER_SIZES, MARGIN_PRESETS, type PageSpec, type SheetPages, type HFContext, type HFRun } from './pages';

export type PrintWhat = 'sheet' | 'book' | 'selection';
export interface PrintOptions { what?: PrintWhat; ignoreArea?: boolean }

/** What the renderer needs from a workbook tab (a BookTab, or an equivalent in tests). */
export interface PrintSource { tab: BookTab; file: string; path: string; hideInk?: boolean }

interface PageJob { sheet: Sheet; layout: SheetPages; spec: PageSpec | null; comments?: { r: number; c: number; cm: CellComment }[]; index: number }

const MAX_PAGES = 2000;

// --------------------------------------------------------------------------------------------
// Planning

/** The pages a print job produces, in order. */
export function planPrint(tab: BookTab, opts: PrintOptions = {}): PageJob[] {
  const book = tab.book;
  const sheets = opts.what === 'book' ? book.sheets.filter((s) => s.visibility === 'visible') : [book.sheet];
  const jobs: PageJob[] = [];
  for (const sheet of sheets) {
    let areas: Range[] | undefined;
    if (opts.what === 'selection') {
      const content = contentRange(sheet);
      areas = sheet.sel.ranges.map((rg) => {
        const out = { ...rg };
        if (rg.r2 >= MAX_ROWS - 1 || rg.c2 >= MAX_COLS - 1) {
          if (!content) return null;
          out.r2 = Math.min(rg.r2, content.r2); out.c2 = Math.min(rg.c2, content.c2);
        }
        return out.r2 >= out.r1 && out.c2 >= out.c1 ? out : null;
      }).filter((x): x is Range => !!x);
    }
    const layout = layoutPages(sheet, { areas, ignoreArea: opts.ignoreArea });
    let i = 0;
    for (const spec of layout.pages) {
      if (jobs.length >= MAX_PAGES) break;
      jobs.push({ sheet, layout, spec, index: i++ });
    }
    if (sheet.print.comments === 'atEnd' && sheet.comments.size) {
      const list = [...sheet.comments].sort((a, b) => a[0] - b[0]).map(([k, cm]) => ({ r: keyRow(k), c: keyCol(k), cm }));
      for (let j = 0; j < list.length; j += 24) jobs.push({ sheet, layout, spec: null, comments: list.slice(j, j + 24), index: i++ });
    }
  }
  return jobs;
}

// --------------------------------------------------------------------------------------------
// Cell rendering

const BORDER_CSS: Record<Border['style'], string> = {
  thin: '1px solid', medium: '2px solid', thick: '3px solid', dashed: '1px dashed', dotted: '1px dotted', double: '3px double', hair: '1px dotted', mediumDashed: '2px dashed', dashDot: '1px dashed',
};

function borderCss(b: Border | undefined, bw: boolean): string | null {
  if (!b) return null;
  return `${BORDER_CSS[b.style] || '1px solid'} ${bw ? '#000' : safeColor(b.color || '#000000')}`;
}

function fontCss(st: CellStyle, base: CellStyle, themeFont: string): string {
  const parts = [`font-family:${fontFamilyCss(st.font || base.font || themeFont || 'Carlito')}`, `font-size:${st.size || base.size || 11}pt`];
  if (st.bold) parts.push('font-weight:700');
  if (st.italic) parts.push('font-style:italic');
  const deco = [st.underline ? 'underline' : '', st.strike ? 'line-through' : ''].filter(Boolean).join(' ');
  if (deco) parts.push(`text-decoration:${deco}`);
  if (st.underline === 'double' || st.underline === 'doubleAccounting') parts.push('text-decoration-style:double');
  if (st.vertAlign) parts.push(`vertical-align:${st.vertAlign === 'superscript' ? 'super' : 'sub'};font-size:${Math.round((st.size || base.size || 11) * 0.7)}pt`);
  return parts.join(';');
}

interface CellOpts { bw: boolean; draft: boolean; grid: boolean; errors: PrintSettings['errors'] }

function displayText(sheet: Sheet, cell: Cell, st: CellStyle, opts: CellOpts): { text: string; color?: string; isNum: boolean } {
  if (sheet.showFormulas && cell.f != null) return { text: '=' + cell.f, isNum: false };
  const v = cell.v ?? null;
  if (v instanceof CellError) {
    if (opts.errors === 'blank') return { text: '', isNum: false };
    if (opts.errors === 'dashes') return { text: '--', isNum: false };
    if (opts.errors === 'na') return { text: '#N/A', isNum: false };
    return { text: v.code, isNum: false };
  }
  if (typeof v === 'number' && v === 0 && !sheet.showZeros) return { text: '', isNum: true };
  const f = formatValue(v, st.numFmt, { date1904: sheet.book.calc.date1904 });
  return { text: f.text, color: f.color, isNum: !!f.isNumber || typeof v === 'number' };
}

/** HTML of a header/footer section. */
export function hfHtml(runs: HFRun[], bw = false): string {
  return runs.map((r) => {
    const css: string[] = [];
    if (r.bold) css.push('font-weight:700');
    if (r.italic) css.push('font-style:italic');
    const deco = [r.underline ? 'underline' : '', r.strike ? 'line-through' : ''].filter(Boolean).join(' ');
    if (deco) css.push(`text-decoration:${deco}`);
    if (r.size) css.push(`font-size:${Math.max(4, Math.min(72, r.size))}pt`);
    if (r.font) css.push(`font-family:${fontFamilyCss(r.font)}`);
    if (r.color && !bw) css.push(`color:${safeColor(r.color)}`);
    if (r.sup || r.sub) css.push(`vertical-align:${r.sup ? 'super' : 'sub'}`);
    return css.length ? `<span style="${escHtml(css.join(';'))}">${escHtml(r.text)}</span>` : escHtml(r.text);
  }).join('');
}

/** Rows (or columns) shown on a page: repeated titles above (left of) the page, then the page's own, without hidden ones. */
function pageIndexes(sheet: Sheet, title: [number, number] | null, from: number, to: number, axis: 'row' | 'col'): { list: number[]; titles: number } {
  const size = (i: number) => (axis === 'row' ? sheet.rowHeight(i) : sheet.colWidth(i));
  const list: number[] = [];
  if (title && from > title[0]) for (let i = title[0]; i <= Math.min(title[1], from - 1); i++) if (size(i) > 0) list.push(i);
  const titles = list.length;
  for (let i = from; i <= to; i++) if (size(i) > 0) list.push(i);
  return { list, titles };
}

/** Sum of row heights (or column widths) from `a` up to, not including, `b` (negative when b < a). */
function offsetBetween(sheet: Sheet, a: number, b: number, axis: 'row' | 'col'): number {
  const size = (i: number) => (axis === 'row' ? sheet.rowHeight(i) : sheet.colWidth(i));
  let s = 0;
  if (b >= a) for (let i = a; i < b; i++) s += size(i);
  else for (let i = b; i < a; i++) s -= size(i);
  return s;
}

/** The cells of a page as an HTML table (rows and columns are sheet indexes in display order). */
export function renderCellsTable(tab: BookTab, sheet: Sheet, rows: number[], cols: number[], opts: CellOpts & { headings: boolean }): string {
  const book = sheet.book;
  const base = book.styles[0] || {};
  const themeFont = book.theme?.bodyFont || 'Carlito';
  const version = book.version + ((tab.engine as any)?.evaluations || 0);
  const rowPos = new Map(rows.map((r, i) => [r, i]));
  const colPos = new Map(cols.map((c, i) => [c, i]));
  const merges = sheet.merges.filter((m) => rows.some((r) => r >= m.r1 && r <= m.r2) && cols.some((c) => c >= m.c1 && c <= m.c2));
  const mergeAt = (r: number, c: number) => merges.find((m) => r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2) || null;
  const tables = sheet.tables;
  const totalW = cols.reduce((s, c) => s + sheet.colWidth(c), 0) + (opts.headings ? HEADING_W : 0);
  const out: string[] = [];
  const gridCss = opts.grid ? 'border:0.5px solid #b7b7b7;' : '';
  out.push(`<table class="ql-pt" style="${escHtml(`width:${totalW}px;${fontCss({}, base, themeFont)}`)}"><colgroup>`);
  if (opts.headings) out.push(`<col style="width:${HEADING_W}px">`);
  for (const c of cols) out.push(`<col style="width:${safeNum(sheet.colWidth(c), 0, 0, 1e5)}px">`);
  out.push('</colgroup>');
  const headCss = `background:${opts.bw ? '#fff' : '#f0f0f0'};border:0.5px solid #999;font:8pt sans-serif;text-align:center;color:#333;padding:0;`;
  if (opts.headings) {
    out.push(`<tr style="height:${HEADING_H}px"><th style="${headCss}"></th>`);
    for (const c of cols) out.push(`<th style="${headCss}">${colName(c)}</th>`);
    out.push('</tr>');
  }
  const occupied = (r: number, c: number) => { const x = sheet.get(r, c); return !!x && ((x.v != null && x.v !== '') || x.f != null); };
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const rh = safeNum(sheet.rowHeight(r), 0, 0, 1e5);
    out.push(`<tr style="height:${rh}px">`);
    if (opts.headings) out.push(`<th style="${headCss}">${r + 1}</th>`);
    for (let j = 0; j < cols.length; j++) {
      const c = cols[j];
      const m = mergeAt(r, c);
      let rs = 1, cs = 1, ar = r, ac = c;
      if (m) {
        // The first visible cell of the merge on this page draws it; the rest are covered.
        const prevR = i > 0 ? rows[i - 1] : -1, prevC = j > 0 ? cols[j - 1] : -1;
        const coveredR = i > 0 && prevR >= m.r1 && prevR <= m.r2 && rowPos.get(prevR) === i - 1;
        const coveredC = j > 0 && prevC >= m.c1 && prevC <= m.c2 && colPos.get(prevC) === j - 1;
        if (coveredR || coveredC) continue;
        while (i + rs < rows.length && rows[i + rs] >= m.r1 && rows[i + rs] <= m.r2) rs++;
        while (j + cs < cols.length && cols[j + cs] >= m.c1 && cols[j + cs] <= m.c2) cs++;
        ar = m.r1; ac = m.c1;
      }
      const cell = sheet.get(ar, ac);
      const sid = cell?.s ?? sheet.rowStyle.get(ar) ?? sheet.colStyle.get(ac) ?? 0;
      let st: CellStyle = book.styles[sid] || {};
      const cf = sheet.cf.length && (tab as any).cf ? tab.cf.evaluate(sheet, ar, ac, version) : null;
      let look: ReturnType<typeof tableCellLook> = null;
      for (const tb of tables) { const l = tableCellLook(tb, ar, ac, book.theme); if (l) { look = l; break; } }
      if (cf?.style) st = { ...st, ...cf.style };
      if (look?.bold && st.bold == null) st = { ...st, bold: true };
      // Borders: a merge takes the outer borders of its edge cells.
      const lastR = m ? m.r2 : r, lastC = m ? m.c2 : c;
      const stAt = (rr: number, cc: number) => sheet.styleOf(rr, cc);
      const bt = st.bt, bl = st.bl;
      const bb = m ? stAt(lastR, ac).bb : st.bb, br = m ? stAt(ar, lastC).br : st.br;
      const css: string[] = [gridCss, 'padding:0 2px;overflow:visible;'];
      const sides: [string, Border | undefined][] = [['top', bt], ['bottom', bb], ['left', bl], ['right', br]];
      for (const [side, b] of sides) { const v = borderCss(b, opts.bw); if (v) css.push(`border-${side}:${v};`); }
      if (!opts.bw) {
        if (look?.borderTop && !bt) css.push(`border-top:1px solid ${safeColor(look.borderTop)};`);
        if (look?.borderBottom && !bb) css.push(`border-bottom:1px solid ${safeColor(look.borderBottom)};`);
        const fill = cf?.fill || cf?.style?.fill || st.fill || look?.fill || (st.pattern?.color && st.pattern.type === 'solid' ? st.pattern.color : undefined);
        if (st.gradient && !cf?.fill) css.push(`background:linear-gradient(${(st.gradient.angle ?? 0) === 90 ? 'to bottom' : 'to right'},${safeColor(st.gradient.from)},${safeColor(st.gradient.to)});`);
        else if (cf?.bar) {
          const pct = Math.max(0, Math.min(100, cf.bar.pct * 100));
          css.push(`background:linear-gradient(to right,${safeColor(cf.bar.color)} ${pct}%,${fill ? safeColor(fill) : 'transparent'} ${pct}%);`);
        } else if (fill) css.push(`background:${safeColor(fill)};`);
      }
      const va = st.vAlign === 'top' ? 'top' : st.vAlign === 'middle' || st.vAlign === 'distributed' || st.vAlign === 'justify' ? 'middle' : 'bottom';
      css.push(`vertical-align:${va};`);
      let inner = '';
      let height = 0;
      for (let k = 0; k < rs; k++) height += sheet.rowHeight(rows[i + k]);
      let width = 0;
      for (let k = 0; k < cs; k++) width += sheet.colWidth(cols[j + k]);
      if (cell && (cell.v != null || cell.f != null) && !(cf?.hideValue)) {
        const disp = displayText(sheet, cell, st, opts);
        let text = disp.text;
        const isNum = disp.isNum && !(sheet.showFormulas && cell.f != null);
        let align = st.hAlign && st.hAlign !== 'general' ? st.hAlign : isNum ? 'right' : typeof cell.v === 'boolean' || cell.v instanceof CellError ? 'center' : 'left';
        if (sheet.rtl && (!st.hAlign || st.hAlign === 'general')) align = align === 'left' ? 'right' : align === 'right' ? 'left' : align;
        const color = opts.bw ? '#000' : safeColor(cf?.style?.color || disp.color || st.color || look?.color || (cell.link ? '#0563c1' : '#000000'));
        const dcss: string[] = [fontCss(st, base, themeFont), `color:${color}`, `max-height:${height}px`, 'line-height:1.2'];
        if (st.indent) dcss.push(`padding-${align === 'right' ? 'right' : 'left'}:${st.indent * 9}px`);
        // Numbers that do not fit show ### (General numbers first lose digits).
        const charW = ((st.size || base.size || 11) / 11) * 7;
        if (isNum && typeof cell.v === 'number' && !st.wrap && !st.shrink && text.length * charW > width - 4) {
          if (!st.numFmt || st.numFmt === 'General') text = formatGeneral(cell.v, Math.max(1, Math.floor((width - 4) / charW)));
          if (text.length * charW > width - 4) text = '#'.repeat(Math.max(1, Math.floor((width - 4) / charW)));
        }
        if (st.wrap || align === 'justify' || align === 'distributed') {
          dcss.push(`width:${width - 4}px`, 'white-space:pre-wrap', 'overflow-wrap:anywhere', 'overflow:hidden', `text-align:${align === 'distributed' ? 'justify' : align}`);
        } else if (align === 'fill' && text) {
          const n = Math.max(1, Math.floor((width - 4) / Math.max(1, text.length * charW)));
          text = text.repeat(n);
          dcss.push(`width:${width - 4}px`, 'white-space:pre', 'overflow:hidden');
        } else {
          // Text overflows into empty neighbors on the page, as on screen.
          let extraR = 0, extraL = 0;
          if (!m && !isNum) {
            if (align === 'left' || align === 'center' || align === 'centerContinuous') for (let k = j + 1; k < cols.length && !occupied(r, cols[k]) && !mergeAt(r, cols[k]); k++) extraR += sheet.colWidth(cols[k]);
            if (align === 'right' || align === 'center' || align === 'centerContinuous') for (let k = j - 1; k >= 0 && !occupied(r, cols[k]) && !mergeAt(r, cols[k]); k--) extraL += sheet.colWidth(cols[k]);
            if (align === 'center' || align === 'centerContinuous') { const e = Math.min(extraL, extraR); extraL = e; extraR = e; }
          }
          const ta = align === 'centerContinuous' ? 'center' : align;
          dcss.push(`width:${width - 4 + extraL + extraR}px`, extraL ? `margin-left:${-extraL}px` : '', 'white-space:pre', 'overflow:hidden', `text-align:${ta}`);
        }
        if (st.rotate && st.rotate !== 255) dcss.push(`transform:rotate(${-Math.max(-90, Math.min(90, st.rotate))}deg)`);
        if (st.rotate === 255) dcss.push('writing-mode:vertical-rl;text-orientation:upright');
        const body = cell.rich && !(sheet.showFormulas && cell.f != null)
          ? cell.rich.map((run) => (run.s ? `<span style="${escHtml(runCss(run.s, opts.bw))}">${escHtml(run.t)}</span>` : escHtml(run.t))).join('')
          : escHtml(text);
        inner = `<div style="${escHtml(dcss.filter(Boolean).join(';'))}">${body}</div>`;
      }
      out.push(`<td${rs > 1 ? ` rowspan="${rs}"` : ''}${cs > 1 ? ` colspan="${cs}"` : ''} style="${escHtml(css.join(''))}">${inner}</td>`);
    }
    out.push('</tr>');
  }
  out.push('</table>');
  return out.join('');
}

function runCss(s: Partial<CellStyle>, bw: boolean): string {
  const css: string[] = [];
  if (s.font) css.push(`font-family:${fontFamilyCss(s.font)}`);
  if (s.size) css.push(`font-size:${s.size}pt`);
  if (s.bold != null) css.push(`font-weight:${s.bold ? 700 : 400}`);
  if (s.italic != null) css.push(`font-style:${s.italic ? 'italic' : 'normal'}`);
  const deco = [s.underline ? 'underline' : '', s.strike ? 'line-through' : ''].filter(Boolean).join(' ');
  if (deco) css.push(`text-decoration:${deco}`);
  if (s.color && !bw) css.push(`color:${safeColor(s.color)}`);
  if (s.vertAlign) css.push(`vertical-align:${s.vertAlign === 'superscript' ? 'super' : 'sub'}`);
  return css.join(';');
}

// --------------------------------------------------------------------------------------------
// Pages


/** Charts and pictures that fall on the page body, positioned relative to its first cell. */
function objectsHtml(src: PrintSource, sheet: Sheet, spec: PageSpec, bodyW: number, bodyH: number, draft: boolean): string {
  if (draft) return '';
  const out: string[] = [];
  const objs = [...sheet.charts.map((d) => ({ kind: 'chart' as const, d })), ...sheet.drawings.map((d) => ({ kind: 'drawing' as const, d }))];
  for (const o of objs) {
    const at = o.d.at;
    if (at.r > spec.r2 || at.c > spec.c2) continue;
    const ext = objectExtent(sheet, at);
    if (ext.r < spec.r1 || ext.c < spec.c1) continue;
    const x = offsetBetween(sheet, spec.c1, at.c, 'col') + at.dx;
    const y = offsetBetween(sheet, spec.r1, at.r, 'row') + at.dy;
    if (x > bodyW || y > bodyH || x + at.w < 0 || y + at.h < 0) continue;
    let body = '';
    try {
      if (o.kind === 'chart') body = renderChartSvg(src.tab, o.d, at.w, at.h);
      else {
        const d = o.d;
        if (d.kind === 'ink' && src.hideInk) continue;
        if (d.kind === 'image' || d.kind === 'signature') {
          const s = safeImageSrc(d.src);
          if (!s) continue;
          body = `<img src="${escHtml(s)}" alt="${escHtml(d.alt || '')}" style="width:100%;height:100%;object-fit:${d.lockAspect === false ? 'fill' : 'contain'}">`;
        } else if (d.kind === 'icon' && d.svg) body = sanitizeGeneratedHtml(d.svg);
        else body = renderShapeSvg(d);
      }
    } catch { continue; }
    const rot = o.kind === 'drawing' && o.d.rotate ? `transform:rotate(${Number(o.d.rotate) || 0}deg);` : '';
    out.push(`<div class="ql-pt-obj" style="left:${safeNum(x, 0)}px;top:${safeNum(y, 0)}px;width:${safeNum(at.w, 0, 0)}px;height:${safeNum(at.h, 0, 0)}px;${rot}">${body}</div>`);
  }
  return out.join('');
}

function hfSection(cls: string, parts: { left: string; center: string; right: string }, ctx: HFContext, pos: string, bw: boolean, themeFont: string): HTMLElement {
  const el = h('div', { class: `ql-pp-hf ${cls}`, style: `${pos};font-family:${fontFamilyCss(themeFont)};` });
  for (const k of ['left', 'center', 'right'] as const) {
    const part = h('div', { class: `ql-pp-hf-${k}` });
    part.innerHTML = hfHtml(parseHeaderFooter(parts[k] || '', ctx), bw);
    el.appendChild(part);
  }
  return el;
}

/** Renders one page of a print job. */
export function renderPage(src: PrintSource, job: PageJob, ctx: HFContext): HTMLElement {
  const sheet = job.sheet;
  const p = sheet.print;
  const [wIn, hIn] = paperInches(p);
  const bw = !!p.blackAndWhite, draft = !!p.draft;
  const page = h('div', { class: 'ql-print-page', 'data-size': `${wIn}x${hIn}`, style: `width:${wIn}in;height:${hIn}in;position:relative;overflow:hidden;background:#fff;color:#000;box-sizing:border-box;` });
  const themeFont = sheet.book.theme?.bodyFont || 'Carlito';
  const hf = headerFooterFor(p, job.index, ctx.page);
  page.appendChild(hfSection('ql-pp-header', hf.header, ctx, `top:${p.margins.header}in;left:${p.margins.left}in;right:${p.margins.right}in`, bw, themeFont));
  page.appendChild(hfSection('ql-pp-footer', hf.footer, ctx, `bottom:${p.margins.footer}in;left:${p.margins.left}in;right:${p.margins.right}in`, bw, themeFont));
  const availW = (wIn - p.margins.left - p.margins.right) * 96, availH = (hIn - p.margins.top - p.margins.bottom) * 96;
  const area = h('div', { class: 'ql-pp-area', style: `position:absolute;left:${p.margins.left}in;top:${p.margins.top}in;width:${availW}px;height:${availH}px;overflow:hidden;` });
  page.appendChild(area);
  if (!job.spec) {
    // Comments and notes listed at the end of the sheet.
    const rows = (job.comments || []).map(({ r, c, cm }) => `<tr><td class="ql-pt-cref">${escHtml(cellName(r, c))}</td><td><b>${escHtml(cm.author)}</b>${cm.note ? ' (note)' : ''}: ${escHtml(cm.text)}${(cm.replies || []).map((rp) => `<div>↳ <b>${escHtml(rp.author)}</b>: ${escHtml(rp.text)}</div>`).join('')}</td></tr>`).join('');
    area.innerHTML = `<table class="ql-pt-comments"><tr><th>Cell</th><th>Comment</th></tr>${rows}</table>`;
    return page;
  }
  const spec = job.spec;
  const lay = job.layout;
  const k = lay.scale / 100;
  const R = pageIndexes(sheet, lay.titleRows, spec.r1, spec.r2, 'row');
  const C = pageIndexes(sheet, lay.titleCols, spec.c1, spec.c2, 'col');
  const headW = p.headings ? HEADING_W : 0, headH = p.headings ? HEADING_H : 0;
  const titleW = C.list.slice(0, C.titles).reduce((s, c) => s + sheet.colWidth(c), 0);
  const titleH = R.list.slice(0, R.titles).reduce((s, r) => s + sheet.rowHeight(r), 0);
  const bodyW = C.list.slice(C.titles).reduce((s, c) => s + sheet.colWidth(c), 0);
  const bodyH = R.list.slice(R.titles).reduce((s, r) => s + sheet.rowHeight(r), 0);
  const contentW = headW + titleW + bodyW, contentH = headH + titleH + bodyH;
  const offX = p.centerH ? Math.max(0, (availW - contentW * k) / 2) : 0;
  const offY = p.centerV ? Math.max(0, (availH - contentH * k) / 2) : 0;
  const inner = h('div', { class: 'ql-pp-inner', style: `position:absolute;left:${offX}px;top:${offY}px;width:${contentW}px;height:${contentH}px;transform:scale(${k});transform-origin:0 0;${bw ? 'filter:grayscale(1);' : ''}` });
  const bg = sheet.background && !bw && !draft ? safeImageSrc(sheet.background) : '';
  const tableHtml = renderCellsTable(src.tab, sheet, R.list, C.list, { bw, draft, grid: p.gridlines && !draft, errors: p.errors, headings: p.headings });
  inner.innerHTML = (bg ? `<div class="ql-pt-bg" style="left:${headW}px;top:${headH}px;width:${titleW + bodyW}px;height:${titleH + bodyH}px;background-image:url(&quot;${escHtml(bg)}&quot;)"></div>` : '') + tableHtml +
    `<div class="ql-pt-objs" style="left:${headW + titleW}px;top:${headH + titleH}px;width:${bodyW}px;height:${bodyH}px">${objectsHtml(src, sheet, spec, bodyW, bodyH, draft)}</div>`;
  area.appendChild(inner);
  return page;
}

/** Renders every page of a print job. */
export function renderPages(src: PrintSource, opts: PrintOptions = {}): HTMLElement[] {
  const jobs = planPrint(src.tab, opts);
  const date = new Date();
  const pages: HTMLElement[] = [];
  const book = opts.what === 'book';
  let n = 0;
  let prevSheet: Sheet | null = null;
  for (const job of jobs) {
    // Page numbers continue through the workbook unless a sheet sets its first page number.
    if (job.sheet !== prevSheet) { prevSheet = job.sheet; n = job.sheet.print.firstPage ?? (book ? n + 1 : 1); } else n++;
    pages.push(renderPage(src, job, { page: n, pages: jobs.length, date, file: src.file, path: src.path, sheet: job.sheet.name }));
  }
  return pages;
}

function sourceFor(app: App): PrintSource {
  const tab = app.tab;
  return { tab, file: tab.fileName || tab.title, path: tab.uri && !tab.uri.startsWith('fsa:') && !tab.uri.startsWith('content:') ? tab.uri : tab.fileName || tab.title, hideInk: !!(tab as any).inkHidden };
}

// --------------------------------------------------------------------------------------------
// Printing

let pendingOptions: PrintOptions | null = null;

const PRINT_CSS = `
#ql-print-root { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
#ql-print-root .ql-print-page, .ql-pp-view .ql-print-page { font-family: "Carlito", sans-serif; }
.ql-pt { border-collapse: collapse; table-layout: fixed; border-spacing: 0; position: relative; }
.ql-pt td, .ql-pt th { padding: 0; overflow: visible; white-space: nowrap; }
.ql-pt-objs { position: absolute; overflow: hidden; pointer-events: none; }
.ql-pt-obj { position: absolute; transform-origin: 50% 50%; }
.ql-pt-obj svg { width: 100%; height: 100%; display: block; }
.ql-pt-bg { position: absolute; background-repeat: repeat; }
.ql-pp-hf { position: absolute; display: flex; gap: 8px; font-size: 10pt; white-space: pre-line; }
.ql-pp-hf > div { flex: 1; min-width: 0; }
.ql-pp-hf-center { text-align: center; }
.ql-pp-hf-right { text-align: right; }
.ql-pt-comments { border-collapse: collapse; width: 100%; font-size: 10pt; }
.ql-pt-comments th, .ql-pt-comments td { border: 1px solid #999; padding: 4px 6px; text-align: left; vertical-align: top; }
.ql-pt-cref { width: 64px; font-weight: 600; }
`;

function ensurePrintCss(): void {
  if (document.getElementById('ql-print-css')) return;
  document.head.appendChild(h('style', { id: 'ql-print-css' }, PRINT_CSS));
}

async function waitForAssets(root: HTMLElement): Promise<void> {
  const imgs = [...root.querySelectorAll('img')];
  await Promise.all(imgs.map((img) => (img.complete ? Promise.resolve() : (img.decode?.() || Promise.resolve()).catch(() => undefined))));
  try { await Promise.race([(document as any).fonts?.ready, new Promise((r) => setTimeout(r, 1500))]); } catch { /* */ }
}

/**
 * Builds the hidden print pages and switches the document into print mode. Returns a cleanup
 * that removes them. Options default to the ones chosen in Print Preview (once), else the active sheet.
 */
export async function preparePrint(app: App, opts?: PrintOptions): Promise<() => void> {
  const o = opts || pendingOptions || {};
  pendingOptions = null;
  ensurePrintCss();
  const pages = renderPages(sourceFor(app), o);
  if (!pages.length) throw new Error("There's nothing to print on this sheet.");
  document.getElementById('ql-print-root')?.remove();
  document.getElementById('ql-print-style')?.remove();
  const root = h('div', { id: 'ql-print-root' });
  for (const p of pages) root.appendChild(p);
  // One @page size per paper size in the job (named pages), the first as the default.
  const sizes = [...new Set(pages.map((p) => p.dataset.size!))];
  const rules = sizes.map((s, i) => { const [w, hh] = s.split('x'); return `@page ql-p${i} { size: ${w}in ${hh}in; margin: 0; } #ql-print-root .ql-print-page[data-size="${s}"] { page: ql-p${i}; }`; });
  const [w0, h0] = sizes[0].split('x');
  const style = h('style', { id: 'ql-print-style' }, `@page { size: ${w0}in ${h0}in; margin: 0; } ${rules.join(' ')}`);
  document.head.appendChild(style);
  document.body.appendChild(root);
  document.body.classList.add('ql-printing');
  await waitForAssets(root);
  return () => {
    root.remove();
    style.remove();
    document.body.classList.remove('ql-printing');
  };
}

// --------------------------------------------------------------------------------------------
// Print Preview

function scalingOf(p: PrintSettings): string {
  if (p.fitWide === 1 && p.fitTall === 1) return 'fitSheet';
  if (p.fitWide === 1 && !p.fitTall) return 'fitCols';
  if (!p.fitWide && p.fitTall === 1) return 'fitRows';
  if (!p.fitWide && !p.fitTall && (p.scale || 100) === 100) return 'none';
  return 'custom';
}

function marginsOf(p: PrintSettings): string {
  for (const [k, v] of Object.entries(MARGIN_PRESETS)) if (JSON.stringify(v.m) === JSON.stringify(p.margins)) return k;
  return 'custom';
}

/** Full-screen Print Preview with settings, page navigation and zoom. */
export async function showPrintPreview(app: App): Promise<void> {
  ensurePrintCss();
  const state: PrintOptions & { page: number; zoom: 'fit' | 'width' | 'actual'; margins: boolean } = { what: 'sheet', ignoreArea: false, page: 0, zoom: 'fit', margins: false };
  let pages: HTMLElement[] = [];

  const view = h('div', { class: 'ql-pp-view', tabindex: '0', 'aria-label': 'Page preview' });
  const counter = h('input', { class: 'ql-pp-counter', type: 'number', min: '1', 'aria-label': 'Page number', 'data-own-enter': '' }) as HTMLInputElement;
  const total = h('span', { class: 'ql-help' });
  const zoomBtn = h('button', { class: 'ql-btn', type: 'button', title: 'Zoom to page' }, icon('zoom_in'), ' Zoom');
  const marginsBtn = h('button', { class: 'ql-btn', type: 'button', title: 'Show margins' }, icon('border_outer'), ' Margins');
  const nav = h('div', { class: 'ql-pp-nav' },
    h('button', { class: 'ql-iconbtn', type: 'button', 'aria-label': 'Previous page', onclick: () => go(state.page - 1) }, icon('chevron_left')),
    counter, total,
    h('button', { class: 'ql-iconbtn', type: 'button', 'aria-label': 'Next page', onclick: () => go(state.page + 1) }, icon('chevron_right')),
    h('span', { class: 'ql-spacer' }), marginsBtn, zoomBtn);
  counter.addEventListener('change', () => go((parseInt(counter.value, 10) || 1) - 1));
  zoomBtn.addEventListener('click', () => { state.zoom = state.zoom === 'fit' ? 'width' : state.zoom === 'width' ? 'actual' : 'fit'; show(); });
  marginsBtn.addEventListener('click', () => { state.margins = !state.margins; marginsBtn.classList.toggle('ql-on', state.margins); show(); });
  view.addEventListener('keydown', (e) => { if (e.key === 'PageDown' || e.key === 'ArrowRight') { e.preventDefault(); go(state.page + 1); } if (e.key === 'PageUp' || e.key === 'ArrowLeft') { e.preventDefault(); go(state.page - 1); } });

  const sel = (label: string, options: [string, string][], value: string, on: (v: string) => void) => {
    const s = h('select', { 'aria-label': label }) as HTMLSelectElement;
    for (const [v, l] of options) s.appendChild(h('option', { value: v }, l));
    s.value = value;
    s.addEventListener('change', () => on(s.value));
    return h('label', { class: 'ql-pp-setting' }, h('span', { class: 'ql-label' }, label), s);
  };
  const side = h('div', { class: 'ql-pp-side' });
  const buildSide = () => {
    const p = app.sheet.print;
    clear(side);
    const ignore = h('input', { type: 'checkbox', checked: !!state.ignoreArea }) as HTMLInputElement;
    ignore.addEventListener('change', () => { state.ignoreArea = ignore.checked; rerender(); });
    side.append(
      h('h3', { class: 'ql-field-heading' }, 'Settings'),
      sel('Print what', [['sheet', 'Print Active Sheet'], ['book', 'Print Entire Workbook'], ['selection', 'Print Selection']], state.what || 'sheet', (v) => { state.what = v as PrintWhat; rerender(); }),
      h('label', { class: 'ql-pp-check' }, ignore, ' Ignore print area'),
      sel('Orientation', [['portrait', 'Portrait'], ['landscape', 'Landscape']], p.orientation, (v) => { app.exec('layout.orientation', v); rerender(); }),
      sel('Paper size', Object.keys(PAPER_SIZES).map((k) => [k, `${k} (${PAPER_SIZES[k][0]}" × ${PAPER_SIZES[k][1]}")`]), p.paper, (v) => { app.exec('layout.size', v); rerender(); }),
      sel('Margins', [['normal', 'Normal'], ['wide', 'Wide'], ['narrow', 'Narrow'], ['custom', 'Custom…']], marginsOf(p), (v) => { if (v === 'custom') void setup('margins'); else { app.exec('layout.margins', v); rerender(); } }),
      sel('Scaling', [['none', 'No Scaling'], ['fitSheet', 'Fit Sheet on One Page'], ['fitCols', 'Fit All Columns on One Page'], ['fitRows', 'Fit All Rows on One Page'], ['custom', 'Custom Scaling…']], scalingOf(p), (v) => { if (v === 'custom') void setup('page'); else { app.exec('layout.scaling', v); rerender(); } }),
      h('button', { class: 'ql-link', type: 'button', onclick: () => void setup('page') }, 'Page Setup…'),
    );
  };
  const setup = async (section: string) => { await app.exec('layout.pageSetup', section); rerender(); };

  const go = (i: number) => { state.page = Math.max(0, Math.min(pages.length - 1, i)); show(); };
  const show = () => {
    clear(view);
    counter.value = String(pages.length ? state.page + 1 : 0);
    counter.max = String(pages.length);
    total.textContent = `of ${pages.length}`;
    if (!pages.length) { view.appendChild(h('p', { class: 'ql-pp-empty' }, "There's nothing to print. Type something on the sheet, or choose what to print in Settings.")); return; }
    const page = pages[state.page];
    const [wIn, hIn] = (page.dataset.size || '8.5x11').split('x').map(Number);
    const pw = wIn * 96, ph = hIn * 96;
    const vw = Math.max(100, view.clientWidth - 24), vh = Math.max(100, view.clientHeight - 24);
    const z = state.zoom === 'actual' ? 1 : state.zoom === 'width' ? vw / pw : Math.min(vw / pw, vh / ph);
    const clone = page.cloneNode(true) as HTMLElement;
    if (state.margins) clone.classList.add('ql-pp-margins');
    const wrap = h('div', { class: 'ql-pp-pagewrap', style: `width:${pw * z}px;height:${ph * z}px` }, clone);
    clone.style.transform = `scale(${z})`;
    clone.style.transformOrigin = '0 0';
    view.appendChild(wrap);
    zoomBtn.lastChild!.textContent = state.zoom === 'fit' ? ' Zoom' : state.zoom === 'width' ? ' 100%' : ' Fit';
  };
  const rerender = () => {
    try { pages = renderPages(sourceFor(app), state); } catch (e) { pages = []; toast(`Print Preview: ${(e as Error).message}`); }
    state.page = Math.min(state.page, Math.max(0, pages.length - 1));
    buildSide();
    show();
  };

  const root = h('div', { class: 'ql-pp' + (isPhone() ? ' ql-pp-phone' : '') }, side, h('div', { class: 'ql-pp-main' }, nav, view));
  const ro = typeof ResizeObserver !== 'undefined' ? new ResizeObserver(() => show()) : null;
  const result = await openDialog({
    title: 'Print Preview', full: true, content: root, ok: 'Print', cancel: 'Close',
    buttons: [{ label: 'Page Setup…', onClick: () => void setup('page') }],
    onOpen: () => { rerender(); ro?.observe(view); setTimeout(() => { show(); view.focus(); }, 60); },
  });
  ro?.disconnect();
  if (result) {
    pendingOptions = { what: state.what, ignoreArea: state.ignoreArea };
    app.exec('file.print');
  }
}
