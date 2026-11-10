/**
 * Grid painter: draws the four pane quadrants (frozen corner, frozen rows, frozen columns, scrolling area),
 * then headers, selection, previews and decorations.
 *
 * Per cell: background (cell fill → table band → conditional fill), gridlines are covered by fills,
 * text with spreadsheet alignment rules and overflow into empty neighbors, data bars, icons, sparklines,
 * then borders, comment and error indicators.
 */
import type { Grid } from './grid';
import type { CellStyle, Border, Scalar, Cell } from '../model/types';
import { CellError } from '../model/types';
import { cellKey, colName, MAX_COLS, MAX_ROWS, type Range } from '../model/address';
import { formatValue, formatGeneral } from '../model/numfmt';
import { tableCellLook } from '../model/tableStyles';
import { ICON_SETS, type CFResult } from './cf';
import { fontFamilyCss } from '../fonts';
import { cellKind, MAP_STYLE } from '../heritage/audit';

/** Drawn radius of a touch selection handle, in CSS pixels. */
const HANDLE_R = 9;

export interface PaintCache { text: Map<string, number>; version: number }

interface Theme {
  bg: string; text: string; grid: string; headerBg: string; headerText: string; headerSel: string; headerSelText: string; headerLine: string;
  accent: string; selFill: string; dark: boolean; frozenLine: string; pageBreak: string;
}

let themeCache: { key: string; t: Theme } | null = null;

function theme(g: Grid): Theme {
  const dark = document.documentElement.dataset.theme === 'dark' && g.app.prefs.darkSheet;
  const key = `${document.documentElement.dataset.theme}:${dark}:${g.sheet.gridColor}`;
  if (themeCache?.key === key) return themeCache.t;
  const css = getComputedStyle(document.documentElement);
  const v = (n: string, d: string) => css.getPropertyValue(n).trim() || d;
  const uiDark = document.documentElement.dataset.theme === 'dark';
  const t: Theme = {
    bg: dark ? '#1f1f1f' : '#ffffff',
    text: dark ? '#e8e8e8' : '#000000',
    grid: g.sheet.gridColor || (dark ? '#3a3a3a' : '#e1e1e1'),
    headerBg: v('--ql-header-bg', uiDark ? '#2a2a2a' : '#f3f3f3'),
    headerText: v('--ql-header-text', uiDark ? '#bdbdbd' : '#444444'),
    headerSel: v('--ql-header-sel', uiDark ? '#2f4a3a' : '#d2e8da'),
    headerSelText: v('--ql-accent', '#1e7145'),
    headerLine: uiDark ? '#444' : '#d0d0d0',
    accent: v('--ql-accent', '#1e7145'),
    selFill: dark ? 'rgba(120,210,160,0.2)' : 'rgba(30,113,69,0.12)',
    dark,
    frozenLine: uiDark ? '#777' : '#9a9a9a',
    pageBreak: '#2f6fb0',
  };
  themeCache = { key, t };
  return t;
}

const AUTO = (t: Theme) => t.text;

export function fontCss(st: CellStyle, z: number, base: CellStyle): string {
  const size = (st.size || base.size || 11) * (4 / 3) * z;
  const fam = fontFamilyCss(st.font || base.font || 'Carlito');
  const sz = st.vertAlign ? size * 0.7 : size;
  return `${st.italic ? 'italic ' : ''}${st.bold ? '700 ' : '400 '}${sz.toFixed(2)}px ${fam}`;
}

function measure(g: Grid, ctx: CanvasRenderingContext2D, font: string, text: string): number {
  const key = font + '\u0000' + text;
  let w = g.cache.text.get(key);
  if (w == null) {
    ctx.font = font;
    w = ctx.measureText(text).width;
    if (g.cache.text.size > 20000) g.cache.text.clear();
    g.cache.text.set(key, w);
  }
  return w;
}

export function paint(g: Grid): void {
  const ctx = g.ctx;
  const t = theme(g);
  const s = g.sheet;
  const dpr = g.dpr;
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  ctx.fillStyle = t.bg;
  ctx.fillRect(0, 0, g.width, g.height);
  const fr = s.freeze.rows, fc = s.freeze.cols;
  const hx = g.headerW, hy = g.headerH;
  const fw = g.frozenW, fh = g.frozenH;
  const vis = g.visible();
  const lastR = Math.min(MAX_ROWS - 1, vis.r2 + 1), lastC = Math.min(MAX_COLS - 1, vis.c2 + 1);
  // Quadrants: [rows range, cols range, clip]
  const quads: { r1: number; r2: number; c1: number; c2: number; clip: [number, number, number, number] }[] = [];
  quads.push({ r1: vis.r1, r2: lastR, c1: vis.c1, c2: lastC, clip: [hx + fw, hy + fh, g.width - hx - fw, g.height - hy - fh] });
  if (fr) quads.push({ r1: 0, r2: fr - 1, c1: vis.c1, c2: lastC, clip: [hx + fw, hy, g.width - hx - fw, fh] });
  if (fc) quads.push({ r1: vis.r1, r2: lastR, c1: 0, c2: fc - 1, clip: [hx, hy + fh, fw, g.height - hy - fh] });
  if (fr && fc) quads.push({ r1: 0, r2: fr - 1, c1: 0, c2: fc - 1, clip: [hx, hy, fw, fh] });
  for (const q of quads) {
    if (q.clip[2] <= 0 || q.clip[3] <= 0) continue;
    ctx.save();
    ctx.beginPath();
    ctx.rect(...q.clip);
    ctx.clip();
    paintCells(g, ctx, t, q.r1, q.r2, q.c1, q.c2);
    paintSelection(g, ctx, t);
    ctx.restore();
  }
  if (fr || fc || s.split) {
    ctx.strokeStyle = t.frozenLine;
    ctx.lineWidth = 1;
    ctx.beginPath();
    if (fr) { ctx.moveTo(0, hy + fh + 0.5); ctx.lineTo(g.width, hy + fh + 0.5); }
    if (fc) { ctx.moveTo(hx + fw + 0.5, 0); ctx.lineTo(hx + fw + 0.5, g.height); }
    ctx.stroke();
  }
  if (s.showHeadings) paintHeaders(g, ctx, t, vis);
  paintPreviews(g, ctx, t);
}

// ----------------------------------------------------------------------------------------------
// Cells

interface CellPaint { r: number; c: number; x: number; y: number; w: number; h: number; cell?: Cell; st: CellStyle; cf: CFResult | null; table: ReturnType<typeof tableCellLook> }

function paintCells(g: Grid, ctx: CanvasRenderingContext2D, t: Theme, r1: number, r2: number, c1: number, c2: number): void {
  const s = g.sheet;
  const book = s.book;
  const base = book.styles[0] || {};
  const z = g.zoom;
  const version = book.version + g.tab.engine.evaluations;
  // Merges intersecting this area
  const merges = s.merges.filter((m) => m.r1 <= r2 && m.r2 >= r1 && m.c1 <= c2 && m.c2 >= c1);
  const inMerge = (r: number, c: number) => merges.find((m) => r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2);
  // Gridlines
  if (s.showGrid && g.tab.viewMode !== 'pageLayout' || (s.showGrid && g.tab.viewMode === 'pageLayout')) {
    ctx.strokeStyle = t.grid;
    ctx.lineWidth = 1;
    ctx.beginPath();
    const yTop = g.rowY(r1), yBot = g.rowY(r2) + g.rowH(r2);
    const xLeft = g.colX(c1), xRight = g.colX(c2) + g.colW(c2);
    for (let c = c1; c <= c2; c++) { const w = g.colW(c); if (!w) continue; const x = Math.round(g.colX(c) + w) - 0.5; ctx.moveTo(x, yTop); ctx.lineTo(x, yBot); }
    for (let r = r1; r <= r2; r++) { const h = g.rowH(r); if (!h) continue; const y = Math.round(g.rowY(r) + h) - 0.5; ctx.moveTo(xLeft, y); ctx.lineTo(xRight, y); }
    ctx.stroke();
  }
  // Pass 1: fills. Collect paint info.
  const rowsInfo: CellPaint[][] = [];
  const tables = s.tables.filter((tb) => tb.range.r1 <= r2 && tb.range.r2 >= r1 && tb.range.c1 <= c2 && tb.range.c2 >= c1);
  for (let r = r1; r <= r2; r++) {
    const rh = g.rowH(r);
    const row: CellPaint[] = [];
    rowsInfo.push(row);
    if (!rh) continue;
    for (let c = c1; c <= c2; c++) {
      const cw = g.colW(c);
      if (!cw) continue;
      const m = inMerge(r, c);
      if (m && (r !== m.r1 || c !== m.c1)) {
        // The merge's top-left may be outside the painted area: paint it once from the first visible cell.
        if (!(r === Math.max(m.r1, r1) && c === Math.max(m.c1, c1))) continue;
      }
      const rr = m ? m.r1 : r, cc = m ? m.c1 : c;
      const cell = s.cells.get(cellKey(rr, cc));
      const sid = cell?.s ?? s.rowStyle.get(rr) ?? s.colStyle.get(cc) ?? 0;
      let st = book.styles[sid] || {};
      const cf = s.cf.length ? g.tab.cf.evaluate(s, rr, cc, version) : null;
      let table = null;
      for (const tb of tables) { const l = tableCellLook(tb, rr, cc, book.theme); if (l) { table = l; break; } }
      if (cf?.style) st = { ...st, ...cf.style };
      const rect = m ? g.rangeRect(m) : { x: g.colX(c), y: g.rowY(r), w: cw, h: rh };
      const fill = cf?.fill || cf?.style?.fill || st.fill || table?.fill;
      if (fill || st.gradient || st.pattern) {
        if (st.gradient && !cf?.fill) {
          const gr = ctx.createLinearGradient(rect.x, rect.y, (st.gradient.angle ?? 0) === 90 ? rect.x : rect.x + rect.w, (st.gradient.angle ?? 0) === 90 ? rect.y + rect.h : rect.y);
          gr.addColorStop(0, st.gradient.from); gr.addColorStop(1, st.gradient.to);
          ctx.fillStyle = gr;
        } else ctx.fillStyle = fill || st.pattern?.color || t.bg;
        ctx.fillRect(Math.floor(rect.x), Math.floor(rect.y), Math.ceil(rect.w), Math.ceil(rect.h));
        if (st.pattern && st.pattern.type !== 'solid' && st.pattern.color) paintPattern(ctx, rect, st.pattern.type, st.pattern.color);
      } else if (m) {
        ctx.fillStyle = t.bg;
        ctx.fillRect(Math.floor(rect.x) + 1, Math.floor(rect.y) + 1, Math.ceil(rect.w) - 2, Math.ceil(rect.h) - 2);
        if (s.showGrid) { ctx.strokeStyle = t.grid; ctx.lineWidth = 1; ctx.strokeRect(Math.round(rect.x) - 0.5, Math.round(rect.y) - 0.5, Math.round(rect.w), Math.round(rect.h)); }
      }
      if (cf?.bar) paintDataBar(ctx, rect, cf.bar);
      row.push({ r: rr, c: cc, ...rect, cell, st, cf, table });
    }
  }
  // Pass 2: text (with overflow into empty neighbors)
  const occupied = (r: number, c: number) => { const x = s.cells.get(cellKey(r, c)); return !!x && (x.v != null && x.v !== '' || x.f != null); };
  const mapView = g.tab.mapView;
  for (const row of rowsInfo) {
    for (const p of row) {
      const cell = p.cell;
      if (mapView) { paintMapCell(g, ctx, t, p); continue; }
      if (!cell || (cell.v == null && cell.f == null)) {
        if (cell?.v == null && cell?.f == null) { paintSparkline(g, ctx, p); }
        continue;
      }
      if (p.cf?.hideValue && !p.cf.icon) continue;
      paintText(g, ctx, t, p, base, z, occupied, c1, c2);
    }
  }
  // Text that overflows in from the left of the painted area.
  if (c1 > 0 && !mapView) {
    for (let r = r1; r <= r2; r++) {
      if (!g.rowH(r)) continue;
      for (let c = Math.max(0, c1 - 15); c < c1; c++) {
        const cell = s.cells.get(cellKey(r, c));
        if (!cell || typeof cell.v !== 'string' || !cell.v) continue;
        const st = book.styles[cell.s ?? 0] || {};
        if (st.wrap || (st.hAlign && st.hAlign !== 'general' && st.hAlign !== 'left') || inMerge(r, c) || s.mergeAt(r, c)) continue;
        paintText(g, ctx, t, { r, c, x: g.colX(c), y: g.rowY(r), w: g.colW(c), h: g.rowH(r), cell, st, cf: null, table: null }, base, z, occupied, c1, c2, true);
      }
    }
  }
  // Pass 3: borders
  for (const row of rowsInfo) for (const p of row) paintBorders(g, ctx, p, t);
  // Indicators: comments, errors, spill outline, invalid-data circles
  for (const row of rowsInfo) for (const p of row) {
    const k = cellKey(p.r, p.c);
    const cm = s.comments.get(k);
    if (cm) {
      ctx.fillStyle = cm.note ? '#c00000' : cm.resolved ? '#9e9e9e' : '#7b3fa0';
      ctx.beginPath();
      ctx.moveTo(p.x + p.w - 7 * Math.min(1, z + 0.2), p.y + 1); ctx.lineTo(p.x + p.w - 1, p.y + 1); ctx.lineTo(p.x + p.w - 1, p.y + 7 * Math.min(1, z + 0.2));
      ctx.fill();
    }
    if (p.cf?.icon) paintIcon(ctx, p, p.cf.icon, z);
  }
  for (const c of g.tab.circles) {
    if (c.r < r1 || c.r > r2 || c.c < c1 || c.c > c2) continue;
    const rc = g.cellRect(c.r, c.c);
    ctx.strokeStyle = '#d00000'; ctx.lineWidth = 1.5;
    ctx.beginPath(); ctx.ellipse(rc.x + rc.w / 2, rc.y + rc.h / 2, rc.w / 2 + 3, rc.h / 2 + 2, 0, 0, Math.PI * 2); ctx.stroke();
  }
  // Filter / table header buttons
  for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) {
    const fb = g.filterButtonAt(r, c);
    if (!fb) continue;
    const rc = g.cellRect(r, c);
    const bs = Math.min(18, rc.h - 2);
    const bx = rc.x + rc.w - bs - 2, by = rc.y + rc.h - bs - 1;
    const active = fb.source === 'filter' ? !!s.filter?.columns[fb.col] : false;
    ctx.fillStyle = t.dark ? '#333' : '#f7f7f7';
    ctx.strokeStyle = t.dark ? '#666' : '#a6a6a6';
    ctx.lineWidth = 1;
    ctx.fillRect(bx, by, bs, bs); ctx.strokeRect(bx + 0.5, by + 0.5, bs - 1, bs - 1);
    ctx.fillStyle = t.dark ? '#ddd' : '#444';
    ctx.beginPath();
    const cx = bx + bs / 2, cy = by + bs / 2;
    if (active) { ctx.moveTo(cx - 4, cy - 3); ctx.lineTo(cx + 4, cy - 3); ctx.lineTo(cx + 1, cy + 1); ctx.lineTo(cx + 1, cy + 4); ctx.lineTo(cx - 1, cy + 4); ctx.lineTo(cx - 1, cy + 1); ctx.closePath(); }
    else { ctx.moveTo(cx - 3.5, cy - 1.5); ctx.lineTo(cx + 3.5, cy - 1.5); ctx.lineTo(cx, cy + 2.5); ctx.closePath(); }
    ctx.fill();
  }
  // Page breaks (Page Break Preview, or after printing/preview)
  if (g.tab.viewMode !== 'normal' || (g as any).showPageBreaks) paintPageBreaks(g, ctx, t, r1, r2, c1, c2);
}

function paintPattern(ctx: CanvasRenderingContext2D, rect: { x: number; y: number; w: number; h: number }, type: string, color: string): void {
  ctx.save();
  ctx.beginPath(); ctx.rect(rect.x, rect.y, rect.w, rect.h); ctx.clip();
  ctx.strokeStyle = color; ctx.lineWidth = 1;
  ctx.beginPath();
  const step = /light/i.test(type) ? 6 : 3;
  if (/Horizontal/i.test(type)) for (let y = rect.y; y < rect.y + rect.h; y += step) { ctx.moveTo(rect.x, y + 0.5); ctx.lineTo(rect.x + rect.w, y + 0.5); }
  else if (/Vertical/i.test(type)) for (let x = rect.x; x < rect.x + rect.w; x += step) { ctx.moveTo(x + 0.5, rect.y); ctx.lineTo(x + 0.5, rect.y + rect.h); }
  else for (let d = -rect.h; d < rect.w; d += step) { ctx.moveTo(rect.x + d, rect.y + rect.h); ctx.lineTo(rect.x + d + rect.h, rect.y); }
  ctx.stroke();
  ctx.restore();
}

function paintDataBar(ctx: CanvasRenderingContext2D, rect: { x: number; y: number; w: number; h: number }, bar: NonNullable<CFResult['bar']>): void {
  const pad = 2;
  const full = rect.w - pad * 2;
  const axisX = rect.x + pad + full * (bar.axis || 0);
  const len = full * bar.pct;
  const x = bar.neg ? axisX - len : axisX;
  const y = rect.y + pad + 1, h = rect.h - pad * 2 - 2;
  const color = bar.neg ? '#ff4b4b' : bar.color;
  if (bar.gradient) {
    const gr = ctx.createLinearGradient(x, 0, x + Math.max(1, len), 0);
    gr.addColorStop(0, color); gr.addColorStop(1, '#ffffff');
    ctx.fillStyle = gr;
  } else ctx.fillStyle = color;
  ctx.fillRect(x, y, len, h);
  ctx.strokeStyle = color; ctx.lineWidth = 1;
  ctx.strokeRect(x + 0.5, y + 0.5, Math.max(0, len - 1), h - 1);
  if (bar.axis) { ctx.strokeStyle = '#000'; ctx.setLineDash([2, 2]); ctx.beginPath(); ctx.moveTo(axisX + 0.5, rect.y); ctx.lineTo(axisX + 0.5, rect.y + rect.h); ctx.stroke(); ctx.setLineDash([]); }
}

function paintIcon(ctx: CanvasRenderingContext2D, p: CellPaint, icon: { set: string; index: number }, z: number): void {
  const def = ICON_SETS[icon.set];
  if (!def) return;
  const name = def.icons[icon.index] || 'circle';
  const color = def.colors[icon.index] || '#666';
  const size = Math.min(p.h - 2, 16 * z);
  ctx.font = `${size}px "Material Symbols Outlined"`;
  ctx.fillStyle = color;
  ctx.textBaseline = 'middle';
  ctx.textAlign = 'left';
  ctx.fillText(name, p.x + 2, p.y + p.h / 2);
}

/** Map View: a tinted box and a glyph for the cell's kind (" label, # number, + formula, ≠ odd formula). */
function paintMapCell(g: Grid, ctx: CanvasRenderingContext2D, t: Theme, p: CellPaint): void {
  const kind = cellKind(g.sheet, p.r, p.c, p.cell);
  if (kind === 'empty') return;
  const look = MAP_STYLE[kind];
  ctx.fillStyle = t.dark ? look.dark : look.fill;
  ctx.fillRect(Math.floor(p.x) + 1, Math.floor(p.y) + 1, Math.ceil(p.w) - 2, Math.ceil(p.h) - 2);
  ctx.fillStyle = t.dark ? '#f0f0f0' : '#222';
  ctx.font = `700 ${Math.max(8, Math.round(13 * g.zoom))}px ui-monospace, Cousine, monospace`;
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.fillText(look.glyph, p.x + p.w / 2, p.y + p.h / 2 + 1);
}

function displayText(g: Grid, cell: Cell, st: CellStyle, widthPx: number, z: number): { text: string; color?: string; isNumber: boolean; fill?: string } {
  const s = g.sheet;
  if (s.showFormulas && cell.f != null) return { text: '=' + cell.f, isNumber: false };
  const v = cell.v ?? null;
  if (typeof v === 'number' && v === 0 && !s.showZeros) return { text: '', isNumber: true };
  const charW = 7 * z;
  const width = Math.max(1, Math.floor((widthPx - 4) / charW));
  const f = formatValue(v, st.numFmt, { date1904: s.book.calc.date1904, width: Math.min(11, width) });
  return { text: f.text, color: f.color, isNumber: !!f.isNumber || typeof v === 'number', fill: f.fill };
}

/** Box geometry of an in-cell checkbox (shared by painting and hit testing). */
export function checkboxRect(p: { x: number; y: number; w: number; h: number }, st: CellStyle, z: number): { x: number; y: number; s: number } {
  const s = Math.max(6, Math.min(p.h - 4, p.w - 4, 14 * z));
  const x = st.hAlign === 'left' ? p.x + 3 * z : st.hAlign === 'right' ? p.x + p.w - s - 3 * z : p.x + (p.w - s) / 2;
  return { x, y: p.y + (p.h - s) / 2, s };
}

function paintCheckbox(ctx: CanvasRenderingContext2D, t: Theme, p: CellPaint, on: boolean, st: CellStyle, z: number): void {
  const b = checkboxRect(p, st, z);
  const r = Math.max(1.5, b.s * 0.18);
  ctx.save();
  ctx.beginPath();
  ctx.moveTo(b.x + r, b.y); ctx.arcTo(b.x + b.s, b.y, b.x + b.s, b.y + b.s, r); ctx.arcTo(b.x + b.s, b.y + b.s, b.x, b.y + b.s, r); ctx.arcTo(b.x, b.y + b.s, b.x, b.y, r); ctx.arcTo(b.x, b.y, b.x + b.s, b.y, r); ctx.closePath();
  ctx.fillStyle = on ? (st.color || t.accent) : t.dark ? '#1e1e1e' : '#ffffff';
  ctx.fill();
  ctx.lineWidth = Math.max(1, z);
  ctx.strokeStyle = on ? (st.color || t.accent) : t.dark ? '#9a9a9a' : '#6f6f6f';
  ctx.stroke();
  if (on) {
    ctx.strokeStyle = '#ffffff'; ctx.lineWidth = Math.max(1.5, b.s * 0.14); ctx.lineCap = 'round'; ctx.lineJoin = 'round';
    ctx.beginPath(); ctx.moveTo(b.x + b.s * 0.24, b.y + b.s * 0.52); ctx.lineTo(b.x + b.s * 0.43, b.y + b.s * 0.7); ctx.lineTo(b.x + b.s * 0.77, b.y + b.s * 0.31); ctx.stroke();
  }
  ctx.restore();
}

function paintText(g: Grid, ctx: CanvasRenderingContext2D, t: Theme, p: CellPaint, base: CellStyle, z: number, occupied: (r: number, c: number) => boolean, c1: number, c2: number, overflowOnly = false): void {
  const cell = p.cell!;
  const st = p.st;
  // In-cell checkbox (Insert ▸ Checkbox): TRUE/FALSE drawn as a box.
  if (st.checkbox && typeof cell.v === 'boolean' && !(g.sheet.showFormulas && cell.f != null)) { if (!overflowOnly) paintCheckbox(ctx, t, p, cell.v, st, z); return; }
  const effective: CellStyle = { ...st };
  if (p.table?.bold && st.bold == null) effective.bold = true;
  const disp = displayText(g, cell, effective, p.w, z);
  let text = disp.text;
  if (!text && !p.cf?.icon) return;
  const font = fontCss(effective, z, base);
  const color = p.cf?.style?.color || disp.color || st.color || p.table?.color || (cell.link ? (t.dark ? '#7fb2ff' : '#0563c1') : AUTO(t));
  const v = cell.v;
  const isNum = disp.isNumber && !(g.sheet.showFormulas && cell.f != null);
  let align = st.hAlign && st.hAlign !== 'general' ? st.hAlign : (isNum ? 'right' : typeof v === 'boolean' || v instanceof CellError ? 'center' : 'left');
  if (g.sheet.rtl && (!st.hAlign || st.hAlign === 'general')) align = align === 'left' ? 'right' : align === 'right' ? 'left' : align;
  const padX = 2 * z + (st.indent ? st.indent * 9 * z : 0);
  const iconPad = p.cf?.icon ? Math.min(p.h - 2, 16 * z) + 2 : 0;
  let textW = measure(g, ctx, font, text);
  const avail = p.w - padX * 2 - iconPad;
  // Numbers that do not fit show ### (General numbers first try fewer digits).
  if (isNum && textW > avail && !st.wrap && !st.shrink && typeof v === 'number' && !g.sheet.showFormulas) {
    if (!st.numFmt || st.numFmt === 'General') {
      for (let wch = 10; wch >= 1 && textW > avail; wch--) { text = formatGeneral(v, wch); textW = measure(g, ctx, font, text); }
    }
    if (textW > avail) { const n = Math.max(1, Math.floor(avail / measure(g, ctx, font, '#'))); text = '#'.repeat(n); textW = measure(g, ctx, font, text); }
  }
  ctx.font = font;
  ctx.fillStyle = color;
  ctx.textBaseline = 'alphabetic';
  const size = (effective.size || base.size || 11) * (4 / 3) * z;
  const lineH = size * 1.2;
  // Vertical position
  const valign = st.vAlign || 'bottom';
  // Clip region: own cell, or extended across empty neighbors for overflowing text.
  let clipX = p.x, clipW = p.w;
  const canOverflow = !isNum && !st.wrap && !st.shrink && !g.sheet.mergeAt(p.r, p.c) && typeof v === 'string' && !st.rotate && align !== 'fill' && align !== 'justify' && align !== 'distributed';
  if (canOverflow && textW > avail) {
    if (align === 'left' || align === 'center') {
      let c = p.c + 1, right = p.x + p.w;
      while (right < p.x + padX + textW + (align === 'center' ? -(textW - avail) / 2 : 0) + 4 && c < MAX_COLS && !occupied(p.r, c) && !g.sheet.mergeAt(p.r, c) && c <= c2 + 40) { right += g.colW(c); c++; }
      clipW = right - p.x;
    }
    if (align === 'right' || align === 'center') {
      let c = p.c - 1, left = p.x;
      while (left > p.x + p.w - padX - textW - 4 && c >= 0 && !occupied(p.r, c) && !g.sheet.mergeAt(p.r, c) && c >= c1 - 40) { left -= g.colW(c); c--; }
      clipW += p.x - left; clipX = left;
    }
  } else if (overflowOnly) return;
  if (overflowOnly && clipX + clipW <= g.colX(c1)) return;
  ctx.save();
  ctx.beginPath();
  ctx.rect(clipX, p.y, clipW, p.h);
  ctx.clip();
  let lines: string[] = [text];
  let scale = 1;
  if (st.wrap && !isNum) lines = wrapText(g, ctx, font, text, avail);
  else if (text.includes('\n')) lines = st.wrap ? lines : [text.replace(/\n/g, ' ')];
  if (st.shrink && textW > avail && avail > 0) scale = avail / textW;
  if (scale !== 1) { ctx.font = font.replace(/([\d.]+)px/, (_m, n) => `${(parseFloat(n) * scale).toFixed(2)}px`); }
  const totalH = lines.length * lineH;
  let y0: number;
  if (valign === 'top') y0 = p.y + 2 + size * 0.95;
  else if (valign === 'middle' || valign === 'justify' || valign === 'distributed') y0 = p.y + (p.h - totalH) / 2 + size * 0.95;
  else y0 = p.y + p.h - totalH - 1 * z + size * 0.95;
  if (st.vertAlign === 'superscript') y0 -= size * 0.35;
  if (st.vertAlign === 'subscript') y0 += size * 0.15;
  if (st.rotate && st.rotate !== 0) {
    paintRotated(ctx, lines.join(' '), p, st.rotate, size, align);
    ctx.restore();
    return;
  }
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const lw = measure(g, ctx, ctx.font, line) * (scale !== 1 ? 1 : 1);
    let x: number;
    if (align === 'right') x = p.x + p.w - padX - lw;
    else if (align === 'center' || align === 'centerContinuous') x = p.x + (p.w - lw) / 2;
    else x = p.x + padX + iconPad;
    if (align === 'fill' && line) {
      const reps = Math.max(1, Math.floor((p.w - padX * 2) / Math.max(1, lw)));
      ctx.fillText(line.repeat(reps), p.x + padX, y0 + i * lineH);
      continue;
    }
    const y = y0 + i * lineH;
    if (cell.rich && cell.rich.length && lines.length === 1 && !g.sheet.showFormulas) paintRich(g, ctx, cell, effective, base, z, x, y, color);
    else ctx.fillText(line, x, y);
    // Accounting fill character (*x): repeat to fill the gap
    if (disp.fill && i === 0 && lines.length === 1) {
      const fw = measure(g, ctx, ctx.font, disp.fill);
      void fw;
    }
    const ul = st.underline;
    if (ul || st.strike) {
      ctx.strokeStyle = color; ctx.lineWidth = Math.max(1, size / 14);
      ctx.beginPath();
      if (ul) { const uy = y + size * 0.12; ctx.moveTo(x, uy); ctx.lineTo(x + lw, uy); if (ul === 'double' || ul === 'doubleAccounting') { ctx.moveTo(x, uy + size * 0.12); ctx.lineTo(x + lw, uy + size * 0.12); } }
      if (st.strike) { const sy = y - size * 0.3; ctx.moveTo(x, sy); ctx.lineTo(x + lw, sy); }
      ctx.stroke();
    }
  }
  ctx.restore();
}

function paintRich(g: Grid, ctx: CanvasRenderingContext2D, cell: Cell, st: CellStyle, base: CellStyle, z: number, x: number, y: number, color: string): void {
  let cx = x;
  for (const run of cell.rich!) {
    const rs = { ...st, ...(run.s || {}) };
    const f = fontCss(rs, z, base);
    ctx.font = f;
    ctx.fillStyle = run.s?.color || color;
    ctx.fillText(run.t, cx, y);
    cx += measure(g, ctx, f, run.t);
  }
}

function paintRotated(ctx: CanvasRenderingContext2D, text: string, p: CellPaint, rot: number, size: number, align: string): void {
  ctx.save();
  if (rot === 255) {
    // vertical stacked letters
    const chars = [...text];
    let y = p.y + size + 2;
    ctx.textAlign = 'center';
    for (const ch of chars) { ctx.fillText(ch, p.x + p.w / 2, y); y += size * 1.05; }
    ctx.restore();
    return;
  }
  const ang = (-rot * Math.PI) / 180;
  ctx.translate(p.x + (align === 'right' ? p.w - 4 : align === 'center' ? p.w / 2 : 4), p.y + p.h - 4);
  ctx.rotate(ang);
  ctx.textAlign = align === 'center' ? 'center' : 'left';
  ctx.fillText(text, 0, 0);
  ctx.restore();
}

function wrapText(g: Grid, ctx: CanvasRenderingContext2D, font: string, text: string, width: number): string[] {
  const out: string[] = [];
  for (const para of text.split('\n')) {
    const words = para.split(/(\s+)/);
    let line = '';
    for (const w of words) {
      const test = line + w;
      if (line && measure(g, ctx, font, test.trimEnd()) > width) {
        out.push(line.trimEnd());
        line = w.trimStart();
        // break very long words
        while (measure(g, ctx, font, line) > width && line.length > 1) {
          let k = line.length - 1;
          while (k > 1 && measure(g, ctx, font, line.slice(0, k)) > width) k--;
          out.push(line.slice(0, k));
          line = line.slice(k);
        }
      } else line = test;
    }
    out.push(line.trimEnd());
  }
  return out;
}

const BORDER_W: Record<Border['style'], number> = { thin: 1, hair: 1, dotted: 1, dashed: 1, dashDot: 1, mediumDashed: 2, medium: 2, thick: 3, double: 3 };

function strokeBorder(ctx: CanvasRenderingContext2D, b: Border, x1: number, y1: number, x2: number, y2: number, t: Theme): void {
  const w = BORDER_W[b.style] || 1;
  ctx.strokeStyle = b.color || (t.dark ? '#d0d0d0' : '#000000');
  ctx.lineWidth = w;
  ctx.setLineDash(b.style === 'dotted' || b.style === 'hair' ? [1, 2] : b.style === 'dashed' || b.style === 'mediumDashed' ? [4, 2] : b.style === 'dashDot' ? [5, 2, 1, 2] : []);
  const off = w % 2 ? 0.5 : 0;
  if (b.style === 'double') {
    ctx.lineWidth = 1;
    ctx.beginPath();
    if (y1 === y2) { ctx.moveTo(x1, y1 - 1 + 0.5); ctx.lineTo(x2, y2 - 1 + 0.5); ctx.moveTo(x1, y1 + 1 + 0.5); ctx.lineTo(x2, y2 + 1 + 0.5); }
    else { ctx.moveTo(x1 - 1 + 0.5, y1); ctx.lineTo(x2 - 1 + 0.5, y2); ctx.moveTo(x1 + 1 + 0.5, y1); ctx.lineTo(x2 + 1 + 0.5, y2); }
    ctx.stroke();
    return;
  }
  ctx.beginPath();
  if (y1 === y2) { ctx.moveTo(x1, Math.round(y1) + off); ctx.lineTo(x2, Math.round(y2) + off); }
  else { ctx.moveTo(Math.round(x1) + off, y1); ctx.lineTo(Math.round(x2) + off, y2); }
  ctx.stroke();
  ctx.setLineDash([]);
}

function paintBorders(g: Grid, ctx: CanvasRenderingContext2D, p: CellPaint, t: Theme): void {
  const st = p.st;
  const x1 = p.x - 0.5, y1 = p.y - 0.5, x2 = p.x + p.w - 0.5, y2 = p.y + p.h - 0.5;
  if (st.bt) strokeBorder(ctx, st.bt, x1, y1, x2, y1, t);
  if (st.bb) strokeBorder(ctx, st.bb, x1, y2, x2, y2, t);
  else if (p.table?.borderBottom) strokeBorder(ctx, { style: 'thin', color: p.table.borderBottom }, x1, y2, x2, y2, t);
  if (st.bl) strokeBorder(ctx, st.bl, x1, y1, x1, y2, t);
  if (st.br) strokeBorder(ctx, st.br, x2, y1, x2, y2, t);
  if (p.table?.borderTop && !st.bt) strokeBorder(ctx, { style: 'thin', color: p.table.borderTop }, x1, y1, x2, y1, t);
  if (st.diagDown) { ctx.strokeStyle = st.diagDown.color || '#000'; ctx.lineWidth = 1; ctx.beginPath(); ctx.moveTo(p.x, p.y); ctx.lineTo(p.x + p.w, p.y + p.h); ctx.stroke(); }
  if (st.diagUp) { ctx.strokeStyle = st.diagUp.color || '#000'; ctx.lineWidth = 1; ctx.beginPath(); ctx.moveTo(p.x, p.y + p.h); ctx.lineTo(p.x + p.w, p.y); ctx.stroke(); }
  void g;
}

function paintSparkline(g: Grid, ctx: CanvasRenderingContext2D, p: CellPaint): void {
  const sp = g.sheet.sparklines.find((x) => x.cell.r === p.r && x.cell.c === p.c);
  if (!sp) return;
  const m = g.tab.engine.evaluateMatrix(sp.data, g.sheet, p.r, p.c);
  const vals = m.flat().map((v: Scalar) => (typeof v === 'number' ? v : null));
  const nums = vals.filter((v): v is number => v != null);
  if (!nums.length) return;
  const min = Math.min(...nums, sp.type === 'column' ? 0 : Infinity), max = Math.max(...nums, sp.type === 'column' ? 0 : -Infinity);
  const pad = 3;
  const W = p.w - pad * 2, H = p.h - pad * 2;
  const y = (v: number) => p.y + pad + H - ((v - min) / (max - min || 1)) * H;
  ctx.save();
  ctx.beginPath(); ctx.rect(p.x, p.y, p.w, p.h); ctx.clip();
  const color = sp.color || '#2f6fb0';
  if (sp.type === 'line') {
    ctx.strokeStyle = color; ctx.lineWidth = 1.2;
    ctx.beginPath();
    let first = true;
    vals.forEach((v, i) => { if (v == null) return; const x = p.x + pad + (W * i) / Math.max(1, vals.length - 1); if (first) ctx.moveTo(x, y(v)); else ctx.lineTo(x, y(v)); first = false; });
    ctx.stroke();
    if (sp.markers || sp.high || sp.low) vals.forEach((v, i) => {
      if (v == null) return;
      const isHi = v === Math.max(...nums), isLo = v === Math.min(...nums);
      if (!(sp.markers || (sp.high && isHi) || (sp.low && isLo))) return;
      ctx.fillStyle = sp.high && isHi ? '#2e7d32' : sp.low && isLo ? '#c00000' : color;
      ctx.beginPath(); ctx.arc(p.x + pad + (W * i) / Math.max(1, vals.length - 1), y(v), 2, 0, Math.PI * 2); ctx.fill();
    });
  } else {
    const bw = W / vals.length;
    const zeroY = sp.type === 'winloss' ? p.y + p.h / 2 : y(Math.max(min, 0));
    vals.forEach((v, i) => {
      if (v == null) return;
      ctx.fillStyle = v < 0 ? '#c00000' : color;
      const x = p.x + pad + i * bw + bw * 0.15;
      if (sp.type === 'winloss') { const hh = H / 2 - 1; ctx.fillRect(x, v >= 0 ? zeroY - hh : zeroY + 1, bw * 0.7, hh); }
      else { const yy = y(v); ctx.fillRect(x, Math.min(yy, zeroY), bw * 0.7, Math.abs(zeroY - yy) || 1); }
    });
  }
  ctx.restore();
}

function paintPageBreaks(g: Grid, ctx: CanvasRenderingContext2D, t: Theme, r1: number, r2: number, c1: number, c2: number): void {
  const breaks = (g as any).pageBreaks as { rows: number[]; cols: number[]; area: Range } | undefined;
  if (!breaks) return;
  ctx.strokeStyle = t.pageBreak; ctx.lineWidth = g.tab.viewMode === 'pageBreak' ? 2 : 1;
  ctx.setLineDash([6, 4]);
  ctx.beginPath();
  const a = breaks.area;
  for (const r of breaks.rows) if (r >= r1 && r <= r2 + 1) { const y = g.rowY(r) - 0.5; ctx.moveTo(g.colX(Math.max(a.c1, c1)), y); ctx.lineTo(g.colX(Math.min(a.c2, c2)) + g.colW(Math.min(a.c2, c2)), y); }
  for (const c of breaks.cols) if (c >= c1 && c <= c2 + 1) { const x = g.colX(c) - 0.5; ctx.moveTo(x, g.rowY(Math.max(a.r1, r1))); ctx.lineTo(x, g.rowY(Math.min(a.r2, r2)) + g.rowH(Math.min(a.r2, r2))); }
  ctx.stroke();
  ctx.setLineDash([]);
  if (g.tab.viewMode === 'pageBreak') {
    // Gray out cells outside the print area and label pages.
    ctx.fillStyle = t.dark ? 'rgba(0,0,0,0.45)' : 'rgba(128,128,128,0.35)';
    const ax1 = g.colX(a.c1), ay1 = g.rowY(a.r1), ax2 = g.colX(a.c2) + g.colW(a.c2), ay2 = g.rowY(a.r2) + g.rowH(a.r2);
    ctx.fillRect(0, 0, g.width, ay1); ctx.fillRect(0, ay2, g.width, g.height); ctx.fillRect(0, ay1, ax1, ay2 - ay1); ctx.fillRect(ax2, ay1, g.width, ay2 - ay1);
    ctx.strokeStyle = t.pageBreak; ctx.lineWidth = 3; ctx.strokeRect(ax1, ay1, ax2 - ax1, ay2 - ay1);
    ctx.fillStyle = 'rgba(128,128,128,0.35)';
    ctx.font = `bold ${Math.round(36 * g.zoom)}px system-ui, sans-serif`;
    ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
    const rs = [a.r1, ...breaks.rows.filter((r) => r > a.r1 && r <= a.r2), a.r2 + 1];
    const cs = [a.c1, ...breaks.cols.filter((c) => c > a.c1 && c <= a.c2), a.c2 + 1];
    let n = 1;
    for (let j = 0; j < cs.length - 1; j++) for (let i = 0; i < rs.length - 1; i++) {
      const x = (g.colX(cs[j]) + g.colX(cs[j + 1])) / 2, y = (g.rowY(rs[i]) + g.rowY(rs[i + 1])) / 2;
      ctx.fillText(`Page ${n++}`, x, y);
    }
    ctx.textAlign = 'left';
  }
}

// ----------------------------------------------------------------------------------------------
// Headers

function paintHeaders(g: Grid, ctx: CanvasRenderingContext2D, t: Theme, vis: { r1: number; r2: number; c1: number; c2: number }): void {
  const s = g.sheet;
  const hx = g.headerW, hy = g.headerH, z = g.zoom;
  const fr = s.freeze.rows, fc = s.freeze.cols;
  const font = `${Math.round(11 * Math.max(0.8, z) * 4 / 3)}px system-ui, -apple-system, "Segoe UI", sans-serif`;
  ctx.font = font;
  ctx.textBaseline = 'middle';
  ctx.textAlign = 'center';
  // Selected rows/cols for highlighting
  const selCols = new Set<number>(), selRows = new Set<number>();
  const fullCols = new Set<number>(), fullRows = new Set<number>();
  for (const rg of s.sel.ranges) {
    const big = (rg.c2 - rg.c1) * (rg.r2 - rg.r1) > 4e6;
    for (let c = Math.max(rg.c1, 0); c <= Math.min(rg.c2, vis.c2 + 1); c++) { selCols.add(c); if (rg.r1 === 0 && rg.r2 >= MAX_ROWS - 1) fullCols.add(c); }
    for (let r = Math.max(rg.r1, 0); r <= Math.min(rg.r2, vis.r2 + 1); r++) { selRows.add(r); if (rg.c1 === 0 && rg.c2 >= MAX_COLS - 1) fullRows.add(r); }
    void big;
  }
  // Column headers
  const colRanges: [number, number, number, number][] = [];
  if (fc) colRanges.push([0, fc - 1, hx, g.frozenW]);
  colRanges.push([vis.c1, Math.min(MAX_COLS - 1, vis.c2 + 1), hx + g.frozenW, g.width - hx - g.frozenW]);
  for (const [a, b, clipX, clipW] of colRanges) {
    ctx.save(); ctx.beginPath(); ctx.rect(clipX, 0, clipW, hy); ctx.clip();
    ctx.fillStyle = t.headerBg; ctx.fillRect(clipX, 0, clipW, hy);
    for (let c = a; c <= b; c++) {
      const w = g.colW(c); if (!w) continue;
      const x = g.colX(c);
      const sel = selCols.has(c);
      if (sel) { ctx.fillStyle = fullCols.has(c) ? t.accent : t.headerSel; ctx.fillRect(x, 0, w, hy); }
      ctx.fillStyle = fullCols.has(c) ? '#ffffff' : sel ? t.headerSelText : t.headerText;
      ctx.fillText(g.app.prefs.r1c1 ? String(c + 1) : colName(c), x + w / 2, hy / 2 + 1);
      ctx.fillStyle = t.headerLine; ctx.fillRect(Math.round(x + w) - 1, 0, 1, hy);
      if (sel && !fullCols.has(c)) { ctx.fillStyle = t.accent; ctx.fillRect(x, hy - 2, w, 2); }
    }
    ctx.restore();
  }
  // Row headers
  const rowRanges: [number, number, number, number][] = [];
  if (fr) rowRanges.push([0, fr - 1, hy, g.frozenH]);
  rowRanges.push([vis.r1, Math.min(MAX_ROWS - 1, vis.r2 + 1), hy + g.frozenH, g.height - hy - g.frozenH]);
  ctx.textAlign = 'right';
  for (const [a, b, clipY, clipH] of rowRanges) {
    ctx.save(); ctx.beginPath(); ctx.rect(0, clipY, hx, clipH); ctx.clip();
    ctx.fillStyle = t.headerBg; ctx.fillRect(0, clipY, hx, clipH);
    for (let r = a; r <= b; r++) {
      const hh = g.rowH(r); if (!hh) continue;
      const y = g.rowY(r);
      const sel = selRows.has(r);
      if (sel) { ctx.fillStyle = fullRows.has(r) ? t.accent : t.headerSel; ctx.fillRect(0, y, hx, hh); }
      ctx.fillStyle = fullRows.has(r) ? '#ffffff' : sel ? t.headerSelText : t.headerText;
      if (hh > 8) ctx.fillText(String(r + 1), hx - 5, y + hh / 2 + 1);
      ctx.fillStyle = t.headerLine; ctx.fillRect(0, Math.round(y + hh) - 1, hx, 1);
      if (sel && !fullRows.has(r)) { ctx.fillStyle = t.accent; ctx.fillRect(hx - 2, y, 2, hh); }
    }
    ctx.restore();
  }
  ctx.textAlign = 'left';
  // Outline (group) bars are drawn by the outline layer in the shell; header borders:
  ctx.fillStyle = t.headerLine;
  ctx.fillRect(0, hy - 1, g.width, 1);
  ctx.fillRect(hx - 1, 0, 1, g.height);
  // Corner (select all)
  ctx.fillStyle = t.headerBg; ctx.fillRect(0, 0, hx - 1, hy - 1);
  ctx.fillStyle = t.headerText;
  ctx.beginPath(); ctx.moveTo(hx - 4, hy - 12 * Math.min(1, z)); ctx.lineTo(hx - 4, hy - 4); ctx.lineTo(hx - 12 * Math.min(1, z), hy - 4); ctx.closePath(); ctx.fill();
  // Column resize tooltip
  if (g.resizePreview) {
    const p = g.resizePreview;
    const label = p.kind === 'col' ? `Width: ${(Math.max(0, p.size - 5) / 7).toFixed(2)} (${p.size} pixels)` : `Height: ${(p.size * 0.75).toFixed(2)} (${p.size} pixels)`;
    ctx.font = '12px system-ui, sans-serif';
    const w = ctx.measureText(label).width + 12;
    const x = p.kind === 'col' ? g.colX(p.index) + p.size * z : hx + 4, y = p.kind === 'col' ? hy + 4 : g.rowY(p.index) + p.size * z;
    ctx.fillStyle = '#ffffe1'; ctx.strokeStyle = '#767676';
    ctx.fillRect(x + 4, y + 4, w, 20); ctx.strokeRect(x + 4.5, y + 4.5, w - 1, 19);
    ctx.fillStyle = '#000'; ctx.textBaseline = 'middle'; ctx.fillText(label, x + 10, y + 14);
    ctx.strokeStyle = t.accent; ctx.setLineDash([3, 3]); ctx.beginPath();
    if (p.kind === 'col') { const lx = Math.round(g.colX(p.index) + p.size * z) + 0.5; ctx.moveTo(lx, 0); ctx.lineTo(lx, g.height); }
    else { const ly = Math.round(g.rowY(p.index) + p.size * z) + 0.5; ctx.moveTo(0, ly); ctx.lineTo(g.width, ly); }
    ctx.stroke(); ctx.setLineDash([]);
  }
}

// ----------------------------------------------------------------------------------------------
// Selection and previews

function paintSelection(g: Grid, ctx: CanvasRenderingContext2D, t: Theme): void {
  const s = g.sheet;
  const sel = s.sel;
  const multi = sel.ranges.length > 1;
  const ac = g.cellRect(sel.r, sel.c);
  for (let i = 0; i < sel.ranges.length; i++) {
    const rg = sel.ranges[i];
    const rc = g.rangeRect(rg);
    // The active cell is left unshaded: fill the range minus the active cell (even-odd rule).
    ctx.fillStyle = t.selFill;
    ctx.beginPath();
    ctx.rect(rc.x, rc.y, rc.w, rc.h);
    if (sel.r >= rg.r1 && sel.r <= rg.r2 && sel.c >= rg.c1 && sel.c <= rg.c2) ctx.rect(ac.x, ac.y, ac.w, ac.h);
    ctx.fill('evenodd');
    if (multi) { ctx.strokeStyle = t.accent; ctx.lineWidth = 1; ctx.strokeRect(rc.x + 0.5, rc.y + 0.5, rc.w - 1, rc.h - 1); }
  }
  const act = sel.ranges[sel.ranges.length - 1];
  const rc = g.rangeRect(act);
  ctx.strokeStyle = t.accent;
  ctx.lineWidth = 2;
  if (!multi) ctx.strokeRect(Math.round(rc.x) - 0.5, Math.round(rc.y) - 0.5, Math.round(rc.w) + 1, Math.round(rc.h) + 1);
  ctx.lineWidth = multi ? 2 : 1;
  if (multi || act.r1 !== act.r2 || act.c1 !== act.c2) { ctx.strokeStyle = t.accent; ctx.strokeRect(Math.round(ac.x) + 0.5, Math.round(ac.y) + 0.5, Math.round(ac.w) - 1, Math.round(ac.h) - 1); }
  // Fill handle / touch handles
  if (!multi && !g.editing) {
    if (g.fillArmed) {
      ctx.fillStyle = t.bg; ctx.fillRect(rc.x + rc.w - 7, rc.y + rc.h - 7, 14, 14);
      ctx.fillStyle = t.accent; ctx.fillRect(rc.x + rc.w - 6, rc.y + rc.h - 6, 12, 12);
    } else if (g.touchMode && g.app.prefs.touchHandles) {
      ctx.fillStyle = t.bg; ctx.strokeStyle = t.accent; ctx.lineWidth = 2;
      for (const [x, y] of [[rc.x, rc.y], [rc.x + rc.w, rc.y + rc.h]]) { ctx.beginPath(); ctx.arc(x, y, HANDLE_R, 0, Math.PI * 2); ctx.fill(); ctx.stroke(); }
    } else if (g.app.prefs.fillHandle) {
      ctx.fillStyle = t.bg; ctx.fillRect(rc.x + rc.w - 4, rc.y + rc.h - 4, 8, 8);
      ctx.fillStyle = t.accent; ctx.fillRect(rc.x + rc.w - 3, rc.y + rc.h - 3, 6, 6);
    }
  }
  // Spill range outline for the active cell
  const sp = g.tab.engine.spillRange(s, sel.r, sel.c) || (() => { const x = s.get(sel.r, sel.c); if (x?.sp != null) { const ar = Math.floor(x.sp / MAX_COLS), acc = x.sp % MAX_COLS; return g.tab.engine.spillRange(s, ar, acc); } return null; })();
  if (sp) {
    const r2 = g.rangeRect(sp);
    ctx.strokeStyle = '#4472c4'; ctx.lineWidth = 1; ctx.setLineDash([4, 2]);
    ctx.strokeRect(r2.x + 0.5, r2.y + 0.5, r2.w - 1, r2.h - 1); ctx.setLineDash([]);
  }
  // Copy/cut marquee
  const mq = g.tab.marquee;
  if (mq && mq.sheet === s) {
    const r3 = g.rangeRect(mq.range);
    ctx.strokeStyle = t.accent; ctx.lineWidth = 2; ctx.setLineDash([5, 4]);
    ctx.lineDashOffset = -((Date.now() / 60) % 9);
    ctx.strokeRect(r3.x + 1, r3.y + 1, r3.w - 2, r3.h - 2);
    ctx.setLineDash([]); ctx.lineDashOffset = 0;
  }
  // Formula reference highlights
  for (const hl of g.refHighlights) {
    if (hl.sheet && hl.sheet.toLowerCase() !== s.name.toLowerCase()) continue;
    const r4 = g.rangeRect(clampRange(hl.range, g));
    ctx.strokeStyle = hl.color; ctx.lineWidth = 2;
    ctx.fillStyle = hl.color + '18';
    ctx.fillRect(r4.x, r4.y, r4.w, r4.h);
    ctx.strokeRect(r4.x + 1, r4.y + 1, r4.w - 2, r4.h - 2);
    ctx.fillStyle = hl.color;
    for (const [x, y] of [[r4.x, r4.y], [r4.x + r4.w, r4.y], [r4.x, r4.y + r4.h], [r4.x + r4.w, r4.y + r4.h]]) ctx.fillRect(x - 2.5, y - 2.5, 5, 5);
  }
  if (g.pointRange) {
    const r5 = g.rangeRect(g.pointRange);
    ctx.strokeStyle = '#2f6fb0'; ctx.lineWidth = 2; ctx.setLineDash([4, 3]);
    ctx.lineDashOffset = -((Date.now() / 60) % 7);
    ctx.strokeRect(r5.x + 1, r5.y + 1, r5.w - 2, r5.h - 2); ctx.setLineDash([]);
    setTimeout(() => g.invalidate(), 150);
  }
  // Trace arrows
  for (const a of g.tab.arrows) {
    if (a.from.sheet !== s || a.to.sheet !== s) continue;
    const f = g.cellRect(a.from.r, a.from.c), tt = g.cellRect(a.to.r, a.to.c);
    const x1 = f.x + f.w / 2, y1 = f.y + f.h / 2, x2 = tt.x + tt.w / 2, y2 = tt.y + tt.h / 2;
    ctx.strokeStyle = '#2f6fb0'; ctx.fillStyle = '#2f6fb0'; ctx.lineWidth = 1.5;
    ctx.beginPath(); ctx.moveTo(x1, y1); ctx.lineTo(x2, y2); ctx.stroke();
    ctx.beginPath(); ctx.arc(x1, y1, 3, 0, Math.PI * 2); ctx.fill();
    const ang = Math.atan2(y2 - y1, x2 - x1);
    ctx.beginPath(); ctx.moveTo(x2, y2); ctx.lineTo(x2 - 9 * Math.cos(ang - 0.4), y2 - 9 * Math.sin(ang - 0.4)); ctx.lineTo(x2 - 9 * Math.cos(ang + 0.4), y2 - 9 * Math.sin(ang + 0.4)); ctx.closePath(); ctx.fill();
  }
}

function clampRange(rg: Range, g: Grid): Range {
  const vis = g.visible();
  return { r1: rg.r1, c1: rg.c1, r2: Math.min(rg.r2, Math.max(vis.r2 + 2, rg.r1)), c2: Math.min(rg.c2, Math.max(vis.c2 + 2, rg.c1)) };
}

function paintPreviews(g: Grid, ctx: CanvasRenderingContext2D, t: Theme): void {
  const p = g.preview;
  if (!p) return;
  const rc = g.rangeRect(p.range);
  ctx.strokeStyle = p.kind === 'fill' ? '#7f7f7f' : t.accent;
  ctx.lineWidth = p.kind === 'fill' ? 1 : 2;
  ctx.setLineDash(p.kind === 'fill' ? [3, 3] : []);
  ctx.strokeRect(rc.x + 0.5, rc.y + 0.5, rc.w - 1, rc.h - 1);
  ctx.setLineDash([]);
  if (p.kind === 'move') {
    ctx.font = '12px system-ui, sans-serif';
    const label = rangeLabel(p.range);
    ctx.fillStyle = '#ffffe1'; ctx.fillRect(rc.x + rc.w + 4, rc.y + rc.h + 4, ctx.measureText(label).width + 10, 18);
    ctx.fillStyle = '#000'; ctx.textBaseline = 'middle'; ctx.fillText(label, rc.x + rc.w + 9, rc.y + rc.h + 13);
  }
}

function rangeLabel(rg: Range): string {
  const a = `${colName(rg.c1)}${rg.r1 + 1}`, b = `${colName(rg.c2)}${rg.r2 + 1}`;
  return a === b ? a : `${a}:${b}`;
}
