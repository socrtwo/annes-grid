/**
 * Draw tab: freehand ink over the grid with pens, pencil and highlighter (pressure-aware on pens
 * and styluses), an eraser, lasso selection, ink-to-shape recognition and stroke replay.
 *
 * While a drawing tool is active a transparent canvas covers the grid (inside grid.overlay). Strokes
 * drawn in quick succession are saved together as one DrawingObject of kind 'ink', anchored to the
 * cell under its top-left corner. The stroke geometry is kept in `svg` as path elements; it is
 * parsed back (parseInk) and re-rendered from numbers only, so no stored markup reaches the page.
 */
import { def, type MenuEntry } from './registry';
import type { App } from '../app';
import type { DrawingObject } from '../model/types';
import type { Sheet } from '../model/workbook';
import { newId } from '../model/workbook';
import { objectHooks } from '../grid/objects';
import { toast, promptDialog } from '../ui/dialog';
import { h } from '../ui/dom';
import { safeColor, safeNum } from '../util/sanitize';

// ============================================================================================
// Ink data

export type InkTool = 'pen' | 'pencil' | 'highlighter';
export interface InkStroke { tool: InkTool; color: string; width: number; opacity: number; points: [number, number][]; widths?: number[] }

const num = (n: number) => String(Math.round(n * 10) / 10);

/** Serializes strokes (points relative to the object's top-left, unzoomed) as SVG markup. Pen strokes with varying width are split into runs. */
export function inkSvg(strokes: InkStroke[], w: number, h: number): string {
  const parts: string[] = [];
  strokes.forEach((s, i) => {
    const color = safeColor(s.color, '#000000');
    const attrs = (width: number) => `stroke="${color}" stroke-width="${num(width)}" stroke-opacity="${num(s.opacity)}" fill="none" stroke-linecap="${s.tool === 'highlighter' ? 'square' : 'round'}" stroke-linejoin="round" data-s="${i}" data-tool="${s.tool}"`;
    const pts = s.points;
    if (!pts.length) return;
    if (!s.widths || s.tool !== 'pen') {
      parts.push(`<path d="${pathD(pts)}" ${attrs(s.width)}/>`);
      return;
    }
    // Runs of similar width (0.5px buckets) keep the markup small while following pressure.
    let start = 0;
    const bucket = (k: number) => Math.round((s.widths![k] ?? s.width) * 2) / 2;
    for (let k = 1; k <= pts.length; k++) {
      if (k === pts.length || bucket(k) !== bucket(start)) {
        const run = pts.slice(Math.max(0, start - 1), k);
        parts.push(`<path d="${pathD(run.length > 1 ? run : [run[0], run[0]])}" ${attrs(bucket(start))}/>`);
        start = k;
      }
    }
  });
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${num(w)} ${num(h)}" preserveAspectRatio="none" data-ink="1">${parts.join('')}</svg>`;
}

function pathD(pts: [number, number][]): string {
  return pts.map((p, i) => `${i ? 'L' : 'M'}${num(p[0])} ${num(p[1])}`).join(' ');
}

/** Parses ink markup written by inkSvg (or any simple path list) back into strokes. Unknown content is ignored. */
export function parseInk(svg: string | undefined): { strokes: InkStroke[]; w: number; h: number } {
  const out: InkStroke[] = [];
  if (!svg) return { strokes: out, w: 0, h: 0 };
  const vb = /viewBox="\s*[-\d.]+\s+[-\d.]+\s+([\d.]+)\s+([\d.]+)\s*"/.exec(svg);
  const byIndex = new Map<string, InkStroke>();
  for (const m of svg.matchAll(/<path\b([^>]*)\/?>/g)) {
    const a = m[1];
    const attr = (n: string) => new RegExp(`\\b${n}="([^"]*)"`).exec(a)?.[1];
    const d = attr('d') || '';
    if (!/^[MLml0-9.,\s-]*$/.test(d)) continue;
    const pts: [number, number][] = [];
    for (const p of d.matchAll(/[ML]\s*(-?[\d.]+)[\s,]+(-?[\d.]+)/gi)) pts.push([+p[1], +p[2]]);
    if (!pts.length) continue;
    const key = attr('data-s') ?? `p${byIndex.size}`;
    const tool = (['pen', 'pencil', 'highlighter'].includes(attr('data-tool') || '') ? attr('data-tool') : 'pen') as InkTool;
    const width = Math.max(0.1, Math.min(200, Number(attr('stroke-width')) || 2));
    const prev = byIndex.get(key);
    if (prev) {
      const skip = prev.points.length && pts.length && prev.points[prev.points.length - 1][0] === pts[0][0] && prev.points[prev.points.length - 1][1] === pts[0][1] ? 1 : 0;
      for (const p of pts.slice(skip)) { prev.points.push(p); prev.widths!.push(width); }
      prev.width = Math.max(prev.width, width);
      continue;
    }
    const s: InkStroke = { tool, color: safeColor(attr('stroke'), '#000000'), width, opacity: Math.max(0.05, Math.min(1, Number(attr('stroke-opacity') ?? 1) || 1)), points: pts, widths: pts.map(() => width) };
    byIndex.set(key, s);
    out.push(s);
  }
  for (const s of out) if (s.widths && s.widths.every((x) => x === s.widths![0])) delete s.widths;
  return { strokes: out, w: vb ? +vb[1] : 0, h: vb ? +vb[2] : 0 };
}

export function strokeBounds(strokes: InkStroke[]): { x1: number; y1: number; x2: number; y2: number } {
  let x1 = Infinity, y1 = Infinity, x2 = -Infinity, y2 = -Infinity;
  for (const s of strokes) for (const [x, y] of s.points) { const pad = s.width / 2; x1 = Math.min(x1, x - pad); y1 = Math.min(y1, y - pad); x2 = Math.max(x2, x + pad); y2 = Math.max(y2, y + pad); }
  return { x1, y1, x2, y2 };
}

// --------------------------------------------------------------------------------------------
// Shape recognition

export type Recognized = { shape: 'line'; x1: number; y1: number; x2: number; y2: number } | { shape: 'rect' | 'ellipse' | 'triangle'; x: number; y: number; w: number; h: number; points?: [number, number][] };

function dist(a: [number, number], b: [number, number]): number { return Math.hypot(a[0] - b[0], a[1] - b[1]); }

function segDist(p: [number, number], a: [number, number], b: [number, number]): number {
  const dx = b[0] - a[0], dy = b[1] - a[1];
  const L = dx * dx + dy * dy;
  if (!L) return dist(p, a);
  const t = Math.max(0, Math.min(1, ((p[0] - a[0]) * dx + (p[1] - a[1]) * dy) / L));
  return Math.hypot(p[0] - a[0] - t * dx, p[1] - a[1] - t * dy);
}

/** Ramer-Douglas-Peucker simplification. */
export function simplify(pts: [number, number][], eps: number): [number, number][] {
  if (pts.length < 3) return pts.slice();
  let idx = 0, max = 0;
  for (let i = 1; i < pts.length - 1; i++) { const d = segDist(pts[i], pts[0], pts[pts.length - 1]); if (d > max) { max = d; idx = i; } }
  if (max <= eps) return [pts[0], pts[pts.length - 1]];
  const a = simplify(pts.slice(0, idx + 1), eps), b = simplify(pts.slice(idx), eps);
  return [...a.slice(0, -1), ...b];
}

/** Recognizes a line, rectangle, ellipse or triangle from one stroke, or returns null. */
export function recognizeShape(points: [number, number][]): Recognized | null {
  if (points.length < 2) return null;
  let x1 = Infinity, y1 = Infinity, x2 = -Infinity, y2 = -Infinity;
  for (const [x, y] of points) { x1 = Math.min(x1, x); y1 = Math.min(y1, y); x2 = Math.max(x2, x); y2 = Math.max(y2, y); }
  const w = x2 - x1, hh = y2 - y1;
  const diag = Math.hypot(w, hh);
  if (diag < 8) return null;
  let len = 0;
  for (let i = 1; i < points.length; i++) len += dist(points[i - 1], points[i]);
  const first = points[0], last = points[points.length - 1];
  const ends = dist(first, last);
  // Straight: the path is barely longer than the distance between its ends.
  if (ends > 0.6 * diag && len < ends * 1.12) return { shape: 'line', x1: first[0], y1: first[1], x2: last[0], y2: last[1] };
  const closed = ends < Math.max(0.25 * diag, 12);
  if (!closed) return null;
  // Closed figure: count corners of the simplified outline.
  const loop: [number, number][] = [...points, first];
  let simp = simplify(loop, Math.max(4, 0.08 * diag));
  if (simp.length > 2 && dist(simp[0], simp[simp.length - 1]) < 0.2 * diag) simp = simp.slice(0, -1);
  // Merge near-collinear corners
  const corners = simp.filter((p, i) => {
    const a = simp[(i - 1 + simp.length) % simp.length], b = simp[(i + 1) % simp.length];
    const v1 = [p[0] - a[0], p[1] - a[1]], v2 = [b[0] - p[0], b[1] - p[1]];
    const cos = (v1[0] * v2[0] + v1[1] * v2[1]) / (Math.hypot(v1[0], v1[1]) * Math.hypot(v2[0], v2[1]) || 1);
    return cos < 0.85;
  });
  // Ellipse test: distance of points to the ellipse inscribed in the bounding box
  const cx = (x1 + x2) / 2, cy = (y1 + y2) / 2, rx = w / 2 || 1, ry = hh / 2 || 1;
  let err = 0;
  for (const [x, y] of points) { const d = Math.hypot((x - cx) / rx, (y - cy) / ry); err += Math.abs(d - 1); }
  err /= points.length;
  if (corners.length === 3) return { shape: 'triangle', x: x1, y: y1, w, h: hh, points: corners };
  if (corners.length === 4 && err > 0.07) return { shape: 'rect', x: x1, y: y1, w, h: hh };
  if (err < 0.16) return { shape: 'ellipse', x: x1, y: y1, w, h: hh };
  if (corners.length === 4 || corners.length === 5) return { shape: 'rect', x: x1, y: y1, w, h: hh };
  return null;
}

// ============================================================================================
// Rendering (objects layer hook)

function renderInk(d: DrawingObject): HTMLElement | null {
  if (d.kind !== 'ink') return null;
  const { strokes, w, h: hh } = parseInk(d.svg);
  const wrap = h('div', { class: 'ql-drawing-svg ql-ink' });
  wrap.innerHTML = inkSvg(strokes, w || d.at.w, hh || d.at.h).replace('<svg ', `<svg width="${safeNum(d.at.w, 120, 1, 20000)}" height="${safeNum(d.at.h, 120, 1, 20000)}" `);
  return wrap;
}
objectHooks.render.push((d) => renderInk(d));

// ============================================================================================
// Ink mode

interface PenSettings { color: string; width: number }
const DEFAULTS: Record<InkTool, PenSettings> = { pen: { color: '#1f3b57', width: 2.5 }, pencil: { color: '#444444', width: 1.5 }, highlighter: { color: '#ffeb3b', width: 14 } };
const PREF_KEY = 'quillgrid.ink';

function loadPens(): Record<InkTool, PenSettings> {
  try { const j = JSON.parse(localStorage.getItem(PREF_KEY) || '{}'); return { pen: { ...DEFAULTS.pen, ...j.pen }, pencil: { ...DEFAULTS.pencil, ...j.pencil }, highlighter: { ...DEFAULTS.highlighter, ...j.highlighter } }; } catch { return { ...DEFAULTS }; }
}
const pens = loadPens();
const savePens = () => { try { localStorage.setItem(PREF_KEY, JSON.stringify(pens)); } catch { /* unavailable */ } };

type Mode = InkTool | 'eraser' | 'lasso';
let lastPen: InkTool = 'pen';
let autoShapes = false;

interface LiveStroke extends InkStroke { abs: [number, number][] }

class InkSession {
  canvas: HTMLCanvasElement;
  ctx: CanvasRenderingContext2D;
  mode: Mode;
  pending: LiveStroke[] = [];
  cur: LiveStroke | null = null;
  lasso: [number, number][] | null = null;
  erased = new Map<string, Set<number>>();
  timer: any = null;
  private onKey = (e: KeyboardEvent) => { if (e.key === 'Escape' && !document.querySelector('.ql-overlay')) { e.preventDefault(); this.app.exec('draw.select'); } };
  private onResize = () => this.size();

  constructor(public app: App, mode: Mode) {
    this.mode = mode;
    const g = app.grid;
    this.canvas = h('canvas', { class: 'ql-ink-canvas', 'aria-label': 'Drawing surface (Esc to stop drawing)' }) as HTMLCanvasElement;
    this.ctx = this.canvas.getContext('2d')!;
    g.overlay.appendChild(this.canvas);
    this.size();
    this.canvas.addEventListener('pointerdown', (e) => this.down(e));
    this.canvas.addEventListener('pointermove', (e) => this.move(e));
    this.canvas.addEventListener('pointerup', (e) => this.up(e));
    this.canvas.addEventListener('pointercancel', (e) => this.up(e));
    this.canvas.addEventListener('wheel', (e) => { this.commit(); g.scroller.scrollTop += e.deltaY; g.scroller.scrollLeft += e.deltaX; e.preventDefault(); }, { passive: false });
    document.addEventListener('keydown', this.onKey, true);
    window.addEventListener('resize', this.onResize);
    this.cursor();
  }

  get grid() { return this.app.grid; }
  get sheet(): Sheet { return this.app.sheet; }

  size(): void {
    const r = this.grid.overlay.getBoundingClientRect();
    const dpr = window.devicePixelRatio || 1;
    this.canvas.width = Math.max(1, Math.round(r.width * dpr));
    this.canvas.height = Math.max(1, Math.round(r.height * dpr));
    this.canvas.style.width = `${r.width}px`;
    this.canvas.style.height = `${r.height}px`;
    this.ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    this.redraw();
  }

  cursor(): void {
    this.canvas.style.cursor = this.mode === 'eraser' ? 'cell' : this.mode === 'lasso' ? 'crosshair' : 'crosshair';
    this.canvas.dataset.mode = this.mode;
  }

  setMode(m: Mode): void { this.commit(); this.mode = m; this.cursor(); }

  pos(e: PointerEvent): [number, number] {
    const r = this.canvas.getBoundingClientRect();
    return [e.clientX - r.left, e.clientY - r.top];
  }

  down(e: PointerEvent): void {
    if (e.button !== 0 && e.pointerType === 'mouse') return;
    e.preventDefault();
    this.canvas.setPointerCapture(e.pointerId);
    clearTimeout(this.timer);
    const p = this.pos(e);
    if (this.mode === 'lasso') { this.lasso = [p]; return; }
    if (this.mode === 'eraser') { this.eraseAt(p); return; }
    const tool = this.mode;
    const pen = pens[tool];
    const z = this.grid.zoom;
    this.cur = { tool, color: pen.color, width: pen.width * z, opacity: tool === 'highlighter' ? 0.4 : tool === 'pencil' ? 0.85 : 1, points: [], abs: [p], widths: [this.widthFor(e, tool)] };
  }

  widthFor(e: PointerEvent, tool: InkTool): number {
    const base = pens[tool].width * this.grid.zoom;
    if (tool !== 'pen') return base;
    // Mouse reports 0.5 while pressed; pens and touch report real pressure.
    const p = e.pointerType === 'mouse' || !e.pressure ? 0.5 : e.pressure;
    return Math.max(0.5, base * (0.35 + p * 1.3));
  }

  move(e: PointerEvent): void {
    if (this.mode === 'eraser' && e.buttons) { this.eraseAt(this.pos(e)); return; }
    if (this.lasso) { this.lasso.push(this.pos(e)); this.redraw(); return; }
    if (!this.cur) return;
    const events = (e as any).getCoalescedEvents?.() as PointerEvent[] | undefined;
    for (const ev of events?.length ? events : [e]) {
      const p = this.pos(ev);
      const last = this.cur.abs[this.cur.abs.length - 1];
      if (Math.hypot(p[0] - last[0], p[1] - last[1]) < 1.2) continue;
      this.cur.abs.push(p);
      this.cur.widths!.push(this.widthFor(ev, this.cur.tool));
      this.drawSegment(this.cur, this.cur.abs.length - 1);
    }
  }

  up(e: PointerEvent): void {
    try { this.canvas.releasePointerCapture(e.pointerId); } catch { /* not captured */ }
    if (this.mode === 'eraser') { this.applyErase(); return; }
    if (this.lasso) { const poly = this.lasso; this.lasso = null; this.redraw(); this.lassoSelect(poly); return; }
    const s = this.cur;
    this.cur = null;
    if (!s) return;
    if (s.abs.length === 1) { s.abs.push([s.abs[0][0] + 0.5, s.abs[0][1] + 0.5]); s.widths!.push(s.widths![0]); }
    if (autoShapes && s.tool !== 'highlighter') {
      const rec = recognizeShape(s.abs);
      if (rec) { this.commit(); addRecognizedShape(this.app, rec, s.color, pens[s.tool].width, (x, y) => this.grid.objects.anchorFor(x, y), this.grid.zoom); this.redraw(); return; }
    }
    this.pending.push(s);
    this.timer = setTimeout(() => this.commit(), 1500);
  }

  drawStroke(s: LiveStroke, upto = s.abs.length): void {
    for (let i = 1; i < upto; i++) this.drawSegment(s, i);
    if (upto === 1) this.drawSegment(s, 0);
  }

  drawSegment(s: LiveStroke, i: number): void {
    const c = this.ctx;
    const a = s.abs[Math.max(0, i - 1)], b = s.abs[i];
    c.save();
    c.globalAlpha = s.opacity;
    c.strokeStyle = s.color;
    c.lineWidth = s.widths?.[i] ?? s.width;
    c.lineCap = s.tool === 'highlighter' ? 'square' : 'round';
    c.lineJoin = 'round';
    if (s.tool === 'highlighter') c.globalCompositeOperation = 'multiply';
    c.beginPath(); c.moveTo(a[0], a[1]); c.lineTo(b[0], b[1]); c.stroke();
    c.restore();
  }

  redraw(): void {
    const r = this.canvas.getBoundingClientRect();
    this.ctx.clearRect(0, 0, r.width || this.canvas.width, r.height || this.canvas.height);
    for (const s of this.pending) this.drawStroke(s);
    if (this.cur) this.drawStroke(this.cur);
    if (this.lasso && this.lasso.length > 1) {
      const c = this.ctx;
      c.save(); c.setLineDash([5, 4]); c.strokeStyle = '#2f6fb0'; c.lineWidth = 1.2; c.beginPath();
      this.lasso.forEach((p, i) => (i ? c.lineTo(p[0], p[1]) : c.moveTo(p[0], p[1])));
      c.stroke(); c.restore();
    }
  }

  /** Saves pending strokes as one ink object. */
  commit(): void {
    clearTimeout(this.timer);
    if (!this.pending.length) return;
    const strokes = this.pending;
    this.pending = [];
    this.redraw();
    if (this.app.isReadOnly()) return;
    saveStrokes(this.app, strokes);
  }

  /** Ink objects on the sheet with their absolute (overlay) rectangles. */
  inkObjects(): { d: DrawingObject; x: number; y: number; z: number }[] {
    const g = this.grid, z = g.zoom;
    return this.sheet.drawings.filter((d) => d.kind === 'ink').map((d) => ({ d, x: g.colX(d.at.c) + d.at.dx * z, y: g.rowY(d.at.r) + d.at.dy * z, z }));
  }

  eraseAt(p: [number, number]): void {
    // Uncommitted strokes first
    const before = this.pending.length;
    this.pending = this.pending.filter((s) => !s.abs.some((q) => Math.hypot(q[0] - p[0], q[1] - p[1]) < 8 + s.width / 2));
    if (this.pending.length !== before) this.redraw();
    for (const o of this.inkObjects()) {
      const { d, x, y, z } = o;
      if (p[0] < x - 8 || p[1] < y - 8 || p[0] > x + d.at.w * z + 8 || p[1] > y + d.at.h * z + 8) continue;
      const ink = parseInk(d.svg);
      const sx = ink.w ? d.at.w / ink.w : 1, sy = ink.h ? d.at.h / ink.h : 1;
      ink.strokes.forEach((s, i) => {
        if (s.points.some(([px, py], k) => {
          const ax = x + px * sx * z, ay = y + py * sy * z;
          if (Math.hypot(ax - p[0], ay - p[1]) < 8 + (s.width * z) / 2) return true;
          if (!k) return false;
          const [qx, qy] = s.points[k - 1];
          return segDist(p, [x + qx * sx * z, y + qy * sy * z], [ax, ay]) < 6 + (s.width * z) / 2;
        })) {
          const set = this.erased.get(d.id) || new Set<number>();
          set.add(i);
          this.erased.set(d.id, set);
          const el = this.grid.overlay.querySelector<HTMLElement>(`.ql-object[data-id="${CSS.escape(d.id)}"] path[data-s="${i}"]`);
          el?.parentElement?.querySelectorAll<SVGElement>(`path[data-s="${i}"]`).forEach((x2) => (x2.style.opacity = '0.15'));
        }
      });
    }
  }

  applyErase(): void {
    if (!this.erased.size) return;
    const erased = this.erased;
    this.erased = new Map();
    const s = this.sheet;
    this.app.book.transact('Erase Ink', () => {
      this.app.book.recordSheet(s);
      for (const [id, set] of erased) {
        const i = s.drawings.findIndex((d) => d.id === id);
        if (i < 0) continue;
        const d = s.drawings[i];
        const ink = parseInk(d.svg);
        const keep = ink.strokes.filter((_x, k) => !set.has(k));
        if (!keep.length) { s.drawings.splice(i, 1); continue; }
        s.drawings[i] = { ...d, svg: inkSvg(keep, ink.w || d.at.w, ink.h || d.at.h) };
      }
    });
  }

  lassoSelect(poly: [number, number][]): void {
    if (poly.length < 3) return;
    const inside = (p: [number, number]) => {
      let c = false;
      for (let i = 0, j = poly.length - 1; i < poly.length; j = i++) {
        const [xi, yi] = poly[i], [xj, yj] = poly[j];
        if (yi > p[1] !== yj > p[1] && p[0] < ((xj - xi) * (p[1] - yi)) / (yj - yi || 1e-9) + xi) c = !c;
      }
      return c;
    };
    const hits = this.inkObjects().filter(({ d, x, y, z }) => inside([x + (d.at.w * z) / 2, y + (d.at.h * z) / 2]));
    if (!hits.length) { toast('No ink inside the lasso.'); return; }
    let id = hits[0].d.id;
    if (hits.length > 1) id = mergeInk(this.app, hits.map((x) => x.d)) || id;
    this.app.exec('draw.select');
    this.app.grid.objects.select(id);
    toast(hits.length > 1 ? `${hits.length} ink drawings grouped and selected` : 'Ink selected');
  }

  close(): void {
    this.commit();
    clearTimeout(this.timer);
    document.removeEventListener('keydown', this.onKey, true);
    window.removeEventListener('resize', this.onResize);
    this.canvas.remove();
  }
}

/** Saves strokes (in overlay coordinates) as a new ink object. Returns its id. */
function saveStrokes(app: App, strokes: LiveStroke[]): string | null {
  const g = app.grid, z = g.zoom;
  let x1 = Infinity, y1 = Infinity, x2 = -Infinity, y2 = -Infinity;
  for (const s of strokes) s.abs.forEach(([x, y], i) => { const pad = (s.widths?.[i] ?? s.width) / 2 + 1; x1 = Math.min(x1, x - pad); y1 = Math.min(y1, y - pad); x2 = Math.max(x2, x + pad); y2 = Math.max(y2, y + pad); });
  if (!Number.isFinite(x1)) return null;
  x1 = Math.max(g.headerW, x1); y1 = Math.max(g.headerH, y1);
  const anchor = g.objects.anchorFor(x1, y1);
  const ax = g.colX(anchor.c) + anchor.dx * z, ay = g.rowY(anchor.r) + anchor.dy * z;
  const w = Math.max(4, (x2 - ax) / z), hh = Math.max(4, (y2 - ay) / z);
  const local: InkStroke[] = strokes.map((s) => ({ tool: s.tool, color: s.color, width: s.width / z, opacity: s.opacity, points: s.abs.map(([x, y]) => [(x - ax) / z, (y - ay) / z] as [number, number]), widths: s.widths?.map((wd) => wd / z) }));
  const d: DrawingObject = { id: newId('ink'), kind: 'ink', at: { ...anchor, w, h: hh }, svg: inkSvg(local, w, hh), alt: 'Ink drawing', stroke: strokes[0].color };
  const s = app.sheet;
  app.book.transact('Ink', () => { app.book.recordSheet(s); s.drawings.push(d); });
  return d.id;
}

/** Combines several ink objects into one (lasso selection). */
function mergeInk(app: App, objs: DrawingObject[]): string | null {
  const g = app.grid, z = g.zoom;
  const strokes: LiveStroke[] = [];
  for (const d of objs) {
    const ink = parseInk(d.svg);
    const ox = g.colX(d.at.c) + d.at.dx * z, oy = g.rowY(d.at.r) + d.at.dy * z;
    const sx = ink.w ? d.at.w / ink.w : 1, sy = ink.h ? d.at.h / ink.h : 1;
    for (const s of ink.strokes) strokes.push({ ...s, width: s.width * z, widths: s.widths?.map((x) => x * z), abs: s.points.map(([x, y]) => [ox + x * sx * z, oy + y * sy * z]) });
  }
  const s = app.sheet;
  let id: string | null = null;
  app.book.transact('Group Ink', () => {
    app.book.recordSheet(s);
    s.drawings = s.drawings.filter((d) => !objs.includes(d));
    id = saveStrokes(app, strokes);
  });
  return id;
}

function addRecognizedShape(app: App, rec: Recognized, color: string, width: number, anchorFor: (x: number, y: number) => { r: number; c: number; dx: number; dy: number }, z: number): void {
  const s = app.sheet;
  let d: DrawingObject;
  if (rec.shape === 'line') {
    const x = Math.min(rec.x1, rec.x2), y = Math.min(rec.y1, rec.y2);
    const w = Math.max(2, Math.abs(rec.x2 - rec.x1) / z), hh = Math.max(2, Math.abs(rec.y2 - rec.y1) / z);
    const up = (rec.x2 - rec.x1) * (rec.y2 - rec.y1) < 0;
    d = { id: newId('shp'), kind: 'shape', shape: 'line', at: { ...anchorFor(x, y), w, h: hh }, fill: 'none', stroke: color, strokeWidth: width, alt: 'Line', ...(up ? { fx: { flipV: true } } : {}) };
  } else {
    d = { id: newId('shp'), kind: 'shape', shape: rec.shape, at: { ...anchorFor(rec.x, rec.y), w: Math.max(4, rec.w / z), h: Math.max(4, rec.h / z) }, fill: 'none', stroke: color, strokeWidth: width, alt: { rect: 'Rectangle', ellipse: 'Oval', triangle: 'Triangle' }[rec.shape] };
  }
  app.book.transact('Ink to Shape', () => { app.book.recordSheet(s); s.drawings.push(d); });
}

let session: InkSession | null = null;

function enter(app: App, mode: Mode): void {
  const s = app.sheet;
  if (s.protection.on && !s.protection.allow.includes('objects')) { app.exec('review.protectionNotice'); return; }
  if (mode === 'pen' || mode === 'pencil' || mode === 'highlighter') lastPen = mode;
  if (session && session.app.grid !== app.grid) { session.close(); session = null; }
  if (session) session.setMode(mode);
  else session = new InkSession(app, mode);
  app.grid.objects.select(null);
  app.ui.status?.(mode === 'eraser' ? 'Eraser: drag across ink to erase it. Esc to stop.' : mode === 'lasso' ? 'Lasso: circle ink to select it. Esc to stop.' : 'Drawing: Esc or Draw ▸ Select to stop.');
  app.refresh();
}

function exitInk(app: App): void {
  if (session) { session.close(); session = null; app.ui.status?.(''); }
  app.refresh();
}

const inMode = (m: Mode) => () => session?.mode === m;

// ============================================================================================
// Ink to shape / math, replay

function lastInk(app: App): DrawingObject | undefined {
  const sel = app.grid.objects.selected;
  const ds = app.sheet.drawings;
  return (sel && ds.find((d) => d.id === sel && d.kind === 'ink')) || [...ds].reverse().find((d) => d.kind === 'ink');
}

function inkToShape(app: App): void {
  session?.commit();
  const d = lastInk(app);
  if (!d) { toast('Draw a shape first (a line, rectangle, oval or triangle), then use Ink to Shape.'); return; }
  const ink = parseInk(d.svg);
  const stroke = ink.strokes[ink.strokes.length - 1];
  const sx = ink.w ? d.at.w / ink.w : 1, sy = ink.h ? d.at.h / ink.h : 1;
  const rec = stroke && recognizeShape(stroke.points.map(([x, y]) => [x * sx, y * sy]));
  if (!rec) { toast('The last stroke does not look like a line, rectangle, oval or triangle.'); return; }
  const s = app.sheet;
  app.book.transact('Ink to Shape', () => {
    app.book.recordSheet(s);
    const i = s.drawings.indexOf(d);
    const rest = ink.strokes.slice(0, -1);
    if (rest.length) s.drawings[i] = { ...d, svg: inkSvg(rest, ink.w || d.at.w, ink.h || d.at.h) };
    else s.drawings.splice(i, 1);
    // Shape coordinates are relative to the ink object's anchor (unzoomed px)
    const base = { r: d.at.r, c: d.at.c };
    const at = (x: number, y: number, w: number, hh: number) => ({ r: base.r, c: base.c, dx: d.at.dx + x, dy: d.at.dy + y, w: Math.max(4, w), h: Math.max(4, hh) });
    let shape: DrawingObject;
    if (rec.shape === 'line') {
      const x = Math.min(rec.x1, rec.x2), y = Math.min(rec.y1, rec.y2);
      const up = (rec.x2 - rec.x1) * (rec.y2 - rec.y1) < 0;
      shape = { id: newId('shp'), kind: 'shape', shape: 'line', at: at(x, y, Math.abs(rec.x2 - rec.x1), Math.abs(rec.y2 - rec.y1)), fill: 'none', stroke: stroke.color, strokeWidth: stroke.width, alt: 'Line', ...(up ? { fx: { flipV: true } } : {}) };
    } else shape = { id: newId('shp'), kind: 'shape', shape: rec.shape, at: at(rec.x, rec.y, rec.w, rec.h), fill: 'none', stroke: stroke.color, strokeWidth: stroke.width, alt: { rect: 'Rectangle', ellipse: 'Oval', triangle: 'Triangle' }[rec.shape] };
    s.drawings.push(shape);
  });
  toast(`Converted to ${{ line: 'a line', rect: 'a rectangle', ellipse: 'an oval', triangle: 'a triangle' }[rec.shape]}`);
}

async function inkToMath(app: App): Promise<void> {
  session?.commit();
  const d = lastInk(app);
  if (!d) { toast('Write an equation in ink first, then use Ink to Math.'); return; }
  const text = await promptDialog('Ink to Math', 'Handwriting recognition is not available offline in Anne\'s Grid, so type the equation you wrote. It replaces the ink as a text box (for example: x = (-b ± √(b²-4ac)) / 2a). To compute it, enter it as a formula in a cell instead.', '');
  if (text == null || !text.trim()) return;
  const s = app.sheet;
  app.book.transact('Ink to Math', () => {
    app.book.recordSheet(s);
    const i = s.drawings.indexOf(d);
    const tb: DrawingObject = { id: newId('txt'), kind: 'textbox', at: { ...d.at, w: Math.max(d.at.w, 120), h: Math.max(28, Math.min(d.at.h, 60)) }, text: text.trim(), font: { font: 'Cambria Math', size: 14, italic: true }, fill: 'none', stroke: 'none', alt: 'Equation' };
    if (i >= 0) s.drawings.splice(i, 1, tb); else s.drawings.push(tb);
  });
}

async function replay(app: App): Promise<void> {
  session?.commit();
  const g = app.grid;
  const sel = g.objects.selected;
  const inks = app.sheet.drawings.filter((d) => d.kind === 'ink' && (!sel || d.id === sel || !app.sheet.drawings.some((x) => x.id === sel && x.kind === 'ink')));
  if (!inks.length) { toast('There is no ink on this sheet to replay.'); return; }
  const z = g.zoom;
  const layer = h('div', { class: 'ql-ink-replay' });
  g.overlay.appendChild(layer);
  const hidden: HTMLElement[] = [];
  const paths: SVGPathElement[] = [];
  for (const d of inks) {
    const el = g.overlay.querySelector<HTMLElement>(`.ql-object[data-id="${CSS.escape(d.id)}"]`);
    if (el) { el.style.visibility = 'hidden'; hidden.push(el); }
    const ink = parseInk(d.svg);
    const box = h('div', { class: 'ql-ink-replay-item', style: `left:${g.colX(d.at.c) + d.at.dx * z}px;top:${g.rowY(d.at.r) + d.at.dy * z}px;width:${d.at.w * z}px;height:${d.at.h * z}px` });
    box.innerHTML = inkSvg(ink.strokes, ink.w || d.at.w, ink.h || d.at.h).replace('<svg ', '<svg width="100%" height="100%" ');
    layer.appendChild(box);
    box.querySelectorAll('path').forEach((p) => { p.setAttribute('pathLength', '1'); p.style.strokeDasharray = '1'; p.style.strokeDashoffset = '1'; paths.push(p); });
  }
  const stop = h('button', { class: 'ql-btn ql-ink-replay-stop', type: 'button' }, 'Stop Replay');
  let cancelled = false;
  stop.addEventListener('click', () => { cancelled = true; });
  layer.appendChild(stop);
  const total = Math.min(8000, Math.max(1200, paths.length * 220));
  const each = total / Math.max(1, paths.length);
  for (const p of paths) {
    if (cancelled) break;
    await new Promise<void>((resolve) => {
      const anim = p.animate?.([{ strokeDashoffset: 1 }, { strokeDashoffset: 0 }], { duration: each, fill: 'forwards', easing: 'linear' });
      if (anim) anim.onfinish = () => resolve(); else { p.style.strokeDashoffset = '0'; setTimeout(resolve, each); }
    });
  }
  await new Promise((r) => setTimeout(r, cancelled ? 0 : 400));
  layer.remove();
  for (const el of hidden) el.style.visibility = '';
}

// ============================================================================================
// Commands

const COLORS: [string, string][] = [['#000000', 'Black'], ['#1f3b57', 'Navy'], ['#2f6fb0', 'Blue'], ['#c00000', 'Red'], ['#2e7d32', 'Green'], ['#7b3fa0', 'Purple'], ['#e07b28', 'Orange'], ['#ffffff', 'White']];
const HIGHLIGHTS: [string, string][] = [['#ffeb3b', 'Yellow'], ['#76ff03', 'Green'], ['#18ffff', 'Turquoise'], ['#ff80ab', 'Pink'], ['#ffab40', 'Orange']];
const WIDTHS: Record<InkTool, number[]> = { pen: [1, 1.5, 2.5, 3.5, 5, 8], pencil: [0.75, 1, 1.5, 2.5, 4], highlighter: [8, 12, 14, 20, 28] };

function penMenu(tool: InkTool): () => MenuEntry[] {
  return () => [
    { heading: 'Color' },
    ...(tool === 'highlighter' ? HIGHLIGHTS : COLORS).map(([c, n]) => ({ id: 'draw.penColor', arg: { tool, color: c }, label: n, swatch: c, checked: pens[tool].color === c })),
    { id: 'draw.penColor', arg: { tool }, label: 'More Colors…', icon: 'palette' },
    { separator: true },
    { heading: 'Thickness' },
    ...WIDTHS[tool].map((w) => ({ id: 'draw.penWidth', arg: { tool, width: w }, label: `${w} pt`, checked: pens[tool].width === w })),
  ];
}

const objectsOk = (app: App) => !(app.sheet.protection.on && !app.sheet.protection.allow.includes('objects'));

def([
  { id: 'draw.select', label: 'Select', icon: 'arrow_selector_tool', tab: 'Draw', group: 'Drawing Tools', kind: 'toggle', noRecord: true, tip: 'Stop drawing and select cells and objects again (Esc)', run: (app) => exitInk(app), isActive: () => !session },
  { id: 'draw.lasso', label: 'Lasso Select', icon: 'gesture_select', tab: 'Draw', group: 'Drawing Tools', kind: 'toggle', noRecord: true, edits: true, tip: 'Draw around ink to select it (several drawings are grouped into one)', run: (app) => (session?.mode === 'lasso' ? exitInk(app) : enter(app, 'lasso')), isActive: inMode('lasso') },
  { id: 'draw.pen', label: 'Pen', icon: 'stylus', tab: 'Draw', group: 'Drawing Tools', kind: 'split', noRecord: true, edits: true, popular: true, tip: 'Draw with a pen; line width follows stylus pressure', keywords: 'ink draw handwriting stylus', run: (app) => (session?.mode === 'pen' ? exitInk(app) : enter(app, 'pen')), isActive: inMode('pen'), menu: penMenu('pen') },
  { id: 'draw.pencil', label: 'Pencil', icon: 'edit', tab: 'Draw', group: 'Drawing Tools', kind: 'split', noRecord: true, edits: true, tip: 'Sketch with a thin pencil line', run: (app) => (session?.mode === 'pencil' ? exitInk(app) : enter(app, 'pencil')), isActive: inMode('pencil'), menu: penMenu('pencil') },
  { id: 'draw.highlighter', label: 'Highlighter', icon: 'ink_highlighter', tab: 'Draw', group: 'Drawing Tools', kind: 'split', noRecord: true, edits: true, tip: 'Mark up cells with a see-through highlighter', run: (app) => (session?.mode === 'highlighter' ? exitInk(app) : enter(app, 'highlighter')), isActive: inMode('highlighter'), menu: penMenu('highlighter') },
  { id: 'draw.eraser', label: 'Eraser', icon: 'ink_eraser', tab: 'Draw', group: 'Drawing Tools', kind: 'toggle', noRecord: true, edits: true, tip: 'Drag across ink strokes to erase them', run: (app) => (session?.mode === 'eraser' ? exitInk(app) : enter(app, 'eraser')), isActive: inMode('eraser') },
  {
    id: 'draw.penColor', label: 'Pen Color', icon: 'palette', tab: 'Draw', group: 'Drawing Tools', noRecord: true, internal: true,
    run: async (app, arg?: { tool?: InkTool; color?: string }) => {
      const tool = arg?.tool || lastPen;
      let color = arg?.color;
      if (!color) { const { colorMenu } = await import('../ui/dialog'); const c = await colorMenu(app.ui.anchorFor(`draw.${tool}`), { title: 'Pen color', highlight: tool === 'highlighter' }); if (!c) return; color = c; }
      pens[tool].color = safeColor(color, pens[tool].color); savePens();
      enter(app, tool);
    },
  },
  { id: 'draw.penWidth', label: 'Pen Thickness', icon: 'line_weight', tab: 'Draw', group: 'Drawing Tools', noRecord: true, internal: true, run: (app, arg?: { tool?: InkTool; width?: number }) => { const tool = arg?.tool || lastPen; if (arg?.width) { pens[tool].width = Math.max(0.25, Math.min(60, arg.width)); savePens(); } enter(app, tool); } },
  { id: 'draw.inkToShape', label: 'Ink to Shape', icon: 'category', tab: 'Draw', group: 'Convert', edits: true, tip: 'Turn the last stroke (or the selected ink) into a straight line, rectangle, oval or triangle', keywords: 'recognize shape convert ink', isEnabled: objectsOk, run: (app) => inkToShape(app) },
  { id: 'draw.autoShapes', label: 'Convert Strokes to Shapes', icon: 'auto_fix_high', tab: 'Draw', group: 'Convert', kind: 'toggle', noRecord: true, tip: 'While on, each line, rectangle, oval or triangle you draw becomes a shape as soon as you lift the pen', run: (app) => { autoShapes = !autoShapes; toast(autoShapes ? 'Strokes that look like shapes become shapes' : 'Strokes stay as ink'); app.refresh(); }, isActive: () => autoShapes },
  { id: 'draw.inkToMath', label: 'Ink to Math', icon: 'function', tab: 'Draw', group: 'Convert', edits: true, tip: 'Replace handwritten math with a typed equation text box (offline handwriting recognition is not available)', keywords: 'equation handwriting math', isEnabled: objectsOk, run: (app) => inkToMath(app) },
  { id: 'draw.replay', label: 'Ink Replay', icon: 'replay', tab: 'Draw', group: 'Replay', noRecord: true, tip: 'Watch the ink on this sheet (or the selected drawing) being drawn again, stroke by stroke', run: (app) => replay(app) },
  {
    id: 'draw.deleteInk', label: 'Delete All Ink', icon: 'delete', tab: 'Draw', group: 'Drawing Tools', edits: true, tip: 'Remove every ink drawing from this sheet', isEnabled: objectsOk,
    run: (app) => { const s = app.sheet; const n = s.drawings.filter((d) => d.kind === 'ink').length; if (!n) { toast('There is no ink on this sheet.'); return; } app.book.transact('Delete All Ink', () => { app.book.recordSheet(s); s.drawings = s.drawings.filter((d) => d.kind !== 'ink'); }); toast(`${n} ink drawing${n === 1 ? '' : 's'} deleted`); },
  },
]);
