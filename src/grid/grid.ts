/**
 * Spreadsheet grid: a canvas renderer with virtual scrolling, frozen panes, selection, fill handle,
 * drag-move, header resizing, in-cell editing and touch support.
 *
 * Layout: the root element holds a <canvas> that paints headers and cells, a transparent scroller
 * (overflow:auto) on top that provides native scrolling (including touch panning and inertia) and
 * receives pointer events, an overlay for charts and pictures, and the in-cell editor.
 */
import type { App, BookTab } from '../app';
import type { Sheet } from '../model/workbook';
import { MAX_COLS, MAX_ROWS, normRange, cellKey, type Range } from '../model/address';
import { Axis } from './axis';
import { paint, checkboxRect, type PaintCache } from './render';
import { CellEditor } from './editor';
import { ObjectLayer } from './objects';
import { keyName } from '../util/keys';
import { h } from '../ui/dom';

/** Finger-sized hit radius of a touch selection handle, in CSS pixels. */
const HANDLE_HIT = 22;

export type Hit =
  | { area: 'corner' }
  | { area: 'colHeader'; c: number; border?: number }
  | { area: 'rowHeader'; r: number; border?: number }
  | { area: 'cell'; r: number; c: number; fill?: boolean; edge?: boolean; filter?: { col: number }; dropdown?: boolean; handle?: 'start' | 'end' }
  | { area: 'none' };

type DragMode =
  | { kind: 'select'; anchor: { r: number; c: number }; add: boolean }
  | { kind: 'cols'; anchor: number; add: boolean }
  | { kind: 'rows'; anchor: number; add: boolean }
  | { kind: 'resizeCol'; c: number; startX: number; startW: number; cols: number[] }
  | { kind: 'resizeRow'; r: number; startY: number; startH: number; rows: number[] }
  | { kind: 'fill'; src: Range; target: Range | null }
  | { kind: 'move'; src: Range; dr: number; dc: number; grabR: number; grabC: number; copy: boolean }
  | { kind: 'ref'; anchor: { r: number; c: number } }
  | { kind: 'handle'; which: 'start' | 'end'; fixed: { r: number; c: number } };

export class Grid {
  el: HTMLElement;
  canvas: HTMLCanvasElement;
  scroller: HTMLElement;
  spacer: HTMLElement;
  overlay: HTMLElement;
  editor: CellEditor;
  objects: ObjectLayer;
  ctx: CanvasRenderingContext2D;
  rows!: Axis;
  cols!: Axis;
  axisVersion = '';
  width = 0;
  height = 0;
  dpr = 1;
  headerW = 40;
  headerH = 20;
  private raf = 0;
  private drag: DragMode | null = null;
  private autoScrollTimer: any = null;
  private lastPointer = { x: 0, y: 0 };
  private resizeObs: ResizeObserver | null = null;
  extent = { rows: 200, cols: 40 };
  cache: PaintCache = { text: new Map(), version: -1 };
  /** Fill-handle / move previews. */
  preview: { range: Range; kind: 'fill' | 'move' } | null = null;
  /** Reference highlights shown while editing a formula. */
  refHighlights: { sheet: string | null; range: Range; color: string }[] = [];
  /** Cell being pointed to while building a formula. */
  pointRange: Range | null = null;
  touchMode = false;
  private tapInfo: { x: number; y: number; t: number; moved: boolean; id: number } | null = null;
  /** Touch long-press that has fired and is waiting for the finger to lift (context menu) or drag (select a range). */
  private pressSel: { cx: number; cy: number; hit: Hit } | null = null;
  private longPressTimer: any = null;
  private pinch: { d: number; z: number } | null = null;
  private touches = new Map<number, { x: number; y: number }>();

  constructor(public app: App, public tab: BookTab) {
    this.el = h('div', { class: 'ql-grid', tabindex: '-1' });
    this.canvas = h('canvas', { class: 'ql-grid-canvas', 'aria-hidden': 'true' });
    this.spacer = h('div', { class: 'ql-grid-spacer' });
    this.scroller = h('div', { class: 'ql-grid-scroller', tabindex: '0', role: 'grid', 'aria-label': 'Worksheet' }, this.spacer);
    this.overlay = h('div', { class: 'ql-grid-overlay' });
    this.el.append(this.canvas, this.scroller, this.overlay);
    this.ctx = this.canvas.getContext('2d')!;
    this.editor = new CellEditor(this);
    this.objects = new ObjectLayer(this);
    this.attach();
  }

  get sheet(): Sheet { return this.tab.book.sheet; }
  get zoom(): number { return this.sheet.zoom || 1; }
  get editing(): boolean { return this.editor.active; }

  destroy(): void {
    cancelAnimationFrame(this.raf);
    this.resizeObs?.disconnect();
    this.editor.destroy();
  }

  focus(): void {
    if (this.editor.active) this.editor.focus();
    else this.scroller.focus({ preventScroll: true });
  }

  // ------------------------------------------------------------------------------------------
  // Geometry

  ensureAxes(): void {
    const s = this.sheet;
    const sig = `${s.id}:${this.tab.book.version}:${s.colW.size}:${s.rowH.size}:${s.hiddenRows.size}:${s.hiddenCols.size}:${s.defaultColW}:${s.defaultRowH}`;
    if (sig === this.axisVersion && this.rows) return;
    this.axisVersion = sig;
    this.cols = new Axis(s.defaultColW, s.colW, s.hiddenCols, MAX_COLS);
    this.rows = new Axis(s.defaultRowH, s.rowH, s.hiddenRows, MAX_ROWS);
  }

  get frozenW(): number { return this.cols.offset(this.sheet.freeze.cols) * this.zoom; }
  get frozenH(): number { return this.rows.offset(this.sheet.freeze.rows) * this.zoom; }
  get scrollX(): number { return this.scroller.scrollLeft; }
  get scrollY(): number { return this.scroller.scrollTop; }

  /** Left edge of a column in canvas CSS px. */
  colX(c: number): number {
    const z = this.zoom, fc = this.sheet.freeze.cols;
    const off = this.cols.offset(c) * z;
    if (c < fc) return this.headerW + off;
    return this.headerW + off - this.scrollX;
  }

  rowY(r: number): number {
    const z = this.zoom, fr = this.sheet.freeze.rows;
    const off = this.rows.offset(r) * z;
    if (r < fr) return this.headerH + off;
    return this.headerH + off - this.scrollY;
  }

  colW(c: number): number { return this.cols.size(c) * this.zoom; }
  rowH(r: number): number { return this.rows.size(r) * this.zoom; }

  colAt(x: number): number {
    const z = this.zoom;
    const fx = this.headerW + this.frozenW;
    if (x < fx) return this.cols.indexAt((x - this.headerW) / z);
    return this.cols.indexAt((x - this.headerW + this.scrollX) / z);
  }

  rowAt(y: number): number {
    const z = this.zoom;
    const fy = this.headerH + this.frozenH;
    if (y < fy) return this.rows.indexAt((y - this.headerH) / z);
    return this.rows.indexAt((y - this.headerH + this.scrollY) / z);
  }

  /** Cell rectangle in canvas CSS px (merged cells cover their whole area). */
  cellRect(r: number, c: number, merged = true): { x: number; y: number; w: number; h: number } {
    let r2 = r, c2 = c;
    if (merged) { const m = this.sheet.mergeAt(r, c); if (m) { r = m.r1; c = m.c1; r2 = m.r2; c2 = m.c2; } }
    const x = this.colX(c), y = this.rowY(r);
    return { x, y, w: this.colX(c2) + this.colW(c2) - x, h: this.rowY(r2) + this.rowH(r2) - y };
  }

  rangeRect(rg: Range): { x: number; y: number; w: number; h: number } {
    const x = this.colX(rg.c1), y = this.rowY(rg.r1);
    return { x, y, w: this.colX(rg.c2) + this.colW(Math.min(rg.c2, MAX_COLS - 1)) - x, h: this.rowY(rg.r2) + this.rowH(Math.min(rg.r2, MAX_ROWS - 1)) - y };
  }

  /** Visible row/col index ranges in the scrolling quadrant. */
  visible(): { r1: number; r2: number; c1: number; c2: number } {
    const fr = this.sheet.freeze.rows, fc = this.sheet.freeze.cols;
    const r1 = Math.max(fr, this.rowAt(this.headerH + this.frozenH));
    const c1 = Math.max(fc, this.colAt(this.headerW + this.frozenW));
    return { r1, c1, r2: this.rowAt(this.height), c2: this.colAt(this.width) };
  }

  // ------------------------------------------------------------------------------------------
  // Layout and painting

  layout(): void {
    const rect = this.el.getBoundingClientRect();
    this.width = Math.max(1, rect.width);
    this.height = Math.max(1, rect.height);
    this.dpr = window.devicePixelRatio || 1;
    this.canvas.width = Math.round(this.width * this.dpr);
    this.canvas.height = Math.round(this.height * this.dpr);
    this.canvas.style.width = `${this.width}px`;
    this.canvas.style.height = `${this.height}px`;
    this.updateExtent();
    this.invalidate();
  }

  /** Scrollable extent: used range plus room to grow, extended as the user scrolls toward the end. */
  updateExtent(): void {
    this.ensureAxes();
    const s = this.sheet;
    const used = s.usedRange();
    const sel = s.sel;
    const vis = this.rows ? this.visible() : { r2: 0, c2: 0 };
    const rows = Math.min(MAX_ROWS, Math.max(this.extent.rows, (used?.r2 ?? 0) + 60, sel.r + 60, vis.r2 + 40, 200));
    const cols = Math.min(MAX_COLS, Math.max(this.extent.cols, (used?.c2 ?? 0) + 15, sel.c + 15, vis.c2 + 10, 30));
    this.extent = { rows, cols };
    const z = this.zoom;
    const fc = s.freeze.cols, fr = s.freeze.rows;
    const w = (this.cols.offset(cols) - this.cols.offset(fc)) * z + this.headerW + this.frozenW;
    const hgt = (this.rows.offset(rows) - this.rows.offset(fr)) * z + this.headerH + this.frozenH;
    // Account for the frozen part: the scroll range covers only the unfrozen area.
    this.spacer.style.width = `${Math.max(w + (this.cols.offset(fc) * z ? 0 : 0), this.width)}px`;
    this.spacer.style.height = `${Math.max(hgt, this.height)}px`;
  }

  invalidate(): void {
    if (this.raf) return;
    this.raf = requestAnimationFrame(() => {
      this.raf = 0;
      this.render();
    });
  }

  render(): void {
    if (!this.width) return;
    this.ensureAxes();
    this.measureHeaders();
    paint(this);
    this.objects.position();
    this.editor.position();
    if (this.tab.marquee) setTimeout(() => this.invalidate(), 120);
  }

  private measureHeaders(): void {
    const z = this.zoom;
    this.headerH = this.sheet.showHeadings ? Math.round(Math.max(16, 20 * z)) : 0;
    if (!this.sheet.showHeadings) { this.headerW = 0; return; }
    const lastRow = this.rowAt(this.height) + 1;
    const digits = String(lastRow).length;
    this.ctx.font = `${Math.round(11 * Math.max(0.8, z) * 4 / 3)}px system-ui, sans-serif`;
    this.headerW = Math.ceil(this.ctx.measureText('0'.repeat(Math.max(2, digits))).width + 14);
  }

  // ------------------------------------------------------------------------------------------
  // Hit testing

  hit(x: number, y: number): Hit {
    if (x < this.headerW && y < this.headerH) return { area: 'corner' };
    const touchSlop = this.touchMode ? 10 : 4;
    if (y < this.headerH) {
      const c = this.colAt(x);
      const right = this.colX(c) + this.colW(c);
      const left = this.colX(c);
      if (Math.abs(x - right) <= touchSlop) return { area: 'colHeader', c, border: c };
      if (Math.abs(x - left) <= touchSlop && c > 0) { let p = c - 1; while (p > 0 && this.cols.isHidden(p)) p--; return { area: 'colHeader', c, border: p }; }
      return { area: 'colHeader', c };
    }
    if (x < this.headerW) {
      const r = this.rowAt(y);
      const bottom = this.rowY(r) + this.rowH(r);
      const top = this.rowY(r);
      if (Math.abs(y - bottom) <= touchSlop - 1) return { area: 'rowHeader', r, border: r };
      if (Math.abs(y - top) <= touchSlop - 1 && r > 0) { let p = r - 1; while (p > 0 && this.rows.isHidden(p)) p--; return { area: 'rowHeader', r, border: p }; }
      return { area: 'rowHeader', r };
    }
    const r = this.rowAt(y), c = this.colAt(x);
    const sel = this.sheet.sel;
    const act = sel.ranges[sel.ranges.length - 1];
    const ar = this.rangeRect(act);
    const out: Hit = { area: 'cell', r, c };
    // fill handle (bottom-right corner of the active range)
    const fx = ar.x + ar.w, fy = ar.y + ar.h;
    const hs = this.touchMode ? 14 : 5;
    if (this.touchMode && this.app.prefs.touchHandles) {
      // Touch: both corners are selection handles (the bottom-right one extends the range, it does
      // not fill), with a finger-sized target. The nearer handle wins on a small range.
      if (sel.ranges.length === 1 && !this.editing) {
        const ds = Math.hypot(x - ar.x, y - ar.y), de = Math.hypot(x - fx, y - fy);
        if (Math.min(ds, de) <= HANDLE_HIT) (out as any).handle = de <= ds ? 'end' : 'start';
      }
    } else if (this.app.prefs.fillHandle && sel.ranges.length === 1 && Math.abs(x - fx) <= hs && Math.abs(y - fy) <= hs) (out as any).fill = true;
    // selection edge (drag to move)
    if (!(out as any).fill && !this.touchMode) {
      const onV = (Math.abs(x - ar.x) <= 3 || Math.abs(x - (ar.x + ar.w)) <= 3) && y >= ar.y - 3 && y <= ar.y + ar.h + 3;
      const onH = (Math.abs(y - ar.y) <= 3 || Math.abs(y - (ar.y + ar.h)) <= 3) && x >= ar.x - 3 && x <= ar.x + ar.w + 3;
      if (onV || onH) (out as any).edge = true;
    }
    // AutoFilter / table header buttons
    const fb = this.filterButtonAt(r, c);
    if (fb) {
      const rc = this.cellRect(r, c);
      const bs = Math.min(18, rc.h - 2);
      if (x >= rc.x + rc.w - bs - 2 && y >= rc.y + rc.h - bs - 1) (out as any).filter = { col: fb.col };
    }
    // Data validation dropdown arrow (active cell only)
    if (r === sel.r && c === sel.c && this.listValidation(r, c)) {
      const rc = this.cellRect(r, c);
      if (x >= rc.x + rc.w && x <= rc.x + rc.w + 18 && y >= rc.y && y <= rc.y + rc.h) (out as any).dropdown = true;
    }
    return out;
  }

  /** Whether (x, y) is on the in-cell checkbox of (r, c) (Insert ▸ Checkbox). */
  checkboxAt(r: number, c: number, x: number, y: number): boolean {
    const st = this.sheet.styleOf(r, c);
    if (!st.checkbox || typeof this.sheet.value(r, c) !== 'boolean') return false;
    const b = checkboxRect(this.cellRect(r, c), st, this.zoom);
    const pad = this.touchMode ? 6 : 2;
    return x >= b.x - pad && x <= b.x + b.s + pad && y >= b.y - pad && y <= b.y + b.s + pad;
  }

  /** AutoFilter / table filter button for a header cell. */
  filterButtonAt(r: number, c: number): { col: number; source: 'filter' | 'table'; tableId?: string } | null {
    const f = this.sheet.filter;
    if (f && r === f.range.r1 && c >= f.range.c1 && c <= f.range.c2) return { col: c - f.range.c1, source: 'filter' };
    for (const t of this.sheet.tables) if (t.opts.header && t.opts.filterButton && r === t.range.r1 && c >= t.range.c1 && c <= t.range.c2) return { col: c - t.range.c1, source: 'table', tableId: t.id };
    return null;
  }

  listValidation(r: number, c: number) {
    return this.sheet.validations.find((v) => v.type === 'list' && v.showDropdown !== false && v.ranges.some((rg) => r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2));
  }

  // ------------------------------------------------------------------------------------------
  // Selection

  get sel() { return this.sheet.sel; }
  get activeRange(): Range { const s = this.sel; return s.ranges[s.ranges.length - 1]; }

  /** Selects a cell or range; extends merged cells so the selection always covers them fully. */
  select(r: number, c: number, opts: { extendTo?: { r: number; c: number }; add?: boolean; range?: Range; keepActive?: boolean; noScroll?: boolean } = {}): void {
    const s = this.sheet;
    r = clamp(r, 0, MAX_ROWS - 1); c = clamp(c, 0, MAX_COLS - 1);
    let rg: Range;
    if (opts.range) rg = normRange(opts.range);
    else if (opts.extendTo) rg = normRange({ r1: r, c1: c, r2: opts.extendTo.r, c2: opts.extendTo.c });
    else rg = { r1: r, c1: c, r2: r, c2: c };
    rg = this.expandMerges(rg);
    if (opts.add) s.sel.ranges.push(rg);
    else s.sel.ranges = [rg];
    if (!opts.keepActive) { s.sel.r = r; s.sel.c = c; }
    if (!opts.noScroll) this.scrollIntoView(opts.extendTo ? opts.extendTo.r : r, opts.extendTo ? opts.extendTo.c : c);
    if (this.app.recording) this.app.recordStep('nav.select', { ranges: s.sel.ranges, r: s.sel.r, c: s.sel.c });
    this.tab.circles = this.tab.circles;
    this.invalidate();
    this.app.refresh();
  }

  expandMerges(rg: Range): Range {
    let changed = true;
    let out = { ...rg };
    let guard = 0;
    while (changed && guard++ < 20) {
      changed = false;
      for (const m of this.sheet.merges) {
        if (m.r1 <= out.r2 && m.r2 >= out.r1 && m.c1 <= out.c2 && m.c2 >= out.c1) {
          const n = { r1: Math.min(out.r1, m.r1), c1: Math.min(out.c1, m.c1), r2: Math.max(out.r2, m.r2), c2: Math.max(out.c2, m.c2) };
          if (n.r1 !== out.r1 || n.c1 !== out.c1 || n.r2 !== out.r2 || n.c2 !== out.c2) { out = n; changed = true; }
        }
      }
    }
    return out;
  }

  scrollIntoView(r: number, c: number): void {
    this.ensureAxes();
    const s = this.sheet;
    if (r >= this.extent.rows - 5 || c >= this.extent.cols - 3) { this.extent = { rows: Math.max(this.extent.rows, r + 60), cols: Math.max(this.extent.cols, c + 15) }; this.updateExtent(); }
    const z = this.zoom;
    if (r >= s.freeze.rows) {
      const top = (this.rows.offset(r) - this.rows.offset(s.freeze.rows)) * z;
      const bottom = top + this.rowH(r);
      const viewH = this.height - this.headerH - this.frozenH - 16;
      const base = this.rows.offset(s.freeze.rows) * z; void base;
      if (top < this.scrollY) this.scroller.scrollTop = top;
      else if (bottom > this.scrollY + viewH) this.scroller.scrollTop = bottom - viewH;
    }
    if (c >= s.freeze.cols) {
      const left = (this.cols.offset(c) - this.cols.offset(s.freeze.cols)) * z;
      const right = left + this.colW(c);
      const viewW = this.width - this.headerW - this.frozenW - 16;
      if (left < this.scrollX) this.scroller.scrollLeft = left;
      else if (right > this.scrollX + viewW) this.scroller.scrollLeft = right - viewW;
    }
    this.invalidate();
  }

  /** Converts scroll offsets to account for frozen panes: scroll positions are measured from the first unfrozen row/col. */
  private fixScrollOrigin(): void { /* the spacer includes the frozen area; colX/rowY subtract scroll only for unfrozen cells */ }

  // ------------------------------------------------------------------------------------------
  // Events

  private attach(): void {
    const sc = this.scroller;
    sc.addEventListener('scroll', () => {
      // Grow the scroll extent near the end so the sheet feels unbounded.
      const nearBottom = sc.scrollTop + sc.clientHeight > sc.scrollHeight - 400;
      const nearRight = sc.scrollLeft + sc.clientWidth > sc.scrollWidth - 300;
      if ((nearBottom && this.extent.rows < MAX_ROWS) || (nearRight && this.extent.cols < MAX_COLS)) {
        this.extent = { rows: nearBottom ? Math.min(MAX_ROWS, this.extent.rows + 200) : this.extent.rows, cols: nearRight ? Math.min(MAX_COLS, this.extent.cols + 20) : this.extent.cols };
        this.updateExtent();
      }
      this.sheet.scroll = { r: this.rowAt(this.headerH + this.frozenH), c: this.colAt(this.headerW + this.frozenW) };
      this.render();
    }, { passive: true });
    sc.addEventListener('pointerdown', (e) => this.onPointerDown(e));
    sc.addEventListener('pointermove', (e) => this.onPointerMove(e));
    sc.addEventListener('pointerup', (e) => this.onPointerUp(e));
    sc.addEventListener('pointercancel', (e) => this.onPointerUp(e, true));
    // The scroller pans natively (touch-action: pan-x pan-y), so a finger drag that is meant to
    // select has to claim the gesture here, or the browser starts scrolling and cancels the pointer.
    // pointerdown runs before touchstart, so a handle or fill drag is already known by then.
    sc.addEventListener('touchstart', (e) => { if (this.drag && e.cancelable) e.preventDefault(); }, { passive: false });
    sc.addEventListener('touchmove', (e) => { if ((this.drag || this.pressSel) && e.cancelable) e.preventDefault(); }, { passive: false });
    sc.addEventListener('dblclick', (e) => this.onDblClick(e));
    sc.addEventListener('contextmenu', (e) => { e.preventDefault(); if (this.touchMode && (this.tapInfo || this.pressSel || this.drag)) return; this.contextMenu(e.clientX, e.clientY); });
    sc.addEventListener('keydown', (e) => this.onKeyDown(e));
    sc.addEventListener('wheel', (e) => {
      if (e.ctrlKey || (this.app.prefs.zoomOnRoll && !e.shiftKey)) {
        e.preventDefault();
        this.setZoom(this.zoom * (e.deltaY < 0 ? 1.1 : 1 / 1.1));
      }
    }, { passive: false });
    sc.addEventListener('copy', (e) => { if (!this.editing) { e.preventDefault(); this.app.exec('edit.copy', { event: e }); } });
    sc.addEventListener('cut', (e) => { if (!this.editing) { e.preventDefault(); this.app.exec('edit.cut', { event: e }); } });
    sc.addEventListener('paste', (e) => { if (!this.editing) { e.preventDefault(); this.app.exec('edit.paste', { event: e }); } });
    sc.addEventListener('dragover', (e) => { if (e.dataTransfer?.types.includes('Files')) e.preventDefault(); });
    sc.addEventListener('drop', (e) => {
      const files = e.dataTransfer?.files;
      if (files && files.length) { e.preventDefault(); this.app.exec('file.openDropped', [...files]); }
    });
    this.resizeObs = new ResizeObserver(() => this.layout());
    this.resizeObs.observe(this.el);
  }

  private local(e: { clientX: number; clientY: number }): { x: number; y: number } {
    const r = this.el.getBoundingClientRect();
    return { x: e.clientX - r.left, y: e.clientY - r.top };
  }

  private onPointerDown(e: PointerEvent): void {
    const { x, y } = this.local(e);
    this.lastPointer = { x, y };
    this.touchMode = e.pointerType === 'touch' || e.pointerType === 'pen';
    if (e.pointerType === 'touch') {
      this.touches.set(e.pointerId, { x, y });
      if (this.touches.size === 2) {
        const [a, b] = [...this.touches.values()];
        this.pinch = { d: Math.hypot(a.x - b.x, a.y - b.y), z: this.zoom };
        clearTimeout(this.longPressTimer);
        this.tapInfo = null;
        this.pressSel = null;
        if (this.drag && (this.drag.kind === 'select' || this.drag.kind === 'handle' || this.drag.kind === 'rows' || this.drag.kind === 'cols')) this.drag = null;
        return;
      }
    }
    const hit = this.hit(x, y);
    if (e.button === 2) {
      // Right-click inside the selection keeps it; elsewhere selects the cell first.
      if (hit.area === 'cell' && !this.inSelection(hit.r, hit.c)) this.select(hit.r, hit.c, { noScroll: true });
      if (hit.area === 'colHeader' && !this.wholeColsSelected().includes(hit.c)) this.selectCols(hit.c, hit.c, false);
      if (hit.area === 'rowHeader' && !this.wholeRowsSelected().includes(hit.r)) this.selectRows(hit.r, hit.r, false);
      return;
    }
    // Touch: let the browser pan; decide on pointerup whether it was a tap.
    if (e.pointerType === 'touch' && !(hit.area === 'cell' && (hit.fill || hit.handle))) {
      this.tapInfo = { x, y, t: Date.now(), moved: false, id: e.pointerId };
      this.pressSel = null;
      clearTimeout(this.longPressTimer);
      this.longPressTimer = setTimeout(() => {
        if (this.tapInfo && !this.tapInfo.moved) {
          const ht = this.hit(x, y);
          if (ht.area === 'cell' && !this.inSelection(ht.r, ht.c) && !this.editor.active) this.select(ht.r, ht.c, { noScroll: true });
          if (this.editor.active || ht.area === 'corner') { this.contextMenu(e.clientX, e.clientY); this.tapInfo = null; return; }
          // Press and hold, then either lift (context menu) or keep dragging (select a range).
          this.pressSel = { cx: e.clientX, cy: e.clientY, hit: ht };
          try { navigator.vibrate?.(12); } catch { /* */ }
        }
      }, 550);
      return;
    }
    // A touch on a selection handle is still a tap until the finger moves.
    if (e.pointerType === 'touch' && hit.area === 'cell' && hit.handle) { clearTimeout(this.longPressTimer); this.pressSel = null; this.tapInfo = { x, y, t: Date.now(), moved: false, id: e.pointerId }; }
    e.preventDefault();
    this.scroller.setPointerCapture(e.pointerId);
    const add = e.ctrlKey || e.metaKey;
    if (this.editor.active && this.editor.canPoint()) {
      if (hit.area === 'cell') { this.drag = { kind: 'ref', anchor: { r: hit.r, c: hit.c } }; this.editor.pointAt({ r1: hit.r, c1: hit.c, r2: hit.r, c2: hit.c }); return; }
    } else if (this.editor.active) this.commitEdit('none');
    switch (hit.area) {
      case 'corner': this.select(0, 0, { range: { r1: 0, c1: 0, r2: MAX_ROWS - 1, c2: MAX_COLS - 1 }, keepActive: true, noScroll: true }); break;
      case 'colHeader':
        if (hit.border != null) {
          const cols = this.wholeColsSelected().includes(hit.border) ? this.wholeColsSelected() : [hit.border];
          this.drag = { kind: 'resizeCol', c: hit.border, startX: x, startW: this.cols.size(hit.border), cols };
        } else {
          if (e.shiftKey) this.selectCols(this.sel.c, hit.c, false);
          else { this.selectCols(hit.c, hit.c, add); this.drag = { kind: 'cols', anchor: hit.c, add }; }
        }
        break;
      case 'rowHeader':
        if (hit.border != null) {
          const rows = this.wholeRowsSelected().includes(hit.border) ? this.wholeRowsSelected() : [hit.border];
          this.drag = { kind: 'resizeRow', r: hit.border, startY: y, startH: this.rows.size(hit.border), rows };
        } else {
          if (e.shiftKey) this.selectRows(this.sel.r, hit.r, false);
          else { this.selectRows(hit.r, hit.r, add); this.drag = { kind: 'rows', anchor: hit.r, add }; }
        }
        break;
      case 'cell': {
        if (hit.filter) { this.app.exec('data.filterMenu', { r: hit.r, c: hit.c }); return; }
        if (hit.dropdown) { this.app.exec('data.validationList', { r: hit.r, c: hit.c }); return; }
        if (this.checkboxAt(hit.r, hit.c, x, y)) { this.select(hit.r, hit.c, { noScroll: true }); this.app.exec('cell.toggleCheckbox', { r: hit.r, c: hit.c }); return; }
        if (hit.fill) { this.drag = { kind: 'fill', src: { ...this.activeRange }, target: null }; return; }
        if (hit.handle) {
          const ar = this.activeRange;
          this.drag = { kind: 'handle', which: hit.handle, fixed: hit.handle === 'start' ? { r: ar.r2, c: ar.c2 } : { r: ar.r1, c: ar.c1 } };
          return;
        }
        if (hit.edge && !e.shiftKey) { const ar = this.activeRange; this.drag = { kind: 'move', src: { ...ar }, dr: 0, dc: 0, grabR: hit.r, grabC: hit.c, copy: add }; return; }
        if (e.shiftKey) { this.select(this.sel.r, this.sel.c, { extendTo: { r: hit.r, c: hit.c }, keepActive: true }); this.drag = { kind: 'select', anchor: { r: this.sel.r, c: this.sel.c }, add: false }; }
        else {
          this.select(hit.r, hit.c, { add, noScroll: true });
          this.drag = { kind: 'select', anchor: { r: hit.r, c: hit.c }, add };
          // Ctrl+click on a hyperlink follows it.
          const link = this.sheet.get(hit.r, hit.c)?.link;
          if (link && (e.ctrlKey || e.metaKey || this.app.isReadOnly())) { this.app.followLink(link); this.drag = null; }
        }
        break;
      }
    }
  }

  private onPointerMove(e: PointerEvent): void {
    const { x, y } = this.local(e);
    this.lastPointer = { x, y };
    if (e.pointerType === 'touch') {
      if (this.touches.has(e.pointerId)) this.touches.set(e.pointerId, { x, y });
      if (this.pinch && this.touches.size === 2) {
        const [a, b] = [...this.touches.values()];
        const d = Math.hypot(a.x - b.x, a.y - b.y);
        this.setZoom(this.pinch.z * (d / this.pinch.d), true);
        return;
      }
      if (this.tapInfo && Math.hypot(x - this.tapInfo.x, y - this.tapInfo.y) > 8) { this.tapInfo.moved = true; clearTimeout(this.longPressTimer); }
      if (this.pressSel && !this.drag && this.tapInfo?.moved) {
        // Held, then dragged: select from the pressed cell (or row / column heading) to the finger.
        const ht = this.pressSel.hit;
        this.pressSel = null;
        if (ht.area === 'cell') { this.select(ht.r, ht.c, { noScroll: true }); this.drag = { kind: 'select', anchor: { r: ht.r, c: ht.c }, add: false }; }
        else if (ht.area === 'colHeader') { this.selectCols(ht.c, ht.c, false); this.drag = { kind: 'cols', anchor: ht.c, add: false }; }
        else if (ht.area === 'rowHeader') { this.selectRows(ht.r, ht.r, false); this.drag = { kind: 'rows', anchor: ht.r, add: false }; }
      }
      if (!this.drag) return;
      if (this.drag.kind === 'handle' && this.tapInfo && !this.tapInfo.moved) return;
    }
    if (!this.drag) { this.updateCursor(x, y); return; }
    this.dragTo(x, y);
    this.autoScroll(x, y);
  }

  private dragTo(x: number, y: number): void {
    const d = this.drag!;
    const r = clamp(this.rowAt(Math.max(this.headerH, y)), 0, MAX_ROWS - 1), c = clamp(this.colAt(Math.max(this.headerW, x)), 0, MAX_COLS - 1);
    switch (d.kind) {
      case 'select': {
        const s = this.sheet.sel;
        const rg = this.expandMerges(normRange({ r1: d.anchor.r, c1: d.anchor.c, r2: r, c2: c }));
        s.ranges[s.ranges.length - 1] = rg;
        this.invalidate(); this.app.refresh();
        break;
      }
      case 'handle': {
        const s = this.sheet.sel;
        s.ranges = [this.expandMerges(normRange({ r1: d.fixed.r, c1: d.fixed.c, r2: r, c2: c }))];
        this.invalidate(); this.app.refresh();
        break;
      }
      case 'cols': this.selectCols(d.anchor, c, d.add, true); break;
      case 'rows': this.selectRows(d.anchor, r, d.add, true); break;
      case 'resizeCol': {
        const w = Math.max(0, Math.round(d.startW + (x - d.startX) / this.zoom));
        this.preview = null;
        this.resizePreview = { kind: 'col', index: d.c, size: w };
        this.invalidate();
        break;
      }
      case 'resizeRow': {
        const hh = Math.max(0, Math.round(d.startH + (y - d.startY) / this.zoom));
        this.resizePreview = { kind: 'row', index: d.r, size: hh };
        this.invalidate();
        break;
      }
      case 'fill': {
        const src = d.src;
        // Fill extends along the axis with the larger distance outside the source.
        const dr = r < src.r1 ? r - src.r1 : r > src.r2 ? r - src.r2 : 0;
        const dc = c < src.c1 ? c - src.c1 : c > src.c2 ? c - src.c2 : 0;
        let target: Range | null = null;
        if (Math.abs(dr) >= Math.abs(dc) && dr) target = dr > 0 ? { ...src, r2: r } : { ...src, r1: r };
        else if (dc) target = dc > 0 ? { ...src, c2: c } : { ...src, c1: c };
        else if (r >= src.r1 && r <= src.r2 && c >= src.c1 && c <= src.c2 && (r < src.r2 || c < src.c2)) {
          // Dragging back inside the selection clears (shrinks) it.
          target = r < src.r2 ? { ...src, r2: r } : { ...src, c2: c };
        }
        d.target = target;
        this.preview = target ? { range: target, kind: 'fill' } : null;
        this.invalidate();
        break;
      }
      case 'move': {
        d.dr = r - d.grabR; d.dc = c - d.grabC;
        const t = { r1: d.src.r1 + d.dr, c1: d.src.c1 + d.dc, r2: d.src.r2 + d.dr, c2: d.src.c2 + d.dc };
        if (t.r1 < 0) { d.dr -= t.r1; } if (t.c1 < 0) { d.dc -= t.c1; }
        this.preview = { range: { r1: d.src.r1 + d.dr, c1: d.src.c1 + d.dc, r2: d.src.r2 + d.dr, c2: d.src.c2 + d.dc }, kind: 'move' };
        this.invalidate();
        break;
      }
      case 'ref': this.editor.pointAt(normRange({ r1: d.anchor.r, c1: d.anchor.c, r2: r, c2: c })); break;
    }
  }

  resizePreview: { kind: 'col' | 'row'; index: number; size: number } | null = null;

  private autoScroll(x: number, y: number): void {
    clearInterval(this.autoScrollTimer);
    const edge = 24;
    let dx = 0, dy = 0;
    if (x > this.width - edge) dx = 30; else if (x < this.headerW + this.frozenW + 4 && this.drag?.kind !== 'resizeCol') dx = -30;
    if (y > this.height - edge) dy = 30; else if (y < this.headerH + this.frozenH + 4 && this.drag?.kind !== 'resizeRow') dy = -30;
    if (!dx && !dy) return;
    if (this.drag?.kind === 'resizeCol' || this.drag?.kind === 'resizeRow') return;
    this.autoScrollTimer = setInterval(() => {
      if (!this.drag) { clearInterval(this.autoScrollTimer); return; }
      this.scroller.scrollLeft += dx; this.scroller.scrollTop += dy;
      this.dragTo(this.lastPointer.x, this.lastPointer.y);
    }, 50);
  }

  private onPointerUp(e: PointerEvent, cancel = false): void {
    clearInterval(this.autoScrollTimer);
    if (e.pointerType === 'touch') {
      this.touches.delete(e.pointerId);
      if (this.pinch) { if (this.touches.size < 2) { this.pinch = null; this.app.refresh(); } return; }
      clearTimeout(this.longPressTimer);
      const t = this.tapInfo;
      const press = this.pressSel;
      this.tapInfo = null;
      this.pressSel = null;
      // A handle that was touched but not dragged is a tap on the cell underneath.
      if (t && !t.moved && this.drag?.kind === 'handle') this.drag = null;
      if (t && !t.moved && !cancel && !this.drag) {
        if (press) { this.contextMenu(press.cx, press.cy); return; }
        const { x, y } = this.local(e);
        this.onTap(x, y, e);
        return;
      }
    }
    const d = this.drag;
    this.drag = null;
    if (!d || cancel) { this.preview = null; this.resizePreview = null; this.invalidate(); return; }
    switch (d.kind) {
      case 'resizeCol': {
        const p = this.resizePreview; this.resizePreview = null;
        if (p) this.app.exec('format.colWidth', { cols: d.cols, px: p.size });
        break;
      }
      case 'resizeRow': {
        const p = this.resizePreview; this.resizePreview = null;
        if (p) this.app.exec('format.rowHeight', { rows: d.rows, px: p.size });
        break;
      }
      case 'fill': {
        this.preview = null;
        if (d.target) this.app.exec('edit.fillDrag', { src: d.src, target: d.target, copy: e.ctrlKey || e.metaKey });
        break;
      }
      case 'move': {
        this.preview = null;
        if (d.dr || d.dc) this.app.exec('edit.moveRange', { src: d.src, dr: d.dr, dc: d.dc, copy: d.copy || e.ctrlKey || e.metaKey });
        break;
      }
      case 'ref': this.editor.focus(); break;
    }
    this.invalidate();
    this.app.refresh();
    if (d.kind !== 'ref') this.focus();
  }

  private lastTap = { t: 0, r: -1, c: -1 };

  private onTap(x: number, y: number, e: PointerEvent): void {
    const hit = this.hit(x, y);
    if (this.editor.active && this.editor.canPoint() && hit.area === 'cell') { this.editor.pointAt({ r1: hit.r, c1: hit.c, r2: hit.r, c2: hit.c }); return; }
    if (this.editor.active) this.commitEdit('none');
    if (hit.area === 'cell') {
      if (hit.filter) { this.app.exec('data.filterMenu', { r: hit.r, c: hit.c }); return; }
      if (hit.dropdown) { this.app.exec('data.validationList', { r: hit.r, c: hit.c }); return; }
      if (this.checkboxAt(hit.r, hit.c, x, y)) { this.select(hit.r, hit.c, { noScroll: true }); this.app.exec('cell.toggleCheckbox', { r: hit.r, c: hit.c }); return; }
      const now = Date.now();
      const dbl = now - this.lastTap.t < 350 && this.lastTap.r === hit.r && this.lastTap.c === hit.c;
      this.lastTap = { t: now, r: hit.r, c: hit.c };
      if (dbl) { this.startEdit({ mode: 'edit' }); return; }
      const link = this.sheet.get(hit.r, hit.c)?.link;
      if (link && this.inSelection(hit.r, hit.c) && this.sel.r === hit.r && this.sel.c === hit.c) { this.app.followLink(link); return; }
      this.select(hit.r, hit.c, { noScroll: true });
    } else if (hit.area === 'colHeader') this.selectCols(hit.c, hit.c, false);
    else if (hit.area === 'rowHeader') this.selectRows(hit.r, hit.r, false);
    else if (hit.area === 'corner') this.select(0, 0, { range: { r1: 0, c1: 0, r2: MAX_ROWS - 1, c2: MAX_COLS - 1 }, keepActive: true, noScroll: true });
    void e;
  }

  private onDblClick(e: MouseEvent): void {
    const { x, y } = this.local(e);
    const hit = this.hit(x, y);
    if (hit.area === 'colHeader' && hit.border != null) { this.app.exec('format.autofitCols', { cols: this.wholeColsSelected().includes(hit.border) ? this.wholeColsSelected() : [hit.border] }); return; }
    if (hit.area === 'rowHeader' && hit.border != null) { this.app.exec('format.autofitRows', { rows: this.wholeRowsSelected().includes(hit.border) ? this.wholeRowsSelected() : [hit.border] }); return; }
    if (hit.area === 'cell') {
      if (hit.fill) { this.app.exec('edit.fillDown', { auto: true }); return; }
      if (hit.edge) { this.app.exec('nav.edge', { dir: edgeDir(x, y, this.rangeRect(this.activeRange)) }); return; }
      const pv = this.sheet.pivots.find((pt) => pt.lastRange && hit.r >= pt.lastRange.r1 && hit.r <= pt.lastRange.r2 && hit.c >= pt.lastRange.c1 && hit.c <= pt.lastRange.c2);
      if (pv) { this.app.exec('pivot.drillDown', { r: hit.r, c: hit.c }); return; }
      if (this.sheet.get(hit.r, hit.c)?.link && !this.app.prefs.editInCell) { this.app.followLink(this.sheet.get(hit.r, hit.c)!.link!); return; }
      this.startEdit({ mode: 'edit' });
    }
  }

  private updateCursor(x: number, y: number): void {
    const hit = this.hit(x, y);
    let cur = 'cell';
    if (hit.area === 'colHeader') cur = hit.border != null ? 'col-resize' : 's-resize-arrow';
    else if (hit.area === 'rowHeader') cur = hit.border != null ? 'row-resize' : 'e-resize-arrow';
    else if (hit.area === 'cell') {
      if (hit.fill) cur = 'crosshair';
      else if (hit.edge) cur = 'move';
      else if (hit.filter || hit.dropdown) cur = 'pointer';
      else if (this.sheet.get(hit.r, hit.c)?.link) cur = 'pointer';
    } else if (hit.area === 'corner') cur = 'default';
    this.scroller.dataset.cursor = cur;
    // Comment / note hover
    if (hit.area === 'cell') this.objects.hoverComment(hit.r, hit.c);
    else this.objects.hoverComment(-1, -1);
  }

  inSelection(r: number, c: number): boolean {
    return this.sel.ranges.some((rg) => r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2);
  }

  wholeColsSelected(): number[] {
    const out: number[] = [];
    for (const rg of this.sel.ranges) if (rg.r1 === 0 && rg.r2 >= MAX_ROWS - 1) for (let c = rg.c1; c <= rg.c2; c++) out.push(c);
    return out;
  }

  wholeRowsSelected(): number[] {
    const out: number[] = [];
    for (const rg of this.sel.ranges) if (rg.c1 === 0 && rg.c2 >= MAX_COLS - 1) for (let r = rg.r1; r <= rg.r2; r++) out.push(r);
    return out;
  }

  selectCols(a: number, b: number, add: boolean, dragging = false): void {
    const rg = { r1: 0, c1: Math.min(a, b), r2: MAX_ROWS - 1, c2: Math.max(a, b) };
    const s = this.sheet.sel;
    if (dragging) s.ranges[s.ranges.length - 1] = rg;
    else if (add) s.ranges.push(rg);
    else s.ranges = [rg];
    if (!dragging) { s.r = this.sheet.scroll.r >= this.sheet.freeze.rows ? Math.max(this.rowAt(this.headerH + this.frozenH), 0) : 0; s.c = a; }
    this.invalidate(); this.app.refresh();
  }

  selectRows(a: number, b: number, add: boolean, dragging = false): void {
    const rg = { r1: Math.min(a, b), c1: 0, r2: Math.max(a, b), c2: MAX_COLS - 1 };
    const s = this.sheet.sel;
    if (dragging) s.ranges[s.ranges.length - 1] = rg;
    else if (add) s.ranges.push(rg);
    else s.ranges = [rg];
    if (!dragging) { s.r = a; s.c = Math.max(this.colAt(this.headerW + this.frozenW), 0); }
    this.invalidate(); this.app.refresh();
  }

  setZoom(z: number, live = false): void {
    const nz = Math.max(0.1, Math.min(4, z));
    if (Math.abs(nz - this.zoom) < 0.001) return;
    this.sheet.zoom = Math.round(nz * 100) / 100;
    this.axisVersion = '';
    this.updateExtent();
    this.invalidate();
    if (!live) this.app.refresh();
    else this.app.ui.renderStatus?.();
  }

  contextMenu(x: number, y: number): void {
    const { x: lx, y: ly } = this.local({ clientX: x, clientY: y });
    const hit = this.hit(lx, ly);
    this.app.ui.gridContextMenu(x, y, hit);
  }

  // ------------------------------------------------------------------------------------------
  // Keyboard

  private onKeyDown(e: KeyboardEvent): void {
    if (this.editor.active) return; // the editor handles its own keys
    const k = keyName(e);
    if (!k) return;
    const s = this.sel;
    const shift = e.shiftKey;
    const ctrl = e.ctrlKey || e.metaKey;
    // Space toggles an in-cell checkbox
    if (e.key === ' ' && !ctrl && !e.altKey && !shift && this.sheet.styleOf(s.r, s.c).checkbox) { e.preventDefault(); this.app.exec('cell.toggleCheckbox', { r: s.r, c: s.c }); return; }
    // 1-2-3 style: "/" opens the slash menu when that option is on.
    if (e.key === '/' && !ctrl && !e.altKey && this.app.prefs.slashKey) { e.preventDefault(); this.app.exec('tools.slashMenu'); return; }
    // Typing starts editing
    if (!ctrl && !e.altKey && e.key.length === 1 && !this.app.isReadOnly()) {
      e.preventDefault();
      this.startEdit({ mode: 'enter', text: e.key });
      return;
    }
    const move = (dr: number, dc: number) => {
      if (shift) this.extendSelection(dr, dc, ctrl);
      else this.moveActive(dr, dc, ctrl);
    };
    switch (e.key) {
      case 'ArrowUp': e.preventDefault(); move(-1, 0); return;
      case 'ArrowDown': e.preventDefault(); move(1, 0); return;
      case 'ArrowLeft': e.preventDefault(); move(0, -1); return;
      case 'ArrowRight': e.preventDefault(); move(0, 1); return;
      case 'PageDown': e.preventDefault(); if (e.altKey) this.page(0, 1, shift); else this.page(1, 0, shift); return;
      case 'PageUp': e.preventDefault(); if (e.altKey) this.page(0, -1, shift); else this.page(-1, 0, shift); return;
      case 'Home':
        e.preventDefault();
        if (ctrl) { if (shift) this.select(s.r, s.c, { extendTo: { r: this.sheet.freeze.rows, c: this.sheet.freeze.cols }, keepActive: true }); else this.select(this.sheet.freeze.rows, this.sheet.freeze.cols); }
        else if (shift) this.select(s.r, s.c, { extendTo: { r: this.activeEnd().r, c: 0 }, keepActive: true });
        else this.select(s.r, this.sheet.freeze.cols && this.sheet.freeze.cols <= s.c ? this.sheet.freeze.cols : 0);
        return;
      case 'End':
        if (ctrl) { e.preventDefault(); const u = this.sheet.usedRange(false) || { r2: 0, c2: 0 }; if (shift) this.select(s.r, s.c, { extendTo: { r: u.r2, c: u.c2 }, keepActive: true }); else this.select(u.r2, u.c2); return; }
        break;
      case 'Enter':
        if (!ctrl && !e.altKey) { e.preventDefault(); this.cycleInSelection(shift ? -1 : 1, this.app.prefs.enterDirection); return; }
        break;
      case 'Tab':
        if (!ctrl && !e.altKey) { e.preventDefault(); this.cycleInSelection(shift ? -1 : 1, 'right'); return; }
        break;
      case 'F2': if (!ctrl && !shift && !e.altKey) { e.preventDefault(); this.startEdit({ mode: 'edit' }); return; } break;
      case 'Delete': if (!ctrl && !shift && !e.altKey) { e.preventDefault(); this.app.exec('edit.clearContents'); return; } break;
      case 'Backspace': if (!ctrl && !shift && !e.altKey) { e.preventDefault(); this.startEdit({ mode: 'enter', text: '' }); return; } break;
      case 'Escape': if (this.tab.marquee) { this.tab.marquee = null; this.invalidate(); return; } break;
      case 'ContextMenu': { e.preventDefault(); const rc = this.cellRect(s.r, s.c); const b = this.el.getBoundingClientRect(); this.contextMenu(b.left + rc.x + rc.w / 2, b.top + rc.y + rc.h / 2); return; }
    }
    // Let the browser raise copy/cut/paste events: they carry the clipboard data without permission prompts.
    if (k === 'Mod-c' || k === 'Mod-x' || k === 'Mod-v' || k === 'Mod-Insert' || k === 'Shift-Insert' || k === 'Shift-Delete') return;
    if (e.key === 'F10' && shift) { e.preventDefault(); const rc = this.cellRect(s.r, s.c); const b = this.el.getBoundingClientRect(); this.contextMenu(b.left + rc.x + rc.w / 2, b.top + rc.y + rc.h / 2); return; }
    if (this.app.handleKey(k)) { e.preventDefault(); return; }
  }

  activeEnd(): { r: number; c: number } {
    const a = this.activeRange, s = this.sel;
    return { r: a.r1 === s.r ? a.r2 : a.r1, c: a.c1 === s.c ? a.c2 : a.c1 };
  }

  moveActive(dr: number, dc: number, jump = false): void {
    const s = this.sel;
    let r = s.r, c = s.c;
    // Leaving a merged cell moves from its edge.
    const m = this.sheet.mergeAt(r, c);
    if (m) { if (dr > 0) r = m.r2; if (dc > 0) c = m.c2; if (dr < 0) r = m.r1; if (dc < 0) c = m.c1; }
    if (jump) ({ r, c } = this.jump(r, c, dr, dc));
    else {
      if (dr) { const n = this.rows.nextVisible(r, dr > 0 ? 1 : -1); if (n >= 0) r = n; }
      if (dc) { const n = this.cols.nextVisible(c, dc > 0 ? 1 : -1); if (n >= 0) c = n; }
    }
    const m2 = this.sheet.mergeAt(r, c);
    if (m2) { r = m2.r1; c = m2.c1; }
    this.select(r, c);
  }

  extendSelection(dr: number, dc: number, jump = false): void {
    const s = this.sel;
    const end = this.activeEnd();
    let { r, c } = end;
    if (jump) ({ r, c } = this.jump(r, c, dr, dc));
    else {
      if (dr) { const n = this.rows.nextVisible(r, dr > 0 ? 1 : -1); if (n >= 0) r = n; }
      if (dc) { const n = this.cols.nextVisible(c, dc > 0 ? 1 : -1); if (n >= 0) c = n; }
    }
    this.select(s.r, s.c, { extendTo: { r, c }, keepActive: true });
  }

  /** Ctrl+Arrow: jump to the edge of the current data region. */
  jump(r: number, c: number, dr: number, dc: number): { r: number; c: number } {
    const has = (rr: number, cc: number) => { const x = this.sheet.get(rr, cc); return !!x && (x.v != null && x.v !== '' || x.f != null); };
    const maxR = MAX_ROWS - 1, maxC = MAX_COLS - 1;
    const used = this.sheet.usedRange(false);
    const step = (rr: number, cc: number) => ({ r: rr + dr, c: cc + dc });
    const inBounds = (p: { r: number; c: number }) => p.r >= 0 && p.c >= 0 && p.r <= maxR && p.c <= maxC;
    let p = { r, c };
    const nxt = step(r, c);
    if (!inBounds(nxt)) return p;
    if (has(r, c) && has(nxt.r, nxt.c)) {
      // run to the last filled cell
      while (inBounds(step(p.r, p.c)) && has(step(p.r, p.c).r, step(p.r, p.c).c)) p = step(p.r, p.c);
      return p;
    }
    // skip blanks to the next filled cell, or to the sheet edge
    p = nxt;
    const limitR = dr > 0 ? (used ? Math.max(used.r2, r) : r) : 0, limitC = dc > 0 ? (used ? Math.max(used.c2, c) : c) : 0;
    while (inBounds(p) && !has(p.r, p.c)) {
      if ((dr > 0 && p.r > limitR) || (dc > 0 && p.c > limitC)) return { r: dr ? maxR : p.r, c: dc ? maxC : p.c };
      if ((dr < 0 && p.r <= 0) || (dc < 0 && p.c <= 0)) return { r: Math.max(0, p.r), c: Math.max(0, p.c) };
      p = step(p.r, p.c);
    }
    return inBounds(p) ? p : { r: clamp(p.r, 0, maxR), c: clamp(p.c, 0, maxC) };
  }

  private page(dr: number, dc: number, shift: boolean): void {
    const vis = this.visible();
    const rowsPer = Math.max(1, vis.r2 - vis.r1);
    const colsPer = Math.max(1, vis.c2 - vis.c1);
    const s = this.sel;
    const r = clamp((shift ? this.activeEnd().r : s.r) + dr * rowsPer, 0, MAX_ROWS - 1);
    const c = clamp((shift ? this.activeEnd().c : s.c) + dc * colsPer, 0, MAX_COLS - 1);
    if (dr) this.scroller.scrollTop += dr * rowsPer * this.sheet.defaultRowH * this.zoom;
    if (dc) this.scroller.scrollLeft += dc * colsPer * this.sheet.defaultColW * this.zoom;
    if (shift) this.select(s.r, s.c, { extendTo: { r, c }, keepActive: true });
    else this.select(r, c);
  }

  /** Enter/Tab: move within a multi-cell selection without losing it, or move the active cell. */
  cycleInSelection(dir: 1 | -1, axis: 'down' | 'right' | 'up' | 'left'): void {
    const s = this.sel;
    const rg = this.activeRange;
    const single = rg.r1 === rg.r2 && rg.c1 === rg.c2 && s.ranges.length === 1;
    const d = axis === 'up' || axis === 'left' ? -dir : dir;
    const vertical = axis === 'down' || axis === 'up';
    if (single) {
      if (!this.app.prefs.enterMoves && axis !== 'right') return;
      if (vertical) this.moveActive(d, 0); else this.moveActive(0, d);
      return;
    }
    let { r, c } = s;
    if (vertical) {
      r += d;
      if (r > rg.r2) { r = rg.r1; c++; if (c > rg.c2) c = rg.c1; }
      if (r < rg.r1) { r = rg.r2; c--; if (c < rg.c1) c = rg.c2; }
    } else {
      c += d;
      if (c > rg.c2) { c = rg.c1; r++; if (r > rg.r2) r = rg.r1; }
      if (c < rg.c1) { c = rg.c2; r--; if (r < rg.r1) r = rg.r2; }
    }
    s.r = r; s.c = c;
    this.scrollIntoView(r, c);
    this.invalidate(); this.app.refresh();
  }

  // ------------------------------------------------------------------------------------------
  // Editing

  startEdit(opts: { mode: 'enter' | 'edit'; text?: string; fromBar?: boolean }): void {
    if (this.app.isReadOnly()) { this.app.exec('review.protectionNotice'); return; }
    const { r, c } = this.sel;
    if (!this.app.exec('cell.canEdit', { r, c })) return;
    this.editor.start(r, c, opts);
  }

  commitEdit(move: 'down' | 'up' | 'right' | 'left' | 'none' | 'enter' = 'enter', opts: { array?: boolean; fill?: boolean } = {}): boolean {
    return this.editor.commit(move, opts);
  }

  cancelEdit(): void { this.editor.cancel(); }
}

function clamp(v: number, a: number, b: number): number { return Math.max(a, Math.min(b, v)); }

function edgeDir(x: number, y: number, rc: { x: number; y: number; w: number; h: number }): 'up' | 'down' | 'left' | 'right' {
  const d = [Math.abs(y - rc.y), Math.abs(y - (rc.y + rc.h)), Math.abs(x - rc.x), Math.abs(x - (rc.x + rc.w))];
  const i = d.indexOf(Math.min(...d));
  return (['up', 'down', 'left', 'right'] as const)[i];
}

export { cellKey };
