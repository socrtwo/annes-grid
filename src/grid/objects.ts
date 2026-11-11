/**
 * Floating objects over the grid: charts, pictures, shapes, text boxes, plus comment pop-ups.
 * Objects are anchored to a cell with a pixel offset and move with scrolling; they can be selected,
 * dragged, resized with handles, deleted, and double-clicked to edit.
 */
import type { Grid } from './grid';
import type { ChartDef, DrawingObject, CellComment } from '../model/types';
import { h, clear } from '../ui/dom';
import { cellKey } from '../model/address';
import { renderChartSvg } from '../charts/render';
import { renderShapeSvg } from '../charts/shapes';
import { safeImageSrc, sanitizeGeneratedHtml } from '../util/sanitize';

type Obj = { kind: 'chart'; def: ChartDef } | { kind: 'drawing'; def: DrawingObject };

/**
 * Extension points for command modules: custom drawing content (ink strokes, form controls) and
 * pointer handling (a form control clicked outside design mode performs its action instead of
 * being selected). A hook returning true has handled the event.
 */
export const objectHooks = {
  render: [] as ((d: DrawingObject, g: Grid) => HTMLElement | null)[],
  pointerdown: [] as ((g: Grid, id: string, e: PointerEvent) => boolean)[],
  dblclick: [] as ((g: Grid, id: string, e: MouseEvent) => boolean)[],
};

export class ObjectLayer {
  private els = new Map<string, { el: HTMLElement; sig: string }>();
  selected: string | null = null;
  private commentEl: HTMLElement;
  private hoverKey = -1;
  showAllComments = false;

  constructor(private g: Grid) {
    this.commentEl = h('div', { class: 'ql-comment-pop', role: 'tooltip' });
    g.overlay.appendChild(this.commentEl);
  }

  private objects(): Obj[] {
    const s = this.g.sheet;
    return [...s.charts.map((def) => ({ kind: 'chart' as const, def })), ...s.drawings.map((def) => ({ kind: 'drawing' as const, def }))];
  }

  selectedObject(): Obj | null {
    if (!this.selected) return null;
    return this.objects().find((o) => o.def.id === this.selected) || null;
  }

  select(id: string | null): void {
    this.selected = id;
    this.position();
    this.g.app.refresh();
  }

  /** Re-creates/positions object elements after scrolling, zooming or edits. */
  position(): void {
    const g = this.g;
    const z = g.zoom;
    const seen = new Set<string>();
    const tab = g.tab;
    for (const o of this.objects()) {
      const d = o.def;
      seen.add(d.id);
      const at = d.at;
      const x = g.colX(at.c) + at.dx * z, y = g.rowY(at.r) + at.dy * z;
      const w = at.w * z, hh = at.h * z;
      const sig = JSON.stringify(d) + ':' + (o.kind === 'chart' || (o.kind === 'drawing' && o.def.control) ? tab.engine.evaluations + ':' + tab.book.version : '') + ':' + z;
      let rec = this.els.get(d.id);
      if (!rec) {
        const el = h('div', { class: `ql-object ql-object-${o.kind}`, 'data-id': d.id, 'data-kind': o.kind === 'drawing' ? o.def.kind : 'chart', tabindex: '-1' });
        this.bind(el, d.id);
        g.overlay.appendChild(el);
        rec = { el, sig: '' };
        this.els.set(d.id, rec);
      }
      if (rec.sig !== sig) {
        rec.sig = sig;
        const body = h('div', { class: 'ql-object-body' });
        try {
          if (o.kind === 'chart') body.innerHTML = renderChartSvg(tab, o.def, at.w, at.h);
          else body.appendChild(objectHooks.render.reduce<HTMLElement | null>((acc, f) => acc || f(o.def, g), null) || drawingContent(o.def));
        } catch (e) { body.textContent = `Could not draw: ${(e as Error).message}`; }
        clear(rec.el);
        rec.el.appendChild(body);
        if (o.kind === 'drawing' && o.def.rotate) body.style.transform = `rotate(${o.def.rotate}deg)`;
        body.style.width = `${at.w}px`; body.style.height = `${at.h}px`;
        body.style.transform = `${o.kind === 'drawing' && o.def.rotate ? `rotate(${o.def.rotate}deg) ` : ''}scale(${z})`;
        body.style.transformOrigin = '0 0';
        if (o.def.id === this.selected) for (const hd of ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w']) rec.el.appendChild(h('div', { class: `ql-handle ql-h-${hd}`, 'data-h': hd }));
      }
      rec.el.classList.toggle('ql-selected', d.id === this.selected);
      if (d.id === this.selected && !rec.el.querySelector('.ql-handle')) for (const hd of ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w']) rec.el.appendChild(h('div', { class: `ql-handle ql-h-${hd}`, 'data-h': hd }));
      if (d.id !== this.selected) rec.el.querySelectorAll('.ql-handle').forEach((x) => x.remove());
      // Hide behind frozen panes: clip to the scrolling area unless anchored in the frozen part.
      const s = g.sheet;
      const minX = at.c < s.freeze.cols ? g.headerW : g.headerW + g.frozenW;
      const minY = at.r < s.freeze.rows ? g.headerH : g.headerH + g.frozenH;
      rec.el.style.left = `${x}px`; rec.el.style.top = `${y}px`;
      rec.el.style.width = `${w}px`; rec.el.style.height = `${hh}px`;
      const clipL = Math.max(0, minX - x), clipT = Math.max(0, minY - y);
      rec.el.style.clipPath = clipL || clipT ? `inset(${clipT}px 0 0 ${clipL}px)` : '';
      rec.el.style.display = x + w < minX || y + hh < minY || x > g.width || y > g.height ? 'none' : '';
    }
    for (const [id, rec] of this.els) if (!seen.has(id)) { rec.el.remove(); this.els.delete(id); }
    if (this.selected && !seen.has(this.selected)) this.selected = null;
    this.positionComments();
  }

  private bind(el: HTMLElement, id: string): void {
    el.addEventListener('pointerdown', (e) => {
      if (objectHooks.pointerdown.some((f) => f(this.g, id, e))) { e.stopPropagation(); return; }
      e.stopPropagation();
      const g = this.g;
      if (g.editing) g.commitEdit('none');
      const handle = (e.target as HTMLElement).dataset.h;
      if (this.selected !== id) this.select(id);
      const o = this.objects().find((x) => x.def.id === id);
      if (!o) return;
      e.preventDefault();
      el.setPointerCapture(e.pointerId);
      const startX = e.clientX, startY = e.clientY;
      const at0 = { ...o.def.at };
      const z = g.zoom;
      const abs0 = { x: g.colX(at0.c) + at0.dx * z, y: g.rowY(at0.r) + at0.dy * z };
      let moved = false;
      const move = (ev: PointerEvent) => {
        const dx = (ev.clientX - startX) / z, dy = (ev.clientY - startY) / z;
        if (Math.abs(dx) + Math.abs(dy) > 2) moved = true;
        if (!moved) return;
        const at = o.def.at;
        if (handle) {
          let { w, h: hh } = at0;
          let nx = 0, ny = 0;
          if (handle.includes('e')) w = Math.max(16, at0.w + dx);
          if (handle.includes('s')) hh = Math.max(16, at0.h + dy);
          if (handle.includes('w')) { w = Math.max(16, at0.w - dx); nx = at0.w - w; }
          if (handle.includes('n')) { hh = Math.max(16, at0.h - dy); ny = at0.h - hh; }
          if (ev.shiftKey || (o.kind === 'drawing' && o.def.lockAspect)) { const ratio = at0.w / at0.h; if (handle.length === 2) hh = w / ratio; }
          const p = this.anchorFor(abs0.x + nx * z, abs0.y + ny * z);
          Object.assign(at, p, { w, h: hh });
        } else {
          Object.assign(at, this.anchorFor(abs0.x + dx * z, abs0.y + dy * z));
        }
        this.position();
      };
      const up = () => {
        el.removeEventListener('pointermove', move);
        el.removeEventListener('pointerup', up);
        if (moved) {
          const after = { ...o.def.at };
          Object.assign(o.def.at, at0);
          g.app.exec('object.setAnchor', { id, at: after });
        }
      };
      el.addEventListener('pointermove', move);
      el.addEventListener('pointerup', up);
    });
    el.addEventListener('dblclick', (e) => {
      if (objectHooks.dblclick.some((f) => f(this.g, id, e))) { e.stopPropagation(); return; }
      e.stopPropagation();
      const o = this.objects().find((x) => x.def.id === id);
      if (!o) return;
      if (o.kind === 'chart') this.g.app.exec('chart.edit', { id });
      else if (o.def.kind === 'textbox' || o.def.kind === 'shape' || o.def.kind === 'wordart') this.g.app.exec('drawing.editText', { id });
      else this.g.app.exec('picture.format', { id });
    });
    el.addEventListener('contextmenu', (e) => { e.preventDefault(); e.stopPropagation(); this.select(id); this.g.app.ui.objectContextMenu(e.clientX, e.clientY, id); });
    el.addEventListener('keydown', (e) => {
      if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); this.g.app.exec('object.delete', { id }); }
      if (e.key === 'Escape') { this.select(null); this.g.focus(); }
    });
  }

  /** Converts an absolute canvas position to a cell anchor with offsets. */
  anchorFor(x: number, y: number): { r: number; c: number; dx: number; dy: number } {
    const g = this.g;
    const c = Math.max(0, g.colAt(Math.max(g.headerW, x)));
    const r = Math.max(0, g.rowAt(Math.max(g.headerH, y)));
    return { r, c, dx: Math.max(0, (x - g.colX(c)) / g.zoom), dy: Math.max(0, (y - g.rowY(r)) / g.zoom) };
  }

  // ------------------------------------------------------------------------------------------
  // Comments and notes

  hoverComment(r: number, c: number): void {
    const k = r < 0 ? -1 : cellKey(r, c);
    if (k === this.hoverKey) return;
    this.hoverKey = k;
    this.positionComments();
  }

  private positionComments(): void {
    const g = this.g;
    const s = g.sheet;
    const el = this.commentEl;
    const k = this.hoverKey;
    const cm = k >= 0 ? s.comments.get(k) : undefined;
    // Data validation input message for the active cell
    const sel = s.sel;
    const dv = s.validations.find((v) => v.inputMessage && v.ranges.some((rg) => sel.r >= rg.r1 && sel.r <= rg.r2 && sel.c >= rg.c1 && sel.c <= rg.c2));
    clear(el);
    if (cm) {
      const r = Math.floor(k / 16384), c = k % 16384;
      const rc = g.cellRect(r, c);
      el.append(...commentContent(cm));
      el.classList.add('ql-show');
      this.fitPopup(el, rc, 'side');
    } else if (dv && !g.editing) {
      const rc = g.cellRect(sel.r, sel.c);
      if (dv.inputTitle) el.appendChild(h('b', null, dv.inputTitle));
      el.appendChild(h('div', null, dv.inputMessage!));
      el.classList.add('ql-show');
      this.fitPopup(el, rc, 'below');
    } else el.classList.remove('ql-show');
    // "Show All Notes" mode
    g.overlay.querySelectorAll('.ql-note-pinned').forEach((x) => x.remove());
    if (this.showAllComments) {
      for (const [ck, c2] of s.comments) {
        if (!c2.note) continue;
        const r = Math.floor(ck / 16384), c = ck % 16384;
        const rc = g.cellRect(r, c);
        if (rc.x > g.width || rc.y > g.height || rc.x + rc.w < 0 || rc.y + rc.h < 0) continue;
        const pin = h('div', { class: 'ql-comment-pop ql-show ql-note-pinned' }, ...commentContent(c2));
        g.overlay.appendChild(pin);
        this.fitPopup(pin, rc, 'side');
      }
    }
  }

  /**
   * Places a note or input message next to its cell ('side': to the right, else to the left) or under it
   * ('below': else above), keeping it inside the grid so that it is not cut off at the window's edge.
   */
  private fitPopup(el: HTMLElement, rc: { x: number; y: number; w: number; h: number }, where: 'side' | 'below'): void {
    const g = this.g;
    const w = el.offsetWidth, hh = el.offsetHeight;
    let x = where === 'side' ? rc.x + rc.w + 8 : rc.x + 12;
    let y = where === 'side' ? rc.y : rc.y + rc.h + 6;
    if (x + w > g.width - 2) x = where === 'side' && rc.x - 8 - w >= 2 ? rc.x - 8 - w : g.width - 2 - w;
    if (y + hh > g.height - 2) y = where === 'below' && rc.y - 6 - hh >= 2 ? rc.y - 6 - hh : g.height - 2 - hh;
    el.style.left = `${Math.max(2, x)}px`;
    el.style.top = `${Math.max(2, y)}px`;
  }
}

function commentContent(cm: CellComment): HTMLElement[] {
  const out: HTMLElement[] = [h('div', { class: 'ql-comment-author' }, cm.author + (cm.note ? ':' : '')), h('div', { class: 'ql-comment-text' }, cm.text)];
  for (const rp of cm.replies || []) out.push(h('div', { class: 'ql-comment-reply' }, h('b', null, rp.author + ': '), rp.text));
  if (cm.resolved) out.push(h('div', { class: 'ql-help' }, 'Resolved'));
  return out;
}

function drawingContent(d: DrawingObject): HTMLElement {
  if (d.kind === 'image' || d.kind === 'signature') {
    const img = h('img', { src: safeImageSrc(d.src), alt: d.alt || '', draggable: 'false' }) as HTMLImageElement;
    img.style.width = '100%'; img.style.height = '100%'; img.style.objectFit = d.lockAspect === false ? 'fill' : 'contain';
    return img;
  }
  if (d.kind === 'icon' && d.svg) { const w = h('div', { class: 'ql-drawing-svg' }); w.innerHTML = sanitizeGeneratedHtml(d.svg); return w; }
  const w = h('div', { class: 'ql-drawing-svg' });
  w.innerHTML = renderShapeSvg(d);
  return w;
}
