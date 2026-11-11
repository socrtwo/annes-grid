/**
 * Floating objects: arrange (order, align, rotate, group), Picture Format and Shape Format commands,
 * plus the shared helpers the Insert, Chart and Table modules use to add and change objects.
 */
import { def, type MenuEntry } from './registry';
import type { App } from '../app';
import type { ChartDef, DrawingObject } from '../model/types';
import type { Sheet } from '../model/workbook';
import { newId } from '../model/workbook';
import { openDialog, toast, colorMenu, THEME_COLORS, STANDARD_COLORS, alertDialog } from '../ui/dialog';
import { h, pickFile, readFileAsDataURL, download, isPhone } from '../ui/dom';
import { renderShapeSvg, shapeLabel, SHAPES, shapePreviewSvg, WORDART_STYLES } from '../charts/shapes';
import { escHtml, isScriptUrl, safeNum } from '../util/sanitize';
import { quoteSheet } from '../model/address';

export type AnyObj = { kind: 'chart'; def: ChartDef; sheet: Sheet } | { kind: 'drawing'; def: DrawingObject; sheet: Sheet };

// --------------------------------------------------------------------------------------------
// Lookup and mutation helpers

/** The object with the given id on any sheet (or the selected object when no id is given). */
export function findObject(app: App, id?: string | null): AnyObj | null {
  const want = id ?? app.tab?.grid?.objects?.selected ?? null;
  if (!want) return null;
  const order = [app.sheet, ...app.book.sheets.filter((s) => s !== app.sheet)];
  for (const sheet of order) {
    const c = sheet.charts.find((x) => x.id === want);
    if (c) return { kind: 'chart', def: c, sheet };
    const d = sheet.drawings.find((x) => x.id === want);
    if (d) return { kind: 'drawing', def: d, sheet };
  }
  return null;
}

const idOf = (arg: any): string | undefined => (arg && typeof arg === 'object' && typeof arg.id === 'string' ? arg.id : typeof arg === 'string' && /^[a-z]{1,4}[0-9a-z]+$/i.test(arg) ? undefined : undefined);

export function selectedDrawing(app: App, arg?: any): DrawingObject | null {
  const o = findObject(app, idOf(arg));
  return o?.kind === 'drawing' ? o.def : null;
}
export function selectedChart(app: App, arg?: any): ChartDef | null {
  const o = findObject(app, idOf(arg));
  return o?.kind === 'chart' ? o.def : null;
}

/** Whether objects on the active sheet may be changed (sheet protection). */
export function canEditObjects(app: App, sheet: Sheet = app.sheet, quiet = false): boolean {
  if (sheet.protection.on && !sheet.protection.allow.includes('objects')) {
    if (!quiet) toast('The sheet is protected. Unprotect it to insert or change objects.');
    return false;
  }
  return true;
}

/** Changes an object as one undoable step. */
export function mutateObject(app: App, o: AnyObj, label: string, fn: (d: any) => void): void {
  if (!canEditObjects(app, o.sheet)) return;
  app.book.transact(label, () => {
    app.book.recordSheet(o.sheet);
    fn(o.def);
  });
  app.tab.grid?.objects.position();
}

export function updateDrawing(app: App, arg: any, label: string, fn: (d: DrawingObject) => void, quietIfNone = false): boolean {
  const o = findObject(app, idOf(arg));
  if (!o || o.kind !== 'drawing') { if (!quietIfNone) toast('Select a picture, shape or text box first.'); return false; }
  mutateObject(app, o, label, fn);
  return true;
}

export function updateChart(app: App, arg: any, label: string, fn: (c: ChartDef) => void): boolean {
  const o = findObject(app, idOf(arg));
  if (!o || o.kind !== 'chart') { toast('Select a chart first.'); return false; }
  mutateObject(app, o, label, fn);
  return true;
}

/** A free anchor near the middle of the visible grid for an object of w × h (px at 100%). */
export function viewAnchor(app: App, w: number, h: number): { r: number; c: number; dx: number; dy: number } {
  const g = app.tab?.grid;
  if (!g || !g.width) { const { r, c } = app.sheet.sel; return { r, c, dx: 0, dy: 0 }; }
  const z = g.zoom;
  const x = Math.max(g.headerW + g.frozenW + 8, (g.width - w * z) / 2);
  const y = Math.max(g.headerH + g.frozenH + 8, (g.height - h * z) / 2);
  return g.objects.anchorFor(x, y);
}

/** Anchor at the active cell (pictures, pasted objects). */
export function cellAnchor(app: App): { r: number; c: number; dx: number; dy: number } {
  const { r, c } = app.sheet.sel;
  return { r, c, dx: 0, dy: 0 };
}

/** Adds a drawing to the active sheet as one undoable step and selects it. */
export function addDrawing(app: App, d: Omit<DrawingObject, 'id'> & { id?: string }, label: string): DrawingObject | null {
  const sheet = app.sheet;
  if (!canEditObjects(app, sheet)) return null;
  const obj = { ...d, id: d.id || newId(d.kind === 'image' ? 'img' : 'shp') } as DrawingObject;
  // Cascade objects inserted at the same spot so each stays visible.
  for (let i = 0; i < 20 && [...sheet.charts, ...sheet.drawings].some((x) => x.at.r === obj.at.r && x.at.c === obj.at.c && Math.abs(x.at.dx - obj.at.dx) < 4 && Math.abs(x.at.dy - obj.at.dy) < 4); i++) obj.at = { ...obj.at, dx: obj.at.dx + 16, dy: obj.at.dy + 16 };
  app.book.transact(label, () => {
    app.book.recordSheet(sheet);
    sheet.drawings.push(obj);
  });
  app.tab.grid?.objects.select(obj.id);
  return obj;
}

export function addChart(app: App, ch: ChartDef, label = 'Insert Chart'): ChartDef | null {
  const sheet = app.sheet;
  if (!canEditObjects(app, sheet)) return null;
  app.book.transact(label, () => {
    app.book.recordSheet(sheet);
    sheet.charts.push(ch);
  });
  app.tab.grid?.objects.select(ch.id);
  return ch;
}

// --------------------------------------------------------------------------------------------
// Gallery dialog (pictures of choices: chart types, shapes, table styles, icons...)

export interface GalleryItem<T> { value: T; label: string; html?: string; group?: string; title?: string }

export function galleryDialog<T>(title: string, items: GalleryItem<T>[], opts: { ok?: string; help?: string; current?: T; size?: 'small' | 'large'; search?: boolean; extra?: HTMLElement } = {}): Promise<T | null> {
  return new Promise((resolve) => {
    let chosen: T | null = opts.current ?? null;
    let picked = false;
    const wrap = h('div', { class: 'ql-objgal-wrap' });
    const grid = h('div', { class: 'ql-objgal' + (opts.size === 'small' ? ' ql-objgal-small' : '') });
    const buttons: HTMLElement[] = [];
    let lastGroup: string | undefined;
    for (const it of items) {
      if (it.group && it.group !== lastGroup) { grid.appendChild(h('div', { class: 'ql-objgal-head' }, it.group)); lastGroup = it.group; }
      const b = h('button', { class: 'ql-objgal-item' + (opts.current === it.value ? ' ql-sel' : ''), type: 'button', title: it.title || it.label, 'data-label': it.label.toLowerCase() });
      if (it.html) { const pv = h('span', { class: 'ql-objgal-pv' }); pv.innerHTML = it.html; b.appendChild(pv); }
      b.appendChild(h('span', { class: 'ql-objgal-label' }, it.label));
      b.addEventListener('click', () => { chosen = it.value; buttons.forEach((x) => x.classList.toggle('ql-sel', x === b)); });
      b.addEventListener('dblclick', () => { chosen = it.value; picked = true; (wrap.closest('.ql-dialog')?.querySelector('.ql-primary') as HTMLElement | null)?.click(); });
      buttons.push(b);
      grid.appendChild(b);
    }
    if (opts.search) {
      const inp = h('input', { type: 'search', class: 'ql-objgal-search', placeholder: 'Search', 'aria-label': 'Search' }) as HTMLInputElement;
      inp.addEventListener('input', () => { const q = inp.value.toLowerCase().trim(); buttons.forEach((b) => { b.style.display = !q || (b.dataset.label || '').includes(q) ? '' : 'none'; }); });
      wrap.appendChild(inp);
    }
    wrap.appendChild(grid);
    if (opts.extra) wrap.appendChild(opts.extra);
    void picked;
    openDialog({ title, content: wrap, wide: true, ok: opts.ok || 'OK', help: opts.help }).then((r) => resolve(r ? chosen : null));
  });
}

// --------------------------------------------------------------------------------------------
// Picture processing (canvas): corrections, color, crop, background removal, borders, effects

export interface PictureFx {
  brightness?: number; contrast?: number; saturation?: number; // percent, 100 = unchanged
  color?: 'none' | 'grayscale' | 'sepia' | 'washout' | 'bw' | 'invert';
  crop?: { t: number; r: number; b: number; l: number }; // percent of each edge
  removeBg?: { tol: number; color?: string };
  shape?: 'none' | 'rounded' | 'ellipse';
  border?: { color: string; width: number } | null;
  shadow?: boolean;
  soft?: number;
  flipH?: boolean; flipV?: boolean;
}

function loadImage(src: string): Promise<HTMLImageElement> {
  return new Promise((res, rej) => {
    const img = new Image();
    img.onload = () => res(img);
    img.onerror = () => rej(new Error('The picture could not be read.'));
    img.src = src;
  });
}

/** Renders the original picture with the effects applied, as a PNG data URL. */
export async function renderPicture(src: string, fx: PictureFx): Promise<string> {
  const img = await loadImage(src);
  const nw = img.naturalWidth || 1, nh = img.naturalHeight || 1;
  const c = fx.crop || { t: 0, r: 0, b: 0, l: 0 };
  const sx = (nw * c.l) / 100, sy = (nh * c.t) / 100;
  const sw = Math.max(1, nw * (1 - (c.l + c.r) / 100)), sh = Math.max(1, nh * (1 - (c.t + c.b) / 100));
  const scale = Math.min(1, 2400 / Math.max(sw, sh));
  const w = Math.round(sw * scale), hh = Math.round(sh * scale);
  const pad = fx.shadow ? Math.round(Math.max(6, Math.min(w, hh) * 0.04)) : 0;
  const bw = fx.border ? Math.max(0, fx.border.width) : 0;
  const cv = document.createElement('canvas');
  cv.width = w + pad * 2; cv.height = hh + pad * 2;
  const ctx = cv.getContext('2d');
  if (!ctx) throw new Error('Pictures cannot be edited in this browser.');
  // 1. the picture with filters
  const layer = document.createElement('canvas');
  layer.width = w; layer.height = hh;
  const lc = layer.getContext('2d')!;
  const filters: string[] = [];
  if (fx.brightness != null && fx.brightness !== 100) filters.push(`brightness(${fx.brightness}%)`);
  if (fx.contrast != null && fx.contrast !== 100) filters.push(`contrast(${fx.contrast}%)`);
  if (fx.saturation != null && fx.saturation !== 100) filters.push(`saturate(${fx.saturation}%)`);
  switch (fx.color) {
    case 'grayscale': filters.push('grayscale(1)'); break;
    case 'sepia': filters.push('sepia(0.85)'); break;
    case 'washout': filters.push('brightness(1.45) contrast(0.55)'); break;
    case 'bw': filters.push('grayscale(1) contrast(4)'); break;
    case 'invert': filters.push('invert(1)'); break;
  }
  if (fx.soft) filters.push(`blur(${fx.soft}px)`);
  lc.filter = filters.join(' ') || 'none';
  lc.save();
  lc.translate(fx.flipH ? w : 0, fx.flipV ? hh : 0);
  lc.scale(fx.flipH ? -1 : 1, fx.flipV ? -1 : 1);
  lc.drawImage(img, sx, sy, sw, sh, 0, 0, w, hh);
  lc.restore();
  // 2. background removal: color-key the corner color (or a chosen color)
  if (fx.removeBg) {
    const data = lc.getImageData(0, 0, w, hh);
    const px = data.data;
    let key: number[];
    if (fx.removeBg.color && /^#[0-9a-f]{6}$/i.test(fx.removeBg.color)) key = [1, 3, 5].map((i) => parseInt(fx.removeBg!.color!.slice(i, i + 2), 16));
    else {
      const corners = [0, (w - 1) * 4, (hh - 1) * w * 4, ((hh - 1) * w + w - 1) * 4];
      key = [0, 1, 2].map((k) => Math.round(corners.reduce((a, i) => a + px[i + k], 0) / 4));
    }
    const tol = Math.max(0, Math.min(255, fx.removeBg.tol ?? 40));
    for (let i = 0; i < px.length; i += 4) {
      const d = Math.max(Math.abs(px[i] - key[0]), Math.abs(px[i + 1] - key[1]), Math.abs(px[i + 2] - key[2]));
      if (d <= tol) px[i + 3] = 0;
      else if (d <= tol * 1.5) px[i + 3] = Math.round(px[i + 3] * ((d - tol) / (tol * 0.5 || 1)));
    }
    lc.putImageData(data, 0, 0);
  }
  // 3. shape clip, shadow, border
  const path = new Path2D();
  if (fx.shape === 'ellipse') path.ellipse(pad + w / 2, pad + hh / 2, w / 2, hh / 2, 0, 0, Math.PI * 2);
  else if (fx.shape === 'rounded') { const r = Math.min(w, hh) * 0.12; (path as any).roundRect ? (path as any).roundRect(pad, pad, w, hh, r) : path.rect(pad, pad, w, hh); }
  else path.rect(pad, pad, w, hh);
  if (fx.shadow) {
    ctx.save();
    ctx.shadowColor = 'rgba(0,0,0,0.4)'; ctx.shadowBlur = pad; ctx.shadowOffsetX = pad / 3; ctx.shadowOffsetY = pad / 3;
    ctx.fillStyle = '#ffffff';
    ctx.fill(path);
    ctx.restore();
    ctx.save(); ctx.globalCompositeOperation = 'destination-out'; ctx.fill(path); ctx.restore();
    // Keep only the shadow outside the picture, then draw the picture.
  }
  ctx.save();
  ctx.clip(path);
  ctx.drawImage(layer, pad, pad);
  ctx.restore();
  if (bw > 0) { ctx.save(); ctx.lineWidth = bw * 2; ctx.strokeStyle = fx.border!.color; ctx.clip(path); ctx.stroke(path); ctx.restore(); }
  return cv.toDataURL('image/png');
}

const isDefaultFx = (fx: PictureFx) => !Object.entries(fx).some(([k, v]) => v != null && v !== false && !(k === 'color' && v === 'none') && !(k === 'shape' && v === 'none') && !((k === 'brightness' || k === 'contrast' || k === 'saturation') && v === 100));

/** Applies picture effects (merged into the drawing's fx) and re-renders its data URL. */
export async function applyPictureFx(app: App, arg: any, patch: Partial<PictureFx>, label: string): Promise<void> {
  const o = findObject(app, idOf(arg));
  if (!o || o.kind !== 'drawing' || (o.def.kind !== 'image' && o.def.kind !== 'signature')) { toast('Select a picture first.'); return; }
  if (!canEditObjects(app, o.sheet)) return;
  const d = o.def;
  const orig = d.orig || d.src || '';
  if (!orig || isScriptUrl(orig)) { toast('This picture cannot be edited.'); return; }
  const fx: PictureFx = { ...(d.fx || {}), ...patch };
  for (const k of Object.keys(fx) as (keyof PictureFx)[]) if (fx[k] == null) delete fx[k];
  let src: string;
  try { src = isDefaultFx(fx) ? orig : await renderPicture(orig, fx); } catch (e) { toast((e as Error).message); return; }
  const oldCrop = d.fx?.crop, newCrop = fx.crop;
  mutateObject(app, o, label, (x: DrawingObject) => {
    x.orig = orig;
    x.fx = fx;
    x.src = src;
    if (JSON.stringify(oldCrop || null) !== JSON.stringify(newCrop || null)) {
      // Keep the visible scale when the crop changes.
      const before = oldCrop || { t: 0, r: 0, b: 0, l: 0 }, after = newCrop || { t: 0, r: 0, b: 0, l: 0 };
      const fw = (100 - after.l - after.r) / Math.max(1, 100 - before.l - before.r), fh = (100 - after.t - after.b) / Math.max(1, 100 - before.t - before.b);
      x.at = { ...x.at, w: Math.max(8, x.at.w * fw), h: Math.max(8, x.at.h * fh) };
    }
  });
}

// --------------------------------------------------------------------------------------------
// Size & properties dialog

async function sizeDialog(app: App, arg: any): Promise<void> {
  const o = findObject(app, idOf(arg));
  if (!o) { toast('Select a picture, shape or chart first.'); return; }
  const at = o.def.at;
  const isDrawing = o.kind === 'drawing';
  const d = o.def as DrawingObject;
  const ratio = at.h / (at.w || 1);
  let natural: { w: number; h: number } | null = null;
  if (isDrawing && d.kind === 'image' && (d.orig || d.src)) { try { const img = await loadImage(d.orig || d.src!); natural = { w: img.naturalWidth, h: img.naturalHeight }; } catch { /* ignore */ } }
  const r = await openDialog({
    title: o.kind === 'chart' ? 'Format Chart Area: Size' : isDrawing && d.kind === 'image' ? 'Format Picture: Size & Properties' : 'Format Shape: Size & Properties',
    fields: [
      { key: 'h', label: 'Height (px)', type: 'number', value: Math.round(at.h), min: 4, half: true },
      { key: 'w', label: 'Width (px)', type: 'number', value: Math.round(at.w), min: 4, half: true },
      ...(isDrawing ? [{ key: 'rot', label: 'Rotation (°)', type: 'number' as const, value: d.rotate || 0, min: -360, max: 360, half: true }] : []),
      { key: 'pct', label: 'Scale (%)', type: 'number', value: 100, min: 1, max: 1000, half: true },
      { key: 'lock', label: 'Lock aspect ratio', type: 'checkbox', value: isDrawing ? d.lockAspect ?? d.kind === 'image' : false },
      ...(natural ? [{ key: 'orig', label: `Reset to original size (${natural.w} × ${natural.h})`, type: 'checkbox' as const, value: false }] : []),
      { key: 'alt', label: 'Alternative text (describes the object for screen readers)', type: 'textarea', rows: 3, value: isDrawing ? d.alt || '' : (o.def as ChartDef).title || '' , visibleIf: () => isDrawing },
    ],
    onChange: (v, ctl, key) => {
      if (key === 'w' && v.lock && v.w > 0) ctl.set('h', Math.round(v.w * ratio));
      if (key === 'h' && v.lock && v.h > 0) ctl.set('w', Math.round(v.h / ratio));
      if (key === 'pct' && v.pct > 0) { ctl.set('w', Math.round((at.w * v.pct) / 100)); ctl.set('h', Math.round((at.h * v.pct) / 100)); }
    },
    validate: (v) => (!(v.w >= 4) || !(v.h >= 4) ? 'Enter a width and height of at least 4 pixels.' : null),
  });
  if (!r) return;
  mutateObject(app, o, 'Size and Properties', (x: any) => {
    let w = r.w, hh = r.h;
    if (r.orig && natural) { const k = Math.min(1, 640 / Math.max(natural.w, natural.h)); w = natural.w * k; hh = natural.h * k; }
    x.at = { ...x.at, w, h: hh };
    if (isDrawing) { x.rotate = ((Number(r.rot) || 0) % 360 + 360) % 360 || undefined; x.lockAspect = !!r.lock; x.alt = String(r.alt || '') || undefined; }
  });
}

// --------------------------------------------------------------------------------------------
// Edit text / content of shapes

const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 48, 60, 72];

async function editText(app: App, arg: any): Promise<void> {
  const o = findObject(app, idOf(arg));
  if (!o || o.kind !== 'drawing') { toast('Select a shape, text box or Decorative Text first.'); return; }
  const d = o.def;
  if (d.shape === 'embed') {
    const choice = await openDialog({ title: d.text || 'Embedded File', content: h('div', { class: 'ql-msg' }, h('p', null, `"${d.text || 'file'}" is embedded in this workbook.`), h('p', { class: 'ql-help' }, 'Save a copy to open it with another app.')), ok: 'Save a Copy…', cancel: 'Close' });
    if (choice && d.src && d.src.startsWith('data:') && !isScriptUrl(d.src)) {
      const blob = await (await fetch(d.src)).blob();
      const { saveAs } = await import('../platform/native');
      await saveAs(blob, d.text || 'embedded-file').catch(() => download(blob, d.text || 'embedded-file'));
    }
    return;
  }
  if (d.shape === 'slicer') { app.exec('insert.slicerPick', { id: d.id }); return; }
  if (d.shape?.startsWith('diagram:')) { app.exec('insert.diagram', { id: d.id }); return; }
  if (d.kind === 'icon') { app.exec('insert.icons', { id: d.id }); return; }
  if (d.kind === 'textbox' && d.shape === 'equation') { app.exec('insert.equation', { id: d.id }); return; }
  if (d.kind === 'image' || d.kind === 'signature') { app.exec('picture.format', { id: d.id }); return; }
  const font = d.font || {};
  const r = await openDialog({
    title: d.kind === 'wordart' ? 'Edit Decorative Text' : d.kind === 'textbox' ? 'Edit Text Box' : `Edit Text: ${shapeLabel(d.shape)}`,
    preview: true, previewStyle: 'height:150px;display:flex;align-items:center;justify-content:center;background:#fff',
    fields: [
      { key: 'text', label: 'Text', type: 'textarea', rows: 4, value: d.text || '', autofocus: true },
      { key: 'font', label: 'Font', type: 'combo', value: font.font || '', options: ['Carlito', 'Calibri', 'Arial', 'Georgia', 'Times New Roman', 'Verdana', 'Courier New', 'Trebuchet MS', 'Impact', 'Comic Sans MS'], half: true, placeholder: 'Default' },
      { key: 'size', label: 'Size', type: 'select', options: ['', ...FONT_SIZES.map(String)].map((s) => ({ value: s, label: s || 'Auto' })), value: font.size ? String(font.size) : '', half: true },
      { key: 'bold', label: 'Bold', type: 'checkbox', value: d.kind === 'wordart' ? font.bold !== false : !!font.bold, half: true },
      { key: 'italic', label: 'Italic', type: 'checkbox', value: !!font.italic, half: true },
      ...(font.color && !/^#[0-9a-f]{6}$/i.test(font.color) ? [{ key: 'keepFill', label: 'Keep the current text fill (gradient)', type: 'checkbox' as const, value: true }] : []),
      { key: 'color', label: 'Text color', type: 'color', value: /^#[0-9a-f]{6}$/i.test(font.color || '') ? font.color : /#[0-9a-f]{6}/i.exec(font.color || '')?.[0] || (d.kind === 'textbox' ? '#000000' : d.kind === 'wordart' ? '#2f6fb0' : '#ffffff'), half: true, visibleIf: (v) => !v.keepFill },
      ...(d.kind !== 'wordart' ? [
        { key: 'hAlign', label: 'Horizontal alignment', type: 'select' as const, options: [{ value: '', label: 'Default' }, { value: 'left', label: 'Left' }, { value: 'center', label: 'Center' }, { value: 'right', label: 'Right' }], value: font.hAlign || '', half: true },
        { key: 'vAlign', label: 'Vertical alignment', type: 'select' as const, options: [{ value: '', label: 'Default' }, { value: 'top', label: 'Top' }, { value: 'middle', label: 'Middle' }, { value: 'bottom', label: 'Bottom' }], value: font.vAlign || '', half: true },
      ] : []),
    ],
    onChange: (v, ctl) => {
      const tmp: DrawingObject = { ...d, at: { ...d.at }, text: v.text, font: fontFrom(v, d) };
      const aw = safeNum(d.at.w, 200, 1, 20000), ah = safeNum(d.at.h, 120, 1, 20000);
      const k = Math.min(1, 320 / aw, 140 / ah);
      ctl.setPreview(`<div style="width:${aw * k}px;height:${ah * k}px"><div style="transform:scale(${k});transform-origin:0 0;width:${aw}px;height:${ah}px">${renderShapeSvg(tmp)}</div></div>`);
    },
  });
  if (!r) return;
  mutateObject(app, o, 'Edit Text', (x: DrawingObject) => { x.text = String(r.text ?? ''); x.font = fontFrom(r, d); });
}

function fontFrom(v: Record<string, any>, d: DrawingObject): DrawingObject['font'] {
  const f: NonNullable<DrawingObject['font']> = { ...(d.font || {}) };
  const set = (k: keyof typeof f, val: any) => { if (val === '' || val == null || val === false) delete f[k]; else (f as any)[k] = val; };
  set('font', String(v.font || '').trim());
  set('size', v.size ? Number(v.size) : '');
  if (d.kind === 'wordart') f.bold = !!v.bold; else set('bold', !!v.bold);
  set('italic', !!v.italic);
  if (!v.keepFill) set('color', v.color);
  if ('hAlign' in v) set('hAlign', v.hAlign);
  if ('vAlign' in v) set('vAlign', v.vAlign);
  return Object.keys(f).length ? f : undefined;
}

// --------------------------------------------------------------------------------------------
// Menus for fills, outlines and effects

const PALETTE_ROW = [...THEME_COLORS[0], ...STANDARD_COLORS];
function colorEntries(id: string, extra: MenuEntry[] = []): MenuEntry[] {
  return [
    ...extra,
    { heading: 'Theme Colors' },
    ...THEME_COLORS[0].map((c) => ({ id, arg: c, label: c, swatch: c })),
    { heading: 'Standard Colors' },
    ...STANDARD_COLORS.map((c) => ({ id, arg: c, label: c, swatch: c })),
    { separator: true },
    { id, arg: '__more', label: 'More Colors…', icon: 'palette' },
  ];
}
void PALETTE_ROW;

async function pickMore(app: App, id: string): Promise<string | null> {
  const c = await colorMenu(app.ui.anchorFor(id), { title: 'Colors' });
  return c || null;
}

const GRADIENTS: [string, string][] = [
  ['Light Linear', 'gradient:#ffffff:{c}:90'], ['Dark Linear', 'gradient:{c}:#1f3b57:90'], ['Linear Diagonal', 'gradient:{c}:#ffffff:45'], ['Linear Left', 'gradient:{c}:#ffffff:0'],
];

function currentDrawingColor(app: App): string {
  const d = selectedDrawing(app);
  const f = d?.fill;
  return f && /^#[0-9a-f]{6}$/i.test(f) && !/^#(f[0-9a-f]){3}$/i.test(f) ? f : app.book.theme.colors[4] || '#2f6fb0';
}

// --------------------------------------------------------------------------------------------
// Arrange

function reorder(app: App, arg: any, how: 'forward' | 'backward' | 'front' | 'back'): void {
  const o = findObject(app, idOf(arg));
  if (!o) { toast('Select an object first.'); return; }
  mutateObject(app, o, how === 'front' ? 'Bring to Front' : how === 'back' ? 'Send to Back' : how === 'forward' ? 'Bring Forward' : 'Send Backward', () => {
    const arr: any[] = o.kind === 'chart' ? o.sheet.charts : o.sheet.drawings;
    const i = arr.findIndex((x) => x.id === o.def.id);
    if (i < 0) return;
    const [x] = arr.splice(i, 1);
    const j = how === 'front' ? arr.length : how === 'back' ? 0 : how === 'forward' ? Math.min(arr.length, i + 1) : Math.max(0, i - 1);
    arr.splice(j, 0, x);
  });
  const el = app.tab.grid?.overlay?.querySelector?.(`[data-id="${CSS.escape(o.def.id)}"]`) as HTMLElement | null;
  // Re-append elements in the new order so the DOM stacking follows the model.
  if (el && app.tab.grid) {
    const ids = [...o.sheet.charts, ...o.sheet.drawings].map((x) => x.id);
    for (const id of ids) { const e = app.tab.grid.overlay.querySelector(`[data-id="${CSS.escape(id)}"]`); if (e) app.tab.grid.overlay.appendChild(e); }
  }
}

function colEdges(sheet: Sheet, upTo: number): number[] { const out = [0]; let x = 0; for (let c = 0; c <= upTo; c++) { x += sheet.colWidth(c); out.push(x); } return out; }
function rowEdges(sheet: Sheet, upTo: number): number[] { const out = [0]; let y = 0; for (let r = 0; r <= upTo; r++) { y += sheet.rowHeight(r); out.push(y); } return out; }

/** Absolute position (px at 100%) of an anchor and back. */
export function anchorToXY(sheet: Sheet, at: { r: number; c: number; dx: number; dy: number }): { x: number; y: number } {
  let x = 0, y = 0;
  for (let c = 0; c < at.c; c++) x += sheet.colWidth(c);
  for (let r = 0; r < at.r; r++) y += sheet.rowHeight(r);
  return { x: x + at.dx, y: y + at.dy };
}
export function xyToAnchor(sheet: Sheet, x: number, y: number): { r: number; c: number; dx: number; dy: number } {
  x = Math.max(0, x); y = Math.max(0, y);
  let c = 0, cx = 0;
  while (c < 16383 && cx + sheet.colWidth(c) <= x) { cx += sheet.colWidth(c); c++; }
  let r = 0, ry = 0;
  while (r < 1048575 && ry + sheet.rowHeight(r) <= y) { ry += sheet.rowHeight(r); r++; }
  return { r, c, dx: x - cx, dy: y - ry };
}

function align(app: App, how: string): void {
  const o = findObject(app);
  if (!o) { toast('Select an object first.'); return; }
  const s = o.sheet;
  if (how === 'distributeH' || how === 'distributeV') {
    const all = [...s.charts, ...s.drawings];
    if (all.length < 3) { toast('Distributing needs at least three objects on the sheet.'); return; }
    mutateObject(app, o, 'Distribute', () => {
      const pos = all.map((d) => ({ d, ...anchorToXY(s, d.at) }));
      const key = how === 'distributeH' ? 'x' : 'y';
      const size = (d: any) => (how === 'distributeH' ? d.at.w : d.at.h);
      pos.sort((a, b) => a[key] - b[key]);
      const first = pos[0], last = pos[pos.length - 1];
      const span = last[key] + size(last.d) - first[key];
      const total = pos.reduce((a, p) => a + size(p.d), 0);
      const gap = (span - total) / (pos.length - 1);
      let cur = first[key];
      for (const p of pos) {
        const nx = how === 'distributeH' ? cur : p.x, ny = how === 'distributeH' ? p.y : cur;
        Object.assign(p.d.at, xyToAnchor(s, nx, ny));
        cur += size(p.d) + gap;
      }
    });
    return;
  }
  if (how === 'snap') { app.prefs && ((app as any).snapToGrid = !(app as any).snapToGrid); toast((app as any).snapToGrid ? 'Snap to Grid is on for moved objects.' : 'Snap to Grid is off.'); return; }
  mutateObject(app, o, 'Align', (d: any) => {
    const { x, y } = anchorToXY(s, d.at);
    const w = d.at.w, hh = d.at.h;
    const ce = colEdges(s, d.at.c + 60), re = rowEdges(s, d.at.r + 200);
    const nearest = (edges: number[], v: number) => edges.reduce((a, b) => (Math.abs(b - v) < Math.abs(a - v) ? b : a), edges[0]);
    let nx = x, ny = y;
    switch (how) {
      case 'left': nx = nearest(ce, x); break;
      case 'right': nx = nearest(ce, x + w) - w; break;
      case 'center': { const c0 = nearest(ce, x), c1 = Math.max(nearest(ce, x + w), c0 + 1); nx = (c0 + c1) / 2 - w / 2; break; }
      case 'top': ny = nearest(re, y); break;
      case 'bottom': ny = nearest(re, y + hh) - hh; break;
      case 'middle': { const r0 = nearest(re, y), r1 = Math.max(nearest(re, y + hh), r0 + 1); ny = (r0 + r1) / 2 - hh / 2; break; }
    }
    Object.assign(d.at, xyToAnchor(s, nx, ny));
  });
}

function rotate(app: App, how: string | number): void {
  const o = findObject(app);
  if (!o) { toast('Select an object first.'); return; }
  if (o.kind === 'chart') { toast('Charts cannot be rotated.'); return; }
  const d = o.def;
  if (how === 'more') { void sizeDialog(app, { id: d.id }); return; }
  if (how === 'flipH' || how === 'flipV') {
    if (d.kind === 'image') { void applyPictureFx(app, { id: d.id }, { [how]: !(d.fx as any)?.[how] }, how === 'flipH' ? 'Flip Horizontal' : 'Flip Vertical'); return; }
    mutateObject(app, o, how === 'flipH' ? 'Flip Horizontal' : 'Flip Vertical', (x: DrawingObject) => { x.fx = { ...(x.fx || {}), [how]: !x.fx?.[how] }; });
    return;
  }
  const delta = how === 'left' ? -90 : how === 'right' ? 90 : Number(how) || 90;
  mutateObject(app, o, 'Rotate', (x: DrawingObject) => { const r = (((x.rotate || 0) + delta) % 360 + 360) % 360; x.rotate = r || undefined; });
}

// --------------------------------------------------------------------------------------------
// Commands

const hasObj = (app: App) => !!app.tab?.grid?.objects?.selectedObject();
const hasPicture = (app: App) => { const o = app.tab?.grid?.objects?.selectedObject(); return o?.kind === 'drawing' && (o.def.kind === 'image' || o.def.kind === 'signature'); };
const hasShape = (app: App) => { const o = app.tab?.grid?.objects?.selectedObject(); return o?.kind === 'drawing' && o.def.kind !== 'image'; };

def([
  {
    id: 'object.setAnchor', label: 'Move or Resize Object', icon: 'open_with', tab: 'Page Layout', group: 'Arrange', internal: true, edits: true,
    run: (app, arg: { id: string; at: Partial<ChartDef['at']> }) => {
      const o = findObject(app, arg?.id);
      if (!o || !arg?.at) return false;
      const at = { ...arg.at } as ChartDef['at'];
      if ((app as any).snapToGrid && at.dx != null && at.dy != null) {
        // Snap the top-left corner to the nearest cell corner.
        if (at.dx > o.sheet.colWidth(at.c ?? o.def.at.c) / 2) { at.c = (at.c ?? o.def.at.c) + 1; } at.dx = 0;
        if (at.dy > o.sheet.rowHeight(at.r ?? o.def.at.r) / 2) { at.r = (at.r ?? o.def.at.r) + 1; } at.dy = 0;
      }
      for (const k of ['w', 'h'] as const) if (at[k] != null) at[k] = Math.max(4, Math.min(20000, Number(at[k]) || 4));
      mutateObject(app, o, 'Move Object', (d) => { d.at = { ...d.at, ...at }; });
      return true;
    },
  },
  {
    id: 'object.delete', label: 'Delete Object', icon: 'delete', tab: 'Page Layout', group: 'Arrange', edits: true, keywords: 'remove chart picture shape',
    run: (app, arg) => {
      const o = findObject(app, idOf(arg));
      if (!o) { toast('Select an object first.'); return false; }
      if (!canEditObjects(app, o.sheet)) return false;
      app.book.transact('Delete Object', () => {
        app.book.recordSheet(o.sheet);
        if (o.kind === 'chart') o.sheet.charts = o.sheet.charts.filter((x) => x.id !== o.def.id);
        else o.sheet.drawings = o.sheet.drawings.filter((x) => x.id !== o.def.id);
      });
      app.tab.grid?.objects.select(null);
      app.focus();
      return true;
    },
  },
  {
    id: 'object.paste', label: 'Paste Object', icon: 'content_paste', tab: 'Home', group: 'Clipboard', internal: true, edits: true,
    run: (app, arg: { kind: 'chart' | 'drawing'; def: any }) => {
      if (!arg?.def) return false;
      const copy = JSON.parse(JSON.stringify(arg.def));
      const src = arg.kind === 'chart' ? findObject(app, copy.id) : null;
      copy.id = newId(arg.kind === 'chart' ? 'ch' : 'shp');
      const sel = app.sheet.sel;
      const same = [...app.sheet.charts, ...app.sheet.drawings].some((x) => x.at.r === sel.r && x.at.c === sel.c);
      copy.at = { ...copy.at, r: sel.r, c: sel.c, dx: same ? 12 : 0, dy: same ? 12 : 0 };
      if (arg.kind === 'chart') {
        // Charts copied to another sheet keep pointing at their source data.
        if (src && src.sheet !== app.sheet) for (const s of copy.series || []) for (const k of ['values', 'categories', 'nameRef', 'sizes']) if (s[k] && !/!/.test(s[k]) && !/^[{"]/.test(s[k])) s[k] = `${quoteSheet(src.sheet.name)}!${s[k]}`;
        return addChart(app, copy, 'Paste Chart');
      }
      if (copy.src && isScriptUrl(copy.src)) delete copy.src;
      return addDrawing(app, copy, 'Paste Object');
    },
  },
  { id: 'object.bringForward', label: 'Bring Forward', icon: 'flip_to_front', tab: 'Page Layout', group: 'Arrange', kind: 'split', edits: true, isEnabled: hasObj, run: (app, arg) => reorder(app, arg, 'forward'), menu: () => [{ id: 'object.bringForward', label: 'Bring Forward', icon: 'flip_to_front' }, { id: 'object.bringToFront', label: 'Bring to Front', icon: 'vertical_align_top' }] },
  { id: 'object.sendBackward', label: 'Send Backward', icon: 'flip_to_back', tab: 'Page Layout', group: 'Arrange', kind: 'split', edits: true, isEnabled: hasObj, run: (app, arg) => reorder(app, arg, 'backward'), menu: () => [{ id: 'object.sendBackward', label: 'Send Backward', icon: 'flip_to_back' }, { id: 'object.sendToBack', label: 'Send to Back', icon: 'vertical_align_bottom' }] },
  { id: 'object.bringToFront', label: 'Bring to Front', icon: 'vertical_align_top', tab: 'Page Layout', group: 'Arrange', edits: true, isEnabled: hasObj, run: (app, arg) => reorder(app, arg, 'front') },
  { id: 'object.sendToBack', label: 'Send to Back', icon: 'vertical_align_bottom', tab: 'Page Layout', group: 'Arrange', edits: true, isEnabled: hasObj, run: (app, arg) => reorder(app, arg, 'back') },
  {
    id: 'object.align', label: 'Align', icon: 'align_horizontal_left', tab: 'Page Layout', group: 'Arrange', kind: 'menu', edits: true,
    run: (app, arg) => align(app, typeof arg === 'string' ? arg : 'left'),
    menu: (app) => [
      { id: 'object.align', arg: 'left', label: 'Align Left (to cell edge)', icon: 'align_horizontal_left' }, { id: 'object.align', arg: 'center', label: 'Align Center', icon: 'align_horizontal_center' }, { id: 'object.align', arg: 'right', label: 'Align Right', icon: 'align_horizontal_right' },
      { separator: true },
      { id: 'object.align', arg: 'top', label: 'Align Top', icon: 'align_vertical_top' }, { id: 'object.align', arg: 'middle', label: 'Align Middle', icon: 'align_vertical_center' }, { id: 'object.align', arg: 'bottom', label: 'Align Bottom', icon: 'align_vertical_bottom' },
      { separator: true },
      { id: 'object.align', arg: 'distributeH', label: 'Distribute Horizontally', icon: 'horizontal_distribute' }, { id: 'object.align', arg: 'distributeV', label: 'Distribute Vertically', icon: 'vertical_distribute' },
      { separator: true },
      { id: 'object.align', arg: 'snap', label: 'Snap to Grid', icon: 'grid_on', checked: !!(app as any).snapToGrid },
    ],
  },
  {
    id: 'object.group', label: 'Group', icon: 'group_work', tab: 'Page Layout', group: 'Arrange', kind: 'menu', edits: true,
    run: () => toast('Grouping needs several objects selected at once. Multiple object selection isn\'t available yet; use Align and Distribute to arrange objects together.', { ms: 5000 }),
    menu: () => [{ id: 'object.group', label: 'Group', icon: 'group_work' }, { id: 'object.group', label: 'Regroup', icon: 'workspaces' }, { id: 'object.group', label: 'Ungroup', icon: 'join_inner' }],
  },
  {
    id: 'object.rotate', label: 'Rotate', icon: 'rotate_right', tab: 'Page Layout', group: 'Arrange', kind: 'menu', edits: true,
    run: (app, arg) => rotate(app, arg ?? 'right'),
    menu: () => [
      { id: 'object.rotate', arg: 'right', label: 'Rotate Right 90°', icon: 'rotate_right' }, { id: 'object.rotate', arg: 'left', label: 'Rotate Left 90°', icon: 'rotate_left' },
      { id: 'object.rotate', arg: 'flipV', label: 'Flip Vertical', icon: 'swap_vert' }, { id: 'object.rotate', arg: 'flipH', label: 'Flip Horizontal', icon: 'swap_horiz' },
      { separator: true }, { id: 'object.rotate', arg: 'more', label: 'More Rotation Options…', icon: 'tune' },
    ],
  },
  { id: 'object.flipH', label: 'Flip Horizontal', icon: 'swap_horiz', tab: 'Page Layout', group: 'Arrange', edits: true, isEnabled: hasObj, run: (app) => rotate(app, 'flipH') },
  { id: 'object.flipV', label: 'Flip Vertical', icon: 'swap_vert', tab: 'Page Layout', group: 'Arrange', edits: true, isEnabled: hasObj, run: (app) => rotate(app, 'flipV') },
  { id: 'object.nudge', label: 'Nudge Object', icon: 'open_with', tab: 'Page Layout', group: 'Arrange', internal: true, edits: true, run: (app, arg: { dx: number; dy: number }) => { const o = findObject(app); if (!o) return; mutateObject(app, o, 'Nudge', (d) => { const p = anchorToXY(o.sheet, d.at); Object.assign(d.at, xyToAnchor(o.sheet, p.x + (arg?.dx || 0), p.y + (arg?.dy || 0))); }); } },

  // ---- Picture Format
  {
    id: 'picture.removeBg', label: 'Remove Background', icon: 'auto_fix_high', tab: 'Picture Format', group: 'Adjust', edits: true, isEnabled: hasPicture,
    run: async (app, arg) => {
      const d = selectedDrawing(app, arg);
      if (!d) { toast('Select a picture first.'); return; }
      const cur = d.fx?.removeBg;
      const r = await openDialog({
        title: 'Remove Background', help: 'Makes the background color transparent. By default the color in the picture\'s corners is removed.',
        fields: [
          { key: 'mode', label: 'Color to remove', type: 'radio', options: [{ value: 'corner', label: 'Corner color (automatic)' }, { value: 'pick', label: 'This color:' }], value: cur?.color ? 'pick' : 'corner' },
          { key: 'color', label: 'Color', type: 'color', value: cur?.color || '#ffffff', visibleIf: (v) => v.mode === 'pick' },
          { key: 'tol', label: 'Tolerance', type: 'range', min: 0, max: 160, value: cur?.tol ?? 40 },
        ],
        buttons: cur ? [{ label: 'Keep Background', onClick: (_v, ctl) => ctl.close({ __clear: true }) }] : [],
      });
      if (!r) return;
      await applyPictureFx(app, arg, { removeBg: r.__clear ? undefined : { tol: Number(r.tol), color: r.mode === 'pick' ? r.color : undefined } }, 'Remove Background');
    },
  },
  {
    id: 'picture.corrections', label: 'Corrections', icon: 'tune', tab: 'Picture Format', group: 'Adjust', kind: 'split', edits: true, isEnabled: hasPicture,
    run: async (app, arg) => {
      if (arg && typeof arg === 'object' && 'brightness' in arg) { await applyPictureFx(app, null, arg, 'Corrections'); return; }
      const d = selectedDrawing(app, arg);
      if (!d) { toast('Select a picture first.'); return; }
      const fx = (d.fx || {}) as PictureFx;
      const src = escHtml(d.orig || d.src || '');
      const r = await openDialog({
        title: 'Picture Corrections', preview: true, previewStyle: 'display:flex;justify-content:center;background:#fff',
        fields: [
          { key: 'brightness', label: 'Brightness', type: 'range', min: 20, max: 200, value: fx.brightness ?? 100, suffix: '%' },
          { key: 'contrast', label: 'Contrast', type: 'range', min: 20, max: 300, value: fx.contrast ?? 100, suffix: '%' },
          { key: 'saturation', label: 'Color saturation', type: 'range', min: 0, max: 300, value: fx.saturation ?? 100, suffix: '%' },
        ],
        onChange: (v, ctl) => ctl.setPreview(`<img alt="" src="${src}" style="max-width:100%;max-height:160px;filter:brightness(${v.brightness}%) contrast(${v.contrast}%) saturate(${v.saturation}%)">`),
        buttons: [{ label: 'Reset', onClick: (_v, ctl) => { ctl.set('brightness', 100); ctl.set('contrast', 100); ctl.set('saturation', 100); } }],
      });
      if (r) await applyPictureFx(app, arg, { brightness: r.brightness, contrast: r.contrast, saturation: r.saturation }, 'Corrections');
    },
    menu: () => {
      const out: MenuEntry[] = [{ heading: 'Brightness / Contrast' }];
      for (const b of [80, 100, 120]) for (const c of [80, 100, 120]) out.push({ id: 'picture.corrections', arg: { brightness: b, contrast: c }, label: `Brightness ${b > 100 ? '+' : ''}${b - 100}%, Contrast ${c > 100 ? '+' : ''}${c - 100}%`, icon: 'brightness_6' });
      out.push({ separator: true }, { id: 'picture.corrections', label: 'Picture Corrections Options…', icon: 'tune' });
      return out;
    },
  },
  {
    id: 'picture.color', label: 'Color', icon: 'palette', tab: 'Picture Format', group: 'Adjust', kind: 'menu', edits: true, isEnabled: hasPicture,
    run: (app, arg) => applyPictureFx(app, null, { color: (typeof arg === 'string' ? arg : 'grayscale') as PictureFx['color'] }, 'Picture Color'),
    menu: () => [['none', 'No Recolor'], ['grayscale', 'Grayscale'], ['sepia', 'Sepia'], ['washout', 'Washout'], ['bw', 'Black and White'], ['invert', 'Negative']].map(([a, l]) => ({ id: 'picture.color', arg: a, label: l, icon: 'palette' })),
  },
  {
    id: 'picture.effects', label: 'Picture Effects', icon: 'blur_on', tab: 'Picture Format', group: 'Adjust', kind: 'menu', edits: true, isEnabled: hasPicture,
    run: (app, arg) => {
      const presets: Record<string, Partial<PictureFx>> = {
        none: { shadow: undefined, shape: undefined, soft: undefined, border: undefined }, shadow: { shadow: true }, noShadow: { shadow: undefined },
        rounded: { shape: 'rounded' }, oval: { shape: 'ellipse' }, rect: { shape: undefined }, soft: { soft: 2 }, sharp: { soft: undefined },
        frame: { border: { color: '#ffffff', width: 8 }, shadow: true }, polaroid: { border: { color: '#ffffff', width: 12 }, shadow: true, shape: undefined },
      };
      return applyPictureFx(app, null, presets[String(arg)] || presets.shadow, 'Picture Effects');
    },
    menu: () => [
      { id: 'picture.effects', arg: 'shadow', label: 'Drop Shadow', icon: 'shadow' }, { id: 'picture.effects', arg: 'noShadow', label: 'No Shadow', icon: 'block' },
      { separator: true }, { id: 'picture.effects', arg: 'rounded', label: 'Rounded Corners', icon: 'rounded_corner' }, { id: 'picture.effects', arg: 'oval', label: 'Oval', icon: 'circle' }, { id: 'picture.effects', arg: 'rect', label: 'Rectangle', icon: 'crop_square' },
      { separator: true }, { id: 'picture.effects', arg: 'soft', label: 'Soft Edges', icon: 'blur_on' }, { id: 'picture.effects', arg: 'sharp', label: 'No Soft Edges', icon: 'blur_off' },
      { separator: true }, { id: 'picture.effects', arg: 'frame', label: 'Simple Frame, White', icon: 'filter_frames' }, { id: 'picture.effects', arg: 'polaroid', label: 'Snapshot Frame', icon: 'photo' },
      { separator: true }, { id: 'picture.effects', arg: 'none', label: 'Remove All Effects', icon: 'layers_clear' },
    ],
  },
  {
    id: 'picture.change', label: 'Change Picture', icon: 'published_with_changes', tab: 'Picture Format', group: 'Adjust', edits: true, isEnabled: hasPicture,
    run: async (app, arg) => {
      const o = findObject(app, idOf(arg));
      if (!o || o.kind !== 'drawing' || o.def.kind !== 'image') { toast('Select a picture first.'); return; }
      const files = await pickFile('image/*');
      if (!files.length) return;
      const src = await readFileAsDataURL(files[0]);
      if (!/^data:image\//.test(src)) { toast('That file is not a picture.'); return; }
      let ratio = 1;
      try { const img = await loadImage(src); ratio = img.naturalHeight / (img.naturalWidth || 1); } catch { toast('The picture could not be read.'); return; }
      mutateObject(app, o, 'Change Picture', (d: DrawingObject) => { d.src = src; delete d.orig; delete d.fx; d.at = { ...d.at, h: Math.max(8, d.at.w * ratio) }; d.alt = d.alt || files[0].name.replace(/\.[^.]+$/, ''); });
    },
  },
  {
    id: 'picture.reset', label: 'Reset Picture', icon: 'restart_alt', tab: 'Picture Format', group: 'Adjust', kind: 'split', edits: true, isEnabled: hasPicture,
    run: async (app, arg) => {
      const o = findObject(app, idOf(arg));
      if (!o || o.kind !== 'drawing' || o.def.kind !== 'image') { toast('Select a picture first.'); return; }
      const size = arg === 'size';
      let nat: { w: number; h: number } | null = null;
      if (size) { try { const img = await loadImage(o.def.orig || o.def.src || ''); nat = { w: img.naturalWidth, h: img.naturalHeight }; } catch { /* ignore */ } }
      mutateObject(app, o, 'Reset Picture', (d: DrawingObject) => {
        if (d.orig) { d.src = d.orig; delete d.orig; }
        delete d.fx; delete d.rotate;
        if (nat) { const k = Math.min(1, 640 / Math.max(nat.w, nat.h)); d.at = { ...d.at, w: nat.w * k, h: nat.h * k }; }
      });
    },
    menu: () => [{ id: 'picture.reset', label: 'Reset Picture', icon: 'restart_alt' }, { id: 'picture.reset', arg: 'size', label: 'Reset Picture & Size', icon: 'aspect_ratio' }],
  },
  {
    id: 'picture.border', label: 'Picture Border', icon: 'border_outer', tab: 'Picture Format', group: 'Picture Styles', kind: 'menu', edits: true, isEnabled: hasPicture,
    run: async (app, arg) => {
      const d = selectedDrawing(app);
      const cur = (d?.fx as PictureFx | undefined)?.border;
      if (arg === '__more') { const c = await pickMore(app, 'picture.border'); if (c) await applyPictureFx(app, null, { border: { color: c, width: cur?.width || 3 } }, 'Picture Border'); return; }
      if (arg === 'none') return applyPictureFx(app, null, { border: undefined }, 'Picture Border');
      if (typeof arg === 'number') return applyPictureFx(app, null, { border: { color: cur?.color || '#000000', width: arg } }, 'Picture Border');
      return applyPictureFx(app, null, { border: { color: typeof arg === 'string' ? arg : cur?.color || '#000000', width: cur?.width || 3 } }, 'Picture Border');
    },
    menu: () => colorEntries('picture.border', [{ id: 'picture.border', arg: 'none', label: 'No Outline', icon: 'block' }, { label: 'Weight', icon: 'line_weight', submenu: [1, 2, 3, 5, 8, 12].map((n) => ({ id: 'picture.border', arg: n, label: `${n} px` })) }, { separator: true }]),
  },
  { id: 'picture.shadow', label: 'Picture Shadow', icon: 'shadow', tab: 'Picture Format', group: 'Picture Styles', kind: 'toggle', edits: true, isEnabled: hasPicture, isActive: (app) => !!selectedDrawing(app)?.fx?.shadow, run: (app) => applyPictureFx(app, null, { shadow: selectedDrawing(app)?.fx?.shadow ? undefined : true }, 'Picture Shadow') },
  {
    id: 'picture.altText', label: 'Alt Text', icon: 'image_search', tab: 'Picture Format', group: 'Accessibility', edits: true, isEnabled: hasObj,
    run: async (app, arg) => {
      const o = findObject(app, idOf(arg));
      if (!o) { toast('Select a picture, shape or chart first.'); return; }
      const cur = o.kind === 'drawing' ? o.def.alt || '' : (o.def as any).alt || '';
      const r = await openDialog({
        title: 'Alt Text', help: 'Describe the object and its context for people who use screen readers (one or two sentences).',
        fields: [
          { key: 'alt', label: 'Description', type: 'textarea', rows: 4, value: cur, autofocus: true },
          { key: 'decorative', label: 'Mark as decorative (screen readers skip it)', type: 'checkbox', value: !!(o.def as any).fx?.decorative },
        ],
      });
      if (!r) return;
      mutateObject(app, o, 'Alt Text', (d: any) => {
        d.alt = r.decorative ? undefined : String(r.alt || '').trim() || undefined;
        if (o.kind === 'drawing') { d.fx = { ...(d.fx || {}) }; if (r.decorative) d.fx.decorative = true; else delete d.fx.decorative; if (!Object.keys(d.fx).length) delete d.fx; }
      });
    },
  },
  {
    id: 'picture.crop', label: 'Crop', icon: 'crop', tab: 'Picture Format', group: 'Size', kind: 'split', edits: true, isEnabled: hasPicture,
    run: async (app, arg) => {
      const d = selectedDrawing(app, arg);
      if (!d || d.kind !== 'image') { toast('Select a picture first.'); return; }
      if (typeof arg === 'string') {
        const shapes: Record<string, Partial<PictureFx>> = { oval: { shape: 'ellipse' }, rounded: { shape: 'rounded' }, rect: { shape: undefined } };
        if (arg === 'square' || arg === '4:3' || arg === '16:9' || arg === '3:2') {
          const [a, b] = arg === 'square' ? [1, 1] : arg.split(':').map(Number);
          try {
            const img = await loadImage(d.orig || d.src || '');
            const W = img.naturalWidth, H = img.naturalHeight;
            const target = a / b;
            let crop = { t: 0, r: 0, b: 0, l: 0 };
            if (W / H > target) { const cut = (1 - (H * target) / W) * 50; crop = { t: 0, b: 0, l: cut, r: cut }; } else { const cut = (1 - W / target / H) * 50; crop = { l: 0, r: 0, t: cut, b: cut }; }
            await applyPictureFx(app, arg && typeof arg === 'object' ? arg : null, { crop }, 'Crop to Aspect Ratio');
          } catch { toast('The picture could not be read.'); }
          return;
        }
        if (shapes[arg]) return applyPictureFx(app, null, shapes[arg], 'Crop to Shape');
      }
      const fx = (d.fx || {}) as PictureFx;
      const c = fx.crop || { t: 0, r: 0, b: 0, l: 0 };
      const src = escHtml(d.orig || d.src || '');
      const r = await openDialog({
        title: 'Crop Picture', help: 'Percent of the original picture to trim from each edge.', preview: true, previewStyle: 'display:flex;justify-content:center;background:#fff',
        fields: [
          { key: 't', label: 'Top (%)', type: 'number', value: Math.round(c.t * 10) / 10, min: 0, max: 95, half: true }, { key: 'b', label: 'Bottom (%)', type: 'number', value: Math.round(c.b * 10) / 10, min: 0, max: 95, half: true },
          { key: 'l', label: 'Left (%)', type: 'number', value: Math.round(c.l * 10) / 10, min: 0, max: 95, half: true }, { key: 'r', label: 'Right (%)', type: 'number', value: Math.round(c.r * 10) / 10, min: 0, max: 95, half: true },
          { key: 'shape', label: 'Crop to shape', type: 'select', options: [{ value: 'none', label: 'Rectangle' }, { value: 'rounded', label: 'Rounded Rectangle' }, { value: 'ellipse', label: 'Oval' }], value: fx.shape || 'none' },
        ],
        onChange: (v, ctl) => ctl.setPreview(`<img alt="" src="${src}" style="max-width:100%;max-height:170px;clip-path:${v.shape === 'ellipse' ? 'ellipse(50% 50%)' : `inset(${+v.t || 0}% ${+v.r || 0}% ${+v.b || 0}% ${+v.l || 0}%${v.shape === 'rounded' ? ' round 12%' : ''})`}">`),
        validate: (v) => ((+v.t || 0) + (+v.b || 0) >= 98 || (+v.l || 0) + (+v.r || 0) >= 98 ? 'The crop would remove the whole picture.' : null),
      });
      if (!r) return;
      const crop = { t: Math.max(0, +r.t || 0), r: Math.max(0, +r.r || 0), b: Math.max(0, +r.b || 0), l: Math.max(0, +r.l || 0) };
      await applyPictureFx(app, null, { crop: crop.t || crop.r || crop.b || crop.l ? crop : undefined, shape: r.shape === 'none' ? undefined : r.shape }, 'Crop');
    },
    menu: () => [
      { id: 'picture.crop', label: 'Crop…', icon: 'crop' }, { separator: true },
      { label: 'Crop to Shape', icon: 'interests', submenu: [{ id: 'picture.crop', arg: 'rect', label: 'Rectangle' }, { id: 'picture.crop', arg: 'rounded', label: 'Rounded Rectangle' }, { id: 'picture.crop', arg: 'oval', label: 'Oval' }] },
      { label: 'Aspect Ratio', icon: 'aspect_ratio', submenu: [{ id: 'picture.crop', arg: 'square', label: 'Square 1:1' }, { id: 'picture.crop', arg: '4:3', label: 'Landscape 4:3' }, { id: 'picture.crop', arg: '3:2', label: 'Landscape 3:2' }, { id: 'picture.crop', arg: '16:9', label: 'Landscape 16:9' }] },
    ],
  },
  { id: 'picture.size', label: 'Size and Properties', icon: 'photo_size_select_large', tab: 'Picture Format', group: 'Size', edits: true, isEnabled: hasObj, run: (app, arg) => sizeDialog(app, arg) },
  { id: 'picture.format', label: 'Format Picture', icon: 'photo_size_select_large', tab: 'Picture Format', group: 'Size', edits: true, run: (app, arg) => sizeDialog(app, arg) },

  // ---- Shape Format
  { id: 'drawing.editText', label: 'Edit Text', icon: 'edit_note', tab: 'Shape Format', group: 'Insert Shapes', edits: true, run: (app, arg) => editText(app, arg) },
  {
    id: 'drawing.fill', label: 'Shape Fill', icon: 'format_color_fill', tab: 'Shape Format', group: 'Shape Styles', kind: 'menu', edits: true, isEnabled: hasShape,
    run: async (app, arg) => {
      let v = typeof arg === 'string' ? arg : '#2f6fb0';
      if (v === '__more') { const c = await pickMore(app, 'drawing.fill'); if (!c) return; v = c; }
      if (v.startsWith('gradient:') && v.includes('{c}')) v = v.replace('{c}', currentDrawingColor(app));
      updateDrawing(app, null, 'Shape Fill', (d) => { if (d.kind === 'wordart' || d.kind === 'icon') { if (d.kind === 'icon') d.fill = v; else d.font = { ...(d.font || {}), color: v }; } else d.fill = v; });
    },
    menu: () => colorEntries('drawing.fill', [{ id: 'drawing.fill', arg: 'none', label: 'No Fill', icon: 'format_color_reset' }, { label: 'Gradient', icon: 'gradient', submenu: GRADIENTS.map(([l, g]) => ({ id: 'drawing.fill', arg: g, label: l })) }, { separator: true }]),
  },
  {
    id: 'drawing.outline', label: 'Shape Outline', icon: 'border_color', tab: 'Shape Format', group: 'Shape Styles', kind: 'menu', edits: true, isEnabled: hasShape,
    run: async (app, arg) => {
      if (arg && typeof arg === 'object') { updateDrawing(app, null, 'Shape Outline', (d) => { if (arg.width != null) d.strokeWidth = arg.width; if (arg.dash != null) d.fx = { ...(d.fx || {}), dash: arg.dash === 'solid' ? undefined : arg.dash }; }); return; }
      let v = typeof arg === 'string' ? arg : '#1f3b57';
      if (v === '__more') { const c = await pickMore(app, 'drawing.outline'); if (!c) return; v = c; }
      updateDrawing(app, null, 'Shape Outline', (d) => { if (d.kind === 'wordart') d.fx = { ...(d.fx || {}), textOutline: v }; else d.stroke = v; });
    },
    menu: () => colorEntries('drawing.outline', [
      { id: 'drawing.outline', arg: 'none', label: 'No Outline', icon: 'block' },
      { label: 'Weight', icon: 'line_weight', submenu: [0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6].map((n) => ({ id: 'drawing.outline', arg: { width: n }, label: `${n} pt` })) },
      { label: 'Dashes', icon: 'line_style', submenu: [['solid', 'Solid'], ['dash', 'Dash'], ['dot', 'Round Dot']].map(([a, l]) => ({ id: 'drawing.outline', arg: { dash: a }, label: l })) },
      { separator: true },
    ]),
  },
  {
    id: 'drawing.effects', label: 'Shape Effects', icon: 'auto_awesome', tab: 'Shape Format', group: 'Shape Styles', kind: 'menu', edits: true, isEnabled: hasShape,
    run: (app, arg) => updateDrawing(app, null, 'Shape Effects', (d) => {
      const fx = { ...(d.fx || {}) };
      const a = String(arg || 'shadow');
      if (a === 'shadow') fx.shadow = !fx.shadow;
      else if (a === 'none') { delete fx.shadow; delete fx.glow; delete fx.soft; }
      else if (a === 'soft') fx.soft = fx.soft ? undefined : 2;
      else if (a === 'noGlow') delete fx.glow;
      else if (a.startsWith('glow:')) fx.glow = a.slice(5);
      for (const k of Object.keys(fx)) if (fx[k] == null) delete fx[k];
      d.fx = Object.keys(fx).length ? fx : undefined;
    }),
    menu: (app) => {
      const fx = selectedDrawing(app)?.fx || {};
      return [
        { id: 'drawing.effects', arg: 'shadow', label: 'Shadow', icon: 'shadow', checked: !!fx.shadow },
        { label: 'Glow', icon: 'flare', submenu: [{ id: 'drawing.effects', arg: 'noGlow', label: 'No Glow' }, ...['#ffc000', '#4a90c2', '#5aa33f', '#e07b28', '#c00000', '#7b3fa0'].map((c) => ({ id: 'drawing.effects', arg: 'glow:' + c, label: c, swatch: c }))] },
        { id: 'drawing.effects', arg: 'soft', label: 'Soft Edges', icon: 'blur_on', checked: !!fx.soft },
        { separator: true }, { id: 'drawing.effects', arg: 'none', label: 'No Effects', icon: 'layers_clear' },
      ];
    },
  },
  {
    id: 'drawing.textFill', label: 'Text Fill', icon: 'format_color_text', tab: 'Shape Format', group: 'WordArt Styles', kind: 'menu', edits: true, isEnabled: hasShape,
    run: async (app, arg) => {
      let v = typeof arg === 'string' ? arg : '#000000';
      if (v === '__more') { const c = await pickMore(app, 'drawing.textFill'); if (!c) return; v = c; }
      updateDrawing(app, null, 'Text Fill', (d) => {
        if (v.startsWith('gradient:') || v === 'none') { d.fx = { ...(d.fx || {}), textFill: v }; if (d.kind === 'wordart') d.font = { ...(d.font || {}), color: v }; }
        else { d.font = { ...(d.font || {}), color: v }; if (d.fx?.textFill) { const fx = { ...d.fx }; delete fx.textFill; d.fx = fx; } }
      });
    },
    menu: () => colorEntries('drawing.textFill', [{ label: 'Gradient', icon: 'gradient', submenu: [['Blue to Purple', 'gradient:#2f6fb0:#7b3fa0:0'], ['Gold', 'gradient:#e3b505:#9e480e:90'], ['Sunset', 'gradient:#ff6a00:#ee0979:90'], ['Ocean', 'gradient:#00b0f0:#1f3b57:90']].map(([l, g]) => ({ id: 'drawing.textFill', arg: g, label: l })) }, { separator: true }]),
  },
  {
    id: 'drawing.textOutline', label: 'Text Outline', icon: 'text_format', tab: 'Shape Format', group: 'WordArt Styles', kind: 'menu', edits: true, isEnabled: hasShape,
    run: async (app, arg) => {
      if (arg && typeof arg === 'object') { updateDrawing(app, null, 'Text Outline', (d) => { d.fx = { ...(d.fx || {}), textOutlineW: arg.width }; }); return; }
      let v = typeof arg === 'string' ? arg : '#000000';
      if (v === '__more') { const c = await pickMore(app, 'drawing.textOutline'); if (!c) return; v = c; }
      updateDrawing(app, null, 'Text Outline', (d) => { const fx = { ...(d.fx || {}) }; if (v === 'none') delete fx.textOutline; else fx.textOutline = v; d.fx = Object.keys(fx).length ? fx : undefined; });
    },
    menu: () => colorEntries('drawing.textOutline', [{ id: 'drawing.textOutline', arg: 'none', label: 'No Outline', icon: 'block' }, { label: 'Weight', icon: 'line_weight', submenu: [0.5, 1, 1.5, 2.5, 4].map((n) => ({ id: 'drawing.textOutline', arg: { width: n }, label: `${n} pt` })) }, { separator: true }]),
  },
  { id: 'drawing.size', label: 'Size', icon: 'aspect_ratio', tab: 'Shape Format', group: 'Size', edits: true, isEnabled: hasObj, run: (app, arg) => sizeDialog(app, arg) },
  {
    id: 'drawing.changeShape', label: 'Change Shape', icon: 'category', tab: 'Shape Format', group: 'Insert Shapes', edits: true, isEnabled: hasShape,
    run: async (app, arg) => {
      const d = selectedDrawing(app);
      if (!d || d.kind !== 'shape') { toast('Select a shape first.'); return; }
      const v = typeof arg === 'string' ? arg : await galleryDialog('Change Shape', SHAPES.map((s) => ({ value: s.id, label: s.label, group: s.group, html: shapePreviewSvg(s.id, 34) })), { size: 'small', current: d.shape, search: true });
      if (v) updateDrawing(app, null, 'Change Shape', (x) => { x.shape = v; });
    },
  },
  {
    id: 'drawing.wordArtStyle', label: 'Decorative Text Styles', icon: 'format_shapes', tab: 'Shape Format', group: 'WordArt Styles', kind: 'menu', edits: true, isEnabled: hasShape,
    run: (app, arg) => {
      const st = WORDART_STYLES.find((s) => s.id === arg) || WORDART_STYLES[0];
      updateDrawing(app, null, 'Decorative Text Style', (d) => {
        d.font = { ...(d.font || {}), color: st.fill, bold: st.bold !== false, ...(st.font ? { font: st.font } : {}) };
        d.fx = { ...(d.fx || {}), textOutline: st.outline, textOutlineW: st.outlineW, shadow: st.shadow || undefined, glow: st.glow };
        for (const k of Object.keys(d.fx)) if (d.fx[k] == null) delete d.fx[k];
      });
    },
    menu: () => WORDART_STYLES.map((s) => ({ id: 'drawing.wordArtStyle', arg: s.id, label: s.label, icon: 'format_shapes' })),
  },
  {
    id: 'object.formatPane', label: 'Format Object', icon: 'format_shapes', tab: 'Shape Format', group: 'Shape Styles', edits: true, isEnabled: hasObj,
    run: (app, arg) => { const o = findObject(app, idOf(arg)); if (!o) return toast('Select an object first.'); if (o.kind === 'chart') return app.exec('chart.edit', { id: o.def.id }); if (o.def.kind === 'image') return sizeDialog(app, arg); return editText(app, { id: o.def.id }); },
  },
]);


/** Shows a short note when a dialog is not usable on this platform. */
export async function unsupported(what: string): Promise<void> { await alertDialog(what, `${what} isn't available on this device.`); }
void isPhone;
