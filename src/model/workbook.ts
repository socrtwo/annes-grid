/**
 * Workbook and sheet model with transactional undo/redo.
 *
 * Cells are stored sparsely per sheet in a Map keyed by cellKey(r, c). Every mutation made inside
 * Workbook.transact() is recorded: single-cell edits as before-images of the touched cells, and
 * structural edits (insert/delete rows, sort, sheet settings) as a full before-snapshot of the sheet.
 * Undo restores the before-images and captures the current state as the redo record, so undo and
 * redo are symmetric and never need per-operation inverse logic.
 */
import { cellKey, keyRow, keyCol, MAX_COLS, MAX_ROWS, type Range } from './address';
import type {
  Cell, CellStyle, CellComment, DataValidation, CFRule, AutoFilter, TableDef, ChartDef, DrawingObject, Sparkline,
  PrintSettings, SheetProtection, AllowEditRange, Scenario, PivotDef, DefinedName, NamedCellStyle, Scalar,
} from './types';

export const DEFAULT_COL_W = 64;   // px at 100%
export const DEFAULT_ROW_H = 20;   // px at 100%

let idSeq = 0;
export const newId = (p = 'id') => `${p}${Date.now().toString(36)}${(idSeq++).toString(36)}`;

export function defaultPrint(): PrintSettings {
  return {
    orientation: 'portrait', paper: 'Letter',
    margins: { top: 0.75, bottom: 0.75, left: 0.7, right: 0.7, header: 0.3, footer: 0.3 },
    scale: 100, gridlines: false, headings: false, centerH: false, centerV: false,
    header: { left: '', center: '', right: '' }, footer: { left: '', center: '&[Page]', right: '' },
    order: 'downThenOver',
  };
}

export class Sheet {
  id = newId('s');
  name: string;
  cells = new Map<number, Cell>();
  colW = new Map<number, number>();
  rowH = new Map<number, number>();
  hiddenRows = new Set<number>();
  hiddenCols = new Set<number>();
  rowLevel = new Map<number, number>();
  colLevel = new Map<number, number>();
  rowStyle = new Map<number, number>();
  colStyle = new Map<number, number>();
  merges: Range[] = [];
  freeze = { rows: 0, cols: 0 };
  split: { x: number; y: number } | null = null;
  comments = new Map<number, CellComment>();
  validations: DataValidation[] = [];
  cf: CFRule[] = [];
  filter: AutoFilter | null = null;
  tables: TableDef[] = [];
  charts: ChartDef[] = [];
  drawings: DrawingObject[] = [];
  sparklines: Sparkline[] = [];
  pivots: PivotDef[] = [];
  scenarios: Scenario[] = [];
  print: PrintSettings = defaultPrint();
  rowBreaks: number[] = [];
  colBreaks: number[] = [];
  protection: SheetProtection = { on: false, allow: ['selectLocked', 'selectUnlocked'] };
  tabColor: string | null = null;
  visibility: 'visible' | 'hidden' | 'veryHidden' = 'visible';
  showGrid = true;
  showHeadings = true;
  showFormulas = false;
  showZeros = true;
  rtl = false;
  defaultColW = DEFAULT_COL_W;
  defaultRowH = DEFAULT_ROW_H;
  zoom = 1;
  gridColor: string | null = null;
  /** View state (not undone). */
  sel = { r: 0, c: 0, ranges: [{ r1: 0, c1: 0, r2: 0, c2: 0 }] as Range[] };
  scroll = { r: 0, c: 0 };
  /** Code name used by macros. */
  codeName = '';
  /** Sheet background picture (data URL), printed behind the cells. */
  background: string | null = null;
  /** Ranges that stay editable on a protected sheet (Allow Edit Ranges). */
  editRanges: AllowEditRange[] = [];

  constructor(public book: Workbook, name: string) { this.name = name; }

  // ------------------------------------------------------------------------------------------
  // Cells

  get(r: number, c: number): Cell | undefined { return this.cells.get(cellKey(r, c)); }
  value(r: number, c: number): Scalar { const x = this.cells.get(cellKey(r, c)); return x?.v ?? null; }

  /** Sets or clears (cell === undefined or empty) a cell, recording the old value for undo. */
  set(r: number, c: number, cell: Cell | undefined): void {
    const k = cellKey(r, c);
    this.book.recordCell(this, k);
    if (!cell || (cell.v == null && cell.f == null && !cell.s && !cell.link && !cell.rich && cell.sp == null && !cell.arr)) this.cells.delete(k);
    else this.cells.set(k, cell);
    this.book.noteChange(this, k);
  }

  /** Merges partial changes into a cell. */
  patch(r: number, c: number, p: Partial<Cell>): void {
    const cur = this.get(r, c);
    const next: Cell = { ...(cur || {}), ...p };
    for (const k of Object.keys(next) as (keyof Cell)[]) if (next[k] === undefined) delete next[k];
    this.set(r, c, next);
  }

  /** Sets a cell value without touching undo history (calc results, spills). */
  setComputed(r: number, c: number, cell: Cell | undefined): void {
    const k = cellKey(r, c);
    if (!cell || (cell.v == null && cell.f == null && !cell.s && cell.sp == null && !cell.link && !cell.arr)) this.cells.delete(k);
    else this.cells.set(k, cell);
  }

  styleOf(r: number, c: number): CellStyle {
    const s = this.get(r, c)?.s ?? this.rowStyle.get(r) ?? this.colStyle.get(c) ?? 0;
    return this.book.styles[s] || {};
  }

  styleIdOf(r: number, c: number): number {
    return this.get(r, c)?.s ?? this.rowStyle.get(r) ?? this.colStyle.get(c) ?? 0;
  }

  /** Bounds of cells that have content or formatting. */
  usedRange(includeFormat = true): Range | null {
    let r1 = Infinity, c1 = Infinity, r2 = -1, c2 = -1;
    for (const [k, cell] of this.cells) {
      if (!includeFormat && cell.v == null && cell.f == null) continue;
      const r = keyRow(k), c = keyCol(k);
      if (r < r1) r1 = r; if (r > r2) r2 = r; if (c < c1) c1 = c; if (c > c2) c2 = c;
    }
    if (r2 < 0) return null;
    return { r1, c1, r2, c2 };
  }

  /** Iterates the existing cells inside a range (sparse; fast for large ranges). */
  forEachIn(rg: Range, fn: (r: number, c: number, cell: Cell) => void): void {
    const area = (rg.r2 - rg.r1 + 1) * (rg.c2 - rg.c1 + 1);
    if (area > this.cells.size * 2) {
      for (const [k, cell] of this.cells) {
        const r = keyRow(k), c = keyCol(k);
        if (r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2) fn(r, c, cell);
      }
    } else {
      for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) { const cell = this.cells.get(cellKey(r, c)); if (cell) fn(r, c, cell); }
    }
  }

  // ------------------------------------------------------------------------------------------
  // Geometry

  colWidth(c: number): number { return this.hiddenCols.has(c) ? 0 : this.colW.get(c) ?? this.defaultColW; }
  rowHeight(r: number): number { return this.hiddenRows.has(r) ? 0 : this.rowH.get(r) ?? this.defaultRowH; }

  mergeAt(r: number, c: number): Range | null {
    for (const m of this.merges) if (r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2) return m;
    return null;
  }

  // ------------------------------------------------------------------------------------------
  // Serialization (used for undo snapshots and the native file format)

  toJSON(): any {
    const cells: [number, Cell][] = [];
    for (const [k, v] of this.cells) cells.push([k, v]);
    return {
      id: this.id, name: this.name, cells,
      colW: [...this.colW], rowH: [...this.rowH], hiddenRows: [...this.hiddenRows], hiddenCols: [...this.hiddenCols],
      rowLevel: [...this.rowLevel], colLevel: [...this.colLevel], rowStyle: [...this.rowStyle], colStyle: [...this.colStyle],
      merges: this.merges, freeze: this.freeze, split: this.split, comments: [...this.comments], validations: this.validations, cf: this.cf,
      filter: this.filter, tables: this.tables, charts: this.charts, drawings: this.drawings, sparklines: this.sparklines, pivots: this.pivots,
      scenarios: this.scenarios, print: this.print, rowBreaks: this.rowBreaks, colBreaks: this.colBreaks, protection: this.protection,
      tabColor: this.tabColor, visibility: this.visibility, showGrid: this.showGrid, showHeadings: this.showHeadings, showFormulas: this.showFormulas,
      showZeros: this.showZeros, rtl: this.rtl, defaultColW: this.defaultColW, defaultRowH: this.defaultRowH, zoom: this.zoom, gridColor: this.gridColor, codeName: this.codeName,
      background: this.background, editRanges: this.editRanges,
    };
  }

  static fromJSON(book: Workbook, j: any): Sheet {
    const s = new Sheet(book, String(j.name ?? 'Sheet'));
    s.loadJSON(j);
    return s;
  }

  loadJSON(j: any): void {
    const clone = structuredCloneSafe(j);
    if (clone.id) this.id = clone.id;
    this.name = String(clone.name ?? this.name);
    this.cells = new Map((clone.cells || []).map(([k, v]: [number, Cell]) => [k, reviveCell(v)]));
    this.colW = sizeMap(clone.colW, 4000); this.rowH = sizeMap(clone.rowH, 4000);
    this.hiddenRows = new Set(clone.hiddenRows || []); this.hiddenCols = new Set(clone.hiddenCols || []);
    this.rowLevel = new Map(clone.rowLevel || []); this.colLevel = new Map(clone.colLevel || []);
    this.rowStyle = new Map(clone.rowStyle || []); this.colStyle = new Map(clone.colStyle || []);
    this.merges = clone.merges || []; this.freeze = clone.freeze || { rows: 0, cols: 0 }; this.split = clone.split || null;
    this.comments = new Map(clone.comments || []);
    this.validations = clone.validations || []; this.cf = clone.cf || []; this.filter = clone.filter || null;
    this.tables = clone.tables || []; this.charts = anchored(clone.charts); this.drawings = anchored(clone.drawings); this.sparklines = clone.sparklines || [];
    this.pivots = clone.pivots || []; this.scenarios = (clone.scenarios || []).map((x: Scenario) => ({ ...x, values: x.values.map(reviveScalar) }));
    this.print = { ...defaultPrint(), ...(clone.print || {}) };
    this.rowBreaks = clone.rowBreaks || []; this.colBreaks = clone.colBreaks || [];
    this.protection = clone.protection || { on: false, allow: ['selectLocked', 'selectUnlocked'] };
    this.tabColor = clone.tabColor ?? null; this.visibility = clone.visibility || 'visible';
    this.showGrid = clone.showGrid ?? true; this.showHeadings = clone.showHeadings ?? true; this.showFormulas = !!clone.showFormulas;
    this.showZeros = clone.showZeros ?? true; this.rtl = !!clone.rtl;
    this.defaultColW = finite(clone.defaultColW, DEFAULT_COL_W, 1, 4000); this.defaultRowH = finite(clone.defaultRowH, DEFAULT_ROW_H, 1, 4000);
    this.zoom = finite(clone.zoom, 1, 0.1, 5); this.gridColor = clone.gridColor ?? null; this.codeName = clone.codeName || '';
    this.background = clone.background ?? null; this.editRanges = clone.editRanges || [];
  }
}

// A workbook file is untrusted input: sizes and positions are interpolated into markup and CSS by
// the renderers, so anything that should be a number is forced to be one when a file is loaded.
function finite(v: unknown, fallback: number, min: number, max: number): number {
  const n = typeof v === 'number' ? v : NaN;
  return Number.isFinite(n) && n > 0 ? Math.min(max, Math.max(min, n)) : fallback;
}

function sizeMap(entries: unknown, max: number): Map<number, number> {
  const out = new Map<number, number>();
  if (!Array.isArray(entries)) return out;
  for (const e of entries) {
    if (!Array.isArray(e)) continue;
    const k = e[0], v = e[1];
    if (Number.isInteger(k) && k >= 0 && typeof v === 'number' && Number.isFinite(v) && v >= 0) out.set(k, Math.min(max, v));
  }
  return out;
}

function anchored<T extends { at: { r: number; c: number; dx: number; dy: number; w: number; h: number } }>(list: unknown): T[] {
  if (!Array.isArray(list)) return [];
  const num = (v: unknown, fb: number, min: number, max: number) => (typeof v === 'number' && Number.isFinite(v) ? Math.min(max, Math.max(min, v)) : fb);
  const out: T[] = [];
  for (const o of list as any[]) {
    if (!o || typeof o !== 'object') continue;
    const a = o.at && typeof o.at === 'object' ? o.at : {};
    o.at = { ...a, r: Math.trunc(num(a.r, 0, 0, 1048575)), c: Math.trunc(num(a.c, 0, 0, 16383)), dx: num(a.dx, 0, -1e5, 1e5), dy: num(a.dy, 0, -1e5, 1e5), w: num(a.w, 200, 1, 20000), h: num(a.h, 120, 1, 20000) };
    if ('rotate' in o) o.rotate = num(o.rotate, 0, -3600, 3600);
    out.push(o as T);
  }
  return out;
}

/** JSON round trip that keeps CellError values (serialized as {"#err": code}). */
export function structuredCloneSafe<T>(j: T): T {
  return JSON.parse(JSON.stringify(j, errReplacer), errReviver);
}

import { CellError, type ErrorCode } from './types';
export function errReplacer(_k: string, v: any): any { return v instanceof CellError ? { '#err': v.code } : v; }
export function errReviver(_k: string, v: any): any { return v && typeof v === 'object' && typeof v['#err'] === 'string' && Object.keys(v).length === 1 ? new CellError(v['#err'] as ErrorCode) : v; }
function reviveScalar(v: any): Scalar { return v && typeof v === 'object' && '#err' in v ? new CellError(v['#err']) : v; }
function reviveCell(c: any): Cell { if (c && c.v && typeof c.v === 'object' && !(c.v instanceof CellError)) c.v = reviveScalar(c.v); return c; }

// --------------------------------------------------------------------------------------------
// Workbook

export interface BookProps {
  title: string; subject: string; author: string; manager: string; company: string; category: string;
  keywords: string; comments: string; status: string; created: number; modified: number; lastModifiedBy: string;
  custom: Record<string, string>;
}

export interface CalcSettings {
  mode: 'auto' | 'manual' | 'autoNoTables';
  iterative: boolean; maxIterations: number; maxChange: number;
  precisionAsDisplayed: boolean;
  date1904: boolean;
  recalcBeforeSave: boolean;
}

export interface ThemeDef { name: string; colors: string[]; headingFont: string; bodyFont: string; /** Theme effects name (stored with the theme). */ effects?: string }

export const DEFAULT_THEME: ThemeDef = {
  name: 'Quill',
  // dark1, light1, dark2, light2, accent1..6, hyperlink, followed
  colors: ['#000000', '#ffffff', '#1f3b57', '#e8eef4', '#2f6fb0', '#e07b28', '#7a8a99', '#e3b505', '#4a90c2', '#5aa33f', '#1565c0', '#7b3fa0'],
  headingFont: 'Carlito', bodyFont: 'Carlito',
};

interface TxRecord {
  label: string;
  cells: Map<Sheet, Map<number, Cell | undefined>>;
  sheets: Map<string, any>;   // sheet id -> full before-snapshot
  book: any | null;            // workbook-level before-snapshot (sheet list, names, settings)
  selection?: { sheet: string; sel: Sheet['sel'] };
}

export interface ChangeSet { cells: Map<Sheet, Set<number>>; structural: boolean }

export class Workbook {
  sheets: Sheet[] = [];
  styles: CellStyle[] = [{}];
  private styleIndex = new Map<string, number>([['{}', 0]]);
  names: DefinedName[] = [];
  cellStyles: NamedCellStyle[] = [];
  customFormats: string[] = [];
  props: BookProps = { title: '', subject: '', author: '', manager: '', company: '', category: '', keywords: '', comments: '', status: '', created: Date.now(), modified: Date.now(), lastModifiedBy: '', custom: {} };
  calc: CalcSettings = { mode: 'auto', iterative: false, maxIterations: 100, maxChange: 0.001, precisionAsDisplayed: false, date1904: false, recalcBeforeSave: true };
  theme: ThemeDef = { ...DEFAULT_THEME, colors: [...DEFAULT_THEME.colors] };
  structureLocked = false;
  structurePassword?: string;
  activeSheet = 0;
  /** JavaScript macro modules. */
  modules: { name: string; code: string }[] = [];
  /** Shared (co-authoring style) markers, reserved. */
  readOnlyRecommended = false;
  finalMarked = false;

  undoStack: TxRecord[] = [];
  redoStack: TxRecord[] = [];
  private tx: TxRecord | null = null;
  private txDepth = 0;
  /** Changes since the last recalculation. */
  pending: ChangeSet = { cells: new Map(), structural: false };
  /** Called after every completed top-level transaction. */
  onChange: ((label: string) => void) | null = null;
  version = 0;

  constructor(init = true) {
    if (init) this.sheets.push(new Sheet(this, 'Sheet1'));
  }

  get sheet(): Sheet { return this.sheets[Math.min(this.activeSheet, this.sheets.length - 1)]; }

  sheetByName(name: string): Sheet | undefined {
    const n = name.toLowerCase();
    return this.sheets.find((s) => s.name.toLowerCase() === n);
  }
  sheetById(id: string): Sheet | undefined { return this.sheets.find((s) => s.id === id); }

  uniqueSheetName(base = 'Sheet'): string {
    let i = this.sheets.length + 1;
    let name = `${base}${i}`;
    while (this.sheetByName(name)) name = `${base}${++i}`;
    return name;
  }

  // ------------------------------------------------------------------------------------------
  // Styles

  styleId(st: CellStyle): number {
    const { key, clean } = canonicalStyle(st);
    let id = this.styleIndex.get(key);
    if (id == null) { id = this.styles.length; this.styles.push(clean); this.styleIndex.set(key, id); }
    return id;
  }

  /** Style id obtained by applying a patch to an existing style id. undefined in the patch removes a property. */
  derivedStyle(base: number, patch: Partial<CellStyle>): number {
    const st: any = { ...(this.styles[base] || {}) };
    for (const [k, v] of Object.entries(patch)) { if (v === undefined || v === null) delete st[k]; else st[k] = v; }
    return this.styleId(st);
  }

  // ------------------------------------------------------------------------------------------
  // Transactions

  /** Runs fn as one undoable step. Nested calls join the outer step. */
  transact<T>(label: string, fn: () => T): T {
    const outer = this.txDepth === 0;
    if (outer) this.tx = { label, cells: new Map(), sheets: new Map(), book: null, selection: { sheet: this.sheet.id, sel: JSON.parse(JSON.stringify(this.sheet.sel)) } };
    this.txDepth++;
    try {
      return fn();
    } finally {
      this.txDepth--;
      if (outer) {
        const t = this.tx!;
        this.tx = null;
        if (t.cells.size || t.sheets.size || t.book) {
          this.undoStack.push(t);
          if (this.undoStack.length > 200) this.undoStack.shift();
          this.redoStack = [];
          this.version++;
          this.props.modified = Date.now();
        }
        this.onChange?.(label);
      }
    }
  }

  get inTransaction(): boolean { return this.txDepth > 0; }

  recordCell(sheet: Sheet, k: number): void {
    const t = this.tx;
    if (!t || t.sheets.has(sheet.id)) return;
    let m = t.cells.get(sheet);
    if (!m) { m = new Map(); t.cells.set(sheet, m); }
    if (!m.has(k)) { const cur = sheet.cells.get(k); m.set(k, cur ? { ...cur } : undefined); }
  }

  /** Records a full snapshot of the sheet before a structural change. Call before mutating. */
  recordSheet(sheet: Sheet): void {
    const t = this.tx;
    if (!t || t.sheets.has(sheet.id)) return;
    // Cells already recorded individually are folded into the snapshot: restore their before-image first.
    const snap = sheet.toJSON();
    const m = t.cells.get(sheet);
    if (m) {
      const cells = new Map<number, Cell>(snap.cells);
      for (const [k, v] of m) { if (v) cells.set(k, v); else cells.delete(k); }
      snap.cells = [...cells];
      t.cells.delete(sheet);
    }
    t.sheets.set(sheet.id, structuredCloneSafe(snap));
    this.pending.structural = true;
  }

  /** Records the workbook-level state (sheet list and order, names, settings) before changing it. */
  recordBook(): void {
    const t = this.tx;
    if (!t || t.book) return;
    t.book = this.bookMeta();
    this.pending.structural = true;
  }

  private bookMeta(): any {
    return structuredCloneSafe({
      order: this.sheets.map((s) => s.id),
      sheets: Object.fromEntries(this.sheets.map((s) => [s.id, s.toJSON()])),
      names: this.names, cellStyles: this.cellStyles, customFormats: this.customFormats, props: this.props, calc: this.calc, theme: this.theme,
      structureLocked: this.structureLocked, structurePassword: this.structurePassword, activeSheet: this.activeSheet, modules: this.modules,
      baseStyle: this.styles[0],
    });
  }

  private restoreBookMeta(m: any): void {
    const byId = new Map(this.sheets.map((s) => [s.id, s]));
    this.sheets = m.order.map((id: string) => {
      const s = byId.get(id) || new Sheet(this, 'Sheet');
      s.loadJSON(m.sheets[id]);
      return s;
    });
    this.names = m.names; this.cellStyles = m.cellStyles; this.customFormats = m.customFormats; this.props = m.props; this.calc = m.calc; this.theme = m.theme;
    this.structureLocked = m.structureLocked; this.structurePassword = m.structurePassword; this.activeSheet = Math.min(m.activeSheet, this.sheets.length - 1); this.modules = m.modules || [];
    if (m.baseStyle) this.styles[0] = m.baseStyle;
  }

  noteChange(sheet: Sheet, k: number): void {
    let s = this.pending.cells.get(sheet);
    if (!s) { s = new Set(); this.pending.cells.set(sheet, s); }
    s.add(k);
  }

  takePending(): ChangeSet {
    const p = this.pending;
    this.pending = { cells: new Map(), structural: false };
    return p;
  }

  canUndo(): boolean { return this.undoStack.length > 0; }
  canRedo(): boolean { return this.redoStack.length > 0; }
  undoLabel(): string { return this.undoStack[this.undoStack.length - 1]?.label || ''; }
  redoLabel(): string { return this.redoStack[this.redoStack.length - 1]?.label || ''; }

  undo(): string | null { return this.swap(this.undoStack, this.redoStack); }
  redo(): string | null { return this.swap(this.redoStack, this.undoStack); }

  private swap(from: TxRecord[], to: TxRecord[]): string | null {
    const t = from.pop();
    if (!t) return null;
    const inverse: TxRecord = { label: t.label, cells: new Map(), sheets: new Map(), book: null, selection: { sheet: this.sheet.id, sel: JSON.parse(JSON.stringify(this.sheet.sel)) } };
    if (t.book) {
      inverse.book = this.bookMeta();
      this.restoreBookMeta(t.book);
      this.pending.structural = true;
    }
    for (const [id, snap] of t.sheets) {
      const sh = this.sheetById(id);
      if (!sh) continue;
      if (!t.book) inverse.sheets.set(id, structuredCloneSafe(sh.toJSON()));
      sh.loadJSON(snap);
      this.pending.structural = true;
    }
    for (const [sh, m] of t.cells) {
      if (!this.sheets.includes(sh)) continue;
      const inv = new Map<number, Cell | undefined>();
      for (const [k, v] of m) {
        const cur = sh.cells.get(k);
        inv.set(k, cur ? { ...cur } : undefined);
        if (v) sh.cells.set(k, { ...v }); else sh.cells.delete(k);
        this.noteChange(sh, k);
      }
      inverse.cells.set(sh, inv);
    }
    if (t.selection) {
      const i = this.sheets.findIndex((s) => s.id === t.selection!.sheet);
      if (i >= 0) { this.activeSheet = i; this.sheets[i].sel = t.selection.sel; }
    }
    to.push(inverse);
    this.version++;
    this.onChange?.(t.label);
    return t.label;
  }

  clearHistory(): void { this.undoStack = []; this.redoStack = []; }

  // ------------------------------------------------------------------------------------------
  // Serialization

  toJSON(): any {
    return {
      format: 'quillgrid', version: 1,
      sheets: this.sheets.map((s) => s.toJSON()), styles: this.styles, names: this.names, cellStyles: this.cellStyles,
      customFormats: this.customFormats, props: this.props, calc: this.calc, theme: this.theme, structureLocked: this.structureLocked,
      structurePassword: this.structurePassword, activeSheet: this.activeSheet, modules: this.modules, finalMarked: this.finalMarked,
    };
  }

  static fromJSON(j: any): Workbook {
    const b = new Workbook(false);
    const clone = structuredCloneSafe(j);
    b.styles = Array.isArray(clone.styles) && clone.styles.length ? clone.styles : [{}];
    b.styleIndex = new Map();
    b.styles.forEach((s, i) => { const k = canonicalStyle(s).key; if (!b.styleIndex.has(k)) b.styleIndex.set(k, i); });
    b.sheets = (clone.sheets || []).map((sj: any) => Sheet.fromJSON(b, sj));
    if (!b.sheets.length) b.sheets.push(new Sheet(b, 'Sheet1'));
    b.names = clone.names || []; b.cellStyles = clone.cellStyles || []; b.customFormats = clone.customFormats || [];
    b.props = { ...b.props, ...(clone.props || {}) }; b.calc = { ...b.calc, ...(clone.calc || {}) };
    b.theme = { ...DEFAULT_THEME, ...(clone.theme || {}) };
    b.structureLocked = !!clone.structureLocked; b.structurePassword = clone.structurePassword;
    b.activeSheet = Math.max(0, Math.min(clone.activeSheet || 0, b.sheets.length - 1));
    b.modules = clone.modules || []; b.finalMarked = !!clone.finalMarked;
    b.pending.structural = true;
    return b;
  }
}

/** Canonical form of a style: drops defaults (false, empty, locked: true) and sorts keys. */
export function canonicalStyle(st: CellStyle): { key: string; clean: CellStyle } {
  const clean: any = {};
  for (const k of Object.keys(st).sort()) {
    const v = (st as any)[k];
    if (k === 'locked') { if (v === false) clean.locked = false; continue; }
    if (v !== undefined && v !== null && v !== false && v !== '') clean[k] = v;
  }
  return { key: JSON.stringify(clean), clean };
}

export { MAX_COLS, MAX_ROWS };
