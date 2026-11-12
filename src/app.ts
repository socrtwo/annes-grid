/** Anne's Grid application core: workbook tabs, command execution, recalculation, autosave. */
import { isSafeExternalUrl } from './util/sanitize';
import { Workbook, type Sheet } from './model/workbook';
import { Engine } from './calc/engine';
import { CFEvaluator } from './grid/cf';
import { getCommand, allCommands, type Command } from './commands/registry';
import { loadPrefs, savePrefs, type Prefs } from './prefs';
import { kvSet, kvDel } from './util/store';
import { toast } from './ui/dialog';
import { canonicalKey } from './util/keys';
import type { Shell } from './ui/shell';
import type { Grid } from './grid/grid';
import type { Range } from './model/address';

let tabCounter = 0;

export type ViewMode = 'normal' | 'pageBreak' | 'pageLayout';

export class BookTab {
  id = `t${Date.now().toString(36)}${tabCounter++}`;
  title = 'Book';
  fileName = '';
  format = 'xlsx';
  uri: string | null = null;
  dirty = false;
  created = Date.now();
  saved = 0;
  readOnly = false;
  viewMode: ViewMode = 'normal';
  el!: HTMLElement;
  grid!: Grid;
  engine: Engine;
  cf: CFEvaluator;
  recovery: any = null;
  /** Storage key for workbooks without a unique URI. */
  key = '';
  /** Clipboard marquee (cut/copy source) shown on the grid. */
  marquee: { sheet: Sheet; range: Range; cut: boolean } | null = null;
  /** Watch Window entries. */
  watches: { sheet: string; r: number; c: number }[] = [];
  /** Cells to trace on the grid (precedent/dependent arrows). */
  arrows: { from: { sheet: Sheet; r: number; c: number }; to: { sheet: Sheet; r: number; c: number } }[] = [];
  /** Circled invalid data. */
  circles: { r: number; c: number }[] = [];
  /** Map View (1-2-3 Release 3): cells drawn as their kind (label, number, formula) instead of their values. */
  mapView = false;
  constructor(public app: App, public book: Workbook) {
    this.engine = new Engine(book);
    this.cf = new CFEvaluator(this.engine);
  }
  get sheet(): Sheet { return this.book.sheet; }
  get displayName(): string { return this.fileName || this.title; }
}

export class App {
  prefs: Prefs = loadPrefs();
  docs: BookTab[] = [];
  active!: BookTab;
  ui!: Shell;
  recording: { name: string; steps: { id: string; arg?: any }[] } | null = null;
  recordingPaused = false;
  private refreshQueued = false;
  private autosaveTimer: any = null;
  private listeners = new Set<(app: App) => void>();
  lastCommand: { id: string; arg?: any } | null = null;
  keyIndex = new Map<string, string>();
  /** Office-style clipboard history (Clipboard pane). */
  clipboard: { text: string; html: string; time: number; source?: string }[] = [];

  constructor(public host: HTMLElement) {
    this.rebuildKeyIndex();
  }

  get tab(): BookTab { return this.active; }
  get book(): Workbook { return this.active.book; }
  get sheet(): Sheet { return this.active.book.sheet; }
  get grid(): Grid { return this.active.grid; }
  get engine(): Engine { return this.active.engine; }

  onRefresh(fn: (app: App) => void): () => void { this.listeners.add(fn); return () => this.listeners.delete(fn); }

  savePrefs(): void { savePrefs(this.prefs); this.rebuildKeyIndex(); }

  rebuildKeyIndex(): void {
    this.keyIndex.clear();
    for (const c of allCommands()) for (const k of c.keys || []) this.keyIndex.set(canonicalKey(k), c.id);
    for (const [k, id] of Object.entries(this.prefs.keymap)) {
      if (id) this.keyIndex.set(canonicalKey(k), id); else this.keyIndex.delete(canonicalKey(k));
    }
  }

  // ------------------------------------------------------------------------------------------
  // Workbooks

  createTab(book?: Workbook, meta: Partial<BookTab> = {}): BookTab {
    if (!book) {
      book = new Workbook();
      book.props.author = this.prefs.author;
      book.props.created = Date.now();
      if (this.prefs.defaultFont || this.prefs.defaultSize) {
        book.styles[0] = { ...(this.prefs.defaultFont ? { font: this.prefs.defaultFont } : {}), ...(this.prefs.defaultSize ? { size: this.prefs.defaultSize } : {}) };
      }
      const want = Math.max(1, Math.min(255, this.prefs.newSheets || 1));
      while (book.sheets.length < want) book.sheets.push(new (book.sheets[0].constructor as any)(book, book.uniqueSheetName()));
    }
    const tab = new BookTab(this, book);
    Object.assign(tab, meta);
    book.onChange = () => this.bookChanged(tab);
    this.ui.mountTab(tab);
    this.docs.push(tab);
    tab.engine.rebuild();
    tab.engine.recalc(true);
    this.activate(tab);
    return tab;
  }

  activate(tab: BookTab): void {
    this.active = tab;
    for (const d of this.docs) d.el.classList.toggle('ql-hidden-tab', d !== tab);
    this.ui.tabsChanged();
    tab.grid.layout();
    this.refresh();
    document.title = `${tab.displayName}${tab.dirty ? ' •' : ''} – Anne\'s Grid`;
  }

  closeTab(tab: BookTab): void {
    const i = this.docs.indexOf(tab);
    if (i < 0) return;
    tab.grid.destroy();
    tab.el.remove();
    this.docs.splice(i, 1);
    kvDel('draft:' + tab.id);
    if (!this.docs.length) this.createTab();
    else this.activate(this.docs[Math.max(0, i - 1)]);
  }

  replaceBook(tab: BookTab, book: Workbook): void {
    tab.book = book;
    book.onChange = () => this.bookChanged(tab);
    tab.engine = new Engine(book);
    tab.cf = new CFEvaluator(tab.engine);
    tab.engine.rebuild();
    tab.engine.recalc(true);
    tab.grid.layout();
    this.refresh();
  }

  /** Called after every transaction, undo and redo. */
  bookChanged(tab: BookTab): void {
    // In manual mode the engine evaluates only the edited formulas; dependents wait for F9.
    tab.engine.recalc();
    tab.cf.invalidate();
    if (!tab.dirty) { tab.dirty = true; this.ui.tabsChanged(); document.title = `${tab.displayName} • – Anne\'s Grid`; }
    this.scheduleAutosave();
    tab.grid.invalidate();
    this.refresh();
  }

  /** Recalculates now (F9). full = rebuild dependency graph (Ctrl+Alt+F9). */
  calculate(full = false): void {
    const t = this.tab;
    if (full) t.book.pending.structural = true;
    t.engine.recalc(true);
    t.cf.invalidate();
    t.grid.invalidate();
    this.refresh();
  }

  refresh(): void {
    if (this.refreshQueued) return;
    this.refreshQueued = true;
    requestAnimationFrame(() => {
      this.refreshQueued = false;
      if (!this.active) return;
      this.ui.refresh();
      for (const l of this.listeners) l(this);
    });
  }

  // ------------------------------------------------------------------------------------------
  // Commands

  handleKey(key: string): boolean {
    const id = this.keyIndex.get(canonicalKey(key));
    if (!id) return false;
    const c = getCommand(id);
    if (!c) return false;
    if (c.isEnabled && !c.isEnabled(this)) return false;
    this.exec(id);
    return true;
  }

  isReadOnly(): boolean {
    return this.tab.readOnly || this.book.finalMarked;
  }

  exec(id: string, arg?: any): unknown {
    const c = getCommand(id);
    if (!c) { toast(`Unknown command: ${id}`); return false; }
    if (c.edits && this.isReadOnly()) { toast('This workbook is read-only.'); return false; }
    if (c.isEnabled && !c.isEnabled(this)) { toast(`${c.label} is not available here.`); return false; }
    if (c.edits !== false && this.grid?.editing && !c.noRecord && !id.startsWith('cell.') && !id.startsWith('edit.formula') && !this.grid.commitEdit('none')) return false;
    if (this.recording && !c.noRecord && id !== 'dev.recordMacro' && id !== 'dev.stopRecording') this.recordStep(id, arg);
    if (!c.noRecord && id !== 'edit.repeat') this.lastCommand = { id, arg };
    try {
      const r = c.run(this, arg);
      if (r instanceof Promise) r.catch((e) => this.reportError(c, e)).finally(() => this.refresh());
      this.refresh();
      return r;
    } catch (e) {
      this.reportError(c, e);
      return false;
    }
  }

  private reportError(c: Command, e: unknown): void {
    console.error(e);
    toast(`${c.label}: ${(e as any)?.message || e}`);
  }

  recordStep(id: string, arg?: any): void {
    if (!this.recording || this.recordingPaused) return;
    const steps = this.recording.steps;
    const last = steps[steps.length - 1];
    if (id === 'nav.select' && last?.id === 'nav.select') { last.arg = arg; return; }
    steps.push(arg === undefined ? { id } : { id, arg: JSON.parse(JSON.stringify(arg)) });
  }

  // ------------------------------------------------------------------------------------------
  // Autosave and drafts

  scheduleAutosave(): void {
    if (!this.prefs.autosave || this.autosaveTimer) return;
    this.autosaveTimer = setTimeout(() => { this.autosaveTimer = null; void this.saveDrafts(); }, Math.max(3, this.prefs.autosaveSeconds) * 1000);
  }

  async saveDrafts(): Promise<void> {
    for (const tab of this.docs) {
      if (!tab.dirty) continue;
      try {
        await kvSet('draft:' + tab.id, { id: tab.id, json: tab.book.toJSON(), title: tab.title, fileName: tab.fileName, format: tab.format, uri: tab.uri, key: tab.key, time: Date.now() });
      } catch { /* storage full */ }
    }
    this.ui.status?.('Draft saved');
  }

  followLink(href: string): void {
    if (href.startsWith('#')) { this.exec('nav.goto', href.slice(1)); return; }
    if (!isSafeExternalUrl(href)) { toast(`Blocked a link with an unsupported or unsafe address: ${href.slice(0, 60)}`); return; }
    import('./platform/native').then(({ openExternal }) => openExternal(href));
  }

  focus(): void { this.active?.grid.focus(); }
}
