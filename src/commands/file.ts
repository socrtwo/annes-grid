/** File (backstage) commands: new, open, save, export, print, share, info, versions, options. */
import { def, commandCount } from './registry';
import type { App, BookTab } from '../app';
import { Workbook } from '../model/workbook';
import { FORMATS, formatById, formatForName, importFile, exportFile, PasswordError, sniff, type SniffResult } from '../io';
import { article } from '../io/sniff';
import { openDocument, saveAs, writeBack, printDocument, shareFile, readUri, openExternal, OPEN_ACCEPT, OPEN_MIMES, openInQuillApp, isIOS, type QuillApp } from '../platform/native';
import { openDialog, toast, confirmDialog, choiceDialog, alertDialog, promptDialog } from '../ui/dialog';
import { kvGet, kvSet, kvKeys, kvDel } from '../util/store';
import { TEMPLATES } from '../templates';
import { h } from '../ui/dom';
import { showRecoveryReport, showManualRecovery, recoverySummary } from '../ui/recovery';
import { sanitizePrefs, DEFAULT_REPLACEMENTS } from '../prefs';
import { SPELL_DIALECTS } from '../proof/proofing';

/** Key for a workbook's cached copy, Recent entry, snapshots and AI notes. */
export function docKey(tab: BookTab): string {
  return (tab.uri && !tab.uri.startsWith('fsa:') ? tab.uri : '') || tab.key || `doc:${tab.id}`;
}

export function localKey(name: string, data: ArrayBuffer): string {
  const b = new Uint8Array(data);
  let hh = 0x811c9dc5;
  for (let i = 0; i < b.length; i++) hh = Math.imul(hh ^ b[i], 0x01000193);
  return `doc:${name}:${b.length.toString(36)}.${(hh >>> 0).toString(36)}`;
}

async function knownKey(app: App, name: string, data: ArrayBuffer): Promise<string | null> {
  const b = new Uint8Array(data);
  for (const r of app.prefs.recent) {
    if (r.name !== name || !r.key.startsWith('doc:')) continue;
    const c = await kvGet('file:' + r.key);
    const d = c?.data instanceof ArrayBuffer ? new Uint8Array(c.data) : null;
    if (d && d.length === b.length && d.every((x, i) => x === b[i])) return r.key;
  }
  return null;
}

/** An untouched new workbook can be replaced by the next file opened. */
function isPristine(tab: BookTab): boolean {
  return !tab.dirty && !tab.uri && tab.book.sheets.length <= 1 && tab.book.sheet.cells.size === 0 && !tab.book.undoStack.length;
}

const SIBLING: Record<string, { app: QuillApp; what: string }> = {
  document: { app: 'Anne\'s Pen', what: 'its tables (one sheet each) and its text' },
  presentation: { app: 'Anne\'s Stage', what: 'a sheet per slide with its text and tables' },
  email: { app: 'Anne\'s Post', what: 'the message headers and body, and its spreadsheet and CSV attachments as sheets' },
};

/**
 * A document, presentation or message was opened: offer the Anne's Office app made for it, and always
 * offer to import it here. Resolves 'import', 'opened' (handed to the other app) or 'cancel'.
 */
async function chooseForeignOpen(sn: SniffResult, data: ArrayBuffer, name: string): Promise<'import' | 'opened' | 'cancel'> {
  const sib = SIBLING[sn.category];
  if (!sib) return 'import';
  const short = name.replace(/^.*[\\/]/, '');
  const canImport = !!formatById(sn.format)?.canImport;
  const choices = [`Open in ${sib.app}`, ...(canImport ? ['Import into Anne\'s Grid'] : [])];
  const c = await choiceDialog(`Open ${sn.label}`, `${short} is ${article(sn.label)} ${sn.label}. ${sib.app} is the Anne\'s Office app for it.${canImport ? ` Anne\'s Grid can also import ${sib.what} into a new workbook.` : ''}`, choices);
  if (c === 1) return 'import';
  if (c !== 0) return 'cancel';
  const mime = formatById(sn.format)?.mime || 'application/octet-stream';
  const r = await openInQuillApp(sib.app, data, short, mime);
  if (r.opened) { toast(`Opened ${short} in ${sib.app}`); return 'opened'; }
  const why = r.reason === 'unsupported'
    ? (isIOS() ? `On iPhone and iPad, apps cannot hand files to each other. Open ${short} from the Files app in ${sib.app}.` : `A web page cannot hand files to other apps. Open ${short} in ${sib.app} yourself (install it from the Anne\'s Office suite if you do not have it).`)
    : r.reason === 'not-installed' ? `${sib.app} does not seem to be installed${r.message ? ` (${r.message})` : ''}. Install it from the Anne\'s Office suite to open ${short} there.`
      : `${short} could not be handed to ${sib.app}${r.message ? `: ${r.message}` : '.'}`;
  if (!canImport) { await alertDialog(`${sib.app} is not available`, why); return 'cancel'; }
  const again = await choiceDialog(`${sib.app} is not available`, `${why} You can import it into Anne\'s Grid instead.`, ['Import into Anne\'s Grid']);
  return again === 0 ? 'import' : 'cancel';
}

export async function openBuffer(app: App, data: ArrayBuffer, name: string, uri: string | null, opts: { repair?: boolean; key?: string; password?: string; importDocuments?: boolean } = {}): Promise<BookTab | null> {
  try {
    if (opts.repair) toast(`Repairing ${name}…`);
    if (!opts.repair && !opts.password && !opts.importDocuments) {
      const sn = sniff(data, name);
      if (sn.category === 'document' || sn.category === 'presentation' || sn.category === 'email') {
        const how = await chooseForeignOpen(sn, data, name);
        if (how !== 'import') return null;
      }
    }
    let res;
    try {
      res = await importFile(data, name, { repair: opts.repair, password: opts.password });
    } catch (e) {
      if (e instanceof PasswordError) {
        const pw = await promptDialog('Password', `${name} is encrypted. Enter its password:`, '', { type: 'password' });
        if (pw == null) return null;
        return openBuffer(app, data, name, uri, { ...opts, password: pw });
      }
      throw e;
    }
    const { book, format, recovery, warnings, note } = res;
    const fmt = formatById(format) || formatForName(name);
    const cur = app.tab;
    const base = name.replace(/\.[^.]+$/, '');
    const canKeep = !!fmt?.canExport && !['csv', 'tsv', 'txt', 'html', 'md'].includes(fmt.id);
    const meta = recovery
      ? { title: `${base} (Recovered)`, fileName: `${base} (Recovered).xlsx`, format: 'xlsx', uri: null, key: '' }
      : { title: base, fileName: name, format: fmt?.canExport ? fmt.id : 'xlsx', uri: canKeep || ['csv', 'tsv', 'txt'].includes(fmt?.id || '') ? uri : null, key: uri && !uri.startsWith('fsa:') ? '' : opts.key || (await knownKey(app, name, data)) || localKey(name, data) };
    let tab: BookTab;
    if (cur && isPristine(cur)) {
      Object.assign(cur, meta);
      app.replaceBook(cur, book);
      cur.dirty = false;
      tab = cur;
      app.ui.tabsChanged();
    } else tab = app.createTab(book, meta);
    tab.dirty = false;
    tab.recovery = recovery || null;
    if (recovery) tab.dirty = true;
    else {
      addRecent(app, tab);
      await kvSet('file:' + docKey(tab), { name, data, uri, when: Date.now() });
    }
    app.activate(tab);
    const lead = note ? `${note} ` : '';
    const details = () => alertDialog('Import Notes', [note, ...(warnings || [])].filter(Boolean).join('\n'));
    if (recovery) toast(`${name} was damaged. ${recoverySummary(recovery)}`, { action: 'Details', onAction: () => app.exec('file.recoveryReport'), ms: 9000 });
    else if (warnings?.length) toast(`${lead}Opened ${name} (${warnings.length} note${warnings.length > 1 ? 's' : ''})`, { action: 'Details', onAction: details, ms: 9000 });
    else if (format === 'xls' || (fmt && !fmt.canExport)) toast(`${lead}Opened ${name}. Save as .xlsx to keep your changes.`, { ms: note ? 9000 : 6000 });
    else if (note) toast(`${lead}`, { ms: 8000 });
    else toast(`Opened ${name}`);
    return tab;
  } catch (e: any) {
    console.error(e);
    toast(`Could not open ${name}: ${e?.message || e}`, { action: 'Repair', onAction: () => openBuffer(app, data, name, uri, { repair: true }), ms: 9000 });
    return null;
  }
}

function addRecent(app: App, tab: BookTab) {
  const key = docKey(tab);
  const list = app.prefs.recent.filter((r) => r.key !== key);
  const old = app.prefs.recent.find((r) => r.key === key);
  list.unshift({ key, name: tab.displayName, when: Date.now(), uri: tab.uri || undefined, format: tab.format, pinned: old?.pinned });
  app.prefs.recent = [...list.filter((r) => r.pinned), ...list.filter((r) => !r.pinned)].slice(0, 40);
  app.savePrefs();
}

async function snapshot(app: App, tab: BookTab, label: string) {
  await kvSet(`snap:${docKey(tab)}:${Date.now()}`, { label, json: tab.book.toJSON(), when: Date.now(), name: tab.displayName });
}

export async function saveTab(app: App, tab: BookTab, opts: { as?: boolean; format?: string; copy?: boolean } = {}): Promise<boolean> {
  if (tab.grid.editing) tab.grid.commitEdit('none');
  let format = opts.format || (formatById(tab.format)?.canExport ? tab.format : app.prefs.defaultSaveFormat);
  let fmt = formatById(format);
  if (!fmt?.canExport) { format = 'xlsx'; fmt = formatById('xlsx')!; }
  // Single-sheet text formats lose sheets, formulas and formatting.
  if (!opts.copy && ['csv', 'tsv', 'txt', 'html', 'md'].includes(format)) {
    const lossy = tab.book.sheets.length > 1 || [...tab.book.sheet.cells.values()].some((c) => c.f != null || c.s);
    if (lossy) {
      const c = await choiceDialog('Keep Current Format?', `${fmt.label} keeps only the values of the active sheet. Formulas, formatting and other sheets will be lost.`, ['Save as Excel Workbook (.xlsx)', `Keep ${fmt.label}`]);
      if (c < 0) return false;
      if (c === 0) { format = 'xlsx'; fmt = formatById('xlsx')!; opts = { ...opts, as: true }; }
    }
  }
  if (tab.book.calc.recalcBeforeSave && tab.book.calc.mode === 'manual') tab.engine.recalc(true);
  const base = (tab.fileName || tab.title || 'Book').replace(/\.[^.]+$/, '');
  const name = `${base}.${fmt.ext[0]}`;
  tab.book.props.lastModifiedBy = app.prefs.author;
  const version = tab.book.version;
  let password: string | undefined;
  if (format === 'qgwe') {
    const r = await openDialog({ title: 'Encrypt Workbook', fields: [{ key: 'p1', label: 'Password', type: 'password' }, { key: 'p2', label: 'Confirm', type: 'password' }], validate: (x) => (!x.p1 ? 'Enter a password.' : x.p1 !== x.p2 ? 'Passwords do not match.' : null) });
    if (!r) return false;
    password = r.p1;
  }
  const blob = await exportFile(tab.book, format, { sheetIndex: tab.book.activeSheet, password });
  if (!opts.as && !opts.copy && tab.uri && format === tab.format) {
    try {
      if (await writeBack(tab.uri, blob)) { markSaved(app, tab, version); toast(`Saved ${tab.displayName}`); await snapshot(app, tab, 'Saved'); return true; }
    } catch (e) { console.warn(e); }
  }
  const r = await saveAs(blob, name);
  if (!r) return false;
  if (!opts.copy) {
    tab.fileName = r.name; tab.title = r.name.replace(/\.[^.]+$/, ''); tab.format = formatForName(r.name)?.id || format; tab.uri = r.uri;
    markSaved(app, tab, version);
    addRecent(app, tab);
    await kvSet('file:' + docKey(tab), { name: r.name, data: await blob.arrayBuffer(), uri: r.uri, when: Date.now() });
    await snapshot(app, tab, 'Saved');
  }
  toast(`Saved ${r.name}`);
  return true;
}

function markSaved(app: App, tab: BookTab, version: number) {
  tab.saved = Date.now();
  if (tab.book.version === version) { tab.dirty = false; kvDel('draft:' + tab.id); } else app.scheduleAutosave();
  app.ui.tabsChanged();
  document.title = `${tab.displayName}${tab.dirty ? ' •' : ''} – Anne\'s Grid`;
}

def([
  {
    id: 'file.new', label: 'New Workbook', icon: 'note_add', tab: 'Backstage', group: 'New', keys: ['Mod-n'], popular: true, noRecord: true,
    run: (app, arg) => {
      const t = TEMPLATES.find((x) => x.id === arg);
      const prev = app.tab && isPristine(app.tab) ? app.tab : null;
      if (t && t.id !== 'blank') {
        const book = t.build();
        book.props.author = app.prefs.author;
        book.clearHistory();
        const tab = app.createTab(book, { title: t.name });
        if (prev) app.closeTab(prev);
        app.activate(tab);
        app.ui.hideBackstage();
        return;
      }
      if (prev) { app.ui.hideBackstage(); app.focus(); return; }
      app.createTab();
      app.ui.hideBackstage();
    },
  },
  { id: 'file.newFromTemplate', label: 'New from Template…', icon: 'library_books', tab: 'Backstage', group: 'New', noRecord: true, run: (app) => app.ui.showBackstage('new') },
  {
    id: 'file.open', label: 'Open…', icon: 'folder_open', tab: 'Backstage', group: 'Open', keys: ['Mod-o', 'Mod-F12'], popular: true, noRecord: true,
    run: async (app) => {
      const f = await openDocument(OPEN_ACCEPT, OPEN_MIMES);
      if (!f) return;
      await openBuffer(app, f.data, f.name, f.uri);
      app.ui.hideBackstage();
    },
  },
  {
    id: 'file.openDropped', label: 'Open Dropped Files', icon: 'file_open', tab: 'Backstage', group: 'Open', internal: true, noRecord: true,
    run: async (app, files: File[]) => { for (const f of files || []) await openBuffer(app, await f.arrayBuffer(), f.name, null); },
  },
  {
    id: 'file.openRepair', label: 'Open and Repair…', icon: 'healing', tab: 'Backstage', group: 'Open', noRecord: true,
    tip: 'Open a damaged workbook: rebuild its package from raw bytes, repair its XML, salvage cell data, and show what was recovered',
    keywords: 'corrupt damaged recover fix broken salvage',
    run: async (app) => {
      const f = await openDocument('*/*', ['*/*']);
      if (!f) return;
      const tab = await openBuffer(app, f.data, f.name, f.uri, { repair: true });
      app.ui.hideBackstage();
      if (tab?.recovery) await showRecoveryReport(tab.recovery, { open: () => app.activate(tab) } as any);
    },
  },
  {
    id: 'file.recoverText', label: 'Recover Data from Any File…', icon: 'text_snippet', tab: 'Backstage', group: 'Open', noRecord: true,
    tip: 'Extract readable values and text from any file, however damaged',
    keywords: 'corrupt salvage strings extract',
    run: async () => { await showManualRecovery(); },
  },
  {
    id: 'file.recoveryReport', label: 'Recovery Report', icon: 'fact_check', tab: 'Backstage', group: 'Open', noRecord: true,
    isEnabled: (app) => !!app.tab?.recovery,
    run: async (app) => {
      const r = app.tab?.recovery;
      if (!r) { toast('This workbook was not recovered from a damaged file.'); return; }
      await showRecoveryReport(r, {} as any);
    },
  },
  { id: 'file.recoveryHelp', label: 'Manual Recovery Options', icon: 'support', tab: 'Backstage', group: 'Open', noRecord: true, run: () => showManualRecovery() },
  {
    id: 'file.openRecent', label: 'Open Recent', icon: 'history', tab: 'Backstage', group: 'Open', kind: 'menu', noRecord: true,
    run: async (app, key) => {
      if (!key) return app.ui.showBackstage('open');
      const rec = app.prefs.recent.find((r) => r.key === key);
      if (!rec) return;
      const existing = app.docs.find((d) => docKey(d) === key);
      if (existing) { app.activate(existing); app.ui.hideBackstage(); return; }
      let f = rec.uri ? await readUri(rec.uri) : null;
      if (!f) { const cached = await kvGet('file:' + key); if (cached) f = { name: cached.name, data: cached.data, uri: cached.uri }; }
      if (!f) { toast('That file is no longer available. Use Open to locate it.'); return; }
      await openBuffer(app, f.data, f.name, f.uri, { key });
      app.ui.hideBackstage();
    },
    menu: (app) => app.prefs.recent.slice(0, 12).map((r) => ({ id: 'file.openRecent', arg: r.key, label: (r.pinned ? '📌 ' : '') + r.name })),
  },
  { id: 'file.clearRecent', label: 'Clear Unpinned Recent Workbooks', icon: 'delete_sweep', tab: 'Backstage', group: 'Open', noRecord: true, run: (app) => { app.prefs.recent = app.prefs.recent.filter((r) => r.pinned); app.savePrefs(); app.ui.showBackstage('open'); } },
  { id: 'file.save', label: 'Save', icon: 'save', tab: 'Backstage', group: 'Save', keys: ['Mod-s', 'Shift-F12'], popular: true, noRecord: true, run: (app) => saveTab(app, app.tab) },
  { id: 'file.saveAs', label: 'Save As…', icon: 'save_as', tab: 'Backstage', group: 'Save', keys: ['F12'], popular: true, noRecord: true, run: (app) => app.ui.showBackstage('saveas') },
  { id: 'file.saveAll', label: 'Save All', icon: 'library_add_check', tab: 'Backstage', group: 'Save', noRecord: true, run: async (app) => { for (const d of app.docs) if (d.dirty) await saveTab(app, d); } },
  {
    id: 'file.saveAsFormat', label: 'Save As Format', icon: 'save_as', tab: 'Backstage', group: 'Save', internal: true, noRecord: true,
    run: async (app, format: string) => { await saveTab(app, app.tab, { as: true, format }); app.ui.hideBackstage(); },
  },
  {
    id: 'file.exportAs', label: 'Export a Copy', icon: 'ios_share', tab: 'Backstage', group: 'Export', internal: true, noRecord: true,
    run: async (app, format: string) => { await saveTab(app, app.tab, { copy: true, format }); },
  },
  {
    id: 'file.saveTemplate', label: 'Save as Template', icon: 'bookmark_add', tab: 'Backstage', group: 'Save', noRecord: true,
    run: async (app) => {
      const name = await promptDialog('Save as Template', 'Template name', app.tab.title);
      if (!name) return;
      await kvSet(`tpl:${name}`, { name, json: app.book.toJSON(), when: Date.now() });
      toast(`Template "${name}" saved`);
    },
  },
  { id: 'file.autosave', label: 'AutoSave', icon: 'autorenew', tab: 'Backstage', group: 'Save', kind: 'toggle', noRecord: true, run: (app) => { app.prefs.autosave = !app.prefs.autosave; app.savePrefs(); toast(app.prefs.autosave ? 'AutoSave on (drafts saved automatically)' : 'AutoSave off'); }, isActive: (app) => app.prefs.autosave },
]);

for (const f of FORMATS.filter((x) => x.canExport)) {
  def({ id: `file.export.${f.id}`, label: `Export as ${f.label}`, icon: 'ios_share', tab: 'Backstage', group: 'Export', noRecord: true, run: (app) => saveTab(app, app.tab, { copy: true, format: f.id }) });
}

def([
  {
    id: 'file.print', label: 'Print…', icon: 'print', tab: 'Backstage', group: 'Print', keys: ['Mod-p'], popular: true, noRecord: true,
    run: async (app) => {
      app.ui.hideBackstage();
      const { preparePrint } = await import('../print/print');
      const cleanup = await preparePrint(app);
      // printDocument resolves when the print UI has closed; the print pages must stay until then,
      // because the print UI lays the page out again each time its preview refreshes.
      try { await printDocument(app.tab.displayName); } finally { setTimeout(cleanup, 300); }
    },
  },
  { id: 'file.quickPrint', label: 'Quick Print', icon: 'print', tab: 'Backstage', group: 'Print', noRecord: true, run: (app) => app.exec('file.print') },
  { id: 'file.printPreview', label: 'Print Preview', icon: 'preview', tab: 'Backstage', group: 'Print', keys: ['Mod-F2'], noRecord: true, run: async (app) => { app.ui.hideBackstage(); const { showPrintPreview } = await import('../print/print'); await showPrintPreview(app); } },
  { id: 'file.exportPdf', label: 'Export as PDF', icon: 'picture_as_pdf', tab: 'Backstage', group: 'Export', popular: true, noRecord: true, run: (app) => { toast('In the print dialog choose "Save as PDF".'); app.exec('file.print'); } },
  {
    id: 'file.share', label: 'Share…', icon: 'share', tab: 'Backstage', group: 'Share', popular: true, noRecord: true,
    run: async (app) => {
      const r = await openDialog({ title: 'Share', fields: [{ key: 'fmt', label: 'Send as', type: 'radio', options: [...FORMATS.filter((f) => f.canExport && f.id !== 'qgwe').map((f) => ({ value: f.id, label: f.label })), { value: 'pdf', label: 'PDF (via print service)' }], value: 'xlsx' }] });
      if (!r) return;
      if (r.fmt === 'pdf') return app.exec('file.exportPdf');
      const fmt = formatById(r.fmt)!;
      const name = `${(app.tab.fileName || app.tab.title).replace(/\.[^.]+$/, '')}.${fmt.ext[0]}`;
      const blob = await exportFile(app.book, r.fmt, { sheetIndex: app.book.activeSheet });
      await shareFile(blob, name, app.tab.title);
    },
  },
  { id: 'file.email', label: 'Email as Attachment', icon: 'attach_email', tab: 'Backstage', group: 'Share', noRecord: true, run: (app) => app.exec('file.share') },
  {
    id: 'file.copyMarkdown', label: 'Copy Sheet as Markdown', icon: 'markdown', tab: 'Backstage', group: 'Share', noRecord: true,
    run: async (app) => { const { sheetToMarkdown } = await import('../io/html'); await navigator.clipboard?.writeText(sheetToMarkdown(app.sheet)); toast('Markdown table copied'); },
  },
  {
    id: 'file.close', label: 'Close Workbook', icon: 'close', tab: 'Backstage', group: 'Close', keys: ['Mod-w', 'Mod-F4'], popular: true, noRecord: true,
    run: async (app, arg) => {
      const tab = app.docs.find((d) => d.id === arg) || app.tab;
      if (tab.dirty && app.prefs.confirmClose) {
        const c = await choiceDialog('Save Changes?', `Save changes to "${tab.displayName}" before closing?`, ['Save', "Don't Save"]);
        if (c < 0) return;
        if (c === 0 && !(await saveTab(app, tab))) return;
      }
      app.closeTab(tab);
      app.ui.hideBackstage();
    },
  },
  { id: 'file.closeAll', label: 'Close All Workbooks', icon: 'tab_close', tab: 'Backstage', group: 'Close', noRecord: true, run: async (app) => { for (const d of [...app.docs]) { app.activate(d); await app.exec('file.close', d.id); } } },
  {
    id: 'file.revert', label: 'Revert to Saved', icon: 'settings_backup_restore', tab: 'Backstage', group: 'Info', noRecord: true,
    run: async (app) => {
      const tab = app.tab;
      const cached = await kvGet('file:' + docKey(tab));
      if (!cached) { toast('No saved version to revert to.'); return; }
      if (!(await confirmDialog('Revert', 'Discard all changes since the last save?'))) return;
      const { book } = await importFile(cached.data, cached.name);
      app.replaceBook(tab, book); tab.dirty = false; app.ui.tabsChanged();
    },
  },
  { id: 'file.exit', label: 'Exit', icon: 'logout', tab: 'Backstage', group: 'Close', noRecord: true, run: async (app) => { await app.saveDrafts(); const { App: CapApp } = await import('@capacitor/app'); try { await CapApp.exitApp(); } catch { window.close(); } } },
]);

// Info ---------------------------------------------------------------------------------------
def([
  {
    id: 'file.properties', label: 'Properties…', icon: 'info', tab: 'Backstage', group: 'Info', noRecord: true,
    run: async (app) => {
      const p = app.book.props;
      const r = await openDialog({ title: 'Workbook Properties', wide: true, fields: [
        { key: 'title', label: 'Title', type: 'text', value: p.title }, { key: 'subject', label: 'Subject', type: 'text', value: p.subject },
        { key: 'author', label: 'Author', type: 'text', value: p.author || app.prefs.author, half: true }, { key: 'manager', label: 'Manager', type: 'text', value: p.manager, half: true },
        { key: 'company', label: 'Company', type: 'text', value: p.company, half: true }, { key: 'category', label: 'Category', type: 'text', value: p.category, half: true },
        { key: 'keywords', label: 'Tags / keywords', type: 'text', value: p.keywords }, { key: 'status', label: 'Status', type: 'text', value: p.status },
        { key: 'comments', label: 'Comments', type: 'textarea', rows: 3, value: p.comments },
        { key: 'custom', label: 'Custom properties (name=value per line)', type: 'textarea', rows: 3, value: Object.entries(p.custom || {}).map(([k, x]) => `${k}=${x}`).join('\n') },
        { key: 'stats', type: 'info', html: `<p class="ql-help">Created: ${new Date(p.created).toLocaleString()} · Modified: ${new Date(p.modified).toLocaleString()} · Last saved: ${app.tab.saved ? new Date(app.tab.saved).toLocaleString() : '—'}</p>` },
      ] });
      if (!r) return;
      app.book.transact('Properties', () => {
        app.book.recordBook();
        Object.assign(app.book.props, { title: r.title, subject: r.subject, author: r.author, manager: r.manager, company: r.company, category: r.category, keywords: r.keywords, status: r.status, comments: r.comments });
        app.book.props.custom = Object.fromEntries(String(r.custom).split('\n').map((l) => l.split('=')).filter((x) => x[0]?.trim()).map(([k, ...rest]) => [k.trim(), rest.join('=').trim()]));
      });
    },
  },
  {
    id: 'file.inspect', label: 'Inspect Workbook…', icon: 'policy', tab: 'Backstage', group: 'Info', noRecord: true,
    run: async (app) => {
      const book = app.book;
      let comments = 0, notes = 0, hiddenRows = 0, hiddenCols = 0, links = 0;
      for (const s of book.sheets) { for (const c of s.comments.values()) c.note ? notes++ : comments++; hiddenRows += s.hiddenRows.size; hiddenCols += s.hiddenCols.size; for (const c of s.cells.values()) if (c.link) links++; }
      const hiddenSheets = book.sheets.filter((s) => s.visibility !== 'visible').length;
      const r = await openDialog({ title: 'Document Inspector', help: 'Select the content to remove before sharing. This cannot be undone with Undo for workbook properties.', fields: [
        { key: 'comments', label: `Comments and notes (${comments + notes})`, type: 'checkbox', value: comments + notes > 0 },
        { key: 'props', label: 'Document properties and personal information', type: 'checkbox', value: true },
        { key: 'hiddenSheets', label: `Hidden worksheets (${hiddenSheets})`, type: 'checkbox', value: false },
        { key: 'hidden', label: `Hidden rows and columns (${hiddenRows + hiddenCols}) — unhide`, type: 'checkbox', value: false },
        { key: 'links', label: `Hyperlinks (${links})`, type: 'checkbox', value: false },
        { key: 'hf', label: 'Headers and footers', type: 'checkbox', value: false },
        { key: 'macros', label: `Macro modules (${book.modules.length})`, type: 'checkbox', value: false },
      ], ok: 'Remove' });
      if (!r) return;
      book.transact('Inspect Workbook', () => {
        book.recordBook();
        for (const s of book.sheets) {
          if (r.comments) s.comments.clear();
          if (r.hidden) { s.hiddenRows.clear(); s.hiddenCols.clear(); }
          if (r.links) for (const [k, c] of s.cells) if (c.link) { const { link: _l, ...rest } = c; void _l; s.cells.set(k, rest); }
          if (r.hf) { s.print.header = { left: '', center: '', right: '' }; s.print.footer = { left: '', center: '', right: '' }; }
        }
        if (r.hiddenSheets) book.sheets = book.sheets.filter((s) => s.visibility === 'visible');
        book.activeSheet = Math.min(book.activeSheet, book.sheets.length - 1);
        if (r.props) { const keep = book.props.title; Object.assign(book.props, { subject: '', author: '', manager: '', company: '', category: '', keywords: '', comments: '', status: '', lastModifiedBy: '', custom: {} }); book.props.title = keep; }
        if (r.macros) book.modules = [];
      });
      toast('Workbook inspected and cleaned');
    },
  },
  { id: 'file.encrypt', label: 'Encrypt with Password…', icon: 'enhanced_encryption', tab: 'Backstage', group: 'Info', noRecord: true, run: (app) => saveTab(app, app.tab, { copy: true, format: 'qgwe' }) },
  {
    id: 'file.openEncrypted', label: 'Open Encrypted Workbook…', icon: 'lock_open', tab: 'Backstage', group: 'Open', noRecord: true,
    run: async (app) => {
      const f = await openDocument('.qgwe', ['application/octet-stream', '*/*']);
      if (!f) return;
      await openBuffer(app, f.data, f.name, null);
    },
  },
  { id: 'file.versions', label: 'Version History', icon: 'history', tab: 'Backstage', group: 'Info', popular: true, noRecord: true, origin: 'Snapshots (Scrivener) / version history (online editors)', run: (app) => { app.ui.hideBackstage(); app.ui.openPane('versions'); } },
  { id: 'file.snapshot', label: 'Take Snapshot', icon: 'add_a_photo', tab: 'Backstage', group: 'Info', keys: ['Mod-Alt-5'], noRecord: true, run: async (app) => { const label = await promptDialog('Take Snapshot', 'Snapshot name', `Snapshot ${new Date().toLocaleString()}`); if (label) { await snapshot(app, app.tab, label); toast('Snapshot saved'); } } },
  {
    id: 'file.recover', label: 'Recover Unsaved Workbooks', icon: 'restore_page', tab: 'Backstage', group: 'Info', noRecord: true,
    run: async (app) => {
      const keys = (await kvKeys('draft:')).filter((k) => !app.docs.some((d) => 'draft:' + d.id === k));
      if (!keys.length) { toast('No unsaved drafts found.'); return; }
      const drafts = await Promise.all(keys.map(async (k) => ({ k, d: await kvGet(k) })));
      const r = await openDialog({ title: 'Workbook Recovery', fields: [{ key: 'k', label: 'Available drafts', type: 'list', multiple: true, rows: 8, options: drafts.map(({ k, d }) => ({ value: k, label: `${d.fileName || d.title} — ${new Date(d.time).toLocaleString()}` })), value: keys }], buttons: [{ label: 'Discard All', onClick: async (_v, ctl) => { for (const k of keys) await kvDel(k); ctl.close(); toast('Drafts discarded'); } }], ok: 'Recover' });
      if (!r?.k?.length) return;
      for (const { k, d } of drafts.filter((x) => r.k.includes(x.k))) {
        try {
          const tab = app.createTab(Workbook.fromJSON(d.json), { title: d.title, fileName: d.fileName, format: d.format, uri: d.uri, key: d.key || '' });
          tab.dirty = true;
          await kvDel(k);
        } catch (e) { toast(`Could not restore ${d.title}: ${(e as Error).message}`); }
      }
    },
  },
]);

// Options ------------------------------------------------------------------------------------
def([
  {
    id: 'file.options', label: 'Options…', icon: 'settings', tab: 'Backstage', group: 'Options', popular: true, noRecord: true,
    run: async (app) => {
      const p = app.prefs;
      const calc = app.book.calc;
      const r = await openDialog({ title: 'Anne\'s Grid Options', wide: true, fields: [
        { key: 'h1', type: 'heading', label: 'General' },
        { key: 'author', label: 'User name', type: 'text', value: p.author, half: true },
        { key: 'initials', label: 'Initials', type: 'text', value: p.initials, half: true },
        { key: 'theme', label: 'App theme', type: 'select', options: ['system', 'light', 'dark'], value: p.theme, half: true },
        { key: 'darkSheet', label: 'Dark cells in dark theme', type: 'checkbox', value: p.darkSheet, half: true },
        { key: 'uiScale', label: 'Toolbar size', type: 'select', options: [{ value: '0.85', label: 'Compact' }, { value: '1', label: 'Normal' }, { value: '1.15', label: 'Large (touch)' }], value: String(p.uiScale), half: true },
        { key: 'newSheets', label: 'Sheets in a new workbook', type: 'number', min: 1, max: 255, value: p.newSheets, half: true },
        { key: 'defaultFont', label: 'Default font (new workbooks)', type: 'text', value: p.defaultFont || 'Carlito', half: true },
        { key: 'defaultSize', label: 'Default size', type: 'number', value: p.defaultSize || 11, half: true },
        { key: 'qatBottom', label: 'Quick Access Bar at the bottom (near keyboard)', type: 'checkbox', value: p.qatPosition === 'bottom', half: true },
        { key: 'qatLabels', label: 'Show labels on Quick Access Bar', type: 'checkbox', value: p.qatLabels, half: true },
        { key: 'h2', type: 'heading', label: 'Formulas' },
        { key: 'calcMode', label: 'Workbook calculation', type: 'select', options: [{ value: 'auto', label: 'Automatic' }, { value: 'autoNoTables', label: 'Automatic except data tables' }, { value: 'manual', label: 'Manual' }], value: calc.mode, half: true },
        { key: 'recalcBeforeSave', label: 'Recalculate before saving', type: 'checkbox', value: calc.recalcBeforeSave, half: true },
        { key: 'iterative', label: 'Enable iterative calculation', type: 'checkbox', value: calc.iterative, half: true },
        { key: 'maxIterations', label: 'Maximum iterations', type: 'number', value: calc.maxIterations, half: true },
        { key: 'maxChange', label: 'Maximum change', type: 'number', step: 0.0001, value: calc.maxChange, half: true },
        { key: 'date1904', label: 'Use 1904 date system', type: 'checkbox', value: calc.date1904, half: true },
        { key: 'r1c1', label: 'R1C1 column headings', type: 'checkbox', value: p.r1c1, half: true },
        { key: 'showFunctionTips', label: 'Formula AutoComplete and argument tips', type: 'checkbox', value: p.showFunctionTips, half: true },
        { key: 'h3', type: 'heading', label: 'Editing' },
        { key: 'enterMoves', label: 'After pressing Enter, move selection', type: 'checkbox', value: p.enterMoves, half: true },
        { key: 'enterDirection', label: 'Direction', type: 'select', options: ['down', 'right', 'up', 'left'], value: p.enterDirection, half: true },
        { key: 'fillHandle', label: 'Enable fill handle and cell drag-and-drop', type: 'checkbox', value: p.fillHandle, half: true },
        { key: 'editInCell', label: 'Allow editing directly in cells', type: 'checkbox', value: p.editInCell, half: true },
        { key: 'autoComplete', label: 'AutoComplete for cell values', type: 'checkbox', value: p.autoComplete, half: true },
        { key: 'autoPercent', label: 'Automatic percent entry', type: 'checkbox', value: p.autoPercent, half: true },
        { key: 'flashFill', label: 'Automatically Flash Fill', type: 'checkbox', value: p.flashFill, half: true },
        { key: 'fixedDecimals', label: 'Automatically insert a decimal point (places, 0 = off)', type: 'number', value: p.fixedDecimals, half: true },
        { key: 'dateOrder', label: 'Typed dates are', type: 'select', options: [{ value: 'mdy', label: 'Month/Day/Year' }, { value: 'dmy', label: 'Day/Month/Year' }, { value: 'ymd', label: 'Year-Month-Day' }], value: p.dateOrder, half: true },
        { key: 'zoomOnRoll', label: 'Zoom on roll with mouse wheel', type: 'checkbox', value: p.zoomOnRoll, half: true },
        { key: 'touchHandles', label: 'Touch selection handles', type: 'checkbox', value: p.touchHandles, half: true },
        { key: 'h4', type: 'heading', label: 'Save and Proofing' },
        { key: 'autosave', label: 'AutoSave drafts', type: 'checkbox', value: p.autosave, half: true },
        { key: 'autosaveSeconds', label: 'Every (seconds)', type: 'number', value: p.autosaveSeconds, half: true },
        { key: 'defaultSaveFormat', label: 'Default save format', type: 'select', options: FORMATS.filter((f) => f.canExport && f.id !== 'qgwe').map((f) => ({ value: f.id, label: f.label })), value: p.defaultSaveFormat, half: true },
        { key: 'confirmClose', label: 'Ask to save when closing', type: 'checkbox', value: p.confirmClose, half: true },
        { key: 'autoCorrect', label: 'AutoCorrect as you type', type: 'checkbox', value: p.autoCorrect, half: true },
        { key: 'spellDialect', label: 'English spelling', type: 'select', options: SPELL_DIALECTS, value: p.spellDialect || 'any', half: true },
        { key: 'h5', type: 'heading', label: 'Customize' },
        { key: 'bq', type: 'button', label: 'Customize Quick Access Bar…', onClick: () => app.exec('file.customizeQat') },
        { key: 'br', type: 'button', label: 'Customize Tabs (Ribbon)…', onClick: () => app.exec('file.customizeRibbon') },
        { key: 'bk', type: 'button', label: 'Customize Keyboard…', onClick: () => app.exec('file.customizeKeyboard') },
        { key: 'bl', type: 'button', label: 'Edit Custom Lists…', onClick: () => app.exec('file.customLists') },
        { key: 'ba', type: 'button', label: 'AutoCorrect Options…', onClick: () => app.exec('file.autocorrectOptions') },
        { key: 'bx', type: 'button', label: 'Export / Import All Settings…', onClick: () => app.exec('file.exportSettings') },
      ] });
      if (!r) return;
      Object.assign(p, {
        author: r.author || 'Author', initials: r.initials || (r.author || 'A')[0], theme: r.theme, darkSheet: r.darkSheet, uiScale: parseFloat(r.uiScale), newSheets: Math.max(1, Math.min(255, r.newSheets || 1)),
        defaultFont: r.defaultFont === 'Carlito' ? '' : r.defaultFont, defaultSize: r.defaultSize === 11 ? 0 : r.defaultSize,
        qatPosition: r.qatBottom ? 'bottom' : 'top', qatLabels: r.qatLabels, r1c1: r.r1c1, showFunctionTips: r.showFunctionTips,
        enterMoves: r.enterMoves, enterDirection: r.enterDirection, fillHandle: r.fillHandle, editInCell: r.editInCell, autoComplete: r.autoComplete, autoPercent: r.autoPercent, flashFill: r.flashFill,
        fixedDecimals: Math.max(0, r.fixedDecimals || 0), dateOrder: r.dateOrder, zoomOnRoll: r.zoomOnRoll, touchHandles: r.touchHandles,
        autosave: r.autosave, autosaveSeconds: Math.max(3, r.autosaveSeconds || 10), defaultSaveFormat: r.defaultSaveFormat, confirmClose: r.confirmClose, autoCorrect: r.autoCorrect, spellDialect: r.spellDialect,
      });
      app.savePrefs();
      const nc = { mode: r.calcMode, recalcBeforeSave: r.recalcBeforeSave, iterative: r.iterative, maxIterations: Math.max(1, r.maxIterations || 100), maxChange: Math.max(0, r.maxChange || 0.001), date1904: r.date1904 };
      if (JSON.stringify(nc) !== JSON.stringify({ mode: calc.mode, recalcBeforeSave: calc.recalcBeforeSave, iterative: calc.iterative, maxIterations: calc.maxIterations, maxChange: calc.maxChange, date1904: calc.date1904 })) {
        app.book.transact('Calculation Options', () => { app.book.recordBook(); Object.assign(app.book.calc, nc); });
        app.calculate(true);
      }
      app.ui.applyPrefs();
    },
  },
  { id: 'file.customizeQat', label: 'Customize Quick Access Bar…', icon: 'tune', tab: 'Backstage', group: 'Options', popular: true, noRecord: true, run: (app) => app.ui.customizeQat() },
  { id: 'file.customizeRibbon', label: 'Customize Tabs…', icon: 'view_quilt', tab: 'Backstage', group: 'Options', noRecord: true, run: (app) => app.ui.customizeRibbon() },
  { id: 'file.customizeKeyboard', label: 'Customize Keyboard…', icon: 'keyboard', tab: 'Backstage', group: 'Options', noRecord: true, run: (app) => app.ui.customizeKeyboard() },
  {
    id: 'file.customLists', label: 'Custom Lists…', icon: 'format_list_numbered', tab: 'Backstage', group: 'Options', noRecord: true,
    tip: 'Lists the fill handle continues (e.g. North, East, South, West)',
    run: async (app) => {
      const r = await openDialog({ title: 'Custom Lists', help: 'One list per line, items separated by commas. Built-in day and month lists are always available.', fields: [{ key: 'lists', label: 'Lists', type: 'textarea', rows: 10, value: app.prefs.customLists.map((l) => l.items.join(', ')).join('\n') }], buttons: [{ label: 'Import from Selection', onClick: (_v, ctl) => { const rg = app.sheet.sel.ranges[0]; const items: string[] = []; app.sheet.forEachIn(rg, (_r, _c, cell) => { if (typeof cell.v === 'string') items.push(cell.v); }); if (items.length) ctl.set('lists', (ctl.get('lists') ? ctl.get('lists') + '\n' : '') + items.join(', ')); } }] });
      if (!r) return;
      app.prefs.customLists = String(r.lists).split('\n').map((l) => ({ items: l.split(',').map((x) => x.trim()).filter(Boolean) })).filter((l) => l.items.length > 1);
      app.savePrefs();
    },
  },
  {
    id: 'file.autocorrectOptions', label: 'AutoCorrect Options…', icon: 'auto_fix_normal', tab: 'Backstage', group: 'Options', noRecord: true,
    run: async (app) => {
      const r = await openDialog({ title: 'AutoCorrect', wide: true, fields: [
        { key: 'on', label: 'Replace text as you type', type: 'checkbox', value: app.prefs.autoCorrect },
        { key: 'list', label: 'Replacements (one per line: typo => correction)', type: 'textarea', rows: 12, value: Object.entries(app.prefs.replacements).map(([k, x]) => `${k} => ${x}`).join('\n') },
      ], buttons: [{ label: 'Reset to Defaults', onClick: (_v, ctl) => ctl.set('list', Object.entries(DEFAULT_REPLACEMENTS).map(([k, x]) => `${k} => ${x}`).join('\n')) }] });
      if (!r) return;
      app.prefs.autoCorrect = r.on;
      app.prefs.replacements = Object.fromEntries(String(r.list).split('\n').map((l) => l.split('=>')).filter((x) => x.length === 2 && x[0].trim()).map(([k, x]) => [k.trim(), x.trim()]));
      app.savePrefs();
    },
  },
  {
    id: 'file.customDictionary', label: 'Custom Dictionary…', icon: 'book', tab: 'Backstage', group: 'Options', noRecord: true,
    run: async (app) => {
      const r = await openDialog({ title: 'Custom Dictionary', fields: [{ key: 'words', label: 'Words (one per line)', type: 'textarea', rows: 12, value: app.prefs.dictionary.join('\n') }, { key: 'ign', label: 'Ignored words (one per line)', type: 'textarea', rows: 4, value: app.prefs.ignoredWords.join('\n') }] });
      if (!r) return;
      app.prefs.dictionary = String(r.words).split('\n').map((w) => w.trim()).filter(Boolean);
      app.prefs.ignoredWords = String(r.ign).split('\n').map((w) => w.trim()).filter(Boolean);
      app.savePrefs();
    },
  },
  {
    id: 'file.exportSettings', label: 'Export / Import Settings…', icon: 'import_export', tab: 'Backstage', group: 'Options', noRecord: true,
    run: async (app) => {
      const r = await openDialog({ title: 'Settings Transfer', help: 'Copy this text to back up your Quick Access Bar, custom tabs, shortcuts, macros, custom lists, AutoCorrect and preferences. Paste settings text here and choose Import to restore.', fields: [{ key: 'json', label: 'Settings (JSON)', type: 'textarea', rows: 12, value: JSON.stringify({ ...app.prefs, recent: [] }, null, 1) }], ok: 'Import', buttons: [{ label: 'Copy', onClick: (x) => { navigator.clipboard?.writeText(x.json); toast('Copied'); } }] });
      if (!r) return;
      let p: unknown;
      try { p = JSON.parse(r.json); } catch { /* reported below */ }
      if (!p || typeof p !== 'object' || Array.isArray(p)) { toast('That is not valid settings JSON.'); return; }
      const before = JSON.parse(JSON.stringify(app.prefs));
      Object.assign(app.prefs, sanitizePrefs(p), { recent: app.prefs.recent });
      try { app.ui.applyPrefs(); app.ui.rebuild(); } catch (e: any) {
        Object.assign(app.prefs, before); app.ui.applyPrefs(); app.ui.rebuild();
        toast(`These settings could not be applied: ${e?.message || e}`);
        return;
      }
      app.savePrefs();
      toast('Settings imported');
    },
  },
  { id: 'file.resetCustomizations', label: 'Reset All Customizations', icon: 'restart_alt', tab: 'Backstage', group: 'Options', noRecord: true, run: async (app) => { if (!(await confirmDialog('Reset', 'Reset the Quick Access Bar, custom tabs and keyboard shortcuts to defaults?'))) return; const { DEFAULT_QAT } = await import('../prefs'); app.prefs.qat = [...DEFAULT_QAT]; app.prefs.customTabs = []; app.prefs.hiddenTabs = []; app.prefs.keymap = {}; app.savePrefs(); app.ui.rebuild(); } },
]);

// Help -------------------------------------------------------------------------------------------
def([
  { id: 'help.help', label: 'Help', icon: 'help', tab: 'Help', group: 'Help', keys: ['F1'], noRecord: true, run: () => openExternal('https://github.com/socrtwo/QuillGrid#readme') },
  { id: 'help.shortcuts', label: 'Keyboard Shortcuts', icon: 'keyboard', tab: 'Help', group: 'Help', keys: ['Mod-/'], noRecord: true, run: (app) => app.ui.showShortcuts() },
  { id: 'help.commandReference', label: 'Command Reference', icon: 'menu_book', tab: 'Help', group: 'Help', noRecord: true, run: (app) => app.ui.showCommandReference() },
  { id: 'help.tour', label: 'Show the Quick Tour', icon: 'tour', tab: 'Help', group: 'Help', noRecord: true, run: (app) => app.ui.tour() },
  {
    id: 'help.about', label: 'About Anne\'s Grid', icon: 'info', tab: 'Help', group: 'Help', noRecord: true,
    run: async () => {
      const { functionNames } = await import('../calc/functions');
      await alertDialog('About Anne\'s Grid', h('div', { class: 'ql-msg' },
        h('p', null, h('b', null, `Anne\'s Grid ${__APP_VERSION__}`), ' — an original, open-source spreadsheet (MIT License). A companion to Anne\'s Pen.'),
        h('p', null, `${commandCount()} commands and ${functionNames().length} worksheet functions. Every command is available to the Quick Access Bar, custom tabs, keyboard shortcuts, macros and command search.`),
        h('p', null, 'Anne\'s Grid is not affiliated with, endorsed by, or sponsored by Microsoft, Google, Apple, The Document Foundation, or any other vendor. Product names are used only to describe file-format compatibility or the origin of an idea.'),
        h('p', null, 'Open-source components: JSZip (MIT/GPLv3 dual, used under MIT), nspell and hunspell-asm (MIT) with SCOWL English dictionaries (MIT and BSD-style), Material Symbols (Apache 2.0), Capacitor (MIT), Electron (MIT), fonts Carlito, Caladea, Gelasio, EB Garamond, Courier Prime, Libre Baskerville, Lato, Merriweather, Open Sans (SIL OFL 1.1) and Tinos, Arimo, Cousine (Apache 2.0). Their licenses are in Help ▸ Third-Party Notices.')));
    },
  },
  {
    id: 'help.notices', label: 'Third-Party Notices', icon: 'license', tab: 'Help', group: 'Help', noRecord: true,
    run: async () => {
      let text = '';
      try { const r = await fetch('./third-party-notices.txt'); if (r.ok) text = await r.text(); } catch { /* offline web build without the file */ }
      await alertDialog('Third-Party Notices', h('pre', { class: 'ql-notices' }, text || 'The notices file is missing from this build. The licenses are listed in docs/LEGAL.md.'));
    },
  },
  { id: 'help.feedback', label: 'Report an Issue', icon: 'bug_report', tab: 'Help', group: 'Help', noRecord: true, run: () => openExternal('https://github.com/socrtwo/QuillGrid/issues') },
  { id: 'help.functionReference', label: 'Function Reference', icon: 'functions', tab: 'Help', group: 'Help', noRecord: true, run: (app) => app.exec('formulas.insertFunction', { browse: true }) },
]);

void writeBack;
