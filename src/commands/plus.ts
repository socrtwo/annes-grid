/**
 * Plus tab: tools in the manner of the other well-known spreadsheet add-ins, Kutools for Excel, Ablebits
 * Ultimate Suite, Power Tools for Google Sheets, XLTools and Microsoft's Inquire, re-implemented from
 * their public descriptions (nominative use; nothing is copied). Groups: View, Range, Combine, Convert,
 * Insert, Delete, Lookup, Filter, Print, Security, Workbook, Inquire, Query.
 */
import type { App } from '../app';
import { Sheet, Workbook, newId } from '../model/workbook';
import type { Cell, CellStyle, DataValidation, DrawingObject, Scalar } from '../model/types';
import { CellError } from '../model/types';
import { cellKey, keyRow, keyCol, colName, cellName, rangeName, parseRange, MAX_ROWS, MAX_COLS, type Range } from '../model/address';
import { selRanges, activeRange, activeCell, clipToUsed, setInput } from './ops';
import { defTools, externalLinkIn, P, selCells, selBlock, singleCell, hasContent, selectHits, sheetProtected, structureLocked, emptySheet, numOf, textOf, plural, report, pickSheets, sheetsByNames, shownText, argOrAsk, columnChoices, headerGuess, writeOut, checkWritable, canChange, changeCells, patchStyle, blockOf, movedCell, hasMergeIn, insertRows, deleteRowSpans, mergeSpans, copyText, rangeFromText, textRows, valueRows, writeBlock, addSheetWithRows, HEAD, fileStem, escapeHtml, rowIsBlank } from './utilShared';
import { menuOf, uniqueSheetName } from './utilities';
import { transformValues } from './tools';
import { openDialog, toast, promptDialog, choiceDialog, alertDialog, confirmDialog } from '../ui/dialog';
import { registerPane, typingIn } from '../ui/panes';
import { h, clear } from '../ui/dom';
import { insertRule } from './conditional';
import { setPrint } from './pagelayout';
import { compareSheetCells } from './tools';
import { kvGet, kvSet, kvKeys, kvDel } from '../util/store';
import { importFile } from '../io';

const G_VIEW = 'View';
const G_RANGE = 'Range';
const G_COMBINE = 'Combine';
const G_CONVERT = 'Convert';
const G_INSERT = 'Insert';
const G_DELETE = 'Delete';
const G_LOOKUP = 'Lookup';
const G_FILTER = 'Filter';
const G_PRINT = 'Print';
const G_SECURITY = 'Security';
const G_BOOK = 'Workbook';
const G_INQUIRE = 'Inquire';
const G_QUERY = 'Query';

// ============================================================================================
// View

const READING_RULE = 'cf_readingLayout';
const readingLayout = { on: false, color: '#fff3cd', unsub: null as null | (() => void), last: '' };

function applyReadingLayout(app: App): void {
  const s = app.sheet;
  const key = `${s.id}:${s.sel.r}:${s.sel.c}`;
  if (key === readingLayout.last) return;
  readingLayout.last = key;
  for (const sh of app.book.sheets) {
    const had = sh.cf.some((r) => r.id === READING_RULE);
    sh.cf = sh.cf.filter((r) => r.id !== READING_RULE);
    if (sh === s) { sh.cf.unshift({ id: READING_RULE, type: 'expression', formulas: ['TRUE'], ranges: [{ r1: s.sel.r, r2: s.sel.r, c1: 0, c2: MAX_COLS - 1 }, { r1: 0, r2: MAX_ROWS - 1, c1: s.sel.c, c2: s.sel.c }], style: { fill: readingLayout.color }, priority: -1 }); }
    else if (had) { /* removed */ }
  }
  app.tab.cf.invalidate(); app.grid.invalidate();
}

function stopReadingLayout(app: App): void {
  readingLayout.on = false; readingLayout.last = '';
  readingLayout.unsub?.(); readingLayout.unsub = null;
  for (const sh of app.book.sheets) sh.cf = sh.cf.filter((r) => r.id !== READING_RULE);
  app.tab.cf.invalidate(); app.grid.invalidate();
}

registerPane({
  id: 'formulaEditor', title: 'Formula Editor',
  render(app, el) { el.dataset.sig = ''; this.update!(app, el, {}); },
  update(app, el) {
    const s = app.sheet; const { r, c } = activeCell(app);
    const cell = s.get(r, c);
    const text = cell?.f != null ? '=' + cell.f : cell?.v == null ? '' : typeof cell.v === 'string' ? cell.v : shownText(s, r, c);
    const sig = `${app.tab.id}:${s.id}:${r}:${c}:${app.book.version}`;
    if (el.dataset.sig === sig || typingIn(el)) return;
    el.dataset.sig = sig;
    clear(el);
    const ta = h('textarea', { class: 'ql-formula-editor', rows: 10, spellcheck: false, style: 'width:100%;font-family:monospace;font-size:14px;resize:vertical' }) as HTMLTextAreaElement;
    ta.value = text;
    const apply = () => { if (!canChange(s, r, c)) { toast('That cell is locked.'); return; } app.book.transact('Edit Cell', () => setInput(app, s, r, c, ta.value)); el.dataset.sig = ''; app.refresh(); };
    ta.addEventListener('keydown', (e) => { if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); apply(); } });
    el.appendChild(h('div', { class: 'ql-help' }, `${s.name}!${cellName(r, c)} — Ctrl+Enter applies`));
    el.appendChild(ta);
    el.appendChild(h('div', { class: 'ql-row' }, h('button', { class: 'ql-btn ql-primary', onclick: apply }, 'Apply'), h('button', { class: 'ql-btn', onclick: () => { el.dataset.sig = ''; (this as any).update(app, el, {}); } }, 'Revert'), h('button', { class: 'ql-btn', onclick: () => { ta.value = prettyFormula(ta.value); } }, 'Pretty-print')));
  },
});

/** Indents a formula one argument per line so long nested formulas can be read. */
export function prettyFormula(f: string): string {
  if (!f.startsWith('=')) return f;
  let out = ''; let depth = 0; let inStr = false;
  for (let i = 0; i < f.length; i++) {
    const ch = f[i];
    if (ch === '"') { inStr = !inStr; out += ch; continue; }
    if (inStr) { out += ch; continue; }
    if (ch === '(') { depth++; out += '(\n' + '  '.repeat(depth); }
    else if (ch === ')') { depth = Math.max(0, depth - 1); out += '\n' + '  '.repeat(depth) + ')'; }
    else if (ch === ',' || ch === ';') out += ch + '\n' + '  '.repeat(depth);
    else out += ch;
  }
  return out.replace(/\n\s*\n/g, '\n');
}

defTools([
  {
    id: 'plus.readingLayout', label: 'Reading Layout', icon: 'highlight_alt', tab: P, group: G_VIEW, kind: 'toggle', noRecord: true, popular: true, keys: ['Mod-Alt-h'], tip: 'Highlight the row and column of the active cell as you move, so a wide table is easy to read', keywords: 'highlight active row column crosshair', origin: 'Kutools (Reading Layout)',
    run: (app) => { if (readingLayout.on) { stopReadingLayout(app); return; } readingLayout.on = true; applyReadingLayout(app); readingLayout.unsub = app.onRefresh((a) => { if (readingLayout.on) applyReadingLayout(a); }); },
    isActive: () => readingLayout.on,
  },
  { id: 'plus.readingLayoutColor', label: 'Reading Layout Color…', icon: 'palette', tab: P, group: G_VIEW, noRecord: true, run: async (app) => { const r = await openDialog({ title: 'Reading Layout Color', fields: [{ key: 'c', label: 'Highlight color', type: 'color', value: readingLayout.color }] }); if (!r) return; readingLayout.color = r.c; readingLayout.last = ''; if (readingLayout.on) applyReadingLayout(app); } },
  { id: 'plus.formulaEditor', label: 'Formula Editor Pane', icon: 'edit_square', tab: P, group: G_VIEW, noRecord: true, popular: true, tip: 'A large, multi-line editor for the active cell\'s formula or text, with pretty-printing', keywords: 'larger formula bar enhanced edit bar', origin: 'Kutools (Enhanced Edit Bar) / Excel Labs (Advanced Formula Environment)', run: (app) => app.ui.togglePane('formulaEditor') },
  { id: 'plus.sheetVisibility', label: 'Workbook and Sheet Visibility…', icon: 'visibility', tab: P, group: G_VIEW, edits: true, tip: 'Tick the sheets that should be visible', origin: 'Kutools (Hide/Unhide Workbooks and Sheets)', run: async (app) => { const pick = await pickSheets(app, 'Sheet Visibility', { help: 'Checked sheets are visible; unchecked ones are hidden.', dflt: 'all' }); if (!pick) return; const vis = new Set(pick.sheets); if (!vis.size) { toast('At least one sheet must stay visible.'); return; } app.book.transact('Sheet Visibility', () => { for (const s of app.book.sheets) { app.book.recordSheet(s); s.visibility = vis.has(s) ? 'visible' : 'hidden'; } }); if (!vis.has(app.sheet)) app.exec('sheet.activate', app.book.sheets.indexOf([...vis][0])); app.ui.rebuild(); } },
  { id: 'plus.syncSelection', label: 'Same Selection on All Sheets', icon: 'sync', tab: P, group: G_VIEW, noRecord: true, tip: 'Give every sheet the active cell, selection and zoom of this one', origin: 'Kutools (Synchronize Worksheets)', run: (app) => { const a = app.sheet; for (const s of app.book.sheets) { if (s === a) continue; s.sel = JSON.parse(JSON.stringify(a.sel)); s.scroll = { ...a.scroll }; s.zoom = a.zoom; } toast('Selection synchronized'); } },
  {
    id: 'plus.snippets', label: 'Insert a Saved Snippet…', icon: 'auto_awesome_motion', tab: P, group: G_VIEW, edits: true, kind: 'split', tip: 'A library of reusable text, formulas and ranges: save the selection as a snippet, insert one later', keywords: 'autotext snippets library reuse', origin: 'Kutools (AutoText)',
    run: async (app, arg?: 'save' | string) => {
      const keys = (await kvKeys('snippet:')).map((k) => k.slice(8));
      if (arg === 'save') { const s = app.sheet; const rg = clipToUsed(s, activeRange(app)); const name = await promptDialog('Save Snippet', 'Snippet name', shownText(s, rg.r1, rg.c1).slice(0, 30)); if (!name) return; await kvSet('snippet:' + name, { rows: blockOf(s, rg), when: Date.now() }); toast(`Snippet "${name}" saved`); return; }
      if (!keys.length) { toast('No snippets yet: use Save the Selection as a Snippet.'); return; }
      const name = typeof arg === 'string' && keys.includes(arg) ? arg : await (async () => { const i = await choiceDialog('Insert Snippet', 'Snippet', keys); return i < 0 ? null : keys[i]; })();
      if (!name) return;
      const snip = await kvGet<{ rows: (Cell | undefined)[][] }>('snippet:' + name); if (!snip) return;
      const { r, c } = activeCell(app); const s = app.sheet;
      const target = { r1: r, c1: c, r2: r + snip.rows.length - 1, c2: c + (snip.rows[0]?.length || 1) - 1 };
      if (!checkWritable(app, [target])) return;
      app.book.transact('Insert Snippet', () => snip.rows.forEach((row, i) => row.forEach((cell, j) => s.set(r + i, c + j, cell ? movedCell(cell, 0, 0) : undefined))));
    },
    menu: () => [{ id: 'plus.snippets', arg: 'save', label: 'Save the Selection as a Snippet…', icon: 'save' }, { id: 'plus.snippets', label: 'Insert a Snippet…', icon: 'auto_awesome_motion' }, { id: 'plus.deleteSnippet', label: 'Delete a Snippet…', icon: 'delete' }],
  },
  { id: 'plus.deleteSnippet', label: 'Delete a Snippet…', icon: 'delete', tab: P, group: G_VIEW, noRecord: true, run: async () => { const keys = (await kvKeys('snippet:')).map((k) => k.slice(8)); if (!keys.length) { toast('No snippets.'); return; } const i = await choiceDialog('Delete Snippet', 'Snippet', keys); if (i < 0) return; await kvDel('snippet:' + keys[i]); toast('Deleted'); } },
  { id: 'plus.menuView', label: 'View', icon: 'visibility', tab: P, group: G_VIEW, kind: 'menu', noRecord: true, run: () => undefined, menu: () => menuOf(['plus.readingLayout', 'plus.readingLayoutColor', 'plus.formulaEditor', 'plus.sheetVisibility', 'plus.syncSelection', 'plus.snippets', 'view.navigator', 'file.snapshot', 'file.versions']) },
]);

// ============================================================================================
// Range

/** Fullwidth ⇄ halfwidth (ASCII) forms. */
export function toHalfWidth(s: string): string { return s.replace(/[\uff01-\uff5e]/g, (ch) => String.fromCharCode(ch.charCodeAt(0) - 0xfee0)).replace(/\u3000/g, ' '); }
export function toFullWidth(s: string): string { return s.replace(/[!-~]/g, (ch) => String.fromCharCode(ch.charCodeAt(0) + 0xfee0)).replace(/ /g, '\u3000'); }

/** Rich-text runs that subscript the digits of a chemical formula (H2O → H₂O as subscript runs). */
export function chemicalRuns(s: string): { t: string; s?: Partial<CellStyle> }[] | null {
  if (!/^[A-Z][a-z]?\d*(?:[A-Z(][^ ]*)?$|^[A-Za-z()\d]+$/.test(s) || !/\d/.test(s) || !/[A-Za-z]/.test(s)) return null;
  const runs: { t: string; s?: Partial<CellStyle> }[] = [];
  for (const m of s.matchAll(/(\d+)|(\D+)/g)) runs.push(m[1] ? { t: m[1], s: { vertAlign: 'subscript' } } : { t: m[2] });
  return runs;
}

export function missingNumbers(values: Scalar[]): number[] {
  const nums = values.filter((v): v is number => typeof v === 'number' && Number.isInteger(v)).sort((a, b) => a - b);
  if (nums.length < 2) return [];
  const set = new Set(nums); const out: number[] = [];
  for (let n = nums[0]; n <= nums[nums.length - 1]; n++) if (!set.has(n)) out.push(n);
  return out;
}

export function combinations(lists: string[][]): string[][] {
  let out: string[][] = [[]];
  for (const list of lists) { const next: string[][] = []; for (const prefix of out) for (const item of list) next.push([...prefix, item]); out = next; }
  return out;
}

async function preventTyping(app: App, kind: 'duplicates' | 'noSpaces' | 'textOnly' | 'numbersOnly' | 'dates' | 'custom', formula?: string): Promise<void> {
  const s = app.sheet; const rg = activeRange(app);
  const first = cellName(rg.r1, rg.c1);
  const col = `$${colName(rg.c1)}:$${colName(rg.c1)}`;
  let dv: Omit<DataValidation, 'id'>;
  switch (kind) {
    case 'duplicates': dv = { ranges: [rg], type: 'custom', formula1: `COUNTIF(${col},${first})<=1`, errorTitle: 'Duplicate', errorMessage: 'This value is already in the column.', errorStyle: 'stop' }; break;
    case 'noSpaces': dv = { ranges: [rg], type: 'custom', formula1: `ISERROR(FIND(" ",${first}))`, errorTitle: 'Spaces', errorMessage: 'Spaces are not allowed here.', errorStyle: 'stop' }; break;
    case 'textOnly': dv = { ranges: [rg], type: 'custom', formula1: `ISTEXT(${first})`, errorTitle: 'Text only', errorMessage: 'Only text is allowed here.', errorStyle: 'stop' }; break;
    case 'numbersOnly': dv = { ranges: [rg], type: 'decimal', operator: 'between', formula1: '-1E+300', formula2: '1E+300', errorTitle: 'Numbers only', errorMessage: 'Only numbers are allowed here.', errorStyle: 'stop' }; break;
    case 'dates': dv = { ranges: [rg], type: 'date', operator: 'between', formula1: '1', formula2: '2958465', errorTitle: 'Dates only', errorMessage: 'Enter a date.', errorStyle: 'stop' }; break;
    default: dv = { ranges: [rg], type: 'custom', formula1: (formula || 'TRUE').replace(/^=/, ''), errorTitle: 'Not allowed', errorMessage: 'The entry does not meet the rule.', errorStyle: 'stop' };
  }
  app.book.transact('Prevent Typing', () => { app.book.recordSheet(s); s.validations = s.validations.filter((v) => !v.ranges.some((vr) => vr.r1 === rg.r1 && vr.c1 === rg.c1 && vr.r2 === rg.r2 && vr.c2 === rg.c2)); s.validations.push({ ...dv, id: newId('dv') }); });
  toast(`Rule set on ${rangeName(rg)}`);
}

defTools([
  { id: 'plus.flipH', label: 'Flip Range Horizontally', icon: 'flip', tab: P, group: G_RANGE, edits: true, origin: 'Kutools / Ablebits (Flip)', run: (app) => app.exec('tools.reverse', 'cols') },
  { id: 'plus.flipV', label: 'Flip Range Vertically', icon: 'flip', tab: P, group: G_RANGE, edits: true, origin: 'Kutools / Ablebits (Flip)', run: (app) => app.exec('tools.reverse', 'rows') },
  {
    id: 'plus.rangeHelper', label: 'Range Helper: Add, Subtract or Intersect…', icon: 'select_all', tab: P, group: G_RANGE, noRecord: true, tip: 'Build a selection from several ranges: add a range, subtract one, or keep only the overlap', origin: 'Kutools (Select Range Helper)',
    run: async (app, arg?: { op: 'add' | 'subtract' | 'intersect'; range: string }) => {
      const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Range Helper', fields: [{ key: 'op', label: 'Operation', type: 'radio', value: 'add', options: [{ value: 'add', label: 'Add a range to the selection' }, { value: 'subtract', label: 'Remove a range from the selection' }, { value: 'intersect', label: 'Keep only the cells both have' }] }, { key: 'range', label: 'Range', type: 'text', value: '' }] }); return r ? { op: r.op, range: r.range } : null; });
      if (!o) return;
      const rg = parseRange(o.range.replace(/\$/g, '')); if (!rg) { toast('That range is not valid.'); return; }
      const s = app.sheet;
      if (o.op === 'add') { s.sel.ranges = [...s.sel.ranges, rg]; }
      else if (o.op === 'subtract') { const { subtractRange } = await import('./dataops'); s.sel.ranges = s.sel.ranges.flatMap((x) => subtractRange(x, rg)); if (!s.sel.ranges.length) s.sel.ranges = [{ r1: s.sel.r, c1: s.sel.c, r2: s.sel.r, c2: s.sel.c }]; }
      else { const { intersection } = await import('../model/address'); const next = s.sel.ranges.map((x) => intersection(x, rg)).filter((x): x is Range => !!x); if (!next.length) { toast('The ranges do not overlap.'); return; } s.sel.ranges = next; }
      s.sel.r = s.sel.ranges[0].r1; s.sel.c = s.sel.ranges[0].c1; app.grid.invalidate(); app.refresh();
    },
  },
  {
    id: 'plus.preventTyping', label: 'Prevent Typing…', icon: 'block', tab: P, group: G_RANGE, edits: true, popular: true, kind: 'split', tip: 'Data validation presets: no duplicates in the column, no spaces, text only, numbers only, dates only, or your own formula', keywords: 'validation prevent duplicates no spaces', origin: 'Kutools (Prevent Typing / Prevent Duplicate Entries)',
    run: async (app, arg?: 'duplicates' | 'noSpaces' | 'textOnly' | 'numbersOnly' | 'dates' | 'custom') => { const kind = await argOrAsk(arg, async () => { const i = await choiceDialog('Prevent Typing', 'In the selected cells, allow:', ['No duplicates in the column', 'No spaces', 'Text only', 'Numbers only', 'Dates only', 'Custom formula…']); return i < 0 ? null : (['duplicates', 'noSpaces', 'textOnly', 'numbersOnly', 'dates', 'custom'] as const)[i]; }); if (!kind) return; let f: string | undefined; if (kind === 'custom') { const t = await promptDialog('Custom Rule', 'Formula that must be TRUE (relative to the first cell)', `=LEN(${cellName(activeRange(app).r1, activeRange(app).c1)})<=10`); if (!t) return; f = t; } await preventTyping(app, kind, f); },
    menu: () => [{ id: 'plus.preventTyping', arg: 'duplicates', label: 'Prevent Duplicates in the Column' }, { id: 'plus.preventTyping', arg: 'noSpaces', label: 'Prevent Spaces' }, { id: 'plus.preventTyping', arg: 'textOnly', label: 'Text Only' }, { id: 'plus.preventTyping', arg: 'numbersOnly', label: 'Numbers Only' }, { id: 'plus.preventTyping', arg: 'dates', label: 'Dates Only' }, { id: 'plus.preventTyping', arg: 'custom', label: 'Custom Formula…' }, { separator: true }, { id: 'data.validation', label: 'Data Validation…' }],
  },
  {
    id: 'plus.dropdownFromColumn', label: 'Drop-Down List from the Column\'s Values', icon: 'arrow_drop_down_circle', tab: P, group: G_RANGE, edits: true, popular: true, tip: 'Make the selected cells a drop-down list of the distinct values already in their column (or the selection)', keywords: 'dropdown list unique values validation', origin: 'Kutools (Drop-down List) / Ablebits',
    run: (app) => { const s = app.sheet; const rg = activeRange(app); const vals = new Set<string>(); const u = s.usedRange(false); const scan = singleCell(app) && u ? { r1: u.r1, r2: u.r2, c1: rg.c1, c2: rg.c1 } : clipToUsed(s, rg); for (let r = scan.r1; r <= scan.r2; r++) for (let c = scan.c1; c <= scan.c2; c++) { const t = shownText(s, r, c).trim(); if (t) vals.add(t); } const list = [...vals].sort((a, b) => a.localeCompare(b)); if (!list.length) { toast('No values found.'); return; } app.book.transact('Drop-Down List', () => { app.book.recordSheet(s); s.validations.push({ id: newId('dv'), ranges: [rg], type: 'list', formula1: list.join(',').slice(0, 255), showDropdown: true, allowBlank: true }); }); toast(`Drop-down with ${plural(list.length, 'choice')} set`); },
  },
  {
    id: 'plus.coloredDropdown', label: 'Color the Drop-Down Choices', icon: 'palette', tab: P, group: G_RANGE, edits: true, tip: 'Give each choice of the selected cells\' drop-down list its own fill color (conditional formatting rules)', origin: 'Kutools (Colored Drop-down List)',
    run: (app) => { const s = app.sheet; const rg = activeRange(app); const dv = s.validations.find((v) => v.type === 'list' && v.ranges.some((x) => rg.r1 >= x.r1 && rg.r1 <= x.r2 && rg.c1 >= x.c1 && rg.c1 <= x.c2)); if (!dv?.formula1) { toast('The active cell has no drop-down list.'); return; } const items = dv.formula1.startsWith('=') ? [] : dv.formula1.split(',').map((x) => x.trim()).filter(Boolean); if (!items.length) { toast('Only lists typed into the validation (a,b,c) can be colored.'); return; } const palette = ['#c8e6c9', '#bbdefb', '#fff9c4', '#ffccbc', '#e1bee7', '#b2dfdb', '#f8bbd0', '#d7ccc8']; app.book.transact('Color Drop-Down', () => { app.book.recordSheet(s); items.forEach((it, i) => insertRule(s, { type: 'cellIs', operator: 'equal', formulas: [`"${it.replace(/"/g, '""')}"`], ranges: dv.ranges.map((x) => ({ ...x })), style: { fill: palette[i % palette.length] } })); }); app.tab.cf.invalidate(); app.grid.invalidate(); toast(`${plural(items.length, 'choice')} colored`); },
  },
  { id: 'plus.selectUnmerged', label: 'Select Unmerged Cells', icon: 'grid_on', tab: P, group: G_RANGE, noRecord: true, origin: 'Kutools', run: (app) => { const s = app.sheet; const hits = selCells(app).filter((x) => !s.mergeAt(x.r, x.c)).map((x) => ({ r: x.r, c: x.c })); selectHits(app, hits, 'unmerged cells'); } },
  { id: 'plus.mergeSameCells', label: 'Merge Cells with the Same Value', icon: 'call_merge', tab: P, group: G_RANGE, edits: true, tip: 'In each selected column, merge runs of adjacent cells that hold the same value (for report layouts)', origin: 'Kutools (Merge Same Cells)', run: (app) => { const s = app.sheet; const rg = clipToUsed(s, activeRange(app)); let n = 0; app.book.transact('Merge Same Cells', () => { app.book.recordSheet(s); for (let c = rg.c1; c <= rg.c2; c++) { let start = rg.r1; for (let r = rg.r1 + 1; r <= rg.r2 + 1; r++) { const same = r <= rg.r2 && shownText(s, r, c) !== '' && shownText(s, r, c) === shownText(s, start, c); if (!same) { if (r - 1 > start) { s.merges.push({ r1: start, r2: r - 1, c1: c, c2: c }); for (let rr = start + 1; rr <= r - 1; rr++) { const cur = s.get(rr, c); s.set(rr, c, cur?.s ? { s: cur.s } : undefined); } n++; } start = r; } } } }); app.grid.invalidate(); toast(`${plural(n, 'run')} merged`); } },
  { id: 'plus.menuRange', label: 'Range', icon: 'select_all', tab: P, group: G_RANGE, kind: 'menu', noRecord: true, run: () => undefined, menu: () => menuOf(['plus.flipH', 'plus.flipV', 'tools.transposeInPlace', 'util.swapRanges', 'tools.shuffle', 'util.selectRandom', null, 'plus.rangeHelper', 'plus.selectUnmerged', 'util.selectDuplicates', 'util.selectMaxMin', 'util.compareRanges', null, 'plus.preventTyping', 'plus.dropdownFromColumn', 'plus.coloredDropdown', 'plus.mergeSameCells', 'tools.unmergeFill', 'util.mergeToOneCell']) },
]);

// ============================================================================================
// Combine

function sheetRows(s: Sheet): { header: string[]; rows: Scalar[][]; rg: Range } | null {
  const u = s.usedRange(false); if (!u) return null;
  const rows = valueRows(s, u);
  const header = (rows[0] || []).map((v, j) => textOf(v) || colName(u.c1 + j));
  return { header, rows: rows.slice(1), rg: u };
}

async function combineSheets(app: App, arg?: { names?: string[]; mode: 'stack' | 'byHeader' | 'position' | 'fn'; fn?: string; nameColumn?: boolean }): Promise<void> {
  if (structureLocked(app)) return;
  const pick = arg ? { sheets: sheetsByNames(app.book, arg.names), values: arg } : await pickSheets(app, 'Combine Sheets', { help: 'Stack: rows under each other (one header). By header: columns matched by name. By position: values of the same cell added up (or counted, averaged…).', extra: [{ key: 'mode', label: 'How', type: 'radio', value: 'byHeader', options: [{ value: 'stack', label: 'Stack the rows (same column order)' }, { value: 'byHeader', label: 'Stack, matching columns by header name' }, { value: 'position', label: 'Consolidate by position: sum the same cells' }] }, { key: 'fn', label: 'Function for consolidation', type: 'select', value: 'SUM', options: ['SUM', 'COUNT', 'AVERAGE', 'MAX', 'MIN'], visibleIf: (v) => v.mode === 'position' }, { key: 'nameColumn', label: 'Add a column with the sheet name', type: 'checkbox', value: true, visibleIf: (v) => v.mode !== 'position' }] });
  if (!pick || !pick.sheets.length) return;
  const v = pick.values as any; const mode = v.mode || 'byHeader';
  const sheets = pick.sheets.filter((s) => s.usedRange(false));
  if (!sheets.length) { toast('The chosen sheets are empty.'); return; }
  const out: Scalar[][] = [];
  if (mode === 'position') {
    const fn = String(v.fn || 'SUM').toUpperCase();
    const acc = new Map<string, number[]>(); let maxR = 0, maxC = 0;
    for (const s of sheets) { const u = s.usedRange(false)!; maxR = Math.max(maxR, u.r2 + 1); maxC = Math.max(maxC, u.c2 + 1); s.forEachIn(u, (r, c, cell) => { if (typeof cell.v === 'number') (acc.get(`${r},${c}`) || acc.set(`${r},${c}`, []).get(`${r},${c}`)!).push(cell.v); }); }
    const first = sheets[0]; const u0 = first.usedRange(false)!;
    for (let r = 0; r < maxR; r++) { const row: Scalar[] = []; for (let c = 0; c < maxC; c++) { const xs = acc.get(`${r},${c}`); if (!xs) { const t = first.value(r, c); row.push(typeof t === 'string' ? t : null); continue; } row.push(fn === 'COUNT' ? xs.length : fn === 'AVERAGE' ? xs.reduce((a, b) => a + b, 0) / xs.length : fn === 'MAX' ? Math.max(...xs) : fn === 'MIN' ? Math.min(...xs) : xs.reduce((a, b) => a + b, 0)); } out.push(row); }
    void u0;
    addSheetWithRows(app, 'Consolidated', out);
    return;
  }
  const nameCol = v.nameColumn !== false;
  if (mode === 'stack') {
    const first = sheetRows(sheets[0])!;
    out.push([...(nameCol ? ['Sheet'] : []), ...first.header].map((t) => ({ v: t, style: HEAD })) as any);
    for (const s of sheets) { const d = sheetRows(s)!; for (const row of d.rows) out.push([...(nameCol ? [s.name] : []), ...row]); }
  } else {
    const headers: string[] = [];
    const data = sheets.map((s) => sheetRows(s)!);
    for (const d of data) for (const hh of d.header) if (!headers.some((x) => x.toLowerCase() === hh.toLowerCase())) headers.push(hh);
    out.push([...(nameCol ? ['Sheet'] : []), ...headers].map((t) => ({ v: t, style: HEAD })) as any);
    data.forEach((d, i) => { const map = d.header.map((hh) => headers.findIndex((x) => x.toLowerCase() === hh.toLowerCase())); for (const row of d.rows) { const line: Scalar[] = headers.map(() => null); row.forEach((val, j) => { if (map[j] >= 0) line[map[j]] = val; }); out.push([...(nameCol ? [sheets[i].name] : []), ...line]); } });
  }
  addSheetWithRows(app, 'Combined', out as any, { freeze: 1 });
  toast(`${plural(sheets.length, 'sheet')} combined (${out.length - 1} rows)`);
}

async function splitToSheets(app: App, arg?: { col: number; header?: boolean }): Promise<void> {
  if (emptySheet(app) || structureLocked(app)) return;
  const s = app.sheet; const rg = selBlock(app); const header = headerGuess(s, rg);
  const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Split Rows into Sheets', fields: [{ key: 'col', label: 'One sheet per distinct value in column', type: 'select', value: String(rg.c1), options: columnChoices(s, rg, header) }, { key: 'header', label: 'Repeat the header row on each sheet', type: 'checkbox', value: header }] }); return r ? { col: Number(r.col), header: !!r.header } : null; });
  if (!o) return;
  const first = rg.r1 + (o.header ? 1 : 0);
  const groups = new Map<string, number[]>();
  for (let r = first; r <= rg.r2; r++) { const k = shownText(s, r, o.col) || '(blank)'; (groups.get(k) || groups.set(k, []).get(k)!).push(r); }
  if (groups.size > 200) { toast(`That column has ${groups.size} distinct values; choose one with fewer.`); return; }
  const book = app.book; const widths = Array.from({ length: rg.c2 - rg.c1 + 1 }, (_x, j) => s.colW.get(rg.c1 + j));
  book.transact('Split into Sheets', () => {
    book.recordBook();
    let at = book.activeSheet; const taken = new Set<string>();
    for (const [name, rows] of groups) {
      const sheetName = uniqueSheetName(book, name, taken); taken.add(sheetName.toLowerCase());
      const ns = new Sheet(book, sheetName);
      let out = 0;
      if (o.header) { blockOf(s, { r1: rg.r1, r2: rg.r1, c1: rg.c1, c2: rg.c2 })[0].forEach((cell, j) => ns.set(0, j, movedCell(cell, -rg.r1, -rg.c1))); out = 1; ns.freeze = { rows: 1, cols: 0 }; }
      for (const r of rows) { blockOf(s, { r1: r, r2: r, c1: rg.c1, c2: rg.c2 })[0].forEach((cell, j) => ns.set(out, j, movedCell(cell, out - r, -rg.c1))); out++; }
      widths.forEach((w, j) => { if (w) ns.colW.set(j, w); });
      book.sheets.splice(++at, 0, ns);
    }
  });
  app.ui.rebuild(); app.refresh();
  toast(`${plural(groups.size, 'sheet')} created`);
}

/** Joins a lookup table's columns onto the main table by key (a VLOOKUP done as values). */
export function mergeTables(main: Scalar[][], lookup: Scalar[][], mainKey: number[], lookupKey: number[], bring: number[], opts: { caseSensitive?: boolean; multi?: 'first' | 'last' | 'join' } = {}): { rows: Scalar[][]; matched: number } {
  const norm = (vs: Scalar[]) => vs.map((v) => (v == null ? '' : opts.caseSensitive ? String(v) : String(v).toLowerCase().trim())).join('\u0001');
  const index = new Map<string, Scalar[][]>();
  for (const row of lookup) { const k = norm(lookupKey.map((c) => row[c])); (index.get(k) || index.set(k, []).get(k)!).push(row); }
  let matched = 0;
  const rows = main.map((row) => {
    const hits = index.get(norm(mainKey.map((c) => row[c])));
    if (!hits) return [...row, ...bring.map(() => null as Scalar)];
    matched++;
    const pick = opts.multi === 'last' ? hits[hits.length - 1] : hits[0];
    return [...row, ...bring.map((c) => (opts.multi === 'join' ? hits.map((hh) => textOf(hh[c])).filter(Boolean).join(', ') : pick[c] ?? null))];
  });
  return { rows, matched };
}

defTools([
  { id: 'util.combineSheets', label: 'Combine Sheets…', icon: 'merge', tab: P, group: G_COMBINE, edits: true, popular: true, tip: 'Stack the rows of several sheets on one sheet (matching columns by header if you like) or consolidate them cell by cell', keywords: 'combine sheets append consolidate merge', origin: 'Ablebits (Combine Sheets / Consolidate Sheets) / Kutools (Combine) / Power Tools', run: (app, arg) => combineSheets(app, arg) },
  { id: 'util.splitToSheets', label: 'Split Rows into Sheets by Column…', icon: 'call_split', tab: P, group: G_COMBINE, edits: true, popular: true, tip: 'One new sheet per distinct value of a column (one per customer, region, month…)', keywords: 'split data sheets by value group', origin: 'Kutools (Split Data) / Ablebits (Split Table)', run: (app, arg) => splitToSheets(app, arg) },
  {
    id: 'plus.mergeTables', label: 'Merge Two Tables by Key…', icon: 'join_inner', tab: P, group: G_COMBINE, edits: true, popular: true, tip: 'Bring columns from another table (even on another sheet) into this one, matched by one or more key columns: a lookup done once, as values', keywords: 'join vlookup merge tables lookup columns', origin: 'Ablebits (Merge Two Tables) / Kutools (Tables Merge)',
    run: async (app, arg?: { lookup: string; mainKey: number[]; lookupKey: number[]; bring: number[]; multi?: 'first' | 'last' | 'join' }) => {
      if (emptySheet(app)) return;
      const s = app.sheet; const rg = selBlock(app); const header = headerGuess(s, rg);
      const o = await argOrAsk(arg, async () => {
        const r1 = await openDialog({ title: 'Merge Two Tables: lookup table', help: `Main table: ${rangeName(rg)} on ${s.name}. Give the range of the lookup table, including its header row.`, fields: [{ key: 'lookup', label: 'Lookup table range', type: 'text', value: '', placeholder: 'Sheet2!A1:F200' }] });
        if (!r1) return null;
        const L = rangeFromText(app, r1.lookup); if (!L) { toast('That range is not valid.'); return null; }
        const r2 = await openDialog({ title: 'Merge Two Tables: columns', fields: [{ key: 'mainKey', label: 'Key column in this table', type: 'list', multiple: true, rows: 5, options: columnChoices(s, rg, header).map((x) => x.label), value: [columnChoices(s, rg, header)[0].label] }, { key: 'lookupKey', label: 'Key column in the lookup table', type: 'list', multiple: true, rows: 5, options: columnChoices(L.sheet, L.range, true).map((x) => x.label), value: [columnChoices(L.sheet, L.range, true)[0].label] }, { key: 'bring', label: 'Columns to bring in', type: 'list', multiple: true, rows: 6, options: columnChoices(L.sheet, L.range, true).map((x) => x.label) }, { key: 'multi', label: 'When several rows match', type: 'select', value: 'first', options: [{ value: 'first', label: 'Use the first match' }, { value: 'last', label: 'Use the last match' }, { value: 'join', label: 'Join all matches with commas' }] }] });
        if (!r2) return null;
        const idx = (labels: string[], choices: { value: string; label: string }[]) => (Array.isArray(labels) ? labels : [labels]).map((l) => Number(choices.find((c) => c.label === l)?.value)).filter((n) => !isNaN(n));
        return { lookup: r1.lookup, mainKey: idx(r2.mainKey, columnChoices(s, rg, header)), lookupKey: idx(r2.lookupKey, columnChoices(L.sheet, L.range, true)), bring: idx(r2.bring, columnChoices(L.sheet, L.range, true)), multi: r2.multi };
      });
      if (!o || !o.mainKey.length || !o.lookupKey.length || !o.bring.length) return;
      const L = rangeFromText(app, o.lookup); if (!L) return;
      const main = valueRows(s, { ...rg, r1: rg.r1 + (header ? 1 : 0) });
      const lookup = valueRows(L.sheet, { ...L.range, r1: L.range.r1 + 1 });
      const res = mergeTables(main, lookup, o.mainKey.map((c) => c - rg.c1), o.lookupKey.map((c) => c - L.range.c1), o.bring.map((c) => c - L.range.c1), { multi: o.multi });
      const target = { r1: rg.r1, r2: rg.r2, c1: rg.c2 + 1, c2: rg.c2 + o.bring.length };
      if (!checkWritable(app, [target])) return;
      app.book.transact('Merge Tables', () => { if (header) o.bring.forEach((c, j) => writeOut(s, rg.r1, target.c1 + j, { v: shownText(L.sheet, L.range.r1, c), style: HEAD })); res.rows.forEach((row, i) => row.slice(main[0]?.length ?? row.length - o.bring.length).forEach((v, j) => { if (v != null) writeOut(s, rg.r1 + (header ? 1 : 0) + i, target.c1 + j, v); })); });
      toast(`${res.matched} of ${plural(main.length, 'row')} matched`);
    },
  },
  { id: 'plus.combineColumnsKeep', label: 'Combine Columns (Keep the Originals)', icon: 'join', tab: P, group: G_COMBINE, edits: true, tip: 'Join the selected columns into a new column to the right, leaving the originals in place', origin: 'Kutools (Combine Rows, Columns or Cells without Losing Data)', run: (app) => app.exec('tools.mergeColumns', { sep: ' ', skipBlanks: true, keep: true }) },
  { id: 'plus.menuCombine', label: 'Combine', icon: 'merge', tab: P, group: G_COMBINE, kind: 'menu', noRecord: true, run: () => undefined, menu: () => menuOf(['util.combineSheets', 'util.splitToSheets', 'plus.mergeTables', 'util.joinRowsByKey', 'util.importStacked', null, 'tools.mergeColumns', 'plus.combineColumnsKeep', 'util.mergeToOneCell', 'tools.splitColumn', 'plus.splitNames', 'data.consolidate']) },
]);

// ============================================================================================
// Convert

export interface NameParts { title: string; first: string; middle: string; last: string; suffix: string }
const TITLES = ['mr', 'mrs', 'ms', 'miss', 'dr', 'prof', 'rev', 'sir', 'hon', 'capt', 'col', 'lt', 'sgt', 'fr', 'sr', 'st'];
const SUFFIXES = ['jr', 'sr', 'ii', 'iii', 'iv', 'v', 'phd', 'md', 'esq', 'dds', 'cpa', 'ret'];
const PARTICLES = ['van', 'von', 'de', 'der', 'den', 'del', 'della', 'di', 'da', 'du', 'la', 'le', 'mac', 'mc', 'st.', 'ter', 'ten', 'bin', 'ibn', 'al', 'el', 'af', 'zu'];

/** Splits "Dr. John A. Smith Jr.", "Smith, John A." or "SMITH John" into title, first, middle, last and suffix. */
export function splitName(full: string): NameParts {
  let s = full.replace(/\s+/g, ' ').trim();
  const out: NameParts = { title: '', first: '', middle: '', last: '', suffix: '' };
  if (!s) return out;
  const clean = (w: string) => w.replace(/[.,]+$/g, '').toLowerCase();
  if (s.includes(',')) {
    const [last, rest = ''] = s.split(',').map((x) => x.trim());
    const restWords = rest.split(' ').filter(Boolean);
    if (restWords.length && SUFFIXES.includes(clean(restWords[restWords.length - 1]))) out.suffix = restWords.pop()!;
    if (restWords.length && TITLES.includes(clean(restWords[0]))) out.title = restWords.shift()!;
    out.last = last; out.first = restWords.shift() || ''; out.middle = restWords.join(' ');
    return out;
  }
  const words = s.split(' ');
  if (words.length > 1 && TITLES.includes(clean(words[0]))) out.title = words.shift()!;
  if (words.length > 1 && SUFFIXES.includes(clean(words[words.length - 1]))) out.suffix = words.pop()!;
  if (words.length === 1) { out.first = words[0]; return out; }
  let lastStart = words.length - 1;
  while (lastStart > 1 && PARTICLES.includes(clean(words[lastStart - 1]))) lastStart--;
  out.last = words.slice(lastStart).join(' ');
  out.first = words[0];
  out.middle = words.slice(1, lastStart).join(' ');
  return out;
}

defTools([
  {
    id: 'plus.splitNames', label: 'Split Names…', icon: 'person', tab: P, group: G_CONVERT, edits: true, popular: true, tip: 'Full names into title, first, middle, last and suffix columns; handles "Last, First", particles like van and de, and Jr./III', keywords: 'split names first last title suffix', origin: 'Ablebits (Split Names) / Kutools (Split Names)',
    run: async (app, arg?: { parts: (keyof NameParts)[] }) => {
      const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Split Names', fields: [{ key: 'title', label: 'Title (Mr, Dr…)', type: 'checkbox', value: false, half: true }, { key: 'first', label: 'First name', type: 'checkbox', value: true, half: true }, { key: 'middle', label: 'Middle name(s)', type: 'checkbox', value: true, half: true }, { key: 'last', label: 'Last name', type: 'checkbox', value: true, half: true }, { key: 'suffix', label: 'Suffix (Jr, III…)', type: 'checkbox', value: false, half: true }] }); return r ? { parts: (['title', 'first', 'middle', 'last', 'suffix'] as (keyof NameParts)[]).filter((k) => r[k]) } : null; });
      if (!o || !o.parts.length) return;
      const s = app.sheet; const rg = clipToUsed(s, activeRange(app)); const header = headerGuess(s, rg);
      const target = { r1: rg.r1, r2: rg.r2, c1: rg.c2 + 1, c2: rg.c2 + o.parts.length };
      if (!checkWritable(app, [target])) return;
      const labels: Record<string, string> = { title: 'Title', first: 'First name', middle: 'Middle name', last: 'Last name', suffix: 'Suffix' };
      app.book.transact('Split Names', () => { for (let r = rg.r1; r <= rg.r2; r++) { if (r === rg.r1 && header) { o.parts.forEach((p, j) => writeOut(s, r, target.c1 + j, { v: labels[p], style: HEAD })); continue; } const t = shownText(s, r, rg.c1); if (!t) continue; const parts = splitName(t); o.parts.forEach((p, j) => writeOut(s, r, target.c1 + j, parts[p])); } });
      app.grid.select(target.r1, target.c1, { range: target });
    },
  },
  { id: 'plus.halfWidth', label: 'Full-Width to Half-Width Characters', icon: 'text_fields', tab: P, group: G_CONVERT, edits: true, tip: 'Ｅｘｃｅｌ → Excel: ASCII letters, digits and punctuation in their full-width forms become normal', origin: 'Kutools (Full Width to Half Width)', run: (app) => transformValues(app, 'Half Width', (v) => toHalfWidth(String(v))) },
  { id: 'plus.fullWidth', label: 'Half-Width to Full-Width Characters', icon: 'text_fields', tab: P, group: G_CONVERT, edits: true, origin: 'Kutools', run: (app) => transformValues(app, 'Full Width', (v) => toFullWidth(String(v))) },
  { id: 'plus.chemicalSubscripts', label: 'Subscript the Numbers in Chemical Formulas', icon: 'science', tab: P, group: G_CONVERT, edits: true, tip: 'H2O, CO2 and C6H12O6 get their digits as subscripts', origin: 'Kutools (Subscript Numbers in Chemical Formulas)', run: (app) => { let n = 0; changeCells(app, 'Chemical Subscripts', (cell) => { if (typeof cell.v !== 'string') return undefined; const runs = chemicalRuns(cell.v); if (!runs) return undefined; n++; return { ...cell, rich: runs }; }); void n; } },
  { id: 'plus.superscriptOrdinals', label: 'Superscript Ordinal Suffixes (1st, 2nd…)', icon: 'superscript', tab: P, group: G_CONVERT, edits: true, tip: '1st, 2nd, 3rd, 4th: the suffix becomes superscript', origin: 'Kutools (Super/Subscript)', run: (app) => changeCells(app, 'Superscript Ordinals', (cell) => { if (typeof cell.v !== 'string' || !/\d(st|nd|rd|th)\b/i.test(cell.v)) return undefined; const runs: { t: string; s?: Partial<CellStyle> }[] = []; let last = 0; for (const m of cell.v.matchAll(/(\d)(st|nd|rd|th)\b/gi)) { const i = m.index! + 1; runs.push({ t: cell.v.slice(last, i) }); runs.push({ t: m[2], s: { vertAlign: 'superscript' } }); last = i + m[2].length; } runs.push({ t: cell.v.slice(last) }); return { ...cell, rich: runs.filter((r) => r.t) }; }) },
  {
    id: 'plus.tableToList', label: 'Cross-Tab to List, or List to Cross-Tab', icon: 'pivot_table_chart', tab: P, group: G_CONVERT, edits: true, kind: 'split', tip: 'A matrix (rows × columns) becomes a three-column list, or a three-column list (row, column, value) becomes a matrix', origin: 'Ablebits (Convert Range) / Kutools (Transpose Table Dimensions)',
    run: async (app, arg?: 'toList' | 'toTable') => {
      const s = app.sheet; const rg = selBlock(app);
      if (arg === 'toTable') {
        const rows = valueRows(s, rg).slice(1).filter((r) => r[0] != null && r[1] != null);
        const rKeys = [...new Set(rows.map((r) => textOf(r[0])))], cKeys = [...new Set(rows.map((r) => textOf(r[1])))];
        const m = new Map(rows.map((r) => [`${textOf(r[0])}\u0001${textOf(r[1])}`, r[2] ?? null]));
        const out: Scalar[][] = [['', ...cKeys], ...rKeys.map((rk) => [rk, ...cKeys.map((ck) => m.get(`${rk}\u0001${ck}`) ?? null)])];
        addSheetWithRows(app, 'Cross-tab', out, { freeze: 1 }); return;
      }
      const rows = valueRows(s, rg); const header = rows[0] || [];
      const out: Scalar[][] = [['Row', 'Column', 'Value']];
      for (let i = 1; i < rows.length; i++) for (let j = 1; j < rows[i].length; j++) if (rows[i][j] != null && rows[i][j] !== '') out.push([rows[i][0], header[j], rows[i][j]]);
      addSheetWithRows(app, 'List', out, { freeze: 1 });
    },
    menu: () => [{ id: 'plus.tableToList', arg: 'toList', label: 'Cross-tab → list (row, column, value)' }, { id: 'plus.tableToList', arg: 'toTable', label: 'List (row, column, value) → cross-tab' }],
  },
  { id: 'plus.menuConvert', label: 'Convert', icon: 'transform', tab: P, group: G_CONVERT, kind: 'menu', noRecord: true, run: () => undefined, menu: () => menuOf(['plus.splitNames', 'plus.tableToList', 'tools.unpivot', 'util.columnToTable', 'util.tableToColumn', null, 'tools.textToNumbers', 'util.numbersToText', 'tools.fixDates', 'util.datesToText', 'util.timeToHours', 'util.convertUnits', 'util.numberBase', 'tools.romanNumerals', 'tools.numberToWords', null, 'plus.halfWidth', 'plus.fullWidth', 'plus.chemicalSubscripts', 'plus.superscriptOrdinals', 'tools.removeAccents', 'tools.changeCase']) },
]);

// ============================================================================================
// Insert

defTools([
  {
    id: 'plus.titleRows', label: 'Repeat the Header Row Every N Rows…', icon: 'table_rows', tab: P, group: G_INSERT, edits: true, tip: 'Insert a copy of the header row after every N data rows (for printed lists)', origin: 'Kutools (Insert Titles Rows)',
    run: async (app, arg?: { every: number }) => { const o = await argOrAsk(arg, async () => { const t = await promptDialog('Repeat the Header Row', 'After every … rows', '20', { type: 'number' }); return t ? { every: Number(t) || 20 } : null; }); if (!o || o.every < 1) return; const s = app.sheet; const rg = selBlock(app); if (sheetProtected(app, 'insertRows')) return; const head = blockOf(s, { r1: rg.r1, r2: rg.r1, c1: rg.c1, c2: rg.c2 })[0]; const spots: number[] = []; for (let r = rg.r1 + 1 + o.every; r <= rg.r2; r += o.every) spots.push(r); app.book.transact('Title Rows', () => { for (const at of spots.reverse()) { insertRows(s, at, 1); head.forEach((cell, j) => s.set(at, rg.c1 + j, movedCell(cell, at - rg.r1, 0))); } }); app.grid.updateExtent(); app.grid.layout(); toast(`${plural(spots.length, 'header row')} inserted`); },
  },
  {
    id: 'plus.pageBreakEvery', label: 'Insert a Page Break Every N Rows…', icon: 'insert_page_break', tab: P, group: G_INSERT, edits: true, origin: 'Kutools (Insert Page Break Every Row)',
    run: async (app, arg?: { every: number }) => { const o = await argOrAsk(arg, async () => { const t = await promptDialog('Page Break Every N Rows', 'Every … rows', '50', { type: 'number' }); return t ? { every: Number(t) || 50 } : null; }); if (!o || o.every < 1) return; const s = app.sheet; const rg = selBlock(app); const breaks: number[] = []; for (let r = rg.r1 + o.every; r <= rg.r2; r += o.every) breaks.push(r); app.book.transact('Page Breaks', () => { app.book.recordSheet(s); s.rowBreaks = [...new Set([...s.rowBreaks, ...breaks])].sort((a, b) => a - b); }); (app.grid as any).showPageBreaks = true; app.grid.invalidate(); toast(`${plural(breaks.length, 'page break')} inserted`); },
  },
  { id: 'plus.findMissing', label: 'Find Missing Sequence Numbers', icon: 'format_list_numbered', tab: P, group: G_INSERT, edits: true, kind: 'split', tip: 'Which numbers are missing from the sequence in the selection: list them on a new sheet, or insert blank rows for them', origin: 'Kutools (Find Missing Sequence Number)', run: (app, arg?: 'list' | 'rows') => { const s = app.sheet; const rg = clipToUsed(s, activeRange(app)); const vals: Scalar[] = []; for (let r = rg.r1; r <= rg.r2; r++) vals.push(s.value(r, rg.c1)); const missing = missingNumbers(vals); if (!missing.length) { toast('No gaps in the sequence.'); return; } if (arg === 'rows') { if (sheetProtected(app, 'insertRows')) return; const rows = vals.map((v, i) => ({ v, r: rg.r1 + i })); app.book.transact('Insert Missing', () => { for (const m of [...missing].reverse()) { const after = rows.filter((x) => typeof x.v === 'number' && x.v < m).pop(); const at = after ? after.r + 1 : rg.r1; insertRows(s, at, 1); writeOut(s, at, rg.c1, { v: m, style: { fill: '#fff3cd' } }); for (const x of rows) if (x.r >= at) x.r++; } }); app.grid.updateExtent(); app.grid.layout(); toast(`${plural(missing.length, 'row')} inserted`); return; } report(app, 'Missing Numbers', ['Missing'], missing.map((m) => [m])); }, menu: () => [{ id: 'plus.findMissing', arg: 'list', label: 'List the missing numbers' }, { id: 'plus.findMissing', arg: 'rows', label: 'Insert a highlighted row for each' }] },
  {
    id: 'plus.combinations', label: 'List All Combinations…', icon: 'grid_view', tab: P, group: G_INSERT, edits: true, tip: 'Every combination of the values in the selected columns (sizes × colors × materials) on a new sheet', origin: 'Kutools (List All Combinations)',
    run: (app) => { const s = app.sheet; const rg = clipToUsed(s, activeRange(app)); const lists: string[][] = []; for (let c = rg.c1; c <= rg.c2; c++) { const l: string[] = []; for (let r = rg.r1; r <= rg.r2; r++) { const t = shownText(s, r, c); if (t) l.push(t); } if (l.length) lists.push(l); } const total = lists.reduce((a, l) => a * l.length, 1); if (total > 100000) { toast(`${total.toLocaleString()} combinations is too many.`); return; } const rows = combinations(lists); addSheetWithRows(app, 'Combinations', rows.map((r) => r as Scalar[])); toast(`${plural(rows.length, 'combination')} listed`); },
  },
  { id: 'plus.bullets', label: 'Add Bullets', icon: 'format_list_bulleted', tab: P, group: G_INSERT, edits: true, kind: 'split', origin: 'Kutools (Insert Bullet)', run: (app, arg?: string) => transformValues(app, 'Bullets', (v) => (String(v).startsWith(arg || '• ') ? undefined : (arg || '• ') + String(v))), menu: () => ['• ', '– ', '▪ ', '◦ ', '✓ ', '→ '].map((b) => ({ id: 'plus.bullets', arg: b, label: `${b}bullet` })) },
  { id: 'plus.numbering', label: 'Add Numbering (1. 2. 3.)', icon: 'format_list_numbered', tab: P, group: G_INSERT, edits: true, kind: 'split', origin: 'Kutools (Insert Numbering)', run: (app, arg?: 'dot' | 'paren' | 'letter' | 'roman') => { const s = app.sheet; let i = 0; const { toRoman } = requireTools(); app.book.transact('Numbering', () => { for (const x of selCells(app, { existingOnly: true })) { if (!hasContent(x.cell) || x.cell!.f != null) continue; i++; const n = arg === 'letter' ? String.fromCharCode(96 + ((i - 1) % 26) + 1) + '.' : arg === 'roman' ? toRoman(i).toLowerCase() + '.' : arg === 'paren' ? `${i})` : `${i}.`; s.set(x.r, x.c, { ...x.cell!, v: `${n} ${shownText(s, x.r, x.c)}` }); } }); }, menu: () => [{ id: 'plus.numbering', arg: 'dot', label: '1. 2. 3.' }, { id: 'plus.numbering', arg: 'paren', label: '1) 2) 3)' }, { id: 'plus.numbering', arg: 'letter', label: 'a. b. c.' }, { id: 'plus.numbering', arg: 'roman', label: 'i. ii. iii.' }] },
  { id: 'plus.batchCheckboxes', label: 'Check Boxes in Every Selected Cell', icon: 'check_box', tab: P, group: G_INSERT, edits: true, popular: true, tip: 'Each selected cell becomes an in-cell check box (unchecked, unless it already holds TRUE)', origin: 'Kutools (Batch Insert Check Boxes)', run: (app) => { const s = app.sheet; app.book.transact('Check Boxes', () => { for (const x of selCells(app, { fill: true })) { const cur = x.cell; s.set(x.r, x.c, { ...(cur || {}), v: cur?.v === true, s: app.book.derivedStyle(s.styleIdOf(x.r, x.c), { checkbox: true, hAlign: 'center' }) }); } }); } },
  { id: 'plus.removeCheckboxes', label: 'Remove Check Boxes', icon: 'check_box_outline_blank', tab: P, group: G_DELETE, edits: true, origin: 'Kutools (Batch Delete Check Boxes)', run: (app) => { const s = app.sheet; app.book.transact('Remove Check Boxes', () => { for (const x of selCells(app, { existingOnly: true })) { if (!s.styleOf(x.r, x.c).checkbox) continue; s.set(x.r, x.c, { ...x.cell!, s: app.book.derivedStyle(x.cell!.s ?? 0, { checkbox: undefined }) || undefined }); } }); } },
  {
    id: 'plus.watermark', label: 'Insert a Watermark…', icon: 'water_drop', tab: P, group: G_INSERT, edits: true, tip: 'Large faint diagonal text (DRAFT, CONFIDENTIAL) behind the cells', origin: 'Kutools (Insert Watermark)',
    run: async (app, arg?: { text: string }) => { const o = await argOrAsk(arg, async () => { const t = await promptDialog('Watermark', 'Text', 'DRAFT'); return t ? { text: t } : null; }); if (!o) return; const { addDrawing } = await import('./objects'); const vis = app.grid.visible(); const d: Omit<DrawingObject, 'id'> = { kind: 'textbox', at: { r: Math.max(0, vis.r1 + 2), c: Math.max(0, vis.c1 + 1), dx: 0, dy: 0, w: 480, h: 120 }, text: o.text, rotate: -30, fill: undefined, stroke: undefined, font: { size: 72, bold: true, color: '#d0d0d0' }, alt: `Watermark: ${o.text}` }; addDrawing(app, d, 'Insert Watermark'); },
  },
  { id: 'plus.datePicker', label: 'Pick a Date…', icon: 'calendar_month', tab: P, group: G_INSERT, edits: true, keys: ['Mod-Alt-;'], tip: 'Choose a date from a calendar control and put it in the active cell', origin: 'Kutools (Date Picker) / XLTools (Popup Calendar)', run: async (app, arg?: string) => { const iso = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Pick a Date', fields: [{ key: 'd', label: 'Date', type: 'date', value: new Date().toISOString().slice(0, 10), autofocus: true }] }); return r?.d ? String(r.d) : null; }); if (!iso) return; const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso); if (!m) return; const { partsToSerial } = await import('../model/dates'); const { r, c } = activeCell(app); app.book.transact('Pick a Date', () => writeOut(app.sheet, r, c, { v: partsToSerial(+m[1], +m[2], +m[3], app.book.calc.date1904), style: app.sheet.styleOf(r, c).numFmt ? {} : { numFmt: 'yyyy-mm-dd' } })); } },
  { id: 'plus.insertRandomPassword', label: 'Insert Random Passwords', icon: 'password', tab: P, group: G_INSERT, edits: true, tip: 'A strong random password (16 characters) in each selected cell', run: (app) => { const alphabet = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789!@#$%&*?'; const gen = () => { const a = new Uint32Array(16); crypto.getRandomValues(a); return [...a].map((x) => alphabet[x % alphabet.length]).join(''); }; app.book.transact('Random Passwords', () => { for (const x of selCells(app, { fill: true })) writeOut(app.sheet, x.r, x.c, gen()); }); } },
  { id: 'plus.menuInsert', label: 'Insert', icon: 'add_box', tab: P, group: G_INSERT, kind: 'menu', noRecord: true, run: () => undefined, menu: () => menuOf(['plus.titleRows', 'plus.pageBreakEvery', 'util.insertRowsEvery', 'util.insertRowAtChange', null, 'plus.findMissing', 'plus.combinations', 'tools.sequence', 'tools.randomData', 'plus.insertRandomPassword', null, 'plus.bullets', 'plus.numbering', 'plus.batchCheckboxes', 'plus.watermark', 'plus.datePicker', 'insert.picture']) },
]);

// ============================================================================================
// Delete

defTools([
  { id: 'plus.deleteBlankRowsBook', label: 'Delete Blank Rows on All Sheets', icon: 'delete_sweep', tab: P, group: G_DELETE, edits: true, origin: 'Kutools (Delete Blank Rows)', run: (app) => { const cur = app.book.activeSheet; let n = 0; for (let i = 0; i < app.book.sheets.length; i++) { const s = app.book.sheets[i]; if (s.protection.on) continue; const u = s.usedRange(false); if (!u) continue; const lines: number[] = []; for (let r = u.r1; r <= u.r2; r++) if (rowIsBlank(s, r)) lines.push(r); if (lines.length) { app.book.transact('Delete Blank Rows', () => deleteRowSpans(s, mergeSpans(lines.map((l) => [l, l] as [number, number])))); n += lines.length; } } app.exec('sheet.activate', cur); app.grid.updateExtent(); toast(`${plural(n, 'blank row')} deleted`); } },
  { id: 'plus.deleteVisibleRows', label: 'Delete the Visible (Filtered) Rows', icon: 'filter_alt_off', tab: P, group: G_DELETE, edits: true, tip: 'Delete the rows a filter shows and keep the hidden ones (the opposite of Delete Hidden Rows)', origin: 'Kutools (Delete Hidden/Visible Rows)', run: (app) => { const s = app.sheet; const rg = selBlock(app); const lines: number[] = []; for (let r = rg.r1; r <= rg.r2; r++) if (!s.hiddenRows.has(r)) lines.push(r); if (!lines.length || sheetProtected(app, 'deleteRows')) return; app.book.transact('Delete Visible Rows', () => deleteRowSpans(s, mergeSpans(lines.map((l) => [l, l] as [number, number])))); app.grid.updateExtent(); app.grid.layout(); toast(`${plural(lines.length, 'row')} deleted`); } },
  { id: 'plus.removeValidation', label: 'Remove Data Validation', icon: 'rule', tab: P, group: G_DELETE, edits: true, kind: 'split', origin: 'Kutools (Clear Data Validation Restrictions)', run: (app, arg?: 'selection' | 'sheet' | 'book') => { const s = app.sheet; let n = 0; app.book.transact('Remove Validation', () => { for (const sh of arg === 'book' ? app.book.sheets : [s]) { app.book.recordSheet(sh); if (arg === 'selection') { const sel = selRanges(app); const before = sh.validations.length; sh.validations = sh.validations.filter((v) => !v.ranges.some((vr) => sel.some((rg) => vr.r1 <= rg.r2 && vr.r2 >= rg.r1 && vr.c1 <= rg.c2 && vr.c2 >= rg.c1))); n += before - sh.validations.length; } else { n += sh.validations.length; sh.validations = []; } } }); toast(`${plural(n, 'validation rule')} removed`); }, menu: () => [{ id: 'plus.removeValidation', arg: 'selection', label: 'From the selection' }, { id: 'plus.removeValidation', arg: 'sheet', label: 'From this sheet' }, { id: 'plus.removeValidation', arg: 'book', label: 'From the whole workbook' }] },
  { id: 'plus.removeLinksBook', label: 'Remove All Hyperlinks in the Workbook', icon: 'link_off', tab: P, group: G_DELETE, edits: true, origin: 'Kutools (Remove Hyperlinks)', run: (app) => { let n = 0; app.book.transact('Remove Hyperlinks', () => { for (const s of app.book.sheets) for (const [k, cell] of [...s.cells]) { if (!cell.link) continue; const { link: _l, linkTip: _t, ...rest } = cell; void _l; void _t; s.set(keyRow(k), keyCol(k), Object.keys(rest).length ? rest : undefined); n++; } }); toast(`${plural(n, 'hyperlink')} removed`); } },
  { id: 'plus.deleteFormControls', label: 'Delete All Form Controls', icon: 'smart_button', tab: P, group: G_DELETE, edits: true, origin: 'Kutools (Delete Web Controls)', run: (app) => { const s = app.sheet; if (sheetProtected(app, 'objects')) return; const n = s.drawings.filter((d) => d.control).length; app.book.transact('Delete Controls', () => { app.book.recordSheet(s); s.drawings = s.drawings.filter((d) => !d.control); }); app.grid.invalidate(); toast(`${plural(n, 'control')} deleted`); } },
  { id: 'plus.menuDelete', label: 'Delete', icon: 'delete', tab: P, group: G_DELETE, kind: 'menu', noRecord: true, run: () => undefined, menu: () => menuOf(['util.deleteEmptyRows', 'plus.deleteBlankRowsBook', 'util.deleteHiddenRows', 'plus.deleteVisibleRows', 'util.deleteEmptyCols', 'util.deleteHiddenCols', null, 'plus.removeValidation', 'plus.removeLinksBook', 'plus.removeCheckboxes', 'plus.deleteFormControls', 'util.deleteAllObjects', 'util.deleteCommentsBook', 'util.removeCFBook', 'util.removeFormattingParts']) },
]);

// ============================================================================================
// Lookup

/** Builds a lookup formula for the Super Lookup wizard. */
export function lookupFormula(o: { kind: 'exact' | 'bottom' | 'multi' | 'many' | 'approx'; lookup: string; keys: string; returns: string; keys2?: string; lookup2?: string; sep?: string; ifNotFound?: string }): string {
  const nf = o.ifNotFound != null && o.ifNotFound !== '' ? `,"${o.ifNotFound.replace(/"/g, '""')}"` : ',""';
  switch (o.kind) {
    case 'exact': return `XLOOKUP(${o.lookup},${o.keys},${o.returns}${nf})`;
    case 'bottom': return `XLOOKUP(${o.lookup},${o.keys},${o.returns}${nf},0,-1)`;
    case 'approx': return `XLOOKUP(${o.lookup},${o.keys},${o.returns}${nf},-1)`;
    case 'multi': return `XLOOKUP(1,(${o.keys}=${o.lookup})*(${o.keys2}=${o.lookup2}),${o.returns}${nf})`;
    case 'many': return `TEXTJOIN("${o.sep || ', '}",TRUE,FILTER(${o.returns},${o.keys}=${o.lookup}${nf}))`;
  }
}

defTools([
  {
    id: 'plus.superLookup', label: 'Lookup Wizard…', icon: 'manage_search', tab: P, group: G_LOOKUP, edits: true, popular: true, tip: 'Write the right lookup formula for you: exact match, last match (from the bottom), two conditions, all matches joined, or nearest value', keywords: 'vlookup xlookup wizard lookup from bottom multiple conditions one to many', origin: 'Kutools (Super LOOKUP) / Ablebits (Vlookup Wizard)',
    run: async (app, arg?: Parameters<typeof lookupFormula>[0]) => {
      const { r, c } = activeCell(app);
      const o = await argOrAsk(arg, async () => { const x = await openDialog({ title: 'Lookup Wizard', help: `The formula goes into ${cellName(r, c)}. Use absolute references ($A$2:$A$100) for the ranges.`, fields: [{ key: 'kind', label: 'Kind of lookup', type: 'select', value: 'exact', options: [{ value: 'exact', label: 'Exact match (first match)' }, { value: 'bottom', label: 'Last match (search from the bottom)' }, { value: 'multi', label: 'Match two conditions' }, { value: 'many', label: 'All matches joined into one cell' }, { value: 'approx', label: 'Nearest value not larger (tax bands, grades)' }] }, { key: 'lookup', label: 'Look for (cell)', type: 'text', value: c > 0 ? cellName(r, c - 1) : 'A2', half: true }, { key: 'keys', label: 'In this column', type: 'text', value: '$A$2:$A$100', half: true }, { key: 'lookup2', label: 'And also for (cell)', type: 'text', value: '', half: true, visibleIf: (v) => v.kind === 'multi' }, { key: 'keys2', label: 'In this column', type: 'text', value: '', half: true, visibleIf: (v) => v.kind === 'multi' }, { key: 'returns', label: 'Return the value from', type: 'text', value: '$B$2:$B$100' }, { key: 'sep', label: 'Separator', type: 'text', value: ', ', visibleIf: (v) => v.kind === 'many' }, { key: 'ifNotFound', label: 'If nothing is found, show', type: 'text', value: '' }] }); return x as any; });
      if (!o) return;
      if (!checkWritable(app, [{ r1: r, c1: c, r2: r, c2: c }])) return;
      const f = lookupFormula(o);
      app.book.transact('Lookup Wizard', () => writeOut(app.sheet, r, c, { f }));
      toast(`=${f}`);
    },
  },
  { id: 'plus.lookupFromBottom', label: 'Lookup from the Bottom', icon: 'vertical_align_bottom', tab: P, group: G_LOOKUP, edits: true, run: (app) => app.exec('plus.superLookup', { kind: 'bottom', lookup: cellName(app.sheet.sel.r, Math.max(0, app.sheet.sel.c - 1)), keys: '$A$2:$A$100', returns: '$B$2:$B$100' }) },
  { id: 'plus.lookupMany', label: 'Lookup All Matches Joined', icon: 'join', tab: P, group: G_LOOKUP, edits: true, run: (app) => app.exec('plus.superLookup', { kind: 'many', lookup: cellName(app.sheet.sel.r, Math.max(0, app.sheet.sel.c - 1)), keys: '$A$2:$A$100', returns: '$B$2:$B$100', sep: ', ' }) },
  { id: 'plus.menuLookup', label: 'Lookup', icon: 'manage_search', tab: P, group: G_LOOKUP, kind: 'menu', noRecord: true, run: () => undefined, menu: () => menuOf(['plus.superLookup', 'plus.lookupFromBottom', 'plus.lookupMany', 'plus.mergeTables', 'tools.fuzzyMatch', 'util.replaceFromList', 'formulas.insertFunction']) },
]);

// ============================================================================================
// Filter

export interface FilterCond { col: number; op: string; value: string; value2?: string }

export function rowPasses(s: Sheet, r: number, conds: FilterCond[], logic: 'and' | 'or'): boolean {
  const test = (cd: FilterCond) => {
    const v = s.value(r, cd.col); const t = shownText(s, r, cd.col); const low = t.toLowerCase(); const needle = cd.value.toLowerCase(); const n = numOf(v); const x = Number(cd.value), y = Number(cd.value2);
    switch (cd.op) {
      case 'equals': return low === needle || (n != null && !isNaN(x) && n === x);
      case 'notEquals': return !(low === needle || (n != null && !isNaN(x) && n === x));
      case 'contains': return low.includes(needle);
      case 'notContains': return !low.includes(needle);
      case 'starts': return low.startsWith(needle);
      case 'ends': return low.endsWith(needle);
      case 'gt': return n != null && n > x; case 'ge': return n != null && n >= x; case 'lt': return n != null && n < x; case 'le': return n != null && n <= x;
      case 'between': return n != null && n >= Math.min(x, y) && n <= Math.max(x, y);
      case 'blank': return t === ''; case 'notBlank': return t !== '';
      case 'bold': return !!s.styleOf(r, cd.col).bold; case 'formula': return s.get(r, cd.col)?.f != null; case 'comment': return s.comments.has(cellKey(r, cd.col)); case 'merged': return !!s.mergeAt(r, cd.col);
      case 'fill': return !!s.styleOf(r, cd.col).fill && (!cd.value || (s.styleOf(r, cd.col).fill || '').toLowerCase() === needle); case 'error': return v instanceof CellError;
      case 'upper': return t !== '' && t === t.toUpperCase() && t !== t.toLowerCase(); case 'lower': return t !== '' && t === t.toLowerCase() && t !== t.toUpperCase();
      case 'lenGt': return t.length > x; case 'lenLt': return t.length < x;
      case 'regex': try { return new RegExp(cd.value, 'i').test(t); } catch { return false; }
      default: return true;
    }
  };
  return logic === 'and' ? conds.every(test) : conds.some(test);
}

const FILTER_OPS = [{ value: 'equals', label: 'equals' }, { value: 'notEquals', label: 'does not equal' }, { value: 'contains', label: 'contains' }, { value: 'notContains', label: 'does not contain' }, { value: 'starts', label: 'starts with' }, { value: 'ends', label: 'ends with' }, { value: 'gt', label: '>' }, { value: 'ge', label: '≥' }, { value: 'lt', label: '<' }, { value: 'le', label: '≤' }, { value: 'between', label: 'between' }, { value: 'blank', label: 'is blank' }, { value: 'notBlank', label: 'is not blank' }, { value: 'regex', label: 'matches regex' }, { value: 'lenGt', label: 'text longer than' }, { value: 'lenLt', label: 'text shorter than' }, { value: 'bold', label: 'is bold' }, { value: 'formula', label: 'is a formula' }, { value: 'comment', label: 'has a comment' }, { value: 'merged', label: 'is merged' }, { value: 'fill', label: 'has a fill color' }, { value: 'error', label: 'is an error' }, { value: 'upper', label: 'is UPPERCASE' }, { value: 'lower', label: 'is lowercase' }];

function hideRows(app: App, rg: Range, keep: (r: number) => boolean, label: string, first: number): number {
  const s = app.sheet; let shown = 0;
  app.book.transact(label, () => { app.book.recordSheet(s); for (let r = first; r <= rg.r2; r++) { if (keep(r)) { s.hiddenRows.delete(r); shown++; } else s.hiddenRows.add(r); } });
  app.grid.layout();
  return shown;
}

defTools([
  {
    id: 'plus.superFilter', label: 'Filter Rows by Several Conditions…', icon: 'filter_list', tab: P, group: G_FILTER, edits: true, popular: true, tip: 'Up to four conditions on different columns, combined with AND or OR, including text length, case, bold, formulas, comments and regular expressions; rows that fail are hidden', keywords: 'super filter multiple criteria and or', origin: 'Kutools (Super Filter / Special Filter)',
    run: async (app, arg?: { conds: FilterCond[]; logic: 'and' | 'or' }) => {
      if (emptySheet(app)) return;
      const s = app.sheet; const rg = selBlock(app); const header = headerGuess(s, rg); const cols = columnChoices(s, rg, header);
      const o = await argOrAsk(arg, async () => { const fields: any[] = []; for (let i = 1; i <= 4; i++) fields.push({ key: `c${i}`, label: `Column ${i}`, type: 'select', value: i === 1 ? String(rg.c1) : '', options: [{ value: '', label: '(none)' }, ...cols], half: true }, { key: `o${i}`, label: 'Condition', type: 'select', value: 'contains', options: FILTER_OPS, half: true }, { key: `v${i}`, label: 'Value', type: 'text', value: '', half: true }, { key: `w${i}`, label: 'And', type: 'text', value: '', half: true, visibleIf: (v: any) => v[`o${i}`] === 'between' }); fields.push({ key: 'logic', label: 'Combine', type: 'radio', value: 'and', options: [{ value: 'and', label: 'All conditions (AND)' }, { value: 'or', label: 'Any condition (OR)' }] }); const r = await openDialog({ title: 'Filter Rows by Several Conditions', wide: true, fields }); if (!r) return null; const conds: FilterCond[] = []; for (let i = 1; i <= 4; i++) if (r[`c${i}`] !== '') conds.push({ col: Number(r[`c${i}`]), op: r[`o${i}`], value: String(r[`v${i}`] ?? ''), value2: String(r[`w${i}`] ?? '') }); return { conds, logic: r.logic }; });
      if (!o || !o.conds.length) return;
      const shown = hideRows(app, rg, (r) => rowPasses(s, r, o.conds, o.logic), 'Filter Rows', rg.r1 + (header ? 1 : 0));
      toast(`${shown} of ${plural(rg.r2 - rg.r1 + 1 - (header ? 1 : 0), 'row')} shown`);
    },
  },
  { id: 'plus.filterBold', label: 'Show Only Rows with a Bold Key Cell', icon: 'format_bold', tab: P, group: G_FILTER, edits: true, origin: 'Kutools (Filter Bold)', run: (app) => { const s = app.sheet; const rg = selBlock(app); const c = activeCell(app).c; hideRows(app, rg, (r) => !!s.styleOf(r, c).bold, 'Filter Bold', rg.r1 + 1); } },
  { id: 'plus.filterComments', label: 'Show Only Rows with a Comment', icon: 'comment', tab: P, group: G_FILTER, edits: true, origin: 'Kutools (Filter Comment)', run: (app) => { const s = app.sheet; const rg = selBlock(app); const c = activeCell(app).c; hideRows(app, rg, (r) => s.comments.has(cellKey(r, c)), 'Filter Comments', rg.r1 + 1); } },
  { id: 'plus.filterFormulas', label: 'Show Only Rows with a Formula', icon: 'function', tab: P, group: G_FILTER, edits: true, origin: 'Kutools (Filter Formula)', run: (app) => { const s = app.sheet; const rg = selBlock(app); const c = activeCell(app).c; hideRows(app, rg, (r) => s.get(r, c)?.f != null, 'Filter Formulas', rg.r1 + 1); } },
  { id: 'plus.filterColor', label: 'Show Only Rows with the Active Cell\'s Fill', icon: 'format_color_fill', tab: P, group: G_FILTER, edits: true, origin: 'Kutools (Filter by Color)', run: (app) => { const s = app.sheet; const rg = selBlock(app); const { r: ar, c } = activeCell(app); const fill = (s.styleOf(ar, c).fill || '').toLowerCase(); hideRows(app, rg, (r) => (s.styleOf(r, c).fill || '').toLowerCase() === fill, 'Filter by Color', rg.r1 + 1); } },
  { id: 'plus.showAllRows', label: 'Show All Rows Again', icon: 'visibility', tab: P, group: G_FILTER, edits: true, run: (app) => { const s = app.sheet; app.book.transact('Show All Rows', () => { app.book.recordSheet(s); s.hiddenRows.clear(); }); app.grid.layout(); } },
  {
    id: 'plus.subtotalEvery', label: 'Subtotal Every N Rows…', icon: 'functions', tab: P, group: G_FILTER, edits: true, tip: 'Insert a SUM row after every N rows of the selected numeric columns (page subtotals)', origin: 'Kutools (Paging Subtotals)',
    run: async (app, arg?: { every: number }) => { const o = await argOrAsk(arg, async () => { const t = await promptDialog('Subtotal Every N Rows', 'Every … rows', '25', { type: 'number' }); return t ? { every: Number(t) || 25 } : null; }); if (!o || o.every < 1) return; const s = app.sheet; const rg = selBlock(app); const header = headerGuess(s, rg); if (sheetProtected(app, 'insertRows')) return; const first = rg.r1 + (header ? 1 : 0); const spots: number[] = []; for (let r = first + o.every; r <= rg.r2 + 1; r += o.every) spots.push(r); if (spots[spots.length - 1] !== rg.r2 + 1) spots.push(rg.r2 + 1); app.book.transact('Paging Subtotals', () => { let offset = 0; for (const at0 of spots) { const at = at0 + offset; insertRows(s, at, 1); const from = at - Math.min(o.every, at - first - offset * 0); for (let c = rg.c1; c <= rg.c2; c++) { let nums = 0; for (let r = Math.max(first, at - o.every); r < at; r++) if (typeof s.value(r, c) === 'number') nums++; if (nums) writeOut(s, at, c, { f: `SUBTOTAL(9,${cellName(Math.max(first, at - o.every), c)}:${cellName(at - 1, c)})`, style: { bold: true, bt: { style: 'thin' } } }); } void from; offset++; } }); app.grid.updateExtent(); app.grid.layout(); },
  },
  { id: 'plus.menuFilter', label: 'Filter', icon: 'filter_list', tab: P, group: G_FILTER, kind: 'menu', noRecord: true, run: () => undefined, menu: () => menuOf(['plus.superFilter', 'plus.filterBold', 'plus.filterComments', 'plus.filterFormulas', 'plus.filterColor', 'plus.showAllRows', null, 'data.filter', 'data.advancedFilter', 'data.filterBySelected', 'util.condRows', 'plus.subtotalEvery', 'data.subtotal']) },
]);

// ============================================================================================
// Print

defTools([
  { id: 'plus.printSelections', label: 'Print the Selected Areas Only', icon: 'print', tab: P, group: G_PRINT, noRecord: true, tip: 'Set the print area to the selected areas (several are allowed) and print', origin: 'Kutools (Print Multiple Selection Wizard)', run: (app) => { const s = app.sheet; const areas = selRanges(app).map((rg) => rangeName(clipToUsed(s, rg))); setPrint(app, 'Print Area', (p) => { p.area = areas.join(','); }); app.exec('file.print'); } },
  { id: 'plus.printCurrentPage', label: 'Print Only the Page of the Active Cell', icon: 'print', tab: P, group: G_PRINT, noRecord: true, origin: 'Kutools (Print Current Page)', run: async (app) => { const { computePageBreaks } = await import('../print/pages'); const s = app.sheet; const pb = computePageBreaks(s); const { r, c } = activeCell(app); const rows = [pb.area.r1, ...pb.rows, pb.area.r2 + 1], cols = [pb.area.c1, ...pb.cols, pb.area.c2 + 1]; let r1 = rows[0], r2 = rows[rows.length - 1] - 1, c1 = cols[0], c2 = cols[cols.length - 1] - 1; for (let i = 0; i < rows.length - 1; i++) if (r >= rows[i] && r < rows[i + 1]) { r1 = rows[i]; r2 = rows[i + 1] - 1; } for (let i = 0; i < cols.length - 1; i++) if (c >= cols[i] && c < cols[i + 1]) { c1 = cols[i]; c2 = cols[i + 1] - 1; } const old = s.print.area; setPrint(app, 'Print Page', (p) => { p.area = rangeName({ r1, c1, r2, c2 }); }); await app.exec('file.print'); setPrint(app, 'Print Area', (p) => { p.area = old; }); } },
  { id: 'plus.borderEachPage', label: 'Border Around Each Printed Page', icon: 'border_outer', tab: P, group: G_PRINT, edits: true, origin: 'Kutools (Add Border to Each Page)', run: async (app) => { const { computePageBreaks } = await import('../print/pages'); const s = app.sheet; const pb = computePageBreaks(s); const rows = [pb.area.r1, ...pb.rows, pb.area.r2 + 1], cols = [pb.area.c1, ...pb.cols, pb.area.c2 + 1]; const b = { style: 'medium' as const, color: '#333333' }; app.book.transact('Page Borders', () => { for (let i = 0; i < rows.length - 1; i++) for (let j = 0; j < cols.length - 1; j++) { const rg = { r1: rows[i], r2: rows[i + 1] - 1, c1: cols[j], c2: cols[j + 1] - 1 }; for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) { const p: Partial<CellStyle> = {}; if (r === rg.r1) p.bt = b; if (r === rg.r2) p.bb = b; if (c === rg.c1) p.bl = b; if (c === rg.c2) p.br = b; if (Object.keys(p).length) s.set(r, c, { ...(s.get(r, c) || {}), s: app.book.derivedStyle(s.styleIdOf(r, c), p) }); } } }); toast(`${(rows.length - 1) * (cols.length - 1)} pages outlined`); } },
  { id: 'plus.removePageBreaks', label: 'Remove All Manual Page Breaks', icon: 'insert_page_break', tab: P, group: G_PRINT, edits: true, run: (app) => { const s = app.sheet; app.book.transact('Remove Page Breaks', () => { app.book.recordSheet(s); s.rowBreaks = []; s.colBreaks = []; }); app.grid.invalidate(); } },
  { id: 'plus.menuPrint', label: 'Print', icon: 'print', tab: P, group: G_PRINT, kind: 'menu', noRecord: true, run: () => undefined, menu: () => menuOf(['plus.printSelections', 'plus.printCurrentPage', 'plus.borderEachPage', 'plus.pageBreakEvery', 'plus.removePageBreaks', null, 'util.fitOnePage', 'util.printTitlesAll', 'util.pageSetupToSheets', 'util.headerFooterAll', 'file.printPreview', 'layout.pageSetup']) },
]);

// ============================================================================================
// Security

const ENC_PREFIX = '\u{1F512}';

async function deriveKey(password: string, salt: Uint8Array): Promise<CryptoKey> {
  const base = await crypto.subtle.importKey('raw', new TextEncoder().encode(password), 'PBKDF2', false, ['deriveKey']);
  return crypto.subtle.deriveKey({ name: 'PBKDF2', salt: salt as BufferSource, iterations: 150000, hash: 'SHA-256' }, base, { name: 'AES-GCM', length: 256 }, false, ['encrypt', 'decrypt']);
}
const b64 = (u: Uint8Array) => btoa(String.fromCharCode(...u));
const unb64 = (s: string) => Uint8Array.from(atob(s), (ch) => ch.charCodeAt(0));

/** Marks an encrypted constant that carries its exact value and type (the number itself, not its rounded display text). */
const ENC_TYPED = '\u0001v:';

export async function encryptText(text: string, password: string): Promise<string> {
  const salt = crypto.getRandomValues(new Uint8Array(16)), iv = crypto.getRandomValues(new Uint8Array(12));
  const key = await deriveKey(password, salt);
  const data = new Uint8Array(await crypto.subtle.encrypt({ name: 'AES-GCM', iv }, key, new TextEncoder().encode(text)));
  return `${ENC_PREFIX}${b64(salt)}.${b64(iv)}.${b64(data)}`;
}
export async function decryptText(text: string, password: string): Promise<string | null> {
  if (!text.startsWith(ENC_PREFIX)) return null;
  const [s, i, d] = text.slice(ENC_PREFIX.length).split('.');
  try { const key = await deriveKey(password, unb64(s)); const out = await crypto.subtle.decrypt({ name: 'AES-GCM', iv: unb64(i) }, key, unb64(d)); return new TextDecoder().decode(out); } catch { return null; }
}

defTools([
  {
    id: 'plus.encryptCells', label: 'Encrypt the Selected Cells…', icon: 'enhanced_encryption', tab: P, group: G_SECURITY, edits: true, popular: true, tip: 'Replace the selected values with AES-256 encrypted text that only Decrypt Cells with the password can read (the formulas and values are not kept anywhere else)', keywords: 'encrypt password cells aes', origin: 'Kutools (Encrypt Cells)',
    run: async (app, arg?: { password: string }) => { const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Encrypt Cells', help: 'Formulas are encrypted as their text. Lost passwords cannot be recovered.', fields: [{ key: 'password', label: 'Password', type: 'password', value: '', autofocus: true }, { key: 'again', label: 'Confirm', type: 'password', value: '' }] }); if (!r) return null; if (!r.password || r.password !== r.again) { toast('The passwords do not match.'); return null; } return { password: String(r.password) }; }); if (!o) return; const s = app.sheet; const jobs: { r: number; c: number; cell: Cell; text: string }[] = []; for (const x of selCells(app, { existingOnly: true })) { const cell = x.cell!; if (!hasContent(cell) || !canChange(s, x.r, x.c)) continue; if (typeof cell.v === 'string' && cell.f == null && cell.v.startsWith(ENC_PREFIX)) continue; const text = cell.f != null ? '=' + cell.f : typeof cell.v === 'string' || typeof cell.v === 'number' || typeof cell.v === 'boolean' ? ENC_TYPED + JSON.stringify(cell.v) : shownText(s, x.r, x.c); jobs.push({ r: x.r, c: x.c, cell, text }); } const enc = await Promise.all(jobs.map((j) => encryptText(j.text, o.password))); app.book.transact('Encrypt Cells', () => jobs.forEach((j, i) => s.set(j.r, j.c, { ...(j.cell.s ? { s: j.cell.s } : {}), v: enc[i] }))); toast(`${plural(jobs.length, 'cell')} encrypted`); },
  },
  {
    id: 'plus.decryptCells', label: 'Decrypt the Selected Cells…', icon: 'no_encryption', tab: P, group: G_SECURITY, edits: true, origin: 'Kutools (Decrypt Cells)',
    run: async (app, arg?: { password: string }) => { const o = await argOrAsk(arg, async () => { const p = await promptDialog('Decrypt Cells', 'Password', '', { type: 'password' }); return p == null ? null : { password: p }; }); if (!o) return; const s = app.sheet; const jobs = selCells(app, { existingOnly: true }).filter((x) => typeof x.cell?.v === 'string' && x.cell.v.startsWith(ENC_PREFIX)); const dec = await Promise.all(jobs.map((j) => decryptText(j.cell!.v as string, o.password))); let n = 0, bad = 0; app.book.transact('Decrypt Cells', () => jobs.forEach((j, i) => { const t = dec[i]; if (t == null) { bad++; return; } if (t.startsWith(ENC_TYPED)) { let v: unknown; try { v = JSON.parse(t.slice(ENC_TYPED.length)); } catch { bad++; return; } s.set(j.r, j.c, { ...(j.cell!.s ? { s: j.cell!.s } : {}), v: v as string | number | boolean }); } else setInput(app, s, j.r, j.c, t); n++; })); toast(`${plural(n, 'cell')} decrypted${bad ? `; ${bad} failed (wrong password?)` : ''}`); },
  },
  { id: 'plus.hideFormulas', label: 'Hide the Formulas When Protected', icon: 'visibility_off', tab: P, group: G_SECURITY, edits: true, tip: 'Mark the selected cells so their formulas are not shown in the formula bar once the sheet is protected', origin: 'Kutools (Hide Formulas)', run: (app) => patchStyle(app, 'Hide Formulas', { hidden: true }) },
  { id: 'plus.unhideFormulas', label: 'Show the Formulas When Protected', icon: 'visibility', tab: P, group: G_SECURITY, edits: true, run: (app) => patchStyle(app, 'Unhide Formulas', { hidden: undefined }) },
  {
    id: 'plus.mailingLinks', label: 'Create Email Links from Rows…', icon: 'mail', tab: P, group: G_SECURITY, edits: true, tip: 'From a table with an address column, build a mailto: link per row with a subject and a body that can use {Column} placeholders; the links open your mail program', keywords: 'mail merge email send mailto', origin: 'Kutools (Create Mailing List / Send Emails)',
    run: async (app, arg?: { to: number; subject: string; body: string }) => {
      if (emptySheet(app)) return;
      const s = app.sheet; const rg = selBlock(app); const header = headerGuess(s, rg); const cols = columnChoices(s, rg, header);
      const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Create Email Links', help: 'Use {Header} in the subject and body to insert the row\'s values.', fields: [{ key: 'to', label: 'Address column', type: 'select', value: String(rg.c1), options: cols }, { key: 'subject', label: 'Subject', type: 'text', value: '' }, { key: 'body', label: 'Body', type: 'textarea', rows: 5, value: '' }] }); return r ? { to: Number(r.to), subject: String(r.subject || ''), body: String(r.body || '') } : null; });
      if (!o) return;
      const names = Array.from({ length: rg.c2 - rg.c1 + 1 }, (_x, j) => (header ? shownText(s, rg.r1, rg.c1 + j) : colName(rg.c1 + j)));
      const target = { r1: rg.r1, r2: rg.r2, c1: rg.c2 + 1, c2: rg.c2 + 1 }; if (!checkWritable(app, [target])) return;
      let n = 0;
      app.book.transact('Email Links', () => { for (let r = rg.r1 + (header ? 1 : 0); r <= rg.r2; r++) { const to = shownText(s, r, o.to).trim(); if (!to.includes('@')) continue; const fill = (t: string) => t.replace(/\{([^}]+)\}/g, (_m, nm: string) => { const j = names.findIndex((x) => x.toLowerCase() === nm.toLowerCase()); return j >= 0 ? shownText(s, r, rg.c1 + j) : ''; }); const link = `mailto:${to}?subject=${encodeURIComponent(fill(o.subject))}&body=${encodeURIComponent(fill(o.body))}`; s.set(r, target.c1, { v: 'Email', link, s: app.book.styleId({ color: '#0563c1', underline: 'single' }) }); n++; } if (header) writeOut(s, rg.r1, target.c1, { v: 'Email link', style: HEAD }); });
      toast(`${plural(n, 'email link')} created`);
    },
  },
  { id: 'plus.menuSecurity', label: 'Security', icon: 'security', tab: P, group: G_SECURITY, kind: 'menu', noRecord: true, run: () => undefined, menu: () => menuOf(['plus.encryptCells', 'plus.decryptCells', 'plus.hideFormulas', 'plus.unhideFormulas', null, 'util.protectSheets', 'util.unprotectSheets', 'review.protectWorkbook', 'file.encrypt', 'util.lockNonEmpty', 'util.lockFormulas', null, 'plus.mailingLinks']) },
]);

// ============================================================================================
// Workbook

defTools([
  { id: 'plus.splitWorkbook', label: 'Split the Workbook into Files…', icon: 'folder_zip', tab: P, group: G_BOOK, noRecord: true, origin: 'Kutools (Split Workbook)', run: (app) => app.exec('util.exportSheetsFiles') },
  {
    id: 'plus.formatConverter', label: 'Convert Files to Another Format…', icon: 'transform', tab: P, group: G_BOOK, noRecord: true, popular: true, tip: 'Pick several spreadsheet files (any format Anne\'s Grid opens, including .xls, .ods, Lotus and Quattro) and save each as .xlsx, .ods, .csv or another format', keywords: 'batch convert xls xlsx ods csv', origin: 'Kutools (Format Converter)',
    run: async (app, arg?: { format: string }) => {
      const { FORMATS, exportFile } = await import('../io');
      const { pickFiles } = await import('./utilShared');
      const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Convert Files', fields: [{ key: 'format', label: 'Save each file as', type: 'select', value: 'xlsx', options: FORMATS.filter((f) => f.canExport && ['xlsx', 'ods', 'csv', 'tsv', 'html', 'json', 'qgw', 'fods'].includes(f.id)).map((f) => ({ value: f.id, label: f.label })) }] }); return r ? { format: r.format } : null; });
      if (!o) return;
      const files = await pickFiles('*/*', true); if (!files.length) return;
      const fmt = FORMATS.find((f) => f.id === o.format)!; const { saveAs } = await import('../platform/native');
      let n = 0;
      for (const f of files) { try { const res = await importFile(await f.arrayBuffer(), f.name); const blob = await exportFile(res.book, fmt.id, { sheetIndex: 0 }); const r = await saveAs(blob, `${fileStem(f.name)}.${fmt.ext[0]}`); if (!r) break; n++; } catch (e) { toast(`${f.name}: ${(e as Error).message}`); } }
      toast(`${plural(n, 'file')} converted`);
      void app;
    },
  },
  { id: 'plus.autoOpenList', label: 'Pin This Workbook to Recent', icon: 'push_pin', tab: P, group: G_BOOK, noRecord: true, tip: 'Keep this workbook at the top of the Open ▸ Recent list', origin: 'Kutools (Auto Open Those Workbooks Next Time)', run: async (app) => { const { docKey } = await import('./file'); const key = docKey(app.tab); const e = app.prefs.recent.find((r) => r.key === key); if (!e) { toast('Save the workbook first.'); return; } e.pinned = !e.pinned; app.savePrefs(); toast(e.pinned ? 'Pinned' : 'Unpinned'); } },
  { id: 'plus.menuWorkbook', label: 'Workbook', icon: 'folder', tab: P, group: G_BOOK, kind: 'menu', noRecord: true, run: () => undefined, menu: () => menuOf(['plus.splitWorkbook', 'plus.formatConverter', 'util.sheetsToNewWorkbook', 'util.combineSheets', 'util.importAsSheets', 'plus.autoOpenList', 'util.saveBackup', 'util.copyPath', 'util.insertFileName', 'util.openContainingFolder', 'util.cleanWorkbook']) },
]);

// ============================================================================================
// Inquire

function sheetRefsIn(f: string): string[] { const out: string[] = []; for (const m of f.replace(/"(?:[^"]|"")*"/g, '').matchAll(/(?:'([^']+)'|([A-Za-z_][\w.]*))!/g)) out.push(m[1] || m[2]); return out; }

defTools([
  {
    id: 'plus.workbookAnalysis', label: 'Workbook Analysis Report', icon: 'analytics', tab: P, group: G_INQUIRE, edits: true, popular: true, tip: 'A structured report: sheets (hidden, very hidden, protected), formulas by kind (with errors, array, external, circular risk), names, validations, conditional formats, comments, objects, merged cells, links', keywords: 'inquire workbook analysis audit', origin: 'Inquire (Workbook Analysis)',
    run: (app) => {
      const book = app.book; const rows: Scalar[][] = [];
      const add = (cat: string, item: string, n: number | string, detail = '') => rows.push([cat, item, n, detail]);
      add('Workbook', 'Sheets', book.sheets.length); add('Workbook', 'Hidden sheets', book.sheets.filter((s) => s.visibility === 'hidden').length, book.sheets.filter((s) => s.visibility === 'hidden').map((s) => s.name).join(', ')); add('Workbook', 'Very hidden sheets', book.sheets.filter((s) => s.visibility === 'veryHidden').length); add('Workbook', 'Protected sheets', book.sheets.filter((s) => s.protection.on).length); add('Workbook', 'Structure protected', book.structureLocked ? 'Yes' : 'No'); add('Workbook', 'Defined names', book.names.length); add('Workbook', 'Names with errors', book.names.filter((n) => /#REF!/.test(n.ref)).length); add('Workbook', 'Hidden names', book.names.filter((n) => n.hidden).length); add('Workbook', 'Macro modules', book.modules.length); add('Workbook', 'Custom styles', book.cellStyles.filter((c) => !c.builtin).length);
      let cells = 0, formulas = 0, errors = 0, external = 0, arrays = 0, text = 0, numbers = 0, dates = 0, links = 0, longF = 0, cross = 0, blankFormulas = 0;
      const fnCount = new Map<string, number>();
      for (const s of book.sheets) for (const cell of s.cells.values()) { if (!hasContent(cell)) continue; cells++; if (cell.link) links++; if (cell.f != null) { formulas++; if (cell.v instanceof CellError) errors++; if (externalLinkIn(cell.f)) external++; if (cell.arr) arrays++; if (cell.f.length > 200) longF++; if (sheetRefsIn(cell.f).length) cross++; if (cell.v == null || cell.v === '') blankFormulas++; for (const m of cell.f.matchAll(/([A-Z][A-Z0-9._]*)\(/g)) fnCount.set(m[1], (fnCount.get(m[1]) || 0) + 1); } else if (typeof cell.v === 'string') text++; else if (typeof cell.v === 'number') numbers++; }
      void dates;
      add('Cells', 'Cells with content', cells); add('Cells', 'Text', text); add('Cells', 'Numbers', numbers); add('Cells', 'Hyperlinks', links);
      add('Formulas', 'Formulas', formulas); add('Formulas', 'With error values', errors); add('Formulas', 'Referring to other sheets', cross); add('Formulas', 'Referring to other workbooks', external); add('Formulas', 'Legacy array formulas', arrays); add('Formulas', 'Longer than 200 characters', longF); add('Formulas', 'Returning blank', blankFormulas);
      for (const [fn, n] of [...fnCount].sort((a, b) => b[1] - a[1]).slice(0, 25)) add('Functions', fn, n);
      for (const s of book.sheets) { add('Sheet: ' + s.name, 'Used range', s.usedRange(false) ? rangeName(s.usedRange(false)!) : '(empty)'); add('Sheet: ' + s.name, 'Data validation rules', s.validations.length); add('Sheet: ' + s.name, 'Conditional format rules', s.cf.length); add('Sheet: ' + s.name, 'Comments and notes', s.comments.size); add('Sheet: ' + s.name, 'Merged ranges', s.merges.length); add('Sheet: ' + s.name, 'Charts', s.charts.length); add('Sheet: ' + s.name, 'Pictures and shapes', s.drawings.length); add('Sheet: ' + s.name, 'Tables', s.tables.length); add('Sheet: ' + s.name, 'PivotTables', s.pivots.length); add('Sheet: ' + s.name, 'Hidden rows / columns', `${s.hiddenRows.size} / ${s.hiddenCols.size}`); }
      report(app, 'Workbook Analysis', ['Category', 'Item', 'Count', 'Details'], rows, { colW: [160, 220, 90, 320] });
    },
  },
  {
    id: 'plus.sheetRelationships', label: 'Sheet Relationship Report', icon: 'hub', tab: P, group: G_INQUIRE, edits: true, tip: 'Which sheets refer to which, through formulas, names, charts and validation, with the number of references', origin: 'Inquire (Worksheet Relationship)',
    run: (app) => { const book = app.book; const m = new Map<string, number>(); const bump = (from: string, to: string) => { const k = `${from}\u0001${to}`; m.set(k, (m.get(k) || 0) + 1); }; for (const s of book.sheets) { for (const cell of s.cells.values()) if (cell.f != null) for (const t of sheetRefsIn(cell.f)) if (t.toLowerCase() !== s.name.toLowerCase()) bump(s.name, t); for (const ch of s.charts) for (const se of ch.series) for (const t of sheetRefsIn(se.values + (se.categories || ''))) bump(`${s.name} (chart)`, t); for (const dv of s.validations) for (const t of sheetRefsIn((dv.formula1 || '') + (dv.formula2 || ''))) bump(`${s.name} (validation)`, t); for (const pv of s.pivots) for (const t of sheetRefsIn(pv.source)) bump(`${s.name} (pivot)`, t); } for (const n of book.names) for (const t of sheetRefsIn(n.ref)) bump(`Name: ${n.name}`, t); if (!m.size) { toast('No cross-sheet references.'); return; } report(app, 'Sheet Relationships', ['From', 'To', 'References', 'Exists'], [...m].map(([k, n]) => { const [a, b] = k.split('\u0001'); return [a, b, n, book.sheetByName(b) ? 'Yes' : 'MISSING'] as Scalar[]; }).sort((x, y) => Number(y[2]) - Number(x[2]))); },
  },
  {
    id: 'plus.cellRelationship', label: 'Cell Relationship Report', icon: 'account_tree', tab: P, group: G_INQUIRE, edits: true, tip: 'For the active cell: everything it depends on and everything that depends on it, two levels deep, as a list', origin: 'Inquire (Cell Relationship)',
    run: (app) => { const s = app.sheet; const { r, c } = activeCell(app); const eng = app.engine; const rows: Scalar[][] = []; const seen = new Set<string>(); const walk = (sh: Sheet, rr: number, cc: number, dir: 'precedent' | 'dependent', depth: number) => { if (depth > 2) return; const list = dir === 'precedent' ? eng.precedentsOf(sh, rr, cc).map((p) => ({ sheet: p.sheet, r: p.r1, c: p.c1, r2: p.r2, c2: p.c2 })) : eng.dependentsOfCell(sh, rr, cc).map((d) => ({ sheet: d.sheet, r: d.r, c: d.c, r2: d.r, c2: d.c })); for (const p of list) { const key = `${dir}:${p.sheet.id}:${p.r}:${p.c}`; if (seen.has(key)) continue; seen.add(key); const addr = p.r === p.r2 && p.c === p.c2 ? cellName(p.r, p.c) : rangeName({ r1: p.r, c1: p.c, r2: p.r2, c2: p.c2 }); const cell = p.sheet.get(p.r, p.c); rows.push([dir === 'precedent' ? `Depends on (level ${depth})` : `Used by (level ${depth})`, `${p.sheet.name}!${addr}`, cell?.f != null ? '=' + cell.f : '', cell?.v instanceof CellError ? cell.v.code : cell?.v ?? null]); if (p.r === p.r2 && p.c === p.c2) walk(p.sheet, p.r, p.c, dir, depth + 1); } }; walk(s, r, c, 'precedent', 1); walk(s, r, c, 'dependent', 1); if (!rows.length) { toast('The active cell has no precedents or dependents.'); return; } report(app, `Relationships ${cellName(r, c)}`, ['Relation', 'Cell', 'Formula', 'Value'], [[`Cell`, `${s.name}!${cellName(r, c)}`, s.get(r, c)?.f != null ? '=' + s.get(r, c)!.f : '', s.value(r, c)], ...rows], { colW: [160, 140, 320, 120] }); },
  },
  {
    id: 'plus.compareFiles', label: 'Compare with Another Workbook File…', icon: 'difference', tab: P, group: G_INQUIRE, edits: true, popular: true, tip: 'Open another workbook file and list every cell that differs from this one, sheet by sheet (values and formulas)', keywords: 'compare files versions differences diff', origin: 'Inquire (Compare Files) / Ablebits (Compare Two Sheets)',
    run: async (app) => { const { pickFiles } = await import('./utilShared'); const files = await pickFiles('.xlsx,.xlsm,.xls,.ods,.csv,.qgw,.fods,.tsv,.json', false); if (!files.length) return; const res = await importFile(await files[0].arrayBuffer(), files[0].name); const other = res.book; const rows: Scalar[][] = []; for (const s of app.book.sheets) { const o = other.sheetByName(s.name); if (!o) { rows.push([s.name, '', 'Sheet missing in the other file', '', '']); continue; } for (const d of compareSheetCells(s, o, true)) rows.push([s.name, cellName(d.r, d.c), d.kind, textOf(d.a), textOf(d.b)]); } for (const o of other.sheets) if (!app.book.sheetByName(o.name)) rows.push([o.name, '', 'Sheet only in the other file', '', '']); if (!rows.length) { toast('The workbooks are identical.'); return; } report(app, 'File Comparison', ['Sheet', 'Cell', 'Difference', 'This workbook', files[0].name], rows, { colW: [110, 60, 160, 220, 220] }); },
  },
  { id: 'plus.menuInquire', label: 'Inquire', icon: 'troubleshoot', tab: P, group: G_INQUIRE, kind: 'menu', noRecord: true, run: () => undefined, menu: () => menuOf(['plus.workbookAnalysis', 'plus.sheetRelationships', 'plus.cellRelationship', 'plus.compareFiles', 'tools.compareSheets', 'util.findProblems', 'util.cleanWorkbook', 'util.listNames', 'util.formulasReport', 'formulas.dependencyDiagram', 'formulas.tracePrecedents', 'formulas.traceDependents']) },
]);

// ============================================================================================
// Query: a small SQL for sheets

type Tok = { t: 'id' | 'num' | 'str' | 'op' | 'kw' | 'punct' | 'end'; v: string };
const KEYWORDS = new Set(['SELECT', 'FROM', 'WHERE', 'GROUP', 'BY', 'ORDER', 'LIMIT', 'AND', 'OR', 'NOT', 'LIKE', 'IN', 'BETWEEN', 'IS', 'NULL', 'AS', 'ASC', 'DESC', 'DISTINCT', 'HAVING', 'OFFSET', 'TRUE', 'FALSE']);

function sqlTokens(src: string): Tok[] {
  const out: Tok[] = []; let i = 0;
  while (i < src.length) {
    const ch = src[i];
    if (/\s/.test(ch)) { i++; continue; }
    if (ch === "'" ) { let j = i + 1, s = ''; while (j < src.length) { if (src[j] === "'" && src[j + 1] === "'") { s += "'"; j += 2; continue; } if (src[j] === "'") break; s += src[j++]; } out.push({ t: 'str', v: s }); i = j + 1; continue; }
    if (ch === '"' || ch === '[' || ch === '`') { const close = ch === '[' ? ']' : ch; const j = src.indexOf(close, i + 1); out.push({ t: 'id', v: src.slice(i + 1, j < 0 ? src.length : j) }); i = j < 0 ? src.length : j + 1; continue; }
    const num = /^-?\d+(?:\.\d+)?(?:e[+-]?\d+)?/i.exec(src.slice(i)); if (num && (out.length === 0 || out[out.length - 1].t === 'op' || out[out.length - 1].t === 'punct' || out[out.length - 1].t === 'kw' || !num[0].startsWith('-'))) { out.push({ t: 'num', v: num[0] }); i += num[0].length; continue; }
    const op = /^(<>|!=|<=|>=|=|<|>|\|\|)/.exec(src.slice(i)); if (op) { out.push({ t: 'op', v: op[0] }); i += op[0].length; continue; }
    if ('(),*.'.includes(ch)) { out.push({ t: 'punct', v: ch }); i++; continue; }
    const id = /^[A-Za-z_][\w.$%]*/.exec(src.slice(i)); if (id) { const up = id[0].toUpperCase(); out.push(KEYWORDS.has(up) ? { t: 'kw', v: up } : { t: 'id', v: id[0] }); i += id[0].length; continue; }
    throw new Error(`Unexpected character "${ch}" at position ${i + 1}`);
  }
  out.push({ t: 'end', v: '' });
  return out;
}

type Expr = { k: 'col'; name: string } | { k: 'lit'; v: Scalar } | { k: 'bin'; op: string; a: Expr; b: Expr } | { k: 'not'; a: Expr } | { k: 'in'; a: Expr; list: Expr[]; neg: boolean } | { k: 'between'; a: Expr; lo: Expr; hi: Expr; neg: boolean } | { k: 'isnull'; a: Expr; neg: boolean } | { k: 'like'; a: Expr; pat: Expr; neg: boolean } | { k: 'agg'; fn: string; arg: Expr | null; distinct?: boolean } | { k: 'fn'; name: string; args: Expr[] };
interface Query { distinct: boolean; cols: { e: Expr; alias?: string }[]; star: boolean; from: string; where: Expr | null; group: Expr[]; having: Expr | null; order: { e: Expr; desc: boolean }[]; limit: number | null; offset: number }

class SqlParser {
  i = 0;
  constructor(private toks: Tok[]) {}
  peek(): Tok { return this.toks[this.i]; }
  next(): Tok { return this.toks[this.i++]; }
  kw(k: string): boolean { if (this.peek().t === 'kw' && this.peek().v === k) { this.i++; return true; } return false; }
  expect(k: string): void { if (!this.kw(k)) throw new Error(`Expected ${k} near "${this.peek().v || 'end'}"`); }
  punct(p: string): boolean { if (this.peek().t === 'punct' && this.peek().v === p) { this.i++; return true; } return false; }
  parse(): Query {
    this.expect('SELECT');
    const distinct = this.kw('DISTINCT');
    const cols: Query['cols'] = []; let star = false;
    do { if (this.punct('*')) { star = true; continue; } const e = this.expr(); let alias: string | undefined; if (this.kw('AS')) alias = this.next().v; else if (this.peek().t === 'id') alias = this.next().v; cols.push({ e, alias }); } while (this.punct(','));
    let from = ''; if (this.kw('FROM')) { from = this.next().v; while (this.punct('.')) from += '.' + this.next().v; }
    const where = this.kw('WHERE') ? this.expr() : null;
    const group: Expr[] = []; if (this.kw('GROUP')) { this.expect('BY'); do group.push(this.expr()); while (this.punct(',')); }
    const having = this.kw('HAVING') ? this.expr() : null;
    const order: Query['order'] = []; if (this.kw('ORDER')) { this.expect('BY'); do { const e = this.expr(); const desc = this.kw('DESC') ? true : (this.kw('ASC'), false); order.push({ e, desc }); } while (this.punct(',')); }
    let limit: number | null = null, offset = 0; if (this.kw('LIMIT')) { limit = Number(this.next().v); if (this.punct(',')) { offset = limit; limit = Number(this.next().v); } } if (this.kw('OFFSET')) offset = Number(this.next().v);
    if (this.peek().t !== 'end') throw new Error(`Unexpected "${this.peek().v}"`);
    return { distinct, cols, star, from, where, group, having, order, limit, offset };
  }
  expr(): Expr { return this.or(); }
  or(): Expr { let a = this.and(); while (this.kw('OR')) a = { k: 'bin', op: 'OR', a, b: this.and() }; return a; }
  and(): Expr { let a = this.not(); while (this.kw('AND')) a = { k: 'bin', op: 'AND', a, b: this.not() }; return a; }
  not(): Expr { if (this.kw('NOT')) return { k: 'not', a: this.not() }; return this.cmp(); }
  cmp(): Expr {
    const a = this.add();
    const neg = this.kw('NOT');
    if (this.kw('IN')) { this.punct('('); const list: Expr[] = []; do list.push(this.expr()); while (this.punct(',')); this.punct(')'); return { k: 'in', a, list, neg }; }
    if (this.kw('BETWEEN')) { const lo = this.add(); this.expect('AND'); const hi = this.add(); return { k: 'between', a, lo, hi, neg }; }
    if (this.kw('LIKE')) return { k: 'like', a, pat: this.add(), neg };
    if (this.kw('IS')) { const n = this.kw('NOT'); this.expect('NULL'); return { k: 'isnull', a, neg: n }; }
    if (neg) throw new Error('NOT must be followed by IN, BETWEEN, LIKE or an expression');
    if (this.peek().t === 'op' && ['=', '<>', '!=', '<', '>', '<=', '>='].includes(this.peek().v)) { const op = this.next().v; return { k: 'bin', op: op === '!=' ? '<>' : op, a, b: this.add() }; }
    return a;
  }
  add(): Expr { let a = this.mul(); while (this.peek().t === 'op' && this.peek().v === '||') { this.next(); a = { k: 'bin', op: '||', a, b: this.mul() }; } while ((this.peek().t === 'num' && this.peek().v.startsWith('-')) || (this.peek().t === 'id' && (this.peek().v === '+' || this.peek().v === '-'))) { break; } return a; }
  mul(): Expr { return this.primary(); }
  primary(): Expr {
    const t = this.next();
    if (t.t === 'num') return { k: 'lit', v: Number(t.v) };
    if (t.t === 'str') return { k: 'lit', v: t.v };
    if (t.t === 'kw' && t.v === 'NULL') return { k: 'lit', v: null };
    if (t.t === 'kw' && (t.v === 'TRUE' || t.v === 'FALSE')) return { k: 'lit', v: t.v === 'TRUE' };
    if (t.t === 'punct' && t.v === '(') { const e = this.expr(); this.punct(')'); return e; }
    if (t.t === 'id') {
      const up = t.v.toUpperCase();
      if (this.punct('(')) {
        if (['COUNT', 'SUM', 'AVG', 'MIN', 'MAX'].includes(up)) { if (this.punct('*')) { this.punct(')'); return { k: 'agg', fn: up, arg: null }; } const distinct = this.kw('DISTINCT'); const arg = this.expr(); this.punct(')'); return { k: 'agg', fn: up, arg, distinct }; }
        const args: Expr[] = []; if (!this.punct(')')) { do args.push(this.expr()); while (this.punct(',')); this.punct(')'); } return { k: 'fn', name: up, args };
      }
      return { k: 'col', name: t.v };
    }
    throw new Error(`Unexpected "${t.v || 'end of query'}"`);
  }
}

function likeRe(p: string): RegExp { return new RegExp('^' + p.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/%/g, '.*').replace(/_/g, '.') + '$', 'i'); }
function cmpVals(a: Scalar, b: Scalar): number { if (a == null && b == null) return 0; if (a == null) return -1; if (b == null) return 1; if (typeof a === 'number' && typeof b === 'number') return a - b; const na = Number(a), nb = Number(b); if (!isNaN(na) && !isNaN(nb) && String(a).trim() !== '' && String(b).trim() !== '') return na - nb; return String(a).toLowerCase().localeCompare(String(b).toLowerCase()); }

function evalExpr(e: Expr, row: Record<string, Scalar>, group?: Record<string, Scalar>[]): Scalar {
  switch (e.k) {
    case 'lit': return e.v;
    case 'col': { const k = Object.keys(row).find((x) => x.toLowerCase() === e.name.toLowerCase()); if (k === undefined) throw new Error(`Unknown column "${e.name}"`); return row[k]; }
    case 'not': return !truthy(evalExpr(e.a, row, group));
    case 'bin': {
      if (e.op === 'AND') return truthy(evalExpr(e.a, row, group)) && truthy(evalExpr(e.b, row, group));
      if (e.op === 'OR') return truthy(evalExpr(e.a, row, group)) || truthy(evalExpr(e.b, row, group));
      const a = evalExpr(e.a, row, group), b = evalExpr(e.b, row, group);
      if (e.op === '||') return textOf(a) + textOf(b);
      if (a == null || b == null) return false;
      const c = cmpVals(a, b);
      return e.op === '=' ? c === 0 : e.op === '<>' ? c !== 0 : e.op === '<' ? c < 0 : e.op === '>' ? c > 0 : e.op === '<=' ? c <= 0 : c >= 0;
    }
    case 'in': { const a = evalExpr(e.a, row, group); const hit = e.list.some((x) => cmpVals(a, evalExpr(x, row, group)) === 0); return e.neg ? !hit : hit; }
    case 'between': { const a = evalExpr(e.a, row, group); const lo = evalExpr(e.lo, row, group), hi = evalExpr(e.hi, row, group); const hit = a != null && cmpVals(a, lo) >= 0 && cmpVals(a, hi) <= 0; return e.neg ? !hit : hit; }
    case 'isnull': { const a = evalExpr(e.a, row, group); const isNull = a == null || a === ''; return e.neg ? !isNull : isNull; }
    case 'like': { const a = evalExpr(e.a, row, group); const hit = likeRe(textOf(evalExpr(e.pat, row, group))).test(textOf(a)); return e.neg ? !hit : hit; }
    case 'agg': {
      const rows = group || [row];
      let vals = e.arg ? rows.map((r) => evalExpr(e.arg!, r)).filter((v) => v != null && v !== '') : rows.map(() => 1 as Scalar);
      if (e.distinct) { const seen = new Set<string>(); vals = vals.filter((v) => { const k = textOf(v).toLowerCase(); if (seen.has(k)) return false; seen.add(k); return true; }); }
      const nums = vals.filter((v): v is number => typeof v === 'number');
      switch (e.fn) { case 'COUNT': return vals.length; case 'SUM': return nums.reduce((a, b) => a + b, 0); case 'AVG': return nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : null; case 'MIN': return vals.length ? vals.reduce((a, b) => (cmpVals(a, b) <= 0 ? a : b)) : null; default: return vals.length ? vals.reduce((a, b) => (cmpVals(a, b) >= 0 ? a : b)) : null; }
    }
    case 'fn': {
      const args = e.args.map((a) => evalExpr(a, row, group));
      switch (e.name) {
        case 'UPPER': return textOf(args[0]).toUpperCase(); case 'LOWER': return textOf(args[0]).toLowerCase(); case 'TRIM': return textOf(args[0]).trim(); case 'LENGTH': case 'LEN': return textOf(args[0]).length;
        case 'ROUND': return typeof args[0] === 'number' ? Math.round(args[0] * Math.pow(10, Number(args[1]) || 0)) / Math.pow(10, Number(args[1]) || 0) : args[0]; case 'ABS': return typeof args[0] === 'number' ? Math.abs(args[0]) : args[0];
        case 'COALESCE': case 'IFNULL': return args.find((a) => a != null && a !== '') ?? null; case 'SUBSTR': case 'SUBSTRING': return textOf(args[0]).substr(Math.max(0, Number(args[1]) - 1), args[2] != null ? Number(args[2]) : undefined);
        case 'LEFT': return textOf(args[0]).slice(0, Number(args[1])); case 'RIGHT': return textOf(args[0]).slice(-Number(args[1])); case 'CONCAT': return args.map(textOf).join(''); case 'YEAR': case 'MONTH': case 'DAY': { if (typeof args[0] !== 'number') return null; const { serialToParts } = requireDates(); const p = serialToParts(args[0]); return e.name === 'YEAR' ? p.y : e.name === 'MONTH' ? p.m : p.d; }
        default: throw new Error(`Unknown function ${e.name}`);
      }
    }
  }
}
const truthy = (v: Scalar): boolean => !!v && v !== 0 && v !== '';
function hasAgg(e: Expr): boolean { switch (e.k) { case 'agg': return true; case 'bin': return hasAgg(e.a) || hasAgg(e.b); case 'not': return hasAgg(e.a); case 'in': return hasAgg(e.a) || e.list.some(hasAgg); case 'between': return hasAgg(e.a) || hasAgg(e.lo) || hasAgg(e.hi); case 'isnull': return hasAgg(e.a); case 'like': return hasAgg(e.a); case 'fn': return e.args.some(hasAgg); default: return false; } }
function exprLabel(e: Expr): string { switch (e.k) { case 'col': return e.name; case 'lit': return textOf(e.v); case 'agg': return `${e.fn}(${e.arg ? exprLabel(e.arg) : '*'})`; case 'fn': return `${e.name}(${e.args.map(exprLabel).join(', ')})`; case 'bin': return `${exprLabel(e.a)} ${e.op} ${exprLabel(e.b)}`; default: return 'expr'; } }

/** Runs a SQL SELECT over rows of records. */
export function runSql(sql: string, tables: Record<string, Record<string, Scalar>[]>, defaultTable?: string): { columns: string[]; rows: Scalar[][] } {
  const q = new SqlParser(sqlTokens(sql)).parse();
  const tname = q.from || defaultTable || Object.keys(tables)[0];
  const key = Object.keys(tables).find((k) => k.toLowerCase() === tname.toLowerCase());
  if (!key) throw new Error(`Unknown sheet "${tname}"`);
  let rows = tables[key];
  if (q.where) rows = rows.filter((r) => truthy(evalExpr(q.where!, r)));
  const aggregated = q.group.length || q.cols.some((c) => hasAgg(c.e));
  let groups: { rep: Record<string, Scalar>; rows: Record<string, Scalar>[] }[];
  if (aggregated) { const m = new Map<string, { rep: Record<string, Scalar>; rows: Record<string, Scalar>[] }>(); for (const r of rows) { const k = q.group.map((g) => textOf(evalExpr(g, r)).toLowerCase()).join('\u0001'); (m.get(k) || m.set(k, { rep: r, rows: [] }).get(k)!).rows.push(r); } groups = [...m.values()]; if (!rows.length && !q.group.length) groups = [{ rep: {}, rows: [] }]; }
  else groups = rows.map((r) => ({ rep: r, rows: [r] }));
  if (q.having) groups = groups.filter((g) => truthy(evalExpr(q.having!, g.rep, g.rows)));
  const columns = q.star ? Object.keys(rows[0] || tables[key][0] || {}) : [];
  const cols = q.star ? [...columns.map((c) => ({ e: { k: 'col', name: c } as Expr, alias: c })), ...q.cols] : q.cols;
  const names = cols.map((c) => c.alias || exprLabel(c.e));
  let out = groups.map((g) => ({ g, vals: cols.map((c) => evalExpr(c.e, g.rep, aggregated ? g.rows : undefined)) }));
  if (q.order.length) out.sort((x, y) => { for (const o of q.order) { const ia = names.findIndex((n) => n.toLowerCase() === exprLabel(o.e).toLowerCase()); const a = ia >= 0 ? x.vals[ia] : evalExpr(o.e, x.g.rep, aggregated ? x.g.rows : undefined), b = ia >= 0 ? y.vals[ia] : evalExpr(o.e, y.g.rep, aggregated ? y.g.rows : undefined); const c = cmpVals(a, b); if (c) return o.desc ? -c : c; } return 0; });
  if (q.distinct) { const seen = new Set<string>(); out = out.filter((x) => { const k = x.vals.map((v) => textOf(v).toLowerCase()).join('\u0001'); if (seen.has(k)) return false; seen.add(k); return true; }); }
  if (q.offset) out = out.slice(q.offset);
  if (q.limit != null) out = out.slice(0, q.limit);
  return { columns: names, rows: out.map((x) => x.vals) };
}

function sheetRecords(s: Sheet): Record<string, Scalar>[] {
  const u = s.usedRange(false); if (!u) return [];
  const rows = valueRows(s, u);
  const header = (rows[0] || []).map((v, j) => textOf(v) || colName(u.c1 + j));
  return rows.slice(1).map((r) => Object.fromEntries(header.map((hh, j) => [hh, r[j] ?? null])));
}

let lastSql = '';

defTools([
  {
    id: 'plus.sqlQuery', label: 'SQL Query on the Sheets…', icon: 'database', tab: P, group: G_QUERY, edits: true, popular: true, tip: 'SELECT … FROM Sheet1 WHERE … GROUP BY … ORDER BY … LIMIT …, with COUNT, SUM, AVG, MIN, MAX, LIKE, IN, BETWEEN and DISTINCT, over any sheet (first row = column names); the result goes on a new sheet', keywords: 'sql select query where group by order by', origin: 'XLTools (SQL Queries)',
    run: async (app, arg?: { sql: string }) => {
      const tables: Record<string, Record<string, Scalar>[]> = {}; for (const s of app.book.sheets) tables[s.name] = sheetRecords(s);
      const cols = Object.keys(tables[app.sheet.name]?.[0] || {});
      const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'SQL Query', wide: true, help: `Sheets are tables; the first row gives the column names. Quote names with spaces in double quotes or [brackets]. Columns of ${app.sheet.name}: ${cols.join(', ') || '(none)'}.`, fields: [{ key: 'sql', label: 'Query', type: 'textarea', rows: 6, value: lastSql || `SELECT * FROM [${app.sheet.name}]${cols.length ? ` WHERE [${cols[0]}] IS NOT NULL` : ''} LIMIT 100`, autofocus: true }] }); return r ? { sql: String(r.sql || '') } : null; });
      if (!o || !o.sql.trim()) return;
      lastSql = o.sql;
      let res: { columns: string[]; rows: Scalar[][] };
      try { res = runSql(o.sql, tables, app.sheet.name); } catch (e) { await alertDialog('SQL Query', (e as Error).message); return; }
      report(app, 'Query', res.columns, res.rows.map((r) => r.map((v) => v)), { colW: res.columns.map(() => 120) });
      toast(`${plural(res.rows.length, 'row')} returned`);
    },
  },
  { id: 'plus.sqlDistinct', label: 'Distinct Values of the Column (SQL)', icon: 'database', tab: P, group: G_QUERY, edits: true, run: (app) => { const s = app.sheet; const u = s.usedRange(false); if (!u) return; const name = textOf(s.value(u.r1, activeCell(app).c)) || colName(activeCell(app).c); app.exec('plus.sqlQuery', { sql: `SELECT [${name}], COUNT(*) AS Count FROM [${s.name}] GROUP BY [${name}] ORDER BY Count DESC` }); } },
  { id: 'plus.menuQuery', label: 'Query', icon: 'database', tab: P, group: G_QUERY, kind: 'menu', noRecord: true, run: () => undefined, menu: () => menuOf(['plus.sqlQuery', 'plus.sqlDistinct', 'data.getData', 'data.advancedFilter', 'tools.groupSummary', 'insert.pivotTable']) },
]);

function requireTools(): { toRoman: (n: number) => string } { return toolsModule; }
import * as toolsModule from './tools';
import * as datesModule from '../model/dates';
function requireDates(): typeof datesModule { return datesModule; }

void Workbook; void MAX_ROWS; void selRanges; void structureLocked; void textRows; void writeBlock; void escapeHtml; void confirmDialog; void numOf; void parseRange; void shownText;
