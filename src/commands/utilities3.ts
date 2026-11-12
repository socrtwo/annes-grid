/**
 * Utilities tab, part 3: Format & Layout, Numbers & Dates, Text, Formulas (ASAP Utilities-style tools,
 * re-implemented; see utilities.ts).
 */
import type { App } from '../app';
import { Sheet } from '../model/workbook';
import type { Cell, CellStyle, Scalar } from '../model/types';
import { CellError } from '../model/types';
import { cellKey, keyRow, keyCol, colName, cellName, rangeName, type Range } from '../model/address';
import { mapRefs, shiftFormula, tokenize } from '../calc/parser';
import { serialToParts, partsToSerial, daysInMonth, DAYS, MONTHS } from '../model/dates';
import { isDateFormat } from '../model/numfmt';
import { selRanges, activeRange, activeCell, clipToUsed } from './ops';
import { defTools, externalLinkIn, U, selCells, selBlock, singleCell, hasContent, selectHits, sheetProtected, emptySheet, numOf, textOf, plural, report, shownText, argOrAsk, columnChoices, headerGuess, writeOut, checkWritable, canChange, changeCells, patchStyle, movedCell, blockOf, copyText, toCsvField } from './utilShared';
import { menuOf, cellMatches, findMatching, type SelectCriteria } from './utilities';
import { transformValues, textToNumber, numberToWords, toRoman, fromRoman } from './tools';
import { openDialog, toast, promptDialog, choiceDialog } from '../ui/dialog';
import { clearRulesIn } from './conditional';
import { setPrint } from './pagelayout';

const G_FORMAT = 'Format & Layout';
const G_NUMBERS = 'Numbers & Dates';
const G_TEXT = 'Text';
const G_FORMULAS = 'Formulas';

// ============================================================================================
// Format & Layout

const STYLE_PRESETS: { value: string; label: string; style: Partial<CellStyle> }[] = [
  { value: 'bold', label: 'Bold', style: { bold: true } }, { value: 'italic', label: 'Italic', style: { italic: true } },
  { value: 'red', label: 'Red font', style: { color: '#c62828' } }, { value: 'yellow', label: 'Yellow fill', style: { fill: '#fff59d' } }, { value: 'green', label: 'Green fill', style: { fill: '#c8e6c9' } }, { value: 'blue', label: 'Light blue fill', style: { fill: '#bbdefb' } },
  { value: 'strike', label: 'Strikethrough', style: { strike: true } }, { value: 'border', label: 'Outline border', style: { bt: { style: 'thin' }, bb: { style: 'thin' }, bl: { style: 'thin' }, br: { style: 'thin' } } },
  { value: 'wrap', label: 'Wrap text', style: { wrap: true } }, { value: 'number2', label: 'Number, 2 decimals', style: { numFmt: '#,##0.00' } }, { value: 'percent', label: 'Percent', style: { numFmt: '0.0%' } }, { value: 'date', label: 'Date', style: { numFmt: 'yyyy-mm-dd' } },
];

const REMOVE_PARTS: { value: string; label: string; keys: (keyof CellStyle)[] }[] = [
  { value: 'fills', label: 'Fill colors and patterns', keys: ['fill', 'pattern', 'gradient'] }, { value: 'borders', label: 'Borders', keys: ['bt', 'bb', 'bl', 'br', 'diagUp', 'diagDown'] }, { value: 'fontColors', label: 'Font colors', keys: ['color'] },
  { value: 'fontStyles', label: 'Bold, italic, underline, strikethrough', keys: ['bold', 'italic', 'underline', 'strike', 'vertAlign'] }, { value: 'fonts', label: 'Font names and sizes', keys: ['font', 'size'] }, { value: 'alignment', label: 'Alignment, wrap, indent, rotation', keys: ['hAlign', 'vAlign', 'wrap', 'shrink', 'indent', 'rotate'] },
  { value: 'numFmts', label: 'Number formats', keys: ['numFmt'] }, { value: 'protection', label: 'Locked/hidden flags', keys: ['locked', 'hidden'] },
];

function removeStyleParts(app: App, parts: string[], scope: 'selection' | 'sheet' | 'workbook'): number {
  const keys = new Set<keyof CellStyle>(REMOVE_PARTS.filter((p) => parts.includes(p.value)).flatMap((p) => p.keys));
  const patch: Partial<CellStyle> = {}; for (const k of keys) (patch as any)[k] = undefined;
  let n = 0;
  const strip = (s: Sheet, r: number, c: number, cell: Cell | undefined) => { const sid = s.styleIdOf(r, c); if (!sid) return; const st = s.book.styles[sid] || {}; if (!Object.keys(st).some((k) => keys.has(k as keyof CellStyle))) return; const id = s.book.derivedStyle(sid, patch); s.set(r, c, { ...(cell || {}), s: id || undefined }); n++; };
  app.book.transact('Remove Formatting', () => {
    if (scope === 'selection') { for (const x of selCells(app)) strip(app.sheet, x.r, x.c, x.cell); }
    else for (const s of scope === 'sheet' ? [app.sheet] : app.book.sheets) { for (const [k, cell] of [...s.cells]) strip(s, keyRow(k), keyCol(k), cell); if (keys.has('fill') || keys.has('font')) { app.book.recordSheet(s); for (const [r, sid] of [...s.rowStyle]) s.rowStyle.set(r, s.book.derivedStyle(sid, patch)); for (const [c, sid] of [...s.colStyle]) s.colStyle.set(c, s.book.derivedStyle(sid, patch)); } }
  });
  return n;
}

const BAND_COLORS = ['#e3f2fd', '#fff8e1', '#e8f5e9', '#fce4ec', '#ede7f6', '#f1f8e9'];

/** Static formatting the conditional formats currently produce, per cell (for Freeze Conditional Formatting). */
function frozenCfStyles(app: App, s: Sheet, rg: Range): { r: number; c: number; style: Partial<CellStyle> }[] {
  const out: { r: number; c: number; style: Partial<CellStyle> }[] = [];
  if (!s.cf.length) return out;
  for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) {
    const res = app.tab.cf.evaluate(s, r, c, app.book.version);
    if (!res) continue;
    const style: Partial<CellStyle> = { ...(res.style || {}) };
    if (res.fill) style.fill = res.fill;
    if (res.bar) style.fill = res.bar.color;
    if (Object.keys(style).length) out.push({ r, c, style });
  }
  return out;
}

defTools([
  {
    id: 'util.formatByContent', label: 'Format Cells by Content…', icon: 'format_paint', tab: U, group: G_FORMAT, edits: true, popular: true, tip: 'Apply a format only to the cells that meet a condition (contains text, greater than, blank, errors, dates…), as fixed formatting rather than a rule', keywords: 'conditional static format cells containing', origin: 'ASAP Utilities (Apply formatting only to cells with…)',
    run: async (app, arg?: SelectCriteria & { preset?: string; style?: Partial<CellStyle>; scope?: 'selection' | 'sheet' }) => {
      const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Format Cells by Content', fields: [
        { key: 'content', label: 'Cells whose content', type: 'select', value: 'contains', options: [{ value: 'equals', label: 'Is equal to' }, { value: 'contains', label: 'Contains' }, { value: 'starts', label: 'Starts with' }, { value: 'ends', label: 'Ends with' }, { value: 'greater', label: 'Is greater than' }, { value: 'less', label: 'Is less than' }, { value: 'between', label: 'Is between' }, { value: 'blank', label: 'Is blank' }, { value: 'nonblank', label: 'Is not blank' }, { value: 'number', label: 'Is a number' }, { value: 'text', label: 'Is text' }, { value: 'date', label: 'Is a date' }, { value: 'error', label: 'Is an error' }, { value: 'formula', label: 'Is a formula' }, { value: 'negative', label: 'Is negative' }, { value: 'duplicate', label: 'Is a duplicate' }, { value: 'regex', label: 'Matches a regular expression' }] },
        { key: 'value', label: 'Value', type: 'text', value: '', half: true }, { key: 'value2', label: 'And', type: 'text', value: '', half: true, visibleIf: (v) => v.content === 'between' }, { key: 'matchCase', label: 'Match case', type: 'checkbox', value: false },
        { key: 'preset', label: 'Apply', type: 'select', value: 'yellow', options: STYLE_PRESETS.map((p) => ({ value: p.value, label: p.label })) },
        { key: 'scope', label: 'Look in', type: 'radio', value: singleCell(app) ? 'sheet' : 'selection', options: [{ value: 'selection', label: 'The selection' }, { value: 'sheet', label: 'The whole sheet' }] },
      ] }); return r as any; });
      if (!o) return;
      const style = o.style || STYLE_PRESETS.find((p) => p.value === o.preset)?.style || { fill: '#fff59d' };
      const hits = findMatching(app, o, o.scope || 'selection');
      if (!hits.length) { toast('No cells match.'); return; }
      const s = app.sheet;
      app.book.transact('Format by Content', () => { for (const p of hits) s.set(p.r, p.c, { ...(s.get(p.r, p.c) || {}), s: app.book.derivedStyle(s.styleIdOf(p.r, p.c), style) }); });
      toast(`${plural(hits.length, 'cell')} formatted`);
    },
  },
  {
    id: 'util.bandRows', label: 'Color Every Nth Row or Column…', icon: 'view_agenda', tab: U, group: G_FORMAT, edits: true, popular: true, tip: 'Shade every 2nd (3rd…) row or column of the selection with a fixed fill (no table or rule needed)', keywords: 'banded rows zebra stripes alternate shading', origin: 'ASAP Utilities / Kutools (Alternate Row/Column Shading)',
    run: async (app, arg?: { axis: 'row' | 'col'; every: number; color: string; start?: number }) => {
      const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Color Every Nth Row or Column', fields: [{ key: 'axis', label: 'Shade', type: 'radio', value: 'row', options: [{ value: 'row', label: 'Rows' }, { value: 'col', label: 'Columns' }] }, { key: 'every', label: 'Every', type: 'number', value: 2, min: 1, max: 50, half: true }, { key: 'start', label: 'Starting at the', type: 'number', value: 2, min: 1, half: true, suffix: 'th' }, { key: 'color', label: 'Fill color', type: 'color', value: '#e3f2fd' }] }); return r ? { axis: r.axis, every: Number(r.every) || 2, start: Number(r.start) || 1, color: r.color } : null; });
      if (!o) return;
      const s = app.sheet; const rg = selBlock(app); if (!checkWritable(app, [rg])) return;
      const first = o.axis === 'row' ? rg.r1 : rg.c1, last = o.axis === 'row' ? rg.r2 : rg.c2;
      app.book.transact('Band Rows', () => { for (let l = first + (o.start || 1) - 1; l <= last; l += o.every) { if (o.axis === 'row') for (let c = rg.c1; c <= rg.c2; c++) s.set(l, c, { ...(s.get(l, c) || {}), s: app.book.derivedStyle(s.styleIdOf(l, c), { fill: o.color }) }); else for (let r = rg.r1; r <= rg.r2; r++) s.set(r, l, { ...(s.get(r, l) || {}), s: app.book.derivedStyle(s.styleIdOf(r, l), { fill: o.color }) }); } });
    },
  },
  {
    id: 'util.bandByGroup', label: 'Color Groups of Rows by Value…', icon: 'layers', tab: U, group: G_FORMAT, edits: true, tip: 'Alternate two fills each time the value in a column changes, so each group of rows stands out', keywords: 'highlight groups alternate change of value', origin: 'ASAP Utilities / Kutools (Distinguish Differences)',
    run: async (app, arg?: { col: number; colors?: string[] }) => {
      if (emptySheet(app)) return;
      const s = app.sheet; const rg = selBlock(app); const header = headerGuess(s, rg);
      const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Color Groups of Rows', fields: [{ key: 'col', label: 'Group by the value in column', type: 'select', value: String(rg.c1), options: columnChoices(s, rg, header) }, { key: 'rainbow', label: 'Use a different color for each group (otherwise two alternating)', type: 'checkbox', value: false }] }); return r ? { col: Number(r.col), colors: r.rainbow ? BAND_COLORS : [BAND_COLORS[0], '#ffffff'] } : null; });
      if (!o) return;
      if (!checkWritable(app, [rg])) return;
      const colors = o.colors || [BAND_COLORS[0], '#ffffff'];
      app.book.transact('Color Groups', () => { let g = 0; for (let r = rg.r1 + (header ? 1 : 0); r <= rg.r2; r++) { if (r > rg.r1 + (header ? 1 : 0) && shownText(s, r, o.col) !== shownText(s, r - 1, o.col)) g++; const fill = colors[g % colors.length]; for (let c = rg.c1; c <= rg.c2; c++) s.set(r, c, { ...(s.get(r, c) || {}), s: app.book.derivedStyle(s.styleIdOf(r, c), { fill: fill === '#ffffff' ? undefined : fill }) }); } });
    },
  },
  {
    id: 'util.removeFormattingParts', label: 'Remove Formatting Parts…', icon: 'format_clear', tab: U, group: G_FORMAT, edits: true, popular: true, tip: 'Remove only the fills, borders, font colors, font styles, alignment or number formats, from the selection, the sheet or the whole workbook', keywords: 'clear fills borders colors keep values', origin: 'ASAP Utilities',
    run: async (app, arg?: { parts: string[]; scope: 'selection' | 'sheet' | 'workbook' }) => {
      const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Remove Formatting Parts', fields: [...REMOVE_PARTS.map((p) => ({ key: p.value, label: p.label, type: 'checkbox' as const, value: p.value === 'fills' })), { key: 'scope', label: 'From', type: 'radio', value: 'selection', options: [{ value: 'selection', label: 'The selection' }, { value: 'sheet', label: 'This sheet' }, { value: 'workbook', label: 'The whole workbook' }] }] }); return r ? { parts: REMOVE_PARTS.map((p) => p.value).filter((v) => r[v]), scope: r.scope } : null; });
      if (!o || !o.parts.length) return;
      const n = removeStyleParts(app, o.parts, o.scope);
      app.grid.invalidate(); toast(`${plural(n, 'cell')} changed`);
    },
  },
  { id: 'util.removeFills', label: 'Remove All Fill Colors', icon: 'format_color_reset', tab: U, group: G_FORMAT, edits: true, run: (app) => { removeStyleParts(app, ['fills'], singleCell(app) ? 'sheet' : 'selection'); app.grid.invalidate(); } },
  { id: 'util.removeBorders', label: 'Remove All Borders', icon: 'border_clear', tab: U, group: G_FORMAT, edits: true, run: (app) => { removeStyleParts(app, ['borders'], singleCell(app) ? 'sheet' : 'selection'); app.grid.invalidate(); } },
  { id: 'util.removeFontColors', label: 'Remove All Font Colors', icon: 'format_color_text', tab: U, group: G_FORMAT, edits: true, run: (app) => { removeStyleParts(app, ['fontColors'], singleCell(app) ? 'sheet' : 'selection'); app.grid.invalidate(); } },
  { id: 'util.removeCFBook', label: 'Remove Conditional Formatting from the Workbook', icon: 'rule_folder', tab: U, group: G_FORMAT, edits: true, origin: 'ASAP Utilities', run: (app) => { let n = 0; app.book.transact('Remove Conditional Formatting', () => { for (const s of app.book.sheets) { app.book.recordSheet(s); n += s.cf.length; s.cf = []; } }); app.tab.cf.invalidate(); app.grid.invalidate(); toast(`${plural(n, 'rule')} removed`); } },
  {
    id: 'util.freezeCF', label: 'Freeze Conditional Formatting', icon: 'ac_unit', tab: U, group: G_FORMAT, edits: true, popular: true, tip: 'Turn the formatting that conditional formatting rules currently show into ordinary cell formatting and remove the rules (for sharing or sorting by color)', keywords: 'static conditional formatting convert rules', origin: 'ASAP Utilities (Freeze conditional formatting)',
    run: (app, arg?: { keepRules?: boolean }) => {
      const s = app.sheet; if (!s.cf.length) { toast('This sheet has no conditional formatting.'); return; }
      const u = s.usedRange(); if (!u) return;
      const rg = singleCell(app) ? u : clipToUsed(s, activeRange(app));
      const frozen = frozenCfStyles(app, s, rg);
      app.book.transact('Freeze Conditional Formatting', () => { for (const f of frozen) s.set(f.r, f.c, { ...(s.get(f.r, f.c) || {}), s: app.book.derivedStyle(s.styleIdOf(f.r, f.c), f.style) }); if (!arg?.keepRules) clearRulesIn(s, [rg]); });
      app.tab.cf.invalidate(); app.grid.invalidate();
      toast(`${plural(frozen.length, 'cell')} formatted permanently`);
    },
  },
  {
    id: 'util.fontAllSheets', label: 'Set the Font on All Sheets…', icon: 'font_download', tab: U, group: G_FORMAT, edits: true, tip: 'Change the font name and/or size of every cell in the workbook', origin: 'ASAP Utilities (Change font on all sheets)',
    run: async (app, arg?: { font?: string; size?: number }) => {
      const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Font on All Sheets', fields: [{ key: 'font', label: 'Font (empty = keep)', type: 'text', value: '', half: true }, { key: 'size', label: 'Size (0 = keep)', type: 'number', value: 0, min: 0, max: 200, half: true }] }); return r ? { font: r.font || undefined, size: Number(r.size) || undefined } : null; });
      if (!o || (!o.font && !o.size)) return;
      const patch: Partial<CellStyle> = {}; if (o.font) patch.font = o.font; if (o.size) patch.size = o.size;
      let n = 0;
      app.book.transact('Font on All Sheets', () => { for (const s of app.book.sheets) { for (const [k, cell] of [...s.cells]) { s.set(keyRow(k), keyCol(k), { ...cell, s: app.book.derivedStyle(cell.s ?? 0, patch) }); n++; } } if (o.font) app.prefs.defaultFont = o.font; if (o.size) app.prefs.defaultSize = o.size; app.savePrefs(); });
      app.grid.layout(); toast(`${plural(n, 'cell')} updated`);
    },
  },
  {
    id: 'util.borderBlocks', label: 'Outline Each Block of Data', icon: 'border_outer', tab: U, group: G_FORMAT, edits: true, tip: 'Draw a border around every separate block of cells on the sheet (or in the selection)', origin: 'ASAP Utilities',
    run: async (app) => {
      const s = app.sheet; const u = s.usedRange(false); if (!u) { toast('The sheet is empty.'); return; }
      const { currentRegion } = await import('./ops');
      const area = singleCell(app) ? u : clipToUsed(s, activeRange(app));
      const seen = new Set<number>(); const blocks: Range[] = [];
      s.forEachIn(area, (r, c, cell) => { if (!hasContent(cell) || seen.has(cellKey(r, c))) return; const rg = currentRegion(s, r, c); blocks.push(rg); for (let rr = rg.r1; rr <= rg.r2; rr++) for (let cc = rg.c1; cc <= rg.c2; cc++) seen.add(cellKey(rr, cc)); });
      const b = { style: 'medium' as const, color: '#333333' };
      app.book.transact('Outline Blocks', () => { for (const rg of blocks) for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) { const p: Partial<CellStyle> = {}; if (r === rg.r1) p.bt = b; if (r === rg.r2) p.bb = b; if (c === rg.c1) p.bl = b; if (c === rg.c2) p.br = b; if (Object.keys(p).length) s.set(r, c, { ...(s.get(r, c) || {}), s: app.book.derivedStyle(s.styleIdOf(r, c), p) }); } });
      toast(`${plural(blocks.length, 'block')} outlined`);
    },
  },
  { id: 'util.negativeRed', label: 'Show Negative Numbers in Red', icon: 'remove_circle', tab: U, group: G_FORMAT, edits: true, tip: 'Add [Red] to the number format of the selected cells so negatives are red', origin: 'ASAP Utilities', run: (app) => { patchStyle(app, 'Negative Red', (st) => { const f = st.numFmt || '#,##0.00'; if (/\[Red\]/i.test(f)) return null; const parts = f.split(';'); return { numFmt: parts.length >= 2 ? `${parts[0]};[Red]${parts[1].replace(/^\[Red\]/i, '')}${parts.slice(2).map((p) => ';' + p).join('')}` : `${f};[Red]-${f}` }; }); } },
  { id: 'util.unwrapSheet', label: 'Unwrap Text on the Whole Sheet', icon: 'wrap_text', tab: U, group: G_FORMAT, edits: true, run: (app) => { const s = app.sheet; let n = 0; app.book.transact('Unwrap', () => { for (const [k, cell] of [...s.cells]) { if (!(s.book.styles[cell.s ?? 0] || {}).wrap) continue; s.set(keyRow(k), keyCol(k), { ...cell, s: app.book.derivedStyle(cell.s ?? 0, { wrap: undefined }) || undefined }); n++; } }); toast(`${plural(n, 'cell')} unwrapped`); } },
  { id: 'util.fitOnePage', label: 'Fit the Sheet on One Page', icon: 'fit_page', tab: U, group: G_FORMAT, edits: true, kind: 'split', origin: 'ASAP Utilities', run: (app, arg?: 'both' | 'wide' | 'tall') => { setPrint(app, 'Fit to Page', (p) => { p.fitWide = arg === 'tall' ? undefined : 1; p.fitTall = arg === 'wide' ? undefined : 1; }); toast('Print scaling set'); }, menu: () => [{ id: 'util.fitOnePage', arg: 'both', label: 'One page' }, { id: 'util.fitOnePage', arg: 'wide', label: 'One page wide' }, { id: 'util.fitOnePage', arg: 'tall', label: 'One page tall' }] },
  { id: 'util.printTitlesAll', label: 'Repeat the First Row When Printing, on All Sheets', icon: 'print', tab: U, group: G_FORMAT, edits: true, origin: 'ASAP Utilities', run: (app) => { for (const s of app.book.sheets) setPrint(app, 'Print Titles', (p) => { p.titleRows = '$1:$1'; }, s); toast('Row 1 repeats on every printed page of every sheet'); } },
  { id: 'util.printAreaAllSheets', label: 'Print Area = Used Range on All Sheets', icon: 'crop', tab: U, group: G_FORMAT, edits: true, run: (app) => { for (const s of app.book.sheets) { const u = s.usedRange(false); setPrint(app, 'Print Area', (p) => { p.area = u ? rangeName(u) : undefined; }, s); } toast('Print areas set'); } },
  {
    id: 'util.deleteUnusedStyles', label: 'Delete Unused Cell Styles', icon: 'style', tab: U, group: G_FORMAT, edits: true, tip: 'Remove custom cell styles that no cell uses (files from other programs can carry thousands)', origin: 'ASAP Utilities / Inquire (Clean Excess Cell Formatting)',
    run: (app) => { const book = app.book; const used = new Set<string>(); const before = book.cellStyles.length; book.transact('Delete Unused Styles', () => { book.recordBook(); book.cellStyles = book.cellStyles.filter((cs) => cs.builtin || used.has(cs.name) || cs.name === 'Normal'); }); toast(`${before - book.cellStyles.length} unused styles removed`); },
  },
  { id: 'util.menuFormat', label: 'Format & Layout', icon: 'format_paint', tab: U, group: G_FORMAT, kind: 'menu', noRecord: true, run: () => undefined, menu: () => menuOf(['util.formatByContent', 'util.bandRows', 'util.bandByGroup', 'util.colorDuplicates', 'util.borderBlocks', 'util.negativeRed', 'util.fontAllSheets', null, 'util.removeFormattingParts', 'util.removeFills', 'util.removeBorders', 'util.removeFontColors', 'util.unwrapSheet', 'fmt.clearFormats', 'util.deleteUnusedStyles', null, 'util.freezeCF', 'util.removeCFBook', 'cond.manageRules', null, 'util.fitOnePage', 'util.printTitlesAll', 'util.printAreaAllSheets', 'util.pageSetupToSheets', 'util.headerFooterAll', 'layout.pageSetup']) },
]);

// ============================================================================================
// Numbers & Dates

export type NumOp = 'add' | 'subtract' | 'multiply' | 'divide' | 'power' | 'negate' | 'abs' | 'percentOf' | 'round' | 'roundTo' | 'ceiling' | 'floor' | 'int' | 'sqrt' | 'log' | 'reciprocal' | 'addPercent';

export function applyNumOp(v: number, op: NumOp, x = 0): number {
  switch (op) {
    case 'add': return v + x; case 'subtract': return v - x; case 'multiply': return v * x; case 'divide': return x === 0 ? NaN : v / x; case 'power': return Math.pow(v, x);
    case 'negate': return -v; case 'abs': return Math.abs(v); case 'percentOf': return v / 100; case 'addPercent': return v + v * x / 100;
    case 'round': { const m = Math.pow(10, x), a = Math.abs(v) * m; /* 1.005 is stored a hair below: the nudge rounds it the way it reads */ return Math.sign(v) * Math.round(a + a * 4 * Number.EPSILON) / m; }
    case 'roundTo': return x === 0 ? v : Math.round(v / x) * x; case 'ceiling': return x === 0 ? v : Math.ceil(v / x) * x; case 'floor': return x === 0 ? v : Math.floor(v / x) * x;
    case 'int': return Math.trunc(v); case 'sqrt': return v < 0 ? NaN : Math.sqrt(v); case 'log': return v <= 0 ? NaN : Math.log10(v); case 'reciprocal': return v === 0 ? NaN : 1 / v;
  }
}

const OP_LABELS: { value: NumOp; label: string; needs?: boolean; formula: (f: string, x: number) => string }[] = [
  { value: 'add', label: 'Add', needs: true, formula: (f, x) => `(${f})+${x}` }, { value: 'subtract', label: 'Subtract', needs: true, formula: (f, x) => `(${f})-${x}` }, { value: 'multiply', label: 'Multiply by', needs: true, formula: (f, x) => `(${f})*${x}` }, { value: 'divide', label: 'Divide by', needs: true, formula: (f, x) => `(${f})/${x}` },
  { value: 'power', label: 'Raise to the power', needs: true, formula: (f, x) => `(${f})^${x}` }, { value: 'addPercent', label: 'Increase by percent', needs: true, formula: (f, x) => `(${f})*(1+${x}/100)` }, { value: 'percentOf', label: 'Divide by 100 (to a fraction)', formula: (f) => `(${f})/100` },
  { value: 'negate', label: 'Change the sign', formula: (f) => `-(${f})` }, { value: 'abs', label: 'Absolute value', formula: (f) => `ABS(${f})` }, { value: 'round', label: 'Round to decimals', needs: true, formula: (f, x) => `ROUND(${f},${x})` }, { value: 'roundTo', label: 'Round to a multiple of', needs: true, formula: (f, x) => `MROUND(${f},${x})` },
  { value: 'ceiling', label: 'Round up to a multiple of', needs: true, formula: (f, x) => `CEILING(${f},${x})` }, { value: 'floor', label: 'Round down to a multiple of', needs: true, formula: (f, x) => `FLOOR(${f},${x})` }, { value: 'int', label: 'Drop the decimals', formula: (f) => `TRUNC(${f})` },
  { value: 'sqrt', label: 'Square root', formula: (f) => `SQRT(${f})` }, { value: 'log', label: 'Log base 10', formula: (f) => `LOG10(${f})` }, { value: 'reciprocal', label: 'Reciprocal (1/x)', formula: (f) => `1/(${f})` },
];

function numberOperation(app: App, op: NumOp, x: number, formulas: boolean, label: string): void {
  const def = OP_LABELS.find((o) => o.value === op)!;
  let n = 0;
  changeCells(app, label, (cell) => {
    if (cell.f != null) { if (!formulas) return undefined; n++; return { ...cell, f: def.formula(cell.f, x) }; }
    const v = typeof cell.v === 'number' ? cell.v : typeof cell.v === 'string' ? textToNumber(cell.v) : null;
    if (v == null) return undefined;
    const out = applyNumOp(v, op, x);
    if (isNaN(out) || !isFinite(out)) return { ...cell, v: new CellError('#NUM!') };
    n++;
    return { ...cell, v: out };
  }, { formulas });
}

const UNITS: Record<string, Record<string, number>> = {
  length: { m: 1, km: 1000, cm: 0.01, mm: 0.001, in: 0.0254, ft: 0.3048, yd: 0.9144, mi: 1609.344, nmi: 1852, pt: 0.0254 / 72, pica: 0.0254 / 6 },
  mass: { kg: 1, g: 0.001, mg: 1e-6, t: 1000, lb: 0.45359237, oz: 0.028349523125, st: 6.35029318 },
  area: { 'm²': 1, 'km²': 1e6, 'cm²': 1e-4, ha: 1e4, acre: 4046.8564224, 'ft²': 0.09290304, 'in²': 0.00064516, 'mi²': 2589988.110336, 'yd²': 0.83612736 },
  volume: { l: 1, ml: 0.001, 'm³': 1000, 'gal (US)': 3.785411784, 'gal (UK)': 4.54609, qt: 0.946352946, pt: 0.473176473, cup: 0.2365882365, 'fl oz': 0.0295735295625, tbsp: 0.01478676478125, tsp: 0.00492892159375, 'ft³': 28.316846592, 'in³': 0.016387064 },
  speed: { 'm/s': 1, 'km/h': 1 / 3.6, mph: 0.44704, knot: 0.514444, 'ft/s': 0.3048 },
  time: { s: 1, min: 60, h: 3600, day: 86400, week: 604800, year: 31557600, ms: 0.001 },
  energy: { J: 1, kJ: 1000, cal: 4.184, kcal: 4184, Wh: 3600, kWh: 3.6e6, BTU: 1055.05585262, eV: 1.602176634e-19 },
  pressure: { Pa: 1, kPa: 1000, bar: 1e5, atm: 101325, psi: 6894.757293168, mmHg: 133.322387415, inHg: 3386.389 },
  data: { B: 1, KB: 1024, MB: 1048576, GB: 1073741824, TB: 1099511627776, bit: 0.125, Kbit: 128, Mbit: 131072 },
};

export function convertUnit(v: number, from: string, to: string): number | null {
  if (from === to) return v;
  const temps: Record<string, [(x: number) => number, (x: number) => number]> = { '°C': [(x) => x, (x) => x], '°F': [(x) => (x - 32) * 5 / 9, (x) => x * 9 / 5 + 32], K: [(x) => x - 273.15, (x) => x + 273.15] };
  if (temps[from] && temps[to]) return temps[to][1](temps[from][0](v));
  for (const cat of Object.values(UNITS)) if (cat[from] != null && cat[to] != null) return v * cat[from] / cat[to];
  return null;
}

/** A number as the text to store: its display text when it has a format of its own, otherwise every digit (never 1.23457E+15). */
function numberAsText(app: App, v: number, r: number, c: number): string {
  const fmt = app.sheet.styleOf(r, c).numFmt;
  if (fmt && fmt !== 'General' && fmt !== '@') return shownText(app.sheet, r, c);
  if (Number.isSafeInteger(v)) return String(v);
  const s = String(+v.toPrecision(15));
  return /e/i.test(s) ? (Math.abs(v) < 1e21 ? v.toLocaleString('en-US', { useGrouping: false, maximumFractionDigits: 20 }) : shownText(app.sheet, r, c)) : s;
}

export function toBase(n: number, base: number): string { return Math.trunc(n).toString(base).toUpperCase(); }
export function fromBase(s: string, base: number): number | null { const t = s.trim().replace(base === 16 ? /^0[xX]/ : base === 2 ? /^0[bB]/ : base === 8 ? /^0[oO]/ : /^(?!)/, ''); if (!t) return null; const n = parseInt(t, base); if (isNaN(n) || n < 0) return null; return n.toString(base).toUpperCase() === t.toUpperCase().replace(/^0+(?=.)/, '') ? n : null; }

function dateParts(app: App, v: number) { return serialToParts(v, app.book.calc.date1904); }

defTools([
  {
    id: 'util.numberOperation', label: 'Apply a Calculation to the Numbers…', icon: 'calculate', tab: U, group: G_NUMBERS, edits: true, popular: true, tip: 'Add, subtract, multiply, divide, raise, round or change the sign of every selected number in place (formulas can be wrapped too)', keywords: 'operation multiply divide add subtract round negate paste special', origin: 'ASAP Utilities / Kutools (Operation Tools)',
    run: async (app, arg?: { op: NumOp; x?: number; formulas?: boolean }) => {
      const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Apply a Calculation', fields: [{ key: 'op', label: 'Calculation', type: 'select', value: 'multiply', options: OP_LABELS.map((x) => ({ value: x.value, label: x.label })) }, { key: 'x', label: 'Value', type: 'number', value: 1, step: 0.01, visibleIf: (v) => !!OP_LABELS.find((x) => x.value === v.op)?.needs }, { key: 'formulas', label: 'Also wrap formulas in the calculation', type: 'checkbox', value: false }] }); return r ? { op: r.op, x: Number(r.x), formulas: !!r.formulas } : null; });
      if (!o) return;
      numberOperation(app, o.op, o.x ?? 0, !!o.formulas, 'Apply Calculation');
    },
  },
  { id: 'util.changeSign', label: 'Change the Sign of Numbers', icon: 'exposure', tab: U, group: G_NUMBERS, edits: true, popular: true, tip: 'Negative becomes positive and positive negative', origin: 'ASAP Utilities / Kutools (Change Sign of Values)', run: (app) => numberOperation(app, 'negate', 0, false, 'Change Sign') },
  { id: 'util.absValues', label: 'Make All Numbers Positive', icon: 'add_circle', tab: U, group: G_NUMBERS, edits: true, origin: 'Kutools (Change Sign of Values)', run: (app) => numberOperation(app, 'abs', 0, false, 'Absolute Values') },
  { id: 'util.percentToFraction', label: 'Divide by 100 (Percent Points to Fractions)', icon: 'percent', tab: U, group: G_NUMBERS, edits: true, tip: '45 becomes 0.45 and is formatted as a percent', run: (app) => { numberOperation(app, 'percentOf', 0, false, 'To Fractions'); patchStyle(app, 'Percent Format', { numFmt: '0.00%' }, { onlyExisting: true }); } },
  {
    id: 'util.numbersToText', label: 'Store Numbers as Text', icon: 'text_fields', tab: U, group: G_NUMBERS, edits: true, tip: 'Keep the numbers but store them as text (leading zeros, IDs, ZIP codes), with a Text number format', origin: 'ASAP Utilities / Kutools (Convert between Text and Number)',
    run: (app) => { changeCells(app, 'Numbers to Text', (cell, r, c) => (typeof cell.v === 'number' ? { ...cell, v: numberAsText(app, cell.v, r, c), s: app.book.derivedStyle(cell.s ?? 0, { numFmt: '@' }) } : undefined)); },
  },
  {
    id: 'util.padZeros', label: 'Pad with Leading Zeros…', icon: '123', tab: U, group: G_NUMBERS, edits: true, popular: true, tip: 'Make every value the same width by adding zeros in front (00042), as text or as a number format', keywords: 'leading zeros pad width fixed length', origin: 'ASAP Utilities / Kutools',
    run: async (app, arg?: { width: number; how: 'text' | 'format' }) => {
      const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Pad with Leading Zeros', fields: [{ key: 'width', label: 'Total width (digits)', type: 'number', value: 5, min: 1, max: 40 }, { key: 'how', label: 'How', type: 'radio', value: 'text', options: [{ value: 'text', label: 'Store as text (00042)' }, { value: 'format', label: 'Keep the number; only display with zeros' }] }] }); return r ? { width: Number(r.width) || 1, how: r.how } : null; });
      if (!o) return;
      if (o.how === 'format') { patchStyle(app, 'Leading Zeros', { numFmt: '0'.repeat(o.width) }, { onlyExisting: true }); return; }
      changeCells(app, 'Leading Zeros', (cell) => { const t = typeof cell.v === 'number' ? String(cell.v) : typeof cell.v === 'string' ? cell.v.trim() : ''; if (!/^-?\d+(\.\d+)?$/.test(t)) return undefined; const neg = t.startsWith('-'); const body = t.replace('-', ''); const [i, d] = body.split('.'); return { ...cell, v: `${neg ? '-' : ''}${i.padStart(o.width, '0')}${d != null ? '.' + d : ''}`, s: app.book.derivedStyle(cell.s ?? 0, { numFmt: '@' }) }; });
    },
  },
  { id: 'util.removeLeadingZeros', label: 'Remove Leading Zeros', icon: 'backspace', tab: U, group: G_NUMBERS, edits: true, origin: 'Kutools (Remove Leading Zeros)', run: (app) => transformValues(app, 'Remove Leading Zeros', (v) => (typeof v === 'string' && /^0+\d/.test(v.trim()) ? v.trim().replace(/^0+(?=\d)/, '') : undefined)) },
  {
    id: 'util.shiftDates', label: 'Add to or Subtract from Dates…', icon: 'event_repeat', tab: U, group: G_NUMBERS, edits: true, popular: true, tip: 'Move every selected date by a number of days, weeks, months or years, or set it to the start or end of its month, quarter or year', keywords: 'dates add days months years shift end of month', origin: 'ASAP Utilities / Ablebits (Date & Time Wizard)',
    run: async (app, arg?: { n: number; unit: 'days' | 'weeks' | 'months' | 'years' | 'startMonth' | 'endMonth' | 'startYear' | 'endYear' | 'startQuarter' | 'endQuarter' }) => {
      const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Shift Dates', fields: [{ key: 'unit', label: 'Change each date', type: 'select', value: 'days', options: [{ value: 'days', label: 'By a number of days' }, { value: 'weeks', label: 'By a number of weeks' }, { value: 'months', label: 'By a number of months' }, { value: 'years', label: 'By a number of years' }, { value: 'startMonth', label: 'To the first day of its month' }, { value: 'endMonth', label: 'To the last day of its month' }, { value: 'startQuarter', label: 'To the first day of its quarter' }, { value: 'endQuarter', label: 'To the last day of its quarter' }, { value: 'startYear', label: 'To January 1 of its year' }, { value: 'endYear', label: 'To December 31 of its year' }] }, { key: 'n', label: 'Amount (negative to subtract)', type: 'number', value: 1, visibleIf: (v) => ['days', 'weeks', 'months', 'years'].includes(v.unit) }] }); return r ? { n: Number(r.n) || 0, unit: r.unit } : null; });
      if (!o) return;
      const d1904 = app.book.calc.date1904;
      changeCells(app, 'Shift Dates', (cell, r, c) => {
        if (typeof cell.v !== 'number' || !isDateFormat(app.sheet.styleOf(r, c).numFmt)) return undefined;
        const frac = cell.v - Math.floor(cell.v);
        const p = serialToParts(Math.floor(cell.v), d1904);
        let y = p.y, m = p.m, d = p.d;
        switch (o.unit) {
          case 'days': return { ...cell, v: cell.v + o.n };
          case 'weeks': return { ...cell, v: cell.v + o.n * 7 };
          case 'months': { const t = (y * 12 + m - 1) + o.n; y = Math.floor(t / 12); m = t - y * 12 + 1; d = Math.min(d, daysInMonth(y, m)); break; }
          case 'years': y += o.n; d = Math.min(d, daysInMonth(y, m)); break;
          case 'startMonth': d = 1; break; case 'endMonth': d = daysInMonth(y, m); break;
          case 'startQuarter': m = Math.floor((m - 1) / 3) * 3 + 1; d = 1; break; case 'endQuarter': m = Math.floor((m - 1) / 3) * 3 + 3; d = daysInMonth(y, m); break;
          case 'startYear': m = 1; d = 1; break; case 'endYear': m = 12; d = 31; break;
        }
        return { ...cell, v: partsToSerial(y, m, d, d1904) + frac };
      });
    },
  },
  {
    id: 'util.splitDates', label: 'Split Dates into Year, Month and Day Columns', icon: 'calendar_view_month', tab: U, group: G_NUMBERS, edits: true, tip: 'Beside each date column, write its year, month, day and weekday', origin: 'ASAP Utilities',
    run: (app) => { const s = app.sheet; const rg = clipToUsed(s, activeRange(app)); const target = { r1: rg.r1, r2: rg.r2, c1: rg.c2 + 1, c2: rg.c2 + 4 }; if (!checkWritable(app, [target])) return; let n = 0; app.book.transact('Split Dates', () => { for (let r = rg.r1; r <= rg.r2; r++) { const v = s.value(r, rg.c1); if (typeof v !== 'number') { if (r === rg.r1 && typeof v === 'string') { ['Year', 'Month', 'Day', 'Weekday'].forEach((t, i) => writeOut(s, r, target.c1 + i, t)); } continue; } const p = dateParts(app, v); writeOut(s, r, target.c1, p.y); writeOut(s, r, target.c1 + 1, p.m); writeOut(s, r, target.c1 + 2, p.d); writeOut(s, r, target.c1 + 3, DAYS[p.dow]); n++; } }); toast(`${plural(n, 'date')} split`); },
  },
  {
    id: 'util.datesToText', label: 'Convert Dates to Text…', icon: 'event_note', tab: U, group: G_NUMBERS, edits: true, tip: 'Replace dates with their text in a pattern (yyyy-mm-dd, d mmmm yyyy, mm/dd/yyyy…)', origin: 'ASAP Utilities',
    run: async (app, arg?: { pattern: string }) => {
      const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Dates to Text', fields: [{ key: 'pattern', label: 'Pattern', type: 'combo', value: 'yyyy-mm-dd', options: ['yyyy-mm-dd', 'd mmmm yyyy', 'mmmm d, yyyy', 'mm/dd/yyyy', 'dd/mm/yyyy', 'dd.mm.yyyy', 'ddd d mmm yyyy', 'yyyymmdd'] }] }); return r ? { pattern: String(r.pattern) } : null; });
      if (!o) return;
      const fmt = (v: number) => { const p = dateParts(app, v); const two = (n: number) => String(n).padStart(2, '0'); const map: Record<string, string> = { yyyy: String(p.y), yy: String(p.y).slice(-2), mmmm: MONTHS[p.m - 1], mmm: MONTHS[p.m - 1].slice(0, 3), mm: two(p.m), m: String(p.m), dddd: DAYS[p.dow], ddd: DAYS[p.dow].slice(0, 3), dd: two(p.d), d: String(p.d) }; /* one pass, so the letters of an inserted month or day name are not read as codes */ return o.pattern.replace(/yyyy|yy|mmmm|mmm|mm|m|dddd|ddd|dd|d/g, (tok) => map[tok]); };
      changeCells(app, 'Dates to Text', (cell, r, c) => (typeof cell.v === 'number' && isDateFormat(app.sheet.styleOf(r, c).numFmt) ? { ...cell, v: fmt(cell.v), s: app.book.derivedStyle(cell.s ?? 0, { numFmt: '@' }) } : undefined));
    },
  },
  { id: 'util.weekdayNames', label: 'Weekday Names Beside Dates', icon: 'today', tab: U, group: G_NUMBERS, edits: true, run: (app) => { const s = app.sheet; const rg = clipToUsed(s, activeRange(app)); const target = { r1: rg.r1, r2: rg.r2, c1: rg.c2 + 1, c2: rg.c2 + 1 }; if (!checkWritable(app, [target])) return; app.book.transact('Weekday Names', () => { for (let r = rg.r1; r <= rg.r2; r++) { const v = s.value(r, rg.c1); if (typeof v === 'number') writeOut(s, r, target.c1, DAYS[dateParts(app, v).dow]); } }); } },
  { id: 'util.timeToHours', label: 'Convert Times to Decimal Hours', icon: 'schedule', tab: U, group: G_NUMBERS, edits: true, kind: 'split', tip: '8:30 becomes 8.5 (or 510 minutes, or 30600 seconds)', origin: 'Kutools (Convert Time)', run: (app, arg?: 'hours' | 'minutes' | 'seconds') => { const k = arg === 'minutes' ? 1440 : arg === 'seconds' ? 86400 : 24; changeCells(app, 'Time to Number', (cell) => (typeof cell.v === 'number' ? { ...cell, v: Math.round(cell.v * k * 1e6) / 1e6, s: app.book.derivedStyle(cell.s ?? 0, { numFmt: '0.##' }) } : undefined)); }, menu: () => [{ id: 'util.timeToHours', arg: 'hours', label: 'To hours (8:30 → 8.5)' }, { id: 'util.timeToHours', arg: 'minutes', label: 'To minutes (8:30 → 510)' }, { id: 'util.timeToHours', arg: 'seconds', label: 'To seconds' }] },
  { id: 'util.hoursToTime', label: 'Convert Decimal Hours to Times', icon: 'timer', tab: U, group: G_NUMBERS, edits: true, tip: '8.5 becomes 8:30', run: (app) => changeCells(app, 'Hours to Time', (cell) => (typeof cell.v === 'number' ? { ...cell, v: cell.v / 24, s: app.book.derivedStyle(cell.s ?? 0, { numFmt: '[h]:mm' }) } : undefined)) },
  {
    id: 'util.convertUnits', label: 'Convert Units…', icon: 'straighten', tab: U, group: G_NUMBERS, edits: true, popular: true, tip: 'Convert the selected numbers between units of length, mass, area, volume, speed, time, temperature, energy, pressure and data size', keywords: 'convert units metric imperial cm inches kg lb celsius', origin: 'Kutools (Unit Conversion)',
    run: async (app, arg?: { from: string; to: string; dest?: 'replace' | 'right' }) => {
      const all = [...new Set([...Object.values(UNITS).flatMap((c) => Object.keys(c)), '°C', '°F', 'K'])];
      const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Convert Units', fields: [{ key: 'from', label: 'From', type: 'combo', value: 'cm', options: all, half: true }, { key: 'to', label: 'To', type: 'combo', value: 'in', options: all, half: true }, { key: 'dest', label: 'Put the result', type: 'radio', value: 'replace', options: [{ value: 'replace', label: 'In place' }, { value: 'right', label: 'In the column to the right' }] }] }); return r ? { from: String(r.from).trim(), to: String(r.to).trim(), dest: r.dest } : null; });
      if (!o) return;
      if (convertUnit(1, o.from, o.to) == null) { toast(`Cannot convert ${o.from} to ${o.to}.`); return; }
      if (o.dest === 'right') { const s = app.sheet; const rg = clipToUsed(s, activeRange(app)); const target = { r1: rg.r1, r2: rg.r2, c1: rg.c2 + 1, c2: rg.c2 + 1 + (rg.c2 - rg.c1) }; if (!checkWritable(app, [target])) return; app.book.transact('Convert Units', () => { for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) { const v = s.value(r, c); if (typeof v === 'number') writeOut(s, r, target.c1 + c - rg.c1, convertUnit(v, o.from, o.to)!); } }); return; }
      changeCells(app, 'Convert Units', (cell) => (typeof cell.v === 'number' ? { ...cell, v: convertUnit(cell.v, o.from, o.to)! } : undefined));
    },
  },
  {
    id: 'util.numberBase', label: 'Convert Number Base…', icon: 'tag', tab: U, group: G_NUMBERS, edits: true, tip: 'Decimal, binary, octal, hexadecimal and Roman numerals, in place', origin: 'Kutools (Number Systems Conversion)',
    run: async (app, arg?: { from: number; to: number }) => {
      const bases = [{ value: '10', label: 'Decimal' }, { value: '2', label: 'Binary' }, { value: '8', label: 'Octal' }, { value: '16', label: 'Hexadecimal' }, { value: '36', label: 'Base 36' }, { value: '0', label: 'Roman numerals' }];
      const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Convert Number Base', fields: [{ key: 'from', label: 'From', type: 'select', value: '10', options: bases, half: true }, { key: 'to', label: 'To', type: 'select', value: '16', options: bases, half: true }] }); return r ? { from: Number(r.from), to: Number(r.to) } : null; });
      if (!o) return;
      changeCells(app, 'Convert Base', (cell) => {
        const t = typeof cell.v === 'number' ? String(cell.v) : typeof cell.v === 'string' ? cell.v : null; if (t == null) return undefined;
        const n = o.from === 0 ? fromRoman(t) : o.from === 10 ? numOf(t) : fromBase(t, o.from);
        if (n == null) return undefined;
        if (o.to === 10) return { ...cell, v: n };
        if (o.to === 0) return n >= 1 && n <= 3999 ? { ...cell, v: toRoman(Math.trunc(n)) } : undefined;
        return { ...cell, v: toBase(n, o.to), s: app.book.derivedStyle(cell.s ?? 0, { numFmt: '@' }) };
      });
    },
  },
  {
    id: 'util.sumBelow', label: 'Insert SUM Below Each Column', icon: 'functions', tab: U, group: G_NUMBERS, edits: true, popular: true, kind: 'split', tip: 'A SUM (or AVERAGE, COUNT, MAX, MIN) formula under every selected column, and to the right of every row', origin: 'ASAP Utilities / Kutools (AutoSum variants)',
    run: (app, arg?: { fn?: string; where?: 'below' | 'right' | 'both' }) => {
      const s = app.sheet; const rg = clipToUsed(s, activeRange(app)); const fn = arg?.fn || 'SUM'; const where = arg?.where || 'below';
      const targets: Range[] = []; if (where !== 'right') targets.push({ r1: rg.r2 + 1, r2: rg.r2 + 1, c1: rg.c1, c2: rg.c2 }); if (where !== 'below') targets.push({ r1: rg.r1, r2: rg.r2, c1: rg.c2 + 1, c2: rg.c2 + 1 });
      if (!checkWritable(app, targets)) return;
      app.book.transact('Insert Totals', () => { if (where !== 'right') for (let c = rg.c1; c <= rg.c2; c++) writeOut(s, rg.r2 + 1, c, { f: `${fn}(${cellName(rg.r1, c)}:${cellName(rg.r2, c)})`, style: { bold: true, bt: { style: 'thin' } } }); if (where !== 'below') for (let r = rg.r1; r <= rg.r2; r++) writeOut(s, r, rg.c2 + 1, { f: `${fn}(${cellName(r, rg.c1)}:${cellName(r, rg.c2)})`, style: { bold: true } }); });
    },
    menu: () => ['SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN'].flatMap((fn) => [{ id: 'util.sumBelow', arg: { fn, where: 'below' }, label: `${fn} below each column` }, { id: 'util.sumBelow', arg: { fn, where: 'right' }, label: `${fn} right of each row` }]),
  },
  { id: 'util.menuNumbers', label: 'Numbers & Dates', icon: 'pin', tab: U, group: G_NUMBERS, kind: 'menu', noRecord: true, run: () => undefined, menu: () => menuOf(['util.numberOperation', 'util.changeSign', 'util.absValues', 'util.percentToFraction', 'tools.roundValues', 'util.sumBelow', null, 'tools.textToNumbers', 'util.numbersToText', 'util.padZeros', 'util.removeLeadingZeros', 'util.numberBase', 'tools.romanNumerals', 'tools.numberToWords', 'util.convertUnits', null, 'tools.fixDates', 'util.shiftDates', 'util.splitDates', 'util.datesToText', 'util.weekdayNames', 'util.timeToHours', 'util.hoursToTime', 'tools.dateCalculator', 'tools.calendar']) },
]);

// ============================================================================================
// Text

export function insertAt(s: string, text: string, where: 'start' | 'end' | 'position' | 'beforeUpper' | 'everyN' | 'beforeNumber', n = 1): string {
  switch (where) {
    case 'start': return text + s;
    case 'end': return s + text;
    case 'position': { const i = Math.max(0, Math.min(s.length, n - 1)); return s.slice(0, i) + text + s.slice(i); }
    case 'beforeUpper': return s.replace(/(?<=\p{Ll})(\p{Lu})/gu, text + '$1');
    case 'beforeNumber': return s.replace(/(?<=\D)(\d)/g, text + '$1').replace(/^(\d)/, text + '$1');
    case 'everyN': { if (n < 1) return s; const out: string[] = []; for (let i = 0; i < s.length; i += n) out.push(s.slice(i, i + n)); return out.join(text); }
  }
}

export function removeByPosition(s: string, mode: 'first' | 'last' | 'from' | 'before' | 'after', n: number, len = 1, marker = ''): string {
  switch (mode) {
    case 'first': return s.slice(n);
    case 'last': return s.slice(0, Math.max(0, s.length - n));
    case 'from': return s.slice(0, Math.max(0, n - 1)) + s.slice(Math.max(0, n - 1) + len);
    case 'before': { const i = s.indexOf(marker); return i < 0 ? s : s.slice(i + marker.length); }
    case 'after': { const i = s.indexOf(marker); return i < 0 ? s : s.slice(0, i); }
  }
}

export function nthWord(s: string, n: number, sep = ' '): string { const words = sep === ' ' ? s.trim().split(/\s+/) : s.split(sep); const i = n < 0 ? words.length + n : n - 1; return words[i] ?? ''; }
export function between(s: string, a: string, b: string): string { const i = a ? s.indexOf(a) : 0; if (i < 0) return ''; const start = i + a.length; const j = b ? s.indexOf(b, start) : s.length; return j < 0 ? '' : s.slice(start, j); }
export function swapAround(s: string, sep: string, join = ' '): string { const i = s.indexOf(sep); if (i < 0) return s; return `${s.slice(i + sep.length).trim()}${join}${s.slice(0, i).trim()}`; }
export function keepOnly(s: string, what: 'digits' | 'letters' | 'alnum' | 'custom', chars = ''): string {
  if (what === 'digits') return s.replace(/[^\d.,-]/g, '');
  if (what === 'letters') return s.replace(/[^\p{L}\s]/gu, '');
  if (what === 'alnum') return s.replace(/[^\p{L}\p{N}\s]/gu, '');
  const set = new Set([...chars]); return [...s].filter((ch) => set.has(ch)).join('');
}
export const wordCount = (s: string): number => (s.trim() ? s.trim().split(/\s+/).length : 0);

defTools([
  {
    id: 'util.insertText', label: 'Insert Text at a Position…', icon: 'add_notes', tab: U, group: G_TEXT, edits: true, popular: true, tip: 'Add text at the start, the end, a character position, before each capital letter, before the first digit, or every N characters', keywords: 'add text position prefix suffix insert characters', origin: 'ASAP Utilities / Kutools (Add Text)',
    run: async (app, arg?: { text: string; where: 'start' | 'end' | 'position' | 'beforeUpper' | 'everyN' | 'beforeNumber'; n?: number }) => {
      const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Insert Text', fields: [{ key: 'text', label: 'Text to insert', type: 'text', value: '', autofocus: true }, { key: 'where', label: 'Where', type: 'select', value: 'start', options: [{ value: 'start', label: 'At the start' }, { value: 'end', label: 'At the end' }, { value: 'position', label: 'Before character number…' }, { value: 'everyN', label: 'Every N characters' }, { value: 'beforeUpper', label: 'Before each capital letter (camelCase → camel Case)' }, { value: 'beforeNumber', label: 'Before the first digit' }] }, { key: 'n', label: 'N', type: 'number', value: 3, min: 1, visibleIf: (v) => v.where === 'position' || v.where === 'everyN' }] }); return r ? { text: String(r.text ?? ''), where: r.where, n: Number(r.n) || 1 } : null; });
      if (!o) return;
      transformValues(app, 'Insert Text', (v) => insertAt(typeof v === 'number' ? String(v) : String(v), o.text, o.where, o.n), { numbersToo: true });
    },
  },
  {
    id: 'util.removeByPosition', label: 'Remove Characters by Position…', icon: 'backspace', tab: U, group: G_TEXT, edits: true, popular: true, tip: 'Delete the first or last N characters, a run of characters from a position, or everything before or after a marker', keywords: 'remove first last characters position substring', origin: 'ASAP Utilities / Kutools (Remove by Position) / Ablebits',
    run: async (app, arg?: { mode: 'first' | 'last' | 'from' | 'before' | 'after'; n?: number; len?: number; marker?: string }) => {
      const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Remove Characters by Position', fields: [{ key: 'mode', label: 'Remove', type: 'select', value: 'first', options: [{ value: 'first', label: 'The first N characters' }, { value: 'last', label: 'The last N characters' }, { value: 'from', label: 'N characters from a position' }, { value: 'before', label: 'Everything before a marker (and the marker)' }, { value: 'after', label: 'Everything after a marker (and the marker)' }] }, { key: 'n', label: 'N / position', type: 'number', value: 1, min: 1, half: true, visibleIf: (v) => v.mode !== 'before' && v.mode !== 'after' }, { key: 'len', label: 'How many', type: 'number', value: 1, min: 1, half: true, visibleIf: (v) => v.mode === 'from' }, { key: 'marker', label: 'Marker text', type: 'text', value: '', visibleIf: (v) => v.mode === 'before' || v.mode === 'after' }] }); return r ? { mode: r.mode, n: Number(r.n) || 1, len: Number(r.len) || 1, marker: r.marker || '' } : null; });
      if (!o) return;
      transformValues(app, 'Remove by Position', (v) => removeByPosition(String(v), o.mode, o.n || 1, o.len, o.marker));
    },
  },
  { id: 'util.reverseText', label: 'Reverse the Characters', icon: 'swap_horiz', tab: U, group: G_TEXT, edits: true, tip: 'abc becomes cba', origin: 'ASAP Utilities / Kutools (Reverse Text Order)', run: (app) => transformValues(app, 'Reverse Text', (v) => [...String(v)].reverse().join('')) },
  { id: 'util.reverseWords', label: 'Reverse the Word Order', icon: 'swap_horiz', tab: U, group: G_TEXT, edits: true, tip: '"John Smith" becomes "Smith John"', origin: 'Kutools', run: (app) => transformValues(app, 'Reverse Words', (v) => String(v).trim().split(/\s+/).reverse().join(' ')) },
  {
    id: 'util.swapAround', label: 'Swap Text Around a Separator…', icon: 'swap_horizontal_circle', tab: U, group: G_TEXT, edits: true, popular: true, tip: '"Smith, John" becomes "John Smith" (or the reverse): exchange the parts before and after a comma, dash or any text', keywords: 'last first name order comma', origin: 'Genealogy and mailing-list practice',
    run: async (app, arg?: { sep: string; join: string }) => {
      const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Swap Around a Separator', fields: [{ key: 'sep', label: 'Separator', type: 'combo', value: ',', options: [',', ' ', '-', ';', '/', ':'], half: true }, { key: 'join', label: 'Join the parts with', type: 'combo', value: ' ', options: [' ', ', ', ' - ', ''], half: true }] }); return r ? { sep: String(r.sep), join: String(r.join) } : null; });
      if (!o) return;
      transformValues(app, 'Swap Around Separator', (v) => swapAround(String(v), o.sep, o.join));
    },
  },
  {
    id: 'util.keepOnly', label: 'Keep Only Digits, Letters…', icon: 'filter_alt', tab: U, group: G_TEXT, edits: true, tip: 'Strip everything from the text except digits, letters, letters and digits, or the characters you list', origin: 'ASAP Utilities / Ablebits (Extract Text)',
    run: async (app, arg?: { what: 'digits' | 'letters' | 'alnum' | 'custom'; chars?: string }) => {
      const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Keep Only', fields: [{ key: 'what', label: 'Keep', type: 'radio', value: 'digits', options: [{ value: 'digits', label: 'Digits (and . , -)' }, { value: 'letters', label: 'Letters and spaces' }, { value: 'alnum', label: 'Letters, digits and spaces' }, { value: 'custom', label: 'These characters:' }] }, { key: 'chars', label: 'Characters', type: 'text', value: '', visibleIf: (v) => v.what === 'custom' }] }); return r ? { what: r.what, chars: r.chars } : null; });
      if (!o) return;
      transformValues(app, 'Keep Only', (v) => { const out = keepOnly(String(v), o.what, o.chars); if (o.what === 'digits' && /^-?(0|[1-9]\d*)(\.\d+)?$/.test(out)) return Number(out); return out; }, { numbersToo: false });
    },
  },
  {
    id: 'util.extractWord', label: 'Extract a Word, or Text Between Markers…', icon: 'content_cut', tab: U, group: G_TEXT, edits: true, popular: true, tip: 'Put the Nth word (or the last), or the text between two markers, in the column to the right', keywords: 'nth word last word extract between substring', origin: 'ASAP Utilities / Ablebits (Extract Text)',
    run: async (app, arg?: { mode: 'word' | 'between' | 'before' | 'after'; n?: number; a?: string; b?: string }) => {
      const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Extract Text', fields: [{ key: 'mode', label: 'Extract', type: 'select', value: 'word', options: [{ value: 'word', label: 'Word number N (negative = from the end)' }, { value: 'between', label: 'Text between two markers' }, { value: 'before', label: 'Text before a marker' }, { value: 'after', label: 'Text after a marker' }] }, { key: 'n', label: 'N', type: 'number', value: 1, visibleIf: (v) => v.mode === 'word' }, { key: 'a', label: 'Start marker', type: 'text', value: '', visibleIf: (v) => v.mode !== 'word', half: true }, { key: 'b', label: 'End marker', type: 'text', value: '', visibleIf: (v) => v.mode === 'between', half: true }] }); return r ? { mode: r.mode, n: Number(r.n) || 1, a: r.a || '', b: r.b || '' } : null; });
      if (!o) return;
      const s = app.sheet; const rg = clipToUsed(s, activeRange(app)); const target = { r1: rg.r1, r2: rg.r2, c1: rg.c2 + 1, c2: rg.c2 + 1 + (rg.c2 - rg.c1) };
      if (!checkWritable(app, [target])) return;
      app.book.transact('Extract Text', () => { for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) { const t = shownText(s, r, c); if (!t) continue; const out = o.mode === 'word' ? nthWord(t, o.n || 1) : o.mode === 'between' ? between(t, o.a || '', o.b || '') : o.mode === 'before' ? (t.includes(o.a || '') && o.a ? t.slice(0, t.indexOf(o.a)) : '') : (o.a && t.includes(o.a) ? t.slice(t.indexOf(o.a) + o.a.length) : ''); writeOut(s, r, target.c1 + c - rg.c1, out); } });
      app.grid.select(target.r1, target.c1, { range: target });
    },
  },
  { id: 'util.lineBreaksToSpaces', label: 'Replace Line Breaks with Spaces', icon: 'wrap_text', tab: U, group: G_TEXT, edits: true, origin: 'ASAP Utilities', run: (app) => transformValues(app, 'Line Breaks to Spaces', (v) => String(v).replace(/\r?\n/g, ' ').replace(/ {2,}/g, ' ').trim()) },
  { id: 'util.separatorToLineBreaks', label: 'Replace a Separator with Line Breaks…', icon: 'keyboard_return', tab: U, group: G_TEXT, edits: true, run: async (app, arg?: string) => { const sep = await argOrAsk(arg, () => promptDialog('Separator to Line Breaks', 'Separator to replace with a line break', ', ')); if (sep == null) return; transformValues(app, 'Separator to Line Breaks', (v) => String(v).split(sep).join('\n')); patchStyle(app, 'Wrap', { wrap: true }, { onlyExisting: true }); } },
  { id: 'util.removeAllSpaces', label: 'Remove All Spaces', icon: 'space_bar', tab: U, group: G_TEXT, edits: true, origin: 'Kutools (Remove Spaces)', run: (app) => transformValues(app, 'Remove All Spaces', (v) => String(v).replace(/[\s ]+/g, '')) },
  { id: 'util.removeLeadingSpaces', label: 'Remove Leading Spaces Only', icon: 'format_indent_decrease', tab: U, group: G_TEXT, edits: true, origin: 'ASAP Utilities', run: (app) => transformValues(app, 'Remove Leading Spaces', (v) => String(v).replace(/^[\s ]+/, '')) },
  { id: 'util.removeTrailingSpaces', label: 'Remove Trailing Spaces Only', icon: 'format_indent_increase', tab: U, group: G_TEXT, edits: true, origin: 'ASAP Utilities', run: (app) => transformValues(app, 'Remove Trailing Spaces', (v) => String(v).replace(/[\s ]+$/, '')) },
  {
    id: 'util.padText', label: 'Pad Text to a Width…', icon: 'format_align_justify', tab: U, group: G_TEXT, edits: true, tip: 'Make every text the same length by adding a character on the left or right (for fixed-width exports)', origin: 'ASAP Utilities',
    run: async (app, arg?: { width: number; char: string; side: 'left' | 'right' }) => {
      const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Pad Text', fields: [{ key: 'width', label: 'Width', type: 'number', value: 10, min: 1, max: 500, half: true }, { key: 'char', label: 'Pad with', type: 'text', value: ' ', half: true }, { key: 'side', label: 'Side', type: 'radio', value: 'right', options: [{ value: 'right', label: 'Right (left-aligned text)' }, { value: 'left', label: 'Left (right-aligned text)' }] }] }); return r ? { width: Number(r.width) || 1, char: String(r.char || ' ')[0], side: r.side } : null; });
      if (!o) return;
      transformValues(app, 'Pad Text', (v) => { const t = typeof v === 'number' ? String(v) : String(v); return o.side === 'left' ? t.padStart(o.width, o.char) : t.padEnd(o.width, o.char); }, { numbersToo: true });
    },
  },
  { id: 'util.textLengths', label: 'Text Length Beside Each Cell', icon: 'straighten', tab: U, group: G_TEXT, edits: true, kind: 'split', tip: 'The number of characters (or words) of each selected cell, in the column to the right', origin: 'ASAP Utilities / Kutools (Count Words)', run: (app, arg?: 'chars' | 'words') => { const s = app.sheet; const rg = clipToUsed(s, activeRange(app)); const target = { r1: rg.r1, r2: rg.r2, c1: rg.c2 + 1, c2: rg.c2 + 1 }; if (!checkWritable(app, [target])) return; app.book.transact('Text Length', () => { for (let r = rg.r1; r <= rg.r2; r++) { const t = shownText(s, r, rg.c1); if (t) writeOut(s, r, target.c1, arg === 'words' ? wordCount(t) : [...t].length); } }); }, menu: () => [{ id: 'util.textLengths', arg: 'chars', label: 'Characters' }, { id: 'util.textLengths', arg: 'words', label: 'Words' }] },
  { id: 'util.trimWorkbook', label: 'Trim Spaces in the Whole Workbook', icon: 'cleaning_services', tab: U, group: G_TEXT, edits: true, origin: 'ASAP Utilities', run: (app) => { let n = 0; app.book.transact('Trim Workbook', () => { for (const s of app.book.sheets) for (const [k, cell] of [...s.cells]) { if (typeof cell.v !== 'string' || cell.f != null) continue; const t = cell.v.replace(/[   ]/g, ' ').replace(/ {2,}/g, ' ').trim(); if (t !== cell.v) { const next: Cell = { ...cell, v: t }; delete next.rich; s.set(keyRow(k), keyCol(k), next); n++; } } }); toast(`${plural(n, 'cell')} trimmed`); } },
  { id: 'util.menuText', label: 'Text', icon: 'abc', tab: U, group: G_TEXT, kind: 'menu', noRecord: true, run: () => undefined, menu: () => menuOf(['tools.changeCase', 'tools.trimSpaces', 'util.removeAllSpaces', 'util.removeLeadingSpaces', 'util.removeTrailingSpaces', 'util.trimWorkbook', 'tools.cleanText', 'tools.removeAccents', null, 'util.insertText', 'tools.addText', 'util.removeByPosition', 'tools.removeChars', 'util.keepOnly', 'util.padText', null, 'util.extractWord', 'tools.extract', 'tools.splitColumn', 'tools.mergeColumns', 'util.swapAround', 'util.reverseText', 'util.reverseWords', null, 'util.lineBreaksToSpaces', 'util.separatorToLineBreaks', 'util.textLengths', 'tools.ocrCleanup', 'tools.numberToWords']) },
]);

// ============================================================================================
// Formulas

export type RefStyle = 'absolute' | 'relative' | 'rowAbs' | 'colAbs' | 'toggle';

export function changeRefStyle(f: string, style: RefStyle): string {
  return mapRefs(f, (ref) => {
    const r = { ...ref };
    const set = (ar: boolean, ac: boolean) => { r.ar1 = ar; r.ar2 = ar; r.ac1 = ac; r.ac2 = ac; };
    if (style === 'absolute') set(true, true); else if (style === 'relative') set(false, false); else if (style === 'rowAbs') set(true, false); else if (style === 'colAbs') set(false, true);
    else set(!(ref.ar1 && ref.ac1), !(ref.ar1 && ref.ac1));
    return r;
  });
}

/** Replaces defined names in a formula with their references (names scoped to the sheet first). */
export function expandNames(f: string, names: { name: string; ref: string }[]): string {
  const toks = tokenize(f);
  let out = ''; let last = 0;
  for (const t of toks) {
    if (t.type !== 'name') continue;
    const nm = names.find((n) => n.name.toLowerCase() === t.text.toLowerCase());
    if (!nm) continue;
    out += f.slice(last, t.pos) + '(' + nm.ref.replace(/^=/, '') + ')';
    last = t.end;
  }
  return out + f.slice(last);
}

function refTextNoSheet(r: { r1: number; c1: number; r2: number; c2: number; kind: string }): string { return r.kind === 'cell' || (r.r1 === r.r2 && r.c1 === r.c2) ? cellName(r.r1, r.c1) : `${cellName(r.r1, r.c1)}:${cellName(r.r2, r.c2)}`; }

/** Text-level replacement of references by names (used by Replace References with Names). */
export function applyNamesText(f: string, names: { name: string; ref: string; sheet?: string | null }[], homeSheet: string): string {
  const toks = tokenize(f); let out = ''; let last = 0;
  for (const t of toks) {
    if (t.type !== 'ref' || !t.ref || t.ref.invalid) continue;
    const ref = t.ref;
    const plain = refTextNoSheet(ref).toUpperCase();
    const hit = names.find((n) => { const m = /^=?(?:'([^']+)'|([^!']+))?!?(.+)$/.exec(n.ref.replace(/^=/, '')); if (!m) return false; const sh = m[1] || m[2]; const body = m[3].replace(/\$/g, '').toUpperCase(); const refSheet = (ref.sheet || homeSheet).toLowerCase(); if (sh && sh.toLowerCase() !== refSheet) return false; if (!sh && refSheet !== homeSheet.toLowerCase()) return false; return body === plain; });
    if (!hit) continue;
    out += f.slice(last, t.pos) + hit.name; last = t.end;
  }
  return out + f.slice(last);
}

/** Exact formula buffer for Copy/Paste Formulas Exactly. */
let exactBuffer: { cells: (Cell | undefined)[][]; from: string } | null = null;

const WRAPS: { value: string; label: string; make: (f: string, x: string) => string; needs?: boolean }[] = [
  { value: 'iferror', label: 'IFERROR(…, value)', make: (f, x) => `IFERROR(${f},${x || '""'})`, needs: true }, { value: 'round', label: 'ROUND(…, decimals)', make: (f, x) => `ROUND(${f},${x || 2})`, needs: true }, { value: 'abs', label: 'ABS(…)', make: (f) => `ABS(${f})` },
  { value: 'text', label: 'TEXT(…, format)', make: (f, x) => `TEXT(${f},"${x || '0.00'}")`, needs: true }, { value: 'ifblank', label: 'IF(…="", "", …)', make: (f) => `IF(${f}="","",${f})` }, { value: 'n', label: 'N(…)', make: (f) => `N(${f})` }, { value: 'sum', label: 'SUM(…)', make: (f) => `SUM(${f})` },
  { value: 'custom', label: 'Custom: use {f} for the formula', make: (f, x) => (x || '{f}').replace(/\{f\}/g, f), needs: true },
];

defTools([
  {
    id: 'util.refStyle', label: 'Change Reference Style…', icon: 'attach_money', tab: U, group: G_FORMULAS, edits: true, popular: true, kind: 'split', tip: 'Make every reference in the selected formulas absolute ($A$1), relative (A1), row-absolute (A$1) or column-absolute ($A1)', keywords: 'absolute relative references dollar F4 all formulas', origin: 'ASAP Utilities / Kutools (Convert Refers) / Ablebits',
    run: async (app, arg?: RefStyle) => {
      const style = await argOrAsk(arg, async () => { const i = await choiceDialog('Change Reference Style', 'Make the references in the selected formulas:', ['Absolute ($A$1)', 'Relative (A1)', 'Row absolute (A$1)', 'Column absolute ($A1)']); return i < 0 ? null : (['absolute', 'relative', 'rowAbs', 'colAbs'] as RefStyle[])[i]; });
      if (!style) return;
      changeCells(app, 'Reference Style', (cell) => (cell.f != null ? { ...cell, f: changeRefStyle(cell.f, style) } : undefined), { formulas: true });
    },
    menu: () => [{ id: 'util.refStyle', arg: 'absolute', label: 'Absolute ($A$1)' }, { id: 'util.refStyle', arg: 'relative', label: 'Relative (A1)' }, { id: 'util.refStyle', arg: 'rowAbs', label: 'Row absolute (A$1)' }, { id: 'util.refStyle', arg: 'colAbs', label: 'Column absolute ($A1)' }],
  },
  { id: 'util.copyExact', label: 'Copy Formulas Exactly', icon: 'content_copy', tab: U, group: G_FORMULAS, noRecord: true, keys: ['Mod-Alt-c'], tip: 'Copy the selection so that Paste Formulas Exactly puts the same formulas elsewhere without adjusting their references', origin: 'ASAP Utilities / Kutools (Exact Copy)', run: (app) => { const s = app.sheet; const rg = clipToUsed(s, activeRange(app)); exactBuffer = { cells: blockOf(s, rg), from: rangeName(rg) }; toast(`${rangeName(rg)} copied (exact)`); } },
  { id: 'util.pasteExact', label: 'Paste Formulas Exactly', icon: 'content_paste', tab: U, group: G_FORMULAS, edits: true, keys: ['Mod-Alt-v'], tip: 'Paste what Copy Formulas Exactly copied, formulas unchanged', origin: 'ASAP Utilities / Kutools (Exact Copy)', run: (app) => { if (!exactBuffer) { toast('Use Copy Formulas Exactly first.'); return; } const s = app.sheet; const { r, c } = activeCell(app); const b = exactBuffer; const target = { r1: r, c1: c, r2: r + b.cells.length - 1, c2: c + (b.cells[0]?.length || 1) - 1 }; if (!checkWritable(app, [target])) return; app.book.transact('Paste Exact', () => b.cells.forEach((row, i) => row.forEach((cell, j) => { const x = cell ? { ...cell } : undefined; if (x) delete x.sp; s.set(r + i, c + j, x); }))); app.grid.select(r, c, { range: target }); } },
  { id: 'util.formulasToText', label: 'Convert Formulas to Text', icon: 'code', tab: U, group: G_FORMULAS, edits: true, tip: 'Show the formula itself (=SUM(A1:A5)) as text, so it can be edited, searched or moved without changing', origin: 'ASAP Utilities / Kutools (Convert Formula to Text)', run: (app) => changeCells(app, 'Formulas to Text', (cell) => (cell.f != null ? { ...(cell.s ? { s: cell.s } : {}), v: '=' + cell.f } : undefined), { formulas: true }) },
  { id: 'util.textToFormulas', label: 'Convert Text to Formulas', icon: 'function', tab: U, group: G_FORMULAS, edits: true, tip: 'Text that starts with = becomes a working formula again', origin: 'ASAP Utilities / Kutools (Convert Text to Formula)', run: (app) => changeCells(app, 'Text to Formulas', (cell) => (typeof cell.v === 'string' && cell.v.startsWith('=') && cell.v.length > 1 ? { ...(cell.s ? { s: cell.s } : {}), f: cell.v.slice(1) } : undefined)) },
  {
    id: 'util.wrapFormulas', label: 'Wrap Formulas in a Function…', icon: 'data_object', tab: U, group: G_FORMULAS, edits: true, popular: true, tip: 'Put IFERROR, ROUND, ABS, TEXT or your own expression around every selected formula (constants can be wrapped too)', keywords: 'iferror round wrap formulas all at once', origin: 'ASAP Utilities / Kutools (Formula Helper)',
    run: async (app, arg?: { wrap: string; x?: string; constants?: boolean }) => {
      const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Wrap Formulas', fields: [{ key: 'wrap', label: 'Wrap in', type: 'select', value: 'iferror', options: WRAPS.map((w) => ({ value: w.value, label: w.label })) }, { key: 'x', label: 'Argument / expression', type: 'text', value: '', visibleIf: (v) => !!WRAPS.find((w) => w.value === v.wrap)?.needs }, { key: 'constants', label: 'Also wrap constants (numbers become formulas)', type: 'checkbox', value: false }] }); return r ? { wrap: r.wrap, x: r.x, constants: !!r.constants } : null; });
      if (!o) return;
      const w = WRAPS.find((x) => x.value === o.wrap)!;
      changeCells(app, 'Wrap Formulas', (cell) => { if (cell.f != null) return { ...cell, f: w.make(cell.f, o.x || '') }; if (!o.constants || cell.v == null) return undefined; const lit = typeof cell.v === 'number' ? String(cell.v) : typeof cell.v === 'boolean' ? String(cell.v).toUpperCase() : `"${String(cell.v).replace(/"/g, '""')}"`; return { ...(cell.s ? { s: cell.s } : {}), f: w.make(lit, o.x || '') }; }, { formulas: true });
    },
  },
  { id: 'util.namesToRefs', label: 'Replace Names with Their References', icon: 'label_off', tab: U, group: G_FORMULAS, edits: true, tip: 'In the selected formulas, replace each defined name with the range it stands for', origin: 'ASAP Utilities / Kutools (Replace Range Names)', run: (app) => { const names = app.book.names.filter((n) => !n.sheet || n.sheet === app.sheet.name); changeCells(app, 'Names to References', (cell) => { if (cell.f == null) return undefined; const f = expandNames(cell.f, names); return f === cell.f ? undefined : { ...cell, f }; }, { formulas: true }); } },
  { id: 'util.refsToNames', label: 'Replace References with Names', icon: 'label', tab: U, group: G_FORMULAS, edits: true, tip: 'In the selected formulas, use a defined name wherever a reference matches one (Apply Names)', origin: 'ASAP Utilities (Apply names)', run: (app) => { const names = app.book.names.filter((n) => !n.sheet || n.sheet === app.sheet.name); if (!names.length) { toast('The workbook has no defined names.'); return; } changeCells(app, 'Apply Names', (cell) => { if (cell.f == null) return undefined; const f = applyNamesText(cell.f, names, app.sheet.name); return f === cell.f ? undefined : { ...cell, f }; }, { formulas: true }); } },
  { id: 'util.formulaTextRight', label: 'Show Formulas as Text in the Next Column', icon: 'subtitles', tab: U, group: G_FORMULAS, edits: true, tip: 'Document a column of formulas: their text goes into the column to the right', origin: 'ASAP Utilities', run: (app) => { const s = app.sheet; const rg = clipToUsed(s, activeRange(app)); const target = { r1: rg.r1, r2: rg.r2, c1: rg.c2 + 1, c2: rg.c2 + 1 }; if (!checkWritable(app, [target])) return; let n = 0; app.book.transact('Formula Text', () => { for (let r = rg.r1; r <= rg.r2; r++) { const cell = s.get(r, rg.c1); if (cell?.f != null) { writeOut(s, r, target.c1, '=' + cell.f); n++; } } }); toast(`${plural(n, 'formula')} written`); } },
  {
    id: 'util.shiftRefs', label: 'Shift All References in Formulas…', icon: 'open_with', tab: U, group: G_FORMULAS, edits: true, tip: 'Move every relative reference in the selected formulas by a number of rows and columns without moving the cells', origin: 'ASAP Utilities',
    run: async (app, arg?: { dr: number; dc: number }) => { const o = await argOrAsk(arg, async () => { const r = await openDialog({ title: 'Shift References', fields: [{ key: 'dr', label: 'Rows (negative = up)', type: 'number', value: 0, half: true }, { key: 'dc', label: 'Columns (negative = left)', type: 'number', value: 0, half: true }] }); return r ? { dr: Number(r.dr) || 0, dc: Number(r.dc) || 0 } : null; }); if (!o || (!o.dr && !o.dc)) return; changeCells(app, 'Shift References', (cell) => (cell.f != null ? { ...cell, f: shiftFormula(cell.f, o.dr, o.dc) } : undefined), { formulas: true }); },
  },
  {
    id: 'util.colorByContentType', label: 'Color Cells by Content Type', icon: 'palette', tab: U, group: G_FORMULAS, edits: true, popular: true, tip: 'The classic model audit: formulas blue, constants black, links to other sheets green, text labels as they are', keywords: 'audit formulas constants colors', origin: 'Financial modelling convention',
    run: (app) => { const s = app.sheet; const u = s.usedRange(); if (!u) return; let n = 0; app.book.transact('Color by Content', () => { s.forEachIn(singleCell(app) ? u : clipToUsed(s, activeRange(app)), (r, c, cell) => { let color: string | undefined; if (cell.f != null) color = /!/.test(cell.f) ? '#2e7d32' : '#1565c0'; else if (typeof cell.v === 'number') color = '#000000'; else return; s.set(r, c, { ...cell, s: app.book.derivedStyle(cell.s ?? 0, { color }) }); n++; }); }); toast(`${plural(n, 'cell')} colored (blue = formulas, green = links to other sheets, black = numbers)`); },
  },
  { id: 'util.formulasToValuesBook', label: 'Convert All Formulas in the Workbook to Values', icon: 'functions', tab: U, group: G_FORMULAS, edits: true, origin: 'ASAP Utilities', run: async (app, arg) => { const { confirmDialog } = await import('../ui/dialog'); if (arg === undefined && !(await confirmDialog('Formulas to Values', 'Replace every formula on every sheet with its current result? This can be undone.', 'Convert'))) return; let n = 0; app.book.transact('Formulas to Values', () => { for (const s of app.book.sheets) for (const [k, cell] of [...s.cells]) { if (cell.f == null && cell.sp == null) continue; const next: Cell = { ...(cell.s ? { s: cell.s } : {}), ...(cell.v != null ? { v: cell.v } : {}) }; s.set(keyRow(k), keyCol(k), next.v == null && !next.s ? undefined : next); n++; } }); toast(`${plural(n, 'formula')} converted`); } },
  {
    id: 'util.breakExternalLinks', label: 'Break Links to Other Workbooks', icon: 'link_off', tab: U, group: G_FORMULAS, edits: true, tip: 'Replace every formula that refers to another workbook ([Book.xlsx]Sheet!A1) with its last value', origin: 'ASAP Utilities (Remove all external links)',
    run: (app) => { let n = 0; app.book.transact('Break External Links', () => { for (const s of app.book.sheets) for (const [k, cell] of [...s.cells]) { if (cell.f == null || !externalLinkIn(cell.f)) continue; const next: Cell = { ...(cell.s ? { s: cell.s } : {}), ...(cell.v != null ? { v: cell.v } : {}) }; s.set(keyRow(k), keyCol(k), next.v == null && !next.s ? undefined : next); n++; } }); toast(n ? `${plural(n, 'external link')} replaced by values` : 'No links to other workbooks.'); },
  },
  {
    id: 'util.formulasReport', label: 'List All Formulas', icon: 'list_alt', tab: U, group: G_FORMULAS, edits: true, popular: true, tip: 'A sheet with every formula in the workbook (or the selection): sheet, cell, formula and current value', origin: 'ASAP Utilities / Kutools',
    run: (app) => { const rows: Scalar[][] = []; const sheets = singleCell(app) ? app.book.sheets : [app.sheet]; for (const s of sheets) { const cells = singleCell(app) ? [...s.cells].map(([k, cell]) => ({ r: keyRow(k), c: keyCol(k), cell })) : selCells(app, { existingOnly: true }); for (const x of cells) { const cell = x.cell!; if (cell.f == null) continue; rows.push([s.name, cellName(x.r, x.c), '=' + cell.f, cell.v instanceof CellError ? cell.v.code : cell.v ?? null, (cell.f.match(/[A-Z]+\(/g) || []).map((m) => m.slice(0, -1)).join(', ')]); } } if (!rows.length) { toast('No formulas found.'); return; } report(app, 'Formulas', ['Sheet', 'Cell', 'Formula', 'Value', 'Functions'], rows, { colW: [100, 60, 360, 120, 160] }); },
  },
  { id: 'util.menuFormulas', label: 'Formulas', icon: 'function', tab: U, group: G_FORMULAS, kind: 'menu', noRecord: true, run: () => undefined, menu: () => menuOf(['util.refStyle', 'util.copyExact', 'util.pasteExact', 'util.shiftRefs', null, 'tools.formulasToValues', 'util.formulasToValuesBook', 'util.formulasToText', 'util.textToFormulas', 'util.wrapFormulas', 'util.formulaTextRight', null, 'util.namesToRefs', 'util.refsToNames', 'util.breakExternalLinks', null, 'util.colorByContentType', 'util.formulasReport', 'formulas.errorChecking', 'formulas.selectInconsistent', 'util.countSumByColor']) },
]);

void textOf; void selectHits; void cellMatches; void numberToWords; void movedCell; void copyText; void toCsvField; void colName; void sheetProtected; void selRanges;
