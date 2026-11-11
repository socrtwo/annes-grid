/**
 * Conditional formatting (Home ▸ Styles ▸ Conditional Formatting): quick highlight and top/bottom rules,
 * data bars, color scales, icon sets, the New Formatting Rule dialog, the Rules Manager and clearing rules.
 * Rules are evaluated by src/grid/cf.ts.
 */
import { def, getCommand, type MenuEntry } from './registry';
import type { App } from '../app';
import type { Sheet } from '../model/workbook';
import { newId } from '../model/workbook';
import type { CFRule, CellStyle, Border } from '../model/types';
import { ICON_SETS } from '../grid/cf';
import { openDialog, toast, showMenu, confirmDialog, type Field, type DialogControl } from '../ui/dialog';
import { h, icon, clear } from '../ui/dom';
import { safeCssValue } from '../util/sanitize';

/** A colour from a rule, which may come from an untrusted file, safe to put in inline CSS. */
const safeColor = (c: unknown, fallback = '#000000'): string => safeCssValue(c, 60) ?? fallback;
import { selRanges, activeCell, clipToUsed } from './ops';
import { rangeName, parseRange, intersects, splitSheetRef, MAX_ROWS, MAX_COLS, type Range } from '../model/address';

const TAB = 'Home';
const GROUP = 'Styles';

// ============================================================================================
// Presets

export const FORMAT_PRESETS: { id: string; label: string; style: Partial<CellStyle> }[] = [
  { id: 'lightRedDarkRed', label: 'Light Red Fill with Dark Red Text', style: { fill: '#ffc7ce', color: '#9c0006' } },
  { id: 'yellowDarkYellow', label: 'Yellow Fill with Dark Yellow Text', style: { fill: '#ffeb9c', color: '#9c5700' } },
  { id: 'greenDarkGreen', label: 'Green Fill with Dark Green Text', style: { fill: '#c6efce', color: '#006100' } },
  { id: 'lightRed', label: 'Light Red Fill', style: { fill: '#ffc7ce' } },
  { id: 'redText', label: 'Red Text', style: { color: '#9c0006' } },
  { id: 'redBorder', label: 'Red Border', style: { bt: { style: 'thin', color: '#c00000' }, bb: { style: 'thin', color: '#c00000' }, bl: { style: 'thin', color: '#c00000' }, br: { style: 'thin', color: '#c00000' } } },
];

export const BAR_COLORS: { name: string; color: string }[] = [
  { name: 'Blue', color: '#638ec6' }, { name: 'Green', color: '#63c384' }, { name: 'Red', color: '#ff555a' },
  { name: 'Orange', color: '#ffb628' }, { name: 'Light Blue', color: '#008aef' }, { name: 'Purple', color: '#d6007b' },
];

/** Color scale presets, colors listed from the lowest value to the highest. */
export const COLOR_SCALES: { label: string; colors: string[] }[] = [
  { label: 'Green - Yellow - Red', colors: ['#f8696b', '#ffeb84', '#63be7b'] },
  { label: 'Red - Yellow - Green', colors: ['#63be7b', '#ffeb84', '#f8696b'] },
  { label: 'Green - White - Red', colors: ['#f8696b', '#fcfcff', '#63be7b'] },
  { label: 'Red - White - Green', colors: ['#63be7b', '#fcfcff', '#f8696b'] },
  { label: 'Blue - White - Red', colors: ['#f8696b', '#fcfcff', '#5a8ac6'] },
  { label: 'Red - White - Blue', colors: ['#5a8ac6', '#fcfcff', '#f8696b'] },
  { label: 'White - Red', colors: ['#f8696b', '#fcfcff'] },
  { label: 'Red - White', colors: ['#fcfcff', '#f8696b'] },
  { label: 'Green - White', colors: ['#fcfcff', '#63be7b'] },
  { label: 'White - Green', colors: ['#63be7b', '#fcfcff'] },
  { label: 'Green - Yellow', colors: ['#ffef9c', '#63be7b'] },
  { label: 'Yellow - Green', colors: ['#63be7b', '#ffef9c'] },
];

export const ICON_SET_LABELS: Record<string, { label: string; group: string }> = {
  '3Arrows': { label: '3 Arrows (Colored)', group: 'Directional' },
  '3ArrowsGray': { label: '3 Arrows (Gray)', group: 'Directional' },
  '3Triangles': { label: '3 Triangles', group: 'Directional' },
  '4Arrows': { label: '4 Arrows (Colored)', group: 'Directional' },
  '5Arrows': { label: '5 Arrows (Colored)', group: 'Directional' },
  '3TrafficLights1': { label: '3 Traffic Lights', group: 'Shapes' },
  '3Signs': { label: '3 Signs', group: 'Shapes' },
  '4RedToBlack': { label: 'Red To Black', group: 'Shapes' },
  '4TrafficLights': { label: '4 Traffic Lights', group: 'Shapes' },
  '3Symbols': { label: '3 Symbols (Circled)', group: 'Indicators' },
  '3Symbols2': { label: '3 Symbols (Uncircled)', group: 'Indicators' },
  '3Flags': { label: '3 Flags', group: 'Indicators' },
  '3Stars': { label: '3 Stars', group: 'Ratings' },
  '4Rating': { label: '4 Ratings', group: 'Ratings' },
  '5Quarters': { label: '5 Quarters', group: 'Ratings' },
  '5Rating': { label: '5 Ratings', group: 'Ratings' },
  '5Boxes': { label: '5 Boxes', group: 'Ratings' },
};

export const TIME_PERIODS: { value: NonNullable<CFRule['timePeriod']>; label: string }[] = [
  { value: 'yesterday', label: 'Yesterday' }, { value: 'today', label: 'Today' }, { value: 'tomorrow', label: 'Tomorrow' },
  { value: 'last7Days', label: 'In the last 7 days' }, { value: 'lastWeek', label: 'Last week' }, { value: 'thisWeek', label: 'This week' },
  { value: 'nextWeek', label: 'Next week' }, { value: 'lastMonth', label: 'Last month' }, { value: 'thisMonth', label: 'This month' }, { value: 'nextMonth', label: 'Next month' },
];

const OPERATORS: { value: NonNullable<CFRule['operator']>; label: string; sym: string }[] = [
  { value: 'between', label: 'between', sym: 'between' }, { value: 'notBetween', label: 'not between', sym: 'not between' },
  { value: 'equal', label: 'equal to', sym: '=' }, { value: 'notEqual', label: 'not equal to', sym: '<>' },
  { value: 'greaterThan', label: 'greater than', sym: '>' }, { value: 'lessThan', label: 'less than', sym: '<' },
  { value: 'greaterThanOrEqual', label: 'greater than or equal to', sym: '>=' }, { value: 'lessThanOrEqual', label: 'less than or equal to', sym: '<=' },
];

// ============================================================================================
// Pure rule helpers (exported for tests)

/** Converts typed input to a rule formula: "=A1" stays a formula, numbers stay numbers, other text is quoted. */
export function valueToFormula(text: string): string {
  const t = String(text ?? '').trim();
  if (t.startsWith('=')) return t.slice(1);
  if (t === '') return '""';
  if (/^[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?$/.test(t)) return String(Number(t));
  if (/^(TRUE|FALSE)$/i.test(t)) return t.toUpperCase();
  return `"${t.replace(/^"|"$/g, '').replace(/"/g, '""')}"`;
}

/** Inverse of valueToFormula for showing a rule's value in an input box. */
export function formulaToValue(f: string | undefined): string {
  if (f == null) return '';
  if (/^-?\d+(\.\d+)?([eE][-+]?\d+)?$/.test(f)) return f;
  const m = /^"((?:[^"]|"")*)"$/.exec(f);
  if (m) return m[1].replace(/""/g, '"');
  return '=' + f;
}

/** Renumbers priorities 1..n in the current order (sorted by priority). */
export function normalizePriorities(rules: CFRule[]): CFRule[] {
  const sorted = [...rules].sort((a, b) => (a.priority ?? 0) - (b.priority ?? 0));
  sorted.forEach((r, i) => { r.priority = i + 1; });
  return sorted;
}

/** Adds a rule at the top priority (new rules take precedence, as in spreadsheets). Mutates sheet.cf. */
export function insertRule(sheet: Sheet, rule: Omit<CFRule, 'id'> & { id?: string }): CFRule {
  const r: CFRule = { ...rule, id: rule.id || newId('cf') } as CFRule;
  for (const x of sheet.cf) x.priority = (x.priority ?? 0) + 1;
  r.priority = 0;
  sheet.cf = normalizePriorities([...sheet.cf, r]);
  return r;
}

/** Removes `cut` from `a`, returning up to four rectangles. */
export function subtractRange(a: Range, cut: Range): Range[] {
  if (!intersects(a, cut)) return [a];
  const out: Range[] = [];
  if (a.r1 < cut.r1) out.push({ r1: a.r1, c1: a.c1, r2: cut.r1 - 1, c2: a.c2 });
  if (a.r2 > cut.r2) out.push({ r1: cut.r2 + 1, c1: a.c1, r2: a.r2, c2: a.c2 });
  const r1 = Math.max(a.r1, cut.r1), r2 = Math.min(a.r2, cut.r2);
  if (a.c1 < cut.c1) out.push({ r1, c1: a.c1, r2, c2: cut.c1 - 1 });
  if (a.c2 > cut.c2) out.push({ r1, c1: cut.c2 + 1, r2, c2: a.c2 });
  return out;
}

/** Removes rule coverage from the given ranges; rules left with no cells are deleted. Returns the count removed. */
export function clearRulesIn(sheet: Sheet, cuts: Range[]): number {
  const before = sheet.cf.length;
  const next: CFRule[] = [];
  for (const rule of sheet.cf) {
    let ranges = rule.ranges;
    for (const cut of cuts) ranges = ranges.flatMap((x) => subtractRange(x, cut));
    if (ranges.length) next.push(ranges === rule.ranges ? rule : { ...rule, ranges });
  }
  sheet.cf = normalizePriorities(next);
  return before - sheet.cf.length;
}

/** Rules that touch any of the ranges (or all rules when ranges is null), in priority order. */
export function rulesFor(sheet: Sheet, ranges: Range[] | null): CFRule[] {
  const list = [...sheet.cf].sort((a, b) => (a.priority ?? 0) - (b.priority ?? 0));
  return ranges ? list.filter((r) => r.ranges.some((x) => ranges.some((y) => intersects(x, y)))) : list;
}

export function rangesText(ranges: Range[]): string {
  return ranges.map((r) => '=' + rangeName(r).replace(/([A-Z]+)(\d+)/g, '$$$1$$$2')).join(',').replace(/,=/g, ',');
}

export function parseRangesText(text: string): Range[] | null {
  const parts = text.replace(/^=/, '').split(/[,;]/).map((p) => p.trim()).filter(Boolean);
  if (!parts.length) return null;
  const out: Range[] = [];
  for (const p of parts) {
    const rg = parseRange(splitSheetRef(p.replace(/^=/, '')).ref);
    if (!rg) return null;
    out.push(rg);
  }
  return out;
}

export function describeRule(r: CFRule): string {
  const v = (i: number) => formulaToValue(r.formulas?.[i]);
  switch (r.type) {
    case 'cellIs': {
      const op = OPERATORS.find((o) => o.value === r.operator);
      return r.operator === 'between' || r.operator === 'notBetween' ? `Cell Value ${op?.label} ${v(0)} and ${v(1)}` : `Cell Value ${op?.sym ?? '='} ${v(0)}`;
    }
    case 'expression': return `Formula: =${r.formulas?.[0] ?? ''}`;
    case 'colorScale': return (r.colors?.length ?? 3) === 2 ? '2-Color Scale' : '3-Color Scale';
    case 'dataBar': return 'Data Bar';
    case 'iconSet': return `Icon Set: ${ICON_SET_LABELS[r.iconSet || '3TrafficLights1']?.label || r.iconSet}`;
    case 'top10': return `${r.bottom ? 'Bottom' : 'Top'} ${r.rank ?? 10}${r.percent ? '%' : ''}`;
    case 'aboveAverage': return `${r.above === false ? 'Below' : 'Above'} Average${r.equalAverage ? ' or Equal' : ''}${r.stdDev ? ` (${r.stdDev} std dev)` : ''}`;
    case 'duplicateValues': return 'Duplicate Values';
    case 'uniqueValues': return 'Unique Values';
    case 'containsText': return `Cell Value contains "${r.text ?? ''}"`;
    case 'notContainsText': return `Cell Value does not contain "${r.text ?? ''}"`;
    case 'beginsWith': return `Cell Value begins with "${r.text ?? ''}"`;
    case 'endsWith': return `Cell Value ends with "${r.text ?? ''}"`;
    case 'containsBlanks': return 'Blanks';
    case 'notContainsBlanks': return 'No Blanks';
    case 'containsErrors': return 'Errors';
    case 'notContainsErrors': return 'No Errors';
    case 'timePeriod': return `Dates Occurring: ${TIME_PERIODS.find((t) => t.value === r.timePeriod)?.label ?? r.timePeriod}`;
  }
  return r.type;
}

/** Inline CSS that previews a rule's format. */
export function previewCss(r: Partial<CFRule>): string {
  if (r.type === 'colorScale') return `background:linear-gradient(90deg,${(r.colors || ['#f8696b', '#ffeb84', '#63be7b']).map((c) => safeColor(c, 'transparent')).join(',')})`;
  if (r.type === 'dataBar') { const bar = safeColor(r.barColor || '#638ec6', '#638ec6'); return r.gradientBar === false ? `background:linear-gradient(90deg,${bar} 70%,transparent 70%)` : `background:linear-gradient(90deg,${bar},#ffffff 80%)`; }
  return styleCss(r.style || {});
}

export function styleCss(st: Partial<CellStyle>): string {
  const parts: string[] = [];
  if (st.fill) parts.push(`background:${safeColor(st.fill, 'transparent')}`);
  if (st.color) parts.push(`color:${safeColor(st.color)}`);
  if (st.bold) parts.push('font-weight:700');
  if (st.italic) parts.push('font-style:italic');
  const deco = [st.underline ? 'underline' : '', st.strike ? 'line-through' : ''].filter(Boolean).join(' ');
  if (deco) parts.push(`text-decoration:${deco}`);
  if (st.bt || st.bb || st.bl || st.br) parts.push(`border:1px solid ${safeColor((st.bt || st.bb || st.bl || st.br)!.color || '#000')}`);
  return parts.join(';');
}

// ============================================================================================
// UI helpers

function canFormat(app: App, sheet: Sheet = app.sheet): boolean {
  if (sheet.protection.on && !sheet.protection.allow.includes('formatCells')) { toast('The sheet is protected. Unprotect it to change conditional formatting.'); return false; }
  return true;
}

function commitRule(app: App, rule: Omit<CFRule, 'id'>, label: string): void {
  const sheet = app.sheet;
  if (!canFormat(app, sheet)) return;
  app.book.transact(label, () => { app.book.recordSheet(sheet); insertRule(sheet, rule); });
}

function targetRanges(app: App): Range[] {
  return selRanges(app).map((r) => ({ ...r }));
}

function selectionNumbers(app: App): number[] {
  const s = app.sheet;
  const out: number[] = [];
  for (const rg0 of selRanges(app)) {
    const rg = clipToUsed(s, rg0);
    s.forEachIn(rg, (_r, _c, cell) => { if (typeof cell.v === 'number') out.push(cell.v); });
    if (out.length > 5000) break;
  }
  return out;
}

const roundNice = (x: number) => (Math.abs(x) >= 100 ? Math.round(x) : Math.round(x * 100) / 100);

function presetOptions(): { value: string; label: string; style?: string }[] {
  return [...FORMAT_PRESETS.map((p) => ({ value: p.id, label: p.label, style: styleCss(p.style) })), { value: 'custom', label: 'Custom Format…' }];
}

/** Small format dialog: font color, bold, italic, underline, strikethrough, fill, border, number format. */
export async function customFormatDialog(init: Partial<CellStyle> = {}): Promise<Partial<CellStyle> | null> {
  const b = init.bt || init.bb || init.bl || init.br;
  const r = await openDialog({
    title: 'Format Cells', preview: 'AaBbCcYyZz 123', previewStyle: 'padding:10px;border:1px solid var(--border);' + styleCss(init),
    fields: [
      { key: 'heading1', type: 'heading', label: 'Font' },
      { key: 'useColor', label: 'Font color', type: 'checkbox', value: !!init.color, half: true },
      { key: 'color', label: 'Color', type: 'color', value: init.color || '#9c0006', half: true },
      { key: 'bold', label: 'Bold', type: 'checkbox', value: !!init.bold, half: true },
      { key: 'italic', label: 'Italic', type: 'checkbox', value: !!init.italic, half: true },
      { key: 'underline', label: 'Underline', type: 'checkbox', value: !!init.underline, half: true },
      { key: 'strike', label: 'Strikethrough', type: 'checkbox', value: !!init.strike, half: true },
      { key: 'heading2', type: 'heading', label: 'Fill' },
      { key: 'useFill', label: 'Fill color', type: 'checkbox', value: !!init.fill, half: true },
      { key: 'fill', label: 'Color', type: 'color', value: init.fill || '#ffc7ce', half: true },
      { key: 'heading3', type: 'heading', label: 'Border' },
      { key: 'border', label: 'Border', type: 'select', options: [{ value: '', label: 'None' }, { value: 'thin', label: 'Outline (thin)' }, { value: 'medium', label: 'Outline (medium)' }, { value: 'dashed', label: 'Outline (dashed)' }], value: b?.style || '', half: true },
      { key: 'borderColor', label: 'Border color', type: 'color', value: b?.color || '#c00000', half: true },
      { key: 'numFmt', label: 'Number format (optional)', type: 'text', value: init.numFmt || '', placeholder: 'e.g. 0.00% or "Late"' },
    ],
    onChange: (v, ctl) => ctl.setPreview(`<div style="${styleCss(buildStyle(v))}">AaBbCcYyZz 123</div>`),
  });
  return r ? buildStyle(r) : null;
}

function buildStyle(v: Record<string, any>): Partial<CellStyle> {
  const st: Partial<CellStyle> = {};
  if (v.useColor) st.color = v.color;
  if (v.bold) st.bold = true;
  if (v.italic) st.italic = true;
  if (v.underline) st.underline = 'single';
  if (v.strike) st.strike = true;
  if (v.useFill) st.fill = v.fill;
  if (v.border) { const bd: Border = { style: v.border, color: v.borderColor }; st.bt = bd; st.bb = bd; st.bl = bd; st.br = bd; }
  if (v.numFmt) st.numFmt = v.numFmt;
  return st;
}

/** Shows the preset select with "Custom Format…" support; resolves the chosen style. */
async function resolveStyle(preset: string, custom: Partial<CellStyle> | null): Promise<Partial<CellStyle> | null> {
  if (preset === 'custom') return custom || (await customFormatDialog());
  return FORMAT_PRESETS.find((p) => p.id === preset)?.style || FORMAT_PRESETS[0].style;
}

// ============================================================================================
// Quick rules (the classic "value + format preset" dialogs)

type QuickKind = 'greaterThan' | 'lessThan' | 'between' | 'equal' | 'text' | 'date' | 'duplicate' | 'top' | 'topPct' | 'bottom' | 'bottomPct' | 'above' | 'below';

interface QuickArg { value?: string | number; value2?: string | number; preset?: string; style?: Partial<CellStyle>; text?: string; period?: CFRule['timePeriod']; unique?: boolean; rank?: number; option?: string }

const QUICK: Record<QuickKind, { title: string; prompt: string }> = {
  greaterThan: { title: 'Greater Than', prompt: 'Format cells that are GREATER THAN:' },
  lessThan: { title: 'Less Than', prompt: 'Format cells that are LESS THAN:' },
  between: { title: 'Between', prompt: 'Format cells that are BETWEEN:' },
  equal: { title: 'Equal To', prompt: 'Format cells that are EQUAL TO:' },
  text: { title: 'Text That Contains', prompt: 'Format cells that contain the text:' },
  date: { title: 'A Date Occurring', prompt: 'Format cells that contain a date occurring:' },
  duplicate: { title: 'Duplicate Values', prompt: 'Format cells that contain:' },
  top: { title: 'Top 10 Items', prompt: 'Format cells that rank in the TOP:' },
  topPct: { title: 'Top 10%', prompt: 'Format cells that rank in the TOP (percent):' },
  bottom: { title: 'Bottom 10 Items', prompt: 'Format cells that rank in the BOTTOM:' },
  bottomPct: { title: 'Bottom 10%', prompt: 'Format cells that rank in the BOTTOM (percent):' },
  above: { title: 'Above Average', prompt: 'Format cells that are ABOVE AVERAGE:' },
  below: { title: 'Below Average', prompt: 'Format cells that are BELOW AVERAGE:' },
};

/** Builds the rule for a quick-rule kind from dialog values (exported for tests). */
export function quickRule(kind: QuickKind, ranges: Range[], v: QuickArg, style: Partial<CellStyle>): Omit<CFRule, 'id'> {
  const base = { ranges, style };
  switch (kind) {
    case 'greaterThan': return { ...base, type: 'cellIs', operator: 'greaterThan', formulas: [valueToFormula(String(v.value ?? ''))] };
    case 'lessThan': return { ...base, type: 'cellIs', operator: 'lessThan', formulas: [valueToFormula(String(v.value ?? ''))] };
    case 'equal': return { ...base, type: 'cellIs', operator: 'equal', formulas: [valueToFormula(String(v.value ?? ''))] };
    case 'between': return { ...base, type: 'cellIs', operator: 'between', formulas: [valueToFormula(String(v.value ?? '')), valueToFormula(String(v.value2 ?? ''))] };
    case 'text': return { ...base, type: 'containsText', text: String(v.text ?? v.value ?? '') };
    case 'date': return { ...base, type: 'timePeriod', timePeriod: v.period || 'today' };
    case 'duplicate': return { ...base, type: v.unique ? 'uniqueValues' : 'duplicateValues' };
    case 'top': case 'topPct': case 'bottom': case 'bottomPct':
      return { ...base, type: 'top10', rank: Math.max(1, Math.round(Number(v.rank ?? v.value ?? 10)) || 10), percent: kind.endsWith('Pct'), bottom: kind.startsWith('bottom') };
    case 'above': case 'below': {
      const above = kind === 'above';
      const opt = v.option || 'avg';
      const sd = /^sd(\d)$/.exec(opt);
      return { ...base, type: 'aboveAverage', above, equalAverage: opt === 'eq', ...(sd ? { stdDev: Number(sd[1]) } : {}) };
    }
  }
}

async function quickDialog(app: App, kind: QuickKind, arg?: QuickArg): Promise<void> {
  if (!canFormat(app)) return;
  const ranges = targetRanges(app);
  if (arg && (arg.style || arg.preset) && (arg.value != null || arg.text != null || arg.period || arg.rank != null || arg.option || arg.unique != null || ['duplicate', 'above', 'below'].includes(kind))) {
    const style = arg.style || FORMAT_PRESETS.find((p) => p.id === arg.preset)?.style || FORMAT_PRESETS[0].style;
    commitRule(app, quickRule(kind, ranges, arg, style), QUICK[kind].title);
    return;
  }
  const nums = selectionNumbers(app);
  const sorted = [...nums].sort((a, b) => a - b);
  const avg = nums.length ? roundNice(nums.reduce((a, b) => a + b, 0) / nums.length) : 0;
  const fields: Field[] = [];
  const presetField: Field = { key: 'preset', label: 'with', type: 'select', options: presetOptions(), value: kind === 'top' || kind === 'topPct' || kind === 'above' ? 'lightRedDarkRed' : 'lightRedDarkRed' };
  switch (kind) {
    case 'greaterThan': case 'lessThan': case 'equal':
      fields.push({ key: 'value', label: QUICK[kind].prompt, type: 'text', value: String(nums.length ? (kind === 'equal' ? sorted[0] : avg) : ''), autofocus: true, help: 'Type a number or text, or a formula or reference starting with =' });
      break;
    case 'between':
      fields.push({ key: 'value', label: QUICK[kind].prompt, type: 'text', value: String(sorted.length ? roundNice(sorted[Math.floor(sorted.length * 0.25)]) : ''), half: true, autofocus: true },
        { key: 'value2', label: 'and', type: 'text', value: String(sorted.length ? roundNice(sorted[Math.floor(sorted.length * 0.75)]) : ''), half: true });
      break;
    case 'text': fields.push({ key: 'text', label: QUICK[kind].prompt, type: 'text', value: '', autofocus: true }); break;
    case 'date': fields.push({ key: 'period', label: QUICK[kind].prompt, type: 'select', options: TIME_PERIODS, value: 'yesterday' }); break;
    case 'duplicate': fields.push({ key: 'unique', label: QUICK[kind].prompt, type: 'select', options: [{ value: 'dup', label: 'Duplicate values' }, { value: 'unique', label: 'Unique values' }], value: 'dup' }); break;
    case 'top': case 'bottom': case 'topPct': case 'bottomPct':
      fields.push({ key: 'rank', label: QUICK[kind].prompt, type: 'number', min: 1, max: kind.endsWith('Pct') ? 100 : 1000, step: 1, value: 10, suffix: kind.endsWith('Pct') ? '%' : 'items' }); break;
    case 'above': case 'below':
      fields.push({ key: 'option', label: QUICK[kind].prompt, type: 'select', options: [{ value: 'avg', label: 'for the selected range' }, { value: 'eq', label: kind === 'above' ? 'equal to or above average' : 'equal to or below average' }, { value: 'sd1', label: `1 std dev ${kind}` }, { value: 'sd2', label: `2 std dev ${kind}` }, { value: 'sd3', label: `3 std dev ${kind}` }], value: 'avg' });
      break;
  }
  fields.push(presetField);
  let custom: Partial<CellStyle> | null = null;
  const r = await openDialog({
    title: QUICK[kind].title, fields, preview: 'AaBbCcYyZz', previewStyle: 'padding:8px;border:1px solid var(--border);' + styleCss(FORMAT_PRESETS[0].style),
    onChange: async (v, ctl, key) => {
      if (key === 'preset' && v.preset === 'custom') { const st = await customFormatDialog(custom || {}); if (st) custom = st; else ctl.set('preset', 'lightRedDarkRed'); }
      const st = v.preset === 'custom' ? custom || {} : FORMAT_PRESETS.find((p) => p.id === ctl.get('preset'))?.style || {};
      ctl.setPreview(`<div style="${styleCss(st)};padding:2px 6px">AaBbCcYyZz</div>`);
    },
    validate: (v) => {
      if ((kind === 'greaterThan' || kind === 'lessThan' || kind === 'equal' || kind === 'between') && String(v.value ?? '').trim() === '') return 'Enter a value.';
      if (kind === 'between' && String(v.value2 ?? '').trim() === '') return 'Enter the second value.';
      if (kind.startsWith('top') || kind.startsWith('bottom')) { const n = Number(v.rank); if (!(n >= 1) || (kind.endsWith('Pct') && n > 100)) return kind.endsWith('Pct') ? 'Enter a percentage between 1 and 100.' : 'Enter a whole number of 1 or more.'; }
      for (const k of ['value', 'value2']) { const t = String(v[k] ?? '').trim(); if (t.startsWith('=') && app.engine.compile(t.slice(1)) instanceof Error) return 'The formula is not valid.'; }
      return null;
    },
  });
  if (!r) return;
  const style = await resolveStyle(r.preset, custom);
  if (!style) return;
  commitRule(app, quickRule(kind, ranges, { ...r, unique: r.unique === 'unique', rank: r.rank }, style), QUICK[kind].title);
}

// ============================================================================================
// Data bars, color scales, icon sets

export function dataBarRule(ranges: Range[], color: string, gradient: boolean): Omit<CFRule, 'id'> {
  return { ranges, type: 'dataBar', barColor: color, gradientBar: gradient, showValue: true, stops: [{ type: 'min' }, { type: 'max' }] };
}

export function colorScaleRule(ranges: Range[], colors: string[]): Omit<CFRule, 'id'> {
  return { ranges, type: 'colorScale', colors: [...colors], stops: colors.length === 2 ? [{ type: 'min' }, { type: 'max' }] : [{ type: 'min' }, { type: 'percentile', value: 50 }, { type: 'max' }] };
}

export function iconSetRule(ranges: Range[], set: string): Omit<CFRule, 'id'> {
  const n = ICON_SETS[set]?.n ?? 3;
  const stops: NonNullable<CFRule['stops']> = [];
  for (let i = 0; i < n; i++) stops.push({ type: 'percent', value: Math.round((100 * i) / n) });
  return { ranges, type: 'iconSet', iconSet: set, stops, showValue: true };
}

// ============================================================================================
// New / Edit Formatting Rule dialog

const RULE_KINDS = [
  { value: 'scale', label: 'Format all cells based on their values' },
  { value: 'contains', label: 'Format only cells that contain' },
  { value: 'rank', label: 'Format only top or bottom ranked values' },
  { value: 'avg', label: 'Format only values that are above or below average' },
  { value: 'dup', label: 'Format only unique or duplicate values' },
  { value: 'formula', label: 'Use a formula to determine which cells to format' },
];

const STOP_TYPES = [
  { value: 'min', label: 'Lowest Value' }, { value: 'max', label: 'Highest Value' }, { value: 'num', label: 'Number' },
  { value: 'percent', label: 'Percent' }, { value: 'percentile', label: 'Percentile' }, { value: 'formula', label: 'Formula' },
];

function kindOf(r?: CFRule): string {
  if (!r) return 'scale';
  switch (r.type) {
    case 'colorScale': case 'dataBar': case 'iconSet': return 'scale';
    case 'top10': return 'rank';
    case 'aboveAverage': return 'avg';
    case 'duplicateValues': case 'uniqueValues': return 'dup';
    case 'expression': return 'formula';
  }
  return 'contains';
}

/** Opens the rule editor. Resolves to the edited/created rule (not yet stored), or null. */
export async function ruleDialog(app: App, init: CFRule | undefined, ranges: Range[], opts: { showAppliesTo?: boolean } = {}): Promise<CFRule | null> {
  let style: Partial<CellStyle> = { ...(init?.style || (init ? {} : FORMAT_PRESETS[0].style)) };
  const r = init;
  const stop = (i: number, d: string) => r?.stops?.[i]?.type || d;
  const stopV = (i: number, d: string | number = '') => (r?.stops?.[i]?.value ?? d) as any;
  const scaleStyle = r?.type === 'dataBar' ? 'dataBar' : r?.type === 'iconSet' ? 'iconSet' : r?.type === 'colorScale' && r.colors?.length === 2 ? '2color' : '3color';
  const containsSub = !r ? 'cellValue' : r.type === 'cellIs' ? 'cellValue' : ['containsText', 'notContainsText', 'beginsWith', 'endsWith'].includes(r.type) ? 'text' : r.type === 'timePeriod' ? 'dates' : r.type === 'containsBlanks' ? 'blanks' : r.type === 'notContainsBlanks' ? 'noBlanks' : r.type === 'containsErrors' ? 'errors' : r.type === 'notContainsErrors' ? 'noErrors' : 'cellValue';
  const iconSet = r?.iconSet || '3TrafficLights1';
  const isScale = (v: Record<string, any>) => v.kind === 'scale';
  const isCS = (v: Record<string, any>) => isScale(v) && (v.scaleStyle === '2color' || v.scaleStyle === '3color');
  const stopTypesFor = (allowMinMax: boolean) => STOP_TYPES.filter((s) => allowMinMax || (s.value !== 'min' && s.value !== 'max'));
  const iconN = (v: Record<string, any>) => ICON_SETS[v.iconSet]?.n ?? 3;
  const iconDefault = (n: number, i: number) => Math.round((100 * i) / n);
  const fields: Field[] = [
    { key: 'kind', label: 'Select a rule type', type: 'radio', options: RULE_KINDS, value: kindOf(r) },
    ...(opts.showAppliesTo ? [{ key: 'appliesTo', label: 'Applies to', type: 'text' as const, value: rangesText(ranges) }] : []),
    { key: 'h1', type: 'heading', label: 'Edit the rule description' },
    // Scales, bars, icons
    { key: 'scaleStyle', label: 'Format style', type: 'select', options: [{ value: '2color', label: '2-Color Scale' }, { value: '3color', label: '3-Color Scale' }, { value: 'dataBar', label: 'Data Bar' }, { value: 'iconSet', label: 'Icon Sets' }], value: scaleStyle, visibleIf: isScale },
    { key: 'minType', label: 'Minimum', type: 'select', options: stopTypesFor(true).filter((s) => s.value !== 'max'), value: stop(0, 'min'), half: true, visibleIf: (v) => isCS(v) || (isScale(v) && v.scaleStyle === 'dataBar') },
    { key: 'minValue', label: 'Value', type: 'text', value: String(stopV(0, '')), half: true, visibleIf: (v) => (isCS(v) || (isScale(v) && v.scaleStyle === 'dataBar')) && v.minType !== 'min' },
    { key: 'minColor', label: 'Minimum color', type: 'color', value: r?.colors?.[0] || '#f8696b', visibleIf: isCS },
    { key: 'midType', label: 'Midpoint', type: 'select', options: stopTypesFor(false), value: r?.colors?.length === 3 ? stop(1, 'percentile') : 'percentile', half: true, visibleIf: (v) => isCS(v) && v.scaleStyle === '3color' },
    { key: 'midValue', label: 'Value', type: 'text', value: String(r?.colors?.length === 3 ? stopV(1, 50) : 50), half: true, visibleIf: (v) => isCS(v) && v.scaleStyle === '3color' },
    { key: 'midColor', label: 'Midpoint color', type: 'color', value: r?.colors?.length === 3 ? r.colors[1] : '#ffeb84', visibleIf: (v) => isCS(v) && v.scaleStyle === '3color' },
    { key: 'maxType', label: 'Maximum', type: 'select', options: stopTypesFor(true).filter((s) => s.value !== 'min'), value: stop(r?.type === 'colorScale' ? (r.colors?.length ?? 3) - 1 : 1, 'max'), half: true, visibleIf: (v) => isCS(v) || (isScale(v) && v.scaleStyle === 'dataBar') },
    { key: 'maxValue', label: 'Value', type: 'text', value: String(stopV(r?.type === 'colorScale' ? (r.colors?.length ?? 3) - 1 : 1, '')), half: true, visibleIf: (v) => (isCS(v) || (isScale(v) && v.scaleStyle === 'dataBar')) && v.maxType !== 'max' },
    { key: 'maxColor', label: 'Maximum color', type: 'color', value: r?.colors?.[(r.colors?.length ?? 1) - 1] || '#63be7b', visibleIf: isCS },
    { key: 'barColor', label: 'Bar color', type: 'color', value: r?.barColor || '#638ec6', half: true, visibleIf: (v) => isScale(v) && v.scaleStyle === 'dataBar' },
    { key: 'barFill', label: 'Fill', type: 'select', options: [{ value: 'gradient', label: 'Gradient Fill' }, { value: 'solid', label: 'Solid Fill' }], value: r?.gradientBar === false ? 'solid' : 'gradient', half: true, visibleIf: (v) => isScale(v) && v.scaleStyle === 'dataBar' },
    { key: 'barOnly', label: 'Show Bar Only', type: 'checkbox', value: r?.type === 'dataBar' && r.showValue === false, visibleIf: (v) => isScale(v) && v.scaleStyle === 'dataBar' },
    { key: 'iconSet', label: 'Icon style', type: 'select', options: Object.keys(ICON_SETS).map((k) => ({ value: k, label: ICON_SET_LABELS[k]?.label || k })), value: iconSet, visibleIf: (v) => isScale(v) && v.scaleStyle === 'iconSet' },
    { key: 'reverse', label: 'Reverse Icon Order', type: 'checkbox', value: !!r?.reverseIcons, half: true, visibleIf: (v) => isScale(v) && v.scaleStyle === 'iconSet' },
    { key: 'iconOnly', label: 'Show Icon Only', type: 'checkbox', value: r?.type === 'iconSet' && r.showValue === false, half: true, visibleIf: (v) => isScale(v) && v.scaleStyle === 'iconSet' },
    ...[1, 2, 3, 4].flatMap((i): Field[] => [
      { key: `it${i}`, label: `Icon ${i + 1} when value is >=`, type: 'select', options: stopTypesFor(false), value: r?.type === 'iconSet' ? stop(i, 'percent') : 'percent', half: true, visibleIf: (v) => isScale(v) && v.scaleStyle === 'iconSet' && i < iconN(v) },
      { key: `iv${i}`, label: 'Value', type: 'text', value: String(r?.type === 'iconSet' ? stopV(i, iconDefault(ICON_SETS[iconSet]?.n ?? 3, i)) : iconDefault(3, i)), half: true, visibleIf: (v) => isScale(v) && v.scaleStyle === 'iconSet' && i < iconN(v) },
    ]),
    // Cells that contain
    { key: 'sub', label: 'Format only cells with', type: 'select', options: [{ value: 'cellValue', label: 'Cell Value' }, { value: 'text', label: 'Specific Text' }, { value: 'dates', label: 'Dates Occurring' }, { value: 'blanks', label: 'Blanks' }, { value: 'noBlanks', label: 'No Blanks' }, { value: 'errors', label: 'Errors' }, { value: 'noErrors', label: 'No Errors' }], value: containsSub, visibleIf: (v) => v.kind === 'contains' },
    { key: 'op', label: 'Operator', type: 'select', options: OPERATORS.map((o) => ({ value: o.value, label: o.label })), value: r?.operator || 'between', visibleIf: (v) => v.kind === 'contains' && v.sub === 'cellValue' },
    { key: 'v1', label: 'Value', type: 'text', value: formulaToValue(r?.type === 'cellIs' ? r.formulas?.[0] : undefined), half: true, visibleIf: (v) => v.kind === 'contains' && v.sub === 'cellValue' },
    { key: 'v2', label: 'and', type: 'text', value: formulaToValue(r?.type === 'cellIs' ? r.formulas?.[1] : undefined), half: true, visibleIf: (v) => v.kind === 'contains' && v.sub === 'cellValue' && (v.op === 'between' || v.op === 'notBetween') },
    { key: 'textOp', label: 'Text', type: 'select', options: [{ value: 'containsText', label: 'containing' }, { value: 'notContainsText', label: 'not containing' }, { value: 'beginsWith', label: 'beginning with' }, { value: 'endsWith', label: 'ending with' }], value: ['containsText', 'notContainsText', 'beginsWith', 'endsWith'].includes(r?.type || '') ? r!.type : 'containsText', half: true, visibleIf: (v) => v.kind === 'contains' && v.sub === 'text' },
    { key: 'text', label: 'Text', type: 'text', value: r?.text || '', half: true, visibleIf: (v) => v.kind === 'contains' && v.sub === 'text' },
    { key: 'period', label: 'Dates', type: 'select', options: TIME_PERIODS, value: r?.timePeriod || 'today', visibleIf: (v) => v.kind === 'contains' && v.sub === 'dates' },
    // Ranked
    { key: 'topBottom', label: 'Format values that rank in the', type: 'select', options: [{ value: 'top', label: 'Top' }, { value: 'bottom', label: 'Bottom' }], value: r?.bottom ? 'bottom' : 'top', half: true, visibleIf: (v) => v.kind === 'rank' },
    { key: 'rank', label: 'Count', type: 'number', min: 1, step: 1, value: r?.rank ?? 10, half: true, visibleIf: (v) => v.kind === 'rank' },
    { key: 'percent', label: '% of the selected range', type: 'checkbox', value: !!r?.percent, visibleIf: (v) => v.kind === 'rank' },
    // Average
    { key: 'avgOpt', label: 'Format values that are', type: 'select', options: [{ value: 'above', label: 'above' }, { value: 'below', label: 'below' }, { value: 'eqAbove', label: 'equal or above' }, { value: 'eqBelow', label: 'equal or below' }, { value: 'sd1a', label: '1 std dev above' }, { value: 'sd1b', label: '1 std dev below' }, { value: 'sd2a', label: '2 std dev above' }, { value: 'sd2b', label: '2 std dev below' }, { value: 'sd3a', label: '3 std dev above' }, { value: 'sd3b', label: '3 std dev below' }], value: r?.type === 'aboveAverage' ? (r.stdDev ? `sd${r.stdDev}${r.above === false ? 'b' : 'a'}` : r.equalAverage ? (r.above === false ? 'eqBelow' : 'eqAbove') : r.above === false ? 'below' : 'above') : 'above', help: 'the average for the selected range', visibleIf: (v) => v.kind === 'avg' },
    // Duplicates
    { key: 'dupOpt', label: 'Format all', type: 'select', options: [{ value: 'dup', label: 'duplicate' }, { value: 'unique', label: 'unique' }], value: r?.type === 'uniqueValues' ? 'unique' : 'dup', help: 'values in the selected range', visibleIf: (v) => v.kind === 'dup' },
    // Formula
    { key: 'formula', label: 'Format values where this formula is true', type: 'text', value: r?.type === 'expression' ? '=' + (r.formulas?.[0] ?? '') : '=', placeholder: '=$B2>100', help: 'Write the formula for the top-left cell of the range; relative references adjust for the other cells.', visibleIf: (v) => v.kind === 'formula' },
    // Format
    { key: 'fmtBtn', label: 'Format…', type: 'button', visibleIf: (v) => !isScale(v), onClick: async (_v, ctl) => { const st = await customFormatDialog(style); if (st) { style = st; updatePreview(ctl); } } },
    { key: 'stopIfTrue', label: 'Stop If True', type: 'checkbox', value: !!r?.stopIfTrue },
  ];
  let ctlRef: DialogControl | null = null;
  const updatePreview = (ctl: DialogControl) => {
    const v = ctl.values();
    const draft = build(v, true);
    ctl.setPreview(`<div style="${previewCss(draft || {})};padding:4px 8px;min-height:22px">${v.kind === 'scale' && v.scaleStyle === 'iconSet' ? '' : 'AaBbCcYyZz'}</div>`);
  };
  const stopFrom = (t: string, val: string): NonNullable<CFRule['stops']>[number] => (t === 'min' || t === 'max' ? { type: t } : { type: t as any, value: t === 'formula' ? String(val).replace(/^=/, '') : Number(val) });
  const build = (v: Record<string, any>, draft = false): CFRule | null => {
    const appliesTo = opts.showAppliesTo ? parseRangesText(String(v.appliesTo || '')) : ranges;
    if (!appliesTo) return null;
    const base: CFRule = { id: init?.id || newId('cf'), ranges: appliesTo, type: 'cellIs', priority: init?.priority, stopIfTrue: !!v.stopIfTrue || undefined };
    switch (v.kind) {
      case 'scale':
        if (v.scaleStyle === 'dataBar') return { ...base, type: 'dataBar', barColor: v.barColor, gradientBar: v.barFill !== 'solid', showValue: !v.barOnly, stops: [stopFrom(v.minType, v.minValue), stopFrom(v.maxType, v.maxValue)] };
        if (v.scaleStyle === 'iconSet') {
          const n = ICON_SETS[v.iconSet]?.n ?? 3;
          const stops: NonNullable<CFRule['stops']> = [{ type: 'percent', value: 0 }];
          for (let i = 1; i < n; i++) stops.push(stopFrom(v[`it${i}`], v[`iv${i}`]));
          return { ...base, type: 'iconSet', iconSet: v.iconSet, reverseIcons: !!v.reverse || undefined, showValue: !v.iconOnly, stops };
        }
        if (v.scaleStyle === '2color') return { ...base, type: 'colorScale', colors: [v.minColor, v.maxColor], stops: [stopFrom(v.minType, v.minValue), stopFrom(v.maxType, v.maxValue)] };
        return { ...base, type: 'colorScale', colors: [v.minColor, v.midColor, v.maxColor], stops: [stopFrom(v.minType, v.minValue), stopFrom(v.midType, v.midValue), stopFrom(v.maxType, v.maxValue)] };
      case 'contains': {
        const st = { style };
        switch (v.sub) {
          case 'cellValue': return { ...base, ...st, type: 'cellIs', operator: v.op, formulas: v.op === 'between' || v.op === 'notBetween' ? [valueToFormula(v.v1), valueToFormula(v.v2)] : [valueToFormula(v.v1)] };
          case 'text': return { ...base, ...st, type: v.textOp, text: String(v.text || '') };
          case 'dates': return { ...base, ...st, type: 'timePeriod', timePeriod: v.period };
          case 'blanks': return { ...base, ...st, type: 'containsBlanks' };
          case 'noBlanks': return { ...base, ...st, type: 'notContainsBlanks' };
          case 'errors': return { ...base, ...st, type: 'containsErrors' };
          case 'noErrors': return { ...base, ...st, type: 'notContainsErrors' };
        }
        return null;
      }
      case 'rank': return { ...base, style, type: 'top10', bottom: v.topBottom === 'bottom', rank: Math.max(1, Math.round(Number(v.rank) || 10)), percent: !!v.percent };
      case 'avg': {
        const o = String(v.avgOpt);
        const sd = /^sd(\d)([ab])$/.exec(o);
        const above = sd ? sd[2] === 'a' : o === 'above' || o === 'eqAbove';
        return { ...base, style, type: 'aboveAverage', above, equalAverage: o.startsWith('eq') || undefined, ...(sd ? { stdDev: Number(sd[1]) } : {}) };
      }
      case 'dup': return { ...base, style, type: v.dupOpt === 'unique' ? 'uniqueValues' : 'duplicateValues' };
      case 'formula': {
        const f = String(v.formula || '').trim().replace(/^=/, '');
        if (!f && !draft) return null;
        return { ...base, style, type: 'expression', formulas: [f] };
      }
    }
    return null;
  };
  const res = await openDialog({
    title: init ? 'Edit Formatting Rule' : 'New Formatting Rule', fields, wide: true, preview: true, previewStyle: 'border:1px solid var(--border);padding:4px',
    onOpen: (ctl) => { ctlRef = ctl; },
    onChange: (_v, ctl) => updatePreview(ctl),
    validate: (v) => {
      if (opts.showAppliesTo && !parseRangesText(String(v.appliesTo || ''))) return 'Type a valid range for "Applies to" (for example =$A$1:$A$20).';
      if (v.kind === 'formula') {
        const f = String(v.formula || '').trim().replace(/^=/, '');
        if (!f) return 'Type a formula.';
        if (app.engine.compile(f) instanceof Error) return 'The formula is not valid.';
      }
      if (v.kind === 'contains' && v.sub === 'cellValue') {
        if (!String(v.v1 ?? '').trim()) return 'Enter a value.';
        if ((v.op === 'between' || v.op === 'notBetween') && !String(v.v2 ?? '').trim()) return 'Enter the second value.';
      }
      if (v.kind === 'scale' && v.scaleStyle !== 'iconSet') {
        for (const [t, val] of [[v.minType, v.minValue], [v.midType, v.midValue], [v.maxType, v.maxValue]]) {
          if (t && t !== 'min' && t !== 'max' && t !== 'formula' && isNaN(Number(val))) return 'Enter a number for each Number, Percent or Percentile value.';
        }
      }
      return null;
    },
  });
  void ctlRef;
  if (!res) return null;
  return build(res);
}

// ============================================================================================
// Rules Manager

async function manageRules(app: App): Promise<void> {
  const book = app.book;
  let sheet = app.sheet;
  const selection = targetRanges(app);
  let scope: string = 'selection';
  let work: CFRule[] = JSON.parse(JSON.stringify(normalizePriorities(sheet.cf.map((r) => ({ ...r })))));
  let selectedId: string | null = null;
  const dirty = new Set<string>();
  const content = h('div', { class: 'ql-cfmgr' });
  const list = h('div', { class: 'ql-cfmgr-list' });
  const scopeSel = h('select', { 'aria-label': 'Show formatting rules for' }, h('option', { value: 'selection' }, 'Current Selection'), ...book.sheets.map((s) => h('option', { value: 'sheet:' + s.id }, s === sheet ? 'This Worksheet' : `Sheet: ${s.name}`))) as HTMLSelectElement;
  const visible = () => {
    const sorted = [...work].sort((a, b) => (a.priority ?? 0) - (b.priority ?? 0));
    return scope === 'selection' ? sorted.filter((r) => r.ranges.some((x) => selection.some((y) => intersects(x, y)))) : sorted;
  };
  const render = () => {
    clear(list);
    const rows = visible();
    const tbl = h('table', { class: 'ql-cfmgr-table' }, h('thead', null, h('tr', null, h('th', null, 'Rule (applied in order shown)'), h('th', null, 'Format'), h('th', null, 'Applies to'), h('th', null, 'Stop If True'))));
    const tb = h('tbody');
    for (const r of rows) {
      const applies = h('input', { class: 'ql-input', value: rangesText(r.ranges), 'aria-label': 'Applies to', spellcheck: 'false' }) as HTMLInputElement;
      applies.addEventListener('change', () => { const rg = parseRangesText(applies.value); if (!rg) { toast('That is not a valid range.'); applies.value = rangesText(r.ranges); return; } r.ranges = rg; dirty.add(r.id); });
      applies.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); e.stopPropagation(); applies.blur(); } });
      const stopCb = h('input', { type: 'checkbox', checked: !!r.stopIfTrue, 'aria-label': 'Stop If True', disabled: r.type === 'colorScale' || r.type === 'dataBar' || r.type === 'iconSet' }) as HTMLInputElement;
      stopCb.addEventListener('change', () => { r.stopIfTrue = stopCb.checked || undefined; dirty.add(r.id); });
      const tr = h('tr', { class: r.id === selectedId ? 'ql-on' : '', tabindex: '0' },
        h('td', null, describeRule(r)),
        h('td', null, h('span', { class: 'ql-cf-swatch', style: previewCss(r) }, r.type === 'iconSet' ? '' : 'AaBbCcYyZz')),
        h('td', null, applies),
        h('td', { class: 'ql-center' }, stopCb));
      tr.addEventListener('click', (e) => { if ((e.target as HTMLElement).tagName === 'INPUT') { selectedId = r.id; return; } selectedId = r.id; render(); });
      tr.addEventListener('dblclick', () => { selectedId = r.id; void editSel(); });
      tb.appendChild(tr);
    }
    if (!rows.length) tb.appendChild(h('tr', null, h('td', { colspan: '4', class: 'ql-help' }, scope === 'selection' ? 'No rules apply to the selected cells.' : 'This sheet has no conditional formatting rules.')));
    tbl.appendChild(tb);
    list.appendChild(tbl);
  };
  const editSel = async () => {
    const r = work.find((x) => x.id === selectedId);
    if (!r) { toast('Select a rule first.'); return; }
    const nr = await ruleDialog(app, r, r.ranges, { showAppliesTo: true });
    if (!nr) return;
    const i = work.indexOf(r);
    work[i] = { ...nr, priority: r.priority };
    dirty.add(r.id);
    render();
  };
  const move = (dir: -1 | 1) => {
    const rows = visible();
    const i = rows.findIndex((x) => x.id === selectedId);
    const j = i + dir;
    if (i < 0 || j < 0 || j >= rows.length) return;
    const a = rows[i], b = rows[j];
    const pa = a.priority ?? 0;
    a.priority = b.priority; b.priority = pa;
    work = normalizePriorities(work);
    render();
  };
  scopeSel.addEventListener('change', () => {
    const v = scopeSel.value;
    if (v === 'selection') { scope = 'selection'; }
    else {
      const sh = book.sheets.find((s) => 'sheet:' + s.id === v);
      if (sh && sh !== sheet) {
        if (dirty.size || JSON.stringify(work) !== JSON.stringify(normalizePriorities(sheet.cf.map((r) => ({ ...r }))))) apply();
        sheet = sh;
        work = JSON.parse(JSON.stringify(normalizePriorities(sh.cf.map((r) => ({ ...r })))));
      }
      scope = 'sheet';
    }
    selectedId = null;
    render();
  });
  const apply = () => {
    if (app.isReadOnly()) { toast('This workbook is read-only.'); return; }
    if (!canFormat(app, sheet)) return;
    const target = sheet;
    const next = normalizePriorities(work.map((r) => ({ ...r })));
    if (JSON.stringify(next) === JSON.stringify(normalizePriorities(target.cf.map((r) => ({ ...r }))))) return;
    book.transact('Manage Rules', () => { book.recordSheet(target); target.cf = next; });
    dirty.clear();
  };
  const bar = h('div', { class: 'ql-row ql-wrap' },
    h('button', { class: 'ql-btn', type: 'button', onclick: async () => { const nr = await ruleDialog(app, undefined, scope === 'selection' ? selection : selection, { showAppliesTo: true }); if (!nr) return; for (const x of work) x.priority = (x.priority ?? 0) + 1; nr.priority = 0; work = normalizePriorities([...work, nr]); selectedId = nr.id; render(); } }, icon('add'), ' New Rule…'),
    h('button', { class: 'ql-btn', type: 'button', onclick: () => void editSel() }, icon('edit'), ' Edit Rule…'),
    h('button', { class: 'ql-btn', type: 'button', onclick: () => { const i = work.findIndex((x) => x.id === selectedId); if (i < 0) { toast('Select a rule first.'); return; } work.splice(i, 1); work = normalizePriorities(work); selectedId = null; render(); } }, icon('delete'), ' Delete Rule'),
    h('button', { class: 'ql-btn', type: 'button', onclick: () => { const r = work.find((x) => x.id === selectedId); if (!r) { toast('Select a rule first.'); return; } const copy: CFRule = { ...JSON.parse(JSON.stringify(r)), id: newId('cf'), priority: (r.priority ?? 0) + 0.5 }; work = normalizePriorities([...work, copy]); selectedId = copy.id; render(); } }, icon('content_copy'), ' Duplicate Rule'),
    h('button', { class: 'ql-iconbtn', type: 'button', title: 'Move Up', 'aria-label': 'Move rule up', onclick: () => move(-1) }, icon('arrow_upward')),
    h('button', { class: 'ql-iconbtn', type: 'button', title: 'Move Down', 'aria-label': 'Move rule down', onclick: () => move(1) }, icon('arrow_downward')));
  if (!visible().length && work.length) { scope = 'sheet'; scopeSel.value = 'sheet:' + sheet.id; }
  content.append(h('div', { class: 'ql-row' }, h('label', { class: 'ql-label' }, 'Show formatting rules for: '), scopeSel), bar, list);
  render();
  const r = await openDialog({ title: 'Conditional Formatting Rules Manager', content, wide: true, buttons: [{ label: 'Apply', onClick: () => { apply(); render(); } }] });
  if (r) apply();
}

// ============================================================================================
// Clearing rules

function clearSelection(app: App): void {
  const sheet = app.sheet;
  if (!canFormat(app, sheet)) return;
  const cuts = targetRanges(app);
  if (!rulesFor(sheet, cuts).length) { toast('The selected cells have no conditional formatting.'); return; }
  app.book.transact('Clear Rules', () => { app.book.recordSheet(sheet); clearRulesIn(sheet, cuts); });
}

async function clearSheet(app: App): Promise<void> {
  const sheet = app.sheet;
  if (!canFormat(app, sheet)) return;
  if (!sheet.cf.length) { toast('This sheet has no conditional formatting.'); return; }
  app.book.transact('Clear Rules', () => { app.book.recordSheet(sheet); sheet.cf = []; });
}

function clearTable(app: App): void {
  const sheet = app.sheet;
  const { r, c } = activeCell(app);
  const t = sheet.tables.find((x) => r >= x.range.r1 && r <= x.range.r2 && c >= x.range.c1 && c <= x.range.c2);
  if (!t) { toast('Select a cell in a table first.'); return; }
  if (!canFormat(app, sheet)) return;
  app.book.transact('Clear Rules', () => { app.book.recordSheet(sheet); clearRulesIn(sheet, [t.range]); });
}

function clearPivot(app: App): void {
  const sheet = app.sheet;
  const { r, c } = activeCell(app);
  const p = sheet.pivots.find((x) => x.lastRange && r >= x.lastRange.r1 && r <= x.lastRange.r2 && c >= x.lastRange.c1 && c <= x.lastRange.c2);
  if (!p?.lastRange) { toast('Select a cell in a PivotTable first.'); return; }
  if (!canFormat(app, sheet)) return;
  app.book.transact('Clear Rules', () => { app.book.recordSheet(sheet); clearRulesIn(sheet, [p.lastRange!]); });
}

// ============================================================================================
// Menu

function condMenu(app: App): MenuEntry[] {
  const inTable = (() => { const s = app.sheet; const { r, c } = s.sel; return s.tables.some((t) => r >= t.range.r1 && r <= t.range.r2 && c >= t.range.c1 && c <= t.range.c2); })();
  const inPivot = (() => { const s = app.sheet; const { r, c } = s.sel; return s.pivots.some((p) => p.lastRange && r >= p.lastRange.r1 && r <= p.lastRange.r2 && c >= p.lastRange.c1 && c <= p.lastRange.c2); })();
  const iconGroups = ['Directional', 'Shapes', 'Indicators', 'Ratings'];
  return [
    {
      label: 'Highlight Cells Rules', icon: 'border_color', submenu: [
        { id: 'cond.greaterThan', label: 'Greater Than…', icon: 'chevron_right' },
        { id: 'cond.lessThan', label: 'Less Than…', icon: 'chevron_left' },
        { id: 'cond.between', label: 'Between…', icon: 'unfold_less' },
        { id: 'cond.equalTo', label: 'Equal To…', icon: 'drag_handle' },
        { id: 'cond.textContains', label: 'Text that Contains…', icon: 'match_case' },
        { id: 'cond.dateOccurring', label: 'A Date Occurring…', icon: 'event' },
        { id: 'cond.duplicateValues', label: 'Duplicate Values…', icon: 'content_copy' },
        { separator: true },
        { id: 'cond.newRule', label: 'More Rules…', icon: 'more_horiz' },
      ],
    },
    {
      label: 'Top/Bottom Rules', icon: 'format_list_numbered', submenu: [
        { id: 'cond.top10', label: 'Top 10 Items…', icon: 'vertical_align_top' },
        { id: 'cond.top10Percent', label: 'Top 10%…', icon: 'percent' },
        { id: 'cond.bottom10', label: 'Bottom 10 Items…', icon: 'vertical_align_bottom' },
        { id: 'cond.bottom10Percent', label: 'Bottom 10%…', icon: 'percent' },
        { id: 'cond.aboveAverage', label: 'Above Average…', icon: 'trending_up' },
        { id: 'cond.belowAverage', label: 'Below Average…', icon: 'trending_down' },
        { separator: true },
        { id: 'cond.newRule', label: 'More Rules…', icon: 'more_horiz' },
      ],
    },
    {
      label: 'Data Bars', icon: 'data_thresholding', submenu: [
        { heading: 'Gradient Fill' },
        ...BAR_COLORS.map((b) => ({ id: 'cond.dataBar', arg: { color: b.color, gradient: true }, label: `${b.name} Data Bar`, swatch: `linear-gradient(90deg,${b.color},#ffffff)` })),
        { heading: 'Solid Fill' },
        ...BAR_COLORS.map((b) => ({ id: 'cond.dataBar', arg: { color: b.color, gradient: false }, label: `Solid ${b.name} Data Bar`, swatch: b.color })),
        { separator: true },
        { id: 'cond.newRule', label: 'More Rules…', icon: 'more_horiz' },
      ],
    },
    {
      label: 'Color Scales', icon: 'gradient', submenu: [
        ...COLOR_SCALES.map((cs, i) => ({ id: 'cond.colorScale', arg: i, label: `${cs.label} Color Scale`, swatch: `linear-gradient(90deg,${[...cs.colors].reverse().join(',')})` })),
        { separator: true },
        { id: 'cond.newRule', label: 'More Rules…', icon: 'more_horiz' },
      ],
    },
    {
      label: 'Icon Sets', icon: 'traffic', submenu: [
        ...iconGroups.flatMap((g) => [{ heading: g } as MenuEntry, ...Object.keys(ICON_SETS).filter((k) => ICON_SET_LABELS[k]?.group === g).map((k) => ({ id: 'cond.iconSet', arg: k, label: ICON_SET_LABELS[k].label, icon: ICON_SETS[k].icons[ICON_SETS[k].n - 1] }))]),
        { separator: true },
        { id: 'cond.newRule', label: 'More Rules…', icon: 'more_horiz' },
      ],
    },
    { separator: true },
    { id: 'cond.newRule', label: 'New Rule…', icon: 'add_box' },
    {
      label: 'Clear Rules', icon: 'format_color_reset', submenu: [
        { id: 'cond.clearSelection', label: 'Clear Rules from Selected Cells' },
        { id: 'cond.clearSheet', label: 'Clear Rules from Entire Sheet' },
        ...(inTable ? [{ id: 'cond.clearTable', label: 'Clear Rules from This Table' }] : []),
        ...(inPivot ? [{ id: 'cond.clearPivot', label: 'Clear Rules from This PivotTable' }] : []),
      ],
    },
    { id: 'cond.manage', label: 'Manage Rules…', icon: 'tune' },
  ];
}

const quickCmd = (id: string, label: string, ic: string, kind: QuickKind, tip: string): any => ({
  id, label, icon: ic, tab: TAB, group: GROUP, edits: true, tip, keywords: 'conditional formatting highlight rule',
  run: (app: App, arg?: QuickArg) => quickDialog(app, kind, arg),
});

def([
  {
    id: 'cond.menu', label: 'Conditional Formatting', icon: 'format_color_fill', tab: TAB, group: GROUP, kind: 'menu', edits: true, popular: true,
    tip: 'Highlight interesting cells, emphasize unusual values and visualize data with bars, color scales and icons', keywords: 'conditional formatting rules highlight',
    run: (app) => { const c = getCommand('cond.menu'); showMenu(c!.menu!(app), { anchor: app.ui.anchorFor('cond.menu'), onPick: (e) => { if (e.id) app.exec(e.id, e.arg); } }); },
    menu: condMenu,
  },
  quickCmd('cond.greaterThan', 'Highlight Greater Than', 'chevron_right', 'greaterThan', 'Format cells greater than a value'),
  quickCmd('cond.lessThan', 'Highlight Less Than', 'chevron_left', 'lessThan', 'Format cells less than a value'),
  quickCmd('cond.between', 'Highlight Between', 'unfold_less', 'between', 'Format cells between two values'),
  quickCmd('cond.equalTo', 'Highlight Equal To', 'drag_handle', 'equal', 'Format cells equal to a value'),
  quickCmd('cond.textContains', 'Highlight Text that Contains', 'match_case', 'text', 'Format cells that contain some text'),
  quickCmd('cond.dateOccurring', 'Highlight a Date Occurring', 'event', 'date', 'Format dates such as yesterday, this week or next month'),
  quickCmd('cond.duplicateValues', 'Highlight Duplicate Values', 'content_copy', 'duplicate', 'Format duplicate or unique values'),
  quickCmd('cond.top10', 'Top 10 Items', 'vertical_align_top', 'top', 'Format the highest values'),
  quickCmd('cond.top10Percent', 'Top 10%', 'percent', 'topPct', 'Format the highest percent of values'),
  quickCmd('cond.bottom10', 'Bottom 10 Items', 'vertical_align_bottom', 'bottom', 'Format the lowest values'),
  quickCmd('cond.bottom10Percent', 'Bottom 10%', 'percent', 'bottomPct', 'Format the lowest percent of values'),
  quickCmd('cond.aboveAverage', 'Above Average', 'trending_up', 'above', 'Format values above the average'),
  quickCmd('cond.belowAverage', 'Below Average', 'trending_down', 'below', 'Format values below the average'),
  {
    id: 'cond.dataBar', label: 'Data Bars', icon: 'data_thresholding', tab: TAB, group: GROUP, kind: 'split', edits: true, keywords: 'conditional formatting bar',
    tip: 'Show a bar in each cell whose length represents the value',
    run: (app, arg?: { color?: string; gradient?: boolean }) => { if (!canFormat(app)) return; commitRule(app, dataBarRule(targetRanges(app), arg?.color || BAR_COLORS[0].color, arg?.gradient !== false), 'Data Bars'); },
    menu: (app) => (condMenu(app).find((e) => e.label === 'Data Bars')!.submenu || []),
  },
  {
    id: 'cond.colorScale', label: 'Color Scales', icon: 'gradient', tab: TAB, group: GROUP, kind: 'split', edits: true, keywords: 'conditional formatting heat map',
    tip: 'Shade cells with a color gradient by value',
    run: (app, i?: number | string[]) => { if (!canFormat(app)) return; const colors = Array.isArray(i) ? i : COLOR_SCALES[Number(i) || 0]?.colors || COLOR_SCALES[0].colors; commitRule(app, colorScaleRule(targetRanges(app), colors), 'Color Scales'); },
    menu: (app) => (condMenu(app).find((e) => e.label === 'Color Scales')!.submenu || []),
  },
  {
    id: 'cond.iconSet', label: 'Icon Sets', icon: 'traffic', tab: TAB, group: GROUP, kind: 'split', edits: true, keywords: 'conditional formatting icons arrows traffic lights',
    tip: 'Show an icon in each cell according to its value',
    run: (app, set?: string) => { if (!canFormat(app)) return; commitRule(app, iconSetRule(targetRanges(app), set && ICON_SETS[set] ? set : '3TrafficLights1'), 'Icon Sets'); },
    menu: (app) => (condMenu(app).find((e) => e.label === 'Icon Sets')!.submenu || []),
  },
  {
    id: 'cond.newRule', label: 'New Formatting Rule', icon: 'add_box', tab: TAB, group: GROUP, edits: true, keywords: 'conditional formatting formula rule',
    tip: 'Create a conditional formatting rule of any type, including formula-based rules',
    run: async (app, arg?: { rule?: Partial<CFRule> }) => {
      if (!canFormat(app)) return;
      if (arg?.rule?.type) { commitRule(app, { ...arg.rule, ranges: arg.rule.ranges || targetRanges(app) } as CFRule, 'New Rule'); return; }
      const rule = await ruleDialog(app, undefined, targetRanges(app));
      if (rule) commitRule(app, rule, 'New Rule');
    },
  },
  { id: 'cond.manage', label: 'Manage Rules', icon: 'tune', tab: TAB, group: GROUP, edits: true, keywords: 'conditional formatting rules manager priority', tip: 'Edit, delete, reorder and change the ranges of conditional formatting rules', run: (app) => manageRules(app) },
  { id: 'cond.clearSelection', label: 'Clear Rules from Selected Cells', icon: 'format_color_reset', tab: TAB, group: GROUP, edits: true, keywords: 'conditional formatting remove', run: (app) => clearSelection(app) },
  { id: 'cond.clearSheet', label: 'Clear Rules from Entire Sheet', icon: 'format_color_reset', tab: TAB, group: GROUP, edits: true, keywords: 'conditional formatting remove', run: (app) => clearSheet(app) },
  { id: 'cond.clearTable', label: 'Clear Rules from This Table', icon: 'format_color_reset', tab: TAB, group: GROUP, edits: true, keywords: 'conditional formatting remove', run: (app) => clearTable(app) },
  { id: 'cond.clearPivot', label: 'Clear Rules from This PivotTable', icon: 'format_color_reset', tab: TAB, group: GROUP, edits: true, keywords: 'conditional formatting remove', run: (app) => clearPivot(app) },
]);

void MAX_ROWS; void MAX_COLS; void confirmDialog;
