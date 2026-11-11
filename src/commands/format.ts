/** Home tab formatting: font, alignment, number formats, borders, cell styles and the Format Cells dialog. */
import { def } from './registry';
import type { App } from '../app';
import type { CellStyle, Border, NamedCellStyle } from '../model/types';
import { applyStyle, toggleStyle, activeStyle, activeCell, selRanges } from './ops';
import { FORMAT_CATEGORIES, formatDecimals, withDecimals, formatValue } from '../model/numfmt';
import { colorMenu, openDialog, toast, confirmDialog, THEME_COLORS, STANDARD_COLORS } from '../ui/dialog';
import { FONT_LIST } from '../fonts';
import { MAX_COLS, MAX_ROWS, type Range } from '../model/address';
import { h } from '../ui/dom';
import { escHtml } from '../util/sanitize';

const SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];

function lastColor(app: App, key: string, dflt: string): string { return (app as any)['last_' + key] || dflt; }
function setLastColor(app: App, key: string, v: string): void { (app as any)['last_' + key] = v; const el = app.ui.root.querySelector<HTMLElement>(`[data-cmd="${key === 'fill' ? 'fmt.fill' : 'fmt.color'}"] .ql-colorbar`); if (el) el.style.background = v; }

def([
  { id: 'fmt.font', label: 'Font', icon: 'font_download', tab: 'Home', group: 'Font', kind: 'font', keys: ['Mod-Shift-f'], edits: true, run: (app, f) => { if (!f) return app.exec('fmt.fontDialog'); applyStyle(app, { font: String(f) }, 'Font'); }, menu: () => FONT_LIST.map((f) => ({ arg: f, label: f, preview: `font-family:"${f}"` })) },
  { id: 'fmt.size', label: 'Font Size', icon: 'format_size', tab: 'Home', group: 'Font', kind: 'size', edits: true, run: (app, v) => { const n = Number(v); if (!(n >= 1 && n <= 409)) { toast('The size must be between 1 and 409.'); return; } applyStyle(app, { size: n }, 'Font Size'); }, menu: () => SIZES.map((s) => ({ id: 'fmt.size', arg: s, label: String(s) })), value: (app) => activeStyle(app).size || app.book.styles[0]?.size || 11 },
  { id: 'fmt.grow', label: 'Increase Font Size', icon: 'text_increase', tab: 'Home', group: 'Font', keys: ['Mod-Shift-.'], edits: true, run: (app) => applyStyle(app, (st) => ({ size: SIZES.find((s) => s > (st.size || 11)) || (st.size || 11) + 10 }), 'Grow Font') },
  { id: 'fmt.shrink', label: 'Decrease Font Size', icon: 'text_decrease', tab: 'Home', group: 'Font', keys: ['Mod-Shift-,'], edits: true, run: (app) => applyStyle(app, (st) => ({ size: [...SIZES].reverse().find((s) => s < (st.size || 11)) || Math.max(1, (st.size || 11) - 1) }), 'Shrink Font') },
  { id: 'fmt.bold', label: 'Bold', icon: 'format_bold', tab: 'Home', group: 'Font', keys: ['Mod-b', 'Mod-2'], kind: 'toggle', popular: true, edits: true, run: (app) => toggleStyle(app, 'bold', 'Bold'), isActive: (app) => !!activeStyle(app).bold },
  { id: 'fmt.italic', label: 'Italic', icon: 'format_italic', tab: 'Home', group: 'Font', keys: ['Mod-i', 'Mod-3'], kind: 'toggle', popular: true, edits: true, run: (app) => toggleStyle(app, 'italic', 'Italic'), isActive: (app) => !!activeStyle(app).italic },
  {
    id: 'fmt.underline', label: 'Underline', icon: 'format_underlined', tab: 'Home', group: 'Font', keys: ['Mod-u', 'Mod-4'], kind: 'split', popular: true, edits: true,
    run: (app, arg) => { const cur = activeStyle(app).underline; const want = typeof arg === 'string' ? arg : 'single'; applyStyle(app, { underline: cur === want ? undefined : want as CellStyle['underline'] }, 'Underline'); },
    isActive: (app) => !!activeStyle(app).underline,
    menu: () => [{ id: 'fmt.underline', arg: 'single', label: 'Underline' }, { id: 'fmt.underline', arg: 'double', label: 'Double Underline' }, { id: 'fmt.underline', arg: 'singleAccounting', label: 'Single Accounting' }, { id: 'fmt.underline', arg: 'doubleAccounting', label: 'Double Accounting' }],
  },
  { id: 'fmt.doubleUnderline', label: 'Double Underline', icon: 'format_underlined', tab: 'Home', group: 'Font', edits: true, run: (app) => app.exec('fmt.underline', 'double') },
  { id: 'fmt.strike', label: 'Strikethrough', icon: 'strikethrough_s', tab: 'Home', group: 'Font', keys: ['Mod-5'], kind: 'toggle', edits: true, run: (app) => toggleStyle(app, 'strike', 'Strikethrough'), isActive: (app) => !!activeStyle(app).strike },
  { id: 'fmt.superscript', label: 'Superscript', icon: 'superscript', tab: 'Home', group: 'Font', kind: 'toggle', edits: true, run: (app) => applyStyle(app, { vertAlign: activeStyle(app).vertAlign === 'superscript' ? undefined : 'superscript' }, 'Superscript'), isActive: (app) => activeStyle(app).vertAlign === 'superscript' },
  { id: 'fmt.subscript', label: 'Subscript', icon: 'subscript', tab: 'Home', group: 'Font', kind: 'toggle', edits: true, run: (app) => applyStyle(app, { vertAlign: activeStyle(app).vertAlign === 'subscript' ? undefined : 'subscript' }, 'Subscript'), isActive: (app) => activeStyle(app).vertAlign === 'subscript' },
  {
    id: 'fmt.color', label: 'Font Color', icon: 'format_color_text', tab: 'Home', group: 'Font', kind: 'color', popular: true, edits: true,
    run: async (app, arg) => {
      if (arg === 'menu') { const c = await colorMenu(app.ui.anchorFor('fmt.color'), { auto: 'Automatic', title: 'Font Color' }); if (c === undefined) return; if (c) setLastColor(app, 'color', c); applyStyle(app, { color: c || undefined }, 'Font Color'); return; }
      applyStyle(app, { color: typeof arg === 'string' ? arg : lastColor(app, 'color', '#c00000') }, 'Font Color');
    },
  },
  {
    id: 'fmt.fill', label: 'Fill Color', icon: 'format_color_fill', tab: 'Home', group: 'Font', kind: 'color', popular: true, edits: true,
    run: async (app, arg) => {
      if (arg === 'menu') { const c = await colorMenu(app.ui.anchorFor('fmt.fill'), { none: 'No Fill', title: 'Fill Color' }); if (c === undefined) return; if (c) setLastColor(app, 'fill', c); applyStyle(app, { fill: c || undefined, pattern: undefined, gradient: undefined }, 'Fill Color'); return; }
      applyStyle(app, { fill: typeof arg === 'string' ? arg : lastColor(app, 'fill', '#ffff00') }, 'Fill Color');
    },
  },
  { id: 'fmt.noFill', label: 'No Fill', icon: 'format_color_reset', tab: 'Home', group: 'Font', edits: true, run: (app) => applyStyle(app, { fill: undefined, pattern: undefined, gradient: undefined }, 'No Fill') },
  { id: 'fmt.fontDialog', label: 'Format Cells: Font…', icon: 'text_format', tab: 'Home', group: 'Font', keys: ['Mod-Shift-p'], edits: true, run: (app) => app.exec('fmt.cellsDialog', 'font') },
]);

// --------------------------------------------------------------------------------------------
// Borders

function borderPatch(app: App, kind: string, b: Border | undefined): (st: CellStyle, r: number, c: number) => Partial<CellStyle> {
  const rgs = selRanges(app);
  const rg = rgs[rgs.length - 1];
  const r2 = Math.min(rg.r2, MAX_ROWS - 1), c2 = Math.min(rg.c2, MAX_COLS - 1);
  return (_st, r, c) => {
    const top = r === rg.r1, bottom = r === r2, left = c === rg.c1, right = c === c2;
    switch (kind) {
      case 'bottom': return bottom ? { bb: b } : {};
      case 'top': return top ? { bt: b } : {};
      case 'left': return left ? { bl: b } : {};
      case 'right': return right ? { br: b } : {};
      case 'none': return { bt: undefined, bb: undefined, bl: undefined, br: undefined, diagUp: undefined, diagDown: undefined };
      case 'all': return { bt: b, bb: b, bl: b, br: b };
      case 'outside': return { ...(top ? { bt: b } : {}), ...(bottom ? { bb: b } : {}), ...(left ? { bl: b } : {}), ...(right ? { br: b } : {}) };
      case 'inside': return { ...(!top ? { bt: b } : {}), ...(!bottom ? { bb: b } : {}), ...(!left ? { bl: b } : {}), ...(!right ? { br: b } : {}) };
      case 'thickOutside': { const t: Border = { style: 'medium', color: b?.color }; return { ...(top ? { bt: t } : {}), ...(bottom ? { bb: t } : {}), ...(left ? { bl: t } : {}), ...(right ? { br: t } : {}) }; }
      case 'bottomDouble': return bottom ? { bb: { style: 'double', color: b?.color } } : {};
      case 'thickBottom': return bottom ? { bb: { style: 'medium', color: b?.color } } : {};
      case 'topBottom': return { ...(top ? { bt: b } : {}), ...(bottom ? { bb: b } : {}) };
      case 'topThickBottom': return { ...(top ? { bt: b } : {}), ...(bottom ? { bb: { style: 'medium', color: b?.color } } : {}) };
      case 'topDoubleBottom': return { ...(top ? { bt: b } : {}), ...(bottom ? { bb: { style: 'double', color: b?.color } } : {}) };
      case 'diagDown': return { diagDown: b };
      case 'diagUp': return { diagUp: b };
      case 'insideH': return { ...(!top ? { bt: b } : {}), ...(!bottom ? { bb: b } : {}) };
      case 'insideV': return { ...(!left ? { bl: b } : {}), ...(!right ? { br: b } : {}) };
    }
    return {};
  };
}

const BORDER_STYLES: Border['style'][] = ['thin', 'hair', 'dotted', 'dashed', 'dashDot', 'mediumDashed', 'medium', 'thick', 'double'];

def([
  {
    id: 'fmt.borders', label: 'Borders', icon: 'border_bottom', tab: 'Home', group: 'Font', kind: 'split', popular: true, edits: true,
    run: (app, kind) => {
      const a = app as any;
      const k = typeof kind === 'string' ? kind : a.lastBorder || 'bottom';
      a.lastBorder = k;
      const b: Border = { style: a.borderStyle || 'thin', color: a.borderColor };
      applyStyle(app, borderPatch(app, k, k === 'none' ? undefined : b), 'Borders');
    },
    menu: (app) => {
      const a = app as any;
      return [
        { heading: 'Borders' },
        { id: 'fmt.borders', arg: 'bottom', label: 'Bottom Border', icon: 'border_bottom' }, { id: 'fmt.borders', arg: 'top', label: 'Top Border', icon: 'border_top' }, { id: 'fmt.borders', arg: 'left', label: 'Left Border', icon: 'border_left' }, { id: 'fmt.borders', arg: 'right', label: 'Right Border', icon: 'border_right' },
        { separator: true },
        { id: 'fmt.borders', arg: 'none', label: 'No Border', icon: 'border_clear' }, { id: 'fmt.borders', arg: 'all', label: 'All Borders', icon: 'border_all' }, { id: 'fmt.borders', arg: 'outside', label: 'Outside Borders', icon: 'border_outer' }, { id: 'fmt.borders', arg: 'thickOutside', label: 'Thick Outside Borders', icon: 'border_style' },
        { separator: true },
        { id: 'fmt.borders', arg: 'bottomDouble', label: 'Bottom Double Border', icon: 'border_bottom' }, { id: 'fmt.borders', arg: 'thickBottom', label: 'Thick Bottom Border', icon: 'border_bottom' }, { id: 'fmt.borders', arg: 'topBottom', label: 'Top and Bottom Border', icon: 'border_horizontal' }, { id: 'fmt.borders', arg: 'topThickBottom', label: 'Top and Thick Bottom Border', icon: 'border_horizontal' }, { id: 'fmt.borders', arg: 'topDoubleBottom', label: 'Top and Double Bottom Border', icon: 'border_horizontal' },
        { id: 'fmt.borders', arg: 'inside', label: 'Inside Borders', icon: 'border_inner' }, { id: 'fmt.borders', arg: 'insideH', label: 'Inside Horizontal', icon: 'border_horizontal' }, { id: 'fmt.borders', arg: 'insideV', label: 'Inside Vertical', icon: 'border_vertical' }, { id: 'fmt.borders', arg: 'diagDown', label: 'Diagonal Down', icon: 'south_east' }, { id: 'fmt.borders', arg: 'diagUp', label: 'Diagonal Up', icon: 'north_east' },
        { heading: 'Draw Borders' },
        { id: 'fmt.borderColor', label: 'Line Color…', icon: 'border_color', swatch: a.borderColor || '#000000' },
        { label: 'Line Style', icon: 'line_style', submenu: BORDER_STYLES.map((s) => ({ id: 'fmt.borderStyle', arg: s, label: s, checked: (a.borderStyle || 'thin') === s })) },
        { id: 'fmt.cellsDialog', arg: 'border', label: 'More Borders…', icon: 'border_all' },
      ];
    },
  },
  { id: 'fmt.borderColor', label: 'Border Line Color', icon: 'border_color', tab: 'Home', group: 'Font', run: async (app) => { const c = await colorMenu(app.ui.anchorFor('fmt.borders'), { auto: 'Automatic' }); if (c !== undefined) (app as any).borderColor = c || undefined; } },
  { id: 'fmt.borderStyle', label: 'Border Line Style', icon: 'line_style', tab: 'Home', group: 'Font', run: (app, s) => { (app as any).borderStyle = s || 'thin'; } },
  ...(['bottom', 'top', 'left', 'right', 'all', 'outside', 'none', 'thickOutside', 'bottomDouble'] as const).map((k) => ({ id: `fmt.border.${k}`, label: `${{ bottom: 'Bottom', top: 'Top', left: 'Left', right: 'Right', all: 'All', outside: 'Outside', none: 'No', thickOutside: 'Thick Outside', bottomDouble: 'Bottom Double' }[k]} Border${k === 'all' || k === 'outside' || k === 'thickOutside' ? 's' : ''}`, icon: k === 'none' ? 'border_clear' : k === 'all' ? 'border_all' : k === 'outside' || k === 'thickOutside' ? 'border_outer' : `border_${k === 'bottomDouble' ? 'bottom' : k}`, tab: 'Home', group: 'Font', keys: k === 'outside' ? ['Mod-Shift-7'] : k === 'none' ? ['Mod-Shift--'] : undefined, edits: true, run: (app: App) => app.exec('fmt.borders', k) })),
]);

// --------------------------------------------------------------------------------------------
// Alignment

const align = (id: string, label: string, icon: string, key: 'hAlign' | 'vAlign', v: string, keys?: string[]) => ({
  id, label, icon, tab: 'Home', group: 'Alignment', kind: 'toggle' as const, keys, edits: true,
  run: (app: App) => applyStyle(app, { [key]: activeStyle(app)[key] === v ? undefined : v } as any, label),
  isActive: (app: App) => activeStyle(app)[key] === v,
});

def([
  align('fmt.alignTop', 'Top Align', 'vertical_align_top', 'vAlign', 'top'),
  align('fmt.alignMiddle', 'Middle Align', 'vertical_align_center', 'vAlign', 'middle'),
  align('fmt.alignBottom', 'Bottom Align', 'vertical_align_bottom', 'vAlign', 'bottom'),
  align('fmt.alignLeft', 'Align Left', 'format_align_left', 'hAlign', 'left'),
  align('fmt.alignCenter', 'Center', 'format_align_center', 'hAlign', 'center'),
  align('fmt.alignRight', 'Align Right', 'format_align_right', 'hAlign', 'right'),
  align('fmt.alignJustify', 'Justify', 'format_align_justify', 'hAlign', 'justify'),
  align('fmt.alignFill', 'Fill (repeat)', 'format_line_spacing', 'hAlign', 'fill'),
  align('fmt.alignCenterAcross', 'Center Across Selection', 'align_horizontal_center', 'hAlign', 'centerContinuous'),
  align('fmt.alignDistributed', 'Distributed', 'format_align_justify', 'hAlign', 'distributed'),
  { id: 'fmt.indentInc', label: 'Increase Indent', icon: 'format_indent_increase', tab: 'Home', group: 'Alignment', keys: ['Mod-Alt-Tab'], edits: true, run: (app) => applyStyle(app, (st) => ({ indent: Math.min(250, (st.indent || 0) + 1), hAlign: st.hAlign === 'right' ? 'right' : 'left' }), 'Increase Indent') },
  { id: 'fmt.indentDec', label: 'Decrease Indent', icon: 'format_indent_decrease', tab: 'Home', group: 'Alignment', keys: ['Mod-Alt-Shift-Tab'], edits: true, run: (app) => applyStyle(app, (st) => ({ indent: Math.max(0, (st.indent || 0) - 1) || undefined }), 'Decrease Indent') },
  { id: 'fmt.wrap', label: 'Wrap Text', icon: 'wrap_text', tab: 'Home', group: 'Alignment', kind: 'toggle', popular: true, edits: true, run: (app) => { toggleStyle(app, 'wrap', 'Wrap Text'); app.exec('format.autofitRows', { quiet: true }); }, isActive: (app) => !!activeStyle(app).wrap },
  { id: 'fmt.shrinkToFit', label: 'Shrink to Fit', icon: 'fit_screen', tab: 'Home', group: 'Alignment', kind: 'toggle', edits: true, run: (app) => toggleStyle(app, 'shrink', 'Shrink to Fit'), isActive: (app) => !!activeStyle(app).shrink },
  {
    id: 'fmt.orientation', label: 'Orientation', icon: 'text_rotation_angleup', tab: 'Home', group: 'Alignment', kind: 'menu', edits: true,
    run: (app, rot) => { if (rot === 'dialog') return app.exec('fmt.cellsDialog', 'alignment'); applyStyle(app, { rotate: Number(rot) || undefined }, 'Orientation'); },
    menu: (app) => { const cur = activeStyle(app).rotate || 0; return [{ id: 'fmt.orientation', arg: 45, label: 'Angle Counterclockwise', icon: 'text_rotation_angleup', checked: cur === 45 }, { id: 'fmt.orientation', arg: -45, label: 'Angle Clockwise', icon: 'text_rotation_angledown', checked: cur === -45 }, { id: 'fmt.orientation', arg: 255, label: 'Vertical Text', icon: 'text_rotate_vertical', checked: cur === 255 }, { id: 'fmt.orientation', arg: 90, label: 'Rotate Text Up', icon: 'text_rotate_up', checked: cur === 90 }, { id: 'fmt.orientation', arg: -90, label: 'Rotate Text Down', icon: 'text_rotation_down', checked: cur === -90 }, { id: 'fmt.orientation', arg: 0, label: 'Horizontal', icon: 'text_rotation_none', checked: !cur }, { separator: true }, { id: 'fmt.orientation', arg: 'dialog', label: 'Format Cell Alignment…', icon: 'tune' }]; },
  },
  { id: 'fmt.rtl', label: 'Right-to-Left Text', icon: 'format_textdirection_r_to_l', tab: 'Home', group: 'Alignment', kind: 'toggle', edits: true, run: (app) => toggleStyle(app, 'rtl', 'Text Direction'), isActive: (app) => !!activeStyle(app).rtl },
  { id: 'fmt.alignmentDialog', label: 'Format Cells: Alignment…', icon: 'format_align_center', tab: 'Home', group: 'Alignment', edits: true, run: (app) => app.exec('fmt.cellsDialog', 'alignment') },
]);

// Merge ---------------------------------------------------------------------------------------

async function merge(app: App, mode: 'center' | 'across' | 'merge' | 'unmerge'): Promise<void> {
  const s = app.sheet;
  const book = app.book;
  const rgs = selRanges(app);
  const dataLoss = mode !== 'unmerge' && rgs.some((rg) => {
    let n = 0;
    s.forEachIn(rg, (_r, _c, cell) => { if (cell.v != null || cell.f != null) n++; });
    return n > 1;
  });
  if (dataLoss && !(await confirmDialog('Merge Cells', 'Merging cells only keeps the upper-left value and discards other values.', 'OK'))) return;
  book.transact(mode === 'unmerge' ? 'Unmerge Cells' : 'Merge Cells', () => {
    book.recordSheet(s);
    for (const rg of rgs) {
      s.merges = s.merges.filter((m) => !(m.r1 <= rg.r2 && m.r2 >= rg.r1 && m.c1 <= rg.c2 && m.c2 >= rg.c1));
      if (mode === 'unmerge') continue;
      const rows: Range[] = mode === 'across' ? Array.from({ length: rg.r2 - rg.r1 + 1 }, (_, i) => ({ r1: rg.r1 + i, c1: rg.c1, r2: rg.r1 + i, c2: rg.c2 })) : [rg];
      for (const m of rows) {
        if (m.r1 === m.r2 && m.c1 === m.c2) continue;
        for (let r = m.r1; r <= m.r2; r++) for (let c = m.c1; c <= m.c2; c++) if (r !== m.r1 || c !== m.c1) { const cur = s.get(r, c); if (cur && (cur.v != null || cur.f != null)) s.set(r, c, cur.s ? { s: cur.s } : undefined); }
        s.merges.push(m);
      }
    }
    if (mode === 'center') for (const rg of rgs) { const cur = s.get(rg.r1, rg.c1); s.patch(rg.r1, rg.c1, { s: book.derivedStyle(cur?.s ?? 0, { hAlign: 'center' }) }); }
  });
}

def([
  {
    id: 'fmt.merge', label: 'Merge & Center', icon: 'cell_merge', tab: 'Home', group: 'Alignment', kind: 'split', popular: true, edits: true,
    run: (app, mode) => { const s = app.sheet; const { r, c } = activeCell(app); if (!mode && s.mergeAt(r, c)) return merge(app, 'unmerge'); merge(app, (mode as any) || 'center'); },
    isActive: (app) => { const { r, c } = activeCell(app); return !!app.sheet.mergeAt(r, c); },
    menu: () => [{ id: 'fmt.merge', arg: 'center', label: 'Merge & Center', icon: 'cell_merge' }, { id: 'fmt.merge', arg: 'across', label: 'Merge Across', icon: 'view_column' }, { id: 'fmt.merge', arg: 'merge', label: 'Merge Cells', icon: 'crop_square' }, { id: 'fmt.merge', arg: 'unmerge', label: 'Unmerge Cells', icon: 'splitscreen' }],
  },
  { id: 'fmt.mergeAcross', label: 'Merge Across', icon: 'view_column', tab: 'Home', group: 'Alignment', edits: true, run: (app) => merge(app, 'across') },
  { id: 'fmt.mergeCells', label: 'Merge Cells', icon: 'crop_square', tab: 'Home', group: 'Alignment', edits: true, run: (app) => merge(app, 'merge') },
  { id: 'fmt.unmerge', label: 'Unmerge Cells', icon: 'splitscreen', tab: 'Home', group: 'Alignment', edits: true, run: (app) => merge(app, 'unmerge') },
]);

// --------------------------------------------------------------------------------------------
// Number formats

export const NUMBER_PRESETS: { id: string; label: string; fmt: string; icon: string; keys?: string[] }[] = [
  { id: 'general', label: 'General', fmt: 'General', icon: 'abc', keys: ['Mod-Shift-`', 'Mod-Shift-~'] },
  { id: 'number', label: 'Number', fmt: '0.00', icon: 'pin' },
  { id: 'numberComma', label: 'Number with Separator', fmt: '#,##0.00', icon: 'pin', keys: ['Mod-Shift-1', 'Mod-!'] },
  { id: 'currency', label: 'Currency', fmt: '$#,##0.00', icon: 'attach_money', keys: ['Mod-Shift-4', 'Mod-$'] },
  { id: 'accounting', label: 'Accounting', fmt: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)', icon: 'request_quote' },
  { id: 'shortDate', label: 'Short Date', fmt: 'm/d/yyyy', icon: 'calendar_today', keys: ['Mod-Shift-3', 'Mod-#'] },
  { id: 'longDate', label: 'Long Date', fmt: 'dddd, mmmm d, yyyy', icon: 'event' },
  { id: 'isoDate', label: 'ISO Date', fmt: 'yyyy-mm-dd', icon: 'event' },
  { id: 'time', label: 'Time', fmt: 'h:mm:ss AM/PM', icon: 'schedule', keys: ['Mod-Shift-2', 'Mod-@'] },
  { id: 'percentage', label: 'Percentage', fmt: '0%', icon: 'percent', keys: ['Mod-Shift-5', 'Mod-%'] },
  { id: 'fraction', label: 'Fraction', fmt: '# ?/?', icon: 'calculate' },
  { id: 'scientific', label: 'Scientific', fmt: '0.00E+00', icon: 'science', keys: ['Mod-Shift-6', 'Mod-^'] },
  { id: 'text', label: 'Text', fmt: '@', icon: 'text_fields' },
];

function formatLabel(fmt: string | undefined): string {
  if (!fmt || fmt === 'General') return 'General';
  const p = NUMBER_PRESETS.find((x) => x.fmt === fmt);
  if (p) return p.label;
  for (const cat of FORMAT_CATEGORIES) if (cat.formats.includes(fmt)) return cat.name;
  if (/[ymd]/i.test(fmt.replace(/"[^"]*"/g, '')) && !/^[#0.,]+$/.test(fmt)) return 'Date';
  if (/%/.test(fmt)) return 'Percentage';
  if (/[$€£¥]/.test(fmt)) return 'Currency';
  return 'Custom';
}

def([
  {
    id: 'fmt.numberFormat', label: 'Number Format', icon: 'pin', tab: 'Home', group: 'Number', kind: 'combo', edits: true,
    value: (app) => formatLabel(activeStyle(app).numFmt),
    run: (app, fmt) => { if (fmt === 'more') return app.exec('fmt.cellsDialog', 'number'); if (typeof fmt === 'string') applyStyle(app, { numFmt: fmt === 'General' ? undefined : fmt }, 'Number Format'); },
    menu: (app) => {
      const { r, c } = activeCell(app);
      const v = app.sheet.value(r, c);
      const sample = (f: string) => (typeof v === 'number' ? formatValue(v, f, { date1904: app.book.calc.date1904 }).text : '');
      return [...NUMBER_PRESETS.map((p) => ({ id: 'fmt.numberFormat', arg: p.fmt, label: p.label + (sample(p.fmt) ? `   ${sample(p.fmt)}` : ''), icon: p.icon })), { separator: true }, { id: 'fmt.numberFormat', arg: 'more', label: 'More Number Formats…', icon: 'more_horiz' }];
    },
  },
  ...NUMBER_PRESETS.map((p) => ({ id: `fmt.num.${p.id}`, label: `${p.label} Format`, icon: p.icon, tab: 'Home', group: 'Number', keys: p.keys, edits: true, run: (app: App) => applyStyle(app, { numFmt: p.fmt === 'General' ? undefined : p.fmt }, `${p.label} Format`) })),
  {
    id: 'fmt.accounting', label: 'Accounting Number Format', icon: 'request_quote', tab: 'Home', group: 'Number', kind: 'split', popular: true, edits: true,
    run: (app, sym) => {
      const s = typeof sym === 'string' ? sym : '$';
      applyStyle(app, { numFmt: `_("${s}"* #,##0.00_);_("${s}"* \\(#,##0.00\\);_("${s}"* "-"??_);_(@_)` }, 'Accounting Format');
    },
    menu: () => [{ id: 'fmt.accounting', arg: '$', label: '$ English (United States)' }, { id: 'fmt.accounting', arg: '£', label: '£ English (United Kingdom)' }, { id: 'fmt.accounting', arg: '€', label: '€ Euro' }, { id: 'fmt.accounting', arg: '¥', label: '¥ Japanese / Chinese' }, { id: 'fmt.accounting', arg: '₹', label: '₹ Indian Rupee' }, { id: 'fmt.accounting', arg: 'CHF', label: 'CHF Swiss Franc' }, { separator: true }, { id: 'fmt.numberFormat', arg: 'more', label: 'More Accounting Formats…' }],
  },
  { id: 'fmt.percent', label: 'Percent Style', icon: 'percent', tab: 'Home', group: 'Number', popular: true, edits: true, run: (app) => applyStyle(app, { numFmt: '0%' }, 'Percent Style') },
  { id: 'fmt.comma', label: 'Comma Style', icon: 'format_quote', tab: 'Home', group: 'Number', edits: true, run: (app) => applyStyle(app, { numFmt: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)' }, 'Comma Style') },
  {
    id: 'fmt.decInc', label: 'Increase Decimal', icon: 'decimal_increase', tab: 'Home', group: 'Number', edits: true,
    run: (app) => { const { r, c } = activeCell(app); const v = app.sheet.value(r, c); const cur = typeof v === 'number' ? (String(v).split('.')[1]?.length ?? 0) : 0; applyStyle(app, (st) => ({ numFmt: withDecimals(st.numFmt, 1, formatDecimals(st.numFmt) < 0 ? cur : undefined) }), 'Increase Decimal'); },
  },
  {
    id: 'fmt.decDec', label: 'Decrease Decimal', icon: 'decimal_decrease', tab: 'Home', group: 'Number', edits: true,
    run: (app) => { const { r, c } = activeCell(app); const v = app.sheet.value(r, c); const cur = typeof v === 'number' ? (String(v).split('.')[1]?.length ?? 0) : 0; applyStyle(app, (st) => ({ numFmt: withDecimals(st.numFmt, -1, formatDecimals(st.numFmt) < 0 ? cur : undefined) }), 'Decrease Decimal'); },
  },
  { id: 'fmt.numberDialog', label: 'Format Cells: Number…', icon: 'pin', tab: 'Home', group: 'Number', edits: true, run: (app) => app.exec('fmt.cellsDialog', 'number') },
]);

// --------------------------------------------------------------------------------------------
// Cell styles

export const BUILTIN_CELL_STYLES: NamedCellStyle[] = [
  { name: 'Normal', style: {}, builtin: true },
  { name: 'Good', style: { fill: '#c6efce', color: '#006100' }, builtin: true },
  { name: 'Bad', style: { fill: '#ffc7ce', color: '#9c0006' }, builtin: true },
  { name: 'Neutral', style: { fill: '#ffeb9c', color: '#9c5700' }, builtin: true },
  { name: 'Calculation', style: { fill: '#f2f2f2', color: '#fa7d00', bold: true, bt: { style: 'thin', color: '#7f7f7f' }, bb: { style: 'thin', color: '#7f7f7f' }, bl: { style: 'thin', color: '#7f7f7f' }, br: { style: 'thin', color: '#7f7f7f' } }, builtin: true },
  { name: 'Check Cell', style: { fill: '#a5a5a5', color: '#ffffff', bold: true, bt: { style: 'double', color: '#3f3f3f' }, bb: { style: 'double', color: '#3f3f3f' }, bl: { style: 'double', color: '#3f3f3f' }, br: { style: 'double', color: '#3f3f3f' } }, builtin: true },
  { name: 'Explanatory Text', style: { italic: true, color: '#7f7f7f' }, builtin: true },
  { name: 'Input', style: { fill: '#ffcc99', color: '#3f3f76', bt: { style: 'thin', color: '#7f7f7f' }, bb: { style: 'thin', color: '#7f7f7f' }, bl: { style: 'thin', color: '#7f7f7f' }, br: { style: 'thin', color: '#7f7f7f' } }, builtin: true },
  { name: 'Linked Cell', style: { color: '#fa7d00', bb: { style: 'double', color: '#ff8001' } }, builtin: true },
  { name: 'Note', style: { fill: '#ffffcc', bt: { style: 'thin', color: '#b2b2b2' }, bb: { style: 'thin', color: '#b2b2b2' }, bl: { style: 'thin', color: '#b2b2b2' }, br: { style: 'thin', color: '#b2b2b2' } }, builtin: true },
  { name: 'Output', style: { fill: '#f2f2f2', color: '#3f3f3f', bold: true, bt: { style: 'thin', color: '#3f3f3f' }, bb: { style: 'thin', color: '#3f3f3f' }, bl: { style: 'thin', color: '#3f3f3f' }, br: { style: 'thin', color: '#3f3f3f' } }, builtin: true },
  { name: 'Warning Text', style: { color: '#ff0000' }, builtin: true },
  { name: 'Title', style: { size: 18, color: '#1f3b57', bold: false }, builtin: true },
  { name: 'Heading 1', style: { size: 15, bold: true, color: '#1f3b57', bb: { style: 'thick', color: '#2f6fb0' } }, builtin: true },
  { name: 'Heading 2', style: { size: 13, bold: true, color: '#1f3b57', bb: { style: 'thick', color: '#9ebde0' } }, builtin: true },
  { name: 'Heading 3', style: { bold: true, color: '#1f3b57', bb: { style: 'medium', color: '#9ebde0' } }, builtin: true },
  { name: 'Heading 4', style: { bold: true, color: '#1f3b57' }, builtin: true },
  { name: 'Total', style: { bold: true, bt: { style: 'thin', color: '#2f6fb0' }, bb: { style: 'double', color: '#2f6fb0' } }, builtin: true },
  { name: 'Comma', style: { numFmt: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)' }, builtin: true },
  { name: 'Comma [0]', style: { numFmt: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)' }, builtin: true },
  { name: 'Currency', style: { numFmt: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)' }, builtin: true },
  { name: 'Currency [0]', style: { numFmt: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)' }, builtin: true },
  { name: 'Percent', style: { numFmt: '0%' }, builtin: true },
  ...[1, 2, 3, 4, 5, 6].flatMap((i) => [
    { name: `Accent${i}`, style: { fill: ['#2f6fb0', '#e07b28', '#7a8a99', '#e3b505', '#4a90c2', '#5aa33f'][i - 1], color: '#ffffff' }, builtin: true },
    { name: `40% - Accent${i}`, style: { fill: ['#acc5e0', '#f3caa9', '#cad0d6', '#f4e29b', '#b7d3e7', '#bddaaf'][i - 1] }, builtin: true },
    { name: `20% - Accent${i}`, style: { fill: ['#d5e2f0', '#f9e4d4', '#e4e8eb', '#f9f0cd', '#dbe9f3', '#deedd7'][i - 1] }, builtin: true },
  ]),
];

def([
  {
    id: 'fmt.cellStyles', label: 'Cell Styles', icon: 'style', tab: 'Home', group: 'Styles', kind: 'gallery', edits: true,
    run: (app, name) => {
      if (name === 'new') return app.exec('fmt.newCellStyle');
      if (name === 'merge') return toast('Styles from other open workbooks are merged when you paste formatted cells.');
      const all = [...BUILTIN_CELL_STYLES, ...app.book.cellStyles];
      const st = all.find((s) => s.name === name);
      if (!st) return;
      // A named style replaces the cell's formatting (keeping nothing but the style).
      app.book.transact(`Style: ${st.name}`, () => {
        for (const rg of selRanges(app)) {
          const area = { r1: rg.r1, c1: rg.c1, r2: Math.min(rg.r2, (app.sheet.usedRange()?.r2 ?? rg.r1) + (rg.r2 - rg.r1 < 5000 ? rg.r2 : 0)), c2: Math.min(rg.c2, rg.c1 + 2000) };
          for (let r = area.r1; r <= Math.min(area.r2, rg.r2); r++) for (let c = area.c1; c <= area.c2; c++) app.sheet.patch(r, c, { s: app.book.styleId(st.style) || undefined });
        }
      });
    },
    menu: (app) => [...BUILTIN_CELL_STYLES, ...app.book.cellStyles].map((s) => ({ id: 'fmt.cellStyles', arg: s.name, label: s.name, preview: styleCss(s.style) })).concat([{ separator: true } as any, { id: 'fmt.cellStyles', arg: 'new', label: 'New Cell Style…' } as any]),
  },
  {
    id: 'fmt.newCellStyle', label: 'New Cell Style…', icon: 'add', tab: 'Home', group: 'Styles', edits: true,
    run: async (app) => {
      const r = await openDialog({ title: 'Style', fields: [{ key: 'name', label: 'Style name', type: 'text', value: 'Style 1', autofocus: true }], help: 'The new style takes the formatting of the active cell.' });
      if (!r?.name) return;
      app.book.transact('New Cell Style', () => {
        app.book.recordBook();
        app.book.cellStyles = app.book.cellStyles.filter((s) => s.name !== r.name);
        app.book.cellStyles.push({ name: r.name, style: { ...activeStyle(app) } });
      });
    },
  },
]);

export function styleCss(st: CellStyle): string {
  return [st.fill ? `background:${st.fill}` : '', st.color ? `color:${st.color}` : '', st.bold ? 'font-weight:700' : '', st.italic ? 'font-style:italic' : '', st.size && st.size > 12 ? 'font-size:13px' : '', st.bb ? `border-bottom:${st.bb.style === 'double' ? '3px double' : st.bb.style === 'thick' ? '3px solid' : '1px solid'} ${st.bb.color || '#000'}` : ''].filter(Boolean).join(';');
}

// --------------------------------------------------------------------------------------------
// Format Cells dialog

def([
  {
    id: 'fmt.cellsDialog', label: 'Format Cells…', icon: 'table_view', tab: 'Home', group: 'Cells', keys: ['Mod-1'], popular: true, edits: true,
    run: async (app, page?: string) => {
      const st = activeStyle(app);
      const { r, c } = activeCell(app);
      const v = app.sheet.value(r, c);
      const cats = FORMAT_CATEGORIES.map((x) => x.name).concat(['Custom']);
      const curCat = formatLabel(st.numFmt);
      const allFormats = [...new Set([...FORMAT_CATEGORIES.flatMap((x) => x.formats), ...app.book.customFormats])];
      const preview = (fmt: string) => (v == null ? '' : formatValue(v, fmt, { date1904: app.book.calc.date1904 }).text);
      const borderOpts = [{ value: '', label: '(keep)' }, { value: 'none', label: 'None' }, ...BORDER_STYLES.map((s) => ({ value: s, label: s }))];
      const res = await openDialog({
        title: 'Format Cells', wide: true,
        fields: [
          { key: 'hn', type: 'heading', label: 'Number' },
          { key: 'cat', label: 'Category', type: 'select', options: cats, value: cats.includes(curCat) ? curCat : 'Custom', half: true },
          { key: 'fmt', label: 'Type (format code)', type: 'combo', options: allFormats, value: st.numFmt || 'General', half: true },
          { key: 'sample', type: 'info', html: `<p class="ql-help">Sample: <b data-sample>${escHtml(preview(st.numFmt || 'General'))}</b></p>` },
          { key: 'ha', type: 'heading', label: 'Alignment' },
          { key: 'hAlign', label: 'Horizontal', type: 'select', options: ['general', 'left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed'], value: st.hAlign || 'general', half: true },
          { key: 'vAlign', label: 'Vertical', type: 'select', options: ['top', 'middle', 'bottom', 'justify', 'distributed'], value: st.vAlign || 'bottom', half: true },
          { key: 'indent', label: 'Indent', type: 'number', min: 0, max: 250, value: st.indent || 0, half: true },
          { key: 'rotate', label: 'Orientation (degrees, -90…90; 255 = vertical)', type: 'number', min: -90, max: 255, value: st.rotate || 0, half: true },
          { key: 'wrap', label: 'Wrap text', type: 'checkbox', value: !!st.wrap, half: true },
          { key: 'shrink', label: 'Shrink to fit', type: 'checkbox', value: !!st.shrink, half: true },
          { key: 'merge', label: 'Merge cells', type: 'checkbox', value: !!app.sheet.mergeAt(r, c), half: true },
          { key: 'rtl', label: 'Right-to-left', type: 'checkbox', value: !!st.rtl, half: true },
          { key: 'hf', type: 'heading', label: 'Font' },
          { key: 'font', label: 'Font', type: 'combo', options: FONT_LIST, value: st.font || app.book.styles[0]?.font || 'Carlito', half: true },
          { key: 'size', label: 'Size', type: 'number', min: 1, max: 409, step: 0.5, value: st.size || 11, half: true },
          { key: 'bold', label: 'Bold', type: 'checkbox', value: !!st.bold, half: true },
          { key: 'italic', label: 'Italic', type: 'checkbox', value: !!st.italic, half: true },
          { key: 'underline', label: 'Underline', type: 'select', options: [{ value: '', label: 'None' }, { value: 'single', label: 'Single' }, { value: 'double', label: 'Double' }, { value: 'singleAccounting', label: 'Single Accounting' }, { value: 'doubleAccounting', label: 'Double Accounting' }], value: st.underline || '', half: true },
          { key: 'color', label: 'Color', type: 'color', value: st.color || '#000000', half: true },
          { key: 'strike', label: 'Strikethrough', type: 'checkbox', value: !!st.strike, half: true },
          { key: 'vert', label: 'Effects', type: 'select', options: [{ value: '', label: 'Normal' }, { value: 'superscript', label: 'Superscript' }, { value: 'subscript', label: 'Subscript' }], value: st.vertAlign || '', half: true },
          { key: 'hb', type: 'heading', label: 'Border' },
          { key: 'bPreset', label: 'Presets', type: 'select', options: [{ value: '', label: '(individual edges below)' }, { value: 'none', label: 'None' }, { value: 'outside', label: 'Outline' }, { value: 'inside', label: 'Inside' }, { value: 'all', label: 'All' }], value: '' },
          { key: 'bt', label: 'Top', type: 'select', options: borderOpts, value: '', half: true },
          { key: 'bb', label: 'Bottom', type: 'select', options: borderOpts, value: '', half: true },
          { key: 'bl', label: 'Left', type: 'select', options: borderOpts, value: '', half: true },
          { key: 'br', label: 'Right', type: 'select', options: borderOpts, value: '', half: true },
          { key: 'bcolor', label: 'Border color', type: 'color', value: '#000000', half: true },
          { key: 'bstyle', label: 'Preset line style', type: 'select', options: BORDER_STYLES, value: 'thin', half: true },
          { key: 'hfill', type: 'heading', label: 'Fill' },
          { key: 'noFill', label: 'No fill', type: 'checkbox', value: !st.fill, half: true },
          { key: 'fill', label: 'Background color', type: 'color', value: st.fill || '#ffffff', half: true },
          { key: 'pattern', label: 'Pattern style', type: 'select', options: ['', 'solid', 'gray125', 'lightGray', 'darkGray', 'lightHorizontal', 'lightVertical', 'lightDown', 'lightUp', 'darkHorizontal', 'darkVertical', 'darkDown', 'darkUp'], value: st.pattern?.type || '', half: true },
          { key: 'patternColor', label: 'Pattern color', type: 'color', value: st.pattern?.color || '#000000', half: true },
          { key: 'hp', type: 'heading', label: 'Protection' },
          { key: 'locked', label: 'Locked', type: 'checkbox', value: st.locked !== false, half: true },
          { key: 'hidden', label: 'Hidden (formula not shown when the sheet is protected)', type: 'checkbox', value: !!st.hidden, half: true },
        ],
        onChange: (vals, ctl, key) => {
          if (key === 'cat') {
            const cat = FORMAT_CATEGORIES.find((x) => x.name === vals.cat);
            if (cat) ctl.set('fmt', cat.formats[0]);
          }
          const el = ctl.body.querySelector('[data-sample]');
          if (el) el.textContent = preview(vals.fmt || 'General');
        },
        onOpen: (ctl) => {
          if (page) ctl.body.querySelector(`[data-key="h${page === 'number' ? 'n' : page === 'alignment' ? 'a' : page === 'font' ? 'f' : page === 'border' ? 'b' : page === 'fill' ? 'fill' : 'p'}"]`)?.scrollIntoView();
        },
      });
      if (!res) return;
      const patch: Partial<CellStyle> = {
        numFmt: !res.fmt || res.fmt === 'General' ? undefined : res.fmt,
        hAlign: res.hAlign === 'general' ? undefined : res.hAlign, vAlign: res.vAlign === 'bottom' ? undefined : res.vAlign,
        indent: res.indent || undefined, rotate: res.rotate || undefined, wrap: res.wrap || undefined, shrink: res.shrink || undefined, rtl: res.rtl || undefined,
        font: res.font, size: res.size, bold: res.bold || undefined, italic: res.italic || undefined, underline: res.underline || undefined, color: res.color === '#000000' ? undefined : res.color,
        strike: res.strike || undefined, vertAlign: res.vert || undefined,
        fill: res.noFill ? undefined : res.fill, pattern: res.pattern && res.pattern !== 'solid' ? { type: res.pattern, color: res.patternColor } : undefined,
        locked: res.locked ? undefined : false, hidden: res.hidden || undefined,
      };
      if (res.fmt && res.fmt !== 'General' && !allFormats.includes(res.fmt)) { app.book.customFormats.push(res.fmt); }
      app.book.transact('Format Cells', () => {
        applyStyle(app, patch, 'Format Cells');
        const b = (x: string): Border | undefined | null => (!x ? null : x === 'none' ? undefined : { style: x as Border['style'], color: res.bcolor });
        if (res.bPreset) applyStyle(app, borderPatch(app, res.bPreset, res.bPreset === 'none' ? undefined : { style: res.bstyle, color: res.bcolor }), 'Borders');
        const edges: Partial<CellStyle> = {};
        for (const k of ['bt', 'bb', 'bl', 'br'] as const) { const v2 = b(res[k]); if (v2 !== null) (edges as any)[k] = v2; }
        if (Object.keys(edges).length) {
          applyStyle(app, borderPatch(app, 'outside', undefined), 'Borders');
          applyStyle(app, (_st, rr, cc) => {
            const rg = selRanges(app)[selRanges(app).length - 1];
            const p: Partial<CellStyle> = {};
            if ('bt' in edges && rr === rg.r1) p.bt = edges.bt; if ('bb' in edges && rr === rg.r2) p.bb = edges.bb;
            if ('bl' in edges && cc === rg.c1) p.bl = edges.bl; if ('br' in edges && cc === rg.c2) p.br = edges.br;
            return p;
          }, 'Borders');
        }
        if (res.merge !== !!app.sheet.mergeAt(r, c)) void merge(app, res.merge ? 'merge' : 'unmerge');
      });
    },
  },
  { id: 'fmt.clearFormats', label: 'Clear Formats', icon: 'format_clear', tab: 'Home', group: 'Font', edits: true, run: (app) => app.exec('edit.clearFormats') },
  {
    id: 'fmt.styleDialog', label: 'Style…', icon: 'style', tab: 'Home', group: 'Styles', keys: ["Alt-'"], edits: true,
    run: async (app) => {
      const names = [...BUILTIN_CELL_STYLES, ...app.book.cellStyles].map((s) => s.name);
      const r = await openDialog({ title: 'Style', fields: [{ key: 'name', label: 'Style name', type: 'combo', options: names, value: 'Normal' }], buttons: [{ label: 'New from Active Cell', onClick: (_v, ctl) => { ctl.close(); app.exec('fmt.newCellStyle'); } }] });
      if (r?.name) app.exec('fmt.cellStyles', r.name);
    },
  },
]);

export { THEME_COLORS, STANDARD_COLORS, h };
