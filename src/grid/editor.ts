/**
 * In-cell editor and formula assistance.
 *
 * One editing session is shared by the in-cell textarea and the formula bar, whichever has focus.
 * Modes: "enter" (started by typing; arrow keys commit and move, or point at cells while a formula
 * expects an operand), "edit" (F2/double-click; arrow keys move the caret), "point" (building a
 * reference by clicking or with the arrow keys). Formula references are color-highlighted on the grid,
 * and function names autocomplete with an argument tip.
 */
import type { Grid } from './grid';
import { h, clear } from '../ui/dom';
import { tokenize, type Token } from '../calc/parser';
import { colName, cellName, MAX_ROWS, MAX_COLS, normRange, type Range } from '../model/address';
import { FUNCTIONS, functionNames } from '../calc/functions';
import { formatValue, isDateFormat } from '../model/numfmt';
import { fontCss } from './render';
import { keyName } from '../util/keys';

const REF_COLORS = ['#2f6fb0', '#c0392b', '#7b3fa0', '#2e8b57', '#b8860b', '#008b8b', '#d2691e', '#c71585'];

export class CellEditor {
  active = false;
  r = 0;
  c = 0;
  mode: 'enter' | 'edit' | 'point' = 'enter';
  original = '';
  ta: HTMLTextAreaElement;
  bar: HTMLTextAreaElement | null = null;
  private popup: HTMLElement;
  private tip: HTMLElement;
  private acItems: { label: string; insert: string; desc: string; kind: string }[] = [];
  private acIndex = 0;
  /** Range inserted by the current point operation (replaced while pointing continues). */
  private pointSpan: { start: number; end: number } | null = null;
  private pointRange: Range | null = null;
  private pointSheet: string | null = null;
  startSheet = '';

  constructor(private g: Grid) {
    this.ta = h('textarea', { class: 'ql-cell-editor', spellcheck: 'false', autocapitalize: 'off', autocomplete: 'off', 'aria-label': 'Cell editor', rows: '1' }) as HTMLTextAreaElement;
    this.popup = h('div', { class: 'ql-ac-popup', role: 'listbox' });
    this.tip = h('div', { class: 'ql-arg-tip' });
    g.el.append(this.ta, this.popup, this.tip);
    this.bindInput(this.ta, 'cell');
  }

  private unbind = new AbortController();

  destroy(): void { this.unbind.abort(); this.ta.remove(); this.popup.remove(); this.tip.remove(); }

  /** The formula bar's input is attached by the shell. */
  attachBar(bar: HTMLTextAreaElement): void {
    this.bar = bar;
    this.bindInput(bar, 'bar');
  }

  get input(): HTMLTextAreaElement {
    return this.bar && document.activeElement === this.bar ? this.bar : this.ta;
  }

  get text(): string { return this.input.value; }

  focus(): void { (this.bar && this.barFocused ? this.bar : this.ta).focus({ preventScroll: true }); }
  private barFocused = false;

  // ------------------------------------------------------------------------------------------

  start(r: number, c: number, opts: { mode: 'enter' | 'edit'; text?: string; fromBar?: boolean }): void {
    (this as any).lastTyped = ''; // AutoComplete compares with what was typed in this edit, not the one before
    const g = this.g;
    const s = g.sheet;
    this.active = true;
    this.r = r; this.c = c;
    this.startSheet = s.name;
    this.mode = opts.mode;
    this.barFocused = !!opts.fromBar;
    const cell = s.get(r, c);
    this.original = editText(g, r, c);
    const text = opts.text != null ? opts.text : this.original;
    this.ta.value = text;
    if (this.bar) this.bar.value = text;
    this.ta.classList.add('ql-show');
    this.position();
    void cell;
    if (opts.fromBar) this.bar?.focus();
    else {
      this.ta.focus({ preventScroll: true });
      const end = this.ta.value.length;
      this.ta.setSelectionRange(end, end);
    }
    this.update();
    g.invalidate();
    g.app.refresh();
  }

  position(): void {
    if (!this.active) return;
    const g = this.g;
    if (g.sheet.name !== this.startSheet) { this.ta.classList.remove('ql-show'); return; }
    const rc = g.cellRect(this.r, this.c);
    const st = g.sheet.styleOf(this.r, this.c);
    const base = g.sheet.book.styles[0] || {};
    this.ta.style.font = fontCss(st, g.zoom, base);
    this.ta.style.left = `${Math.max(g.headerW, rc.x)}px`;
    this.ta.style.top = `${Math.max(g.headerH, rc.y)}px`;
    this.ta.style.minWidth = `${rc.w}px`;
    this.ta.style.minHeight = `${rc.h}px`;
    this.ta.style.maxWidth = `${Math.max(rc.w, g.width - rc.x - 8)}px`;
    this.ta.style.textAlign = st.hAlign === 'right' ? 'right' : st.hAlign === 'center' ? 'center' : 'left';
    this.ta.style.whiteSpace = st.wrap ? 'pre-wrap' : 'pre';
    this.ta.style.color = st.color || '';
    this.ta.style.background = st.fill || '';
    // grow to fit content
    this.ta.style.width = 'auto';
    this.ta.style.height = 'auto';
    this.ta.style.width = `${Math.min(Math.max(rc.w, this.ta.scrollWidth + 4), g.width - Math.max(g.headerW, rc.x) - 4)}px`;
    this.ta.style.height = `${Math.max(rc.h, this.ta.scrollHeight)}px`;
    this.ta.classList.add('ql-show');
  }

  canPoint(): boolean {
    if (!this.active) return false;
    const t = this.input.value;
    if (!t.startsWith('=') && !t.startsWith('+') && !(t.startsWith('-') && t.length > 1)) return false;
    if (this.pointSpan) return true;
    const pos = this.input.selectionStart ?? t.length;
    return expectsOperand(t.slice(0, pos));
  }

  /** Inserts or updates a reference at the caret (clicking or arrow keys while building a formula). */
  pointAt(rg: Range): void {
    const inp = this.input;
    const g = this.g;
    const t = inp.value;
    const sameSheet = g.sheet.name === this.startSheet;
    const prefix = sameSheet ? '' : quoteName(g.sheet.name) + '!';
    const ref = prefix + (rg.r1 === rg.r2 && rg.c1 === rg.c2 ? cellName(rg.r1, rg.c1) : `${cellName(rg.r1, rg.c1)}:${cellName(rg.r2, rg.c2)}`);
    let start: number, end: number;
    if (this.pointSpan) ({ start, end } = this.pointSpan);
    else { start = inp.selectionStart ?? t.length; end = inp.selectionEnd ?? start; }
    inp.value = t.slice(0, start) + ref + t.slice(end);
    this.pointSpan = { start, end: start + ref.length };
    this.pointRange = rg;
    this.pointSheet = g.sheet.name;
    inp.setSelectionRange(start + ref.length, start + ref.length);
    this.mode = this.mode === 'edit' ? 'edit' : 'point';
    g.pointRange = rg;
    this.sync(inp);
    this.update();
  }

  private stopPointing(): void {
    this.pointSpan = null;
    this.pointRange = null;
    this.g.pointRange = null;
    if (this.mode === 'point') this.mode = 'enter';
  }

  // ------------------------------------------------------------------------------------------

  private bindInput(inp: HTMLTextAreaElement, which: 'cell' | 'bar'): void {
    // Every workbook's editor listens to the one shared formula bar: only the editor of the
    // workbook in front may answer, or an entry is written into each open workbook's active cell.
    const mine = () => which === 'cell' || this.g.app.active?.grid === this.g;
    const o = { signal: this.unbind.signal };
    inp.addEventListener('input', () => {
      if (!mine()) return;
      if (!this.active) { this.g.startEdit({ mode: 'edit', fromBar: which === 'bar', text: inp.value }); return; }
      this.stopPointing();
      this.sync(inp);
      this.autoCompleteValue(inp);
      this.update();
      this.position();
    }, o);
    inp.addEventListener('keydown', (e) => { if (mine()) this.onKey(e, inp, which); }, o);
    inp.addEventListener('focus', () => { if (!mine()) return; this.barFocused = which === 'bar'; if (which === 'bar' && !this.active) this.g.startEdit({ mode: 'edit', fromBar: true }); }, o);
    inp.addEventListener('click', () => { if (mine() && this.active) { this.stopPointing(); if (this.mode === 'enter') this.mode = 'edit'; this.update(); } }, o);
    inp.addEventListener('blur', () => setTimeout(() => this.hidePopup(), 150), o);
  }

  private sync(from: HTMLTextAreaElement): void {
    const other = from === this.ta ? this.bar : this.ta;
    if (other && other.value !== from.value) other.value = from.value;
  }

  private onKey(e: KeyboardEvent, inp: HTMLTextAreaElement, which: 'cell' | 'bar'): void {
    if (!this.active) return;
    // Keys pressed while an input method is composing belong to it (Enter picks a candidate).
    if (e.isComposing || e.keyCode === 229) return;
    const g = this.g;
    // Autocomplete list navigation
    if (this.popup.classList.contains('ql-show')) {
      if (e.key === 'ArrowDown' || e.key === 'ArrowUp') { e.preventDefault(); this.acIndex = (this.acIndex + (e.key === 'ArrowDown' ? 1 : -1) + this.acItems.length) % this.acItems.length; this.acMoved = true; this.renderPopup(); return; }
      if (e.key === 'Tab' || (e.key === 'Enter' && this.acItems.length && this.acTyped())) { e.preventDefault(); this.acceptAc(); return; }
      if (e.key === 'Escape') { e.preventDefault(); this.hidePopup(); return; }
    }
    const k = keyName(e);
    if (e.key === 'Enter') {
      if (e.altKey || (e.ctrlKey && e.metaKey)) { e.preventDefault(); insertAtCaret(inp, '\n'); this.sync(inp); this.position(); return; }
      e.preventDefault();
      if (e.ctrlKey && e.shiftKey) { this.commit('none', { array: true }); return; }
      if (e.ctrlKey || e.metaKey) { this.commit('none', { fill: true }); return; }
      this.commit(e.shiftKey ? 'up' : 'enter');
      return;
    }
    if (e.key === 'Tab') { e.preventDefault(); this.commit(e.shiftKey ? 'left' : 'right'); return; }
    if (e.key === 'Escape') { e.preventDefault(); this.cancel(); return; }
    if (e.key === 'F4' || (k === 'Mod-t' && /mac/i.test(navigator.platform))) { e.preventDefault(); cycleAbsolute(inp); this.sync(inp); this.update(); return; }
    if (e.key === 'F2') { e.preventDefault(); this.mode = this.mode === 'edit' ? 'enter' : 'edit'; this.stopPointing(); this.g.app.ui.status(this.mode === 'edit' ? 'Edit' : 'Enter'); return; }
    if (e.key === 'F9' && !e.shiftKey && !e.ctrlKey) {
      // Evaluate the selected part of a formula in place.
      const a = inp.selectionStart ?? 0, b = inp.selectionEnd ?? 0;
      if (b > a && inp.value.startsWith('=')) {
        e.preventDefault();
        const part = inp.value.slice(a, b);
        const v = g.tab.engine.evaluateScalar(part, g.sheet, this.r, this.c);
        const txt = typeof v === 'string' ? `"${v.replace(/"/g, '""')}"` : v == null ? '0' : String(v);
        inp.setRangeText(txt, a, b, 'select');
        this.sync(inp);
      }
      return;
    }
    if (e.key.startsWith('Arrow') && !e.altKey) {
      const dr = e.key === 'ArrowDown' ? 1 : e.key === 'ArrowUp' ? -1 : 0;
      const dc = e.key === 'ArrowRight' ? 1 : e.key === 'ArrowLeft' ? -1 : 0;
      if ((this.mode === 'enter' || this.mode === 'point') && this.canPoint()) {
        e.preventDefault();
        const base = this.pointRange || { r1: this.r, c1: this.c, r2: this.r, c2: this.c };
        let rg: Range;
        if (e.shiftKey && this.pointRange) rg = normRange({ r1: base.r1, c1: base.c1, r2: clamp(base.r2 + dr, 0, MAX_ROWS - 1), c2: clamp(base.c2 + dc, 0, MAX_COLS - 1) });
        else { const r = clamp((this.pointRange ? base.r2 : this.r) + dr, 0, MAX_ROWS - 1), c = clamp((this.pointRange ? base.c2 : this.c) + dc, 0, MAX_COLS - 1); rg = { r1: r, c1: c, r2: r, c2: c }; }
        this.pointAt(rg);
        g.scrollIntoView(rg.r2, rg.c2);
        return;
      }
      if (this.mode === 'enter' && which === 'cell' && !e.shiftKey && !(e.ctrlKey || e.metaKey)) {
        e.preventDefault();
        this.commit(dr > 0 ? 'down' : dr < 0 ? 'up' : dc > 0 ? 'right' : 'left');
        return;
      }
    }
    // Ctrl+; date, Ctrl+Shift+; time while editing
    if (k === 'Mod-;') { e.preventDefault(); const d = new Date(); insertAtCaret(inp, `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`); this.sync(inp); return; }
    if (k === 'Mod-Shift-;' || k === 'Mod-:') { e.preventDefault(); const d = new Date(); insertAtCaret(inp, `${d.getHours()}:${String(d.getMinutes()).padStart(2, '0')}`); this.sync(inp); return; }
    if (k === 'Mod-a' && inp.value.startsWith('=')) {
      // Ctrl+A after a function name opens the Function Arguments dialog
      const fn = enclosingFunction(inp.value, inp.selectionStart ?? 0);
      if (fn) { e.preventDefault(); g.app.exec('formulas.insertFunction', { name: fn.name, fromEditor: true }); return; }
    }
    if (k === 'Shift-F3') { e.preventDefault(); g.app.exec('formulas.insertFunction', { fromEditor: true }); return; }
    if (k === 'F3') { e.preventDefault(); g.app.exec('formulas.pasteName'); return; }
  }

  /** Enter takes a suggestion only after the arrows chose one; otherwise it enters the cell as typed (Tab always takes it). */
  private acMoved = false;
  private acTyped(): boolean { return this.acMoved; }

  /** Commits the edit. Returns false when validation rejected the value (editing continues). */
  commit(move: 'down' | 'up' | 'right' | 'left' | 'none' | 'enter', opts: { array?: boolean; fill?: boolean } = {}): boolean {
    if (!this.active) return true;
    const g = this.g;
    let text = this.input.value;
    // Return to the sheet where editing started
    if (g.sheet.name !== this.startSheet) {
      const idx = g.sheet.book.sheets.findIndex((s) => s.name === this.startSheet);
      if (idx >= 0) g.sheet.book.activeSheet = idx;
    }
    if (text.startsWith('=')) text = closeParens(text);
    const changed = text !== this.original || opts.array || opts.fill;
    this.endUi();
    if (changed) {
      const ok = g.app.exec('cell.enter', { r: this.r, c: this.c, text, array: opts.array, fill: opts.fill });
      if (ok === false) {
        // keep editing
        this.active = true;
        this.ta.value = text;
        if (this.bar) this.bar.value = text;
        this.ta.classList.add('ql-show');
        this.position();
        this.ta.focus();
        return false;
      }
    }
    this.active = false;
    g.refHighlights = [];
    g.pointRange = null;
    if (move === 'enter') g.cycleInSelection(1, g.app.prefs.enterDirection);
    else if (move === 'down') g.cycleInSelection(1, 'down');
    else if (move === 'up') g.cycleInSelection(-1, 'down');
    else if (move === 'right') g.cycleInSelection(1, 'right');
    else if (move === 'left') g.cycleInSelection(-1, 'right');
    g.invalidate();
    g.app.refresh();
    g.focus();
    return true;
  }

  cancel(): void {
    if (!this.active) return;
    this.active = false;
    this.endUi();
    const g = this.g;
    if (g.sheet.name !== this.startSheet) {
      const idx = g.sheet.book.sheets.findIndex((s) => s.name === this.startSheet);
      if (idx >= 0) g.sheet.book.activeSheet = idx;
    }
    if (this.bar) this.bar.value = this.original;
    g.refHighlights = [];
    g.pointRange = null;
    g.invalidate();
    g.app.refresh();
    g.focus();
  }

  private endUi(): void {
    this.ta.classList.remove('ql-show');
    this.ta.value = '';
    this.hidePopup();
    this.tip.classList.remove('ql-show');
    this.pointSpan = null;
    this.pointRange = null;
  }

  // ------------------------------------------------------------------------------------------
  // Formula assistance

  private update(): void {
    const g = this.g;
    const t = this.input.value;
    g.refHighlights = [];
    if (t.startsWith('=')) {
      let toks: Token[] = [];
      try { toks = tokenize(t.slice(1)); } catch { /* incomplete formula */ }
      const colors = new Map<string, string>();
      let i = 0;
      for (const tk of toks) {
        if (tk.type !== 'ref' || !tk.ref || tk.ref.invalid) continue;
        const key = tk.text.replace(/\$/g, '').toUpperCase();
        if (!colors.has(key)) colors.set(key, REF_COLORS[i++ % REF_COLORS.length]);
        g.refHighlights.push({ sheet: tk.ref.sheet || this.startSheet, range: { r1: tk.ref.r1, c1: tk.ref.c1, r2: tk.ref.r2, c2: tk.ref.c2 }, color: colors.get(key)! });
      }
      this.functionAssist(t);
    } else {
      this.hidePopup();
      this.tip.classList.remove('ql-show');
    }
    g.invalidate();
  }

  private functionAssist(t: string): void {
    const inp = this.input;
    const pos = inp.selectionStart ?? t.length;
    const before = t.slice(0, pos);
    const m = /([A-Za-z_][A-Za-z0-9_.]*)$/.exec(before);
    const g = this.g;
    if (m && !/\d$/.test(m[1].replace(/[A-Za-z_.]/g, '')) && !inString(before) && m[1].length >= 1 && !/^[A-Z]{1,3}\d+$/i.test(m[1])) {
      const prefix = m[1].toUpperCase();
      const items: CellEditor['acItems'] = [];
      for (const name of functionNames()) if (name.startsWith(prefix)) { const f = FUNCTIONS.get(name)!; items.push({ label: name, insert: name + '(', desc: f.desc, kind: 'fn' }); }
      for (const n of g.sheet.book.names) if (n.name.toUpperCase().startsWith(prefix) && (!n.sheet || n.sheet === g.sheet.name)) items.push({ label: n.name, insert: n.name, desc: n.comment || n.ref, kind: 'name' });
      for (const sh of g.sheet.book.sheets) for (const tb of sh.tables) if (tb.name.toUpperCase().startsWith(prefix)) items.push({ label: tb.name, insert: tb.name + '[', desc: 'Table', kind: 'table' });
      if (items.length && !(items.length === 1 && items[0].label.toUpperCase() === prefix)) {
        this.acItems = items.slice(0, 40);
        this.acIndex = 0;
        this.acMoved = false;
        this.renderPopup();
      } else this.hidePopup();
    } else this.hidePopup();
    // Argument tip
    const fn = enclosingFunction(t, pos);
    const def = fn && FUNCTIONS.get(fn.name);
    if (def && g.app.prefs.showFunctionTips) {
      clear(this.tip);
      this.tip.append(h('b', null, def.name), '(');
      def.params.forEach((p, i) => {
        if (i) this.tip.append(', ');
        const isCur = i === fn!.arg || (p === '...' && fn!.arg >= i) || (i === def.params.length - 1 && fn!.arg > i && /\.\.\./.test(def.params.join(',')));
        this.tip.append(isCur ? h('b', { class: 'ql-arg-cur' }, p) : p);
      });
      this.tip.append(')');
      this.placeBelow(this.tip, true);
    } else this.tip.classList.remove('ql-show');
  }

  private renderPopup(): void {
    clear(this.popup);
    this.acItems.forEach((it, i) => {
      const row = h('div', { class: 'ql-ac-item' + (i === this.acIndex ? ' ql-sel' : ''), role: 'option' },
        h('span', { class: 'ql-icon' }, it.kind === 'fn' ? 'function' : it.kind === 'table' ? 'table' : 'label'), h('b', null, it.label));
      row.title = it.desc;
      row.addEventListener('mousedown', (e) => { e.preventDefault(); this.acIndex = i; this.acceptAc(); });
      this.popup.appendChild(row);
    });
    const cur = this.acItems[this.acIndex];
    if (cur) this.popup.appendChild(h('div', { class: 'ql-ac-desc' }, cur.desc));
    this.placeBelow(this.popup, false);
    (this.popup.children[this.acIndex] as HTMLElement | undefined)?.scrollIntoView({ block: 'nearest' });
  }

  /**
   * Shows the autocomplete list or the argument tip under the text being edited, or above it when the cell is
   * near the bottom of the visible grid (the end of the sheet area, or the on-screen keyboard) and there is
   * more room above. The list is kept inside the grid horizontally and shortened to the room it has.
   */
  private placeBelow(el: HTMLElement, tipRow: boolean): void {
    const g = this.g;
    const inBar = this.bar && document.activeElement === this.bar;
    const base = inBar ? this.bar!.getBoundingClientRect() : this.ta.getBoundingClientRect();
    const host = g.el.getBoundingClientRect();
    el.style.maxHeight = '';
    el.classList.add('ql-show');
    const vv = window.visualViewport;
    const visTop = Math.max(host.top, vv ? vv.offsetTop : 0), visBottom = Math.min(host.bottom, vv ? vv.offsetTop + vv.height : innerHeight);
    const gap = tipRow ? 2 : this.tip.classList.contains('ql-show') && !this.tip.classList.contains('ql-above') ? 26 : 2;
    const tipAbove = !tipRow && this.tip.classList.contains('ql-show') && this.tip.classList.contains('ql-above') ? this.tip.offsetHeight + 2 : 0;
    const below = visBottom - (base.bottom + gap) - 4, above = base.top - visTop - 6 - tipAbove;
    const hgt = el.offsetHeight;
    const up = hgt > below && above > below && !inBar;
    const room = Math.max(40, up ? above : below);
    if (hgt > room) el.style.maxHeight = `${room}px`;
    const shownH = Math.min(hgt, room);
    el.classList.toggle('ql-above', up);
    el.style.top = `${(up ? base.top - 2 - tipAbove - shownH : base.bottom + gap) - host.top}px`;
    el.style.left = `${Math.max(0, Math.min(base.left - host.left, host.width - el.offsetWidth - 2))}px`;
  }

  private hidePopup(): void { this.popup.classList.remove('ql-show'); this.acItems = []; }

  private acceptAc(): void {
    const it = this.acItems[this.acIndex];
    if (!it) return;
    const inp = this.input;
    const pos = inp.selectionStart ?? inp.value.length;
    const before = inp.value.slice(0, pos);
    const m = /([A-Za-z_][A-Za-z0-9_.]*)$/.exec(before);
    const start = m ? pos - m[1].length : pos;
    inp.setRangeText(it.insert, start, pos, 'end');
    this.sync(inp);
    this.hidePopup();
    this.update();
    inp.focus();
  }

  /** AutoComplete for plain text: completes from values already in the column. */
  private autoCompleteValue(inp: HTMLTextAreaElement): void {
    const g = this.g;
    if (!g.app.prefs.autoComplete || inp.value.startsWith('=') || !inp.value || (inp.selectionStart ?? 0) !== inp.value.length) return;
    const typed = inp.value;
    if ((this as any).lastTyped && (this as any).lastTyped.length > typed.length) { (this as any).lastTyped = typed; return; } // backspacing
    (this as any).lastTyped = typed;
    if (/^[\d.,\-+$%]/.test(typed)) return;
    const s = g.sheet;
    const low = typed.toLowerCase();
    const found = new Set<string>();
    for (const dir of [-1, 1]) {
      for (let r = this.r + dir, n = 0; r >= 0 && r < MAX_ROWS && n < 2000; r += dir, n++) {
        const v = s.value(r, this.c);
        if (v == null || v === '') break;
        if (typeof v === 'string' && v.toLowerCase().startsWith(low) && v.length > typed.length) found.add(v);
      }
    }
    if (found.size !== 1) return;
    const [match] = [...found];
    inp.value = typed + match.slice(typed.length);
    inp.setSelectionRange(typed.length, match.length);
    this.sync(inp);
  }
}

// ----------------------------------------------------------------------------------------------
// Helpers

/** Text shown when editing a cell: the formula, or the value in an editable form. */
export function editText(g: Grid, r: number, c: number): string {
  const s = g.sheet;
  const cell = s.get(r, c);
  if (!cell) return '';
  if (cell.f != null) return (cell.arr ? '{=' : '=') + cell.f + (cell.arr ? '}' : '');
  const v = cell.v;
  if (v == null) return '';
  if (typeof v === 'string') {
    // A leading apostrophe keeps text that would otherwise parse as a number or formula.
    return /^[=+\-]|^\d|^(TRUE|FALSE)$/i.test(v) && !isTextFmt(g, r, c) ? "'" + v : v;
  }
  if (typeof v === 'number') {
    const fmt = s.styleOf(r, c).numFmt;
    if (isDateFormat(fmt)) {
      const hasTime = /[hs]/i.test(fmt!.replace(/"[^"]*"/g, ''));
      const hasDate = /[dy]|m{3,}/i.test(fmt!.replace(/"[^"]*"/g, '')) || /m\/|\/m|d/.test(fmt!);
      const f = hasDate && hasTime ? 'm/d/yyyy h:mm:ss' : hasTime ? 'h:mm:ss' : 'm/d/yyyy';
      return formatValue(v, f, { date1904: s.book.calc.date1904 }).text.replace(/:00$/, (m0) => (hasTime && !/s/.test(fmt!) ? '' : m0)).replace(/ 0:00$/, '');
    }
    if (fmt && /%/.test(fmt) && !/"[^"]*%/.test(fmt)) return String(+(v * 100).toPrecision(15)) + '%';
    return String(+v.toPrecision(15));
  }
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  return v.code;
}

function isTextFmt(g: Grid, r: number, c: number): boolean { return g.sheet.styleOf(r, c).numFmt === '@'; }

function clamp(v: number, a: number, b: number): number { return Math.max(a, Math.min(b, v)); }

function quoteName(n: string): string { return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(n) ? n : `'${n.replace(/'/g, "''")}'`; }

function insertAtCaret(inp: HTMLTextAreaElement, text: string): void {
  const a = inp.selectionStart ?? inp.value.length, b = inp.selectionEnd ?? a;
  inp.setRangeText(text, a, b, 'end');
}

function inString(s: string): boolean { return ((s.match(/"/g) || []).length % 2) === 1; }

/** True when the formula text up to the caret ends where an operand (reference) can be inserted. */
export function expectsOperand(before: string): boolean {
  const t = before.replace(/\s+$/, '');
  if (inString(t)) return false;
  return /[=(,+\-*/^&<>:;{%]$/.test(t) || /^[=+\-]$/.test(t);
}

/** Function call enclosing the caret and the 0-based argument index. */
export function enclosingFunction(t: string, pos: number): { name: string; arg: number } | null {
  let depth = 0, arg = 0;
  let inStr = false;
  for (let i = pos - 1; i >= 0; i--) {
    const ch = t[i];
    if (ch === '"') { inStr = !inStr; continue; }
    if (inStr) continue;
    if (ch === ')') depth++;
    else if (ch === '(') {
      if (depth === 0) {
        const m = /([A-Za-z_][A-Za-z0-9_.]*)\s*$/.exec(t.slice(0, i));
        return m ? { name: m[1].toUpperCase(), arg } : null;
      }
      depth--;
    } else if ((ch === ',') && depth === 0) arg++;
  }
  return null;
}

/** F4: cycles A1 → $A$1 → A$1 → $A1 → A1 for the reference at or before the caret. */
export function cycleAbsolute(inp: HTMLTextAreaElement): void {
  const t = inp.value;
  const pos = inp.selectionStart ?? t.length;
  const re = /(\$?[A-Za-z]{1,3}\$?\d{1,7})(:\$?[A-Za-z]{1,3}\$?\d{1,7})?/g;
  let m: RegExpExecArray | null;
  let hit: RegExpExecArray | null = null;
  while ((m = re.exec(t))) { if (m.index <= pos && re.lastIndex >= pos - 1) hit = m; if (m.index > pos) break; }
  if (!hit) return;
  const cycle = (ref: string) => {
    const mm = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})$/.exec(ref)!;
    const ac = !!mm[1], ar = !!mm[3];
    const [nc, nr] = !ac && !ar ? [true, true] : ac && ar ? [false, true] : !ac && ar ? [true, false] : [false, false];
    return `${nc ? '$' : ''}${mm[2]}${nr ? '$' : ''}${mm[4]}`;
  };
  const parts = hit[0].split(':').map(cycle);
  const repl = parts.join(':');
  inp.setRangeText(repl, hit.index, hit.index + hit[0].length, 'end');
}

/** Adds missing closing parentheses on commit, as spreadsheets do. */
export function closeParens(t: string): string {
  let depth = 0, inStr = false;
  for (const ch of t) { if (ch === '"') inStr = !inStr; else if (!inStr) { if (ch === '(') depth++; else if (ch === ')') depth--; } }
  if (inStr) t += '"';
  return depth > 0 ? t + ')'.repeat(depth) : t;
}

export { colName };
