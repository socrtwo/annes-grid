/** Application shell: title bar, Quick Access Bar, tabs + ribbon, formula bar, grid area, sheet tabs, panes, status bar. */
import type { App, BookTab } from '../app';
import { getCommand, allCommands, searchCommands, type Command, type MenuEntry } from '../commands/registry';
import { h, icon, clear, isPhone } from './dom';
import { showMenu, closeMenus, openDialog, toast } from './dialog';
import { DEFAULT_TABS, tabsFor, type RibbonTab } from './ribbonLayout';
import { prettyKey, canonicalKey, keyName } from '../util/keys';
import { PANES, type PaneDef } from './panes';
import { customizeQat, customizeRibbon, customizeKeyboard } from './customize';
import { renderBackstage } from './backstage';
import { Grid, type Hit } from '../grid/grid';
import { cellName, colName, MAX_COLS, MAX_ROWS, rangeName, keyRow, keyCol } from '../model/address';
import { CellError } from '../model/types';
import { formatValue } from '../model/numfmt';
import { FONT_LIST } from '../fonts';
import { activeStyle } from '../commands/ops';
import { editText } from '../grid/editor';
import { QAT_STANDARD, QAT_EXTRAS } from '../prefs';

export class Shell {
  root: HTMLElement;
  titlebar!: HTMLElement;
  qatEl!: HTMLElement;
  qatBottom!: HTMLElement;
  tabsEl!: HTMLElement;
  ribbonEl!: HTMLElement;
  docTabsEl!: HTMLElement;
  formulaBar!: HTMLElement;
  nameBox!: HTMLInputElement;
  barInput!: HTMLTextAreaElement;
  workspace!: HTMLElement;
  paneEl!: HTMLElement;
  sheetBar!: HTMLElement;
  statusEl!: HTMLElement;
  backstageEl!: HTMLElement;
  titleBtn!: HTMLElement;
  currentTab = 'home';
  ribbonOpen = true;
  pane: { def: PaneDef; el: HTMLElement; opts: any } | null = null;
  statusMsg = '';
  private statusTimer: any;
  private statsKey = '';
  private statsCache: Record<string, number | null> = {};

  constructor(public app: App) {
    this.root = app.host;
    this.build();
    this.applyPrefs();
    document.addEventListener('keydown', (e) => this.globalKey(e), true);
  }

  // ----------------------------------------------------------------------------------------
  // Layout

  build(): void {
    clear(this.root);
    this.root.className = 'ql-app';
    this.titlebar = h('header', { class: 'ql-titlebar' });
    const fileBtn = h('button', { class: 'ql-iconbtn ql-filebtn', title: 'File', 'aria-label': 'File menu', onclick: () => this.showBackstage() }, icon('menu'));
    this.qatEl = h('div', { class: 'ql-qat', role: 'toolbar', 'aria-label': 'Quick Access Bar' });
    this.titleBtn = h('button', { class: 'ql-title', title: 'Rename / workbook info', onclick: () => this.renameDoc() }, 'Book');
    const search = h('button', { class: 'ql-iconbtn', title: 'Search commands (Alt+Q)', 'aria-label': 'Search commands', onclick: () => this.commandPalette() }, icon('search'));
    const more = h('button', { class: 'ql-iconbtn', title: 'More', 'aria-label': 'More', onclick: (e: MouseEvent) => this.moreMenu(e.currentTarget as HTMLElement) }, icon('more_vert'));
    // Shown only while the Quick Access Bar is hidden, in its place, so there is always a visible way to get it back.
    const qatBack = h('button', { class: 'ql-iconbtn ql-qat-restore', title: 'Show Quick Access Bar', 'aria-label': 'Show Quick Access Bar', onclick: () => { this.app.prefs.qatHidden = false; this.app.savePrefs(); this.applyPrefs(); toast('Quick Access Bar shown'); } }, icon('bolt'));
    this.titlebar.append(fileBtn, this.qatEl, qatBack, this.titleBtn, h('span', { class: 'ql-spacer' }), search, more);
    this.tabsEl = h('nav', { class: 'ql-tabs', role: 'tablist' });
    this.ribbonEl = h('div', { class: 'ql-ribbon', role: 'toolbar' });
    this.docTabsEl = h('div', { class: 'ql-doctabs' });
    this.buildFormulaBar();
    this.workspace = h('div', { class: 'ql-workspace' });
    this.paneEl = h('aside', { class: 'ql-pane', 'aria-hidden': 'true' });
    const main = h('main', { class: 'ql-main' }, this.workspace, this.paneEl);
    this.sheetBar = h('div', { class: 'ql-sheetbar', role: 'tablist', 'aria-label': 'Sheets' });
    this.qatBottom = h('div', { class: 'ql-qat ql-qat-bottom', role: 'toolbar' });
    this.statusEl = h('footer', { class: 'ql-status' });
    this.backstageEl = h('div', { class: 'ql-backstage', 'aria-hidden': 'true' });
    this.root.append(this.titlebar, this.tabsEl, this.ribbonEl, this.docTabsEl, this.formulaBar, main, this.sheetBar, this.qatBottom, this.statusEl, this.backstageEl, h('div', { class: 'ql-toasts' }));
    this.renderQat();
    this.renderTabs();
    this.renderRibbon();
  }

  private buildFormulaBar(): void {
    this.nameBox = h('input', { class: 'ql-namebox', 'aria-label': 'Name Box', title: 'Name Box: type a cell, range or name and press Enter', spellcheck: 'false', autocomplete: 'off' }) as HTMLInputElement;
    this.nameBox.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') { e.preventDefault(); const v = this.nameBox.value.trim(); if (v) this.app.exec('nav.goto', v); this.app.focus(); }
      if (e.key === 'Escape') { this.refresh(); this.app.focus(); }
    });
    this.nameBox.addEventListener('focus', () => this.nameBox.select());
    const nameDrop = h('button', { class: 'ql-iconbtn ql-sm', title: 'Defined names', 'aria-label': 'Defined names', onclick: (e: MouseEvent) => {
      const names = this.app.book.names.filter((n) => !n.hidden);
      const tables = this.app.book.sheets.flatMap((s) => s.tables.map((t) => t.name));
      showMenu([...names.map((n) => ({ id: 'nav.goto', arg: n.name, label: n.name })), ...tables.map((t) => ({ id: 'nav.goto', arg: t, label: t, icon: 'table' })), ...(names.length + tables.length ? [] : [{ label: '(no defined names)' }]), { separator: true }, { id: 'formulas.nameManager', label: 'Name Manager…', icon: 'badge' }], { anchor: e.currentTarget as HTMLElement, onPick: (en) => en.id && this.app.exec(en.id, en.arg) });
    } }, icon('arrow_drop_down'));
    const cancel = h('button', { class: 'ql-iconbtn ql-fb-btn', title: 'Cancel (Esc)', 'aria-label': 'Cancel', onclick: () => this.app.grid.cancelEdit() }, icon('close'));
    const enter = h('button', { class: 'ql-iconbtn ql-fb-btn', title: 'Enter', 'aria-label': 'Enter', onclick: () => this.app.grid.commitEdit('none') }, icon('check'));
    const fx = h('button', { class: 'ql-iconbtn ql-fx', title: 'Insert Function (Shift+F3)', 'aria-label': 'Insert Function', onclick: () => this.app.exec('formulas.insertFunction') }, h('i', null, 'fx'));
    this.barInput = h('textarea', { class: 'ql-fb-input', rows: '1', spellcheck: 'false', autocomplete: 'off', 'aria-label': 'Formula bar' }) as HTMLTextAreaElement;
    const expand = h('button', { class: 'ql-iconbtn ql-sm', title: 'Expand formula bar', 'aria-label': 'Expand formula bar', onclick: () => { this.formulaBar.classList.toggle('ql-fb-expanded'); } }, icon('expand_more'));
    this.formulaBar = h('div', { class: 'ql-formulabar' }, h('div', { class: 'ql-namebox-wrap' }, this.nameBox, nameDrop), cancel, enter, fx, this.barInput, expand);
  }

  rebuild(): void {
    this.renderQat();
    this.renderTabs();
    this.renderRibbon();
    this.refresh();
  }

  mountTab(tab: BookTab): void {
    tab.el = h('div', { class: 'ql-book' });
    tab.grid = new Grid(this.app, tab);
    tab.grid.editor.attachBar(this.barInput);
    tab.el.appendChild(tab.grid.el);
    this.workspace.appendChild(tab.el);
    requestAnimationFrame(() => tab.grid.layout());
  }

  // ----------------------------------------------------------------------------------------
  // Quick Access Bar

  renderQat(): void {
    const target = this.app.prefs.qatPosition === 'bottom' ? this.qatBottom : this.qatEl;
    clear(this.qatEl); clear(this.qatBottom);
    for (const id of this.app.prefs.qat) {
      if (id === '|') { target.appendChild(h('span', { class: 'ql-qat-sep' })); continue; }
      const c = getCommand(id);
      if (!c) continue;
      target.appendChild(this.commandButton(c, 'qat'));
    }
    target.appendChild(h('button', { class: 'ql-iconbtn ql-qat-customize', title: 'Customize Quick Access Bar', 'aria-label': 'Customize Quick Access Bar', onclick: (e: MouseEvent) => this.qatMenu(e.currentTarget as HTMLElement) }, icon('arrow_drop_down')));
  }

  qatMenu(anchor: HTMLElement): void {
    const q = this.app.prefs.qat;
    const item = (id: string) => ({ id, label: getCommand(id)?.label || id, checked: q.includes(id), arg: '__toggle' });
    const present = (id: string) => !!getCommand(id);
    showMenu([
      { heading: 'Customize Quick Access Bar' },
      ...QAT_STANDARD.filter(present).map(item),
      { separator: true },
      ...QAT_EXTRAS.filter(present).map(item),
      { separator: true },
      { id: '__addAll', label: 'Add All Choices (Save, Print, then A–Z)', icon: 'sort_by_alpha' },
      { id: '__more', label: 'More Commands…', icon: 'library_add' },
      { id: '__pos', label: this.app.prefs.qatPosition === 'bottom' ? 'Show Above the Ribbon' : 'Show at the Bottom (near keyboard)', icon: 'swap_vert' },
      { id: '__labels', label: this.app.prefs.qatLabels ? 'Hide Command Labels' : 'Show Command Labels', icon: 'label' },
      { id: '__hide', label: 'Hide Quick Access Bar', icon: 'visibility_off' },
      { id: '__ribbon', label: this.app.prefs.ribbon === 'full' ? 'Collapse the Ribbon' : 'Show the Ribbon', icon: 'expand_less' },
    ], {
      anchor, onPick: (e) => {
        const p = this.app.prefs;
        if (e.arg === '__toggle' && e.id) { p.qat = q.includes(e.id) ? q.filter((x) => x !== e.id) : [...q, e.id]; }
        else if (e.id === '__addAll') { this.app.exec('view.qatAddAll'); return; }
        else if (e.id === '__more') { this.customizeQat(); return; }
        else if (e.id === '__pos') { this.app.exec('view.qatAtBottom'); return; }
        else if (e.id === '__labels') p.qatLabels = !p.qatLabels;
        else if (e.id === '__hide') { p.qatHidden = true; toast('Quick Access Bar hidden. The ⚡ button in the title bar shows it again.', { ms: 8000, action: 'Undo', onAction: () => { p.qatHidden = false; this.app.savePrefs(); this.applyPrefs(); } }); }
        else if (e.id === '__ribbon') p.ribbon = p.ribbon === 'full' ? 'tabs' : 'full';
        this.app.savePrefs();
        this.applyPrefs();
        this.renderQat();
        this.refresh();
      },
    });
  }

  // ----------------------------------------------------------------------------------------
  // Tabs and ribbon

  renderTabs(): void {
    clear(this.tabsEl);
    this.tabsEl.appendChild(h('button', { class: 'ql-tab ql-tab-file', role: 'tab', onclick: () => this.showBackstage() }, 'File'));
    for (const t of tabsFor(this.app)) {
      const b = h('button', { class: 'ql-tab' + (t.contextual ? ' ql-contextual' : '') + (t.id === this.currentTab ? ' ql-active' : ''), role: 'tab', 'data-tab': t.id, style: t.accent ? `--tab-accent:${t.accent}` : '' }, t.label);
      b.addEventListener('click', () => {
        if (this.currentTab === t.id && this.app.prefs.ribbon !== 'hidden') { this.ribbonOpen = !this.ribbonOpen; this.root.classList.toggle('ql-ribbon-closed', !this.ribbonOpen); return; }
        this.currentTab = t.id; this.ribbonOpen = true; this.root.classList.remove('ql-ribbon-closed');
        this.renderTabs(); this.renderRibbon(); this.refresh();
        b.scrollIntoView({ inline: 'nearest', block: 'nearest' });
      });
      b.addEventListener('dblclick', () => { this.app.exec('view.collapseRibbon'); });
      this.tabsEl.appendChild(b);
    }
  }

  ribbonTab(): RibbonTab {
    return tabsFor(this.app).find((t) => t.id === this.currentTab) || DEFAULT_TABS[0];
  }

  renderRibbon(): void {
    clear(this.ribbonEl);
    const tab = this.ribbonTab();
    for (const g of tab.groups) {
      const items = h('div', { class: 'ql-group-items' });
      for (const id of g.items) {
        const c = getCommand(id);
        if (!c) continue;
        items.appendChild(this.ribbonItem(c));
      }
      const label = h('div', { class: 'ql-group-label' }, g.label, g.launcher ? h('button', { class: 'ql-launcher', title: `${g.label} settings`, 'aria-label': `${g.label} settings`, onclick: () => this.app.exec(g.launcher!) }, icon('open_in_new')) : null);
      this.ribbonEl.appendChild(h('section', { class: 'ql-group', 'data-group': g.id }, items, label));
    }
    this.ribbonEl.scrollLeft = 0;
  }

  ribbonItem(c: Command): HTMLElement {
    switch (c.kind) {
      case 'font': return this.fontBox(c);
      case 'size': return this.sizeBox(c);
      case 'gallery': return this.gallery(c);
      case 'combo': return this.comboBox(c);
      default: return this.commandButton(c, 'ribbon');
    }
  }

  commandButton(c: Command, where: 'ribbon' | 'qat' | 'mini'): HTMLElement {
    const showLabel = where === 'ribbon' || (where === 'qat' && this.app.prefs.qatLabels);
    const key = c.keys?.[0] ? ` (${prettyKey(c.keys[0])})` : '';
    const main = h('button', { class: 'ql-cmd', 'data-cmd': c.id, title: (c.tip ? `${c.label}: ${c.tip}` : c.label) + key, 'aria-label': c.label, type: 'button' }, icon(c.icon), showLabel ? h('span', { class: 'ql-cmd-label' }, c.label.replace(/…$/, '')) : null);
    if (c.kind === 'color') main.appendChild(h('span', { class: 'ql-colorbar' }));
    let longPress: any;
    main.addEventListener('pointerdown', () => { longPress = setTimeout(() => { longPress = 'fired'; this.longPress(c, main); }, 550); });
    main.addEventListener('pointerup', () => { if (longPress !== 'fired') clearTimeout(longPress); });
    main.addEventListener('pointerleave', () => { if (longPress !== 'fired') clearTimeout(longPress); });
    main.addEventListener('click', () => {
      if (longPress === 'fired') { longPress = null; return; }
      if (c.kind === 'menu') { this.openCommandMenu(c, main); return; }
      this.app.exec(c.id);
      if (where !== 'qat' && !['menu', 'split', 'color'].includes(c.kind || '')) setTimeout(() => this.app.focus(), 0);
    });
    main.addEventListener('mousedown', (e) => e.preventDefault());
    if (c.kind === 'menu') main.appendChild(icon('arrow_drop_down', 'ql-caret'));
    if (c.kind === 'split' || c.kind === 'color') {
      const arrow = h('button', { class: 'ql-split-arrow', type: 'button', title: `${c.label} options`, 'aria-label': `${c.label} options` }, icon('arrow_drop_down'));
      arrow.addEventListener('mousedown', (e) => e.preventDefault());
      arrow.addEventListener('click', () => (c.kind === 'color' ? this.app.exec(c.id, 'menu') : this.openCommandMenu(c, wrap)));
      const wrap = h('span', { class: 'ql-split', 'data-cmd-wrap': c.id }, main, arrow);
      return wrap;
    }
    return main;
  }

  longPress(c: Command, el: HTMLElement): void {
    const inQat = this.app.prefs.qat.includes(c.id);
    showMenu([
      { id: '__qat', label: inQat ? 'Remove from Quick Access Bar' : 'Add to Quick Access Bar', icon: inQat ? 'remove_circle' : 'add_circle' },
      ...(this.app.prefs.qatHidden ? [{ id: '__showQat', label: 'Show Quick Access Bar', icon: 'bolt' }] : []),
      ...(c.id === 'edit.painter' ? [{ id: '__sticky', label: 'Keep Format Painter On', icon: 'format_paint' }] : []),
      { id: '__key', label: 'Assign Shortcut…', icon: 'keyboard' },
      { id: '__info', label: `About: ${c.label}`, icon: 'info' },
    ], {
      anchor: el, onPick: (e) => {
        if (e.id === '__qat') { this.app.prefs.qat = inQat ? this.app.prefs.qat.filter((x) => x !== c.id) : [...this.app.prefs.qat, c.id]; this.app.savePrefs(); this.renderQat(); this.refresh(); toast(inQat ? 'Removed from Quick Access Bar' : 'Added to Quick Access Bar'); }
        if (e.id === '__showQat') this.app.exec('view.showQat');
        if (e.id === '__sticky') this.app.exec('edit.painter', 'sticky');
        if (e.id === '__key') this.customizeKeyboard(c.id);
        if (e.id === '__info') toast(`${c.label} — ${c.tab}${c.group ? ' ▸ ' + c.group : ''}${c.keys?.length ? ' — ' + c.keys.map(prettyKey).join(', ') : ''}${c.origin ? ' — inspired by ' + c.origin : ''}`, { ms: 5000 });
      },
    });
  }

  openCommandMenu(c: Command, anchor: HTMLElement): void {
    const entries = c.menu ? c.menu(this.app) : [];
    showMenu(entries, { anchor, onPick: (e) => this.pickEntry(c, e) });
  }

  pickEntry(c: Command, e: MenuEntry): void {
    if (!e.id) return;
    if (e.id === c.id) this.app.exec(c.id, e.arg);
    else this.app.exec(e.id, e.arg);
    setTimeout(() => this.app.focus(), 0);
  }

  fontBox(c: Command): HTMLElement {
    const b = h('button', { class: 'ql-combo ql-fontbox', 'data-cmd': c.id, 'data-combo': 'font', title: 'Font', type: 'button' }, h('span', { class: 'ql-combo-text' }, 'Font'), icon('arrow_drop_down'));
    b.addEventListener('mousedown', (e) => e.preventDefault());
    b.addEventListener('click', () => {
      const theme = this.app.book.theme;
      const entries: MenuEntry[] = [
        { heading: 'Theme Fonts' }, { arg: theme.headingFont, label: `${theme.headingFont} (Headings)`, preview: `font-family:"${theme.headingFont}"` }, { arg: theme.bodyFont, label: `${theme.bodyFont} (Body)`, preview: `font-family:"${theme.bodyFont}"` },
        { heading: 'All Fonts' }, ...FONT_LIST.map((f) => ({ arg: f, label: f, preview: `font-family:"${f}"` })),
      ];
      const input = h('input', { class: 'ql-menu-search', placeholder: 'Search fonts or type a name…' }) as HTMLInputElement;
      const m = showMenu(entries, { anchor: b, onPick: (e) => { this.app.exec('fmt.font', e.arg); this.app.focus(); }, className: 'ql-font-menu' });
      m.prepend(input);
      input.addEventListener('input', () => {
        const q = input.value.toLowerCase();
        m.querySelectorAll<HTMLElement>('.ql-menu-item').forEach((it) => { it.style.display = it.textContent!.toLowerCase().includes(q) ? '' : 'none'; });
      });
      input.addEventListener('keydown', (e) => { if (e.key === 'Enter' && input.value) { closeMenus(); this.app.exec('fmt.font', input.value); this.app.focus(); } });
      setTimeout(() => { if (!isPhone()) input.focus(); }, 20);
    });
    return b;
  }

  sizeBox(c: Command): HTMLElement {
    const wrap = h('span', { class: 'ql-sizebox', 'data-cmd-wrap': c.id });
    const input = h('input', { class: 'ql-size-input', 'data-combo': c.id, inputmode: 'decimal', 'aria-label': c.label, title: c.label }) as HTMLInputElement;
    input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') { e.preventDefault(); const val = parseFloat(input.value); if (!isNaN(val)) this.app.exec(c.id, val); this.app.focus(); }
      if (e.key === 'Escape') this.app.focus();
    });
    const arrow = h('button', { class: 'ql-split-arrow', type: 'button', 'aria-label': `${c.label} list` }, icon('arrow_drop_down'));
    arrow.addEventListener('mousedown', (e) => e.preventDefault());
    arrow.addEventListener('click', () => { if (c.menu) this.openCommandMenu(c, wrap); else this.app.exec(c.id); });
    wrap.append(input, arrow);
    return wrap;
  }

  /** Text combo (Number Format box, Scale to Fit boxes): shows the current value, opens a menu. */
  comboBox(c: Command): HTMLElement {
    const b = h('button', { class: 'ql-combo ql-wide-combo', 'data-cmd': c.id, 'data-combo': c.id, title: c.label, type: 'button' }, h('span', { class: 'ql-combo-text' }, c.label), icon('arrow_drop_down'));
    b.addEventListener('mousedown', (e) => e.preventDefault());
    b.addEventListener('click', () => this.openCommandMenu(c, b));
    return b;
  }

  gallery(c: Command): HTMLElement {
    const wrap = h('div', { class: 'ql-gallery', 'data-gallery': c.id });
    const strip = h('div', { class: 'ql-gallery-strip' });
    const more = h('button', { class: 'ql-gallery-more', title: 'More', 'aria-label': `More ${c.label}` }, icon('expand_more'));
    more.addEventListener('mousedown', (e) => e.preventDefault());
    more.addEventListener('click', () => this.openCommandMenu(c, wrap));
    wrap.append(strip, more);
    return wrap;
  }

  fillGalleries(): void {
    this.ribbonEl.querySelectorAll<HTMLElement>('[data-gallery]').forEach((wrap) => {
      const id = wrap.dataset.gallery!;
      const c = getCommand(id);
      if (!c?.menu) return;
      const strip = wrap.querySelector('.ql-gallery-strip')!;
      const entries = c.menu(this.app).filter((e) => !e.heading && !e.separator).slice(0, 12);
      const sig = JSON.stringify(entries.map((e) => [e.label, e.checked, e.preview]));
      if (strip.getAttribute('data-sig') === sig) return;
      strip.setAttribute('data-sig', sig);
      clear(strip as HTMLElement);
      for (const e of entries) {
        const chip = h('button', { class: 'ql-gallery-item' + (e.checked ? ' ql-on' : ''), title: e.label || '', type: 'button' }, h('span', { style: e.preview || '' }, e.label || ''));
        chip.addEventListener('mousedown', (ev) => ev.preventDefault());
        chip.addEventListener('click', () => { this.app.exec(e.id || id, e.arg); this.app.focus(); });
        strip.appendChild(chip);
      }
    });
  }

  // ----------------------------------------------------------------------------------------
  // Refresh (state → UI)

  refresh(): void {
    const app = this.app;
    if (!app.active) return;
    for (const t of DEFAULT_TABS.filter((x) => x.contextual)) {
      const el = this.tabsEl.querySelector<HTMLElement>(`[data-tab="${t.id}"]`);
      let show = false;
      try { show = t.contextual!(app); } catch { /* */ }
      if (el) el.style.display = show ? '' : 'none';
      if (!show && this.currentTab === t.id) { this.currentTab = 'home'; this.renderTabs(); this.renderRibbon(); }
    }
    for (const b of this.root.querySelectorAll<HTMLElement>('[data-cmd]')) {
      const c = getCommand(b.dataset.cmd!);
      if (!c) continue;
      let active = false, enabled = true;
      try { active = !!c.isActive?.(app); enabled = c.isEnabled ? c.isEnabled(app) : true; } catch { /* */ }
      b.classList.toggle('ql-on', active);
      b.setAttribute('aria-pressed', String(active));
      b.classList.toggle('ql-disabled', !enabled);
    }
    const st = activeStyle(app);
    const base = app.book.styles[0] || {};
    const fontEl = this.ribbonEl.querySelector<HTMLElement>('[data-combo=font] .ql-combo-text');
    if (fontEl) { const f = st.font || base.font || app.book.theme.bodyFont; fontEl.textContent = f; fontEl.style.fontFamily = `"${f}"`; }
    this.root.querySelectorAll<HTMLInputElement>('input[data-combo]').forEach((inp) => {
      if (document.activeElement === inp) return;
      const c = getCommand(inp.dataset.combo!);
      const v = c?.value?.(app);
      if (v != null) inp.value = String(v);
      else if (inp.dataset.combo === 'fmt.size') inp.value = String(st.size || base.size || 11);
    });
    this.root.querySelectorAll<HTMLElement>('button[data-combo]').forEach((b) => {
      const c = getCommand(b.dataset.combo!);
      const v = c?.value?.(app);
      const t = b.querySelector('.ql-combo-text');
      if (t && v != null) t.textContent = String(v);
    });
    this.fillGalleries();
    this.refreshFormulaBar();
    this.renderSheetBar();
    this.renderStatus();
    if (this.pane?.def.update) this.pane.def.update(app, this.pane.el, this.pane.opts);
    this.titleBtn.textContent = app.tab.displayName + (app.tab.dirty ? ' •' : '');
    this.root.classList.toggle('ql-no-formulabar', !app.prefs.showFormulaBar);
    app.tab.grid.invalidate();
  }

  refreshFormulaBar(): void {
    const app = this.app;
    const g = app.grid;
    const s = app.sheet;
    const { r, c } = s.sel;
    if (document.activeElement !== this.nameBox) {
      const rg = s.sel.ranges[s.sel.ranges.length - 1];
      const nm = app.book.names.find((n) => !n.hidden && (n.ref.replace(/\$/g, '').replace(/^=/, '').toLowerCase() === `${s.name.toLowerCase()}!${rangeName(rg).toLowerCase()}` || n.ref.replace(/\$/g, '').replace(/^=/, '').replace(/^'|'(?=!)/g, '').toLowerCase() === `${s.name.toLowerCase()}!${rangeName(rg).toLowerCase()}`));
      const dragSize = rg.r1 !== rg.r2 || rg.c1 !== rg.c2 ? null : null;
      this.nameBox.value = nm ? nm.name : app.prefs.r1c1 ? `R${r + 1}C${c + 1}` : cellName(r, c);
      void dragSize;
    }
    if (!g.editing && document.activeElement !== this.barInput) {
      const cell = s.get(r, c);
      const hidden = s.protection.on && cell?.s && app.book.styles[cell.s]?.hidden;
      this.barInput.value = hidden ? '' : editText(g, r, c);
    }
  }

  // ----------------------------------------------------------------------------------------
  // Sheet tabs

  renderSheetBar(): void {
    const app = this.app;
    const book = app.book;
    const sig = JSON.stringify([book.sheets.map((s) => [s.name, s.tabColor, s.visibility, s.protection.on]), book.activeSheet, app.tab.id]);
    if (this.sheetBar.dataset.sig === sig) return;
    this.sheetBar.dataset.sig = sig;
    clear(this.sheetBar);
    const nav = h('div', { class: 'ql-sheetnav' },
      h('button', { class: 'ql-iconbtn ql-sm', title: 'Previous sheet (Ctrl+PgUp)', 'aria-label': 'Previous sheet', onclick: () => app.exec('sheet.prev') }, icon('chevron_left')),
      h('button', { class: 'ql-iconbtn ql-sm', title: 'Next sheet (Ctrl+PgDn)', 'aria-label': 'Next sheet', onclick: () => app.exec('sheet.next') }, icon('chevron_right')),
      h('button', { class: 'ql-iconbtn ql-sm', title: 'All sheets', 'aria-label': 'All sheets', onclick: (e: MouseEvent) => showMenu(book.sheets.map((s, i) => ({ id: 'sheet.activate', arg: i, label: s.name + (s.visibility !== 'visible' ? ' (hidden)' : ''), checked: i === book.activeSheet })), { anchor: e.currentTarget as HTMLElement, onPick: (en) => en.id && app.exec(en.id, en.arg) }) }, icon('more_horiz')));
    const tabs = h('div', { class: 'ql-sheettabs' });
    book.sheets.forEach((s, i) => {
      if (s.visibility !== 'visible') return;
      const b = h('button', { class: 'ql-sheettab' + (i === book.activeSheet ? ' ql-active' : ''), role: 'tab', 'aria-selected': String(i === book.activeSheet), draggable: 'true', style: s.tabColor ? `--tab-color:${s.tabColor}` : '' }, s.protection.on ? icon('lock', 'ql-tab-lock') : null, s.name);
      b.addEventListener('click', () => app.exec('sheet.activate', i));
      b.addEventListener('dblclick', () => app.exec('sheet.rename', i));
      b.addEventListener('contextmenu', (e) => { e.preventDefault(); app.exec('sheet.activate', i); this.sheetMenu(e.clientX, e.clientY); });
      b.addEventListener('dragstart', (e) => { e.dataTransfer!.setData('text/x-ql-sheet', String(i)); e.dataTransfer!.effectAllowed = 'move'; });
      b.addEventListener('dragover', (e) => { if (e.dataTransfer?.types.includes('text/x-ql-sheet')) { e.preventDefault(); b.classList.add('ql-drop'); } });
      b.addEventListener('dragleave', () => b.classList.remove('ql-drop'));
      b.addEventListener('drop', (e) => { e.preventDefault(); b.classList.remove('ql-drop'); const from = +e.dataTransfer!.getData('text/x-ql-sheet'); if (from !== i) app.exec('sheet.move', { from, to: i, copy: e.ctrlKey }); });
      let lp: any;
      b.addEventListener('pointerdown', (e) => { if (e.pointerType === 'touch') lp = setTimeout(() => { app.exec('sheet.activate', i); this.sheetMenu(e.clientX, e.clientY); }, 550); });
      b.addEventListener('pointerup', () => clearTimeout(lp));
      b.addEventListener('pointerleave', () => clearTimeout(lp));
      tabs.appendChild(b);
    });
    const add = h('button', { class: 'ql-iconbtn ql-sm ql-addsheet', title: 'New sheet (Shift+F11)', 'aria-label': 'New sheet', onclick: () => app.exec('sheet.insert') }, icon('add_circle'));
    this.sheetBar.append(nav, tabs, add);
    requestAnimationFrame(() => tabs.querySelector('.ql-active')?.scrollIntoView({ inline: 'nearest', block: 'nearest' }));
  }

  sheetMenu(x: number, y: number): void {
    const app = this.app;
    const s = app.sheet;
    showMenu([
      { id: 'sheet.insert', label: 'Insert Sheet', icon: 'add' },
      { id: 'sheet.delete', label: 'Delete', icon: 'delete' },
      { id: 'sheet.rename', label: 'Rename', icon: 'edit' },
      { id: 'sheet.moveCopy', label: 'Move or Copy…', icon: 'drive_file_move' },
      { id: 'sheet.duplicate', label: 'Duplicate', icon: 'content_copy' },
      { id: 'sheet.tabColor', label: 'Tab Color', icon: 'palette' },
      { id: 'review.protectSheet', label: s.protection.on ? 'Unprotect Sheet…' : 'Protect Sheet…', icon: 'lock' },
      { separator: true },
      { id: 'sheet.hide', label: 'Hide', icon: 'visibility_off' },
      { id: 'sheet.unhide', label: 'Unhide…', icon: 'visibility' },
      { id: 'sheet.selectAll', label: 'Select All Sheets', icon: 'select_all' },
      { separator: true },
      { id: 'sheet.exportCsv', label: 'Save This Sheet as CSV…', icon: 'download' },
    ], { x, y, onPick: (e) => e.id && app.exec(e.id, e.arg) });
  }

  // ----------------------------------------------------------------------------------------
  // Status bar

  renderStatus(): void {
    const app = this.app;
    const tab = app.tab;
    const g = tab.grid;
    const s = app.sheet;
    clear(this.statusEl);
    const item = (text: string, onclick?: () => void, title?: string, cls = '') => h('button', { class: 'ql-status-item ' + cls, onclick, title: title || text, type: 'button' }, text);
    const mode = g.editing ? (g.editor.mode === 'point' ? 'Point' : g.editor.mode === 'edit' ? 'Edit' : 'Enter') : 'Ready';
    this.statusEl.appendChild(h('span', { class: 'ql-status-mode' }, mode));
    if (app.book.calc.mode === 'manual' && app.book.pending.cells.size) this.statusEl.appendChild(item('Calculate', () => app.exec('formulas.calcNow'), 'Recalculate (F9)', 'ql-status-warn'));
    if (tab.engine.circular.length) this.statusEl.appendChild(item('Circular References', () => app.exec('formulas.circularRefs'), 'Show circular references', 'ql-status-warn'));
    if (app.recording) this.statusEl.appendChild(item(app.recordingPaused ? '⏸ Paused' : '● REC', () => app.exec('dev.stopRecording'), 'Stop recording macro', 'ql-status-rec'));
    if (s.protection.on || app.book.structureLocked) this.statusEl.appendChild(item('🔒', () => app.exec('review.protectSheet'), 'Protection'));
    if (app.book.finalMarked) this.statusEl.appendChild(item('Marked as Final', () => app.exec('review.markFinal'), 'Marked as final'));
    if (s.filter && Object.keys(s.filter.columns).length) {
      const shown = this.filteredCount();
      this.statusEl.appendChild(h('span', { class: 'ql-status-msg' }, shown));
    }
    if (this.statusMsg) this.statusEl.appendChild(h('span', { class: 'ql-status-msg' }, this.statusMsg));
    this.statusEl.appendChild(h('span', { class: 'ql-spacer' }));
    // Selection statistics
    const stats = this.selectionStats();
    const fmt = activeStyle(app).numFmt;
    const show = (k: string, label: string) => {
      const v = stats[k];
      if (v == null || !app.prefs.statusItems.includes(k)) return;
      const text = k === 'count' || k === 'numCount' ? String(v) : formatValue(v, fmt && !/@/.test(fmt) ? fmt : 'General', { date1904: app.book.calc.date1904 }).text;
      this.statusEl.appendChild(item(`${label}: ${text}`, () => { navigator.clipboard?.writeText(String(v)); toast(`${label} copied`); }, `${label} (click to copy)`));
    };
    if (stats.count != null && stats.count > 1) { show('average', 'Average'); show('count', 'Count'); show('numCount', 'Numerical Count'); show('min', 'Min'); show('max', 'Max'); show('sum', 'Sum'); }
    this.statusEl.appendChild(h('button', { class: 'ql-iconbtn ql-sm', title: 'Customize status bar', 'aria-label': 'Customize status bar', onclick: (e: MouseEvent) => this.statusMenu(e.currentTarget as HTMLElement) }, icon('tune')));
    const views: [string, string, string][] = [['normal', 'grid_on', 'Normal'], ['pageLayout', 'description', 'Page Layout'], ['pageBreak', 'splitscreen', 'Page Break Preview']];
    for (const [m, ic, l] of views) this.statusEl.appendChild(h('button', { class: 'ql-iconbtn ql-status-view' + (tab.viewMode === m ? ' ql-on' : ''), title: l, 'aria-label': l, onclick: () => app.exec(`view.${m}`) }, icon(ic)));
    const z = s.zoom || 1;
    const slider = h('input', { type: 'range', class: 'ql-zoom-slider', min: '10', max: '400', step: '5', value: String(Math.round(z * 100)), 'aria-label': 'Zoom' }) as HTMLInputElement;
    slider.addEventListener('input', () => g.setZoom(+slider.value / 100, true));
    slider.addEventListener('change', () => g.setZoom(+slider.value / 100));
    this.statusEl.append(h('button', { class: 'ql-iconbtn', title: 'Zoom out', 'aria-label': 'Zoom out', onclick: () => app.exec('view.zoomOut') }, icon('remove')), slider,
      h('button', { class: 'ql-iconbtn', title: 'Zoom in', 'aria-label': 'Zoom in', onclick: () => app.exec('view.zoomIn') }, icon('add')),
      h('button', { class: 'ql-status-item ql-zoom', title: 'Zoom…', onclick: () => app.exec('view.zoom') }, `${Math.round(z * 100)}%`));
  }

  private filteredCount(): string {
    const s = this.app.sheet;
    const f = s.filter!;
    let total = 0, shown = 0;
    for (let r = f.range.r1 + 1; r <= f.range.r2; r++) { total++; if (!s.hiddenRows.has(r)) shown++; }
    return `${shown} of ${total} records found`;
  }

  selectionStats(): Record<string, number | null> {
    const app = this.app;
    const s = app.sheet;
    const key = `${app.tab.id}:${s.id}:${app.book.version}:${app.tab.engine.evaluations}:${JSON.stringify(s.sel.ranges)}`;
    if (key === this.statsKey) return this.statsCache;
    this.statsKey = key;
    let count = 0, numCount = 0, sum = 0, min = Infinity, max = -Infinity;
    const seen = new Set<number>();
    let cells = 0;
    for (const rg of s.sel.ranges) {
      s.forEachIn(rg, (r, c, cell) => {
        const k = r * MAX_COLS + c;
        if (seen.has(k)) return;
        seen.add(k);
        if (s.hiddenRows.has(r) || s.hiddenCols.has(c)) return;
        cells++;
        const v = cell.v;
        if (v == null || v === '') return;
        count++;
        if (typeof v === 'number') { numCount++; sum += v; if (v < min) min = v; if (v > max) max = v; }
      });
    }
    void cells;
    this.statsCache = count ? { count, numCount, sum: numCount ? sum : 0, average: numCount ? sum / numCount : null, min: numCount ? min : null, max: numCount ? max : null } : { count: 0 };
    return this.statsCache;
  }

  statusMenu(anchor: HTMLElement): void {
    const p = this.app.prefs;
    const items: [string, string][] = [['average', 'Average'], ['count', 'Count'], ['numCount', 'Numerical Count'], ['min', 'Minimum'], ['max', 'Maximum'], ['sum', 'Sum']];
    showMenu([{ heading: 'Show on status bar' }, ...items.map(([k, l]) => ({ arg: k, label: l, checked: p.statusItems.includes(k) }))], {
      anchor, onPick: (e) => { const k = e.arg as string; p.statusItems = p.statusItems.includes(k) ? p.statusItems.filter((x) => x !== k) : [...p.statusItems, k]; this.app.savePrefs(); this.statsKey = ''; this.renderStatus(); },
    });
  }

  status(msg: string): void {
    this.statusMsg = msg;
    clearTimeout(this.statusTimer);
    if (msg && !msg.startsWith('●')) this.statusTimer = setTimeout(() => { this.statusMsg = ''; this.renderStatus(); }, 3000);
    this.renderStatus();
  }

  // ----------------------------------------------------------------------------------------
  // Workbook tabs strip and title

  tabsChanged(): void {
    const app = this.app;
    clear(this.docTabsEl);
    this.root.classList.toggle('ql-multi-doc', app.docs.length > 1);
    for (const d of app.docs) {
      const t = h('div', { class: 'ql-doctab' + (d === app.active ? ' ql-active' : ''), role: 'tab' },
        h('button', { class: 'ql-doctab-name', onclick: () => app.activate(d) }, d.displayName + (d.dirty ? ' •' : '')),
        h('button', { class: 'ql-doctab-close', 'aria-label': `Close ${d.displayName}`, onclick: () => { app.activate(d); app.exec('file.close', d.id); } }, icon('close')));
      this.docTabsEl.appendChild(t);
    }
    this.docTabsEl.appendChild(h('button', { class: 'ql-iconbtn', title: 'New workbook', 'aria-label': 'New workbook', onclick: () => app.exec('file.new') }, icon('add')));
    if (app.active) this.titleBtn.textContent = app.active.displayName + (app.active.dirty ? ' •' : '');
    this.sheetBar.dataset.sig = '';
  }

  async renameDoc(): Promise<void> {
    const tab = this.app.tab;
    const r = await openDialog({ title: 'Workbook', fields: [{ key: 'name', label: 'File name', type: 'text', value: tab.fileName || tab.title }, { key: 'title', label: 'Title property', type: 'text', value: this.app.book.props.title }] });
    if (!r) return;
    tab.fileName = r.name; tab.title = r.name.replace(/\.[^.]+$/, '');
    this.app.book.props.title = r.title;
    this.tabsChanged();
  }

  moreMenu(anchor: HTMLElement): void {
    const app = this.app;
    showMenu([
      { id: 'file.save', label: 'Save', icon: 'save' }, { id: 'file.share', label: 'Share', icon: 'share' }, { id: 'file.print', label: 'Print / PDF', icon: 'print' },
      { separator: true },
      { id: 'view.navigator', label: 'Navigator', icon: 'list_alt' }, { id: 'review.showComments', label: 'Comments', icon: 'forum' }, { id: 'edit.find', label: 'Find & Replace', icon: 'search' }, { id: 'formulas.watchWindow', label: 'Watch Window', icon: 'visibility' },
      { separator: true },
      { id: 'view.showQat', label: app.prefs.qatHidden ? 'Show Quick Access Bar' : 'Hide Quick Access Bar', icon: 'bolt' },
      { id: 'view.qatAtBottom', label: app.prefs.qatPosition === 'bottom' ? 'Quick Access Bar Above the Ribbon' : 'Quick Access Bar at the Bottom', icon: 'swap_vert' },
      { id: 'file.customizeQat', label: 'Customize Quick Access Bar…', icon: 'tune' }, { id: 'file.customizeRibbon', label: 'Customize Tabs…', icon: 'view_quilt' },
      { id: 'file.customizeKeyboard', label: 'Customize Keyboard…', icon: 'keyboard' }, { id: 'file.options', label: 'Options…', icon: 'settings' },
      { id: 'help.about', label: 'About', icon: 'info' },
    ], { anchor, onPick: (e) => e.id && app.exec(e.id, e.arg) });
  }

  // ----------------------------------------------------------------------------------------
  // Panes

  openPane(id: string, opts: any = {}): void {
    const def = PANES[id];
    if (!def) { toast(`Unknown pane ${id}`); return; }
    if (this.pane?.def.id === id) { this.pane.opts = opts; def.update?.(this.app, this.pane.el, opts); def.onOpen?.(this.app, this.pane.el, opts); return; }
    this.closePane();
    const body = h('div', { class: 'ql-pane-body' });
    const head = h('div', { class: 'ql-pane-head' }, h('h2', null, def.title), h('button', { class: 'ql-iconbtn', 'aria-label': 'Close pane', onclick: () => this.closePane() }, icon('close')));
    clear(this.paneEl);
    this.paneEl.append(head, body);
    this.paneEl.setAttribute('aria-hidden', 'false');
    this.root.classList.add('ql-pane-open');
    this.pane = { def, el: body, opts };
    def.render(this.app, body, opts);
    def.onOpen?.(this.app, body, opts);
    this.refresh();
    requestAnimationFrame(() => this.app.grid.layout());
  }

  closePane(): void {
    if (!this.pane) return;
    this.pane.def.onClose?.(this.app);
    this.pane = null;
    clear(this.paneEl);
    this.paneEl.setAttribute('aria-hidden', 'true');
    this.root.classList.remove('ql-pane-open');
    this.app.refresh();
    requestAnimationFrame(() => this.app.grid.layout());
  }

  togglePane(id: string, opts?: any): void { if (this.pane?.def.id === id) this.closePane(); else this.openPane(id, opts); }
  paneOpen(id: string): boolean { return this.pane?.def.id === id; }

  // ----------------------------------------------------------------------------------------
  // Backstage

  showBackstage(page = 'home'): void {
    closeMenus();
    if (this.app.grid?.editing) this.app.grid.commitEdit('none');
    this.backstageEl.setAttribute('aria-hidden', 'false');
    this.root.classList.add('ql-backstage-open');
    renderBackstage(this.app, this.backstageEl, page);
  }
  hideBackstage(): void {
    this.backstageEl.setAttribute('aria-hidden', 'true');
    this.root.classList.remove('ql-backstage-open');
    clear(this.backstageEl);
  }

  customizeQat(): void { customizeQat(this.app); }
  customizeRibbon(): void { customizeRibbon(this.app); }
  customizeKeyboard(id?: string): void { customizeKeyboard(this.app, id); }

  // ----------------------------------------------------------------------------------------
  // Command palette

  commandPalette(): void {
    const app = this.app;
    const overlay = h('div', { class: 'ql-overlay ql-palette-overlay' });
    const input = h('input', { class: 'ql-palette-input', placeholder: `Search ${allCommands().length} commands… (e.g. "freeze", "pivot", "merge")`, 'aria-label': 'Search commands' }) as HTMLInputElement;
    const list = h('div', { class: 'ql-palette-list', role: 'listbox' });
    const box = h('div', { class: 'ql-palette', role: 'dialog' }, input, list);
    overlay.appendChild(box);
    document.body.appendChild(overlay);
    let results: Command[] = [];
    let sel = 0;
    const close = () => { overlay.remove(); app.focus(); };
    const run = (c: Command) => { close(); app.exec(c.id); };
    const draw = () => {
      clear(list);
      const q = input.value.trim();
      results = q ? searchCommands(q, 80) : [...app.prefs.pinnedCommands.map((id) => getCommand(id)!).filter(Boolean), ...allCommands().filter((c) => c.popular).slice(0, 40)];
      results.forEach((c, i) => {
        const pinned = app.prefs.qat.includes(c.id);
        const row = h('div', { class: 'ql-palette-row' + (i === sel ? ' ql-sel' : ''), role: 'option' },
          icon(c.icon), h('span', { class: 'ql-palette-label' }, c.label), h('span', { class: 'ql-palette-tab' }, `${c.tab === 'Backstage' ? 'File' : c.tab}${c.group ? ' ▸ ' + c.group : ''}`),
          c.keys?.[0] ? h('kbd', null, prettyKey(c.keys[0])) : null,
          h('button', { class: 'ql-iconbtn ql-palette-pin', title: pinned ? 'Remove from Quick Access Bar' : 'Add to Quick Access Bar', onclick: (e: Event) => { e.stopPropagation(); app.prefs.qat = pinned ? app.prefs.qat.filter((x) => x !== c.id) : [...app.prefs.qat, c.id]; app.savePrefs(); this.renderQat(); draw(); } }, icon(pinned ? 'push_pin' : 'keep')));
        row.addEventListener('click', () => run(c));
        list.appendChild(row);
      });
      if (!results.length) list.appendChild(h('div', { class: 'ql-help' }, 'No matching commands.'));
    };
    input.addEventListener('input', () => { sel = 0; draw(); });
    input.addEventListener('keydown', (e) => {
      if (e.key === 'ArrowDown') { sel = Math.min(results.length - 1, sel + 1); draw(); list.children[sel]?.scrollIntoView({ block: 'nearest' }); e.preventDefault(); }
      if (e.key === 'ArrowUp') { sel = Math.max(0, sel - 1); draw(); list.children[sel]?.scrollIntoView({ block: 'nearest' }); e.preventDefault(); }
      if (e.key === 'Enter' && results[sel]) { e.preventDefault(); run(results[sel]); }
      if (e.key === 'Escape') close();
    });
    overlay.addEventListener('mousedown', (e) => { if (e.target === overlay) close(); });
    draw();
    setTimeout(() => input.focus(), 20);
  }

  // ----------------------------------------------------------------------------------------
  // Context menus

  gridContextMenu(x: number, y: number, hit: Hit): void {
    const app = this.app;
    const s = app.sheet;
    const entries: MenuEntry[] = [];
    const pasteMenu: MenuEntry[] = [
      { id: 'edit.paste', label: 'Paste', icon: 'content_paste' }, { id: 'edit.pasteValues', label: 'Values', icon: '123' }, { id: 'edit.pasteFormulas', label: 'Formulas', icon: 'function' },
      { id: 'edit.pasteFormats', label: 'Formatting', icon: 'format_paint' }, { id: 'edit.pasteTranspose', label: 'Transpose', icon: 'swap_horiz' }, { id: 'edit.pasteLink', label: 'Paste Link', icon: 'link' },
      { id: 'edit.pasteNoBorders', label: 'No Borders', icon: 'border_clear' }, { id: 'edit.pasteColumnWidths', label: 'Keep Source Column Widths', icon: 'width' }, { separator: true }, { id: 'edit.pasteSpecial', label: 'Paste Special…', icon: 'content_paste_go' },
    ];
    entries.push({ id: 'edit.cut', label: 'Cut', icon: 'content_cut' }, { id: 'edit.copy', label: 'Copy', icon: 'content_copy' }, { label: 'Paste Options', icon: 'content_paste', submenu: pasteMenu });
    entries.push({ separator: true });
    if (hit.area === 'colHeader') {
      entries.push({ id: 'cells.insertCols', label: 'Insert', icon: 'add' }, { id: 'cells.deleteCols', label: 'Delete', icon: 'delete' }, { id: 'edit.clearContents', label: 'Clear Contents', icon: 'backspace' },
        { separator: true }, { id: 'fmt.cellsDialog', label: 'Format Cells…', icon: 'table_view' }, { id: 'format.colWidthDialog', label: 'Column Width…', icon: 'width' }, { id: 'format.autofitCols', label: 'AutoFit Column Width', icon: 'fit_width' },
        { id: 'format.hideCols', label: 'Hide', icon: 'visibility_off' }, { id: 'format.unhideCols', label: 'Unhide', icon: 'visibility' }, { id: 'data.group', label: 'Group', icon: 'unfold_less' });
    } else if (hit.area === 'rowHeader') {
      entries.push({ id: 'cells.insertRows', label: 'Insert', icon: 'add' }, { id: 'cells.deleteRows', label: 'Delete', icon: 'delete' }, { id: 'edit.clearContents', label: 'Clear Contents', icon: 'backspace' },
        { separator: true }, { id: 'fmt.cellsDialog', label: 'Format Cells…', icon: 'table_view' }, { id: 'format.rowHeightDialog', label: 'Row Height…', icon: 'height' }, { id: 'format.autofitRows', label: 'AutoFit Row Height', icon: 'fit_screen' },
        { id: 'format.hideRows', label: 'Hide', icon: 'visibility_off' }, { id: 'format.unhideRows', label: 'Unhide', icon: 'visibility' }, { id: 'data.group', label: 'Group', icon: 'unfold_less' });
    } else {
      const { r, c } = s.sel;
      const cell = s.get(r, c);
      const cm = s.comments.get(r * MAX_COLS + c);
      entries.push({ id: 'ai.quickAnalysis', label: 'Quick Analysis', icon: 'bolt' }, { separator: true },
        { id: 'cells.insertDialog', label: 'Insert…', icon: 'add_box' }, { id: 'cells.deleteDialog', label: 'Delete…', icon: 'disabled_by_default' }, { id: 'edit.clearContents', label: 'Clear Contents', icon: 'backspace' },
        { separator: true },
        { label: 'Filter', icon: 'filter_alt', submenu: [{ id: 'data.filterBySelected', label: "Filter by Selected Cell's Value" }, { id: 'data.filterByColor', label: "Filter by Selected Cell's Color" }, { id: 'data.reapply', label: 'Reapply' }, { id: 'data.clearFilter', label: 'Clear Filter' }] },
        { label: 'Sort', icon: 'sort', submenu: [{ id: 'data.sortAZ', label: 'Sort A to Z' }, { id: 'data.sortZA', label: 'Sort Z to A' }, { id: 'data.sortByColor', label: "Put Selected Cell Color on Top" }, { id: 'data.sort', label: 'Custom Sort…' }] },
        { id: 'data.fromTable', label: 'Get Data from Table/Range…', icon: 'table_chart' },
        { separator: true },
        cm && !cm.note ? { id: 'review.editComment', label: 'Edit Comment', icon: 'edit_note' } : { id: 'review.newComment', label: 'New Comment', icon: 'add_comment' },
        cm?.note ? { id: 'review.editNote', label: 'Edit Note', icon: 'sticky_note_2' } : { id: 'review.newNote', label: 'New Note', icon: 'sticky_note_2' },
        ...(cm ? [{ id: 'review.deleteComment', label: cm.note ? 'Delete Note' : 'Delete Comment', icon: 'delete' }] : []),
        { separator: true },
        { id: 'fmt.cellsDialog', label: 'Format Cells…', icon: 'table_view' },
        { id: 'data.pickFromList', label: 'Pick From Drop-down List…', icon: 'list' },
        { id: 'formulas.defineName', label: 'Define Name…', icon: 'badge' },
        cell?.link ? { label: 'Link', icon: 'link', submenu: [{ id: 'insert.link', label: 'Edit Link…' }, { id: 'insert.openLink', label: 'Open Link' }, { id: 'insert.removeLink', label: 'Remove Link' }] } : { id: 'insert.link', label: 'Link…', icon: 'link' },
        { id: 'review.smartLookup', label: 'Smart Lookup', icon: 'travel_explore' });
    }
    showMenu(entries, { x, y, onPick: (en) => { if (en.id) app.exec(en.id, en.arg); } });
  }

  objectContextMenu(x: number, y: number, id: string): void {
    const app = this.app;
    const o = app.grid.objects.selectedObject();
    const isChart = o?.kind === 'chart';
    showMenu([
      { id: 'edit.cut', label: 'Cut', icon: 'content_cut' }, { id: 'edit.copy', label: 'Copy', icon: 'content_copy' }, { id: 'object.delete', arg: { id }, label: 'Delete', icon: 'delete' },
      { separator: true },
      ...(isChart ? [{ id: 'chart.edit', arg: { id }, label: 'Edit Chart…', icon: 'edit' }, { id: 'chart.selectData', label: 'Select Data…', icon: 'dataset' }, { id: 'chart.changeType', label: 'Change Chart Type…', icon: 'bar_chart' }, { id: 'chart.saveImage', label: 'Save as Picture…', icon: 'image' }] as MenuEntry[]
        : [{ id: 'drawing.editText', arg: { id }, label: 'Edit Text', icon: 'edit' }, { id: 'picture.altText', label: 'Edit Alt Text…', icon: 'image_search' }, { id: 'picture.size', label: 'Size and Properties…', icon: 'photo_size_select_large' }] as MenuEntry[]),
      { separator: true },
      { id: 'object.bringToFront', label: 'Bring to Front', icon: 'flip_to_front' }, { id: 'object.sendToBack', label: 'Send to Back', icon: 'flip_to_back' },
    ], { x, y, onPick: (e) => e.id && app.exec(e.id, e.arg) });
  }

  // ----------------------------------------------------------------------------------------

  anchorFor(id: string): HTMLElement {
    const els = [...this.root.querySelectorAll<HTMLElement>(`[data-cmd-wrap="${id}"], [data-cmd="${id}"]`)].filter((e) => e.offsetParent);
    return els[0] || this.ribbonEl;
  }

  applyPrefs(): void {
    const p = this.app.prefs;
    const dark = p.theme === 'dark' || (p.theme === 'system' && window.matchMedia('(prefers-color-scheme: dark)').matches);
    document.documentElement.dataset.theme = dark ? 'dark' : 'light';
    this.root.classList.toggle('ql-ribbon-collapsed', p.ribbon !== 'full');
    this.root.classList.toggle('ql-qat-at-bottom', p.qatPosition === 'bottom');
    this.root.classList.toggle('ql-qat-hidden', p.qatHidden);
    this.root.classList.toggle('ql-touch', p.touchMode);
    this.root.classList.toggle('ql-high-contrast', p.highContrastUI);
    this.root.style.setProperty('--ql-ui-scale', String(p.uiScale || 1));
    if (!this.app.docs.length) return;
    this.renderQat();
    for (const d of this.app.docs) d.grid.invalidate();
    this.refresh();
  }

  globalKey(e: KeyboardEvent): void {
    const k = keyName(e);
    if (!k) return;
    if (e.key === 'Escape') {
      if (document.querySelector('.ql-menu')) { closeMenus(); return; }
      if (this.root.classList.contains('ql-backstage-open')) { this.hideBackstage(); return; }
      if (document.body.classList.contains('ql-fullscreen')) { document.body.classList.remove('ql-fullscreen'); return; }
    }
    if (this.app.active && this.app.grid.el.contains(e.target as Node)) return; // the grid handles its own keys
    const t = e.target as HTMLElement;
    if (t && (t.tagName === 'INPUT' || t.tagName === 'TEXTAREA' || t.tagName === 'SELECT')) return;
    if (document.querySelector('.ql-overlay')) return;
    const id = this.app.keyIndex.get(canonicalKey(k));
    if (id && /^(file|view|help|dev\.macros|edit\.find|edit\.replace|edit\.undo|edit\.redo)/.test(id)) { e.preventDefault(); this.app.exec(id); }
  }

  async showShortcuts(): Promise<void> {
    const rows = allCommands().filter((c) => (c.keys && c.keys.length) || Object.values(this.app.prefs.keymap).includes(c.id));
    const custom = Object.entries(this.app.prefs.keymap);
    const nav: [string, string][] = [['Arrow keys', 'Move one cell'], ['Ctrl+Arrow', 'Move to the edge of the data region'], ['Shift+Arrow', 'Extend the selection'], ['Ctrl+Shift+Arrow', 'Extend to the edge of the data region'], ['Home / Ctrl+Home / Ctrl+End', 'Start of row / A1 / last used cell'], ['PgUp / PgDn, Alt+PgUp / Alt+PgDn', 'Move one screen up/down/left/right'], ['Enter / Shift+Enter', 'Move down/up (within a selection)'], ['Tab / Shift+Tab', 'Move right/left'], ['F2', 'Edit the active cell'], ['Alt+Enter', 'New line in a cell'], ['Ctrl+Enter', 'Fill the selection with the entry'], ['Ctrl+Shift+Enter', 'Enter as an array formula'], ['F4 (while editing)', 'Cycle absolute/relative reference'], ['Esc', 'Cancel editing']];
    const table = h('table', { class: 'ql-stats ql-shortcuts' },
      ...nav.map(([a, b]) => h('tr', null, h('td', null, b), h('td', null, a), h('td', { class: 'ql-help' }, 'Grid'))),
      ...rows.sort((a, b) => a.tab.localeCompare(b.tab) || a.label.localeCompare(b.label)).map((c) => {
        const keys = [...(c.keys || []), ...custom.filter(([, id]) => id === c.id).map(([k2]) => k2)];
        return h('tr', null, h('td', null, c.label), h('td', null, keys.map(prettyKey).join(', ')), h('td', { class: 'ql-help' }, c.tab === 'Backstage' ? 'File' : c.tab));
      }));
    await openDialog({ title: 'Keyboard Shortcuts', wide: true, content: table, cancel: false });
  }

  async showCommandReference(): Promise<void> {
    const byTab = new Map<string, Command[]>();
    for (const c of allCommands()) { const t = c.tab === 'Backstage' ? 'File' : c.tab; if (!byTab.has(t)) byTab.set(t, []); byTab.get(t)!.push(c); }
    const div = h('div', { class: 'ql-cmdref' });
    div.appendChild(h('p', { class: 'ql-help' }, `${allCommands().length} commands. Every command can be added to the Quick Access Bar, a custom tab, a keyboard shortcut, or a macro.`));
    for (const [t, cs] of byTab) {
      div.appendChild(h('h3', null, `${t} (${cs.length})`));
      div.appendChild(h('ul', null, ...cs.sort((a, b) => a.label.localeCompare(b.label)).map((c) => h('li', null, h('b', null, c.label), c.keys?.length ? ` — ${c.keys.map(prettyKey).join(', ')}` : '', c.origin ? h('span', { class: 'ql-help' }, ` (from ${c.origin})`) : ''))));
    }
    await openDialog({ title: 'Command Reference', wide: true, content: div, cancel: false });
  }

  async tour(): Promise<void> {
    const steps = [
      ['Everything is a command', 'Every feature of Anne\'s Grid is a command you can find with the search button (or Alt+Q), place on the Quick Access Bar, put on your own tab, assign a key, or record in a macro.'],
      ['Quick Access Bar', 'Tap ▾ at the end of the Quick Access Bar to add common commands, or choose More Commands… to pick from the full library. Long-press any ribbon button to add it. Options lets you move the bar to the bottom, near the keyboard.'],
      ['Cells and formulas', 'Type in a cell and press Enter. Start with = for a formula: click cells to insert references, and pick functions from the list as you type. Formulas that return several values spill into neighboring cells.'],
      ['Touch', 'On a phone or tablet, tap to select, double-tap to edit, and drag either round handle to extend a selection. Or press and hold a cell, a row number or a column letter and then drag to select a range. Pinch to zoom, and press and hold then lift for the context menu.'],
      ['Files', 'Open and save .xlsx, .ods, .csv and more with the system file picker. Damaged workbooks are repaired on open; File ▸ Open ▸ Open and Repair forces it. Print or Save as PDF through the system print service.'],
    ];
    for (let i = 0; i < steps.length; i++) {
      const r = await openDialog({ title: `${steps[i][0]} (${i + 1}/${steps.length})`, content: h('p', { class: 'ql-msg' }, steps[i][1]), ok: i === steps.length - 1 ? 'Start' : 'Next', cancel: 'Skip tour' });
      if (!r) break;
    }
    this.app.prefs.firstRun = false;
    this.app.savePrefs();
  }
}

export { colName, keyRow, keyCol, MAX_ROWS, CellError };
