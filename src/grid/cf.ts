/**
 * Conditional formatting evaluation: per-rule statistics are computed once per recalc and cached; each
 * visible cell then resolves to style overrides, a color-scale fill, a data bar or an icon.
 */
import type { Sheet } from '../model/workbook';
import type { CFRule, CellStyle, Scalar } from '../model/types';
import { CellError } from '../model/types';
import type { Engine } from '../calc/engine';
import { compare, criteria, toScalar } from '../calc/values';
import { serialToParts, partsToSerial, jsDateToSerial } from '../model/dates';

export interface CFResult { style?: Partial<CellStyle>; fill?: string; bar?: { pct: number; color: string; gradient: boolean; neg?: boolean; axis?: number }; icon?: { set: string; index: number }; hideValue?: boolean }

interface RuleStats { nums: number[]; min: number; max: number; avg: number; std: number; counts?: Map<string, number>; threshold?: number; sorted: number[] }

export const ICON_SETS: Record<string, { n: number; icons: string[]; colors: string[] }> = {
  '3Arrows': { n: 3, icons: ['arrow_downward', 'arrow_forward', 'arrow_upward'], colors: ['#c00000', '#e3a21a', '#2e7d32'] },
  '3ArrowsGray': { n: 3, icons: ['arrow_downward', 'arrow_forward', 'arrow_upward'], colors: ['#808080', '#808080', '#808080'] },
  '3Flags': { n: 3, icons: ['flag', 'flag', 'flag'], colors: ['#c00000', '#e3a21a', '#2e7d32'] },
  '3TrafficLights1': { n: 3, icons: ['circle', 'circle', 'circle'], colors: ['#c00000', '#e3a21a', '#2e7d32'] },
  '3Signs': { n: 3, icons: ['change_history', 'square', 'circle'], colors: ['#c00000', '#e3a21a', '#2e7d32'] },
  '3Symbols': { n: 3, icons: ['cancel', 'error', 'check_circle'], colors: ['#c00000', '#e3a21a', '#2e7d32'] },
  '3Symbols2': { n: 3, icons: ['close', 'priority_high', 'check'], colors: ['#c00000', '#e3a21a', '#2e7d32'] },
  '3Stars': { n: 3, icons: ['star_outline', 'star_half', 'star'], colors: ['#e3a21a', '#e3a21a', '#e3a21a'] },
  '3Triangles': { n: 3, icons: ['arrow_drop_down', 'remove', 'arrow_drop_up'], colors: ['#c00000', '#e3a21a', '#2e7d32'] },
  '4Arrows': { n: 4, icons: ['arrow_downward', 'south_east', 'north_east', 'arrow_upward'], colors: ['#c00000', '#e3a21a', '#e3a21a', '#2e7d32'] },
  '4RedToBlack': { n: 4, icons: ['circle', 'circle', 'circle', 'circle'], colors: ['#000000', '#808080', '#ff8080', '#c00000'] },
  '4Rating': { n: 4, icons: ['signal_cellular_alt_1_bar', 'signal_cellular_alt_2_bar', 'signal_cellular_alt', 'signal_cellular_4_bar'], colors: ['#2f6fb0', '#2f6fb0', '#2f6fb0', '#2f6fb0'] },
  '4TrafficLights': { n: 4, icons: ['circle', 'circle', 'circle', 'circle'], colors: ['#000000', '#c00000', '#e3a21a', '#2e7d32'] },
  '5Arrows': { n: 5, icons: ['arrow_downward', 'south_east', 'arrow_forward', 'north_east', 'arrow_upward'], colors: ['#c00000', '#e3a21a', '#e3a21a', '#e3a21a', '#2e7d32'] },
  '5Quarters': { n: 5, icons: ['radio_button_unchecked', 'clock_loader_10', 'clock_loader_40', 'clock_loader_60', 'circle'], colors: ['#404040', '#404040', '#404040', '#404040', '#404040'] },
  '5Rating': { n: 5, icons: ['signal_cellular_0_bar', 'signal_cellular_alt_1_bar', 'signal_cellular_alt_2_bar', 'signal_cellular_alt', 'signal_cellular_4_bar'], colors: ['#2f6fb0', '#2f6fb0', '#2f6fb0', '#2f6fb0', '#2f6fb0'] },
  '5Boxes': { n: 5, icons: ['check_box_outline_blank', 'grid_view', 'grid_view', 'grid_view', 'square'], colors: ['#2f6fb0', '#2f6fb0', '#2f6fb0', '#2f6fb0', '#2f6fb0'] },
};

export class CFEvaluator {
  private stats = new Map<string, RuleStats>();
  private version = -1;

  constructor(private engine: Engine) {}

  invalidate(): void { this.version = -1; }

  private ensure(sheet: Sheet, version: number): void {
    if (this.version === version) return;
    this.stats.clear();
    this.version = version;
    void sheet;
  }

  private ruleStats(sheet: Sheet, rule: CFRule): RuleStats {
    const key = sheet.id + ':' + rule.id;
    let st = this.stats.get(key);
    if (st) return st;
    const nums: number[] = [];
    const counts = new Map<string, number>();
    const wantCounts = rule.type === 'duplicateValues' || rule.type === 'uniqueValues';
    for (const rg of rule.ranges) {
      const b = this.engine.usedBounds(sheet);
      const r2 = Math.min(rg.r2, b.r2), c2 = Math.min(rg.c2, b.c2);
      for (let r = rg.r1; r <= r2; r++) for (let c = rg.c1; c <= c2; c++) {
        const v = sheet.value(r, c);
        if (typeof v === 'number') nums.push(v);
        if (wantCounts && v != null && v !== '') { const k = typeof v === 'string' ? 's:' + v.toLowerCase() : 'n:' + String(v); counts.set(k, (counts.get(k) || 0) + 1); }
      }
    }
    const sorted = [...nums].sort((a, b) => a - b);
    const min = sorted[0] ?? 0, max = sorted[sorted.length - 1] ?? 0;
    const avg = nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : 0;
    const std = nums.length > 1 ? Math.sqrt(nums.reduce((a, b) => a + (b - avg) ** 2, 0) / (nums.length - 1)) : 0;
    st = { nums, min, max, avg, std, counts, sorted };
    if (rule.type === 'top10') {
      const n = rule.percent ? Math.max(1, Math.floor((sorted.length * (rule.rank ?? 10)) / 100)) : (rule.rank ?? 10);
      st.threshold = rule.bottom ? sorted[Math.min(n, sorted.length) - 1] : sorted[Math.max(0, sorted.length - n)];
    }
    this.stats.set(key, st);
    return st;
  }

  private stopValue(sheet: Sheet, rule: CFRule, st: RuleStats, i: number, fallback: 'min' | 'max' | 'mid'): number {
    const s = rule.stops?.[i];
    if (!s) return fallback === 'min' ? st.min : fallback === 'max' ? st.max : (st.min + st.max) / 2;
    switch (s.type) {
      case 'min': return st.min;
      case 'max': return st.max;
      case 'num': return Number(s.value ?? 0);
      case 'percent': return st.min + ((st.max - st.min) * Number(s.value ?? 50)) / 100;
      case 'percentile': return percentile(st.sorted, Number(s.value ?? 50) / 100);
      case 'formula': { const v = this.engine.evaluateScalar(String(s.value ?? '0'), sheet, rule.ranges[0].r1, rule.ranges[0].c1); return typeof v === 'number' ? v : 0; }
    }
    return 0;
  }

  /** Conditional formatting for one cell, merging all matching rules in priority order. */
  evaluate(sheet: Sheet, r: number, c: number, version: number): CFResult | null {
    if (!sheet.cf.length) return null;
    this.ensure(sheet, version);
    let out: CFResult | null = null;
    const rules = sheet.cf.length > 1 ? [...sheet.cf].sort((a, b) => (a.priority ?? 0) - (b.priority ?? 0)) : sheet.cf;
    for (const rule of rules) {
      const rg = rule.ranges.find((x) => r >= x.r1 && r <= x.r2 && c >= x.c1 && c <= x.c2);
      if (!rg) continue;
      const v = sheet.value(r, c);
      const res = this.applyRule(sheet, rule, r, c, v, rule.ranges[0]);
      if (!res) continue;
      out ||= {};
      if (res.style) out.style = { ...(res.style), ...(out.style || {}) };
      if (res.fill && !out.fill) out.fill = res.fill;
      if (res.bar && !out.bar) out.bar = res.bar;
      if (res.icon && !out.icon) out.icon = res.icon;
      if (res.hideValue) out.hideValue = true;
      if (rule.stopIfTrue) break;
    }
    return out;
  }

  private applyRule(sheet: Sheet, rule: CFRule, r: number, c: number, v: Scalar, first: { r1: number; c1: number }): CFResult | null {
    const st = rule.style ? { style: rule.style } : {};
    const base = { r: first.r1, c: first.c1 };
    const evalF = (i: number): Scalar => this.engine.evaluateScalar(rule.formulas?.[i] ?? '0', sheet, r, c, base);
    switch (rule.type) {
      case 'cellIs': {
        if (v == null || v instanceof CellError) return null;
        const a = evalF(0);
        const b = rule.formulas && rule.formulas.length > 1 ? evalF(1) : null;
        const cmp = (x: Scalar, y: Scalar) => compare(x, y);
        let ok = false;
        switch (rule.operator) {
          case 'between': ok = cmp(v, a) >= 0 && cmp(v, b) <= 0; break;
          case 'notBetween': ok = cmp(v, a) < 0 || cmp(v, b) > 0; break;
          case 'equal': ok = cmp(v, a) === 0; break;
          case 'notEqual': ok = cmp(v, a) !== 0; break;
          case 'greaterThan': ok = cmp(v, a) > 0; break;
          case 'lessThan': ok = cmp(v, a) < 0; break;
          case 'greaterThanOrEqual': ok = cmp(v, a) >= 0; break;
          case 'lessThanOrEqual': ok = cmp(v, a) <= 0; break;
        }
        return ok ? st : null;
      }
      case 'expression': {
        const x = evalF(0);
        return x === true || (typeof x === 'number' && x !== 0) ? st : null;
      }
      case 'containsText': return typeof v === 'string' && v.toLowerCase().includes((rule.text || '').toLowerCase()) ? st : null;
      case 'notContainsText': return !(typeof v === 'string' && v.toLowerCase().includes((rule.text || '').toLowerCase())) ? st : null;
      case 'beginsWith': return typeof v === 'string' && v.toLowerCase().startsWith((rule.text || '').toLowerCase()) ? st : null;
      case 'endsWith': return typeof v === 'string' && v.toLowerCase().endsWith((rule.text || '').toLowerCase()) ? st : null;
      case 'containsBlanks': return v == null || (typeof v === 'string' && !v.trim()) ? st : null;
      case 'notContainsBlanks': return !(v == null || (typeof v === 'string' && !v.trim())) ? st : null;
      case 'containsErrors': return v instanceof CellError ? st : null;
      case 'notContainsErrors': return !(v instanceof CellError) ? st : null;
      case 'duplicateValues': case 'uniqueValues': {
        if (v == null || v === '') return null;
        const s = this.ruleStats(sheet, rule);
        const k = typeof v === 'string' ? 's:' + v.toLowerCase() : 'n:' + String(v);
        const n = s.counts!.get(k) || 0;
        return (rule.type === 'duplicateValues' ? n > 1 : n === 1) ? st : null;
      }
      case 'top10': {
        if (typeof v !== 'number') return null;
        const s = this.ruleStats(sheet, rule);
        if (s.threshold == null) return null;
        return (rule.bottom ? v <= s.threshold : v >= s.threshold) ? st : null;
      }
      case 'aboveAverage': {
        if (typeof v !== 'number') return null;
        const s = this.ruleStats(sheet, rule);
        const sd = rule.stdDev ? s.std * rule.stdDev : 0;
        const above = rule.above !== false;
        const ok = above ? (rule.equalAverage ? v >= s.avg + sd : v > s.avg + sd) : (rule.equalAverage ? v <= s.avg - sd : v < s.avg - sd);
        return ok ? st : null;
      }
      case 'timePeriod': {
        if (typeof v !== 'number') return null;
        return inPeriod(v, rule.timePeriod || 'today', sheet.book.calc.date1904) ? st : null;
      }
      case 'colorScale': {
        if (typeof v !== 'number') return null;
        const s = this.ruleStats(sheet, rule);
        const cols = rule.colors?.length ? rule.colors : ['#f8696b', '#ffeb84', '#63be7b'];
        if (cols.length === 2) {
          const lo = this.stopValue(sheet, rule, s, 0, 'min'), hi = this.stopValue(sheet, rule, s, 1, 'max');
          return { fill: mix(cols[0], cols[1], hi === lo ? 0.5 : clamp01((v - lo) / (hi - lo))) };
        }
        const lo = this.stopValue(sheet, rule, s, 0, 'min'), mid = rule.stops?.[1] ? this.stopValue(sheet, rule, s, 1, 'mid') : percentile(s.sorted, 0.5), hi = this.stopValue(sheet, rule, s, 2, 'max');
        if (v <= mid) return { fill: mix(cols[0], cols[1], mid === lo ? 1 : clamp01((v - lo) / (mid - lo))) };
        return { fill: mix(cols[1], cols[2], hi === mid ? 0 : clamp01((v - mid) / (hi - mid))) };
      }
      case 'dataBar': {
        if (typeof v !== 'number') return null;
        const s = this.ruleStats(sheet, rule);
        const lo = Math.min(0, this.stopValue(sheet, rule, s, 0, 'min')), hi = Math.max(0, this.stopValue(sheet, rule, s, 1, 'max'));
        const span = hi - lo || 1;
        const axis = lo < 0 ? -lo / span : 0;
        const pct = clamp01(Math.abs(v) / span);
        return { bar: { pct, color: rule.barColor || '#638ec6', gradient: rule.gradientBar !== false, neg: v < 0, axis }, hideValue: rule.showValue === false };
      }
      case 'iconSet': {
        if (typeof v !== 'number') return null;
        const s = this.ruleStats(sheet, rule);
        const set = ICON_SETS[rule.iconSet || '3TrafficLights1'] || ICON_SETS['3TrafficLights1'];
        // thresholds default to equal percent bands
        let idx = 0;
        for (let i = 1; i < set.n; i++) {
          const stop = rule.stops?.[i];
          const th = stop ? this.stopValue(sheet, rule, s, i, 'mid') : s.min + ((s.max - s.min) * i) / set.n;
          if (v >= th) idx = i;
        }
        if (rule.reverseIcons) idx = set.n - 1 - idx;
        return { icon: { set: rule.iconSet || '3TrafficLights1', index: idx }, hideValue: rule.showValue === false };
      }
    }
    return null;
  }
}

function percentile(sorted: number[], p: number): number {
  if (!sorted.length) return 0;
  const x = (sorted.length - 1) * p, lo = Math.floor(x), hi = Math.ceil(x);
  return sorted[lo] + (sorted[hi] - sorted[lo]) * (x - lo);
}

function clamp01(x: number): number { return Math.max(0, Math.min(1, x)); }

export function mix(a: string, b: string, t: number): string {
  const pa = hex(a), pb = hex(b);
  const m = pa.map((x, i) => Math.round(x + (pb[i] - x) * t));
  return '#' + m.map((x) => x.toString(16).padStart(2, '0')).join('');
}

export function hex(c: string): number[] {
  let s = c.replace('#', '');
  if (s.length === 3) s = s.split('').map((x) => x + x).join('');
  return [0, 2, 4].map((i) => parseInt(s.slice(i, i + 2), 16) || 0);
}

function inPeriod(v: number, p: string, d1904: boolean): boolean {
  const today = Math.floor(jsDateToSerial(new Date(), d1904));
  const day = Math.floor(v);
  const tp = serialToParts(today, d1904);
  const weekStart = today - tp.dow;
  switch (p) {
    case 'today': return day === today;
    case 'yesterday': return day === today - 1;
    case 'tomorrow': return day === today + 1;
    case 'last7Days': return day > today - 7 && day <= today;
    case 'thisWeek': return day >= weekStart && day < weekStart + 7;
    case 'lastWeek': return day >= weekStart - 7 && day < weekStart;
    case 'nextWeek': return day >= weekStart + 7 && day < weekStart + 14;
    case 'thisMonth': { const s = partsToSerial(tp.y, tp.m, 1, d1904); return day >= s && day < partsToSerial(tp.y, tp.m + 1, 1, d1904); }
    case 'lastMonth': { const s = partsToSerial(tp.y, tp.m - 1, 1, d1904); return day >= s && day < partsToSerial(tp.y, tp.m, 1, d1904); }
    case 'nextMonth': { const s = partsToSerial(tp.y, tp.m + 1, 1, d1904); return day >= s && day < partsToSerial(tp.y, tp.m + 2, 1, d1904); }
  }
  return false;
}

void criteria; void toScalar;
