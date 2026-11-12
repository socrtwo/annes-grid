/**
 * Number format engine: renders values with spreadsheet format codes
 * (General, 0.00, #,##0, 0%, 0.00E+00, # ?/?, dates and times, [Red], [>=100], "text", @, _x, *x).
 */
import { CellError, type Scalar } from './types';
import { serialToParts, MONTHS, DAYS } from './dates';

export interface Formatted { text: string; color?: string; fill?: string; isNumber?: boolean }

interface Section {
  tokens: Tok[];
  cond?: { op: string; v: number };
  color?: string;
  kind: 'number' | 'date' | 'text' | 'general';
  hasNeg?: boolean;
}

type Tok =
  | { t: 'lit'; v: string }
  | { t: 'digit'; ch: '0' | '#' | '?' }
  | { t: 'dot' }
  | { t: 'comma' }
  | { t: 'pct' }
  | { t: 'exp'; sign: '+' | '-'; upper: boolean }
  | { t: 'slash' }
  | { t: 'text' }
  | { t: 'space'; ch: string }
  | { t: 'fill'; ch: string }
  | { t: 'date'; v: string }
  | { t: 'ampm'; v: string }
  | { t: 'elapsed'; v: 'h' | 'm' | 's'; n: number }
  | { t: 'general' };

const COLORS: Record<string, string> = {
  black: '#000000', white: '#ffffff', red: '#ff0000', green: '#008000', blue: '#0000ff', yellow: '#ffff00', magenta: '#ff00ff', cyan: '#00ffff',
};
const INDEXED = ['#000000', '#ffffff', '#ff0000', '#00ff00', '#0000ff', '#ffff00', '#ff00ff', '#00ffff', '#800000', '#008000', '#000080', '#808000', '#800080', '#008080', '#c0c0c0', '#808080', '#9999ff', '#993366', '#ffffcc', '#ccffff', '#660066', '#ff8080', '#0066cc', '#ccccff', '#000080', '#ff00ff', '#ffff00', '#00ffff', '#800080', '#800000', '#008080', '#0000ff', '#00ccff', '#ccffff', '#ccffcc', '#ffff99', '#99ccff', '#ff99cc', '#cc99ff', '#ffcc99', '#3366ff', '#33cccc', '#99cc00', '#ffcc00', '#ff9900', '#ff6600', '#666699', '#969696', '#003366', '#339966', '#003300', '#333300', '#993300', '#993366', '#333399', '#333333'];

const cache = new Map<string, Section[]>();

/** Built-in format ids used by OOXML. */
export const BUILTIN_FORMATS: Record<number, string> = {
  0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??',
  14: 'm/d/yyyy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss', 22: 'm/d/yyyy h:mm',
  37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)', 39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)',
  44: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)', 45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mm:ss.0', 48: '##0.0E+0', 49: '@',
  41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)', 42: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)', 43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
};

export const FORMAT_CATEGORIES: { name: string; formats: string[] }[] = [
  { name: 'General', formats: ['General'] },
  { name: 'Number', formats: ['0', '0.00', '#,##0', '#,##0.00', '#,##0;-#,##0', '#,##0.00;[Red]-#,##0.00', '#,##0.00_);(#,##0.00)', '#,##0.00_);[Red](#,##0.00)'] },
  { name: 'Currency', formats: ['$#,##0', '$#,##0.00', '$#,##0.00;[Red]$#,##0.00', '$#,##0.00_);($#,##0.00)', '$#,##0.00_);[Red]($#,##0.00)', '[$€-x-euro2] #,##0.00', '[$£-809]#,##0.00', '[$¥-411]#,##0'] },
  { name: 'Accounting', formats: ['_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)', '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)', '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)'] },
  { name: 'Date', formats: ['m/d/yyyy', 'dddd, mmmm d, yyyy', 'yyyy-mm-dd', 'd-mmm', 'd-mmm-yy', 'dd-mmm-yy', 'mmm-yy', 'mmmm-yy', 'mmmm d, yyyy', 'm/d/yy', 'mm/dd/yy', 'd/m/yyyy', 'dd/mm/yyyy', 'mmmmm', 'mmmmm-yy', 'm/d/yyyy h:mm', 'm/d/yy h:mm AM/PM'] },
  { name: 'Time', formats: ['h:mm', 'h:mm:ss', 'h:mm AM/PM', 'h:mm:ss AM/PM', 'mm:ss', 'mm:ss.0', '[h]:mm:ss', 'm/d/yyyy h:mm'] },
  { name: 'Percentage', formats: ['0%', '0.0%', '0.00%', '0.000%'] },
  { name: 'Fraction', formats: ['# ?/?', '# ??/??', '# ???/???', '# ?/2', '# ?/4', '# ?/8', '# ??/16', '# ?/10', '# ??/100'] },
  { name: 'Scientific', formats: ['0.00E+00', '0.0E+00', '0E+00', '##0.0E+0'] },
  { name: 'Text', formats: ['@'] },
  { name: 'Special', formats: ['00000', '00000-0000', '[<=9999999]###-####;(###) ###-####', '000-00-0000'] },
];

function parseFormat(fmt: string): Section[] {
  const hit = cache.get(fmt);
  if (hit) return hit;
  const raw = splitSections(fmt);
  // An empty section among several (0;-0;;@) shows nothing; only a wholly empty format is General.
  const sections = raw.map((src) => (raw.length > 1 && src.trim() === '' ? { tokens: [], kind: 'number' } as Section : parseSection(src)));
  if (cache.size > 2000) cache.clear();
  cache.set(fmt, sections);
  return sections;
}

function splitSections(fmt: string): string[] {
  const out: string[] = [];
  let cur = '';
  let inQ = false, inB = false;
  for (let i = 0; i < fmt.length; i++) {
    const ch = fmt[i];
    if (ch === '"' && !inB) inQ = !inQ;
    else if (ch === '\\' && !inQ) { cur += ch + (fmt[i + 1] ?? ''); i++; continue; }
    else if (ch === '[' && !inQ) inB = true;
    else if (ch === ']' && !inQ) inB = false;
    else if (ch === ';' && !inQ && !inB) { out.push(cur); cur = ''; continue; }
    cur += ch;
  }
  out.push(cur);
  return out;
}

function parseSection(src: string): Section {
  const tokens: Tok[] = [];
  const sec: Section = { tokens, kind: 'number' };
  let i = 0;
  const lower = src.toLowerCase();
  if (lower.trim() === 'general' || lower.trim() === '') { sec.kind = 'general'; tokens.push({ t: 'general' }); return sec; }
  let hasDate = false, hasText = false, hasDigit = false, hasGeneral = false;
  while (i < src.length) {
    const ch = src[i];
    const rest = lower.slice(i);
    if (ch === '"') {
      const j = src.indexOf('"', i + 1);
      const end = j < 0 ? src.length : j;
      tokens.push({ t: 'lit', v: src.slice(i + 1, end) });
      i = end + 1; continue;
    }
    if (ch === '\\') { tokens.push({ t: 'lit', v: src[i + 1] ?? '' }); i += 2; continue; }
    if (ch === '_') { tokens.push({ t: 'space', ch: src[i + 1] ?? ' ' }); i += 2; continue; }
    if (ch === '*') { tokens.push({ t: 'fill', ch: src[i + 1] ?? ' ' }); i += 2; continue; }
    if (ch === '[') {
      const j = src.indexOf(']', i);
      const inner = src.slice(i + 1, j < 0 ? src.length : j);
      i = j < 0 ? src.length : j + 1;
      const il = inner.toLowerCase();
      let m: RegExpExecArray | null;
      if (COLORS[il]) sec.color = COLORS[il];
      else if ((m = /^color\s*(\d+)$/.exec(il))) sec.color = INDEXED[(+m[1] - 1) % INDEXED.length];
      else if ((m = /^(<=|>=|<>|<|>|=)\s*(-?[\d.]+(?:e[+-]?\d+)?)$/.exec(il))) sec.cond = { op: m[1], v: +m[2] };
      else if (/^(h+|m+|s+)$/.test(il)) { tokens.push({ t: 'elapsed', v: il[0] as 'h' | 'm' | 's', n: il.length }); hasDate = true; }
      else if (inner.startsWith('$')) { const sym = inner.slice(1).split('-')[0]; if (sym) tokens.push({ t: 'lit', v: sym }); }
      continue;
    }
    // "General" with literal text around it, e.g. \(General\) written by Excel 4.0.
    if (rest.startsWith('general')) { tokens.push({ t: 'general' }); i += 7; hasGeneral = true; continue; }
    if (rest.startsWith('am/pm')) { tokens.push({ t: 'ampm', v: src.slice(i, i + 5) }); i += 5; hasDate = true; continue; }
    if (rest.startsWith('a/p')) { tokens.push({ t: 'ampm', v: src.slice(i, i + 3) }); i += 3; hasDate = true; continue; }
    if ('ymdhsge'.includes(rest[0]) && /^(y+|m+|d+|h+|s+|e+|g+)/.test(rest) && !(rest[0] === 'e' && /^e[+-]/.test(rest))) {
      const m = /^(y+|m+|d+|h+|s+|e+|g+)/.exec(rest)!;
      tokens.push({ t: 'date', v: m[1] });
      i += m[1].length; hasDate = true; continue;
    }
    if (ch === '0' || ch === '#' || ch === '?') { tokens.push({ t: 'digit', ch }); i++; hasDigit = true; continue; }
    if (ch === '.') {
      // fractional seconds: ss.0
      if (hasDate && /^\.0+/.test(rest)) { const m = /^\.0+/.exec(rest)!; tokens.push({ t: 'date', v: m[0] }); i += m[0].length; continue; }
      tokens.push({ t: 'dot' }); i++; continue;
    }
    if (ch === ',') { tokens.push({ t: 'comma' }); i++; continue; }
    if (ch === '%') { tokens.push({ t: 'pct' }); i++; continue; }
    if ((ch === 'E' || ch === 'e') && (src[i + 1] === '+' || src[i + 1] === '-')) { tokens.push({ t: 'exp', sign: src[i + 1] as '+' | '-', upper: ch === 'E' }); i += 2; continue; }
    if (ch === '/' && hasDigit && !hasDate) { tokens.push({ t: 'slash' }); i++; continue; }
    if (ch === '@') { tokens.push({ t: 'text' }); i++; hasText = true; continue; }
    tokens.push({ t: 'lit', v: ch }); i++;
  }
  if (hasDate) sec.kind = 'date';
  else if (hasGeneral && !hasDigit) sec.kind = 'general';
  else if (hasText && !hasDigit) sec.kind = 'text';
  sec.hasNeg = tokens.some((t) => t.t === 'lit' && /[-(]/.test(t.v));
  // Resolve m/mm as minutes when after h or before s.
  if (hasDate) {
    for (let k = 0; k < tokens.length; k++) {
      const t = tokens[k];
      if (t.t !== 'date' || !/^m{1,2}$/.test(t.v)) continue;
      let prev = k - 1; while (prev >= 0 && tokens[prev].t !== 'date' && tokens[prev].t !== 'elapsed') prev--;
      let next = k + 1; while (next < tokens.length && tokens[next].t !== 'date' && tokens[next].t !== 'elapsed') next++;
      const pv = prev >= 0 ? (tokens[prev] as any).v : '';
      const nv = next < tokens.length ? (tokens[next] as any).v : '';
      if (/^h/.test(pv) || /^s/.test(nv)) tokens[k] = { t: 'date', v: t.v === 'm' ? 'M' : 'MM' };
    }
  }
  return sec;
}

function condMatch(c: { op: string; v: number }, x: number): boolean {
  switch (c.op) {
    case '<': return x < c.v; case '<=': return x <= c.v; case '>': return x > c.v; case '>=': return x >= c.v;
    case '=': return x === c.v; case '<>': return x !== c.v;
  }
  return false;
}

/** Picks the section for a number; returns whether the sign should be stripped. */
function pickSection(secs: Section[], x: number): { sec: Section; abs: boolean } {
  const numSecs = secs.filter((s) => s.kind !== 'text' || secs.length === 1);
  if (secs.some((s) => s.cond)) {
    for (let i = 0; i < Math.min(2, secs.length); i++) if (secs[i].cond && condMatch(secs[i].cond!, x)) return { sec: secs[i], abs: i > 0 && x < 0 && !!secs[i].hasNeg };
    const other = secs.find((s, i) => i >= 1 && !s.cond && s.kind !== 'text') || secs[secs.length - 1];
    return { sec: other, abs: x < 0 && secs.indexOf(other) > 0 };
  }
  if (x > 0 || numSecs.length === 1) return { sec: secs[0], abs: false };
  if (x < 0) return { sec: secs[1] || secs[0], abs: !!secs[1] };
  return { sec: secs[2] || secs[0], abs: false };
}

export function isDateFormat(fmt: string | undefined): boolean {
  if (!fmt) return false;
  return parseFormat(fmt).some((s) => s.kind === 'date');
}

export function isTextFormat(fmt: string | undefined): boolean {
  return fmt === '@';
}

export function isPercentFormat(fmt: string | undefined): boolean {
  return !!fmt && parseFormat(fmt)[0]?.tokens.some((t) => t.t === 'pct');
}

/** Formats a scalar for display. width is the column width in characters for General, when known. */
export function formatValue(v: Scalar, fmt: string | undefined, opts: { date1904?: boolean; width?: number } = {}): Formatted {
  if (v == null) return { text: '' };
  if (v instanceof CellError) return { text: v.code };
  if (typeof v === 'boolean') return { text: v ? 'TRUE' : 'FALSE' };
  const f = fmt || 'General';
  const secs = parseFormat(f);
  if (typeof v === 'string') {
    const ts = secs.length >= 4 ? secs[3] : secs.find((s) => s.kind === 'text');
    if (!ts) return { text: v };
    return { text: renderText(ts, v), color: ts.color };
  }
  if (!Number.isFinite(v)) return { text: '#NUM!' };
  const { sec, abs } = pickSection(secs, v);
  const x = abs ? Math.abs(v) : v;
  let text: string;
  if (sec.kind === 'general') text = sec.tokens.length > 1 ? renderText(sec, formatGeneral(x, opts.width)) : formatGeneral(x, opts.width);
  else if (sec.kind === 'date') {
    if (x < 0 || x > 2958465.99999) return { text: '#'.repeat(Math.max(1, Math.min(opts.width ?? 8, 20))), isNumber: true };
    text = renderDate(sec, x, !!opts.date1904);
  } else if (sec.kind === 'text') text = renderText(sec, formatGeneral(v));
  else text = renderNumber(sec, x, !abs && x < 0 && secs.indexOf(sec) === 0);
  const fillTok = sec.tokens.find((t) => t.t === 'fill') as { ch: string } | undefined;
  return { text, color: sec.color, fill: fillTok?.ch, isNumber: true };
}

function renderText(sec: Section, s: string): string {
  let out = '';
  for (const t of sec.tokens) {
    if (t.t === 'text') out += s;
    else if (t.t === 'lit') out += t.v;
    else if (t.t === 'space') out += ' ';
    else if (t.t === 'general') out += s;
  }
  return out;
}

/** General format: up to 11 significant characters, switching to scientific for very large/small. */
export function formatGeneral(x: number, width?: number): string {
  if (x === 0) return '0';
  const maxLen = Math.max(5, Math.min(11, width ?? 11));
  const ax = Math.abs(x);
  if (ax >= 1e11 || ax < 1e-9) return sci(x, Math.max(0, maxLen - 6));
  // Ten significant digits, except that an 11-digit whole number is shown in full.
  let s = String(+x.toPrecision(ax >= 1e10 ? 11 : 10));
  if (s.includes('e')) s = sci(x, 5);
  if (s.replace('-', '').length > maxLen) {
    const intLen = Math.floor(Math.log10(ax)) + 1;
    const dec = Math.max(0, maxLen - Math.max(intLen, 1) - 1 - (x < 0 ? 1 : 0));
    s = String(+x.toFixed(Math.min(dec, 10)));
    if (s.replace('-', '').length > maxLen && intLen > maxLen) s = sci(x, Math.max(0, maxLen - 6));
  }
  return s;
}

function sci(x: number, decimals: number): string {
  const s = x.toExponential(Math.min(20, decimals));
  const [m, e] = s.split('e');
  const n = +e;
  const mant = decimals > 0 ? m.replace(/\.?0+$/, '') : m;
  return `${mant}E${n < 0 ? '-' : '+'}${String(Math.abs(n)).padStart(2, '0')}`;
}

function renderNumber(sec: Section, x: number, showMinus: boolean): string {
  const toks = sec.tokens;
  let pctCount = 0;
  for (const t of toks) if (t.t === 'pct') pctCount++;
  let val = Math.abs(x) * Math.pow(100, pctCount);
  // Trailing commas scale by 1000.
  const lastDigit = findLastIndex(toks, (t) => t.t === 'digit');
  let scale = 0;
  for (let k = lastDigit + 1; k < toks.length && toks[k].t === 'comma'; k++) scale++;
  if (scale) val /= Math.pow(1000, scale);
  const expIdx = toks.findIndex((t) => t.t === 'exp');
  const slashIdx = toks.findIndex((t) => t.t === 'slash');
  let body: string;
  if (slashIdx >= 0) body = renderFraction(toks, val, slashIdx);
  else if (expIdx >= 0) body = renderSci(toks, val, expIdx);
  else body = renderFixed(toks, val, lastDigit, scale);
  const isZero = /^[^1-9]*$/.test(body);
  return (showMinus && !isZero ? '-' : '') + body;
}

function findLastIndex<T>(a: T[], f: (x: T) => boolean): number { for (let i = a.length - 1; i >= 0; i--) if (f(a[i])) return i; return -1; }

function renderFixed(toks: Tok[], val: number, lastDigit: number, scaleCommas: number): string {
  const dotIdx = toks.findIndex((t) => t.t === 'dot');
  const intToks = dotIdx >= 0 ? toks.slice(0, dotIdx) : toks;
  const fracToks = dotIdx >= 0 ? toks.slice(dotIdx + 1) : [];
  const fracDigits = fracToks.filter((t) => t.t === 'digit').length;
  const grouping = intToks.some((t, k) => t.t === 'comma' && k < (dotIdx >= 0 ? dotIdx : lastDigit + 1) - 0 && intToks.slice(k + 1).some((u) => u.t === 'digit'));
  void scaleCommas;
  const rounded = roundHalfAway(val, fracDigits);
  let [ip, fp = ''] = rounded.toFixed(fracDigits).split('.');
  if (ip === '0') ip = '';
  // A section with no digit placeholders and no decimal point (e.g. "pos";"neg" or [Red]"x") shows only its literals.
  if (dotIdx < 0 && !intToks.some((t) => t.t === 'digit')) return intToks.map(tokText).join('');
  const out = layoutInteger(intToks, ip, grouping);
  if (dotIdx < 0) return out;
  // Fraction part
  let fout = '';
  let fi = 0;
  const fpTrim = trimFrac(fp, fracToks);
  for (const t of fracToks) {
    if (t.t === 'digit') { const d = fpTrim[fi++]; fout += d ?? (t.ch === '0' ? '0' : t.ch === '?' ? ' ' : ''); }
    else fout += tokText(t);
  }
  return out + '.' + fout;
}

/** Places integer digits into placeholders, right-aligned; surplus digits go to the leftmost placeholder. */
function layoutInteger(intToks: Tok[], ip: string, grouping: boolean): string {
  const ph = intToks.map((t, k) => (t.t === 'digit' ? k : -1)).filter((k) => k >= 0);
  if (!ph.length) return intToks.map(tokText).join('') + (ip && ip !== '0' ? ip : '');
  const first = ph[0], last = ph[ph.length - 1];
  const contiguous = intToks.slice(first, last + 1).every((t) => t.t === 'digit' || t.t === 'comma');
  if (contiguous) {
    const chars = ph.map((k) => (intToks[k] as { ch: string }).ch);
    const firstZero = chars.indexOf('0');
    const minLen = firstZero < 0 ? 0 : chars.length - firstZero;
    let digits = ip.length < minLen ? ip.padStart(minLen, '0') : ip;
    const padQ = Math.max(0, chars.length - Math.max(digits.length, minLen) - (chars.slice(0, chars.length - digits.length).filter((c) => c !== '?').length));
    if (grouping && digits) digits = digits.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    const qPad = chars.slice(0, Math.max(0, chars.length - Math.max(ip.length, minLen))).filter((c) => c === '?').length;
    void padQ;
    return intToks.slice(0, first).map(tokText).join('') + ' '.repeat(qPad) + digits + intToks.slice(last + 1).map(tokText).join('');
  }
  let di = ip.length;
  const parts: string[] = new Array(intToks.length).fill('');
  for (let j = ph.length - 1; j >= 0; j--) {
    const k = ph[j];
    const ch = (intToks[k] as { ch: string }).ch;
    if (j === 0 && di > 0) { parts[k] = ip.slice(0, di); di = 0; }
    else if (di > 0) { parts[k] = ip[di - 1]; di--; }
    else parts[k] = ch === '0' ? '0' : ch === '?' ? ' ' : '';
  }
  return intToks.map((t, k) => (t.t === 'digit' ? parts[k] : tokText(t))).join('');
}

function trimFrac(fp: string, fracToks: Tok[]): string {
  const ph = fracToks.filter((t) => t.t === 'digit') as { ch: string }[];
  let s = fp;
  for (let k = ph.length - 1; k >= 0 && s.length > k; k--) {
    if (ph[k].ch !== '0' && s[k] === '0') s = s.slice(0, k); else break;
  }
  return s;
}

function tokText(t: Tok): string {
  switch (t.t) {
    case 'lit': return t.v;
    case 'space': return ' ';
    case 'pct': return '%';
    case 'comma': return '';
    case 'fill': return '';
    default: return '';
  }
}

function roundHalfAway(x: number, d: number): number {
  const m = Math.pow(10, d);
  return Math.round((x * m) * (1 + Number.EPSILON)) / m;
}

function renderSci(toks: Tok[], val: number, expIdx: number): string {
  const mant = toks.slice(0, expIdx);
  const expT = toks[expIdx] as { sign: string; upper: boolean };
  const expDigits = toks.slice(expIdx + 1).filter((t) => t.t === 'digit').length || 1;
  const dotIdx = mant.findIndex((t) => t.t === 'dot');
  const intPh = (dotIdx >= 0 ? mant.slice(0, dotIdx) : mant).filter((t) => t.t === 'digit').length || 1;
  const fracPh = dotIdx >= 0 ? mant.slice(dotIdx + 1).filter((t) => t.t === 'digit').length : 0;
  let e = val === 0 ? 0 : Math.floor(Math.log10(val));
  // Engineering-style when intPh > 1: exponent is a multiple of intPh.
  if (intPh > 1) e = Math.floor(e / intPh) * intPh;
  let m = val / Math.pow(10, e);
  m = roundHalfAway(m, fracPh);
  if (m >= Math.pow(10, intPh)) { m /= Math.pow(10, intPh > 1 ? intPh : 1); e += intPh > 1 ? intPh : 1; }
  let ms = m.toFixed(fracPh);
  if (fracPh && mant.slice(dotIdx + 1).every((t) => t.t !== 'digit' || (t as any).ch === '#')) ms = ms.replace(/\.?0+$/, '');
  const pre = mant.filter((t) => t.t === 'lit').map((t) => (t as any).v).join('');
  const es = String(Math.abs(e)).padStart(expDigits, '0');
  return `${pre}${ms}${expT.upper ? 'E' : 'e'}${e < 0 ? '-' : expT.sign === '+' ? '+' : ''}${es}`;
}

function renderFraction(toks: Tok[], val: number, slashIdx: number): string {
  const before = toks.slice(0, slashIdx);
  const after = toks.slice(slashIdx + 1);
  // Split "# ?/?" into integer part, numerator and denominator placeholders.
  const digitPos = before.map((t, k) => (t.t === 'digit' ? k : -1)).filter((k) => k >= 0);
  let numStart = digitPos.length ? digitPos[digitPos.length - 1] : 0;
  while (numStart > 0 && before[numStart - 1].t === 'digit') numStart--;
  const intToks = before.slice(0, numStart);
  const hasInt = intToks.some((t) => t.t === 'digit');
  const numPh = before.slice(numStart).filter((t) => t.t === 'digit').length;
  const denToks = after.filter((t) => t.t === 'digit' || (t.t === 'lit' && /\d/.test(t.v)));
  const fixedDen = after.filter((t) => t.t === 'lit' && /^\d+$/.test(t.v)).map((t) => (t as any).v).join('');
  const denPh = denToks.filter((t) => t.t === 'digit').length || 1;
  const whole = hasInt ? Math.floor(val) : 0;
  const frac = val - whole;
  let num: number, den: number;
  if (fixedDen) { den = +fixedDen; num = Math.round(frac * den); }
  else [num, den] = bestFraction(frac, Math.pow(10, denPh) - 1);
  let w = whole;
  if (den && num === den) { w++; num = 0; }
  const intStr = hasInt ? (w || num === 0 ? String(w) : '') : '';
  if (num === 0 && hasInt) return (intStr || '0') + ' '.repeat(numPh + 1 + denPh + 1);
  const numStr = String(hasInt ? num : num + w * den).padStart(numPh, ' ');
  const denStr = String(den).padEnd(denPh, ' ');
  return `${intStr}${intStr ? ' ' : ''}${numStr}/${denStr}`;
}

function bestFraction(x: number, maxDen: number): [number, number] {
  let best = [0, 1], err = Infinity;
  for (let d = 1; d <= maxDen; d++) {
    const n = Math.round(x * d);
    const e = Math.abs(x - n / d);
    if (e < err - 1e-12) { err = e; best = [n, d]; if (e < 1e-10) break; }
  }
  return best as [number, number];
}

function renderDate(sec: Section, x: number, date1904: boolean): string {
  const hasSubSec = sec.tokens.find((t) => t.t === 'date' && t.v.startsWith('.'));
  const p = serialToParts(hasSubSec ? x : Math.round(x * 86400) / 86400, date1904);
  const ampm = sec.tokens.some((t) => t.t === 'ampm');
  let out = '';
  for (const t of sec.tokens) {
    switch (t.t) {
      case 'lit': out += t.v; break;
      case 'space': out += ' '; break;
      case 'ampm': {
        const pm = p.hh >= 12;
        if (t.v.length === 3) out += (pm ? 'P' : 'A')[t.v[0] === 'a' ? 'toLowerCase' : 'toUpperCase']();
        else out += pm ? 'PM' : 'AM';
        break;
      }
      case 'elapsed': {
        const totalSec = Math.round(x * 86400);
        const v = t.v === 'h' ? Math.floor(totalSec / 3600) : t.v === 'm' ? Math.floor(totalSec / 60) : totalSec;
        out += String(v).padStart(t.n, '0');
        break;
      }
      case 'date': out += dateTok(t.v, p, ampm); break;
      case 'digit': out += t.ch === '0' ? '0' : ''; break;
      case 'comma': out += ','; break;
      case 'dot': out += '.'; break;
      default: out += tokText(t);
    }
  }
  return out;
}

function dateTok(v: string, p: ReturnType<typeof serialToParts>, ampm: boolean): string {
  if (v.startsWith('.')) { const n = v.length - 1; return '.' + String(Math.floor(p.ms / Math.pow(10, 3 - n))).padStart(n, '0').slice(0, n); }
  switch (v) {
    case 'yy': case 'y': return String(p.y % 100).padStart(2, '0');
    case 'yyy': case 'yyyy': return String(p.y);
    case 'm': return String(p.m);
    case 'mm': return String(p.m).padStart(2, '0');
    case 'mmm': return MONTHS[p.m - 1].slice(0, 3);
    case 'mmmm': return MONTHS[p.m - 1];
    case 'mmmmm': return MONTHS[p.m - 1][0];
    case 'M': return String(p.mm);
    case 'MM': return String(p.mm).padStart(2, '0');
    case 'd': return String(p.d);
    case 'dd': return String(p.d).padStart(2, '0');
    case 'ddd': return DAYS[p.dow].slice(0, 3);
    case 'dddd': return DAYS[p.dow];
    case 'h': { const hh = ampm ? ((p.hh + 11) % 12) + 1 : p.hh; return String(hh); }
    case 'hh': { const hh = ampm ? ((p.hh + 11) % 12) + 1 : p.hh; return String(hh).padStart(2, '0'); }
    case 's': return String(p.ss);
    case 'ss': return String(p.ss).padStart(2, '0');
    case 'e': case 'ee': return String(p.y);
    case 'g': case 'gg': case 'ggg': return '';
  }
  if (/^y+$/.test(v)) return String(p.y);
  if (/^m{6,}$/.test(v)) return MONTHS[p.m - 1];
  if (/^d{5,}$/.test(v)) return DAYS[p.dow];
  return v;
}

/** Decimal places shown by a format (for Increase/Decrease Decimal). */
export function formatDecimals(fmt: string | undefined): number {
  if (!fmt || fmt === 'General') return -1;
  const sec = parseFormat(fmt)[0];
  const dot = sec.tokens.findIndex((t) => t.t === 'dot');
  if (dot < 0) return 0;
  let n = 0;
  for (let k = dot + 1; k < sec.tokens.length && sec.tokens[k].t === 'digit'; k++) n++;
  return n;
}

/** Changes the number of decimals in every numeric section of a format. */
export function withDecimals(fmt: string | undefined, delta: number, current?: number): string {
  if (!fmt || fmt === 'General') {
    const d = Math.max(0, (current ?? 0) + delta);
    return d ? '0.' + '0'.repeat(d) : '0';
  }
  return splitSections(fmt).map((s) => {
    if (/[ymdhs]/i.test(s.replace(/"[^"]*"|\[[^\]]*\]/g, '')) || s === '@') return s;
    if (delta > 0) {
      if (/\d|#|\?/.test(s) && /\./.test(s)) return s.replace(/(\.[0#?]*)/, `$1${'0'.repeat(delta)}`);
      return s.replace(/([0#?])(?![0#?,]*[0#?])/, `$1.${'0'.repeat(delta)}`);
    }
    let out = s;
    for (let i = 0; i < -delta; i++) out = out.replace(/\.([0#?]*)[0#?]/, (_m, a) => (a ? '.' + a : ''));
    return out;
  }).join(';');
}

/** Parses typed input (numbers with currency, %, thousands, fractions, scientific). */
export function parseNumberInput(text: string): { value: number; format?: string } | null {
  let t = text.trim();
  if (!t) return null;
  let neg = false;
  if (/^\(.*\)$/.test(t)) { neg = true; t = t.slice(1, -1).trim(); }
  if (t.startsWith('-')) { neg = !neg; t = t.slice(1).trim(); } else if (t.startsWith('+')) t = t.slice(1).trim();
  let currency = '';
  const cm = /^([$€£¥₹])\s*/.exec(t) || /\s*([$€£¥₹])$/.exec(t);
  if (cm) { currency = cm[1]; t = t.replace(cm[0], '').trim(); }
  if (t.startsWith('-') && !neg) { neg = true; t = t.slice(1); }
  let pct = false;
  if (t.endsWith('%')) { pct = true; t = t.slice(0, -1).trim(); }
  let m: RegExpExecArray | null;
  // Fractions: "1 1/2", "3/4" (only when unambiguous; 3/4 alone is a date in most locales so require a whole part)
  if ((m = /^(\d+)\s+(\d+)\/(\d+)$/.exec(t))) {
    const v = +m[1] + +m[2] / +m[3];
    if (!+m[3]) return null;
    return { value: neg ? -v : v, format: m[3].length > 1 ? '# ??/??' : '# ?/?' };
  }
  const grouped = /^\d{1,3}(,\d{3})+(\.\d*)?$/.test(t);
  if (!grouped && !/^(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(t)) return null;
  const v = +t.replace(/,/g, '');
  if (!Number.isFinite(v)) return null;
  const value = (neg ? -v : v) / (pct ? 100 : 1);
  const dec = (t.split('.')[1] || '').replace(/[eE].*/, '').length;
  let format: string | undefined;
  if (currency) format = `${currency}#,##0${dec ? '.00' : ''}`;
  else if (pct) format = dec ? '0.' + '0'.repeat(Math.min(dec, 4)) + '%' : '0%';
  else if (grouped) format = dec ? '#,##0.00' : '#,##0';
  else if (/[eE]/.test(t)) format = '0.00E+00';
  if (currency && neg && text.trim().startsWith('(')) format = `${currency}#,##0${dec ? '.00' : ''}_);(${currency}#,##0${dec ? '.00' : ''})`;
  return { value, format };
}
