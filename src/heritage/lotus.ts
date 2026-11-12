/**
 * Lotus 1-2-3 formula entry.
 *
 * Translates 1-2-3 formulas (@functions, A1..B5 ranges, #AND# / #OR# / #NOT#, sheet letters such
 * as B:A1) into this app's formula language, keeping 1-2-3's meaning: its operator precedence
 * (unary minus below ^, & and the logical operators lowest), its zero-based offsets (@CHOOSE,
 * @INDEX, @VLOOKUP, @MID, @FIND...), its argument orders (@PMT, @PV, @FV, @IRR...) and its
 * treatment of labels in statistical ranges (@AVG counts labels as zero; @PUREAVG ignores them).
 *
 * The translator is a small Pratt parser over 1-2-3 tokens that emits a fully equivalent formula
 * with parentheses added wherever the two precedence tables disagree.
 */
import { colIndex, quoteSheet } from '../model/address';

type Tok =
  | { t: 'num'; v: string }
  | { t: 'str'; v: string }
  | { t: 'fn'; name: string }
  | { t: 'ref'; text: string; simple: boolean }
  | { t: 'name'; v: string }
  | { t: 'op'; v: string }
  | { t: 'lp' } | { t: 'rp' } | { t: 'sep' };

type LNode =
  | { k: 'num'; v: string }
  | { k: 'str'; v: string }
  | { k: 'ref'; text: string }
  | { k: 'name'; v: string }
  | { k: 'fn'; name: string; args: LNode[] }
  | { k: 'un'; op: string; a: LNode }
  | { k: 'bin'; op: string; a: LNode; b: LNode };

export interface LotusOptions {
  /** Workbook sheet names, so 1-2-3 sheet letters (A:, B:, ...) become sheet references. */
  sheetNames?: string[];
}

export class LotusError extends Error {}

const CELLP = String.raw`(?:\$?[A-Za-z]{1,2}:)?\$?[A-Za-z]{1,3}\$?\d{1,7}`;
const RE_REF = new RegExp(String.raw`^(${CELLP})(?:(?:\.{1,2}|:)(${CELLP}))?(?![A-Za-z0-9_(])`);

function refText(cell: string, o: LotusOptions): { sheet: string | null; cell: string } {
  const m = /^(\$?([A-Za-z]{1,2}):)?(.*)$/.exec(cell)!;
  if (!m[1]) return { sheet: null, cell: m[3].toUpperCase() };
  const idx = colIndex(m[2].toUpperCase());
  const name = o.sheetNames?.[idx] ?? m[2].toUpperCase();
  return { sheet: name, cell: m[3].toUpperCase() };
}

function tokenize(src: string, o: LotusOptions): Tok[] {
  const out: Tok[] = [];
  let i = 0;
  while (i < src.length) {
    const ch = src[i];
    const rest = src.slice(i);
    if (/\s/.test(ch)) { i++; continue; }
    if (ch === '"') {
      let j = i + 1, s = '';
      for (;;) {
        if (j >= src.length) throw new LotusError('Unterminated string');
        if (src[j] === '"') { if (src[j + 1] === '"') { s += '"'; j += 2; continue; } break; }
        s += src[j++];
      }
      out.push({ t: 'str', v: s });
      i = j + 1;
      continue;
    }
    if (rest.startsWith('@@')) { out.push({ t: 'fn', name: '@' }); i += 2; continue; }
    if (ch === '@') {
      const m = /^@([A-Za-z][A-Za-z0-9.]*)/.exec(rest);
      if (!m) throw new LotusError('Function name expected after @');
      out.push({ t: 'fn', name: m[1].toUpperCase() });
      i += m[0].length;
      continue;
    }
    const lg = /^#(AND|OR|NOT)#/i.exec(rest);
    if (lg) { out.push({ t: 'op', v: '#' + lg[1].toUpperCase() + '#' }); i += lg[0].length; continue; }
    const r = RE_REF.exec(rest);
    if (r) {
      const a = refText(r[1], o);
      let text: string;
      if (r[2]) {
        const b = refText(r[2], o);
        const sheetA = a.sheet, sheetB = b.sheet ?? a.sheet;
        if (sheetA && sheetB && sheetA !== sheetB) text = `${quoteSheet(sheetA)}:${quoteSheet(sheetB)}!${a.cell}:${b.cell}`;
        else text = `${sheetA ? quoteSheet(sheetA) + '!' : ''}${a.cell}:${b.cell}`;
      } else text = `${a.sheet ? quoteSheet(a.sheet) + '!' : ''}${a.cell}`;
      out.push({ t: 'ref', text, simple: true });
      i += r[0].length;
      continue;
    }
    const n = /^(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?/.exec(rest);
    if (n) { out.push({ t: 'num', v: n[0] }); i += n[0].length; continue; }
    const id = /^\$?([A-Za-z_][A-Za-z0-9_]*)/.exec(rest);
    if (id) { out.push({ t: 'name', v: id[1] }); i += id[0].length; continue; }
    const op = /^(<>|<=|>=|[-+*/^&=<>])/.exec(rest);
    if (op) { out.push({ t: 'op', v: op[1] }); i += op[1].length; continue; }
    if (ch === '(') { out.push({ t: 'lp' }); i++; continue; }
    if (ch === ')') { out.push({ t: 'rp' }); i++; continue; }
    if (ch === ',' || ch === ';') { out.push({ t: 'sep' }); i++; continue; }
    throw new LotusError(`Unexpected character "${ch}"`);
  }
  return out;
}

// 1-2-3 precedence (higher binds tighter).
const BIN_PREC: Record<string, number> = { '^': 7, '*': 5, '/': 5, '+': 4, '-': 4, '=': 3, '<>': 3, '<': 3, '>': 3, '<=': 3, '>=': 3, '#AND#': 1, '#OR#': 1, '&': 1 };
/** Functions that take no parentheses in 1-2-3. */
const NO_ARGS = new Set(['PI', 'TRUE', 'FALSE', 'NA', 'ERR', 'NOW', 'TODAY', 'RAND']);

function parseTokens(toks: Tok[]): LNode {
  let p = 0;
  const peek = () => toks[p];
  const expr = (minPrec: number): LNode => {
    let left = prefix();
    for (;;) {
      const t = peek();
      if (!t || t.t !== 'op' || !(t.v in BIN_PREC)) break;
      const prec = BIN_PREC[t.v];
      if (prec < minPrec) break;
      p++;
      const right = expr(prec + 1);
      left = { k: 'bin', op: t.v, a: left, b: right };
    }
    return left;
  };
  const prefix = (): LNode => {
    const t = toks[p++];
    if (!t) throw new LotusError('Formula ends too soon');
    switch (t.t) {
      case 'num': return { k: 'num', v: t.v };
      case 'str': return { k: 'str', v: t.v };
      case 'ref': return { k: 'ref', text: t.text };
      case 'name': return { k: 'name', v: t.v };
      case 'op':
        if (t.v === '-' || t.v === '+') return { k: 'un', op: t.v, a: expr(6) };
        if (t.v === '#NOT#') return { k: 'un', op: '#NOT#', a: expr(2) };
        throw new LotusError(`Unexpected operator ${t.v}`);
      case 'lp': {
        const e = expr(0);
        if (toks[p++]?.t !== 'rp') throw new LotusError('Missing )');
        return e;
      }
      case 'fn': {
        const args: LNode[] = [];
        if (peek()?.t === 'lp') {
          p++;
          if (peek()?.t === 'rp') p++;
          else {
            for (;;) {
              args.push(expr(0));
              const n = toks[p++];
              if (n?.t === 'rp') break;
              if (n?.t !== 'sep') throw new LotusError('Missing ) after function arguments');
            }
          }
        } else if (!NO_ARGS.has(t.name)) throw new LotusError(`@${t.name} needs arguments`);
        return { k: 'fn', name: t.name, args };
      }
      default: throw new LotusError('Unexpected symbol');
    }
  };
  const e = expr(0);
  if (p < toks.length) throw new LotusError('Unexpected text after the formula');
  return e;
}

// ---------------------------------------------------------------------------------------------
// Emission

// Target precedence (higher binds tighter).
const OUT_PREC: Record<string, number> = { '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '&': 2, '+': 3, '-': 3, '*': 4, '/': 4, '^': 5 };
const ATOM = 99;

interface Out { s: string; p: number }

const wrap = (o: Out, parent: number, right = false) => (o.p < parent || (right && o.p === parent) ? `(${o.s})` : o.s);
const call = (name: string, args: string[]): Out => ({ s: `${name}(${args.join(',')})`, p: ATOM });
const isNum = (n: LNode | undefined): n is { k: 'num'; v: string } => n?.k === 'num';
/** 1-2-3 offsets count from 0; the target counts from 1. */
const plus1 = (n: LNode, s: string) => (isNum(n) ? String(+n.v + 1) : n.k === 'bin' || n.k === 'un' ? `(${s})+1` : `${s}+1`);
/** An argument that is itself an expression, parenthesised so it can take part in a larger one. */
const par = (n: LNode, s: string) => (n.k === 'bin' || n.k === 'un' ? `(${s})` : s);
const neg = (n: LNode, s: string) => (isNum(n) ? '-' + n.v : `-(${s})`);

/** Same name and meaning in both languages. */
const SAME = new Set(('ABS ACOS ASIN ATAN ATAN2 COS SIN TAN EXP LN LOG SQRT ROUND ROUNDUP ROUNDDOWN TRUNC PI RAND SUM SUMPRODUCT SUMSQ SUMXMY2 '
  + 'IF NOW TODAY DATE TIME DATEVALUE TIMEVALUE DAY MONTH HOUR MINUTE SECOND UPPER LOWER PROPER TRIM EXACT LEFT RIGHT CHAR CODE VALUE N '
  + 'ISNUMBER ISNA NPV SLN SYD DDB VDB ROWS MEDIAN RANK LARGE SMALL PRODUCT EVEN ODD SIGN FACT COMBIN PERMUT DEGREES RADIANS SUMIF COUNTIF '
  + 'SHEETS DATEDIF WORKDAY NETWORKDAYS QUOTIENT GEOMEAN HARMEAN KURT SKEW CLEAN ISBETWEEN WEIGHTAVG SEMEAN REFCONVERT XINDEX').split(' '));

/** Simple renames. */
const RENAME: Record<string, string> = {
  // Labels in ranges count as zero in the plain statistical functions; the PURE versions ignore them.
  AVG: 'AVERAGEA', COUNT: 'COUNTA', MAX: 'MAXA', MIN: 'MINA', STD: 'STDEVPA', STDS: 'STDEVA', VAR: 'VARPA', VARS: 'VARA',
  PUREAVG: 'AVERAGE', PURECOUNT: 'COUNT', PUREMAX: 'MAX', PUREMIN: 'MIN', PURESTD: 'STDEV.P', PURESTDS: 'STDEV.S', PUREVAR: 'VAR.P', PUREVARS: 'VAR.S',
  LENGTH: 'LEN', REPEAT: 'REPT', S: 'T', ISSTRING: 'ISTEXT', ISRANGE: 'ISREF', ISEMPTY: 'ISBLANK', ISERR: 'ISERROR', COLS: 'COLUMNS',
  D360: 'DAYS360', DAYS360: 'DAYS360', '@': 'INDIRECT', CELL: 'CELL', CELLPOINTER: 'CELL', ROUNDM: 'MROUND',
};

const DB: Record<string, string> = { DSUM: 'DSUM', DAVG: 'DAVERAGE', DCOUNT: 'DCOUNTA', DMAX: 'DMAX', DMIN: 'DMIN', DSTD: 'DSTDEVP', DSTDS: 'DSTDEV', DVAR: 'DVARP', DVARS: 'DVAR', DGET: 'DGET' };

/** Every @function the translator knows (for help text and for telling 1-2-3 names apart). */
export const LOTUS_FUNCTIONS: string[] = [
  ...SAME, ...Object.keys(RENAME).filter((k) => k !== '@'), ...Object.keys(DB), '@@',
  'TRUE', 'FALSE', 'NA', 'ERR', 'CHOOSE', 'INDEX', 'VLOOKUP', 'HLOOKUP', 'MID', 'FIND', 'REPLACE', 'STRING', 'YEAR', 'WEEKDAY',
  'PMT', 'PV', 'FV', 'TERM', 'CTERM', 'RATE', 'IRR', 'INT', 'MOD',
].sort();

function fnOut(name: string, nodes: LNode[], a: string[]): Out {
  const n = nodes.length;
  const need = (lo: number, hi = lo) => { if (n < lo || n > hi) throw new LotusError(`@${name} takes ${lo === hi ? lo : `${lo} to ${hi}`} argument${hi === 1 ? '' : 's'}`); };
  if (name in DB) {
    need(3);
    const f = nodes[1];
    const field = isNum(f) ? String(+f.v + 1) : f.k === 'str' ? a[1] : `IF(ISNUMBER(${a[1]}),${a[1]}+1,${a[1]})`;
    return call(DB[name], [a[0], field, a[2]]);
  }
  switch (name) {
    case 'TRUE': case 'FALSE': need(0); return { s: name, p: ATOM };
    case 'NA': need(0); return call('NA', []);
    case 'ERR': need(0); return { s: '#VALUE!', p: ATOM };
    case 'CHOOSE': if (n < 2) need(2, 255); return call('CHOOSE', [plus1(nodes[0], a[0]), ...a.slice(1)]);
    case 'INDEX': need(3, 4); return call('INDEX', [a[0], plus1(nodes[2], a[2]), plus1(nodes[1], a[1])]);
    case 'VLOOKUP': case 'HLOOKUP':
      // Numbers match the largest value not above them; text must match exactly.
      need(3); return call(name, [a[0], a[1], plus1(nodes[2], a[2]), `ISNUMBER(${a[0]})`]);
    case 'MID': need(3); return call('MID', [a[0], plus1(nodes[1], a[1]), a[2]]);
    // 1-2-3 truncates toward zero, and its remainder takes the sign of the dividend (here INT rounds down and MOD
    // takes the sign of the divisor).
    case 'INT': need(1); return call('TRUNC', [a[0]]);
    case 'MOD': { need(2); const x = par(nodes[0], a[0]), y = par(nodes[1], a[1]); return { s: `${x}-${y}*TRUNC(${x}/${y})`, p: OUT_PREC['-'] }; }
    case 'FIND': need(3); return { s: `FIND(${a[0]},${a[1]},${plus1(nodes[2], a[2])})-1`, p: OUT_PREC['-'] };
    case 'REPLACE': need(4); return call('REPLACE', [a[0], plus1(nodes[1], a[1]), a[2], a[3]]);
    case 'STRING': need(2); return call('FIXED', [a[0], a[1], 'TRUE']);
    case 'YEAR': need(1); return { s: `YEAR(${a[0]})-1900`, p: OUT_PREC['-'] };
    case 'WEEKDAY': need(1); return call('WEEKDAY', [a[0], '3']); // 0 = Monday ... 6 = Sunday
    case 'PMT': need(3); return call('PMT', [a[1], a[2], neg(nodes[0], a[0])]);
    case 'PV': need(3); return call('PV', [a[1], a[2], neg(nodes[0], a[0])]);
    case 'FV': need(3); return call('FV', [a[1], a[2], neg(nodes[0], a[0])]);
    case 'TERM': need(3); return call('NPER', [a[1], neg(nodes[0], a[0]), '0', a[2]]);
    case 'CTERM': need(3); return call('NPER', [a[0], '0', neg(nodes[2], a[2]), a[1]]);
    case 'RATE': need(3); return call('RRI', [a[2], a[1], a[0]]);
    case 'IRR': need(2); return call('IRR', [a[1], a[0]]);
  }
  if (name in RENAME) return call(RENAME[name], a);
  return call(name, a);
}

function emit(n: LNode): Out {
  switch (n.k) {
    case 'num': return { s: n.v, p: ATOM };
    case 'str': return { s: `"${n.v.replace(/"/g, '""')}"`, p: ATOM };
    case 'ref': return { s: n.text, p: ATOM };
    case 'name': return { s: n.v, p: ATOM };
    case 'fn': {
      if (NO_ARGS.has(n.name) && !n.args.length && !['TRUE', 'FALSE', 'ERR'].includes(n.name)) return call(RENAME[n.name] || n.name, []);
      return fnOut(n.name, n.args, n.args.map((x) => emit(x).s));
    }
    case 'un': {
      if (n.op === '#NOT#') return call('NOT', [emit(n.a).s]);
      if (n.op === '+') return emit(n.a);
      return { s: '-' + wrap(emit(n.a), 6), p: 6 };
    }
    case 'bin': {
      if (n.op === '#AND#' || n.op === '#OR#') {
        // Flatten chains of the same operator into one call.
        const items: LNode[] = [];
        const collect = (x: LNode) => { if (x.k === 'bin' && x.op === n.op) { collect(x.a); collect(x.b); } else items.push(x); };
        collect(n);
        return call(n.op === '#AND#' ? 'AND' : 'OR', items.map((x) => emit(x).s));
      }
      const p = OUT_PREC[n.op];
      return { s: wrap(emit(n.a), p) + n.op + wrap(emit(n.b), p, true), p };
    }
  }
}

/** Translates a 1-2-3 formula (with or without a leading +, - or @) to a formula without the leading "=". Throws LotusError. */
export function translateLotus(src: string, o: LotusOptions = {}): string {
  let t = src.trim();
  if (t.startsWith('=')) t = t.slice(1);
  if (t.startsWith('+')) t = t.slice(1);
  const toks = tokenize(t, o);
  if (!toks.length) throw new LotusError('Empty formula');
  return emit(parseTokens(toks)).s;
}

/** True when the text uses 1-2-3 syntax: @functions, A1..B5 ranges, or #AND#/#OR#/#NOT#. */
export function hasLotusSyntax(text: string): boolean {
  return /@@?\(|@[A-Za-z]/.test(text) || /\$?[A-Za-z]{1,3}\$?\d+\.\.?\$?[A-Za-z]{1,3}\$?\d+/.test(text) || /#(AND|OR|NOT)#/i.test(text);
}

/**
 * What a typed entry means under 1-2-3 formula entry, or null when it is not a 1-2-3 formula:
 * - "@SUM(A1..A5)", "+A1*@AVG(B1..B3)", "-@PI": 1-2-3 formulas, translated with 1-2-3 meaning.
 * - "=..." entries are translated only when they use 1-2-3-only syntax (".." ranges, #AND#) or
 *   @functions the app does not know; "=@INDEX(...)" keeps its implicit-intersection meaning.
 */
export function lotusEntry(text: string, isKnownFunction: (name: string) => boolean, o: LotusOptions = {}): string | null {
  const t = text.trim();
  if (t.length < 2) return null;
  const first = t[0];
  const known = (nm: string) => LOTUS_FUNCTIONS.includes(nm) || isKnownFunction(nm);
  const atNames = [...t.matchAll(/@([A-Za-z][A-Za-z0-9.]*)/g)].map((m) => m[1].toUpperCase());
  if (first === '@') {
    // "@john (work)" stays text: every @name must be a function.
    if (!/^@(@\(|[A-Za-z])/.test(t) || !atNames.every(known)) return null;
  } else if (first === '+' || first === '-') {
    if (!hasLotusSyntax(t) || !atNames.every(known)) return null;
  } else if (first === '=') {
    const rangeOrLogic = /\$?[A-Za-z]{1,3}\$?\d+\.\.?\$?[A-Za-z]{1,3}\$?\d+/.test(t) || /#(AND|OR|NOT)#/i.test(t);
    const names = [...t.matchAll(/@([A-Za-z][A-Za-z0-9.]*)\s*\(/g)].map((m) => m[1].toUpperCase());
    const lotusOnly = names.length > 0 && names.every((nm) => !isKnownFunction(nm) && LOTUS_FUNCTIONS.includes(nm));
    if (!rangeOrLogic && !lotusOnly && !/@@\(/.test(t)) return null;
  } else return null;
  try {
    return translateLotus(t, o);
  } catch {
    return null;
  }
}

/** 1-2-3 label prefixes: ' left, " right, ^ centered, \ repeating. Returns the label and alignment, or null. */
export function lotusLabel(text: string): { text: string; align: 'left' | 'right' | 'center' | 'fill' } | null {
  if (text.length < 2) return null;
  const a = ({ "'": 'left', '"': 'right', '^': 'center', '\\': 'fill' } as const)[text[0] as "'" | '"' | '^' | '\\'];
  if (!a) return null;
  // A label wrapped in quotes ("text") is ordinary text, not a prefix.
  if (text[0] === '"' && text.endsWith('"')) return null;
  return { text: text.slice(1), align: a };
}
