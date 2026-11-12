/**
 * Analysis ToolPak: statistical and engineering analyses whose results are written as labeled values to a new
 * sheet or an output range. Every tool is a pure function from numbers to an output matrix (tested in
 * tests/unit/data.test.ts); the command only gathers input and writes the result.
 */
import { def } from './registry';
import type { App } from '../app';
import type { Sheet } from '../model/workbook';
import { newId } from '../model/workbook';
import type { Scalar, Matrix, ChartDef } from '../model/types';
import { CellError } from '../model/types';
import { resolveRangeText, rangeRefText } from '../model/structure';
import { activeRange, clipToUsed, currentRegion } from './ops';
import { openDialog, toast, alertDialog, confirmDialog, type Field } from '../ui/dialog';
import { tSf, tInv, fCdf, fSf, normSf, normInv, invert } from '../calc/functions/stats';
import { blocked, addSheetAfter, writeMatrix } from './data';

const NA = () => new CellError('#N/A');
const DIV0 = () => new CellError('#DIV/0!');
const tSf2 = (t: number, df: number) => tSf(t, df);

export interface Series { label: string; data: number[] }
export interface ToolOutput { m: Scalar[][]; bold?: number[]; chart?: { cats: [number, number, number]; vals: [number, number, number][]; type: ChartDef['type']; title: string } }

// --------------------------------------------------------------------------------------------
// Basic statistics

export const mean = (x: number[]) => x.reduce((a, b) => a + b, 0) / x.length;
export function variance(x: number[], sample = true): number { const m = mean(x); return x.reduce((a, b) => a + (b - m) ** 2, 0) / (x.length - (sample ? 1 : 0)); }
export function median(x: number[]): number { const s = [...x].sort((a, b) => a - b); const n = s.length; return n % 2 ? s[(n - 1) / 2] : (s[n / 2 - 1] + s[n / 2]) / 2; }
export function mode(x: number[]): number | null {
  const counts = new Map<number, number>(); let best: number | null = null, bc = 1;
  for (const v of x) { const c = (counts.get(v) || 0) + 1; counts.set(v, c); if (c > bc) { bc = c; best = v; } }
  return best;
}
export function skewness(x: number[]): number { const n = x.length, m = mean(x), s = Math.sqrt(variance(x)); return (n / ((n - 1) * (n - 2))) * x.reduce((a, b) => a + ((b - m) / s) ** 3, 0); }
export function kurtosis(x: number[]): number {
  const n = x.length, m = mean(x), s = Math.sqrt(variance(x));
  return ((n * (n + 1)) / ((n - 1) * (n - 2) * (n - 3))) * x.reduce((a, b) => a + ((b - m) / s) ** 4, 0) - (3 * (n - 1) ** 2) / ((n - 2) * (n - 3));
}
const fin = (v: number): Scalar => (Number.isFinite(v) ? +v.toPrecision(15) : DIV0());
const fInvUpper = (alpha: number, d1: number, d2: number) => invert((x) => fCdf(x, d1, d2), 1 - alpha, 0, 10);
const fInvLower = (alpha: number, d1: number, d2: number) => invert((x) => fCdf(x, d1, d2), alpha, 0, 1);

// --------------------------------------------------------------------------------------------
// Tools

export function descriptiveStatistics(series: Series[], o: { summary?: boolean; conf?: number | null; kLarge?: number | null; kSmall?: number | null } = { summary: true }): ToolOutput {
  const rows: Scalar[][] = [];
  const width = series.length * 2;
  const put = (i: number, j: number, v: Scalar) => { while (rows.length <= i) rows.push(new Array(width).fill(null)); rows[i][j] = v; };
  series.forEach((s, k) => {
    const x = s.data, n = x.length, j = k * 2;
    const items: [string, Scalar][] = [];
    if (o.summary !== false) {
      const md = mode(x);
      const sd = Math.sqrt(variance(x));
      items.push(['Mean', fin(mean(x))], ['Standard Error', fin(sd / Math.sqrt(n))], ['Median', fin(median(x))], ['Mode', md == null ? NA() : md],
        ['Standard Deviation', fin(sd)], ['Sample Variance', fin(variance(x))], ['Kurtosis', n > 3 ? fin(kurtosis(x)) : DIV0()], ['Skewness', n > 2 ? fin(skewness(x)) : DIV0()],
        ['Range', fin(Math.max(...x) - Math.min(...x))], ['Minimum', Math.min(...x)], ['Maximum', Math.max(...x)], ['Sum', fin(x.reduce((a, b) => a + b, 0))], ['Count', n]);
    }
    const sorted = [...x].sort((a, b) => b - a);
    if (o.kLarge) items.push([`Largest(${o.kLarge})`, o.kLarge <= n ? sorted[o.kLarge - 1] : new CellError('#NUM!')]);
    if (o.kSmall) items.push([`Smallest(${o.kSmall})`, o.kSmall <= n ? sorted[n - o.kSmall] : new CellError('#NUM!')]);
    if (o.conf) items.push([`Confidence Level(${(o.conf * 100).toFixed(1)}%)`, n > 1 ? fin(tInv(1 - (1 - o.conf) / 2, n - 1) * Math.sqrt(variance(x)) / Math.sqrt(n)) : DIV0()]);
    put(0, j, s.label);
    items.forEach(([l, v], i) => { put(i + 2, j, l); put(i + 2, j + 1, v); });
  });
  return { m: rows, bold: [0] };
}

export function histogram(data: number[], bins: number[] | null, o: { pareto?: boolean; cumulative?: boolean; chart?: boolean } = {}): ToolOutput {
  let edges = bins && bins.length ? [...bins].sort((a, b) => a - b) : null;
  if (!edges) {
    const mn = Math.min(...data), mx = Math.max(...data);
    const k = Math.max(1, Math.round(Math.sqrt(data.length)));
    const w = (mx - mn) / k || 1;
    edges = []; for (let i = 0; i <= k; i++) edges.push(+(mn + w * i).toPrecision(12));
  }
  const freq = new Array(edges.length + 1).fill(0);
  for (const v of data) { let i = edges.findIndex((e) => v <= e); if (i < 0) i = edges.length; freq[i]++; }
  const labels: Scalar[] = [...edges, 'More'];
  const total = data.length;
  const m: Scalar[][] = [['Bin', 'Frequency', ...(o.cumulative ? ['Cumulative %'] : [])]];
  let run = 0;
  labels.forEach((l, i) => { run += freq[i]; m.push([l, freq[i], ...(o.cumulative ? [total ? run / total : 0] : [])]); });
  if (o.pareto) {
    const order = labels.map((_, i) => i).sort((a, b) => freq[b] - freq[a] || a - b);
    const off = m[0].length + 1;
    m[0].push(null, 'Bin', 'Frequency', ...(o.cumulative ? ['Cumulative %'] : []));
    let r2 = 0;
    order.forEach((i, k) => { r2 += freq[i]; const row = m[k + 1]; while (row.length < off) row.push(null); row.push(labels[i], freq[i], ...(o.cumulative ? [total ? r2 / total : 0] : [])); });
  }
  return { m, bold: [0], chart: o.chart ? { cats: [1, labels.length, 0], vals: [[1, labels.length, 1]], type: 'column', title: 'Histogram' } : undefined };
}

export function correlationMatrix(series: Series[], cov = false): ToolOutput {
  const k = series.length;
  const m: Scalar[][] = [[null, ...series.map((s) => s.label)]];
  for (let i = 0; i < k; i++) {
    const row: Scalar[] = [series[i].label];
    for (let j = 0; j < k; j++) {
      if (j > i) { row.push(null); continue; }
      const x = series[i].data, y = series[j].data, n = Math.min(x.length, y.length);
      const mx = mean(x.slice(0, n)), my = mean(y.slice(0, n));
      let sxy = 0, sxx = 0, syy = 0;
      for (let t = 0; t < n; t++) { sxy += (x[t] - mx) * (y[t] - my); sxx += (x[t] - mx) ** 2; syy += (y[t] - my) ** 2; }
      row.push(fin(cov ? sxy / n : i === j ? 1 : sxy / Math.sqrt(sxx * syy)));
    }
    m.push(row);
  }
  return { m, bold: [0] };
}

/** Inverts a symmetric positive (semi)definite matrix by Gauss–Jordan elimination; null when singular. */
export function invertMatrix(a: number[][]): number[][] | null {
  const n = a.length;
  const M = a.map((r, i) => [...r, ...r.map((_, j) => (i === j ? 1 : 0))]);
  for (let c = 0; c < n; c++) {
    let p = c;
    for (let r = c + 1; r < n; r++) if (Math.abs(M[r][c]) > Math.abs(M[p][c])) p = r;
    if (Math.abs(M[p][c]) < 1e-12 * (1 + Math.abs(M[c][c]))) return null;
    [M[c], M[p]] = [M[p], M[c]];
    const d = M[c][c];
    for (let j = 0; j < 2 * n; j++) M[c][j] /= d;
    for (let r = 0; r < n; r++) if (r !== c) { const f = M[r][c]; if (f) for (let j = 0; j < 2 * n; j++) M[r][j] -= f * M[c][j]; }
  }
  return M.map((r) => r.slice(n));
}

export interface RegressionFit { coef: number[]; se: number[]; r2: number; adjR2: number; seY: number; n: number; dfReg: number; dfRes: number; ssReg: number; ssRes: number; F: number; sigF: number; predicted: number[]; residuals: number[] }

export function regressionFit(y: number[], X: number[][], constZero = false): RegressionFit {
  const n = y.length, k = X[0]?.length || 0;
  const D = X.map((r) => (constZero ? [...r] : [1, ...r]));
  const p = D[0].length;
  const XtX = Array.from({ length: p }, (_, i) => Array.from({ length: p }, (_, j) => D.reduce((s, r) => s + r[i] * r[j], 0)));
  const Xty = Array.from({ length: p }, (_, i) => D.reduce((s, r, t) => s + r[i] * y[t], 0));
  const inv = invertMatrix(XtX);
  if (!inv) throw new Error('The X range values are collinear (the regression cannot be computed).');
  const b = inv.map((row) => row.reduce((s, v, j) => s + v * Xty[j], 0));
  const predicted = D.map((r) => r.reduce((s, v, j) => s + v * b[j], 0));
  const residuals = y.map((v, t) => v - predicted[t]);
  const ssRes = residuals.reduce((s, e) => s + e * e, 0);
  const my = mean(y);
  const ssTot = constZero ? y.reduce((s, v) => s + v * v, 0) : y.reduce((s, v) => s + (v - my) ** 2, 0);
  const ssReg = ssTot - ssRes;
  const dfReg = k, dfRes = n - k - (constZero ? 0 : 1);
  const mse = ssRes / dfRes;
  const r2 = ssTot ? ssReg / ssTot : 1;
  const F = (ssReg / dfReg) / mse;
  const coef = constZero ? [0, ...b] : b;
  const se = constZero ? [NaN, ...inv.map((r, i) => Math.sqrt(mse * r[i]))] : inv.map((r, i) => Math.sqrt(mse * r[i]));
  return { coef, se, r2, adjR2: 1 - (1 - r2) * ((n - (constZero ? 0 : 1)) / dfRes), seY: Math.sqrt(mse), n, dfReg, dfRes, ssReg, ssRes, F, sigF: fSf(F, dfReg, dfRes), predicted, residuals };
}

export function regression(y: Series, xs: Series[], o: { constZero?: boolean; conf?: number; residuals?: boolean; stdResiduals?: boolean } = {}): ToolOutput {
  const n = y.data.length;
  const X = y.data.map((_, t) => xs.map((s) => s.data[t]));
  const f = regressionFit(y.data, X, !!o.constZero);
  const conf = o.conf ?? 0.95;
  const tc95 = tInv(0.975, f.dfRes), tc = tInv(1 - (1 - conf) / 2, f.dfRes);
  const m: Scalar[][] = [
    ['SUMMARY OUTPUT'], [],
    ['Regression Statistics'],
    ['Multiple R', fin(Math.sqrt(Math.max(0, f.r2)))], ['R Square', fin(f.r2)], ['Adjusted R Square', fin(f.adjR2)], ['Standard Error', fin(f.seY)], ['Observations', n], [],
    ['ANOVA'],
    [null, 'df', 'SS', 'MS', 'F', 'Significance F'],
    ['Regression', f.dfReg, fin(f.ssReg), fin(f.ssReg / f.dfReg), fin(f.F), fin(f.sigF)],
    ['Residual', f.dfRes, fin(f.ssRes), fin(f.ssRes / f.dfRes)],
    ['Total', f.dfReg + f.dfRes, fin(f.ssReg + f.ssRes)], [],
    [null, 'Coefficients', 'Standard Error', 't Stat', 'P-value', 'Lower 95%', 'Upper 95%', `Lower ${(conf * 100).toFixed(1)}%`, `Upper ${(conf * 100).toFixed(1)}%`],
  ];
  const names = ['Intercept', ...xs.map((s) => s.label)];
  f.coef.forEach((b, i) => {
    const se = f.se[i];
    if (!Number.isFinite(se)) { m.push([names[i], 0, NA(), NA(), NA(), NA(), NA(), NA(), NA()]); return; }
    const t = b / se;
    m.push([names[i], fin(b), fin(se), fin(t), fin(2 * tSf2(Math.abs(t), f.dfRes)), fin(b - tc95 * se), fin(b + tc95 * se), fin(b - tc * se), fin(b + tc * se)]);
  });
  const bold = [0, 2, 9, 10, 15];
  if (o.residuals || o.stdResiduals) {
    m.push([], [], ['RESIDUAL OUTPUT'], []);
    bold.push(m.length - 2);
    const sdE = Math.sqrt(f.ssRes / (n - 1));
    m.push(['Observation', `Predicted ${y.label}`, 'Residuals', ...(o.stdResiduals ? ['Standard Residuals'] : [])]);
    bold.push(m.length - 1);
    f.predicted.forEach((p, t) => m.push([t + 1, fin(p), fin(f.residuals[t]), ...(o.stdResiduals ? [fin(f.residuals[t] / sdE)] : [])]));
  }
  return { m, bold };
}

export type TTestKind = 'paired' | 'equal' | 'unequal';

export function tTest(a: Series, b: Series, kind: TTestKind, hmd = 0, alpha = 0.05): ToolOutput {
  const n1 = a.data.length, n2 = b.data.length;
  const m1 = mean(a.data), m2 = mean(b.data), v1 = variance(a.data), v2 = variance(b.data);
  const rows: Scalar[][] = [];
  const title = kind === 'paired' ? 't-Test: Paired Two Sample for Means' : kind === 'equal' ? 't-Test: Two-Sample Assuming Equal Variances' : 't-Test: Two-Sample Assuming Unequal Variances';
  rows.push([title], [], [null, a.label, b.label], ['Mean', fin(m1), fin(m2)], ['Variance', fin(v1), fin(v2)], ['Observations', n1, n2]);
  let t: number, df: number;
  if (kind === 'paired') {
    if (n1 !== n2) throw new Error('Paired samples must have the same number of observations.');
    const d = a.data.map((x, i) => x - b.data[i]);
    const md = mean(d), sd = Math.sqrt(variance(d));
    const mx = m1, my = m2;
    let sxy = 0, sxx = 0, syy = 0;
    for (let i = 0; i < n1; i++) { sxy += (a.data[i] - mx) * (b.data[i] - my); sxx += (a.data[i] - mx) ** 2; syy += (b.data[i] - my) ** 2; }
    rows.push(['Pearson Correlation', fin(sxy / Math.sqrt(sxx * syy))]);
    rows.push(['Hypothesized Mean Difference', hmd]);
    df = n1 - 1;
    t = (md - hmd) / (sd / Math.sqrt(n1));
  } else if (kind === 'equal') {
    const pv = ((n1 - 1) * v1 + (n2 - 1) * v2) / (n1 + n2 - 2);
    rows.push(['Pooled Variance', fin(pv)], ['Hypothesized Mean Difference', hmd]);
    df = n1 + n2 - 2;
    t = (m1 - m2 - hmd) / Math.sqrt(pv * (1 / n1 + 1 / n2));
  } else {
    rows.push(['Hypothesized Mean Difference', hmd]);
    const q1 = v1 / n1, q2 = v2 / n2;
    df = Math.round((q1 + q2) ** 2 / (q1 * q1 / (n1 - 1) + q2 * q2 / (n2 - 1)));
    t = (m1 - m2 - hmd) / Math.sqrt(q1 + q2);
  }
  rows.push(['df', df], ['t Stat', fin(t)], ['P(T<=t) one-tail', fin(tSf2(Math.abs(t), df))], ['t Critical one-tail', fin(tInv(1 - alpha, df))], ['P(T<=t) two-tail', fin(2 * tSf2(Math.abs(t), df))], ['t Critical two-tail', fin(tInv(1 - alpha / 2, df))]);
  return { m: rows, bold: [0, 2] };
}

export function zTest(a: Series, b: Series, var1: number, var2: number, hmd = 0, alpha = 0.05): ToolOutput {
  const n1 = a.data.length, n2 = b.data.length, m1 = mean(a.data), m2 = mean(b.data);
  const z = (m1 - m2 - hmd) / Math.sqrt(var1 / n1 + var2 / n2);
  return {
    m: [['z-Test: Two Sample for Means'], [], [null, a.label, b.label], ['Mean', fin(m1), fin(m2)], ['Known Variance', var1, var2], ['Observations', n1, n2], ['Hypothesized Mean Difference', hmd],
      ['z', fin(z)], ['P(Z<=z) one-tail', fin(normSf(Math.abs(z)))], ['z Critical one-tail', fin(normInv(1 - alpha))], ['P(Z<=z) two-tail', fin(2 * normSf(Math.abs(z)))], ['z Critical two-tail', fin(normInv(1 - alpha / 2))]],
    bold: [0, 2],
  };
}

export function fTest(a: Series, b: Series, alpha = 0.05): ToolOutput {
  const n1 = a.data.length, n2 = b.data.length, v1 = variance(a.data), v2 = variance(b.data);
  const F = v1 / v2, d1 = n1 - 1, d2 = n2 - 1;
  const upper = F >= 1;
  return {
    m: [['F-Test Two-Sample for Variances'], [], [null, a.label, b.label], ['Mean', fin(mean(a.data)), fin(mean(b.data))], ['Variance', fin(v1), fin(v2)], ['Observations', n1, n2], ['df', d1, d2],
      ['F', fin(F)], ['P(F<=f) one-tail', fin(upper ? fSf(F, d1, d2) : fCdf(F, d1, d2))], ['F Critical one-tail', fin(upper ? fInvUpper(alpha, d1, d2) : fInvLower(alpha, d1, d2))]],
    bold: [0, 2],
  };
}

const summaryRow = (label: Scalar, x: number[]): Scalar[] => [label, x.length, fin(x.reduce((s, v) => s + v, 0)), fin(mean(x)), x.length > 1 ? fin(variance(x)) : DIV0()];

export function anovaSingle(groups: Series[], alpha = 0.05): ToolOutput {
  const all = groups.flatMap((g) => g.data);
  const G = mean(all), N = all.length, k = groups.length;
  const ssb = groups.reduce((s, g) => s + g.data.length * (mean(g.data) - G) ** 2, 0);
  const ssw = groups.reduce((s, g) => { const m = mean(g.data); return s + g.data.reduce((a, v) => a + (v - m) ** 2, 0); }, 0);
  const dfb = k - 1, dfw = N - k;
  const F = (ssb / dfb) / (ssw / dfw);
  const m: Scalar[][] = [['Anova: Single Factor'], [], ['SUMMARY'], ['Groups', 'Count', 'Sum', 'Average', 'Variance'], ...groups.map((g) => summaryRow(g.label, g.data)), [], [],
    ['ANOVA'], ['Source of Variation', 'SS', 'df', 'MS', 'F', 'P-value', 'F crit'],
    ['Between Groups', fin(ssb), dfb, fin(ssb / dfb), fin(F), fin(fSf(F, dfb, dfw)), fin(fInvUpper(alpha, dfb, dfw))],
    ['Within Groups', fin(ssw), dfw, fin(ssw / dfw)], [], ['Total', fin(ssb + ssw), N - 1]];
  return { m, bold: [0, 3, 6 + groups.length, 7 + groups.length] };
}

/** Two-factor ANOVA without replication: data[i][j] = row i, column j. */
export function anovaTwoNoRep(data: number[][], rowLabels: string[], colLabels: string[], alpha = 0.05): ToolOutput {
  const r = data.length, c = data[0].length;
  const all = data.flat(), G = mean(all);
  const rowM = data.map((row) => mean(row)), colM = colLabels.map((_, j) => mean(data.map((row) => row[j])));
  const ssr = c * rowM.reduce((s, m) => s + (m - G) ** 2, 0), ssc = r * colM.reduce((s, m) => s + (m - G) ** 2, 0);
  const sst = all.reduce((s, v) => s + (v - G) ** 2, 0), sse = sst - ssr - ssc;
  const dfr = r - 1, dfc = c - 1, dfe = dfr * dfc;
  const mse = sse / dfe;
  const Fr = (ssr / dfr) / mse, Fc = (ssc / dfc) / mse;
  const m: Scalar[][] = [['Anova: Two-Factor Without Replication'], [], ['SUMMARY', 'Count', 'Sum', 'Average', 'Variance'],
    ...data.map((row, i) => summaryRow(rowLabels[i], row)), [], ...colLabels.map((l, j) => summaryRow(l, data.map((row) => row[j]))), [], [],
    ['ANOVA'], ['Source of Variation', 'SS', 'df', 'MS', 'F', 'P-value', 'F crit'],
    ['Rows', fin(ssr), dfr, fin(ssr / dfr), fin(Fr), fin(fSf(Fr, dfr, dfe)), fin(fInvUpper(alpha, dfr, dfe))],
    ['Columns', fin(ssc), dfc, fin(ssc / dfc), fin(Fc), fin(fSf(Fc, dfc, dfe)), fin(fInvUpper(alpha, dfc, dfe))],
    ['Error', fin(sse), dfe, fin(mse)], [], ['Total', fin(sst), r * c - 1]];
  return { m, bold: [0, 2] };
}

/** Two-factor ANOVA with replication: blocks of `reps` rows per sample. */
export function anovaTwoRep(data: number[][], reps: number, sampleLabels: string[], colLabels: string[], alpha = 0.05): ToolOutput {
  const a = Math.floor(data.length / reps), b = data[0].length;
  if (a < 1 || a * reps !== data.length) throw new Error('Input range must contain the same number of rows for every sample.');
  const cell = (i: number, j: number) => data.slice(i * reps, (i + 1) * reps).map((row) => row[j]);
  const all = data.flat(), G = mean(all), N = all.length;
  const sampleM = rowsN(a).map((i) => mean(data.slice(i * reps, (i + 1) * reps).flat()));
  const colM = rowsN(b).map((j) => mean(data.map((row) => row[j])));
  const ssA = b * reps * sampleM.reduce((s, m) => s + (m - G) ** 2, 0);
  const ssB = a * reps * colM.reduce((s, m) => s + (m - G) ** 2, 0);
  let ssAB = 0, ssW = 0;
  for (let i = 0; i < a; i++) for (let j = 0; j < b; j++) { const x = cell(i, j), m = mean(x); ssAB += reps * (m - sampleM[i] - colM[j] + G) ** 2; ssW += x.reduce((s, v) => s + (v - m) ** 2, 0); }
  const dfA = a - 1, dfB = b - 1, dfAB = dfA * dfB, dfW = a * b * (reps - 1);
  const msW = ssW / dfW;
  const row = (name: string, ss: number, df: number): Scalar[] => { const F = (ss / df) / msW; return [name, fin(ss), df, fin(ss / df), fin(F), fin(fSf(F, df, dfW)), fin(fInvUpper(alpha, df, dfW))]; };
  const m: Scalar[][] = [['Anova: Two-Factor With Replication'], []];
  const bold = [0];
  const block = (label: string, cols: number[][]) => {
    m.push(['SUMMARY', ...colLabels, 'Total']); bold.push(m.length - 1);
    m.push([label]);
    const tot = cols.flat();
    m.push(['Count', ...cols.map((x) => x.length), tot.length]);
    m.push(['Sum', ...cols.map((x) => fin(x.reduce((s, v) => s + v, 0))), fin(tot.reduce((s, v) => s + v, 0))]);
    m.push(['Average', ...cols.map((x) => fin(mean(x))), fin(mean(tot))]);
    m.push(['Variance', ...cols.map((x) => (x.length > 1 ? fin(variance(x)) : DIV0())), fin(variance(tot))]);
    m.push([]);
  };
  for (let i = 0; i < a; i++) block(sampleLabels[i] || `Sample ${i + 1}`, rowsN(b).map((j) => cell(i, j)));
  m.push(['Total']);
  m.push(['Count', ...rowsN(b).map((j) => data.length)]);
  m.push(['Sum', ...rowsN(b).map((j) => fin(data.reduce((s, r) => s + r[j], 0)))]);
  m.push(['Average', ...colM.map(fin)]);
  m.push(['Variance', ...rowsN(b).map((j) => fin(variance(data.map((r) => r[j]))))]);
  m.push([], [], ['ANOVA'], ['Source of Variation', 'SS', 'df', 'MS', 'F', 'P-value', 'F crit']);
  bold.push(m.length - 2, m.length - 1);
  m.push(row('Sample', ssA, dfA), row('Columns', ssB, dfB), row('Interaction', ssAB, dfAB), ['Within', fin(ssW), dfW, fin(msW)], [], ['Total', fin(ssA + ssB + ssAB + ssW), N - 1]);
  return { m, bold };
}

function rowsN(n: number): number[] { return Array.from({ length: n }, (_, i) => i); }

export function movingAverage(data: number[], interval: number, stdErrors = false): ToolOutput {
  const out: Scalar[][] = [];
  const f: number[] = [];
  for (let t = 0; t < data.length; t++) {
    if (t < interval - 1) { out.push([NA()]); f.push(NaN); continue; }
    const w = data.slice(t - interval + 1, t + 1);
    const v = mean(w); f.push(v);
    out.push([fin(v)]);
  }
  if (stdErrors) out.forEach((row, t) => {
    if (t < 2 * interval - 2) { row.push(NA()); return; }
    let s = 0; for (let k = t - interval + 1; k <= t; k++) s += (data[k] - f[k]) ** 2;
    row.push(fin(Math.sqrt(s / interval)));
  });
  return { m: out };
}

export function exponentialSmoothing(data: number[], damping: number, stdErrors = false): ToolOutput {
  const a = 1 - damping;
  const f: number[] = [NaN];
  const out: Scalar[][] = [[NA()]];
  for (let t = 1; t < data.length; t++) {
    const v = t === 1 ? data[0] : a * data[t - 1] + damping * f[t - 1];
    f.push(v); out.push([fin(v)]);
  }
  if (stdErrors) out.forEach((row, t) => {
    if (t < 4) { row.push(NA()); return; }
    let s = 0; for (let k = t - 2; k <= t; k++) s += (data[k] - f[k]) ** 2;
    row.push(fin(Math.sqrt(s / 3)));
  });
  return { m: out };
}

export function rankAndPercentile(series: Series[]): ToolOutput {
  const m: Scalar[][] = [];
  series.forEach((s, k) => {
    const j = k * 4;
    const put = (i: number, c: number, v: Scalar) => { while (m.length <= i) m.push([]); const row = m[i]; while (row.length < c) row.push(null); row[c] = v; };
    put(0, j, 'Point'); put(0, j + 1, s.label); put(0, j + 2, 'Rank'); put(0, j + 3, 'Percent');
    const n = s.data.length;
    const sortedAsc = [...s.data].sort((a, b) => a - b);
    const order = s.data.map((v, i) => i).sort((a, b) => s.data[b] - s.data[a] || a - b);
    order.forEach((i, r) => {
      const v = s.data[i];
      const rank = s.data.filter((x) => x > v).length + 1;
      const below = sortedAsc.filter((x) => x < v).length;
      const pct = n > 1 ? Math.floor((below / (n - 1)) * 1000) / 1000 : 1;
      put(r + 1, j, i + 1); put(r + 1, j + 1, v); put(r + 1, j + 2, rank); put(r + 1, j + 3, pct);
    });
  });
  return { m, bold: [0] };
}

/** Seeded random generator (mulberry32); seed 0 means random. */
export function rng(seed = 0): () => number {
  let a = seed ? Math.floor(seed) >>> 0 : (Math.random() * 2 ** 32) >>> 0;
  return () => { a |= 0; a = (a + 0x6d2b79f5) | 0; let t = Math.imul(a ^ (a >>> 15), 1 | a); t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t; return ((t ^ (t >>> 14)) >>> 0) / 4294967296; };
}

export type Dist = { kind: 'uniform'; a: number; b: number } | { kind: 'normal'; mean: number; sd: number } | { kind: 'bernoulli'; p: number } | { kind: 'binomial'; p: number; n: number }
  | { kind: 'poisson'; lambda: number } | { kind: 'patterned'; from: number; to: number; step: number; repeatEach: number; repeatSeq: number } | { kind: 'discrete'; values: number[]; probs: number[] };

export function randomNumbers(vars: number, count: number, d: Dist, seed = 0): ToolOutput {
  const R = rng(seed);
  const normal = () => { let u = 0; while (!u) u = R(); return Math.sqrt(-2 * Math.log(u)) * Math.cos(2 * Math.PI * R()); };
  const draw = (): number => {
    switch (d.kind) {
      case 'uniform': return d.a + (d.b - d.a) * R();
      case 'normal': return d.mean + d.sd * normal();
      case 'bernoulli': return R() < d.p ? 1 : 0;
      case 'binomial': { let k = 0; for (let i = 0; i < d.n; i++) if (R() < d.p) k++; return k; }
      case 'poisson': { const L = Math.exp(-d.lambda); let k = 0, p = 1; do { k++; p *= R(); } while (p > L && k < 1e6); return k - 1; }
      case 'discrete': { const u = R() * d.probs.reduce((s, x) => s + x, 0); let acc = 0; for (let i = 0; i < d.values.length; i++) { acc += d.probs[i]; if (u <= acc) return d.values[i]; } return d.values[d.values.length - 1]; }
      default: return 0;
    }
  };
  if (d.kind === 'patterned') {
    const seq: number[] = [];
    for (let s = 0; s < Math.max(1, d.repeatSeq); s++) for (let v = d.from; d.step > 0 ? v <= d.to + 1e-12 : v >= d.to - 1e-12; v += d.step || 1) for (let e = 0; e < Math.max(1, d.repeatEach); e++) seq.push(+v.toPrecision(15));
    return { m: seq.map((v) => rowsN(Math.max(1, vars)).map(() => v)) };
  }
  const m: Scalar[][] = [];
  for (let i = 0; i < count; i++) { const row: Scalar[] = []; for (let j = 0; j < vars; j++) row.push(+draw().toPrecision(15)); m.push(row); }
  return { m };
}

export function sampling(data: number[], method: 'periodic' | 'random', n: number, seed = 0): ToolOutput {
  if (method === 'periodic') { const out: Scalar[][] = []; for (let i = n - 1; i < data.length; i += n) out.push([data[i]]); return { m: out }; }
  const R = rng(seed);
  return { m: rowsN(n).map(() => [data[Math.floor(R() * data.length)]]) };
}

export type Complex = [number, number];

const NUM = '(?:\\d+\\.?\\d*|\\.\\d+)(?:[eE][+-]?\\d+)?';
const RE_REAL = new RegExp(`^[+-]?${NUM}$`), RE_IMAG = new RegExp(`^([+-]?)(${NUM})?[ij]$`), RE_BOTH = new RegExp(`^([+-]?${NUM})([+-])(${NUM})?[ij]$`);

export function parseComplex(v: Scalar): Complex | null {
  if (typeof v === 'number') return [v, 0];
  if (typeof v !== 'string') return null;
  const s = v.replace(/\s/g, '');
  if (RE_REAL.test(s)) return [+s, 0];
  let m = RE_IMAG.exec(s);
  if (m) return [0, (m[1] === '-' ? -1 : 1) * (m[2] ? +m[2] : 1)];
  m = RE_BOTH.exec(s);
  if (m) return [+m[1], (m[2] === '-' ? -1 : 1) * (m[3] ? +m[3] : 1)];
  return null;
}

export function formatComplex([re, im]: Complex): string {
  const f = (x: number) => String(+x.toPrecision(15));
  const r = Math.abs(re) < 1e-13 ? 0 : re, i = Math.abs(im) < 1e-13 ? 0 : im;
  if (!i) return f(r);
  const is = i === 1 ? '' : i === -1 ? '-' : f(i);
  if (!r) return `${is}i`;
  return `${f(r)}${i > 0 ? '+' : ''}${is}i`;
}

export function fft(x: Complex[], inverse = false): Complex[] {
  const n = x.length;
  if (n & (n - 1)) throw new Error('Fourier analysis needs 2, 4, 8, … up to 4096 values.');
  const a = x.map((c) => [...c] as Complex);
  for (let i = 1, j = 0; i < n; i++) { let bit = n >> 1; for (; j & bit; bit >>= 1) j ^= bit; j ^= bit; if (i < j) [a[i], a[j]] = [a[j], a[i]]; }
  for (let len = 2; len <= n; len <<= 1) {
    const ang = ((inverse ? 2 : -2) * Math.PI) / len;
    for (let i = 0; i < n; i += len) for (let k = 0; k < len / 2; k++) {
      const wr = Math.cos(ang * k), wi = Math.sin(ang * k);
      const [ur, ui] = a[i + k], [vr0, vi0] = a[i + k + len / 2];
      const vr = vr0 * wr - vi0 * wi, vi = vr0 * wi + vi0 * wr;
      a[i + k] = [ur + vr, ui + vi]; a[i + k + len / 2] = [ur - vr, ui - vi];
    }
  }
  return inverse ? a.map(([r, i]) => [r / n, i / n]) : a;
}

export function fourier(values: Scalar[], inverse = false): ToolOutput {
  const xs = values.map(parseComplex);
  if (xs.some((x) => !x)) throw new Error('Input range contains non-numeric data.');
  if (xs.length > 4096) throw new Error('Fourier analysis is limited to 4096 values.');
  return { m: fft(xs as Complex[], inverse).map((c) => [formatComplex(c)]) };
}

// --------------------------------------------------------------------------------------------
// Input reading and the dialog

type Grouping = 'columns' | 'rows';

function readSeries(sheet: Sheet, rg: { r1: number; c1: number; r2: number; c2: number }, grouped: Grouping, labels: boolean, prefix = 'Column'): Series[] {
  const out: Series[] = [];
  const lines = grouped === 'columns' ? rowsN(rg.c2 - rg.c1 + 1).map((i) => rg.c1 + i) : rowsN(rg.r2 - rg.r1 + 1).map((i) => rg.r1 + i);
  lines.forEach((line, k) => {
    const cells = grouped === 'columns' ? rowsN(rg.r2 - rg.r1 + 1).map((i) => sheet.value(rg.r1 + i, line)) : rowsN(rg.c2 - rg.c1 + 1).map((i) => sheet.value(line, rg.c1 + i));
    const label = labels ? String(cells[0] ?? '') : `${grouped === 'columns' ? prefix : 'Row'} ${k + 1}`;
    const vals = (labels ? cells.slice(1) : cells);
    if (vals.some((v) => typeof v === 'string' && v.trim() !== '')) throw new Error(`Input range contains non-numeric data (${label}). Check "Labels" if the first ${grouped === 'columns' ? 'row' : 'column'} holds names.`);
    out.push({ label, data: vals.filter((v): v is number => typeof v === 'number') });
  });
  return out;
}

const TOOLS = [
  'Anova: Single Factor', 'Anova: Two-Factor With Replication', 'Anova: Two-Factor Without Replication', 'Correlation', 'Covariance', 'Descriptive Statistics',
  'Exponential Smoothing', 'F-Test Two-Sample for Variances', 'Fourier Analysis', 'Histogram', 'Moving Average', 'Random Number Generation', 'Rank and Percentile',
  'Regression', 'Sampling', 't-Test: Paired Two Sample for Means', 't-Test: Two-Sample Assuming Equal Variances', 't-Test: Two-Sample Assuming Unequal Variances', 'z-Test: Two Sample for Means',
] as const;

const outputFields = (): Field[] => [
  { key: 'out', type: 'radio', label: 'Output options', options: [{ value: 'sheet', label: 'New worksheet' }, { value: 'range', label: 'Output range' }], value: 'sheet' },
  { key: 'outRange', type: 'text', label: 'Output range (top-left cell)', value: '', visibleIf: (v) => v.out === 'range' },
  { key: 'sheetName', type: 'text', label: 'New sheet name', value: '', placeholder: 'optional', visibleIf: (v) => v.out === 'sheet' },
];

async function toolDialog(app: App, tool: string): Promise<void> {
  const sheet = app.sheet, book = app.book;
  const sel = activeRange(app);
  const region = sel.r1 === sel.r2 && sel.c1 === sel.c2 ? currentRegion(sheet, sel.r1, sel.c1) : clipToUsed(sheet, sel);
  const defIn = rangeRefText(sheet, region, true, sheet);
  const hasLabels = typeof sheet.value(region.r1, region.c1) === 'string' || typeof sheet.value(region.r1, region.c2) === 'string';
  const inField = (key = 'input', label = 'Input range', value = defIn): Field => ({ key, type: 'text', label, value });
  const grouped: Field = { key: 'grouped', type: 'radio', label: 'Grouped by', options: [{ value: 'columns', label: 'Columns' }, { value: 'rows', label: 'Rows' }], value: 'columns' };
  const labels = (label = 'Labels in first row'): Field => ({ key: 'labels', type: 'checkbox', label, value: hasLabels });
  const alpha: Field = { key: 'alpha', type: 'number', label: 'Alpha', value: 0.05, step: 0.01, min: 0.0001, max: 0.9999, half: true };
  const twoRanges = (): Field[] => [
    inField('in1', 'Variable 1 range', rangeRefText(sheet, { ...region, c2: region.c1 }, true, sheet)),
    inField('in2', 'Variable 2 range', rangeRefText(sheet, { ...region, c1: Math.min(region.c2, region.c1 + 1), c2: Math.min(region.c2, region.c1 + 1) }, true, sheet)),
  ];
  let fields: Field[] = [];
  switch (tool) {
    case 'Descriptive Statistics': fields = [inField(), grouped, labels('Labels in first row/column'), { key: 'summary', type: 'checkbox', label: 'Summary statistics', value: true }, { key: 'confOn', type: 'checkbox', label: 'Confidence Level for Mean', value: true, half: true }, { key: 'conf', type: 'number', label: '%', value: 95, half: true }, { key: 'kL', type: 'number', label: 'Kth Largest', value: '' as any, half: true }, { key: 'kS', type: 'number', label: 'Kth Smallest', value: '' as any, half: true }]; break;
    case 'Histogram': fields = [inField(), { key: 'bins', type: 'text', label: 'Bin range (optional)', value: '' }, labels(), { key: 'pareto', type: 'checkbox', label: 'Pareto (sorted histogram)', value: false }, { key: 'cum', type: 'checkbox', label: 'Cumulative Percentage', value: false }, { key: 'chart', type: 'checkbox', label: 'Chart output', value: true }]; break;
    case 'Correlation': case 'Covariance': case 'Rank and Percentile': fields = [inField(), grouped, labels('Labels in first row/column')]; break;
    case 'Anova: Single Factor': fields = [inField(), grouped, labels('Labels in first row/column'), alpha]; break;
    case 'Anova: Two-Factor Without Replication': fields = [inField(), { key: 'labels', type: 'checkbox', label: 'Labels (first row and column)', value: hasLabels }, alpha]; break;
    case 'Anova: Two-Factor With Replication': fields = [inField(), { key: 'reps', type: 'number', label: 'Rows per sample', value: 2, min: 1, half: true }, alpha, { key: 'i', type: 'info', label: 'The first row and the first column of the input range must contain labels.' }]; break;
    case 'Regression': fields = [inField('y', 'Input Y range', rangeRefText(sheet, { ...region, c1: region.c2 }, true, sheet)), inField('x', 'Input X range', rangeRefText(sheet, { ...region, c2: Math.max(region.c1, region.c2 - 1) }, true, sheet)), labels('Labels'), { key: 'const0', type: 'checkbox', label: 'Constant is zero', value: false }, { key: 'conf', type: 'number', label: 'Confidence level (%)', value: 95, half: true }, { key: 'resid', type: 'checkbox', label: 'Residuals', value: false, half: true }, { key: 'stdResid', type: 'checkbox', label: 'Standardized residuals', value: false, half: true }]; break;
    case 't-Test: Paired Two Sample for Means': case 't-Test: Two-Sample Assuming Equal Variances': case 't-Test: Two-Sample Assuming Unequal Variances':
      fields = [...twoRanges(), { key: 'hmd', type: 'number', label: 'Hypothesized Mean Difference', value: 0, half: true }, alpha, labels('Labels')]; break;
    case 'z-Test: Two Sample for Means': fields = [...twoRanges(), { key: 'hmd', type: 'number', label: 'Hypothesized Mean Difference', value: 0 }, { key: 'v1', type: 'number', label: 'Variable 1 variance (known)', value: 1, half: true }, { key: 'v2', type: 'number', label: 'Variable 2 variance (known)', value: 1, half: true }, alpha, labels('Labels')]; break;
    case 'F-Test Two-Sample for Variances': fields = [...twoRanges(), alpha, labels('Labels')]; break;
    case 'Moving Average': fields = [inField(), labels('Labels in first row'), { key: 'interval', type: 'number', label: 'Interval', value: 3, min: 2, half: true }, { key: 'stderr', type: 'checkbox', label: 'Standard errors', value: false, half: true }]; break;
    case 'Exponential Smoothing': fields = [inField(), { key: 'damping', type: 'number', label: 'Damping factor', value: 0.3, min: 0, max: 1, step: 0.05, half: true }, labels('Labels'), { key: 'stderr', type: 'checkbox', label: 'Standard errors', value: false, half: true }]; break;
    case 'Sampling': fields = [inField(), labels('Labels'), { key: 'method', type: 'radio', label: 'Sampling method', options: [{ value: 'periodic', label: 'Periodic' }, { value: 'random', label: 'Random' }], value: 'random' }, { key: 'n', type: 'number', label: 'Period / Number of samples', value: 10, min: 1 }]; break;
    case 'Fourier Analysis': fields = [inField(), labels('Labels in first row'), { key: 'inverse', type: 'checkbox', label: 'Inverse', value: false }]; break;
    case 'Random Number Generation': fields = [
      { key: 'vars', type: 'number', label: 'Number of variables', value: 1, min: 1, half: true }, { key: 'count', type: 'number', label: 'Number of random numbers', value: 10, min: 1, half: true },
      { key: 'dist', type: 'select', label: 'Distribution', options: ['Uniform', 'Normal', 'Bernoulli', 'Binomial', 'Poisson', 'Patterned', 'Discrete'], value: 'Uniform' },
      { key: 'p1', type: 'number', label: 'Parameter 1 (Uniform: between · Normal: mean · Bernoulli/Binomial: p · Poisson: λ · Patterned: from)', value: 0 },
      { key: 'p2', type: 'number', label: 'Parameter 2 (Uniform: and · Normal: std. dev. · Binomial: trials · Patterned: to)', value: 1 },
      { key: 'p3', type: 'number', label: 'Patterned: step', value: 1, half: true, visibleIf: (v) => v.dist === 'Patterned' }, { key: 'p4', type: 'number', label: 'Repeat each number', value: 1, half: true, visibleIf: (v) => v.dist === 'Patterned' },
      { key: 'p5', type: 'number', label: 'Repeat the sequence', value: 1, visibleIf: (v) => v.dist === 'Patterned' },
      { key: 'disc', type: 'text', label: 'Value and probability input range (2 columns)', value: '', visibleIf: (v) => v.dist === 'Discrete' },
      { key: 'seed', type: 'number', label: 'Random seed (optional)', value: '' as any },
    ]; break;
  }
  const res = await openDialog({ title: tool, fields: [...fields, ...outputFields()], wide: true });
  if (!res) return;
  let out: ToolOutput;
  const rng0 = (key: string) => { const x = resolveRangeText(book, String(res[key] || ''), sheet); if (!x) throw new Error(`The ${key === 'y' ? 'Y' : key === 'x' ? 'X' : 'input'} range is not valid.`); return { sheet: x.sheet, range: clipToUsed(x.sheet, x.range) }; };
  const one = (key: string) => { const x = rng0(key); const s = readSeries(x.sheet, x.range, x.range.c2 > x.range.c1 && x.range.r2 === x.range.r1 ? 'rows' : 'columns', !!res.labels, 'Variable'); return s[0]; };
  try {
    switch (tool) {
      case 'Descriptive Statistics': { const x = rng0('input'); out = descriptiveStatistics(readSeries(x.sheet, x.range, res.grouped, !!res.labels), { summary: !!res.summary, conf: res.confOn ? (res.conf || 95) / 100 : null, kLarge: res.kL || null, kSmall: res.kS || null }); break; }
      case 'Histogram': {
        const x = rng0('input');
        const data = readSeries(x.sheet, x.range, 'columns', !!res.labels).flatMap((s) => s.data);
        let bins: number[] | null = null;
        if (String(res.bins || '').trim()) { const b = rng0('bins'); bins = readSeries(b.sheet, b.range, 'columns', !!res.labels).flatMap((s) => s.data); }
        out = histogram(data, bins, { pareto: !!res.pareto, cumulative: !!res.cum, chart: !!res.chart });
        break;
      }
      case 'Correlation': case 'Covariance': { const x = rng0('input'); out = correlationMatrix(readSeries(x.sheet, x.range, res.grouped, !!res.labels), tool === 'Covariance'); break; }
      case 'Rank and Percentile': { const x = rng0('input'); out = rankAndPercentile(readSeries(x.sheet, x.range, res.grouped, !!res.labels)); break; }
      case 'Anova: Single Factor': { const x = rng0('input'); out = anovaSingle(readSeries(x.sheet, x.range, res.grouped, !!res.labels), res.alpha); break; }
      case 'Anova: Two-Factor Without Replication': {
        const x = rng0('input'); const L = !!res.labels; const rg = x.range;
        const data: number[][] = [];
        for (let r = rg.r1 + (L ? 1 : 0); r <= rg.r2; r++) data.push(rowsN(rg.c2 - rg.c1 + 1 - (L ? 1 : 0)).map((j) => Number(x.sheet.value(r, rg.c1 + (L ? 1 : 0) + j))));
        if (data.flat().some((v) => !Number.isFinite(v))) throw new Error('Input range contains non-numeric data.');
        const rl = rowsN(data.length).map((i) => (L ? String(x.sheet.value(rg.r1 + 1 + i, rg.c1) ?? '') : `Row ${i + 1}`));
        const cl = rowsN(data[0].length).map((j) => (L ? String(x.sheet.value(rg.r1, rg.c1 + 1 + j) ?? '') : `Column ${j + 1}`));
        out = anovaTwoNoRep(data, rl, cl, res.alpha); break;
      }
      case 'Anova: Two-Factor With Replication': {
        const x = rng0('input'); const rg = x.range; const reps = Math.max(1, Math.round(res.reps || 1));
        const data: number[][] = [];
        for (let r = rg.r1 + 1; r <= rg.r2; r++) data.push(rowsN(rg.c2 - rg.c1).map((j) => Number(x.sheet.value(r, rg.c1 + 1 + j))));
        if (data.flat().some((v) => !Number.isFinite(v))) throw new Error('Input range contains non-numeric data.');
        const sl = rowsN(Math.floor(data.length / reps)).map((i) => String(x.sheet.value(rg.r1 + 1 + i * reps, rg.c1) ?? `Sample ${i + 1}`));
        const cl = rowsN(rg.c2 - rg.c1).map((j) => String(x.sheet.value(rg.r1, rg.c1 + 1 + j) ?? ''));
        out = anovaTwoRep(data, reps, sl, cl, res.alpha); break;
      }
      case 'Regression': {
        const y = rng0('y'), xr = rng0('x');
        const ys = readSeries(y.sheet, y.range, 'columns', !!res.labels)[0];
        const xs = readSeries(xr.sheet, xr.range, 'columns', !!res.labels, 'X Variable');
        if (xs.some((s) => s.data.length !== ys.data.length)) throw new Error('Input ranges must have the same number of rows (and no blank cells).');
        if (xs.length > 16) throw new Error('Regression accepts at most 16 X columns.');
        if (!res.labels) { ys.label = 'Y'; xs.forEach((s, i) => { s.label = `X Variable ${xs.length > 1 ? i + 1 : ''}`.trim(); }); }
        out = regression(ys, xs, { constZero: !!res.const0, conf: (res.conf || 95) / 100, residuals: !!res.resid, stdResiduals: !!res.stdResid }); break;
      }
      case 't-Test: Paired Two Sample for Means': case 't-Test: Two-Sample Assuming Equal Variances': case 't-Test: Two-Sample Assuming Unequal Variances': {
        const a = one('in1'), b = one('in2');
        if (!res.labels) { a.label = 'Variable 1'; b.label = 'Variable 2'; }
        out = tTest(a, b, tool.includes('Paired') ? 'paired' : tool.includes('Unequal') ? 'unequal' : 'equal', res.hmd || 0, res.alpha); break;
      }
      case 'z-Test: Two Sample for Means': { const a = one('in1'), b = one('in2'); if (!res.labels) { a.label = 'Variable 1'; b.label = 'Variable 2'; } out = zTest(a, b, res.v1, res.v2, res.hmd || 0, res.alpha); break; }
      case 'F-Test Two-Sample for Variances': { const a = one('in1'), b = one('in2'); if (!res.labels) { a.label = 'Variable 1'; b.label = 'Variable 2'; } out = fTest(a, b, res.alpha); break; }
      case 'Moving Average': { const s = one('input'); out = movingAverage(s.data, Math.max(2, Math.round(res.interval || 3)), !!res.stderr); break; }
      case 'Exponential Smoothing': { const s = one('input'); out = exponentialSmoothing(s.data, Math.min(1, Math.max(0, res.damping ?? 0.3)), !!res.stderr); break; }
      case 'Sampling': { const x = rng0('input'); const data = readSeries(x.sheet, x.range, 'columns', !!res.labels).flatMap((s) => s.data); out = sampling(data, res.method, Math.max(1, Math.round(res.n || 1)), res.seed || 0); break; }
      case 'Fourier Analysis': {
        const x = rng0('input'); const vals: Scalar[] = [];
        for (let r = x.range.r1 + (res.labels ? 1 : 0); r <= x.range.r2; r++) for (let c = x.range.c1; c <= x.range.c2; c++) vals.push(x.sheet.value(r, c));
        out = fourier(vals, !!res.inverse); break;
      }
      case 'Random Number Generation': {
        const k = res.dist as string;
        let d: Dist;
        if (k === 'Uniform') d = { kind: 'uniform', a: res.p1 ?? 0, b: res.p2 ?? 1 };
        else if (k === 'Normal') d = { kind: 'normal', mean: res.p1 ?? 0, sd: res.p2 ?? 1 };
        else if (k === 'Bernoulli') d = { kind: 'bernoulli', p: res.p1 ?? 0.5 };
        else if (k === 'Binomial') d = { kind: 'binomial', p: res.p1 ?? 0.5, n: Math.round(res.p2 ?? 10) };
        else if (k === 'Poisson') d = { kind: 'poisson', lambda: res.p1 ?? 1 };
        else if (k === 'Patterned') d = { kind: 'patterned', from: res.p1 ?? 1, to: res.p2 ?? 10, step: res.p3 || 1, repeatEach: res.p4 || 1, repeatSeq: res.p5 || 1 };
        else {
          const x = rng0('disc');
          const values: number[] = [], probs: number[] = [];
          for (let r = x.range.r1; r <= x.range.r2; r++) { values.push(Number(x.sheet.value(r, x.range.c1))); probs.push(Number(x.sheet.value(r, x.range.c1 + 1))); }
          if (Math.abs(probs.reduce((s, p) => s + p, 0) - 1) > 1e-6) throw new Error('The probabilities in the value and probability range must add up to 1.');
          d = { kind: 'discrete', values, probs };
        }
        if ((d.kind === 'bernoulli' || d.kind === 'binomial') && (d.p < 0 || d.p > 1)) throw new Error('The probability must be between 0 and 1.');
        out = randomNumbers(Math.max(1, Math.round(res.vars || 1)), Math.max(1, Math.round(res.count || 1)), d, res.seed || 0); break;
      }
      default: return;
    }
  } catch (e) { void alertDialog(tool, (e as Error).message); return; }
  await writeOutput(app, out, res, tool);
}

async function writeOutput(app: App, out: ToolOutput, res: Record<string, any>, tool: string): Promise<void> {
  const book = app.book;
  let target: Sheet = app.sheet, r0 = 0, c0 = 0;
  if (res.out === 'range') {
    const x = resolveRangeText(book, String(res.outRange || ''), app.sheet);
    if (!x) { toast('The output range is not valid.'); return; }
    target = x.sheet; r0 = x.range.r1; c0 = x.range.c1;
    let occupied = false;
    out.m.forEach((row, i) => row.forEach((_, j) => { const cell = target.get(r0 + i, c0 + j); if (cell && (cell.v != null || cell.f != null)) occupied = true; }));
    if (occupied && !(await confirmDialog(tool, 'Output range will overwrite existing data. Press OK to overwrite.'))) return;
  }
  book.transact(tool, () => {
    if (res.out === 'range') book.recordSheet(target); // the chart added below is part of the sheet, not of a cell
    if (res.out !== 'range') target = addSheetAfter(book, String(res.sheetName || '').trim() || tool.replace(/[:]/g, ''));
    writeMatrix(target, r0, c0, out.m as Matrix);
    for (const i of out.bold || []) out.m[i]?.forEach((v, j) => { if (v != null) { const cell = target.get(r0 + i, c0 + j); if (cell) target.set(r0 + i, c0 + j, { ...cell, s: book.derivedStyle(cell.s ?? 0, { bold: true }) }); } });
    if (res.out !== 'range') { target.colW.set(c0, 150); }
    if (out.chart) {
      const ch = out.chart;
      const ref = (r1: number, r2: number, c: number) => rangeRefText(target, { r1: r0 + r1, c1: c0 + c, r2: r0 + r2, c2: c0 + c });
      target.charts.push({ id: newId('ch'), type: ch.type, title: ch.title, series: ch.vals.map(([a, b, c]) => ({ name: String(out.m[0][c] ?? 'Frequency'), values: ref(a, b, c), categories: ref(ch.cats[0], ch.cats[1], ch.cats[2]) })), at: { r: r0, c: c0 + (out.m[0]?.length || 3) + 1, dx: 0, dy: 0, w: 420, h: 260 }, legend: 'none', gridlines: true });
    }
  });
  if (target !== app.sheet) app.exec('sheet.activate', book.sheets.indexOf(target));
}

async function analysisTools(app: App): Promise<void> {
  if (blocked(app)) return;
  const res = await openDialog({ title: 'Data Analysis', fields: [{ key: 'tool', type: 'list', label: 'Analysis Tools', options: [...TOOLS], value: 'Descriptive Statistics', rows: 12 }] });
  if (!res?.tool) return;
  await toolDialog(app, res.tool);
}

def([
  { id: 'data.analysisTools', label: 'Data Analysis', icon: 'query_stats', tab: 'Data', group: 'Analyze', edits: true, keywords: 'analysis toolpak regression anova t-test histogram descriptive statistics correlation fourier sampling', run: (app, arg?: string) => (arg && (TOOLS as readonly string[]).includes(arg) ? toolDialog(app, arg) : analysisTools(app)) },
]);
