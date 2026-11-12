/** Logical functions, LET/LAMBDA and the LAMBDA helper functions (MAP, REDUCE, SCAN, BYROW, BYCOL, MAKEARRAY). */
import { defFn, getFn } from './registry';
import {
  bool, fail, isErr, isLambda, isMatrix, isRef, iterValues, toMatrix, toScalar, num, int, compare, ERR, CellError,
  type Ctx, type Value, type LambdaV,
} from '../values';
import type { Node } from '../parser';
import type { Matrix, Scalar } from '../../model/types';

/**
 * The function argument of a LAMBDA helper: a LAMBDA, or the bare name of a function (BYROW(range, SUM)), which
 * arrives as #NAME? carrying the name and stands for a LAMBDA that passes its arguments on to that function.
 */
function lambdaArg(v: Value | undefined, arity: number): LambdaV | null {
  if (v && isLambda(v)) return v;
  const e = v as unknown as { code?: string; detail?: string };
  if (v instanceof CellError && e.code === '#NAME?' && e.detail && getFn(e.detail)) {
    const params = Array.from({ length: arity }, (_, i) => `_X${i}`);
    return { k: 'lambda', params, body: { t: 'fn', name: String(e.detail).toUpperCase(), args: params.map((name) => ({ t: 'name', name })) }, scope: null } as LambdaV;
  }
  return null;
}

/** One result of a LAMBDA helper's function: a single value; an array of several is #CALC!, not its first cell. */
function oneValue(v: Value, ctx: Ctx): Scalar {
  if (isMatrix(v) && (v.length > 1 || (v[0]?.length ?? 0) > 1)) return ERR.calc;
  if (isRef(v) && (v.r2 > v.r1 || v.c2 > v.c1)) return ERR.calc;
  return scalarOf(v, ctx);
}

/** Evaluates a node to a scalar for a condition (single cell or 1x1). */
function evalScalar(n: Node, ctx: Ctx): Scalar {
  const v = ctx.evalNode(n);
  return toScalar(v, ctx);
}

/** Element-wise IF when the condition is an array. */
function ifArray(cond: Value, ctx: Ctx, t: Node | undefined, f: Node | undefined): Value {
  const m = toMatrix(cond, ctx);
  const tv = t ? ctx.evalNode(t) : true;
  const fv = f ? ctx.evalNode(f) : false;
  const pick = (v: Value, r: number, c: number): Scalar => {
    if (isMatrix(v) || isRef(v)) {
      const mm = toMatrix(v, ctx);
      const rr = mm.length === 1 ? 0 : r, cc = (mm[0]?.length ?? 1) === 1 ? 0 : c;
      return mm[rr]?.[cc] ?? ERR.na;
    }
    return toScalar(v, ctx);
  };
  return m.map((row, r) => row.map((x, c) => {
    if (isErr(x)) return x;
    try { return bool(x) ? pick(tv, r, c) : pick(fv, r, c); } catch (e) { if ((e as any)?.err instanceof CellError) return ERR.value; throw e; }
  }));
}

defFn([
  {
    name: 'IF', min: 2, max: 3, category: 'Logical', desc: 'Returns one value if a condition is TRUE and another if it is FALSE.', params: ['logical_test', 'value_if_true', '[value_if_false]'],
    lazy: (n, ctx) => {
      const cv = ctx.evalNode(n[0]);
      if ((isRef(cv) && (cv.r1 !== cv.r2 || cv.c1 !== cv.c2)) || (isMatrix(cv) && (cv.length > 1 || cv[0].length > 1))) return ifArray(cv, ctx, n[1], n[2]);
      const c = toScalar(cv, ctx);
      if (isErr(c)) return c;
      if (bool(c)) return n[1].t === 'missing' ? 0 : ctx.evalNode(n[1]);
      if (n.length < 3) return false;
      return n[2].t === 'missing' ? 0 : ctx.evalNode(n[2]);
    },
  },
  {
    name: 'IFS', min: 2, max: 254, category: 'Logical', desc: 'Checks conditions in order and returns the value for the first TRUE one.', params: ['logical_test1', 'value_if_true1', '...'],
    lazy: (n, ctx) => {
      if (n.length % 2) return ERR.value;
      for (let i = 0; i < n.length; i += 2) {
        const c = evalScalar(n[i], ctx);
        if (isErr(c)) return c;
        if (bool(c)) return ctx.evalNode(n[i + 1]);
      }
      return ERR.na;
    },
  },
  {
    name: 'IFERROR', min: 2, max: 2, category: 'Logical', desc: 'Returns value_if_error if the expression is an error, otherwise the value of the expression.', params: ['value', 'value_if_error'],
    lazy: (n, ctx) => {
      let v: Value;
      try { v = ctx.evalNode(n[0]); } catch (e) { if (!((e as any)?.err instanceof CellError)) throw e; v = (e as any).err; }
      if (isMatrix(v) || (isRef(v) && (v.r1 !== v.r2 || v.c1 !== v.c2))) {
        const m = toMatrix(v, ctx);
        let alt: Scalar | null = null;
        return m.map((row) => row.map((x) => (isErr(x) ? (alt ??= toScalar(ctx.evalNode(n[1]), ctx)) : x)));
      }
      const s = isRef(v) ? toScalar(v, ctx) : v;
      return isErr(s) ? ctx.evalNode(n[1]) : v;
    },
  },
  {
    name: 'IFNA', min: 2, max: 2, category: 'Logical', desc: 'Returns value_if_na if the expression is #N/A, otherwise the value.', params: ['value', 'value_if_na'],
    lazy: (n, ctx) => {
      const v = ctx.evalNode(n[0]);
      if (isMatrix(v) || (isRef(v) && (v.r1 !== v.r2 || v.c1 !== v.c2))) {
        const m = toMatrix(v, ctx);
        return m.map((row) => row.map((x) => (isErr(x) && x.code === '#N/A' ? toScalar(ctx.evalNode(n[1]), ctx) : x)));
      }
      const s = isRef(v) ? toScalar(v, ctx) : v;
      return isErr(s) && s.code === '#N/A' ? ctx.evalNode(n[1]) : v;
    },
  },
  {
    name: 'AND', min: 1, max: 255, category: 'Logical', desc: 'TRUE if all arguments are TRUE.', params: ['logical1', '[logical2]', '...'], arrays: 'all',
    impl: (args, ctx) => {
      let seen = false;
      for (const a of args) for (const { v, fromRef } of iterValues(a, ctx)) {
        if (isErr(v)) return v;
        if (v == null || (fromRef && typeof v === 'string')) continue;
        seen = true;
        if (!bool(v)) return false;
      }
      return seen ? true : ERR.value;
    },
  },
  {
    name: 'OR', min: 1, max: 255, category: 'Logical', desc: 'TRUE if any argument is TRUE.', params: ['logical1', '[logical2]', '...'], arrays: 'all',
    impl: (args, ctx) => {
      let seen = false, any = false;
      for (const a of args) for (const { v, fromRef } of iterValues(a, ctx)) {
        if (isErr(v)) return v;
        if (v == null || (fromRef && typeof v === 'string')) continue;
        seen = true;
        if (bool(v)) any = true;
      }
      return seen ? any : ERR.value;
    },
  },
  {
    name: 'XOR', min: 1, max: 254, category: 'Logical', desc: 'TRUE if an odd number of arguments are TRUE.', params: ['logical1', '[logical2]', '...'], arrays: 'all',
    impl: (args, ctx) => {
      let n = 0, seen = false;
      for (const a of args) for (const { v, fromRef } of iterValues(a, ctx)) {
        if (isErr(v)) return v;
        if (v == null || (fromRef && typeof v === 'string')) continue;
        seen = true;
        if (bool(v)) n++;
      }
      return seen ? n % 2 === 1 : ERR.value;
    },
  },
  { name: 'NOT', min: 1, max: 1, category: 'Logical', desc: 'Reverses a logical value.', params: ['logical'], impl: ([a]) => !bool(a as Scalar) },
  { name: 'TRUE', min: 0, max: 0, category: 'Logical', desc: 'The logical value TRUE.', params: [], impl: () => true },
  { name: 'FALSE', min: 0, max: 0, category: 'Logical', desc: 'The logical value FALSE.', params: [], impl: () => false },
  {
    name: 'SWITCH', min: 3, max: 254, category: 'Logical', desc: 'Compares an expression with a list of values and returns the result for the first match.', params: ['expression', 'value1', 'result1', '[default_or_value2]', '...'],
    lazy: (n, ctx) => {
      const x = evalScalar(n[0], ctx);
      if (isErr(x)) return x;
      let i = 1;
      for (; i + 1 < n.length; i += 2) {
        const v = evalScalar(n[i], ctx);
        if (compare(x, v) === 0 && typeof x === typeof v) return ctx.evalNode(n[i + 1]);
      }
      return i < n.length ? ctx.evalNode(n[i]) : ERR.na;
    },
  },
  {
    name: 'CHOOSE', min: 2, max: 255, category: 'Lookup & Reference', desc: 'Chooses a value from a list by index.', params: ['index_num', 'value1', '[value2]', '...'], returnsRef: true,
    lazy: (n, ctx) => {
      const iv = ctx.evalNode(n[0]);
      if (isMatrix(iv) || (isRef(iv) && (iv.r1 !== iv.r2 || iv.c1 !== iv.c2))) {
        const m = toMatrix(iv, ctx);
        return m.map((row) => row.map((x) => {
          if (isErr(x)) return x;
          const i = Math.trunc(num(x));
          if (i < 1 || i >= n.length) return ERR.value;
          return toScalar(ctx.evalNode(n[i]), ctx);
        }));
      }
      const i = Math.trunc(num(toScalar(iv, ctx)));
      if (i < 1 || i >= n.length) return ERR.value;
      return ctx.evalNode(n[i]);
    },
  },
  {
    name: 'LET', min: 3, max: 253, category: 'Logical', desc: 'Assigns names to calculation results for use inside the formula.', params: ['name1', 'name_value1', 'calculation_or_name2', '...'],
    lazy: (n, ctx) => {
      if (n.length % 2 === 0) return ERR.value;
      const scope = new Map(ctx.scope || []);
      for (let i = 0; i + 1 < n.length; i += 2) {
        const nm = n[i];
        if (nm.t !== 'name') return ERR.value;
        scope.set(nm.name.toUpperCase(), ctx.evalNode(n[i + 1], scope));
      }
      return ctx.evalNode(n[n.length - 1], scope);
    },
  },
  {
    name: 'LAMBDA', min: 1, max: 254, category: 'Logical', desc: 'Creates a reusable function from parameters and a calculation.', params: ['[parameter1]', '...', 'calculation'],
    lazy: (n, ctx) => {
      const params: string[] = [];
      for (const p of n.slice(0, -1)) { if (p.t !== 'name') return ERR.value; params.push(p.name); }
      return { k: 'lambda', params, body: n[n.length - 1], scope: ctx.scope } as LambdaV;
    },
  },
  {
    name: 'ISOMITTED', min: 1, max: 1, category: 'Information', desc: 'TRUE if a LAMBDA parameter was not supplied.', params: ['argument'],
    lazy: (n, ctx) => {
      const a = n[0];
      if (a.t !== 'name') return false;
      const v = ctx.scope?.get(a.name.toUpperCase());
      return v === null || v === undefined;
    },
  },
  {
    name: 'MAP', min: 2, max: 254, category: 'Logical', desc: 'Applies a LAMBDA to each value of one or more arrays.', params: ['array1', '[array2]', '...', 'lambda'], arrays: 'all',
    impl: (args, ctx) => {
      const fn = args[args.length - 1];
      const f = lambdaArg(fn, args.length - 1);
      if (!f) return ERR.value;
      const mats = args.slice(0, -1).map((a) => toMatrix(a, ctx));
      const rows = mats[0].length, cols = mats[0][0]?.length ?? 0;
      return mats[0].map((row, r) => row.map((_, c) => oneValue(ctx.callLambda(f, mats.map((m) => (m[r]?.[c] ?? ERR.na))), ctx))).slice(0, rows).map((x) => x.slice(0, cols));
    },
  },
  {
    name: 'REDUCE', min: 3, max: 3, category: 'Logical', desc: 'Reduces an array to an accumulated value with a LAMBDA.', params: ['initial_value', 'array', 'lambda'], arrays: [1],
    impl: ([init, arr, fn], ctx) => {
      const f = lambdaArg(fn, 2);
      if (!f) return ERR.value;
      let acc: Value = init ?? 0;
      for (const row of toMatrix(arr, ctx)) for (const x of row) acc = ctx.callLambda(f, [acc, x]);
      return acc;
    },
  },
  {
    name: 'SCAN', min: 3, max: 3, category: 'Logical', desc: 'Applies a LAMBDA to each value and returns each intermediate value.', params: ['initial_value', 'array', 'lambda'], arrays: [1],
    impl: ([init, arr, fn], ctx) => {
      const f = lambdaArg(fn, 2);
      if (!f) return ERR.value;
      let acc: Value = init ?? 0;
      return toMatrix(arr, ctx).map((row) => row.map((x) => { acc = ctx.callLambda(f, [acc, x]); return oneValue(acc, ctx); }));
    },
  },
  {
    name: 'BYROW', min: 2, max: 2, category: 'Logical', desc: 'Applies a LAMBDA to each row and returns a column of results.', params: ['array', 'lambda'], arrays: [0],
    impl: ([arr, fn], ctx) => {
      const f = lambdaArg(fn, 1);
      if (!f) return ERR.value;
      return toMatrix(arr, ctx).map((row) => [oneValue(ctx.callLambda(f, [[row]]), ctx)]);
    },
  },
  {
    name: 'BYCOL', min: 2, max: 2, category: 'Logical', desc: 'Applies a LAMBDA to each column and returns a row of results.', params: ['array', 'lambda'], arrays: [0],
    impl: ([arr, fn], ctx) => {
      const f = lambdaArg(fn, 1);
      if (!f) return ERR.value;
      const m = toMatrix(arr, ctx);
      const cols = m[0]?.length ?? 0;
      const out: Scalar[] = [];
      for (let c = 0; c < cols; c++) out.push(oneValue(ctx.callLambda(f, [m.map((r) => [r[c]])]), ctx));
      return [out];
    },
  },
  {
    name: 'MAKEARRAY', min: 3, max: 3, category: 'Logical', desc: 'Builds an array of the given size by calling a LAMBDA for each position.', params: ['rows', 'columns', 'lambda'],
    impl: ([rs, cs, fn], ctx) => {
      if (!isLambda(fn)) return ERR.value;
      const rows = int(rs as Scalar), cols = int(cs as Scalar);
      if (rows < 1 || cols < 1 || rows * cols > 1e6) return ERR.value;
      const out: Matrix = [];
      for (let r = 1; r <= rows; r++) { const row: Scalar[] = []; for (let c = 1; c <= cols; c++) row.push(scalarOf(ctx.callLambda(fn, [r, c]), ctx)); out.push(row); }
      return out;
    },
  },
]);

function scalarOf(v: Value, ctx: Ctx): Scalar {
  if (isLambda(v)) return ERR.calc;
  if (isMatrix(v) || isRef(v)) return toScalar(v, ctx);
  if (v && typeof v === 'object' && (v as any).k === 'refs') return ERR.value;
  return v as Scalar;
}

void fail;
