// Regression tests for defects found in the source review.
import { describe, it, expect } from 'vitest';
import '../../src/calc/functions';
import { Workbook } from '../../src/model/workbook';
import { Engine } from '../../src/calc/engine';
import { parseCell } from '../../src/model/address';
import { CellError, type Scalar } from '../../src/model/types';
import { formatValue, formatGeneral } from '../../src/model/numfmt';
import { parseDateTime } from '../../src/model/dates';
import { shiftCells, insertCellsShift } from '../../src/model/structure';
import { Sheet } from '../../src/model/workbook';
import { translateLotus } from '../../src/heritage/lotus';
import { parseCreateTable } from '../../src/io/sqlite';
import { exportOds, importOds } from '../../src/io/ods';
import { exportXlsx } from '../../src/io/xlsx-export';
import { importXlsx } from '../../src/io/xlsx-import';

const norm = (v: Scalar): any => (v instanceof CellError ? v.code : v);

function book(data: Record<string, any> = {}) {
  const b = new Workbook();
  const e = new Engine(b);
  const s = b.sheet;
  b.transact('t', () => { for (const [a, v] of Object.entries(data)) { const p = parseCell(a)!; s.set(p.r, p.c, typeof v === 'string' && v.startsWith('=') ? { f: v.slice(1) } : { v }); } });
  e.recalc();
  const val = (f: string) => { b.transact('f', () => s.set(20, 20, { f })); e.recalc(); return norm(s.value(20, 20)); };
  return { b, e, s, val };
}

describe('ranges in array formulas', () => {
  it('keep their full size beyond the used range', () => {
    const { val } = book({ A1: 2, A2: 3 });
    expect(val('SUMPRODUCT(--(A1:A10=""))')).toBe(8);
    expect(val('SUM(A1:A5*{1;1;1;1;1})')).toBe(5);
    expect(val('SUM(IF(A1:A6="",1,0))')).toBe(4);
    expect(val('ROWS(A1:A5&"x")')).toBe(5);
    expect(val('MATCH(TRUE,A1:A6="",0)')).toBe(3);
  });
  it('still clip whole columns to the data', () => {
    const { val } = book({ A1: 2, A2: 3 });
    expect(val('SUM(A:A*2)')).toBe(10);
    expect(val('ROWS(A:A&"")')).toBeLessThan(100);
  });
});

describe('arithmetic', () => {
  it('reports an overflow as #NUM! where IFERROR can catch it', () => {
    const { val } = book({ A1: 1e308 });
    expect(val('IFERROR(A1*10,"x")')).toBe('x');
    expect(val('ISERROR(2^1024)')).toBe(true);
    expect(val('1/0')).toBe('#DIV/0!');
  });
  it('MOD stays below the divisor for huge quotients', () => {
    const { val } = book();
    expect(val('MOD(1E+20,3)')).toBeLessThan(3);
    expect(val('MOD(-7,3)')).toBe(2);
    expect(val('MOD(7,-3)')).toBe(-2);
    expect(val('MOD(10,2.5)')).toBe(0);
  });
});

describe('number formats', () => {
  it('keeps a literal point in a date format', () => {
    expect(formatValue(45000.52, 'dd.mm.yyyy').text).toBe('15.03.2023');
    expect(formatValue(45000.52, 'yyyy.mm.dd').text).toBe('2023.03.15');
  });
  it('shows nothing for an empty section', () => {
    expect(formatValue(0, '0;-0;;@').text).toBe('');
    expect(formatValue(5, '0;-0;;@').text).toBe('5');
    expect(formatValue(-5, '0;;').text).toBe('');
    expect(formatValue(5, ';;;').text).toBe('');
    expect(formatValue('t', ';;;').text).toBe('');
    expect(formatValue(5, '').text).toBe('5');
  });
  it('shows every digit of an 11-digit whole number in General', () => {
    expect(formatGeneral(12345678901)).toBe('12345678901');
    expect(formatGeneral(-12345678901)).toBe('-12345678901');
    expect(formatGeneral(1234567890.5)).toBe('1234567891');
    expect(formatGeneral(0.1 + 0.2)).toBe('0.3');
    expect(formatGeneral(123456.789)).toBe('123456.789');
  });
});

describe('typed dates and times', () => {
  it('leaves a date before 1900 as text', () => {
    expect(parseDateTime('7/4/1776')).toBeNull();
    expect(parseDateTime('1850-06-01')).toBeNull();
    expect(parseDateTime('Jan 1850')).toBeNull();
    expect(parseDateTime('7/4/1976')).not.toBeNull();
  });
  it('accepts elapsed times of 100 hours or more', () => {
    expect(parseDateTime('100:00')?.serial).toBeCloseTo(100 / 24, 10);
    expect(parseDateTime('10:30')?.serial).toBeCloseTo(10.5 / 24, 10);
  });
});

describe('table columns when sheet columns are inserted or deleted', () => {
  function withTable() {
    const { b, e, s } = book({ B1: 'Name', C1: 'Qty', B2: 'x', C2: 2, B3: 'y', C3: 3, F1: '=SUM(Sales[Qty])' });
    s.tables.push({ id: 't1', name: 'Sales', range: { r1: 0, c1: 1, r2: 2, c2: 2 }, columns: [{ name: 'Name' }, { name: 'Qty' }], style: '', opts: { header: true } } as any);
    e.recalc();
    return { b, e, s };
  }
  const names = (s: any) => s.tables[0].columns.map((c: any) => c.name);
  it('a column inserted at the left edge moves the table', () => {
    const { b, e, s } = withTable();
    b.transact('i', () => shiftCells(s, 'col', 1, 1));
    e.recalc();
    expect(names(s)).toEqual(['Name', 'Qty']);
    expect(s.tables[0].range).toMatchObject({ c1: 2, c2: 3 });
    expect(s.value(0, 6)).toBe(5);
  });
  it('a column inserted just right of the table is not part of it', () => {
    const { b, s } = withTable();
    b.transact('i', () => shiftCells(s, 'col', 3, 1));
    expect(names(s)).toEqual(['Name', 'Qty']);
  });
  it('a column inserted inside the table joins it', () => {
    const { b, s } = withTable();
    b.transact('i', () => shiftCells(s, 'col', 2, 1));
    expect(names(s)).toEqual(['Name', 'Column3', 'Qty']);
    expect(s.tables[0].range).toMatchObject({ c1: 1, c2: 3 });
  });
  it('deleting columns that start left of the table removes only the table columns inside the span', () => {
    const { b, s } = withTable();
    b.transact('d', () => shiftCells(s, 'col', 0, -2));
    expect(names(s)).toEqual(['Qty']);
    expect(s.tables[0].range).toMatchObject({ c1: 0, c2: 0 });
  });
});

describe('importers and exporters', () => {
  it('parenthesises a 1-2-3 offset before adding one', () => {
    expect(translateLotus('@CHOOSE(A1>5,"lo","hi")')).toBe('CHOOSE((A1>5)+1,"lo","hi")');
    expect(translateLotus('@CHOOSE(A1,"lo","hi")')).toBe('CHOOSE(A1+1,"lo","hi")');
  });
  it('finds the rowid column however the primary key is declared', () => {
    expect(parseCreateTable('CREATE TABLE u (id INTEGER NOT NULL, name VARCHAR, PRIMARY KEY (id))').rowidCol).toBe(0);
    expect(parseCreateTable('CREATE TABLE u ("id" integer NOT NULL PRIMARY KEY AUTOINCREMENT, name text)').rowidCol).toBe(0);
    expect(parseCreateTable('CREATE TABLE u (id INTEGER PRIMARY KEY, name TEXT)').rowidCol).toBe(0);
    expect(parseCreateTable('CREATE TABLE u (id INT PRIMARY KEY, name TEXT)').rowidCol).toBe(-1);
    expect(parseCreateTable('CREATE TABLE u (a INTEGER, b INTEGER, PRIMARY KEY (a, b))').rowidCol).toBe(-1);
    expect(parseCreateTable('CREATE TABLE u (id INTEGER PRIMARY KEY, name TEXT) WITHOUT ROWID').rowidCol).toBe(-1);
  });
  it('ignores comments in a table definition', () => {
    const t = parseCreateTable("CREATE TABLE cm (id INTEGER PRIMARY KEY, -- the user's id, unique\n name TEXT, /* a, b */\n age INT DEFAULT (1+2))");
    expect(t.cols).toEqual(['id', 'name', 'age']);
  });
  it('writes well-formed OpenDocument when a sheet name needs escaping', async () => {
    const b = new Workbook();
    b.transact('t', () => { b.sheet.name = 'P&L'; b.sheet.set(0, 0, { v: 1 }); b.sheet.set(1, 0, { v: 2 }); });
    b.names.push({ name: 'N', ref: "'P&L'!$A$1", sheet: null } as any);
    b.sheet.filter = { range: { r1: 0, c1: 0, r2: 1, c2: 0 }, columns: {} } as any;
    const back = await importOds(await (await exportOds(b)).arrayBuffer());
    expect(back.sheets[0].name).toBe('P&L');
    expect(back.sheets[0].value(1, 0)).toBe(2);
    expect(back.names.find((n) => n.name === 'N')).toBeTruthy();
  });
  it('keeps hyperlink tooltips and everything after the first # in .xlsx', async () => {
    const b = new Workbook();
    b.transact('t', () => {
      b.sheet.set(0, 0, { v: 'a', link: '#Sheet1!B2', linkTip: 'tip' });
      b.sheet.set(1, 0, { v: 'b', link: 'https://ex.com/a#frag#x', linkTip: 'more' });
    });
    const back = await importXlsx(await (await exportXlsx(b)).arrayBuffer());
    expect(back.sheet.get(0, 0)).toMatchObject({ link: '#Sheet1!B2', linkTip: 'tip' });
    expect(back.sheet.get(1, 0)).toMatchObject({ link: 'https://ex.com/a#frag#x', linkTip: 'more' });
  });
});

describe('undo after a sheet has been re-created', () => {
  it('still undoes and redoes cell edits on that sheet', () => {
    const b = new Workbook();
    b.transact('add', () => { b.recordBook(); b.sheets.push(new Sheet(b, 'Two')); });
    b.transact('type', () => b.sheets[1].set(0, 0, { v: 'x' }));
    b.undo(); b.undo();
    expect(b.sheets.length).toBe(1);
    b.redo(); b.redo();
    expect(b.sheets[1].value(0, 0)).toBe('x');
    b.undo();
    expect(b.sheets[1].value(0, 0) ?? null).toBeNull();
  });
  it('redoes a step that wrote cells before and after a structural change', () => {
    const b = new Workbook(); const s = b.sheet;
    b.transact('mix', () => { s.set(0, 0, { v: 1 }); shiftCells(s, 'row', 0, 1); s.set(0, 0, { v: 2 }); });
    expect([s.value(0, 0), s.value(1, 0)]).toEqual([2, 1]);
    b.undo();
    expect([s.value(0, 0) ?? null, s.value(1, 0) ?? null]).toEqual([null, null]);
    b.redo();
    expect([b.sheet.value(0, 0), b.sheet.value(1, 0)]).toEqual([2, 1]);
  });
});

describe('references when rows, columns or cells are inserted and deleted', () => {
  it('insert cells (shift down) widens a range that straddles the insertion', () => {
    const { b, e, s } = book({ A1: 1, A2: 2, A3: 3, A4: 4, A5: 5, B1: '=SUM(A1:A5)', C1: '=A4' });
    b.transact('i', () => insertCellsShift(s, { r1: 2, c1: 0, r2: 2, c2: 0 }, 'down'));
    e.recalc();
    expect(s.get(0, 1)?.f).toBe('SUM(A1:A6)');
    expect(s.value(0, 1)).toBe(15);
    expect(s.get(0, 2)?.f).toBe('A5');
  });
  it('delete cells (shift up) shrinks a straddling range and breaks a reference to the deleted cell', () => {
    const { b, e, s } = book({ A1: 1, A2: 2, A3: 3, A4: 4, A5: 5, B1: '=SUM(A1:A5)', C1: '=A3', D1: '=A5' });
    b.transact('d', () => insertCellsShift(s, { r1: 2, c1: 0, r2: 2, c2: 0 }, 'up'));
    e.recalc();
    expect(s.get(0, 1)?.f).toBe('SUM(A1:A4)');
    expect(s.value(0, 1)).toBe(12);
    expect(s.get(0, 2)?.f).toBe('#REF!');
    expect(s.get(0, 3)?.f).toBe('A4');
  });
  it('a range typed backwards is adjusted when a row inside it is deleted', () => {
    const { b, e, s } = book({ A1: 1, A2: 2, A3: 3, A4: 4, A5: 5, B1: '=SUM(A4:A2)' });
    b.transact('d', () => shiftCells(s, 'row', 2, -1));
    e.recalc();
    expect(s.get(0, 1)?.f).toBe('SUM(A3:A2)');
    expect(s.value(0, 1)).toBe(6);
  });
  it('a validation or conditional-format formula survives the deletion of its first row', () => {
    const { b, s } = book({ A1: 1 });
    s.validations.push({ ranges: [{ r1: 0, c1: 0, r2: 9, c2: 0 }], type: 'custom', formula1: 'A1>5' } as any);
    s.cf.push({ id: 'r1', type: 'expression', formulas: ['$B1+A1>5'], ranges: [{ r1: 0, c1: 0, r2: 5, c2: 0 }], priority: 1 } as any);
    b.transact('d', () => shiftCells(s, 'row', 0, -1));
    expect(s.validations[0].formula1).toBe('A1>5');
    expect(s.validations[0].ranges[0]).toMatchObject({ r1: 0, r2: 8 });
    expect(s.cf[0].formulas).toEqual(['$B1+A1>5']);
  });
  it('reference text outside formulas follows inserted rows', () => {
    const { b, s } = book({ B1: 1 });
    s.scenarios.push({ name: 'x', cells: '$B$1,$B$2', values: [1, 2] });
    s.print.titleRows = '$1:$1';
    b.transact('i', () => shiftCells(s, 'row', 0, 2));
    expect(s.scenarios[0].cells).toBe('$B$3,$B$4');
    expect(s.print.titleRows).toBe('$3:$3');
  });
  it('filter criteria stay with their columns', () => {
    const { b, s } = book({ A1: 'id', B1: 'k', C1: 'v' });
    s.filter = { range: { r1: 0, c1: 0, r2: 3, c2: 2 }, columns: { 1: { values: ['a'] } as any }, sort: { col: 2, desc: false } };
    b.transact('d', () => shiftCells(s, 'col', 0, -1));
    expect(Object.keys(s.filter!.columns)).toEqual(['0']);
    expect(s.filter!.sort).toEqual({ col: 1, desc: false });
  });
});
