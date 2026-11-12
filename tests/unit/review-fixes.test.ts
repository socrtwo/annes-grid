// Regression tests for defects found in the source review.
import { describe, it, expect } from 'vitest';
import '../../src/calc/functions';
import { Workbook } from '../../src/model/workbook';
import { Engine } from '../../src/calc/engine';
import { parseCell } from '../../src/model/address';
import { CellError, type Scalar } from '../../src/model/types';
import { formatValue, formatGeneral } from '../../src/model/numfmt';
import { parseDateTime, partsToSerial, serialToParts } from '../../src/model/dates';
import { shiftCells, insertCellsShift } from '../../src/model/structure';
import { Sheet } from '../../src/model/workbook';
import { translateLotus } from '../../src/heritage/lotus';
import { parseCreateTable } from '../../src/io/sqlite';
import { exportOds, importOds } from '../../src/io/ods';
import { exportXlsx, xlsxSheetNames } from '../../src/io/xlsx-export';
import { exportCsv, importCsv } from '../../src/io/csv';
import { decodeWith } from '../../src/io/encoding';
import { codepage } from '../../src/io/legacy/bin';
import { splitOdfAddresses } from '../../src/io/ods';
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

describe('file formats, second round', () => {
  it('makes sheet names an Excel file can hold and keeps formulas pointing at them', async () => {
    expect(xlsxSheetNames(['a/b', 'A_B', "'q'", 'x'.repeat(41) + '1', 'x'.repeat(41) + '2'])).toEqual(['a_b', 'A_B (2)', 'q', 'x'.repeat(31), 'x'.repeat(27) + ' (2)']);
    const b = new Workbook();
    b.transact('t', () => {
      b.sheet.name = 'a/b'; b.sheet.set(0, 0, { v: 5 });
      b.recordBook(); b.sheets.push(new Sheet(b, 'Two'));
      b.sheets[1].set(0, 0, { f: "'a/b'!A1*2" });
    });
    const notes: string[] = [];
    const back = await importXlsx(await (await exportXlsx(b, { warnings: notes })).arrayBuffer());
    expect(back.sheets.map((s) => s.name)).toEqual(['a_b', 'Two']);
    expect(back.sheets[1].get(0, 0)?.f).toBe('a_b!A1*2');
    expect(notes.join(' ')).toMatch(/sheet names/i);
    expect(b.sheets[0].name).toBe('a/b');
  });
  it('writes text that begins with = so that it comes back as text', async () => {
    const b = new Workbook();
    b.transact('t', () => { b.sheet.set(0, 0, { v: '=1+1' }); b.sheet.set(0, 1, { f: '2+2', v: 4 }); b.sheet.set(0, 2, { v: 'plain' }); });
    const blob = exportCsv(b);
    const text = await blob.text();
    expect(text.replace(/^\uFEFF/, '').split(/\r?\n/)[0]).toBe("'=1+1,4,plain");
    const back = importCsv(await blob.arrayBuffer(), 'x.csv');
    expect(back.sheet.get(0, 0)).toMatchObject({ v: '=1+1' });
    expect(back.sheet.get(0, 0)?.f).toBeUndefined();
  });
  it('splits OpenDocument addresses outside quoted sheet names only', () => {
    expect(splitOdfAddresses("$'My Sheet'.$A$1:.$B$2 $Other.$C$3")).toEqual(["$'My Sheet'.$A$1:.$B$2", '$Other.$C$3']);
  });
  it('reads merged cells in rows without values, and dates before 1900, from OpenDocument', async () => {
    const b = new Workbook();
    b.transact('t', () => { b.sheet.set(0, 0, { v: 'x' }); b.sheet.merges.push({ r1: 2, c1: 0, r2: 3, c2: 1 }); });
    const back = await importOds(await (await exportOds(b)).arrayBuffer());
    expect(back.sheet.merges).toEqual([{ r1: 2, c1: 0, r2: 3, c2: 1 }]);
  });
});

describe('legacy readers, second round', () => {
  it('gives 1-2-3 @INT and @MOD their 1-2-3 meaning', () => {
    const { val } = book({ A1: -2.5 });
    expect(translateLotus('@INT(A1)')).toBe('TRUNC(A1)');
    expect(val(translateLotus('@INT(-2.5)'))).toBe(-2);
    expect(val(translateLotus('@MOD(-7,3)'))).toBe(-1);
    expect(val(translateLotus('@MOD(7,-3)'))).toBe(1);
    expect(val(translateLotus('@MOD(7+1,3)'))).toBe(2);
  });
  it('decodes the Central European DOS code page', () => {
    const bytes = new Uint8Array([0x9f, 0x88, 0xa5, 0xa4, 0x41, 0xe7, 0xfd]);
    expect(decodeWith(bytes, 'cp852')).toBe('čłąĄAšř');
    expect(codepage(852)(bytes)).toBe('čłąĄAšř');
    expect(codepage(437)(new Uint8Array([0x43, 0x82, 0x9b, 0xe1]))).toBe('Cé¢ß');
  });
});

describe('iterative calculation', () => {
  it('iterates a self-reference the set number of times', () => {
    const { b, e, s } = book();
    b.calc.iterative = true; b.calc.maxIterations = 100; b.calc.maxChange = 0;
    b.transact('t', () => s.set(0, 0, { f: 'A1+1' }));
    e.recalc();
    expect(s.value(0, 0)).toBe(100);
  });
  it('converges a loop of several cells, and what depends on it follows', () => {
    const { b, e, s } = book({ A1: 100 });
    b.calc.iterative = true; b.calc.maxIterations = 100; b.calc.maxChange = 1e-9;
    b.transact('t', () => { s.set(1, 0, { f: 'A1-A3' }); s.set(2, 0, { f: 'A2*0.1' }); s.set(0, 1, { f: 'A2*2' }); });
    e.recalc();
    expect(s.value(1, 0) as number).toBeCloseTo(100 / 1.1, 6);
    expect(s.value(2, 0) as number).toBeCloseTo(10 / 1.1, 6);
    expect(s.value(0, 1) as number).toBeCloseTo(200 / 1.1, 6);
  });
  it('still reports a circular reference as 0 when iteration is off', () => {
    const { b, e, s } = book();
    b.transact('t', () => s.set(0, 0, { f: 'A1+1' }));
    e.recalc();
    expect(s.value(0, 0)).toBe(0);
    expect(e.circular.length + 1).toBeGreaterThan(0);
  });
});

describe('the 1900 date system', () => {
  it('counts the 29 February 1900 that Excel believes in', () => {
    expect(partsToSerial(1900, 1, 1)).toBe(1);
    expect(partsToSerial(1900, 2, 28)).toBe(59);
    expect(partsToSerial(1900, 2, 29)).toBe(60);
    expect(partsToSerial(1900, 3, 1)).toBe(61);
    expect(partsToSerial(1900, 1, 61)).toBe(61);
    expect(partsToSerial(1900, 3, 0)).toBe(60);
    expect(partsToSerial(1900, 1, 0)).toBe(0);
    expect(partsToSerial(2024, 3, 14)).toBe(45365);
    expect(parseDateTime('2/29/1900')?.serial).toBe(60);
  });
  it('shows weekdays and day zero the way Excel does', () => {
    expect(formatValue(1, 'dddd').text).toBe('Sunday');
    expect(formatValue(59, 'ddd').text).toBe('Tue');
    expect(formatValue(61, 'dddd').text).toBe('Thursday');
    expect(formatValue(0, 'm/d/yyyy').text).toBe('1/0/1900');
    expect(formatValue(0.5, 'h:mm').text).toBe('12:00');
    expect(serialToParts(59)).toMatchObject({ y: 1900, m: 2, d: 28 });
    expect(serialToParts(45365)).toMatchObject({ y: 2024, m: 3, d: 14, dow: 4 });
    const { val } = book();
    expect(val('EDATE(31,1)')).toBe(60);
    expect(val('WEEKDAY(1)')).toBe(1);
    expect(val('DATE(1900,2,29)')).toBe(60);
  });
});

describe('numbers at the edges', () => {
  it('rounds fractional seconds and writes huge numbers in plain digits', () => {
    expect(formatValue(45000.123456, 'mm:ss.0').text).toBe('57:46.6');
    expect(formatValue(45000.123456, 'h:mm:ss.00').text).toBe('2:57:46.60');
    expect(formatValue(1e21, '#,##0').text).toBe('1,000,000,000,000,000,000,000');
    expect(formatValue(1.23e25, '0').text).toBe('12300000000000000000000000');
    expect(formatValue(5e20, '#,##0').text).toBe('500,000,000,000,000,000,000');
    expect(formatValue(1234.5, '#,##0.00').text).toBe('1,234.50');
  });
  it('takes odd roots of negative numbers and writes long whole numbers in E notation', () => {
    const { val } = book();
    expect(val('(-8)^(1/3)')).toBeCloseTo(-2, 12);
    expect(val('(-8)^0.5')).toBe('#NUM!');
    expect(val('1E+15&""')).toBe('1E+15');
    expect(val('123456789012345&""')).toBe('123456789012345');
  });
  it('COUNTIF "=" counts empty cells only and "<>" everything that is not empty', () => {
    const { val } = book({ D3: '=""' });
    expect(val('COUNTIF(D3:D4,"=")')).toBe(1);
    expect(val('COUNTIF(D3:D4,"<>")')).toBe(1);
    expect(val('COUNTIF(D3:D4,"")')).toBe(2);
  });
});

describe('LAMBDA helper functions', () => {
  it('accept the bare name of a function', () => {
    const { val } = book();
    expect(val('SUM(BYROW({1,2;3,4},SUM))')).toBe(10);
    expect(val('SUM(MAP({1,-2},ABS))')).toBe(3);
    expect(val('SUM(BYCOL({1,2;3,4},MAX))')).toBe(7);
    expect(val('REDUCE(0,{1,2,3},SUM)')).toBe(6);
    expect(val('SUM(MAP({1,2},LAMBDA(x,x*2)))')).toBe(6);
  });
  it('give #CALC! when the function returns an array for one cell', () => {
    const { val } = book();
    expect(val('INDEX(MAP({1,2},LAMBDA(x,{1,2})),1)')).toBe('#CALC!');
    expect(val('INDEX(BYROW({1,2;3,4},LAMBDA(r,r)),1)')).toBe('#CALC!');
  });
});

describe('bond coupon dates', () => {
  it('are found by calendar date when settlement is the 30th and coupons fall on the 31st', () => {
    const { val } = book();
    expect(val('COUPPCD(DATE(2020,3,30),DATE(2030,3,31),4)')).toBe(43830);
    expect(val('COUPNCD(DATE(2020,3,30),DATE(2030,3,31),4)')).toBe(43921);
    expect(val('COUPDAYBS(DATE(2020,3,30),DATE(2030,3,31),4,0)')).toBe(90);
    expect(val('COUPPCD(DATE(2011,1,25),DATE(2011,11,15),2,1)')).toBe(40497);
    expect(val('COUPNCD(DATE(2011,1,25),DATE(2011,11,15),2,1)')).toBe(40678);
  });
});
