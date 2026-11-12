import { describe, it, expect } from 'vitest';
import '../../src/calc/functions';
import '../../src/commands';
import { Workbook, Sheet } from '../../src/model/workbook';
import { Engine } from '../../src/calc/engine';
import { parseCell, parseRange, cellName } from '../../src/model/address';
import { partsToSerial, serialToParts } from '../../src/model/dates';
import { defaultPrefs } from '../../src/prefs';
import { getCommand, allCommands } from '../../src/commands/registry';
import { DEFAULT_TABS } from '../../src/ui/ribbonLayout';
import { cellMatches, findMatching, resetLastCell, uniqueSheetName, menuOf } from '../../src/commands/utilities';
import { replaceText, sortOrderBy, reshape } from '../../src/commands/utilities2';
import { applyNumOp, convertUnit, toBase, fromBase, insertAt, removeByPosition, nthWord, between, swapAround, keepOnly, wordCount, changeRefStyle, expandNames, applyNamesText } from '../../src/commands/utilities3';
import { decodeEntities, encodeEntities, stripTags, urlDecode, urlParts, htmlTable } from '../../src/commands/utilities4';
import { prettyFormula, toHalfWidth, toFullWidth, chemicalRuns, missingNumbers, combinations, mergeTables, splitName, lookupFormula, rowPasses, encryptText, decryptText, runSql } from '../../src/commands/plus';
import { soundex, metaphoneKey, parseGedcomDate, gedcomRange, ymdCell, ymdText, ageBetween, julianToGregorian, gregorianToJulian, quakerMonth, surnameFirst, titleCase, smartPunctuation, straightPunctuation, fractionGlyphs, readingEase, isbnInfo, luhnValid, slugify, amortization, agingBucket, benford, fiscalYear, fingerprint, clusterValues, iqrOutliers, crc32, base64Encode, base64Decode, easter, usHolidays, weekNumberIso, dmsToDecimal, decimalToDms, haversineKm, bearingDeg, compassPoint } from '../../src/commands/extras';
import { naturalCompare, mergeSpans, selCells, externalLinkIn } from '../../src/commands/utilShared';

/** App stand-in with enough of the grid and shell for the plug-in tab commands to run headless. */
function fakeApp() {
  const book = new Workbook();
  const engine = new Engine(book);
  book.onChange = () => engine.recalc();
  const noop = () => undefined;
  const app: any = {
    book, engine, prefs: defaultPrefs(), recording: null, docs: [], clipboard: [],
    get sheet() { return book.sheet; },
    get tab() { return app._tab; },
    grid: {
      select: (r: number, c: number, o: any = {}) => { book.sheet.sel = { r: o.keepActive ? book.sheet.sel.r : r, c: o.keepActive ? book.sheet.sel.c : c, ranges: [o.range || { r1: r, c1: c, r2: r, c2: c }] }; },
      selectRows: (a: number, b: number) => { book.sheet.sel = { r: a, c: 0, ranges: [{ r1: a, c1: 0, r2: b, c2: 16383 }] }; },
      selectCols: (a: number, b: number) => { book.sheet.sel = { r: 0, c: a, ranges: [{ r1: 0, c1: a, r2: 1048575, c2: b }] }; },
      objects: { select: noop, selectedObject: () => null }, invalidate: noop, layout: noop, updateExtent: noop, scrollIntoView: noop, setZoom: (z: number) => { book.sheet.zoom = z; },
      visible: () => ({ r1: 0, r2: 40, c1: 0, c2: 12 }), scroller: { scrollTop: 0, scrollLeft: 0 },
    },
    isReadOnly: () => false,
    exec: (id: string, arg?: any) => { if (id === 'sheet.activate') { book.activeSheet = arg; return true; } const c = getCommand(id); return c ? c.run(app, arg) : false; },
    refresh() {}, savePrefs() {}, calculate() { engine.recalc(); }, followLink() {}, activate() {}, onRefresh() { return () => undefined; },
    createTab(b: Workbook) { return { book: b, dirty: false }; },
    ui: { status: noop, openPane: noop, togglePane: noop, paneOpen: () => false, rebuild: noop, applyPrefs: noop, anchorFor: () => document.body, commandPalette: noop, hideBackstage: noop, showBackstage: noop },
  };
  app._tab = { book, engine, displayName: 'Test', title: 'Test', fileName: '', id: 't1', cf: { evaluate: () => null, invalidate: noop }, viewMode: 'normal' };
  const put = (a: string, v: any, sh: Sheet = book.sheet) => { const p = parseCell(a)!; book.transact('t', () => sh.set(p.r, p.c, typeof v === 'string' && v.startsWith('=') ? { f: v.slice(1) } : { v })); };
  const get = (a: string, sh: Sheet = book.sheet) => { engine.recalc(); const p = parseCell(a)!; return sh.value(p.r, p.c); };
  const cell = (a: string, sh: Sheet = book.sheet) => { const p = parseCell(a)!; return sh.get(p.r, p.c); };
  const select = (ref: string) => { const rg = parseRange(ref)!; book.sheet.sel = { r: rg.r1, c: rg.c1, ranges: [rg] }; };
  return { app, book, engine, put, get, cell, select };
}

describe('plug-in tabs are registered', () => {
  it('has Utilities, Plus and Extras tabs whose items all exist', () => {
    const tabs = DEFAULT_TABS.filter((t) => ['utilities', 'plus', 'extras'].includes(t.id));
    expect(tabs.map((t) => t.label)).toEqual(['Utilities', 'Plus', 'Extras']);
    const missing: string[] = [];
    for (const t of tabs) for (const g of t.groups) for (const id of g.items) if (!getCommand(id)) missing.push(id);
    expect(missing).toEqual([]);
  });
  it('registers several hundred new commands with unique ids, labels and icons', () => {
    const list = allCommands().filter((c) => c.tab === 'Utilities' || c.tab === 'Plus' || c.tab === 'Extras');
    expect(list.length).toBeGreaterThan(400);
    const ids = new Set(list.map((c) => c.id));
    expect(ids.size).toBe(list.length);
    expect(list.filter((c) => !c.label || !c.icon || !c.group)).toEqual([]);
  });
  it('category menus only list real commands', () => {
    const entries = menuOf(['util.selectByContent', 'nope.missing', null, 'tools.trimSpaces']);
    expect(entries.map((e) => e.id || 'sep')).toEqual(['util.selectByContent', 'sep', 'tools.trimSpaces']);
    for (const c of allCommands().filter((x) => x.kind === 'menu' && /^(util|plus|extra)\./.test(x.id))) {
      const { app } = fakeApp();
      for (const e of c.menu!(app)) if (e.id) expect(getCommand(e.id), `${c.id} → ${e.id}`).toBeTruthy();
    }
  });
});

describe('Utilities: select, sheets, range', () => {
  it('matches cells by content and format', () => {
    const { app, put, select } = fakeApp();
    put('A1', 'apple'); put('A2', 42); put('A3', 'Banana'); put('A4', -3); put('A5', '=1/0');
    select('A1:A5');
    const s = app.sheet;
    expect(cellMatches(s, 0, 0, s.get(0, 0), { content: 'contains', value: 'APP' })).toBe(true);
    expect(cellMatches(s, 0, 0, s.get(0, 0), { content: 'contains', value: 'APP', matchCase: true })).toBe(false);
    expect(findMatching(app, { content: 'number' }).map((p) => cellName(p.r, p.c))).toEqual(['A2', 'A4']);
    expect(findMatching(app, { content: 'negative' }).map((p) => cellName(p.r, p.c))).toEqual(['A4']);
    expect(findMatching(app, { content: 'error' }).map((p) => cellName(p.r, p.c))).toEqual(['A5']);
    expect(findMatching(app, { content: 'greater', value: '10' }).map((p) => cellName(p.r, p.c))).toEqual(['A2']);
  });
  it('select commands change the selection', () => {
    const { app, put, select } = fakeApp();
    put('A1', 1); put('A2', 'x'); put('A3', 3); put('B1', '=A1+1');
    select('A1:B3');
    app.exec('util.selectNumbers');
    expect(app.sheet.sel.ranges.map((r: any) => cellName(r.r1, r.c1))).toEqual(['A1', 'A3']);
    app.exec('util.reselectPrevious');
    expect(app.sheet.sel.ranges).toEqual([{ r1: 0, c1: 0, r2: 2, c2: 1 }]);
    select('A1');
    app.exec('util.extendToLastRow');
    expect(app.sheet.sel.ranges[0]).toEqual({ r1: 0, c1: 0, r2: 2, c2: 0 });
    app.exec('util.selectSameValue');
    expect(app.sheet.sel.ranges.length).toBe(1);
  });
  it('conditional row delete removes matching rows', async () => {
    const { app, put, select, get } = fakeApp();
    put('A1', 'Name'); put('B1', 'Status'); put('A2', 'a'); put('B2', 'keep'); put('A3', 'b'); put('B3', 'drop'); put('A4', 'c'); put('B4', 'keep');
    select('A1:B4');
    await app.exec('util.condRows', { key: '1', content: 'equals', value: 'drop', action: 'delete', skipHeader: true });
    expect(get('A3')).toBe('c');
    expect(get('A4')).toBeNull();
  });
  it('resets the last cell and reports sheets', () => {
    const { app, book, put } = fakeApp();
    put('A1', 'x');
    book.transact('fmt', () => book.sheet.set(50, 20, { s: book.styleId({ bold: true }) }));
    expect(book.sheet.usedRange()!.r2).toBe(50);
    book.transact('reset', () => resetLastCell(book.sheet));
    expect(book.sheet.usedRange()!.r2).toBe(0);
    app.exec('util.sheetsReport');
    expect(book.sheets.length).toBe(2);
    expect(book.sheets[1].value(1, 0)).toBe('Sheet1');
  });
  it('inserts, sorts, renames and indexes sheets', async () => {
    const { app, book } = fakeApp();
    expect(uniqueSheetName(book, 'Sheet1')).toBe('Sheet1 (2)');
    await app.exec('util.insertSheets', { names: ['Zeta', 'alpha', 'Sheet10', 'Sheet2'] });
    expect(book.sheets.map((s) => s.name)).toEqual(['Sheet1', 'Zeta', 'alpha', 'Sheet10', 'Sheet2']);
    await app.exec('util.sortSheets', { by: 'natural' });
    expect(book.sheets.map((s) => s.name)).toEqual(['alpha', 'Sheet1', 'Sheet2', 'Sheet10', 'Zeta']);
    await app.exec('util.renameSheetsReplace', { find: 'Sheet', replace: 'Tab' });
    expect(book.sheets.map((s) => s.name)).toEqual(['alpha', 'Tab1', 'Tab2', 'Tab10', 'Zeta']);
    app.exec('util.sheetIndex');
    expect(book.sheets[0].name).toBe('Index');
    expect(book.sheets[0].get(1, 0)?.link).toBe("#'alpha'!A1");
    await app.exec('util.deleteEmptySheets', { yes: true });
    expect(book.sheets.map((s) => s.name)).toEqual(['Index']);
  });
  it('advanced find and replace works on values, formulas and comments', async () => {
    const { app, put, get, book } = fakeApp();
    put('A1', 'red apple'); put('A2', '=UPPER("red")'); book.sheet.comments.set(0, { author: 'a', text: 'red note', time: 0 });
    app.sheet.sel = { r: 0, c: 0, ranges: [{ r1: 0, c1: 0, r2: 0, c2: 0 }] };
    expect(replaceText('Red Apple', { find: 'red', replace: 'green' })).toBe('green Apple');
    expect(replaceText('Red Apple', { find: 'red', replace: 'green', matchCase: true })).toBeNull();
    expect(replaceText('a1b22', { find: '(\\d+)', replace: '[$1]', regex: true })).toBe('a[1]b[22]');
    expect(replaceText('red', { find: 'red', replace: 'x', wholeCell: true })).toBe('x');
    expect(replaceText('red apple', { find: 'red', replace: 'x', wholeCell: true })).toBeNull();
    await app.exec('util.findReplaceAdvanced', { find: 'red', replace: 'green', scope: 'workbook', inFormulas: true, inComments: true });
    expect(get('A1')).toBe('green apple');
    expect(get('A2')).toBe('GREEN');
    expect(book.sheet.comments.get(0)!.text).toBe('green note');
  });
  it('sorts by natural order, length and last word; swaps ranges; colors duplicates', async () => {
    expect(sortOrderBy(['file10', 'file2', 'file1'], 'natural')).toEqual([2, 1, 0]);
    expect(sortOrderBy(['ccc', 'a', 'bb'], 'length')).toEqual([1, 2, 0]);
    expect(sortOrderBy(['John Smith', 'Ann Adams'], 'lastWord')).toEqual([1, 0]);
    expect(naturalCompare('Sheet2', 'Sheet10')).toBeLessThan(0);
    const { app, put, get, select } = fakeApp();
    put('A1', 1); put('A2', 2); put('C1', 'x'); put('C2', 'y');
    select('A1:A2');
    await app.exec('util.swapRanges', { b: 'C1:C2' });
    expect(get('A1')).toBe('x'); expect(get('C2')).toBe(2);
    put('D1', 'dup'); put('D2', 'dup'); put('D3', 'one');
    select('D1:D3');
    app.exec('util.colorDuplicates');
    expect(app.sheet.styleOf(0, 3).fill).toBeTruthy();
    expect(app.sheet.styleOf(2, 3).fill).toBeUndefined();
    expect(mergeSpans([[5, 6], [1, 2], [3, 3]])).toEqual([[1, 3], [5, 6]]);
  });
});

describe('Utilities: fill, columns & rows, comments', () => {
  it('fills with values, formulas, lists and row numbers', async () => {
    const { app, put, get, cell, select } = fakeApp();
    put('A1', 10); put('A2', 20); put('A3', 30);
    select('B1:B3');
    await app.exec('util.fillWith', { text: '=A1*2' });
    expect(cell('B3')!.f).toBe('A3*2');
    expect(get('B3')).toBe(60);
    select('C1:C4');
    await app.exec('util.fillCycle', { items: ['x', 'y'], dir: 'down' });
    expect([get('C1'), get('C2'), get('C3'), get('C4')]).toEqual(['x', 'y', 'x', 'y']);
    select('D1:D3');
    app.exec('util.fillRowNumbers');
    expect(get('D3')).toBe(3);
    select('E1');
    put('E1', '=A1+1');
    app.exec('util.fillDownToLast');
    expect(cell('E3')!.f).toBe('A3+1');
  });
  it('inserts rows every nth row and at value changes, deletes empty rows and columns', async () => {
    const { app, put, get, select } = fakeApp();
    ['a', 'a', 'b', 'b', 'c'].forEach((v, i) => put(`A${i + 1}`, v));
    select('A1:A5');
    await app.exec('util.insertRowAtChange', { col: 0, what: 'row', count: 1 });
    expect([get('A1'), get('A2'), get('A3'), get('A4'), get('A5'), get('A6'), get('A7')]).toEqual(['a', 'a', null, 'b', 'b', null, 'c']);
    select('A1');
    await app.exec('util.deleteEmptyRows', { scope: 'sheet' });
    expect([get('A1'), get('A2'), get('A3'), get('A4'), get('A5')]).toEqual(['a', 'a', 'b', 'b', 'c']);
    select('A1:A4');
    await app.exec('util.insertRowsEvery', { every: 2, count: 1 });
    expect([get('A1'), get('A2'), get('A3'), get('A4')]).toEqual(['a', 'a', null, 'b']);
    put('D1', 'far');
    select('A1');
    app.exec('util.deleteEmptyCols');
    expect(get('B1')).toBe('far');
  });
  it('reshapes columns, reorders columns and moves rows', async () => {
    expect(reshape([1, 2, 3, 4, 5], 2)).toEqual([[1, 2], [3, 4], [5, null]]);
    const { app, put, get, select } = fakeApp();
    put('A1', 'b'); put('B1', 'a'); put('C1', 'c'); put('A2', 2); put('B2', 1); put('C2', 3);
    select('A1:C2');
    app.exec('util.sortColumnsByHeader');
    expect([get('A1'), get('B1'), get('C1')]).toEqual(['a', 'b', 'c']);
    expect([get('A2'), get('B2'), get('C2')]).toEqual([1, 2, 3]);
    await app.exec('util.reorderColumns', { order: ['c', 'a'] });
    expect([get('A1'), get('B1'), get('C1')]).toEqual(['c', 'a', 'b']);
    select('A2:C2');
    app.exec('util.moveRowUp');
    expect([get('A1'), get('A2')]).toEqual([3, 'c']);
  });
  it('combines rows by key and duplicates rows', async () => {
    const { app, put, get, select } = fakeApp();
    put('A1', 'k'); put('B1', 'n'); put('A2', 'x'); put('B2', 1); put('A3', 'y'); put('B3', 5); put('A4', 'x'); put('B4', 2);
    select('A1:B4');
    await app.exec('util.joinRowsByKey', { key: 0, op: 'sum' });
    expect([get('A2'), get('B2'), get('A3'), get('B3'), get('A4')]).toEqual(['x', 3, 'y', 5, null]);
    select('A2:B3');
    await app.exec('util.duplicateRows', { times: 2 });
    expect([get('A2'), get('A3'), get('A4'), get('A5')]).toEqual(['x', 'x', 'y', 'y']);
  });
  it('copies comments to cells and lists them', () => {
    const { app, put, get, book, select } = fakeApp();
    put('A1', 'v'); book.sheet.comments.set(0, { author: 'me', text: 'hello', time: 0, note: true });
    select('A1');
    app.exec('util.commentsToCells', 'right');
    expect(get('B1')).toBe('hello');
    app.exec('util.commentsReport');
    expect(book.sheets[1].value(1, 5)).toBe('hello');
  });
});

describe('Utilities: format, numbers, text, formulas', () => {
  it('number operations and conversions', () => {
    expect(applyNumOp(10, 'multiply', 3)).toBe(30);
    expect(applyNumOp(-2.345, 'round', 2)).toBe(-2.35);
    expect(applyNumOp(7, 'roundTo', 5)).toBe(5);
    expect(applyNumOp(100, 'addPercent', 10)).toBe(110);
    expect(convertUnit(1, 'in', 'cm')).toBeCloseTo(2.54);
    expect(convertUnit(100, '°C', '°F')).toBeCloseTo(212);
    expect(convertUnit(1, 'kg', 'm')).toBeNull();
    expect(toBase(255, 16)).toBe('FF');
    expect(fromBase('ff', 16)).toBe(255);
    expect(fromBase('101', 2)).toBe(5);
  });
  it('applies calculations, pads zeros, shifts dates in place', async () => {
    const { app, put, get, cell, select, book } = fakeApp();
    put('A1', 5); put('A2', '=A1*2'); put('A3', 7);
    select('A1:A3');
    await app.exec('util.numberOperation', { op: 'multiply', x: 10, formulas: true });
    expect(get('A1')).toBe(50); expect(cell('A2')!.f).toBe('(A1*2)*10'); expect(get('A3')).toBe(70);
    app.exec('util.changeSign');
    expect(get('A3')).toBe(-70);
    put('B1', 42); select('B1');
    await app.exec('util.padZeros', { width: 5, how: 'text' });
    expect(get('B1')).toBe('00042');
    const d = partsToSerial(2024, 1, 31);
    book.transact('d', () => book.sheet.set(0, 2, { v: d, s: book.styleId({ numFmt: 'yyyy-mm-dd' }) }));
    select('C1');
    await app.exec('util.shiftDates', { n: 1, unit: 'months' });
    const p = serialToParts(get('C1') as number);
    expect([p.y, p.m, p.d]).toEqual([2024, 2, 29]);
    await app.exec('util.shiftDates', { n: 0, unit: 'startQuarter' });
    expect(serialToParts(get('C1') as number).d).toBe(1);
  });
  it('text tools', () => {
    expect(insertAt('abc', '-', 'position', 2)).toBe('a-bc');
    expect(insertAt('camelCase', ' ', 'beforeUpper')).toBe('camel Case');
    expect(insertAt('abcdef', '-', 'everyN', 2)).toBe('ab-cd-ef');
    expect(removeByPosition('hello', 'first', 2)).toBe('llo');
    expect(removeByPosition('hello', 'last', 2)).toBe('hel');
    expect(removeByPosition('hello', 'from', 2, 2)).toBe('hlo');
    expect(removeByPosition('a-b-c', 'before', 0, 0, '-')).toBe('b-c');
    expect(removeByPosition('a-b-c', 'after', 0, 0, '-')).toBe('a');
    expect(nthWord('one two three', 2)).toBe('two');
    expect(nthWord('one two three', -1)).toBe('three');
    expect(between('x [a] y', '[', ']')).toBe('a');
    expect(swapAround('Smith, John', ',', ' ')).toBe('John Smith');
    expect(keepOnly('ab-12.5c', 'digits')).toBe('-12.5');
    expect(keepOnly('a1 b2', 'letters')).toBe('a b');
    expect(wordCount('  a b  c ')).toBe(3);
  });
  it('formula tools', async () => {
    expect(changeRefStyle('A1+Sheet2!B2:C3', 'absolute')).toBe('$A$1+Sheet2!$B$2:$C$3');
    expect(changeRefStyle('$A$1', 'relative')).toBe('A1');
    expect(changeRefStyle('A1', 'rowAbs')).toBe('A$1');
    expect(expandNames('SUM(Sales)*2', [{ name: 'Sales', ref: 'Sheet1!$B$2:$B$9' }])).toBe('SUM((Sheet1!$B$2:$B$9))*2');
    expect(applyNamesText('SUM($B$2:$B$9)', [{ name: 'Sales', ref: 'Sheet1!$B$2:$B$9' }], 'Sheet1')).toBe('SUM(Sales)');
    const { app, put, get, cell, select } = fakeApp();
    put('A1', 1); put('A2', '=A1/0');
    select('A2');
    await app.exec('util.refStyle', 'absolute');
    expect(cell('A2')!.f).toBe('$A$1/0');
    await app.exec('util.wrapFormulas', { wrap: 'iferror', x: '"n/a"' });
    expect(cell('A2')!.f).toBe('IFERROR($A$1/0,"n/a")');
    expect(get('A2')).toBe('n/a');
    app.exec('util.formulasToText');
    expect(get('A2')).toBe('=IFERROR($A$1/0,"n/a")');
    app.exec('util.textToFormulas');
    expect(cell('A2')!.f).toBe('IFERROR($A$1/0,"n/a")');
    app.exec('util.copyExact');
    select('C5');
    app.exec('util.pasteExact');
    expect(cell('C5')!.f).toBe('IFERROR($A$1/0,"n/a")');
  });
  it('format by content, banding and freeze conditional formatting', async () => {
    const { app, put, select, book } = fakeApp();
    put('A1', 5); put('A2', 50); put('A3', 'x');
    select('A1:A3');
    await app.exec('util.formatByContent', { content: 'greater', value: '10', preset: 'bold', scope: 'selection' });
    expect(book.sheet.styleOf(1, 0).bold).toBe(true);
    expect(book.sheet.styleOf(0, 0).bold).toBeUndefined();
    await app.exec('util.bandRows', { axis: 'row', every: 2, color: '#eeeeee', start: 1 });
    expect(book.sheet.styleOf(0, 0).fill).toBe('#eeeeee');
    expect(book.sheet.styleOf(1, 0).fill).toBeUndefined();
    await app.exec('util.removeFormattingParts', { parts: ['fills'], scope: 'sheet' });
    expect(book.sheet.styleOf(0, 0).fill).toBeUndefined();
    expect(book.sheet.styleOf(1, 0).bold).toBe(true);
  });
});

describe('Utilities: web, info, export', () => {
  it('html and url helpers', () => {
    expect(decodeEntities('a &amp; b &eacute; &#8217; &#x41;')).toBe('a & b é ’ A');
    expect(encodeEntities('<a>é')).toBe('&lt;a&gt;&#233;');
    expect(stripTags('<p>Hi <b>there</b><br>you &amp; me</p>')).toBe('Hi there\nyou & me');
    expect(urlDecode('a%20b+c')).toBe('a b c');
    expect(urlParts('https://example.com/p/q?x=1#frag')).toEqual({ protocol: 'https', host: 'example.com', path: '/p/q', query: 'x=1', hash: 'frag' });
    const { app, put } = fakeApp();
    put('A1', 'h'); put('A2', 'v<');
    expect(htmlTable(app.sheet, { r1: 0, c1: 0, r2: 1, c2: 0 }, { header: true })).toBe('<table>\n<tr><th>h</th></tr>\n<tr><td>v&lt;</td></tr>\n</table>');
  });
  it('activates links, lists values with counts and finds problems', () => {
    const { app, put, get, cell, select, book } = fakeApp();
    put('A1', 'www.example.com'); put('A2', 'bob@example.com'); put('A3', 'plain');
    select('A1:A3');
    app.exec('util.activateLinks');
    expect(cell('A1')!.link).toBe('http://www.example.com');
    expect(cell('A2')!.link).toBe('mailto:bob@example.com');
    expect(cell('A3')!.link).toBeUndefined();
    put('B1', 'a'); put('B2', 'a'); put('B3', 'b');
    select('B1:B3');
    app.exec('util.valueCounts');
    expect(book.sheets[1].value(1, 0)).toBe('a');
    expect(book.sheets[1].value(1, 1)).toBe(2);
    book.activeSheet = 0;
    put('C1', '=Missing!A1');
    app.exec('util.findProblems');
    const rep = book.sheetByName('Problems')!;
    expect(String(rep.value(1, 2))).toMatch(/Broken reference|Link to a missing sheet|Error value/);
    void get;
  });
  it('tracks the last tool and usage', () => {
    const { app, put, select, book } = fakeApp();
    put('A1', ' x '); select('A1');
    app.exec('util.removeAllSpaces');
    expect(book.sheet.value(0, 0)).toBe('x');
    put('A1', ' y '); select('A1');
    app.exec('util.lastTool');
    expect(book.sheet.value(0, 0)).toBe('y');
    app.exec('util.usageReport');
    expect(book.sheets[1].value(1, 1)).toBeGreaterThanOrEqual(2);
  });
});

describe('Plus tab', () => {
  it('pretty-prints formulas and converts widths', () => {
    expect(prettyFormula('=IF(A1,SUM(B1:B2),0)')).toBe('=IF(\n  A1,\n  SUM(\n    B1:B2\n  ),\n  0\n)');
    expect(toHalfWidth('Ｅｘｃｅｌ　１')).toBe('Excel 1');
    expect(toFullWidth('A1')).toBe('Ａ１');
    expect(chemicalRuns('H2O')).toEqual([{ t: 'H' }, { t: '2', s: { vertAlign: 'subscript' } }, { t: 'O' }]);
    expect(chemicalRuns('hello')).toBeNull();
    expect(missingNumbers([1, 2, 5, 'x', 7])).toEqual([3, 4, 6]);
    expect(combinations([['S', 'M'], ['red', 'blue']])).toEqual([['S', 'red'], ['S', 'blue'], ['M', 'red'], ['M', 'blue']]);
  });
  it('splits names', () => {
    expect(splitName('Dr. John A. Smith Jr.')).toEqual({ title: 'Dr.', first: 'John', middle: 'A.', last: 'Smith', suffix: 'Jr.' });
    expect(splitName('Smith, John A.')).toEqual({ title: '', first: 'John', middle: 'A.', last: 'Smith', suffix: '' });
    expect(splitName('Ludwig van Beethoven')).toEqual({ title: '', first: 'Ludwig', middle: '', last: 'van Beethoven', suffix: '' });
    expect(splitName('Cher').first).toBe('Cher');
  });
  it('merges tables by key', () => {
    const main = [['a', 1], ['b', 2], ['c', 3]];
    const lookup = [['B', 'bee'], ['a', 'ay'], ['a', 'ay2']];
    const r = mergeTables(main, lookup, [0], [0], [1]);
    expect(r.rows).toEqual([['a', 1, 'ay'], ['b', 2, 'bee'], ['c', 3, null]]);
    expect(r.matched).toBe(2);
    expect(mergeTables(main, lookup, [0], [0], [1], { multi: 'join' }).rows[0][2]).toBe('ay, ay2');
  });
  it('lookup wizard formulas and filter conditions', () => {
    expect(lookupFormula({ kind: 'bottom', lookup: 'A2', keys: '$A$2:$A$9', returns: '$B$2:$B$9' })).toBe('XLOOKUP(A2,$A$2:$A$9,$B$2:$B$9,"",0,-1)');
    expect(lookupFormula({ kind: 'many', lookup: 'A2', keys: 'K', returns: 'V', sep: '; ' })).toBe('TEXTJOIN("; ",TRUE,FILTER(V,K=A2,""))');
    const { app, put } = fakeApp();
    put('A1', 'Apple'); put('B1', 12); put('A2', 'pear'); put('B2', 3);
    expect(rowPasses(app.sheet, 0, [{ col: 0, op: 'starts', value: 'ap' }, { col: 1, op: 'gt', value: '10' }], 'and')).toBe(true);
    expect(rowPasses(app.sheet, 1, [{ col: 0, op: 'upper', value: '' }, { col: 1, op: 'gt', value: '10' }], 'or')).toBe(false);
    expect(rowPasses(app.sheet, 1, [{ col: 0, op: 'lower', value: '' }], 'and')).toBe(true);
    expect(rowPasses(app.sheet, 0, [{ col: 1, op: 'between', value: '10', value2: '20' }], 'and')).toBe(true);
  });
  it('super filter hides rows and show-all restores them', async () => {
    const { app, put, select, book } = fakeApp();
    put('A1', 'n'); put('A2', 1); put('A3', 20); put('A4', 3);
    select('A1:A4');
    await app.exec('plus.superFilter', { conds: [{ col: 0, op: 'gt', value: '5' }], logic: 'and' });
    expect([...book.sheet.hiddenRows].sort()).toEqual([1, 3]);
    app.exec('plus.showAllRows');
    expect(book.sheet.hiddenRows.size).toBe(0);
  });
  it('encrypts and decrypts cells', async () => {
    const enc = await encryptText('secret 42', 'pw');
    expect(enc.startsWith('🔒')).toBe(true);
    expect(await decryptText(enc, 'pw')).toBe('secret 42');
    expect(await decryptText(enc, 'wrong')).toBeNull();
    const { app, put, get, select } = fakeApp();
    put('A1', 'hello'); put('A2', 7); select('A1:A2');
    await app.exec('plus.encryptCells', { password: 'x' });
    expect(String(get('A1'))).toMatch(/^🔒/);
    await app.exec('plus.decryptCells', { password: 'x' });
    expect(get('A1')).toBe('hello'); expect(get('A2')).toBe(7);
  });
  it('runs SQL over sheet records', () => {
    const t = { Sales: [{ Region: 'East', Amount: 10, Rep: 'Ann' }, { Region: 'West', Amount: 25, Rep: 'Bob' }, { Region: 'East', Amount: 5, Rep: 'Cid' }, { Region: 'North', Amount: null, Rep: 'Dee' }] };
    expect(runSql('SELECT Rep FROM Sales WHERE Amount > 7 ORDER BY Amount DESC', t).rows).toEqual([['Bob'], ['Ann']]);
    const g = runSql('SELECT Region, SUM(Amount) AS Total, COUNT(*) AS N FROM Sales GROUP BY Region ORDER BY Total DESC', t);
    expect(g.columns).toEqual(['Region', 'Total', 'N']);
    expect(g.rows).toEqual([['West', 25, 1], ['East', 15, 2], ['North', 0, 1]]);
    expect(runSql("SELECT * FROM Sales WHERE Rep LIKE 'A%' OR Region IN ('North')", t).rows.length).toBe(2);
    expect(runSql('SELECT COUNT(*) FROM Sales WHERE Amount IS NULL', t).rows).toEqual([[1]]);
    expect(runSql('SELECT DISTINCT Region FROM Sales', t).rows.length).toBe(3);
    expect(runSql("SELECT UPPER(Rep) AS R FROM [Sales] WHERE Amount BETWEEN 5 AND 10 LIMIT 1", t).rows).toEqual([['ANN']]);
    expect(() => runSql('SELECT Nope FROM Sales', t)).toThrow(/Unknown column/);
  });
  it('splits rows into sheets and combines them back', async () => {
    const { app, put, select, book } = fakeApp();
    put('A1', 'Region'); put('B1', 'Amt'); put('A2', 'E'); put('B2', 1); put('A3', 'W'); put('B3', 2); put('A4', 'E'); put('B4', 3);
    select('A1:B4');
    await app.exec('util.splitToSheets', { col: 0, header: true });
    expect(book.sheets.map((s) => s.name)).toEqual(['Sheet1', 'E', 'W']);
    expect(book.sheets[1].value(2, 1)).toBe(3);
    book.activeSheet = 0;
    await app.exec('util.combineSheets', { names: ['E', 'W'], mode: 'stack', nameColumn: true });
    const c = book.sheetByName('Combined')!;
    expect(c.value(0, 0)).toBe('Sheet');
    expect(c.value(1, 0)).toBe('E');
    expect(c.value(3, 2)).toBe(2);
  });
  it('workbook analysis and cell relationship reports run', () => {
    const { app, put, select, book } = fakeApp();
    put('A1', 1); put('A2', '=A1*2'); put('A3', '=A2+1');
    select('A2');
    app.exec('plus.cellRelationship');
    const rep = book.sheets[1];
    expect(rep.value(1, 0)).toBe('Cell');
    expect(rep.cells.size).toBeGreaterThan(6);
    book.activeSheet = 0;
    app.exec('plus.workbookAnalysis');
    expect(book.sheetByName('Workbook Analysis')!.value(0, 0)).toBe('Category');
  });
});

describe('Extras tab', () => {
  it('soundex and phonetic keys', () => {
    expect(soundex('Robert')).toBe('R163'); expect(soundex('Rupert')).toBe('R163'); expect(soundex('Tymczak')).toBe('T522'); expect(soundex('Pfister')).toBe('P236'); expect(soundex('Ashcraft')).toBe('A261'); expect(soundex('Honeyman')).toBe('H555');
    expect(metaphoneKey('Smith')).toBe(metaphoneKey('Smyth'));
    expect(metaphoneKey('Knight')).toBe('NT');
  });
  it('genealogy dates and ages', () => {
    expect(parseGedcomDate('ABT 1850')).toMatchObject({ kind: 'about', y: 1850 });
    expect(parseGedcomDate('BEF 12 MAR 1799')).toMatchObject({ kind: 'before', y: 1799, m: 3, d: 12 });
    expect(parseGedcomDate('BET 1800 AND 1810')).toMatchObject({ kind: 'between', y: 1800, y2: 1810 });
    expect(parseGedcomDate('1 FEB 1699/00')).toMatchObject({ kind: 'exact', y: 1699, dual: 1700 });
    expect(parseGedcomDate('nonsense').kind).toBe('unknown');
    const rg = gedcomRange(parseGedcomDate('BET 1800 AND 1810'))!;
    expect(rg[0]).toEqual({ y: 1800, m: 1, d: 1 }); expect(rg[1]).toEqual({ y: 1810, m: 12, d: 31 });
    expect(gedcomRange(parseGedcomDate('BEF 1 MAR 1799'))![1]).toEqual({ y: 1799, m: 2, d: 28 });
    expect(parseGedcomDate('5 March 1850')).toMatchObject({ y: 1850, m: 3, d: 5 });
    expect(ymdText(ymdCell({ y: 1850, m: 3, d: 5 }).v as any === '1850-03-05' ? { y: 1850, m: 3, d: 5 } : { y: 0, m: 1, d: 1 })).toBe('1850-03-05');
    expect(ageBetween(partsToSerial(1950, 6, 15), partsToSerial(2024, 6, 14))).toEqual({ years: 73, months: 11, days: 30 });
    expect(julianToGregorian(1752, 9, 2)).toEqual({ y: 1752, m: 9, d: 13 });
    expect(gregorianToJulian(1752, 9, 13)).toEqual({ y: 1752, m: 9, d: 2 });
    expect(quakerMonth(1, 1720)).toBe(3); expect(quakerMonth(11, 1720)).toBe(1); expect(quakerMonth(1, 1800)).toBe(1);
    expect(surnameFirst('John van der Berg')).toBe('van der Berg, John');
  });
  it('publishing helpers', () => {
    expect(titleCase('the lord of the rings: a tale of the end')).toBe('The Lord of the Rings: A Tale of the End');
    expect(titleCase('working with NASA and the fbi')).toBe('Working with NASA and the Fbi');
    expect(smartPunctuation('"Hi," she said -- it\'s 1990-1995...')).toBe('“Hi,” she said – it’s 1990–1995…');
    expect(straightPunctuation('“Hi” — ‘x’')).toBe('"Hi" -- \'x\'');
    expect(fractionGlyphs('1/2 cup and 3/4')).toBe('½ cup and ¾');
    expect(readingEase('The cat sat. The dog ran. It was fun.')).toBeGreaterThan(90);
    expect(isbnInfo('0-306-40615-2')).toMatchObject({ valid: true, isbn13: '9780306406157' });
    expect(isbnInfo('9780306406157')).toMatchObject({ valid: true, isbn10: '0306406152' });
    expect(isbnInfo('0306406153').valid).toBe(false);
    expect(luhnValid('4539 1488 0343 6467')).toBe(true); expect(luhnValid('4539 1488 0343 6468')).toBe(false);
    expect(slugify('Crème Brûlée & Co.')).toBe('creme-brulee-co');
  });
  it('finance helpers', () => {
    const rows = amortization(1000, 0.12, 12, partsToSerial(2024, 1, 15));
    expect(rows.length).toBe(12);
    expect(rows[0][2]).toBeCloseTo(88.85, 2);
    expect(rows[11][5]).toBe(0);
    expect(agingBucket(0)).toBe('Current'); expect(agingBucket(45)).toBe('31–60'); expect(agingBucket(200)).toBe('Over 90');
    const b = benford([1, 1, 1, 2, 2, 3, 9]);
    expect(b[0]).toMatchObject({ digit: 1, count: 3 });
    expect(b[0].expected).toBeCloseTo(0.301, 3);
    expect(fiscalYear(partsToSerial(2024, 8, 1), 7)).toEqual({ fy: 2025, quarter: 1 });
    expect(fiscalYear(partsToSerial(2024, 3, 1), 1)).toEqual({ fy: 2024, quarter: 1 });
  });
  it('data quality helpers', () => {
    expect(fingerprint('New York, NY')).toBe(fingerprint('ny new york'));
    const cl = clusterValues(['New York', 'new york ', 'York New', 'Boston']);
    expect(cl.length).toBe(1); expect(cl[0].variants[0].text).toBe('New York');
    expect(iqrOutliers([1, 2, 3, 4, 100])).toMatchObject({ lo: expect.any(Number) });
    expect(crc32('hello')).toBe('3610a686');
    expect(base64Decode(base64Encode('héllo'))).toBe('héllo');
  });
  it('dates and geography helpers', () => {
    expect(easter(2024)).toEqual({ m: 3, d: 31 }); expect(easter(2025)).toEqual({ m: 4, d: 20 });
    const h = usHolidays(2026);
    expect(h.find((x) => x.name === 'Thanksgiving Day')!.serial).toBe(partsToSerial(2026, 11, 26));
    expect(h.find((x) => x.name === 'Independence Day')!.serial).toBe(partsToSerial(2026, 7, 3));
    expect(weekNumberIso(partsToSerial(2024, 12, 30))).toBe(1);
    expect(dmsToDecimal('41°24\'12.2"N')).toBeCloseTo(41.403389, 5);
    expect(dmsToDecimal('2 10 26.5 W')).toBeCloseTo(-2.174028, 5);
    expect(decimalToDms(-2.174028, 'lon')).toBe('2°10\'26.5"W');
    expect(haversineKm(51.5007, -0.1246, 40.6892, -74.0445)).toBeCloseTo(5570, -1);
    expect(compassPoint(bearingDeg(0, 0, 0, 10))).toBe('E');
  });
  it('extras commands run on a sheet', async () => {
    const { app, put, get, select, book } = fakeApp();
    put('A1', 'Name'); put('A2', 'Robert'); put('A3', 'Rupert');
    select('A1:A3');
    app.exec('extra.soundex');
    expect(get('B2')).toBe('R163');
    put('C1', 'Amount'); put('C2', 100); put('C3', -40);
    select('C1:C3');
    await app.exec('extra.runningBalance', { opening: 10, mode: 'single' });
    expect(get('D3')).toBe(70);
    put('E1', 'the quick brown fox'); select('E1');
    app.exec('extra.titleCase');
    expect(get('E1')).toBe('The Quick Brown Fox');
    put('F1', 'a'); put('F2', 'a '); put('F3', 'A'); put('F4', 'b');
    select('F1:F4');
    await app.exec('extra.clusterValues', { method: 'fingerprint', apply: true });
    expect([get('F1'), get('F2'), get('F3')]).toEqual(['a', 'a', 'a']);
    await app.exec('extra.holidays', { year: 2025 });
    expect(book.sheets[book.sheets.length - 1].value(0, 0)).toBe('Holiday');
    expect(selCells(app).length).toBeGreaterThan(0);
  });
});

describe('review fixes: undo and copies', () => {
  it('undoes and redoes a tool that inserts rows and then writes cells', async () => {
    const { app, book, put, get, select } = fakeApp();
    put('A1', 'a'); put('B1', 1); put('A2', 'b'); put('B2', 2); put('A3', 'c'); put('B3', 3);
    select('A1:B3');
    await app.exec('util.duplicateRows', { times: 2 });
    const after = ['A1', 'A2', 'A3', 'A4', 'A5', 'A6'].map((a) => get(a));
    expect(after).toEqual(['a', 'a', 'b', 'b', 'c', 'c']);
    book.undo();
    expect(['A1', 'A2', 'A3', 'A4'].map((a) => get(a))).toEqual(['a', 'b', 'c', null]);
    expect(['B1', 'B2', 'B3'].map((a) => get(a))).toEqual([1, 2, 3]);
    book.redo();
    expect(['A1', 'A2', 'A3', 'A4', 'A5', 'A6'].map((a) => get(a))).toEqual(after);
  });
  it('gives each copied sheet its own id', async () => {
    const { app, book, put } = fakeApp();
    put('A1', 'orig');
    await app.exec('util.copySheetTimes', { count: 2 });
    expect(book.sheets.length).toBe(3);
    expect(new Set(book.sheets.map((s) => s.id)).size).toBe(3);
    expect(book.sheets[1].value(0, 0)).toBe('orig');
  });
  it('writes month and weekday names without reading their letters as codes', async () => {
    const { app, book, put, get, select } = fakeApp();
    put('A1', partsToSerial(2024, 9, 4)); put('A2', partsToSerial(2024, 12, 25));
    book.transact('f', () => { for (const r of [0, 1]) book.sheet.patch(r, 0, { s: book.derivedStyle(0, { numFmt: 'yyyy-mm-dd' }) }); });
    select('A1:A2');
    await app.exec('util.datesToText', { pattern: 'dddd d mmmm yyyy' });
    expect(get('A1')).toBe('Wednesday 4 September 2024');
    expect(get('A2')).toBe('Wednesday 25 December 2024');
  });
  it('reads zero-padded hexadecimal and counts first digits of fractions', () => {
    expect(fromBase('0B', 16)).toBe(11);
    expect(fromBase('0BAD', 16)).toBe(2989);
    expect(fromBase('0x1F', 16)).toBe(31);
    expect(fromBase('0b101', 2)).toBe(5);
    const b = benford([0.25, 0.31, 0.5, 20.5, 100, 3]);
    expect(b.map((x) => x.count)).toEqual([1, 2, 2, 0, 1, 0, 0, 0, 0]);
  });
});

describe('review fixes: filtered lists, protection and tools', () => {
  function filtered() {
    const f = fakeApp();
    const rows = [['k', 'v'], ['a', 1], ['b', 2], ['a', 3], ['b', 4], ['a', 5]];
    rows.forEach((row, r) => { f.put(`A${r + 1}`, row[0]); f.put(`B${r + 1}`, row[1]); });
    const s = f.book.sheet;
    s.filter = { range: { r1: 0, c1: 0, r2: 5, c2: 1 }, columns: { 0: { values: ['a'] } as any } };
    s.hiddenRows = new Set([2, 4]);
    return f;
  }
  it('Clear Contents leaves rows hidden by the filter alone', async () => {
    const { app, get, select } = filtered();
    select('B2:B6');
    await app.exec('edit.clearContents');
    expect(['B2', 'B3', 'B4', 'B5', 'B6'].map((a) => get(a) ?? null)).toEqual([null, 2, null, 4, null]);
  });
  it('Delete Rows removes only the rows that are showing', async () => {
    const { app, book, get } = filtered();
    app.grid.selectRows(1, 5);
    await app.exec('cells.deleteRows');
    expect(['A1', 'A2', 'A3', 'A4'].map((a) => get(a) ?? null)).toEqual(['k', 'b', 'b', null]);
    expect(['B2', 'B3'].map((a) => get(a))).toEqual([2, 4]);
    book.undo();
    expect(get('A6')).toBe('a');
  });
  it('refuses to unlock or overwrite cells on a protected sheet', async () => {
    const { app, book, put, get, select } = fakeApp();
    put('A1', 'secret');
    book.sheet.protection = { on: true, allow: ['selectLocked', 'selectUnlocked'] } as any;
    select('A1');
    await app.exec('util.unlockSelection');
    expect(book.sheet.styleOf(0, 0).locked).not.toBe(false);
    await app.exec('plus.insertRandomPassword');
    await app.exec('plus.batchCheckboxes');
    expect(get('A1')).toBe('secret');
  });
  it('encrypts and decrypts without changing values or types', async () => {
    const { app, book, put, cell, select } = fakeApp();
    put('A1', 3.14159); put('A2', '00123'); put('A3', 0.1 + 0.2); put('A4', true); put('A5', '=1+1'); put('A6', 45000.75);
    book.transact('f', () => book.sheet.patch(0, 0, { s: book.derivedStyle(0, { numFmt: '0.00' }) }));
    select('A1:A6');
    await app.exec('plus.encryptCells', { password: 'pw' });
    expect(typeof cell('A1')!.v).toBe('string');
    expect(String(cell('A1')!.v)).not.toContain('3.14');
    await app.exec('plus.decryptCells', { password: 'pw' });
    expect(cell('A1')!.v).toBe(3.14159);
    expect(cell('A2')!.v).toBe('00123');
    expect(cell('A3')!.v).toBe(0.1 + 0.2);
    expect(cell('A4')!.v).toBe(true);
    expect(cell('A5')!.f).toBe('1+1');
    expect(cell('A6')!.v).toBe(45000.75);
    expect(book.sheet.styleOf(0, 0).numFmt).toBe('0.00');
  });
  it('tells a link to another workbook from a table reference or bracketed text', () => {
    expect(externalLinkIn('SUM(Table1[Qty])')).toBeNull();
    expect(externalLinkIn('"see [1]"&A2')).toBeNull();
    expect(externalLinkIn('Table1[[#Totals],[Qty]]+1')).toBeNull();
    expect(externalLinkIn('[Book.xlsx]Sheet1!A1*2')).toBe('[Book.xlsx]');
    expect(externalLinkIn("SUM('[My Book.xlsx]Data'!A1:A3)")).toBe('[My Book.xlsx]');
  });
  it('stores numbers as text in full and keeps leading zeros when stripping characters', async () => {
    const { app, put, get, select } = fakeApp();
    put('A1', 1234567890123456); put('A2', 123456789.123456); put('A3', 'ZIP: 02134'); put('A4', 'Room 12 - A');
    select('A1:A2');
    await app.exec('util.numbersToText');
    expect(get('A1')).toBe('1234567890123456');
    expect(get('A2')).toBe('123456789.123456');
    select('A3:A4');
    await app.exec('util.keepOnly', { what: 'digits' });
    expect(get('A3')).toBe('02134');
    expect(get('A4')).toBe('12-');
  });
  it('fills every selected cell, also beyond the used range', async () => {
    const { app, put, get, select } = fakeApp();
    put('A1', 1);
    select('D5:E10');
    await app.exec('util.fillToday');
    expect(typeof get('D5')).toBe('number');
    expect(typeof get('E10')).toBe('number');
    select('G2:G4');
    await app.exec('plus.batchCheckboxes');
    expect(['G2', 'G3', 'G4'].map((a) => get(a))).toEqual([false, false, false]);
  });
  it('keeps a sheet of the user called index when making a sheet index', async () => {
    const { app, book, put } = fakeApp();
    book.transact('n', () => { book.sheets.push(new Sheet(book, 'index')); });
    put('A1', 'mine', book.sheets[1]);
    await app.exec('util.sheetIndex');
    expect(book.sheets.some((s) => s.name === 'index' && s.value(0, 0) === 'mine')).toBe(true);
    const n = book.sheets.length;
    await app.exec('util.sheetIndex');
    expect(book.sheets.length).toBe(n);
  });
});
