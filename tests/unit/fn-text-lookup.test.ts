import { describe, it, expect } from 'vitest';
import '../../src/calc/functions/logical';
import '../../src/calc/functions/text';
import '../../src/calc/functions/datetime';
import '../../src/calc/functions/info';
import '../../src/calc/functions/lookup';
import '../../src/calc/functions/dynamic';
import { defFn, getFn } from '../../src/calc/functions/registry';
import { Workbook } from '../../src/model/workbook';
import { Engine } from '../../src/calc/engine';
import { parseCell } from '../../src/model/address';
import { CellError, type Scalar } from '../../src/model/types';
import { iterValues, type Value } from '../../src/calc/values';

// Minimal SUM/COUNT for aggregation tests when the math module is not loaded.
if (!getFn('SUM')) {
  defFn({ name: 'SUM', min: 1, max: 255, category: 'Math & Trig', desc: '', params: [], arrays: 'all', impl: (args, ctx) => { let s = 0; for (const a of args) for (const { v } of iterValues(a as Value, ctx)) if (typeof v === 'number') s += v; return s; } });
}
if (!getFn('COUNT')) {
  defFn({ name: 'COUNT', min: 1, max: 255, category: 'Statistical', desc: '', params: [], arrays: 'all', impl: (args, ctx) => { let s = 0; for (const a of args) for (const { v } of iterValues(a as Value, ctx)) if (typeof v === 'number') s++; return s; } });
}

const norm = (v: Scalar): any => (v instanceof CellError ? v.code : v);

function book(data: Record<string, any> = {}) {
  const b = new Workbook();
  const e = new Engine(b);
  const s = b.sheet;
  const set = (a: string, v: any, sheet = s) => {
    const p = parseCell(a)!;
    b.transact('t', () => sheet.set(p.r, p.c, typeof v === 'string' && v.startsWith('=') ? { f: v.slice(1) } : { v }));
  };
  for (const [k, v] of Object.entries(data)) set(k, v);
  const get = (a: string, sheet = s) => { e.recalc(); const p = parseCell(a)!; return norm(sheet.value(p.r, p.c)); };
  /** Evaluates a formula in cell Z100 and returns a scalar (errors as codes). */
  const val = (f: string) => { e.recalc(); return norm(e.evaluateScalar(f, s, 99, 25)); };
  const arr = (f: string) => { e.recalc(); return e.evaluateMatrix(f, s, 99, 25).map((r) => r.map(norm)); };
  return { b, e, s, set, get, val, arr };
}

describe('text functions', () => {
  const t = book({ A1: 'Hello World', A2: '  a   b  ', A3: 12.5, A4: true, A5: 'x' });
  it('basic slicing', () => {
    expect(t.val('LEFT(A1,5)')).toBe('Hello');
    expect(t.val('LEFT(A1)')).toBe('H');
    expect(t.val('RIGHT(A1,5)')).toBe('World');
    expect(t.val('RIGHT(A1,0)')).toBe('');
    expect(t.val('MID(A1,7,3)')).toBe('Wor');
    expect(t.val('MID(A1,20,3)')).toBe('');
    expect(t.val('MID(A1,0,3)')).toBe('#VALUE!');
    expect(t.val('LEFT(A1,-1)')).toBe('#VALUE!');
    expect(t.val('LEN(A1)')).toBe(11);
    expect(t.val('LEN(A3)')).toBe(4);
    expect(t.val('LENB("abc")')).toBe(3);
    expect(t.val('LEFTB("abc",2)')).toBe('ab');
  });
  it('case and cleanup', () => {
    expect(t.val('UPPER("abc")')).toBe('ABC');
    expect(t.val('LOWER("ABC")')).toBe('abc');
    expect(t.val('PROPER("hello wORLD 2nd it\'s")')).toBe("Hello World 2Nd It'S");
    expect(t.val('TRIM(A2)')).toBe('a b');
    expect(t.val('CLEAN("a"&CHAR(10)&"b")')).toBe('ab');
    expect(t.val('EXACT("abc","ABC")')).toBe(false);
    expect(t.val('EXACT("abc","abc")')).toBe(true);
  });
  it('FIND and SEARCH', () => {
    expect(t.val('FIND("o",A1)')).toBe(5);
    expect(t.val('FIND("o",A1,6)')).toBe(8);
    expect(t.val('FIND("O",A1)')).toBe('#VALUE!');
    expect(t.val('FIND("",A1)')).toBe(1);
    expect(t.val('FIND("o",A1,20)')).toBe('#VALUE!');
    expect(t.val('SEARCH("O",A1)')).toBe(5);
    expect(t.val('SEARCH("w*d",A1)')).toBe(7);
    expect(t.val('SEARCH("?o",A1)')).toBe(4);
    expect(t.val('SEARCH("~*","a*b")')).toBe(2);
    expect(t.val('SEARCH("z",A1)')).toBe('#VALUE!');
    expect(t.val('FINDB("l",A1)')).toBe(3);
    expect(t.val('SEARCHB("L",A1)')).toBe(3);
  });
  it('REPLACE, SUBSTITUTE, REPT', () => {
    expect(t.val('REPLACE("abcdef",2,3,"X")')).toBe('aXef');
    expect(t.val('REPLACEB("abcdef",1,0,"X")')).toBe('Xabcdef');
    expect(t.val('SUBSTITUTE("a-b-c","-","+")')).toBe('a+b+c');
    expect(t.val('SUBSTITUTE("a-b-c","-","+",2)')).toBe('a-b+c');
    expect(t.val('SUBSTITUTE("a-b-c","-","+",3)')).toBe('a-b-c');
    expect(t.val('SUBSTITUTE("abc","","x")')).toBe('abc');
    expect(t.val('SUBSTITUTE("abc","b","x",0)')).toBe('#VALUE!');
    expect(t.val('REPT("ab",3)')).toBe('ababab');
    expect(t.val('REPT("ab",-1)')).toBe('#VALUE!');
    expect(t.val('REPT("ab",20000)')).toBe('#VALUE!');
  });
  it('CHAR, CODE, UNICHAR, UNICODE', () => {
    expect(t.val('CHAR(65)')).toBe('A');
    expect(t.val('CHAR(128)')).toBe('€');
    expect(t.val('CHAR(0)')).toBe('#VALUE!');
    expect(t.val('CHAR(256)')).toBe('#VALUE!');
    expect(t.val('CODE("A")')).toBe(65);
    expect(t.val('CODE("€")')).toBe(128);
    expect(t.val('CODE("")')).toBe('#VALUE!');
    expect(t.val('UNICHAR(9731)')).toBe('☃');
    expect(t.val('UNICHAR(0)')).toBe('#VALUE!');
    expect(t.val('UNICODE("☃")')).toBe(9731);
    expect(t.val('UNICODE("😀")')).toBe(128512);
  });
  it('CONCAT, CONCATENATE, TEXTJOIN', () => {
    const b = book({ A1: 'a', A2: null, A3: 'c', B1: 1, B2: 2, B3: 3 });
    expect(b.val('CONCAT(A1:A3,B1:B3)')).toBe('ac123');
    expect(b.val('CONCATENATE("x",1,TRUE)')).toBe('x1TRUE');
    expect(b.arr('CONCATENATE(B1:B3,"!")')).toEqual([['1!'], ['2!'], ['3!']]);
    expect(b.val('TEXTJOIN(",",TRUE,A1:A3)')).toBe('a,c');
    expect(b.val('TEXTJOIN(",",FALSE,A1:A3)')).toBe('a,,c');
    expect(b.val('TEXTJOIN({"-","+"},TRUE,B1:B3,"z")')).toBe('1-2+3-z');
    expect(b.val('TEXTJOIN(",",TRUE,1/0,"a")')).toBe('#DIV/0!');
    expect(b.val('CONCAT(1/0)')).toBe('#DIV/0!');
  });
  it('T, VALUE, NUMBERVALUE', () => {
    expect(t.val('T("abc")')).toBe('abc');
    expect(t.val('T(12)')).toBe('');
    expect(t.val('T(NA())')).toBe('#N/A');
    expect(t.val('VALUE("1,234.5")')).toBe(1234.5);
    expect(t.val('VALUE("$10")')).toBe(10);
    expect(t.val('VALUE("50%")')).toBe(0.5);
    expect(t.val('VALUE("12:00")')).toBe(0.5);
    expect(t.val('VALUE("abc")')).toBe('#VALUE!');
    expect(t.val('VALUE(TRUE)')).toBe('#VALUE!');
    expect(t.val('NUMBERVALUE("2.500,27",",",".")')).toBe(2500.27);
    expect(t.val('NUMBERVALUE("3.5%")')).toBeCloseTo(0.035, 12);
    expect(t.val('NUMBERVALUE("1 000")')).toBe(1000);
    expect(t.val('NUMBERVALUE("")')).toBe(0);
    expect(t.val('NUMBERVALUE("1.2.3")')).toBe('#VALUE!');
    expect(t.val('NUMBERVALUE("1.000,5",".",".")')).toBe('#VALUE!');
    expect(t.val('NUMBERVALUE("1,5,2",",",".")')).toBe('#VALUE!');
  });
  it('TEXT, FIXED, DOLLAR', () => {
    expect(t.val('TEXT(1234.567,"#,##0.00")')).toBe('1,234.57');
    expect(t.val('TEXT(0.256,"0.0%")')).toBe('25.6%');
    expect(t.val('TEXT(45366,"yyyy-mm-dd")')).toBe('2024-03-15');
    expect(t.val('TEXT(45366,"dddd, mmmm d")')).toBe('Friday, March 15');
    expect(t.val('TEXT("3/15/2024","mmm yyyy")')).toBe('Mar 2024');
    expect(t.val('TEXT(0.75,"h:mm AM/PM")')).toBe('6:00 PM');
    expect(t.val('TEXT("abc","0.00")')).toBe('abc');
    expect(t.val('TEXT(5,"000")')).toBe('005');
    expect(t.val('FIXED(1234.567)')).toBe('1,234.57');
    expect(t.val('FIXED(1234.567,1,TRUE)')).toBe('1234.6');
    expect(t.val('FIXED(-1234.567,-2)')).toBe('-1,200');
    expect(t.val('FIXED(2.5,0)')).toBe('3');
    expect(t.val('DOLLAR(1234.567)')).toBe('$1,234.57');
    expect(t.val('DOLLAR(-1234.567,1)')).toBe('($1,234.6)');
    expect(t.val('DOLLAR(1234.567,-2)')).toBe('$1,200');
  });
  it('TEXTBEFORE and TEXTAFTER', () => {
    expect(t.val('TEXTBEFORE("a-b-c","-")')).toBe('a');
    expect(t.val('TEXTBEFORE("a-b-c","-",2)')).toBe('a-b');
    expect(t.val('TEXTBEFORE("a-b-c","-",-1)')).toBe('a-b');
    expect(t.val('TEXTAFTER("a-b-c","-")')).toBe('b-c');
    expect(t.val('TEXTAFTER("a-b-c","-",-1)')).toBe('c');
    expect(t.val('TEXTAFTER("a-b-c","-",3)')).toBe('#N/A');
    expect(t.val('TEXTAFTER("a-b-c","-",3,,,"none")')).toBe('none');
    expect(t.val('TEXTAFTER("a-b-c","-",0)')).toBe('#VALUE!');
    expect(t.val('TEXTAFTER("aXbxc","x",1,1)')).toBe('bxc');
    expect(t.val('TEXTAFTER("aXbxc","x",1,0)')).toBe('c');
    expect(t.val('TEXTBEFORE("a-b","-",2,,1)')).toBe('a-b');
    expect(t.val('TEXTAFTER("a-b","-",2,,1)')).toBe('');
    expect(t.val('TEXTAFTER("a,b;c",{",",";"},2)')).toBe('c');
    expect(t.val('TEXTBEFORE("abc","")')).toBe('');
    expect(t.val('TEXTAFTER("abc","")')).toBe('abc');
  });
  it('TEXTSPLIT', () => {
    expect(t.arr('TEXTSPLIT("a,b,c",",")')).toEqual([['a', 'b', 'c']]);
    expect(t.arr('TEXTSPLIT("a,b;c",",",";")')).toEqual([['a', 'b'], ['c', '#N/A']]);
    expect(t.arr('TEXTSPLIT("a,b;c",",",";",,,"-")')).toEqual([['a', 'b'], ['c', '-']]);
    expect(t.arr('TEXTSPLIT("a,,b",",")')).toEqual([['a', '', 'b']]);
    expect(t.arr('TEXTSPLIT("a,,b",",",,TRUE)')).toEqual([['a', 'b']]);
    expect(t.arr('TEXTSPLIT("a-b+c",{"-","+"})')).toEqual([['a', 'b', 'c']]);
    expect(t.arr('TEXTSPLIT("aXbxc","x",,,1)')).toEqual([['a', 'b', 'c']]);
    expect(t.arr('TEXTSPLIT("a;b",,";")')).toEqual([['a'], ['b']]);
  });
  it('regex functions', () => {
    expect(t.val('REGEXTEST("abc123","\\d+")')).toBe(true);
    expect(t.val('REGEXTEST("ABC","abc")')).toBe(false);
    expect(t.val('REGEXTEST("ABC","abc",1)')).toBe(true);
    expect(t.val('REGEXEXTRACT("abc123def45","\\d+")')).toBe('123');
    expect(t.arr('REGEXEXTRACT("abc123def45","\\d+",1)')).toEqual([['123'], ['45']]);
    expect(t.arr('REGEXEXTRACT("John Smith","(\\w+) (\\w+)",2)')).toEqual([['John', 'Smith']]);
    expect(t.val('REGEXEXTRACT("abc","\\d")')).toBe('#N/A');
    expect(t.val('REGEXREPLACE("a1b2c3","\\d","#")')).toBe('a#b#c#');
    expect(t.val('REGEXREPLACE("a1b2c3","\\d","#",2)')).toBe('a1b#c3');
    expect(t.val('REGEXREPLACE("a1b2c3","\\d","#",-1)')).toBe('a1b2c#');
    expect(t.val('REGEXREPLACE("John Smith","(\\w+) (\\w+)","$2, $1")')).toBe('Smith, John');
    expect(t.val('REGEXTEST("a","(")')).toBe('#VALUE!');
  });
  it('ARRAYTOTEXT, VALUETOTEXT', () => {
    expect(t.val('ARRAYTOTEXT({1,"a";TRUE,2})')).toBe('1, a, TRUE, 2');
    expect(t.val('ARRAYTOTEXT({1,"a";TRUE,2},1)')).toBe('{1,"a";TRUE,2}');
    expect(t.val('VALUETOTEXT("a")')).toBe('a');
    expect(t.val('VALUETOTEXT("a",1)')).toBe('"a"');
    expect(t.val('VALUETOTEXT(NA())')).toBe('#N/A');
    expect(t.val('ARRAYTOTEXT({1},2)')).toBe('#VALUE!');
  });
  it('locale and misc', () => {
    expect(t.val('ASC("ＡＢＣ　１")')).toBe('ABC 1');
    expect(t.val('DBCS("AB 1")')).toBe('ＡＢ　１');
    expect(t.val('JIS("A")')).toBe('Ａ');
    expect(t.val('ASC("ガ")')).toBe('ｶﾞ');
    expect(t.val('DBCS("ｶﾞ")')).toBe('ガ');
    expect(t.val('BAHTTEXT(1234)')).toBe('หนึ่งพันสองร้อยสามสิบสี่บาทถ้วน');
    expect(t.val('BAHTTEXT(21.5)')).toBe('ยี่สิบเอ็ดบาทห้าสิบสตางค์');
    expect(t.val('BAHTTEXT(0)')).toBe('ศูนย์บาทถ้วน');
    expect(t.val('BAHTTEXT(11)')).toBe('สิบเอ็ดบาทถ้วน');
    expect(t.val('BAHTTEXT(1000001)')).toBe('หนึ่งล้านเอ็ดบาทถ้วน');
    const b = book({ A1: 'ab', A2: 5, A3: 'cd' });
    expect(b.val('PHONETIC(A1:A3)')).toBe('abcd');
  });
});

describe('date and time functions', () => {
  const t = book({ A1: 45366, A2: '=DATE(2024,1,31)' });
  it('DATE and parts', () => {
    expect(t.val('DATE(2024,3,15)')).toBe(45366);
    expect(t.val('DATE(2024,14,1)')).toBe(45689); // 2025-02-01
    expect(t.val('DATE(2024,0,1)')).toBe(45261); // 2023-12-01
    expect(t.val('DATE(2024,3,0)')).toBe(45351); // 2024-02-29
    expect(t.val('DATE(100,1,1)')).toBe(36526); // year 2000
    expect(t.val('DATE(1900,2,29)')).toBe(60);
    expect(t.val('DATE(1900,3,1)')).toBe(61);
    expect(t.val('DATE(1900,1,0)')).toBe(0);
    expect(t.val('DATE(1899,12,31)')).toBe(t.val('DATE(3799,12,31)'));
    expect(t.val('DATE(-1,1,1)')).toBe('#NUM!');
    expect(t.val('DATE(10000,1,1)')).toBe('#NUM!');
    expect(t.val('YEAR(A1)')).toBe(2024);
    expect(t.val('MONTH(A1)')).toBe(3);
    expect(t.val('DAY(A1)')).toBe(15);
    expect(t.val('DAY(60)')).toBe(29);
    expect(t.val('MONTH(60)')).toBe(2);
    expect(t.val('DAY(0)')).toBe(0);
    expect(t.val('YEAR(0)')).toBe(1900);
    expect(t.val('DAY(-1)')).toBe('#NUM!');
    expect(t.val('YEAR("2024-06-01")')).toBe(2024);
    expect(t.val('HOUR(0.75)')).toBe(18);
    expect(t.val('MINUTE("10:45:30")')).toBe(45);
    expect(t.val('SECOND(TIME(1,2,3))')).toBe(3);
    expect(t.val('SECOND(0.999999999)')).toBe(0);
  });
  it('TIME, DATEVALUE, TIMEVALUE', () => {
    expect(t.val('TIME(12,0,0)')).toBe(0.5);
    expect(t.val('TIME(25,0,0)')).toBeCloseTo(1 / 24, 12);
    expect(t.val('TIME(0,-1,0)')).toBe('#NUM!');
    expect(t.val('TIME(1,-30,0)')).toBeCloseTo(0.5 / 24, 12);
    expect(t.val('DATEVALUE("2024-03-15")')).toBe(45366);
    expect(t.val('DATEVALUE("3/15/2024 10:00")')).toBe(45366);
    expect(t.val('DATEVALUE("15-Mar-2024")')).toBe(45366);
    expect(t.val('DATEVALUE("xyz")')).toBe('#VALUE!');
    expect(t.val('DATEVALUE(45366)')).toBe('#VALUE!');
    expect(t.val('TIMEVALUE("6:00 PM")')).toBe(0.75);
    expect(t.val('TIMEVALUE("2024-01-01 06:00")')).toBe(0.25);
  });
  it('DAYS, DAYS360, DATEDIF', () => {
    expect(t.val('DAYS(DATE(2024,3,15),DATE(2024,1,1))')).toBe(74);
    expect(t.val('DAYS("2024-03-15","2024-01-01")')).toBe(74);
    expect(t.val('DAYS360(DATE(2011,2,28),DATE(2011,3,31))')).toBe(30);
    expect(t.val('DAYS360(DATE(2011,1,30),DATE(2011,12,31))')).toBe(330);
    expect(t.val('DAYS360(DATE(2011,1,30),DATE(2011,12,31),TRUE)')).toBe(330);
    expect(t.val('DAYS360(DATE(2011,1,1),DATE(2011,12,31),TRUE)')).toBe(359);
    expect(t.val('DAYS360(DATE(2011,1,1),DATE(2011,12,31))')).toBe(360);
    expect(t.val('DATEDIF(DATE(2001,1,1),DATE(2003,1,1),"Y")')).toBe(2);
    expect(t.val('DATEDIF(DATE(2001,6,1),DATE(2002,8,15),"D")')).toBe(440);
    expect(t.val('DATEDIF(DATE(2001,6,1),DATE(2002,8,15),"YD")')).toBe(75);
    expect(t.val('DATEDIF(DATE(2001,6,1),DATE(2002,8,15),"M")')).toBe(14);
    expect(t.val('DATEDIF(DATE(2001,6,1),DATE(2002,8,15),"YM")')).toBe(2);
    expect(t.val('DATEDIF(DATE(2001,6,20),DATE(2002,8,15),"MD")')).toBe(26);
    expect(t.val('DATEDIF(DATE(2003,1,1),DATE(2001,1,1),"Y")')).toBe('#NUM!');
    expect(t.val('DATEDIF(DATE(2001,1,1),DATE(2003,1,1),"Q")')).toBe('#NUM!');
  });
  it('EDATE, EOMONTH', () => {
    expect(t.val('EDATE(A2,1)')).toBe(45351); // 2024-02-29
    expect(t.val('EDATE(A2,-2)')).toBe(45260); // 2023-11-30
    expect(t.val('EOMONTH(A1,0)')).toBe(45382);
    expect(t.val('EOMONTH(A1,-1)')).toBe(45351);
    expect(t.val('EOMONTH(A1,1.9)')).toBe(45412);
    expect(t.val('EDATE(1,-1)')).toBe('#NUM!');
  });
  it('WEEKDAY, WEEKNUM, ISOWEEKNUM', () => {
    expect(t.val('WEEKDAY(A1)')).toBe(6); // Friday
    expect(t.val('WEEKDAY(A1,2)')).toBe(5);
    expect(t.val('WEEKDAY(A1,3)')).toBe(4);
    expect(t.val('WEEKDAY(A1,16)')).toBe(7); // Saturday-based: Friday is 7
    expect(t.val('WEEKDAY(1)')).toBe(1);
    expect(t.val('WEEKDAY(A1,4)')).toBe('#NUM!');
    expect(t.val('WEEKNUM(DATE(2024,1,1))')).toBe(1);
    expect(t.val('WEEKNUM(DATE(2024,3,9))')).toBe(10);
    expect(t.val('WEEKNUM(DATE(2024,3,9),2)')).toBe(10);
    expect(t.val('WEEKNUM(DATE(2024,3,10),2)')).toBe(10);
    expect(t.val('WEEKNUM(DATE(2024,3,10))')).toBe(11);
    expect(t.val('WEEKNUM(DATE(2021,1,1),21)')).toBe(53);
    expect(t.val('ISOWEEKNUM(DATE(2021,1,4))')).toBe(1);
    expect(t.val('ISOWEEKNUM(DATE(2024,12,30))')).toBe(1);
    expect(t.val('ISOWEEKNUM(A1)')).toBe(11);
  });
  it('NETWORKDAYS and WORKDAY', () => {
    const b = book({ H1: 45370, H2: 45371 }); // 2024-03-19, 2024-03-20
    expect(b.val('NETWORKDAYS(DATE(2024,3,1),DATE(2024,3,31))')).toBe(21);
    expect(b.val('NETWORKDAYS(DATE(2024,3,1),DATE(2024,3,31),H1:H2)')).toBe(19);
    expect(b.val('NETWORKDAYS(DATE(2024,3,31),DATE(2024,3,1))')).toBe(-21);
    expect(b.val('NETWORKDAYS.INTL(DATE(2024,3,1),DATE(2024,3,31),11)')).toBe(26);
    expect(b.val('NETWORKDAYS.INTL(DATE(2024,3,1),DATE(2024,3,31),"0000011")')).toBe(21);
    // Every day a weekend: there are no working days to count, but nothing is wrong with the request.
    expect(b.val('NETWORKDAYS.INTL(DATE(2024,3,1),DATE(2024,3,31),"1111111")')).toBe(0);
    expect(b.val('WORKDAY.INTL(DATE(2024,3,1),5,"1111111")')).toBe('#VALUE!');
    expect(b.val('NETWORKDAYS.INTL(DATE(2024,3,1),DATE(2024,3,31),8)')).toBe('#NUM!');
    expect(b.val('WORKDAY(DATE(2024,3,15),1)')).toBe(45369); // Monday 18th
    expect(b.val('WORKDAY(DATE(2024,3,15),3,H1:H2)')).toBe(45373); // skips 19th, 20th
    expect(b.val('WORKDAY(DATE(2024,3,18),-1)')).toBe(45366);
    expect(b.val('WORKDAY.INTL(DATE(2024,3,15),1,7)')).toBe(45368); // Fri/Sat weekend -> Sunday 17th
    expect(b.val('WORKDAY.INTL(DATE(2024,3,15),1,"0000000")')).toBe(45367);
  });
  it('YEARFRAC all bases', () => {
    expect(t.val('YEARFRAC(DATE(2012,1,1),DATE(2012,7,30))')).toBeCloseTo(0.58055556, 7);
    expect(t.val('YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),1)')).toBeCloseTo(0.57650273, 7);
    expect(t.val('YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),2)')).toBeCloseTo(0.58611111, 7);
    expect(t.val('YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),3)')).toBeCloseTo(0.57808219, 7);
    expect(t.val('YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),4)')).toBeCloseTo(0.58055556, 7);
    expect(t.val('YEARFRAC(DATE(2010,1,1),DATE(2013,7,1),1)')).toBeCloseTo(1277 / 365.25, 10);
    expect(t.val('YEARFRAC(DATE(2011,2,28),DATE(2012,2,29),0)')).toBe(1);
    expect(t.val('YEARFRAC(DATE(2012,7,30),DATE(2012,1,1))')).toBeCloseTo(0.58055556, 7);
    expect(t.val('YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),5)')).toBe('#NUM!');
  });
  it('NOW and TODAY are volatile', () => {
    const b = book({ A1: '=TODAY()', A2: '=NOW()' });
    const today = b.get('A1') as number, now = b.get('A2') as number;
    expect(Number.isInteger(today)).toBe(true);
    expect(now).toBeGreaterThanOrEqual(today);
    expect(now - today).toBeLessThan(1);
  });
});

describe('information functions', () => {
  const t = book({ A1: 1, A2: 'x', A3: true, A4: '=1/0', A5: '=NA()', B1: '=A1*2' });
  it('IS functions', () => {
    expect(t.val('ISBLANK(C1)')).toBe(true);
    expect(t.val('ISBLANK(A1)')).toBe(false);
    expect(t.val('ISBLANK("")')).toBe(false);
    expect(t.val('ISERR(A4)')).toBe(true);
    expect(t.val('ISERR(A5)')).toBe(false);
    expect(t.val('ISERROR(A5)')).toBe(true);
    expect(t.val('ISNA(A5)')).toBe(true);
    expect(t.val('ISNA(A4)')).toBe(false);
    expect(t.val('ISNUMBER(A1)')).toBe(true);
    expect(t.val('ISNUMBER("1")')).toBe(false);
    expect(t.val('ISTEXT(A2)')).toBe(true);
    expect(t.val('ISNONTEXT(A2)')).toBe(false);
    expect(t.val('ISNONTEXT(C1)')).toBe(true);
    expect(t.val('ISLOGICAL(A3)')).toBe(true);
    expect(t.val('ISEVEN(2.5)')).toBe(true);
    expect(t.val('ISODD(-3)')).toBe(true);
    expect(t.val('ISEVEN("x")')).toBe('#VALUE!');
    expect(t.val('ISREF(A1)')).toBe(true);
    expect(t.val('ISREF(A1:B2)')).toBe(true);
    expect(t.val('ISREF(1)')).toBe(false);
    expect(t.val('ISREF(INDEX(A1:A3,2))')).toBe(true);
    expect(t.val('ISFORMULA(B1)')).toBe(true);
    expect(t.val('ISFORMULA(A1)')).toBe(false);
    expect(t.arr('ISNUMBER(A1:A3)')).toEqual([[true], [false], [false]]);
  });
  it('ERROR.TYPE, N, NA, TYPE', () => {
    expect(t.val('ERROR.TYPE(A4)')).toBe(2);
    expect(t.val('ERROR.TYPE(A5)')).toBe(7);
    expect(t.val('ERROR.TYPE(1)')).toBe('#N/A');
    expect(t.val('N(A3)')).toBe(1);
    expect(t.val('N("7")')).toBe(0);
    expect(t.val('N(A4)')).toBe('#DIV/0!');
    expect(t.val('NA()')).toBe('#N/A');
    expect(t.val('TYPE(1)')).toBe(1);
    expect(t.val('TYPE("a")')).toBe(2);
    expect(t.val('TYPE(TRUE)')).toBe(4);
    expect(t.val('TYPE(A4)')).toBe(16);
    expect(t.val('TYPE({1,2})')).toBe(64);
    expect(t.val('TYPE(A1:A2)')).toBe(64);
    expect(t.val('TYPE(C1)')).toBe(1);
  });
  it('SHEET, SHEETS, CELL, INFO', () => {
    const b = book({ A1: 5 });
    b.b.transact('add', () => { const s2 = new (b.s.constructor as any)(b.b, 'Data'); b.b.sheets.push(s2); });
    b.b.transact('add', () => { const s3 = new (b.s.constructor as any)(b.b, 'Other'); b.b.sheets.push(s3); });
    expect(b.val('SHEET()')).toBe(1);
    expect(b.val('SHEET(Data!A1)')).toBe(2);
    expect(b.val('SHEET("Other")')).toBe(3);
    expect(b.val('SHEET("Nope")')).toBe('#N/A');
    expect(b.val('SHEETS()')).toBe(3);
    expect(b.val('SHEETS(Sheet1:Other!A1)')).toBe(3);
    expect(b.val('SHEETS(A1)')).toBe(1);
    expect(b.val('CELL("address",B3)')).toBe('$B$3');
    expect(b.val('CELL("address",Data!B3)')).toBe('Data!$B$3');
    expect(b.val('CELL("row",B3)')).toBe(3);
    expect(b.val('CELL("col",B3)')).toBe(2);
    expect(b.val('CELL("contents",A1)')).toBe(5);
    expect(b.val('CELL("type",A1)')).toBe('v');
    expect(b.val('CELL("type",A9)')).toBe('b');
    expect(b.val('CELL("format",A1)')).toBe('G');
    expect(b.val('CELL("protect",A1)')).toBe(1);
    expect(b.val('CELL("width",A1)')).toBe(8);
    expect(b.val('CELL("bogus",A1)')).toBe('#VALUE!');
    expect(b.val('INFO("numfile")')).toBe(3);
    expect(b.val('INFO("recalc")')).toBe('Automatic');
    b.b.transact('fmt', () => { b.s.set(1, 0, { v: 'txt', s: b.b.styleId({ numFmt: '0.00%', hAlign: 'center' }) }); });
    expect(b.val('CELL("format",A2)')).toBe('P2');
    expect(b.val('CELL("prefix",A2)')).toBe('^');
  });
});

describe('lookup and reference', () => {
  const data = {
    A1: 'Name', B1: 'Qty', C1: 'Price',
    A2: 'apple', B2: 10, C2: 1.5,
    A3: 'banana', B3: 20, C3: 0.25,
    A4: 'cherry', B4: 30, C4: 3,
    A5: 'date', B5: 40, C5: 5,
    E1: 1, E2: 3, E3: 5, E4: 7, E5: 9,
    F1: 9, F2: 7, F3: 5, F4: 3, F5: 1,
  };
  const t = book(data);
  it('VLOOKUP and HLOOKUP', () => {
    expect(t.val('VLOOKUP("banana",A2:C5,2,FALSE)')).toBe(20);
    expect(t.val('VLOOKUP("BANANA",A2:C5,3,0)')).toBe(0.25);
    expect(t.val('VLOOKUP("b*",A2:C5,2,FALSE)')).toBe(20);
    expect(t.val('VLOOKUP("?herry",A2:C5,2,FALSE)')).toBe(30);
    expect(t.val('VLOOKUP("fig",A2:C5,2,FALSE)')).toBe('#N/A');
    expect(t.val('VLOOKUP("c",A2:C5,2)')).toBe(20);
    expect(t.val('VLOOKUP("zzz",A2:C5,2,TRUE)')).toBe(40);
    expect(t.val('VLOOKUP("a",A2:C5,2,TRUE)')).toBe('#N/A');
    expect(t.val('VLOOKUP("apple",A2:C5,4,FALSE)')).toBe('#REF!');
    expect(t.val('VLOOKUP("apple",A2:C5,0,FALSE)')).toBe('#VALUE!');
    expect(t.val('VLOOKUP(6,E1:F5,2)')).toBe(5);
    expect(t.val('VLOOKUP(6,E1:F5,2,)')).toBe('#N/A');
    expect(t.val('HLOOKUP("Qty",A1:C5,3,FALSE)')).toBe(20);
    expect(t.val('HLOOKUP("Q",A1:C5,2)')).toBe('apple');
    expect(t.arr('VLOOKUP({"date";"apple"},A2:C5,2,FALSE)')).toEqual([[40], [10]]);
  });
  it('VLOOKUP approximate on unsorted data uses binary search', () => {
    const b = book({ A1: 1, A2: 5, A3: 2, A4: 8, A5: 3, B1: 'a', B2: 'b', B3: 'c', B4: 'd', B5: 'e' });
    // Binary search: mid=A3(2)<=4 -> right; mid=A4(8)>4 -> left; result A3.
    expect(b.val('VLOOKUP(4,A1:B5,2)')).toBe('c');
    expect(b.val('MATCH(4,A1:A5)')).toBe(3);
  });
  it('MATCH', () => {
    expect(t.val('MATCH(5,E1:E5,0)')).toBe(3);
    expect(t.val('MATCH(6,E1:E5)')).toBe(3);
    expect(t.val('MATCH(6,E1:E5,1)')).toBe(3);
    expect(t.val('MATCH(0,E1:E5,1)')).toBe('#N/A');
    expect(t.val('MATCH(6,F1:F5,-1)')).toBe(2);
    expect(t.val('MATCH(10,F1:F5,-1)')).toBe('#N/A');
    expect(t.val('MATCH("ch*",A1:A5,0)')).toBe(4);
    expect(t.val('MATCH("CHERRY",A1:A5,0)')).toBe(4);
    expect(t.val('MATCH(2,{1,2,2,2,3})')).toBe(4);
    expect(t.val('MATCH("5",E1:E5,0)')).toBe('#N/A');
    expect(t.val('MATCH(5,A1:C5,0)')).toBe('#N/A');
    expect(t.val('MATCH(5,{1,3,5,7},0)')).toBe(3);
  });
  it('INDEX', () => {
    expect(t.val('INDEX(A1:C5,3,2)')).toBe(20);
    expect(t.val('INDEX(A2:A5,2)')).toBe('banana');
    expect(t.val('INDEX(A1:C1,2)')).toBe('Qty');
    expect(t.arr('INDEX(A1:C5,0,2)')).toEqual([['Qty'], [10], [20], [30], [40]]);
    expect(t.arr('INDEX(A1:C5,2,0)')).toEqual([['apple', 10, 1.5]]);
    expect(t.val('INDEX(A1:C5,6,1)')).toBe('#REF!');
    expect(t.val('INDEX(A1:C5,-1,1)')).toBe('#VALUE!');
    expect(t.val('INDEX({1,2;3,4},2,1)')).toBe(3);
    expect(t.val('INDEX({1,2,3},2)')).toBe(2);
    expect(t.arr('INDEX({1,2;3,4},0,2)')).toEqual([[2], [4]]);
    expect(t.val('INDEX((A1:A5,E1:E5),2,1,2)')).toBe(3);
    expect(t.val('INDEX((A1:A5,E1:E5),2,1,3)')).toBe('#REF!');
    expect(t.val('SUM(INDEX(B2:B5,1):INDEX(B2:B5,3))')).toBe(60);
    expect(t.arr('INDEX(E1:E5,{1;3})')).toEqual([[1], [5]]);
    expect(t.val('INDEX(B1:B5,MATCH("cherry",A1:A5,0))')).toBe(30);
  });
  it('XLOOKUP', () => {
    expect(t.val('XLOOKUP("cherry",A2:A5,B2:B5)')).toBe(30);
    expect(t.val('XLOOKUP("fig",A2:A5,B2:B5)')).toBe('#N/A');
    expect(t.val('XLOOKUP("fig",A2:A5,B2:B5,"none")')).toBe('none');
    expect(t.arr('XLOOKUP("banana",A2:A5,B2:C5)')).toEqual([[20, 0.25]]);
    expect(t.val('XLOOKUP(6,E1:E5,F1:F5,,-1)')).toBe(5);
    expect(t.val('XLOOKUP(6,E1:E5,F1:F5,,1)')).toBe(3);
    expect(t.val('XLOOKUP(6,E1:E5,F1:F5,,0)')).toBe('#N/A');
    expect(t.val('XLOOKUP("b*",A2:A5,B2:B5,,2)')).toBe(20);
    expect(t.val('XLOOKUP("b*",A2:A5,B2:B5)')).toBe('#N/A');
    expect(t.val('XLOOKUP(6,E1:E5,F1:F5,,-1,2)')).toBe(5);
    expect(t.val('XLOOKUP(6,E1:E5,F1:F5,,1,2)')).toBe(3);
    expect(t.val('XLOOKUP(6,F1:F5,E1:E5,,1,-2)')).toBe(3);
    expect(t.val('XLOOKUP(6,F1:F5,E1:E5,,-1,-2)')).toBe(5);
    expect(t.val('XLOOKUP("Price",A1:C1,A4:C4)')).toBe(3);
    expect(t.val('XLOOKUP(1,{1,2,1},{"a","b","c"},,0,-1)')).toBe('c');
    expect(t.val('XLOOKUP(1,A2:A5,B2:B4)')).toBe('#VALUE!');
    expect(t.val('XLOOKUP("a*",A2:A5,B2:B5,,2,2)')).toBe('#VALUE!');
    expect(t.val('SUM(XLOOKUP("banana",A2:A5,B2:B5):XLOOKUP("date",A2:A5,B2:B5))')).toBe(90);
    expect(t.val('ISREF(XLOOKUP("banana",A2:A5,B2:B5))')).toBe(true);
  });
  it('XMATCH', () => {
    expect(t.val('XMATCH("cherry",A2:A5)')).toBe(3);
    expect(t.val('XMATCH(6,E1:E5,1)')).toBe(4);
    expect(t.val('XMATCH(6,E1:E5,-1)')).toBe(3);
    expect(t.val('XMATCH("?a*",A2:A5,2)')).toBe(2);
    expect(t.val('XMATCH(1,{1,2,1},0,-1)')).toBe(3);
    expect(t.val('XMATCH(7,E1:E5,0,2)')).toBe(4);
    expect(t.val('XMATCH(7,E1:E5,0,3)')).toBe('#VALUE!');
  });
  it('LOOKUP', () => {
    expect(t.val('LOOKUP(6,E1:E5,F1:F5)')).toBe(5);
    expect(t.val('LOOKUP(0,E1:E5,F1:F5)')).toBe('#N/A');
    expect(t.val('LOOKUP("c",A2:A5,B2:B5)')).toBe(20);
    expect(t.val('LOOKUP(6,E1:F5)')).toBe(5);
    expect(t.val('LOOKUP("b",{"a","b","c";1,2,3})')).toBe(2);
    expect(t.val('LOOKUP(9,E1:E5,G1)')).toBe(null);
  });
  it('ROW, COLUMN, ROWS, COLUMNS, ADDRESS, AREAS', () => {
    expect(t.val('ROW()')).toBe(100);
    expect(t.val('COLUMN()')).toBe(26);
    expect(t.val('ROW(C7)')).toBe(7);
    expect(t.val('COLUMN(C7)')).toBe(3);
    expect(t.arr('ROW(A2:A4)')).toEqual([[2], [3], [4]]);
    expect(t.arr('COLUMN(B1:D1)')).toEqual([[2, 3, 4]]);
    expect(t.val('ROWS(A1:C5)')).toBe(5);
    expect(t.val('COLUMNS(A1:C5)')).toBe(3);
    expect(t.val('ROWS({1,2;3,4;5,6})')).toBe(3);
    expect(t.val('ROWS(A:A)')).toBe(1048576);
    expect(t.val('ADDRESS(2,3)')).toBe('$C$2');
    expect(t.val('ADDRESS(2,3,2)')).toBe('C$2');
    expect(t.val('ADDRESS(2,3,3)')).toBe('$C2');
    expect(t.val('ADDRESS(2,3,4)')).toBe('C2');
    expect(t.val('ADDRESS(2,3,1,FALSE)')).toBe('R2C3');
    expect(t.val('ADDRESS(2,3,4,FALSE)')).toBe('R[2]C[3]');
    expect(t.val('ADDRESS(1,1,1,TRUE,"My Sheet")')).toBe("'My Sheet'!$A$1");
    expect(t.val('ADDRESS(1,1,1,TRUE,"Data")')).toBe('Data!$A$1');
    expect(t.val('ADDRESS(0,1)')).toBe('#VALUE!');
    expect(t.val('AREAS((A1:B2,C3,D4:E5))')).toBe(3);
    expect(t.val('AREAS(A1:B2)')).toBe(1);
  });
  it('OFFSET and INDIRECT', () => {
    expect(t.val('OFFSET(A1,2,1)')).toBe(20);
    expect(t.val('SUM(OFFSET(B1,1,0,4,1))')).toBe(100);
    expect(t.arr('OFFSET(A1,1,1,2,2)')).toEqual([[10, 1.5], [20, 0.25]]);
    expect(t.val('OFFSET(A1,-1,0)')).toBe('#REF!');
    expect(t.val('OFFSET(A1,0,0,0,1)')).toBe('#REF!');
    expect(t.val('SUM(OFFSET(B5,0,0,-2))')).toBe(70);
    expect(t.val('ROWS(OFFSET(A1,0,0,3,2))')).toBe(3);
    expect(t.val('INDIRECT("B3")')).toBe(20);
    expect(t.val('INDIRECT("R3C2",FALSE)')).toBe(20);
    expect(t.val('SUM(INDIRECT("B2:B5"))')).toBe(100);
    expect(t.val('INDIRECT("nonsense!!")')).toBe('#REF!');
    expect(t.val('INDIRECT("")')).toBe('#REF!');
  });
  it('OFFSET and INDIRECT recalc on change (volatile)', () => {
    const b = book({ A1: 1, A2: 2, B1: '=SUM(OFFSET(A1,0,0,2,1))', B2: '=INDIRECT("A"&2)*10' });
    expect(b.get('B1')).toBe(3);
    expect(b.get('B2')).toBe(20);
    b.set('A2', 5);
    expect(b.get('B1')).toBe(6);
    expect(b.get('B2')).toBe(50);
  });
  it('FORMULATEXT, HYPERLINK, IMAGE, stubs', () => {
    const b = book({ A1: 1, A2: '=A1+1' });
    expect(b.val('FORMULATEXT(A2)')).toBe('=A1+1');
    expect(b.val('FORMULATEXT(A1)')).toBe('#N/A');
    expect(b.val('HYPERLINK("https://example.com","Example")')).toBe('Example');
    expect(b.val('HYPERLINK("https://example.com")')).toBe('https://example.com');
    expect(b.val('IMAGE("https://example.com/a.png")')).toBe('https://example.com/a.png');
    expect(b.val('GETPIVOTDATA("Sales",A1)')).toBe('#REF!');
    expect(b.val('RTD("x","","t")')).toBe('#N/A');
  });
});

describe('dynamic arrays', () => {
  const t = book({
    A1: 'b', B1: 2, C1: 'x',
    A2: 'a', B2: 3, C2: 'y',
    A3: 'c', B3: 1, C3: 'x',
    A4: 'a', B4: 3, C4: 'z',
  });
  it('SEQUENCE', () => {
    expect(t.arr('SEQUENCE(3)')).toEqual([[1], [2], [3]]);
    expect(t.arr('SEQUENCE(2,3,10,-1)')).toEqual([[10, 9, 8], [7, 6, 5]]);
    expect(t.val('SEQUENCE(0)')).toBe('#CALC!');
    expect(t.val('SEQUENCE(-1)')).toBe('#VALUE!');
  });
  it('FILTER', () => {
    expect(t.arr('FILTER(A1:B4,B1:B4>2)')).toEqual([['a', 3], ['a', 3]]);
    expect(t.arr('FILTER(A1:C1,{1,0,1})')).toEqual([['b', 'x']]);
    expect(t.val('FILTER(A1:B4,B1:B4>5)')).toBe('#CALC!');
    expect(t.val('FILTER(A1:B4,B1:B4>5,"none")')).toBe('none');
    expect(t.val('FILTER(A1:B4,B1:B3>5)')).toBe('#VALUE!');
    expect(t.arr('FILTER(A1:A4,(B1:B4>1)*(C1:C4="x"))')).toEqual([['b']]);
  });
  it('SORT and SORTBY', () => {
    expect(t.arr('SORT(A1:B4)')).toEqual([['a', 3], ['a', 3], ['b', 2], ['c', 1]]);
    expect(t.arr('SORT(A1:B4,2,-1)')).toEqual([['a', 3], ['a', 3], ['b', 2], ['c', 1]]);
    expect(t.arr('SORT(A1:C4,{3,2},{1,-1})')).toEqual([['b', 2, 'x'], ['c', 1, 'x'], ['a', 3, 'y'], ['a', 3, 'z']]);
    expect(t.arr('SORT({3,1,2},,,TRUE)')).toEqual([[1, 2, 3]]);
    expect(t.arr('SORT({"b";2;TRUE;"a";1})')).toEqual([[1], [2], ['a'], ['b'], [true]]);
    expect(t.val('SORT(A1:B4,3)')).toBe('#VALUE!');
    expect(t.val('SORT(A1:B4,1,2)')).toBe('#VALUE!');
    expect(t.arr('SORTBY(A1:A4,B1:B4)')).toEqual([['c'], ['b'], ['a'], ['a']]);
    expect(t.arr('SORTBY(A1:A4,C1:C4,-1,B1:B4,1)')).toEqual([['a'], ['a'], ['c'], ['b']]);
    expect(t.arr('SORTBY({"x","y","z"},{3,1,2})')).toEqual([['y', 'z', 'x']]);
    expect(t.val('SORTBY(A1:A4,B1:B3)')).toBe('#VALUE!');
  });
  it('UNIQUE', () => {
    expect(t.arr('UNIQUE(A1:A4)')).toEqual([['b'], ['a'], ['c']]);
    expect(t.arr('UNIQUE(A1:A4,,TRUE)')).toEqual([['b'], ['c']]);
    expect(t.arr('UNIQUE(A1:B4)')).toEqual([['b', 2], ['a', 3], ['c', 1]]);
    expect(t.arr('UNIQUE({1,1,2},TRUE)')).toEqual([[1, 2]]);
    expect(t.arr('UNIQUE({"A";"a";"b"})')).toEqual([['A'], ['b']]);
    expect(t.val('UNIQUE({1;1},,TRUE)')).toBe('#CALC!');
  });
  it('TAKE, DROP, EXPAND, CHOOSEROWS, CHOOSECOLS', () => {
    expect(t.arr('TAKE(A1:C4,2)')).toEqual([['b', 2, 'x'], ['a', 3, 'y']]);
    expect(t.arr('TAKE(A1:C4,-1,-2)')).toEqual([[3, 'z']]);
    expect(t.arr('TAKE(A1:C4,,1)')).toEqual([['b'], ['a'], ['c'], ['a']]);
    expect(t.val('TAKE(A1:C4,0)')).toBe('#CALC!');
    expect(t.arr('TAKE({1,2},5)')).toEqual([[1, 2]]);
    expect(t.arr('DROP(A1:C4,3)')).toEqual([['a', 3, 'z']]);
    expect(t.arr('DROP(A1:C4,-3,2)')).toEqual([['x']]);
    expect(t.val('DROP(A1:C4,4)')).toBe('#CALC!');
    expect(t.arr('EXPAND({1,2},2,3)')).toEqual([[1, 2, '#N/A'], ['#N/A', '#N/A', '#N/A']]);
    expect(t.arr('EXPAND({1,2},2,,0)')).toEqual([[1, 2], [0, 0]]);
    expect(t.val('EXPAND({1,2;3,4},1)')).toBe('#VALUE!');
    expect(t.arr('CHOOSEROWS(A1:B4,1,-1)')).toEqual([['b', 2], ['a', 3]]);
    expect(t.arr('CHOOSEROWS(A1:B4,{2,3})')).toEqual([['a', 3], ['c', 1]]);
    expect(t.val('CHOOSEROWS(A1:B4,5)')).toBe('#VALUE!');
    expect(t.arr('CHOOSECOLS(A1:C2,3,1)')).toEqual([['x', 'b'], ['y', 'a']]);
    expect(t.val('CHOOSECOLS(A1:C2,0)')).toBe('#VALUE!');
  });
  it('HSTACK, VSTACK, TOCOL, TOROW, WRAPROWS, WRAPCOLS, TRANSPOSE', () => {
    expect(t.arr('HSTACK({1;2},{3;4;5})')).toEqual([[1, 3], [2, 4], ['#N/A', 5]]);
    expect(t.arr('VSTACK({1,2},{3})')).toEqual([[1, 2], [3, '#N/A']]);
    expect(t.arr('VSTACK(1,"a")')).toEqual([[1], ['a']]);
    expect(t.arr('TOCOL({1,2;3,4})')).toEqual([[1], [2], [3], [4]]);
    expect(t.arr('TOCOL({1,2;3,4},0,TRUE)')).toEqual([[1], [3], [2], [4]]);
    expect(t.arr('TOROW({1,2;3,4})')).toEqual([[1, 2, 3, 4]]);
    const b = book({ A1: 1, A2: null, A3: '=1/0', A4: 4 });
    expect(b.arr('TOCOL(A1:A4,1)')).toEqual([[1], ['#DIV/0!'], [4]]);
    expect(b.arr('TOCOL(A1:A4,3)')).toEqual([[1], [4]]);
    expect(b.arr('TOROW(A1:A4,2)')).toEqual([[1, null, 4]]);
    expect(t.arr('WRAPROWS({1,2,3,4,5},2)')).toEqual([[1, 2], [3, 4], [5, '#N/A']]);
    expect(t.arr('WRAPROWS({1,2,3,4,5},2,0)')).toEqual([[1, 2], [3, 4], [5, 0]]);
    expect(t.arr('WRAPCOLS({1,2,3,4,5},2)')).toEqual([[1, 3, 5], [2, 4, '#N/A']]);
    expect(t.val('WRAPCOLS({1,2;3,4},2)')).toBe('#VALUE!');
    expect(t.val('WRAPCOLS({1,2},0)')).toBe('#NUM!');
    expect(t.arr('TRANSPOSE({1,2,3})')).toEqual([[1], [2], [3]]);
    expect(t.arr('TRANSPOSE(A1:B2)')).toEqual([['b', 'a'], [2, 3]]);
  });
  it('TRIMRANGE', () => {
    const b = book({ B2: 1, C3: 2 });
    expect(b.arr('TRIMRANGE(A1:E10)')).toEqual([[1, null], [null, 2]]);
    expect(b.val('ROWS(TRIMRANGE(A1:E10))')).toBe(2);
    expect(b.val('ROWS(TRIMRANGE(A1:E10,1))')).toBe(9);
    expect(b.val('ROWS(TRIMRANGE(A:A))')).toBe('#REF!');
    expect(b.arr('TRIMRANGE({0,0;0,1})')).toEqual([[0, 0], [0, 1]]);
  });
  it('spills into the sheet', () => {
    const b = book({ A1: 3, A2: 1, A3: 2, C1: '=SORT(A1:A3)' });
    expect(b.get('C1')).toBe(1);
    expect(b.get('C3')).toBe(3);
  });
});

describe('GROUPBY and PIVOTBY', () => {
  const t = book({
    A1: 'Region', B1: 'Product', C1: 'Sales',
    A2: 'East', B2: 'X', C2: 10,
    A3: 'West', B3: 'Y', C3: 20,
    A4: 'East', B4: 'Y', C4: 5,
    A5: 'West', B5: 'X', C5: 7,
    A6: 'East', B6: 'X', C6: 3,
  });
  it('GROUPBY with an eta-reduced function and with a LAMBDA', () => {
    expect(t.arr('GROUPBY(A2:A6,C2:C6,SUM)')).toEqual([['East', 18], ['West', 27], ['Total', 45]]);
    expect(t.arr('GROUPBY(A2:A6,C2:C6,LAMBDA(x,COUNT(x)),,0)')).toEqual([['East', 3], ['West', 2]]);
    expect(t.arr('GROUPBY(A1:A6,C1:C6,SUM,3)')).toEqual([['Region', 'Sales'], ['East', 18], ['West', 27], ['Total', 45]]);
    expect(t.arr('GROUPBY(A1:A6,C1:C6,SUM)')).toEqual([['East', 18], ['West', 27], ['Total', 45]]);
    expect(t.arr('GROUPBY(A2:A6,C2:C6,SUM,,,-2)')).toEqual([['West', 27], ['East', 18], ['Total', 45]]);
    expect(t.arr('GROUPBY(A2:A6,C2:C6,SUM,,0,-1)')).toEqual([['West', 27], ['East', 18]]);
    expect(t.arr('GROUPBY(A2:A6,C2:C6,SUM,,0,,B2:B6="X")')).toEqual([['East', 13], ['West', 7]]);
  });
  it('GROUPBY with two fields and subtotals', () => {
    expect(t.arr('GROUPBY(A2:B6,C2:C6,SUM,0,2)')).toEqual([
      ['East', 'X', 13], ['East', 'Y', 5], ['East', null, 18],
      ['West', 'X', 7], ['West', 'Y', 20], ['West', null, 27],
      ['Total', null, 45],
    ]);
  });
  it('PIVOTBY', () => {
    expect(t.arr('PIVOTBY(A2:A6,B2:B6,C2:C6,SUM)')).toEqual([
      [null, 'X', 'Y', 'Total'],
      ['East', 13, 5, 18],
      ['West', 7, 20, 27],
      ['Total', 20, 25, 45],
    ]);
    expect(t.arr('PIVOTBY(A2:A6,B2:B6,C2:C6,SUM,,0,,0)')).toEqual([
      [null, 'X', 'Y'],
      ['East', 13, 5],
      ['West', 7, 20],
    ]);
  });
  it('rejects a non-function aggregator', () => {
    expect(t.val('GROUPBY(A2:A6,C2:C6,5)')).toBe('#VALUE!');
    expect(t.val('GROUPBY(A2:A6,C2:C6,NOSUCHFN)')).toBe('#NAME?');
  });
});
