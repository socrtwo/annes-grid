import { describe, it, expect } from 'vitest';
import { Workbook } from '../../src/model/workbook';
import { CellError } from '../../src/model/types';
import { layoutPages, computePageBreaks, fitScale, parseTitleSpan, expandHeaderFooter, parseHeaderFooter, headerFooterFor, printAreas, paperInches } from '../../src/print/pages';
import { renderCellsTable, renderPages, planPrint } from '../../src/print/print';

function sheetWith(rows: number, cols: number) {
  const b = new Workbook();
  const s = b.sheet;
  b.transact('fill', () => { for (let r = 0; r < rows; r++) for (let c = 0; c < cols; c++) s.set(r, c, { v: r * cols + c }); });
  return { b, s };
}

const ctx = { page: 3, pages: 7, date: new Date(2026, 0, 2, 13, 5), file: 'Budget.xlsx', path: '/docs/Budget.xlsx', sheet: 'Q1' };

describe('page breaks', () => {
  it('breaks rows and columns by the printable area', () => {
    const { s } = sheetWith(100, 20);
    // Letter portrait, 0.75" top/bottom: (11 - 1.5) * 96 = 912px tall -> 45 rows of 20px; (8.5 - 1.4) * 96 = 681.6px -> 10 columns of 64px.
    const l = layoutPages(s);
    expect(l.rowBreaks[0]).toBe(45);
    expect(l.colBreaks[0]).toBe(10);
    expect(l.pages.length).toBe(3 * 2);
    // Down, then over: the first pages go down the first columns.
    expect(l.pages[0]).toMatchObject({ r1: 0, c1: 0 });
    expect(l.pages[1]).toMatchObject({ r1: 45, c1: 0 });
    s.print.order = 'overThenDown';
    expect(layoutPages(s).pages[1]).toMatchObject({ r1: 0, c1: 10 });
  });

  it('honors manual breaks, and ignores them when fitting to pages', () => {
    const { s } = sheetWith(30, 3);
    s.rowBreaks = [10, 20];
    expect(computePageBreaks(s).rows).toEqual([10, 20]);
    s.print.fitWide = 1; s.print.fitTall = 1;
    expect(computePageBreaks(s).rows).toEqual([]);
    expect(layoutPages(s).pages.length).toBe(1);
  });

  it('fit to one page scales down and puts everything on one page', () => {
    const { s } = sheetWith(200, 30);
    s.print.fitWide = 1; s.print.fitTall = 1;
    const l = layoutPages(s);
    expect(l.scale).toBeLessThan(100);
    expect(l.scale).toBe(fitScale(s, l.areas));
    expect(l.pages.length).toBe(1);
    // Fit all columns: one page wide, as many tall as needed.
    s.print.fitTall = undefined;
    const l2 = layoutPages(s);
    expect(l2.colBreaks).toEqual([]);
    expect(l2.rowBreaks.length).toBeGreaterThan(0);
  });

  it('repeated title rows take room on later pages', () => {
    const { s } = sheetWith(100, 2);
    const plain = layoutPages(s).rowBreaks[1];
    s.print.titleRows = '$1:$2';
    const titled = layoutPages(s).rowBreaks[1];
    expect(titled - layoutPages(s).rowBreaks[0]).toBe(43);
    expect(titled).toBeLessThan(plain);
  });

  it('print area parts, landscape and titles parsing', () => {
    const { s } = sheetWith(10, 10);
    s.print.area = 'A1:B3,$D$4:$E$5';
    expect(printAreas(s)).toEqual([{ r1: 0, c1: 0, r2: 2, c2: 1 }, { r1: 3, c1: 3, r2: 4, c2: 4 }]);
    expect(layoutPages(s).pages.length).toBe(2);
    s.print.orientation = 'landscape';
    expect(paperInches(s.print)).toEqual([11, 8.5]);
    expect(parseTitleSpan('$1:$3', 'row')).toEqual([0, 2]);
    expect(parseTitleSpan("'Sheet 1'!$B:$C", 'col')).toEqual([1, 2]);
    expect(parseTitleSpan('A1', 'row')).toBeNull();
  });

  it('skips pages whose rows are all hidden', () => {
    const { s } = sheetWith(60, 1);
    for (let r = 45; r < 60; r++) s.hiddenRows.add(r);
    expect(layoutPages(s).pages.length).toBe(1);
  });
});

describe('header and footer codes', () => {
  it('expands readable and file codes', () => {
    expect(expandHeaderFooter('Page &[Page] of &[Pages]', ctx)).toBe('Page 3 of 7');
    expect(expandHeaderFooter('&P/&N &A &F &Z', ctx)).toBe('3/7 Q1 Budget.xlsx /docs/Budget.xlsx');
    expect(expandHeaderFooter('&[Tab] - &[File]', ctx)).toBe('Q1 - Budget.xlsx');
    expect(expandHeaderFooter('&P+2 &P-1', ctx)).toBe('5 2');
    expect(expandHeaderFooter('A && B', ctx)).toBe('A & B');
    expect(expandHeaderFooter('&D', ctx)).toBe(ctx.date.toLocaleDateString());
    expect(expandHeaderFooter('&[Time]', ctx)).toMatch(/1:05|13:05/);
  });
  it('parses formatting codes into runs', () => {
    const runs = parseHeaderFooter('&"Lato,Bold Italic"&14Title&B plain&KFF0000 red', ctx);
    expect(runs[0]).toMatchObject({ text: 'Title', font: 'Lato', bold: true, italic: true, size: 14 });
    expect(runs[1]).toMatchObject({ text: ' plain', bold: false });
    expect(runs[2]).toMatchObject({ text: ' red', color: '#ff0000' });
  });
  it('first and even page headers', () => {
    const { s } = sheetWith(1, 1);
    const p = s.print;
    p.differentFirst = true; p.firstHeader = { left: 'first', center: '', right: '' };
    p.differentOddEven = true; p.evenHeader = { left: 'even', center: '', right: '' };
    p.header = { left: 'odd', center: '', right: '' };
    expect(headerFooterFor(p, 0, 1).header.left).toBe('first');
    expect(headerFooterFor(p, 1, 2).header.left).toBe('even');
    expect(headerFooterFor(p, 2, 3).header.left).toBe('odd');
  });
});

describe('HTML pages', () => {
  it('renders merges, borders, number formats and alignment', () => {
    const b = new Workbook();
    const s = b.sheet;
    b.transact('t', () => {
      s.set(0, 0, { v: 'Title', s: b.styleId({ bold: true, bb: { style: 'thin', color: '#ff0000' } }) });
      s.set(1, 0, { v: 1234.5, s: b.styleId({ numFmt: '#,##0.00', fill: '#ffff00' }) });
      s.set(1, 1, { v: 'x<y' });
      s.merges.push({ r1: 0, c1: 0, r2: 0, c2: 2 });
    });
    const html = renderCellsTable({ book: b, engine: {} } as any, s, [0, 1], [0, 1, 2], { bw: false, draft: false, grid: true, errors: undefined, headings: true });
    const host = document.createElement('div');
    host.innerHTML = html;
    const tds = host.querySelectorAll('tr')[1].querySelectorAll('td');
    expect(tds.length).toBe(1);
    expect(tds[0].getAttribute('colspan')).toBe('3');
    expect(tds[0].getAttribute('style')).toContain('border-bottom:1px solid #ff0000');
    expect(tds[0].textContent).toBe('Title');
    expect(tds[0].innerHTML).toContain('font-weight:700');
    const row2 = host.querySelectorAll('tr')[2];
    expect(row2.querySelector('th')!.textContent).toBe('2');
    const cells = row2.querySelectorAll('td');
    expect(cells[0].textContent).toBe('1,234.50');
    expect(cells[0].getAttribute('style')).toContain('background:#ffff00');
    expect(cells[0].innerHTML).toContain('text-align:right');
    expect(cells[1].innerHTML).toContain('x&lt;y');
    expect(host.querySelectorAll('col').length).toBe(4);
    expect(host.querySelector('tr')!.textContent).toBe('ABC');
  });

  it('renders pages with header/footer, page numbers and error display options', () => {
    const { b, s } = sheetWith(60, 2);
    s.print.footer = { left: '', center: 'Page &[Page] of &[Pages]', right: '&[Tab]' };
    s.print.errors = 'dashes';
    b.transact('e', () => s.set(0, 0, { v: new CellError('#DIV/0!') }));
    const tab = { book: b, engine: {}, fileName: 'x.xlsx', title: 'x' } as any;
    const pages = renderPages({ tab, file: 'x.xlsx', path: 'x.xlsx' });
    expect(pages.length).toBe(2);
    expect(pages[1].querySelector('.ql-pp-footer')!.textContent).toBe('Page 2 of 2Sheet1');
    expect(pages[0].dataset.size).toBe('8.5x11');
    expect(pages[0].querySelector('td')!.textContent).toBe('--');
    // Comments at the end add a page.
    b.transact('c', () => { b.recordSheet(s); s.comments.set(0, { author: 'Ann', text: 'Check', time: 0 }); });
    s.print.comments = 'atEnd';
    expect(planPrint(tab).length).toBe(3);
    expect(renderPages({ tab, file: 'x', path: 'x' })[2].textContent).toContain('Ann');
  });

  it('prints the selection only', () => {
    const { b, s } = sheetWith(100, 20);
    s.sel = { r: 0, c: 0, ranges: [{ r1: 0, c1: 0, r2: 4, c2: 1 }] };
    const tab = { book: b, engine: {} } as any;
    const jobs = planPrint(tab, { what: 'selection' });
    expect(jobs.length).toBe(1);
    expect(jobs[0].spec).toMatchObject({ r1: 0, r2: 4, c1: 0, c2: 1 });
  });
});

describe('fit to a number of pages', () => {
  it('never runs over the requested page count', () => {
    for (const [rows, cols, wide, tall] of [[99, 3, 1, 2], [5, 21, 2, 0], [200, 3, 0, 3], [130, 14, 1, 3], [61, 25, 2, 2]] as const) {
      const { s } = sheetWith(rows, cols);
      s.print.fitWide = wide; s.print.fitTall = tall;
      const l = layoutPages(s);
      const acrossPages = new Set(l.pages.map((p) => p.c1)).size, downPages = new Set(l.pages.map((p) => p.r1)).size;
      if (wide) expect(acrossPages, `${rows}x${cols} wide`).toBeLessThanOrEqual(wide);
      if (tall) expect(downPages, `${rows}x${cols} tall`).toBeLessThanOrEqual(tall);
    }
  });
  it('keeps to the page count with repeated title rows', () => {
    const { s } = sheetWith(200, 3);
    s.print.titleRows = '$1:$3'; s.print.fitTall = 3;
    expect(new Set(layoutPages(s).pages.map((p) => p.r1)).size).toBeLessThanOrEqual(3);
  });
});
