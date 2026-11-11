// A workbook file is untrusted: nothing in it may reach markup as anything but text or a number.
import { describe, it, expect } from 'vitest';
import { Workbook } from '../../src/model/workbook';
import { previewCss, styleCss } from '../../src/commands/conditional';
import { exportHtml, importHtml } from '../../src/io/html';
import { safeImageSrc } from '../../src/util/sanitize';

const EVIL = '1px"><img src=x onerror=alert(1)>';

function hostile(): Workbook {
  const j: any = new Workbook().toJSON();
  const s = j.sheets[0];
  s.colW = [[0, EVIL], [1, 90]];
  s.rowH = [[0, EVIL], [2, 30]];
  s.defaultColW = EVIL; s.defaultRowH = EVIL; s.zoom = EVIL;
  s.drawings = [{ id: 'd1', kind: 'ink', svg: '<svg></svg>', rotate: EVIL, at: { r: 0, c: 0, dx: EVIL, dy: 0, w: EVIL, h: EVIL } }];
  s.charts = [{ id: 'c1', at: { r: EVIL, c: 0, dx: 0, dy: 0, w: 300, h: EVIL } }];
  return Workbook.fromJSON(j);
}

describe('loading an untrusted workbook', () => {
  it('keeps only numeric sizes and positions', () => {
    const s = hostile().sheets[0];
    expect(s.colW.has(0)).toBe(false);
    expect(s.colW.get(1)).toBe(90);
    expect(s.rowH.has(0)).toBe(false);
    expect(s.rowH.get(2)).toBe(30);
    for (const v of [s.defaultColW, s.defaultRowH, s.zoom, s.colWidth(0), s.rowHeight(0)]) expect(typeof v).toBe('number');
    for (const o of [...s.drawings, ...s.charts]) for (const v of Object.values(o.at)) expect(typeof v).toBe('number');
    expect(typeof s.drawings[0].rotate).toBe('number');
    expect(s.charts[0].at.w).toBe(300);
  });

  it('round-trips ordinary sizes unchanged', () => {
    const b = new Workbook();
    b.sheet.colW.set(3, 120.5); b.sheet.rowH.set(4, 33); b.sheet.zoom = 1.5;
    const s = Workbook.fromJSON(b.toJSON()).sheets[0];
    expect(s.colW.get(3)).toBe(120.5);
    expect(s.rowH.get(4)).toBe(33);
    expect(s.zoom).toBe(1.5);
  });
});

describe('markup built from document data', () => {
  it('does not let a conditional-format colour leave its declaration', () => {
    const bad = 'red"><img src=x onerror=alert(1)>';
    for (const css of [
      styleCss({ fill: bad, color: bad, bt: { style: 'thin', color: bad } } as any),
      previewCss({ type: 'colorScale', colors: [bad, '#ffffff'] } as any),
      previewCss({ type: 'dataBar', barColor: bad } as any),
    ]) expect(css).not.toMatch(/[<>"]/);
    expect(styleCss({ fill: '#ff0000', color: 'var(--ql-accent1)' } as any)).toBe('background:#ff0000;color:var(--ql-accent1)');
  });

  it('drops script links however the scheme is disguised', async () => {
    const b = new Workbook();
    b.transact('t', () => {
      b.sheet.set(0, 0, { v: 'a', link: 'java\tscript:alert(1)' });
      b.sheet.set(1, 0, { v: 'b', link: 'https://example.com/' });
    });
    const html = await exportHtml(b).text();
    expect(html).not.toMatch(/script:/i);
    expect(html).toContain('href="https://example.com/"');
    const back = importHtml('<table><tr><td><a href=" JaVa&#9;Script:alert(1)">x</a></td><td><a href="https://example.com/">y</a></td></tr></table>');
    expect(back.sheet.get(0, 0)?.link).toBeUndefined();
    expect(back.sheet.get(0, 1)?.link).toBe('https://example.com/');
  });

  it('accepts embedded pictures only', () => {
    expect(safeImageSrc('data:image/png;base64,AAAA')).toBe('data:image/png;base64,AAAA');
    expect(safeImageSrc('blob:https://app/1')).toBe('blob:https://app/1');
    expect(safeImageSrc('https://example.invalid/beacon.png')).toBe('');
    expect(safeImageSrc('javascript:alert(1)')).toBe('');
    expect(safeImageSrc({} as any)).toBe('');
  });
});
