// Touch selection test: real touch gestures (CDP touch events, so the browser's own panning and
// pointercancel behaviour apply) on a phone-sized viewport.
// Usage: npm run build && npm run e2e:touch   (set QL_URL to test an already-running server instead)
import { spawn } from 'node:child_process';
import { existsSync } from 'node:fs';
import { chromium } from 'playwright';

const PORT = 4181;
let server = null;
let url = process.env.QL_URL;
if (!url) {
  server = spawn('npx', ['vite', 'preview', '--port', String(PORT), '--strictPort'], { stdio: 'ignore', detached: true });
  url = `http://localhost:${PORT}/`;
  for (let i = 0; i < 60; i++) {
    try { if ((await fetch(url)).ok) break; } catch { /* not up yet */ }
    await new Promise((r) => setTimeout(r, 500));
  }
}

const exe = existsSync('/opt/pw-browsers/chromium') ? { executablePath: '/opt/pw-browsers/chromium' } : {};
const browser = await chromium.launch(exe);
const failures = [];
const check = (name, got, want) => { if (JSON.stringify(got) !== JSON.stringify(want)) failures.push(`${name}: got ${JSON.stringify(got)}, want ${JSON.stringify(want)}`); };

try {
  const ctx = await browser.newContext({ viewport: { width: 412, height: 915 }, hasTouch: true, isMobile: true, deviceScaleFactor: 2 });
  const p = await ctx.newPage();
  const errs = [];
  p.on('pageerror', (e) => errs.push('pageerror: ' + e.message));
  await p.addInitScript(() => localStorage.setItem('quillgrid.prefs.v1', JSON.stringify({ firstRun: false })));
  await p.goto(url, { waitUntil: 'networkidle' });
  await p.waitForSelector('.ql-grid-scroller');
  await p.waitForTimeout(500);
  const cdp = await ctx.newCDPSession(p);
  const wait = (ms) => p.waitForTimeout(ms);
  const touch = (type, pts) => cdp.send('Input.dispatchTouchEvent', { type, touchPoints: pts.map(([x, y]) => ({ x, y })) });
  /** Page coordinates of a point in a cell: fx, fy are fractions of the cell (0,0 top-left; 1,1 bottom-right). */
  const at = (r, c, fx = 0.5, fy = 0.5) => p.evaluate(([r, c, fx, fy]) => { const g = ql.grid, b = g.el.getBoundingClientRect(), rc = g.cellRect(r, c); return [b.left + rc.x + rc.w * fx, b.top + rc.y + rc.h * fy]; }, [r, c, fx, fy]);
  const sel = () => p.evaluate(() => { const s = ql.sheet.sel, a = s.ranges[s.ranges.length - 1]; return [a.r1, a.c1, a.r2, a.c2]; });
  const scroll = () => p.evaluate(() => [ql.grid.scroller.scrollLeft, ql.grid.scroller.scrollTop]);
  const tap = async (pt) => { await touch('touchStart', [pt]); await wait(40); await touch('touchEnd', []); await wait(80); };
  /** Finger down at `from`, optional hold, move to `to` in steps, lift. */
  const drag = async (from, to, { hold = 0, steps = 12 } = {}) => {
    await touch('touchStart', [from]);
    await wait(hold || 30);
    for (let i = 1; i <= steps; i++) { await touch('touchMove', [[from[0] + ((to[0] - from[0]) * i) / steps, from[1] + ((to[1] - from[1]) * i) / steps]]); await wait(16); }
    await touch('touchEnd', []);
    await wait(120);
  };
  const menuOpen = () => p.evaluate(() => !!document.querySelector('.ql-menu'));
  const closeMenu = async () => { await p.keyboard.press('Escape'); await wait(80); };

  // Tap selects a cell.
  await tap(await at(2, 1));
  check('tap selects', await sel(), [2, 1, 2, 1]);

  // The bottom-right handle extends the selection (it used to start a fill, and the browser's
  // panning cancelled the drag).
  await drag(await at(2, 1, 1, 1), await at(7, 3));
  check('end handle extends', await sel(), [2, 1, 7, 3]);
  check('end handle does not scroll', await scroll(), [0, 0]);
  check('end handle does not fill', await p.evaluate(() => ql.sheet.value(5, 2)), null);

  // The top-left handle moves the other corner.
  await drag(await at(2, 1, 0, 0), await at(1, 0));
  check('start handle extends', await sel(), [1, 0, 7, 3]);

  // A handle that is touched without dragging is a tap on the cell underneath.
  await tap(await at(8, 4, 0.2, 0.3));
  check('tap near a handle selects the cell', await sel(), [8, 4, 8, 4]);

  // Press and hold, then drag: selects a range from the pressed cell, with no context menu.
  await drag(await at(4, 1), await at(9, 3), { hold: 700 });
  check('hold and drag selects', await sel(), [4, 1, 9, 3]);
  check('hold and drag does not scroll', await scroll(), [0, 0]);
  check('hold and drag shows no menu', await menuOpen(), false);

  // Press and hold a row heading, then drag: selects whole rows.
  const hdr = await p.evaluate(() => { const g = ql.grid, b = g.el.getBoundingClientRect(); return [b.left + g.headerW / 2, b.top]; });
  const rowY = async (r) => (await at(r, 0))[1];
  await drag([hdr[0], await rowY(3)], [hdr[0], await rowY(6)], { hold: 700 });
  check('hold and drag on row headings', await p.evaluate(() => ql.grid.wholeRowsSelected()), [3, 4, 5, 6]);

  // Press and hold, then lift: the context menu.
  const c = await at(5, 2);
  await touch('touchStart', [c]); await wait(700); await touch('touchEnd', []); await wait(150);
  check('hold and lift opens the menu', await menuOpen(), true);
  await closeMenu();

  // Press and hold the bottom-right handle, then drag: fills.
  await p.evaluate(() => { ql.sheet.set(0, 5, { v: 7 }); ql.grid.select(0, 5); });
  await wait(100);
  await drag(await at(0, 5, 1, 1), await at(3, 5), { hold: 650 });
  await wait(200);
  check('hold handle and drag fills', await p.evaluate(() => [ql.sheet.value(1, 5), ql.sheet.value(3, 5)]), [7, 7]);
  check('hold handle and drag shows no menu', await menuOpen(), false);

  // A plain drag still pans the sheet and leaves the selection alone.
  await tap(await at(2, 1));
  await drag(await at(20, 2), await at(8, 2), { steps: 20 });
  await wait(600);
  check('drag pans', (await scroll())[1] > 50, true);
  check('drag keeps the selection', await sel(), [2, 1, 2, 1]);

  check('page errors', errs, []);
} finally {
  await browser.close();
  if (server) { try { process.kill(-server.pid); } catch { /* */ } }
}
if (failures.length) { console.error('FAIL\n' + failures.join('\n')); process.exit(1); }
console.log('touch e2e: ok');
