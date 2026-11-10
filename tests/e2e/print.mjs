// Print lifetime test: on a phone browser print() returns at once and the print UI keeps laying
// the live page out for its preview, so the print pages must stay until the user is back in the app.
// Usage: npm run build && npm run e2e:print   (set QL_URL to test an already-running server instead)
import { spawn } from 'node:child_process';
import { existsSync } from 'node:fs';
import { chromium } from 'playwright';

const PORT = 4182;
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
  const ctx = await browser.newContext({ viewport: { width: 412, height: 915 }, hasTouch: true, isMobile: true, userAgent: 'Mozilla/5.0 (Linux; Android 14; Pixel 8) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Mobile Safari/537.36' });
  const p = await ctx.newPage();
  await p.addInitScript(() => {
    localStorage.setItem('quillgrid.prefs.v1', JSON.stringify({ firstRun: false }));
    // A phone browser: print() opens the print UI and returns immediately, and afterprint comes early.
    window.print = () => { window.__printed = (window.__printed || 0) + 1; setTimeout(() => window.dispatchEvent(new Event('afterprint')), 50); };
  });
  await p.goto(url, { waitUntil: 'networkidle' });
  await p.waitForSelector('.ql-grid-scroller');
  await p.waitForTimeout(500);
  await p.evaluate(() => { ql.sheet.set(0, 0, { v: 'Printed cell' }); });
  p.on('pageerror', (e) => failures.push('pageerror: ' + e.message));
  await p.evaluate(() => { window.__printErr = null; Promise.resolve(ql.exec('file.print')).catch((e) => { window.__printErr = String(e); }); });
  await p.waitForFunction(() => window.__printed === 1, null, { timeout: 10000 });
  // Well past every timer the old code used (800 ms): the print pages are still what would print.
  await p.waitForTimeout(3000);
  const state = () => p.evaluate(() => ({ pages: document.querySelectorAll('#ql-print-root .ql-print-page').length > 0, printing: document.body.classList.contains('ql-printing') }));
  check('pages kept while the print UI is open', await state(), { pages: true, printing: true });
  await p.emulateMedia({ media: 'print' });
  check("print layout shows only the pages", await p.evaluate(() => { const root = document.getElementById("ql-print-root"); return { app: [...document.body.children].filter((e) => e !== root && getComputedStyle(e).display !== "none").length, root: root ? getComputedStyle(root).display : "missing" }; }), { app: 0, root: "block" });
  await p.emulateMedia({ media: 'screen' });
  // Back in the app: the first touch ends the print session and the pages go.
  await p.touchscreen.tap(200, 500);
  await p.waitForTimeout(800);
  check('pages removed afterwards', await state(), { pages: false, printing: false });
} finally {
  await browser.close();
  if (server) { try { process.kill(-server.pid); } catch { /* */ } }
}
if (failures.length) { console.error('FAIL\n' + failures.join('\n')); process.exit(1); }
console.log('print e2e: ok');
