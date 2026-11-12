/** Anne's Grid entry point. */
import './util/legacy'; // first: carries data over from the app's former name
import 'material-symbols/outlined.css';
import './styles/app.css';
import './ui/scrollcue.css';
import { installFonts, onFontsReady } from './fonts';
import './calc/functions';
import './commands';
import { App } from './app';
import { Shell } from './ui/shell';
import { launchFile, onFileOpened, onBackButton, isNative, isDesktop, isCapacitor, type OpenedFile } from './platform/native';
import { closeMenus, inDialog, toast } from './ui/dialog';
import { kvKeys } from './util/store';
import { allCommands, getCommand } from './commands/registry';
import { installScrollCues } from './ui/scrollcue';

// A second line of defence for the web app: should markup from a file ever reach the page unescaped, script in it
// (inline handlers, injected <script>, javascript: links) does not run. The desktop app sends the same policy as a
// response header; the Android and iOS apps are left as they are (their WebView bridge has not been tested with it).
if (!isNative() && !isCapacitor() && !import.meta.env.DEV && /^https?:$/.test(location.protocol)) {
  const meta = document.createElement('meta');
  meta.httpEquiv = 'Content-Security-Policy';
  meta.content = [
    "default-src 'self'",
    "script-src 'self' 'unsafe-eval' 'wasm-unsafe-eval' blob: https://cdn.jsdelivr.net",
    "worker-src 'self' blob:",
    "connect-src 'self' https: http://localhost:* http://127.0.0.1:* data: blob:",
    "img-src 'self' data: blob: https:",
    "media-src 'self' data: blob:",
    "style-src 'self' 'unsafe-inline'",
    "font-src 'self' data:",
    "object-src 'none'",
    "base-uri 'self'",
  ].join('; ');
  document.head.appendChild(meta);
}

installFonts();

// Scroll cues (thin bars, edge shadows and a "More" chip) on every scroll container that can hold choices:
// menus, dialog bodies and their list boxes, panes, Backstage, the command palette, the cell editor's
// autocomplete, galleries. Not on the grid itself, the ribbon or the other horizontal strips.
installScrollCues({
  selector: [
    '.ql-menu', '.ql-dialog-body', '.ql-cmdlist', '.ql-listbox', '.ql-pane-body', '.ql-bs-nav', '.ql-bs-body',
    '.ql-palette-list', '.ql-ac-popup', '.ql-symbol-grid', '.ql-notices', '.ql-dialog-preview', '.ql-ai-out',
    '.ql-spell-context', '.ql-dict-list', '.ql-pp-side', '.ql-pp-view', '.ql-fx-args', '.ql-namemgr-list',
    '.ql-cfmgr-list', '.ql-pvt-fields', '.ql-pvt-rec-list', '.ql-pvt-rec-preview', '.ql-objgal', '.ql-ai-log',
    '.ql-ai-diff-wrap', '.ql-ai-chips', '.ql-ai-result', '.ql-macro-mods', '.ql-macro-out', '.ql-data-checklist',
    '.ql-data-preview', '.ql-slash-items', '.ql-depgraph', '.ql-ai-delib-text',
  ].join(', '),
});

const host = document.getElementById('app')!;
const app = new App(host);
app.ui = new Shell(app);
app.rebuildKeyIndex();
app.createTab();
(window as any).ql = app;
(window as any).__qlAllIds = allCommands().map((c) => c.id);
(window as any).__qlGetCommand = getCommand;
(window as any).__qlIO = () => import('./io');
(window as any).__qlAi = async () => ({ ...(await import('./ai/client')), ...(await import('./ai/store')), ...(await import('./ai/providers')) });

// Canvas text needs the web fonts: preload the default ones and repaint when they arrive.
try { void (document as any).fonts?.load('14.67px Carlito'); void (document as any).fonts?.load('16px "Material Symbols Outlined"'); } catch { /* */ }
onFontsReady(() => { for (const d of app.docs) { d.grid.cache.text.clear(); d.grid.invalidate(); } });

async function start() {
  if (!isNative() && /[?&]ql-oauth=/.test(location.search)) {
    try {
      const { handleOAuthRedirect } = await import('./ai/oauth');
      const { setApiKey, saveAiSettings } = await import('./ai/store');
      const ok = await handleOAuthRedirect(async (key) => {
        await setApiKey('openrouter', key);
        saveAiSettings((s) => {
          s.providers.openrouter = s.providers.openrouter || { id: 'openrouter', model: 'openrouter/free', fastModel: 'openrouter/free' };
          if (!s.active) s.active = 'openrouter';
        });
      });
      if (ok) toast('Signed in to OpenRouter and saved the key. You can close this tab and return to Anne\'s Grid.', { ms: 15000 });
    } catch (e: any) {
      toast(`OpenRouter sign-in did not complete: ${e?.message || e}`, { ms: 10000 });
    }
  }
  const { openBuffer } = await import('./commands/file');
  // Listen first, so a file opened while the launch file loads is queued rather than lost.
  let ready = false;
  const queued: OpenedFile[] = [];
  onFileOpened((f) => { if (ready) void openBuffer(app, f.data, f.name, f.uri); else queued.push(f); });
  const lf = await launchFile();
  if (lf) await openBuffer(app, lf.data, lf.name, lf.uri);
  const openedAny = !!lf || queued.length > 0;
  for (const f of queued.splice(0)) await openBuffer(app, f.data, f.name, f.uri);
  ready = true;
  const drafts = await kvKeys('draft:');
  if (drafts.length) toast(`${drafts.length} unsaved workbook(s) can be recovered.`, { action: 'Recover', onAction: () => app.exec('file.recover'), ms: 8000 });
  if (app.prefs.firstRun && !openedAny) setTimeout(() => app.ui.tour(), 600);
  app.focus();
}
start();

if (!isNative() && !isCapacitor() && 'serviceWorker' in navigator && /^https?:$/.test(location.protocol) && !import.meta.env.DEV) {
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('./sw.js').catch(() => undefined);
    // The files of this first load were fetched before the worker existed: hand it the list so they are cached too.
    navigator.serviceWorker.ready.then((reg) => {
      const urls = [...new Set(performance.getEntriesByType('resource').map((x) => x.name).filter((u) => u.startsWith(location.origin)))];
      reg.active?.postMessage({ type: 'precache', urls });
    }).catch(() => undefined);
  });
}

onBackButton(() => {
  if (document.querySelector('.ql-menu')) { closeMenus(); return true; }
  if (inDialog()) { document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' })); return true; }
  if (host.classList.contains('ql-backstage-open')) { app.ui.hideBackstage(); return true; }
  if (app.ui.pane) { app.ui.closePane(); return true; }
  if (app.grid.editing) { app.grid.cancelEdit(); return true; }
  if (document.body.classList.contains('ql-fullscreen')) { document.body.classList.remove('ql-fullscreen'); return true; }
  void app.saveDrafts();
  return false;
});

window.addEventListener('beforeunload', (e) => {
  void app.saveDrafts();
  if ((!isNative() || isDesktop()) && app.docs.some((d) => d.dirty)) { e.preventDefault(); e.returnValue = ''; }
});
document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'hidden') void app.saveDrafts(); });
window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', () => app.ui.applyPrefs());
