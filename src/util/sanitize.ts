/**
 * HTML and URL sanitizing for document-controlled content.
 *
 * Documents come from anywhere (DOCX/HTML/EPUB/.qgw imports, the clipboard), and on Android
 * the WebView origin can reach the native bridge, so any string from a document that ends up
 * in innerHTML or navigation must pass through here.
 */

/** Schemes Anne's Grid will hand to the system browser or mail/phone apps. */
const SAFE_SCHEMES = /^(https?|mailto|tel|ftp|sms|geo):/i;

export function isSafeExternalUrl(url: string): boolean {
  const u = url.trim();
  return SAFE_SCHEMES.test(u);
}

/** True for URLs that must never be stored or followed (script-capable schemes). */
export function isScriptUrl(url: string): boolean {
  // Browsers ignore whitespace and control characters inside the scheme.
  const u = url.replace(/[\u0000- \u007f-\u009f]+/g, '').toLowerCase();
  return /^(javascript|vbscript|livescript):/.test(u) || /^data:(text\/html|application\/xhtml|image\/svg)/.test(u);
}

// Form controls stay (task-list checkboxes): without on* attributes they cannot run script.
const DANGEROUS_TAGS = 'script,iframe,frame,frameset,object,embed,applet,base,link,meta,noscript,template,portal';
const URL_ATTRS = new Set(['href', 'src', 'xlink:href', 'action', 'formaction', 'poster', 'background', 'lowsrc', 'dynsrc', 'data', 'cite', 'longdesc']);

/**
 * Remove script vectors from a DOM subtree in place: event-handler attributes, script-capable
 * URLs, srcdoc, and elements that execute or load active content.
 */
export function stripActiveContent(root: Element | DocumentFragment): void {
  root.querySelectorAll(DANGEROUS_TAGS).forEach((el) => {
    // Keep MathJax sources (handled by the HTML importer before this runs).
    if (el.tagName === 'SCRIPT' && /^math\/tex/.test(el.getAttribute('type') || '')) return;
    el.remove();
  });
  const all = root instanceof Element ? [root, ...root.querySelectorAll('*')] : [...root.querySelectorAll('*')];
  for (const el of all) {
    for (const a of [...el.attributes]) {
      const n = a.name.toLowerCase();
      if (n.startsWith('on') || n === 'srcdoc' || n === 'formaction') el.removeAttribute(a.name);
      else if (URL_ATTRS.has(n) && isScriptUrl(a.value)) el.removeAttribute(a.name);
      else if (n === 'style' && /expression\s*\(|javascript:|behavior\s*:|-moz-binding/i.test(a.value)) el.removeAttribute(a.name);
    }
  }
}

const GEN_TAGS = new Set(['P', 'SPAN', 'A', 'B', 'I', 'EM', 'STRONG', 'U', 'SUP', 'SUB', 'BR', 'DIV', 'SMALL']);
const SAFE_STYLE = /^[\w\s:;.,%#()'"-]*$/;

/**
 * Allow-list sanitizer for generated-block caches (TOC, index, bibliography...): only the
 * tags and attributes Anne's Grid itself generates survive.
 */
export function sanitizeGeneratedHtml(html: string): string {
  if (!html) return '';
  const t = document.createElement('template');
  t.innerHTML = html;
  const walk = (parent: ParentNode) => {
    for (const node of [...parent.childNodes]) {
      if (node.nodeType === 3) continue;
      if (node.nodeType !== 1) { node.remove(); continue; }
      const el = node as Element;
      walk(el);
      if (!GEN_TAGS.has(el.tagName)) { el.replaceWith(...el.childNodes); continue; }
      for (const a of [...el.attributes]) {
        const keep = a.name === 'class' || (a.name === 'href' && a.value.startsWith('#')) || (a.name === 'style' && SAFE_STYLE.test(a.value) && !/url\s*\(|expression/i.test(a.value));
        if (!keep) el.removeAttribute(a.name);
      }
    }
  };
  walk(t.content);
  return t.innerHTML;
}

/** Escape text for HTML element content and quoted attribute values. */
export function escHtml(s: unknown): string {
  return String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

/** A CSS color from document data, or the fallback when it is not a plain color value. */
export function safeColor(c: unknown, fallback = '#000000'): string {
  const s = String(c ?? '');
  return /^(#[0-9a-f]{3,8}|rgba?\([\d\s.,%]+\)|hsla?\([\d\s.,%deg]+\)|[a-z]{3,20})$/i.test(s) ? s : fallback;
}

/**
 * An image source from document data, or '' when it is not embedded data. Remote URLs are refused
 * so that opening or printing a workbook cannot be used to call out to a server.
 */
export function safeImageSrc(src: unknown): string {
  if (typeof src !== 'string' || !src) return '';
  return /^(data:image\/(png|jpe?g|gif|webp|bmp|svg\+xml)[;,]|blob:)/i.test(src) ? src : '';
}

/** A finite number from document data, clamped. */
export function safeNum(v: unknown, fallback: number, min = -1e6, max = 1e6): number {
  const n = Number(v);
  return Number.isFinite(n) ? Math.min(max, Math.max(min, n)) : fallback;
}

/**
 * One CSS property value from document data ("1px solid #444", "35%", "var(--ql-accent1)"), or
 * null. No ; : quotes, slashes or braces, so it cannot add declarations or rules, and only
 * var/calc/color functions, so it cannot load anything.
 */
export function safeCssValue(v: unknown, max = 120): string | null {
  if (typeof v !== 'string' && typeof v !== 'number') return null;
  const s = String(v).trim();
  if (!s || s.length > max || !/^[#\w\s.,%()+-]+$/.test(s)) return null;
  for (const m of s.matchAll(/([\w-]*)\s*\(/g)) if (!/^(var|calc|rgba?|hsla?)$/i.test(m[1])) return null;
  return s;
}
