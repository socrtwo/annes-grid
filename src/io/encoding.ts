/**
 * Text encoding detection and decoding for imported text files (CSV, TSV, fixed width, SYLK,
 * DIF, ...).
 *
 * Detection order: byte order marks; BOM-less UTF-16 (zero-byte pattern); strict UTF-8; then the
 * legacy code pages, each decoded and scored for plausibility:
 *   - single-byte: Windows-1250..1258, Mac OS Roman, DOS code pages 437 and 850, KOI8-R;
 *   - multi-byte: Shift-JIS, EUC-JP, GBK/GB18030, Big5, EUC-KR (rejected when not valid).
 * Single-byte candidates are scored by whether non-ASCII characters are letters that sit inside
 * words of the same script with plausible letter case; multi-byte candidates by script mix
 * (kana for Japanese, Hangul for Korean) and a small list of very common Han characters to tell
 * simplified (GBK) from traditional (Big5) Chinese.
 *
 * TextDecoder covers every label used here except the DOS code pages 437 and 850, which are
 * decoded from the tables below.
 */

const CP437_HIGH = 'ÇüéâäàåçêëèïîìÄÅÉæÆôöòûùÿÖÜ¢£¥₧ƒáíóúñÑªº¿⌐¬½¼¡«»░▒▓│┤╡╢╖╕╣║╗╝╜╛┐└┴┬├─┼╞╟╚╔╩╦╠═╬╧╨╤╥╙╘╒╓╫╪┘┌█▄▌▐▀αßΓπΣσµτΦΘΩδ∞φε∩≡±≥≤⌠⌡÷≈°∙·√ⁿ²■ ';
const CP850_HIGH = 'ÇüéâäàåçêëèïîìÄÅÉæÆôöòûùÿÖÜø£Ø×ƒáíóúñÑªº¿®¬½¼¡«»░▒▓│┤ÁÂÀ©╣║╗╝¢¥┐└┴┬├─┼ãÃ╚╔╩╦╠═╬¤ðÐÊËÈıÍÎÏ┘┌█▄¦Ì▀ÓßÔÒõÕµþÞÚÛÙýÝ¯´­±‗¾¶§÷¸°¨·¹³²■ ';

/** Central European DOS (Latin 2), used by Polish, Czech, Slovak and Hungarian dBASE and Lotus files. */
const CP852_HIGH = 'ÇüéâäůćçłëŐőîŹÄĆÉĹĺôöĽľŚśÖÜŤťŁ×čáíóúĄąŽžĘę¬źČş«»░▒▓│┤ÁÂĚŞ╣║╗╝Żż┐└┴┬├─┼Ăă╚╔╩╦╠═╬¤đĐĎËďŇÍÎě┘┌█▄ŢŮ▀ÓßÔŃńňŠšŔÚŕŰýÝţ´­˝˛ˇ˘§÷¸°¨˙űŘř■ ';

const DOS_TABLES: Record<string, string> = { cp437: CP437_HIGH, cp850: CP850_HIGH, cp852: CP852_HIGH };

/** Normalizes an encoding label: 'CP-1252' -> 'windows-1252', 'ibm437' -> 'cp437', 'mac' -> 'macintosh'. */
export function normalizeEncoding(label: string): string {
  const l = label.trim().toLowerCase().replace(/_/g, '-');
  if (/^(cp|ibm|dos)-?437$/.test(l) || l === 'oem-us') return 'cp437';
  if (/^(cp|ibm|dos)-?850$/.test(l)) return 'cp850';
  if (/^(cp|ibm|dos)-?852$/.test(l)) return 'cp852';
  const w = /^(?:cp|windows|win)-?(125\d)$/.exec(l);
  if (w) return `windows-${w[1]}`;
  if (l === 'mac' || l === 'macroman' || l === 'mac-roman' || l === 'x-mac-roman') return 'macintosh';
  if (l === 'sjis' || l === 'shift-jis' || l === 'cp932' || l === 'ms932' || l === 'windows-31j') return 'shift_jis';
  if (l === 'cp936' || l === 'gb2312') return 'gbk';
  if (l === 'cp950') return 'big5';
  if (l === 'cp949' || l === 'uhc') return 'euc-kr';
  if (l === 'unicode' || l === 'utf-16') return 'utf-16le';
  if (l === 'latin1' || l === 'latin-1') return 'windows-1252';
  if (l === 'cp866' || l === 'ibm866') return 'ibm866';
  return l;
}

/** Decodes bytes with an encoding label (TextDecoder labels plus cp437 and cp850). */
export function decodeWith(u8: Uint8Array, encoding: string): string {
  const enc = normalizeEncoding(encoding);
  const table = DOS_TABLES[enc];
  if (table) {
    let s = '';
    const CH = 8192;
    for (let i = 0; i < u8.length; i += CH) {
      const part: string[] = [];
      const end = Math.min(u8.length, i + CH);
      for (let j = i; j < end; j++) { const b = u8[j]; part.push(b < 128 ? String.fromCharCode(b) : table[b - 128]); }
      s += part.join('');
    }
    return s;
  }
  return new TextDecoder(enc).decode(u8);
}

function tryDecode(u8: Uint8Array, enc: string, fatal = false): string | null {
  try {
    if (DOS_TABLES[enc]) return decodeWith(u8, enc);
    return new TextDecoder(enc, { fatal }).decode(u8);
  } catch { return null; }
}

export interface EncodingGuess { encoding: string; bom: number; confidence: 'certain' | 'high' | 'low' }

const SINGLE_BYTE = ['windows-1252', 'windows-1250', 'windows-1251', 'windows-1253', 'windows-1254', 'windows-1255', 'windows-1256', 'windows-1257', 'windows-1258', 'macintosh', 'cp850', 'cp437', 'koi8-r'];
const MULTI_BYTE = ['shift_jis', 'euc-jp', 'gbk', 'big5', 'euc-kr'];

const RE_LETTER = /\p{L}/u;
const RE_UPPER = /\p{Lu}/u;
const RE_LOWER = /\p{Ll}/u;
const SCRIPTS: [string, RegExp][] = [
  ['Latin', /\p{Script=Latin}/u], ['Cyrillic', /\p{Script=Cyrillic}/u], ['Greek', /\p{Script=Greek}/u],
  ['Hebrew', /\p{Script=Hebrew}/u], ['Arabic', /\p{Script=Arabic}/u],
];
const scriptOf = (ch: string): string => { for (const [n, re] of SCRIPTS) if (re.test(ch)) return n; return ''; };
/** Symbols common in real data (currency, degree, fractions, quotes, dashes). */
const COMMON_SYMBOLS = new Set('€£¥¢°±×÷µ§©®™½¼¾·–—‘’“”„…•«»¿¡ªº²³¹ '.split(''));

/** Plausibility score of decoded text: higher is better (per non-ASCII character). */
function scoreSingleByte(s: string): number {
  let score = 0, n = 0;
  for (let i = 0; i < s.length; i++) {
    const code = s.charCodeAt(i);
    if (code < 128) continue;
    n++;
    const ch = s[i];
    if ((code >= 0x80 && code < 0xa0) || code === 0xfffd) { score -= 5; continue; }
    if (RE_LETTER.test(ch)) {
      const sc = scriptOf(ch);
      const prev = s[i - 1] || ' ', next = s[i + 1] || ' ';
      let good = 1;
      for (const nb of [prev, next]) {
        if (RE_LETTER.test(nb)) {
          const ns = nb.charCodeAt(0) < 128 ? 'Latin' : scriptOf(nb);
          good += ns === sc ? 1 : -3;
        }
      }
      // An upper-case letter right after a lower-case one ("cafŽ") is unlikely.
      if (RE_UPPER.test(ch) && RE_LOWER.test(prev)) good -= 2;
      // Final letter forms (Greek final sigma, Hebrew finals) only end words.
      if ('ςךםןףץ'.includes(ch) && RE_LETTER.test(next)) good -= 3;
      // Accented Latin letters sit in words with plain ASCII letters; a word made only of
      // accented letters is more likely another script decoded as Latin.
      if (sc === 'Latin') {
        let a = i, b = i;
        while (a > 0 && RE_LETTER.test(s[a - 1])) a--;
        while (b < s.length - 1 && RE_LETTER.test(s[b + 1])) b++;
        if (b > a && !/[A-Za-z]/.test(s.slice(a, b + 1))) good -= 2;
      }
      score += good;
    } else if (COMMON_SYMBOLS.has(ch)) score += 0.5;
    else if (code >= 0x2500 && code <= 0x25ff) score -= 0.5; // box drawing: plausible in DOS files only
    else score -= 1.5;
  }
  if (!n) return 0;
  // Slavic text decoded as Cyrillic has about 40% vowels; other scripts read as Cyrillic do not.
  const cyr = s.match(/\p{Script=Cyrillic}/gu);
  if (cyr && cyr.length >= 6) {
    const vowels = cyr.filter((c) => 'аеёиоуыэюяіїєАЕЁИОУЫЭЮЯІЇЄ'.includes(c)).length;
    if (vowels / cyr.length < 0.25) score -= 2 * n;
  }
  return score / n;
}

const SIMPLIFIED = '的一是不了人我在有他这中大来上国个到说们为子和你地出道也时年会对能就要以发生学可下过';
const TRADITIONAL = '的一是不了人我在有他這中大來上國個到說們為子和你地出道也時年會對能就要以發生學可下過';

function scoreMultiByte(s: string, enc: string): number {
  let han = 0, kana = 0, hangul = 0, halfKana = 0, other = 0, n = 0, simp = 0, trad = 0;
  for (const ch of s) {
    const code = ch.codePointAt(0)!;
    if (code < 128) continue;
    n++;
    if (code === 0xfffd || (code >= 0xe000 && code <= 0xf8ff)) other += 3;
    else if (code >= 0x3040 && code <= 0x30ff) kana++;
    else if (code >= 0xff61 && code <= 0xff9f) halfKana++;
    else if (code >= 0xac00 && code <= 0xd7a3) hangul++;
    else if (code >= 0x4e00 && code <= 0x9fff) { han++; if (SIMPLIFIED.includes(ch) && !TRADITIONAL.includes(ch)) simp++; if (TRADITIONAL.includes(ch) && !SIMPLIFIED.includes(ch)) trad++; }
    else if ((code >= 0x3000 && code <= 0x303f) || (code >= 0xff01 && code <= 0xff5e)) { /* CJK punctuation, full width forms */ }
    else other++;
  }
  if (!n) return 0;
  let score = 2 - (other * 2 + halfKana * 1.5) / n;
  if (enc === 'shift_jis' || enc === 'euc-jp') score += kana / n > 0.05 ? 1.5 : -0.5;
  else if (enc === 'euc-kr') score += hangul / n > 0.3 ? 2 * (hangul / n) - 3 * (han / n) : -1;
  else score += han / n > 0.5 ? 0.5 : -0.5;
  if (enc === 'gbk') score += Math.min(2, (simp - trad) * 0.3);
  if (enc === 'big5') score += Math.min(2, (trad - simp) * 0.3);
  return score;
}

/** Guesses the encoding of text bytes. */
export function detectEncoding(data: Uint8Array | ArrayBuffer, sampleSize = 256 * 1024): EncodingGuess {
  const u8 = data instanceof Uint8Array ? data : new Uint8Array(data);
  if (u8[0] === 0xef && u8[1] === 0xbb && u8[2] === 0xbf) return { encoding: 'utf-8', bom: 3, confidence: 'certain' };
  if (u8[0] === 0xff && u8[1] === 0xfe) return { encoding: 'utf-16le', bom: 2, confidence: 'certain' };
  if (u8[0] === 0xfe && u8[1] === 0xff) return { encoding: 'utf-16be', bom: 2, confidence: 'certain' };
  const sample = u8.subarray(0, Math.min(u8.length, sampleSize));
  if (sample.length >= 4) {
    let zo = 0, ze = 0;
    const n = Math.min(sample.length, 4000);
    for (let i = 0; i < n; i++) if (sample[i] === 0) { if (i & 1) zo++; else ze++; }
    if (zo > n / 4 && ze < n / 50) return { encoding: 'utf-16le', bom: 0, confidence: 'high' };
    if (ze > n / 4 && zo < n / 50) return { encoding: 'utf-16be', bom: 0, confidence: 'high' };
  }
  let ascii = true;
  for (let i = 0; i < sample.length; i++) if (sample[i] >= 0x80) { ascii = false; break; }
  if (ascii) return { encoding: 'utf-8', bom: 0, confidence: 'high' };
  // Strict UTF-8 (a sample may cut a sequence at its end: drop up to 3 trailing bytes).
  let cut = sample.length;
  if (cut < u8.length) { let k = cut - 1; while (k > cut - 4 && k >= 0 && (sample[k] & 0xc0) === 0x80) k--; if (k >= 0 && sample[k] >= 0xc0) cut = k; }
  if (tryDecode(sample.subarray(0, cut), 'utf-8', true) != null) return { encoding: 'utf-8', bom: 0, confidence: 'high' };

  let best = 'windows-1252', bestScore = -Infinity;
  const consider = (enc: string, score: number) => { if (score > bestScore + 1e-9) { best = enc; bestScore = score; } };
  for (const enc of SINGLE_BYTE) {
    const s = tryDecode(sample, enc);
    if (s != null) consider(enc, scoreSingleByte(s));
  }
  // Multi-byte: must decode without errors (allowing a sequence cut at the end of the sample).
  let mbCut = sample.length;
  if (mbCut < u8.length) { mbCut = sample.length - 1; while (mbCut > 0 && sample[mbCut] >= 0x80) mbCut--; mbCut++; }
  for (const enc of MULTI_BYTE) {
    const s = tryDecode(sample.subarray(0, mbCut), enc, true);
    if (s != null) consider(enc, scoreMultiByte(s, enc));
  }
  return { encoding: best, bom: 0, confidence: bestScore > 1 ? 'high' : 'low' };
}

/** Decodes text bytes, detecting the encoding unless one is given ('auto' detects too). */
export function decodeAuto(data: Uint8Array | ArrayBuffer, encoding?: string): { text: string; encoding: string } {
  const u8 = data instanceof Uint8Array ? data : new Uint8Array(data);
  if (encoding && encoding !== 'auto') {
    const enc = normalizeEncoding(encoding);
    let start = 0;
    if (enc === 'utf-8' && u8[0] === 0xef && u8[1] === 0xbb && u8[2] === 0xbf) start = 3;
    if ((enc === 'utf-16le' && u8[0] === 0xff && u8[1] === 0xfe) || (enc === 'utf-16be' && u8[0] === 0xfe && u8[1] === 0xff)) start = 2;
    const text = tryDecode(u8.subarray(start), enc);
    if (text != null) return { text, encoding: enc };
  }
  const g = detectEncoding(u8);
  return { text: tryDecode(u8.subarray(g.bom), g.encoding) ?? new TextDecoder('windows-1252').decode(u8), encoding: g.encoding };
}
