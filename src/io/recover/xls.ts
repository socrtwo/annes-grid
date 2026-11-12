/**
 * Excel 97-2003 (.xls, BIFF8) reader, with best-effort support for Excel 5.0/95 (BIFF5/7, code
 * pages) and the single-sheet BIFF2-4 formats.
 *
 * The OLE2 compound file is opened with readCfb (doc97.ts) and the "Workbook" (BIFF8) or "Book"
 * (BIFF5) stream is walked record by record. The reader decodes:
 *
 *  - workbook globals: BOF/EOF, BOUNDSHEET, SST (with CONTINUE, split strings, rich and phonetic
 *    runs), FONT, FORMAT, XF, PALETTE, DATEMODE, CODEPAGE, NAME, SUPBOOK, EXTERNSHEET,
 *    EXTERNNAME, PROTECT/PASSWORD, FILEPASS (encryption is reported, not decrypted);
 *  - cells: LABELSST, LABEL, RSTRING, NUMBER, RK, MULRK, INTEGER (BIFF2), BOOLERR, BLANK,
 *    MULBLANK, FORMULA with its cached result (and the STRING record that follows a string
 *    result), SHRFMLA and ARRAY;
 *  - formulas: the parsed-expression (Ptg) token stream is turned back into formula text,
 *    including 3D references through EXTERNSHEET/SUPBOOK, defined names, add-in and "future"
 *    functions (ptgNameX), array constants and shared formulas (ptgRefN/ptgAreaN relative to
 *    the cell). A formula that cannot be decoded keeps its cached value only;
 *  - sheet layout: COLINFO, ROW, DEFCOLWIDTH, MERGEDCELLS, WINDOW2/PANE (frozen panes), SCL,
 *    NOTE/OBJ/TXO comments, HLINK hyperlinks, PROTECT.
 *
 * In lenient mode (recovery), a record with an impossible length or an unknown identifier
 * followed by garbage triggers a resynchronization: the reader scans forward for the next
 * plausible record header (a known identifier whose length leads to another known header) and
 * continues. When the compound file itself is damaged, the raw file bytes are scanned for BIFF
 * BOF records and walked the same way.
 */
import { Workbook, Sheet } from '../../model/workbook';
import { CellError, errorFromCode, type Border, type Cell, type CellStyle, type Scalar } from '../../model/types';
import { cellKey, colName, MAX_COLS, MAX_ROWS } from '../../model/address';
import { BUILTIN_FORMATS } from '../../model/numfmt';
import { isOle2, readCfb } from './doc97';
import { cp1252 } from './salvage';

export const XLS_ENCRYPTED_MSG = 'This workbook is password-protected (encrypted); Anne\'s Grid cannot decrypt it.';

export interface BiffStats {
  /** BIFF version: 2, 3, 4, 5 (Excel 5.0/95) or 8 (Excel 97-2003). */
  biff: number;
  /** Where the record stream came from. */
  source: 'cfb' | 'raw' | 'scan';
  sheets: number;
  cells: number;
  formulas: number;
  formulasDecoded: number;
  sharedStrings: number;
  records: number;
  resyncs: number;
  skippedBytes: number;
  codepage: number;
  encrypted: boolean;
}

export interface BiffResult { book: Workbook; warnings: string[]; stats: BiffStats }

// ------------------------------------------------------------------------------------------
// Record identifiers

const R = {
  FORMULA: 0x0006, EOF: 0x000a, CALCMODE: 0x000d, PROTECT: 0x0012, PASSWORD: 0x0013, EXTERNSHEET: 0x0017, NAME: 0x0018,
  WINDOWPROTECT: 0x0019, NOTE: 0x001c, FORMAT_OLD: 0x001e, DATEMODE: 0x0022, EXTERNNAME: 0x0023, COLWIDTH2: 0x0024,
  FILEPASS: 0x002f, FONT: 0x0031, CONTINUE: 0x003c, PANE: 0x0041, CODEPAGE: 0x0042, XF2: 0x0043, DEFCOLWIDTH: 0x0055,
  OBJ: 0x005d, COLINFO: 0x007d, BOUNDSHEET: 0x0085, PALETTE: 0x0092, SCL: 0x00a0, MULRK: 0x00bd, MULBLANK: 0x00be,
  RSTRING: 0x00d6, XF: 0x00e0, MERGEDCELLS: 0x00e5, SST: 0x00fc, LABELSST: 0x00fd, SUPBOOK: 0x01ae, TXO: 0x01b6,
  HLINK: 0x01b8, BLANK: 0x0201, NUMBER: 0x0203, LABEL: 0x0204, BOOLERR: 0x0205, FORMULA3: 0x0206, STRING: 0x0207,
  ROW: 0x0208, ARRAY: 0x0221, ARRAY2: 0x0021, WINDOW2: 0x023e, XF3: 0x0243, RK: 0x027e, FORMAT: 0x041e, FORMULA4: 0x0406,
  XF4: 0x0443, SHRFMLA: 0x04bc, BOF: 0x0809, BOF2: 0x0009, BOF3: 0x0209, BOF4: 0x0409,
  BLANK2: 0x0001, INTEGER2: 0x0002, NUMBER2: 0x0003, LABEL2: 0x0004, BOOLERR2: 0x0005, STRING2: 0x0007, ROW2: 0x0008,
  FONT34: 0x0231, ROW34: 0x0208, SHEETPR: 0x0081,
} as const;

/** Identifiers seen in real BIFF2-8 streams; used to judge plausibility during resynchronization. */
const KNOWN = new Set<number>([
  0x0001, 0x0002, 0x0003, 0x0004, 0x0005, 0x0006, 0x0007, 0x0008, 0x0009, 0x000a, 0x000c, 0x000d, 0x000e, 0x000f, 0x0010,
  0x0011, 0x0012, 0x0013, 0x0014, 0x0015, 0x0016, 0x0017, 0x0018, 0x0019, 0x001a, 0x001b, 0x001c, 0x001d, 0x001e, 0x001f,
  0x0020, 0x0021, 0x0022, 0x0023, 0x0024, 0x0025, 0x0026, 0x0027, 0x0028, 0x0029, 0x002a, 0x002b, 0x002f, 0x0031, 0x003c,
  0x003d, 0x003e, 0x0040, 0x0041, 0x0042, 0x0043, 0x0044, 0x0045, 0x004d, 0x0051, 0x0055, 0x0059, 0x005a, 0x005b, 0x005c,
  0x005d, 0x005e, 0x005f, 0x0060, 0x0063, 0x007d, 0x007e, 0x0080, 0x0081, 0x0082, 0x0083, 0x0084, 0x0085, 0x0086, 0x0087,
  0x0088, 0x008c, 0x008d, 0x008e, 0x008f, 0x0092, 0x0093, 0x0094, 0x0095, 0x0096, 0x0097, 0x0098, 0x0099, 0x009a, 0x009b,
  0x009c, 0x009d, 0x009e, 0x00a0, 0x00a1, 0x00a9, 0x00aa, 0x00ab, 0x00b0, 0x00b1, 0x00b2, 0x00b6, 0x00bd, 0x00be, 0x00bf,
  0x00c0, 0x00c1, 0x00c2, 0x00c5, 0x00c6, 0x00c7, 0x00c8, 0x00c9, 0x00ca, 0x00cb, 0x00cc, 0x00cd, 0x00ce, 0x00cf, 0x00d0,
  0x00d1, 0x00d2, 0x00d3, 0x00d5, 0x00d6, 0x00d7, 0x00d8, 0x00d9, 0x00da, 0x00db, 0x00dc, 0x00dd, 0x00de, 0x00e0, 0x00e1,
  0x00e2, 0x00e3, 0x00e5, 0x00e9, 0x00ea, 0x00eb, 0x00ec, 0x00ed, 0x00ef, 0x00f1, 0x00f6, 0x00fc, 0x00fd, 0x00ff, 0x0100,
  0x0103, 0x0105, 0x0122, 0x0123, 0x0124, 0x0125, 0x013d, 0x0140, 0x0141, 0x015f, 0x0160, 0x0161, 0x0162, 0x01a5, 0x01a9,
  0x01aa, 0x01ab, 0x01ad, 0x01ae, 0x01af, 0x01b0, 0x01b1, 0x01b2, 0x01b5, 0x01b6, 0x01b7, 0x01b8, 0x01ba, 0x01bb, 0x01bc,
  0x01be, 0x01c0, 0x01c1, 0x01c2, 0x0200, 0x0201, 0x0203, 0x0204, 0x0205, 0x0206, 0x0207, 0x0208, 0x0209, 0x020b, 0x0218,
  0x021b, 0x021d, 0x0221, 0x0223, 0x0225, 0x0231, 0x0236, 0x023e, 0x0243, 0x0293, 0x027e, 0x0406, 0x0409, 0x041e, 0x0443,
  0x04bc, 0x0800, 0x0801, 0x0802, 0x0803, 0x0804, 0x0805, 0x0806, 0x0807, 0x0808, 0x0809, 0x080a, 0x080b, 0x0810, 0x0812,
  0x0850, 0x0851, 0x0852, 0x0853, 0x0854, 0x0855, 0x0856, 0x0857, 0x0858, 0x0859, 0x085a, 0x085b, 0x085c, 0x0860, 0x0861,
  0x0862, 0x0863, 0x0864, 0x0865, 0x0866, 0x0867, 0x0868, 0x0869, 0x086a, 0x086b, 0x086c, 0x086d, 0x0870, 0x0871, 0x0872,
  0x0873, 0x0874, 0x0875, 0x0876, 0x0877, 0x0878, 0x0879, 0x087a, 0x087b, 0x087c, 0x087d, 0x087e, 0x087f, 0x0880, 0x0881,
  0x0882, 0x0883, 0x0884, 0x0885, 0x0886, 0x0887, 0x0888, 0x0889, 0x088a, 0x088b, 0x088c, 0x088d, 0x088e, 0x088f, 0x0890,
  0x0891, 0x0892, 0x0893, 0x0894, 0x0895, 0x0896, 0x0897, 0x0898, 0x0899, 0x089a, 0x089b, 0x089c, 0x089d, 0x089e, 0x08a0,
  0x08a1, 0x08a3, 0x08a4, 0x08a5, 0x08a6, 0x08a7, 0x08c1, 0x08c2, 0x08c3, 0x08c4, 0x08c5, 0x08c6, 0x08c7, 0x08c8, 0x08c9,
  0x08ca, 0x08cb, 0x08cc, 0x08cd, 0x08ce, 0x1001, 0x1002, 0x1003, 0x1004, 0x1005, 0x1006, 0x1007, 0x1008, 0x1009, 0x100a,
  0x100b, 0x100c, 0x100d, 0x100e, 0x100f, 0x1014, 0x1015, 0x1016, 0x1017, 0x1018, 0x1019, 0x101a, 0x101b, 0x101c, 0x101d,
  0x101e, 0x101f, 0x1020, 0x1021, 0x1022, 0x1024, 0x1025, 0x1026, 0x1027, 0x1032, 0x1033, 0x1034, 0x1035, 0x1036, 0x1037,
  0x103a, 0x103c, 0x103d, 0x103e, 0x1041, 0x1043, 0x1044, 0x1045, 0x1046, 0x104a, 0x104b, 0x104e, 0x104f, 0x1050, 0x1051,
  0x105c, 0x105d, 0x105f, 0x1060, 0x1061, 0x1062, 0x1063, 0x1064, 0x1065, 0x1066,
]);

const isBofId = (id: number) => id === R.BOF || id === R.BOF2 || id === R.BOF3 || id === R.BOF4;

// ------------------------------------------------------------------------------------------
// Byte helpers

const u16 = (b: Uint8Array, o: number) => (o + 2 <= b.length ? b[o] | (b[o + 1] << 8) : 0);
const i16 = (b: Uint8Array, o: number) => { const v = u16(b, o); return v & 0x8000 ? v - 0x10000 : v; };
const u32 = (b: Uint8Array, o: number) => (o + 4 <= b.length ? (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0 : 0);
function f64(b: Uint8Array, o: number): number {
  if (o + 8 > b.length) throw new Error('truncated number');
  return new DataView(b.buffer, b.byteOffset + o, 8).getFloat64(0, true);
}
function rk(v: number): number {
  let n: number;
  if (v & 2) n = v >> 2; // 30-bit signed integer
  else {
    const dv = new DataView(new ArrayBuffer(8));
    dv.setUint32(0, 0, true);
    dv.setUint32(4, v & 0xfffffffc, true);
    n = dv.getFloat64(0, true);
  }
  return v & 1 ? n / 100 : n;
}

class EndOfData extends Error { constructor() { super('unexpected end of record data'); } }

/** Reads across a record and its CONTINUE records. Strings re-read the high-byte flag at each boundary. */
class Cursor {
  pi = 0;
  p: number;
  constructor(public parts: Uint8Array[], start = 0) { this.p = start; }
  private ensure(): void { while (this.pi < this.parts.length && this.p >= this.parts[this.pi].length) { this.pi++; this.p = 0; } }
  eof(): boolean { this.ensure(); return this.pi >= this.parts.length; }
  u8(): number { this.ensure(); if (this.pi >= this.parts.length) throw new EndOfData(); return this.parts[this.pi][this.p++]; }
  u16(): number { const a = this.u8(); return a | (this.u8() << 8); }
  u32(): number { const a = this.u16(); return (a | (this.u16() << 16)) >>> 0; }
  skip(n: number): void {
    while (n > 0) {
      this.ensure();
      if (this.pi >= this.parts.length) throw new EndOfData();
      const take = Math.min(n, this.parts[this.pi].length - this.p);
      this.p += take; n -= take;
    }
  }
  bytes(n: number): Uint8Array {
    const out = new Uint8Array(n);
    for (let i = 0; i < n; i++) out[i] = this.u8();
    return out;
  }
  /** n characters of a BIFF8 string; at a CONTINUE boundary a new option byte gives the width. */
  chars(n: number, high: boolean): string {
    let s = '';
    while (n > 0) {
      if (this.pi >= this.parts.length || this.p >= this.parts[this.pi].length) {
        this.pi++; this.p = 0;
        if (this.pi >= this.parts.length) throw new EndOfData();
        high = !!(this.parts[this.pi][this.p++] & 1);
        continue;
      }
      const part = this.parts[this.pi];
      const avail = part.length - this.p;
      if (high) {
        const take = Math.min(n, avail >> 1);
        if (!take) { this.p = part.length; continue; }
        for (let k = 0; k < take; k++) s += String.fromCharCode(part[this.p + 2 * k] | (part[this.p + 2 * k + 1] << 8));
        this.p += take * 2; n -= take;
      } else {
        const take = Math.min(n, avail);
        for (let k = 0; k < take; k++) s += String.fromCharCode(part[this.p + k]);
        this.p += take; n -= take;
      }
    }
    return s;
  }
}

const CODEPAGES: Record<number, string> = {
  367: 'ascii', 437: 'ibm437', 720: 'ibm720', 737: 'ibm737', 775: 'ibm775', 850: 'ibm850', 852: 'ibm852', 855: 'ibm855', 857: 'ibm857',
  860: 'ibm860', 861: 'ibm861', 862: 'ibm862', 863: 'ibm863', 864: 'ibm864', 865: 'ibm865', 866: 'ibm866', 869: 'ibm869', 874: 'windows-874',
  932: 'shift_jis', 936: 'gbk', 949: 'euc-kr', 950: 'big5', 1200: 'utf-16le', 1250: 'windows-1250', 1251: 'windows-1251',
  1252: 'windows-1252', 1253: 'windows-1253', 1254: 'windows-1254', 1255: 'windows-1255', 1256: 'windows-1256', 1257: 'windows-1257',
  1258: 'windows-1258', 1361: 'johab', 10000: 'macintosh', 32768: 'macintosh', 32769: 'windows-1252',
};

function makeDecoder(cp: number): (b: Uint8Array) => string {
  // Some runtimes decode "windows-1252" as ISO-8859-1; map the 0x80-0x9F range explicitly.
  if (cp === 1252 || cp === 32769 || !CODEPAGES[cp] || cp === 1200) return (b) => { let s = ''; for (const x of b) s += cp1252(x); return s; };
  const label = CODEPAGES[cp] && cp !== 1200 ? CODEPAGES[cp] : 'windows-1252';
  try {
    const d = new TextDecoder(label);
    return (b) => d.decode(b);
  } catch {
    return (b) => { let s = ''; for (const x of b) s += String.fromCharCode(x); return s; };
  }
}

// ------------------------------------------------------------------------------------------
// Built-in function table (BIFF function index -> name, fixed argument count or -1)

const FN_SRC =
  '0 COUNT -1|1 IF -1|2 ISNA 1|3 ISERROR 1|4 SUM -1|5 AVERAGE -1|6 MIN -1|7 MAX -1|8 ROW -1|9 COLUMN -1|10 NA 0|11 NPV -1|' +
  '12 STDEV -1|13 DOLLAR -1|14 FIXED -1|15 SIN 1|16 COS 1|17 TAN 1|18 ATAN 1|19 PI 0|20 SQRT 1|21 EXP 1|22 LN 1|23 LOG10 1|' +
  '24 ABS 1|25 INT 1|26 SIGN 1|27 ROUND 2|28 LOOKUP -1|29 INDEX -1|30 REPT 2|31 MID 3|32 LEN 1|33 VALUE 1|34 TRUE 0|35 FALSE 0|' +
  '36 AND -1|37 OR -1|38 NOT 1|39 MOD 2|40 DCOUNT 3|41 DSUM 3|42 DAVERAGE 3|43 DMIN 3|44 DMAX 3|45 DSTDEV 3|46 VAR -1|47 DVAR 3|' +
  '48 TEXT 2|49 LINEST -1|50 TREND -1|51 LOGEST -1|52 GROWTH -1|53 GOTO -1|54 HALT -1|55 RETURN -1|56 PV -1|57 FV -1|58 NPER -1|' +
  '59 PMT -1|60 RATE -1|61 MIRR 3|62 IRR -1|63 RAND 0|64 MATCH -1|65 DATE 3|66 TIME 3|67 DAY 1|68 MONTH 1|69 YEAR 1|' +
  '70 WEEKDAY -1|71 HOUR 1|72 MINUTE 1|73 SECOND 1|74 NOW 0|75 AREAS 1|76 ROWS 1|77 COLUMNS 1|78 OFFSET -1|79 ABSREF 2|' +
  '80 RELREF 2|81 ARGUMENT -1|82 SEARCH -1|83 TRANSPOSE 1|84 ERROR -1|85 STEP 0|86 TYPE 1|87 ECHO -1|88 SET.NAME -1|89 CALLER 0|' +
  '90 DEREF 1|91 WINDOWS -1|92 SERIES -1|93 DOCUMENTS -1|94 ACTIVE.CELL 0|95 SELECTION 0|96 RESULT -1|97 ATAN2 2|98 ASIN 1|' +
  '99 ACOS 1|100 CHOOSE -1|101 HLOOKUP -1|102 VLOOKUP -1|103 LINKS -1|104 INPUT -1|105 ISREF 1|106 GET.FORMULA 1|' +
  '107 GET.NAME -1|108 SET.VALUE 2|109 LOG -1|110 EXEC -1|111 CHAR 1|112 LOWER 1|113 UPPER 1|114 PROPER 1|115 LEFT -1|' +
  '116 RIGHT -1|117 EXACT 2|118 TRIM 1|119 REPLACE 4|120 SUBSTITUTE -1|121 CODE 1|122 NAMES -1|123 DIRECTORY -1|124 FIND -1|' +
  '125 CELL -1|126 ISERR 1|127 ISTEXT 1|128 ISNUMBER 1|129 ISBLANK 1|130 T 1|131 N 1|132 FOPEN -1|133 FCLOSE 1|134 FSIZE 1|' +
  '135 FREADLN 1|136 FREAD 2|137 FWRITELN 2|138 FWRITE 2|139 FPOS -1|140 DATEVALUE 1|141 TIMEVALUE 1|142 SLN 3|143 SYD 4|' +
  '144 DDB -1|145 GET.DEF -1|146 REFTEXT -1|147 TEXTREF -1|148 INDIRECT -1|149 REGISTER -1|150 CALL -1|151 ADD.BAR -1|' +
  '152 ADD.MENU -1|153 ADD.COMMAND -1|154 ENABLE.COMMAND -1|155 CHECK.COMMAND -1|156 RENAME.COMMAND -1|157 SHOW.BAR -1|' +
  '158 DELETE.MENU -1|159 DELETE.COMMAND -1|160 GET.CHART.ITEM -1|161 DIALOG.BOX 1|162 CLEAN 1|163 MDETERM 1|164 MINVERSE 1|' +
  '165 MMULT 2|166 FILES -1|167 IPMT -1|168 PPMT -1|169 COUNTA -1|170 CANCEL.KEY -1|171 FOR -1|172 WHILE 1|173 BREAK 0|' +
  '174 NEXT 0|175 INITIATE 2|176 REQUEST 2|177 POKE 3|178 EXECUTE 2|179 TERMINATE 1|180 RESTART -1|181 HELP -1|182 GET.BAR -1|' +
  '183 PRODUCT -1|184 FACT 1|185 GET.CELL -1|186 GET.WORKSPACE 1|187 GET.WINDOW -1|188 GET.DOCUMENT -1|189 DPRODUCT 3|' +
  '190 ISNONTEXT 1|191 GET.NOTE -1|192 NOTE -1|193 STDEVP -1|194 VARP -1|195 DSTDEVP 3|196 DVARP 3|197 TRUNC -1|198 ISLOGICAL 1|' +
  '199 DCOUNTA 3|200 DELETE.BAR 1|201 UNREGISTER 1|204 USDOLLAR -1|205 FINDB -1|206 SEARCHB -1|207 REPLACEB 4|208 LEFTB -1|' +
  '209 RIGHTB -1|210 MIDB 3|211 LENB 1|212 ROUNDUP 2|213 ROUNDDOWN 2|214 ASC 1|215 DBCS 1|216 RANK -1|219 ADDRESS -1|' +
  '220 DAYS360 -1|221 TODAY 0|222 VDB -1|223 ELSE 0|224 ELSE.IF 1|225 END.IF 0|226 FOR.CELL -1|227 MEDIAN -1|228 SUMPRODUCT -1|' +
  '229 SINH 1|230 COSH 1|231 TANH 1|232 ASINH 1|233 ACOSH 1|234 ATANH 1|235 DGET 3|236 CREATE.OBJECT -1|237 VOLATILE -1|' +
  '238 LAST.ERROR 0|239 CUSTOM.UNDO -1|240 CUSTOM.REPEAT -1|241 FORMULA.CONVERT -1|242 GET.LINK.INFO -1|243 TEXT.BOX -1|' +
  '244 INFO 1|245 GROUP 0|246 GET.OBJECT -1|247 DB -1|248 PAUSE -1|251 RESUME -1|252 FREQUENCY 2|253 ADD.TOOLBAR -1|' +
  '254 DELETE.TOOLBAR 1|256 RESET.TOOLBAR 1|257 EVALUATE 1|258 GET.TOOLBAR -1|259 GET.TOOL -1|260 SPELLING.CHECK -1|' +
  '261 ERROR.TYPE 1|262 APP.TITLE -1|263 WINDOW.TITLE -1|264 SAVE.TOOLBAR -1|265 ENABLE.TOOL 3|266 PRESS.TOOL 3|' +
  '267 REGISTER.ID -1|268 GET.WORKBOOK -1|269 AVEDEV -1|270 BETADIST -1|271 GAMMALN 1|272 BETAINV -1|273 BINOMDIST 4|' +
  '274 CHIDIST 2|275 CHIINV 2|276 COMBIN 2|277 CONFIDENCE 3|278 CRITBINOM 3|279 EVEN 1|280 EXPONDIST 3|281 FDIST 3|282 FINV 3|' +
  '283 FISHER 1|284 FISHERINV 1|285 FLOOR 2|286 GAMMADIST 4|287 GAMMAINV 3|288 CEILING 2|289 HYPGEOMDIST 4|290 LOGNORMDIST 3|' +
  '291 LOGINV 3|292 NEGBINOMDIST 3|293 NORMDIST 4|294 NORMSDIST 1|295 NORMINV 3|296 NORMSINV 1|297 STANDARDIZE 3|298 ODD 1|' +
  '299 PERMUT 2|300 POISSON 3|301 TDIST 3|302 WEIBULL 4|303 SUMXMY2 2|304 SUMX2MY2 2|305 SUMX2PY2 2|306 CHITEST 2|307 CORREL 2|' +
  '308 COVAR 2|309 FORECAST 3|310 FTEST 2|311 INTERCEPT 2|312 PEARSON 2|313 RSQ 2|314 STEYX 2|315 SLOPE 2|316 TTEST 4|' +
  '317 PROB -1|318 DEVSQ -1|319 GEOMEAN -1|320 HARMEAN -1|321 SUMSQ -1|322 KURT -1|323 SKEW -1|324 ZTEST -1|325 LARGE 2|' +
  '326 SMALL 2|327 QUARTILE 2|328 PERCENTILE 2|329 PERCENTRANK -1|330 MODE -1|331 TRIMMEAN 2|332 TINV 2|334 MOVIE.COMMAND -1|' +
  '335 GET.MOVIE -1|336 CONCATENATE -1|337 POWER 2|338 PIVOT.ADD.DATA -1|339 GET.PIVOT.TABLE -1|340 GET.PIVOT.FIELD -1|' +
  '341 GET.PIVOT.ITEM -1|342 RADIANS 1|343 DEGREES 1|344 SUBTOTAL -1|345 SUMIF -1|346 COUNTIF 2|347 COUNTBLANK 1|' +
  '348 SCENARIO.GET -1|349 OPTIONS.LISTS.GET 1|350 ISPMT 4|351 DATEDIF 3|352 DATESTRING 1|353 NUMBERSTRING 2|354 ROMAN -1|' +
  '355 OPEN.DIALOG -1|356 SAVE.DIALOG -1|357 VIEW.GET -1|358 GETPIVOTDATA -1|359 HYPERLINK -1|360 PHONETIC 1|361 AVERAGEA -1|' +
  '362 MAXA -1|363 MINA -1|364 STDEVPA -1|365 VARPA -1|366 STDEVA -1|367 VARA -1|368 BAHTTEXT 1|369 THAIDAYOFWEEK 1|' +
  '370 THAIDIGIT 1|371 THAIMONTHOFYEAR 1|372 THAINUMSOUND 1|373 THAINUMSTRING 1|374 THAISTRINGLENGTH 1|375 ISTHAIDIGIT 1|' +
  '376 ROUNDBAHTDOWN 1|377 ROUNDBAHTUP 1|378 THAIYEAR 1|379 RTD -1';

/** BIFF built-in function table: index -> [name, fixed argument count (-1 = variable)]. */
export const BIFF_FUNCTIONS: Map<number, [string, number]> = new Map(
  FN_SRC.split('|').map((e) => { const [i, n, a] = e.split(' '); return [+i, [n, +a]] as [number, [string, number]]; }),
);

const ERR_CODES: Record<number, string> = { 0x00: '#NULL!', 0x07: '#DIV/0!', 0x0f: '#VALUE!', 0x17: '#REF!', 0x1d: '#NAME?', 0x24: '#NUM!', 0x2a: '#N/A', 0x2b: '#GETTING_DATA' };
const errValue = (code: number): CellError => errorFromCode(ERR_CODES[code] || '#N/A') || new CellError('#N/A');

const BUILTIN_NAMES = ['Consolidate_Area', 'Auto_Open', 'Auto_Close', 'Extract', 'Database', 'Criteria', 'Print_Area', 'Print_Titles', 'Recorder', 'Data_Form', 'Auto_Activate', 'Auto_Deactivate', 'Sheet_Title', '_FilterDatabase'];

/** Number formats 5-8 are locale currency formats that writers normally store explicitly. */
const EXTRA_BUILTIN_FORMATS: Record<number, string> = {
  5: '"$"#,##0_);\\("$"#,##0\\)', 6: '"$"#,##0_);[Red]\\("$"#,##0\\)', 7: '"$"#,##0.00_);\\("$"#,##0.00\\)', 8: '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
};

// Default BIFF8 color palette (indexes 8-63); 0-7 repeat the first eight entries.
const DEFAULT_PALETTE = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
  '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF', '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
  '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99', '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
  '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
];

const BORDER_STYLES: (Border['style'] | null)[] = [null, 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashed', 'dashDot', 'mediumDashed', 'dashDot'];
const PATTERNS = ['none', 'solid', 'mediumGray', 'darkGray', 'lightGray', 'darkHorizontal', 'darkVertical', 'darkDown', 'darkUp', 'darkGrid', 'darkTrellis', 'lightHorizontal', 'lightVertical', 'lightDown', 'lightUp', 'lightGrid', 'lightTrellis', 'gray125', 'gray0625'];
const H_ALIGN: CellStyle['hAlign'][] = ['general', 'left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed'];
const V_ALIGN: CellStyle['vAlign'][] = ['top', 'middle', 'bottom', 'justify', 'distributed'];

// ------------------------------------------------------------------------------------------
// Record walking

interface Rec { id: number; off: number; data: Uint8Array; conts: Uint8Array[] }

interface WalkStats { resyncs: number; skippedBytes: number; truncated: boolean }

function headerPlausible(s: Uint8Array, p: number, depth: number): boolean {
  if (p === s.length) return true;
  if (p + 4 > s.length) return false;
  const id = u16(s, p), len = u16(s, p + 2);
  if (!KNOWN.has(id) || len > 8224 || p + 4 + len > s.length) return false;
  if (id === 0 && len === 0) return false;
  return depth <= 0 || headerPlausible(s, p + 4 + len, depth - 1);
}

function walk(s: Uint8Array, start: number, lenient: boolean, st: WalkStats): Rec[] {
  const out: Rec[] = [];
  let p = start;
  while (p + 4 <= s.length) {
    const id = u16(s, p), len = u16(s, p + 2);
    let ok = len <= 8224 && p + 4 + len <= s.length;
    if (ok && lenient && !KNOWN.has(id)) ok = headerPlausible(s, p + 4 + len, 1);
    if (ok && lenient && id === 0 && len === 0) ok = false;
    if (!ok) {
      if (!lenient) {
        if (p + 4 + len > s.length) { st.truncated = true; break; }
        throw new Error(`Damaged record at byte ${p}`);
      }
      let q = -1;
      for (let k = p + 1; k + 4 <= s.length; k++) if (headerPlausible(s, k, 2)) { q = k; break; }
      if (q < 0) { st.skippedBytes += s.length - p; if (p + 4 + len > s.length) st.truncated = true; break; }
      st.resyncs++; st.skippedBytes += q - p; p = q;
      continue;
    }
    const data = s.subarray(p + 4, p + 4 + len);
    if (id === R.CONTINUE && out.length) out[out.length - 1].conts.push(data);
    else out.push({ id, off: p, data, conts: [] });
    p += 4 + len;
  }
  return out;
}

// ------------------------------------------------------------------------------------------
// Reader state

interface Font { name: string; size: number; bold: boolean; italic: boolean; underline: number; strike: boolean; color: number; script: number }
interface Xf { font: number; fmt: number; locked: boolean; hidden: boolean; h: number; v: number; wrap: boolean; shrink: boolean; indent: number; rot: number;
  bl: number; br: number; bt: number; bb: number; cl: number; cr: number; ct: number; cb: number; pattern: number; fore: number; back: number }
interface SupBook { self: boolean; addin: boolean; path: string; sheets: string[]; names: string[] }
interface NameRec { name: string; itab: number; hidden: boolean; builtin: boolean; rgce: Uint8Array; extra: Uint8Array; text?: string | null }
interface BoundSheet { name: string; pos: number; state: number; type: number }
interface SharedF { r1: number; r2: number; c1: number; c2: number; rgce: Uint8Array; extra: Uint8Array; array: boolean }

class Reader {
  ver = 8;
  dec8: (b: Uint8Array) => string = makeDecoder(1252);
  codepage = 1252;
  sst: string[] = [];
  fonts: Font[] = [];
  formats = new Map<number, string>();
  formatSeq = 0;
  xfs: Xf[] = [];
  xfStyle = new Map<number, number>();
  palette = DEFAULT_PALETTE.slice();
  bound: BoundSheet[] = [];
  sup: SupBook[] = [];
  xti: { sup: number; first: number; last: number }[] = [];
  ext5: { name: string; names: string[] }[] = [];
  names: NameRec[] = [];
  warnings: string[] = [];
  formulas = 0;
  decoded = 0;
  cells = 0;
  book = new Workbook(false);
  constructor(public lenient: boolean) {}

  warn(msg: string): void { if (this.warnings.length < 200 && !this.warnings.includes(msg)) this.warnings.push(msg); }

  // ---- strings
  /** BIFF8 XLUnicodeString (cch16) or ShortXLUnicodeString (cch8); BIFF2-5 byte strings in the code page. */
  str(c: Cursor, cch8 = false): string {
    const n = cch8 ? c.u8() : c.u16();
    if (this.ver < 8) return this.dec8(c.bytes(n));
    const fl = c.u8();
    const runs = fl & 8 ? c.u16() : 0;
    const ext = fl & 4 ? c.u32() : 0;
    const s = c.chars(n, !!(fl & 1));
    if (runs) c.skip(4 * runs);
    if (ext) c.skip(ext);
    return s;
  }
  strAt(b: Uint8Array, off: number, cch8 = false): string { return this.str(new Cursor([b], off), cch8); }

  color(i: number): string | undefined {
    if (i === 0x7fff || i === 64 || i === 65 || i === 0x51 || i < 0) return undefined;
    // BIFF3-4 use 5-bit indexes: 24 and 25 are the system text and window colors (automatic).
    if (this.ver < 5 && (i === 24 || i === 25 || i === 31)) return undefined;
    const idx = i < 8 ? i : i - 8;
    const hex = this.palette[idx];
    return hex ? `#${hex.toLowerCase()}` : undefined;
  }

  // ---- styles
  styleFor(ixfe: number): number {
    if (ixfe <= 0 && !this.xfs[0]) return 0;
    const hit = this.xfStyle.get(ixfe);
    if (hit != null) return hit;
    const xf = this.xfs[ixfe];
    let id = 0;
    if (xf) {
      const st: CellStyle = {};
      const fi = xf.font >= 4 && this.ver >= 5 ? xf.font - 1 : xf.font;
      const f = this.fonts[fi], f0 = this.fonts[0];
      if (f) {
        if (!f0 || f.name !== f0.name) st.font = f.name;
        if (!f0 || f.size !== f0.size) st.size = f.size;
        if (f.bold) st.bold = true;
        if (f.italic) st.italic = true;
        if (f.strike) st.strike = true;
        if (f.underline === 1) st.underline = 'single'; else if (f.underline === 2) st.underline = 'double';
        else if (f.underline === 0x21) st.underline = 'singleAccounting'; else if (f.underline === 0x22) st.underline = 'doubleAccounting';
        if (f.script === 1) st.vertAlign = 'superscript'; else if (f.script === 2) st.vertAlign = 'subscript';
        const col = this.color(f.color);
        if (col && col !== '#000000') st.color = col;
      }
      const fmt = this.formats.get(xf.fmt) ?? BUILTIN_FORMATS[xf.fmt] ?? EXTRA_BUILTIN_FORMATS[xf.fmt];
      if (fmt && fmt.toLowerCase() !== 'general') st.numFmt = fmt;
      if (xf.h > 0 && H_ALIGN[xf.h]) st.hAlign = H_ALIGN[xf.h];
      if (xf.v !== 2 && V_ALIGN[xf.v]) st.vAlign = V_ALIGN[xf.v];
      if (xf.wrap) st.wrap = true;
      if (xf.shrink) st.shrink = true;
      if (xf.indent) st.indent = xf.indent;
      if (xf.rot) st.rotate = xf.rot === 255 ? 255 : xf.rot <= 90 ? xf.rot : 90 - xf.rot;
      if (!xf.locked) st.locked = false;
      if (xf.hidden) st.hidden = true;
      const bd = (s: number, c: number): Border | undefined => { const name = BORDER_STYLES[s]; if (!name) return undefined; const col = this.color(c); return col && col !== '#000000' ? { style: name, color: col } : { style: name }; };
      const bl = bd(xf.bl, xf.cl), br = bd(xf.br, xf.cr), bt = bd(xf.bt, xf.ct), bb = bd(xf.bb, xf.cb);
      if (bl) st.bl = bl; if (br) st.br = br; if (bt) st.bt = bt; if (bb) st.bb = bb;
      if (xf.pattern === 1) { const c = this.color(xf.fore); if (c) st.fill = c; }
      else if (xf.pattern > 1) {
        const fore = this.color(xf.fore), back = this.color(xf.back);
        st.pattern = { type: PATTERNS[xf.pattern] || 'gray125', ...(fore ? { color: fore } : {}) };
        if (back) st.fill = back;
      }
      id = Object.keys(st).length ? this.book.styleId(st) : 0;
    }
    this.xfStyle.set(ixfe, id);
    return id;
  }

  attrCache = new Map<number, number>();
  /** Style from BIFF2 cell attributes: XF index (or font, format and alignment given directly). */
  attrStyle(b0: number, b1: number, b2: number): number {
    const key = (b0 << 16) | (b1 << 8) | b2;
    const hit = this.attrCache.get(key);
    if (hit != null) return hit;
    const xf: Xf = { font: (b1 >> 6) & 3, fmt: b1 & 0x3f, locked: !(b0 & 0x40), hidden: !!(b0 & 0x80), h: b2 & 7, v: 2, wrap: false, shrink: false, indent: 0, rot: 0,
      bl: b2 & 0x08 ? 1 : 0, br: b2 & 0x10 ? 1 : 0, bt: b2 & 0x20 ? 1 : 0, bb: b2 & 0x40 ? 1 : 0, cl: 0x7fff, cr: 0x7fff, ct: 0x7fff, cb: 0x7fff,
      pattern: b2 & 0x80 ? 17 : 0, fore: 0x7fff, back: 0x7fff };
    const ix = 0x100000 + key;
    this.xfs[ix] = xf;
    const id = this.styleFor(ix);
    delete this.xfs[ix];
    this.attrCache.set(key, id);
    return id;
  }

  // ---- globals records
  globals(rec: Rec): void {
    const d = rec.data;
    switch (rec.id) {
      case R.FILEPASS: throw new Error(XLS_ENCRYPTED_MSG);
      case R.CODEPAGE: {
        this.codepage = u16(d, 0);
        this.dec8 = makeDecoder(this.codepage);
        break;
      }
      case R.DATEMODE: this.book.calc.date1904 = u16(d, 0) === 1; break;
      case R.SST: {
        const c = new Cursor([d, ...rec.conts], 8);
        const unique = u32(d, 4);
        try {
          for (let i = 0; i < unique && !c.eof(); i++) this.sst.push(this.str(c));
        } catch (e) {
          if (!this.lenient) throw e;
          this.warn(`The shared string table is damaged; ${this.sst.length} of ${unique} strings were read.`);
        }
        if (this.sst.length < unique && this.lenient) this.warn(`The shared string table is short: ${this.sst.length} of ${unique} strings.`);
        break;
      }
      case R.FONT: case R.FONT34: {
        const size = u16(d, 0) / 20;
        const gr = u16(d, 2);
        let f: Font;
        if (this.ver >= 5) {
          f = { size, italic: !!(gr & 2), strike: !!(gr & 8), color: u16(d, 4), bold: u16(d, 6) >= 700, script: u16(d, 8), underline: d[10], name: this.strAt(d, 14, true) };
        } else if (this.ver >= 3) {
          f = { size, bold: !!(gr & 1), italic: !!(gr & 2), underline: gr & 4 ? 1 : 0, strike: !!(gr & 8), color: u16(d, 4), script: 0, name: this.dec8(d.subarray(7, 7 + d[6])) };
        } else {
          f = { size, bold: !!(gr & 1), italic: !!(gr & 2), underline: gr & 4 ? 1 : 0, strike: !!(gr & 8), color: 0x7fff, script: 0, name: this.dec8(d.subarray(5, 5 + d[4])) };
        }
        this.fonts.push(f);
        break;
      }
      case R.FORMAT: case R.FORMAT_OLD: {
        let idx: number, code: string;
        if (this.ver >= 8) { idx = u16(d, 0); code = this.strAt(d, 2); }
        else if (this.ver >= 5) { idx = u16(d, 0); code = this.dec8(d.subarray(3, 3 + d[2])); }
        else if (this.ver === 4) { idx = this.formatSeq++; code = this.dec8(d.subarray(3, 3 + d[2])); }
        else { idx = this.formatSeq++; code = this.dec8(d.subarray(1, 1 + d[0])); }
        this.formats.set(idx, code);
        if (!Object.values(BUILTIN_FORMATS).includes(code) && !this.book.customFormats.includes(code) && code.toLowerCase() !== 'general') this.book.customFormats.push(code);
        break;
      }
      case R.XF: case R.XF2: case R.XF3: case R.XF4: this.xfs.push(this.parseXf(rec)); break;
      case R.PALETTE: {
        const n = u16(d, 0);
        for (let i = 0; i < n && 2 + i * 4 + 3 <= d.length; i++) {
          const o = 2 + i * 4;
          this.palette[i] = [d[o], d[o + 1], d[o + 2]].map((x) => x.toString(16).padStart(2, '0')).join('').toUpperCase();
        }
        this.xfStyle.clear();
        break;
      }
      case R.BOUNDSHEET: {
        const name = this.ver >= 8 ? this.strAt(d, 6, true) : this.dec8(d.subarray(7, 7 + d[6]));
        this.bound.push({ pos: u32(d, 0), state: d[4] & 3, type: d[5], name });
        break;
      }
      case R.SUPBOOK: {
        const ctab = u16(d, 0), cch = u16(d, 2);
        if (cch === 0x0401) this.sup.push({ self: true, addin: false, path: '', sheets: [], names: [] });
        else if (cch === 0x3a01) this.sup.push({ self: false, addin: true, path: '', sheets: [], names: [] });
        else {
          const c = new Cursor([d, ...rec.conts], 2);
          let path = '';
          const sheets: string[] = [];
          try { path = this.str(c); for (let i = 0; i < ctab; i++) sheets.push(this.str(c)); } catch { /* partial */ }
          this.sup.push({ self: false, addin: false, path: cleanPath(path), sheets, names: [] });
        }
        break;
      }
      case R.EXTERNNAME: {
        if (this.ver >= 8) {
          const sb = this.sup[this.sup.length - 1];
          let nm = '';
          try { nm = this.strAt(d, 6, true); } catch { /* ignore */ }
          if (sb) sb.names.push(nm);
        } else {
          const e = this.ext5[this.ext5.length - 1];
          if (e) e.names.push(this.dec8(d.subarray(7, 7 + d[6])));
        }
        break;
      }
      case R.EXTERNSHEET: {
        if (this.ver >= 8) {
          const c = new Cursor([d, ...rec.conts]);
          try {
            const n = c.u16();
            for (let i = 0; i < n; i++) this.xti.push({ sup: c.u16(), first: c.u16(), last: c.u16() });
          } catch { /* partial */ }
        } else {
          const n = d[0];
          const raw = d.subarray(1, 1 + n);
          let name = this.dec8(raw);
          if (raw[0] === 3 || raw[0] === 4 || raw[0] === 1) name = this.dec8(raw.subarray(1));
          this.ext5.push({ name, names: [] });
        }
        break;
      }
      case R.NAME: case 0x0218: if (this.ver >= 5) this.parseName(rec); break;
      case R.PROTECT: if (u16(d, 0)) this.book.structureLocked = true; break;
      case R.PASSWORD: if (u16(d, 0)) this.book.structurePassword = u16(d, 0).toString(16).toUpperCase(); break;
    }
  }

  parseXf(rec: Rec): Xf {
    const d = rec.data;
    const xf: Xf = { font: 0, fmt: 0, locked: true, hidden: false, h: 0, v: 2, wrap: false, shrink: false, indent: 0, rot: 0, bl: 0, br: 0, bt: 0, bb: 0, cl: 0x7fff, cr: 0x7fff, ct: 0x7fff, cb: 0x7fff, pattern: 0, fore: 0x7fff, back: 0x7fff };
    if (this.ver >= 8) {
      xf.font = u16(d, 0); xf.fmt = u16(d, 2);
      const t = u16(d, 4); xf.locked = !!(t & 1); xf.hidden = !!(t & 2);
      const al = d[6]; xf.h = al & 7; xf.wrap = !!(al & 8); xf.v = (al >> 4) & 7;
      xf.rot = d[7]; xf.indent = d[8] & 15; xf.shrink = !!(d[8] & 16);
      const b1 = u32(d, 10), b2 = u32(d, 14), col = u16(d, 18);
      xf.bl = b1 & 15; xf.br = (b1 >>> 4) & 15; xf.bt = (b1 >>> 8) & 15; xf.bb = (b1 >>> 12) & 15;
      xf.cl = (b1 >>> 16) & 0x7f; xf.cr = (b1 >>> 23) & 0x7f; xf.ct = b2 & 0x7f; xf.cb = (b2 >>> 7) & 0x7f;
      xf.pattern = (b2 >>> 26) & 0x3f; xf.fore = col & 0x7f; xf.back = (col >>> 7) & 0x7f;
    } else if (this.ver >= 5) {
      xf.font = u16(d, 0); xf.fmt = u16(d, 2);
      const t = u16(d, 4); xf.locked = !!(t & 1); xf.hidden = !!(t & 2);
      const al = u16(d, 6); xf.h = al & 7; xf.wrap = !!(al & 8); xf.v = (al >> 4) & 7;
      const ori = (al >> 8) & 3; xf.rot = ori === 1 ? 255 : ori === 2 ? 90 : ori === 3 ? 180 : 0;
      const a = u32(d, 8), b = u32(d, 12);
      xf.fore = a & 0x7f; xf.back = (a >>> 7) & 0x7f; xf.pattern = (a >>> 16) & 0x3f; xf.bb = (a >>> 22) & 7; xf.cb = (a >>> 25) & 0x7f;
      xf.bt = b & 7; xf.bl = (b >>> 3) & 7; xf.br = (b >>> 6) & 7; xf.ct = (b >>> 9) & 0x7f; xf.cl = (b >>> 16) & 0x7f; xf.cr = (b >>> 23) & 0x7f;
    } else if (this.ver >= 3) {
      xf.font = d[0]; xf.fmt = d[1];
      xf.locked = !!(d[2] & 1); xf.hidden = !!(d[2] & 2);
      const al = this.ver === 4 ? d[4] : u16(d, 4);
      xf.h = al & 7; xf.wrap = !!(al & 8);
      if (this.ver === 4) { xf.v = (al >> 4) & 3; const ori = (al >> 6) & 3; xf.rot = ori === 1 ? 255 : ori === 2 ? 90 : ori === 3 ? 180 : 0; }
      if (d.length >= 12) {
        // Background: pattern (6 bits), foreground and background colors (5 bits each).
        const bg = u16(d, 6);
        xf.pattern = bg & 0x3f; xf.fore = (bg >> 6) & 0x1f; xf.back = (bg >> 11) & 0x1f;
        // Borders: style (3 bits) and color (5 bits) for top, left, bottom, right.
        const bd = u32(d, 8);
        xf.bt = bd & 7; xf.ct = (bd >>> 3) & 0x1f; xf.bl = (bd >>> 8) & 7; xf.cl = (bd >>> 11) & 0x1f;
        xf.bb = (bd >>> 16) & 7; xf.cb = (bd >>> 19) & 0x1f; xf.br = (bd >>> 24) & 7; xf.cr = (bd >>> 27) & 0x1f;
      }
    } else {
      xf.font = d[0]; xf.fmt = d[2] & 0x3f; xf.h = d[3] & 7;
      xf.locked = !(d[2] & 0x40); xf.hidden = !!(d[2] & 0x80);
    }
    return xf;
  }

  parseName(rec: Rec): void {
    const d = rec.data;
    const gr = u16(d, 0);
    const cch = d[3];
    const cce = u16(d, 4);
    let name: string, rgceOff: number, itab: number;
    if (this.ver >= 8) {
      itab = u16(d, 8);
      const c = new Cursor([d, ...rec.conts], 14);
      const fl = c.u8();
      name = c.chars(cch, !!(fl & 1));
      rgceOff = c.pi === 0 ? c.p : d.length;
    } else if (this.ver >= 5) {
      itab = u16(d, 8);
      name = this.dec8(d.subarray(14, 14 + cch));
      rgceOff = 14 + cch;
    } else {
      itab = 0;
      const off = this.ver === 2 ? 5 : 6;
      name = this.dec8(d.subarray(off, off + cch));
      rgceOff = off + cch;
    }
    const builtin = !!(gr & 0x20);
    if (builtin && name.length >= 1) {
      const code = name.charCodeAt(0);
      name = BUILTIN_NAMES[code] ?? name;
    }
    if (name.startsWith('_xlnm.')) name = name.slice(6);
    this.names.push({ name, itab, hidden: !!(gr & 1), builtin, rgce: d.subarray(rgceOff, rgceOff + cce), extra: d.subarray(rgceOff + cce) });
  }

  // ---- formula decoding

  sheetRef(i: number): string { return quoteSheet(this.bound[i]?.name ?? `Sheet${i + 1}`); }

  /** Sheet prefix for a BIFF8 XTI index. */
  xtiPrefix(ixti: number): string | null {
    const x = this.xti[ixti];
    if (!x) return null;
    const sb = this.sup[x.sup];
    if (!sb || sb.self) {
      if (x.first === 0xfffe || x.first === 0xffff) return null;
      const a = this.bound[x.first]?.name ?? `Sheet${x.first + 1}`;
      const b = this.bound[x.last]?.name ?? a;
      return x.first === x.last ? quoteSheet(a) : quoteSheet(`${a}:${b}`);
    }
    const sh = sb.sheets[x.first] ?? '';
    const file = sb.path.replace(/^.*[\\/]/, '');
    return `'[${file.replace(/'/g, "''")}]${sh.replace(/'/g, "''")}'`;
  }

  /** Decode a parsed expression into formula text (no leading '='). Throws on unknown tokens. */
  decode(rgce: Uint8Array, extra: Uint8Array, r: number, c: number, shared: boolean): string {
    const v = this.ver, b8 = v >= 8;
    const st: string[] = [];
    const xc = new Cursor([extra]);
    const pop = () => { const x = st.pop(); if (x === undefined) throw new Error('formula stack underflow'); return x; };
    const maxRow = b8 ? 0xffff : 0x3fff;
    const ref = (row: number, col: number, rowRel: boolean, colRel: boolean, relative: boolean): string => {
      if (relative) {
        if (rowRel) row = b8 ? (r + (row << 16 >> 16)) & 0xffff : (r + (row << 18 >> 18)) & 0x3fff;
        if (colRel) col = (c + (col << 24 >> 24)) & 0xff;
      }
      if (row > MAX_ROWS - 1 || col > MAX_COLS - 1) throw new Error('reference out of range');
      return `${colRel ? '' : '$'}${colName(col)}${rowRel ? '' : '$'}${row + 1}`;
    };
    const readRef = (o: number, relative: boolean): [string, number] => {
      if (b8) { const row = u16(rgce, o), cf = u16(rgce, o + 2); return [ref(row, cf & 0x3fff, !!(cf & 0x8000), !!(cf & 0x4000), relative), 4]; }
      const rf = u16(rgce, o); return [ref(rf & 0x3fff, rgce[o + 2], !!(rf & 0x8000), !!(rf & 0x4000), relative), 3];
    };
    const readArea = (o: number, relative: boolean): [string, number] => {
      let r1: number, r2: number, c1: number, c2: number, r1r: boolean, r2r: boolean, c1r: boolean, c2r: boolean, n: number;
      if (b8) {
        r1 = u16(rgce, o); r2 = u16(rgce, o + 2); const f1 = u16(rgce, o + 4), f2 = u16(rgce, o + 6);
        c1 = f1 & 0x3fff; c2 = f2 & 0x3fff; r1r = !!(f1 & 0x8000); c1r = !!(f1 & 0x4000); r2r = !!(f2 & 0x8000); c2r = !!(f2 & 0x4000); n = 8;
      } else {
        const a = u16(rgce, o), b = u16(rgce, o + 2);
        r1 = a & 0x3fff; r2 = b & 0x3fff; r1r = !!(a & 0x8000); c1r = !!(a & 0x4000); r2r = !!(b & 0x8000); c2r = !!(b & 0x4000);
        c1 = rgce[o + 4]; c2 = rgce[o + 5]; n = 6;
      }
      if (!relative && r1 === 0 && r2 === maxRow) return [`${c1r ? '' : '$'}${colName(c1)}:${c2r ? '' : '$'}${colName(c2)}`, n];
      if (!relative && c1 === 0 && c2 === 0xff) return [`${r1r ? '' : '$'}${r1 + 1}:${r2r ? '' : '$'}${r2 + 1}`, n];
      return [`${ref(r1, c1, r1r, c1r, relative)}:${ref(r2, c2, r2r, c2r, relative)}`, n];
    };
    const nameText = (idx: number): string => {
      const nm = this.names[idx - 1];
      if (!nm) throw new Error('unknown name');
      return nm.name;
    };
    const readArrayConst = (): string => {
      // BIFF8 stores the last column and row index; earlier versions store the counts, with 0 columns meaning 256.
      const nc = xc.u8(), nr = xc.u16();
      const cols = b8 ? nc + 1 : nc || 256;
      const rows = b8 ? nr + 1 : nr;
      const out: string[] = [];
      for (let i = 0; i < rows; i++) {
        const row: string[] = [];
        for (let j = 0; j < cols; j++) {
          const t = xc.u8();
          if (t === 0) { xc.skip(8); row.push(''); }
          else if (t === 1) { const b = xc.bytes(8); row.push(numText(f64(b, 0))); }
          else if (t === 2) { const s = b8 ? this.str(xc) : this.dec8(xc.bytes(xc.u8())); row.push(`"${s.replace(/"/g, '""')}"`); }
          else if (t === 4) { const b = xc.bytes(8); row.push(b[0] ? 'TRUE' : 'FALSE'); }
          else if (t === 16) { const b = xc.bytes(8); row.push(ERR_CODES[b[0]] || '#N/A'); }
          else throw new Error('bad array constant');
        }
        out.push(row.join(','));
      }
      return `{${out.join(';')}}`;
    };
    let i = 0;
    while (i < rgce.length) {
      const ptg = rgce[i++];
      if (ptg >= 0x03 && ptg <= 0x11) {
        const b = pop(), a = pop();
        const op = OPS[ptg];
        st.push(ptg === 0x10 ? `(${a},${b})` : `${a}${op}${b}`);
        continue;
      }
      switch (ptg) {
        case 0x01: case 0x02: throw new Error('ptgExp/ptgTbl outside a cell formula');
        case 0x12: st.push(`+${pop()}`); continue;
        case 0x13: st.push(`-${pop()}`); continue;
        case 0x14: st.push(`${pop()}%`); continue;
        case 0x15: st.push(`(${pop()})`); continue;
        case 0x16: st.push(''); continue;
        case 0x17: {
          let s: string;
          if (b8) { const cur = new Cursor([rgce], i); s = this.str(cur, true); i = cur.p; }
          else { const n = rgce[i]; s = this.dec8(rgce.subarray(i + 1, i + 1 + n)); i += 1 + n; }
          st.push(`"${s.replace(/"/g, '""')}"`);
          continue;
        }
        case 0x18: throw new Error('extended token');
        case 0x19: {
          const gr = rgce[i];
          if (gr & 0x04) { const n = u16(rgce, i + 1); i += 3 + 2 * (n + 1); continue; }
          i += v >= 3 ? 3 : 2;
          if (gr & 0x10) st.push(`SUM(${pop()})`);
          continue;
        }
        case 0x1c: st.push(ERR_CODES[rgce[i]] || '#N/A'); i += 1; continue;
        case 0x1d: st.push(rgce[i] ? 'TRUE' : 'FALSE'); i += 1; continue;
        case 0x1e: st.push(String(u16(rgce, i))); i += 2; continue;
        case 0x1f: st.push(numText(f64(rgce, i))); i += 8; continue;
      }
      if (ptg < 0x20 || ptg > 0x7f) throw new Error(`unknown token 0x${ptg.toString(16)}`);
      const t = (ptg & 0x1f) | 0x20;
      switch (t) {
        case 0x20: i += 7; st.push(readArrayConst()); break;
        case 0x21: {
          const idx = v >= 4 ? u16(rgce, i) : rgce[i]; i += v >= 4 ? 2 : 1;
          const fn = BIFF_FUNCTIONS.get(idx);
          if (!fn || fn[1] < 0) throw new Error(`unknown fixed function ${idx}`);
          const args: string[] = [];
          for (let k = 0; k < fn[1]; k++) args.unshift(pop());
          st.push(`${fn[0]}(${args.join(',')})`);
          break;
        }
        case 0x22: {
          const n = rgce[i] & 0x7f;
          const idx = (v >= 4 ? u16(rgce, i + 1) : rgce[i + 1]) & 0x7fff; i += v >= 4 ? 3 : 2;
          const args: string[] = [];
          for (let k = 0; k < n; k++) args.unshift(pop());
          if (idx === 255) {
            const nm = (args.shift() || '').replace(/^_xlfn\./i, '').replace(/^_xlws\./i, '');
            st.push(`${nm}(${args.join(',')})`);
          } else {
            const fn = BIFF_FUNCTIONS.get(idx);
            if (!fn) throw new Error(`unknown function ${idx}`);
            st.push(`${fn[0]}(${args.join(',')})`);
          }
          break;
        }
        case 0x23: { const idx = u16(rgce, i); i += b8 ? 4 : v >= 5 ? 14 : v >= 3 ? 10 : 7; st.push(nameText(idx)); break; }
        case 0x24: { const [s, n] = readRef(i, false); i += n; st.push(s); break; }
        case 0x25: { const [s, n] = readArea(i, false); i += n; st.push(s); break; }
        case 0x26: i += b8 ? 6 : 6; if (b8) { const n = xc.u16(); xc.skip(n * 8); } break;
        case 0x27: case 0x28: i += b8 ? 6 : 6; break;
        case 0x29: i += 2; break;
        case 0x2a: i += b8 ? 4 : 3; st.push('#REF!'); break;
        case 0x2b: i += b8 ? 8 : 6; st.push('#REF!'); break;
        case 0x2c: { const [s, n] = readRef(i, true); i += n; st.push(s); break; }
        case 0x2d: { const [s, n] = readArea(i, true); i += n; st.push(s); break; }
        case 0x2e: case 0x2f: i += 2; break;
        case 0x39: {
          if (b8) {
            const ixti = u16(rgce, i), ilbl = u16(rgce, i + 2); i += 6;
            const x = this.xti[ixti], sb = x ? this.sup[x.sup] : undefined;
            if (!sb || sb.self) st.push(nameText(ilbl));
            else {
              const nm = sb.names[ilbl - 1];
              if (nm == null) throw new Error('unknown external name');
              st.push(sb.addin ? nm : `[${sb.path.replace(/^.*[\\/]/, '')}]${nm}`);
            }
          } else {
            const ixals = i16(rgce, i), ilbl = u16(rgce, i + 10); i += 24;
            if (ixals < 0) st.push(nameText(ilbl));
            else { const nm = this.ext5[ixals - 1]?.names[ilbl - 1]; if (nm == null) throw new Error('unknown external name'); st.push(nm); }
          }
          break;
        }
        case 0x3a: case 0x3b: case 0x3c: case 0x3d: {
          let prefix: string | null;
          if (b8) { prefix = this.xtiPrefix(u16(rgce, i)); i += 2; }
          else {
            const ixals = i16(rgce, i), f = i16(rgce, i + 10), l = i16(rgce, i + 12); i += 14;
            if (f >= 0) prefix = f === l || l < 0 ? this.sheetRef(f) : quoteSheet(`${this.bound[f]?.name ?? `Sheet${f + 1}`}:${this.bound[l]?.name ?? `Sheet${l + 1}`}`);
            else { const e = this.ext5[Math.abs(ixals) - 1]; prefix = e ? quoteSheet(e.name) : null; }
          }
          if (t === 0x3a) { const [s, n] = readRef(i, shared); i += n; st.push(prefix ? `${prefix}!${s}` : '#REF!'); }
          else if (t === 0x3b) { const [s, n] = readArea(i, shared); i += n; st.push(prefix ? `${prefix}!${s}` : '#REF!'); }
          else { i += t === 0x3c ? (b8 ? 4 : 3) : (b8 ? 8 : 6); st.push(prefix ? `${prefix}!#REF!` : '#REF!'); }
          break;
        }
        default: throw new Error(`unknown token 0x${ptg.toString(16)}`);
      }
    }
    if (st.length !== 1) throw new Error('unbalanced formula');
    return st[0];
  }
}

const OPS: Record<number, string> = { 3: '+', 4: '-', 5: '*', 6: '/', 7: '^', 8: '&', 9: '<', 10: '<=', 11: '=', 12: '>=', 13: '>', 14: '<>', 15: ' ', 16: ',', 17: ':' };

function numText(n: number): string {
  if (!Number.isFinite(n)) return '#NUM!';
  const s = String(n);
  return s.includes('e') ? n.toExponential().toUpperCase().replace('E+', 'E+') : s;
}

function cleanPath(p: string): string {
  // Encoded file paths: \x01 = volume, \x02 = same volume, \x03 = down a directory, \x04 = up a directory.
  return p.replace(/^[\x01\x02]/, '').replace(/\x03/g, '/').replace(/\x04/g, '../').replace(/[\x00-\x1f]/g, '');
}

export function quoteSheet(name: string): string {
  return /^[A-Za-z_À-￿][\w.À-￿]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name) && !/^(R\d*C\d*|TRUE|FALSE)$/i.test(name)
    ? name : `'${name.replace(/'/g, "''")}'`;
}

// ------------------------------------------------------------------------------------------
// Sheet substreams

interface SheetCtx {
  sheet: Sheet;
  shared: Map<string, SharedF>;
  pendingExp: { r: number; c: number; br: number; bc: number; cell: Cell }[];
  lastFormula: Cell | null;
  objText: Map<number, string>;
  lastObj: number;
  notes: { r: number; c: number; id: number; author: string; text?: string }[];
  note5: { r: number; c: number; text: string } | null;
  frozen: boolean;
}

function processSheetRec(rd: Reader, ctx: SheetCtx, rec: Rec): void {
  const d = rec.data, v = rd.ver, sh = ctx.sheet;
  const b2 = v === 2;
  const put = (r: number, c: number, cell: Cell) => {
    if (r >= MAX_ROWS || c >= MAX_COLS) return;
    if (cell.s === 0) delete cell.s;
    sh.setComputed(r, c, cell);
    if (cell.v != null || cell.f != null) rd.cells++;
  };
  const sid = (ixfe: number) => rd.styleFor(ixfe);
  /** BIFF2 cells carry 3 attribute bytes (XF index, font and format, alignment and borders) instead of an XF number. */
  const a2 = (o = 4) => rd.attrStyle(d[o], d[o + 1], d[o + 2]);
  switch (rec.id) {
    case R.LABELSST: {
      const idx = u32(d, 6);
      const s = rd.sst[idx];
      if (s === undefined) { rd.warn('Some cells refer to shared strings that were lost.'); put(u16(d, 0), u16(d, 2), { v: null, s: sid(u16(d, 4)) }); break; }
      put(u16(d, 0), u16(d, 2), { v: s, s: sid(u16(d, 4)) });
      break;
    }
    case R.LABEL: case R.RSTRING: {
      const s = rd.str(new Cursor([d, ...rec.conts], 6));
      put(u16(d, 0), u16(d, 2), { v: s, s: sid(u16(d, 4)) });
      break;
    }
    case R.LABEL2: put(u16(d, 0), u16(d, 2), { v: rd.dec8(d.subarray(8, 8 + d[7])), s: a2() }); break;
    case R.BLANK2: { const s = a2(); if (s) put(u16(d, 0), u16(d, 2), { s }); break; }
    case R.NUMBER: put(u16(d, 0), u16(d, 2), { v: f64(d, 6), s: sid(u16(d, 4)) }); break;
    case R.NUMBER2: put(u16(d, 0), u16(d, 2), { v: f64(d, 7), s: a2() }); break;
    case R.INTEGER2: put(u16(d, 0), u16(d, 2), { v: u16(d, 7), s: a2() }); break;
    case R.RK: put(u16(d, 0), u16(d, 2), { v: rk(u32(d, 6)), s: sid(u16(d, 4)) }); break;
    case R.MULRK: {
      const r = u16(d, 0), c0 = u16(d, 2);
      const n = Math.floor((d.length - 6) / 6);
      for (let k = 0; k < n; k++) put(r, c0 + k, { v: rk(u32(d, 4 + k * 6 + 2)), s: sid(u16(d, 4 + k * 6)) });
      break;
    }
    case R.BOOLERR: case R.BOOLERR2: {
      const o = rec.id === R.BOOLERR2 ? 7 : 6;
      const val: Scalar = d[o + 1] ? errValue(d[o]) : !!d[o];
      put(u16(d, 0), u16(d, 2), { v: val, s: b2 ? a2() : sid(u16(d, 4)) });
      break;
    }
    case R.BLANK: { const s = sid(u16(d, 4)); if (s) put(u16(d, 0), u16(d, 2), { s }); break; }
    case R.MULBLANK: {
      const r = u16(d, 0), c0 = u16(d, 2);
      const n = Math.floor((d.length - 6) / 2);
      for (let k = 0; k < n; k++) { const s = sid(u16(d, 4 + 2 * k)); if (s) put(r, c0 + k, { s }); }
      break;
    }
    case R.FORMULA: case R.FORMULA3: case R.FORMULA4: {
      const r = u16(d, 0), c = u16(d, 2);
      const numOff = b2 ? 7 : 6;
      let val: Scalar = null;
      let wantString = false;
      if (d[numOff + 6] === 0xff && d[numOff + 7] === 0xff) {
        const t = d[numOff];
        if (t === 0) wantString = true;
        else if (t === 1) val = !!d[numOff + 2];
        else if (t === 2) val = errValue(d[numOff + 2]);
        else if (t === 3) val = '';
      } else val = f64(d, numOff);
      let cce: number, off: number;
      if (b2) { cce = d[16]; off = 17; }
      else if (v >= 5) { cce = u16(d, 20); off = 22; }
      else { cce = u16(d, 16); off = 18; }
      const rgce = d.subarray(off, off + cce), extra = d.subarray(off + cce);
      const cell: Cell = { v: val, s: b2 ? a2() : sid(u16(d, 4)) };
      rd.formulas++;
      if (rgce[0] === 0x01 && rgce.length === 5) {
        ctx.pendingExp.push({ r, c, br: u16(rgce, 1), bc: v >= 8 ? u16(rgce, 3) : rgce[3], cell });
      } else if (rgce[0] === 0x02) {
        // Data table (TABLE): only the cached value can be kept.
      } else {
        try { cell.f = rd.decode(rgce, extra, r, c, false); rd.decoded++; }
        catch (e: any) { rd.warn(`Some formulas could not be decoded and keep their last calculated values (${e?.message || e}).`); }
      }
      put(r, c, cell);
      ctx.lastFormula = wantString ? cell : null;
      if (wantString) cell.v = '';
      break;
    }
    case R.STRING: case R.STRING2: {
      if (ctx.lastFormula) {
        ctx.lastFormula.v = rec.id === R.STRING2 ? rd.dec8(d.subarray(1, 1 + d[0])) : rd.str(new Cursor([d, ...rec.conts]));
        ctx.lastFormula = null;
      }
      break;
    }
    case R.SHRFMLA: case R.ARRAY: case R.ARRAY2: {
      const r1 = u16(d, 0), r2 = u16(d, 2), c1 = d[4], c2 = d[5];
      let cce: number, off: number;
      if (rec.id === R.SHRFMLA) { cce = u16(d, 8); off = 10; }
      else if (v >= 5) { cce = u16(d, 12); off = 14; }
      else if (v >= 3) { cce = u16(d, 8); off = 10; }
      else { cce = d[7]; off = 8; }
      ctx.shared.set(`${r1},${c1}`, { r1, r2, c1, c2, rgce: d.subarray(off, off + cce), extra: d.subarray(off + cce), array: rec.id !== R.SHRFMLA });
      break;
    }
    case R.COLINFO: {
      const c1 = u16(d, 0), c2 = Math.min(u16(d, 2), 255), w = u16(d, 4), gr = u16(d, 8);
      const px = Math.max(0, Math.round((w / 256) * 7));
      for (let c = c1; c <= c2; c++) {
        if (w && Math.abs(px - sh.defaultColW) > 1) sh.colW.set(c, px);
        if (gr & 1) sh.hiddenCols.add(c);
        const lvl = (gr >> 8) & 7;
        if (lvl) sh.colLevel.set(c, lvl);
        const s = sid(u16(d, 6));
        if (s) sh.colStyle.set(c, s);
      }
      break;
    }
    case R.COLWIDTH2: {
      const c1 = d[0], c2 = d[1], w = u16(d, 2);
      for (let c = c1; c <= c2; c++) sh.colW.set(c, Math.round((w / 256) * 7));
      break;
    }
    case R.DEFCOLWIDTH: { const w = u16(d, 0); if (w && w !== 8) sh.defaultColW = Math.round(w * 7 + 5); break; }
    case R.ROW: case R.ROW2: {
      const r = u16(d, 0), ht = u16(d, 6);
      if (v >= 5) {
        const gr = u16(d, 12);
        if (gr & 0x20) sh.hiddenRows.add(r);
        if ((gr & 0x40) && !(ht & 0x8000)) sh.rowH.set(r, Math.round(((ht & 0x7fff) / 20) * (4 / 3)));
        const lvl = gr & 7;
        if (lvl) sh.rowLevel.set(r, lvl);
        if (gr & 0x80) { const s = sid(u16(d, 14) & 0xfff); if (s) sh.rowStyle.set(r, s); }
      } else if (!(ht & 0x8000)) sh.rowH.set(r, Math.round(((ht & 0x7fff) / 20) * (4 / 3)));
      break;
    }
    case R.MERGEDCELLS: {
      const n = u16(d, 0);
      for (let k = 0; k < n && 2 + k * 8 + 8 <= d.length; k++) {
        const o = 2 + k * 8;
        sh.merges.push({ r1: u16(d, o), r2: u16(d, o + 2), c1: u16(d, o + 4), c2: u16(d, o + 6) });
      }
      break;
    }
    case R.WINDOW2: {
      if (v < 5) break;
      const gr = u16(d, 0);
      sh.showFormulas = !!(gr & 1);
      sh.showGrid = !!(gr & 2);
      sh.showHeadings = !!(gr & 4);
      sh.showZeros = !!(gr & 0x10);
      sh.rtl = !!(gr & 0x40);
      ctx.frozen = !!(gr & 8);
      break;
    }
    case R.PANE: {
      const x = u16(d, 0), y = u16(d, 2);
      if (ctx.frozen) sh.freeze = { rows: y, cols: x };
      break;
    }
    case R.SCL: { const n = u16(d, 0), dd = u16(d, 2); if (n && dd) sh.zoom = n / dd; break; }
    case R.PROTECT: if (u16(d, 0)) sh.protection = { ...sh.protection, on: true }; break;
    case R.PASSWORD: if (u16(d, 0)) sh.protection = { ...sh.protection, passwordHash: u16(d, 0).toString(16).toUpperCase() }; break;
    case R.OBJ: {
      if (v >= 8 && u16(d, 0) === 0x15) ctx.lastObj = u16(d, 6);
      break;
    }
    case R.TXO: {
      if (v < 8) break;
      const cch = u16(d, 10);
      if (cch && rec.conts.length) {
        try {
          const cur = new Cursor(rec.conts, 0);
          const high = !!(cur.u8() & 1);
          ctx.objText.set(ctx.lastObj, cur.chars(cch, high));
        } catch { /* comment text lost */ }
      }
      break;
    }
    case R.NOTE: {
      const r = u16(d, 0), c = u16(d, 2);
      if (v >= 8) {
        let author = '';
        try { author = rd.strAt(d, 8); } catch { /* ignore */ }
        ctx.notes.push({ r, c, id: u16(d, 6), author });
      } else {
        const n = u16(d, 4);
        const text = rd.dec8(d.subarray(6, 6 + n));
        if (r === 0xffff && ctx.note5) ctx.note5.text += text;
        else { ctx.note5 = { r, c, text }; ctx.notes.push({ r, c, id: -1, author: '', text: '' }); }
        const last = ctx.notes[ctx.notes.length - 1];
        if (last && ctx.note5) last.text = ctx.note5.text;
      }
      break;
    }
    case R.HLINK: {
      try {
        const link = parseHlink(d);
        if (!link) break;
        const r1 = u16(d, 0), r2 = u16(d, 2), c1 = u16(d, 4), c2 = u16(d, 6);
        for (let r = r1; r <= Math.min(r2, r1 + 1000); r++) for (let c = c1; c <= Math.min(c2, c1 + 255); c++) {
          const cur = sh.get(r, c);
          sh.setComputed(r, c, { ...(cur || { v: null }), link });
        }
      } catch { rd.warn('A hyperlink could not be read.'); }
      break;
    }
  }
}

function utf16(b: Uint8Array, o: number, chars: number): string {
  let s = '';
  for (let i = 0; i < chars && o + 2 * i + 1 < b.length; i++) s += String.fromCharCode(b[o + 2 * i] | (b[o + 2 * i + 1] << 8));
  return s.replace(/\0+$/, '');
}

const URL_CLSID = [0xe0, 0xc9, 0xea, 0x79, 0xf9, 0xba, 0xce, 0x11, 0x8c, 0x82, 0x00, 0xaa, 0x00, 0x4b, 0xa9, 0x0b];
const FILE_CLSID = [0x03, 0x03, 0, 0, 0, 0, 0, 0, 0xc0, 0, 0, 0, 0, 0, 0, 0x46];
const sameBytes = (b: Uint8Array, o: number, sig: number[]) => sig.every((x, i) => b[o + i] === x);

function parseHlink(d: Uint8Array): string | null {
  const flags = u32(d, 28);
  let p = 32;
  if (flags & 0x10) { const n = u32(d, p); p += 4 + n * 2; }
  if (flags & 0x80) { const n = u32(d, p); p += 4 + n * 2; }
  let target = '';
  if (flags & 0x01) {
    if (flags & 0x100) { const n = u32(d, p); target = utf16(d, p + 4, n); p += 4 + n * 2; }
    else if (sameBytes(d, p, URL_CLSID)) { const n = u32(d, p + 16); target = utf16(d, p + 20, n / 2); p += 20 + n; }
    else if (sameBytes(d, p, FILE_CLSID)) {
      p += 16;
      const up = u16(d, p); const an = u32(d, p + 2);
      let ansi = '';
      for (let i = 0; i < an; i++) { const ch = d[p + 6 + i]; if (!ch) break; ansi += String.fromCharCode(ch); }
      p += 6 + an + 24;
      const cb = u32(d, p); p += 4;
      let uni = '';
      if (cb > 0) { const nb = u32(d, p); uni = utf16(d, p + 6, nb / 2); p += cb; }
      target = '../'.repeat(up) + (uni || ansi).replace(/\\/g, '/');
    } else return null;
  }
  let loc = '';
  if (flags & 0x08) { const n = u32(d, p); loc = utf16(d, p + 4, n); }
  if (!target && !loc) return null;
  return target ? (loc ? `${target}#${loc}` : target) : `#${loc}`;
}

function finishSheet(rd: Reader, ctx: SheetCtx): void {
  const sh = ctx.sheet;
  for (const p of ctx.pendingExp) {
    const sf = ctx.shared.get(`${p.br},${p.bc}`);
    if (!sf) {
      // A shared or array formula whose SHRFMLA/ARRAY record was lost keeps its value.
      continue;
    }
    if (sf.array) {
      if (p.r === sf.r1 && p.c === sf.c1) {
        try { p.cell.f = rd.decode(sf.rgce, sf.extra, p.r, p.c, false); p.cell.arr = { r1: sf.r1, c1: sf.c1, r2: sf.r2, c2: sf.c2 }; rd.decoded++; }
        catch { rd.warn('Some array formulas could not be decoded and keep their values.'); }
      } else rd.formulas--;
    } else {
      try { p.cell.f = rd.decode(sf.rgce, sf.extra, p.r, p.c, true); rd.decoded++; }
      catch { rd.warn('Some shared formulas could not be decoded and keep their values.'); }
    }
    sh.setComputed(p.r, p.c, p.cell);
  }
  for (const n of ctx.notes) {
    const text = n.text ?? ctx.objText.get(n.id);
    if (text == null) continue;
    sh.comments.set(cellKey(n.r, n.c), { author: n.author, text, time: Date.now(), note: true });
  }
}

// ------------------------------------------------------------------------------------------
// Entry points

/** The BIFF record stream of a file: the Workbook/Book stream of a compound file, or a raw BIFF2-4 file. */
function findStream(u8: Uint8Array, warnings: string[]): { s: Uint8Array; source: 'cfb' | 'raw'; damaged: boolean; hint: number } | null {
  if (isOle2(u8)) {
    const cfb = readCfb(u8);
    if (!cfb) return null;
    let key = [...cfb.streams.keys()].find((k) => k.toLowerCase() === 'workbook');
    let hint = 8;
    if (!key) { key = [...cfb.streams.keys()].find((k) => k.toLowerCase() === 'book'); hint = 5; }
    if (!key) {
      // Works 6-9 (.xlr) and other programs keep the BIFF stream under another name.
      key = [...cfb.streams.keys()].find((k) => { const st = cfb.streams.get(k)!; return isBofId(u16(st, 0)) && u16(st, 2) >= 4 && u16(st, 2) <= 20; });
      hint = 8;
    }
    if (!key) {
      if ([...cfb.streams.keys()].some((k) => k === 'EncryptedPackage' || k === 'EncryptionInfo')) throw new Error(XLS_ENCRYPTED_MSG);
      return null;
    }
    const damaged = cfb.damaged.length > 0;
    if (damaged) warnings.push(`The compound file is damaged (${cfb.damaged.join(', ')}).`);
    return { s: cfb.streams.get(key)!, source: 'cfb', damaged, hint };
  }
  if (isBofId(u16(u8, 0)) && u16(u8, 2) >= 4 && u16(u8, 2) <= 20) return { s: u8, source: 'raw', damaged: false, hint: 0 };
  return null;
}

/** Offsets of BIFF BOF records in arbitrary bytes (used when the compound file cannot be read). */
export function scanBofs(u8: Uint8Array): number[] {
  const out: number[] = [];
  for (let i = 0; i + 8 <= u8.length; i++) {
    if (u8[i + 1] !== 0x08 || u8[i] !== 0x09) continue;
    const len = u16(u8, i + 2), vers = u16(u8, i + 4), dt = u16(u8, i + 6);
    if ((len === 16 && vers === 0x0600) || (len >= 8 && len <= 16 && vers === 0x0500)) {
      if (dt === 0x0005 || dt === 0x0010 || dt === 0x0020 || dt === 0x0040 || dt === 0x0006 || dt === 0x0100) out.push(i);
    }
  }
  return out;
}

function parseStream(s: Uint8Array, lenient: boolean, hint: number): { rd: Reader; walk: WalkStats; records: number } {
  const ws: WalkStats = { resyncs: 0, skippedBytes: 0, truncated: false };
  const rd = new Reader(lenient);
  let start = 0;
  if (lenient && !isBofId(u16(s, 0))) {
    const b = scanBofs(s);
    if (b.length) start = b[0];
  }
  const recs = walk(s, start, lenient, ws);
  if (ws.truncated) {
    if (!lenient) throw new Error('The workbook stream is truncated.');
    rd.warn('The workbook stream ends early; the data after the break is lost.');
  }
  let i = recs.findIndex((r) => isBofId(r.id));
  if (i < 0) throw new Error('No BIFF BOF record found.');
  const bof = recs[i];
  if (bof.id === R.BOF2) rd.ver = 2;
  else if (bof.id === R.BOF3) rd.ver = 3;
  else if (bof.id === R.BOF4) rd.ver = 4;
  else {
    const vers = u16(bof.data, 0);
    rd.ver = vers === 0x0600 ? 8 : vers === 0x0500 ? 5 : bof.data.length >= 16 ? 8 : hint || 5;
  }
  if (rd.ver === 8) rd.dec8 = makeDecoder(1252);
  const firstDt = bof.data.length >= 4 ? u16(bof.data, 2) : 0x0010;
  const sheets: { ctx: SheetCtx; bound: number }[] = [];
  const newCtx = (name: string): SheetCtx => {
    const sheet = new Sheet(rd.book, name);
    return { sheet, shared: new Map(), pendingExp: [], lastFormula: null, objText: new Map(), lastObj: -1, notes: [], note5: null, frozen: false };
  };
  const guard = (fn: () => void) => {
    try { fn(); }
    catch (e: any) {
      if (e?.message === XLS_ENCRYPTED_MSG) throw e;
      if (!lenient) throw e;
      rd.warn(`A damaged record was skipped (${e?.message || e}).`);
    }
  };

  if (rd.ver === 4 && firstDt === 0x0100) {
    // Excel 4.0 workbook (BIFF4W): a globals substream lists the sheets (BUNDLESHEET), then each sheet
    // follows as a complete BIFF4 stream with its own fonts, formats and cell styles.
    const bundles: { pos: number; name: string }[] = [];
    for (i = i + 1; i < recs.length && recs[i].id !== R.EOF; i++) {
      const rec = recs[i], d = rec.data;
      if (rec.id === 0x008f && d.length >= 7) {
        const name = 7 + d[6] <= d.length ? rd.dec8(d.subarray(7, 7 + d[6])) : rd.dec8(d.subarray(5, 5 + d[4]));
        bundles.push({ pos: u32(d, 0), name });
      } else guard(() => rd.globals(rec));
    }
    let k = 0;
    for (i = i + 1; i < recs.length; i++) {
      if (!isBofId(recs[i].id)) continue;
      const bofRec = recs[i];
      const dt = u16(bofRec.data, 2);
      const bundle = bundles.find((b) => b.pos === bofRec.off) || bundles[k];
      k++;
      // Each sheet has its own style tables.
      rd.fonts = []; rd.xfs = []; rd.formats = new Map(); rd.formatSeq = 0; rd.xfStyle.clear(); rd.attrCache.clear();
      const ctx = dt === 0x0010 || dt === 0x0040 ? newCtx(bundle?.name || `Sheet${sheets.length + 1}`) : null;
      let depth = 1;
      for (i = i + 1; i < recs.length; i++) {
        const rec = recs[i];
        if (isBofId(rec.id)) { depth++; continue; }
        if (rec.id === R.EOF) { depth--; if (!depth) break; continue; }
        if (depth > 1 || !ctx) continue;
        guard(() => rd.globals(rec));
        guard(() => processSheetRec(rd, ctx, rec));
      }
      if (ctx) { finishSheet(rd, ctx); sheets.push({ ctx, bound: -1 }); }
    }
  } else if (rd.ver < 5 || firstDt === 0x0010 || firstDt === 0x0040) {
    // Single-sheet file (BIFF2-4, or a stream that starts at a worksheet): globals and cells share one substream.
    const ctx = newCtx('Sheet1');
    for (i = i + 1; i < recs.length && recs[i].id !== R.EOF; i++) {
      const rec = recs[i];
      if (isBofId(rec.id)) { // embedded chart: skip to its EOF
        let depth = 1;
        while (++i < recs.length && depth) { if (isBofId(recs[i].id)) depth++; else if (recs[i].id === R.EOF) depth--; }
        i--; continue;
      }
      guard(() => rd.globals(rec));
      guard(() => processSheetRec(rd, ctx, rec));
    }
    finishSheet(rd, ctx);
    sheets.push({ ctx, bound: -1 });
  } else {
    // Workbook globals substream
    for (i = i + 1; i < recs.length && recs[i].id !== R.EOF; i++) {
      const rec = recs[i];
      if (isBofId(rec.id)) break;
      guard(() => rd.globals(rec));
    }
    const byPos = new Map<number, number>();
    rd.bound.forEach((b, k) => byPos.set(b.pos, k));
    const used = new Set<number>();
    let order = 0;
    // Sheet substreams
    for (; i < recs.length; i++) {
      if (!isBofId(recs[i].id)) continue;
      const bofRec = recs[i];
      const dt = u16(bofRec.data, 2);
      let bi = byPos.get(bofRec.off);
      if (bi == null || used.has(bi)) { while (used.has(order)) order++; bi = order < rd.bound.length ? order : undefined; }
      if (bi != null) used.add(bi);
      const isWork = dt === 0x0010 || dt === 0x0040 || (dt !== 0x0020 && dt !== 0x0006 && dt !== 0x0005);
      const bound = bi != null ? rd.bound[bi] : undefined;
      let depth = 1;
      const ctx = isWork && (!bound || bound.type === 0 || bound.type === 1) ? newCtx(bound?.name || `Sheet${sheets.length + 1}`) : null;
      for (i = i + 1; i < recs.length; i++) {
        const rec = recs[i];
        if (isBofId(rec.id)) {
          // An embedded chart substream nests inside a worksheet; a BOF at depth 1 after a lost EOF starts the next sheet.
          if (lenient && u16(rec.data, 2) !== 0x0020) { i--; break; }
          depth++; continue;
        }
        if (rec.id === R.EOF) { depth--; if (!depth) break; continue; }
        if (depth > 1 || !ctx) continue;
        guard(() => processSheetRec(rd, ctx, rec));
      }
      if (ctx) {
        finishSheet(rd, ctx);
        if (bound) {
          ctx.sheet.visibility = bound.state === 1 ? 'hidden' : bound.state === 2 ? 'veryHidden' : 'visible';
        }
        sheets.push({ ctx, bound: bi ?? -1 });
      }
    }
  }

  // Sheets in BOUNDSHEET order; unmatched ones after.
  sheets.sort((a, b) => (a.bound < 0 ? 1e9 : a.bound) - (b.bound < 0 ? 1e9 : b.bound));
  const seen = new Set<string>();
  for (const { ctx } of sheets) {
    let nm = ctx.sheet.name.replace(/[\\/?*[\]:]/g, '_').slice(0, 31) || 'Sheet';
    let k = 2;
    while (seen.has(nm.toLowerCase())) nm = `${ctx.sheet.name.slice(0, 27)} (${k++})`;
    seen.add(nm.toLowerCase());
    ctx.sheet.name = nm;
    rd.book.sheets.push(ctx.sheet);
  }

  // Defined names (decoded after all sheets so 3D references and other names resolve).
  for (const nm of rd.names) {
    let text: string | null = null;
    try { text = nm.rgce.length ? rd.decode(nm.rgce, nm.extra, 0, 0, false) : null; } catch { text = null; }
    if (text == null) continue;
    const scope = nm.itab > 0 ? rd.bound[nm.itab - 1]?.name ?? null : null;
    if (nm.name === 'Print_Area' && scope) {
      const sh = rd.book.sheetByName(scope);
      if (sh) sh.print.area = text.replace(/^(?:'[^']*(?:''[^']*)*'|[^!]+)!/, '');
    } else if (nm.name === 'Print_Titles' && scope) {
      const sh = rd.book.sheetByName(scope);
      if (sh) {
        for (const part of text.replace(/^\(|\)$/g, '').split(',')) {
          const ref = part.replace(/^.*!/, '');
          if (/^\$?\d+:\$?\d+$/.test(ref)) sh.print.titleRows = ref; else if (/^\$?[A-Z]+:\$?[A-Z]+$/.test(ref)) sh.print.titleCols = ref;
        }
      }
    }
    if (nm.name === '_FilterDatabase') continue;
    rd.book.names.push({ name: nm.name, ref: text, sheet: scope, ...(nm.hidden ? { hidden: true } : {}) });
  }
  return { rd, walk: ws, records: recs.length };
}

function emptyStats(source: BiffStats['source']): BiffStats {
  return { biff: 0, source, sheets: 0, cells: 0, formulas: 0, formulasDecoded: 0, sharedStrings: 0, records: 0, resyncs: 0, skippedBytes: 0, codepage: 0, encrypted: false };
}

function result(rd: Reader, ws: WalkStats, records: number, source: BiffStats['source'], warnings: string[]): BiffResult {
  if (!rd.book.sheets.length) rd.book.sheets.push(new Sheet(rd.book, 'Sheet1'));
  rd.book.activeSheet = 0;
  rd.book.clearHistory();
  return {
    book: rd.book,
    warnings: [...warnings, ...rd.warnings],
    stats: {
      biff: rd.ver, source, sheets: rd.book.sheets.length, cells: rd.cells, formulas: rd.formulas, formulasDecoded: rd.decoded,
      sharedStrings: rd.sst.length, records, resyncs: ws.resyncs, skippedBytes: ws.skippedBytes, codepage: rd.codepage, encrypted: false,
    },
  };
}

/**
 * Read a BIFF workbook. Accepts a whole .xls file (compound file) or a raw BIFF stream.
 * Strict mode throws on damage; lenient mode recovers what it can and reports warnings.
 * Encrypted workbooks always throw XLS_ENCRYPTED_MSG.
 */
export function readBiff(u8: Uint8Array, opts: { lenient?: boolean } = {}): BiffResult {
  const lenient = !!opts.lenient;
  const warnings: string[] = [];
  let found: ReturnType<typeof findStream> = null;
  try { found = findStream(u8, warnings); }
  catch (e: any) { if (e?.message === XLS_ENCRYPTED_MSG || !lenient) throw e; }
  let best: BiffResult | null = null;
  if (found) {
    if (found.damaged && !lenient) throw new Error(`Damaged compound file: ${warnings.join(' ')}`);
    try {
      const { rd, walk: ws, records } = parseStream(found.s, lenient, found.hint);
      best = result(rd, ws, records, found.source, warnings);
    } catch (e: any) {
      if (e?.message === XLS_ENCRYPTED_MSG || !lenient) throw e;
      warnings.push(`The workbook stream could not be read (${e?.message || e}).`);
    }
  } else if (!lenient) {
    throw new Error(isOle2(u8) ? 'No Workbook stream was found in the compound file.' : 'Not an Excel 97-2003 (BIFF) workbook.');
  }
  // Lenient: also scan the raw bytes for BIFF substreams and keep whichever recovers more.
  if (lenient && (!found || found.damaged || !best || best.stats.cells === 0)) {
    const bofs = scanBofs(u8);
    if (bofs.length) {
      try {
        const { rd, walk: ws, records } = parseStream(u8.subarray(bofs[0]), true, 8);
        const scanned = result(rd, ws, records, 'scan', [...warnings, 'The workbook was rebuilt from BIFF records found by scanning the raw bytes.']);
        if (!best || scanned.stats.cells > best.stats.cells) best = scanned;
      } catch (e: any) {
        if (e?.message === XLS_ENCRYPTED_MSG) throw e;
      }
    }
  }
  if (!best) {
    if (!lenient) throw new Error('No workbook data found.');
    const book = new Workbook();
    book.clearHistory();
    return { book, warnings: [...warnings, 'No BIFF records could be found.'], stats: emptyStats('scan') };
  }
  return best;
}

/** Normal .xls import (strict). */
export async function importXls(data: ArrayBuffer, name: string, warnings?: string[]): Promise<Workbook> {
  const { book, warnings: w } = readBiff(new Uint8Array(data), { lenient: false });
  if (warnings) warnings.push(...w);
  if (name && !book.props.title) book.props.title = name.replace(/\.[^.]+$/, '');
  return book;
}
