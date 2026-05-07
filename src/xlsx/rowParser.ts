/**
 * Incremental SAX parser for `xl/worksheets/sheetN.xml`.
 *
 * Designed to be fed arbitrary chunks of decoded XML text. It emits complete
 * `Row` values as `</row>` is closed and never holds more than one in-flight
 * row in memory regardless of sheet size.
 *
 * The parser only understands the narrow XLSX subset relevant to cell values:
 *   <row>                            row container
 *     <c r="A1" t="?" s="?">         cell with optional type / style index
 *       <v>...</v>                   value (numeric, shared-string index, …)
 *       <is><t>...</t></is>          inline string (t="inlineStr")
 *       <f>...</f>                   formula text (ignored — we use cached <v>)
 *     </c>
 *   </row>
 *
 * Anything else (col widths, merge cells, conditional formats, etc.) flows
 * through as ignored markup. Namespace prefixes on any tag are accepted; the
 * parser matches by local name.
 *
 * Implementation notes:
 *   - Bytes that arrive mid-tag, mid-text, mid-comment, or mid-CDATA stay in
 *     `pending` until the rest of the construct arrives. The fuzz tests
 *     (`tests/rowParser.fuzz.test.ts`) verify that splitting the same input
 *     at every byte position yields identical output.
 *   - All decoded text accumulates into `textBuf` and is flushed only at
 *     `</v>` / `</t>`, so entity boundaries that straddle a chunk are still
 *     decoded correctly (the `<` that ends the text segment has already
 *     arrived by then).
 */

import { decodeXml } from '../utils/decodeXml.js';
import { excelSerialToDate } from '../utils/excelDate.js';
import type { CellValue, Row } from '../types.js';

export interface RowParserContext {
  sharedStrings: readonly string[];
  /** cellXfs indices that resolve to date number formats. */
  dateFormatStyleIds: ReadonlySet<number>;
  /** When true, numeric cells with a date-styled `s=` are returned as `Date`. */
  parseDates: boolean;
}

export interface RowParser {
  /** Feed a chunk of decoded sheet XML. Returns rows completed during this push. */
  push(chunk: string): Row[];
  /** Signal end-of-stream and flush any trailing row. */
  end(): Row[];
}

// ─── ASCII codes (kept inline for speed in the hot loop) ─────────────────────
const CC_LT = 0x3c;
const CC_GT = 0x3e;
const CC_SLASH = 0x2f;
const CC_QMARK = 0x3f;
const CC_EXCL = 0x21;
const CC_DQUOTE = 0x22;
const CC_SQUOTE = 0x27;

interface PendingCell {
  col: number;
  type: string;
  styleIdx: number;
  valueText: string;
  inlineText: string;
  hasV: boolean;
}

function localName(tag: string): string {
  const colonIdx = tag.indexOf(':');
  return colonIdx === -1 ? tag : tag.slice(colonIdx + 1);
}

function colFromRef(ref: string): number {
  let col = 0;
  for (let i = 0; i < ref.length; i++) {
    const c = ref.charCodeAt(i);
    if (c >= 0x41 && c <= 0x5a) col = col * 26 + (c - 0x40); // A-Z
    else if (c >= 0x61 && c <= 0x7a) col = col * 26 + (c - 0x60); // a-z
    else break; // hit a digit
  }
  return col - 1;
}

function parseAttrs(s: string): Map<string, string> {
  const out = new Map<string, string>();
  const n = s.length;
  let i = 0;
  while (i < n) {
    while (i < n && s.charCodeAt(i) <= 0x20) i++;
    if (i >= n) break;
    const nameStart = i;
    while (i < n) {
      const c = s.charCodeAt(i);
      if (c === 0x3d /* = */ || c <= 0x20) break;
      i++;
    }
    const name = s.slice(nameStart, i);
    while (i < n && s.charCodeAt(i) <= 0x20) i++;
    if (i >= n || s.charCodeAt(i) !== 0x3d) continue;
    i++;
    while (i < n && s.charCodeAt(i) <= 0x20) i++;
    const quote = s.charCodeAt(i);
    if (quote !== CC_DQUOTE && quote !== CC_SQUOTE) continue;
    i++;
    const valStart = i;
    while (i < n && s.charCodeAt(i) !== quote) i++;
    const value = s.slice(valStart, i);
    if (i < n) i++;
    out.set(name, value);
  }
  return out;
}

/**
 * Find the closing `>` of a tag, respecting attribute quoting.
 *
 * XLSX never embeds `>` literally inside attribute values (Excel encodes them
 * as `&gt;`), but a strictly conformant parser must still skip quoted regions.
 */
function findTagEnd(s: string, start: number): number {
  let inQuote = 0;
  for (let i = start; i < s.length; i++) {
    const c = s.charCodeAt(i);
    if (inQuote !== 0) {
      if (c === inQuote) inQuote = 0;
    } else if (c === CC_DQUOTE || c === CC_SQUOTE) {
      inQuote = c;
    } else if (c === CC_GT) {
      return i;
    }
  }
  return -1;
}

export function createRowParser(ctx: RowParserContext): RowParser {
  let pending = '';

  // Element-state flags — what scope are we currently inside?
  let inRow = false;
  let inV = false;
  let inIs = false;
  let inT = false; // <t> inside <is>
  // <f> formula text is intentionally not collected — XLSX always emits the
  // cached numeric/string result in the sibling <v>, which is what we want.

  // Current cell context (null between cells / outside row).
  let cell: PendingCell | null = null;

  // Auto-incrementing column index used when a `<c>` lacks an `r=` attribute
  // (some non-Excel producers omit it).
  let colCursor = 0;

  // Text accumulators — populated only while inV or inT.
  //
  // We keep raw (undecoded) text in `textRaw` and only run `decodeXml` on it
  // at flush time (closing `</v>` or `</t>`). Decoding per-chunk would break
  // when an entity like `&amp;` is split across chunks ("am" + "p;").
  //
  // CDATA segments enter the value verbatim (no entity decoding inside CDATA),
  // so when we hit a CDATA boundary we drain `textRaw` through `decodeXml`
  // into `textOut`, then append CDATA content as-is.
  let textRaw = '';
  let textOut = '';

  // Cells collected for the current row. We use an array (not a sparse Map)
  // so that the order of `</c>` events drives row construction; the column
  // index inside each PendingCell handles gaps.
  let rowCells: PendingCell[] = [];

  // Rows completed during the current push() call. Returned and reset.
  let completed: Row[] = [];

  function startElement(rawName: string, attrs: Map<string, string>): void {
    const name = localName(rawName);
    switch (name) {
      case 'row':
        inRow = true;
        rowCells = [];
        colCursor = 0;
        break;
      case 'c': {
        if (!inRow) return;
        const ref = attrs.get('r');
        const col = ref !== undefined ? colFromRef(ref) : colCursor;
        colCursor = col + 1;
        cell = {
          col,
          type: attrs.get('t') ?? '',
          styleIdx: parseStyleIdx(attrs.get('s')),
          valueText: '',
          inlineText: '',
          hasV: false,
        };
        break;
      }
      case 'v':
        if (cell) {
          inV = true;
          textRaw = '';
          textOut = '';
        }
        break;
      case 'is':
        if (cell) inIs = true;
        break;
      case 't':
        if (inIs && cell) {
          inT = true;
          textRaw = '';
          textOut = '';
        }
        break;
      case 'f':
        // Formula tag — no extra state; we ignore its character data because
        // textBuf is only fed while inV || inT is true.
        break;
    }
  }

  function endElement(rawName: string): void {
    const name = localName(rawName);
    switch (name) {
      case 'row':
        if (inRow) completed.push(buildRow(rowCells));
        inRow = false;
        rowCells = [];
        break;
      case 'c':
        if (cell) {
          rowCells.push(cell);
          cell = null;
        }
        break;
      case 'v':
        if (inV && cell) {
          cell.valueText = textOut + decodeXml(textRaw);
          cell.hasV = true;
        }
        inV = false;
        textRaw = '';
        textOut = '';
        break;
      case 't':
        if (inT && cell) cell.inlineText += textOut + decodeXml(textRaw);
        inT = false;
        textRaw = '';
        textOut = '';
        break;
      case 'is':
        inIs = false;
        break;
    }
  }

  function parseStyleIdx(raw: string | undefined): number {
    if (raw === undefined) return -1;
    const n = Number.parseInt(raw, 10);
    return Number.isFinite(n) ? n : -1;
  }

  function buildRow(cells: PendingCell[]): Row {
    if (cells.length === 0) return [];
    let maxCol = -1;
    for (const c of cells) if (c.col > maxCol) maxCol = c.col;
    const row: Row = new Array(maxCol + 1).fill(null);
    for (const c of cells) row[c.col] = convertCell(c);
    return row;
  }

  function convertCell(c: PendingCell): CellValue {
    switch (c.type) {
      case 's': {
        const idx = Number.parseInt(c.valueText, 10);
        return Number.isFinite(idx) ? ctx.sharedStrings[idx] ?? null : null;
      }
      case 'str':
        return c.hasV ? c.valueText : null;
      case 'inlineStr':
        return c.inlineText;
      case 'b':
        return c.valueText === '1' || c.valueText === 'true';
      case 'e':
        return c.valueText || null;
      case 'd': {
        // Strict-schema ISO-8601 date string.
        const t = Date.parse(c.valueText);
        return Number.isNaN(t) ? c.valueText : new Date(t);
      }
      case 'n':
      case '': {
        if (!c.hasV) return null;
        const num = Number.parseFloat(c.valueText);
        if (Number.isNaN(num)) return null;
        if (
          ctx.parseDates &&
          c.styleIdx >= 0 &&
          ctx.dateFormatStyleIds.has(c.styleIdx)
        ) {
          return excelSerialToDate(num);
        }
        return num;
      }
      default:
        // Unknown type — treat as string-ish.
        return c.valueText || c.inlineText || null;
    }
  }

  function dispatchTag(content: string): void {
    if (content.length === 0) return;

    if (content.charCodeAt(0) === CC_SLASH) {
      // </name> — strip leading '/' and any trailing whitespace
      let end = content.length;
      while (end > 1 && content.charCodeAt(end - 1) <= 0x20) end--;
      endElement(content.slice(1, end));
      return;
    }

    let body = content;
    let selfClose = false;
    if (body.charCodeAt(body.length - 1) === CC_SLASH) {
      selfClose = true;
      body = body.slice(0, -1);
    }

    // Trim trailing whitespace before parsing name.
    let nameEnd = 0;
    while (nameEnd < body.length) {
      const c = body.charCodeAt(nameEnd);
      if (c <= 0x20 || c === CC_SLASH) break;
      nameEnd++;
    }
    const name = body.slice(0, nameEnd);
    const attrs = parseAttrs(body.slice(nameEnd));

    startElement(name, attrs);
    if (selfClose) endElement(name);
  }

  function processChunk(buf: string): void {
    const n = buf.length;
    let i = 0;
    const collecting = (): boolean => inV || inT;

    while (i < n) {
      const ltIdx = buf.indexOf('<', i);

      if (ltIdx === -1) {
        // Pure character data till EOF of this buffer.
        if (collecting()) textRaw += buf.slice(i);
        i = n;
        break;
      }

      if (collecting() && ltIdx > i) {
        textRaw += buf.slice(i, ltIdx);
      }

      // Need at least one char after '<' to classify the construct.
      if (ltIdx + 1 >= n) {
        i = ltIdx;
        break;
      }

      const c1 = buf.charCodeAt(ltIdx + 1);

      if (c1 === CC_QMARK) {
        // Processing instruction <? ... ?>
        const close = buf.indexOf('?>', ltIdx + 2);
        if (close === -1) {
          i = ltIdx;
          break;
        }
        i = close + 2;
        continue;
      }

      if (c1 === CC_EXCL) {
        // <!--, <![CDATA[, or <!DOCTYPE>. Need 9 chars minimum to
        // disambiguate "<![CDATA[".
        if (n - ltIdx < 9) {
          i = ltIdx;
          break;
        }
        if (
          buf.charCodeAt(ltIdx + 2) === 0x2d /* - */ &&
          buf.charCodeAt(ltIdx + 3) === 0x2d
        ) {
          const close = buf.indexOf('-->', ltIdx + 4);
          if (close === -1) {
            i = ltIdx;
            break;
          }
          i = close + 3;
          continue;
        }
        if (
          buf.charCodeAt(ltIdx + 2) === 0x5b /* [ */ &&
          buf.startsWith('CDATA[', ltIdx + 3)
        ) {
          const close = buf.indexOf(']]>', ltIdx + 9);
          if (close === -1) {
            i = ltIdx;
            break;
          }
          if (collecting()) {
            // Drain accumulated raw text through entity decoding before
            // appending CDATA verbatim — CDATA content is literal, including
            // any `&amp;` etc., so it must not pass through decodeXml.
            textOut += decodeXml(textRaw);
            textRaw = '';
            textOut += buf.slice(ltIdx + 9, close);
          }
          i = close + 3;
          continue;
        }
        // <!DOCTYPE …> or <!ENTITY …> — Excel never emits these. Skip to '>'.
        const close = findTagEnd(buf, ltIdx + 2);
        if (close === -1) {
          i = ltIdx;
          break;
        }
        i = close + 1;
        continue;
      }

      // Element tag (start, end, or self-closing).
      const close = findTagEnd(buf, ltIdx + 1);
      if (close === -1) {
        i = ltIdx;
        break;
      }
      dispatchTag(buf.slice(ltIdx + 1, close));
      i = close + 1;
    }

    pending = i < n ? buf.slice(i) : '';
  }

  function push(chunk: string): Row[] {
    completed = [];
    processChunk(pending + chunk);
    const out = completed;
    completed = [];
    return out;
  }

  function end(): Row[] {
    completed = [];
    pending = '';
    const out = completed;
    completed = [];
    return out;
  }

  return { push, end };
}
