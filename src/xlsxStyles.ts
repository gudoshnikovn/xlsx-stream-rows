/**
 * Style → date-format detection from `xl/styles.xml`.
 *
 * XLSX has no dedicated date type in transitional schema (Excel's output): a
 * date cell is a numeric `<c>` whose `s=` attribute points at a `cellXfs/<xf>`
 * entry whose `numFmtId` resolves to a date format. To convert numeric values
 * to `Date`, we need to know which `cellXfs` indices are date-styled.
 *
 * `parseDateFormatMask` returns the set of `cellXfs` indices (the values that
 * appear in `<c s="...">`) backed by either a built-in date format ID
 * (ECMA-376 Part 1 §18.8.30) or a custom `<numFmt>` whose `formatCode`
 * contains date tokens.
 */

import { decodeXml } from './decodeXml.js';

// ECMA-376 Part 1 §18.8.30, table of built-in number formats:
//   14–22  date / date-time
//   27–36  locale-specific date variants (East-Asian)
//   45     mm:ss              (time)
//   46     [h]:mm:ss          (elapsed time)
//   47     mm:ss.0            (time with tenths)
//   50–58  locale-specific date variants (East-Asian)
const BUILTIN_DATE_IDS: ReadonlySet<number> = new Set([
  14, 15, 16, 17, 18, 19, 20, 21, 22,
  27, 28, 29, 30, 31, 32, 33, 34, 35, 36,
  45, 46, 47,
  50, 51, 52, 53, 54, 55, 56, 57, 58,
]);

/**
 * Detect whether a `formatCode` string represents a date/time format.
 *
 * Strips quoted literals (`"text"`, `'text'`) and bracketed sections
 * (`[Red]`, `[$-409]`), then tests for any unambiguous date/time token.
 *
 * `m` alone is ambiguous (month vs. minute depending on neighbouring tokens),
 * but every realistic month-only format contains an `m`. We accept `m` as a
 * date marker — false positives produce `Date` objects from numeric cells
 * that visually display like dates, which is what the user asked for anyway.
 */
function isDateFormatCode(code: string): boolean {
  const stripped = code
    .replace(/"[^"]*"/g, '')
    .replace(/'[^']*'/g, '')
    .replace(/\[[^\]]*\]/g, '');
  return /[ymdhsYMDHS]/.test(stripped);
}

const TAG_PREFIX = '(?:[a-zA-Z][\\w-]*:)?';

const NUMFMT_RE = new RegExp(
  `<${TAG_PREFIX}numFmt\\b([^/>]*)/?>`,
  'g',
);

const CELL_XFS_RE = new RegExp(
  `<${TAG_PREFIX}cellXfs\\b[^>]*>([\\s\\S]*?)</${TAG_PREFIX}cellXfs>`,
);

const XF_RE = new RegExp(`<${TAG_PREFIX}xf\\b([^>]*?)/?>`, 'g');

function attr(attrs: string, name: string): string | undefined {
  const re = new RegExp(`\\b${name}\\s*=\\s*["']([^"']*)["']`);
  const m = re.exec(attrs);
  return m ? m[1] : undefined;
}

export function parseDateFormatMask(stylesXml: string): Set<number> {
  const customDateIds = new Set<number>();

  for (const m of stylesXml.matchAll(NUMFMT_RE)) {
    const attrs = m[1] ?? '';
    const idStr = attr(attrs, 'numFmtId');
    const codeStr = attr(attrs, 'formatCode');
    if (idStr === undefined || codeStr === undefined) continue;
    const id = Number.parseInt(idStr, 10);
    if (!Number.isFinite(id)) continue;
    if (isDateFormatCode(decodeXml(codeStr))) customDateIds.add(id);
  }

  const result = new Set<number>();
  const cellXfsMatch = CELL_XFS_RE.exec(stylesXml);
  if (!cellXfsMatch) return result;
  const body = cellXfsMatch[1] ?? '';

  let xfIndex = 0;
  for (const m of body.matchAll(XF_RE)) {
    const attrs = m[1] ?? '';
    const idStr = attr(attrs, 'numFmtId');
    if (idStr !== undefined) {
      const id = Number.parseInt(idStr, 10);
      if (Number.isFinite(id) && (BUILTIN_DATE_IDS.has(id) || customDateIds.has(id))) {
        result.add(xfIndex);
      }
    }
    xfIndex++;
  }

  return result;
}
