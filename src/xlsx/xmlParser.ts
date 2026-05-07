/**
 * Pure XML parsers for the small XLSX parts that fit in memory:
 *   - `extractSheetNames` / `parseSheets` — `xl/workbook.xml`
 *   - `parseSharedStrings`                — `xl/sharedStrings.xml`
 *
 * These are regex-driven on purpose: the parts are bounded in size, the XML
 * shape is narrow and well-defined by ECMA-376, and a full SAX scan would be
 * overkill. The streaming row parser lives in `rowParser.ts` instead.
 *
 * All parsers tolerate any XML namespace prefix (e.g. `<x:sheet>`).
 */

import { decodeXml } from '../utils/decodeXml.js';

const PREFIX = '(?:[a-zA-Z][\\w-]*:)?';

// ─── workbook.xml — sheet index ──────────────────────────────────────────────

export interface WorkbookSheet {
  /** Sheet name as displayed in Excel. */
  name: string;
  /** Relationship ID — looked up in `xl/_rels/workbook.xml.rels`. */
  rId: string;
  /** sheetId attribute (rarely used by consumers, kept for completeness). */
  sheetId: string;
}

const SHEET_TAG_RE = new RegExp(`<${PREFIX}sheet\\b([^>]*)/?>`, 'g');

export function parseSheets(workbookXml: string): WorkbookSheet[] {
  const sheets: WorkbookSheet[] = [];
  for (const m of workbookXml.matchAll(SHEET_TAG_RE)) {
    const attrs = m[1] ?? '';
    const name = attrValue(attrs, 'name');
    const rId = relIdValue(attrs);
    const sheetId = attrValue(attrs, 'sheetId') ?? '';
    if (name !== undefined && rId !== undefined) {
      sheets.push({ name: decodeXml(name), rId, sheetId });
    }
  }
  return sheets;
}

export function extractSheetNames(workbookXml: string): string[] {
  return parseSheets(workbookXml).map((s) => s.name);
}

// ─── sharedStrings.xml — global string-deduplication table ──────────────────

const SI_RE = new RegExp(
  `<${PREFIX}si\\b[^>]*>([\\s\\S]*?)</${PREFIX}si>`,
  'g',
);
const RPH_RE = new RegExp(
  `<${PREFIX}rPh\\b[^>]*>[\\s\\S]*?</${PREFIX}rPh>`,
  'g',
);
const T_RE = new RegExp(
  `<${PREFIX}t\\b[^>]*>([\\s\\S]*?)</${PREFIX}t>`,
  'g',
);

/**
 * Parse a `<sst>` document into the ordered string table.
 *
 * Handles:
 *   - plain `<si><t>text</t></si>`
 *   - rich text `<si><r>...<t>chunk</t>...</r></si>` (concatenated in order)
 *   - empty cells `<si><t/></si>` → empty string
 *   - phonetic guides `<rPh>...</rPh>` (Asian locales) — stripped
 */
export function parseSharedStrings(xml: string): string[] {
  const out: string[] = [];
  for (const m of xml.matchAll(SI_RE)) {
    const body = (m[1] ?? '').replace(RPH_RE, '');
    let text = '';
    for (const tm of body.matchAll(T_RE)) {
      text += decodeXml(tm[1] ?? '');
    }
    out.push(text);
  }
  return out;
}

// ─── attribute helpers ──────────────────────────────────────────────────────

function attrValue(attrs: string, name: string): string | undefined {
  const re = new RegExp(`\\b${name}\\s*=\\s*["']([^"']*)["']`);
  const m = re.exec(attrs);
  return m ? m[1] : undefined;
}

/**
 * Look up the relationship-id attribute. The OOXML convention is `r:id`, but
 * the namespace prefix is producer-controlled, so we accept any prefix
 * pointing at `id` plus a bare `id` fallback.
 */
function relIdValue(attrs: string): string | undefined {
  const prefixed = /(?:^|\s)([a-zA-Z][\w-]*):id\s*=\s*["']([^"']*)["']/.exec(attrs);
  if (prefixed) return prefixed[2];
  return attrValue(attrs, 'id');
}
