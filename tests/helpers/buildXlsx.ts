/**
 * Synthetic XLSX builder for integration tests.
 *
 * Produces a byte-accurate, OPC-conformant package on top of `buildZip`. Lets
 * tests vary the package layout (move parts out of `xl/`, change the
 * relationship-id namespace prefix, omit sharedStrings/styles, switch to the
 * strict SpreadsheetML namespace) without hand-crafting ZIPs each time.
 */

import { buildZip, type EntrySpec } from './buildZip.js';

const TRANSITIONAL_SS_NS =
  'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const STRICT_SS_NS = 'http://purl.oclc.org/ooxml/spreadsheetml/main';

const TRANSITIONAL_REL_NS =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const STRICT_REL_NS = 'http://purl.oclc.org/ooxml/officeDocument/relationships';

const PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

export interface BuildXlsxSheet {
  /** Sheet name as it appears in the workbook. */
  name: string;
  /** Inner XML body — `<sheetData>…</sheetData>`, namespace-free. */
  sheetData: string;
}

export interface BuildXlsxOptions {
  sheets: BuildXlsxSheet[];
  /** Optional shared-string table. Each entry becomes one `<si><t>…</t></si>`. */
  sharedStrings?: string[];
  /**
   * Raw `<sst>` XML override — supersedes `sharedStrings` when provided.
   * Use to inject self-closing `<si/>` or other non-standard SST shapes.
   */
  rawSharedStringsXml?: string;
  /** Optional `<styleSheet>` XML body — provided whole so tests control format codes. */
  stylesXml?: string;
  /** Override workbook part location (default `xl/workbook.xml`). */
  workbookPath?: string;
  /** Override sharedStrings part location (default `xl/sharedStrings.xml`). */
  sharedStringsPath?: string;
  /** Override styles part location (default `xl/styles.xml`). */
  stylesPath?: string;
  /** Override sheet directory (default `xl/worksheets/`). */
  sheetsDir?: string;
  /** Use the strict SpreadsheetML namespace and rels namespace. */
  strict?: boolean;
  /** Use this XML namespace prefix for SpreadsheetML elements (e.g. 'x'). */
  ssPrefix?: string;
  /** Compress entries with deflate (8) instead of stored (0). */
  deflate?: boolean;
}

function xmlEscape(s: string): string {
  return s.replace(/[&<>"']/g, (c) => {
    switch (c) {
      case '&':
        return '&amp;';
      case '<':
        return '&lt;';
      case '>':
        return '&gt;';
      case '"':
        return '&quot;';
      default:
        return '&apos;';
    }
  });
}

function utf8(s: string): Uint8Array {
  return new TextEncoder().encode(s);
}

function dirname(path: string): string {
  const i = path.lastIndexOf('/');
  return i === -1 ? '' : path.slice(0, i);
}

function basename(path: string): string {
  const i = path.lastIndexOf('/');
  return i === -1 ? path : path.slice(i + 1);
}

function relsFor(partPath: string): string {
  const dir = dirname(partPath);
  const name = basename(partPath);
  return dir.length > 0 ? `${dir}/_rels/${name}.rels` : `_rels/${name}.rels`;
}

function relPath(fromDir: string, toPath: string): string {
  // Compute a path relative to fromDir suitable as a Target= in OPC rels.
  // For simplicity we use absolute (root-relative) targets with a leading '/'
  // — the resolver normalises both forms.
  return `/${toPath}`;
}

export async function buildXlsx(opts: BuildXlsxOptions): Promise<Uint8Array> {
  const ssNs = opts.strict ? STRICT_SS_NS : TRANSITIONAL_SS_NS;
  const relNs = opts.strict ? STRICT_REL_NS : TRANSITIONAL_REL_NS;
  const prefix = opts.ssPrefix ? `${opts.ssPrefix}:` : '';
  const xmlnsAttr = opts.ssPrefix
    ? `xmlns:${opts.ssPrefix}="${ssNs}"`
    : `xmlns="${ssNs}"`;
  const method: 0 | 8 = opts.deflate ? 8 : 0;

  const workbookPath = opts.workbookPath ?? 'xl/workbook.xml';
  const hasSharedStrings = opts.sharedStrings !== undefined || opts.rawSharedStringsXml !== undefined;
  const sharedStringsPath = hasSharedStrings
    ? opts.sharedStringsPath ?? 'xl/sharedStrings.xml'
    : undefined;
  const stylesPath = opts.stylesXml
    ? opts.stylesPath ?? 'xl/styles.xml'
    : undefined;
  const sheetsDir = opts.sheetsDir ?? 'xl/worksheets';

  // Allocate sheet paths and rIds.
  const sheetMeta = opts.sheets.map((s, i) => ({
    name: s.name,
    sheetData: s.sheetData,
    path: `${sheetsDir}/sheet${i + 1}.xml`,
    rId: `rId${i + 1}`,
  }));

  // sharedStrings/styles get rIds after the sheets to keep numbering simple.
  let nextRId = sheetMeta.length + 1;
  const sharedStringsRId = sharedStringsPath ? `rId${nextRId++}` : undefined;
  const stylesRId = stylesPath ? `rId${nextRId++}` : undefined;

  // ─── [Content_Types].xml ──────────────────────────────────────────────────
  // The library does not actually require this file for resolution, but real
  // OPC packages always include it; we mirror that for realism.
  const contentTypes =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/${workbookPath}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  ${sheetMeta
    .map(
      (s) =>
        `<Override PartName="/${s.path}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`,
    )
    .join('\n  ')}
  ${sharedStringsPath ? `<Override PartName="/${sharedStringsPath}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>` : ''}
  ${stylesPath ? `<Override PartName="/${stylesPath}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>` : ''}
</Types>`;

  // ─── _rels/.rels ──────────────────────────────────────────────────────────
  const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${PKG_REL_NS}">
  <Relationship Id="rId1" Type="${relNs}/officeDocument" Target="${relPath('', workbookPath)}"/>
</Relationships>`;

  // ─── workbook.xml ─────────────────────────────────────────────────────────
  const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<${prefix}workbook ${xmlnsAttr} xmlns:r="${relNs}">
  <${prefix}sheets>
    ${sheetMeta
      .map(
        (s) =>
          `<${prefix}sheet name="${xmlEscape(s.name)}" sheetId="${
            sheetMeta.indexOf(s) + 1
          }" r:id="${s.rId}"/>`,
      )
      .join('\n    ')}
  </${prefix}sheets>
</${prefix}workbook>`;

  // ─── workbook .rels ───────────────────────────────────────────────────────
  const workbookDir = dirname(workbookPath);
  const workbookRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${PKG_REL_NS}">
  ${sheetMeta
    .map(
      (s) =>
        `<Relationship Id="${s.rId}" Type="${relNs}/worksheet" Target="${relPath(workbookDir, s.path)}"/>`,
    )
    .join('\n  ')}
  ${
    sharedStringsRId && sharedStringsPath
      ? `<Relationship Id="${sharedStringsRId}" Type="${relNs}/sharedStrings" Target="${relPath(workbookDir, sharedStringsPath)}"/>`
      : ''
  }
  ${
    stylesRId && stylesPath
      ? `<Relationship Id="${stylesRId}" Type="${relNs}/styles" Target="${relPath(workbookDir, stylesPath)}"/>`
      : ''
  }
</Relationships>`;

  // ─── sharedStrings.xml ────────────────────────────────────────────────────
  const sharedStringsXml = opts.rawSharedStringsXml
    ? opts.rawSharedStringsXml
    : opts.sharedStrings
      ? `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="${ssNs}" count="${opts.sharedStrings.length}" uniqueCount="${opts.sharedStrings.length}">
${opts.sharedStrings.map((s) => `  <si><t>${xmlEscape(s)}</t></si>`).join('\n')}
</sst>`
      : undefined;

  // ─── sheet parts ──────────────────────────────────────────────────────────
  const sheetXmls = sheetMeta.map(
    (s) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<${prefix}worksheet ${xmlnsAttr} xmlns:r="${relNs}">
${s.sheetData}
</${prefix}worksheet>`,
  );

  // ─── assemble ZIP ─────────────────────────────────────────────────────────
  const entries: EntrySpec[] = [
    { name: '[Content_Types].xml', data: utf8(contentTypes), method },
    { name: '_rels/.rels', data: utf8(rootRels), method },
    { name: workbookPath, data: utf8(workbookXml), method },
    { name: relsFor(workbookPath), data: utf8(workbookRels), method },
  ];
  if (sharedStringsPath && sharedStringsXml) {
    entries.push({ name: sharedStringsPath, data: utf8(sharedStringsXml), method });
  }
  if (stylesPath && opts.stylesXml) {
    entries.push({ name: stylesPath, data: utf8(opts.stylesXml), method });
  }
  for (let i = 0; i < sheetMeta.length; i++) {
    entries.push({ name: sheetMeta[i]!.path, data: utf8(sheetXmls[i]!), method });
  }

  return buildZip(entries);
}
