/**
 * Open Packaging Conventions (ECMA-376 Part 2) part resolution.
 *
 * XLSX is an OPC package: a ZIP archive of "parts" whose locations are
 * **not** fixed by the standard. Instead, parts are discovered by following
 * relationships:
 *
 *   1. The package-level `_rels/.rels` lists relationships from the package
 *      root, including the one of type `…/officeDocument` that points at the
 *      workbook part.
 *   2. The workbook's own `.rels` (sibling file in `<dir>/_rels/<name>.rels`)
 *      lists relationships from the workbook to its sharedStrings, styles,
 *      and individual sheet parts.
 *
 * Excel's convention is `xl/workbook.xml` etc., but spec-compliant producers
 * (LibreOffice, Google Sheets exports, custom tooling) are free to place
 * parts anywhere. Resolving via relationships rather than hardcoded paths
 * means we don't break on those.
 */

import { decodeXml } from '../utils/decodeXml.js';

export interface PackagePaths {
  /** ZIP entry path of the workbook part. */
  workbook: string;
  /** ZIP entry path of `xl/sharedStrings.xml`, if the workbook has one. */
  sharedStrings?: string;
  /** ZIP entry path of `xl/styles.xml`, if present. */
  styles?: string;
  /**
   * Map from relationship Id (`r:id` attribute on `<sheet>`) to the
   * resolved ZIP entry path of that sheet part.
   */
  sheetByRId: Map<string, string>;
}

interface Relationship {
  id: string;
  type: string;
  target: string;
  /**
   * Per OPC §9.3, `Target` is normally relative to the part owning the
   * .rels file; if `TargetMode="External"` the relationship points outside
   * the package. We ignore external relationships.
   */
  external: boolean;
}

const REL_TAG_RE = /<Relationship\b([^>]*?)\/?>/g;

const REL_TYPE_OFFICE_DOC = '/officeDocument';
const REL_TYPE_SHARED_STRINGS = '/sharedStrings';
const REL_TYPE_STYLES = '/styles';
const REL_TYPE_WORKSHEET = '/worksheet';

function attrValue(attrs: string, name: string): string | undefined {
  // Attribute names in OPC rels XML are case-sensitive (`Id`, `Type`,
  // `Target`, `TargetMode`).
  const re = new RegExp(`\\b${name}\\s*=\\s*["']([^"']*)["']`);
  const m = re.exec(attrs);
  return m ? m[1] : undefined;
}

function parseRelationships(xml: string): Relationship[] {
  const out: Relationship[] = [];
  for (const m of xml.matchAll(REL_TAG_RE)) {
    const attrs = m[1] ?? '';
    const id = attrValue(attrs, 'Id');
    const type = attrValue(attrs, 'Type');
    const target = attrValue(attrs, 'Target');
    if (id === undefined || type === undefined || target === undefined) continue;
    const external = (attrValue(attrs, 'TargetMode') ?? '').toLowerCase() === 'external';
    out.push({ id, type, target: decodeXml(target), external });
  }
  return out;
}

/**
 * Resolve a relationship Target to a package-root-relative ZIP entry path.
 *
 * - Absolute targets (leading `/`) are root-relative directly.
 * - Relative targets join with the directory of the part owning the .rels
 *   file. Path components `.` and `..` are normalised.
 */
function resolveTarget(target: string, owningPartDir: string): string {
  let raw: string;
  if (target.startsWith('/')) {
    raw = target.slice(1);
  } else {
    raw = owningPartDir.length > 0 ? `${owningPartDir}/${target}` : target;
  }
  const stack: string[] = [];
  for (const segment of raw.split('/')) {
    if (segment === '' || segment === '.') continue;
    if (segment === '..') {
      stack.pop();
      continue;
    }
    stack.push(segment);
  }
  return stack.join('/');
}

function dirname(path: string): string {
  const i = path.lastIndexOf('/');
  return i === -1 ? '' : path.slice(0, i);
}

function basename(path: string): string {
  const i = path.lastIndexOf('/');
  return i === -1 ? path : path.slice(i + 1);
}

/**
 * Compute the conventional `.rels` path for a part: `<dir>/_rels/<name>.rels`.
 * If the part is at the package root, the rels file is `_rels/<name>.rels`.
 */
export function relsPathFor(partPath: string): string {
  const dir = dirname(partPath);
  const name = basename(partPath);
  return dir.length > 0 ? `${dir}/_rels/${name}.rels` : `_rels/${name}.rels`;
}

/**
 * Resolve the workbook part path from the root rels XML.
 *
 * Throws if no `officeDocument` relationship exists — in that case the
 * package isn't a Word/Excel/PowerPoint file, just a generic OPC archive.
 */
export function resolveWorkbookPath(rootRelsXml: string): string | undefined {
  for (const rel of parseRelationships(rootRelsXml)) {
    if (rel.external) continue;
    if (rel.type.endsWith(REL_TYPE_OFFICE_DOC)) {
      return resolveTarget(rel.target, '');
    }
  }
  return undefined;
}

/**
 * Resolve the full set of parts the streaming reader needs.
 *
 * Inputs are the *contents* of the relevant XML files; the caller is
 * responsible for fetching them from the ZIP. Returning `undefined` for
 * `sharedStrings` or `styles` means the workbook does not declare that
 * relationship — both are optional in OPC.
 */
export function resolvePackagePaths(
  rootRelsXml: string,
  workbookXml: string,
  workbookRelsXml: string,
): PackagePaths | undefined {
  const workbookPath = resolveWorkbookPath(rootRelsXml);
  if (workbookPath === undefined) return undefined;

  const workbookDir = dirname(workbookPath);
  const rels = parseRelationships(workbookRelsXml);

  const relsById = new Map<string, Relationship>();
  for (const rel of rels) relsById.set(rel.id, rel);

  let sharedStrings: string | undefined;
  let styles: string | undefined;
  for (const rel of rels) {
    if (rel.external) continue;
    if (rel.type.endsWith(REL_TYPE_SHARED_STRINGS) && sharedStrings === undefined) {
      sharedStrings = resolveTarget(rel.target, workbookDir);
    } else if (rel.type.endsWith(REL_TYPE_STYLES) && styles === undefined) {
      styles = resolveTarget(rel.target, workbookDir);
    }
  }

  // Cross-reference workbook's <sheet r:id="..."> with worksheet rels.
  const sheetByRId = new Map<string, string>();
  const sheetRefs = collectSheetRIds(workbookXml);
  for (const rId of sheetRefs) {
    const rel = relsById.get(rId);
    if (rel === undefined || rel.external) continue;
    if (!rel.type.endsWith(REL_TYPE_WORKSHEET)) continue;
    sheetByRId.set(rId, resolveTarget(rel.target, workbookDir));
  }

  const result: PackagePaths = {
    workbook: workbookPath,
    sheetByRId,
  };
  if (sharedStrings !== undefined) result.sharedStrings = sharedStrings;
  if (styles !== undefined) result.styles = styles;
  return result;
}

/**
 * Walk `<sheet>` tags in workbook.xml and collect their `r:id` (or any
 * prefix:id) attribute values, in document order.
 */
function collectSheetRIds(workbookXml: string): string[] {
  const out: string[] = [];
  const sheetRe = /<(?:[a-zA-Z][\w-]*:)?sheet\b([^>]*)\/?>/g;
  for (const m of workbookXml.matchAll(sheetRe)) {
    const attrs = m[1] ?? '';
    const prefixed = /(?:^|\s)[a-zA-Z][\w-]*:id\s*=\s*["']([^"']*)["']/.exec(attrs);
    if (prefixed) {
      out.push(prefixed[1] ?? '');
      continue;
    }
    const bare = attrValue(attrs, 'id');
    if (bare !== undefined) out.push(bare);
  }
  return out;
}
