/**
 * XLSX I/O orchestration — wires the ZIP layer, OPC resolver, XML parsers,
 * and incremental row parser into a streaming `AsyncIterable<Row>`.
 *
 * Lifecycle:
 *   1. Read the ZIP Central Directory (~64 KiB tail read).
 *   2. Read `_rels/.rels` and `[Content_Types].xml`. Resolve workbook part
 *      via OPC (no hardcoded paths — see `opcResolver.ts`).
 *   3. Read workbook.xml + its .rels. Resolve sharedStrings, styles,
 *      and the target sheet path.
 *   4. Fetch sharedStrings and (when `parseDates`) styles in parallel.
 *   5. Open a decompression stream over the target sheet, pipe through
 *      `TextDecoderStream`, feed the row parser, yield rows.
 *
 * Cancellation: three equivalent stop mechanisms tear the pipeline down:
 *   - `maxRows` reached → generator returns; `finally` cancels the reader.
 *   - `for await … break` → iterator `return()` runs the `finally`.
 *   - `AbortSignal.aborted` → next read rejects with `signal.reason`.
 *
 * Aborting before the first row is yielded works too: we check the signal
 * after every async step, so callers can abort during sharedStrings fetch.
 */

import {
  InvalidOpcPackageError,
  SharedStringsTooLargeError,
  SheetNotFoundError,
} from '../errors.js';
import {
  type ZipEntry,
  openDecompressedStream,
  readEntryToString,
  readZipEntries,
} from '../zip/reader.js';
import {
  type PackagePaths,
  relsPathFor,
  resolvePackagePaths,
  resolveWorkbookPath,
} from '../zip/opcResolver.js';
import { parseSharedStrings, parseSingleSharedString, parseSheets } from './xmlParser.js';
import { parseDateFormatMask } from './styles.js';
import { createRowParser } from './rowParser.js';
import { checkAbort, abortable } from '../utils/abort.js';
import type { Row } from '../types.js';

const ROOT_RELS_PATH = '_rels/.rels';
const DEFAULT_SHARED_STRINGS_MAX = 64 * 1024 * 1024; // 64 MiB
const SMALL_PART_LIMIT = 4 * 1024 * 1024; // workbook.xml / .rels / styles cap

export interface WorkbookInfo {
  filename: string;
  sheetNames: string[];
  format: 'xlsx';
}

export interface XlsxStreamOptions {
  maxRows?: number;
  sheetName?: string;
  parseDates?: boolean;
  sharedStringsMaxBytes?: number;
  signal?: AbortSignal;
}

interface ResolvedOptions {
  maxRows: number;
  sheetName: string | undefined;
  parseDates: boolean;
  sharedStringsMaxBytes: number;
  signal: AbortSignal | undefined;
}

function resolveOptions(o: XlsxStreamOptions | undefined): ResolvedOptions {
  return {
    maxRows: o?.maxRows ?? Number.POSITIVE_INFINITY,
    sheetName: o?.sheetName,
    parseDates: o?.parseDates ?? true,
    sharedStringsMaxBytes: o?.sharedStringsMaxBytes ?? DEFAULT_SHARED_STRINGS_MAX,
    signal: o?.signal,
  };
}

interface Resolved {
  paths: PackagePaths;
  workbookXml: string;
  entryByPath: Map<string, ZipEntry>;
  sheetEntryByName: Map<string, ZipEntry>;
  sheetNames: string[];
}

/**
 * Walk the OPC indirection: read `_rels/.rels`, workbook part, workbook
 * rels; build the PackagePaths and the (sheet name → ZIP entry) map the
 * caller actually needs.
 */
async function resolvePackage(
  file: Blob,
  signal: AbortSignal | undefined,
): Promise<Resolved> {
  const entries = await abortable(readZipEntries(file), signal);
  checkAbort(signal);

  const entryByPath = new Map<string, ZipEntry>();
  for (const e of entries) entryByPath.set(e.filename, e);

  const rootRelsEntry = entryByPath.get(ROOT_RELS_PATH);
  if (rootRelsEntry === undefined) {
    throw new InvalidOpcPackageError(`Missing ${ROOT_RELS_PATH} — not a valid OPC package`);
  }
  const rootRels = await abortable(
    readEntryToString(file, rootRelsEntry, SMALL_PART_LIMIT),
    signal,
  );
  checkAbort(signal);

  const workbookPath = resolveWorkbookPath(rootRels);
  if (workbookPath === undefined) {
    throw new InvalidOpcPackageError(
      'Root rels has no relationship of type ".../officeDocument" — not an XLSX',
    );
  }
  const workbookEntry = entryByPath.get(workbookPath);
  if (workbookEntry === undefined) {
    throw new InvalidOpcPackageError(
      `Workbook part "${workbookPath}" referenced by rels is not in the archive`,
    );
  }

  const workbookRelsPath = relsPathFor(workbookPath);
  const workbookRelsEntry = entryByPath.get(workbookRelsPath);
  if (workbookRelsEntry === undefined) {
    throw new InvalidOpcPackageError(`Missing workbook rels at "${workbookRelsPath}"`);
  }

  const [workbookXml, workbookRelsXml] = await abortable(
    Promise.all([
      readEntryToString(file, workbookEntry, SMALL_PART_LIMIT),
      readEntryToString(file, workbookRelsEntry, SMALL_PART_LIMIT),
    ]),
    signal,
  );
  checkAbort(signal);

  const paths = resolvePackagePaths(rootRels, workbookXml, workbookRelsXml);
  if (paths === undefined) {
    throw new InvalidOpcPackageError('Failed to resolve OPC package layout');
  }

  // Build (sheet name → ZIP entry) by joining workbook.xml's sheet list
  // (which gives name + rId in document order) with paths.sheetByRId.
  const sheetEntryByName = new Map<string, ZipEntry>();
  const sheetNames: string[] = [];
  for (const sheet of parseSheets(workbookXml)) {
    sheetNames.push(sheet.name);
    const partPath = paths.sheetByRId.get(sheet.rId);
    if (partPath === undefined) continue; // dangling rId — skip
    const entry = entryByPath.get(partPath);
    if (entry === undefined) continue;
    sheetEntryByName.set(sheet.name, entry);
  }

  return { paths, workbookXml, entryByPath, sheetEntryByName, sheetNames };
}

/**
 * Return sheet names without reading any row data.
 *
 * Reads only the OPC indirection chain: ~64 KiB tail (Central Directory) +
 * `_rels/.rels` + workbook part + workbook rels. Total ≈ 100 KiB on a
 * typical workbook regardless of file size.
 */
export async function openXlsxWorkbook(file: File): Promise<WorkbookInfo> {
  const { sheetNames } = await resolvePackage(file, undefined);
  return { filename: file.name, sheetNames, format: 'xlsx' };
}

/**
 * Stream rows from the chosen sheet.
 *
 * Memory profile (1 GB workbook with small sharedStrings):
 *   ~ sharedStrings size + 64 KiB DecompressionStream window + one Row
 *   = a few MiB peak, independent of sheet length.
 *
 * Note: the function returns a fresh AsyncIterable each call. Iterating
 * twice will run the pipeline twice.
 */
export function streamXlsxRows(
  file: File,
  options?: XlsxStreamOptions,
): AsyncIterable<Row> {
  const opts = resolveOptions(options);
  return {
    [Symbol.asyncIterator](): AsyncIterator<Row> {
      return streamXlsxRowsImpl(file, opts);
    },
  };
}

async function* streamXlsxRowsImpl(
  file: File,
  opts: ResolvedOptions,
): AsyncGenerator<Row, void, unknown> {
  const { signal } = opts;
  checkAbort(signal);

  const { paths, entryByPath, sheetEntryByName, sheetNames } = await resolvePackage(
    file,
    signal,
  );

  const targetName = opts.sheetName ?? sheetNames[0];
  if (targetName === undefined) {
    throw new InvalidOpcPackageError('Workbook contains no sheets');
  }
  const sheetEntry = sheetEntryByName.get(targetName);
  if (sheetEntry === undefined) {
    throw new SheetNotFoundError(targetName, sheetNames);
  }

  const stylesPromise =
    opts.parseDates && paths.styles
      ? loadDateFormatMask(file, entryByPath, paths.styles)
      : Promise.resolve(new Set<number>());

  let sharedStrings: string[];
  let dateFormatStyleIds: Set<number>;

  if (paths.sharedStrings !== undefined && opts.maxRows < Number.POSITIVE_INFINITY) {
    // Two-pass lazy loading: collect only the shared-string indices used by
    // the first maxRows rows, then load just those entries. Memory is
    // proportional to the number of rows read, not to sharedStrings size.
    const neededIndices = await abortable(
      collectSharedStringIndices(file, sheetEntry, opts.maxRows, signal),
      signal,
    );
    checkAbort(signal);
    [sharedStrings, dateFormatStyleIds] = await abortable(
      Promise.all([
        loadSharedStringsSelective(file, entryByPath, paths.sharedStrings, neededIndices),
        stylesPromise,
      ]),
      signal,
    );
  } else {
    [sharedStrings, dateFormatStyleIds] = await abortable(
      Promise.all([
        paths.sharedStrings
          ? loadSharedStrings(file, entryByPath, paths.sharedStrings, opts.sharedStringsMaxBytes)
          : Promise.resolve<string[]>([]),
        stylesPromise,
      ]),
      signal,
    );
  }
  checkAbort(signal);

  const stream = await abortable(openDecompressedStream(file, sheetEntry), signal);
  // DOM lib types TextDecoderStream's writable as WritableStream<BufferSource>
  // but pipeThrough expects matching <Uint8Array, …> — same shim used in
  // zip/reader.ts. The runtime contract is identical.
  const td = new TextDecoderStream('utf-8') as unknown as ReadableWritablePair<
    string,
    Uint8Array
  >;
  const reader = stream.pipeThrough(td).getReader();

  const parser = createRowParser({
    sharedStrings,
    dateFormatStyleIds,
    parseDates: opts.parseDates,
  });

  let yielded = 0;

  try {
    while (yielded < opts.maxRows) {
      checkAbort(signal);
      const { done, value } = await abortable(reader.read(), signal);
      if (done) {
        for (const row of parser.end()) {
          yield row;
          yielded++;
          if (yielded >= opts.maxRows) return;
          checkAbort(signal);
        }
        return;
      }
      for (const row of parser.push(value)) {
        yield row;
        yielded++;
        if (yielded >= opts.maxRows) return;
        // Re-check between buffered rows: when the parser drains many rows
        // from a single chunk, the consumer may have aborted in between.
        checkAbort(signal);
      }
    }
  } finally {
    // Triggered on natural completion, maxRows return, iterator break,
    // and abort. Cancelling the reader propagates cancellation up the
    // pipeline (DecompressionStream → Blob slice), so no further bytes
    // are fetched or decompressed.
    try {
      await reader.cancel();
    } catch {
      /* the reader may already be closed; ignore */
    }
  }
}

// ─── helpers ────────────────────────────────────────────────────────────────

async function loadSharedStrings(
  file: Blob,
  entryByPath: Map<string, ZipEntry>,
  path: string,
  maxBytes: number,
): Promise<string[]> {
  const entry = entryByPath.get(path);
  if (entry === undefined) {
    throw new InvalidOpcPackageError(`sharedStrings part "${path}" not in archive`);
  }
  if (entry.uncompressedSize > maxBytes) {
    throw new SharedStringsTooLargeError(entry.uncompressedSize, maxBytes);
  }
  const xml = await readEntryToString(file, entry, maxBytes);
  return parseSharedStrings(xml);
}

async function loadDateFormatMask(
  file: Blob,
  entryByPath: Map<string, ZipEntry>,
  path: string,
): Promise<Set<number>> {
  const entry = entryByPath.get(path);
  if (entry === undefined) return new Set();
  const xml = await readEntryToString(file, entry, SMALL_PART_LIMIT);
  return parseDateFormatMask(xml);
}

// ─── lazy sharedStrings helpers ──────────────────────────────────────────────

function ssiLocalName(rawName: string): string {
  const colon = rawName.indexOf(':');
  return colon === -1 ? rawName : rawName.slice(colon + 1);
}

function ssiTagClose(s: string, start: number): number {
  let q = 0;
  for (let i = start; i < s.length; i++) {
    const c = s.charCodeAt(i);
    if (q !== 0) { if (c === q) q = 0; }
    else if (c === 0x22 || c === 0x27) q = c;
    else if (c === 0x3e) return i;
  }
  return -1;
}

/**
 * Pass 1 of the lazy-sharedStrings path: stream the sheet XML and collect the
 * set of shared-string indices referenced by the first `maxRows` rows.
 */
async function collectSharedStringIndices(
  file: Blob,
  sheetEntry: ZipEntry,
  maxRows: number,
  signal: AbortSignal | undefined,
): Promise<Set<number>> {
  const indices = new Set<number>();

  const stream = await abortable(openDecompressedStream(file, sheetEntry), signal);
  const td = new TextDecoderStream('utf-8') as unknown as ReadableWritablePair<string, Uint8Array>;
  const reader = stream.pipeThrough(td).getReader();

  let pending = '';
  let rowsDone = 0;
  let inRow = false;
  let isSCell = false;
  let inV = false;
  let vText = '';

  function onStart(rawName: string, attrsStr: string, selfClose: boolean): boolean {
    const name = ssiLocalName(rawName);
    switch (name) {
      case 'row':
        inRow = true;
        break;
      case 'c':
        if (inRow) isSCell = /\bt\s*=\s*["']s["']/.test(attrsStr);
        break;
      case 'v':
        if (isSCell && inRow) { inV = true; vText = ''; }
        break;
    }
    return selfClose ? onEnd(rawName) : false;
  }

  function onEnd(rawName: string): boolean {
    const name = ssiLocalName(rawName.trimEnd());
    switch (name) {
      case 'row':
        inRow = false;
        isSCell = false;
        rowsDone++;
        if (rowsDone >= maxRows) return true;
        break;
      case 'c':
        isSCell = false;
        inV = false;
        break;
      case 'v':
        if (inV) {
          const idx = Number.parseInt(vText.trim(), 10);
          if (Number.isFinite(idx)) indices.add(idx);
          inV = false;
        }
        break;
    }
    return false;
  }

  // Returns true when maxRows is reached and further processing can stop.
  function processChunk(input: string): boolean {
    const buf = pending + input;
    const n = buf.length;
    let i = 0;

    while (i < n) {
      const lt = buf.indexOf('<', i);

      if (lt === -1) {
        if (inV) vText += buf.slice(i);
        pending = '';
        return false;
      }

      if (inV && lt > i) vText += buf.slice(i, lt);

      if (lt + 1 >= n) {
        pending = buf.slice(lt);
        return false;
      }

      const c1 = buf.charCodeAt(lt + 1);

      if (c1 === 0x3f /* ? */) {
        const end = buf.indexOf('?>', lt + 2);
        if (end === -1) { pending = buf.slice(lt); return false; }
        i = end + 2;
        continue;
      }
      if (c1 === 0x21 /* ! */) {
        if (buf.charCodeAt(lt + 2) === 0x2d && buf.charCodeAt(lt + 3) === 0x2d) {
          const end = buf.indexOf('-->', lt + 4);
          if (end === -1) { pending = buf.slice(lt); return false; }
          i = end + 3;
        } else {
          const end = buf.indexOf('>', lt + 2);
          if (end === -1) { pending = buf.slice(lt); return false; }
          i = end + 1;
        }
        continue;
      }

      const gt = ssiTagClose(buf, lt + 1);
      if (gt === -1) { pending = buf.slice(lt); return false; }

      const content = buf.slice(lt + 1, gt);
      let stop = false;

      if (content.charCodeAt(0) === 0x2f) {
        stop = onEnd(content.slice(1));
      } else {
        let body = content;
        const selfClose = body.charCodeAt(body.length - 1) === 0x2f;
        if (selfClose) body = body.slice(0, -1).trimEnd();
        let nameEnd = 0;
        while (nameEnd < body.length) {
          const c = body.charCodeAt(nameEnd);
          if (c <= 0x20 || c === 0x2f) break;
          nameEnd++;
        }
        stop = onStart(body.slice(0, nameEnd), body.slice(nameEnd), selfClose);
      }

      i = gt + 1;
      if (stop) { pending = ''; return true; }
    }

    pending = i < n ? buf.slice(i) : '';
    return false;
  }

  try {
    while (true) {
      checkAbort(signal);
      const { done, value } = await abortable(reader.read(), signal);
      if (done) break;
      if (processChunk(value)) break;
    }
  } finally {
    try { await reader.cancel(); } catch { /* already closed */ }
  }

  return indices;
}

// Regex to find <si> (or <ns:si>) opening and closing tags in sharedStrings.xml.
const SI_OPEN_RE = /<(?:[a-zA-Z][\w-]*:)?si(?:\s[^>]*|\/?)>/;
const SI_CLOSE_RE = /<\/(?:[a-zA-Z][\w-]*:)?si>/;

/**
 * Pass 2 of the lazy-sharedStrings path: stream sharedStrings.xml and keep
 * only the entries whose indices appear in `needed`. Returns a sparse array
 * (unneeded indices are `undefined`) so that `sharedStrings[idx] ?? null`
 * in the row parser continues to work without changes.
 */
async function loadSharedStringsSelective(
  file: Blob,
  entryByPath: Map<string, ZipEntry>,
  path: string,
  needed: Set<number>,
): Promise<string[]> {
  const entry = entryByPath.get(path);
  if (entry === undefined) {
    throw new InvalidOpcPackageError(`sharedStrings part "${path}" not in archive`);
  }

  const result: string[] = [];
  if (needed.size === 0) return result;

  const maxNeeded = Math.max(...needed);

  const stream = await openDecompressedStream(file, entry);
  const td = new TextDecoderStream('utf-8') as unknown as ReadableWritablePair<string, Uint8Array>;
  const reader = stream.pipeThrough(td).getReader();

  let buf = '';
  let siIdx = 0;
  let inSi = false;
  let siAccum = '';
  let done = false;

  function processBuf(): void {
    while (!done && buf.length > 0) {
      if (!inSi) {
        const m = SI_OPEN_RE.exec(buf);
        if (!m) {
          // Keep from last '<' to handle tags that span chunk boundaries.
          const lt = buf.lastIndexOf('<');
          buf = lt >= 0 ? buf.slice(lt) : '';
          return;
        }
        const selfClose = m[0].endsWith('/>');
        buf = buf.slice(m.index + m[0].length);
        if (selfClose) {
          if (needed.has(siIdx)) result[siIdx] = '';
          siIdx++;
          if (siIdx > maxNeeded) done = true;
        } else {
          inSi = true;
          siAccum = '';
        }
      } else {
        const m = SI_CLOSE_RE.exec(buf);
        if (!m) {
          // Accumulate all but a trailing window that might hold a partial tag.
          const lt = buf.lastIndexOf('<');
          if (lt > 0) {
            siAccum += buf.slice(0, lt);
            buf = buf.slice(lt);
          } else if (lt === -1) {
            siAccum += buf;
            buf = '';
          }
          // lt === 0: '<' is at the start — keep buf as-is until more data.
          return;
        }
        siAccum += buf.slice(0, m.index);
        buf = buf.slice(m.index + m[0].length);
        inSi = false;
        if (needed.has(siIdx)) result[siIdx] = parseSingleSharedString(siAccum);
        siIdx++;
        siAccum = '';
        if (siIdx > maxNeeded) done = true;
      }
    }
  }

  try {
    while (true) {
      const { done: readDone, value } = await reader.read();
      if (readDone) break;
      buf += value;
      processBuf();
      if (done) break;
    }
  } finally {
    try { await reader.cancel(); } catch { /* already closed */ }
  }

  return result;
}
