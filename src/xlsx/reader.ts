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
import { parseSharedStrings, parseSheets } from './xmlParser.js';
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

  // Fetch sharedStrings + (optionally) styles in parallel.
  const sharedStringsPromise = paths.sharedStrings
    ? loadSharedStrings(file, entryByPath, paths.sharedStrings, opts.sharedStringsMaxBytes)
    : Promise.resolve<string[]>([]);
  const stylesPromise =
    opts.parseDates && paths.styles
      ? loadDateFormatMask(file, entryByPath, paths.styles)
      : Promise.resolve(new Set<number>());

  const [sharedStrings, dateFormatStyleIds] = await abortable(
    Promise.all([sharedStringsPromise, stylesPromise]),
    signal,
  );
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
