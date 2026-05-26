/**
 * CSV streaming reader. Pipes `File.stream()` through `TextDecoderStream`
 * into the incremental CSV parser, yielding rows on demand.
 *
 * Memory profile: the parser keeps at most one in-flight row, the decoder
 * keeps a small internal window, and the `Blob.stream()` underneath pulls
 * bytes lazily. A 1 GB CSV reads in the same memory envelope as a 1 KB one.
 */

import { createCsvParser } from './parser.js';
import { checkAbort, abortable } from '../utils/abort.js';
import type { Row } from '../types.js';

export interface CsvStreamOptions {
  /** Stop after yielding this many rows. */
  maxRows?: number;
  /**
   * Text encoding to use. Defaults to 'utf-8'. UTF-8 / UTF-16 LE / UTF-16 BE
   * BOMs are auto-detected and override this option.
   */
  encoding?: string;
  /**
   * Field separator character. Defaults to `'\t'` for `.tsv` files, `','`
   * for everything else. Pass any single character to override.
   */
  separator?: string;
  signal?: AbortSignal;
}

export interface CsvWorkbookInfo {
  filename: string;
  sheetNames: string[];
  format: 'csv';
}

interface ResolvedCsvOptions {
  maxRows: number;
  encoding: string;
  separator: string | undefined;
  signal: AbortSignal | undefined;
}

function resolveOptions(o: CsvStreamOptions | undefined): ResolvedCsvOptions {
  return {
    maxRows: o?.maxRows ?? Number.POSITIVE_INFINITY,
    encoding: o?.encoding ?? 'utf-8',
    separator: o?.separator,
    signal: o?.signal,
  };
}

function separatorForFile(file: File, explicit: string | undefined): string {
  if (explicit !== undefined) return explicit;
  const ext = file.name.slice(file.name.lastIndexOf('.') + 1).toLowerCase();
  return ext === 'tsv' ? '\t' : ',';
}

interface BomDetect {
  encoding: string;
  skip: number;
}

/**
 * Detect a leading BOM (UTF-8 EF BB BF, UTF-16 LE FF FE, UTF-16 BE FE FF).
 *
 * If found, returns the corresponding encoding and the byte count to skip.
 * Otherwise returns the caller-provided encoding with skip = 0. UTF-8 BOM
 * is in principle stripped by `TextDecoderStream` automatically, but doing
 * it ourselves keeps behaviour uniform across browsers.
 */
async function detectBom(file: Blob, fallback: string): Promise<BomDetect> {
  const head = new Uint8Array(await file.slice(0, 3).arrayBuffer());
  if (head.length >= 3 && head[0] === 0xef && head[1] === 0xbb && head[2] === 0xbf) {
    return { encoding: 'utf-8', skip: 3 };
  }
  if (head.length >= 2) {
    if (head[0] === 0xff && head[1] === 0xfe) return { encoding: 'utf-16le', skip: 2 };
    if (head[0] === 0xfe && head[1] === 0xff) return { encoding: 'utf-16be', skip: 2 };
  }
  return { encoding: fallback, skip: 0 };
}

export async function openCsvWorkbook(file: File, signal?: AbortSignal): Promise<CsvWorkbookInfo> {
  if (signal?.aborted) throw signal.reason ?? new DOMException('Aborted', 'AbortError');
  const baseName = file.name.replace(/\.[^./\\]+$/, '');
  return {
    filename: file.name,
    sheetNames: [baseName.length > 0 ? baseName : file.name],
    format: 'csv',
  };
}

/**
 * Stream rows from a CSV file. Each row is `string[]` because CSV has no
 * type information; values are returned to the caller verbatim. The function
 * still returns `Row` (`CellValue[]`) so the unified `streamRows` API can
 * mix CSV with XLSX/XLS without a type discriminator.
 */
export function streamCsvRows(file: File, options?: CsvStreamOptions): AsyncIterable<Row> {
  const opts = resolveOptions(options);
  return {
    [Symbol.asyncIterator](): AsyncIterator<Row> {
      return streamCsvRowsImpl(file, opts);
    },
  };
}

async function* streamCsvRowsImpl(
  file: File,
  opts: ResolvedCsvOptions,
): AsyncGenerator<Row, void, unknown> {
  const { signal } = opts;
  checkAbort(signal);

  const { encoding, skip } = await abortable(detectBom(file, opts.encoding), signal);
  checkAbort(signal);

  const bytes = skip > 0 ? file.slice(skip) : file;
  // DOM lib types TextDecoderStream's writable as WritableStream<BufferSource>;
  // pipeThrough wants <Uint8Array, …>. Same shim as elsewhere.
  const td = new TextDecoderStream(encoding) as unknown as ReadableWritablePair<
    string,
    Uint8Array
  >;
  const reader = bytes.stream().pipeThrough(td).getReader();
  const parser = createCsvParser(separatorForFile(file, opts.separator));

  let yielded = 0;

  try {
    while (yielded < opts.maxRows) {
      checkAbort(signal);
      const { done, value } = await abortable(reader.read(), signal);
      if (done) {
        for (const row of parser.end()) {
          yield row as Row;
          yielded++;
          if (yielded >= opts.maxRows) return;
          checkAbort(signal);
        }
        return;
      }
      for (const row of parser.push(value)) {
        yield row as Row;
        yielded++;
        if (yielded >= opts.maxRows) return;
        checkAbort(signal);
      }
    }
  } finally {
    try {
      await reader.cancel();
    } catch {
      /* ignore */
    }
  }
}
