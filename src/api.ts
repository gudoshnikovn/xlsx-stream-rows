/**
 * Unified public API across XLSX, XLS, and CSV.
 *
 * Format is detected by magic bytes (ZIP, OLE2) with extension as fallback.
 * Each call dispatches to the right adapter:
 *   - XLSX → `xlsx/reader` (true streaming)
 *   - CSV  → `csv/reader`  (true streaming)
 *   - XLS  → `xls/adapter` (delegated; loads whole file into memory,
 *                           bounded by `xlsMaxBytes`)
 *
 * Options that don't apply to a given format are silently ignored — e.g.
 * `parseDates` has no meaning for CSV (everything is a string), and
 * `csvEncoding` is ignored for XLSX.
 */

import { detectFormat, type SpreadsheetFormat } from './formatDetect.js';
import { openXlsxWorkbook, streamXlsxRows } from './xlsx/reader.js';
import { openCsvWorkbook, streamCsvRows } from './csv/reader.js';
import { openXlsWorkbook, streamXlsRows } from './xls/adapter.js';
import type { Row } from './types.js';

export interface WorkbookInfo {
  filename: string;
  sheetNames: string[];
  format: SpreadsheetFormat;
}

export interface ReadOptions {
  /** Stop after yielding this many rows. Default: unlimited. */
  maxRows?: number;
  /** Sheet to read. Default: first sheet. Ignored for CSV. */
  sheetName?: string;
  /**
   * XLSX/XLS: convert numeric date-styled cells to `Date`. Default `true`.
   * Has no effect on CSV (CSV produces strings only).
   */
  parseDates?: boolean;
  /** XLSX-only: cap on `xl/sharedStrings.xml` uncompressed size. Default 64 MiB. */
  sharedStringsMaxBytes?: number;
  /** XLS-only: cap on the file size loaded into memory. Default 50 MiB. */
  xlsMaxBytes?: number;
  /**
   * CSV-only: explicit text encoding (e.g. `'windows-1251'`). UTF-8 / UTF-16
   * BOMs are auto-detected and override this. Default `'utf-8'`.
   */
  csvEncoding?: string;
  /** Cancel the read at any point. Aborting before the first row works. */
  signal?: AbortSignal;
}

/** Options accepted by `openWorkbook`. */
export interface OpenWorkbookOptions {
  /** XLS-only: cap on the file size loaded into memory. Default 50 MiB. */
  xlsMaxBytes?: number;
  /** Cancel the operation. */
  signal?: AbortSignal;
}

/**
 * List sheet names without reading any row data.
 *
 * - XLSX: ~100 KiB read (Central Directory + workbook.xml).
 * - CSV : zero I/O — returns one pseudo-sheet named after the file.
 * - XLS : full file load (bounded by `xlsMaxBytes`, default 50 MiB).
 */
export async function openWorkbook(
  file: File,
  options?: OpenWorkbookOptions,
): Promise<WorkbookInfo> {
  const format = await detectFormat(file);
  switch (format) {
    case 'xlsx':
      return openXlsxWorkbook(file, options?.signal);
    case 'csv':
      return openCsvWorkbook(file, options?.signal);
    case 'xls':
      return openXlsWorkbook(file, options);
  }
}

/**
 * Stream rows from the chosen sheet. Returns an `AsyncIterable<Row>`.
 *
 * Memory profile is dictated by the format:
 *   - XLSX: peak ≈ sharedStrings size + a few MiB. Independent of file size.
 *   - CSV : peak ≈ one row + decoder window. Independent of file size.
 *   - XLS : peak ≈ file size (no streaming primitive in the BIFF format).
 *
 * Stop mechanisms (all equivalent in teardown): `maxRows`, `break` out of
 * the loop, or `AbortSignal`. Aborting before the first row works.
 */
export function streamRows(file: File, options?: ReadOptions): AsyncIterable<Row> {
  return {
    [Symbol.asyncIterator](): AsyncIterator<Row> {
      return dispatch(file, options ?? {});
    },
  };
}

async function* dispatch(
  file: File,
  o: ReadOptions,
): AsyncGenerator<Row, void, unknown> {
  if (o.signal?.aborted) {
    throw o.signal.reason ?? new DOMException('Aborted', 'AbortError');
  }
  const format = await detectFormat(file);

  let inner: AsyncIterable<Row>;
  switch (format) {
    case 'xlsx': {
      const xlsxOpts: Parameters<typeof streamXlsxRows>[1] = {};
      if (o.maxRows !== undefined) xlsxOpts.maxRows = o.maxRows;
      if (o.sheetName !== undefined) xlsxOpts.sheetName = o.sheetName;
      if (o.parseDates !== undefined) xlsxOpts.parseDates = o.parseDates;
      if (o.sharedStringsMaxBytes !== undefined)
        xlsxOpts.sharedStringsMaxBytes = o.sharedStringsMaxBytes;
      if (o.signal !== undefined) xlsxOpts.signal = o.signal;
      inner = streamXlsxRows(file, xlsxOpts);
      break;
    }
    case 'csv': {
      const csvOpts: Parameters<typeof streamCsvRows>[1] = {};
      if (o.maxRows !== undefined) csvOpts.maxRows = o.maxRows;
      if (o.csvEncoding !== undefined) csvOpts.encoding = o.csvEncoding;
      if (o.signal !== undefined) csvOpts.signal = o.signal;
      inner = streamCsvRows(file, csvOpts);
      break;
    }
    case 'xls': {
      const xlsOpts: Parameters<typeof streamXlsRows>[1] = {};
      if (o.maxRows !== undefined) xlsOpts.maxRows = o.maxRows;
      if (o.sheetName !== undefined) xlsOpts.sheetName = o.sheetName;
      if (o.parseDates !== undefined) xlsOpts.parseDates = o.parseDates;
      if (o.xlsMaxBytes !== undefined) xlsOpts.xlsMaxBytes = o.xlsMaxBytes;
      if (o.signal !== undefined) xlsOpts.signal = o.signal;
      inner = streamXlsRows(file, xlsOpts);
      break;
    }
  }

  for await (const row of inner) yield row;
}

/**
 * Convenience: collect up to `maxRows` rows into an array.
 *
 * Equivalent to:
 * ```ts
 * const rows = [];
 * for await (const r of streamRows(file, options)) rows.push(r);
 * ```
 */
export async function readRows(file: File, options?: ReadOptions): Promise<Row[]> {
  const out: Row[] = [];
  for await (const row of streamRows(file, options)) out.push(row);
  return out;
}
