/**
 * XLS adapter — thin wrapper over the optional `xlsx` peer dependency.
 *
 * XLS uses Microsoft BIFF inside an OLE2 Compound Document container; both
 * formats predate ZIP and require a full FAT-table parse, so there is no
 * streaming path. We load the entire file into memory, hand it to `xlsx`,
 * and yield rows from the resulting workbook.
 *
 * Because XLS reads are unbounded by design, the adapter enforces a default
 * 50 MiB cap (`xlsMaxBytes`). Callers can opt in to larger files by passing
 * `Infinity` — the memory cost is on them.
 *
 * `xlsx` is loaded via dynamic import so consumers who only read XLSX/CSV
 * never pay its bundle cost. If the package is missing we surface
 * `XlsxPackageMissingError` with an actionable install hint.
 */

import {
  XlsFileTooLargeError,
  XlsxPackageMissingError,
  SheetNotFoundError,
} from '../errors.js';
import type { CellValue, Row } from '../types.js';

export interface XlsStreamOptions {
  maxRows?: number;
  sheetName?: string;
  /** Hard cap on file size. Default 50 MiB. Set Infinity to disable. */
  xlsMaxBytes?: number;
  parseDates?: boolean;
  signal?: AbortSignal;
}

export interface XlsWorkbookInfo {
  filename: string;
  sheetNames: string[];
  format: 'xls';
}

const DEFAULT_XLS_MAX = 50 * 1024 * 1024;

interface ResolvedXlsOptions {
  maxRows: number;
  sheetName: string | undefined;
  xlsMaxBytes: number;
  parseDates: boolean;
  signal: AbortSignal | undefined;
}

function resolveOptions(o: XlsStreamOptions | undefined): ResolvedXlsOptions {
  return {
    maxRows: o?.maxRows ?? Number.POSITIVE_INFINITY,
    sheetName: o?.sheetName,
    xlsMaxBytes: o?.xlsMaxBytes ?? DEFAULT_XLS_MAX,
    parseDates: o?.parseDates ?? true,
    signal: o?.signal,
  };
}

function checkAbort(signal: AbortSignal | undefined): void {
  if (signal?.aborted) {
    throw signal.reason ?? new DOMException('Aborted', 'AbortError');
  }
}

interface XlsxModule {
  read: (data: ArrayBuffer | Uint8Array, opts?: { type?: string; cellDates?: boolean }) => XlsxWorkbook;
  utils: {
    sheet_to_json: (
      sheet: XlsxSheet,
      opts?: { header?: 1; raw?: boolean; defval?: unknown; blankrows?: boolean },
    ) => unknown[][];
  };
}

interface XlsxWorkbook {
  SheetNames: string[];
  Sheets: Record<string, XlsxSheet>;
}

type XlsxSheet = Record<string, unknown>;

let cachedXlsx: XlsxModule | null = null;

async function loadXlsx(): Promise<XlsxModule> {
  if (cachedXlsx !== null) return cachedXlsx;
  try {
    // Dynamic import keeps `xlsx` out of the static dep graph for consumers
    // who only need XLSX/CSV. The string literal is intentionally
    // non-resolvable at bundle time; tsup will leave it for the runtime.
    const mod = (await import('xlsx')) as unknown as XlsxModule | { default: XlsxModule };
    cachedXlsx = 'default' in mod ? mod.default : mod;
    return cachedXlsx;
  } catch {
    throw new XlsxPackageMissingError();
  }
}

interface LoadedWorkbook {
  wb: XlsxWorkbook;
  xlsx: XlsxModule;
}

async function loadWorkbook(file: Blob, opts: ResolvedXlsOptions): Promise<LoadedWorkbook> {
  if (file.size > opts.xlsMaxBytes) {
    throw new XlsFileTooLargeError(file.size, opts.xlsMaxBytes);
  }
  checkAbort(opts.signal);
  const xlsx = await loadXlsx();
  checkAbort(opts.signal);
  const buf = await file.arrayBuffer();
  checkAbort(opts.signal);
  return { wb: xlsx.read(buf, { type: 'array', cellDates: opts.parseDates }), xlsx };
}

export async function openXlsWorkbook(file: File): Promise<XlsWorkbookInfo> {
  const opts = resolveOptions(undefined);
  const { wb } = await loadWorkbook(file, opts);
  return { filename: file.name, sheetNames: [...wb.SheetNames], format: 'xls' };
}

export function streamXlsRows(file: File, options?: XlsStreamOptions): AsyncIterable<Row> {
  const opts = resolveOptions(options);
  return {
    [Symbol.asyncIterator](): AsyncIterator<Row> {
      return streamXlsRowsImpl(file, opts);
    },
  };
}

async function* streamXlsRowsImpl(
  file: File,
  opts: ResolvedXlsOptions,
): AsyncGenerator<Row, void, unknown> {
  const { signal } = opts;
  const { wb, xlsx } = await loadWorkbook(file, opts);

  const sheetName = opts.sheetName ?? wb.SheetNames[0];
  if (sheetName === undefined) {
    throw new SheetNotFoundError('(none)', wb.SheetNames);
  }
  const sheet = wb.Sheets[sheetName];
  if (sheet === undefined) {
    throw new SheetNotFoundError(sheetName, wb.SheetNames);
  }

  // header: 1 → array-of-arrays. raw: true → keep numeric/Date as-is rather
  // than stringifying. blankrows: false → match XLSX behaviour where empty
  // rows aren't yielded.
  const rows = xlsx.utils.sheet_to_json(sheet, {
    header: 1,
    raw: true,
    blankrows: false,
  });

  let yielded = 0;
  for (const raw of rows) {
    checkAbort(signal);
    yield normaliseRow(raw);
    yielded++;
    if (yielded >= opts.maxRows) return;
  }
}

function normaliseRow(raw: unknown[]): Row {
  const out: CellValue[] = new Array(raw.length);
  for (let i = 0; i < raw.length; i++) {
    out[i] = normaliseCell(raw[i]);
  }
  return out;
}

function normaliseCell(v: unknown): CellValue {
  if (v === undefined || v === null) return null;
  if (typeof v === 'string') return v;
  if (typeof v === 'number') return Number.isFinite(v) ? v : null;
  if (typeof v === 'boolean') return v;
  if (v instanceof Date) return v;
  // Anything else (e.g. xlsx's error-cell objects) → string fallback.
  return String(v);
}
