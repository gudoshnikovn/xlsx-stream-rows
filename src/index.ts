// ─── Unified public API ──────────────────────────────────────────────────────
// The primary surface — auto-detects format and streams rows uniformly.

export {
  openWorkbook,
  streamRows,
  readRows,
  type WorkbookInfo,
  type ReadOptions,
} from './api.js';

export { detectFormat, type SpreadsheetFormat } from './formatDetect.js';

export type { CellValue, Row } from './types.js';

// ─── Errors ──────────────────────────────────────────────────────────────────

export {
  XlsxStreamError,
  NotAZipError,
  Zip64NotSupportedError,
  InvalidLocalHeaderError,
  UnsupportedCompressionError,
  EntryTooLargeError,
  InvalidOpcPackageError,
  SharedStringsTooLargeError,
  SheetNotFoundError,
  XlsFileTooLargeError,
  XlsxPackageMissingError,
  FormatNotSupportedError,
} from './errors.js';

// ─── Per-format adapters ─────────────────────────────────────────────────────
// Useful when the caller already knows the format and wants the typed
// per-format options without going through the unified dispatcher.

export {
  openXlsxWorkbook,
  streamXlsxRows,
  type XlsxStreamOptions,
} from './xlsxStreamReader.js';

export {
  openCsvWorkbook,
  streamCsvRows,
  type CsvStreamOptions,
  type CsvWorkbookInfo,
} from './csvStreamReader.js';

export {
  openXlsWorkbook,
  streamXlsRows,
  type XlsStreamOptions,
  type XlsWorkbookInfo,
} from './xlsAdapter.js';

// ─── Lower-level building blocks ─────────────────────────────────────────────
// Exported so power users can compose their own pipelines (e.g. read raw ZIP
// entries, drive the row parser from custom XML).

export {
  readZipEntries,
  openEntryStream,
  openDecompressedStream,
  readEntryToString,
  type ZipEntry,
} from './zipReader.js';

export { decodeXml } from './decodeXml.js';
export { excelSerialToDate } from './excelDate.js';
export { parseDateFormatMask } from './xlsxStyles.js';
export {
  parseSheets,
  extractSheetNames,
  parseSharedStrings,
  type WorkbookSheet,
} from './xlsxXmlParser.js';
export {
  createRowParser,
  type RowParser,
  type RowParserContext,
} from './rowParser.js';
export { createCsvParser, type CsvParser } from './csvParser.js';
export {
  resolvePackagePaths,
  resolveWorkbookPath,
  relsPathFor,
  type PackagePaths,
} from './opcResolver.js';
