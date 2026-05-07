// Public surface — incremental. The full streaming API
// (`openWorkbook`, `streamRows`, `readRows`) is still being built; for now we
// export the layers that are already implemented so they can be used directly
// while higher layers land.

export {
  readZipEntries,
  openEntryStream,
  openDecompressedStream,
  readEntryToString,
  type ZipEntry,
} from './zipReader.js';

export {
  XlsxStreamError,
  NotAZipError,
  Zip64NotSupportedError,
  InvalidLocalHeaderError,
  UnsupportedCompressionError,
  EntryTooLargeError,
} from './errors.js';

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
export type { CellValue, Row } from './types.js';
