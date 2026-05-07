// Public surface — incremental. The full streaming API
// (`openWorkbook`, `streamRows`, `readRows`) is still being built; for now we
// export the ZIP layer so it can be used directly while higher layers land.

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
