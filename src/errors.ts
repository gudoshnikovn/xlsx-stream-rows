/**
 * Error hierarchy for xlsx-stream.
 *
 * All errors thrown by the library inherit from `XlsxStreamError` so that
 * callers can catch the family with a single `instanceof` check.
 */

export class XlsxStreamError extends Error {
  constructor(message: string) {
    super(message);
    this.name = new.target.name;
  }
}

/** EOCD signature not found within 65,557 bytes of EOF (APPNOTE.TXT §4.3.16). */
export class NotAZipError extends XlsxStreamError {}

/**
 * Central Directory or any entry uses ZIP64 extension sentinels
 * (`0xFFFFFFFF` / `0xFFFF`). ZIP64 support is deferred — see SPEC.
 */
export class Zip64NotSupportedError extends XlsxStreamError {}

/** Local File Header signature mismatch — file is corrupted or offset wrong. */
export class InvalidLocalHeaderError extends XlsxStreamError {}

/** ZIP entry uses a compression method other than 0 (stored) or 8 (deflate). */
export class UnsupportedCompressionError extends XlsxStreamError {
  readonly method: number;
  constructor(method: number) {
    super(`Unsupported ZIP compression method: ${method} (only 0 and 8 are supported)`);
    this.method = method;
  }
}

/** Bounded read of an entry exceeded the caller's size cap. */
export class EntryTooLargeError extends XlsxStreamError {
  readonly filename: string;
  readonly uncompressedSize: number;
  readonly limit: number;
  constructor(filename: string, uncompressedSize: number, limit: number) {
    super(
      `ZIP entry "${filename}" is ${uncompressedSize} bytes uncompressed, ` +
        `exceeds limit of ${limit}`,
    );
    this.filename = filename;
    this.uncompressedSize = uncompressedSize;
    this.limit = limit;
  }
}

/**
 * The OPC package is structurally invalid: missing `_rels/.rels`, no
 * `officeDocument` relationship, or a referenced part that the ZIP does
 * not contain. Almost always means the file is not actually an XLSX.
 */
export class InvalidOpcPackageError extends XlsxStreamError {}

/**
 * `xl/sharedStrings.xml` (uncompressed) exceeds `sharedStringsMaxBytes`.
 *
 * Pathological workbooks with millions of unique strings are rare but
 * possible; raising the limit trades the streaming memory guarantee.
 */
export class SharedStringsTooLargeError extends XlsxStreamError {
  readonly uncompressedSize: number;
  readonly limit: number;
  constructor(uncompressedSize: number, limit: number) {
    super(
      `sharedStrings.xml is ${uncompressedSize} bytes uncompressed, ` +
        `exceeds sharedStringsMaxBytes (${limit}). Raise the option ` +
        `or ask the producer to use inline strings.`,
    );
    this.uncompressedSize = uncompressedSize;
    this.limit = limit;
  }
}

/** The requested `sheetName` was not found in `xl/workbook.xml`. */
export class SheetNotFoundError extends XlsxStreamError {
  readonly sheetName: string;
  constructor(sheetName: string, available?: readonly string[]) {
    const list = available && available.length > 0 ? ` (available: ${available.join(', ')})` : '';
    super(`Sheet "${sheetName}" not found in workbook${list}`);
    this.sheetName = sheetName;
  }
}
