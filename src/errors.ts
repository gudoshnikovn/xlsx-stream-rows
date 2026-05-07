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
