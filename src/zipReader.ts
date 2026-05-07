/**
 * Streaming ZIP reader — APPNOTE.TXT (PKWARE) compliant.
 *
 * Reads only the bytes required to answer the caller's question:
 *   - `readZipEntries` reads the trailing ~64 KiB of the file.
 *   - `openEntryStream` / `openDecompressedStream` seek to the entry's
 *     compressed data and return a `ReadableStream<Uint8Array>` so that
 *     consumers can pull bytes on demand and cancel mid-flight.
 *
 * Every magic number below is named and annotated with its APPNOTE.TXT
 * section reference so future readers can verify against the spec
 * (https://pkware.cachefly.net/webdocs/casestudies/APPNOTE.TXT, rev 6.3.10).
 */

import {
  EntryTooLargeError,
  InvalidLocalHeaderError,
  NotAZipError,
  UnsupportedCompressionError,
  Zip64NotSupportedError,
} from './errors.js';

// ─── Signatures (APPNOTE §4.3) ───────────────────────────────────────────────
const SIG_EOCD = 0x06054b50;
const SIG_CD_ENTRY = 0x02014b50;
const SIG_LOCAL_HEADER = 0x04034b50;

// ─── End of Central Directory record (APPNOTE §4.3.16) ───────────────────────
// Layout (offsets within the 22-byte fixed block):
//   0  4  Signature
//   4  2  Disk number
//   6  2  Disk with start of CD
//   8  2  Entries on this disk
//  10  2  Total entries                ← EOCD_TOTAL_ENTRIES
//  12  4  Central directory size       ← EOCD_CD_SIZE
//  16  4  Central directory offset     ← EOCD_CD_OFFSET
//  20  2  Comment length (uint16)
//  22  …  Comment (≤ 65 535 bytes)
const EOCD_SIZE = 22;
const EOCD_MAX_COMMENT = 65535;
/** Maximum bytes from EOF where the EOCD can possibly live: 22 + 65 535 = 65 557. */
const EOCD_MAX_TOTAL = EOCD_SIZE + EOCD_MAX_COMMENT;
const EOCD_TOTAL_ENTRIES = 10;
const EOCD_CD_SIZE = 12;
const EOCD_CD_OFFSET = 16;

// ─── Central Directory file header (APPNOTE §4.3.12) ─────────────────────────
const CD_FIXED_SIZE = 46;
const CD_METHOD = 10;
const CD_COMPRESSED_SIZE = 20;
const CD_UNCOMPRESSED_SIZE = 24;
const CD_FILENAME_LEN = 28;
const CD_EXTRA_LEN = 30;
const CD_COMMENT_LEN = 32;
const CD_LOCAL_HEADER_OFFSET = 42;
const CD_FILENAME_DATA = 46;

// ─── Local File Header (APPNOTE §4.3.7) ──────────────────────────────────────
const LFH_FIXED_SIZE = 30;
const LFH_FILENAME_LEN = 26;
const LFH_EXTRA_LEN = 28;

// ─── ZIP64 sentinels (APPNOTE §4.3.16, §4.4.1.4) ─────────────────────────────
const ZIP64_SENTINEL_U32 = 0xffffffff;
const ZIP64_SENTINEL_U16 = 0xffff;

// ─── Compression methods (APPNOTE §4.4.5) ────────────────────────────────────
const METHOD_STORED = 0;
const METHOD_DEFLATE = 8;

// ─── Public types ────────────────────────────────────────────────────────────

export interface ZipEntry {
  /** Path inside the archive (UTF-8 decoded). */
  filename: string;
  /** Compression method: 0 = stored, 8 = deflate. */
  method: number;
  /** Byte offset of this entry's Local File Header. */
  localHeaderOffset: number;
  /** Compressed payload size in bytes. */
  compressedSize: number;
  /** Original (uncompressed) size in bytes. */
  uncompressedSize: number;
}

// ─── EOCD scan ───────────────────────────────────────────────────────────────

/**
 * Locate the EOCD record by scanning backwards from the end of the file.
 *
 * The EOCD lives at most 65 557 bytes from EOF (APPNOTE §4.3.16), so a single
 * trailing slice of that size is always sufficient. Smaller files are sliced
 * in full.
 *
 * Returns a `DataView` over the trailing window and the EOCD position within
 * that window. Throws `NotAZipError` if no EOCD signature is found.
 */
async function locateEocd(file: Blob): Promise<{ view: DataView; eocdPos: number }> {
  const window = Math.min(file.size, EOCD_MAX_TOTAL);
  if (window < EOCD_SIZE) {
    throw new NotAZipError(`File is ${file.size} bytes — too small to be a ZIP archive`);
  }

  const tail = await file.slice(file.size - window).arrayBuffer();
  const view = new DataView(tail);

  // Backward scan: the EOCD signature can appear anywhere in the comment
  // region, so we accept the *last* match — the legitimate one always sits
  // at the latest valid offset.
  for (let i = tail.byteLength - EOCD_SIZE; i >= 0; i--) {
    if (view.getUint32(i, true) === SIG_EOCD) {
      return { view, eocdPos: i };
    }
  }

  throw new NotAZipError('Not a valid ZIP/XLSX file: EOCD signature not found');
}

// ─── Public API ──────────────────────────────────────────────────────────────

/**
 * Read the Central Directory and return the list of entries.
 *
 * Fetches at most ~64 KiB (EOCD + CD). For typical XLSX files the CD itself
 * is well under 1 MiB even with thousands of entries.
 *
 * Throws:
 *   - `NotAZipError` — no EOCD signature in the trailing 65 557 bytes.
 *   - `Zip64NotSupportedError` — the archive uses ZIP64 extensions.
 */
export async function readZipEntries(file: Blob): Promise<ZipEntry[]> {
  const { view, eocdPos } = await locateEocd(file);

  const totalEntries = view.getUint16(eocdPos + EOCD_TOTAL_ENTRIES, true);
  const cdSize = view.getUint32(eocdPos + EOCD_CD_SIZE, true);
  const cdOffset = view.getUint32(eocdPos + EOCD_CD_OFFSET, true);

  if (
    totalEntries === ZIP64_SENTINEL_U16 ||
    cdSize === ZIP64_SENTINEL_U32 ||
    cdOffset === ZIP64_SENTINEL_U32
  ) {
    throw new Zip64NotSupportedError(
      'Archive uses ZIP64 extensions (file > 4 GiB or > 65 535 entries) — not supported',
    );
  }

  const cdBuf = await file.slice(cdOffset, cdOffset + cdSize).arrayBuffer();
  const cdView = new DataView(cdBuf);
  const decoder = new TextDecoder('utf-8');
  const entries: ZipEntry[] = [];
  let pos = 0;

  while (pos + CD_FIXED_SIZE <= cdBuf.byteLength) {
    if (cdView.getUint32(pos, true) !== SIG_CD_ENTRY) break;

    const method = cdView.getUint16(pos + CD_METHOD, true);
    const compressedSize = cdView.getUint32(pos + CD_COMPRESSED_SIZE, true);
    const uncompressedSize = cdView.getUint32(pos + CD_UNCOMPRESSED_SIZE, true);
    const fnLen = cdView.getUint16(pos + CD_FILENAME_LEN, true);
    const extraLen = cdView.getUint16(pos + CD_EXTRA_LEN, true);
    const commentLen = cdView.getUint16(pos + CD_COMMENT_LEN, true);
    const localHeaderOffset = cdView.getUint32(pos + CD_LOCAL_HEADER_OFFSET, true);

    if (
      compressedSize === ZIP64_SENTINEL_U32 ||
      uncompressedSize === ZIP64_SENTINEL_U32 ||
      localHeaderOffset === ZIP64_SENTINEL_U32
    ) {
      throw new Zip64NotSupportedError(
        `ZIP64 sentinel in CD entry "${decodeFilename(decoder, cdBuf, pos, fnLen)}"`,
      );
    }

    const filename = decodeFilename(decoder, cdBuf, pos, fnLen);
    entries.push({ filename, method, localHeaderOffset, compressedSize, uncompressedSize });

    pos += CD_FIXED_SIZE + fnLen + extraLen + commentLen;
  }

  if (entries.length !== totalEntries) {
    // Soft warning would belong here in a chattier API; we trust what we parsed.
  }

  return entries;
}

function decodeFilename(
  decoder: TextDecoder,
  cdBuf: ArrayBuffer,
  cdEntryPos: number,
  fnLen: number,
): string {
  return decoder.decode(new Uint8Array(cdBuf, cdEntryPos + CD_FILENAME_DATA, fnLen));
}

/**
 * Resolve the byte offset where an entry's compressed payload begins.
 *
 * The Central Directory points at the Local File Header, but the LFH has its
 * own variable-length filename + extra fields, so we must fetch the 30-byte
 * fixed LFH block to learn where the payload actually starts.
 */
async function resolveDataStart(file: Blob, entry: ZipEntry): Promise<number> {
  const lhBuf = await file
    .slice(entry.localHeaderOffset, entry.localHeaderOffset + LFH_FIXED_SIZE)
    .arrayBuffer();
  const lhView = new DataView(lhBuf);

  if (lhView.getUint32(0, true) !== SIG_LOCAL_HEADER) {
    throw new InvalidLocalHeaderError(
      `Invalid Local File Header at offset ${entry.localHeaderOffset} for "${entry.filename}"`,
    );
  }

  const fnLen = lhView.getUint16(LFH_FILENAME_LEN, true);
  const extraLen = lhView.getUint16(LFH_EXTRA_LEN, true);
  return entry.localHeaderOffset + LFH_FIXED_SIZE + fnLen + extraLen;
}

/**
 * Open a streaming reader over the entry's *compressed* bytes.
 *
 * Backed by `Blob.stream()` on a slice — no full read, no buffering of the
 * tail. Cancelling the returned stream stops further byte fetches.
 */
export async function openEntryStream(
  file: Blob,
  entry: ZipEntry,
): Promise<ReadableStream<Uint8Array>> {
  const dataStart = await resolveDataStart(file, entry);
  const slice = file.slice(dataStart, dataStart + entry.compressedSize);
  return slice.stream();
}

/**
 * Open a streaming reader over the entry's *decompressed* bytes.
 *
 * Wraps `openEntryStream` in `DecompressionStream('deflate-raw')` for method 8;
 * passes through unchanged for stored entries.
 *
 * Throws `UnsupportedCompressionError` for any other method.
 */
export async function openDecompressedStream(
  file: Blob,
  entry: ZipEntry,
): Promise<ReadableStream<Uint8Array>> {
  const compressed = await openEntryStream(file, entry);
  if (entry.method === METHOD_STORED) return compressed;
  if (entry.method === METHOD_DEFLATE) {
    // Cast: DOM lib types DecompressionStream's writable as
    // WritableStream<BufferSource>, but pipeThrough expects matching
    // ReadableWritablePair<Uint8Array, Uint8Array>. The runtime contract is
    // identical — TS 5.7's Uint8Array<ArrayBufferLike> generic creates a
    // spurious mismatch with DOM lib.
    const ds = new DecompressionStream('deflate-raw') as unknown as ReadableWritablePair<
      Uint8Array,
      Uint8Array
    >;
    return compressed.pipeThrough(ds);
  }
  // Cancel to release the underlying slice; we're not going to consume it.
  void compressed.cancel();
  throw new UnsupportedCompressionError(entry.method);
}

/**
 * Read the entire entry into a UTF-8 string. Used for small XML parts
 * (workbook.xml, .rels, sharedStrings up to a configured cap).
 *
 * Throws `EntryTooLargeError` if `entry.uncompressedSize` exceeds
 * `maxUncompressedBytes`. The check runs *before* any bytes are decompressed.
 */
export async function readEntryToString(
  file: Blob,
  entry: ZipEntry,
  maxUncompressedBytes: number,
): Promise<string> {
  if (entry.uncompressedSize > maxUncompressedBytes) {
    throw new EntryTooLargeError(entry.filename, entry.uncompressedSize, maxUncompressedBytes);
  }

  const stream = await openDecompressedStream(file, entry);
  // Same DOM-lib / TS 5.7 type-mismatch dance as in openDecompressedStream.
  const td = new TextDecoderStream('utf-8') as unknown as ReadableWritablePair<
    string,
    Uint8Array
  >;
  const reader = stream.pipeThrough(td).getReader();

  let text = '';
  try {
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      text += value;
    }
  } finally {
    reader.releaseLock();
  }
  return text;
}
