// Minimal streaming ZIP reader using File.slice() — no full-file reads.
//
// All byte offsets below come from the ZIP format specification (PKWARE APPNOTE.TXT).
// The format has been stable since 1989 — these values will not change.

export interface ZipEntry {
  filename: string;
  method: number; // 0 = stored, 8 = deflate
  localHeaderOffset: number;
  compressedSize: number;
  uncompressedSize: number;
}

// ─── ZIP format signatures ────────────────────────────────────────────────────

const SIG_EOCD = 0x06054b50; // End of Central Directory record
const SIG_CD_ENTRY = 0x02014b50; // Central Directory entry
const SIG_LOCAL_HEADER = 0x04034b50; // Local File Header

// ─── End of Central Directory Record (EOCD) — 22-byte fixed block ────────────
// Offset  Size  Field
//   0      4    Signature (SIG_EOCD)
//   4      2    Disk number
//   6      2    Disk with start of central directory
//   8      2    Entries on this disk
//  10      2    Total entries
//  12      4    Central directory size   ← EOCD_CD_SIZE
//  16      4    Central directory offset ← EOCD_CD_OFFSET
//  20      2    Comment length
//  22+    var   Comment (max 65535 bytes)

const EOCD_SIZE = 22;
const EOCD_MAX_COMMENT = 65535;
const EOCD_CD_SIZE = 12;
const EOCD_CD_OFFSET = 16;

// ─── Central Directory Entry — 46-byte fixed block + variable fields ──────────
// Offset  Size  Field
//   0      4    Signature (SIG_CD_ENTRY)
//   4      2    Version made by
//   6      2    Version needed
//   8      2    General purpose flags
//  10      2    Compression method       ← CD_METHOD
//  12      2    Last mod time
//  14      2    Last mod date
//  16      4    CRC-32
//  20      4    Compressed size          ← CD_COMPRESSED_SIZE
//  24      4    Uncompressed size        ← CD_UNCOMPRESSED_SIZE
//  28      2    Filename length          ← CD_FILENAME_LEN
//  30      2    Extra field length       ← CD_EXTRA_LEN
//  32      2    File comment length      ← CD_COMMENT_LEN
//  34      2    Disk number start
//  36      2    Internal attributes
//  38      4    External attributes
//  42      4    Local header offset      ← CD_LOCAL_HEADER_OFFSET
//  46+    var   Filename

const CD_FIXED_SIZE = 46;
const CD_METHOD = 10;
const CD_COMPRESSED_SIZE = 20;
const CD_UNCOMPRESSED_SIZE = 24;
const CD_FILENAME_LEN = 28;
const CD_EXTRA_LEN = 30;
const CD_COMMENT_LEN = 32;
const CD_LOCAL_HEADER_OFFSET = 42;
const CD_FILENAME_DATA = 46;

// ─── Local File Header — 30-byte fixed block + variable fields ────────────────
// Offset  Size  Field
//   0      4    Signature (SIG_LOCAL_HEADER)
//   4      2    Version needed
//   6      2    General purpose flags
//   8      2    Compression method
//  10      2    Last mod time
//  12      2    Last mod date
//  14      4    CRC-32
//  18      4    Compressed size
//  22      4    Uncompressed size
//  26      2    Filename length          ← LFH_FILENAME_LEN
//  28      2    Extra field length       ← LFH_EXTRA_LEN
//  30+    var   Filename
//  30+fn+ex    Compressed data

const LFH_FIXED_SIZE = 30;
const LFH_FILENAME_LEN = 26;
const LFH_EXTRA_LEN = 28;

// ─── Public API ───────────────────────────────────────────────────────────────

// Reads the ZIP Central Directory to get the list of all entries.
// Only fetches the last ~64 KB of the file (EOCD + directory), not the full file.
export async function readZipEntries(file: File): Promise<ZipEntry[]> {
  const searchSize = Math.min(file.size, EOCD_SIZE + EOCD_MAX_COMMENT);
  const tail = await file.slice(file.size - searchSize).arrayBuffer();
  const view = new DataView(tail);

  // Scan backwards for EOCD signature — it can be anywhere in the last 65 KB
  // if the ZIP comment field is used.
  let eocdPos = -1;
  for (let i = tail.byteLength - EOCD_SIZE; i >= 0; i--) {
    if (view.getUint32(i, true) === SIG_EOCD) { eocdPos = i; break; }
  }
  if (eocdPos === -1) throw new Error('Not a valid ZIP/XLSX file');

  const cdSize = view.getUint32(eocdPos + EOCD_CD_SIZE, true);
  const cdOffset = view.getUint32(eocdPos + EOCD_CD_OFFSET, true);

  const cdBuf = await file.slice(cdOffset, cdOffset + cdSize).arrayBuffer();
  const cdView = new DataView(cdBuf);
  const entries: ZipEntry[] = [];
  let pos = 0;

  while (pos < cdBuf.byteLength) {
    if (cdView.getUint32(pos, true) !== SIG_CD_ENTRY) break;

    const method = cdView.getUint16(pos + CD_METHOD, true);
    const compressedSize = cdView.getUint32(pos + CD_COMPRESSED_SIZE, true);
    const uncompressedSize = cdView.getUint32(pos + CD_UNCOMPRESSED_SIZE, true);
    const fnLen = cdView.getUint16(pos + CD_FILENAME_LEN, true);
    const extraLen = cdView.getUint16(pos + CD_EXTRA_LEN, true);
    const commentLen = cdView.getUint16(pos + CD_COMMENT_LEN, true);
    const localHeaderOffset = cdView.getUint32(pos + CD_LOCAL_HEADER_OFFSET, true);
    const filename = new TextDecoder().decode(new Uint8Array(cdBuf, pos + CD_FILENAME_DATA, fnLen));

    entries.push({ filename, method, localHeaderOffset, compressedSize, uncompressedSize });
    pos += CD_FIXED_SIZE + fnLen + extraLen + commentLen;
  }

  return entries;
}

// Reads the compressed payload of a ZIP entry using a direct file seek.
// Pass maxCompressedBytes to read only a bounded slice — safe for huge files.
export async function readEntryData(
  file: File,
  entry: ZipEntry,
  maxCompressedBytes?: number,
): Promise<{ data: ArrayBuffer; partial: boolean }> {
  const lhBuf = await file
    .slice(entry.localHeaderOffset, entry.localHeaderOffset + LFH_FIXED_SIZE)
    .arrayBuffer();
  const lhView = new DataView(lhBuf);

  if (lhView.getUint32(0, true) !== SIG_LOCAL_HEADER) {
    throw new Error('Invalid local file header');
  }

  const fnLen = lhView.getUint16(LFH_FILENAME_LEN, true);
  const extraLen = lhView.getUint16(LFH_EXTRA_LEN, true);
  const dataStart = entry.localHeaderOffset + LFH_FIXED_SIZE + fnLen + extraLen;

  const limit = maxCompressedBytes !== undefined
    ? Math.min(maxCompressedBytes, entry.compressedSize)
    : entry.compressedSize;

  const data = await file.slice(dataStart, dataStart + limit).arrayBuffer();
  return { data, partial: limit < entry.compressedSize };
}

// Decompresses a ZIP entry payload to a UTF-8 string using the browser's native
// DecompressionStream. For partial=true, a truncation error at the end is expected
// and swallowed — the caller gets whatever was successfully decompressed.
export async function decompressToText(
  data: ArrayBuffer,
  method: number,
  partial = false,
): Promise<string> {
  if (method === 0) return new TextDecoder().decode(data); // stored — no compression
  if (method !== 8) throw new Error(`Unsupported ZIP compression method: ${method}`);

  const ds = new DecompressionStream('deflate-raw');
  const writer = ds.writable.getWriter();
  const reader = ds.readable.getReader();
  const decoder = new TextDecoder('utf-8', { stream: true });
  let text = '';

  const pump = (async () => {
    try { await writer.write(new Uint8Array(data)); await writer.close(); }
    catch { /* partial write — expected */ }
  })();

  try {
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      text += decoder.decode(value, { stream: true });
    }
  } catch {
    if (!partial) throw new Error('Decompression failed');
    // Truncation mid-stream is expected when we read less than the full entry.
  }

  await pump;
  return text;
}
