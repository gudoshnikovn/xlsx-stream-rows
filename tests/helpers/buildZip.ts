/**
 * Hand-rolled ZIP archive builder for tests.
 *
 * Produces a byte-accurate APPNOTE.TXT-compliant archive with zero, one, or
 * many entries (stored or deflate-raw). CRC-32 fields are left as zero — the
 * library under test does not verify CRC, and zeroing keeps the helper small.
 *
 * Optional knobs let tests forge edge cases:
 *   - `comment`        — non-empty EOCD comment (exercises the EOCD scan).
 *   - `forceZip64`     — write ZIP64 sentinels into the EOCD totals.
 *   - `corruptLfh`     — clobber the LFH signature of a chosen entry.
 */

const SIG_LFH = 0x04034b50;
const SIG_CD = 0x02014b50;
const SIG_EOCD = 0x06054b50;

export interface EntrySpec {
  name: string;
  data: Uint8Array;
  method: 0 | 8;
}

export interface BuildOptions {
  comment?: Uint8Array;
  forceZip64?: 'totalEntries' | 'cdSize' | 'cdOffset';
  corruptLfh?: number; // entry index whose LFH signature should be smashed
}

interface BuiltEntry {
  spec: EntrySpec;
  payload: Uint8Array;       // bytes actually written (compressed or stored)
  uncompressedSize: number;
  compressedSize: number;
  localHeaderOffset: number;
}

async function deflateRaw(data: Uint8Array): Promise<Uint8Array> {
  const cs = new CompressionStream('deflate-raw');
  const writer = cs.writable.getWriter();
  // CompressionStream insists on a BufferSource; copy to a plain Uint8Array
  // to avoid SharedArrayBuffer / typed-array-view issues.
  void writer.write(new Uint8Array(data));
  void writer.close();

  const reader = cs.readable.getReader();
  const chunks: Uint8Array[] = [];
  let total = 0;
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    chunks.push(value);
    total += value.byteLength;
  }
  const out = new Uint8Array(total);
  let pos = 0;
  for (const c of chunks) {
    out.set(c, pos);
    pos += c.byteLength;
  }
  return out;
}

function utf8(s: string): Uint8Array {
  return new TextEncoder().encode(s);
}

function writeU16(view: DataView, offset: number, value: number): void {
  view.setUint16(offset, value, true);
}

function writeU32(view: DataView, offset: number, value: number): void {
  view.setUint32(offset, value, true);
}

export async function buildZip(
  entries: EntrySpec[],
  options: BuildOptions = {},
): Promise<Uint8Array> {
  // Build payload + per-entry record up front so we know offsets.
  const built: BuiltEntry[] = [];
  let cursor = 0;

  for (const spec of entries) {
    const payload =
      spec.method === 8 ? await deflateRaw(spec.data) : new Uint8Array(spec.data);
    const nameBytes = utf8(spec.name);
    const localHeaderOffset = cursor;

    built.push({
      spec,
      payload,
      uncompressedSize: spec.data.byteLength,
      compressedSize: payload.byteLength,
      localHeaderOffset,
    });

    // LFH (30 fixed) + filename + payload
    cursor += 30 + nameBytes.byteLength + payload.byteLength;
  }

  const cdStart = cursor;

  // Compute CD size.
  let cdSize = 0;
  for (const e of built) cdSize += 46 + utf8(e.spec.name).byteLength;
  const eocdStart = cdStart + cdSize;

  const comment = options.comment ?? new Uint8Array(0);
  const totalSize = eocdStart + 22 + comment.byteLength;

  const buf = new Uint8Array(totalSize);
  const view = new DataView(buf.buffer);

  // ─── Local file headers + payloads ─────────────────────────────────────────
  for (let i = 0; i < built.length; i++) {
    const e = built[i]!;
    const nameBytes = utf8(e.spec.name);
    const off = e.localHeaderOffset;
    const sig = options.corruptLfh === i ? 0xdeadbeef : SIG_LFH;
    writeU32(view, off + 0, sig);
    writeU16(view, off + 4, 20); // version needed
    writeU16(view, off + 6, 0); // flags
    writeU16(view, off + 8, e.spec.method);
    writeU16(view, off + 10, 0); // mod time
    writeU16(view, off + 12, 0); // mod date
    writeU32(view, off + 14, 0); // crc32 (left zero)
    writeU32(view, off + 18, e.compressedSize);
    writeU32(view, off + 22, e.uncompressedSize);
    writeU16(view, off + 26, nameBytes.byteLength);
    writeU16(view, off + 28, 0); // extra length
    buf.set(nameBytes, off + 30);
    buf.set(e.payload, off + 30 + nameBytes.byteLength);
  }

  // ─── Central directory ─────────────────────────────────────────────────────
  let cdPos = cdStart;
  for (const e of built) {
    const nameBytes = utf8(e.spec.name);
    writeU32(view, cdPos + 0, SIG_CD);
    writeU16(view, cdPos + 4, 20); // version made by
    writeU16(view, cdPos + 6, 20); // version needed
    writeU16(view, cdPos + 8, 0); // flags
    writeU16(view, cdPos + 10, e.spec.method);
    writeU16(view, cdPos + 12, 0); // mod time
    writeU16(view, cdPos + 14, 0); // mod date
    writeU32(view, cdPos + 16, 0); // crc32
    writeU32(view, cdPos + 20, e.compressedSize);
    writeU32(view, cdPos + 24, e.uncompressedSize);
    writeU16(view, cdPos + 28, nameBytes.byteLength);
    writeU16(view, cdPos + 30, 0); // extra length
    writeU16(view, cdPos + 32, 0); // comment length
    writeU16(view, cdPos + 34, 0); // disk number
    writeU16(view, cdPos + 36, 0); // internal attrs
    writeU32(view, cdPos + 38, 0); // external attrs
    writeU32(view, cdPos + 42, e.localHeaderOffset);
    buf.set(nameBytes, cdPos + 46);
    cdPos += 46 + nameBytes.byteLength;
  }

  // ─── EOCD ──────────────────────────────────────────────────────────────────
  const totalEntries =
    options.forceZip64 === 'totalEntries' ? 0xffff : built.length;
  const cdSizeField = options.forceZip64 === 'cdSize' ? 0xffffffff : cdSize;
  const cdOffsetField = options.forceZip64 === 'cdOffset' ? 0xffffffff : cdStart;

  writeU32(view, eocdStart + 0, SIG_EOCD);
  writeU16(view, eocdStart + 4, 0); // disk number
  writeU16(view, eocdStart + 6, 0); // disk with CD
  writeU16(view, eocdStart + 8, totalEntries); // entries on this disk
  writeU16(view, eocdStart + 10, totalEntries); // total entries
  writeU32(view, eocdStart + 12, cdSizeField);
  writeU32(view, eocdStart + 16, cdOffsetField);
  writeU16(view, eocdStart + 20, comment.byteLength);
  if (comment.byteLength > 0) buf.set(comment, eocdStart + 22);

  return buf;
}

/** Wrap the given bytes in a `File` so tests can call APIs that take `File`. */
export function asFile(bytes: Uint8Array, name = 'test.zip'): File {
  // Slice to a fresh ArrayBuffer to keep the File constructor's BlobPart typing
  // happy under TS 5.7's Uint8Array<ArrayBufferLike> generic.
  return new File([bytes.slice().buffer], name);
}

/** Drain a ReadableStream<Uint8Array> into a single concatenated buffer. */
export async function drainStream(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  const reader = stream.getReader();
  const chunks: Uint8Array[] = [];
  let total = 0;
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    chunks.push(value);
    total += value.byteLength;
  }
  const out = new Uint8Array(total);
  let pos = 0;
  for (const c of chunks) {
    out.set(c, pos);
    pos += c.byteLength;
  }
  return out;
}
