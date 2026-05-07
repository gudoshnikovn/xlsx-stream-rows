import { describe, expect, it } from 'vitest';

import {
  InvalidLocalHeaderError,
  NotAZipError,
  UnsupportedCompressionError,
  Zip64NotSupportedError,
  openDecompressedStream,
  readEntryToString,
  readZipEntries,
} from '../src/index.js';
import { asFile, buildZip, drainStream } from './helpers/buildZip.js';

const utf8 = (s: string): Uint8Array => new TextEncoder().encode(s);
const utf8Decode = (b: Uint8Array): string => new TextDecoder().decode(b);

describe('readZipEntries', () => {
  it('lists a single stored entry', async () => {
    const data = utf8('Hello, World!');
    const zip = await buildZip([{ name: 'hello.txt', data, method: 0 }]);
    const entries = await readZipEntries(asFile(zip));

    expect(entries).toHaveLength(1);
    expect(entries[0]).toMatchObject({
      filename: 'hello.txt',
      method: 0,
      compressedSize: data.byteLength,
      uncompressedSize: data.byteLength,
    });
  });

  it('lists a single deflated entry', async () => {
    const data = utf8('x'.repeat(2048)); // compresses well
    const zip = await buildZip([{ name: 'big.txt', data, method: 8 }]);
    const entries = await readZipEntries(asFile(zip));

    expect(entries).toHaveLength(1);
    expect(entries[0]?.method).toBe(8);
    expect(entries[0]?.uncompressedSize).toBe(2048);
    expect(entries[0]?.compressedSize).toBeLessThan(2048);
  });

  it('lists multiple entries in archive order', async () => {
    const zip = await buildZip([
      { name: 'a.txt', data: utf8('one'), method: 0 },
      { name: 'b.txt', data: utf8('two'), method: 8 },
      { name: 'nested/c.txt', data: utf8('three'), method: 0 },
    ]);
    const entries = await readZipEntries(asFile(zip));
    expect(entries.map((e) => e.filename)).toEqual(['a.txt', 'b.txt', 'nested/c.txt']);
  });

  it('finds the EOCD when a non-empty comment is present', async () => {
    // Comment of moderate length — exercises the backward EOCD scan.
    const comment = new Uint8Array(1234).fill(0x41); // 1234 'A's
    const zip = await buildZip([{ name: 'a.txt', data: utf8('hi'), method: 0 }], {
      comment,
    });
    const entries = await readZipEntries(asFile(zip));
    expect(entries[0]?.filename).toBe('a.txt');
  });

  it('handles UTF-8 filenames', async () => {
    const zip = await buildZip([{ name: 'тест/файл.txt', data: utf8('x'), method: 0 }]);
    const entries = await readZipEntries(asFile(zip));
    expect(entries[0]?.filename).toBe('тест/файл.txt');
  });

  it('throws NotAZipError on garbage', async () => {
    const noise = new Uint8Array(1024);
    crypto.getRandomValues(noise);
    await expect(readZipEntries(asFile(noise))).rejects.toBeInstanceOf(NotAZipError);
  });

  it('throws NotAZipError on file too small to contain an EOCD', async () => {
    await expect(readZipEntries(asFile(new Uint8Array(10)))).rejects.toBeInstanceOf(
      NotAZipError,
    );
  });

  it('throws Zip64NotSupportedError when EOCD totalEntries is the sentinel', async () => {
    const zip = await buildZip([{ name: 'a.txt', data: utf8('x'), method: 0 }], {
      forceZip64: 'totalEntries',
    });
    await expect(readZipEntries(asFile(zip))).rejects.toBeInstanceOf(Zip64NotSupportedError);
  });

  it('throws Zip64NotSupportedError when CD size is the sentinel', async () => {
    const zip = await buildZip([{ name: 'a.txt', data: utf8('x'), method: 0 }], {
      forceZip64: 'cdSize',
    });
    await expect(readZipEntries(asFile(zip))).rejects.toBeInstanceOf(Zip64NotSupportedError);
  });

  it('throws Zip64NotSupportedError when CD offset is the sentinel', async () => {
    const zip = await buildZip([{ name: 'a.txt', data: utf8('x'), method: 0 }], {
      forceZip64: 'cdOffset',
    });
    await expect(readZipEntries(asFile(zip))).rejects.toBeInstanceOf(Zip64NotSupportedError);
  });
});

describe('openDecompressedStream', () => {
  it('round-trips a stored entry byte-for-byte', async () => {
    const data = utf8('stored payload — 42');
    const zip = await buildZip([{ name: 'a.bin', data, method: 0 }]);
    const file = asFile(zip);
    const entry = (await readZipEntries(file))[0]!;
    const stream = await openDecompressedStream(file, entry);
    const got = await drainStream(stream);
    expect(utf8Decode(got)).toBe('stored payload — 42');
  });

  it('round-trips a deflated entry byte-for-byte', async () => {
    const original = 'lorem ipsum '.repeat(500);
    const zip = await buildZip([{ name: 'big.txt', data: utf8(original), method: 8 }]);
    const file = asFile(zip);
    const entry = (await readZipEntries(file))[0]!;
    const stream = await openDecompressedStream(file, entry);
    const got = await drainStream(stream);
    expect(utf8Decode(got)).toBe(original);
  });

  it('throws UnsupportedCompressionError for non-stored / non-deflate methods', async () => {
    const zip = await buildZip([{ name: 'a.txt', data: utf8('x'), method: 0 }]);
    const file = asFile(zip);
    const entry = (await readZipEntries(file))[0]!;
    // Forge an unsupported method without rebuilding the archive.
    const forged = { ...entry, method: 12 };
    await expect(openDecompressedStream(file, forged)).rejects.toBeInstanceOf(
      UnsupportedCompressionError,
    );
  });

  it('throws InvalidLocalHeaderError when the LFH signature is corrupted', async () => {
    const zip = await buildZip(
      [{ name: 'a.txt', data: utf8('x'), method: 0 }],
      { corruptLfh: 0 },
    );
    const file = asFile(zip);
    const entry = (await readZipEntries(file))[0]!;
    await expect(openDecompressedStream(file, entry)).rejects.toBeInstanceOf(
      InvalidLocalHeaderError,
    );
  });
});

describe('readEntryToString', () => {
  it('decodes a deflated UTF-8 XML part', async () => {
    const xml = '<?xml version="1.0"?><root><a>тест</a></root>';
    const zip = await buildZip([{ name: 'doc.xml', data: utf8(xml), method: 8 }]);
    const file = asFile(zip);
    const entry = (await readZipEntries(file))[0]!;
    expect(await readEntryToString(file, entry, 1024 * 1024)).toBe(xml);
  });

  it('throws when the entry exceeds the size cap', async () => {
    const xml = 'x'.repeat(10_000);
    const zip = await buildZip([{ name: 'doc.xml', data: utf8(xml), method: 8 }]);
    const file = asFile(zip);
    const entry = (await readZipEntries(file))[0]!;
    await expect(readEntryToString(file, entry, 1024)).rejects.toThrow(/exceeds limit/);
  });
});
