import { describe, expect, it } from 'vitest';

import { openCsvWorkbook, streamCsvRows } from '../src/csvStreamReader.js';
import type { Row } from '../src/types.js';

const utf8 = (s: string): Uint8Array => new TextEncoder().encode(s);

function csvFile(content: string | Uint8Array, name = 'data.csv'): File {
  const bytes = typeof content === 'string' ? utf8(content) : content;
  return new File([bytes.slice().buffer], name);
}

async function collect<T>(iter: AsyncIterable<T>): Promise<T[]> {
  const out: T[] = [];
  for await (const x of iter) out.push(x);
  return out;
}

describe('openCsvWorkbook', () => {
  it('returns a single pseudo-sheet named after the file', async () => {
    const info = await openCsvWorkbook(csvFile('a,b\n', 'data.csv'));
    expect(info).toEqual({ filename: 'data.csv', sheetNames: ['data'], format: 'csv' });
  });

  it('uses the full filename when there is no extension', async () => {
    const info = await openCsvWorkbook(csvFile('a,b\n', 'README'));
    expect(info.sheetNames).toEqual(['README']);
  });
});

describe('streamCsvRows', () => {
  it('streams rows from a basic CSV file', async () => {
    const rows = await collect(streamCsvRows(csvFile('a,b,c\n1,2,3\n')));
    expect(rows).toEqual([
      ['a', 'b', 'c'],
      ['1', '2', '3'],
    ]);
  });

  it('strips a UTF-8 BOM', async () => {
    const bytes = new Uint8Array([0xef, 0xbb, 0xbf, ...utf8('a,b\n')]);
    const rows = await collect(streamCsvRows(csvFile(bytes)));
    expect(rows).toEqual([['a', 'b']]);
  });

  it('decodes UTF-16 LE via BOM', async () => {
    // UTF-16 LE BOM + "ab\n"
    const bytes = new Uint8Array([
      0xff, 0xfe,
      'a'.charCodeAt(0), 0x00,
      'b'.charCodeAt(0), 0x00,
      0x0a, 0x00,
    ]);
    const rows = await collect(streamCsvRows(csvFile(bytes)));
    expect(rows).toEqual([['ab']]);
  });

  it('honors maxRows', async () => {
    const big = Array.from({ length: 1000 }, (_, i) => `${i}`).join('\n');
    const rows = await collect(streamCsvRows(csvFile(big), { maxRows: 5 }));
    expect(rows.map((r) => r[0])).toEqual(['0', '1', '2', '3', '4']);
  });

  it('stops on `break` out of the for-await loop', async () => {
    const big = Array.from({ length: 1000 }, (_, i) => `${i}`).join('\n');
    const collected: Row[] = [];
    for await (const r of streamCsvRows(csvFile(big))) {
      if (Number(r[0]) > 2) break;
      collected.push(r);
    }
    expect(collected.map((r) => r[0])).toEqual(['0', '1', '2']);
  });

  it('rejects with the AbortSignal reason when aborted before iteration', async () => {
    const ac = new AbortController();
    const reason = new Error('cancel');
    ac.abort(reason);
    await expect(
      collect(streamCsvRows(csvFile('a\nb\nc\n'), { signal: ac.signal })),
    ).rejects.toBe(reason);
  });
});
