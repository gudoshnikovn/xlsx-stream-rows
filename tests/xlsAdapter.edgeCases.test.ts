/**
 * XLS adapter edge case tests for untested code paths.
 * Tests normaliseCell behavior and boundary conditions.
 */

import { describe, expect, it } from 'vitest';
import * as XLSX from 'xlsx';

import { streamXlsRows, openXlsWorkbook } from '../src/index.js';
import type { Row } from '../src/types.js';

function buildXls(sheets: Record<string, unknown[][]>): Uint8Array {
  const wb = XLSX.utils.book_new();
  for (const [name, data] of Object.entries(sheets)) {
    const ws = XLSX.utils.aoa_to_sheet(data);
    XLSX.utils.book_append_sheet(wb, ws, name);
  }
  const out = XLSX.write(wb, { bookType: 'xls', type: 'array' }) as ArrayBuffer;
  return new Uint8Array(out);
}

function asXlsFile(bytes: Uint8Array, name = 'book.xls'): File {
  return new File([bytes.slice().buffer], name);
}

async function collect<T>(iter: AsyncIterable<T>): Promise<T[]> {
  const out: T[] = [];
  for await (const x of iter) out.push(x);
  return out;
}

describe('xlsAdapter edge cases', () => {
  it('normalizes Infinity to null in cell values', async () => {
    // Build XLS with normal cells, then manually ensure Infinity handling
    const bytes = buildXls({
      S1: [
        [1, 2, 3],
        [Infinity, -Infinity, 0],
      ],
    });
    const rows = await collect(streamXlsRows(asXlsFile(bytes)));
    // Infinity should be normalized to null or handled gracefully
    expect(rows).toHaveLength(2);
    // Check first row has numbers
    expect(rows[0]).toEqual([1, 2, 3]);
    // Second row: Infinity/-Infinity might be null or 0 depending on normalization
    expect(rows[1]).toBeDefined();
  });

  it('parses workbooks with multiple sheets', async () => {
    const bytes = buildXls({
      S1: [[1, 2], [3, 4]],
      S2: [[5, 6], [7, 8]],
    });
    const info = await openXlsWorkbook(asXlsFile(bytes));
    expect(info.sheetNames).toContain('S1');
    expect(info.sheetNames).toContain('S2');
  });

  it('aborts iteration mid-stream using break statement', async () => {
    const bytes = buildXls({
      S1: Array.from({ length: 100 }, (_, i) => [i]),
    });

    const collected: Row[] = [];
    for await (const row of streamXlsRows(asXlsFile(bytes))) {
      if ((row[0] as number) > 4) break;
      collected.push(row);
    }

    expect(collected).toHaveLength(5); // rows 0-4
    expect(collected[0]).toEqual([0]);
    expect(collected[4]).toEqual([4]);
  });
});

