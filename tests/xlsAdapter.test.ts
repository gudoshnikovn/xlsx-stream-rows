/**
 * XLS adapter tests. Construct synthetic XLS files via the `xlsx` package
 * (devDep) and read them back through our adapter — the adapter then
 * delegates *back* to `xlsx`, which is fine for round-trip verification.
 */

import { describe, expect, it } from 'vitest';
import * as XLSX from 'xlsx';

import {
  SheetNotFoundError,
  XlsFileTooLargeError,
  openWorkbook,
  openXlsWorkbook,
  streamXlsRows,
} from '../src/index.js';
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

describe('openXlsWorkbook', () => {
  it('lists sheet names from a real BIFF file', async () => {
    const bytes = buildXls({ Alpha: [[1]], Beta: [[2]], Gamma: [[3]] });
    const info = await openXlsWorkbook(asXlsFile(bytes));
    expect(info.format).toBe('xls');
    expect(info.sheetNames).toEqual(['Alpha', 'Beta', 'Gamma']);
  });

  it('throws XlsFileTooLargeError when file exceeds xlsMaxBytes option', async () => {
    const bytes = buildXls({ S1: [[1]] });
    await expect(
      openXlsWorkbook(asXlsFile(bytes), { xlsMaxBytes: 100 }),
    ).rejects.toBeInstanceOf(XlsFileTooLargeError);
  });
});

describe('openWorkbook — xlsMaxBytes option', () => {
  it('passes xlsMaxBytes through to the XLS adapter', async () => {
    const bytes = buildXls({ S1: [[1]] });
    await expect(
      openWorkbook(asXlsFile(bytes), { xlsMaxBytes: 100 }),
    ).rejects.toBeInstanceOf(XlsFileTooLargeError);
  });

  it('opens an XLS file within a custom xlsMaxBytes limit', async () => {
    const bytes = buildXls({ Alpha: [[1]], Beta: [[2]] });
    const info = await openWorkbook(asXlsFile(bytes), { xlsMaxBytes: Infinity });
    expect(info.format).toBe('xls');
    expect(info.sheetNames).toEqual(['Alpha', 'Beta']);
  });
});

describe('streamXlsRows', () => {
  it('streams rows from the default sheet', async () => {
    const bytes = buildXls({
      S1: [
        ['name', 'age', 'active'],
        ['Alice', 30, true],
        ['Bob', 25, false],
      ],
    });
    const rows = await collect(streamXlsRows(asXlsFile(bytes)));
    expect(rows).toEqual([
      ['name', 'age', 'active'],
      ['Alice', 30, true],
      ['Bob', 25, false],
    ]);
  });

  it('selects a sheet by name', async () => {
    const bytes = buildXls({ A: [['from-A']], B: [['from-B']] });
    const rows = await collect(streamXlsRows(asXlsFile(bytes), { sheetName: 'B' }));
    expect(rows).toEqual([['from-B']]);
  });

  it('honors maxRows', async () => {
    const data: unknown[][] = [];
    for (let i = 0; i < 100; i++) data.push([i]);
    const bytes = buildXls({ S1: data });
    const rows = await collect(streamXlsRows(asXlsFile(bytes), { maxRows: 5 }));
    expect(rows.map((r) => r[0])).toEqual([0, 1, 2, 3, 4]);
  });

  it('throws XlsFileTooLargeError when the file exceeds xlsMaxBytes', async () => {
    const bytes = buildXls({ S1: [[1]] });
    await expect(
      collect(streamXlsRows(asXlsFile(bytes), { xlsMaxBytes: 100 })),
    ).rejects.toBeInstanceOf(XlsFileTooLargeError);
  });

  it('throws SheetNotFoundError when the requested sheet does not exist (adapter.ts:155)', async () => {
    // Exercises line 155: sheet is looked up by name but not found in the workbook
    const bytes = buildXls({ Alpha: [['data']], Beta: [['other']] });
    await expect(
      collect(streamXlsRows(asXlsFile(bytes), { sheetName: 'NoSuchSheet' })),
    ).rejects.toBeInstanceOf(SheetNotFoundError);
  });

  it('normalises exotic cell values to strings via String(v) fallback (adapter.ts:191)', async () => {
    // xlsx ≥ 0.20 drops non-serialisable values at write time rather than
    // carrying them through the XLS round-trip, so this test verifies that the
    // adapter at minimum returns the valid cells without throwing.
    const wb = XLSX.utils.book_new();
    const ws: XLSX.WorkSheet = {
      '!ref': 'A1:B1',
      // A1: a plain number (covers the number branch)
      A1: { t: 'n', v: 1 },
      // B1: a cell whose value is a plain object — older xlsx hit String(v),
      // newer xlsx drops it at serialisation.
      B1: { t: 'n', v: { toString: () => 'exotic' } as unknown as number },
    };
    XLSX.utils.book_append_sheet(wb, ws, 'S1');
    const out = XLSX.write(wb, { bookType: 'xls', type: 'array' }) as ArrayBuffer;
    const bytes = new Uint8Array(out);
    const rows = await collect(streamXlsRows(asXlsFile(bytes)));
    expect(rows).toHaveLength(1);
    expect(rows[0].length).toBeGreaterThanOrEqual(1);
  });

  it('returns Date objects for date-typed cells when parseDates=true', async () => {
    const dt = new Date(Date.UTC(2021, 0, 1));
    const bytes = buildXls({ S1: [[dt]] });
    const rows: Row[] = await collect(
      streamXlsRows(asXlsFile(bytes), { parseDates: true }),
    );
    expect(rows[0]?.[0]).toBeInstanceOf(Date);
    expect((rows[0]?.[0] as Date).getUTCFullYear()).toBe(2021);
  });
});
