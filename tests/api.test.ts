/**
 * Tests for the unified `openWorkbook` / `streamRows` / `readRows` API.
 * Verifies format detection (magic bytes + extension fallback) and routing.
 */

import { describe, expect, it } from 'vitest';
import * as XLSX from 'xlsx';

import {
  FormatNotSupportedError,
  detectFormat,
  openWorkbook,
  readRows,
  streamRows,
} from '../src/index.js';
import { asFile } from './helpers/buildZip.js';
import { buildXlsx } from './helpers/buildXlsx.js';

const utf8 = (s: string): Uint8Array => new TextEncoder().encode(s);

function asNamedFile(bytes: Uint8Array, name: string): File {
  return new File([bytes.slice().buffer], name);
}

function buildXls(rows: unknown[][]): Uint8Array {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), 'S1');
  return new Uint8Array(XLSX.write(wb, { bookType: 'xls', type: 'array' }) as ArrayBuffer);
}

describe('detectFormat', () => {
  it('detects XLSX by ZIP magic regardless of extension', async () => {
    const xlsx = await buildXlsx({
      sheets: [{ name: 'A', sheetData: '<sheetData/>' }],
    });
    expect(await detectFormat(asNamedFile(xlsx, 'mystery'))).toBe('xlsx');
  });

  it('detects XLS by OLE2 magic regardless of extension', async () => {
    expect(await detectFormat(asNamedFile(buildXls([[1]]), 'mystery'))).toBe('xls');
  });

  it('falls back to .csv extension when no magic matches', async () => {
    expect(await detectFormat(asNamedFile(utf8('a,b\n'), 'data.csv'))).toBe('csv');
  });

  it('falls back to .xlsx extension on an empty file', async () => {
    expect(await detectFormat(asNamedFile(new Uint8Array(0), 'empty.xlsx'))).toBe('xlsx');
  });

  it('throws FormatNotSupportedError for unknown extensions and no magic', async () => {
    await expect(detectFormat(asNamedFile(utf8('x'), 'mystery.dat'))).rejects.toBeInstanceOf(
      FormatNotSupportedError,
    );
  });
});

describe('openWorkbook', () => {
  it('routes XLSX files to the XLSX adapter', async () => {
    const xlsx = await buildXlsx({
      sheets: [
        { name: 'Sheet1', sheetData: '<sheetData/>' },
        { name: 'Sheet2', sheetData: '<sheetData/>' },
      ],
    });
    const info = await openWorkbook(asFile(xlsx, 'wb.xlsx'));
    expect(info).toEqual({
      filename: 'wb.xlsx',
      sheetNames: ['Sheet1', 'Sheet2'],
      format: 'xlsx',
    });
  });

  it('routes CSV files to the CSV adapter', async () => {
    const info = await openWorkbook(asNamedFile(utf8('a,b\n'), 'data.csv'));
    expect(info.format).toBe('csv');
    expect(info.sheetNames).toEqual(['data']);
  });

  it('routes XLS files to the XLS adapter', async () => {
    const info = await openWorkbook(asNamedFile(buildXls([[1]]), 'book.xls'));
    expect(info.format).toBe('xls');
    expect(info.sheetNames).toEqual(['S1']);
  });
});

describe('streamRows / readRows', () => {
  it('streams rows from XLSX through the unified API', async () => {
    const xlsx = await buildXlsx({
      sharedStrings: ['Alice'],
      sheets: [
        {
          name: 'A',
          sheetData:
            '<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>42</v></c></row></sheetData>',
        },
      ],
    });
    expect(await readRows(asFile(xlsx, 'wb.xlsx'))).toEqual([['Alice', 42]]);
  });

  it('streams rows from CSV through the unified API', async () => {
    expect(await readRows(asNamedFile(utf8('a,b\n1,2\n'), 'd.csv'))).toEqual([
      ['a', 'b'],
      ['1', '2'],
    ]);
  });

  it('streams rows from XLS through the unified API', async () => {
    const xls = buildXls([
      ['n', 'v'],
      ['x', 1],
    ]);
    expect(await readRows(asNamedFile(xls, 'book.xls'))).toEqual([
      ['n', 'v'],
      ['x', 1],
    ]);
  });

  it('honors maxRows uniformly across formats', async () => {
    const csv = utf8(Array.from({ length: 50 }, (_, i) => i).join('\n'));
    const rows = await readRows(asNamedFile(csv, 'big.csv'), { maxRows: 3 });
    expect(rows.map((r) => r[0])).toEqual(['0', '1', '2']);
  });

  it('aborts before the first row when the signal is already aborted', async () => {
    const ac = new AbortController();
    const reason = new Error('pre-abort');
    ac.abort(reason);
    await expect(
      readRows(asNamedFile(utf8('a\nb\n'), 'd.csv'), { signal: ac.signal }),
    ).rejects.toBe(reason);
  });
});
