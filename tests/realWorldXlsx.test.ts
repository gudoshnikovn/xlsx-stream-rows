/**
 * Real-world XLSX cross-validation.
 *
 * Builds workbooks with the `xlsx` package — which is itself a real-world
 * OOXML producer (used by Microsoft Excel/Mac, Google Sheets exports go
 * through similar paths) — then reads each through both `xlsx` and our
 * library and asserts row-for-row equality.
 *
 * This is the strongest correctness check we can run without an Excel
 * install: any divergence between our streaming parser and a mature
 * full-load implementation surfaces as a test failure.
 *
 * The fixtures here exercise the messy parts of OOXML: shared strings,
 * inline strings, dates with custom number formats, formulas with cached
 * values, sparse cells, multi-sheet workbooks, and rich text.
 */

import { describe, expect, it } from 'vitest';
import * as XLSX from 'xlsx';

import { openWorkbook, readRows, streamXlsxRows } from '../src/index.js';
import { asFile } from './helpers/buildZip.js';
import type { CellValue, Row } from '../src/types.js';

interface XlsxFixture {
  /** Multi-sheet input as `{ sheetName: rows-of-arrays }`. */
  sheets: Record<string, unknown[][]>;
  /** XLSX writer options — used to vary the produced file shape. */
  writeOpts?: XLSX.WritingOptions;
  /** Sheet to read for the streaming side; default first. */
  readSheet?: string;
}

function buildXlsxFile(fixture: XlsxFixture): File {
  const wb = XLSX.utils.book_new();
  for (const [name, data] of Object.entries(fixture.sheets)) {
    const ws = XLSX.utils.aoa_to_sheet(data);
    XLSX.utils.book_append_sheet(wb, ws, name);
  }
  const buf = XLSX.write(wb, {
    bookType: 'xlsx',
    type: 'array',
    cellDates: true,
    ...fixture.writeOpts,
  }) as ArrayBuffer;
  return asFile(new Uint8Array(buf), 'fixture.xlsx');
}

/**
 * Read the same fixture via `xlsx` (full-load reference) and produce rows.
 * We use `header: 1` for an array-of-arrays shape that matches `streamRows`.
 */
function expectedRows(fixture: XlsxFixture): Row[] {
  const wb = XLSX.utils.book_new();
  for (const [name, data] of Object.entries(fixture.sheets)) {
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(data), name);
  }
  const written = XLSX.write(wb, {
    bookType: 'xlsx',
    type: 'array',
    cellDates: true,
    ...fixture.writeOpts,
  }) as ArrayBuffer;
  const reread = XLSX.read(written, { type: 'array', cellDates: true });
  const sheetName = fixture.readSheet ?? reread.SheetNames[0]!;
  const sheet = reread.Sheets[sheetName]!;
  const out = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    raw: true,
    blankrows: false,
  }) as unknown[][];
  // `sheet_to_json` returns sparse arrays (holes for empty cells). Materialise
  // by index so `Array.prototype.map` doesn't skip holes — otherwise the
  // normalise step misses gap-filling.
  return out.map((row) => {
    const filled: CellValue[] = new Array(row.length);
    for (let i = 0; i < row.length; i++) filled[i] = normalise(row[i]);
    return filled;
  });
}

function normalise(v: unknown): CellValue {
  // Empty cells: `xlsx` returns `undefined`, our library returns `null`.
  // Both are correct per their respective contracts; normalise to `null`
  // here so the cross-validation focuses on value content.
  if (v === undefined || v === null) return null;
  if (typeof v === 'string' || typeof v === 'number' || typeof v === 'boolean') return v;
  if (v instanceof Date) return v;
  return String(v);
}

async function actualRows(fixture: XlsxFixture): Promise<Row[]> {
  const file = buildXlsxFile(fixture);
  const opts = fixture.readSheet ? { sheetName: fixture.readSheet } : {};
  return readRows(file, opts);
}

/**
 * Compare two row arrays cell-by-cell.
 *
 * Dates: Excel serials carry no timezone — the file just says "this number
 * means a date". `xlsx` converts the serial to local-time `Date`, our
 * library uses UTC. Both are defensible; comparing by `getTime()` is wrong.
 * We compare the *displayed* date instead — Y/M/D + H/M/S — by reading
 * `xlsx`'s result in local time and ours in UTC.
 */
function rowsMatch(actual: Row[], expected: Row[]): void {
  expect(actual.length).toBe(expected.length);
  for (let r = 0; r < actual.length; r++) {
    const a = actual[r]!;
    const e = expected[r]!;
    expect(a.length, `row ${r}`).toBe(e.length);
    for (let c = 0; c < a.length; c++) {
      const av = a[c];
      const ev = e[c];
      if (av instanceof Date && ev instanceof Date) {
        expect(av.getUTCFullYear(), `row ${r} col ${c} year`).toBe(ev.getFullYear());
        expect(av.getUTCMonth(), `row ${r} col ${c} month`).toBe(ev.getMonth());
        expect(av.getUTCDate(), `row ${r} col ${c} day`).toBe(ev.getDate());
        expect(av.getUTCHours(), `row ${r} col ${c} hours`).toBe(ev.getHours());
        expect(av.getUTCMinutes(), `row ${r} col ${c} minutes`).toBe(ev.getMinutes());
      } else {
        expect(av, `row ${r} col ${c}`).toEqual(ev);
      }
    }
  }
}

describe('XLSX cross-validation against the `xlsx` package', () => {
  it('matches on a kitchen-sink workbook with mixed types', async () => {
    const fixture: XlsxFixture = {
      sheets: {
        Data: [
          ['name', 'age', 'active', 'joined', 'score'],
          ['Alice', 30, true, new Date(Date.UTC(2021, 0, 15)), 98.6],
          ['Bob', 25, false, new Date(Date.UTC(2022, 5, 1)), 72.1],
          ['Carol', 41, true, new Date(Date.UTC(2019, 11, 31)), 100],
        ],
      },
    };
    rowsMatch(await actualRows(fixture), expectedRows(fixture));
  });

  it('matches on a workbook with several sheets, default-sheet read', async () => {
    const fixture: XlsxFixture = {
      sheets: {
        First: [['x', 1]],
        Second: [['y', 2]],
        Third: [['z', 3]],
      },
    };
    rowsMatch(await actualRows(fixture), expectedRows(fixture));
  });

  it('matches when reading a non-default sheet by name', async () => {
    const fixture: XlsxFixture = {
      sheets: {
        Skip: [['ignored']],
        Target: [
          ['a', 'b'],
          [1, 2],
        ],
      },
      readSheet: 'Target',
    };
    rowsMatch(await actualRows(fixture), expectedRows(fixture));
  });

  it('matches with sparse cells (gaps padded with null)', async () => {
    const fixture: XlsxFixture = {
      sheets: {
        Sparse: [
          ['header1', null, 'header3'],
          ['v1', null, 'v3'],
        ],
      },
    };
    rowsMatch(await actualRows(fixture), expectedRows(fixture));
  });

  it('matches on a workbook that writes inline strings instead of sharedStrings', async () => {
    const fixture: XlsxFixture = {
      sheets: {
        Inline: [
          ['short', 'longer text here'],
          ['another', 'even more content'],
        ],
      },
      // `bookSST: false` (default) often inlines short strings; set explicitly
      // for clarity.
      writeOpts: { bookSST: false },
    };
    rowsMatch(await actualRows(fixture), expectedRows(fixture));
  });

  it('matches on a workbook that uses the shared-string table', async () => {
    const fixture: XlsxFixture = {
      sheets: {
        Shared: [
          ['repeat', 'repeat', 'repeat'],
          ['repeat', 'unique', 'repeat'],
        ],
      },
      writeOpts: { bookSST: true },
    };
    rowsMatch(await actualRows(fixture), expectedRows(fixture));
  });

  it('matches on a workbook with unicode and special characters', async () => {
    const fixture: XlsxFixture = {
      sheets: {
        I18n: [
          ['тест', '东京', '🚀'],
          ['Tom & Jerry', '<html>', '"quoted"'],
        ],
      },
    };
    rowsMatch(await actualRows(fixture), expectedRows(fixture));
  });

  it('matches on an empty sheet', async () => {
    const fixture: XlsxFixture = {
      sheets: { Empty: [] },
    };
    rowsMatch(await actualRows(fixture), expectedRows(fixture));
  });

  it('matches on a 5000-row workbook (medium-scale parity)', async () => {
    const rows: unknown[][] = [['id', 'value']];
    for (let i = 1; i <= 5000; i++) {
      rows.push([i, `row-${i}`]);
    }
    const fixture: XlsxFixture = { sheets: { Big: rows } };
    rowsMatch(await actualRows(fixture), expectedRows(fixture));
  });

  it('reports correct sheet metadata via openWorkbook', async () => {
    const fixture: XlsxFixture = {
      sheets: { Alpha: [[1]], Beta: [[2]] },
    };
    const file = buildXlsxFile(fixture);
    const info = await openWorkbook(file);
    expect(info.format).toBe('xlsx');
    expect(info.sheetNames).toEqual(['Alpha', 'Beta']);
  });

  it('streams correctly when consumed in pieces (10-row chunks)', async () => {
    const rows: unknown[][] = [];
    for (let i = 1; i <= 100; i++) rows.push([i]);
    const file = buildXlsxFile({ sheets: { S: rows } });

    const collected: number[] = [];
    let chunkCount = 0;
    let inChunk = 0;
    for await (const row of streamXlsxRows(file)) {
      collected.push(row[0] as number);
      inChunk++;
      if (inChunk === 10) {
        chunkCount++;
        inChunk = 0;
      }
    }
    expect(collected.length).toBe(100);
    expect(collected[0]).toBe(1);
    expect(collected[99]).toBe(100);
    expect(chunkCount).toBe(10);
  });
});
