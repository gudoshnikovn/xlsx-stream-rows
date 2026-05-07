/**
 * Memory smoke test — the load-bearing claim of the library.
 *
 * Builds a synthetic XLSX with a million rows and verifies two things:
 *
 *   1. **Functional**: the streaming pipeline consumes all rows correctly
 *      from a sheet whose uncompressed XML is far larger than what most
 *      "load-everything" libraries can handle in a tab.
 *   2. **Memory** (when `globalThis.gc` is available): peak heap growth
 *      stays bounded relative to the sheet size — a non-streaming reader
 *      would hold the entire uncompressed sheet (50+ MiB for this fixture);
 *      we assert peak growth < uncompressed sheet size.
 *
 * The test is gated on `XLSX_STREAM_MEMORY_TEST=1` because building the
 * synthetic file takes a few seconds and would slow every CI run.
 *
 * Run it locally:
 *   XLSX_STREAM_MEMORY_TEST=1 NODE_OPTIONS=--expose-gc \
 *     npx vitest run tests/memory.smoke.test.ts --pool=forks --poolOptions.forks.execArgv=--expose-gc
 *
 * Without `--expose-gc`, the heap-bound assertion is skipped and the test
 * runs as a functional smoke test only (still logs measured numbers).
 */

import { describe, expect, it } from 'vitest';

import { streamXlsxRows } from '../src/index.js';
import { asFile } from './helpers/buildZip.js';
import { buildXlsx } from './helpers/buildXlsx.js';

const ENABLED = process.env.XLSX_STREAM_MEMORY_TEST === '1';
const describeMaybe = ENABLED ? describe : describe.skip;

const MIB = 1024 * 1024;

const HAS_GC =
  typeof (globalThis as { gc?: () => void }).gc === 'function';

function heapUsed(): number {
  const gc = (globalThis as { gc?: () => void }).gc;
  if (gc) gc();
  return process.memoryUsage().heapUsed;
}

describeMaybe('memory smoke', () => {
  it(
    'streams 1M rows without holding the sheet in memory',
    async () => {
      const ROWS = 1_000_000;

      // Build the file inside a nested async function so the giant `body`
      // string and intermediate buffers go out of scope before we measure
      // baseline heap. Returning only the File handle keeps the closure
      // empty.
      async function buildBigFile(): Promise<{ file: File; uncompressedSheetBytes: number }> {
        let body = '<sheetData>';
        for (let i = 1; i <= ROWS; i++) {
          body += `<row r="${i}"><c r="A${i}"><v>${i}</v></c></row>`;
        }
        body += '</sheetData>';
        const uncompressedSheetBytes = body.length;
        const xlsx = await buildXlsx({
          deflate: true,
          sheets: [{ name: 'big', sheetData: body }],
        });
        return { file: asFile(xlsx), uncompressedSheetBytes };
      }

      const { file, uncompressedSheetBytes } = await buildBigFile();

      const baseline = heapUsed();
      let peak = 0;

      let count = 0;
      let last = 0;
      for await (const row of streamXlsxRows(file, { maxRows: ROWS })) {
        count++;
        last = row[0] as number;
        if (count % 50_000 === 0) {
          const cur = heapUsed();
          if (cur > peak) peak = cur;
        }
      }
      expect(count).toBe(ROWS);
      expect(last).toBe(ROWS);

      const growth = Math.max(0, peak - baseline);
      console.log(
        `   ↳ rows=${count}, file=${(file.size / MIB).toFixed(1)} MiB compressed, ` +
          `sheet uncompressed=${(uncompressedSheetBytes / MIB).toFixed(1)} MiB, ` +
          `peak heap growth=${(growth / MIB).toFixed(1)} MiB ` +
          `(gc=${HAS_GC ? 'forced' : 'not available — measurement includes garbage'})`,
      );

      if (HAS_GC) {
        // With explicit GC available, peak growth must stay below the size
        // of the uncompressed sheet — a non-streaming reader cannot.
        expect(growth).toBeLessThan(uncompressedSheetBytes);
      }
    },
    { timeout: 120_000 },
  );

  it(
    'reads only the rows requested when maxRows is small (no full decompression)',
    async () => {
      const ROWS = 200_000;
      let body = '<sheetData>';
      for (let i = 1; i <= ROWS; i++) {
        body += `<row r="${i}"><c r="A${i}"><v>${i}</v></c></row>`;
      }
      body += '</sheetData>';

      const xlsx = await buildXlsx({
        deflate: true,
        sheets: [{ name: 'big', sheetData: body }],
      });
      const file = asFile(xlsx);

      const start = performance.now();
      let count = 0;
      for await (const row of streamXlsxRows(file, { maxRows: 50 })) {
        count++;
        void row;
      }
      const elapsed = performance.now() - start;
      expect(count).toBe(50);
      // 50 rows from a 200k-row sheet should be fast — far less than the
      // time it would take to decompress the whole sheet. We use a generous
      // upper bound to stay stable across machines.
      expect(elapsed).toBeLessThan(1000);
      console.log(`   ↳ first 50 of ${ROWS} rows in ${elapsed.toFixed(0)} ms`);
    },
    { timeout: 60_000 },
  );
});
