# xlsx-stream-rows

**Streaming, chunked, low-memory spreadsheet reader for the browser.** Reads XLSX and CSV files row-by-row without loading the entire file into memory — a 1 GB workbook reads in roughly the same memory envelope as a 1 MB one. XLS is also supported via a unified interface, though it requires a full file load (see [Format support](#format-support)).

Zero dependencies for XLSX and CSV. TypeScript. ESM + CJS. Works in browsers and Web Workers.

```sh
npm install xlsx-stream-rows
```

---

## Why

Most browser-side spreadsheet libraries materialise the entire file before yielding a single row. A 300 MB XLSX usually peaks at 1–2 GB of JS heap and crashes the tab. `xlsx-stream-rows` treats `File` as a handle, not a byte array: it reads only the ZIP Central Directory at the end of the archive, then pipes the target sheet through `DecompressionStream` into an incremental SAX parser, **emitting rows lazily, on demand**.

If you've been searching for terms like _streaming xlsx parser_, _chunked spreadsheet reader_, _lazy xlsx_, _partial xlsx parsing_, _incremental xlsx reader_, _async iterator over xlsx rows_, _row-by-row xlsx_, _on-demand spreadsheet loading_, _memory-bounded xlsx_, or _read large xlsx in browser without OOM_ — that's what this library does.

---

## Features

- **True streaming.** `AsyncIterable<Row>` — pull rows on demand, stop whenever. No batch parse, no buffered intermediate result.
- **Bounded memory.** Peak heap is proportional to the data you consume, not the file size. Measured on a 1,000,000-row XLSX (51 MiB uncompressed sheet, 7.2 MiB on disk): **1.5 MiB peak heap growth**, ~34× smaller than the sheet (`tests/memory.smoke.test.ts`).
- **Three stop mechanisms, all equivalent.** `maxRows`, `break` out of `for await`, or `AbortSignal`. Pipeline tears down promptly: no further bytes fetched, decompressed, or parsed.
- **Format coverage.** XLSX / XLSM / XLTX / XLTM (streaming), CSV and TSV (streaming, custom separator supported), XLS (delegated to optional `xlsx` peer dep, bounded by `xlsMaxBytes`).
- **Auto-detect.** ZIP / OLE2 magic-byte sniff with filename extension as fallback.
- **Standards-grounded.** Every byte offset and XML path traces to [PKWARE APPNOTE.TXT](https://pkware.cachefly.net/webdocs/casestudies/APPNOTE.TXT) (ZIP) or [ECMA-376](https://www.ecma-international.org/publications-and-standards/standards/ecma-376/) (OOXML), not reverse-engineered from other libraries.
- **OPC-correct.** Resolves package parts via `_rels/.rels` indirection rather than hardcoded paths, so files from LibreOffice, Google Sheets, or custom generators work too.
- **Strict + transitional schemas.** Both SpreadsheetML namespaces accepted. All `ST_CellType` values handled (`n`, `s`, `str`, `inlineStr`, `b`, `e`, `d`).
- **Worker-safe.** No DOM APIs. Runs inside a Web Worker.
- **Zero dependencies for XLSX/CSV.** Web Platform APIs only (`File`, `Blob`, `DecompressionStream`, `TextDecoderStream`).

---

## Quick start

```ts
import { openWorkbook, streamRows, readRows } from 'xlsx-stream-rows';

// 1. List sheets without reading row data (≈ 100 KiB read regardless of file size).
const info = await openWorkbook(file);
console.log(info.sheetNames, info.format); // ['Sheet1', 'Data'], 'xlsx'

// 2. Stream rows one at a time. Memory stays bounded for files of any size.
for await (const row of streamRows(file, { maxRows: 100 })) {
  console.log(row); // (string | number | boolean | Date | null)[]
}

// 3. Convenience: collect into an array.
const rows = await readRows(file, { sheetName: 'Data', maxRows: 1000 });
```

### Bounded read — preview the first N rows of a huge file

```ts
// Reads only the bytes needed for the first 50 rows. The decompression
// stream is cancelled as soon as the limit is hit; nothing past that point
// is fetched from the file.
const preview = await readRows(file, { maxRows: 50 });
```

### Stop on a predicate

```ts
// `break` cancels the underlying pipeline — no further bytes are read.
for await (const row of streamRows(file)) {
  if (row[0] === 'END') break;
  process(row);
}
```

### Cancel from the outside

```ts
const ac = new AbortController();
cancelButton.onclick = () => ac.abort();

try {
  for await (const row of streamRows(file, { signal: ac.signal })) {
    insertIntoTable(row);
  }
} catch (e) {
  if (e === ac.signal.reason) console.log('user cancelled');
  else throw e;
}
```

`AbortSignal` cancels in-flight metadata fetches too — you can abort _before_ the first row is yielded (e.g. while `sharedStrings.xml` is still downloading).

### Progress reporting

The pull-based design makes progress trivial — count rows yourself, or report file-position progress via your UI:

```ts
let count = 0;
for await (const row of streamRows(file)) {
  count++;
  if (count % 1000 === 0) updateProgress(count);
  await processRow(row);
}
```

For byte-level progress you can wrap the input file with a `Proxy` over `slice()` that reports cumulative bytes; ask in an issue and we'll add an example.

### Run inside a Web Worker

The library has no DOM dependencies, so move heavy spreadsheet imports off the main thread:

```ts
// worker.ts
import { streamRows } from 'xlsx-stream-rows';

self.onmessage = async (e: MessageEvent<File>) => {
  const file = e.data;
  for await (const row of streamRows(file, { maxRows: 10_000 })) {
    self.postMessage({ type: 'row', row });
  }
  self.postMessage({ type: 'done' });
};
```

```ts
// main.ts
const worker = new Worker(new URL('./worker.ts', import.meta.url), { type: 'module' });
worker.onmessage = (e) => { /* … */ };
worker.postMessage(file); // File is structured-cloneable
```

### Pick a sheet by name

```ts
const info = await openWorkbook(file);
console.log(info.sheetNames); // ['Summary', 'Q1', 'Q2', 'Q3', 'Q4']

for await (const row of streamRows(file, { sheetName: 'Q3' })) {
  // …
}
```

### Date detection (XLSX)

XLSX stores dates as numbers styled with a date format. By default, numeric cells whose style references a date `numFmtId` (built-in 14–22, 27–36, 45–47, 50–58, plus custom `<numFmt>` containing date tokens) are returned as `Date`. Disable this if you want raw serial numbers:

```ts
const rows = await readRows(file, { parseDates: false }); // 44197 instead of new Date(2021, 0, 1)
```

### CSV with non-UTF encoding

UTF-8, UTF-16 LE, and UTF-16 BE BOMs are auto-detected. For other encodings (e.g. Windows-1251 / cp1251), pass `csvEncoding`:

```ts
const rows = await readRows(file, { csvEncoding: 'windows-1251' });
```

---

## Format support

| Extension(s) | Format | Streaming | Memory peak | Dependency |
|---|---|---|---|---|
| `.xlsx`, `.xlsm`, `.xltx`, `.xltm` | OOXML (Excel 2007+) | ✅ true streaming | ≈ sharedStrings size + few MiB | none — Web APIs only |
| `.csv` | CSV (RFC 4180, comma) | ✅ true streaming | ≈ one row + decoder window | none — Web APIs only |
| `.tsv` | TSV (tab-separated) | ✅ true streaming | ≈ one row + decoder window | none — Web APIs only |
| `.xls` | Legacy XLS (BIFF/OLE2) | ❌ full file load | ≤ `xlsMaxBytes` (default 50 MiB) | optional `xlsx` peer dep |

**Format detection** uses ZIP / OLE2 magic bytes first, filename extension as fallback. A file without an extension is treated as CSV. A file with an unrecognised extension throws `FormatNotSupportedError`.

**OOXML variants** (`.xlsm`, `.xltx`, `.xltm`) share the same ZIP + OPC structure as `.xlsx` — the streaming parser handles them identically.

**Encoding** (CSV/TSV): UTF-8 by default. UTF-8 BOM, UTF-16 LE BOM, and UTF-16 BE BOM are auto-detected. Any other encoding (e.g. `'windows-1251'`) can be passed via the `csvEncoding` option — any label recognised by the [WHATWG Encoding Standard](https://encoding.spec.whatwg.org/#names-and-labels) is accepted.

**Custom separator** (CSV/TSV): pass `separator` to override the default (`,` for `.csv`, `\t` for `.tsv`):
```ts
// semicolon-separated
const rows = await readRows(file, { separator: ';' });
```

**XLS is a legacy binary format** (Microsoft BIFF inside OLE2) that has no streaming primitive. The library loads the entire file into memory and delegates parsing to the optional [`xlsx`](https://sheetjs.com/) peer dependency. XLS support exists to give your code a single unified entry point — `openWorkbook` / `streamRows` — regardless of what file the user hands you. If you know you'll never receive `.xls` files, you don't need the peer dep at all.

Install `xlsx` only if you need XLS support (SheetJS no longer publishes to the npm registry; install from their CDN):

```sh
npm install https://cdn.sheetjs.com/xlsx-0.20.3/xlsx-0.20.3.tgz
```

---

## API

```ts
type CellValue = string | number | boolean | Date | null;
type Row = CellValue[];

interface WorkbookInfo {
  filename: string;
  sheetNames: string[];
  format: 'xlsx' | 'xls' | 'csv';
}

interface ReadOptions {
  /** Stop after yielding this many rows. */
  maxRows?: number;
  /** Sheet to read. Default: first sheet. Ignored for CSV. */
  sheetName?: string;
  /** XLSX/XLS only: convert numeric date-styled cells to `Date`. Default true. */
  parseDates?: boolean;
  /** XLSX only: cap on `xl/sharedStrings.xml` uncompressed size. Default 64 MiB. */
  sharedStringsMaxBytes?: number;
  /** XLS only: cap on the file size loaded into memory. Default 50 MiB. */
  xlsMaxBytes?: number;
  /** CSV/TSV only: text encoding. UTF-8/16 BOMs are auto-detected. Default 'utf-8'. */
  csvEncoding?: string;
  /** CSV/TSV only: field separator. Default ',' for .csv, '\t' for .tsv. */
  separator?: string;
  /** Cancel the read at any point — including before the first row is yielded. */
  signal?: AbortSignal;
}

function openWorkbook(file: File, options?: OpenWorkbookOptions): Promise<WorkbookInfo>;
function streamRows(file: File, options?: ReadOptions): AsyncIterable<Row>;
function readRows(file: File, options?: ReadOptions): Promise<Row[]>;
```

Per-format adapters (`openXlsxWorkbook`, `streamXlsxRows`, `openCsvWorkbook`, `streamCsvRows`, `openXlsWorkbook`, `streamXlsRows`) and lower-level building blocks (`readZipEntries`, `createRowParser`, `createCsvParser`, `resolvePackagePaths`, …) are also exported for power users.

### Errors

All errors inherit from `XlsxStreamError` so callers can catch the family with one `instanceof` check.

| Error | When |
|-------|------|
| `NotAZipError` | EOCD signature not found within 65,557 bytes of EOF |
| `Zip64NotSupportedError` | File uses ZIP64 extensions (>4 GiB / >65,535 entries) |
| `InvalidLocalHeaderError` | LFH signature mismatch — file corrupted |
| `UnsupportedCompressionError` | ZIP method is not 0 (stored) or 8 (deflate) |
| `InvalidOpcPackageError` | Missing `_rels/.rels` or `officeDocument` relationship |
| `SharedStringsTooLargeError` | `sharedStrings.xml` exceeds `sharedStringsMaxBytes` |
| `SheetNotFoundError` | Requested `sheetName` not in workbook |
| `EntryTooLargeError` | Bounded entry read exceeded its uncompressed cap |
| `XlsFileTooLargeError` | XLS file exceeds `xlsMaxBytes` |
| `XlsxPackageMissingError` | XLS read attempted without the `xlsx` peer dep installed |
| `FormatNotSupportedError` | File extension is set but unrecognised and magic bytes don't match any known format |

Pipeline cancellation (`maxRows` hit, iterator `return()`, `AbortSignal`) is **not** an error — the iterator simply ends, or rejects with `signal.reason`.

---

## How it works

XLSX files are ZIP archives of XML parts. The Central Directory at the end of the archive lists every entry's offset and size, so:

1. Fetch the trailing 64 KiB of the file → locate the EOCD record → read the Central Directory.
2. Resolve the workbook part via OPC relationships (`_rels/.rels` → `xl/_rels/workbook.xml.rels`).
3. Open a `ReadableStream` over the target sheet's compressed bytes (`Blob.slice().stream()`), pipe through `DecompressionStream('deflate-raw')` and `TextDecoderStream`, feed the result to a hand-rolled SAX state machine that emits rows as `</row>` closes.
4. Cancelling the iterator (`maxRows`, `break`, `AbortSignal`) cancels the reader, which propagates up the pipeline — no more bytes are fetched.

For CSV/TSV: same shape, simpler — `file.stream() → TextDecoderStream → RFC-4180 parser` with a configurable field separator (`,` for CSV, `\t` for TSV, or any custom character).

For XLS: no streaming primitive exists in the BIFF/OLE2 format, so we delegate to the `xlsx` package and bound the file size to keep memory predictable.

---

## Limitations

- **No ZIP64.** Files with the Central Directory or any individual entry over 4 GiB, or with more than 65,535 entries, are rejected with `Zip64NotSupportedError`. In practice you can comfortably read 5–10 GB of sheet data; multi-GB compressed XLSX with monster sharedStrings tables are the edge case to watch.
- **No formula evaluation.** Formula cells return the cached `<v>` written by the producer; if absent, `null`. We do not re-evaluate.
- **No merged-cell expansion.** Cells appear exactly where the XML places them.
- **No password-protected files.**

---

## Standards

- ZIP container — [PKWARE APPNOTE.TXT 6.3.10](https://pkware.cachefly.net/webdocs/casestudies/APPNOTE.TXT)
- OOXML — [ECMA-376 5th Edition Part 1 (2016)](https://www.ecma-international.org/publications-and-standards/standards/ecma-376/) (SpreadsheetML) and Part 2 (2021) (OPC)
- CSV — [RFC 4180](https://datatracker.ietf.org/doc/html/rfc4180)

---

## Try it in your browser

The repo includes a self-contained playground:

```sh
git clone https://github.com/gudoshnikovn/xlsx-stream-rows
cd xlsx-stream-rows
npm install && npm run build
npx serve .   # or any static server
# open http://localhost:3000/examples/playground.html
```

Drop a real spreadsheet in, set `maxRows`, watch it stream.

## Development

```sh
npm install
npm test            # Node test suite (260+ tests, including fuzz)
npm run test:browser  # same suite in real Chromium / Firefox / WebKit via Playwright
npm run test:memory   # 1M-row memory smoke (set --pool=forks for forced GC)
npm run build         # tsup → dist/ (ESM + CJS + d.ts)
npm run typecheck
```

Requires Node 20+ for `File`, `Blob.stream()`, and `DecompressionStream` globals. CI runs on Node 20 / 22 plus all three browsers.

---

## License

MIT
