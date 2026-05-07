# xlsx-stream

Zero-dependency streaming XLSX / CSV / XLS reader for the browser. Peak memory is proportional to the data returned, not to the file size — a 1 GB workbook reads in the same memory envelope as a 1 MB one.

> **Status:** early development. XLSX streaming is implemented end-to-end (`openXlsxWorkbook`, `streamXlsxRows`); CSV streaming, XLS delegation, and the unified `streamRows` / `openWorkbook` / `readRows` public API across all three formats are next.

## Quick start

```ts
import { openXlsxWorkbook, streamXlsxRows } from 'xlsx-stream';

// List sheets without reading row data (≈ 100 KiB read regardless of file size)
const info = await openXlsxWorkbook(file);
console.log(info.sheetNames);

// Stream rows on demand. Memory peak is bounded by sharedStrings size,
// not the file or sheet length.
for await (const row of streamXlsxRows(file, { maxRows: 100 })) {
  console.log(row); // (string | number | boolean | Date | null)[]
}

// Cancel from the outside (user clicks "Cancel"):
const ac = new AbortController();
for await (const row of streamXlsxRows(file, { signal: ac.signal })) {
  if (somethingHappened) ac.abort();
}
```

## Why

Browser-based spreadsheet readers historically materialise the entire file before yielding any data. A 300 MB XLSX often peaks at 1–2 GB of JS heap and crashes the tab. `xlsx-stream` treats `File` as a handle: it reads only the ZIP Central Directory at the end of the archive, then streams the target sheet through `DecompressionStream` into an incremental XML parser, yielding rows on demand.

## Standards

Every byte-level decision traces to a published specification:

- ZIP container — [PKWARE APPNOTE.TXT 6.3.10](https://pkware.cachefly.net/webdocs/casestudies/APPNOTE.TXT)
- XLSX format — [ECMA-376 5th Edition](https://www.ecma-international.org/publications-and-standards/standards/ecma-376/) (Part 1: SpreadsheetML, Part 2: OPC packaging)

## Development

```sh
npm install
npm test         # vitest
npm run build    # tsup → dist/ (ESM + CJS + d.ts)
npm run typecheck
```

Requires Node 20+ for `File`, `Blob.stream()`, and `DecompressionStream` globals.

## License

MIT
