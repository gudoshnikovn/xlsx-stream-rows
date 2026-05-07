# xlsx-stream

Zero-dependency streaming XLSX / CSV / XLS reader for the browser. Peak memory is proportional to the data returned, not to the file size — a 1 GB workbook reads in the same memory envelope as a 1 MB one.

> **Status:** beta. All three formats (XLSX, CSV, XLS) work through the unified `openWorkbook` / `streamRows` / `readRows` API. Format is auto-detected by magic bytes with extension fallback.

## Quick start

```ts
import { openWorkbook, streamRows, readRows } from 'xlsx-stream';

// List sheets without reading row data
const info = await openWorkbook(file);
console.log(info.sheetNames, info.format); // ['Sheet1', 'Data'], 'xlsx'

// Stream rows on demand — bounded memory regardless of file size (XLSX/CSV)
for await (const row of streamRows(file, { maxRows: 100 })) {
  console.log(row); // (string | number | boolean | Date | null)[]
}

// Convenience: collect into an array
const rows = await readRows(file, { sheetName: 'Data', maxRows: 1000 });

// Cancel from the outside (user clicks "Cancel")
const ac = new AbortController();
for await (const row of streamRows(file, { signal: ac.signal })) {
  if (userCancelled) ac.abort();
}
```

## Format support

| Format | Streaming | Memory peak | Dependency |
|--------|-----------|-------------|------------|
| XLSX / XLSM | yes | ≈ sharedStrings size + ~5 MiB | none |
| CSV | yes | ≈ one row + decoder window | none |
| XLS | no (delegated) | ≈ file size, capped by `xlsMaxBytes` | optional `xlsx` peer dep |

Install `xlsx` only if you need XLS support:

```sh
npm install xlsx
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
