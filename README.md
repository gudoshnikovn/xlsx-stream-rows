# xlsx-stream

Zero-dependency streaming XLSX / CSV / XLS reader for the browser. Peak memory is proportional to the data returned, not to the file size — a 1 GB workbook reads in the same memory envelope as a 1 MB one.

> **Status:** early development. The ZIP layer (Phase 1) is implemented; XLSX parsing, CSV streaming, and the public `streamRows` / `openWorkbook` / `readRows` API are in progress.

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
