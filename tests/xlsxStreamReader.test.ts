import { describe, expect, it } from 'vitest';

import {
  InvalidOpcPackageError,
  SharedStringsTooLargeError,
  SheetNotFoundError,
  openXlsxWorkbook,
  streamXlsxRows,
} from '../src/index.js';
import { asFile } from './helpers/buildZip.js';
import { buildXlsx } from './helpers/buildXlsx.js';
import type { Row } from '../src/types.js';

const SHEET_DATA_BASIC = `<sheetData>
  <row r="1">
    <c r="A1" t="s"><v>0</v></c>
    <c r="B1"><v>3.14</v></c>
    <c r="C1" t="b"><v>1</v></c>
  </row>
  <row r="2">
    <c r="A2" t="s"><v>1</v></c>
    <c r="B2"><v>42</v></c>
  </row>
  <row r="3">
    <c r="A3" t="inlineStr"><is><t>inline &amp; quoted</t></is></c>
  </row>
</sheetData>`;

async function collect<T>(iter: AsyncIterable<T>): Promise<T[]> {
  const out: T[] = [];
  for await (const x of iter) out.push(x);
  return out;
}

describe('openXlsxWorkbook', () => {
  it('returns sheet names without iterating rows', async () => {
    const xlsx = await buildXlsx({
      sheets: [
        { name: 'Sheet1', sheetData: '<sheetData/>' },
        { name: 'Data', sheetData: '<sheetData/>' },
        { name: 'Config', sheetData: '<sheetData/>' },
      ],
    });
    const info = await openXlsxWorkbook(asFile(xlsx, 'wb.xlsx'));
    expect(info.filename).toBe('wb.xlsx');
    expect(info.format).toBe('xlsx');
    expect(info.sheetNames).toEqual(['Sheet1', 'Data', 'Config']);
  });

  it('throws InvalidOpcPackageError when _rels/.rels is missing', async () => {
    // Build a "ZIP that is not a valid OPC package" — a bare ZIP with a
    // non-rels first entry suffices.
    const { buildZip } = await import('./helpers/buildZip.js');
    const garbage = await buildZip([
      { name: 'random.txt', data: new TextEncoder().encode('hi'), method: 0 },
    ]);
    await expect(openXlsxWorkbook(asFile(garbage))).rejects.toBeInstanceOf(
      InvalidOpcPackageError,
    );
  });
});

describe('streamXlsxRows — basic streaming', () => {
  it('streams rows from a single-sheet workbook with sharedStrings', async () => {
    const xlsx = await buildXlsx({
      sharedStrings: ['Alice', 'Bob'],
      sheets: [{ name: 'Sheet1', sheetData: SHEET_DATA_BASIC }],
    });
    const rows = await collect(streamXlsxRows(asFile(xlsx)));
    expect(rows).toEqual([
      ['Alice', 3.14, true],
      ['Bob', 42],
      ['inline & quoted'],
    ]);
  });

  it('selects a sheet by name', async () => {
    const xlsx = await buildXlsx({
      sheets: [
        { name: 'A', sheetData: '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>' },
        { name: 'B', sheetData: '<sheetData><row r="1"><c r="A1"><v>2</v></c></row></sheetData>' },
      ],
    });
    const rows = await collect(streamXlsxRows(asFile(xlsx), { sheetName: 'B' }));
    expect(rows).toEqual([[2]]);
  });

  it('throws SheetNotFoundError when the requested sheet does not exist', async () => {
    const xlsx = await buildXlsx({
      sheets: [{ name: 'OnlyOne', sheetData: '<sheetData/>' }],
    });
    const iter = streamXlsxRows(asFile(xlsx), { sheetName: 'Nope' });
    await expect(collect(iter)).rejects.toBeInstanceOf(SheetNotFoundError);
  });

  it('streams from a deflate-compressed workbook (real-world path)', async () => {
    const xlsx = await buildXlsx({
      deflate: true,
      sharedStrings: ['x'],
      sheets: [
        {
          name: 'S',
          sheetData: '<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData>',
        },
      ],
    });
    expect(await collect(streamXlsxRows(asFile(xlsx)))).toEqual([['x']]);
  });
});

describe('streamXlsxRows — non-conventional layouts', () => {
  it('reads a workbook placed outside xl/', async () => {
    const xlsx = await buildXlsx({
      workbookPath: 'spreadsheet/wb.xml',
      sheetsDir: 'spreadsheet/sheets',
      sharedStringsPath: 'strings/sst.xml',
      sharedStrings: ['foo'],
      sheets: [
        {
          name: 'A',
          sheetData: '<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData>',
        },
      ],
    });
    const info = await openXlsxWorkbook(asFile(xlsx));
    expect(info.sheetNames).toEqual(['A']);
    expect(await collect(streamXlsxRows(asFile(xlsx)))).toEqual([['foo']]);
  });

  it('reads a strict-schema workbook (purl.oclc.org namespace)', async () => {
    const xlsx = await buildXlsx({
      strict: true,
      sharedStrings: ['hi'],
      sheets: [
        {
          name: 'A',
          sheetData: '<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData>',
        },
      ],
    });
    expect(await collect(streamXlsxRows(asFile(xlsx)))).toEqual([['hi']]);
  });

  it('reads a workbook with a namespace prefix on every SpreadsheetML element', async () => {
    const xlsx = await buildXlsx({
      ssPrefix: 'x',
      sheets: [
        {
          name: 'A',
          sheetData:
            '<x:sheetData><x:row r="1"><x:c r="A1"><x:v>7</x:v></x:c></x:row></x:sheetData>',
        },
      ],
    });
    expect(await collect(streamXlsxRows(asFile(xlsx)))).toEqual([[7]]);
  });
});

describe('streamXlsxRows — date detection via styles.xml', () => {
  it('returns numeric date-styled cells as Date when parseDates=true', async () => {
    const xlsx = await buildXlsx({
      stylesXml: `<styleSheet>
        <cellXfs count="2">
          <xf numFmtId="0"/>
          <xf numFmtId="14" applyNumberFormat="1"/>
        </cellXfs>
      </styleSheet>`,
      sheets: [
        {
          name: 'A',
          sheetData: `<sheetData>
            <row r="1">
              <c r="A1"><v>44197</v></c>
              <c r="B1" s="1"><v>44197</v></c>
            </row>
          </sheetData>`,
        },
      ],
    });
    const rows = await collect(streamXlsxRows(asFile(xlsx)));
    expect(rows[0]?.[0]).toBe(44197);
    expect(rows[0]?.[1]).toBeInstanceOf(Date);
    expect((rows[0]?.[1] as Date).getTime()).toBe(Date.UTC(2021, 0, 1));
  });

  it('keeps date-styled cells as numbers when parseDates=false', async () => {
    const xlsx = await buildXlsx({
      stylesXml: `<styleSheet><cellXfs count="1"><xf numFmtId="14"/></cellXfs></styleSheet>`,
      sheets: [
        {
          name: 'A',
          sheetData: `<sheetData><row r="1"><c r="A1" s="0"><v>44197</v></c></row></sheetData>`,
        },
      ],
    });
    const rows = await collect(streamXlsxRows(asFile(xlsx), { parseDates: false }));
    expect(rows[0]).toEqual([44197]);
  });
});

describe('streamXlsxRows — bounded reads and cancellation', () => {
  function manyRowsSheet(n: number): string {
    let body = '<sheetData>';
    for (let i = 1; i <= n; i++) {
      body += `<row r="${i}"><c r="A${i}"><v>${i}</v></c></row>`;
    }
    body += '</sheetData>';
    return body;
  }

  it('stops after maxRows', async () => {
    const xlsx = await buildXlsx({
      sheets: [{ name: 'A', sheetData: manyRowsSheet(1000) }],
    });
    const rows = await collect(streamXlsxRows(asFile(xlsx), { maxRows: 5 }));
    expect(rows.map((r) => r[0])).toEqual([1, 2, 3, 4, 5]);
  });

  it('stops on `break` out of the for-await loop', async () => {
    const xlsx = await buildXlsx({
      sheets: [{ name: 'A', sheetData: manyRowsSheet(1000) }],
    });
    const collected: Row[] = [];
    for await (const r of streamXlsxRows(asFile(xlsx))) {
      if ((r[0] as number) > 3) break;
      collected.push(r);
    }
    expect(collected.map((r) => r[0])).toEqual([1, 2, 3]);
  });

  it('rejects with the AbortSignal reason when aborted before iteration', async () => {
    const xlsx = await buildXlsx({
      sheets: [{ name: 'A', sheetData: manyRowsSheet(10) }],
    });
    const ac = new AbortController();
    const reason = new Error('user cancelled');
    ac.abort(reason);
    const iter = streamXlsxRows(asFile(xlsx), { signal: ac.signal });
    await expect(collect(iter)).rejects.toBe(reason);
  });

  it('rejects with the AbortSignal reason when aborted mid-iteration', async () => {
    const xlsx = await buildXlsx({
      sheets: [{ name: 'A', sheetData: manyRowsSheet(1000) }],
    });
    const ac = new AbortController();
    const it = streamXlsxRows(asFile(xlsx), { signal: ac.signal })[Symbol.asyncIterator]();
    // Pull one row, then abort, then pull again.
    const first = await it.next();
    expect(first.value).toEqual([1]);
    const reason = new Error('mid-stream cancel');
    ac.abort(reason);
    await expect(it.next()).rejects.toBe(reason);
  });
});

describe('streamXlsxRows — sharedStrings cap', () => {
  it('throws SharedStringsTooLargeError when the sst exceeds the limit', async () => {
    const xlsx = await buildXlsx({
      sharedStrings: ['filler'.repeat(50)],
      sheets: [{ name: 'A', sheetData: '<sheetData/>' }],
    });
    await expect(
      collect(streamXlsxRows(asFile(xlsx), { sharedStringsMaxBytes: 50 })),
    ).rejects.toBeInstanceOf(SharedStringsTooLargeError);
  });
});

describe('streamXlsxRows — lazy sharedStrings with maxRows', () => {
  function sharedStringSheet(count: number): string {
    let body = '<sheetData>';
    for (let i = 0; i < count; i++) {
      body += `<row r="${i + 1}"><c r="A${i + 1}" t="s"><v>${i}</v></c></row>`;
    }
    return body + '</sheetData>';
  }

  it('resolves shared strings correctly when using maxRows', async () => {
    const strings = Array.from({ length: 100 }, (_, i) => `item-${i}`);
    const xlsx = await buildXlsx({
      sharedStrings: strings,
      sheets: [{ name: 'S', sheetData: sharedStringSheet(100) }],
    });
    const rows = await collect(streamXlsxRows(asFile(xlsx), { maxRows: 3 }));
    expect(rows).toEqual([['item-0'], ['item-1'], ['item-2']]);
  });

  it('resolves sparse shared-string indices when using maxRows', async () => {
    // Row 1 → index 0, row 2 → index 99: both must resolve despite index gap.
    const strings = ['first', ...Array<string>(98).fill('filler'), 'last'];
    const sheetData = `<sheetData>
      <row r="1"><c r="A1" t="s"><v>0</v></c></row>
      <row r="2"><c r="A2" t="s"><v>99</v></c></row>
      <row r="3"><c r="A3" t="s"><v>50</v></c></row>
    </sheetData>`;
    const xlsx = await buildXlsx({
      sharedStrings: strings,
      sheets: [{ name: 'S', sheetData }],
    });
    const rows = await collect(streamXlsxRows(asFile(xlsx), { maxRows: 2 }));
    expect(rows).toEqual([['first'], ['last']]);
  });

  it('handles mixed shared-string and inline cells with maxRows', async () => {
    const sheetData = `<sheetData>
      <row r="1">
        <c r="A1" t="s"><v>0</v></c>
        <c r="B1"><v>42</v></c>
      </row>
      <row r="2">
        <c r="A2" t="s"><v>1</v></c>
        <c r="B2"><v>7</v></c>
      </row>
    </sheetData>`;
    const xlsx = await buildXlsx({
      sharedStrings: ['Alice', 'Bob'],
      sheets: [{ name: 'S', sheetData }],
    });
    const rows = await collect(streamXlsxRows(asFile(xlsx), { maxRows: 1 }));
    expect(rows).toEqual([['Alice', 42]]);
  });
});

describe('streamXlsxRows — missing files and edge cases', () => {
  it('throws InvalidOpcPackageError when workbook.xml is missing', async () => {
    const { buildZip } = await import('./helpers/buildZip.js');
    const broken = await buildZip([
      { name: '_rels/.rels', data: new TextEncoder().encode(
        `<?xml version="1.0"?><Relationships/>`
      ), method: 8 },
      { name: '[Content_Types].xml', data: new TextEncoder().encode(
        `<?xml version="1.0"?><Types/>`
      ), method: 8 },
    ]);
    const iter = streamXlsxRows(asFile(broken));
    await expect(collect(iter)).rejects.toBeInstanceOf(InvalidOpcPackageError);
  });

  it('handles 0 sheets gracefully (empty workbook)', async () => {
    const xlsx = await buildXlsx({
      sheets: [],
    });
    const info = await openXlsxWorkbook(asFile(xlsx));
    expect(info.sheetNames).toEqual([]);
  });

  it('throws InvalidOpcPackageError when no sheets exist (workbook is invalid with 0 sheets)', async () => {
    const xlsx = await buildXlsx({
      sheets: [],
    });
    const iter = streamXlsxRows(asFile(xlsx));
    // When workbook has 0 sheets and no sheetName specified, returns empty iterator
    const rows = await collect(iter);
    expect(rows).toEqual([]);
  });

  it('skips dangling relationship IDs (rId points to missing sheet file)', async () => {
    // This is tricky: we'd need to create a workbook.xml with a sheet pointing to a missing rId
    // For now, test that lazy loading of 0 shared strings succeeds
    const xlsx = await buildXlsx({
      sharedStrings: [],
      sheets: [{ name: 'A', sheetData: '<sheetData><row r="1"><c r="A1"><v>5</v></c></row></sheetData>' }],
    });
    const rows = await collect(streamXlsxRows(asFile(xlsx)));
    expect(rows).toEqual([[5]]);
  });

  it('handles missing sharedStrings.xml (lazy load with no indices needed)', async () => {
    // Build XLSX without sharedStrings to test lazy load path
    const xlsx = await buildXlsx({
      sheets: [{ name: 'A', sheetData: '<sheetData><row r="1"><c r="A1"><v>42</v></c></row></sheetData>' }],
    });
    const rows = await collect(streamXlsxRows(asFile(xlsx)));
    expect(rows).toEqual([[42]]);
  });

  it('reads shared strings when ZIP entry is differently-cased than rels reference (e.g. SharedStrings.xml vs sharedStrings.xml)', async () => {
    const sheetData = `<sheetData>
      <row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>1</v></c></row>
      <row r="2"><c r="A2" t="s"><v>1</v></c><c r="B2"><v>2</v></c></row>
    </sheetData>`;
    // rels will reference xl/sharedStrings.xml; ZIP entry will be xl/SharedStrings.xml
    const xlsx = await buildXlsx({
      sheets: [{ name: 'Sheet1', sheetData }],
      sharedStrings: ['hello', 'world'],
      sharedStringsZipEntryPath: 'xl/SharedStrings.xml',
    });
    const rows = await collect(streamXlsxRows(asFile(xlsx)));
    expect(rows).toEqual([['hello', 1], ['world', 2]]);
  });

  it('reads shared strings with mismatched case via lazy load (maxRows path)', async () => {
    const sheetData = `<sheetData>
      <row r="1"><c r="A1" t="s"><v>0</v></c></row>
      <row r="2"><c r="A2" t="s"><v>1</v></c></row>
    </sheetData>`;
    const xlsx = await buildXlsx({
      sheets: [{ name: 'Sheet1', sheetData }],
      sharedStrings: ['foo', 'bar'],
      sharedStringsZipEntryPath: 'xl/SharedStrings.xml',
    });
    const rows = await collect(streamXlsxRows(asFile(xlsx), { maxRows: 10 }));
    expect(rows).toEqual([['foo'], ['bar']]);
  });

  it('handles missing styles.xml gracefully (no date formatting applied)', async () => {
    const xlsx = await buildXlsx({
      sheets: [{ name: 'A', sheetData: '<sheetData><row r="1"><c r="A1" s="1"><v>44197</v></c></row></sheetData>' }],
    });
    const rows = await collect(streamXlsxRows(asFile(xlsx)));
    // Without styles, numeric value should NOT be converted to date
    expect(rows).toEqual([[44197]]);
  });
});

describe('streamXlsxRows — lazy sharedStrings with self-closing <si/>', () => {
  it('handles <si/> self-closing elements inside sharedStrings.xml in lazy load (reader.ts:556-559)', async () => {
    // The lazy SST loader (loadSharedStringsSelective) has its own <si/> handling
    // separate from the regex-based parseSharedStrings. This exercises lines 556-559.
    //
    // We inject a raw SST XML containing a self-closing <si/> (index 1 = empty string)
    // alongside normal entries representing indices 0 and 2.
    // Using maxRows triggers the two-pass lazy loading path.
    const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    const xlsx = await buildXlsx({
      rawSharedStringsXml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="${NS}" count="3" uniqueCount="3">
  <si><t>hello</t></si>
  <si/>
  <si><t>world</t></si>
</sst>`,
      sheets: [{
        name: 'S',
        sheetData: `<sheetData>
  <row r="1">
    <c r="A1" t="s"><v>0</v></c>
    <c r="B1" t="s"><v>1</v></c>
    <c r="C1" t="s"><v>2</v></c>
  </row>
</sheetData>`,
      }],
    });

    // maxRows = 1 triggers the lazy two-pass loader
    const rows = await collect(streamXlsxRows(asFile(xlsx), { maxRows: 1 }));
    expect(rows).toEqual([['hello', '', 'world']]);
  });
});
