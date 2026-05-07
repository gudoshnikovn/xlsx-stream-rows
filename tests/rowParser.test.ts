import { describe, expect, it } from 'vitest';

import { createRowParser, type RowParserContext } from '../src/xlsx/rowParser.js';
import type { Row } from '../src/types.js';

const ctx = (overrides: Partial<RowParserContext> = {}): RowParserContext => ({
  sharedStrings: [],
  dateFormatStyleIds: new Set<number>(),
  parseDates: true,
  ...overrides,
});

function parseAll(xml: string, c: RowParserContext = ctx()): Row[] {
  const p = createRowParser(c);
  const rows = p.push(xml);
  rows.push(...p.end());
  return rows;
}

describe('createRowParser — single-chunk parsing', () => {
  it('parses a numeric cell', () => {
    const xml = `<sheetData><row r="1"><c r="A1"><v>42</v></c></row></sheetData>`;
    expect(parseAll(xml)).toEqual([[42]]);
  });

  it('parses a shared-string cell using the provided string table', () => {
    const xml = `<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData>`;
    expect(parseAll(xml, ctx({ sharedStrings: ['Alice'] }))).toEqual([['Alice']]);
  });

  it('parses an inline string', () => {
    const xml = `<sheetData><row r="1">
      <c r="A1" t="inlineStr"><is><t>hello</t></is></c>
    </row></sheetData>`;
    expect(parseAll(xml)).toEqual([['hello']]);
  });

  it('concatenates multiple <t> runs inside <is>', () => {
    const xml = `<sheetData><row r="1">
      <c r="A1" t="inlineStr"><is><r><t>foo</t></r><r><t>bar</t></r></is></c>
    </row></sheetData>`;
    expect(parseAll(xml)).toEqual([['foobar']]);
  });

  it('parses a boolean cell', () => {
    const xml = `<sheetData>
      <row r="1"><c r="A1" t="b"><v>1</v></c></row>
      <row r="2"><c r="A2" t="b"><v>0</v></c></row>
    </sheetData>`;
    expect(parseAll(xml)).toEqual([[true], [false]]);
  });

  it('parses a formula cell using the cached <v>', () => {
    const xml = `<sheetData><row r="1">
      <c r="A1"><f>SUM(B1:C1)</f><v>5</v></c>
    </row></sheetData>`;
    expect(parseAll(xml)).toEqual([[5]]);
  });

  it('parses an error cell as a string', () => {
    const xml = `<sheetData><row r="1">
      <c r="A1" t="e"><v>#DIV/0!</v></c>
    </row></sheetData>`;
    expect(parseAll(xml)).toEqual([['#DIV/0!']]);
  });

  it('parses a strict-schema ISO 8601 date cell (t="d")', () => {
    const xml = `<sheetData><row r="1">
      <c r="A1" t="d"><v>2021-01-01T00:00:00Z</v></c>
    </row></sheetData>`;
    const rows = parseAll(xml);
    expect(rows[0]?.[0]).toBeInstanceOf(Date);
    expect((rows[0]?.[0] as Date).toISOString()).toBe('2021-01-01T00:00:00.000Z');
  });

  it('handles sparse cells, padding gaps with null', () => {
    const xml = `<sheetData><row r="1">
      <c r="A1"><v>1</v></c>
      <c r="C1"><v>3</v></c>
    </row></sheetData>`;
    expect(parseAll(xml)).toEqual([[1, null, 3]]);
  });

  it('returns null for an explicitly empty <c r="A1"/>', () => {
    const xml = `<sheetData><row r="1">
      <c r="A1"/><c r="B1"><v>2</v></c>
    </row></sheetData>`;
    expect(parseAll(xml)).toEqual([[null, 2]]);
  });

  it('uses positional fallback when r= is absent', () => {
    // Some non-Excel producers emit cells without the r attribute.
    const xml = `<sheetData><row>
      <c><v>10</v></c><c><v>20</v></c><c><v>30</v></c>
    </row></sheetData>`;
    expect(parseAll(xml)).toEqual([[10, 20, 30]]);
  });

  it('decodes column letters past Z (AA = 27th column)', () => {
    const xml = `<sheetData><row r="1">
      <c r="AA1"><v>27</v></c>
    </row></sheetData>`;
    const row = parseAll(xml)[0]!;
    expect(row.length).toBe(27);
    expect(row[26]).toBe(27);
  });

  it('decodes XML entities inside cell values', () => {
    const xml = `<sheetData><row r="1">
      <c r="A1" t="inlineStr"><is><t>Tom &amp; Jerry &lt;3</t></is></c>
    </row></sheetData>`;
    expect(parseAll(xml)).toEqual([['Tom & Jerry <3']]);
  });

  it('converts a numeric date-styled cell to a Date when parseDates=true', () => {
    const xml = `<sheetData><row r="1">
      <c r="A1" s="1"><v>44197</v></c>
    </row></sheetData>`;
    const rows = parseAll(xml, ctx({ dateFormatStyleIds: new Set([1]) }));
    expect(rows[0]?.[0]).toBeInstanceOf(Date);
    expect((rows[0]?.[0] as Date).getTime()).toBe(Date.UTC(2021, 0, 1));
  });

  it('keeps a numeric date-styled cell as number when parseDates=false', () => {
    const xml = `<sheetData><row r="1">
      <c r="A1" s="1"><v>44197</v></c>
    </row></sheetData>`;
    const rows = parseAll(
      xml,
      ctx({ dateFormatStyleIds: new Set([1]), parseDates: false }),
    );
    expect(rows[0]).toEqual([44197]);
  });

  it('tolerates namespace prefixes on every tag', () => {
    const xml = `<x:sheetData>
      <x:row r="1">
        <x:c r="A1" t="s"><x:v>0</x:v></x:c>
      </x:row>
    </x:sheetData>`;
    expect(parseAll(xml, ctx({ sharedStrings: ['ok'] }))).toEqual([['ok']]);
  });

  it('skips processing instructions, comments and DOCTYPE markers', () => {
    const xml =
      `<?xml version="1.0"?>` +
      `<!DOCTYPE foo>` +
      `<!-- top comment -->` +
      `<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>`;
    expect(parseAll(xml)).toEqual([[1]]);
  });

  it('handles CDATA sections inside text content', () => {
    const xml = `<sheetData><row r="1">
      <c r="A1" t="inlineStr"><is><t><![CDATA[<raw & text>]]></t></is></c>
    </row></sheetData>`;
    expect(parseAll(xml)).toEqual([['<raw & text>']]);
  });

  it('parses many rows in one shot', () => {
    let body = '<sheetData>';
    for (let i = 1; i <= 100; i++) {
      body += `<row r="${i}"><c r="A${i}"><v>${i}</v></c></row>`;
    }
    body += '</sheetData>';
    const rows = parseAll(body);
    expect(rows).toHaveLength(100);
    expect(rows[0]).toEqual([1]);
    expect(rows[99]).toEqual([100]);
  });
});
