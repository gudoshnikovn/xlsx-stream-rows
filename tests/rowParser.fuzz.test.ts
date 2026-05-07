/**
 * Chunk-boundary fuzz: split the same XML at every byte position and assert
 * the parser produces identical output. This is the load-bearing correctness
 * test for the streaming SAX state machine — if the parser holds state
 * correctly across `push()` calls, partitioning cannot change the output.
 */

import { describe, expect, it } from 'vitest';

import { createRowParser, type RowParserContext } from '../src/xlsx/rowParser.js';
import type { Row } from '../src/types.js';

const ctx = (overrides: Partial<RowParserContext> = {}): RowParserContext => ({
  sharedStrings: [],
  dateFormatStyleIds: new Set<number>(),
  parseDates: true,
  ...overrides,
});

function parseInOneShot(xml: string, c: RowParserContext): Row[] {
  const p = createRowParser(c);
  const rows = p.push(xml);
  rows.push(...p.end());
  return rows;
}

function parseAtSplit(xml: string, splitAt: number, c: RowParserContext): Row[] {
  const p = createRowParser(c);
  const rows: Row[] = [];
  rows.push(...p.push(xml.slice(0, splitAt)));
  rows.push(...p.push(xml.slice(splitAt)));
  rows.push(...p.end());
  return rows;
}

function parseAtAllSplits(xml: string, c: RowParserContext): void {
  const expected = parseInOneShot(xml, c);
  for (let i = 0; i <= xml.length; i++) {
    const got = parseAtSplit(xml, i, c);
    expect(got, `split at byte ${i}`).toEqual(expected);
  }
}

function parseAtTriSplit(
  xml: string,
  a: number,
  b: number,
  c: RowParserContext,
): Row[] {
  const p = createRowParser(c);
  const rows: Row[] = [];
  rows.push(...p.push(xml.slice(0, a)));
  rows.push(...p.push(xml.slice(a, b)));
  rows.push(...p.push(xml.slice(b)));
  rows.push(...p.end());
  return rows;
}

describe('rowParser chunk-boundary fuzz', () => {
  it('survives every binary split of a representative XLSX sheet fragment', () => {
    const xml =
      `<?xml version="1.0" encoding="UTF-8"?>` +
      `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
      `<sheetData>` +
      `<row r="1">` +
      `<c r="A1" t="s"><v>0</v></c>` +
      `<c r="B1"><v>3.14</v></c>` +
      `<c r="C1" t="b"><v>1</v></c>` +
      `</row>` +
      `<row r="2">` +
      `<c r="A2" t="inlineStr"><is><t>Tom &amp; Jerry</t></is></c>` +
      `<c r="C2"><f>SUM(B1)</f><v>3.14</v></c>` +
      `</row>` +
      `</sheetData>` +
      `</worksheet>`;
    parseAtAllSplits(xml, ctx({ sharedStrings: ['Alice'] }));
  });

  it('survives every binary split of namespace-prefixed markup', () => {
    const xml =
      `<x:worksheet xmlns:x="...">` +
      `<x:sheetData>` +
      `<x:row r="1"><x:c r="A1"><x:v>1</x:v></x:c></x:row>` +
      `<x:row r="2"><x:c r="A2"><x:v>2</x:v></x:c></x:row>` +
      `</x:sheetData>` +
      `</x:worksheet>`;
    parseAtAllSplits(xml, ctx());
  });

  it('survives splits inside character references', () => {
    const xml =
      `<sheetData><row r="1">` +
      `<c r="A1" t="inlineStr"><is><t>&amp;&lt;&gt;&quot;&apos;&#39;&#x27;</t></is></c>` +
      `</row></sheetData>`;
    parseAtAllSplits(xml, ctx());
  });

  it('survives splits inside CDATA sections', () => {
    const xml =
      `<sheetData><row r="1">` +
      `<c r="A1" t="inlineStr"><is><t><![CDATA[hello & world]]></t></is></c>` +
      `</row></sheetData>`;
    parseAtAllSplits(xml, ctx());
  });

  it('survives 3-way splits across larger inputs', () => {
    let body = '<sheetData>';
    for (let i = 1; i <= 20; i++) {
      body += `<row r="${i}"><c r="A${i}"><v>${i}</v></c></row>`;
    }
    body += '</sheetData>';

    const expected = parseInOneShot(body, ctx());
    // Sample the cubic space rather than enumerate every (a, b) pair to keep
    // the test fast — 20-byte stride is enough to expose state-machine bugs.
    for (let a = 0; a <= body.length; a += 20) {
      for (let b = a; b <= body.length; b += 20) {
        const got = parseAtTriSplit(body, a, b, ctx());
        expect(got, `splits at ${a}, ${b}`).toEqual(expected);
      }
    }
  });
});
