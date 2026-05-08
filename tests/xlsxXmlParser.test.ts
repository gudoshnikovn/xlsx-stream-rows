import { describe, expect, it } from 'vitest';

import {
  extractSheetNames,
  parseSharedStrings,
  parseSheets,
} from '../src/xlsx/xmlParser.js';

describe('parseSheets / extractSheetNames', () => {
  it('parses a flat workbook with three sheets', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <sheets>
          <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
          <sheet name="Data" sheetId="2" r:id="rId2"/>
          <sheet name="Config" sheetId="3" r:id="rId3"/>
        </sheets>
      </workbook>`;
    expect(extractSheetNames(xml)).toEqual(['Sheet1', 'Data', 'Config']);
    expect(parseSheets(xml)[0]).toEqual({ name: 'Sheet1', rId: 'rId1', sheetId: '1' });
  });

  it('handles namespace-prefixed elements', () => {
    const xml = `<x:workbook xmlns:x="..." xmlns:r="...">
      <x:sheets><x:sheet name="A" sheetId="1" r:id="rId1"/></x:sheets>
    </x:workbook>`;
    expect(extractSheetNames(xml)).toEqual(['A']);
  });

  it('decodes entities inside the name attribute', () => {
    const xml = `<workbook xmlns:r="">
      <sheets><sheet name="Tom &amp; Jerry" sheetId="1" r:id="rId1"/></sheets>
    </workbook>`;
    expect(extractSheetNames(xml)).toEqual(['Tom & Jerry']);
  });

  it('accepts a different namespace prefix on the relationship-id attribute', () => {
    const xml = `<workbook xmlns:rels="...">
      <sheets><sheet name="A" sheetId="1" rels:id="rId7"/></sheets>
    </workbook>`;
    expect(parseSheets(xml)[0]?.rId).toBe('rId7');
  });

  it('accepts a bare id attribute (no namespace prefix) — xmlParser.ts:115', () => {
    // Some non-Excel producers emit id= without a namespace prefix.
    // This exercises the relIdValue fallback: return attrValue(attrs, 'id').
    const xml = `<workbook>
      <sheets><sheet name="Sheet1" sheetId="1" id="rId42"/></sheets>
    </workbook>`;
    const sheets = parseSheets(xml);
    expect(sheets[0]?.rId).toBe('rId42');
    expect(sheets[0]?.name).toBe('Sheet1');
  });
});

describe('parseSharedStrings', () => {
  it('extracts plain <si><t> entries', () => {
    const xml = `<sst>
      <si><t>hello</t></si>
      <si><t>world</t></si>
    </sst>`;
    expect(parseSharedStrings(xml)).toEqual(['hello', 'world']);
  });

  it('concatenates rich-text runs in order', () => {
    const xml = `<sst><si>
      <r><rPr><b/></rPr><t>Bold</t></r>
      <r><t> and </t></r>
      <r><rPr><i/></rPr><t>italic</t></r>
    </si></sst>`;
    expect(parseSharedStrings(xml)).toEqual(['Bold and italic']);
  });

  it('strips phonetic guides (rPh) so they do not contaminate the value', () => {
    const xml = `<sst><si>
      <t>東京</t>
      <rPh sb="0" eb="2"><t>とうきょう</t></rPh>
    </si></sst>`;
    expect(parseSharedStrings(xml)).toEqual(['東京']);
  });

  it('decodes XML entities inside text', () => {
    const xml = `<sst><si><t>Tom &amp; Jerry &lt;3</t></si></sst>`;
    expect(parseSharedStrings(xml)).toEqual(['Tom & Jerry <3']);
  });

  it('preserves whitespace in <t xml:space="preserve">', () => {
    const xml = `<sst><si><t xml:space="preserve">  hello  </t></si></sst>`;
    expect(parseSharedStrings(xml)).toEqual(['  hello  ']);
  });

  it('handles namespace prefixes', () => {
    const xml = `<x:sst><x:si><x:t>a</x:t></x:si></x:sst>`;
    expect(parseSharedStrings(xml)).toEqual(['a']);
  });

  it('returns empty string for empty <si><t></t></si>', () => {
    const xml = `<sst><si><t></t></si></sst>`;
    expect(parseSharedStrings(xml)).toEqual(['']);
  });

  it('handles self-closing <si/> (empty shared string)', () => {
    const xml = `<sst>
      <si><t>first</t></si>
      <si/>
      <si><t>third</t></si>
    </sst>`;
    // Self-closing <si/> is treated as an empty string
    expect(parseSharedStrings(xml)).toEqual(['first', '', 'third']);
  });

  it('parseSingleSharedString: extracts plain text from a single <si> element', () => {
    const xml = `<si><t>hello world</t></si>`;
    // Note: parseSharedStrings is for full <sst>, but we test the shared string parsing behavior
    // by confirming plain text works in the full context
    const fullXml = `<sst>${xml}</sst>`;
    expect(parseSharedStrings(fullXml)).toEqual(['hello world']);
  });

  it('parseSingleSharedString: concatenates multiple <r> runs in <si>', () => {
    const xml = `<sst><si>
      <r><t>Part1</t></r>
      <r><t>Part2</t></r>
      <r><t>Part3</t></r>
    </si></sst>`;
    expect(parseSharedStrings(xml)).toEqual(['Part1Part2Part3']);
  });

  it('extracts sheets with missing name attribute (handles malformed workbooks gracefully)', () => {
    const xml = `<?xml version="1.0"?>
      <workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <sheets>
          <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
          <sheet sheetId="2" r:id="rId2"/>
        </sheets>
      </workbook>`;
    const sheets = parseSheets(xml);
    // Should gracefully handle missing name attribute
    expect(sheets).toHaveLength(1); // Only Sheet1 is parsed (malformed sheet skipped)
    expect(sheets[0]).toEqual({ name: 'Sheet1', rId: 'rId1', sheetId: '1' });
  });
});
