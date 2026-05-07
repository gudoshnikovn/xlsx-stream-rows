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
});
