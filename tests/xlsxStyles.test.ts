import { describe, expect, it } from 'vitest';

import { parseDateFormatMask } from '../src/xlsxStyles.js';

describe('parseDateFormatMask', () => {
  it('flags cellXfs entries that reference built-in date formats', () => {
    // numFmtId 0 = General (numeric), 14 = m/d/yyyy (date), 22 = m/d/yyyy h:mm
    const xml = `<styleSheet>
      <cellXfs count="3">
        <xf numFmtId="0"/>
        <xf numFmtId="14" applyNumberFormat="1"/>
        <xf numFmtId="22" applyNumberFormat="1"/>
      </cellXfs>
    </styleSheet>`;
    const mask = parseDateFormatMask(xml);
    expect(mask.has(0)).toBe(false);
    expect(mask.has(1)).toBe(true);
    expect(mask.has(2)).toBe(true);
  });

  it('flags entries that reference a custom date numFmt', () => {
    const xml = `<styleSheet>
      <numFmts count="1">
        <numFmt numFmtId="164" formatCode="dd.mm.yyyy"/>
      </numFmts>
      <cellXfs count="2">
        <xf numFmtId="0"/>
        <xf numFmtId="164" applyNumberFormat="1"/>
      </cellXfs>
    </styleSheet>`;
    expect(parseDateFormatMask(xml)).toEqual(new Set([1]));
  });

  it('does not flag custom non-date numFmts', () => {
    const xml = `<styleSheet>
      <numFmts count="1">
        <numFmt numFmtId="164" formatCode="#,##0.00 &quot;руб.&quot;"/>
      </numFmts>
      <cellXfs count="2">
        <xf numFmtId="0"/>
        <xf numFmtId="164"/>
      </cellXfs>
    </styleSheet>`;
    expect(parseDateFormatMask(xml).size).toBe(0);
  });

  it('ignores date tokens that appear only inside quoted literals', () => {
    const xml = `<styleSheet>
      <numFmts count="1">
        <numFmt numFmtId="164" formatCode="#,##0 &quot;y in stock&quot;"/>
      </numFmts>
      <cellXfs count="1">
        <xf numFmtId="164"/>
      </cellXfs>
    </styleSheet>`;
    expect(parseDateFormatMask(xml).size).toBe(0);
  });

  it('handles namespace-prefixed style elements', () => {
    const xml = `<x:styleSheet>
      <x:cellXfs count="2">
        <x:xf numFmtId="0"/>
        <x:xf numFmtId="14"/>
      </x:cellXfs>
    </x:styleSheet>`;
    expect(parseDateFormatMask(xml)).toEqual(new Set([1]));
  });

  it('returns an empty mask when cellXfs is missing', () => {
    expect(parseDateFormatMask('<styleSheet/>').size).toBe(0);
  });
});
