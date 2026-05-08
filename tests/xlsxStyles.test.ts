import { describe, expect, it } from 'vitest';

import { parseDateFormatMask } from '../src/xlsx/styles.js';

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

  it('detects date format with lowercase date tokens (yyyy, mm, dd)', () => {
    const xml = `<styleSheet>
      <numFmts count="1">
        <numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>
      </numFmts>
      <cellXfs count="1">
        <xf numFmtId="164"/>
      </cellXfs>
    </styleSheet>`;
    expect(parseDateFormatMask(xml)).toEqual(new Set([0]));
  });

  it('detects date format with mixed case tokens (YYYY, MM, DD)', () => {
    const xml = `<styleSheet>
      <numFmts count="1">
        <numFmt numFmtId="164" formatCode="YYYY-MM-DD"/>
      </numFmts>
      <cellXfs count="1">
        <xf numFmtId="164"/>
      </cellXfs>
    </styleSheet>`;
    expect(parseDateFormatMask(xml)).toEqual(new Set([0]));
  });

  it('ignores quoted literals in format code (e.g., "mm" inside quotes is not a date token)', () => {
    const xml = `<styleSheet>
      <numFmts count="1">
        <numFmt numFmtId="164" formatCode="&quot;minutes: mm&quot;"/>
      </numFmts>
      <cellXfs count="1">
        <xf numFmtId="164"/>
      </cellXfs>
    </styleSheet>`;
    expect(parseDateFormatMask(xml).size).toBe(0);
  });

  it('flags entries with built-in date format IDs 45-47 (Excel built-in date/time formats)', () => {
    const xml = `<styleSheet>
      <cellXfs count="4">
        <xf numFmtId="0"/>
        <xf numFmtId="45"/>
        <xf numFmtId="46"/>
        <xf numFmtId="47"/>
      </cellXfs>
    </styleSheet>`;
    const mask = parseDateFormatMask(xml);
    expect(mask.has(1)).toBe(true);
    expect(mask.has(2)).toBe(true);
    expect(mask.has(3)).toBe(true);
  });

  it('handles xf entries without numFmtId attribute (missing numFmtId defaults to 0)', () => {
    const xml = `<styleSheet>
      <cellXfs count="2">
        <xf/>
        <xf numFmtId="14"/>
      </cellXfs>
    </styleSheet>`;
    const mask = parseDateFormatMask(xml);
    expect(mask.has(0)).toBe(false);
    expect(mask.has(1)).toBe(true);
  });

  // Branch coverage for the numFmt parsing loop (styles.ts:73-78)

  it('skips numFmt entry when numFmtId attribute is missing (styles.ts:76 continue branch)', () => {
    // idStr === undefined → continue; the entry is ignored, no custom date IDs added
    const xml = `<styleSheet>
      <numFmts count="1">
        <numFmt formatCode="yyyy-mm-dd"/>
      </numFmts>
      <cellXfs count="1">
        <xf numFmtId="164"/>
      </cellXfs>
    </styleSheet>`;
    // numFmt has no numFmtId, so it is skipped; xf references 164 which is not in built-ins
    expect(parseDateFormatMask(xml).size).toBe(0);
  });

  it('skips numFmt entry when formatCode attribute is missing (styles.ts:76 continue branch)', () => {
    // codeStr === undefined → continue; the entry is ignored
    const xml = `<styleSheet>
      <numFmts count="1">
        <numFmt numFmtId="164"/>
      </numFmts>
      <cellXfs count="1">
        <xf numFmtId="164"/>
      </cellXfs>
    </styleSheet>`;
    // numFmt has no formatCode → skipped; 164 not a built-in date ID
    expect(parseDateFormatMask(xml).size).toBe(0);
  });

  it('skips numFmt entry when numFmtId is non-numeric (styles.ts:78 continue branch)', () => {
    // Number.parseInt("abc") = NaN → !Number.isFinite(NaN) → continue
    const xml = `<styleSheet>
      <numFmts count="1">
        <numFmt numFmtId="abc" formatCode="yyyy-mm-dd"/>
      </numFmts>
      <cellXfs count="1">
        <xf numFmtId="164"/>
      </cellXfs>
    </styleSheet>`;
    expect(parseDateFormatMask(xml).size).toBe(0);
  });
});
