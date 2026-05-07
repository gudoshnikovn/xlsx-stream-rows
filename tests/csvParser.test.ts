import { describe, expect, it } from 'vitest';

import { createCsvParser } from '../src/csv/parser.js';

function parseAll(input: string): string[][] {
  const p = createCsvParser();
  const rows = p.push(input);
  rows.push(...p.end());
  return rows;
}

describe('createCsvParser', () => {
  it('parses a simple comma-separated row', () => {
    expect(parseAll('a,b,c')).toEqual([['a', 'b', 'c']]);
  });

  it('parses multiple rows separated by LF', () => {
    expect(parseAll('a,b\nc,d\n')).toEqual([
      ['a', 'b'],
      ['c', 'd'],
    ]);
  });

  it('parses CRLF and bare-CR line endings', () => {
    expect(parseAll('a\r\nb\r\nc')).toEqual([['a'], ['b'], ['c']]);
    expect(parseAll('a\rb\rc')).toEqual([['a'], ['b'], ['c']]);
  });

  it('handles empty fields', () => {
    expect(parseAll('a,,b')).toEqual([['a', '', 'b']]);
    expect(parseAll(',a,')).toEqual([['', 'a', '']]);
  });

  it('preserves whitespace verbatim (no trimming per RFC 4180)', () => {
    expect(parseAll(' a , b ')).toEqual([[' a ', ' b ']]);
  });

  it('parses quoted fields containing commas', () => {
    expect(parseAll('"a,b",c')).toEqual([['a,b', 'c']]);
  });

  it('parses quoted fields containing newlines', () => {
    expect(parseAll('"line1\nline2",x')).toEqual([['line1\nline2', 'x']]);
  });

  it('decodes escaped quotes ("" → ")', () => {
    expect(parseAll('"a""b","c""d"')).toEqual([['a"b', 'c"d']]);
  });

  it('handles a single trailing newline without emitting an extra row', () => {
    expect(parseAll('a,b\n')).toEqual([['a', 'b']]);
  });

  it('does emit empty rows that were explicitly present', () => {
    expect(parseAll('a\n\nb')).toEqual([['a'], [''], ['b']]);
  });

  it('parses unicode content', () => {
    expect(parseAll('тест,привет\n东京,北京')).toEqual([
      ['тест', 'привет'],
      ['东京', '北京'],
    ]);
  });

  it('flushes a trailing field when the file ends without a newline', () => {
    expect(parseAll('a,b,c')).toEqual([['a', 'b', 'c']]);
  });

  it('flushes an unterminated quoted field at EOF as the field value', () => {
    // RFC says malformed; we stay lenient.
    expect(parseAll('"unterminated')).toEqual([['unterminated']]);
  });

  it('returns no rows for empty input', () => {
    expect(parseAll('')).toEqual([]);
  });
});
