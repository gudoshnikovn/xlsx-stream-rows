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

  it('handles AfterQuote state with junk characters after closing quote', () => {
    // After closing quote, junk is ignored/skipped in lenient parser
    expect(parseAll('"a"b,c')).toEqual([['a', 'c']]);
  });

  it('handles empty fields after pushing empty string', () => {
    // Test parser behavior with consecutive delimiters (empty fields)
    expect(parseAll('a,,b')).toEqual([['a', '', 'b']]);
    expect(parseAll('a,,,b')).toEqual([['a', '', '', 'b']]);
  });

  it('handles delimiter-only rows (all commas, no data)', () => {
    expect(parseAll(',,,')).toEqual([['', '', '', '']]);
  });

  // ─── CR at exact chunk boundary (pendingCR = true path) ──────────────────────

  it('correctly handles \\r at the end of a chunk (FieldStart state, parser.ts:104-105)', () => {
    // The \\r ends the first chunk with no \\n — pendingCR = true is set.
    // The second chunk starts with \\n (which is consumed by the pendingCR handler).
    const p = createCsvParser();
    const first = p.push('row1\r');   // \\r at chunk boundary → pendingCR = true
    const second = p.push('\nrow2'); // pending \\r consumes the \\n
    const last = p.end();
    expect([...first, ...second, ...last]).toEqual([['row1'], ['row2']]);
  });

  it('\\r at chunk boundary without following \\n (FieldStart, pendingCR then non-LF)', () => {
    // pendingCR = true, but next chunk starts with something other than \\n
    const p = createCsvParser();
    const first = p.push('row1\r');
    const second = p.push('row2'); // no \\n follows the \\r → row2 is a new row
    const last = p.end();
    expect([...first, ...second, ...last]).toEqual([['row1'], ['row2']]);
  });

  it('correctly handles \\r at end of chunk in AfterQuote state (parser.ts:175-176)', () => {
    // Quoted field followed by \\r at chunk boundary.
    const p = createCsvParser();
    const first = p.push('"val"\r'); // \\r at chunk boundary in AfterQuote state
    const second = p.push('\nnext'); // \\n consumed by pending CR handler
    const last = p.end();
    expect([...first, ...second, ...last]).toEqual([['val'], ['next']]);
  });
});
