/**
 * Chunk-boundary fuzz for the CSV parser. Splits the same input at every
 * byte position and verifies identical output — the load-bearing
 * correctness check for the RFC-4180 state machine.
 */

import { describe, expect, it } from 'vitest';

import { createCsvParser } from '../src/csv/parser.js';

function parseInOneShot(input: string): string[][] {
  const p = createCsvParser();
  const rows = p.push(input);
  rows.push(...p.end());
  return rows;
}

function parseAtSplit(input: string, splitAt: number): string[][] {
  const p = createCsvParser();
  const rows: string[][] = [];
  rows.push(...p.push(input.slice(0, splitAt)));
  rows.push(...p.push(input.slice(splitAt)));
  rows.push(...p.end());
  return rows;
}

function parseAtTriSplit(input: string, a: number, b: number): string[][] {
  const p = createCsvParser();
  const rows: string[][] = [];
  rows.push(...p.push(input.slice(0, a)));
  rows.push(...p.push(input.slice(a, b)));
  rows.push(...p.push(input.slice(b)));
  rows.push(...p.end());
  return rows;
}

function expectAllSplitsMatch(input: string): void {
  const expected = parseInOneShot(input);
  for (let i = 0; i <= input.length; i++) {
    expect(parseAtSplit(input, i), `split at byte ${i}`).toEqual(expected);
  }
}

describe('csvParser chunk-boundary fuzz', () => {
  it('survives every binary split of mixed quoted / unquoted content', () => {
    expectAllSplitsMatch('a,b,"c,d","e\nf"\nx,y,z\n"q""q",last');
  });

  it('survives every binary split across CRLF / LF / CR line endings', () => {
    expectAllSplitsMatch('a,b\r\nc,d\ne,f\rg,h');
  });

  it('survives every binary split inside escaped-quote runs', () => {
    expectAllSplitsMatch('"""quoted""","plain"');
  });

  it('survives every binary split of a unicode-heavy input', () => {
    expectAllSplitsMatch('тест,"привет, мир"\n东京,"\n北京"\n');
  });

  it('survives 3-way splits across larger inputs', () => {
    let body = '';
    for (let i = 0; i < 30; i++) {
      body += `row${i},"v,${i}","x\n${i}"\n`;
    }
    const expected = parseInOneShot(body);
    for (let a = 0; a <= body.length; a += 17) {
      for (let b = a; b <= body.length; b += 17) {
        expect(parseAtTriSplit(body, a, b), `splits at ${a},${b}`).toEqual(expected);
      }
    }
  });
});
