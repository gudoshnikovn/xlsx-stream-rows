import { describe, expect, it } from 'vitest';

import { excelSerialToDate } from '../src/excelDate.js';

const utc = (y: number, m: number, d: number, h = 0, mi = 0, s = 0): number =>
  Date.UTC(y, m - 1, d, h, mi, s);

describe('excelSerialToDate', () => {
  it('maps serial 1 to 1900-01-01 UTC', () => {
    expect(excelSerialToDate(1).getTime()).toBe(utc(1900, 1, 1));
  });

  it('maps serial 2 to 1900-01-02 UTC', () => {
    expect(excelSerialToDate(2).getTime()).toBe(utc(1900, 1, 2));
  });

  it('maps serial 59 to 1900-02-28 UTC (last real date before the 1900 leap-year bug)', () => {
    expect(excelSerialToDate(59).getTime()).toBe(utc(1900, 2, 28));
  });

  it('maps serial 61 to 1900-03-01 UTC (first date after the spurious leap day)', () => {
    expect(excelSerialToDate(61).getTime()).toBe(utc(1900, 3, 1));
  });

  it('maps serial 25569 to 1970-01-01 UTC (the Unix epoch alignment point)', () => {
    expect(excelSerialToDate(25569).getTime()).toBe(0);
  });

  it('maps serial 44197 to 2021-01-01 UTC (a contemporary check)', () => {
    expect(excelSerialToDate(44197).getTime()).toBe(utc(2021, 1, 1));
  });

  it('encodes fractional time-of-day as milliseconds', () => {
    // 44197.5 = 2021-01-01 12:00:00 UTC
    expect(excelSerialToDate(44197.5).getTime()).toBe(utc(2021, 1, 1, 12, 0, 0));
  });
});
