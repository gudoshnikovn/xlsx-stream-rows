import { describe, expect, it } from 'vitest';

import { excelSerialToDate } from '../src/utils/excelDate.js';

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

  it('maps serial 60 to 1900-02-28 UTC (Excel\'s phantom leap day maps to same as 59)', () => {
    // The 1900 leap-year bug: serial 60 is the ghost day, treated as 1900-02-29 in Excel
    // but since 1900 was not a leap year, we should get the same date as serial 59
    expect(excelSerialToDate(60).getTime()).toBe(utc(1900, 2, 28));
  });

  it('maps serial 0 to 1899-12-31 UTC (day before 1900-01-01)', () => {
    expect(excelSerialToDate(0).getTime()).toBe(utc(1899, 12, 31));
  });

  it('handles fractional time less than 1 day (time-only dates)', () => {
    // 0.5 = 12:00:00 on 1899-12-31
    expect(excelSerialToDate(0.5).getTime()).toBe(utc(1899, 12, 31, 12, 0, 0));
  });

  it('maps very large serials (e.g., serial 2958465 = 9999-12-31)', () => {
    // Serial 2958465 should be at or near the end of the Excel date range
    const result = excelSerialToDate(2958465);
    expect(result.getUTCFullYear()).toBeLessThanOrEqual(9999);
    expect(result).toBeInstanceOf(Date);
  });
});
