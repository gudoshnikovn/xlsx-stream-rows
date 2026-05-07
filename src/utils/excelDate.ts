/**
 * Convert an Excel date serial number to a JavaScript `Date`.
 *
 * Excel encodes dates as the number of days since "1900-01-01" with a famous
 * quirk: it inherits Lotus 1-2-3's bug of treating 1900 as a leap year, so
 * serial 60 maps to a non-existent "1900-02-29". We follow Excel's actual
 * behaviour (compatible with Microsoft and LibreOffice):
 *
 *   serial < 60   → days since 1900-01-01 (no shift)
 *   serial >= 60  → days since 1900-01-01 minus one day (the spurious leap day)
 *
 * The fractional part is interpreted as time-of-day (0.5 = 12:00, etc.).
 *
 * Returned `Date` is UTC-based — Excel serials carry no timezone, so we use
 * UTC throughout to keep the result stable across machines.
 */

const EPOCH_MS = Date.UTC(1900, 0, 1); // 1900-01-01 00:00:00 UTC
const MS_PER_DAY = 86_400_000;

export function excelSerialToDate(serial: number): Date {
  const adjusted = serial < 60 ? serial : serial - 1;
  return new Date(EPOCH_MS + (adjusted - 1) * MS_PER_DAY);
}
