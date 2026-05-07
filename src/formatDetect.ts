/**
 * File format detection by magic bytes (preferred) and filename extension
 * (fallback). Reads only the first 8 bytes of the file.
 */

import { FormatNotSupportedError } from './errors.js';

export type SpreadsheetFormat = 'xlsx' | 'xls' | 'csv';

const MAGIC_ZIP = [0x50, 0x4b, 0x03, 0x04]; // "PK\x03\x04" — XLSX/XLSM
const MAGIC_OLE2 = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]; // XLS

function startsWith(bytes: Uint8Array, magic: readonly number[]): boolean {
  if (bytes.length < magic.length) return false;
  for (let i = 0; i < magic.length; i++) {
    if (bytes[i] !== magic[i]) return false;
  }
  return true;
}

function extensionOf(filename: string): string {
  const i = filename.lastIndexOf('.');
  return i === -1 ? '' : filename.slice(i + 1).toLowerCase();
}

/**
 * Detect the format of `file` by:
 *   1. Reading the first 8 bytes and matching against ZIP / OLE2 magic.
 *   2. Falling back to the filename extension when no magic matches.
 *
 * Throws `FormatNotSupportedError` if neither path produces a known format.
 */
export async function detectFormat(file: File): Promise<SpreadsheetFormat> {
  if (file.size > 0) {
    const head = new Uint8Array(
      await file.slice(0, Math.min(file.size, 8)).arrayBuffer(),
    );
    if (startsWith(head, MAGIC_ZIP)) return 'xlsx';
    if (startsWith(head, MAGIC_OLE2)) return 'xls';
  }

  switch (extensionOf(file.name)) {
    case 'xlsx':
    case 'xlsm':
      return 'xlsx';
    case 'xls':
      return 'xls';
    case 'csv':
      return 'csv';
    default:
      throw new FormatNotSupportedError(file.name);
  }
}
