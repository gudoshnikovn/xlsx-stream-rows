/**
 * Incremental RFC 4180 CSV parser.
 *
 * Designed for streaming consumption: feed text chunks via `push()`, receive
 * complete rows as they finish. State persists across calls so chunk
 * boundaries inside a field, inside a quoted region, or in the middle of a
 * CRLF line ending all parse correctly. The fuzz suite
 * (`tests/csvParser.fuzz.test.ts`) splits the same input at every byte
 * position and verifies identical output.
 *
 * RFC 4180 rules applied:
 *   - Comma is the field separator. (No semicolon/tab auto-detect.)
 *   - Fields may be wrapped in `"…"`. A literal `"` inside a quoted field is
 *     written as `""`.
 *   - CRLF is the canonical line ending; we also accept bare LF and bare CR.
 *   - Whitespace is preserved verbatim.
 *
 * Empty-row semantics: a newline always finalises the current row (even an
 * empty one), so `"a\n\nb"` yields three rows `[["a"], [""], ["b"]]`. A
 * trailing newline does **not** produce a phantom empty row at EOF — the
 * `end()` flush only emits when the state machine is mid-field/mid-row.
 *
 * Non-strict tolerances (real-world files):
 *   - Characters appearing immediately after a closing quote and before the
 *     next delimiter are silently dropped (RFC says they are forbidden).
 *   - A bare `"` inside an unquoted field is taken literally.
 *   - An unterminated quoted field at EOF flushes the accumulated text as
 *     the field's value.
 */

export interface CsvParser {
  /** Feed a chunk of decoded text. Returns rows completed during this push. */
  push(chunk: string): string[][];
  /** Flush a trailing partial field/row, if any. Call once at EOF. */
  end(): string[][];
}

const enum State {
  /** At the very start of a field — next char decides quoted vs unquoted. */
  FieldStart = 0,
  /** Inside an unquoted field. */
  Unquoted = 1,
  /** Inside a quoted field. */
  Quoted = 2,
  /** Just consumed a `"` while in `Quoted`; awaiting follow-up. */
  AfterQuote = 3,
}

const CC_QUOTE = 0x22;
const CC_CR = 0x0d;
const CC_LF = 0x0a;

export function createCsvParser(separator = ','): CsvParser {
  const CC_SEP = separator.charCodeAt(0);
  let state: State = State.FieldStart;
  let field = '';
  let row: string[] = [];
  /** True if we just consumed a `\r` whose `\n` partner might land in the
   * next chunk; suppresses a duplicate row emission. */
  let pendingCR = false;
  let completed: string[][] = [];

  function endField(): void {
    row.push(field);
    field = '';
    state = State.FieldStart;
  }

  function emitRow(): void {
    row.push(field);
    completed.push(row);
    field = '';
    row = [];
    state = State.FieldStart;
  }

  function processChunk(buf: string): void {
    const n = buf.length;
    let i = 0;

    if (pendingCR) {
      pendingCR = false;
      if (i < n && buf.charCodeAt(i) === CC_LF) i++;
    }

    while (i < n) {
      switch (state) {
        case State.FieldStart: {
          const c = buf.charCodeAt(i);
          if (c === CC_QUOTE) {
            state = State.Quoted;
            i++;
          } else if (c === CC_SEP) {
            endField();
            i++;
          } else if (c === CC_LF) {
            emitRow();
            i++;
          } else if (c === CC_CR) {
            emitRow();
            i++;
            if (i < n) {
              if (buf.charCodeAt(i) === CC_LF) i++;
            } else {
              pendingCR = true;
            }
          } else {
            state = State.Unquoted;
          }
          break;
        }
        case State.Unquoted: {
          // Fast scan to the next field/row terminator.
          let j = i;
          while (j < n) {
            const c = buf.charCodeAt(j);
            if (c === CC_SEP || c === CC_CR || c === CC_LF) break;
            j++;
          }
          if (j > i) {
            field += buf.slice(i, j);
            i = j;
          }
          if (i >= n) return; // wait for more
          const c = buf.charCodeAt(i);
          if (c === CC_SEP) {
            endField();
            i++;
          } else if (c === CC_LF) {
            emitRow();
            i++;
          } else {
            // CR
            emitRow();
            i++;
            if (i < n) {
              if (buf.charCodeAt(i) === CC_LF) i++;
            } else {
              pendingCR = true;
            }
          }
          break;
        }
        case State.Quoted: {
          // Scan to next `"` — quoted regions cross delimiters and newlines
          // verbatim per RFC 4180.
          const q = buf.indexOf('"', i);
          if (q === -1) {
            field += buf.slice(i);
            return; // wait for more
          }
          if (q > i) field += buf.slice(i, q);
          state = State.AfterQuote;
          i = q + 1;
          break;
        }
        case State.AfterQuote: {
          const c = buf.charCodeAt(i);
          if (c === CC_QUOTE) {
            // Escaped quote: `""` → literal `"`
            field += '"';
            state = State.Quoted;
            i++;
          } else if (c === CC_SEP) {
            endField();
            i++;
          } else if (c === CC_LF) {
            emitRow();
            i++;
          } else if (c === CC_CR) {
            emitRow();
            i++;
            if (i < n) {
              if (buf.charCodeAt(i) === CC_LF) i++;
            } else {
              pendingCR = true;
            }
          } else {
            // Junk after closing quote — silently consume.
            i++;
          }
          break;
        }
      }
    }
  }

  function push(chunk: string): string[][] {
    completed = [];
    if (chunk.length > 0) processChunk(chunk);
    const out = completed;
    completed = [];
    return out;
  }

  function end(): string[][] {
    completed = [];
    // Treat an unterminated quoted field as if its closing quote arrived.
    if (state === State.Quoted) state = State.AfterQuote;
    // Flush only when the state machine is genuinely mid-field or mid-row.
    if (
      state === State.Unquoted ||
      state === State.AfterQuote ||
      row.length > 0 ||
      field.length > 0
    ) {
      emitRow();
    }
    pendingCR = false;
    const out = completed;
    completed = [];
    return out;
  }

  return { push, end };
}
