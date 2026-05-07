/**
 * Decode XML character references and the five predefined entities.
 *
 * Supported:
 *   - Named entities: &amp; &lt; &gt; &quot; &apos;  (XML 1.0 §4.6)
 *   - Numeric decimal: &#NNNN;
 *   - Numeric hex:     &#xHHHH;  (case-insensitive)
 *
 * Unknown entities are left as-is — XLSX files emitted by spec-compliant
 * producers (Excel, LibreOffice) only ever use the five named entities and
 * numeric character references, so silently passing through anything else
 * keeps the function fast on the common path.
 */

const NAMED: Record<string, string> = {
  amp: '&',
  lt: '<',
  gt: '>',
  quot: '"',
  apos: "'",
};

export function decodeXml(s: string): string {
  // Fast path — most XML chunks contain no entities at all.
  if (s.indexOf('&') === -1) return s;

  return s.replace(/&(#[xX][0-9a-fA-F]+|#\d+|[a-zA-Z]+);/g, (match, body: string) => {
    if (body.charCodeAt(0) === 35 /* '#' */) {
      const code =
        body.charCodeAt(1) === 120 || body.charCodeAt(1) === 88 // 'x' or 'X'
          ? parseInt(body.slice(2), 16)
          : parseInt(body.slice(1), 10);
      if (!Number.isFinite(code) || code < 0 || code > 0x10ffff) return match;
      try {
        return String.fromCodePoint(code);
      } catch {
        return match;
      }
    }
    const named = NAMED[body];
    return named !== undefined ? named : match;
  });
}
