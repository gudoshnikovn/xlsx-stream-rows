import { describe, expect, it } from 'vitest';

import { decodeXml } from '../src/decodeXml.js';

describe('decodeXml', () => {
  it('passes through strings without entities (fast path)', () => {
    expect(decodeXml('hello world')).toBe('hello world');
    expect(decodeXml('')).toBe('');
  });

  it('decodes the five named entities', () => {
    expect(decodeXml('&amp;')).toBe('&');
    expect(decodeXml('&lt;')).toBe('<');
    expect(decodeXml('&gt;')).toBe('>');
    expect(decodeXml('&quot;')).toBe('"');
    expect(decodeXml('&apos;')).toBe("'");
  });

  it('decodes decimal character references', () => {
    expect(decodeXml('&#39;')).toBe("'");
    expect(decodeXml('&#65;')).toBe('A');
    expect(decodeXml('&#1046;')).toBe('Ж');
  });

  it('decodes hexadecimal character references (both x and X)', () => {
    expect(decodeXml('&#x27;')).toBe("'");
    expect(decodeXml('&#X27;')).toBe("'");
    expect(decodeXml('&#x41;')).toBe('A');
    expect(decodeXml('&#x1F600;')).toBe('😀');
  });

  it('decodes mixed entities in one string', () => {
    expect(decodeXml('Tom &amp; Jerry &lt;&gt; &quot;hi&quot;')).toBe('Tom & Jerry <> "hi"');
  });

  it('passes through unknown named entities verbatim', () => {
    expect(decodeXml('&nbsp;')).toBe('&nbsp;');
    expect(decodeXml('&unknown;')).toBe('&unknown;');
  });

  it('passes through malformed numeric refs verbatim', () => {
    expect(decodeXml('&#abc;')).toBe('&#abc;');
    // Out-of-range code points
    expect(decodeXml('&#x110000;')).toBe('&#x110000;');
  });

  it('handles entities at string boundaries', () => {
    expect(decodeXml('&amp;hello')).toBe('&hello');
    expect(decodeXml('hello&amp;')).toBe('hello&');
    expect(decodeXml('&amp;&amp;')).toBe('&&');
  });
});
