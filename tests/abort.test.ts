/**
 * Unit tests for the `abortable` utility function (src/utils/abort.ts).
 *
 * Covers the two branches missed by integration tests:
 *   - Line 10: signal is already aborted when abortable() is called
 *     (returns a pre-rejected Promise without registering any listener)
 *   - Line 14: onAbort callback fires while the wrapped Promise is still pending
 *     (mid-flight abort path)
 */

import { describe, expect, it } from 'vitest';

import { abortable, checkAbort } from '../src/utils/abort.js';

describe('checkAbort', () => {
  it('does nothing when signal is undefined', () => {
    expect(() => checkAbort(undefined)).not.toThrow();
  });

  it('does nothing when signal is not aborted', () => {
    const ac = new AbortController();
    expect(() => checkAbort(ac.signal)).not.toThrow();
  });

  it('throws the abort reason when signal is already aborted', () => {
    const ac = new AbortController();
    const reason = new Error('cancelled');
    ac.abort(reason);
    expect(() => checkAbort(ac.signal)).toThrow(reason);
  });
});

describe('abortable', () => {
  it('returns the original promise when signal is undefined', async () => {
    const result = await abortable(Promise.resolve(42), undefined);
    expect(result).toBe(42);
  });

  it('rejects immediately when signal is already aborted before call (abort.ts:10)', async () => {
    // signal.aborted is true at the time abortable() is called.
    // This exercises line 10: return Promise.reject(signal.reason ?? ...)
    const ac = new AbortController();
    const reason = new Error('pre-aborted');
    ac.abort(reason);

    await expect(abortable(Promise.resolve(99), ac.signal)).rejects.toBe(reason);
  });

  it('rejects with DOMException when signal has no reason and was already aborted (abort.ts:10)', async () => {
    const ac = new AbortController();
    ac.abort(); // no explicit reason → signal.reason is undefined
    await expect(abortable(Promise.resolve(1), ac.signal)).rejects.toBeInstanceOf(DOMException);
  });

  it('rejects with abort reason when signal fires mid-flight (abort.ts:14)', async () => {
    // Creates a promise that never resolves, then aborts the signal.
    // This exercises the onAbort listener registered at line 16 which calls
    // reject(signal.reason ...) at line 14.
    const ac = new AbortController();
    const reason = new Error('mid-flight abort');

    const neverResolves = new Promise<never>(() => { /* intentionally empty */ });
    const wrapped = abortable(neverResolves, ac.signal);

    // Abort after the listener has been registered.
    ac.abort(reason);

    await expect(wrapped).rejects.toBe(reason);
  });

  it('resolves normally when signal is provided but never fired', async () => {
    const ac = new AbortController(); // never aborted
    const result = await abortable(Promise.resolve('ok'), ac.signal);
    expect(result).toBe('ok');
  });

  it('rejects with underlying error when the wrapped promise rejects (non-abort path)', async () => {
    const ac = new AbortController(); // never aborted
    const err = new Error('original error');
    await expect(abortable(Promise.reject(err), ac.signal)).rejects.toBe(err);
  });
});
