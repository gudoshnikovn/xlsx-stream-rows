import { defineConfig } from 'vitest/config';

/**
 * Browser-mode configuration. Runs the same test suite (minus Node-only
 * tests) inside a real browser via Playwright, validating that
 * `DecompressionStream`, `Blob.stream()`, `TextDecoderStream`, and `File`
 * behave as expected on the platforms we target.
 *
 * Usage:
 *   npm run test:browser                 # default: chromium
 *   npm run test:browser -- --browser.name=firefox
 *   npm run test:browser -- --browser.name=webkit
 *
 * Run `npx playwright install <browser>` once per machine.
 */
export default defineConfig({
  test: {
    include: ['tests/**/*.test.ts'],
    // Memory smoke uses Node's `process.memoryUsage` — not available in
    // browser. Already gated by env var, but exclude defensively.
    exclude: ['tests/memory.smoke.test.ts', 'tests/xlsAdapter.test.ts'],
    browser: {
      enabled: true,
      provider: 'playwright',
      headless: true,
      name: 'chromium',
    },
  },
});
