import { defineConfig } from '@playwright/test';

// Playwright is configured to drive the system Chrome install (channel: 'chrome')
// because the Playwright-managed Chromium bundle does not yet publish binaries
// for ubuntu26.04 on this host. The tests themselves are plain DOM/fetch and
// don't depend on anything Chrome-specific.
export default defineConfig({
  testDir: './tests',
  testMatch: ['**/*.spec.ts', '**/*.spec.js'],
  // Conformance and interop suites have their own configs; keep the
  // smoke suite (`npm test`) scoped to tests/render-test & friends.
  testIgnore: ['**/tests/conformance/**', '**/tests/interop/**'],
  fullyParallel: true,
  retries: 0,
  workers: 4,
  reporter: 'list',
  use: {
    baseURL: `http://localhost:${process.env.PORT || 8765}`,
    trace: 'on-first-retry',
    screenshot: 'only-on-failure',
  },
  projects: [
    {
      name: 'chrome',
      use: {
        browserName: 'chromium',
        channel: 'chrome',
      },
    },
  ],
  webServer: {
    command: 'node scripts/dev-server.mjs',
    url: `http://localhost:${process.env.PORT || 8765}/tests/harness.html`,
    reuseExistingServer: !process.env.CI,
    timeout: 30_000,
  },
});
