import { defineConfig } from '@playwright/test';

// Visual-diff interop harness against the manifests in the sibling
// ooxml-reference-corpus checkout. Kept separate from `playwright.config.ts`
// so that:
//   - The main suite (tests/render-test) is unaffected.
//   - `npm run test:interop` runs *only* the interop runner.
//   - Screenshot output lands in its own report dir.
//
// See tests/interop/README.md for usage.
export default defineConfig({
    testDir: './tests/interop',
    testMatch: ['**/*.spec.ts'],
    // Screenshot comparison needs consistent rendering, so force single-worker
    // sequential runs. The overhead of a fresh `#files` upload per test is
    // small compared to the determinism gain.
    fullyParallel: false,
    workers: 1,
    retries: 0,
    reporter: [['list'], ['html', { outputFolder: 'playwright-report-interop', open: 'never' }]],
    outputDir: 'test-results-interop',
    // Pin baseline storage to `tests/interop/baselines/<feature>.png`.
    // Playwright's default is `<spec>-snapshots/<name>-<project>-<platform>.png`;
    // flattening it keeps the committed set tidy and matches the directory
    // the README documents.
    snapshotPathTemplate: 'tests/interop/baselines/{arg}{ext}',
    use: {
        baseURL: 'http://localhost:8765',
        trace: 'on-first-retry',
        screenshot: 'only-on-failure',
        // Deterministic viewport so baselines are comparable across runs.
        viewport: { width: 1024, height: 1200 },
    },
    // Playwright's built-in screenshot comparison. Threshold is the
    // per-pixel colour-distance tolerance (0..1); maxDiffPixels caps the
    // number of differing pixels overall. Tuned for anti-aliasing and
    // minor font-rendering drift but tight enough to catch real regressions.
    expect: {
        toHaveScreenshot: {
            threshold: 0.001,
            maxDiffPixels: 100,
        },
    },
    projects: [
        {
            name: 'chrome',
            use: {
                browserName: 'chromium',
                channel: 'chrome',
                headless: true,
                // The system Chrome channel shares a user-data-dir across
                // every launch by default, so leftover tab state (an open
                // file-input selection, a cached render) from the user's
                // interactive Chrome session can bleed into test pages.
                // Force a throwaway profile dir AND fresh storage state
                // per test — between them we get hermetic isolation.
                storageState: { cookies: [], origins: [] },
                launchOptions: {
                    args: [
                        '--no-default-browser-check',
                        '--no-first-run',
                        '--disable-features=PersistentStorage',
                        '--disable-background-networking',
                        '--disable-session-crashed-bubble',
                        '--disable-restore-session-state',
                    ],
                },
            },
        },
    ],
    webServer: {
        command: 'node scripts/dev-server.mjs',
        url: 'http://localhost:8765/',
        // Reuse the dev server when running locally alongside other Playwright
        // configs; start a fresh one in CI.
        reuseExistingServer: !process.env.CI,
        timeout: 30_000,
    },
});
