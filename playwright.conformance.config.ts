import { defineConfig } from '@playwright/test';

// Conformance runner config. Reads every `docx/*.json` manifest in the
// sibling `ooxml-reference-corpus` checkout, renders the fixture through
// docxjs, evaluates each render_assertion, and writes one result JSON
// per feature under `../ooxml-validate/conformance/results/docxjs/`.
//
// Kept separate from `playwright.config.ts` so:
//   - npm test (smoke suite) is unaffected.
//   - npm run test:conformance runs ONLY the conformance spec.
//   - Failure semantics (see tests/conformance/docx-conformance.spec.ts)
//     only error out on evaluator errors, not fail verdicts.
export default defineConfig({
    testDir: './tests/conformance',
    testMatch: ['**/*.spec.ts'],
    // Deterministic iteration order keeps the per-run log readable and
    // each result file is independent, so there's no ordering risk.
    fullyParallel: true,
    workers: 4,
    retries: 0,
    reporter: 'list',
    use: {
        baseURL: 'http://localhost:8765',
        trace: 'on-first-retry',
        screenshot: 'only-on-failure',
        // Matches the interop config; a stable viewport size keeps the
        // visual_ssim cropped screenshot reproducible between runs.
        viewport: { width: 1024, height: 1200 },
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
        url: 'http://localhost:8765/tests/harness.html',
        reuseExistingServer: !process.env.CI,
        timeout: 30_000,
    },
});
