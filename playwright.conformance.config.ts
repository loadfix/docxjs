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
        baseURL: `http://localhost:${process.env.PORT || 8765}`,
        trace: 'on-first-retry',
        screenshot: 'only-on-failure',
        // Letter page is 8.5in x 11in at 150dpi = 1275x1650, matching
        // the PDF-derived reference PNGs in the corpus. Set a viewport
        // slightly larger than the wrapper so Playwright doesn't clip
        // the screenshot; the visual_ssim evaluator crops to the
        // .docx-wrapper bounding box.
        viewport: { width: 1400, height: 1800 },
        deviceScaleFactor: 1.5625,
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
