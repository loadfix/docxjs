// @ts-check
//
// Visual-diff interop harness.
//
// Discovers feature manifests in the sibling `ooxml-reference-corpus`
// checkout, renders each associated `.docx` fixture through the docxjs demo
// page, and screenshots the rendered `.docx-wrapper` for comparison against
// a committed baseline PNG.
//
// The corpus is NOT required for the repo to build. If the sibling checkout
// is missing, every test in this file is skipped with a pointer to where it
// should live.
//
// Run modes (see tests/interop/README.md for full usage):
//   npm run test:interop             # compare to committed baselines
//   npm run test:interop -- --update-snapshots   # regenerate baselines

import { test, expect } from '@playwright/test';
import { readdirSync, readFileSync, existsSync, statSync } from 'node:fs';
import { join, resolve } from 'node:path';

// Resolve the sibling corpus relative to this file. Layout on disk:
//   /home/ben/code/docxjs/                       (this repo)
//   /home/ben/code/ooxml-reference-corpus/       (the corpus)
const CORPUS_ROOT = resolve(__dirname, '..', '..', '..', 'ooxml-reference-corpus');
const MANIFESTS_DIR = join(CORPUS_ROOT, 'features', 'docx');
const FIXTURES_ROOT = join(CORPUS_ROOT, 'fixtures');

interface ManifestEntry {
    /** Relative manifest path under features/docx/, e.g. "bold-text.json". */
    manifestFile: string;
    /** Feature slug without extension, e.g. "bold-text". */
    featureName: string;
    /** Absolute path to the .docx fixture this manifest points at. */
    fixturePath: string;
    /** The manifest's `id` field, e.g. "docx/bold-text". */
    manifestId: string;
}

/**
 * Enumerate every `*.json` manifest under `features/docx/` and resolve its
 * `.docx` fixture via `fixtures.machine`. Returns an empty list if the
 * corpus directory is absent — callers are expected to branch on that.
 */
function listManifests(): ManifestEntry[] {
    if (!existsSync(MANIFESTS_DIR)) return [];
    const entries: ManifestEntry[] = [];
    for (const name of readdirSync(MANIFESTS_DIR).sort()) {
        if (!name.endsWith('.json')) continue;
        const manifestPath = join(MANIFESTS_DIR, name);
        if (!statSync(manifestPath).isFile()) continue;
        let manifest: any;
        try {
            manifest = JSON.parse(readFileSync(manifestPath, 'utf-8'));
        } catch (err) {
            // Corrupt manifest — surface as a failing test rather than
            // silently skipping the whole suite.
            entries.push({
                manifestFile: name,
                featureName: name.replace(/\.json$/, ''),
                fixturePath: '',
                manifestId: `INVALID:${(err as Error).message}`,
            });
            continue;
        }
        // fixtures.machine is a slash-separated logical name like
        // "docx/bold-text" — resolve to fixtures/docx/bold-text.docx.
        const machine: string | undefined = manifest?.fixtures?.machine;
        const fixturePath = machine
            ? join(FIXTURES_ROOT, `${machine}.docx`)
            : '';
        entries.push({
            manifestFile: name,
            featureName: name.replace(/\.json$/, ''),
            fixturePath,
            manifestId: manifest?.id ?? '',
        });
    }
    return entries;
}

const manifests = listManifests();
const corpusPresent = existsSync(MANIFESTS_DIR);

test.describe('interop visual diff', () => {
    if (!corpusPresent) {
        // One explicit skipped test so the run output tells the user *why*
        // everything is inert. The message names the expected sibling path.
        test.skip(`corpus not found at ${CORPUS_ROOT}`, () => {
            // Clone it with:
            //   git clone https://github.com/loadfix/ooxml-reference-corpus \
            //       ../ooxml-reference-corpus
            // then re-run `npm run test:interop`.
        });
        return;
    }

    if (manifests.length === 0) {
        test.skip('no manifests under features/docx/', () => {
            // The corpus is present but features/docx/ is empty. Likely a
            // pre-release checkout; nothing to compare against.
        });
        return;
    }

    for (const m of manifests) {
        const label = `renders docx/${m.featureName}`;

        if (m.manifestId.startsWith('INVALID:')) {
            test(label, () => {
                throw new Error(`manifest ${m.manifestFile} is not valid JSON: ${m.manifestId.slice('INVALID:'.length)}`);
            });
            continue;
        }

        if (!m.fixturePath || !existsSync(m.fixturePath)) {
            // A manifest without a resolvable fixture is a corpus-side
            // problem; skip rather than fail so the rest of the suite still
            // runs. The message pinpoints the missing file.
            test.skip(label, () => {
                // Manifest declares fixtures.machine but the .docx is
                // missing: ${m.fixturePath || '(fixtures.machine absent)'}
            });
            continue;
        }

        test(label, async ({ page }) => {
            await page.goto('/');
            // The demo's `#files` <input> triggers renderDocx on change.
            await page.locator('#files').setInputFiles(m.fixturePath);
            // `.docx-wrapper` is emitted by renderAsync once the document
            // finishes parsing — waiting for it guarantees screenshots
            // capture the rendered output, not the empty chrome.
            await page.waitForSelector('.docx-wrapper', { timeout: 15_000 });
            // A short settling window so late-loading resources (fonts,
            // images) paint before capture. Most fixtures are text-only.
            await page.waitForTimeout(250);

            // Baseline file name is stable and OS-agnostic inside the repo;
            // Playwright stores platform-specific variants under
            // `{projectName}/{spec}-snapshots/`.
            await expect(page.locator('.docx-wrapper')).toHaveScreenshot(
                `${m.featureName}.png`,
            );
        });
    }
});
