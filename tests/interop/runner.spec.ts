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
            // Ensure the demo is ready — #files must exist and renderDocx
            // must be exposed on window. If the demo is still evaluating
            // its <script>, setInputFiles could fire before the listener
            // is attached and the file silently goes nowhere.
            await page.waitForFunction(
                () => !!document.querySelector('#files') && typeof (window as any).renderDocx === 'function',
                { timeout: 10_000 },
            );

            // The demo's `#files` <input> triggers renderDocx on change.
            // We call renderDocx directly via page.evaluate rather than
            // relying on the input's change event so we can AWAIT the
            // render — the event-based path returns control before the
            // async render finishes, which caused stale-render leakage
            // during initial harness development.
            const fixtureBytes = readFileSync(m.fixturePath);
            const renderResult = await page.evaluate(async ({ name, bytes }) => {
                // Wipe the document container up-front so any residual
                // wrapper from a previous render can't be screenshotted.
                const container = document.querySelector('#document-container');
                if (container) container.innerHTML = '';
                const thumbs = document.querySelector('#thumbnails-container');
                if (thumbs) thumbs.innerHTML = '';

                const file = new File([new Uint8Array(bytes)], name, {
                    type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                });
                const fn = (window as any).renderDocx;
                if (typeof fn !== 'function') return { ok: false, error: 'renderDocx missing' };
                try {
                    await fn(file);
                    const wrappers = document.querySelectorAll('.docx-wrapper');
                    const text = wrappers[0]?.textContent?.slice(0, 120) ?? '';
                    return { ok: true, wrapperCount: wrappers.length, firstText: text };
                } catch (e) {
                    return { ok: false, error: (e as Error).message };
                }
            }, { name: `${m.featureName}.docx`, bytes: Array.from(fixtureBytes) });
            // Surface render status in the test log so baseline drift is
            // easy to triage.
            if (!renderResult.ok) {
                throw new Error(`renderDocx failed for ${m.featureName}: ${renderResult.error}`);
            }
            console.log(`[${m.featureName}] wrappers=${renderResult.wrapperCount} text="${renderResult.firstText}"`);

            // `.docx-wrapper` is emitted by renderAsync once the document
            // finishes parsing. We scope to the main `#document-container`
            // because the demo *also* calls renderThumbnails, which creates
            // a second .docx-wrapper-class subtree inside
            // #thumbnails-container.
            const wrapper = page.locator('#document-container .docx-wrapper');
            await wrapper.waitFor({ timeout: 15_000 });
            // Cross-check the wrapper's text against what renderDocx told
            // us above — if they diverge we've got a stale wrapper and the
            // screenshot would capture the wrong fixture.
            const wrapperText = await wrapper.evaluate((el) => el.textContent?.slice(0, 120) ?? '');
            if (wrapperText !== renderResult.firstText) {
                throw new Error(
                    `[${m.featureName}] wrapper text mismatch — expected "${renderResult.firstText}", saw "${wrapperText}"`,
                );
            }
            // A short settling window so late-loading resources (fonts,
            // images) paint before capture. Most fixtures are text-only.
            await page.waitForTimeout(250);

            // Baseline file name is stable and OS-agnostic inside the repo;
            // Playwright stores platform-specific variants under
            // `{projectName}/{spec}-snapshots/`.
            await expect(wrapper).toHaveScreenshot(`${m.featureName}.png`);
        });
    }
});
