// Playwright conformance runner — docxjs edition.
//
// Discovers every manifest in `../ooxml-reference-corpus/features/docx/`
// that carries a `render_assertions` block, renders the committed
// machine fixture through docxjs in a headless page, evaluates each
// render_assertion, and writes a result JSON shaped exactly like
// `FeatureResult.to_dict()` from
// `ooxml-validate/src/ooxml_validate/conformance.py`.
//
// Output:
//   ../ooxml-validate/conformance/results/docxjs/<feature_id>.json
// where feature_id is e.g. `docx/bold-text`, so the file lands at
// `docxjs/docx/bold-text.json`.
//
// Test-failure semantics mirror the Python runner:
//   - a `fail` verdict is real signal (we write it) but does NOT fail
//     the test — the matrix page is where that surfaces;
//   - an `error` verdict DOES fail the test, because it means the
//     assertion could not be evaluated at all (page load failed,
//     fixture couldn't be fetched, etc.).
//
// The corpus is an optional sibling checkout. If it's missing, every
// test in this file is `.skip`-ed with a pointer to the clone command.

import { test, expect } from '@playwright/test';
import {
    existsSync,
    mkdirSync,
    readFileSync,
    readdirSync,
    statSync,
    writeFileSync,
} from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import {
    evaluateAssertion,
    type AssertionVerdict,
    type RenderAssertion,
    type VisualSsimAssertion,
} from './evaluator';

// Layout on disk:
//   /home/ben/code/docxjs/                       (this repo)
//   /home/ben/code/ooxml-reference-corpus/       (the corpus)
//   /home/ben/code/ooxml-validate/               (result sink)
// Resolved from this spec so relative paths survive `cwd` changes.
const REPO_ROOT = resolve(__dirname, '..', '..');
const CORPUS_ROOT = resolve(REPO_ROOT, '..', 'ooxml-reference-corpus');
const MANIFESTS_DIR = join(CORPUS_ROOT, 'features', 'docx');
const FIXTURES_ROOT = join(CORPUS_ROOT, 'fixtures');
const REFS_ROOT = join(CORPUS_ROOT, 'refs');
const VALIDATE_ROOT = resolve(REPO_ROOT, '..', 'ooxml-validate');
const RESULTS_ROOT = join(VALIDATE_ROOT, 'conformance', 'results');

const LIBRARY = 'docxjs';
// Free-form version string, captured in every result JSON for
// traceability. Mirrors the FeatureResult.tool_version field.
const TOOL_VERSION = (() => {
    try {
        const pkg = JSON.parse(readFileSync(join(REPO_ROOT, 'package.json'), 'utf-8'));
        return String(pkg.version ?? '');
    } catch {
        return '';
    }
})();

interface ManifestEntry {
    featureName: string;       // e.g. "bold-text"
    featureId: string;         // e.g. "docx/bold-text"
    manifestPath: string;      // absolute
    fixturePath: string;       // absolute, resolved via fixtures.machine
    renderAssertions: RenderAssertion[];
    manifest: any;
}

function listManifests(): ManifestEntry[] {
    if (!existsSync(MANIFESTS_DIR)) return [];
    const out: ManifestEntry[] = [];
    for (const name of readdirSync(MANIFESTS_DIR).sort()) {
        if (!name.endsWith('.json')) continue;
        const p = join(MANIFESTS_DIR, name);
        if (!statSync(p).isFile()) continue;
        let manifest: any;
        try {
            manifest = JSON.parse(readFileSync(p, 'utf-8'));
        } catch {
            continue;
        }
        const renderAssertions: RenderAssertion[] | undefined = manifest.render_assertions;
        if (!renderAssertions || renderAssertions.length === 0) continue;
        // `roles` defaults to ['authoring','rendering'] per the schema.
        // Include the feature unless it explicitly omits 'rendering'.
        const roles: string[] | undefined = manifest.roles;
        if (roles && !roles.includes('rendering')) continue;
        const machine: string | undefined = manifest?.fixtures?.machine;
        const fixturePath = machine ? join(FIXTURES_ROOT, `${machine}.docx`) : '';
        out.push({
            featureName: name.replace(/\.json$/, ''),
            featureId: manifest.id ?? `docx/${name.replace(/\.json$/, '')}`,
            manifestPath: p,
            fixturePath,
            renderAssertions,
            manifest,
        });
    }
    return out;
}

/**
 * Resolve a visual_ssim assertion's reference PNG. If the manifest
 * overrides `reference_png`, it's corpus-root-relative; otherwise fall
 * back to `refs/docx/<feature-name>-page1.png` which matches the
 * filenames the corpus currently ships.
 */
function resolveReferencePng(
    entry: ManifestEntry,
): (a: VisualSsimAssertion) => string {
    return (a) => {
        if (a.reference_png) {
            return resolve(CORPUS_ROOT, a.reference_png);
        }
        return join(REFS_ROOT, 'docx', `${entry.featureName}-page1.png`);
    };
}

/**
 * Produce a FeatureResult.to_dict()-shaped object from a verdict list
 * plus metadata. Keys and order match the Python dataclass exactly so
 * the matrix page can read either source without branching.
 */
function makeResultJson(
    entry: ManifestEntry,
    verdicts: AssertionVerdict[],
): Record<string, unknown> {
    // Aggregate status: error beats fail beats pass. Matches the
    // Python `run_feature` escalation rule.
    let aggregate: 'pass' | 'fail' | 'error' = 'pass';
    for (const v of verdicts) {
        if (v.status === 'error') { aggregate = 'error'; break; }
        if (v.status === 'fail' && aggregate === 'pass') aggregate = 'fail';
    }
    return {
        feature_id: entry.featureId,
        library: LIBRARY,
        status: aggregate,
        fixture_path: entry.fixturePath,
        // ISO-8601 UTC with Z suffix, second precision — same format
        // as the Python side (`"%Y-%m-%dT%H:%M:%SZ"`).
        run_at: new Date().toISOString().replace(/\.\d{3}Z$/, 'Z'),
        tool_version: TOOL_VERSION,
        assertions: verdicts.map((v) => ({
            id: v.id,
            status: v.status,
            detail: v.detail,
        })),
    };
}

function writeResultJson(
    entry: ManifestEntry,
    result: Record<string, unknown>,
): string {
    // Output path: results/<library>/<feature_id>.json. feature_id
    // contains a slash ("docx/bold-text") so the library dir gets a
    // nested `docx/` child, matching the README layout.
    const out = join(RESULTS_ROOT, LIBRARY, `${entry.featureId}.json`);
    mkdirSync(dirname(out), { recursive: true });
    writeFileSync(out, JSON.stringify(result, null, 2) + '\n', 'utf-8');
    return out;
}

const manifests = listManifests();

test.describe('docx conformance', () => {
    if (!existsSync(MANIFESTS_DIR)) {
        test.skip(`corpus not found at ${CORPUS_ROOT}`, () => {
            // Clone it with:
            //   git clone https://github.com/loadfix/ooxml-reference-corpus \
            //       ../ooxml-reference-corpus
        });
        return;
    }
    if (!existsSync(VALIDATE_ROOT)) {
        test.skip(`ooxml-validate not found at ${VALIDATE_ROOT}`, () => {
            // Clone it with:
            //   git clone https://github.com/loadfix/ooxml-validate \
            //       ../ooxml-validate
        });
        return;
    }
    if (manifests.length === 0) {
        test.skip('no docx manifests with render_assertions under features/docx/', () => {});
        return;
    }

    for (const entry of manifests) {
        test.describe(entry.featureId, () => {
            // Guard: a manifest can point at a fixture that isn't yet in
            // the corpus (e.g. `heading-1.json` declares
            // `fixtures.machine` but `heading-1.docx` hasn't been
            // committed). That's a corpus-side data gap, not a docxjs
            // regression — skip with a loud message and still write an
            // error-status result JSON so the matrix knows why the cell
            // is empty.
            if (!entry.fixturePath || !existsSync(entry.fixturePath)) {
                test.skip(`renders and evaluates render_assertions (fixture missing)`, () => {
                    // Skipping — fixture file: ${entry.fixturePath || '(fixtures.machine absent)'}
                });
                const result = makeResultJson(entry, [{
                    id: '<fixture>',
                    status: 'error',
                    detail: `fixture not found: ${entry.fixturePath || '(fixtures.machine missing)'}`,
                }]);
                writeResultJson(entry, result);
                return;
            }

            test(`renders and evaluates render_assertions`, async ({ page }) => {

                await page.goto('/tests/harness.html');

                // Fetch the fixture bytes Node-side and hand them to the
                // page as a byte array — simpler than extending the dev
                // server to serve the sibling corpus path, and consistent
                // with the interop runner's approach.
                const fixtureBytes = Array.from(readFileSync(entry.fixturePath));

                const renderOutcome = await page.evaluate(async ({ name, bytes }) => {
                    // Fresh container every run so no residual wrapper
                    // from a prior test bleeds through.
                    const existing = document.querySelector('#render-root');
                    if (existing) existing.remove();
                    const container = document.createElement('div');
                    container.id = 'render-root';
                    document.body.appendChild(container);
                    const blob = new Blob([new Uint8Array(bytes)], {
                        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                    });
                    try {
                        // eslint-disable-next-line @typescript-eslint/no-explicit-any
                        await (window as any).docx.renderAsync(blob, container);
                        const wrappers = document.querySelectorAll('.docx-wrapper');
                        return { ok: true, wrapperCount: wrappers.length };
                    } catch (e) {
                        return { ok: false, error: (e as Error).message };
                    }
                }, { name: `${entry.featureName}.docx`, bytes: fixtureBytes });

                // Collect verdicts across all assertions, even when the
                // render itself failed — the result JSON needs to tell
                // the matrix page *why* we couldn't evaluate.
                const verdicts: AssertionVerdict[] = [];
                if (!renderOutcome.ok) {
                    for (const a of entry.renderAssertions) {
                        verdicts.push({
                            id: a.id,
                            status: 'error',
                            detail: `docx.renderAsync failed: ${renderOutcome.error}`,
                        });
                    }
                } else if (renderOutcome.wrapperCount === 0) {
                    for (const a of entry.renderAssertions) {
                        verdicts.push({
                            id: a.id,
                            status: 'error',
                            detail: 'renderAsync produced no .docx-wrapper element.',
                        });
                    }
                } else {
                    const refResolver = resolveReferencePng(entry);
                    for (const a of entry.renderAssertions) {
                        // eslint-disable-next-line no-await-in-loop
                        const verdict = await evaluateAssertion(page, a, refResolver);
                        verdicts.push(verdict);
                    }
                }

                const resultJson = makeResultJson(entry, verdicts);
                const written = writeResultJson(entry, resultJson);
                console.log(`[${entry.featureId}] ${resultJson.status} -> ${written}`);

                // Per the prompt: `fail` verdicts are recorded without
                // failing the test (they're real signal). Only `error`
                // verdicts — where we literally couldn't evaluate —
                // fail the Playwright test.
                expect(
                    resultJson.status,
                    `aggregate status=${resultJson.status}; details: ${JSON.stringify(resultJson.assertions, null, 2)}`,
                ).not.toBe('error');
            });
        });
    }
});
