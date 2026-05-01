// Node harness that loads a DOCX fixture through the built library in a jsdom
// environment and asserts the track-changes pipeline produces the expected
// attributes, classes, and sidebar cards. Not a replacement for the Karma
// e2e suite — used as a smoke test for the Phase 1-3 changes so we catch
// regressions that `npm run build` alone doesn't surface.

import { readFileSync } from 'node:fs';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { JSDOM } from 'jsdom';

const here = dirname(fileURLToPath(import.meta.url));
const repo = resolve(here, '..');

// Stand up a DOM before importing the library — the library captures
// `document` / `Range` / `HTMLElement` from globals at module evaluation time.
const dom = new JSDOM('<!doctype html><html><body></body></html>', {
    url: 'http://localhost/',
});
globalThis.window = dom.window;
globalThis.document = dom.window.document;
globalThis.Range = dom.window.Range;
globalThis.HTMLElement = dom.window.HTMLElement;
globalThis.Node = dom.window.Node;
globalThis.Element = dom.window.Element;
globalThis.DOMParser = dom.window.DOMParser;
globalThis.XMLSerializer = dom.window.XMLSerializer;
globalThis.requestAnimationFrame = (cb) => setTimeout(cb, 0);
globalThis.cancelAnimationFrame = (h) => clearTimeout(h);
// Stub out Highlight/CSS.highlights — the comments code path uses them but
// they're optional (guarded by `globalThis.Highlight`).
globalThis.CSS = { highlights: new Map() };

// The UMD bundle is the one `npm run build` writes on every invocation
// (the `.mjs` output is only produced on `--environment BUILD:production`).
// Evaluate it against our globals so we can consume the `docx` namespace.
await import('jszip').then((m) => {
    globalThis.JSZip = m.default;
});
const umd = readFileSync(`${repo}/dist/docx-preview.js`, 'utf8');
// UMD checks `typeof exports === 'object'` and `typeof define`; with neither
// present it attaches to `globalThis` under the `name` configured in rollup
// ("docx"). We run it as a Function so it sees our globals.
new Function('require', umd)(() => ({})); // require stub for any node fallback
const { parseAsync, renderDocument } = globalThis.docx;

const failures = [];
const warnings = [];
function assert(cond, msg) {
    if (!cond) failures.push(msg);
}
function note(msg) {
    warnings.push(msg);
}

async function renderFixture(path, options) {
    const buf = readFileSync(resolve(repo, 'tests/render-test', path, 'document.docx'));
    // The library's parseAsync accepts anything JSZip can load — a Uint8Array
    // or Buffer works fine in Node.
    const doc = await parseAsync(buf, options);
    const nodes = await renderDocument(doc, options);
    const container = document.createElement('div');
    for (const n of nodes) container.appendChild(n);
    return { doc, container };
}

// ── 1. Revision fixture, changes.show = false (legacy off-state) ───────────
{
    const { container } = await renderFixture('revision', { renderChanges: false });
    assert(
        container.querySelectorAll('ins, del').length === 0,
        '1a: revision fixture with changes off should render no <ins>/<del>',
    );
    assert(
        container.querySelectorAll('.docx-change-bar').length === 0,
        '1b: no change bars when changes are off',
    );
}

// ── 2. Revision fixture, changes.show = true ───────────────────────────────
{
    const { container } = await renderFixture('revision', {
        changes: { show: true, legend: true, changeBar: true },
    });
    const insEls = container.querySelectorAll('ins');
    const delEls = container.querySelectorAll('del');
    note(`2·: revision fixture produced ${insEls.length} <ins> and ${delEls.length} <del>`);
    assert(
        insEls.length + delEls.length > 0,
        '2a: revision fixture should produce at least one <ins> or <del>',
    );
    // Every rendered change should have a change-id and author palette class.
    const changes = [...insEls, ...delEls];
    for (const el of changes) {
        assert(
            el.dataset.changeKind === 'insertion' || el.dataset.changeKind === 'deletion',
            `2b: change element missing data-change-kind (got ${el.dataset.changeKind})`,
        );
        assert(
            el.dataset.changeId !== undefined,
            '2c: change element missing data-change-id',
        );
        assert(
            [...el.classList].some((c) => c.startsWith('docx-change-author-')),
            '2d: change element missing author palette class',
        );
    }
    // Change bars should appear on the paragraph ancestors.
    const bars = container.querySelectorAll('.docx-change-bar');
    assert(bars.length > 0, '2e: expected at least one paragraph to carry a change bar');
    // Legend should list the authors we saw.
    const legendItems = container.querySelectorAll('.docx-legend-item');
    assert(
        legendItems.length > 0,
        '2f: legend should list at least one author',
    );
    note(`2g: legend rendered ${legendItems.length} author(s)`);
}

// ── 3. Backwards compatibility: renderChanges: true with no `changes` block ─
{
    const { container } = await renderFixture('revision', { renderChanges: true });
    const changes = container.querySelectorAll('ins, del');
    assert(
        changes.length > 0,
        '3a: legacy renderChanges:true should still produce <ins>/<del>',
    );
    assert(
        container.querySelectorAll('.docx-change-bar').length > 0,
        '3b: legacy renderChanges:true should also produce change bars',
    );
}

// ── 4. Prototype-pollution guard (paraId) ──────────────────────────────────
// Not directly testable without a crafted DOCX, but at least assert that the
// guard function is in the shipped bundle. Grep the source of the bundle.
{
    const bundle = readFileSync(`${repo}/dist/docx-preview.js`, 'utf8');
    assert(
        bundle.includes('SAFE_PARA_ID'),
        '4a: SAFE_PARA_ID regex guard should be present in the bundle',
    );
}

// ── 5. Non-revision fixture: the new code should be completely inert ───────
{
    const { container } = await renderFixture('text', { changes: { show: true } });
    // Even with changes.show on, a document with no revisions should have no
    // change-bars, no legend, no ins/del.
    assert(
        container.querySelectorAll('.docx-change-bar, .docx-legend').length === 0,
        '5a: document with no revisions should have no change-bar / legend decoration',
    );
}

// ── 6. Readonly / accept-reject plumbing ──────────────────────────────────
{
    const calls = [];
    const { container } = await renderFixture('revision', {
        changes: { show: true, readOnly: false },
        changeCallbacks: {
            onChangeAccept: (id, kind) => calls.push(['accept', id, kind]),
            onChangeReject: (id, kind) => calls.push(['reject', id, kind]),
        },
    });
    // Each change element should have an injected ✓/✕ button wrapper.
    const actions = container.querySelectorAll('.docx-change-actions');
    assert(
        actions.length > 0,
        '6a: with readOnly:false, change action buttons should be injected',
    );
    // Simulate a click on the first Accept.
    const firstAccept = container.querySelector('.docx-change-accept');
    if (firstAccept) {
        // bubble so the delegated listener on the wrapper picks it up
        firstAccept.dispatchEvent(new dom.window.MouseEvent('click', { bubbles: true, cancelable: true }));
    }
    assert(
        calls.length === 1 && calls[0][0] === 'accept',
        `6b: clicking ✓ should invoke onChangeAccept once (got ${JSON.stringify(calls)})`,
    );
}

// ── report ─────────────────────────────────────────────────────────────────
console.log('--- track-changes harness ---');
for (const w of warnings) console.log(`  · ${w}`);
if (failures.length) {
    console.error(`\n${failures.length} FAILURE(S):`);
    for (const f of failures) console.error(`  ✗ ${f}`);
    process.exit(1);
} else {
    console.log(`\n✓ all ${6} scenarios passed`);
}
