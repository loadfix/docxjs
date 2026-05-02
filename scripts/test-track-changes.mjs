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
const {
    parseAsync, renderDocument, renderThumbnails, defaultOptions,
    isSafeHyperlinkHref, sanitizeCssColor, sanitizeFontFamily,
    isSafeCssIdent, escapeCssStringContent, keyBy, mergeDeep,
} = globalThis.docx;

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

// ── 6. Rich fixture: formatting revisions (w:rPrChange) ────────────────────
{
    const { container } = await renderFixture('revision-rich', {
        changes: { show: true, showFormatting: true },
    });
    const formatting = container.querySelectorAll('.docx-formatting-revision');
    note(`6·: revision-rich produced ${formatting.length} formatting revision(s)`);
    assert(
        formatting.length > 0,
        '6a: revision-rich fixture should produce at least one formatting-revision element',
    );
    for (const el of formatting) {
        assert(
            el.dataset.changeKind === 'formatting',
            '6b: formatting-revision element missing data-change-kind=formatting',
        );
        assert(
            el.getAttribute('title') && el.getAttribute('title').includes(':'),
            '6c: formatting-revision should have a title summarising the change',
        );
    }
    // Change bars should light up on the paragraphs touched by rPrChange even
    // though there's no <ins>/<del> in the doc.
    assert(
        container.querySelectorAll('.docx-change-bar').length > 0,
        '6d: formatting-only revisions should still trigger change bars',
    );
    // showFormatting:false should suppress them.
    const { container: hidden } = await renderFixture('revision-rich', {
        changes: { show: true, showFormatting: false },
    });
    assert(
        hidden.querySelectorAll('.docx-formatting-revision').length === 0,
        '6e: showFormatting:false should suppress formatting revisions',
    );
}

// ── 7. Read-only: no mutation UI on comment cards or change elements ──────
{
    const { container } = await renderFixture('revision', {
        renderComments: true,
        changes: { show: true },
    });
    // Edit / Delete / Reply buttons should never appear — the library is
    // read-only. Same for inline ✓/✕ change-accept/reject buttons and the
    // sidebar Accept-all / Reject-all toolbar buttons.
    const forbidden = [
        '.docx-comment-edit-btn',
        '.docx-comment-delete-btn',
        '.docx-comment-reply-btn',
        '.docx-comment-add-btn',
        '.docx-change-accept',
        '.docx-change-reject',
        '.docx-change-actions',
        '.docx-comment-editor',
        '.docx-comment-reply-composer',
        '.docx-new-comment-composer',
    ];
    for (const sel of forbidden) {
        assert(
            container.querySelectorAll(sel).length === 0,
            `7a: no ${sel} element should ever be rendered`,
        );
    }
}

// ── 8. Sidebar layout modes ────────────────────────────────────────────────
{
    // Packed: no marginTop offsets on any card. No scroll listener overhead.
    const { container: packed } = await renderFixture('revision', {
        renderComments: true,
        comments: { sidebar: true, layout: 'packed' },
    });
    const packedMargins = [...packed.querySelectorAll('.docx-sidebar-comment')]
        .map(el => el.style.marginTop || '');
    assert(
        packedMargins.every(m => m === ''),
        `8a: packed mode should leave marginTop untouched (got ${JSON.stringify(packedMargins)})`,
    );

    // Anchored: setupSidebarScrollSync runs. In jsdom layout is trivial so we
    // don't assert specific offsets — just that the pass doesn't throw.
    // (Real alignment is verified in the browser via Playwright.)
    const { container: anchored } = await renderFixture('revision', {
        renderComments: true,
        comments: { sidebar: true, layout: 'anchored' },
    });
    assert(
        !!anchored.querySelector('.docx-comment-sidebar'),
        '8b: anchored mode should still render the sidebar',
    );
}

// ── 9. Comments sidebar has no Comments collapse/expand toggle ─────────────
{
    const { container } = await renderFixture('revision', {
        renderComments: true,
        comments: { sidebar: true },
    });
    assert(
        container.querySelectorAll('.docx-sidebar-toggle').length === 0,
        '9a: no .docx-sidebar-toggle elements should exist',
    );
    assert(
        container.querySelectorAll('.docx-sidebar-collapsed').length === 0,
        '9b: no .docx-sidebar-collapsed elements should exist',
    );
}

// ── 10. Thumbnail API ─────────────────────────────────────────────────────
// Thumbnails-per-page: the API should produce at least one thumbnail per
// rendered <section>. In jsdom there's no layout so paginateSection falls
// back to 1 thumbnail per section, which is what we assert here. The
// multi-page path is covered by Playwright in-browser checks.
{
    assert(typeof renderThumbnails === 'function', '10a: docx.renderThumbnails export should be a function');

    const { container: main } = await renderFixture('text');
    document.body.appendChild(main);
    const thumbs = document.createElement('div');
    document.body.appendChild(thumbs);

    const sectionCount = main.querySelectorAll('section.docx').length;
    note(`10·: text fixture produced ${sectionCount} section(s)`);
    assert(sectionCount > 0, '10b: fixture should render at least one section');

    const handle = renderThumbnails(main, thumbs);
    assert(typeof handle?.dispose === 'function', '10c: renderThumbnails should return { dispose }');

    const thumbEls = thumbs.querySelectorAll('.docx-thumbnail');
    assert(
        thumbEls.length >= sectionCount,
        `10d: expected >= ${sectionCount} thumbnails (one per page), got ${thumbEls.length}`,
    );
    for (let i = 0; i < thumbEls.length; i++) {
        const t = thumbEls[i];
        assert(!!t.querySelector('.docx-thumbnail-preview'), `10e·${i}: missing preview element`);
        assert(t.querySelector('.docx-thumbnail-label')?.textContent === String(i + 1), `10f·${i}: label should be "${i + 1}"`);
        assert(t.getAttribute('tabindex') === '0', `10g·${i}: tabindex=0 expected`);
    }

    // Idempotent re-run: should replace, not append.
    const handle2 = renderThumbnails(main, thumbs);
    assert(
        thumbs.querySelectorAll('.docx-thumbnail').length === thumbEls.length,
        '10h: re-running should replace (not append)',
    );

    // Style injected once.
    assert(
        document.head.querySelectorAll('style[data-docxjs-thumbnails]').length === 1,
        '10i: exactly one injected <style> expected',
    );

    // dispose clears the thumbnail container.
    handle2.dispose();
    assert(thumbs.children.length === 0, '10j: dispose() should clear the thumbnail container');

    main.remove();
    thumbs.remove();
}

// ── 11. Security: altChunk removed, hyperlink scheme allowlist ────────────
// These come from SECURITY_REVIEW.md findings 1 and 2. The fixes strip the
// altChunk renderer entirely and validate hyperlink hrefs against an
// allowlist before emitting them. No DOCX fixture needed — we check the
// exported surface directly.
{
    assert(
        !Object.prototype.hasOwnProperty.call(defaultOptions, 'renderAltChunks'),
        '11a: defaultOptions.renderAltChunks should be removed (SECURITY_REVIEW.md #1)',
    );
    // No altChunk iframe should appear in the rendered output under any
    // option combination — the switch case returns null now.
    assert(
        typeof isSafeHyperlinkHref === 'function',
        '11b: isSafeHyperlinkHref should be exported',
    );
    // Unsafe schemes
    const unsafe = [
        'javascript:alert(1)',
        'JAVASCRIPT:alert(1)',
        '  javascript:alert(1)  ',
        'data:text/html,<script>alert(1)</script>',
        'vbscript:msgbox(1)',
        'file:///etc/passwd',
        'blob:http://attacker/foo',
    ];
    for (const u of unsafe) {
        assert(!isSafeHyperlinkHref(u), `11c: should reject "${u}"`);
    }
    // Safe schemes + relatives + anchor
    const safe = [
        'https://example.com/',
        'http://example.com/',
        'mailto:a@example.com',
        'tel:+1234567890',
        '#anchor',
        '',
        null,
        'relative/path.html',
        '../up/one.html',
    ];
    for (const s of safe) {
        assert(isSafeHyperlinkHref(s), `11d: should accept "${s}"`);
    }
}

// ── 12. Security: CSS sanitizers and prototype-pollution guards ───────────
// Findings 3, 4, 6 from SECURITY_REVIEW.md — the helpers are pure functions
// so we poke them directly with malicious inputs.
{
    // Color sanitizer: allowlist hex / rgb*/hsl*; everything else → null.
    assert(sanitizeCssColor('ff0000') === '#ff0000', '12a: bare 6-hex accepted');
    assert(sanitizeCssColor('#fff') === '#fff', '12b: #3-hex accepted');
    assert(sanitizeCssColor('rgb(1,2,3)') === 'rgb(1,2,3)', '12c: rgb() accepted');
    assert(sanitizeCssColor('red; display: block') === null, '12d: CSS break-out rejected');
    assert(sanitizeCssColor('ff0000}a{') === null, '12e: hex with trailing braces rejected');
    assert(sanitizeCssColor('<img>') === null, '12f: non-color string rejected');
    assert(sanitizeCssColor('"></style><script>') === null, '12g: HTML injection string rejected');
    assert(sanitizeCssColor(null) === null, '12h: null rejected');

    // Font family sanitizer: quotes wrap, backslashes/quotes stripped.
    assert(sanitizeFontFamily('Arial') === "'Arial'", '12i: plain family quoted');
    const evilFont = sanitizeFontFamily('Arial"; } * { background: red } a {');
    assert(
        !evilFont.includes('}') && !evilFont.includes('{') && !evilFont.includes('"'),
        `12j: font-family break-out chars stripped (got ${evilFont})`,
    );
    assert(sanitizeFontFamily('') === 'sans-serif', '12k: empty falls back to sans-serif');
    assert(sanitizeFontFamily(null) === 'sans-serif', '12l: null falls back to sans-serif');

    // Safe CSS identifier check: alnum + underscore only.
    assert(isSafeCssIdent('accent1') === true, '12m: accent1 is safe');
    assert(isSafeCssIdent('foo_bar') === true, '12n: underscores ok');
    assert(isSafeCssIdent('foo-bar') === false, '12o: dash rejected');
    assert(isSafeCssIdent('foo bar') === false, '12p: space rejected');
    assert(isSafeCssIdent('__proto__') === true, '12q: __proto__ is a safe CSS ident (but filtered elsewhere)');
    assert(isSafeCssIdent('}@import url(x);') === false, '12r: injection rejected');

    // CSS string escape: backslashes doubled, quotes backslash-escaped.
    assert(escapeCssStringContent('a"b') === 'a\\"b', '12s: double quotes escaped');
    assert(escapeCssStringContent('a\\b') === 'a\\\\b', '12t: backslashes escaped');
    assert(escapeCssStringContent('plain') === 'plain', '12u: plain passthrough');

    // keyBy / mergeDeep prototype pollution guards.
    const before = {}.polluted;
    const keyed = keyBy([{ id: '__proto__', polluted: true }, { id: 'ok', polluted: false }], x => x.id);
    assert(keyed['__proto__'] === undefined, '12v: keyBy rejects __proto__ key');
    assert(keyed['ok']?.polluted === false, '12w: keyBy still stores safe keys');
    assert(({}).polluted === before, '12x: Object.prototype not polluted by keyBy');

    const before2 = {}.pollutedViaMerge;
    const merged = mergeDeep({}, JSON.parse('{"__proto__": {"pollutedViaMerge": true}}'));
    assert(({}).pollutedViaMerge === before2, '12y: Object.prototype not polluted by mergeDeep');
    // The guarded key should not be copied.
    assert(!('pollutedViaMerge' in merged), '12z: mergeDeep drops __proto__ key');

    // constructor / prototype are also filtered.
    const keyed2 = keyBy([{ id: 'constructor', marker: true }], x => x.id);
    assert(keyed2['constructor'] === undefined, '12aa: keyBy rejects constructor key');
}

// ── report ─────────────────────────────────────────────────────────────────
console.log('--- track-changes harness ---');
for (const w of warnings) console.log(`  · ${w}`);
if (failures.length) {
    console.error(`\n${failures.length} FAILURE(S):`);
    for (const f of failures) console.error(`  ✗ ${f}`);
    process.exit(1);
} else {
    console.log(`\n✓ all ${12} scenarios passed`);
}
