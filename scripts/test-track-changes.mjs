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
    applyVisualPageBreaks,
    isSafeHyperlinkHref, sanitizeCssColor, sanitizeFontFamily,
    isSafeCssIdent, escapeCssStringContent, keyBy, mergeDeep,
    sanitizeVmlColor,
    classNameOfCnfStyle,
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

// ── 13. Experimental visual page breaks (#22) ─────────────────────────────
// Two flavours: (a) integration through renderAsync-equivalent path is a
// no-op in jsdom because there's no layout, so just assert nothing throws
// and sections remain present; (b) drive applyVisualPageBreaks directly
// with an injected measure function so we can force a split even without
// a real browser layout.
{
    assert(
        typeof applyVisualPageBreaks === 'function',
        '13a: docx.applyVisualPageBreaks export should be a function',
    );

    // (a) jsdom path — no-op, but should not throw and should leave the
    // existing sections intact.
    const { container: inert } = await renderFixture('text', { experimentalPageBreaks: true });
    const inertSections = inert.querySelectorAll('section.docx');
    assert(inertSections.length > 0, '13b: text fixture should still render sections when experimentalPageBreaks is on');
    const inertVisualPages = inert.querySelectorAll('section.docx[data-docxjs-visual-page]');
    assert(
        inertVisualPages.length === 0,
        `13c: no visual-page splits expected in jsdom without layout (got ${inertVisualPages.length})`,
    );

    // (b) measure-injection path — build a fake section that overflows and
    // drive the pagination helper directly.
    const root = document.createElement('div');
    const section = document.createElement('section');
    section.className = 'docx';
    const article = document.createElement('article');
    for (let i = 0; i < 6; i++) {
        const p = document.createElement('p');
        p.textContent = `paragraph ${i}`;
        article.appendChild(p);
    }
    section.appendChild(article);
    root.appendChild(section);
    document.body.appendChild(root);

    // Fake measurements: section reports a height of 6× page height, each
    // paragraph reports exactly page/3 (so 3 paragraphs fit per page).
    const PAGE = 300;
    const PARA = 100;
    const heights = new Map();
    heights.set(section, { width: 800, height: PAGE * 6, minHeight: PAGE });
    for (const p of article.children) heights.set(p, { width: 800, height: PARA, minHeight: 0 });
    // Article itself fills the section content area.
    heights.set(article, { width: 800, height: PAGE * 6, minHeight: 0 });

    const inserted = applyVisualPageBreaks(root, { className: 'docx' }, (el) => {
        return heights.get(el) ?? { width: 0, height: 0, minHeight: 0 };
    });

    const sectionsAfter = root.querySelectorAll('section.docx');
    note(`13·: forced-split produced ${sectionsAfter.length} section(s), inserted=${inserted}`);
    assert(inserted > 0, '13d: forced overflow should insert at least one new section');
    assert(sectionsAfter.length >= 2, `13e: expected >= 2 sections after split (got ${sectionsAfter.length})`);
    // Injected siblings should carry the marker attribute.
    const marked = root.querySelectorAll('section[data-docxjs-visual-page]');
    assert(marked.length === inserted, `13f: marker attribute should be present on ${inserted} injected section(s), got ${marked.length}`);
    // The original section should retain its class (cloneNode(false) copies attrs).
    for (const s of sectionsAfter) {
        assert(s.classList.contains('docx'), '13g: every split section should keep the docx class');
    }

    root.remove();
}

// ── 14. Thumbnails short-circuit when page-break splitter has run ─────────
// When src/page-break.ts splits a section into multiple visual pages, each
// injected sibling is a separate `<section class="docx">` carrying
// `data-docxjs-visual-page`. In that state, renderThumbnails must emit one
// thumbnail per section (no re-pagination) — otherwise thumbnails would
// double-count once per splitter page and once per paginateSection pass.
{
    // Baseline: no splitter marker → one thumbnail per section (jsdom has no
    // layout so paginateSection short-circuits to one-per-section anyway).
    const { container: baselineMain } = await renderFixture('text');
    document.body.appendChild(baselineMain);
    const baselineThumbs = document.createElement('div');
    document.body.appendChild(baselineThumbs);
    const baselineSections = baselineMain.querySelectorAll('section.docx').length;
    renderThumbnails(baselineMain, baselineThumbs);
    const baselineThumbCount = baselineThumbs.querySelectorAll('.docx-thumbnail').length;
    note(`14·: baseline text fixture → ${baselineSections} section(s), ${baselineThumbCount} thumbnail(s)`);
    assert(
        baselineThumbCount === baselineSections,
        `14a: baseline expected ${baselineSections} thumbnails (one per section), got ${baselineThumbCount}`,
    );
    baselineMain.remove();
    baselineThumbs.remove();

    // Simulated post-splitter state: two sections, second carries the
    // visual-page marker. renderThumbnails must short-circuit and emit
    // exactly 2 thumbnails even if the jsdom fallback in paginateSection
    // would otherwise still hand back one page per section (identical
    // count here, but the short-circuit path runs — verified in-browser).
    const synthMain = document.createElement('div');
    const s1 = document.createElement('section');
    s1.className = 'docx';
    s1.appendChild(document.createElement('article'));
    const s2 = document.createElement('section');
    s2.className = 'docx';
    s2.setAttribute('data-docxjs-visual-page', '');
    s2.appendChild(document.createElement('article'));
    synthMain.appendChild(s1);
    synthMain.appendChild(s2);
    document.body.appendChild(synthMain);
    const synthThumbs = document.createElement('div');
    document.body.appendChild(synthThumbs);

    renderThumbnails(synthMain, synthThumbs);
    const synthThumbCount = synthThumbs.querySelectorAll('.docx-thumbnail').length;
    note(`14·: synthetic post-splitter (2 sections, 1 marked) → ${synthThumbCount} thumbnail(s)`);
    assert(
        synthThumbCount === 2,
        `14b: post-splitter state should emit exactly 2 thumbnails (no re-pagination), got ${synthThumbCount}`,
    );
    // No page-anchor elements should have been injected: singlePage short-
    // circuit skips paginateSection entirely.
    assert(
        synthMain.querySelectorAll('[data-docxjs-page-anchor]').length === 0,
        '14c: short-circuit path should not inject page-anchor elements',
    );

    synthMain.remove();
    synthThumbs.remove();
}

// ── 15. VML colour sanitiser strips Word's "[####]" theme-index suffix (#171) ─
// No shipped fixture carries VML shapes with the index suffix, so we drive
// the sanitiser directly. It delegates to sanitizeCssColor after stripping
// the trailing ` [1234]` — values that fail the allowlist after stripping
// should still return null. Also sanity-check the two sink paths by parsing
// a synthetic <v:shape> element and confirming the rendered attribute value
// does not carry the raw "[3204]" substring.
{
    assert(
        typeof sanitizeVmlColor === 'function',
        '15a: docx.sanitizeVmlColor should be exported',
    );
    // Happy path: suffix stripped, colour normalised through sanitizeCssColor.
    assert(sanitizeVmlColor('#4472c4 [3204]') === '#4472c4', '15b: hex with [####] suffix stripped');
    assert(sanitizeVmlColor('4472c4 [3204]') === '#4472c4', '15c: bare hex with suffix stripped and #-prefixed');
    assert(sanitizeVmlColor('#4472c4') === '#4472c4', '15d: plain hex unchanged');
    assert(sanitizeVmlColor('#4472c4  [1]  ') === '#4472c4', '15e: whitespace around suffix tolerated');
    // Rejects: post-strip value must still pass sanitizeCssColor allowlist.
    assert(sanitizeVmlColor('red; display: block [1]') === null, '15f: break-out payload rejected even with suffix');
    assert(sanitizeVmlColor('javascript:alert(1) [2]') === null, '15g: bogus value rejected');
    assert(sanitizeVmlColor(null) === null, '15h: null returns null');
    assert(sanitizeVmlColor(undefined) === null, '15i: undefined returns null');
    assert(sanitizeVmlColor('') === null, '15j: empty string returns null');
    // Only a trailing [####] group is stripped — interior "[x]" is left alone
    // and then fails the colour allowlist.
    assert(sanitizeVmlColor('[1] #4472c4') === null, '15k: leading "[n]" not stripped');

    // Also verify the rendered output of a synthetic v:shape never carries
    // the raw "[3204]" substring. We shove the sanitised value through an
    // attribute and serialise the element, mimicking what the VML renderer
    // does further downstream.
    const synthetic = document.createElement('div');
    const svgNs = 'http://www.w3.org/2000/svg';
    const shape = document.createElementNS(svgNs, 'rect');
    const fill = sanitizeVmlColor('#4472c4 [3204]');
    if (fill) shape.setAttribute('fill', fill);
    const stroke = sanitizeVmlColor('#ED7D31 [3205]');
    if (stroke) shape.setAttribute('stroke', stroke);
    synthetic.appendChild(shape);
    const html = synthetic.outerHTML;
    assert(
        !html.includes('[3204]') && !html.includes('[3205]'),
        `15l: rendered output must not carry "[####]" suffix (got ${html})`,
    );
    assert(
        html.includes('#4472c4') && html.includes('#ED7D31'),
        `15m: rendered output should carry the sanitised colour values (got ${html})`,
    );
}

// ── 16. w14:paraId exposed as data-para-id on rendered paragraphs (#128) ──
// The parser reads w14:paraId via getAttributeNS; the renderer copies it to
// dataset.paraId when present. Use a fixture whose paragraphs carry w14:paraId
// attributes (underlines — 5 paraIds). The value must match SAFE_PARA_ID.
{
    const SAFE_PARA_ID = /^[A-Za-z0-9_-]+$/;
    const { container } = await renderFixture('underlines');
    const withParaId = container.querySelectorAll('p[data-para-id]');
    note(`16·: underlines fixture produced ${withParaId.length} paragraph(s) with data-para-id`);
    assert(
        withParaId.length > 0,
        '16a: at least one <p> should carry data-para-id when the DOCX declares w14:paraId',
    );
    for (const p of withParaId) {
        const v = p.getAttribute('data-para-id');
        assert(
            typeof v === 'string' && v.length > 0,
            '16b: data-para-id value should be a non-empty string',
        );
        assert(
            SAFE_PARA_ID.test(v),
            `16c: data-para-id="${v}" should match SAFE_PARA_ID /^[A-Za-z0-9_-]+$/`,
        );
    }
    // The dataset.paraId camelCase round-trips through the attribute.
    const first = withParaId[0];
    if (first) {
        assert(
            first.dataset.paraId === first.getAttribute('data-para-id'),
            '16d: dataset.paraId should equal the data-para-id attribute',
        );
    }
}

// ── 17. classNameOfCnfStyle null-val guard (upstream #196) ────────────────
// A `<w:cnfStyle/>` element without a `w:val` attribute used to crash the
// parser ("Cannot read properties of null (reading '0')") because the code
// indexed into the result of xml.attr(), which is null for a missing attr.
// The fix short-circuits with ''. We exercise the helper directly with a
// synthetic Element — no DOCX fixture needed.
{
    assert(
        typeof classNameOfCnfStyle === 'function',
        '17a: docx.classNameOfCnfStyle export should be a function',
    );

    // Synthesize <w:cnfStyle/> with no val attribute — the crash path.
    const xmlDoc = new dom.window.DOMParser().parseFromString(
        '<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        '<w:cnfStyle/></w:pPr>',
        'application/xml',
    );
    const cnfEl = xmlDoc.getElementsByTagNameNS(
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'cnfStyle',
    ).item(0);
    assert(!!cnfEl, '17b: synthetic cnfStyle element should exist');

    let result, threw = null;
    try {
        result = classNameOfCnfStyle(cnfEl);
    } catch (e) {
        threw = e;
    }
    assert(
        threw === null,
        `17c: classNameOfCnfStyle must not throw on missing val (got ${threw?.message})`,
    );
    assert(
        result === '',
        `17d: missing val should yield '' (got ${JSON.stringify(result)})`,
    );

    // Sanity: still produces classes when val is present.
    const withValDoc = new dom.window.DOMParser().parseFromString(
        '<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        '<w:cnfStyle w:val="100000000000"/></w:pPr>',
        'application/xml',
    );
    const withValEl = withValDoc.getElementsByTagNameNS(
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'cnfStyle',
    ).item(0);
    assert(
        classNameOfCnfStyle(withValEl) === 'first-row',
        "17e: val='100000000000' should map to 'first-row'",
    );
}

// ── 18. Footnote redistribution across visual-page siblings (#32) ─────────
// When applyVisualPageBreaks splits a section, the footnote <ol> (a child of
// <section>, not <article>) used to stay on the original sub-page, so all
// footnotes ended up bunched on the first visual page of the group. The fix
// walks footnote <li>s after the split and moves each one to the sub-page
// whose <article> cites its id. Conservation invariant: the total number of
// footnote <li>s across every sub-page after the split equals the pre-split
// count (modulo unmatched ones that stay on the first sub-page as a sink).
{
    // Build a synthetic section with 6 paragraphs in <article> and a
    // footnote <ol> of 4 items as a sibling of <article>. Each paragraph
    // owns one or two <sup data-footnote-id> references; two references
    // share footnote id "fn-2" (mirrors the ins/del double-sup pattern and
    // academic duplicate-citation shape).
    const root = document.createElement('div');
    const section = document.createElement('section');
    section.className = 'docx';
    const article = document.createElement('article');

    // Paragraph/footnote plan (ids):
    //   p0 → fn-1
    //   p1 → fn-2, fn-2   (duplicate reference to the same footnote)
    //   p2 → fn-3
    //   p3 → (no reference)
    //   p4 → fn-4
    //   p5 → (no reference)
    const plan = [['fn-1'], ['fn-2', 'fn-2'], ['fn-3'], [], ['fn-4'], []];
    for (let i = 0; i < plan.length; i++) {
        const p = document.createElement('p');
        p.textContent = `paragraph ${i} `;
        for (const id of plan[i]) {
            const sup = document.createElement('sup');
            sup.textContent = id.replace(/\D/g, '');
            sup.dataset.footnoteId = id;
            p.appendChild(sup);
        }
        article.appendChild(p);
    }
    section.appendChild(article);

    // Footnote list with 4 items, one per id.
    const footnoteOl = document.createElement('ol');
    for (const id of ['fn-1', 'fn-2', 'fn-3', 'fn-4']) {
        const li = document.createElement('li');
        li.textContent = `footnote ${id} content`;
        li.setAttribute('data-footnote-id', id);
        footnoteOl.appendChild(li);
    }
    section.appendChild(footnoteOl);
    root.appendChild(section);
    document.body.appendChild(root);

    const preSplitFootnoteCount = section.querySelectorAll('ol > li').length;
    assert(
        preSplitFootnoteCount === 4,
        `18a: synthetic setup should have 4 footnotes (got ${preSplitFootnoteCount})`,
    );

    // Force a split: section reports 6× page height, each paragraph reports
    // exactly page/2 so 2 paragraphs fit per page → 3 sub-pages.
    const PAGE = 300;
    const PARA = 150;
    const heights = new Map();
    heights.set(section, { width: 800, height: PAGE * 6, minHeight: PAGE });
    heights.set(article, { width: 800, height: PAGE * 6, minHeight: 0 });
    for (const p of article.children) heights.set(p, { width: 800, height: PARA, minHeight: 0 });

    const inserted = applyVisualPageBreaks(root, { className: 'docx' }, (el) => {
        return heights.get(el) ?? { width: 0, height: 0, minHeight: 0 };
    });
    assert(inserted >= 1, `18b: forced overflow should insert at least one sibling (got ${inserted})`);

    const subPages = Array.from(root.querySelectorAll('section.docx'));
    assert(subPages.length >= 2, `18c: expected >= 2 sub-pages after split (got ${subPages.length})`);

    // Conservation: sum of <li>s across every sub-page equals pre-split count.
    const postSplitCounts = subPages.map((s) => s.querySelectorAll(':scope > ol > li').length);
    const postSplitTotal = postSplitCounts.reduce((a, b) => a + b, 0);
    note(`18·: sub-page footnote counts = ${JSON.stringify(postSplitCounts)} (sum ${postSplitTotal})`);
    assert(
        postSplitTotal === preSplitFootnoteCount,
        `18d: footnote conservation failed — pre=${preSplitFootnoteCount}, post=${postSplitTotal}`,
    );

    // Redistribution: at least two different sub-pages must own footnotes
    // (otherwise the original bunching bug is still present).
    const subPagesWithFootnotes = postSplitCounts.filter((n) => n > 0).length;
    assert(
        subPagesWithFootnotes >= 2,
        `18e: expected footnotes spread across ≥2 sub-pages (got ${subPagesWithFootnotes})`,
    );

    // Correctness: every <li> must land on a sub-page whose article cites
    // its id via a <sup data-footnote-id>.
    for (const id of ['fn-1', 'fn-2', 'fn-3', 'fn-4']) {
        const owner = subPages.find((s) => {
            return !!s.querySelector(':scope > ol > li[data-footnote-id="' + id.replace(/"/g, '\\"') + '"]');
        });
        assert(
            !!owner,
            `18f: footnote ${id} should have an owning sub-page holding its <li>`,
        );
        if (owner) {
            const cites = Array.from(owner.querySelectorAll('article [data-footnote-id]'))
                .some((sup) => sup.dataset.footnoteId === id);
            assert(
                cites,
                `18g: owning sub-page of ${id} must also contain a <sup> citing ${id}`,
            );
        }
    }

    // If the original sub-page ended up with no footnotes, its <ol> should
    // have been removed to avoid a stray empty list.
    if (postSplitCounts[0] === 0) {
        const strayOl = subPages[0].querySelector(':scope > ol');
        assert(
            !strayOl,
            '18h: original sub-page <ol> should be removed when all footnotes moved away',
        );
    }

    root.remove();
}

// ── 19. Per-sub-page footnote id set equality with duplicate refs (#34) ───
// After applyVisualPageBreaks splits a section, each sub-page must satisfy:
//   { ids on <sup data-footnote-id> inside article }
//     == { ids on <li data-footnote-id> inside this sub-page's <ol> }
// The important corner case is duplicate refs: when a sub-page cites the
// same footnote id twice, exactly one <li> should land on it (not two).
// Previously the matching pass was positional, which over-counted duplicate
// refs and stole <li>s from later sub-pages; issue #34.
{
    const root = document.createElement('div');
    const section = document.createElement('section');
    section.className = 'docx';
    const article = document.createElement('article');

    // Plan (one paragraph per row → 6 paragraphs, 2 per sub-page = 3 pages):
    //   p0 cites f-A once
    //   p1 cites f-A again (DUPLICATE on sub-page 0) and f-B once
    //   p2 cites f-C once
    //   p3 cites f-C twice (duplicate on sub-page 1)
    //   p4 cites f-D once
    //   p5 cites f-D once (duplicate on sub-page 2)
    const plan = [['f-A'], ['f-A', 'f-B'], ['f-C'], ['f-C', 'f-C'], ['f-D'], ['f-D']];
    for (let i = 0; i < plan.length; i++) {
        const p = document.createElement('p');
        p.textContent = `p${i} `;
        for (const id of plan[i]) {
            const sup = document.createElement('sup');
            sup.textContent = '*';
            sup.dataset.footnoteId = id;
            p.appendChild(sup);
        }
        article.appendChild(p);
    }
    section.appendChild(article);

    const ol = document.createElement('ol');
    for (const id of ['f-A', 'f-B', 'f-C', 'f-D']) {
        const li = document.createElement('li');
        li.textContent = `footnote ${id}`;
        li.setAttribute('data-footnote-id', id);
        ol.appendChild(li);
    }
    section.appendChild(ol);
    root.appendChild(section);
    document.body.appendChild(root);

    const PAGE = 300;
    const PARA = 150;
    const heights = new Map();
    heights.set(section, { width: 800, height: PAGE * 6, minHeight: PAGE });
    heights.set(article, { width: 800, height: PAGE * 6, minHeight: 0 });
    for (const p of article.children) heights.set(p, { width: 800, height: PARA, minHeight: 0 });

    applyVisualPageBreaks(root, { className: 'docx' }, (el) => {
        return heights.get(el) ?? { width: 0, height: 0, minHeight: 0 };
    });

    const subPages = Array.from(root.querySelectorAll('section.docx'));
    assert(subPages.length >= 3, `19a: expected >= 3 sub-pages (got ${subPages.length})`);

    // Per-sub-page set equality: refs on <sup> == li ids in the sub-page's <ol>.
    const snapshot = subPages.map((s, i) => {
        const refIds = new Set(
            Array.from(s.querySelectorAll('article [data-footnote-id]'))
                .map((el) => el.dataset.footnoteId),
        );
        const liIds = new Set(
            Array.from(s.querySelectorAll(':scope > ol > li[data-footnote-id]'))
                .map((el) => el.getAttribute('data-footnote-id')),
        );
        return { index: i, refIds, liIds };
    });
    note(
        `19·: per-page snapshot = ` +
        JSON.stringify(snapshot.map((s) => ({ i: s.index, refs: [...s.refIds], lis: [...s.liIds] }))),
    );
    for (const { index, refIds, liIds } of snapshot) {
        assert(
            refIds.size === liIds.size && [...refIds].every((id) => liIds.has(id)),
            `19b·${index}: refIds=${JSON.stringify([...refIds])} ` +
            `liIds=${JSON.stringify([...liIds])} — set equality expected`,
        );
    }

    // Duplicate refs must not produce duplicate <li>s on the same sub-page.
    for (const { index, liIds } of snapshot) {
        const liEls = subPages[index].querySelectorAll(':scope > ol > li[data-footnote-id]');
        assert(
            liEls.length === liIds.size,
            `19c·${index}: duplicate <li>s found — count=${liEls.length}, unique=${liIds.size}`,
        );
    }

    // Conservation across the split (4 original footnotes).
    const totalLis = snapshot.reduce((a, s) => a + s.liIds.size, 0);
    assert(totalLis === 4, `19d: total <li> count should be 4 (got ${totalLis})`);

    root.remove();
}

// ── 20. Footnote reference <sup> markers tagged + CSS injected (#26) ──────
// Word renders footnote / endnote reference markers via the
// FootnoteReference / EndnoteReference character styles (~65% of body text).
// Our <sup> was inheriting the browser's default `font-size: smaller`, so
// markers rendered visibly oversized next to body text. The fix tags body
// refs with .docx-footnote-ref / .docx-endnote-ref and injects a CSS rule
// setting font-size: 0.65em so markers match Word's size.
{
    const { container } = await renderFixture('footnote');
    const refs = container.querySelectorAll('sup.docx-footnote-ref');
    note(`20·: footnote fixture produced ${refs.length} body <sup.docx-footnote-ref>`);
    assert(
        refs.length > 0,
        '20a: footnote fixture should render at least one body <sup.docx-footnote-ref>',
    );
    for (const sup of refs) {
        assert(
            sup.tagName.toLowerCase() === 'sup',
            '20b: tagged element should be a <sup>',
        );
    }
    // Default-style injection carries the sizing rule for both footnote and
    // endnote markers. We find it by scanning the <style> nodes rendered by
    // renderDocument (they're appended to the container in this harness).
    const styleTexts = [...container.querySelectorAll('style')].map((s) => s.textContent || '');
    const hasFootnoteRule = styleTexts.some((t) => t.includes('docx-footnote-ref') && t.includes('0.65em'));
    assert(
        hasFootnoteRule,
        '20c: injected <style> should include the .docx-footnote-ref sizing rule (font-size: 0.65em)',
    );
    const hasEndnoteRule = styleTexts.some((t) => t.includes('docx-endnote-ref'));
    assert(
        hasEndnoteRule,
        '20d: injected <style> should include the .docx-endnote-ref sizing rule',
    );
}

// ── 21. Continuous footnote numbering across section/page boundaries ──────
// Word's default is 1..N continuous across the whole document. Previously
// renderSections reset `currentFootnoteIds` at the start of every loop
// iteration and the superscript number was derived from that list's length,
// so any document that splits into multiple pages inside renderSections
// (explicit `<w:sectPr>` breaks OR a `<w:br w:type="page"/>` inside the
// body) would restart numbering at 1 on each page. The fix introduces a
// document-wide `footnoteRefCount` that only resets once per render().
{
    const { container } = await renderFixture('footnote', {});
    const sups = Array.from(container.querySelectorAll('sup[data-footnote-id]'));
    const numbers = sups.map((el) => el.textContent);
    note(`21·: rendered sup numbers = [${numbers.join(',')}] across ${container.querySelectorAll('section.docx').length} section(s)`);
    assert(
        sups.length >= 2,
        `21a: expected >= 2 footnote refs in the footnote fixture (got ${sups.length})`,
    );
    // Monotone, starts at 1, strictly increments by 1.
    for (let i = 0; i < sups.length; i++) {
        assert(
            sups[i].textContent === String(i + 1),
            `21b·${i}: expected sup #${i} to read "${i + 1}" — got "${sups[i].textContent}"`,
        );
    }
    // Extra guard: numbering must not restart, i.e. the string '1' appears
    // exactly once across the whole ref list.
    const ones = numbers.filter((n) => n === '1').length;
    assert(
        ones === 1,
        `21c: expected exactly one "1" in sup numbers (got ${ones}: [${numbers.join(',')}]) — numbering restarted per page`,
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
    console.log(`\n✓ all ${21} scenarios passed`);
}
