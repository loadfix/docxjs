// Post-render visual pagination.
//
// The renderer (`src/html-renderer.ts`) emits one `<section>` per logical
// page-break run. A section's CSS `min-height` is set from the DOCX sectPr
// pageSize, but its *actual* rendered height can be many multiples of that
// when the content doesn't fit: the section visually spans N pages as one
// tall continuous block. Thumbnails already handle this via clone + clip
// (see `paginateSection` in `src/thumbnails.ts`), but the main view does not.
//
// This module walks the rendered sections in the main view and, for any
// section whose height overflows its page-sized minHeight by more than a
// small slack, slices it into multiple page-shaped sibling sections. The
// split is made at child boundaries inside the section's `<article>` — we
// never split *inside* a child element. The original section keeps its
// headers/footers; later pages do not (v1 limitation — see TODO below).
//
// The function takes an optional `measureFn` so it can be driven in jsdom
// (where getBoundingClientRect / getComputedStyle return zero) by a stub.
//
// Security: DOCX-derived strings are never interpolated into CSS or
// innerHTML. We use cloneNode(false) to copy the section shell (its
// class, id, style attribute), then setAttribute / appendChild to move
// child nodes. The same pattern thumbnails.ts uses.

export interface Measurement {
    width: number;
    height: number;
    minHeight: number;
}

export type MeasureFn = (el: HTMLElement) => Measurement;

const defaultMeasure: MeasureFn = (el) => {
    const win = el.ownerDocument.defaultView;
    const cs = win?.getComputedStyle(el);
    const rect = el.getBoundingClientRect();
    const width = (cs ? parseFloat(cs.width) : 0) || rect.width || 0;
    const height = (cs ? parseFloat(cs.height) : 0) || rect.height || 0;
    const minHeight = cs ? parseFloat(cs.minHeight) || 0 : 0;
    return { width, height, minHeight };
};

// Marker so we can identify sections we injected, mostly for tests and
// debugging. Consumers shouldn't rely on it.
const VISUAL_PAGE_MARKER = 'data-docxjs-visual-page';

/**
 * Walk every top-level `<section>` emitted by the renderer under
 * `bodyContainer` and split those whose rendered height exceeds the page
 * height (min-height from pageSize) by more than a small slack factor.
 *
 * For each overflowing section we locate its content `<article>` and walk
 * children, keeping a running offset; when the next child would push past
 * the page height we move it (and all following children) into a fresh
 * `<article>` inside a cloned page-shaped sibling section. Repeat until
 * every resulting section fits.
 *
 * Returns the number of *new* sections inserted (useful for tests).
 */
export function applyVisualPageBreaks(
    bodyContainer: HTMLElement,
    options: { className?: string; slack?: number } = {},
    measureFn: MeasureFn = defaultMeasure,
): number {
    const className = options.className ?? 'docx';
    // Allow a little slack — rendered content tends to have sub-pixel
    // fractions over its min-height even when it "fits". 1.1× is the ratio
    // used elsewhere in the codebase for the same kind of guard.
    const slack = options.slack ?? 1.1;

    const sections = Array.from(
        bodyContainer.querySelectorAll<HTMLElement>(`section.${className}`),
    );

    let inserted = 0;
    for (const section of sections) {
        // Only process sections the renderer produced; skip ones we may have
        // already injected on a previous pass if the caller re-runs.
        if (section.hasAttribute(VISUAL_PAGE_MARKER)) continue;
        const subPages = splitSection(section, measureFn, slack);
        if (subPages.length > 1) {
            inserted += subPages.length - 1;
            // After the content split, footnotes still live on the original
            // section (they're children of <section>, not <article>). Move
            // each <li> to whichever sub-page cites its footnote id (via
            // matching `data-footnote-id` on <sup> and <li>), so footnotes
            // appear on the page that references them.
            redistributeFootnotes(subPages);
        }
    }
    return inserted;
}

function splitSection(section: HTMLElement, measureFn: MeasureFn, slack: number): HTMLElement[] {
    const { height, minHeight } = measureFn(section);
    const pageHeight = minHeight > 0 ? minHeight : 0;

    // No measurable page height (jsdom, or ignoreHeight) or the content
    // already fits within slack — nothing to do.
    if (pageHeight <= 0) return [section];
    if (height <= pageHeight * slack) return [section];

    // Find the single content `<article>` — that's what `createSectionContent`
    // in html-renderer.ts emits for each section. If there isn't one we bail;
    // anything else (headers, footers, footnote containers) we leave alone.
    //
    // NOTE: v1 keeps headers/footers on the first sub-page only. Later
    // sub-pages get a fresh empty article in a cloned section shell — no
    // header/footer is duplicated. Known limitation: a real Word print
    // would repeat the header/footer on every page. Tracked as out of
    // scope for this change.
    const article = section.querySelector<HTMLElement>(':scope > article');
    if (!article) return [section];

    // Walk the children, measuring cumulative offset. When we'd exceed the
    // page height, start a new section from that child.
    //
    // TODO (v1 limitation): if a single child is taller than the page we
    // leave it where it is and accept overflow on that sub-page. Splitting
    // inside a child (e.g. mid-paragraph or mid-table) is out of scope.
    const children = Array.from(article.children) as HTMLElement[];
    if (children.length === 0) return [section];

    const articleTopOffset = offsetWithinSection(article, section, measureFn);

    const subPages: HTMLElement[] = [section];
    let currentArticle = article;
    let currentTop = articleTopOffset; // offset of currentArticle's top within its section
    let runningHeight = 0;              // height consumed so far inside currentArticle
    let currentSection = section;

    for (let i = 0; i < children.length; i++) {
        const child = children[i];
        const { height: ch } = measureFn(child);

        // If this child on its own would break to a new page (and we've
        // already placed something on the current page), spin up a new
        // section starting with this child.
        const roomLeft = pageHeight - currentTop - runningHeight;
        const willOverflow = ch > roomLeft;

        if (willOverflow && runningHeight > 0) {
            // Close out the current page and start a new one with `child`.
            const newSection = cloneSectionShell(currentSection);
            const newArticle = cloneArticleShell(currentArticle);
            newSection.appendChild(newArticle);

            // Insert right after the current section.
            currentSection.parentNode!.insertBefore(newSection, currentSection.nextSibling);
            subPages.push(newSection);

            // Move `child` and all following children into the new article.
            for (let j = i; j < children.length; j++) {
                newArticle.appendChild(children[j]);
            }

            // Continue walking in the new section.
            currentSection = newSection;
            currentArticle = newArticle;
            currentTop = 0; // fresh section, article starts at top (no header)
            runningHeight = ch;

            // children[i] is now inside newArticle, already accounted for.
            continue;
        }

        runningHeight += ch;
    }
    return subPages;
}

// After splitting, the footnote `<ol>` emitted by the renderer still lives on
// the original section (subPages[0]) because it's a child of `<section>`, not
// of the split `<article>`. Each `<li>` carries `data-footnote-id` from
// renderNotes and each body `<sup>` carries the same from
// renderFootnoteReference — move each `<li>` to the sub-page whose body cites
// its id.
//
// Matching by id (not by `<sup>` position) avoids two bugs: (1) duplicate
// citations of the same footnote id — common in academic writing — would
// previously double-count and steal footnotes from later sub-pages; (2) stray
// `<sup>` elements that aren't footnote refs (e.g. superscript formatting)
// would inflate the count.
//
// Security: `data-footnote-id` values originate from DOCX (elem.id on
// footnote refs / notes), i.e. attacker-controlled per CLAUDE.md. We never
// interpolate them into a CSS selector. Instead we iterate the wildcard
// `[data-footnote-id]` selector and compare values via `dataset.footnoteId`
// in JS. (If a future change wants the selector form, wrap the id with
// `CSS.escape` before interpolation.)
function redistributeFootnotes(subPages: HTMLElement[]): void {
    const original = subPages[0];
    const originalOls = Array.from(
        original.querySelectorAll<HTMLOListElement>(':scope > ol'),
    );
    if (originalOls.length === 0) return;

    // Collect the set of unique footnote ids cited in each sub-page's
    // article. Deduping is important: academic writing cites the same
    // footnote id multiple times and we want one `<li>` per id, not N.
    const refsBySubPage: Set<string>[] = subPages.map((page) => {
        const ids = new Set<string>();
        const sups = page.querySelectorAll<HTMLElement>('article [data-footnote-id]');
        for (const sup of Array.from(sups)) {
            const id = sup.dataset.footnoteId;
            if (id) ids.add(id);
        }
        return ids;
    });

    for (const originalOl of originalOls) {
        // Skip endnote lists if we can detect them. The renderer emits
        // footnotes and endnotes as bare `<ol>` siblings; the endnote list
        // comes last and sits on the final visual page only (see
        // renderSections in html-renderer.ts). If the first `<li>`'s id
        // prefix indicates an endnote, leave the list alone.
        const firstLi = originalOl.querySelector<HTMLLIElement>(':scope > li');
        if (firstLi && firstLi.id && /^docx-endnote/i.test(firstLi.id)) {
            continue;
        }

        const lis = Array.from(originalOl.children) as HTMLElement[];
        // Lazy-created per-sub-page `<ol>` receiving `<li>`s. subPages[0]
        // reuses the original `<ol>` so its attribute shape is preserved.
        const targetOls = new Map<HTMLElement, HTMLOListElement>();
        targetOls.set(original, originalOl);

        for (const li of lis) {
            const id = li.dataset.footnoteId;
            if (!id) continue; // No identity → can't match; leave in sink.

            // Find the first sub-page whose article cites this id.
            let ownerIdx = -1;
            for (let p = 0; p < subPages.length; p++) {
                if (refsBySubPage[p].has(id)) {
                    ownerIdx = p;
                    break;
                }
            }

            // No reference found → leave the li in the original `<ol>` as a
            // safe sink so no footnote content is lost. ownerIdx === 0
            // means the first sub-page (i.e. the original) already owns it;
            // the `<li>` is already in originalOl, no move required.
            if (ownerIdx <= 0) continue;

            const owner = subPages[ownerIdx];
            let ol = targetOls.get(owner);
            if (!ol) {
                ol = originalOl.cloneNode(false) as HTMLOListElement;
                ol.removeAttribute('id');
                owner.appendChild(ol);
                targetOls.set(owner, ol);
            }
            ol.appendChild(li);
        }

        // If the original `<ol>` was fully drained, remove it so an empty
        // list isn't rendered on the first sub-page.
        if (originalOl.children.length === 0) {
            originalOl.remove();
        }

        // Set `start` on each per-sub-page `<ol>` so the rendered numbering
        // is continuous across pages. Without this, each `<ol>` would begin
        // at 1 because the browser's ordered-list counter is per-list —
        // so a doc with 15 footnotes split across 8 pages would render the
        // trailing lists as "1.", "1." 2.", "1." 2." 3." 4.", … instead of
        // "1.", "2." 3.", "4." 5." 6." 7.", … Walk sub-pages in document
        // order and accumulate the number of <li>s seen so far.
        let cumulative = 0;
        for (const page of subPages) {
            const ol = targetOls.get(page);
            if (!ol) continue;
            if (cumulative > 0) {
                ol.setAttribute('start', String(cumulative + 1));
            }
            cumulative += ol.children.length;
        }
    }
}

// Compute `child`'s offsetTop relative to `ancestor`, using measurement.
// We prefer offsetTop when available (HTMLElement.offsetTop is the common
// browser path), falling back to bounding-rect arithmetic. In jsdom both
// typically return 0 and callers will already have short-circuited on
// pageHeight <= 0 upstream — but we keep the calculation defensive.
function offsetWithinSection(child: HTMLElement, ancestor: HTMLElement, measureFn: MeasureFn): number {
    // getBoundingClientRect-based diff is the most robust cross-env path.
    const cRect = child.getBoundingClientRect();
    const aRect = ancestor.getBoundingClientRect();
    const delta = cRect.top - aRect.top;
    if (Number.isFinite(delta) && delta >= 0) return delta;
    // Last-resort fallback: assume article sits flush at section top.
    void measureFn;
    return 0;
}

// Make a new empty `<section>` that mirrors the shape of `source`:
// same tag, same attributes (class, style, id → but we strip id so it
// doesn't collide), but zero children. We use cloneNode(false) so the
// browser handles attribute copying without invoking the HTML parser on
// any DOCX-derived string. No innerHTML or template literals touch DOCX
// data here.
function cloneSectionShell(source: HTMLElement): HTMLElement {
    const shell = source.cloneNode(false) as HTMLElement;
    // An id on the original section is fine; duplicating it is not.
    shell.removeAttribute('id');
    shell.setAttribute(VISUAL_PAGE_MARKER, '');
    return shell;
}

function cloneArticleShell(source: HTMLElement): HTMLElement {
    const shell = source.cloneNode(false) as HTMLElement;
    shell.removeAttribute('id');
    return shell;
}
