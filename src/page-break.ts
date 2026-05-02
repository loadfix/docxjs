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
        inserted += splitSection(section, measureFn, slack);
    }
    return inserted;
}

function splitSection(section: HTMLElement, measureFn: MeasureFn, slack: number): number {
    const { height, minHeight } = measureFn(section);
    const pageHeight = minHeight > 0 ? minHeight : 0;

    // No measurable page height (jsdom, or ignoreHeight) or the content
    // already fits within slack — nothing to do.
    if (pageHeight <= 0) return 0;
    if (height <= pageHeight * slack) return 0;

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
    if (!article) return 0;

    // Walk the children, measuring cumulative offset. When we'd exceed the
    // page height, start a new section from that child.
    //
    // TODO (v1 limitation): if a single child is taller than the page we
    // leave it where it is and accept overflow on that sub-page. Splitting
    // inside a child (e.g. mid-paragraph or mid-table) is out of scope.
    const children = Array.from(article.children) as HTMLElement[];
    if (children.length === 0) return 0;

    const articleTopOffset = offsetWithinSection(article, section, measureFn);

    let inserted = 0;
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
            inserted++;

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
    return inserted;
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
