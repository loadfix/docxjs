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
            // each <li> to whichever sub-page contains its matching <sup>
            // reference so footnotes appear on the page that cites them.
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
// of the split `<article>`. Each `<li>` corresponds to a footnote whose number
// (1-based index in the `<ol>`) matches the text of some `<sup>` reference
// inside one of the sub-page articles. Move each `<li>` to the sub-page that
// cites it so footnotes appear on the visual page containing their reference.
//
// The original `<ol>` may have no class or inline style (the renderer emits a
// bare `<ol>`) — we still clone its attribute shape with cloneNode(false) so
// if customisation is added later it's preserved on every sub-page. No DOCX
// strings are interpolated; all operations are DOM moves between rendered
// nodes, with reference matching on the rendered numeric text content.
function redistributeFootnotes(subPages: HTMLElement[]): void {
    const original = subPages[0];
    const originalOls = Array.from(
        original.querySelectorAll<HTMLOListElement>(':scope > ol'),
    );
    if (originalOls.length === 0) return;

    // Collect the set of footnote-reference numbers cited in each sub-page.
    // Matching by `<sup>` text content inside the section's `<article>`. The
    // renderer's `renderFootnoteReference` emits `<sup>N</sup>` where N is
    // the 1-based footnote index; the same number can appear in multiple
    // `<sup>`s (e.g. inside an ins/del wrapper) but we only need to know
    // *which* sub-page owns each footnote number.
    const refsBySubPage: Set<number>[] = subPages.map((page) => {
        const nums = new Set<number>();
        const sups = page.querySelectorAll<HTMLElement>('article sup');
        for (const sup of Array.from(sups)) {
            const t = (sup.textContent ?? '').trim();
            if (!/^\d+$/.test(t)) continue;
            nums.add(parseInt(t, 10));
        }
        return nums;
    });

    for (const originalOl of originalOls) {
        // Skip endnote lists if we can detect them. The renderer emits
        // footnotes and endnotes as bare `<ol>` siblings with no class, so
        // there's nothing to discriminate on in today's output. If the
        // element or its items carry an `docx-endnote` id prefix, leave it
        // alone out of caution.
        const firstLi = originalOl.querySelector<HTMLLIElement>(':scope > li');
        if (firstLi && firstLi.id && /^docx-endnote/i.test(firstLi.id)) {
            continue;
        }

        const lis = Array.from(originalOl.children) as HTMLElement[];
        // Lazy-created per-sub-page `<ol>` receiving `<li>`s. subPages[0]
        // reuses the original `<ol>` so its attribute shape is preserved.
        const targetOls = new Map<HTMLElement, HTMLOListElement>();
        targetOls.set(original, originalOl);

        for (let i = 0; i < lis.length; i++) {
            const li = lis[i];
            const footnoteNumber = i + 1; // 1-based, matches <sup> text
            // Find the first sub-page whose article cites this number.
            let ownerIdx = -1;
            for (let p = 0; p < subPages.length; p++) {
                if (refsBySubPage[p].has(footnoteNumber)) {
                    ownerIdx = p;
                    break;
                }
            }

            // No reference found → leave the li in the original `<ol>` as a
            // safe sink so no footnote content is lost.
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
