// Library-level thumbnail API.
//
// Given the container that `renderAsync` / `renderDocument` wrote into and a
// second container the caller owns, build a scaled-down visual preview per
// page with a numeric label, and wire up click-to-scroll + active-page
// tracking via IntersectionObserver.
//
// The renderer emits one `<section>` per logical page break run, but each
// section can overflow across many Word-sized pages when the content is
// taller than the section's page height. We detect this by comparing each
// section's rendered height against its CSS `min-height` (set from the DOCX
// sectPr pageSize). For each visual page we clone the section and use
// `clip-path` + a negative translate to show only the relevant vertical
// slice.
//
// Public entry point: `renderThumbnails(mainContainer, thumbnailContainer, options?)`.

export interface ThumbnailsOptions {
    /** Thumbnail width in px. Default 120. */
    width?: number;
    /** Show a page number label below each thumbnail. Default true. */
    showPageNumbers?: boolean;
    /** Class applied to the currently-active thumbnail. Default `${className}-thumbnail-active`. */
    activeClassName?: string;
    /** Base class name, matches `Options.className` on the renderer. Default `docx`. */
    className?: string;
}

export interface ThumbnailsHandle {
    dispose: () => void;
}

const STYLE_MARKER = 'data-docxjs-thumbnails';
// Marker set by the page-break splitter (src/page-break.ts) on every
// injected sibling section when `experimentalPageBreaks: true`. If any
// section in the main container carries it, we know the splitter has
// already produced one section per visual page — thumbnails must then
// skip their own sub-pagination and emit one thumbnail per section.
const VISUAL_PAGE_MARKER = 'data-docxjs-visual-page';

// One visual page: the section it belongs to, its position within that
// section's flow, and the scroll-target element used when the user clicks.
// When the section fits in a single page `scrollTarget === section`; for
// multi-page sections we insert invisible anchors inside the section so
// scrolling targets the exact page boundary.
interface VisualPage {
    section: HTMLElement;
    scrollTarget: HTMLElement;
    // Pixel offset inside the section where this page starts.
    topOffset: number;
    // Page dimensions in pixels (matches the section's CSS page size).
    pageWidth: number;
    pageHeight: number;
}

function findScrollingAncestor(el: HTMLElement | null): HTMLElement | null {
    let cur: HTMLElement | null = el?.parentElement ?? null;
    while (cur) {
        const cs = cur.ownerDocument.defaultView?.getComputedStyle(cur);
        if (cs) {
            const oy = cs.overflowY;
            if ((oy === 'auto' || oy === 'scroll') && cur.scrollHeight > cur.clientHeight) {
                return cur;
            }
        }
        cur = cur.parentElement;
    }
    return null;
}

function ensureStyle(doc: Document, className: string, activeClassName: string): void {
    const head = doc.head;
    if (!head) return;
    if (head.querySelector(`style[${STYLE_MARKER}]`)) return;

    const style = doc.createElement('style');
    style.setAttribute(STYLE_MARKER, '');
    style.textContent = `
.${className}-thumbnail {
    display: flex;
    flex-direction: column;
    align-items: center;
    margin: 0.5rem auto;
    cursor: pointer;
    outline: none;
}
.${className}-thumbnail:focus-visible .${className}-thumbnail-preview {
    box-shadow: 0 0 0 2px #4a90e2, 0 0 10px rgba(0, 0, 0, 0.5);
}
.${className}-thumbnail-preview {
    overflow: hidden;
    background: white;
    box-shadow: 0 0 10px rgba(0, 0, 0, 0.5);
    box-sizing: content-box;
    border: 2px solid transparent;
    position: relative;
}
.${className}-thumbnail-label {
    font-size: 0.75rem;
    color: white;
    margin-top: 0.25rem;
    text-align: center;
    line-height: 1.2;
}
.${activeClassName} .${className}-thumbnail-preview {
    border-color: #4a90e2;
}
`;
    head.appendChild(style);
}

// Read an element's CSS dimensions in pixels. Falls back to
// getBoundingClientRect when computed styles yield 0 (e.g. element not laid
// out yet).
function measure(el: HTMLElement, win: Window | null): { width: number; height: number; minHeight: number } {
    const cs = win?.getComputedStyle(el);
    const rect = el.getBoundingClientRect();
    const width = (cs ? parseFloat(cs.width) : 0) || rect.width || 0;
    const height = (cs ? parseFloat(cs.height) : 0) || rect.height || 0;
    const minHeight = cs ? parseFloat(cs.minHeight) || 0 : 0;
    return { width, height, minHeight };
}

// Single-page short-circuit. Used when the page-break splitter has
// already run and guaranteed that each section represents exactly one
// visual page — sub-pagination would double-count.
function singlePage(section: HTMLElement, win: Window | null): VisualPage[] {
    const { width, height, minHeight } = measure(section, win);
    const pageHeight = minHeight > 0 ? minHeight : height;
    return [{
        section, scrollTarget: section,
        topOffset: 0, pageWidth: width, pageHeight,
    }];
}

// Splits a rendered section into its visual pages. Returns at least one
// entry even if the measurements are zero (jsdom case); callers should
// tolerate degenerate geometry.
function paginateSection(section: HTMLElement, win: Window | null): VisualPage[] {
    const { width, height, minHeight } = measure(section, win);
    // `min-height` is set by the renderer from the DOCX pageSize; treat it
    // as the canonical page height. Fall back to the rendered height if the
    // renderer has been configured with ignoreHeight.
    const pageHeight = minHeight > 0 ? minHeight : height;
    const pageWidth = width;

    if (pageHeight <= 0 || height <= 0) {
        // No layout info — fall back to treating the whole section as one
        // "page". This also matches the jsdom harness path.
        return [{
            section, scrollTarget: section,
            topOffset: 0, pageWidth, pageHeight,
        }];
    }

    const pageCount = Math.max(1, Math.ceil(height / pageHeight));

    // Single-page section: just the whole thing.
    if (pageCount === 1) {
        return [{
            section, scrollTarget: section,
            topOffset: 0, pageWidth, pageHeight,
        }];
    }

    // Multi-page section: attach a page-sized anchor block per page inside
    // the section. Each anchor is absolutely positioned, page-height tall,
    // and page-width wide, so IntersectionObserver sees it as a real
    // rectangle (not a zero-size point) and picks the anchor that dominates
    // the centre band of the viewport. The anchors are invisible and don't
    // intercept pointer events, so the real content underneath stays
    // interactive.
    const cs = win?.getComputedStyle(section);
    if (cs && cs.position === 'static') {
        section.style.position = 'relative';
    }
    const pages: VisualPage[] = [];
    for (let i = 0; i < pageCount; i++) {
        let anchor = section.querySelector<HTMLElement>(`[data-docxjs-page-anchor="${i}"]`);
        if (!anchor) {
            anchor = section.ownerDocument.createElement('div');
            anchor.setAttribute('data-docxjs-page-anchor', String(i));
            anchor.setAttribute('aria-hidden', 'true');
            anchor.style.cssText = [
                'position:absolute',
                `top:${i * pageHeight}px`,
                'left:0',
                `width:${pageWidth}px`,
                `height:${pageHeight}px`,
                'pointer-events:none',
                'visibility:hidden',
            ].join(';');
            section.appendChild(anchor);
        }
        pages.push({
            section, scrollTarget: anchor,
            topOffset: i * pageHeight,
            pageWidth, pageHeight,
        });
    }
    return pages;
}

export function renderThumbnails(
    mainContainer: HTMLElement,
    thumbnailContainer: HTMLElement,
    options?: ThumbnailsOptions,
): ThumbnailsHandle {
    const width = options?.width ?? 120;
    const showPageNumbers = options?.showPageNumbers ?? true;
    const className = options?.className ?? 'docx';
    const activeClassName = options?.activeClassName ?? `${className}-thumbnail-active`;

    const doc = thumbnailContainer.ownerDocument;
    const win = doc.defaultView;

    ensureStyle(mainContainer.ownerDocument, className, activeClassName);

    // Replace, don't append. Idempotent on re-run.
    thumbnailContainer.innerHTML = '';

    const sections = Array.from(
        mainContainer.querySelectorAll<HTMLElement>(`section.${className}`),
    );

    // If the page-break splitter (src/page-break.ts, gated by
    // `experimentalPageBreaks: true`) has already run, every section in the
    // container is exactly one visual page — injected siblings carry the
    // `data-docxjs-visual-page` attribute, and the original (first) section
    // is also one page by construction. Short-circuit to one thumbnail per
    // section instead of re-paginating, which would otherwise double-count.
    const splitterRan = mainContainer.querySelector(
        `section[${VISUAL_PAGE_MARKER}]`,
    ) !== null;

    // Expand each section into its visual pages.
    const pages: VisualPage[] = [];
    for (const section of sections) {
        const sectionPages = splitterRan
            ? singlePage(section, win)
            : paginateSection(section, win);
        for (const p of sectionPages) {
            pages.push(p);
        }
    }

    const pairs: Array<{ scrollTarget: HTMLElement; thumb: HTMLElement }> = [];

    for (let i = 0; i < pages.length; i++) {
        const { section, scrollTarget, topOffset, pageWidth, pageHeight } = pages[i];
        const pageNum = i + 1;

        const thumb = doc.createElement('div');
        thumb.className = `${className}-thumbnail`;
        thumb.setAttribute('role', 'button');
        thumb.setAttribute('tabindex', '0');
        thumb.setAttribute('aria-label', `Go to page ${pageNum}`);
        thumb.dataset.page = String(pageNum);

        const preview = doc.createElement('div');
        preview.className = `${className}-thumbnail-preview`;

        // Clone the rendered section once per thumbnail. cloneNode(true)
        // preserves structure and attributes without invoking the HTML
        // parser on any DOCX-derived strings.
        const clone = section.cloneNode(true) as HTMLElement;
        clone.setAttribute('aria-hidden', 'true');
        clone.removeAttribute('id');
        clone.style.boxShadow = 'none';
        clone.style.margin = '0';
        clone.style.flexShrink = '0';

        let scale = 1;
        let previewHeight = 0;
        if (pageWidth > 0) {
            scale = width / pageWidth;
            previewHeight = pageHeight * scale;
        }

        preview.style.width = `${width}px`;
        if (previewHeight > 0) {
            preview.style.height = `${previewHeight}px`;
        }

        // Compose the transform so only this page's slice is visible. The
        // clone is scaled and then translated up by its page offset (scaled),
        // so the preview box shows rows [topOffset..topOffset+pageHeight]
        // from the section.
        const translate = -topOffset * scale;
        clone.style.transform = `translateY(${translate}px) scale(${scale})`;
        clone.style.transformOrigin = '0 0';

        preview.appendChild(clone);
        thumb.appendChild(preview);

        if (showPageNumbers) {
            const label = doc.createElement('div');
            label.className = `${className}-thumbnail-label`;
            label.textContent = String(pageNum);
            thumb.appendChild(label);
        }

        const goTo = () => {
            scrollTarget.scrollIntoView({ behavior: 'smooth', block: 'start' });
        };
        thumb.addEventListener('click', goTo);
        thumb.addEventListener('keydown', (ev: Event) => {
            const ke = ev as KeyboardEvent;
            if (ke.key === 'Enter' || ke.key === ' ') {
                ke.preventDefault();
                goTo();
            }
        });

        thumbnailContainer.appendChild(thumb);
        pairs.push({ scrollTarget, thumb });
    }

    // Active-page tracking. IntersectionObserver watches each scroll target
    // and we pick whichever one is most "inside" the centre strip.
    let observer: IntersectionObserver | null = null;
    const IO =
        (win as any)?.IntersectionObserver ??
        (globalThis as any).IntersectionObserver;

    if (IO && pairs.length > 0) {
        const scrollRoot = findScrollingAncestor(mainContainer);
        const visibility = new Map<HTMLElement, number>();

        observer = new IO((entries: IntersectionObserverEntry[]) => {
            for (const entry of entries) {
                visibility.set(entry.target as HTMLElement, entry.intersectionRatio);
            }
            let bestIdx = -1;
            let bestRatio = -1;
            for (let i = 0; i < pairs.length; i++) {
                const r = visibility.get(pairs[i].scrollTarget) ?? 0;
                if (r > bestRatio) {
                    bestRatio = r;
                    bestIdx = i;
                }
            }
            for (let i = 0; i < pairs.length; i++) {
                pairs[i].thumb.classList.toggle(activeClassName, i === bestIdx && bestRatio > 0);
            }
        }, {
            root: scrollRoot,
            rootMargin: '-45% 0px -45% 0px',
            threshold: [0, 0.01, 0.5, 1],
        }) as IntersectionObserver;

        for (const { scrollTarget } of pairs) observer.observe(scrollTarget);
    }

    return {
        dispose() {
            if (observer) {
                observer.disconnect();
                observer = null;
            }
            thumbnailContainer.innerHTML = '';
            // Remove any page-boundary anchors we inserted into sections.
            for (const section of sections) {
                section.querySelectorAll('[data-docxjs-page-anchor]').forEach(n => n.remove());
            }
        },
    };
}
