import { WordDocument } from './word-document';
import { DocumentParser } from './document-parser';
import { HtmlRenderer } from './html-renderer';
import { h } from './html';
import { applyVisualPageBreaks } from './page-break';

export { renderThumbnails } from './thumbnails';
export type { ThumbnailsOptions, ThumbnailsHandle } from './thumbnails';
// Exported for the Node jsdom harness so it can exercise the null-val
// guard for upstream issue #196 directly, without needing a crafted DOCX.
export { classNameOfCnfStyle } from './document-parser';
// Exported so consumers and tests can drive the pagination pass directly;
// renderAsync calls this internally when `experimentalPageBreaks` is on.
export { applyVisualPageBreaks } from './page-break';
export type { MeasureFn, Measurement } from './page-break';

// Security helpers re-exported so they can be unit-tested and (optionally)
// reused by embedding code. Pure functions; see SECURITY_REVIEW.md for
// context and the behaviours each one enforces.
export { isSafeHyperlinkHref } from './html-renderer';
export {
    sanitizeCssColor,
    sanitizeFontFamily,
    isSafeCssIdent,
    escapeCssStringContent,
    keyBy,
    mergeDeep,
} from './utils';
// VML colour sanitiser — strips Word's "[####]" theme-index suffix before
// delegating to sanitizeCssColor. Exported for unit testing; see #171.
export { sanitizeVmlColor } from './vml/vml';
// Field instruction tokenizer — exported for harness testing.
export { parseFieldInstruction } from './fields/instruction';
export type { ParsedFieldInstruction } from './fields/instruction';
// Squarified treemap layout — exported so the jsdom harness can assert
// on rectangle positions without needing a chartEx fixture. See
// Bruls et al., 'Squarified Treemaps' (2000).
export { layoutTreemap } from './charts/render';
export type { TreemapLayoutRect } from './charts/render';

export interface CommentsOptions {
    sidebar?: boolean;
    highlight?: boolean;
    /**
     * Sidebar card layout mode.
     * - 'anchored' (default): each card aligns vertically with its anchor
     *   text, pushing down on collision, re-flowing as the document scrolls.
     *   Matches Word's "All Markup" convention.
     * - 'packed': cards stack flush in document order with no gaps. Simpler
     *   and more compact; no scroll re-flow.
     */
    layout?: 'anchored' | 'packed';
}

export interface ChangesOptions {
    show?: boolean;
    showInsertions?: boolean;
    showDeletions?: boolean;
    showMoves?: boolean;
    showFormatting?: boolean;
    colorByAuthor?: boolean;
    changeBar?: boolean;
    legend?: boolean;
    sidebarCards?: boolean;
}

export interface Options {
    inWrapper: boolean;
    hideWrapperOnPrint: boolean;
    ignoreWidth: boolean;
    ignoreHeight: boolean;
    ignoreFonts: boolean;
    breakPages: boolean;
    debug: boolean;
    experimental: boolean;
    className: string;
    trimXmlDeclaration: boolean;
    renderHeaders: boolean;
    renderFooters: boolean;
    renderFootnotes: boolean;
	renderEndnotes: boolean;
    ignoreLastRenderedPageBreak: boolean;
	useBase64URL: boolean;
	renderChanges: boolean;
    renderComments: boolean;
    /**
     * Experimental: when true, after the renderer emits its sections, walk
     * each one and split any that overflow their page-sized `min-height`
     * into multiple page-shaped sibling sections. This matches the visual
     * pagination that `renderThumbnails` already produces on the thumbnail
     * side (see `src/thumbnails.ts`). Default `false` to preserve current
     * behaviour for existing consumers. See #22.
     */
    experimentalPageBreaks?: boolean;
    comments: CommentsOptions;
    changes: ChangesOptions;
    h: typeof h;
}

export const defaultOptions: Options = {
    ignoreHeight: false,
    ignoreWidth: false,
    ignoreFonts: false,
    breakPages: true,
    debug: false,
    experimental: false,
    className: "docx",
    inWrapper: true,
    hideWrapperOnPrint: false,
    trimXmlDeclaration: true,
    ignoreLastRenderedPageBreak: true,
    renderHeaders: true,
    renderFooters: true,
    renderFootnotes: true,
	renderEndnotes: true,
	useBase64URL: false,
	renderChanges: false,
    renderComments: false,
    experimentalPageBreaks: false,
    comments: {
        sidebar: true,
        highlight: true,
        layout: 'anchored',
    },
    changes: {
        show: false,
        showInsertions: true,
        showDeletions: true,
        showMoves: true,
        showFormatting: true,
        colorByAuthor: true,
        changeBar: true,
        legend: true,
        sidebarCards: true,
    },
    h: h
};

function mergeOptions(userOptions?: Partial<Options>): Options {
    const ops = { ...defaultOptions, ...userOptions };
    // `renderChanges: true` is the legacy switch. If a caller sets it without
    // also passing `changes.show`, honour the legacy intent.
    if (userOptions?.renderChanges && userOptions?.changes?.show === undefined) {
        ops.changes = { ...defaultOptions.changes, ...userOptions.changes, show: true };
    }
    return ops;
}

export function parseAsync(data: Blob | any, userOptions?: Partial<Options>): Promise<WordDocument>  {
    const ops = mergeOptions(userOptions);
    return WordDocument.load(data, new DocumentParser(ops), ops);
}

export async function renderDocument(document: any, userOptions?: Partial<Options>): Promise<Node[]> {
    const ops = mergeOptions(userOptions);
    const renderer = new HtmlRenderer();
    return await renderer.render(document, ops);
}

export async function renderAsync(data: Blob | any, bodyContainer: HTMLElement, styleContainer?: HTMLElement, userOptions?: Partial<Options>): Promise<WordDocument> {
	const doc = await parseAsync(data, userOptions);
	const nodes = await renderDocument(doc, userOptions);

    styleContainer ??= bodyContainer;
    styleContainer.innerHTML = "";
    bodyContainer.innerHTML = "";

    for (let n of nodes) {
        const c = n.nodeName === "STYLE" ? styleContainer : bodyContainer;
        c.appendChild(n);
    }

    // Visual pagination must run after nodes are attached — the split logic
    // relies on layout-driven measurements (getBoundingClientRect / computed
    // styles) that return zero for detached elements. See #22.
    const ops = mergeOptions(userOptions);
    if (ops.experimentalPageBreaks) {
        applyVisualPageBreaks(bodyContainer, { className: ops.className });
    }

    return doc;
}
