import { WordDocument } from './word-document';
import { DocumentParser } from './document-parser';
import { HtmlRenderer } from './html-renderer';
import { h } from './html';

export { renderThumbnails } from './thumbnails';
export type { ThumbnailsOptions, ThumbnailsHandle } from './thumbnails';

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
    renderAltChunks: boolean;
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
    renderAltChunks: true,
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

export function parseAsync(data: Blob | any, userOptions?: Partial<Options>): Promise<any>  {
    const ops = mergeOptions(userOptions);
    return WordDocument.load(data, new DocumentParser(ops), ops);
}

export async function renderDocument(document: any, userOptions?: Partial<Options>): Promise<any> {
    const ops = mergeOptions(userOptions);
    const renderer = new HtmlRenderer();
    return await renderer.render(document, ops);
}

export async function renderAsync(data: Blob | any, bodyContainer: HTMLElement, styleContainer?: HTMLElement, userOptions?: Partial<Options>): Promise<any> {
	const doc = await parseAsync(data, userOptions);
	const nodes = await renderDocument(doc, userOptions);

    styleContainer ??= bodyContainer;
    styleContainer.innerHTML = "";
    bodyContainer.innerHTML = "";

    for (let n of nodes) {
        const c = n.nodeName === "STYLE" ? styleContainer : bodyContainer;
        c.appendChild(n);
    }

    return doc;
}
