[![npm version](https://badge.fury.io/js/docx-preview.svg)](https://www.npmjs.com/package/docx-preview)
[![Support Ukraine](https://img.shields.io/badge/Support-Ukraine-blue?style=flat&logo=adguard)](https://war.ukraine.ua/)

# docxjs
Docx rendering library.

This repository is a fork of [docxjs](https://github.com/VolodymyrBaydalka/docxjs) by Volodymyr Baydalka. It builds on their original work with additional features — visual page breaks, a comments sidebar, track-changes rendering, repeated headers/footers on split pages, and various security and parsing fixes. Credit for the foundational library goes to the original author.

Goal
----
Goal of this project is to render/convert DOCX document into HTML document with keeping HTML semantic as much as possible.
That means library is limited by HTML capabilities (for example Google Docs renders *.docx document on canvas as an image).

Installation
-----
```
npm install docx-preview
```

Usage
-----
```html
<!--lib uses jszip-->
<script src="https://unpkg.com/jszip/dist/jszip.min.js"></script>
<script src="docx-preview.min.js"></script>
<script>
    var docData = <document Blob>;

    docx.renderAsync(docData, document.getElementById("container"))
        .then(x => console.log("docx: finished"));
</script>
<body>
    ...
    <div id="container"></div>
    ...
</body>
```
API
---
```ts
// renders document into specified element
renderAsync(
    document: Blob | ArrayBuffer | Uint8Array, // any type supported by JSZip.loadAsync
    bodyContainer: HTMLElement, // element to render document content
    styleContainer: HTMLElement, // element to render document styles, numberings, fonts. If null, bodyContainer is used.
    options: {
        className: string = "docx",                    // class name/prefix for default and document style classes
        inWrapper: boolean = true,                     // render a wrapper around document content
        hideWrapperOnPrint: boolean = false,           // disable wrapper styles on print
        ignoreWidth: boolean = false,                  // disable rendering width of page
        ignoreHeight: boolean = false,                 // disable rendering height of page
        ignoreFonts: boolean = false,                  // disable fonts rendering
        breakPages: boolean = true,                    // enable page breaking on DOCX page-break elements
        ignoreLastRenderedPageBreak: boolean = true,   // disable page breaking on lastRenderedPageBreak elements
        experimental: boolean = false,                 // enable experimental features (tab stops calculation)
        trimXmlDeclaration: boolean = true,            // strip XML declaration before parsing
        useBase64URL: boolean = false,                 // convert images/fonts to base64 URLs instead of object URLs
        renderHeaders: boolean = true,                 // enable headers rendering
        renderFooters: boolean = true,                 // enable footers rendering
        renderFootnotes: boolean = true,               // enable footnotes rendering
        renderEndnotes: boolean = true,                // enable endnotes rendering
        renderChanges: boolean = false,                // legacy switch for track-changes; prefer `changes.show` below
        renderComments: boolean = false,               // enable comments rendering
        experimentalPageBreaks: boolean = false,       // after rendering, split sections whose content overflows
                                                       // the page-sized min-height into multiple page-shaped siblings.
                                                       // Repeats headers/footers and splits oversized tables at row
                                                       // boundaries. Off by default to preserve current behaviour.
        comments: {
            sidebar: boolean = true,                   // render comments in a right-margin sidebar
            highlight: boolean = true,                 // highlight the anchor text via CSS Highlight API
            layout: 'anchored' | 'packed' = 'anchored' // 'anchored' aligns each card to its anchor; 'packed' stacks flush
        },
        changes: {
            show: boolean = false,                     // master switch for track-changes rendering
            showInsertions: boolean = true,
            showDeletions: boolean = true,
            showMoves: boolean = true,
            showFormatting: boolean = true,
            colorByAuthor: boolean = true,             // color each author's revisions distinctly
            changeBar: boolean = true,                 // render a marginal bar next to changed paragraphs
            legend: boolean = true,                    // inject a legend listing authors
            sidebarCards: boolean = true               // anchor a card per revision in the right margin
        },
        debug: boolean = false
    }): Promise<WordDocument>

// Render per-page thumbnails of an already-rendered document into a sibling container.
renderThumbnails(
    mainContainer: HTMLElement,      // the container renderAsync wrote into
    thumbnailContainer: HTMLElement, // where thumbnails should be placed
    options?: ThumbnailsOptions
): ThumbnailsHandle

/// ==== experimental / internal API ===
// this API can be used to modify a document before rendering
// renderAsync = parseAsync + renderDocument

// parse document and return internal document object
parseAsync(
    document: Blob | ArrayBuffer | Uint8Array,
    options: Options
): Promise<WordDocument>

// render internal document object and return list of nodes
renderDocument(
    wordDocument: WordDocument,
    options: Options
): Promise<Node[]>

// Drive the visual-page-break pass directly (normally called by renderAsync when
// experimentalPageBreaks is true). Useful for consumers that assemble the DOM
// themselves.
applyVisualPageBreaks(
    bodyContainer: HTMLElement,
    options?: { className?: string; slack?: number },
    measureFn?: MeasureFn
): number
```

Thumbnails
------
`renderThumbnails` is a first-class export that renders per-page thumbnails of an already-rendered document into a separate container. Each thumbnail is a clipped clone of the matching page section.

Table of contents
------
Table-of-contents rendering is not yet supported. Word builds TOCs via TOC fields, and field evaluation is not implemented (see http://officeopenxml.com/WPtableOfContents.php).

Breaks
------
The library breaks pages when:
- a manual page break `<w:br w:type="page"/>` is inserted
- an application page break `<w:lastRenderedPageBreak/>` is inserted — Word emits these (set `ignoreLastRenderedPageBreak: false` to honour them)
- page settings change between paragraphs — e.g. portrait to landscape

In addition, when `experimentalPageBreaks: true`, the library performs visual pagination after rendering: any section whose rendered height exceeds its page-sized `min-height` is split into multiple page-shaped sibling sections. Headers/footers are cloned onto every sub-page, and oversized tables are split at row boundaries (preserving `colgroup` and repeating `thead`). Mid-paragraph splitting is not yet implemented — a paragraph taller than a page will overflow its sub-page.

By default `ignoreLastRenderedPageBreak` is `true`. Set it to `false` to break on `<w:lastRenderedPageBreak/>` points.

Status and stability
------
The public surface (`renderAsync`, `parseAsync`, `renderDocument`, `renderThumbnails`) is stable. Newer features behind `experimental*` flags (`experimentalPageBreaks`, revision sidebar cards, etc.) may still change in shape as they settle.

Contributing
------
This fork commits `dist/` alongside source changes so consumers can pull from git directly. If you open a PR, rebuild `dist/` (via `npm run build`) before committing so the bundled output stays in sync with `src/`.
