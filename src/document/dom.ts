export enum DomType {
    Document = "document",
    Paragraph = "paragraph",
    Run = "run",
    Break = "break",
    NoBreakHyphen = "noBreakHyphen",
    Table = "table",
    Row = "row",
    Cell = "cell",
    Hyperlink = "hyperlink",
    SmartTag = "smartTag",
    Drawing = "drawing",
    Image = "image",
    DrawingShape = "drawingShape",
    DrawingGroup = "drawingGroup",
    Text = "text",
    Tab = "tab",
    Symbol = "symbol",
    BookmarkStart = "bookmarkStart",
    BookmarkEnd = "bookmarkEnd",
    Footer = "footer",
    Header = "header",
    FootnoteReference = "footnoteReference", 
	EndnoteReference = "endnoteReference",
    Footnote = "footnote",
    Endnote = "endnote",
    SimpleField = "simpleField",
    ComplexField = "complexField",
    Instruction = "instruction",
	VmlPicture = "vmlPicture",
	MmlMath = "mmlMath",
	MmlMathParagraph = "mmlMathParagraph",
	MmlFraction = "mmlFraction",
	MmlFunction = "mmlFunction",
	MmlFunctionName = "mmlFunctionName",
	MmlNumerator = "mmlNumerator",
	MmlDenominator = "mmlDenominator",
	MmlRadical = "mmlRadical",
	MmlBase = "mmlBase",
	MmlDegree = "mmlDegree",
	MmlSuperscript = "mmlSuperscript",
	MmlSubscript = "mmlSubscript",
	MmlPreSubSuper = "mmlPreSubSuper",
	MmlSubArgument = "mmlSubArgument",
	MmlSuperArgument = "mmlSuperArgument",
	MmlNary = "mmlNary",
	MmlDelimiter = "mmlDelimiter",
	MmlRun = "mmlRun",
	MmlEquationArray = "mmlEquationArray",
	MmlLimit = "mmlLimit",
	MmlLimitLower = "mmlLimitLower",
	MmlMatrix = "mmlMatrix",
	MmlMatrixRow = "mmlMatrixRow",
	MmlBox = "mmlBox",
	MmlBar = "mmlBar",
	MmlGroupChar = "mmlGroupChar",
	MmlAccent = "mmlAccent",
	MmlBorderBox = "mmlBorderBox",
	MmlSubSuperscript = "mmlSubSuperscript",
	MmlPhantom = "mmlPhantom",
	MmlGroup = "mmlGroup",
	VmlElement = "vmlElement",
	Inserted = "inserted",
	Deleted = "deleted",
	DeletedText = "deletedText",
	MoveFrom = "moveFrom",
	MoveTo = "moveTo",
	Comment = "comment",
	CommentReference = "commentReference",
	CommentRangeStart = "commentRangeStart",
	CommentRangeEnd = "commentRangeEnd",
    AltChunk = "altChunk",
    Sdt = "sdt",
    // East-Asian / bidi typography (see document-parser.ts parseRuby,
    // parseFitText, parseBidiOverride). Ruby wraps rubyBase and rt children;
    // the renderer emits native HTML <ruby><rt>. FitText wraps a run with a
    // target width in twips. BidiOverride wraps content in <bdo dir="…">.
    Ruby = "ruby",
    RubyBase = "rubyBase",
    RubyText = "rubyText",
    FitText = "fitText",
    BidiOverride = "bidiOverride",
    Chart = "chart",
    // Modern 2013+ chart parts (sunburst / waterfall / funnel / treemap /
    // histogram / pareto / box-whisker). Rendered as a labelled
    // placeholder rather than a real SVG chart. See
    // src/charts/chartex-part.ts and the renderChartEx branch of
    // html-renderer.ts.
    ChartEx = "chartEx",
    // SmartArt reference that could not be rendered as its Fallback
    // drawing (either no mc:Fallback existed or its content was
    // unrecognised). Rendered as a labelled placeholder with the
    // layout URN carried in a data attribute. See
    // parseSmartArtReference in document-parser.ts.
    SmartArt = "smartArt"
}

// Structured Document Tag (content control). Parsed from w:sdt when
// w:sdtPr contains a w:alias or w:tag — otherwise parseSdt unwraps
// directly to sdtContent children (unless a typed form control is
// detected below, which also forces a wrapper). See parseSdt in
// document-parser.ts and the DomType.Sdt branch of renderElement in
// html-renderer.ts.
export interface WmlSdt extends OpenXmlElement {
    sdtAlias?: string;
    sdtTag?: string;
    sdtControl?: SdtControl;
}

// Typed form-control metadata extracted from w:sdtPr. All DOCX-derived
// string fields (displayText, value, format, fullDate) are attacker
// controlled — they must only reach the DOM via setAttribute /
// textContent, never innerHTML or className interpolation.
export type SdtControl =
    | SdtCheckboxControl
    | SdtDropdownControl
    | SdtDateControl
    | SdtPictureControl
    | SdtGalleryControl;

export interface SdtCheckboxControl {
    type: "checkbox";
    checked: boolean;
    checkedChar?: number;
    uncheckedChar?: number;
}

export interface SdtDropdownItem {
    displayText: string;
    value: string;
}

export interface SdtDropdownControl {
    // Shared shape for w:dropDownList and w:comboBox — the only rendering
    // difference Word draws between them (editable combo vs strict list)
    // doesn't apply in a read-only HTML view, so both emit <select>.
    type: "dropdown";
    items: SdtDropdownItem[];
}

export interface SdtDateControl {
    type: "date";
    format?: string;
    fullDate?: string;
}

export interface SdtPictureControl {
    type: "picture";
}

export interface SdtGalleryControl {
    type: "gallery";
}

export interface Revision {
    id?: string;
    author?: string;
    date?: string;
}

// Captured from w:rPrChange / w:pPrChange. Phase 2 (#4) only records the
// revision metadata and a small, bounded summary of what changed so the
// renderer can tell the reader "bold, font-size"; computing the full
// previous property set is Phase 3+.
export interface FormattingRevision extends Revision {
    // Short summary of changed property keys, e.g. ['bold', 'fontSize'].
    // Bounded length keeps the rendered title attribute manageable.
    changedProps?: string[];
}

export interface OpenXmlElement {
    type: DomType;
    children?: OpenXmlElement[];
    cssStyle?: Record<string, string>;
    props?: Record<string, any>;

	styleName?: string; //style name
	className?: string; //class mods

    // Populated for DomType.Inserted / DomType.Deleted / MoveFrom / MoveTo
    // and for paragraph-mark revisions (w:pPr/w:rPr/w:ins|w:del).
    // See Track Changes Phase 1 (#3).
    revision?: Revision;
    // 'inserted' or 'deleted' — only set on paragraphs whose paragraph mark
    // itself was inserted/deleted. Rendered as a pilcrow in Phase 2 (#4).
    paragraphMarkRevisionKind?: 'inserted' | 'deleted';
    // Populated on runs (w:rPrChange) and paragraphs (w:pPrChange).
    // See Track Changes Phase 2 (#4).
    formattingRevision?: FormattingRevision;

    parent?: OpenXmlElement;
}

export abstract class OpenXmlElementBase implements OpenXmlElement {
    type: DomType;
    children?: OpenXmlElement[] = [];
    cssStyle?: Record<string, string> = {};
    props?: Record<string, any>;

    className?: string;
    styleName?: string;

    parent?: OpenXmlElement;
}

export interface WmlHyperlink extends OpenXmlElement {
	id?: string;
    anchor?: string;
    tooltip?: string;
    targetFrame?: string;
}

export interface WmlAltChunk extends OpenXmlElement {
	id?: string;
}

export interface WmlSmartTag extends OpenXmlElement {
	uri?: string;
    element?: string;
}

// w:ruby — East-Asian phonetic guide (furigana). Children are RubyBase and
// RubyText sub-elements; rubyPr is captured as a bag of numeric/bool props.
// See parseRuby in document-parser.ts and renderRuby in html-renderer.ts.
export interface WmlRuby extends OpenXmlElement {
    rubyPr?: {
        // w:hps — half-point font size for the ruby text
        hps?: number;
        // w:hpsBaseText — half-point font size for the base characters
        hpsBaseText?: number;
        // w:hpsRaise — amount to raise the ruby above its base
        hpsRaise?: number;
        // w:lid — language identifier (e.g. "ja-JP")
        lid?: string;
        // w:rubyAlign — center | distributeLetter | distributeSpace | left | right | rightVertical
        rubyAlign?: string;
    };
}

// w:fitText — "fit text" character / run wrapper. val is target width in twips,
// id groups consecutive fitText runs into a single stretched unit.
export interface WmlFitText extends OpenXmlElement {
    // Target width in twips (1/20 pt). Numeric by construction in the parser.
    width?: number;
    id?: string;
}

// w:bdo — explicit bidi override ("ltr" | "rtl"). The `dir` field is
// allowlisted in the parser before it ever reaches the DOM.
export interface WmlBidiOverride extends OpenXmlElement {
    dir?: "ltr" | "rtl";
}

export interface WmlNoteReference extends OpenXmlElement {
    id: string;
}

export interface WmlBreak extends OpenXmlElement{
    break: "page" | "lastRenderedPageBreak" | "textWrapping";
}

export interface WmlText extends OpenXmlElement{
    text: string;
}

export interface WmlSymbol extends OpenXmlElement {
    font: string;
    char: number;
}

export interface WmlTable extends OpenXmlElement {
    columns?: WmlTableColumn[];
    cellStyle?: Record<string, string>;

	colBandSize?: number;
	rowBandSize?: number;
}

export interface WmlTableRow extends OpenXmlElement {
	isHeader?: boolean;
    gridBefore?: number;
    gridAfter?: number;
    // w:cantSplit — when true, the visual-pagination splitter must not
    // break a page inside this row. See splitTableAtRowBoundary in
    // page-break.ts.
    cantSplit?: boolean;
    // Set when w:trPr contains a w:ins or w:del element (Track Changes
    // Phase 3, #5). The row itself is inserted or deleted, distinct from
    // revisions to its content.
    rowRevisionKind?: 'inserted' | 'deleted';
}

export interface WmlTableCell extends OpenXmlElement {
	verticalMerge?: 'restart' | 'continue' | string;
    span?: number;
}

export interface IDomImage extends OpenXmlElement {
    src: string;
    srcRect: number[];
    rotation: number;
    // wp:docPr/@descr, pic:cNvPr/@descr, or a:blip/@descr. Accessibility
    // alt text for screen readers. Empty string when absent/decorative.
    altText?: string;
}

export interface WmlTableColumn {
    width?: string;
}

export interface IDomNumbering {
    id: string;
    level: number;
    start: number;
    pStyleName: string;
    pStyle: Record<string, string>;
    rStyle: Record<string, string>;
    levelText?: string;
    suff: string;
    format?: string;
    bullet?: NumberingPicBullet;
    // When set, every %N placeholder in the level text must render as
    // arabic regardless of that level's own numFmt (w:isLgl).
    isLgl?: boolean;
    // w:lvlRestart — the ancestor level whose occurrence resets this
    // counter. Omitted / -1 means "default" (restart at parent).
    restart?: number;
    // w:lvlJc — left | right | center | start | end.
    justification?: string;
}

export interface NumberingPicBullet {
    id: number;
    src: string;
    style?: string;
}
