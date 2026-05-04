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
    Sdt = "sdt"
}

// Structured Document Tag (content control). Parsed from w:sdt when
// w:sdtPr contains a w:alias or w:tag — otherwise parseSdt unwraps
// directly to sdtContent children. See parseSdt in document-parser.ts
// and the DomType.Sdt branch of renderElement in html-renderer.ts.
export interface WmlSdt extends OpenXmlElement {
    sdtAlias?: string;
    sdtTag?: string;
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
