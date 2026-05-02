import { WordDocument } from './word-document';
import {
	DomType, WmlTable, IDomNumbering,
	WmlHyperlink, IDomImage, OpenXmlElement, WmlTableColumn, WmlTableCell, WmlText, WmlSymbol, WmlBreak, WmlNoteReference,
	WmlSmartTag,
	WmlTableRow
} from './document/dom';
import { Options } from './docx-preview';
import { DocumentElement } from './document/document';
import { WmlParagraph } from './document/paragraph';
import {
	asArray, encloseFontFamily, escapeClassName, escapeCssStringContent,
	isSafeCssIdent, isString, keyBy, mergeDeep, sanitizeCssColor, sanitizeFontFamily,
} from './utils';
import { computePixelToPoint, updateTabStop } from './javascript';
import { FontTablePart } from './font-table/font-table';
import { FooterHeaderReference, SectionProperties } from './document/section';
import { WmlRun } from './document/run';
import { WmlBookmarkStart } from './document/bookmarks';
import { IDomStyle } from './document/style';
import { WmlBaseNote, WmlFootnote } from './notes/elements';
import { ThemePart } from './theme/theme-part';
import { BaseHeaderFooterPart } from './header-footer/parts';
import { Part } from './common/part';
import { VmlElement } from './vml/vml';
import { WmlComment, WmlCommentRangeStart, WmlCommentRangeEnd, WmlCommentReference } from './comments/elements';
import { cx, h, ns } from './html';

// URL schemes safe to emit as the `href` of a rendered hyperlink in a
// read-only document viewer. Anything outside this list (most importantly
// `javascript:`, `data:`, `vbscript:`, `blob:`, `file:`) is rejected — the
// link is rendered as inert plain text instead. Fragment-only (`#foo`) and
// scheme-relative / path-relative URLs are treated as safe because they can't
// carry script. See SECURITY_REVIEW.md #2.
const SAFE_HREF_SCHEMES = new Set(['http:', 'https:', 'mailto:', 'tel:', 'ftp:', 'ftps:']);

/**
 * Returns `true` iff `raw` can be safely emitted as the `href` of an `<a>` in
 * a read-only docx viewer. The allowlist accepts absolute URLs with known
 * safe schemes, fragment-only URLs (`#anchor`), and relative paths. All
 * other schemes — notably `javascript:` — are rejected.
 */
export function isSafeHyperlinkHref(raw: string | null | undefined): boolean {
	if (raw == null) return true; // empty href is inert
	if (typeof raw !== 'string') return false;
	const trimmed = raw.trim();
	if (trimmed === '') return true;
	if (trimmed.startsWith('#')) return true;
	try {
		// Resolve against a synthetic base so relative URLs parse but never
		// inherit the host document's base URI (which could be file:// in
		// embedding apps and pass the scheme allowlist accidentally).
		const parsed = new URL(trimmed, 'http://docxjs.invalid/');
		return SAFE_HREF_SCHEMES.has(parsed.protocol);
	} catch {
		// Non-URL strings that aren't fragments are almost always relative
		// paths like `foo/bar.html`. Treat as safe — no scheme, no sink.
		return !/^[a-z][a-z0-9+.-]*:/i.test(trimmed);
	}
}

// Internal-only — addresses a rendered change element via data-change-kind.
// Not exported; the library is read-only so consumers never see this.
type ChangeKind =
	| 'insertion'
	| 'deletion'
	| 'formatting'
	| 'move'
	| 'paragraphMark'
	| 'rowInsertion'
	| 'rowDeletion';

interface CellPos {
	col: number;
	row: number;
}

interface Section {
	sectProps: SectionProperties;
	elements: OpenXmlElement[];
	pageBreak: boolean;
}

declare const Highlight: any;

type CellVerticalMergeType = Record<number, HTMLTableCellElement>;

export class HtmlRenderer {

	className: string = "docx";
	rootSelector: string;
	document: WordDocument;
	options: Options;
	styleMap: Record<string, IDomStyle> = {};
	currentPart: Part = null;

	tableVerticalMerges: CellVerticalMergeType[] = [];
	currentVerticalMerge: CellVerticalMergeType = null;
	tableCellPositions: CellPos[] = [];
	currentCellPosition: CellPos = null;

	footnoteMap: Record<string, WmlFootnote> = {};
	endnoteMap: Record<string, WmlFootnote> = {};
	currentFootnoteIds: string[];
	currentEndnoteIds: string[] = [];
	usedHederFooterParts: any[] = [];

	defaultTabSize: string;
	currentTabs: any[] = [];

	commentHighlight: any;
	commentMap: Record<string, Range> = {};
	commentAnchorElements: Record<string, HTMLElement[]> = {};
	sidebarContainer: HTMLElement = null;
	sidebarCommentElements: Record<string, HTMLElement> = {};
	// Track-changes revision cards in the sidebar, keyed by change id. Used by
	// the anchored layout pass alongside sidebarCommentElements.
	revisionCardElements: Map<string, HTMLElement> = new Map();

	// Track-changes (#3): per-render author → palette index, plus all rendered
	// change elements so the post-render change-bar pass can walk them.
	changeAuthorIndex: Map<string, number> = new Map();
	changeElements: HTMLElement[] = [];
	// Parallel to changeElements: metadata for each change element so that the
	// sidebar card pass, the delegated accept/reject handler, and the
	// change-bar pass can all read from the same record.
	changeMeta: Array<{
		el: HTMLElement;
		id?: string;
		kind: ChangeKind;
		author?: string;
		date?: string;
		summary: string;
	}> = [];
	// Tracks the two halves of each move so a click on either scrolls to the
	// counterpart. Keyed by move id — note the parser feeds us a DOCX-derived
	// string here, so the renderer must never interpolate it into a CSS
	// selector or innerHTML (we only use dataset and addEventListener).
	moveElements: Map<string, { from?: HTMLElement; to?: HTMLElement }> = new Map();
	static CHANGE_PALETTE_SIZE = 8;

	tasks: Promise<any>[] = [];
	postRenderTasks: any[] = [];
	h = h;

	get useSidebar(): boolean {
		return this.options.renderComments && (this.options.comments?.sidebar !== false);
	}

	get useHighlight(): boolean {
		return this.options.renderComments && (this.options.comments?.highlight !== false);
	}

	get sidebarLayout(): 'anchored' | 'packed' {
		return this.options.comments?.layout === 'packed' ? 'packed' : 'anchored';
	}

	get showChanges(): boolean {
		return !!this.options.changes?.show;
	}

	async render(document: WordDocument, options: Options): Promise<Node[]> {
		this.document = document;
		this.options = options;
		this.className = options.className;
		this.rootSelector = options.inWrapper ? `.${this.className}-wrapper` : ':root';
		this.h = options.h ?? h;
		this.styleMap = null;
		this.tasks = [];
		this.commentAnchorElements = {};
		this.sidebarCommentElements = {};
		this.revisionCardElements = new Map();
		this.sidebarContainer = null;
		this.changeAuthorIndex = new Map();
		this.changeElements = [];
		this.changeMeta = [];
		this.moveElements = new Map();

		if (this.options.renderComments && this.useHighlight && globalThis.Highlight) {
			this.commentHighlight = new Highlight();
		}

		const result: Node[] = [...this.renderDefaultStyle()];

		if (document.themePart) {
			result.push(...this.renderTheme(document.themePart));
		}

		if (document.stylesPart != null) {
			this.styleMap = this.processStyles(document.stylesPart.styles);
			result.push(...this.renderStyles(document.stylesPart.styles));
		}

		if (document.numberingPart) {
			this.prodessNumberings(document.numberingPart.domNumberings);

			result.push(...await this.renderNumbering(document.numberingPart.domNumberings));
			//result.push(...await this.renderNumbering2(document.numberingPart.domNumberings));
		}

		if (document.footnotesPart) {
			this.footnoteMap = keyBy(document.footnotesPart.notes, x => x.id);
		}

		if (document.endnotesPart) {
			this.endnoteMap = keyBy(document.endnotesPart.notes, x => x.id);
		}

		if (document.settingsPart) {
			this.defaultTabSize = document.settingsPart.settings?.defaultTabStop;
		}

		if (!options.ignoreFonts && document.fontTablePart)
			result.push(...await this.renderFontTable(document.fontTablePart));

		var sectionElements = this.renderSections(document.documentPart.body);

		if (this.options.inWrapper) {
			if (this.useSidebar) {
				result.push(this.renderWrapperWithSidebar(sectionElements));
			} else {
				result.push(this.renderWrapper(sectionElements));
			}
		} else {
			result.push(...sectionElements);
		}

		if (this.commentHighlight && this.useHighlight) {
			(CSS as any).highlights.set(`${this.className}-comments`, this.commentHighlight);
		} else {
			// Clear any highlight registered by a previous render of the same
			// renderer instance; required so toggling comments.highlight off and
			// re-rendering actually removes the text highlights.
			(CSS as any).highlights?.delete(`${this.className}-comments`);
		}

		if (this.showChanges) {
			this.finalizeChangesRendering(result);
		}

		this.postRenderTasks.forEach(t => t());

		await Promise.allSettled(this.tasks);

		this.refreshTabStops();

		return result;
	}

	renderTheme(themePart: ThemePart) {
		// All string fields here come from DOCX theme XML; every interpolation
		// point is sanitized before it reaches CSS. Invalid values are dropped
		// rather than coerced to a default, so a malformed theme produces
		// fewer custom properties instead of unsafe ones. See
		// SECURITY_REVIEW.md #4.
		const variables: Record<string, string> = {};
		const fontScheme = themePart.theme?.fontScheme;

		if (fontScheme) {
			if (fontScheme.majorFont?.latinTypeface) {
				variables['--docx-majorHAnsi-font'] = sanitizeFontFamily(fontScheme.majorFont.latinTypeface);
			}

			if (fontScheme.minorFont?.latinTypeface) {
				variables['--docx-minorHAnsi-font'] = sanitizeFontFamily(fontScheme.minorFont.latinTypeface);
			}
		}

		const colorScheme = themePart.theme?.colorScheme;

		if (colorScheme) {
			for (const [k, v] of Object.entries(colorScheme.colors)) {
				// Both key and value are attacker-controlled: key becomes part
				// of the custom-property name, value becomes the color. Bail on
				// anything that isn't a safe identifier / hex color.
				if (!isSafeCssIdent(k)) continue;
				const color = sanitizeCssColor(v);
				if (!color) continue;
				variables[`--docx-${k}-color`] = color;
			}
		}

		const cssText = this.styleToString(`.${this.className}`, variables);
		return [
			this.h({ tagName: "#comment", children: ["docxjs document theme values"] }),
			this.h({ tagName: "style", children: [cssText] })
		];
	}

	async renderFontTable(fontsPart: FontTablePart) {
		const result = [];

		for (let f of fontsPart.fonts) {
			for (let ref of f.embedFontRefs) {
				try{
					const fontData = await this.document.loadFont(ref.id, ref.key);
					const cssValues = {
						'font-family': encloseFontFamily(f.name),
						'src': `url(${fontData})`
					};

					if (ref.type == "bold" || ref.type == "boldItalic") {
						cssValues['font-weight'] = 'bold';
					}

					if (ref.type == "italic" || ref.type == "boldItalic") {
						cssValues['font-style'] = 'italic';
					}

					result.push(this.h({ tagName: "#comment", children: [`docxjs ${f.name} font`] }));
					result.push(this.h({ tagName: "style", children: [this.styleToString(`@font-face`, cssValues)] }));
				} catch(e) {
					if (this.options.debug) console.warn(`Can't load font with id ${ref.id} and key ${ref.key}`);
				}
			}
		}

		return result;
	}

	processStyleName(className: string): string {
		return className ? `${this.className}_${escapeClassName(className)}` : this.className;
	}

	processStyles(styles: IDomStyle[]) {
		const stylesMap = keyBy(styles.filter(x => x.id != null), x => x.id);

		for (const style of styles.filter(x => x.basedOn)) {
			var baseStyle = stylesMap[style.basedOn];

			if (baseStyle) {
				style.paragraphProps = mergeDeep(style.paragraphProps, baseStyle.paragraphProps);
				style.runProps = mergeDeep(style.runProps, baseStyle.runProps);

				for (const baseValues of baseStyle.styles) {
					const styleValues = style.styles.find(x => x.target == baseValues.target);

					if (styleValues) {
						this.copyStyleProperties(baseValues.values, styleValues.values);
					} else {
						style.styles.push({ ...baseValues, values: { ...baseValues.values } });
					}
				}
			}
			else if (this.options.debug)
				console.warn(`Can't find base style ${style.basedOn}`);
		}

		for (let style of styles) {
			style.cssName = this.processStyleName(style.id);
		}

		return stylesMap;
	}

	prodessNumberings(numberings: IDomNumbering[]) {
		for (let num of numberings.filter(n => n.pStyleName)) {
			const style = this.findStyle(num.pStyleName);

			if (style?.paragraphProps?.numbering) {
				style.paragraphProps.numbering.level = num.level;
			}
		}
	}

	processElement(element: OpenXmlElement) {
		if (element.children) {
			for (var e of element.children) {
				e.parent = element;

				if (e.type == DomType.Table) {
					this.processTable(e);
				}
				else {
					this.processElement(e);
				}
			}
		}
	}

	processTable(table: WmlTable) {
		for (var r of table.children) {
			for (var c of r.children) {
				c.cssStyle = this.copyStyleProperties(table.cellStyle, c.cssStyle, [
					"border-left", "border-right", "border-top", "border-bottom",
					"padding-left", "padding-right", "padding-top", "padding-bottom"
				]);

				this.processElement(c);
			}
		}
	}

	copyStyleProperties(input: Record<string, string>, output: Record<string, string>, attrs: string[] = null): Record<string, string> {
		if (!input)
			return output;

		if (output == null) output = {};
		if (attrs == null) attrs = Object.getOwnPropertyNames(input);

		for (var key of attrs) {
			if (input.hasOwnProperty(key) && !output.hasOwnProperty(key))
				output[key] = input[key];
		}

		return output;
	}

	createPageElement(className: string, props: SectionProperties, docStyle: Record<string, any>) {
		const style: Record<string, string> = { ...docStyle };

		if (props) {
			if (props.pageMargins) {
				style.paddingLeft = props.pageMargins.left;
				style.paddingRight = props.pageMargins.right;
				style.paddingTop = props.pageMargins.top;
				style.paddingBottom = props.pageMargins.bottom;
			}

			if (props.pageSize) {
				if (!this.options.ignoreWidth)
					style.width = props.pageSize.width;
				if (!this.options.ignoreHeight)
					style.minHeight = props.pageSize.height;
			}
		}

		return this.h({ tagName: "section", className, style }) as HTMLElement;
	}

	createSectionContent(props: SectionProperties) {
		const style: Record<string, string> = {};

		if (props.columns && props.columns.numberOfColumns) {
			style.columnCount = `${props.columns.numberOfColumns}`;
			style.columnGap = props.columns.space;

			if (props.columns.separator) {
				style.columnRule = "1px solid black";
			}
		}

		return this.h({ tagName: "article", style }) ;
	}	

	renderSections(document: DocumentElement): HTMLElement[] {
		const result = [];

		this.processElement(document);
		const sections = this.splitBySection(document.children, document.props);
		const pages = this.groupByPageBreaks(sections);
		let prevProps = null;

		for (let i = 0, l = pages.length; i < l; i++) {			
			this.currentFootnoteIds = [];

			const section = pages[i][0];
			let props = section.sectProps;
			const pageElement = this.createPageElement(this.className, props, document.cssStyle);

			this.options.renderHeaders && this.renderHeaderFooter(props.headerRefs, props,
				result.length, prevProps != props, pageElement);

			for (const sect of pages[i]) {
				var contentElement = this.createSectionContent(sect.sectProps);
				this.renderElements(sect.elements, contentElement);
				pageElement.appendChild(contentElement);
				props = sect.sectProps;
			}

			if (this.options.renderFootnotes) {
				const notes = this.renderNotes(this.currentFootnoteIds, this.footnoteMap);
				notes && pageElement.appendChild(notes);
			}

			if (this.options.renderEndnotes && i == l - 1) {
				const notes = this.renderNotes(this.currentEndnoteIds, this.endnoteMap);
				notes && pageElement.appendChild(notes);
			}

			this.options.renderFooters && this.renderHeaderFooter(props.footerRefs, props,
				result.length, prevProps != props, pageElement);

			result.push(pageElement);
			prevProps = props;
		}

		return result;
	}

	renderHeaderFooter(refs: FooterHeaderReference[], props: SectionProperties, page: number, firstOfSection: boolean, into: HTMLElement) {
		if (!refs) return;

		var ref = (props.titlePage && firstOfSection ? refs.find(x => x.type == "first") : null)
			?? (page % 2 == 1 ? refs.find(x => x.type == "even") : null)
			?? refs.find(x => x.type == "default");

		var part = ref && this.document.findPartByRelId(ref.id, this.document.documentPart) as BaseHeaderFooterPart;

		if (part) {
			this.currentPart = part;
			if (!this.usedHederFooterParts.includes(part.path)) {
				this.processElement(part.rootElement);
				this.usedHederFooterParts.push(part.path);
			}
			const [el] = this.renderElements([part.rootElement], into) as HTMLElement[];

			if (props?.pageMargins) {
				if (part.rootElement.type === DomType.Header) {
					el.style.marginTop = `calc(${props.pageMargins.header} - ${props.pageMargins.top})`;
					el.style.minHeight = `calc(${props.pageMargins.top} - ${props.pageMargins.header})`;
				}
				else if (part.rootElement.type === DomType.Footer) {
					el.style.marginBottom = `calc(${props.pageMargins.footer} - ${props.pageMargins.bottom})`;
					el.style.minHeight = `calc(${props.pageMargins.bottom} - ${props.pageMargins.footer})`;
				}
			}

			this.currentPart = null;
		}
	}

	isPageBreakElement(elem: OpenXmlElement): boolean {
		if (elem.type != DomType.Break)
			return false;

		if ((elem as WmlBreak).break == "lastRenderedPageBreak")
			return !this.options.ignoreLastRenderedPageBreak;

		return (elem as WmlBreak).break == "page";
	}

	isPageBreakSection(prev: SectionProperties, next: SectionProperties): boolean {
		if (!prev) return false;
		if (!next) return false;

		return prev.pageSize?.orientation != next.pageSize?.orientation
			|| prev.pageSize?.width != next.pageSize?.width
			|| prev.pageSize?.height != next.pageSize?.height;
	}

	splitBySection(elements: OpenXmlElement[], defaultProps: SectionProperties): Section[] {
		var current: Section = { sectProps: null, elements: [], pageBreak: false };
		var result = [current];

		for (let elem of elements) {
			if (elem.type == DomType.Paragraph) {
				const s = this.findStyle((elem as WmlParagraph).styleName);

				if (s?.paragraphProps?.pageBreakBefore) {
					// TODO(correctness): see SECURITY_REVIEW.md #7
					current.sectProps = sectProps;
					current.pageBreak = true;
					current = { sectProps: null, elements: [], pageBreak: false };
					result.push(current);
				}
			}

			current.elements.push(elem);

			if (elem.type == DomType.Paragraph) {
				const p = elem as WmlParagraph;

				var sectProps = p.sectionProps;
				var pBreakIndex = -1;
				var rBreakIndex = -1;

				if (this.options.breakPages && p.children) {
					pBreakIndex = p.children.findIndex(r => {
						rBreakIndex = r.children?.findIndex(this.isPageBreakElement.bind(this)) ?? -1;
						return rBreakIndex != -1;
					});
				}

				if (sectProps || pBreakIndex != -1) {
					current.sectProps = sectProps;
					current.pageBreak = pBreakIndex != -1;
					current = { sectProps: null, elements: [], pageBreak: false };
					result.push(current);
				}

				if (pBreakIndex != -1) {
					let breakRun = p.children[pBreakIndex];
					let splitRun = rBreakIndex < breakRun.children.length - 1;

					if (pBreakIndex < p.children.length - 1 || splitRun) {
						var children = elem.children;
						var newParagraph = { ...elem, children: children.slice(pBreakIndex) };
						elem.children = children.slice(0, pBreakIndex);
						current.elements.push(newParagraph);

						if (splitRun) {
							let runChildren = breakRun.children;
							let newRun = { ...breakRun, children: runChildren.slice(0, rBreakIndex) };
							elem.children.push(newRun);
							breakRun.children = runChildren.slice(rBreakIndex);
						}
					}
				}
			}
		}

		let currentSectProps = null;

		for (let i = result.length - 1; i >= 0; i--) {
			if (result[i].sectProps == null) {
				result[i].sectProps = currentSectProps ?? defaultProps;
			} else {
				currentSectProps = result[i].sectProps
			}
		}

		return result;
	}

	groupByPageBreaks(sections: Section[]): Section[][] {
		let current = [];
		let prev: SectionProperties;
		const result: Section[][] = [current];

		for (let s of sections) {
			current.push(s);

			if (this.options.ignoreLastRenderedPageBreak || s.pageBreak || this.isPageBreakSection(prev, s.sectProps))
				result.push(current = []);

			prev = s.sectProps;
		}

		return result.filter(x => x.length > 0);
	}

	renderWrapper(children: HTMLElement[]) {
		return this.h({ tagName: "div", className: `${this.className}-wrapper`, children });
	}

	renderWrapperWithSidebar(sectionElements: HTMLElement[]) {
		const c = this.className;

		const docContainer = this.h({ tagName: "div", className: `${c}-doc-container`, children: sectionElements }) as HTMLElement;

		this.sidebarContainer = this.h({
			tagName: "div",
			className: `${c}-comment-sidebar ${c}-sidebar-${this.sidebarLayout}`
		}) as HTMLElement;

		const contentArea = this.h({
			tagName: "div",
			className: `${c}-sidebar-content`,
			children: []
		}) as HTMLElement;

		this.sidebarContainer.appendChild(contentArea);

		this.renderSidebarComments(contentArea);

		const wrapper = this.h({
			tagName: "div",
			className: `${c}-wrapper`,
			children: [docContainer, this.sidebarContainer]
		}) as HTMLElement;

		this.later(() => {
			this.setupSidebarScrollSync(docContainer, contentArea);
		});

		return wrapper;
	}

	setupSidebarScrollSync(docContainer: HTMLElement, sidebarContent: HTMLElement) {
		// Packed mode: natural CSS flow already stacks cards flush with zero
		// gaps. No measurement or scroll listener needed.
		if (this.sidebarLayout === 'packed') return;

		// Anchored mode: the sidebar grows vertically to match the document
		// and rides the same scroll container, so each card's position is
		// computed once relative to the sidebar-content flow and stays
		// visually locked to its anchor without per-scroll recalculation.
		// We still re-run on resize (fonts, images, reflow), since that
		// genuinely changes anchor positions inside the content.
		const CARD_GAP = 8;

		const positionCards = () => {
			const anchored: Array<{ el: HTMLElement; anchor: HTMLElement; desiredTop: number }> = [];
			for (const [id, sidebarEl] of Object.entries(this.sidebarCommentElements)) {
				if (!sidebarEl.isConnected) continue;
				const anchor = this.commentAnchorElements[id]?.[0];
				if (!anchor?.isConnected) continue;
				anchored.push({ el: sidebarEl, anchor, desiredTop: 0 });
			}
			for (const meta of this.changeMeta) {
				const card = this.revisionCardElements.get(meta.id ?? '');
				if (!card?.isConnected || !meta.el.isConnected) continue;
				anchored.push({ el: card, anchor: meta.el, desiredTop: 0 });
			}
			if (anchored.length === 0) return;

			// Compute each card's target Y in sidebar-content coordinate space.
			// The sidebar rides the doc scroll in anchored mode, so this math
			// is scroll-invariant and only needs to run once after layout.
			const sidebarRect = sidebarContent.getBoundingClientRect();
			for (const entry of anchored) {
				const r = entry.anchor.getBoundingClientRect();
				entry.desiredTop = r.top - sidebarRect.top + sidebarContent.scrollTop;
			}

			// Sort by document order and push down on collision so cards
			// never overlap. Dense clusters end up below their anchors.
			anchored.sort((a, b) => a.desiredTop - b.desiredTop);
			for (const { el } of anchored) el.style.marginTop = "";

			let floor = -Infinity;
			for (const entry of anchored) {
				const target = Math.max(entry.desiredTop, floor);
				const naturalTop = entry.el.offsetTop;
				const offset = target - naturalTop;
				if (offset > 0) entry.el.style.marginTop = `${offset}px`;
				floor = target + entry.el.offsetHeight + CARD_GAP;
			}
		};

		let rafId: number;
		const schedule = () => {
			cancelAnimationFrame(rafId);
			rafId = requestAnimationFrame(positionCards);
		};

		if (typeof ResizeObserver !== "undefined") {
			const ro = new ResizeObserver(schedule);
			ro.observe(docContainer);
			for (const el of Object.values(this.sidebarCommentElements)) {
				if (el.isConnected) ro.observe(el);
			}
		}

		// Initial pass after layout settles (fonts, images).
		setTimeout(positionCards, 100);
	}

	renderSidebarComments(container: HTMLElement) {
		const commentsPart = this.document.commentsPart;
		if (!commentsPart) return;

		const comments = commentsPart.topLevelComments.length > 0
			? commentsPart.topLevelComments
			: commentsPart.comments;

		for (const comment of comments) {
			const el = this.renderSidebarComment(comment, false);
			if (el) container.appendChild(el);
		}
	}

	renderSidebarComment(comment: WmlComment, isReply: boolean): HTMLElement {
		const c = this.className;

		const headerChildren: Node[] = [
			this.h({ tagName: "span", className: `${c}-comment-author`, children: [comment.author ?? "Unknown"] }),
			this.h({ tagName: "span", className: `${c}-comment-date`, children: [comment.date ? new Date(comment.date).toLocaleString() : ""] })
		];

		if (comment.done) {
			headerChildren.push(this.h({ tagName: "span", className: `${c}-comment-done`, children: ["Done"] }));
		}

		const header = this.h({
			tagName: "div",
			className: `${c}-comment-header`,
			children: headerChildren
		}) as HTMLElement;

		const bodyEl = this.h({
			tagName: "div",
			className: `${c}-comment-body`,
			children: this.renderElements(comment.children)
		}) as HTMLElement;

		const children: Node[] = [header, bodyEl];

		if (comment.replies && comment.replies.length > 0) {
			const repliesContainer = this.h({
				tagName: "div",
				className: `${c}-comment-replies`,
				children: comment.replies.map(r => this.renderSidebarComment(r, true))
			}) as HTMLElement;

			const threadToggle = this.h({
				tagName: "button",
				className: `${c}-thread-toggle`,
				children: [`${comment.replies.length} ${comment.replies.length === 1 ? 'reply' : 'replies'}`]
			}) as HTMLButtonElement;

			children.push(threadToggle);
			children.push(repliesContainer);

			this.later(() => {
				threadToggle.addEventListener("click", (ev) => {
					ev.stopPropagation();
					repliesContainer.classList.toggle(`${c}-replies-collapsed`);
					threadToggle.classList.toggle(`${c}-thread-collapsed`);
				});
			});
		}

		const commentEl = this.h({
			tagName: "div",
			className: cx(`${c}-sidebar-comment`, isReply && `${c}-sidebar-reply`),
			children
		}) as HTMLElement;

		commentEl.dataset.commentId = comment.id;

		if (!isReply) {
			this.sidebarCommentElements[comment.id] = commentEl;

			this.later(() => {
				commentEl.addEventListener("click", () => {
					const anchors = this.commentAnchorElements[comment.id];
					if (anchors && anchors.length > 0) {
						anchors[0].scrollIntoView({ behavior: "smooth", block: "center" });
					}
				});
			});
		}

		return commentEl;
	}

	renderDefaultStyle() {
		var c = this.className;
		var wrapperStyle = `
.${c}-wrapper { background: gray; padding: 30px; padding-bottom: 0px; display: flex; flex-flow: column; align-items: center; } 
.${c}-wrapper>section.${c} { background: white; box-shadow: 0 0 10px rgba(0, 0, 0, 0.5); margin-bottom: 30px; }`;
		if (this.options.hideWrapperOnPrint) {
			wrapperStyle = `@media not print { ${wrapperStyle} }`;
		}
		var styleText = `${wrapperStyle}
.${c} { color: black; hyphens: auto; text-underline-position: from-font; }
section.${c} { box-sizing: border-box; display: flex; flex-flow: column nowrap; position: relative; overflow: hidden; }
section.${c}>article { margin-bottom: auto; z-index: 1; }
section.${c}>footer { z-index: 1; }
.${c} table { border-collapse: collapse; }
.${c} table td, .${c} table th { vertical-align: top; }
.${c} p { margin: 0pt; min-height: 1em; }
.${c} span { white-space: pre-wrap; overflow-wrap: break-word; }
.${c} a { color: inherit; text-decoration: inherit; }
.${c} svg { fill: transparent; }
`;

		if (this.options.renderComments) {
			if (this.useSidebar) {
				styleText += `
.${c}-wrapper { flex-flow: row !important; align-items: flex-start !important; }
.${c}-doc-container { flex: 1; display: flex; flex-flow: column; align-items: center; min-width: 0; overflow: auto; padding: 30px; padding-bottom: 0; }
.${c}-doc-container>section.${c} { background: white; box-shadow: 0 0 10px rgba(0, 0, 0, 0.5); margin-bottom: 30px; }
.${c}-comment-sidebar { width: 320px; min-width: 260px; background: #fafafa; border-left: 1px solid #ddd; display: flex; flex-direction: column; transition: width 0.2s, min-width 0.2s, padding 0.2s; }
/* packed mode: panel stays pinned as a short compact list at the top of the viewport. */
.${c}-comment-sidebar.${c}-sidebar-packed { position: sticky; top: 0; height: 100vh; overflow: hidden; align-self: flex-start; }
/* anchored mode: panel grows to match the document height and rides the same scroll container so each card stays next to its anchor. The toolbar inside it is sticky so it's always visible. */
.${c}-comment-sidebar.${c}-sidebar-anchored { align-self: stretch; }
.${c}-sidebar-packed .${c}-sidebar-content { flex: 1; overflow-y: auto; padding: 8px; }
.${c}-sidebar-anchored .${c}-sidebar-content { padding: 8px; }
.${c}-sidebar-comment { background: white; border: 1px solid #e0e0e0; border-radius: 6px; padding: 10px; margin-bottom: 8px; cursor: pointer; transition: box-shadow 0.2s, border-color 0.2s; }
.${c}-sidebar-comment:hover { border-color: #4a90d9; box-shadow: 0 1px 4px rgba(74, 144, 217, 0.2); }
.${c}-sidebar-reply { margin-left: 16px; border-left: 3px solid #4a90d9; background: #f8fbff; }
.${c}-comment-header { display: flex; align-items: baseline; gap: 8px; margin-bottom: 4px; flex-wrap: wrap; }
.${c}-comment-author { font-weight: 600; font-size: 0.85rem; color: #333; }
.${c}-comment-date { font-size: 0.75rem; color: #999; }
.${c}-comment-done { font-size: 0.7rem; background: #4caf50; color: white; padding: 1px 6px; border-radius: 3px; }
.${c}-comment-body { font-size: 0.85rem; color: #444; margin-bottom: 6px; line-height: 1.4; }
.${c}-comment-body p { margin: 2px 0; }
.${c}-comment-replies { margin-top: 6px; }
.${c}-replies-collapsed { display: none; }
.${c}-thread-toggle { background: none; border: none; color: #4a90d9; cursor: pointer; font-size: 0.8rem; padding: 2px 0; margin-top: 4px; }
.${c}-thread-toggle:hover { text-decoration: underline; }
.${c}-thread-collapsed::before { content: "▶ "; }
.${c}-thread-toggle:not(.${c}-thread-collapsed)::before { content: "▼ "; }
.${c}-comment-focused { border-color: #ff9800 !important; box-shadow: 0 0 8px rgba(255, 152, 0, 0.4) !important; }
.${c}-comment-anchor-start { cursor: pointer; }
::highlight(${c}-comments) { background-color: rgba(255, 212, 0, 0.35); }
.${c}-no-highlight .${c}-comment-anchor-start { cursor: default; }
`;
			} else {
				styleText += `
.${c}-comment-ref { cursor: default; }
.${c}-comment-popover { display: none; z-index: 1000; padding: 0.5rem; background: white; position: absolute; box-shadow: 0 0 0.25rem rgba(0, 0, 0, 0.25); width: 30ch; }
.${c}-comment-ref:hover~.${c}-comment-popover { display: block; }
.${c}-comment-author,.${c}-comment-date { font-size: 0.875rem; color: #888; }
`;
			}
		};

		if (this.showChanges) {
			styleText += this.changesStyles();
		}

		return [
			this.h({ tagName: "#comment", children: ["docxjs library predefined styles"] }),
			this.h({ tagName: "style", children: [styleText] })
		];
	}

	private changesStyles(): string {
		const c = this.className;
		// WCAG-AA contrast on white. Same hue for ins and del per author so
		// authorship stays visually trackable when both forms appear.
		const palette = [
			"#2563eb", "#dc2626", "#16a34a", "#9333ea",
			"#ea580c", "#0891b2", "#c026d3", "#65a30d"
		];
		let css = `
.${c} ins { text-decoration: underline; text-decoration-thickness: 2px; background: transparent; }
.${c} del { text-decoration: line-through; text-decoration-thickness: 2px; }
.${c} .${c}-move-from { text-decoration: line-through double; text-decoration-thickness: 1px; cursor: pointer; }
.${c} .${c}-move-to { text-decoration: underline double; text-decoration-thickness: 1px; cursor: pointer; }
.${c} .${c}-formatting-revision { text-decoration: underline dotted; text-decoration-thickness: 1px; cursor: help; }
.${c}-paragraph-mark { margin-left: 2px; font-weight: bold; user-select: none; }
.${c}-paragraph-mark-deleted { text-decoration: line-through; }
.${c}-row-inserted > td { background: color-mix(in srgb, currentColor 8%, transparent); }
.${c}-row-deleted > td { background: color-mix(in srgb, currentColor 10%, transparent); text-decoration: line-through; text-decoration-color: currentColor; text-decoration-thickness: 2px; }
.${c}-revision-kind { margin-left: auto; font-size: 0.7rem; padding: 1px 6px; border: 1px solid currentColor; border-radius: 3px; text-transform: uppercase; }
.${c}-revision-card { border-left: 3px solid currentColor; }
.${c}-change-bar { position: relative; }
.${c}-change-bar::before { content: ""; position: absolute; left: -12px; top: 0; bottom: 0; width: 2px; background: currentColor; opacity: 0.55; }
.${c}-legend { display: flex; flex-wrap: wrap; gap: 12px; align-items: center; padding: 8px 12px; margin: 0 auto 12px; background: #f5f5f5; border: 1px solid #ddd; border-radius: 4px; font-size: 0.85rem; color: #333; max-width: calc(100% - 60px); }
.${c}-legend-label { font-weight: 600; margin-right: 4px; }
.${c}-legend-item { display: inline-flex; align-items: center; gap: 4px; }
.${c}-legend-swatch { display: inline-block; width: 12px; height: 12px; border-radius: 2px; }
`;
		for (let i = 0; i < HtmlRenderer.CHANGE_PALETTE_SIZE; i++) {
			css += `.${c}-change-author-${i} { color: ${palette[i]}; text-decoration-color: ${palette[i]}; }\n`;
		}
		return css;
	}

	// renderNumbering2(numberingPart: NumberingPartProperties, container: HTMLElement): HTMLElement {
	//     let css = "";
	//     const numberingMap = keyBy(numberingPart.abstractNumberings, x => x.id);
	//     const bulletMap = keyBy(numberingPart.bulletPictures, x => x.id);
	//     const topCounters = [];

	//     for(let num of numberingPart.numberings) {
	//         const absNum = numberingMap[num.abstractId];

	//         for(let lvl of absNum.levels) {
	//             const className = this.numberingClass(num.id, lvl.level);
	//             let listStyleType = "none";

	//             if(lvl.text && lvl.format == 'decimal') {
	//                 const counter = this.numberingCounter(num.id, lvl.level);

	//                 if (lvl.level > 0) {
	//                     css += this.styleToString(`p.${this.numberingClass(num.id, lvl.level - 1)}`, {
	//                         "counter-reset": counter
	//                     });
	//                 } else {
	//                     topCounters.push(counter);
	//                 }

	//                 css += this.styleToString(`p.${className}:before`, {
	//                     "content": this.levelTextToContent(lvl.text, num.id),
	//                     "counter-increment": counter
	//                 });
	//             } else if(lvl.bulletPictureId) {
	//                 let pict = bulletMap[lvl.bulletPictureId];
	//                 let variable = `--${this.className}-${pict.referenceId}`.toLowerCase();

	//                 css += this.styleToString(`p.${className}:before`, {
	//                     "content": "' '",
	//                     "display": "inline-block",
	//                     "background": `var(${variable})`
	//                 }, pict.style);

	//                 this.document.loadNumberingImage(pict.referenceId).then(data => {
	//                     var text = `.${this.className}-wrapper { ${variable}: url(${data}) }`;
	//                     container.appendChild(createStyleElement(text));
	//                 });
	//             } else {
	//                 listStyleType = this.numFormatToCssValue(lvl.format);
	//             }

	//             css += this.styleToString(`p.${className}`, {
	//                 "display": "list-item",
	//                 "list-style-position": "inside",
	//                 "list-style-type": listStyleType,
	//                 //TODO
	//                 //...num.style
	//             });
	//         }
	//     }

	//     if (topCounters.length > 0) {
	//         css += this.styleToString(`.${this.className}-wrapper`, {
	//             "counter-reset": topCounters.join(" ")
	//         });
	//     }

	//     return createStyleElement(css);
	// }

	async renderNumbering(numberings: IDomNumbering[]) {
		var styleText = "";
		var resetCounters = [];

		for (var num of numberings) {
			// `num.id` / `num.level` / `num.bullet.src` land in CSS selectors
			// and custom-property names. Skip any entry whose identifiers
			// aren't plain alphanumeric/underscore — otherwise an attacker
			// DOCX could break out of the selector. See SECURITY_REVIEW.md #3.
			if (!isSafeCssIdent(String(num.id)) || !Number.isInteger(num.level)) {
				continue;
			}

			var selector = `p.${this.numberingClass(num.id, num.level)}`;
			var listStyleType = "none";

			if (num.bullet) {
				if (!isSafeCssIdent(String(num.bullet.src))) {
					continue;
				}
				let valiable = `--${this.className}-${num.bullet.src}`.toLowerCase();

				// `num.bullet.style` is a raw VML style attribute from DOCX.
				// Dropping it entirely is safest; width/height can still be
				// expressed via the sanitized `num.pStyle` path on the
				// selector below. See SECURITY_REVIEW.md #3.
				styleText += this.styleToString(`${selector}:before`, {
					"content": "' '",
					"display": "inline-block",
					"background": `var(${valiable})`
				});

				try {
					const imgData = await this.document.loadNumberingImage(num.bullet.src);
					styleText += `${this.rootSelector} { ${valiable}: url(${imgData}) }`;
				} catch(e) {
					if (this.options.debug) console.warn(`Can't load numbering image with src ${num.bullet.src}`);
				}
			}
			else if (num.levelText) {
				let counter = this.numberingCounter(num.id, num.level);
				const counterReset = counter + " " + (num.start - 1);
				if (num.level > 0) {
					styleText += this.styleToString(`p.${this.numberingClass(num.id, num.level - 1)}`, {
						"counter-set": counterReset
					});
				}
				// reset all level counters with start value
				resetCounters.push(counterReset);

				// `levelTextToContent` escapes the attacker-controlled literal
				// chunks before composing the `content` expression; counter
				// names / numformat are validated above.
				styleText += this.styleToString(`${selector}:before`, {
					"content": this.levelTextToContent(num.levelText, num.suff, num.id, this.numFormatToCssValue(num.format)),
					"counter-increment": counter,
					...num.rStyle,
				});
			}
			else {
				listStyleType = this.numFormatToCssValue(num.format);
			}

			styleText += this.styleToString(selector, {
				"display": "list-item",
				"list-style-position": "inside",
				"list-style-type": listStyleType,
				...num.pStyle
			});
		}

		if (resetCounters.length > 0) {
			styleText += this.styleToString(this.rootSelector, {
				"counter-reset": resetCounters.join(" ")
			});
		}

		return [
			this.h({ tagName: "#comment", children: ["docxjs document numbering styles"] }),
			this.h({ tagName: "style", children: [styleText] })
		];
	}

	renderStyles(styles: IDomStyle[]) {
		var styleText = "";
		const stylesMap = this.styleMap;
		const defautStyles = keyBy(styles.filter(s => s.isDefault), s => s.target);

		for (const style of styles) {
			var subStyles = style.styles;

			if (style.linked) {
				var linkedStyle = style.linked && stylesMap[style.linked];

				if (linkedStyle)
					subStyles = subStyles.concat(linkedStyle.styles);
				else if (this.options.debug)
					console.warn(`Can't find linked style ${style.linked}`);
			}

			for (const subStyle of subStyles) {
				//TODO temporary disable modificators until test it well
				var selector = `${style.target ?? ''}.${style.cssName}`; //${subStyle.mod ?? ''} 

				if (style.target != subStyle.target)
					selector += ` ${subStyle.target}`;

				if (defautStyles[style.target] == style)
					selector = `.${this.className} ${style.target}, ` + selector;

				styleText += this.styleToString(selector, subStyle.values);
			}
		}

		return [
			this.h({ tagName: "#comment", children: ["docxjs document styles"] }),
			this.h({ tagName: "style", children: [styleText] })
		];
	}

	renderNotes(noteIds: string[], notesMap: Record<string, WmlBaseNote>) {
		var notes = noteIds.map(id => notesMap[id]).filter(x => x);

		if (notes.length > 0) {
			return this.h({ tagName: "ol", children: this.renderElements(notes) });
		}
	}

	renderElement(elem: OpenXmlElement): Node | Node[] {
		switch (elem.type) {
			case DomType.Paragraph:
				return this.renderParagraph(elem as WmlParagraph);

			case DomType.BookmarkStart:
				return this.renderBookmarkStart(elem as WmlBookmarkStart);

			case DomType.BookmarkEnd:
				return null; //ignore bookmark end

			case DomType.Run:
				return this.renderRun(elem as WmlRun);

			case DomType.Table:
				return this.renderTable(elem);

			case DomType.Row:
				return this.renderTableRow(elem);

			case DomType.Cell:
				return this.renderTableCell(elem);

			case DomType.Hyperlink:
				return this.renderHyperlink(elem);
			
			case DomType.SmartTag:
				return this.renderSmartTag(elem);

			case DomType.Drawing:
				return this.renderDrawing(elem);

			case DomType.Image:
				return this.renderImage(elem as IDomImage);

			case DomType.Text:
				return this.renderText(elem as WmlText);

			case DomType.Text:
				return this.renderText(elem as WmlText);

			case DomType.DeletedText:
				return this.renderDeletedText(elem as WmlText);
	
			case DomType.Tab:
				return this.renderTab(elem);

			case DomType.Symbol:
				return this.renderSymbol(elem as WmlSymbol);

			case DomType.Break:
				return this.renderBreak(elem as WmlBreak);

			case DomType.Footer:
				return this.renderContainer(elem, "footer");

			case DomType.Header:
				return this.renderContainer(elem, "header");

			case DomType.Footnote:
			case DomType.Endnote:
				return this.renderContainer(elem, "li");

			case DomType.FootnoteReference:
				return this.renderFootnoteReference(elem as WmlNoteReference);

			case DomType.EndnoteReference:
				return this.renderEndnoteReference(elem as WmlNoteReference);

			case DomType.NoBreakHyphen:
				return this.h({ tagName: "wbr" });

			case DomType.VmlPicture:
				return this.renderVmlPicture(elem);

			case DomType.VmlElement:
				return this.renderVmlElement(elem as VmlElement);
	
			case DomType.MmlMath:
				return this.renderContainerNS(elem, ns.mathML, "math", { xmlns: ns.mathML });
	
			case DomType.MmlMathParagraph:
				return this.renderContainer(elem, "span");

			case DomType.MmlFraction:
				return this.renderContainerNS(elem, ns.mathML, "mfrac");

			case DomType.MmlBase:
				return this.renderContainerNS(elem, ns.mathML, 
					elem.parent.type == DomType.MmlMatrixRow ? "mtd" : "mrow");

			case DomType.MmlNumerator:
			case DomType.MmlDenominator:
			case DomType.MmlFunction:
			case DomType.MmlLimit:
			case DomType.MmlBox:
				return this.renderContainerNS(elem, ns.mathML, "mrow");

			case DomType.MmlGroupChar:
				return this.renderMmlGroupChar(elem);

			case DomType.MmlLimitLower:
				return this.renderContainerNS(elem, ns.mathML, "munder");

			case DomType.MmlMatrix:
				return this.renderContainerNS(elem, ns.mathML, "mtable");

			case DomType.MmlMatrixRow:
				return this.renderContainerNS(elem, ns.mathML, "mtr");
	
			case DomType.MmlRadical:
				return this.renderMmlRadical(elem);

			case DomType.MmlSuperscript:
				return this.renderContainerNS(elem, ns.mathML, "msup");

			case DomType.MmlSubscript:
				return this.renderContainerNS(elem, ns.mathML, "msub");

			case DomType.MmlDegree:
			case DomType.MmlSuperArgument:
			case DomType.MmlSubArgument:
				return this.renderContainerNS(elem, ns.mathML, "mn");

			case DomType.MmlFunctionName:
				return this.renderContainerNS(elem, ns.mathML, "ms");
	
			case DomType.MmlDelimiter:
				return this.renderMmlDelimiter(elem);

			case DomType.MmlRun:
				return this.renderMmlRun(elem);

			case DomType.MmlNary:
				return this.renderMmlNary(elem);

			case DomType.MmlPreSubSuper:
				return this.renderMmlPreSubSuper(elem);

			case DomType.MmlBar:
				return this.renderMmlBar(elem);
	
			case DomType.MmlEquationArray:
				return this.renderMllList(elem);

			case DomType.Inserted:
				return this.renderInserted(elem);

			case DomType.Deleted:
				return this.renderDeleted(elem);

			case DomType.MoveFrom:
				return this.renderMoveFrom(elem);

			case DomType.MoveTo:
				return this.renderMoveTo(elem);

			case DomType.CommentRangeStart:
				return this.renderCommentRangeStart(elem);

			case DomType.CommentRangeEnd:
				return this.renderCommentRangeEnd(elem);

			case DomType.CommentReference:
				return this.renderCommentReference(elem);

			case DomType.AltChunk:
				// AltChunk rendering removed for security — the old implementation
				// assigned attacker-controlled HTML to iframe.srcdoc (same-origin,
				// no sandbox). The read-only viewer does not need alt chunks; the
				// parser still produces a node so consumers can detect them.
				// See SECURITY_REVIEW.md #1.
				return null;
		}

		return null;
	}
	renderElements(elems: OpenXmlElement[], into?: Node): Node[] {
		if (elems == null)
			return null;

		var result = elems.flatMap(e => this.renderElement(e)).filter(e => e != null);

		if (into)
			result.forEach(c => into.appendChild(isString(c) ? document.createTextNode(c) : c));

		return result;
	}

	renderContainer<T extends keyof HTMLElementTagNameMap>(elem: OpenXmlElement, tagName: T): HTMLElementTagNameMap[T] {
		return this.h({ tagName, children: this.renderElements(elem.children) }) as any;
	}

	renderContainerNS(elem: OpenXmlElement, ns: ns, tagName: string, props?: Record<string, any>) {
		return this.h({ ns, tagName, children: this.renderElements(elem.children), ...props });
	}

	renderParagraph(elem: WmlParagraph) {
		var result = this.toHTML(elem, ns.html, "p");

		const style = this.findStyle(elem.styleName);
		elem.tabs ??= style?.paragraphProps?.tabs;  //TODO

		const numbering = elem.numbering ?? style?.paragraphProps?.numbering;

		if (numbering) {
			result.classList.add(this.numberingClass(numbering.id, numbering.level));
		}

		if (this.showChanges && elem.paragraphMarkRevisionKind) {
			this.appendParagraphMarkRevision(result, elem);
		}

		this.applyFormattingRevision(result, elem);

		// Expose the Word-assigned paraId so consumers can correlate/navigate
		// paragraphs by their stable DOCX identifier. The value is attacker-
		// controlled, but dataset.* attribute-encodes it — no innerHTML sink.
		if (elem.paraId) {
			result.dataset.paraId = elem.paraId;
		}

		return result;
	}

	private appendParagraphMarkRevision(paragraphEl: HTMLElement, elem: WmlParagraph) {
		const c = this.className;
		const kind = elem.paragraphMarkRevisionKind;
		const rev = elem.revision;
		if (!kind) return;

		const classes = [`${c}-paragraph-mark`, `${c}-paragraph-mark-${kind}`];
		if (rev?.author && this.options.changes?.colorByAuthor !== false) {
			classes.push(`${c}-change-author-${this.getAuthorIndex(rev.author)}`);
		}

		const mark = this.h({
			tagName: "span",
			className: classes.join(" "),
			children: ["¶"]
		}) as HTMLElement;
		if (rev?.id) mark.dataset.changeId = rev.id;
		if (rev?.author) mark.dataset.author = rev.author;
		if (rev?.date) mark.dataset.date = rev.date;
		mark.dataset.changeKind = "paragraphMark";
		mark.setAttribute("aria-label", kind === "inserted" ? "Paragraph inserted" : "Paragraph mark deleted");

		paragraphEl.appendChild(mark);
		this.changeElements.push(mark);
		this.changeMeta.push({
			el: mark, id: rev?.id, kind: "paragraphMark",
			author: rev?.author, date: rev?.date,
			summary: this.summarizeChange(mark, "paragraphMark"),
		});
	}

	renderHyperlink(elem: WmlHyperlink) {
		const res = this.toH(elem, ns.html, "a");
		let rawHref = '';

		if (elem.id) {
			const rel = this.document.documentPart.rels.find(it => it.id == elem.id && it.targetMode === "External");
			rawHref = rel?.target ?? '';
		}

		// Validate the scheme before emitting. DOCX hyperlink targets are
		// attacker-controlled, so `javascript:` / `data:` / etc. must never
		// land in an `<a href>`. Unsafe targets drop the href entirely and
		// render the visible link text as an inert span. See
		// SECURITY_REVIEW.md #2.
		if (rawHref && !isSafeHyperlinkHref(rawHref)) {
			// Render the children without wrapping them in <a> — produces plain
			// text (or whatever runs the hyperlink contained) with no sink.
			return this.h({
				ns: ns.html,
				tagName: "span",
				className: res.className,
				style: res.style,
				children: res.children,
			});
		}

		let href = rawHref;
		// Anchor fragments are opaque to the URL parser from the host page's
		// perspective; safe to append. See SECURITY_REVIEW.md #2 note.
		if (elem.anchor) {
			href += `#${elem.anchor}`;
		}

		res.href = href;
		return this.h(res);
	}
	
	renderSmartTag(elem: WmlSmartTag) {
		return this.renderContainer(elem, "span");
	}
	
	renderCommentRangeStart(commentStart: WmlCommentRangeStart) {
		if (!this.options.renderComments)
			return null;

		if (this.useSidebar) {
			const anchor = this.h({ tagName: "span", className: `${this.className}-comment-anchor-start` }) as HTMLElement;
			anchor.dataset.commentId = commentStart.id;

			if (!this.commentAnchorElements[commentStart.id]) {
				this.commentAnchorElements[commentStart.id] = [];
			}
			this.commentAnchorElements[commentStart.id].push(anchor);

			if (this.useHighlight) {
				const rng = new Range();
				this.commentHighlight?.add(rng);
				this.later(() => rng.setStart(anchor, 0));
				this.commentMap[commentStart.id] = rng;
			}

			this.later(() => {
				anchor.addEventListener("click", () => {
					const sidebarEl = this.sidebarCommentElements[commentStart.id];
					if (sidebarEl) {
						sidebarEl.scrollIntoView({ behavior: "smooth", block: "center" });
						sidebarEl.classList.add(`${this.className}-comment-focused`);
						setTimeout(() => sidebarEl.classList.remove(`${this.className}-comment-focused`), 2000);
					}
				});
			});

			return anchor;
		}

		const rng = new Range();
		this.commentHighlight?.add(rng);

		const result = this.h({ tagName: "#comment", children: [`start of comment #${commentStart.id}`] });
		this.later(() => rng.setStart(result, 0));
		this.commentMap[commentStart.id] = rng;

		return result
	}

	renderCommentRangeEnd(commentEnd: WmlCommentRangeEnd) {
		if (!this.options.renderComments)
			return null;

		if (this.useSidebar) {
			const anchor = this.h({ tagName: "span", className: `${this.className}-comment-anchor-end` }) as HTMLElement;
			anchor.dataset.commentId = commentEnd.id;

			if (this.useHighlight) {
				const rng = this.commentMap[commentEnd.id];
				this.later(() => rng?.setEnd(anchor, 0));
			}

			return anchor;
		}

		const rng = this.commentMap[commentEnd.id];
		const result = this.h({ tagName: "#comment", children: [`end of comment #${commentEnd.id}`] });
		this.later(() => rng?.setEnd(result, 0));

		return result;
	}

	renderCommentReference(commentRef: WmlCommentReference) {
		if (!this.options.renderComments)
			return null;

		if (this.useSidebar) {
			return this.h({ tagName: "#comment", children: [`comment ref #${commentRef.id}`] });
		}

		var comment = this.document.commentsPart?.commentMap[commentRef.id];

		if (!comment)
			return null;

		const commentRefEl = this.h({ tagName: "span", className: `${this.className}-comment-ref`, children: ['💬'] });
		const commentsContainerEl = this.h({
			tagName: "div", className: `${this.className}-comment-popover`, children: [
				this.h({ tagName: 'div', className: `${this.className}-comment-author`, children: [comment.author] }),
				this.h({ tagName: 'div', className: `${this.className}-comment-date`, children: [new Date(comment.date).toLocaleString()] }),
				...this.renderElements(comment.children)
			]
		});

		return this.h({ tagName:  "#fragment", children: [
			this.h({ tagName: "#comment", children: [`comment #${comment.id} by ${comment.author} on ${comment.date}`] }),
			commentRefEl,
			commentsContainerEl
		] });
	}

	renderDrawing(elem: OpenXmlElement) {
		var result = this.toHTML(elem, ns.html, "div");

		result.style.display = "inline-block";
		result.style.position = "relative";
		result.style.textIndent = "0px";

		return result;
	}

	renderImage(elem: IDomImage) {
		let result = this.toHTML(elem, ns.html, "img", []);
		let transform = elem.cssStyle?.transform;

		if (elem.srcRect && elem.srcRect.some(x => x != 0)) {
			var [left, top, right, bottom] = elem.srcRect;
			transform = `scale(${1 / (1 - left - right)}, ${1 / (1 - top - bottom)})`;
			result.style['clip-path'] = `rect(${(100 * top).toFixed(2)}% ${(100 * (1 - right)).toFixed(2)}% ${(100 * (1 - bottom)).toFixed(2)}% ${(100 * left).toFixed(2)}%)`;
		}

		if (elem.rotation)
			transform = `rotate(${elem.rotation}deg) ${transform ?? ''}`;

		result.style.transform = transform?.trim();

		if (this.document) {
			this.tasks.push(this.document.loadDocumentImage(elem.src, this.currentPart).then(x => {
				result.src = x;
			}));
		}

		return result;
	}

	renderText(elem: WmlText) {
		return this.h(elem.text);
	}

	renderDeletedText(elem: WmlText) {
		return (this.showChanges && this.options.changes?.showDeletions !== false)
			? this.renderText(elem)
			: null;
	}

	renderBreak(elem: WmlBreak) {
		return elem.break == "textWrapping" ? this.h({ tagName: "br" }) : null;
	}

	renderInserted(elem: OpenXmlElement): Node | Node[] {
		if (this.showChanges && this.options.changes?.showInsertions !== false) {
			const node = this.renderContainer(elem, "ins");
			this.applyChangeAttributes(node, elem, "insertion");
			return node;
		}
		return this.renderElements(elem.children);
	}

	renderDeleted(elem: OpenXmlElement): Node {
		if (this.showChanges && this.options.changes?.showDeletions !== false) {
			const node = this.renderContainer(elem, "del");
			this.applyChangeAttributes(node, elem, "deletion");
			return node;
		}
		return null;
	}

	renderMoveFrom(elem: OpenXmlElement): Node | Node[] {
		if (!this.showChanges || this.options.changes?.showMoves === false) {
			return null;
		}
		const node = this.renderContainer(elem, "span") as HTMLElement;
		node.classList.add(`${this.className}-move-from`);
		this.applyChangeAttributes(node, elem, "move");
		this.registerMove(node, elem, "from");
		return node;
	}

	renderMoveTo(elem: OpenXmlElement): Node | Node[] {
		if (!this.showChanges || this.options.changes?.showMoves === false) {
			return this.renderElements(elem.children);
		}
		const node = this.renderContainer(elem, "span") as HTMLElement;
		node.classList.add(`${this.className}-move-to`);
		this.applyChangeAttributes(node, elem, "move");
		this.registerMove(node, elem, "to");
		return node;
	}

	private registerMove(node: HTMLElement, elem: OpenXmlElement, half: "from" | "to") {
		const id = elem.revision?.id;
		if (!id) return;
		node.dataset.moveId = id;

		const pair = this.moveElements.get(id) ?? {};
		pair[half] = node;
		this.moveElements.set(id, pair);

		this.later(() => {
			node.addEventListener("click", (ev) => {
				const entry = this.moveElements.get(id);
				const counterpart = half === "from" ? entry?.to : entry?.from;
				if (counterpart) {
					ev.preventDefault();
					counterpart.scrollIntoView({ behavior: "smooth", block: "center" });
				}
			});
		});
	}

	// Populates data-change-id/author/date and a palette-index class on a
	// rendered <ins>/<del>/move/etc. See Track Changes Phase 1 (#3).
	applyChangeAttributes(node: HTMLElement, elem: OpenXmlElement, kind: ChangeKind) {
		const rev = elem.revision;
		if (!rev) return;

		if (rev.id) node.dataset.changeId = rev.id;
		if (rev.author) node.dataset.author = rev.author;
		if (rev.date) node.dataset.date = rev.date;
		node.dataset.changeKind = kind;

		if (rev.author && this.options.changes?.colorByAuthor !== false) {
			const idx = this.getAuthorIndex(rev.author);
			node.classList.add(`${this.className}-change-author-${idx}`);
		}

		this.changeElements.push(node);
		this.changeMeta.push({
			el: node,
			id: rev.id,
			kind,
			author: rev.author,
			date: rev.date,
			summary: this.summarizeChange(node, kind),
		});
	}

	private summarizeChange(node: HTMLElement, kind: ChangeKind): string {
		const MAX = 80;
		const truncate = (s: string) => {
			const clean = s.replace(/\s+/g, " ").trim();
			return clean.length > MAX ? clean.slice(0, MAX - 1) + "…" : clean;
		};

		switch (kind) {
			case "insertion":
			case "move": {
				const text = truncate(node.textContent ?? "");
				return text ? `Inserted: "${text}"` : "Inserted content";
			}
			case "deletion": {
				const text = truncate(node.textContent ?? "");
				return text ? `Deleted: "${text}"` : "Deleted content";
			}
			case "paragraphMark":
				return "Paragraph mark changed";
			case "rowInsertion":
				return "Row inserted";
			case "rowDeletion":
				return "Row deleted";
			case "formatting": {
				const title = node.getAttribute("title");
				return title ?? "Formatting changed";
			}
		}
	}

	private getAuthorIndex(author: string): number {
		let idx = this.changeAuthorIndex.get(author);
		if (idx === undefined) {
			idx = this.changeAuthorIndex.size % HtmlRenderer.CHANGE_PALETTE_SIZE;
			this.changeAuthorIndex.set(author, idx);
		}
		return idx;
	}

	renderSymbol(elem: WmlSymbol) {
		return this.h({ tagName: "span", children: [String.fromCharCode(elem.char)], style: { fontFamily: elem.font } });
	}

	renderFootnoteReference(elem: WmlNoteReference) {
		this.currentFootnoteIds.push(elem.id);
		return this.h({ tagName: "sup", children: [`${this.currentFootnoteIds.length}`] });
	}

	renderEndnoteReference(elem: WmlNoteReference) {
		this.currentEndnoteIds.push(elem.id);
		return this.h({ tagName: "sup", children: [`${this.currentEndnoteIds.length}`] });
	}

	renderTab(elem: OpenXmlElement) {
		var tabSpan = this.h({ tagName: "span", children: ["\u2003"] }) as HTMLElement;//"&nbsp;";

		if (this.options.experimental) {
			tabSpan.className = this.tabStopClass();
			var stops = findParent<WmlParagraph>(elem, DomType.Paragraph)?.tabs;
			this.currentTabs.push({ stops, span: tabSpan });
		}

		return tabSpan;
	}

	renderBookmarkStart(elem: WmlBookmarkStart) {
		return this.h({ tagName: "span", id: elem.name });
	}

	renderRun(elem: WmlRun) {
		if (elem.fieldRun)
			return null;

		let children = this.renderElements(elem.children);

		if (elem.verticalAlign) {
			children = [this.h({ tagName: elem.verticalAlign, children: this.renderElements(elem.children) })];
		}

		const result = this.toHTML(elem, ns.html, "span", children);

		if (elem.id)
			result.id = elem.id;

		this.applyFormattingRevision(result, elem);

		return result;
	}

	// Marks a run or paragraph as touched by a w:rPrChange / w:pPrChange.
	// The element stays rendered with its *current* formatting — the visible
	// revision is a dotted underline in the author's colour plus a title
	// attribute summarising what changed. The element is registered with the
	// change-bar pass so the paragraph still gets a margin bar even when it
	// only contains formatting revisions.
	applyFormattingRevision(node: HTMLElement, elem: OpenXmlElement) {
		const fr = elem.formattingRevision;
		if (!fr) return;
		if (!this.showChanges || this.options.changes?.showFormatting === false) return;

		const c = this.className;
		node.classList.add(`${c}-formatting-revision`);
		if (fr.id) node.dataset.changeId = fr.id;
		if (fr.author) node.dataset.author = fr.author;
		if (fr.date) node.dataset.date = fr.date;

		if (fr.author && this.options.changes?.colorByAuthor !== false) {
			node.classList.add(`${c}-change-author-${this.getAuthorIndex(fr.author)}`);
		}

		const changed = fr.changedProps && fr.changedProps.length
			? fr.changedProps.join(", ")
			: "formatting";
		const who = fr.author ? `${fr.author} changed` : "Changed";
		node.setAttribute("title", `${who}: ${changed}`);
		node.dataset.changeKind = "formatting";
		if (fr.id) node.dataset.changeId = fr.id;

		this.changeElements.push(node);
		this.changeMeta.push({
			el: node, id: fr.id, kind: "formatting",
			author: fr.author, date: fr.date,
			summary: `${who}: ${changed}`,
		});
	}

	renderTable(elem: WmlTable) {
		this.tableCellPositions.push(this.currentCellPosition);
		this.tableVerticalMerges.push(this.currentVerticalMerge);
		this.currentVerticalMerge = {};
		this.currentCellPosition = { col: 0, row: 0 };

		const children = [];

		if (elem.columns)
			children.push(this.renderTableColumns(elem.columns));

		children.push(...this.renderElements(elem.children));

		this.currentVerticalMerge = this.tableVerticalMerges.pop();
		this.currentCellPosition = this.tableCellPositions.pop();
		return this.toHTML(elem, ns.html, "table", children);
	}

	renderTableColumns(columns: WmlTableColumn[]) {
		const children = columns.map(x => this.h({ tagName: "col", style: { width: x.width } }));
		return this.h({ tagName: "colgroup", children });
	}

	renderTableRow(elem: WmlTableRow) {
		this.currentCellPosition.col = 0;

		const children = [];

		if (elem.gridBefore)
			children.push(this.renderTableCellPlaceholder(elem.gridBefore));

		children.push(...this.renderElements(elem.children));

		if (elem.gridAfter)
			children.push(this.renderTableCellPlaceholder(elem.gridAfter));

		this.currentCellPosition.row++;

		const tr = this.toHTML(elem, ns.html, "tr", children) as HTMLTableRowElement;

		if (this.showChanges && elem.rowRevisionKind) {
			this.applyRowRevision(tr, elem);
		}
		this.applyFormattingRevision(tr, elem);

		return tr;
	}

	// Rows can't be wrapped in <ins>/<del> (invalid HTML), so we decorate the
	// <tr> with a class and data-attrs and lean on CSS for the strikethrough
	// overlay on deletions.
	private applyRowRevision(tr: HTMLTableRowElement, elem: WmlTableRow) {
		const kind = elem.rowRevisionKind;
		if (!kind) return;
		if (kind === "inserted" && this.options.changes?.showInsertions === false) return;
		if (kind === "deleted" && this.options.changes?.showDeletions === false) return;

		const c = this.className;
		tr.classList.add(`${c}-row-${kind}`);
		const rev = elem.revision;
		if (rev?.id) tr.dataset.changeId = rev.id;
		if (rev?.author) tr.dataset.author = rev.author;
		if (rev?.date) tr.dataset.date = rev.date;
		const metaKind: ChangeKind = kind === "inserted" ? "rowInsertion" : "rowDeletion";
		tr.dataset.changeKind = metaKind;

		if (rev?.author && this.options.changes?.colorByAuthor !== false) {
			tr.classList.add(`${c}-change-author-${this.getAuthorIndex(rev.author)}`);
		}

		this.changeElements.push(tr);
		this.changeMeta.push({
			el: tr, id: rev?.id, kind: metaKind,
			author: rev?.author, date: rev?.date,
			summary: this.summarizeChange(tr, metaKind),
		});
	}

	renderTableCellPlaceholder(colSpan: number) {
		return this.h({ tagName: "td", colSpan, style: { border: "none" } });
	}

	renderTableCell(elem: WmlTableCell) {
		let result = this.toHTML(elem, ns.html, "td");

		const key = this.currentCellPosition.col;

		if (elem.verticalMerge) {
			if (elem.verticalMerge == "restart") {
				this.currentVerticalMerge[key] = result;
				result.rowSpan = 1;
			} else if (this.currentVerticalMerge[key]) {
				this.currentVerticalMerge[key].rowSpan += 1;
				result.style.display = "none";
			}
		} else {
			this.currentVerticalMerge[key] = null;
		}

		if (elem.span)
			result.colSpan = elem.span;

		this.currentCellPosition.col += result.colSpan;

		return result;
	}

	renderVmlPicture(elem: OpenXmlElement) {
		return this.renderContainer(elem, "div");
	}

	renderVmlElement(elem: VmlElement): SVGElement {
		// TODO(correctness): see SECURITY_REVIEW.md #8
		// TODO(security): sanitize cssStyleText before emitting to style attribute — see SECURITY_REVIEW.md #5
		var container = this.h({ ns: ns.svg, tagName: "svg", style: elem.cssStyleText }) as SVGElement;

		const result = this.renderVmlChildElement(elem);

		if (elem.imageHref?.id) {
			this.tasks.push(this.document?.loadDocumentImage(elem.imageHref.id, this.currentPart)
				.then(x => result.setAttribute("href", x)));
		}

		container.appendChild(result);

		requestAnimationFrame(() => {
			const bb = (container.firstElementChild as any).getBBox();

			container.setAttribute("width", `${Math.ceil(bb.x +  bb.width)}`);
			container.setAttribute("height", `${Math.ceil(bb.y + bb.height)}`);
		});

		return container;
	}

	renderVmlChildElement(elem: VmlElement): any {
		const result = this.createSvgElement(elem.tagName as any);
		Object.entries(elem.attrs).forEach(([k, v]) => result.setAttribute(k, v));

		for (let child of elem.children) {
			if (child.type == DomType.VmlElement) {
				result.appendChild(this.renderVmlChildElement(child as VmlElement));
			} else {
				result.appendChild(...asArray(this.renderElement(child as any)));
			}
		}

		return result;
	}

	renderMmlRadical(elem: OpenXmlElement) {
		const base = elem.children.find(el => el.type == DomType.MmlBase);

		if (elem.props?.hideDegree) {
			return this.createMathMLElement("msqrt", null, this.renderElements([base]));
		}

		const degree = elem.children.find(el => el.type == DomType.MmlDegree);
		return this.createMathMLElement("mroot", null, this.renderElements([base, degree]));
	}

	renderMmlDelimiter(elem: OpenXmlElement) {		
		const children = [];

		children.push(this.createMathMLElement("mo", null, [elem.props.beginChar ?? '(']));
		children.push(...this.renderElements(elem.children));
		children.push(this.createMathMLElement("mo", null, [elem.props.endChar ?? ')']));

		return this.createMathMLElement("mrow", null, children);
	}

	renderMmlNary(elem: OpenXmlElement) {		
		const children = [];
		const grouped = keyBy(elem.children, x => x.type);

		const sup = grouped[DomType.MmlSuperArgument];
		const sub = grouped[DomType.MmlSubArgument];
		const supElem = sup ? this.createMathMLElement("mo", null, asArray(this.renderElement(sup))) : null;
		const subElem = sub ? this.createMathMLElement("mo", null, asArray(this.renderElement(sub))) : null;

		const charElem = this.createMathMLElement("mo", null, [elem.props?.char ?? '\u222B']);

		if (supElem || subElem) {
			children.push(this.createMathMLElement("munderover", null, [charElem, subElem, supElem]));
		} else if(supElem) {
			children.push(this.createMathMLElement("mover", null, [charElem, supElem]));
		} else if(subElem) {
			children.push(this.createMathMLElement("munder", null, [charElem, subElem]));
		} else {
			children.push(charElem);
		}

		children.push(...this.renderElements(grouped[DomType.MmlBase].children));

		return this.createMathMLElement("mrow", null, children);
	}

	renderMmlPreSubSuper(elem: OpenXmlElement) {
		const children = [];
		const grouped = keyBy(elem.children, x => x.type);

		const sup = grouped[DomType.MmlSuperArgument];
		const sub = grouped[DomType.MmlSubArgument];
		const supElem = sup ? this.createMathMLElement("mo", null, asArray(this.renderElement(sup))) : null;
		const subElem = sub ? this.createMathMLElement("mo", null, asArray(this.renderElement(sub))) : null;
		const stubElem = this.createMathMLElement("mo", null);

		children.push(this.createMathMLElement("msubsup", null, [stubElem, subElem, supElem]));
		children.push(...this.renderElements(grouped[DomType.MmlBase].children));

		return this.createMathMLElement("mrow", null, children);
	}

	renderMmlGroupChar(elem: OpenXmlElement) {
		const tagName = elem.props.verticalJustification === "bot" ? "mover" : "munder";
		const result = this.renderContainerNS(elem, ns.mathML, tagName);

		if (elem.props.char) {
			result.appendChild(this.createMathMLElement("mo", null, [elem.props.char]));
		}

		return result;
	}

	renderMmlBar(elem: OpenXmlElement) {
		const style = {} as any;

		switch(elem.props.position) {
			case "top": style.textDecoration = "overline"; break
			case "bottom": style.textDecoration = "underline"; break
		}

		return this.renderContainerNS(elem, ns.mathML, "mrow", { style }) as MathMLElement;
	}

	renderMmlRun(elem: OpenXmlElement) {
		return this.toHTML(elem, ns.mathML, "ms");
	}

	renderMllList(elem: OpenXmlElement) {
		const children = this.renderElements(elem.children).map(x => this.createMathMLElement("mtr", null, [
			this.createMathMLElement("mtd", null, [x])
		]));

		return this.toHTML(elem, ns.mathML, "mtable", children);
	}

	toH(elem: OpenXmlElement, ns: ns, tagName: string, children: Node[] = null) {
		const { "$lang": lang, ...style } = elem.cssStyle ?? {};
		const className = cx(elem.className, elem.styleName && this.processStyleName(elem.styleName));
		return { ns, tagName, className, lang, style, children: children ?? this.renderElements(elem.children) } as any;
	}

	toHTML(elem: OpenXmlElement, ns: ns, tagName: string, children: Node[] = null) {
		return this.h(this.toH(elem, ns, tagName, children)) as any;
	}

	findStyle(styleName: string) {
		return styleName && this.styleMap?.[styleName];
	}

	numberingClass(id: string, lvl: number) {
		return `${this.className}-num-${id}-${lvl}`;
	}

	tabStopClass() {
		return `${this.className}-tab-stop`;
	}

	styleToString(selectors: string, values: Record<string, string>, cssText: string = null) {
		let result = `${selectors} {\r\n`;

		for (const key in values) {
			if (key.startsWith('$'))
				continue;
			
			result += `  ${key}: ${values[key]};\r\n`;
		}

		if (cssText)
			result += cssText;

		return result + "}\r\n";
	}

	numberingCounter(id: string, lvl: number) {
		return `${this.className}-num-${id}-${lvl}`;
	}

	levelTextToContent(text: string, suff: string, id: string, numformat: string) {
		// text, id, and numformat are all derived from DOCX. Callers have
		// already validated `id` and `numformat`; the `text` body is the last
		// DOCX-controlled value that lands inside a CSS `content` string, so
		// we escape `\` and `"` before embedding. Without this, a crafted
		// levelText of `"}a{background:url(…)}"` would break out of the
		// declaration block. See SECURITY_REVIEW.md #3.
		const suffMap = {
			"tab": "\\9",
			"space": "\\a0",
		};

		// Split literal text from counter placeholders (`%1`, `%2`, ...) so we
		// can escape each literal segment without touching the generated
		// `counter(...)` function. The placeholder regex matches the original
		// behaviour.
		const parts: string[] = [];
		let last = 0;
		const re = /%\d+/g;
		let m: RegExpExecArray | null;
		while ((m = re.exec(text)) !== null) {
			if (m.index > last) {
				parts.push(`"${escapeCssStringContent(text.slice(last, m.index))}"`);
			}
			const lvl = parseInt(m[0].substring(1), 10) - 1;
			parts.push(`counter(${this.numberingCounter(id, lvl)}, ${numformat})`);
			last = re.lastIndex;
		}
		if (last < text.length) {
			parts.push(`"${escapeCssStringContent(text.slice(last))}"`);
		}

		const suffToken = suffMap[suff];
		if (suffToken) {
			parts.push(`"${suffToken}"`);
		}

		// CSS `content` values can be composed of multiple space-separated
		// string / counter() fragments.
		return parts.length > 0 ? parts.join(' ') : '""';
	}

	numFormatToCssValue(format: string) {
		var mapping = {
			none: "none",
			bullet: "disc",
			decimal: "decimal",
			lowerLetter: "lower-alpha",
			upperLetter: "upper-alpha",
			lowerRoman: "lower-roman",
			upperRoman: "upper-roman",
			decimalZero: "decimal-leading-zero", // 01,02,03,...
			// ordinal: "", // 1st, 2nd, 3rd,...
			// ordinalText: "", //First, Second, Third, ...
			// cardinalText: "", //One,Two Three,...
			// numberInDash: "", //-1-,-2-,-3-, ...
			// hex: "upper-hexadecimal",
			aiueo: "katakana",
			aiueoFullWidth: "katakana",
			chineseCounting: "simp-chinese-informal",
			chineseCountingThousand: "simp-chinese-informal",
			chineseLegalSimplified: "simp-chinese-formal", // 中文大写
			chosung: "hangul-consonant",
			ideographDigital: "cjk-ideographic",
			ideographTraditional: "cjk-heavenly-stem", // 十天干
			ideographLegalTraditional: "trad-chinese-formal",
			ideographZodiac: "cjk-earthly-branch", // 十二地支
			iroha: "katakana-iroha",
			irohaFullWidth: "katakana-iroha",
			japaneseCounting: "japanese-informal",
			japaneseDigitalTenThousand: "cjk-decimal",
			japaneseLegal: "japanese-formal",
			thaiNumbers: "thai",
			koreanCounting: "korean-hangul-formal",
			koreanDigital: "korean-hangul-formal",
			koreanDigital2: "korean-hanja-informal",
			hebrew1: "hebrew",
			hebrew2: "hebrew",
			hindiNumbers: "devanagari",
			ganada: "hangul",
			taiwaneseCounting: "cjk-ideographic",
			taiwaneseCountingThousand: "cjk-ideographic",
			taiwaneseDigital:  "cjk-decimal",
		};

		// `format` comes from DOCX. Only emit values from the explicit allow
		// list — dropping an unknown format to "decimal" prevents raw DOCX
		// strings from landing inside a CSS `counter(..., <fmt>)` expression
		// or `list-style-type:` declaration. See SECURITY_REVIEW.md #3.
		return mapping[format] ?? 'decimal';
	}

	refreshTabStops() {
		if (!this.options.experimental)
			return;

		setTimeout(() => {
			const pixelToPoint = computePixelToPoint();

			for (let tab of this.currentTabs) {
				updateTabStop(tab.span, tab.stops, this.defaultTabSize, pixelToPoint);
			}
		}, 500);
	}

	createElementNS(ns: any, tagName: string, props?: Partial<Record<any, any>>, children?: any[]) {
		return this.h({ ns, tagName, children, ...props }) as any;
	}

	createElement<T extends keyof HTMLElementTagNameMap>(tagName: T, props?: Partial<Record<keyof HTMLElementTagNameMap[T], any>>, children?: any[]): HTMLElementTagNameMap[T] {
		return this.createElementNS(ns.html, tagName, props, children);
	}

	createMathMLElement<T extends keyof MathMLElementTagNameMap>(tagName: T, props?: Partial<Record<keyof MathMLElementTagNameMap[T], any>>, children?: any[]): MathMLElementTagNameMap[T] {
		return this.createElementNS(ns.mathML, tagName, props, children);
	}

	createSvgElement<T extends keyof SVGElementTagNameMap>(tagName: T, props?: Partial<Record<keyof SVGElementTagNameMap[T], any>>, children?: any[]): SVGElementTagNameMap[T] {
		return this.createElementNS(ns.svg, tagName, props, children);
	}

	later(func: Function) {
		this.postRenderTasks.push(func);
	}

	// Apply change bars to ancestor blocks of each rendered <ins>/<del>,
	// and inject the author legend. Runs once per render() after the tree
	// is built; see Track Changes Phase 1 (#3).
	private finalizeChangesRendering(result: Node[]) {
		const c = this.className;
		const opts = this.options.changes ?? {};

		if (opts.changeBar !== false) {
			for (const el of this.changeElements) {
				const block = this.findBlockAncestor(el);
				if (!block) continue;
				block.classList.add(`${c}-change-bar`);
				// Inherit the author colour so ::before uses `currentColor` to
				// draw the bar. We read the first author-index class; paragraphs
				// touched by multiple authors show the first one's colour.
				if (!block.style.color) {
					const match = Array.from(el.classList).find(n => n.startsWith(`${c}-change-author-`));
					if (match) block.classList.add(match);
				}
			}
		}

		if (opts.legend !== false && this.changeAuthorIndex.size > 0) {
			const legend = this.buildLegend();
			if (legend) {
				// Prefer inserting at the top of the wrapper so the legend sits
				// above the document when it's present; fall back to prepending
				// as a sibling of the first rendered element.
				const wrapper = this.findWrapper(result);
				if (wrapper) {
					wrapper.insertBefore(legend, wrapper.firstChild);
				} else if (result.length) {
					const insertAt = result.findIndex(n => n.nodeName !== "STYLE" && n.nodeType === 1);
					if (insertAt >= 0) result.splice(insertAt, 0, legend);
					else result.push(legend);
				}
			}
		}

		this.extendSidebarWithChanges();
	}

	// Appends a read-only revision card per unique change id to the comments
	// sidebar (when the sidebar is active and changes.sidebarCards isn't
	// disabled). Cards carry author, date, kind badge, and a short summary.
	// Clicking a card scrolls the document to the change.
	private extendSidebarWithChanges() {
		const c = this.className;
		const opts = this.options.changes ?? {};
		if (opts.sidebarCards === false) return;
		if (!this.useSidebar || !this.sidebarContainer) return;

		const content = this.sidebarContainer.querySelector(`.${c}-sidebar-content`);
		if (!content) return;

		// Only include top-level changes (each revision id appears once; for
		// moves that means a single card regardless of from/to halves).
		const seen = new Set<string>();
		const unique = this.changeMeta.filter(m => {
			if (!m.id || seen.has(m.id)) return false;
			seen.add(m.id);
			return true;
		});

		for (const meta of unique) {
			const card = this.buildRevisionCard(meta);
			content.appendChild(card);
			if (meta.id) this.revisionCardElements.set(meta.id, card);
		}
	}

	private buildRevisionCard(
		meta: { el: HTMLElement; id?: string; kind: ChangeKind; author?: string; date?: string; summary: string },
	): HTMLElement {
		const c = this.className;
		const opts = this.options.changes ?? {};

		const authorIdxClass = meta.author && opts.colorByAuthor !== false
			? `${c}-change-author-${this.getAuthorIndex(meta.author)}`
			: "";

		const headerChildren: Node[] = [
			this.h({ tagName: "span", className: `${c}-comment-author ${authorIdxClass}`, children: [meta.author ?? "Unknown"] }),
			this.h({ tagName: "span", className: `${c}-comment-date`, children: [meta.date ? new Date(meta.date).toLocaleString() : ""] }),
			this.h({ tagName: "span", className: `${c}-revision-kind`, children: [this.kindLabel(meta.kind)] }),
		];

		const card = this.h({
			tagName: "div",
			className: `${c}-sidebar-comment ${c}-revision-card`,
			children: [
				this.h({ tagName: "div", className: `${c}-comment-header`, children: headerChildren }),
				this.h({ tagName: "div", className: `${c}-comment-body`, children: [meta.summary] }),
			]
		}) as HTMLElement;

		card.addEventListener("click", () => {
			meta.el.scrollIntoView({ behavior: "smooth", block: "center" });
		});

		return card;
	}

	private kindLabel(kind: ChangeKind): string {
		switch (kind) {
			case "insertion": return "Inserted";
			case "deletion": return "Deleted";
			case "move": return "Moved";
			case "formatting": return "Formatted";
			case "paragraphMark": return "Paragraph mark";
			case "rowInsertion": return "Row added";
			case "rowDeletion": return "Row removed";
		}
	}

	private findBlockAncestor(el: HTMLElement): HTMLElement | null {
		let cur: HTMLElement | null = el.parentElement;
		while (cur) {
			const tag = cur.tagName;
			if (tag === "P" || tag === "LI" || tag === "TR" || tag === "H1" || tag === "H2" ||
				tag === "H3" || tag === "H4" || tag === "H5" || tag === "H6") {
				return cur;
			}
			if (tag === "SECTION" || tag === "BODY" || tag === "ARTICLE") return null;
			cur = cur.parentElement;
		}
		return null;
	}

	private findWrapper(result: Node[]): HTMLElement | null {
		const wrapperClass = `${this.className}-wrapper`;
		for (const node of result) {
			if (node instanceof HTMLElement && node.classList.contains(wrapperClass)) {
				return node;
			}
		}
		return null;
	}

	private buildLegend(): HTMLElement | null {
		const c = this.className;
		const items: Node[] = [
			this.h({ tagName: "span", className: `${c}-legend-label`, children: ["Changes by:"] })
		];
		const authors = [...this.changeAuthorIndex.entries()].sort((a, b) => a[1] - b[1]);
		for (const [author, idx] of authors) {
			items.push(this.h({
				tagName: "span",
				className: `${c}-legend-item`,
				children: [
					this.h({ tagName: "span", className: `${c}-legend-swatch ${c}-change-author-${idx}`, style: { background: "currentColor" } }),
					author
				]
			}));
		}
		return this.h({ tagName: "div", className: `${c}-legend`, children: items }) as HTMLElement;
	}
}

function findParent<T extends OpenXmlElement>(elem: OpenXmlElement, type: DomType): T {
	var parent = elem.parent;

	while (parent != null && parent.type != type)
		parent = parent.parent;

	return <T>parent;
}

