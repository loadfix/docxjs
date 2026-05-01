import { WordDocument } from './word-document';
import {
	DomType, WmlTable, IDomNumbering,
	WmlHyperlink, IDomImage, OpenXmlElement, WmlTableColumn, WmlTableCell, WmlText, WmlSymbol, WmlBreak, WmlNoteReference,
	WmlSmartTag,
	WmlAltChunk,
	WmlTableRow
} from './document/dom';
import { Options } from './docx-preview';
import { DocumentElement } from './document/document';
import { WmlParagraph } from './document/paragraph';
import { asArray, encloseFontFamily, escapeClassName, isString, keyBy, mergeDeep } from './utils';
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
import { CommentEventCallbacks } from './docx-preview';

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
		kind: import('./docx-preview').ChangeKind;
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

	get isReadOnly(): boolean {
		return this.options.comments?.readOnly !== false;
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
		const variables = {};
		const fontScheme = themePart.theme?.fontScheme;

		if (fontScheme) {
			if (fontScheme.majorFont) {
				variables['--docx-majorHAnsi-font'] = fontScheme.majorFont.latinTypeface;
			}

			if (fontScheme.minorFont) {
				variables['--docx-minorHAnsi-font'] = fontScheme.minorFont.latinTypeface;
			}
		}

		const colorScheme = themePart.theme?.colorScheme;

		if (colorScheme) {
			for (let [k, v] of Object.entries(colorScheme.colors)) {
				variables[`--docx-${k}-color`] = `#${v}`;
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

		this.sidebarContainer = this.h({ tagName: "div", className: `${c}-comment-sidebar` }) as HTMLElement;

		const toggleBtn = this.h({
			tagName: "button",
			className: `${c}-sidebar-toggle`,
			children: ["Comments"],
			title: "Toggle comments sidebar"
		}) as HTMLButtonElement;

		const highlightToggle = this.useHighlight ? this.h({
			tagName: "label",
			className: `${c}-highlight-toggle`,
			children: [
				this.h({ tagName: "input", type: "checkbox", checked: true }) as HTMLInputElement,
				" Highlight"
			]
		}) as HTMLElement : null;

		const toolbarChildren: Node[] = [toggleBtn];
		if (highlightToggle) toolbarChildren.push(highlightToggle);

		if (!this.isReadOnly) {
			const addBtn = this.h({
				tagName: "button",
				className: `${c}-comment-add-btn`,
				children: ["+ Comment"],
				title: "Add a comment on selected text"
			}) as HTMLButtonElement;
			toolbarChildren.push(addBtn);

			this.later(() => {
				addBtn.addEventListener("click", () => {
					const sel = document.getSelection();
					if (!sel || sel.isCollapsed) return;
					const range = sel.getRangeAt(0).cloneRange();
					if (!docContainer.contains(range.commonAncestorContainer)) return;
					this.showNewCommentComposer(contentArea, range);
				});
			});
		}

		const toolbar = this.h({
			tagName: "div",
			className: `${c}-comment-toolbar`,
			children: toolbarChildren
		}) as HTMLElement;

		const contentArea = this.h({
			tagName: "div",
			className: `${c}-sidebar-content`,
			children: []
		}) as HTMLElement;

		this.sidebarContainer.appendChild(toolbar);
		this.sidebarContainer.appendChild(contentArea);

		this.renderSidebarComments(contentArea);

		const wrapper = this.h({
			tagName: "div",
			className: `${c}-wrapper`,
			children: [docContainer, this.sidebarContainer]
		}) as HTMLElement;

		this.later(() => {
			toggleBtn.addEventListener("click", () => {
				this.sidebarContainer.classList.toggle(`${c}-sidebar-collapsed`);
			});

			if (highlightToggle) {
				const checkbox = highlightToggle.querySelector("input") as HTMLInputElement;
				checkbox.addEventListener("change", () => {
					if (checkbox.checked) {
						if (this.commentHighlight) {
							(CSS as any).highlights.set(`${c}-comments`, this.commentHighlight);
						}
						docContainer.classList.remove(`${c}-no-highlight`);
					} else {
						(CSS as any).highlights?.delete(`${c}-comments`);
						docContainer.classList.add(`${c}-no-highlight`);
					}
				});
			}

			this.setupSidebarScrollSync(docContainer, contentArea);
		});

		return wrapper;
	}

	setupSidebarScrollSync(docContainer: HTMLElement, sidebarContent: HTMLElement) {
		const wrapper = docContainer.parentElement;
		if (!wrapper) return;

		const CARD_GAP = 8;

		const positionComments = () => {
			const sidebarRect = sidebarContent.getBoundingClientRect();
			const ordered: Array<{ el: HTMLElement; desiredTop: number }> = [];

			for (const [commentId, sidebarEl] of Object.entries(this.sidebarCommentElements)) {
				if (!sidebarEl.isConnected) continue;
				const anchors = this.commentAnchorElements[commentId];
				const firstAnchor = anchors?.[0];
				if (!firstAnchor || !firstAnchor.isConnected) continue;

				const anchorRect = firstAnchor.getBoundingClientRect();
				const desiredTop = anchorRect.top - sidebarRect.top + sidebarContent.scrollTop;
				ordered.push({ el: sidebarEl, desiredTop });
			}

			// Sort by desired vertical position so cards stack in document order even
			// if the comments dictionary insertion order disagrees.
			ordered.sort((a, b) => a.desiredTop - b.desiredTop);

			// Clear prior offsets so offsetTop reflects the natural flow position.
			for (const { el } of ordered) el.style.marginTop = "";

			let floor = -Infinity;
			for (const { el, desiredTop } of ordered) {
				const target = Math.max(desiredTop, floor);
				const naturalTop = el.offsetTop;
				const offset = target - naturalTop;
				if (offset > 0) el.style.marginTop = `${offset}px`;
				floor = target + el.offsetHeight + CARD_GAP;
			}
		};

		let rafId: number;
		const throttledPosition = () => {
			cancelAnimationFrame(rafId);
			rafId = requestAnimationFrame(positionComments);
		};

		wrapper.addEventListener("scroll", throttledPosition, { passive: true });
		docContainer.addEventListener("scroll", throttledPosition, { passive: true });

		if (typeof ResizeObserver !== "undefined") {
			const ro = new ResizeObserver(throttledPosition);
			ro.observe(docContainer);
			for (const el of Object.values(this.sidebarCommentElements)) {
				if (el.isConnected) ro.observe(el);
			}
		}

		setTimeout(positionComments, 100);
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
		const callbacks = this.options.commentCallbacks ?? {};
		const readOnly = this.isReadOnly;

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

		let replyContainerRef: HTMLElement = null;

		if (!readOnly) {
			const actionsEl = this.h({
				tagName: "div",
				className: `${c}-comment-actions`,
				children: []
			}) as HTMLElement;

			const editBtn = this.h({ tagName: "button", className: `${c}-comment-edit-btn`, children: ["Edit"] }) as HTMLButtonElement;
			const deleteBtn = this.h({ tagName: "button", className: `${c}-comment-delete-btn`, children: ["Delete"] }) as HTMLButtonElement;
			actionsEl.appendChild(editBtn);
			actionsEl.appendChild(deleteBtn);

			let replyBtn: HTMLButtonElement = null;
			if (!isReply) {
				replyBtn = this.h({ tagName: "button", className: `${c}-comment-reply-btn`, children: ["Reply"] }) as HTMLButtonElement;
				actionsEl.appendChild(replyBtn);
			}

			children.push(actionsEl);

			this.later(() => {
				editBtn.addEventListener("click", (ev) => {
					ev.stopPropagation();
					this.openInlineEditor(bodyEl, bodyEl.textContent ?? "", (newText) => {
						callbacks.onCommentEdit?.(comment.id, newText);
					});
				});

				deleteBtn.addEventListener("click", (ev) => {
					ev.stopPropagation();
					this.openInlineConfirm(actionsEl, "Delete this comment?", () => {
						callbacks.onCommentDelete?.(comment.id);
					});
				});

				if (replyBtn) {
					replyBtn.addEventListener("click", (ev) => {
						ev.stopPropagation();
						const host = replyContainerRef ?? (() => {
							const el = this.h({ tagName: "div", className: `${c}-comment-replies` }) as HTMLElement;
							commentEl.insertBefore(el, null);
							replyContainerRef = el;
							return el;
						})();
						this.openReplyComposer(host, (text) => {
							callbacks.onCommentReply?.(comment.id, text);
						});
					});
				}
			});
		}

		if (comment.replies && comment.replies.length > 0) {
			const repliesContainer = this.h({
				tagName: "div",
				className: `${c}-comment-replies`,
				children: comment.replies.map(r => this.renderSidebarComment(r, true))
			}) as HTMLElement;

			replyContainerRef = repliesContainer;

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

	private openInlineEditor(bodyEl: HTMLElement, currentText: string, onSave: (text: string) => void) {
		const c = this.className;
		if (bodyEl.querySelector(`.${c}-comment-editor`)) return;

		const originalContent = Array.from(bodyEl.childNodes);
		const textarea = this.h({ tagName: "textarea", className: `${c}-comment-editor` }) as HTMLTextAreaElement;
		textarea.value = currentText;

		const save = this.h({ tagName: "button", className: `${c}-comment-editor-save`, children: ["Save"] }) as HTMLButtonElement;
		const cancel = this.h({ tagName: "button", className: `${c}-comment-editor-cancel`, children: ["Cancel"] }) as HTMLButtonElement;
		const actions = this.h({ tagName: "div", className: `${c}-comment-editor-actions`, children: [save, cancel] }) as HTMLElement;

		bodyEl.replaceChildren(textarea, actions);
		textarea.focus();
		textarea.select();

		const restore = () => bodyEl.replaceChildren(...originalContent);

		save.addEventListener("click", (ev) => {
			ev.stopPropagation();
			const next = textarea.value;
			if (next !== currentText) onSave(next);
			restore();
		});
		cancel.addEventListener("click", (ev) => {
			ev.stopPropagation();
			restore();
		});
		textarea.addEventListener("click", (ev) => ev.stopPropagation());
		textarea.addEventListener("keydown", (ev) => {
			if (ev.key === "Escape") { ev.preventDefault(); restore(); }
			if (ev.key === "Enter" && (ev.metaKey || ev.ctrlKey)) { ev.preventDefault(); save.click(); }
		});
	}

	private openInlineConfirm(hostEl: HTMLElement, message: string, onConfirm: () => void) {
		const c = this.className;
		if (hostEl.querySelector(`.${c}-comment-confirm`)) return;

		const msg = this.h({ tagName: "span", className: `${c}-comment-confirm-msg`, children: [message] }) as HTMLElement;
		const yes = this.h({ tagName: "button", className: `${c}-comment-confirm-yes`, children: ["Yes"] }) as HTMLButtonElement;
		const no = this.h({ tagName: "button", className: `${c}-comment-confirm-no`, children: ["No"] }) as HTMLButtonElement;
		const wrap = this.h({ tagName: "div", className: `${c}-comment-confirm`, children: [msg, yes, no] }) as HTMLElement;

		hostEl.appendChild(wrap);

		yes.addEventListener("click", (ev) => {
			ev.stopPropagation();
			wrap.remove();
			onConfirm();
		});
		no.addEventListener("click", (ev) => {
			ev.stopPropagation();
			wrap.remove();
		});
	}

	private openReplyComposer(hostEl: HTMLElement, onSubmit: (text: string) => void) {
		const c = this.className;
		if (hostEl.querySelector(`.${c}-comment-reply-composer`)) return;

		const textarea = this.h({ tagName: "textarea", className: `${c}-comment-editor` }) as HTMLTextAreaElement;
		textarea.placeholder = "Write a reply...";
		const submit = this.h({ tagName: "button", className: `${c}-comment-editor-save`, children: ["Reply"] }) as HTMLButtonElement;
		const cancel = this.h({ tagName: "button", className: `${c}-comment-editor-cancel`, children: ["Cancel"] }) as HTMLButtonElement;
		const actions = this.h({ tagName: "div", className: `${c}-comment-editor-actions`, children: [submit, cancel] }) as HTMLElement;
		const composer = this.h({ tagName: "div", className: `${c}-comment-reply-composer`, children: [textarea, actions] }) as HTMLElement;

		hostEl.appendChild(composer);
		textarea.focus();

		submit.addEventListener("click", (ev) => {
			ev.stopPropagation();
			const text = textarea.value.trim();
			if (text) onSubmit(text);
			composer.remove();
		});
		cancel.addEventListener("click", (ev) => {
			ev.stopPropagation();
			composer.remove();
		});
		textarea.addEventListener("click", (ev) => ev.stopPropagation());
		textarea.addEventListener("keydown", (ev) => {
			if (ev.key === "Escape") { ev.preventDefault(); composer.remove(); }
			if (ev.key === "Enter" && (ev.metaKey || ev.ctrlKey)) { ev.preventDefault(); submit.click(); }
		});
	}

	private showNewCommentComposer(contentArea: HTMLElement, range: Range) {
		const c = this.className;
		const existing = contentArea.querySelector(`.${c}-new-comment-composer`);
		if (existing) existing.remove();

		const textarea = this.h({ tagName: "textarea", className: `${c}-comment-editor` }) as HTMLTextAreaElement;
		textarea.placeholder = "Write a comment on the selected text...";
		const submit = this.h({ tagName: "button", className: `${c}-comment-editor-save`, children: ["Add"] }) as HTMLButtonElement;
		const cancel = this.h({ tagName: "button", className: `${c}-comment-editor-cancel`, children: ["Cancel"] }) as HTMLButtonElement;
		const actions = this.h({ tagName: "div", className: `${c}-comment-editor-actions`, children: [submit, cancel] }) as HTMLElement;
		const composer = this.h({ tagName: "div", className: `${c}-new-comment-composer`, children: [textarea, actions] }) as HTMLElement;

		contentArea.insertBefore(composer, contentArea.firstChild);
		textarea.focus();

		submit.addEventListener("click", () => {
			const text = textarea.value.trim();
			if (text) this.options.commentCallbacks?.onCommentAdd?.(range, text);
			composer.remove();
		});
		cancel.addEventListener("click", () => composer.remove());
		textarea.addEventListener("keydown", (ev) => {
			if (ev.key === "Escape") { ev.preventDefault(); composer.remove(); }
			if (ev.key === "Enter" && (ev.metaKey || ev.ctrlKey)) { ev.preventDefault(); submit.click(); }
		});
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
.${c}-comment-sidebar { width: 320px; min-width: 260px; background: #fafafa; border-left: 1px solid #ddd; display: flex; flex-direction: column; position: sticky; top: 0; height: 100vh; overflow: hidden; transition: width 0.2s, min-width 0.2s, padding 0.2s; }
.${c}-sidebar-collapsed { width: 0 !important; min-width: 0 !important; padding: 0 !important; border: none !important; overflow: hidden; }
.${c}-sidebar-collapsed .${c}-sidebar-content,
.${c}-sidebar-collapsed .${c}-comment-toolbar > *:not(.${c}-sidebar-toggle) { display: none; }
.${c}-comment-toolbar { display: flex; align-items: center; gap: 8px; padding: 8px 12px; border-bottom: 1px solid #ddd; background: #f5f5f5; flex-shrink: 0; flex-wrap: wrap; }
.${c}-sidebar-toggle { cursor: pointer; background: #fff; border: 1px solid #ccc; border-radius: 4px; padding: 4px 10px; font-size: 0.8rem; }
.${c}-sidebar-toggle:hover { background: #e8e8e8; }
.${c}-highlight-toggle { font-size: 0.8rem; display: flex; align-items: center; gap: 4px; cursor: pointer; white-space: nowrap; }
.${c}-comment-add-btn { cursor: pointer; background: #4a90d9; color: white; border: none; border-radius: 4px; padding: 4px 10px; font-size: 0.8rem; }
.${c}-comment-add-btn:hover { background: #357abd; }
.${c}-sidebar-content { flex: 1; overflow-y: auto; padding: 8px; }
.${c}-sidebar-comment { background: white; border: 1px solid #e0e0e0; border-radius: 6px; padding: 10px; margin-bottom: 8px; cursor: pointer; transition: box-shadow 0.2s, border-color 0.2s; }
.${c}-sidebar-comment:hover { border-color: #4a90d9; box-shadow: 0 1px 4px rgba(74, 144, 217, 0.2); }
.${c}-sidebar-reply { margin-left: 16px; border-left: 3px solid #4a90d9; background: #f8fbff; }
.${c}-comment-header { display: flex; align-items: baseline; gap: 8px; margin-bottom: 4px; flex-wrap: wrap; }
.${c}-comment-author { font-weight: 600; font-size: 0.85rem; color: #333; }
.${c}-comment-date { font-size: 0.75rem; color: #999; }
.${c}-comment-done { font-size: 0.7rem; background: #4caf50; color: white; padding: 1px 6px; border-radius: 3px; }
.${c}-comment-body { font-size: 0.85rem; color: #444; margin-bottom: 6px; line-height: 1.4; }
.${c}-comment-body p { margin: 2px 0; }
.${c}-comment-actions { display: flex; gap: 6px; }
.${c}-comment-actions button { background: none; border: 1px solid #ddd; border-radius: 3px; padding: 2px 8px; font-size: 0.75rem; cursor: pointer; color: #666; }
.${c}-comment-actions button:hover { background: #f0f0f0; border-color: #bbb; }
.${c}-comment-delete-btn:hover { color: #d32f2f !important; border-color: #d32f2f !important; }
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
.${c}-comment-editor { width: 100%; min-height: 60px; box-sizing: border-box; font: inherit; font-size: 0.85rem; padding: 6px; border: 1px solid #bbb; border-radius: 4px; resize: vertical; }
.${c}-comment-editor-actions { display: flex; gap: 6px; margin-top: 6px; }
.${c}-comment-editor-actions button { background: none; border: 1px solid #ddd; border-radius: 3px; padding: 2px 10px; font-size: 0.75rem; cursor: pointer; color: #666; }
.${c}-comment-editor-save { background: #4a90d9 !important; color: white !important; border-color: #4a90d9 !important; }
.${c}-comment-editor-save:hover { background: #357abd !important; }
.${c}-comment-confirm { display: flex; align-items: center; gap: 6px; margin-left: auto; font-size: 0.75rem; color: #d32f2f; }
.${c}-comment-confirm button { background: none; border: 1px solid #ddd; border-radius: 3px; padding: 2px 8px; font-size: 0.75rem; cursor: pointer; }
.${c}-comment-confirm-yes { color: white !important; background: #d32f2f !important; border-color: #d32f2f !important; }
.${c}-new-comment-composer,.${c}-comment-reply-composer { background: white; border: 1px solid #4a90d9; border-radius: 6px; padding: 10px; margin-bottom: 8px; }
.${c}-comment-reply-composer { margin: 6px 0; }
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
.${c}-change-actions { display: none; margin-left: 4px; user-select: none; vertical-align: baseline; font-size: 0.75em; }
.${c}-change-actions button { border: 1px solid currentColor; background: white; color: currentColor; cursor: pointer; padding: 0 4px; border-radius: 3px; line-height: 1; margin-right: 2px; }
.${c}-change-actions button:hover { background: currentColor; color: white; }
.${c} ins:hover > .${c}-change-actions,
.${c} del:hover > .${c}-change-actions,
.${c} .${c}-move-from:hover > .${c}-change-actions,
.${c} .${c}-move-to:hover > .${c}-change-actions,
.${c} .${c}-formatting-revision:hover > .${c}-change-actions,
.${c}-paragraph-mark:hover > .${c}-change-actions { display: inline-flex; gap: 2px; }
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
			var selector = `p.${this.numberingClass(num.id, num.level)}`;
			var listStyleType = "none";

			if (num.bullet) {
				let valiable = `--${this.className}-${num.bullet.src}`.toLowerCase();

				styleText += this.styleToString(`${selector}:before`, {
					"content": "' '",
					"display": "inline-block",
					"background": `var(${valiable})`
				}, num.bullet.style);

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
				return this.renderAltChunk(elem);
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
		res.href = '';

		if (elem.id) {
			const rel = this.document.documentPart.rels.find(it => it.id == elem.id && it.targetMode === "External");
			res.href = rel?.target ?? res.href;
		}

		if (elem.anchor) {
			res.href += `#${elem.anchor}`;
		}

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

	renderAltChunk(elem: WmlAltChunk) {
		if (!this.options.renderAltChunks)
			return null;

		var result = this.h({ tagName: "iframe" }) as HTMLIFrameElement;
		
		this.tasks.push(this.document.loadAltChunk(elem.id, this.currentPart).then(x => {
			result.srcdoc = x;
		}));

		return result;
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
	applyChangeAttributes(node: HTMLElement, elem: OpenXmlElement, kind: import('./docx-preview').ChangeKind) {
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

	private summarizeChange(node: HTMLElement, kind: import('./docx-preview').ChangeKind): string {
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
		const metaKind: import('./docx-preview').ChangeKind = kind === "inserted" ? "rowInsertion" : "rowDeletion";
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
		const suffMap = {
			"tab": "\\9",
			"space": "\\a0",
		};

		var result = text.replace(/%\d*/g, s => {
			let lvl = parseInt(s.substring(1), 10) - 1;
			return `"counter(${this.numberingCounter(id, lvl)}, ${numformat})"`;
		});

		return `"${result}${suffMap[suff] ?? ""}"`;
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

		return mapping[format] ?? format;
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

		if (opts.readOnly === false) {
			this.wireChangeActionDelegate(result);
			this.injectInlineChangeActions();
			this.extendSidebarWithChanges();
		}
	}

	// Inject ✓/✕ buttons into each change element with an id. Uses a single
	// delegated listener attached to the wrapper; this avoids N listeners on a
	// large document and keeps us resilient to elements being moved by CSS.
	private injectInlineChangeActions() {
		const c = this.className;
		for (const meta of this.changeMeta) {
			if (!meta.id) continue;
			if (meta.el.querySelector(`.${c}-change-actions`)) continue;
			// Row revisions are <tr> — inserting <span> children is invalid.
			// Skip inline buttons for rows; the sidebar card still works.
			if (meta.el.tagName === "TR") continue;

			const accept = this.h({
				tagName: "button",
				className: `${c}-change-accept`,
				children: ["✓"],
				title: "Accept change"
			}) as HTMLButtonElement;
			const reject = this.h({
				tagName: "button",
				className: `${c}-change-reject`,
				children: ["✕"],
				title: "Reject change"
			}) as HTMLButtonElement;
			const wrap = this.h({
				tagName: "span",
				className: `${c}-change-actions`,
				children: [accept, reject]
			}) as HTMLElement;
			meta.el.appendChild(wrap);
		}
	}

	private wireChangeActionDelegate(result: Node[]) {
		const c = this.className;
		const wrapper = this.findWrapper(result);
		const root = wrapper ?? this.findFirstElementRoot(result);
		if (!root) return;
		const callbacks = this.options.changeCallbacks ?? {};

		root.addEventListener("click", (ev) => {
			const target = ev.target as HTMLElement | null;
			if (!target) return;
			const btn = target.closest(`.${c}-change-accept, .${c}-change-reject`) as HTMLElement | null;
			if (!btn) return;

			// Find the owning change element via dataset.changeKind.
			const owner = btn.closest<HTMLElement>("[data-change-id][data-change-kind]");
			if (!owner) return;

			ev.preventDefault();
			ev.stopPropagation();

			const id = owner.dataset.changeId!;
			const kind = owner.dataset.changeKind as import('./docx-preview').ChangeKind;
			if (btn.classList.contains(`${c}-change-accept`)) {
				callbacks.onChangeAccept?.(id, kind);
			} else {
				callbacks.onChangeReject?.(id, kind);
			}
		});
	}

	// Locates the first element node in the render output so we can delegate
	// events from its subtree when there's no wrapper.
	private findFirstElementRoot(result: Node[]): HTMLElement | null {
		for (const n of result) {
			if (n instanceof HTMLElement) return n;
		}
		return null;
	}

	// Adds revision cards to the comments sidebar (if active) and wires up
	// the "Accept all" / "Reject all" toolbar buttons.
	private extendSidebarWithChanges() {
		const c = this.className;
		const opts = this.options.changes ?? {};
		if (opts.sidebarCards === false) return;
		if (!this.useSidebar || !this.sidebarContainer) return;

		const content = this.sidebarContainer.querySelector(`.${c}-sidebar-content`);
		const toolbar = this.sidebarContainer.querySelector(`.${c}-comment-toolbar`);
		if (!content) return;

		// Only include top-level changes (each revision id appears once; for
		// moves that means a single card regardless of from/to halves).
		const seen = new Set<string>();
		const unique = this.changeMeta.filter(m => {
			if (!m.id || seen.has(m.id)) return false;
			seen.add(m.id);
			return true;
		});
		const callbacks = this.options.changeCallbacks ?? {};

		for (const meta of unique) {
			content.appendChild(this.buildRevisionCard(meta, callbacks));
		}

		if (toolbar && opts.readOnly === false) {
			const acceptAll = this.h({
				tagName: "button",
				className: `${c}-sidebar-toggle`,
				children: ["Accept all"]
			}) as HTMLButtonElement;
			const rejectAll = this.h({
				tagName: "button",
				className: `${c}-sidebar-toggle`,
				children: ["Reject all"]
			}) as HTMLButtonElement;
			toolbar.appendChild(acceptAll);
			toolbar.appendChild(rejectAll);
			acceptAll.addEventListener("click", () => callbacks.onChangeAcceptAll?.());
			rejectAll.addEventListener("click", () => callbacks.onChangeRejectAll?.());
		}
	}

	private buildRevisionCard(
		meta: { el: HTMLElement; id?: string; kind: import('./docx-preview').ChangeKind; author?: string; date?: string; summary: string },
		callbacks: import('./docx-preview').ChangeEventCallbacks,
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

		const body = this.h({
			tagName: "div",
			className: `${c}-comment-body`,
			children: [meta.summary]
		}) as HTMLElement;

		const children: Node[] = [
			this.h({ tagName: "div", className: `${c}-comment-header`, children: headerChildren }),
			body,
		];

		if (opts.readOnly === false && meta.id) {
			const accept = this.h({ tagName: "button", children: ["Accept"] }) as HTMLButtonElement;
			const reject = this.h({ tagName: "button", children: ["Reject"] }) as HTMLButtonElement;
			accept.addEventListener("click", (ev) => {
				ev.stopPropagation();
				callbacks.onChangeAccept?.(meta.id!, meta.kind);
			});
			reject.addEventListener("click", (ev) => {
				ev.stopPropagation();
				callbacks.onChangeReject?.(meta.id!, meta.kind);
			});
			children.push(this.h({
				tagName: "div",
				className: `${c}-comment-actions`,
				children: [accept, reject]
			}));
		}

		const card = this.h({
			tagName: "div",
			className: `${c}-sidebar-comment ${c}-revision-card`,
			children
		}) as HTMLElement;

		card.addEventListener("click", () => {
			meta.el.scrollIntoView({ behavior: "smooth", block: "center" });
		});

		return card;
	}

	private kindLabel(kind: import('./docx-preview').ChangeKind): string {
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

