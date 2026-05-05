import {
	DomType, WmlTable, IDomNumbering,
	WmlHyperlink, WmlSmartTag, IDomImage, OpenXmlElement, WmlTableColumn, WmlTableCell,
	WmlTableRow, NumberingPicBullet, WmlText, WmlSymbol, WmlBreak, WmlNoteReference,
	WmlAltChunk, Revision, FormattingRevision, WmlSdt, SdtControl, SdtCheckboxControl,
	WmlRuby, WmlFitText, WmlBidiOverride
} from './document/dom';
import {
	DrawingShape, DrawingGroup, DrawingChart, DrawingChartEx, DrawingSmartArt,
	GradientFill, PatternFill, ShapeEffects,
} from './document/drawing';
import { sanitizeCssColor } from './utils';
import { ColourRef, ColourModifiers, isAllowedSchemeSlot, DEFAULT_THEME_PALETTE, buildThemeColorReference } from './drawing/theme';
import { DocumentElement } from './document/document';
import { WmlParagraph, parseParagraphProperties, parseParagraphProperty } from './document/paragraph';
import { parseSectionProperties, SectionProperties } from './document/section';
import xml from './parser/xml-parser';
import { parseRunProperties, WmlRun } from './document/run';
import { parseBookmarkEnd, parseBookmarkStart } from './document/bookmarks';
import { IDomStyle, IDomSubStyle } from './document/style';
import { WmlFieldChar, WmlFieldSimple, WmlInstructionText } from './document/fields';
import { convertLength, LengthUsage, LengthUsageType } from './document/common';
import { parseVmlElement } from './vml/vml';
import { WmlComment, WmlCommentRangeEnd, WmlCommentRangeStart, WmlCommentReference } from './comments/elements';
import { encloseFontFamily } from './utils';

function parseRevisionAttrs(elem: Element): Revision {
	return {
		id: xml.attr(elem, "id"),
		author: xml.attr(elem, "author"),
		date: xml.attr(elem, "date")
	};
}

// Produces a short, bounded list of changed property names from a w:rPrChange
// or w:pPrChange element's previous-properties child (w:rPr / w:pPr).
// We cap at 5 entries to keep the rendered title attribute short.
const FORMATTING_PROP_NAMES: Record<string, string> = {
	b: "bold", i: "italic", u: "underline", strike: "strikethrough",
	sz: "font size", rFonts: "font", color: "color", highlight: "highlight",
	jc: "alignment", ind: "indent", spacing: "spacing", numPr: "numbering",
	pStyle: "style", rStyle: "style"
};

function parseFormattingRevision(elem: Element): FormattingRevision {
	const rev: FormattingRevision = {
		id: xml.attr(elem, "id"),
		author: xml.attr(elem, "author"),
		date: xml.attr(elem, "date"),
		changedProps: []
	};
	const prev = xml.elements(elem).find(e => e.localName === "rPr" || e.localName === "pPr");
	if (prev) {
		const seen = new Set<string>();
		for (const child of xml.elements(prev)) {
			const pretty = FORMATTING_PROP_NAMES[child.localName] ?? child.localName;
			if (seen.has(pretty)) continue;
			seen.add(pretty);
			rev.changedProps.push(pretty);
			if (rev.changedProps.length >= 5) break;
		}
	}
	return rev;
}

// Exported for unit-testing the null-attribute guard added for upstream
// issue #196 — when a `<w:cnfStyle/>` element appears without a `w:val`
// attribute, `xml.attr` returns `null` and the previous indexing into that
// `null` crashed the parser. The guard returns '' so callers get an empty
// className instead.
export function classNameOfCnfStyle(c: Element): string {
	const val = xml.attr(c, "val");
	if (!val) return '';
	const classes = [
		'first-row', 'last-row', 'first-col', 'last-col',
		'odd-col', 'even-col', 'odd-row', 'even-row',
		'ne-cell', 'nw-cell', 'se-cell', 'sw-cell'
	];

	return classes.filter((_, i) => val[i] == '1').join(' ');
}

export var autos = {
	shd: "inherit",
	color: "black",
	borderColor: "black",
	highlight: "transparent"
};

const supportedNamespaceURIs = [];

// Walks up parentNode links looking for an Element with a matching
// localName. Stops at the document root. Used by parseSmartArtReference
// to find an enclosing <mc:AlternateContent> without threading it
// through every caller.
function findAncestorByLocalName(start: Element | null, localName: string): Element | null {
	let n: Node | null = start ? start.parentNode : null;
	while (n && n.nodeType === 1) {
		if ((n as Element).localName === localName) return n as Element;
		n = n.parentNode;
	}
	return null;
}

const mmlTagMap = {
	"oMath": DomType.MmlMath,
	"oMathPara": DomType.MmlMathParagraph,
	"f": DomType.MmlFraction,
	"func": DomType.MmlFunction,
	"fName": DomType.MmlFunctionName,
	"num": DomType.MmlNumerator,
	"den": DomType.MmlDenominator,
	"rad": DomType.MmlRadical,
	"deg": DomType.MmlDegree,
	"e": DomType.MmlBase,
	"sSup": DomType.MmlSuperscript,
	"sSub": DomType.MmlSubscript,
	"sPre": DomType.MmlPreSubSuper,
	"sup": DomType.MmlSuperArgument,
	"sub": DomType.MmlSubArgument,
	"d": DomType.MmlDelimiter,
	"nary": DomType.MmlNary,
	"eqArr": DomType.MmlEquationArray,
	"lim": DomType.MmlLimit,
	"limLow": DomType.MmlLimitLower,
	"m": DomType.MmlMatrix,
	"mr": DomType.MmlMatrixRow,
	"box": DomType.MmlBox,
	"bar": DomType.MmlBar,
	"groupChr": DomType.MmlGroupChar,
	"acc": DomType.MmlAccent,
	"borderBox": DomType.MmlBorderBox,
	"sSubSup": DomType.MmlSubSuperscript,
	"phant": DomType.MmlPhantom,
	"sGroup": DomType.MmlGroup
}

export interface DocumentParserOptions {
	ignoreWidth: boolean;
	debug: boolean;
}

export class DocumentParser {
	options: DocumentParserOptions;

	constructor(options?: Partial<DocumentParserOptions>) {
		this.options = {
			ignoreWidth: false,
			debug: false,
			...options
		};
	}

	parseNotes(xmlDoc: Element, elemName: string, elemClass: any): any[] {
		var result = [];

		for (let el of xml.elements(xmlDoc, elemName)) {
			const node = new elemClass();
			node.id = xml.attr(el, "id");
			node.noteType = xml.attr(el, "type");
			node.children = this.parseBodyElements(el);
			result.push(node);
		}

		return result;
	}

	parseComments(xmlDoc: Element): any[] {
		var result = [];

		for (let el of xml.elements(xmlDoc, "comment")) {
			const item = new WmlComment();
			item.id = xml.attr(el, "id");
			item.author = xml.attr(el, "author");
			item.initials = xml.attr(el, "initials");
			item.date = xml.attr(el, "date");

			// Parse paraId for threading support (w14:paraId or w:paraId)
			const paraId = el.getAttributeNS("http://schemas.microsoft.com/office/word/2010/wordml", "paraId")
				?? el.getAttribute("w14:paraId")
				?? xml.attr(el, "paraId");
			if (paraId) {
				item.paraId = paraId;
			}

			item.children = this.parseBodyElements(el);
			result.push(item);
		}

		return result;
	}

	parseDocumentFile(xmlDoc: Element): DocumentElement {
		var xbody = xml.element(xmlDoc, "body");
		var background = xml.element(xmlDoc, "background");
		var sectPr = xml.element(xbody, "sectPr");

		return {
			type: DomType.Document,
			children: this.parseBodyElements(xbody),
			props: sectPr ? parseSectionProperties(sectPr, xml) : {} as SectionProperties,
			cssStyle: background ? this.parseBackground(background) : {},
		};
	}

	parseBackground(elem: Element): any {
		var result = {};
		var color = xmlUtil.colorAttr(elem, "color");

		if (color) {
			result["background-color"] = color;
		}

		return result;
	}

	parseBodyElements(element: Element): OpenXmlElement[] {
		var children = [];

		for (const elem of xml.elements(element)) {
			switch (elem.localName) {
				case "p":
					children.push(this.parseParagraph(elem));
					break;

				case "altChunk":
					children.push(this.parseAltChunk(elem));
					break;

				case "tbl":
					children.push(this.parseTable(elem));
					break;

				case "sdt":
					children.push(...this.parseSdt(elem, e => this.parseBodyElements(e)));
					break;
			}
		}

		return children;
	}

	parseStylesFile(xstyles: Element): IDomStyle[] {
		var result = [];

		for (const n of xml.elements(xstyles)) {
			switch (n.localName) {
				case "style":
					result.push(this.parseStyle(n));
					break;

				case "docDefaults":
					result.push(this.parseDefaultStyles(n));
					break;
			}
		}

		return result;
	}

	parseDefaultStyles(node: Element): IDomStyle {
		var result = <IDomStyle>{
			id: null,
			name: null,
			target: null,
			basedOn: null,
			styles: []
		};

		for (const c of xml.elements(node)){
			switch (c.localName) {
				case "rPrDefault":
					var rPr = xml.element(c, "rPr");

					if (rPr)
						result.styles.push({
							target: "span",
							values: this.parseDefaultProperties(rPr, {})
						});
					break;

				case "pPrDefault":
					var pPr = xml.element(c, "pPr");

					if (pPr)
						result.styles.push({
							target: "p",
							values: this.parseDefaultProperties(pPr, {})
						});
					break;
			}
		}

		return result;
	}

	parseStyle(node: Element): IDomStyle {
		var result = <IDomStyle>{
			id: xml.attr(node, "styleId"),
			isDefault: xml.boolAttr(node, "default"),
			name: null,
			target: null,
			basedOn: null,
			styles: [],
			linked: null
		};

		switch (xml.attr(node, "type")) {
			case "paragraph": result.target = "p"; break;
			case "table": result.target = "table"; break;
			case "character": result.target = "span"; break;
			//case "numbering": result.target = "p"; break;
		}

		for (const n of xml.elements(node)) {
			switch (n.localName) {
				case "basedOn":
					result.basedOn = xml.attr(n, "val");
					break;

				case "name":
					result.name = xml.attr(n, "val");
					break;

				case "link":
					result.linked = xml.attr(n, "val");
					break;

				case "next":
					result.next = xml.attr(n, "val");
					break;

				case "aliases":
					result.aliases = xml.attr(n, "val").split(",");
					break;

				case "pPr":
					result.styles.push({
						target: "p",
						values: this.parseDefaultProperties(n, {})
					});
					result.paragraphProps = parseParagraphProperties(n, xml);
					break;

				case "rPr":
					result.styles.push({
						target: "span",
						values: this.parseDefaultProperties(n, {})
					});
					result.runProps = parseRunProperties(n, xml);
					break;

				case "tblPr":
				case "tcPr":
					result.styles.push({
						target: "td", //TODO: maybe move to processor
						values: this.parseDefaultProperties(n, {})
					});
					break;

				case "tblStylePr":
					for (let s of this.parseTableStyle(n))
						result.styles.push(s);
					break;

				case "rsid":
				case "qFormat":
				case "hidden":
				case "semiHidden":
				case "unhideWhenUsed":
				case "autoRedefine":
				case "uiPriority":
					//TODO: ignore
					break;

				default:
					this.options.debug && console.warn(`DOCX: Unknown style element: ${n.localName}`);
			}
		}

		return result;
	}

	parseTableStyle(node: Element): IDomSubStyle[] {
		var result = [];

		var type = xml.attr(node, "type");
		var selector = "";
		var modificator = "";

		switch (type) {
			case "firstRow":
				modificator = ".first-row";
				selector = "tr.first-row td";
				break;
			case "lastRow":
				modificator = ".last-row";
				selector = "tr.last-row td";
				break;
			case "firstCol":
				modificator = ".first-col";
				selector = "td.first-col";
				break;
			case "lastCol":
				modificator = ".last-col";
				selector = "td.last-col";
				break;
			case "band1Vert":
				modificator = ":not(.no-vband)";
				selector = "td.odd-col";
				break;
			case "band2Vert":
				modificator = ":not(.no-vband)";
				selector = "td.even-col";
				break;
			case "band1Horz":
				modificator = ":not(.no-hband)";
				selector = "tr.odd-row";
				break;
			case "band2Horz":
				modificator = ":not(.no-hband)";
				selector = "tr.even-row";
				break;
			default: return [];
		}

		for (const n of xml.elements(node)) {
			switch (n.localName) {
				case "pPr":
					result.push({
						target: `${selector} p`,
						mod: modificator,
						values: this.parseDefaultProperties(n, {})
					});
					break;

				case "rPr":
					result.push({
						target: `${selector} span`,
						mod: modificator,
						values: this.parseDefaultProperties(n, {})
					});
					break;

				case "tblPr":
				case "tcPr":
					result.push({
						target: selector, //TODO: maybe move to processor
						mod: modificator,
						values: this.parseDefaultProperties(n, {})
					});
					break;
			}
		}

		return result;
	}

	parseNumberingFile(node: Element): IDomNumbering[] {
		// Group abstract-num levels by their abstractNumId so each <w:num>
		// can clone and optionally override them. The previous implementation
		// kept a single abstractNumId -> numId map, which silently dropped
		// every <w:num> but the last when several nums shared an abstract —
		// a pattern common in legal / spec documents where the second list
		// is supposed to restart its counter. Emitting one entry per
		// (numId, level) lets each num own a distinct CSS counter.
		var abstractLevels: Record<string, IDomNumbering[]> = {};
		var bullets = [];
		var numElements: Element[] = [];

		for (const n of xml.elements(node)) {
			switch (n.localName) {
				case "abstractNum":
					var absId = xml.attr(n, "abstractNumId");
					abstractLevels[absId] = this.parseAbstractNumbering(n, bullets);
					break;

				case "numPicBullet":
					bullets.push(this.parseNumberingPicBullet(n));
					break;

				case "num":
					// Defer until all abstractNums are collected (DOCX files
					// may list nums before abstracts, though this is rare).
					numElements.push(n);
					break;
			}
		}

		var result: IDomNumbering[] = [];
		for (const n of numElements) {
			var numId = xml.attr(n, "numId");
			var abstractNumId = xml.elementAttr(n, "abstractNumId", "val");
			var baseLevels = abstractLevels[abstractNumId];
			if (!baseLevels) continue;

			// Parse lvlOverride entries keyed by ilvl.
			var overrides: Record<number, { start?: number; level?: IDomNumbering }> = {};
			for (const child of xml.elements(n, "lvlOverride")) {
				var ilvl = xml.intAttr(child, "ilvl");
				var entry: { start?: number; level?: IDomNumbering } = {};
				for (const sub of xml.elements(child)) {
					if (sub.localName === "startOverride") {
						entry.start = xml.intAttr(sub, "val");
					} else if (sub.localName === "lvl") {
						entry.level = this.parseNumberingLevel(abstractNumId, sub, bullets);
					}
				}
				overrides[ilvl] = entry;
			}

			for (const base of baseLevels) {
				// Shallow-clone the level so each num owns its own copy.
				// pStyle / rStyle are nested objects the renderer mutates via
				// spread, so we clone those too.
				var clone: IDomNumbering = {
					...base,
					pStyle: { ...base.pStyle },
					rStyle: { ...base.rStyle },
					id: numId,
				};
				var ov = overrides[base.level];
				if (ov) {
					if (ov.level) {
						// Property-level override: any field present on the
						// override <w:lvl> replaces the abstract value.
						if (ov.level.start !== undefined) clone.start = ov.level.start;
						if (ov.level.levelText !== undefined) clone.levelText = ov.level.levelText;
						if (ov.level.format !== undefined) clone.format = ov.level.format;
						if (ov.level.suff !== undefined) clone.suff = ov.level.suff;
						if (ov.level.restart !== undefined) clone.restart = ov.level.restart;
						if (ov.level.justification !== undefined) clone.justification = ov.level.justification;
						if (ov.level.isLgl !== undefined) clone.isLgl = ov.level.isLgl;
						if (ov.level.bullet !== undefined) clone.bullet = ov.level.bullet;
						if (ov.level.pStyleName !== undefined) clone.pStyleName = ov.level.pStyleName;
						if (ov.level.pStyle && Object.keys(ov.level.pStyle).length) {
							clone.pStyle = { ...clone.pStyle, ...ov.level.pStyle };
						}
						if (ov.level.rStyle && Object.keys(ov.level.rStyle).length) {
							clone.rStyle = { ...clone.rStyle, ...ov.level.rStyle };
						}
					}
					// startOverride wins over any property-level start.
					if (ov.start !== undefined) clone.start = ov.start;
				}
				result.push(clone);
			}
		}

		return result;
	}

	parseNumberingPicBullet(elem: Element): NumberingPicBullet {
		var pict = xml.element(elem, "pict");
		var shape = pict && xml.element(pict, "shape");
		var imagedata = shape && xml.element(shape, "imagedata");

		return imagedata ? {
			id: xml.intAttr(elem, "numPicBulletId"),
			src: xml.attr(imagedata, "id"),
			style: xml.attr(shape, "style")
		} : null;
	}

	parseAbstractNumbering(node: Element, bullets: any[]): IDomNumbering[] {
		var result = [];
		var id = xml.attr(node, "abstractNumId");

		for (const n of xml.elements(node)) {
			switch (n.localName) {
				case "lvl":
					result.push(this.parseNumberingLevel(id, n, bullets));
					break;
			}
		}

		return result;
	}

	parseNumberingLevel(id: string, node: Element, bullets: any[]): IDomNumbering {
		var result: IDomNumbering = {
			id: id,
			level: xml.intAttr(node, "ilvl"),
			start: 1,
			pStyleName: undefined,
			pStyle: {},
			rStyle: {},
			suff: "tab"
		};

		for (const n of xml.elements(node)) {
			switch (n.localName) {
				case "start":
					result.start = xml.intAttr(n, "val");
					break;

				case "pPr":
					this.parseDefaultProperties(n, result.pStyle);
					break;

				case "rPr":
					this.parseDefaultProperties(n, result.rStyle);
					break;

				case "lvlPicBulletId":
					var bulletId = xml.intAttr(n, "val");
					result.bullet = bullets.find(x => x?.id == bulletId);
					break;

				case "lvlText":
					result.levelText = xml.attr(n, "val");
					break;

				case "pStyle":
					result.pStyleName = xml.attr(n, "val");
					break;

				case "numFmt":
					result.format = xml.attr(n, "val");
					break;

				case "suff":
					result.suff = xml.attr(n, "val");
					break;

				case "lvlRestart":
					// w:val is the 1-based ancestor level at which this level's
					// counter restarts. Absence of the element (undefined here)
					// means "default — restart at the parent level".
					result.restart = xml.intAttr(n, "val");
					break;

				case "lvlJc":
					// left | right | center | start | end; validated at render
					// time before emission into CSS.
					result.justification = xml.attr(n, "val");
					break;

				case "isLgl":
					// w:isLgl is a boolean toggle-element. When present with no
					// val / val=1 / val="true", every %N placeholder in lvlText
					// must render as arabic, regardless of that lower level's
					// own numFmt.
					var lglVal = xml.attr(n, "val");
					result.isLgl = lglVal === undefined || lglVal === "" ||
						lglVal === "1" || lglVal === "true" || lglVal === "on";
					break;
			}
		}

		return result;
	}

	parseSdt(node: Element, parser: Function): OpenXmlElement[] {
		const sdtContent = xml.element(node, "sdtContent");
		if (!sdtContent) return [];

		const children: OpenXmlElement[] = parser(sdtContent) ?? [];

		// w:sdtPr holds the content-control metadata. w:alias is the visible
		// label ("Publication Date", etc.); w:tag is the programmatic id.
		// When either is present — or a typed form control (checkbox,
		// dropdown, date, picture, gallery) is detected — we wrap the
		// content so the renderer can emit an a11y group and/or form
		// control. Otherwise we unwrap as before.
		const sdtPr = xml.element(node, "sdtPr");
		if (sdtPr) {
			const aliasEl = xml.element(sdtPr, "alias");
			const tagEl = xml.element(sdtPr, "tag");
			const alias = aliasEl ? xml.attr(aliasEl, "val") : null;
			const tag = tagEl ? xml.attr(tagEl, "val") : null;
			const control = this.parseSdtControl(sdtPr);
			if (alias || tag || control) {
				const wrapper: WmlSdt = { type: DomType.Sdt, children };
				if (alias) wrapper.sdtAlias = alias;
				if (tag) wrapper.sdtTag = tag;
				if (control) wrapper.sdtControl = control;
				return [wrapper];
			}
		}

		return children;
	}

	// Inspect w:sdtPr for a typed content-control marker. Matches are by
	// localName so w14:* elements are found regardless of namespace prefix,
	// the same pattern xml.element() uses elsewhere.
	private parseSdtControl(sdtPr: Element): SdtControl | null {
		for (const el of xml.elements(sdtPr)) {
			switch (el.localName) {
				case "checkbox": {
					// w14:checkbox — children w14:checked, w14:checkedState,
					// w14:uncheckedState each carry a w14:val.
					const checkedEl = xml.element(el, "checked");
					const checkedStateEl = xml.element(el, "checkedState");
					const uncheckedStateEl = xml.element(el, "uncheckedState");
					const checkedRaw = checkedEl ? xml.attr(checkedEl, "val") : null;
					const checked = checkedRaw === "1" || checkedRaw === "true";
					// w14:checkedState w14:val is always a hex codepoint
					// ("2611" for ☑, "2610" for ☐). Renderer currently emits a
					// native <input type="checkbox"> and doesn't use these,
					// but they're captured for any future glyph-style render.
					const checkedChar = checkedStateEl ? xml.hexAttr(checkedStateEl, "val") : undefined;
					const uncheckedChar = uncheckedStateEl ? xml.hexAttr(uncheckedStateEl, "val") : undefined;
					const result: SdtCheckboxControl = { type: "checkbox", checked };
					if (checkedChar != null) result.checkedChar = checkedChar;
					if (uncheckedChar != null) result.uncheckedChar = uncheckedChar;
					return result;
				}
				case "dropDownList":
				case "comboBox": {
					// Both carry the same w:listItem children. Word
					// distinguishes editable vs strict, but in a read-only
					// renderer both collapse to a disabled <select>.
					const items: { displayText: string; value: string }[] = [];
					for (const li of xml.elements(el, "listItem")) {
						const displayText = xml.attr(li, "displayText");
						const value = xml.attr(li, "value");
						items.push({
							displayText: displayText ?? value ?? "",
							value: value ?? displayText ?? ""
						});
					}
					return { type: "dropdown", items };
				}
				case "date":
				case "sdtDate": {
					// Both w:date and w14:sdtDate are observed in the wild.
					const formatEl = xml.element(el, "dateFormat");
					const fullDateEl = xml.element(el, "fullDate");
					const format = formatEl ? xml.attr(formatEl, "val") : null;
					// w:date often has a w:fullDate attribute on the element
					// itself rather than a child element.
					const fullDateAttr = xml.attr(el, "fullDate");
					const fullDate = fullDateEl ? xml.attr(fullDateEl, "val") : fullDateAttr;
					return {
						type: "date",
						format: format ?? undefined,
						fullDate: fullDate ?? undefined
					};
				}
				case "picture":
					return { type: "picture" };
				case "docPartList":
				case "docPartObj":
					return { type: "gallery" };
			}
		}
		return null;
	}

	parseInserted(node: Element, parentParser: Function): OpenXmlElement {
		return <OpenXmlElement>{
			type: DomType.Inserted,
			revision: parseRevisionAttrs(node),
			children: parentParser(node)?.children ?? []
		};
	}

	parseDeleted(node: Element, parentParser: Function): OpenXmlElement {
		return <OpenXmlElement>{
			type: DomType.Deleted,
			revision: parseRevisionAttrs(node),
			children: parentParser(node)?.children ?? []
		};
	}

	parseMoveFrom(node: Element, parentParser: Function): OpenXmlElement {
		return <OpenXmlElement>{
			type: DomType.MoveFrom,
			revision: parseRevisionAttrs(node),
			children: parentParser(node)?.children ?? []
		};
	}

	parseMoveTo(node: Element, parentParser: Function): OpenXmlElement {
		return <OpenXmlElement>{
			type: DomType.MoveTo,
			revision: parseRevisionAttrs(node),
			children: parentParser(node)?.children ?? []
		};
	}

	parseAltChunk(node: Element): WmlAltChunk {
		return { type: DomType.AltChunk, children: [], id: xml.attr(node, "id") };
	}

	parseParagraph(node: Element): OpenXmlElement {
		var result = <WmlParagraph>{ type: DomType.Paragraph, children: [] };

		// w14:paraId is the Word-assigned stable identifier for the paragraph.
		// Used internally by the comments threading feature; also exposed in
		// the rendered output as data-para-id (see renderParagraph).
		const paraId = node.getAttributeNS("http://schemas.microsoft.com/office/word/2010/wordml", "paraId")
			?? node.getAttribute("w14:paraId")
			?? xml.attr(node, "paraId");
		if (paraId) {
			result.paraId = paraId;
		}

		for (let el of xml.elements(node)) {
			switch (el.localName) {
				case "pPr":
					this.parseParagraphProperties(el, result);
					break;

				case "r":
					result.children.push(this.parseRun(el, result));
					break;

				case "hyperlink":
					result.children.push(this.parseHyperlink(el, result));
					break;

				case "smartTag":
					result.children.push(this.parseSmartTag(el, result));
					break;

				case "bookmarkStart":
					result.children.push(parseBookmarkStart(el, xml));
					break;

				case "bookmarkEnd":
					result.children.push(parseBookmarkEnd(el, xml));
					break;

				case "commentRangeStart":
					result.children.push(new WmlCommentRangeStart(xml.attr(el, "id")));
					break;

				case "commentRangeEnd":
					result.children.push(new WmlCommentRangeEnd(xml.attr(el, "id")));
					break;

				case "oMath":
				case "oMathPara":
					result.children.push(this.parseMathElement(el));
					break;

				case "sdt":
					result.children.push(...this.parseSdt(el, e => this.parseParagraph(e).children));
					break;

				case "ins":
					result.children.push(this.parseInserted(el, e => this.parseParagraph(e)));
					break;

				case "del":
					result.children.push(this.parseDeleted(el, e => this.parseParagraph(e)));
					break;

				case "moveFrom":
					result.children.push(this.parseMoveFrom(el, e => this.parseParagraph(e)));
					break;

				case "moveTo":
					result.children.push(this.parseMoveTo(el, e => this.parseParagraph(e)));
					break;

				case "fldSimple":
					// A paragraph-level simple field wraps its cached result as
					// child runs (and occasionally nested fldSimple for e.g.
					// HYPERLINK containing formatted text). Recurse on the
					// children so renderSimpleField gets real content; store
					// the instruction so the renderer can wrap the result.
					result.children.push(this.parseFieldSimple(el, result));
					break;

				case "ruby":
					// w:ruby at paragraph scope. OOXML allows ruby as a direct
					// child of <w:p>; Wave 4.1 added the handler at run scope
					// only. Flagging here so paragraph-level ruby doesn't get
					// silently dropped.
					result.children.push(this.parseRuby(el, result));
					break;
			}
		}

		return result;
	}

	parseFieldSimple(node: Element, parent?: OpenXmlElement): WmlFieldSimple {
		const result: WmlFieldSimple = {
			type: DomType.SimpleField,
			instruction: xml.attr(node, "instr"),
			lock: xml.boolAttr(node, "lock", false),
			dirty: xml.boolAttr(node, "dirty", false),
			parent,
			children: [],
		};

		for (const c of xml.elements(node)) {
			switch (c.localName) {
				case "r":
					result.children.push(this.parseRun(c, result));
					break;
				case "hyperlink":
					result.children.push(this.parseHyperlink(c, result));
					break;
				case "fldSimple":
					result.children.push(this.parseFieldSimple(c, result));
					break;
			}
		}

		return result;
	}

	parseParagraphProperties(elem: Element, paragraph: WmlParagraph) {
		this.parseDefaultProperties(elem, paragraph.cssStyle = {}, null, c => {
			if (parseParagraphProperty(c, paragraph, xml))
				return true;

			switch (c.localName) {
				case "pStyle":
					paragraph.styleName = xml.attr(c, "val");
					break;

				case "cnfStyle":
					paragraph.className = values.classNameOfCnfStyle(c);
					break;

				case "framePr":
					this.parseFrame(c, paragraph);
					break;

				case "rPr":
					// Check for paragraph-mark revisions (w:pPr/w:rPr/w:ins or w:del).
					// The metadata is attached to the paragraph; rendering of the mark
					// itself is Phase 2 (#4). See #3.
					for (const rPrChild of xml.elements(c)) {
						if (rPrChild.localName === "ins") {
							paragraph.paragraphMarkRevisionKind = 'inserted';
							paragraph.revision = parseRevisionAttrs(rPrChild);
						} else if (rPrChild.localName === "del") {
							paragraph.paragraphMarkRevisionKind = 'deleted';
							paragraph.revision = parseRevisionAttrs(rPrChild);
						}
					}
					break;

				case "pPrChange":
					paragraph.formattingRevision = parseFormattingRevision(c);
					break;

				default:
					return false;
			}

			return true;
		});
	}

	parseFrame(node: Element, paragraph: WmlParagraph) {
		const dropCap = xml.attr(node, "dropCap");

		if (dropCap !== "drop" && dropCap !== "margin")
			return;

		// w:lines is how many lines the drop cap spans (default 3 per spec).
		// Clamp to a sane range so a hostile DOCX can't emit a huge font-size.
		const linesRaw = xml.intAttr(node, "lines");
		const lines = (Number.isInteger(linesRaw) && linesRaw >= 1 && linesRaw <= 10)
			? linesRaw
			: 3;

		paragraph.dropCap = dropCap;
		paragraph.dropCapLines = lines;
	}

	parseHyperlink(node: Element, parent?: OpenXmlElement): WmlHyperlink {
		var result: WmlHyperlink = <WmlHyperlink>{ type: DomType.Hyperlink, parent: parent, children: [] };

		result.anchor = xml.attr(node, "anchor");
		result.id = xml.attr(node, "id");

		const tooltip = xml.attr(node, "tooltip");
		if (tooltip) result.tooltip = tooltip;

		const targetFrame = xml.attr(node, "tgtFrame");
		if (targetFrame) result.targetFrame = targetFrame;

		for (const c of xml.elements(node)) {
			switch (c.localName) {
				case "r":
					result.children.push(this.parseRun(c, result));
					break;
			}
		}

		return result;
	}

	parseSmartTag(node: Element, parent?: OpenXmlElement): WmlSmartTag {
		var result: WmlSmartTag = { type: DomType.SmartTag, parent, children: [] };
		var uri = xml.attr(node, "uri");
		var element = xml.attr(node, "element");

		if (uri)
			result.uri = uri;

		if (element)
			result.element = element;

		for (const c of xml.elements(node)) {
			switch (c.localName) {
				case "r":
					result.children.push(this.parseRun(c, result));
					break;
				case "smartTag":
					result.children.push(this.parseSmartTag(c, result));
					break;
			}
		}

		return result;
	}

	parseRun(node: Element, parent?: OpenXmlElement): OpenXmlElement {
		var result: WmlRun = <WmlRun>{ type: DomType.Run, parent: parent, children: [] };

		for (let c of xml.elements(node)) {
			c = this.checkAlternateContent(c);

			switch (c.localName) {
				case "t":
					result.children.push(<WmlText>{
						type: DomType.Text,
						text: c.textContent
					});//.replace(" ", "\u00A0"); // TODO
					break;

				case "delText":
					result.children.push(<WmlText>{
						type: DomType.DeletedText,
						text: c.textContent
					});
					break;

				case "commentReference":
					result.children.push(new WmlCommentReference(xml.attr(c, "id")));
					break;

				case "fldSimple":
					result.children.push(this.parseFieldSimple(c, result));
					break;

				case "instrText":
					result.fieldRun = true;
					result.children.push(<WmlInstructionText>{
						type: DomType.Instruction,
						text: c.textContent
					});
					break;

				case "fldChar":
					result.fieldRun = true;
					result.children.push(<WmlFieldChar>{
						type: DomType.ComplexField,
						charType: xml.attr(c, "fldCharType"),
						lock: xml.boolAttr(c, "lock", false),
						dirty: xml.boolAttr(c, "dirty", false)
					});
					break;

				case "noBreakHyphen":
					result.children.push({ type: DomType.NoBreakHyphen });
					break;

				case "softHyphen":
					// U+00AD — browser-discretionary break point; invisible unless wrapped.
					result.children.push(<WmlText>{
						type: DomType.Text,
						text: "­"
					});
					break;

				case "br":
				case "cr":
					// w:cr is Word's explicit carriage return; render identically to a
					// text-wrapping br.
					result.children.push(<WmlBreak>{
						type: DomType.Break,
						break: c.localName === "cr" ? "textWrapping" : (xml.attr(c, "type") || "textWrapping")
					});
					break;

				case "lastRenderedPageBreak":
					result.children.push(<WmlBreak>{
						type: DomType.Break,
						break: "lastRenderedPageBreak"
					});
					break;

				case "sym":
					result.children.push(<WmlSymbol>{
						type: DomType.Symbol,
						font: encloseFontFamily(xml.attr(c, "font")),
						char: xml.hexAttr(c, "char")
					});
					break;

				case "tab":
					result.children.push({ type: DomType.Tab });
					break;

				case "footnoteReference":
					result.children.push(<WmlNoteReference>{
						type: DomType.FootnoteReference,
						id: xml.attr(c, "id")
					});
					break;

				case "endnoteReference":
					result.children.push(<WmlNoteReference>{
						type: DomType.EndnoteReference,
						id: xml.attr(c, "id")
					});
					break;

				case "drawing":
					let d = this.parseDrawing(c);

					if (d)
						result.children.push(d);
					break;

				case "pict":
					result.children.push(this.parseVmlPicture(c));
					break;

				case "ruby":
					result.children.push(this.parseRuby(c, result));
					break;

				case "rPr":
					this.parseRunProperties(c, result);
					break;
			}
		}

		// w:rPr/w:fitText and w:rPr/w:bdo are captured as run-level metadata
		// by parseRunProperties. When either is set, wrap the run in a
		// DomType.FitText / DomType.BidiOverride element so the renderer can
		// emit the appropriate host element without touching renderRun.
		let wrapped: OpenXmlElement = result;
		if (result.bidiOverride) {
			const bidi: WmlBidiOverride = {
				type: DomType.BidiOverride,
				dir: result.bidiOverride,
				parent,
				children: [wrapped]
			};
			wrapped.parent = bidi;
			delete result.bidiOverride;
			wrapped = bidi;
		}
		if (result.fitText) {
			const fit: WmlFitText = {
				type: DomType.FitText,
				width: result.fitText.width,
				id: result.fitText.id,
				parent,
				children: [wrapped]
			};
			wrapped.parent = fit;
			delete result.fitText;
			wrapped = fit;
		}

		return wrapped;
	}

	// w:ruby — East-Asian phonetic guide (furigana). Structure:
	//   <w:ruby>
	//     <w:rubyPr> <w:rt> <w:rubyBase>
	// rubyPr carries layout hints (alignment, font sizes, language). The
	// rt / rubyBase wrappers each contain one or more <w:r>. The browser
	// renders <ruby><rb>base</rb><rt>annotation</rt></ruby> natively, so
	// we just emit wrapper elements here. DOCX-derived strings from this
	// branch reach the DOM only via renderElements of the child runs,
	// which already sanitise text content and style.
	parseRuby(node: Element, parent?: OpenXmlElement): WmlRuby {
		const result: WmlRuby = { type: DomType.Ruby, parent, children: [] };

		for (const c of xml.elements(node)) {
			switch (c.localName) {
				case "rubyPr": {
					const rubyPr: WmlRuby["rubyPr"] = {};
					for (const p of xml.elements(c)) {
						const v = xml.attr(p, "val");
						switch (p.localName) {
							case "rubyAlign":
								// Allowlist rubyAlign — it would reach a CSS value
								// if we ever styled the ruby container.
								if (/^(center|distributeLetter|distributeSpace|left|right|rightVertical|start|end)$/.test(v))
									rubyPr.rubyAlign = v;
								break;
							case "hps": {
								const n = xml.intAttr(p, "val");
								if (Number.isFinite(n) && n > 0 && n < 1000) rubyPr.hps = n;
								break;
							}
							case "hpsBaseText": {
								const n = xml.intAttr(p, "val");
								if (Number.isFinite(n) && n > 0 && n < 1000) rubyPr.hpsBaseText = n;
								break;
							}
							case "hpsRaise": {
								const n = xml.intAttr(p, "val");
								if (Number.isFinite(n) && n >= 0 && n < 1000) rubyPr.hpsRaise = n;
								break;
							}
							case "lid":
								// Language tag — reaches `lang` attribute only,
								// which the browser attribute-encodes.
								if (v) rubyPr.lid = v;
								break;
						}
					}
					result.rubyPr = rubyPr;
					break;
				}
				case "rubyBase": {
					const base: OpenXmlElement = { type: DomType.RubyBase, parent: result, children: [] };
					for (const r of xml.elements(c)) {
						if (r.localName === "r")
							base.children.push(this.parseRun(r, base));
					}
					result.children.push(base);
					break;
				}
				case "rt": {
					const rt: OpenXmlElement = { type: DomType.RubyText, parent: result, children: [] };
					for (const r of xml.elements(c)) {
						if (r.localName === "r")
							rt.children.push(this.parseRun(r, rt));
					}
					result.children.push(rt);
					break;
				}
			}
		}

		return result;
	}

	parseMathElement(elem: Element): OpenXmlElement {
		const propsTag = `${elem.localName}Pr`;
		const result = { type: mmlTagMap[elem.localName], children: [] } as OpenXmlElement;

		for (const el of xml.elements(elem)) {
			const childType = mmlTagMap[el.localName];

			if (childType) {
				result.children.push(this.parseMathElement(el));
			} else if (el.localName == "r") {
				var run = this.parseRun(el);
				run.type = DomType.MmlRun;
				result.children.push(run);
			} else if (el.localName == propsTag) {
				result.props = this.parseMathProperies(el);
			}
		}

		return result;
	}

	parseMathProperies(elem: Element): Record<string, any> {
		const result: Record<string, any> = {};

		for (const el of xml.elements(elem)) {
			switch (el.localName) {
				case "chr": result.char = xml.attr(el, "val"); break;
				case "vertJc": result.verticalJustification = xml.attr(el, "val"); break;
				case "pos": result.position = xml.attr(el, "val"); break;
				case "degHide": result.hideDegree = xml.boolAttr(el, "val"); break;
				case "begChr": result.beginChar = xml.attr(el, "val"); break;
				case "endChr": result.endChar = xml.attr(el, "val"); break;
				case "limLoc": result.limLoc = xml.attr(el, "val"); break;
			}
		}

		return result;
	}

	parseRunProperties(elem: Element, run: WmlRun) {
		this.parseDefaultProperties(elem, run.cssStyle = {}, null, c => {
			switch (c.localName) {
				case "rStyle":
					run.styleName = xml.attr(c, "val");
					break;

				case "vertAlign":
					run.verticalAlign = values.valueOfVertAlign(c, true);
					break;

				case "rPrChange":
					run.formattingRevision = parseFormattingRevision(c);
					break;

				case "fitText": {
					// w:fitText — run-level "fit to width" directive. val is
					// the target width in twips. Numeric via floatAttr so the
					// raw string never reaches CSS.
					const w = xml.floatAttr(c, "val");
					if (Number.isFinite(w) && w > 0) {
						run.fitText = { width: w, id: xml.attr(c, "id") || undefined };
					}
					break;
				}

				case "bdo": {
					// w:bdo — explicit bidi override. The DOCX-derived val is
					// allowlisted against /^(ltr|rtl)$/ before being stored.
					// Any other value is silently dropped.
					const raw = xml.attr(c, "val");
					if (raw === "ltr" || raw === "rtl") {
						run.bidiOverride = raw;
					}
					break;
				}

				default:
					return false;
			}

			return true;
		});
	}

	parseVmlPicture(elem: Element): OpenXmlElement {
		const result = { type: DomType.VmlPicture, children: [] };

		for (const el of xml.elements(elem)) {
			const child = parseVmlElement(el, this);
			child && result.children.push(child);
		}

		return result;
	}

	checkAlternateContent(elem: Element): Element {
		if (elem.localName != 'AlternateContent')
			return elem;

		var choice = xml.element(elem, "Choice");

		if (choice) {
			var requires = xml.attr(choice, "Requires");
			var namespaceURI = elem.lookupNamespaceURI(requires);

			if (supportedNamespaceURIs.includes(namespaceURI))
				return choice.firstElementChild;
		}

		return xml.element(elem, "Fallback")?.firstElementChild;
	}

	parseDrawing(node: Element): OpenXmlElement {
		for (var n of xml.elements(node)) {
			switch (n.localName) {
				case "inline":
				case "anchor":
					return this.parseDrawingWrapper(n);
			}
		}
	}

	parseDrawingWrapper(node: Element): OpenXmlElement {
		var result = <OpenXmlElement>{ type: DomType.Drawing, children: [], cssStyle: {}, props: {} };
		var isAnchor = node.localName == "anchor";
		// Surface the anchor/inline distinction on the parsed element so
		// renderDrawing can project it as a data-* hook. Responsive CSS
		// targets floating anchors under @media (max-width: 768px); non-
		// responsive consumers simply ignore the attribute.
		(result.props ??= {}).isAnchor = isAnchor;

		// DrawingML stores offsets in English Metric Units: 914400 EMU = 1 inch = 96 CSS pixels.
		// For dist*/wrapPolygon we want CSS pixels, so divide by 9525.
		const EMU_PER_PX = 9525;

		// Allowlists for DOCX-derived enum values so they can be used safely in CSS strings.
		const WRAP_TEXT_ALLOWED = new Set(["bothSides", "left", "right", "largest"]);
		const RELATIVE_FROM_ALLOWED = new Set([
			"margin", "page", "column", "character", "leftMargin", "rightMargin",
			"insideMargin", "outsideMargin", "paragraph", "line", "topMargin",
			"bottomMargin"
		]);

		// dist* attributes are EMU integers describing padding around the wrapped shape.
		// Parsed as floats explicitly so DOCX-controlled text never reaches a template
		// literal unchecked.
		const distT = xml.floatAttr(node, "distT", 0);
		const distB = xml.floatAttr(node, "distB", 0);
		const distL = xml.floatAttr(node, "distL", 0);
		const distR = xml.floatAttr(node, "distR", 0);
		const marginTopPx = (distT || 0) / EMU_PER_PX;
		const marginBottomPx = (distB || 0) / EMU_PER_PX;
		const marginLeftPx = (distL || 0) / EMU_PER_PX;
		const marginRightPx = (distR || 0) / EMU_PER_PX;

		let wrapType: "wrapTopAndBottom" | "wrapNone" | "wrapSquare" | "wrapTight" | "wrapThrough" | null = null;
		let wrapText: string | null = null;
		let wrapPolygonPoints: Array<[number, number]> | null = null;
		let simplePos = xml.boolAttr(node, "simplePos");
		let behindDoc = xml.boolAttr(node, "behindDoc");

		// Raw extent in EMU (used for polygon coord scaling). cssStyle width/height
		// keep their existing pt representation for back-compat.
		let extentCx = 0;
		let extentCy = 0;

		let posX = { relative: "page", align: "left", offset: "0", offsetEmu: 0 };
		let posY = { relative: "page", align: "top", offset: "0", offsetEmu: 0 };

		// wp:docPr/@descr — the preferred alt-text source. See parsePicture
		// for the pic:cNvPr/@descr and a:blip/@descr fallbacks.
		let docPrDescr: string | null = null;

		for (var n of xml.elements(node)) {
			switch (n.localName) {
				case "docPr":
					docPrDescr = xml.attr(n, "descr");
					break;

				case "simplePos":
					if (simplePos) {
						posX.offsetEmu = xml.floatAttr(n, "x", 0) || 0;
						posY.offsetEmu = xml.floatAttr(n, "y", 0) || 0;
						posX.offset = xml.lengthAttr(n, "x", LengthUsage.Emu);
						posY.offset = xml.lengthAttr(n, "y", LengthUsage.Emu);
					}
					break;

				case "extent":
					extentCx = xml.floatAttr(n, "cx", 0) || 0;
					extentCy = xml.floatAttr(n, "cy", 0) || 0;
					result.cssStyle["width"] = xml.lengthAttr(n, "cx", LengthUsage.Emu);
					result.cssStyle["height"] = xml.lengthAttr(n, "cy", LengthUsage.Emu);
					break;

				case "positionH":
				case "positionV":
					if (!simplePos) {
						let pos = n.localName == "positionH" ? posX : posY;
						var alignNode = xml.element(n, "align");
						var offsetNode = xml.element(n, "posOffset");

						pos.relative = xml.attr(n, "relativeFrom") ?? pos.relative;

						if (alignNode)
							pos.align = alignNode.textContent;

						if (offsetNode) {
							pos.offset = convertLength(offsetNode.textContent, LengthUsage.Emu);
							const parsed = parseFloat(offsetNode.textContent);
							pos.offsetEmu = Number.isFinite(parsed) ? parsed : 0;
						}
					}
					break;

				case "wrapTopAndBottom":
					wrapType = "wrapTopAndBottom";
					break;

				case "wrapNone":
					wrapType = "wrapNone";
					break;

				case "wrapSquare":
					wrapType = "wrapSquare";
					wrapText = xml.attr(n, "wrapText");
					break;

				case "wrapTight":
				case "wrapThrough":
					wrapType = n.localName as "wrapTight" | "wrapThrough";
					wrapText = xml.attr(n, "wrapText");
					{
						const polyNode = xml.element(n, "wrapPolygon");
						if (polyNode) {
							const pts: Array<[number, number]> = [];
							for (const child of xml.elements(polyNode)) {
								if (child.localName !== "start" && child.localName !== "lineTo")
									continue;
								const x = xml.floatAttr(child, "x", NaN);
								const y = xml.floatAttr(child, "y", NaN);
								if (Number.isFinite(x) && Number.isFinite(y))
									pts.push([x, y]);
							}
							if (pts.length >= 3)
								wrapPolygonPoints = pts;
						}
					}
					break;

				case "graphic":
					var g = this.parseGraphic(n);

					if (g)
						result.children.push(g);
					break;
			}
		}

		// Validate DOCX-derived enums against the allowlists before they reach any
		// CSS string — the browser does not sanitise CSS property values.
		const safeWrapText = wrapText && WRAP_TEXT_ALLOWED.has(wrapText) ? wrapText : null;
		const safeRelativeH = RELATIVE_FROM_ALLOWED.has(posX.relative) ? posX.relative : "page";

		// Decide which side the image floats on for the flowing-text wrap modes.
		// "left" wrapText means text flows only on the left side — the image sits
		// on the right, so it floats right. Mirror for "right". For "bothSides" /
		// "largest" we fall back to the anchor's horizontal alignment (treating
		// "center" as left).
		const floatFromWrapText = (wt: string | null, align: string): "left" | "right" => {
			if (wt === "left") return "right";
			if (wt === "right") return "left";
			if (align === "right") return "right";
			return "left";
		};

		const applyMarginsToStyle = () => {
			if (distT) result.cssStyle["margin-top"] = `${marginTopPx.toFixed(2)}px`;
			if (distB) result.cssStyle["margin-bottom"] = `${marginBottomPx.toFixed(2)}px`;
			if (distL) result.cssStyle["margin-left"] = `${marginLeftPx.toFixed(2)}px`;
			if (distR) result.cssStyle["margin-right"] = `${marginRightPx.toFixed(2)}px`;
		};

		const applyShapeMargin = () => {
			// shape-margin takes a single length. Use the largest dist* so text
			// keeps a safe gap from the shape on every side.
			const maxDist = Math.max(marginTopPx, marginBottomPx, marginLeftPx, marginRightPx);
			if (maxDist > 0)
				result.cssStyle["shape-margin"] = `${maxDist.toFixed(2)}px`;
		};

		const buildPolygonCss = (): string | null => {
			if (!wrapPolygonPoints || wrapPolygonPoints.length < 3) return null;
			if (extentCx <= 0 || extentCy <= 0) return null;
			// Polygon points in wp:wrapPolygon are in EMU relative to the shape
			// bounding box (extent cx/cy). Emitting percentages lets the polygon
			// scale with the image element's box.
			const segs = wrapPolygonPoints.map(([x, y]) => {
				const px = (x / extentCx) * 100;
				const py = (y / extentCy) * 100;
				return `${px.toFixed(2)}% ${py.toFixed(2)}%`;
			});
			return `polygon(${segs.join(", ")})`;
		};

		if (wrapType == "wrapTopAndBottom") {
			result.cssStyle['display'] = 'block';

			if (posX.align) {
				result.cssStyle['text-align'] = posX.align;
				result.cssStyle['width'] = "100%";
			}
		}
		else if (wrapType == "wrapNone") {
			result.cssStyle['display'] = 'block';
			result.cssStyle['position'] = 'relative';
			result.cssStyle["width"] = "0px";
			result.cssStyle["height"] = "0px";

			if (posX.offset)
				result.cssStyle["left"] = posX.offset;
			if (posY.offset)
				result.cssStyle["top"] = posY.offset;
		}
		else if (wrapType == "wrapSquare" || wrapType == "wrapTight" || wrapType == "wrapThrough") {
			const floatSide = floatFromWrapText(safeWrapText, posX.align);

			// relativeFrom=margin → behave like a plain float (current default).
			// relativeFrom=column → inline-block pulled into the text flow via a
			//   negative margin on the opposite side of the float.
			// relativeFrom=paragraph → absolute positioning using posX.offsetEmu
			//   converted to px.
			if (safeRelativeH === "paragraph") {
				result.cssStyle["position"] = "absolute";
				const leftPx = (posX.offsetEmu || 0) / EMU_PER_PX;
				result.cssStyle["left"] = `${leftPx.toFixed(2)}px`;
				if (posY.offsetEmu) {
					const topPx = (posY.offsetEmu || 0) / EMU_PER_PX;
					result.cssStyle["top"] = `${topPx.toFixed(2)}px`;
				}
				applyMarginsToStyle();
			} else if (safeRelativeH === "column") {
				result.cssStyle["display"] = "inline-block";
				result.cssStyle["float"] = floatSide;
				// Pull the shape slightly into the column by negating the
				// corresponding dist* so text wraps tightly against the text
				// column edge.
				applyMarginsToStyle();
				if (floatSide === "left" && distL)
					result.cssStyle["margin-left"] = `${(-marginLeftPx).toFixed(2)}px`;
				else if (floatSide === "right" && distR)
					result.cssStyle["margin-right"] = `${(-marginRightPx).toFixed(2)}px`;
			} else {
				// "margin" (and any other allowlisted relativeFrom) — plain float.
				result.cssStyle["float"] = floatSide;
				applyMarginsToStyle();
			}

			applyShapeMargin();

			if (wrapType === "wrapTight" || wrapType === "wrapThrough") {
				const polyCss = buildPolygonCss();
				if (polyCss)
					result.cssStyle["shape-outside"] = polyCss;
			}
		}
		else if (isAnchor && (posX.align == 'left' || posX.align == 'right')) {
			result.cssStyle["float"] = posX.align;
		}

		// Propagate wp:docPr/@descr down to the contained IDomImage so
		// renderImage can emit it as alt="". parsePicture may have already
		// set altText from pic:cNvPr/a:blip; docPr wins since it's authored
		// at the Word "Alt Text…" UI level.
		if (docPrDescr != null) {
			this.setImageAltText(result, docPrDescr);
		}

		return result;
	}

	// Recursively finds the single IDomImage child (if any) under a drawing
	// wrapper and assigns altText. The tree shape is:
	//   Drawing → (graphic?) → Image. One picture per drawing in practice.
	private setImageAltText(elem: OpenXmlElement, descr: string) {
		if (elem.type === DomType.Image) {
			(elem as IDomImage).altText = descr;
			return;
		}
		if (elem.children) {
			for (const c of elem.children) this.setImageAltText(c, descr);
		}
	}

	parseGraphic(elem: Element): OpenXmlElement {
		var graphicData = xml.element(elem, "graphicData");

		// Inspect graphicData@uri once so we can dispatch chartEx
		// (which embeds <cx:chart>, not <c:chart>) without widening
		// the localName switch below. The classic chart URI is
		// recognised by localName "chart" and continues to go through
		// parseChartReference.
		const uri = graphicData ? xml.attr(graphicData, "uri") : null;
		const CHARTEX_URI = "http://schemas.microsoft.com/office/drawingml/2014/chartex";
		const SMARTART_URI = "http://schemas.openxmlformats.org/drawingml/2006/diagram";

		// SmartArt: <a:graphicData uri="…/diagram"> contains a
		// <dgm:relIds> pointing at the five /word/diagrams/ parts. We
		// don't implement the layout engine, so the best user-visible
		// outcome is to walk back up to any enclosing mc:AlternateContent
		// and render its <mc:Fallback> drawing (which Word itself
		// pre-renders as a picture for exactly this situation). When
		// that fails — either no AlternateContent wrapper or no
		// Fallback — we emit a labelled placeholder. See
		// parseSmartArtReference for the ancestor walk.
		if (uri === SMARTART_URI) {
			return this.parseSmartArtReference(elem, graphicData);
		}

		for (let n of xml.elements(graphicData)) {
			// ChartEx: the child is <cx:chart r:id="..."/>. URI gates
			// the branch so an attacker-supplied local name can't hit
			// the chartEx path for a classic chart.
			if (uri === CHARTEX_URI && n.localName === "chart") {
				return this.parseChartExReference(n);
			}
			switch (n.localName) {
				case "pic":
					return this.parsePicture(n);
				case "wsp":
					// DrawingML shape (rectangle, arrow, callout, text-
					// box, …). See parseDrawingShape — hardened against
					// DOCX-supplied strings via PRESET_GEOMETRY_ALLOWLIST
					// and sanitizeCssColor.
					return this.parseDrawingShape(n);
				case "wgp":
					// DrawingML shape group. Shapes can nest groups, so
					// parseDrawingShapeGroup recurses back through
					// parseGraphicDataChild for every entry.
					return this.parseDrawingShapeGroup(n);
				case "chart":
					// DrawingML chart reference. Actual SVG rendering
					// happens in html-renderer.ts after looking up the
					// ChartPart via the enclosing part's rel map. See
					// src/charts/render.ts and SECURITY above.
					return this.parseChartReference(n);
			}
		}

		return null;
	}

	// <c:chart r:id="rIdX"/> inside <a:graphicData
	// uri="…/drawingml/2006/chart">. We record the rel id only; the
	// renderer resolves it against the document's relationship map.
	// Attacker-controlled string never reaches a CSS selector or class
	// name — only read via findPartByRelId's Map/Object lookup.
	parseChartReference(elem: Element): DrawingChart {
		const relId = xml.attr(elem, "id");
		return {
			type: DomType.Chart,
			relId: relId ?? "",
		};
	}

	// <cx:chart r:id="rIdX"/> inside <a:graphicData
	// uri="http://schemas.microsoft.com/office/drawingml/2014/chartex">.
	// Same shape as parseChartReference but the relId resolves to a
	// ChartExPart and renders as a placeholder.
	parseChartExReference(elem: Element): DrawingChartEx {
		const relId = xml.attr(elem, "id");
		return {
			type: DomType.ChartEx,
			relId: relId ?? "",
		};
	}

	// SmartArt dispatch — see the big comment in parseGraphic above.
	//
	// Two paths out of here:
	//   1. If an <mc:AlternateContent> ancestor exists, re-parse the
	//      <mc:Fallback>'s drawing (almost always a <pic:pic> of the
	//      rendered SmartArt) and return *that* as the substitute.
	//      checkAlternateContent usually handles this at parseRun,
	//      but the SmartArt <a:graphic> can only reach this method
	//      when it did NOT — e.g. the AlternateContent was deeper than
	//      the run level, or the document doesn't wrap the SmartArt
	//      in AlternateContent at all.
	//   2. Otherwise emit a placeholder that a consumer can style.
	//      The layout URN is captured from <dgm:relIds r:lo="…"> via
	//      the layout-part relationship so a future SmartArt renderer
	//      can dispatch on it.
	//
	// Security: r:dm / r:lo / r:qs / r:cs are Id tokens resolved through
	// the part-map (never interpolated into CSS). The layout URN is
	// allowlisted by DiagramLayoutPart before it reaches the DOM as
	// data-smartart-layout. Fallback drawing re-enters parseDrawing*
	// which applies the usual DOCX-string sinks (textContent /
	// setAttribute), so no new trust boundary is introduced.
	parseSmartArtReference(graphic: Element, graphicData: Element | null): OpenXmlElement {
		// Walk ancestors for the nearest mc:AlternateContent. We only
		// consider the *same* AlternateContent that owns this graphic
		// (so we never steal a Fallback from an unrelated sibling).
		const alt = findAncestorByLocalName(graphic, "AlternateContent");
		if (alt) {
			const fallback = xml.element(alt, "Fallback");
			if (fallback) {
				// The Fallback wraps either a <w:drawing> or a bare
				// DrawingML element depending on the generator. The
				// most common shape is <w:drawing><wp:inline>…</w:drawing>
				// but some files put <wp:inline> or even <mc:Choice>-
				// style bare content directly. Try each in order.
				const drawing = xml.element(fallback, "drawing");
				if (drawing) {
					const r = this.parseDrawing(drawing);
					if (r) return r;
				}
				// Bare <wp:inline>/<wp:anchor> fallback.
				for (const c of xml.elements(fallback)) {
					if (c.localName === "inline" || c.localName === "anchor") {
						const r = this.parseDrawingWrapper(c);
						if (r) return r;
					}
				}
				// Some files put a raw <pic:pic> in Fallback. Extremely
				// rare, but handle it so we don't drop content.
				for (const c of xml.elements(fallback)) {
					if (c.localName === "pic") {
						return this.parsePicture(c);
					}
				}
			}
		}

		// No usable Fallback — emit a placeholder. Extract <dgm:relIds>
		// for the rel-id capture so a future layout engine can find the
		// parts. Nothing here reaches a CSS string or selector.
		const relIds: DrawingSmartArt["relIds"] = {};
		if (graphicData) {
			const rel = xml.element(graphicData, "relIds");
			if (rel) {
				const dm = xml.attr(rel, "dm");
				const lo = xml.attr(rel, "lo");
				const qs = xml.attr(rel, "qs");
				const cs = xml.attr(rel, "cs");
				if (dm) relIds.dm = dm;
				if (lo) relIds.lo = lo;
				if (qs) relIds.qs = qs;
				if (cs) relIds.cs = cs;
			}
		}

		return <DrawingSmartArt>{
			type: DomType.SmartArt,
			children: [],
			cssStyle: {},
			relIds,
		};
	}

	// Hard-coded allowlist of DrawingML preset-geometry names that
	// parseDrawingShape is willing to emit. Anything outside the
	// allowlist — and specifically `custGeom`, which we parse but do
	// not render — falls back to `rect`. The strings returned here
	// reach an SVG attribute value and a CSS class name, so they must
	// never be DOCX-controlled.
	//
	// Keep in sync with the switch in presetGeometryToSvgPath in
	// src/drawing/shapes.ts.
	private readonly PRESET_GEOMETRY_ALLOWLIST = new Set([
		"rect", "roundRect", "ellipse", "triangle", "rtTriangle", "diamond",
		"parallelogram", "trapezoid", "pentagon", "hexagon", "octagon", "line",
		"rightArrow", "leftArrow", "upArrow", "downArrow", "leftRightArrow",
		"wedgeRectCallout", "wedgeRoundRectCallout", "wedgeEllipseCallout",
		"star5", "star6", "star8", "cloudCallout",
	]);

	// Parses <a:xfrm> inside <wps:spPr> / <wpg:grpSpPr>. Returns
	// {x, y, cx, cy, rot} — all EMU — or undefined when the element is
	// missing. Callers defaulting missing fields is cheaper than
	// duplicating that here.
	private parseXfrm(
		xfrm: Element | null | undefined,
	): { x: number; y: number; cx: number; cy: number; rot?: number } | undefined {
		if (!xfrm) return undefined;
		const result: { x: number; y: number; cx: number; cy: number; rot?: number } =
			{ x: 0, y: 0, cx: 0, cy: 0 };
		const rotRaw = xml.intAttr(xfrm, "rot", 0);
		if (rotRaw) result.rot = rotRaw / 60000;
		for (const n of xml.elements(xfrm)) {
			switch (n.localName) {
				case "off":
					result.x = xml.floatAttr(n, "x", 0) || 0;
					result.y = xml.floatAttr(n, "y", 0) || 0;
					break;
				case "ext":
					result.cx = xml.floatAttr(n, "cx", 0) || 0;
					result.cy = xml.floatAttr(n, "cy", 0) || 0;
					break;
			}
		}
		return result;
	}

	// Parses <a:solidFill>/<a:gradFill>/<a:blipFill>/<a:noFill>/<a:pattFill>
	// into the shape-fill model. solidFill resolves to a concrete CSS
	// colour (via parseColor + the fall-back palette), gradient and
	// pattern fills preserve their ColourRefs so the renderer can
	// resolve schemeClr against the active theme palette. All colour
	// values pass through sanitizeCssColor on the way in.
	private parseShapeFill(spPr: Element): DrawingShape["fill"] | undefined {
		if (!spPr) return undefined;
		for (const n of xml.elements(spPr)) {
			switch (n.localName) {
				case "noFill":
					return { type: "none" };
				case "solidFill": {
					// Resolve schemeClr / lumMod / lumOff here so the
					// shape model carries a ready-to-use CSS colour for
					// the common solid-fill case. Theme palette is not
					// available at parse time, so we use defaults here
					// and let the renderer apply a theme override via
					// the schemeClr path for gradients / patterns.
					const ref = this.parseColor(n);
					if (ref?.hex) return { type: "solid", color: ref.hex };
					// schemeClr-only solid fill — store as gradient with a
					// single stop so the renderer resolves it with theme.
					if (ref?.scheme) {
						return {
							type: "gradient",
							gradient: {
								kind: "linear",
								stops: [
									{ pos: 0, colour: ref },
									{ pos: 1, colour: ref },
								],
								angle: 0,
							},
						};
					}
					return undefined;
				}
				case "gradFill": {
					const grad = this.parseGradientFill(n);
					if (grad) return { type: "gradient", gradient: grad };
					return undefined;
				}
				case "pattFill": {
					const patt = this.parsePatternFill(n);
					if (patt) return { type: "pattern", pattern: patt };
					return undefined;
				}
				case "blipFill":
					// Image-filled shape — out of scope for v1. TODO.
					return undefined;
			}
		}
		return undefined;
	}

	// Parses <a:ln w="…"> → stroke width (EMU) + optional <a:solidFill>
	// colour. schemeClr colours are resolved against the default palette
	// at parse time (stroke doesn't share the renderer's theme
	// access).
	private parseShapeStroke(spPr: Element): DrawingShape["stroke"] | undefined {
		if (!spPr) return undefined;
		const ln = xml.element(spPr, "ln");
		if (!ln) return undefined;
		const result: DrawingShape["stroke"] = {};
		const w = xml.intAttr(ln, "w", 0);
		if (w > 0) result.width = w;
		const solid = xml.element(ln, "solidFill");
		if (solid) {
			const ref = this.parseColor(solid);
			if (ref?.hex) result.color = ref.hex;
			else if (ref?.scheme) {
				// Resolve against default palette so the colour is
				// usable immediately — stroke doesn't flow through the
				// gradient/defs pipeline. Theme palette override is
				// only applied to shape fills (renderer has access) so
				// stroke defaults are "close enough" for v1.
				const p = DEFAULT_THEME_PALETTE[ref.scheme];
				const sanitized = sanitizeCssColor(p);
				if (sanitized) result.color = sanitized;
			}
		}
		return result;
	}

	// Parse a colour-source element (<a:solidFill>, <a:gs>, <a:fgClr>,
	// <a:outerShdw>, …) into a ColourRef capturing the base colour plus
	// any lumMod / lumOff / tint / shade / alpha modifiers. Returns null
	// if the element carries no recognisable colour.
	//
	// Security: hex values pass through sanitizeCssColor before being
	// stored; scheme slot names are validated against an allowlist in
	// resolveColour before reaching any attribute sink.
	private parseColor(elem: Element): ColourRef | null {
		if (!elem) return null;
		// <a:srgbClr>, <a:schemeClr>, <a:sysClr>, <a:prstClr> can all
		// appear directly or as children of the wrapper element.
		const srgb = xml.element(elem, "srgbClr") ?? (elem.localName === "srgbClr" ? elem : null);
		const scheme = xml.element(elem, "schemeClr") ?? (elem.localName === "schemeClr" ? elem : null);
		const sys = xml.element(elem, "sysClr") ?? (elem.localName === "sysClr" ? elem : null);

		const ref: ColourRef = {};
		let modsSource: Element | null = null;

		if (srgb) {
			const val = xml.attr(srgb, "val");
			const sanitized = sanitizeCssColor(val);
			if (sanitized) ref.hex = sanitized;
			modsSource = srgb;
		} else if (sys) {
			// <a:sysClr lastClr="…"> carries the resolved hex.
			const val = xml.attr(sys, "lastClr") ?? xml.attr(sys, "val");
			const sanitized = sanitizeCssColor(val);
			if (sanitized) ref.hex = sanitized;
			modsSource = sys;
		} else if (scheme) {
			const val = xml.attr(scheme, "val");
			if (isAllowedSchemeSlot(val)) ref.scheme = val;
			modsSource = scheme;
		}

		if (modsSource) {
			const mods: ColourModifiers = {};
			let gotMod = false;
			for (const m of xml.elements(modsSource)) {
				const v = xml.intAttr(m, "val");
				if (v == null || !Number.isFinite(v)) continue;
				switch (m.localName) {
					case "lumMod": mods.lumMod = v; gotMod = true; break;
					case "lumOff": mods.lumOff = v; gotMod = true; break;
					case "tint":   mods.tint   = v; gotMod = true; break;
					case "shade":  mods.shade  = v; gotMod = true; break;
					case "alpha":  mods.alpha  = v; gotMod = true; break;
				}
			}
			if (gotMod) ref.mods = mods;
		}

		if (!ref.hex && !ref.scheme) return null;
		return ref;
	}

	// Parses <a:gradFill> into a GradientFill model. Stops preserve
	// their ColourRef so the renderer can resolve schemeClr against the
	// document's theme palette. Angles are converted from DOCX's
	// 60000ths-of-a-degree units to plain degrees.
	private parseGradientFill(el: Element): GradientFill | null {
		const gsLst = xml.element(el, "gsLst");
		if (!gsLst) return null;
		const stops: Array<{ pos: number; colour: ColourRef }> = [];
		for (const gs of xml.elements(gsLst, "gs")) {
			const posRaw = xml.intAttr(gs, "pos");
			if (posRaw == null || !Number.isFinite(posRaw)) continue;
			const pos = Math.max(0, Math.min(1, posRaw / 100000));
			const colour = this.parseColor(gs);
			if (!colour) continue;
			stops.push({ pos, colour });
		}
		if (stops.length === 0) return null;

		const lin = xml.element(el, "lin");
		const pathEl = xml.element(el, "path");

		if (pathEl) {
			const path = xml.attr(pathEl, "path");
			const radialPath: 'circle' | 'rect' = path === 'rect' ? 'rect' : 'circle';
			return { kind: 'radial', stops, path: radialPath };
		}

		let angle = 0;
		if (lin) {
			const ang = xml.intAttr(lin, "ang");
			if (ang != null && Number.isFinite(ang)) {
				// DOCX stores angles in 60000ths of a degree, measured
				// clockwise from 3-o'clock (east). SVG's linear gradient
				// default (x1=0,y1=0 → x2=1,y2=0) is left→right, which
				// is also 0°. So the raw degree conversion is what the
				// renderer wants.
				angle = (ang / 60000) % 360;
			}
		}
		return { kind: 'linear', stops, angle };
	}

	// Allowlist of preset pattern names we know how to draw. Unknown
	// presets fall back to the foreground colour as a solid fill.
	private readonly PATTERN_ALLOWLIST = new Set([
		"dkDnDiag", "ltDnDiag", "dkUpDiag", "ltUpDiag",
		"dkHorz", "ltHorz", "dkVert", "ltVert",
		"cross", "diagCross",
	]);

	// Parses <a:pattFill prst="…"> → PatternFill. The preset name is
	// validated against the allowlist before being stored; unknown
	// presets still produce a PatternFill so the renderer can emit the
	// fallback solid colour.
	private parsePatternFill(el: Element): PatternFill | null {
		const prst = xml.attr(el, "prst") ?? "";
		const preset = this.PATTERN_ALLOWLIST.has(prst) ? prst : "";
		const fgEl = xml.element(el, "fgClr");
		const bgEl = xml.element(el, "bgClr");
		const fg = fgEl ? this.parseColor(fgEl) : null;
		const bg = bgEl ? this.parseColor(bgEl) : null;
		if (!fg && !bg) return null;
		const result: PatternFill = { preset };
		if (fg) result.fg = fg;
		if (bg) result.bg = bg;
		return result;
	}

	// Allowlist of adjustment names we accept on <a:avLst>. Names are
	// matched before being used as Record<string, number> keys in the
	// renderer, so attacker-controlled values (e.g. "__proto__") never
	// reach `Object.assign` / bracket-indexed lookup.
	private readonly ADJUSTMENT_NAME_ALLOWLIST = new Set([
		"adj", "adj1", "adj2", "adj3", "adj4", "adj5", "adj6", "adj7", "adj8",
	]);

	// Parses <a:avLst><a:gd name="adj" fmla="val N"/> into a plain
	// {name: value} record. Only numeric `val N` formulas are accepted;
	// formula references (`*/ 2 3 4` etc.) are ignored in v1.
	private parseAdjustments(avLst: Element | null | undefined): Record<string, number> | undefined {
		if (!avLst) return undefined;
		const out: Record<string, number> = {};
		let any = false;
		for (const gd of xml.elements(avLst, "gd")) {
			const name = xml.attr(gd, "name");
			if (!name || !this.ADJUSTMENT_NAME_ALLOWLIST.has(name)) continue;
			const fmla = xml.attr(gd, "fmla");
			if (!fmla) continue;
			// Only "val N" formulas in v1 — the computed guide forms
			// (+-, *, +/) need a full formula evaluator.
			const m = /^val\s+(-?\d+(?:\.\d+)?)$/.exec(fmla.trim());
			if (!m) continue;
			const n = Number(m[1]);
			if (!Number.isFinite(n)) continue;
			out[name] = n;
			any = true;
		}
		return any ? out : undefined;
	}

	// Parses <a:effectLst> → ShapeEffects. Handles outer/inner shadow,
	// softEdge, glow, and reflection.
	private parseEffectList(el: Element | null | undefined): ShapeEffects | undefined {
		if (!el) return undefined;
		const result: ShapeEffects = {};
		let any = false;
		for (const n of xml.elements(el)) {
			switch (n.localName) {
				case "outerShdw":
				case "innerShdw": {
					const blurRad = xml.intAttr(n, "blurRad");
					const dist = xml.intAttr(n, "dist");
					const dirRaw = xml.intAttr(n, "dir");
					const colour = this.parseColor(n);
					const entry: NonNullable<ShapeEffects["outerShadow"]> = {};
					if (blurRad != null && Number.isFinite(blurRad)) entry.blurRad = blurRad;
					if (dist != null && Number.isFinite(dist)) entry.dist = dist;
					if (dirRaw != null && Number.isFinite(dirRaw)) entry.dir = (dirRaw / 60000) % 360;
					if (colour) entry.colour = colour;
					if (n.localName === "outerShdw") result.outerShadow = entry;
					else result.innerShadow = entry;
					any = true;
					break;
				}
				case "softEdge": {
					const rad = xml.intAttr(n, "rad");
					if (rad != null && Number.isFinite(rad) && rad > 0) {
						result.softEdge = { rad };
						any = true;
					}
					break;
				}
				case "glow": {
					// <a:glow rad="…"> with a colour child. rad is EMU
					// blur radius; omit the entry entirely if rad is
					// missing or non-positive so the renderer doesn't
					// emit an empty filter chain.
					const rad = xml.intAttr(n, "rad");
					if (rad != null && Number.isFinite(rad) && rad > 0) {
						const colour = this.parseColor(n);
						const entry: NonNullable<ShapeEffects["glow"]> = { rad };
						if (colour) entry.colour = colour;
						result.glow = entry;
						any = true;
					}
					break;
				}
				case "reflection": {
					// <a:reflection stA=".." endA=".." stPos=".." endPos=".."
					//   dist=".." dir=".." fadeDir=".." rotWithShape=".."/>
					// All attributes optional; the renderer picks
					// sensible defaults. `dir` / `fadeDir` are in
					// 60000ths of a degree; `stPos` / `endPos` are in
					// 1/1000ths of a percent. `rotWithShape` defaults
					// to true (reflection rotates with the shape).
					const stA = xml.intAttr(n, "stA");
					const endA = xml.intAttr(n, "endA");
					const dist = xml.intAttr(n, "dist");
					const dirRaw = xml.intAttr(n, "dir");
					const fadeDirRaw = xml.intAttr(n, "fadeDir");
					const stPos = xml.intAttr(n, "stPos");
					const endPos = xml.intAttr(n, "endPos");
					const rotWithShape = xml.boolAttr(n, "rotWithShape");
					const entry: NonNullable<ShapeEffects["reflection"]> = {};
					if (stA != null && Number.isFinite(stA)) entry.stA = stA;
					if (endA != null && Number.isFinite(endA)) entry.endA = endA;
					if (dist != null && Number.isFinite(dist)) entry.dist = dist;
					if (dirRaw != null && Number.isFinite(dirRaw)) entry.dir = (dirRaw / 60000) % 360;
					if (fadeDirRaw != null && Number.isFinite(fadeDirRaw)) entry.fadeDir = (fadeDirRaw / 60000) % 360;
					if (stPos != null && Number.isFinite(stPos)) entry.stPos = stPos;
					if (endPos != null && Number.isFinite(endPos)) entry.endPos = endPos;
					if (rotWithShape != null) entry.rotWithShape = rotWithShape;
					result.reflection = entry;
					any = true;
					break;
				}
			}
		}
		return any ? result : undefined;
	}

	// Parses <wps:bodyPr> — text-frame insets. All four sides default
	// to the DrawingML defaults (91440/45720 EMU) at render time.
	private parseBodyPr(elem: Element): DrawingShape["bodyPr"] | undefined {
		if (!elem) return undefined;
		const result: DrawingShape["bodyPr"] = {};
		const lIns = xml.intAttr(elem, "lIns");
		const tIns = xml.intAttr(elem, "tIns");
		const rIns = xml.intAttr(elem, "rIns");
		const bIns = xml.intAttr(elem, "bIns");
		if (lIns != null) result.lIns = lIns;
		if (tIns != null) result.tIns = tIns;
		if (rIns != null) result.rIns = rIns;
		if (bIns != null) result.bIns = bIns;
		return result;
	}

	// Parses <a:custGeom> into a CustomGeometry. Only <a:pathLst> is
	// consumed; <a:avLst>, <a:gdLst>, <a:ahLst>, <a:cxnLst>, and
	// <a:rect> are out of scope for v1 (adjustment values, guide
	// formulas, and connection points parameterise preset geometries
	// rather than the path list; the text rectangle will matter once
	// nested text boxes track it).
	//
	// Each <a:path> contributes one `d` string composed entirely of
	// hard-coded SVG command letters (`M L C Q A Z`) plus numeric
	// tokens parsed via xml.intAttr. A non-finite or missing number
	// collapses to 0. No DOCX-supplied string ever reaches the output
	// `d` attribute.
	private parseCustGeom(
		custGeom: Element | null | undefined,
	): import("./drawing/shapes").CustomGeometry | undefined {
		if (!custGeom) return undefined;
		const paths: { w: number; h: number; d: string }[] = [];
		const pathLst = xml.element(custGeom, "pathLst");
		if (!pathLst) return undefined;
		for (const pathEl of xml.elements(pathLst, "path")) {
			const w = this.safeNum(xml.intAttr(pathEl, "w", 0));
			const h = this.safeNum(xml.intAttr(pathEl, "h", 0));
			if (w <= 0 || h <= 0) continue;
			let d = "";
			// Track the pen position so <a:arcTo> can compute its end
			// point. DrawingML arcTo is relative to the current pen,
			// not absolute.
			let penX = 0;
			let penY = 0;
			for (const cmd of xml.elements(pathEl)) {
				switch (cmd.localName) {
					case "moveTo": {
						const pt = xml.element(cmd, "pt");
						if (!pt) break;
						const x = this.safeNum(xml.intAttr(pt, "x", 0));
						const y = this.safeNum(xml.intAttr(pt, "y", 0));
						d += `M ${x} ${y} `;
						penX = x;
						penY = y;
						break;
					}
					case "lnTo": {
						const pt = xml.element(cmd, "pt");
						if (!pt) break;
						const x = this.safeNum(xml.intAttr(pt, "x", 0));
						const y = this.safeNum(xml.intAttr(pt, "y", 0));
						d += `L ${x} ${y} `;
						penX = x;
						penY = y;
						break;
					}
					case "cubicBezTo": {
						const pts = xml.elements(cmd, "pt");
						if (pts.length < 3) break;
						const x1 = this.safeNum(xml.intAttr(pts[0], "x", 0));
						const y1 = this.safeNum(xml.intAttr(pts[0], "y", 0));
						const x2 = this.safeNum(xml.intAttr(pts[1], "x", 0));
						const y2 = this.safeNum(xml.intAttr(pts[1], "y", 0));
						const x = this.safeNum(xml.intAttr(pts[2], "x", 0));
						const y = this.safeNum(xml.intAttr(pts[2], "y", 0));
						d += `C ${x1} ${y1} ${x2} ${y2} ${x} ${y} `;
						penX = x;
						penY = y;
						break;
					}
					case "quadBezTo": {
						const pts = xml.elements(cmd, "pt");
						if (pts.length < 2) break;
						const x1 = this.safeNum(xml.intAttr(pts[0], "x", 0));
						const y1 = this.safeNum(xml.intAttr(pts[0], "y", 0));
						const x = this.safeNum(xml.intAttr(pts[1], "x", 0));
						const y = this.safeNum(xml.intAttr(pts[1], "y", 0));
						d += `Q ${x1} ${y1} ${x} ${y} `;
						penX = x;
						penY = y;
						break;
					}
					case "arcTo": {
						// DrawingML arc: centred at (penX + wR, penY + hR),
						// starts at `stAng` (60000ths of a degree), sweeps
						// `swAng`. Convert to an SVG A command.
						const wR = this.safeNum(xml.intAttr(cmd, "wR", 0));
						const hR = this.safeNum(xml.intAttr(cmd, "hR", 0));
						const stAng = this.safeNum(xml.intAttr(cmd, "stAng", 0));
						const swAng = this.safeNum(xml.intAttr(cmd, "swAng", 0));
						if (wR === 0 || hR === 0) break;
						const stRad = (stAng / 60000) * (Math.PI / 180);
						const swRad = (swAng / 60000) * (Math.PI / 180);
						// Ellipse centre: the pen sits on the ellipse at
						// angle stAng, so centre = pen - (wR*cos, hR*sin).
						const cx = penX - wR * Math.cos(stRad);
						const cy = penY - hR * Math.sin(stRad);
						const endAng = stRad + swRad;
						const endX = cx + wR * Math.cos(endAng);
						const endY = cy + hR * Math.sin(endAng);
						const largeArc = Math.abs(swRad) > Math.PI ? 1 : 0;
						const sweep = swRad >= 0 ? 1 : 0;
						d += `A ${wR} ${hR} 0 ${largeArc} ${sweep} ${endX} ${endY} `;
						penX = endX;
						penY = endY;
						break;
					}
					case "close": {
						d += "Z ";
						break;
					}
				}
			}
			d = d.trim();
			if (d) {
				paths.push({ w, h, d });
			}
		}
		return { paths };
	}

	// Coerces any number-like to a finite value or 0. Used throughout
	// parseCustGeom so a malformed DOCX attribute can never leak
	// NaN / Infinity into the SVG path data.
	private safeNum(n: number | null | undefined): number {
		return typeof n === "number" && Number.isFinite(n) ? n : 0;
	}

	parseDrawingShape(elem: Element): DrawingShape {
		const result: DrawingShape = {
			type: DomType.DrawingShape,
			children: [],
		};

		const spPr = xml.element(elem, "spPr");
		if (spPr) {
			const xfrm = xml.element(spPr, "xfrm");
			result.xfrm = this.parseXfrm(xfrm) ?? { x: 0, y: 0, cx: 0, cy: 0 };

			const prstGeom = xml.element(spPr, "prstGeom");
			if (prstGeom) {
				const prst = xml.attr(prstGeom, "prst");
				// Validate against the allowlist — never interpolate an
				// attacker-controlled string into SVG/CSS.
				result.presetGeometry = this.PRESET_GEOMETRY_ALLOWLIST.has(prst) ? prst : "rect";
				const avLst = xml.element(prstGeom, "avLst");
				const adjustments = this.parseAdjustments(avLst);
				if (adjustments) result.presetAdjustments = adjustments;
			} else if (xml.element(spPr, "custGeom")) {
				// <a:custGeom> — caller-defined path list. Parsed into
				// CustomGeometry; presetGeometry is left blank so the
				// renderer takes the custGeom branch. Leave the plain
				// 'rect' as a safe fallback in case parsing yields no
				// usable paths.
				result.presetGeometry = "rect";
				result.hasCustomGeometry = true;
				const custGeom = this.parseCustGeom(xml.element(spPr, "custGeom"));
				if (custGeom && custGeom.paths.length > 0) {
					result.custGeom = custGeom;
				}
			} else {
				result.presetGeometry = "rect";
			}

			result.fill = this.parseShapeFill(spPr);
			result.stroke = this.parseShapeStroke(spPr);
			const effects = this.parseEffectList(xml.element(spPr, "effectLst"));
			if (effects) result.effects = effects;
		} else {
			result.presetGeometry = "rect";
			result.xfrm = { x: 0, y: 0, cx: 0, cy: 0 };
		}

		// <wps:bodyPr> sits as a sibling of <wps:spPr> inside <wps:wsp>.
		const bodyPr = xml.element(elem, "bodyPr");
		if (bodyPr) {
			result.bodyPr = this.parseBodyPr(bodyPr);
		}

		// <wps:txbx>/<w:txbxContent> — shape-carried text. Parsed via
		// the normal body-element pipeline so its paragraphs inherit
		// body-paragraph sanitisation.
		const txbx = xml.element(elem, "txbx");
		const txbxContent = txbx ? xml.element(txbx, "txbxContent") : null;
		if (txbxContent) {
			result.txbxParagraphs = this.parseBodyElements(txbxContent);
		}

		// TODO: <wps:style> theme-ref resolution. <a:effectLst> is now
		// parsed above (outer/inner shadow + softEdge); glow and
		// reflection remain out of scope.
		return result;
	}

	parseDrawingShapeGroup(elem: Element): DrawingGroup {
		const result: DrawingGroup = {
			type: DomType.DrawingGroup,
			children: [],
		};

		const grpSpPr = xml.element(elem, "grpSpPr");
		if (grpSpPr) {
			const xfrm = xml.element(grpSpPr, "xfrm");
			if (xfrm) {
				result.xfrm = this.parseXfrm(xfrm) ?? { x: 0, y: 0, cx: 0, cy: 0 };
				// <a:chOff> / <a:chExt> are siblings of <a:off>/<a:ext>
				// inside a group's xfrm. They define the child
				// coordinate space.
				const chOff = xml.element(xfrm, "chOff");
				const chExt = xml.element(xfrm, "chExt");
				result.childOffset = {
					x: chOff ? (xml.floatAttr(chOff, "x", 0) || 0) : 0,
					y: chOff ? (xml.floatAttr(chOff, "y", 0) || 0) : 0,
					cx: chExt ? (xml.floatAttr(chExt, "cx", 0) || 0) : (result.xfrm?.cx ?? 0),
					cy: chExt ? (xml.floatAttr(chExt, "cy", 0) || 0) : (result.xfrm?.cy ?? 0),
				};
			}
		}

		// Children: <wps:wsp>, nested <wpg:wgp>, <pic:pic>.
		for (const n of xml.elements(elem)) {
			switch (n.localName) {
				case "wsp":
					result.children.push(this.parseDrawingShape(n));
					break;
				case "wgp":
					result.children.push(this.parseDrawingShapeGroup(n));
					break;
				case "pic":
					result.children.push(this.parsePicture(n));
					break;
			}
		}

		return result;
	}

	parsePicture(elem: Element): IDomImage {
		var result = <IDomImage>{ type: DomType.Image, src: "", cssStyle: {} };
		var blipFill = xml.element(elem, "blipFill");
		var blip = xml.element(blipFill, "blip");
		var srcRect = xml.element(blipFill, "srcRect");

		result.src = xml.attr(blip, "embed");

		// Alt text: pic:nvPicPr/pic:cNvPr/@descr is the in-picture source;
		// a:blip/@descr appears in some newer files. parseDrawingWrapper may
		// overwrite this with wp:docPr/@descr — the authoring-level value.
		const nvPicPr = xml.element(elem, "nvPicPr");
		const cNvPr = nvPicPr ? xml.element(nvPicPr, "cNvPr") : null;
		const picDescr = cNvPr ? xml.attr(cNvPr, "descr") : null;
		const blipDescr = blip ? xml.attr(blip, "descr") : null;
		if (picDescr != null) result.altText = picDescr;
		else if (blipDescr != null) result.altText = blipDescr;

		if (srcRect) {
			result.srcRect = [
				xml.intAttr(srcRect, "l", 0) / 100000,
				xml.intAttr(srcRect, "t", 0) / 100000,
				xml.intAttr(srcRect, "r", 0) / 100000,
				xml.intAttr(srcRect, "b", 0) / 100000,
			];
		}

		var spPr = xml.element(elem, "spPr");
		var xfrm = xml.element(spPr, "xfrm");

		result.cssStyle["position"] = "relative";

		if (xfrm) {
			result.rotation = xml.intAttr(xfrm, "rot", 0) / 60000;

			for (var n of xml.elements(xfrm)) {
				switch (n.localName) {
					case "ext":
						result.cssStyle["width"] = xml.lengthAttr(n, "cx", LengthUsage.Emu);
						result.cssStyle["height"] = xml.lengthAttr(n, "cy", LengthUsage.Emu);
						break;

					case "off":
						result.cssStyle["left"] = xml.lengthAttr(n, "x", LengthUsage.Emu);
						result.cssStyle["top"] = xml.lengthAttr(n, "y", LengthUsage.Emu);
						break;
				}
			}
		}

		return result;
	}

	parseTable(node: Element): WmlTable {
		var result: WmlTable = { type: DomType.Table, children: [] };

		for (const c of xml.elements(node)) {
			switch (c.localName) {
				case "tr":
					result.children.push(this.parseTableRow(c));
					break;

				case "tblGrid":
					result.columns = this.parseTableColumns(c);
					break;

				case "tblPr":
					this.parseTableProperties(c, result);
					break;
			}
		}

		return result;
	}

	parseTableColumns(node: Element): WmlTableColumn[] {
		var result = [];

		for (const n of xml.elements(node)) {
			switch (n.localName) {
				case "gridCol":
					result.push({ width: xml.lengthAttr(n, "w") });
					break;
			}
		}

		return result;
	}

	parseTableProperties(elem: Element, table: WmlTable) {
		table.cssStyle = {};
		table.cellStyle = {};

		this.parseDefaultProperties(elem, table.cssStyle, table.cellStyle, c => {
			switch (c.localName) {
				case "tblStyle":
					table.styleName = xml.attr(c, "val");
					break;

				case "tblLook":
					table.className = values.classNameOftblLook(c);
					break;

				case "tblpPr":
					this.parseTablePosition(c, table);
					break;

				case "tblStyleColBandSize":
					table.colBandSize = xml.intAttr(c, "val");
					break;

				case "tblStyleRowBandSize":
					table.rowBandSize = xml.intAttr(c, "val");
					break;


				case "hidden":
					table.cssStyle["display"] = "none";
					break;

				default:
					return false;
			}

			return true;
		});

		switch (table.cssStyle["text-align"]) {
			case "center":
				delete table.cssStyle["text-align"];
				table.cssStyle["margin-left"] = "auto";
				table.cssStyle["margin-right"] = "auto";
				break;

			case "right":
				delete table.cssStyle["text-align"];
				table.cssStyle["margin-left"] = "auto";
				break;
		}
	}

	parseTablePosition(node: Element, table: WmlTable) {
		var topFromText = xml.lengthAttr(node, "topFromText");
		var bottomFromText = xml.lengthAttr(node, "bottomFromText");
		var rightFromText = xml.lengthAttr(node, "rightFromText");
		var leftFromText = xml.lengthAttr(node, "leftFromText");

		// OOXML w:tblpPr positioning attributes. We honour:
		//   - page anchors: switch to position:absolute with left/top from
		//     tblpX / tblpY (twips → pt).
		//   - text / margin anchors: keep float:left and fold the four
		//     *FromText padding values into the margins (legacy behaviour).
		//   - tblpXSpec = "center" or "right": reset margin-left/right to
		//     auto so the table centers / right-aligns regardless of
		//     float. Spec values beat any numeric tblpX.
		const horzAnchor = xml.attr(node, "horzAnchor");
		const vertAnchor = xml.attr(node, "vertAnchor");
		const tblpXSpec = xml.attr(node, "tblpXSpec");
		const tblpYSpec = xml.attr(node, "tblpYSpec");
		const tblpX = xml.lengthAttr(node, "tblpX");
		const tblpY = xml.lengthAttr(node, "tblpY");

		const pageAnchored = horzAnchor === "page" || vertAnchor === "page";

		if (pageAnchored) {
			table.cssStyle["position"] = "absolute";
			if (tblpX) table.cssStyle["left"] = tblpX;
			if (tblpY) table.cssStyle["top"] = tblpY;
		} else {
			table.cssStyle["float"] = "left";
		}

		table.cssStyle["margin-bottom"] = values.addSize(table.cssStyle["margin-bottom"], bottomFromText);
		table.cssStyle["margin-left"] = values.addSize(table.cssStyle["margin-left"], leftFromText);
		table.cssStyle["margin-right"] = values.addSize(table.cssStyle["margin-right"], rightFromText);
		table.cssStyle["margin-top"] = values.addSize(table.cssStyle["margin-top"], topFromText);

		if (tblpXSpec === "center") {
			table.cssStyle["margin-left"] = "auto";
			table.cssStyle["margin-right"] = "auto";
		} else if (tblpXSpec === "right") {
			table.cssStyle["margin-left"] = "auto";
		}

		// tblpYSpec has no good desktop-web analogue (bottom / inside /
		// outside are page-relative OOXML positions). Only record it for
		// downstream tooling; not projected onto CSS.
		if (tblpYSpec) {
			table.cssStyle["$tblp-y-spec"] = tblpYSpec;
		}
	}

	parseTableRow(node: Element): WmlTableRow {
		var result: WmlTableRow = { type: DomType.Row, children: [] };

		for (const c of xml.elements(node)) {
			switch (c.localName) {
				case "tc":
					result.children.push(this.parseTableCell(c));
					break;

				case "bookmarkStart":
					// Row-level bookmarks carry w:colFirst / w:colLast and
					// denote a column range. Kept in `children` alongside the
					// cells so renderTableRow can project them onto the matching
					// <td>s. Bookmarks without col range still round-trip here
					// but render as a plain inline anchor.
					result.children.push(parseBookmarkStart(c, xml));
					break;

				case "bookmarkEnd":
					result.children.push(parseBookmarkEnd(c, xml));
					break;

				case "trPr":
				case "tblPrEx":
					this.parseTableRowProperties(c, result);
					break;
			}
		}

		return result;
	}

	parseTableRowProperties(elem: Element, row: WmlTableRow) {
		row.cssStyle = this.parseDefaultProperties(elem, {}, null, c => {
			switch (c.localName) {
				case "cnfStyle":
					row.className = values.classNameOfCnfStyle(c);
					break;

				case "tblHeader":
					row.isHeader = xml.boolAttr(c, "val");
					break;

				case "cantSplit":
					// Default "val" for a presence-only boolean element in
					// OOXML is true. Parse explicitly so <w:cantSplit
					// w:val="false"/> also round-trips correctly.
					row.cantSplit = xml.boolAttr(c, "val", true);
					break;

				case "gridBefore":
					row.gridBefore = xml.intAttr(c, "val");
					break;

				case "gridAfter":
					row.gridAfter = xml.intAttr(c, "val");
					break;

				case "ins":
					row.revision = parseRevisionAttrs(c);
					row.rowRevisionKind = "inserted";
					break;

				case "del":
					row.revision = parseRevisionAttrs(c);
					row.rowRevisionKind = "deleted";
					break;

				case "trPrChange":
					row.formattingRevision = parseFormattingRevision(c);
					break;

				default:
					return false;
			}

			return true;
		});
	}

	parseTableCell(node: Element): OpenXmlElement {
		var result: WmlTableCell = { type: DomType.Cell, children: [] };

		for (const c of xml.elements(node)) {
			switch (c.localName) {
				case "tbl":
					result.children.push(this.parseTable(c));
					break;

				case "p":
					result.children.push(this.parseParagraph(c));
					break;

				case "tcPr":
					this.parseTableCellProperties(c, result);
					break;
			}
		}

		return result;
	}

	parseTableCellProperties(elem: Element, cell: WmlTableCell) {
		cell.cssStyle = this.parseDefaultProperties(elem, {}, null, c => {
			switch (c.localName) {
				case "gridSpan":
					cell.span = xml.intAttr(c, "val", null);
					break;

				case "vMerge":
					cell.verticalMerge = xml.attr(c, "val") ?? "continue";
					break;

				case "cnfStyle":
					cell.className = values.classNameOfCnfStyle(c);
					break;

				default:
					return false;
			}

			return true;
		});

		this.parseTableCellVerticalText(elem, cell);
	}

	parseTableCellVerticalText(elem: Element, cell: WmlTableCell) {
		const directionMap = {
			"btLr": {
				writingMode: "vertical-rl",
				transform: "rotate(180deg)"
			},
			"lrTb": {
				writingMode: "vertical-lr",
				transform: "none"
			},
			"tbRl": {
				writingMode: "vertical-rl",
				transform: "none"
			}
		};

		for (const c of xml.elements(elem)) {
			if (c.localName === "textDirection") {
				const direction = xml.attr(c, "val");
				const style = directionMap[direction] || { writingMode: "horizontal-tb" };
				cell.cssStyle["writing-mode"] = style.writingMode;
				cell.cssStyle["transform"] = style.transform;
			}
		}
	}

	parseDefaultProperties(elem: Element, style: Record<string, string> = null, childStyle: Record<string, string> = null, handler: (prop: Element) => boolean = null): Record<string, string> {
		style = style || {};

		for (const c of xml.elements(elem)) {
			if (handler?.(c))
				continue;

			switch (c.localName) {
				case "jc":
					style["text-align"] = values.valueOfJc(c);
					break;

				case "textAlignment":
					style["vertical-align"] = values.valueOfTextAlignment(c);
					break;

				case "color": {
					style["color"] = xmlUtil.colorAttr(c, "val", null, autos.color);
					// Sideband: renderer substitutes the concrete hex (with
					// any themeTint/themeShade applied) back into `color`.
					// Literal w:val wins when both are present.
					const tref = xmlUtil.themeColorReference(c, "themeColor", "val");
					if (tref) style["$themeColor-color"] = tref;
					break;
				}

				case "sz":
					style["font-size"] = style["min-height"] = xml.lengthAttr(c, "val", LengthUsage.FontSize);
					break;

				case "shd":
					values.applyShd(c, style);
					break;

				case "highlight": {
					style["background-color"] = xmlUtil.colorAttr(c, "val", null, autos.highlight);
					const tref = xmlUtil.themeColorReference(c, "themeColor", "val");
					if (tref) style["$themeColor-background-color"] = tref;
					break;
				}

				case "vertAlign":
					//TODO
					// style.verticalAlign = values.valueOfVertAlign(c);
					break;

				case "position":
					style.verticalAlign = xml.lengthAttr(c, "val", LengthUsage.FontSize);
					break;

				case "tcW":
					if (this.options.ignoreWidth)
						break;

				case "tblW":
					style["width"] = values.valueOfSize(c, "w");
					break;

				case "trHeight":
					this.parseTrHeight(c, style);
					break;

				case "strike":
					style["text-decoration"] = xml.boolAttr(c, "val", true) ? "line-through" : "none"
					break;

				case "dstrike":
					// Word's double strikethrough. Combine line-through with the "double"
					// style (identical longhand sinks as the decoration line).
					style["text-decoration"] = xml.boolAttr(c, "val", true) ? "line-through double" : "none";
					break;

				case "b":
					style["font-weight"] = xml.boolAttr(c, "val", true) ? "bold" : "normal";
					break;

				case "i":
					style["font-style"] = xml.boolAttr(c, "val", true) ? "italic" : "normal";
					break;

				case "caps":
					style["text-transform"] = xml.boolAttr(c, "val", true) ? "uppercase" : "none";
					break;

				case "smallCaps":
					style["font-variant"] = xml.boolAttr(c, "val", true) ? "small-caps" : "none";
					break;

				case "u":
					this.parseUnderline(c, style);
					break;

				case "ind":
				case "tblInd":
					this.parseIndentation(c, style);
					break;

				case "rFonts":
					this.parseFont(c, style);
					break;

				case "tblBorders":
					this.parseBorderProperties(c, childStyle || style);
					break;

				case "tblCellSpacing":
					style["border-spacing"] = values.valueOfMargin(c);
					style["border-collapse"] = "separate";
					break;

				case "pBdr":
					this.parseBorderProperties(c, style);
					break;

				case "bdr": {
					style["border"] = values.valueOfBorder(c);
					const tref = values.themeRefOfBorder(c);
					if (tref) style["$themeColor-border"] = tref;
					break;
				}

				case "tcBorders":
					this.parseBorderProperties(c, style);
					break;

				case "vanish":
				case "specVanish":
					// specVanish marks field-code-generated hidden text; for a read-only
					// viewer it collapses to the same display:none as w:vanish.
					if (xml.boolAttr(c, "val", true))
						style["display"] = "none";
					break;

				case "kern":
					// w:kern@val is the half-point threshold above which kerning kicks in.
					// We can't know the rendered font size at parse time (it may come
					// from a style inherited later), so v1 just enables kerning whenever
					// the property is present. CSS default is "auto" so absence of
					// property leaves the browser to decide.
					if (xml.boolAttr(c, "val", true))
						style["font-kerning"] = "normal";
					break;

				case "w":
					// Character scale — integer percentage (e.g. val="150" == 150%).
					// font-stretch accepts a percentage in CSS Fonts Level 4.
					{
						const pct = xml.intAttr(c, "val");
						if (pct != null)
							style["font-stretch"] = `${pct}%`;
					}
					break;

				case "emboss":
					// Light highlight top-left, dark shadow bottom-right.
					if (xml.boolAttr(c, "val", true))
						style["text-shadow"] = "1px 1px 1px #fff, -1px -1px 1px #000";
					break;

				case "imprint":
					// Inverse of emboss: dark top-left, light bottom-right.
					if (xml.boolAttr(c, "val", true))
						style["text-shadow"] = "1px 1px 1px #000, -1px -1px 1px #fff";
					break;

				case "outline":
					// WebKit stroke; non-WebKit browsers fall back to normal text.
					if (xml.boolAttr(c, "val", true)) {
						style["-webkit-text-stroke"] = "1px currentColor";
						style["color"] = "transparent";
					}
					break;

				case "shadow":
					if (xml.boolAttr(c, "val", true))
						style["text-shadow"] = "2px 2px 2px rgba(0,0,0,0.5)";
					break;

				case "noWrap":
					// Applies to table cells (<w:tcPr><w:noWrap/>). A
					// presence-only boolean in OOXML defaults to true.
					if (xml.boolAttr(c, "val", true))
						style["white-space"] = "nowrap";
					break;

				case "tblCellMar":
				case "tcMar":
					this.parseMarginProperties(c, childStyle || style);
					break;

				case "tblLayout":
					style["table-layout"] = values.valueOfTblLayout(c);
					break;

				case "vAlign":
					style["vertical-align"] = values.valueOfTextAlignment(c);
					break;

				case "spacing":
					if (elem.localName == "pPr") {
						this.parseSpacing(c, style);
					} else if (elem.localName == "rPr") {
						// Run-level character spacing (tracking). val is in twentieths
						// of a point. Dxa's mul (0.05) converts twips → pt, identical
						// conversion.
						style["letter-spacing"] = xml.lengthAttr(c, "val", LengthUsage.Dxa);
					}
					break;

				case "wordWrap":
					if (xml.boolAttr(c, "val")) //TODO: test with examples
						style["overflow-wrap"] = "break-word";
					break;

				case "suppressAutoHyphens":
					style["hyphens"] = xml.boolAttr(c, "val", true) ? "none" : "auto";
					break;

				case "lang":
					style["$lang"] = xml.attr(c, "val");
					break;

				case "rtl":
				case "bidi":
					if (xml.boolAttr(c, "val", true))
						style["direction"] = "rtl";
					break;

				case "cs":
					// Run marked as "complex script". For a read-only viewer
					// without a Unicode character-class database, the best we
					// can do is surface a direction hint — renderRun / rPr
					// consumers on an RTL paragraph will already have
					// direction:rtl. We avoid forcing rtl here (cs can mean
					// Arabic/Hebrew but also Thai/Vietnamese, which are LTR),
					// so this is a no-op beyond acknowledgement.
					break;

				case "bCs":
					// Complex-script bold. v1: mirror w:b.
					style["font-weight"] = xml.boolAttr(c, "val", true) ? "bold" : "normal";
					break;

				case "iCs":
					// Complex-script italic. v1: mirror w:i.
					style["font-style"] = xml.boolAttr(c, "val", true) ? "italic" : "normal";
					break;

				case "szCs": {
					// Complex-script font size (half-points). Without a per-
					// character script classifier we can't selectively apply
					// szCs only to complex-script chars, so v1 treats it as
					// a secondary font-size — recorded in the side channel
					// $cs-font-size and also set as font-size when no w:sz
					// has been seen. DOCX value is numeric via lengthAttr.
					const csSize = xml.lengthAttr(c, "val", LengthUsage.FontSize);
					if (csSize) {
						style["$cs-font-size"] = csSize;
						if (!style["font-size"])
							style["font-size"] = csSize;
					}
					break;
				}

				case "em": {
					// w:em — East-Asian emphasis marks (dot/comma/circle/
					// underDot). DOCX-derived val goes through an explicit
					// allowlist before being mapped to a hand-picked CSS
					// string, so no raw DOCX text reaches the style sink.
					const raw = xml.attr(c, "val");
					switch (raw) {
						case "dot":
							style["text-emphasis"] = "filled dot";
							break;
						case "comma":
							style["text-emphasis"] = "filled sesame";
							break;
						case "circle":
							style["text-emphasis"] = "filled circle";
							break;
						case "underDot":
							style["text-emphasis"] = "filled dot";
							style["text-emphasis-position"] = "under";
							break;
						// "none" or anything else → no-op
					}
					break;
				}

				case "tabs": //ignore - tabs is parsed by other parser
				case "outlineLvl": //TODO
				case "contextualSpacing": //TODO
				case "tblStyleColBandSize": //TODO
				case "tblStyleRowBandSize": //TODO
				case "webHidden": //TODO - maybe web-hidden should be implemented
				case "pageBreakBefore": //TODO - maybe ignore
				case "suppressLineNumbers": //TODO - maybe ignore
				case "keepLines": //TODO - maybe ignore
				case "keepNext": //TODO - maybe ignore
				case "widowControl": //TODO - maybe ignore
				case "noProof": //ignore spellcheck
					//TODO ignore
					break;

				default:
					if (this.options.debug)
						console.warn(`DOCX: Unknown document element: ${elem.localName}.${c.localName}`);
					break;
			}
		}

		return style;
	}

	parseUnderline(node: Element, style: Record<string, string>) {
		var val = xml.attr(node, "val");

		if (val == null)
			return;

		switch (val) {
			case "dash":
			case "dashDotDotHeavy":
			case "dashDotHeavy":
			case "dashedHeavy":
			case "dashLong":
			case "dashLongHeavy":
			case "dotDash":
			case "dotDotDash":
				style["text-decoration"] = "underline dashed";
				break;

			case "dotted":
			case "dottedHeavy":
				style["text-decoration"] = "underline dotted";
				break;

			case "double":
				style["text-decoration"] = "underline double";
				break;

			case "single":
			case "thick":
				style["text-decoration"] = "underline";
				break;

			case "wave":
			case "wavyDouble":
			case "wavyHeavy":
				style["text-decoration"] = "underline wavy";
				break;

			case "words":
				style["text-decoration"] = "underline";
				break;

			case "none":
				style["text-decoration"] = "none";
				break;
		}

		var col = xmlUtil.colorAttr(node, "color");

		if (col)
			style["text-decoration-color"] = col;
	}

	parseFont(node: Element, style: Record<string, string>) {
		var ascii = xml.attr(node, "ascii");
		var asciiTheme = values.themeValue(node, "asciiTheme");
		var eastAsia = xml.attr(node, "eastAsia");
		var fonts = [ascii, asciiTheme, eastAsia].filter(x => x).map(x => encloseFontFamily(x));

		if (fonts.length > 0)
			style["font-family"] = [...new Set(fonts)].join(', ');
	}

	parseIndentation(node: Element, style: Record<string, string>) {
		var firstLine = xml.lengthAttr(node, "firstLine");
		var hanging = xml.lengthAttr(node, "hanging");
		var left = xml.lengthAttr(node, "left");
		var start = xml.lengthAttr(node, "start");
		var right = xml.lengthAttr(node, "right");
		var end = xml.lengthAttr(node, "end");

		if (firstLine) style["text-indent"] = firstLine;
		if (hanging) style["text-indent"] = `-${hanging}`;
		if (left || start) style["margin-inline-start"] = left || start;
		if (right || end) style["margin-inline-end"] = right || end;
	}

	parseSpacing(node: Element, style: Record<string, string>) {
		var before = xml.lengthAttr(node, "before");
		var after = xml.lengthAttr(node, "after");
		var line = xml.intAttr(node, "line", null);
		var lineRule = xml.attr(node, "lineRule");

		if (before) style["margin-top"] = before;
		if (after) style["margin-bottom"] = after;

		if (line !== null) {
			switch (lineRule) {
				case "auto":
					style["line-height"] = `${(line / 240).toFixed(2)}`;
					break;

				case "atLeast":
					style["line-height"] = `calc(100% + ${line / 20}pt)`;
					break;

				default:
					style["line-height"] = style["min-height"] = `${line / 20}pt`
					break;
			}
		}
	}

	parseMarginProperties(node: Element, output: Record<string, string>) {
		for (const c of xml.elements(node)) {
			switch (c.localName) {
				case "left":
					output["padding-left"] = values.valueOfMargin(c);
					break;

				case "right":
					output["padding-right"] = values.valueOfMargin(c);
					break;

				case "top":
					output["padding-top"] = values.valueOfMargin(c);
					break;

				case "bottom":
					output["padding-bottom"] = values.valueOfMargin(c);
					break;
			}
		}
	}

	parseTrHeight(node: Element, output: Record<string, string>) {
		switch (xml.attr(node, "hRule")) {
			case "exact":
				output["height"] = xml.lengthAttr(node, "val");
				break;

			case "atLeast":
			default:
				output["height"] = xml.lengthAttr(node, "val");
				// min-height doesn't work for tr
				//output["min-height"] = xml.sizeAttr(node, "val");  
				break;
		}
	}

	parseBorderProperties(node: Element, output: Record<string, string>) {
		const setBorder = (prop: string, c: Element) => {
			output[prop] = values.valueOfBorder(c);
			const tref = values.themeRefOfBorder(c);
			if (tref) output[`$themeColor-${prop}`] = tref;
		};
		for (const c of xml.elements(node)) {
			switch (c.localName) {
				case "start":
				case "left":
					setBorder("border-left", c);
					break;

				case "end":
				case "right":
					setBorder("border-right", c);
					break;

				case "top":
					setBorder("border-top", c);
					break;

				case "bottom":
					setBorder("border-bottom", c);
					break;

				// Diagonal borders. These don't map to any CSS border
				// property, so we stash them under `$`-prefixed metadata
				// keys (which styleToString + html.ts both skip when
				// applying styles) and let renderTableCell read them back
				// to emit an inline SVG overlay.
				case "tl2br":
					output["$diag-tlbr"] = values.valueOfBorder(c);
					break;

				case "tr2bl":
					output["$diag-trbl"] = values.valueOfBorder(c);
					break;
			}
		}
	}
}

const knownColors = ['black', 'blue', 'cyan', 'darkBlue', 'darkCyan', 'darkGray', 'darkGreen', 'darkMagenta', 'darkRed', 'darkYellow', 'green', 'lightGray', 'magenta', 'none', 'red', 'white', 'yellow'];

class xmlUtil {
	static colorAttr(node: Element, attrName: string, defValue: string = null, autoColor: string = 'black', themeAttrName: string = "themeColor") {
		var v = xml.attr(node, attrName);

		if (v) {
			if (v == "auto") {
				return autoColor;
			} else if (knownColors.includes(v)) {
				return v;
			}

			return `#${v}`;
		}

		var themeColor = xml.attr(node, themeAttrName);

		return themeColor ? `var(--docx-${themeColor}-color)` : defValue;
	}

	// Builds a `$themeColor` sideband placeholder for a `w:color` /
	// `w:shd` / `w:bdr` element. Returns null when the element carries
	// no themeColor/themeFill attribute, the slot is not allowlisted,
	// or a literal `w:val`/`w:fill` hex is present — the literal wins
	// (matching xmlUtil.colorAttr's fallback order and Word's own
	// rendering precedence).
	static themeColorReference(node: Element, themeAttrName: string = "themeColor", literalAttrName?: string): string | null {
		if (literalAttrName) {
			const literal = xml.attr(node, literalAttrName);
			if (literal && literal !== "auto") return null;
		}
		const slot = xml.attr(node, themeAttrName);
		if (!slot) return null;
		const tint = xml.attr(node, "themeTint");
		const shade = xml.attr(node, "themeShade");
		return buildThemeColorReference(slot, tint, shade);
	}
}

class values {
	static themeValue(c: Element, attr: string) {
		var val = xml.attr(c, attr);
		return val ? `var(--docx-${val}-font)` : null;
	}

	static valueOfSize(c: Element, attr: string) {
		var type = LengthUsage.Dxa;

		switch (xml.attr(c, "type")) {
			case "dxa": break;
			case "pct": type = LengthUsage.Percent; break;
			case "auto": return "auto";
		}

		return xml.lengthAttr(c, attr, type);
	}

	static valueOfMargin(c: Element) {
		return xml.lengthAttr(c, "w");
	}

	// w:shd pattern → CSS background. For v1:
	//   - clear / nil / null             → plain background-color from fill
	//   - pctN (where N is a known digit) → color-mix fill/color at N%
	//   - diagStripe / horzStripe /
	//     vertStripe / diagCross /
	//     thinDiagStripe / thinHorzStripe /
	//     thinVertStripe / thinDiagCross  → repeating-linear-gradient
	//   - anything else                   → plain background-color
	// Pattern keys go through an allowlist (SHD_PATTERN_RE below) before
	// selecting a gradient template; the fill/color colours pass through
	// xmlUtil.colorAttr which already constrains the output to hex / known
	// keyword / var(--docx-theme-…). No attacker DOCX string is ever
	// interpolated into CSS here. See SECURITY_REVIEW.md #3/#4.
	static applyShd(c: Element, style: Record<string, string>) {
		const fill = xmlUtil.colorAttr(c, "fill", null, autos.shd, "themeFill");
		const color = xmlUtil.colorAttr(c, "color", null, autos.shd, "themeColor");
		const val = xml.attr(c, "val");

		// Baseline fill; pattern templates may override `background` below.
		if (fill != null) {
			style["background-color"] = fill;
			// Sideband: when the fill came via w:themeFill (optionally
			// with w:themeTint / w:themeShade), let the renderer swap
			// the resolved hex in at render time. A literal w:fill wins
			// over the themeFill reference.
			const tref = xmlUtil.themeColorReference(c, "themeFill", "fill");
			if (tref) style["$themeColor-background-color"] = tref;
		}

		// Nothing to pattern if val is missing / clear / nil.
		if (!val || val === "clear" || val === "nil") return;

		// Strict allowlist: pctN, thin-prefixed and plain stripe / cross
		// names, or the raw Word enum values. Anything that doesn't match
		// falls through to the plain fill set above.
		const SHD_PATTERN_RE = /^(pct\d{1,2}|thin[A-Z][A-Za-z]+|[a-z][A-Za-z]+)$/;
		if (!SHD_PATTERN_RE.test(val)) return;

		const base = fill ?? "transparent";
		const fg = color ?? "black";

		// Percent shading: blend fg into the base at N%.  Word's pctN is
		// the strength of the foreground dot pattern, so the result is a
		// fill that is N% foreground and (100-N)% base. color-mix works
		// in all evergreen browsers; older browsers fall back to the
		// plain background-color assigned above.
		const pctMatch = /^pct(\d{1,2})$/.exec(val);
		if (pctMatch) {
			const pct = Math.min(100, Math.max(0, parseInt(pctMatch[1], 10)));
			style["background-color"] = `color-mix(in srgb, ${fg} ${pct}%, ${base})`;
			return;
		}

		// Stripe / cross templates. The numeric widths are hard-coded; no
		// DOCX-derived values land inside the template string.
		const templates: Record<string, string> = {
			horzStripe:     `repeating-linear-gradient(0deg, ${fg} 0 2px, ${base} 2px 6px)`,
			thinHorzStripe: `repeating-linear-gradient(0deg, ${fg} 0 1px, ${base} 1px 3px)`,
			vertStripe:     `repeating-linear-gradient(90deg, ${fg} 0 2px, ${base} 2px 6px)`,
			thinVertStripe: `repeating-linear-gradient(90deg, ${fg} 0 1px, ${base} 1px 3px)`,
			diagStripe:     `repeating-linear-gradient(45deg, ${fg} 0 2px, ${base} 2px 6px)`,
			thinDiagStripe: `repeating-linear-gradient(45deg, ${fg} 0 1px, ${base} 1px 3px)`,
			reverseDiagStripe:     `repeating-linear-gradient(-45deg, ${fg} 0 2px, ${base} 2px 6px)`,
			thinReverseDiagStripe: `repeating-linear-gradient(-45deg, ${fg} 0 1px, ${base} 1px 3px)`,
		};

		const tpl = templates[val];
		if (tpl) {
			style["background-image"] = tpl;
			style["background-color"] = base;
			return;
		}

		// diagCross / horzCross / thin variants: two stacked gradients.
		const crossTemplates: Record<string, string> = {
			diagCross:     `repeating-linear-gradient(45deg, ${fg} 0 2px, transparent 2px 6px), repeating-linear-gradient(-45deg, ${fg} 0 2px, transparent 2px 6px)`,
			thinDiagCross: `repeating-linear-gradient(45deg, ${fg} 0 1px, transparent 1px 3px), repeating-linear-gradient(-45deg, ${fg} 0 1px, transparent 1px 3px)`,
			horzCross:     `repeating-linear-gradient(0deg, ${fg} 0 2px, transparent 2px 6px), repeating-linear-gradient(90deg, ${fg} 0 2px, transparent 2px 6px)`,
			thinHorzCross: `repeating-linear-gradient(0deg, ${fg} 0 1px, transparent 1px 3px), repeating-linear-gradient(90deg, ${fg} 0 1px, transparent 1px 3px)`,
		};

		const crossTpl = crossTemplates[val];
		if (crossTpl) {
			style["background-image"] = crossTpl;
			style["background-color"] = base;
		}
		// Otherwise: unsupported pattern (e.g. unrecognised word); leave
		// the baseline fill in place.
	}

	static valueOfBorder(c: Element) {
		var type = values.parseBorderType(xml.attr(c, "val"));

		if (type == "none")
			return "none";

		var color = xmlUtil.colorAttr(c, "color");
		var size = xml.lengthAttr(c, "sz", LengthUsage.Border);

		// A border element with no `w:color` (and no `w:themeColor`) means
		// "use the default foreground" — same as `auto`. Without this
		// fallback the composite string becomes "<size> <type> null", which
		// browsers reject as invalid CSS and the whole declaration is
		// dropped — the failure mode the corpus's border-around-paragraph
		// fixture hits.
		if (color == null || color == "auto")
			color = autos.borderColor;

		return `${size} ${type} ${color}`;
	}

	// Companion to `valueOfBorder` for the `$themeColor-<prop>` sideband.
	// Borders emit a composite "<size> <type> <color>" string, so when
	// the color portion is theme-bound the renderer needs to know the
	// concrete colour to splice in. Returns null when the border has no
	// themeColor (the composite string's literal colour is used as-is).
	// Matches xmlUtil.colorAttr's fallback order: a literal `color`
	// value that isn't `auto` wins over any themeColor declaration.
	static themeRefOfBorder(c: Element): string | null {
		if (values.parseBorderType(xml.attr(c, "val")) == "none") return null;
		const literal = xml.attr(c, "color");
		if (literal && literal !== "auto") return null;
		return xmlUtil.themeColorReference(c);
	}

	static parseBorderType(type: string) {
		switch (type) {
			case "single": return "solid";
			case "dashDotStroked": return "solid";
			case "dashed": return "dashed";
			case "dashSmallGap": return "dashed";
			case "dotDash": return "dotted";
			case "dotDotDash": return "dotted";
			case "dotted": return "dotted";
			case "double": return "double";
			case "doubleWave": return "double";
			case "inset": return "inset";
			case "nil": return "none";
			case "none": return "none";
			case "outset": return "outset";
			case "thick": return "solid";
			case "thickThinLargeGap": return "solid";
			case "thickThinMediumGap": return "solid";
			case "thickThinSmallGap": return "solid";
			case "thinThickLargeGap": return "solid";
			case "thinThickMediumGap": return "solid";
			case "thinThickSmallGap": return "solid";
			case "thinThickThinLargeGap": return "solid";
			case "thinThickThinMediumGap": return "solid";
			case "thinThickThinSmallGap": return "solid";
			case "threeDEmboss": return "solid";
			case "threeDEngrave": return "solid";
			case "triple": return "double";
			case "wave": return "solid";
		}

		return 'solid';
	}

	static valueOfTblLayout(c: Element) {
		var type = xml.attr(c, "val");
		return type == "fixed" ? "fixed" : "auto";
	}

	static classNameOfCnfStyle(c: Element) {
		return classNameOfCnfStyle(c);
	}

	static valueOfJc(c: Element) {
		var type = xml.attr(c, "val");

		// Map Word justification tokens (w:jc/@w:val) onto CSS text-align
		// values. "both" and "distribute" both render as full
		// justification: "both" justifies every line except the last, and
		// "distribute" additionally justifies the final line by expanding
		// inter-character spacing — CSS has no native "distribute"
		// keyword, so "justify" is the closest renderable approximation.
		// "start" and "end" collapse to physical "left"/"right" for
		// backward compatibility with existing render-test snapshots.
		switch (type) {
			case "start":
			case "left": return "left";
			case "center": return "center";
			case "end":
			case "right": return "right";
			case "both": return "justify";
			case "distribute": return "justify";
		}

		return type;
	}

	static valueOfVertAlign(c: Element, asTagName: boolean = false) {
		var type = xml.attr(c, "val");

		switch (type) {
			case "subscript": return "sub";
			case "superscript": return asTagName ? "sup" : "super";
		}

		return asTagName ? null : type;
	}

	static valueOfTextAlignment(c: Element) {
		var type = xml.attr(c, "val");

		switch (type) {
			case "auto":
			case "baseline": return "baseline";
			case "top": return "top";
			case "center": return "middle";
			case "bottom": return "bottom";
		}

		return type;
	}

	static addSize(a: string, b: string): string {
		if (a == null) return b;
		if (b == null) return a;

		return `calc(${a} + ${b})`; //TODO
	}

	static classNameOftblLook(c: Element) {
		const val = xml.hexAttr(c, "val", 0);
		let className = "";

		if (xml.boolAttr(c, "firstRow") || (val & 0x0020)) className += " first-row";
		if (xml.boolAttr(c, "lastRow") || (val & 0x0040)) className += " last-row";
		if (xml.boolAttr(c, "firstColumn") || (val & 0x0080)) className += " first-col";
		if (xml.boolAttr(c, "lastColumn") || (val & 0x0100)) className += " last-col";
		if (xml.boolAttr(c, "noHBand") || (val & 0x0200)) className += " no-hband";
		if (xml.boolAttr(c, "noVBand") || (val & 0x0400)) className += " no-vband";

		return className.trim();
	}
}
