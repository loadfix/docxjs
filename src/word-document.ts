import { OutputType } from "jszip";

import { DocumentParser } from './document-parser';
import { Relationship, RelationshipTypes } from './common/relationship';
import { Part } from './common/part';
import { FontTablePart } from './font-table/font-table';
import { OpenXmlPackage } from './common/open-xml-package';
import { DocumentPart } from './document/document-part';
import { DocumentElement } from './document/document';
import { blobToBase64, resolvePath, splitPath } from './utils';
import { NumberingPart } from './numbering/numbering-part';
import { StylesPart } from './styles/styles-part';
import { FooterPart, HeaderPart } from "./header-footer/parts";
import { ExtendedPropsPart } from "./document-props/extended-props-part";
import { CorePropsPart } from "./document-props/core-props-part";
import { ThemePart } from "./theme/theme-part";
import { EndnotesPart, FootnotesPart } from "./notes/parts";
import { SettingsPart } from "./settings/settings-part";
import { CustomPropsPart } from "./document-props/custom-props-part";
import { CommentsPart } from "./comments/comments-part";
import { CommentsExtendedPart } from "./comments/comments-extended-part";
import { ChartPart } from "./charts/chart-part";
import { ChartExPart } from "./charts/chartex-part";
import {
	DiagramLayoutPart, DiagramDataPart, DiagramQuickStylePart,
	DiagramColorsPart, DiagramDrawingPart,
} from "./smartart/smartart-parts";
import { ContentType } from "./common/content-types";
import { convertVectorImage, detectVectorFormat } from "./common/vector-image";

// Glossary document (/word/glossary/document.xml) — a full WordprocessingML
// part that stores reusable building blocks (AutoText, Quick Parts, cover
// pages, etc). docxjs parses it with the same parseDocumentFile pipeline as
// the main body so all sanitisation paths apply, but doesn't render it in
// the default output; consumers can read WordDocument.glossaryDocument for
// the parsed tree.
export class GlossaryDocumentPart extends Part {
	private _documentParser: DocumentParser;
	body: DocumentElement;

	constructor(pkg: OpenXmlPackage, path: string, parser: DocumentParser) {
		super(pkg, path);
		this._documentParser = parser;
	}

	parseXml(root: Element) {
		// <w:glossaryDocument> wraps a <w:docParts> list of building blocks.
		// The parser can walk the glossary tree but most callers only need
		// the root node; we pass the root element directly so
		// parseDocumentFile can look for <w:body> / <w:background> if Word
		// happens to write them inside the glossary (rare, but harmless).
		this.body = this._documentParser.parseDocumentFile(root);
	}
}

// Custom XML data part (/customXml/item*.xml). Referenced by <w:dataBinding
// w:storeItemID="…"/> on an <w:sdt>. The itemID GUID lives in a sibling
// /customXml/itemProps*.xml rel; we parse the companion props to extract it
// so html-renderer.ts can look the part up by GUID.
//
// Storing the raw XML document lets the SDT resolver evaluate the
// dataBinding XPath against it via document.evaluate. Evaluation is scoped
// to this part's document — never the rendered HTML document.
export class CustomXmlPart extends Part {
	// Parsed XML document for the item*.xml — evaluated against at render
	// time. Kept as an XML Document (not serialised).
	xmlDoc: Document;
	// GUID from /customXml/itemProps*.xml <ds:datastoreItem ds:itemID="{…}"/>.
	// Uppercased, braces preserved, to match the form Word emits.
	itemId: string;

	async load(): Promise<any> {
		this.rels = await this._package.loadRelationships(this.path);
		const xmlText = await this._package.load(this.path);
		if (xmlText) {
			try {
				this.xmlDoc = this._package.parseXmlDocument(xmlText);
			} catch {
				// Malformed custom-xml — leave xmlDoc null; the SDT
				// resolver falls back to the sdtContent placeholder.
				this.xmlDoc = null;
			}
		}
	}

	setItemId(id: string) {
		this.itemId = id ?? null;
	}
}

// Companion itemProps*.xml part — just a lookup from the datastoreItem's
// ds:itemID GUID back to the CustomXmlPart that referenced it.
export class CustomXmlPropsPart extends Part {
	itemId: string;
	parseXml(root: Element) {
		// <ds:datastoreItem ds:itemID="{…}"/>. We read the attribute by
		// localName to sidestep the datastore namespace prefix.
		const attrs = Array.from(root.attributes ?? []);
		const idAttr = attrs.find(a => a.localName === "itemID");
		this.itemId = idAttr?.value ?? null;
	}
}

const topLevelRels = [
	{ type: RelationshipTypes.OfficeDocument, target: "word/document.xml" },
	{ type: RelationshipTypes.ExtendedProperties, target: "docProps/app.xml" },
	{ type: RelationshipTypes.CoreProperties, target: "docProps/core.xml" },
	{ type: RelationshipTypes.CustomProperties, target: "docProps/custom.xml" },
];

// Document-part-scoped relationships we want to pre-load (glossary lives on
// the document's rels, not the top-level .rels). These are probed after the
// main document part is loaded.
const documentPartRels = [
	{ type: RelationshipTypes.GlossaryDocument, target: "glossary/document.xml" },
];

export class WordDocument {
	private _package: OpenXmlPackage;
	private _parser: DocumentParser;
	private _options: any;

	rels: Relationship[];
	parts: Part[] = [];
	partsMap: Record<string, Part> = {};
	contentTypes: ContentType[] = [];

	documentPart: DocumentPart;
	fontTablePart: FontTablePart;
	numberingPart: NumberingPart;
	stylesPart: StylesPart;
	footnotesPart: FootnotesPart;
	endnotesPart: EndnotesPart;
	themePart: ThemePart;
	corePropsPart: CorePropsPart;
	extendedPropsPart: ExtendedPropsPart;
	settingsPart: SettingsPart;
	commentsPart: CommentsPart;
	commentsExtendedPart: CommentsExtendedPart;
	// Glossary document (building blocks) — parsed but not rendered by
	// default; consumers can walk the tree for auto-text / quick-part
	// inspection. See GlossaryDocumentPart above.
	glossaryDocumentPart?: GlossaryDocumentPart;
	// Parsed custom XML parts keyed by their datastoreItem GUID (uppercased,
	// braces preserved). Used to resolve <w:dataBinding w:storeItemID="…"/>
	// on <w:sdt> elements. XPath evaluation happens in html-renderer.ts
	// against the part's XML document — never the rendered HTML document.
	customXmlParts: CustomXmlPart[] = [];

	static async load(blob: Blob | any, parser: DocumentParser, options: any): Promise<WordDocument> {
		var d = new WordDocument();

		d._options = options;
		d._parser = parser;
		d._package = await OpenXmlPackage.load(blob, options);
		d.rels = await d._package.loadRelationships();
		d.contentTypes = await d._package.loadContentTypes();

		await Promise.all(topLevelRels.map(rel => {
			const r = d.rels.find(x => x.type === rel.type) ?? rel; //fallback
			return d.loadRelationshipPart(r.target, r.type);
		}));

		// Glossary document + custom XML parts hang off the main document
		// part's rels, not the package-level .rels. loadRelationshipPart
		// already recursed into those rels when it loaded the main
		// DocumentPart, so both parts types are already in partsMap. We
		// just need to backfill the GUID on each CustomXmlPart from its
		// sibling itemProps part.
		for (const part of d.customXmlParts) {
			const propsRel = part.rels?.find(r => r.type === RelationshipTypes.CustomXmlProps);
			if (!propsRel) continue;
			const [partFolder] = splitPath(part.path);
			const propsPath = resolvePath(propsRel.target, partFolder);
			const propsPart = d.partsMap[propsPath] as CustomXmlPropsPart | undefined;
			if (propsPart?.itemId) {
				part.setItemId(propsPart.itemId);
			}
		}

		if (d.commentsPart) {
			const extComments = d.commentsExtendedPart?.comments ?? [];
			d.commentsPart.buildThreading(extComments);
		}

		return d;
	}

	// Look up a parsed custom XML part by the GUID referenced by a
	// <w:dataBinding w:storeItemID="…"/>. GUIDs are compared
	// case-insensitively and with braces stripped — different producers
	// emit slightly different formatting, but the underlying value is the
	// same. Returns null when no part matches or the part failed to parse.
	findCustomXmlByStoreItemId(storeItemID: string): CustomXmlPart | null {
		if (!storeItemID) return null;
		const norm = normalizeGuid(storeItemID);
		for (const part of this.customXmlParts) {
			if (!part.xmlDoc) continue;
			if (normalizeGuid(part.itemId) === norm) return part;
		}
		return null;
	}

	// Public accessor for the parsed glossary document body (building
	// blocks / auto-text). Undefined when the package has no glossary.
	get glossaryDocument(): DocumentElement | undefined {
		return this.glossaryDocumentPart?.body;
	}

	save(type = "blob"): Promise<any> {
		return this._package.save(type);
	}

	private async loadRelationshipPart(path: string, type: string): Promise<Part> {
		if (this.partsMap[path])
			return this.partsMap[path];

		if (!this._package.get(path))
			return null;

		let part: Part = null;

		switch (type) {
			case RelationshipTypes.OfficeDocument:
				this.documentPart = part = new DocumentPart(this._package, path, this._parser);
				break;

			case RelationshipTypes.FontTable:
				this.fontTablePart = part = new FontTablePart(this._package, path);
				break;

			case RelationshipTypes.Numbering:
				this.numberingPart = part = new NumberingPart(this._package, path, this._parser);
				break;

			case RelationshipTypes.Styles:
				this.stylesPart = part = new StylesPart(this._package, path, this._parser);
				break;

			case RelationshipTypes.Theme:
				this.themePart = part = new ThemePart(this._package, path);
				break;

			case RelationshipTypes.Footnotes:
				this.footnotesPart = part = new FootnotesPart(this._package, path, this._parser);
				break;

			case RelationshipTypes.Endnotes:
				this.endnotesPart = part = new EndnotesPart(this._package, path, this._parser);
				break;

			case RelationshipTypes.Footer:
				part = new FooterPart(this._package, path, this._parser);
				break;

			case RelationshipTypes.Header:
				part = new HeaderPart(this._package, path, this._parser);
				break;

			case RelationshipTypes.CoreProperties:
				this.corePropsPart = part = new CorePropsPart(this._package, path);
				break;

			case RelationshipTypes.ExtendedProperties:
				this.extendedPropsPart = part = new ExtendedPropsPart(this._package, path);
				break;

			case RelationshipTypes.CustomProperties:
				part = new CustomPropsPart(this._package, path);
				break;
	
			case RelationshipTypes.Settings:
				this.settingsPart = part = new SettingsPart(this._package, path);
				break;

			case RelationshipTypes.Comments:
				this.commentsPart = part = new CommentsPart(this._package, path, this._parser);
				break;

			case RelationshipTypes.CommentsExtended:
				this.commentsExtendedPart = part = new CommentsExtendedPart(this._package, path);
				break;

			case RelationshipTypes.Chart:
				// Chart parts live under /word/charts/. The renderer
				// resolves them from the enclosing (document / header /
				// footer) part's relationship map.
				part = new ChartPart(this._package, path);
				break;

			case RelationshipTypes.ChartEx:
				// Modern 2013+ chart parts (sunburst, waterfall, funnel,
				// treemap, ...). Rendered as a labelled placeholder
				// rather than a real chart — see src/charts/chartex-part.ts.
				part = new ChartExPart(this._package, path);
				break;

			case RelationshipTypes.DiagramLayout:
				// SmartArt layout definition. Loaded for the uniqueId
				// URN only; the layout algorithm itself is not
				// implemented. See src/smartart/smartart-parts.ts.
				part = new DiagramLayoutPart(this._package, path);
				break;

			case RelationshipTypes.DiagramData:
				// SmartArt data-model (tree of points with text).
				// Currently loaded only so the package resolver walks
				// its own relationships (embedded images). A future
				// list/hierarchy renderer will consume it.
				part = new DiagramDataPart(this._package, path);
				break;

			case RelationshipTypes.DiagramQuickStyle:
				part = new DiagramQuickStylePart(this._package, path);
				break;

			case RelationshipTypes.DiagramColors:
				part = new DiagramColorsPart(this._package, path);
				break;

			case RelationshipTypes.DiagramDrawing:
				// Microsoft extension: cached drawing with manual
				// layout overrides. Loaded so its image rels resolve.
				part = new DiagramDrawingPart(this._package, path);
				break;

			case RelationshipTypes.GlossaryDocument:
				// Glossary document (building blocks). Parsed through the
				// same parseDocumentFile pipeline as the main body — all
				// sanitisation paths apply — but not rendered by default.
				this.glossaryDocumentPart = part = new GlossaryDocumentPart(this._package, path, this._parser);
				break;

			case RelationshipTypes.CustomXml: {
				// Custom XML data part (item*.xml). Referenced by
				// <w:dataBinding w:storeItemID="…"/> on <w:sdt>.
				const xmlPart = new CustomXmlPart(this._package, path);
				this.customXmlParts.push(xmlPart);
				part = xmlPart;
				break;
			}

			case RelationshipTypes.CustomXmlProps:
				// Companion to CustomXml — parses the datastoreItem's GUID.
				// The load() pass above links the GUID back onto the
				// matching CustomXmlPart.
				part = new CustomXmlPropsPart(this._package, path);
				break;
		}

		if (part == null)
			return Promise.resolve(null);

		this.partsMap[path] = part;
		this.parts.push(part);

		await part.load();

		if (part.rels?.length > 0) {
			const [folder] = splitPath(part.path);
			await Promise.all(part.rels.map(rel => this.loadRelationshipPart(resolvePath(rel.target, folder), rel.type)));
		}

		return part;
	}

	async loadDocumentImage(id: string, part?: Part): Promise<string> {
		const path = this.getPathById(part ?? this.documentPart, id);
		if (!path) return null;
		const blob = await this._package.load(path, "blob");
		return this.blobToImageURL(blob, path);
	}

	async loadNumberingImage(id: string): Promise<string> {
		const path = this.getPathById(this.numberingPart, id);
		if (!path) return null;
		const blob = await this._package.load(path, "blob");
		return this.blobToImageURL(blob, path);
	}

	private async blobToImageURL(blob: Blob, path: string): Promise<string> {
		if (!blob) return null;

		// Browsers can't render legacy WMF / EMF directly. Detect by the
		// part path's extension (hard-coded regex in detectVectorFormat —
		// no DOCX string is passed to any decoder) and route through the
		// vector-image helper, which uses WMFJS / EMFJS when available and
		// falls back to a placeholder SVG otherwise.
		const vector = detectVectorFormat(path);
		if (vector) {
			return convertVectorImage(blob, vector);
		}

		const url = this.blobToURL(blob, path);
		return typeof url === 'string' ? url : await url;
	}

	async loadFont(id: string, key: string): Promise<string> {
		const path = this.getPathById(this.fontTablePart, id);
		if (!path) return null;
		const x = await this._package.load(path, "uint8array");
		return x ? this.blobToURL(new Blob([deobfuscate(x, key)]), path) : x;
	}

	// @internal Dead code. The renderer no longer consumes alt chunks — the
	// previous implementation assigned attacker-controlled HTML to iframe.srcdoc
	// (same-origin, no sandbox). Kept as a separate surface for now, but not
	// called from anywhere in the library. See SECURITY_REVIEW.md #1.
	async loadAltChunk(id: string, part?: Part): Promise<string> {
		const path = this.getPathById(part ?? this.documentPart, id);
		return path ? this._package.load(path, "string") : Promise.resolve(null);
	}

	private blobToURL(blob: Blob, path?: string): string | Promise<string> {
		if (!blob)
			return null;

		if (path) {
			const ct = this.contentTypes.find(x => x.partName === path || (x.extension && path.endsWith(`.${x.extension}`)));
			blob = ct ? new Blob([blob], { type: ct.contentType }) : blob;
		}

		if (this._options.useBase64URL) {
			return blobToBase64(blob);
		}

		return URL.createObjectURL(blob);
	}

	findPartByRelId(id: string, basePart: Part = null) {
		var rel = (basePart.rels ?? this.rels).find(r => r.id == id);
		const folder = basePart ? splitPath(basePart.path)[0] : '';
		return rel ? this.partsMap[resolvePath(rel.target, folder)] : null;
	}

	getPathById(part: Part, id: string): string {
		const rel = part.rels.find(x => x.id == id);
		const [folder] = splitPath(part.path);
		return rel ? resolvePath(rel.target, folder) : null;
	}
}

// Canonicalise a datastoreItem GUID for case-insensitive, brace-insensitive
// comparison. Both `{A1B2C3D4-…}` and `a1b2c3d4-…` are treated as the same
// id. Returns null for null/empty input so two absent ids never "match".
function normalizeGuid(s: string | null | undefined): string | null {
	if (!s) return null;
	return s.replace(/[{}]/g, '').toUpperCase();
}

export function deobfuscate(data: Uint8Array, guidKey: string) {
	const len = 16;
	const trimmed = guidKey.replace(/{|}|-/g, "");
	const numbers = new Array(len);

	for (let i = 0; i < len; i++)
		numbers[len - i - 1] = parseInt(trimmed.substring(i * 2, i * 2 + 2), 16);

	for (let i = 0; i < 32; i++)
		data[i] = data[i] ^ numbers[i % len]

	// FIXME: return type
	return data as any;
}