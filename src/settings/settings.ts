import { DocumentParser } from "../document-parser";
import { Length } from "../document/common";
import { XmlParser } from "../parser/xml-parser";

export interface WmlSettings {
	defaultTabStop: Length;
	footnoteProps: NoteProperties;
	endnoteProps: NoteProperties;
	autoHyphenation: boolean;
	evenAndOddHeaders: boolean;
	// Presence of w:documentProtection means the author set a protection
	// policy. All string attrs are DOCX-derived and therefore attacker
	// controlled — the renderer must never interpolate them into CSS or
	// innerHTML. We allowlist `edit` against a fixed set.
	documentProtection?: DocumentProtection;
}

export interface DocumentProtection {
	// ST_DocProtect: readOnly | trackedChanges | comments | forms | none.
	edit?: 'readOnly' | 'trackedChanges' | 'comments' | 'forms' | 'none';
	enforcement: boolean;
	formatting: boolean;
}

export interface NoteProperties {
	nummeringFormat: string;
	defaultNoteIds: string[];
}

export function parseSettings(elem: Element, xml: XmlParser) {
	var result = {} as WmlSettings; 

	for (let el of xml.elements(elem)) {
		switch(el.localName) {
			case "defaultTabStop": result.defaultTabStop = xml.lengthAttr(el, "val"); break;
			case "footnotePr": result.footnoteProps = parseNoteProperties(el, xml); break;
			case "endnotePr": result.endnoteProps = parseNoteProperties(el, xml); break;
			case "autoHyphenation": result.autoHyphenation = xml.boolAttr(el, "val"); break;
			// `w:evenAndOddHeaders` is a toggle element — presence means "on"
			// unless `w:val="false"` is explicit. Default when absent: false.
			case "evenAndOddHeaders": result.evenAndOddHeaders = xml.boolAttr(el, "val", true); break;
			case "documentProtection": result.documentProtection = parseDocumentProtection(el, xml); break;
		}
	}

    return result;
}

// w:documentProtection carries edit/enforcement metadata for a protected
// document. All values reach the DOM only as sanitised attribute strings
// or allowlisted enum checks — never innerHTML, never className.
const ALLOWED_EDIT = new Set(['readOnly', 'trackedChanges', 'comments', 'forms', 'none']);

export function parseDocumentProtection(elem: Element, xml: XmlParser): DocumentProtection {
	const editAttr = xml.attr(elem, "edit");
	return {
		edit: ALLOWED_EDIT.has(editAttr) ? editAttr as DocumentProtection["edit"] : undefined,
		enforcement: xml.boolAttr(elem, "enforcement", false),
		formatting: xml.boolAttr(elem, "formatting", false),
	};
}

export function parseNoteProperties(elem: Element, xml: XmlParser) {
	var result = {
		defaultNoteIds: []
	} as NoteProperties; 

	for (let el of xml.elements(elem)) {
		switch(el.localName) {
			case "numFmt": 
				result.nummeringFormat = xml.attr(el, "val");
				break;

			case "footnote": 
			case "endnote": 
				result.defaultNoteIds.push(xml.attr(el, "id"));
				break;
		}
	}

    return result;
}