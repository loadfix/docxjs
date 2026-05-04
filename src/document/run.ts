import { XmlParser } from "../parser/xml-parser";
import { CommonProperties, parseCommonProperty } from "./common";
import { OpenXmlElement } from "./dom";

export interface WmlRun extends OpenXmlElement, RunProperties {
    id?: string;
    verticalAlign?: string;
	fieldRun?: boolean;
	// w:rPr/w:fitText — target width in twips (1/20 pt) and optional id
	// grouping consecutive fitText runs. Parsed via parseFloat so no raw
	// DOCX string reaches the renderer's CSS sink.
	fitText?: { width: number; id?: string };
	// w:rPr/w:bdo — explicit bidi override. The parser allowlists the raw
	// DOCX val against /^(ltr|rtl)$/; anything else is dropped.
	bidiOverride?: "ltr" | "rtl";
}

export interface RunProperties extends CommonProperties {

}

export function parseRunProperties(elem: Element, xml: XmlParser): RunProperties {
    let result = <RunProperties>{};

    for(let el of xml.elements(elem)) {
        parseRunProperty(el, result, xml);
    }

    return result;
}

export function parseRunProperty(elem: Element, props: RunProperties, xml: XmlParser) {
    if (parseCommonProperty(elem, props, xml))
        return true;

    return false;
}