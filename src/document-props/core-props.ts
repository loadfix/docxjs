import { XmlParser } from "../parser/xml-parser";

export interface CorePropsDeclaration {
    title: string,
    description: string,
    subject: string,
    creator: string,
    keywords: string,
    language: string,
    lastModifiedBy: string,
    revision: number,
    // ISO-8601 timestamps from dcterms:created / dcterms:modified. These are
    // parsed as plain strings (not Date objects) — the renderer puts them in
    // a data-* attribute verbatim when Options.emitDocumentProps is on.
    created?: string,
    modified?: string,
}

export function parseCoreProps(root: Element, xmlParser: XmlParser): CorePropsDeclaration {
    const result = <CorePropsDeclaration>{};

    for (let el of xmlParser.elements(root)) {
        switch (el.localName) {
            case "title": result.title = el.textContent; break;
            case "description": result.description = el.textContent; break;
            case "subject": result.subject = el.textContent; break;
            case "creator": result.creator = el.textContent; break;
            case "keywords": result.keywords = el.textContent; break;
            case "language": result.language = el.textContent; break;
            case "lastModifiedBy": result.lastModifiedBy = el.textContent; break;
            case "revision": el.textContent && (result.revision = parseInt(el.textContent)); break;
            // dcterms:created / dcterms:modified — namespaced children on the
            // core-properties root. localName is "created" / "modified".
            case "created": result.created = el.textContent; break;
            case "modified": result.modified = el.textContent; break;
        }
    }

    return result;
}