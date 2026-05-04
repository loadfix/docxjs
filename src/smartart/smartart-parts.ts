import { Part } from "../common/part";
import { OpenXmlPackage } from "../common/open-xml-package";
import xml from "../parser/xml-parser";

// Minimal SmartArt part loaders.
//
// SmartArt is backed by five sibling parts inside /word/diagrams/:
//   data*.xml         — the data-model (tree of points with text)
//   layout*.xml       — the layout definition (algorithm + styling)
//   quickStyle*.xml   — style presets
//   colors*.xml       — colour transform
//   drawing*.xml      — Microsoft extension with cached manual layout
//
// v1 of the SmartArt support only needs the layout URN (to tag the
// placeholder <div> with data-smartart-layout) and the data-model text
// (reserved for a future list/hierarchy renderer — not used yet). The
// other two parts are loaded only so the package resolver walks their
// own relationships and preloads any embedded media.
//
// Security: the layout URN is matched against SMARTART_LAYOUT_ALLOWLIST
// in document-parser.ts before reaching the DOM. Data-model text
// reaches the DOM via textContent only, if ever consumed.

// Matches the canonical SmartArt layout URN shape. Anything that does
// not satisfy this is discarded (never reaches data-*). Kept here
// alongside the part loader so the guard lives with the parse step.
const SMARTART_LAYOUT_URN_RE =
    /^urn:microsoft\.com\/office\/officeart\/\d{4}\/\d+\/layout\/[a-z0-9]+$/i;

export class DiagramLayoutPart extends Part {
    // uniqueId attribute from <dgm:layoutDef>. Empty string when the
    // attribute is missing or does not pass the allowlist regex.
    layoutId: string = "";

    constructor(pkg: OpenXmlPackage, path: string) {
        super(pkg, path);
    }

    protected parseXml(root: Element) {
        // Root element is <dgm:layoutDef uniqueId="urn:...">. Some files
        // wrap it in <mc:AlternateContent>; in that case we still look
        // for the uniqueId on the first descendant <layoutDef>.
        let uniqueId = xml.attr(root, "uniqueId");
        if (!uniqueId) {
            const def = firstDescendant(root, "layoutDef");
            if (def) uniqueId = xml.attr(def, "uniqueId");
        }
        if (uniqueId && SMARTART_LAYOUT_URN_RE.test(uniqueId)) {
            this.layoutId = uniqueId;
        }
    }
}

export class DiagramDataPart extends Part {
    constructor(pkg: OpenXmlPackage, path: string) {
        super(pkg, path);
    }
    // No parsing yet — a future list/hierarchy renderer will reach for
    // <dgm:dataModel><dgm:ptLst><dgm:pt> points here. Keeping the class
    // lets word-document.ts register the part and walk its relationships.
}

export class DiagramQuickStylePart extends Part {
    constructor(pkg: OpenXmlPackage, path: string) {
        super(pkg, path);
    }
}

export class DiagramColorsPart extends Part {
    constructor(pkg: OpenXmlPackage, path: string) {
        super(pkg, path);
    }
}

export class DiagramDrawingPart extends Part {
    constructor(pkg: OpenXmlPackage, path: string) {
        super(pkg, path);
    }
}

function firstDescendant(root: Element, localName: string): Element | null {
    for (const c of xml.elements(root)) {
        if (c.localName === localName) return c;
        const d = firstDescendant(c, localName);
        if (d) return d;
    }
    return null;
}
