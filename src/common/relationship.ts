import { XmlParser } from "../parser/xml-parser";

export interface Relationship {
    id: string,
    type: RelationshipTypes | string,
    target: string
    targetMode: "" | "External" | string 
}

export enum RelationshipTypes {
    OfficeDocument = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
    FontTable = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable",
    Image = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
    Numbering = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
    Styles = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
    StylesWithEffects = "http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects",
    Theme = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
    Settings = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings",
    WebSettings = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings",
    Hyperlink = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
    Footnotes = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
	Endnotes = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes",
    Footer = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
    Header = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
    ExtendedProperties = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
    CoreProperties = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
	CustomProperties = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/custom-properties",
	Comments = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
    CommentsExtended = "http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
    AltChunk = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk",
    Chart = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
    // Modern 2013+ chart parts. Relationship type differs from Chart;
    // loaded as a `ChartExPart` and rendered as a placeholder. See
    // src/charts/chartex-part.ts.
    ChartEx = "http://schemas.microsoft.com/office/2014/relationships/chartEx",
    // SmartArt — five related parts referenced from <dgm:relIds>:
    //   r:dm → /word/diagrams/data*.xml (data model)
    //   r:lo → /word/diagrams/layout*.xml (layout definition)
    //   r:qs → /word/diagrams/quickStyle*.xml (quick-style)
    //   r:cs → /word/diagrams/colors*.xml (colour transform)
    // A sixth, `diagramDrawing`, may be present on some files and points at
    // an extra cached drawing part that Word uses to persist manual layout
    // overrides. Registering these types makes the package resolver walk
    // their own relationships so any images they embed get preloaded. We
    // do not currently parse the diagram XML itself — the Fallback drawing
    // is what reaches the DOM (see parseSmartArtReference in
    // document-parser.ts). See also TODO.md "SmartArt (<dgm:relIds>)".
    DiagramData = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData",
    DiagramLayout = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout",
    DiagramQuickStyle = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle",
    DiagramColors = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors",
    DiagramDrawing = "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing",
    // Glossary document (building blocks / auto-text). A full document part
    // at /word/glossary/document.xml; docxjs parses it with the same
    // pipeline as the main body but does not render it by default.
    GlossaryDocument = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument",
    // Custom XML data (<w:dataBinding> targets). One or more parts at
    // /customXml/item*.xml, each with a companion itemProps*.xml rel
    // that carries the ds:itemID (the GUID we match against storeItemID
    // on <w:dataBinding>).
    CustomXml = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml",
    CustomXmlProps = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps",
}

export function parseRelationships(root: Element, xml: XmlParser): Relationship[] {
    return xml.elements(root).map(e => <Relationship>{
        id: xml.attr(e, "Id"),
        type: xml.attr(e, "Type"),
        target: xml.attr(e, "Target"),
        targetMode: xml.attr(e, "TargetMode")
    });
}