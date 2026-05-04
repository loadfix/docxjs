import { DomType, OpenXmlElement } from './dom';
import { CustomGeometry } from '../drawing/shapes';

// DrawingML shape — the shapes authored as "Insert > Shapes" in Word
// (rectangles, ellipses, arrows, callouts, stars, …). Parsed from
// <wps:wsp> inside an <a:graphic>/<a:graphicData>. Rendered by
// src/drawing/shapes.ts as a positioned <div> containing a single
// inline <svg> with the preset geometry path, plus (optionally) a
// sibling text-frame <div> holding the <wps:txbx> paragraphs.
//
// All numeric values are EMU (914400 EMU = 1 inch). The renderer
// converts with emuToPx = emu => emu / 9525.
export interface DrawingShape extends OpenXmlElement {
    type: DomType.DrawingShape;
    // Preset geometry name from <a:prstGeom prst="…">. Always one of
    // the hard-coded allowlist in presetGeometryToSvgPath(); anything
    // else is coerced to 'rect' by the parser so the value is safe to
    // embed in attribute text. Empty string marks a <a:custGeom>.
    presetGeometry?: string;
    xfrm?: {
        x: number;
        y: number;
        cx: number;
        cy: number;
        // <a:xfrm rot="…"> — 60000ths of a degree. Already converted
        // to degrees by the parser.
        rot?: number;
    };
    fill?:
        | { type: 'solid'; color: string }
        | { type: 'none' };
    stroke?: {
        color?: string;
        width?: number; // EMU
    };
    bodyPr?: {
        lIns?: number; // EMU; default 91440
        tIns?: number;
        rIns?: number;
        bIns?: number;
    };
    // <w:p> paragraphs inside <wps:txbx>/<w:txbxContent>. Rendered via
    // the existing renderElement pipeline so txbx content inherits all
    // body-paragraph sanitisation.
    txbxParagraphs?: OpenXmlElement[];
    // Flag kept for back-compat with earlier wave-3.1 callers; set true
    // whenever custGeom was seen regardless of whether parsing yielded
    // usable paths.
    hasCustomGeometry?: boolean;
    // Parsed <a:custGeom> paths. When present, the renderer emits one
    // <path> per entry (scaled from path-local coords to the shape's
    // render size) and ignores presetGeometry. All values are numeric
    // — the `d` string is composed from hard-coded SVG command letters
    // plus validated numbers, never from DOCX strings.
    custGeom?: CustomGeometry;
}

// DrawingML shape group — <wpg:wgp>. Groups can nest. The group
// establishes a child coordinate space via a:chOff / a:chExt, which
// the renderer emits as the viewBox of a nested <svg> so child
// positions translate for free.
export interface DrawingGroup extends OpenXmlElement {
    type: DomType.DrawingGroup;
    xfrm?: {
        x: number;
        y: number;
        cx: number;
        cy: number;
    };
    childOffset?: {
        x: number;
        y: number;
        cx: number;
        cy: number;
    };
    // Children may be DrawingShape, DrawingGroup, or IDomImage.
    children: OpenXmlElement[];
}

// Chart reference inside an <a:graphic>/<a:graphicData uri="…/chart">.
// The relationship id resolves to a ChartPart via the enclosing part's
// relationship map; the renderer looks that part up and turns it into
// SVG. Parsed by parseChartReference in document-parser.ts; rendered
// by the DomType.Chart branch of html-renderer.
export interface DrawingChart extends OpenXmlElement {
    type: DomType.Chart;
    // r:id of the chart part inside the document's relationships.
    // Attacker-controlled — validated against the rel map on lookup,
    // never interpolated into a CSS class or selector.
    relId: string;
}
