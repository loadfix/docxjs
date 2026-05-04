import { DomType, OpenXmlElement } from './dom';

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
    // TODO: custGeom path rendering (tracked upstream).
    hasCustomGeometry?: boolean;
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
