import { DomType, OpenXmlElement } from './dom';
import { CustomGeometry } from '../drawing/shapes';
import { ColourRef } from '../drawing/theme';

// Parsed <a:gradFill>. Each stop carries a ColourRef (so
// schemeClr / lumMod / lumOff resolution happens at render time once
// the theme palette is known) plus a position in [0..1].
export interface GradientFill {
    kind: 'linear' | 'radial';
    stops: Array<{ pos: number; colour: ColourRef }>;
    // Linear only: final gradient angle in degrees, already converted
    // from DOCX's 60000ths. 0 = left→right.
    angle?: number;
    // Radial only: 'circle' | 'rect' — selects SVG <radialGradient>
    // vs an emulation via <linearGradient>. For v1 we use
    // <radialGradient> for both.
    path?: 'circle' | 'rect';
}

// Parsed <a:pattFill prst="…">. The two colours resolve at render time;
// the hatching itself is emitted from a hard-coded catalogue in
// src/drawing/shapes.ts.
export interface PatternFill {
    preset: string; // validated against PATTERN_ALLOWLIST in the parser
    fg?: ColourRef;
    bg?: ColourRef;
}

// Parsed <a:effectLst>. Handles outer/inner shadow, softEdge, glow,
// and reflection.
export interface ShapeEffects {
    outerShadow?: {
        blurRad?: number; // EMU
        dist?: number;    // EMU
        dir?: number;     // degrees, converted from 60000ths
        colour?: ColourRef;
    };
    innerShadow?: {
        blurRad?: number;
        dist?: number;
        dir?: number;
        colour?: ColourRef;
    };
    softEdge?: {
        rad: number; // EMU
    };
    // <a:glow rad="…"><a:srgbClr|schemeClr/></a:glow>. Rendered as a
    // coloured halo around the shape via an SVG filter chain.
    glow?: {
        rad: number;       // EMU blur radius
        colour?: ColourRef;
    };
    // <a:reflection>. Rendered as a mirrored DOM twin of the shape
    // positioned below, faded via a CSS mask-image gradient. `dir`
    // controls the offset angle of the reflection (defaults to 90° =
    // straight down); `fadeDir` controls the angle of the fade mask
    // independently of `dir`. `stPos` / `endPos` set gradient stops
    // (1/1000ths; 0 = at shape edge, 100000 = fully faded).
    // `rotWithShape=false` counter-rotates the reflection so it stays
    // level when the shape itself is rotated.
    reflection?: {
        stA?: number;          // start alpha, 1000ths (0..100000)
        endA?: number;         // end alpha, 1000ths
        dist?: number;         // distance EMU
        dir?: number;          // degrees, converted from 60000ths
        fadeDir?: number;      // degrees, converted from 60000ths
        stPos?: number;        // 1/1000ths (0..100000)
        endPos?: number;       // 1/1000ths (0..100000)
        rotWithShape?: boolean; // default true
    };
}

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
        | { type: 'none' }
        | { type: 'gradient'; gradient: GradientFill }
        | { type: 'pattern'; pattern: PatternFill };
    stroke?: {
        color?: string;
        width?: number; // EMU
    };
    // Parsed <a:avLst><a:gd name="adj|adj1|adj2|…" fmla="val N"/>.
    // Only numeric `val N` formulas are accepted; keys are matched
    // against a hard-coded allowlist (adj, adj1, adj2, adj3) before
    // being used by presetGeometryToSvgPath.
    presetAdjustments?: Record<string, number>;
    // Parsed <a:effectLst>. Rendered as SVG <filter> elements in the
    // shape's <defs>; the `filter=` attribute references them by
    // counter-generated id.
    effects?: ShapeEffects;
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

// ChartEx reference inside an <a:graphic>/<a:graphicData
// uri="http://schemas.microsoft.com/office/drawingml/2014/chartex">.
// Resolves to a `ChartExPart`; rendered as a placeholder div by the
// DomType.ChartEx branch of html-renderer (never as a real SVG chart
// in v1). See src/charts/chartex-part.ts.
export interface DrawingChartEx extends OpenXmlElement {
    type: DomType.ChartEx;
    // r:id of the chartEx part. Same safety notes as DrawingChart.relId.
    relId: string;
}

// SmartArt placeholder — emitted by parseSmartArtReference when the
// SmartArt <a:graphicData uri="…/diagram"> cannot be replaced by its
// <mc:Fallback> sibling (either none existed or its content was
// unrecognised). The renderer turns this into a bland labelled div
// with `data-smartart-layout` carrying an allowlisted URN so a host
// stylesheet can distinguish SmartArt kinds without granting the
// attacker any CSS write primitive. Real SmartArt rendering remains
// unimplemented. See parseSmartArtReference in document-parser.ts.
export interface DrawingSmartArt extends OpenXmlElement {
    type: DomType.SmartArt;
    // Layout URN parsed out of the referenced /word/diagrams/layoutN.xml
    // part (e.g. "urn:microsoft.com/office/officeart/2005/8/layout/list1").
    // Empty string when the part could not be resolved or the URN did
    // not match the allowlist. Validated against SMARTART_LAYOUT_ALLOWLIST
    // in document-parser.ts before being emitted.
    layoutId?: string;
    // The five r:id values from <dgm:relIds>. Captured so a future
    // layout engine can resolve the diagram parts; currently only
    // layoutId is used for rendering.
    relIds?: {
        dm?: string;
        lo?: string;
        qs?: string;
        cs?: string;
    };
}
