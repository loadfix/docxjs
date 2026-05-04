// DrawingML preset-geometry renderer.
//
// Parsed <wps:wsp> shapes (rectangles, ellipses, arrows, callouts,
// stars, …) and <wpg:wgp> groups land here. For each shape we emit
// an absolutely-positioned <div> wrapping an inline <svg> containing
// a single <path d="…"> computed from the preset-geometry name. For
// groups we emit a nested <svg viewBox="chOff.x chOff.y chExt.cx
// chExt.cy"> so child coordinates translate for free (same trick
// used by vml/vml.ts for <v:group>).
//
// Security: every DOCX-derived string used in an SVG attribute goes
// through either sanitizeCssColor (fill / stroke colour) or the
// hard-coded PRESET_GEOMETRY_PATHS allowlist (preset name). Numerics
// come in already parsed by xml.intAttr/floatAttr, so no attacker
// string can reach a template literal.

import { sanitizeCssColor } from '../utils';
import { ns } from '../html';
import {
    DrawingShape, DrawingGroup,
    GradientFill, PatternFill, ShapeEffects,
} from '../document/drawing';
import { resolveColour } from './theme';

// Parsed <a:custGeom>. `paths` holds one entry per <a:path> inside
// <a:pathLst>; the `d` string is already in path-local coordinates
// (0..w, 0..h). The renderer scales those to the shape's px render
// size. See customGeometryToSvgPaths below.
//
// Security: every number reaching `d` has been coerced through
// Number.isFinite in the parser; the command letters are hard-coded
// in the parser/renderer. No DOCX string ever reaches an SVG
// attribute via this type.
export interface CustomGeometry {
    paths: Array<{
        w: number;   // path-local coordinate space width
        h: number;   // path-local coordinate space height
        d: string;   // SVG path data string, path-local coords
    }>;
}

// Rescales a custGeom path `d` string from path-local coordinates
// (0..w, 0..h) to the shape's render pixel coordinates. Because the
// `d` string is composed by the parser from a fixed vocabulary of
// command letters plus numeric tokens, we can safely tokenize and
// multiply without risk of injecting attacker-controlled content —
// any non-numeric token that isn't a known command letter is
// dropped.
const VALID_COMMANDS = new Set(['M', 'L', 'C', 'Q', 'A', 'Z']);

function scalePathD(
    d: string,
    pathW: number,
    pathH: number,
    renderW: number,
    renderH: number,
): string {
    if (pathW <= 0 || pathH <= 0) return '';
    const sx = renderW / pathW;
    const sy = renderH / pathH;
    // Tokens are whitespace-separated in parser output. Inside an
    // arc's 6-tuple, args 3 (x-axis-rotation) and 4/5 (large-arc /
    // sweep flags) are not coordinates; track argument index per
    // command to scale only the positional values.
    const tokens = d.split(/\s+/).filter(Boolean);
    let out = '';
    let cmd = '';
    let argIdx = 0;
    const argsPerCmd: Record<string, number> = {
        M: 2, L: 2, C: 6, Q: 4, A: 7, Z: 0,
    };
    for (const t of tokens) {
        if (VALID_COMMANDS.has(t)) {
            cmd = t;
            argIdx = 0;
            out += (out ? ' ' : '') + cmd;
            continue;
        }
        const n = Number(t);
        if (!Number.isFinite(n)) continue;
        let v = n;
        // For arc commands the 3rd (x-axis rotation), 4th (large-arc
        // flag), and 5th (sweep flag) arguments are not scaled. Radii
        // (args 0, 1) and endpoint (args 5, 6) are scaled.
        if (cmd === 'A') {
            const ai = argIdx % 7;
            if (ai === 0) v = n * sx;        // rx
            else if (ai === 1) v = n * sy;   // ry
            else if (ai === 2) v = n;         // x-axis-rotation
            else if (ai === 3) v = n;         // large-arc-flag
            else if (ai === 4) v = n;         // sweep-flag
            else if (ai === 5) v = n * sx;   // x
            else v = n * sy;                  // y
        } else if (argsPerCmd[cmd]) {
            const ai = argIdx % argsPerCmd[cmd];
            v = ai % 2 === 0 ? n * sx : n * sy;
        }
        out += ' ' + (Number.isFinite(v) ? v : 0);
        argIdx++;
    }
    return out;
}

/**
 * Convert each parsed custGeom path into a render-space SVG `d`
 * string scaled to (renderWidth, renderHeight) px. Returns one entry
 * per path. Empty array if `renderWidth` / `renderHeight` are
 * non-positive.
 */
export function customGeometryToSvgPaths(
    custGeom: CustomGeometry,
    renderWidth: number,
    renderHeight: number,
): string[] {
    if (!custGeom || !custGeom.paths || renderWidth <= 0 || renderHeight <= 0) {
        return [];
    }
    const out: string[] = [];
    for (const p of custGeom.paths) {
        if (!p || !p.d) continue;
        if (!Number.isFinite(p.w) || !Number.isFinite(p.h)) continue;
        if (p.w <= 0 || p.h <= 0) continue;
        const scaled = scalePathD(p.d, p.w, p.h, renderWidth, renderHeight);
        if (scaled) out.push(scaled);
    }
    return out;
}

// Default DrawingML text-frame insets, EMU. Used when <wps:bodyPr>
// omits lIns/tIns/rIns/bIns.
const DEFAULT_INSET_LR_EMU = 91440;
const DEFAULT_INSET_TB_EMU = 45720;

// Preset-geometry path generator. Returns an SVG `d` attribute string
// scaled to the shape's pixel width/height, or null for custGeom /
// unknown preset names. The caller is expected to fall back to a
// plain rectangle when this returns null.
//
// Coordinate conventions: (0, 0) is top-left, (w, h) is bottom-right.
// Every branch of the switch is hard-coded — `prst` is never
// interpolated into the returned string.
export function presetGeometryToSvgPath(
    prst: string,
    w: number,
    h: number,
    adjustments?: Record<string, number> | null,
): string | null {
    if (!prst || w <= 0 || h <= 0) return null;

    // Helpers for the arrow / callout / star paths.
    const cx = w / 2;
    const cy = h / 2;

    // Read a numeric adjustment `name` as a fraction of the shape's
    // `base` dimension. DOCX stores adj values in 100,000ths. Returns
    // `dflt` (also a fraction) when the adjustment is missing or
    // non-finite. Caller is responsible for clamping / using the
    // result.
    const adj = (name: string, dflt: number): number => {
        if (!adjustments) return dflt;
        const raw = adjustments[name];
        if (typeof raw !== 'number' || !Number.isFinite(raw)) return dflt;
        return raw / 100000;
    };

    switch (prst) {
        case 'rect':
            return `M0,0 L${w},0 L${w},${h} L0,${h} Z`;

        case 'roundRect': {
            // <a:avLst><a:gd name="adj" fmla="val N"/> overrides the
            // corner radius. N is in 100,000ths of min(w,h); the
            // default (no avLst) is 10%.
            const frac = Math.max(0, Math.min(0.5, adj('adj', 0.1)));
            const r = Math.min(w, h) * frac;
            return (
                `M${r},0 L${w - r},0 Q${w},0 ${w},${r} ` +
                `L${w},${h - r} Q${w},${h} ${w - r},${h} ` +
                `L${r},${h} Q0,${h} 0,${h - r} ` +
                `L0,${r} Q0,0 ${r},0 Z`
            );
        }

        case 'ellipse':
            // Single cubic-bezier ellipse via the SVG `A` arc command.
            return (
                `M0,${cy} A${cx},${cy} 0 1,0 ${w},${cy} ` +
                `A${cx},${cy} 0 1,0 0,${cy} Z`
            );

        case 'triangle':
            // Isoceles, apex at top-centre.
            return `M${cx},0 L${w},${h} L0,${h} Z`;

        case 'rtTriangle':
            // Right-angle at bottom-left.
            return `M0,0 L0,${h} L${w},${h} Z`;

        case 'diamond':
            return `M${cx},0 L${w},${cy} L${cx},${h} L0,${cy} Z`;

        case 'parallelogram': {
            const skew = w * 0.25;
            return `M${skew},0 L${w},0 L${w - skew},${h} L0,${h} Z`;
        }

        case 'trapezoid': {
            const skew = w * 0.25;
            return `M${skew},0 L${w - skew},0 L${w},${h} L0,${h} Z`;
        }

        case 'pentagon': {
            // Regular pentagon, apex at top.
            const points = regularPolygon(5, cx, cy, w, h, -Math.PI / 2);
            return polygonToPath(points);
        }

        case 'hexagon': {
            // Two flat sides left/right, points top-bottom.
            const dx = w * 0.25;
            return (
                `M${dx},0 L${w - dx},0 L${w},${cy} ` +
                `L${w - dx},${h} L${dx},${h} L0,${cy} Z`
            );
        }

        case 'octagon': {
            const d = w * 0.2929;
            const e = h * 0.2929;
            return (
                `M${d},0 L${w - d},0 L${w},${e} L${w},${h - e} ` +
                `L${w - d},${h} L${d},${h} L0,${h - e} L0,${e} Z`
            );
        }

        case 'line':
            // Diagonal — caller should set fill=none; stroke applies.
            return `M0,0 L${w},${h}`;

        case 'rightArrow': {
            // Shaft height 50% of h; arrow head 40% of w.
            const head = w * 0.6;
            const shaftTop = h * 0.25;
            const shaftBot = h * 0.75;
            return (
                `M0,${shaftTop} L${head},${shaftTop} L${head},0 ` +
                `L${w},${cy} L${head},${h} L${head},${shaftBot} ` +
                `L0,${shaftBot} Z`
            );
        }

        case 'leftArrow': {
            const head = w * 0.4;
            const shaftTop = h * 0.25;
            const shaftBot = h * 0.75;
            return (
                `M${w},${shaftTop} L${head},${shaftTop} L${head},0 ` +
                `L0,${cy} L${head},${h} L${head},${shaftBot} ` +
                `L${w},${shaftBot} Z`
            );
        }

        case 'upArrow': {
            const head = h * 0.4;
            const shaftL = w * 0.25;
            const shaftR = w * 0.75;
            return (
                `M${shaftL},${h} L${shaftL},${head} L0,${head} ` +
                `L${cx},0 L${w},${head} L${shaftR},${head} ` +
                `L${shaftR},${h} Z`
            );
        }

        case 'downArrow': {
            const head = h * 0.6;
            const shaftL = w * 0.25;
            const shaftR = w * 0.75;
            return (
                `M${shaftL},0 L${shaftL},${head} L0,${head} ` +
                `L${cx},${h} L${w},${head} L${shaftR},${head} ` +
                `L${shaftR},0 Z`
            );
        }

        case 'leftRightArrow': {
            const headL = w * 0.2;
            const headR = w * 0.8;
            const shaftTop = h * 0.25;
            const shaftBot = h * 0.75;
            return (
                `M0,${cy} L${headL},0 L${headL},${shaftTop} ` +
                `L${headR},${shaftTop} L${headR},0 L${w},${cy} ` +
                `L${headR},${h} L${headR},${shaftBot} ` +
                `L${headL},${shaftBot} L${headL},${h} Z`
            );
        }

        case 'wedgeRectCallout': {
            // adj1 / adj2 — tail x/y as signed fractions of w/h,
            // measured from the shape centre. A value of 100000
            // = 100% of (w or h) from centre. Defaults place the tail
            // down-left of the rect (Word's out-of-the-box look).
            const tailX = cx + w * adj('adj1', -0.2);
            const tailY = cy + h * adj('adj2', 0.625);
            return (
                `M0,0 L${w},0 L${w},${h} ` +
                `L${w * 0.5},${h} L${tailX},${tailY} L${w * 0.2},${h} ` +
                `L0,${h} Z`
            );
        }

        case 'wedgeRoundRectCallout': {
            // adj1 / adj2 — tail position (see wedgeRectCallout).
            // adj3 — corner-radius fraction (min(w,h) × adj3/100000).
            const r = Math.min(w, h) * Math.max(0, Math.min(0.5, adj('adj3', 0.1)));
            const tailX = cx + w * adj('adj1', -0.2);
            const tailY = cy + h * adj('adj2', 0.625);
            return (
                `M${r},0 L${w - r},0 Q${w},0 ${w},${r} ` +
                `L${w},${h - r} Q${w},${h} ${w - r},${h} ` +
                `L${w * 0.5},${h} L${tailX},${tailY} L${w * 0.2},${h} ` +
                `L${r},${h} Q0,${h} 0,${h - r} ` +
                `L0,${r} Q0,0 ${r},0 Z`
            );
        }

        case 'wedgeEllipseCallout': {
            // Body ellipse with a triangular tail. adj1/adj2 position
            // the tail tip as a signed fraction of w/h from centre.
            const ang = (Math.PI * 5) / 4;
            const ex = cx + Math.cos(ang) * cx;
            const ey = cy + Math.sin(ang) * cy;
            const tailX = cx + w * adj('adj1', -0.2);
            const tailY = cy + h * adj('adj2', 0.625);
            return (
                `M${ex},${ey} L${tailX},${tailY} L${cx * 0.6},${h * 0.95} ` +
                `A${cx},${cy} 0 1,1 ${ex},${ey} Z`
            );
        }

        case 'star5': {
            // adj controls inner-outer vertex ratio (default 0.4).
            const ratio = Math.max(0.1, Math.min(0.9, adj('adj', 0.4)));
            const points = star(5, cx, cy, w, h, ratio, -Math.PI / 2);
            return polygonToPath(points);
        }

        case 'star6': {
            const ratio = Math.max(0.1, Math.min(0.9, adj('adj', 0.5)));
            const points = star(6, cx, cy, w, h, ratio, -Math.PI / 2);
            return polygonToPath(points);
        }

        case 'star8': {
            const ratio = Math.max(0.1, Math.min(0.9, adj('adj', 0.55)));
            const points = star(8, cx, cy, w, h, ratio, -Math.PI / 2);
            return polygonToPath(points);
        }

        case 'cloudCallout': {
            // Cloud is several overlapping arcs; the tail is a small
            // circle off to the bottom-left. A faithful rendering is
            // big — this approximation uses a squashed ellipse for the
            // body and two circles for the tail.
            const tail1X = w * 0.1;
            const tail1Y = h * 1.05;
            const tail2X = w * -0.05;
            const tail2Y = h * 1.2;
            const tail1R = Math.min(w, h) * 0.05;
            const tail2R = Math.min(w, h) * 0.03;
            return (
                `M0,${cy} A${cx},${cy} 0 1,0 ${w},${cy} ` +
                `A${cx},${cy} 0 1,0 0,${cy} Z ` +
                `M${tail1X - tail1R},${tail1Y} ` +
                `A${tail1R},${tail1R} 0 1,0 ${tail1X + tail1R},${tail1Y} ` +
                `A${tail1R},${tail1R} 0 1,0 ${tail1X - tail1R},${tail1Y} Z ` +
                `M${tail2X - tail2R},${tail2Y} ` +
                `A${tail2R},${tail2R} 0 1,0 ${tail2X + tail2R},${tail2Y} ` +
                `A${tail2R},${tail2R} 0 1,0 ${tail2X - tail2R},${tail2Y} Z`
            );
        }

        default:
            // custGeom or unknown preset — caller falls back to <rect>.
            return null;
    }
}

// A regular polygon with `n` points fitted to the (w, h) bounding box.
function regularPolygon(
    n: number,
    cx: number,
    cy: number,
    w: number,
    h: number,
    startAng: number,
): Array<[number, number]> {
    const rx = w / 2;
    const ry = h / 2;
    const pts: Array<[number, number]> = [];
    for (let i = 0; i < n; i++) {
        const ang = startAng + (i * 2 * Math.PI) / n;
        pts.push([cx + Math.cos(ang) * rx, cy + Math.sin(ang) * ry]);
    }
    return pts;
}

// Alternating outer/inner-radius polygon (2*n points). `innerRatio` is
// the inner radius as a fraction of the outer radius.
function star(
    points: number,
    cx: number,
    cy: number,
    w: number,
    h: number,
    innerRatio: number,
    startAng: number,
): Array<[number, number]> {
    const rx = w / 2;
    const ry = h / 2;
    const irx = rx * innerRatio;
    const iry = ry * innerRatio;
    const total = points * 2;
    const pts: Array<[number, number]> = [];
    for (let i = 0; i < total; i++) {
        const outer = i % 2 === 0;
        const ang = startAng + (i * Math.PI) / points;
        const tx = outer ? rx : irx;
        const ty = outer ? ry : iry;
        pts.push([cx + Math.cos(ang) * tx, cy + Math.sin(ang) * ty]);
    }
    return pts;
}

function polygonToPath(points: Array<[number, number]>): string {
    if (points.length === 0) return '';
    const [fx, fy] = points[0];
    let d = `M${fx},${fy}`;
    for (let i = 1; i < points.length; i++) {
        d += ` L${points[i][0]},${points[i][1]}`;
    }
    return d + ' Z';
}

// Emit a <linearGradient> / <radialGradient> into `defs` and return
// its id. Returns null when the gradient has no usable stops.
//
// Security: id is counter-generated (never DOCX-sourced); every colour
// value passes through resolveColour + sanitizeCssColor before being
// set on an SVG attribute.
function appendGradient(
    defs: SVGDefsElement,
    grad: GradientFill,
    context: ShapeRenderContext,
): string | null {
    if (!grad || !grad.stops || grad.stops.length === 0) return null;

    const tag = grad.kind === 'radial' ? 'radialGradient' : 'linearGradient';
    const el = document.createElementNS(ns.svg, tag);
    const id = safeId(context.nextId('docx-grad'));
    if (!id) return null;
    el.setAttribute('id', id);

    if (grad.kind === 'linear') {
        // Convert degrees → unit-square endpoints. 0° = left→right;
        // clockwise positive (matches SVG conventions).
        const ang = typeof grad.angle === 'number' && Number.isFinite(grad.angle)
            ? grad.angle
            : 0;
        const rad = (ang * Math.PI) / 180;
        const cx = 0.5, cy = 0.5;
        const x1 = cx - Math.cos(rad) * 0.5;
        const y1 = cy - Math.sin(rad) * 0.5;
        const x2 = cx + Math.cos(rad) * 0.5;
        const y2 = cy + Math.sin(rad) * 0.5;
        el.setAttribute('x1', x1.toFixed(4));
        el.setAttribute('y1', y1.toFixed(4));
        el.setAttribute('x2', x2.toFixed(4));
        el.setAttribute('y2', y2.toFixed(4));
    } else {
        // Radial defaults: centre at 50,50 with r=50. Caller's `path`
        // hint (circle|rect) affects which gradient type we chose at
        // the tag level — SVG has no rectangular radial gradient so
        // `rect` falls through to radialGradient too.
        el.setAttribute('cx', '0.5');
        el.setAttribute('cy', '0.5');
        el.setAttribute('r', '0.5');
    }

    for (const s of grad.stops) {
        const stop = document.createElementNS(ns.svg, 'stop');
        const pos = Math.max(0, Math.min(1, typeof s.pos === 'number' ? s.pos : 0));
        stop.setAttribute('offset', `${(pos * 100).toFixed(2)}%`);
        const hex = resolveColour(s.colour, context.themePalette);
        const c = sanitizeCssColor(hex);
        stop.setAttribute('stop-color', c ?? '#000000');
        const alpha = s.colour?.mods?.alpha;
        if (typeof alpha === 'number' && Number.isFinite(alpha)) {
            const o = Math.max(0, Math.min(1, alpha / 100000));
            stop.setAttribute('stop-opacity', o.toFixed(3));
        }
        el.appendChild(stop);
    }

    defs.appendChild(el);
    return id;
}

// Emit an SVG <pattern> into `defs` for the supplied PatternFill. The
// resulting pattern uses objectBoundingBox units so the hatch scales
// with the shape — but we also set a 8×8 user-space tile so the
// hatching has a consistent visual density. Returns null when the
// preset isn't in the allowlist (caller then falls back to fg solid).
function appendPattern(
    defs: SVGDefsElement,
    patt: PatternFill,
    context: ShapeRenderContext,
): string | null {
    if (!patt || !patt.preset) return null;

    const fg = sanitizeCssColor(resolveColour(patt.fg, context.themePalette)) ?? '#000000';
    const bg = sanitizeCssColor(resolveColour(patt.bg, context.themePalette)) ?? '#FFFFFF';

    const id = safeId(context.nextId('docx-patt'));
    if (!id) return null;

    // 8×8 tile in userSpaceOnUse — keeps the hatch visually consistent
    // across shape sizes. patternUnits default is objectBoundingBox,
    // which would make a 1×1 tile invisible at common shape sizes.
    const size = 8;
    const pattern = document.createElementNS(ns.svg, 'pattern');
    pattern.setAttribute('id', id);
    pattern.setAttribute('width', String(size));
    pattern.setAttribute('height', String(size));
    pattern.setAttribute('patternUnits', 'userSpaceOnUse');

    // Background fill.
    const bgRect = document.createElementNS(ns.svg, 'rect');
    bgRect.setAttribute('x', '0');
    bgRect.setAttribute('y', '0');
    bgRect.setAttribute('width', String(size));
    bgRect.setAttribute('height', String(size));
    bgRect.setAttribute('fill', bg);
    pattern.appendChild(bgRect);

    // Stroke widths: dk* variants are thicker (~1.5px), lt* thinner (~0.5px).
    const isDk = patt.preset.startsWith('dk');
    const sw = isDk ? 1.5 : 0.75;

    const addLine = (x1: number, y1: number, x2: number, y2: number) => {
        const line = document.createElementNS(ns.svg, 'line');
        line.setAttribute('x1', String(x1));
        line.setAttribute('y1', String(y1));
        line.setAttribute('x2', String(x2));
        line.setAttribute('y2', String(y2));
        line.setAttribute('stroke', fg);
        line.setAttribute('stroke-width', String(sw));
        pattern.appendChild(line);
    };

    switch (patt.preset) {
        case 'dkDnDiag':
        case 'ltDnDiag':
            // Top-left → bottom-right diagonal. Emit two strokes per
            // tile so the hatch tiles cleanly across tile boundaries.
            addLine(-1, 0, size + 1, size + 2);
            addLine(-1, -size, size + 1, 2);
            addLine(-1, size, size + 1, size * 2 + 2);
            break;
        case 'dkUpDiag':
        case 'ltUpDiag':
            // Bottom-left → top-right.
            addLine(-1, size, size + 1, -2);
            addLine(-1, size * 2, size + 1, size - 2);
            addLine(-1, 0, size + 1, -size - 2);
            break;
        case 'dkHorz':
        case 'ltHorz':
            addLine(0, size / 2, size, size / 2);
            break;
        case 'dkVert':
        case 'ltVert':
            addLine(size / 2, 0, size / 2, size);
            break;
        case 'cross':
            addLine(0, size / 2, size, size / 2);
            addLine(size / 2, 0, size / 2, size);
            break;
        case 'diagCross':
            addLine(-1, -1, size + 1, size + 1);
            addLine(-1, size + 1, size + 1, -1);
            break;
        default:
            return null;
    }

    defs.appendChild(pattern);
    return id;
}

// Emit an <a:effectLst> as an SVG <filter> into `defs` and return the
// id. `widthPx` / `heightPx` are the shape's render size — used to
// size the filter region generously so drop-shadows aren't clipped.
function appendEffects(
    defs: SVGDefsElement,
    effects: ShapeEffects,
    widthPx: number,
    heightPx: number,
    context: ShapeRenderContext,
): string | null {
    if (!effects) return null;
    const hasEffect =
        effects.outerShadow || effects.innerShadow || effects.softEdge;
    if (!hasEffect) return null;

    const id = safeId(context.nextId('docx-filter'));
    if (!id) return null;

    const filter = document.createElementNS(ns.svg, 'filter');
    filter.setAttribute('id', id);
    // Expand the filter region to accommodate ~50% shadow offset.
    filter.setAttribute('x', '-50%');
    filter.setAttribute('y', '-50%');
    filter.setAttribute('width', '200%');
    filter.setAttribute('height', '200%');
    filter.setAttribute('filterUnits', 'objectBoundingBox');

    // Outer shadow → feDropShadow. feDropShadow already samples the
    // source alpha, so a single primitive covers the common case.
    if (effects.outerShadow) {
        const s = effects.outerShadow;
        const { dx, dy } = shadowOffset(s.dir, s.dist);
        const blur = emuToShadowPx(s.blurRad) / 2;
        const colour = sanitizeCssColor(
            resolveColour(s.colour, context.themePalette),
        ) ?? '#000000';
        const alpha = s.colour?.mods?.alpha;
        const opacity = typeof alpha === 'number' && Number.isFinite(alpha)
            ? Math.max(0, Math.min(1, alpha / 100000))
            : 0.5;

        const fe = document.createElementNS(ns.svg, 'feDropShadow');
        fe.setAttribute('dx', dx.toFixed(2));
        fe.setAttribute('dy', dy.toFixed(2));
        fe.setAttribute('stdDeviation', blur.toFixed(2));
        fe.setAttribute('flood-color', colour);
        fe.setAttribute('flood-opacity', opacity.toFixed(3));
        filter.appendChild(fe);
    }

    // Inner shadow: a common recipe is SourceAlpha → offset → blur →
    // composite with inverted source. Full emulation is complex; for
    // v1 we degrade to a feDropShadow that sits on top of the fill —
    // visually noticeable but not pixel-perfect. TODO: full inner
    // shadow via feComposite / feFlood.
    if (effects.innerShadow) {
        const s = effects.innerShadow;
        const { dx, dy } = shadowOffset(s.dir, s.dist);
        const blur = emuToShadowPx(s.blurRad) / 2;
        const colour = sanitizeCssColor(
            resolveColour(s.colour, context.themePalette),
        ) ?? '#000000';
        const fe = document.createElementNS(ns.svg, 'feDropShadow');
        fe.setAttribute('dx', dx.toFixed(2));
        fe.setAttribute('dy', dy.toFixed(2));
        fe.setAttribute('stdDeviation', blur.toFixed(2));
        fe.setAttribute('flood-color', colour);
        fe.setAttribute('flood-opacity', '0.6');
        filter.appendChild(fe);
    }

    // Soft edge — blur the alpha channel. feGaussianBlur with
    // `in="SourceAlpha"` produces the classic fade.
    if (effects.softEdge) {
        const rad = emuToShadowPx(effects.softEdge.rad);
        const blur = document.createElementNS(ns.svg, 'feGaussianBlur');
        blur.setAttribute('in', 'SourceGraphic');
        blur.setAttribute('stdDeviation', rad.toFixed(2));
        filter.appendChild(blur);
    }

    defs.appendChild(filter);
    return id;
}

function shadowOffset(dirDeg: number | undefined, distEmu: number | undefined): { dx: number; dy: number } {
    const dir = typeof dirDeg === 'number' && Number.isFinite(dirDeg) ? dirDeg : 0;
    const distPx = emuToShadowPx(distEmu);
    const rad = (dir * Math.PI) / 180;
    return {
        dx: Math.cos(rad) * distPx,
        dy: Math.sin(rad) * distPx,
    };
}

// Standard DOCX EMU → px conversion (1px = 9525 EMU). Defined locally
// so this file doesn't need to call back into the renderer for a
// simple constant.
function emuToShadowPx(emu: number | undefined): number {
    if (typeof emu !== 'number' || !Number.isFinite(emu)) return 0;
    return emu / 9525;
}

// Belt-and-braces: the id returned by `context.nextId` is composed
// entirely from the prefix we pass and a monotonic integer. This
// re-validates before it reaches an `id=` attribute so any future
// regression that lets DOCX data near the prefix is caught here.
function safeId(id: string | null | undefined): string | null {
    if (typeof id !== 'string') return null;
    if (!/^[a-z][a-z0-9-]*$/i.test(id)) return null;
    return id;
}

// Callback used by renderShape to render <wps:txbx> paragraphs via
// the main HtmlRenderer pipeline (inheriting all body-paragraph
// sanitisation).
export type TextRenderer = (paragraphs: any[]) => Node[];

// Per-render-pass context threading counter-generated ids and the
// active theme palette through renderShape / renderShapeGroup. Callers
// in html-renderer.ts build one of these once per document render.
export interface ShapeRenderContext {
    // Monotonic id generator. Returns strings like `docx-grad-1`,
    // `docx-patt-2`, etc. — never touched by DOCX-sourced data, so the
    // returned id is always safe to interpolate into a CSS selector or
    // SVG attribute.
    nextId(prefix: string): string;
    // Theme palette lookup. Returns the resolved 6-digit hex for a
    // scheme slot (`accent1`, `dk1`, …) or undefined when the slot is
    // not present. The underlying data comes from DOCX; resolveColour
    // re-validates it before emitting.
    themePalette?: Record<string, string> | null;
}

// Default context used when the caller doesn't supply one (harness /
// standalone tests). The counter is module-local so repeated invocations
// still get unique ids.
let defaultIdCounter = 0;
function defaultContext(): ShapeRenderContext {
    return {
        nextId(prefix: string) {
            defaultIdCounter += 1;
            return `${prefix}-${defaultIdCounter}`;
        },
    };
}

// Sanitised CSS-identifier prefixes we are willing to prepend to the
// generated counter. The renderer only ever passes literal strings
// (`docx-grad`, `docx-patt`, `docx-filter`) so this just guards against
// future refactors; DOCX strings do not flow here.
const SAFE_ID_PREFIX_RE = /^[a-z][a-z0-9-]*$/;

/**
 * Render a single DrawingML shape. Returns a positioned <div>
 * containing one inline <svg> (the preset-geometry outline / fill)
 * plus, optionally, a sibling <div class="docx-shape-text"> for the
 * <wps:txbx> paragraphs.
 *
 * All DOCX-derived strings (`fill.color`, `stroke.color`,
 * `presetGeometry`) are validated against `sanitizeCssColor` or the
 * hard-coded allowlist in `presetGeometryToSvgPath` before they
 * reach an SVG attribute. Numerics come in already parsed.
 */
export function renderShape(
    shape: DrawingShape,
    emuToPx: (emu: number) => number,
    renderText?: TextRenderer,
    ctx?: ShapeRenderContext,
): HTMLElement {
    const context = ctx ?? defaultContext();
    const widthPx = emuToPx(shape.xfrm?.cx ?? 0);
    const heightPx = emuToPx(shape.xfrm?.cy ?? 0);
    const leftPx = emuToPx(shape.xfrm?.x ?? 0);
    const topPx = emuToPx(shape.xfrm?.y ?? 0);

    const wrapper = document.createElement('div');
    wrapper.className = 'docx-shape';
    wrapper.style.position = 'absolute';
    wrapper.style.left = `${leftPx.toFixed(2)}px`;
    wrapper.style.top = `${topPx.toFixed(2)}px`;
    wrapper.style.width = `${widthPx.toFixed(2)}px`;
    wrapper.style.height = `${heightPx.toFixed(2)}px`;

    const rot = shape.xfrm?.rot;
    if (rot && !Number.isNaN(rot)) {
        wrapper.style.transform = `rotate(${rot}deg)`;
    }

    // --- SVG outline + fill ---
    const svg = document.createElementNS(ns.svg, 'svg');
    svg.setAttribute('xmlns', ns.svg);
    svg.setAttribute('viewBox', `0 0 ${widthPx} ${heightPx}`);
    svg.setAttribute('width', '100%');
    svg.setAttribute('height', '100%');
    svg.style.position = 'absolute';
    svg.style.inset = '0';
    svg.style.overflow = 'visible';

    // Determine which `d` strings to emit. Custom geometry wins when
    // present; otherwise fall back to the preset allowlist. Final
    // fallback is a plain rectangle so an unknown preset still
    // renders something visible.
    const customDs =
        shape.custGeom && shape.custGeom.paths && shape.custGeom.paths.length > 0
            ? customGeometryToSvgPaths(shape.custGeom, widthPx, heightPx)
            : [];
    const presetD = customDs.length === 0
        ? (presetGeometryToSvgPath(
                shape.presetGeometry || 'rect',
                widthPx, heightPx,
                shape.presetAdjustments,
            )
            ?? presetGeometryToSvgPath('rect', widthPx, heightPx))
        : null;
    const dStrings = customDs.length > 0 ? customDs : (presetD ? [presetD] : []);

    // <defs> block accumulates gradients, patterns, and filters. Only
    // counter-generated ids reach it, so the id attribute can never
    // carry DOCX-controlled data. `defs` is only appended to the SVG
    // when it has at least one child.
    const defs = document.createElementNS(ns.svg, 'defs');

    // --- Fill resolution ---
    let fillAttr = '#4472C4';
    if (shape.fill && shape.fill.type === 'solid') {
        const c = sanitizeCssColor(shape.fill.color);
        fillAttr = c ?? 'none';
    } else if (shape.fill && shape.fill.type === 'none') {
        fillAttr = 'none';
    } else if (shape.fill && shape.fill.type === 'gradient') {
        const id = appendGradient(defs, shape.fill.gradient, context);
        if (id) fillAttr = `url(#${id})`;
    } else if (shape.fill && shape.fill.type === 'pattern') {
        const id = appendPattern(defs, shape.fill.pattern, context);
        if (id) fillAttr = `url(#${id})`;
        else {
            // Unknown preset — fall back to the fg colour as a solid.
            const fg = resolveColour(shape.fill.pattern.fg, context.themePalette);
            const c = sanitizeCssColor(fg);
            if (c) fillAttr = c;
        }
    }
    // Preset `line` is an open path — force fill none regardless of
    // what the author set, otherwise the browser closes the stroke
    // into a triangle and floods it.
    if (shape.presetGeometry === 'line' && customDs.length === 0) {
        fillAttr = 'none';
    }

    let strokeAttr: string | null = '#2F5496';
    let strokeWidthAttr: string | null = '1';
    if (shape.stroke) {
        strokeAttr = null;
        strokeWidthAttr = null;
        const stroke = sanitizeCssColor(shape.stroke.color);
        if (stroke) strokeAttr = stroke;
        if (shape.stroke.width != null && Number.isFinite(shape.stroke.width)) {
            const wPx = emuToPx(shape.stroke.width);
            strokeWidthAttr = `${wPx.toFixed(2)}`;
        }
    }

    // --- Effects (shadow / softEdge) → SVG <filter> ---
    const filterId = shape.effects
        ? appendEffects(defs, shape.effects, widthPx, heightPx, context)
        : null;

    if (defs.childNodes.length > 0) svg.appendChild(defs);

    for (const d of dStrings) {
        const path = document.createElementNS(ns.svg, 'path');
        path.setAttribute('d', d);
        path.setAttribute('fill', fillAttr);
        if (strokeAttr) path.setAttribute('stroke', strokeAttr);
        if (strokeWidthAttr) path.setAttribute('stroke-width', strokeWidthAttr);
        if (filterId) path.setAttribute('filter', `url(#${filterId})`);
        svg.appendChild(path);
    }
    wrapper.appendChild(svg);

    // --- Text frame (<wps:txbx>/<w:txbxContent>) ---
    if (shape.txbxParagraphs && shape.txbxParagraphs.length > 0 && renderText) {
        const text = document.createElement('div');
        text.className = 'docx-shape-text';
        text.style.position = 'absolute';
        text.style.inset = '0';
        text.style.boxSizing = 'border-box';

        const lIns = shape.bodyPr?.lIns ?? DEFAULT_INSET_LR_EMU;
        const tIns = shape.bodyPr?.tIns ?? DEFAULT_INSET_TB_EMU;
        const rIns = shape.bodyPr?.rIns ?? DEFAULT_INSET_LR_EMU;
        const bIns = shape.bodyPr?.bIns ?? DEFAULT_INSET_TB_EMU;
        text.style.paddingLeft = `${emuToPx(lIns).toFixed(2)}px`;
        text.style.paddingTop = `${emuToPx(tIns).toFixed(2)}px`;
        text.style.paddingRight = `${emuToPx(rIns).toFixed(2)}px`;
        text.style.paddingBottom = `${emuToPx(bIns).toFixed(2)}px`;
        text.style.overflow = 'hidden';

        const rendered = renderText(shape.txbxParagraphs);
        for (const n of rendered) {
            if (n) text.appendChild(n);
        }
        wrapper.appendChild(text);
    }

    return wrapper;
}

/**
 * Render a DrawingML shape group. Emits a nested <svg viewBox> so
 * child coordinates declared in the group's child-coord space
 * translate for free — the browser applies the viewBox transform.
 * Non-shape children (e.g. embedded <pic:pic>) are wrapped in a
 * <foreignObject> so their HTML renders inside the SVG coordinate
 * space.
 */
export function renderShapeGroup(
    group: DrawingGroup,
    emuToPx: (emu: number) => number,
    renderChild: (child: any) => Node | null,
    // Context is threaded through for consistency with renderShape —
    // unused at the group level today, but exposed so a future
    // gradient-on-group path doesn't need to change the signature.
    _ctx?: ShapeRenderContext,
): HTMLElement {
    const widthPx = emuToPx(group.xfrm?.cx ?? 0);
    const heightPx = emuToPx(group.xfrm?.cy ?? 0);
    const leftPx = emuToPx(group.xfrm?.x ?? 0);
    const topPx = emuToPx(group.xfrm?.y ?? 0);

    const wrapper = document.createElement('div');
    wrapper.className = 'docx-shape-group';
    wrapper.style.position = 'absolute';
    wrapper.style.left = `${leftPx.toFixed(2)}px`;
    wrapper.style.top = `${topPx.toFixed(2)}px`;
    wrapper.style.width = `${widthPx.toFixed(2)}px`;
    wrapper.style.height = `${heightPx.toFixed(2)}px`;

    // SVG coordinates stay in EMU so child xfrm values (also EMU) map
    // 1:1 into the viewBox. Only the outer wrapper converts to px.
    const chOff = group.childOffset ?? { x: 0, y: 0, cx: group.xfrm?.cx ?? 0, cy: group.xfrm?.cy ?? 0 };
    const svg = document.createElementNS(ns.svg, 'svg');
    svg.setAttribute('xmlns', ns.svg);
    svg.setAttribute(
        'viewBox',
        `${chOff.x} ${chOff.y} ${chOff.cx} ${chOff.cy}`,
    );
    svg.setAttribute('width', '100%');
    svg.setAttribute('height', '100%');
    svg.style.overflow = 'visible';

    // We render each child into a <foreignObject> so the child's own
    // HTML (an <svg> wrapper from renderShape, or an <img> from the
    // pic branch) sits at the right position inside the group's
    // viewBox-defined coord space. Using the child's xfrm directly
    // gives it the group-local EMU position; the viewBox maps that to
    // the wrapper pixels.
    for (const child of group.children ?? []) {
        const node = renderChild(child);
        if (!node) continue;
        const x = (child as any).xfrm?.x ?? 0;
        const y = (child as any).xfrm?.y ?? 0;
        const cx = (child as any).xfrm?.cx ?? chOff.cx;
        const cy = (child as any).xfrm?.cy ?? chOff.cy;

        const fo = document.createElementNS(ns.svg, 'foreignObject');
        fo.setAttribute('x', String(x));
        fo.setAttribute('y', String(y));
        fo.setAttribute('width', String(cx));
        fo.setAttribute('height', String(cy));
        // The child already carries absolute positioning with its own
        // xfrm-derived left/top; strip that so it sits flush in the
        // foreignObject.
        if (node instanceof HTMLElement) {
            node.style.position = 'relative';
            node.style.left = '0';
            node.style.top = '0';
            node.style.width = '100%';
            node.style.height = '100%';
        }
        fo.appendChild(node);
        svg.appendChild(fo);
    }

    wrapper.appendChild(svg);
    return wrapper;
}
