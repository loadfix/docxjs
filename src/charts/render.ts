import { sanitizeCssColor } from "../utils";
import {
    ChartAxisColorRef,
    ChartAxisStyle,
    ChartExFunnelModel,
    ChartExHistogramModel,
    ChartExTreeDataModel,
    ChartExTreeNode,
    ChartExWaterfallModel,
    ChartModel,
    ChartSeries,
} from "./model";
import { resolveSchemeColor } from "../drawing/theme";

// Options threaded into the SVG chart renderer. Currently carries only
// the document's theme palette, which axis chrome (lines / tick labels
// / gridlines) resolves schemeClr references against. Series and
// data-point colours are still pre-resolved by chart-part.ts.
export interface RenderChartOptions {
    themePalette?: Record<string, string> | null;
}

// Zero-dependency SVG chart renderer. Consumes the shape produced by
// `src/charts/chart-part.ts` and emits an `<svg>` element into the
// caller's document.
//
// Layout is deliberately simple:
// - Fixed viewBox (600 x 400) so charts scale with their container.
// - Plot area is inset by CHART_PADDING plus a dynamic title / legend /
//   axis-label reserve.
// - No animations, no tooltips, no interactivity — this is a
//   read-only viewer.
//
// Security: every string on ChartModel originates in untrusted DOCX
// XML. Text is emitted via `textContent` (so the browser HTML-encodes
// it); colours round-trip through `sanitizeCssColor`.

const SVG_NS = "http://www.w3.org/2000/svg";

const VIEW_W = 600;
const VIEW_H = 400;
const PADDING = 16;
const TITLE_HEIGHT = 28;
const LEGEND_ROW_HEIGHT = 18;
const AXIS_LABEL_WIDTH = 40;
const AXIS_LABEL_HEIGHT = 24;
const FONT_SIZE = 11;

// Fallback palette used when a series has no <a:srgbClr>. Colours are
// Office-2016-ish defaults; all pre-sanitised so we can emit them
// directly as SVG attribute values.
const DEFAULT_PALETTE = [
    "#4472C4", "#ED7D31", "#A5A5A5", "#FFC000",
    "#5B9BD5", "#70AD47", "#264478", "#9E480E",
    "#636363", "#997300",
];

// Axis-chrome defaults. Applied whenever the axis style's
// corresponding ref is null or fails to resolve to a safe colour.
// Tick labels are mid-grey to match the original behaviour; axis
// lines are a slightly darker grey; gridlines default to none
// (matching the v1 renderer which didn't emit gridlines at all, so
// we preserve that behaviour unless the axis explicitly declared
// <c:majorGridlines>).
const DEFAULT_AXIS_LINE = "#888";
const DEFAULT_TICK_LABEL = "#555";

interface ResolvedAxisStyle {
    line: string;
    tickLabel: string;
    gridline: string | null;
}

function resolveAxisRef(
    ref: ChartAxisColorRef,
    palette: Record<string, string> | null | undefined,
): string | null {
    if (!ref) return null;
    if (ref.kind === "literal") {
        return sanitizeCssColor(ref.color);
    }
    const resolved = resolveSchemeColor(ref.slot, palette ?? undefined);
    return resolved ? sanitizeCssColor(resolved) : null;
}

function resolveAxisStyle(
    style: ChartAxisStyle | undefined,
    palette: Record<string, string> | null | undefined,
): ResolvedAxisStyle {
    const s = style ?? { line: null, tickLabel: null, gridline: null };
    return {
        line: resolveAxisRef(s.line, palette) ?? DEFAULT_AXIS_LINE,
        tickLabel: resolveAxisRef(s.tickLabel, palette) ?? DEFAULT_TICK_LABEL,
        // Gridlines fall back to null so the renderer knows to skip
        // emission entirely when the DOCX didn't declare them.
        gridline: resolveAxisRef(s.gridline, palette),
    };
}

interface AxisColors {
    catAxis: ResolvedAxisStyle;
    valAxis: ResolvedAxisStyle;
}

export function renderChart(model: ChartModel, options?: RenderChartOptions): SVGElement {
    const doc: Document = document;
    const svg = doc.createElementNS(SVG_NS, "svg") as SVGElement;
    svg.setAttribute("viewBox", `0 0 ${VIEW_W} ${VIEW_H}`);
    svg.setAttribute("role", "img");
    svg.setAttribute("preserveAspectRatio", "xMidYMid meet");
    svg.style.maxWidth = "100%";
    svg.style.height = "auto";
    svg.style.display = "block";

    const palette = options?.themePalette ?? null;
    const catAxis = resolveAxisStyle(model.catAxis, palette);
    const valAxis = resolveAxisStyle(model.valAxis, palette);

    const title = (model.title ?? "").trim();
    const titleBottom = title ? PADDING + TITLE_HEIGHT : PADDING;
    if (title) {
        const t = mkText(doc, VIEW_W / 2, PADDING + TITLE_HEIGHT - 10, title, {
            fontSize: 14, fontWeight: "600", anchor: "middle",
        });
        svg.appendChild(t);
    }

    // Compute legend layout (simple single-row wrapped to multiple rows
    // at the bottom of the chart).
    const legendEntries = model.showLegend
        ? model.series.filter((s) => s.values.length > 0)
        : [];
    const legendLayout = layoutLegend(legendEntries, VIEW_W - 2 * PADDING, doc);
    const legendTop = VIEW_H - PADDING - legendLayout.height;

    const plotTop = titleBottom + 4;
    const plotBottom = legendLayout.height > 0
        ? legendTop - 4
        : VIEW_H - PADDING;

    switch (model.kind) {
        case "column":
        case "bar":
            renderBarChart(svg, doc, model, {
                left: PADDING + AXIS_LABEL_WIDTH,
                right: VIEW_W - PADDING,
                top: plotTop,
                bottom: plotBottom - AXIS_LABEL_HEIGHT,
                horizontal: model.kind === "bar",
            }, { catAxis, valAxis });
            break;
        case "line":
            renderLineChart(svg, doc, model, {
                left: PADDING + AXIS_LABEL_WIDTH,
                right: VIEW_W - PADDING,
                top: plotTop,
                bottom: plotBottom - AXIS_LABEL_HEIGHT,
            }, { catAxis, valAxis });
            break;
        case "pie":
            renderPieChart(svg, doc, model, {
                left: PADDING,
                right: VIEW_W - PADDING,
                top: plotTop,
                bottom: plotBottom,
            });
            break;
        default:
            // Unknown / chartEx / 3D — render a placeholder so the
            // reader knows something was there.
            const msg = mkText(doc, VIEW_W / 2, VIEW_H / 2, "[chart]", {
                fontSize: 12, anchor: "middle", fill: "#888",
            });
            svg.appendChild(msg);
            break;
    }

    // Legend last so it sits on top of the plot.
    for (const node of legendLayout.nodes(legendTop)) svg.appendChild(node);

    return svg;
}

// -----------------------------------------------------------------
// Bar / column
// -----------------------------------------------------------------

function renderBarChart(
    svg: SVGElement, doc: Document, model: ChartModel,
    rect: { left: number; right: number; top: number; bottom: number; horizontal: boolean },
    axisColors: AxisColors,
) {
    const series = model.series.filter((s) => s.values.length > 0);
    if (series.length === 0) return;

    const categories = maxLengthCategories(series);
    const catCount = categories.length;
    if (catCount === 0) return;

    const stacked = model.grouping === "stacked" || model.grouping === "percentStacked";
    const percentStacked = model.grouping === "percentStacked";

    // Compute min/max taking stacking into account.
    let valMin = 0;
    let valMax = 0;
    if (stacked) {
        for (let i = 0; i < catCount; i++) {
            let sum = 0;
            for (const s of series) {
                const v = finite(s.values[i]);
                if (v == null) continue;
                sum += v;
            }
            if (percentStacked) sum = sum === 0 ? 0 : 1;
            valMax = Math.max(valMax, sum);
            valMin = Math.min(valMin, sum);
        }
    } else {
        for (const s of series) {
            for (const v of s.values) {
                const f = finite(v);
                if (f == null) continue;
                valMax = Math.max(valMax, f);
                valMin = Math.min(valMin, f);
            }
        }
    }
    if (valMax === valMin) valMax = valMin + 1;
    const { min: yMin, max: yMax, ticks } = niceScale(valMin, valMax, 5);

    const horizontal = rect.horizontal;
    const plotW = rect.right - rect.left;
    const plotH = rect.bottom - rect.top;

    // Axis-line colour is taken from the axis that each edge represents:
    // the value-axis line runs perpendicular to its ticks, and the
    // category axis sits along the plot's edge opposite the tick labels.
    const valLineColor = axisColors.valAxis.line;
    const catLineColor = axisColors.catAxis.line;
    const valTickColor = axisColors.valAxis.tickLabel;
    const catTickColor = axisColors.catAxis.tickLabel;
    const valGrid = axisColors.valAxis.gridline;

    // Axes.
    svg.appendChild(mkLine(doc, rect.left, rect.top, rect.left, rect.bottom, valLineColor));
    svg.appendChild(mkLine(doc, rect.left, rect.bottom, rect.right, rect.bottom, catLineColor));

    // Tick labels (and optional major gridlines) for the value axis.
    for (const t of ticks) {
        if (horizontal) {
            const x = rect.left + ((t - yMin) / (yMax - yMin)) * plotW;
            if (valGrid && t !== yMin) {
                svg.appendChild(mkLine(doc, x, rect.top, x, rect.bottom, valGrid));
            }
            svg.appendChild(mkLine(doc, x, rect.bottom, x, rect.bottom + 4, valLineColor));
            svg.appendChild(mkText(doc, x, rect.bottom + 16, formatTick(t), {
                fontSize: FONT_SIZE, anchor: "middle", fill: valTickColor,
            }));
        } else {
            const y = rect.bottom - ((t - yMin) / (yMax - yMin)) * plotH;
            if (valGrid && t !== yMin) {
                svg.appendChild(mkLine(doc, rect.left, y, rect.right, y, valGrid));
            }
            svg.appendChild(mkLine(doc, rect.left - 4, y, rect.left, y, valLineColor));
            svg.appendChild(mkText(doc, rect.left - 6, y + 4, formatTick(t), {
                fontSize: FONT_SIZE, anchor: "end", fill: valTickColor,
            }));
        }
    }

    // Category slots.
    const slotSize = (horizontal ? plotH : plotW) / catCount;
    const groupPad = Math.min(slotSize * 0.15, 8);
    const innerGroup = slotSize - 2 * groupPad;
    const barsPerSlot = stacked ? 1 : series.length;
    const barSize = Math.max(1, innerGroup / Math.max(1, barsPerSlot));

    for (let i = 0; i < catCount; i++) {
        // Category label.
        if (horizontal) {
            const yMid = rect.top + slotSize * (i + 0.5);
            svg.appendChild(mkText(doc, rect.left - 6, yMid + 4, categories[i] ?? "", {
                fontSize: FONT_SIZE, anchor: "end", fill: catTickColor,
            }));
        } else {
            const xMid = rect.left + slotSize * (i + 0.5);
            svg.appendChild(mkText(doc, xMid, rect.bottom + 16, categories[i] ?? "", {
                fontSize: FONT_SIZE, anchor: "middle", fill: catTickColor,
            }));
        }

        if (stacked) {
            let cumulative = 0;
            // For percent-stacked we need the per-slot total up front.
            let slotTotal = 0;
            if (percentStacked) {
                for (const s of series) {
                    const v = finite(s.values[i]);
                    if (v == null) continue;
                    slotTotal += v;
                }
            }
            for (let si = 0; si < series.length; si++) {
                const raw = finite(series[si].values[i]);
                if (raw == null || raw === 0) continue;
                const value = percentStacked
                    ? (slotTotal === 0 ? 0 : raw / slotTotal)
                    : raw;
                // Per-data-point override by point index `i` wins over
                // the series colour; series colour wins over palette.
                const color = pointColor(series[si], i, si);
                appendBar(svg, doc, {
                    horizontal, rect, yMin, yMax, plotW, plotH,
                    slotIndex: i, slotSize, barIndex: 0,
                    barSize: innerGroup, groupPad,
                    start: cumulative, end: cumulative + value,
                    color,
                });
                cumulative += value;
            }
        } else {
            for (let si = 0; si < series.length; si++) {
                const value = finite(series[si].values[i]);
                if (value == null) continue;
                const color = pointColor(series[si], i, si);
                appendBar(svg, doc, {
                    horizontal, rect, yMin, yMax, plotW, plotH,
                    slotIndex: i, slotSize, barIndex: si,
                    barSize, groupPad,
                    start: 0, end: value,
                    color,
                });
            }
        }
    }
}

function appendBar(svg: SVGElement, doc: Document, p: {
    horizontal: boolean;
    rect: { left: number; right: number; top: number; bottom: number };
    yMin: number; yMax: number; plotW: number; plotH: number;
    slotIndex: number; slotSize: number; barIndex: number;
    barSize: number; groupPad: number;
    start: number; end: number; color: string;
}) {
    const scale = (v: number) => (v - p.yMin) / (p.yMax - p.yMin);
    const s0 = scale(Math.min(p.start, p.end));
    const s1 = scale(Math.max(p.start, p.end));

    if (p.horizontal) {
        const y = p.rect.top + p.slotSize * p.slotIndex + p.groupPad
            + p.barSize * p.barIndex;
        const x0 = p.rect.left + s0 * p.plotW;
        const x1 = p.rect.left + s1 * p.plotW;
        const rect = doc.createElementNS(SVG_NS, "rect");
        rect.setAttribute("x", fmt(Math.min(x0, x1)));
        rect.setAttribute("y", fmt(y));
        rect.setAttribute("width", fmt(Math.max(0, Math.abs(x1 - x0))));
        rect.setAttribute("height", fmt(Math.max(0, p.barSize)));
        rect.setAttribute("fill", p.color);
        svg.appendChild(rect);
    } else {
        const x = p.rect.left + p.slotSize * p.slotIndex + p.groupPad
            + p.barSize * p.barIndex;
        const y0 = p.rect.bottom - s0 * p.plotH;
        const y1 = p.rect.bottom - s1 * p.plotH;
        const rect = doc.createElementNS(SVG_NS, "rect");
        rect.setAttribute("x", fmt(x));
        rect.setAttribute("y", fmt(Math.min(y0, y1)));
        rect.setAttribute("width", fmt(Math.max(0, p.barSize)));
        rect.setAttribute("height", fmt(Math.max(0, Math.abs(y1 - y0))));
        rect.setAttribute("fill", p.color);
        svg.appendChild(rect);
    }
}

// -----------------------------------------------------------------
// Line
// -----------------------------------------------------------------

function renderLineChart(
    svg: SVGElement, doc: Document, model: ChartModel,
    rect: { left: number; right: number; top: number; bottom: number },
    axisColors: AxisColors,
) {
    const series = model.series.filter((s) => s.values.length > 0);
    if (series.length === 0) return;

    const categories = maxLengthCategories(series);
    const catCount = categories.length;
    if (catCount === 0) return;

    let valMin = Infinity;
    let valMax = -Infinity;
    for (const s of series) {
        for (const v of s.values) {
            const f = finite(v);
            if (f == null) continue;
            if (f < valMin) valMin = f;
            if (f > valMax) valMax = f;
        }
    }
    if (!Number.isFinite(valMin) || !Number.isFinite(valMax)) return;
    // Pad min side toward zero if all positive, for nicer framing.
    if (valMin > 0 && valMin / (valMax - valMin || 1) < 0.5) valMin = 0;
    if (valMax === valMin) valMax = valMin + 1;
    const { min: yMin, max: yMax, ticks } = niceScale(valMin, valMax, 5);

    const plotW = rect.right - rect.left;
    const plotH = rect.bottom - rect.top;

    const valLineColor = axisColors.valAxis.line;
    const catLineColor = axisColors.catAxis.line;
    const valTickColor = axisColors.valAxis.tickLabel;
    const catTickColor = axisColors.catAxis.tickLabel;
    const valGrid = axisColors.valAxis.gridline;

    svg.appendChild(mkLine(doc, rect.left, rect.top, rect.left, rect.bottom, valLineColor));
    svg.appendChild(mkLine(doc, rect.left, rect.bottom, rect.right, rect.bottom, catLineColor));

    for (const t of ticks) {
        const y = rect.bottom - ((t - yMin) / (yMax - yMin)) * plotH;
        if (valGrid && t !== yMin) {
            svg.appendChild(mkLine(doc, rect.left, y, rect.right, y, valGrid));
        }
        svg.appendChild(mkLine(doc, rect.left - 4, y, rect.left, y, valLineColor));
        svg.appendChild(mkText(doc, rect.left - 6, y + 4, formatTick(t), {
            fontSize: FONT_SIZE, anchor: "end", fill: valTickColor,
        }));
    }

    const xStep = catCount > 1 ? plotW / (catCount - 1) : 0;
    for (let i = 0; i < catCount; i++) {
        const x = rect.left + xStep * i;
        svg.appendChild(mkText(doc, x, rect.bottom + 16, categories[i] ?? "", {
            fontSize: FONT_SIZE, anchor: "middle", fill: catTickColor,
        }));
    }

    for (let si = 0; si < series.length; si++) {
        const s = series[si];
        // The line itself uses the series colour — a stroke with per-
        // segment colours would require rendering N sub-paths, which
        // isn't worth the complexity for v1.
        const lineColor = seriesColor(s, si);
        const entries: { x: number; y: number; pointIndex: number }[] = [];
        for (let i = 0; i < catCount; i++) {
            const v = finite(s.values[i]);
            if (v == null) continue;
            const x = catCount > 1 ? rect.left + xStep * i : rect.left + plotW / 2;
            const y = rect.bottom - ((v - yMin) / (yMax - yMin)) * plotH;
            entries.push({ x, y, pointIndex: i });
        }
        if (entries.length === 0) continue;
        const polyline = doc.createElementNS(SVG_NS, "polyline");
        polyline.setAttribute("points", entries.map((e) => `${fmt(e.x)},${fmt(e.y)}`).join(" "));
        polyline.setAttribute("fill", "none");
        polyline.setAttribute("stroke", lineColor);
        polyline.setAttribute("stroke-width", "2");
        svg.appendChild(polyline);
        // Data-point markers honour per-point overrides.
        for (const e of entries) {
            const circle = doc.createElementNS(SVG_NS, "circle");
            circle.setAttribute("cx", fmt(e.x));
            circle.setAttribute("cy", fmt(e.y));
            circle.setAttribute("r", "2.5");
            circle.setAttribute("fill", pointColor(s, e.pointIndex, si));
            svg.appendChild(circle);
        }
    }
}

// -----------------------------------------------------------------
// Pie
// -----------------------------------------------------------------

function renderPieChart(
    svg: SVGElement, doc: Document, model: ChartModel,
    rect: { left: number; right: number; top: number; bottom: number },
) {
    // For a pie chart, the first series provides the slice values and
    // categories provide the slice labels/legend entries.
    const series = model.series.find((s) => s.values.length > 0);
    if (!series) return;

    const values = series.values.map((v) => finite(v) ?? 0);
    const total = values.reduce((a, b) => a + (b > 0 ? b : 0), 0);
    if (total <= 0) return;

    const cx = (rect.left + rect.right) / 2;
    const cy = (rect.top + rect.bottom) / 2;
    const r = Math.max(0, Math.min(rect.right - rect.left, rect.bottom - rect.top) / 2 - 8);
    if (r <= 0) return;

    let startAngle = -Math.PI / 2;
    for (let i = 0; i < values.length; i++) {
        const v = values[i];
        if (v <= 0) continue;
        const angle = (v / total) * 2 * Math.PI;
        const endAngle = startAngle + angle;
        const path = pieSlicePath(cx, cy, r, startAngle, endAngle);
        const slice = doc.createElementNS(SVG_NS, "path");
        slice.setAttribute("d", path);
        // Per-slice precedence:
        //   1. <c:dPt idx="i"> override (the common case in real Word
        //      files — each slice gets its own colour).
        //   2. Series-level colour.
        //   3. Palette rotation by slice index.
        const override = series.dataPointOverrides?.get(i)?.color;
        const sanitisedOverride = override != null ? sanitizeCssColor(override) : null;
        const color = sanitisedOverride
            ?? sanitizeCssColor(series.color)
            ?? DEFAULT_PALETTE[i % DEFAULT_PALETTE.length];
        slice.setAttribute("fill", color);
        slice.setAttribute("stroke", "#fff");
        slice.setAttribute("stroke-width", "1");
        svg.appendChild(slice);
        startAngle = endAngle;
    }
}

function pieSlicePath(cx: number, cy: number, r: number, a0: number, a1: number): string {
    // Full-circle special case (avoids the degenerate arc that 0-length
    // endpoints produce).
    if (a1 - a0 >= 2 * Math.PI - 1e-9) {
        return `M ${fmt(cx - r)} ${fmt(cy)} A ${fmt(r)} ${fmt(r)} 0 1 0 ${fmt(cx + r)} ${fmt(cy)} A ${fmt(r)} ${fmt(r)} 0 1 0 ${fmt(cx - r)} ${fmt(cy)} Z`;
    }
    const x0 = cx + r * Math.cos(a0);
    const y0 = cy + r * Math.sin(a0);
    const x1 = cx + r * Math.cos(a1);
    const y1 = cy + r * Math.sin(a1);
    const large = a1 - a0 > Math.PI ? 1 : 0;
    return `M ${fmt(cx)} ${fmt(cy)} L ${fmt(x0)} ${fmt(y0)} A ${fmt(r)} ${fmt(r)} 0 ${large} 1 ${fmt(x1)} ${fmt(y1)} Z`;
}

// -----------------------------------------------------------------
// Legend
// -----------------------------------------------------------------

interface LegendLayout {
    height: number;
    nodes: (top: number) => Element[];
}

function layoutLegend(entries: ChartSeries[], maxWidth: number, doc: Document): LegendLayout {
    if (entries.length === 0) return { height: 0, nodes: () => [] };

    const swatchW = 10;
    const gap = 6;
    const entryGap = 16;
    // Estimate each entry's width from the title length. This is an
    // approximation — no getBBox because many consumers attach the SVG
    // to the DOM later. FONT_SIZE * 0.55 is a reasonable average for
    // most proportional fonts.
    const estWidth = (title: string) => swatchW + gap + Math.max(12, title.length * FONT_SIZE * 0.55);

    const rows: { entry: ChartSeries; seriesIndex: number; width: number }[][] = [[]];
    let rowWidth = 0;
    for (let i = 0; i < entries.length; i++) {
        const e = entries[i];
        const w = estWidth(e.title || `Series ${i + 1}`);
        const needed = rowWidth === 0 ? w : w + entryGap;
        if (rowWidth + needed > maxWidth && rows[rows.length - 1].length > 0) {
            rows.push([]);
            rowWidth = 0;
        }
        rows[rows.length - 1].push({ entry: e, seriesIndex: i, width: w });
        rowWidth += needed;
    }

    const height = rows.length * LEGEND_ROW_HEIGHT;

    return {
        height,
        nodes: (top: number) => {
            const out: Element[] = [];
            for (let r = 0; r < rows.length; r++) {
                const row = rows[r];
                const totalWidth = row.reduce((a, b, i) => a + b.width + (i === 0 ? 0 : entryGap), 0);
                let x = (VIEW_W - totalWidth) / 2;
                const y = top + r * LEGEND_ROW_HEIGHT + LEGEND_ROW_HEIGHT / 2;
                // Row key used by scheduleLegendOverflowAdjust to group
                // swatches + labels that should shift together. Safe
                // static string (no DOCX input).
                const rowKey = `r${r}`;
                for (const { entry, seriesIndex, width } of row) {
                    const color = seriesColor(entry, seriesIndex);
                    const swatch = doc.createElementNS(SVG_NS, "rect");
                    swatch.setAttribute("x", fmt(x));
                    swatch.setAttribute("y", fmt(y - swatchW / 2));
                    swatch.setAttribute("width", fmt(swatchW));
                    swatch.setAttribute("height", fmt(swatchW));
                    swatch.setAttribute("fill", color);
                    swatch.setAttribute("data-legend-row", rowKey);
                    out.push(swatch);
                    const label = mkText(doc, x + swatchW + gap, y + 4, entry.title || `Series ${seriesIndex + 1}`, {
                        fontSize: FONT_SIZE, anchor: "start", fill: "#333",
                    });
                    // Markers for the post-attach overflow pass. The
                    // estimated text width is the label portion only
                    // (the swatch+gap are added back during overflow
                    // comparisons — simpler to stash the original
                    // estimate than re-derive it).
                    label.setAttribute("data-legend-entry", "1");
                    label.setAttribute("data-legend-est-w", fmt(width - swatchW - gap));
                    label.setAttribute("data-legend-row", rowKey);
                    out.push(label);
                    x += width + entryGap;
                }
            }
            return out;
        },
    };
}

// -----------------------------------------------------------------
// Helpers
// -----------------------------------------------------------------

function seriesColor(s: ChartSeries, index: number): string {
    return sanitizeCssColor(s.color) ?? DEFAULT_PALETTE[index % DEFAULT_PALETTE.length];
}

// Per-data-point colour lookup with fallback chain:
//   1. <c:dPt idx="pointIndex"> fill override if present.
//   2. The series colour (<c:ser><c:spPr><a:solidFill>).
//   3. Palette rotation by series index (so different series stay
//      visually distinct even when nothing is explicitly coloured).
//
// Both override and series strings round-trip through sanitizeCssColor
// so a malformed override can never reach `fill` unfiltered.
function pointColor(s: ChartSeries, pointIndex: number, seriesIndex: number): string {
    const override = s.dataPointOverrides?.get(pointIndex)?.color;
    if (override != null) {
        const safe = sanitizeCssColor(override);
        if (safe) return safe;
    }
    return seriesColor(s, seriesIndex);
}

function finite(v: number): number | null {
    return Number.isFinite(v) ? v : null;
}

function maxLengthCategories(series: ChartSeries[]): string[] {
    let longest: string[] = [];
    let count = 0;
    for (const s of series) {
        const len = Math.max(s.values.length, s.categories.length);
        if (len > count) {
            count = len;
            longest = s.categories.slice(0, len);
        }
    }
    // Pad missing categories with empty strings (matching Word's
    // behaviour when categories are shorter than values).
    while (longest.length < count) longest.push("");
    return longest;
}

function fmt(n: number): string {
    if (!Number.isFinite(n)) return "0";
    return Math.abs(n) < 0.005 ? "0" : n.toFixed(2);
}

function formatTick(n: number): string {
    if (!Number.isFinite(n)) return "";
    if (Math.abs(n) >= 1000) return n.toFixed(0);
    if (Math.abs(n) >= 10) return n.toFixed(1);
    return n.toFixed(2).replace(/\.?0+$/, "") || "0";
}

// Returns a "nice" axis domain + tick list for [min, max]. Based on
// the classic Heckbert algorithm — no external dep.
function niceScale(min: number, max: number, targetTicks: number) {
    const range = niceNum(max - min, false);
    const step = niceNum(range / Math.max(1, targetTicks - 1), true);
    const niceMin = Math.floor(min / step) * step;
    const niceMax = Math.ceil(max / step) * step;
    const ticks: number[] = [];
    // Guard against pathological inputs blowing out the tick count.
    const maxTicks = 32;
    for (let v = niceMin, i = 0; v <= niceMax + step * 0.5 && i < maxTicks; v += step, i++) {
        ticks.push(Number(v.toFixed(10)));
    }
    return { min: niceMin, max: niceMax, ticks };
}

function niceNum(range: number, round: boolean): number {
    if (range <= 0) return 1;
    const exponent = Math.floor(Math.log10(range));
    const fraction = range / Math.pow(10, exponent);
    let niceFraction: number;
    if (round) {
        if (fraction < 1.5) niceFraction = 1;
        else if (fraction < 3) niceFraction = 2;
        else if (fraction < 7) niceFraction = 5;
        else niceFraction = 10;
    } else {
        if (fraction <= 1) niceFraction = 1;
        else if (fraction <= 2) niceFraction = 2;
        else if (fraction <= 5) niceFraction = 5;
        else niceFraction = 10;
    }
    return niceFraction * Math.pow(10, exponent);
}

function mkLine(doc: Document, x1: number, y1: number, x2: number, y2: number, stroke: string = "#888"): SVGElement {
    const el = doc.createElementNS(SVG_NS, "line") as SVGElement;
    el.setAttribute("x1", fmt(x1));
    el.setAttribute("y1", fmt(y1));
    el.setAttribute("x2", fmt(x2));
    el.setAttribute("y2", fmt(y2));
    el.setAttribute("stroke", stroke);
    el.setAttribute("stroke-width", "1");
    return el;
}

function mkText(
    doc: Document, x: number, y: number, text: string,
    opts: { fontSize?: number; anchor?: string; fill?: string; fontWeight?: string } = {},
): SVGElement {
    const el = doc.createElementNS(SVG_NS, "text") as SVGElement;
    el.setAttribute("x", fmt(x));
    el.setAttribute("y", fmt(y));
    if (opts.anchor) el.setAttribute("text-anchor", opts.anchor);
    el.setAttribute("font-size", String(opts.fontSize ?? FONT_SIZE));
    if (opts.fontWeight) el.setAttribute("font-weight", opts.fontWeight);
    if (opts.fill) el.setAttribute("fill", opts.fill);
    // DOCX-derived strings reach the DOM via textContent only.
    el.textContent = text;
    return el;
}

// -----------------------------------------------------------------
// Post-attach legend overflow adjustment
// -----------------------------------------------------------------
//
// layoutLegend estimates each entry's width from `title.length *
// FONT_SIZE * 0.55`. That estimate is right most of the time but
// wrong enough often enough (CJK / unusually wide fonts / very
// long titles) to justify a second pass once the SVG is in the DOM
// and `getBBox` returns real measurements.
//
// The adjustment is intentionally cheap:
//   - We only re-measure `<text>` elements that carry a `data-legend-entry`
//     marker (set by layoutLegend below).
//   - If any entry's measured width exceeds its estimated width by more
//     than TOLERANCE_PX, we assume the estimate was too generous and
//     shift each legend row's contents leftward to stop them spilling
//     past the right edge. We do not re-flow the whole chart.
//
// SSR-safe: `requestAnimationFrame` and `isConnected` guards prevent
// the hook from firing inside jsdom or non-browser environments.

export function scheduleLegendOverflowAdjust(svg: SVGElement) {
    if (typeof requestAnimationFrame !== "function") return;
    // isConnected is unreliable in jsdom-lite (undefined). The `!svg.isConnected`
    // check also rejects SVGs that are parented but not yet in the document.
    // Deferring to rAF gives the caller one tick to attach.
    requestAnimationFrame(() => {
        if (!svg.isConnected) return;
        try {
            adjustLegendIfNeeded(svg);
        } catch {
            // Never let measurement blow up the render.
        }
    });
}

function adjustLegendIfNeeded(svg: SVGElement) {
    const entries = Array.from(
        svg.querySelectorAll("text[data-legend-entry]"),
    ) as SVGTextElement[];
    if (entries.length === 0) return;

    let overflow = false;
    for (const entry of entries) {
        if (typeof entry.getBBox !== "function") return;
        const bbox = entry.getBBox();
        const estAttr = entry.getAttribute("data-legend-est-w");
        const est = estAttr ? parseFloat(estAttr) : NaN;
        if (!Number.isFinite(est)) continue;
        if (bbox.width > est + 2) {
            overflow = true;
            break;
        }
    }
    if (!overflow) return;

    // Cheap remediation: find each legend row (grouped by y-coord on
    // the `data-legend-row` attribute) and shift the whole row left
    // so its rightmost entry stays within VIEW_W - PADDING. If that
    // would push the row off the left edge, clip the overflow and
    // let the SVG's `overflow: hidden` crop it (still better than
    // extending past the plot boundary).
    const byRow = new Map<string, SVGGraphicsElement[]>();
    const all = Array.from(
        svg.querySelectorAll("[data-legend-row]"),
    ) as SVGGraphicsElement[];
    for (const node of all) {
        const key = node.getAttribute("data-legend-row") ?? "";
        const arr = byRow.get(key) ?? [];
        arr.push(node);
        byRow.set(key, arr);
    }

    for (const nodes of byRow.values()) {
        let rowRight = -Infinity;
        for (const node of nodes) {
            if (typeof node.getBBox !== "function") continue;
            const bbox = node.getBBox();
            const right = bbox.x + bbox.width;
            if (right > rowRight) rowRight = right;
        }
        if (!Number.isFinite(rowRight)) continue;
        const allowed = VIEW_W - PADDING;
        if (rowRight <= allowed) continue;
        const shift = Math.min(rowRight - allowed, VIEW_W);
        for (const node of nodes) {
            // Stack shifts on top of any existing transform. The
            // nodes we emit carry no transform at render time, so
            // this is effectively set-once.
            const prev = node.getAttribute("transform") ?? "";
            const next = `${prev} translate(${(-shift).toFixed(2)},0)`.trim();
            node.setAttribute("transform", next);
        }
    }
}

// -----------------------------------------------------------------
// ChartEx: sunburst and treemap
// -----------------------------------------------------------------
//
// ChartEx schema uses a multi-level category tree (<cx:strDim
// type="cat"><cx:lvl>) plus a flat value dim (<cx:numDim type="val">).
// chartex-part.ts does the data-model work; here we only take the
// parsed tree and lay it out in SVG.
//
// Security: same rules as the classic chart renderer — labels reach
// the DOM via textContent only, colours via sanitizeCssColor. SVG
// path `d` strings are built from hard-coded command letters plus
// numbers formatted through `fmt`. No DOCX content reaches `d`.

// Minimum angular sweep (radians) before we render a sunburst label.
// ~0.12 rad ≈ 7°, below which the text would collide with the arc
// boundary even at the outer ring. Small slices still render the arc,
// just without a label.
const SUNBURST_LABEL_THRESHOLD = 0.12;
// Minimum rectangle width/height (px) before we render a treemap label.
const TREEMAP_LABEL_MIN_W = 32;
const TREEMAP_LABEL_MIN_H = 14;

export function renderSunburst(model: ChartExTreeDataModel): SVGElement {
    const doc: Document = document;
    const svg = doc.createElementNS(SVG_NS, "svg") as SVGElement;
    svg.setAttribute("viewBox", `0 0 ${VIEW_W} ${VIEW_H}`);
    svg.setAttribute("role", "img");
    svg.setAttribute("preserveAspectRatio", "xMidYMid meet");
    svg.style.maxWidth = "100%";
    svg.style.height = "auto";
    svg.style.display = "block";

    const title = (model.title ?? "").trim();
    const titleBottom = title ? PADDING + TITLE_HEIGHT : PADDING;
    if (title) {
        svg.appendChild(mkText(doc, VIEW_W / 2, PADDING + TITLE_HEIGHT - 10, title, {
            fontSize: 14, fontWeight: "600", anchor: "middle",
        }));
    }

    const plotTop = titleBottom + 4;
    const plotBottom = VIEW_H - PADDING;
    const cx = VIEW_W / 2;
    const cy = (plotTop + plotBottom) / 2;
    // Leave an inner hole so nested rings aren't squashed onto a point.
    const outerR = Math.max(0, Math.min(VIEW_W - 2 * PADDING, plotBottom - plotTop) / 2 - 4);
    if (outerR <= 0 || model.root.value <= 0 || model.maxDepth === 0) {
        // Nothing to render — emit a neutral "empty" marker so the
        // caller's wrapper still has something inside.
        svg.appendChild(mkText(doc, VIEW_W / 2, VIEW_H / 2, title || "[sunburst]", {
            fontSize: 12, anchor: "middle", fill: "#888",
        }));
        return svg;
    }
    // Innermost ring gets a small radius so it reads as a hub; each
    // subsequent ring adds uniform thickness.
    const innerR = Math.min(outerR * 0.15, 24);
    const ringThickness = (outerR - innerR) / model.maxDepth;

    // Walk the tree, emitting an SVG <path> arc for each non-root node.
    // Each child's angular sweep is proportional to its value share
    // of the parent; siblings are placed consecutively.
    renderSunburstNode(
        svg, doc, model.root, cx, cy, innerR, ringThickness,
        -Math.PI / 2, 2 * Math.PI, 0,
    );

    return svg;
}

function renderSunburstNode(
    svg: SVGElement, doc: Document, node: ChartExTreeNode,
    cx: number, cy: number, innerR: number, ringThickness: number,
    startAngle: number, sweep: number, paletteIdx: number,
) {
    const total = node.value > 0 ? node.value : sumChildValue(node);
    if (total <= 0 || node.children.length === 0) return;

    let angle = startAngle;
    for (let i = 0; i < node.children.length; i++) {
        const child = node.children[i];
        // Leaf values that degenerated to 0 (negative / NaN source)
        // contribute nothing to the sweep; skip them silently.
        if (child.value <= 0) continue;

        const share = child.value / total;
        const slice = sweep * share;
        const r0 = innerR + ringThickness * child.level;
        const r1 = r0 + ringThickness;

        // Colour precedence: explicit tree-node colour (set by
        // parse-time dataPt override or propagated from an ancestor);
        // fall back to palette rotation seeded by the top-level index
        // so the visual grouping follows the root's children.
        const basePaletteIdx = child.level === 0 ? i : paletteIdx;
        const color = sanitizeCssColor(child.color)
            ?? DEFAULT_PALETTE[basePaletteIdx % DEFAULT_PALETTE.length];

        const path = doc.createElementNS(SVG_NS, "path");
        path.setAttribute("d", sunburstArcPath(cx, cy, r0, r1, angle, angle + slice));
        path.setAttribute("fill", color);
        path.setAttribute("stroke", "#fff");
        path.setAttribute("stroke-width", "1");
        svg.appendChild(path);

        // Label the slice when there's room. "Enough room" is an
        // approximation: the midpoint of the arc, rotated back so
        // text sits upright. Skip very small sweeps to stop labels
        // overlapping.
        if (slice >= SUNBURST_LABEL_THRESHOLD && child.label) {
            const midAngle = angle + slice / 2;
            const midR = (r0 + r1) / 2;
            const tx = cx + midR * Math.cos(midAngle);
            const ty = cy + midR * Math.sin(midAngle);
            // Baseline adjustment keeps the label centred on the mid-
            // radius rather than hanging below it.
            const text = mkText(doc, tx, ty + 4, child.label, {
                fontSize: Math.min(FONT_SIZE, Math.max(8, ringThickness * 0.45)),
                anchor: "middle",
                fill: "#fff",
            });
            svg.appendChild(text);
        }

        // Recurse into descendants — they occupy the outer rings within
        // this slice's angular range.
        renderSunburstNode(
            svg, doc, child, cx, cy, innerR, ringThickness,
            angle, slice, basePaletteIdx,
        );

        angle += slice;
    }
}

function sumChildValue(node: ChartExTreeNode): number {
    let total = 0;
    for (const c of node.children) total += c.value;
    return total;
}

// Builds a filled annular sector (pie-slice ring) from (r0, r1) ×
// (a0, a1). SVG path:
//   M outer-start  A outer-arc  L inner-end  A inner-arc  Z
// All tokens are command letters plus numbers — no DOCX string ever
// reaches the `d` attribute.
function sunburstArcPath(
    cx: number, cy: number, r0: number, r1: number,
    a0: number, a1: number,
): string {
    const large = a1 - a0 > Math.PI ? 1 : 0;
    const x0o = cx + r1 * Math.cos(a0);
    const y0o = cy + r1 * Math.sin(a0);
    const x1o = cx + r1 * Math.cos(a1);
    const y1o = cy + r1 * Math.sin(a1);
    const x0i = cx + r0 * Math.cos(a1);
    const y0i = cy + r0 * Math.sin(a1);
    const x1i = cx + r0 * Math.cos(a0);
    const y1i = cy + r0 * Math.sin(a0);
    return `M ${fmt(x0o)} ${fmt(y0o)}`
        + ` A ${fmt(r1)} ${fmt(r1)} 0 ${large} 1 ${fmt(x1o)} ${fmt(y1o)}`
        + ` L ${fmt(x0i)} ${fmt(y0i)}`
        + ` A ${fmt(r0)} ${fmt(r0)} 0 ${large} 0 ${fmt(x1i)} ${fmt(y1i)}`
        + ` Z`;
}

export function renderTreemap(model: ChartExTreeDataModel): SVGElement {
    const doc: Document = document;
    const svg = doc.createElementNS(SVG_NS, "svg") as SVGElement;
    svg.setAttribute("viewBox", `0 0 ${VIEW_W} ${VIEW_H}`);
    svg.setAttribute("role", "img");
    svg.setAttribute("preserveAspectRatio", "xMidYMid meet");
    svg.style.maxWidth = "100%";
    svg.style.height = "auto";
    svg.style.display = "block";

    const title = (model.title ?? "").trim();
    const titleBottom = title ? PADDING + TITLE_HEIGHT : PADDING;
    if (title) {
        svg.appendChild(mkText(doc, VIEW_W / 2, PADDING + TITLE_HEIGHT - 10, title, {
            fontSize: 14, fontWeight: "600", anchor: "middle",
        }));
    }

    const plotTop = titleBottom + 4;
    const plotBottom = VIEW_H - PADDING;
    const plot = {
        x: PADDING, y: plotTop,
        w: VIEW_W - 2 * PADDING, h: plotBottom - plotTop,
    };
    if (plot.w <= 0 || plot.h <= 0 || model.root.value <= 0) {
        svg.appendChild(mkText(doc, VIEW_W / 2, VIEW_H / 2, title || "[treemap]", {
            fontSize: 12, anchor: "middle", fill: "#888",
        }));
        return svg;
    }

    // Squarified layout (Bruls et al. 2000) — produces rectangles with
    // aspect ratios closer to 1:1 than the slice-and-dice alternative,
    // matching Excel / Word's own rendering more closely. The entry
    // point renderTreemap is unchanged; the swap is internal.
    layoutSquarifiedTree(model.root, plot.x, plot.y, plot.w, plot.h);
    renderTreemapNodes(svg, doc, model.root, 0);
    return svg;
}

// Each node owns a {x, y, w, h} rectangle computed during layout. We
// stash the rect on the node itself rather than a side map because
// the tree is single-use (we rebuild it per part). This keeps render
// traversal trivial.
interface LaidOutNode extends ChartExTreeNode {
    _x?: number; _y?: number; _w?: number; _h?: number;
}

// Recursively lay out `node`'s subtree. The parent's rect is written
// to the node, then we squarify its direct children across that rect
// and recurse into any non-leaf children with their computed rects.
function layoutSquarifiedTree(
    node: ChartExTreeNode, x: number, y: number, w: number, h: number,
) {
    const ln = node as LaidOutNode;
    ln._x = x; ln._y = y; ln._w = w; ln._h = h;

    if (node.children.length === 0) return;
    if (w <= 0 || h <= 0) {
        // Propagate an empty rect down so descendants don't inherit
        // stale values from a previous layout.
        for (const c of node.children) layoutSquarifiedTree(c, x, y, 0, 0);
        return;
    }

    const layout = squarifiedLayout(node.children, { x, y, width: w, height: h });
    for (const r of layout) {
        layoutSquarifiedTree(r.node, r.x, r.y, r.width, r.height);
    }
}

// -----------------------------------------------------------------
// Squarified treemap layout
// -----------------------------------------------------------------
//
// See Bruls et al., 'Squarified Treemaps' (2000).
//
// Given a rectangle and a set of child values, greedily pack children
// into rows along the shorter side. For each candidate child, compute
// the worst aspect ratio among the row's rectangles; commit the child
// to the row if the ratio improves, otherwise close the row, peel it
// off the parent, and start a new row on the remaining space.

export interface TreemapLayoutRect {
    node: ChartExTreeNode;
    x: number;
    y: number;
    width: number;
    height: number;
}

interface PlainRect { x: number; y: number; width: number; height: number; }

export function squarifiedLayout(
    children: ChartExTreeNode[],
    rect: PlainRect,
): TreemapLayoutRect[] {
    const out: TreemapLayoutRect[] = [];
    if (children.length === 0) return out;
    if (!(rect.width > 0) || !(rect.height > 0)) return out;

    // Single-child fast path: it fills the whole rect regardless of
    // value sign.
    if (children.length === 1) {
        out.push({
            node: children[0],
            x: rect.x, y: rect.y,
            width: rect.width, height: rect.height,
        });
        return out;
    }

    // Coerce values to finite non-negative numbers, preserving the
    // child reference. Sort descending — the algorithm assumes it.
    const items = children
        .map((node) => {
            const raw = parseFloat(node.value as unknown as string);
            const v = Number.isFinite(raw) && raw > 0 ? raw : 0;
            return { node, value: v };
        })
        .sort((a, b) => b.value - a.value);

    const total = items.reduce((s, i) => s + i.value, 0);
    if (total <= 0) {
        // Every child is zero / negative. Match Wave 7.2 semantics:
        // leaves still render with a label but no area. Emit zero-area
        // rects collapsed against the parent's top-left corner.
        for (const it of items) {
            out.push({
                node: it.node,
                x: rect.x, y: rect.y,
                width: 0, height: 0,
            });
        }
        return out;
    }

    // Scale value-space so the sum equals the rect's area. Keeping the
    // algorithm in scaled units avoids recomputing the scale factor
    // every step.
    const area = rect.width * rect.height;
    const scale = area / total;
    const scaled = items.map((i) => ({ node: i.node, area: i.value * scale }));

    squarifyInto(scaled, [], { ...rect }, out);
    return out;
}

interface ScaledItem { node: ChartExTreeNode; area: number; }

function squarifyInto(
    remaining: ScaledItem[],
    row: ScaledItem[],
    rect: PlainRect,
    out: TreemapLayoutRect[],
): void {
    // Iterate instead of recursing to dodge stack growth on pathological
    // inputs. Each loop either extends the current row or closes it.
    while (true) {
        if (remaining.length === 0) {
            if (row.length > 0) layoutRow(row, rect, out);
            return;
        }
        const w = Math.min(rect.width, rect.height);
        if (w <= 0) {
            // No space left; drop remaining as zero-area against the
            // current rect origin. Unreachable in practice because
            // layoutRow peels off a strip of positive width whenever
            // the row had positive area — defensive.
            for (const it of [...row, ...remaining]) {
                out.push({ node: it.node, x: rect.x, y: rect.y, width: 0, height: 0 });
            }
            return;
        }
        const head = remaining[0];
        const extended = row.length === 0
            ? [head]
            : [...row, head];
        // Zero-area items never worsen the row (numerator collapses);
        // cheaper to skip the ratio comparison and just keep going.
        if (row.length === 0 || worstRatio(extended, w) <= worstRatio(row, w)) {
            row = extended;
            remaining = remaining.slice(1);
        } else {
            rect = layoutRow(row, rect, out);
            row = [];
        }
    }
}

function worstRatio(row: ScaledItem[], w: number): number {
    let s = 0;
    let rmax = -Infinity;
    let rmin = Infinity;
    for (const r of row) {
        s += r.area;
        if (r.area > rmax) rmax = r.area;
        if (r.area < rmin) rmin = r.area;
    }
    if (s <= 0) return Infinity;
    const w2 = w * w;
    const s2 = s * s;
    // rmin can be 0 — when the row contains a zero-area item, its
    // aspect ratio is unbounded. Report Infinity so the row closes
    // rather than absorbing a sliver of nothing.
    const a = (w2 * rmax) / s2;
    const b = rmin > 0 ? s2 / (w2 * rmin) : Infinity;
    return Math.max(a, b);
}

// Place `row` along the shorter side of `rect`; return the remaining
// rect for subsequent rows.
function layoutRow(row: ScaledItem[], rect: PlainRect, out: TreemapLayoutRect[]): PlainRect {
    let sum = 0;
    for (const r of row) sum += r.area;
    if (sum <= 0) {
        // Row is all zero-area. Emit collapsed rects at the origin and
        // leave the parent rect untouched so the next row has the full
        // remaining space.
        for (const r of row) {
            out.push({ node: r.node, x: rect.x, y: rect.y, width: 0, height: 0 });
        }
        return rect;
    }
    const horizontal = rect.width >= rect.height;
    if (horizontal) {
        // Row is a vertical strip of width `sum / height` along the
        // left edge.
        const stripW = sum / rect.height;
        let cy = rect.y;
        for (const r of row) {
            const hh = rect.height * (r.area / sum);
            out.push({
                node: r.node,
                x: rect.x, y: cy,
                width: stripW, height: hh,
            });
            cy += hh;
        }
        return {
            x: rect.x + stripW, y: rect.y,
            width: Math.max(0, rect.width - stripW), height: rect.height,
        };
    } else {
        // Row is a horizontal strip of height `sum / width` along the
        // top edge.
        const stripH = sum / rect.width;
        let cx = rect.x;
        for (const r of row) {
            const ww = rect.width * (r.area / sum);
            out.push({
                node: r.node,
                x: cx, y: rect.y,
                width: ww, height: stripH,
            });
            cx += ww;
        }
        return {
            x: rect.x, y: rect.y + stripH,
            width: rect.width, height: Math.max(0, rect.height - stripH),
        };
    }
}

// Harness-facing helper. Lays out a node list into `rect` and returns
// the positioned rectangles directly, so jsdom tests can assert on
// positions without needing a chart fixture or DOM.
export function layoutTreemap(
    nodes: ChartExTreeNode[],
    rect: PlainRect,
): TreemapLayoutRect[] {
    return squarifiedLayout(nodes, rect);
}

function renderTreemapNodes(
    svg: SVGElement, doc: Document, node: ChartExTreeNode, paletteIdx: number,
) {
    // Walk depth-first; render each leaf as a <rect> with its label.
    // Intermediate nodes don't render themselves — only their leaves
    // do — but their children inherit the seeded palette index so
    // colours cluster visually.
    for (let i = 0; i < node.children.length; i++) {
        const child = node.children[i];
        const seedIdx = child.level === 0 ? i : paletteIdx;
        if (child.children.length === 0) {
            renderTreemapLeaf(svg, doc, child, seedIdx);
        } else {
            renderTreemapNodes(svg, doc, child, seedIdx);
        }
    }
}

function renderTreemapLeaf(
    svg: SVGElement, doc: Document, node: ChartExTreeNode, paletteIdx: number,
) {
    const ln = node as LaidOutNode;
    const x = ln._x ?? 0;
    const y = ln._y ?? 0;
    const w = ln._w ?? 0;
    const h = ln._h ?? 0;
    if (w <= 0 || h <= 0) return;

    const color = sanitizeCssColor(node.color)
        ?? DEFAULT_PALETTE[paletteIdx % DEFAULT_PALETTE.length];

    const rect = doc.createElementNS(SVG_NS, "rect");
    // Use higher precision here than the default two-decimal `fmt`:
    // squarified rows share edges and need to line up sub-pixel so we
    // don't leave visible gaps or overlaps between siblings.
    rect.setAttribute("x", fmt6(x));
    rect.setAttribute("y", fmt6(y));
    rect.setAttribute("width", fmt6(w));
    rect.setAttribute("height", fmt6(h));
    rect.setAttribute("fill", color);
    rect.setAttribute("stroke", "#fff");
    rect.setAttribute("stroke-width", "1");
    svg.appendChild(rect);

    // Label only when the rect is large enough to hold readable text.
    if (node.label && w >= TREEMAP_LABEL_MIN_W && h >= TREEMAP_LABEL_MIN_H) {
        const label = mkText(doc, x + 4, y + 14, node.label, {
            fontSize: Math.min(FONT_SIZE, Math.max(9, Math.floor(h / 3))),
            anchor: "start",
            fill: "#fff",
        });
        svg.appendChild(label);
    }
}

function fmt6(n: number): string {
    if (!Number.isFinite(n)) return "0";
    // Trim trailing zeros to keep the emitted SVG compact.
    return Number(n.toFixed(6)).toString();
}

// -----------------------------------------------------------------
// ChartEx: waterfall / funnel / histogram
// -----------------------------------------------------------------
//
// Flat chartEx kinds. All three share the same outer SVG shell
// (title + plot rect) as sunburst / treemap; only the inner layout
// differs. Factored into a single helper to keep the per-kind
// renderers focused on geometry.
//
// Security: every label reaches the DOM via textContent; every
// colour round-trips through sanitizeCssColor before hitting a
// `fill` attribute. Numerics are already parseFloat-filtered on the
// model so we only defend against the degenerate cases (empty data,
// all-zero totals, NaN after arithmetic).

// Waterfall palette. Positive contributions use the green; negative
// contributions use the red; subtotals and totals use the blue. The
// per-point <cx:dataPt> override wins over all three — matches the
// same precedence rule as the classic pie/bar renderer.
const WATERFALL_POSITIVE = "#548235";
const WATERFALL_NEGATIVE = "#C00000";
const WATERFALL_TOTAL = "#4472C4";

// Funnel bars get a gradient-free solid fill — we rotate through the
// default palette so adjacent trapezoids stay visually distinct.
// Users with <cx:dataPt> overrides obviously win.

// Histogram default bin count when neither binSize nor binCount is
// declared. Matches PowerPoint's "Automatic" fallback closely enough
// for a viewer (the canonical Scott / Sturges rules would need the
// sample standard deviation, which is more work than this pass
// warrants).
const HISTOGRAM_DEFAULT_BIN_COUNT = 10;

// Upper bound on computed bin count. Guards against pathological
// binSize values (e.g. 0.0001 against a [0, 10000] range) producing
// a DOM-exploding number of <rect>s.
const HISTOGRAM_MAX_BINS = 200;

function chartExShell(
    title: string,
    emptyLabel: string,
): {
    svg: SVGElement;
    doc: Document;
    plot: { x: number; y: number; w: number; h: number };
    emptyIf(cond: boolean): SVGElement | null;
} {
    const doc: Document = document;
    const svg = doc.createElementNS(SVG_NS, "svg") as SVGElement;
    svg.setAttribute("viewBox", `0 0 ${VIEW_W} ${VIEW_H}`);
    svg.setAttribute("role", "img");
    svg.setAttribute("preserveAspectRatio", "xMidYMid meet");
    svg.style.maxWidth = "100%";
    svg.style.height = "auto";
    svg.style.display = "block";

    const clean = (title ?? "").trim();
    const titleBottom = clean ? PADDING + TITLE_HEIGHT : PADDING;
    if (clean) {
        svg.appendChild(mkText(doc, VIEW_W / 2, PADDING + TITLE_HEIGHT - 10, clean, {
            fontSize: 14, fontWeight: "600", anchor: "middle",
        }));
    }

    const plotTop = titleBottom + 4;
    const plotBottom = VIEW_H - PADDING;
    const plot = {
        x: PADDING + AXIS_LABEL_WIDTH,
        y: plotTop,
        w: VIEW_W - 2 * PADDING - AXIS_LABEL_WIDTH,
        h: plotBottom - plotTop - AXIS_LABEL_HEIGHT,
    };

    return {
        svg, doc, plot,
        emptyIf(cond: boolean) {
            if (!cond) return null;
            svg.appendChild(mkText(doc, VIEW_W / 2, VIEW_H / 2, clean || emptyLabel, {
                fontSize: 12, anchor: "middle", fill: "#888",
            }));
            return svg;
        },
    };
}

export function renderWaterfall(model: ChartExWaterfallModel): SVGElement {
    const shell = chartExShell(model.title, "[waterfall]");
    const empty = shell.emptyIf(
        model.points.length === 0 || shell.plot.w <= 0 || shell.plot.h <= 0,
    );
    if (empty) return empty;
    const { svg, doc, plot } = shell;

    // Walk once to compute the min/max of the rect spans. Running
    // total tracks the cumulative sum up to (but not including) the
    // current point; each point contributes [before, after] where:
    //   - normal:   before = runningTotal; after = runningTotal + value
    //   - subtotal: before = 0;            after = runningTotal + value
    //                                      runningTotal jumps to after.
    //   - total:    before = 0;            after = runningTotal + value
    //                                      runningTotal does NOT change.
    interface Span { before: number; after: number; }
    const spans: Span[] = [];
    let running = 0;
    for (const p of model.points) {
        let before: number;
        let after: number;
        if (p.type === "normal") {
            before = running;
            after = running + p.value;
            running = after;
        } else if (p.type === "subtotal") {
            before = 0;
            after = running + p.value;
            running = after;
        } else {
            before = 0;
            after = running + p.value;
            // Total doesn't advance the running sum — it's a cap,
            // visualised separately from the running contributions.
        }
        spans.push({ before, after });
    }

    let valMin = 0;
    let valMax = 0;
    for (const s of spans) {
        if (s.before < valMin) valMin = s.before;
        if (s.after < valMin) valMin = s.after;
        if (s.before > valMax) valMax = s.before;
        if (s.after > valMax) valMax = s.after;
    }
    if (valMax === valMin) valMax = valMin + 1;
    const { min: yMin, max: yMax, ticks } = niceScale(valMin, valMax, 5);

    // Axis chrome. Default greys — waterfall chartEx has no axis
    // style metadata in the model, unlike classic barChart.
    svg.appendChild(mkLine(doc, plot.x, plot.y, plot.x, plot.y + plot.h, DEFAULT_AXIS_LINE));
    svg.appendChild(mkLine(doc, plot.x, plot.y + plot.h, plot.x + plot.w, plot.y + plot.h, DEFAULT_AXIS_LINE));

    for (const t of ticks) {
        const y = plot.y + plot.h - ((t - yMin) / (yMax - yMin)) * plot.h;
        svg.appendChild(mkLine(doc, plot.x - 4, y, plot.x, y, DEFAULT_AXIS_LINE));
        svg.appendChild(mkText(doc, plot.x - 6, y + 4, formatTick(t), {
            fontSize: FONT_SIZE, anchor: "end", fill: DEFAULT_TICK_LABEL,
        }));
    }

    const n = model.points.length;
    const slotW = plot.w / n;
    const barPad = Math.min(slotW * 0.15, 8);
    const barW = Math.max(1, slotW - 2 * barPad);

    for (let i = 0; i < n; i++) {
        const p = model.points[i];
        const span = spans[i];
        const xMid = plot.x + slotW * (i + 0.5);

        // Category label.
        svg.appendChild(mkText(doc, xMid, plot.y + plot.h + 16, p.label, {
            fontSize: FONT_SIZE, anchor: "middle", fill: DEFAULT_TICK_LABEL,
        }));

        const scaled = (v: number) => plot.y + plot.h
            - ((v - yMin) / (yMax - yMin)) * plot.h;
        const y0 = scaled(Math.max(span.before, span.after));
        const y1 = scaled(Math.min(span.before, span.after));
        const barH = Math.max(0, y1 - y0);

        // Colour precedence: explicit per-point override from <cx:dataPt>
        // wins over the type-based default.
        const typeColor =
            p.type === "normal"
                ? (p.value >= 0 ? WATERFALL_POSITIVE : WATERFALL_NEGATIVE)
                : WATERFALL_TOTAL;
        const color = sanitizeCssColor(p.color) ?? typeColor;

        const rect = doc.createElementNS(SVG_NS, "rect");
        rect.setAttribute("x", fmt(xMid - barW / 2));
        rect.setAttribute("y", fmt(y0));
        rect.setAttribute("width", fmt(barW));
        rect.setAttribute("height", fmt(barH));
        rect.setAttribute("fill", color);
        svg.appendChild(rect);
    }

    return svg;
}

export function renderFunnel(model: ChartExFunnelModel): SVGElement {
    const shell = chartExShell(model.title, "[funnel]");
    // Prepare working arrays so we can reserve label space before the
    // "empty" check decides we have nothing to render.
    const n = model.points.length;
    let maxVal = 0;
    for (const p of model.points) {
        if (p.value > maxVal) maxVal = p.value;
    }

    const empty = shell.emptyIf(
        n === 0 || maxVal <= 0 || shell.plot.w <= 0 || shell.plot.h <= 0,
    );
    if (empty) return empty;
    const { svg, doc, plot } = shell;

    // Reserve the right-hand third of the plot for labels so the
    // widest trapezoid doesn't collide with its own "category: value"
    // annotation.
    const labelReserve = Math.min(160, plot.w * 0.33);
    const funnelW = Math.max(1, plot.w - labelReserve);
    const cx = plot.x + funnelW / 2;

    // Evenly spaced vertical stack — one trapezoid per point.
    const bandH = plot.h / n;
    const vertPad = Math.min(bandH * 0.1, 6);
    const inner = bandH - 2 * vertPad;

    for (let i = 0; i < n; i++) {
        const p = model.points[i];
        const nextVal = i + 1 < n ? model.points[i + 1].value : p.value;

        // Top / bottom widths proportional to this + next value.
        // Negative / non-finite values are clamped to zero by the parser.
        const topHalf = (p.value / maxVal) * funnelW / 2;
        const botHalf = (Math.min(nextVal, p.value) / maxVal) * funnelW / 2;

        const y0 = plot.y + bandH * i + vertPad;
        const y1 = y0 + inner;

        const x0t = cx - topHalf;
        const x1t = cx + topHalf;
        const x0b = cx - botHalf;
        const x1b = cx + botHalf;

        const color = sanitizeCssColor(p.color)
            ?? DEFAULT_PALETTE[i % DEFAULT_PALETTE.length];

        const poly = doc.createElementNS(SVG_NS, "polygon");
        const points = [
            `${fmt(x0t)},${fmt(y0)}`,
            `${fmt(x1t)},${fmt(y0)}`,
            `${fmt(x1b)},${fmt(y1)}`,
            `${fmt(x0b)},${fmt(y1)}`,
        ].join(" ");
        poly.setAttribute("points", points);
        poly.setAttribute("fill", color);
        poly.setAttribute("stroke", "#fff");
        poly.setAttribute("stroke-width", "1");
        svg.appendChild(poly);

        // Right-hand label: "Label: value". Label and value both come
        // from the model (textContent-safe).
        const labelX = plot.x + funnelW + 8;
        const labelY = (y0 + y1) / 2 + 4;
        const labelText = p.label
            ? `${p.label}: ${formatTick(p.value)}`
            : formatTick(p.value);
        svg.appendChild(mkText(doc, labelX, labelY, labelText, {
            fontSize: FONT_SIZE, anchor: "start", fill: "#333",
        }));
    }

    return svg;
}

export function renderHistogram(model: ChartExHistogramModel): SVGElement {
    const shell = chartExShell(model.title, "[histogram]");

    const values = model.values;
    const n = values.length;

    let minV = Infinity;
    let maxV = -Infinity;
    for (const v of values) {
        if (v < minV) minV = v;
        if (v > maxV) maxV = v;
    }
    const haveRange = n > 0 && Number.isFinite(minV) && Number.isFinite(maxV);

    const empty = shell.emptyIf(
        !haveRange || shell.plot.w <= 0 || shell.plot.h <= 0,
    );
    if (empty) return empty;
    const { svg, doc, plot } = shell;

    // Binning. Precedence mirrors PowerPoint's dialog:
    //   1. binSize (explicit bucket width) — walk from underflow
    //      (or min) to overflow (or max) in steps of binSize.
    //   2. binCount — divide the range into `binCount` equal-width
    //      buckets.
    //   3. Default — HISTOGRAM_DEFAULT_BIN_COUNT equal-width buckets.
    //
    // Edge cases:
    //   - underflow: any value < underflow lands in a single leading
    //     bin labelled "<= underflow".
    //   - overflow: any value > overflow lands in a single trailing
    //     bin labelled "> overflow".
    //   - min === max: a single bin spanning [min, min+1] keeps the
    //     renderer from dividing by zero.
    const { underflow, overflow } = model.binning;
    const useUnderflow = underflow != null && underflow > minV;
    const useOverflow = overflow != null && overflow < maxV;
    const rangeLo = useUnderflow ? underflow! : minV;
    let rangeHi = useOverflow ? overflow! : maxV;
    if (rangeHi <= rangeLo) rangeHi = rangeLo + 1;

    let binSize = model.binning.binSize;
    let binCount = model.binning.binCount;
    if (binSize == null) {
        const count = binCount != null
            ? binCount
            : Math.max(1, Math.min(HISTOGRAM_MAX_BINS, HISTOGRAM_DEFAULT_BIN_COUNT));
        binSize = (rangeHi - rangeLo) / count;
    }
    if (!(binSize > 0)) binSize = rangeHi - rangeLo;
    // Derive binCount from binSize (in case both were given and binSize
    // wins; or when binSize came from the default path).
    binCount = Math.max(1, Math.min(
        HISTOGRAM_MAX_BINS,
        Math.ceil((rangeHi - rangeLo) / binSize),
    ));
    // Tighten rangeHi to the last bin's right edge for label consistency.
    rangeHi = rangeLo + binSize * binCount;

    interface Bin { lo: number; hi: number; count: number; label: string; }
    const bins: Bin[] = [];
    if (useUnderflow) {
        bins.push({
            lo: -Infinity, hi: rangeLo, count: 0,
            label: `<= ${formatTick(rangeLo)}`,
        });
    }
    for (let i = 0; i < binCount; i++) {
        const lo = rangeLo + binSize * i;
        const hi = lo + binSize;
        bins.push({
            lo, hi, count: 0,
            label: `${formatTick(lo)}-${formatTick(hi)}`,
        });
    }
    if (useOverflow) {
        bins.push({
            lo: rangeHi, hi: Infinity, count: 0,
            label: `> ${formatTick(rangeHi)}`,
        });
    }

    // Bin each value. Right edge is exclusive except for the final
    // "normal" bin, which is inclusive so the maximum value still lands.
    for (const v of values) {
        if (useUnderflow && v < rangeLo) {
            bins[0].count++;
            continue;
        }
        if (useOverflow && v > rangeHi) {
            bins[bins.length - 1].count++;
            continue;
        }
        let idx = Math.floor((v - rangeLo) / binSize);
        if (idx < 0) idx = 0;
        if (idx >= binCount) idx = binCount - 1;
        const offset = useUnderflow ? 1 : 0;
        bins[idx + offset].count++;
    }

    // Value axis: bar counts.
    let maxCount = 0;
    for (const b of bins) {
        if (b.count > maxCount) maxCount = b.count;
    }
    if (maxCount === 0) maxCount = 1;
    const { min: yMin, max: yMax, ticks } = niceScale(0, maxCount, 5);

    svg.appendChild(mkLine(doc, plot.x, plot.y, plot.x, plot.y + plot.h, DEFAULT_AXIS_LINE));
    svg.appendChild(mkLine(doc, plot.x, plot.y + plot.h, plot.x + plot.w, plot.y + plot.h, DEFAULT_AXIS_LINE));

    for (const t of ticks) {
        const y = plot.y + plot.h - ((t - yMin) / (yMax - yMin)) * plot.h;
        svg.appendChild(mkLine(doc, plot.x - 4, y, plot.x, y, DEFAULT_AXIS_LINE));
        svg.appendChild(mkText(doc, plot.x - 6, y + 4, formatTick(t), {
            fontSize: FONT_SIZE, anchor: "end", fill: DEFAULT_TICK_LABEL,
        }));
    }

    const slotW = plot.w / bins.length;
    const barPad = Math.min(slotW * 0.1, 4);
    const barW = Math.max(1, slotW - 2 * barPad);

    const baseColor = sanitizeCssColor(model.seriesColor) ?? DEFAULT_PALETTE[0];

    for (let i = 0; i < bins.length; i++) {
        const b = bins[i];
        const xLeft = plot.x + slotW * i + barPad;
        const y0 = plot.y + plot.h - ((b.count - yMin) / (yMax - yMin)) * plot.h;
        const y1 = plot.y + plot.h;

        const override = model.dataPointOverrides.get(i);
        const color = (override ? sanitizeCssColor(override) : null) ?? baseColor;

        const rect = doc.createElementNS(SVG_NS, "rect");
        rect.setAttribute("x", fmt(xLeft));
        rect.setAttribute("y", fmt(y0));
        rect.setAttribute("width", fmt(barW));
        rect.setAttribute("height", fmt(Math.max(0, y1 - y0)));
        rect.setAttribute("fill", color);
        svg.appendChild(rect);

        svg.appendChild(mkText(doc, xLeft + barW / 2, plot.y + plot.h + 16, b.label, {
            fontSize: FONT_SIZE, anchor: "middle", fill: DEFAULT_TICK_LABEL,
        }));
    }

    return svg;
}
