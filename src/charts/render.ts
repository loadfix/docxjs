import { sanitizeCssColor } from "../utils";
import { ChartModel, ChartSeries } from "./model";

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

export function renderChart(model: ChartModel): SVGElement {
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
            });
            break;
        case "line":
            renderLineChart(svg, doc, model, {
                left: PADDING + AXIS_LABEL_WIDTH,
                right: VIEW_W - PADDING,
                top: plotTop,
                bottom: plotBottom - AXIS_LABEL_HEIGHT,
            });
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

    // Axes.
    svg.appendChild(mkLine(doc, rect.left, rect.top, rect.left, rect.bottom));
    svg.appendChild(mkLine(doc, rect.left, rect.bottom, rect.right, rect.bottom));

    // Tick labels for the value axis.
    for (const t of ticks) {
        if (horizontal) {
            const x = rect.left + ((t - yMin) / (yMax - yMin)) * plotW;
            svg.appendChild(mkLine(doc, x, rect.bottom, x, rect.bottom + 4));
            svg.appendChild(mkText(doc, x, rect.bottom + 16, formatTick(t), {
                fontSize: FONT_SIZE, anchor: "middle", fill: "#555",
            }));
        } else {
            const y = rect.bottom - ((t - yMin) / (yMax - yMin)) * plotH;
            svg.appendChild(mkLine(doc, rect.left - 4, y, rect.left, y));
            svg.appendChild(mkText(doc, rect.left - 6, y + 4, formatTick(t), {
                fontSize: FONT_SIZE, anchor: "end", fill: "#555",
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
                fontSize: FONT_SIZE, anchor: "end", fill: "#555",
            }));
        } else {
            const xMid = rect.left + slotSize * (i + 0.5);
            svg.appendChild(mkText(doc, xMid, rect.bottom + 16, categories[i] ?? "", {
                fontSize: FONT_SIZE, anchor: "middle", fill: "#555",
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

    svg.appendChild(mkLine(doc, rect.left, rect.top, rect.left, rect.bottom));
    svg.appendChild(mkLine(doc, rect.left, rect.bottom, rect.right, rect.bottom));

    for (const t of ticks) {
        const y = rect.bottom - ((t - yMin) / (yMax - yMin)) * plotH;
        svg.appendChild(mkLine(doc, rect.left - 4, y, rect.left, y));
        svg.appendChild(mkText(doc, rect.left - 6, y + 4, formatTick(t), {
            fontSize: FONT_SIZE, anchor: "end", fill: "#555",
        }));
    }

    const xStep = catCount > 1 ? plotW / (catCount - 1) : 0;
    for (let i = 0; i < catCount; i++) {
        const x = rect.left + xStep * i;
        svg.appendChild(mkText(doc, x, rect.bottom + 16, categories[i] ?? "", {
            fontSize: FONT_SIZE, anchor: "middle", fill: "#555",
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

function mkLine(doc: Document, x1: number, y1: number, x2: number, y2: number): SVGElement {
    const el = doc.createElementNS(SVG_NS, "line") as SVGElement;
    el.setAttribute("x1", fmt(x1));
    el.setAttribute("y1", fmt(y1));
    el.setAttribute("x2", fmt(x2));
    el.setAttribute("y2", fmt(y2));
    el.setAttribute("stroke", "#888");
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
