// In-memory representation of a single <c:chartSpace> part, parsed from
// /word/charts/chart*.xml. Only the fields the SVG renderer actually
// consumes live here — formulas (<c:f>), numbering formats, themed
// gradients, axis tick details, etc. are intentionally dropped on the
// floor for v1.
//
// Security: every string on this model may originate in an untrusted
// DOCX. Series titles and category labels must reach the DOM via
// `textContent` only; colours must pass through `sanitizeCssColor`
// before becoming an SVG attribute value.

export type ChartKind =
    | "bar"       // c:barChart with barDir=bar (horizontal bars)
    | "column"    // c:barChart with barDir=col (vertical bars — Word's "column")
    | "line"      // c:lineChart
    | "pie"       // c:pieChart / c:doughnutChart (treated the same for v1)
    | "unknown";

// Per-data-point override parsed from <c:dPt>. Currently only the
// fill colour is extracted; other overrides (explode, border, marker
// shape) would be added here. Map key is the 0-based point index
// (<c:idx val="N"/>).
export interface ChartDataPointOverride {
    // Sanitised fill colour or null when the <c:dPt> had no usable
    // <a:solidFill>. Parsed from <c:spPr/a:solidFill/a:srgbClr> or
    // <a:schemeClr> (resolved via the theme palette).
    color?: string | null;
}

export interface ChartSeries {
    // <c:tx/c:strRef/c:strCache/c:pt/c:v> or the inline <c:tx/c:v>. Used
    // for the legend entry. Never interpolated into HTML/CSS.
    title: string;
    // Sanitised colour (`#hex` / `rgb(...)` / `null`). Null means "fall
    // back to palette". Parsed from <c:spPr/a:solidFill/a:srgbClr> or
    // <a:schemeClr>.
    color: string | null;
    // Finite numeric values from <c:val/c:numRef/c:numCache/c:pt/c:v>.
    // Non-finite entries (empty, "#N/A", huge) are filtered out.
    values: number[];
    // Matching category labels from <c:cat> — same length as `values`
    // when parsing succeeds, otherwise padded/truncated by the renderer.
    categories: string[];
    // Per-data-point overrides from <c:dPt>. Empty map when the
    // series didn't declare any. Indexed by point index, not by
    // insertion order — see `parseSeries` in chart-part.ts.
    dataPointOverrides: Map<number, ChartDataPointOverride>;
}

// Axis chrome colour reference. Carries either a sanitised literal
// (`"#4472C4"`), a scheme-slot placeholder to resolve against the
// document's theme palette at render time (`"schemeClr:accent1"`), or
// null when the DOCX had no usable colour and the renderer should
// fall back to its hard-coded defaults.
//
// Keeping axes as a string-placeholder union (rather than a full
// ColourRef) lets chart parsing stay palette-independent: the theme
// part may load before or after a chart part, and axes don't need
// modifier support for v1.
export type ChartAxisColorRef =
    | { kind: "literal"; color: string }
    | { kind: "scheme"; slot: string }
    | null;

// Axis chrome colour triplet. Each field is the reference collected
// at parse time; the renderer resolves each one against the active
// theme palette (passed via RenderChartOptions.themePalette).
//
// Parsed from <c:catAx> / <c:valAx>:
//   - `line`      ← <c:spPr><a:ln><a:solidFill> (the axis line stroke)
//   - `tickLabel` ← <c:txPr>/<a:p>/<a:pPr>/<a:defRPr><a:solidFill>
//   - `gridline`  ← <c:majorGridlines><c:spPr><a:ln><a:solidFill>
export interface ChartAxisStyle {
    line: ChartAxisColorRef;
    tickLabel: ChartAxisColorRef;
    gridline: ChartAxisColorRef;
}

export interface ChartModel {
    // Unique key (the chart part's path without extension). Currently
    // unused by the renderer but handy for debugging.
    key: string;
    // <c:title> text, if any.
    title: string;
    // Whether <c:legend> was present. When false we skip legend layout.
    showLegend: boolean;
    // Normalised chart kind. The renderer switches on this.
    kind: ChartKind;
    // Grouping for bar/column charts: "clustered" | "stacked" |
    // "percentStacked" | "standard". v1 only special-cases "stacked" and
    // "percentStacked"; anything else renders as clustered.
    grouping: string;
    // Ordered list of series. Pie charts use the first series only.
    series: ChartSeries[];
    // Axis chrome colours parsed from <c:catAx> and <c:valAx>. When a
    // chart has neither axis (pie / doughnut), both entries are the
    // all-null default and the renderer falls back to built-in greys.
    catAxis: ChartAxisStyle;
    valAxis: ChartAxisStyle;
}

// The allowlisted set of chartEx root-element local names. Anything
// outside this list is coerced to `"unknown"` before reaching the
// data-chart-kind attribute in the DOM (see ChartExPlaceholder).
export type ChartExKind =
    | "sunburst"
    | "waterfall"
    | "funnel"
    | "treemap"
    | "histogram"
    | "pareto"
    | "box_whisker"
    | "unknown";

// Placeholder emitted for `/word/charts/chartEx*.xml` parts whose kind
// we don't yet render as a real SVG chart (waterfall / funnel /
// histogram / pareto / box_whisker). The renderer produces a labelled
// `<div class="docx-chartex-placeholder">` so the reader sees
// something where a chart would otherwise be missing.
export interface ChartExPlaceholder {
    // Discriminator separating the placeholder from parsed data models.
    shape: "placeholder";
    // Unique key (the chartEx part's path without extension).
    key: string;
    // Extracted <cx:title> text, or "" when the chartEx had no title.
    // Always reaches the DOM via textContent.
    title: string;
    // Normalised chartEx type — allowlisted against ChartExKind before
    // the renderer writes it into the `data-chart-kind` attribute.
    kind: ChartExKind;
}

// A node in the hierarchical category tree used by both sunburst and
// treemap. Built from <cx:strDim type="cat"><cx:lvl> multi-level
// category data: each `<cx:lvl>` describes one depth of the tree and
// each `<cx:pt>` at level N declares its parent's idx at level N-1
// (via the `parent` / legacy `parentIdx` attribute).
//
// `value` on a leaf is the raw <cx:numDim type="val"> entry for that
// point. On an intermediate node `value` is the sum of its
// descendants' leaf values — computed by `buildCategoryTree`.
//
// `color` is the sanitised fill colour from the matching <cx:dataPt>
// override, or null when the series didn't declare one (renderer
// falls back to the palette by leaf index).
//
// Security: `label` is attacker-controlled and reaches the DOM only
// via textContent. `color` has already passed sanitizeCssColor.
export interface ChartExTreeNode {
    label: string;
    value: number;
    color: string | null;
    children: ChartExTreeNode[];
    // 0-based depth. Root placeholder is at level -1 and never renders;
    // the first real level is 0 (innermost ring for sunburst).
    level: number;
    // Original leaf index within the flattened <cx:numDim> (used to
    // look up per-point overrides and to rotate the palette). -1 on
    // intermediate nodes.
    leafIndex: number;
}

// Parsed data model for hierarchical chartEx kinds (sunburst / treemap).
// Both share the same tree shape; only the layout differs at render time.
export interface ChartExTreeDataModel {
    shape: "data";
    key: string;
    title: string;
    // Restricted to the kinds that use the tree shape.
    kind: "sunburst" | "treemap";
    // Root of the parsed category tree. Children are the level-0
    // nodes. The root itself is a synthetic container with label "".
    root: ChartExTreeNode;
    // Maximum depth across the tree (root = 0, first real level = 1,
    // etc.). Used by the sunburst renderer to size rings.
    maxDepth: number;
}

// A single entry in a waterfall chart. `type` is one of:
//   - "normal"    — a positive or negative contribution. The bar floats
//                   from the running total before this value to the
//                   running total after.
//   - "subtotal"  — the bar is anchored to the 0-baseline. Downstream
//                   contributions resume from this subtotal as their
//                   new baseline.
//   - "total"     — a final bar anchored to 0 and not contributing to
//                   the running sum (matches Word's rendering).
//
// Security: `label` is an attacker-controlled DOCX string and reaches
// the DOM via textContent only. `value` is parseFloat-filtered. `type`
// is allowlisted to one of the three literals.
export interface WaterfallPoint {
    label: string;
    value: number;
    type: "normal" | "subtotal" | "total";
    // Per-data-point override from <cx:dataPt idx>. Already sanitised.
    color: string | null;
}

export interface ChartExWaterfallModel {
    shape: "data";
    key: string;
    title: string;
    kind: "waterfall";
    points: WaterfallPoint[];
}

// A simple ordered list of {label, value} for funnel charts.
// `label` via textContent; `value` finite-checked; `color` sanitised.
export interface FunnelPoint {
    label: string;
    value: number;
    color: string | null;
}

export interface ChartExFunnelModel {
    shape: "data";
    key: string;
    title: string;
    kind: "funnel";
    points: FunnelPoint[];
}

// Histogram binning parameters from <cx:binning>. At least one of
// `binSize` or `binCount` is set; the renderer prefers `binSize` when
// both are present (matches PowerPoint's fallback order).
export interface HistogramBinning {
    binSize: number | null;
    binCount: number | null;
    underflow: number | null;
    overflow: number | null;
}

export interface ChartExHistogramModel {
    shape: "data";
    key: string;
    title: string;
    kind: "histogram";
    // Raw (unbinned) values from <cx:numDim type="val">. The renderer
    // bins them at layout time using `binning`.
    values: number[];
    binning: HistogramBinning;
    // Series-level colour (pre-sanitised) and a map of per-bin
    // overrides (by bin index). Populated when <cx:dataPt> entries
    // were present.
    seriesColor: string | null;
    dataPointOverrides: Map<number, string>;
}

// Parsed data model union — one variant per kind of chartEx we
// render as real SVG. Kept as a discriminated union on `kind` so the
// renderer dispatch is type-safe.
export type ChartExDataModel =
    | ChartExTreeDataModel
    | ChartExWaterfallModel
    | ChartExFunnelModel
    | ChartExHistogramModel;

// Union consumed by renderChartEx in html-renderer.ts.
export type ChartExModel = ChartExPlaceholder | ChartExDataModel;
