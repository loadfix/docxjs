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

// Placeholder emitted for `/word/charts/chartEx*.xml` parts. We don't
// attempt to render these as real charts in v1; the renderer instead
// produces a labelled `<div class="docx-chartex-placeholder">` so
// the reader sees something where a chart would otherwise be missing.
export interface ChartExPlaceholder {
    // Unique key (the chartEx part's path without extension).
    key: string;
    // Extracted <cx:title> text, or "" when the chartEx had no title.
    // Always reaches the DOM via textContent.
    title: string;
    // Normalised chartEx type — allowlisted against ChartExKind before
    // the renderer writes it into the `data-chart-kind` attribute.
    kind: ChartExKind;
}
