import { Part } from "../common/part";
import { OpenXmlPackage } from "../common/open-xml-package";
import xml from "../parser/xml-parser";
import { sanitizeCssColor } from "../utils";
import { ChartDataPointOverride, ChartKind, ChartModel, ChartSeries } from "./model";
import { resolveSchemeColor } from "./theme";

// Parses a `/word/charts/chart*.xml` part into a `ChartModel`. The
// renderer (`src/charts/render.ts`) owns layout; this file only knows
// the XML schema.
//
// Out of scope for v1:
// - <c:f> formulas (we render whatever `c:*Cache` captured, which is
//   what Word itself displays when it opens the file without Excel)
// - <c:chartEx> modern chart types (sunburst/waterfall/funnel/treemap)
// - themed / gradient / pattern fills — we accept <a:solidFill> only
// - value-axis formatting, log scale, reversed axes
// - combo charts (multiple plot-area chart elements with different kinds)
//
// Everything parsed here is attacker-controlled. `ChartSeries.title`
// and `ChartSeries.categories` reach the DOM via `textContent` only;
// `ChartSeries.color` always round-trips through `sanitizeCssColor` so
// an attacker can't inject a CSS/SVG attribute payload.

export class ChartPart extends Part {
    chart: ChartModel;

    constructor(pkg: OpenXmlPackage, path: string) {
        super(pkg, path);
    }

    protected parseXml(root: Element) {
        this.chart = parseChartSpace(root, this.path);
    }
}

function parseChartSpace(chartSpace: Element, path: string): ChartModel {
    const key = deriveKey(path);
    const chart = xml.element(chartSpace, "chart");
    if (!chart) {
        return {
            key, title: "", showLegend: false, kind: "unknown",
            grouping: "clustered", series: [],
        };
    }

    const titleEl = xml.element(chart, "title");
    const legendEl = xml.element(chart, "legend");
    const plotArea = xml.element(chart, "plotArea");

    const title = titleEl ? extractRichText(titleEl) : "";
    const showLegend = legendEl != null;

    // The first recognised chart-kind element inside <c:plotArea> wins.
    // Combo charts would expose multiple — we intentionally take only
    // the first (documented as a scope cut).
    let kind: ChartKind = "unknown";
    let grouping = "clustered";
    let series: ChartSeries[] = [];

    if (plotArea) {
        for (const child of xml.elements(plotArea)) {
            const k = localNameToKind(child.localName);
            if (k === "unknown") continue;
            kind = k;
            grouping = xml.attr(xml.element(child, "grouping") ?? null, "val") ?? "clustered";
            // Bar vs column is carried by <c:barDir val="col"|"bar">.
            if (k === "column" || k === "bar") {
                const barDir = xml.attr(xml.element(child, "barDir") ?? null, "val");
                if (barDir === "bar") kind = "bar";
                else if (barDir === "col") kind = "column";
            }
            series = xml.elements(child, "ser").map(parseSeries);
            break;
        }
    }

    return { key, title, showLegend, kind, grouping, series };
}

function deriveKey(path: string): string {
    // "word/charts/chart2.xml" → "chart2". Safe to use as a debug label
    // (never reaches the DOM as an attribute or class name).
    const segs = path.split("/");
    const file = segs[segs.length - 1] ?? "";
    const dot = file.lastIndexOf(".");
    return dot > 0 ? file.slice(0, dot) : file;
}

function localNameToKind(name: string): ChartKind {
    switch (name) {
        case "barChart": return "column"; // overridden by barDir inspection above
        case "bar3DChart": return "column";
        case "lineChart": return "line";
        case "line3DChart": return "line";
        case "pieChart": return "pie";
        case "pie3DChart": return "pie";
        case "doughnutChart": return "pie";
        default: return "unknown";
    }
}

// Walks <c:title><c:tx><c:rich><a:p><a:r><a:t> without imposing a rigid
// path — some files skip <c:tx> or put the text in <c:strRef> instead.
// Returns a plain string suitable for `textContent`.
function extractRichText(titleEl: Element): string {
    // Prefer cached string reference if present (populated when the title
    // is a cell reference like ='Sheet1'!$A$1).
    const strRef = findDescendant(titleEl, "strRef");
    if (strRef) {
        const cache = xml.element(strRef, "strCache");
        if (cache) {
            const pt = xml.element(cache, "pt");
            if (pt) {
                const v = xml.element(pt, "v");
                if (v) return v.textContent ?? "";
            }
        }
    }
    // Otherwise concatenate every <a:t> under the title (preserves
    // spacing across runs).
    const parts: string[] = [];
    walkTextRuns(titleEl, parts);
    return parts.join("");
}

function walkTextRuns(node: Element, out: string[]) {
    for (const c of xml.elements(node)) {
        if (c.localName === "t") {
            out.push(c.textContent ?? "");
        } else {
            walkTextRuns(c, out);
        }
    }
}

function findDescendant(node: Element, localName: string): Element | null {
    for (const c of xml.elements(node)) {
        if (c.localName === localName) return c;
        const d = findDescendant(c, localName);
        if (d) return d;
    }
    return null;
}

function parseSeries(serEl: Element): ChartSeries {
    const title = parseSeriesTitle(serEl);
    const color = parseSeriesColor(serEl);
    const categories = parseCategoryLabels(xml.element(serEl, "cat"));
    const values = parseNumericValues(xml.element(serEl, "val"));
    const dataPointOverrides = parseDataPointOverrides(serEl);

    return { title, color, categories, values, dataPointOverrides };
}

// Walks <c:ser> for <c:dPt> children and extracts per-point colour
// overrides. <c:idx val="N"/> is the 0-based point index, <c:spPr>
// carries the fill. Indices are bounded identically to extractPoints
// so a pathological file can't allocate an unbounded Map.
function parseDataPointOverrides(serEl: Element): Map<number, ChartDataPointOverride> {
    const out = new Map<number, ChartDataPointOverride>();
    for (const dPt of xml.elements(serEl, "dPt")) {
        const idxEl = xml.element(dPt, "idx");
        const rawIdx = idxEl ? xml.attr(idxEl, "val") : null;
        const idx = rawIdx != null ? parseInt(rawIdx, 10) : NaN;
        if (!Number.isFinite(idx) || idx < 0 || idx >= MAX_POINTS) continue;

        const spPr = xml.element(dPt, "spPr");
        if (!spPr) continue;
        const color = parseSolidFillColor(spPr);
        if (color == null) continue;
        out.set(idx, { color });
    }
    return out;
}

function parseSeriesTitle(serEl: Element): string {
    const tx = xml.element(serEl, "tx");
    if (!tx) return "";
    // Most common shape: <c:tx><c:strRef><c:strCache><c:pt><c:v>Label</c:v>.
    const strRef = xml.element(tx, "strRef");
    if (strRef) {
        const cache = xml.element(strRef, "strCache");
        if (cache) {
            const pt = xml.element(cache, "pt");
            if (pt) {
                const v = xml.element(pt, "v");
                if (v) return v.textContent ?? "";
            }
        }
    }
    // Some generators emit <c:tx><c:v>Label</c:v> directly.
    const v = xml.element(tx, "v");
    if (v) return v.textContent ?? "";
    // Or rich-text inside <c:tx><c:rich>.
    const rich = xml.element(tx, "rich");
    if (rich) {
        const parts: string[] = [];
        walkTextRuns(rich, parts);
        return parts.join("");
    }
    return "";
}

function parseSeriesColor(serEl: Element): string | null {
    const spPr = xml.element(serEl, "spPr");
    if (!spPr) return null;
    return parseSolidFillColor(spPr);
}

// Extracts a single solid-fill colour from a <c:spPr> or similar
// shape-properties element. Handles both <a:srgbClr val="RRGGBB"/>
// and <a:schemeClr val="accent1"/>. Returns a sanitised `#hex` /
// `rgb(...)` string, or null when the element carries no usable
// fill (gradient, pattern, missing).
//
// Security: the `val` attribute is attacker-controlled.
// sanitizeCssColor rejects anything that isn't a bare hex triplet
// or a recognised CSS colour; resolveSchemeColor matches `val`
// against a hard-coded allowlist of scheme keys and returns null
// otherwise. Both paths converge on a string that's safe to emit
// as an SVG attribute.
function parseSolidFillColor(spPr: Element): string | null {
    const solidFill = xml.element(spPr, "solidFill");
    if (!solidFill) return null;
    const srgb = xml.element(solidFill, "srgbClr");
    if (srgb) {
        const raw = xml.attr(srgb, "val");
        return sanitizeCssColor(raw);
    }
    const scheme = xml.element(solidFill, "schemeClr");
    if (scheme) {
        const raw = xml.attr(scheme, "val");
        const resolved = resolveSchemeColor(raw);
        // The resolved palette entry is already `#RRGGBB`, but we
        // still round-trip through sanitizeCssColor to keep the
        // security invariant stated in the module header.
        return resolved ? sanitizeCssColor(resolved) : null;
    }
    return null;
}

// Parses <c:cat>/<c:strRef>/<c:strCache>/<c:pt idx="N">/<c:v>label.
// Also handles <c:numRef>/<c:numCache> (when categories are numeric).
// Returns an array ordered by idx (gaps filled with "").
function parseCategoryLabels(catEl: Element | null): string[] {
    if (!catEl) return [];
    const cache = findCache(catEl, ["strCache", "numCache"]);
    if (!cache) return [];
    return extractPoints(cache, (v) => v);
}

function parseNumericValues(valEl: Element | null): number[] {
    if (!valEl) return [];
    const cache = findCache(valEl, ["numCache", "strCache"]);
    if (!cache) return [];
    return extractPoints(cache, (v) => {
        const n = parseFloat(v);
        return Number.isFinite(n) ? n : NaN;
    });
}

function findCache(parent: Element, localNames: string[]): Element | null {
    for (const name of localNames) {
        const ref = xml.element(parent, name === "numCache" ? "numRef" : "strRef");
        if (ref) {
            const cache = xml.element(ref, name);
            if (cache) return cache;
        }
    }
    // Some files embed the cache without the ref wrapper.
    for (const name of localNames) {
        const cache = xml.element(parent, name);
        if (cache) return cache;
    }
    return null;
}

// Extracts <c:pt idx="N"> values in idx order. Bounded to 4096 entries
// so a pathological DOCX can't turn the renderer into a denial-of-
// service (SVG with millions of nodes).
const MAX_POINTS = 4096;

function extractPoints<T>(cache: Element, mapValue: (v: string) => T): T[] {
    const pts = xml.elements(cache, "pt");
    const pairs: { idx: number; value: T }[] = [];
    for (const pt of pts) {
        const rawIdx = xml.attr(pt, "idx");
        const idx = rawIdx != null ? parseInt(rawIdx, 10) : NaN;
        if (!Number.isFinite(idx) || idx < 0 || idx >= MAX_POINTS) continue;
        const v = xml.element(pt, "v");
        const text = v ? (v.textContent ?? "") : "";
        pairs.push({ idx, value: mapValue(text) });
    }
    pairs.sort((a, b) => a.idx - b.idx);
    if (pairs.length === 0) return [];
    const maxIdx = pairs[pairs.length - 1].idx;
    const out = new Array<T>(maxIdx + 1);
    for (const { idx, value } of pairs) out[idx] = value;
    return out as T[];
}
