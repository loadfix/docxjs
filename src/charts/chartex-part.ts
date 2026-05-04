import { Part } from "../common/part";
import { OpenXmlPackage } from "../common/open-xml-package";
import xml from "../parser/xml-parser";
import { ChartExKind, ChartExPlaceholder } from "./model";

// Parses a `/word/charts/chartEx*.xml` (modern 2013+) part into a
// `ChartExPlaceholder`. We do not render sunburst / waterfall /
// funnel / treemap / histogram / pareto / box-whisker charts as
// actual SVG — the layout requirements are substantial enough that
// a placeholder with the chart title is a better short-term outcome
// than silently dropping the drawing.
//
// Security: the chartEx root element's localName becomes the
// `data-chart-kind` attribute in the DOM, so it MUST be allowlisted
// before reaching HtmlRenderer. `kind` is coerced to "unknown" for
// anything outside CHARTEX_KINDS. The title text reaches the DOM
// via textContent only.

const CHARTEX_KINDS: Record<string, ChartExKind> = {
    sunburst: "sunburst",
    waterfall: "waterfall",
    funnel: "funnel",
    treemap: "treemap",
    histogram: "histogram",
    pareto: "pareto",
    // ChartEx spells this as "boxWhisker"; we normalise to snake_case
    // so the attribute looks natural in HTML (`data-chart-kind="box_whisker"`).
    boxWhisker: "box_whisker",
    clusteredColumn: "unknown",
    regionMap: "unknown",
};

export class ChartExPart extends Part {
    chart: ChartExPlaceholder;

    constructor(pkg: OpenXmlPackage, path: string) {
        super(pkg, path);
    }

    protected parseXml(root: Element) {
        this.chart = parseChartExSpace(root, this.path);
    }
}

function parseChartExSpace(root: Element, path: string): ChartExPlaceholder {
    const key = deriveKey(path);

    // <cx:chartSpace> wraps <cx:chart> which contains <cx:plotArea>
    // and optionally <cx:title>. The first recognised plot element
    // identifies the chart kind.
    const chart = xml.element(root, "chart");
    const plotArea = chart ? xml.element(chart, "plotArea") : null;
    const plotSurface = plotArea ? xml.element(plotArea, "plotSurface") : null;

    let kind: ChartExKind = "unknown";
    if (plotArea) {
        // <cx:plotArea><cx:plotAreaRegion><cx:series layoutId="sunburst">
        const plotAreaRegion = xml.element(plotArea, "plotAreaRegion");
        if (plotAreaRegion) {
            for (const seriesEl of xml.elements(plotAreaRegion, "series")) {
                const layoutId = xml.attr(seriesEl, "layoutId");
                if (layoutId && CHARTEX_KINDS[layoutId]) {
                    kind = CHARTEX_KINDS[layoutId];
                    break;
                }
            }
        }
    }
    // Some chartEx generators put the layoutId on the root child instead.
    if (kind === "unknown" && plotSurface) {
        const layoutId = xml.attr(plotSurface, "layoutId");
        if (layoutId && CHARTEX_KINDS[layoutId]) {
            kind = CHARTEX_KINDS[layoutId];
        }
    }

    const title = chart ? extractTitle(chart) : "";

    return { key, title, kind };
}

function deriveKey(path: string): string {
    const segs = path.split("/");
    const file = segs[segs.length - 1] ?? "";
    const dot = file.lastIndexOf(".");
    return dot > 0 ? file.slice(0, dot) : file;
}

// <cx:title><cx:tx><cx:rich><a:p><a:r><a:t>…</a:t>. The shape differs
// between generators — we walk every descendant <a:t> under the title
// and concatenate, which matches how src/charts/chart-part.ts handles
// the classic <c:title>.
function extractTitle(chart: Element): string {
    const titleEl = xml.element(chart, "title");
    if (!titleEl) return "";
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
