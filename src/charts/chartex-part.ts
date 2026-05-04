import { Part } from "../common/part";
import { OpenXmlPackage } from "../common/open-xml-package";
import xml from "../parser/xml-parser";
import { sanitizeCssColor } from "../utils";
import { resolveSchemeColor } from "../drawing/theme";
import {
    ChartExDataModel,
    ChartExKind,
    ChartExModel,
    ChartExPlaceholder,
    ChartExTreeNode,
} from "./model";

// Parses a `/word/charts/chartEx*.xml` (modern 2013+) part into either
// a `ChartExDataModel` (when the layoutId is one we render as real
// SVG — currently sunburst or treemap) or a `ChartExPlaceholder`
// (waterfall / funnel / histogram / pareto / box_whisker, and
// anything we don't recognise).
//
// Security: the chartEx root element's localName becomes the
// `data-chart-kind` attribute in the DOM, so it MUST be allowlisted
// before reaching HtmlRenderer. `kind` is coerced to "unknown" for
// anything outside CHARTEX_KINDS. Title text, category labels and
// series names all reach the DOM via textContent only. All numerics
// are parsed with parseFloat and filtered through Number.isFinite.
// Colours round-trip through sanitizeCssColor.

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

// Upper bound on <cx:pt> entries per dimension level. Matches the
// MAX_POINTS ceiling on the classic chart parser; keeps a pathological
// file from materialising an unbounded tree.
const MAX_POINTS = 4096;

export class ChartExPart extends Part {
    chart: ChartExModel;

    constructor(pkg: OpenXmlPackage, path: string) {
        super(pkg, path);
    }

    protected parseXml(root: Element) {
        this.chart = parseChartExSpace(root, this.path);
    }
}

function parseChartExSpace(root: Element, path: string): ChartExModel {
    const key = deriveKey(path);

    // <cx:chartSpace> wraps <cx:chart> which contains <cx:plotArea>
    // and optionally <cx:title>. The first recognised plot element
    // identifies the chart kind.
    const chart = xml.element(root, "chart");
    const plotArea = chart ? xml.element(chart, "plotArea") : null;
    const plotSurface = plotArea ? xml.element(plotArea, "plotSurface") : null;
    const plotAreaRegion = plotArea ? xml.element(plotArea, "plotAreaRegion") : null;

    let kind: ChartExKind = "unknown";
    let firstSeries: Element | null = null;
    if (plotAreaRegion) {
        // <cx:plotArea><cx:plotAreaRegion><cx:series layoutId="sunburst">
        for (const seriesEl of xml.elements(plotAreaRegion, "series")) {
            const layoutId = xml.attr(seriesEl, "layoutId");
            if (layoutId && CHARTEX_KINDS[layoutId]) {
                kind = CHARTEX_KINDS[layoutId];
                firstSeries = seriesEl;
                break;
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

    // Try to parse a full data model for the kinds we actually render.
    // Anything else (waterfall / funnel / histogram / pareto /
    // box_whisker / unknown) falls back to the placeholder shape —
    // see renderChartEx in html-renderer.ts for the matching dispatch.
    if ((kind === "sunburst" || kind === "treemap") && firstSeries) {
        const dataModel = tryParseDataModel(root, firstSeries, key, title, kind);
        if (dataModel) return dataModel;
        // Parsing failed (missing data / empty tree); fall through to
        // the placeholder so the reader still sees something labelled.
    }

    const placeholder: ChartExPlaceholder = { shape: "placeholder", key, title, kind };
    return placeholder;
}

// Full data-model parser for sunburst / treemap. Returns null when
// the chart has no usable hierarchy — caller falls back to the
// placeholder shape.
function tryParseDataModel(
    root: Element,
    seriesEl: Element,
    key: string,
    title: string,
    kind: "sunburst" | "treemap",
): ChartExDataModel | null {
    // <cx:dataId val="0"> on the series points at a <cx:data id="0">
    // inside <cx:chartData>. When missing, default to the first
    // <cx:data> block (common in minimal generators).
    const dataIdEl = xml.element(seriesEl, "dataId");
    const dataId = dataIdEl ? xml.attr(dataIdEl, "val") : null;

    const chartData = xml.element(root, "chartData");
    if (!chartData) return null;

    let dataEl: Element | null = null;
    for (const d of xml.elements(chartData, "data")) {
        if (dataId == null || xml.attr(d, "id") === dataId) {
            dataEl = d;
            break;
        }
    }
    if (!dataEl) return null;

    // Extract the multi-level categories and the single value dimension.
    const catDim = findDimension(dataEl, "strDim", "cat")
        ?? findDimension(dataEl, "numDim", "cat"); // rare — numeric categories
    const valDim = findDimension(dataEl, "numDim", "val");

    if (!catDim || !valDim) return null;

    const catLevels = parseStringLevels(catDim);
    const values = parseNumericLevel(valDim);

    if (catLevels.length === 0 || values.length === 0) return null;

    // Per-data-point colour overrides from <cx:dataPt idx="N">. Indexed
    // by leaf point index (same domain as the value dim).
    const dataPointColors = parseDataPointOverrides(seriesEl);

    const root0 = buildCategoryTree(catLevels, values, dataPointColors);
    if (!root0 || root0.children.length === 0) return null;

    const maxDepth = computeMaxDepth(root0);

    return { shape: "data", key, title, kind, root: root0, maxDepth };
}

// Finds a <cx:strDim>/<cx:numDim> child of <cx:data> whose `type`
// attribute matches. <cx:data> may contain several of each — most
// commonly one `cat` dim and one `val` dim, but sunburst / treemap
// sometimes declare additional size/colour dims which we ignore.
function findDimension(dataEl: Element, localName: string, type: string): Element | null {
    for (const d of xml.elements(dataEl, localName)) {
        if (xml.attr(d, "type") === type) return d;
    }
    return null;
}

// Level shape from <cx:strDim> (one entry per <cx:lvl>). `points[idx]`
// is the label at that index in this level; `parents[idx]` is the idx
// of the parent in the previous (shallower) level, or -1 when the
// attribute is absent / unparseable. Level 0 always has parents = [].
interface CategoryLevel {
    points: string[];
    parents: number[];
}

function parseStringLevels(dim: Element): CategoryLevel[] {
    const out: CategoryLevel[] = [];
    for (const lvl of xml.elements(dim, "lvl")) {
        const level = parseStringLevel(lvl);
        if (level) out.push(level);
    }
    return out;
}

function parseStringLevel(lvl: Element): CategoryLevel | null {
    const ptCountAttr = xml.attr(lvl, "ptCount");
    const ptCount = ptCountAttr != null ? parseInt(ptCountAttr, 10) : NaN;
    const declared = Number.isFinite(ptCount) && ptCount >= 0 && ptCount < MAX_POINTS
        ? ptCount : 0;

    const points: string[] = new Array(declared).fill("");
    const parents: number[] = new Array(declared).fill(-1);
    let maxIdx = declared - 1;

    for (const pt of xml.elements(lvl, "pt")) {
        const rawIdx = xml.attr(pt, "idx");
        const idx = rawIdx != null ? parseInt(rawIdx, 10) : NaN;
        if (!Number.isFinite(idx) || idx < 0 || idx >= MAX_POINTS) continue;
        if (idx > maxIdx) {
            // Grow to accommodate a higher idx than ptCount advertised.
            while (points.length <= idx) points.push("");
            while (parents.length <= idx) parents.push(-1);
            maxIdx = idx;
        }
        points[idx] = pt.textContent ?? "";
        // Parent pointer: the spec uses `parent`; older generators used
        // `parentIdx`. Both are 0-based indices into the previous level.
        const parentAttr = xml.attr(pt, "parent") ?? xml.attr(pt, "parentIdx");
        if (parentAttr != null) {
            const p = parseInt(parentAttr, 10);
            if (Number.isFinite(p) && p >= 0 && p < MAX_POINTS) parents[idx] = p;
        }
    }

    if (points.length === 0) return null;
    return { points, parents };
}

// Numeric value dim: a single <cx:lvl> whose <cx:pt> elements carry
// finite numbers. Values outside the range or non-numeric entries
// become NaN and are treated as zero when aggregating.
function parseNumericLevel(dim: Element): number[] {
    const lvl = xml.element(dim, "lvl");
    if (!lvl) return [];

    const ptCountAttr = xml.attr(lvl, "ptCount");
    const ptCount = ptCountAttr != null ? parseInt(ptCountAttr, 10) : NaN;
    const declared = Number.isFinite(ptCount) && ptCount >= 0 && ptCount < MAX_POINTS
        ? ptCount : 0;

    const values: number[] = new Array(declared).fill(NaN);
    let maxIdx = declared - 1;

    for (const pt of xml.elements(lvl, "pt")) {
        const rawIdx = xml.attr(pt, "idx");
        const idx = rawIdx != null ? parseInt(rawIdx, 10) : NaN;
        if (!Number.isFinite(idx) || idx < 0 || idx >= MAX_POINTS) continue;
        if (idx > maxIdx) {
            while (values.length <= idx) values.push(NaN);
            maxIdx = idx;
        }
        const raw = pt.textContent ?? "";
        const n = parseFloat(raw);
        values[idx] = Number.isFinite(n) ? n : NaN;
    }

    return values;
}

// Per-leaf colour overrides. Keyed by leaf point index (same domain
// as the values array). Treemap / sunburst commonly colour the
// top-level groups, so callers must also walk up to an ancestor when
// the leaf's own entry is missing — handled in buildCategoryTree.
function parseDataPointOverrides(seriesEl: Element): Map<number, string> {
    const out = new Map<number, string>();
    for (const dPt of xml.elements(seriesEl, "dataPt")) {
        const rawIdx = xml.attr(dPt, "idx");
        const idx = rawIdx != null ? parseInt(rawIdx, 10) : NaN;
        if (!Number.isFinite(idx) || idx < 0 || idx >= MAX_POINTS) continue;

        const spPr = xml.element(dPt, "spPr");
        if (!spPr) continue;
        const color = parseSolidFillColor(spPr);
        if (color) out.set(idx, color);
    }
    return out;
}

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
        return resolved ? sanitizeCssColor(resolved) : null;
    }
    return null;
}

// Builds the category tree from the level-by-level arrays.
//
// ChartEx category hierarchies come out flat: one <cx:lvl> per depth,
// and each <cx:pt> in a deeper level references its parent via the
// `parent` attribute. The deepest level's indices are also the leaf
// indices into the value dim — that's the key that ties values,
// colours and labels together.
//
// Edge cases handled:
//   - single-level data: every point in level 0 becomes a leaf that
//     sits directly under the synthetic root.
//   - empty intermediate level: that whole subtree is dropped (can't
//     recover labels we don't have).
//   - missing parent: the point attaches to a synthetic "Other"
//     container at its own level so it still renders.
//   - NaN or negative leaf values: treated as 0 so the slice disappears
//     rather than breaking angular math. The tree node is still emitted
//     so the label remains navigable.
function buildCategoryTree(
    levels: CategoryLevel[],
    values: number[],
    colors: Map<number, string>,
): ChartExTreeNode {
    const root: ChartExTreeNode = {
        label: "", value: 0, color: null, children: [], level: -1, leafIndex: -1,
    };

    if (levels.length === 0) return root;

    // Nodes per level, in idx order, so deeper levels can link back via
    // their `parent` attribute.
    const perLevel: ChartExTreeNode[][] = [];

    // Level 0 attaches to the root.
    const lvl0 = levels[0];
    const level0Nodes: ChartExTreeNode[] = lvl0.points.map((label, idx) => {
        const isLeaf = levels.length === 1;
        const leafIndex = isLeaf ? idx : -1;
        const raw = isLeaf ? values[idx] : NaN;
        const value = Number.isFinite(raw) && raw > 0 ? raw : 0;
        const color = isLeaf ? (colors.get(idx) ?? null) : null;
        return { label, value, color, children: [], level: 0, leafIndex };
    });
    for (const node of level0Nodes) root.children.push(node);
    perLevel.push(level0Nodes);

    // Each subsequent level: attach each node to its parent at the
    // previous level. The deepest level's idx is also the value idx.
    for (let d = 1; d < levels.length; d++) {
        const lvl = levels[d];
        const isDeepest = d === levels.length - 1;
        const parentLevel = perLevel[d - 1];
        const levelNodes: ChartExTreeNode[] = [];
        for (let idx = 0; idx < lvl.points.length; idx++) {
            const label = lvl.points[idx];
            const parentIdx = lvl.parents[idx];
            // Leaves inherit the colour from their own leaf index first;
            // if that's missing, walk up the chain in case an ancestor
            // was coloured instead.
            const leafIndex = isDeepest ? idx : -1;
            let color: string | null = null;
            if (isDeepest) {
                color = colors.get(idx) ?? null;
            }
            const raw = isDeepest ? values[idx] : NaN;
            const value = Number.isFinite(raw) && raw > 0 ? raw : 0;
            const node: ChartExTreeNode = {
                label, value, color, children: [], level: d, leafIndex,
            };

            const parent = parentIdx >= 0 && parentIdx < parentLevel.length
                ? parentLevel[parentIdx]
                : null;
            if (parent) {
                parent.children.push(node);
            } else {
                // Orphan — attach to synthetic "Other" container on the
                // root so the slice still renders somewhere rather than
                // silently disappearing.
                root.children.push(node);
            }
            levelNodes.push(node);
        }
        perLevel.push(levelNodes);
    }

    // Second pass: propagate summed values up from leaves and propagate
    // ancestor colours down to leaves whose own override was missing.
    sumValues(root);
    propagateColors(root, null);

    return root;
}

// Sum leaf values up through the tree. Intermediate nodes that already
// had a non-zero value (unusual but legal) are preserved and used as a
// lower bound; typically everything comes from the leaves.
function sumValues(node: ChartExTreeNode): number {
    if (node.children.length === 0) {
        return node.value;
    }
    let total = 0;
    for (const child of node.children) {
        total += sumValues(child);
    }
    if (total > node.value) node.value = total;
    return node.value;
}

// Push a parent's colour down onto any descendant that didn't have its
// own <cx:dataPt> override. This matches Word's rendering, where a
// sunburst slice inherits its parent's fill unless the data point
// explicitly specifies one.
function propagateColors(node: ChartExTreeNode, inherited: string | null) {
    const effective = node.color ?? inherited;
    if (node.color == null && inherited != null) {
        node.color = inherited;
    }
    for (const child of node.children) propagateColors(child, effective);
}

function computeMaxDepth(root: ChartExTreeNode): number {
    let max = 0;
    function visit(n: ChartExTreeNode, d: number) {
        if (d > max) max = d;
        for (const c of n.children) visit(c, d + 1);
    }
    for (const c of root.children) visit(c, 1);
    return max;
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
// the classic <c:title>. We also honour the simpler <cx:tx><cx:txData>
// <cx:v> form used by more compact generators.
function extractTitle(chart: Element): string {
    const titleEl = xml.element(chart, "title");
    if (!titleEl) return "";
    const parts: string[] = [];
    walkTextRuns(titleEl, parts);
    return parts.join("");
}

function walkTextRuns(node: Element, out: string[]) {
    for (const c of xml.elements(node)) {
        if (c.localName === "t" || c.localName === "v") {
            out.push(c.textContent ?? "");
        } else {
            walkTextRuns(c, out);
        }
    }
}
